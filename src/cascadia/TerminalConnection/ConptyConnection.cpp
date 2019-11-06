// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

#include "pch.h"
#include "ConptyConnection.h"

#include <windows.h>
#include <sstream>

#include "ConptyConnection.g.cpp"

#include "../../types/inc/Utils.hpp"
#include "../../types/inc/UTF8OutPipeReader.hpp"

using namespace ::Microsoft::Console;

#pragma region Environment Strings

// A case-insensitive wide-character map is used to store environment variables
// due to documented requirements:
//
//      "All strings in the environment block must be sorted alphabetically by name.
//      The sort is case-insensitive, Unicode order, without regard to locale.
//      Because the equal sign is a separator, it must not be used in the name of
//      an environment variable."
//      https://docs.microsoft.com/en-us/windows/desktop/ProcThread/changing-environment-variables
//
struct WStringCaseInsensitiveCompare
{
    [[nodiscard]] bool operator()(const std::wstring& lhs, const std::wstring& rhs) const noexcept
    {
        return (::_wcsicmp(lhs.c_str(), rhs.c_str()) < 0);
    }
};

using EnvironmentVariableMapW = std::map<std::wstring, std::wstring, WStringCaseInsensitiveCompare>;

[[nodiscard]] HRESULT UpdateEnvironmentMapW(EnvironmentVariableMapW& map) noexcept;

[[nodiscard]] HRESULT EnvironmentMapToEnvironmentStringsW(EnvironmentVariableMapW& map,
                                                          std::vector<wchar_t>& newEnvVars) noexcept;

// Function Description:
// - Updates an EnvironmentVariableMapW with the current process's unicode
//   environment variables ignoring ones already set in the provided map.
// Arguments:
// - map: The map to populate with the current processes's environment variables.
// Return Value:
// - S_OK if we succeeded, or an appropriate HRESULT for failing
[[nodiscard]] __declspec(noinline) inline HRESULT UpdateEnvironmentMapW(EnvironmentVariableMapW& map) noexcept
{
    LPWCH currentEnvVars{};
    auto freeCurrentEnv = wil::scope_exit([&] {
        if (currentEnvVars)
        {
            (void)FreeEnvironmentStringsW(currentEnvVars);
            currentEnvVars = nullptr;
        }
    });

    RETURN_IF_NULL_ALLOC(currentEnvVars = ::GetEnvironmentStringsW());

    // Each entry is NULL-terminated; block is guaranteed to be double-NULL terminated at a minimum.
    for (wchar_t const* lastCh{ currentEnvVars }; *lastCh != '\0'; ++lastCh)
    {
        // Copy current entry into temporary map.
        const size_t cchEntry{ ::wcslen(lastCh) };
        const std::wstring_view entry{ lastCh, cchEntry };

        // Every entry is of the form "name=value\0".
        auto pos = entry.find_first_of(L"=", 0, 1);
        RETURN_HR_IF(E_UNEXPECTED, pos == std::wstring::npos);

        std::wstring name{ entry.substr(0, pos) }; // portion before '='
        std::wstring value{ entry.substr(pos + 1) }; // portion after '='

        // Don't replace entries that already exist.
        map.try_emplace(std::move(name), std::move(value));
        lastCh += cchEntry;
    }

    return S_OK;
}

// Function Description:
// - Creates a new environment block using the provided vector as appropriate
//   (resizing if needed) based on the provided environment variable map
//   matching the format of GetEnvironmentStringsW.
// Arguments:
// - map: The map to populate the new environment block vector with.
// - newEnvVars: The vector that will be used to create the new environment block.
// Return Value:
// - S_OK if we succeeded, or an appropriate HRESULT for failing
[[nodiscard]] __declspec(noinline) inline HRESULT EnvironmentMapToEnvironmentStringsW(EnvironmentVariableMapW& map, std::vector<wchar_t>& newEnvVars) noexcept
{
    // Clear environment block before use.
    constexpr size_t cbChar{ sizeof(decltype(newEnvVars.begin())::value_type) };

    if (!newEnvVars.empty())
    {
        ::SecureZeroMemory(newEnvVars.data(), newEnvVars.size() * cbChar);
    }

    // Resize environment block to fit map.
    size_t cchEnv{ 2 }; // For the block's double NULL-terminator.
    for (const auto& [name, value] : map)
    {
        // Final form of "name=value\0".
        cchEnv += name.size() + 1 + value.size() + 1;
    }
    newEnvVars.resize(cchEnv);

    // Ensure new block is wiped if we exit due to failure.
    auto zeroNewEnv = wil::scope_exit([&] {
        ::SecureZeroMemory(newEnvVars.data(), newEnvVars.size() * cbChar);
    });

    // Transform each map entry and copy it into the new environment block.
    LPWCH pEnvVars{ newEnvVars.data() };
    size_t cbRemaining{ cchEnv * cbChar };
    for (const auto& [name, value] : map)
    {
        // Final form of "name=value\0" for every entry.
        {
            const size_t cchSrc{ name.size() };
            const size_t cbSrc{ cchSrc * cbChar };
            RETURN_HR_IF(E_OUTOFMEMORY, memcpy_s(pEnvVars, cbRemaining, name.c_str(), cbSrc) != 0);
            pEnvVars += cchSrc;
            cbRemaining -= cbSrc;
        }

        RETURN_HR_IF(E_OUTOFMEMORY, memcpy_s(pEnvVars, cbRemaining, L"=", cbChar) != 0);
        ++pEnvVars;
        cbRemaining -= cbChar;

        {
            const size_t cchSrc{ value.size() };
            const size_t cbSrc{ cchSrc * cbChar };
            RETURN_HR_IF(E_OUTOFMEMORY, memcpy_s(pEnvVars, cbRemaining, value.c_str(), cbSrc) != 0);
            pEnvVars += cchSrc;
            cbRemaining -= cbSrc;
        }

        RETURN_HR_IF(E_OUTOFMEMORY, memcpy_s(pEnvVars, cbRemaining, L"\0", cbChar) != 0);
        ++pEnvVars;
        cbRemaining -= cbChar;
    }

    // Environment block only has to be NULL-terminated, but double NULL-terminate anyway.
    RETURN_HR_IF(E_OUTOFMEMORY, memcpy_s(pEnvVars, cbRemaining, L"\0\0", cbChar * 2) != 0);
    cbRemaining -= cbChar * 2;

    RETURN_HR_IF(E_UNEXPECTED, cbRemaining != 0);

    zeroNewEnv.release(); // success; don't wipe new environment block on exit

    return S_OK;
}

#pragma endregion

namespace winrt::Microsoft::Terminal::TerminalConnection::implementation
{
    // Function Description:
    // - creates some basic anonymous pipes and passes them to CreatePseudoConsole
    // Arguments:
    // - size: The size of the conpty to create, in characters.
    // - phInput: Receives the handle to the newly-created anonymous pipe for writing input to the conpty.
    // - phOutput: Receives the handle to the newly-created anonymous pipe for reading the output of the conpty.
    // - phPty: Receives a token value to identify this conpty
    static HRESULT _CreatePseudoConsoleAndPipes(COORD size, DWORD dwFlags, HANDLE* phInput, HANDLE* phOutput, HPCON* phPC)
    {
        RETURN_HR_IF(E_INVALIDARG, phPC == nullptr || phInput == nullptr || phOutput == nullptr);

        wil::unique_hfile outPipeOurSide, outPipePseudoConsoleSide;
        wil::unique_hfile inPipeOurSide, inPipePseudoConsoleSide;

        RETURN_IF_WIN32_BOOL_FALSE(CreatePipe(&inPipePseudoConsoleSide, &inPipeOurSide, NULL, 0));
        RETURN_IF_WIN32_BOOL_FALSE(CreatePipe(&outPipeOurSide, &outPipePseudoConsoleSide, NULL, 0));
        RETURN_IF_FAILED(CreatePseudoConsole(size, inPipePseudoConsoleSide.get(), outPipePseudoConsoleSide.get(), dwFlags, phPC));
        *phInput = inPipeOurSide.release();
        *phOutput = outPipeOurSide.release();
        return S_OK;
    }

    // Prepares the `lpAttributeList` member of a STARTUPINFOEX for attaching a
    //      client application to a pseudoconsole.
    // Prior to calling this function, hPC should be initialized with a call to
    //      CreatePseudoConsole, and the pAttrList should be initialized with a call
    //      to InitializeProcThreadAttributeList. The caller of
    //      InitializeProcThreadAttributeList should add one to the dwAttributeCount
    //      param when creating the attribute list for usage by this function.
    static HRESULT _AttachPseudoConsole(HPCON hPC, LPPROC_THREAD_ATTRIBUTE_LIST lpAttributeList)
    {
        RETURN_IF_WIN32_BOOL_FALSE(UpdateProcThreadAttribute(lpAttributeList,
                                                             0,
                                                             PROC_THREAD_ATTRIBUTE_PSEUDOCONSOLE,
                                                             hPC,
                                                             sizeof(HPCON),
                                                             NULL,
                                                             NULL));
        return S_OK;
    }

    void ConptyConnection::_CreatePseudoConsole()
    {
        COORD dimensions{ gsl::narrow_cast<SHORT>(_initialCols), gsl::narrow_cast<SHORT>(_initialRows) };
        THROW_IF_FAILED(_CreatePseudoConsoleAndPipes(dimensions, 0, &_inPipe, &_outPipe, &_hPC));

        STARTUPINFOEX siEx{ 0 };
        siEx.StartupInfo.cb = sizeof(STARTUPINFOEX);
        siEx.StartupInfo.dwFlags = STARTF_USESTDHANDLES;

        size_t size{};
        InitializeProcThreadAttributeList(NULL, 1, 0, (PSIZE_T)&size);
        auto attrList = std::make_unique<std::byte[]>(size);
        siEx.lpAttributeList = reinterpret_cast<PPROC_THREAD_ATTRIBUTE_LIST>(attrList.get());
        THROW_IF_WIN32_BOOL_FALSE(InitializeProcThreadAttributeList(siEx.lpAttributeList, 1, 0, (PSIZE_T)&size));

        THROW_IF_FAILED(_AttachPseudoConsole(_hPC.get(), siEx.lpAttributeList));

        std::wstring cmdline{ _commandline.c_str() }; // mutable copy -- required for CreateProcessW

        EnvironmentVariableMapW extraEnvVars;

        {
            // Convert connection Guid to string and ignore the enclosing '{}'.
            std::wstring wsGuid{ Utils::GuidToString(_guid) };
            wsGuid.pop_back();

            const wchar_t* const pwszGuid{ wsGuid.data() + 1 };

            // Ensure every connection has the unique identifier in the environment.
            extraEnvVars.emplace(L"WT_SESSION", pwszGuid);
        }

        std::vector<wchar_t> newEnvVars;
        auto zeroNewEnv = wil::scope_exit([&] {
            ::SecureZeroMemory(newEnvVars.data(),
                               newEnvVars.size() * sizeof(decltype(newEnvVars.begin())::value_type));
        });

        EnvironmentVariableMapW tempEnvMap{ extraEnvVars };
        auto zeroEnvMap = wil::scope_exit([&] {
            // Can't zero the keys, but at least we can zero the values.
            for (auto& [name, value] : tempEnvMap)
            {
                ::SecureZeroMemory(value.data(), value.size() * sizeof(decltype(value.begin())::value_type));
            }

            tempEnvMap.clear();
        });

        THROW_IF_FAILED(UpdateEnvironmentMapW(tempEnvMap));
        THROW_IF_FAILED(EnvironmentMapToEnvironmentStringsW(tempEnvMap, newEnvVars));

        LPWCH lpEnvironment = newEnvVars.empty() ? nullptr : newEnvVars.data();

        // If we have a startingTitle, create a mutable character buffer to add
        // it to the STARTUPINFO.
        std::wstring mutableTitle{};
        if (!_startingTitle.empty())
        {
            mutableTitle = _startingTitle;
            siEx.StartupInfo.lpTitle = mutableTitle.data();
        }

        THROW_IF_WIN32_BOOL_FALSE(CreateProcessW(
            nullptr,
            cmdline.data(),
            nullptr, // lpProcessAttributes
            nullptr, // lpThreadAttributes
            false, // bInheritHandles
            EXTENDED_STARTUPINFO_PRESENT | CREATE_UNICODE_ENVIRONMENT, // dwCreationFlags
            lpEnvironment, // lpEnvironment
            _startingDirectory.c_str(),
            &siEx.StartupInfo, // lpStartupInfo
            &_piClient // lpProcessInformation
            ));

        DeleteProcThreadAttributeList(siEx.lpAttributeList);
    }

    ConptyConnection::ConptyConnection(const hstring& commandline,
                                       const hstring& startingDirectory,
                                       const hstring& startingTitle,
                                       const uint32_t initialRows,
                                       const uint32_t initialCols,
                                       const guid& initialGuid) :
        _initialRows{ initialRows },
        _initialCols{ initialCols },
        _commandline{ commandline },
        _startingDirectory{ startingDirectory },
        _startingTitle{ startingTitle },
        _guid{ initialGuid }
    {
        if (_guid == guid{})
        {
            _guid = Utils::CreateGuid();
        }
    }

    winrt::guid ConptyConnection::Guid() const noexcept
    {
        return _guid;
    }

    winrt::event_token ConptyConnection::TerminalOutput(Microsoft::Terminal::TerminalConnection::TerminalOutputEventArgs const& handler)
    {
        return _outputHandlers.add(handler);
    }

    void ConptyConnection::TerminalOutput(winrt::event_token const& token) noexcept
    {
        _outputHandlers.remove(token);
    }

    winrt::event_token ConptyConnection::TerminalDisconnected(Microsoft::Terminal::TerminalConnection::TerminalDisconnectedEventArgs const& handler)
    {
        return _disconnectHandlers.add(handler);
    }

    void ConptyConnection::TerminalDisconnected(winrt::event_token const& token) noexcept
    {
        _disconnectHandlers.remove(token);
    }

    void ConptyConnection::Start()
    {
        _CreatePseudoConsole();

        _startTime = std::chrono::high_resolution_clock::now();

        // Create our own output handling thread
        // This must be done after the pipes are populated.
        // Each connection needs to make sure to drain the output from its backing host.
        _hOutputThread.reset(CreateThread(
            nullptr,
            0,
            [](LPVOID lpParameter) {
                ConptyConnection* const pInstance = (ConptyConnection*)lpParameter;
                return pInstance->_OutputThread();
            },
            this,
            0,
            nullptr));

        THROW_LAST_ERROR_IF_NULL(_hOutputThread);

        _connected = true;
    }

    void ConptyConnection::WriteInput(hstring const& data)
    {
        if (!_connected || _closing.load())
        {
            return;
        }

        // convert from UTF-16LE to UTF-8 as ConPty expects UTF-8
        std::string str = winrt::to_string(data);
        bool fSuccess = !!WriteFile(_inPipe.get(), str.c_str(), (DWORD)str.length(), nullptr, nullptr);
        fSuccess;
    }

    void ConptyConnection::Resize(uint32_t rows, uint32_t columns)
    {
        if (!_connected)
        {
            _initialRows = rows;
            _initialCols = columns;
        }
        else if (!_closing.load())
        {
            ResizePseudoConsole(_hPC.get(), { Utils::ClampToShortMax(columns, 1), Utils::ClampToShortMax(rows, 1) });
        }
    }

    void ConptyConnection::Close()
    {
        if (!_connected)
        {
            return;
        }

        if (!_closing.exchange(true))
        {
            _hPC.reset();

            _inPipe.reset();
            _outPipe.reset();

            // Tear down our output thread -- now that the output pipe was closed on the
            // far side, we can run down our local reader.
            WaitForSingleObject(_hOutputThread.get(), INFINITE);
            _hOutputThread.reset();

            // Wait for the client to terminate.
            WaitForSingleObject(_piClient.hProcess, INFINITE);

            _piClient.reset();
        }
    }

    DWORD ConptyConnection::_OutputThread()
    {
        UTF8OutPipeReader pipeReader{ _outPipe.get() };
        std::string_view strView{};

        // process the data of the output pipe in a loop
        while (true)
        {
            HRESULT result = pipeReader.Read(strView);
            if (FAILED(result) || result == S_FALSE)
            {
                if (_closing.load())
                {
                    // This is okay, break out to kill the thread
                    return 0;
                }

                _disconnectHandlers();
                return (DWORD)-1;
            }

            if (strView.empty())
            {
                return 0;
            }

            if (!_recievedFirstByte)
            {
                auto now = std::chrono::high_resolution_clock::now();
                std::chrono::duration<double> delta = now - _startTime;

                TraceLoggingWrite(g_hTerminalConnectionProvider,
                                  "RecievedFirstByte",
                                  TraceLoggingDescription("An event emitted when the connection recieves the first byte"),
                                  TraceLoggingGuid(_guid, "SessionGuid", "The WT_SESSION's GUID"),
                                  TraceLoggingFloat64(delta.count(), "Duration"),
                                  TraceLoggingKeyword(MICROSOFT_KEYWORD_MEASURES),
                                  TelemetryPrivacyDataTag(PDT_ProductAndServicePerformance));
                _recievedFirstByte = true;
            }

            // Convert buffer to hstring
            auto hstr{ winrt::to_hstring(strView) };

            // Pass the output to our registered event handlers
            _outputHandlers(hstr);
        }

        return 0;
    }

}
