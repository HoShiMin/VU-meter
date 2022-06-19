#define WIN32_LEAN_AND_MEAN
#define WIN32_NO_STATUS
#define NOMINMAX
#include <Windows.h>

#include <mmdeviceapi.h>
#include <audiopolicy.h>
#include <endpointvolume.h>

#include <hidusage.h>

#include <string>
#include <atomic>
#include <charconv>

namespace
{
    
class Console
{
public:
    using Char = wchar_t;
    
    class Attr
    {
    private:
        unsigned short m_value;

    public:
        Attr() : m_value(FOREGROUND_RED | FOREGROUND_GREEN | FOREGROUND_BLUE)
        {
        }

        Attr(unsigned short value) : m_value(value)
        {
        }

        Attr& operator = (unsigned short value)
        {
            m_value = value;
            return *this;
        }
    };

private:
    HWND m_hwnd;
    HANDLE m_stdin;
    HANDLE m_stdout;

public:
    Console()
        : m_hwnd(GetConsoleWindow())
        , m_stdin(GetStdHandle(STD_INPUT_HANDLE))
        , m_stdout(GetStdHandle(STD_OUTPUT_HANDLE))
    {
    }
        
    bool setResizeable(bool resizeable)
    {
        const DWORD style = GetWindowLongW(m_hwnd, GWL_STYLE);
        return SetWindowLongW(
            m_hwnd,
            GWL_STYLE,
            resizeable
            ? (style | WS_MAXIMIZEBOX | WS_SIZEBOX)
            : (style & ~WS_MAXIMIZEBOX & ~WS_SIZEBOX)
        );
    }

    bool setInteractable(bool interactable)
    {
        DWORD mode = 0;
        if (!GetConsoleMode(m_stdin, &mode))
        {
            return false;
        }

        mode = interactable
            ? (mode | (ENABLE_QUICK_EDIT_MODE | ENABLE_EXTENDED_FLAGS))
            : (mode & ~(ENABLE_QUICK_EDIT_MODE | ENABLE_EXTENDED_FLAGS));

        return !!SetConsoleMode(m_stdin, mode);
    }

    bool resize(unsigned short width, unsigned short height)
    {
        CONSOLE_SCREEN_BUFFER_INFOEX info{};
        info.cbSize = sizeof(info);
        const bool queryStatus = !!GetConsoleScreenBufferInfoEx(m_stdout, &info);
        if (!queryStatus)
        {
            return false;
        }

        info.dwSize = { static_cast<short>(width), static_cast<short>(height) };
        info.srWindow = { 0, 0, static_cast<short>(width), static_cast<short>(height) };
        const bool setStatus = !!SetConsoleScreenBufferInfoEx(m_stdout, &info);

        return setStatus;
    }

    bool setCursorView(bool visible, unsigned long cellFillingPercentage)
    {
        CONSOLE_CURSOR_INFO cursor{ cellFillingPercentage, visible };
        return !!SetConsoleCursorInfo(m_stdout, &cursor);
    }

    bool write(const wchar_t* buf, unsigned int numberOfChars)
    {
        unsigned long written = 0;
        return !!WriteConsoleW(m_stdout, buf, numberOfChars, &written, nullptr);
    }

    template <typename T>
    bool write(const T& buf)
    {
        return write(reinterpret_cast<const wchar_t*>(&buf), sizeof(buf) / sizeof(wchar_t));
    }

    bool writeTo(const wchar_t* buf, unsigned int numberOfChars, unsigned short x, unsigned short y)
    {
        unsigned long written = 0;
        return !!WriteConsoleOutputCharacterW(m_stdout, buf, numberOfChars, { static_cast<short>(x), static_cast<short>(y) }, &written);
    }

    template <typename T>
    bool writeTo(const T& buf, unsigned short x, unsigned short y)
    {
        return writeTo(reinterpret_cast<const wchar_t*>(&buf), sizeof(buf) / sizeof(wchar_t), x, y);
    }

    bool writeAttrsTo(const unsigned short* attrs, unsigned int numberOfAttrs, unsigned int x, unsigned int y)
    {
        unsigned long written = 0;
        return !!WriteConsoleOutputAttribute(m_stdout, attrs, numberOfAttrs, { static_cast<short>(x), static_cast<short>(y) }, &written);
    }

    template <typename T>
    bool writeAttrsTo(const T& attrs, unsigned int x, unsigned int y)
    {
        return writeAttrsTo(reinterpret_cast<const unsigned short*>(&attrs), sizeof(attrs) / sizeof(unsigned short), static_cast<short>(x), static_cast<short>(y));
    }
};



Console g_console;



template <template <typename> typename Data>
class ScreenBuffer
{
private:
    Console& m_console;
    Data<Console::Char> m_charsBuffer;
    Data<Console::Attr> m_attrsBuffer;

public:
    ScreenBuffer(Console& console) : m_console(console), m_charsBuffer{}, m_attrsBuffer{}
    {
    }

    void render()
    {
        m_console.writeTo(m_charsBuffer, 0, 0);
        m_console.writeAttrsTo(m_attrsBuffer, 0, 0);
    }

    const Data<Console::Char>& chars() const
    {
        return m_charsBuffer;
    }

    Data<Console::Char>& chars()
    {
        return m_charsBuffer;
    }

    const Data<Console::Attr>& attrs() const
    {
        return m_attrsBuffer;
    }

    Data<Console::Attr>& attrs()
    {
        return m_attrsBuffer;
    }
};



template <typename CellType>
struct VUMeter
{
    CellType leftBracket;
    CellType volume[3];
    CellType rightBracket;
    CellType delimiter;
    CellType bar[26];
    CellType line[32];
    CellType histogram[/*row:*/16][/*col:*/32];

    void displayVolume(unsigned char percentage)
    {
        if constexpr (std::is_same_v<CellType, Console::Char>)
        {
            for (auto& c : volume)
            {
                c = L' ';
            }

            char ansiValue[sizeof(volume) / sizeof(*volume)]{};
            std::to_chars(&ansiValue[0], &ansiValue[sizeof(ansiValue)], percentage);
            for (auto i = 0u; i < sizeof(ansiValue); ++i)
            {
                if (!ansiValue[i])
                {
                    break;
                }
                volume[i] = static_cast<wchar_t>(ansiValue[i]);
            }
        }
        else if constexpr (std::is_same_v<CellType, Console::Attr>)
        {
            for (auto i = 0u; i < std::size(bar); ++i)
            {
                const auto cellPercent = (static_cast<unsigned int>(i) * 100u) / (sizeof(bar) / sizeof(*bar));

                const unsigned short color = (cellPercent < 60)
                    ? BACKGROUND_GREEN
                    : ((cellPercent < 80)
                        ? (BACKGROUND_GREEN | BACKGROUND_RED)
                        : BACKGROUND_RED);
                
                if (cellPercent < percentage)
                {
                    bar[i] = color;
                }
                else
                {
                    bar[i] = 0;
                }
            }

            for (auto row = 0u; row < 16u; ++row)
            {
                for (auto col = 0u; col < 32u; ++col)
                {
                    histogram[row][col] = histogram[row][col + 1];
                }
            }


            for (auto row = 0u; row < 16u; ++row)
            {
                auto& cell = histogram[row][31];

                const auto cellPercent = (static_cast<unsigned int>(16u - (row + 1)) * 100u) / 16u;

                const unsigned short color = (cellPercent < 60)
                    ? BACKGROUND_GREEN
                    : ((cellPercent < 80)
                        ? (BACKGROUND_GREEN | BACKGROUND_RED)
                        : BACKGROUND_RED);

                if (cellPercent < percentage)
                {
                    cell = color;
                }
                else
                {
                    cell = 0;
                }
            }
        }
    }

    template <typename T = CellType, std::enable_if_t<std::is_same_v<T, Console::Char>, bool> = true>
    VUMeter()
        : leftBracket(L'[')
        , volume{ L' ', L' ', L' ' }
        , rightBracket(L']')
        , delimiter(L' ')
        , bar{}
        , line{}
        , histogram{}
    {
        for (auto& c : line)
        {
            c = L'=';
        }
    }

    template <typename T = CellType, std::enable_if_t<std::is_same_v<T, Console::Attr>, bool> = true>
    VUMeter()
    {
    }
};



class CoInit
{
private:
    HRESULT m_result;

public:
    CoInit() : m_result(CoInitializeEx(nullptr, COINITBASE_MULTITHREADED))
    {
    }

    CoInit(const CoInit&) = delete;
    CoInit(CoInit&&) = delete;

    ~CoInit()
    {
        if (SUCCEEDED(m_result))
        {
            CoUninitialize();
        }
    }

    CoInit& operator = (const CoInit&) = delete;
    CoInit& operator = (CoInit&&) = delete;

    HRESULT result() const
    {
        return m_result;
    }
};

// For a production-ready code you should use CComPtr from ATL,
// but  we wrote our own wrapper to avoid dependencies on ATL:
template <typename ComObj>
class ComPtr
{
private:
    ComObj* m_obj;

public:
    ComPtr() : m_obj(nullptr)
    {
    }

    ComPtr(const ComPtr&) = delete;
    ComPtr(ComPtr&&) = delete;

    ~ComPtr()
    {
        if (m_obj)
        {
            m_obj->Release();
        }
    }

    ComPtr& operator = (const ComPtr&) = delete;
    ComPtr& operator = (ComPtr&&) = delete;

    void** pointerTo()
    {
        return reinterpret_cast<void**>(&m_obj);
    }

    ComObj** typedPointerTo()
    {
        return &m_obj;
    }

    ComObj* operator -> ()
    {
        return m_obj;
    }

    const ComObj* operator -> () const
    {
        return m_obj;
    }

    ComObj& operator * ()
    {
        return *m_obj;
    }

    const ComObj& operator * () const
    {
        return *m_obj;
    }

    bool valid() const
    {
        return m_obj != nullptr;
    }
};


class KbLed
{
private:
    static inline std::atomic<unsigned char> s_numLockToggled = false;
    static inline std::atomic<unsigned char> s_capsLockToggled = false;
    static inline std::atomic<unsigned char> s_scrollLockToggled = false;

private:
    static DWORD __stdcall keyboardListenerThread(void*)
    {
        WNDCLASS wndClass{};
        wndClass.lpfnWndProc = [](HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) -> LRESULT
        {
            if (msg == WM_INPUT)
            {
                RAWINPUT input{};
                unsigned int rawInputSize = sizeof(input);
                const auto copiedBytes = GetRawInputData(reinterpret_cast<HRAWINPUT>(lParam), RID_INPUT, &input, &rawInputSize, sizeof(RAWINPUTHEADER));

                if (copiedBytes == 1)
                {
                    return 0;
                }

                if (input.header.dwType == RIM_TYPEKEYBOARD)
                {
                    const auto& key = input.data.keyboard;
                    if (key.Flags == RI_KEY_MAKE)
                    {
                        switch (key.VKey)
                        {
                        case VK_NUMLOCK:
                        {
                            s_numLockToggled.fetch_xor(1); // Atomic "not" (b = !b)
                            break;
                        }
                        case VK_CAPITAL:
                        {
                            s_capsLockToggled.fetch_xor(1);
                            break;
                        }
                        case VK_SCROLL:
                        {
                            s_scrollLockToggled.fetch_xor(1);
                            break;
                        }
                        }
                    }
                }
            }
            return DefWindowProcW(hwnd, msg, wParam, lParam);
        };
        wndClass.hInstance = GetModuleHandleW(nullptr);
        wndClass.lpszClassName = L"MessageOnlyClass-{7AA75419-C6BC-45A7-8A3A-F2CC20C2C8CC}";

        if (!RegisterClassW(&wndClass))
        {
            return GetLastError();
        }

        const HWND hwnd = CreateWindowW(wndClass.lpszClassName, L"VU-meter", 0, 0, 0, 0, 0, HWND_MESSAGE, 0, 0, 0);
        if (!hwnd)
        {
            UnregisterClassW(wndClass.lpszClassName, GetModuleHandleW(nullptr));
            return GetLastError();
        }

        RAWINPUTDEVICE inputDev{};

        inputDev.usUsagePage = HID_USAGE_PAGE_GENERIC;
        inputDev.usUsage = HID_USAGE_GENERIC_KEYBOARD;
        inputDev.dwFlags = RIDEV_INPUTSINK;
        inputDev.hwndTarget = hwnd;

        const bool registerStatus = !!RegisterRawInputDevices(&inputDev, 1, sizeof(inputDev));
        if (!registerStatus)
        {
            return GetLastError();
        }

        MSG msg{};
        while (GetMessageW(&msg, nullptr, 0, 0))
        {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        return 0;
    }

    static bool startKeyboardListener()
    {
        const HANDLE hThread = CreateThread(nullptr, 0, KbLed::keyboardListenerThread, nullptr, 0, nullptr);
        if (!hThread)
        {
            return false;
        }
        CloseHandle(hThread);
        return true;
    }

public:
    static bool init()
    {
        unsigned char keyboard[256]{};
        const bool kbStateStatus = !!GetKeyboardState(keyboard);
        if (!kbStateStatus)
        {
            return false;
        }

        s_numLockToggled = ((keyboard[VK_NUMLOCK] & 1ui8) != 0ui8);
        s_capsLockToggled = ((keyboard[VK_CAPITAL] & 1ui8) != 0ui8);
        s_scrollLockToggled = ((keyboard[VK_SCROLL] & 1ui8) != 0ui8);

        return startKeyboardListener();
    }

    static bool setLedState(bool numLock, bool capsLock, bool scrollLock)
    {
        const bool needToPressNumLock = (static_cast<unsigned char>(numLock) != s_numLockToggled);
        const bool needToPressCapsLock = (static_cast<unsigned char>(capsLock) != s_capsLockToggled);
        const bool needToPressScrollLock = (static_cast<unsigned char>(scrollLock) != s_scrollLockToggled);

        unsigned char inputsCount = 0;
        unsigned short requiredKeys[3]{};

        if (needToPressNumLock)
        {
            requiredKeys[inputsCount] = VK_NUMLOCK;
            ++inputsCount;
        }

        if (needToPressCapsLock)
        {
            requiredKeys[inputsCount] = VK_CAPITAL;
            ++inputsCount;
        }

        if (needToPressScrollLock)
        {
            requiredKeys[inputsCount] = VK_SCROLL;
            ++inputsCount;
        }

        if (!inputsCount)
        {
            return true;
        }

        INPUT inputs[6]{};

        for (auto i = 0u; i < inputsCount; ++i)
        {
            const auto inputIndex = i * 2;
            inputs[inputIndex].type = INPUT_KEYBOARD;
            inputs[inputIndex].ki.wVk = requiredKeys[i];
            inputs[inputIndex].ki.dwFlags = 0;
            inputs[inputIndex + 1].type = INPUT_KEYBOARD;
            inputs[inputIndex + 1].ki.wVk = requiredKeys[i];
            inputs[inputIndex + 1].ki.dwFlags = KEYEVENTF_KEYUP;
        }

        const UINT sentCount = SendInput(inputsCount * 2, inputs, sizeof(INPUT));
        return sentCount > 0;
    }
};


template <size_t sampleCount>
class VUAvg
{
private:
    unsigned int m_samples[sampleCount];
    unsigned int m_pos;
    unsigned int m_min;
    unsigned int m_max;

public:
    VUAvg() : m_samples{}, m_pos(0), m_min(0), m_max(0)
    {
    }

    unsigned int addSampleAndGetItsAvgPercentage(unsigned int sample)
    {
        m_samples[m_pos % sampleCount] = sample;
        ++m_pos;

        m_min = m_samples[0];
        m_max = m_samples[0];
        for (const auto& entry : m_samples)
        {
            if (entry < m_min)
            {
                m_min = entry;
            }

            if (entry > m_max)
            {
                m_max = entry;
            }
        }

        const auto relativeMax = m_max - m_min;
        const auto relativeSample = sample - m_min;

        if (relativeMax == 0)
        {
            return 0;
        }

        return (relativeSample * 100) / relativeMax;
    }
};


} // namespace




int wmain(/*int argc, wchar_t* argv[]*/)
{
    KbLed::init();

    g_console.setInteractable(false);
    g_console.setResizeable(false);
    g_console.resize(32, 18);
    g_console.setCursorView(false, 50);
    
    ScreenBuffer<VUMeter> screen(g_console);

    const CoInit coInit;
    if (FAILED(coInit.result()))
    {
        return coInit.result();
    }

    ComPtr<IMMDeviceEnumerator> deviceEnumerator;
    const HRESULT enumeratorStatus = CoCreateInstance(
        __uuidof(MMDeviceEnumerator),
        NULL,
        CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator),
        deviceEnumerator.pointerTo());

    if (FAILED(enumeratorStatus))
    {
        return enumeratorStatus;
    }

    ComPtr<IMMDevice> audioRenderer;
    const HRESULT getEndpointStatus = deviceEnumerator->GetDefaultAudioEndpoint(EDataFlow::eRender, ERole::eMultimedia, audioRenderer.typedPointerTo());
    if (FAILED(getEndpointStatus))
    {
        return getEndpointStatus;
    }

    ComPtr<IAudioMeterInformation> meter;
    const HRESULT activateStatus = audioRenderer->Activate(__uuidof(IAudioMeterInformation), CLSCTX_ALL, NULL, meter.pointerTo());
    if (FAILED(activateStatus))
    {
        return activateStatus;
    }


    VUAvg<32> averagePercentage; // Average by 32 samples

    while (true)
    {
        float peak = 0.0f;
        const HRESULT volumeStatus = meter->GetPeakValue(&peak);
        if (FAILED(volumeStatus))
        {
            return volumeStatus;
        }

        const auto percentage = static_cast<unsigned char>(peak * 100ui8);

        screen.chars().displayVolume(percentage);
        screen.attrs().displayVolume(percentage);
        screen.render();

        const auto percentageForLeds = averagePercentage.addSampleAndGetItsAvgPercentage(percentage);

        if (percentageForLeds < 25)
        {
            KbLed::setLedState(false, false, false);
        }
        else if (percentageForLeds < 50)
        {
            KbLed::setLedState(true, false, false);
        }
        else if (percentageForLeds < 75)
        {
            KbLed::setLedState(true, true, false);
        }
        else
        {
            KbLed::setLedState(true, true, true);
        }

        Sleep(40);
    }

    return 0;
}
