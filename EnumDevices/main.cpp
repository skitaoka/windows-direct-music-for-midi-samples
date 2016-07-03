#include <sstream>

#include <wrl.h>
#include <dmusicc.h>

namespace {
std::wostream& operator << (std::wostream& out, DMUS_PORTCAPS const& caps) {
  out << L"GUID: "
    << caps.guidPort.Data1 << '-'
    << caps.guidPort.Data2 << '-'
    << caps.guidPort.Data3 << '-'
    << caps.guidPort.Data4;

  switch (caps.dwClass) {
  case DMUS_PC_INPUTCLASS:
    out << L" IN";
    break;
  case DMUS_PC_OUTPUTCLASS:
    out << L" OUT";
    break;
  default:
    out << L" UNKNOWN";
    break;
  }

  switch (caps.dwType) {
  case DMUS_PORT_WINMM_DRIVER:
    out << L" WINMM";
    break;
  case DMUS_PORT_USER_MODE_SYNTH:
    out << L" SYNTH";
    break;
  case DMUS_PORT_KERNEL_MODE:
    out << L" WDM";
    break;
  default:
    out << L" UNKNOWN";
    break;
  }

  out
    << L" MemorySize=" << caps.dwMemorySize
    << L" MaxChannelGroups=" << caps.dwMaxChannelGroups
    << L" MaxVoices=" << caps.dwMaxVoices
    << L" MaxAudioChannels=" << caps.dwMaxAudioChannels;
  if (caps.dwEffectFlags != DMUS_EFFECT_NONE) {
    out << L" EffectFlags=";
    if (caps.dwEffectFlags & DMUS_EFFECT_REVERB) {
      out << L'R';
    }
    if (caps.dwEffectFlags & DMUS_EFFECT_CHORUS) {
      out << L'C';
    }
    if (caps.dwEffectFlags & DMUS_EFFECT_DELAY) {
      out << L'D';
    }
  }

  out << L" \"" << caps.wszDescription << L"\"";

  return out;
}
}

int main() {
  if SUCCEEDED(::CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED)) {
    {
      Microsoft::WRL::ComPtr<IDirectMusic> music;
      if SUCCEEDED(::CoCreateInstance(CLSID_DirectMusic, nullptr,
                                      CLSCTX_INPROC, IID_IDirectMusic,
                                      reinterpret_cast<void**>(music.ReleaseAndGetAddressOf()))) {

        std::wostringstream out;

        DMUS_PORTCAPS caps = {sizeof(caps)};
        for (DWORD i = 0;; ++i) {
          HRESULT const hr = music->EnumPort(i, &caps);
          if (hr == S_FALSE) {
            break;
          }
          if FAILED(hr) {
            continue;
          }
          out << L"[" << i << L"]: " << caps << L'\n';
        }

        std::wstring const text = out.str();
        ::MessageBox(HWND_DESKTOP, text.c_str(), L"Ports", MB_OK);
      }
    }
    ::CoUninitialize();
  }

  return 0;
}