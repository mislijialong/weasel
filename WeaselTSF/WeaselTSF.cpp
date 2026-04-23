#include "stdafx.h"

#include <WeaselIPCData.h>
#include <thread>
#include <shellapi.h>
#include <tlhelp32.h>
#include "WeaselTSF.h"
#include "CandidateList.h"
#include "LanguageBar.h"
#include "Compartment.h"
#include "ResponseParser.h"

static void error_message(const WCHAR* msg) {
  static DWORD next_tick = 0;
  DWORD now = GetTickCount();
  if (now > next_tick) {
    next_tick = now + 10000;  // (ms)
    MessageBox(NULL, msg, get_weasel_ime_name().c_str(), MB_ICONERROR | MB_OK);
  }
}

namespace {
constexpr UINT kInlineAiPollIntervalMs = 40;

void InlineAiTsfDebug(const std::wstring& message) {
  OutputDebugStringW((L"inline_ai_tsf: " + message + L"\n").c_str());
}
}  // namespace

WeaselTSF::WeaselTSF() {
  _cRef = 1;

  _dwThreadMgrEventSinkCookie = TF_INVALID_COOKIE;

  _dwTextEditSinkCookie = TF_INVALID_COOKIE;
  _dwTextLayoutSinkCookie = TF_INVALID_COOKIE;
  _dwThreadFocusSinkCookie = TF_INVALID_COOKIE;
  _fTestKeyDownPending = FALSE;
  _fTestKeyUpPending = FALSE;

  _fCUASWorkaroundTested = _fCUASWorkaroundEnabled = FALSE;

  _cand = new CCandidateList(this);

  DllAddRef();
}

WeaselTSF::~WeaselTSF() {
  _CancelInlineAiSync();
  DllRelease();
}

STDAPI WeaselTSF::QueryInterface(REFIID riid, void** ppvObject) {
  if (ppvObject == NULL)
    return E_INVALIDARG;

  *ppvObject = NULL;

  if (IsEqualIID(riid, IID_IUnknown) ||
      IsEqualIID(riid, IID_ITfTextInputProcessor))
    *ppvObject = (ITfTextInputProcessor*)this;
  else if (IsEqualIID(riid, IID_ITfTextInputProcessorEx))
    *ppvObject = (ITfTextInputProcessorEx*)this;
  else if (IsEqualIID(riid, IID_ITfThreadMgrEventSink))
    *ppvObject = (ITfThreadMgrEventSink*)this;
  else if (IsEqualIID(riid, IID_ITfTextEditSink))
    *ppvObject = (ITfTextEditSink*)this;
  else if (IsEqualIID(riid, IID_ITfTextLayoutSink))
    *ppvObject = (ITfTextLayoutSink*)this;
  else if (IsEqualIID(riid, IID_ITfKeyEventSink))
    *ppvObject = (ITfKeyEventSink*)this;
  else if (IsEqualIID(riid, IID_ITfCompositionSink))
    *ppvObject = (ITfCompositionSink*)this;
  else if (IsEqualIID(riid, IID_ITfEditSession))
    *ppvObject = (ITfEditSession*)this;
  else if (IsEqualIID(riid, IID_ITfThreadFocusSink))
    *ppvObject = (ITfThreadFocusSink*)this;
  else if (IsEqualIID(riid, IID_ITfDisplayAttributeProvider))
    *ppvObject = (ITfDisplayAttributeProvider*)this;

  if (*ppvObject) {
    AddRef();
    return S_OK;
  }
  return E_NOINTERFACE;
}

STDAPI_(ULONG) WeaselTSF::AddRef() {
  return ++_cRef;
}

STDAPI_(ULONG) WeaselTSF::Release() {
  LONG cr = --_cRef;

  assert(_cRef >= 0);

  if (_cRef == 0)
    delete this;

  return cr;
}

STDAPI WeaselTSF::Activate(ITfThreadMgr* pThreadMgr, TfClientId tfClientId) {
  return ActivateEx(pThreadMgr, tfClientId, 0U);
}

STDAPI WeaselTSF::Deactivate() {
  _CancelInlineAiSync();
  m_client.EndSession();

  _InitTextEditSink(com_ptr<ITfDocumentMgr>());

  _UninitThreadMgrEventSink();

  _UninitKeyEventSink();
  _UninitPreservedKey();

  _UninitLanguageBar();

  _UninitCompartment();

  _UninitThreadMgrEventSink();

  _pThreadMgr = NULL;

  _tfClientId = TF_CLIENTID_NULL;

  _cand->DestroyAll();

  return S_OK;
}

VOID CALLBACK WeaselTSF::_InlineAiPollTimerProc(_In_ HWND hwnd,
                                                _In_ UINT uMsg,
                                                _In_ UINT_PTR idEvent,
                                                _In_ DWORD dwTime) {
  WeaselTSF* self = reinterpret_cast<WeaselTSF*>(idEvent);
  if (!self) {
    return;
  }
  if (!self->_SyncInlineAiState()) {
    self->_CancelInlineAiSync();
  }
}

void WeaselTSF::_ScheduleInlineAiSync() {
  if (_inlineAiPollTimer != 0) {
    return;
  }
  _inlineAiPollTimer = reinterpret_cast<UINT_PTR>(this);
  SetTimer(nullptr, _inlineAiPollTimer, kInlineAiPollIntervalMs,
           &WeaselTSF::_InlineAiPollTimerProc);
}

void WeaselTSF::_CancelInlineAiSync() {
  if (_inlineAiPollTimer == 0) {
    return;
  }
  KillTimer(nullptr, _inlineAiPollTimer);
  _inlineAiPollTimer = 0;
}

bool WeaselTSF::_SyncInlineAiState() {
  if (!_pEditSessionContext) {
    return false;
  }
  if (!m_client.Echo()) {
    return false;
  }
  if (!m_client.SyncSession()) {
    return false;
  }
  weasel::Config config;
  auto context = std::make_shared<weasel::Context>();
  std::wstring commit;
  weasel::ResponseParser parser(&commit, context.get(), &_status, &config,
                                &_cand->style());
  const bool ok = m_client.GetResponseData(std::ref(parser));
  if (!ok) {
    InlineAiTsfDebug(L"sync_session response missing");
    return false;
  }
  InlineAiTsfDebug(L"sync_session ok composing=" +
                   std::to_wstring(_status.composing) +
                   L" inline_preedit=" +
                   std::to_wstring(config.inline_preedit) +
                   L" inline_ai_requesting=" +
                   std::to_wstring(config.inline_ai_requesting) +
                   L" preedit=" + context->preedit.str +
                   L" commit=" + commit);
  _UpdateLanguageBar(_status);
  if (!commit.empty()) {
    if (!_IsComposing()) {
      _StartComposition(_pEditSessionContext,
                        _fCUASWorkaroundEnabled && !config.inline_preedit);
    }
    _InsertText(_pEditSessionContext, commit);
    _EndComposition(_pEditSessionContext, false);
    _committed = TRUE;
  }
  if (_status.composing && !_IsComposing()) {
    _StartComposition(_pEditSessionContext,
                      _fCUASWorkaroundEnabled && !config.inline_preedit);
  } else if (!_status.composing && _IsComposing()) {
    _EndComposition(_pEditSessionContext, true);
  }
  if (_IsComposing() && config.inline_preedit) {
    _ShowInlinePreedit(_pEditSessionContext, context);
  }
  _UpdateCompositionWindow(_pEditSessionContext);
  _UpdateUI(*context, _status);
  return config.inline_ai_requesting;
}

STDAPI WeaselTSF::ActivateEx(ITfThreadMgr* pThreadMgr,
                             TfClientId tfClientId,
                             DWORD dwFlags) {
  com_ptr<ITfDocumentMgr> pDocMgrFocus;
  _activateFlags = dwFlags;

  _pThreadMgr = pThreadMgr;
  _tfClientId = tfClientId;

  if (!_InitThreadMgrEventSink())
    goto ExitError;

  if ((_pThreadMgr->GetFocus(&pDocMgrFocus) == S_OK) &&
      (pDocMgrFocus != NULL)) {
    _InitTextEditSink(pDocMgrFocus);
  }

  if (!_InitKeyEventSink())
    goto ExitError;

  // if (!_InitDisplayAttributeGuidAtom())
  //	goto ExitError;
  //	some app might init failed because it not provide DisplayAttributeInfo,
  // like some opengl stuff
  _InitDisplayAttributeGuidAtom();

  if (!_InitPreservedKey())
    goto ExitError;

  if (!_InitLanguageBar())
    goto ExitError;

  if (!_IsKeyboardOpen())
    _SetKeyboardOpen(TRUE);

  if (!_InitCompartment())
    goto ExitError;
  if (!_InitThreadFocusSink())
    goto ExitError;

  _EnsureServerConnected();

  return S_OK;

ExitError:
  Deactivate();
  return E_FAIL;
}

STDMETHODIMP WeaselTSF::OnSetThreadFocus() {
  std::wstring _ToggleImeOnOpenClose{};
  RegGetStringValue(HKEY_CURRENT_USER, L"Software\\Rime\\weasel",
                    L"ToggleImeOnOpenClose", _ToggleImeOnOpenClose);
  _isToOpenClose = (_ToggleImeOnOpenClose == L"yes");
  if (m_client.Echo()) {
    m_client.ProcessKeyEvent(0);
    weasel::ResponseParser parser(NULL, NULL, &_status, NULL, &_cand->style());
    bool ok = m_client.GetResponseData(std::ref(parser));
    if (ok)
      _UpdateLanguageBar(_status);
  }
  return S_OK;
}
STDMETHODIMP WeaselTSF::OnKillThreadFocus() {
  _AbortComposition();
  return S_OK;
}
BOOL WeaselTSF::_InitThreadFocusSink() {
  com_ptr<ITfSource> pSource;
  if (FAILED(_pThreadMgr->QueryInterface(&pSource)))
    return FALSE;
  if (FAILED(pSource->AdviseSink(IID_ITfThreadFocusSink,
                                 (ITfThreadFocusSink*)this,
                                 &_dwThreadFocusSinkCookie)))
    return FALSE;
  return TRUE;
}
void WeaselTSF::_UninitThreadFocusSink() {
  com_ptr<ITfSource> pSource;
  if (FAILED(_pThreadMgr->QueryInterface(&pSource)))
    return;
  if (FAILED(pSource->UnadviseSink(_dwThreadFocusSinkCookie)))
    return;
}

STDMETHODIMP WeaselTSF::OnActivated(REFCLSID clsid,
                                    REFGUID guidProfile,
                                    BOOL isActivated) {
  if (!IsEqualCLSID(clsid, c_clsidTextService)) {
    return S_OK;
  }

  if (isActivated) {
    _ShowLanguageBar(TRUE);
    _UpdateLanguageBar(_status);
  } else {
    _DeleteCandidateList();
    _ShowLanguageBar(FALSE);
  }
  return S_OK;
}

void WeaselTSF::_Reconnect() {
  m_client.Disconnect();
  m_client.Connect(NULL);
  m_client.StartSession();
  weasel::ResponseParser parser(NULL, NULL, &_status, NULL, &_cand->style());
  bool ok = m_client.GetResponseData(std::ref(parser));
  if (ok) {
    _UpdateLanguageBar(_status);
  }
}

static unsigned int retry = 0;

bool WeaselTSF::_EnsureServerConnected() {
  if (!m_client.Echo()) {
    _Reconnect();
    retry++;
    if (retry >= 6) {
      HANDLE hMutex = CreateMutex(NULL, TRUE, L"WeaselDeployerExclusiveMutex");
      const auto count_server_process = []() -> int {
        int count = 0;
        HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (snap == INVALID_HANDLE_VALUE)
          return 0;
        PROCESSENTRY32 pe;
        pe.dwSize = sizeof(pe);
        if (Process32First(snap, &pe)) {
          do {
            if (_wcsicmp(pe.szExeFile, L"WeaselServer.exe") == 0)
              count++;
          } while (Process32Next(snap, &pe));
        }
        CloseHandle(snap);
        return count;
      };
      if (!m_client.Echo() && GetLastError() != ERROR_ALREADY_EXISTS &&
          !count_server_process()) {
        std::wstring dir = _GetRootDir();
        std::thread th([dir, this]() {
          ShellExecuteW(NULL, L"open", (dir + L"\\start_service.bat").c_str(),
                        NULL, dir.c_str(), SW_HIDE);
          // wait 500ms, then reconnect
          std::this_thread::sleep_for(std::chrono::milliseconds(500));
          _Reconnect();
        });
        th.detach();
      }
      if (hMutex) {
        CloseHandle(hMutex);
      }
      retry = 0;
    }
    return (m_client.Echo() != 0);
  } else {
    return true;
  }
}
