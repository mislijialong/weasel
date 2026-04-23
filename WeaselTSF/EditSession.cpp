#include "stdafx.h"
#include "WeaselTSF.h"
#include "CandidateList.h"
#include "ResponseParser.h"

namespace {
void InlineAiTsfDebug(const std::wstring& message) {
  OutputDebugStringW((L"inline_ai_tsf: " + message + L"\n").c_str());
}
}  // namespace

STDAPI WeaselTSF::DoEditSession(TfEditCookie ec) {
  // get commit string from server
  std::wstring commit;
  weasel::Config config;
  auto context = std::make_shared<weasel::Context>();
  weasel::ResponseParser parser(&commit, context.get(), &_status, &config,
                                &_cand->style());

  bool ok = m_client.GetResponseData(std::ref(parser));
  InlineAiTsfDebug(L"edit_session ok=" + std::to_wstring(ok) +
                   L" composing=" + std::to_wstring(_status.composing) +
                   L" inline_preedit=" +
                   std::to_wstring(config.inline_preedit) +
                   L" inline_ai_requesting=" +
                   std::to_wstring(config.inline_ai_requesting) +
                   L" preedit=" + context->preedit.str +
                   L" commit=" + commit);

  _UpdateLanguageBar(_status);

  if (ok) {
    if (!commit.empty()) {
      // For auto-selecting, commit and preedit can both exist.
      // Commit and close the original composition first.
      if (!_IsComposing()) {
        _StartComposition(_pEditSessionContext,
                          _fCUASWorkaroundEnabled && !config.inline_preedit);
      }
      _InsertText(_pEditSessionContext, commit);
      _EndComposition(_pEditSessionContext, false);
      _committed = TRUE;
    } else {
      _committed = FALSE;
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
  }

  _UpdateUI(*context, _status);
  if (config.inline_ai_requesting) {
    _ScheduleInlineAiSync();
  } else {
    _CancelInlineAiSync();
  }

  return TRUE;
}
