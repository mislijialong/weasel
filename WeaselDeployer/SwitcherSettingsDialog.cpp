#include "stdafx.h"

#include "SwitcherSettingsDialog.h"

#include "Configurator.h"
#include "WeaselDeployer.h"

#include <WeaselUtility.h>
#include <algorithm>
#include <set>
#include <rime_levers_api.h>

namespace {

UINT BuildKeyboardModifierMask(const BYTE key_state[256]) {
  if (!key_state) {
    return 0;
  }

  UINT modifiers = 0;
  if ((key_state[VK_CONTROL] & 0x80) != 0 ||
      (key_state[VK_LCONTROL] & 0x80) != 0 ||
      (key_state[VK_RCONTROL] & 0x80) != 0) {
    modifiers |= ibus::CONTROL_MASK;
  }
  if ((key_state[VK_MENU] & 0x80) != 0 ||
      (key_state[VK_LMENU] & 0x80) != 0 ||
      (key_state[VK_RMENU] & 0x80) != 0) {
    modifiers |= ibus::ALT_MASK;
  }
  if ((key_state[VK_SHIFT] & 0x80) != 0 ||
      (key_state[VK_LSHIFT] & 0x80) != 0 ||
      (key_state[VK_RSHIFT] & 0x80) != 0) {
    modifiers |= ibus::SHIFT_MASK;
  }
  if ((key_state[VK_LWIN] & 0x80) != 0 || (key_state[VK_RWIN] & 0x80) != 0) {
    modifiers |= ibus::SUPER_MASK;
  }
  return modifiers;
}

void MarkVirtualKeyState(UINT vkey, bool down, BYTE key_state[256]) {
  if (!key_state || vkey >= 256) {
    return;
  }

  const BYTE value = down ? static_cast<BYTE>(key_state[vkey] | 0x80)
                          : static_cast<BYTE>(key_state[vkey] & ~0x80);
  key_state[vkey] = value;

  switch (vkey) {
    case VK_SHIFT:
    case VK_LSHIFT:
    case VK_RSHIFT:
      key_state[VK_SHIFT] = value;
      break;
    case VK_CONTROL:
    case VK_LCONTROL:
    case VK_RCONTROL:
      key_state[VK_CONTROL] = value;
      break;
    case VK_MENU:
    case VK_LMENU:
    case VK_RMENU:
      key_state[VK_MENU] = value;
      break;
    default:
      break;
  }
}

constexpr const wchar_t* kDefaultAiHotkey = L"Control+3";
constexpr const wchar_t* kDefaultInstructionLookupPrefix = L"sS";

std::wstring TrimInstructionLookupPrefixText(const std::wstring& text) {
  size_t begin = 0;
  while (begin < text.size() &&
         iswspace(static_cast<wint_t>(text[begin])) != 0) {
    ++begin;
  }
  size_t end = text.size();
  while (end > begin &&
         iswspace(static_cast<wint_t>(text[end - 1])) != 0) {
    --end;
  }
  return text.substr(begin, end - begin);
}

bool IsValidInstructionLookupPrefixText(const std::wstring& text) {
  if (text.empty()) {
    return true;
  }
  for (wchar_t ch : text) {
    if ((ch < L'A' || ch > L'Z') && (ch < L'a' || ch > L'z')) {
      return false;
    }
  }
  return true;
}

std::string BuildInstructionLookupRecognizerPattern(
    const std::wstring& prefix_text) {
  return "^" + wtou8(prefix_text) + "[a-z]+$";
}

}  // namespace

void AIAssistantHotkeyEdit::InitializeFromConfigText(
    const std::wstring& raw_text) {
  binding_ = AiHotkeyBinding();
  model_text_ = raw_text;
  canonical_config_text_.clear();
  has_parse_error_ = false;
  is_empty_ = raw_text.empty();
  preserve_raw_text_ = false;
  showing_partial_capture_ = false;
  partial_modifiers_ = 0;

  if (raw_text.empty()) {
    display_text_.clear();
    SetDisplayText(display_text_);
    NotifyParentStateChanged();
    return;
  }

  AiHotkeyBinding parsed_binding;
  if (TryParseAiHotkeyConfig(wtou8(raw_text), &parsed_binding)) {
    binding_ = parsed_binding;
    canonical_config_text_ = FormatAiHotkeyForConfig(binding_);
    const AiHotkeyValidationResult validation =
        ValidateAiHotkeyBinding(binding_);
    if (validation.ok) {
      model_text_ = u8tow(canonical_config_text_);
      display_text_ = FormatAiHotkeyForDisplay(binding_);
    } else {
      preserve_raw_text_ = true;
      display_text_ = raw_text;
    }
  } else {
    has_parse_error_ = true;
    preserve_raw_text_ = true;
    display_text_ = raw_text;
  }

  SetDisplayText(display_text_);
  NotifyParentStateChanged();
}

std::wstring AIAssistantHotkeyEdit::GetDisplayText() const {
  return display_text_;
}

LRESULT AIAssistantHotkeyEdit::OnGetDlgCode(UINT,
                                            WPARAM,
                                            LPARAM,
                                            BOOL&) {
  return DLGC_WANTALLKEYS | DLGC_WANTCHARS;
}

LRESULT AIAssistantHotkeyEdit::OnFocus(UINT u_msg,
                                       WPARAM,
                                       LPARAM,
                                       BOOL& b_handled) {
  b_handled = FALSE;
  if (u_msg == WM_SETFOCUS) {
    SetSel(0, -1);
  } else if (showing_partial_capture_) {
    RestoreCommittedDisplay();
    return 0;
  }
  NotifyParentStateChanged();
  return 0;
}

LRESULT AIAssistantHotkeyEdit::OnKeyDown(UINT,
                                         WPARAM w_param,
                                         LPARAM l_param,
                                         BOOL&) {
  BYTE key_state[256] = {0};
  ::GetKeyboardState(key_state);
  MarkVirtualKeyState(static_cast<UINT>(w_param), true, key_state);

  const UINT modifiers = BuildKeyboardModifierMask(key_state);
  const UINT vkey = static_cast<UINT>(w_param);
  if (IsAiHotkeyModifierVirtualKey(vkey)) {
    UpdatePartialCapture(modifiers);
    return 0;
  }

  if ((vkey == VK_BACK || vkey == VK_DELETE) && modifiers == 0) {
    ClearBinding();
    return 0;
  }

  AiHotkeyBinding binding;
  const KeyInfo key_info(l_param);
  if (TryBuildAiHotkeyFromWinKey(vkey, key_info.isExtended != 0, key_state,
                                 &binding)) {
    CommitBinding(binding);
  }
  return 0;
}

LRESULT AIAssistantHotkeyEdit::OnKeyUp(UINT,
                                       WPARAM w_param,
                                       LPARAM,
                                       BOOL&) {
  const UINT vkey = static_cast<UINT>(w_param);
  if (!IsAiHotkeyModifierVirtualKey(vkey) || !showing_partial_capture_) {
    return 0;
  }

  BYTE key_state[256] = {0};
  ::GetKeyboardState(key_state);
  MarkVirtualKeyState(vkey, false, key_state);
  const UINT modifiers = BuildKeyboardModifierMask(key_state);
  if (modifiers == 0) {
    RestoreCommittedDisplay();
  } else {
    UpdatePartialCapture(modifiers);
  }
  return 0;
}

LRESULT AIAssistantHotkeyEdit::OnBlockedInput(UINT,
                                              WPARAM,
                                              LPARAM,
                                              BOOL&) {
  return 0;
}

void AIAssistantHotkeyEdit::ClearBinding() {
  binding_ = AiHotkeyBinding();
  model_text_.clear();
  display_text_.clear();
  canonical_config_text_.clear();
  has_parse_error_ = false;
  is_empty_ = true;
  preserve_raw_text_ = false;
  showing_partial_capture_ = false;
  partial_modifiers_ = 0;
  SetDisplayText(display_text_);
  NotifyParentStateChanged();
}

void AIAssistantHotkeyEdit::CommitBinding(const AiHotkeyBinding& binding) {
  binding_ = binding;
  canonical_config_text_ = FormatAiHotkeyForConfig(binding_);
  model_text_ = u8tow(canonical_config_text_);
  display_text_ = FormatAiHotkeyForDisplay(binding_);
  has_parse_error_ = false;
  is_empty_ = false;
  preserve_raw_text_ = false;
  showing_partial_capture_ = false;
  partial_modifiers_ = 0;
  SetDisplayText(display_text_);
  NotifyParentStateChanged();
}

void AIAssistantHotkeyEdit::NotifyParentStateChanged() const {
  if (GetParent().IsWindow()) {
    GetParent().SendMessageW(WM_AI_HOTKEY_STATE_CHANGED, 0, 0);
  }
}

void AIAssistantHotkeyEdit::RestoreCommittedDisplay() {
  showing_partial_capture_ = false;
  partial_modifiers_ = 0;
  if (preserve_raw_text_) {
    display_text_ = model_text_;
  } else if (is_empty_) {
    display_text_.clear();
  } else {
    display_text_ = FormatAiHotkeyForDisplay(binding_);
  }
  SetDisplayText(display_text_);
  NotifyParentStateChanged();
}

void AIAssistantHotkeyEdit::SetDisplayText(const std::wstring& text) {
  display_text_ = text;
  if (!IsWindow()) {
    return;
  }

  CString current_text;
  GetWindowTextW(current_text);
  if (text != std::wstring(current_text.GetString())) {
    SetWindowTextW(text.c_str());
  }
  if (::GetFocus() == m_hWnd) {
    SetSel(0, -1);
  }
}

void AIAssistantHotkeyEdit::UpdatePartialCapture(UINT modifiers) {
  showing_partial_capture_ = modifiers != 0;
  partial_modifiers_ = modifiers;
  display_text_ = FormatAiHotkeyModifiersForDisplay(modifiers, true);
  SetDisplayText(display_text_);
  NotifyParentStateChanged();
}

SwitcherSettingsDialog::SwitcherSettingsDialog(RimeSwitcherSettings* settings)
    : settings_(settings),
      ai_settings_(nullptr),
      instruction_lookup_schema_settings_(nullptr),
      loaded_(false),
      schema_modified_(false),
      ai_hotkey_modified_(false),
      instruction_lookup_prefix_modified_(false),
      ai_settings_saved_(false),
      initial_ai_hotkey_model_text_(),
      initial_instruction_lookup_prefix_() {
  api_ = (RimeLeversApi*)rime_get_api()->find_module("levers")->get_api();
  if (api_) {
    ai_settings_ =
        api_->custom_settings_init("weasel", "Weasel::SwitcherSettingsDialog");
    instruction_lookup_schema_settings_ =
        api_->custom_settings_init("rime_ice",
                                   "Weasel::SwitcherSettingsDialog");
  }
}

SwitcherSettingsDialog::~SwitcherSettingsDialog() {
  if (api_ && instruction_lookup_schema_settings_) {
    api_->custom_settings_destroy(instruction_lookup_schema_settings_);
    instruction_lookup_schema_settings_ = nullptr;
  }
  if (api_ && ai_settings_) {
    api_->custom_settings_destroy(ai_settings_);
    ai_settings_ = nullptr;
  }
}

void SwitcherSettingsDialog::Populate() {
  if (!settings_) {
    return;
  }

  RimeSchemaList available = {0};
  api_->get_available_schema_list(settings_, &available);
  RimeSchemaList selected = {0};
  api_->get_selected_schema_list(settings_, &selected);
  schema_list_.DeleteAllItems();
  int index = 0;
  std::set<RimeSchemaInfo*> recruited;
  for (size_t i = 0; i < selected.size; ++i) {
    const char* schema_id = selected.list[i].schema_id;
    for (size_t j = 0; j < available.size; ++j) {
      RimeSchemaListItem& item(available.list[j]);
      RimeSchemaInfo* info = (RimeSchemaInfo*)item.reserved;
      if (!strcmp(item.schema_id, schema_id) &&
          recruited.find(info) == recruited.end()) {
        recruited.insert(info);
        std::wstring item_text = u8tow(item.name);
        schema_list_.AddItem(index, 0, item_text.c_str());
        schema_list_.SetItemData(index, (DWORD_PTR)info);
        schema_list_.SetCheckState(index, TRUE);
        ++index;
        break;
      }
    }
  }
  for (size_t i = 0; i < available.size; ++i) {
    RimeSchemaListItem& item(available.list[i]);
    RimeSchemaInfo* info = (RimeSchemaInfo*)item.reserved;
    if (recruited.find(info) == recruited.end()) {
      recruited.insert(info);
      std::wstring item_text = u8tow(item.name);
      schema_list_.AddItem(index, 0, item_text.c_str());
      schema_list_.SetItemData(index, (DWORD_PTR)info);
      ++index;
    }
  }

  const std::wstring ai_hotkey = LoadAIAssistantTriggerHotkey();
  ai_hotkey_.InitializeFromConfigText(ai_hotkey);
  initial_ai_hotkey_model_text_ = ai_hotkey_.GetModelText();
  const std::wstring instruction_lookup_prefix = LoadInstructionLookupPrefix();
  instruction_lookup_prefix_.SetWindowTextW(instruction_lookup_prefix.c_str());
  initial_instruction_lookup_prefix_ = instruction_lookup_prefix;

  loaded_ = true;
  schema_modified_ = false;
  ai_hotkey_modified_ = false;
  instruction_lookup_prefix_modified_ = false;
  ai_settings_saved_ = false;
  UpdateAIAssistantHotkeyUi();
}

void SwitcherSettingsDialog::ShowDetails(RimeSchemaInfo* info) {
  if (!info) {
    return;
  }

  std::string details;
  if (const char* name = api_->get_schema_name(info)) {
    details += name;
  }
  if (const char* author = api_->get_schema_author(info)) {
    (details += "\n\n") += author;
  }
  if (const char* description = api_->get_schema_description(info)) {
    (details += "\n\n") += description;
  }
  description_.SetWindowTextW(u8tow(details.c_str()).c_str());
}

std::wstring SwitcherSettingsDialog::BuildShortcutOverviewText(
    const std::wstring& ai_hotkey) const {
  std::wstring text = LoadStringResource(IDS_STR_AI_HOTKEY_OVERVIEW_TITLE);
  text += L"\r\n";
  text += LoadStringResource(IDS_STR_AI_HOTKEY_OVERVIEW_FIXED_IME);
  text += L"\r\n";
  text += LoadStringResource(IDS_STR_AI_HOTKEY_OVERVIEW_AI_PREFIX);
  text += ai_hotkey.empty() ? kDefaultAiHotkey : ai_hotkey;
  text += L"\r\n";
  text += LoadStringResource(IDS_STR_AI_HOTKEY_OVERVIEW_FIXED_ESC);
  text += L"\r\n\r\n";
  text += LoadStringResource(IDS_STR_AI_HOTKEY_OVERVIEW_NOTE);
  return text;
}

std::wstring SwitcherSettingsDialog::GetCurrentAIAssistantHotkeyHintText() const {
  if (ai_hotkey_.IsShowingPartialCapture()) {
    return LoadStringResource(IDS_STR_AI_HOTKEY_HINT_PARTIAL);
  }
  if (ai_hotkey_.HasParseError()) {
    return LoadStringResource(IDS_STR_AI_HOTKEY_HINT_PARSE_ERROR);
  }
  if (ai_hotkey_.IsEmpty()) {
    return LoadStringResource(IDS_STR_AI_HOTKEY_HINT_EMPTY);
  }
  if (ai_hotkey_.HasRecognizedBinding()) {
    const AiHotkeyValidationResult validation =
        ValidateAiHotkeyBinding(ai_hotkey_.GetBinding());
    if (validation.reserved_conflict) {
      return LoadStringResource(IDS_STR_AI_HOTKEY_ERR_RESERVED);
    }
    if (!validation.ok) {
      return LoadStringResource(IDS_STR_AI_HOTKEY_ERR_REQUIRE_MOD);
    }
  }
  return LoadStringResource(IDS_STR_AI_HOTKEY_HINT_DEFAULT);
}

bool SwitcherSettingsDialog::IsCurrentAIAssistantHotkeySavable() const {
  if (ai_hotkey_.HasParseError()) {
    return false;
  }
  if (ai_hotkey_.IsEmpty()) {
    return true;
  }
  if (!ai_hotkey_.HasRecognizedBinding()) {
    return false;
  }
  return ValidateAiHotkeyBinding(ai_hotkey_.GetBinding()).ok;
}

std::wstring SwitcherSettingsDialog::LoadAIAssistantTriggerHotkey() const {
  if (!api_ || !ai_settings_) {
    return kDefaultAiHotkey;
  }
  if (!api_->load_settings(ai_settings_)) {
    return kDefaultAiHotkey;
  }
  RimeConfig config = {0};
  if (!api_->settings_get_config(ai_settings_, &config)) {
    return kDefaultAiHotkey;
  }
  const char* value =
      rime_get_api()->config_get_cstring(&config, "ai_assistant/trigger_hotkey");
  if (!value || !*value) {
    return kDefaultAiHotkey;
  }
  return u8tow(value);
}

std::wstring SwitcherSettingsDialog::LoadInstructionLookupPrefix() const {
  if (!api_ || !ai_settings_) {
    return kDefaultInstructionLookupPrefix;
  }
  if (!api_->load_settings(ai_settings_)) {
    return kDefaultInstructionLookupPrefix;
  }
  RimeConfig config = {0};
  if (!api_->settings_get_config(ai_settings_, &config)) {
    return kDefaultInstructionLookupPrefix;
  }
  const char* value = rime_get_api()->config_get_cstring(
      &config, "ai_assistant/instruction_lookup_prefix");
  if (!value || !*value) {
    return kDefaultInstructionLookupPrefix;
  }
  const std::wstring trimmed = TrimInstructionLookupPrefixText(u8tow(value));
  if (trimmed.empty() || !IsValidInstructionLookupPrefixText(trimmed)) {
    return kDefaultInstructionLookupPrefix;
  }
  return trimmed;
}

std::wstring SwitcherSettingsDialog::GetCurrentInstructionLookupPrefix() const {
  CString text;
  instruction_lookup_prefix_.GetWindowTextW(text);
  return TrimInstructionLookupPrefixText(std::wstring(text.GetString()));
}

bool SwitcherSettingsDialog::IsCurrentInstructionLookupPrefixSavable() const {
  return IsValidInstructionLookupPrefixText(GetCurrentInstructionLookupPrefix());
}

std::wstring SwitcherSettingsDialog::LoadStringResource(UINT id) const {
  CString text;
  text.LoadStringW(id);
  return std::wstring(text.GetString());
}

void SwitcherSettingsDialog::UpdateAIAssistantHotkeyUi() {
  if (!loaded_) {
    return;
  }

  ai_hotkey_modified_ =
      ai_hotkey_.GetModelText() != initial_ai_hotkey_model_text_;
  instruction_lookup_prefix_modified_ =
      GetCurrentInstructionLookupPrefix() != initial_instruction_lookup_prefix_;
  ai_hotkey_hint_.SetWindowTextW(GetCurrentAIAssistantHotkeyHintText().c_str());

  if (schema_list_.GetNextItem(-1, LVNI_SELECTED) < 0) {
    description_.SetWindowTextW(
        BuildShortcutOverviewText(ai_hotkey_.GetDisplayText()).c_str());
  }

  const bool allow_ok =
      !(ai_hotkey_modified_ && !IsCurrentAIAssistantHotkeySavable()) &&
      !(instruction_lookup_prefix_modified_ &&
        !IsCurrentInstructionLookupPrefixSavable());
  if (CWindow ok_button = GetDlgItem(IDOK)) {
    ok_button.EnableWindow(allow_ok ? TRUE : FALSE);
  }
}

LRESULT SwitcherSettingsDialog::OnInitDialog(UINT,
                                             WPARAM,
                                             LPARAM,
                                             BOOL&) {
  schema_list_.SubclassWindow(GetDlgItem(IDC_SCHEMA_LIST));
  schema_list_.SetExtendedListViewStyle(LVS_EX_FULLROWSELECT,
                                        LVS_EX_FULLROWSELECT);

  CString schema_name;
  schema_name.LoadStringW(IDS_STR_SCHEMA_NAME);
  schema_list_.AddColumn(schema_name, 0);
  CRect rect;
  schema_list_.GetClientRect(&rect);
  schema_list_.SetColumnWidth(0, rect.Width() - 20);

  description_.Attach(GetDlgItem(IDC_SCHEMA_DESCRIPTION));
  ai_hotkey_hint_.Attach(GetDlgItem(IDC_AI_HOTKEY_HINT));

  ai_hotkey_.SubclassWindow(GetDlgItem(IDC_AI_HOTKEY));
  ai_hotkey_.SetReadOnly(TRUE);
  ai_hotkey_.EnableWindow(TRUE);
  instruction_lookup_prefix_.Attach(
      GetDlgItem(IDC_AI_INSTRUCTION_LOOKUP_PREFIX));

  if (HWND legacy_button = GetDlgItem(IDC_GET_SCHEMATA)) {
    ::EnableWindow(legacy_button, FALSE);
    ::ShowWindow(legacy_button, SW_HIDE);
  }

  Populate();

  CenterWindow();
  BringWindowToTop();
  return TRUE;
}

LRESULT SwitcherSettingsDialog::OnClose(UINT, WPARAM, LPARAM, BOOL&) {
  EndDialog(IDCANCEL);
  return 0;
}

LRESULT SwitcherSettingsDialog::OnAIAssistantHotkeyStateChanged(UINT,
                                                                WPARAM,
                                                                LPARAM,
                                                                BOOL&) {
  UpdateAIAssistantHotkeyUi();
  return 0;
}

LRESULT SwitcherSettingsDialog::OnInstructionLookupPrefixChanged(WORD,
                                                                 WORD,
                                                                 HWND,
                                                                 BOOL&) {
  UpdateAIAssistantHotkeyUi();
  return 0;
}

LRESULT SwitcherSettingsDialog::OnOK(WORD, WORD code, HWND, BOOL&) {
  if (schema_modified_ && settings_ && schema_list_.GetItemCount() != 0) {
    const char** selection = new const char*[schema_list_.GetItemCount()];
    int count = 0;
    for (int i = 0; i < schema_list_.GetItemCount(); ++i) {
      if (!schema_list_.GetCheckState(i)) {
        continue;
      }
      RimeSchemaInfo* info =
          (RimeSchemaInfo*)(schema_list_.GetItemData(i));
      if (info) {
        selection[count++] = api_->get_schema_id(info);
      }
    }
    if (count == 0) {
      MSG_BY_IDS(IDS_STR_ERR_AT_LEAST_ONE_SEL, IDS_STR_NOT_REGULAR,
                 MB_OK | MB_ICONEXCLAMATION);
      delete[] selection;
      return 0;
    }
    api_->select_schemas(settings_, selection, count);
    delete[] selection;
  }

  if ((ai_hotkey_modified_ || instruction_lookup_prefix_modified_) &&
      ai_settings_) {
    if (!IsCurrentAIAssistantHotkeySavable()) {
      ai_hotkey_.SetFocus();
      return 0;
    }
    if (!IsCurrentInstructionLookupPrefixSavable()) {
      instruction_lookup_prefix_.SetFocus();
      return 0;
    }

    std::string ai_hotkey_utf8 = ai_hotkey_.IsEmpty()
                                     ? std::string("Control+3")
                                     : ai_hotkey_.GetCanonicalConfigText();
    if (ai_hotkey_utf8.empty()) {
      ai_hotkey_utf8 = "Control+3";
    }
    std::wstring instruction_lookup_prefix =
        GetCurrentInstructionLookupPrefix();
    if (instruction_lookup_prefix.empty()) {
      instruction_lookup_prefix = kDefaultInstructionLookupPrefix;
    }
    const std::string instruction_lookup_prefix_utf8 =
        wtou8(instruction_lookup_prefix);
    if (!api_->customize_string(ai_settings_, "ai_assistant/trigger_hotkey",
                                ai_hotkey_utf8.c_str()) ||
        !api_->customize_string(ai_settings_,
                                "ai_assistant/instruction_lookup_prefix",
                                instruction_lookup_prefix_utf8.c_str()) ||
        !api_->save_settings(ai_settings_)) {
      MessageBoxW(L"AI 配置保存失败，请重试。", L"智能输入法",
                  MB_OK | MB_ICONEXCLAMATION);
      return 0;
    }
    ai_settings_saved_ = true;

    if (instruction_lookup_prefix_modified_ && instruction_lookup_schema_settings_) {
      const std::string recognizer_pattern =
          BuildInstructionLookupRecognizerPattern(instruction_lookup_prefix);
      if (!api_->customize_string(instruction_lookup_schema_settings_,
                                  "instruction_lookup/prefix",
                                  instruction_lookup_prefix_utf8.c_str()) ||
          !api_->customize_string(
              instruction_lookup_schema_settings_,
              "recognizer/patterns/instruction_lookup",
              recognizer_pattern.c_str()) ||
          !api_->save_settings(instruction_lookup_schema_settings_)) {
        MessageBoxW(
            L"AI 指令检索前缀已保存，但 rime_ice 方案同步失败；重新部署后新前缀仍可用，旧的 sS 规则可能仍保留。",
            L"智能输入法", MB_OK | MB_ICONEXCLAMATION);
      }
    }
  }

  EndDialog(code);
  return 0;
}

LRESULT SwitcherSettingsDialog::OnSchemaListItemChanged(int,
                                                        LPNMHDR notification,
                                                        BOOL&) {
  LPNMLISTVIEW list_view = reinterpret_cast<LPNMLISTVIEW>(notification);
  if (!loaded_ || !list_view || list_view->iItem < 0 ||
      list_view->iItem >= schema_list_.GetItemCount()) {
    return 0;
  }

  if ((list_view->uNewState & LVIS_STATEIMAGEMASK) !=
      (list_view->uOldState & LVIS_STATEIMAGEMASK)) {
    schema_modified_ = true;
  } else if ((list_view->uNewState & LVIS_SELECTED) &&
             !(list_view->uOldState & LVIS_SELECTED)) {
    ShowDetails((RimeSchemaInfo*)(schema_list_.GetItemData(list_view->iItem)));
  }
  return 0;
}
