#include <AIAssistantHotkey.h>

#include <algorithm>
#include <cctype>

namespace {

constexpr UINT kSupportedHotkeyModifiers = ibus::CONTROL_MASK |
                                           ibus::SHIFT_MASK |
                                           ibus::ALT_MASK |
                                           ibus::SUPER_MASK |
                                           ibus::META_MASK |
                                           ibus::HYPER_MASK;

std::string TrimAsciiWhitespace(const std::string& input) {
  size_t first = 0;
  while (first < input.size() &&
         std::isspace(static_cast<unsigned char>(input[first]))) {
    ++first;
  }
  size_t last = input.size();
  while (last > first &&
         std::isspace(static_cast<unsigned char>(input[last - 1]))) {
    --last;
  }
  return input.substr(first, last - first);
}

std::string ToLowerAscii(std::string input) {
  std::transform(input.begin(), input.end(), input.begin(),
                 [](unsigned char ch) {
                   return static_cast<char>(std::tolower(ch));
                 });
  return input;
}

UINT NormalizeAiHotkeyKeycode(UINT keycode) {
  if (keycode >= 'A' && keycode <= 'Z') {
    return keycode - 'A' + 'a';
  }
  if (keycode == ibus::grave) {
    return '`';
  }
  switch (keycode) {
    case '!':
      return '1';
    case '@':
      return '2';
    case '#':
      return '3';
    case '$':
      return '4';
    case '%':
      return '5';
    case '^':
      return '6';
    case '&':
      return '7';
    case '*':
      return '8';
    case '(':
      return '9';
    case ')':
      return '0';
    case '_':
      return '-';
    case '+':
      return '=';
    case ':':
      return ';';
    case '"':
      return '\'';
    case '<':
      return ',';
    case '>':
      return '.';
    case '?':
      return '/';
    case '{':
      return '[';
    case '}':
      return ']';
    case '|':
      return '\\';
    case '~':
      return static_cast<UINT>(ibus::grave);
    default:
      break;
  }
  return keycode;
}

bool TryParseNamedHotkeyKey(const std::string& token, UINT* keycode) {
  if (!keycode) {
    return false;
  }
  if (token.empty()) {
    return false;
  }
  if (token.size() == 1) {
    const unsigned char ch = static_cast<unsigned char>(token[0]);
    if (std::isalnum(ch) || ch == '/' || ch == '\\' || ch == ',' || ch == '.' ||
        ch == ';' || ch == '\'' || ch == '-' || ch == '=' || ch == '[' ||
        ch == ']' || ch == '`') {
      *keycode = (ch == '`') ? static_cast<UINT>(ibus::grave)
                             : static_cast<UINT>(ch);
      return true;
    }
  }

  if (token == "space" || token == "spacebar") {
    *keycode = ibus::space;
    return true;
  }
  if (token == "tab") {
    *keycode = ibus::Tab;
    return true;
  }
  if (token == "enter" || token == "return") {
    *keycode = ibus::Return;
    return true;
  }
  if (token == "esc" || token == "escape") {
    *keycode = ibus::Escape;
    return true;
  }
  if (token == "backspace") {
    *keycode = ibus::BackSpace;
    return true;
  }
  if (token == "delete" || token == "del") {
    *keycode = ibus::Delete;
    return true;
  }
  if (token == "insert" || token == "ins") {
    *keycode = ibus::Insert;
    return true;
  }
  if (token == "home") {
    *keycode = ibus::Home;
    return true;
  }
  if (token == "end") {
    *keycode = ibus::End;
    return true;
  }
  if (token == "pageup" || token == "page_up" || token == "pgup" ||
      token == "prior") {
    *keycode = ibus::Page_Up;
    return true;
  }
  if (token == "pagedown" || token == "page_down" || token == "pgdn" ||
      token == "next") {
    *keycode = ibus::Page_Down;
    return true;
  }
  if (token == "left") {
    *keycode = ibus::Left;
    return true;
  }
  if (token == "right") {
    *keycode = ibus::Right;
    return true;
  }
  if (token == "up") {
    *keycode = ibus::Up;
    return true;
  }
  if (token == "down") {
    *keycode = ibus::Down;
    return true;
  }
  if (token == "grave" || token == "backquote" || token == "backtick") {
    *keycode = ibus::grave;
    return true;
  }
  if (token == "slash") {
    *keycode = '/';
    return true;
  }
  if (token == "backslash") {
    *keycode = '\\';
    return true;
  }
  if (token == "comma") {
    *keycode = ',';
    return true;
  }
  if (token == "period" || token == "dot") {
    *keycode = '.';
    return true;
  }
  if (token == "semicolon") {
    *keycode = ';';
    return true;
  }
  if (token == "quote" || token == "apostrophe") {
    *keycode = '\'';
    return true;
  }
  if (token == "minus" || token == "hyphen") {
    *keycode = '-';
    return true;
  }
  if (token == "equal" || token == "equals") {
    *keycode = '=';
    return true;
  }
  if (token == "leftbracket" || token == "lbracket" ||
      token == "openbracket") {
    *keycode = '[';
    return true;
  }
  if (token == "rightbracket" || token == "rbracket" ||
      token == "closebracket") {
    *keycode = ']';
    return true;
  }
  if (token.size() > 1 && token[0] == 'f') {
    bool all_digits = true;
    int function_key = 0;
    for (size_t i = 1; i < token.size(); ++i) {
      const unsigned char ch = static_cast<unsigned char>(token[i]);
      if (!std::isdigit(ch)) {
        all_digits = false;
        break;
      }
      function_key = function_key * 10 + (ch - '0');
      if (function_key > 24) {
        return false;
      }
    }
    if (all_digits && function_key >= 1) {
      *keycode = static_cast<UINT>(ibus::F1 + function_key - 1);
      return true;
    }
  }

  return false;
}

bool TryGetHotkeyKeyLabels(UINT keycode,
                           std::string* config_label,
                           std::wstring* display_label) {
  if (!config_label || !display_label) {
    return false;
  }

  const UINT normalized_keycode = NormalizeAiHotkeyKeycode(keycode);
  if (normalized_keycode >= 'a' && normalized_keycode <= 'z') {
    const char upper = static_cast<char>(normalized_keycode - 'a' + 'A');
    *config_label = std::string(1, upper);
    *display_label = std::wstring(1, static_cast<wchar_t>(upper));
    return true;
  }
  if (normalized_keycode >= '0' && normalized_keycode <= '9') {
    const char digit = static_cast<char>(normalized_keycode);
    *config_label = std::string(1, digit);
    *display_label = std::wstring(1, static_cast<wchar_t>(digit));
    return true;
  }
  switch (normalized_keycode) {
    case '/':
    case '\\':
    case ',':
    case '.':
    case ';':
    case '\'':
    case '-':
    case '=':
    case '[':
    case ']': {
      const char ch = static_cast<char>(normalized_keycode);
      *config_label = std::string(1, ch);
      *display_label = std::wstring(1, static_cast<wchar_t>(ch));
      return true;
    }
    case '`':
      *config_label = "`";
      *display_label = L"`";
      return true;
    case ibus::space:
      *config_label = "Space";
      *display_label = L"Space";
      return true;
    case ibus::Tab:
      *config_label = "Tab";
      *display_label = L"Tab";
      return true;
    case ibus::Return:
      *config_label = "Enter";
      *display_label = L"Enter";
      return true;
    case ibus::Escape:
      *config_label = "Escape";
      *display_label = L"Esc";
      return true;
    case ibus::BackSpace:
      *config_label = "Backspace";
      *display_label = L"Backspace";
      return true;
    case ibus::Delete:
      *config_label = "Delete";
      *display_label = L"Delete";
      return true;
    case ibus::Insert:
      *config_label = "Insert";
      *display_label = L"Insert";
      return true;
    case ibus::Home:
      *config_label = "Home";
      *display_label = L"Home";
      return true;
    case ibus::End:
      *config_label = "End";
      *display_label = L"End";
      return true;
    case ibus::Page_Up:
      *config_label = "PageUp";
      *display_label = L"PageUp";
      return true;
    case ibus::Page_Down:
      *config_label = "PageDown";
      *display_label = L"PageDown";
      return true;
    case ibus::Left:
      *config_label = "Left";
      *display_label = L"Left";
      return true;
    case ibus::Right:
      *config_label = "Right";
      *display_label = L"Right";
      return true;
    case ibus::Up:
      *config_label = "Up";
      *display_label = L"Up";
      return true;
    case ibus::Down:
      *config_label = "Down";
      *display_label = L"Down";
      return true;
    default:
      break;
  }

  if (normalized_keycode >= ibus::F1 && normalized_keycode <= ibus::F24) {
    const int function_key = normalized_keycode - ibus::F1 + 1;
    *config_label = "F" + std::to_string(function_key);
    *display_label =
        L"F" + std::to_wstring(static_cast<long long>(function_key));
    return true;
  }

  return false;
}

void AppendModifierConfigLabels(UINT modifiers, std::string* text) {
  if (!text) {
    return;
  }
  const auto append = [text](const char* token) {
    if (!text->empty()) {
      *text += "+";
    }
    *text += token;
  };
  if ((modifiers & ibus::CONTROL_MASK) != 0) {
    append("Control");
  }
  if ((modifiers & ibus::ALT_MASK) != 0) {
    append("Alt");
  }
  if ((modifiers & ibus::SHIFT_MASK) != 0) {
    append("Shift");
  }
  if ((modifiers & ibus::SUPER_MASK) != 0) {
    append("Win");
  }
  if ((modifiers & ibus::META_MASK) != 0) {
    append("Meta");
  }
  if ((modifiers & ibus::HYPER_MASK) != 0) {
    append("Hyper");
  }
}

void AppendModifierDisplayLabels(UINT modifiers,
                                 std::wstring* text,
                                 bool trailing_plus) {
  if (!text) {
    return;
  }
  const auto append = [text](const wchar_t* token) {
    if (!text->empty()) {
      *text += L"+";
    }
    *text += token;
  };
  if ((modifiers & ibus::CONTROL_MASK) != 0) {
    append(L"Ctrl");
  }
  if ((modifiers & ibus::ALT_MASK) != 0) {
    append(L"Alt");
  }
  if ((modifiers & ibus::SHIFT_MASK) != 0) {
    append(L"Shift");
  }
  if ((modifiers & ibus::SUPER_MASK) != 0) {
    append(L"Win");
  }
  if ((modifiers & ibus::META_MASK) != 0) {
    append(L"Meta");
  }
  if ((modifiers & ibus::HYPER_MASK) != 0) {
    append(L"Hyper");
  }
  if (trailing_plus && !text->empty()) {
    *text += L"+";
  }
}

UINT BuildHotkeyModifierMask(const BYTE key_state[256]) {
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

bool TryMapVirtualKeyToHotkeyKeycode(UINT vkey, bool /*extended*/, UINT* keycode) {
  if (!keycode) {
    return false;
  }

  if ((vkey >= 'A' && vkey <= 'Z') || (vkey >= '0' && vkey <= '9')) {
    *keycode = vkey;
    return true;
  }
  if (vkey >= VK_F1 && vkey <= VK_F24) {
    *keycode = static_cast<UINT>(ibus::F1 + (vkey - VK_F1));
    return true;
  }

  switch (vkey) {
    case VK_SPACE:
      *keycode = ibus::space;
      return true;
    case VK_TAB:
      *keycode = ibus::Tab;
      return true;
    case VK_RETURN:
      *keycode = ibus::Return;
      return true;
    case VK_ESCAPE:
      *keycode = ibus::Escape;
      return true;
    case VK_BACK:
      *keycode = ibus::BackSpace;
      return true;
    case VK_DELETE:
      *keycode = ibus::Delete;
      return true;
    case VK_INSERT:
      *keycode = ibus::Insert;
      return true;
    case VK_HOME:
      *keycode = ibus::Home;
      return true;
    case VK_END:
      *keycode = ibus::End;
      return true;
    case VK_PRIOR:
      *keycode = ibus::Page_Up;
      return true;
    case VK_NEXT:
      *keycode = ibus::Page_Down;
      return true;
    case VK_LEFT:
      *keycode = ibus::Left;
      return true;
    case VK_RIGHT:
      *keycode = ibus::Right;
      return true;
    case VK_UP:
      *keycode = ibus::Up;
      return true;
    case VK_DOWN:
      *keycode = ibus::Down;
      return true;
    case VK_OEM_1:
      *keycode = ';';
      return true;
    case VK_OEM_PLUS:
      *keycode = '=';
      return true;
    case VK_OEM_COMMA:
      *keycode = ',';
      return true;
    case VK_OEM_MINUS:
      *keycode = '-';
      return true;
    case VK_OEM_PERIOD:
      *keycode = '.';
      return true;
    case VK_OEM_2:
      *keycode = '/';
      return true;
    case VK_OEM_3:
      *keycode = ibus::grave;
      return true;
    case VK_OEM_4:
      *keycode = '[';
      return true;
    case VK_OEM_5:
      *keycode = '\\';
      return true;
    case VK_OEM_6:
      *keycode = ']';
      return true;
    case VK_OEM_7:
      *keycode = '\'';
      return true;
    default:
      break;
  }

  return false;
}

bool IsModifierKeycode(UINT keycode) {
  switch (keycode) {
    case ibus::Shift_L:
    case ibus::Shift_R:
    case ibus::Control_L:
    case ibus::Control_R:
    case ibus::Alt_L:
    case ibus::Alt_R:
    case ibus::Meta_L:
    case ibus::Meta_R:
    case ibus::Super_L:
    case ibus::Super_R:
    case ibus::Hyper_L:
    case ibus::Hyper_R:
      return true;
    default:
      return false;
  }
}

}  // namespace

bool TryParseAiHotkeyConfig(const std::string& raw_hotkey,
                            AiHotkeyBinding* binding) {
  if (!binding) {
    return false;
  }

  *binding = AiHotkeyBinding();
  const std::string hotkey = TrimAsciiWhitespace(raw_hotkey);
  if (hotkey.empty()) {
    return false;
  }

  bool has_key = false;
  UINT parsed_keycode = 0;
  UINT parsed_modifiers = 0;
  size_t begin = 0;
  while (begin <= hotkey.size()) {
    const size_t end = hotkey.find('+', begin);
    const std::string token_raw =
        end == std::string::npos ? hotkey.substr(begin)
                                 : hotkey.substr(begin, end - begin);
    const std::string token = ToLowerAscii(TrimAsciiWhitespace(token_raw));
    if (token.empty()) {
      return false;
    }

    UINT modifier = 0;
    if (token == "ctrl" || token == "control") {
      modifier = ibus::CONTROL_MASK;
    } else if (token == "alt") {
      modifier = ibus::ALT_MASK;
    } else if (token == "shift") {
      modifier = ibus::SHIFT_MASK;
    } else if (token == "win" || token == "windows" || token == "super") {
      modifier = ibus::SUPER_MASK;
    } else if (token == "meta") {
      modifier = ibus::META_MASK;
    } else if (token == "hyper") {
      modifier = ibus::HYPER_MASK;
    }

    if (modifier != 0) {
      parsed_modifiers |= modifier;
    } else {
      if (has_key || !TryParseNamedHotkeyKey(token, &parsed_keycode)) {
        return false;
      }
      has_key = true;
    }

    if (end == std::string::npos) {
      break;
    }
    begin = end + 1;
  }

  if (!has_key) {
    return false;
  }

  *binding = AiHotkeyBinding(NormalizeAiHotkeyKeycode(parsed_keycode),
                             parsed_modifiers & kSupportedHotkeyModifiers,
                             true);
  return true;
}

std::string FormatAiHotkeyForConfig(const AiHotkeyBinding& binding) {
  if (!binding.valid) {
    return std::string();
  }

  std::string text;
  AppendModifierConfigLabels(binding.modifiers & kSupportedHotkeyModifiers,
                             &text);

  std::string key_label;
  std::wstring unused_display;
  if (!TryGetHotkeyKeyLabels(binding.keycode, &key_label, &unused_display)) {
    return std::string();
  }
  if (!text.empty()) {
    text += "+";
  }
  text += key_label;
  return text;
}

std::wstring FormatAiHotkeyForDisplay(const AiHotkeyBinding& binding) {
  if (!binding.valid) {
    return std::wstring();
  }

  std::wstring text;
  AppendModifierDisplayLabels(binding.modifiers & kSupportedHotkeyModifiers,
                              &text, false);

  std::string unused_config;
  std::wstring key_label;
  if (!TryGetHotkeyKeyLabels(binding.keycode, &unused_config, &key_label)) {
    return std::wstring();
  }
  if (!text.empty()) {
    text += L"+";
  }
  text += key_label;
  return text;
}

std::wstring FormatAiHotkeyModifiersForDisplay(UINT modifiers,
                                               bool trailing_plus) {
  std::wstring text;
  AppendModifierDisplayLabels(modifiers & kSupportedHotkeyModifiers, &text,
                              trailing_plus);
  return text;
}

bool TryBuildAiHotkeyFromWinKey(UINT vkey,
                                bool extended,
                                BYTE key_state[256],
                                AiHotkeyBinding* binding) {
  if (!binding) {
    return false;
  }

  *binding = AiHotkeyBinding();
  if (IsAiHotkeyModifierVirtualKey(vkey)) {
    return false;
  }

  UINT keycode = 0;
  if (!TryMapVirtualKeyToHotkeyKeycode(vkey, extended, &keycode)) {
    return false;
  }

  const UINT modifiers =
      key_state ? BuildHotkeyModifierMask(key_state) : 0U;
  *binding = AiHotkeyBinding(NormalizeAiHotkeyKeycode(keycode),
                             modifiers & kSupportedHotkeyModifiers, true);
  return true;
}

AiHotkeyValidationResult ValidateAiHotkeyBinding(
    const AiHotkeyBinding& binding) {
  AiHotkeyValidationResult result;
  if (!binding.valid) {
    result.message = L"Hotkey format is invalid.";
    return result;
  }

  if (binding.keycode == 0 || IsModifierKeycode(binding.keycode)) {
    result.message = L"Press a non-modifier key to complete the hotkey.";
    return result;
  }

  const UINT modifiers = binding.modifiers & kSupportedHotkeyModifiers;
  if (modifiers == 0) {
    result.message =
        L"The AI panel hotkey must include Ctrl, Alt, Shift, or Win.";
    return result;
  }

  if (NormalizeAiHotkeyKeycode(binding.keycode) == ibus::space &&
      modifiers == ibus::CONTROL_MASK) {
    result.reserved_conflict = true;
    result.message = L"Ctrl+Space is reserved for the fixed IME toggle.";
    return result;
  }

  result.ok = true;
  return result;
}

bool IsAiHotkeyMatch(const AiHotkeyBinding& binding,
                     const weasel::KeyEvent& key_event) {
  if (!binding.valid) {
    return false;
  }

  if (NormalizeAiHotkeyKeycode(key_event.keycode) !=
      NormalizeAiHotkeyKeycode(binding.keycode)) {
    return false;
  }

  const UINT modifiers = key_event.mask & kSupportedHotkeyModifiers;
  return modifiers == (binding.modifiers & kSupportedHotkeyModifiers);
}

bool IsAiHotkeyModifierVirtualKey(UINT vkey) {
  switch (vkey) {
    case VK_SHIFT:
    case VK_LSHIFT:
    case VK_RSHIFT:
    case VK_CONTROL:
    case VK_LCONTROL:
    case VK_RCONTROL:
    case VK_MENU:
    case VK_LMENU:
    case VK_RMENU:
    case VK_LWIN:
    case VK_RWIN:
      return true;
    default:
      return false;
  }
}
