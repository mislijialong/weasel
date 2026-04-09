#pragma once

#include <windows.h>

#include <string>

#include <KeyEvent.h>

struct AiHotkeyBinding {
  AiHotkeyBinding() : keycode(0), modifiers(0), valid(false) {}
  AiHotkeyBinding(UINT keycode_value, UINT modifier_mask, bool is_valid = true)
      : keycode(keycode_value), modifiers(modifier_mask), valid(is_valid) {}

  UINT keycode;
  UINT modifiers;
  bool valid;
};

struct AiHotkeyValidationResult {
  AiHotkeyValidationResult() : ok(false), reserved_conflict(false), message() {}

  bool ok;
  bool reserved_conflict;
  std::wstring message;
};

bool TryParseAiHotkeyConfig(const std::string& raw_hotkey,
                            AiHotkeyBinding* binding);
std::string FormatAiHotkeyForConfig(const AiHotkeyBinding& binding);
std::wstring FormatAiHotkeyForDisplay(const AiHotkeyBinding& binding);
std::wstring FormatAiHotkeyModifiersForDisplay(UINT modifiers,
                                               bool trailing_plus = true);
bool TryBuildAiHotkeyFromWinKey(UINT vkey,
                                bool extended,
                                BYTE key_state[256],
                                AiHotkeyBinding* binding);
AiHotkeyValidationResult ValidateAiHotkeyBinding(
    const AiHotkeyBinding& binding);
bool IsAiHotkeyMatch(const AiHotkeyBinding& binding,
                     const weasel::KeyEvent& key_event);
bool IsAiHotkeyModifierVirtualKey(UINT vkey);
