#include "stdafx.h"

#include <AIAssistantHotkey.h>
#include <boost/detail/lightweight_test.hpp>

#include <cstring>
#include <string>

namespace {

BYTE* ClearKeyState(BYTE (&key_state)[256]) {
  memset(key_state, 0, sizeof(key_state));
  return key_state;
}

void MarkKeyDown(BYTE (&key_state)[256], UINT vkey) {
  key_state[vkey] = 0x80;
}

void ExpectRoundTrip(const std::string& raw_hotkey,
                     const std::string& expected_config,
                     const std::wstring& expected_display) {
  AiHotkeyBinding binding;
  BOOST_TEST(TryParseAiHotkeyConfig(raw_hotkey, &binding));
  BOOST_TEST(binding.valid);
  BOOST_TEST(ValidateAiHotkeyBinding(binding).ok);
  BOOST_TEST(FormatAiHotkeyForConfig(binding) == expected_config);
  BOOST_TEST(FormatAiHotkeyForDisplay(binding) == expected_display);
}

void test_round_trip() {
  ExpectRoundTrip("Control+3", "Control+3", L"Ctrl+3");
  ExpectRoundTrip("Control+Shift+P", "Control+Shift+P", L"Ctrl+Shift+P");
  ExpectRoundTrip("Alt+/", "Alt+/", L"Alt+/");
  ExpectRoundTrip("Control+Backspace", "Control+Backspace",
                  L"Ctrl+Backspace");
  ExpectRoundTrip("Control+Left", "Control+Left", L"Ctrl+Left");
  ExpectRoundTrip("Control+Tab", "Control+Tab", L"Ctrl+Tab");
  ExpectRoundTrip("Control+F24", "Control+F24", L"Ctrl+F24");
}

void test_aliases_and_canonical_order() {
  AiHotkeyBinding binding;
  BOOST_TEST(TryParseAiHotkeyConfig("ctrl+shift+p", &binding));
  BOOST_TEST(FormatAiHotkeyForConfig(binding) == "Control+Shift+P");

  BOOST_TEST(TryParseAiHotkeyConfig("Shift+Control+slash", &binding));
  BOOST_TEST(FormatAiHotkeyForConfig(binding) == "Control+Shift+/");

  BOOST_TEST(TryParseAiHotkeyConfig("Control+escape", &binding));
  BOOST_TEST(FormatAiHotkeyForDisplay(binding) == L"Ctrl+Esc");
}

void test_reject_bare_keys() {
  AiHotkeyBinding binding;

  BOOST_TEST(TryParseAiHotkeyConfig("A", &binding));
  BOOST_TEST(!ValidateAiHotkeyBinding(binding).ok);

  BOOST_TEST(TryParseAiHotkeyConfig("F2", &binding));
  BOOST_TEST(!ValidateAiHotkeyBinding(binding).ok);

  BOOST_TEST(TryParseAiHotkeyConfig("Escape", &binding));
  BOOST_TEST(!ValidateAiHotkeyBinding(binding).ok);
}

void test_reject_modifier_only_and_reserved() {
  AiHotkeyBinding binding;
  BOOST_TEST(!TryParseAiHotkeyConfig("Ctrl", &binding));
  BOOST_TEST(!TryParseAiHotkeyConfig("Ctrl+Shift", &binding));
  BOOST_TEST(!TryParseAiHotkeyConfig("Control+F25", &binding));

  bool threw = false;
  try {
    BOOST_TEST(!TryParseAiHotkeyConfig("Control+F99999999999999999999",
                                       &binding));
  } catch (...) {
    threw = true;
  }
  BOOST_TEST(!threw);

  BOOST_TEST(TryParseAiHotkeyConfig("Control+Space", &binding));
  const AiHotkeyValidationResult validation =
      ValidateAiHotkeyBinding(binding);
  BOOST_TEST(!validation.ok);
  BOOST_TEST(validation.reserved_conflict);
}

void test_build_from_windows_keys() {
  BYTE key_state[256];
  AiHotkeyBinding binding;

  ClearKeyState(key_state);
  MarkKeyDown(key_state, VK_CONTROL);
  MarkKeyDown(key_state, VK_SHIFT);
  BOOST_TEST(TryBuildAiHotkeyFromWinKey('P', false, key_state, &binding));
  BOOST_TEST(binding.valid);
  BOOST_TEST(FormatAiHotkeyForConfig(binding) == "Control+Shift+P");

  ClearKeyState(key_state);
  MarkKeyDown(key_state, VK_CONTROL);
  BOOST_TEST(TryBuildAiHotkeyFromWinKey(VK_LEFT, false, key_state, &binding));
  BOOST_TEST(FormatAiHotkeyForConfig(binding) == "Control+Left");

  ClearKeyState(key_state);
  BOOST_TEST(TryBuildAiHotkeyFromWinKey(VK_BACK, false, key_state, &binding));
  BOOST_TEST(FormatAiHotkeyForConfig(binding) == "Backspace");
  BOOST_TEST(!ValidateAiHotkeyBinding(binding).ok);
}

void test_runtime_match() {
  AiHotkeyBinding binding;
  BOOST_TEST(TryParseAiHotkeyConfig("Control+Shift+P", &binding));
  BOOST_TEST(IsAiHotkeyMatch(
      binding, weasel::KeyEvent('p', ibus::CONTROL_MASK | ibus::SHIFT_MASK)));
  BOOST_TEST(!IsAiHotkeyMatch(
      binding, weasel::KeyEvent('p', ibus::CONTROL_MASK)));

  BOOST_TEST(TryParseAiHotkeyConfig("Control+Shift+1", &binding));
  BOOST_TEST(IsAiHotkeyMatch(
      binding, weasel::KeyEvent('!', ibus::CONTROL_MASK | ibus::SHIFT_MASK)));

  BOOST_TEST(TryParseAiHotkeyConfig("Control+Shift+/", &binding));
  BOOST_TEST(IsAiHotkeyMatch(
      binding, weasel::KeyEvent('?', ibus::CONTROL_MASK | ibus::SHIFT_MASK)));
}

}  // namespace

int _tmain(int, _TCHAR*[]) {
  test_round_trip();
  test_aliases_and_canonical_order();
  test_reject_bare_keys();
  test_reject_modifier_only_and_reserved();
  test_build_from_windows_keys();
  test_runtime_match();
  return boost::report_errors();
}
