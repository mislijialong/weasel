#pragma once
#include <AIAssistantHotkey.h>
#include <InputContentStore.h>
#include <WeaselIPC.h>
#include <WeaselUI.h>
#include <atomic>
#include <filesystem>
#include <map>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

#include <rime_api.h>

struct CaseInsensitiveCompare {
  bool operator()(const std::string& str1, const std::string& str2) const {
    std::string str1Lower, str2Lower;
    std::transform(str1.begin(), str1.end(), std::back_inserter(str1Lower),
                   [](char c) { return std::tolower(c); });
    std::transform(str2.begin(), str2.end(), std::back_inserter(str2Lower),
                   [](char c) { return std::tolower(c); });
    return str1Lower < str2Lower;
  }
};

typedef std::map<std::string, bool> AppOptions;
typedef std::map<std::string, AppOptions, CaseInsensitiveCompare>
    AppOptionsByAppName;

struct SessionStatus {
  SessionStatus()
      : style(weasel::UIStyle()),
        __synced(false),
        session_id(0),
        client_app() {
    RIME_STRUCT(RimeStatus, status);
  }
  weasel::UIStyle style;
  RimeStatus status;
  bool __synced;
  RimeSessionId session_id;
  std::string client_app;
};
typedef std::map<DWORD, SessionStatus> SessionStatusMap;
typedef DWORD WeaselSessionId;

struct AIAssistantConfig {
  AIAssistantConfig()
      : enabled(false),
        stream(true),
        login_required(false),
        debug_dump_context(false),
        trigger_hotkey("Control+3"),
        trigger_binding('3', ibus::CONTROL_MASK),
        endpoint(),
        api_key(),
        model("gpt-5"),
        debug_dump_path("ai_context_dump.txt"),
        panel_url(),
        panel_allowed_origin(),
        system_prompt(),
        reasoning_effort("low"),
        login_url(),
        login_state_path("ai_login_state.json"),
        login_token_key("token"),
        refresh_token_endpoint(),
        mqtt_url(),
        mqtt_topic_template("/mqtt/topic/sino/lamp/oauth/token/login/{uuid}"),
        mqtt_username(),
        mqtt_password(),
        mqtt_timeout_ms(120000),
        max_history_chars(2048),
        timeout_ms(30000) {}
  bool enabled;
  bool stream;
  bool login_required;
  bool debug_dump_context;
  std::string trigger_hotkey;
  AiHotkeyBinding trigger_binding;
  std::string endpoint;
  std::string api_key;
  std::string model;
  std::string debug_dump_path;
  std::string panel_url;
  std::string panel_allowed_origin;
  std::string system_prompt;
  std::string reasoning_effort;
  std::string login_url;
  std::string login_state_path;
  std::string login_token_key;
  std::string refresh_token_endpoint;
  std::string mqtt_url;
  std::string mqtt_topic_template;
  std::string mqtt_username;
  std::string mqtt_password;
  int mqtt_timeout_ms;
  int max_history_chars;
  int timeout_ms;
};

struct AIPanelInstitutionOption {
  AIPanelInstitutionOption() : id(), name(), app_key(), template_content() {}
  AIPanelInstitutionOption(const std::wstring& option_id,
                           const std::wstring& option_name,
                           const std::wstring& option_app_key,
                           const std::wstring& option_template_content =
                               std::wstring())
      : id(option_id),
        name(option_name),
        app_key(option_app_key),
        template_content(option_template_content) {}
  std::wstring id;
  std::wstring name;
  std::wstring app_key;
  std::wstring template_content;
};

struct AIPanelRuntime {
  AIPanelRuntime()
      : panel_hwnd(nullptr),
        status_hwnd(nullptr),
        output_hwnd(nullptr),
        webview_hwnd(nullptr),
        webview_controller(nullptr),
        webview(nullptr),
        request_hwnd(nullptr),
        confirm_hwnd(nullptr),
        cancel_hwnd(nullptr),
        target_hwnd(nullptr),
        request_id(0),
        ipc_id(0),
        context_text(),
        status_text(),
        output_text(),
        institution_options(),
        selected_institution_id(),
        panel_width(540),
        panel_height(360),
        last_panel_x(0),
        last_panel_y(0),
        has_last_panel_position(false),
        resize_tracking(false),
        resize_edges(0),
        resize_start_screen_x(0),
        resize_start_screen_y(0),
        resize_start_window_x(0),
        resize_start_window_y(0),
        resize_start_width(0),
        resize_start_height(0),
        focus_pending(false),
        webview_ready(false),
        requesting(false),
        institutions_loading(false),
        completed(false),
        has_error(false) {}
  HWND panel_hwnd;
  HWND status_hwnd;
  HWND output_hwnd;
  HWND webview_hwnd;
  void* webview_controller;
  void* webview;
  HWND request_hwnd;
  HWND confirm_hwnd;
  HWND cancel_hwnd;
  HWND target_hwnd;
  uint64_t request_id;
  WeaselSessionId ipc_id;
  std::wstring context_text;
  std::wstring status_text;
  std::wstring output_text;
  std::vector<AIPanelInstitutionOption> institution_options;
  std::wstring selected_institution_id;
  int panel_width;
  int panel_height;
  int last_panel_x;
  int last_panel_y;
  bool has_last_panel_position;
  bool resize_tracking;
  int resize_edges;
  int resize_start_screen_x;
  int resize_start_screen_y;
  int resize_start_window_x;
  int resize_start_window_y;
  int resize_start_width;
  int resize_start_height;
  bool focus_pending;
  bool webview_ready;
  bool requesting;
  bool institutions_loading;
  bool completed;
  bool has_error;
};

struct SystemCommandLaunchRequest {
  std::string command_id;
  std::filesystem::path preferred_output_dir;
};

class RimeWithWeaselHandler : public weasel::RequestHandler {
 public:
  RimeWithWeaselHandler(weasel::UI* ui);
  virtual ~RimeWithWeaselHandler();
  virtual void Initialize();
  virtual void Finalize();
  virtual DWORD FindSession(WeaselSessionId ipc_id);
  virtual DWORD AddSession(LPWSTR buffer, EatLine eat = 0);
  virtual DWORD RemoveSession(WeaselSessionId ipc_id);
  virtual BOOL ProcessKeyEvent(weasel::KeyEvent keyEvent,
                               WeaselSessionId ipc_id,
                               EatLine eat);
  virtual void CommitComposition(WeaselSessionId ipc_id);
  virtual void ClearComposition(WeaselSessionId ipc_id);
  virtual void SelectCandidateOnCurrentPage(size_t index,
                                            WeaselSessionId ipc_id);
  virtual bool HighlightCandidateOnCurrentPage(size_t index,
                                               WeaselSessionId ipc_id,
                                               EatLine eat);
  virtual bool ChangePage(bool backward, WeaselSessionId ipc_id, EatLine eat);
  virtual void FocusIn(DWORD param, WeaselSessionId ipc_id);
  virtual void FocusOut(DWORD param, WeaselSessionId ipc_id);
  virtual void UpdateInputPosition(RECT const& rc, WeaselSessionId ipc_id);
  virtual void StartMaintenance();
  virtual void EndMaintenance();
  virtual void SetOption(WeaselSessionId ipc_id,
                         const std::string& opt,
                         bool val);
  virtual void UpdateColorTheme(BOOL darkMode);

  void OnUpdateUI(std::function<void()> const& cb);
  void OnSystemCommand(
      std::function<void(const SystemCommandLaunchRequest&)> const& cb);

 private:
  void _Setup();
  bool _IsDeployerRunning();
  void _UpdateUI(WeaselSessionId ipc_id);
  void _LoadSchemaSpecificSettings(WeaselSessionId ipc_id,
                                   const std::string& schema_id);
  void _LoadAppInlinePreeditSet(WeaselSessionId ipc_id,
                                bool ignore_app_name = false);
  bool _ShowMessage(weasel::Context& ctx, weasel::Status& status);
  bool _Respond(WeaselSessionId ipc_id,
                EatLine eat,
                const std::wstring* extra_commit = nullptr);
  void _ReadClientInfo(WeaselSessionId ipc_id, LPWSTR buffer);
  void _GetCandidateInfo(weasel::CandidateInfo& cinfo, RimeContext& ctx);
  void _GetStatus(weasel::Status& stat,
                  WeaselSessionId ipc_id,
                  weasel::Context& ctx);
  void _GetContext(weasel::Context& ctx, RimeSessionId session_id);
  void _UpdateShowNotifications(RimeConfig* config, bool initialize = false);
  void _LoadAIAssistantConfig(RimeConfig* config);
  bool _EnsureAIAssistantLogin();
  bool _ForceAIAssistantRelogin();
  bool _StartAIAssistantLoginFlow();
  bool _IsAIAssistantLoggedIn();
  void _StopAIAssistantLoginFlow();
  void _RunAIAssistantLoginListener(const std::string& client_id);
  bool _TryProcessAIAssistantTrigger(weasel::KeyEvent keyEvent,
                                     WeaselSessionId ipc_id,
                                     EatLine eat);
  bool _PrepareAIPanelInstitutionOptionsForOpen(
      std::vector<AIPanelInstitutionOption>* options,
      bool* options_ready,
      bool* relogin_started,
      std::string* error_message);
  bool _EnsureAIPanelWindow();
  bool _OpenAIPanel(WeaselSessionId ipc_id,
                    HWND target_hwnd,
                    uint64_t request_id,
                    const std::vector<AIPanelInstitutionOption>* initial_options =
                        nullptr,
                    bool institutions_ready = false);
  void _CloseAIPanel();
  void _DestroyAIPanel();
  void _SetAIPanelStatus(const std::wstring& status_text);
  void _ResetAIPanelOutput();
  void _AppendAIPanelOutput(const std::wstring& chunk);
  void _CompleteAIPanel(bool has_error, const std::wstring& error_text);
  void _ResizeAIPanelWebView();
  void _RequestAIPanelGeneration();
  void _ConfirmAIPanelOutput();
  void _CancelAIPanelOutput();
  void _ExecuteAIPanelSystemCommand(const std::wstring& command_id);
  void _RefreshAIPanelInstitutionOptions();
  void _SetAIPanelContextText(const std::wstring& context_text);
  void _SetAIPanelInstitutionSelection(const std::wstring& institution_id);
  void _HandleAIPanelWritebackRequest(const std::wstring& text);
  void _ApplyAIPanelSizeAndReposition(int requested_width,
                                      int requested_height,
                                      bool prefer_anchor_position = false);
  void _StartAIAssistantStreamRequest(uint64_t request_id,
                                      const std::wstring& context_text);
  bool _SendTextToTargetWindow(HWND target_hwnd, const std::wstring& text);
  static LRESULT CALLBACK AIAssistantPanelWndProc(HWND hwnd,
                                                  UINT msg,
                                                  WPARAM wParam,
                                                  LPARAM lParam);
  void _AppendCommittedText(WeaselSessionId ipc_id, const std::wstring& text);
  std::wstring _CollectAIAssistantContext(WeaselSessionId ipc_id,
                                          const std::wstring& current_text);
  std::wstring _TakePendingCommitText(RimeSessionId session_id);
  bool _TryHandleSystemCommandCommit(std::wstring* commit_text,
                                     WeaselSessionId ipc_id);
  std::filesystem::path _ReadSystemCommandOutputDir(WeaselSessionId ipc_id) const;
  std::string _GetInputContentContextKey(WeaselSessionId ipc_id);
  std::string _GetContextCacheKey(WeaselSessionId ipc_id) const;

  void _UpdateInlinePreeditStatus(WeaselSessionId ipc_id);

  RimeSessionId to_session_id(WeaselSessionId ipc_id) {
    return m_session_status_map[ipc_id].session_id;
  }
  SessionStatus& get_session_status(WeaselSessionId ipc_id) {
    return m_session_status_map[ipc_id];
  }
  SessionStatus& new_session_status(WeaselSessionId ipc_id) {
    return m_session_status_map[ipc_id] = SessionStatus();
  }

  AppOptionsByAppName m_app_options;
  weasel::UI* m_ui;  // reference
  DWORD m_active_session;
  bool m_disabled;
  std::string m_last_schema_id;
  std::string m_last_app_name;
  weasel::UIStyle m_base_style;
  std::map<std::string, bool> m_show_notifications;
  std::map<std::string, bool> m_show_notifications_base;
  std::function<void()> _UpdateUICallback;
  std::function<void(const SystemCommandLaunchRequest&)>
      m_system_command_callback;

  static void OnNotify(void* context_object,
                       uintptr_t session_id,
                       const char* message_type,
                       const char* message_value);
  static std::string m_message_type;
  static std::string m_message_value;
  static std::string m_message_label;
  static std::string m_option_name;
  static std::mutex m_notifier_mutex;
  SessionStatusMap m_session_status_map;
  bool m_current_dark_mode;
  bool m_global_ascii_mode;
  int m_show_notifications_time;
  DWORD m_pid;
  AIAssistantConfig m_ai_config;
  InputContentStore m_input_content_store;
  std::string m_input_active_context_key;
  std::thread m_ai_login_thread;
  std::mutex m_ai_login_mutex;
  std::atomic<bool> m_ai_login_pending;
  std::atomic<bool> m_ai_login_stop;
  std::string m_ai_login_token;
  std::string m_ai_login_tenant_id;
  std::string m_ai_login_refresh_token;
  std::string m_ai_login_client_id;
  AIPanelRuntime m_ai_panel;
  std::mutex m_ai_panel_mutex;
  RECT m_last_input_rect;
  bool m_has_last_input_rect;
  std::atomic<uint64_t> m_ai_request_seq;
};
