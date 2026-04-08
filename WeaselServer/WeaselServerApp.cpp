#include "stdafx.h"
#include "WeaselServerApp.h"
#include <filesystem>
#include <system_error>
#include <vector>

namespace {

std::wstring ExpandEnvironmentVariables(const std::wstring& input) {
  if (input.empty()) {
    return input;
  }
  const DWORD required = ExpandEnvironmentStringsW(input.c_str(), nullptr, 0);
  if (required == 0) {
    return input;
  }
  std::wstring expanded(required, L'\0');
  const DWORD written = ExpandEnvironmentStringsW(
      input.c_str(), expanded.data(), static_cast<DWORD>(expanded.size()));
  if (written == 0) {
    return input;
  }
  expanded.resize(written > 0 ? written - 1 : 0);
  return expanded;
}

void AppendUniquePath(std::vector<fs::path>* paths, const fs::path& path) {
  if (!paths || path.empty()) {
    return;
  }
  const fs::path normalized = path.lexically_normal();
  for (const fs::path& existing : *paths) {
    if (existing.lexically_normal() == normalized) {
      return;
    }
  }
  paths->push_back(normalized);
}

std::vector<fs::path> BuildOutputDirCandidates(
    const SystemCommandLaunchRequest& request) {
  std::vector<fs::path> candidates;
  AppendUniquePath(&candidates, request.preferred_output_dir);
  AppendUniquePath(&candidates,
                   fs::path(ExpandEnvironmentVariables(L"%USERPROFILE%\\Desktop")));
  AppendUniquePath(&candidates, WeaselUserDataPath());
  return candidates;
}

fs::path BuildDocumentPath(const fs::path& dir,
                           const wchar_t* extension,
                           int suffix) {
  SYSTEMTIME now = {0};
  GetLocalTime(&now);
  wchar_t stem[64] = {0};
  swprintf_s(stem, L"weasel_note_%04u%02u%02u_%02u%02u%02u", now.wYear,
             now.wMonth, now.wDay, now.wHour, now.wMinute, now.wSecond);

  std::wstring filename(stem);
  if (suffix > 0) {
    filename.append(L"_").append(std::to_wstring(suffix));
  }
  filename.append(L".").append(extension);
  return dir / filename;
}

bool CreateEmptyDocument(const fs::path& path) {
  std::error_code ec;
  const fs::path parent = path.parent_path();
  if (!parent.empty()) {
    fs::create_directories(parent, ec);
    if (ec) {
      return false;
    }
  }

  HANDLE handle = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr,
                              CREATE_NEW, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (handle == INVALID_HANDLE_VALUE) {
    return false;
  }
  CloseHandle(handle);
  return true;
}

}  // namespace

WeaselServerApp::WeaselServerApp()
    : m_handler(std::make_unique<RimeWithWeaselHandler>(&m_ui)),
      tray_icon(m_ui) {
  // m_handler.reset(new RimeWithWeaselHandler(&m_ui));
  m_handler->OnSystemCommand(
      [this](const SystemCommandLaunchRequest& request) {
        EnqueueSystemCommand(request);
      });
  StartLauncherWorker();
  m_server.SetRequestHandler(m_handler.get());
  SetupMenuHandlers();
}

WeaselServerApp::~WeaselServerApp() {
  StopLauncherWorker();
}

int WeaselServerApp::Run() {
  if (!m_server.Start())
    return -1;

  // win_sparkle_set_appcast_url("http://localhost:8000/weasel/update/appcast.xml");
  win_sparkle_set_registry_path("Software\\Rime\\Weasel\\Updates");
  if (GetThreadUILanguage() ==
      MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL))
    win_sparkle_set_lang("zh-TW");
  else if (GetThreadUILanguage() ==
           MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED))
    win_sparkle_set_lang("zh-CN");
  else
    win_sparkle_set_lang("en");
  win_sparkle_init();
  m_ui.Create(m_server.GetHWnd());

  m_handler->Initialize();
  m_handler->OnUpdateUI([this]() { tray_icon.Refresh(); });

  tray_icon.Create(m_server.GetHWnd());
  tray_icon.Refresh();

  int ret = m_server.Run();

  m_handler->Finalize();
  m_ui.Destroy();
  tray_icon.RemoveIcon();
  win_sparkle_cleanup();

  return ret;
}

void WeaselServerApp::SetupMenuHandlers() {
  std::filesystem::path dir = install_dir();
  m_server.AddMenuHandler(ID_WEASELTRAY_QUIT,
                          [this] { return m_server.Stop() == 0; });
  m_server.AddMenuHandler(ID_WEASELTRAY_DEPLOY,
                          std::bind(execute, dir / L"WeaselDeployer.exe",
                                    std::wstring(L"/deploy")));
  m_server.AddMenuHandler(
      ID_WEASELTRAY_SETTINGS,
      std::bind(execute, dir / L"WeaselDeployer.exe", std::wstring()));
  m_server.AddMenuHandler(
      ID_WEASELTRAY_DICT_MANAGEMENT,
      std::bind(execute, dir / L"WeaselDeployer.exe", std::wstring(L"/dict")));
  m_server.AddMenuHandler(
      ID_WEASELTRAY_SYNC,
      std::bind(execute, dir / L"WeaselDeployer.exe", std::wstring(L"/sync")));
  m_server.AddMenuHandler(ID_WEASELTRAY_HOMEPAGE,
                          std::bind(open, L"https://rime.im/"));
  m_server.AddMenuHandler(ID_WEASELTRAY_CHECKUPDATE, check_update);
  m_server.AddMenuHandler(ID_WEASELTRAY_INSTALLDIR, std::bind(explore, dir));
  m_server.AddMenuHandler(ID_WEASELTRAY_USERCONFIG,
                          std::bind(explore, WeaselUserDataPath()));
  m_server.AddMenuHandler(ID_WEASELTRAY_LOGDIR,
                          std::bind(explore, WeaselLogPath()));
}

void WeaselServerApp::EnqueueSystemCommand(
    const SystemCommandLaunchRequest& request) {
  {
    std::lock_guard<std::mutex> lock(m_launcher_mutex);
    m_launch_queue.push(request);
  }
  m_launcher_cv.notify_one();
}

void WeaselServerApp::StartLauncherWorker() {
  m_launcher_stop = false;
  m_launcher_thread = std::thread([this]() { RunLauncherWorker(); });
}

void WeaselServerApp::StopLauncherWorker() {
  {
    std::lock_guard<std::mutex> lock(m_launcher_mutex);
    m_launcher_stop = true;
  }
  m_launcher_cv.notify_all();
  if (m_launcher_thread.joinable()) {
    m_launcher_thread.join();
  }
}

void WeaselServerApp::RunLauncherWorker() {
  while (true) {
    SystemCommandLaunchRequest request;
    {
      std::unique_lock<std::mutex> lock(m_launcher_mutex);
      m_launcher_cv.wait(lock, [this]() {
        return m_launcher_stop || !m_launch_queue.empty();
      });
      if (m_launch_queue.empty()) {
        if (m_launcher_stop) {
          return;
        }
        continue;
      }
      request = m_launch_queue.front();
      m_launch_queue.pop();
    }
    ExecuteSystemCommand(request);
  }
}

bool WeaselServerApp::ExecuteSystemCommand(
    const SystemCommandLaunchRequest& request) {
  const std::string& command_id = request.command_id;
  if (command_id == "jsq" || command_id == "calc") {
    return execute(L"calc.exe", std::wstring());
  }
  if (command_id == "notepad") {
    return execute(L"notepad.exe", std::wstring());
  }
  if (command_id == "mspaint") {
    return execute(L"mspaint.exe", std::wstring());
  }
  if (command_id == "explorer") {
    return execute(L"explorer.exe", std::wstring());
  }
  if (command_id == "gh") {
    return open(L"https://github.com");
  }
  if (command_id == "bd") {
    return open(L"https://www.baidu.com");
  }
  if (command_id == "wb") {
    return open(L"https://weibo.com");
  }
  if (command_id == "g") {
    return open(L"https://google.com");
  }
  if (command_id == "kb") {
    return open(L"https://www.baidu.com");
  }
  if (command_id == "yt") {
    return open(L"https://youtube.com");
  }
  if (command_id == "rili" || command_id == "calendar" || command_id == "cal") {
    return open(L"outlookcal:");
  }
  if (command_id == "txt" || command_id == "md") {
    const wchar_t* extension = command_id == "md" ? L"md" : L"txt";
    for (const fs::path& dir : BuildOutputDirCandidates(request)) {
      for (int suffix = 0; suffix < 100; ++suffix) {
        const fs::path path = BuildDocumentPath(dir, extension, suffix);
        if (!CreateEmptyDocument(path)) {
          continue;
        }
        const std::wstring args = L"\"" + path.wstring() + L"\"";
        return execute(L"notepad.exe", args);
      }
    }
  }
  return false;
}
