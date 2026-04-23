## AI Assistant

当前代码里，AI 相关能力分成两条主线：

### 1. React 面板链路

- 按 `ai_assistant.trigger_hotkey` 打开远端 React 面板，默认 `Control+3`。
- Host 负责：
  - 收集输入上下文
  - 维护登录态、`token`、`tenantId`、用户信息缓存
  - 拉取 AI 指令/机构列表
  - 接收前端确认后的回写文本
- 主请求由前端面板自己发起，Host 不再直连旧的 `endpoint`。

### 2. TSF 内联 AI 链路

- 输入 `ai_assistant.instruction_lookup_prefix`（默认 `sS`）后，走动态 AI 指令检索。
- 选中某个 AI 指令后，进入 TSF 内联 preedit。
- 内联前缀会带上指令名，例如 `【续写】，`。
- 内联提交后，Host 会调用：
  - `ai_assistant.ai_api_base + "/chat-messages"`
  - `Authorization: Bearer {指令 appKey}`
  - `inputs.Token = 登录 token`
  - `inputs.tenantid = 当前 tenantId`
  - `user = 用户信息里的 id`
  - `query = 内联录入内容`
- 接口按流式返回时，内容直接上屏；结束后退出内联状态。

### 当前配置示例

```yaml
ai_assistant:
  enabled: true

  trigger_hotkey: "ctrl+3"
  instruction_lookup_prefix: "sS"
  inline_instruction_enabled: true
  inline_instruction_prefix: "/"

  ai_api_base: "https://copilot.sino-bridge.com:90/v1"

  login_required: true
  login_url: "https://copilot.sino-bridge.com:85?uuid={uuid}&fromType=plugIn&operationType=login"
  login_state_path: "ai_login_state.json"
  login_token_key: "token"
  refresh_token_endpoint: ""
  mqtt_url: "wss://copilot.sino-bridge.com:85/mqttSocket/mqtt"
  mqtt_topic_template: "/mqtt/topic/sino/lamp/oauth/token/login/{uuid}"
  mqtt_username: ""
  mqtt_password: ""
  mqtt_timeout_ms: 120000

  panel_url: "https://copilot.sino-bridge.com/toolbox/#/rime-with-weasel"
  panel_allowed_origin: "https://copilot.sino-bridge.com"

  max_history_chars: 2048
  timeout_ms: 30000

  debug_dump_context: true
  debug_dump_path: "ai_context_dump.txt"
```

### 配置说明

- `enabled`：AI 总开关。
- `trigger_hotkey`：打开 React 面板的快捷键。
- `instruction_lookup_prefix`：AI 指令检索前缀，当前 `sS` 这条链路会进入“候选词选指令”。
- `inline_instruction_enabled`：是否启用 TSF 内联 AI。
- `inline_instruction_prefix`：内联显示前缀字符，默认 `/`；当前主要用于 preedit 展示，不是主要入口。
- `ai_api_base`：内联 AI 请求基地址，提交时会拼成 `/chat-messages`。
- `login_required` / `login_url` / `login_state_path` / `login_token_key`：登录流程相关配置。
- `refresh_token_endpoint`：刷新 token 接口；留空时会根据登录地址推导候选刷新地址。
- `mqtt_*`：扫码登录/消息通知相关配置。
- `panel_url`：React 面板地址。
- `panel_allowed_origin`：允许和 Host 通信的前端源；为空时会从 `panel_url` 推导。
- `max_history_chars`：收集上下文时保留的最大历史字符数。
- `timeout_ms`：网络请求超时，登录、用户信息、机构列表、内联 AI 都会用到。
- `debug_dump_context` / `debug_dump_path`：是否把上下文落盘，便于联调。

### 登录与缓存

- 登录成功后，身份信息和用户信息会落盘缓存。
- 重启电脑后缓存仍保留。
- 当相关接口返回 `401` 时，会清理失效状态并触发重新登录。

### Panel 消息协议

Host -> Panel（WebView `message` 事件）：

```json
{
  "v": "1.0",
  "type": "host.sync",
  "payload": {
    "context": "当前引用输入内容",
    "status": "上下文已就绪，请在前端面板中发起请求。",
    "output": "",
    "requesting": false,
    "institutionsLoading": false,
    "selectedInstitutionId": "",
    "auth": {
      "token": "",
      "tenantId": "",
      "refreshToken": ""
    },
    "institutions": [
      { "id": "1", "name": "Agent A", "appKey": "app-xxx" }
    ]
  }
}
```

Panel -> Host（`window.chrome.webview.postMessage(...)`）：

- `ui.ready`
- `ui.context.changed`
- `ui.select.institution`
- `ui.writeback.confirm`
- `ui.cancel`
- `ui.drag.start`
- `ui.panel.resize`
- `ui.auth.refresh_request`
- `ui.request`
- `ui.confirm`
- `ui.system_command`

说明：

- `ui.request` 现在只用于把上下文/状态同步给 Host，真正的面板请求仍由前端自己完成。
- `ui.writeback.confirm` 才会把最终文本回写到目标输入框。

### React 端最小对接

接收 Host 同步：

```ts
window.chrome?.webview?.addEventListener("message", (event: any) => {
  const message = typeof event.data === "string"
    ? JSON.parse(event.data)
    : event.data
  if (message?.type === "host.sync") {
    // 使用 message.payload.context / auth / institutions
  }
})
```

发送回写：

```ts
window.chrome?.webview?.postMessage(JSON.stringify({
  type: "ui.writeback.confirm",
  payload: { text: finalText }
}))
```

### 已下线配置

下面这些旧 Host 直连生成配置已经下线，不要再配置：

- `ai_assistant/stream`
- `ai_assistant/endpoint`
- `ai_assistant/api_key`
- `ai_assistant/model`
- `ai_assistant/reasoning_effort`

### 已知限制

- Rime 仍无法直接读取任意应用里“已经存在的全文”，上下文只能基于输入法可观测内容重建。
- 内联 AI 期间会屏蔽再次触发 AI 指令和系统指令候选，避免冲突。
