## AI Assistant (Remote React Panel)

当前实现为「Host 负责触发与状态同步，React 面板负责流式请求」。

### 行为

- 按 `ai_assistant.trigger_hotkey` 触发面板（默认 `Ctrl + 3`）。
- 若当前有未上屏 preedit，会先上屏，避免内容丢失。
- Host 只负责：
  - 收集上下文（输入法历史 + 当前预编辑）
  - 处理登录态与 token/tenantId/refreshToken
  - 接收面板确认回写并写回目标输入框
- AI 流式请求在 React 面板内完成（Host 不再发起主请求）。

### `weasel.yaml` 配置示例

```yaml
ai_assistant:
  enabled: true
  trigger_hotkey: "Control+3"
  instruction_lookup_prefix: "sS"
  login_required: true
  debug_dump_context: false
  debug_dump_path: "ai_context_dump.txt"

  panel_url: "https://copilot.sino-bridge.com/toolbox/#/rime-with-weasel"
  panel_allowed_origin: "https://copilot.sino-bridge.com"

  login_url: "https://copilot.sino-bridge.com/copilot-web-app/login?uuid={uuid}&fromType=plugIn&operationType=login"
  login_state_path: "ai_login_state.json"
  login_token_key: "token"
  refresh_token_endpoint: "https://copilot.sino-bridge.com/api/oauth/anyTenant/refresh"
  mqtt_url: "wss://copilot.sino-bridge.com/mqtt"
  mqtt_topic_template: "/mqtt/topic/sino/lamp/oauth/token/login/{uuid}"
  mqtt_username: ""
  mqtt_password: ""
  mqtt_timeout_ms: 120000

  max_history_chars: 2048
  timeout_ms: 30000
```

说明：

- `trigger_hotkey` 可配置触发快捷键，默认 `Control+3`。
- `instruction_lookup_prefix` 可配置 `sS` 这类 AI 指令检索前缀，默认 `sS`。
- `panel_url` 必填，触发快捷键命中时会直接 `Navigate(panel_url)`。
- `panel_allowed_origin` 为空时，会从 `panel_url` 自动推导 origin。
- Host 到 React 的 token/tenantId/上下文通过 `postMessage` 发送，不走 URL 参数。

### 消息协议

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
- `ui.context.changed`，payload: `{ "text": "..." }`
- `ui.select.institution`，payload: `{ "id": "..." }`
- `ui.writeback.confirm`，payload: `{ "text": "最终确认回写文本" }`
- `ui.cancel`
- `ui.drag.start`
- `ui.auth.refresh_request`

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

发送命令给 Host：

```ts
window.chrome?.webview?.postMessage(JSON.stringify({
  type: "ui.writeback.confirm",
  payload: { text: finalText }
}))
```

### 已知限制

- Rime 仍无法直接读取任意应用内“已存在全文”；上下文基于输入法可观测内容重建。
- 焦点切换时按 app 维度隔离上下文，避免跨应用串内容。
