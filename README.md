# OneNote Web → PDF 批量导出

用 Playwright 在现有 X11 桌面上打开 OneNote 网页版，由你手动登录，然后按
**笔记本 → 分区组 → 分区 → 页面/子页面** 的层次逐页调用 OneNote 自带的
**“文件 → 打印”**，把每个页面保存为 PDF。

不修改笔记内容，不打印浏览器界面，也不使用 `Ctrl+P` 或浏览器菜单。

## 为什么不用现成项目

| 项目 | 限制 |
|---|---|
| `enoola/Microsoft-OneNote-Exporter` | 导出 Markdown/附件，不生成逐页 PDF |
| `jli-software/OneNoteBackupExporter` | 依赖 Windows + OneNote 桌面 COM API |
| `jwhco/OneNoteExporterAkaPublisher` | PowerShell + Windows 桌面版 |

本工具针对 Linux + 现有 X11 + OneNote 网页版。

## 环境

```bash
cd onenote-pdf
uv sync
uv run playwright install chromium
```

`uv` 管理 Python 环境和依赖；Chromium 由 Playwright 下载。使用现有 X11 桌面
时设置 `DISPLAY=:0`（或你的实际显示号），不要用 Xvfb。

## 登录

```bash
DISPLAY=:0 uv run python onenote_pdf.py login
```

会在现有 X11 屏幕打开 Chromium。完成微软登录/MFA 后，登录状态保存在
`.profile/`。工具不会读取或保存密码。

## 导出

最省事的方式是直接给笔记本 URL：

```bash
DISPLAY=:0 uv run python onenote_pdf.py export \
  --url 'https://onedrive.live.com/...' \
  --output output
```

如果浏览器里已经打开了笔记本，也可以让工具自动识别当前页面：

```bash
DISPLAY=:0 uv run python onenote_pdf.py export --output output
```

### 默认纸张

默认 **A0 横向**，适合 OneNote 的超宽无限画布：

```bash
# 默认就是 A0 横向
uv run python onenote_pdf.py export --url '...' --output output

# 改纸张 / 纵向
uv run python onenote_pdf.py export --url '...' --output output --paper A2
uv run python onenote_pdf.py export --url '...' --output output --paper A4 --portrait
```

可用纸张：`A0 A1 A2 A3 A4 Letter Legal`。

OneNote 的一个页面可能被它自己的打印引擎分成多张 A0 PDF 页；工具不会重排或
裁切笔记内容。

## 打印重试逻辑

每个 OneNote 页面最多尝试 **5 次**：

1. 选中页面后直接调用 OneNote「文件 → 打印」。
2. 如果 OneNote 提示“此页面上的某些图像仍在下载中”，或打印按钮持续灰色，
   关闭提示，把页面滚动到最底部，再滚动回最顶部。
3. 重新尝试打印。
4. 5 次仍失败则记录错误并以非零状态结束，不会把空文件当成成功。

滚动使用固定步长（400px / 80ms），用来触发 OneNote 的懒加载图片。每页最多
5 次，避免单个坏页面卡住整个批量任务。

## 输出结构

```text
output/
  └── 示例笔记本 [01234567]/
      ├── 示例分区/
      │   ├── 01 页面一 [89abcdef]/
      │   │   └── 01 页面一 [89abcdef].pdf
      │   └── 02 页面二 [fedcba98]/
      │       └── 02 页面二 [fedcba98].pdf
      └── 分区组/
          └── 子分区/
              └── ...
```

- 目录层级保留笔记本、分区组、分区、子页面关系。
- 页面目录带稳定短 ID，避免同名页面互相覆盖。
- `manifest.json` 记录每页状态；重跑时跳过已成功页面，失败页面会重试。
- `INDEX.md` 是全部已导出页面的索引，按分区列出，可直接点击打开 PDF。
- 笔记本目录的短 ID 由 OneNote 的 `sourcedoc` 生成，不随 URL 查询参数变化。
- 想强制重导：加 `--no-resume`。

## 连接已启动的 Chromium

如果不想让工具自己启动浏览器，可以连接一个已按下面要求启动的 Chromium：

```bash
chromium --remote-debugging-port=9223 \
  --user-data-dir=.profile \
  --kiosk-printing
```

然后：

```bash
uv run python onenote_pdf.py export \
  --cdp-url http://127.0.0.1:9223 \
  --print-dir .print-staging \
  --output output
```

`--print-dir` 必须是专用暂存目录：工具会在每次打印前删除其中的 `*.pdf`
临时文件，避免把旧文件误当成新结果。

注意：`--cdp-url` 模式下工具不会修改浏览器偏好，浏览器必须已经：

- 以 `--kiosk-printing` 启动；
- 默认打印目标为“另存为 PDF”；
- 保存目录与 `--print-dir` 一致。

工具自己启动浏览器时会自动写入这些偏好。

## 常用参数

```text
--url URL             笔记本 URL；不传则用当前已打开的笔记本
--output DIR          输出根目录（默认 output）
--profile DIR         持久 Chromium profile（默认 .profile）
--print-dir DIR       Chromium “另存为 PDF”暂存目录（默认 .print-staging）
--paper {A0,A1,...}   纸张（默认 A0）
--portrait            纵向（默认横向）
--max-pages N         最多成功导出 N 页，便于试跑
--section NAME        只导出名称/ID 匹配的分区
--no-resume           忽略 manifest，重新导出
--login-timeout SEC   等待人工登录/打开笔记本的时间（默认 900）
--nav-timeout SEC     等待 OneNote 导航的时间（默认 120）
--print-timeout SEC   单页打印总超时（默认 180；内部按 5 次重试）
```

## 调试

```bash
uv run python onenote_pdf.py inspect --cdp-url http://127.0.0.1:9223
```

只输出 frame、role、选择器计数和页面标题，不转储笔记正文。

## 已知限制

- 只支持 OneNote for the web；OneNote 桌面版请用 Windows COM 方案。
- 加密分区无法自动解锁，会记录错误并继续其他分区。
- OneNote 新主页的笔记本列表可能延迟加载；批量导出整个账户时建议逐个传
  `--url`，或先在浏览器中打开目标笔记本。
- 打印预览生成可能很慢（大页面几十秒），期间打印按钮会是灰色；工具会等待，
  超时后按上面的滚动重试逻辑处理。
- PDF 的分页由 OneNote 自己的打印引擎决定；本工具只负责逐页触发和保存。
