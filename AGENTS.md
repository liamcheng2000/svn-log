# AGENTS.md

Windows 桌面工具（中文界面）：tkinter + openpyxl，通过 subprocess 调用 `svn.exe`。GUI 固定使用标准库 tkinter，运行期依赖仅 openpyxl。无服务端，也没有 lint/typecheck 配置。`dist/`、`build/`、`*.spec` 是生成物，已 gitignore。

## 常用命令

- 安装依赖：`python -m pip install -r requirements.txt`（仅 openpyxl；打包另需 PyInstaller）
- 源码运行：`python main.py`（或 `run.bat`）
- 全部测试：`python -m unittest discover -s tests -v`（不访问网络 / SVN）
- 单个测试：`python -m unittest tests.test_core.ExcelTests.test_rich_message_alignment_and_height -v`
- 打包：不要直接运行 `build_exe.bat`（结尾有 `pause`，非交互会挂住），执行：

```bat
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "SVN日志导出" --version-file version_info.txt --exclude-module numpy --add-data "docs\user_manual.md;docs" --paths . main.py
```

产物 `dist\SVN日志导出.exe`（约 24 MB）。构建机为 Python 3.14 + PyInstaller 6.21。

## 版本与发布

- 版本号涉及 4 个文件，必须同步：`app/config.py` 的 `APP_VERSION`、`version_info.txt`（`filevers`/`prodvers` + `FileVersion`/`ProductVersion`）、`docs/user_manual.md` 顶部与版本记录、`README.md` 当前版本与版本记录。
- `tests/test_help.py` 会校验说明书包含 `v{APP_VERSION}`，漏改会测试失败。
- 当前发布方式：直接提交 `main`，打 annotated tag `vX.Y.Z`，推送 `main` 和标签（没有 PR 流程）。
- 任何用户可见的界面/行为变更，都要同步更新 `docs/user_manual.md`（帮助窗口内容）和 `README.md`。

## 打包硬性要求

- 必须 `--exclude-module numpy`：构建机的 NumPy 2.x 会经 openpyxl 的可选导入被打进 exe，在不支持 AVX2 的老 CPU 上启动即崩溃（已实际发生）。openpyxl 去掉 NumPy 可正常工作。
- 必须 `--add-data "docs\user_manual.md;docs"`，否则帮助按钮读不到说明书。运行时查找顺序：exe 同级 `docs\` → 打包资源 `_MEIPASS/docs` → 源码根目录 `docs/`。
- `.bat` 由 `.gitattributes` 固定为 CRLF，改批处理时不要用 LF 重写。

## 实现细节（容易踩坑）

- SVN 客户端：`app/svn_client.py` 自动定位 `svn.exe`（PATH → 环境变量 `SVN_EXE` → TortoiseSVN 注册表 → 常见目录），固定 `--non-interactive`，复用 SVN 自身凭据缓存。不要加 `--no-auth-cache`，也不要改成交互式登录。
- 版本范围两项留空 = `--limit 10`（最近 10 条），不是全量历史；日期模式结束边界是当天 `T23:59:59`。
- 查询在后台线程执行，结果经 `queue.Queue` 回主线程 `_poll_queue`；不要在工作线程直接操作 Tk 控件。
- Excel 富文本：`CellRichText` 的 run 不能只含空白，换行必须并入下一个 `TextBlock`，否则 Excel 打开时提示「字符串属性」并修复、丢失加粗和换行。测试已断言 XML 不含 `<t>\s*</t>` 且带 `xml:space="preserve"`。
- 日历弹窗必须锚定被点击的「日历」按钮，并用 `_clamp_popup_position` 做屏幕边界收敛；不要按主窗口定位（曾导致小屏跑到屏幕外）。
- 偏好保存在 `%APPDATA%\SVNLogApp\config.json`，禁止把密码/凭据写入其中。
- 所有源码/文档为 UTF-8；Windows 控制台里中文或文件名显示乱码通常只是代码页问题，不代表文件损坏。

## 结构

- `main.py`：入口（DPI 适配、异常弹窗）。
- `app/ui.py`：主窗口、日历弹窗、后台线程与消息循环。
- `app/svn_client.py`：svn.exe 定位、命令构造、XML 解析、错误翻译。
- `app/excel_writer.py`：三页报表、富文本与列样式。
- `app/help_window.py`：Markdown 使用说明窗口。
- `tests/`：普通 unittest 模块，可直接运行；不创建 Tk 窗口。
- 界面文案与用户文档保持中文。
