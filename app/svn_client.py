"""SVN 客户端封装：定位 svn.exe、执行 log 查询、解析 XML、翻译错误。

设计要点:
* 只用 svn 命令行（TortoiseSVN 自带），不依赖 pysvn。
* 固定 ``--non-interactive``，避免 svn 弹窗把界面卡死。
* 认证依赖 SVN 自身的凭据缓存；失败时引导用户打开认证终端。
"""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
import threading
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from typing import Callable, Iterable

from .models import FileChange, LogEntry

CREATE_NO_WINDOW = 0x08000000


class SvnError(Exception):
    """带中文提示的 SVN 异常。"""

    def __init__(self, message: str, raw: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.raw = raw


class Cancelled(Exception):
    """用户主动取消。"""


# --------------------------------------------------------------------------- #
# svn.exe 定位
# --------------------------------------------------------------------------- #

_CANDIDATE_DIRS = (
    r"C:\Program Files\TortoiseSVN\bin",
    r"C:\Program Files (x86)\TortoiseSVN\bin",
    r"C:\Program Files\SlikSvn\bin",
    r"C:\Program Files (x86)\SlikSvn\bin",
    r"C:\Program Files\VisualSVN Server\bin",
    r"C:\Program Files (x86)\VisualSVN Server\bin",
)


def _from_registry() -> str:
    try:
        import winreg
    except ImportError:  # 非 Windows
        return ""

    keys = (
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\TortoiseSVN"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\TortoiseSVN"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\TortoiseSVN"),
    )
    for root, path in keys:
        try:
            with winreg.OpenKey(root, path) as key:
                for value_name in ("Directory", "InstallPath", "ProcPath"):
                    try:
                        value = winreg.QueryValueEx(key, value_name)[0]
                    except OSError:
                        continue
                    if not isinstance(value, str) or not value:
                        continue
                    candidates = (
                        value,
                        os.path.join(value, "bin", "svn.exe"),
                        os.path.join(os.path.dirname(value), "bin", "svn.exe"),
                    )
                    for candidate in candidates:
                        if os.path.isfile(candidate) and candidate.lower().endswith(".exe"):
                            return candidate
        except OSError:
            continue
    return ""


def find_svn_exe(preferred: str = "") -> str:
    """按 偏好路径 → SVN_EXE → PATH → 注册表 → 常见目录 的顺序查找。"""
    if preferred and os.path.isfile(preferred):
        return preferred

    env_path = os.environ.get("SVN_EXE", "")
    if env_path and os.path.isfile(env_path):
        return env_path

    found = shutil.which("svn") or shutil.which("svn.exe")
    if found:
        return found

    from_registry = _from_registry()
    if from_registry:
        return from_registry

    for folder in _CANDIDATE_DIRS:
        candidate = os.path.join(folder, "svn.exe")
        if os.path.isfile(candidate):
            return candidate

    return ""


# --------------------------------------------------------------------------- #
# 命令行构造
# --------------------------------------------------------------------------- #

def parse_patterns(text: str) -> list[str]:
    """把 ".cpp;.h" / "*.cpp, h" 之类的输入规范化成 ['.cpp', '.h']。"""
    normalized = text.replace(",", ";").replace("，", ";").replace(" ", ";").replace("\t", ";")
    patterns: list[str] = []
    for raw in normalized.split(";"):
        item = raw.strip().lower()
        if not item:
            continue
        if item.startswith("*."):
            item = item[1:]
        if not item.startswith("."):
            item = "." + item
        if item != "." and item not in patterns:
            patterns.append(item)
    return patterns


def build_rev_spec(mode: str, start: str, end: str) -> str | None:
    """生成 ``-r`` 参数；返回 None 表示查询全部历史。

    * date 模式: ``{YYYY-MM-DDT00:00:00}:{YYYY-MM-DDT23:59:59}``
    * rev  模式: ``N:M``，留空用 1:HEAD
    """
    start = start.strip()
    end = end.strip()

    if mode == "date":
        if not start and not end:
            return None
        start_part = f"{{{start}T00:00:00}}" if start else "1"
        end_part = f"{{{end}T23:59:59}}" if end else "HEAD"
        return f"{start_part}:{end_part}"

    if not start and not end:
        return None
    start_part = start.upper() if start.upper() == "HEAD" else start
    end_part = end.upper() if end.upper() == "HEAD" else end
    return f"{start_part or '1'}:{end_part or 'HEAD'}"


def build_log_command(
    svn_exe: str,
    url: str,
    rev_spec: str | None,
    limit: int | None = None,
) -> list[str]:
    command = [svn_exe, "log", "--xml", "--verbose", "--non-interactive"]
    if rev_spec:
        command += ["-r", rev_spec]
    if limit:
        # 不指定 -r 时，--limit N 等价于「最近 N 条日志」，即 HEAD-N ~ HEAD
        command += ["--limit", str(limit)]
    command.append(url)
    return command


# --------------------------------------------------------------------------- #
# 错误翻译
# --------------------------------------------------------------------------- #

_ERROR_HINTS: tuple[tuple[str, str], ...] = (
    (
        "E170001",
        "认证失败：服务器要求登录，或当前账号权限不足。\n"
        "请点击「打开认证终端」，在弹出的窗口里输入账号密码完成一次登录（SVN 会缓存凭据），然后重试。",
    ),
    (
        "E215004",
        "认证失败：没有提供有效凭据。\n请点击「打开认证终端」手动登录一次后重试。",
    ),
    (
        "E730061",
        "认证失败：凭据被服务器拒绝。\n请点击「打开认证终端」重新登录。",
    ),
    (
        "E175002",
        "无法连接服务器：地址错误、端口不通、被代理/VPN 拦截，或服务器未启动。",
    ),
    (
        "E170013",
        "无法连接服务器：请检查 SVN 地址、网络、代理或 VPN 设置。",
    ),
    (
        "E731001",
        "主机名无法解析：请检查网络或 DNS 设置。",
    ),
    (
        "E230001",
        "服务器证书校验失败。请确认服务器可信，并在 TortoiseSVN 中永久接受该证书后重试。",
    ),
    (
        "E195020",
        "服务器证书校验失败。请在 TortoiseSVN 中永久接受该证书后重试。",
    ),
    (
        "E160013",
        "路径不存在：SVN 地址或版本号有误。",
    ),
    (
        "E195012",
        "找不到该路径：SVN 地址或版本号不存在。",
    ),
    (
        "E205000",
        "版本号无效：请检查起始/结束版本是否为数字或 HEAD。",
    ),
    (
        "E200009",
        "SVN 地址格式错误，请使用 https:// 或 http:// 开头的完整地址。",
    ),
    (
        "E000111",
        "SVN 地址无法解析，请检查是否缺少协议头或包含非法字符。",
    ),
)


def describe_error(raw: str) -> str:
    text = (raw or "").strip()
    lowered = text.lower()
    for code, hint in _ERROR_HINTS:
        if code.lower() in lowered:
            return hint
    if "authorization failed" in lowered or "authentication failed" in lowered:
        return _ERROR_HINTS[0][1]
    if "unable to connect" in lowered:
        return _ERROR_HINTS[3][1]
    if "not a working copy" in lowered:
        return "目标不是有效的 SVN 地址，请填写仓库 URL（而不是本地工作副本路径）。"
    if not text:
        return "SVN 执行失败，且没有返回错误信息。"
    return "SVN 执行失败：\n" + text[:1500]


# --------------------------------------------------------------------------- #
# 执行与解析
# --------------------------------------------------------------------------- #

def _run_capture(command: list[str], cancel_event: threading.Event | None = None) -> tuple[int, bytes, str]:
    """执行命令，返回 (返回码, stdout 字节, stderr 文本)。stderr 走临时文件防止管道死锁。"""
    with tempfile.TemporaryFile() as err_file:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=err_file,
            creationflags=CREATE_NO_WINDOW if os.name == "nt" else 0,
        )
        try:
            out, _ = process.communicate(timeout=None)
        except BaseException:
            process.kill()
            raise
        finally:
            if cancel_event is not None and cancel_event.is_set():
                process.kill()
        err_file.seek(0)
        stderr_text = err_file.read().decode("utf-8", errors="replace")
    return process.returncode, out or b"", stderr_text


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _child(element: ET.Element, name: str) -> ET.Element | None:
    for item in element:
        if _local_name(item.tag) == name:
            return item
    return None


def _parse_date(text: str) -> datetime:
    """svn 输出 UTC 时间，例如 2024-06-30T02:12:33.123456Z。"""
    raw = (text or "").strip()
    if not raw:
        return datetime.now()
    if raw.endswith("Z"):
        raw = raw[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(raw)
    except ValueError:
        try:
            parsed = datetime.strptime(raw[:19], "%Y-%m-%dT%H:%M:%S")
        except ValueError:
            return datetime.now()
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone().replace(tzinfo=None)


def parse_log_xml(payload: bytes) -> list[LogEntry]:
    if not payload.strip():
        return []
    try:
        root = ET.fromstring(payload)
    except ET.ParseError as exc:
        raise SvnError(f"无法解析 SVN 返回的日志数据：{exc}") from exc

    entries: list[LogEntry] = []
    for node in root:
        if _local_name(node.tag) != "logentry":
            continue
        try:
            revision = int(node.get("revision", "0"))
        except ValueError:
            continue

        author_node = _child(node, "author")
        date_node = _child(node, "date")
        message_node = _child(node, "msg")
        entry = LogEntry(
            revision=revision,
            author=(author_node.text or "").strip() if author_node is not None else "(未知)",
            date=_parse_date(date_node.text or "" if date_node is not None else ""),
            message=(message_node.text or "").strip() if message_node is not None else "",
        )

        paths_node = _child(node, "paths")
        if paths_node is not None:
            for path_node in paths_node:
                if _local_name(path_node.tag) != "path":
                    continue
                if (path_node.get("kind") or "").lower() != "file":
                    continue  # 只保留文件，忽略目录
                entry.files.append(
                    FileChange(
                        path=(path_node.text or "").strip(),
                        action=(path_node.get("action") or "M").upper(),
                    )
                )
        entries.append(entry)

    entries.sort(key=lambda item: item.revision, reverse=True)
    return entries


def filter_entries(
    entries: Iterable[LogEntry],
    patterns: list[str],
    keep_empty: bool = False,
) -> list[LogEntry]:
    """按扩展名过滤文件；keep_empty=False 时丢弃没有匹配文件的版本。"""
    result: list[LogEntry] = []
    for entry in entries:
        if patterns:
            matched = [f for f in entry.files if f.ext in patterns]
        else:
            matched = list(entry.files)
        if not matched and not keep_empty:
            continue
        result.append(
            LogEntry(
                revision=entry.revision,
                author=entry.author,
                date=entry.date,
                message=entry.message,
                files=matched,
            )
        )
    return result


def fetch_log(
    svn_exe: str,
    url: str,
    rev_spec: str | None,
    limit: int | None = None,
    on_progress: Callable[[int], None] | None = None,
    cancel_event: threading.Event | None = None,
) -> list[LogEntry]:
    """执行 svn log 并返回解析后的日志；支持实时进度与取消。"""
    command = build_log_command(svn_exe, url, rev_spec, limit)
    with tempfile.TemporaryFile() as err_file:
        process = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=err_file,
            creationflags=CREATE_NO_WINDOW if os.name == "nt" else 0,
        )

        chunks: list[bytes] = []
        count = 0
        tail = b""
        try:
            assert process.stdout is not None
            while True:
                if cancel_event is not None and cancel_event.is_set():
                    process.kill()
                    raise Cancelled()
                chunk = process.stdout.read(8192)
                if not chunk:
                    break
                chunks.append(chunk)
                data = tail + chunk
                found = data.count(b"<logentry")
                if found:
                    count += found
                    if on_progress is not None:
                        on_progress(count)
                tail = data[-8:]
            process.wait()
        finally:
            if process.poll() is None:
                process.kill()
                process.wait()

        err_file.seek(0)
        stderr_text = err_file.read().decode("utf-8", errors="replace")

    if cancel_event is not None and cancel_event.is_set():
        raise Cancelled()

    if process.returncode != 0:
        raise SvnError(describe_error(stderr_text), stderr_text)

    payload = b"".join(chunks)
    if on_progress is not None and count == 0:
        on_progress(0)
    return parse_log_xml(payload)


def test_connection(svn_exe: str, url: str) -> str:
    """svn info 探测，返回一句人类可读的结论。"""
    command = [svn_exe, "info", "--xml", "--non-interactive", "-r", "HEAD", url]
    code, out, err = _run_capture(command)
    if code != 0:
        raise SvnError(describe_error(err), err)

    try:
        root = ET.fromstring(out)
    except ET.ParseError:
        return "连接成功（无法解析版本信息）。"

    entry = None
    for node in root:
        if _local_name(node.tag) == "entry":
            entry = node
            break
    if entry is None:
        return "连接成功。"

    revision = entry.get("revision", "?")
    kind = entry.get("kind", "")
    commit = _child(entry, "commit")
    author = ""
    when = ""
    if commit is not None:
        author_node = _child(commit, "author")
        date_node = _child(commit, "date")
        author = (author_node.text or "").strip() if author_node is not None else ""
        when = _parse_date(date_node.text or "").strftime("%Y-%m-%d %H:%M") if date_node is not None else ""
    suffix = f"，最后提交 {when} by {author}" if when else ""
    return f"连接成功：{kind or '仓库'}，HEAD = r{revision}{suffix}"


def open_auth_terminal(svn_exe: str, url: str) -> None:
    """弹出真实控制台执行一次 svn log，让用户手动登录并缓存凭据。"""
    # 批处理里 % 需要转义成 %%，否则 URL 里的 %20 之类会被 cmd 当变量展开
    safe_svn = svn_exe.replace('"', "").replace("%", "%%")
    safe_url = url.replace('"', "").replace("%", "%%")
    bat_content = (
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        "title SVN 认证 - 输入账号密码后等待命令结束\r\n"
        "echo 即将访问: " + safe_url + "\r\n"
        "echo 若提示认证，请输入 SVN 账号密码；完成后关闭本窗口即可。\r\n"
        "echo.\r\n"
        f'"{safe_svn}" log --limit 1 -r HEAD "{safe_url}"\r\n'
        "echo.\r\n"
        "echo 命令结束。凭据已由 SVN 缓存，可回到工具重试。\r\n"
        "pause\r\n"
    )
    handle, path = tempfile.mkstemp(prefix="svn_auth_", suffix=".bat")
    with os.fdopen(handle, "w", encoding="utf-8", newline="") as fp:
        fp.write(bat_content)
    if os.name == "nt":
        os.startfile(path)  # noqa: S606 - 明确由用户触发
    else:
        subprocess.Popen(["sh", path])
