#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import shutil
import subprocess
import sys
import time
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PIL import Image, ImageDraw
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


SERVICE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = SERVICE_DIR / "config" / "defaults.json"


def _load_config() -> dict[str, Any]:
    config_path = Path(os.environ.get("INVOICE_CHECK_CONFIG", DEFAULT_CONFIG_PATH)).expanduser()
    data: dict[str, Any] = {}
    if config_path.is_file():
        data = json.loads(config_path.read_text(encoding="utf-8"))
    env_map = {
        "adb_path": "INVOICE_CHECK_ADB_PATH",
        "vlm_api": "INVOICE_CHECK_VLM_API",
        "vlm_model": "INVOICE_CHECK_VLM_MODEL",
        "paddle_vl_api": "INVOICE_CHECK_PADDLE_VL_API",
        "paddle_vl_model": "INVOICE_CHECK_PADDLE_VL_MODEL",
        "remote_dir": "INVOICE_CHECK_REMOTE_DIR",
        "remote_image": "INVOICE_CHECK_REMOTE_IMAGE",
    }
    for key, env_name in env_map.items():
        if os.environ.get(env_name):
            data[key] = os.environ[env_name]
    numeric_envs: dict[str, tuple[str, type]] = {
        "vlm_temperature": ("INVOICE_CHECK_VLM_TEMPERATURE", float),
        "vlm_top_p": ("INVOICE_CHECK_VLM_TOP_P", float),
        "vlm_top_k": ("INVOICE_CHECK_VLM_TOP_K", int),
        "vlm_max_tokens": ("INVOICE_CHECK_VLM_MAX_TOKENS", int),
        "vlm_timeout": ("INVOICE_CHECK_VLM_TIMEOUT", float),
        "paddle_vl_temperature": ("INVOICE_CHECK_PADDLE_VL_TEMPERATURE", float),
        "paddle_vl_max_tokens": ("INVOICE_CHECK_PADDLE_VL_MAX_TOKENS", int),
        "paddle_vl_timeout": ("INVOICE_CHECK_PADDLE_VL_TIMEOUT", float),
    }
    for key, (env_name, caster) in numeric_envs.items():
        if os.environ.get(env_name):
            data[key] = caster(os.environ[env_name])
    if os.environ.get("INVOICE_CHECK_VLM_ENABLE_THINKING"):
        raw = os.environ["INVOICE_CHECK_VLM_ENABLE_THINKING"].strip().lower()
        data["vlm_enable_thinking"] = raw in {"1", "true", "yes", "on"}
    return data


CONFIG = _load_config()
ADB = Path(CONFIG.get("adb_path", "adb")).expanduser()
VLM_API = CONFIG.get("vlm_api", "http://127.0.0.1:8081/v1/chat/completions")
VLM_MODEL = CONFIG.get("vlm_model", "qwen-vl")
PADDLE_VL_API = CONFIG.get("paddle_vl_api", "http://127.0.0.1:8090/v1/chat/completions")
PADDLE_VL_MODEL = CONFIG.get("paddle_vl_model", "paddleocr-vl")
REMOTE_DIR = CONFIG.get("remote_dir", "/sdcard/Pictures/invoice-check")
REMOTE_IMAGE = CONFIG.get("remote_image", "qr_current.png")
VLM_TEMPERATURE = float(CONFIG.get("vlm_temperature", 0.01))
VLM_TOP_P = float(CONFIG.get("vlm_top_p", 0.8))
VLM_TOP_K = int(CONFIG.get("vlm_top_k", 20))
VLM_MAX_TOKENS = int(CONFIG.get("vlm_max_tokens", 512))
VLM_TIMEOUT = float(CONFIG.get("vlm_timeout", 30))
VLM_ENABLE_THINKING = bool(CONFIG.get("vlm_enable_thinking", False))
PADDLE_VL_TEMPERATURE = float(CONFIG.get("paddle_vl_temperature", 0))
PADDLE_VL_MAX_TOKENS = int(CONFIG.get("paddle_vl_max_tokens", 768))
PADDLE_VL_TIMEOUT = float(CONFIG.get("paddle_vl_timeout", 15))
EMIT_PROGRESS_EVENTS = bool(CONFIG.get("emit_progress_events", False))


@dataclass
class Candidate:
    id: int
    label_hint: str
    box: list[int]

    @property
    def center(self) -> tuple[int, int]:
        x1, y1, x2, y2 = self.box
        return (x1 + x2) // 2, (y1 + y2) // 2

    def as_json(self) -> dict[str, Any]:
        x, y = self.center
        return {"id": self.id, "label_hint": self.label_hint, "box": self.box, "center": [x, y]}


def run(cmd: list[str], *, capture: bool = False, check: bool = True, timeout: float = 30) -> subprocess.CompletedProcess[str]:
    return subprocess.run(cmd, text=True, capture_output=capture, check=check, timeout=timeout)


def adb(args: list[str], *, capture: bool = False, check: bool = True, timeout: float = 30) -> subprocess.CompletedProcess[str]:
    cmd = [str(ADB), *args]
    try:
        return run(cmd, capture=capture, check=check, timeout=timeout)
    except subprocess.TimeoutExpired:
        recover_adb_server()
        return run(cmd, capture=capture, check=check, timeout=timeout)


def recover_adb_server() -> None:
    # ADB occasionally wedges when prior commands are interrupted. Best effort only.
    subprocess.run([str(ADB), "kill-server"], text=True, capture_output=True, timeout=5, check=False)
    time.sleep(0.5)
    subprocess.run([str(ADB), "start-server"], text=True, capture_output=True, timeout=10, check=False)
    time.sleep(0.5)


def adb_shell(command: str, *, capture: bool = False, check: bool = True, timeout: float = 30) -> subprocess.CompletedProcess[str]:
    # Use one shell string so Android shell handles quoting and globs consistently.
    return adb(["shell", command], capture=capture, check=check, timeout=timeout)


def get_size() -> tuple[int, int]:
    proc = adb(["shell", "wm", "size"], capture=True)
    line = proc.stdout.strip().splitlines()[-1]
    size = line.split(":")[-1].strip()
    width, height = size.split("x")
    return int(width), int(height)


def screenshot(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("wb") as f:
        subprocess.run([str(ADB), "exec-out", "screencap", "-p"], check=True, stdout=f, timeout=20)


def tap(x: int, y: int) -> None:
    adb(["shell", "input", "tap", str(x), str(y)], timeout=10)


def keyevent(code: int) -> None:
    adb(["shell", "input", "keyevent", str(code)], timeout=10)


def swipe(x1: int, y1: int, x2: int, y2: int, duration_ms: int = 300) -> None:
    adb([
        "shell",
        "input",
        "swipe",
        str(x1),
        str(y1),
        str(x2),
        str(y2),
        str(duration_ms),
    ], timeout=10)


def stage_image(local_image: Path) -> None:
    if not local_image.is_file():
        raise FileNotFoundError(local_image)
    adb_shell(f"mkdir -p {REMOTE_DIR!r}", timeout=10)
    adb_shell(f"find {REMOTE_DIR!r} -maxdepth 1 -type f -delete", timeout=10)
    remote = f"{REMOTE_DIR}/{REMOTE_IMAGE}"
    adb(["push", str(local_image), remote], timeout=20)
    adb([
        "shell",
        "am",
        "broadcast",
        "-a",
        "android.intent.action.MEDIA_SCANNER_SCAN_FILE",
        "-d",
        f"file://{remote}",
    ], timeout=10)


def cleanup_phone_album() -> None:
    adb_shell(f"mkdir -p {REMOTE_DIR!r}", timeout=10, check=False)
    adb_shell(f"find {REMOTE_DIR!r} -maxdepth 1 -type f -delete", timeout=10, check=False)
    adb([
        "shell",
        "am",
        "broadcast",
        "-a",
        "android.intent.action.MEDIA_SCANNER_SCAN_FILE",
        "-d",
        f"file://{REMOTE_DIR}/",
    ], timeout=10, check=False)


def is_textish_pixel(rgb: tuple[int, int, int]) -> bool:
    r, g, b = rgb
    dark = r < 120 and g < 120 and b < 120
    green = g > 110 and r < 90 and b < 140
    gray = 120 <= r <= 215 and abs(r - g) < 16 and abs(g - b) < 16
    white = r > 225 and g > 225 and b > 225
    red = r > 180 and g < 90 and b < 100
    return dark or green or gray or white or red


def connected_components_text_boxes(image: Image.Image) -> list[list[int]]:
    width, height = image.size
    pix = image.load()
    visited: set[tuple[int, int]] = set()
    comps: list[list[int]] = []
    y_min = 70
    y_max = min(height, 900)
    for y in range(y_min, y_max):
        for x in range(width):
            if (x, y) in visited or not is_textish_pixel(pix[x, y]):
                continue
            stack = [(x, y)]
            visited.add((x, y))
            xs: list[int] = []
            ys: list[int] = []
            for px, py in stack:
                xs.append(px)
                ys.append(py)
                for nx in (px - 1, px, px + 1):
                    for ny in (py - 1, py, py + 1):
                        if (
                            0 <= nx < width
                            and y_min <= ny < y_max
                            and (nx, ny) not in visited
                            and is_textish_pixel(pix[nx, ny])
                        ):
                            visited.add((nx, ny))
                            stack.append((nx, ny))
            if len(xs) >= 8:
                x1, y1, x2, y2 = min(xs), min(ys), max(xs), max(ys)
                if (x2 - x1) >= 2 and (y2 - y1) >= 4:
                    comps.append([x1, y1, x2, y2])

    lines: list[list[Any]] = []
    for x1, y1, x2, y2 in sorted(comps, key=lambda b: (b[1], b[0])):
        if y2 - y1 > 90 or x2 - x1 > width * 0.9:
            continue
        placed = False
        cy = (y1 + y2) / 2
        for line in lines:
            lx1, ly1, lx2, ly2, items = line
            lcy = (ly1 + ly2) / 2
            if abs(cy - lcy) < 18:
                items.append([x1, y1, x2, y2])
                line[0] = min(lx1, x1)
                line[1] = min(ly1, y1)
                line[2] = max(lx2, x2)
                line[3] = max(ly2, y2)
                placed = True
                break
        if not placed:
            lines.append([x1, y1, x2, y2, [[x1, y1, x2, y2]]])

    boxes: list[list[int]] = []
    for _lx1, _ly1, _lx2, _ly2, items in lines:
        run: list[list[int]] = []
        last: list[int] | None = None
        for item in sorted(items, key=lambda b: b[0]):
            if last is not None and item[0] - last[2] > 42:
                if run:
                    boxes.append(_padded_run_box(run, width, height))
                run = []
            run.append(item)
            last = item
        if run:
            boxes.append(_padded_run_box(run, width, height))
    return boxes


def _padded_run_box(run: list[list[int]], width: int, height: int) -> list[int]:
    x1 = min(c[0] for c in run)
    y1 = min(c[1] for c in run)
    x2 = max(c[2] for c in run)
    y2 = max(c[3] for c in run)
    return [max(0, x1 - 16), max(0, y1 - 14), min(width, x2 + 16), min(height, y2 + 14)]


def color_button_boxes(image: Image.Image) -> list[tuple[str, list[int]]]:
    width, height = image.size
    pix = image.load()
    boxes: list[tuple[str, list[int]]] = []
    rows: list[tuple[int, int, int]] = []
    for y in range(height):
        xs: list[int] = []
        for x in range(width):
            r, g, b = pix[x, y]
            if g > 145 and r < 90 and b < 140:
                xs.append(x)
        if len(xs) > 45:
            rows.append((min(xs), y, max(xs)))
    if rows:
        groups: list[list[tuple[int, int, int]]] = []
        for row in rows:
            if groups and row[1] - groups[-1][-1][1] <= 2:
                groups[-1].append(row)
            else:
                groups.append([row])
        for group in groups:
            if len(group) < 8:
                continue
            x1 = min(r[0] for r in group)
            y1 = min(r[1] for r in group)
            x2 = max(r[2] for r in group)
            y2 = max(r[1] for r in group) + 1
            boxes.append(("green_button_or_selected_tab", [x1, y1, x2, y2]))
    return boxes


def anchor_boxes(width: int, height: int) -> list[tuple[str, list[int]]]:
    return [
        ("top_left_close_or_back", [0, 75, 85, 155]),
        ("top_right_menu", [620, 75, 715, 155]),
        ("wechat_home_search_icon", [540, 75, 635, 155]),
        ("wechat_search_input", [70, 80, 690, 175]),
        ("wechat_recent_fiscal_ticket_search", [20, 330, 280, 420]),
        ("wechat_search_result_service_card", [20, 270, 705, 595]),
        ("wechat_search_result_ticket_check_button", [170, 485, 430, 575]),
        ("scan_tab_expected_area", [460, 165, 620, 235]),
        ("service_tab_expected_area", [170, 570, 290, 665]),
        ("service_account_bottom_menu_ticket_check", [500, 1475, min(width, 715), min(height, 1630)]),
        ("service_popup_ticket_check_row", [0, 1480, width, min(height, 1630)]),
        ("camera_album_icon", [580, 1460, 700, 1600]),
        ("album_first_image_slot", [178, 150, 365, 340]),
        ("dialog_confirm_button", [150, 860, 570, 970]),
    ]


def build_candidates(
    image_path: Path,
    annotated_path: Path,
    extra_boxes: list[tuple[str, list[int]]] | None = None,
    page_hint: str = "",
) -> list[Candidate]:
    image = Image.open(image_path).convert("RGB")
    width, height = image.size
    is_launcher = looks_like_android_launcher(page_hint)
    raw: list[tuple[str, list[int]]] = []
    if extra_boxes:
        raw.extend(extra_boxes)
    if not is_launcher:
        for box in connected_components_text_boxes(image):
            x1, y1, x2, y2 = box
            if x2 - x1 >= 18 and y2 - y1 >= 12:
                raw.append(("cv_text_or_ui_region", box))
        raw.extend(color_button_boxes(image))
        raw.extend(anchor_boxes(width, height))

    deduped: list[tuple[str, list[int]]] = []
    for label, box in raw:
        x1, y1, x2, y2 = box
        if x2 <= x1 or y2 <= y1:
            continue
        area = (x2 - x1) * (y2 - y1)
        if area > width * height * 0.35:
            continue
        if any(_iou(box, existing) > 0.82 for _label, existing in deduped):
            continue
        deduped.append((label, box))

    priority: list[tuple[str, list[int]]] = []
    regular: list[tuple[str, list[int]]] = []
    for item in deduped:
        label, box = item
        if (
            label.startswith("paddle_launcher_app:微信")
            or label.startswith("paddle_ocr:微信")
            or label.startswith("heuristic_launcher_app:微信")
        ):
            priority.append(item)
        else:
            regular.append(item)

    # Keep high-confidence OCR app targets first; fill the rest top-to-bottom.
    regular = sorted(regular, key=lambda item: (item[1][1], item[1][0]))
    deduped = priority + regular
    deduped = deduped[:60]
    candidates = [Candidate(i + 1, label, box) for i, (label, box) in enumerate(deduped)]

    out = image.copy()
    draw = ImageDraw.Draw(out)
    colors = ["red", "blue", "orange", "purple", "lime", "cyan", "magenta", "yellow"]
    for c in candidates:
        color = colors[(c.id - 1) % len(colors)]
        x1, y1, x2, y2 = c.box
        draw.rectangle(c.box, outline=color, width=4)
        draw.rectangle([x1, max(0, y1 - 28), x1 + 58, y1], fill=color)
        draw.text((x1 + 5, max(0, y1 - 25)), str(c.id), fill="white")
    out.save(annotated_path)
    image.close()
    return candidates


def looks_like_android_launcher(page_hint: str) -> bool:
    if not page_hint:
        return False
    launcher_markers = [
        "日历",
        "钱包",
        "游戏中心",
        "小米视频",
        "米家",
        "文件管理",
        "小米社区",
        "拼多多",
        "淘宝",
    ]
    workflow_markers = ["微信", "财政票据", "票据查验", "扫码查验", "服务通知", "相册", "图库"]
    wechat_internal_markers = ["我的二维码", "添加朋友", "收照片", "收红包", "财政票据", "票据查验", "扫码查验", "服务通知", "相册", "图库"]
    if "微信" in page_hint and not any(marker in page_hint for marker in wechat_internal_markers):
        return True
    if any(marker in page_hint for marker in workflow_markers):
        return False
    return sum(1 for marker in launcher_markers if marker in page_hint) >= 3


def simplified_wechat_launcher_boxes(page_hint: str, width: int, height: int) -> list[tuple[str, list[int]]]:
    if "微信" not in page_hint:
        return []
    if any(marker in page_hint for marker in ["我的二维码", "添加朋友", "收照片", "财政票据", "票据查验", "扫码查验"]):
        return []
    # The test phone launcher is intentionally simplified to a single WeChat
    # icon near the upper-left. PaddleOCR may return only plain text without
    # LOC tokens, so provide a deterministic tappable app cell fallback.
    return [("heuristic_launcher_app:微信", [40, 80, min(width, 180), min(height, 260)])]


def _iou(a: list[int], b: list[int]) -> float:
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    ix1, iy1 = max(ax1, bx1), max(ay1, by1)
    ix2, iy2 = min(ax2, bx2), min(ay2, by2)
    if ix2 <= ix1 or iy2 <= iy1:
        return 0.0
    inter = (ix2 - ix1) * (iy2 - iy1)
    area_a = (ax2 - ax1) * (ay2 - ay1)
    area_b = (bx2 - bx1) * (by2 - by1)
    return inter / float(area_a + area_b - inter)


def maybe_ask_paddle_vl(image_path: Path) -> str:
    """Optional OCR-ish page summary via PaddleOCR-VL-compatible OpenAI endpoint."""
    try:
        b64 = base64.b64encode(image_path.read_bytes()).decode("ascii")
        payload = {
            "model": PADDLE_VL_MODEL,
            "temperature": PADDLE_VL_TEMPERATURE,
            "max_tokens": PADDLE_VL_MAX_TOKENS,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": "Read all visible Chinese UI text from this phone screenshot, including the bottom dock app labels. Include text locations if your format supports it.",
                        },
                        {"type": "image_url", "image_url": {"url": "data:image/png;base64," + b64}},
                    ],
                }
            ],
        }
        resp = post_json(PADDLE_VL_API, payload, timeout=PADDLE_VL_TIMEOUT)
        return resp.get("choices", [{}])[0].get("message", {}).get("content", "")[:8000]
    except Exception as exc:  # noqa: BLE001 - best-effort helper for MVP traces
        return f"paddle_vl_unavailable: {type(exc).__name__}: {exc}"


def paddle_text_boxes(raw: str, width: int, height: int) -> list[tuple[str, list[int]]]:
    """Parse PaddleOCR-VL location-token text into candidate boxes.

    The vLLM PaddleOCR-VL endpoint returns lines like:
    ``微信<|LOC_118|><|LOC_720|>...`` where LOC coordinates are normalized
    to a 0..1000 grid. Convert them to screenshot pixels.
    """
    boxes: list[tuple[str, list[int]]] = []
    pattern = re.compile(r"(.+?)((?:<\|LOC_\d+\|>){8})")
    loc_pattern = re.compile(r"<\|LOC_(\d+)\|>")
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        match = pattern.search(line)
        if not match:
            continue
        text = match.group(1).strip()
        locs = [int(x) for x in loc_pattern.findall(match.group(2))]
        if len(locs) != 8:
            continue
        xs = locs[0::2]
        ys = locs[1::2]
        x1 = round(min(xs) * width / 1000)
        x2 = round(max(xs) * width / 1000)
        y1 = round(min(ys) * height / 1000)
        y2 = round(max(ys) * height / 1000)
        if x2 <= x1 or y2 <= y1:
            continue
        padded = [max(0, x1 - 18), max(0, y1 - 16), min(width, x2 + 18), min(height, y2 + 16)]
        boxes.append((f"paddle_ocr:{text}", padded))
        if text in {"微信", "WeChat"}:
            # Launcher app labels sit below the icon. Add a larger tappable cell
            # centered above the label, so tapping opens the app instead of just
            # hitting the text baseline.
            cx = (x1 + x2) // 2
            cell = [max(0, cx - 75), max(0, y1 - 150), min(width, cx + 75), min(height, y2 + 20)]
            boxes.append(("paddle_launcher_app:微信", cell))
    return boxes


def call_vlm(
    image_path: Path,
    candidates: list[Candidate],
    paddle_hint: str,
    step: int,
    stop_at: str,
    task_phase: str = "verify",
) -> dict[str, Any]:
    b64 = base64.b64encode(image_path.read_bytes()).decode("ascii")
    cand_json = [c.as_json() for c in candidates]
    instruction = {
        "task": (
            "Return from the invoice verification result page to the 财政票据 service account main/chat page."
            if task_phase == "return_to_entry"
            else "Verify the currently staged invoice QR code in WeChat fiscal invoice verification."
        ),
        "task_phase": task_phase,
        "current_step": step,
        "allowed_actions": [
            "tap_candidate",
            "back",
            "wait",
            "swipe_left",
            "swipe_right",
            "home",
            "stop_success",
            "stop_user_action_required",
            "stop_failed",
        ],
        "policy": [
            "Use candidate_id instead of raw coordinates.",
            "Always choose the smallest/specific candidate that directly covers the target, not a broad container.",
            f"Current stop_at mode is {stop_at}.",
            f"Current task_phase is {task_phase}.",
            "If task_phase is return_to_entry, do not start a new verification. Use back/wait/home/navigation until the 财政票据 service account main/chat page is visible, then stop_success.",
            "If task_phase is return_to_entry, success means the screen is the 财政票据 service account main/chat page with its bottom service menu area, not the invoice result page, not the check form, and not the camera or album.",
            "If task_phase is return_to_entry and the screen is invoice detail/result, 票据查验 form, camera scan page, or album picker, choose back.",
            "If stop_at is album and the photo picker/album page is visible, your action MUST be stop_success with candidate_id null. Do not tap any photo or QR candidate.",
            "A Xiaomi account expired credential banner/toast on the Android launcher is unrelated to this task. Ignore it or wait briefly; do not stop_user_action_required for that banner alone.",
            "Canonical entry strategy: the normal task entrance is the 财政票据 service account main/chat page in WeChat. From there open the bottom 票据查验 service menu.",
            "If the screen is not WeChat, not Android launcher, and not a manual security/permission/login blocker, choose home to return to the Android launcher, then open WeChat and navigate back to 财政票据.",
            "If inside WeChat but not already in 财政票据 or an invoice-check child page, navigate by WeChat search to 财政票据 before continuing.",
            "Hard rule for the current test phone: Android Home returns to a simple launcher page with only WeChat. On Android launcher/home screen, if WeChat is not clearly visible, choose home once or wait for banners to disappear before swiping.",
            "On Android launcher/home screen, only tap WeChat when a candidate label_hint explicitly contains paddle_launcher_app:微信 or paddle_ocr:微信. Otherwise prefer home/wait first; use swipe only after Home has failed to reveal WeChat.",
            "Page relationship map:",
            "0. Android launcher/home screen -> find and tap the WeChat/微信 app icon only when the candidate clearly contains the WeChat green two-bubble icon or label 微信. If WeChat is not clearly visible, use swipe_left or swipe_right to change launcher page. Never guess by grid position.",
            "1. WeChat home/chat list -> tap search icon.",
            "2. WeChat search page -> if recent/search suggestion 财政票据 is visible, tap it.",
            "3. WeChat search results for 财政票据 -> tap the 财政票据 service account or its 票据查验 button.",
            "4. 财政票据 service account chat/home -> tap the bottom 票据查验 menu item.",
            "5. Fiscal invoice H5 loading/header page -> wait until the form or scan tab appears.",
            "6. 票据查验 form page -> tap 扫码查验.",
            "7. WeChat camera scan page -> tap album/gallery.",
            "8. Photo picker/album page -> select unique QR image, unless stop_at is album.",
            "9. After QR selection -> wait for result; success page shows 票据详情/票据信息/金额合计/查验次数.",
            "Back behavior: search page back returns to WeChat home; search results back returns to search page; service account/chat back returns to search results or WeChat; H5 detail back returns to check form; camera back returns to check form; album back returns to camera.",
            "Important: after tapping 票据查验 on the service account, the pjcy.mof.gov.cn H5 page may show a bare loading/header page first. In task_phase verify, if pjcy.mof.gov.cn is visible but form controls or 扫码查验 tab are not visible yet, choose wait. Do not tap top-left close/back during loading.",
            "Launcher behavior: this phone is configured so Home should reveal the intended WeChat-only launcher page. Do not tap random app icons unless the icon is likely WeChat/微信.",
            "If on 财政票据 service account home, open 服务/服务号提供的服务 then choose 票据查验.",
            "If on 票据查验 form, choose 扫码查验.",
            "If on WeChat camera scan page, choose album/gallery.",
            "If in photo picker, choose the only QR/image thumbnail.",
            "If a fiscal invoice landing/header page is still loading and there is no clear scan/form control, wait instead of tapping random menu or album anchors.",
            "If task_phase is verify and invoice detail/票据信息/查验次数/金额合计 is visible, stop_success.",
            "If 查验异常/server transient appears, tap 确认 if needed, then retry from scan page unless retries are exhausted.",
            "If phone is locked, WeChat not logged in, permission prompt, or captcha/manual security appears, stop_user_action_required.",
        ],
        "paddle_vl_hint": paddle_hint,
        "candidates": cand_json,
        "response_schema": {
            "action": "tap_candidate|back|wait|swipe_left|swipe_right|home|stop_success|stop_user_action_required|stop_failed",
            "candidate_id": "number or null",
            "screen_state": "short string",
            "result_status": "in_progress|success|server_transient|user_action_required|failed",
            "reason": "short Chinese explanation",
            "confidence": "0..1",
        },
    }
    payload = {
        "model": VLM_MODEL,
        "temperature": VLM_TEMPERATURE,
        "top_p": VLM_TOP_P,
        "top_k": VLM_TOP_K,
        "max_tokens": VLM_MAX_TOKENS,
        "chat_template_kwargs": {"enable_thinking": VLM_ENABLE_THINKING},
        "messages": [
            {
                "role": "system",
                "content": "You are the main controller for an Android phone invoice-check task. /no_think Return exactly one JSON object. No markdown. Do not output raw coordinates.",
            },
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": json.dumps(instruction, ensure_ascii=False) + "\n/no_think",
                    },
                    {"type": "image_url", "image_url": {"url": "data:image/png;base64," + b64}},
                ],
            },
        ],
    }
    resp = post_json(VLM_API, payload, timeout=VLM_TIMEOUT)
    content = resp.get("choices", [{}])[0].get("message", {}).get("content", "")
    return parse_json_object(content)


def should_home_recover(decision: dict[str, Any]) -> bool:
    text = " ".join(
        str(decision.get(key, "")) for key in ("screen_state", "reason", "result_status", "action")
    ).lower()
    manual_blockers = [
        "locked",
        "锁屏",
        "permission",
        "权限",
        "captcha",
        "验证码",
        "security",
        "安全",
        "login",
        "登录",
        "credential",
        "凭证",
    ]
    if any(marker in text for marker in manual_blockers):
        return False
    recover_markers = [
        "not wechat",
        "non-wechat",
        "other app",
        "wrong app",
        "非微信",
        "其他应用",
        "错误应用",
        "文件管理",
        "file manager",
        "variflight",
        "飞常准",
    ]
    return any(marker in text for marker in recover_markers)


def is_ignorable_system_banner(decision: dict[str, Any]) -> bool:
    text = " ".join(str(decision.get(key, "")) for key in ("screen_state", "reason", "result_status"))
    return "小米账号" in text and ("登录凭证" in text or "凭证失效" in text or "重新登录" in text)


def post_json(url: str, payload: dict[str, Any], *, timeout: float) -> dict[str, Any]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def parse_json_object(text: str) -> dict[str, Any]:
    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = re.sub(r"^```(?:json)?\s*", "", stripped)
        stripped = re.sub(r"\s*```$", "", stripped)
    try:
        return json.loads(stripped)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", stripped, re.S)
        if not match:
            raise
        return json.loads(match.group(0))


def discover_qr_tasks(input_dir: Path) -> list[dict[str, Any]]:
    allowed = {".png", ".jpg", ".jpeg", ".webp"}
    tasks: list[dict[str, Any]] = []
    for path in sorted(input_dir.iterdir()):
        if path.is_file() and path.suffix.lower() in allowed:
            tasks.append({"id": safe_task_id(path), "qr_path": str(path.resolve()), "source_name": path.name})
    return tasks


def safe_task_id(path: Path) -> str:
    stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", path.stem).strip("._")
    return stem or f"qr_{abs(hash(path.name))}"


def load_single_summary(summary_path: Path) -> dict[str, Any]:
    if not summary_path.is_file():
        return {}
    return json.loads(summary_path.read_text(encoding="utf-8"))


def extract_result_conclusion(summary: dict[str, Any]) -> dict[str, Any]:
    validation_screenshot = summary.get("validation_screenshot")
    text = ""
    for step in summary.get("steps", []):
        if step.get("screenshot") == validation_screenshot:
            text = str(step.get("paddle_hint", ""))
            break
    if not text:
        for step in reversed(summary.get("steps", [])):
            decision = step.get("decision", {})
            if decision.get("result_status") == "success" and step.get("task_phase") == "verify":
                text = str(step.get("paddle_hint", ""))
                break
    classification = classify_business_result_text(text)
    business_result = summary.get("business_result") or classification.get("business_result")
    fields: dict[str, str | None] = {}
    if business_result == "verified" or (business_result is None and text):
        fields = {
            "票据代码": _extract_after_label(text, "票据代码"),
            "票据号码": _extract_after_label(text, "票据号码"),
            "校验码": _extract_after_label(text, "校验码"),
            "开票日期": _extract_after_label(text, "开票日期"),
            "金额合计": _extract_after_label(text, "金额合计"),
            "查验次数": _extract_after_label(text, "查验次数"),
        }
    return {
        "business_result": business_result,
        "result_message": summary.get("result_message") or classification.get("message"),
        "fields": {key: value for key, value in fields.items() if value},
        "ocr_excerpt": text[:1200],
    }


def _extract_after_label(text: str, label: str) -> str | None:
    pattern = re.compile(re.escape(label) + r"\s*[:：]?\s*([^\n\r|，,。 ]{1,64})")
    match = pattern.search(text)
    return match.group(1).strip() if match else None


def classify_business_result_text(text: str) -> dict[str, str]:
    normalized = _compact_text(text, limit=2000)
    compact = normalized.replace(" ", "")
    if not compact:
        return {}
    daily_markers = [
        "超过当日限制查验次数",
        "超过当日查验次数",
        "当日限制查验次数",
        "请明日进行查验",
    ]
    if any(marker in compact for marker in daily_markers):
        return {
            "business_result": "daily_limit",
            "screen_state": "超过当日查验次数限制提示",
            "message": "超过当日可核验次数，请明日进行查验。",
            "stop_reason": "daily_limit_reached",
        }
    verified_markers = ["票据详情", "票据信息", "金额合计", "查验次数"]
    if sum(1 for marker in verified_markers if marker in compact) >= 2:
        return {
            "business_result": "verified",
            "screen_state": "票据详情/查验结果页",
            "message": "已获取票据详情页。",
            "stop_reason": "invoice_detail_captured",
        }
    invalid_markers = [
        "查无此票",
        "未查询到",
        "票据不存在",
        "二维码无效",
        "票据代码或号码错误",
        "票据信息有误",
        "查验失败",
        "不能查验",
        "无法查验",
    ]
    if any(marker in compact for marker in invalid_markers):
        return {
            "business_result": "invalid_or_not_found",
            "screen_state": "票据查验失败/不可查提示",
            "message": normalized[:220],
            "stop_reason": "invoice_invalid_or_not_found",
        }
    return {}


def write_batch_outputs(out_dir: Path, results: list[dict[str, Any]]) -> Path:
    screenshots_dir = out_dir / "validation_screenshots"
    screenshots_dir.mkdir(parents=True, exist_ok=True)
    for idx, item in enumerate(results, start=1):
        screenshot_path = item.get("validation_screenshot")
        if not screenshot_path:
            continue
        src = Path(screenshot_path)
        if src.is_file():
            dst = screenshots_dir / validation_screenshot_filename(item, idx, src.suffix)
            shutil.copy2(src, dst)
            item["packaged_validation_screenshot"] = str(dst)
            item["validation_screenshot_filename"] = dst.name

    report_json = out_dir / "batch_report.json"
    report_json.write_text(json.dumps({"results": results}, ensure_ascii=False, indent=2), encoding="utf-8")

    report_xlsx = out_dir / "发票核验结果清单.xlsx"
    write_excel_report(report_xlsx, results)

    archive_path = out_dir / "invoice_check_results.zip"
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in [report_xlsx]:
            zf.write(path, path.relative_to(out_dir))
        for path in screenshots_dir.glob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(out_dir))
    return archive_path


def validation_screenshot_filename(item: dict[str, Any], idx: int, suffix: str) -> str:
    source = safe_filename_stem(str(item.get("source_name") or item.get("task_id") or f"invoice_{idx}"))
    return f"{idx:03d}_{source}_核验结果{suffix or '.png'}"


def safe_filename_stem(value: str) -> str:
    stem = Path(value).stem or value
    stem = re.sub(r"[\\/:*?\"<>|]+", "_", stem)
    stem = re.sub(r"\s+", "_", stem).strip("._ ")
    return stem[:80] or "invoice"


def write_excel_report(path: Path, results: list[dict[str, Any]]) -> None:
    wb = Workbook()
    summary_ws = wb.active
    summary_ws.title = "汇总"
    total = len(results)
    success = sum(1 for item in results if item.get("status") == "success")
    skipped = sum(1 for item in results if item.get("status") == "skipped")
    failed = total - success - skipped
    summary_rows = [
        ("校验总张数", total),
        ("已取得核验结果张数", success),
        ("因当日限制跳过张数", skipped),
        ("失败/需人工处理张数", failed),
        ("核验结果截图目录", "validation_screenshots/"),
        ("结果压缩包", "invoice_check_results.zip"),
    ]
    for row in summary_rows:
        summary_ws.append(row)
    summary_ws.column_dimensions["A"].width = 28
    summary_ws.column_dimensions["B"].width = 48
    for cell in summary_ws["A"]:
        cell.font = Font(bold=True)
    for row in summary_ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws = wb.active
    ws = wb.create_sheet("发票核验结果")
    ws.title = "发票核验结果"
    headers = [
        "序号",
        "来源文件",
        "任务状态",
        "核验结论",
        "结果说明",
        "尝试次数",
        "票据代码",
        "票据号码",
        "校验码",
        "开票日期",
        "金额合计",
        "查验次数",
        "错误类型",
        "核验结果截图文件名",
        "核验结果截图路径",
        "OCR摘要",
    ]
    ws.append(headers)
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for idx, item in enumerate(results, start=1):
        fields = item.get("conclusion", {}).get("fields", {})
        conclusion = item.get("conclusion", {})
        packaged_shot = item.get("packaged_validation_screenshot") or item.get("validation_screenshot")
        packaged_name = item.get("validation_screenshot_filename") or (Path(packaged_shot).name if packaged_shot else "")
        row = [
            idx,
            item.get("source_name"),
            item.get("status"),
            conclusion.get("business_result") or item.get("business_result"),
            conclusion.get("result_message") or item.get("result_message"),
            item.get("attempts"),
            fields.get("票据代码"),
            fields.get("票据号码"),
            fields.get("校验码"),
            fields.get("开票日期"),
            fields.get("金额合计"),
            fields.get("查验次数"),
            item.get("error_type"),
            packaged_name,
            packaged_shot,
            item.get("conclusion", {}).get("ocr_excerpt"),
        ]
        ws.append(row)
        row_no = ws.max_row
        if packaged_shot:
            ws.cell(row=row_no, column=15).hyperlink = packaged_shot
            ws.cell(row=row_no, column=15).style = "Hyperlink"
        status_cell = ws.cell(row=row_no, column=3)
        if item.get("status") == "success":
            status_cell.fill = PatternFill("solid", fgColor="E2F0D9")
        else:
            status_cell.fill = PatternFill("solid", fgColor="FCE4D6")

    widths = [8, 32, 14, 20, 42, 10, 18, 18, 18, 16, 16, 16, 28, 36, 58, 80]
    for i, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def build_single_result(args: argparse.Namespace, summary: dict[str, Any]) -> dict[str, Any]:
    source = Path(args.qr_image)
    return {
        "task_id": safe_task_id(source),
        "source_name": source.name,
        "qr_path": str(source.resolve()),
        "status": summary.get("status", "failed"),
        "attempts": 1,
        "attempt_records": [{
            "attempt": 1,
            "pass": "single",
            "summary_path": str(Path(args.output_dir).resolve() / "summary.json"),
            "status": summary.get("status", "failed"),
            "error_type": summary.get("error_type"),
        }],
        "summary_path": str(Path(args.output_dir).resolve() / "summary.json"),
        "validation_screenshot": summary.get("validation_screenshot"),
        "final_screenshot": summary.get("final_screenshot"),
        "stop_reason": summary.get("stop_reason"),
        "error_type": summary.get("error_type"),
        "business_result": summary.get("business_result"),
        "result_message": summary.get("result_message"),
        "conclusion": extract_result_conclusion(summary),
    }


def write_single_outputs(out_dir: Path, result: dict[str, Any]) -> Path:
    archive_path = write_batch_outputs(out_dir, [result])
    summary_path = out_dir / "single_summary.json"
    payload = {
        "status": result.get("status", "failed"),
        "total": 1,
        "success": 1 if result.get("status") == "success" else 0,
        "failed": 0 if result.get("status") == "success" else 1,
        "archive": str(archive_path),
        "report_json": str(out_dir / "batch_report.json"),
        "report_xlsx": str(out_dir / "发票核验结果清单.xlsx"),
        "screenshots_dir": str(out_dir / "validation_screenshots"),
        "result": result,
    }
    summary_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return archive_path


def archive_process_pages(out_dir: Path) -> Path | None:
    process_pages = sorted(out_dir.glob("step_*.png"))
    if not process_pages:
        return None
    archive_path = out_dir / "process_trace.zip"
    with zipfile.ZipFile(archive_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for page in process_pages:
            zf.write(page, page.name)
    for page in process_pages:
        try:
            page.unlink()
        except OSError:
            pass
    return archive_path


def archive_batch_process_pages(out_dir: Path) -> None:
    runs_dir = out_dir / "runs"
    if not runs_dir.is_dir():
        return
    for attempt_dir in runs_dir.glob("*/*/attempt_*"):
        if attempt_dir.is_dir():
            archive_process_pages(attempt_dir)


def emit_progress(message: str, **fields: Any) -> None:
    """Print a human-readable progress line for AgentD Task Output."""
    if not EMIT_PROGRESS_EVENTS:
        return
    suffix = ""
    if fields:
        clean_fields = {
            key: str(value)
            for key, value in fields.items()
            if value is not None and str(value) != ""
        }
        if clean_fields:
            suffix = " | " + " | ".join(f"{key}={value}" for key, value in clean_fields.items())
    print(f"Progress: {message}{suffix}", flush=True)


def _compact_text(value: Any, *, limit: int = 220) -> str:
    text = str(value or "").strip()
    text = re.sub(r"\s+", " ", text)
    if len(text) > limit:
        return text[: limit - 1].rstrip() + "…"
    return text


def emit_agent(kind: str, message: str, **fields: Any) -> None:
    """Print a user-facing agent trace line for long-running Task Output."""
    labels = {
        "observe": "Agent观察",
        "think": "Agent判断",
        "act": "Agent动作",
        "result": "Agent结果",
    }
    label = labels.get(kind, "Agent日志")
    suffix = ""
    clean_fields = {
        key: _compact_text(value, limit=160)
        for key, value in fields.items()
        if value is not None and str(value) != ""
    }
    if clean_fields:
        suffix = "（" + "；".join(f"{key}：{value}" for key, value in clean_fields.items()) + "）"
    print(f"{label}：{message}{suffix}", flush=True)


def summarize_decision(decision: dict[str, Any], candidate: Candidate | None = None) -> str:
    action = str(decision.get("action") or "unknown")
    state = str(decision.get("screen_state") or "")
    reason = str(decision.get("reason") or "")
    target = ""
    if candidate is not None:
        x, y = candidate.center
        target = f" -> candidate {candidate.id} ({candidate.label_hint}) at ({x},{y})"
    summary = f"{action}{target}"
    if state:
        summary += f"; screen={state}"
    if reason:
        summary += f"; reason={reason}"
    return summary


def describe_action(decision: dict[str, Any], candidate: Candidate | None = None) -> str:
    action = str(decision.get("action") or "unknown").strip()
    if action == "tap_candidate" and candidate is not None:
        x, y = candidate.center
        return f"点击候选项 {candidate.id}（{candidate.label_hint}，坐标 {x},{y}）"
    mapping = {
        "wait": "等待页面响应",
        "home": "按 Android Home 键",
        "back": "按 Android Back 键",
        "swipe_left": "向左滑动桌面/页面",
        "swipe_right": "向右滑动桌面/页面",
        "stop_success": "停止并判定查验成功",
        "stop_failed": "停止并判定查验失败",
        "stop_user_action_required": "停止并等待用户人工处理",
    }
    return mapping.get(action, f"执行动作 {action}")


def run_agent(args: argparse.Namespace) -> int:
    global EMIT_PROGRESS_EVENTS
    EMIT_PROGRESS_EVENTS = bool(getattr(args, "emit_progress_events", False))
    out_dir = Path(args.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    trace_path = out_dir / "trace.jsonl"
    summary_path = out_dir / "summary.json"

    emit_progress(
        "invoice verification started",
        output_dir=out_dir,
        max_steps=args.max_steps,
        stop_at=args.stop_at,
    )
    emit_agent(
        "act",
        "开始发票查验。我会先准备手机相册中的二维码，然后逐屏观察微信界面并决定下一步操作。",
        输出目录=out_dir,
        最大步数=args.max_steps,
    )
    width, height = get_size()
    emit_progress("phone screen detected", size=f"{width}x{height}")
    emit_agent("observe", "已连接到手机并读取屏幕尺寸。", 屏幕=f"{width}x{height}")
    summary: dict[str, Any] = {
        "status": "running",
        "stop_at": args.stop_at,
        "screen_size": [width, height],
        "steps": [],
        "output_dir": str(out_dir),
    }

    if args.stage_image:
        emit_progress("staging QR image to phone album", image=Path(args.stage_image).resolve(), remote_dir=REMOTE_DIR)
        emit_agent("act", "把待查验二维码推送到手机相册，供微信扫码页选择。", 文件=Path(args.stage_image).resolve())
        stage_image(Path(args.stage_image).resolve())
        emit_progress("QR image staged", remote=f"{REMOTE_DIR}/{REMOTE_IMAGE}")
        emit_agent("result", "二维码已放入手机相册。", 手机路径=f"{REMOTE_DIR}/{REMOTE_IMAGE}")
        time.sleep(args.after_stage_sleep)

    server_transient_count = 0
    home_recover_count = 0
    task_phase = "verify"
    validation_screenshot: str | None = None
    for step in range(1, args.max_steps + 1):
        raw_path = out_dir / f"step_{step:03d}.png"
        annotated_path = out_dir / f"step_{step:03d}_candidates.png"
        emit_progress(f"step {step}/{args.max_steps}: capturing phone screen", screenshot=raw_path)
        emit_agent("observe", f"第 {step} 步：截取当前手机屏幕。")
        screenshot(raw_path)
        emit_progress(f"step {step}/{args.max_steps}: reading visible UI text")
        emit_agent("observe", f"第 {step} 步：识别屏幕文字和可点击区域。")
        paddle_hint = maybe_ask_paddle_vl(raw_path) if args.use_paddle_hint else ""
        paddle_boxes = paddle_text_boxes(paddle_hint, width, height) if args.use_paddle_hint else []
        if args.use_paddle_hint and not any("微信" in label for label, _box in paddle_boxes):
            paddle_boxes.extend(simplified_wechat_launcher_boxes(paddle_hint, width, height))
        page_is_launcher = looks_like_android_launcher(paddle_hint)
        candidates = build_candidates(raw_path, annotated_path, paddle_boxes, page_hint=paddle_hint)
        emit_progress(
            f"step {step}/{args.max_steps}: candidates built",
            candidate_count=len(candidates),
            annotated=annotated_path,
        )
        hint_summary = _compact_text(paddle_hint, limit=180)
        if hint_summary and hint_summary.count("📶") < 5:
            emit_agent("observe", "屏幕理解结果已生成。", 画面=hint_summary)
        else:
            emit_agent("observe", "屏幕理解结果已生成。")
        emit_agent("observe", "已标注可点击候选区域。", 候选数量=len(candidates), 标注图=annotated_path)
        terminal_result = classify_business_result_text(paddle_hint)
        if terminal_result:
            decision = {
                "action": "stop_success",
                "candidate_id": None,
                "screen_state": terminal_result["screen_state"],
                "result_status": "success",
                "business_result": terminal_result["business_result"],
                "reason": terminal_result["message"],
                "confidence": 1.0,
            }
        else:
            decision = call_vlm(annotated_path, candidates, paddle_hint, step, args.stop_at, task_phase)
        planned_candidate = None
        if str(decision.get("action", "")).strip() == "tap_candidate":
            cid = int(decision.get("candidate_id") or 0)
            planned_candidate = next((c for c in candidates if c.id == cid), None)
        emit_progress(
            f"step {step}/{args.max_steps}: decision",
            detail=summarize_decision(decision, planned_candidate),
        )
        emit_agent(
            "think",
            f"第 {step} 步决策完成。",
            当前画面=decision.get("screen_state"),
            理由=decision.get("reason"),
            计划动作=describe_action(decision, planned_candidate),
            置信度=decision.get("confidence"),
        )

        record = {
            "step": step,
            "task_phase": task_phase,
            "screenshot": str(raw_path),
            "annotated": str(annotated_path),
            "candidate_count": len(candidates),
            "candidates": [c.as_json() for c in candidates],
            "paddle_hint": paddle_hint,
            "decision": decision,
        }
        with trace_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
        summary["steps"].append(record)
        if args.emit_json_events:
            print(json.dumps({"step": step, "decision": decision}, ensure_ascii=False), flush=True)

        if terminal_result:
            summary["status"] = "success"
            summary["business_result"] = terminal_result["business_result"]
            summary["result_message"] = terminal_result["message"]
            summary["stop_reason"] = terminal_result["stop_reason"]
            summary["validation_screenshot"] = str(raw_path)
            summary["final_screenshot"] = str(raw_path)
            emit_agent(
                "result",
                "已识别到业务终态，停止当前票据查验，不再重复点击或重试。",
                结论=terminal_result["business_result"],
                说明=terminal_result["message"],
                结果截图=raw_path,
            )
            break

        action = str(decision.get("action", "")).strip()
        result_status = str(decision.get("result_status", "")).strip()
        screen_state = str(decision.get("screen_state", "")).lower()
        if task_phase == "verify" and args.stop_at == "album" and (
            "photo picker" in screen_state
            or "album" in screen_state
            or "相册" in screen_state
            or "图库" in screen_state
        ):
            emit_progress(f"step {step}/{args.max_steps}: reached album page, stopping as requested")
            emit_agent("result", "已到达相册选择页，按当前 stop_at 设置停止。")
            summary["status"] = "success"
            summary["stop_reason"] = "stop_at_album"
            summary["final_screenshot"] = str(raw_path)
            break
        if action == "stop_success" or result_status == "success":
            if task_phase == "verify" and args.return_to_entry:
                validation_screenshot = str(raw_path)
                summary["validation_screenshot"] = validation_screenshot
                task_phase = "return_to_entry"
                emit_progress(
                    f"step {step}/{args.max_steps}: verification result captured, returning to entry page",
                    validation_screenshot=validation_screenshot,
                )
                emit_agent("result", "已经捕获查验结果页，接下来返回入口页。", 结果截图=validation_screenshot)
                if args.emit_json_events:
                    print(
                        json.dumps(
                            {
                                "step": step,
                                "host_guard": "validation_captured_return_to_entry",
                                "validation_screenshot": validation_screenshot,
                                "fallback": "back",
                            },
                            ensure_ascii=False,
                        ),
                        flush=True,
                    )
                keyevent(4)
                time.sleep(args.after_action_sleep)
                continue
            summary["status"] = "success"
            summary["final_screenshot"] = str(raw_path)
            if task_phase == "return_to_entry":
                summary["stop_reason"] = "returned_to_entry"
                if validation_screenshot:
                    summary["validation_screenshot"] = validation_screenshot
            emit_progress(f"step {step}/{args.max_steps}: task succeeded", final_screenshot=raw_path)
            emit_agent("result", "查验流程成功结束。", 最终截图=raw_path)
            break
        if result_status == "server_transient":
            server_transient_count += 1
            emit_progress(
                f"step {step}/{args.max_steps}: server transient reported",
                count=server_transient_count,
                max_retries=args.max_server_retries,
            )
            emit_agent("think", "识别到服务端临时异常，先等待或重试，不立刻判定任务失败。", 次数=server_transient_count)
            if server_transient_count > args.max_server_retries:
                summary["status"] = "failed"
                summary["error_type"] = "max_server_transient_retries"
                summary["final_screenshot"] = str(raw_path)
                emit_progress("stopping after too many server transient errors", final_screenshot=raw_path)
                emit_agent("result", "连续服务端临时异常超过上限，停止本次查验。", 最终截图=raw_path)
                break
        if action in {"stop_user_action_required", "stop_failed"}:
            if is_ignorable_system_banner(decision):
                emit_progress(f"step {step}/{args.max_steps}: ignorable system banner, waiting")
                emit_agent("think", "当前像是可忽略的系统提示，先等待页面恢复。")
                if args.emit_json_events:
                    print(
                        json.dumps(
                            {
                                "step": step,
                                "host_guard": "ignorable_system_banner_wait",
                                "fallback": "wait",
                            },
                            ensure_ascii=False,
                        ),
                        flush=True,
                    )
                time.sleep(args.wait_seconds)
                continue
            if should_home_recover(decision) and home_recover_count < args.max_home_recovers:
                home_recover_count += 1
                emit_progress(
                    f"step {step}/{args.max_steps}: wrong app/page, pressing Home for recovery",
                    recover_count=home_recover_count,
                )
                emit_agent("act", "当前不在预期微信页面，按 Home 回到桌面后重新定位。", 恢复次数=home_recover_count)
                if args.emit_json_events:
                    print(
                        json.dumps(
                            {
                                "step": step,
                                "host_guard": "non_wechat_home_recover",
                                "home_recover_count": home_recover_count,
                            },
                            ensure_ascii=False,
                        ),
                        flush=True,
                    )
                keyevent(3)
                time.sleep(args.after_action_sleep)
                continue
            summary["status"] = "user_action_required" if action == "stop_user_action_required" else "failed"
            summary["error_type"] = decision.get("screen_state", action)
            summary["final_screenshot"] = str(raw_path)
            emit_progress(
                f"step {step}/{args.max_steps}: stopped",
                status=summary["status"],
                error_type=summary["error_type"],
                final_screenshot=raw_path,
            )
            emit_agent("result", "流程停止，需要用户查看原因或人工介入。", 状态=summary["status"], 原因=summary["error_type"])
            break
        if action == "wait":
            emit_progress(f"step {step}/{args.max_steps}: waiting", seconds=args.wait_seconds)
            emit_agent("act", "等待页面加载或状态变化。", 秒数=args.wait_seconds)
            time.sleep(args.wait_seconds)
            continue
        if action == "home":
            emit_progress(f"step {step}/{args.max_steps}: pressing Android Home")
            emit_agent("act", "按 Android Home 键。")
            keyevent(3)
            time.sleep(args.after_action_sleep)
            continue
        if action == "back":
            emit_progress(f"step {step}/{args.max_steps}: pressing Android Back")
            emit_agent("act", "按 Android Back 键。")
            keyevent(4)
            time.sleep(args.after_action_sleep)
            continue
        if action == "swipe_left":
            if "launcher" in screen_state and ("not visible" in screen_state or "未" in str(decision.get("reason", ""))):
                emit_progress(f"step {step}/{args.max_steps}: launcher missing WeChat, pressing Home before swipe")
                emit_agent("act", "桌面上没有看到微信，先按 Home 复位再继续。")
                if args.emit_json_events:
                    print(
                        json.dumps(
                            {"step": step, "host_guard": "launcher_missing_wechat_home_first", "fallback": "home"},
                            ensure_ascii=False,
                        ),
                        flush=True,
                    )
                keyevent(3)
                time.sleep(args.after_action_sleep)
                continue
            emit_progress(f"step {step}/{args.max_steps}: swiping left")
            emit_agent("act", "向左滑动，尝试寻找微信或目标入口。")
            swipe(int(width * 0.82), int(height * 0.55), int(width * 0.18), int(height * 0.55))
            time.sleep(args.after_action_sleep)
            continue
        if action == "swipe_right":
            if "launcher" in screen_state and ("not visible" in screen_state or "未" in str(decision.get("reason", ""))):
                emit_progress(f"step {step}/{args.max_steps}: launcher missing WeChat, pressing Home before swipe")
                emit_agent("act", "桌面上没有看到微信，先按 Home 复位再继续。")
                if args.emit_json_events:
                    print(
                        json.dumps(
                            {"step": step, "host_guard": "launcher_missing_wechat_home_first", "fallback": "home"},
                            ensure_ascii=False,
                        ),
                        flush=True,
                    )
                keyevent(3)
                time.sleep(args.after_action_sleep)
                continue
            emit_progress(f"step {step}/{args.max_steps}: swiping right")
            emit_agent("act", "向右滑动，尝试寻找微信或目标入口。")
            swipe(int(width * 0.18), int(height * 0.55), int(width * 0.82), int(height * 0.55))
            time.sleep(args.after_action_sleep)
            continue
        if action == "tap_candidate":
            cid = int(decision.get("candidate_id") or 0)
            candidate = next((c for c in candidates if c.id == cid), None)
            if candidate is None:
                if page_is_launcher:
                    emit_progress(
                        f"step {step}/{args.max_steps}: invalid launcher candidate, swiping left instead",
                        bad_candidate_id=cid,
                    )
                    emit_agent("think", "模型给出的桌面候选项无效，改为滑动查找。", 候选编号=cid)
                    if args.emit_json_events:
                        print(
                            json.dumps(
                                {
                                    "step": step,
                                    "host_guard": "launcher_invalid_candidate_blocked",
                                    "bad_candidate_id": cid,
                                    "fallback": "swipe_left",
                                },
                                ensure_ascii=False,
                            ),
                            flush=True,
                        )
                    swipe(int(width * 0.82), int(height * 0.55), int(width * 0.18), int(height * 0.55))
                    time.sleep(args.after_action_sleep)
                    continue
                summary["status"] = "failed"
                summary["error_type"] = "invalid_candidate_id"
                summary["bad_decision"] = decision
                summary["final_screenshot"] = str(raw_path)
                emit_progress(
                    f"step {step}/{args.max_steps}: invalid candidate, stopping",
                    bad_candidate_id=cid,
                    final_screenshot=raw_path,
                )
                emit_agent("result", "模型给出的候选编号不存在，停止以避免误点。", 候选编号=cid)
                break
            if page_is_launcher and "微信" not in candidate.label_hint:
                emit_progress(
                    f"step {step}/{args.max_steps}: blocked non-WeChat launcher tap, swiping left instead",
                    candidate=candidate.label_hint,
                )
                emit_agent("think", "候选项不是微信入口，阻止误点并改为继续滑动。", 候选=candidate.label_hint)
                if args.emit_json_events:
                    print(
                        json.dumps(
                            {
                                "step": step,
                                "host_guard": "launcher_non_wechat_tap_blocked",
                                "blocked_candidate": candidate.as_json(),
                                "fallback": "swipe_left",
                            },
                            ensure_ascii=False,
                        ),
                        flush=True,
                    )
                swipe(int(width * 0.82), int(height * 0.55), int(width * 0.18), int(height * 0.55))
                time.sleep(args.after_action_sleep)
                continue
            x, y = candidate.center
            emit_progress(
                f"step {step}/{args.max_steps}: tapping candidate",
                candidate_id=candidate.id,
                label=candidate.label_hint,
                center=f"{x},{y}",
            )
            emit_agent("act", f"点击候选项 {candidate.id}。", 名称=candidate.label_hint, 坐标=f"{x},{y}")
            tap(x, y)
            time.sleep(args.after_action_sleep)
            continue

        summary["status"] = "failed"
        summary["error_type"] = "unknown_action"
        summary["bad_decision"] = decision
        summary["final_screenshot"] = str(raw_path)
        emit_progress(
            f"step {step}/{args.max_steps}: unknown action, stopping",
            action=action,
            final_screenshot=raw_path,
        )
        emit_agent("result", "收到未知动作，停止以避免误操作。", 动作=action)
        break
    else:
        summary["status"] = "failed"
        summary["error_type"] = "max_steps_exceeded"
        emit_progress("maximum step count reached", max_steps=args.max_steps)
        emit_agent("result", "达到最大步数，本次短跑停止。", 最大步数=args.max_steps)

    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    emit_progress("invoice verification finished", status=summary["status"], summary=summary_path)
    emit_agent("result", "本次发票查验任务已写入结果文件。", 状态=summary["status"], 摘要=summary_path)
    if args.emit_json_events:
        print(json.dumps({"summary": str(summary_path), "status": summary["status"]}, ensure_ascii=False))
    return 0 if summary["status"] == "success" else 2


def run_batch(args: argparse.Namespace) -> int:
    out_dir = Path(args.output_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    input_dir = Path(args.batch_input_dir).resolve()
    tasks = discover_qr_tasks(input_dir)
    if not tasks:
        raise SystemExit(f"No QR image files found in {input_dir}")

    results_by_id: dict[str, dict[str, Any]] = {}
    failed: list[dict[str, Any]] = []
    daily_limit_reached = False
    for index, task in enumerate(tasks):
        result = run_batch_task(args, task, out_dir, pass_name="first")
        results_by_id[task["id"]] = result
        if result.get("business_result") == "daily_limit":
            daily_limit_reached = True
            for remaining in tasks[index + 1:]:
                results_by_id[remaining["id"]] = skipped_due_to_daily_limit_result(remaining)
            break
        if result["status"] != "success":
            failed.append(task)
            if args.failed_task_sleep > 0:
                time.sleep(args.failed_task_sleep)

    if failed and not daily_limit_reached and not args.no_failed_replay:
        replay_failed: list[dict[str, Any]] = []
        if args.failed_replay_sleep > 0:
            time.sleep(args.failed_replay_sleep)
        for task in failed:
            result = run_batch_task(args, task, out_dir, pass_name="replay")
            results_by_id[task["id"]] = result
            if result["status"] != "success":
                replay_failed.append(task)
                if args.failed_task_sleep > 0:
                    time.sleep(args.failed_task_sleep)
        failed = replay_failed

    ordered_results = [results_by_id[task["id"]] for task in tasks]
    archive_path = write_batch_outputs(out_dir, ordered_results)
    archive_batch_process_pages(out_dir)
    batch_summary = {
        "status": "success" if not failed else "partial_failed",
        "total": len(tasks),
        "success": sum(1 for item in ordered_results if item["status"] == "success"),
        "skipped": sum(1 for item in ordered_results if item["status"] == "skipped"),
        "technical_failed": sum(1 for item in ordered_results if item["status"] not in {"success", "skipped"}),
        "failed": sum(1 for item in ordered_results if item["status"] not in {"success", "skipped"}),
        "archive": str(archive_path),
        "report_json": str(out_dir / "batch_report.json"),
        "report_xlsx": str(out_dir / "发票核验结果清单.xlsx"),
        "screenshots_dir": str(out_dir / "validation_screenshots"),
        "results": ordered_results,
    }
    (out_dir / "batch_summary.json").write_text(
        json.dumps(batch_summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(batch_summary, ensure_ascii=False))
    return 0 if batch_summary["status"] == "success" else 2


def skipped_due_to_daily_limit_result(task: dict[str, Any]) -> dict[str, Any]:
    return {
        "task_id": task["id"],
        "source_name": task["source_name"],
        "qr_path": task["qr_path"],
        "status": "skipped",
        "attempts": 0,
        "attempt_records": [],
        "summary_path": None,
        "validation_screenshot": None,
        "final_screenshot": None,
        "error_type": None,
        "business_result": "skipped_daily_limit",
        "result_message": "前序票据已触发当日查验次数限制，本张未继续查验。",
        "conclusion": {
            "business_result": "skipped_daily_limit",
            "result_message": "前序票据已触发当日查验次数限制，本张未继续查验。",
            "fields": {},
            "ocr_excerpt": "",
        },
    }


def run_batch_task(args: argparse.Namespace, task: dict[str, Any], batch_out_dir: Path, *, pass_name: str) -> dict[str, Any]:
    attempts_dir = batch_out_dir / "runs" / task["id"] / pass_name
    attempts_dir.mkdir(parents=True, exist_ok=True)
    last_summary: dict[str, Any] = {}
    attempt_records: list[dict[str, Any]] = []
    for attempt in range(1, args.batch_attempts + 1):
        attempt_dir = attempts_dir / f"attempt_{attempt:02d}"
        single_args = argparse.Namespace(**vars(args))
        single_args.output_dir = str(attempt_dir)
        single_args.stage_image = task["qr_path"]
        single_args.stop_at = "result"
        single_args.return_to_entry = True
        code = run_agent(single_args)
        last_summary = load_single_summary(attempt_dir / "summary.json")
        if not args.keep_phone_album:
            cleanup_phone_album()
        attempt_records.append({
            "attempt": attempt,
            "pass": pass_name,
            "exit_code": code,
            "summary_path": str(attempt_dir / "summary.json"),
            "status": last_summary.get("status", "failed"),
            "error_type": last_summary.get("error_type"),
        })
        if code == 0 and last_summary.get("status") == "success":
            conclusion = extract_result_conclusion(last_summary)
            return {
                "task_id": task["id"],
                "source_name": task["source_name"],
                "qr_path": task["qr_path"],
                "status": "success",
                "attempts": len(attempt_records),
                "attempt_records": attempt_records,
                "summary_path": str(attempt_dir / "summary.json"),
                "validation_screenshot": last_summary.get("validation_screenshot"),
                "final_screenshot": last_summary.get("final_screenshot"),
                "stop_reason": last_summary.get("stop_reason"),
                "business_result": last_summary.get("business_result"),
                "result_message": last_summary.get("result_message"),
                "conclusion": conclusion,
            }
        if args.attempt_sleep > 0:
            time.sleep(args.attempt_sleep)

    return {
        "task_id": task["id"],
        "source_name": task["source_name"],
        "qr_path": task["qr_path"],
        "status": last_summary.get("status", "failed") if last_summary else "failed",
        "attempts": len(attempt_records),
        "attempt_records": attempt_records,
        "summary_path": attempt_records[-1]["summary_path"] if attempt_records else None,
        "validation_screenshot": last_summary.get("validation_screenshot") if last_summary else None,
        "final_screenshot": last_summary.get("final_screenshot") if last_summary else None,
        "error_type": last_summary.get("error_type") if last_summary else "no_summary",
        "business_result": last_summary.get("business_result") if last_summary else None,
        "result_message": last_summary.get("result_message") if last_summary else None,
        "conclusion": extract_result_conclusion(last_summary) if last_summary else {},
    }


def cmd_health(_: argparse.Namespace) -> int:
    payload = {
        "success": True,
        "service": "invoice-check-cli",
        "config": {
            "adb_path": str(ADB),
            "vlm_api": VLM_API,
            "vlm_model": VLM_MODEL,
            "paddle_vl_api": PADDLE_VL_API,
            "paddle_vl_model": PADDLE_VL_MODEL,
            "remote_dir": REMOTE_DIR,
            "remote_image": REMOTE_IMAGE,
            "vlm_sampling": {
                "temperature": VLM_TEMPERATURE,
                "top_p": VLM_TOP_P,
                "top_k": VLM_TOP_K,
                "max_tokens": VLM_MAX_TOKENS,
                "timeout": VLM_TIMEOUT,
                "enable_thinking": VLM_ENABLE_THINKING,
            },
            "paddle_vl_sampling": {
                "temperature": PADDLE_VL_TEMPERATURE,
                "max_tokens": PADDLE_VL_MAX_TOKENS,
                "timeout": PADDLE_VL_TIMEOUT,
            },
        },
    }
    print(json.dumps(payload, ensure_ascii=False))
    return 0


def add_runtime_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--max-steps", type=int, default=int(CONFIG.get("max_steps", 80)))
    parser.add_argument("--max-server-retries", type=int, default=int(CONFIG.get("max_server_retries", 2)))
    parser.add_argument("--after-stage-sleep", type=float, default=float(CONFIG.get("after_stage_sleep", 1.5)))
    parser.add_argument("--after-action-sleep", type=float, default=float(CONFIG.get("after_action_sleep", 2.5)))
    parser.add_argument("--wait-seconds", type=float, default=float(CONFIG.get("wait_seconds", 3.0)))
    parser.add_argument("--max-home-recovers", type=int, default=int(CONFIG.get("max_home_recovers", 3)))
    parser.add_argument("--use-paddle-hint", action=argparse.BooleanOptionalAction, default=bool(CONFIG.get("use_paddle_hint", True)))
    parser.add_argument("--cleanup-phone-album", action=argparse.BooleanOptionalAction, default=bool(CONFIG.get("cleanup_phone_album", True)))
    parser.add_argument("--keep-phone-album", action="store_true")
    parser.add_argument(
        "--emit-json-events",
        action="store_true",
        help="Also print per-step JSON debug events to stdout. Human-readable Task Output is the default.",
    )
    parser.add_argument(
        "--emit-progress-events",
        action="store_true",
        help="Also print low-level Progress: debug lines to stdout. Hidden by default.",
    )


def cmd_verify_one(args: argparse.Namespace) -> int:
    args.stage_image = str(Path(args.qr_image).resolve())
    args.stop_at = "result"
    args.return_to_entry = True
    try:
        code = run_agent(args)
        out_dir = Path(args.output_dir).resolve()
        summary = load_single_summary(out_dir / "summary.json")
        result = build_single_result(args, summary)
        archive_path = write_single_outputs(out_dir, result)
        process_archive = archive_process_pages(out_dir)
        if process_archive:
            summary["process_trace_archive"] = str(process_archive)
            (out_dir / "summary.json").write_text(
                json.dumps(summary, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        emit_agent(
            "result",
            "最终交付物已生成：Excel 清单、核验结果截图目录和结果压缩包。",
            Excel=out_dir / "发票核验结果清单.xlsx",
            截图目录=out_dir / "validation_screenshots",
            压缩包=archive_path,
        )
        return code
    finally:
        if args.cleanup_phone_album and not args.keep_phone_album:
            cleanup_phone_album()


def cmd_verify_batch(args: argparse.Namespace) -> int:
    args.batch_input_dir = args.input_dir
    args.stop_at = "result"
    args.return_to_entry = True
    return run_batch(args)


def main() -> int:
    parser = argparse.ArgumentParser(description="Invoice check phone-use CLI service for AgentD.")
    sub = parser.add_subparsers(dest="command", required=True)

    health = sub.add_parser("health", help="Print service configuration.")
    health.set_defaults(func=cmd_health)

    one = sub.add_parser("verify-one", help="Verify one QR image through the Android phone.")
    one.add_argument("--qr-image", required=True)
    one.add_argument("--output-dir", required=True)
    add_runtime_args(one)
    one.set_defaults(func=cmd_verify_one)

    batch = sub.add_parser("verify-batch", help="Verify a directory of QR images.")
    batch.add_argument("--input-dir", required=True, help="Directory containing extracted QR images.")
    batch.add_argument("--output-dir", required=True)
    add_runtime_args(batch)
    batch.add_argument("--batch-attempts", "--attempts", type=int, default=int(CONFIG.get("batch_attempts", 3)))
    batch.add_argument("--attempt-sleep", type=float, default=float(CONFIG.get("attempt_sleep", 2.0)))
    batch.add_argument("--failed-task-sleep", type=float, default=float(CONFIG.get("failed_task_sleep", 10.0)))
    batch.add_argument("--failed-replay-sleep", type=float, default=float(CONFIG.get("failed_replay_sleep", 30.0)))
    batch.add_argument("--no-failed-replay", action="store_true")
    batch.set_defaults(func=cmd_verify_batch)

    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
