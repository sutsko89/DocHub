"""Служебный скрипт для скриншотов интерфейса (в сборку не попадает).

Примеры:
    python dev_screenshot.py preview/main.png
    set UI_THEME=dark && python dev_screenshot.py preview/dark.png
    set UI_PAGE=additional && set UI_MODE=ip && python dev_screenshot.py preview/ip.png
    set UI_CAPTURE=screen && python dev_screenshot.py preview/screen.png

Переменные окружения:
    UI_THEME    light | dark
    UI_PAGE     contract | additional
    UI_MODE     ooo | ip
    UI_FOCUS    подпись поля, на которое поставить фокус
    UI_CAPTURE  window (PrintWindow, по умолчанию) | screen (снимок экрана целиком)
"""

from __future__ import annotations

import ctypes
import os
import sys
import tkinter as tk
from ctypes import wintypes

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
os.chdir(BASE)

import main  # noqa: E402
import ui_macos as ui  # noqa: E402


def capture_window(window) -> "object":
    """Снимок окна через PrintWindow — работает, даже если окно перекрыто."""
    from PIL import Image

    class BITMAPINFOHEADER(ctypes.Structure):
        _fields_ = [
            ("biSize", wintypes.DWORD),
            ("biWidth", wintypes.LONG),
            ("biHeight", wintypes.LONG),
            ("biPlanes", wintypes.WORD),
            ("biBitCount", wintypes.WORD),
            ("biCompression", wintypes.DWORD),
            ("biSizeImage", wintypes.DWORD),
            ("biXPelsPerMeter", wintypes.LONG),
            ("biYPelsPerMeter", wintypes.LONG),
            ("biClrUsed", wintypes.DWORD),
            ("biClrImportant", wintypes.DWORD),
        ]

    user32 = ctypes.windll.user32
    gdi32 = ctypes.windll.gdi32
    user32.GetWindowDC.restype = wintypes.HDC
    user32.GetWindowDC.argtypes = [wintypes.HWND]
    user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
    user32.GetParent.argtypes = [wintypes.HWND]
    user32.GetParent.restype = wintypes.HWND
    user32.PrintWindow.argtypes = [wintypes.HWND, wintypes.HDC, wintypes.UINT]
    user32.PrintWindow.restype = wintypes.BOOL
    user32.ReleaseDC.argtypes = [wintypes.HWND, wintypes.HDC]
    gdi32.CreateCompatibleDC.restype = wintypes.HDC
    gdi32.CreateCompatibleDC.argtypes = [wintypes.HDC]
    gdi32.CreateCompatibleBitmap.restype = wintypes.HBITMAP
    gdi32.CreateCompatibleBitmap.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
    gdi32.SelectObject.argtypes = [wintypes.HDC, wintypes.HGDIOBJ]
    gdi32.SelectObject.restype = wintypes.HGDIOBJ
    gdi32.GetDIBits.argtypes = [
        wintypes.HDC, wintypes.HBITMAP, wintypes.UINT, wintypes.UINT,
        ctypes.c_void_p, ctypes.c_void_p, wintypes.UINT,
    ]
    gdi32.DeleteObject.argtypes = [wintypes.HGDIOBJ]
    gdi32.DeleteDC.argtypes = [wintypes.HDC]

    hwnd = user32.GetParent(window.winfo_id()) or window.winfo_id()
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    width, height = rect.right - rect.left, rect.bottom - rect.top

    window_dc = user32.GetWindowDC(hwnd)
    memory_dc = gdi32.CreateCompatibleDC(window_dc)
    bitmap = gdi32.CreateCompatibleBitmap(window_dc, width, height)
    gdi32.SelectObject(memory_dc, bitmap)
    user32.PrintWindow(hwnd, memory_dc, 2)

    header = BITMAPINFOHEADER()
    header.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    header.biWidth = width
    header.biHeight = -height
    header.biPlanes = 1
    header.biBitCount = 32
    header.biCompression = 0

    buffer = ctypes.create_string_buffer(width * height * 4)
    gdi32.GetDIBits(memory_dc, bitmap, 0, height, buffer, ctypes.byref(header), 0)
    image = Image.frombuffer("RGB", (width, height), buffer, "raw", "BGRX", 0, 1).copy()

    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(memory_dc)
    user32.ReleaseDC(hwnd, window_dc)
    return image


def compose(app) -> "object":
    from PIL import Image

    root_image = capture_window(app.root)
    origin = (app.root.winfo_rootx(), app.root.winfo_rooty())
    for child in app.root.winfo_children():
        if not isinstance(child, tk.Toplevel) or not child.winfo_ismapped():
            continue
        layer = capture_window(child)
        offset = (child.winfo_rootx() - origin[0], child.winfo_rooty() - origin[1])
        root_image.paste(layer, offset)
    return root_image


def main_() -> int:
    output = sys.argv[1] if len(sys.argv) > 1 else os.path.join(BASE, "preview", "shot.png")
    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)

    app = main.ContractApp()
    theme = os.environ.get("UI_THEME")
    if theme:
        ui.set_appearance_mode(theme)

    page = os.environ.get("UI_PAGE")
    if page in app.pages:
        app._select_page(page)

    if os.environ.get("UI_CAPTURE") == "screen":
        app.root.attributes("-topmost", True)
        app.root.lift()

    sheet = os.environ.get("UI_SHEET")
    if sheet == "alert":
        app.root.after(300, lambda: ui.alert(
            app.root, "Ошибки в данных",
            "Проверьте следующие поля:\n\n• ОГРН должен содержать 13 цифр\n"
            "• Расчётный счёт должен содержать 20 цифр",
            "error",
        ))

    def grab() -> None:
        try:
            app.root.update_idletasks()
            app.root.update()
            if os.environ.get("UI_CAPTURE") == "screen":
                from PIL import ImageGrab

                x, y = app.root.winfo_rootx(), app.root.winfo_rooty()
                width, height = app.root.winfo_width(), app.root.winfo_height()
                image = ImageGrab.grab(
                    bbox=(x - 14, y - 14, x + width + 14, y + height + 14), all_screens=True,
                )
            else:
                image = compose(app)
            image.save(output)
            print("saved:", output)
        finally:
            app.root.after(50, app.root.destroy)

    app.root.after(2600, grab)
    app.root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main_())
