#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Перцептивный хеш (pHash) вложения для «глубокого анализа» — поиск ВИЗУАЛЬНО
одинаковых фотоотчётов/сканов, даже если у файлов разные имена/размеры или один
и тот же кадр вставлен в разные PDF.

Использование:  python3 phash.py <path> <mime>
Вывод (stdout, JSON):  {"ok": true, "phashes": ["<hex16>", ...]}
Для изображения — один хеш; для PDF — по хешу на страницу (до MAX_PAGES).
Любая ошибка -> {"ok": false, "error": "..."} и код возврата 0 (Node решает сам).
"""
import sys
import json

MAX_PAGES = 5          # сколько страниц PDF хешируем (фотоотчёты обычно короткие)
RENDER_DPI = 100       # достаточно для устойчивого pHash, но не тяжело


def phash_image(pil_img):
    import imagehash
    return str(imagehash.phash(pil_img))  # 16 hex-символов (64 бита)


def hashes_for_pdf(path):
    import fitz  # PyMuPDF
    from PIL import Image
    out = []
    doc = fitz.open(path)
    try:
        for i, page in enumerate(doc):
            if i >= MAX_PAGES:
                break
            pix = page.get_pixmap(dpi=RENDER_DPI)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            out.append(phash_image(img))
    finally:
        doc.close()
    return out


def hashes_for_image(path):
    from PIL import Image
    with Image.open(path) as img:
        return [phash_image(img.convert("RGB"))]


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"ok": False, "error": "no path"}))
        return
    path = sys.argv[1]
    mime = (sys.argv[2] if len(sys.argv) > 2 else "").lower()
    try:
        if mime == "application/pdf" or path.lower().endswith(".pdf"):
            phashes = hashes_for_pdf(path)
        else:
            phashes = hashes_for_image(path)
        # Убираем пустые/дубли, сохраняя порядок
        seen = set()
        uniq = []
        for h in phashes:
            if h and h not in seen:
                seen.add(h)
                uniq.append(h)
        print(json.dumps({"ok": True, "phashes": uniq}))
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": str(e)}))


if __name__ == "__main__":
    main()
