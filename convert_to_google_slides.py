#!/usr/bin/env python3
"""
convert_to_google_slides.py
============================
Parses slides.html and creates a Google Slides presentation via the Slides API.

Prerequisites
-------------
pip install google-auth google-auth-oauthlib google-api-python-client beautifulsoup4

Authentication
--------------
1. Go to https://console.cloud.google.com and create a project.
2. Enable the Google Slides API and Google Drive API.
3. Create OAuth 2.0 credentials (Desktop app) and download as `credentials.json`.
4. Place `credentials.json` next to this script.

Usage
-----
python convert_to_google_slides.py               # creates a new presentation
python convert_to_google_slides.py --open        # also opens it in the browser
python convert_to_google_slides.py --id PRES_ID  # update an existing presentation
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import webbrowser
from pathlib import Path
from typing import Any

from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ── Configuration ──────────────────────────────────────────────────────────────

SLIDES_HTML = Path(__file__).parent / "slides.html"
CREDENTIALS_FILE = Path(__file__).parent / "credentials.json"
TOKEN_FILE = Path(__file__).parent / "token.json"
SCOPES = [
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive.file",
]

# Slide dimensions (19:9 widescreen, in EMU — 1 inch = 914400 EMU)
SLIDE_WIDTH = 9_144_000   # 10 inches
SLIDE_HEIGHT = 5_143_500  # 5.625 inches

# Palette
COLORS = {
    "white":       {"red": 0.973, "green": 0.980, "blue": 0.992},
    "slate_900":   {"red": 0.059, "green": 0.090, "blue": 0.165},
    "slate_800":   {"red": 0.118, "green": 0.161, "blue": 0.231},
    "slate_300":   {"red": 0.796, "green": 0.835, "blue": 0.886},
    "blue_600":    {"red": 0.149, "green": 0.502, "blue": 0.941},
    "blue_800":    {"red": 0.118, "green": 0.251, "blue": 0.690},
    "violet_700":  {"red": 0.486, "green": 0.231, "blue": 0.929},
    "teal_700":    {"red": 0.059, "green": 0.463, "blue": 0.431},
    "pink_600":    {"red": 0.859, "green": 0.149, "blue": 0.467},
    "sky_300":     {"red": 0.482, "green": 0.827, "blue": 0.980},
}


# ── Auth ───────────────────────────────────────────────────────────────────────

def get_credentials() -> Credentials:
    creds: Credentials | None = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                sys.exit(
                    f"[ERROR] {CREDENTIALS_FILE} not found.\n"
                    "Download OAuth credentials from Google Cloud Console and save as credentials.json."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
    return creds


# ── HTML Parsing ───────────────────────────────────────────────────────────────

class SlideData:
    def __init__(self, index: int, title: str, theme: str, elements: list[dict]):
        self.index = index
        self.title = title
        self.theme = theme
        self.elements = elements  # list of {type, text, level}


def parse_slides(html_path: Path) -> list[SlideData]:
    soup = BeautifulSoup(html_path.read_text(encoding="utf-8"), "html.parser")
    slides: list[SlideData] = []
    for div in soup.find_all("div", class_="slide"):
        classes = div.get("class", [])
        theme = next((c for c in classes if c != "slide"), "content")
        index = int(div.get("data-slide", 0))
        title = div.get("data-title", f"Slide {index}")
        elements: list[dict] = []
        for tag in div.find_all(["div", "h1", "h2", "h3", "p", "li"]):
            cls = tag.get("class", [])
            text = tag.get_text(" ", strip=True)
            if not text:
                continue
            if "label" in cls:
                elements.append({"type": "label", "text": text})
            elif tag.name == "h1":
                elements.append({"type": "h1", "text": text})
            elif tag.name == "h2":
                elements.append({"type": "h2", "text": text})
            elif tag.name == "h3":
                elements.append({"type": "h3", "text": text})
            elif tag.name == "li":
                elements.append({"type": "bullet", "text": text})
            elif tag.name == "p" and "stat" not in cls and "stat-label" not in cls:
                elements.append({"type": "body", "text": text})
            elif "stat" in cls:
                elements.append({"type": "stat", "text": text})
            elif "stat-label" in cls:
                elements.append({"type": "stat_label", "text": text})
        slides.append(SlideData(index, title, theme, elements))
    slides.sort(key=lambda s: s.index)
    return slides


# ── Request Builders ───────────────────────────────────────────────────────────

def emu(pt: float) -> int:
    """Points to EMU (1 pt = 12700 EMU)."""
    return int(pt * 12700)


def pt(pixels: float) -> float:
    """Approximate px → pt (96 dpi → 72 dpi)."""
    return pixels * 0.75


def rgb(color_key: str) -> dict:
    return {"rgbColor": COLORS[color_key]}


def solid_fill(color_key: str) -> dict:
    return {"solidFill": {"color": rgb(color_key)}}


def text_style(
    bold: bool = False,
    font_size_pt: float = 14,
    color_key: str = "white",
    font: str = "Google Sans",
) -> dict:
    return {
        "bold": bold,
        "fontFamily": font,
        "fontSize": {"magnitude": font_size_pt, "unit": "PT"},
        "foregroundColor": rgb(color_key),
    }


def bg_gradient(theme: str) -> dict:
    """Return a background fill request dict for a given theme."""
    gradients = {
        "cover":     ("blue_800", "violet_700"),
        "highlight": ("teal_700", "blue_800"),
        "closing":   ("violet_700", "pink_600"),
        "section":   ("slate_800", "slate_800"),
        "content":   ("slate_800", "slate_800"),
    }
    c1, c2 = gradients.get(theme, ("slate_800", "slate_800"))
    return {
        "linearGradient": {
            "colorStops": [
                {"color": rgb(c1), "position": 0.0},
                {"color": rgb(c2), "position": 1.0},
            ],
            "angle": 135,
        }
    }


def add_text_box(
    object_id: str,
    slide_id: str,
    text: str,
    x: int, y: int, w: int, h: int,
    font_size: float = 14,
    bold: bool = False,
    color: str = "white",
    font: str = "Google Sans",
    align: str = "LEFT",
) -> list[dict]:
    """Return a list of API requests to create a styled text box."""
    reqs = [
        {
            "createShape": {
                "objectId": object_id,
                "shapeType": "TEXT_BOX",
                "elementProperties": {
                    "pageObjectId": slide_id,
                    "size": {"width": {"magnitude": w, "unit": "EMU"},
                             "height": {"magnitude": h, "unit": "EMU"}},
                    "transform": {"scaleX": 1, "scaleY": 1,
                                  "translateX": x, "translateY": y,
                                  "unit": "EMU"},
                },
            }
        },
        {
            "insertText": {
                "objectId": object_id,
                "text": text,
            }
        },
        {
            "updateTextStyle": {
                "objectId": object_id,
                "style": text_style(bold=bold, font_size_pt=font_size, color_key=color, font=font),
                "fields": "bold,fontFamily,fontSize,foregroundColor",
            }
        },
        {
            "updateParagraphStyle": {
                "objectId": object_id,
                "style": {"alignment": align},
                "fields": "alignment",
            }
        },
    ]
    return reqs


# ── Slide Layout Builder ───────────────────────────────────────────────────────

def build_slide_requests(slide: SlideData, slide_id: str, page_object_id: str) -> list[dict]:
    """Generate all API requests to populate a single slide."""
    reqs: list[dict] = []
    counter = [0]

    def uid(prefix: str = "obj") -> str:
        counter[0] += 1
        return f"{slide_id}_{prefix}_{counter[0]}"

    W = SLIDE_WIDTH
    H = SLIDE_HEIGHT
    PAD_X = emu(56)
    PAD_Y = emu(48)
    INNER_W = W - 2 * PAD_X

    # ── Background ──
    reqs.append({
        "updatePageProperties": {
            "objectId": page_object_id,
            "pageProperties": {
                "pageBackgroundFill": {
                    "stretchedPictureFill": None,
                    "solidFill": {"color": rgb("slate_800")},
                }
            },
            "fields": "pageBackgroundFill.solidFill",
        }
    })

    # ── Logo (top-right) — skip on cover and closing ──
    if slide.theme not in ("cover", "closing"):
        reqs += add_text_box(
            uid("logo"), page_object_id,
            "Hook",
            x=W - emu(120), y=emu(20),
            w=emu(100), h=emu(32),
            font_size=16, bold=True, color="blue_600",
        )

    # ── Slide number (bottom-right) ──
    reqs += add_text_box(
        uid("num"), page_object_id,
        str(slide.index),
        x=W - emu(56), y=H - emu(32),
        w=emu(40), h=emu(24),
        font_size=10, color="slate_300", align="RIGHT",
    )

    # ── Lay out text elements in a simple vertical stack ──
    y_cursor = PAD_Y
    bullets: list[str] = []

    def flush_bullets():
        nonlocal y_cursor
        if not bullets:
            return
        text = "\n".join(f"→  {b}" for b in bullets)
        reqs += add_text_box(
            uid("bullets"), page_object_id,
            text,
            x=PAD_X, y=y_cursor,
            w=INNER_W, h=emu(14 * len(bullets) + 8),
            font_size=13, color="slate_300",
        )
        y_cursor += emu(14 * len(bullets) + 16)
        bullets.clear()

    for el in slide.elements:
        t = el["type"]
        txt = el["text"]

        if t == "label":
            flush_bullets()
            reqs += add_text_box(
                uid("label"), page_object_id, txt.upper(),
                x=PAD_X, y=y_cursor, w=INNER_W, h=emu(20),
                font_size=10, color="slate_300",
            )
            y_cursor += emu(24)

        elif t == "h1":
            flush_bullets()
            fs = 44 if slide.theme in ("cover", "closing") else 36
            reqs += add_text_box(
                uid("h1"), page_object_id, txt,
                x=PAD_X, y=y_cursor, w=INNER_W, h=emu(fs * 1.4),
                font_size=fs, bold=True, color="white",
            )
            y_cursor += emu(fs * 1.4 + 10)

        elif t == "h2":
            flush_bullets()
            reqs += add_text_box(
                uid("h2"), page_object_id, txt,
                x=PAD_X, y=y_cursor, w=INNER_W, h=emu(38),
                font_size=26, bold=True, color="white",
            )
            y_cursor += emu(48)

        elif t == "h3":
            flush_bullets()
            reqs += add_text_box(
                uid("h3"), page_object_id, txt,
                x=PAD_X, y=y_cursor, w=INNER_W, h=emu(24),
                font_size=15, bold=True, color="sky_300",
            )
            y_cursor += emu(28)

        elif t == "body":
            flush_bullets()
            reqs += add_text_box(
                uid("body"), page_object_id, txt,
                x=PAD_X, y=y_cursor, w=INNER_W, h=emu(48),
                font_size=13, color="slate_300",
            )
            y_cursor += emu(52)

        elif t == "bullet":
            bullets.append(txt)

        elif t == "stat":
            flush_bullets()
            reqs += add_text_box(
                uid("stat"), page_object_id, txt,
                x=PAD_X, y=y_cursor, w=emu(200), h=emu(48),
                font_size=32, bold=True, color="sky_300",
            )
            y_cursor += emu(48)

        elif t == "stat_label":
            flush_bullets()
            reqs += add_text_box(
                uid("slabel"), page_object_id, txt,
                x=PAD_X, y=y_cursor, w=INNER_W, h=emu(20),
                font_size=11, color="slate_300",
            )
            y_cursor += emu(24)

    flush_bullets()
    return reqs


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(description="Convert slides.html → Google Slides")
    parser.add_argument("--id", dest="pres_id", default=None, help="Existing presentation ID to update")
    parser.add_argument("--open", dest="open_browser", action="store_true", help="Open in browser when done")
    args = parser.parse_args()

    print("[1/4] Authenticating …")
    creds = get_credentials()
    slides_service = build("slides", "v1", credentials=creds)

    print("[2/4] Parsing slides.html …")
    slides = parse_slides(SLIDES_HTML)
    print(f"      Found {len(slides)} slides.")

    # ── Create or fetch presentation ──
    if args.pres_id:
        pres_id = args.pres_id
        print(f"[3/4] Updating existing presentation {pres_id} …")
        # Delete all existing slides first
        pres = slides_service.presentations().get(presentationId=pres_id).execute()
        existing_pages = pres.get("slides", [])
        if existing_pages:
            delete_reqs = [{"deleteObject": {"objectId": p["objectId"]}} for p in existing_pages]
            slides_service.presentations().batchUpdate(
                presentationId=pres_id,
                body={"requests": delete_reqs},
            ).execute()
    else:
        print("[3/4] Creating new presentation …")
        pres = slides_service.presentations().create(
            body={
                "title": "Hook — Business Case",
                "slides": [],
                "pageSize": {
                    "width": {"magnitude": SLIDE_WIDTH, "unit": "EMU"},
                    "height": {"magnitude": SLIDE_HEIGHT, "unit": "EMU"},
                },
            }
        ).execute()
        pres_id = pres["presentationId"]
        # Delete the default blank slide
        default_page = pres["slides"][0]["objectId"] if pres.get("slides") else None
        if default_page:
            slides_service.presentations().batchUpdate(
                presentationId=pres_id,
                body={"requests": [{"deleteObject": {"objectId": default_page}}]},
            ).execute()

    print(f"[4/4] Building {len(slides)} slides …")
    for slide in slides:
        sid = f"slide_{slide.index:03d}"
        page_oid = f"page_{slide.index:03d}"

        # 1. Add blank slide
        slides_service.presentations().batchUpdate(
            presentationId=pres_id,
            body={
                "requests": [
                    {
                        "createSlide": {
                            "objectId": page_oid,
                            "insertionIndex": slide.index - 1,
                            "slideLayoutReference": {"predefinedLayout": "BLANK"},
                        }
                    }
                ]
            },
        ).execute()

        # 2. Populate slide content
        content_reqs = build_slide_requests(slide, sid, page_oid)
        if content_reqs:
            slides_service.presentations().batchUpdate(
                presentationId=pres_id,
                body={"requests": content_reqs},
            ).execute()

        print(f"      ✓ Slide {slide.index}: {slide.title}")

    url = f"https://docs.google.com/presentation/d/{pres_id}/edit"
    print(f"\n✅  Done! Presentation ID: {pres_id}")
    print(f"   Open: {url}")

    if args.open_browser:
        webbrowser.open(url)


if __name__ == "__main__":
    main()
