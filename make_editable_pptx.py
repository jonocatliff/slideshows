"""
Editable PPTX: faithfully reproduces the design-system colours, typography,
and layouts using python-pptx shapes and text — fully editable in Google Slides.

Slide order: Intro, 1-Antigravity, 2-Plugin, 3-Folder, 4-Prompt
"""
import pathlib, io, textwrap
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import asyncio
from playwright.async_api import async_playwright

# ── paths ──────────────────────────────────────────────────────────────────
BASE = pathlib.Path(__file__).parent
OUT  = BASE / "editable.pptx"

# ── slide dimensions: 16:9 ─────────────────────────────────────────────────
W = 9144000   # EMU  (≈ 25.4 cm / 10 in)
H = 5143500   # EMU  (≈ 14.29 cm)

# ── design-system tokens ───────────────────────────────────────────────────
C_PRIMARY      = RGBColor(0xcc, 0x78, 0x5c)   # coral
C_INK          = RGBColor(0x14, 0x14, 0x13)
C_BODY         = RGBColor(0x3d, 0x3d, 0x3a)
C_MUTED        = RGBColor(0x6c, 0x6a, 0x64)
C_MUTED_SOFT   = RGBColor(0x8e, 0x8b, 0x82)
C_CANVAS       = RGBColor(0xfa, 0xf9, 0xf5)
C_SURFACE_CARD = RGBColor(0xef, 0xe9, 0xde)
C_DARK         = RGBColor(0x18, 0x17, 0x15)
C_DARK_EL      = RGBColor(0x25, 0x23, 0x20)
C_DARK_SOFT    = RGBColor(0x1f, 0x1e, 0x1b)
C_ON_DARK      = RGBColor(0xfa, 0xf9, 0xf5)
C_ON_DARK_SOFT = RGBColor(0xa0, 0x9d, 0x96)
C_TEAL         = RGBColor(0x5d, 0xb8, 0xa6)
C_AMBER        = RGBColor(0xe8, 0xa5, 0x5a)
C_SUCCESS      = RGBColor(0x5d, 0xb8, 0x72)
C_WHITE        = RGBColor(0xff, 0xff, 0xff)


# ── helpers ────────────────────────────────────────────────────────────────
def cm(v): return int(v * 360000)     # centimetres → EMU
def pct(v, total): return int(v / 100 * total)


def add_rect(slide, x, y, w, h, fill, alpha=None):
    """Add a filled rectangle, return the shape."""
    shape = slide.shapes.add_shape(1, x, y, w, h)   # MSO_SHAPE_TYPE.RECTANGLE = 1
    shape.line.fill.background()
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    return shape


def add_textbox(slide, x, y, w, h, text, font_pt, bold=False, italic=False,
                color=None, align=PP_ALIGN.LEFT, wrap=True, line_spacing=None):
    """Add a textbox; return the frame."""
    tf = slide.shapes.add_textbox(x, y, w, h).text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    if line_spacing:
        p.line_spacing = line_spacing
    return tf


def add_label(slide, x, y, w, h, text, size=10, color=None, bold=False, align=PP_ALIGN.LEFT):
    return add_textbox(slide, x, y, w, h, text, size, bold=bold, color=color, align=align)


def slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def rounded_rect(slide, x, y, w, h, fill, radius_emu=None):
    """Add a rectangle with slight rounding (workaround: pptx uses freeform or adjustVal)."""
    shape = add_rect(slide, x, y, w, h, fill)
    if radius_emu:
        # Set corner rounding via XML
        sp = shape._element
        spPr = sp.find(qn('p:spPr'))
        prstGeom = spPr.find(qn('a:prstGeom'))
        if prstGeom is not None:
            avLst = prstGeom.find(qn('a:avLst'))
            if avLst is None:
                from lxml import etree
                avLst = etree.SubElement(prstGeom, qn('a:avLst'))
            # adj is in 1/100 000ths of shape width; clamp to 50 000 (50%)
            adj_val = min(50000, int(radius_emu / w * 100000))
            from lxml import etree
            gd = etree.SubElement(avLst, qn('a:gd'))
            gd.set('name', 'adj')
            gd.set('fmla', f'val {adj_val}')
            prstGeom.set('prst', 'roundRect')
    return shape


def add_bullet_row(slide, x, y, dot_color, title, body, title_col, body_col, dot_r=cm(0.12)):
    """Dot + bold title + body on one line."""
    add_rect(slide, x, y + cm(0.12), dot_r, dot_r, dot_color)
    add_textbox(slide, x + cm(0.32), y - cm(0.05), cm(7.5), cm(0.55),
                title, 10, bold=True, color=title_col)
    add_textbox(slide, x + cm(0.32) + cm(2.2), y - cm(0.05), cm(5.5), cm(0.55),
                body, 10, color=body_col)


def add_feature_card(slide, x, y, w, h, title, body, title_col=C_INK, body_col=C_BODY):
    rounded_rect(slide, x, y, w, h, C_SURFACE_CARD, radius_emu=cm(0.3))
    add_textbox(slide, x + cm(0.5), y + cm(0.4), w - cm(1.0), cm(0.55),
                title, 13, bold=True, color=title_col)
    add_textbox(slide, x + cm(0.5), y + cm(1.0), w - cm(1.0), h - cm(1.2),
                body, 11, color=body_col)


def add_dark_card(slide, x, y, w, h, title, body, title_col=C_ON_DARK, body_col=C_ON_DARK_SOFT):
    rounded_rect(slide, x, y, w, h, C_DARK_EL, radius_emu=cm(0.3))
    add_textbox(slide, x + cm(0.5), y + cm(0.35), w - cm(1.0), cm(0.55),
                title, 13, bold=True, color=title_col)
    add_textbox(slide, x + cm(0.5), y + cm(0.95), w - cm(1.0), h - cm(1.1),
                body, 10.5, color=body_col)


def eyebrow(slide, x, y, text, color=C_PRIMARY):
    add_textbox(slide, x, y, cm(12), cm(0.4), text.upper(), 9,
                bold=True, color=color)


def display_h1(slide, x, y, w, h, text, color=C_INK):
    tf = add_textbox(slide, x, y, w, h, text, 36, color=color)
    return tf


def display_h1_dark(slide, x, y, w, h, text):
    return display_h1(slide, x, y, w, h, text, color=C_ON_DARK)


def lead_text(slide, x, y, w, h, text, color=C_BODY):
    add_textbox(slide, x, y, w, h, text, 14, color=color)


# ── screenshot helper (for the orbit illustration on slide 1) ──────────────
async def capture_orbit(html_file, out_path):
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page = await browser.new_page(viewport={"width": 960, "height": 720})
        url = f"file://{(BASE / html_file).resolve()}"
        await page.goto(url, wait_until="networkidle")
        await page.wait_for_timeout(600)
        # Crop just the hero-right illustration
        elem = await page.query_selector('.hero-right')
        if elem:
            box = await elem.bounding_box()
            img = await page.screenshot(
                type="png",
                clip={"x": box["x"], "y": box["y"],
                      "width": box["width"], "height": box["height"]}
            )
        else:
            img = await page.screenshot(type="png")
        await browser.close()
    out_path.write_bytes(img)
    return img


# ── nav bar (thin dark strip) ──────────────────────────────────────────────
def add_nav(slide, slide_label, dark=False, prev_url=None, next_url=None):
    bar_h = cm(0.9)
    bg = C_DARK if dark else C_CANVAS
    add_rect(slide, 0, 0, W, bar_h, bg)

    # brand text
    lbl_col = C_ON_DARK if dark else C_INK
    add_textbox(slide, cm(0.7), cm(0.15), cm(4), cm(0.6),
                "✦ Antigravity", 10, bold=True, color=lbl_col)

    # counter
    counter_col = C_ON_DARK_SOFT if dark else C_MUTED
    add_textbox(slide, W - cm(3.5), cm(0.15), cm(3), cm(0.6),
                slide_label, 9, color=counter_col, align=PP_ALIGN.RIGHT)


# ═══════════════════════════════════════════════════════════════════════════
# SLIDES
# ═══════════════════════════════════════════════════════════════════════════

def build_intro(prs):
    """Index / intro slide — dark background, 2×2 grid of cards."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, C_DARK)
    add_nav(slide, "Intro", dark=True)

    nav_h = cm(0.9)
    body_top = nav_h + cm(0.6)
    body_h   = H - nav_h - cm(0.6)

    # Eyebrow
    add_textbox(slide, cm(1.5), body_top, cm(12), cm(0.5),
                "ANTIGRAVITY · CLAUDE CODE", 9, bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)

    # H1
    add_textbox(slide, cm(1.5), body_top + cm(0.6), W - cm(3), cm(2.2),
                "Building with Claude Code", 40, color=C_ON_DARK, align=PP_ALIGN.CENTER)

    # Subtitle
    add_textbox(slide, pct(20, W), body_top + cm(2.9), pct(60, W), cm(0.8),
                "From zero to your first prompt — in four slides.", 14,
                color=C_ON_DARK_SOFT, align=PP_ALIGN.CENTER)

    # 2×2 card grid
    card_w = cm(8.5)
    card_h = cm(2.6)
    gap    = cm(0.35)
    grid_w = card_w * 2 + gap
    gx     = (W - grid_w) // 2
    gy     = body_top + cm(4.1)

    cards = [
        ("01", "What is Antigravity?",      "An agentic dev environment built around Claude."),
        ("02", "Claude Code as a Plugin",   "Not a separate app — a plugin, wired in."),
        ("03", "Adding a Project Folder",   "Drop in a folder. Claude Code indexes it."),
        ("04", "Sending a Prompt",          "Type what you want. Claude Code acts."),
    ]

    positions = [
        (gx,            gy),
        (gx + card_w + gap, gy),
        (gx,            gy + card_h + gap),
        (gx + card_w + gap, gy + card_h + gap),
    ]

    for (num, title, body), (cx, cy) in zip(cards, positions):
        rounded_rect(slide, cx, cy, card_w, card_h, C_DARK_EL, radius_emu=cm(0.3))
        add_textbox(slide, cx + cm(0.45), cy + cm(0.25), cm(2), cm(0.4),
                    num, 9, bold=True, color=C_PRIMARY)
        add_textbox(slide, cx + cm(0.45), cy + cm(0.65), card_w - cm(0.9), cm(0.55),
                    title, 14, bold=True, color=C_ON_DARK)
        add_textbox(slide, cx + cm(0.45), cy + cm(1.25), card_w - cm(0.9), cm(1.0),
                    body, 11, color=C_ON_DARK_SOFT)

    # Footnote
    add_textbox(slide, 0, H - cm(0.6), W, cm(0.5),
                "Use ← → arrow keys to navigate", 9,
                color=C_MUTED, align=PP_ALIGN.CENTER)


def build_slide1(prs, orbit_img_bytes=None):
    """Slide 1 — What is Antigravity? (cream canvas)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, C_CANVAS)
    add_nav(slide, "1 / 4", dark=False)

    nav_h = cm(0.9)
    lx = cm(1.2)
    ly = nav_h + cm(1.0)

    eyebrow(slide, lx, ly, "01 — Antigravity")
    display_h1(slide, lx, ly + cm(0.55), cm(11), cm(2.6), "Meet\nAntigravity.", C_INK)
    lead_text(slide, lx, ly + cm(3.3), cm(10.5), cm(1.2),
              "An agentic dev environment built around Claude —\nAI and code in the same place.", C_BODY)

    # Orbit illustration card (screenshot) or placeholder
    card_x = cm(13.2)
    card_y = nav_h + cm(0.8)
    card_w = cm(11.6)
    card_h = cm(7.8)
    rounded_rect(slide, card_x, card_y, card_w, card_h, C_DARK, radius_emu=cm(0.4))

    if orbit_img_bytes:
        slide.shapes.add_picture(
            io.BytesIO(orbit_img_bytes),
            card_x, card_y, card_w, card_h,
        )
    else:
        add_textbox(slide, card_x + cm(0.5), card_y + card_h // 2 - cm(0.3),
                    card_w - cm(1), cm(0.6),
                    "[ Antigravity orbit illustration ]", 11,
                    color=C_ON_DARK_SOFT, align=PP_ALIGN.CENTER)

    # Label overlays inside card
    for lbl, lx2, ly2, col in [
        ("Plugins",  card_x + card_w - cm(1.8), card_y + card_h - cm(0.7), C_PRIMARY),
        ("Agents",   card_x + cm(0.4),           card_y + cm(0.4),           C_TEAL),
        ("Projects", card_x + cm(0.4),            card_y + card_h - cm(0.7), C_AMBER),
    ]:
        add_textbox(slide, lx2, ly2, cm(1.8), cm(0.4), lbl, 9, bold=True, color=col)

    # Three feature cards at the bottom
    fc_y  = nav_h + cm(9.0)
    fc_w  = cm(8.0)
    fc_h  = cm(2.5)
    fc_gap = cm(0.35)
    fc_x0 = cm(1.2)

    feats = [
        ("Plugin-first",   "Every capability is a plugin. Add, remove, compose."),
        ("Project-aware",  "Drop in a folder. Your whole codebase becomes live context."),
        ("One-prompt flows", "One prompt. Agents, tools, and code edits — handled."),
    ]
    for i, (t, b) in enumerate(feats):
        add_feature_card(slide, fc_x0 + i * (fc_w + fc_gap), fc_y, fc_w, fc_h, t, b)


def build_slide2(prs):
    """Slide 2 — Claude Code as a Plugin (dark)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, C_DARK)
    add_nav(slide, "2 / 4", dark=True)

    nav_h = cm(0.9)
    lx = cm(1.2)
    ly = nav_h + cm(1.0)

    eyebrow(slide, lx, ly, "02 — Plugin", color=C_PRIMARY)
    display_h1_dark(slide, lx, ly + cm(0.55), cm(10.5), cm(2.4),
                    "Claude Code lives\ninside Antigravity.")
    lead_text(slide, lx, ly + cm(3.15), cm(10.5), cm(0.8),
              "Not a separate app. A plugin — toggled on, wired into your project.", C_ON_DARK_SOFT)

    # Bullet points
    bullets = [
        (C_PRIMARY, "Instant context.", "Reads your codebase the moment it's enabled."),
        (C_TEAL,    "Side-by-side.",    "Terminal panel in the same window — no switching."),
        (C_AMBER,   "Composable.",      "Mix with other plugins in one flow."),
    ]
    by = ly + cm(4.15)
    for dot_col, title, body in bullets:
        rounded_rect(slide, lx, by + cm(0.18), cm(0.18), cm(0.18), dot_col)
        add_textbox(slide, lx + cm(0.36), by, cm(2.4), cm(0.45), title, 11, bold=True, color=C_ON_DARK)
        add_textbox(slide, lx + cm(2.8),  by, cm(7.6), cm(0.45), body,  11, color=C_ON_DARK_SOFT)
        by += cm(0.7)

    # Plugin panel mockup (right side)
    px = cm(13.0)
    py = nav_h + cm(0.8)
    pw = cm(12.0)
    ph = cm(10.5)
    rounded_rect(slide, px, py, pw, ph, C_DARK_EL, radius_emu=cm(0.3))

    # Titlebar
    add_rect(slide, px, py, pw, cm(0.65), C_DARK_SOFT)
    add_textbox(slide, px + cm(1.0), py + cm(0.1), pw - cm(1.2), cm(0.45),
                "Antigravity — Plugins", 10, color=C_ON_DARK_SOFT)
    for i, col in enumerate([RGBColor(0xc6,0x45,0x45), RGBColor(0xe8,0xa5,0x5a), RGBColor(0x5d,0xb8,0x72)]):
        cx2 = px + cm(0.25) + i * cm(0.3)
        add_rect(slide, cx2, py + cm(0.22), cm(0.2), cm(0.2), col)

    # Plugin rows
    plugins = [
        ("✦ Claude Code",       "AI pair programmer & agent", True),
        ("⌥ Git & Branches",    "Inline diff, blame, history", True),
        ("✓ Test Runner",        "Jest, Vitest, pytest",        True),
        ("◇ Deploy",             "Vercel, Railway, Fly.io",    False),
        ("≡ Linter & Formatter", "ESLint, Prettier, Ruff",     False),
    ]
    row_h = cm(1.35)
    ry = py + cm(0.8)
    for icon_name, desc, enabled in plugins:
        row_bg = RGBColor(0x20, 0x1e, 0x1a) if enabled else C_DARK_SOFT
        rounded_rect(slide, px + cm(0.35), ry, pw - cm(0.7), cm(1.2), row_bg, radius_emu=cm(0.2))
        if enabled and "Claude" in icon_name:
            border_shape = slide.shapes.add_shape(1, px + cm(0.35), ry, pw - cm(0.7), cm(1.2))
            border_shape.fill.background()
            border_shape.line.color.rgb = C_PRIMARY
            border_shape.line.width = Emu(19050)
        add_textbox(slide, px + cm(1.0), ry + cm(0.12), pw - cm(2.5), cm(0.45),
                    icon_name, 11, bold=True,
                    color=C_ON_DARK if enabled else C_ON_DARK_SOFT)
        add_textbox(slide, px + cm(1.0), ry + cm(0.6), pw - cm(2.5), cm(0.4),
                    desc, 10, color=C_ON_DARK_SOFT)
        # Toggle
        tog_col = C_PRIMARY if enabled else C_DARK_EL
        rounded_rect(slide, px + pw - cm(1.0), ry + cm(0.45), cm(0.55), cm(0.3), tog_col, radius_emu=cm(0.15))
        ry += row_h

    # Hint bar
    hb_y = py + ph + cm(0.25)
    rounded_rect(slide, px, hb_y, pw, cm(0.8), RGBColor(0x22, 0x1a, 0x14), radius_emu=cm(0.25))
    add_textbox(slide, px + cm(0.5), hb_y + cm(0.15), pw - cm(1.0), cm(0.5),
                "Claude Code is on by default. Everything else is optional.", 10,
                color=C_ON_DARK_SOFT)


def build_slide3(prs):
    """Slide 3 — Adding a Project Folder (cream)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, C_CANVAS)
    add_nav(slide, "3 / 4", dark=False)

    nav_h = cm(0.9)
    lx = cm(1.2)
    ly = nav_h + cm(0.9)

    eyebrow(slide, lx, ly, "03 — Setup")
    display_h1(slide, lx, ly + cm(0.55), cm(11), cm(1.6), "Add your project folder.", C_INK)
    lead_text(slide, lx, ly + cm(2.25), cm(11), cm(0.7),
              "Three steps. Drop in a folder, Claude Code indexes it, you're ready.", C_BODY)

    # Step list
    steps = [
        ("done",     "1", "Open Antigravity",        "Land on the Projects screen."),
        ("active",   "2", "Drop in your folder",     "Drag a folder into the drop zone, or click Open folder."),
        ("upcoming", "3", "Claude Code indexes it",  "Files scanned. Prompt bar unlocks. You're ready."),
    ]
    sx = lx
    sy = ly + cm(3.15)
    step_h = cm(2.2)
    for state, num, title, body in steps:
        dot_fill = C_PRIMARY if state in ("done", "active") else C_CANVAS
        dot_col  = C_WHITE if state == "done" else (C_PRIMARY if state == "active" else C_MUTED)
        add_rect(slide, sx, sy, cm(0.65), cm(0.65), dot_fill)
        add_textbox(slide, sx + cm(0.05), sy + cm(0.1), cm(0.55), cm(0.45),
                    "✓" if state == "done" else num, 9, bold=True, color=dot_col, align=PP_ALIGN.CENTER)
        title_col = C_INK if state != "upcoming" else C_MUTED
        body_col  = C_BODY if state != "upcoming" else C_MUTED_SOFT
        add_textbox(slide, sx + cm(1.0), sy + cm(0.05), cm(9.2), cm(0.45), title, 12, bold=True, color=title_col)
        add_textbox(slide, sx + cm(1.0), sy + cm(0.55), cm(9.2), cm(0.7),  body,  11, color=body_col)
        if state != "upcoming":
            add_rect(slide, sx + cm(0.28), sy + cm(0.65), cm(0.07), step_h - cm(0.65), RGBColor(0xe6,0xdf,0xd8))
        sy += step_h

    # File explorer mockup (right side)
    mx = cm(13.0)
    my = nav_h + cm(0.7)
    mw = cm(12.2)
    mh = cm(11.0)
    rounded_rect(slide, mx, my, mw, mh, C_DARK, radius_emu=cm(0.4))

    # Titlebar
    add_rect(slide, mx, my, mw, cm(0.65), C_DARK_EL)
    add_textbox(slide, mx + cm(1.0), my + cm(0.12), mw - cm(1.2), cm(0.42),
                "Antigravity — Open Project", 10, color=C_ON_DARK_SOFT)
    for i, col in enumerate([RGBColor(0xc6,0x45,0x45), RGBColor(0xe8,0xa5,0x5a), RGBColor(0x5d,0xb8,0x72)]):
        add_rect(slide, mx + cm(0.25) + i * cm(0.3), my + cm(0.22), cm(0.2), cm(0.2), col)

    # Left sidebar (file tree)
    sb_w = cm(3.6)
    add_rect(slide, mx, my + cm(0.65), sb_w, mh - cm(0.65), C_DARK_SOFT)
    tree = [
        (0, "📁 my-project", C_AMBER),
        (1, "  📂 src",      C_ON_DARK_SOFT),
        (2, "    index.ts",  C_TEAL),
        (2, "    app.ts",    C_ON_DARK_SOFT),
        (1, "  📂 tests",    C_ON_DARK_SOFT),
        (1, "  package.json",C_ON_DARK_SOFT),
    ]
    ty = my + cm(0.85)
    for _, label, col in tree:
        add_textbox(slide, mx + cm(0.2), ty, sb_w - cm(0.3), cm(0.42), label, 9, color=col)
        ty += cm(0.48)

    # Drop zone (right portion of mockup)
    dz_x = mx + sb_w + cm(0.2)
    dz_y = my + cm(0.9)
    dz_w = mw - sb_w - cm(0.4)
    dz_h = cm(4.5)

    # Dashed border using a shape with no fill
    dz_shape = slide.shapes.add_shape(1, dz_x, dz_y, dz_w, dz_h)
    dz_shape.fill.background()
    dz_shape.line.color.rgb = C_PRIMARY
    dz_shape.line.width = Emu(19050)
    dz_shape.line.dash_style = 4   # DASH

    add_textbox(slide, dz_x, dz_y + cm(1.3), dz_w, cm(0.5),
                "Drop a folder here", 12, bold=True, color=C_ON_DARK_SOFT, align=PP_ALIGN.CENTER)
    add_textbox(slide, dz_x, dz_y + cm(1.9), dz_w, cm(0.4),
                "or click to browse", 10, color=C_ON_DARK_SOFT, align=PP_ALIGN.CENTER)

    # Open folder button
    btn_x = dz_x + (dz_w - cm(2.8)) // 2
    btn_y = dz_y + cm(2.5)
    rounded_rect(slide, btn_x, btn_y, cm(2.8), cm(0.55), RGBColor(0x2e, 0x1f, 0x16), radius_emu=cm(0.15))
    add_textbox(slide, btn_x, btn_y + cm(0.08), cm(2.8), cm(0.4),
                "Open folder…", 10, bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)

    # Recent label
    rec_y = dz_y + dz_h + cm(0.4)
    add_textbox(slide, dz_x, rec_y, dz_w, cm(0.4), "RECENT", 8, bold=True, color=C_ON_DARK_SOFT)
    for proj, path in [("my-project", "~/Documents/my-project"), ("api-service", "~/Work/api-service")]:
        rec_y += cm(0.5)
        rounded_rect(slide, dz_x, rec_y, dz_w, cm(0.72), C_DARK_EL, radius_emu=cm(0.2))
        add_textbox(slide, dz_x + cm(0.3), rec_y + cm(0.07), dz_w - cm(1.8), cm(0.3),
                    proj, 10, bold=True, color=C_ON_DARK)
        add_textbox(slide, dz_x + cm(0.3), rec_y + cm(0.38), dz_w - cm(1.8), cm(0.28),
                    path, 9, color=C_ON_DARK_SOFT)
        rounded_rect(slide, dz_x + dz_w - cm(1.2), rec_y + cm(0.2), cm(1.0), cm(0.3),
                     RGBColor(0x2e, 0x1f, 0x16), radius_emu=cm(0.1))
        add_textbox(slide, dz_x + dz_w - cm(1.2), rec_y + cm(0.2), cm(1.0), cm(0.3),
                    "Open", 8, bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)


def build_slide4(prs):
    """Slide 4 — Sending a Prompt (dark)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_bg(slide, C_DARK)
    add_nav(slide, "4 / 4", dark=True)

    nav_h = cm(0.9)
    lx = cm(1.2)
    ly = nav_h + cm(0.9)

    eyebrow(slide, lx, ly, "04 — Prompt", color=C_PRIMARY)
    display_h1_dark(slide, lx, ly + cm(0.55), cm(10.5), cm(1.6), "Apples an bananas.")
    lead_text(slide, lx, ly + cm(2.25), cm(10.5), cm(0.7),
              "Type what you want. Claude Code reads your codebase and acts.", C_ON_DARK_SOFT)

    # Anatomy rows
    anatomy = [
        ("verb",    "Start with an action.",  "Add, Fix, Refactor, Write tests for."),
        ("scope",   "Name the thing.",         "A file, a function, a feature."),
        ("context", "Add constraints.",         '"without breaking tests", "using the pattern in auth.ts".'),
    ]
    ay = ly + cm(3.2)
    for badge, title, body in anatomy:
        rounded_rect(slide, lx, ay, cm(1.4), cm(0.42), RGBColor(0x2e, 0x1f, 0x16), radius_emu=cm(0.12))
        add_textbox(slide, lx, ay + cm(0.05), cm(1.4), cm(0.35),
                    badge, 9, bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)
        add_textbox(slide, lx + cm(1.6), ay, cm(3.2), cm(0.42), title, 10, bold=True, color=C_ON_DARK)
        add_textbox(slide, lx + cm(4.9), ay, cm(5.8), cm(0.42), body,  10, color=C_ON_DARK_SOFT)
        ay += cm(0.75)

    # Right side: prompt input + terminal
    rx = cm(13.0)
    ry = nav_h + cm(0.8)
    rw = cm(12.0)

    # Prompt input box
    pi_h = cm(1.8)
    rounded_rect(slide, rx, ry, rw, pi_h, C_DARK_EL, radius_emu=cm(0.3))
    # Coral border
    pi_border = slide.shapes.add_shape(1, rx, ry, rw, pi_h)
    pi_border.fill.background()
    pi_border.line.color.rgb = C_PRIMARY
    pi_border.line.width = Emu(19050)

    add_textbox(slide, rx + cm(0.4), ry + cm(0.18), rw - cm(0.8), cm(0.8),
                'Add a rate-limit middleware to src/app.ts, using the pattern in src/auth.ts, without breaking existing tests',
                11, color=C_ON_DARK)
    # Bottom strip
    add_rect(slide, rx, ry + pi_h - cm(0.55), rw, cm(0.55), C_DARK_SOFT)
    add_textbox(slide, rx + cm(0.3), ry + pi_h - cm(0.5), cm(3), cm(0.4),
                "⌘ K  Commands     @  Files", 9, color=C_ON_DARK_SOFT)
    # Send button
    rounded_rect(slide, rx + rw - cm(1.6), ry + pi_h - cm(0.5), cm(1.3), cm(0.4), C_PRIMARY, radius_emu=cm(0.12))
    add_textbox(slide, rx + rw - cm(1.6), ry + pi_h - cm(0.48), cm(1.3), cm(0.38),
                "Send →", 9, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)

    # Terminal
    tm_y = ry + pi_h + cm(0.28)
    tm_h = cm(5.5)
    rounded_rect(slide, rx, tm_y, rw, tm_h, C_DARK_SOFT, radius_emu=cm(0.3))
    # Titlebar
    add_rect(slide, rx, tm_y, rw, cm(0.6), C_DARK_EL)
    add_textbox(slide, rx + cm(1.0), tm_y + cm(0.1), rw - cm(1.2), cm(0.42),
                "Claude Code — Terminal", 10, color=C_ON_DARK_SOFT)
    for i, col in enumerate([RGBColor(0xc6,0x45,0x45), RGBColor(0xe8,0xa5,0x5a), RGBColor(0x5d,0xb8,0x72)]):
        add_rect(slide, rx + cm(0.25) + i * cm(0.3), tm_y + cm(0.2), cm(0.18), cm(0.18), col)

    terminal_lines = [
        ("> claude \"Add rate-limit middleware…\"", C_ON_DARK),
        ("", None),
        ("Reading src/app.ts, src/auth.ts, tests/ (14 files)", C_ON_DARK_SOFT),
        ("", None),
        ("1.  Install express-rate-limit", C_ON_DARK_SOFT),
        ("2.  Create src/middleware/rateLimiter.ts", C_ON_DARK_SOFT),
        ("3.  Mount after authMiddleware in app.ts", C_ON_DARK_SOFT),
        ("", None),
        ("✓  package.json updated", C_SUCCESS),
        ("✓  rateLimiter.ts created", C_SUCCESS),
        ("✓  app.ts patched", C_SUCCESS),
        ("✓  All 14 tests passing", C_SUCCESS),
        ("", None),
        ("Done. ▌", C_PRIMARY),
    ]
    tl_y = tm_y + cm(0.72)
    for line_text, line_col in terminal_lines:
        if line_text and line_col:
            add_textbox(slide, rx + cm(0.4), tl_y, rw - cm(0.8), cm(0.36),
                        line_text, 9.5, color=line_col)
        tl_y += cm(0.32)

    # Three mini-cards at the bottom
    mc_w  = cm(3.8)
    mc_h  = cm(1.8)
    mc_y  = tm_y + tm_h + cm(0.28)
    mc_gap = cm(0.2)
    mc_titles = ["Reads & reasons", "Writes & runs", "Iterates"]
    mc_bodies = [
        "Full codebase before writing a line.",
        "Files edited. Tests run automatically.",
        "Self-corrects if tests fail.",
    ]
    for i in range(3):
        mc_x = rx + i * (mc_w + mc_gap)
        add_dark_card(slide, mc_x, mc_y, mc_w, mc_h, mc_titles[i], mc_bodies[i])

    # Coral CTA band at very bottom
    cta_y = H - cm(2.2)
    cta_x = cm(1.2)
    cta_w = W - cm(2.4)
    rounded_rect(slide, cta_x, cta_y, cta_w, cm(1.9), C_PRIMARY, radius_emu=cm(0.3))
    add_textbox(slide, cta_x + cm(0.7), cta_y + cm(0.2), cta_w - cm(5), cm(0.65),
                "You're ready to build.", 20, color=C_WHITE)
    add_textbox(slide, cta_x + cm(0.7), cta_y + cm(0.9), cta_w - cm(5), cm(0.55),
                "Open Antigravity. Drop in a project. Send your first prompt.", 11,
                color=RGBColor(0xff, 0xff, 0xff))
    rounded_rect(slide, cta_x + cta_w - cm(4.0), cta_y + cm(0.55), cm(3.4), cm(0.75),
                 C_WHITE, radius_emu=cm(0.2))
    add_textbox(slide, cta_x + cta_w - cm(4.0), cta_y + cm(0.62), cm(3.4), cm(0.55),
                "Back to overview →", 11, bold=True, color=C_PRIMARY, align=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════

async def main():
    # Capture the orbit illustration from slide-1 to embed as image
    orbit_cache = BASE / "_orbit_cache.png"
    if not orbit_cache.exists():
        print("Capturing orbit illustration…")
        img = await capture_orbit("slide-1.html", orbit_cache)
        orbit_bytes = img
    else:
        orbit_bytes = orbit_cache.read_bytes()

    print("Building editable.pptx…")
    prs = Presentation()
    prs.slide_width  = Emu(W)
    prs.slide_height = Emu(H)

    build_intro(prs)
    print("  Intro slide done")
    build_slide1(prs, orbit_bytes)
    print("  Slide 1 done")
    build_slide2(prs)
    print("  Slide 2 done")
    build_slide3(prs)
    print("  Slide 3 done")
    build_slide4(prs)
    print("  Slide 4 done")

    prs.save(OUT)
    print(f"\nSaved → {OUT}")


asyncio.run(main())
