---
name: presentation
description: Generate beautiful HTML slideshow presentations using the Anthropic-style design system (cream canvas + coral + dark navy, Cormorant Garamond serif + Inter sans). Produces an index page + N slides with keyboard navigation (← →), responsive layouts, and visual-first SVG graphics. Optionally exports to PowerPoint via Playwright screenshots and/or native python-pptx shapes. Use when the user types /presentation or asks to "build a slideshow", "create a presentation", "make slides about X".
---

# /presentation — Anthropic-style HTML slideshow generator

Builds a multi-slide HTML presentation that looks like a Claude.com editorial landing page: warm cream canvas, coral CTAs, dark navy product surfaces, Cormorant Garamond serif headlines, Inter sans body. Each slide has a fixed nav bar with `← →` keyboard navigation. The index page IS slide 0 in the sequence — it's an intro, not just an aggregator.

## When to invoke

User says any of:
- `/presentation`
- "build a slideshow"
- "create an HTML presentation"
- "make slides about X"
- "build a deck about X"

## Ask first (one round only)

If the user hasn't specified, ask in a single message:
1. **Topic** — what is the presentation about?
2. **Slide count** — how many content slides? (default: 4)
3. **Slide titles** — give each slide a one-line title
4. **PPTX export?** — want PowerPoint versions too? (default: no)

Don't ask follow-ups after that. Make sensible decisions and ship.

## File structure to produce

```
<project>/
├── design-system.md          # (already exists or copy from skill)
├── index.html                # Intro / overview slide
├── slide-1.html
├── slide-2.html
├── slide-N.html
└── slide-base.css            # Shared nav + utility styles
```

If exporting to PPTX:
```
├── make_screenshot_pptx.py   # Pixel-perfect via Playwright
├── make_editable_pptx.py     # Editable via python-pptx shapes
├── pixel-perfect.pptx
└── editable.pptx
```

---

## Design system tokens (copy verbatim into every slide)

```css
:root {
  --primary: #cc785c;        /* coral — CTAs, accents */
  --primary-active: #a9583e;
  --ink: #141413;            /* warm near-black */
  --body: #3d3d3a;
  --muted: #6c6a64;
  --muted-soft: #8e8b82;
  --hairline: #e6dfd8;
  --canvas: #faf9f5;         /* cream — NOT white */
  --surface-card: #efe9de;   /* cream cards */
  --surface-dark: #181715;   /* dark navy product surfaces */
  --surface-dark-elevated: #252320;
  --surface-dark-soft: #1f1e1b;
  --on-dark: #faf9f5;
  --on-dark-soft: #a09d96;
  --accent-teal: #5db8a6;
  --accent-amber: #e8a55a;

  --serif: 'Cormorant Garamond', 'Tiempos Headline', Georgia, serif;
  --sans: 'Inter', -apple-system, sans-serif;
  --mono: 'JetBrains Mono', ui-monospace, monospace;
}
```

Load fonts via:
```html
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:wght@400;500&family=Inter:wght@400;500&family=JetBrains+Mono:wght@400&display=swap" rel="stylesheet" />
```

## Style rules (non-negotiable)

1. **Alternate surfaces** — cream → dark → cream → dark. Never two same-coloured slides in a row. Index is dark.
2. **Serif for display, sans for body.** Cormorant Garamond weight 400 (never bold) with negative letter-spacing (`-1px` to `-1.5px`) for all H1s.
3. **Coral is scarce.** Only on primary CTAs, badges, accents, and full-bleed coral callout cards. Never on body text.
4. **Visual-first.** Every slide must be understandable WITHOUT reading the body text — use SVG mockups, illustrations, product chrome (terminal windows, file explorers, plugin panels, code blocks).
5. **Less is more on copy.** Lead text = one short sentence. Bullets = `<strong>label.</strong> short phrase.` — not paragraphs.
6. **No nav text content.** Buttons like "Back" / "Next slide" inside the body are forbidden — the nav bar handles all navigation.

## Required components

### Nav bar (every slide, fixed top)

```html
<nav class="slide-nav">  <!-- add class="dark" on dark slides -->
  <a href="index.html" class="nav-brand">
    <svg width="20" height="20" viewBox="0 0 28 28" fill="none">
      <path d="M14 2L15.2 12.8L22 6L15.2 14L26 14L15.2 14.8L22 22L14.8 15.2L14 26L13.2 15.2L6 22L12.8 14.8L2 14L12.8 13.2L6 6L13.2 12.8L14 2Z" fill="#cc785c"/>
    </svg>
    <span class="nav-brand-name">{Project Name}</span>
  </a>
  <div class="nav-controls">
    <a href="{prev}" class="nav-btn">←</a>
    <span class="nav-counter">{N} / {total}</span>
    <a href="{next}" class="nav-btn">→</a>
  </div>
</nav>
```

The spike-mark SVG path above is the Anthropic 8-spoke radial — use it as-is everywhere.

### Keyboard navigation (every slide, before `</body>`)

```html
<script>
  document.addEventListener('keydown', e => {
    if (e.key === 'ArrowRight') window.location.href = '{next}';
    if (e.key === 'ArrowLeft')  window.location.href = '{prev}';
  });
</script>
```

On the first slide (index), only the right arrow is bound. On the last slide, only the left arrow.

### Visual elements to use liberally

- **Code window** — dark `surface-dark-soft` background, traffic-light dots (red `#c64545` / amber `#e8a55a` / green `#5db872`), JetBrains Mono terminal output with colour-coded prefixes (✓ green for success, coral for "Done", teal for prompts, amber for filenames)
- **Plugin/file panel mockup** — dark `surface-dark-elevated` with titlebar + rows (each with a small SVG icon + name + description + optional toggle pill)
- **Step list** — circle-numbered steps (done = coral fill + ✓, active = coral border, upcoming = grey)
- **Drop zone** — dashed coral border on a `rgba(204,120,92,0.04)` tint
- **Feature card grid** — 3-up cards on `surface-card` cream
- **Coral CTA band** — full-bleed `--primary` background with serif h2 + white button

## slide-base.css (drop in as-is)

```css
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root { /* paste tokens block from above */ }

html, body { height: 100%; font-family: var(--sans); }

.slide-nav {
  position: fixed; top: 0; left: 0; right: 0; height: 56px; z-index: 100;
  display: flex; align-items: center; justify-content: space-between;
  padding: 0 32px;
  background: rgba(250,249,245,0.92); backdrop-filter: blur(8px);
  border-bottom: 1px solid var(--hairline);
}
.slide-nav.dark {
  background: rgba(24,23,21,0.92);
  border-bottom-color: rgba(250,249,245,0.06);
}
.nav-brand { display: flex; align-items: center; gap: 8px; text-decoration: none; }
.nav-brand-name { font-size: 14px; font-weight: 500; color: var(--ink); }
.slide-nav.dark .nav-brand-name { color: var(--on-dark); }
.nav-controls { display: flex; align-items: center; gap: 8px; }
.nav-btn {
  display: flex; align-items: center; justify-content: center;
  width: 36px; height: 36px; border-radius: 9999px;
  background: var(--canvas); border: 1px solid var(--hairline);
  color: var(--ink); text-decoration: none;
}
.slide-nav.dark .nav-btn {
  background: var(--surface-dark-elevated);
  border-color: rgba(250,249,245,0.08);
  color: var(--on-dark);
}
.nav-btn.disabled { opacity: 0.25; pointer-events: none; }
.nav-counter { font-size: 12px; font-weight: 500; letter-spacing: 1px; color: var(--muted); min-width: 48px; text-align: center; }
.slide-nav.dark .nav-counter { color: var(--on-dark-soft); }

.slide { min-height: 100vh; padding-top: 56px; }

.dot { width: 10px; height: 10px; border-radius: 50%; }
.dot.red { background: #c64545; }
.dot.amber { background: #e8a55a; }
.dot.green { background: #5db872; }
```

## Slide pattern — every content slide

```
┌─ Fixed nav bar (56px) ──────────────────────────────┐
├─ Eyebrow: "01 — {SECTION}" (coral, uppercase, 12px) ┤
├─ H1: Cormorant 40-60px, weight 400, letter-spacing -1px
├─ Lead: ONE short sentence, 16px, body/muted colour
├─ Two-column body:                                    │
│    LEFT  → 3 bullet rows (dot + bold label + phrase) │
│    RIGHT → visual mockup (code window / panel / etc) │
└─ Keyboard nav script                                 ┘
```

Don't deviate. The pattern is the brand.

---

## Workflow (do all steps)

1. **Confirm topic + slide count + titles** (one round of questions max).
2. **Write `design-system.md`** at project root if it doesn't exist — copy from the reference project at `/Users/jonocatliff/Documents/Claude Code/presentations-skool/design-system.md`.
3. **Write `slide-base.css`** with the block above.
4. **Write `index.html`** — dark surface, intro h1, 2-column or 2×2 card grid linking to each slide. Counter shows "Intro". Left arrow disabled. Include keyboard nav (right arrow only).
5. **Write `slide-1.html` through `slide-N.html`** — alternate cream / dark surfaces. Each one:
   - Has the nav bar with correct prev/next links and counter "N / total"
   - Has an eyebrow + serif H1 + one-line lead
   - Has a visual element on the right (use SVG, not images)
   - Ends with keyboard nav script
6. **Last slide ends with a coral CTA band** linking back to `index.html`.
7. **If PPTX requested**, write both Python scripts (reference: `/Users/jonocatliff/Documents/Claude Code/presentations-skool/make_screenshot_pptx.py` and `make_editable_pptx.py`). Install deps: `pip3 install python-pptx pillow playwright --break-system-packages && python3 -m playwright install chromium`.
8. **Confirm in one sentence** which files were created and tell the user to open `index.html` to start.

## Do / Don't

**Do:**
- Use SVG for every visual — no external image files
- Show real product chrome (terminals, plugin panels, file trees) over abstract marketing illustrations
- Let whitespace breathe — 72-96px section padding, 32px card padding
- Match content tone to Anthropic — literary, considered, not breathless

**Don't:**
- Don't use pure white (`#fff`) for cream backgrounds — always `#faf9f5`
- Don't bold Cormorant — weight 400 only
- Don't put coral on body text or as a fourth surface tone (no purple/green sections)
- Don't write paragraphs in bullets — one short clause per item
- Don't add hover styling beyond what's already in the base
- Don't add "Next slide" buttons inside slide content — nav bar only

## Reference implementation

The canonical example lives at `/Users/jonocatliff/Documents/Claude Code/presentations-skool/`. Read those files when you need a pixel-accurate template for any component.
