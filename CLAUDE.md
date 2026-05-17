# Asset Master — CLAUDE.md

## Project Context

PowerPoint Office Web Add-in (Task Pane). Vanilla HTML/CSS/JS, no build step.
Hosted on GitHub Pages: https://yaromstuff.github.io/asset-master-addin/
Repo: https://github.com/YaromStuff/asset-master-addin

---

## Save Gotchas — Lessons Learned

### PowerPoint JS API

**`addImage` requires API 1.4, not 1.3.**
The manifest says `MinVersion="1.3"` for installation compatibility, but `slide.shapes.addImage()` silently fails on 1.3. Wrap in try/catch and surface a clear Hebrew error.

**`getSelectedSlides()` requires API 1.5.**
Not available at 1.3 or 1.4. Always wrap in try/catch and fall back to `slides.items[0]`. Never assume selected slide is accessible.

**`addImage` takes pure base64 — no data URI prefix.**
Strip `data:image/...;base64,` before passing to `addImage`. Use `fr.result.split(",")[1]`.

**Shape size/position is in points, not pixels.**
Standard slide = 720pt × 540pt (10" × 7.5" at 72pt/in). Always calculate layout in points.

**Load order matters.**
Call `slides.load("items")` then `await ctx.sync()` before accessing `slides.items`. Forgetting the load call throws a cryptic "PropertyNotLoaded" error.

**All Office API calls must be inside `Office.onReady()`.**
Never call `PowerPoint.run` or any Office API at the top level. It will be undefined on load.

---

### Google Drive

**CORS blocks `fetch()` on direct Drive URLs.**
`https://drive.google.com/uc?export=download&id=FILE_ID` may be blocked by CORS in the add-in iframe. The thumbnail URL (`thumbnail?id=...&sz=w200`) works for display but not for fetching binary. Document this limitation to the user.

**Two different URL patterns for the same file:**
- Thumbnail (display only): `https://drive.google.com/thumbnail?id=FILE_ID&sz=w200`
- Download (insert): `https://drive.google.com/uc?export=download&id=FILE_ID`

**File must be shared as "Anyone with the link" — not just org.**
Private files silently return an HTML error page instead of the binary. The fetch succeeds (200) but the content is HTML, causing base64 corruption.

---

### GitHub CLI (zsh)

**`gh api -f source[branch]=main` fails in zsh.**
zsh interprets `[branch]` as a glob pattern → "no matches found". Always use `--field` instead of `-f` for bracket-notation keys:
```bash
# WRONG (zsh glob expansion)
gh api ... -f source[branch]=main

# CORRECT
gh api ... --method POST --field "source[branch]=main" --field "source[path]=/"
```

**Enabling GitHub Pages via CLI:**
```bash
gh api repos/OWNER/REPO/pages --method POST \
  --field "source[branch]=main" \
  --field "source[path]=/"
```

---

### Office Add-in Manifest

**resid values must be ≤ 32 characters.** Keep them short: `addin.url`, `icon.32`, `btn.label`. Long resids silently fail validation.

**`GITHUB_PAGES_URL` must be replaced in 8+ places in manifest.xml.**
The manifest references icons, source URLs, and support URLs — all must point to the live HTTPS URL. GitHub Pages URL format: `https://OWNER.github.io/REPO`

**`FunctionFile` resid must exist even if unused.**
The manifest schema requires it in `DesktopFormFactor`. Point it to the same `addin.url` as the task pane.

**`DefaultLocale: he-IL`** — set this if the add-in is Hebrew-only. Affects how Office displays the add-in in the store/dialog.

---

### UI / Accessibility

**`#a19f9d` on white fails WCAG AA (only ~2.85:1).**
Use `#6e6b69` as the minimum for secondary/muted text — passes 4.5:1. Never use `#a19f9d` or lighter for body text.

**`rgba(255,255,255,0.65)` on `#0078d4` is borderline.**
Use at least `0.82` opacity for white text on the Office blue. Lower values fail at small font sizes.

**`<div>` acting as a button/tab needs: `role`, `tabindex`, `keydown` handler.**
Without these three, keyboard users are completely locked out. Pattern:
```js
el.setAttribute("role", "button");
el.setAttribute("tabindex", "0");
el.addEventListener("keydown", e => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); doAction(); }
});
```

**Status messages need `aria-live="polite"` to reach screen readers.**
A div that gets text injected via JS is invisible to screen readers unless it has `aria-live`. Add `role="status" aria-live="polite" aria-atomic="true"` to `#st`.

**`@media (prefers-reduced-motion: reduce)` must disable ALL animations.**
Cover: shimmer skeleton, spin loader, card hover transform, insert flash, image fade-in. Users who request reduced motion get seizures/nausea from ignored animations.

**Structural icons → SVG. Decorative/data icons → emoji OK.**
The search icon (structural) was a 🔍 emoji — replaced with inline SVG. Category icons (🖥️ 🎨 ⭐ 🔊) come from user data and are acceptable as emoji.

---

### Vanilla JS Patterns

**Use versioned sessionStorage keys to avoid stale cache.**
`sessionStorage.getItem("am_cache_v1")` — bump the version suffix if the `assets.json` schema changes.

**`const` DOM refs can't be reassigned — use a wrapper element.**
Never plan to do `mainEl = newElement` if `mainEl` is `const`. Use a persistent wrapper div and update its `.innerHTML` or `appendChild`.

**`type="button"` on every button generated via innerHTML.**
Buttons without an explicit type default to `type="submit"`. Inside any ancestor form they'll submit. Always explicit.

**`URL.createObjectURL` must always be followed by `URL.revokeObjectURL`.**
Memory leak otherwise. Call revoke inside the `onload`/`onerror` callbacks.

---

### Workflow

**When the user provides a complete, detailed spec + auto mode is active → skip brainstorming.**
The brainstorming skill is for refining unclear ideas. A 200-line Hebrew spec with architecture, error handling, and deployment instructions is an approved design. Don't invoke the skill's hard gate — go straight to implementation.

**Run `ui-ux-pro-max` *after* building, not before.**
For spec-driven tasks, build first then run the design review to catch what the spec missed (contrast, ARIA, focus states, reduced-motion). Don't let it block execution.
