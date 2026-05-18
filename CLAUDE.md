
# Asset Master — CLAUDE.md

## Project Context

PowerPoint Office Web Add-in (Task Pane). Vanilla HTML/CSS/JS, no build step.
- **Repo:** https://github.com/YaromStuff/asset-master-addin
- **Live URL:** https://yaromstuff.github.io/asset-master-addin/ (served from `gh-pages` branch)
- **Drive folder:** https://drive.google.com/drive/folders/1kuLZ5Ayvn6Fdl7CK7o0qJWVQVumtE3VE
- **Local files:** `assets-addin/` inside this directory
- **Manifest:** `assets-addin/asset master.xml` — must also copy to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` to activate

### Architecture
```
index.html → Drive API (googleapis.com) → lists subfolders → lists files per folder
```
Live load from Drive on every task pane open. Cache: `sessionStorage` key `am_cache_v2` (bump on URL format or schema change).

### Key files
- `assets-addin/index.html` — all UI + JS
- `assets-addin/asset master.xml` — PowerPoint manifest
- `.github/workflows/deploy.yml` — injects API key at build time, deploys to `gh-pages`

### Deployment flow
1. Push `main` → Action runs
2. Action replaces `%%GOOGLE_API_KEY%%` w/ Secret
3. Deploy → `gh-pages` branch → Pages serves it

---

## Maintenance Checklist (before compact / handoff)

- [ ] Add new gotchas from this session
- [ ] Update manifest XML if Pages URL changed + re-copy to `wef` folder
- [ ] Update `AppDomains` in manifest if new external domains fetched
- [ ] Delete duplicate `CLAUDE.md` inside `assets-addin/` if reappears

---

## Gotchas — Lessons Learned

### PowerPoint JS API

**`addImage` is `undefined` on API < 1.4 — always use `setSelectedDataAsync` as fallback.**
`slide.shapes.addImage` is literally `undefined` on older versions → throws "not a function". Fallback:
```js
Office.context.document.setSelectedDataAsync(b64, {
  coercionType: Office.CoercionType.Image,
  imageLeft: left, imageTop: top,
  imageWidth: imgW, imageHeight: imgH, // all in points
}, r => r.status === Office.AsyncResultStatus.Failed ? reject(...) : resolve());
```
`setSelectedDataAsync` works from API 1.1+. Detect with `typeof slide.shapes.addImage !== "function"` before calling.

**`getSelectedSlides()` requires API 1.5.** Always fall back to `ctx.presentation.slides.items[0]`.

**`addAudioFromBase64` requires API 1.5.** Likely unavailable if `addImage` (1.4) isn't. Fall back: download + show "הוספה → שמע → שמע במחשב שלי".

**SVG: pass base64 directly — never rasterize.**
PowerPoint 2016+ accepts SVG as native vector in `addImage`. Base64-encode the raw SVG bytes. For dimensions, parse `viewBox` from SVG text — don't rely on `naturalWidth` (returns 0 when SVG has no explicit `width/height`).

**`addImage` takes pure base64 — no data URI prefix.**
Strip `data:image/...;base64,` via `fr.result.split(",")[1]`. Same for `setSelectedDataAsync`.

**Shape size/position in points.**
Standard slide = 720pt × 540pt (10" × 7.5" at 72pt/in). `setSelectedDataAsync` image options also use points.

**Load order matters.**
`slides.load("items")` → `await ctx.sync()` before accessing `slides.items`. Missing load → "PropertyNotLoaded".

**All Office API calls must be inside `Office.onReady()`.**

---

### Google Drive API

**Use `alt=media` for full-quality download — not `uc?export=download`.**
`uc?export=download` can return compressed/resized versions. Correct endpoint:
```
https://www.googleapis.com/drive/v3/files/{id}?alt=media&key={KEY}
```
CORS-compatible, works for publicly shared files with API key.

**File must be shared as "Anyone with the link".**
Private files return HTTP 200 with HTML error page → silent base64 corruption on insert.

**Drive API `thumbnailLink` uses `=s220` suffix.** Replace w/ `=s200` for display thumbnails.

**Drive API v3 requires API key even for public folders.** Store as GitHub Secret, inject at build. Restrict key to `https://yaromstuff.github.io/*` in Google Cloud Console.

**`googleapis.com` must be in manifest `AppDomains`** alongside `drive.google.com` and `lh3.googleusercontent.com`.

**ROOT_FOLDER_ID is the last path segment of the folder URL.**

---

### Manifest + Install

**Manifest in `wef` is separate from the project file.** After updating `assets-addin/asset master.xml`, manually copy to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`. Stale `wef` copy = "cannot load add-in" error in PowerPoint. Restart PowerPoint after copying.

**Manifest filename irrelevant.** PowerPoint reads any `.xml` in `wef` root. Subdirs ignored.

**resid values must be ≤ 32 chars.**

**`FunctionFile` resid must exist even if unused** (required by schema in `DesktopFormFactor`).

**GitHub Pages URL:** `https://OWNER.github.io/REPO` — no trailing slash in manifest URLs.

---

### GitHub CLI (zsh)

**`gh api -f source[branch]=main` fails in zsh** — `[branch]` treated as glob. Use `--field`:
```bash
gh api ... --method POST --field "source[branch]=main" --field "source[path]=/"
```

**Pushing `.github/workflows/` requires `workflow` scope:**
```bash
gh auth refresh -h github.com --scopes workflow
```

**Multiple accounts — push uses wrong account.** `gh auth switch` changes gh but not macOS keychain:
```bash
TOKEN=$(gh auth token)
git -c "credential.helper=" -c "credential.helper=!f(){ echo \"username=x-token\"; echo \"password=$TOKEN\"; };f" push
```

---

### API Key Security

**Never commit API keys to public repo.** Use `%%PLACEHOLDER%%` in source, inject via GitHub Actions:
```yaml
run: sed -i "s|%%MY_KEY%%|${MY_KEY}|g" index.html
```
Key visible in deployed HTML source but absent from git history.

---

### UI / Accessibility

**Contrast minimums:** `#6e6b69` for muted text on white (4.5:1). `rgba(255,255,255,0.82)` min on `#0078d4` blue.
**`<div>` as button:** needs `role="button"`, `tabindex="0"`, keydown handler for Enter/Space.
**Status messages:** `role="status" aria-live="polite" aria-atomic="true"`.
**`@media (prefers-reduced-motion: reduce)`** must disable all animations (shimmer, spin, hover transform, flash, fade-in).

---

### Vanilla JS

**Versioned sessionStorage keys.** Current: `am_cache_v2`. Bump when URL format or data schema changes.
**`URL.createObjectURL` must be revoked** inside `onload`/`onerror` to avoid memory leaks.
**`type="button"` on all buttons** — default is `type="submit"`.
**`const` DOM refs can't be reassigned** — use a persistent wrapper div.

---

### Workflow

**After every code change — update this file.** Add new gotchas, remove outdated ones, keep it clean.
**Run `ui-ux-pro-max` after building**, not before.
**Keep only one CLAUDE.md — at project root.** Delete any copy in `assets-addin/`.
