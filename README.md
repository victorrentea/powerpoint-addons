# powerpoint-addons

PowerPoint (Mac) task-pane add-in "Victor Tools", button on the Home tab. One pane, no tabs: every button is in sight (details in each button's tooltip).
- **Code block** — paste code, pick language/theme/size → syntax-highlighted JetBrains Mono text box (VS Code Dark+/Light+ colors).
- **Selected picture → Invert** — select a picture → "Invert selected image" replaces it with its negative (black ↔ white, alpha kept). Needs PowerPointApi 1.10 (Mac 16.105+).
- **Selected picture → Remove background** — cuts the subject out of the selected picture with BiRefNet on the Mac's GPU, trimmed to the subject and placed where the subject was. Moved here from Victor Addons' ⌘⇧V bezel (2026-10-06): PowerPoint is the only place it's needed.
- **Animated group → Add selection** — select an animated group plus the shapes to add: they join the group, which keeps its id, so its animation and its place in the build order stay. Office.js has no animation API and ungroup + group makes a new id the animation doesn't know, so the pane exports the slide (`exportAsBase64`, PowerPointApi 1.8), `web/regroup.js` moves the shapes inside the `<p:grpSp>` in the XML (into the group's own scaled coordinate space), and the rebuilt slide replaces the original (`insertSlidesFromBase64` after it, then delete). The added shapes lose their own effects (PowerPoint animates only top-level shapes); placeholders, tables and rotated groups are refused. Tests: `npm test`.
- **Text-heavy slides** — ranks slides by word count; click to jump.

## How it's wired
- `web/` static task pane, served over HTTPS by `server.js` on `https://localhost:44344` (Office loads add-ins only over https).
- `certs/` self-signed localhost cert (git-ignored), trusted in the login keychain for SSL.
- `manifest.xml` sideloaded by copying to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/victor-tools.xml`.
- Requires PowerPointApi 1.5.
- ✂️ Background removal can't run in the pane: `server.js` answers `POST /bg-remove` (PNG in, trimmed PNG + its box out) through `bg-remove.js`, which keeps `bg-remove/cutout_server.py` (BiRefNet, PyTorch MPS fp16, edge bleed) warm over a JSON-lines pipe. It lives as long as PowerPoint does: a 30 s `pgrep` watcher in `server.js` loads it when PowerPoint is running and stops it when PowerPoint quits (~5 GB; ~10 s cold, ~1 s warm). A 30-min idle timeout came first and missed: the clicks come hours apart, so nearly every one hit a cold model. `uv run --project bg-remove` builds the venv on first use (needs `/opt/homebrew/bin/uv`); its stderr goes to `/tmp/powerpoint-addons-bg-remove.log`. Deps and the Hub commit of `ZhengPeng7/BiRefNet` are pinned.

## Run
Installed as LaunchAgent `ro.victorrentea.powerpoint-addons` (source: `launchagent.plist`), starts at login, log in `/tmp/powerpoint-addons.log`.
Restart after editing `server.js`: `launchctl kickstart -k gui/$(id -u)/ro.victorrentea.powerpoint-addons`.
Edits under `web/` need no restart (served with `no-store`), just reopen the pane.
