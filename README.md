# powerpoint-addons

PowerPoint (Mac) task-pane add-in "Victor Tools", button on the Home tab:
- **Code block** — paste code, pick language/theme/size → syntax-highlighted JetBrains Mono text box (VS Code Dark+/Light+ colors).
- **Image** — select a picture → "Invert selected image" replaces it with its negative (black ↔ white, alpha kept). Needs PowerPointApi 1.10 (Mac 16.105+).
- **Image → Remove background** — cuts the subject out of the selected picture with BiRefNet on the Mac's GPU, trimmed to the subject and placed where the subject was. Moved here from Victor Addons' ⌘⇧V bezel (2026-10-06): PowerPoint is the only place it's needed.
- **Text-heavy slides** — ranks slides by word count; click to jump.

## How it's wired
- `web/` static task pane, served over HTTPS by `server.js` on `https://localhost:44344` (Office loads add-ins only over https).
- `certs/` self-signed localhost cert (git-ignored), trusted in the login keychain for SSL.
- `manifest.xml` sideloaded by copying to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/victor-tools.xml`.
- Requires PowerPointApi 1.5.
- ✂️ Background removal can't run in the pane: `server.js` answers `POST /bg-remove` (PNG in, trimmed PNG + its box out) through `bg-remove.js`, which keeps `bg-remove/cutout_server.py` (BiRefNet, PyTorch MPS fp16, edge bleed) warm over a JSON-lines pipe. Opening the Image tab prewarms it (~10 s cold, ~1–3 s warm); it exits after 30 min idle. `uv run --project bg-remove` builds the venv on first use (needs `/opt/homebrew/bin/uv`); its stderr goes to `/tmp/powerpoint-addons-bg-remove.log`. Deps and the Hub commit of `ZhengPeng7/BiRefNet` are pinned.

## Run
Installed as LaunchAgent `ro.victorrentea.powerpoint-addons` (source: `launchagent.plist`), starts at login, log in `/tmp/powerpoint-addons.log`.
Restart after editing `server.js`: `launchctl kickstart -k gui/$(id -u)/ro.victorrentea.powerpoint-addons`.
Edits under `web/` need no restart (served with `no-store`), just reopen the pane.
