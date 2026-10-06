# victor-ppt-addin

PowerPoint (Mac) task-pane add-in "Victor Tools", button on the Home tab:
- **Code block** — paste code, pick language/theme/size → syntax-highlighted JetBrains Mono text box (VS Code Dark+/Light+ colors).
- **Image** — select a picture → "Invert selected image" replaces it with its negative (black ↔ white, alpha kept). Needs PowerPointApi 1.10 (Mac 16.105+).
- **Text-heavy slides** — ranks slides by word count; click to jump.

## How it's wired
- `web/` static task pane, served over HTTPS by `server.js` on `https://localhost:44344` (Office loads add-ins only over https).
- `certs/` self-signed localhost cert (git-ignored), trusted in the login keychain for SSL.
- `manifest.xml` sideloaded by copying to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/victor-tools.xml`.
- Requires PowerPointApi 1.5.

## Run
Installed as LaunchAgent `ro.victorrentea.ppt-addin` (source: `launchagent.plist`), starts at login, log in `/tmp/ppt-addin.log`.
Restart after editing `server.js`: `launchctl kickstart -k gui/$(id -u)/ro.victorrentea.ppt-addin`.
Edits under `web/` need no restart (served with `no-store`), just reopen the pane.
