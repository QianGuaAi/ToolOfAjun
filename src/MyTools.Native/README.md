# MyTools Native

This directory contains the C/C++ native rewrite track for MyTools.

Current phase: Phase 1/2 native shell and foundation services, plus Phase 3 Codex Profiles scaffold.

Implemented scope:

- Win32 entry point.
- Single-instance mutex.
- Main window with native menu.
- Direct2D + DirectWrite rendering.
- Per-monitor DPI awareness.
- Tray icon with show/exit menu.
- Startup, runtime, and crash log files beside the executable.
- DPAPI smoke test.
- Foundation services for config, binary reads, atomic binary/UTF-8 file writes, cancellable scans, background tasks, process launch, TCP probes, and global hotkeys.
- A Codex current-folder backup service that can create DPAPI-protected `.bak.dpapi` backups from an explicit menu action or before a confirmed profile switch.
- A Codex diff summary service that can compare the first readable DPAPI-protected profile with the current `~/.codex` files by presence, size, line count, equality, and SHA-256 fingerprint without returning file contents.
- A Codex profile metadata edit service that can rename the first readable profile and edit its note, remark, or tags from confirmed explicit menu actions; rename trims and rejects empty values, values over 120 characters, and control characters, and synchronizes `active.json` when the renamed profile is the current active profile, while note/remark/tags use `PromptText`, trim input, allow empty values to clear fields, cap note/remark at 500 characters and tags at 200 characters, and reject C0/C1 control characters.
- A Codex current-folder import service that can DPAPI-protect the current `~/.codex/config.toml` and `auth.json` into a profile-library item after a confirmed explicit menu action.
- A Codex profile export service that can locate the first readable DPAPI-protected profile by display name and atomically export `config.toml` and `auth.json` to a user-selected directory after confirmation, while rejecting the current `~/.codex` directory to avoid overwriting current account files without backup.
- A Codex `.codexbox` package service that can export portable encrypted profile packages using the WPF-compatible `CDXB` PBKDF2/AES-CBC/HMAC format, and import them by re-protecting config/auth content with the current Windows user's DPAPI after explicit menu path/password/conflict-policy dialogs.
- A Codex latest-backup restore service that can select the newest DPAPI-protected `.bak.dpapi` backup and atomically restore `config.toml` and `auth.json` after a confirmed explicit menu action.
- A Codex explicit profile switch service that can locate the first readable DPAPI-protected profile by display name, back up the current `~/.codex`, and atomically write `config.toml`, `auth.json`, and `active.json` after a confirmed explicit menu action.
- A small C ABI JSON schema token helper for the lower-level native boundary.
- Module navigation for Home and the first Codex Profiles scaffold.
- Codex Profiles currently probes local file presence, reads DPAPI-protected profile-library metadata summaries by `items` object, marks `active.json`'s active profile, and exposes explicit menu actions for refresh, first-profile diff, current-folder backup, confirmed first-profile apply, confirmed current-folder import, confirmed latest-backup restore, `Export first profile files...`, `Rename first profile...`, `Edit first profile note...`, `Edit first profile remark...`, `Edit first profile tags...`, `Export .codexbox...`, and `Import .codexbox...`. The native UI only collects folders, `PromptText` metadata text, save/open paths, passwords, and the `.codexbox` import conflict policy, then dispatches through `CodexProfileModule`; profile-file export, rename, metadata read/write, and encrypted package import/export behavior stays in services, with `CodexProfileEditService::UpdateProfileMetadata()` writing the DPAPI-protected `profiles.json` and synchronizing `active.json` when the renamed profile is active, and `CodexProfileBoxService::ImportBox()` applying the selected Rename / Skip / Replace conflict policy. Plain navigation still does not create backups, generate diffs, rename profiles, edit note/remark/tags, import profiles, export profile files, import/export `.codexbox` packages, restore backups, expose embedded config/auth secrets, switch profiles, overwrite `~/.codex`, or contact networks.

Not implemented in this phase:

- FRP tunnel migration.
- Screenshot/recording.
- Multimedia.
- Schedule.
- WeChat tools.
- Native installer.

Build target:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\native-eval.ps1 -Build
```
