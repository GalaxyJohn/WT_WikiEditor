# War Thunder Wiki Tree Editor

Browser-side Tampermonkey editor for War Thunder Wiki aviation tech tree pages.

## Features

- Toggleable edit mode
- Hover selection for tree cells and rows
- Custom context menu
- Insert, edit, copy, paste, delete, and drag vehicle cards
- Folder create/unpack support
- Row and rank editing
- Undo and redo
- Save changed data

## Install

Download Tempermonkey or AddGuard.
Chromium v93+ - [Tampermonkey](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo?hl=ko)
Firefox v102+ - [Tampermonkey](https://addons.mozilla.org/ko/firefox/addon/tampermonkey/)
Others - [AdGuard](https://adguard.com/ko/welcome.html)

https://github.com/GalaxyJohn/WT_WikiEditor/releases/download/release/wtte.user.js


## Default controls

- `Ctrl+Shift+E`: toggle edit mode
- `Ctrl+Click`: multi-select cards for foldering
- `Ctrl+Z`: undo
- `Ctrl+Y`: redo

## Repo notes

- `wiki.html` is treated as a local analysis fixture and is ignored by git.
- If you want to commit a local snapshot for debugging, remove `wiki.html` from `.gitignore`.
