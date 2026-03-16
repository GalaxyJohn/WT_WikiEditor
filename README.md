# War Thunder Wiki Tree Editor

Browser-side Tampermonkey editor for War Thunder Wiki aviation tech tree pages.

## Features

- Toggleable edit mode
- Hover selection for tree cells and rows
- Custom context menu
- Insert, edit, copy, paste, delete, and drag vehicle cards
- Folder create/unpack support
- Row and rank editing
- Arrow show/hide toggle
- Undo and redo
- Local persistence with `localStorage`

## Install

Download Tempermonkey or AddGuard.
- Chromium v98+ - [Tampermonkey](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo?hl=ko)
- Firefox v102+ - [Tampermonkey](https://addons.mozilla.org/ko/firefox/addon/tampermonkey/)
- Others - [AdGuard](https://adguard.com/ko/welcome.html)

In desktop Chrome/Edge 138+:
1. Open extension settings by right-clicking the Tampermonkey icon (1) and selecting "Manage Extension" (2).
<img width="389" height="543" alt="Image" src="https://github.com/user-attachments/assets/d3ed4d74-a7df-43d3-9ef9-3b11ddb1c5f5" />

2. Locate and enable the "Allow User Scripts" toggle.
<img width="880" height="105" alt="Image" src="https://github.com/user-attachments/assets/d3c045b3-740c-4758-bd00-811116f54a90" />

Then click [here](https://github.com/GalaxyJohn/WT_WikiEditor/releases/download/release/wtte.user.js) to download the latest released script.

## Default controls
### Keyboard actions
- `Right click`: Open context menu
- `Ctrl + Click`: Multi-select / deselect cells
- `Ctrl + Drag`: Multi-select cells along the drag path
- `Shift + Drag`: Select a rectangular range
- `Click + tab`: Open custom tree preset selection menu
- `Click folder`: Open / close folder
### Mouse actions
- `Ctrl + Shift + E`: Toggle edit mode
- `Ctrl + Z`: Undo
- `Ctrl + Y` or `Ctrl + Shift + Z`: Redo
- `Ctrl + C`: Copy selected equipment or folder
- `Ctrl + V`: Paste equipment or folder into the selected cell
- `Delete`: If any equipment exists in the selected range, delete the equipment; if all cells are empty, delete the row
- `Backspace`: When a range is selected, delete the row / rank / equipment structure within that range
- `Esc`: Close right-click menu, close modal, cancel drag

## Repo notes

- `wiki.html` is treated as a local analysis fixture and is ignored by git.
- If you want to commit a local snapshot for debugging, remove `wiki.html` from `.gitignore`.
