# War Thunder Wiki Tree Editor

Browser-side Tampermonkey editor for War Thunder Wiki aviation tech tree pages.

## Main file

- `wtte.user.js`: userscript for `https://wiki.warthunder.com/*`

## Features

- Toggleable edit mode
- Hover selection for tree cells and rows
- Custom context menu
- Insert, edit, copy, paste, delete, and drag vehicle cards
- Folder create/unpack support
- Row and rank editing
- Undo and redo
- Arrow show/hide toggle
- Korean and English UI toggle
- Local persistence with `localStorage`

## Install

1. Open Tampermonkey.
2. Create a new userscript.
3. Paste the contents of `wtte.user.js`.
4. Save and open a War Thunder Wiki aircraft tree page.

## Default controls

- `Ctrl+Shift+E`: toggle edit mode
- `Ctrl+Click`: multi-select cards for foldering
- `Ctrl+Z`: undo
- `Ctrl+Y`: redo

## Repo notes

- `wiki.html` is treated as a local analysis fixture and is ignored by git.
- If you want to commit a local snapshot for debugging, remove `wiki.html` from `.gitignore`.
