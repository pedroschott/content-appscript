# Google Apps Script: Bulk Insert & Delete Drive Images in Google Sheets

## Overview

This script helps you **automatically insert images from a Google Drive folder into a Google Sheets spreadsheet**, perfectly resizing and arranging them in a grid. You can also delete all inserted images with one click.  
Ideal for catalogs, visual dashboards, or any sheet that needs fast, bulk image insertion.

---

## Features

- **Bulk insert images:** Adds all images from a Drive folder, randomized and auto-resized.
- **Neat grid layout:** Set starting row/column, columns per row, and cell size.
- **No duplicate placements:** Fills only empty cells, leaves existing images untouched.
- **Batch delete:** Instantly remove all images from your active sheet.
- **Customizable:** Change grid size, starting points, and image sizes in the script.

---

## How It Works

### Insert Random Drive Images

- Loads all image files from a specified Google Drive folder.
- Randomizes order and finds empty spots in a specified cell grid (default: 8 images per row, starting at row 2, column L).
- Resizes images to fit cell size.
- Skips cells that already contain images.

### Delete All Images

- Removes **all images** from the active sheet.

---

## Setup & Usage

1. **Copy the script** into your Google Sheet via Extensions > Apps Script.
2. Replace `'YOUR GOOGLE DRIVE FOLDER ID'` with your actual Drive folder ID.
3. (Optional) Change these in the script to match your layout:
    - `startColumn` (default: 12, which is column L)
    - `imagesPerRow` (default: 8)
    - `cellSize` (default: 100 px)
    - `startRow` (default: 2)
4. **Run** `insertRandomDriveImages` to insert images.
5. **Run** `deleteAllImages` to remove all images.

---

## Permissions

The script will ask for Google Drive and Sheets permissions the first time you run it.

---

## Notes

- Works only on the active sheet.
- Only images from the chosen folder are inserted.
- For best performance, avoid very large folders (>1000 images).

---

## License

MIT – Free to use, modify, and share.
