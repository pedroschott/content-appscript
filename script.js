/**
 * Inserts random images from a Google Drive folder into a Google Sheet grid.
 * Configure the folderId, startColumn, imagesPerRow, cellSize, and startRow as needed.
 */
function insertRandomDriveImages() {
  const folderId     = 'YOUR GOOGLE DRIVE FOLDER ID';
  const startColumn  = 12;    // e.g., Column L (A=1, B=2, ...)
  const imagesPerRow = 8;
  const cellSize     = 100;   // px
  const startRow     = 2;     // skip header row

  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const rowsCnt = lastRow - startRow + 1;

  // 1) Pull all image blobs once
  const iter = DriveApp.getFolderById(folderId).getFiles();
  const blobs = [];
  while (iter.hasNext()) {
    const f = iter.next();
    if (f.getMimeType().startsWith('image/')) blobs.push(f.getBlob());
  }
  if (!blobs.length) {
    Logger.log('No images in folder!');
    return;
  }

  // 2) Shuffle images
  for (let i = blobs.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [blobs[i], blobs[j]] = [blobs[j], blobs[i]];
  }

  // 3) Find existing image-anchors
  const filled = new Set();
  sheet.getImages().forEach(img => {
    const c = img.getAnchorCell();
    if (
      c.getRow()    >= startRow &&
      c.getColumn() >= startColumn &&
      c.getColumn() < startColumn + imagesPerRow
    ) {
      filled.add(`${c.getRow()},${c.getColumn()}`);
    }
  });

  // 4) Build & shuffle empty-cell list
  const empty = [];
  for (let r = startRow; r <= lastRow; r++) {
    for (let o = 0; o < imagesPerRow; o++) {
      const c = startColumn + o;
      if (!filled.has(`${r},${c}`)) empty.push({ row: r, col: c });
    }
  }
  for (let i = empty.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [empty[i], empty[j]] = [empty[j], empty[i]];
  }

  // 5) Batch-resize grid
  sheet.setColumnWidths(startColumn, imagesPerRow, cellSize);
  sheet.setRowHeights(startRow, rowsCnt, cellSize);

  // 6) Insert images into empty cells
  let idx = 0;
  empty.forEach(cell => {
    const img = sheet.insertImage(blobs[idx], cell.col, cell.row);
    const scale = Math.min(cellSize / img.getWidth(), cellSize / img.getHeight());
    img.setWidth(img.getWidth() * scale).setHeight(img.getHeight() * scale);
    idx = (idx + 1) % blobs.length;
  });

  Logger.log(`Inserted ${empty.length} images at ${new Date().toLocaleTimeString()}`);
}

/**
 * Deletes all images from the active sheet.
 */
function deleteAllImages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const images = sheet.getImages();

  images.forEach(img => {
    img.remove();
  });

  Logger.log(`Deleted ${images.length} images at ${new Date().toLocaleTimeString()}`);
}
