function insertRandomDriveImagesUnique() {
  const folderId     = 'YOUR GOOGLE DRIVE FOLDER ID';
  const startColumn  = 12;    // H
  const imagesPerRow = 8;
  const cellSize     = 100;   // px
  const startRow     = 2;     // skip header

  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const rowsCnt = lastRow - startRow + 1;

  // 1) Pull all image blobs
  const blobs = [];
  const files = DriveApp.getFolderById(folderId).getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().startsWith('image/')) {
      blobs.push(f.getBlob());
    }
  }
  if (blobs.length === 0) {
    Logger.log('No images found in folder!');
    return;
  }

  // 2) Shuffle blobs (Fisher–Yates)
  for (let i = blobs.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [blobs[i], blobs[j]] = [blobs[j], blobs[i]];
  }

  // 3) Find which cells already have images
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

  // 4) Build & shuffle list of empty cells
  const empty = [];
  for (let r = startRow; r <= lastRow; r++) {
    for (let offset = 0; offset < imagesPerRow; offset++) {
      const c = startColumn + offset;
      if (!filled.has(`${r},${c}`)) {
        empty.push({ row: r, col: c });
      }
    }
  }
  for (let i = empty.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [empty[i], empty[j]] = [empty[j], empty[i]];
  }

  // 5) Resize rows & columns
  sheet.setColumnWidths(startColumn, imagesPerRow, cellSize);
  sheet.setRowHeights(startRow, rowsCnt, cellSize);

  // 6) Insert each blob into one unique empty cell
  const toInsert = Math.min(blobs.length, empty.length);
  for (let i = 0; i < toInsert; i++) {
    const { row, col } = empty[i];
    const blob = blobs[i];
    const img  = sheet.insertImage(blob, col, row);
    // scale to fit
    const scale = Math.min(cellSize / img.getWidth(), cellSize / img.getHeight());
    img.setWidth(img.getWidth() * scale)
       .setHeight(img.getHeight() * scale);
  }

  Logger.log(`Inserted ${toInsert} unique images at ${new Date().toLocaleTimeString()}`);
}

function deleteAllImages() {
  const sheet  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const images = sheet.getImages();
  images.forEach(img => img.remove());
  Logger.log(`Deleted ${images.length} images at ${new Date().toLocaleTimeString()}`);
}
