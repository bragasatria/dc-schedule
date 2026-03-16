/**
 * @file calculator.js
 * Core business logic for optimizing DC cable schedule assignments into rolls/drums.
 * utilizing ExcelJS to accurately map formats, widths, and complex layout properties.
 */
import ExcelJS from 'exceljs';

const MAX_LENGTH = 1000;
const MAX_EXECUTION_MS = 20000;

const hasUnassigned = (list) => list.some(item => item['Roll/Drum'] === null);

const calculateRolls = (dataList, headerNames) => {
  const startTime = performance.now();
  const keyLength = headerNames[1];

  let cablesToProcess = dataList.map((item, index) => ({
    ...item,
    _originalIndex: index,
    'Roll/Drum': item['Roll/Drum'] !== undefined ? item['Roll/Drum'] : null
  }));

  for (let i = 0; i < cablesToProcess.length; i++) {
    const rawVal = cablesToProcess[i][keyLength];
    const parsedVal = parseInt(rawVal, 10);
    if (isNaN(parsedVal)) throw new Error(`Data panjang harus angka. Cek nilai: "${rawVal}"`);
    cablesToProcess[i][keyLength] = parsedVal;
  }

  let rollCounter = 1;
  const MAX_ROLLS = 100;

  while (hasUnassigned(cablesToProcess) && rollCounter <= MAX_ROLLS) {
    if ((performance.now() - startTime) > MAX_EXECUTION_MS) {
      throw new Error(`Proses terlalu lama (>20 detik). Periksa kembali file input.`);
    }

    const unassignedCables = cablesToProcess.filter(item => item['Roll/Drum'] === null);
    if (unassignedCables.length === 0) break;

    let dp = { 0: { sum: 0, indices: [] } };

    for (const cable of unassignedCables) {
      const originalIndex = cable._originalIndex;
      const length = cable[keyLength];
      const currentWeights = Object.keys(dp).map(Number);

      for (const w of currentWeights) {
        const newW = w + length;
        if (newW <= MAX_LENGTH) {
          const currentSum = dp[w].sum;
          if (!dp[newW] || dp[newW].sum < (currentSum + length)) {
            dp[newW] = { sum: currentSum + length, indices: [...dp[w].indices, originalIndex] };
          }
        }
      }
    }

    const dpWeights = Object.keys(dp).map(Number).sort((a, b) => b - a);
    if (dpWeights.length > 1) {
      const bestIndices = dp[dpWeights[0]].indices;
      for (const idx of bestIndices) {
        cablesToProcess[idx]['Roll/Drum'] = rollCounter;
      }
    } else {
      break; 
    }

    rollCounter++;
  }

  return cablesToProcess.map(item => {
    const { _originalIndex, ...cleanItem } = item;
    return cleanItem;
  });
};

/**
 * Handles processing of raw Excel file Buffer, iterates sheet-by-sheet to 
 * resolve values, dynamically style layouts, generate Summary tabs, 
 * and return the compiled Output Blob mapped exactly to the original PHP Logic.
 */
export const processAndGenerateExcel = async (arrayBuffer, imageBuffer) => {
  const workbookInput = new ExcelJS.Workbook();
  await workbookInput.xlsx.load(arrayBuffer);

  const workbookOutput = new ExcelJS.Workbook();
  const globalSummaryData = [];
  
  let imageId = null;
  if (imageBuffer) {
    imageId = workbookOutput.addImage({
      buffer: imageBuffer,
      extension: 'png',
    });
  }

  // Creates Summary Sheet rigidly at index 0 
  const summarySheet = workbookOutput.addWorksheet('Summary DC Schedule');

  let sheetIndex = 0;

  for (const sheetIn of workbookInput.worksheets) {
    const sheetName = sheetIn.name;
    
    // Dynamically identify the header row to prevent dropping the first line of data
    let headerRowIdx = 1;
    let headerRow = sheetIn.getRow(headerRowIdx);
    
    const getVal = (val) => {
      if (val === null || val === undefined) return null;
      if (typeof val === 'object') {
        return val.richText ? val.richText.map(t => t.text).join('') : (val.result || val.text || val);
      }
      return val;
    };

    let labelHeader = getVal(headerRow.getCell(1).value);
    let lengthHeader = getVal(headerRow.getCell(2).value);

    // If row 1 is actually a merged title or completely empty, fallback to reading headers from row 2
    if (!labelHeader || !lengthHeader) {
        headerRowIdx = 2;
        headerRow = sheetIn.getRow(headerRowIdx);
        labelHeader = getVal(headerRow.getCell(1).value) || 'String Label';
        lengthHeader = getVal(headerRow.getCell(2).value) || 'Total Length';
    }

    const headers = [labelHeader, lengthHeader];
    const dataPerSheet = [];

    const maxRow = sheetIn.rowCount;
    const dataStartRow = headerRowIdx + 1; // Extract strictly starting AFTER the headers

    for (let r = dataStartRow; r <= maxRow; r++) { 
      const row = sheetIn.getRow(r);
      const valA = getVal(row.getCell(1).value);
      const valB = getVal(row.getCell(2).value);

      if (valA === null || valB === null) continue;

      dataPerSheet.push({
        [headers[0]]: valA,
        [headers[1]]: valB,
        'Roll/Drum': null
      });
    }

    if (dataPerSheet.length === 0) continue;

    let processedData;
    try {
      processedData = calculateRolls(dataPerSheet, headers);
    } catch (e) {
      throw new Error(`Gagal memproses sheet "${sheetName}". ${e.message}`);
    }

    const rolls = processedData.map(i => i['Roll/Drum']).filter(Boolean);
    const totalRollsSheet = new Set(rolls).size;

    globalSummaryData.push({ sheet_name: sheetName, total_roll: totalRollsSheet });

    // Output Sheet naming limitation logic
    let safeName = sheetName.replace(/[\\/?*[\]:]/g, '_');
    let finalSheetName = "Schedule " + safeName.substring(0, 31 - "Schedule ".length);
    if (finalSheetName === "Schedule ") finalSheetName = "Schedule Sheet";

    let suffixCount = 1;
    let originalName = finalSheetName;
    while (workbookOutput.getWorksheet(finalSheetName)) {
      const suffix = ` (${suffixCount})`;
      finalSheetName = originalName.substring(0, 31 - suffix.length) + suffix;
      suffixCount++;
    }

    const sheetOut = workbookOutput.addWorksheet(finalSheetName);

    // Title Merges & Fonts
    sheetOut.mergeCells('A1:C3');
    const titleCell = sheetOut.getCell('A1');
    titleCell.value = `DC SCHEDULE PROGRAM\n${sheetName}`;
    titleCell.font = { bold: true, size: 16, name: 'Calibri' };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    if (imageId !== null) {
      sheetOut.addImage(imageId, {
        tl: { col: 0, row: 0 },
        ext: { width: 140, height: 60 }
      });
    }

    // Set Dynamic Tab Colors (Red/Black Toggle)
    sheetIndex++;
    const tabColor = (sheetIndex % 2 === 1) ? 'FFFF0000' : 'FF000000';
    sheetOut.properties.tabColor = { argb: tabColor };

    const setBorders = (rowNum, cols = ['A','B','C']) => {
       cols.forEach(c => {
         sheetOut.getCell(`${c}${rowNum}`).border = { 
             top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} 
         };
       });
    };

    let currentRow = 6;
    
    // Explicitly Hardcode Row 6 Headers to prevent mapping overrides!

    const headerFillStyle = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
    
    sheetOut.getCell('A6').value = 'String Label';
    sheetOut.getCell('A6').font = { bold: true };
    sheetOut.getCell('A6').alignment = { horizontal: 'center', vertical: 'middle' };
    sheetOut.getCell('A6').fill = headerFillStyle;


    sheetOut.getCell('B6').value = 'Total Length';  
    sheetOut.getCell('B6').font = { bold: true };
    sheetOut.getCell('B6').alignment = { horizontal: 'center', vertical: 'middle' };
    sheetOut.getCell('B6').fill = headerFillStyle;

    sheetOut.getCell('C6').value = 'Roll/Drum';
    sheetOut.getCell('C6').font = { bold: true };
    sheetOut.getCell('C6').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    sheetOut.getCell('C6').fill = headerFillStyle;
    
    setBorders(currentRow);
    
    let colAWidth = 15;
    let colBWidth = 15;

    // Explicitly start appending array data from Row 7
    currentRow = 7;
    for (const rowVal of processedData) {
      const rOut = sheetOut.getRow(currentRow);
      rOut.getCell(1).value = rowVal[headers[0]];
      rOut.getCell(2).value = rowVal[headers[1]];
      rOut.getCell(3).value = rowVal['Roll/Drum'];
      
      rOut.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
      rOut.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
      rOut.getCell(3).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

      setBorders(currentRow);
      
      const lenA = String(rowVal[headers[0]] || '').length;
      const lenB = String(rowVal[headers[1]] || '').length;
      if (lenA > colAWidth) colAWidth = lenA;
      if (lenB > colBWidth) colBWidth = lenB;

      currentRow++;
    }

    // Process Summary Resume Logic Table Appended at Bottom
    currentRow += 4;
    const rollsSummary = {};
    for (const item of processedData) {
      const rNum = item['Roll/Drum'];
      if (rNum) {
        if (!rollsSummary[rNum]) rollsSummary[rNum] = { labels: [], total: 0 };
        rollsSummary[rNum].labels.push(item[headers[0]]);
        rollsSummary[rNum].total += item[headers[1]];
      }
    }

    const sHeaderRowOut = sheetOut.getRow(currentRow);
    const sumCellA = sHeaderRowOut.getCell(1);
    const sumCellB = sHeaderRowOut.getCell(2);
    const sumCellC = sHeaderRowOut.getCell(3);

    sumCellA.value = 'Nomor Roll/Drum';
    sumCellB.value = 'Total Length (m)';
    sumCellC.value = 'Detail String Label';

    const resumeHeaderFillStyle = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
    const headerAlignStyle = { horizontal: 'center', vertical: 'middle', wrapText: true };

    sumCellA.font = { bold: true }; sumCellA.fill = resumeHeaderFillStyle; sumCellA.alignment = headerAlignStyle;
    sumCellB.font = { bold: true }; sumCellB.fill = resumeHeaderFillStyle; sumCellB.alignment = headerAlignStyle;
    sumCellC.font = { bold: true }; sumCellC.fill = resumeHeaderFillStyle; sumCellC.alignment = headerAlignStyle;

    setBorders(currentRow);
    
    currentRow++;

    const sortedRolls = Object.keys(rollsSummary).sort((a,b) => parseInt(a) - parseInt(b));
    for (const rNum of sortedRolls) {
      const det = rollsSummary[rNum];
      const rOut = sheetOut.getRow(currentRow);
      
      const rCell1 = rOut.getCell(1);
      const rCell2 = rOut.getCell(2);
      const rCell3 = rOut.getCell(3);

      rCell1.value = parseInt(rNum);
      rCell2.value = det.total;
      rCell3.value = det.labels.join(", ");

      rCell1.alignment = { horizontal: 'center', vertical: 'middle' };
      rCell2.alignment = { horizontal: 'center', vertical: 'middle' };
      rCell3.alignment = { vertical: 'middle', wrapText: true };

      setBorders(currentRow);
      currentRow++;
    }

    // Fix column widths matching PHP autofit restrictions
    sheetOut.getColumn(1).width = Math.min(colAWidth + 5, 40);
    sheetOut.getColumn(2).width = Math.min(colBWidth + 5, 40);
    sheetOut.getColumn(3).width = 60; 
  }

  if (globalSummaryData.length === 0) {
    throw new Error('Data tidak ditemukan. Pastikan file Excel input memiliki format baris yang sesuai.');
  }

  if (imageId !== null) {
    summarySheet.addImage(imageId, {
      tl: { col: 0, row: 1 }, // A2
      ext: { width: 150, height: 50 }
    });
  }

  // --- Populate Summary Data Tab ---
  let row = 6;
  const sHeaderRow = summarySheet.getRow(row);
  sHeaderRow.getCell(1).value = "Sheet Name";
  sHeaderRow.getCell(2).value = "For";
  sHeaderRow.getCell(3).value = "Total Roll/Drum";
  sHeaderRow.font = { bold: true, name: 'Calibri' };
  sHeaderRow.alignment = { horizontal: 'center' };
  ['A','B','C'].forEach(c => {
    const cell = summarySheet.getCell(`${c}${row}`);
    cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
  });

  row++;
  let grandTotal = 0;
  
  globalSummaryData.sort((a,b) => a.sheet_name.localeCompare(b.sheet_name));

  for (const item of globalSummaryData) {
    grandTotal += item.total_roll;
    let type = "DC Cable";
    if (item.sheet_name.includes("(+)") || item.sheet_name.includes("Positif")) type = "Positif";
    else if (item.sheet_name.includes("(-)") || item.sheet_name.includes("Negatif")) type = "Negatif";
    else if (item.sheet_name.includes("HR") || item.sheet_name.includes("Homerun")) type = "Homerun";

    const cRow = summarySheet.getRow(row);
    cRow.getCell(1).value = item.sheet_name;
    cRow.getCell(2).value = type;
    cRow.getCell(3).value = item.total_roll;
    
    ['A','B','C'].forEach(c => {
        summarySheet.getCell(`${c}${row}`).border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
    cRow.getCell(3).alignment = { horizontal: 'center' };
    row++;
  }

  const gtRow = summarySheet.getRow(row);
  summarySheet.mergeCells(`A${row}:B${row}`);
  gtRow.getCell(1).value = "Grand Total Roll/Drum";
  gtRow.getCell(1).alignment = { horizontal: 'center' };
  gtRow.getCell(1).font = { bold: true };
  gtRow.getCell(1).border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };

  gtRow.getCell(3).value = grandTotal;
  gtRow.getCell(3).font = { bold: true };
  gtRow.getCell(3).alignment = { horizontal: 'center' };
  gtRow.getCell(3).border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };

  summarySheet.getColumn(1).width = 25;
  summarySheet.getColumn(2).width = 15;
  summarySheet.getColumn(3).width = 20;

  // Generate buffer
  return await workbookOutput.xlsx.writeBuffer();
};
