/**
 * Pure Offline Native OOXML (.xlsx) Generator
 * Exports the exact Arkan Group Tour Expenses workbook with:
 * - Merged headers & per-day vertical merges (Cols I-P)
 * - Live Excel formulas (=H4-F4+1, =SUM(...)) + cached values
 * - Exact Calibri styling, thin black borders, date formats, landscape A4 setup
 * - Embedded Prepared by signature PNG image
 */
(function () {
  'use strict';

  // ---- CRC-32 for uncompressed ZIP (Store method) ----
  const CRC_TABLE = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    CRC_TABLE[n] = c >>> 0;
  }

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function strToUtf8(str) {
    return new TextEncoder().encode(str);
  }

  function buildZip(entries) {
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    for (const entry of entries) {
      const nameBytes = strToUtf8(entry.name);
      const dataBytes = entry.data instanceof Uint8Array ? entry.data : strToUtf8(entry.data);
      const crc = crc32(dataBytes);
      const size = dataBytes.length;

      // Local file header (30 bytes + name)
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(localHeader.buffer);
      lv.setUint32(0, 0x04034b50, true); // signature
      lv.setUint16(4, 20, true);         // version needed
      lv.setUint16(6, 0x0800, true);     // UTF-8 flag
      lv.setUint16(8, 0, true);          // method: 0 (store)
      lv.setUint16(10, 0, true);         // mod time
      lv.setUint16(12, 0x5421, true);    // mod date
      lv.setUint32(14, crc, true);
      lv.setUint32(18, size, true);
      lv.setUint32(22, size, true);
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);
      localHeader.set(nameBytes, 30);

      localParts.push(localHeader, dataBytes);

      // Central directory header (46 bytes + name)
      const cdHeader = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(cdHeader.buffer);
      cv.setUint32(0, 0x02014b50, true); // signature
      cv.setUint16(4, 20, true);         // version made by
      cv.setUint16(6, 20, true);         // version needed
      cv.setUint16(8, 0x0800, true);     // UTF-8 flag
      cv.setUint16(10, 0, true);         // method: store
      cv.setUint16(12, 0, true);
      cv.setUint16(14, 0x5421, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, size, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);
      cv.setUint16(32, 0, true);
      cv.setUint16(34, 0, true);
      cv.setUint16(36, 0, true);
      cv.setUint32(38, 0, true);
      cv.setUint32(42, offset, true);
      cdHeader.set(nameBytes, 46);

      centralParts.push(cdHeader);
      offset += localHeader.length + dataBytes.length;
    }

    const centralOffset = offset;
    let centralSize = 0;
    for (const cp of centralParts) centralSize += cp.length;

    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, entries.length, true);
    ev.setUint16(10, entries.length, true);
    ev.setUint32(12, centralSize, true);
    ev.setUint32(16, centralOffset, true);
    ev.setUint16(20, 0, true);

    const totalLen = offset + centralSize + eocd.length;
    const out = new Uint8Array(totalLen);
    let pos = 0;
    for (const p of localParts) { out.set(p, pos); pos += p.length; }
    for (const p of centralParts) { out.set(p, pos); pos += p.length; }
    out.set(eocd, pos);
    return out;
  }

  // ---- Helpers ----
  function escXml(val) {
    return String(val ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function isoToExcelSerial(iso) {
    if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(iso)) return null;
    const [y, m, d] = iso.split('-').map(Number);
    const utc = Date.UTC(y, m - 1, d);
    const epoch = Date.UTC(1899, 11, 30);
    return Math.round((utc - epoch) / 86400000);
  }

  function numVal(v) {
    const n = Number(v);
    return Number.isFinite(n) && v !== '' && v !== null && v !== undefined ? n : 0;
  }

  function strCell(ref, text, style) {
    if (text === '' || text === null || text === undefined) {
      return `<c r="${ref}" s="${style}"/>`;
    }
    return `<c r="${ref}" t="inlineStr" s="${style}"><is><t xml:space="preserve">${escXml(text)}</t></is></c>`;
  }

  function numCell(ref, val, style, showZero) {
    if ((val === '' || val === null || val === undefined || Number(val) === 0) && !showZero) {
      return `<c r="${ref}" s="${style}"/>`;
    }
    return `<c r="${ref}" s="${style}"><v>${Number(val)}</v></c>`;
  }

  function formulaCell(ref, formula, cachedVal, style) {
    return `<c r="${ref}" s="${style}"><f>${escXml(formula)}</f><v>${cachedVal}</v></c>`;
  }

  // Render signature (+ optional signature date) into a crisp transparent PNG Uint8Array
  async function renderSignaturePngBytes(sigDataUrl, showDate, dateText, blendWhite) {
    if (!sigDataUrl) return null;
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => {
        try {
          const canvas = document.createElement('canvas');
          const sigW = 240;
          const sigH = 74;
          const extraDateW = (showDate && dateText) ? 140 : 0;
          canvas.width = sigW + extraDateW;
          canvas.height = sigH;
          const ctx = canvas.getContext('2d');
          ctx.clearRect(0, 0, canvas.width, canvas.height);

          // Draw signature aspect-fitted into sigW x sigH
          const ratio = Math.min(sigW / img.width, sigH / img.height);
          const drawW = img.width * ratio;
          const drawH = img.height * ratio;
          const dx = (sigW - drawW) / 2;
          const dy = sigH - drawH;
          ctx.drawImage(img, dx, dy, drawW, drawH);

          // If blendWhite is enabled, convert near-white pixels to transparent
          if (blendWhite) {
            const imgData = ctx.getImageData(0, 0, sigW, sigH);
            const d = imgData.data;
            for (let i = 0; i < d.length; i += 4) {
              const lum = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
              if (lum > 235) {
                d[i + 3] = 0;
              } else if (lum > 180) {
                d[i + 3] = Math.round(d[i + 3] * ((235 - lum) / 55));
              }
            }
            ctx.putImageData(imgData, 0, 0);
          }

          if (showDate && dateText) {
            ctx.fillStyle = '#111111';
            ctx.font = 'italic 20px "Segoe Script", "Comic Sans MS", Calibri, sans-serif';
            ctx.textBaseline = 'bottom';
            ctx.fillText(dateText, sigW + 6, sigH - 8);
          }

          const dataUrl = canvas.toDataURL('image/png');
          const b64 = dataUrl.split(',')[1];
          const bin = atob(b64);
          const bytes = new Uint8Array(bin.length);
          for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
          resolve({ bytes, width: canvas.width, height: canvas.height });
        } catch (err) {
          resolve(null);
        }
      };
      img.onerror = () => resolve(null);
      img.src = sigDataUrl;
    });
  }

  function buildStylesXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="d\\-m\\-yyyy"/>
    <numFmt numFmtId="165" formatCode="d/m/yyyy"/>
  </numFmts>
  <fonts count="5">
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><b/><sz val="14"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><sz val="12"/><color rgb="FF000000"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FF000000"/></left>
      <right style="thin"><color rgb="FF000000"/></right>
      <top style="thin"><color rgb="FF000000"/></top>
      <bottom style="thin"><color rgb="FF000000"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="13">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="164" fontId="3" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="164" fontId="3" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
</styleSheet>`;
  }

  async function generateXlsxUint8Array(payload) {
    const {
      companyName,
      reportTitle,
      destination,
      purpose,
      designation,
      startDate,
      endDate,
      days,
      signatureLabels,
      signatureDataUrl,
      showSigDate,
      sigDate,
      sigBlendMultiply
    } = payload;

    const startSerial = isoToExcelSerial(startDate);
    const endSerial = isoToExcelSerial(endDate);
    const totalDaysVal = (startSerial && endSerial && endSerial >= startSerial)
      ? (endSerial - startSerial + 1)
      : 0;

    const merges = [
      'A1:P1',
      'A2:P2',
      'A6:A7',
      'B6:C6',
      'D6:I6',
      'J6:N6'
    ];

    const rowsXml = [];

    // Row 1: Company Name
    rowsXml.push(`<row r="1" ht="19.5" customHeight="1">${strCell('A1', companyName, 1)}</row>`);
    // Row 2: Subtitle
    rowsXml.push(`<row r="2" ht="16.5" customHeight="1">${strCell('A2', reportTitle, 2)}</row>`);
    // Row 3: Blank
    rowsXml.push(`<row r="3" ht="10" customHeight="1"/>`);
    // Row 4: Destination & Duration
    let r4 = strCell('A4', `Destination: ${destination}`, 3);
    r4 += strCell('E4', 'Duration:', 3);
    r4 += startSerial ? `<c r="F4" s="5"><v>${startSerial}</v></c>` : strCell('F4', startDate, 3);
    r4 += strCell('G4', 'to', 4);
    r4 += endSerial ? `<c r="H4" s="6"><v>${endSerial}</v></c>` : strCell('H4', endDate, 3);
    rowsXml.push(`<row r="4" ht="16" customHeight="1">${r4}</row>`);

    // Row 5: Purpose, Total Day, Travellers Designation
    let r5 = strCell('A5', `Purpose: ${purpose}`, 3);
    r5 += strCell('E5', 'Total Day:', 3);
    r5 += (startSerial && endSerial)
      ? formulaCell('F5', 'H4-F4+1', totalDaysVal, 3)
      : numCell('F5', totalDaysVal, 3, true);
    r5 += strCell('I5', `Travellers Designation: ${designation}`, 3);
    rowsXml.push(`<row r="5" ht="16" customHeight="1">${r5}</row>`);

    // Row 6: Table Header Row 1
    let r6 = '';
    r6 += strCell('A6', 'Date', 7);
    r6 += strCell('B6', 'Destination', 7) + strCell('C6', '', 7);
    r6 += strCell('D6', 'Transportation Expenses', 7)
      + strCell('E6', '', 7) + strCell('F6', '', 7)
      + strCell('G6', '', 7) + strCell('H6', '', 7) + strCell('I6', '', 7);
    r6 += strCell('J6', 'Meal Expenses', 7)
      + strCell('K6', '', 7) + strCell('L6', '', 7)
      + strCell('M6', '', 7) + strCell('N6', '', 7);
    r6 += strCell('O6', 'Living Cost', 7);
    r6 += strCell('P6', 'Grand Total', 7);
    rowsXml.push(`<row r="6" ht="26" customHeight="1">${r6}</row>`);

    // Row 7: Table Header Row 2
    const h2 = [
      ['A', ''], ['B', 'From'], ['C', 'To'],
      ['D', 'Bus'], ['E', 'Metro/ Rail '], ['F', 'Auto'], ['G', 'Bike'], ['H', 'Rickshaw'], ['I', 'Transport Total'],
      ['J', 'Breakfast'], ['K', 'Lunch'], ['L', 'Dinner'], ['M', 'Tea Break'], ['N', 'Meal Total'],
      ['O', 'Hotel'], ['P', 'Total Day Expense']
    ];
    let r7 = h2.map(([col, txt]) => strCell(`${col}7`, txt, 7)).join('');
    rowsXml.push(`<row r="7" ht="28" customHeight="1">${r7}</row>`);

    // Data Rows starting at Row 8
    let curRow = 8;
    const sumCols = { D: 0, E: 0, F: 0, G: 0, H: 0, I: 0, J: 0, K: 0, L: 0, M: 0, N: 0, O: 0, P: 0 };

    const safeDays = (days && days.length) ? days : [{
      date: startDate || '',
      routes: [{ from: '', to: '', bus: '', metro: '', auto: '', bike: '', rickshaw: '' }],
      breakfast: '', lunch: '', dinner: '', teaBreak: '', hotel: ''
    }];

    for (const day of safeDays) {
      const routes = (day.routes && day.routes.length)
        ? day.routes
        : [{ from: '', to: '', bus: '', metro: '', auto: '', bike: '', rickshaw: '' }];
      const startR = curRow;
      const endR = curRow + routes.length - 1;

      let dayTrans = 0;
      for (const rt of routes) {
        const b = numVal(rt.bus), m = numVal(rt.metro), a = numVal(rt.auto), bk = numVal(rt.bike), rk = numVal(rt.rickshaw);
        sumCols.D += b; sumCols.E += m; sumCols.F += a; sumCols.G += bk; sumCols.H += rk;
        dayTrans += b + m + a + bk + rk;
      }

      const bf = numVal(day.breakfast), ln = numVal(day.lunch), dn = numVal(day.dinner), tb = numVal(day.teaBreak);
      const dayMeal = bf + ln + dn + tb;
      const dayHotel = numVal(day.hotel);
      const dayGrand = dayTrans + dayMeal + dayHotel;

      sumCols.I += dayTrans;
      sumCols.J += bf; sumCols.K += ln; sumCols.L += dn; sumCols.M += tb;
      sumCols.N += dayMeal;
      sumCols.O += dayHotel;
      sumCols.P += dayGrand;

      const daySerial = isoToExcelSerial(day.date);

      routes.forEach((rt, idx) => {
        const r = startR + idx;
        let rowCells = '';

        // Col A: Date on first route row only
        if (idx === 0) {
          rowCells += daySerial
            ? `<c r="A${r}" s="8"><v>${daySerial}</v></c>`
            : strCell(`A${r}`, day.date || '', 8);
        } else {
          rowCells += `<c r="A${r}" s="8"/>`;
        }

        // Col B & C: From & To
        rowCells += strCell(`B${r}`, rt.from || '', 9);
        rowCells += strCell(`C${r}`, rt.to || '', 9);

        // Cols D..H: Transport modes
        rowCells += numCell(`D${r}`, rt.bus, 10, false);
        rowCells += numCell(`E${r}`, rt.metro, 10, false);
        rowCells += numCell(`F${r}`, rt.auto, 10, false);
        rowCells += numCell(`G${r}`, rt.bike, 10, false);
        rowCells += numCell(`H${r}`, rt.rickshaw, 10, false);

        // Cols I..P: First row holds values/formulas, sub-rows empty bordered cells (merged vertically)
        if (idx === 0) {
          rowCells += formulaCell(`I${r}`, `SUM(D${startR}:H${endR})`, dayTrans, 10);
          rowCells += numCell(`J${r}`, day.breakfast, 10, false);
          rowCells += numCell(`K${r}`, day.lunch, 10, false);
          rowCells += numCell(`L${r}`, day.dinner, 10, false);
          rowCells += numCell(`M${r}`, day.teaBreak, 10, false);
          rowCells += formulaCell(`N${r}`, `SUM(J${startR}:M${startR})`, dayMeal, 10);
          rowCells += numCell(`O${r}`, day.hotel, 10, false);
          rowCells += formulaCell(`P${r}`, `SUM(I${startR},N${startR},O${startR})`, dayGrand, 10);
        } else {
          for (const col of ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']) {
            rowCells += `<c r="${col}${r}" s="10"/>`;
          }
        }

        rowsXml.push(`<row r="${r}" ht="15.75" customHeight="1">${rowCells}</row>`);
      });

      if (endR > startR) {
        for (const col of ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']) {
          merges.push(`${col}${startR}:${col}${endR}`);
        }
      }

      curRow = endR + 1;
    }

    const lastDataRow = curRow - 1;
    const totalRow = curRow;

    // Total Row
    let totCells = strCell(`A${totalRow}`, 'Total', 11)
      + `<c r="B${totalRow}" s="11"/>`
      + `<c r="C${totalRow}" s="11"/>`;
    for (const col of ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P']) {
      totCells += formulaCell(`${col}${totalRow}`, `SUM(${col}8:${col}${lastDataRow})`, sumCols[col], 11);
    }
    rowsXml.push(`<row r="${totalRow}" ht="18" customHeight="1">${totCells}</row>`);

    // 3 Blank spacing rows for signature
    rowsXml.push(`<row r="${totalRow + 1}" ht="15" customHeight="1"/>`);
    rowsXml.push(`<row r="${totalRow + 2}" ht="15" customHeight="1"/>`);
    rowsXml.push(`<row r="${totalRow + 3}" ht="15" customHeight="1"/>`);

    // Signature Labels Row (Row totalRow + 4, matching original Row 115)
    const sigRow = totalRow + 4;
    const labels = signatureLabels || ['Prepared by', 'Dept. Head', 'Audit Dept.', 'Accounts', 'GM (Operation)', 'Director'];
    const sigCols = ['B', 'D', 'F', 'H', 'J', 'N'];
    let sigCells = sigCols.map((col, i) => strCell(`${col}${sigRow}`, labels[i] || '', 12)).join('');
    rowsXml.push(`<row r="${sigRow}" ht="16" customHeight="1">${sigCells}</row>`);

    // Optional embedded signature image
    const sigPng = await renderSignaturePngBytes(signatureDataUrl, showSigDate, sigDate, sigBlendMultiply);
    const hasDrawing = Boolean(sigPng && sigPng.bytes && sigPng.bytes.length > 0);

    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cols>
    <col min="1" max="1" width="11.5" customWidth="1"/>
    <col min="2" max="3" width="16.5" customWidth="1"/>
    <col min="4" max="5" width="9" customWidth="1"/>
    <col min="6" max="6" width="10" customWidth="1"/>
    <col min="7" max="7" width="8" customWidth="1"/>
    <col min="8" max="8" width="10.5" customWidth="1"/>
    <col min="9" max="9" width="11.5" customWidth="1"/>
    <col min="10" max="14" width="9.5" customWidth="1"/>
    <col min="15" max="15" width="9" customWidth="1"/>
    <col min="16" max="16" width="11.5" customWidth="1"/>
  </cols>
  <sheetData>
    ${rowsXml.join('\n    ')}
  </sheetData>
  <mergeCells count="${merges.length}">
    ${merges.map(m => `<mergeCell ref="${m}"/>`).join('\n    ')}
  </mergeCells>
  <pageMargins left="0.5" right="0.5" top="0.55" bottom="0.55" header="0.25" footer="0.25"/>
  <pageSetup paperSize="9" scale="80" orientation="landscape"/>
  ${hasDrawing ? '<drawing r:id="rId1"/>' : ''}
</worksheet>`;

    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${hasDrawing ? '<Default Extension="png" ContentType="image/png"/>' : ''}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.styles+xml"/>
  ${hasDrawing ? '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' : ''}
</Types>`;

    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Tour Expenses" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

    const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

    const zipEntries = [
      { name: '[Content_Types].xml', data: contentTypes },
      { name: '_rels/.rels', data: rootRels },
      { name: 'xl/workbook.xml', data: workbookXml },
      { name: 'xl/_rels/workbook.xml.rels', data: workbookRels },
      { name: 'xl/styles.xml', data: buildStylesXml() },
      { name: 'xl/worksheets/sheet1.xml', data: sheetXml }
    ];

    if (hasDrawing) {
      // Zero-indexed row anchor right above sigRow (totalRow is 1-based, so 0-based row = totalRow)
      const fromRow0 = totalRow;
      const toRow0 = sigRow - 1;
      const toCol0 = (showSigDate && sigDate) ? 2 : 1;

      const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;

      const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/signature.png"/>
</Relationships>`;

      const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from>
      <xdr:col>0</xdr:col>
      <xdr:colOff>450000</xdr:colOff>
      <xdr:row>${fromRow0}</xdr:row>
      <xdr:rowOff>60000</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>${toCol0}</xdr:col>
      <xdr:colOff>900000</xdr:colOff>
      <xdr:row>${toRow0}</xdr:row>
      <xdr:rowOff>120000</xdr:rowOff>
    </xdr:to>
    <xdr:pic>
      <xdr:nvPicPr>
        <xdr:cNvPr id="2" name="PreparedBySignature"/>
        <xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr>
      </xdr:nvPicPr>
      <xdr:blipFill>
        <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1"/>
        <a:stretch><a:fillRect/></a:stretch>
      </xdr:blipFill>
      <xdr:spPr>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
      </xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>`;

      zipEntries.push(
        { name: 'xl/worksheets/_rels/sheet1.xml.rels', data: sheetRels },
        { name: 'xl/drawings/drawing1.xml', data: drawingXml },
        { name: 'xl/drawings/_rels/drawing1.xml.rels', data: drawingRels },
        { name: 'xl/media/signature.png', data: sigPng.bytes }
      );
    }

    return buildZip(zipEntries);
  }

  window.TourXlsxExporter = {
    async download(payload, filename) {
      const bytes = await generateXlsxUint8Array(payload);
      const blob = new Blob([bytes], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename || 'Tour_Expenses.xlsx';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 200);
    },
    generateBytes: generateXlsxUint8Array
  };
})();
