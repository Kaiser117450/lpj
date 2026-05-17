const ExcelJS = require("exceljs");

const BRAND_PRIMARY = "FFD4A017";
const BRAND_DARK = "FF8A6810";
const ZEBRA = "FFFAF7F0";
const TEXT_DARK = "FF1F2937";

function safeStr(v) { return (v == null ? "" : String(v)).trim(); }

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const body = req.body || {};
    const entries = Array.isArray(body.entries) ? body.entries : [];
    const title = safeStr(body.title) || "BUKTI NOTA LPJ";
    const subtitle = safeStr(body.subtitle) || "Ayam Guling Enakko";

    if (!entries.length) {
      return res.status(400).json({ error: "Tidak ada entri untuk diekspor" });
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = "AGE BALI — LPJ Report Builder";
    wb.created = new Date();
    wb.title = title;

    const ws = wb.addWorksheet("LPJ Report", {
      pageSetup: {
        paperSize: 9, // A4
        orientation: "portrait",
        fitToPage: true, fitToWidth: 1, fitToHeight: 0,
        margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 },
      },
      views: [{ state: "frozen", ySplit: 5 }],
    });

    // Title rows
    ws.mergeCells("A1:F1");
    const tCell = ws.getCell("A1");
    tCell.value = title;
    tCell.font = { name: "Arial", size: 18, bold: true, color: { argb: TEXT_DARK } };
    tCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(1).height = 30;

    ws.mergeCells("A2:F2");
    const sCell = ws.getCell("A2");
    sCell.value = subtitle;
    sCell.font = { name: "Arial", size: 12, italic: true, color: { argb: BRAND_DARK } };
    sCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(2).height = 20;

    ws.mergeCells("A3:F3");
    const dCell = ws.getCell("A3");
    dCell.value = `Dicetak: ${new Date().toLocaleString("id-ID", {
      weekday: "long", year: "numeric", month: "long", day: "numeric",
      hour: "2-digit", minute: "2-digit",
    })}`;
    dCell.font = { name: "Arial", size: 10, color: { argb: "FF6B7280" } };
    dCell.alignment = { horizontal: "center", vertical: "middle" };
    ws.getRow(3).height = 18;

    ws.getRow(4).height = 8; // spacer

    // Column setup
    ws.columns = [
      { key: "no", width: 6 },
      { key: "kode", width: 22 },
      { key: "jenis", width: 32 },
      { key: "tanggal", width: 18 },
      { key: "transfer", width: 14 },
      { key: "nota", width: 14 },
    ];

    // Header row at row 5
    const headerRow = ws.getRow(5);
    headerRow.values = ["NO", "KODE", "JENIS PEMBELIAN", "TANGGAL", "JML TRANSFER", "JML NOTA"];
    headerRow.height = 26;
    headerRow.eachCell((cell) => {
      cell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BRAND_PRIMARY } };
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = {
        top: { style: "thin", color: { argb: BRAND_DARK } },
        bottom: { style: "medium", color: { argb: BRAND_DARK } },
        left: { style: "thin", color: { argb: BRAND_DARK } },
        right: { style: "thin", color: { argb: BRAND_DARK } },
      };
    });

    // Data rows starting row 6
    entries.forEach((e, i) => {
      const rowNum = 6 + i;
      const row = ws.getRow(rowNum);
      row.values = [
        i + 1,
        safeStr(e.kode),
        safeStr(e.jenis) || "—",
        safeStr(e.tanggal) || "—",
        (e.transfer && e.transfer.length) || 0,
        (e.nota && e.nota.length) || 0,
      ];
      row.height = 22;
      const zebra = i % 2 === 1;
      row.eachCell((cell, colNum) => {
        cell.font = { name: "Arial", size: 10, color: { argb: TEXT_DARK } };
        if (zebra) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: ZEBRA } };
        }
        cell.border = {
          top: { style: "hair", color: { argb: "FFD1D5DB" } },
          bottom: { style: "hair", color: { argb: "FFD1D5DB" } },
          left: { style: "hair", color: { argb: "FFD1D5DB" } },
          right: { style: "hair", color: { argb: "FFD1D5DB" } },
        };
        if (colNum === 1 || colNum === 4 || colNum === 5 || colNum === 6) {
          cell.alignment = { horizontal: "center", vertical: "middle" };
        } else {
          cell.alignment = { horizontal: "left", vertical: "middle", indent: 1 };
        }
        if (colNum === 2) {
          cell.font = { name: "Consolas", size: 9, color: { argb: "FF6B7280" } };
        }
      });
    });

    // Totals row
    const totalRowNum = 6 + entries.length;
    const totalTransfer = entries.reduce((s, e) => s + ((e.transfer && e.transfer.length) || 0), 0);
    const totalNota = entries.reduce((s, e) => s + ((e.nota && e.nota.length) || 0), 0);
    const totRow = ws.getRow(totalRowNum);
    totRow.values = ["", "", `TOTAL (${entries.length} entri)`, "", totalTransfer, totalNota];
    ws.mergeCells(totalRowNum, 1, totalRowNum, 4);
    totRow.height = 26;
    totRow.eachCell((cell) => {
      cell.font = { name: "Arial", size: 11, bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: BRAND_DARK } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.border = {
        top: { style: "medium", color: { argb: BRAND_DARK } },
        bottom: { style: "medium", color: { argb: BRAND_DARK } },
        left: { style: "thin", color: { argb: BRAND_DARK } },
        right: { style: "thin", color: { argb: BRAND_DARK } },
      };
    });
    // Right-align label cell
    ws.getCell(totalRowNum, 1).alignment = { horizontal: "right", vertical: "middle", indent: 1 };

    // Footer info
    const footerRow = totalRowNum + 2;
    ws.mergeCells(footerRow, 1, footerRow, 6);
    const fCell = ws.getCell(footerRow, 1);
    fCell.value = "AGE BALI — Ayam Guling Enakko • Dicetak otomatis oleh LPJ Report Builder";
    fCell.font = { name: "Arial", size: 9, italic: true, color: { argb: "FF6B7280" } };
    fCell.alignment = { horizontal: "center", vertical: "middle" };

    const buffer = await wb.xlsx.writeBuffer();
    const safeName = title.replace(/[^\w\s-]/g, "").replace(/\s+/g, "_") || "LPJ_Report";

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName}.xlsx"`);
    res.setHeader("Cache-Control", "no-store");
    res.status(200).send(Buffer.from(buffer));
  } catch (err) {
    console.error("XLSX export error:", err);
    res.status(500).json({ error: err.message || "Internal server error" });
  }
};

module.exports.config = {
  api: {
    bodyParser: { sizeLimit: "20mb" },
  },
};
