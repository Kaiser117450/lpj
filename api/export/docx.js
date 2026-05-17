const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, PageNumber, VerticalAlign,
} = require("docx");

const BRAND_PRIMARY = "D4A017";
const BRAND_DARK = "8A6810";
const TEXT_DARK = "1F2937";
const TEXT_MUTED = "6B7280";
const ROW_ZEBRA = "FAF7F0";

function b64toBuf(dataUrl) {
  if (!dataUrl || typeof dataUrl !== "string") return null;
  const matches = /^data:image\/([a-zA-Z0-9+.-]+);base64,(.*)$/.exec(dataUrl);
  if (!matches) return null;
  const ext = matches[1].toLowerCase() === "jpeg" ? "jpg" : matches[1].toLowerCase();
  return { buf: Buffer.from(matches[2], "base64"), ext };
}

function imageType(ext) {
  if (ext === "jpg" || ext === "jpeg") return "jpg";
  if (ext === "png") return "png";
  if (ext === "gif") return "gif";
  if (ext === "bmp") return "bmp";
  return "jpg";
}

function makeBorder(color = "D1D5DB") {
  return { style: BorderStyle.SINGLE, size: 6, color };
}

function makeBorders(color) {
  const b = makeBorder(color);
  return { top: b, bottom: b, left: b, right: b };
}

function buildImageParagraphs(images, label) {
  const result = [];
  if (!images || !images.length) {
    result.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "— tidak ada —", italics: true, color: TEXT_MUTED, font: "Arial", size: 18 })],
    }));
    return result;
  }

  for (let i = 0; i < images.length; i++) {
    const img = images[i];
    const parsed = b64toBuf(img.url);
    if (!parsed) continue;
    result.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: i === 0 ? 0 : 120, after: 60 },
      children: [
        new ImageRun({
          type: imageType(parsed.ext),
          data: parsed.buf,
          transformation: { width: 200, height: 260 },
          altText: { title: label, description: `${label} ${i + 1}`, name: label.toLowerCase() },
        }),
      ],
    }));
  }

  if (!result.length) {
    result.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "— gambar tidak valid —", italics: true, color: TEXT_MUTED, font: "Arial", size: 18 })],
    }));
  }
  return result;
}

function makeCell({ children, width, shading, borders, valign }) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    margins: { top: 120, bottom: 120, left: 140, right: 140 },
    borders: borders || makeBorders(),
    shading,
    verticalAlign: valign || VerticalAlign.CENTER,
    children,
  });
}

function formatDateID() {
  const d = new Date();
  return d.toLocaleDateString("id-ID", {
    weekday: "long", year: "numeric", month: "long", day: "numeric",
  });
}

module.exports = async function handler(req, res) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const body = req.body || {};
    const entries = Array.isArray(body.entries) ? body.entries : [];
    const title = (body.title || "BUKTI NOTA LPJ").toString().trim();
    const subtitle = (body.subtitle || "Ayam Guling Enakko").toString().trim();

    if (!entries.length) {
      return res.status(400).json({ error: "Tidak ada entri untuk diekspor" });
    }

    // Column widths in DXA (1 inch = 1440 DXA). Total ≈ 9360 DXA (~6.5")
    const COL_NO = 700;
    const COL_JENIS = 2400;
    const COL_TRANSFER = 3100;
    const COL_NOTA = 3160;
    const COL_W = [COL_NO, COL_JENIS, COL_TRANSFER, COL_NOTA];
    const TABLE_W = COL_W.reduce((a, b) => a + b, 0);

    const rows = [];

    // Header row
    rows.push(new TableRow({
      tableHeader: true,
      children: ["NO", "JENIS PEMBELIAN", "TGL & BUKTI TRANSFER", "BUKTI NOTA"].map((label, i) =>
        makeCell({
          width: COL_W[i],
          shading: { fill: BRAND_PRIMARY, type: ShadingType.CLEAR, color: "auto" },
          borders: makeBorders(BRAND_DARK),
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: label, bold: true, font: "Arial", size: 22, color: "FFFFFF",
            })],
          })],
        })
      ),
    }));

    // Data rows
    entries.forEach((entry, idx) => {
      const isZebra = idx % 2 === 1;
      const rowShading = isZebra
        ? { fill: ROW_ZEBRA, type: ShadingType.CLEAR, color: "auto" }
        : undefined;

      const transferChildren = [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [
            new TextRun({
              text: entry.tanggal || "—",
              bold: true, font: "Arial", size: 22, color: TEXT_DARK,
            }),
          ],
        }),
        ...buildImageParagraphs(entry.transfer, "Transfer"),
      ];

      const notaChildren = buildImageParagraphs(entry.nota, "Nota");

      rows.push(new TableRow({
        cantSplit: false,
        children: [
          makeCell({
            width: COL_NO,
            shading: rowShading,
            valign: VerticalAlign.CENTER,
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({
                text: String(idx + 1), bold: true, font: "Arial", size: 24, color: BRAND_DARK,
              })],
            })],
          }),
          makeCell({
            width: COL_JENIS,
            shading: rowShading,
            valign: VerticalAlign.CENTER,
            children: [
              new Paragraph({
                children: [new TextRun({
                  text: entry.jenis || "—",
                  bold: true, font: "Arial", size: 22, color: TEXT_DARK,
                })],
              }),
              new Paragraph({
                spacing: { before: 100 },
                children: [new TextRun({
                  text: entry.kode || "",
                  font: "Consolas", size: 16, color: TEXT_MUTED,
                })],
              }),
            ],
          }),
          makeCell({
            width: COL_TRANSFER,
            shading: rowShading,
            children: transferChildren,
          }),
          makeCell({
            width: COL_NOTA,
            shading: rowShading,
            children: notaChildren,
          }),
        ],
      }));
    });

    // Totals row
    const totalTransfer = entries.reduce((s, e) => s + (e.transfer?.length || 0), 0);
    const totalNota = entries.reduce((s, e) => s + (e.nota?.length || 0), 0);
    const darkShading = { fill: BRAND_DARK, type: ShadingType.CLEAR, color: "auto" };

    rows.push(new TableRow({
      children: [
        new TableCell({
          columnSpan: 2,
          width: { size: COL_NO + COL_JENIS, type: WidthType.DXA },
          margins: { top: 120, bottom: 120, left: 140, right: 140 },
          borders: makeBorders(BRAND_DARK),
          shading: darkShading,
          verticalAlign: VerticalAlign.CENTER,
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: `TOTAL (${entries.length} entri)`,
              bold: true, font: "Arial", size: 22, color: "FFFFFF",
            })],
          })],
        }),
        makeCell({
          width: COL_TRANSFER,
          shading: darkShading,
          borders: makeBorders(BRAND_DARK),
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: `${totalTransfer} bukti transfer`,
              bold: true, font: "Arial", size: 22, color: "FFFFFF",
            })],
          })],
        }),
        makeCell({
          width: COL_NOTA,
          shading: darkShading,
          borders: makeBorders(BRAND_DARK),
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: `${totalNota} bukti nota`,
              bold: true, font: "Arial", size: 22, color: "FFFFFF",
            })],
          })],
        }),
      ],
    }));

    const doc = new Document({
      creator: "AGE BALI — LPJ Report Builder",
      title,
      description: "Laporan Pertanggungjawaban Bukti Nota & Transfer",
      styles: {
        default: { document: { run: { font: "Arial", size: 22 } } },
      },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 },
          },
        },
        headers: {
          default: new Header({
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              border: { bottom: { color: BRAND_PRIMARY, style: BorderStyle.SINGLE, size: 12 } },
              spacing: { after: 100 },
              children: [
                new TextRun({ text: "AYAM GULING ENAKKO ", bold: true, font: "Arial", size: 26, color: BRAND_PRIMARY }),
                new TextRun({ text: "• AGE BALI", bold: true, font: "Arial", size: 20, color: TEXT_MUTED }),
              ],
            })],
          }),
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({ text: "Halaman ", font: "Arial", size: 16, color: TEXT_MUTED }),
                new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: TEXT_MUTED, bold: true }),
                new TextRun({ text: " dari ", font: "Arial", size: 16, color: TEXT_MUTED }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: TEXT_MUTED, bold: true }),
                new TextRun({ text: "  •  Dicetak otomatis oleh LPJ Report Builder", font: "Arial", size: 14, color: TEXT_MUTED, italics: true }),
              ],
            })],
          }),
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [new TextRun({
              text: title, bold: true, font: "Arial", size: 32, color: TEXT_DARK,
            })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 80 },
            children: [new TextRun({
              text: subtitle, font: "Arial", size: 22, color: BRAND_DARK, italics: true,
            })],
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 320 },
            children: [new TextRun({
              text: `Dicetak: ${formatDateID()}`,
              font: "Arial", size: 18, color: TEXT_MUTED,
            })],
          }),
          new Table({
            width: { size: TABLE_W, type: WidthType.DXA },
            columnWidths: COL_W,
            rows,
          }),
          new Paragraph({
            spacing: { before: 320 },
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: "______________________________",
              font: "Arial", size: 20, color: TEXT_MUTED,
            })],
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({
              text: "Manajer Operasional",
              font: "Arial", size: 20, bold: true, color: TEXT_DARK,
            })],
          }),
        ],
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    const safeName = title.replace(/[^\w\s-]/g, "").replace(/\s+/g, "_") || "LPJ_Report";

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName}.docx"`);
    res.setHeader("Cache-Control", "no-store");
    res.status(200).send(buffer);
  } catch (err) {
    console.error("DOCX export error:", err);
    res.status(500).json({ error: err.message || "Internal server error" });
  }
};

module.exports.config = {
  api: {
    bodyParser: { sizeLimit: "20mb" },
  },
};
