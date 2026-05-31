import { 
  Document, 
  Packer, 
  Paragraph, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  AlignmentType, 
  VerticalAlign, 
  TextRun,
  BorderStyle
} from 'docx';
import { saveAs } from 'file-saver';
import { CardData, CardStatus } from '../types';

// Map status levels to hex colors for the MS Word cells
const HEX_COLORS: { [key in CardStatus]: string } = {
  Inicio: "EF4444", // Rojo standard vibrant
  Proceso: "F97316", // Naranja
  Logrado: "22C55E", // Verde
  Destacado: "0EA5E9" // Celeste / Azul claro
};

export const exportDeckToWord = async (
  cards: CardData[], 
  columnsCount: number = 3, 
  fileName: string = "tarjetas_tabu_editables"
) => {
  if (cards.length === 0) {
    throw new Error("No hay tarjetas activas en la baraja para exportar.");
  }

  const cardsPerPage = columnsCount * 4; // 12 cards for 3 columns, 16 cards for 4 columns, 8 for 2 columns
  const pages: CardData[][] = [];
  
  for (let i = 0; i < cards.length; i += cardsPerPage) {
    pages.push(cards.slice(i, i + cardsPerPage));
  }

  const docElements: any[] = [];
  const percentSize = Math.floor(100 / columnsCount);

  pages.forEach((pageCards, pageIdx) => {
    const tableRows: TableRow[] = [];

    // Group cards of this page into rows of size columnsCount (2, 3, or 4)
    for (let i = 0; i < pageCards.length; i += columnsCount) {
      const chunk = pageCards.slice(i, i + columnsCount);
      
      // Create cells for this row
      const cellList: TableCell[] = chunk.map((card) => {
        const status = (card.nivelLogro || 'Inicio') as CardStatus;
        const hexColor = HEX_COLORS[status] || HEX_COLORS.Inicio;

        return new TableCell({
          width: {
            size: percentSize,
            type: WidthType.PERCENTAGE,
          },
          // Setup margins inside the cell to give cards a real margins representation (padding)
          margins: {
            top: 140,
            bottom: 140,
            left: 140,
            right: 140
          },
          // Outer card border: thin black border
          borders: {
            top: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
            bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
            left: { style: BorderStyle.SINGLE, size: 12, color: "000000" },
            right: { style: BorderStyle.SINGLE, size: 12, color: "000000" }
          },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Table({
              width: {
                size: 100,
                type: WidthType.PERCENTAGE
              },
              borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.NONE }
              },
              rows: [
                // Row 1: Solid colored header banner matching the status color
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 100, type: WidthType.PERCENTAGE },
                      shading: { fill: hexColor },
                      margins: { top: 120, bottom: 120, left: 100, right: 100 },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: card.secreto.toUpperCase(),
                              bold: true,
                              size: 22, // 11pt
                              color: "FFFFFF"
                            })
                          ]
                        })
                      ]
                    })
                  ]
                }),
                // Row 2: Contains category tag & level indicator (just below the banner)
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 100, type: WidthType.PERCENTAGE },
                      margins: { top: 60, bottom: 40, left: 80, right: 80 },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: `${(card.categoria || "GENERAL").toUpperCase()}  |  ${status.toUpperCase()}`,
                              size: 13, // 6.5pt
                              color: "94A3B8",
                              bold: true
                            })
                          ]
                        })
                      ]
                    })
                  ]
                }),
                // Row 3: Card body containing tabu words and socratic hint at bottom
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 100, type: WidthType.PERCENTAGE },
                      margins: { top: 140, bottom: 140, left: 100, right: 100 },
                      children: [
                        // taboo words centered with spacing & elegant line separators
                        ...card.tabu.flatMap((taboWord, tabWordIdx) => [
                          new Paragraph({
                            alignment: AlignmentType.CENTER,
                            children: [
                              new TextRun({
                                text: taboWord.toUpperCase(),
                                bold: true,
                                size: 18, // 9pt
                                color: "1E293B" // Deep charcoal
                              })
                            ]
                          }),
                          tabWordIdx < card.tabu.length - 1 ? new Paragraph({
                            alignment: AlignmentType.CENTER,
                            spacing: { before: 10, after: 10 },
                            children: [
                              new TextRun({
                                text: "─────",
                                color: "E2E8F0",
                                size: 8
                              })
                            ]
                          }) : null
                        ].filter(Boolean) as Paragraph[]),

                        // Dotted divider before hint
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          spacing: { before: 120, after: 120 },
                          children: [
                            new TextRun({
                              text: "┈┈┈┈┈┈┈┈┈┈┈┈┈┈",
                              color: "CBD5E1",
                              size: 12
                            })
                          ]
                        }),

                        // Socratic hint centered below
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "pista:  ",
                              size: 14, // 7pt
                              color: "4F46E5", // Elegant indigo
                              bold: true
                            }),
                            new TextRun({
                              text: (card.pista || "Sin ayuda registrada.").toLowerCase(),
                              italics: true,
                              size: 14, // 7pt
                              color: "475569"
                            })
                          ]
                        })
                      ]
                    })
                  ]
                })
              ]
            })
          ]
        });
      });

      // Handle padding cells for columns to fit precisely
      while (cellList.length < columnsCount) {
        cellList.push(
          new TableCell({
            width: { size: percentSize, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE }
            },
            children: []
          })
        );
      }

      tableRows.push(
        new TableRow({
          children: cellList,
          cantSplit: true
        })
      );
    }

    // Pad table rows to hit exactly 4 rows for clean print formatting alignment
    while (tableRows.length < 4) {
      const emptyRowCells: TableCell[] = [];
      for (let c = 0; c < columnsCount; c++) {
        emptyRowCells.push(
          new TableCell({
            width: { size: percentSize, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE }
            },
            children: [new Paragraph({ children: [] })]
          })
        );
      }
      tableRows.push(
        new TableRow({
          children: emptyRowCells,
          cantSplit: true
        })
      );
    }

    // Push the compiled page table
    docElements.push(
      new Table({
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        rows: tableRows
      })
    );

    // If there is another page, inject a page break after the current table
    if (pageIdx < pages.length - 1) {
      docElements.push(
        new Paragraph({
          pageBreakBefore: true,
          children: []
        })
      );
    }
  });

  // Construct the global printable deck Document
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 720,
            bottom: 720,
            left: 720,
            right: 720,
          },
        },
      },
      children: docElements
    }]
  });

  const blob = await Packer.toBlob(doc);
  // Clean file name
  const cleanName = fileName.trim().replace(/[^a-zA-Z0-9_\-]/g, "_") || "tarjetas_tabu_editables";
  saveAs(blob, `${cleanName}.docx`);
};
