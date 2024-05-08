import { Workbook, downloadXlsx, AlignHorizontal, Spreadsheet, Row, CellType } from "./excellent.js"
(() => {
  const downloadButton = document.getElementById("download") as HTMLButtonElement;

  downloadButton.addEventListener("click", async (event) => {
    event.preventDefault();

    const workbook = createWorkbook();

    try {
      await downloadXlsx("example.xlsx", workbook);
      console.log("Done");
    }
    catch (err) {
      console.error(err);
    }
  });
})()

function createTimesTablesSheet(): Spreadsheet {

  const titleRow: Row = {
    cells: [
      {
        alignHorizontal: AlignHorizontal.Center,
        data: "My Cool Times Tables",
        type: CellType.String,
        span: 11,
        bold: true,
        italic: true,
        underline: true
      }
    ]
  };

  const headerRow: Row = {
    cells: [
      { data: "", type: CellType.String }
    ]
  };

  for (let x = 1; x <= 10; x++) {
    headerRow.cells.push({
      type: CellType.Number,
      data: x,
      bold: true
    });
  }

  const dataRows = [];

  for (let y = 1; y <= 10; y++) {
    const row: Row = {
      cells: [
        {
          data: y,
          bold: true,
          type: CellType.Number
        }
      ]
    }

    for (let x = 1; x <= 10; x++) {
      row.cells.push({
        type: CellType.Number,
        data: x * y
      })
    }

    dataRows.push(row);
  }

  const timesTables: Spreadsheet = {
    title: "Times Tables",
    rows: [
      titleRow,
      headerRow,
      ...dataRows,
    ]
  }

  return timesTables;
}

function createWorkbook(): Workbook {

  const workbook: Workbook = {
    sheets: [
      createTimesTablesSheet(),
      {
        title: "Sheet 1",
        rows: [
          {
            cells: [
              { data: "Column 1", bold: true, span: 2, alignHorizontal: AlignHorizontal.Center, type: CellType.String },
              { data: "Column 2", bold: true, type: CellType.String },
            ]
          },
          {
            cells: [
              { data: "Row 1, Col 1", type: CellType.String },
              { data: "Row 1, Col 2", type: CellType.String },
              { data: "Row 1, Col 3", type: CellType.String },
            ]
          },
          {
            cells: [
              { data: "I am across all the rows!", span: 3, alignHorizontal: AlignHorizontal.Right, type: CellType.String }
            ]
          }
        ]
      }
    ],
    author: "Jim Bob"
  }
  return workbook;
}