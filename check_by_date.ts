interface Student {
  id: string;
  name: string;
}

const main = (workbook: ExcelScript.Workbook): void => {
  const readSheet = workbook.getWorksheet("前期");
  const writeSheet = workbook.getWorksheet("検索シート");
  const checkMark = writeSheet.getRange('C2').getValue() === "ペナルティ" ? "●" : "〇";

  const findColumnByDate = (targetDate: number): string[] => {
    const targetRange = readSheet.getRange('D1:G1').getValues()[0];

    const maxCol = 36 + 1;
    const row = targetRange.indexOf(targetDate);
    return  row < 0
      ? []
      : readSheet.getRangeByIndexes(1, row + 3, maxCol, 1).getValues().flat();
  };

  const clearCells = (): void => {
    const maxCol = writeSheet.getRange().getColumnCount();
    writeSheet
      .getRange(`C7:D${maxCol}`)
      .clear(ExcelScript.ClearApplyTo.all);
  };

  const writeCells = (cells: string[]): void => {
    const students: Student[] = [];

    const ids = readSheet.getRange('A2:A36').getValues();
    const names = readSheet.getRange('B2:B36').getValues();

    cells.forEach((cell, idx) => {
      if (cell[0] === checkMark) {
        students.push({
          id: ids[idx][0] as string,
          name: names[idx][0] as string,
        });
      }
    })

    const startCol = 7;
    students.forEach((student, col) => {
      writeSheet
        .getRange(`C${col + startCol}`)
        .setValue(student.id);
      writeSheet
        .getRange(`D${col + startCol}`)
        .setValue(student.name);
    });
  };
  
  const checkByDate = (): void => {
    const targetDate = writeSheet.getRange('C4').getValue() as number;
    const cells = findColumnByDate(targetDate);

    clearCells();
    writeCells(cells);
  }

  checkByDate();
}
