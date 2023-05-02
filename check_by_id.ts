const toDateFormat = (serialValue: number): string => {
  // シリアル値からDate型へ変換
  const date = new Date(Math.round((serialValue - 25569) * 86400 * 1000));
  return `${date.getMonth() + 1}月${date.getDate()}日`;
};

const main = (workbook: ExcelScript.Workbook): void => {
  const readSheet = workbook.getWorksheet("前期");
  const writeSheet = workbook.getWorksheet("検索シート");
  const checkMark = writeSheet.getRange('C2').getValue() === "ペナルティ" ? "●" : "〇";

  const findColumnByDate = (targetId: string): string[] => {
    const targetRange: string[] = readSheet.getRange('A2:A36').getValues().flat();

    const maxRow = 152;
    const col = targetRange.indexOf(targetId);
    return col < 0
      ? []
      : readSheet.getRangeByIndexes(col + 1, 3, 1, maxRow).getValues().flat();
  };

  const clearCells = (): void => {
    const maxCol = writeSheet.getRange().getColumnCount();
    writeSheet
        .getRange(`G7:G${maxCol}`)
        .clear(ExcelScript.ClearApplyTo.contents);
  };

  const writeCells = (cells: string[]): void => {
    const dates: number[] = [];
    const targetDates: number[] = readSheet.getRange('D1:EY1').getValues()[0] as number[];

    cells.forEach((cell, idx) => {
      if (cell[0] === checkMark) {
        dates.push(targetDates[idx]);
      }
    })

    const startCol = 7;
    dates.forEach((id, col) => {
      writeSheet
        .getRange(`G${col + startCol}`)
        .setValue(toDateFormat(id));
    });
  };

  const checkById = (): void => {
    const targetId = writeSheet.getRange('G4').getValue() as string;
    const cells = findColumnByDate(targetId);

    clearCells();
    writeCells(cells);
  }

  checkById();
}
