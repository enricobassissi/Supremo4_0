function getLotsaValues() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSoci = ss.getSheetByName("Soci");
    const values3D = sheetSoci
      .getRangeList(['A', 'G', 'I'])
      .getRanges()
      .map(range.getValues());
    
  }