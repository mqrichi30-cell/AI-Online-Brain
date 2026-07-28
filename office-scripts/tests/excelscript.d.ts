declare namespace ExcelScript {
  interface Range { getValues(): (string | number | boolean)[][]; }
  interface Worksheet { getUsedRange(): Range | undefined; }
  interface Workbook {
    getWorksheet(name: string): Worksheet | undefined;
    getActiveWorksheet(): Worksheet;
  }
}
