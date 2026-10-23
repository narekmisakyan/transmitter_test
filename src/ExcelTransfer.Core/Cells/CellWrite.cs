namespace ExcelTransfer.Core.Cells;

/// <summary>A value to write into a worksheet cell (1-based row and column). A Blank value clears the cell.</summary>
public readonly record struct CellWrite(int Row, int Column, CellValue Value);
