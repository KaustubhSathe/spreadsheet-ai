export function getCellId(row: number, col: number): string {
  const colStr = String.fromCharCode(65 + col);
  const rowStr = (row + 1).toString();
  return colStr + rowStr;
} 