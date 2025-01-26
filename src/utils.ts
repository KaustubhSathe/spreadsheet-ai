export function getCellId(row: number, col: number): string {
  const colLabel = String.fromCharCode(65 + col);
  return `${colLabel}${row + 1}`;
}

export function evaluateFormula(formula: string, data: any): string {
  if (!formula.startsWith('=')) {
    return formula;
  }

  try {
    // Remove the '=' sign and evaluate the formula
    const expression = formula.substring(1);
    // Replace cell references with their values
    const evaluatedExpression = expression.replace(/[A-Z]\d+/g, (match) => {
      return data[match]?.computed || '0';
    });
    return eval(evaluatedExpression).toString();
  } catch (error) {
    return '#ERROR!';
  }
} 