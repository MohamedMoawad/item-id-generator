export function evaluateFormula(formula: string, values: { [internalName: string]: string }): string {
  const expression = formula.replace(/\{([A-Za-z_][A-Za-z0-9_]*)\}/g, (_match: string, name: string) => {
    const numeric = Number(values[name]);
    return Number.isFinite(numeric) ? String(numeric) : '0';
  });
  if (!/^[\d+\-*/().\s]+$/.test(expression)) {
    return '';
  }
  const result = compute(expression.replace(/\s+/g, ''));
  if (!Number.isFinite(result)) {
    return '';
  }
  return String(Math.round(result * 10000) / 10000);
}

export function applyFormulas(
  values: { [internalName: string]: string },
  columns: { internalName: string; formula?: string }[]
): { [internalName: string]: string } {
  const next = { ...values };
  columns.forEach((column) => {
    if (!column.formula) {
      return;
    }
    next[column.internalName] = evaluateFormula(column.formula, next);
  });
  return next;
}

export function columnTotal(rows: { values: { [internalName: string]: string } }[], internalName: string): number {
  return rows.reduce((sum, row) => {
    const numeric = Number(row.values[internalName]);
    return sum + (Number.isFinite(numeric) ? numeric : 0);
  }, 0);
}

class Expression {
  private index = 0;

  public constructor(private readonly expression: string) {}

  public compute(): number {
    return this.parseExpression();
  }

  private parseExpression(): number {
    let value = this.parseTerm();
    while (this.expression.charAt(this.index) === '+' || this.expression.charAt(this.index) === '-') {
      const operator = this.expression.charAt(this.index);
      this.index += 1;
      const right = this.parseTerm();
      value = operator === '+' ? value + right : value - right;
    }
    return value;
  }

  private parseTerm(): number {
    let value = this.parseFactor();
    while (this.expression.charAt(this.index) === '*' || this.expression.charAt(this.index) === '/') {
      const operator = this.expression.charAt(this.index);
      this.index += 1;
      const right = this.parseFactor();
      value = operator === '*' ? value * right : (right === 0 ? 0 : value / right);
    }
    return value;
  }

  private parseFactor(): number {
    if (this.expression.charAt(this.index) === '(') {
      this.index += 1;
      const value = this.parseExpression();
      if (this.expression.charAt(this.index) === ')') {
        this.index += 1;
      }
      return value;
    }
    const start = this.index;
    if (this.expression.charAt(this.index) === '-') {
      this.index += 1;
    }
    while (this.index < this.expression.length && /[\d.]/.test(this.expression.charAt(this.index))) {
      this.index += 1;
    }
    return Number(this.expression.slice(start, this.index)) || 0;
  }
}

function compute(expression: string): number {
  return new Expression(expression).compute();
}
