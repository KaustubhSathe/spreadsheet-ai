import { Spreadsheet } from '../src/main';

describe('Formula Evaluation', () => {
  let spreadsheet: Spreadsheet;

  beforeEach(() => {
    document.body.innerHTML = `
      <div id="app">
        <div id="spreadsheet-container"></div>
        <input id="cell-name" />
        <input id="formula-input" />
        <div class="sheet-tabs">
          <div class="add-sheet">+</div>
        </div>
      </div>
    `;
    spreadsheet = new Spreadsheet('app');
  });

  it('should evaluate basic arithmetic', () => {
    spreadsheet['updateCell']('A1', '=1+1');
    expect(spreadsheet['data']['A1'].computed).toBe('2');
  });

  it('should handle cell references', () => {
    spreadsheet['updateCell']('A1', '1');
    spreadsheet['updateCell']('B1', '2');
    spreadsheet['updateCell']('C1', '=A1+B1');
    expect(spreadsheet['data']['C1'].computed).toBe('3');
  });

  it('should return original value for non-formulas', () => {
    spreadsheet['updateCell']('A1', '123');
    expect(spreadsheet['data']['A1'].computed).toBe('123');
    spreadsheet['updateCell']('A1', 'text');
    expect(spreadsheet['data']['A1'].computed).toBe('text');
  });

  it('should handle invalid formulas', () => {
    spreadsheet['updateCell']('A1', '=invalid');
    expect(spreadsheet['data']['A1'].computed).toBe('#ERROR!');
    spreadsheet['updateCell']('A1', '=1/0');
    expect(spreadsheet['data']['A1'].computed).toBe('#DIV/0!');
  });
}); 