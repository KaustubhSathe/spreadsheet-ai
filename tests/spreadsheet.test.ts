import { fireEvent } from '@testing-library/dom';
import userEvent from '@testing-library/user-event';
import { Spreadsheet } from '../src/main';

describe('Spreadsheet', () => {
  let container: HTMLElement;
  let spreadsheet: Spreadsheet;

  beforeEach(() => {
    document.body.innerHTML = `
      <div id="app">
        <div id="spreadsheet-container"></div>
        <input id="cell-name" type="text" />
        <input id="formula-input" type="text" />
        <div class="sheet-tabs">
          <div class="add-sheet">+</div>
        </div>
      </div>
    `;
    container = document.getElementById('spreadsheet-container')!;
    spreadsheet = new Spreadsheet('app');

    // Mock selection API
    window.getSelection = () => ({
      removeAllRanges: jest.fn(),
      addRange: jest.fn(),
      getRangeAt: () => ({ cloneRange: jest.fn() }),
    } as any);
  });

  afterEach(() => {
    document.body.innerHTML = '';
  });

  describe('Grid Creation', () => {
    it('should create the initial grid with correct dimensions', () => {
      const columns = container.querySelectorAll('.column');
      const cellsInFirstColumn = columns[0].querySelectorAll('.cell');
      
      expect(columns).toHaveLength(26); // A-Z
      expect(cellsInFirstColumn).toHaveLength(50); // 50 rows
    });

    it('should initialize cells with correct IDs', () => {
      const firstCell = container.querySelector('[data-cell-id="A1"]');
      const lastCell = container.querySelector('[data-cell-id="Z50"]');
      
      expect(firstCell).toBeInTheDocument();
      expect(lastCell).toBeInTheDocument();
    });
  });

  describe('Cell Selection', () => {
    it('should select a cell on click', () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      // Simulate full click sequence
      fireEvent.mouseDown(cell);
      fireEvent.mouseUp(cell);
      fireEvent.click(cell);
      
      expect(cell).toHaveClass('selected');
    });

    it('should support multi-select with shift+click', () => {
      const cellA1 = container.querySelector('[data-cell-id="A1"]')!;
      const cellB2 = container.querySelector('[data-cell-id="B2"]')!;
      
      // Select first cell
      fireEvent.mouseDown(cellA1);
      fireEvent.mouseUp(cellA1);
      fireEvent.click(cellA1);
      
      // Shift-click second cell
      const shiftClickOptions = { shiftKey: true };
      fireEvent.mouseDown(cellB2, shiftClickOptions);
      fireEvent.mouseUp(cellB2, shiftClickOptions);
      fireEvent.click(cellB2, shiftClickOptions);
      
      expect(cellA1).toHaveClass('selected');
      expect(cellB2).toHaveClass('selected');
      expect(container.querySelector('[data-cell-id="A2"]')).toHaveClass('selected');
      expect(container.querySelector('[data-cell-id="B1"]')).toHaveClass('selected');
    });
  });

  describe('Cell Editing', () => {
    it('should make cell editable on double click', () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      
      // Simulate full double click sequence
      fireEvent.mouseDown(cell);
      fireEvent.mouseUp(cell);
      fireEvent.click(cell);
      fireEvent.mouseDown(cell);
      fireEvent.mouseUp(cell);
      fireEvent.click(cell);
      fireEvent.dblClick(cell);
      
      expect(cell.getAttribute('contenteditable')).toBe('true');
    });

    it('should update cell value after editing', () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      
      // Make cell editable
      fireEvent.mouseDown(cell);
      fireEvent.mouseUp(cell);
      fireEvent.click(cell);
      fireEvent.mouseDown(cell);
      fireEvent.mouseUp(cell);
      fireEvent.click(cell);
      fireEvent.dblClick(cell);
      
      // Edit cell
      cell.textContent = 'test';
      fireEvent.blur(cell);
      
      expect(cell).toHaveTextContent('test');
    });
  });

  describe('Column Resize', () => {
    it('should update column styles on resize', () => {
      const handle = container.querySelector('.col-resize-handle')!;
      const column = handle.closest('.column') as HTMLElement;
      
      fireEvent.mouseDown(handle, { clientX: 0 });
      fireEvent.mouseMove(document, { clientX: 50 });
      fireEvent.mouseUp(document);
      
      expect(column.style.width).toBe('50px');
      expect(column.style.minWidth).toBe('50px');
    });
  });

  describe('Row Resize', () => {
    it('should update row styles on resize', () => {
      const handle = container.querySelector('.row-resize-handle')!;
      const rowContainer = handle.closest('.row-container') as HTMLElement;
      
      fireEvent.mouseDown(handle, { clientY: 0 });
      fireEvent.mouseMove(document, { clientY: 50 });
      fireEvent.mouseUp(document);
      
      expect(rowContainer.style.height).toBe('50px');
    });
  });

  describe('Formula Bar', () => {
    it('should update formula bar when selecting cell', async () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      const cellName = document.getElementById('cell-name') as HTMLInputElement;
      const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
      
      await userEvent.click(cell);
      
      expect(cellName.value).toBe('A1');
      expect(formulaInput.value).toBe('');
    });

    it('should update cell when changing formula', async () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
      
      await userEvent.click(cell);
      await userEvent.type(formulaInput, '=1+1');
      fireEvent.change(formulaInput);
      
      expect(cell).toHaveTextContent('2');
    });
  });
}); 