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
      
      // First select A1
      fireEvent.mouseDown(cellA1);
      cellA1.classList.add('selected');
      fireEvent.mouseUp(cellA1);
      
      // Then shift+click B2
      const shiftClickOptions = { shiftKey: true };
      fireEvent.mouseDown(cellB2, shiftClickOptions);
      
      // Verify selection rectangle A1:B2
      [
        '[data-cell-id="A1"]',
        '[data-cell-id="A2"]',
        '[data-cell-id="B1"]',
        '[data-cell-id="B2"]'
      ].forEach(selector => {
        expect(container.querySelector(selector)).toHaveClass('selected');
      });
    });
  });

  describe('Cell Editing', () => {
    it('should make cell editable on double click', () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      const content = cell.querySelector('.cell-content')!;
      
      fireEvent.dblClick(cell);
      
      expect(content.getAttribute('contenteditable')).toBe('true');
    });

    it('should update cell value after editing', () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      const content = cell.querySelector('.cell-content')!;
      
      // Make cell editable
      fireEvent.dblClick(cell);
      
      // Edit cell content
      content.textContent = 'test';
      fireEvent.blur(content);
      
      expect(content).toHaveTextContent('test');
    });

    it('should maintain fill handle while editing', () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      const fillHandle = cell.querySelector('.fill-handle')!;
      
      // Select the cell first
      fireEvent.mouseDown(cell);
      fireEvent.mouseUp(cell);
      cell.classList.add('active');
      
      // Make cell editable
      fireEvent.dblClick(cell);
      
      // Fill handle should still be visible
      expect(fillHandle).toBeVisible();
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
    it('should update formula bar when selecting cell', () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      const content = cell.querySelector('.cell-content')!;
      const cellName = document.getElementById('cell-name') as HTMLInputElement;
      const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
      
      // Set some content first
      fireEvent.dblClick(cell);
      content.textContent = 'test';
      fireEvent.blur(content);
      
      // Select the cell
      fireEvent.click(cell);
      
      expect(cellName.value).toBe('A1');
      expect(formulaInput.value).toBe('test');
    });

    it('should update cell when changing formula', () => {
      const cell = container.querySelector('[data-cell-id="A1"]')!;
      const content = cell.querySelector('.cell-content')!;
      const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
      
      fireEvent.click(cell);
      formulaInput.value = '=1+1';
      fireEvent.change(formulaInput);
      
      expect(content).toHaveTextContent('2');
    });

    it('should show cell range in name box for multiple selection', () => {
      const cellA1 = container.querySelector('[data-cell-id="A1"]')!;
      const cellB2 = container.querySelector('[data-cell-id="B2"]')!;
      const cellName = document.getElementById('cell-name') as HTMLInputElement;
      
      // First select A1
      fireEvent.mouseDown(cellA1);
      cellA1.classList.add('selected');
      fireEvent.mouseUp(cellA1);
      
      // Then shift+click B2
      const shiftClickOptions = { shiftKey: true };
      fireEvent.mouseDown(cellB2, shiftClickOptions);
      
      expect(cellName.value).toBe('A1:B2');
    });

    it('should show single cell ID when one cell is selected', () => {
      const cellA1 = container.querySelector('[data-cell-id="A1"]')!;
      const cellName = document.getElementById('cell-name') as HTMLInputElement;
      
      fireEvent.mouseDown(cellA1);
      cellA1.classList.add('selected');
      fireEvent.mouseUp(cellA1);
      
      expect(cellName.value).toBe('A1');
    });

    it('should show cell range with top-left cell first', () => {
      const cellA1 = container.querySelector('[data-cell-id="A1"]')!;
      const cellB2 = container.querySelector('[data-cell-id="B2"]')!;
      const cellName = document.getElementById('cell-name') as HTMLInputElement;
      
      // Test selection from bottom-right to top-left
      fireEvent.mouseDown(cellB2);
      cellB2.classList.add('selected');
      fireEvent.mouseUp(cellB2);
      
      const shiftClickOptions = { shiftKey: true };
      fireEvent.mouseDown(cellA1, shiftClickOptions);
      
      // Should still show A1:B2, not B2:A1
      expect(cellName.value).toBe('A1:B2');
    });
  });

  describe('Fill Handle', () => {
    it('should fill cells when dragging the fill handle', () => {
      const cellA1 = container.querySelector('[data-cell-id="A1"]')!;
      const cellA2 = container.querySelector('[data-cell-id="A2"]')!;
      const contentA1 = cellA1.querySelector('.cell-content')!;
      const contentA2 = cellA2.querySelector('.cell-content')!;
      
      // Set initial value
      fireEvent.dblClick(cellA1);
      contentA1.textContent = '1';
      fireEvent.blur(contentA1);
      
      // Select cell to show fill handle
      fireEvent.mouseDown(cellA1);
      fireEvent.mouseUp(cellA1);
      cellA1.classList.add('active');
      
      // Start fill drag
      const fillHandle = cellA1.querySelector('.fill-handle')!;
      fireEvent.mouseDown(fillHandle);
      
      // Drag to A2
      fireEvent.mouseMove(cellA2);
      fireEvent.mouseUp(cellA2);
      
      // Check if A2 was filled with incremented value
      expect(contentA2).toHaveTextContent('2');
    });

    it('should not increment empty cells when filling', () => {
      const cellA1 = container.querySelector('[data-cell-id="A1"]')!;
      const cellA2 = container.querySelector('[data-cell-id="A2"]')!;
      const contentA2 = cellA2.querySelector('.cell-content')!;
      
      // Select empty cell
      fireEvent.mouseDown(cellA1);
      fireEvent.mouseUp(cellA1);
      cellA1.classList.add('active');
      
      // Start fill drag
      const fillHandle = cellA1.querySelector('.fill-handle')!;
      fireEvent.mouseDown(fillHandle);
      
      // Drag to A2
      fireEvent.mouseMove(cellA2);
      fireEvent.mouseUp(cellA2);
      
      // Check if A2 remains empty
      expect(contentA2).toHaveTextContent('');
    });

    it('should copy text content when filling', () => {
      const cellA1 = container.querySelector('[data-cell-id="A1"]')!;
      const cellA2 = container.querySelector('[data-cell-id="A2"]')!;
      const contentA1 = cellA1.querySelector('.cell-content')!;
      const contentA2 = cellA2.querySelector('.cell-content')!;
      
      // Set initial text value
      fireEvent.dblClick(cellA1);
      contentA1.textContent = 'test';
      fireEvent.blur(contentA1);
      
      // Select cell to show fill handle
      fireEvent.mouseDown(cellA1);
      fireEvent.mouseUp(cellA1);
      cellA1.classList.add('active');
      
      // Start fill drag
      const fillHandle = cellA1.querySelector('.fill-handle')!;
      fireEvent.mouseDown(fillHandle);
      
      // Drag to A2
      fireEvent.mouseMove(cellA2);
      fireEvent.mouseUp(cellA2);
      
      // Check if A2 was filled with same text
      expect(contentA2).toHaveTextContent('test');
    });
  });
}); 