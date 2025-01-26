import './style.css';
import { SpreadsheetData, Sheet } from './types';
import { getCellId } from './utils';

export class Spreadsheet {
  private container: HTMLElement;
  private data: SpreadsheetData = {};
  private activeCell: HTMLElement | null = null;
  private rows: number = 50;
  private cols: number = 26;
  private sheets: Sheet[] = [];
  private activeSheetId: string = '';
  private isSelecting: boolean = false;
  private selectionStart: HTMLElement | null = null;
  private selectionEnd: HTMLElement | null = null;
  private selectionAnchor: HTMLElement | null = null;
  private isResizing: boolean = false;
  private resizeStartX: number = 0;
  private resizeStartY: number = 0;
  private resizeElement: HTMLElement | null = null;
  private initialSize: number = 0;

  constructor(containerId: string) {
    this.container = document.getElementById('spreadsheet-container')!;
    this.init();
    this.setupFormulaBar();
    this.setupSheetTabs();
  }

  private init(): void {
    // Create initial sheet
    this.addSheet();
    this.createSpreadsheet();
    this.attachEventListeners();
  }

  private addSheet(): void {
    const sheetNumber = this.sheets.length + 1;
    const newSheet: Sheet = {
      id: crypto.randomUUID(),
      name: `Sheet${sheetNumber}`,
      data: {}
    };
    
    this.sheets.push(newSheet);
    this.activeSheetId = newSheet.id;
    this.data = newSheet.data;
    
    // Initialize cells for the new sheet
    for (let row = 0; row < this.rows; row++) {
      for (let col = 0; col < this.cols; col++) {
        const cellId = getCellId(row, col);
        this.data[cellId] = {
          value: '',
          formula: '',
          computed: ''
        };
      }
    }

    this.updateSheetTabs();
  }

  private setupSheetTabs(): void {
    const addSheetButton = document.querySelector('.add-sheet') as HTMLElement;
    const sheetTabs = document.querySelector('.sheet-tabs') as HTMLElement;

    addSheetButton.addEventListener('click', () => {
      this.addSheet();
    });

    sheetTabs.addEventListener('click', (e) => {
      const tab = (e.target as HTMLElement).closest('.tab') as HTMLElement;
      if (tab && !tab.classList.contains('active')) {
        const sheetId = tab.dataset.sheetId!;
        this.switchSheet(sheetId);
      }
    });
  }

  private switchSheet(sheetId: string): void {
    const sheet = this.sheets.find(s => s.id === sheetId);
    if (sheet) {
      this.activeSheetId = sheetId;
      this.data = sheet.data;
      this.updateSheetTabs();
      this.renderSheet();
    }
  }

  private updateSheetTabs(): void {
    const tabsContainer = document.querySelector('.sheet-tabs') as HTMLElement;
    const addSheetButton = tabsContainer.querySelector('.add-sheet') as HTMLElement;
    
    // Remove all existing tabs
    tabsContainer.querySelectorAll('.tab').forEach(tab => tab.remove());
    
    // Add tabs for each sheet after the add button
    this.sheets.forEach(sheet => {
      const tab = document.createElement('div');
      tab.className = `tab${sheet.id === this.activeSheetId ? ' active' : ''}`;
      tab.textContent = sheet.name;
      tab.dataset.sheetId = sheet.id;
      tabsContainer.appendChild(tab);
    });
  }

  private renderSheet(): void {
    // Clear existing cells
    const cells = this.container.querySelectorAll('td');
    cells.forEach(cell => {
      const cellId = cell.dataset.cellId!;
      cell.textContent = this.data[cellId]?.computed || '';
    });
  }

  private createSpreadsheet(): void {
    const grid = document.createElement('div');
    grid.className = 'spreadsheet';

    // Create main grid container
    const gridContainer = document.createElement('div');
    gridContainer.className = 'grid-container';

    // Row numbers column (including header)
    const rowNumbersColumn = document.createElement('div');
    rowNumbersColumn.className = 'row-numbers';
    
    // Add corner cell to row numbers
    const cornerCell = document.createElement('div');
    cornerCell.className = 'corner-cell';
    rowNumbersColumn.appendChild(cornerCell);
    
    // Add row numbers with resize handles
    for (let row = 0; row < this.rows; row++) {
      const rowContainer = document.createElement('div');
      rowContainer.className = 'row-container';

      const rowNumber = document.createElement('div');
      rowNumber.className = 'row-number';
      rowNumber.textContent = (row + 1).toString();
      
      const rowResizeHandle = document.createElement('div');
      rowResizeHandle.className = 'row-resize-handle';
      rowResizeHandle.dataset.row = row.toString();
      
      rowContainer.appendChild(rowNumber);
      rowContainer.appendChild(rowResizeHandle);
      rowNumbersColumn.appendChild(rowContainer);
    }
    gridContainer.appendChild(rowNumbersColumn);

    // Create columns with headers and resize handles
    for (let col = 0; col < this.cols; col++) {
      const column = document.createElement('div');
      column.className = 'column';
      
      // Add column header with resize handle
      const headerContainer = document.createElement('div');
      headerContainer.className = 'header-container';
      
      const header = document.createElement('div');
      header.className = 'column-header';
      header.textContent = String.fromCharCode(65 + col);
      
      const colResizeHandle = document.createElement('div');
      colResizeHandle.className = 'col-resize-handle';
      colResizeHandle.dataset.col = col.toString();
      
      headerContainer.appendChild(header);
      headerContainer.appendChild(colResizeHandle);
      column.appendChild(headerContainer);
      
      // Create cells for this column
      for (let row = 0; row < this.rows; row++) {
        const cell = document.createElement('div');
        cell.className = 'cell';
        const cellId = getCellId(row, col);
        cell.dataset.cellId = cellId;
        cell.contentEditable = 'false';
        column.appendChild(cell);

        // Initialize cell data
        this.data[cellId] = {
          value: '',
          formula: '',
          computed: ''
        };
      }
      gridContainer.appendChild(column);
    }

    grid.appendChild(gridContainer);
    this.container.appendChild(grid);
  }

  private attachEventListeners(): void {
    // Handle selection start
    this.container.addEventListener('mousedown', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (cell) {
        // First, handle any active editing
        const editingCell = this.container.querySelector('.cell[contenteditable="true"]') as HTMLElement;
        if (editingCell && editingCell !== cell) {
          this.finishEditing(editingCell);
        }

        this.isSelecting = true;
        
        if (e.shiftKey && this.activeCell) {
          // For shift + click, extend selection from active cell
          this.selectionAnchor = this.activeCell;
          this.selectionStart = this.selectionAnchor;
          this.selectionEnd = cell;
          this.updateSelection(); // Update selection immediately
        } else if (!cell.hasAttribute('contenteditable') || cell.getAttribute('contenteditable') === 'false') {
          // Only start new selection if cell isn't being edited
          this.clearSelection();
          this.selectionStart = cell;
          this.selectionEnd = cell;
          this.selectionAnchor = cell;
          cell.classList.add('selected'); // Add selected class immediately
        }
        
        this.activeCell = cell;
        this.updateFormulaBar(); // Update formula bar with selection range
      }
    });

    // Handle selection drag
    this.container.addEventListener('mousemove', (e) => {
      if (this.isSelecting) {
        const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
        if (cell && cell !== this.selectionEnd) {
          this.selectionEnd = cell;
          this.updateSelection();
          this.updateFormulaBar(); // Update formula bar while dragging
        }
      }
    });

    // Handle selection end
    document.addEventListener('mouseup', () => {
      this.isSelecting = false;
    });

    // Handle cell selection on click
    this.container.addEventListener('click', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (cell) {
        // Don't handle click if we're already editing
        if (cell.contentEditable === 'true') {
          return;
        }

        // Remove selected class from previously selected cell
        const prevSelected = this.container.querySelector('.cell.selected');
        if (prevSelected) {
          prevSelected.classList.remove('selected');
        }
        
        this.activeCell = cell;
        cell.classList.add('selected');
        
        const cellId = cell.dataset.cellId!;
        this.updateFormulaBar(cellId); // Update formula bar for single cell
      }
    });

    // Make cell editable on double click
    this.container.addEventListener('dblclick', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (cell) {
        // Keep the selection state but make the cell editable
        cell.setAttribute('contenteditable', 'true');
        const cellId = cell.dataset.cellId!;
        cell.textContent = this.data[cellId].formula || this.data[cellId].value;
        cell.focus();
      }
    });

    // Handle cell blur (finish editing)
    this.container.addEventListener('blur', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (cell) {
        cell.setAttribute('contenteditable', 'false');
        this.finishEditing(cell);
      }
    }, true);

    // Move keydown listener to document level
    document.addEventListener('keydown', (e) => {
      // If no cell is selected or we're editing, don't handle navigation
      if (!this.activeCell || this.activeCell.contentEditable === 'true') {
        if (this.activeCell?.contentEditable === 'true' && e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          this.activeCell.blur();
          
          // Move to the next row after finishing edit
          const currentCellId = this.activeCell.dataset.cellId!;
          const [col, row] = this.parseCellId(currentCellId);
          const nextRow = Math.min(this.rows - 1, row + 1);
          const nextCellId = getCellId(nextRow, col);
          const nextCell = this.container.querySelector(`[data-cell-id="${nextCellId}"]`) as HTMLElement;
          
          if (nextCell) {
            // Remove selected from current cell
            this.activeCell.classList.remove('selected');
            // Select next cell
            nextCell.classList.add('selected');
            this.activeCell = nextCell;
            this.updateFormulaBar(nextCellId); // Update formula bar for single cell
          }
        }
        return;
      }
      
      // Handle navigation
      this.handleKeyNavigation(e);
    });

    // Add resize event listeners
    this.container.addEventListener('mousedown', (e) => {
      const target = e.target as HTMLElement;
      if (target.classList.contains('col-resize-handle')) {
        this.startColumnResize(e, target);
      } else if (target.classList.contains('row-resize-handle')) {
        this.startRowResize(e, target);
      }
    });

    document.addEventListener('mousemove', (e) => {
      if (this.isResizing) {
        this.handleResize(e);
      }
    });

    document.addEventListener('mouseup', () => {
      this.isResizing = false;
      this.resizeElement = null;
    });
  }

  private finishEditing(cell: HTMLElement): void {
    cell.contentEditable = 'false';
    const cellId = cell.dataset.cellId!;
    const newValue = cell.textContent || '';
    this.updateCell(cellId, newValue);
    this.recalculateAll();
    
    // Show computed value but maintain selection state
    cell.textContent = this.data[cellId].computed;
  }

  private updateCell(cellId: string, value: string): void {
    this.data[cellId] = {
      value: value,
      formula: value.startsWith('=') ? value : '',
      computed: this.evaluateFormula(value, this.data)
    };

    const cell = this.container.querySelector(`[data-cell-id="${cellId}"]`);
    if (cell) {
      cell.textContent = this.data[cellId].computed;
    }
  }

  private recalculateAll(): void {
    for (const cellId in this.data) {
      if (this.data[cellId].formula) {
        this.updateCell(cellId, this.data[cellId].formula);
      }
    }
  }

  private setupFormulaBar(): void {
    const cellName = document.getElementById('cell-name') as HTMLInputElement;
    const formulaInput = document.getElementById('formula-input') as HTMLInputElement;

    this.container.addEventListener('focus', (e) => {
      const cell = (e.target as HTMLElement).closest('td') as HTMLElement;
      if (cell) {
        const cellId = cell.dataset.cellId!;
        cellName.value = cellId;
        formulaInput.value = this.data[cellId].formula || this.data[cellId].value;
      }
    }, true);

    formulaInput.addEventListener('change', (e) => {
      if (this.activeCell) {
        const cellId = this.activeCell.dataset.cellId!;
        const value = (e.target as HTMLInputElement).value;
        this.updateCell(cellId, value);
        this.recalculateAll();
      }
    });
  }

  private handleKeyNavigation(e: KeyboardEvent): void {
    // Skip navigation if we're editing a cell
    if (!this.activeCell || this.activeCell.contentEditable === 'true') return;
    
    const currentCellId = this.activeCell.dataset.cellId!;
    const [col, row] = this.parseCellId(currentCellId);
    
    let nextRow = row;
    let nextCol = col;
    
    switch (e.key) {
      case 'ArrowUp':
        nextRow = Math.max(0, row - 1);
        e.preventDefault();
        break;
      case 'ArrowDown':
        nextRow = Math.min(this.rows - 1, row + 1);
        e.preventDefault();
        break;
      case 'ArrowLeft':
        nextCol = Math.max(0, col - 1);
        e.preventDefault();
        break;
      case 'ArrowRight':
        nextCol = Math.min(this.cols - 1, col + 1);
        e.preventDefault();
        break;
      case 'Tab':
        nextCol = e.shiftKey ? Math.max(0, col - 1) : Math.min(this.cols - 1, col + 1);
        if (nextCol === col) {
          nextRow = e.shiftKey ? Math.max(0, row - 1) : Math.min(this.rows - 1, row + 1);
          nextCol = e.shiftKey ? this.cols - 1 : 0;
        }
        e.preventDefault();
        break;
      case 'Enter':
        nextRow = e.shiftKey ? Math.max(0, row - 1) : Math.min(this.rows - 1, row + 1);
        e.preventDefault();
        break;
      default:
        return;
    }

    const nextCellId = getCellId(nextRow, nextCol);
    const nextCell = this.container.querySelector(`[data-cell-id="${nextCellId}"]`) as HTMLElement;
    
    if (nextCell) {
      if (e.shiftKey) {
        // For shift + arrow keys, extend selection from anchor point
        if (!this.selectionAnchor) {
          this.selectionAnchor = this.activeCell;
        }
        this.selectionStart = this.selectionAnchor;
        this.selectionEnd = nextCell;
        this.updateSelection();
        this.updateFormulaBar(); // Update formula bar for shift selection
      } else {
        // For normal navigation, clear selection and select single cell
        this.clearSelection();
        this.selectionStart = nextCell;
        this.selectionEnd = nextCell;
        this.selectionAnchor = nextCell;
        nextCell.classList.add('selected');
        this.updateFormulaBar(nextCellId); // Update formula bar for single cell
      }
      
      this.activeCell = nextCell;
    }
  }

  private parseCellId(cellId: string): [number, number] {
    // Convert "A1" format to [col, row]
    const col = cellId.charCodeAt(0) - 65; // 'A' is 65 in ASCII
    const row = parseInt(cellId.slice(1)) - 1; // Convert 1-based to 0-based
    return [col, row];
  }

  private clearSelection(): void {
    const selectedCells = this.container.querySelectorAll('.cell.selected');
    const editingCell = this.container.querySelector('.cell[contenteditable="true"]');
    
    selectedCells.forEach(cell => {
      if (cell !== editingCell) {
        cell.classList.remove('selected');
      }
    });
  }

  private updateSelection(): void {
    if (!this.selectionStart || !this.selectionEnd) return;

    // Clear previous selection
    this.clearSelection();

    // Get the range of cells to select
    const startId = this.selectionStart.dataset.cellId!;
    const endId = this.selectionEnd.dataset.cellId!;
    const [startCol, startRow] = this.parseCellId(startId);
    const [endCol, endRow] = this.parseCellId(endId);

    // Calculate the selection rectangle
    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = Math.min(startCol, endCol);
    const maxCol = Math.max(startCol, endCol);

    // Select all cells in the range
    for (let row = minRow; row <= maxRow; row++) {
      for (let col = minCol; col <= maxCol; col++) {
        const cellId = getCellId(row, col);
        const cell = this.container.querySelector(`[data-cell-id="${cellId}"]`);
        if (cell) {
          cell.classList.add('selected');
        }
      }
    }
  }

  private startColumnResize(e: MouseEvent, handle: HTMLElement): void {
    this.isResizing = true;
    this.resizeStartX = e.clientX;
    this.resizeElement = handle;
    const column = handle.closest('.column') as HTMLElement;
    if (column) {
      this.initialSize = column.offsetWidth;
    }
  }

  private startRowResize(e: MouseEvent, handle: HTMLElement): void {
    this.isResizing = true;
    this.resizeStartY = e.clientY;
    this.resizeElement = handle;
    const rowContainer = handle.closest('.row-container') as HTMLElement;
    if (rowContainer) {
      this.initialSize = rowContainer.offsetHeight;
    }
  }

  private handleResize(e: MouseEvent): void {
    if (!this.resizeElement) return;

    if (this.resizeElement.classList.contains('col-resize-handle')) {
      // Column resize
      const column = this.resizeElement.closest('.column') as HTMLElement;
      if (column) {
        const deltaX = e.clientX - this.resizeStartX;
        const newWidth = Math.max(50, this.initialSize + deltaX);
        column.style.width = `${newWidth}px`;
        column.style.minWidth = `${newWidth}px`;
      }
    } else if (this.resizeElement.classList.contains('row-resize-handle')) {
      // Row resize
      const rowContainer = this.resizeElement.closest('.row-container') as HTMLElement;
      if (rowContainer) {
        const deltaY = e.clientY - this.resizeStartY;
        const newHeight = Math.max(21, this.initialSize + deltaY);
        
        // Update row container height
        rowContainer.style.height = `${newHeight}px`;
        
        // Get row number from the handle's dataset
        const rowNumber = parseInt(this.resizeElement.dataset.row || '0');
        
        // Update all cells in this row
        for (let col = 0; col < this.cols; col++) {
          const cellId = getCellId(rowNumber, col);
          const cell = this.container.querySelector(`[data-cell-id="${cellId}"]`);
          if (cell) {
            (cell as HTMLElement).style.height = `${newHeight}px`;
          }
        }
      }
    }
  }

  private evaluateFormula(formula: string, data: SpreadsheetData): string {
    if (!formula.startsWith('=')) {
      return formula;
    }

    try {
      const expression = formula.slice(1);
      const evaluatedExpression = expression.replace(/[A-Z]\d+/g, (match) => {
        return data[match]?.computed || '0';
      });
      // Check for division by zero before evaluation
      if (expression.includes('/0')) {
        return '#DIV/0!';
      }
      return eval(evaluatedExpression).toString();
    } catch (e) {
      return '#ERROR!';
    }
  }

  private updateFormulaBar(cellId?: string): void {
    const cellName = document.getElementById('cell-name') as HTMLInputElement;
    const formulaInput = document.getElementById('formula-input') as HTMLInputElement;

    if (this.selectionStart && this.selectionEnd) {
      const startId = this.selectionStart.dataset.cellId!;
      const endId = this.selectionEnd.dataset.cellId!;
      
      if (startId === endId) {
        cellName.value = startId;
      } else {
        // Calculate which cell should be displayed first
        const [startCol, startRow] = this.parseCellId(startId);
        const [endCol, endRow] = this.parseCellId(endId);
        
        const minCell = getCellId(
          Math.min(startRow, endRow),
          Math.min(startCol, endCol)
        );
        const maxCell = getCellId(
          Math.max(startRow, endRow),
          Math.max(startCol, endCol)
        );
        
        cellName.value = `${minCell}:${maxCell}`;
      }
      
      // Show formula/value of active cell
      if (this.activeCell) {
        const activeCellId = this.activeCell.dataset.cellId!;
        formulaInput.value = this.data[activeCellId].formula || this.data[activeCellId].value;
      }
    } else if (cellId) {
      // Single cell selection
      cellName.value = cellId;
      formulaInput.value = this.data[cellId].formula || this.data[cellId].value;
    }
  }
}

// Only initialize if we're in the browser and not in test environment
if (typeof window !== 'undefined' && process.env.NODE_ENV !== 'test') {
  new Spreadsheet('app');
} 