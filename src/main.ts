import './style.css';
import { SpreadsheetData, Sheet } from './types';
import { getCellId, evaluateFormula } from './utils';

class Spreadsheet {
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
      const tab = (e.target as HTMLElement).closest('.tab');
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
      const cell = (e.target as HTMLElement).closest('.cell');
      if (cell) {
        // Store the currently editing cell
        const editingCell = this.container.querySelector('.cell[contenteditable="true"]');
        
        // If clicking a different cell while editing, finish editing first
        if (editingCell && editingCell !== cell) {
          this.finishEditing(editingCell as HTMLElement);
        }

        this.isSelecting = true;
        
        if (e.shiftKey && this.activeCell) {
          // For shift + click, extend selection from active cell
          this.selectionAnchor = this.activeCell;
          this.selectionStart = this.selectionAnchor;
          this.selectionEnd = cell;
        } else if (cell !== editingCell) { // Only clear selection if not clicking the editing cell
          // Normal click starts new selection
          this.clearSelection();
          this.selectionStart = cell;
          this.selectionEnd = cell;
          this.selectionAnchor = cell;
        }
        
        this.activeCell = cell;
        this.updateSelection();
        
        // Update formula bar
        const cellId = cell.dataset.cellId!;
        const cellName = document.getElementById('cell-name') as HTMLInputElement;
        const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
        cellName.value = cellId;
        formulaInput.value = this.data[cellId].formula || this.data[cellId].value;
      }
    });

    // Handle selection drag
    this.container.addEventListener('mousemove', (e) => {
      if (this.isSelecting) {
        const cell = (e.target as HTMLElement).closest('.cell');
        if (cell && cell !== this.selectionEnd) {
          this.selectionEnd = cell;
          this.updateSelection();
        }
      }
    });

    // Handle selection end
    document.addEventListener('mouseup', () => {
      this.isSelecting = false;
    });

    // Handle cell selection on click
    this.container.addEventListener('click', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell');
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
        // Update formula bar
        const cellName = document.getElementById('cell-name') as HTMLInputElement;
        const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
        cellName.value = cellId;
        formulaInput.value = this.data[cellId].formula || this.data[cellId].value;
        
        // Show computed value
        cell.textContent = this.data[cellId].computed;
      }
    });

    // Make cell editable on double click
    this.container.addEventListener('dblclick', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell');
      if (cell) {
        // Keep the selection state but make the cell editable
        cell.contentEditable = 'true';
        const cellId = cell.dataset.cellId!;
        cell.textContent = this.data[cellId].formula || this.data[cellId].value;
        cell.focus();
      }
    });

    // Handle cell blur (finish editing)
    this.container.addEventListener('blur', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell');
      if (cell) {
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
            // Update formula bar
            const cellName = document.getElementById('cell-name') as HTMLInputElement;
            const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
            cellName.value = nextCellId;
            formulaInput.value = this.data[nextCellId].formula || this.data[nextCellId].value;
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
      computed: evaluateFormula(value, this.data)
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
      const cell = (e.target as HTMLElement).closest('td');
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
      } else {
        // For normal navigation, clear selection and select single cell
        this.clearSelection();
        this.selectionStart = nextCell;
        this.selectionEnd = nextCell;
        this.selectionAnchor = nextCell;
        nextCell.classList.add('selected');
      }
      
      this.activeCell = nextCell;
      
      // Update formula bar
      const cellName = document.getElementById('cell-name') as HTMLInputElement;
      const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
      cellName.value = nextCellId;
      formulaInput.value = this.data[nextCellId].formula || this.data[nextCellId].value;
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

    // Store editing state
    const editingCell = this.container.querySelector('.cell[contenteditable="true"]');
    const editingCellId = editingCell?.dataset.cellId;
    const editingContent = editingCell?.textContent || '';
    const selection = window.getSelection();
    const range = selection?.getRangeAt(0);

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

    // Clear previous selection
    this.clearSelection();

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

    // Restore editing state
    if (editingCell && editingCellId) {
      editingCell.contentEditable = 'true';
      editingCell.textContent = editingContent;
      editingCell.focus();
      
      // Restore cursor position
      if (range && selection) {
        selection.removeAllRanges();
        selection.addRange(range);
      }
    }
  }

  private startColumnResize(e: MouseEvent, handle: HTMLElement): void {
    this.isResizing = true;
    this.resizeStartX = e.clientX;
    this.resizeElement = handle;
    const column = handle.closest('.column');
    if (column) {
      this.initialSize = column.offsetWidth;
    }
  }

  private startRowResize(e: MouseEvent, handle: HTMLElement): void {
    this.isResizing = true;
    this.resizeStartY = e.clientY;
    this.resizeElement = handle;
    const rowContainer = handle.closest('.row-container');
    if (rowContainer) {
      this.initialSize = rowContainer.offsetHeight;
    }
  }

  private handleResize(e: MouseEvent): void {
    if (!this.resizeElement) return;

    if (this.resizeElement.classList.contains('col-resize-handle')) {
      // Column resize
      const column = this.resizeElement.closest('.column');
      if (column) {
        const deltaX = e.clientX - this.resizeStartX;
        const newWidth = Math.max(50, this.initialSize + deltaX);
        column.style.width = `${newWidth}px`;
        column.style.minWidth = `${newWidth}px`;
      }
    } else if (this.resizeElement.classList.contains('row-resize-handle')) {
      // Row resize
      const rowContainer = this.resizeElement.closest('.row-container');
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
}

// Initialize the spreadsheet
new Spreadsheet('app'); 