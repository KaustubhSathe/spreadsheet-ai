import './app.css';
import { SpreadsheetData, Sheet, Cell, SpreadSheet } from './types';
import { getCellId } from './utils/grid';
import { supabase, supabaseAnonKey } from './supabase';
import { HyperFormula } from 'hyperformula';
import { checkAuth, handleLogout } from './utils/auth';

export class Spreadsheet {
  private container: HTMLElement;
  private data: SpreadsheetData = {};
  private activeCell: HTMLElement | null = null;
  private rows: number = 50;
  private cols: number = 26;
  private spreadsheet: SpreadSheet | null = null;
  private activeSheetId: number = 0;
  private isSelecting: boolean = false;
  private selectionStart: HTMLElement | null = null;
  private selectionEnd: HTMLElement | null = null;
  private selectionAnchor: HTMLElement | null = null;
  private isResizing: boolean = false;
  private resizeStartX: number = 0;
  private resizeStartY: number = 0;
  private resizeElement: HTMLElement | null = null;
  private initialSize: number = 0;
  private isFilling: boolean = false;
  private fillStartCell: HTMLElement | null = null;
  private fillEndCell: HTMLElement | null = null;
  private title: string = '';
  private hf: HyperFormula;

  constructor(containerId: string) {
    this.container = document.getElementById('spreadsheet-container')!;
    // Initialize HyperFormula with proper config
    this.hf = HyperFormula.buildEmpty({ 
      licenseKey: 'gpl-v3',
      maxColumns: 26,
      maxRows: 50
    });
    this.activeSheetId = 1;
    this.init();
  }

  private init(): void {
    this.createSpreadsheet();
    this.setupFormulaBar();
    this.setupSheetTabs();
    this.setupProfileDropdown();
    this.attachEventListeners();
    this.loadSpreadsheet();
  }

  async loadSpreadsheet() {
    const urlParams = new URLSearchParams(window.location.search);
    const id = urlParams.get('id');
    
    if (id) {
      try {
        const { data: { session }, error: sessionError } = await supabase.auth.getSession();
        if (sessionError || !session) {
          throw new Error('No active session');
        }

        const { data: spreadsheet, error } = await supabase.functions.invoke(`get-spreadsheet?id=${id}`, {
          method: 'GET',
          headers: {
            Authorization: `Bearer ${session.access_token}`,
            'apikey': supabaseAnonKey,
          }
        });

        if (error) throw error;

        if (spreadsheet) {
          this.spreadsheet = spreadsheet;
          this.title = spreadsheet.title;
          document.querySelector('.title-input')?.setAttribute('value', this.title);
          
          // Initialize HyperFormula with all sheets
          spreadsheet.sheets.forEach((sheet, index) => {
              this.hf.addSheet(`Sheet${sheet.sheet_number}`);
          });

          // Find and set sheet with sheet_number 1 as active
          const firstSheet = spreadsheet.sheets.find(s => s.sheet_number === 1);
          if (firstSheet) {
            this.activeSheetId = 0; // HyperFormula index for first sheet
            this.data = firstSheet.data || {};
          }

          this.updateSheetTabs();
          this.renderSheet();
        }
      } catch (error) {
        console.error('Error loading spreadsheet:', error);
        window.location.href = '/dashboard';
      }
    } else {
      window.location.href = '/dashboard';
    }
  }

  private async addSheet() {
    try {
      if (!this.spreadsheet) return;

      const { data: { session }, error: sessionError } = await supabase.auth.getSession();
      if (sessionError || !session) {
        throw new Error('No active session');
      }

      const { data: sheet, error } = await supabase.functions.invoke('create-sheet', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${session.access_token}`,
          'apikey': supabaseAnonKey,
        },
        body: {
          spreadsheetId: this.spreadsheet.id
        }
      });

      if (error) throw error;

      this.spreadsheet.sheets.push(sheet);
      // Convert database sheet_number (1-based) to HyperFormula index (0-based)
      this.activeSheetId = sheet.sheet_number - 1;
      this.hf.addSheet(`Sheet${sheet.sheet_number}`);
      this.data = {};
      
      this.updateSheetTabs();
      this.renderSheet();
    } catch (error) {
      console.error('Error creating new sheet:', error);
      alert('Failed to create new sheet. Please try again.');
    }
  }

  private updateSheetTabs(): void {
    if (!this.spreadsheet) return;
    
    const tabsContainer = document.querySelector('.sheet-tabs') as HTMLElement;
    
    // Remove all existing tabs except add button
    tabsContainer.querySelectorAll('.tab').forEach(tab => tab.remove());
    
    // Sort sheets by sheet_number and add tabs
    const sortedSheets = [...this.spreadsheet.sheets].sort((a, b) => a.sheet_number - b.sheet_number);
    sortedSheets.forEach((sheet) => {
      const tab = document.createElement('div');
      // Convert database sheet_number to array index for comparison
      const sheetIndex = sheet.sheet_number - 1;
      tab.className = `tab${sheetIndex === this.activeSheetId ? ' active' : ''}`;
      // Use actual sheet_number for display
      tab.textContent = `Sheet${sheet.sheet_number}`;
      tab.dataset.sheetId = sheet.id;
      tabsContainer.appendChild(tab);
    });
  }

  private switchSheet(sheetId: string): void {
    if (!this.spreadsheet) return;
    
    const sheet = this.spreadsheet.sheets.find(s => s.id === sheetId);
    if (sheet) {
      this.activeSheetId = sheet.sheet_number - 1;
      this.data = sheet.data || {};
      this.updateSheetTabs();
      this.renderSheet();
    }
  }

  private setupFormulaBar(): void {
    const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
    const cellName = document.getElementById('cell-name') as HTMLInputElement;

    // Handle cell content changes during editing
    document.addEventListener('input', (e) => {
      const content = e.target as HTMLElement;
      if (content.classList.contains('cell-content') && content.contentEditable === 'true') {
        // Update formula bar to match cell content
        formulaInput.value = content.textContent || '';
      }
    });
  }

  private createSpreadsheet(): void {
    const createCell = (): HTMLElement => {
      // Create wrapper div
      const cellWrapper = document.createElement('div');
      cellWrapper.className = 'cell';
      
      // Create content div
      const content = document.createElement('div');
      content.className = 'cell-content';
      content.contentEditable = 'false';
      
      // Create fill handle div
      const fillHandle = document.createElement('div');
      fillHandle.className = 'fill-handle';
      
      // Append both to wrapper
      cellWrapper.appendChild(content);
      cellWrapper.appendChild(fillHandle);
      
      return cellWrapper;
    };

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
        const cell = createCell();
        const cellId = getCellId(row, col);
        cell.dataset.cellId = cellId;
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
    const keydownHandler = async (e: KeyboardEvent) => {
      const activeContent = this.activeCell?.querySelector('.cell-content') as HTMLElement;
      
      if (e.ctrlKey && e.key.toLowerCase() === 's') {
        e.preventDefault();
        await this.saveSpreadsheet();
        return;
      }

      if (e.key === 'Delete' && !this.isEditing()) {
        this.deleteSelectedCells();
        return;
      }

      if (!this.activeCell || activeContent?.contentEditable === 'true') {
        if (activeContent?.contentEditable === 'true' && e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          activeContent.blur();
          
          // Move to the next row after finishing edit
          const currentCellId = this.activeCell.dataset.cellId!;
          const { row, col } = this.parseCellId(currentCellId);
          const nextRow = Math.min(this.rows - 1, row + 1);
          const nextCellId = getCellId(nextRow, col);
          const nextCell = document.querySelector(`[data-cell-id="${nextCellId}"]`) as HTMLElement;
          
          if (nextCell) {
            // Remove selected from current cell
            this.activeCell.classList.remove('selected', 'active');
            // Select next cell
            nextCell.classList.add('selected', 'active');
            this.activeCell = nextCell;
            this.updateFormulaBar(nextCellId);
          }
        }
        return;
      }
      
      this.handleKeyNavigation(e);
    };

    const inputHandler = (e: Event) => {
      const target = e.target as HTMLElement;
      
      if (target.classList.contains('cell-content') && target.contentEditable === 'true') {
        const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
        formulaInput.value = target.textContent || '';
      }
      
      if (target.id === 'formula-input' && this.activeCell) {
        const content = this.activeCell.querySelector('.cell-content');
        if (content) {
          this.handleCellEdit(content as HTMLElement, (target as HTMLInputElement).value);
        }
      }
    };

    const mousedownHandler = (e: MouseEvent) => {
      const target = e.target as HTMLElement;
      
      const cell = target.closest('.cell') as HTMLElement;
      if (cell) {
        // First, handle any active editing
        const editingContent = document.querySelector('.cell-content[contenteditable="true"]') as HTMLElement;
        if (editingContent) {
          const editingCell = editingContent.closest('.cell') as HTMLElement;
          if (editingCell !== cell) {
            this.finishEditing(editingCell);
          }
        }

        this.isSelecting = true;
        
        if (e.shiftKey && this.activeCell) {
          // For shift + click, extend selection from active cell
          this.selectionAnchor = this.activeCell;
          this.selectionStart = this.selectionAnchor;
          this.selectionEnd = cell;
          this.updateSelection(); // Update selection immediately
        } else {
          const content = cell.querySelector('.cell-content') as HTMLElement;
          if (content.contentEditable !== 'true') {
            // Only start new selection if cell isn't being edited
            this.clearSelection();
            this.selectionStart = cell;
            this.selectionEnd = cell;
            this.selectionAnchor = cell;
            cell.classList.add('selected', 'active');
          }
        }
        
        this.activeCell = cell;
        this.updateFormulaBar();
      }
      
      if (target.classList.contains('col-resize-handle')) {
        this.startColumnResize(e, target);
      } else if (target.classList.contains('row-resize-handle')) {
        this.startRowResize(e, target);
      }
      
      const fillHandle = target.closest('.fill-handle');
      if (fillHandle) {
        e.stopPropagation();
        this.isFilling = true;
        this.fillStartCell = fillHandle.parentElement as HTMLElement;
      }
    };

    const mousemoveHandler = (e: MouseEvent) => {
      if (this.isSelecting) {
        const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
        if (cell && cell !== this.selectionEnd) {
          this.selectionEnd = cell;
          this.updateSelection();
          this.updateFormulaBar(); // Update formula bar while dragging
        }
      }
      if (this.isResizing) {
        this.handleResize(e);
      }
      if (this.isFilling) {
        const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
        if (cell && cell !== this.fillEndCell) {
          if (this.fillEndCell) {
            this.clearFillPreview();
          }
          this.fillEndCell = cell;
          this.showFillPreview();
        }
      }
    };

    const mouseupHandler = () => {
      this.isSelecting = false;
      this.isResizing = false;
      if (this.isFilling) {
        this.completeFill();
        this.isFilling = false;
        this.fillStartCell = null;
        this.fillEndCell = null;
      }
    };

    // Handle cell selection on click
    const clickHandler = (e: MouseEvent) => {
      const clickedCell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      
      // Always update dependent cells first
      this.updateDependentCells();

      if (clickedCell) {
        // Update selection and formula bar
        this.activeCell = clickedCell;
        const cellId = clickedCell.dataset.cellId!;
        const content = clickedCell.querySelector('.cell-content') as HTMLElement;
        const cellData = this.data[cellId];

        // Show computed value in cell
        if (!content.contentEditable || content.contentEditable === 'false') {
          content.textContent = cellData?.computed || '';
        }

        // Update formula bar
        const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
        const cellName = document.getElementById('cell-name') as HTMLInputElement;
        cellName.value = cellId;
        formulaInput.value = cellData?.formula || cellData?.value || '';
      }
    };

    // Make cell editable on double click
    const dblclickHandler = (e: MouseEvent) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (cell) {
        this.startEditing(cell);
      }
    };

    // Handle cell blur (finish editing)
    const blurHandler = (e: FocusEvent) => {
      const content = e.target as HTMLElement;
      if (content.classList.contains('cell-content') && content.contentEditable === 'true') {
        const cell = content.parentElement;
        if (cell) {
          const value = content.textContent || '';
          this.handleCellEdit(content as HTMLElement, value);
          content.contentEditable = 'false';
          
          // Force update of dependent cells
          requestAnimationFrame(() => {
            this.updateDependentCells();
          });
        }
      }
    };

    // Attach event listeners using consistently named handlers
    document.addEventListener('keydown', keydownHandler);
    document.addEventListener('input', inputHandler);
    document.addEventListener('mousedown', mousedownHandler);
    document.addEventListener('mousemove', mousemoveHandler);
    document.addEventListener('mouseup', mouseupHandler);
    document.addEventListener('click', clickHandler);
    document.addEventListener('dblclick', dblclickHandler);
    document.addEventListener('blur', blurHandler, true);
  }

  private startEditing(cell: HTMLElement): void {
    const content = cell.querySelector('.cell-content') as HTMLElement;
    const cellId = cell.dataset.cellId!;
    const cellData = this.data[cellId];
    
    // Show formula if exists, otherwise show value
    const editValue = cellData?.formula || cellData?.value || '';
    content.textContent = editValue;
    content.contentEditable = 'true';
    content.focus();
    
    // Update formula bar to match cell content
    const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
    formulaInput.value = editValue;
    
    // Select all text in the cell
    const range = document.createRange();
    range.selectNodeContents(content);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
  }

  private finishEditing(cell: HTMLElement): void {
    const content = cell.querySelector('.cell-content') as HTMLElement;
    content.contentEditable = 'false';
        const cellId = cell.dataset.cellId!;
    const newValue = content.textContent || '';
        this.updateCell(cellId, newValue);
        this.recalculateAll();
    
    // Show computed value
    content.textContent = this.data[cellId].computed;
  }

  private updateCell(cellId: string, value: unknown): void {
    // Ensure value is a string
    const stringValue = typeof value === 'string' ? value : '';
    
    this.data[cellId] = this.data[cellId] || {
      value: '',
      formula: '',
      computed: ''
    };

    const { row, col } = this.parseCellId(cellId);

    if (stringValue.startsWith('=')) {
      // Use HyperFormula for formula evaluation
      this.hf.setCellContents({ sheet: 0, row, col }, stringValue);
      const computed = this.hf.getCellValue({ sheet: 0, row, col });
      
      this.data[cellId] = {
        value: stringValue,
        formula: stringValue,
        computed: computed?.toString() || '#ERROR!'
      };
    } else {
      // Regular value
      this.hf.setCellContents({ sheet: 0, row, col }, stringValue);
      this.data[cellId] = {
        value: stringValue,
        formula: '',
        computed: stringValue
      };
    }
    
    // Update cell display
    const cell = this.container.querySelector(`[data-cell-id="${cellId}"]`);
    if (cell) {
      const content = cell.querySelector('.cell-content');
      if (content) {
        content.textContent = this.data[cellId].computed;
      }
    }
  }

  private recalculateAll(): void {
    for (const cellId in this.data) {
      if (this.data[cellId].formula) {
        this.updateCell(cellId, this.data[cellId].formula);
      }
    }
  }

  private handleCellEdit(cell: HTMLElement, value: string) {
    const cellId = cell.parentElement?.dataset.cellId;
    if (!cellId) return;

    try {
      const { row, col } = this.parseCellId(cellId);
      
      if (value.startsWith('=')) {
        // Set the formula in HyperFormula
        this.hf.setCellContents({
          sheet: 0,
          row,
          col
        }, value);

        // Get the computed value
        const computed = this.hf.getCellValue({ sheet: 0, row, col });
        const displayValue = computed?.toString() || '#ERROR!';

        // Update our data structure
        this.data[cellId] = {
          value: value,
          formula: value,
          computed: displayValue
        };

        // Update cell display with computed value
        cell.textContent = displayValue;

        // Update formula bar with formula
        const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
        formulaInput.value = value;

        // Update all cells that might have formulas
        this.updateDependentCells();
      } else {
        // Regular value
        this.hf.setCellContents({ sheet: 0, row, col }, value);
        
        // Update our data structure
        this.data[cellId] = {
          value: value,
          formula: '',
          computed: value
        };

        // Update both cell and formula bar with value
        cell.textContent = value;
        const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
        formulaInput.value = value;

        // Update any cells that might depend on this one
        this.updateDependentCells();
      }
    } catch (error) {
      console.error('Cell edit error:', error);
      cell.classList.add('error');
      setTimeout(() => cell.classList.remove('error'), 2000);
      
      // Show error in cell and formula bar
      cell.textContent = '#ERROR!';
      const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
      formulaInput.value = value;
    }
  }

  private updateDependentCells(): void {
    // Get all formulas from our data
    const formulaCells = Object.entries(this.data)
      .filter(([_, cellData]) => cellData.formula)
      .map(([cellId, _]) => {
        const { row, col } = this.parseCellId(cellId);
        return { cellId, row, col };
      });

      console.log("formulaCells", formulaCells)

    // Only update cells that contain formulas
    formulaCells.forEach(({ cellId, row, col }) => {
      const computed = this.hf.getCellValue({ sheet: 0, row, col });
      if (computed !== null) {
        const computedValue = computed.toString();
        const cell = document.querySelector(`[data-cell-id="${cellId}"] .cell-content`) as HTMLElement;
        if (cell) {
          cell.textContent = computedValue;
          this.data[cellId].computed = computedValue;
        }
      }
    });
  }

  private parseCellId(cellId: string): { row: number, col: number } {
    const colStr = cellId.match(/[A-Z]+/)?.[0] || '';
    const rowStr = cellId.match(/\d+/)?.[0] || '';
    
    const col = this.columnNameToNumber(colStr);
    const row = parseInt(rowStr) - 1;
    
    return { row, col };
  }

  private columnNameToNumber(name: string): number {
    let sum = 0;
    for (let i = 0; i < name.length; i++) {
      sum *= 26;
      sum += name.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
    }
    return sum - 1;
  }

  private handleKeyNavigation(e: KeyboardEvent): void {
    const activeContent = this.activeCell?.querySelector('.cell-content') as HTMLElement;
    // Skip navigation if we're editing a cell
    if (!this.activeCell || activeContent?.contentEditable === 'true') {
      // Handle Enter and Tab while editing
      if ((e.key === 'Enter' && !e.shiftKey) || e.key === 'Tab') {
        e.preventDefault();
        // Finish editing current cell
        const value = activeContent?.textContent || '';
        this.handleCellEdit(activeContent, value);
        activeContent.contentEditable = 'false';
        
        // Force update of dependent cells
        requestAnimationFrame(() => {
          this.updateDependentCells();
        });

        // Continue with navigation
        const currentCellId = this.activeCell.dataset.cellId!;
        const { row, col } = this.parseCellId(currentCellId);
        
        let nextRow = row;
        let nextCol = col;
        
        if (e.key === 'Enter') {
          nextRow = Math.min(this.rows - 1, row + 1);
        } else if (e.key === 'Tab') {
          nextCol = e.shiftKey ? Math.max(0, col - 1) : Math.min(this.cols - 1, col + 1);
          if (nextCol === col) {
            nextRow = e.shiftKey ? Math.max(0, row - 1) : Math.min(this.rows - 1, row + 1);
            nextCol = e.shiftKey ? this.cols - 1 : 0;
          }
        }

        const nextCellId = getCellId(nextRow, nextCol);
        const nextCell = document.querySelector(`[data-cell-id="${nextCellId}"]`) as HTMLElement;
        if (nextCell) {
          this.activeCell.classList.remove('selected', 'active');
          nextCell.classList.add('selected', 'active');
          this.activeCell = nextCell;
          this.updateFormulaBar(nextCellId);
        }
      }
      return;
    }
    
    // Rest of navigation code...
  }

  private clearSelection(): void {
    const selectedCells = document.querySelectorAll('.cell.selected');
    const editingCell = document.querySelector('.cell[contenteditable="true"]');
    
    selectedCells.forEach(cell => {
      if (cell !== editingCell) {
        cell.classList.remove('selected', 'active'); // Remove both classes
      }
    });
  }

  private updateSelection(): void {
    if (!this.selectionStart || !this.selectionEnd) return;

    this.clearSelection();

    const startId = this.selectionStart.dataset.cellId!;
    const endId = this.selectionEnd.dataset.cellId!;

    const { col: startCol, row: startRow } = this.parseCellId(startId);
    const { col: endCol, row: endRow } = this.parseCellId(endId);

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
          // Make the last cell in the selection active
          if (row === endRow && col === endCol) {
            cell.classList.add('active');
            this.activeCell = cell as HTMLElement;
          }
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

  private updateFormulaBar(cellId?: string): void {
    const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
    const cellName = document.getElementById('cell-name') as HTMLInputElement;

    if (!this.activeCell && !cellId) return;

    const activeCellId = cellId || this.activeCell?.dataset.cellId;
    if (!activeCellId) return;

    cellName.value = activeCellId;
    
    const cellData = this.data[activeCellId];
    if (cellData) {
      formulaInput.value = cellData.formula || cellData.value || '';
    } else {
      formulaInput.value = '';
    }
  }

  private clearFillPreview(): void {
    const previewCells = this.container.querySelectorAll('.fill-preview');
    previewCells.forEach(cell => cell.classList.remove('fill-preview'));
  }

  private showFillPreview(): void {
    if (!this.fillStartCell || !this.fillEndCell) return;

    // First clear any existing selection highlight except for the start cell
    const selectedCells = this.container.querySelectorAll('.cell.selected');
    selectedCells.forEach(cell => {
      if (cell !== this.fillStartCell) {
        cell.classList.remove('selected', 'active');
      }
    });

    const startId = this.fillStartCell.dataset.cellId!;
    const endId = this.fillEndCell.dataset.cellId!;
    const { col: startCol, row: startRow } = this.parseCellId(startId);
    const { col: endCol, row: endRow } = this.parseCellId(endId);

    // Determine if we're filling horizontally or vertically
    const isHorizontal = Math.abs(endCol - startCol) > Math.abs(endRow - startRow);

    if (isHorizontal) {
      // Horizontal fill - keep same row
      const minCol = Math.min(startCol, endCol);
      const maxCol = Math.max(startCol, endCol);
      
      for (let col = minCol; col <= maxCol; col++) {
        const cellId = getCellId(startRow, col);
        const cell = this.container.querySelector(`[data-cell-id="${cellId}"]`);
        if (cell && cell !== this.fillStartCell) {
          cell.classList.add('fill-preview');
        }
      }
    } else {
      // Vertical fill - keep same column
      const minRow = Math.min(startRow, endRow);
      const maxRow = Math.max(startRow, endRow);
      
      for (let row = minRow; row <= maxRow; row++) {
        const cellId = getCellId(row, startCol);
        const cell = this.container.querySelector(`[data-cell-id="${cellId}"]`);
        if (cell && cell !== this.fillStartCell) {
          cell.classList.add('fill-preview');
        }
      }
    }
  }

  private completeFill(): void {
    if (!this.fillStartCell || !this.fillEndCell) return;

    const startId = this.fillStartCell.dataset.cellId!;
    const endId = this.fillEndCell.dataset.cellId!;
    const startValue = this.data[startId].value;
    
    const isNumber = startValue !== '' && !isNaN(Number(startValue));
    const startNum = Number(startValue);

    const { col: startCol, row: startRow } = this.parseCellId(startId);
    const { col: endCol, row: endRow } = this.parseCellId(endId);
    
    // Determine if we're filling horizontally or vertically
    const isHorizontal = Math.abs(endCol - startCol) > Math.abs(endRow - startRow);

    if (isHorizontal) {
      // Horizontal fill
      const minCol = Math.min(startCol, endCol);
      const maxCol = Math.max(startCol, endCol);
      
      for (let col = minCol; col <= maxCol; col++) {
        const cellId = getCellId(startRow, col);
        if (cellId !== startId) {
          let value = startValue;
          
          if (isNumber) {
            const offset = col - startCol;
            value = (startNum + offset).toString();
          } else if (startValue === '') {
            value = '';
          }
          
          this.updateCell(cellId, value);
        }
      }
    } else {
      // Vertical fill
      const minRow = Math.min(startRow, endRow);
      const maxRow = Math.max(startRow, endRow);
      
      for (let row = minRow; row <= maxRow; row++) {
        const cellId = getCellId(row, startCol);
        if (cellId !== startId) {
          let value = startValue;
          
          if (isNumber) {
            const offset = row - startRow;
            value = (startNum + offset).toString();
          } else if (startValue === '') {
            value = '';
          }
          
        this.updateCell(cellId, value);
        }
      }
    }

    this.clearFillPreview();
    
    // Update dependent cells after fill
    requestAnimationFrame(() => {
      this.updateDependentCells();
    });
  }

  // Add this new method
  private setupProfileDropdown(): void {
    const accountIcon = document.querySelector('.account-icon');
    const dropdown = document.querySelector('.profile-dropdown');
    
    if (!accountIcon || !dropdown) return;

    // Toggle dropdown on click
    accountIcon.addEventListener('click', (e) => {
      e.stopPropagation();
      dropdown.classList.toggle('show');
    });

    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
      if (!dropdown.contains(e.target as Node)) {
        dropdown.classList.remove('show');
      }
    });

    // Use handleLogout function for logout
    const logoutButton = dropdown.querySelector('.logout-button');
    if (logoutButton) {
      logoutButton.addEventListener('click', handleLogout);
    }
  }

  private deleteSelectedCells(): void {
    const selectedCells = this.container.querySelectorAll('.cell.selected .cell-content');
    selectedCells.forEach(cell => {
      cell.textContent = '';
      const cellId = (cell.parentElement as HTMLElement).dataset.cellId!;
      this.data[cellId] = {
        value: '',
        formula: '',
        computed: ''
      };
    });
  }

  private isEditing(): boolean {
    return !!this.container.querySelector('.cell-content[contenteditable="true"]');
  }

  private async saveSpreadsheet(): Promise<void> {
    if (!this.spreadsheet) return;
    
    try {
      const { data: { session }, error: sessionError } = await supabase.auth.getSession();
      if (sessionError || !session) {
        throw new Error('No active session');
      }

      const currentSheet = this.spreadsheet.sheets.find(s => s.sheet_number - 1 === this.activeSheetId);
      if (!currentSheet) throw new Error('Current sheet not found');

      const { error } = await supabase.functions.invoke('save-spreadsheet', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${session.access_token}`,
          'apikey': supabaseAnonKey,
        },
        body: {
          spreadsheetId: this.spreadsheet.id,
          title: this.title,
          sheetId: currentSheet.id,
          data: this.data
        }
      });

      if (error) throw error;
      
      // Show save indicator
      const saveIndicator = document.createElement('div');
      saveIndicator.textContent = 'All changes saved';
      saveIndicator.style.cssText = `
        position: fixed;
        top: 16px;
        right: 16px;
        background: #323232;
        color: white;
        padding: 8px 16px;
        border-radius: 4px;
        font-size: 14px;
        opacity: 0;
        transition: opacity 0.3s;
      `;
      document.body.appendChild(saveIndicator);
      
      // Fade in
      setTimeout(() => saveIndicator.style.opacity = '1', 0);
      // Fade out and remove
      setTimeout(() => {
        saveIndicator.style.opacity = '0';
        setTimeout(() => saveIndicator.remove(), 300);
      }, 2000);

    } catch (error) {
      console.error('Error saving spreadsheet:', error);
      alert('Failed to save spreadsheet. Please try again.');
    }
  }

  private renderSheet(): void {
    if (!this.spreadsheet) return;
    
    const sheet = this.spreadsheet.sheets[this.activeSheetId];
    if (!sheet) return;

    // Clear existing cells
    const cells = this.container.querySelectorAll('.cell');
    cells.forEach(cell => {
      const cellId = (cell as HTMLElement).dataset.cellId!;
      const content = cell.querySelector('.cell-content') as HTMLElement;
      content.textContent = this.data[cellId]?.computed || '';
    });

    // Also update HyperFormula with the new sheet's data
    Object.entries(this.data).forEach(([cellId, cellData]) => {
      const { row, col } = this.parseCellId(cellId);
      if (cellData.formula) {
        this.hf.setCellContents({ sheet: this.activeSheetId, row, col }, cellData.formula);
      } else {
        this.hf.setCellContents({ sheet: this.activeSheetId, row, col }, cellData.value);
      }
    });
  }

  private setupSheetTabs(): void {
    const sheetTabs = document.querySelector('.sheet-tabs') as HTMLElement;
    const addSheetButton = document.querySelector('.add-sheet') as HTMLElement;

    // Add sheet button handler
    addSheetButton.addEventListener('click', () => {
      this.addSheet();
    });

    // Handle sheet switching
    sheetTabs.addEventListener('click', (e) => {
      const tab = (e.target as HTMLElement).closest('.tab') as HTMLElement;
      if (tab && !tab.classList.contains('active')) {
        const sheetId = tab.dataset.sheetId!;
        this.switchSheet(sheetId);
      }
    });
  }
}

// Initialize app only after auth check
if (typeof window !== 'undefined' && process.env.NODE_ENV !== 'test') {
  checkAuth().then(isAuthenticated => {
    if (isAuthenticated) {
new Spreadsheet('app'); 
    }
  });
} 