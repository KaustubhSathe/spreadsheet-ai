import './app.css';
import { SpreadsheetData, Sheet, Cell } from './types';
import { getCellId } from './utils';
import { supabase, supabaseAnonKey } from './supabase';
import { HyperFormula } from 'hyperformula';

// Add auth check at the start
async function checkAuth() {
  const { data: { user } } = await supabase.auth.getUser();
  if (!user) {
    window.location.href = '/';
    return false;
  }
  return true;
}

// Add this function after checkAuth()
async function handleLogout() {
  const { error } = await supabase.auth.signOut();
  if (!error) {
    window.location.href = '/';
  }
}

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
    // Add a sheet and store its ID
    const sheetId = this.hf.addSheet('Sheet1');
    this.init();
    this.loadSpreadsheet();
    this.setupFormulaBar();
    this.setupSheetTabs();
    this.setupProfileDropdown();
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

        const { data: spreadsheets, error } = await supabase.functions.invoke(`get-spreadsheet?id=${id}`, {
          method: 'GET',
          headers: {
            Authorization: `Bearer ${session.access_token}`,
            'apikey': supabaseAnonKey,
          }
        }) as { data: Sheet[], error: any };

        if (error) throw error;

        const spreadsheet = spreadsheets[0]; // Get first item since we know it's an array with one item
        if (spreadsheet) {
          this.title = spreadsheet.title;
          document.querySelector('.title-input')?.setAttribute('value', this.title);
          
          if (spreadsheet.data) {
            // First register all cells with HyperFormula
            Object.entries(spreadsheet.data).forEach(([cellId, cellData]: [string, Cell]) => {
              const { row, col } = this.parseCellId(cellId);
              if (cellData.formula) {
                this.hf.setCellContents({ sheet: 0, row, col }, cellData.formula);
              } else {
                this.hf.setCellContents({ sheet: 0, row, col }, cellData.value);
              }
            });

            // Then update UI with computed values
            Object.entries(spreadsheet.data).forEach(([cellId, cellData]: [string, Cell]) => {
              const cell = document.querySelector(`[data-cell-id="${cellId}"] .cell-content`);
              if (cell) {
                const { row, col } = this.parseCellId(cellId);
                const computed = this.hf.getCellValue({ sheet: 0, row, col });
                cell.textContent = computed?.toString() || '';
                this.data[cellId] = {
                  value: cellData.value,
                  formula: cellData.formula,
                  computed: computed?.toString() || ''
                };
              }
            });
          }
        }
      } catch (error) {
        console.error('Error loading spreadsheet:', error);
        window.location.href = '/dashboard';
      }
    } else {
      window.location.href = '/dashboard';
    }
  }

  private init(): void {
    // Create initial sheet
    this.addSheet();
    this.createSpreadsheet();
    this.attachEventListeners();
  }

  private async addSheet() {
    try {
      const { data: { user }, error: userError } = await supabase.auth.getUser();
      if (userError || !user) {
        throw new Error('No active session');
      }

    const sheetNumber = this.sheets.length + 1;
    const newSheet: Sheet = {
      id: crypto.randomUUID(),
        title: `Sheet${sheetNumber}`,
        data: {},
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString(),
        deleted_at: null,
        user_id: user.id
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
    } catch (error) {
      console.error('Error creating new sheet:', error);
      window.location.href = '/';
    }
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
      tab.textContent = sheet.title;
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
        const cell = this.createCell();
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

  private createCell(): HTMLElement {
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
  }

  private attachEventListeners(): void {
    // Handle selection start
    document.addEventListener('mousedown', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
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
    });

    // Handle selection drag
    document.addEventListener('mousemove', (e) => {
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
    document.addEventListener('click', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (cell) {
        const content = cell.querySelector('.cell-content') as HTMLElement;
        // Don't handle click if we're already editing
        if (content.contentEditable === 'true') {
          return;
        }

        // Remove selected and active classes from previously selected cells
        const prevSelected = document.querySelectorAll('.cell.selected');
        prevSelected.forEach(cell => {
          cell.classList.remove('selected', 'active');
        });
        
        this.activeCell = cell;
        cell.classList.add('selected', 'active');
        
        const cellId = cell.dataset.cellId!;
        this.updateFormulaBar(cellId);
      }
    });

    // Make cell editable on double click
    document.addEventListener('dblclick', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (cell) {
        this.startEditing(cell);
      }
    });

    // Handle cell content changes during editing
    document.addEventListener('input', (e) => {
      const content = e.target as HTMLElement;
      if (content.classList.contains('cell-content') && content.contentEditable === 'true') {
        const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
        formulaInput.value = content.textContent || '';
      }
    });

    // Handle cell blur (finish editing)
    document.addEventListener('blur', (e) => {
      const content = e.target as HTMLElement;
      if (content.classList.contains('cell-content') && content.contentEditable === 'true') {
        const cell = content.parentElement;
        if (cell) {
          const value = content.textContent || '';
          this.handleCellEdit(content as HTMLElement, value);
          content.contentEditable = 'false';
        }
      }
    }, true);

    // Move keydown listener to document level
    document.addEventListener('keydown', (e) => {
      // If no cell is selected or we're editing, don't handle navigation
      const activeContent = this.activeCell?.querySelector('.cell-content') as HTMLElement;
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
      
      // Handle navigation
      this.handleKeyNavigation(e);

      // Add keyboard event listener for delete
      if (e.key === 'Delete' && !this.isEditing()) {
        this.deleteSelectedCells();
      }
    });

    // Add resize event listeners
    document.addEventListener('mousedown', (e) => {
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

    // Handle fill handle events
    document.addEventListener('mousedown', (e) => {
      const fillHandle = (e.target as HTMLElement).closest('.fill-handle');
      if (fillHandle) {
        e.stopPropagation();
        this.isFilling = true;
        this.fillStartCell = fillHandle.parentElement as HTMLElement;
      }
    });

    document.addEventListener('mousemove', (e) => {
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
    });

    document.addEventListener('mouseup', () => {
      if (this.isFilling) {
        this.completeFill();
        this.isFilling = false;
        this.fillStartCell = null;
        this.fillEndCell = null;
      }
    });

    // Add save shortcut
    document.addEventListener('keydown', async (e) => {
      if (e.ctrlKey && e.key.toLowerCase() === 's') {  // Check for exact combination
        console.log(e)
        e.preventDefault(); // Prevent browser save dialog
        await this.saveSpreadsheet();
      }
    });

    // Update formula bar when cell is selected
    document.addEventListener('click', (e) => {
      const cell = (e.target as HTMLElement).closest('td');
      if (cell) {
        this.activeCell = cell;
        const cellId = cell.dataset.cellId!;
        const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
        const cellName = document.getElementById('cell-name') as HTMLInputElement;
        
        cellName.value = cellId;
        formulaInput.value = this.data[cellId]?.formula || this.data[cellId]?.value || '';
      }
    });

    // Update cell when formula bar changes
    const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
    formulaInput.addEventListener('input', (e) => {
      if (this.activeCell) {
        const content = this.activeCell.querySelector('.cell-content');
        if (content) {
          const value = (e.target as HTMLInputElement).value;
          this.handleCellEdit(content as HTMLElement, value);
        }
      }
    });
  }

  private setupFormulaBar(): void {
    const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
    const cellName = document.getElementById('cell-name') as HTMLInputElement;

    // Update cell when formula bar changes
    formulaInput.addEventListener('input', (e) => {
      if (this.activeCell) {
        const content = this.activeCell.querySelector('.cell-content') as HTMLElement;
        const value = (e.target as HTMLInputElement).value;
        
        // Update both cell content and formula bar
        content.textContent = value;
        this.handleCellEdit(content, value);
      }
    });

    // Update formula bar when cell is clicked
    document.addEventListener('click', (e) => {
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (cell) {
        this.activeCell = cell;
        const cellId = cell.dataset.cellId!;
        cellName.value = cellId;
        
        // Show formula if exists, otherwise show computed value
        const cellData = this.data[cellId];
        formulaInput.value = cellData?.formula || cellData?.value || '';
        
        // Update cell content to show computed value
        const content = cell.querySelector('.cell-content') as HTMLElement;
        content.textContent = cellData?.computed || '';
      }
    });

    // Handle cell content changes during editing
    document.addEventListener('input', (e) => {
      const content = e.target as HTMLElement;
      if (content.classList.contains('cell-content') && content.contentEditable === 'true') {
        // Update formula bar to match cell content
        formulaInput.value = content.textContent || '';
      }
    });
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

    this.data[cellId].value = stringValue;
    if (stringValue.startsWith('=')) {
      this.data[cellId].formula = stringValue;
      const computed = this.evaluateFormula(stringValue, this.data);
      this.data[cellId].computed = typeof computed === 'string' ? computed : '';
    } else {
      this.data[cellId].formula = '';
      this.data[cellId].computed = stringValue;
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

    console.log(formulaCells)

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
    if (!this.activeCell || activeContent?.contentEditable === 'true') return;
    
    const currentCellId = this.activeCell.dataset.cellId!;
    const { row, col } = this.parseCellId(currentCellId);
    
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
        const { col: startCol, row: startRow } = this.parseCellId(startId);
        const { col: endCol, row: endRow } = this.parseCellId(endId);
        
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

    // Handle logout
    const logoutButton = dropdown.querySelector('.logout-button');
    if (logoutButton) {
      logoutButton.addEventListener('click', async () => {
        const { error } = await supabase.auth.signOut();
        if (!error) {
          window.location.href = '/';
        }
      });
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
    const urlParams = new URLSearchParams(window.location.search);
    const id = urlParams.get('id');
    
    if (!id) return;

    try {
      const { error } = await supabase
        .from('spreadsheets')
        .update({
          title: this.title ?? 'Untitled spreadsheet',
          data: this.data ?? {},
          updated_at: new Date().toISOString()
        })
        .eq('id', id);

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
}

// Initialize app only after auth check
if (typeof window !== 'undefined' && process.env.NODE_ENV !== 'test') {
  checkAuth().then(isAuthenticated => {
    if (isAuthenticated) {
new Spreadsheet('app'); 
    }
  });
} 