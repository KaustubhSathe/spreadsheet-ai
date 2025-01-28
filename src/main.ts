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
  private selectionAnchor: { row: number; col: number } | null = null;
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
  private formulaDropdown: HTMLElement | null = null;
  private readonly HYPERFORMULA_FUNCTIONS: { name: string; description: string }[] = [
    // Math Functions
    { name: 'SUM', description: 'Adds up a series of numbers' },
    { name: 'AVERAGE', description: 'Calculates the arithmetic mean of numbers' },
    { name: 'COUNT', description: 'Counts the number of cells that contain numbers' },
    { name: 'MAX', description: 'Returns the largest value in a set of numbers' },
    { name: 'MIN', description: 'Returns the smallest value in a set of numbers' },
    { name: 'ROUND', description: 'Rounds a number to a specified number of digits' },
    { name: 'ROUNDDOWN', description: 'Rounds a number down toward zero' },
    { name: 'ROUNDUP', description: 'Rounds a number up away from zero' },
    { name: 'POWER', description: 'Returns the result of a number raised to a power' },
    { name: 'SQRT', description: 'Returns the square root of a number' },
    { name: 'ABS', description: 'Returns the absolute value of a number' },
    { name: 'MOD', description: 'Returns the remainder after division' },
    { name: 'PRODUCT', description: 'Multiplies all the numbers given as arguments' },
    { name: 'SUMIF', description: 'Adds the cells specified by a given criteria' },
    { name: 'MEDIAN', description: 'Returns the median of the given numbers' },

    // Logical Functions
    { name: 'IF', description: 'Makes a logical comparison between values' },
    { name: 'AND', description: 'Returns TRUE if all arguments are TRUE' },
    { name: 'OR', description: 'Returns TRUE if any argument is TRUE' },
    { name: 'NOT', description: 'Reverses the logical value of its argument' },
    { name: 'TRUE', description: 'Returns the logical value TRUE' },
    { name: 'FALSE', description: 'Returns the logical value FALSE' },
    { name: 'ISBLANK', description: 'Returns TRUE if the value is blank' },
    { name: 'ISERROR', description: 'Returns TRUE if the value is an error' },

    // Text Functions
    { name: 'CONCAT', description: 'Combines text from multiple cells into one text' },
    { name: 'LEFT', description: 'Returns the leftmost characters from a text value' },
    { name: 'RIGHT', description: 'Returns the rightmost characters from a text value' },
    { name: 'MID', description: 'Returns a specific number of characters from a text string' },
    { name: 'LEN', description: 'Returns the number of characters in a text string' },
    { name: 'LOWER', description: 'Converts text to lowercase' },
    { name: 'UPPER', description: 'Converts text to uppercase' },
    { name: 'TRIM', description: 'Removes spaces from text' },
    { name: 'REPLACE', description: 'Replaces characters within text' },
    { name: 'FIND', description: 'Finds one text value within another (case-sensitive)' },

    // Date & Time Functions
    { name: 'DATE', description: 'Returns the serial number of a particular date' },
    { name: 'DAY', description: 'Converts a serial number to a day of the month' },
    { name: 'MONTH', description: 'Converts a serial number to a month' },
    { name: 'YEAR', description: 'Converts a serial number to a year' },
    { name: 'NOW', description: 'Returns the current date and time' },
    { name: 'TODAY', description: 'Returns the current date' },
    { name: 'TIME', description: 'Returns the serial number of a particular time' },
    { name: 'HOUR', description: 'Converts a serial number to an hour' },
    { name: 'MINUTE', description: 'Converts a serial number to a minute' },
    { name: 'SECOND', description: 'Converts a serial number to a second' },

    // Statistical Functions
    { name: 'COUNTA', description: 'Counts how many values are in a list of arguments' },
    { name: 'COUNTBLANK', description: 'Counts empty cells in a specified range' },
    { name: 'COUNTIF', description: 'Counts cells that meet a criteria' },
    { name: 'AVERAGEA', description: 'Calculates average including text and logical values' },

    // Trigonometric Functions
    { name: 'SIN', description: 'Returns the sine of an angle' },
    { name: 'COS', description: 'Returns the cosine of an angle' },
    { name: 'TAN', description: 'Returns the tangent of an angle' },
    { name: 'ASIN', description: 'Returns the arcsine of a number' },
    { name: 'ACOS', description: 'Returns the arccosine of a number' },
    { name: 'ATAN', description: 'Returns the arctangent of a number' },
    { name: 'SINH', description: 'Returns the hyperbolic sine of a number' },
    { name: 'COSH', description: 'Returns the hyperbolic cosine of a number' },
    { name: 'TANH', description: 'Returns the hyperbolic tangent of a number' },

    // Information Functions
    { name: 'ISTEXT', description: 'Returns TRUE if the value is text' },
    { name: 'ISNUMBER', description: 'Returns TRUE if the value is a number' },
    { name: 'ISLOGICAL', description: 'Returns TRUE if the value is logical' },
    { name: 'ISNONTEXT', description: 'Returns TRUE if the value is not text' },
    { name: 'ISNA', description: 'Returns TRUE if the value is #N/A' },

    // Math & Trig
    { name: 'PI', description: 'Returns the value of pi' },
    { name: 'DEGREES', description: 'Converts radians to degrees' },
    { name: 'RADIANS', description: 'Converts degrees to radians' },
    { name: 'LN', description: 'Returns the natural logarithm of a number' },
    { name: 'LOG', description: 'Returns the logarithm of a number to a specified base' },
    { name: 'LOG10', description: 'Returns the base-10 logarithm of a number' },
    { name: 'EXP', description: 'Returns e raised to the power of a number' }
  ];
  private selectedFormulaIndex: number = -1;

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
    const gridContainer = document.createElement('div');
    gridContainer.className = 'grid-container';

    // Create row numbers column with corner cell
    const rowNumbers = document.createElement('div');
    rowNumbers.className = 'row-numbers';
    
    // Add corner cell
    const cornerCell = document.createElement('div');
    cornerCell.className = 'corner-cell';
    rowNumbers.appendChild(cornerCell);
    
    // Add row numbers with IDs
    for (let i = 0; i < this.rows; i++) {
      const rowContainer = document.createElement('div');
      rowContainer.className = 'row-container';
      
      const rowNumber = document.createElement('div');
      rowNumber.className = 'row-number';
      rowNumber.id = `row${i + 1}`; // Add unique ID
      rowNumber.dataset.index = i.toString(); // Add data attribute for easier selection
      rowNumber.textContent = (i + 1).toString();
      
      const rowResizeHandle = document.createElement('div');
      rowResizeHandle.className = 'row-resize-handle';
      rowResizeHandle.dataset.row = i.toString();
      
      rowContainer.appendChild(rowNumber);
      rowContainer.appendChild(rowResizeHandle);
      rowNumbers.appendChild(rowContainer);
    }
    gridContainer.appendChild(rowNumbers);

    // Create columns A-Z with IDs
    for (let col = 0; col < this.cols; col++) {
      const column = document.createElement('div');
      column.className = 'column';
      
      // Add column header with resize handle
      const headerContainer = document.createElement('div');
      headerContainer.className = 'header-container';
      
      const header = document.createElement('div');
      header.className = 'column-header';
      const colName = String.fromCharCode(65 + col);
      header.id = `col${colName}`; // Add unique ID
      header.textContent = colName;
      
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
        
        const content = document.createElement('div');
        content.className = 'cell-content';
        cell.appendChild(content);

        column.appendChild(cell);
      }
      gridContainer.appendChild(column);
    }
    
    this.container.appendChild(gridContainer);
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

      // Handle cell editing on keydown
      if (activeContent?.contentEditable === 'true') {
        // Handle Enter key
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          activeContent.blur();
          
          // Move to the next row after finishing edit
          const currentCellId = this.activeCell!.dataset.cellId!;
          const { row, col } = this.parseCellId(currentCellId);
          const nextRow = Math.min(this.rows - 1, row + 1);
          const nextCellId = getCellId(nextRow, col);
          const nextCell = document.querySelector(`[data-cell-id="${nextCellId}"]`) as HTMLElement;
          
          if (nextCell) {
            this.activeCell!.classList.remove('selected', 'active');
            nextCell.classList.add('selected', 'active');
            this.activeCell = nextCell;
            this.updateFormulaBar(nextCellId);
          }
          return;
        }

        // Handle formula dropdown on keydown
        requestAnimationFrame(() => {
          const value = activeContent.textContent || '';
          this.handleCellEdit(activeContent, value);
        });
      } else if (!e.ctrlKey && !e.altKey && !e.metaKey && e.key.length === 1) {
        // Start editing on any printable character
        if (this.activeCell) {
          const content = this.activeCell.querySelector('.cell-content') as HTMLElement;
          content.contentEditable = 'true';
          content.textContent = '';
          content.focus();
        }
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
      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (!cell) return;

      // Start selection
      this.isSelecting = true;
      
      if (e.shiftKey && this.activeCell) {
        // For shift+click, set anchor if not already set
        if (!this.selectionAnchor) {
          const { row, col } = this.parseCellId(this.activeCell.dataset.cellId!);
          this.selectionAnchor = { row, col };
        }
        // Handle shift selection
        const { row, col } = this.parseCellId(cell.dataset.cellId!);
        this.handleShiftSelection(row, col);
      } else {
        // Regular selection
        this.clearSelection();
        this.selectionAnchor = null;
        this.selectCell(cell);
        const { row, col } = this.parseCellId(cell.dataset.cellId!);
        this.selectionAnchor = { row, col };
      }

      this.selectionStart = cell;
      this.selectionEnd = cell;
    };

    const mousemoveHandler = (e: MouseEvent) => {
      if (!this.isSelecting || !this.selectionStart) return;

      const cell = (e.target as HTMLElement).closest('.cell') as HTMLElement;
      if (!cell || cell === this.selectionEnd) return;

      this.selectionEnd = cell;
      
      // Get range and select cells
      const { row: endRow, col: endCol } = this.parseCellId(cell.dataset.cellId!);
      this.handleShiftSelection(endRow, endCol);
    };

    const mouseupHandler = () => {
      this.isSelecting = false;
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

    // Add to attachEventListeners
    document.addEventListener('click', (e) => {
      if (!this.formulaDropdown?.contains(e.target as Node)) {
        this.hideFormulaDropdown();
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

  private handleCellEdit(content: HTMLElement, value: string): void {
    const cellId = content.parentElement?.dataset.cellId;
    if (!cellId) return;

    try {
      const { row, col } = this.parseCellId(cellId);
      
      // Handle formula dropdown
      if (value === '=') {
        this.showFormulaDropdown(content);
        // Just store the = without evaluating
        this.data[cellId] = {
          value: value,
          formula: value,
          computed: value
        };
        content.textContent = value;
        return;
      } else if (value.startsWith('=')) {
        const searchTerm = value.substring(1);
        this.showFormulaDropdown(content, searchTerm);
        
        // Only evaluate if the formula is complete (has closing parenthesis)
        const isCompleteFormula = value.includes('(') && value.includes(')');
        if (isCompleteFormula) {
          // Set the formula in HyperFormula
          this.hf.setCellContents({
            sheet: this.activeSheetId,
            row,
            col
          }, value);

          // Get the computed value
          const computed = this.hf.getCellValue({ sheet: this.activeSheetId, row, col });
          const displayValue = computed?.toString() || '';

          // Update our data structure
          this.data[cellId] = {
            value: value,
            formula: value,
            computed: displayValue
          };
        } else {
          // Store incomplete formula without evaluating
          this.data[cellId] = {
            value: value,
            formula: value,
            computed: value
          };
        }
        content.textContent = value;
      } else {
        this.hideFormulaDropdown();
        // Regular value handling...
        this.hf.setCellContents({ sheet: this.activeSheetId, row, col }, value);
        
        this.data[cellId] = {
          value: value,
          formula: '',
          computed: value
        };

        content.textContent = value;
      }

      // Update formula bar
      const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
      formulaInput.value = value;

    } catch (error) {
      console.error('Cell edit error:', error);
      // Don't show error while typing formula
      if (!content.contentEditable || content.contentEditable === 'false') {
        content.classList.add('error');
        setTimeout(() => content.classList.remove('error'), 2000);
        content.textContent = '#ERROR!';
      }
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
    // Skip if formula dropdown is visible
    if (this.formulaDropdown) return;

    if (e.key.startsWith('Arrow')) {
      e.preventDefault();
      
      // Get current active cell position
      const activeCell = this.activeCell;
      if (!activeCell) return;
      
      const { row, col } = this.parseCellId(activeCell.dataset.cellId!);
      let newRow = row;
      let newCol = col;

      // Calculate new position
      switch (e.key) {
        case 'ArrowUp': newRow = Math.max(0, row - 1); break;
        case 'ArrowDown': newRow = Math.min(this.rows - 1, row + 1); break;
        case 'ArrowLeft': newCol = Math.max(0, col - 1); break;
        case 'ArrowRight': newCol = Math.min(this.cols - 1, col + 1); break;
      }

      // Handle shift key for multi-select
      if (e.shiftKey) {
        if (!this.selectionAnchor) {
          this.selectionAnchor = { row, col };
        }
        this.handleShiftSelection(newRow, newCol);
      } else {
        // Clear all selections when shift is released
        this.clearSelection();
        this.selectionAnchor = null;
        const newCellId = getCellId(newRow, newCol);
        const newCell = document.querySelector(`[data-cell-id="${newCellId}"]`) as HTMLElement;
        if (newCell) {
          this.selectCell(newCell);
        }
      }
    }
  }

  private handleShiftSelection(newRow: number, newCol: number): void {
    if (!this.activeCell || !this.selectionAnchor) return;
    
    // Clear previous selection except active cell
    const selectedCells = document.querySelectorAll('.cell.selected:not(.active)');
    selectedCells.forEach(cell => cell.classList.remove('selected'));
    
    // Clear previous header highlights
    this.highlightHeaders(0, this.rows - 1, 0, this.cols - 1, false);

    // Get the range to select using the anchor point
    const minRow = Math.min(this.selectionAnchor.row, newRow);
    const maxRow = Math.max(this.selectionAnchor.row, newRow);
    const minCol = Math.min(this.selectionAnchor.col, newCol);
    const maxCol = Math.max(this.selectionAnchor.col, newCol);

    // Select cells in range
    for (let row = minRow; row <= maxRow; row++) {
      for (let col = minCol; col <= maxCol; col++) {
        const cellId = getCellId(row, col);
        const cell = document.querySelector(`[data-cell-id="${cellId}"]`);
        if (cell) {
          cell.classList.add('selected');
        }
      }
    }

    // Highlight headers for selected range
    this.highlightHeaders(minRow, maxRow, minCol, maxCol, true);

    // Update active cell
    const newCellId = getCellId(newRow, newCol);
    const newCell = document.querySelector(`[data-cell-id="${newCellId}"]`) as HTMLElement;
    if (newCell) {
      if (this.activeCell) {
        this.activeCell.classList.remove('active');
      }
      this.activeCell = newCell;
      this.activeCell.classList.add('active');
    }

    // Update formula bar with range
    const startCellId = getCellId(this.selectionAnchor.row, this.selectionAnchor.col);
    const endCellId = getCellId(newRow, newCol);
    this.updateFormulaBar(`${startCellId}:${endCellId}`);
  }

  private handleFormulaKeydown = (e: KeyboardEvent) => {
    if (!this.formulaDropdown) return;
    
    // Prevent cell navigation when formula dropdown is visible
    if (e.key.startsWith('Arrow')) {
      e.preventDefault();
      e.stopPropagation();
    }

    const items = this.formulaDropdown.querySelectorAll('.formula-item');
    
    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        e.stopPropagation();
        if (this.selectedFormulaIndex === -1) {
          this.selectedFormulaIndex = 0;
        } else {
          this.selectedFormulaIndex = Math.min(this.selectedFormulaIndex + 1, items.length - 1);
        }
        this.updateSelectedFormula();
        break;
        
      case 'ArrowUp':
        e.preventDefault();
        e.stopPropagation();
        if (this.selectedFormulaIndex === -1) {
          this.selectedFormulaIndex = items.length - 1;
        } else {
          this.selectedFormulaIndex = Math.max(this.selectedFormulaIndex - 1, 0);
        }
        this.updateSelectedFormula();
        break;
        
      case 'Tab':
      case 'Enter':
        e.preventDefault();
        e.stopPropagation();
        if (this.selectedFormulaIndex === -1) {
          this.selectedFormulaIndex = 0;
        }
        const selectedItem = items[this.selectedFormulaIndex];
        const formulaName = selectedItem.querySelector('.formula-name')?.textContent;
        if (formulaName) {
          const activeContent = document.querySelector('.cell-content[contenteditable="true"]') as HTMLElement;
          if (activeContent) {
            this.selectFormula(activeContent, formulaName);
          }
        }
        break;
        
      case 'Escape':
        e.preventDefault();
        e.stopPropagation();
        this.hideFormulaDropdown();
        break;
    }
  };

  private updateSelectedFormula(): void {
    if (!this.formulaDropdown) return;
    
    const items = this.formulaDropdown.querySelectorAll('.formula-item');
    items.forEach((item, index) => {
      if (index === this.selectedFormulaIndex) {
        item.classList.add('selected');
        item.scrollIntoView({ block: 'nearest' });
      } else {
        item.classList.remove('selected');
      }
    });
  }

  private selectFormula(inputElement: HTMLElement, formulaName: string): void {
    if (inputElement.contentEditable === 'true') {
      inputElement.textContent = `=${formulaName}()`;
      const range = document.createRange();
      const sel = window.getSelection();
      range.setStart(inputElement.firstChild!, inputElement.textContent!.length - 1);
      range.collapse(true);
      sel?.removeAllRanges();
      sel?.addRange(range);
      this.hideFormulaDropdown();
    }
  }

  private hideFormulaDropdown(): void {
    if (this.formulaDropdown) {
      document.removeEventListener('keydown', this.handleFormulaKeydown);
      this.formulaDropdown.remove();
      this.formulaDropdown = null;
    }
  }

  private clearSelection(): void {
    const selectedCells = document.querySelectorAll('.cell.selected');
    const editingCell = document.querySelector('.cell[contenteditable="true"]');
    
    selectedCells.forEach(cell => {
      if (cell !== editingCell) {
        cell.classList.remove('selected', 'active');
      }
    });

    // Clear header highlights
    this.highlightHeaders(0, this.rows - 1, 0, this.cols - 1, false);
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

  private showFormulaDropdown(inputElement: HTMLElement, searchTerm: string = ''): void {
    this.hideFormulaDropdown();
    this.selectedFormulaIndex = -1;

    const rect = inputElement.getBoundingClientRect();
    const viewport = {
      width: window.innerWidth,
      height: window.innerHeight
    };

    const dropdown = document.createElement('div');
    dropdown.className = 'formula-dropdown';
    
    // Create scrollable container for items
    const itemsContainer = document.createElement('div');
    itemsContainer.className = 'formula-items-container';
    
    // Filter functions based on search term
    const filteredFunctions = this.HYPERFORMULA_FUNCTIONS.filter(fn => 
      fn.name.toLowerCase().includes(searchTerm.toLowerCase())
    );

    if (filteredFunctions.length === 0) {
      this.hideFormulaDropdown();
      return;
    }

    // Add items to container
    filteredFunctions.forEach((fn, index) => {
      const item = document.createElement('div');
      item.className = 'formula-item';
      
      const nameSpan = document.createElement('span');
      nameSpan.className = 'formula-name';
      nameSpan.textContent = fn.name;
      
      const descSpan = document.createElement('span');
      descSpan.className = 'formula-description';
      descSpan.textContent = fn.description;
      
      item.appendChild(nameSpan);
      item.appendChild(descSpan);
      
      item.addEventListener('click', () => this.selectFormula(inputElement, fn.name));
      item.addEventListener('mouseover', () => {
        this.selectedFormulaIndex = index;
        this.updateSelectedFormula();
      });
      
      itemsContainer.appendChild(item);
    });

    // Add items container to dropdown
    dropdown.appendChild(itemsContainer);

    // Add keyboard navigation hint
    const hint = document.createElement('div');
    hint.className = 'formula-hint';
    hint.textContent = '↑↓ to navigate • Tab or Enter to select';
    dropdown.appendChild(hint);

    // Append to body and get dimensions
    document.body.appendChild(dropdown);
    this.formulaDropdown = dropdown;
    const dropdownRect = dropdown.getBoundingClientRect();

    // Calculate available space and scroll positions
    const spaceBelow = viewport.height - rect.bottom;
    const spaceAbove = rect.top;
    const spaceRight = viewport.width - rect.left;
    const scrollX = window.scrollX || document.documentElement.scrollLeft;
    const scrollY = window.scrollY || document.documentElement.scrollTop;

    // Position dropdown
    dropdown.style.position = 'fixed';
    dropdown.style.zIndex = '10000';

    // Horizontal positioning
    if (spaceRight >= dropdownRect.width) {
      // Align with cell's left edge
      dropdown.style.left = `${rect.left + scrollX}px`;
    } else {
      // Align with cell's right edge
      dropdown.style.left = `${rect.right - dropdownRect.width + scrollX}px`;
    }

    // Vertical positioning
    if (spaceBelow >= dropdownRect.height) {
      // Show below the cell
      dropdown.style.top = `${rect.bottom + scrollY}px`;
    } else if (spaceAbove >= dropdownRect.height) {
      // Show above the cell
      dropdown.style.top = `${rect.top - dropdownRect.height + scrollY}px`;
    } else {
      // If no space above or below, show below and make it scrollable
      dropdown.style.top = `${rect.bottom + scrollY}px`;
      const availableHeight = Math.max(spaceBelow, 100); // Minimum 100px height
      dropdown.style.maxHeight = `${availableHeight}px`;
    }

    document.addEventListener('keydown', this.handleFormulaKeydown, true);
  }

  private selectCell(cell: HTMLElement): void {
    if (this.activeCell) {
      this.activeCell.classList.remove('selected', 'active');
    }
    this.activeCell = cell;
    this.activeCell.classList.add('selected', 'active');
    this.updateFormulaBar(cell.dataset.cellId);
  }

  private highlightHeaders(minRow: number, maxRow: number, minCol: number, maxCol: number, highlight: boolean): void {
    // Highlight column headers
    for (let col = minCol; col <= maxCol; col++) {
      const colName = String.fromCharCode(65 + col);
      const colHeader = document.querySelector(`#col${colName}`); // e.g., #colA, #colB
      if (colHeader) {
        if (highlight) {
          colHeader.classList.add('selected');
        } else {
          colHeader.classList.remove('selected');
        }
      }
    }

    // Highlight row headers
    for (let row = minRow; row <= maxRow; row++) {
      const rowNum = row + 1;
      const rowHeader = document.querySelector(`#row${rowNum}`); // e.g., #row1, #row2
      if (rowHeader) {
        if (highlight) {
          rowHeader.classList.add('selected');
        } else {
          rowHeader.classList.remove('selected');
        }
      }
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