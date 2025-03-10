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
    // Add click handler for spreadsheet icon
    const spreadsheetIcon = document.querySelector('.spreadsheet-icon');
    if (spreadsheetIcon) {
      spreadsheetIcon.addEventListener('click', () => {
        window.location.href = '/dashboard';
      });
      spreadsheetIcon.style.cursor = 'pointer';
    }

    this.createSpreadsheet();
    this.setupFormulaBar();
    this.setupSheetTabs();
    this.setupProfileDropdown();
    this.setupFileMenu();
    this.setupEditMenu();
    this.setupViewMenu();
    this.setupInsertMenu();
    this.setupFormatMenu();
    this.setupHelpMenu();
    this.setupTitleInput();
    this.setupGlobalClickHandler();
    this.attachEventListeners();
    this.loadSpreadsheet();
    this.setupClipboardHandlers();
    this.setupToolbar();
    this.setupThemeToggle();
    this.setupResizeHandlers();
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

            // Apply stored styles to all cells
            Object.entries(this.data).forEach(([cellId, cellData]) => {
              const cell = document.querySelector(`[data-cell-id="${cellId}"]`);
              if (cell && cellData.styles) {
                const contentDiv = cell.querySelector('.cell-content') as HTMLElement;
                if (contentDiv) {
                  Object.entries(cellData.styles).forEach(([property, value]) => {
                    contentDiv.style[property] = value;
                  });
                }
              }
            });
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
        const cell = this.createCell(row, col);
        column.appendChild(cell);
      }
      gridContainer.appendChild(column);
    }

    this.container.appendChild(gridContainer);
  }

  private createCell(row: number, col: number): HTMLElement {
        const cell = document.createElement('div');
    const cellId = this.getCellId(col, row);
        cell.className = 'cell';
    cell.setAttribute('data-cell-id', cellId);
        
        const content = document.createElement('div');
        content.className = 'cell-content';
    
    // Apply stored styles if they exist
    if (this.data[cellId]?.styles) {
      Object.entries(this.data[cellId].styles).forEach(([property, value]) => {
        content.style[property] = value;
      });
    }

    // Set content if exists
    if (this.data[cellId]) {
      content.textContent = this.data[cellId].value;
    }

    cell.appendChild(content);
    return cell;
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
    let selectedCells = document.querySelectorAll('.cell.selected:not(.active)');
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

    // Check fonts and font sizes of selected cells
    selectedCells = document.querySelectorAll('.cell.selected');
    const fontSelect = document.querySelector('.font-select') as HTMLSelectElement;
    const fontSizeSelect = document.querySelector('.font-size') as HTMLSelectElement;

    if (selectedCells.length > 0) {
      let commonFont: string | null = null;
      let commonSize: string | null = null;
      let hasFontConflict = false;
      let hasSizeConflict = false;

      selectedCells.forEach((cell) => {
        const contentDiv = cell.querySelector('.cell-content') as HTMLElement;
        if (contentDiv) {
          // Check font
          const currentFont = contentDiv.style.fontFamily || 'Arial';
          const cleanFont = currentFont.replace(/['"]/g, '');
          if (commonFont === null) {
            commonFont = cleanFont;
          } else if (commonFont !== cleanFont) {
            hasFontConflict = true;
          }

          // Check font size
          const currentSize = contentDiv.style.fontSize || '11px';
          const sizeNumber = parseInt(currentSize);
          if (commonSize === null) {
            commonSize = sizeNumber.toString();
          } else if (commonSize !== sizeNumber.toString()) {
            hasSizeConflict = true;
          }
        }
      });

      // Update dropdowns
      if (fontSelect) {
        fontSelect.value = hasFontConflict ? '' : (commonFont || 'Arial');
      }
      if (fontSizeSelect) {
        fontSizeSelect.value = hasSizeConflict ? '' : (commonSize || '11');
      }
    }
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

    // Add theme toggle before logout button
    const themeToggle = document.createElement('button');
    themeToggle.className = 'theme-toggle-button';
    themeToggle.innerHTML = `
      <span class="material-icons">dark_mode</span>
      <span>Dark theme</span>
      <span class="material-icons theme-check">check</span>
    `;

    // Initialize theme from localStorage
    const isDarkTheme = localStorage.getItem('darkTheme') === 'true';
    document.documentElement.classList.toggle('dark-theme', isDarkTheme);
    themeToggle.querySelector('.theme-check')?.classList.toggle('active', isDarkTheme);

    themeToggle.addEventListener('click', () => {
      document.documentElement.classList.toggle('dark-theme');
      const isNowDark = document.documentElement.classList.contains('dark-theme');
      localStorage.setItem('darkTheme', isNowDark.toString());
      themeToggle.querySelector('.theme-check')?.classList.toggle('active', isNowDark);
    });

    // Insert theme toggle before logout button
    const logoutButton = dropdown.querySelector('.logout-button');
    if (logoutButton) {
      dropdown.insertBefore(themeToggle, logoutButton);
    }

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
    // Remove active state from previous cell
    this.activeCell?.classList.remove('active');
    
    // Set new active cell
    this.activeCell = cell;
    cell.classList.add('active', 'selected');

    // Update cell name display
    const cellId = cell.getAttribute('data-cell-id');
    if (cellId) {
      const cellNameInput = document.getElementById('cell-name') as HTMLInputElement;
      if (cellNameInput) {
        cellNameInput.value = cellId;
      }

      // Update formula input
      const formulaInput = document.getElementById('formula-input') as HTMLInputElement;
      if (formulaInput) {
        formulaInput.value = this.data[cellId]?.formula || '';
      }

      // Update font select and font size to match cell's styles
      const contentDiv = cell.querySelector('.cell-content') as HTMLElement;
      if (contentDiv) {
        const fontSelect = document.querySelector('.font-select') as HTMLSelectElement;
        const fontSizeSelect = document.querySelector('.font-size') as HTMLSelectElement;

        if (fontSelect) {
          const currentFont = contentDiv.style.fontFamily || 'Arial';
          const cleanFont = currentFont.replace(/['"]/g, '');
          fontSelect.value = cleanFont;
        }

        if (fontSizeSelect) {
          const currentSize = contentDiv.style.fontSize || '11px';
          const sizeNumber = parseInt(currentSize);
          fontSizeSelect.value = sizeNumber.toString();
        }
      }
    }

    // Sync format buttons with cell styles
    this.syncFormatButtons(cell);
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

  private setupFileMenu(): void {
    const fileMenu = document.querySelector('.menu-item:first-child');
    if (!fileMenu) return;

    const dropdown = document.createElement('div');
    dropdown.className = 'menu-dropdown';
    
    const menuItems = [
      { label: 'New spreadsheet', icon: 'add', action: () => window.location.href = '/dashboard' },
      { label: 'Open a CSV', icon: 'upload_file', action: () => this.openCSVFile() },
      { label: 'Make a copy', icon: 'content_copy', action: this.makeSpreadsheetCopy.bind(this) },
      { label: 'Download as CSV', icon: 'download', action: this.downloadAsCSV.bind(this) },
      null, // separator
      { label: 'Rename', icon: 'drive_file_rename_outline', action: this.focusTitleInput.bind(this) },
      { label: 'Move to trash', icon: 'delete', action: this.moveToTrash.bind(this) },
      null, // separator
      { label: 'Details', icon: 'info', action: this.showDetails.bind(this) }
    ];

    menuItems.forEach(item => {
      if (!item) {
        const separator = document.createElement('div');
        separator.className = 'menu-dropdown-separator';
        dropdown.appendChild(separator);
        return;
      }

      const menuItem = document.createElement('div');
      menuItem.className = 'menu-dropdown-item';
      menuItem.innerHTML = `
        <div class="menu-item-left">
          <span class="material-icons">${item.icon}</span><span>${item.label}</span>
        </div>
      `;
      menuItem.addEventListener('click', item.action);
      dropdown.appendChild(menuItem);
    });

    fileMenu.appendChild(dropdown);

    // Add click handler to toggle menu
    fileMenu.addEventListener('click', (e) => {
      e.stopPropagation();
      document.querySelectorAll('.menu-dropdown').forEach(menu => {
        if (menu !== dropdown) {
          menu.classList.remove('active');
        }
      });
      dropdown.classList.toggle('active');
    });
  }

  private setupEditMenu(): void {
    const editMenu = document.querySelector('.menu-item:nth-child(2)');
    if (!editMenu) return;

    const dropdown = document.createElement('div');
    dropdown.className = 'menu-dropdown';

    const menuItems = [
      { 
        label: 'Undo', 
        icon: 'undo', 
        shortcut: 'Ctrl+Z'
      },
      { 
        label: 'Redo', 
        icon: 'redo', 
        shortcut: 'Ctrl+Y'
      },
      null, // separator
      { 
        label: 'Cut', 
        icon: 'content_cut', 
        shortcut: 'Ctrl+X'
      },
      { 
        label: 'Copy', 
        icon: 'content_copy', 
        shortcut: 'Ctrl+C'
      },
      { 
        label: 'Paste', 
        icon: 'content_paste', 
        shortcut: 'Ctrl+V'
      },
      null, // separator
      { 
        label: 'Delete', 
        icon: 'backspace', 
        shortcut: 'Delete'
      }
    ];

    menuItems.forEach(item => {
      if (!item) {
        const separator = document.createElement('div');
        separator.className = 'menu-dropdown-separator';
        dropdown.appendChild(separator);
        return;
      }

      const menuItem = document.createElement('div');
      menuItem.className = 'menu-dropdown-item';
      menuItem.innerHTML = `
        <div class="menu-item-left">
          <span class="material-icons">${item.icon}</span>
          ${item.label}
        </div>
        <span class="shortcut">${item.shortcut}</span>
      `;
      dropdown.appendChild(menuItem);
    });

    editMenu.appendChild(dropdown);

    // Add click handler to toggle menu
    editMenu.addEventListener('click', (e) => {
      e.stopPropagation();
      document.querySelectorAll('.menu-dropdown').forEach(menu => {
        if (menu !== dropdown) {
          menu.classList.remove('active');
        }
      });
      dropdown.classList.toggle('active');
    });
  }

  private setupViewMenu(): void {
    const viewMenu = document.querySelector('.menu-item:nth-child(3)');
    if (!viewMenu) return;

    const dropdown = document.createElement('div');
    dropdown.className = 'menu-dropdown';
    
    // Show submenu items
    const showSubmenu = document.createElement('div');
    showSubmenu.className = 'submenu-item';
    showSubmenu.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">visibility</span>
        <span>Show</span>
      </div>
      <span class="material-icons submenu-arrow">arrow_right</span>
    `;

    const showDropdown = document.createElement('div');
    showDropdown.className = 'submenu-dropdown';
    showDropdown.innerHTML = `
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span>Formula bar</span>
        </div>
        <span class="material-icons">check</span>
      </div>
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span>Toolbar</span>
        </div>
        <span class="material-icons">check</span>
      </div>
    `;
    showSubmenu.appendChild(showDropdown);

    // Add theme toggle
    const themeToggle = document.createElement('div');
    themeToggle.className = 'menu-dropdown-item';
    themeToggle.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">dark_mode</span>
        <span>Dark theme</span>
      </div>
      <span class="material-icons theme-check">check</span>
    `;

    // Initialize theme from localStorage
    const isDarkTheme = localStorage.getItem('darkTheme') === 'true';
    document.documentElement.classList.toggle('dark-theme', isDarkTheme);
    themeToggle.querySelector('.theme-check')?.classList.toggle('active', isDarkTheme);

    themeToggle.addEventListener('click', () => {
      document.documentElement.classList.toggle('dark-theme');
      const isNowDark = document.documentElement.classList.contains('dark-theme');
      localStorage.setItem('darkTheme', isNowDark.toString());
      themeToggle.querySelector('.theme-check')?.classList.toggle('active', isNowDark);
      dropdown.classList.remove('active');
    });

    // Zoom submenu items
    const zoomSubmenu = document.createElement('div');
    zoomSubmenu.className = 'submenu-item';
    zoomSubmenu.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">zoom_in</span>
        <span>Zoom</span>
      </div>
      <span class="material-icons submenu-arrow">arrow_right</span>
    `;

    const zoomDropdown = document.createElement('div');
    zoomDropdown.className = 'submenu-dropdown';
    const zoomLevels = ['50%', '75%', '90%', '100%', '125%', '150%', '200%'];
    zoomLevels.forEach(level => {
      const item = document.createElement('div');
      item.className = 'menu-dropdown-item';
      item.innerHTML = `
        <div class="menu-item-left">
          <span>${level}</span>
        </div>
        ${level === '100%' ? '<span class="material-icons">check</span>' : ''}
      `;
      zoomDropdown.appendChild(item);
    });
    zoomSubmenu.appendChild(zoomDropdown);

    // Full screen item
    const fullScreenItem = document.createElement('div');
    fullScreenItem.className = 'menu-dropdown-item';
    fullScreenItem.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">fullscreen</span>
        <span>Full screen</span>
      </div>
      <span class="shortcut">F11</span>
    `;

    // Add all items to dropdown
    dropdown.appendChild(showSubmenu);
    dropdown.appendChild(themeToggle);
    dropdown.appendChild(zoomSubmenu);
    dropdown.appendChild(document.createElement('div')).className = 'menu-dropdown-separator';
    dropdown.appendChild(fullScreenItem);

    viewMenu.appendChild(dropdown);

    // Add click handler to toggle menu
    viewMenu.addEventListener('click', (e) => {
      e.stopPropagation();
      document.querySelectorAll('.menu-dropdown').forEach(menu => {
        if (menu !== dropdown) {
          menu.classList.remove('active');
        }
      });
      dropdown.classList.toggle('active');
    });
  }

  private setupGlobalClickHandler(): void {
    // Close all menus when clicking outside
    document.addEventListener('click', () => {
      document.querySelectorAll('.menu-dropdown').forEach(menu => {
        menu.classList.remove('active');
      });
    });
  }

  private openCSVFile(): void {
    // Implementation of openCSVFile method
  }

  private makeSpreadsheetCopy(): void {
    // Implementation of makeSpreadsheetCopy method
  }

  private downloadAsCSV(): void {
    // Implementation of downloadAsCSV method
  }

  private focusTitleInput(): void {
    // Implementation of focusTitleInput method
  }

  private moveToTrash(): void {
    // Implementation of moveToTrash method
  }

  private showDetails(): void {
    // Implementation of showDetails method
  }

  private setupInsertMenu(): void {
    const insertMenu = document.querySelector('.menu-item:nth-child(4)');
    if (!insertMenu) return;

    const dropdown = document.createElement('div');
    dropdown.className = 'menu-dropdown';

    // Rows submenu
    const rowsSubmenu = document.createElement('div');
    rowsSubmenu.className = 'submenu-item';
    rowsSubmenu.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">table_rows</span>
        <span>Rows</span>
      </div>
      <span class="material-icons submenu-arrow">arrow_right</span>
    `;

    const rowsDropdown = document.createElement('div');
    rowsDropdown.className = 'submenu-dropdown';
    rowsDropdown.innerHTML = `
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span>1 row above</span>
        </div>
      </div>
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span>1 row below</span>
        </div>
      </div>
    `;
    rowsSubmenu.appendChild(rowsDropdown);

    // Columns submenu
    const columnsSubmenu = document.createElement('div');
    columnsSubmenu.className = 'submenu-item';
    columnsSubmenu.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">view_column</span>
        <span>Columns</span>
      </div>
      <span class="material-icons submenu-arrow">arrow_right</span>
    `;

    const columnsDropdown = document.createElement('div');
    columnsDropdown.className = 'submenu-dropdown';
    columnsDropdown.innerHTML = `
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span>1 column left</span>
        </div>
      </div>
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span>1 column right</span>
        </div>
      </div>
    `;
    columnsSubmenu.appendChild(columnsDropdown);

    // Functions submenu
    const functionsSubmenu = document.createElement('div');
    functionsSubmenu.className = 'submenu-item';
    functionsSubmenu.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">functions</span>
        <span>Function</span>
      </div>
      <span class="material-icons submenu-arrow">arrow_right</span>
    `;

    const functionGroups = {
      'Date': ['DATE', 'DAY', 'MONTH', 'YEAR', 'NOW', 'TODAY'],
      'Logical': ['IF', 'AND', 'OR', 'NOT', 'TRUE', 'FALSE'],
      'Math': ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN', 'ROUND'],
      'Statistical': ['MEDIAN', 'MODE', 'STDEV', 'VAR'],
      'Text': ['CONCAT', 'LEFT', 'RIGHT', 'MID', 'LEN', 'LOWER', 'UPPER'],
      'Engineering': ['DEC2BIN', 'BIN2DEC', 'DEC2HEX', 'HEX2DEC'],
    };

    const functionsDropdown = document.createElement('div');
    functionsDropdown.className = 'submenu-dropdown function-dropdown';
    
    Object.entries(functionGroups).forEach(([group, functions]) => {
      const groupItem = document.createElement('div');
      groupItem.className = 'submenu-item';
      groupItem.innerHTML = `
        <div class="menu-item-left">
          <span>${group}</span>
        </div>
        <span class="material-icons submenu-arrow">arrow_right</span>
      `;

      const functionsList = document.createElement('div');
      functionsList.className = 'submenu-dropdown';
      functions.forEach(func => {
        functionsList.innerHTML += `
          <div class="menu-dropdown-item">
            <div class="menu-item-left">
              <span>${func}</span>
            </div>
          </div>
        `;
      });
      
      groupItem.appendChild(functionsList);
      functionsDropdown.appendChild(groupItem);
    });
    
    functionsSubmenu.appendChild(functionsDropdown);

    // Create regular menu items
    const menuItems = [
      rowsSubmenu,
      columnsSubmenu,
      null,
      {
        label: 'Sheet',
        icon: 'post_add',
        shortcut: 'Shift+F11'
      },
      {
        label: 'Image',
        icon: 'image',
        shortcut: ''
      },
      functionsSubmenu,
      null,
      {
        label: 'Checkbox',
        icon: 'check_box',
        shortcut: ''
      },
      {
        label: 'Dropdown',
        icon: 'arrow_drop_down_circle',
        shortcut: ''
      },
      {
        label: 'Emoji',
        icon: 'emoji_emotions',
        shortcut: ''
      },
      null,
      {
        label: 'Comment',
        icon: 'comment',
        shortcut: 'Ctrl+Alt+M'
      },
      {
        label: 'Note',
        icon: 'note',
        shortcut: 'Shift+F2'
      }
    ];

    menuItems.forEach(item => {
      if (!item) {
        const separator = document.createElement('div');
        separator.className = 'menu-dropdown-separator';
        dropdown.appendChild(separator);
        return;
      }

      if (item instanceof Element) {
        dropdown.appendChild(item);
      } else {
        const menuItem = document.createElement('div');
        menuItem.className = 'menu-dropdown-item';
        menuItem.innerHTML = `
          <div class="menu-item-left">
            <span class="material-icons">${item.icon}</span>
            <span>${item.label}</span>
          </div>
          ${item.shortcut ? `<span class="shortcut">${item.shortcut}</span>` : ''}
        `;
        dropdown.appendChild(menuItem);
      }
    });

    insertMenu.appendChild(dropdown);

    // Add click handler to toggle menu
    insertMenu.addEventListener('click', (e) => {
      e.stopPropagation();
      document.querySelectorAll('.menu-dropdown').forEach(menu => {
        if (menu !== dropdown) {
          menu.classList.remove('active');
        }
      });
      dropdown.classList.toggle('active');
    });
  }

  private setupFormatMenu(): void {
    const formatMenu = document.querySelector('.menu-item:nth-child(5)');
    if (!formatMenu) return;

    const dropdown = document.createElement('div');
    dropdown.className = 'menu-dropdown';

    // Text formatting submenu
    const textSubmenu = document.createElement('div');
    textSubmenu.className = 'submenu-item';
    textSubmenu.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">text_format</span>
        <span>Text</span>
      </div>
      <span class="material-icons submenu-arrow">arrow_right</span>
    `;

    const textDropdown = document.createElement('div');
    textDropdown.className = 'submenu-dropdown';
    textDropdown.innerHTML = `
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span class="material-icons">format_bold</span>
          <span>Bold</span>
        </div>
        <span class="shortcut">Ctrl+B</span>
      </div>
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span class="material-icons">format_italic</span>
          <span>Italic</span>
        </div>
        <span class="shortcut">Ctrl+I</span>
      </div>
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span class="material-icons">format_underlined</span>
          <span>Underline</span>
        </div>
        <span class="shortcut">Ctrl+U</span>
      </div>
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span class="material-icons">strikethrough_s</span>
          <span>Strikethrough</span>
        </div>
        <span class="shortcut">Alt+Shift+5</span>
      </div>
    `;
    textSubmenu.appendChild(textDropdown);

    // Alignment submenu
    const alignmentSubmenu = document.createElement('div');
    alignmentSubmenu.className = 'submenu-item';
    alignmentSubmenu.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">format_align_left</span>
        <span>Alignment</span>
      </div>
      <span class="material-icons submenu-arrow">arrow_right</span>
    `;

    const alignmentDropdown = document.createElement('div');
    alignmentDropdown.className = 'submenu-dropdown';
    alignmentDropdown.innerHTML = `
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span class="material-icons">format_align_left</span>
          <span>Left</span>
        </div>
        <span class="shortcut">Ctrl+Shift+L</span>
      </div>
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span class="material-icons">format_align_center</span>
          <span>Center</span>
        </div>
        <span class="shortcut">Ctrl+Shift+C</span>
      </div>
      <div class="menu-dropdown-item">
        <div class="menu-item-left">
          <span class="material-icons">format_align_right</span>
          <span>Right</span>
        </div>
        <span class="shortcut">Ctrl+Shift+R</span>
      </div>
    `;
    alignmentSubmenu.appendChild(alignmentDropdown);

    // Font size submenu
    const fontSizeSubmenu = document.createElement('div');
    fontSizeSubmenu.className = 'submenu-item';
    fontSizeSubmenu.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">format_size</span>
        <span>Font size</span>
      </div>
      <span class="material-icons submenu-arrow">arrow_right</span>
    `;

    const fontSizeDropdown = document.createElement('div');
    fontSizeDropdown.className = 'submenu-dropdown font-size-dropdown';
    
    // Generate font sizes from 6 to 36 with step of 2
    for (let size = 6; size <= 36; size += 2) {
      fontSizeDropdown.innerHTML += `
        <div class="menu-dropdown-item">
          <div class="menu-item-left">
            <span>${size}</span>
          </div>
          ${size === 11 ? '<span class="material-icons">check</span>' : ''}
        </div>
      `;
    }
    fontSizeSubmenu.appendChild(fontSizeDropdown);

    // Create regular menu items
    const menuItems = [
      textSubmenu,
      alignmentSubmenu,
      fontSizeSubmenu,
      null,
      {
        label: 'Clear formatting',
        icon: 'format_clear',
        shortcut: 'Ctrl+\\'
      }
    ];

    menuItems.forEach(item => {
      if (!item) {
        const separator = document.createElement('div');
        separator.className = 'menu-dropdown-separator';
        dropdown.appendChild(separator);
        return;
      }

      if (item instanceof Element) {
        dropdown.appendChild(item);
      } else {
        const menuItem = document.createElement('div');
        menuItem.className = 'menu-dropdown-item';
        menuItem.innerHTML = `
          <div class="menu-item-left">
            <span class="material-icons">${item.icon}</span>
            <span>${item.label}</span>
          </div>
          ${item.shortcut ? `<span class="shortcut">${item.shortcut}</span>` : ''}
        `;
        dropdown.appendChild(menuItem);
      }
    });

    formatMenu.appendChild(dropdown);

    // Add click handler to toggle menu
    formatMenu.addEventListener('click', (e) => {
      e.stopPropagation();
      document.querySelectorAll('.menu-dropdown').forEach(menu => {
        if (menu !== dropdown) {
          menu.classList.remove('active');
        }
      });
      dropdown.classList.toggle('active');
    });
  }

  private setupHelpMenu(): void {
    const helpMenu = document.querySelector('.menu-item:last-child');
    if (!helpMenu) return;

    const dropdown = document.createElement('div');
    dropdown.className = 'menu-dropdown';

    const menuItems = [
      {
        label: 'Functions list',
        icon: 'functions',
        shortcut: ''
      },
      {
        label: 'Keyboard shortcuts',
        icon: 'keyboard',
        shortcut: 'Ctrl+/'
      },
      null, // separator
      {
        label: 'Privacy policy',
        icon: 'privacy_tip',
        shortcut: ''
      },
      {
        label: 'Terms of service',
        icon: 'description',
        shortcut: ''
      }
    ];

    menuItems.forEach(item => {
      if (!item) {
        const separator = document.createElement('div');
        separator.className = 'menu-dropdown-separator';
        dropdown.appendChild(separator);
        return;
      }

      const menuItem = document.createElement('div');
      menuItem.className = 'menu-dropdown-item';
      menuItem.innerHTML = `
        <div class="menu-item-left">
          <span class="material-icons">${item.icon}</span>
          <span>${item.label}</span>
        </div>
        ${item.shortcut ? `<span class="shortcut">${item.shortcut}</span>` : ''}
      `;
      dropdown.appendChild(menuItem);
    });

    helpMenu.appendChild(dropdown);

    // Add click handler to toggle menu
    helpMenu.addEventListener('click', (e) => {
      e.stopPropagation();
      document.querySelectorAll('.menu-dropdown').forEach(menu => {
        if (menu !== dropdown) {
          menu.classList.remove('active');
        }
      });
      dropdown.classList.toggle('active');
    });
  }

  private setupTitleInput(): void {
    const titleInput = document.querySelector('.title-input') as HTMLInputElement;
    if (!titleInput) return;

    // Create save indicator element
    const saveIndicator = document.createElement('div');
    saveIndicator.className = 'save-indicator';
    titleInput.parentElement?.appendChild(saveIndicator);

    // Create a debounced save function
    const debouncedSave = this.debounce(async (title: string) => {
      try {
        saveIndicator.textContent = 'Saving...';
        saveIndicator.classList.add('visible', 'loading');

        const { data: { session } } = await supabase.auth.getSession();
        if (!session) return;

        const { error } = await supabase.functions.invoke('save-spreadsheet', {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${session.access_token}`,
          },
          body: {
            spreadsheetId: this.spreadsheet?.id,
            title,
            sheetId: this.spreadsheet?.sheets[this.activeSheetId]?.id,
            data: this.data
          }
        });

        if (error) throw error;

        saveIndicator.textContent = 'Saved';
        saveIndicator.classList.remove('loading');
        setTimeout(() => {
          saveIndicator.classList.remove('visible');
        }, 2000);

      } catch (error) {
        console.error('Error saving title:', error);
        saveIndicator.textContent = 'Failed to save';
        saveIndicator.classList.remove('loading');
        saveIndicator.classList.add('error');
        setTimeout(() => {
          saveIndicator.classList.remove('visible', 'error');
        }, 3000);
      }
    }, 500);

    // Add input event listener
    titleInput.addEventListener('input', (e) => {
      const title = (e.target as HTMLInputElement).value;
      debouncedSave(title);
    });
  }

  // Debounce utility function
  private debounce<T extends (...args: any[]) => any>(
    func: T,
    wait: number
  ): (...args: Parameters<T>) => void {
    let timeout: NodeJS.Timeout;
    return (...args: Parameters<T>) => {
      clearTimeout(timeout);
      timeout = setTimeout(() => func(...args), wait);
    };
  }

  private setupClipboardHandlers(): void {
    document.addEventListener('keydown', (e) => {
      if (!this.activeCell) return;

      if (e.ctrlKey) {
        switch (e.key.toLowerCase()) {
          case 'c':
            e.preventDefault();
            this.handleCopy();
            break;
          case 'x':
            e.preventDefault();
            this.handleCut();
            break;
          case 'v':
            e.preventDefault();
            this.handlePaste();
            break;
        }
      }
    });
  }

  private handleCopy(): void {
    const selectedCells = document.querySelectorAll('.cell.selected');
    if (!selectedCells.length) return;

    const clipboardData: { [key: string]: Cell } = {};
    selectedCells.forEach((cell) => {
      const cellId = cell.getAttribute('data-cell-id');
      if (cellId && this.data[cellId]) {
        clipboardData[cellId] = { ...this.data[cellId] };
      }
    });

    // Store in localStorage since clipboard API doesn't support complex data
    localStorage.setItem('spreadsheet-clipboard', JSON.stringify(clipboardData));
  }

  private handleCut(): void {
    this.handleCopy();
    const selectedCells = document.querySelectorAll('.cell.selected');
    selectedCells.forEach((cell) => {
      const cellId = cell.getAttribute('data-cell-id');
      if (cellId) {
        delete this.data[cellId];
        const contentDiv = cell.querySelector('.cell-content') as HTMLElement;
        if (contentDiv) {
          contentDiv.textContent = '';
        }
      }
    });
  }

  private handlePaste(): void {
    if (!this.activeCell) return;

    const clipboardData = localStorage.getItem('spreadsheet-clipboard');
    if (!clipboardData) return;

    try {
      const copiedCells = JSON.parse(clipboardData) as { [key: string]: Cell };
      const copiedCellIds = Object.keys(copiedCells);
      if (!copiedCellIds.length) return;

      // Get the reference point (top-left cell of copied range)
      const referenceCell = copiedCellIds[0];
      const [refCol, refRow] = this.getCellCoordinates(referenceCell);

      // Get target cell coordinates
      const targetCellId = this.activeCell.getAttribute('data-cell-id');
      if (!targetCellId) return;
      const [targetCol, targetRow] = this.getCellCoordinates(targetCellId);

      // Calculate offset
      const colOffset = targetCol - refCol;
      const rowOffset = targetRow - refRow;

      // Paste cells with offset
      Object.entries(copiedCells).forEach(([cellId, cellData]) => {
        const [col, row] = this.getCellCoordinates(cellId);
        const newCellId = this.getCellId(col + colOffset, row + rowOffset);
        
        // Skip if new position is out of bounds
        if (col + colOffset >= this.cols || row + rowOffset >= this.rows) return;
        
        this.data[newCellId] = { ...cellData };
        const cell = document.querySelector(`[data-cell-id="${newCellId}"]`);
        if (cell) {
          const contentDiv = cell.querySelector('.cell-content') as HTMLElement;
          if (contentDiv) {
            contentDiv.textContent = cellData.value;
          }
        }
      });

    } catch (error) {
      console.error('Error pasting cells:', error);
    }
  }

  private getCellCoordinates(cellId: string): [number, number] {
    const col = cellId.charCodeAt(0) - 65; // Convert A-Z to 0-25
    const row = parseInt(cellId.slice(1)) - 1;
    return [col, row];
  }

  private getCellId(col: number, row: number): string {
    return `${String.fromCharCode(65 + col)}${row + 1}`;
  }

  private setupToolbar(): void {
    const toolbar = document.querySelector('.toolbar');
    if (!toolbar) return;

    // Text formatting buttons
    const formatButtons = {
      bold: toolbar.querySelector('[title="Bold"]'),
      italic: toolbar.querySelector('[title="Italic"]'),
      underline: toolbar.querySelector('[title="Underline"]'),
      strikethrough: toolbar.querySelector('[title="Strikethrough"]')
    };

    // Add click handlers for text formatting
    Object.entries(formatButtons).forEach(([format, button]) => {
      if (button instanceof HTMLElement) {
        button.addEventListener('click', () => {
          this.toggleTextFormat(format as keyof typeof formatButtons);
        });
      }
    });
  }

  private toggleTextFormat(format: 'bold' | 'italic' | 'underline' | 'strikethrough'): void {
    const selectedCells: NodeListOf<Element> = document.querySelectorAll('.cell.selected');
    const cells: Element[] = this.activeCell ? 
      [this.activeCell, ...Array.from(selectedCells)] : 
      Array.from(selectedCells);

    // Get the button to toggle its state
    const button = document.querySelector(`[title="${format.charAt(0).toUpperCase() + format.slice(1)}"]`);
    const isActive = button?.classList.contains('active');

    // Map format to style property
    const styleMap: Record<typeof format, string> = {
      bold: 'fontWeight',
      italic: 'fontStyle',
      underline: 'textDecoration',
      strikethrough: 'textDecoration'
    };

    // Map format to style value
    const valueMap: Record<typeof format, string> = {
      bold: 'bold',
      italic: 'italic',
      underline: 'underline',
      strikethrough: 'line-through'
    };

    cells.forEach((cell) => {
      const cellId = cell.getAttribute('data-cell-id');
      if (!cellId) return;

      const contentDiv = cell.querySelector('.cell-content') as HTMLElement;
      if (!contentDiv) return;

      // Initialize cell data if needed
      if (!this.data[cellId]) {
        this.data[cellId] = { value: '', formula: '', computed: '' };
      }
      if (!this.data[cellId].styles) {
        this.data[cellId].styles = {};
      }

      // Toggle the style
      if (format === 'underline' || format === 'strikethrough') {
        const currentDecoration = contentDiv.style.textDecoration || '';
        const decorations = new Set(currentDecoration.split(' ').filter(Boolean));
        
        if (!isActive) {
          decorations.add(valueMap[format]);
        } else {
          decorations.delete(valueMap[format]);
        }
        
        const newDecoration = Array.from(decorations).join(' ');
        contentDiv.style.textDecoration = newDecoration;
        this.data[cellId].styles.textDecoration = newDecoration;
      } else {
        const styleProperty = styleMap[format];
        const newValue = !isActive ? valueMap[format] : '';
        contentDiv.style[styleProperty] = newValue;
        this.data[cellId].styles[styleProperty] = newValue;
      }
    });

    // Toggle button state
    if (button) {
      button.classList.toggle('active');
    }
  }

  private syncFormatButtons(cell: HTMLElement): void {
    const contentDiv = cell.querySelector('.cell-content') as HTMLElement;
    if (!contentDiv) return;

    const formats = {
      bold: contentDiv.style.fontWeight === 'bold',
      italic: contentDiv.style.fontStyle === 'italic',
      underline: (contentDiv.style.textDecoration || '').includes('underline'),
      strikethrough: (contentDiv.style.textDecoration || '').includes('line-through')
    };

    Object.entries(formats).forEach(([format, isActive]) => {
      const button = document.querySelector(`[title="${format.charAt(0).toUpperCase() + format.slice(1)}"]`);
      if (button) {
        if (isActive) {
          button.classList.add('active');
        } else {
          button.classList.remove('active');
        }
      }
    });
  }

  private setupThemeToggle(): void {
    const viewMenu = document.querySelector('.menu-item:nth-child(3)');
    if (!viewMenu) return;

    // Find the View dropdown
    const dropdown = viewMenu.querySelector('.menu-dropdown');
    if (!dropdown) return;

    // Add theme toggle before the separator
    const themeToggle = document.createElement('div');
    themeToggle.className = 'menu-dropdown-item';
    themeToggle.innerHTML = `
      <div class="menu-item-left">
        <span class="material-icons">dark_mode</span>
        <span>Dark theme</span>
      </div>
      <span class="material-icons theme-check">check</span>
    `;

    // Insert before the last separator
    const separator = dropdown.querySelector('.menu-dropdown-separator:last-of-type');
    if (separator) {
      dropdown.insertBefore(themeToggle, separator);
    }

    // Initialize theme from localStorage
    const isDarkTheme = localStorage.getItem('darkTheme') === 'true';
    document.documentElement.classList.toggle('dark-theme', isDarkTheme);
    themeToggle.querySelector('.theme-check')?.classList.toggle('active', isDarkTheme);

    themeToggle.addEventListener('click', () => {
      document.documentElement.classList.toggle('dark-theme');
      const isNowDark = document.documentElement.classList.contains('dark-theme');
      localStorage.setItem('darkTheme', isNowDark.toString());
      themeToggle.querySelector('.theme-check')?.classList.toggle('active', isNowDark);
    });
  }

  private setupResizeHandlers(): void {
    // Column resize
    document.querySelectorAll('.col-resize-handle').forEach(handle => {
      handle.addEventListener('mousedown', (e: MouseEvent) => {
        e.preventDefault();
        this.isResizing = true;
        this.resizeStartX = e.clientX;
        this.resizeElement = (handle as HTMLElement).closest('.column') as HTMLElement;
        this.initialSize = this.resizeElement.offsetWidth;

        const handleMouseMove = (e: MouseEvent) => {
          if (!this.isResizing) return;
          const delta = e.clientX - this.resizeStartX;
          const newWidth = Math.max(50, this.initialSize + delta); // Minimum width of 50px
          if (this.resizeElement) {
            this.resizeElement.style.width = `${newWidth}px`;
            this.resizeElement.style.minWidth = `${newWidth}px`;
          }
        };

        const handleMouseUp = () => {
          this.isResizing = false;
          document.removeEventListener('mousemove', handleMouseMove);
          document.removeEventListener('mouseup', handleMouseUp);
        };

        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);
      });
    });

    // Row resize
    document.querySelectorAll('.row-resize-handle').forEach(handle => {
      handle.addEventListener('mousedown', (e: MouseEvent) => {
        e.preventDefault();
        this.isResizing = true;
        this.resizeStartY = e.clientY;
        this.resizeElement = (handle as HTMLElement).closest('.row-container') as HTMLElement;
        this.initialSize = this.resizeElement.offsetHeight;

        // Get the row number to find corresponding cells
        const rowNumber = this.resizeElement.querySelector('.row-number')?.textContent;
        const cells = document.querySelectorAll(`.cell[data-cell-id$="${rowNumber}"]`);

        const handleMouseMove = (e: MouseEvent) => {
          if (!this.isResizing) return;
          const delta = e.clientY - this.resizeStartY;
          const newHeight = Math.max(25, this.initialSize + delta); // Minimum height of 25px
          
          if (this.resizeElement) {
            // Update row container height
            this.resizeElement.style.height = `${newHeight}px`;
            
            // Update all cells in this row
            cells.forEach(cell => {
              (cell as HTMLElement).style.height = `${newHeight}px`;
              (cell as HTMLElement).style.minHeight = `${newHeight}px`;
              
              // Update the cell content div height
              const contentDiv = cell.querySelector('.cell-content') as HTMLElement;
              if (contentDiv) {
                contentDiv.style.height = '100%';
                contentDiv.style.lineHeight = `${newHeight}px`;
              }
            });
          }
        };

        const handleMouseUp = () => {
          this.isResizing = false;
          document.removeEventListener('mousemove', handleMouseMove);
          document.removeEventListener('mouseup', handleMouseUp);
        };

        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);
      });
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