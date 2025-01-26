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
    
    // Add row numbers
    for (let row = 0; row < this.rows; row++) {
      const rowNumber = document.createElement('div');
      rowNumber.className = 'row-number';
      rowNumber.textContent = (row + 1).toString();
      rowNumbersColumn.appendChild(rowNumber);
    }
    gridContainer.appendChild(rowNumbersColumn);

    // Create columns with headers
    for (let col = 0; col < this.cols; col++) {
      const column = document.createElement('div');
      column.className = 'column';
      
      // Add column header
      const header = document.createElement('div');
      header.className = 'column-header';
      header.textContent = String.fromCharCode(65 + col);
      column.appendChild(header);
      
      // Create cells for this column
      for (let row = 0; row < this.rows; row++) {
        const cell = document.createElement('div');
        cell.className = 'cell';
        const cellId = getCellId(row, col);
        cell.dataset.cellId = cellId;
        cell.contentEditable = 'true';
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
    this.container.addEventListener('focus', (e) => {
      const cell = (e.target as HTMLElement).closest('td');
      if (cell) {
        this.activeCell = cell;
        const cellId = cell.dataset.cellId!;
        cell.textContent = this.data[cellId].formula || this.data[cellId].value;
      }
    }, true);

    this.container.addEventListener('blur', (e) => {
      const cell = (e.target as HTMLElement).closest('td');
      if (cell) {
        const cellId = cell.dataset.cellId!;
        const newValue = cell.textContent || '';
        this.updateCell(cellId, newValue);
        this.recalculateAll();
      }
    }, true);
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
}

// Initialize the spreadsheet
new Spreadsheet('app'); 