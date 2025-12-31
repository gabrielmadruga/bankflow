import "./style.css";

// ============================================================================
// Constants
// ============================================================================

const COLUMN_INDICES = {
  DATE: 1,
  DESCRIPTION: 2,
  DEBIT: 4,
  CREDIT: 5,
  REFERENCE: 7,
} as const;

const MIN_ROW_LENGTH = 5;

// ============================================================================
// Types
// ============================================================================

interface Transaction {
  date: string;
  description: string;
  reference: string;
  debit: number;
  credit: number;
}

interface AppState {
  income: number;
  outcome: number;
  loadedFiles: string[];
  reset(): void;
  readonly netBalance: number;
}

type ExcelRow = (string | number | null)[] | null;

declare const XLSX: {
  read(data: Uint8Array, opts: { type: string }): { Sheets: Record<string, unknown>; SheetNames: string[] };
  utils: {
    sheet_to_json(sheet: unknown, opts: object): ExcelRow[];
  };
  SSF: {
    parse_date_code(value: number): { y: number; m: number; d: number };
  };
};

// ============================================================================
// DOM Elements
// ============================================================================

function getElement<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Element with id "${id}" not found`);
  return el as T;
}

const elements = {
  dropZone: getElement<HTMLDivElement>("drop-zone"),
  fileInput: getElement<HTMLInputElement>("file-input"),
  fileList: getElement<HTMLElement>("file-list"),
  txBody: getElement<HTMLTableSectionElement>("tx-body"),
  totalIncome: getElement<HTMLElement>("total-income"),
  totalOutcome: getElement<HTMLElement>("total-outcome"),
  netBalance: getElement<HTMLElement>("net-balance"),
};

// ============================================================================
// State
// ============================================================================

const state: AppState = {
  income: 0,
  outcome: 0,
  loadedFiles: [],

  reset() {
    this.income = 0;
    this.outcome = 0;
    this.loadedFiles = [];
  },

  get netBalance() {
    return this.income - this.outcome;
  },
};

// ============================================================================
// Event Listeners
// ============================================================================

function initEventListeners(): void {
  const { dropZone, fileInput } = elements;

  dropZone.addEventListener("click", () => fileInput.click());

  dropZone.addEventListener("dragover", (e: DragEvent) => {
    e.preventDefault();
    dropZone.classList.add("hover");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("hover");
  });

  dropZone.addEventListener("drop", (e: DragEvent) => {
    e.preventDefault();
    dropZone.classList.remove("hover");
    if (e.dataTransfer?.files) {
      handleFiles(e.dataTransfer.files);
    }
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files) {
      handleFiles(fileInput.files);
    }
  });
}

// ============================================================================
// File Handling
// ============================================================================

function handleFiles(files: FileList): void {
  if (!files.length) return;

  Array.from(files).forEach((file) => {
    state.loadedFiles.push(file.name);
    processFile(file);
  });

  updateFileListDisplay();
}

function processFile(file: File): void {
  const reader = new FileReader();

  reader.onerror = () => {
    console.error(`Failed to read file: ${file.name}`);
  };

  reader.onload = (e: ProgressEvent<FileReader>) => {
    try {
      const result = e.target?.result;
      if (!(result instanceof ArrayBuffer)) return;

      const data = new Uint8Array(result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: true,
        defval: null,
        blankrows: false,
      });

      appendTransactions(rows);
    } catch (error) {
      console.error(`Failed to parse Excel file: ${file.name}`, error);
    }
  };

  reader.readAsArrayBuffer(file);
}

// ============================================================================
// Data Parsing
// ============================================================================

function findHeaderRowIndex(rows: ExcelRow[]): number {
  return rows.findIndex((row) => row?.some((cell) => typeof cell === "string" && cell.toUpperCase().includes("FECHA")));
}

function parseExcelDate(value: string | number | null): string {
  if (typeof value === "number") {
    const { y, m, d } = XLSX.SSF.parse_date_code(value);
    return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }
  return value ?? "";
}

function parseMoney(value: string | number | null): number {
  if (value == null) return 0;
  if (typeof value === "number") return value;
  return Number(value.toString().replace(/\./g, "").replace(",", ".")) || 0;
}

function formatCurrency(amount: number): string {
  return `$${amount.toFixed(2)}`;
}

function isValidDataRow(row: ExcelRow): row is (string | number | null)[] {
  if (!row || row.length < MIN_ROW_LENGTH) return false;

  const desc = String(row[COLUMN_INDICES.DESCRIPTION] ?? "").trim();
  return desc.length > 0 && !desc.toUpperCase().includes("SALDO");
}

function parseTransaction(row: (string | number | null)[]): Transaction {
  return {
    date: parseExcelDate(row[COLUMN_INDICES.DATE]),
    description: String(row[COLUMN_INDICES.DESCRIPTION] ?? "").trim(),
    reference: String(row[COLUMN_INDICES.REFERENCE] ?? "").trim(),
    debit: parseMoney(row[COLUMN_INDICES.DEBIT]),
    credit: parseMoney(row[COLUMN_INDICES.CREDIT]),
  };
}

// ============================================================================
// DOM Updates
// ============================================================================

function appendTransactions(rows: ExcelRow[]): void {
  const headerIndex = findHeaderRowIndex(rows);
  const dataRows = rows.slice(headerIndex + 1);

  const fragment = document.createDocumentFragment();

  dataRows.filter(isValidDataRow).forEach((row) => {
    const tx = parseTransaction(row);

    state.income += tx.credit;
    state.outcome += tx.debit;

    fragment.appendChild(createTransactionRow(tx));
  });

  elements.txBody.appendChild(fragment);
  updateDashboard();
}

function createTransactionRow(tx: Transaction): HTMLTableRowElement {
  const { date, description, reference, debit, credit } = tx;
  const tr = document.createElement("tr");

  const cells: { text: string; className?: string }[] = [
    { text: date },
    { text: description },
    { text: reference },
    { text: debit ? `-${debit.toFixed(2)}` : "", className: "outcome" },
    { text: credit ? `+${credit.toFixed(2)}` : "", className: "income" },
  ];

  cells.forEach(({ text, className }) => {
    const td = document.createElement("td");
    td.textContent = text;
    if (className) td.className = className;
    tr.appendChild(td);
  });

  return tr;
}

function updateFileListDisplay(): void {
  elements.fileList.textContent = state.loadedFiles.length ? `Loaded files: ${state.loadedFiles.join(", ")}` : "No files loaded";
}

function updateDashboard(): void {
  elements.totalIncome.textContent = formatCurrency(state.income);
  elements.totalOutcome.textContent = formatCurrency(state.outcome);
  elements.netBalance.textContent = formatCurrency(state.netBalance);
}

// ============================================================================
// Public API
// ============================================================================

function clearData(): void {
  state.reset();
  elements.txBody.innerHTML = "";
  updateFileListDisplay();
  updateDashboard();
}

// Make clearData available globally for HTML onclick
declare global {
  interface Window {
    clearData: typeof clearData;
  }
}
window.clearData = clearData;

// ============================================================================
// Initialize
// ============================================================================

initEventListeners();
updateDashboard();
