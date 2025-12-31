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
  key: string;
  date: string;
  description: string;
  reference: string;
  debit: number;
  credit: number;
}

interface AppState {
  income: number;
  outcome: number;
  loadedFiles: Set<string>;
  transactions: Map<string, Transaction>;
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
  loadedFiles: new Set(),
  transactions: new Map(),

  reset() {
    this.income = 0;
    this.outcome = 0;
    this.loadedFiles = new Set();
    this.transactions = new Map();
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
    state.loadedFiles.add(file.name);
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

      addTransactions(rows);
    } catch (error) {
      console.error(`Failed to parse Excel file: ${file.name}`, error);
    }
  };

  reader.readAsArrayBuffer(file);
}

// ============================================================================
// Date Utilities
// ============================================================================

function parseExcelDate(value: string | number | null): string {
  if (typeof value === "number") {
    const { y, m, d } = XLSX.SSF.parse_date_code(value);
    return `${String(d).padStart(2, "0")}/${String(m).padStart(2, "0")}/${y}`;
  }
  return value ?? "";
}

function parseDateString(dateStr: string): Date | null {
  const parts = dateStr.split("/");
  if (parts.length !== 3) return null;

  const day = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1; // Month is 0-indexed
  const year = parseInt(parts[2], 10);

  if (isNaN(day) || isNaN(month) || isNaN(year)) return null;

  return new Date(year, month, day);
}

function compareTransactionsByDate(a: Transaction, b: Transaction): number {
  const dateA = parseDateString(a.date);
  const dateB = parseDateString(b.date);

  // Handle invalid dates
  if (!dateA && !dateB) return 0;
  if (!dateA) return 1; // Push invalid dates to end
  if (!dateB) return -1;

  return dateB.getTime() - dateA.getTime();
}

// ============================================================================
// Data Parsing
// ============================================================================

function findHeaderRowIndex(rows: ExcelRow[]): number {
  return rows.findIndex((row) => row?.some((cell) => typeof cell === "string" && cell.toUpperCase().includes("FECHA")));
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
  const date = parseExcelDate(row[COLUMN_INDICES.DATE]);
  const description = String(row[COLUMN_INDICES.DESCRIPTION] ?? "").trim();
  const reference = String(row[COLUMN_INDICES.REFERENCE] ?? "").trim();
  const debit = parseMoney(row[COLUMN_INDICES.DEBIT]);
  const credit = parseMoney(row[COLUMN_INDICES.CREDIT]);

  return {
    key: `${date}|${description}|${reference}|${debit}|${credit}`,
    date,
    description,
    reference,
    debit,
    credit,
  };
}

// ============================================================================
// Transaction Management
// ============================================================================

function addTransactions(rows: ExcelRow[]): void {
  const headerIndex = findHeaderRowIndex(rows);
  const dataRows = rows.slice(headerIndex + 1);

  dataRows
    .filter(isValidDataRow)
    .map(parseTransaction)
    .forEach((tx) => {
      state.transactions.set(tx.key, tx);
    });

  recalculateTotals();
  renderTransactions();
}

function getSortedTransactions(): Transaction[] {
  return Array.from(state.transactions.values()).sort(compareTransactionsByDate);
}

function recalculateTotals(): void {
  state.income = 0;
  state.outcome = 0;

  state.transactions.forEach((tx) => {
    state.income += tx.credit;
    state.outcome += tx.debit;
  });
}

// ============================================================================
// DOM Updates
// ============================================================================

function renderTransactions(): void {
  const fragment = document.createDocumentFragment();

  getSortedTransactions().forEach((tx) => {
    fragment.appendChild(createTransactionRow(tx));
  });

  elements.txBody.innerHTML = "";
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
  const files = Array.from(state.loadedFiles);
  elements.fileList.textContent = files.length ? `Loaded files: ${files.join(", ")}` : "No files loaded";
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
