/**
 * FinanceFlow - iOS Optimized Personal Finance Manager
 * LocalStorage-based offline Progressive Web App
 */

// ================= CONSTANTS & DEFAULT STATE =================
const STORAGE_KEYS = {
  ACCOUNTS: 'ff_accounts_v2',
  CATEGORIES: 'ff_categories_v2',
  TRANSACTIONS: 'ff_transactions_v2',
  SETTINGS: 'ff_settings_v2'
};

const DEFAULT_ACCOUNTS = [
  { id: 'acc_sbi', name: 'Bank Acc 1 (SBI)', type: 'bank', balance: 0.00 },
  { id: 'acc_bank2', name: 'Bank Acc 2', type: 'bank', balance: 0.00 },
  { id: 'acc_cash', name: 'Cash', type: 'cash', balance: 0.00 }
];

const DEFAULT_CATEGORIES = {
  expense: ['Food', 'Transport', 'Mobile Recharge', 'Dress', 'Lifestyle', 'Education', 'Selfcare', 'Household'],
  income: ['Salary', 'Pocket Money', 'Allowance', 'Bonus'],
  transfer: ['Food', 'Transport', 'Mobile Recharge', 'Dress', 'Lifestyle', 'Education', 'Selfcare', 'Household', 'Savings']
};

const DEFAULT_SETTINGS = {
  currency: '₹',
  theme: 'theme-dark'
};

// ================= APP STATE =================
let state = {
  accounts: [],
  categories: { expense: [], income: [], transfer: [] },
  transactions: [],
  settings: { ...DEFAULT_SETTINGS },
  currentView: 'dashboard',
  selectedTxMonth: new Date().toISOString().slice(0, 7), // YYYY-MM
  timeMode: 'auto', // 'auto' | 'manual'
  timeAutoInterval: null,
  activeTxSegment: 'all'
};

// Chart instances
let expenseChartInstance = null;
let incomeChartInstance = null;
let monthlyTrendChartInstance = null;

// ================= INITIALIZATION =================
document.addEventListener('DOMContentLoaded', () => {
  loadStoredData();
  applyTheme(state.settings.theme);
  updateCurrencyDisplays();
  initDateHeaders();
  populateAccountsUI();
  populateCategorySelects();
  setupEventListeners();
  renderDashboard();
  renderTransactionsView();
  renderStatsView();
  renderSetupView();
  checkShortcutsQueryParam();

  // Unregister all Service Workers to guarantee live updates on iPhone
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.getRegistrations().then(regs => {
      for (let reg of regs) {
        reg.unregister().catch(() => {});
      }
    });
  }
});

// Load data from LocalStorage
function loadStoredData() {
  try {
    const acc = localStorage.getItem(STORAGE_KEYS.ACCOUNTS);
    state.accounts = acc ? JSON.parse(acc) : JSON.parse(JSON.stringify(DEFAULT_ACCOUNTS));

    const cat = localStorage.getItem(STORAGE_KEYS.CATEGORIES);
    state.categories = cat ? JSON.parse(cat) : JSON.parse(JSON.stringify(DEFAULT_CATEGORIES));

    const tx = localStorage.getItem(STORAGE_KEYS.TRANSACTIONS);
    state.transactions = tx ? JSON.parse(tx) : [];
    // Purge any legacy sample transactions
    state.transactions = state.transactions.filter(t => !t.id.startsWith('tx_init_'));

    // If there are no transactions, ensure accounts start fresh at zero
    if (state.transactions.length === 0) {
      state.accounts.forEach(a => {
        if (a.id === 'acc_sbi' && a.balance === 5000) a.balance = 0;
        if (a.id === 'acc_bank2' && a.balance === 2500) a.balance = 0;
        if (a.id === 'acc_cash' && a.balance === 1000) a.balance = 0;
      });
      saveData('accounts');
      saveData('transactions');
    }

    const set = localStorage.getItem(STORAGE_KEYS.SETTINGS);
    state.settings = set ? { ...DEFAULT_SETTINGS, ...JSON.parse(set) } : { ...DEFAULT_SETTINGS };
  } catch (e) {
    console.error('Failed to load storage, resetting to defaults', e);
    state.accounts = JSON.parse(JSON.stringify(DEFAULT_ACCOUNTS));
    state.categories = JSON.parse(JSON.stringify(DEFAULT_CATEGORIES));
    state.transactions = [];
    state.settings = { ...DEFAULT_SETTINGS };
  }
}

function saveData(key) {
  try {
    if (!key || key === 'accounts') localStorage.setItem(STORAGE_KEYS.ACCOUNTS, JSON.stringify(state.accounts));
    if (!key || key === 'categories') localStorage.setItem(STORAGE_KEYS.CATEGORIES, JSON.stringify(state.categories));
    if (!key || key === 'transactions') localStorage.setItem(STORAGE_KEYS.TRANSACTIONS, JSON.stringify(state.transactions));
    if (!key || key === 'settings') localStorage.setItem(STORAGE_KEYS.SETTINGS, JSON.stringify(state.settings));
  } catch (err) {
    console.error('Error saving data to LocalStorage', err);
  }
}

// ================= EXCEL IMPORT HELPERS =================
function parseExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // Excel serial date code
    return new Date(Math.round((val - 25569) * 86400 * 1000));
  }
  const str = String(val).trim();
  // Try YYYY-MM-DD or YYYY/MM/DD
  let m = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) {
    return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  }
  // Try DD-MM-YYYY or DD/MM/YYYY
  m = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) {
    return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));
  }
  // Try MM-DD-YYYY or MM/DD/YYYY
  m = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})/);
  if (m) {
    const yr = parseInt(m[3]) < 100 ? 2000 + parseInt(m[3]) : parseInt(m[3]);
    return new Date(yr, parseInt(m[1]) - 1, parseInt(m[2]));
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function handleExcelImport(event) {
  const file = event.target.files[0];
  if (!file) return;

  const statusEl = document.getElementById('import-excel-status');
  statusEl.className = 'import-status-text';
  statusEl.classList.remove('hidden');
  statusEl.textContent = 'Reading and importing Excel sheet...';

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });

      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        statusEl.className = 'import-status-text error';
        statusEl.textContent = 'Excel file has no sheets.';
        return;
      }

      // Read rows from the first sheet
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });

      if (!rawRows || rawRows.length === 0) {
        statusEl.className = 'import-status-text error';
        statusEl.textContent = 'Excel sheet is empty or contains no valid rows.';
        return;
      }

      let importedCount = 0;
      let skippedCount = 0;

      rawRows.forEach((row, idx) => {
        // Case-insensitive key lookup helper
        const getCol = (...keys) => {
          for (let k of keys) {
            const rowKey = Object.keys(row).find(rk => rk.trim().toLowerCase() === k.trim().toLowerCase());
            if (rowKey && row[rowKey] !== undefined && row[rowKey] !== '') return row[rowKey];
          }
          return '';
        };

        // 1. Period -> Date (set time to 9:00 AM)
        const periodRaw = getCol('Period', 'Date', 'Transaction Date');
        if (!periodRaw) {
          skippedCount++;
          return;
        }

        let datetimeStr = '';
        const parsedDate = parseExcelDate(periodRaw);
        if (parsedDate && !isNaN(parsedDate.getTime())) {
          const y = parsedDate.getFullYear();
          const m = String(parsedDate.getMonth() + 1).padStart(2, '0');
          const d = String(parsedDate.getDate()).padStart(2, '0');
          // Specified by user: "Period has only date in it and for that time put 9 am only"
          datetimeStr = `${y}-${m}-${d}T09:00`;
        } else {
          const today = new Date().toISOString().slice(0, 10);
          datetimeStr = `${today}T09:00`;
        }

        // 2. Amount: INR -> fallback to Amount
        const amountRaw = getCol('INR', 'Amount', 'Total');
        const cleanAmtStr = String(amountRaw).replace(/[^0-9.-]/g, '');
        const amount = Math.abs(parseFloat(cleanAmtStr));
        if (isNaN(amount) || amount === 0) {
          skippedCount++;
          return;
        }

        // 3. Type: Income/Expense
        const typeRaw = String(getCol('Income/Expense', 'Type', 'Transaction Type')).toLowerCase().trim();
        let txType = 'expense';
        if (typeRaw.includes('inc') || typeRaw === 'income') {
          txType = 'income';
        } else if (typeRaw.includes('trans') || typeRaw === 'transfer') {
          txType = 'transfer';
        } else {
          txType = 'expense';
        }

        // 4. Accounts: Account name (SBI, Bank 2, Cash, etc.)
        const accNameRaw = String(getCol('Accounts', 'Account')).trim();
        let targetAccountId = state.accounts[0]?.id || 'acc_sbi';

        if (accNameRaw) {
          let matchAcc = state.accounts.find(a => 
            a.name.toLowerCase() === accNameRaw.toLowerCase() ||
            a.name.toLowerCase().includes(accNameRaw.toLowerCase()) ||
            accNameRaw.toLowerCase().includes(a.name.toLowerCase())
          );

          if (!matchAcc) {
            const newAccId = 'acc_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
            matchAcc = {
              id: newAccId,
              name: accNameRaw,
              type: accNameRaw.toLowerCase().includes('cash') ? 'cash' : 'bank',
              balance: 0.00
            };
            state.accounts.push(matchAcc);
          }
          targetAccountId = matchAcc.id;
        }

        // 5. Category: Category
        let categoryRaw = String(getCol('Category')).trim() || 'General';
        if (!state.categories[txType].includes(categoryRaw)) {
          state.categories[txType].push(categoryRaw);
        }

        // 6. Note: Note (fallback to Description)
        const noteRaw = String(getCol('Note', 'Description')).trim();

        // 7. Update Balances in real time
        if (txType === 'expense') {
          const acc = state.accounts.find(a => a.id === targetAccountId);
          if (acc) acc.balance = (parseFloat(acc.balance) || 0) - amount;
        } else if (txType === 'income') {
          const acc = state.accounts.find(a => a.id === targetAccountId);
          if (acc) acc.balance = (parseFloat(acc.balance) || 0) + amount;
        } else if (txType === 'transfer') {
          const fromAcc = state.accounts.find(a => a.id === targetAccountId);
          const toAcc = state.accounts.find(a => a.id !== targetAccountId) || state.accounts[0];
          if (fromAcc && toAcc && fromAcc.id !== toAcc.id) {
            fromAcc.balance = (parseFloat(fromAcc.balance) || 0) - amount;
            toAcc.balance = (parseFloat(toAcc.balance) || 0) + amount;
          }
        }

        // 8. Add transaction object
        state.transactions.push({
          id: 'tx_imp_' + Date.now() + '_' + idx + '_' + Math.floor(Math.random() * 1000),
          type: txType,
          amount: amount,
          datetime: datetimeStr,
          category: categoryRaw,
          accountId: targetAccountId,
          fromAccountId: txType === 'transfer' ? targetAccountId : null,
          toAccountId: txType === 'transfer' ? (state.accounts.find(a => a.id !== targetAccountId)?.id || null) : null,
          note: noteRaw,
          timestamp: new Date(datetimeStr).getTime() || Date.now()
        });

        importedCount++;
      });

      // Save updated data
      saveData('accounts');
      saveData('categories');
      saveData('transactions');

      // Update UI
      renderDashboard();
      renderTransactionsView();
      renderStatsView();
      renderSetupView();
      populateAccountsUI();

      statusEl.className = 'import-status-text success';
      statusEl.textContent = `✓ Successfully imported ${importedCount} transactions! Balances updated.`;
      showToast(`Imported ${importedCount} transactions!`);

      event.target.value = '';
    } catch (err) {
      console.error('Excel Import Error:', err);
      statusEl.className = 'import-status-text error';
      statusEl.textContent = 'Failed to import Excel: ' + err.message;
    }
  };

  reader.readAsArrayBuffer(file);
}

// ================= DATE HELPERS =================
function getDeviceLocalDateTime() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const mins = String(now.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day}T${hours}:${mins}`;
}

function formatDisplayDate(dateInput) {
  if (!dateInput) return '';
  const d = new Date(dateInput);
  return d.toLocaleDateString('en-IN', {
    day: 'numeric',
    month: 'short',
    year: 'numeric'
  });
}

function formatDisplayTime(dateInput) {
  if (!dateInput) return '';
  const d = new Date(dateInput);
  return d.toLocaleTimeString('en-IN', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: true
  });
}

function formatCurrency(val) {
  const num = parseFloat(val) || 0;
  return `${state.settings.currency}${num.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function initDateHeaders() {
  const today = new Date();
  const options = { weekday: 'short', month: 'short', day: 'numeric' };
  document.getElementById('current-date-display').textContent = today.toLocaleDateString('en-US', options);
}

// ================= DASHBOARD UI =================
function renderDashboard() {
  // 1. Calculate Total Net Worth (Sum of all accounts)
  const totalBalance = state.accounts.reduce((acc, a) => acc + (parseFloat(a.balance) || 0), 0);
  document.getElementById('dashboard-total-balance').textContent = formatCurrency(totalBalance);

  // 2. Calculate This Month's Income & Expense
  const currentMonth = new Date().toISOString().slice(0, 7);
  let monthIncome = 0;
  let monthExpense = 0;

  state.transactions.forEach(t => {
    if (t.datetime && t.datetime.startsWith(currentMonth)) {
      const amt = parseFloat(t.amount) || 0;
      if (t.type === 'income') monthIncome += amt;
      if (t.type === 'expense') monthExpense += amt;
    }
  });

  document.getElementById('month-income-stat').textContent = `+${formatCurrency(monthIncome)}`;
  document.getElementById('month-expense-stat').textContent = `-${formatCurrency(monthExpense)}`;

  // 3. Render Row 1: Accounts Horizontal Row
  renderAccountsRow();

  // 4. Render Recent Transactions List
  renderRecentTransactions();
}

function renderAccountsRow() {
  const container = document.getElementById('accounts-container');
  container.innerHTML = '';

  state.accounts.forEach(acc => {
    const card = document.createElement('div');
    card.className = 'account-card glass-panel';
    card.onclick = () => openAccountEditModal(acc.id);

    let iconClass = 'bank-sbi';
    let iconSvg = `<svg class="icon-sm" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="5" width="20" height="14" rx="2"></rect><line x1="2" y1="10" x2="22" y2="10"></line></svg>`;

    if (acc.id === 'acc_cash' || acc.type === 'cash') {
      iconClass = 'cash';
      iconSvg = `<svg class="icon-sm" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="6" width="20" height="12" rx="2"></rect><circle cx="12" cy="12" r="3"></circle></svg>`;
    } else if (acc.id === 'acc_bank2') {
      iconClass = 'bank-2';
      iconSvg = `<svg class="icon-sm" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 21h18M3 10h18M5 6l7-3 7 3M4 10v11M20 10v11M8 14v4M12 14v4M16 14v4"></path></svg>`;
    } else if (acc.id !== 'acc_sbi') {
      iconClass = 'custom';
      iconSvg = `<svg class="icon-sm" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><polyline points="12 6 12 12 14 14"></polyline></svg>`;
    }

    card.innerHTML = `
      <div class="account-top">
        <div class="account-icon ${iconClass}">
          ${iconSvg}
        </div>
        <span class="account-tag">${acc.type.toUpperCase()}</span>
      </div>
      <div class="account-name" title="${acc.name}">${acc.name}</div>
      <div class="account-balance">${formatCurrency(acc.balance)}</div>
    `;

    container.appendChild(card);
  });
}

function renderRecentTransactions() {
  const container = document.getElementById('recent-transactions-list');
  container.innerHTML = '';

  const recent = [...state.transactions]
    .sort((a, b) => new Date(b.datetime) - new Date(a.datetime))
    .slice(0, 5);

  if (recent.length === 0) {
    container.innerHTML = `
      <div class="empty-placeholder">
        <p>No transactions recorded yet.</p>
        <p style="font-size:12px;color:var(--text-muted);margin-top:4px;">Tap <strong>Record Transaction</strong> to log your first expense!</p>
      </div>
    `;
    return;
  }

  recent.forEach(tx => {
    container.appendChild(createTransactionListItem(tx));
  });
}

function createTransactionListItem(tx) {
  const item = document.createElement('div');
  item.className = 'tx-item';

  let typeClass = `cat-icon-${tx.type}`;
  let sign = '-';
  let amtClass = 'text-expense';

  if (tx.type === 'income') {
    sign = '+';
    amtClass = 'text-income';
  } else if (tx.type === 'transfer') {
    sign = '⇄ ';
    amtClass = 'text-transfer';
  }

  // Account name
  let accountLabel = '';
  if (tx.type === 'transfer') {
    const fromAcc = state.accounts.find(a => a.id === tx.fromAccountId)?.name || 'Account';
    const toAcc = state.accounts.find(a => a.id === tx.toAccountId)?.name || 'Account';
    accountLabel = `${fromAcc} → ${toAcc}`;
  } else {
    accountLabel = state.accounts.find(a => a.id === tx.accountId)?.name || 'Account';
  }

  const subText = tx.note ? `${accountLabel} • ${tx.note}` : accountLabel;

  item.innerHTML = `
    <div class="tx-item-left">
      <div class="tx-cat-icon ${typeClass}">
        ${tx.type === 'income' ? '↓' : tx.type === 'expense' ? '↑' : '⇄'}
      </div>
      <div class="tx-info-block">
        <span class="tx-cat-name">${tx.category}</span>
        <span class="tx-sub-info">${subText}</span>
      </div>
    </div>
    <div class="tx-item-right">
      <span class="tx-amount-display ${amtClass}">${sign}${formatCurrency(tx.amount)}</span>
      <span class="tx-time-tag">${formatDisplayDate(tx.datetime)} ${formatDisplayTime(tx.datetime)}</span>
    </div>
  `;

  return item;
}

// ================= TRANSACTIONS TAB - DATE-WISE TITLED LIST =================
function renderTransactionsView() {
  updateMonthPickerUI();
  populateFilterOptions();
  renderDateGroupedTransactions();
}

function updateMonthPickerUI() {
  const title = document.getElementById('tx-month-title');
  const [year, month] = state.selectedTxMonth.split('-');
  const dateObj = new Date(parseInt(year), parseInt(month) - 1, 1);
  title.textContent = dateObj.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  // Update hidden select
  const select = document.getElementById('tx-month-select');
  select.innerHTML = '';
  
  // Generate 12 months before and 6 months after
  const cur = new Date();
  for (let i = -12; i <= 6; i++) {
    const d = new Date(cur.getFullYear(), cur.getMonth() + i, 1);
    const val = d.toISOString().slice(0, 7);
    const opt = document.createElement('option');
    opt.value = val;
    opt.textContent = d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
    if (val === state.selectedTxMonth) opt.selected = true;
    select.appendChild(opt);
  }
}

function populateFilterOptions() {
  const accSelect = document.getElementById('filter-account-select');
  if (accSelect) {
    accSelect.innerHTML = '<option value="all">All Accounts</option>';
    state.accounts.forEach(a => {
      accSelect.innerHTML += `<option value="${a.id}">${a.name}</option>`;
    });
  }

  const catSelect = document.getElementById('filter-category-select');
  if (catSelect) {
    catSelect.innerHTML = '<option value="all">All Categories</option>';
    const allCategories = new Set([
      ...state.categories.expense,
      ...state.categories.income,
      ...state.categories.transfer
    ]);
    allCategories.forEach(c => {
      catSelect.innerHTML += `<option value="${c}">${c}</option>`;
    });
  }
}

function renderDateGroupedTransactions() {
  const container = document.getElementById('date-grouped-tx-container');
  if (!container) return;
  container.innerHTML = '';

  const searchQuery = (document.getElementById('tx-search-input')?.value || '').toLowerCase().trim();
  const sortMode = document.getElementById('tx-sort-select')?.value || 'date-desc';
  const filterAcc = document.getElementById('filter-account-select')?.value || 'all';
  const filterCat = document.getElementById('filter-category-select')?.value || 'all';
  const activeSegment = state.activeTxSegment || 'all';

  // Calculate monthly summary totals across ALL transactions in this month
  let monthTotalIncome = 0;
  let monthTotalExpense = 0;
  let monthTotalTransfer = 0;

  state.transactions.forEach(t => {
    if (t.datetime && t.datetime.startsWith(state.selectedTxMonth)) {
      const amt = parseFloat(t.amount) || 0;
      if (t.type === 'income') monthTotalIncome += amt;
      else if (t.type === 'expense') monthTotalExpense += amt;
      else if (t.type === 'transfer') monthTotalTransfer += amt;
    }
  });

  const incEl = document.getElementById('col-income-total');
  const expEl = document.getElementById('col-expense-total');
  const trfEl = document.getElementById('col-transfer-total');
  if (incEl) incEl.textContent = `+${formatCurrency(monthTotalIncome)}`;
  if (expEl) expEl.textContent = `-${formatCurrency(monthTotalExpense)}`;
  if (trfEl) trfEl.textContent = `⇄ ${formatCurrency(monthTotalTransfer)}`;

  // Filter transactions
  let filtered = state.transactions.filter(t => {
    if (!t.datetime || !t.datetime.startsWith(state.selectedTxMonth)) return false;
    if (activeSegment !== 'all' && t.type !== activeSegment) return false;
    if (searchQuery) {
      const matchNote = (t.note || '').toLowerCase().includes(searchQuery);
      const matchCat = (t.category || '').toLowerCase().includes(searchQuery);
      if (!matchNote && !matchCat) return false;
    }
    if (filterAcc !== 'all') {
      if (t.type === 'transfer') {
        if (t.fromAccountId !== filterAcc && t.toAccountId !== filterAcc) return false;
      } else {
        if (t.accountId !== filterAcc) return false;
      }
    }
    if (filterCat !== 'all' && t.category !== filterCat) return false;
    return true;
  });

  // Sort
  filtered.sort((a, b) => {
    if (sortMode === 'date-desc') return new Date(b.datetime) - new Date(a.datetime);
    if (sortMode === 'date-asc') return new Date(a.datetime) - new Date(b.datetime);
    if (sortMode === 'amount-desc') return b.amount - a.amount;
    if (sortMode === 'amount-asc') return a.amount - b.amount;
    return 0;
  });

  if (filtered.length === 0) {
    container.innerHTML = `
      <div class="empty-placeholder glass-panel" style="padding: 36px 20px;">
        <p style="font-size: 15px; font-weight: 600; color: var(--text-main);">No transactions found</p>
        <p style="font-size: 12px; color: var(--text-dim); margin-top: 4px;">Tap Record above or import from Excel in Setup.</p>
      </div>
    `;
    return;
  }

  // Group by date (YYYY-MM-DD)
  const groupsByDate = {};
  filtered.forEach(tx => {
    const dateKey = tx.datetime ? tx.datetime.slice(0, 10) : 'Other';
    if (!groupsByDate[dateKey]) groupsByDate[dateKey] = [];
    groupsByDate[dateKey].push(tx);
  });

  const todayStr = new Date().toISOString().slice(0, 10);
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = yesterday.toISOString().slice(0, 10);

  // Render each date group
  Object.keys(groupsByDate).forEach(dateKey => {
    const txList = groupsByDate[dateKey];
    const groupCard = document.createElement('div');
    groupCard.className = 'date-group-card glass-panel';

    let displayTitle = '';
    const d = new Date(dateKey + 'T00:00:00');
    const formattedDate = d.toLocaleDateString('en-US', { day: 'numeric', month: 'short', year: 'numeric' });
    const weekday = d.toLocaleDateString('en-US', { weekday: 'short' });

    if (dateKey === todayStr) {
      displayTitle = `Today • ${formattedDate}`;
    } else if (dateKey === yesterdayStr) {
      displayTitle = `Yesterday • ${formattedDate}`;
    } else {
      displayTitle = `${weekday}, ${formattedDate}`;
    }

    let dayTotal = 0;
    txList.forEach(t => {
      const amt = parseFloat(t.amount) || 0;
      if (t.type === 'income') dayTotal += amt;
      else if (t.type === 'expense') dayTotal -= amt;
    });

    const daySign = dayTotal > 0 ? '+' : dayTotal < 0 ? '-' : '';
    const dayClass = dayTotal > 0 ? 'text-income' : dayTotal < 0 ? 'text-expense' : '';

    groupCard.innerHTML = `
      <div class="date-group-header">
        <span class="date-group-title">${displayTitle}</span>
        <span class="date-group-day-total ${dayClass}">
          ${dayTotal !== 0 ? `${daySign}${formatCurrency(Math.abs(dayTotal))}` : `${txList.length} record${txList.length > 1 ? 's' : ''}`}
        </span>
      </div>
      <div class="date-group-items"></div>
    `;

    const itemsContainer = groupCard.querySelector('.date-group-items');
    txList.forEach(tx => {
      itemsContainer.appendChild(createDateGroupedTxItem(tx));
    });

    container.appendChild(groupCard);
  });
}

function createDateGroupedTxItem(tx) {
  const item = document.createElement('div');
  item.className = 'tx-item';

  let typeClass = `cat-icon-${tx.type}`;
  let sign = '-';
  let amtClass = 'text-expense';

  if (tx.type === 'income') {
    sign = '+';
    amtClass = 'text-income';
  } else if (tx.type === 'transfer') {
    sign = '⇄ ';
    amtClass = 'text-transfer';
  }

  let accountLabel = '';
  if (tx.type === 'transfer') {
    const fromAcc = state.accounts.find(a => a.id === tx.fromAccountId)?.name || 'Account';
    const toAcc = state.accounts.find(a => a.id === tx.toAccountId)?.name || 'Account';
    accountLabel = `${fromAcc} → ${toAcc}`;
  } else {
    accountLabel = state.accounts.find(a => a.id === tx.accountId)?.name || 'Account';
  }

  const subText = tx.note ? `${accountLabel} • ${tx.note}` : accountLabel;

  item.innerHTML = `
    <div class="tx-item-left">
      <div class="tx-cat-icon ${typeClass}">
        ${tx.type === 'income' ? '↓' : tx.type === 'expense' ? '↑' : '⇄'}
      </div>
      <div class="tx-info-block">
        <span class="tx-cat-name">${tx.category}</span>
        <span class="tx-sub-info">${subText}</span>
      </div>
    </div>
    <div class="tx-item-right">
      <span class="tx-amount-display ${amtClass}">${sign}${formatCurrency(tx.amount)}</span>
      <div style="display:flex;align-items:center;justify-content:flex-end;gap:6px;margin-top:2px;">
        <span class="tx-time-tag">${formatDisplayTime(tx.datetime)}</span>
        <button class="tx-edit-inline-btn" title="Edit transaction">Edit</button>
        <button class="tx-delete-inline-btn" title="Delete transaction">Delete</button>
      </div>
    </div>
  `;

  item.querySelector('.tx-edit-inline-btn').onclick = (e) => {
    e.stopPropagation();
    openRecordModalForEdit(tx.id);
  };

  item.querySelector('.tx-delete-inline-btn').onclick = (e) => {
    e.stopPropagation();
    deleteTransaction(tx.id);
  };

  return item;
}

// ================= TRANSACTION CRUD & BALANCE UPDATES =================
function addTransaction(txData) {
  const amount = parseFloat(txData.amount);
  if (isNaN(amount) || amount <= 0) {
    showToast('Please enter a valid amount.');
    return false;
  }

  // Update Account Balances
  if (txData.type === 'expense') {
    const acc = state.accounts.find(a => a.id === txData.accountId);
    if (!acc) return false;
    acc.balance = (parseFloat(acc.balance) || 0) - amount;
  } else if (txData.type === 'income') {
    const acc = state.accounts.find(a => a.id === txData.accountId);
    if (!acc) return false;
    acc.balance = (parseFloat(acc.balance) || 0) + amount;
  } else if (txData.type === 'transfer') {
    if (txData.fromAccountId === txData.toAccountId) {
      showToast('Cannot transfer to the same account.');
      return false;
    }
    const fromAcc = state.accounts.find(a => a.id === txData.fromAccountId);
    const toAcc = state.accounts.find(a => a.id === txData.toAccountId);
    if (!fromAcc || !toAcc) return false;

    fromAcc.balance = (parseFloat(fromAcc.balance) || 0) - amount;
    toAcc.balance = (parseFloat(toAcc.balance) || 0) + amount;
  }

  // Create new transaction object
  const newTx = {
    id: 'tx_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
    type: txData.type,
    amount: amount,
    datetime: txData.datetime || getDeviceLocalDateTime(),
    category: txData.category,
    accountId: txData.accountId || null,
    fromAccountId: txData.fromAccountId || null,
    toAccountId: txData.toAccountId || null,
    note: (txData.note || '').trim(),
    timestamp: Date.now()
  };

  state.transactions.unshift(newTx);

  saveData('accounts');
  saveData('transactions');

  renderDashboard();
  renderTransactionsView();
  renderStatsView();
  renderSetupView();

  showToast('Transaction recorded successfully!');
  triggerHapticFeedback();
  return true;
}

function updateTransaction(editId, newTxData) {
  const index = state.transactions.findIndex(t => t.id === editId);
  if (index === -1) return false;

  const oldTx = state.transactions[index];
  const oldAmount = parseFloat(oldTx.amount) || 0;
  const newAmount = parseFloat(newTxData.amount);

  if (isNaN(newAmount) || newAmount <= 0) {
    showToast('Please enter a valid amount.');
    return false;
  }

  // 1. Reversal of old transaction balance effect
  if (oldTx.type === 'expense') {
    const acc = state.accounts.find(a => a.id === oldTx.accountId);
    if (acc) acc.balance = (parseFloat(acc.balance) || 0) + oldAmount;
  } else if (oldTx.type === 'income') {
    const acc = state.accounts.find(a => a.id === oldTx.accountId);
    if (acc) acc.balance = (parseFloat(acc.balance) || 0) - oldAmount;
  } else if (oldTx.type === 'transfer') {
    const fromAcc = state.accounts.find(a => a.id === oldTx.fromAccountId);
    const toAcc = state.accounts.find(a => a.id === oldTx.toAccountId);
    if (fromAcc) fromAcc.balance = (parseFloat(fromAcc.balance) || 0) + oldAmount;
    if (toAcc) toAcc.balance = (parseFloat(toAcc.balance) || 0) - oldAmount;
  }

  // 2. Apply new transaction balance effect
  if (newTxData.type === 'expense') {
    const acc = state.accounts.find(a => a.id === newTxData.accountId);
    if (!acc) return false;
    acc.balance = (parseFloat(acc.balance) || 0) - newAmount;
  } else if (newTxData.type === 'income') {
    const acc = state.accounts.find(a => a.id === newTxData.accountId);
    if (!acc) return false;
    acc.balance = (parseFloat(acc.balance) || 0) + newAmount;
  } else if (newTxData.type === 'transfer') {
    if (newTxData.fromAccountId === newTxData.toAccountId) {
      showToast('Cannot transfer to the same account.');
      return false;
    }
    const fromAcc = state.accounts.find(a => a.id === newTxData.fromAccountId);
    const toAcc = state.accounts.find(a => a.id === newTxData.toAccountId);
    if (!fromAcc || !toAcc) return false;

    fromAcc.balance = (parseFloat(fromAcc.balance) || 0) - newAmount;
    toAcc.balance = (parseFloat(toAcc.balance) || 0) + newAmount;
  }

  // 3. Update properties
  oldTx.type = newTxData.type;
  oldTx.amount = newAmount;
  oldTx.category = newTxData.category;
  oldTx.datetime = newTxData.datetime || getDeviceLocalDateTime();
  oldTx.accountId = newTxData.accountId || null;
  oldTx.fromAccountId = newTxData.fromAccountId || null;
  oldTx.toAccountId = newTxData.toAccountId || null;
  oldTx.note = (newTxData.note || '').trim();

  saveData('accounts');
  saveData('transactions');

  renderDashboard();
  renderTransactionsView();
  renderStatsView();
  renderSetupView();

  showToast('Transaction updated!');
  triggerHapticFeedback();
  return true;
}

function deleteTransaction(txId) {
  const index = state.transactions.findIndex(t => t.id === txId);
  if (index === -1) return;

  const tx = state.transactions[index];
  const amount = parseFloat(tx.amount) || 0;

  // Reversal of account balances
  if (tx.type === 'expense') {
    const acc = state.accounts.find(a => a.id === tx.accountId);
    if (acc) acc.balance = (parseFloat(acc.balance) || 0) + amount;
  } else if (tx.type === 'income') {
    const acc = state.accounts.find(a => a.id === tx.accountId);
    if (acc) acc.balance = (parseFloat(acc.balance) || 0) - amount;
  } else if (tx.type === 'transfer') {
    const fromAcc = state.accounts.find(a => a.id === tx.fromAccountId);
    const toAcc = state.accounts.find(a => a.id === tx.toAccountId);
    if (fromAcc) fromAcc.balance = (parseFloat(fromAcc.balance) || 0) + amount;
    if (toAcc) toAcc.balance = (parseFloat(toAcc.balance) || 0) - amount;
  }

  state.transactions.splice(index, 1);

  saveData('accounts');
  saveData('transactions');

  renderDashboard();
  renderTransactionsView();
  renderStatsView();
  renderSetupView();

  showToast('Transaction deleted and balance restored.');
}

// ================= RECORD MODAL DYNAMICS =================
function openRecordModal(defaultType = 'expense') {
  const modal = document.getElementById('record-modal');
  const form = document.getElementById('record-tx-form');
  form.reset();

  const editInput = document.getElementById('edit-tx-id');
  if (editInput) editInput.value = '';

  const modalTitle = document.getElementById('record-modal-title');
  if (modalTitle) modalTitle.textContent = 'Record Transaction';

  const submitLabel = document.getElementById('record-submit-label');
  if (submitLabel) submitLabel.textContent = 'Save Transaction';

  // Set type radio
  const radio = form.querySelector(`input[name="tx-type"][value="${defaultType}"]`);
  if (radio) radio.checked = true;

  updateRecordFormFields(defaultType);

  // Set default datetime to device local time
  const timeInput = document.getElementById('tx-datetime');
  timeInput.value = getDeviceLocalDateTime();
  setLiveTimeMode('auto');

  modal.classList.add('show');
  document.getElementById('tx-amount').focus();
}

function openRecordModalForEdit(txId) {
  const tx = state.transactions.find(t => t.id === txId);
  if (!tx) return;

  const modal = document.getElementById('record-modal');
  const form = document.getElementById('record-tx-form');
  form.reset();

  const editInput = document.getElementById('edit-tx-id');
  if (editInput) editInput.value = tx.id;

  const modalTitle = document.getElementById('record-modal-title');
  if (modalTitle) modalTitle.textContent = 'Edit Transaction';

  const submitLabel = document.getElementById('record-submit-label');
  if (submitLabel) submitLabel.textContent = 'Update Transaction';

  // Select type
  const radio = form.querySelector(`input[name="tx-type"][value="${tx.type}"]`);
  if (radio) radio.checked = true;

  updateRecordFormFields(tx.type);

  // Amount
  document.getElementById('tx-amount').value = tx.amount;

  // Datetime
  setLiveTimeMode('manual');
  const timeInput = document.getElementById('tx-datetime');
  if (tx.datetime) {
    timeInput.value = tx.datetime.slice(0, 16);
  }

  // Category
  populateCategoryDropdown(tx.type, tx.category);

  // Account
  if (tx.type === 'transfer') {
    if (tx.fromAccountId) document.getElementById('tx-from-account').value = tx.fromAccountId;
    if (tx.toAccountId) document.getElementById('tx-to-account').value = tx.toAccountId;
  } else {
    if (tx.accountId) document.getElementById('tx-account').value = tx.accountId;
  }

  // Note
  document.getElementById('tx-note').value = tx.note || '';

  modal.classList.add('show');
}

function closeRecordModal() {
  document.getElementById('record-modal').classList.remove('show');
  const editInput = document.getElementById('edit-tx-id');
  if (editInput) editInput.value = '';

  const modalTitle = document.getElementById('record-modal-title');
  if (modalTitle) modalTitle.textContent = 'Record Transaction';

  const submitLabel = document.getElementById('record-submit-label');
  if (submitLabel) submitLabel.textContent = 'Save Transaction';

  clearInterval(state.timeAutoInterval);
}

function setLiveTimeMode(mode) {
  state.timeMode = mode;
  const autoBtn = document.getElementById('time-auto-btn');
  const manualBtn = document.getElementById('time-manual-btn');
  const statusText = document.getElementById('time-status-text');
  const timeInput = document.getElementById('tx-datetime');

  clearInterval(state.timeAutoInterval);

  if (mode === 'auto') {
    autoBtn.classList.add('active');
    manualBtn.classList.remove('active');
    statusText.textContent = 'Clock synced with mobile device';
    timeInput.value = getDeviceLocalDateTime();

    // Keep live-updating every 10 seconds while modal is open
    state.timeAutoInterval = setInterval(() => {
      if (state.timeMode === 'auto') {
        timeInput.value = getDeviceLocalDateTime();
      }
    }, 10000);
  } else {
    manualBtn.classList.add('active');
    autoBtn.classList.remove('active');
    statusText.textContent = 'Manual date & time selected';
  }
}

function updateRecordFormFields(type) {
  // Category dropdown
  populateCategoryDropdown(type);

  // Account inputs
  const singleGroup = document.getElementById('account-single-group');
  const transferGroup = document.getElementById('account-transfer-group');

  if (type === 'transfer') {
    singleGroup.classList.add('hidden');
    transferGroup.classList.remove('hidden');
    populateTransferAccountSelects();
  } else {
    singleGroup.classList.remove('hidden');
    transferGroup.classList.add('hidden');
    populateSingleAccountSelect();
  }
}

function populateCategoryDropdown(type, selectedCategory = null) {
  const select = document.getElementById('tx-category');
  select.innerHTML = '';

  const categories = state.categories[type] || [];
  categories.forEach(cat => {
    const opt = document.createElement('option');
    opt.value = cat;
    opt.textContent = cat;
    if (selectedCategory && cat === selectedCategory) opt.selected = true;
    select.appendChild(opt);
  });

  // Add category option
  const addOpt = document.createElement('option');
  addOpt.value = '__ADD_NEW__';
  addOpt.textContent = '➕ Add New Category...';
  select.appendChild(addOpt);
}

function populateSingleAccountSelect() {
  const select = document.getElementById('tx-account');
  select.innerHTML = '';
  state.accounts.forEach(a => {
    const opt = document.createElement('option');
    opt.value = a.id;
    opt.textContent = `${a.name} (${formatCurrency(a.balance)})`;
    select.appendChild(opt);
  });
}

function populateTransferAccountSelects() {
  const fromSelect = document.getElementById('tx-from-account');
  const toSelect = document.getElementById('tx-to-account');
  fromSelect.innerHTML = '';
  toSelect.innerHTML = '';

  state.accounts.forEach((a, idx) => {
    const opt1 = document.createElement('option');
    opt1.value = a.id;
    opt1.textContent = a.name;
    fromSelect.appendChild(opt1);

    const opt2 = document.createElement('option');
    opt2.value = a.id;
    opt2.textContent = a.name;
    // Default to a different account for destination
    if (idx === 1 || (idx === 0 && state.accounts.length === 1)) opt2.selected = true;
    toSelect.appendChild(opt2);
  });
}

function handleInlineCategoryAdd() {
  const form = document.getElementById('record-tx-form');
  const activeType = form.querySelector('input[name="tx-type"]:checked').value;
  const input = document.getElementById('inline-new-cat-name');
  const catName = input.value.trim();

  if (!catName) return;

  if (!state.categories[activeType].includes(catName)) {
    state.categories[activeType].push(catName);
    saveData('categories');
  }

  // Refresh dropdown and select this category
  populateCategoryDropdown(activeType, catName);
  document.getElementById('inline-cat-input-group').classList.add('hidden');
  input.value = '';
  showToast(`Category "${catName}" added!`);
}

// ================= EXPORT EXCEL LOGIC =================
function openExcelModal() {
  const modal = document.getElementById('excel-modal');
  updateExcelExportStats();
  modal.classList.add('show');
}

function closeExcelModal() {
  document.getElementById('excel-modal').classList.remove('show');
}

function getFilteredExportTransactions() {
  const timeline = document.getElementById('export-timeline-select').value;
  const now = new Date();

  return state.transactions.filter(t => {
    if (!t.datetime) return false;
    const tDate = new Date(t.datetime);

    if (timeline === 'current-month') {
      const curMonthStr = now.toISOString().slice(0, 7);
      return t.datetime.startsWith(curMonthStr);
    } else if (timeline === 'last-month') {
      const prevDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      const prevMonthStr = prevDate.toISOString().slice(0, 7);
      return t.datetime.startsWith(prevMonthStr);
    } else if (timeline === 'last-3-months') {
      const threeMonthsAgo = new Date(now.getFullYear(), now.getMonth() - 3, 1);
      return tDate >= threeMonthsAgo && tDate <= now;
    } else if (timeline === 'this-year') {
      return tDate.getFullYear() === now.getFullYear();
    } else if (timeline === 'custom') {
      const start = document.getElementById('export-start-date').value;
      const end = document.getElementById('export-end-date').value;
      if (start && t.datetime.slice(0, 10) < start) return false;
      if (end && t.datetime.slice(0, 10) > end) return false;
      return true;
    }
    // all-time
    return true;
  });
}

function updateExcelExportStats() {
  const list = getFilteredExportTransactions();
  const incCount = list.filter(t => t.type === 'income').length;
  const expCount = list.filter(t => t.type === 'expense').length;
  const trfCount = list.filter(t => t.type === 'transfer').length;

  document.getElementById('export-count-income').textContent = incCount;
  document.getElementById('export-count-expense').textContent = expCount;
  document.getElementById('export-count-transfer').textContent = trfCount;
}

function exportExcelFile() {
  if (typeof XLSX === 'undefined') {
    showToast('Excel exporter library loading...');
    return;
  }

  const list = getFilteredExportTransactions();

  if (list.length === 0) {
    showToast('No transactions found in selected timeline to export.');
    return;
  }

  // Create separate sheets
  const incomeRows = [];
  const expenseRows = [];
  const transferRows = [];

  list.forEach(t => {
    const dateFormatted = formatDisplayDate(t.datetime);
    const timeFormatted = formatDisplayTime(t.datetime);
    const amt = parseFloat(t.amount) || 0;

    if (t.type === 'income') {
      const acc = state.accounts.find(a => a.id === t.accountId)?.name || 'Account';
      incomeRows.push({
        'Date': dateFormatted,
        'Time': timeFormatted,
        'Amount': amt,
        'Category': t.category,
        'Account': acc,
        'Note': t.note || ''
      });
    } else if (t.type === 'expense') {
      const acc = state.accounts.find(a => a.id === t.accountId)?.name || 'Account';
      expenseRows.push({
        'Date': dateFormatted,
        'Time': timeFormatted,
        'Amount': amt,
        'Category': t.category,
        'Account': acc,
        'Note': t.note || ''
      });
    } else if (t.type === 'transfer') {
      const fromAcc = state.accounts.find(a => a.id === t.fromAccountId)?.name || 'Account';
      const toAcc = state.accounts.find(a => a.id === t.toAccountId)?.name || 'Account';
      transferRows.push({
        'Date': dateFormatted,
        'Time': timeFormatted,
        'Amount': amt,
        'Category / Reason': t.category,
        'From Account': fromAcc,
        'To Account': toAcc,
        'Note': t.note || ''
      });
    }
  });

  // Summary Row Sheet
  const totalInc = incomeRows.reduce((sum, r) => sum + r.Amount, 0);
  const totalExp = expenseRows.reduce((sum, r) => sum + r.Amount, 0);
  const netSavings = totalInc - totalExp;

  const summaryRows = [
    { 'Metric': 'Export Date', 'Value': new Date().toLocaleString() },
    { 'Metric': 'Total Income', 'Value': totalInc },
    { 'Metric': 'Total Expenses', 'Value': totalExp },
    { 'Metric': 'Net Savings', 'Value': netSavings },
    { 'Metric': '--- Account Balances ---', 'Value': '----------------' }
  ];

  state.accounts.forEach(a => {
    summaryRows.push({
      'Metric': a.name,
      'Value': a.balance
    });
  });

  // Build Workbook
  const wb = XLSX.utils.book_new();

  // Sheet 1: Income
  const wsIncome = XLSX.utils.json_to_sheet(incomeRows.length ? incomeRows : [{ Note: 'No income records' }]);
  XLSX.utils.book_append_sheet(wb, wsIncome, 'Income');

  // Sheet 2: Expense
  const wsExpense = XLSX.utils.json_to_sheet(expenseRows.length ? expenseRows : [{ Note: 'No expense records' }]);
  XLSX.utils.book_append_sheet(wb, wsExpense, 'Expense');

  // Sheet 3: Transfer
  const wsTransfer = XLSX.utils.json_to_sheet(transferRows.length ? transferRows : [{ Note: 'No transfer records' }]);
  XLSX.utils.book_append_sheet(wb, wsTransfer, 'Transfer');

  // Sheet 4: Summary
  const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  // Generate Filename
  const timeline = document.getElementById('export-timeline-select').value;
  const fileName = `Finances_${timeline}_${new Date().toISOString().slice(0, 10)}.xlsx`;

  // Download trigger
  XLSX.writeFile(wb, fileName);
  showToast('Excel spreadsheet downloaded!');
  closeExcelModal();
}

// ================= STATS & CHARTS UI =================
function renderStatsView() {
  const currentMonth = state.selectedTxMonth;
  document.getElementById('stats-chart-month-label').textContent = formatMonthTitle(currentMonth);

  let totalIncome = 0;
  let totalExpense = 0;
  const expenseByCategory = {};
  const incomeByCategory = {};

  state.transactions.forEach(t => {
    if (t.datetime && t.datetime.startsWith(currentMonth)) {
      const amt = parseFloat(t.amount) || 0;
      if (t.type === 'income') {
        totalIncome += amt;
        incomeByCategory[t.category] = (incomeByCategory[t.category] || 0) + amt;
      } else if (t.type === 'expense') {
        totalExpense += amt;
        expenseByCategory[t.category] = (expenseByCategory[t.category] || 0) + amt;
      }
    }
  });

  const netSavings = totalIncome - totalExpense;
  const savingsRate = totalIncome > 0 ? Math.round((netSavings / totalIncome) * 100) : 0;

  document.getElementById('stats-total-income').textContent = formatCurrency(totalIncome);
  document.getElementById('stats-total-expense').textContent = formatCurrency(totalExpense);
  document.getElementById('stats-net-savings').textContent = formatCurrency(netSavings);
  document.getElementById('stats-net-savings').className = `stat-value ${netSavings >= 0 ? 'text-income' : 'text-expense'}`;
  document.getElementById('stats-savings-rate').textContent = `${savingsRate}%`;

  renderExpenseChart(expenseByCategory);
  renderIncomeChart(incomeByCategory);
  renderMonthlyTrendChart();
}

function formatMonthTitle(monthStr) {
  const [year, month] = monthStr.split('-');
  const dateObj = new Date(parseInt(year), parseInt(month) - 1, 1);
  return dateObj.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
}

function renderExpenseChart(dataObj) {
  const ctx = document.getElementById('expenseCategoryChart');
  if (!ctx || typeof Chart === 'undefined') return;

  const labels = Object.keys(dataObj);
  const values = Object.values(dataObj);

  const colors = [
    '#f43f5e', '#fb923c', '#eab308', '#22c55e', '#06b6d4',
    '#6366f1', '#a855f7', '#ec4899', '#14b8a6', '#64748b'
  ];

  if (expenseChartInstance) expenseChartInstance.destroy();

  if (labels.length === 0) {
    document.getElementById('expense-category-breakdown-list').innerHTML = `
      <div class="empty-placeholder"><p>No expenses for this month.</p></div>
    `;
    return;
  }

  expenseChartInstance = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: colors.slice(0, labels.length),
        borderWidth: 2,
        borderColor: '#131823'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false }
      },
      cutout: '68%'
    }
  });

  // Render breakdown list below chart
  const breakdownList = document.getElementById('expense-category-breakdown-list');
  breakdownList.innerHTML = '';
  const total = values.reduce((s, v) => s + v, 0);

  labels.forEach((cat, idx) => {
    const val = values[idx];
    const pct = total > 0 ? Math.round((val / total) * 100) : 0;
    const row = document.createElement('div');
    row.className = 'breakdown-row';
    row.innerHTML = `
      <div class="breakdown-label">
        <span class="breakdown-color-box" style="background:${colors[idx % colors.length]};"></span>
        <span>${cat}</span>
      </div>
      <div>
        <strong>${formatCurrency(val)}</strong>
        <span style="font-size:11px;color:var(--text-dim);margin-left:4px;">(${pct}%)</span>
      </div>
    `;
    breakdownList.appendChild(row);
  });
}

function renderIncomeChart(dataObj) {
  const ctx = document.getElementById('incomeCategoryChart');
  if (!ctx || typeof Chart === 'undefined') return;

  const labels = Object.keys(dataObj);
  const values = Object.values(dataObj);

  const colors = ['#10b981', '#3b82f6', '#8b5cf6', '#f59e0b', '#06b6d4'];

  if (incomeChartInstance) incomeChartInstance.destroy();

  if (labels.length === 0) {
    document.getElementById('income-category-breakdown-list').innerHTML = `
      <div class="empty-placeholder"><p>No income records this month.</p></div>
    `;
    return;
  }

  incomeChartInstance = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: colors.slice(0, labels.length),
        borderWidth: 2,
        borderColor: '#131823'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      cutout: '68%'
    }
  });

  const breakdownList = document.getElementById('income-category-breakdown-list');
  breakdownList.innerHTML = '';
  labels.forEach((cat, idx) => {
    const val = values[idx];
    const row = document.createElement('div');
    row.className = 'breakdown-row';
    row.innerHTML = `
      <div class="breakdown-label">
        <span class="breakdown-color-box" style="background:${colors[idx % colors.length]};"></span>
        <span>${cat}</span>
      </div>
      <strong>${formatCurrency(val)}</strong>
    `;
    breakdownList.appendChild(row);
  });
}

function renderMonthlyTrendChart() {
  const ctx = document.getElementById('monthlyTrendChart');
  if (!ctx || typeof Chart === 'undefined') return;

  // Gather last 6 months data
  const months = [];
  const cur = new Date();
  for (let i = 5; i >= 0; i--) {
    const d = new Date(cur.getFullYear(), cur.getMonth() - i, 1);
    months.push(d.toISOString().slice(0, 7));
  }

  const incomeData = [];
  const expenseData = [];
  const monthLabels = [];

  months.forEach(m => {
    const [y, mon] = m.split('-');
    const label = new Date(parseInt(y), parseInt(mon) - 1, 1).toLocaleDateString('en-US', { month: 'short' });
    monthLabels.push(label);

    let inc = 0;
    let exp = 0;
    state.transactions.forEach(t => {
      if (t.datetime && t.datetime.startsWith(m)) {
        if (t.type === 'income') inc += parseFloat(t.amount) || 0;
        if (t.type === 'expense') exp += parseFloat(t.amount) || 0;
      }
    });
    incomeData.push(inc);
    expenseData.push(exp);
  });

  if (monthlyTrendChartInstance) monthlyTrendChartInstance.destroy();

  monthlyTrendChartInstance = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: monthLabels,
      datasets: [
        {
          label: 'Income',
          data: incomeData,
          backgroundColor: '#10b981',
          borderRadius: 6
        },
        {
          label: 'Expense',
          data: expenseData,
          backgroundColor: '#f43f5e',
          borderRadius: 6
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { grid: { display: false }, ticks: { color: '#94a3b8' } },
        y: { grid: { color: 'rgba(255,255,255,0.06)' }, ticks: { color: '#94a3b8' } }
      },
      plugins: {
        legend: { labels: { color: '#f8fafc', boxWidth: 12 } }
      }
    }
  });
}

// ================= SETUP & CUSTOMIZATION VIEW =================
let activeCategorySetupTab = 'expense';

function renderSetupView() {
  renderSetupAccounts();
  renderSetupCategories();
  document.getElementById('pref-currency').value = state.settings.currency;
  document.getElementById('pref-theme').value = state.settings.theme;

  // Update example shortcut URL
  const loc = window.location.origin + window.location.pathname;
  document.getElementById('shortcut-example-url').textContent = `${loc}?action=add&type=expense&amount=50&category=Food&account=acc_sbi`;
}

function renderSetupAccounts() {
  const container = document.getElementById('setup-accounts-list');
  container.innerHTML = '';

  state.accounts.forEach(a => {
    const item = document.createElement('div');
    item.className = 'account-manage-item';
    item.innerHTML = `
      <div>
        <strong style="font-size:14px;color:var(--text-main);">${a.name}</strong>
        <span style="font-size:12px;color:var(--text-muted);display:block;margin-top:2px;">${formatCurrency(a.balance)}</span>
      </div>
      <button class="ios-link-btn" onclick="openAccountEditModal('${a.id}')">Edit</button>
    `;
    container.appendChild(item);
  });
}

function renderSetupCategories() {
  const container = document.getElementById('setup-category-chips');
  container.innerHTML = '';

  const list = state.categories[activeCategorySetupTab] || [];
  list.forEach((cat, idx) => {
    const chip = document.createElement('div');
    chip.className = 'cat-chip';
    chip.innerHTML = `
      <span>${cat}</span>
      <button class="cat-chip-delete" onclick="deleteCategory('${activeCategorySetupTab}', ${idx})">&times;</button>
    `;
    container.appendChild(chip);
  });
}

function deleteCategory(type, index) {
  if (state.categories[type].length <= 1) {
    showToast('Must keep at least one category.');
    return;
  }
  state.categories[type].splice(index, 1);
  saveData('categories');
  renderSetupCategories();
  populateCategorySelects();
  showToast('Category removed.');
}

function openAccountEditModal(accId) {
  const modal = document.getElementById('account-modal');
  const form = document.getElementById('account-form');
  form.reset();

  if (accId) {
    const acc = state.accounts.find(a => a.id === accId);
    if (!acc) return;
    document.getElementById('account-modal-title').textContent = 'Edit Account';
    document.getElementById('edit-account-id').value = acc.id;
    document.getElementById('account-name-input').value = acc.name;
    document.getElementById('account-balance-input').value = acc.balance;
    document.getElementById('account-type-select').value = acc.type || 'bank';
  } else {
    document.getElementById('account-modal-title').textContent = 'Add New Account';
    document.getElementById('edit-account-id').value = '';
    document.getElementById('account-balance-input').value = '0';
  }

  modal.classList.add('show');
}

function closeAccountModal() {
  document.getElementById('account-modal').classList.remove('show');
}

// ================= THEME & CURRENCY =================
function applyTheme(themeClass) {
  document.body.classList.remove('theme-dark', 'theme-slate', 'theme-light');
  document.body.classList.add(themeClass);
  state.settings.theme = themeClass;
  saveData('settings');
}

function updateCurrencyDisplays() {
  document.getElementById('input-currency-symbol').textContent = state.settings.currency;
  document.getElementById('account-currency-symbol').textContent = state.settings.currency;
}

// ================= TOAST & HAPTICS =================
function showToast(msg) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.classList.add('show');
  setTimeout(() => {
    toast.classList.remove('show');
  }, 2400);
}

function triggerHapticFeedback() {
  if (navigator.vibrate) {
    try { navigator.vibrate(20); } catch (e) {}
  }
}

// ================= SHORTCUTS / QUERY PARAM INTEGRATION =================
function checkShortcutsQueryParam() {
  const params = new URLSearchParams(window.location.search);
  const action = params.get('action');

  if (action === 'add' || params.get('amount')) {
    const type = params.get('type') || 'expense';
    const amount = parseFloat(params.get('amount'));
    const category = params.get('category');
    const account = params.get('account');
    const note = params.get('note') || 'Logged via Shortcut';

    if (!isNaN(amount) && amount > 0) {
      // Find matching account id
      let accId = state.accounts[0]?.id;
      if (account) {
        const found = state.accounts.find(a => a.id.toLowerCase().includes(account.toLowerCase()) || a.name.toLowerCase().includes(account.toLowerCase()));
        if (found) accId = found.id;
      }

      // Find or add category
      let cat = state.categories[type]?.[0] || 'Food';
      if (category) {
        const foundCat = state.categories[type]?.find(c => c.toLowerCase() === category.toLowerCase());
        if (foundCat) cat = foundCat;
        else {
          state.categories[type].push(category);
          saveData('categories');
          cat = category;
        }
      }

      addTransaction({
        type: type,
        amount: amount,
        category: cat,
        accountId: accId,
        datetime: getDeviceLocalDateTime(),
        note: note
      });

      // Clear query string to prevent double submission on reload
      window.history.replaceState({}, document.title, window.location.pathname);
    } else {
      // Open modal pre-populated
      openRecordModal(type);
    }
  }
}

// ================= EVENT LISTENERS =================
function setupEventListeners() {
  // Bottom Tab Navigation
  document.querySelectorAll('.nav-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const targetView = tab.getAttribute('data-view');
      switchView(targetView);
    });
  });

  // FAB Center Plus Button & Header Record Button
  document.getElementById('fab-record-btn').addEventListener('click', () => openRecordModal('expense'));
  document.getElementById('quick-add-header-btn').addEventListener('click', () => openRecordModal('expense'));
  document.getElementById('open-record-modal-btn').addEventListener('click', () => openRecordModal('expense'));

  // Close Record Modal
  document.getElementById('close-record-modal-btn').addEventListener('click', closeRecordModal);
  document.getElementById('record-modal').addEventListener('click', (e) => {
    if (e.target.id === 'record-modal') closeRecordModal();
  });

  // Type change in Record Form
  document.querySelectorAll('input[name="tx-type"]').forEach(radio => {
    radio.addEventListener('change', (e) => {
      updateRecordFormFields(e.target.value);
    });
  });

  // Time mode toggle buttons
  document.getElementById('time-auto-btn').addEventListener('click', () => setLiveTimeMode('auto'));
  document.getElementById('time-manual-btn').addEventListener('click', () => setLiveTimeMode('manual'));
  document.getElementById('tx-datetime').addEventListener('input', () => setLiveTimeMode('manual'));

  // Inline Category Add toggle
  document.getElementById('tx-category').addEventListener('change', (e) => {
    if (e.target.value === '__ADD_NEW__') {
      document.getElementById('inline-cat-input-group').classList.remove('hidden');
      document.getElementById('inline-new-cat-name').focus();
    } else {
      document.getElementById('inline-cat-input-group').classList.add('hidden');
    }
  });

  document.getElementById('quick-add-cat-btn').addEventListener('click', () => {
    document.getElementById('inline-cat-input-group').classList.remove('hidden');
    document.getElementById('inline-new-cat-name').focus();
  });

  document.getElementById('inline-save-cat-btn').addEventListener('click', handleInlineCategoryAdd);
  document.getElementById('inline-cancel-cat-btn').addEventListener('click', () => {
    document.getElementById('inline-cat-input-group').classList.add('hidden');
    const type = document.querySelector('input[name="tx-type"]:checked').value;
    populateCategoryDropdown(type);
  });

  // Record Form Submit
  document.getElementById('record-tx-form').addEventListener('submit', (e) => {
    e.preventDefault();
    const type = document.querySelector('input[name="tx-type"]:checked').value;
    const amount = document.getElementById('tx-amount').value;
    let category = document.getElementById('tx-category').value;
    const datetime = document.getElementById('tx-datetime').value;
    const note = document.getElementById('tx-note').value;

    if (category === '__ADD_NEW__') {
      const inlineVal = document.getElementById('inline-new-cat-name').value.trim();
      if (!inlineVal) {
        showToast('Please type a category name.');
        return;
      }
      category = inlineVal;
      if (!state.categories[type].includes(category)) {
        state.categories[type].push(category);
        saveData('categories');
      }
    }

    let txPayload = {
      type,
      amount,
      category,
      datetime,
      note
    };

    if (type === 'transfer') {
      txPayload.fromAccountId = document.getElementById('tx-from-account').value;
      txPayload.toAccountId = document.getElementById('tx-to-account').value;
    } else {
      txPayload.accountId = document.getElementById('tx-account').value;
    }

    const editId = document.getElementById('edit-tx-id')?.value;
    let success = false;
    if (editId) {
      success = updateTransaction(editId, txPayload);
    } else {
      success = addTransaction(txPayload);
    }

    if (success) {
      closeRecordModal();
    }
  });

  // Export Excel Modal
  document.getElementById('open-excel-modal-btn').addEventListener('click', openExcelModal);
  document.getElementById('close-excel-modal-btn').addEventListener('click', closeExcelModal);
  document.getElementById('excel-modal').addEventListener('click', (e) => {
    if (e.target.id === 'excel-modal') closeExcelModal();
  });
  document.getElementById('export-timeline-select').addEventListener('change', (e) => {
    const customGroup = document.getElementById('custom-range-group');
    if (e.target.value === 'custom') customGroup.classList.remove('hidden');
    else customGroup.classList.add('hidden');
    updateExcelExportStats();
  });
  document.getElementById('export-start-date').addEventListener('change', updateExcelExportStats);
  document.getElementById('export-end-date').addEventListener('change', updateExcelExportStats);
  document.getElementById('confirm-export-excel-btn').addEventListener('click', exportExcelFile);

  // Month navigation in Transactions tab
  document.getElementById('prev-month-btn').addEventListener('click', () => changeMonth(-1));
  document.getElementById('next-month-btn').addEventListener('click', () => changeMonth(1));
  document.getElementById('tx-month-select').addEventListener('change', (e) => {
    state.selectedTxMonth = e.target.value;
    updateMonthPickerUI();
    renderTransactionsView();
    renderStatsView();
  });

  // Transactions Search & Filters
  document.getElementById('tx-search-input').addEventListener('input', (e) => {
    const clearBtn = document.getElementById('tx-search-clear');
    if (e.target.value) clearBtn.classList.remove('hidden');
    else clearBtn.classList.add('hidden');
    renderDateGroupedTransactions();
  });

  document.getElementById('tx-search-clear').addEventListener('click', () => {
    document.getElementById('tx-search-input').value = '';
    document.getElementById('tx-search-clear').classList.add('hidden');
    renderDateGroupedTransactions();
  });

  document.getElementById('tx-sort-select').addEventListener('change', renderDateGroupedTransactions);
  document.getElementById('filter-account-select').addEventListener('change', renderDateGroupedTransactions);
  document.getElementById('filter-category-select').addEventListener('change', renderDateGroupedTransactions);

  // Transactions Segmented Switcher (All / Income / Expense / Transfer)
  document.querySelectorAll('.segment-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.segment-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.activeTxSegment = btn.getAttribute('data-col');
      renderDateGroupedTransactions();
    });
  });

  // Setup View Add & Edit Accounts
  document.getElementById('setup-add-account-btn').addEventListener('click', () => openAccountEditModal(null));
  document.getElementById('edit-accounts-quick-btn').addEventListener('click', () => switchView('setup'));
  document.getElementById('see-all-tx-btn').addEventListener('click', () => switchView('transactions'));
  document.getElementById('close-account-modal-btn').addEventListener('click', closeAccountModal);
  document.getElementById('account-modal').addEventListener('click', (e) => {
    if (e.target.id === 'account-modal') closeAccountModal();
  });

  document.getElementById('account-form').addEventListener('submit', (e) => {
    e.preventDefault();
    const id = document.getElementById('edit-account-id').value;
    const name = document.getElementById('account-name-input').value.trim();
    const balance = parseFloat(document.getElementById('account-balance-input').value) || 0;
    const type = document.getElementById('account-type-select').value;

    if (!name) return;

    if (id) {
      const acc = state.accounts.find(a => a.id === id);
      if (acc) {
        acc.name = name;
        acc.balance = balance;
        acc.type = type;
      }
    } else {
      state.accounts.push({
        id: 'acc_' + Date.now(),
        name,
        balance,
        type
      });
    }

    saveData('accounts');
    renderDashboard();
    renderSetupAccounts();
    populateCategorySelects();
    closeAccountModal();
    showToast('Account updated!');
  });

  // Setup Category Sub-tabs
  document.querySelectorAll('.sub-tab-btn').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.sub-tab-btn').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      activeCategorySetupTab = tab.getAttribute('data-category-type');
      renderSetupCategories();
    });
  });

  document.getElementById('setup-add-category-btn').addEventListener('click', () => {
    const input = document.getElementById('setup-new-category-input');
    const val = input.value.trim();
    if (!val) return;
    if (!state.categories[activeCategorySetupTab].includes(val)) {
      state.categories[activeCategorySetupTab].push(val);
      saveData('categories');
      input.value = '';
      renderSetupCategories();
      populateCategorySelects();
      showToast(`Category "${val}" added!`);
    } else {
      showToast('Category already exists.');
    }
  });

  // Preferences Change
  document.getElementById('pref-currency').addEventListener('change', (e) => {
    state.settings.currency = e.target.value;
    saveData('settings');
    updateCurrencyDisplays();
    renderDashboard();
    renderTransactionsView();
    renderStatsView();
    renderSetupView();
    showToast('Currency updated');
  });

  document.getElementById('pref-theme').addEventListener('change', (e) => {
    applyTheme(e.target.value);
  });

  document.getElementById('theme-toggle-btn').addEventListener('click', () => {
    const current = state.settings.theme;
    const next = current === 'theme-dark' ? 'theme-light' : current === 'theme-light' ? 'theme-slate' : 'theme-dark';
    applyTheme(next);
    document.getElementById('pref-theme').value = next;
    showToast(`Theme switched to ${next.replace('theme-', '')}`);
  });

  // Excel Data Import
  const excelInput = document.getElementById('import-excel-file');
  if (excelInput) excelInput.addEventListener('change', handleExcelImport);

  // Force Refresh & Cache Buster for iOS
  const forceRefreshBtn = document.getElementById('force-refresh-btn');
  if (forceRefreshBtn) {
    forceRefreshBtn.addEventListener('click', async () => {
      showToast('Clearing cache and reloading...');
      try {
        if ('caches' in window) {
          const keys = await caches.keys();
          await Promise.all(keys.map(k => caches.delete(k)));
        }
        if ('serviceWorker' in navigator) {
          const registrations = await navigator.serviceWorker.getRegistrations();
          for (let reg of registrations) {
            await reg.unregister();
          }
        }
      } catch (err) {
        console.error('Cache clear error:', err);
      }
      setTimeout(() => {
        window.location.href = window.location.origin + window.location.pathname + '?t=' + Date.now();
      }, 300);
    });
  }

  // Backup & Restore
  document.getElementById('export-backup-btn').addEventListener('click', exportJsonBackup);
  document.getElementById('import-backup-file').addEventListener('change', importJsonBackup);
  document.getElementById('reset-all-data-btn').addEventListener('click', resetAllData);

  // Test Shortcut Button
  document.getElementById('test-shortcut-btn').addEventListener('click', () => {
    const sample = `${window.location.origin}${window.location.pathname}?action=add&type=expense&amount=50&category=Food&account=acc_sbi&note=ShortcutTest`;
    window.location.href = sample;
  });
}

function switchView(viewName) {
  state.currentView = viewName;
  document.querySelectorAll('.view-section').forEach(sec => sec.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(tab => tab.classList.remove('active'));

  const targetSec = document.getElementById(`view-${viewName}`);
  const targetTab = document.querySelector(`.nav-tab[data-view="${viewName}"]`);
  const pageTitle = document.getElementById('page-title');

  if (targetSec) targetSec.classList.add('active');
  if (targetTab) targetTab.classList.add('active');

  const titles = {
    dashboard: 'Dashboard',
    transactions: 'Transactions',
    stats: 'Analytics',
    setup: 'Settings'
  };
  pageTitle.textContent = titles[viewName] || 'FinanceFlow';

  if (viewName === 'stats') renderStatsView();
  if (viewName === 'transactions') renderTransactionsView();
  if (viewName === 'setup') renderSetupView();
}

function changeMonth(delta) {
  const [year, month] = state.selectedTxMonth.split('-');
  const newDate = new Date(parseInt(year), parseInt(month) - 1 + delta, 1);
  state.selectedTxMonth = newDate.toISOString().slice(0, 7);
  updateMonthPickerUI();
  renderTransactionsView();
  renderStatsView();
}

function populateCategorySelects() {
  populateFilterOptions();
}

function populateAccountsUI() {
  populateSingleAccountSelect();
  populateTransferAccountSelects();
}

// ================= BACKUP & RESTORE =================
function exportJsonBackup() {
  const fullBackup = {
    accounts: state.accounts,
    categories: state.categories,
    transactions: state.transactions,
    settings: state.settings,
    version: '1.0',
    exportedAt: new Date().toISOString()
  };

  const blob = new Blob([JSON.stringify(fullBackup, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `FinanceFlow_Backup_${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  URL.revokeObjectURL(url);
  showToast('Backup file downloaded.');
}

function importJsonBackup(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = JSON.parse(e.target.result);
      if (data.accounts && data.categories && data.transactions) {
        state.accounts = data.accounts;
        state.categories = data.categories;
        state.transactions = data.transactions;
        if (data.settings) state.settings = { ...DEFAULT_SETTINGS, ...data.settings };

        saveData();
        applyTheme(state.settings.theme);
        updateCurrencyDisplays();
        renderDashboard();
        renderTransactionsView();
        renderStatsView();
        renderSetupView();
        showToast('Backup restored successfully!');
      } else {
        showToast('Invalid backup file structure.');
      }
    } catch (err) {
      showToast('Failed to parse backup JSON file.');
    }
  };
  reader.readAsText(file);
}

function resetAllData() {
  if (confirm('Are you sure you want to reset all accounts and transactions? This cannot be undone.')) {
    localStorage.clear();
    loadStoredData();
    applyTheme(state.settings.theme);
    updateCurrencyDisplays();
    renderDashboard();
    renderTransactionsView();
    renderStatsView();
    renderSetupView();
    showToast('All data reset to initial defaults.');
  }
}
