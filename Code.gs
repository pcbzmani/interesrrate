// ============================================================
// Asaktivatti — Google Apps Script Backend
// ============================================================
// SETUP:
//   1. Open Google Sheets > Extensions > Apps Script
//   2. Paste this entire file, replacing all existing code
//   3. Deploy > New Deployment > Web App
//      Execute as: Me | Who can access: Anyone
//   4. Copy the deployment URL into Asaktivatti Settings
// ============================================================

const SHEETS = { BORROWERS: 'Borrowers', TRANSACTIONS: 'Transactions', TRANSFERS: 'Transfers' };

// RateUnit values: PA=% per annum | PM=% per month | PAISE=paise/₹/month | PER100=₹/₹100/month
const B_HDR  = ['ID','Name','Phone','Principal','InterestRate','Months','StartDate',
                 'InterestType','RateUnit','Status','Notes','CreatedAt'];
const T_HDR  = ['ID','BorrowerID','Date','Type','Amount','Description','RunningBalance','CreatedAt'];
const TR_HDR = ['ID','FromID','ToID','Amount','Date','Notes','CreatedAt'];

// ── Entry points ───────────────────────────────────────────
function doGet(e) {
  try {
    return jsonOut(dispatch(e.parameter.action, e.parameter, null));
  } catch(err) { return jsonOut({error: err.message}); }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    return jsonOut(dispatch(d.action, null, d));
  } catch(err) { return jsonOut({error: err.message}); }
}

function dispatch(action, p, d) {
  switch(action) {
    case 'ping':               return {ok:true, ts:new Date().toISOString()};
    case 'getBorrowers':       return getBorrowers();
    case 'getTransactions':    return getTransactions(p ? p.id : null);
    case 'getAllTransactions':  return getTransactions(null);
    case 'getSummary':         return getSummary();
    case 'addBorrower':        return addBorrower(d.borrower);
    case 'updateBorrower':     return updateBorrower(d.borrower);
    case 'deleteBorrower':     return setStatus(d.id, 'DELETED');
    case 'closeLoan':          return setStatus(d.id, 'CLOSED');
    case 'reopenLoan':         return setStatus(d.id, 'ACTIVE');
    case 'recordPayment':      return recordTxn(d.borrowerId,'PAYMENT',d.amount,d.date,d.description,-1);
    case 'recordDisbursement': return recordTxn(d.borrowerId,'DISBURSEMENT',d.amount,d.date,d.description,1);
    case 'recordTransfer':     return processTransfer(d);
    default: return {error:'Unknown action: ' + action};
  }
}

// ── Helpers ────────────────────────────────────────────────
function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(headers);
    s.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#dbeafe');
    s.setFrozenRows(1);
    s.setColumnWidths(1, headers.length, 150);
  }
  return s;
}

function toObjects(s, headers) {
  const data = s.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => {
    const o = {}; headers.forEach((h,i) => o[h] = row[i]); return o;
  });
}

function uid() {
  return Date.now().toString(36).toUpperCase() +
         Math.random().toString(36).slice(2,5).toUpperCase();
}

// ── Interest rate conversion ────────────────────────────────
// All calculations use decimal monthly rate internally
// PA  = % per annum    → divide by 12 and by 100
// PM  = % per month    → divide by 100
// PAISE = paise/₹/month  → same as % per month (1 paise/₹ = 1%)
// PER100 = ₹/₹100/month → same as % per month
function toMonthlyRate(rate, unit) {
  switch(unit) {
    case 'PM':     return rate / 100;
    case 'PAISE':  return rate / 100;
    case 'PER100': return rate / 100;
    default:       return rate / 12 / 100; // PA
  }
}

// ── Loan math ──────────────────────────────────────────────
function calcStats(principal, rate, months, loanType, rateUnit) {
  if (!principal || !months) return {monthlyPayment:0, totalInterest:0, totalRepayable:principal||0};
  const r = toMonthlyRate(rate || 0, rateUnit || 'PA');
  let monthly, interest;

  switch(loanType) {
    case 'SIMPLE':
      interest = principal * r * months;
      monthly  = (principal + interest) / months;
      break;
    case 'FLAT':
      monthly  = (principal / months) + (principal * r);
      interest = monthly * months - principal;
      break;
    case 'INTEREST_ONLY':
      monthly  = principal * r;
      interest = monthly * months;
      break;
    default: // EMI — reducing balance
      monthly  = r === 0
        ? principal / months
        : principal * r * Math.pow(1+r,months) / (Math.pow(1+r,months) - 1);
      interest = monthly * months - principal;
  }

  return {
    monthlyPayment: round2(monthly),
    totalInterest:  round2(Math.max(0, interest)),
    totalRepayable: round2(principal + Math.max(0, interest))
  };
}

function round2(n) { return Math.round(n * 100) / 100; }

// ── Borrowers ──────────────────────────────────────────────
function getBorrowers() {
  const s    = getOrCreateSheet(SHEETS.BORROWERS, B_HDR);
  const rows = toObjects(s, B_HDR);
  const txns = getTransactions(null);
  const today = new Date();

  return rows
    .filter(b => b.Status !== 'DELETED')
    .map(b => {
      const myTxns   = txns.filter(t => t.BorrowerID === b.ID);
      const lastTxn  = myTxns[myTxns.length - 1];
      const outstanding = lastTxn
        ? round2(parseFloat(lastTxn.RunningBalance) || 0)
        : round2(parseFloat(b.Principal) || 0);
      const paid = round2(myTxns
        .filter(t => t.Type === 'PAYMENT')
        .reduce((s,t) => s + (parseFloat(t.Amount)||0), 0));

      const endDate = new Date(b.StartDate);
      endDate.setMonth(endDate.getMonth() + parseInt(b.Months || 0));
      const isOverdue = b.Status === 'ACTIVE' && outstanding > 0.01 && today > endDate;

      const stats = calcStats(
        parseFloat(b.Principal), parseFloat(b.InterestRate),
        parseInt(b.Months), b.InterestType, b.RateUnit || 'PA'
      );

      return {
        ...b,
        outstanding, paid, isOverdue,
        endDate: endDate.toISOString().split('T')[0],
        stats
      };
    });
}

function addBorrower(b) {
  const s  = getOrCreateSheet(SHEETS.BORROWERS, B_HDR);
  const id = uid();
  s.appendRow([
    id,
    b.name,
    b.phone || '',
    parseFloat(b.principal),
    parseFloat(b.interestRate),
    parseInt(b.months),
    b.startDate,
    b.interestType || 'EMI',
    b.rateUnit     || 'PA',
    'ACTIVE',
    b.notes || '',
    new Date().toISOString()
  ]);
  recordTxn(id, 'DISBURSEMENT', parseFloat(b.principal), b.startDate, 'Loan disbursement', 1);
  return {success:true, id};
}

function updateBorrower(b) {
  const s    = getOrCreateSheet(SHEETS.BORROWERS, B_HDR);
  const data = s.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][B_HDR.indexOf('ID')] === b.ID) {
      const r = i + 1;
      const updates = {
        Name: b.Name, Phone: b.Phone || '',
        InterestRate: parseFloat(b.InterestRate),
        RateUnit: b.RateUnit || 'PA',
        Months:   parseInt(b.Months),
        Status:   b.Status || 'ACTIVE',
        Notes:    b.Notes  || ''
      };
      Object.entries(updates).forEach(([key,val]) => {
        const col = B_HDR.indexOf(key);
        if (col >= 0) s.getRange(r, col+1).setValue(val);
      });
      return {success:true};
    }
  }
  return {error:'Borrower not found'};
}

function setStatus(id, status) {
  const s    = getOrCreateSheet(SHEETS.BORROWERS, B_HDR);
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][B_HDR.indexOf('ID')] === id) {
      s.getRange(i+1, B_HDR.indexOf('Status')+1).setValue(status);
      return {success:true};
    }
  }
  return {error:'Not found'};
}

// ── Transactions ───────────────────────────────────────────
function getTransactions(borrowerId) {
  const s = getOrCreateSheet(SHEETS.TRANSACTIONS, T_HDR);
  const txns = toObjects(s, T_HDR);
  return borrowerId ? txns.filter(t => t.BorrowerID === borrowerId) : txns;
}

function recordTxn(borrowerId, type, amount, date, desc, sign) {
  const s    = getOrCreateSheet(SHEETS.TRANSACTIONS, T_HDR);
  const prev = getTransactions(borrowerId);
  const prevBal = prev.length > 0 ? (parseFloat(prev[prev.length-1].RunningBalance) || 0) : 0;
  const newBal  = round2(prevBal + sign * parseFloat(amount));
  const id      = uid();

  s.appendRow([
    id, borrowerId,
    date || new Date().toISOString().split('T')[0],
    type, parseFloat(amount), desc || type,
    newBal, new Date().toISOString()
  ]);

  if (type === 'PAYMENT' && newBal <= 0.01) setStatus(borrowerId, 'CLOSED');
  return {success:true, id, newBalance:newBal};
}

// ── Transfers (P2P debt reassignment) ──────────────────────
function processTransfer(d) {
  const s      = getOrCreateSheet(SHEETS.TRANSFERS, TR_HDR);
  const amount = parseFloat(d.amount);
  const date   = d.date || new Date().toISOString().split('T')[0];
  const bors   = getBorrowers();
  const fromName = (bors.find(b => b.ID === d.fromBorrowerId) || {}).Name || d.fromBorrowerId;
  const toName   = (bors.find(b => b.ID === d.toBorrowerId)   || {}).Name || d.toBorrowerId;

  recordTxn(d.fromBorrowerId, 'PAYMENT',      amount, date, `Transfer to ${toName}`,   -1);
  recordTxn(d.toBorrowerId,   'DISBURSEMENT', amount, date, `Transfer from ${fromName}`, 1);

  const id = uid();
  s.appendRow([id, d.fromBorrowerId, d.toBorrowerId, amount, date,
               d.notes || '', new Date().toISOString()]);
  return {success:true, id};
}

// ── Summary ────────────────────────────────────────────────
function getSummary() {
  const bors   = getBorrowers();
  const txns   = getTransactions(null);
  const active  = bors.filter(b => b.Status === 'ACTIVE');
  const closed  = bors.filter(b => b.Status === 'CLOSED');
  const overdue = bors.filter(b => b.isOverdue);

  const totalPrincipal   = round2(bors.reduce((s,b) => s+(parseFloat(b.Principal)||0), 0));
  const totalOutstanding = round2(active.reduce((s,b) => s+(b.outstanding||0), 0));
  const totalCollected   = round2(txns.filter(t => t.Type==='PAYMENT')
                                      .reduce((s,t) => s+(parseFloat(t.Amount)||0), 0));
  const totalInterest    = round2(bors.reduce((s,b) => s+(b.stats?.totalInterest||0), 0));

  return {
    totalPrincipal, totalOutstanding, totalCollected, totalInterest,
    activeBorrowers:  active.length,
    closedBorrowers:  closed.length,
    overdueBorrowers: overdue.length,
    totalBorrowers:   bors.length
  };
}
