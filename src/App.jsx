import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import XLSXStyle from "xlsx-js-style";
import IncomeRegistration from "./components/IncomeRegistration.jsx";
import ExpenseRegistration from "./components/ExpenseRegistration.jsx";
import { supabase } from "./supabaseClient.js";

const STORAGE_KEY = "dr-hibist-v3";
const AUTH_KEY = "dr-hibist-auth-v1";
const DARK_KEY = "dr-hibist-dark-v1";
const LANG_KEY = "dr-hibist-lang-v1";
const LOGO_KEY = "dr-hibist-logo-v1";

const DEPARTMENTS = [
  "Doctor", "Nurse", "Reception", "Laboratory",
  "Cashier", "Security", "Cleaner", "Finance", "CEO", "Manager",
  "Pharmacy", "ጉዳይ አስፈሳሚ", "Other",
];

const DOCTOR_BASIS_LABELS = {
  all: "ሁሉም አገልግሎቶች",
  card: "ካርድ",
  lab: "ላብራቶሪ",
  imaging: "ኢሜጂንግ",
  consultation: "ኮንሰልቴሽን",
  procedure: "ፕሮሲጀር",
  round: "ራውንድ",
};

// Admin-level menu items
const ADMIN_ITEMS = [
  "Dashboard",
  "HRM",
  "Approve Expenses",
  "Income Registration",
  "Expense Registration",
  "Debt Registration",
  "Reports",
  "Balance Sheet",
  "Create User",
  "Announcements",
  "Messages",
  "Resignations",
  "Settings",
];
const ADMIN_LABELS = {
  en: ["Dashboard", "HRM", "Approve Expenses", "Income Registration", "Expense Registration", "Debt Registration", "Reports", "Balance Sheet", "Create User", "Announcements", "Messages", "Resignations", "Settings"],
  am: ["ዳሽቦርድ", "ተሰራተኞች", "ወጪ አፕሩቭ", "ገቢ ምዝገባ", "ወጪ ምዝገባ", "እዳ ምዝገባ", "ሪፖርቶች", "የሃብትና እዳ ምዝገባ", "ዩዘር ፍጠር", "ማስታወቂያ", "መልዕክቶች", "ስራ መልቀቃ", "ቅንብሮች"],
};

// Finance sub-menu items (inside "Finance" department section)
const FINANCE_ITEMS = [
  "Income Registration",
  "Expense Registration",
  "Debt Registration",
  "Payroll",
];
const FINANCE_LABELS = {
  en: ["Income Registration", "Expense Registration", "Debt Registration", "Payroll"],
  am: ["ገቢ መመዝገቢያ", "ወጪ መመዝገቢያ", "እዳ መመዝገቢያ", "ፔሮል"],
};

const SECTION_TITLES = {
  Dashboard: "Finance Dashboard",
  HRM: "HRM — የሰራተኞች አስተዳደር",
  "Approve Expenses": "Approve Expenses — ወጪ አፕሩቭ",
  Reports: "Reports — ሪፖርቶች",
  "Create User": "Create User — ሰራተኛ አካውንት ፍጠር",
  Announcements: "Announcements — ማስታወቂያ",
  Messages: "Messages — መልዕክቶች",
  Resignations: "Resignations — ስራ መልቀቃ",
  Settings: "Settings",
  "Balance Sheet": "Balance Sheet — የሃብትና እዳ ምዝገባ",
  "Income Registration": "Income Registration",
  "Expense Registration": "Expense Registration",
  "Debt Registration": "Debt Registration",
  Payroll: "Payroll — ፔሮል",
};

const T = {
  en: {
    signIn: "Sign in with your phone number and 4-digit password.",
    password: "Password", login: "Login", createAccount: "Create New Account",
    hideAccount: "Hide Form", fullName: "Full Name", newCreditAccount: "New Credit Account",
    saveNewAccount: "Save Account", darkMode: "Dark Mode", lightMode: "Light Mode", language: "AM",
    actions: "Actions", delete: "Delete",
    confirmDeleteIncome: "Delete this income entry?",
    confirmDeleteExpense: "Delete this expense entry?",
    confirmDeleteDebt: "Delete this debt entry?",
    deletedIncome: "Income deleted.", deletedExpense: "Expense deleted.", deletedDebt: "Debt deleted.",
  },
  am: {
    signIn: "በስልክ ቁጥር እና በ4 አሃዝ ፓስዋርድ ይግቡ።",
    password: "ፓስዋርድ", login: "ግባ", createAccount: "አዲስ አካውንት ፍጠር",
    hideAccount: "ቅጽ ደብቅ", fullName: "ሙሉ ስም", newCreditAccount: "ክሬዲት አካውንት",
    saveNewAccount: "አካውንት አስቀምጥ", darkMode: "ጨለማ", lightMode: "ብርሃን", language: "EN",
    actions: "ድርጊት", delete: "ሰርዝ",
    confirmDeleteIncome: "ይህን የገቢ መዝገብ ለማጥፋት እርግጠኛ ነዎት?",
    confirmDeleteExpense: "ይህን የወጪ መዝገብ ለማጥፋት እርግጠኛ ነዎት?",
    confirmDeleteDebt: "ይህን የጥቅል መዝገብ ለማጥፋት እርግጠኛ ነዎት?",
    deletedIncome: "የገቢ መዝገብ ተሰርዟል።", deletedExpense: "የወጪ ተሰርዟል።", deletedDebt: "የጥቅል ተሰርዟል።",
  },
};

const DEFAULT = {
  organizationName: "Dr Hibist Pediatrics and Medical Center",
  organizationLogo: "", organizationLogoText: "Dr Hibist Pediatrics and Medical Center",
  loginAccount: { fullName: "Dr Hibist Admin", phoneNumber: "0913765565", email: "office@drhibist.com", password: "1234", creditAccount: "ACC-001" },
  incomeEntries: [],
  expenseEntries: [],
  bankAccounts: [],
  debts: [],
  employees: [],
  leaveRequests: [],
  messages: [],
  announcements: [],
  resignations: [],
  balanceSheet: {
    asOfDate: "",
    // Current Assets
    cashAndBank: 0,
    accountsReceivable: 0,
    inventory: 0,
    // Non-Current Assets
    equipmentAndVehicles: 0,
    accumulatedDepreciation: 0,
    // Current Liabilities
    accountsPayable: 0,
    taxPayable: 0,
    // Owner's Equity
    capital: 0,
    retainedEarnings: 0,
  },
};

function today() {
  return new Date().toISOString().slice(0, 10);
}

function loadDb() {
  try {
    const p = JSON.parse(localStorage.getItem(STORAGE_KEY) || "null");
    // Logo is stored separately so it never gets wiped
    const savedLogo = localStorage.getItem(LOGO_KEY) || "";
    if (!p) return { ...DEFAULT, organizationLogo: savedLogo };
    return {
      ...DEFAULT, ...p,
      organizationLogo: savedLogo || p.organizationLogo || "",
      loginAccount: { ...DEFAULT.loginAccount, ...(p.loginAccount || {}) },
      employees: p.employees || [], leaveRequests: p.leaveRequests || [], messages: p.messages || [],
      announcements: p.announcements || [], resignations: p.resignations || [],
      balanceSheet: p.balanceSheet ? { ...DEFAULT.balanceSheet, ...p.balanceSheet } : DEFAULT.balanceSheet,
    };
  } catch { return DEFAULT; }
}
function loadBool(key) { return localStorage.getItem(key) === "true"; }
function loadStr(key, def = "") { return localStorage.getItem(key) || def; }

export default function App() {

  // ========== CORE STATE ==========
  const [db, setDb] = useState(() => loadDb());
  const [isAdmin, setIsAdmin] = useState(() => loadBool(AUTH_KEY));
  const [dark, setDark] = useState(() => loadBool(DARK_KEY));
  const [lang, setLang] = useState(() => loadStr(LANG_KEY, "en"));
  const [section, setSection] = useState("Dashboard");
  const [hrmTab, setHrmTab] = useState("employees");
  const [financeTab, setFinanceTab] = useState("income");
  const [reportTab, setReportTab] = useState("monthly");
  const [toast, setToast] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [menuExpanded, setMenuExpanded] = useState(true);
  const [showDashboardCards, setShowDashboardCards] = useState(false);

  // Auth / portal
  const [loginForm, setLoginForm] = useState({ phoneNumber: "", password: "" });
  const [showNewAcc, setShowNewAcc] = useState(false);
  const [newAccForm, setNewAccForm] = useState({ fullName: "", creditAccount: "", phoneNumber: "", email: "", password: "" });
  const [showEmpPortal, setShowEmpPortal] = useState(false);
  const [empLoggedIn, setEmpLoggedIn] = useState(null);
  const [empLoginForm, setEmpLoginForm] = useState({ phone: "", password: "" });
  const [empMsgBody, setEmpMsgBody] = useState("");

  // Forms
  const [incomeForm, setIncomeForm] = useState({ category: "Clinic income", amount: "", notes: "", bankId: "", date: today() });
  const [expenseForm, setExpenseForm] = useState({ category: "Salary", amount: "", notes: "", bankId: "", date: today() });
  const [bankForm, setBankForm] = useState({ name: "", accountNumber: "", initialBalance: "" });
  const [debtForm, setDebtForm] = useState({ organization: "", total: "", paid: "", status: "Pending", date: today() });
  const [empForm, setEmpForm] = useState({ name: "", department: "Doctor", phone: "", bankAccount: "", hireDate: today(), basicSalary: "" });
  const [leaveForm, setLeaveForm] = useState({ empId: "", days: "", startDate: today(), reason: "", leaveType: "annual" });
  const [msgForm, setMsgForm] = useState({ toId: "", body: "" });
  const [createUserForm, setCreateUserForm] = useState({ empId: "", phone: "", password: "", department: "" });
  const [leaveDaysForm, setLeaveDaysForm] = useState({ empId: "", days: "25" });
  const [payrollBankId, setPayrollBankId] = useState("");
  const [payrollDate, setPayrollDate] = useState(today());
  const [payrollEdit, setPayrollEdit] = useState({});
  const [bsForm, setBsForm] = useState({});
  const [bsEditing, setBsEditing] = useState(false);
  // Leave approval with adjustment
  const [leaveApproveEdit, setLeaveApproveEdit] = useState({}); // {reqId: adjustedDays}
  const [leaveRejectReason, setLeaveRejectReason] = useState({}); // {reqId: reason}
  // Announcements
  const [announcementForm, setAnnouncementForm] = useState({ title: "", body: "" });
  // Messages: selected thread employee
  const [msgThreadEmpId, setMsgThreadEmpId] = useState("");
  // Resignations
  const [resignForm, setResignForm] = useState({ reason: "", lastDay: "" });
  // Chat input for inline thread reply
  const [threadReply, setThreadReply] = useState("");
  // Employee portal extras
  const [empPortalTab, setEmpPortalTab] = useState("home"); // home | messages | announcements | leave | resign
  const [empChangePwForm, setEmpChangePwForm] = useState({ current: "", newPw: "", confirm: "" });
  const [empPeerMsgForm, setEmpPeerMsgForm] = useState({ toId: "", body: "" });
  // Expense inline edit
  const [editExpenseId, setEditExpenseId] = useState(null);
  const [editExpenseForm, setEditExpenseForm] = useState({});
  // Income inline edit
  const [editIncomeId, setEditIncomeId] = useState(null);
  const [editIncomeForm, setEditIncomeForm] = useState({});
  // Debt inline edit
  const [editDebtId, setEditDebtId] = useState(null);
  const [editDebtForm, setEditDebtForm] = useState({});
  // Custom department
  const [customDept, setCustomDept] = useState("");
  // Report date range filter
  const [reportDateFrom, setReportDateFrom] = useState("");
  const [reportDateTo, setReportDateTo] = useState("");

  // Doctor weekly revenue calculator
  const [doctorWeeklyForm, setDoctorWeeklyForm] = useState({});
  // { [empId]: { cardCount:"", cardPrice:"", labTotal:"", imagingCount:"", imagingPrice:"", duty:"", procedure:"", round:"", bankId:"" } }
  const [doctorWeeklyHistory, setDoctorWeeklyHistory] = useState(() => {
    try { return JSON.parse(localStorage.getItem("dr-hibist-weekly-v1") || "{}"); } catch { return {}; }
  });

  // ========== LOAD FROM SUPABASE ON STARTUP ==========
  useEffect(() => {
    const loadFromSupabase = async () => {
      try {
        const { data, error } = await supabase
          .from('clinic_data')
          .select('data')
          .eq('id', 'app_state')
          .single();

        if (error || !data) {
          console.log('No remote data found — using local data.');
          return;
        }

        const remote = data.data;
        if (!remote) return;

        // Merge: remote data ን local ን ጋር ማዋሃድ
        const savedLogo = localStorage.getItem(LOGO_KEY) || remote.organizationLogo || '';
        const merged = {
          ...DEFAULT,
          ...remote,
          organizationLogo: savedLogo,
          loginAccount: { ...DEFAULT.loginAccount, ...(remote.loginAccount || {}) },
          employees: remote.employees || [],
          leaveRequests: remote.leaveRequests || [],
          messages: remote.messages || [],
          announcements: remote.announcements || [],
          resignations: remote.resignations || [],
          balanceSheet: remote.balanceSheet
            ? { ...DEFAULT.balanceSheet, ...remote.balanceSheet }
            : DEFAULT.balanceSheet,
        };
        setDb(merged);
        console.log('Data loaded from Supabase successfully.');
      } catch (err) {
        console.error('Failed to load from Supabase:', err);
      }
    };
    loadFromSupabase();
  }, []); // once on mount

  // ========== AUTO-SAVE TO SUPABASE ==========
  useEffect(() => {
    if (!db) return;
    const autoSaveTimer = setTimeout(async () => {
      try {
        const { error } = await supabase
          .from('clinic_data')
          .upsert({ id: 'app_state', data: db });
        if (error) {
          console.error('Error auto-syncing to Supabase:', error);
        } else {
          console.log('Changes synced to Supabase automatically.');
        }
      } catch (err) {
        console.error('Auto-sync failed:', err);
      }
    }, 1500);
    return () => clearTimeout(autoSaveTimer);
  }, [db]);

  useEffect(() => { localStorage.setItem(STORAGE_KEY, JSON.stringify(db)); }, [db]);
  useEffect(() => { localStorage.setItem(AUTH_KEY, String(isAdmin)); }, [isAdmin]);
  useEffect(() => { localStorage.setItem(DARK_KEY, String(dark)); }, [dark]);
  useEffect(() => { localStorage.setItem(LANG_KEY, lang); }, [lang]);
  useEffect(() => { localStorage.setItem("dr-hibist-weekly-v1", JSON.stringify(doctorWeeklyHistory)); }, [doctorWeeklyHistory]);
  useEffect(() => {
    if (!toast) return;
    const id = setTimeout(() => setToast(null), 2600);
    return () => clearTimeout(id);
  }, [toast]);

  const t = T[lang] || T.en;

  // ---------- computed ----------
  const bankSummaries = useMemo(() => (db.bankAccounts || []).map(b => {
    const inc = db.incomeEntries.filter(e => Number(e.bankId) === b.id).reduce((s, e) => s + Number(e.amount || 0), 0);
    // Only approved expenses deduct from balance
    const exp = db.expenseEntries.filter(e => Number(e.bankId) === b.id && (e.status === "Approved" || !e.status)).reduce((s, e) => s + Number(e.amount || 0), 0);
    return { ...b, balance: Number(b.initialBalance || 0) + inc - exp };
  }), [db]);

  const yearly = useMemo(() => {
    const inc = db.incomeEntries.reduce((s, e) => s + Number(e.amount || 0), 0);
    const exp = db.expenseEntries.filter(e => e.status === "Approved" || !e.status).reduce((s, e) => s + Number(e.amount || 0), 0);
    const td = db.debts.reduce((s, d) => s + Number(d.total || 0), 0);
    const pd = db.debts.reduce((s, d) => s + Number(d.paid || 0), 0);
    return { totalIncome: inc, totalExpense: exp, netProfit: inc - exp, totalDebt: td, paidDebt: pd, remainingDebt: td - pd };
  }, [db]);

  const monthly = useMemo(() => {
    const now = new Date(); const yr = now.getFullYear(); const mo = now.getMonth();
    const f = arr => arr.filter(e => { const d = new Date(e.date); return d.getFullYear() === yr && d.getMonth() === mo; });
    const mInc = f(db.incomeEntries).reduce((s, e) => s + Number(e.amount || 0), 0);
    const mExp = f(db.expenseEntries).filter(e => e.status === "Approved" || !e.status).reduce((s, e) => s + Number(e.amount || 0), 0);
    const mDebt = f(db.debts).reduce((s, d) => s + Number(d.total || 0), 0);
    const mPaid = f(db.debts).reduce((s, d) => s + Number(d.paid || 0), 0);
    return { monthIncome: mInc, monthExpense: mExp, monthDebt: mDebt, monthPaidDebt: mPaid, remainingProfit: mInc - mExp };
  }, [db]);

  const groupedExpenses = useMemo(() => {
    const m = {};
    db.expenseEntries.filter(e => e.status === "Approved" || !e.status).forEach(e => { m[e.category] = (m[e.category] || 0) + Number(e.amount || 0); });
    return Object.entries(m).map(([cat, amt]) => ({ cat, amt }));
  }, [db]);

  const deptStats = useMemo(() => {
    const m = {};
    (db.employees || []).forEach(e => { m[e.department] = (m[e.department] || 0) + 1; });
    return Object.entries(m).map(([dept, count]) => ({ dept, count })).sort((a, b) => b.count - a.count);
  }, [db]);

  const unreadAdmin = useMemo(() =>
    (db.messages || []).filter(m => m.to === "admin" && !m.readByAdmin).length
    , [db]);

  const unreadEmp = useMemo(() => {
    if (!empLoggedIn) return 0;
    return (db.messages || []).filter(m => m.to === String(empLoggedIn.id) && !m.readByEmp).length;
  }, [db, empLoggedIn]);

  const pendingLeave = useMemo(() =>
    (db.leaveRequests || []).filter(r => r.status === "Pending").length
    , [db]);

  const pendingExpenses = useMemo(() =>
    (db.expenseEntries || []).filter(e => e.status === "Pending")
    , [db]);

  const pendingResignations = useMemo(() =>
    (db.resignations || []).filter(r => r.status === "Pending").length
    , [db]);

  const unreadAnnouncements = useMemo(() => {
    if (!empLoggedIn) return 0;
    return (db.announcements || []).filter(a => !(a.readBy || []).includes(empLoggedIn.id)).length;
  }, [db, empLoggedIn]);

  // ---------- handlers ----------
  function showToast(type, message) { setToast({ type, message }); }

  function handleLogin(e) {
    e.preventDefault();
    const ph = loginForm.phoneNumber.trim(); const pw = loginForm.password.trim();
    if (!ph || !pw) { showToast("error", "Phone and password required."); return; }
    const acc = db.loginAccount || {};
    if (ph !== acc.phoneNumber?.trim() || pw !== acc.password?.trim()) { showToast("error", "Incorrect credentials."); return; }
    setIsAdmin(true); showToast("success", "Login successful.");
  }

  function handleCreateAccount(e) {
    e.preventDefault();
    const { fullName, creditAccount, phoneNumber, email, password } = newAccForm;
    if (!fullName.trim() || !phoneNumber.trim() || !password.trim()) { showToast("error", "Fill required fields."); return; }
    if (!/^\d{4}$/.test(password.trim())) { showToast("error", "Password must be 4 digits."); return; }
    setDb(p => ({
      ...p, loginAccount: { fullName: fullName.trim(), phoneNumber: phoneNumber.trim(), email: email.trim(), password: password.trim(), creditAccount: creditAccount.trim() },
      incomeEntries: [], expenseEntries: [], bankAccounts: [], debts: [], employees: [], leaveRequests: [], messages: []
    }));
    setNewAccForm({ fullName: "", creditAccount: "", phoneNumber: "", email: "", password: "" });
    setShowNewAcc(false); showToast("success", "Account created. Data reset.");
  }

  const addIncome = (e) => {
    e.preventDefault();
    if (!incomeForm.category || !incomeForm.amount || !incomeForm.bankId) { showToast("error", "Fill all fields."); return; }
    const amt = Number(incomeForm.amount); if (isNaN(amt) || amt <= 0) { showToast("error", "Invalid amount."); return; }
    setDb(p => ({ ...p, incomeEntries: [{ id: Date.now(), bankId: Number(incomeForm.bankId), amount: amt, category: incomeForm.category, source: incomeForm.notes ? `${incomeForm.category} — ${incomeForm.notes}` : incomeForm.category, date: incomeForm.date }, ...p.incomeEntries] }));
    setIncomeForm(p => ({ ...p, amount: "", notes: "", bankId: "", category: "Clinic income" })); showToast("success", "ገቢ ተቀምጧል።");
  };

  const addExpense = (e) => {
    e.preventDefault();
    if (!expenseForm.category || !expenseForm.amount || !expenseForm.bankId) { showToast("error", "Fill all fields."); return; }
    const amt = Number(expenseForm.amount); if (isNaN(amt) || amt <= 0) { showToast("error", "Invalid amount."); return; }
    setDb(p => ({ ...p, expenseEntries: [{ id: Date.now(), category: expenseForm.category, bankId: Number(expenseForm.bankId), amount: amt, company: "Manual", date: expenseForm.date, note: expenseForm.notes || "Expense", status: "Pending" }, ...p.expenseEntries] }));
    setExpenseForm(p => ({ ...p, amount: "", notes: "", bankId: "" })); showToast("success", "Expense submitted — pending approval.");
  };

  const addBank = (e) => {
    e.preventDefault();
    if (!bankForm.name || !bankForm.accountNumber) { showToast("error", "Name and account number required."); return; }
    setDb(p => ({ ...p, bankAccounts: [...(p.bankAccounts || []), { id: Date.now(), name: bankForm.name, accountNumber: bankForm.accountNumber, initialBalance: Number(bankForm.initialBalance || 0) }] }));
    setBankForm({ name: "", accountNumber: "", initialBalance: "" }); showToast("success", "Bank account added.");
  };

  const deleteBank = (id) => { if (!window.confirm("Delete bank account?")) return; setDb(p => ({ ...p, bankAccounts: p.bankAccounts.filter(b => b.id !== id) })); };

  const addDebt = () => {
    if (!debtForm.organization || !debtForm.total) { showToast("error", "Fill required fields."); return; }
    setDb(p => ({ ...p, debts: [{ id: Date.now(), organization: debtForm.organization, total: Number(debtForm.total), paid: Number(debtForm.paid || 0), status: debtForm.status, date: debtForm.date }, ...p.debts] }));
    setDebtForm({ organization: "", total: "", paid: "", status: "Pending", date: today() });
    showToast("success", "እዳ ተቀምጧል።");
  };

  const deleteIncome = (id) => { if (!window.confirm(t.confirmDeleteIncome)) return; setDb(p => ({ ...p, incomeEntries: p.incomeEntries.filter(e => e.id !== id) })); showToast("success", t.deletedIncome); };

  const startEditIncome = (inc) => {
    setEditIncomeId(inc.id);
    setEditIncomeForm({ bankId: String(inc.bankId), amount: String(inc.amount), date: inc.date, source: inc.source || "" });
  };
  const saveEditIncome = (id) => {
    if (!editIncomeForm.amount || !editIncomeForm.bankId) { showToast("error", "Bank እና Amount ያስፈልጋሉ።"); return; }
    setDb(p => ({
      ...p,
      incomeEntries: p.incomeEntries.map(e => e.id === id
        ? { ...e, bankId: Number(editIncomeForm.bankId), amount: Number(editIncomeForm.amount), date: editIncomeForm.date, source: editIncomeForm.source }
        : e
      ),
    }));
    setEditIncomeId(null); setEditIncomeForm({});
    showToast("success", "ገቢ ተዘምኗል።");
  };

  const deleteExpense = (id) => { if (!window.confirm(t.confirmDeleteExpense)) return; setDb(p => ({ ...p, expenseEntries: p.expenseEntries.filter(e => e.id !== id) })); showToast("success", t.deletedExpense); };

  const startEditExpense = (exp) => {
    setEditExpenseId(exp.id);
    setEditExpenseForm({ category: exp.category, amount: String(exp.amount), date: exp.date, note: exp.note || "", bankId: String(exp.bankId) });
  };

  const saveEditExpense = (id) => {
    if (!editExpenseForm.category || !editExpenseForm.amount) { showToast("error", "Category እና Amount ያስፈልጋሉ።"); return; }
    setDb(p => ({
      ...p,
      expenseEntries: p.expenseEntries.map(e => e.id === id
        ? { ...e, category: editExpenseForm.category, amount: Number(editExpenseForm.amount), date: editExpenseForm.date, note: editExpenseForm.note, bankId: Number(editExpenseForm.bankId) }
        : e
      ),
    }));
    setEditExpenseId(null);
    setEditExpenseForm({});
    showToast("success", "ወጪ ተዘምኗል።");
  };
  const deleteDebt = (id) => { if (!window.confirm(t.confirmDeleteDebt)) return; setDb(p => ({ ...p, debts: p.debts.filter(e => e.id !== id) })); showToast("success", t.deletedDebt); };

  const startEditDebt = (d) => {
    setEditDebtId(d.id);
    setEditDebtForm({ organization: d.organization, total: String(d.total), paid: String(d.paid), status: d.status, date: d.date });
  };
  const saveEditDebt = (id) => {
    if (!editDebtForm.organization || !editDebtForm.total) { showToast("error", "Organization እና Total ያስፈልጋሉ።"); return; }
    setDb(p => ({
      ...p,
      debts: p.debts.map(d => d.id === id
        ? { ...d, organization: editDebtForm.organization, total: Number(editDebtForm.total), paid: Number(editDebtForm.paid || 0), status: editDebtForm.status, date: editDebtForm.date }
        : d
      ),
    }));
    setEditDebtId(null); setEditDebtForm({});
    showToast("success", "እዳ ተዘምኗል።");
  };

  const resetAllData = () => {
    if (!window.confirm("⚠️ ሁሉንም ዳታ ማጥፋት ትፈልጋለህ?\n\nይህ ሁሉንም ገቢ፣ ወጪ፣ እዳ፣ ሰራተኞች፣ ባንኮች ያጠፋል።\nLogin account ብቻ ይቀራል።")) return;
    const saved = db.loginAccount;
    const savedLogo = db.organizationLogo;
    const savedLogoText = db.organizationLogoText;
    const savedOrgName = db.organizationName;
    setDb({
      ...DEFAULT,
      loginAccount: saved,
      organizationLogo: savedLogo,
      organizationLogoText: savedLogoText,
      organizationName: savedOrgName,
    });
    showToast("success", "ሁሉም ዳታ ተጥፏል።");
  };

  const approveExpense = (id) => {
    setDb(p => ({ ...p, expenseEntries: p.expenseEntries.map(e => e.id === id ? { ...e, status: "Approved", approvedDate: new Date().toISOString().slice(0, 10) } : e) }));
    showToast("success", "ወጪ ተፈቅዷል — ከባንክ ተቀንሷል።");
  };

  const rejectExpense = (id) => {
    setDb(p => ({ ...p, expenseEntries: p.expenseEntries.map(e => e.id === id ? { ...e, status: "Rejected" } : e) }));
    showToast("success", "ወጪ ተሰርዟል።");
  };

  // ---- HRM handlers ----
  const addEmployee = (e) => {
    e.preventDefault();
    const { name, department, phone, bankAccount, hireDate, basicSalary } = empForm;
    if (!name.trim() || !basicSalary) { showToast("error", "ስም እና ደሞዝ ያስፈልጋሉ።"); return; }
    const finalDept = department === "Other" ? (customDept.trim() || "Other") : department;
    const sal = Number(basicSalary);
    const emp = {
      id: Date.now(), name: name.trim(), department: finalDept,
      phone: phone.trim(), bankAccount: bankAccount.trim(), hireDate,
      basicSalary: sal, duty: 0, pension: 0, pensionOpt: false, netPay: sal,
      password: "1234", totalLeaveDays: 25, usedLeaveDays: 0,
    };
    setDb(p => ({ ...p, employees: [...(p.employees || []), emp] }));
    setEmpForm({ name: "", department: "Doctor", phone: "", bankAccount: "", hireDate: today(), basicSalary: "" });
    setCustomDept("");
    showToast("success", `${name} ተመዝግቧል።`);
  };

  const deleteEmployee = (id) => {
    if (!window.confirm("ሰራተኛ ይሰረዝ?")) return;
    setDb(p => ({ ...p, employees: p.employees.filter(e => e.id !== id) }));
  };

  const updateEmployeeLeave = (id, field, val) => {
    setDb(p => ({ ...p, employees: p.employees.map(e => e.id === id ? { ...e, [field]: Number(val) } : e) }));
  };

  // Update payroll fields (duty, pension opt-in, salary adjustment) per employee
  const savePayrollEdit = (empId) => {
    const edit = payrollEdit[empId];
    if (!edit) return;
    setDb(p => ({
      ...p,
      employees: p.employees.map(e => {
        if (e.id !== empId) return e;
        const sal = Number(edit.basicSalary ?? e.basicSalary);
        const dut = Number(edit.duty ?? e.duty);
        const pen = (edit.pensionOpt ?? e.pensionOpt) ? Math.round(sal * 0.07 * 100) / 100 : 0;
        const net = Math.round((sal - pen + dut) * 100) / 100;
        return { ...e, basicSalary: sal, duty: dut, pension: pen, pensionOpt: edit.pensionOpt ?? e.pensionOpt, netPay: net };
      }),
    }));
    setPayrollEdit(p => { const n = { ...p }; delete n[empId]; return n; });
    showToast("success", "ፔሮል ተዘምኗል።");
  };

  // ---- Doctor weekly revenue calculator ----
  const getDoctorForm = (empId) => doctorWeeklyForm[String(empId)] || { cardAmount: "", labAmount: "", imagingAmount: "", consultationAmount: "", procedureAmount: "", roundAmount: "", bankId: "", percentage: "15", basis: "all", doctorId: String(empId) };

  const setDoctorField = (empId, field, val) => {
    setDoctorWeeklyForm(p => ({ ...p, [String(empId)]: { ...getDoctorForm(empId), [field]: val } }));
  };

  const getDoctorWeeklyHistoryForEmp = (empId) => {
    const targetId = String(empId);
    return Object.values(doctorWeeklyHistory || {}).flatMap(list => (Array.isArray(list) ? list : [])).filter(entry => String(entry.doctorId || entry.empId || "") === targetId);
  };

  const calcDoctorWeekly = (empId) => {
    const f = getDoctorForm(empId);
    const card = Number(f.cardAmount ?? ((Number(f.cardCount) || 0) * (Number(f.cardPrice) || 0))) || 0;
    const lab = Number(f.labAmount ?? f.labTotal) || 0;
    const imaging = Number(f.imagingAmount ?? ((Number(f.imagingCount) || 0) * (Number(f.imagingPrice) || 0))) || 0;
    const consultation = Number(f.consultationAmount ?? f.duty) || 0;
    const procedure = Number(f.procedureAmount ?? f.procedure) || 0;
    const round = Number(f.roundAmount ?? f.round) || 0;
    const allTotal = card + lab + imaging + consultation + procedure + round;
    const basis = f.basis || "all";
    const selectedTotal = basis === "card"
      ? card
      : basis === "lab"
        ? lab
        : basis === "imaging"
          ? imaging
          : basis === "consultation"
            ? consultation
            : basis === "procedure"
              ? procedure
              : basis === "round"
                ? round
                : allTotal;
    const percent = Number.isFinite(Number(f.percentage)) && Number(f.percentage) > 0 ? Number(f.percentage) : 15;
    const doctorCut = Math.round(selectedTotal * (percent / 100) * 100) / 100;
    return { card, lab, imaging, consultation, procedure, round, total: allTotal, basis, selectedTotal, percent, doctorCut };
  };

  const saveDoctorWeekly = (emp) => {
    const calc = calcDoctorWeekly(emp.id);
    const f = getDoctorForm(emp.id);
    if (calc.total === 0) { showToast("error", "ምንም ዋጋ አልተሞላም።"); return; }
    if (!f.bankId) { showToast("error", "ክፍያ የሚከፈልበት ባንክ ይምረጡ።"); return; }
    const targetDoctorId = String(f.doctorId || emp.id);
    const entry = {
      id: Date.now(),
      date: today(),
      weekLabel: `ሳምንት — ${today()}`,
      cardAmount: f.cardAmount,
      labAmount: f.labAmount,
      imagingAmount: f.imagingAmount,
      consultationAmount: f.consultationAmount,
      procedureAmount: f.procedureAmount,
      roundAmount: f.roundAmount,
      bankId: f.bankId,
      percentage: f.percentage || "15",
      basis: f.basis || "all",
      doctorId: targetDoctorId,
      doctorName: (db.employees || []).find(e => String(e.id) === targetDoctorId)?.name || emp.name,
      status: "Pending",
      ...calc,
    };
    const key = targetDoctorId;
    setDoctorWeeklyHistory(p => ({
      ...p,
      [key]: [entry, ...(p[key] || [])].slice(0, 20),
    }));
    setDoctorWeeklyForm(p => ({ ...p, [String(emp.id)]: { cardAmount: "", labAmount: "", imagingAmount: "", consultationAmount: "", procedureAmount: "", roundAmount: "", bankId: "", percentage: "15", basis: "all", doctorId: String(emp.id) } }));
    showToast("success", `${emp.name} — ${calc.doctorCut.toLocaleString()} Birr (${calc.percent}%) ቀርቧል። Admin approval ይጠብቁ።`);
  };

  const deleteDoctorWeeklyEntry = (empId, entryId) => {
    const key = String(empId);
    setDoctorWeeklyHistory(p => ({ ...p, [key]: (p[key] || []).filter(e => e.id !== entryId) }));
  };

  const approveDoctorWeekly = (empId, entryId) => {
    const key = String(empId);
    setDoctorWeeklyHistory(p => ({
      ...p,
      [key]: (p[key] || []).map(e => e.id === entryId ? { ...e, status: "Approved", approvedDate: today() } : e),
    }));
    showToast("success", "የዶክተር ሳምንታዊ ክፍያ ተፈቅዷል።");
  };

  const rejectDoctorWeekly = (empId, entryId) => {
    const key = String(empId);
    setDoctorWeeklyHistory(p => ({
      ...p,
      [key]: (p[key] || []).map(e => e.id === entryId ? { ...e, status: "Rejected" } : e),
    }));
    showToast("success", "የዶክተር ሳምንታዊ ክፍያ ተሰርዟል።");
  };
  const createUserAccount = (e) => {
    e.preventDefault();
    if (!createUserForm.empId || !createUserForm.password) { showToast("error", "ሰራተኛ እና ፓስዋርድ ያስፈልጋሉ።"); return; }
    if (!/^\d{4,6}$/.test(createUserForm.password)) { showToast("error", "ፓስዋርድ 4 እስከ 6 ቁጥር ብቻ ይሁን።"); return; }
    setDb(p => ({
      ...p,
      employees: p.employees.map(em =>
        em.id === Number(createUserForm.empId)
          ? {
            ...em,
            password: createUserForm.password,
            hasPortalAccess: true,
            ...(createUserForm.phone.trim() ? { phone: createUserForm.phone.trim() } : {}),
            ...(createUserForm.department ? { department: createUserForm.department } : {}),
          }
          : em
      ),
    }));
    setCreateUserForm({ empId: "", phone: "", password: "", department: "" });
    showToast("success", "Portal access ተሰጠ።");
  };

  const revokeUserAccess = (empId) => {
    if (!window.confirm("Portal access ይሰረዝ?")) return;
    setDb(p => ({
      ...p,
      employees: p.employees.map(em => em.id === empId ? { ...em, hasPortalAccess: false, password: "" } : em),
    }));
    showToast("success", "Access ተሰረዘ።");
  };

  // Admin sets annual leave days for employee
  const setEmployeeLeaveDays = (e) => {
    e.preventDefault();
    if (!leaveDaysForm.empId || !leaveDaysForm.days) { showToast("error", "ሰራተኛ እና ቀናት ይምረጡ።"); return; }
    setDb(p => ({
      ...p,
      employees: p.employees.map(em =>
        em.id === Number(leaveDaysForm.empId) ? { ...em, totalLeaveDays: Number(leaveDaysForm.days) } : em
      ),
    }));
    setLeaveDaysForm({ empId: "", days: "25" });
    showToast("success", "የፍቃድ ቀናት ተዘምኗል።");
  };

  const payAllSalaries = () => {
    const emps = db.employees || [];
    if (!emps.length) { showToast("error", "ምንም ሰራተኛ የለም።"); return; }
    if (!payrollBankId) { showToast("error", "ባንክ ምረጥ።"); return; }
    const expenses = emps.map(emp => ({
      id: Date.now() + emp.id, category: "Salary", bankId: Number(payrollBankId),
      amount: emp.netPay, company: emp.name, date: payrollDate,
      note: `ደሞዝ:${emp.basicSalary} ዲውቲ:${emp.duty} ጡረታ:-${emp.pension}`,
      status: "Pending",
    }));
    setDb(p => ({ ...p, expenseEntries: [...expenses, ...p.expenseEntries] }));
    showToast("success", `${emps.length} ሰራተኞች ክፍያ ተፈጽሟል።`);
  };

  const exportPayroll = () => {
    const emps = db.employees || [];
    if (!emps.length) { showToast("error", "ምንም ሰራተኛ የለም።"); return; }
    const ws = XLSXStyle.utils.aoa_to_sheet([
      ["Dr Hibist Pediatrics and Medical Center — Payroll"],
      [`Date: ${new Date().toLocaleString()}`],
      [],
      ["ተ.ቁ", "ስም", "ዲፓርትመንት", "ስልክ ቁጥር", "የባንክ አካውንት", "ደሞዝ (Birr)", "ዲውቲ (Birr)", "7% ጡረታ", "የሚከፈል (Birr)"],
      ...emps.map((e, i) => [
        i + 1, e.name, e.department, e.phone || "", e.bankAccount || "",
        e.basicSalary, e.duty || 0, e.pension || 0, e.netPay,
      ]),
      [],
      ["", "ጠቅላላ", "", "", "",
        emps.reduce((s, e) => s + e.basicSalary, 0),
        emps.reduce((s, e) => s + (e.duty || 0), 0),
        emps.reduce((s, e) => s + (e.pension || 0), 0),
        emps.reduce((s, e) => s + e.netPay, 0),
      ],
    ]);
    ws["!cols"] = [{ wch: 6 }, { wch: 22 }, { wch: 14 }, { wch: 14 }, { wch: 18 }, { wch: 14 }, { wch: 12 }, { wch: 12 }, { wch: 16 }];
    const wb = XLSXStyle.utils.book_new(); XLSXStyle.utils.book_append_sheet(wb, ws, "Payroll");
    const buf = XLSXStyle.write(wb, { bookType: "xlsx", type: "array" });
    const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a"); a.href = url; a.download = `payroll_${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    showToast("success", "Payroll exported to Excel.");
  };

  // ---- Balance Sheet handlers ----
  const saveBalanceSheet = () => {
    const n = (v) => Number(v) || 0;
    setDb(p => ({
      ...p,
      balanceSheet: {
        asOfDate: bsForm.asOfDate,
        cashAndBank: n(bsForm.cashAndBank),
        accountsReceivable: n(bsForm.accountsReceivable),
        inventory: n(bsForm.inventory),
        equipmentAndVehicles: n(bsForm.equipmentAndVehicles),
        accumulatedDepreciation: n(bsForm.accumulatedDepreciation),
        accountsPayable: n(bsForm.accountsPayable),
        taxPayable: n(bsForm.taxPayable),
        capital: n(bsForm.capital),
        retainedEarnings: n(bsForm.retainedEarnings),
      },
    }));
    setBsEditing(false);
    showToast("success", "Balance Sheet ተቀምጧል።");
  };

  const startEditBs = () => {
    const bs = db.balanceSheet || DEFAULT.balanceSheet;
    setBsForm({
      asOfDate: bs.asOfDate || new Date().toISOString().slice(0, 10),
      cashAndBank: bs.cashAndBank || "",
      accountsReceivable: bs.accountsReceivable || "",
      inventory: bs.inventory || "",
      equipmentAndVehicles: bs.equipmentAndVehicles || "",
      accumulatedDepreciation: bs.accumulatedDepreciation || "",
      accountsPayable: bs.accountsPayable || "",
      taxPayable: bs.taxPayable || "",
      capital: bs.capital || "",
      retainedEarnings: bs.retainedEarnings || "",
    });
    setBsEditing(true);
  };

  const exportBalanceSheet = () => {
    const bs = db.balanceSheet || DEFAULT.balanceSheet;
    const orgName = db.organizationLogoText || db.organizationName;
    const etYear = new Date().getFullYear() - 7;
    const n = (v) => Number(v) || 0;
    const fmt = (v) => Number(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    const totalCurrentAssets = n(bs.cashAndBank) + n(bs.accountsReceivable) + n(bs.inventory);
    const totalNonCurrentAssets = n(bs.equipmentAndVehicles) - n(bs.accumulatedDepreciation);
    const totalAssets = totalCurrentAssets + totalNonCurrentAssets;
    const totalCurrentLiab = n(bs.accountsPayable) + n(bs.taxPayable);
    const aData = buildStatementData(db.incomeEntries, db.expenseEntries);
    const netProfit = aData.netProfit;
    const totalEquity = n(bs.capital) + n(bs.retainedEarnings) + netProfit;
    const totalLiabEquity = totalCurrentLiab + totalEquity;

    // ── Style constants ──
    const DARK_BLUE = "1E3A5F";
    const LIGHT_BLUE = "DBEAFE";
    const YELLOW = "FFF9E6";
    const GREEN_BG = "DCFCE7";
    const DARK_BG = "0F172A";
    const WHITE = "FFFFFF";
    const TEXT_GREEN = "166534";
    const TEXT_RED = "991B1B";
    const GRAY_BG = "F8FAFC";

    const hdr = (v) => ({
      v, t: "s",
      s: { font: { bold: true, color: { rgb: WHITE }, sz: 11 }, fill: { fgColor: { rgb: DARK_BLUE } }, alignment: { horizontal: "left", vertical: "center", wrapText: true }, border: { bottom: { style: "medium", color: { rgb: WHITE } } } }
    });
    const secHdr = (v, col = "left") => ({
      v, t: "s",
      s: { font: { bold: true, color: { rgb: "334155" }, sz: 11 }, fill: { fgColor: { rgb: GRAY_BG } }, alignment: { horizontal: col, wrapText: true }, border: { top: { style: "thin", color: { rgb: "CBD5E1" } } } }
    });
    const data = (v, indent = false) => ({
      v, t: "s",
      s: { font: { sz: 11 }, fill: { fgColor: { rgb: indent ? YELLOW : WHITE } }, alignment: { horizontal: "left", indent: indent ? 1 : 0, wrapText: true } }
    });
    const amt = (v, color) => ({
      v: fmt(v), t: "s",
      s: { font: { bold: false, color: { rgb: color || "0F172A" }, sz: 11 }, fill: { fgColor: { rgb: YELLOW } }, alignment: { horizontal: "right" } }
    });
    const subTotal = (label, value, bgColor = LIGHT_BLUE) => [
      { v: label, t: "s", s: { font: { bold: true, sz: 11 }, fill: { fgColor: { rgb: bgColor } }, alignment: { horizontal: "left", wrapText: true }, border: { top: { style: "medium", color: { rgb: DARK_BLUE } } } } },
      { v: fmt(value), t: "s", s: { font: { bold: true, sz: 11 }, fill: { fgColor: { rgb: bgColor } }, alignment: { horizontal: "right" }, border: { top: { style: "medium", color: { rgb: DARK_BLUE } } } } },
    ];
    const grandTotal = (label, value) => [
      { v: label, t: "s", s: { font: { bold: true, color: { rgb: WHITE }, sz: 12 }, fill: { fgColor: { rgb: DARK_BG } }, alignment: { horizontal: "left" }, border: { top: { style: "medium", color: { rgb: "4ADE80" } } } } },
      { v: fmt(value), t: "s", s: { font: { bold: true, color: { rgb: "4ADE80" }, sz: 12 }, fill: { fgColor: { rgb: DARK_BG } }, alignment: { horizontal: "right" }, border: { top: { style: "medium", color: { rgb: "4ADE80" } } } } },
    ];
    const empty = () => [{ v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }];
    const titleCell = (v, sz = 14, bold = true) => ({ v, t: "s", s: { font: { bold, sz, color: { rgb: DARK_BLUE } }, alignment: { horizontal: "center" } } });

    const rows = [
      // Title rows (span all 5 cols via merge)
      [titleCell(`[${orgName}]`, 16, true), { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }],
      [titleCell("የሃብትና እዳ መግለጫ (Balance Sheet)", 18, true), { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }],
      [titleCell(`አ.አ. ሐምሌ 30/${etYear} ቀን — As of ${bs.asOfDate || new Date().toISOString().slice(0, 10)}`, 11, false), { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }],
      empty(),
      // Column headers
      [hdr("ሃብት (Assets)"), hdr("ብር (ETB)"), { v: "", t: "s", s: { fill: { fgColor: { rgb: "E2E8F0" } } } }, hdr("ዕዳ ካፒታል (Liabilities & Equity)"), hdr("ብር (ETB)")],
      // Current Assets | Current Liabilities
      [secHdr("የአሁኑ ጊዜ ሃብት (Current Assets)"), { v: "", t: "s", s: { fill: { fgColor: { rgb: GRAY_BG } } } }, { v: "", t: "s", s: {} }, secHdr("የአሁኑ ጊዜ ዕዳ (Current Liabilities)"), { v: "", t: "s", s: { fill: { fgColor: { rgb: GRAY_BG } } } }],
      [data("ጥሬ ናንያ ወይም ሃ/ጊዜ ገንዘብ (Cash & Bank)", true), amt(bs.cashAndBank || 0), { v: "", t: "s", s: {} }, data("ሌሎቻቸሞች ያልቀፈሉት (Accounts Payable)", true), amt(bs.accountsPayable || 0)],
      [data("ያልቀፈሉ ደረሰኝ ዕዳ (Accounts Receivable)", true), amt(bs.accountsReceivable || 0), { v: "", t: "s", s: {} }, data("ያልቀፈሉ ግብር/ታክስ (Tax Payable)", true), amt(bs.taxPayable || 0)],
      [data("የዕቃ ክምችት (Inventory)", true), amt(bs.inventory || 0), { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }],
      [...subTotal("ጠቅላላ የአሁኑ ጊዜ ሃብት", totalCurrentAssets), { v: "", t: "s", s: {} }, ...subTotal("ጠቅላላ የአሁኑ ጊዜ ዕዳ", totalCurrentLiab)],
      empty(),
      // Non-Current Assets | Owner's Equity
      [secHdr("ቋሚ ጊዜ ሃብት (Non-Current Assets)"), { v: "", t: "s", s: { fill: { fgColor: { rgb: GRAY_BG } } } }, { v: "", t: "s", s: {} }, secHdr("የባለቤቱ ካፒታል (Owner's Equity)"), { v: "", t: "s", s: { fill: { fgColor: { rgb: GRAY_BG } } } }],
      [data("መሳሪያና ተሸከርካሪ (Equipment & Vehicles)", true), amt(bs.equipmentAndVehicles || 0), { v: "", t: "s", s: {} }, data("የመጀመሪያ ሙሉ ካፒታል (Capital)", true), amt(bs.capital || 0)],
      [data("ልቀሰ: የተጠራቀመ እርጅና (Acc. Depreciation)", true), amt(bs.accumulatedDepreciation || 0, TEXT_RED), { v: "", t: "s", s: {} }, data("የተጠራቀሙ ትርፍ (Retained Earnings)", true), amt(bs.retainedEarnings || 0)],
      [{ v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, data("የዓመቱ የተጣራ ትርፍ (Net Profit)", true), { v: fmt(netProfit), t: "s", s: { font: { bold: true, color: { rgb: netProfit >= 0 ? TEXT_GREEN : TEXT_RED }, sz: 11 }, fill: { fgColor: { rgb: YELLOW } }, alignment: { horizontal: "right" } } }],
      [...subTotal("ጠቅላላ ቋሚ ጊዜ ሃብት", totalNonCurrentAssets), { v: "", t: "s", s: {} }, ...subTotal("ጠቅላላ የባለቤቱ ካፒታል", totalEquity, GREEN_BG)],
      empty(),
      [{ v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, ...subTotal("ጠቅላላ ዕዳና ካፒታል (Total Liab. & Equity)", totalLiabEquity)],
      empty(),
      [...grandTotal("ጠቅላላ ሃብት (Total Assets)", totalAssets), { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }, { v: "", t: "s", s: {} }],
    ];

    const ws = XLSXStyle.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 42 }, { wch: 18 }, { wch: 3 }, { wch: 42 }, { wch: 18 }];
    ws["!rows"] = [{ hpt: 22 }, { hpt: 26 }, { hpt: 18 }];
    // Merge title rows across all columns
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: 4 } },
    ];

    const wb = XLSXStyle.utils.book_new();
    XLSXStyle.utils.book_append_sheet(wb, ws, "Balance Sheet");
    const buf = XLSXStyle.write(wb, { bookType: "xlsx", type: "array" });
    const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = `balance_sheet_${(bs.asOfDate || new Date().toISOString().slice(0, 10))}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    showToast("success", "Balance Sheet exported to Excel.");
  };

  // ---- Leave handlers ----
  const submitLeaveRequest = (e) => {
    e.preventDefault();
    if (!leaveForm.empId || !leaveForm.days || !leaveForm.startDate) { showToast("error", "Fill all leave fields."); return; }
    const ltype = leaveForm.leaveType || "annual";
    // For annual leave, check remaining balance
    if (ltype === "annual") {
      const emp = (db.employees || []).find(em => em.id === Number(leaveForm.empId));
      if (emp) {
        const rem = (emp.totalLeaveDays || 0) - (emp.usedLeaveDays || 0);
        if (Number(leaveForm.days) > rem) { showToast("error", `ይህ ሰራተኛ ${rem} ቀን ብቻ ፍቃድ ቀርቷቸዋል።`); return; }
      }
    }
    const req = { id: Date.now(), empId: Number(leaveForm.empId), days: Number(leaveForm.days), startDate: leaveForm.startDate, reason: leaveForm.reason, leaveType: ltype, status: "Pending", requestDate: today() };
    setDb(p => ({ ...p, leaveRequests: [req, ...(p.leaveRequests || [])] }));
    setLeaveForm({ empId: "", days: "", startDate: today(), reason: "", leaveType: "annual" });
    showToast("success", ltype === "emergency" ? "ድንገተኛ ፍቃድ ጥያቄ ቀረበ።" : "ፍቃድ ጥያቄ ቀረበ።");
  };

  const approveLeave = (id) => {
    setDb(p => {
      const req = (p.leaveRequests || []).find(r => r.id === id);
      return {
        ...p,
        leaveRequests: p.leaveRequests.map(r => r.id !== id ? r : { ...r, status: "Approved", approvedDate: today() }),
        // Only deduct from annual balance for annual leave type
        employees: p.employees.map(emp => {
          if (!req || emp.id !== req.empId) return emp;
          if (req.leaveType === "emergency") return emp; // emergency doesn't consume annual days
          return { ...emp, usedLeaveDays: (emp.usedLeaveDays || 0) + req.days };
        }),
      };
    });
    showToast("success", "ፍቃድ ተቀበለ።");
  };

  const rejectLeave = (id) => {
    setDb(p => ({ ...p, leaveRequests: p.leaveRequests.map(r => r.id === id ? { ...r, status: "Rejected" } : r) }));
    showToast("success", "ፍቃድ ተከልክሏል።");
  };

  const deleteLeaveRequest = (id) => {
    if (!window.confirm("ይህን ፍቃድ ጥያቄ ይሰርዙ?")) return;
    setDb(p => ({ ...p, leaveRequests: (p.leaveRequests || []).filter(r => r.id !== id) }));
    showToast("success", "ፍቃድ ጥያቄ ተሰርዟል።");
  };

  const deleteResignation = (id) => {
    if (!window.confirm("ይህን ስራ ልቀቃ ጥያቄ ይሰርዙ?")) return;
    setDb(p => ({ ...p, resignations: (p.resignations || []).filter(r => r.id !== id) }));
    showToast("success", "ስራ ልቀቃ ጥያቄ ተሰርዟል።");
  };

  const deleteMessage = (id) => {
    setDb(p => ({ ...p, messages: (p.messages || []).filter(m => m.id !== id) }));
    showToast("success", "መልዕክት ተሰርዟል።");
  };

  // ---- Employee leave request (from portal) ----
  const empRequestLeave = (e) => {
    e.preventDefault();
    if (!empLoggedIn || !leaveForm.days || !leaveForm.startDate) { showToast("error", "Fill all fields."); return; }
    const ltype = leaveForm.leaveType || "annual";
    if (ltype === "annual") {
      const rem = (empLoggedIn.totalLeaveDays || 25) - (empLoggedIn.usedLeaveDays || 0);
      if (Number(leaveForm.days) > rem) { showToast("error", `${rem} ቀን ብቻ ቀርቷል። ለድንገተኛ ፍቃድ "Emergency" ይምረጡ።`); return; }
    }
    const req = { id: Date.now(), empId: empLoggedIn.id, days: Number(leaveForm.days), startDate: leaveForm.startDate, reason: leaveForm.reason, leaveType: ltype, status: "Pending", requestDate: today() };
    setDb(p => ({ ...p, leaveRequests: [req, ...(p.leaveRequests || [])] }));
    setLeaveForm({ empId: "", days: "", startDate: today(), reason: "", leaveType: "annual" });
    showToast("success", ltype === "emergency" ? "ድንገተኛ ፍቃድ ጥያቄ ቀርቧል። ቀጣይ ዉሳኔ ይጠብቁ።" : "ፍቃድ ጥያቄ ቀርቧል። ቀጣይ ዉሳኔ ይጠብቁ።");
  };

  // ---- Messaging ----
  const sendAdminMsg = (e) => {
    e.preventDefault();
    if (!msgForm.toId || !msgForm.body.trim()) { showToast("error", "Recipient and message required."); return; }
    const msg = { id: Date.now(), from: "admin", to: String(msgForm.toId), body: msgForm.body.trim(), date: today(), readByAdmin: true, readByEmp: false };
    setDb(p => ({ ...p, messages: [msg, ...(p.messages || [])] }));
    setMsgForm({ toId: "", body: "" }); showToast("success", "መልዕክት ተላከ።");
  };

  const sendEmpMsg = (e) => {
    e.preventDefault();
    if (!empMsgBody.trim() || !empLoggedIn) { showToast("error", "መልዕክት ባዶ ነው።"); return; }
    const msg = { id: Date.now(), from: String(empLoggedIn.id), to: "admin", body: empMsgBody.trim(), date: today(), readByAdmin: false, readByEmp: true };
    setDb(p => ({ ...p, messages: [msg, ...(p.messages || [])] }));
    setEmpMsgBody(""); showToast("success", "ለአድሚን ተላከ።");
  };

  const markAdminMsgsRead = () => {
    setDb(p => ({ ...p, messages: p.messages.map(m => m.to === "admin" ? { ...m, readByAdmin: true } : m) }));
  };

  const markEmpMsgsRead = (empId) => {
    setDb(p => ({ ...p, messages: p.messages.map(m => m.to === String(empId) ? { ...m, readByEmp: true } : m) }));
  };

  // ---- Employee portal login by phone + password ----
  // Manager department ብቻ ፓስዋርድ-free login ይጠቀማሉ
  const PASSWORD_FREE_DEPTS = ["Manager"];

  const handleEmpLogin = (e) => {
    e.preventDefault();
    const ph = empLoginForm.phone.trim();
    const pw = empLoginForm.password.trim();
    if (!ph) { showToast("error", "ስልክ ቁጥር ያስፈልጋል።"); return; }

    // ስልክ ቁጥር ብቻ ተጠቅሞ ሰራተኛ ፈልግ
    const empByPhone = (db.employees || []).find(em => em.phone === ph && em.hasPortalAccess);
    if (!empByPhone) { showToast("error", "ስልክ ቁጥር ትክክል አይደለም ወይም portal access የለም።"); return; }

    // ፓስዋርድ-free departments ከሆነ ፓስዋርድ አያስፈልግም
    if (PASSWORD_FREE_DEPTS.includes(empByPhone.department)) {
      setEmpLoggedIn(empByPhone); markEmpMsgsRead(empByPhone.id);
      showToast("success", `እንኳን ደህና መጡ ${empByPhone.name}!`);
      return;
    }

    // ሌሎች departments ፓስዋርድ ያስፈልጋቸዋል
    if (!pw) { showToast("error", "ፓስዋርድ ያስፈልጋል።"); return; }
    if (empByPhone.password !== pw) { showToast("error", "ፓስዋርድ ትክክል አይደለም።"); return; }
    setEmpLoggedIn(empByPhone); markEmpMsgsRead(empByPhone.id);
    showToast("success", `እንኳን ደህና መጡ ${empByPhone.name}!`);
  };

  // refresh empLoggedIn when db.employees changes
  useEffect(() => {
    if (!empLoggedIn) return;
    const updated = (db.employees || []).find(e => e.id === empLoggedIn.id);
    if (updated) setEmpLoggedIn(updated);
  }, [db.employees]);

  // Build income statement data for a given set of entries
  const buildStatementData = (incEntries, expEntries) => {
    const approvedExp = expEntries.filter(e => e.status === "Approved" || !e.status);
    const revenue = incEntries.reduce((s, e) => s + Number(e.amount || 0), 0);

    // COGS: Pharmacy + Lab categories treated as cost of goods
    const cogsCategories = ["Pharmacy", "Lab", "Medicine"];
    const cogs = approvedExp
      .filter(e => cogsCategories.includes(e.category))
      .reduce((s, e) => s + Number(e.amount || 0), 0);
    const grossProfit = revenue - cogs;

    // Build full expense map (non-COGS)
    const opExpMap = {};
    approvedExp
      .filter(e => !cogsCategories.includes(e.category))
      .forEach(e => { opExpMap[e.category] = (opExpMap[e.category] || 0) + Number(e.amount || 0); });

    // Fixed rows matching the image — always shown (0 if no data)
    const salaries = (opExpMap["Salary"] || 0);
    const rent = (opExpMap["Rent"] || 0);
    const utilities = (opExpMap["Electricity"] || 0) + (opExpMap["Water"] || 0) + (opExpMap["Waste"] || 0) + (opExpMap["Utilities"] || 0);
    const transport = (opExpMap["Transport"] || 0) + (opExpMap["Telecom"] || 0);
    const misc = (opExpMap["Misc Expenses"] || 0);

    // Everything else — dynamic rows (e.g. Salary from payroll with different keys)
    const fixedKeys = ["Salary", "Rent", "Electricity", "Water", "Waste", "Utilities", "Transport", "Telecom", "Misc Expenses"];
    const otherMap = {};
    Object.entries(opExpMap).forEach(([cat, amt]) => {
      if (!fixedKeys.includes(cat)) otherMap[cat] = amt;
    });
    const otherRows = Object.entries(otherMap)
      .map(([cat, amt]) => ({ cat, amt }))
      .sort((a, b) => b.amt - a.amt);

    const totalOpEx = salaries + rent + utilities + transport + misc +
      otherRows.reduce((s, x) => s + x.amt, 0);
    const netOperatingIncome = grossProfit - totalOpEx;
    // Ethiopian corporate income tax: 30% on positive profit
    const tax = netOperatingIncome > 0 ? Math.round(netOperatingIncome * 0.30) : 0;
    const netProfit = netOperatingIncome - tax;

    return {
      revenue, cogs, grossProfit,
      salaries, rent, utilities, transport, misc,
      otherRows,
      totalOpEx, netOperatingIncome, tax, netProfit,
    };
  };

  // ---- Excel export — Income Statement with full cell styling ----
  const exportIncomeStatement = (periodLabel, data) => {
    const orgName = db.organizationLogoText || db.organizationName || "Dr Hibist Pediatrics and Medical Center";
    const now = new Date();
    const etYear = now.getFullYear() - 7;
    const fmt = (n) => Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const { revenue, cogs, grossProfit, salaries, rent, utilities, transport, misc, otherRows, totalOpEx, netOperatingIncome, tax, netProfit } = data;

    // ── Style constants ──
    const DARK_BLUE = "1E3A5F";
    const YELLOW = "FFF9E6";
    const GREEN_BG = "F0FDF4";
    const DARK_BG = "0F172A";
    const GRAY_BG = "F8FAFC";
    const WHITE = "FFFFFF";
    const TEXT_GREEN = "166534";
    const TEXT_RED = "991B1B";

    // cell builders
    const c = (v, s) => ({ v: v ?? "", t: "s", s });
    const TITLE = (v, sz = 14) => c(v, { font: { bold: true, sz, color: { rgb: DARK_BLUE } }, alignment: { horizontal: "center" } });
    const SUBTITLE = (v) => c(v, { font: { italic: true, sz: 11, color: { rgb: "64748B" } }, alignment: { horizontal: "center" } });
    const HDR = (v) => c(v, { font: { bold: true, sz: 11, color: { rgb: WHITE } }, fill: { fgColor: { rgb: DARK_BLUE } }, alignment: { horizontal: v === "የሂሳብ መይብ (Account Description)" ? "left" : "center" }, border: { bottom: { style: "medium", color: { rgb: WHITE } } } });
    const SECTION = (v) => c(v, { font: { bold: true, sz: 11, color: { rgb: "334155" } }, fill: { fgColor: { rgb: GRAY_BG } }, alignment: { horizontal: "left" }, border: { top: { style: "thin", color: { rgb: "CBD5E1" } } } });
    const BLANK = (bg = WHITE) => c("", { fill: { fgColor: { rgb: bg } } });
    const DATA_L = (v) => c(v, { font: { sz: 11 }, fill: { fgColor: { rgb: YELLOW } }, alignment: { horizontal: "left", wrapText: true } });
    const DATA_R = (v) => c(fmt(v), { font: { sz: 11 }, fill: { fgColor: { rgb: YELLOW } }, alignment: { horizontal: "right" } });
    const NORMAL_R = (v) => c(fmt(v), { font: { sz: 11 }, alignment: { horizontal: "right" } });
    const BOLD_L = (v, bg = WHITE) => c(v, { font: { bold: true, sz: 11 }, fill: { fgColor: { rgb: bg } }, alignment: { horizontal: "left", wrapText: true }, border: { top: { style: "medium", color: { rgb: DARK_BLUE } } } });
    const BOLD_R = (v, color = "0F172A", bg = WHITE) => c(fmt(v), { font: { bold: true, sz: 11, color: { rgb: color } }, fill: { fgColor: { rgb: bg } }, alignment: { horizontal: "right" }, border: { top: { style: "medium", color: { rgb: DARK_BLUE } } } });
    const NET_PROFIT_L = (v) => c(v, { font: { bold: true, sz: 13, color: { rgb: WHITE } }, fill: { fgColor: { rgb: DARK_BG } }, alignment: { horizontal: "left" }, border: { top: { style: "thick", color: { rgb: "4ADE80" } } } });
    const NET_PROFIT_R = (v, pos) => c(fmt(v), { font: { bold: true, sz: 13, color: { rgb: pos ? "4ADE80" : TEXT_RED } }, fill: { fgColor: { rgb: DARK_BG } }, alignment: { horizontal: "right" }, border: { top: { style: "thick", color: { rgb: "4ADE80" } } } });
    const NET_INC_R = (v, pos) => c(fmt(v), { font: { bold: true, sz: 11, color: { rgb: pos ? TEXT_GREEN : TEXT_RED } }, fill: { fgColor: { rgb: pos ? GREEN_BG : "FEE2E2" } }, alignment: { horizontal: "right" }, border: { top: { style: "medium", color: { rgb: DARK_BLUE } } } });

    const rows = [
      // ── Title block ──
      [TITLE(`[${orgName}]`, 15), BLANK(), BLANK()],
      [TITLE("የገቢና ወጪ መግለጫ (Income Statement)", 17), BLANK(), BLANK()],
      [SUBTITLE(`ለ${etYear} አ.ም. / ${periodLabel}`), BLANK(), BLANK()],
      [BLANK(), BLANK(), BLANK()],

      // ── Column headers ──
      [HDR("የሂሳብ መይብ (Account Description)"), HDR("ነጉስ ድምር (ETB)"), HDR("ዋና ድምር (ETB)")],

      // ── Revenue ──
      [DATA_L("ከሽያጭ ወይም የአጠቃላዮት ገቢ (Revenue)"), BLANK(YELLOW), NORMAL_R(revenue)],
      [DATA_L("ልቅነስ: የሽያጭ ዕቃዎች ዋጋ (Cost of Goods Sold – COGS)"), DATA_R(cogs), BLANK(YELLOW)],
      [BOLD_L("አጠቃላይ ትርፍ (Gross Profit)", GREEN_BG), BLANK(GREEN_BG), BOLD_R(grossProfit, TEXT_GREEN, GREEN_BG)],
      [BLANK(), BLANK(), BLANK()],

      // ── Operating Expenses ──
      [SECTION("የሥራ ማስኬጃ ወጪዎች (Operating Expenses):"), BLANK(GRAY_BG), BLANK(GRAY_BG)],
      [DATA_L("  የሰራተኞች ደሞዝ (Salaries & Wages)"), DATA_R(salaries), BLANK(YELLOW)],
      [DATA_L("  የቢሮ ወይም የሱቅ ኪራይ (Rent Expense)"), DATA_R(rent), BLANK(YELLOW)],
      [DATA_L("  ውሃ፣ መብራትና ስልክ (Utilities)"), DATA_R(utilities), BLANK(YELLOW)],
      [DATA_L("  የትርንስፖርትና ሎጅስቲክስ (Transport)"), DATA_R(transport), BLANK(YELLOW)],
      [DATA_L("  ልዩ ልዩ ወጪዎች (Miscellaneous)"), DATA_R(misc), BLANK(YELLOW)],
      // dynamic other rows
      ...otherRows.map(({ cat, amt: a }) => [DATA_L(`  ${cat}`), DATA_R(a), BLANK(YELLOW)]),

      [BOLD_L("ጠቅላላ የሥራ ማስኬጃ ወጪ (Total Operating Expenses)"), BLANK(), BOLD_R(totalOpEx)],
      [BLANK(), BLANK(), BLANK()],

      // ── Bottom lines ──
      [BOLD_L("የተጣራ የንግድ ትርፍ (Net Operating Income)"), BLANK(), NET_INC_R(netOperatingIncome, netOperatingIncome >= 0)],
      [DATA_L("ልቅነስ: የገቢ ግብር / ታክስ (Income Tax Expense — 30%)"), BLANK(YELLOW), NORMAL_R(tax)],
      [NET_PROFIT_L("የዓመቱ የተጣራ ትርፍ (Net Profit for the Year)"), c("", { fill: { fgColor: { rgb: DARK_BG } } }), NET_PROFIT_R(netProfit, netProfit >= 0)],

      [BLANK(), BLANK(), BLANK()],
      [BLANK(), BLANK(), c(`Generated: ${now.toLocaleString()}`, { font: { italic: true, sz: 9, color: { rgb: "94A3B8" } }, alignment: { horizontal: "right" } })],
    ];

    const ws = XLSXStyle.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 54 }, { wch: 22 }, { wch: 22 }];
    ws["!rows"] = [{ hpt: 22 }, { hpt: 26 }, { hpt: 16 }];
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 2 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 2 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: 2 } },
    ];

    const wb = XLSXStyle.utils.book_new();
    XLSXStyle.utils.book_append_sheet(wb, ws, "Income Statement");
    const buf = XLSXStyle.write(wb, { bookType: "xlsx", type: "array" });
    const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = `income_statement_${periodLabel.replace(/\s/g, "_").toLowerCase()}.xlsx`;
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    showToast("success", "Income Statement exported to Excel.");
  };

  // ---- Leave: approve with adjusted days + reject with reason ----
  const approveLeaveAdjusted = (id) => {
    const adjustedDays = leaveApproveEdit[id];
    setDb(p => {
      const req = (p.leaveRequests || []).find(r => r.id === id);
      const days = adjustedDays ? Number(adjustedDays) : req?.days || 0;
      return {
        ...p,
        leaveRequests: p.leaveRequests.map(r => r.id !== id ? r : { ...r, status: "Approved", approvedDate: today(), approvedDays: days }),
        employees: p.employees.map(emp => {
          if (!req || emp.id !== req.empId) return emp;
          if (req.leaveType === "emergency") return emp;
          return { ...emp, usedLeaveDays: (emp.usedLeaveDays || 0) + days };
        }),
      };
    });
    setLeaveApproveEdit(p => { const n = { ...p }; delete n[id]; return n; });
    showToast("success", "ፍቃድ ተቀበለ።");
  };

  const rejectLeaveWithReason = (id) => {
    const reason = leaveRejectReason[id]?.trim() || "";
    if (!reason) { showToast("error", "የሚከለክሉበት ምክንያት ያስፈልጋል።"); return; }
    setDb(p => ({
      ...p,
      leaveRequests: p.leaveRequests.map(r =>
        r.id === id ? { ...r, status: "Rejected", rejectReason: reason } : r
      ),
    }));
    setLeaveRejectReason(p => { const n = { ...p }; delete n[id]; return n; });
    showToast("success", "ፍቃድ ተከልክሏል።");
  };

  // ---- Announcements ----
  const sendAnnouncement = (e) => {
    e.preventDefault();
    if (!announcementForm.title.trim() || !announcementForm.body.trim()) { showToast("error", "ርዕስ እና ዝርዝር ያስፈልጋሉ።"); return; }
    const ann = { id: Date.now(), title: announcementForm.title.trim(), body: announcementForm.body.trim(), date: today(), readBy: [] };
    setDb(p => ({ ...p, announcements: [ann, ...(p.announcements || [])] }));
    setAnnouncementForm({ title: "", body: "" });
    showToast("success", "ማስታወቂያ ለሁሉም ሰራተኞች ተላከ።");
  };

  const markAnnouncementRead = (annId, empId) => {
    setDb(p => ({
      ...p,
      announcements: (p.announcements || []).map(a =>
        a.id === annId && !(a.readBy || []).includes(empId)
          ? { ...a, readBy: [...(a.readBy || []), empId] }
          : a
      ),
    }));
  };

  const deleteAnnouncement = (id) => {
    setDb(p => ({ ...p, announcements: (p.announcements || []).filter(a => a.id !== id) }));
  };

  // ---- Resignations ----
  const submitResignation = (e) => {
    e.preventDefault();
    if (!empLoggedIn || !resignForm.reason.trim() || !resignForm.lastDay) { showToast("error", "ምክንያት እና የመጨረሻ ቀን ያስፈልጋሉ።"); return; }
    // check 30 days
    const today30 = new Date(); today30.setDate(today30.getDate() + 30);
    const last = new Date(resignForm.lastDay);
    if (last < today30) { showToast("error", "የስራ ልቀቃ ቀን ቢያንስ ከዛሬ 30 ቀን በኋላ መሆን አለበት።"); return; }
    const res = { id: Date.now(), empId: empLoggedIn.id, empName: empLoggedIn.name, department: empLoggedIn.department, reason: resignForm.reason.trim(), lastDay: resignForm.lastDay, status: "Pending", submittedDate: today() };
    setDb(p => ({ ...p, resignations: [res, ...(p.resignations || [])] }));
    setResignForm({ reason: "", lastDay: "" });
    showToast("success", "ስራ ልቀቅ ማሳወቂያ ለአድሚን ተላከ። ቀጣይ ዉሳኔ ይጠብቁ።");
  };

  const acknowledgeResignation = (id) => {
    setDb(p => ({ ...p, resignations: (p.resignations || []).map(r => r.id === id ? { ...r, status: "Acknowledged" } : r) }));
    showToast("success", "ስራ ልቀቃ ተቀበለ።");
  };

  // ---- Chat thread message ----
  const sendThreadReply = (e) => {
    e.preventDefault();
    if (!threadReply.trim() || !msgThreadEmpId) return;
    const msg = { id: Date.now(), from: "admin", to: String(msgThreadEmpId), body: threadReply.trim(), date: today(), readByAdmin: true, readByEmp: false };
    setDb(p => ({ ...p, messages: [msg, ...(p.messages || [])] }));
    setThreadReply("");
  };

  // ---- Employee portal: change password ----
  const empChangePassword = (e) => {
    e.preventDefault();
    if (!empLoggedIn) return;
    if (empChangePwForm.current !== empLoggedIn.password) { showToast("error", "የአሁኑ ፓስዋርድ ትክክል አይደለም።"); return; }
    if (!/^\d{4,6}$/.test(empChangePwForm.newPw)) { showToast("error", "አዲስ ፓስዋርድ 4-6 ቁጥር ብቻ ይሁን።"); return; }
    if (empChangePwForm.newPw !== empChangePwForm.confirm) { showToast("error", "ፓስዋርዶቹ አይዛመዱም።"); return; }
    setDb(p => ({ ...p, employees: p.employees.map(em => em.id === empLoggedIn.id ? { ...em, password: empChangePwForm.newPw } : em) }));
    setEmpChangePwForm({ current: "", newPw: "", confirm: "" });
    showToast("success", "ፓስዋርድ ተቀይሯል።");
  };

  // ---- Employee portal: upload profile photo ----
  const empUploadPhoto = (ev) => {
    const file = ev.target.files?.[0]; if (!file) return;
    const r = new FileReader();
    r.onload = () => {
      setDb(p => ({ ...p, employees: p.employees.map(em => em.id === empLoggedIn.id ? { ...em, photo: String(r.result) } : em) }));
      showToast("success", "ፎቶ ተቀምጧል።");
    };
    r.readAsDataURL(file);
  };

  // ---- Employee peer messaging ----
  const sendEmpPeerMsg = (e) => {
    e.preventDefault();
    if (!empLoggedIn || !empPeerMsgForm.toId || !empPeerMsgForm.body.trim()) { showToast("error", "ተቀባይ እና መልዕክት ያስፈልጋሉ።"); return; }
    const msg = { id: Date.now(), from: String(empLoggedIn.id), to: String(empPeerMsgForm.toId), body: empPeerMsgForm.body.trim(), date: today(), readByAdmin: false, readByEmp: false, isPeer: true };
    setDb(p => ({ ...p, messages: [msg, ...(p.messages || [])] }));
    setEmpPeerMsgForm({ toId: "", body: "" });
    showToast("success", "መልዕክት ተላከ።");
  };

  // ---- Expense receipt print ----
  const printExpenseReceipt = (expenseId) => {
    const exp = (db.expenseEntries || []).find(e => e.id === expenseId);
    if (!exp) return;
    const bank = (db.bankAccounts || []).find(b => b.id === Number(exp.bankId));
    const logo = db.organizationLogo || localStorage.getItem(LOGO_KEY) || "";
    const orgName = db.organizationLogoText || db.organizationName || "Dr Hibist Pediatrics and Medical Center";

    const logoHtml = logo
      ? `<img src="${logo}" alt="logo" style="height:72px;width:auto;object-fit:contain;border-radius:10px;" />`
      : `<div style="font-size:26px;font-weight:900;color:#1e3a5f;letter-spacing:-0.5px;">Dr Hibist</div>`;

    // Watermark — ሙሉ ሽፋን
    const watermarkHtml = logo
      ? `<div style="position:fixed;inset:0;display:flex;align-items:center;justify-content:center;pointer-events:none;z-index:0;overflow:hidden;">
           <img src="${logo}" alt="" style="width:80%;max-width:480px;height:auto;opacity:0.06;transform:rotate(-20deg);" />
         </div>`
      : `<div style="position:fixed;inset:0;display:flex;align-items:center;justify-content:center;pointer-events:none;z-index:0;overflow:hidden;">
           <span style="font-size:110px;font-weight:900;color:#1e3a5f;opacity:0.05;transform:rotate(-20deg);white-space:nowrap;">Dr Hibist</span>
         </div>`;

    const receiptHtml = `
      <html><head><title>Expense Receipt</title>
      <style>
        *{box-sizing:border-box;margin:0;padding:0;}
        body{font-family:'Segoe UI',Arial,sans-serif;padding:40px;max-width:560px;margin:0 auto;position:relative;background:#fff;}
        .header{display:flex;align-items:center;justify-content:space-between;margin-bottom:22px;padding-bottom:16px;border-bottom:3px solid #1e3a5f;}
        .header-left{display:flex;align-items:center;gap:14px;}
        .org-name{font-size:16px;font-weight:800;color:#1e3a5f;line-height:1.3;}
        .org-sub{font-size:11px;color:#64748b;margin-top:2px;}
        .receipt-badge{background:#1e3a5f;color:#fff;font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;letter-spacing:0.06em;text-transform:uppercase;}
        .receipt-no{font-size:26px;font-weight:900;color:#1e3a5f;text-align:right;margin-top:4px;}
        .section{position:relative;z-index:1;}
        .row{display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px dashed #e2e8f0;}
        .row:last-child{border-bottom:none;}
        .label{color:#64748b;font-size:13px;}
        .val{font-weight:700;font-size:13px;text-align:right;color:#1e293b;}
        .total-box{margin-top:20px;background:linear-gradient(135deg,#1e3a5f,#2563eb);border-radius:14px;padding:18px 22px;display:flex;justify-content:space-between;align-items:center;position:relative;z-index:1;}
        .total-label{color:rgba(255,255,255,0.85);font-size:14px;font-weight:600;}
        .total-val{color:#fff;font-size:28px;font-weight:900;letter-spacing:-0.5px;}
        .birr{font-size:14px;font-weight:600;opacity:0.85;margin-left:4px;}
        .status-pill{display:inline-block;padding:4px 12px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:0.03em;}
        .approved{background:#dcfce7;color:#15803d;}
        .pending{background:#fef9c3;color:#a16207;}
        .rejected{background:#fee2e2;color:#b91c1c;}
        .footer{text-align:center;margin-top:24px;font-size:10px;color:#94a3b8;position:relative;z-index:1;border-top:1px solid #f1f5f9;padding-top:14px;line-height:1.7;}
        @media print{body{padding:24px;} @page{margin:10mm;}}
      </style></head><body>
      ${watermarkHtml}
      <div class="header">
        <div class="header-left">
          ${logoHtml}
          <div>
            <div class="org-name">${orgName}</div>
            <div class="org-sub">Official Expense Receipt</div>
          </div>
        </div>
        <div style="text-align:right;">
          <div class="receipt-badge">RECEIPT</div>
          <div class="receipt-no">#${exp.id.toString().slice(-6)}</div>
        </div>
      </div>
      <div class="section">
        <div class="row"><span class="label">ቀን / Date</span><span class="val">${exp.date}</span></div>
        <div class="row"><span class="label">Category</span><span class="val">${exp.category}</span></div>
        <div class="row"><span class="label">Bank Account</span><span class="val">${bank?.name || "—"} (${bank?.accountNumber || "—"})</span></div>
        <div class="row"><span class="label">Note / ማስታወሻ</span><span class="val">${exp.note || "—"}</span></div>
        <div class="row">
          <span class="label">Status</span>
          <span class="val">
            <span class="status-pill ${exp.status === "Approved" ? "approved" : exp.status === "Rejected" ? "rejected" : "pending"}">${exp.status || "Approved"}</span>
          </span>
        </div>
      </div>
      <div class="total-box">
        <span class="total-label">Total Amount</span>
        <span class="total-val">${Number(exp.amount).toLocaleString()}<span class="birr">Birr</span></span>
      </div>
      <div class="footer">
        Generated: ${new Date().toLocaleString()}<br/>
        ${orgName} — Official Expense Receipt
      </div>
      </body></html>
    `;
    const w = window.open("", "_blank", "width=660,height=820");
    if (!w) { showToast("error", "Popup blocked — please allow popups."); return; }
    w.document.write(receiptHtml);
    w.document.close();
    setTimeout(() => w.print(), 400);
  };

  const handleLogoUpload = (ev) => {
    const file = ev.target.files?.[0]; if (!file) return;
    const r = new FileReader();
    r.onload = () => {
      const logoData = String(r.result);
      localStorage.setItem(LOGO_KEY, logoData); // persist separately
      setDb(p => ({ ...p, organizationLogo: logoData }));
      showToast("success", "Logo uploaded.");
    };
    r.readAsDataURL(file);
  };

  // ========== AUTH SCREEN ==========
  if (!isAdmin) {
    // Show employee portal if user switches to it
    if (showEmpPortal) {
      if (empLoggedIn) {
        const emp = empLoggedIn;
        const remaining = (emp.totalLeaveDays || 25) - (emp.usedLeaveDays || 0);
        const myMsgs = (db.messages || []).filter(m => m.from === "admin" && m.to === String(emp.id));

        // Countdown: days left until returning from approved leave
        const approvedLeave = (db.leaveRequests || []).find(r => r.empId === emp.id && r.status === "Approved" && r.startDate);
        let daysUntilReturn = null;
        if (approvedLeave) {
          const start = new Date(approvedLeave.startDate);
          const approvedDays = approvedLeave.approvedDays || approvedLeave.days || 0;
          const returnDate = new Date(start);
          returnDate.setDate(returnDate.getDate() + Number(approvedDays));
          const now = new Date();
          now.setHours(0, 0, 0, 0);
          returnDate.setHours(0, 0, 0, 0);
          const diff = Math.ceil((returnDate - now) / (1000 * 60 * 60 * 24));
          if (diff >= 0) daysUntilReturn = { days: diff, returnDate: returnDate.toISOString().slice(0, 10) };
        }

        // Peer messages involving this employee
        const peerMsgs = (db.messages || []).filter(m =>
          m.isPeer && (m.from === String(emp.id) || m.to === String(emp.id))
        );
        const unreadPeer = peerMsgs.filter(m => m.to === String(emp.id) && !m.readByEmp).length;

        return (
          <div className={`app-shell ${dark ? "dark-mode" : ""}`}>
            <div className="emp-portal-screen">
              <div className="emp-portal-card" style={{ maxWidth: 900 }}>
                {/* Header */}
                <div className="emp-portal-header">
                  <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
                    {emp.photo
                      ? <img src={emp.photo} alt="profile" style={{ width: 60, height: 60, borderRadius: "50%", objectFit: "cover", border: "3px solid #2563eb" }} />
                      : <div style={{ width: 60, height: 60, borderRadius: "50%", background: "#dbeafe", display: "grid", placeItems: "center", fontSize: 24, border: "3px solid #2563eb" }}>👤</div>
                    }
                    <div>
                      <h2 style={{ margin: "0 0 4px" }}>{emp.name}</h2>
                      <span className="dept-badge">{emp.department}</span>
                      {daysUntilReturn !== null && (
                        <div style={{ marginTop: 6, fontSize: 13, background: daysUntilReturn.days === 0 ? "#dcfce7" : "#fef3c7", color: daysUntilReturn.days === 0 ? "#166534" : "#92400e", padding: "3px 10px", borderRadius: 999, display: "inline-block", marginLeft: 8, fontWeight: 700 }}>
                          {daysUntilReturn.days === 0 ? "🏁 ዛሬ ወደ ስራ ይመለሳሉ!" : `🏖️ ወደ ስራ ${daysUntilReturn.days} ቀን ቀርቷል (${daysUntilReturn.returnDate})`}
                        </div>
                      )}
                    </div>
                  </div>
                  <button className="secondary-btn" onClick={() => { setEmpLoggedIn(null); setShowEmpPortal(false); }}>Logout</button>
                </div>
                {toast && <div className={`toast ${toast.type}`}>{toast.message}</div>}

                {/* Portal Tabs */}
                <div className="hrm-tabs" style={{ marginBottom: 20, flexWrap: "wrap" }}>
                  {[
                    ["home", "🏠 መነሻ"],
                    ["leave", "🗓️ ፍቃድ"],
                    ["messages", `💬 መልዕክቶች${(unreadEmp + unreadPeer) > 0 ? ` (${unreadEmp + unreadPeer})` : ""}`],
                    ["announcements", `📢 ማስታወቂያ${unreadAnnouncements > 0 ? ` (${unreadAnnouncements})` : ""}`],
                    ["resign", "📝 ስራ ልቀቃ"],
                    ["profile", "⚙️ መቼቶች"],
                    ...(emp.department === "Finance" ? [["finance", "💼 Finance"]] : []),
                  ].map(([k, l]) => (
                    <button key={k} className={empPortalTab === k ? "hrm-tab active" : "hrm-tab"} onClick={() => setEmpPortalTab(k)}>{l}</button>
                  ))}
                </div>

                {/* ===== HOME TAB ===== */}
                {empPortalTab === "home" && (
                  <div>
                    <div className="emp-info-grid">
                      <div className="emp-info-card">
                        <span className="emp-info-label">ደሞዝ</span>
                        <strong className="emp-info-val">{emp.basicSalary.toLocaleString()} Birr</strong>
                      </div>
                      <div className="emp-info-card">
                        <span className="emp-info-label">ዲውቲ</span>
                        <strong className="emp-info-val">{emp.duty.toLocaleString()} Birr</strong>
                      </div>
                      <div className="emp-info-card">
                        <span className="emp-info-label">7% ጡረታ</span>
                        <strong className="emp-info-val" style={{ color: "#ef4444" }}>-{emp.pension.toFixed(2)} Birr</strong>
                      </div>
                      <div className="emp-info-card">
                        <span className="emp-info-label">የሚከፈል</span>
                        <strong className="emp-info-val" style={{ color: "#16a34a" }}>{emp.netPay.toFixed(2)} Birr</strong>
                      </div>
                      <div className="emp-info-card">
                        <span className="emp-info-label">ጠቅላላ ፍቃድ</span>
                        <strong className="emp-info-val">{emp.totalLeaveDays || 25} ቀን</strong>
                      </div>
                      <div className="emp-info-card">
                        <span className="emp-info-label">የቀረ ፍቃድ</span>
                        <strong className="emp-info-val" style={{ color: remaining < 3 ? "#ef4444" : "#16a34a" }}>{remaining} ቀን{remaining < 3 ? " ⚠️" : ""}</strong>
                      </div>
                    </div>

                    {/* Doctor weekly summary — Doctor dept ብቻ */}
                    {emp.department === "Doctor" && (() => {
                      const history = getDoctorWeeklyHistoryForEmp(emp.id);
                      const approvedHistory = history.filter(h => h.status === "Approved");
                      const totalAllTime = approvedHistory.reduce((s, h) => s + Number(h.total || 0), 0);
                      const totalCutAllTime = approvedHistory.reduce((s, h) => s + Number(h.doctorCut || 0), 0);
                      const lastEntry = approvedHistory[0];
                      const pendingCount = history.filter(h => h.status === "Pending").length;
                      return (
                        <div style={{ marginTop: 16 }}>
                          <div style={{ fontSize: 13, fontWeight: 700, color: "#475569", marginBottom: 10 }}>🩺 የሳምንታዊ ክፍያ ማጠቃለያ</div>

                          {/* Pending notice */}
                          {pendingCount > 0 && (
                            <div style={{ marginBottom: 10, padding: "10px 14px", background: "#fef3c7", borderRadius: 10, border: "1px solid #fde68a", fontSize: 13, color: "#92400e", fontWeight: 700 }}>
                              ⏳ {pendingCount} ሳምንታዊ ክፍያ Admin approval ይጠብቃል...
                            </div>
                          )}

                          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
                            <div style={{ padding: "18px 20px", borderRadius: 14, background: "#f0f9ff", border: "1.5px solid #bae6fd", textAlign: "center" }}>
                              <div style={{ fontSize: 12, color: "#0369a1", fontWeight: 700, marginBottom: 6 }}>📊 ጠቅላላ ሳምንታዊ ገቢ</div>
                              <div style={{ fontSize: 28, fontWeight: 800, color: "#0f172a" }}>
                                {lastEntry ? Number(lastEntry.total).toLocaleString() : "—"}
                              </div>
                              <div style={{ fontSize: 11, color: "#64748b", marginTop: 3 }}>
                                {lastEntry ? `ቀን: ${lastEntry.date}` : "ምንም Approved ክፍያ የለም"}
                              </div>
                              {approvedHistory.length > 1 && (
                                <div style={{ fontSize: 11, color: "#0369a1", marginTop: 4 }}>
                                  ጠቅላላ ({approvedHistory.length} ሳምንት): {totalAllTime.toLocaleString()} Birr
                                </div>
                              )}
                            </div>
                            <div style={{ padding: "18px 20px", borderRadius: 14, background: "#0f172a", border: "2px solid #4ade80", textAlign: "center" }}>
                              <div style={{ fontSize: 12, color: "#4ade80", fontWeight: 700, marginBottom: 6 }}>💰 {lastEntry ? `${Number(lastEntry.percentage || 15)}%` : "15%"} — የእርስዎ ክፍያ</div>
                              <div style={{ fontSize: 28, fontWeight: 800, color: "#4ade80" }}>
                                {lastEntry ? Number(lastEntry.doctorCut).toLocaleString() : "—"}
                              </div>
                              <div style={{ fontSize: 11, color: "#86efac", marginTop: 3 }}>
                                {lastEntry ? "Birr" : "Admin approval ይጠብቁ"}
                              </div>
                              {approvedHistory.length > 1 && (
                                <div style={{ fontSize: 11, color: "#86efac", marginTop: 4 }}>
                                  ጠቅላላ: {totalCutAllTime.toLocaleString()} Birr
                                </div>
                              )}
                            </div>
                          </div>

                          {approvedHistory.length > 0 && (
                            <div style={{ marginTop: 12 }}>
                              <table className="report-table" style={{ fontSize: 13 }}>
                                <thead>
                                  <tr>
                                    <th>ቀን</th>
                                    <th style={{ textAlign: "right" }}>ጠቅላላ ገቢ</th>
                                    <th style={{ textAlign: "right", color: "#16a34a" }}>ክፍያ መቶኛ</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {approvedHistory.slice(0, 5).map(h => (
                                    <tr key={h.id}>
                                      <td>{h.date}</td>
                                      <td style={{ textAlign: "right", fontWeight: 700 }}>{Number(h.total).toLocaleString()} Birr</td>
                                      <td style={{ textAlign: "right", color: "#16a34a", fontWeight: 800 }}>{Number(h.doctorCut).toLocaleString()} Birr</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          )}
                        </div>
                      );
                    })()}

                    {daysUntilReturn !== null && (
                      <div style={{ marginTop: 12, padding: "14px 18px", borderRadius: 12, background: daysUntilReturn.days === 0 ? "#dcfce7" : "#fef3c7", border: `1px solid ${daysUntilReturn.days === 0 ? "#bbf7d0" : "#fde68a"}`, display: "flex", alignItems: "center", gap: 10 }}>
                        <span style={{ fontSize: 28 }}>{daysUntilReturn.days === 0 ? "🏁" : "🏖️"}</span>
                        <div>
                          <div style={{ fontWeight: 700, fontSize: 15, color: daysUntilReturn.days === 0 ? "#166534" : "#92400e" }}>
                            {daysUntilReturn.days === 0 ? "ዛሬ ወደ ስራ ይመለሳሉ!" : `ወደ ስራ ለመመለስ ${daysUntilReturn.days} ቀን ቀርቷል`}
                          </div>
                          <div style={{ fontSize: 13, color: "#64748b" }}>የመመለሻ ቀን: {daysUntilReturn.returnDate}</div>
                        </div>
                      </div>
                    )}
                  </div>
                )}

                {/* ===== LEAVE TAB ===== */}
                {empPortalTab === "leave" && (
                  <div className="emp-section" style={{ marginTop: 0, borderTop: "none", paddingTop: 0 }}>
                    <h3>ፍቃድ ጠይቅ</h3>
                    <form className="form-grid" onSubmit={empRequestLeave}>
                      <select value={leaveForm.leaveType} onChange={e => setLeaveForm(p => ({ ...p, leaveType: e.target.value }))}>
                        <option value="annual">🗓️ ዓመታዊ ፍቃድ (Annual)</option>
                        <option value="emergency">🚨 ድንገተኛ ፍቃድ (Emergency)</option>
                      </select>
                      <input type="number" placeholder="ቀናት" value={leaveForm.days} onChange={e => setLeaveForm(p => ({ ...p, days: e.target.value }))} required />
                      <input type="date" value={leaveForm.startDate} onChange={e => setLeaveForm(p => ({ ...p, startDate: e.target.value }))} />
                      <input placeholder="ምክንያት" value={leaveForm.reason} onChange={e => setLeaveForm(p => ({ ...p, reason: e.target.value }))} style={{ gridColumn: "1/-1" }} />
                      {leaveForm.leaveType === "annual" && (
                        <p style={{ gridColumn: "1/-1", margin: 0, fontSize: 13, color: remaining < 5 ? "#ef4444" : "#16a34a", fontWeight: 700 }}>
                          የቀረ ዓመታዊ ፍቃድ: {remaining} / {emp.totalLeaveDays || 25} ቀን
                        </p>
                      )}
                      {leaveForm.leaveType === "emergency" && (
                        <p style={{ gridColumn: "1/-1", margin: 0, fontSize: 13, color: "#f59e0b", fontWeight: 700 }}>
                          ⚠️ ድንገተኛ ፍቃድ ከዓመታዊ ፍቃድ ቀናት አይቆረጥም።
                        </p>
                      )}
                      <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>ፍቃድ ጠይቅ</button>
                    </form>
                    {/* my leave history */}
                    {(db.leaveRequests || []).filter(r => r.empId === emp.id).length > 0 && (
                      <table className="report-table" style={{ marginTop: 10 }}>
                        <thead><tr><th>ቀናት</th><th>ጀምሮ</th><th>ምክንያት</th><th>ሁኔታ</th><th>ማስታወሻ</th></tr></thead>
                        <tbody>
                          {(db.leaveRequests || []).filter(r => r.empId === emp.id).map(r => (
                            <tr key={r.id}>
                              <td>{r.approvedDays || r.days}</td><td>{r.startDate}</td><td>{r.reason}</td>
                              <td><span className={`status-badge ${r.status === "Approved" ? "approved" : r.status === "Rejected" ? "rejected" : "pending"}`}>{r.status}</span></td>
                              <td style={{ fontSize: 12, color: "#ef4444" }}>{r.rejectReason || ""}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                )}

                {/* ===== MESSAGES TAB ===== */}
                {empPortalTab === "messages" && (
                  <div className="emp-section" style={{ marginTop: 0, borderTop: "none", paddingTop: 0 }}>
                    <h3>💬 መልዕክቶች</h3>

                    {/* Admin → Employee messages */}
                    <div style={{ marginBottom: 20 }}>
                      <h4 style={{ fontSize: 14, margin: "0 0 8px" }}>📥 ከ Admin የተላኩ</h4>
                      {(db.messages || []).filter(m => m.from === "admin" && m.to === String(emp.id)).length === 0
                        ? <p style={{ color: "#94a3b8", fontSize: 13 }}>ምንም መልዕክት የለም።</p>
                        : <div className="msg-list">
                          {(db.messages || []).filter(m => m.from === "admin" && m.to === String(emp.id)).map(m => (
                            <div key={m.id} className="msg-bubble admin-bubble">
                              <div className="msg-from">Admin · {m.date}</div>
                              <p>{m.body}</p>
                            </div>
                          ))}
                        </div>
                      }
                    </div>

                    {/* Employee → Admin reply */}
                    <div style={{ marginBottom: 20 }}>
                      <h4 style={{ fontSize: 14, margin: "0 0 8px" }}>📤 ለ Admin መልስ ላክ</h4>
                      <form onSubmit={sendEmpMsg}>
                        <textarea
                          placeholder="መልዕክትዎን ይጻፉ..."
                          value={empMsgBody}
                          onChange={e => setEmpMsgBody(e.target.value)}
                          style={{ width: "100%", minHeight: 80, padding: "8px 12px", borderRadius: 8, border: "1px solid #cbd5e1", fontSize: 14, resize: "vertical" }}
                          required
                        />
                        <button className="btn" type="submit" style={{ marginTop: 8 }}>ላክ →</button>
                      </form>
                      {/* Sent messages */}
                      {(db.messages || []).filter(m => m.from === String(emp.id) && m.to === "admin").length > 0 && (
                        <div className="msg-list" style={{ marginTop: 12 }}>
                          <p style={{ fontSize: 12, color: "#94a3b8", margin: "0 0 6px" }}>የላክሃቸው:</p>
                          {(db.messages || []).filter(m => m.from === String(emp.id) && m.to === "admin").map(m => (
                            <div key={m.id} className="msg-bubble emp-bubble">
                              <div className="msg-from">እኔ · {m.date}</div>
                              <p>{m.body}</p>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* Peer messages */}
                    <div>
                      <h4 style={{ fontSize: 14, margin: "0 0 8px" }}>👥 ለሌላ ሰራተኛ ላክ</h4>
                      <form onSubmit={sendEmpPeerMsg}>
                        <div className="form-grid" style={{ marginBottom: 8 }}>
                          <select value={empPeerMsgForm.toId} onChange={e => setEmpPeerMsgForm(p => ({ ...p, toId: e.target.value }))} required style={{ gridColumn: "1/-1" }}>
                            <option value="">ሰራተኛ ምረጥ</option>
                            {(db.employees || []).filter(e2 => e2.id !== emp.id).map(e2 => (
                              <option key={e2.id} value={e2.id}>{e2.name} — {e2.department}</option>
                            ))}
                          </select>
                          <textarea placeholder="መልዕክት..." value={empPeerMsgForm.body} onChange={e => setEmpPeerMsgForm(p => ({ ...p, body: e.target.value }))}
                            style={{ gridColumn: "1/-1", minHeight: 70, padding: "8px", borderRadius: 8, border: "1px solid #cbd5e1", fontSize: 13 }} required />
                        </div>
                        <button className="btn" type="submit" style={{ marginTop: 0 }}>ላክ →</button>
                      </form>
                      {/* Peer messages received */}
                      {(db.messages || []).filter(m => m.isPeer && m.to === String(emp.id)).length > 0 && (
                        <div className="msg-list" style={{ marginTop: 12 }}>
                          <p style={{ fontSize: 12, color: "#94a3b8", margin: "0 0 6px" }}>የተቀበሏቸው:</p>
                          {(db.messages || []).filter(m => m.isPeer && m.to === String(emp.id)).map(m => {
                            const sender = (db.employees || []).find(e2 => String(e2.id) === String(m.from));
                            return (
                              <div key={m.id} className="msg-bubble emp-bubble">
                                <div className="msg-from">{sender?.name || "ሰራተኛ"} · {m.date}</div>
                                <p>{m.body}</p>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {/* ===== ANNOUNCEMENTS TAB ===== */}
                {empPortalTab === "announcements" && (
                  <div className="emp-section" style={{ marginTop: 0, borderTop: "none", paddingTop: 0 }}>
                    <h3>📢 ማስታወቂያዎች</h3>
                    {(db.announcements || []).length === 0
                      ? <p style={{ color: "#94a3b8", fontSize: 13 }}>ምንም ማስታወቂያ የለም።</p>
                      : <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
                        {(db.announcements || []).map(a => {
                          const isRead = (a.readBy || []).includes(emp.id);
                          return (
                            <div key={a.id}
                              className="msg-bubble admin-bubble"
                              style={{ borderLeft: `4px solid ${isRead ? "#94a3b8" : "#2563eb"}`, cursor: "pointer" }}
                              onClick={() => markAnnouncementRead(a.id, emp.id)}
                            >
                              <div className="msg-from" style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                <span>📢 {a.title} · {a.date}</span>
                                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                  {isRead
                                    ? <span style={{ color: "#94a3b8", fontSize: 11 }}>✓ አንብቤዋለሁ</span>
                                    : <span style={{ color: "#2563eb", fontSize: 11, fontWeight: 700 }}>● አዲስ</span>
                                  }
                                  {emp.department === "Finance" && (
                                    <button
                                      className="btn-danger btn-sm"
                                      style={{ padding: "2px 8px", fontSize: 11 }}
                                      onClick={e => { e.stopPropagation(); deleteAnnouncement(a.id); }}
                                    >🗑️</button>
                                  )}
                                </div>
                              </div>
                              <p style={{ margin: "8px 0 0", fontSize: 14 }}>{a.body}</p>
                            </div>
                          );
                        })}
                      </div>
                    }
                  </div>
                )}

                {/* ===== RESIGN TAB ===== */}
                {empPortalTab === "resign" && (
                  <div className="emp-section" style={{ marginTop: 0, borderTop: "none", paddingTop: 0 }}>
                    <h3>📝 ስራ መልቀቃ ማሳወቂያ</h3>
                    <p style={{ fontSize: 13, color: "#64748b", margin: "0 0 14px" }}>
                      ⚠️ የስራ ልቀቃ ቀን ቢያንስ ከዛሬ 30 ቀን በኋላ መሆን አለበት።
                    </p>
                    {/* Check if already has pending/acknowledged resignation */}
                    {(db.resignations || []).some(r => r.empId === emp.id && r.status === "Pending") ? (
                      <div style={{ padding: "14px 18px", borderRadius: 12, background: "#fef3c7", border: "1px solid #fde68a" }}>
                        <strong>⏳ ጥያቄዎ ለ Admin ተልኳል። ቀጣይ ዉሳኔ ይጠብቁ።</strong>
                      </div>
                    ) : (
                      <form className="form-grid" onSubmit={submitResignation}>
                        <textarea
                          placeholder="የምትለቁበት ምክንያት *"
                          value={resignForm.reason}
                          onChange={e => setResignForm(p => ({ ...p, reason: e.target.value }))}
                          style={{ gridColumn: "1/-1", minHeight: 90, padding: "8px 12px", borderRadius: 8, border: "1px solid #cbd5e1", fontSize: 14, resize: "vertical" }}
                          required
                        />
                        <div style={{ gridColumn: "1/-1" }}>
                          <label style={{ fontSize: 13, color: "#64748b", display: "block", marginBottom: 4 }}>የመጨረሻ የስራ ቀን *</label>
                          <input type="date" value={resignForm.lastDay} onChange={e => setResignForm(p => ({ ...p, lastDay: e.target.value }))}
                            style={{ width: "100%", padding: "10px 12px", border: "1px solid #cbd5e1", borderRadius: 8 }} required />
                        </div>
                        <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0, background: "#ef4444" }}>
                          📝 ስራ ልቀቃ ለ Admin ላክ
                        </button>
                      </form>
                    )}
                    {/* History */}
                    {(db.resignations || []).filter(r => r.empId === emp.id).length > 0 && (
                      <div style={{ marginTop: 16 }}>
                        <h4 style={{ fontSize: 14, margin: "0 0 8px" }}>የቀደሙ ጥያቄዎች</h4>
                        <table className="report-table">
                          <thead><tr><th>ምክንያት</th><th>የመጨረሻ ቀን</th><th>ሁኔታ</th></tr></thead>
                          <tbody>
                            {(db.resignations || []).filter(r => r.empId === emp.id).map(r => (
                              <tr key={r.id}>
                                <td>{r.reason}</td>
                                <td>{r.lastDay}</td>
                                <td><span className={`status-badge ${r.status === "Acknowledged" ? "approved" : "pending"}`}>{r.status === "Acknowledged" ? "✅ ተቀበለ" : "⏳ Pending"}</span></td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                )}

                {/* ===== PROFILE TAB ===== */}
                {empPortalTab === "profile" && (
                  <div className="emp-section" style={{ marginTop: 0, borderTop: "none", paddingTop: 0 }}>
                    <h3>⚙️ የእኔ መቼቶች</h3>

                    {/* Photo upload */}
                    <div style={{ marginBottom: 24 }}>
                      <h4 style={{ fontSize: 14, margin: "0 0 10px" }}>📷 የፕሮፋይል ፎቶ</h4>
                      <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
                        {emp.photo
                          ? <img src={emp.photo} alt="profile" style={{ width: 72, height: 72, borderRadius: "50%", objectFit: "cover", border: "3px solid #2563eb" }} />
                          : <div style={{ width: 72, height: 72, borderRadius: "50%", background: "#dbeafe", display: "grid", placeItems: "center", fontSize: 28, border: "3px solid #2563eb" }}>👤</div>
                        }
                        <div>
                          <label className="btn" style={{ cursor: "pointer", display: "inline-block", marginTop: 0 }}>
                            📷 ፎቶ ምረጥ
                            <input type="file" accept="image/*" onChange={empUploadPhoto} style={{ display: "none" }} />
                          </label>
                          <p style={{ fontSize: 12, color: "#94a3b8", margin: "6px 0 0" }}>JPG፣ PNG — ከ 2MB ያነሰ</p>
                        </div>
                      </div>
                    </div>

                    {/* Password change — only for non-password-free departments */}
                    {!["Manager"].includes(emp.department) && (
                      <div>
                        <h4 style={{ fontSize: 14, margin: "0 0 10px" }}>🔒 ፓስዋርድ ቀይር</h4>
                        <form className="form-grid" onSubmit={empChangePassword}>
                          <input type="password" placeholder="የአሁኑ ፓስዋርድ" value={empChangePwForm.current}
                            onChange={e => setEmpChangePwForm(p => ({ ...p, current: e.target.value }))} required />
                          <input type="password" placeholder="አዲስ ፓስዋርድ (4-6 ቁጥር)" inputMode="numeric" maxLength={6}
                            value={empChangePwForm.newPw} onChange={e => setEmpChangePwForm(p => ({ ...p, newPw: e.target.value }))} required />
                          <input type="password" placeholder="አዲስ ፓስዋርድ አረጋግጥ" inputMode="numeric" maxLength={6}
                            value={empChangePwForm.confirm} onChange={e => setEmpChangePwForm(p => ({ ...p, confirm: e.target.value }))} required
                            style={{ gridColumn: "1/-1" }} />
                          <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>💾 ፓስዋርድ ቀይር</button>
                        </form>
                      </div>
                    )}
                    {["Manager"].includes(emp.department) && (
                      <p style={{ fontSize: 13, color: "#94a3b8", padding: "10px 14px", background: "#f8fafc", borderRadius: 8 }}>
                        ℹ️ Manager department ፓስዋርድ አይጠቀምም — ስልክ ቁጥር ብቻ ይበቃዎታል።
                      </p>
                    )}
                  </div>
                )}

                {/* Finance Tools — Finance dept only */}
                {emp.department === "Finance" && empPortalTab === "finance" && (
                  <div>
                    <div className="hrm-tabs" style={{ marginBottom: 16 }}>
                      {[["income", "💰 ገቢ"], ["expense", "💸 ወጪ"], ["debt", "📝 እዳ"], ["payroll", "💼 ፔሮል"], ["hrm", "👥 HRM"]].map(([k, l]) => (
                        <button key={k} className={financeTab === k ? "hrm-tab active" : "hrm-tab"} onClick={() => setFinanceTab(k)}>
                          {l}
                          {k === "hrm" && pendingLeave > 0 && <span className="badge">{pendingLeave}</span>}
                        </button>
                      ))}
                    </div>

                    {/* Income */}
                    {financeTab === "income" && (
                      <div>
                        <h4 style={{ margin: "0 0 10px" }}>አዲስ ገቢ ምዝገባ</h4>
                        <IncomeRegistration
                          form={incomeForm}
                          banks={db.bankAccounts || []}
                          onChange={(f, v) => setIncomeForm(p => ({ ...p, [f]: v }))}
                          onSubmit={addIncome}
                        />
                        <h4 style={{ margin: "14px 0 8px" }}>የቅርብ ጊዜ ገቢዎች</h4>
                        <table className="report-table">
                          <thead><tr><th>Bank</th><th>Amount</th><th>Date</th><th>Source</th><th>ድርጊት</th></tr></thead>
                          <tbody>
                            {db.incomeEntries.slice(0, 15).map(e => {
                              const isEditing = editIncomeId === e.id;
                              return isEditing ? (
                                <tr key={e.id} style={{ background: "#f0fdf4" }}>
                                  <td>
                                    <select className="leave-input" value={editIncomeForm.bankId} onChange={ev => setEditIncomeForm(p => ({ ...p, bankId: ev.target.value }))} style={{ width: 110 }}>
                                      {(db.bankAccounts || []).map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                                    </select>
                                  </td>
                                  <td><input className="leave-input" type="number" value={editIncomeForm.amount} onChange={ev => setEditIncomeForm(p => ({ ...p, amount: ev.target.value }))} style={{ width: 90 }} /></td>
                                  <td><input className="leave-input" type="date" value={editIncomeForm.date} onChange={ev => setEditIncomeForm(p => ({ ...p, date: ev.target.value }))} style={{ width: 120 }} /></td>
                                  <td><input className="leave-input" value={editIncomeForm.source} onChange={ev => setEditIncomeForm(p => ({ ...p, source: ev.target.value }))} style={{ width: 110 }} /></td>
                                  <td>
                                    <div style={{ display: "flex", gap: 4 }}>
                                      <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => saveEditIncome(e.id)}>💾</button>
                                      <button className="btn-danger btn-sm" onClick={() => setEditIncomeId(null)}>✕</button>
                                    </div>
                                  </td>
                                </tr>
                              ) : (
                                <tr key={e.id}>
                                  <td>{(db.bankAccounts || []).find(b => b.id === Number(e.bankId))?.name || "?"}</td>
                                  <td>{e.amount.toLocaleString()} Birr</td>
                                  <td>{e.date}</td>
                                  <td>{e.source}</td>
                                  <td>
                                    <div style={{ display: "flex", gap: 4 }}>
                                      <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#2563eb" }} onClick={() => startEditIncome(e)}>✏️</button>
                                      <button className="btn-danger btn-sm" onClick={() => deleteIncome(e.id)}>🗑️ ሰርዝ</button>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {/* Expense */}
                    {financeTab === "expense" && (
                      <div>
                        <h4 style={{ margin: "0 0 10px" }}>አዲስ ወጪ ምዝገባ</h4>
                        <ExpenseRegistration
                          form={expenseForm}
                          banks={db.bankAccounts || []}
                          onChange={(f, v) => setExpenseForm(p => ({ ...p, [f]: v }))}
                          onSubmit={addExpense}
                        />
                        <p style={{ fontSize: 12, color: "#f59e0b", margin: "8px 0 0" }}>⚠️ ወጪ ለ Admin approval ተልኳል — ከዚህ ቀጥሎ admin ሲያፈቅር ከባንክ ይቀነሳል።</p>
                        <h4 style={{ margin: "14px 0 8px" }}>Pending ወጪዎቼ</h4>
                        <table className="report-table">
                          <thead><tr><th>Category</th><th>Amount</th><th>Date</th><th>Status</th><th>ድርጊት</th></tr></thead>
                          <tbody>
                            {db.expenseEntries.slice(0, 10).map(e => (
                              <tr key={e.id}>
                                <td>{e.category}</td>
                                <td>{e.amount.toLocaleString()} Birr</td>
                                <td>{e.date}</td>
                                <td><span className={`status-badge ${e.status === "Approved" ? "approved" : e.status === "Rejected" ? "rejected" : "pending"}`}>{e.status || "Approved"}</span></td>
                                <td>
                                  <div style={{ display: "flex", gap: 4 }}>
                                    {(e.status === "Approved" || !e.status) && (
                                      <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 10px", fontSize: 12, background: "#0f172a" }}
                                        onClick={() => printExpenseReceipt(e.id)}>🖨️ Receipt</button>
                                    )}
                                    <button className="btn-danger btn-sm" onClick={() => deleteExpense(e.id)}>🗑️ ሰርዝ</button>
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {/* Debt */}
                    {financeTab === "debt" && (
                      <div>
                        <h4 style={{ margin: "0 0 10px" }}>አዲስ እዳ ምዝገባ</h4>
                        <div className="form-grid" style={{ marginBottom: 12 }}>
                          <input placeholder="Organization" value={debtForm.organization} onChange={e => setDebtForm(p => ({ ...p, organization: e.target.value }))} />
                          <input type="number" placeholder="Total Debt" value={debtForm.total} onChange={e => setDebtForm(p => ({ ...p, total: e.target.value }))} />
                          <input type="number" placeholder="Paid Amount" value={debtForm.paid} onChange={e => setDebtForm(p => ({ ...p, paid: e.target.value }))} />
                          <select value={debtForm.status} onChange={e => setDebtForm(p => ({ ...p, status: e.target.value }))}>
                            {["Pending", "Partial", "Paid"].map(s => <option key={s}>{s}</option>)}
                          </select>
                          <input type="date" value={debtForm.date} onChange={e => setDebtForm(p => ({ ...p, date: e.target.value }))} />
                        </div>
                        <button className="btn" style={{ marginTop: 0 }} onClick={addDebt}>Save Debt</button>
                        <h4 style={{ margin: "14px 0 8px" }}>የእዳ ዝርዝር</h4>
                        <table className="report-table">
                          <thead><tr><th>Organization</th><th>Total</th><th>Paid</th><th>Remaining</th><th>Status</th><th>ድርጊት</th></tr></thead>
                          <tbody>
                            {db.debts.map(d => (
                              <tr key={d.id}>
                                <td>{d.organization}</td>
                                <td>{d.total.toLocaleString()}</td>
                                <td style={{ color: "#16a34a" }}>{d.paid.toLocaleString()}</td>
                                <td style={{ color: "#ef4444", fontWeight: 700 }}>{(d.total - d.paid).toLocaleString()}</td>
                                <td><span className={`status-badge ${d.status === "Paid" ? "approved" : d.status === "Partial" ? "pending" : "rejected"}`}>{d.status}</span></td>
                                <td>
                                  <button className="btn-danger btn-sm" onClick={() => deleteDebt(d.id)}>🗑️ ሰርዝ</button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    )}

                    {/* Payroll */}
                    {financeTab === "payroll" && (
                      <div>
                        <h4 style={{ margin: "0 0 10px" }}>ፔሮል — ደሞዝ ከፍያ</h4>
                        <div className="form-grid" style={{ marginBottom: 12 }}>
                          <select value={payrollBankId} onChange={e => setPayrollBankId(e.target.value)}>
                            <option value="">ባንክ ምረጥ</option>
                            {(db.bankAccounts || []).map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                          </select>
                          <input type="date" value={payrollDate} onChange={e => setPayrollDate(e.target.value)} />
                        </div>
                        {(db.employees || []).length > 0 && (
                          <div className="payroll-totals">
                            <div className="payroll-total-row"><span>ሰራተኞች</span><strong>{(db.employees || []).length}</strong></div>
                            <div className="payroll-total-row pension-row"><span>ጡረታ</span><strong>-{(db.employees || []).reduce((s, e) => s + e.pension, 0).toFixed(0)} Birr</strong></div>
                            <div className="payroll-total-row net-row"><span>ጠቅላላ የሚከፈል</span><strong>{(db.employees || []).reduce((s, e) => s + e.netPay, 0).toFixed(0)} Birr</strong></div>
                          </div>
                        )}
                        <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
                          <button className="btn" style={{ flex: 1 }} onClick={payAllSalaries}>✅ ደን — ወጪ ምዝግብ</button>
                          <button className="secondary-btn" style={{ flex: 1 }} onClick={exportPayroll}>📥 Excel</button>
                        </div>

                        {/* Per-employee payroll edit */}
                        {(db.employees || []).length > 0 && (
                          <div style={{ marginTop: 16 }}>
                            <h4 style={{ fontSize: 14, margin: "0 0 10px" }}>💼 ዲውቲ እና ደሞዝ ማስተካከያ</h4>
                            <table className="report-table">
                              <thead>
                                <tr><th>ስም</th><th>ደሞዝ</th><th>ዲውቲ</th><th>7% ጡረታ</th><th>የሚከፈል</th><th>አርትዕ</th></tr>
                              </thead>
                              <tbody>
                                {(db.employees || []).map(e2 => {
                                  const ed = payrollEdit[e2.id] || {};
                                  const editing = e2.id in payrollEdit;
                                  const curSal = editing ? (ed.basicSalary ?? e2.basicSalary) : e2.basicSalary;
                                  const curDuty = editing ? (ed.duty ?? (e2.duty || 0)) : (e2.duty || 0);
                                  const curPen = editing ? (ed.pensionOpt ?? e2.pensionOpt) : e2.pensionOpt;
                                  const previewPen = curPen ? (Number(curSal) * 0.07).toFixed(0) : 0;
                                  const previewNet = (Number(curSal) * (curPen ? 0.93 : 1) + Number(curDuty)).toFixed(0);
                                  return (
                                    <tr key={e2.id}>
                                      <td>{e2.name}</td>
                                      <td>{editing
                                        ? <input type="number" className="leave-input" style={{ width: 80 }} value={ed.basicSalary ?? e2.basicSalary}
                                          onChange={ev => setPayrollEdit(p => ({ ...p, [e2.id]: { ...p[e2.id], basicSalary: ev.target.value } }))} />
                                        : `${(e2.basicSalary || 0).toLocaleString()}`}
                                      </td>
                                      <td>{editing
                                        ? <input type="number" className="leave-input" style={{ width: 70 }} value={ed.duty ?? (e2.duty || 0)}
                                          onChange={ev => setPayrollEdit(p => ({ ...p, [e2.id]: { ...p[e2.id], duty: ev.target.value } }))} />
                                        : `${(e2.duty || 0).toLocaleString()}`}
                                      </td>
                                      <td>{editing
                                        ? <label style={{ display: "flex", alignItems: "center", gap: 4 }}>
                                          <input type="checkbox" checked={ed.pensionOpt ?? e2.pensionOpt}
                                            onChange={ev => setPayrollEdit(p => ({ ...p, [e2.id]: { ...p[e2.id], pensionOpt: ev.target.checked } }))} />
                                          {curPen ? <span style={{ color: "#ef4444" }}>-{previewPen}</span> : "—"}
                                        </label>
                                        : (e2.pension > 0 ? <span style={{ color: "#ef4444" }}>-{e2.pension.toFixed(0)}</span> : "—")}
                                      </td>
                                      <td style={{ color: "#16a34a", fontWeight: 700 }}>
                                        {editing ? `${previewNet} Birr` : `${(e2.netPay || e2.basicSalary || 0).toFixed(0)} Birr`}
                                      </td>
                                      <td>
                                        {!editing
                                          ? <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12 }}
                                            onClick={() => setPayrollEdit(p => ({ ...p, [e2.id]: {} }))}>✏️</button>
                                          : <div style={{ display: "flex", gap: 4 }}>
                                            <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }}
                                              onClick={() => savePayrollEdit(e2.id)}>💾</button>
                                            <button className="btn-danger btn-sm"
                                              onClick={() => setPayrollEdit(p => { const n = { ...p }; delete n[e2.id]; return n; })}>✕</button>
                                          </div>
                                        }
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        )}
                      </div>
                    )}

                    {/* HRM tab — inside Finance portal */}
                    {financeTab === "hrm" && (
                      <div>
                        <h4 style={{ margin: "0 0 12px" }}>👥 HRM — የሰራተኞች አስተዳደር</h4>
                        <div className="panel-grid" style={{ marginBottom: 16 }}>
                          <div className="card" style={{ padding: 16 }}>
                            <h4 style={{ margin: "0 0 10px", fontSize: 14 }}>አዲስ ሰራተኛ ምዝገባ</h4>
                            <form className="form-grid" onSubmit={addEmployee}>
                              <input placeholder="ሙሉ ስም" value={empForm.name} onChange={e => setEmpForm(p => ({ ...p, name: e.target.value }))} required />
                              <select value={empForm.department} onChange={e => setEmpForm(p => ({ ...p, department: e.target.value }))}>
                                {DEPARTMENTS.map(d => <option key={d}>{d}</option>)}
                              </select>
                              <input placeholder="ስልክ ቁጥር" value={empForm.phone} onChange={e => setEmpForm(p => ({ ...p, phone: e.target.value }))} />
                              <input placeholder="የባንክ አካውንት ቁጥር" value={empForm.bankAccount} onChange={e => setEmpForm(p => ({ ...p, bankAccount: e.target.value }))} />
                              <input type="date" value={empForm.hireDate} onChange={e => setEmpForm(p => ({ ...p, hireDate: e.target.value }))} title="የተቀጠረበት ቀን" />
                              <input type="number" placeholder="ደሞዝ (Birr)" value={empForm.basicSalary} onChange={e => setEmpForm(p => ({ ...p, basicSalary: e.target.value }))} required />
                              <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>+ ምዝግብ</button>
                            </form>
                          </div>

                          {/* Leave management */}
                          <div className="card" style={{ padding: 16 }}>
                            <h4 style={{ margin: "0 0 10px", fontSize: 14 }}>🗓️ ፍቃድ አስተዳደር {pendingLeave > 0 && <span className="badge">{pendingLeave} Pending</span>}</h4>
                            <table className="report-table">
                              <thead><tr><th>ሰራተኛ</th><th>ቀናት</th><th>ጀምሮ</th><th>አይነት</th><th>ሁኔታ</th><th>ድርጊት</th></tr></thead>
                              <tbody>
                                {(db.leaveRequests || []).length === 0
                                  ? <tr><td colSpan={6} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ጥያቄ የለም።</td></tr>
                                  : (db.leaveRequests || []).map(r => {
                                    const e2 = (db.employees || []).find(e => e.id === r.empId);
                                    return (
                                      <tr key={r.id}>
                                        <td>{e2?.name || r.empId}</td>
                                        <td>{r.days}</td>
                                        <td>{r.startDate}</td>
                                        <td><span className={`status-badge ${r.leaveType === "emergency" ? "rejected" : "pending"}`}>{r.leaveType === "emergency" ? "🚨 Emergency" : "🗓️ Annual"}</span></td>
                                        <td><span className={`status-badge ${r.status === "Approved" ? "approved" : r.status === "Rejected" ? "rejected" : "pending"}`}>{r.status}</span></td>
                                        <td>
                                          <div style={{ display: "flex", gap: 4 }}>
                                            {r.status === "Pending" && <>
                                              <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => approveLeave(r.id)}>✅</button>
                                              <button className="btn-danger btn-sm" onClick={() => rejectLeave(r.id)}>❌</button>
                                            </>}
                                            <button className="btn-danger btn-sm" onClick={() => deleteLeaveRequest(r.id)}>🗑️</button>
                                          </div>
                                        </td>
                                      </tr>
                                    );
                                  })
                                }
                              </tbody>
                            </table>
                          </div>
                        </div>

                        {/* Employee list */}
                        <div className="card" style={{ padding: 16 }}>
                          <h4 style={{ margin: "0 0 10px", fontSize: 14 }}>የሰራተኞች ዝርዝር ({(db.employees || []).length})</h4>
                          <table className="report-table">
                            <thead><tr><th>ስም</th><th>ዲፓርትመንት</th><th>ስልክ</th><th>አካውንት</th><th>ደሞዝ</th><th>ቀጠሮ</th><th>ድርጊት</th></tr></thead>
                            <tbody>
                              {(db.employees || []).length === 0
                                ? <tr><td colSpan={7} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ሰራተኛ የለም።</td></tr>
                                : (db.employees || []).map(e2 => (
                                  <tr key={e2.id}>
                                    <td>{e2.name}</td>
                                    <td><span className="dept-badge">{e2.department}</span></td>
                                    <td>{e2.phone || "—"}</td>
                                    <td>{e2.bankAccount || "—"}</td>
                                    <td style={{ color: "#16a34a", fontWeight: 700 }}>{e2.netPay.toFixed(0)} Birr</td>
                                    <td>{e2.hireDate || "—"}</td>
                                    <td><button className="btn-danger btn-sm" onClick={() => deleteEmployee(e2.id)}>ሰርዝ</button></td>
                                  </tr>
                                ))
                              }
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                  </div>
                )}

                <button className="secondary-btn" style={{ marginTop: 14 }} onClick={() => setShowEmpPortal(false)}>← Back to Login</button>
              </div>
            </div>
          </div>
        );
      }
      // Employee login form
      // ስልክ ቁጥር ሲታይ ፓስዋርድ-free department ነው እንደሆነ ይለካ
      const phoneEmp = (db.employees || []).find(em => em.phone === empLoginForm.phone.trim() && em.hasPortalAccess);
      const isPasswordFree = phoneEmp && ["Manager"].includes(phoneEmp.department);

      return (
        <div className={`app-shell ${dark ? "dark-mode" : ""}`}>
          <div className="auth-screen">
            <div className="auth-card">
              <h2 style={{ margin: "0 0 6px" }}>Employee Portal</h2>
              <p style={{ margin: "0 0 16px", color: "#64748b", fontSize: 13 }}>
                {isPasswordFree
                  ? "ስልክ ቁጥርዎን ያስገቡ — ፓስዋርድ አያስፈልግም።"
                  : "ስልክ ቁጥር እና ፓስዋርድ ያስገቡ።"}
              </p>
              {toast && <div className={`toast ${toast.type}`}>{toast.message}</div>}
              <form className="auth-form" onSubmit={handleEmpLogin} autoComplete="on">
                <label>ስልክ ቁጥር<input
                  type="tel"
                  placeholder="0912345678"
                  value={empLoginForm.phone}
                  onChange={e => setEmpLoginForm(p => ({ ...p, phone: e.target.value }))}
                  autoComplete="username"
                  inputMode="tel"
                  autoCapitalize="off"
                  autoCorrect="off"
                  spellCheck="false"
                  style={{ fontSize: 16 }}
                /></label>
                {!isPasswordFree && (
                  <label>ፓስዋርድ<input
                    type="password"
                    placeholder="1234"
                    value={empLoginForm.password}
                    onChange={e => setEmpLoginForm(p => ({ ...p, password: e.target.value }))}
                    autoComplete="current-password"
                    inputMode="numeric"
                    maxLength={6}
                    style={{ fontSize: 16 }}
                  /></label>
                )}
                <button className="btn" type="submit" style={{ touchAction: "manipulation" }}>ግባ</button>
              </form>
              <button className="secondary-btn" style={{ marginTop: 8 }} onClick={() => setShowEmpPortal(false)}>← Admin Login</button>
            </div>
          </div>
        </div>
      );
    }

    // Admin login screen
    return (
      <div className={`app-shell ${dark ? "dark-mode" : ""}`}>
        <div className="auth-screen">
          <div className="auth-card">
            {/* ── Brand header ── */}
            <div className="auth-brand-top" style={{ flexDirection: "column", alignItems: "center", textAlign: "center", gap: 14, marginBottom: 20 }}>
              {db.organizationLogo
                ? <img className="auth-logo" src={db.organizationLogo} alt="logo" style={{ width: 96, height: 96, borderRadius: 18 }} />
                : <div className="auth-logo-placeholder" style={{ width: 96, height: 96, fontSize: 36, borderRadius: 18 }}>🏥</div>
              }
              <div>
                <h1 style={{ margin: 0, fontSize: 22, fontWeight: 800, color: "#1e3a5f", lineHeight: 1.2 }}>
                  {db.organizationLogoText || db.organizationName}
                </h1>
                <p style={{ margin: "6px 0 0", fontSize: 13, color: "#64748b", fontWeight: 500 }}>
                  Finance Management System
                </p>
              </div>
            </div>

            {/* ── Divider ── */}
            <div style={{ height: 1, background: "linear-gradient(90deg,transparent,#cbd5e1,transparent)", marginBottom: 20 }} />

            <p style={{ color: "#64748b", margin: "0 0 18px", fontSize: 14, textAlign: "center" }}>{t.signIn}</p>
            {toast && <div className={`toast ${toast.type}`}>{toast.message}</div>}
            <form className="auth-form" onSubmit={handleLogin} autoComplete="on">
              <label>Phone Number<input
                type="tel"
                placeholder="0912345678"
                value={loginForm.phoneNumber}
                onChange={e => setLoginForm(p => ({ ...p, phoneNumber: e.target.value }))}
                autoComplete="username"
                inputMode="tel"
                autoCapitalize="off"
                autoCorrect="off"
                spellCheck="false"
                style={{ fontSize: 16 }}
              /></label>
              <label>{t.password}<input
                type="password"
                inputMode="numeric"
                maxLength={4}
                placeholder="••••"
                value={loginForm.password}
                onChange={e => setLoginForm(p => ({ ...p, password: e.target.value }))}
                autoComplete="current-password"
                style={{ fontSize: 16 }}
              /></label>
              <button className="btn" type="submit" style={{ touchAction: "manipulation", padding: "13px", fontSize: 15, borderRadius: 10, letterSpacing: "0.03em" }}>{t.login}</button>
            </form>
            <button className="secondary-btn" style={{ marginTop: 8, width: "100%", textAlign: "center" }} onClick={() => setShowNewAcc(p => !p)}>{showNewAcc ? t.hideAccount : t.createAccount}</button>
            {showNewAcc && (
              <div className="create-account-box">
                <h3>{t.createAccount}</h3>
                <form className="create-account-form" onSubmit={handleCreateAccount}>
                  <label>{t.fullName}<input value={newAccForm.fullName} onChange={e => setNewAccForm(p => ({ ...p, fullName: e.target.value }))} placeholder="Full name" /></label>
                  <label>Phone<input type="tel" value={newAccForm.phoneNumber} onChange={e => setNewAccForm(p => ({ ...p, phoneNumber: e.target.value }))} placeholder="0912345678" /></label>
                  <label>4-Digit Password<input type="password" inputMode="numeric" maxLength={4} value={newAccForm.password} onChange={e => setNewAccForm(p => ({ ...p, password: e.target.value }))} placeholder="1234" /></label>
                  <label>Email<input type="email" value={newAccForm.email} onChange={e => setNewAccForm(p => ({ ...p, email: e.target.value }))} placeholder="name@domain.com" /></label>
                  <label>{t.newCreditAccount}<input value={newAccForm.creditAccount} onChange={e => setNewAccForm(p => ({ ...p, creditAccount: e.target.value }))} placeholder="ACC-001" /></label>
                  <button className="btn" type="submit">{t.saveNewAccount}</button>
                </form>
              </div>
            )}
            <div style={{ marginTop: 12, display: "flex", gap: 8 }}>
              <button className="secondary-btn" style={{ marginTop: 0, flex: 1, background: "#2563eb", color: "white", display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }} onClick={() => setShowEmpPortal(true)}>👤 Employee Portal</button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  // ========== MAIN DASHBOARD ==========
  const maxDept = Math.max(...deptStats.map(d => d.count), 1);
  return (
    <div className={`app-shell ${dark ? "dark-mode" : ""}`}>
      <div className="finance-dashboard">
        {/* SIDEBAR */}
        <aside className="sidebar">
          <div className="sidebar-mobile-top">
            <button
              className="sidebar-menu-toggle"
              onClick={() => setMenuExpanded(p => !p)}
              aria-label="Toggle menu"
            >
              ☰ Menu
            </button>
            <button
              className="hamburger-btn"
              onClick={() => setSidebarOpen(p => !p)}
              aria-label="Toggle menu"
            >
              {sidebarOpen ? "✕" : "☰"}
            </button>
          </div>

          <nav className={`menu${(menuExpanded || sidebarOpen) ? " mobile-open" : " collapsed"}`}>
            <div className="menu-group-label">ADMIN</div>
            {ADMIN_ITEMS.map((item, i) => (
              <button key={item} className={section === item ? "active" : ""} onClick={() => { setSection(item); setSidebarOpen(false); setMenuExpanded(true); if (item === "Messages") markAdminMsgsRead(); }}>
                {(ADMIN_LABELS[lang] || ADMIN_LABELS.en)[i]}
                {item === "Messages" && unreadAdmin > 0 && <span className="badge">{unreadAdmin}</span>}
                {item === "Approve Expenses" && pendingExpenses.length > 0 && <span className="badge">{pendingExpenses.length}</span>}
                {item === "Resignations" && pendingResignations > 0 && <span className="badge">{pendingResignations}</span>}
                {item === "Create User" && pendingLeave > 0 && <span className="badge">{pendingLeave}</span>}
                {item === "HRM" && pendingLeave > 0 && <span className="badge">{pendingLeave}</span>}
              </button>
            ))}
          </nav>
          <button className="logout-btn" onClick={() => { setIsAdmin(false); showToast("success", "Logged out."); }}>Logout</button>
        </aside>

        {/* CONTENT */}
        <main className="content">
          <div className="topbar">
            <h1>{SECTION_TITLES[section]}</h1>
            <div className="topbar-actions">
              <button className="language-toggle" onClick={() => setLang(p => p === "en" ? "am" : "en")}>{t.language}</button>
              <button className="theme-toggle" onClick={() => setDark(p => !p)}>{dark ? t.lightMode : t.darkMode}</button>
              <div className="chip">Dr Hibist Finance</div>
            </div>
          </div>
          {toast && <div className={`toast ${toast.type}`}>{toast.message}</div>}

          {/* ===== DASHBOARD ===== */}
          {section === "Dashboard" && (
            <>
              <div className="dashboard-widget-actions">
                <button className="secondary-btn" onClick={() => setShowDashboardCards(p => !p)}>
                  {showDashboardCards ? "Hide Dashboard Boxes" : "Show Dashboard Boxes"}
                </button>
              </div>

              {!showDashboardCards && (
                <div className="card empty-dashboard-card">
                  <h3>Dashboard boxes are hidden</h3>
                  <p>Use the button above to show the dashboard widgets when needed.</p>
                </div>
              )}

              {showDashboardCards && (
                <>
                  {/* Summary cards */}
                  <section className="summary-grid">
                    <div className="card"><h3>Annual Income</h3><p className="metric">{yearly.totalIncome.toLocaleString()} Birr</p></div>
                    <div className="card"><h3>Annual Expense</h3><p className="metric">{yearly.totalExpense.toLocaleString()} Birr</p></div>
                    <div className="card"><h3>Net Profit</h3><p className="metric" style={{ color: yearly.netProfit >= 0 ? "#16a34a" : "#ef4444" }}>{yearly.netProfit.toLocaleString()} Birr</p></div>
                    <div className="card"><h3>Unpaid Debt</h3><p className="metric">{yearly.remainingDebt.toLocaleString()} Birr</p></div>
                  </section>

                  {/* Monthly summary cards */}
                  <section className="summary-grid" style={{ marginTop: 16 }}>
                    <div className="card" style={{ borderTop: "3px solid #2563eb" }}><h3>📅 This Month Income</h3><p className="metric" style={{ fontSize: 20, color: "#2563eb" }}>{monthly.monthIncome.toLocaleString()} Birr</p></div>
                    <div className="card" style={{ borderTop: "3px solid #ef4444" }}><h3>📅 This Month Expense</h3><p className="metric" style={{ fontSize: 20, color: "#ef4444" }}>{monthly.monthExpense.toLocaleString()} Birr</p></div>
                    <div className="card" style={{ borderTop: "3px solid #16a34a" }}><h3>📅 Monthly Profit</h3><p className="metric" style={{ fontSize: 20, color: monthly.remainingProfit >= 0 ? "#16a34a" : "#ef4444" }}>{monthly.remainingProfit.toLocaleString()} Birr</p></div>
                    <div className="card" style={{ borderTop: "3px solid #f59e0b" }}><h3>📅 This Month Debt</h3><p className="metric" style={{ fontSize: 20, color: "#f59e0b" }}>{monthly.monthDebt.toLocaleString()} Birr</p></div>
                  </section>

                  {/* HRM quick stats */}
                  <section className="summary-grid" style={{ marginTop: 16 }}>
                    <div className="card">
                      <h3>👥 Total Employees</h3>
                      <p className="metric">{(db.employees || []).length}</p>
                    </div>
                    <div className="card">
                      <h3>🏢 Departments</h3>
                      <p className="metric">{deptStats.length}</p>
                    </div>
                    <div className="card">
                      <h3>🗓️ Pending Leave</h3>
                      <p className="metric">{pendingLeave}</p>
                    </div>
                    <div className="card">
                      <h3>💬 Unread Messages</h3>
                      <p className="metric">{unreadAdmin}</p>
                    </div>
                  </section>

                  {/* Bank balances */}
                  <section className="bank-summary-grid">
                    {bankSummaries.map(b => (
                      <div key={b.id} className="card bank-card">
                        <div className="bank-info"><h4>{b.name}</h4><span className="acc-num">{b.accountNumber}</span></div>
                        <p className="bank-balance">{b.balance.toLocaleString()} Birr</p>
                      </div>
                    ))}
                  </section>
                </>
              )}

              {/* Department bar chart */}
              {deptStats.length > 0 && (
                <div className="card" style={{ marginTop: 18 }}>
                  <h3>Department Distribution</h3>
                  <div className="dept-chart">
                    {deptStats.map(d => (
                      <div key={d.dept} className="dept-bar-row">
                        <span className="dept-bar-label">{d.dept}</span>
                        <div className="dept-bar-track">
                          <div className="dept-bar-fill" style={{ width: `${(d.count / maxDept) * 100}%` }} />
                        </div>
                        <span className="dept-bar-count">{d.count}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* Expense chart + Financial snapshot */}
              <section className="panel-grid">
                <div className="card">
                  <h3>Expense by Category</h3>
                  <div className="chart-list">
                    {groupedExpenses.map(({ cat, amt }) => {
                      const w = Math.max((amt / Math.max(...groupedExpenses.map(e => e.amt), 1)) * 100, 8);
                      return (<div key={cat} className="chart-row">
                        <div className="chart-label"><span>{cat}</span><strong>{amt.toLocaleString()} Birr</strong></div>
                        <div className="chart-track"><div className="chart-bar" style={{ width: `${w}%` }} /></div>
                      </div>);
                    })}
                  </div>
                </div>
                <div className="card">
                  <h3>Financial Snapshot</h3>
                  <div className="report-rows">
                    {[["Income", yearly.totalIncome], ["Expense", yearly.totalExpense], ["Profit", yearly.netProfit], ["Outstanding Debt", yearly.remainingDebt]].map(([l, v]) => (
                      <div key={l} className="report-row"><span>{l}</span><strong>{v.toLocaleString()} Birr</strong></div>
                    ))}
                  </div>
                </div>
              </section>

              {/* Leave summary table */}
              {(db.employees || []).filter(e => e.totalLeaveDays > 0).length > 0 && (
                <div className="card" style={{ marginTop: 18 }}>
                  <h3>Employee Leave Summary</h3>
                  <table className="report-table">
                    <thead><tr><th>ስም</th><th>ዲፓርትመንት</th><th>ጠቅላላ ፍቃድ</th><th>የወሰዱ</th><th>የቀረ</th></tr></thead>
                    <tbody>
                      {(db.employees || []).filter(e => e.totalLeaveDays > 0).map(e => {
                        const rem = (e.totalLeaveDays || 0) - (e.usedLeaveDays || 0);
                        return (<tr key={e.id}>
                          <td>{e.name}</td><td>{e.department}</td>
                          <td>{e.totalLeaveDays}</td><td>{e.usedLeaveDays}</td>
                          <td style={{ color: rem < 3 ? "#ef4444" : "#16a34a", fontWeight: 700 }}>{rem}{rem < 3 ? " ⚠️" : ""}</td>
                        </tr>);
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}

          {/* ===== HRM ===== */}
          {section === "HRM" && (
            <section>
              <div className="hrm-tabs" style={{ marginBottom: 18, flexWrap: "wrap" }}>
                {[
                  ["employees", "👥 ሰራተኞች"],
                  ["payroll", "💼 ፔሮል"],
                  ["leave", `🗓️ ፍቃድ${pendingLeave > 0 ? ` (${pendingLeave})` : ""}`],
                  ["messages", `💬 መልዕክቶች${unreadAdmin > 0 ? ` (${unreadAdmin})` : ""}`],
                  ["announcements", "📢 ማስታወቂያ"],
                  ["resignations", `📝 ስራ መልቀቃ${pendingResignations > 0 ? ` (${pendingResignations})` : ""}`],
                  ["settings", "⚙️ መቼቶች"],
                ].map(([k, l]) => (
                  <button key={k} className={hrmTab === k ? "hrm-tab active" : "hrm-tab"} onClick={() => { setHrmTab(k); if (k === "messages") markAdminMsgsRead(); }}>{l}</button>
                ))}
              </div>

              {/* ---- EMPLOYEES TAB ---- */}
              {hrmTab === "employees" && (
                <div className="panel-grid">
                  <div className="card">
                    <h3>አዲስ ሰራተኛ ምዝገባ</h3>
                    <form className="form-grid" onSubmit={addEmployee}>
                      <input placeholder="ሙሉ ስም *" value={empForm.name} onChange={e => setEmpForm(p => ({ ...p, name: e.target.value }))} required />
                      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                        <select value={empForm.department} onChange={e => { setEmpForm(p => ({ ...p, department: e.target.value })); if (e.target.value !== "Other") setCustomDept(""); }}>
                          {DEPARTMENTS.map(d => <option key={d}>{d}</option>)}
                        </select>
                        {empForm.department === "Other" && (
                          <input placeholder="ዲፓርትመንት ስም ጻፍ *" value={customDept} onChange={e => setCustomDept(e.target.value)} required style={{ padding: "8px 10px", border: "1px solid #2563eb", borderRadius: 8, fontSize: 13 }} />
                        )}
                      </div>
                      <input placeholder="ስልክ ቁጥር" value={empForm.phone} onChange={e => setEmpForm(p => ({ ...p, phone: e.target.value }))} />
                      <input placeholder="የባንክ አካውንት ቁጥር" value={empForm.bankAccount} onChange={e => setEmpForm(p => ({ ...p, bankAccount: e.target.value }))} />
                      <input type="date" title="የተቀጠረበት ቀን" value={empForm.hireDate} onChange={e => setEmpForm(p => ({ ...p, hireDate: e.target.value }))} />
                      <input type="number" placeholder="ደሞዝ (Birr) *" value={empForm.basicSalary} onChange={e => setEmpForm(p => ({ ...p, basicSalary: e.target.value }))} required />
                      <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>+ ሰራተኛ ምዝግብ</button>
                    </form>
                  </div>

                  <div className="card">
                    <h3>የሰራተኞች ዝርዝር ({(db.employees || []).length})</h3>
                    <table className="report-table">
                      <thead>
                        <tr><th>ስም</th><th>ዲፓርትመንት</th><th>ስልክ</th><th>አካውንት</th><th>ደሞዝ (Net)</th><th>ቀጠሮ</th><th>ድርጊት</th></tr>
                      </thead>
                      <tbody>
                        {(db.employees || []).length === 0 ? (
                          <tr><td colSpan={7} style={{ textAlign: "center", color: "#94a3b8", padding: 20 }}>ምንም ሰራተኛ አልተመዘገበም።</td></tr>
                        ) : (
                          (db.employees || []).map(e => (
                            <tr key={e.id}>
                              <td style={{ fontWeight: 600 }}>{e.name}</td>
                              <td><span className="dept-badge">{e.department}</span></td>
                              <td>{e.phone || "—"}</td>
                              <td>{e.bankAccount || "—"}</td>
                              <td style={{ color: "#16a34a", fontWeight: 700 }}>{(e.netPay || e.basicSalary).toLocaleString()} Birr</td>
                              <td>{e.hireDate || "—"}</td>
                              <td>
                                <button className="btn-danger btn-sm" onClick={() => deleteEmployee(e.id)}>ሰርዝ</button>
                              </td>
                            </tr>
                          ))
                        )}
                      </tbody>
                    </table>

                    {/* Department stats */}
                    {deptStats.length > 0 && (
                      <div style={{ marginTop: 20 }}>
                        <h4 style={{ margin: "0 0 10px", fontSize: 14, color: "#334155" }}>ዲፓርትመንት ስርጭት</h4>
                        <div className="dept-chart">
                          {deptStats.map(d => (
                            <div key={d.dept} className="dept-bar-row">
                              <span className="dept-bar-label">{d.dept}</span>
                              <div className="dept-bar-track">
                                <div className="dept-bar-fill" style={{ width: `${(d.count / Math.max(...deptStats.map(x => x.count), 1)) * 100}%` }} />
                              </div>
                              <span className="dept-bar-count">{d.count}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* ---- PAYROLL TAB ---- */}
              {hrmTab === "payroll" && (
                <div>
                  <div className="panel-grid" style={{ marginBottom: 18 }}>
                    <div className="card">
                      <h3>ደሞዝ ከፍያ (Payroll)</h3>
                      <p style={{ color: "#64748b", fontSize: 13, margin: "0 0 12px" }}>ሁሉም ሰራተኞች ደሞዝ ሲደን ከዚህ ባንክ ወጪ ይደረጋል።</p>
                      <div className="form-grid" style={{ marginBottom: 12 }}>
                        <select value={payrollBankId} onChange={e => setPayrollBankId(e.target.value)}>
                          <option value="">ባንክ ምረጥ</option>
                          {(db.bankAccounts || []).map(b => <option key={b.id} value={b.id}>{b.name} ({b.accountNumber})</option>)}
                        </select>
                        <input type="date" value={payrollDate} onChange={e => setPayrollDate(e.target.value)} />
                      </div>
                      {(db.employees || []).length > 0 && (
                        <div className="payroll-totals">
                          <div className="payroll-total-row"><span>ሰራተኞች ብዛት</span><strong>{(db.employees || []).length}</strong></div>
                          <div className="payroll-total-row"><span>ጠቅላላ ደሞዝ</span><strong>{(db.employees || []).reduce((s, e) => s + (e.basicSalary || 0), 0).toLocaleString()} Birr</strong></div>
                          <div className="payroll-total-row pension-row"><span>ጡረታ ቅናሽ</span><strong>-{(db.employees || []).reduce((s, e) => s + (e.pension || 0), 0).toFixed(0)} Birr</strong></div>
                          <div className="payroll-total-row net-row"><span>ጠቅላላ የሚከፈል</span><strong>{(db.employees || []).reduce((s, e) => s + (e.netPay || e.basicSalary || 0), 0).toFixed(0)} Birr</strong></div>
                        </div>
                      )}
                      <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
                        <button className="btn" style={{ flex: 1 }} onClick={payAllSalaries}>✅ ደን — ወጪ ምዝግብ</button>
                        <button className="secondary-btn" style={{ flex: 1 }} onClick={exportPayroll}>📥 Excel Export</button>
                      </div>
                    </div>

                    <div className="card">
                      <h3>Payroll Preview</h3>
                      <table className="report-table">
                        <thead><tr><th>ተ.ቁ</th><th>ስም</th><th>ዲፓርትመንት</th><th>ስልክ</th><th>አካውንት</th><th>የሚከፈል</th></tr></thead>
                        <tbody>
                          {(db.employees || []).length === 0 ? (
                            <tr><td colSpan={6} style={{ textAlign: "center", color: "#94a3b8", padding: 20 }}>ምንም ሰራተኛ የለም።</td></tr>
                          ) : (
                            (db.employees || []).map((e, i) => (
                              <tr key={e.id}>
                                <td>{i + 1}</td>
                                <td>{e.name}</td>
                                <td><span className="dept-badge">{e.department}</span></td>
                                <td>{e.phone || "—"}</td>
                                <td>{e.bankAccount || "—"}</td>
                                <td style={{ color: "#16a34a", fontWeight: 700 }}>{(e.netPay || e.basicSalary || 0).toFixed(0)} Birr</td>
                              </tr>
                            ))
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Per-employee payroll edit */}
                  <div className="card">
                    <h3>💼 ፔሮል ማስተካከያ — ዲውቲ፣ ጡረታ፣ ደሞዝ</h3>
                    <p style={{ color: "#64748b", fontSize: 13, margin: "0 0 12px" }}>
                      ለእያንዳዱ ሰራተኛ ዲውቲ ማከል፣ ጡረታ opt-in/out፣ ወይም ደሞዝ ማስተካከል።
                    </p>
                    {(db.employees || []).length === 0 ? (
                      <p style={{ color: "#94a3b8", textAlign: "center", padding: 20 }}>ምንም ሰራተኛ አልተመዘገበም። አስቀድሞ ሰራተኛ ምዝግቡ።</p>
                    ) : (
                      <table className="report-table">
                        <thead>
                          <tr>
                            <th>ስም</th><th>ዲፓርትመንት</th>
                            <th>ደሞዝ</th><th>ዲውቲ</th>
                            <th>7% ጡረታ</th><th>የሚከፈል</th><th>አርትዕ</th>
                          </tr>
                        </thead>
                        <tbody>
                          {(db.employees || []).map(e => {
                            const ed = payrollEdit[e.id] || {};
                            const editing = e.id in payrollEdit;
                            const curSal = editing ? (ed.basicSalary ?? e.basicSalary) : e.basicSalary;
                            const curDuty = editing ? (ed.duty ?? (e.duty || 0)) : (e.duty || 0);
                            const curPen = editing ? (ed.pensionOpt ?? e.pensionOpt) : e.pensionOpt;
                            const previewPen = curPen ? (Number(curSal) * 0.07).toFixed(0) : 0;
                            const previewNet = (Number(curSal) * (curPen ? 0.93 : 1) + Number(curDuty)).toFixed(0);
                            return (
                              <tr key={e.id}>
                                <td>{e.name}</td>
                                <td><span className="dept-badge">{e.department}</span></td>
                                <td>
                                  {editing
                                    ? <input type="number" className="leave-input" style={{ width: 80 }} value={ed.basicSalary ?? e.basicSalary} onChange={ev => setPayrollEdit(p => ({ ...p, [e.id]: { ...p[e.id], basicSalary: ev.target.value } }))} />
                                    : `${(e.basicSalary || 0).toLocaleString()}`
                                  }
                                </td>
                                <td>
                                  {editing
                                    ? <input type="number" className="leave-input" style={{ width: 70 }} value={ed.duty ?? (e.duty || 0)} onChange={ev => setPayrollEdit(p => ({ ...p, [e.id]: { ...p[e.id], duty: ev.target.value } }))} />
                                    : `${(e.duty || 0).toLocaleString()}`
                                  }
                                </td>
                                <td>
                                  {editing
                                    ? <label style={{ display: "flex", alignItems: "center", gap: 4 }}>
                                      <input type="checkbox" checked={ed.pensionOpt ?? e.pensionOpt} onChange={ev => setPayrollEdit(p => ({ ...p, [e.id]: { ...p[e.id], pensionOpt: ev.target.checked } }))} />
                                      {curPen ? <span style={{ color: "#ef4444" }}>-{previewPen}</span> : "—"}
                                    </label>
                                    : (e.pension > 0 ? <span style={{ color: "#ef4444" }}>-{e.pension.toFixed(0)}</span> : "—")
                                  }
                                </td>
                                <td style={{ color: "#16a34a", fontWeight: 700 }}>
                                  {editing ? `${previewNet} Birr` : `${(e.netPay || e.basicSalary || 0).toFixed(0)} Birr`}
                                </td>
                                <td>
                                  {!editing
                                    ? <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 10px", fontSize: 12 }} onClick={() => setPayrollEdit(p => ({ ...p, [e.id]: {} }))}>✏️ አርትዕ</button>
                                    : <div style={{ display: "flex", gap: 4 }}>
                                      <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => savePayrollEdit(e.id)}>💾 አስቀምጥ</button>
                                      <button className="btn-danger btn-sm" onClick={() => setPayrollEdit(p => { const n = { ...p }; delete n[e.id]; return n; })}>✕</button>
                                    </div>
                                  }
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    )}
                  </div>

                </div>
              )}

              {/* ---- LEAVE TAB ---- */}
              {hrmTab === "leave" && (
                <div className="panel-grid">
                  {/* Submit leave request on behalf of employee */}
                  <div className="card">
                    <h3>🗓️ ፍቃድ ጥያቄ ምዝገባ (Admin)</h3>
                    <form className="form-grid" onSubmit={submitLeaveRequest}>
                      <select value={leaveForm.empId} onChange={e => setLeaveForm(p => ({ ...p, empId: e.target.value }))} required>
                        <option value="">ሰራተኛ ምረጥ</option>
                        {(db.employees || []).map(e => <option key={e.id} value={e.id}>{e.name} — {e.department}</option>)}
                      </select>
                      <select value={leaveForm.leaveType} onChange={e => setLeaveForm(p => ({ ...p, leaveType: e.target.value }))}>
                        <option value="annual">🗓️ ዓመታዊ ፍቃድ</option>
                        <option value="emergency">🚨 ድንገተኛ ፍቃድ</option>
                      </select>
                      <input type="number" placeholder="ቀናት ብዛት" value={leaveForm.days} onChange={e => setLeaveForm(p => ({ ...p, days: e.target.value }))} required />
                      <input type="date" value={leaveForm.startDate} onChange={e => setLeaveForm(p => ({ ...p, startDate: e.target.value }))} />
                      <input placeholder="ምክንያት" value={leaveForm.reason} onChange={e => setLeaveForm(p => ({ ...p, reason: e.target.value }))} style={{ gridColumn: "1/-1" }} />
                      <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>ፍቃድ ጥያቄ ምዝግብ</button>
                    </form>

                    {/* Leave balance per employee */}
                    <h4 style={{ margin: "18px 0 10px", fontSize: 14 }}>የሰራተኞች ፍቃድ ባላንስ</h4>
                    <table className="report-table">
                      <thead><tr><th>ስም</th><th>ዲፓርትመንት</th><th>ጠቅላላ</th><th>የወሰዱ</th><th>የቀረ</th></tr></thead>
                      <tbody>
                        {(db.employees || []).length === 0 ? (
                          <tr><td colSpan={5} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ሰራተኛ የለም።</td></tr>
                        ) : (
                          (db.employees || []).map(e => {
                            const rem = (e.totalLeaveDays || 25) - (e.usedLeaveDays || 0);
                            return (
                              <tr key={e.id}>
                                <td>{e.name}</td>
                                <td><span className="dept-badge">{e.department}</span></td>
                                <td>{e.totalLeaveDays || 25}</td>
                                <td>{e.usedLeaveDays || 0}</td>
                                <td style={{ color: rem < 3 ? "#ef4444" : "#16a34a", fontWeight: 700 }}>{rem}{rem < 3 ? " ⚠️" : ""}</td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>

                  {/* Pending leave approvals */}
                  <div className="card">
                    <h3>ፍቃድ ጥያቄዎች {pendingLeave > 0 && <span className="badge">{pendingLeave} Pending</span>}</h3>
                    <table className="report-table">
                      <thead>
                        <tr><th>ሰራተኛ</th><th>አይነት</th><th>ቀናት</th><th>ጀምሮ</th><th>ምክንያት</th><th>ሁኔታ</th><th>ድርጊት</th></tr>
                      </thead>
                      <tbody>
                        {(db.leaveRequests || []).length === 0 ? (
                          <tr><td colSpan={7} style={{ textAlign: "center", color: "#94a3b8", padding: 20 }}>ምንም ፍቃድ ጥያቄ አልቀረበም።</td></tr>
                        ) : (
                          (db.leaveRequests || []).map(r => {
                            const emp = (db.employees || []).find(e => e.id === r.empId);
                            const rem = emp ? (emp.totalLeaveDays || 25) - (emp.usedLeaveDays || 0) : "—";
                            const adjDays = leaveApproveEdit[r.id] !== undefined ? leaveApproveEdit[r.id] : r.days;
                            const rejectR = leaveRejectReason[r.id] || "";
                            return (
                              <tr key={r.id}>
                                <td style={{ fontWeight: 600 }}>{emp?.name || r.empId}</td>
                                <td>
                                  <span className={`status-badge ${r.leaveType === "emergency" ? "rejected" : "pending"}`}>
                                    {r.leaveType === "emergency" ? "🚨 Emergency" : "🗓️ Annual"}
                                  </span>
                                </td>
                                <td>
                                  {r.status === "Pending"
                                    ? <input type="number" min={1} className="leave-input" style={{ width: 54 }} value={adjDays} onChange={ev => setLeaveApproveEdit(p => ({ ...p, [r.id]: ev.target.value }))} />
                                    : (r.approvedDays || r.days)
                                  }
                                </td>
                                <td>{r.startDate}</td>
                                <td style={{ fontSize: 12 }}>{r.reason || "—"}</td>
                                <td>
                                  <span className={`status-badge ${r.status === "Approved" ? "approved" : r.status === "Rejected" ? "rejected" : "pending"}`}>
                                    {r.status}
                                  </span>
                                  {r.rejectReason && <div style={{ fontSize: 11, color: "#ef4444", marginTop: 2 }}>{r.rejectReason}</div>}
                                </td>
                                <td>
                                  {r.status === "Pending" && (
                                    <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                                      <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => approveLeaveAdjusted(r.id)}>✅ ፍቀድ</button>
                                      <input placeholder="ምክንያት..." className="leave-input" style={{ width: 100, marginTop: 2 }} value={rejectR} onChange={ev => setLeaveRejectReason(p => ({ ...p, [r.id]: ev.target.value }))} />
                                      <button className="btn-danger btn-sm" onClick={() => rejectLeaveWithReason(r.id)}>❌ ሰርዝ</button>
                                    </div>
                                  )}
                                  <button className="btn-danger btn-sm" style={{ marginTop: r.status === "Pending" ? 4 : 0 }} onClick={() => deleteLeaveRequest(r.id)}>🗑️</button>
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* ---- MESSAGES TAB ---- */}
              {hrmTab === "messages" && (
                <div style={{ display: "grid", gridTemplateColumns: "280px 1fr", gap: 16, minHeight: "70vh" }}>
                  {/* Left: employee list */}
                  <div className="card" style={{ padding: 16, overflowY: "auto" }}>
                    <h3 style={{ margin: "0 0 12px", fontSize: 15 }}>ሰራተኞች</h3>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {(db.employees || []).length === 0
                        ? <p style={{ color: "#94a3b8", fontSize: 13 }}>ምንም ሰራተኛ አልተመዘገበም።</p>
                        : (db.employees || []).map(e => {
                          const unread = (db.messages || []).filter(
                            m => m.from === String(e.id) && m.to === "admin" && !m.readByAdmin
                          ).length;
                          const isActive = msgThreadEmpId === String(e.id);
                          return (
                            <button
                              key={e.id}
                              onClick={() => { setMsgThreadEmpId(String(e.id)); markAdminMsgsRead(); }}
                              style={{
                                textAlign: "left", padding: "10px 12px", borderRadius: 10,
                                border: `1px solid ${isActive ? "#2563eb" : "#e2e8f0"}`,
                                background: isActive ? "#2563eb" : "transparent",
                                color: isActive ? "white" : "inherit",
                                cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center",
                              }}
                            >
                              <div>
                                <div style={{ fontWeight: 600, fontSize: 14 }}>{e.name}</div>
                                <div style={{ fontSize: 11, opacity: 0.75 }}>{e.department}</div>
                              </div>
                              {unread > 0 && <span className="badge">{unread}</span>}
                            </button>
                          );
                        })
                      }
                    </div>
                    {/* Quick send form */}
                    <div style={{ marginTop: 16, borderTop: "1px solid #e2e8f0", paddingTop: 12 }}>
                      <p style={{ fontSize: 11, color: "#94a3b8", margin: "0 0 8px" }}>ቀጥታ ላክ:</p>
                      <form onSubmit={sendAdminMsg}>
                        <select value={msgForm.toId} onChange={e => setMsgForm(p => ({ ...p, toId: e.target.value }))}
                          style={{ width: "100%", padding: "6px 8px", borderRadius: 6, border: "1px solid #cbd5e1", marginBottom: 6, fontSize: 13 }} required>
                          <option value="">ሰራተኛ ምረጥ</option>
                          {(db.employees || []).map(e => <option key={e.id} value={e.id}>{e.name}</option>)}
                        </select>
                        <textarea placeholder="መልዕክት..." value={msgForm.body} onChange={e => setMsgForm(p => ({ ...p, body: e.target.value }))}
                          style={{ width: "100%", padding: "6px 8px", borderRadius: 6, border: "1px solid #cbd5e1", minHeight: 60, fontSize: 13, resize: "vertical" }} required />
                        <button className="btn" type="submit" style={{ width: "100%", marginTop: 6 }}>ላክ →</button>
                      </form>
                    </div>
                  </div>

                  {/* Right: chat thread */}
                  <div className="card" style={{ display: "flex", flexDirection: "column", padding: 16 }}>
                    {!msgThreadEmpId ? (
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "center", flex: 1, color: "#94a3b8", flexDirection: "column", gap: 8 }}>
                        <span style={{ fontSize: 40 }}>💬</span>
                        <span>በግራ ሰራተኛ ምረጥ — ውይይቱ ይታያል</span>
                      </div>
                    ) : (() => {
                      const selEmp = (db.employees || []).find(e => String(e.id) === msgThreadEmpId);
                      const thread = [...(db.messages || [])].filter(
                        m => (m.from === msgThreadEmpId && m.to === "admin") ||
                          (m.from === "admin" && m.to === msgThreadEmpId)
                      ).sort((a, b) => a.id - b.id);
                      return (
                        <>
                          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, paddingBottom: 10, borderBottom: "1px solid #e2e8f0" }}>
                            <strong style={{ fontSize: 16 }}>{selEmp?.name}</strong>
                            <span className="dept-badge">{selEmp?.department}</span>
                          </div>
                          <div style={{ flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: 10, maxHeight: "calc(70vh - 200px)", padding: "4px 0" }}>
                            {thread.length === 0
                              ? <p style={{ color: "#94a3b8", fontSize: 13, textAlign: "center", marginTop: 40 }}>ምንም መልዕክት የለም።</p>
                              : thread.map(m => {
                                const fromAdmin = m.from === "admin";
                                return (
                                  <div key={m.id} style={{ display: "flex", justifyContent: fromAdmin ? "flex-end" : "flex-start", alignItems: "flex-start", gap: 4 }}>
                                    <div className={`msg-bubble ${fromAdmin ? "admin-bubble" : "emp-bubble"}`} style={{ maxWidth: "65%" }}>
                                      <div className="msg-from">{fromAdmin ? "Admin" : selEmp?.name} · {m.date}</div>
                                      <p>{m.body}</p>
                                    </div>
                                    <button
                                      onClick={() => deleteMessage(m.id)}
                                      title="ሰርዝ"
                                      style={{ background: "none", border: "none", cursor: "pointer", color: "#ef4444", fontSize: 14, padding: "4px", opacity: 0.6, flexShrink: 0, marginTop: 4 }}
                                      onMouseEnter={e => e.currentTarget.style.opacity = "1"}
                                      onMouseLeave={e => e.currentTarget.style.opacity = "0.6"}
                                    >🗑️</button>
                                  </div>
                                );
                              })
                            }
                          </div>
                          <form onSubmit={sendThreadReply} style={{ display: "flex", gap: 8, marginTop: 10, borderTop: "1px solid #e2e8f0", paddingTop: 10 }}>
                            <input style={{ flex: 1, padding: "9px 12px", border: "1px solid #cbd5e1", borderRadius: 8, fontSize: 14 }}
                              placeholder="መልዕክት ይጻፉ..." value={threadReply} onChange={e => setThreadReply(e.target.value)} />
                            <button className="btn" type="submit" style={{ marginTop: 0, padding: "9px 18px" }}>ላክ</button>
                          </form>
                        </>
                      );
                    })()}
                  </div>
                </div>
              )}

              {/* ---- ANNOUNCEMENTS TAB ---- */}
              {hrmTab === "announcements" && (
                <div className="panel-grid">
                  <div className="card">
                    <h3>📢 አዲስ ማስታወቂያ ላክ</h3>
                    <p style={{ margin: "0 0 12px", color: "#64748b", fontSize: 13 }}>ማስታወቂያ ለሁሉም ሰራተኞች Portal ላይ ይታያቸዋል።</p>
                    <form onSubmit={sendAnnouncement}>
                      <div className="form-grid">
                        <input placeholder="ርዕስ (Title)" value={announcementForm.title}
                          onChange={e => setAnnouncementForm(p => ({ ...p, title: e.target.value }))} required style={{ gridColumn: "1/-1" }} />
                        <textarea placeholder="ዝርዝር መልዕክት..." value={announcementForm.body}
                          onChange={e => setAnnouncementForm(p => ({ ...p, body: e.target.value }))}
                          style={{ gridColumn: "1/-1", minHeight: 100 }} required />
                        <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>
                          📢 ለሁሉም ሰራተኞች ላክ
                        </button>
                      </div>
                    </form>
                  </div>
                  <div className="card">
                    <h3>ማስታወቂያ ዝርዝር ({(db.announcements || []).length})</h3>
                    {(db.announcements || []).length === 0
                      ? <p style={{ color: "#94a3b8", fontSize: 13 }}>ምንም ማስታወቂያ አልተላከም።</p>
                      : <div className="msg-list">
                        {(db.announcements || []).map(a => (
                          <div key={a.id} className="msg-bubble admin-bubble">
                            <div className="msg-from" style={{ display: "flex", justifyContent: "space-between" }}>
                              <span>📢 {a.title} · {a.date}</span>
                              <span style={{ color: "#16a34a", fontSize: 11 }}>👁 {(a.readBy || []).length} አነበቡ</span>
                            </div>
                            <p style={{ margin: "6px 0 8px" }}>{a.body}</p>
                            <button className="btn-danger btn-sm" onClick={() => deleteAnnouncement(a.id)}>ሰርዝ</button>
                          </div>
                        ))}
                      </div>
                    }
                  </div>
                </div>
              )}

              {/* ---- RESIGNATIONS TAB ---- */}
              {hrmTab === "resignations" && (
                <div className="card">
                  <h3>📝 ስራ መልቀቃ ጥያቄዎች
                    {pendingResignations > 0 && <span className="badge" style={{ marginLeft: 8 }}>{pendingResignations} Pending</span>}
                  </h3>
                  {(db.resignations || []).length === 0
                    ? <p style={{ color: "#94a3b8", padding: "20px 0" }}>ምንም ስራ መልቀቃ ጥያቄ አልቀረበም።</p>
                    : (
                      <table className="report-table">
                        <thead>
                          <tr><th>ሰራተኛ</th><th>ዲፓርትመንት</th><th>ምክንያት</th><th>የመጨረሻ ቀን</th><th>የቀረበ ቀን</th><th>ሁኔታ</th><th>ድርጊት</th></tr>
                        </thead>
                        <tbody>
                          {(db.resignations || []).map(r => (
                            <tr key={r.id}>
                              <td style={{ fontWeight: 600 }}>{r.empName}</td>
                              <td><span className="dept-badge">{r.department}</span></td>
                              <td>{r.reason}</td>
                              <td style={{ fontWeight: 700, color: "#ef4444" }}>{r.lastDay}</td>
                              <td>{r.submittedDate}</td>
                              <td>
                                <span className={`status-badge ${r.status === "Acknowledged" ? "approved" : "pending"}`}>
                                  {r.status === "Acknowledged" ? "✅ ተቀበለ" : "⏳ Pending"}
                                </span>
                              </td>
                              <td>
                                <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                                  {r.status === "Pending" && (
                                    <button className="btn btn-sm" style={{ marginTop: 0, padding: "5px 12px", fontSize: 12, background: "#16a34a" }}
                                      onClick={() => acknowledgeResignation(r.id)}>
                                      ✅ ተቀበለ
                                    </button>
                                  )}
                                  <button className="btn-danger btn-sm" onClick={() => deleteResignation(r.id)}>🗑️</button>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )
                  }
                </div>
              )}

              {/* ---- SETTINGS TAB (Create User + Leave Days) ---- */}
              {hrmTab === "settings" && (
                <div className="settings-section">
                  {/* Portal access */}
                  <div className="card">
                    <h3>👤 ሰራተኛ Portal Access</h3>
                    <p style={{ margin: "0 0 14px", color: "#64748b", fontSize: 13 }}>
                      ከ HRM ከተመዘገቡ ሰራተኞች ምረጥ፣ ፓስዋርድ ስጥ — Employee Portal ሊጠቀሙ ይችላሉ።
                    </p>
                    <form className="form-grid" onSubmit={createUserAccount}>
                      <select value={createUserForm.empId} onChange={e => setCreateUserForm(p => ({ ...p, empId: e.target.value }))} required
                        style={{ gridColumn: "1/-1" }}>
                        <option value="">ሰራተኛ ምረጥ (ከ HRM)</option>
                        {(db.employees || []).map(e => (
                          <option key={e.id} value={e.id}>{e.name} — {e.department} {e.hasPortalAccess ? "✅" : ""}</option>
                        ))}
                      </select>
                      <input
                        placeholder="ስልክ ቁጥር (ለ Portal login)"
                        type="tel"
                        value={createUserForm.phone}
                        onChange={e => setCreateUserForm(p => ({ ...p, phone: e.target.value }))}
                      />
                      <input
                        placeholder="ፓስዋርድ (4-6 ቁጥር) *"
                        type="password"
                        inputMode="numeric"
                        maxLength={6}
                        value={createUserForm.password}
                        onChange={e => setCreateUserForm(p => ({ ...p, password: e.target.value }))}
                        required
                      />
                      <select value={createUserForm.department} onChange={e => setCreateUserForm(p => ({ ...p, department: e.target.value }))}
                        style={{ gridColumn: "1/-1" }}>
                        <option value="">ዲፓርትመንት ምረጥ (አማራጭ)</option>
                        {DEPARTMENTS.map(d => <option key={d} value={d}>{d}</option>)}
                      </select>
                      <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>✅ Access ስጥ</button>
                    </form>
                    <h4 style={{ margin: "20px 0 10px" }}>Portal Access ያላቸው ሰራተኞች</h4>
                    <table className="report-table">
                      <thead><tr><th>ስም</th><th>ዲፓርትመንት</th><th>Access</th><th>ድርጊት</th></tr></thead>
                      <tbody>
                        {(db.employees || []).length === 0
                          ? <tr><td colSpan={4} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ሰራተኛ አልተመዘገበም።</td></tr>
                          : (db.employees || []).map(e => (
                            <tr key={e.id}>
                              <td>{e.name}</td>
                              <td><span className="dept-badge">{e.department}</span></td>
                              <td>{e.hasPortalAccess
                                ? <span className="status-badge approved">✅ Active</span>
                                : <span className="status-badge pending">No Access</span>}
                              </td>
                              <td>{e.hasPortalAccess && (
                                <button className="btn-danger btn-sm" onClick={() => revokeUserAccess(e.id)}>❌ ሰርዝ</button>
                              )}</td>
                            </tr>
                          ))
                        }
                      </tbody>
                    </table>
                  </div>

                  {/* Annual leave days */}
                  <div className="card">
                    <h3>🗓️ አመታዊ ፍቃድ ቀናት</h3>
                    <p style={{ margin: "0 0 14px", color: "#64748b", fontSize: 13 }}>
                      ለእያንዳንዱ ሰራተኛ አመታዊ ፍቃድ ቀናት ብዛት ይወሰን። Default: 25 ቀናት።
                    </p>
                    <form className="form-grid" onSubmit={setEmployeeLeaveDays}>
                      <select value={leaveDaysForm.empId} onChange={e => setLeaveDaysForm(p => ({ ...p, empId: e.target.value }))} required>
                        <option value="">ሰራተኛ ምረጥ</option>
                        {(db.employees || []).map(e => (
                          <option key={e.id} value={e.id}>{e.name} (አሁን: {e.totalLeaveDays || 25} ቀን)</option>
                        ))}
                      </select>
                      <input type="number" min={1} max={365} placeholder="ቀናት ብዛት"
                        value={leaveDaysForm.days} onChange={e => setLeaveDaysForm(p => ({ ...p, days: e.target.value }))} required />
                      <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>💾 ቀናት አስቀምጥ</button>
                    </form>
                    <h4 style={{ margin: "20px 0 10px" }}>ሰራተኞች — ፍቃድ ባላንስ</h4>
                    <table className="report-table">
                      <thead><tr><th>ስም</th><th>ዲፓርትመንት</th><th>ጠቅላላ</th><th>የወሰዱ</th><th>የቀረ</th></tr></thead>
                      <tbody>
                        {(db.employees || []).length === 0
                          ? <tr><td colSpan={5} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ሰራተኛ የለም።</td></tr>
                          : (db.employees || []).map(e => {
                            const rem = (e.totalLeaveDays || 25) - (e.usedLeaveDays || 0);
                            return (
                              <tr key={e.id}>
                                <td>{e.name}</td>
                                <td><span className="dept-badge">{e.department}</span></td>
                                <td>{e.totalLeaveDays || 25}</td>
                                <td>{e.usedLeaveDays || 0}</td>
                                <td style={{ color: rem < 3 ? "#ef4444" : "#16a34a", fontWeight: 700 }}>{rem}{rem < 3 ? " ⚠️" : ""}</td>
                              </tr>
                            );
                          })
                        }
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

            </section>
          )}

          {/* ===== APPROVE EXPENSES ===== */}
          {section === "Approve Expenses" && (
            <section>
              {/* Pending approvals */}
              <div className="card" style={{ marginBottom: 20 }}>
                <h3>ፔንዲንግ ወጪዎች {pendingExpenses.length > 0 && <span className="badge" style={{ marginLeft: 6 }}>{pendingExpenses.length}</span>}</h3>
                {pendingExpenses.length === 0 ? (
                  <p style={{ color: "#94a3b8" }}>ምንም ፔንዲንግ ወጪ የለም።</p>
                ) : (
                  <table className="report-table">
                    <thead>
                      <tr><th>Category</th><th>Amount</th><th>Bank</th><th>Date</th><th>Note</th><th>Action</th></tr>
                    </thead>
                    <tbody>
                      {pendingExpenses.map(e => (
                        <tr key={e.id}>
                          <td>{e.category}</td>
                          <td><strong>{e.amount.toLocaleString()} Birr</strong></td>
                          <td>{(db.bankAccounts || []).find(b => b.id === Number(e.bankId))?.name || "?"}</td>
                          <td>{e.date}</td>
                          <td>{e.note}</td>
                          <td>
                            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                              <button className="btn btn-sm" style={{ marginTop: 0, padding: "5px 10px", fontSize: 12, background: "#16a34a" }} onClick={() => approveExpense(e.id)}>✅ አፕሩቭ</button>
                              <button className="btn-danger btn-sm" onClick={() => rejectExpense(e.id)}>❌ ሰርዝ</button>
                              <button className="btn btn-sm" style={{ marginTop: 0, padding: "5px 10px", fontSize: 12, background: "#2563eb" }} onClick={() => startEditExpense(e)}>✏️ አርትዕ</button>
                              <button className="btn-danger btn-sm" onClick={() => deleteExpense(e.id)}>🗑️</button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )}
              </div>

              {/* All expenses with status */}
              <div className="card">
                <h3>ሁሉም ወጪዎች</h3>
                <table className="report-table">
                  <thead>
                    <tr><th>Category</th><th>Amount</th><th>Bank</th><th>Date</th><th>Note</th><th>Status</th><th>ድርጊት</th></tr>
                  </thead>
                  <tbody>
                    {db.expenseEntries.length === 0 ? (
                      <tr><td colSpan={7} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ወጪ የለም።</td></tr>
                    ) : (
                      db.expenseEntries.map(e => {
                        const isEditing = editExpenseId === e.id;
                        return isEditing ? (
                          <tr key={e.id} style={{ background: "#eff6ff" }}>
                            <td><input className="leave-input" value={editExpenseForm.category} onChange={ev => setEditExpenseForm(p => ({ ...p, category: ev.target.value }))} style={{ width: 110 }} /></td>
                            <td><input className="leave-input" type="number" value={editExpenseForm.amount} onChange={ev => setEditExpenseForm(p => ({ ...p, amount: ev.target.value }))} style={{ width: 90 }} /></td>
                            <td>
                              <select className="leave-input" value={editExpenseForm.bankId} onChange={ev => setEditExpenseForm(p => ({ ...p, bankId: ev.target.value }))} style={{ width: 110 }}>
                                {(db.bankAccounts || []).map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                              </select>
                            </td>
                            <td><input className="leave-input" type="date" value={editExpenseForm.date} onChange={ev => setEditExpenseForm(p => ({ ...p, date: ev.target.value }))} style={{ width: 120 }} /></td>
                            <td><input className="leave-input" value={editExpenseForm.note} onChange={ev => setEditExpenseForm(p => ({ ...p, note: ev.target.value }))} style={{ width: 110 }} /></td>
                            <td><span className={`status-badge ${e.status === "Approved" ? "approved" : e.status === "Rejected" ? "rejected" : "pending"}`}>{e.status || "Approved"}</span></td>
                            <td>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 10px", fontSize: 12, background: "#16a34a" }} onClick={() => saveEditExpense(e.id)}>💾 አስቀምጥ</button>
                                <button className="btn-danger btn-sm" onClick={() => setEditExpenseId(null)}>✕</button>
                              </div>
                            </td>
                          </tr>
                        ) : (
                          <tr key={e.id}>
                            <td>{e.category}</td>
                            <td>{e.amount.toLocaleString()} Birr</td>
                            <td>{(db.bankAccounts || []).find(b => b.id === Number(e.bankId))?.name || "?"}</td>
                            <td>{e.date}</td>
                            <td>{e.note}</td>
                            <td>
                              <span className={`status-badge ${e.status === "Approved" ? "approved" : e.status === "Rejected" ? "rejected" : "pending"}`}>
                                {e.status || "Approved"}
                              </span>
                            </td>
                            <td>
                              <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                                {e.status === "Pending" && <>
                                  <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => approveExpense(e.id)}>✅</button>
                                  <button className="btn-danger btn-sm" onClick={() => rejectExpense(e.id)}>❌</button>
                                </>}
                                <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#2563eb" }} onClick={() => startEditExpense(e)}>✏️</button>
                                <button className="btn-danger btn-sm" onClick={() => deleteExpense(e.id)}>🗑️</button>
                                {(e.status === "Approved" || !e.status) && (
                                  <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#0f172a" }} onClick={() => printExpenseReceipt(e.id)}>🖨️</button>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* ===== INCOME ===== */}
          {section === "Income Registration" && (
            <section className="panel-grid">
              <IncomeRegistration form={incomeForm} banks={db.bankAccounts || []} onChange={(f, v) => setIncomeForm(p => ({ ...p, [f]: v }))} onSubmit={addIncome} />
              <div className="card">
                <h3>Income Entries</h3>
                <table className="report-table">
                  <thead><tr><th>Bank</th><th>Category</th><th>Amount</th><th>Date</th><th>Source/Notes</th><th>ድርጊት</th></tr></thead>
                  <tbody>
                    {db.incomeEntries.length === 0
                      ? <tr><td colSpan={6} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ገቢ የለም።</td></tr>
                      : db.incomeEntries.map(e => {
                        const isEditing = editIncomeId === e.id;
                        return isEditing ? (
                          <tr key={e.id} style={{ background: "#f0fdf4" }}>
                            <td>
                              <select className="leave-input" value={editIncomeForm.bankId} onChange={ev => setEditIncomeForm(p => ({ ...p, bankId: ev.target.value }))} style={{ width: 110 }}>
                                {(db.bankAccounts || []).map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                              </select>
                            </td>
                            <td><input className="leave-input" value={editIncomeForm.source} onChange={ev => setEditIncomeForm(p => ({ ...p, source: ev.target.value }))} style={{ width: 110 }} /></td>
                            <td><input className="leave-input" type="number" value={editIncomeForm.amount} onChange={ev => setEditIncomeForm(p => ({ ...p, amount: ev.target.value }))} style={{ width: 90 }} /></td>
                            <td><input className="leave-input" type="date" value={editIncomeForm.date} onChange={ev => setEditIncomeForm(p => ({ ...p, date: ev.target.value }))} style={{ width: 120 }} /></td>
                            <td>—</td>
                            <td>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => saveEditIncome(e.id)}>💾</button>
                                <button className="btn-danger btn-sm" onClick={() => setEditIncomeId(null)}>✕</button>
                              </div>
                            </td>
                          </tr>
                        ) : (
                          <tr key={e.id}>
                            <td>{(db.bankAccounts || []).find(b => b.id === Number(e.bankId))?.name || "?"}</td>
                            <td><span className="dept-badge" style={{ background: "#eff6ff", color: "#1d4ed8" }}>{e.category || "—"}</span></td>
                            <td>{e.amount.toLocaleString()} Birr</td>
                            <td>{e.date}</td>
                            <td>{e.source}</td>
                            <td>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#2563eb" }} onClick={() => startEditIncome(e)}>✏️</button>
                                <button className="btn-danger btn-sm" onClick={() => deleteIncome(e.id)}>🗑️</button>
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    }
                  </tbody>
                </table>
              </div>
            </section>
          )}

          {/* ===== EXPENSE ===== */}
          {section === "Expense Registration" && (
            <section className="panel-grid">
              <ExpenseRegistration form={expenseForm} banks={db.bankAccounts || []} onChange={(f, v) => setExpenseForm(p => ({ ...p, [f]: v }))} onSubmit={addExpense} />
              <div className="card">
                <h3>Expense Entries</h3>
                <table className="report-table">
                  <thead>
                    <tr><th>Bank</th><th>Category</th><th>Amount</th><th>Date</th><th>Note</th><th>Status</th><th>ድርጊት</th></tr>
                  </thead>
                  <tbody>
                    {db.expenseEntries.length === 0
                      ? <tr><td colSpan={7} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ወጪ የለም።</td></tr>
                      : db.expenseEntries.map(e => {
                        const isEditing = editExpenseId === e.id;
                        return isEditing ? (
                          <tr key={e.id} style={{ background: "#eff6ff" }}>
                            <td>
                              <select className="leave-input" value={editExpenseForm.bankId} onChange={ev => setEditExpenseForm(p => ({ ...p, bankId: ev.target.value }))} style={{ width: 100 }}>
                                {(db.bankAccounts || []).map(b => <option key={b.id} value={b.id}>{b.name}</option>)}
                              </select>
                            </td>
                            <td><input className="leave-input" value={editExpenseForm.category} onChange={ev => setEditExpenseForm(p => ({ ...p, category: ev.target.value }))} style={{ width: 100 }} /></td>
                            <td><input className="leave-input" type="number" value={editExpenseForm.amount} onChange={ev => setEditExpenseForm(p => ({ ...p, amount: ev.target.value }))} style={{ width: 80 }} /></td>
                            <td><input className="leave-input" type="date" value={editExpenseForm.date} onChange={ev => setEditExpenseForm(p => ({ ...p, date: ev.target.value }))} style={{ width: 120 }} /></td>
                            <td><input className="leave-input" value={editExpenseForm.note} onChange={ev => setEditExpenseForm(p => ({ ...p, note: ev.target.value }))} style={{ width: 100 }} /></td>
                            <td><span className={`status-badge ${e.status === "Approved" ? "approved" : e.status === "Rejected" ? "rejected" : "pending"}`}>{e.status || "Pending"}</span></td>
                            <td>
                              <div style={{ display: "flex", gap: 4 }}>
                                <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => saveEditExpense(e.id)}>💾</button>
                                <button className="btn-danger btn-sm" onClick={() => setEditExpenseId(null)}>✕</button>
                              </div>
                            </td>
                          </tr>
                        ) : (
                          <tr key={e.id}>
                            <td>{(db.bankAccounts || []).find(b => b.id === Number(e.bankId))?.name || "?"}</td>
                            <td>{e.category}</td>
                            <td>{e.amount.toLocaleString()} Birr</td>
                            <td>{e.date}</td>
                            <td>{e.note}</td>
                            <td><span className={`status-badge ${e.status === "Approved" ? "approved" : e.status === "Rejected" ? "rejected" : "pending"}`}>{e.status || "Pending"}</span></td>
                            <td>
                              <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                                <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#2563eb" }} onClick={() => startEditExpense(e)}>✏️</button>
                                <button className="btn-danger btn-sm" onClick={() => deleteExpense(e.id)}>🗑️</button>
                                {(e.status === "Approved" || !e.status) && (
                                  <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#0f172a" }} onClick={() => printExpenseReceipt(e.id)}>🖨️</button>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    }
                  </tbody>
                </table>
              </div>
            </section>
          )
          }

          {/* ===== CREATE USER ===== */}
          {
            section === "Create User" && (
              <section className="settings-section">
                {/* Left: Grant/Revoke portal access */}
                <div className="card">
                  <h3>👤 ሰራተኛ Portal Access</h3>
                  <p style={{ margin: "0 0 14px", color: "#64748b", fontSize: 13 }}>
                    ከ HRM ከተመዘገቡ ሰራተኞች ምረጥ፣ ፓስዋርድ ስጥ — Employee Portal ሊጠቀሙ ይችላሉ።
                  </p>
                  <form className="form-grid" onSubmit={createUserAccount}>
                    <select value={createUserForm.empId} onChange={e => setCreateUserForm(p => ({ ...p, empId: e.target.value }))} required
                      style={{ gridColumn: "1/-1" }}>
                      <option value="">ሰራተኛ ምረጥ (ከ HRM)</option>
                      {(db.employees || []).map(e => (
                        <option key={e.id} value={e.id}>
                          {e.name} — {e.department} {e.hasPortalAccess ? "✅" : ""}
                        </option>
                      ))}
                    </select>
                    <input
                      placeholder="ስልክ ቁጥር (ለ Portal login)"
                      type="tel"
                      value={createUserForm.phone}
                      onChange={e => setCreateUserForm(p => ({ ...p, phone: e.target.value }))}
                    />
                    <input
                      placeholder="ፓስዋርድ (4-6 ቁጥር) *"
                      type="password" inputMode="numeric" maxLength={6}
                      value={createUserForm.password}
                      onChange={e => setCreateUserForm(p => ({ ...p, password: e.target.value }))}
                      required
                    />
                    <select value={createUserForm.department} onChange={e => setCreateUserForm(p => ({ ...p, department: e.target.value }))}
                      style={{ gridColumn: "1/-1" }}>
                      <option value="">ዲፓርትመንት ምረጥ (አማራጭ)</option>
                      {DEPARTMENTS.map(d => <option key={d} value={d}>{d}</option>)}
                    </select>
                    <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>
                      ✅ Access ስጥ
                    </button>
                  </form>

                  <h4 style={{ margin: "20px 0 10px" }}>Portal Access ያላቸው ሰራተኞች</h4>
                  <table className="report-table">
                    <thead><tr><th>ስም</th><th>ዲፓርትመንት</th><th>Access</th><th>ድርጊት</th></tr></thead>
                    <tbody>
                      {(db.employees || []).length === 0
                        ? <tr><td colSpan={4} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ሰራተኛ አልተመዘገበም።</td></tr>
                        : (db.employees || []).map(e => (
                          <tr key={e.id}>
                            <td>{e.name}</td>
                            <td><span className="dept-badge">{e.department}</span></td>
                            <td>
                              {e.hasPortalAccess
                                ? <span className="status-badge approved">✅ Active</span>
                                : <span className="status-badge pending">No Access</span>
                              }
                            </td>
                            <td>
                              {e.hasPortalAccess && (
                                <button className="btn-danger btn-sm" onClick={() => revokeUserAccess(e.id)}>
                                  ❌ ሰርዝ
                                </button>
                              )}
                            </td>
                          </tr>
                        ))
                      }
                    </tbody>
                  </table>
                </div>

                {/* Right: Annual leave days management */}
                <div className="card">
                  <h3>🗓️ አመታዊ ፍቃድ ቀናት አስተዳደር</h3>
                  <p style={{ margin: "0 0 14px", color: "#64748b", fontSize: 13 }}>
                    ለእያንዳንዱ ሰራተኛ አመታዊ ፍቃድ ቀናት ብዛት ይወሰን። Default: 25 ቀናት።
                  </p>
                  <form className="form-grid" onSubmit={setEmployeeLeaveDays}>
                    <select value={leaveDaysForm.empId} onChange={e => setLeaveDaysForm(p => ({ ...p, empId: e.target.value }))} required>
                      <option value="">ሰራተኛ ምረጥ</option>
                      {(db.employees || []).map(e => (
                        <option key={e.id} value={e.id}>{e.name} (አሁን: {e.totalLeaveDays || 25} ቀን)</option>
                      ))}
                    </select>
                    <input
                      type="number" min={1} max={365}
                      placeholder="ቀናት ብዛት"
                      value={leaveDaysForm.days}
                      onChange={e => setLeaveDaysForm(p => ({ ...p, days: e.target.value }))}
                      required
                    />
                    <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>
                      💾 ቀናት አስቀምጥ
                    </button>
                  </form>

                  <h4 style={{ margin: "20px 0 10px" }}>ሰራተኞች — ፍቃድ ባላንስ</h4>
                  <table className="report-table">
                    <thead><tr><th>ስም</th><th>ዲፓርትመንት</th><th>ጠቅላላ</th><th>የወሰዱ</th><th>የቀረ</th></tr></thead>
                    <tbody>
                      {(db.employees || []).length === 0
                        ? <tr><td colSpan={5} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም ሰራተኛ የለም።</td></tr>
                        : (db.employees || []).map(e => {
                          const rem = (e.totalLeaveDays || 25) - (e.usedLeaveDays || 0);
                          return (
                            <tr key={e.id}>
                              <td>{e.name}</td>
                              <td><span className="dept-badge">{e.department}</span></td>
                              <td>{e.totalLeaveDays || 25}</td>
                              <td>{e.usedLeaveDays || 0}</td>
                              <td style={{ color: rem < 3 ? "#ef4444" : "#16a34a", fontWeight: 700 }}>
                                {rem}{rem < 3 ? " ⚠️" : ""}
                              </td>
                            </tr>
                          );
                        })
                      }
                    </tbody>
                  </table>

                  {/* Pending leave requests */}
                  {pendingLeave > 0 && (
                    <>
                      <h4 style={{ margin: "20px 0 10px" }}>🗓️ Pending ፍቃድ ጥያቄዎች <span className="badge">{pendingLeave}</span></h4>
                      <table className="report-table">
                        <thead><tr><th>ሰራተኛ</th><th>የቀረ ፍቃድ</th><th>የጠየቁ</th><th>ይፈቀዱ (ቀናት)</th><th>ጀምሮ</th><th>አይነት</th><th>ምክንያት ሲቀቅ</th><th>ድርጊት</th></tr></thead>
                        <tbody>
                          {(db.leaveRequests || []).filter(r => r.status === "Pending").map(r => {
                            const emp = (db.employees || []).find(e => e.id === r.empId);
                            const rem = emp ? (emp.totalLeaveDays || 25) - (emp.usedLeaveDays || 0) : "—";
                            const adjDays = leaveApproveEdit[r.id] !== undefined ? leaveApproveEdit[r.id] : r.days;
                            const rejectR = leaveRejectReason[r.id] || "";
                            return (
                              <tr key={r.id}>
                                <td>{emp?.name || r.empId}</td>
                                <td style={{ color: typeof rem === "number" && rem < 5 ? "#ef4444" : "#16a34a", fontWeight: 700 }}>{rem} ቀን</td>
                                <td>{r.days} ቀን</td>
                                <td>
                                  <input
                                    type="number" min={1} max={rem}
                                    className="leave-input" style={{ width: 60 }}
                                    value={adjDays}
                                    onChange={ev => setLeaveApproveEdit(p => ({ ...p, [r.id]: ev.target.value }))}
                                  />
                                </td>
                                <td>{r.startDate}</td>
                                <td><span className={`status-badge ${r.leaveType === "emergency" ? "rejected" : "pending"}`}>
                                  {r.leaveType === "emergency" ? "🚨 Emergency" : "🗓️ Annual"}
                                </span></td>
                                <td>
                                  <input
                                    placeholder="ምክንያት..."
                                    className="leave-input" style={{ width: 120 }}
                                    value={rejectR}
                                    onChange={ev => setLeaveRejectReason(p => ({ ...p, [r.id]: ev.target.value }))}
                                  />
                                </td>
                                <td>
                                  <div style={{ display: "flex", gap: 4 }}>
                                    <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => approveLeaveAdjusted(r.id)}>✅ ፍቃድ ስጥ</button>
                                    <button className="btn-danger btn-sm" onClick={() => rejectLeaveWithReason(r.id)}>❌ ሰርዝ</button>
                                    <button className="btn-danger btn-sm" onClick={() => deleteLeaveRequest(r.id)}>🗑️</button>
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </>
                  )}
                </div>
              </section>
            )
          }
          {/* ===== ANNOUNCEMENTS ===== */}
          {
            section === "Announcements" && (
              <section className="panel-grid">
                <div className="card">
                  <h3>📢 አዲስ ማስታወቂያ ላክ</h3>
                  <p style={{ margin: "0 0 12px", color: "#64748b", fontSize: 13 }}>ማስታወቂያ ለሁሉም ሰራተኞች Portal ላይ ይታያቸዋል።</p>
                  <form onSubmit={sendAnnouncement}>
                    <div className="form-grid">
                      <input
                        placeholder="ርዕስ (Title)"
                        value={announcementForm.title}
                        onChange={e => setAnnouncementForm(p => ({ ...p, title: e.target.value }))}
                        required
                        style={{ gridColumn: "1/-1" }}
                      />
                      <textarea
                        placeholder="ዝርዝር መልዕክት..."
                        value={announcementForm.body}
                        onChange={e => setAnnouncementForm(p => ({ ...p, body: e.target.value }))}
                        style={{ gridColumn: "1/-1", minHeight: 100 }}
                        required
                      />
                      <button className="btn" type="submit" style={{ gridColumn: "1/-1", marginTop: 0 }}>
                        📢 ለሁሉም ሰራተኞች ላክ
                      </button>
                    </div>
                  </form>
                </div>
                <div className="card">
                  <h3>ማስታወቂያ ዝርዝር ({(db.announcements || []).length})</h3>
                  {(db.announcements || []).length === 0
                    ? <p style={{ color: "#94a3b8", fontSize: 13 }}>ምንም ማስታወቂያ አልተላከም።</p>
                    : <div className="msg-list">
                      {(db.announcements || []).map(a => (
                        <div key={a.id} className="msg-bubble admin-bubble">
                          <div className="msg-from" style={{ display: "flex", justifyContent: "space-between" }}>
                            <span>📢 {a.title} · {a.date}</span>
                            <span style={{ color: "#16a34a", fontSize: 11 }}>👁 {(a.readBy || []).length} አነበቡ</span>
                          </div>
                          <p style={{ margin: "6px 0 8px" }}>{a.body}</p>
                          <button className="btn-danger btn-sm" onClick={() => deleteAnnouncement(a.id)}>ሰርዝ</button>
                        </div>
                      ))}
                    </div>
                  }
                </div>
              </section>
            )
          }

          {/* ===== MESSAGES (chat thread) ===== */}
          {
            section === "Messages" && (
              <section style={{ display: "grid", gridTemplateColumns: "280px 1fr", gap: 16, minHeight: "75vh" }}>
                {/* Left: employee list */}
                <div className="card" style={{ padding: 16, overflowY: "auto" }}>
                  <h3 style={{ margin: "0 0 12px", fontSize: 15 }}>ሰራተኞች</h3>
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    {(db.employees || []).length === 0
                      ? <p style={{ color: "#94a3b8", fontSize: 13 }}>ምንም ሰራተኛ የለም።</p>
                      : (db.employees || []).map(e => {
                        const unread = (db.messages || []).filter(
                          m => m.from === String(e.id) && m.to === "admin" && !m.readByAdmin
                        ).length;
                        const isActive = msgThreadEmpId === String(e.id);
                        return (
                          <button
                            key={e.id}
                            onClick={() => { setMsgThreadEmpId(String(e.id)); markAdminMsgsRead(); }}
                            style={{
                              textAlign: "left", padding: "10px 12px", borderRadius: 10,
                              border: `1px solid ${isActive ? "#2563eb" : "#e2e8f0"}`,
                              background: isActive ? "#2563eb" : "transparent",
                              color: isActive ? "white" : "inherit",
                              cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center",
                            }}
                          >
                            <div>
                              <div style={{ fontWeight: 600, fontSize: 14 }}>{e.name}</div>
                              <div style={{ fontSize: 11, opacity: 0.75 }}>{e.department}</div>
                            </div>
                            {unread > 0 && <span className="badge">{unread}</span>}
                          </button>
                        );
                      })
                    }
                  </div>

                  {/* Quick send form */}
                  <div style={{ marginTop: 16, borderTop: "1px solid #e2e8f0", paddingTop: 12 }}>
                    <p style={{ fontSize: 11, color: "#94a3b8", margin: "0 0 8px" }}>ቀጥታ ላክ:</p>
                    <form onSubmit={sendAdminMsg}>
                      <select
                        value={msgForm.toId}
                        onChange={e => setMsgForm(p => ({ ...p, toId: e.target.value }))}
                        style={{ width: "100%", padding: "6px 8px", borderRadius: 6, border: "1px solid #cbd5e1", marginBottom: 6, fontSize: 13 }}
                        required
                      >
                        <option value="">ሰራተኛ ምረጥ</option>
                        {(db.employees || []).map(e => <option key={e.id} value={e.id}>{e.name}</option>)}
                      </select>
                      <textarea
                        placeholder="መልዕክት..."
                        value={msgForm.body}
                        onChange={e => setMsgForm(p => ({ ...p, body: e.target.value }))}
                        style={{ width: "100%", padding: "6px 8px", borderRadius: 6, border: "1px solid #cbd5e1", minHeight: 60, fontSize: 13, resize: "vertical" }}
                        required
                      />
                      <button className="btn" type="submit" style={{ width: "100%", marginTop: 6 }}>ላክ →</button>
                    </form>
                  </div>
                </div>

                {/* Right: chat thread */}
                <div className="card" style={{ display: "flex", flexDirection: "column", padding: 16 }}>
                  {!msgThreadEmpId ? (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", flex: 1, color: "#94a3b8", flexDirection: "column", gap: 8 }}>
                      <span style={{ fontSize: 40 }}>💬</span>
                      <span>በግራ በኩል ሰራተኛ ምረጥ — የውይይቱ ዝርዝር ይታያል</span>
                    </div>
                  ) : (() => {
                    const selEmp = (db.employees || []).find(e => String(e.id) === msgThreadEmpId);
                    const thread = [...(db.messages || [])].filter(
                      m => (m.from === msgThreadEmpId && m.to === "admin") ||
                        (m.from === "admin" && m.to === msgThreadEmpId)
                    ).sort((a, b) => a.id - b.id);
                    return (
                      <>
                        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, paddingBottom: 10, borderBottom: "1px solid #e2e8f0" }}>
                          <div>
                            <strong style={{ fontSize: 16 }}>{selEmp?.name}</strong>
                            <span className="dept-badge" style={{ marginLeft: 8 }}>{selEmp?.department}</span>
                          </div>
                        </div>
                        <div style={{ flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: 10, maxHeight: "calc(75vh - 200px)", padding: "4px 0" }}>
                          {thread.length === 0
                            ? <p style={{ color: "#94a3b8", fontSize: 13, textAlign: "center", marginTop: 40 }}>ምንም መልዕክት የለም።</p>
                            : thread.map(m => {
                              const isAdmin = m.from === "admin";
                              return (
                                <div key={m.id} style={{ display: "flex", justifyContent: isAdmin ? "flex-end" : "flex-start", alignItems: "flex-start", gap: 4 }}>
                                  <div className={`msg-bubble ${isAdmin ? "admin-bubble" : "emp-bubble"}`} style={{ maxWidth: "65%" }}>
                                    <div className="msg-from">{isAdmin ? "Admin" : selEmp?.name} · {m.date}</div>
                                    <p>{m.body}</p>
                                  </div>
                                  <button
                                    onClick={() => deleteMessage(m.id)}
                                    title="ሰርዝ"
                                    style={{ background: "none", border: "none", cursor: "pointer", color: "#ef4444", fontSize: 14, padding: "4px", opacity: 0.6, flexShrink: 0, marginTop: 4 }}
                                    onMouseEnter={e => e.currentTarget.style.opacity = "1"}
                                    onMouseLeave={e => e.currentTarget.style.opacity = "0.6"}
                                  >🗑️</button>
                                </div>
                              );
                            })
                          }
                        </div>
                        <form onSubmit={sendThreadReply} style={{ display: "flex", gap: 8, marginTop: 10, borderTop: "1px solid #e2e8f0", paddingTop: 10 }}>
                          <input
                            style={{ flex: 1, padding: "9px 12px", border: "1px solid #cbd5e1", borderRadius: 8, fontSize: 14 }}
                            placeholder="መልዕክት ይጻፉ..."
                            value={threadReply}
                            onChange={e => setThreadReply(e.target.value)}
                          />
                          <button className="btn" type="submit" style={{ marginTop: 0, padding: "9px 18px" }}>ላክ</button>
                        </form>
                      </>
                    );
                  })()}
                </div>
              </section>
            )
          }

          {/* ===== RESIGNATIONS ===== */}
          {
            section === "Resignations" && (
              <section>
                <div className="card">
                  <h3>📝 ስራ መልቀቂያ ጥያቄዎች
                    {pendingResignations > 0 && <span className="badge" style={{ marginLeft: 8 }}>{pendingResignations} Pending</span>}
                  </h3>
                  {(db.resignations || []).length === 0
                    ? <p style={{ color: "#94a3b8" }}>ምንም ስራ መልቀቂያ ጥያቄ አልቀረበም።</p>
                    : (
                      <table className="report-table">
                        <thead>
                          <tr><th>ሰራተኛ</th><th>ዲፓርትመንት</th><th>ምክንያት</th><th>የመጨረሻ የስራ ቀን</th><th>የቀረበ ቀን</th><th>ሁኔታ</th><th>ድርጊት</th></tr>
                        </thead>
                        <tbody>
                          {(db.resignations || []).map(r => (
                            <tr key={r.id}>
                              <td style={{ fontWeight: 600 }}>{r.empName}</td>
                              <td><span className="dept-badge">{r.department}</span></td>
                              <td>{r.reason}</td>
                              <td style={{ fontWeight: 700, color: "#ef4444" }}>{r.lastDay}</td>
                              <td>{r.submittedDate}</td>
                              <td>
                                <span className={`status-badge ${r.status === "Acknowledged" ? "approved" : "pending"}`}>
                                  {r.status === "Acknowledged" ? "✅ ተቀበለ" : "⏳ Pending"}
                                </span>
                              </td>
                              <td>
                                <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                                  {r.status === "Pending" && (
                                    <button
                                      className="btn btn-sm"
                                      style={{ marginTop: 0, padding: "5px 12px", fontSize: 12, background: "#16a34a" }}
                                      onClick={() => acknowledgeResignation(r.id)}
                                    >
                                      ✅ ተቀበለ
                                    </button>
                                  )}
                                  <button className="btn-danger btn-sm" onClick={() => deleteResignation(r.id)}>🗑️ ሰርዝ</button>
                                </div>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )
                  }
                </div>
              </section>
            )
          }

          {/* ===== REPORTS — Income Statement ===== */}
          {
            section === "Reports" && (() => {
              const now = new Date();
              const yr = now.getFullYear(); const mo = now.getMonth();
              const etYear = yr - 7;
              const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

              // Custom date range filter
              const filterByRange = (entries) => {
                const f = reportDateFrom ? new Date(reportDateFrom) : null;
                const t = reportDateTo ? new Date(reportDateTo) : null;
                return entries.filter(e => {
                  const d = new Date(e.date);
                  if (f && d < f) return false;
                  if (t && d > t) return false;
                  return true;
                });
              };

              const monthlyInc = db.incomeEntries.filter(e => { const d = new Date(e.date); return d.getFullYear() === yr && d.getMonth() === mo; });
              const monthlyExp = db.expenseEntries.filter(e => { const d = new Date(e.date); return d.getFullYear() === yr && d.getMonth() === mo; });

              const mData = buildStatementData(monthlyInc, monthlyExp);
              const aData = buildStatementData(db.incomeEntries, db.expenseEntries);
              const rData = buildStatementData(filterByRange(db.incomeEntries), filterByRange(db.expenseEntries));

              const rangeLabel = reportDateFrom && reportDateTo
                ? `${reportDateFrom} → ${reportDateTo}`
                : reportDateFrom ? `${reportDateFrom} →` : reportDateTo ? `→ ${reportDateTo}` : "ሁሉም ቀናት";

              const fmt = (n) => Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

              const IncomeStatement = ({ data, periodLabel, subtitleLabel }) => (
                <div className="income-statement-wrap">
                  {/* org logo if available */}
                  {db.organizationLogo && (
                    <div style={{ textAlign: "center", marginBottom: 10 }}>
                      <img src={db.organizationLogo} alt="logo" style={{ height: 60, objectFit: "contain" }} />
                    </div>
                  )}
                  <div className="is-header">
                    <div className="is-org">[{db.organizationLogoText || db.organizationName}]</div>
                    <div className="is-title">የገቢና ወጪ መግለጫ (Income Statement)</div>
                    <div className="is-subtitle">{subtitleLabel}</div>
                  </div>

                  <table className="is-table">
                    <thead>
                      <tr>
                        <th style={{ width: "58%" }}>የሂሳብ መይብ (Account Description)</th>
                        <th style={{ width: "21%", textAlign: "right" }}>ነጉስ ድምር (ETB)</th>
                        <th style={{ width: "21%", textAlign: "right" }}>ዋና ድምር (ETB)</th>
                      </tr>
                    </thead>
                    <tbody>

                      {/* ── Revenue ── */}
                      <tr className="is-row">
                        <td>ከሽያጭ ወይም የአጠቃላዮት ገቢ (Revenue)</td>
                        <td></td>
                        <td className="is-amt">{fmt(data.revenue)}</td>
                      </tr>
                      <tr className="is-row">
                        <td>ልቅነስ: የሽያጭ ዕቃዎች ዋጋ (Cost of Goods Sold – COGS)</td>
                        <td className="is-amt">{fmt(data.cogs)}</td>
                        <td></td>
                      </tr>
                      <tr className="is-row is-bold is-gross">
                        <td>አጠቃላይ ትርፍ (Gross Profit)</td>
                        <td></td>
                        <td className="is-amt">{fmt(data.grossProfit)}</td>
                      </tr>

                      <tr className="is-spacer"><td colSpan={3}></td></tr>

                      {/* ── Operating Expenses header ── */}
                      <tr className="is-section-header">
                        <td colSpan={3}>የሥራ ማስኬጃ ወጪዎች (Operating Expenses):</td>
                      </tr>

                      {/* Fixed rows — always shown exactly as in the image */}
                      <tr className="is-row is-indent">
                        <td>የሰራተኞች ደሞዝ (Salaries &amp; Wages)</td>
                        <td className="is-amt">{fmt(data.salaries)}</td>
                        <td></td>
                      </tr>
                      <tr className="is-row is-indent">
                        <td>የቢሮ ወይም የሱቅ ኪራይ (Rent Expense)</td>
                        <td className="is-amt">{fmt(data.rent)}</td>
                        <td></td>
                      </tr>
                      <tr className="is-row is-indent">
                        <td>ውሃ፣ መብራትና ስልክ (Utilities)</td>
                        <td className="is-amt">{fmt(data.utilities)}</td>
                        <td></td>
                      </tr>
                      <tr className="is-row is-indent">
                        <td>የትርንስፖርትና ሎጅስቲክስ (Transport)</td>
                        <td className="is-amt">{fmt(data.transport)}</td>
                        <td></td>
                      </tr>
                      <tr className="is-row is-indent">
                        <td>ልዩ ልዩ ወጪዎች (Miscellaneous)</td>
                        <td className="is-amt">{fmt(data.misc)}</td>
                        <td></td>
                      </tr>

                      {/* Dynamic "other" expense categories not in the fixed list */}
                      {data.otherRows.map(({ cat, amt }) => (
                        <tr key={cat} className="is-row is-indent">
                          <td>{cat}</td>
                          <td className="is-amt">{fmt(amt)}</td>
                          <td></td>
                        </tr>
                      ))}

                      {/* Total OpEx */}
                      <tr className="is-row is-bold is-totalopex">
                        <td>ጠቅላላ የሥራ ማስኬጃ ወጪ (Total Operating Expenses)</td>
                        <td></td>
                        <td className="is-amt">{fmt(data.totalOpEx)}</td>
                      </tr>

                      <tr className="is-spacer"><td colSpan={3}></td></tr>

                      {/* ── Bottom lines ── */}
                      <tr className="is-row is-bold">
                        <td>የተጣራ የንግድ ትርፍ (Net Operating Income)</td>
                        <td></td>
                        <td className="is-amt" style={{ color: data.netOperatingIncome >= 0 ? "#16a34a" : "#ef4444" }}>
                          {fmt(data.netOperatingIncome)}
                        </td>
                      </tr>
                      <tr className="is-row">
                        <td>ልቅነስ: የገቢ ግብር / ታክስ (Income Tax Expense — 30%)</td>
                        <td></td>
                        <td className="is-amt">{fmt(data.tax)}</td>
                      </tr>
                      <tr className="is-row is-netprofit">
                        <td>የዓመቱ የተጣራ ትርፍ (Net Profit for the Year)</td>
                        <td></td>
                        <td className="is-amt">{fmt(data.netProfit)}</td>
                      </tr>

                    </tbody>
                  </table>

                  <button
                    className="btn"
                    style={{ marginTop: 18 }}
                    onClick={() => exportIncomeStatement(periodLabel, data)}
                  >
                    📥 Export to Excel
                  </button>
                </div>
              );

              return (
                <section>
                  {/* ── Tab buttons ── */}
                  <div className="report-tabs" style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap", alignItems: "center" }}>
                    {[
                      ["monthly", `📅 ${monthNames[mo]} ${yr}`],
                      ["annual", `📊 አመታዊ ${yr}`],
                      ["custom", "📆 Custom Range"],
                    ].map(([k, l]) => (
                      <button key={k} className={`hrm-tab ${reportTab === k ? "active" : ""}`} onClick={() => setReportTab(k)}>{l}</button>
                    ))}
                  </div>

                  {/* ── Custom date range picker ── */}
                  {reportTab === "custom" && (
                    <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap", marginBottom: 20, padding: "14px 18px", background: "#f8fafc", borderRadius: 12, border: "1px solid #e2e8f0" }}>
                      <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                        <label style={{ fontSize: 12, fontWeight: 700, color: "#64748b" }}>ከ (From)</label>
                        <input type="date" value={reportDateFrom} onChange={e => setReportDateFrom(e.target.value)}
                          style={{ padding: "8px 12px", border: "1px solid #cbd5e1", borderRadius: 8, fontSize: 14 }} />
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                        <label style={{ fontSize: 12, fontWeight: 700, color: "#64748b" }}>እስከ (To)</label>
                        <input type="date" value={reportDateTo} onChange={e => setReportDateTo(e.target.value)}
                          style={{ padding: "8px 12px", border: "1px solid #cbd5e1", borderRadius: 8, fontSize: 14 }} />
                      </div>
                      <button className="secondary-btn" style={{ marginTop: 0 }}
                        onClick={() => { setReportDateFrom(""); setReportDateTo(""); }}>✕ Reset</button>
                      {(reportDateFrom || reportDateTo) && (
                        <div style={{ fontSize: 13, color: "#2563eb", fontWeight: 700, alignSelf: "flex-end", paddingBottom: 6 }}>
                          {filterByRange(db.incomeEntries).length} ገቢ · {filterByRange(db.expenseEntries).length} ወጪ records
                        </div>
                      )}
                    </div>
                  )}

                  {reportTab === "monthly" && (
                    <IncomeStatement
                      data={mData}
                      periodLabel={`${monthNames[mo]}_${yr}`}
                      subtitleLabel={`ለ${etYear} አ.ም. — For the Month of ${monthNames[mo]} ${yr}`}
                    />
                  )}
                  {reportTab === "annual" && (
                    <IncomeStatement
                      data={aData}
                      periodLabel={`Annual_${yr}`}
                      subtitleLabel={`ለ${etYear} አ.ም. — For the Year Ended ${yr}`}
                    />
                  )}
                  {reportTab === "custom" && (
                    <IncomeStatement
                      data={rData}
                      periodLabel={`Custom_${rangeLabel.replace(/[\s→]/g, "_")}`}
                      subtitleLabel={`Custom Range — ${rangeLabel}`}
                    />
                  )}
                </section>
              );
            })()
          }

          {/* ===== BALANCE SHEET — የሃብትና እዳ ምዝገባ ===== */}
          {section === "Balance Sheet" && (() => {
            const bs = db.balanceSheet || DEFAULT.balanceSheet;
            const n = (v) => Number(v) || 0;
            const fmt = (v) => Number(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

            // Computed totals
            const totalCurrentAssets = n(bs.cashAndBank) + n(bs.accountsReceivable) + n(bs.inventory);
            const totalNonCurrentAssets = n(bs.equipmentAndVehicles) - n(bs.accumulatedDepreciation);
            const totalAssets = totalCurrentAssets + totalNonCurrentAssets;
            const totalCurrentLiab = n(bs.accountsPayable) + n(bs.taxPayable);

            // Pull net profit from income statement
            const aData = buildStatementData(db.incomeEntries, db.expenseEntries);
            const netProfit = aData.netProfit;

            const totalEquity = n(bs.capital) + n(bs.retainedEarnings) + netProfit;
            const totalLiabEquity = totalCurrentLiab + totalEquity;

            const etYear = new Date().getFullYear() - 7;

            const Field = ({ label, fkey, wide }) => (
              <div className={`bs-field${wide ? " bs-field-wide" : ""}`}>
                <label className="bs-field-label">{label}</label>
                <input
                  type="number"
                  className="bs-field-input"
                  placeholder="0.00"
                  value={bsForm[fkey] ?? ""}
                  onChange={e => setBsForm(p => ({ ...p, [fkey]: e.target.value }))}
                />
              </div>
            );

            return (
              <section>
                {/* ── FORM (edit mode) ── */}
                {bsEditing ? (
                  <div className="bs-form-card">
                    <div className="bs-form-header">
                      <h2 style={{ margin: 0 }}>📋 Balance Sheet ምዝገባ</h2>
                      <div style={{ display: "flex", gap: 10 }}>
                        <button className="btn" style={{ marginTop: 0 }} onClick={saveBalanceSheet}>💾 አስቀምጥ</button>
                        <button className="secondary-btn" style={{ marginTop: 0 }} onClick={() => setBsEditing(false)}>✕ ሰርዝ</button>
                      </div>
                    </div>

                    <div style={{ marginBottom: 18 }}>
                      <label className="bs-field-label">📅 የሪፖርቱ ቀን (As of Date)</label>
                      <input type="date" className="bs-field-input" style={{ maxWidth: 220 }}
                        value={bsForm.asOfDate} onChange={e => setBsForm(p => ({ ...p, asOfDate: e.target.value }))} />
                    </div>

                    <div className="bs-form-grid">
                      {/* LEFT — Assets */}
                      <div>
                        <div className="bs-section-title">ሃብት (Assets)</div>

                        <div className="bs-sub-title">የአሁኑ ጊዜ ሃብት (Current Assets)</div>
                        <Field label="ጥሬ ናንያ ወይም ሃ/ጊዜ ገንዘብ — Cash & Bank" fkey="cashAndBank" />
                        <Field label="ያልቀፈሉ ደረሰኝ ዕዳ — Accounts Receivable" fkey="accountsReceivable" />
                        <Field label="የዕቃ ክምችት — Inventory" fkey="inventory" />

                        <div className="bs-sub-title" style={{ marginTop: 18 }}>ቋሚ ጊዜ ሃብት (Non-Current Assets)</div>
                        <Field label="መሳሪያና ተሸከርካሪ — Equipment & Vehicles" fkey="equipmentAndVehicles" />
                        <Field label="ልቀሰ: የተጠራቀመ እርጅና — Accumulated Depreciation" fkey="accumulatedDepreciation" />
                      </div>

                      {/* RIGHT — Liabilities & Equity */}
                      <div>
                        <div className="bs-section-title">ዕዳና ካፒታል (Liabilities &amp; Equity)</div>

                        <div className="bs-sub-title">የአሁኑ ጊዜ ዕዳ (Current Liabilities)</div>
                        <Field label="ሌሎቻቸሞች ያልቀፈሉት — Accounts Payable" fkey="accountsPayable" />
                        <Field label="ያልቀፈሉ ግብር/ታክስ — Tax Payable" fkey="taxPayable" />

                        <div className="bs-sub-title" style={{ marginTop: 18 }}>የባለቤቱ ካፒታል (Owner's Equity)</div>
                        <Field label="የመጀመሪያ ሙሉ ካፒታል — Capital" fkey="capital" />
                        <Field label="የተጠራቀሙ ትርፍ — Retained Earnings" fkey="retainedEarnings" />
                        <div className="bs-auto-row">
                          <span>የዓመቱ የተጣራ ትርፍ — Net Profit (auto)</span>
                          <strong style={{ color: netProfit >= 0 ? "#16a34a" : "#ef4444" }}>{fmt(netProfit)} ETB</strong>
                        </div>
                      </div>
                    </div>
                  </div>
                ) : (
                  /* ── DISPLAY MODE — Balance Sheet table matching the image ── */
                  <div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
                      <div>
                        <button className="btn" style={{ marginTop: 0 }} onClick={startEditBs}>✏️ ዝርዝር ምዝግብ / አርትዕ</button>
                        <button className="secondary-btn" style={{ marginTop: 0, marginLeft: 10 }} onClick={exportBalanceSheet}>📥 Export Excel</button>
                      </div>
                    </div>

                    <div className="bs-display-wrap">
                      {/* ── Header ── */}
                      <div className="bs-display-header">
                        {db.organizationLogo && <img src={db.organizationLogo} alt="logo" style={{ height: 52, objectFit: "contain", marginBottom: 6 }} />}
                        <div className="bs-display-org">[{db.organizationLogoText || db.organizationName}]</div>
                        <div className="bs-display-title">የሃብትና እዳ መግለጫ (Balance Sheet)</div>
                        <div className="bs-display-subtitle">
                          አ.አ. ሐምሌ 30/{etYear} ቀን አንፃር ({bs.asOfDate ? `As of ${bs.asOfDate}` : "As of —"})
                        </div>
                      </div>

                      {/* ── Two-column table ── */}
                      <div className="bs-cols">

                        {/* LEFT — Assets */}
                        <table className="bs-table">
                          <thead>
                            <tr>
                              <th>ሃብት (Assets)</th>
                              <th>ብር (ETB)</th>
                            </tr>
                          </thead>
                          <tbody>
                            {/* Current Assets */}
                            <tr className="bs-section-row"><td colSpan={2}>የአሁኑ ጊዜ ሃብት (Current Assets)</td></tr>
                            <tr className="bs-data-row">
                              <td>ጥሬ ናንያ ወይም ሃ/ጊዜ ገንዘብ (Cash &amp; Bank)</td>
                              <td className="bs-amt">{fmt(bs.cashAndBank || 0)}</td>
                            </tr>
                            <tr className="bs-data-row">
                              <td>ያልቀፈሉ ደረሰኝ ዕዳ (Accounts Receivable)</td>
                              <td className="bs-amt">{fmt(bs.accountsReceivable || 0)}</td>
                            </tr>
                            <tr className="bs-data-row">
                              <td>የዕቃ ክምችት (Inventory)</td>
                              <td className="bs-amt">{fmt(bs.inventory || 0)}</td>
                            </tr>
                            <tr className="bs-total-row">
                              <td>ጠቅላላ የአሁኑ ጊዜ ሃብት</td>
                              <td className="bs-amt">{fmt(totalCurrentAssets)}</td>
                            </tr>

                            <tr className="bs-spacer"><td colSpan={2}></td></tr>

                            {/* Non-Current Assets */}
                            <tr className="bs-section-row"><td colSpan={2}>ቋሚ ጊዜ ሃብት (Non-Current Assets)</td></tr>
                            <tr className="bs-data-row">
                              <td>መሳሪያና ተሸከርካሪ (Equipment &amp; Vehicles)</td>
                              <td className="bs-amt">{fmt(bs.equipmentAndVehicles || 0)}</td>
                            </tr>
                            <tr className="bs-data-row">
                              <td>ልቀሰ: የተጠራቀመ እርጅና (Acc. Depreciation)</td>
                              <td className="bs-amt" style={{ color: "#ef4444" }}>({fmt(bs.accumulatedDepreciation || 0)})</td>
                            </tr>
                            <tr className="bs-total-row">
                              <td>ጠቅላላ ቋሚ ጊዜ ሃብት</td>
                              <td className="bs-amt">{fmt(totalNonCurrentAssets)}</td>
                            </tr>

                            <tr className="bs-spacer"><td colSpan={2}></td></tr>
                            <tr className="bs-grand-total-row">
                              <td>ጠቅላላ ሃብት (Total Assets)</td>
                              <td className="bs-amt">{fmt(totalAssets)}</td>
                            </tr>
                          </tbody>
                        </table>

                        {/* RIGHT — Liabilities & Equity */}
                        <table className="bs-table">
                          <thead>
                            <tr>
                              <th>ዕዳ ካፒታል (Liabilities &amp; Equity)</th>
                              <th>ብር (ETB)</th>
                            </tr>
                          </thead>
                          <tbody>
                            {/* Current Liabilities */}
                            <tr className="bs-section-row"><td colSpan={2}>የአሁኑ ጊዜ ዕዳ (Current Liabilities)</td></tr>
                            <tr className="bs-data-row">
                              <td>ሌሎቻቸሞች ያልቀፈሉት (Accounts Payable)</td>
                              <td className="bs-amt">{fmt(bs.accountsPayable || 0)}</td>
                            </tr>
                            <tr className="bs-data-row">
                              <td>ያልቀፈሉ ግብር/ታክስ (Tax Payable)</td>
                              <td className="bs-amt">{fmt(bs.taxPayable || 0)}</td>
                            </tr>
                            <tr className="bs-total-row">
                              <td>ጠቅላላ የአሁኑ ጊዜ ዕዳ</td>
                              <td className="bs-amt">{fmt(totalCurrentLiab)}</td>
                            </tr>

                            <tr className="bs-spacer"><td colSpan={2}></td></tr>

                            {/* Owner's Equity */}
                            <tr className="bs-section-row"><td colSpan={2}>የባለቤቱ ካፒታል (Owner's Equity)</td></tr>
                            <tr className="bs-data-row">
                              <td>የመጀመሪያ ሙሉ ካፒታል (Capital)</td>
                              <td className="bs-amt">{fmt(bs.capital || 0)}</td>
                            </tr>
                            <tr className="bs-data-row">
                              <td>የተጠራቀሙ ትርፍ (Retained Earnings)</td>
                              <td className="bs-amt">{fmt(bs.retainedEarnings || 0)}</td>
                            </tr>
                            <tr className="bs-data-row">
                              <td>የዓመቱ የተጣራ ትርፍ (Net Profit)</td>
                              <td className="bs-amt" style={{ color: netProfit >= 0 ? "#16a34a" : "#ef4444", fontWeight: 700 }}>{fmt(netProfit)}</td>
                            </tr>
                            <tr className="bs-total-row">
                              <td>ጠቅላላ የባለቤቱ ካፒታል</td>
                              <td className="bs-amt">{fmt(totalEquity)}</td>
                            </tr>

                            <tr className="bs-spacer"><td colSpan={2}></td></tr>
                            <tr className="bs-grand-total-row">
                              <td>ጠቅላላ ዕዳና ካፒታል (Total Liabilities &amp; Equity)</td>
                              <td className="bs-amt">{fmt(totalLiabEquity)}</td>
                            </tr>

                            <tr className="bs-spacer"><td colSpan={2}></td></tr>

                            {/* Balance check */}
                            <tr style={{ background: Math.abs(totalAssets - totalLiabEquity) < 1 ? "#dcfce7" : "#fee2e2" }}>
                              <td colSpan={2} style={{ textAlign: "center", fontWeight: 700, fontSize: 13, padding: "8px 14px" }}>
                                {Math.abs(totalAssets - totalLiabEquity) < 1
                                  ? "✅ Balance Sheet ተሟሟቷል — Assets = Liabilities + Equity"
                                  : `⚠️ ልዩነት: ${fmt(Math.abs(totalAssets - totalLiabEquity))} ETB — ዝርዝሮቹን ያረጋግጡ`
                                }
                              </td>
                            </tr>
                          </tbody>
                        </table>

                      </div>{/* /bs-cols */}
                    </div>{/* /bs-display-wrap */}
                  </div>
                )}
              </section>
            );
          })()}

          {/* ===== DEBT REGISTRATION (Finance) ===== */}
          {
            section === "Debt Registration" && (
              <section className="panel-grid">
                <div className="card">
                  <h3>አዲስ እዳ ምዝገባ</h3>
                  <div className="form-grid" style={{ marginBottom: 12 }}>
                    <input placeholder="Organization" value={debtForm.organization} onChange={e => setDebtForm(p => ({ ...p, organization: e.target.value }))} />
                    <input type="number" placeholder="Total Debt" value={debtForm.total} onChange={e => setDebtForm(p => ({ ...p, total: e.target.value }))} />
                    <input type="number" placeholder="Paid Amount" value={debtForm.paid} onChange={e => setDebtForm(p => ({ ...p, paid: e.target.value }))} />
                    <select value={debtForm.status} onChange={e => setDebtForm(p => ({ ...p, status: e.target.value }))}>
                      {["Pending", "Partial", "Paid"].map(s => <option key={s}>{s}</option>)}
                    </select>
                    <input type="date" value={debtForm.date} onChange={e => setDebtForm(p => ({ ...p, date: e.target.value }))} />
                  </div>
                  <button className="btn" onClick={addDebt}>Save Debt</button>
                </div>
                <div className="card">
                  <h3>የእዳ ዝርዝር</h3>
                  <table className="report-table">
                    <thead><tr><th>Organization</th><th>Total</th><th>Paid</th><th>Remaining</th><th>Status</th><th>Date</th><th>{t.actions}</th></tr></thead>
                    <tbody>
                      {db.debts.length === 0
                        ? <tr><td colSpan={7} style={{ textAlign: "center", color: "#94a3b8" }}>ምንም እዳ የለም።</td></tr>
                        : db.debts.map(d => {
                          const isEditing = editDebtId === d.id;
                          return isEditing ? (
                            <tr key={d.id} style={{ background: "#fef3c7" }}>
                              <td><input className="leave-input" value={editDebtForm.organization} onChange={ev => setEditDebtForm(p => ({ ...p, organization: ev.target.value }))} style={{ width: 120 }} /></td>
                              <td><input className="leave-input" type="number" value={editDebtForm.total} onChange={ev => setEditDebtForm(p => ({ ...p, total: ev.target.value }))} style={{ width: 80 }} /></td>
                              <td><input className="leave-input" type="number" value={editDebtForm.paid} onChange={ev => setEditDebtForm(p => ({ ...p, paid: ev.target.value }))} style={{ width: 80 }} /></td>
                              <td style={{ color: "#ef4444", fontWeight: 700 }}>{(Number(editDebtForm.total || 0) - Number(editDebtForm.paid || 0)).toLocaleString()}</td>
                              <td>
                                <select className="leave-input" value={editDebtForm.status} onChange={ev => setEditDebtForm(p => ({ ...p, status: ev.target.value }))} style={{ width: 90 }}>
                                  {["Pending", "Partial", "Paid"].map(s => <option key={s}>{s}</option>)}
                                </select>
                              </td>
                              <td><input className="leave-input" type="date" value={editDebtForm.date} onChange={ev => setEditDebtForm(p => ({ ...p, date: ev.target.value }))} style={{ width: 120 }} /></td>
                              <td>
                                <div style={{ display: "flex", gap: 4 }}>
                                  <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => saveEditDebt(d.id)}>💾</button>
                                  <button className="btn-danger btn-sm" onClick={() => setEditDebtId(null)}>✕</button>
                                </div>
                              </td>
                            </tr>
                          ) : (
                            <tr key={d.id}>
                              <td>{d.organization}</td>
                              <td>{d.total.toLocaleString()}</td>
                              <td style={{ color: "#16a34a" }}>{d.paid.toLocaleString()}</td>
                              <td style={{ color: "#ef4444", fontWeight: 700 }}>{(d.total - d.paid).toLocaleString()}</td>
                              <td><span className={`status-badge ${d.status === "Paid" ? "approved" : d.status === "Partial" ? "pending" : "rejected"}`}>{d.status}</span></td>
                              <td>{d.date}</td>
                              <td>
                                <div style={{ display: "flex", gap: 4 }}>
                                  <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#2563eb" }} onClick={() => startEditDebt(d)}>✏️</button>
                                  <button className="btn-danger btn-sm" onClick={() => deleteDebt(d.id)}>🗑️</button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      }
                    </tbody>
                  </table>
                </div>
              </section>
            )
          }

          {/* ===== PAYROLL (Finance) ===== */}
          {
            section === "Payroll" && (
              <section>
                <div className="panel-grid" style={{ marginBottom: 18 }}>
                  <div className="card">
                    <h3>ደሞዝ ከፍያ (Payroll)</h3>
                    <p style={{ color: "#64748b", fontSize: 13, margin: "0 0 12px" }}>ሁሉም ሰራተኞች ደሞዝ ሲደን ከዚህ ባንክ ወጪ ይደረጋል።</p>
                    <div className="form-grid" style={{ marginBottom: 12 }}>
                      <select value={payrollBankId} onChange={e => setPayrollBankId(e.target.value)}>
                        <option value="">ባንክ ምረጥ</option>
                        {(db.bankAccounts || []).map(b => <option key={b.id} value={b.id}>{b.name} ({b.accountNumber})</option>)}
                      </select>
                      <input type="date" value={payrollDate} onChange={e => setPayrollDate(e.target.value)} />
                    </div>
                    {(db.employees || []).length > 0 && (
                      <div className="payroll-totals">
                        <div className="payroll-total-row"><span>ሰራተኞች</span><strong>{(db.employees || []).length}</strong></div>
                        <div className="payroll-total-row"><span>ጠቅላላ ደሞዝ</span><strong>{(db.employees || []).reduce((s, e) => s + e.basicSalary, 0).toLocaleString()} Birr</strong></div>
                        <div className="payroll-total-row pension-row"><span>ጡረታ</span><strong>-{(db.employees || []).reduce((s, e) => s + (e.pension || 0), 0).toFixed(0)} Birr</strong></div>
                        <div className="payroll-total-row net-row"><span>ጠቅላላ የሚከፈል</span><strong>{(db.employees || []).reduce((s, e) => s + e.netPay, 0).toFixed(0)} Birr</strong></div>
                      </div>
                    )}
                    <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
                      <button className="btn" style={{ flex: 1 }} onClick={payAllSalaries}>✅ ደን — ወጪ ምዝግብ</button>
                      <button className="secondary-btn" style={{ flex: 1 }} onClick={exportPayroll}>📥 Excel Export</button>
                    </div>
                  </div>
                  <div className="card">
                    <h3>Payroll Preview</h3>
                    <table className="report-table">
                      <thead><tr><th>ተ.ቁ</th><th>ስም</th><th>ዲፓርትመንት</th><th>ስልክ</th><th>አካውንት</th><th>ደሞዝ</th></tr></thead>
                      <tbody>{(db.employees || []).map((e, i) => (
                        <tr key={e.id}>
                          <td>{i + 1}</td><td>{e.name}</td><td><span className="dept-badge">{e.department}</span></td>
                          <td>{e.phone || "—"}</td><td>{e.bankAccount || "—"}</td>
                          <td style={{ color: "#16a34a", fontWeight: 700 }}>{e.netPay.toFixed(0)} Birr</td>
                        </tr>
                      ))}</tbody>
                    </table>
                  </div>
                </div>

                {/* Per-employee payroll edit */}
                <div className="card">
                  <h3>💼 ፔሮል ማስተካከያ — ዲውቲ፣ ጡረታ፣ ደሞዝ</h3>
                  <p style={{ color: "#64748b", fontSize: 13, margin: "0 0 12px" }}>ለእያንዳዱ ሰራተኛ ዲውቲ ማከል፣ ጡረታ opt-in/out፣ ወይም ደሞዝ ማስተካከል።</p>
                  <table className="report-table">
                    <thead>
                      <tr>
                        <th>ስም</th><th>ዲፓርትመንት</th>
                        <th>ደሞዝ</th><th>ዲውቲ</th>
                        <th>7% ጡረታ</th><th>የሚከፈል</th><th>አርትዕ</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(db.employees || []).map(e => {
                        const ed = payrollEdit[e.id] || {};
                        const editing = e.id in payrollEdit;
                        const curSal = editing ? (ed.basicSalary ?? e.basicSalary) : e.basicSalary;
                        const curDuty = editing ? (ed.duty ?? (e.duty || 0)) : (e.duty || 0);
                        const curPen = editing ? (ed.pensionOpt ?? e.pensionOpt) : e.pensionOpt;
                        const previewPen = curPen ? (Number(curSal) * 0.07).toFixed(0) : 0;
                        const previewNet = (Number(curSal) * (curPen ? 0.93 : 1) + Number(curDuty)).toFixed(0);
                        return (
                          <tr key={e.id}>
                            <td>{e.name}</td>
                            <td><span className="dept-badge">{e.department}</span></td>
                            <td>
                              {editing
                                ? <input type="number" className="leave-input" style={{ width: 80 }} value={ed.basicSalary ?? e.basicSalary} onChange={ev => setPayrollEdit(p => ({ ...p, [e.id]: { ...p[e.id], basicSalary: ev.target.value } }))} />
                                : `${e.basicSalary.toLocaleString()}`
                              }
                            </td>
                            <td>
                              {editing
                                ? <input type="number" className="leave-input" style={{ width: 70 }} value={ed.duty ?? (e.duty || 0)} onChange={ev => setPayrollEdit(p => ({ ...p, [e.id]: { ...p[e.id], duty: ev.target.value } }))} />
                                : `${(e.duty || 0).toLocaleString()}`
                              }
                            </td>
                            <td>
                              {editing
                                ? <label style={{ display: "flex", alignItems: "center", gap: 4 }}>
                                  <input type="checkbox" checked={ed.pensionOpt ?? e.pensionOpt} onChange={ev => setPayrollEdit(p => ({ ...p, [e.id]: { ...p[e.id], pensionOpt: ev.target.checked } }))} />
                                  {curPen ? <span style={{ color: "#ef4444" }}>-{previewPen}</span> : "—"}
                                </label>
                                : (e.pension > 0 ? <span style={{ color: "#ef4444" }}>-{e.pension.toFixed(0)}</span> : "—")
                              }
                            </td>
                            <td style={{ color: "#16a34a", fontWeight: 700 }}>
                              {editing ? `${previewNet} Birr` : `${e.netPay.toFixed(0)} Birr`}
                            </td>
                            <td>
                              {!editing
                                ? <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 10px", fontSize: 12 }} onClick={() => setPayrollEdit(p => ({ ...p, [e.id]: {} }))}>✏️</button>
                                : <div style={{ display: "flex", gap: 4 }}>
                                  <button className="btn btn-sm" style={{ marginTop: 0, padding: "4px 8px", fontSize: 12, background: "#16a34a" }} onClick={() => savePayrollEdit(e.id)}>💾</button>
                                  <button className="btn-danger btn-sm" onClick={() => setPayrollEdit(p => { const n = { ...p }; delete n[e.id]; return n; })}>✕</button>
                                </div>
                              }
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </section>
            )
          }

          {/* ===== SETTINGS ===== */}
          {
            section === "Settings" && (
              <section className="settings-section">
                <div className="card">
                  <h3>🏦 Bank Accounts</h3>
                  <form className="form-grid" style={{ marginBottom: 20 }} onSubmit={addBank}>
                    <input placeholder="Bank Name" value={bankForm.name} onChange={e => setBankForm(p => ({ ...p, name: e.target.value }))} required />
                    <input placeholder="Account Number" value={bankForm.accountNumber} onChange={e => setBankForm(p => ({ ...p, accountNumber: e.target.value }))} required />
                    <input type="number" placeholder="Initial Balance" value={bankForm.initialBalance} onChange={e => setBankForm(p => ({ ...p, initialBalance: e.target.value }))} />
                    <button className="btn" type="submit" style={{ marginTop: 0 }}>+ Open Bank</button>
                  </form>
                  <table><thead><tr><th>Bank Name</th><th>Account #</th><th>Current Balance</th><th>Action</th></tr></thead><tbody>
                    {bankSummaries.map(b => (
                      <tr key={b.id}>
                        <td>{b.name}</td><td>{b.accountNumber}</td><td>{b.balance.toLocaleString()} Birr</td>
                        <td><button className="btn-danger btn-sm" onClick={() => deleteBank(b.id)}>Delete</button></td>
                      </tr>
                    ))}
                    {bankSummaries.length === 0 && <tr><td colSpan={4} style={{ color: "#94a3b8", textAlign: "center" }}>No banks yet.</td></tr>}
                  </tbody></table>
                </div>
                <div className="card">
                  <h3>General Settings</h3>
                  <p style={{ margin: "0 0 12px", color: "#64748b" }}>Organization: {db.organizationName}</p>
                  <div className="logo-upload-wrap">
                    <label className="upload-label" htmlFor="otxt">Logo Text</label>
                    <input id="otxt" value={db.organizationLogoText || ""} onChange={e => setDb(p => ({ ...p, organizationLogoText: e.target.value }))} placeholder="Enter logo text" />
                    <label className="upload-label" htmlFor="oupload">Upload Logo</label>
                    <input id="oupload" type="file" accept="image/*" onChange={handleLogoUpload} />
                    {db.organizationLogo && (
                      <div className="logo-preview-box">
                        <img className="logo-preview" src={db.organizationLogo} alt="logo" />
                        <button className="btn" onClick={() => { localStorage.removeItem(LOGO_KEY); setDb(p => ({ ...p, organizationLogo: "" })); showToast("success", "Logo removed."); }}>Remove Logo</button>
                      </div>
                    )}
                    {!db.organizationLogo && <p className="muted-text">No logo uploaded.</p>}
                  </div>
                  <div className="settings-login-card">
                    <h3>Login Account</h3>
                    <div className="form-grid">
                      <input placeholder="Full Name" value={db.loginAccount?.fullName || ""} onChange={e => setDb(p => ({ ...p, loginAccount: { ...p.loginAccount, fullName: e.target.value } }))} />
                      <input placeholder="Phone" value={db.loginAccount?.phoneNumber || ""} onChange={e => setDb(p => ({ ...p, loginAccount: { ...p.loginAccount, phoneNumber: e.target.value } }))} />
                      <input type="password" placeholder="Password (4 digits)" maxLength={4} value={db.loginAccount?.password || ""} onChange={e => setDb(p => ({ ...p, loginAccount: { ...p.loginAccount, password: e.target.value } }))} />
                      <input placeholder="Credit Account" value={db.loginAccount?.creditAccount || ""} onChange={e => setDb(p => ({ ...p, loginAccount: { ...p.loginAccount, creditAccount: e.target.value } }))} />
                      <input placeholder="Email" value={db.loginAccount?.email || ""} onChange={e => setDb(p => ({ ...p, loginAccount: { ...p.loginAccount, email: e.target.value } }))} />
                    </div>
                    <button className="btn" onClick={() => showToast("success", "Login account updated.")}>Save Login Account</button>
                  </div>

                  {/* ── Danger Zone ── */}
                  <div style={{ marginTop: 28, padding: "18px 20px", border: "2px solid #fca5a5", borderRadius: 12, background: "#fff1f2" }}>
                    <h3 style={{ margin: "0 0 8px", color: "#dc2626", fontSize: 15 }}>⚠️ Danger Zone — ሁሉንም ዳታ አጥፋ</h3>
                    <p style={{ margin: "0 0 14px", color: "#64748b", fontSize: 13 }}>
                      ሁሉንም ገቢ፣ ወጪ፣ እዳ፣ ሰራተኞች፣ ባንኮችና ሌሎች ዝርዝሮች ይጠፋሉ። Login account ብቻ ይቀራል።
                    </p>
                    <button
                      onClick={resetAllData}
                      style={{ background: "#dc2626", color: "white", border: "none", padding: "10px 20px", borderRadius: 8, fontWeight: 700, cursor: "pointer", fontSize: 14 }}
                    >
                      🗑️ ሁሉንም ዳታ አጥፋ (Reset All Data)
                    </button>
                  </div>
                </div>
              </section>
            )
          }
        </main >
      </div >
    </div >
  );
}
