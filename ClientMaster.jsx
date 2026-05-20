import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { Search, Plus, Building2, User, Users, Trash2, Edit2, Save, X, Link2, Download, Upload, Briefcase, Calendar, AlertCircle, CheckCircle2, Clock, FileText, ChevronRight, ChevronDown, Filter, DollarSign, TrendingUp, Circle, Star, Lock, LogOut, Shield, Activity, UserPlus, Eye, EyeOff, Key } from 'lucide-react';

// ROLES
const ROLES_LIST = ['Owner', 'Manager', 'Editor', 'Viewer'];
const PERMS = {
  'Owner': { canCreate: 1, canEdit: 1, canDelete: 1, canImport: 1, canManageUsers: 1, canViewAudit: 1 },
  'Manager': { canCreate: 1, canEdit: 1, canDelete: 1, canImport: 1, canManageUsers: 0, canViewAudit: 1 },
  'Editor': { canCreate: 1, canEdit: 1, canDelete: 0, canImport: 1, canManageUsers: 0, canViewAudit: 0 },
  'Viewer': { canCreate: 0, canEdit: 0, canDelete: 0, canImport: 0, canManageUsers: 0, canViewAudit: 0 }
};
const ROLE_DESC = {
  'Owner': 'Full access including user management',
  'Manager': 'All client operations plus audit log access',
  'Editor': 'Create and edit clients (no delete)',
  'Viewer': 'Read-only access to all client data'
};
const can = (u, p) => u && PERMS[u.role] && PERMS[u.role][p] === 1;

// CONSTANTS
const ENTITY_TYPES = ['LLC (Single-Member)', 'LLC (Multi-Member)', 'C-Corporation (1120)', 'S-Corporation (1120-S)', 'Partnership (1065)', 'Trust (1041)', 'Non-Profit (990)', 'Sole Proprietorship', 'Foreign Entity', 'Other'];
const FILING_STATUS = ['Single', 'Married Filing Jointly', 'Married Filing Separately', 'Head of Household', 'Qualifying Widow(er)', 'Non-Resident Alien'];
const VISA = ['US Citizen', 'Green Card', 'H-1B', 'F-1', 'L-1', 'O-1', 'B-1/B-2', 'TN', 'E-2', 'Other Non-Resident', 'N/A'];
const STATUS_OPTS = ['Active', 'Inactive', 'Prospective', 'Onboarding', 'Terminated'];
const REL_ROLES = ['Owner', 'Co-Owner', 'Spouse of Owner', 'Child / Dependent of Owner', 'Parent of Owner', 'Sibling of Owner', 'Other Relative', 'Officer / Director', 'Employee (Non-Owner)', 'Trustee', 'Beneficiary', 'Authorized Signatory'];
const FILING_DEADLINE_STATUS = ['Not Started', 'In Progress', 'Documents Pending', 'Ready for Review', 'Filed', 'Filed - Payment Pending', 'Completed', 'Extended', 'On Hold'];
const PAYROLL_FREQ = ['Weekly', 'Bi-Weekly', 'Semi-Monthly', 'Monthly', 'Quarterly', 'On-Demand'];
const PAYROLL_PROVIDER = ['In-House', 'Gusto', 'ADP', 'QuickBooks Payroll', 'Paychex', 'Rippling', 'OnPay', 'SurePayroll', 'Other'];
const PAYROLL_STATUS = ['Active', 'Inactive', 'Setup Pending', 'Suspended'];
const COMPLIANCE_STATUS = ['Not Started', 'In Progress', 'Filed', 'Pending Payment', 'Paid', 'Completed', 'Extended', 'N/A'];
const BILLING_CYCLE = ['One-Time', 'Monthly', 'Quarterly', 'Annual', 'Per Engagement'];
const PAY_STATUS = ['Pending', 'Invoiced', 'Partially Paid', 'Paid', 'Overdue', 'Written Off'];

const PAYROLL_SERVICES = ['Payroll Processing', 'Form 941 Prep', 'Form 940 Prep', 'W-2 / W-3 Filing', '1099-NEC / 1099-MISC', 'State Payroll Returns', 'New Hire Reporting'];

const SERVICE_CATEGORIES = {
  'Tax Services': ['Form 1040', 'Form 1040-NR', 'Form 1120 (C-Corp)', 'Form 1120-S (S-Corp)', 'Form 1065 (Partnership)', 'Form 1041 (Trust)', 'Form 990 (Non-Profit)', 'FBAR (FinCEN 114)', 'FATCA (Form 8938)', 'ITIN Application (W-7)', 'Amended Returns (1040-X)', 'Estimated Tax Planning'],
  'Compliance': ['BOI Reporting', 'Form 5472', 'Form 5471', 'Form 8865', 'State Registrations', 'Sales Tax Filing', 'Annual Reports', 'Resident Agent Services'],
  'Payroll': PAYROLL_SERVICES,
  'Advisory and Other': ['Bookkeeping', 'Advisory / Consulting', 'Entity Formation', 'S-Corp Election', 'Cross-Border Structuring', 'Audit Support', 'IRS Representation', 'Tax Planning', 'Other']
};

const CAT_META = {
  'Tax Services': { color: 'red', desc: 'Federal and state tax returns' },
  'Compliance': { color: 'amber', desc: 'Regulatory filings' },
  'Payroll': { color: 'navy', desc: 'Payroll processing and filings' },
  'Advisory and Other': { color: 'slate', desc: 'Consulting and special services' }
};

const ENTITY_ONLY = ['Form 1120 (C-Corp)', 'Form 1120-S (S-Corp)', 'Form 1065 (Partnership)', 'Form 990 (Non-Profit)', 'Form 5472', 'BOI Reporting', 'Annual Reports', 'Resident Agent Services', 'S-Corp Election', 'Entity Formation'].concat(PAYROLL_SERVICES);
const INDIVIDUAL_ONLY = ['Form 1040', 'Form 1040-NR', 'Amended Returns (1040-X)', 'ITIN Application (W-7)', 'Estimated Tax Planning'];

const MUTEX = [
  { services: ['Form 1040', 'Form 1040-NR'] },
  { services: ['Form 1120 (C-Corp)', 'Form 1120-S (S-Corp)', 'Form 1065 (Partnership)', 'Form 1041 (Trust)', 'Form 990 (Non-Profit)'] }
];

const DEPS = {
  'Payroll Processing': ['Form 941 Prep', 'Form 940 Prep', 'W-2 / W-3 Filing'],
  'S-Corp Election': ['Form 1120-S (S-Corp)'],
  'Entity Formation': ['BOI Reporting']
};

const ENT_FIELDS = [
  { key: 'name', label: 'Entity Name', aliases: ['name', 'entityname', 'companyname', 'clientname'] },
  { key: 'entityType', label: 'Entity Type', aliases: ['entitytype', 'type'] },
  { key: 'ein', label: 'EIN', aliases: ['ein', 'taxid', 'fein'] },
  { key: 'stateOfFormation', label: 'State', aliases: ['state', 'stateofformation'] },
  { key: 'formationDate', label: 'Formation Date', aliases: ['formationdate'] },
  { key: 'taxYearEnd', label: 'Tax Year-End', aliases: ['yearend'] },
  { key: 'mailingAddress', label: 'Mailing Address', aliases: ['address'] },
  { key: 'naicsCode', label: 'NAICS', aliases: ['naics'] },
  { key: 'serviceFee', label: 'Fee', aliases: ['fee', 'amount'] },
  { key: 'billingCycle', label: 'Billing Cycle', aliases: ['billing'] },
  { key: 'preparer', label: 'Preparer', aliases: ['preparer'] },
  { key: 'reviewer', label: 'Reviewer', aliases: ['reviewer'] },
  { key: 'status', label: 'Status', aliases: ['status'] },
  { key: 'notes', label: 'Notes', aliases: ['notes'] }
];

const IND_FIELDS = [
  { key: 'firstName', label: 'First Name', aliases: ['firstname', 'fname'] },
  { key: 'lastName', label: 'Last Name', aliases: ['lastname', 'lname'] },
  { key: 'ssn', label: 'SSN', aliases: ['ssn'] },
  { key: 'itin', label: 'ITIN', aliases: ['itin'] },
  { key: 'dob', label: 'DOB', aliases: ['dob'] },
  { key: 'filingStatus', label: 'Filing Status', aliases: ['filingstatus'] },
  { key: 'visaStatus', label: 'Visa', aliases: ['visa'] },
  { key: 'residencyState', label: 'State', aliases: ['state'] },
  { key: 'mailingAddress', label: 'Address', aliases: ['address'] },
  { key: 'email', label: 'Email', aliases: ['email'] },
  { key: 'phone', label: 'Phone', aliases: ['phone'] },
  { key: 'occupation', label: 'Occupation', aliases: ['occupation'] },
  { key: 'serviceFee', label: 'Fee', aliases: ['fee'] },
  { key: 'billingCycle', label: 'Billing Cycle', aliases: ['billing'] },
  { key: 'preparer', label: 'Preparer', aliases: ['preparer'] },
  { key: 'reviewer', label: 'Reviewer', aliases: ['reviewer'] },
  { key: 'status', label: 'Status', aliases: ['status'] },
  { key: 'notes', label: 'Notes', aliases: ['notes'] }
];

const DEFAULT_DEADLINES = [
  { id: 'd1', name: 'Form 941', period: 'Q1', form: '941', dueDate: '04-30', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd2', name: 'Form 941', period: 'Q2', form: '941', dueDate: '07-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd3', name: 'Form 941', period: 'Q3', form: '941', dueDate: '10-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd4', name: 'Form 941', period: 'Q4', form: '941', dueDate: '01-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd5', name: 'Form 940 (FUTA)', period: 'Annual', form: '940', dueDate: '01-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd6', name: 'W-2 / W-3', period: 'Annual', form: 'W-2', dueDate: '01-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd7', name: '1099-NEC', period: 'Annual', form: '1099', dueDate: '01-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd8', name: 'State W/H', period: 'Q1', form: 'State', dueDate: '04-30', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd9', name: 'State W/H', period: 'Q2', form: 'State', dueDate: '07-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd10', name: 'State W/H', period: 'Q3', form: 'State', dueDate: '10-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd11', name: 'State W/H', period: 'Q4', form: 'State', dueDate: '01-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd12', name: 'SUTA', period: 'Q1', form: 'SUTA', dueDate: '04-30', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd13', name: 'SUTA', period: 'Q2', form: 'SUTA', dueDate: '07-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd14', name: 'SUTA', period: 'Q3', form: 'SUTA', dueDate: '10-31', status: 'Not Started', filedDate: '', amount: '' },
  { id: 'd15', name: 'SUTA', period: 'Q4', form: 'SUTA', dueDate: '01-31', status: 'Not Started', filedDate: '', amount: '' }
];

const newId = (p) => p + '_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7);

const genPeriods = (cycle) => {
  const yr = new Date().getFullYear();
  if (cycle === 'Monthly') {
    return ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'].map(m => ({ id: 'p_' + yr + '_' + m, label: m + ' ' + yr, status: 'Pending', amount: '', paidDate: '', invoiceNum: '' }));
  }
  if (cycle === 'Quarterly') return ['Q1', 'Q2', 'Q3', 'Q4'].map(q => ({ id: 'p_' + yr + '_' + q, label: q + ' ' + yr, status: 'Pending', amount: '', paidDate: '', invoiceNum: '' }));
  if (cycle === 'Annual') return [{ id: 'p_' + yr, label: yr + ' Annual', status: 'Pending', amount: '', paidDate: '', invoiceNum: '' }];
  if (cycle === 'One-Time') return [{ id: 'p_one', label: 'One-Time Fee', status: 'Pending', amount: '', paidDate: '', invoiceNum: '' }];
  if (cycle === 'Per Engagement') return [{ id: 'p_eng', label: 'Engagement 1', status: 'Pending', amount: '', paidDate: '', invoiceNum: '' }];
  return [];
};

const newEntity = () => ({
  id: newId('e'), type: 'entity',
  name: '', entityType: 'LLC (Single-Member)', ein: '', stateOfFormation: 'TX',
  formationDate: '', taxYearEnd: '12-31', fiscalYear: 'Calendar Year',
  registeredAddress: '', mailingAddress: '', naicsCode: '', accountingMethod: 'Cash',
  statesFiled: '', services: [], serviceFee: '', billingCycle: 'Annual',
  paymentStatus: 'Pending', paymentPeriods: genPeriods('Annual'),
  engagementLetterDate: '', preparer: '', reviewer: '',
  status: 'Active', notes: '',
  payroll: {
    payrollStatus: 'Inactive', frequency: 'Bi-Weekly', provider: 'In-House',
    employees: '', firstPayrollDate: '', nextPayrollDue: '',
    statesRegistered: '', federalEinStatus: 'Active',
    stateWithholdingId: '', stateUnemploymentId: '',
    deadlines: JSON.parse(JSON.stringify(DEFAULT_DEADLINES))
  },
  taxFilings: [], createdAt: new Date().toISOString()
});

const newIndividual = () => ({
  id: newId('i'), type: 'individual',
  firstName: '', lastName: '', ssn: '', itin: '', dob: '',
  filingStatus: 'Single', citizenship: 'US Citizen', visaStatus: 'US Citizen',
  residencyState: 'TX', mailingAddress: '', email: '', phone: '',
  occupation: '', spouseName: '', spouseSSN: '', dependents: '',
  priorYearAGI: '', irsIpPin: '', fbarRequired: 'No', fatcaRequired: 'No',
  treatyCountry: '', services: [], serviceFee: '', billingCycle: 'Annual',
  paymentStatus: 'Pending', paymentPeriods: genPeriods('Annual'),
  engagementLetterDate: '', preparer: '', reviewer: '',
  status: 'Active', notes: '', taxFilings: [],
  createdAt: new Date().toISOString()
});

const computeChanges = (oldObj, newObj) => {
  const changes = [];
  const skip = ['id', 'type', 'createdAt', 'displayName', 'searchKey'];
  const allKeys = new Set([...Object.keys(oldObj || {}), ...Object.keys(newObj || {})]);
  allKeys.forEach(k => {
    if (skip.includes(k)) return;
    const o = JSON.stringify(oldObj && oldObj[k]);
    const n = JSON.stringify(newObj && newObj[k]);
    if (o !== n) changes.push({ field: k, from: summary(oldObj && oldObj[k]), to: summary(newObj && newObj[k]) });
  });
  return changes;
};

const summary = (v) => {
  if (v === null || v === undefined || v === '') return '(empty)';
  if (typeof v === 'string') return v.length > 60 ? v.slice(0, 60) + '...' : v;
  if (Array.isArray(v)) return v.length + ' items';
  if (typeof v === 'object') return '(object)';
  return String(v);
};

const K_CLIENTS = 'tfs_clients_v1';
const K_USERS = 'tfs_users_v1';
const K_AUDIT = 'tfs_audit_v1';
const K_SESSION = 'tfs_session_v1';
const K_LASTBACKUP = 'tfs_lastbackup_v1';

// MAIN COMPONENT
export default function ClientMaster() {
  const [data, setData] = useState({ entities: [], individuals: [], relationships: [] });
  const [users, setUsers] = useState([]);
  const [auditLog, setAuditLog] = useState([]);
  const [currentUser, setCurrentUser] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [lastBackup, setLastBackup] = useState('');
  const [search, setSearch] = useState('');
  const [filterType, setFilterType] = useState('all');
  const [filterStatus, setFilterStatus] = useState('all');
  const [selectedId, setSelectedId] = useState(null);
  const [view, setView] = useState('dashboard');
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(null);
  const [showLinkModal, setShowLinkModal] = useState(false);
  const [showNewMenu, setShowNewMenu] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showImportModal, setShowImportModal] = useState(false);
  const [importStep, setImportStep] = useState('upload');
  const [importHeaders, setImportHeaders] = useState([]);
  const [importRows, setImportRows] = useState([]);
  const [importType, setImportType] = useState('entity');
  const [importMapping, setImportMapping] = useState({});
  const [importPreview, setImportPreview] = useState([]);
  const [confirmAction, setConfirmAction] = useState(null);
  const [pickerOpen, setPickerOpen] = useState(false);
  const [pickerCategory, setPickerCategory] = useState(null);
  const [pendingService, setPendingService] = useState(null);
  const [linkRole, setLinkRole] = useState('Owner');
  const [linkOwnership, setLinkOwnership] = useState('');
  const [linkTargetId, setLinkTargetId] = useState('');
  const [editingUser, setEditingUser] = useState(null);
  const [showUserMenu, setShowUserMenu] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const [c, u, a, s, lb] = await Promise.all([
          window.storage.get(K_CLIENTS, false).catch(e => { console.error('load clients', e); return null; }),
          window.storage.get(K_USERS, false).catch(e => { console.error('load users', e); return null; }),
          window.storage.get(K_AUDIT, false).catch(e => { console.error('load audit', e); return null; }),
          window.storage.get(K_SESSION, false).catch(e => { console.error('load session', e); return null; }),
          window.storage.get(K_LASTBACKUP, false).catch(() => null)
        ]);
        if (lb && lb.value) setLastBackup(typeof lb.value === 'string' ? lb.value : '');
        if (c && c.value) {
          const p = typeof c.value === 'string' ? JSON.parse(c.value) : c.value;
          setData({ entities: p.entities || [], individuals: p.individuals || [], relationships: p.relationships || [] });
        }
        let userList = [];
        if (u && u.value) {
          userList = typeof u.value === 'string' ? JSON.parse(u.value) : u.value;
          setUsers(userList);
        }
        if (a && a.value) {
          const al = typeof a.value === 'string' ? JSON.parse(a.value) : a.value;
          setAuditLog(al);
        }
        if (s && s.value) {
          const sd = typeof s.value === 'string' ? JSON.parse(s.value) : s.value;
          if (sd.userId) {
            const u2 = userList.find(x => x.id === sd.userId);
            if (u2 && u2.status === 'Active') setCurrentUser(u2);
          }
        }
      } catch (e) {
        console.error(e);
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  const persistData = async (nd) => { setData(nd); try { await window.storage.set(K_CLIENTS, JSON.stringify(nd), false); } catch (e) { console.error('save clients', e); } };
  const persistUsers = async (nu) => { setUsers(nu); try { await window.storage.set(K_USERS, JSON.stringify(nu), false); } catch (e) { console.error('save users', e); } };
  const persistAudit = async (nl) => { setAuditLog(nl); try { await window.storage.set(K_AUDIT, JSON.stringify(nl.slice(-1000)), false); } catch (e) { console.error('save audit', e); } };
  const persistSession = async (uid) => { try { if (uid) await window.storage.set(K_SESSION, JSON.stringify({ userId: uid }), false); else await window.storage.delete(K_SESSION, false).catch(() => {}); } catch (e) {} };

  const addAuditEntry = async (existing, user, action, targetType, targetId, targetName, changes) => {
    const entry = {
      id: newId('a'), timestamp: new Date().toISOString(),
      userId: user.id, userName: user.fullName || user.username, userRole: user.role,
      action, targetType, targetId: targetId || '', targetName: targetName || '', changes: changes || []
    };
    await persistAudit([...existing, entry]);
  };

  const addAudit = (action, targetType, targetId, targetName, changes) => {
    if (!currentUser) return Promise.resolve();
    return addAuditEntry(auditLog, currentUser, action, targetType, targetId, targetName, changes);
  };

  const showAlert = (title, message) => {
    setConfirmAction({ title, message, confirmLabel: 'OK', danger: false, onConfirm: () => setConfirmAction(null) });
  };

  // AUTH
  const handleLogin = async (username, password) => {
    const u = users.find(x => x.username.toLowerCase() === username.toLowerCase() && x.password === password && x.status === 'Active');
    if (!u) return { ok: false, error: 'Invalid credentials or inactive account' };
    setCurrentUser(u);
    await persistSession(u.id);
    await addAuditEntry(auditLog, u, 'login', 'system', '', '', []);
    return { ok: true };
  };

  const handleLogout = async () => {
    if (currentUser) await addAudit('logout', 'system', '', '', []);
    setCurrentUser(null); setSelectedId(null); setView('dashboard'); setEditing(false); setDraft(null);
    await persistSession(null);
  };

  const handleSetupOwner = async (d) => {
    const owner = { id: newId('u'), username: d.username, password: d.password, fullName: d.fullName, email: d.email, role: 'Owner', status: 'Active', createdAt: new Date().toISOString() };
    await persistUsers([owner]);
    setCurrentUser(owner);
    await persistSession(owner.id);
    await addAuditEntry([], owner, 'system_initialized', 'system', '', '', []);
  };

  const saveUser = async (ud, isNew) => {
    if (isNew) {
      const nu = { id: newId('u'), username: ud.username, password: ud.password, fullName: ud.fullName, email: ud.email, role: ud.role, status: ud.status || 'Active', createdAt: new Date().toISOString() };
      await persistUsers([...users, nu]);
      await addAudit('user_created', 'user', nu.id, nu.fullName, [{ field: 'role', from: '(empty)', to: nu.role }]);
    } else {
      const oldU = users.find(u => u.id === ud.id);
      await persistUsers(users.map(u => u.id === ud.id ? ud : u));
      const ch = computeChanges(oldU, ud).map(c => c.field === 'password' ? { field: 'password', from: '***', to: '*** (changed)' } : c);
      await addAudit('user_updated', 'user', ud.id, ud.fullName, ch);
    }
    setEditingUser(null);
  };

  const deleteUser = (uid) => {
    const u = users.find(x => x.id === uid);
    if (!u) return;
    if (u.role === 'Owner') return showAlert('Cannot Delete', 'The Owner account cannot be deleted.');
    setConfirmAction({
      title: 'Delete user?', message: 'Remove "' + u.fullName + '"? They will no longer be able to log in.',
      confirmLabel: 'Delete', danger: true,
      onConfirm: async () => {
        await persistUsers(users.filter(x => x.id !== uid));
        await addAudit('user_deleted', 'user', u.id, u.fullName, []);
        setConfirmAction(null);
      }
    });
  };

  // CLIENTS
  const allClients = useMemo(() => {
    const ents = data.entities.map(e => ({ ...e, displayName: e.name || '(Unnamed Entity)', searchKey: ((e.name || '') + ' ' + (e.ein || '')).toLowerCase() }));
    const inds = data.individuals.map(i => {
      const fn = ((i.firstName || '') + ' ' + (i.lastName || '')).trim();
      return { ...i, displayName: fn || '(Unnamed Individual)', searchKey: (fn + ' ' + (i.ssn || '') + ' ' + (i.email || '')).toLowerCase() };
    });
    return [...ents, ...inds];
  }, [data]);

  const filteredClients = useMemo(() => {
    const q = search.toLowerCase().trim();
    return allClients
      .filter(c => filterType === 'all' || c.type === filterType)
      .filter(c => filterStatus === 'all' || c.status === filterStatus)
      .filter(c => !q || c.searchKey.includes(q))
      .sort((a, b) => a.displayName.localeCompare(b.displayName));
  }, [allClients, search, filterType, filterStatus]);

  const selectedClient = useMemo(() => selectedId ? (allClients.find(c => c.id === selectedId) || null) : null, [selectedId, allClients]);

  const relatedToSelected = useMemo(() => {
    if (!selectedClient) return [];
    return data.relationships
      .filter(r => r.entityId === selectedClient.id || r.individualId === selectedClient.id)
      .map(r => {
        const otherId = r.entityId === selectedClient.id ? r.individualId : r.entityId;
        const other = allClients.find(c => c.id === otherId);
        return { ...r, other };
      })
      .filter(r => r.other);
  }, [selectedClient, data.relationships, allClients]);

  const handleSelect = (id) => { setSelectedId(id); setView('detail'); setEditing(false); setDraft(null); };
  const startEdit = () => { if (!can(currentUser, 'canEdit')) return showAlert('Permission Denied', 'You cannot edit clients.'); setDraft(JSON.parse(JSON.stringify(selectedClient))); setEditing(true); };
  const cancelEdit = () => { setEditing(false); setDraft(null); setPickerOpen(false); setPendingService(null); };

  const saveEdit = async () => {
    if (!can(currentUser, 'canEdit')) return showAlert('Permission Denied', 'You cannot edit clients.');
    const cleaned = { ...draft };
    delete cleaned.displayName; delete cleaned.searchKey;
    const oldC = selectedClient;
    const ch = computeChanges(oldC, cleaned);
    if (cleaned.type === 'entity') {
      await persistData({ ...data, entities: data.entities.map(e => e.id === cleaned.id ? cleaned : e) });
    } else {
      await persistData({ ...data, individuals: data.individuals.map(i => i.id === cleaned.id ? cleaned : i) });
    }
    await addAudit('client_updated', cleaned.type, cleaned.id, oldC.displayName, ch);
    setEditing(false); setDraft(null);
  };

  const addNewClient = async (type) => {
    if (!can(currentUser, 'canCreate')) return showAlert('Permission Denied', 'You cannot create clients.');
    const fresh = type === 'entity' ? newEntity() : newIndividual();
    if (type === 'entity') await persistData({ ...data, entities: [...data.entities, fresh] });
    else await persistData({ ...data, individuals: [...data.individuals, fresh] });
    await addAudit('client_created', type, fresh.id, '(new ' + type + ')', []);
    setSelectedId(fresh.id); setView('detail'); setDraft(JSON.parse(JSON.stringify(fresh))); setEditing(true);
  };

  const deleteClient = () => {
    if (!can(currentUser, 'canDelete')) return showAlert('Permission Denied', 'You cannot delete clients.');
    if (!selectedClient) return;
    setConfirmAction({
      title: 'Delete this client?',
      message: 'Permanently delete "' + selectedClient.displayName + '" and all relationships. This is logged.',
      confirmLabel: 'Delete', danger: true,
      onConfirm: async () => {
        const newRels = data.relationships.filter(r => r.entityId !== selectedClient.id && r.individualId !== selectedClient.id);
        if (selectedClient.type === 'entity') {
          await persistData({ ...data, entities: data.entities.filter(e => e.id !== selectedClient.id), relationships: newRels });
        } else {
          await persistData({ ...data, individuals: data.individuals.filter(i => i.id !== selectedClient.id), relationships: newRels });
        }
        await addAudit('client_deleted', selectedClient.type, selectedClient.id, selectedClient.displayName, []);
        setSelectedId(null); setView('dashboard'); setEditing(false); setConfirmAction(null);
      }
    });
  };

  const addRelationship = async () => {
    if (!linkTargetId || !selectedClient) return;
    const isE = selectedClient.type === 'entity';
    const nr = { id: newId('r'), entityId: isE ? selectedClient.id : linkTargetId, individualId: isE ? linkTargetId : selectedClient.id, role: linkRole, ownership: linkOwnership, effectiveDate: new Date().toISOString().split('T')[0], notes: '' };
    await persistData({ ...data, relationships: [...data.relationships, nr] });
    await addAudit('relationship_created', 'relationship', nr.id, selectedClient.displayName + ' / ' + linkRole, []);
    setShowLinkModal(false); setLinkTargetId(''); setLinkOwnership(''); setLinkRole('Owner');
  };

  const removeRelationship = (relId) => {
    setConfirmAction({
      title: 'Remove relationship?', message: 'The two clients will no longer be linked.',
      confirmLabel: 'Remove', danger: true,
      onConfirm: async () => {
        const rel = data.relationships.find(r => r.id === relId);
        await persistData({ ...data, relationships: data.relationships.filter(r => r.id !== relId) });
        await addAudit('relationship_deleted', 'relationship', relId, (rel && rel.role) || '', []);
        setConfirmAction(null);
      }
    });
  };

  const triggerDownload = (content, filename, mimeType) => {
    try {
      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.rel = 'noopener';
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { try { document.body.removeChild(a); } catch (_) {} URL.revokeObjectURL(url); }, 250);
      return true;
    } catch (e) {
      console.error('download failed', e);
      return false;
    }
  };

  // EXPORT JSON — complete backup: clients + users + audit log. Re-importing restores everything, including accounts.
  const exportData = () => {
    if (data.entities.length === 0 && data.individuals.length === 0 && users.length === 0) return showAlert('Nothing to Export', 'There is no data to export yet.');
    const backup = { format: 'tfs_full_backup', version: 2, exportedAt: new Date().toISOString(), data, users, auditLog };
    const ok = triggerDownload(JSON.stringify(backup, null, 2), 'tfs_backup_' + new Date().toISOString().split('T')[0] + '.json', 'application/json');
    if (!ok) return showAlert('Download Failed', 'Could not trigger download. Check the browser console for details.');
    const now = new Date().toISOString();
    setLastBackup(now);
    window.storage.set(K_LASTBACKUP, now, false).catch(() => {});
  };

  const exportCSV = () => {
    if (data.entities.length === 0 && data.individuals.length === 0) return showAlert('Nothing to Export', 'You have no clients yet. Add some clients first.');
    const rows = [['Type', 'Name', 'ID', 'Status', 'Services', 'Fee', 'Billing', 'Payment', 'Preparer', 'Reviewer']];
    data.entities.forEach(e => rows.push(['Entity', e.name || '', e.ein || '', e.status || '', (e.services || []).join('; '), e.serviceFee || '', e.billingCycle || '', e.paymentStatus || '', e.preparer || '', e.reviewer || '']));
    data.individuals.forEach(i => rows.push(['Individual', ((i.firstName || '') + ' ' + (i.lastName || '')).trim(), i.ssn || i.itin || '', i.status || '', (i.services || []).join('; '), i.serviceFee || '', i.billingCycle || '', i.paymentStatus || '', i.preparer || '', i.reviewer || '']));
    const csv = rows.map(r => r.map(c => '"' + String(c || '').replace(/"/g, '""') + '"').join(',')).join('\n');
    const ok = triggerDownload('\ufeff' + csv, 'tfs_clients_' + new Date().toISOString().split('T')[0] + '.csv', 'text/csv;charset=utf-8;');
    if (!ok) showAlert('Download Failed', 'Could not trigger download. Check the browser console for details.');
  };

  // Normalize any supported JSON shape into { entities, individuals, relationships }
  const extractClients = (parsed) => {
    const src = (parsed && parsed.data) ? parsed.data : (parsed && parsed.clients ? parsed.clients : parsed);
    if (!src || typeof src !== 'object') return null;
    if (!('entities' in src) && !('individuals' in src) && !('relationships' in src)) return null;
    return { entities: Array.isArray(src.entities) ? src.entities : [], individuals: Array.isArray(src.individuals) ? src.individuals : [], relationships: Array.isArray(src.relationships) ? src.relationships : [] };
  };

  // RESTORE — accepts both the legacy clients-only JSON and the v2 full backup. Replaces current data.
  // mode 'replace' (default) overwrites; mode 'merge' appends clients by id.
  const restoreFromBackup = async (parsed, mode) => {
    const nd = extractClients(parsed);
    const nu = parsed && Array.isArray(parsed.users) ? parsed.users : null;
    const na = parsed && Array.isArray(parsed.auditLog) ? parsed.auditLog : null;
    if (!nd && !nu) return { ok: false, error: 'This file does not look like a Client Master backup (no clients or users found).' };

    let finalData = data;
    if (nd) {
      if (mode === 'merge') {
        const eIds = new Set(data.entities.map(e => e.id));
        const iIds = new Set(data.individuals.map(i => i.id));
        const rIds = new Set(data.relationships.map(r => r.id));
        finalData = {
          entities: [...data.entities, ...nd.entities.filter(e => !eIds.has(e.id))],
          individuals: [...data.individuals, ...nd.individuals.filter(i => !iIds.has(i.id))],
          relationships: [...data.relationships, ...nd.relationships.filter(r => !rIds.has(r.id))]
        };
      } else {
        finalData = nd;
      }
      await persistData(finalData);
    }
    if (nu && nu.length > 0) {
      const finalUsers = mode === 'merge' ? (() => { const ids = new Set(users.map(u => u.id)); const names = new Set(users.map(u => (u.username || '').toLowerCase())); return [...users, ...nu.filter(u => !ids.has(u.id) && !names.has((u.username || '').toLowerCase()))]; })() : nu;
      await persistUsers(finalUsers);
    }
    if (na && na.length > 0) {
      await persistAudit(mode === 'merge' ? [...auditLog, ...na].slice(-1000) : na);
    }
    if (currentUser) await addAudit('backup_restored', 'system', '', (mode === 'merge' ? 'Merged' : 'Replaced') + ' from backup', [{ field: 'clients', from: String(data.entities.length + data.individuals.length), to: String(finalData.entities.length + finalData.individuals.length) }]);
    return { ok: true, entities: (nd ? finalData.entities.length : data.entities.length), individuals: (nd ? finalData.individuals.length : data.individuals.length), users: nu ? nu.length : 0, hadUsers: !!(nu && nu.length > 0) };
  };

  const handleRestoreFile = async (file, mode) => {
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      const res = await restoreFromBackup(parsed, mode);
      if (!res.ok) { showAlert('Restore Failed', res.error); return res; }
      const parts = [];
      parts.push((res.entities + res.individuals) + ' client' + ((res.entities + res.individuals) === 1 ? '' : 's'));
      if (res.users) parts.push(res.users + ' user' + (res.users === 1 ? '' : 's'));
      showAlert('Restore Complete', 'Restored ' + parts.join(' and ') + '.' + (res.hadUsers ? ' Sign in with your existing credentials.' : ''));
      return res;
    } catch (e) {
      console.error('restore', e);
      showAlert('Restore Failed', 'Could not read or parse the file. Make sure it is a JSON backup exported from Client Master.');
      return { ok: false, error: 'parse' };
    }
  };

  const handleImportFile = async (file) => {
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
      if (!json || json.length === 0) return showAlert('Empty File', 'No data found.');
      let hi = 0;
      for (let i = 0; i < json.length; i++) if (json[i].some(c => String(c).trim() !== '')) { hi = i; break; }
      const headers = (json[hi] || []).map(h => String(h || '').trim()).filter(h => h);
      if (headers.length === 0) return showAlert('No Headers', 'No column headers detected.');
      const rows = json.slice(hi + 1).filter(r => r.some(c => String(c).trim() !== '')).map(r => {
        const o = {}; headers.forEach((h, i) => { o[h] = String(r[i] || '').trim(); }); return o;
      });
      if (rows.length === 0) return showAlert('No Data', 'No data rows found.');
      setImportHeaders(headers); setImportRows(rows);
      const auto = {};
      const fields = importType === 'entity' ? ENT_FIELDS : IND_FIELDS;
      headers.forEach(h => {
        const lo = h.toLowerCase().replace(/[^a-z0-9]/g, '');
        const m = fields.find(f => (f.aliases || []).map(a => a.toLowerCase().replace(/[^a-z0-9]/g, '')).some(a => lo === a || lo.includes(a) || a.includes(lo)));
        if (m) auto[h] = m.key;
      });
      setImportMapping(auto); setImportStep('map');
    } catch (e) { showAlert('Error', 'Could not read file.'); }
  };

  const buildPreview = () => {
    const fields = importType === 'entity' ? ENT_FIELDS : IND_FIELDS;
    const fk = {}; fields.forEach(f => { fk[f.key] = f; });
    const pv = importRows.map(row => {
      const obj = importType === 'entity' ? newEntity() : newIndividual();
      Object.keys(importMapping).forEach(h => {
        const k = importMapping[h];
        if (!k || k === '__skip__') return;
        const v = row[h]; if (!v) return;
        if (fk[k]) obj[k] = v;
      });
      return obj;
    });
    setImportPreview(pv); setImportStep('preview');
  };

  const confirmImport = async () => {
    const count = importPreview.length;
    const tl = importType === 'entity' ? 'entities' : 'individuals';
    if (importType === 'entity') await persistData({ ...data, entities: [...data.entities, ...importPreview] });
    else await persistData({ ...data, individuals: [...data.individuals, ...importPreview] });
    await addAudit('import_completed', importType, '', count + ' ' + tl, [{ field: 'imported', from: '0', to: String(count) }]);
    setShowImportModal(false); resetImport();
    showAlert('Import Successful', 'Imported ' + count + ' ' + tl + '.');
  };

  const resetImport = () => { setImportStep('upload'); setImportHeaders([]); setImportRows([]); setImportMapping({}); setImportPreview([]); };

  const stats = useMemo(() => {
    const tE = data.entities.length, tI = data.individuals.length;
    const aP = data.entities.filter(e => (e.services || []).some(s => PAYROLL_SERVICES.includes(s))).length;
    const p941 = data.entities.reduce((s, e) => s + ((e.payroll && e.payroll.deadlines) || []).filter(d => d.name === 'Form 941' && !['Filed', 'Completed', 'Paid'].includes(d.status)).length, 0);
    return { totalEntities: tE, totalIndividuals: tI, activePayroll: aP, pending941: p941 };
  }, [data]);

  const updateDraft = (path, value) => {
    setDraft(prev => {
      const c = JSON.parse(JSON.stringify(prev));
      const k = path.split('.'); let cur = c;
      for (let i = 0; i < k.length - 1; i++) cur = cur[k[i]];
      cur[k[k.length - 1]] = value;
      return c;
    });
  };

  const updateBillingCycle = (nc) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.billingCycle = nc; c.paymentPeriods = genPeriods(nc); return c; });
  const updatePeriod = (i, f, v) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.paymentPeriods[i][f] = v; return c; });
  const addPaymentPeriod = () => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.paymentPeriods = c.paymentPeriods || []; c.paymentPeriods.push({ id: newId('p'), label: 'New Period', status: 'Pending', amount: '', paidDate: '', invoiceNum: '' }); return c; });
  const removePeriod = (i) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.paymentPeriods.splice(i, 1); return c; });
  const updateDeadline = (i, f, v) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.payroll.deadlines[i][f] = v; return c; });
  const setF941Frequency = (freq) => setDraft(p => {
    const c = JSON.parse(JSON.stringify(p));
    c.payroll = c.payroll || {};
    c.payroll.f941Frequency = freq;
    c.payroll.deadlines = c.payroll.deadlines || [];
    if (freq === 'Monthly' && !c.payroll.deadlines.some(d => d.name === 'Form 941' && d.freq === 'Monthly')) {
      const due = { Jan: '02-15', Feb: '03-15', Mar: '04-15', Apr: '05-15', May: '06-15', Jun: '07-15', Jul: '08-15', Aug: '09-15', Sep: '10-15', Oct: '11-15', Nov: '12-15', Dec: '01-15' };
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      const monthly = months.map(m => ({ id: newId('d941m'), name: 'Form 941', period: m, form: '941', freq: 'Monthly', dueDate: due[m], status: 'Not Started', filedDate: '', amount: '' }));
      c.payroll.deadlines = [...c.payroll.deadlines, ...monthly];
    }
    return c;
  });

  const toggleService = (s) => {
    setDraft(prev => {
      const c = JSON.parse(JSON.stringify(prev));
      const sv = c.services || [];
      if (sv.includes(s)) { c.services = sv.filter(x => x !== s); return c; }
      const mg = MUTEX.find(g => g.services.includes(s));
      if (mg) c.services = sv.filter(x => !mg.services.includes(x)); else c.services = [...sv];
      c.services.push(s);
      return c;
    });
    const deps = DEPS[s];
    if (deps && deps.length > 0) {
      setTimeout(() => setDraft(p => {
        if (!p) return p;
        const missing = deps.filter(d => !(p.services || []).includes(d));
        if (missing.length > 0) setPendingService({ parent: s, suggestions: missing });
        return p;
      }), 50);
    }
  };

  const acceptSuggestion = (s) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); const sv = c.services || []; if (!sv.includes(s)) c.services = [...sv, s]; return c; });
  const dismissSuggestions = () => setPendingService(null);
  const addTaxFiling = () => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.taxFilings = c.taxFilings || []; c.taxFilings.push({ id: newId('f'), year: new Date().getFullYear() - 1, formType: c.type === 'entity' ? '1120-S' : '1040', status: 'Not Started', dueDate: '', filedDate: '', extensionFiled: 'No', refundOrDue: '', amount: '', notes: '' }); return c; });
  const updateFiling = (i, f, v) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.taxFilings[i][f] = v; return c; });
  const removeFiling = (i) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.taxFilings.splice(i, 1); return c; });

  const addBank = () => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.bankAccounts = c.bankAccounts || []; c.bankAccounts.push({ id: newId('b'), bankName: '', accountType: 'Bank Account', accountHolder: '', accountNumber: '', routingNumber: '', loginUrl: '', loginUsername: '', loginPassword: '', notes: '' }); return c; });
  const updateBank = (i, f, v) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.bankAccounts[i][f] = v; return c; });
  const removeBank = (i) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.bankAccounts.splice(i, 1); return c; });
  const addEmployee = () => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.payroll = c.payroll || {}; c.payroll.employeeList = c.payroll.employeeList || []; c.payroll.employeeList.push({ id: newId('emp'), fullName: '', status: 'Active', dob: '', ssn: '', position: '', hireDate: '', termDate: '', email: '', phone: '', bankName: '', bankAccountType: 'Checking', bankRouting: '', bankAccount: '', idCardName: '', idCardData: '', notes: '' }); return c; });
  const updateEmployee = (i, f, v) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.payroll.employeeList[i][f] = v; return c; });
  const removeEmployee = (i) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.payroll.employeeList.splice(i, 1); return c; });
  const attachIdCard = (i, name, dataUrl) => setDraft(p => { const c = JSON.parse(JSON.stringify(p)); c.payroll.employeeList[i].idCardName = name; c.payroll.employeeList[i].idCardData = dataUrl; return c; });

  if (!loaded) return <div className="min-h-screen bg-slate-50 flex items-center justify-center"><div className="text-slate-500">Loading...</div></div>;
  if (users.length === 0) return <SetupScreen onSetup={handleSetupOwner} onRestore={handleRestoreFile} hasData={data.entities.length + data.individuals.length > 0} />;
  if (!currentUser) return <LoginScreen onLogin={handleLogin} onRestore={handleRestoreFile} />;

  const current = editing ? draft : selectedClient;
  const hasPayroll = current && current.type === 'entity' && (current.services || []).some(s => PAYROLL_SERVICES.includes(s));
  const fontStack = '-apple-system, BlinkMacSystemFont, "Inter", "Segoe UI", Roboto, sans-serif';

  return (
    <div className="min-h-screen text-slate-900" style={{ fontFamily: fontStack, background: 'linear-gradient(135deg, #fef2f2 0%, #ffffff 50%, #eff6ff 100%)' }}>
      <Header
        currentUser={currentUser} showUserMenu={showUserMenu} setShowUserMenu={setShowUserMenu}
        handleLogout={handleLogout} search={search} setSearch={setSearch}
        filteredClients={filteredClients} handleSelect={handleSelect}
        view={view} setView={setView} setSelectedId={setSelectedId}
        showNewMenu={showNewMenu} setShowNewMenu={setShowNewMenu}
        showExportMenu={showExportMenu} setShowExportMenu={setShowExportMenu}
        addNewClient={addNewClient} exportCSV={exportCSV} exportData={exportData}
        handleRestoreFile={handleRestoreFile}
        setShowImportModal={setShowImportModal} resetImport={resetImport}
      />

      <div className="max-w-7xl mx-auto px-4 py-6">
        {view === 'dashboard' && <Dashboard stats={stats} allClients={allClients} data={data} handleSelect={handleSelect} currentUser={currentUser} exportData={exportData} lastBackup={lastBackup} />}
        {view === 'list' && <ClientList filterType={filterType} setFilterType={setFilterType} filterStatus={filterStatus} setFilterStatus={setFilterStatus} filteredClients={filteredClients} handleSelect={handleSelect} />}
        {view === 'users' && can(currentUser, 'canManageUsers') && <UsersView users={users} currentUser={currentUser} setEditingUser={setEditingUser} deleteUser={deleteUser} />}
        {view === 'audit' && can(currentUser, 'canViewAudit') && <AuditView auditLog={auditLog} users={users} />}
        {view === 'detail' && selectedClient && current && (
          <DetailView
            selectedClient={selectedClient} current={current} editing={editing} draft={draft} currentUser={currentUser}
            startEdit={startEdit} cancelEdit={cancelEdit} saveEdit={saveEdit} deleteClient={deleteClient}
            setShowLinkModal={setShowLinkModal} updateDraft={updateDraft} updateBillingCycle={updateBillingCycle}
            updatePeriod={updatePeriod} addPaymentPeriod={addPaymentPeriod} removePeriod={removePeriod}
            updateDeadline={updateDeadline} setF941Frequency={setF941Frequency} toggleService={toggleService}
            addTaxFiling={addTaxFiling} updateFiling={updateFiling} removeFiling={removeFiling}
            addBank={addBank} updateBank={updateBank} removeBank={removeBank}
            addEmployee={addEmployee} updateEmployee={updateEmployee} removeEmployee={removeEmployee} attachIdCard={attachIdCard}
            hasPayroll={hasPayroll} relatedToSelected={relatedToSelected} removeRelationship={removeRelationship}
            pickerOpen={pickerOpen} setPickerOpen={setPickerOpen}
            pickerCategory={pickerCategory} setPickerCategory={setPickerCategory}
            pendingService={pendingService} acceptSuggestion={acceptSuggestion} dismissSuggestions={dismissSuggestions}
          />
        )}
      </div>

      {showLinkModal && selectedClient && <LinkModal selectedClient={selectedClient} data={data} linkRole={linkRole} setLinkRole={setLinkRole} linkOwnership={linkOwnership} setLinkOwnership={setLinkOwnership} linkTargetId={linkTargetId} setLinkTargetId={setLinkTargetId} addRelationship={addRelationship} onClose={() => setShowLinkModal(false)} />}
      {confirmAction && <ConfirmDialog action={confirmAction} onClose={() => setConfirmAction(null)} />}
      {showImportModal && <ImportModal importStep={importStep} importType={importType} setImportType={setImportType} importHeaders={importHeaders} importRows={importRows} importMapping={importMapping} setImportMapping={setImportMapping} importPreview={importPreview} handleImportFile={handleImportFile} buildPreview={buildPreview} confirmImport={confirmImport} setImportStep={setImportStep} resetImport={resetImport} onClose={() => setShowImportModal(false)} />}
      {editingUser && <UserEditModal user={editingUser} users={users} currentUser={currentUser} onSave={saveUser} onClose={() => setEditingUser(null)} />}
    </div>
  );
}

// LOGO
function TFLogo({ size }) {
  const s = size || 40;
  return (
    <svg viewBox="0 0 64 64" width={s} height={s} xmlns="http://www.w3.org/2000/svg" style={{ filter: 'drop-shadow(0 2px 4px rgba(0,0,0,0.15))' }}>
      <defs>
        <linearGradient id="tflog" x1="0" y1="0" x2="1" y2="1">
          <stop offset="0%" stopColor="#1e3a8a" />
          <stop offset="100%" stopColor="#b91c1c" />
        </linearGradient>
      </defs>
      <rect width="64" height="64" rx="12" fill="url(#tflog)" />
      <rect x="0" y="0" width="22" height="64" fill="#1e3a8a" />
      <polygon points="11,18 13,24 19,24 14.5,28 16.5,34 11,30.5 5.5,34 7.5,28 3,24 9,24" fill="white" />
      <text x="40" y="32" fontFamily="Georgia, serif" fontSize="22" fontWeight="900" fill="white" textAnchor="middle">TF</text>
      <rect x="22" y="44" width="42" height="6" fill="white" />
      <text x="43" y="58" fontFamily="Arial, sans-serif" fontSize="7" fontWeight="700" fill="white" textAnchor="middle" letterSpacing="0.5">FINSERV</text>
    </svg>
  );
}

// SETUP
function SetupScreen({ onSetup, onRestore, hasData }) {
  const [un, setUn] = useState('');
  const [restoreMsg, setRestoreMsg] = useState('');
  const [restoreErr, setRestoreErr] = useState('');
  const [pw, setPw] = useState('');
  const [cpw, setCpw] = useState('');
  const [fn, setFn] = useState('');
  const [em, setEm] = useState('');
  const [err, setErr] = useState('');
  const [show, setShow] = useState(false);

  const submit = async () => {
    if (!un.trim() || !pw || !fn.trim()) return setErr('Username, password, and full name are required');
    if (pw.length < 6) return setErr('Password must be at least 6 characters');
    if (pw !== cpw) return setErr('Passwords do not match');
    setErr('');
    await onSetup({ username: un.trim(), password: pw, fullName: fn.trim(), email: em.trim() });
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: 'linear-gradient(135deg, #fef2f2 0%, #ffffff 50%, #eff6ff 100%)' }}>
      <div className="bg-white rounded-2xl shadow-xl border-t-4 border-red-700 max-w-md w-full p-8">
        <div className="flex items-center gap-3 mb-2">
          <TFLogo size={48} />
          <div>
            <h1 className="text-xl font-bold text-slate-900 leading-tight">Texas FinServ INC</h1>
            <p className="text-xs text-red-700 font-bold uppercase tracking-wider">Initial Setup</p>
          </div>
        </div>
        <h2 className="text-lg font-bold text-slate-900 mt-4 mb-1">Create Owner Account</h2>
        <p className="text-sm text-slate-600 mb-6">You're the first user. Set up the Owner account with full access.</p>
        <div className="space-y-3">
          <SetupField label="Full Name" value={fn} onChange={setFn} placeholder="John Smith" />
          <SetupField label="Email" value={em} onChange={setEm} placeholder="john@texasfinserv.com" type="email" />
          <SetupField label="Username" value={un} onChange={setUn} placeholder="jsmith" />
          <PwField label="Password" value={pw} onChange={setPw} show={show} setShow={setShow} placeholder="Minimum 6 characters" />
          <SetupField label="Confirm Password" value={cpw} onChange={setCpw} type="password" placeholder="Re-enter password" />
          {err && <Errs msg={err} />}
          <button onClick={submit} className="w-full bg-red-700 text-white font-bold py-2.5 rounded-lg hover:bg-red-800 shadow-sm flex items-center justify-center gap-2 mt-2">
            <Shield className="w-4 h-4" /> Create Owner Account
          </button>
        </div>
        <div className="mt-5 pt-4 border-t border-slate-100">
          {hasData && (
            <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-2.5 mb-3 text-xs text-emerald-800 flex items-start gap-2">
              <CheckCircle2 className="w-3.5 h-3.5 mt-0.5 flex-shrink-0" />
              <div>Existing client data was found in storage, but no user accounts. Create an Owner account above to access it, or restore a full backup below.</div>
            </div>
          )}
          {restoreMsg && (
            <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-2.5 mb-3 text-xs text-emerald-800 flex items-start gap-2">
              <CheckCircle2 className="w-3.5 h-3.5 mt-0.5 flex-shrink-0" /> <div>{restoreMsg}</div>
            </div>
          )}
          {restoreErr && (
            <div className="bg-rose-50 border border-rose-200 rounded-lg p-2.5 mb-3 text-xs text-rose-700 flex items-start gap-2">
              <AlertCircle className="w-3.5 h-3.5 mt-0.5 flex-shrink-0" /> <div>{restoreErr}</div>
            </div>
          )}
          <div className="text-xs font-bold text-slate-600 uppercase tracking-wide mb-1.5">Coming from another chat or device?</div>
          <p className="text-xs text-slate-500 mb-2">Import a <strong>JSON full backup</strong> (Export JSON) to recover everything — clients <em>and</em> user accounts. If it contains users, you'll be taken straight to the sign-in page.</p>
          <label className="w-full bg-blue-900 text-white font-bold py-2.5 rounded-lg hover:bg-blue-950 shadow-sm flex items-center justify-center gap-2 cursor-pointer text-sm">
            <Upload className="w-4 h-4" /> Import / Restore JSON Backup
            <input type="file" accept=".json,application/json" className="hidden" onChange={async (e) => {
              const f = e.target.files && e.target.files[0]; e.target.value = '';
              if (!f || !onRestore) return;
              setRestoreMsg(''); setRestoreErr('');
              const res = await onRestore(f, 'replace');
              if (!res || !res.ok) { setRestoreErr('Could not read that file. Make sure it is a JSON backup exported from Client Master.'); return; }
              if (res.hadUsers) { setRestoreMsg('Restored ' + (res.entities + res.individuals) + ' client(s) and ' + res.users + ' user(s). Loading the sign-in page…'); }
              else { setRestoreMsg('Imported ' + (res.entities + res.individuals) + ' client(s). This backup had no user accounts — create your Owner account above to access them.'); }
            }} />
          </label>
        </div>
        <div className="mt-5 pt-4 border-t border-slate-100 text-xs text-slate-500">
          <strong>Note:</strong> Credentials stored in browser storage. For internal firm use only.
        </div>
      </div>
    </div>
  );
}

function SetupField({ label, value, onChange, placeholder, type }) {
  return (
    <div>
      <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">{label}</label>
      <input type={type || 'text'} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-600" />
    </div>
  );
}

function PwField({ label, value, onChange, show, setShow, placeholder }) {
  return (
    <div>
      <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">{label}</label>
      <div className="relative">
        <input type={show ? 'text' : 'password'} value={value} onChange={e => onChange(e.target.value)} placeholder={placeholder} className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 pr-10 focus:outline-none focus:ring-2 focus:ring-red-600" />
        <button onClick={() => setShow(!show)} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
          {show ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
        </button>
      </div>
    </div>
  );
}

function Errs({ msg }) {
  return (
    <div className="bg-rose-50 border border-rose-200 text-rose-700 text-sm rounded-lg px-3 py-2 flex items-center gap-2">
      <AlertCircle className="w-4 h-4 flex-shrink-0" /> {msg}
    </div>
  );
}

// LOGIN
function LoginScreen({ onLogin, onRestore }) {
  const [un, setUn] = useState('');
  const [pw, setPw] = useState('');
  const [err, setErr] = useState('');
  const [show, setShow] = useState(false);
  const [busy, setBusy] = useState(false);
  const [restoreMsg, setRestoreMsg] = useState('');

  const submit = async () => {
    if (!un || !pw) return setErr('Enter username and password');
    setBusy(true);
    const r = await onLogin(un, pw);
    setBusy(false);
    if (!r.ok) setErr(r.error || 'Login failed');
  };

  const doRestore = async (e) => {
    const f = e.target.files && e.target.files[0]; e.target.value = '';
    if (!f || !onRestore) return;
    setErr(''); setRestoreMsg('');
    const res = await onRestore(f, 'replace');
    if (!res || !res.ok) { setErr('Could not read that backup file. Make sure it is a Client Master JSON export.'); return; }
    const c = res.entities + res.individuals;
    setRestoreMsg('Restored ' + c + ' client' + (c === 1 ? '' : 's') + (res.users ? ' and ' + res.users + ' user' + (res.users === 1 ? '' : 's') : '') + '. Sign in below with the restored credentials.');
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4" style={{ background: 'linear-gradient(135deg, #fef2f2 0%, #ffffff 50%, #eff6ff 100%)' }}>
      <div className="bg-white rounded-2xl shadow-xl border-t-4 border-red-700 max-w-sm w-full p-8">
        <div className="flex flex-col items-center mb-6">
          <TFLogo size={64} />
          <h1 className="text-xl font-bold text-slate-900 mt-3">Texas FinServ INC</h1>
          <p className="text-xs text-red-700 font-bold uppercase tracking-wider mt-0.5">Client Master · Sign In</p>
        </div>
        <div className="space-y-3">
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">Username</label>
            <input value={un} onChange={e => setUn(e.target.value)} onKeyDown={e => e.key === 'Enter' && submit()} placeholder="Your username" autoFocus className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-600" />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">Password</label>
            <div className="relative">
              <input type={show ? 'text' : 'password'} value={pw} onChange={e => setPw(e.target.value)} onKeyDown={e => e.key === 'Enter' && submit()} placeholder="Your password" className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 pr-10 focus:outline-none focus:ring-2 focus:ring-red-600" />
              <button onClick={() => setShow(!show)} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600">
                {show ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>
          {err && <Errs msg={err} />}
          {restoreMsg && (
            <div className="bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2 text-xs text-emerald-800 flex items-start gap-2">
              <CheckCircle2 className="w-3.5 h-3.5 mt-0.5 flex-shrink-0" /> <div>{restoreMsg}</div>
            </div>
          )}
          <button onClick={submit} disabled={busy} className="w-full bg-red-700 text-white font-bold py-2.5 rounded-lg hover:bg-red-800 shadow-sm flex items-center justify-center gap-2 disabled:opacity-50">
            <Lock className="w-4 h-4" /> {busy ? 'Signing in...' : 'Sign In'}
          </button>
        </div>
        <div className="mt-5 pt-4 border-t border-slate-100">
          <p className="text-xs text-slate-500 mb-2">Restarting in a new chat? Load your saved <strong>JSON full backup</strong> to restore all clients and accounts, then sign in.</p>
          <label className="w-full border border-blue-900 text-blue-900 font-bold py-2 rounded-lg hover:bg-blue-50 flex items-center justify-center gap-2 cursor-pointer text-sm">
            <Upload className="w-4 h-4" /> Import / Restore JSON Backup
            <input type="file" accept=".json,application/json" className="hidden" onChange={doRestore} />
          </label>
        </div>
      </div>
    </div>
  );
}

// HEADER
function Header(p) {
  const cls = (active) => 'px-3 py-2 rounded-lg text-sm font-medium transition-colors ' + (active ? 'bg-red-50 text-red-700 border border-red-200' : 'text-slate-600 hover:bg-slate-100');
  return (
    <div className="bg-white border-b-4 border-red-700 sticky top-0 z-30 shadow-md">
      <div className="bg-gradient-to-r from-red-700 via-white to-blue-900 h-1" />
      <div className="max-w-7xl mx-auto px-4 py-3 flex items-center gap-3 flex-wrap">
        <div className="flex items-center gap-2.5">
          <TFLogo size={44} />
          <div>
            <h1 className="text-base font-bold tracking-tight text-slate-900 leading-none">Texas FinServ INC</h1>
            <span className="text-[10px] text-red-700 font-bold tracking-wider uppercase">Client Master · CPA Practice</span>
          </div>
        </div>
        <div className="flex-1 min-w-[200px] max-w-md relative">
          <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input type="text" value={p.search} onChange={e => p.setSearch(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter' && p.filteredClients.length > 0) p.handleSelect(p.filteredClients[0].id); }}
            placeholder="Search clients..." className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-200 bg-slate-50 text-sm focus:outline-none focus:ring-2 focus:ring-red-600 focus:bg-white" />
        </div>
        <div className="flex items-center gap-2 flex-wrap">
          <button onClick={() => { p.setView('dashboard'); p.setSelectedId(null); }} className={cls(p.view === 'dashboard')}>Dashboard</button>
          <button onClick={() => { p.setView('list'); p.setSelectedId(null); }} className={cls(p.view === 'list')}>Clients</button>
          {can(p.currentUser, 'canManageUsers') && (
            <button onClick={() => { p.setView('users'); p.setSelectedId(null); }} className={cls(p.view === 'users') + ' flex items-center gap-1'}><Users className="w-4 h-4" />Users</button>
          )}
          {can(p.currentUser, 'canViewAudit') && (
            <button onClick={() => { p.setView('audit'); p.setSelectedId(null); }} className={cls(p.view === 'audit') + ' flex items-center gap-1'}><Activity className="w-4 h-4" />Audit</button>
          )}
          {can(p.currentUser, 'canImport') && (
            <button onClick={() => { p.setShowImportModal(true); p.resetImport(); }} className="px-3 py-2 rounded-lg border border-slate-200 text-slate-700 text-sm font-medium hover:bg-slate-50 flex items-center gap-1">
              <Upload className="w-4 h-4" /> Import
            </button>
          )}
          {can(p.currentUser, 'canCreate') && (
            <div className="relative">
              <button onClick={() => { p.setShowNewMenu(!p.showNewMenu); p.setShowExportMenu(false); }} className="px-3 py-2 rounded-lg bg-red-700 text-white text-sm font-medium hover:bg-red-800 flex items-center gap-1 shadow-sm">
                <Plus className="w-4 h-4" /> New
              </button>
              {p.showNewMenu && (
                <>
                  <div className="fixed inset-0 z-30" onClick={() => p.setShowNewMenu(false)} />
                  <div className="absolute right-0 mt-1 w-48 bg-white rounded-lg shadow-lg border border-slate-200 py-1 z-40">
                    <button onClick={() => { p.addNewClient('entity'); p.setShowNewMenu(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2"><Building2 className="w-4 h-4 text-red-700" /> Entity Client</button>
                    <button onClick={() => { p.addNewClient('individual'); p.setShowNewMenu(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2"><User className="w-4 h-4 text-blue-900" /> Individual Client</button>
                  </div>
                </>
              )}
            </div>
          )}
          <div className="relative">
            <button onClick={() => { p.setShowExportMenu(!p.showExportMenu); p.setShowNewMenu(false); }} className="px-2 py-2 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50">
              <Download className="w-4 h-4" />
            </button>
            {p.showExportMenu && (
              <>
                <div className="fixed inset-0 z-30" onClick={() => p.setShowExportMenu(false)} />
                <div className="absolute right-0 mt-1 w-56 bg-white rounded-lg shadow-lg border border-slate-200 py-1 z-40">
                  <div className="px-3 py-1 text-[10px] font-bold text-slate-400 uppercase tracking-wider">Export</div>
                  <button onClick={() => { p.exportCSV(); p.setShowExportMenu(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50">Export CSV (clients only)</button>
                  <button onClick={() => { p.exportData(); p.setShowExportMenu(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 font-medium text-blue-900">Export JSON (full backup — incl. users)</button>
                  <div className="border-t border-slate-100 my-1" />
                  <div className="px-3 py-1 text-[10px] font-bold text-slate-400 uppercase tracking-wider">Restore</div>
                  <label className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2 cursor-pointer text-red-700 font-medium">
                    <Upload className="w-3.5 h-3.5" /> Restore from Backup…
                    <input type="file" accept=".json,application/json" className="hidden" onChange={(e) => { const f = e.target.files && e.target.files[0]; p.setShowExportMenu(false); if (f) p.handleRestoreFile(f, 'replace'); e.target.value = ''; }} />
                  </label>
                  <label className="w-full text-left px-3 py-2 text-xs hover:bg-slate-50 flex items-center gap-2 cursor-pointer text-slate-500">
                    Merge backup (keep current)
                    <input type="file" accept=".json,application/json" className="hidden" onChange={(e) => { const f = e.target.files && e.target.files[0]; p.setShowExportMenu(false); if (f) p.handleRestoreFile(f, 'merge'); e.target.value = ''; }} />
                  </label>
                </div>
              </>
            )}
          </div>
          <div className="relative">
            <button onClick={() => p.setShowUserMenu(!p.showUserMenu)} className="flex items-center gap-2 px-2 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50">
              <div className="w-7 h-7 rounded-full bg-gradient-to-br from-red-700 to-blue-900 text-white flex items-center justify-center text-xs font-bold">
                {(p.currentUser.fullName || p.currentUser.username).charAt(0).toUpperCase()}
              </div>
              <div className="text-left hidden sm:block">
                <div className="text-xs font-bold text-slate-900 leading-tight">{p.currentUser.fullName || p.currentUser.username}</div>
                <div className="text-[10px] text-red-700 font-bold uppercase tracking-wide leading-tight">{p.currentUser.role}</div>
              </div>
            </button>
            {p.showUserMenu && (
              <>
                <div className="fixed inset-0 z-30" onClick={() => p.setShowUserMenu(false)} />
                <div className="absolute right-0 mt-1 w-56 bg-white rounded-lg shadow-lg border border-slate-200 py-1 z-40">
                  <div className="px-3 py-2 border-b border-slate-100">
                    <div className="text-sm font-bold text-slate-900">{p.currentUser.fullName || p.currentUser.username}</div>
                    <div className="text-xs text-slate-500">@{p.currentUser.username}</div>
                    <div className="text-xs text-red-700 font-bold mt-1">{p.currentUser.role}</div>
                  </div>
                  <button onClick={() => { p.handleLogout(); p.setShowUserMenu(false); }} className="w-full text-left px-3 py-2 text-sm hover:bg-slate-50 flex items-center gap-2 text-rose-600">
                    <LogOut className="w-4 h-4" /> Sign Out
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

// USERS VIEW
function UsersView({ users, currentUser, setEditingUser, deleteUser }) {
  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <Users className="w-5 h-5 text-red-700" />
            <h2 className="text-xl font-bold text-slate-900">Team Members</h2>
          </div>
          <p className="text-sm text-slate-500">Manage employee accounts and permissions</p>
        </div>
        <button onClick={() => setEditingUser({ isNew: true, role: 'Editor', status: 'Active' })} className="px-3 py-2 rounded-lg bg-red-700 text-white text-sm font-medium hover:bg-red-800 flex items-center gap-1 shadow-sm">
          <UserPlus className="w-4 h-4" /> Add Team Member
        </button>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
        {users.map(u => <UserCard key={u.id} user={u} currentUser={currentUser} onEdit={() => setEditingUser({ ...u })} onDelete={() => deleteUser(u.id)} />)}
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-4 mt-6">
        <h3 className="text-sm font-bold text-slate-900 mb-2 flex items-center gap-1.5">
          <Shield className="w-4 h-4 text-red-700" /> Role Permissions
        </h3>
        <div className="space-y-1.5 text-xs">
          {ROLES_LIST.map(r => (
            <div key={r} className="flex items-start gap-2">
              <span className="font-bold text-slate-700 w-20 flex-shrink-0">{r}:</span>
              <span className="text-slate-600">{ROLE_DESC[r]}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function UserCard({ user, currentUser, onEdit, onDelete }) {
  const isMe = user.id === currentUser.id;
  const isOwner = user.role === 'Owner';
  const rc = { 'Owner': 'bg-red-100 text-red-700 border-red-200', 'Manager': 'bg-blue-100 text-blue-900 border-blue-200', 'Editor': 'bg-amber-100 text-amber-700 border-amber-200', 'Viewer': 'bg-slate-100 text-slate-700 border-slate-200' };
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-sm">
      <div className="flex items-start gap-3 mb-3">
        <div className="w-10 h-10 rounded-full bg-gradient-to-br from-red-700 to-blue-900 text-white flex items-center justify-center font-bold flex-shrink-0">
          {(user.fullName || user.username).charAt(0).toUpperCase()}
        </div>
        <div className="flex-1 min-w-0">
          <div className="font-bold text-slate-900 truncate">{user.fullName || user.username}</div>
          <div className="text-xs text-slate-500 truncate">@{user.username}</div>
          {user.email && <div className="text-xs text-slate-500 truncate">{user.email}</div>}
        </div>
      </div>
      <div className="flex items-center gap-2 mb-3 flex-wrap">
        <span className={'px-2 py-0.5 text-xs font-bold rounded border ' + rc[user.role]}>{user.role}</span>
        <span className={'px-2 py-0.5 text-xs font-bold rounded ' + (user.status === 'Active' ? 'bg-emerald-100 text-emerald-700' : 'bg-slate-100 text-slate-600')}>{user.status}</span>
        {isMe && <span className="px-2 py-0.5 text-xs font-bold rounded bg-violet-100 text-violet-700">You</span>}
      </div>
      <div className="flex items-center gap-2">
        <button onClick={onEdit} className="flex-1 px-3 py-1.5 text-xs rounded-lg border border-slate-200 hover:bg-slate-50 flex items-center justify-center gap-1">
          <Edit2 className="w-3 h-3" /> Edit
        </button>
        {!isOwner && !isMe && (
          <button onClick={onDelete} className="px-2 py-1.5 text-xs rounded-lg border border-rose-200 text-rose-600 hover:bg-rose-50">
            <Trash2 className="w-3 h-3" />
          </button>
        )}
      </div>
    </div>
  );
}

function UserEditModal({ user, users, currentUser, onSave, onClose }) {
  const isNew = user.isNew === true;
  const [form, setForm] = useState({ id: user.id || '', username: user.username || '', password: user.password || '', fullName: user.fullName || '', email: user.email || '', role: user.role || 'Editor', status: user.status || 'Active' });
  const [show, setShow] = useState(false);
  const [chPw, setChPw] = useState(isNew);
  const [err, setErr] = useState('');
  const isSelf = !isNew && user.id === currentUser.id;
  const isOwner = !isNew && user.role === 'Owner';
  const roles = ROLES_LIST.filter(r => !(r === 'Owner' && !isOwner));

  const submit = async () => {
    if (!form.username.trim() || !form.fullName.trim()) return setErr('Username and full name are required');
    const dup = users.find(u => u.username.toLowerCase() === form.username.trim().toLowerCase() && u.id !== form.id);
    if (dup) return setErr('Username already taken');
    if (chPw && (!form.password || form.password.length < 6)) return setErr('Password must be at least 6 characters');
    setErr('');
    const payload = { ...form };
    if (!chPw && !isNew) {
      const orig = users.find(u => u.id === form.id);
      payload.password = orig ? orig.password : payload.password;
    }
    await onSave(payload, isNew);
  };

  return (
    <div className="fixed inset-0 bg-slate-900/50 flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-md p-5 border-t-4 border-red-700" onClick={e => e.stopPropagation()}>
        <h3 className="font-bold text-lg mb-4 text-slate-900 flex items-center gap-2">
          <UserPlus className="w-5 h-5 text-red-700" />{isNew ? 'Add Team Member' : 'Edit Team Member'}
        </h3>
        <div className="space-y-3">
          <SetupField label="Full Name" value={form.fullName} onChange={v => setForm({ ...form, fullName: v })} placeholder="Jane Doe" />
          <SetupField label="Email" value={form.email} onChange={v => setForm({ ...form, email: v })} placeholder="jane@texasfinserv.com" type="email" />
          <SetupField label="Username" value={form.username} onChange={v => setForm({ ...form, username: v })} placeholder="jdoe" />
          {!isNew && !chPw ? (
            <button onClick={() => setChPw(true)} className="text-xs text-red-700 hover:underline flex items-center gap-1">
              <Key className="w-3 h-3" /> Change password
            </button>
          ) : (
            <PwField label={isNew ? 'Password' : 'New Password'} value={form.password} onChange={v => setForm({ ...form, password: v })} show={show} setShow={setShow} placeholder="Min 6 characters" />
          )}
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">Role</label>
            <select value={form.role} onChange={e => setForm({ ...form, role: e.target.value })} disabled={isSelf || isOwner} className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 bg-white disabled:bg-slate-50 disabled:text-slate-500">
              {roles.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
            <div className="text-xs text-slate-500 mt-1">{ROLE_DESC[form.role]}</div>
            {(isSelf || isOwner) && <div className="text-xs text-amber-700 mt-1">{isOwner ? 'Owner role cannot be changed' : 'You cannot change your own role'}</div>}
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">Status</label>
            <select value={form.status} onChange={e => setForm({ ...form, status: e.target.value })} disabled={isSelf || isOwner} className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 bg-white disabled:bg-slate-50 disabled:text-slate-500">
              <option value="Active">Active</option>
              <option value="Inactive">Inactive (cannot log in)</option>
            </select>
          </div>
          {err && <Errs msg={err} />}
        </div>
        <div className="flex justify-end gap-2 mt-5">
          <button onClick={onClose} className="px-3 py-2 text-sm rounded-lg border border-slate-200 hover:bg-slate-50">Cancel</button>
          <button onClick={submit} className="px-4 py-2 text-sm rounded-lg bg-red-700 text-white hover:bg-red-800 font-medium shadow-sm">
            {isNew ? 'Create User' : 'Save Changes'}
          </button>
        </div>
      </div>
    </div>
  );
}

// AUDIT VIEW
function AuditView({ auditLog, users }) {
  const [fu, setFu] = useState('all');
  const [fa, setFa] = useState('all');
  const [s, setS] = useState('');
  const sorted = [...auditLog].reverse();
  const filtered = sorted.filter(e => {
    if (fu !== 'all' && e.userId !== fu) return false;
    if (fa !== 'all' && !e.action.includes(fa)) return false;
    if (s) {
      const q = s.toLowerCase();
      if (!(e.userName + ' ' + e.action + ' ' + e.targetName).toLowerCase().includes(q)) return false;
    }
    return true;
  });

  return (
    <div className="space-y-4">
      <div>
        <div className="flex items-center gap-2 mb-1">
          <Activity className="w-5 h-5 text-red-700" />
          <h2 className="text-xl font-bold text-slate-900">Audit Trail</h2>
        </div>
        <p className="text-sm text-slate-500">All changes logged with user, timestamp, and details</p>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-3 flex items-center gap-2 flex-wrap shadow-sm">
        <Filter className="w-4 h-4 text-slate-400" />
        <select value={fu} onChange={e => setFu(e.target.value)} className="text-sm border border-slate-200 rounded-lg px-2 py-1.5 bg-white">
          <option value="all">All users</option>
          {users.map(u => <option key={u.id} value={u.id}>{u.fullName || u.username}</option>)}
        </select>
        <select value={fa} onChange={e => setFa(e.target.value)} className="text-sm border border-slate-200 rounded-lg px-2 py-1.5 bg-white">
          <option value="all">All actions</option>
          <option value="client">Clients</option>
          <option value="user">Users</option>
          <option value="login">Logins</option>
          <option value="created">Created</option>
          <option value="updated">Updated</option>
          <option value="deleted">Deleted</option>
        </select>
        <input value={s} onChange={e => setS(e.target.value)} placeholder="Search..." className="text-sm border border-slate-200 rounded-lg px-2 py-1.5 bg-white flex-1 min-w-[150px]" />
        <span className="text-xs text-slate-400">{filtered.length} entries</span>
      </div>
      <div className="space-y-2">
        {filtered.length === 0 ? (
          <div className="bg-white rounded-xl border border-slate-200 p-8 text-center text-slate-400 text-sm">No audit entries match your filters.</div>
        ) : filtered.slice(0, 200).map(e => <AuditEntry key={e.id} entry={e} />)}
        {filtered.length > 200 && <div className="text-center text-xs text-slate-500 py-2">Showing 200 most recent.</div>}
      </div>
    </div>
  );
}

function AuditEntry({ entry }) {
  const [exp, setExp] = useState(false);
  const dateStr = new Date(entry.timestamp).toLocaleString();
  const ac = { 'client_created': 'bg-emerald-100 text-emerald-700', 'client_updated': 'bg-blue-100 text-blue-900', 'client_deleted': 'bg-rose-100 text-rose-700', 'relationship_created': 'bg-violet-100 text-violet-700', 'relationship_deleted': 'bg-rose-100 text-rose-700', 'user_created': 'bg-emerald-100 text-emerald-700', 'user_updated': 'bg-blue-100 text-blue-900', 'user_deleted': 'bg-rose-100 text-rose-700', 'login': 'bg-slate-100 text-slate-700', 'logout': 'bg-slate-100 text-slate-700', 'system_initialized': 'bg-amber-100 text-amber-700', 'import_completed': 'bg-emerald-100 text-emerald-700' };
  const cls = ac[entry.action] || 'bg-slate-100 text-slate-700';
  const has = entry.changes && entry.changes.length > 0;
  return (
    <div className="bg-white rounded-lg border border-slate-200 shadow-sm overflow-hidden">
      <button onClick={() => has && setExp(!exp)} className="w-full text-left px-4 py-3 hover:bg-slate-50 flex items-start gap-3">
        <div className="w-8 h-8 rounded-full bg-gradient-to-br from-red-700 to-blue-900 text-white flex items-center justify-center text-xs font-bold flex-shrink-0">{entry.userName.charAt(0).toUpperCase()}</div>
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-bold text-sm text-slate-900">{entry.userName}</span>
            <span className="text-xs text-slate-500">({entry.userRole})</span>
            <span className={'text-xs font-bold px-2 py-0.5 rounded ' + cls}>{entry.action.replace(/_/g, ' ')}</span>
          </div>
          {entry.targetName && <div className="text-sm text-slate-700 mt-0.5">{entry.targetType}: <strong>{entry.targetName}</strong></div>}
          <div className="text-xs text-slate-500 mt-0.5">{dateStr}</div>
        </div>
        {has && (
          <div className="flex items-center gap-1 text-xs text-slate-500 flex-shrink-0">
            <span>{entry.changes.length} change{entry.changes.length === 1 ? '' : 's'}</span>
            {exp ? <ChevronDown className="w-3 h-3" /> : <ChevronRight className="w-3 h-3" />}
          </div>
        )}
      </button>
      {exp && has && (
        <div className="border-t border-slate-100 px-4 py-3 bg-slate-50/50">
          <div className="space-y-1.5">
            {entry.changes.map((c, i) => (
              <div key={i} className="text-xs flex items-start gap-2 flex-wrap">
                <span className="font-bold text-slate-700 w-32 flex-shrink-0">{c.field}:</span>
                <span className="text-rose-600 line-through">{c.from || '(empty)'}</span>
                <ChevronRight className="w-3 h-3 text-slate-400 flex-shrink-0" />
                <span className="text-emerald-700 font-medium">{c.to || '(empty)'}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// DASHBOARD
function Dashboard({ stats, allClients, data, handleSelect, currentUser, exportData, lastBackup }) {
  const recent = [...allClients].sort((a, b) => (b.createdAt || '').localeCompare(a.createdAt || '')).slice(0, 6);
  const daysSince = lastBackup ? Math.floor((Date.now() - new Date(lastBackup).getTime()) / 86400000) : null;
  const stale = lastBackup === '' || (daysSince !== null && daysSince >= 1);
  const lastTxt = lastBackup ? new Date(lastBackup).toLocaleString() : 'never';
  return (
    <div className="space-y-6">
      <div className={'rounded-xl border p-4 flex items-start gap-3 flex-wrap ' + (stale ? 'bg-amber-50 border-amber-300' : 'bg-emerald-50 border-emerald-200')}>
        <div className={'w-10 h-10 rounded-lg flex items-center justify-center flex-shrink-0 ' + (stale ? 'bg-amber-100 text-amber-700' : 'bg-emerald-100 text-emerald-700')}>
          {stale ? <AlertCircle className="w-5 h-5" /> : <CheckCircle2 className="w-5 h-5" />}
        </div>
        <div className="flex-1 min-w-[200px]">
          <div className={'font-bold text-sm ' + (stale ? 'text-amber-900' : 'text-emerald-900')}>
            {stale ? 'Back up before you close this chat' : 'Backup is current'}
          </div>
          <div className={'text-xs mt-0.5 ' + (stale ? 'text-amber-800' : 'text-emerald-800')}>
            Last full backup: <strong>{lastTxt}</strong>. The JSON includes all clients <em>and</em> user accounts — import it in a new chat to continue exactly where you left off.
          </div>
        </div>
        <button onClick={exportData} className={'px-4 py-2 rounded-lg text-white text-sm font-bold shadow-sm flex items-center gap-2 ' + (stale ? 'bg-amber-600 hover:bg-amber-700' : 'bg-emerald-600 hover:bg-emerald-700')}>
          <Download className="w-4 h-4" /> Export Full Backup
        </button>
      </div>
      <div>
        <div className="flex items-center gap-2 mb-1">
          <Star className="w-5 h-5 text-red-700" fill="currentColor" />
          <h2 className="text-xl font-bold text-slate-900">Welcome, {currentUser.fullName || currentUser.username}</h2>
        </div>
        <p className="text-sm text-slate-500">{ROLE_DESC[currentUser.role]}</p>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        <StatCard icon={<Building2 className="w-5 h-5" />} label="Entities" value={stats.totalEntities} color="red" />
        <StatCard icon={<User className="w-5 h-5" />} label="Individuals" value={stats.totalIndividuals} color="navy" />
        <StatCard icon={<Briefcase className="w-5 h-5" />} label="Active Payroll" value={stats.activePayroll} color="navy" />
        <StatCard icon={<AlertCircle className="w-5 h-5" />} label="Pending 941s" value={stats.pending941} color="amber" />
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-5 shadow-sm">
        <h3 className="font-bold mb-3 flex items-center gap-2 text-slate-900"><Clock className="w-4 h-4 text-amber-600" /> Recent Clients</h3>
        <div className="space-y-1">
          {recent.map(c => (
            <button key={c.id} onClick={() => handleSelect(c.id)} className="w-full text-left px-3 py-2 rounded-lg hover:bg-red-50/50 flex items-center gap-2 text-sm">
              {c.type === 'entity' ? <Building2 className="w-4 h-4 text-red-700" /> : <User className="w-4 h-4 text-blue-900" />}
              <span className="font-medium">{c.displayName}</span>
              <span className="text-xs text-slate-400 ml-auto">{c.status}</span>
            </button>
          ))}
          {allClients.length === 0 && <div className="text-sm text-slate-400 px-3 py-4">No clients yet.</div>}
        </div>
      </div>
    </div>
  );
}

// CLIENT LIST
function ClientList({ filterType, setFilterType, filterStatus, setFilterStatus, filteredClients, handleSelect }) {
  return (
    <div className="space-y-4">
      <div className="flex items-center gap-2 flex-wrap">
        <Filter className="w-4 h-4 text-slate-400" />
        <select value={filterType} onChange={e => setFilterType(e.target.value)} className="text-sm border border-slate-200 rounded-lg px-2 py-1.5 bg-white">
          <option value="all">All Types</option>
          <option value="entity">Entities</option>
          <option value="individual">Individuals</option>
        </select>
        <select value={filterStatus} onChange={e => setFilterStatus(e.target.value)} className="text-sm border border-slate-200 rounded-lg px-2 py-1.5 bg-white">
          <option value="all">All Statuses</option>
          {STATUS_OPTS.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
        <span className="text-xs text-slate-400 ml-2">{filteredClients.length} clients</span>
      </div>
      <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
        <table className="w-full text-sm">
          <thead className="bg-gradient-to-r from-red-50 to-blue-50 border-b border-slate-200">
            <tr>
              <th className="text-left px-4 py-2.5 font-bold text-slate-700">Type</th>
              <th className="text-left px-4 py-2.5 font-bold text-slate-700">Name</th>
              <th className="text-left px-4 py-2.5 font-bold text-slate-700">ID</th>
              <th className="text-left px-4 py-2.5 font-bold text-slate-700 hidden md:table-cell">Services</th>
              <th className="text-left px-4 py-2.5 font-bold text-slate-700">Status</th>
            </tr>
          </thead>
          <tbody>
            {filteredClients.map(c => (
              <tr key={c.id} onClick={() => handleSelect(c.id)} className="border-b border-slate-100 hover:bg-red-50/30 cursor-pointer">
                <td className="px-4 py-2">{c.type === 'entity' ? <Building2 className="w-4 h-4 text-red-700" /> : <User className="w-4 h-4 text-blue-900" />}</td>
                <td className="px-4 py-2 font-medium">{c.displayName}</td>
                <td className="px-4 py-2 text-slate-500 font-mono text-xs">{c.type === 'entity' ? c.ein : (c.ssn || c.itin)}</td>
                <td className="px-4 py-2 text-slate-500 text-xs hidden md:table-cell">{(c.services || []).slice(0, 2).join(', ')}{(c.services || []).length > 2 ? ' +' + ((c.services || []).length - 2) : ''}</td>
                <td className="px-4 py-2"><StatusBadge status={c.status} /></td>
              </tr>
            ))}
            {filteredClients.length === 0 && <tr><td colSpan="5" className="px-4 py-8 text-center text-slate-400">No clients match.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// DETAIL VIEW
function DetailView(p) {
  return (
    <div className="space-y-4">
      <DetailHeader sc={p.selectedClient} editing={p.editing} cu={p.currentUser} startEdit={p.startEdit} cancelEdit={p.cancelEdit} saveEdit={p.saveEdit} deleteClient={p.deleteClient} setShowLinkModal={p.setShowLinkModal} />
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        <div className="lg:col-span-2 space-y-4">
          <BasicInfo current={p.current} editing={p.editing} updateDraft={p.updateDraft} />
          <ContactSection current={p.current} editing={p.editing} updateDraft={p.updateDraft} />
          {p.current.type === 'individual' && <TaxSection current={p.current} editing={p.editing} updateDraft={p.updateDraft} />}
          <ServicesSection current={p.current} editing={p.editing} draft={p.draft} updateDraft={p.updateDraft} updateBillingCycle={p.updateBillingCycle} toggleService={p.toggleService} pickerOpen={p.pickerOpen} setPickerOpen={p.setPickerOpen} pickerCategory={p.pickerCategory} setPickerCategory={p.setPickerCategory} pendingService={p.pendingService} acceptSuggestion={p.acceptSuggestion} dismissSuggestions={p.dismissSuggestions} />
          <PaymentSection current={p.current} editing={p.editing} draft={p.draft} updatePeriod={p.updatePeriod} addPaymentPeriod={p.addPaymentPeriod} removePeriod={p.removePeriod} />
          <FilingsSection current={p.current} editing={p.editing} draft={p.draft} addTaxFiling={p.addTaxFiling} updateFiling={p.updateFiling} removeFiling={p.removeFiling} />
          <BankAccountsSection current={p.current} editing={p.editing} draft={p.draft} addBank={p.addBank} updateBank={p.updateBank} removeBank={p.removeBank} />
          {p.current.type === 'entity' && p.hasPayroll && <PayrollSection current={p.current} editing={p.editing} updateDraft={p.updateDraft} updateDeadline={p.updateDeadline} setF941Frequency={p.setF941Frequency} />}
          {p.current.type === 'entity' && p.hasPayroll && <EmployeesSection current={p.current} editing={p.editing} draft={p.draft} addEmployee={p.addEmployee} updateEmployee={p.updateEmployee} removeEmployee={p.removeEmployee} attachIdCard={p.attachIdCard} />}
          {p.current.type === 'entity' && !p.hasPayroll && <PayrollHint />}
          <NotesSection current={p.current} editing={p.editing} draft={p.draft} updateDraft={p.updateDraft} />
        </div>
        <div className="space-y-4">
          <RelatedSection sc={p.selectedClient} related={p.relatedToSelected} removeRelationship={p.removeRelationship} setShowLinkModal={p.setShowLinkModal} cu={p.currentUser} />
        </div>
      </div>
    </div>
  );
}

function DetailHeader({ sc, editing, cu, startEdit, cancelEdit, saveEdit, deleteClient, setShowLinkModal }) {
  const isE = sc.type === 'entity';
  const ab = isE ? 'bg-gradient-to-br from-red-100 to-red-200 text-red-700' : 'bg-gradient-to-br from-blue-100 to-blue-200 text-blue-900';
  const id = isE ? (sc.ein || 'No EIN') : (sc.ssn || sc.itin || 'No SSN/ITIN');
  return (
    <div className="bg-white rounded-xl border-l-4 border-red-700 border-t border-r border-b border-slate-200 shadow-sm p-5 flex items-start justify-between gap-4 flex-wrap">
      <div className="flex items-start gap-3">
        <div className={'w-14 h-14 rounded-xl flex items-center justify-center shadow-sm ' + ab}>
          {isE ? <Building2 className="w-7 h-7" /> : <User className="w-7 h-7" />}
        </div>
        <div>
          <div className="text-xs text-red-700 uppercase tracking-wider font-bold">{isE ? sc.entityType : 'Individual Client'}</div>
          <h2 className="text-2xl font-bold text-slate-900 leading-tight">{sc.displayName}</h2>
          <div className="flex items-center gap-3 mt-1.5">
            <StatusBadge status={sc.status} />
            <span className="text-xs text-slate-500 font-mono">{id}</span>
          </div>
        </div>
      </div>
      <div className="flex items-center gap-2">
        {!editing ? (
          <>
            {can(cu, 'canEdit') && <button onClick={() => setShowLinkModal(true)} className="px-3 py-2 text-sm rounded-lg border border-slate-200 hover:bg-slate-50 flex items-center gap-1"><Link2 className="w-4 h-4" /> Link</button>}
            {can(cu, 'canEdit') ? (
              <button onClick={startEdit} className="px-3 py-2 text-sm rounded-lg bg-red-700 text-white hover:bg-red-800 flex items-center gap-1 shadow-sm"><Edit2 className="w-4 h-4" /> Edit</button>
            ) : (
              <span className="px-3 py-2 text-sm text-slate-400 italic flex items-center gap-1"><Lock className="w-3 h-3" /> Read-only</span>
            )}
            {can(cu, 'canDelete') && <button onClick={deleteClient} className="px-2 py-2 text-sm rounded-lg border border-rose-200 text-rose-600 hover:bg-rose-50"><Trash2 className="w-4 h-4" /></button>}
          </>
        ) : (
          <>
            <button onClick={cancelEdit} className="px-3 py-2 text-sm rounded-lg border border-slate-200 hover:bg-slate-50 flex items-center gap-1"><X className="w-4 h-4" /> Cancel</button>
            <button onClick={saveEdit} className="px-3 py-2 text-sm rounded-lg bg-blue-900 text-white hover:bg-blue-950 flex items-center gap-1 shadow-sm"><Save className="w-4 h-4" /> Save</button>
          </>
        )}
      </div>
    </div>
  );
}

function BasicInfo({ current, editing, updateDraft }) {
  const icon = current.type === 'entity' ? <Building2 className="w-4 h-4 text-red-700" /> : <User className="w-4 h-4 text-blue-900" />;
  return (
    <Section title="Basic Information" icon={icon}>
      {current.type === 'entity' ? (
        <Grid>
          <Field label="Entity Name" value={current.name} editing={editing} onChange={v => updateDraft('name', v)} />
          <Field label="Entity Type" value={current.entityType} editing={editing} onChange={v => updateDraft('entityType', v)} options={ENTITY_TYPES} />
          <Field label="EIN" value={current.ein} editing={editing} onChange={v => updateDraft('ein', v)} mono />
          <Field label="State of Formation" value={current.stateOfFormation} editing={editing} onChange={v => updateDraft('stateOfFormation', v)} />
          <Field label="Formation Date" value={current.formationDate} editing={editing} onChange={v => updateDraft('formationDate', v)} type="date" />
          <Field label="Tax Year-End" value={current.taxYearEnd} editing={editing} onChange={v => updateDraft('taxYearEnd', v)} placeholder="12-31" />
          <Field label="NAICS Code" value={current.naicsCode} editing={editing} onChange={v => updateDraft('naicsCode', v)} />
          <Field label="States Filed" value={current.statesFiled} editing={editing} onChange={v => updateDraft('statesFiled', v)} placeholder="TX, CA" />
          <Field label="Status" value={current.status} editing={editing} onChange={v => updateDraft('status', v)} options={STATUS_OPTS} />
        </Grid>
      ) : (
        <Grid>
          <Field label="First Name" value={current.firstName} editing={editing} onChange={v => updateDraft('firstName', v)} />
          <Field label="Last Name" value={current.lastName} editing={editing} onChange={v => updateDraft('lastName', v)} />
          <Field label="SSN" value={current.ssn} editing={editing} onChange={v => updateDraft('ssn', v)} mono />
          <Field label="ITIN" value={current.itin} editing={editing} onChange={v => updateDraft('itin', v)} mono />
          <Field label="Date of Birth" value={current.dob} editing={editing} onChange={v => updateDraft('dob', v)} type="date" />
          <Field label="Filing Status" value={current.filingStatus} editing={editing} onChange={v => updateDraft('filingStatus', v)} options={FILING_STATUS} />
          <Field label="Visa Status" value={current.visaStatus} editing={editing} onChange={v => updateDraft('visaStatus', v)} options={VISA} />
          <Field label="Residency State" value={current.residencyState} editing={editing} onChange={v => updateDraft('residencyState', v)} />
          <Field label="Occupation" value={current.occupation} editing={editing} onChange={v => updateDraft('occupation', v)} />
          <Field label="Status" value={current.status} editing={editing} onChange={v => updateDraft('status', v)} options={STATUS_OPTS} />
        </Grid>
      )}
    </Section>
  );
}

function ContactSection({ current, editing, updateDraft }) {
  return (
    <Section title="Contact and Address" icon={<FileText className="w-4 h-4 text-blue-900" />}>
      <Grid>
        {current.type === 'individual' && <Field label="Email" value={current.email} editing={editing} onChange={v => updateDraft('email', v)} />}
        {current.type === 'individual' && <Field label="Phone" value={current.phone} editing={editing} onChange={v => updateDraft('phone', v)} />}
        {current.type === 'entity' && <Field label="Registered Address" value={current.registeredAddress} editing={editing} onChange={v => updateDraft('registeredAddress', v)} full multiline />}
        <Field label="Mailing Address" value={current.mailingAddress} editing={editing} onChange={v => updateDraft('mailingAddress', v)} full multiline />
      </Grid>
    </Section>
  );
}

function TaxSection({ current, editing, updateDraft }) {
  return (
    <Section title="Tax Profile" icon={<TrendingUp className="w-4 h-4 text-amber-600" />}>
      <Grid>
        <Field label="Spouse Name" value={current.spouseName} editing={editing} onChange={v => updateDraft('spouseName', v)} />
        <Field label="Spouse SSN" value={current.spouseSSN} editing={editing} onChange={v => updateDraft('spouseSSN', v)} mono />
        <Field label="Dependents" value={current.dependents} editing={editing} onChange={v => updateDraft('dependents', v)} multiline />
        <Field label="Prior Year AGI" value={current.priorYearAGI} editing={editing} onChange={v => updateDraft('priorYearAGI', v)} />
        <Field label="IRS IP PIN" value={current.irsIpPin} editing={editing} onChange={v => updateDraft('irsIpPin', v)} mono />
        <Field label="FBAR Required?" value={current.fbarRequired} editing={editing} onChange={v => updateDraft('fbarRequired', v)} options={['Yes', 'No', 'Unknown']} />
        <Field label="FATCA Required?" value={current.fatcaRequired} editing={editing} onChange={v => updateDraft('fatcaRequired', v)} options={['Yes', 'No', 'Unknown']} />
        <Field label="Treaty Country" value={current.treatyCountry} editing={editing} onChange={v => updateDraft('treatyCountry', v)} />
      </Grid>
    </Section>
  );
}

// SERVICES
function ServicesSection(p) {
  const groupBy = (svcs) => {
    const g = {};
    svcs.forEach(s => {
      let cn = 'Other';
      const ks = Object.keys(SERVICE_CATEGORIES);
      for (let i = 0; i < ks.length; i++) if (SERVICE_CATEGORIES[ks[i]].includes(s)) { cn = ks[i]; break; }
      if (!g[cn]) g[cn] = []; g[cn].push(s);
    });
    return g;
  };
  const ed = p.draft ? (p.draft.services || []) : [];
  const gs = groupBy(ed);
  const gc = groupBy(p.current.services || []);
  return (
    <Section title="Services Engaged" icon={<Briefcase className="w-4 h-4 text-red-700" />}>
      {p.editing ? (
        <div className="space-y-4">
          <div>
            <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">Add Services</label>
            <button onClick={() => { p.setPickerOpen(true); p.setPickerCategory(null); }} className="w-full text-left text-sm border border-slate-300 rounded-lg px-3 py-2.5 bg-white hover:border-red-400 flex items-center justify-between">
              <span className="text-slate-600 flex items-center gap-2"><Plus className="w-4 h-4 text-red-700" /> Choose by category</span>
              <ChevronRight className="w-4 h-4 text-slate-400" />
            </button>
            {p.pickerOpen && <ServicePicker clientType={p.current.type} draft={p.draft} category={p.pickerCategory} setCategory={p.setPickerCategory} onClose={() => { p.setPickerOpen(false); p.setPickerCategory(null); }} toggleService={p.toggleService} />}
          </div>
          {p.pendingService && <SuggestionBanner pending={p.pendingService} onAccept={p.acceptSuggestion} onDismiss={p.dismissSuggestions} />}
          {ed.length > 0 && (
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-2">Selected ({ed.length})</label>
              <div className="space-y-2">
                {Object.keys(gs).map(c => <CatGroup key={c} category={c} services={gs[c]} onRemove={p.toggleService} editable />)}
              </div>
            </div>
          )}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3 pt-3 border-t border-slate-100">
            <Field label="Total Fee ($)" value={p.draft.serviceFee} editing={true} onChange={v => p.updateDraft('serviceFee', v)} placeholder="0.00" />
            <div>
              <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">Billing Cycle</label>
              <select value={p.draft.billingCycle || ''} onChange={e => p.updateBillingCycle(e.target.value)} className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 bg-white">
                {BILLING_CYCLE.map(o => <option key={o} value={o}>{o}</option>)}
              </select>
            </div>
            <Field label="Engagement Letter" value={p.draft.engagementLetterDate} editing={true} onChange={v => p.updateDraft('engagementLetterDate', v)} type="date" />
            <Field label="Preparer" value={p.draft.preparer} editing={true} onChange={v => p.updateDraft('preparer', v)} />
            <Field label="Reviewer" value={p.draft.reviewer} editing={true} onChange={v => p.updateDraft('reviewer', v)} />
          </div>
        </div>
      ) : (
        <div className="space-y-4">
          {(p.current.services || []).length === 0 ? (
            <div className="text-sm text-slate-400 italic">No services selected.</div>
          ) : (
            <div className="space-y-2">{Object.keys(gc).map(c => <CatGroup key={c} category={c} services={gc[c]} />)}</div>
          )}
          <div className="grid grid-cols-2 md:grid-cols-3 gap-3 pt-3 border-t border-slate-100">
            <ReadField label="Total Fee" value={p.current.serviceFee ? '$' + p.current.serviceFee : ''} />
            <ReadField label="Billing Cycle" value={p.current.billingCycle} />
            <ReadField label="Preparer" value={p.current.preparer} />
            <ReadField label="Reviewer" value={p.current.reviewer} />
          </div>
        </div>
      )}
    </Section>
  );
}

function ServicePicker({ clientType, draft, category, setCategory, onClose, toggleService }) {
  const isAv = (s) => !((clientType === 'individual' && ENTITY_ONLY.includes(s)) || (clientType === 'entity' && INDIVIDUAL_ONLY.includes(s)));
  const getMx = (s) => {
    const g = MUTEX.find(g => g.services.includes(s));
    if (!g) return null;
    const sel = (draft.services || []);
    return g.services.find(x => x !== s && sel.includes(x));
  };
  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-2xl w-full max-w-2xl max-h-[80vh] flex flex-col overflow-hidden" onClick={e => e.stopPropagation()}>
        <div className="px-5 py-4 border-b-2 border-red-700 flex items-center justify-between bg-gradient-to-r from-red-50 to-blue-50">
          <div className="flex items-center gap-2">
            {category && <button onClick={() => setCategory(null)} className="p-1 hover:bg-white rounded"><ChevronRight className="w-4 h-4 rotate-180 text-slate-600" /></button>}
            <Briefcase className="w-5 h-5 text-red-700" />
            <h3 className="font-bold text-base text-slate-900">{category || 'Choose a Category'}</h3>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600 p-1 hover:bg-white rounded"><X className="w-5 h-5" /></button>
        </div>
        <div className="flex-1 overflow-auto p-5">
          {!category ? <CategoryGrid clientType={clientType} draft={draft} setCategory={setCategory} /> : <SvcList category={category} draft={draft} isAv={isAv} getMx={getMx} toggleService={toggleService} />}
        </div>
        <div className="px-5 py-3 border-t border-slate-200 bg-slate-50 flex items-center justify-between">
          <span className="text-xs text-slate-500">{(draft.services || []).length} selected</span>
          <button onClick={onClose} className="px-4 py-2 text-sm rounded-lg bg-red-700 text-white hover:bg-red-800 font-medium shadow-sm">Done</button>
        </div>
      </div>
    </div>
  );
}

function CategoryGrid({ clientType, draft, setCategory }) {
  const cc = { red: 'border-red-200 hover:border-red-500 bg-red-50/50', amber: 'border-amber-200 hover:border-amber-500 bg-amber-50/50', navy: 'border-blue-200 hover:border-blue-700 bg-blue-50/50', slate: 'border-slate-200 hover:border-slate-500 bg-slate-50/50' };
  const ic = { red: 'text-red-700 bg-red-100', amber: 'text-amber-600 bg-amber-100', navy: 'text-blue-900 bg-blue-100', slate: 'text-slate-600 bg-slate-100' };
  const icons = { 'Tax Services': <FileText className="w-5 h-5" />, 'Compliance': <CheckCircle2 className="w-5 h-5" />, 'Payroll': <Users className="w-5 h-5" />, 'Advisory and Other': <TrendingUp className="w-5 h-5" /> };
  return (
    <div>
      <p className="text-sm text-slate-600 mb-4">Pick a category for this {clientType} client:</p>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
        {Object.keys(SERVICE_CATEGORIES).map(cn => {
          const svcs = SERVICE_CATEGORIES[cn];
          const m = CAT_META[cn];
          const av = svcs.filter(s => !((clientType === 'individual' && ENTITY_ONLY.includes(s)) || (clientType === 'entity' && INDIVIDUAL_ONLY.includes(s))));
          const sn = (draft.services || []).filter(s => svcs.includes(s)).length;
          return (
            <button key={cn} onClick={() => setCategory(cn)} disabled={av.length === 0} className={'p-4 rounded-lg border-2 text-left transition-all disabled:opacity-40 disabled:cursor-not-allowed ' + cc[m.color]}>
              <div className="flex items-start justify-between mb-2">
                <div className={'w-10 h-10 rounded-lg flex items-center justify-center ' + ic[m.color]}>{icons[cn]}</div>
                {sn > 0 && <span className="text-xs font-bold bg-white px-2 py-0.5 rounded-full text-slate-700 border border-slate-200">{sn} selected</span>}
              </div>
              <div className="font-bold text-sm text-slate-900">{cn}</div>
              <div className="text-xs text-slate-600 mt-0.5">{m.desc}</div>
              <div className="text-xs text-slate-500 mt-2">{av.length} available</div>
            </button>
          );
        })}
      </div>
    </div>
  );
}

function SvcList({ category, draft, isAv, getMx, toggleService }) {
  const svcs = SERVICE_CATEGORIES[category];
  const sel = draft.services || [];
  const m = CAT_META[category];
  return (
    <div className="space-y-1">
      {svcs.map(s => {
        if (!isAv(s)) return null;
        return <SvcItem key={s} service={s} isSelected={sel.includes(s)} mutexConflict={getMx(s)} onToggle={() => toggleService(s)} color={m.color} />;
      })}
    </div>
  );
}

function SvcItem({ service, isSelected, mutexConflict, onToggle, color }) {
  const cb = { red: 'bg-red-50 border-red-300 text-red-900', amber: 'bg-amber-50 border-amber-300 text-amber-900', navy: 'bg-blue-50 border-blue-300 text-blue-900', slate: 'bg-slate-50 border-slate-300 text-slate-900' };
  return (
    <button onClick={onToggle} className={'w-full text-left px-3 py-2.5 rounded-lg border transition-all flex items-center justify-between gap-3 ' + (isSelected ? cb[color] : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50')}>
      <div className="flex items-center gap-3 min-w-0 flex-1">
        <div className={'w-5 h-5 rounded border-2 flex items-center justify-center flex-shrink-0 ' + (isSelected ? 'bg-current border-current' : 'border-slate-300')}>
          {isSelected && <CheckCircle2 className="w-3.5 h-3.5 text-white" />}
        </div>
        <div className="min-w-0">
          <div className="text-sm font-medium truncate">{service}</div>
          {mutexConflict && !isSelected && (
            <div className="text-xs text-amber-700 mt-0.5 flex items-center gap-1"><AlertCircle className="w-3 h-3" /> Will replace: {mutexConflict}</div>
          )}
        </div>
      </div>
    </button>
  );
}

function SuggestionBanner({ pending, onAccept, onDismiss }) {
  return (
    <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
      <div className="flex items-start gap-2 mb-2">
        <AlertCircle className="w-4 h-4 text-amber-600 mt-0.5 flex-shrink-0" />
        <div className="flex-1">
          <div className="text-sm font-bold text-amber-900">Recommended additions for "{pending.parent}"</div>
          <div className="text-xs text-amber-800 mt-0.5">These services typically go together:</div>
        </div>
        <button onClick={onDismiss} className="text-amber-600 hover:bg-amber-100 rounded p-0.5"><X className="w-4 h-4" /></button>
      </div>
      <div className="flex flex-wrap gap-1.5 ml-6">
        {pending.suggestions.map(s => (
          <button key={s} onClick={() => onAccept(s)} className="px-2.5 py-1 text-xs rounded-md bg-white border border-amber-300 text-amber-800 hover:bg-amber-100 font-medium flex items-center gap-1">
            <Plus className="w-3 h-3" /> {s}
          </button>
        ))}
      </div>
    </div>
  );
}

function CatGroup({ category, services, onRemove, editable }) {
  const m = CAT_META[category];
  const hc = { red: 'text-red-800 bg-red-50 border-red-200', amber: 'text-amber-800 bg-amber-50 border-amber-200', navy: 'text-blue-900 bg-blue-50 border-blue-200', slate: 'text-slate-700 bg-slate-50 border-slate-200' };
  const cc = { red: 'bg-red-100 text-red-800 border-red-200', amber: 'bg-amber-100 text-amber-800 border-amber-200', navy: 'bg-blue-100 text-blue-900 border-blue-200', slate: 'bg-slate-100 text-slate-800 border-slate-200' };
  const c = m ? m.color : 'slate';
  return (
    <div className={'rounded-lg border p-2.5 ' + hc[c]}>
      <div className="text-xs font-bold uppercase tracking-wide mb-2">{category}</div>
      <div className="flex flex-wrap gap-1.5">
        {services.map(s => (
          <span key={s} className={'inline-flex items-center gap-1.5 px-2.5 py-1 text-xs rounded-md border font-medium ' + cc[c]}>
            {s}
            {editable && <button onClick={() => onRemove(s)} className="hover:bg-white/50 rounded p-0.5"><X className="w-3 h-3" /></button>}
          </span>
        ))}
      </div>
    </div>
  );
}

// PAYMENT
function PaymentSection({ current, editing, draft, updatePeriod, addPaymentPeriod, removePeriod }) {
  const t = editing ? draft : current;
  const periods = t.paymentPeriods || [];
  const tc = periods.filter(p => p.status === 'Paid').reduce((s, p) => s + (parseFloat(p.amount) || 0), 0);
  const tp = periods.filter(p => ['Pending', 'Invoiced', 'Overdue'].includes(p.status)).reduce((s, p) => s + (parseFloat(p.amount) || 0), 0);
  const pc = periods.filter(p => p.status === 'Paid').length;

  return (
    <Section title="Payment Tracking" icon={<DollarSign className="w-4 h-4 text-red-700" />}>
      {periods.length === 0 ? (
        <div className="text-sm text-slate-400 italic text-center py-6">No payment periods. Click Edit to set up.</div>
      ) : (
        <div className="space-y-3">
          <div className="grid grid-cols-3 gap-2">
            <SumCard label="Collected" value={'$' + tc.toLocaleString()} sub={pc + ' of ' + periods.length} color="blue" />
            <SumCard label="Outstanding" value={'$' + tp.toLocaleString()} sub={(periods.length - pc) + ' pending'} color="red" />
            <SumCard label="Cycle" value={t.billingCycle || 'None'} sub={'Per fee: $' + (t.serviceFee || '0')} color="slate" />
          </div>
          <div className="border border-slate-200 rounded-lg overflow-hidden">
            <div className="grid grid-cols-12 gap-2 px-3 py-2 bg-gradient-to-r from-red-50 to-blue-50 border-b border-slate-200 text-xs font-bold text-slate-700 uppercase tracking-wide">
              <div className="col-span-3">Period</div>
              <div className="col-span-2">Status</div>
              <div className="col-span-2">Amount</div>
              <div className="col-span-2">Paid Date</div>
              <div className="col-span-2">Invoice #</div>
              <div className="col-span-1"></div>
            </div>
            {periods.map((pd, i) => <PeriodRow key={pd.id || i} period={pd} idx={i} editing={editing} onUpdate={updatePeriod} onRemove={removePeriod} />)}
          </div>
          {editing && (
            <button onClick={addPaymentPeriod} className="text-xs text-red-700 hover:bg-red-50 rounded-lg px-3 py-1.5 border border-dashed border-red-300 flex items-center gap-1">
              <Plus className="w-3 h-3" /> Add custom period
            </button>
          )}
        </div>
      )}
    </Section>
  );
}

function SumCard({ label, value, sub, color }) {
  const cc = { blue: 'bg-blue-50 border-blue-100 text-blue-900', red: 'bg-red-50 border-red-100 text-red-900', slate: 'bg-slate-50 border-slate-200 text-slate-800' };
  const lc = { blue: 'text-blue-900', red: 'text-red-700', slate: 'text-slate-600' };
  return (
    <div className={'border rounded-lg p-3 ' + cc[color]}>
      <div className={'text-xs font-bold uppercase tracking-wide mb-0.5 ' + lc[color]}>{label}</div>
      <div className="text-lg font-bold">{value}</div>
      <div className="text-xs opacity-70">{sub}</div>
    </div>
  );
}

function PeriodRow({ period, idx, editing, onUpdate, onRemove }) {
  let bg = '';
  if (period.status === 'Paid') bg = 'bg-blue-50/30';
  if (period.status === 'Overdue') bg = 'bg-red-50/30';
  return (
    <div className={'grid grid-cols-12 gap-2 px-3 py-2 items-center border-b border-slate-100 last:border-b-0 ' + bg}>
      <div className="col-span-3 text-sm font-medium text-slate-800">{period.label}</div>
      <div className="col-span-2">
        {editing ? (
          <select value={period.status} onChange={e => onUpdate(idx, 'status', e.target.value)} className="w-full text-xs border border-slate-300 rounded px-2 py-1 bg-white">
            {PAY_STATUS.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        ) : <PaymentBadge status={period.status} />}
      </div>
      <div className="col-span-2">
        {editing ? <input value={period.amount} onChange={e => onUpdate(idx, 'amount', e.target.value)} placeholder="0.00" className="w-full text-xs border border-slate-300 rounded px-2 py-1" /> : <span className="text-sm font-mono">{period.amount ? '$' + period.amount : '—'}</span>}
      </div>
      <div className="col-span-2">
        {editing ? <input type="date" value={period.paidDate} onChange={e => onUpdate(idx, 'paidDate', e.target.value)} className="w-full text-xs border border-slate-300 rounded px-2 py-1" /> : <span className="text-xs text-slate-600">{period.paidDate || '—'}</span>}
      </div>
      <div className="col-span-2">
        {editing ? <input value={period.invoiceNum} onChange={e => onUpdate(idx, 'invoiceNum', e.target.value)} placeholder="—" className="w-full text-xs border border-slate-300 rounded px-2 py-1 font-mono" /> : <span className="text-xs font-mono text-slate-600">{period.invoiceNum || '—'}</span>}
      </div>
      <div className="col-span-1 flex justify-end">
        {editing && <button onClick={() => onRemove(idx)} className="text-rose-500 hover:bg-rose-50 rounded p-1"><Trash2 className="w-3.5 h-3.5" /></button>}
      </div>
    </div>
  );
}

// FILINGS
function FilingsSection({ current, editing, draft, addTaxFiling, updateFiling, removeFiling }) {
  const filings = (editing ? draft : current).taxFilings || [];
  return (
    <Section title="Tax Filings Tracker" icon={<FileText className="w-4 h-4 text-red-700" />}>
      {editing && (
        <button onClick={addTaxFiling} className="mb-3 px-3 py-1.5 text-xs rounded-lg border border-red-200 text-red-700 hover:bg-red-50 flex items-center gap-1">
          <Plus className="w-3 h-3" /> Add Filing
        </button>
      )}
      <div className="space-y-2">
        {filings.length === 0 ? <div className="text-sm text-slate-400">No filings tracked yet.</div> : filings.map((f, i) => <FilingRow key={f.id || i} filing={f} idx={i} editing={editing} updateFiling={updateFiling} removeFiling={removeFiling} />)}
      </div>
    </Section>
  );
}

function FilingRow({ filing, idx, editing, updateFiling, removeFiling }) {
  if (editing) {
    return (
      <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/50">
        <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
          <Field label="Year" value={filing.year} editing={true} onChange={v => updateFiling(idx, 'year', v)} />
          <Field label="Form" value={filing.formType} editing={true} onChange={v => updateFiling(idx, 'formType', v)} />
          <Field label="Status" value={filing.status} editing={true} onChange={v => updateFiling(idx, 'status', v)} options={FILING_DEADLINE_STATUS} />
          <Field label="Due Date" value={filing.dueDate} editing={true} onChange={v => updateFiling(idx, 'dueDate', v)} type="date" />
          <Field label="Filed Date" value={filing.filedDate} editing={true} onChange={v => updateFiling(idx, 'filedDate', v)} type="date" />
          <Field label="Refund/Due" value={filing.refundOrDue} editing={true} onChange={v => updateFiling(idx, 'refundOrDue', v)} options={['Refund', 'Tax Due', 'Zero', 'N/A']} />
          <Field label="Amount" value={filing.amount} editing={true} onChange={v => updateFiling(idx, 'amount', v)} />
          <Field label="Extension?" value={filing.extensionFiled} editing={true} onChange={v => updateFiling(idx, 'extensionFiled', v)} options={['Yes', 'No']} />
        </div>
        <button onClick={() => removeFiling(idx)} className="text-xs text-rose-600 hover:underline mt-2">Remove</button>
      </div>
    );
  }
  return (
    <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/50">
      <div className="flex items-center justify-between gap-3 flex-wrap">
        <div className="flex items-center gap-3">
          <span className="font-bold text-sm">{filing.year} · {filing.formType}</span>
          <FilingStatusBadge status={filing.status} />
        </div>
        <div className="text-xs text-slate-500 flex items-center gap-3">
          {filing.dueDate && <span>Due: {filing.dueDate}</span>}
          {filing.filedDate && <span>Filed: {filing.filedDate}</span>}
          {filing.amount && <span>${filing.amount}</span>}
        </div>
      </div>
    </div>
  );
}

// PAYROLL
function EmployeeCountField({ payroll }) {
  const list = (payroll && payroll.employeeList) || [];
  const total = list.length;
  const active = list.filter(e => (e.status || 'Active') === 'Active').length;
  return (
    <div>
      <div className="text-xs font-bold text-slate-500 uppercase tracking-wide mb-1">
        # Employees <span className="text-slate-400 normal-case font-normal">(auto from list)</span>
      </div>
      <div className="text-sm text-slate-900 font-medium min-h-[20px]">
        {total === 0 ? <span className="text-slate-300 font-normal">— add employees below</span> : (
          <span>{active} active{active !== total ? <span className="text-slate-500"> · {total} total</span> : null}</span>
        )}
      </div>
    </div>
  );
}

function PayrollSection({ current, editing, updateDraft, updateDeadline, setF941Frequency }) {
  return (
    <Section title="Payroll Setup and Compliance" icon={<Users className="w-4 h-4 text-blue-900" />}>
      <div className="bg-blue-50/50 rounded-lg p-3 mb-4 border border-blue-100">
        <div className="text-xs font-bold text-blue-900 uppercase tracking-wide mb-2.5">Payroll Setup</div>
        <Grid>
          <Field label="Status" value={current.payroll && current.payroll.payrollStatus} editing={editing} onChange={v => updateDraft('payroll.payrollStatus', v)} options={PAYROLL_STATUS} />
          <Field label="Frequency" value={current.payroll && current.payroll.frequency} editing={editing} onChange={v => updateDraft('payroll.frequency', v)} options={PAYROLL_FREQ} />
          <Field label="Provider" value={current.payroll && current.payroll.provider} editing={editing} onChange={v => updateDraft('payroll.provider', v)} options={PAYROLL_PROVIDER} />
          <EmployeeCountField payroll={current.payroll} />
          <Field label="States Registered" value={current.payroll && current.payroll.statesRegistered} editing={editing} onChange={v => updateDraft('payroll.statesRegistered', v)} placeholder="TX, CA" />
          <Field label="Federal EIN" value={current.payroll && current.payroll.federalEinStatus} editing={editing} onChange={v => updateDraft('payroll.federalEinStatus', v)} options={['Active', 'Pending', 'Closed']} />
          <Field label="State W/H ID" value={current.payroll && current.payroll.stateWithholdingId} editing={editing} onChange={v => updateDraft('payroll.stateWithholdingId', v)} mono />
          <Field label="State SUTA ID" value={current.payroll && current.payroll.stateUnemploymentId} editing={editing} onChange={v => updateDraft('payroll.stateUnemploymentId', v)} mono />
        </Grid>
      </div>
      <DeadlineMatrix deadlines={(current.payroll && current.payroll.deadlines) || []} editing={editing} onUpdate={updateDeadline} f941Frequency={(current.payroll && current.payroll.f941Frequency) || 'Quarterly'} onSetF941Freq={setF941Frequency} />
    </Section>
  );
}

function PayrollHint() {
  return (
    <div className="bg-slate-50 border border-slate-200 rounded-xl p-4 text-sm text-slate-600 flex items-start gap-2">
      <Users className="w-4 h-4 text-slate-400 mt-0.5 flex-shrink-0" />
      <div><strong className="text-slate-700">Payroll module hidden.</strong> Add a payroll service to enable.</div>
    </div>
  );
}

function DeadlineMatrix({ deadlines, editing, onUpdate, f941Frequency, onSetF941Freq }) {
  const freq = f941Frequency || 'Quarterly';
  const groups = {};
  deadlines.forEach((d, i) => {
    if (!groups[d.name]) groups[d.name] = [];
    groups[d.name].push({ ...d, originalIdx: i });
  });
  return (
    <div>
      <div className="flex items-center gap-2 mb-3">
        <Calendar className="w-4 h-4 text-blue-900" />
        <h4 className="font-bold text-sm text-slate-800">Compliance Deadlines</h4>
      </div>
      <div className="space-y-2">
        {Object.keys(groups).map(fn => {
          if (fn === 'Form 941') {
            const isMonthly = freq === 'Monthly';
            const items = groups[fn].filter(d => isMonthly ? d.freq === 'Monthly' : d.freq !== 'Monthly');
            return <DeadlineGroup key={fn} formName={fn} items={items} editing={editing} onUpdate={onUpdate} freqToggle={{ value: freq, onChange: onSetF941Freq }} />;
          }
          return <DeadlineGroup key={fn} formName={fn} items={groups[fn]} editing={editing} onUpdate={onUpdate} />;
        })}
      </div>
    </div>
  );
}

function DeadlineGroup({ formName, items, editing, onUpdate, freqToggle }) {
  const gc = items.length > 6 ? 'grid-cols-3 md:grid-cols-6' : (items.length > 1 ? 'grid-cols-2 md:grid-cols-4' : 'grid-cols-1');
  const subtitle = freqToggle ? (freqToggle.value === 'Monthly' ? '941 · Monthly deposits' : '941 · Quarterly filing') : ((items[0] && items[0].form) + (items[0] && items[0].dueDate ? ' · Due ' + items[0].dueDate : ''));
  return (
    <div className="border border-slate-200 rounded-lg overflow-hidden">
      <div className="bg-gradient-to-r from-red-50 to-blue-50 px-3 py-2 border-b border-slate-200 flex items-center justify-between gap-2 flex-wrap">
        <div className="font-bold text-sm text-slate-800">{formName}</div>
        <div className="flex items-center gap-2">
          {freqToggle && editing && (
            <div className="flex items-center rounded-md border border-slate-300 overflow-hidden text-xs">
              <button type="button" onClick={() => freqToggle.onChange('Quarterly')} className={'px-2 py-1 font-bold ' + (freqToggle.value !== 'Monthly' ? 'bg-blue-900 text-white' : 'bg-white text-slate-600 hover:bg-slate-50')}>Quarterly</button>
              <button type="button" onClick={() => freqToggle.onChange('Monthly')} className={'px-2 py-1 font-bold ' + (freqToggle.value === 'Monthly' ? 'bg-blue-900 text-white' : 'bg-white text-slate-600 hover:bg-slate-50')}>Monthly</button>
            </div>
          )}
          <div className="text-xs text-slate-500">{subtitle}</div>
        </div>
      </div>
      {items.length === 0 ? (
        <div className="p-3 text-xs text-slate-400 italic">No periods yet — switch to Edit to set up.</div>
      ) : (
        <div className={'grid gap-1 p-2 ' + gc}>
          {items.map(d => <DeadlineCell key={d.id} deadline={d} editing={editing} onUpdate={onUpdate} />)}
        </div>
      )}
    </div>
  );
}

function DeadlineCell({ deadline, editing, onUpdate }) {
  let bg = 'bg-white border-slate-200';
  let icon = <Circle className="w-3.5 h-3.5 text-slate-400" />;
  if (['Filed', 'Completed', 'Paid'].includes(deadline.status)) { bg = 'bg-blue-50 border-blue-200'; icon = <CheckCircle2 className="w-3.5 h-3.5 text-blue-700" />; }
  else if (['In Progress', 'Pending Payment'].includes(deadline.status)) { bg = 'bg-amber-50 border-amber-200'; icon = <Clock className="w-3.5 h-3.5 text-amber-600" />; }
  else if (deadline.status === 'Extended') { bg = 'bg-orange-50 border-orange-200'; icon = <Clock className="w-3.5 h-3.5 text-orange-600" />; }
  const idx = deadline.originalIdx;
  return (
    <div className={'rounded-lg border p-2 ' + bg}>
      <div className="flex items-center justify-between mb-1.5">
        <span className="text-xs font-bold text-slate-700">{deadline.period}</span>
        {icon}
      </div>
      {editing ? (
        <div className="space-y-1">
          <select value={deadline.status} onChange={e => onUpdate(idx, 'status', e.target.value)} className="w-full text-xs border border-slate-300 rounded px-1.5 py-1 bg-white">
            {COMPLIANCE_STATUS.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <input type="date" value={deadline.filedDate} onChange={e => onUpdate(idx, 'filedDate', e.target.value)} className="w-full text-xs border border-slate-300 rounded px-1.5 py-1" />
          <input value={deadline.amount} onChange={e => onUpdate(idx, 'amount', e.target.value)} placeholder="$ Amount" className="w-full text-xs border border-slate-300 rounded px-1.5 py-1" />
        </div>
      ) : (
        <div className="space-y-0.5">
          <div className="text-xs font-medium text-slate-700">{deadline.status}</div>
          {deadline.filedDate && <div className="text-xs text-slate-500">Filed: {deadline.filedDate}</div>}
          {deadline.amount && <div className="text-xs text-slate-500">${deadline.amount}</div>}
        </div>
      )}
    </div>
  );
}

// NOTES & RELATED
function NotesSection({ current, editing, draft, updateDraft }) {
  return (
    <Section title="Notes" icon={<FileText className="w-4 h-4 text-slate-500" />}>
      {editing ? (
        <textarea value={draft.notes} onChange={e => updateDraft('notes', e.target.value)} rows={4} className="w-full text-sm border border-slate-300 rounded-lg p-3 focus:outline-none focus:ring-2 focus:ring-red-600" placeholder="Internal notes..." />
      ) : (
        <div className="text-sm text-slate-600 whitespace-pre-wrap min-h-[40px]">{current.notes || <span className="text-slate-400 italic">No notes</span>}</div>
      )}
    </Section>
  );
}

function RelatedSection({ sc, related, removeRelationship, setShowLinkModal, cu }) {
  const title = sc.type === 'entity' ? 'Related Individuals' : 'Related Entities';
  return (
    <Section title={title} icon={<Link2 className="w-4 h-4 text-blue-900" />}>
      {can(cu, 'canEdit') && (
        <button onClick={() => setShowLinkModal(true)} className="w-full mb-3 px-3 py-2 text-xs rounded-lg border border-dashed border-red-300 text-red-700 hover:bg-red-50 flex items-center justify-center gap-1">
          <Link2 className="w-3 h-3" /> Add Relationship
        </button>
      )}
      {related.length === 0 ? (
        <div className="text-sm text-slate-400 text-center py-3">No relationships yet.</div>
      ) : (
        <div className="space-y-2">
          {related.map(r => (
            <div key={r.id} className="border border-slate-200 rounded-lg p-3 hover:border-red-300 transition-colors">
              <div className="flex items-start gap-2">
                {r.other.type === 'entity' ? <Building2 className="w-4 h-4 text-red-700 mt-0.5" /> : <User className="w-4 h-4 text-blue-900 mt-0.5" />}
                <div className="flex-1 min-w-0">
                  <div className="font-medium text-sm truncate">{r.other.displayName}</div>
                  <div className="text-xs text-slate-500">{r.role}{r.ownership ? ' · ' + r.ownership + '%' : ''}</div>
                </div>
              </div>
              {can(cu, 'canEdit') && <button onClick={() => removeRelationship(r.id)} className="text-xs text-rose-500 hover:underline mt-1">Remove</button>}
            </div>
          ))}
        </div>
      )}
    </Section>
  );
}

// MODALS
function LinkModal({ selectedClient, data, linkRole, setLinkRole, linkOwnership, setLinkOwnership, linkTargetId, setLinkTargetId, addRelationship, onClose }) {
  const cands = selectedClient.type === 'entity' ? data.individuals : data.entities;
  const filt = cands.filter(c => !data.relationships.some(r => (r.entityId === selectedClient.id && r.individualId === c.id) || (r.individualId === selectedClient.id && r.entityId === c.id)));
  return (
    <div className="fixed inset-0 bg-slate-900/50 flex items-center justify-center z-50 p-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-md p-5 border-t-4 border-red-700" onClick={e => e.stopPropagation()}>
        <h3 className="font-bold mb-4 text-slate-900">Add Relationship</h3>
        <div className="space-y-3">
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1 uppercase tracking-wide">{selectedClient.type === 'entity' ? 'Link Individual' : 'Link Entity'}</label>
            <select value={linkTargetId} onChange={e => setLinkTargetId(e.target.value)} className="w-full text-sm border border-slate-200 rounded-lg px-2 py-2 bg-white">
              <option value="">— Select —</option>
              {filt.map(c => {
                const n = c.type === 'entity' ? c.name : ((c.firstName || '') + ' ' + (c.lastName || '')).trim();
                return <option key={c.id} value={c.id}>{n}</option>;
              })}
            </select>
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-700 mb-1 uppercase tracking-wide">Role</label>
            <select value={linkRole} onChange={e => setLinkRole(e.target.value)} className="w-full text-sm border border-slate-200 rounded-lg px-2 py-2 bg-white">
              {REL_ROLES.map(r => <option key={r} value={r}>{r}</option>)}
            </select>
          </div>
          {(linkRole === 'Owner' || linkRole === 'Co-Owner') && (
            <div>
              <label className="block text-xs font-bold text-slate-700 mb-1 uppercase tracking-wide">Ownership %</label>
              <input value={linkOwnership} onChange={e => setLinkOwnership(e.target.value)} placeholder="50" className="w-full text-sm border border-slate-200 rounded-lg px-2 py-2" />
            </div>
          )}
        </div>
        <div className="flex justify-end gap-2 mt-5">
          <button onClick={onClose} className="px-3 py-2 text-sm rounded-lg border border-slate-200 hover:bg-slate-50">Cancel</button>
          <button onClick={addRelationship} disabled={!linkTargetId} className="px-3 py-2 text-sm rounded-lg bg-red-700 text-white hover:bg-red-800 disabled:opacity-50">Add Link</button>
        </div>
      </div>
    </div>
  );
}

function ConfirmDialog({ action, onClose }) {
  const bc = action.danger ? 'bg-rose-600 hover:bg-rose-700' : 'bg-red-700 hover:bg-red-800';
  const ib = action.danger ? 'bg-rose-100 text-rose-600' : 'bg-blue-100 text-blue-900';
  return (
    <div className="fixed inset-0 bg-slate-900/60 flex items-center justify-center z-[60] p-4" onClick={onClose}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-sm p-5" onClick={e => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-3">
          <div className={'w-10 h-10 rounded-full flex items-center justify-center flex-shrink-0 ' + ib}>
            {action.danger ? <AlertCircle className="w-5 h-5" /> : <CheckCircle2 className="w-5 h-5" />}
          </div>
          <div className="flex-1">
            <h3 className="font-bold text-base text-slate-900">{action.title}</h3>
            <p className="text-sm text-slate-600 mt-1">{action.message}</p>
          </div>
        </div>
        <div className="flex justify-end gap-2 mt-4">
          {action.danger && <button onClick={onClose} className="px-4 py-2 text-sm rounded-lg border border-slate-200 hover:bg-slate-50">Cancel</button>}
          <button onClick={action.onConfirm} className={'px-4 py-2 text-sm rounded-lg text-white font-medium ' + bc}>{action.confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

function ImportModal(p) {
  return (
    <div className="fixed inset-0 bg-slate-900/50 flex items-center justify-center z-50 p-4" onClick={p.onClose}>
      <div className="bg-white rounded-xl shadow-xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col border-t-4 border-red-700" onClick={e => e.stopPropagation()}>
        <div className="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <Upload className="w-5 h-5 text-red-700" />
            <h3 className="font-bold text-lg text-slate-900">Import Clients</h3>
            <ImportSteps importStep={p.importStep} />
          </div>
          <button onClick={p.onClose} className="text-slate-400 hover:text-slate-600"><X className="w-5 h-5" /></button>
        </div>
        <div className="flex-1 overflow-auto p-5">
          {p.importStep === 'upload' && <ImportUpload importType={p.importType} setImportType={p.setImportType} handleImportFile={p.handleImportFile} />}
          {p.importStep === 'map' && <ImportMap importHeaders={p.importHeaders} importRows={p.importRows} importMapping={p.importMapping} setImportMapping={p.setImportMapping} importType={p.importType} />}
          {p.importStep === 'preview' && <ImportPreview importPreview={p.importPreview} importType={p.importType} />}
        </div>
        <div className="px-5 py-3 border-t border-slate-200 flex items-center justify-between gap-2 bg-slate-50">
          <button onClick={p.resetImport} className="text-xs text-slate-500 hover:underline">Start over</button>
          <div className="flex items-center gap-2">
            <button onClick={p.onClose} className="px-3 py-2 text-sm rounded-lg border border-slate-200 hover:bg-white">Cancel</button>
            {p.importStep === 'map' && (
              <button onClick={p.buildPreview} disabled={!Object.values(p.importMapping).some(v => v && v !== '__skip__')} className="px-3 py-2 text-sm rounded-lg bg-red-700 text-white hover:bg-red-800 disabled:opacity-50">Preview →</button>
            )}
            {p.importStep === 'preview' && (
              <>
                <button onClick={() => p.setImportStep('map')} className="px-3 py-2 text-sm rounded-lg border border-slate-200 hover:bg-white">← Back</button>
                <button onClick={p.confirmImport} className="px-3 py-2 text-sm rounded-lg bg-blue-900 text-white hover:bg-blue-950 flex items-center gap-1"><CheckCircle2 className="w-4 h-4" /> Confirm ({p.importPreview.length})</button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

function ImportSteps({ importStep }) {
  const sc = (s) => importStep === s ? 'bg-red-100 text-red-700 font-bold' : 'text-slate-400';
  return (
    <div className="flex items-center gap-1 ml-4 text-xs">
      <span className={'px-2 py-1 rounded ' + sc('upload')}>1. Upload</span>
      <ChevronRight className="w-3 h-3 text-slate-300" />
      <span className={'px-2 py-1 rounded ' + sc('map')}>2. Map</span>
      <ChevronRight className="w-3 h-3 text-slate-300" />
      <span className={'px-2 py-1 rounded ' + sc('preview')}>3. Confirm</span>
    </div>
  );
}

function ImportUpload({ importType, setImportType, handleImportFile }) {
  const ec = importType === 'entity' ? 'border-red-700 bg-red-50' : 'border-slate-200 hover:border-slate-300';
  const ic = importType === 'individual' ? 'border-blue-900 bg-blue-50' : 'border-slate-200 hover:border-slate-300';
  return (
    <div className="space-y-4">
      <div>
        <label className="block text-sm font-bold text-slate-700 mb-2">What are you importing?</label>
        <div className="grid grid-cols-2 gap-3">
          <button onClick={() => setImportType('entity')} className={'p-4 rounded-lg border-2 text-left ' + ec}>
            <Building2 className="w-5 h-5 text-red-700 mb-2" />
            <div className="font-bold text-sm">Entity Clients</div>
            <div className="text-xs text-slate-500 mt-1">LLCs, Corps, Partnerships</div>
          </button>
          <button onClick={() => setImportType('individual')} className={'p-4 rounded-lg border-2 text-left ' + ic}>
            <User className="w-5 h-5 text-blue-900 mb-2" />
            <div className="font-bold text-sm">Individual Clients</div>
            <div className="text-xs text-slate-500 mt-1">1040 / 1040-NR</div>
          </button>
        </div>
      </div>
      <div className="border-2 border-dashed border-slate-300 rounded-lg p-8 text-center hover:border-red-400 transition-colors">
        <Upload className="w-10 h-10 text-slate-400 mx-auto mb-3" />
        <input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => { const f = e.target.files && e.target.files[0]; if (f) handleImportFile(f); e.target.value = ''; }} className="block mx-auto text-sm text-slate-700 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-medium file:bg-red-700 file:text-white hover:file:bg-red-800 file:cursor-pointer cursor-pointer" />
        <div className="text-xs text-slate-500 mt-3">.xlsx, .xls, or .csv</div>
      </div>
    </div>
  );
}

function ImportMap({ importHeaders, importRows, importMapping, setImportMapping, importType }) {
  const fields = importType === 'entity' ? ENT_FIELDS : IND_FIELDS;
  return (
    <div className="space-y-3">
      <div className="text-sm text-slate-600">Found <strong>{importRows.length} rows</strong>, <strong>{importHeaders.length} columns</strong>.</div>
      <div className="border border-slate-200 rounded-lg overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gradient-to-r from-red-50 to-blue-50 border-b border-slate-200">
            <tr>
              <th className="text-left px-3 py-2 font-bold text-slate-600 w-2/5">Your Column</th>
              <th className="text-left px-3 py-2 font-bold text-slate-600 w-2/5">Maps To</th>
              <th className="text-left px-3 py-2 font-bold text-slate-600 w-1/5">Sample</th>
            </tr>
          </thead>
          <tbody>
            {importHeaders.map(h => (
              <tr key={h} className="border-b border-slate-100">
                <td className="px-3 py-2 font-medium">{h}</td>
                <td className="px-3 py-2">
                  <select value={importMapping[h] || '__skip__'} onChange={e => setImportMapping(prev => ({ ...prev, [h]: e.target.value }))} className="w-full text-sm border border-slate-200 rounded px-2 py-1 bg-white">
                    <option value="__skip__">— Skip —</option>
                    {fields.map(f => <option key={f.key} value={f.key}>{f.label}</option>)}
                  </select>
                </td>
                <td className="px-3 py-2 text-xs text-slate-500 truncate max-w-[200px]">{importRows[0] && importRows[0][h] ? importRows[0][h] : '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ImportPreview({ importPreview, importType }) {
  const tl = importType === 'entity' ? 'entities' : 'individuals';
  return (
    <div className="space-y-3">
      <div className="text-sm text-slate-600">Ready to import <strong>{importPreview.length} {tl}</strong>.</div>
      <div className="border border-slate-200 rounded-lg overflow-hidden">
        <table className="w-full text-sm">
          <thead className="bg-gradient-to-r from-red-50 to-blue-50 border-b border-slate-200">
            <tr>
              <th className="text-left px-3 py-2 font-bold text-slate-600">#</th>
              <th className="text-left px-3 py-2 font-bold text-slate-600">Name</th>
              <th className="text-left px-3 py-2 font-bold text-slate-600">ID</th>
            </tr>
          </thead>
          <tbody>
            {importPreview.slice(0, 15).map((p, i) => {
              const n = importType === 'entity' ? p.name : ((p.firstName || '') + ' ' + (p.lastName || '')).trim();
              const id = importType === 'entity' ? p.ein : (p.ssn || p.itin);
              return (
                <tr key={i} className="border-b border-slate-100">
                  <td className="px-3 py-2 text-slate-400">{i + 1}</td>
                  <td className="px-3 py-2 font-medium">{n}</td>
                  <td className="px-3 py-2 font-mono text-xs">{id}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
        {importPreview.length > 15 && <div className="bg-slate-50 px-3 py-2 text-xs text-slate-500 text-center">... and {importPreview.length - 15} more</div>}
      </div>
    </div>
  );
}

// UTILITY
function StatCard({ icon, label, value, color }) {
  const c = { red: 'bg-red-50 text-red-700', navy: 'bg-blue-50 text-blue-900', slate: 'bg-slate-100 text-slate-700', amber: 'bg-amber-50 text-amber-700' };
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-sm">
      <div className={'w-9 h-9 rounded-lg flex items-center justify-center mb-2 ' + c[color]}>{icon}</div>
      <div className="text-2xl font-bold">{value}</div>
      <div className="text-xs text-slate-500">{label}</div>
    </div>
  );
}

function Section({ title, icon, children }) {
  const [exp, setExp] = useState(true);
  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm">
      <button onClick={() => setExp(!exp)} className="w-full flex items-center justify-between px-5 py-3.5 border-b border-slate-100 hover:bg-slate-50/50 rounded-t-xl">
        <h3 className="font-bold text-sm text-slate-800 flex items-center gap-2">{icon}{title}</h3>
        {exp ? <ChevronDown className="w-4 h-4 text-slate-400" /> : <ChevronRight className="w-4 h-4 text-slate-400" />}
      </button>
      {exp && <div className="p-5">{children}</div>}
    </div>
  );
}

function Grid({ children }) {
  return <div className="grid grid-cols-1 md:grid-cols-2 gap-3">{children}</div>;
}

function Field({ label, value, editing, onChange, options, type, mono, full, multiline, placeholder }) {
  const wc = full ? 'md:col-span-2' : '';
  const mc = mono ? 'font-mono' : '';
  if (!editing) {
    return (
      <div className={wc}>
        <div className="text-xs font-bold text-slate-500 uppercase tracking-wide mb-1">{label}</div>
        <div className={'text-sm text-slate-900 min-h-[20px] ' + mc + (multiline ? ' whitespace-pre-wrap' : '')}>
          {value || <span className="text-slate-300 font-normal">—</span>}
        </div>
      </div>
    );
  }
  return (
    <div className={wc}>
      <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">{label}</label>
      {options ? (
        <select value={value || ''} onChange={e => onChange(e.target.value)} className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 bg-white focus:outline-none focus:ring-2 focus:ring-red-600">
          <option value="">— Select —</option>
          {options.map(o => <option key={o} value={o}>{o}</option>)}
        </select>
      ) : multiline ? (
        <textarea value={value || ''} onChange={e => onChange(e.target.value)} rows={2} placeholder={placeholder} className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-600" />
      ) : (
        <input type={type || 'text'} value={value || ''} onChange={e => onChange(e.target.value)} placeholder={placeholder} className={'w-full text-sm border border-slate-300 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-red-600 ' + mc} />
      )}
    </div>
  );
}

function ReadField({ label, value }) {
  return (
    <div>
      <div className="text-xs font-bold text-slate-500 uppercase tracking-wide mb-1">{label}</div>
      <div className="text-sm text-slate-900 font-medium">{value || <span className="text-slate-300 font-normal">—</span>}</div>
    </div>
  );
}

function StatusBadge({ status }) {
  const c = { 'Active': 'bg-blue-100 text-blue-900', 'Inactive': 'bg-slate-100 text-slate-600', 'Prospective': 'bg-amber-100 text-amber-700', 'Onboarding': 'bg-red-100 text-red-700', 'Terminated': 'bg-rose-100 text-rose-700' };
  return <span className={'inline-block px-2 py-0.5 text-xs font-bold rounded ' + (c[status] || 'bg-slate-100 text-slate-600')}>{status || 'Unknown'}</span>;
}

function PaymentBadge({ status }) {
  const c = { 'Pending': 'bg-amber-100 text-amber-700', 'Invoiced': 'bg-red-100 text-red-700', 'Partially Paid': 'bg-blue-100 text-blue-900', 'Paid': 'bg-blue-100 text-blue-900', 'Overdue': 'bg-rose-100 text-rose-700', 'Written Off': 'bg-slate-200 text-slate-600' };
  return <span className={'inline-block px-2 py-0.5 text-xs font-bold rounded ' + (c[status] || 'bg-slate-100 text-slate-600')}>{status || '—'}</span>;
}

function FilingStatusBadge({ status }) {
  const c = { 'Not Started': 'bg-slate-100 text-slate-600', 'In Progress': 'bg-amber-100 text-amber-700', 'Documents Pending': 'bg-amber-100 text-amber-700', 'Ready for Review': 'bg-red-100 text-red-700', 'Filed': 'bg-blue-100 text-blue-900', 'Filed - Payment Pending': 'bg-blue-100 text-blue-900', 'Pending Payment': 'bg-amber-100 text-amber-700', 'Paid': 'bg-blue-100 text-blue-900', 'Completed': 'bg-blue-100 text-blue-900', 'Extended': 'bg-orange-100 text-orange-700', 'On Hold': 'bg-slate-200 text-slate-600', 'N/A': 'bg-slate-100 text-slate-400' };
  return <span className={'inline-block px-1.5 py-0.5 text-xs font-bold rounded ' + (c[status] || 'bg-slate-100 text-slate-600')}>{status || '—'}</span>;
}

function KV({ label, value, mono }) {
  return <div className="text-xs"><span className="text-slate-500">{label}:</span> <span className={mono ? 'font-mono' : 'font-medium'}>{value}</span></div>;
}

function BankAccountsSection({ current, editing, draft, addBank, updateBank, removeBank }) {
  const items = (editing ? draft : current).bankAccounts || [];
  return (
    <Section title="Bank & Login Credentials" icon={<Key className="w-4 h-4 text-blue-900" />}>
      <div className="bg-amber-50 border border-amber-200 rounded-lg p-2.5 mb-3 text-xs text-amber-800 flex items-start gap-2">
        <AlertCircle className="w-3.5 h-3.5 mt-0.5 flex-shrink-0" />
        <div><strong>Sensitive data:</strong> Account numbers and passwords are stored in browser storage. Use additional care for high-value accounts.</div>
      </div>
      {editing && <button onClick={addBank} className="mb-3 px-3 py-1.5 text-xs rounded-lg border border-red-200 text-red-700 hover:bg-red-50 flex items-center gap-1"><Plus className="w-3 h-3" /> Add Bank / Card</button>}
      {items.length === 0 ? (
        <div className="text-sm text-slate-400 italic">No accounts added.</div>
      ) : (
        <div className="space-y-3">{items.map((b, i) => <BankCard key={b.id || i} bank={b} idx={i} editing={editing} updateBank={updateBank} removeBank={removeBank} />)}</div>
      )}
    </Section>
  );
}

function BankCard({ bank, idx, editing, updateBank, removeBank }) {
  const [showPw, setShowPw] = useState(false);
  if (editing) return (
    <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/50">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
        <Field label="Bank / Card Name" value={bank.bankName} editing={true} onChange={v => updateBank(idx, 'bankName', v)} placeholder="Chase, Amex Platinum, etc." />
        <Field label="Account Type" value={bank.accountType} editing={true} onChange={v => updateBank(idx, 'accountType', v)} options={['Bank Account', 'Credit Card', 'Investment', 'Loan', 'Other']} />
        <Field label="Account Holder" value={bank.accountHolder} editing={true} onChange={v => updateBank(idx, 'accountHolder', v)} />
        <Field label="Account / Card #" value={bank.accountNumber} editing={true} onChange={v => updateBank(idx, 'accountNumber', v)} mono />
        <Field label="Routing #" value={bank.routingNumber} editing={true} onChange={v => updateBank(idx, 'routingNumber', v)} mono />
        <Field label="Login URL" value={bank.loginUrl} editing={true} onChange={v => updateBank(idx, 'loginUrl', v)} placeholder="https://..." />
        <Field label="Login Username" value={bank.loginUsername} editing={true} onChange={v => updateBank(idx, 'loginUsername', v)} mono />
        <div>
          <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">Login Password</label>
          <div className="relative">
            <input type={showPw ? 'text' : 'password'} value={bank.loginPassword || ''} onChange={e => updateBank(idx, 'loginPassword', e.target.value)} className="w-full text-sm border border-slate-300 rounded-lg px-3 py-2 pr-10 font-mono focus:outline-none focus:ring-2 focus:ring-red-600" />
            <button onClick={() => setShowPw(!showPw)} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400">
              {showPw ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
            </button>
          </div>
        </div>
        <div className="md:col-span-2"><Field label="Notes" value={bank.notes} editing={true} onChange={v => updateBank(idx, 'notes', v)} multiline /></div>
      </div>
      <button onClick={() => removeBank(idx)} className="text-xs text-rose-600 hover:underline mt-2">Remove</button>
    </div>
  );
  return (
    <div className="border-l-4 border-blue-900 border-t border-r border-b border-slate-200 rounded-lg p-3 bg-blue-50/30">
      <div className="font-bold text-sm text-slate-900">{bank.bankName || '(no name)'}</div>
      <div className="text-xs text-blue-900 font-bold uppercase tracking-wide mb-2">{bank.accountType || 'Account'}</div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-x-3 gap-y-1">
        {bank.accountHolder && <KV label="Holder" value={bank.accountHolder} />}
        {bank.accountNumber && <KV label="Account #" value={bank.accountNumber} mono />}
        {bank.routingNumber && <KV label="Routing #" value={bank.routingNumber} mono />}
        {bank.loginUrl && <KV label="URL" value={bank.loginUrl} />}
        {bank.loginUsername && <KV label="Username" value={bank.loginUsername} mono />}
        {bank.loginPassword && (
          <div className="text-xs flex items-center gap-1">
            <span className="text-slate-500">Password:</span>
            <span className="font-mono">{showPw ? bank.loginPassword : '••••••••'}</span>
            <button onClick={() => setShowPw(!showPw)} className="text-slate-400 hover:text-slate-600">{showPw ? <EyeOff className="w-3 h-3" /> : <Eye className="w-3 h-3" />}</button>
          </div>
        )}
      </div>
      {bank.notes && <div className="mt-2 text-xs text-slate-600 italic">{bank.notes}</div>}
    </div>
  );
}

function EmployeesSection({ current, editing, draft, addEmployee, updateEmployee, removeEmployee, attachIdCard }) {
  const items = (((editing ? draft : current).payroll) || {}).employeeList || [];
  return (
    <Section title="Employee Details" icon={<Users className="w-4 h-4 text-blue-900" />}>
      <div className="bg-amber-50 border border-amber-200 rounded-lg p-2.5 mb-3 text-xs text-amber-800 flex items-start gap-2">
        <AlertCircle className="w-3.5 h-3.5 mt-0.5 flex-shrink-0" />
        <div><strong>Sensitive PII:</strong> SSNs and ID documents are stored in browser storage. Handle with care.</div>
      </div>
      {editing && <button onClick={addEmployee} className="mb-3 px-3 py-1.5 text-xs rounded-lg border border-red-200 text-red-700 hover:bg-red-50 flex items-center gap-1"><UserPlus className="w-3 h-3" /> Add Employee</button>}
      {items.length === 0 ? (
        <div className="text-sm text-slate-400 italic">No employees added.</div>
      ) : (
        <div className="space-y-3">{items.map((e, i) => <EmployeeCard key={e.id || i} emp={e} idx={i} editing={editing} updateEmployee={updateEmployee} removeEmployee={removeEmployee} attachIdCard={attachIdCard} />)}</div>
      )}
    </Section>
  );
}

function EmployeeCard({ emp, idx, editing, updateEmployee, removeEmployee, attachIdCard }) {
  const handleFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    if (file.size > 5 * 1024 * 1024) { alert('File too large (max 5MB)'); return; }
    const reader = new FileReader();
    reader.onload = () => attachIdCard(idx, file.name, reader.result);
    reader.readAsDataURL(file);
    e.target.value = '';
  };
  const isImg = emp.idCardData && emp.idCardData.indexOf('data:image/') === 0;
  const status = emp.status || 'Active';
  if (editing) return (
    <div className="border border-slate-200 rounded-lg p-3 bg-slate-50/50">
      <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
        <Field label="Full Name" value={emp.fullName} editing={true} onChange={v => updateEmployee(idx, 'fullName', v)} />
        <Field label="Employment Status" value={status} editing={true} onChange={v => updateEmployee(idx, 'status', v)} options={['Active', 'Inactive', 'Terminated', 'On Leave']} />
        <Field label="Date of Birth" value={emp.dob} editing={true} onChange={v => updateEmployee(idx, 'dob', v)} type="date" />
        <Field label="SSN" value={emp.ssn} editing={true} onChange={v => updateEmployee(idx, 'ssn', v)} mono />
        <Field label="Position" value={emp.position} editing={true} onChange={v => updateEmployee(idx, 'position', v)} />
        <Field label="Hire Date" value={emp.hireDate} editing={true} onChange={v => updateEmployee(idx, 'hireDate', v)} type="date" />
        {(status === 'Inactive' || status === 'Terminated') && <Field label="Termination / Inactive Date" value={emp.termDate} editing={true} onChange={v => updateEmployee(idx, 'termDate', v)} type="date" />}
        <Field label="Email" value={emp.email} editing={true} onChange={v => updateEmployee(idx, 'email', v)} />
        <Field label="Phone" value={emp.phone} editing={true} onChange={v => updateEmployee(idx, 'phone', v)} />
      </div>
      <div className="mt-3 pt-3 border-t border-slate-200">
        <div className="text-xs font-bold text-blue-900 uppercase tracking-wide mb-2 flex items-center gap-1.5"><Key className="w-3.5 h-3.5" /> Direct Deposit / Bank Details</div>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
          <Field label="Bank Name" value={emp.bankName} editing={true} onChange={v => updateEmployee(idx, 'bankName', v)} placeholder="Chase, BoA, etc." />
          <Field label="Account Type" value={emp.bankAccountType || 'Checking'} editing={true} onChange={v => updateEmployee(idx, 'bankAccountType', v)} options={['Checking', 'Savings']} />
          <Field label="Routing #" value={emp.bankRouting} editing={true} onChange={v => updateEmployee(idx, 'bankRouting', v)} mono />
          <Field label="Account #" value={emp.bankAccount} editing={true} onChange={v => updateEmployee(idx, 'bankAccount', v)} mono />
        </div>
      </div>
      <div className="mt-3 pt-3 border-t border-slate-200 grid grid-cols-1 md:grid-cols-2 gap-2">
        <div>
          <label className="block text-xs font-bold text-slate-700 uppercase tracking-wide mb-1.5">ID Card (Image or PDF)</label>
          <input type="file" accept="image/*,.pdf" onChange={handleFile} className="block text-sm text-slate-700 file:mr-2 file:py-1 file:px-2 file:rounded file:border-0 file:text-xs file:font-medium file:bg-red-700 file:text-white" />
          {emp.idCardName && <div className="text-xs text-slate-500 mt-1">Attached: {emp.idCardName}</div>}
        </div>
        <div className="md:col-span-2"><Field label="Notes" value={emp.notes} editing={true} onChange={v => updateEmployee(idx, 'notes', v)} multiline /></div>
      </div>
      <button onClick={() => removeEmployee(idx)} className="text-xs text-rose-600 hover:underline mt-2">Remove employee</button>
    </div>
  );
  const sc = { 'Active': 'bg-emerald-100 text-emerald-700', 'Inactive': 'bg-slate-200 text-slate-600', 'Terminated': 'bg-rose-100 text-rose-700', 'On Leave': 'bg-amber-100 text-amber-700' };
  return (
    <div className={'border border-slate-200 rounded-lg p-3 ' + (status === 'Active' ? 'bg-slate-50/30' : 'bg-slate-100/60')}>
      <div className="flex items-start gap-3">
        <div className="w-10 h-10 rounded-full bg-gradient-to-br from-red-700 to-blue-900 text-white flex items-center justify-center font-bold flex-shrink-0">{(emp.fullName || '?').charAt(0).toUpperCase()}</div>
        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2 flex-wrap">
            <div className="font-bold text-sm text-slate-900">{emp.fullName || '(no name)'}</div>
            <span className={'px-2 py-0.5 text-xs font-bold rounded ' + (sc[status] || sc['Inactive'])}>{status}</span>
          </div>
          {emp.position && <div className="text-xs text-slate-500">{emp.position}</div>}
          <div className="mt-1 grid grid-cols-1 md:grid-cols-2 gap-x-3 gap-y-0.5">
            {emp.dob && <KV label="DOB" value={emp.dob} />}
            {emp.ssn && <KV label="SSN" value={emp.ssn} mono />}
            {emp.hireDate && <KV label="Hired" value={emp.hireDate} />}
            {emp.termDate && (status === 'Inactive' || status === 'Terminated') && <KV label="Ended" value={emp.termDate} />}
            {emp.email && <KV label="Email" value={emp.email} />}
            {emp.phone && <KV label="Phone" value={emp.phone} />}
          </div>
          {(emp.bankName || emp.bankAccount || emp.bankRouting) && (
            <div className="mt-2 rounded-md border border-blue-100 bg-blue-50/40 p-2">
              <div className="text-xs font-bold text-blue-900 uppercase tracking-wide mb-1">Direct Deposit</div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-x-3 gap-y-0.5">
                {emp.bankName && <KV label="Bank" value={emp.bankName + (emp.bankAccountType ? ' (' + emp.bankAccountType + ')' : '')} />}
                {emp.bankRouting && <KV label="Routing #" value={emp.bankRouting} mono />}
                {emp.bankAccount && <KV label="Account #" value={emp.bankAccount} mono />}
              </div>
            </div>
          )}
          {emp.idCardData && (
            <div className="mt-2">
              <a href={emp.idCardData} download={emp.idCardName || 'id-card'} className="text-xs text-red-700 hover:underline flex items-center gap-1">
                <FileText className="w-3 h-3" /> {emp.idCardName || 'ID Card'}
              </a>
              {isImg && <img src={emp.idCardData} alt="ID" className="mt-1 max-w-[200px] max-h-[120px] rounded border border-slate-200" />}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}