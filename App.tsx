/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Search, Plus, Trash2, RefreshCw, Download, FileText, 
  ArrowRight, ArrowLeft, ShieldCheck, Globe, Settings, 
  AlertCircle, CheckCircle2, Clock, XCircle, ExternalLink, 
  RotateCcw, Upload, Menu, Lock, Phone, X, LayoutDashboard, 
  Activity, Database, Share2, Link2, AlertTriangle, Send, 
  Users, UserPlus, FileSpreadsheet, FileX, Archive, Briefcase, 
  Tag, BarChart3, ChevronDown, FolderOpen, HardDrive, Hash, 
  CalendarDays, Pencil, Check, CheckSquare, ArrowLeftRight, 
  SlidersHorizontal, ArrowDownUp, Info
} from 'lucide-react';

// ═══════════════════════════════════════════════════════════════
// Utility functions
// ═══════════════════════════════════════════════════════════════
function hexToRgba(hex: string, alpha: number) {
  let c = (hex || '#10b981').substring(1);
  if (c.length === 3) c = c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
  const r = parseInt(c.substring(0, 2), 16) || 0;
  const g = parseInt(c.substring(2, 4), 16) || 0;
  const b = parseInt(c.substring(4, 6), 16) || 0;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function darkenColor(hex: string, percent: number) {
  let c = (hex || '#10b981').substring(1);
  if (c.length === 3) c = c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
  let r = parseInt(c.substring(0, 2), 16) || 0;
  let g = parseInt(c.substring(2, 4), 16) || 0;
  let b = parseInt(c.substring(4, 6), 16) || 0;
  r = Math.max(0, Math.min(255, Math.round(r * (1 - percent))));
  g = Math.max(0, Math.min(255, Math.round(g * (1 - percent))));
  b = Math.max(0, Math.min(255, Math.round(b * (1 - percent))));
  return `#${r.toString(16).padStart(2,'0')}${g.toString(16).padStart(2,'0')}${b.toString(16).padStart(2,'0')}`;
}

// ═══════════════════════════════════════════════════════════════
// Types
// ═══════════════════════════════════════════════════════════════
interface VisaEntry {
  id: string;
  passportNumber: string;
  applicationNumber: string;
  applicantName?: string;
  nationality: string;
  arrivalPoint: string;
  visaType: string;
  applicantData?: string;
  status: 'Idle' | 'Checking' | 'Found' | 'Error' | 'Retrying';
  statusText?: string;
  visaNumber?: string;
  lastUpdate?: string;
  phoneNumber?: string;
  profession?: string;
  updatedAtUnix?: number;
  dataUpdatedAtUnix?: number;
  applicationDate?: string;
  embassy?: string;
  entriesCount?: string;
  requesterName?: string;
  nameEnglish?: string;
  passportType?: string;
  passportExpiry?: string;
  birthDate?: string;
  birthPlace?: string;
  currentNationality?: string;
  gender?: string;
  purpose?: string;
  documentNumber?: string;
  customId?: string;
  officeId?: string;
  txStatus?: string;
  txStatusType?: 'مستمرة' | 'موقفة' | 'بدون';
  txStatusDate?: string;
  txStatusNote?: string;
  archived?: boolean;
  archivedAt?: number;
  printImageUrl?: string;
  visaImageUrl?: string;
  healthCertStatus?: string;
}

interface Office { id: string; name: string; phone: string; }
interface TxStatusOption { id: string; label: string; type: 'مستمرة'|'موقفة'|'بدون'; hasStatsCard?: boolean; statsIsolate?: boolean; }

const DEFAULT_SITE_IDENTITY = {
  officeNameAr: 'نظام إنجاز التأشيرات',
  officeNameEn: 'Injaz Visa System',
  officeDescriptionAr: 'نظام متابعة الجوازات والتأشيرات',
  officeDescriptionEn: 'Passport and Visa Tracking System',
  heroBadgeAr: 'نظام إدارة التأشيرات',
  heroBadgeEn: 'VISA MANAGEMENT SYSTEM',
  heroTitleAr: 'استعلم عن حالة معاملتك',
  heroTitleEn: 'Track Your Application',
  searchInquiryAr: 'استعلام عن حالة المعاملة',
  searchInquiryEn: 'Check Status',
  searchDescAr: 'أدخل رقم الجواز ورقم الهاتف المسجل',
  searchDescEn: 'Enter passport and registered phone',
  supportPhone: '771234567',
  supportEmail: 'support@injaz.com',
  hqAddressAr: 'اليمن',
  hqAddressEn: 'Yemen',
  themePreset: 'emerald',
  primaryColor: '#10b981',
  secondaryColor: '#0ea5e9',
  searchEnabled: true,
  searchDisabledMessageAr: 'الاستعلام متوقف مؤقتاً للصيانة',
  searchDisabledMessageEn: 'Search temporarily disabled for maintenance',
  logo: ''
};

const NO_PRINT_INDICATORS = [
  'حدث خطأ يرجى التأكد من البيانات المدخلة',
  'يرجى التأكد من البيانات المدخلة',
  'ليس لها برنت', 'لم يصدر لها برنت', 'لا يوجد برنت',
  'بدون برنت', 'بلا برنت'
];

function entryHasPrint(entry: VisaEntry): boolean {
  if (!entry) return false;
  const appNo = (entry.applicationNumber || '').trim();
  if (appNo === '---' || appNo === '' || appNo === 'N/A') return false;
  const appData = (entry.applicantData || '').trim();
  const statText = (entry.statusText || '').trim();
  const containsNoPrint = NO_PRINT_INDICATORS.some(ind => appData.includes(ind) || statText.includes(ind));
  if (containsNoPrint) return false;
  return appNo !== '---' && appNo !== '';
}

function entryHasVisaAndImage(entry: VisaEntry): boolean {
  if (!entry) return false;
  const vNum = (entry.visaNumber || '').trim();
  const hasVisa = vNum !== '' && vNum !== '---' && vNum !== 'N/A';
  const hasImg = !!(entry.visaImageUrl);
  return hasVisa && hasImg;
}

// ═══════════════════════════════════════════════════════════════
// Inline Sub-Components
// ═══════════════════════════════════════════════════════════════

function AdminLoginPage({ lang, adminUsernameInput, setAdminUsernameInput, adminPinInput, setAdminPinInput, adminPinError, handleLoginSubmit, siteIdentity }: any) {
  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-900 via-slate-800 to-slate-900 p-4" dir="rtl">
      <div className="w-full max-w-md">
        <div className="bg-white rounded-3xl shadow-2xl p-8 border border-slate-100">
          <div className="flex flex-col items-center mb-8">
            {siteIdentity?.logo ? (
              <div className="w-20 h-20 rounded-2xl overflow-hidden border border-slate-200 bg-white flex items-center justify-center mb-4 shadow-lg">
                <img src={siteIdentity.logo} alt="Logo" className="w-full h-full object-contain" />
              </div>
            ) : (
              <div className="w-20 h-20 rounded-2xl bg-gradient-to-tr from-emerald-600 to-teal-500 flex items-center justify-center mb-4 shadow-lg">
                <ShieldCheck size={36} className="text-white" />
              </div>
            )}
            <h1 className="text-2xl font-black text-slate-900 text-center">لوحة التحكم الإدارية</h1>
            <p className="text-sm text-slate-400 mt-2 text-center font-medium">سجّل الدخول للوصول إلى النظام</p>
          </div>
          {adminPinError && (
            <div className="mb-5 p-3.5 rounded-2xl bg-red-50 border border-red-200 text-red-700 text-sm font-semibold flex items-center gap-2">
              <AlertCircle size={16} className="shrink-0" />
              <span>{adminPinError}</span>
            </div>
          )}
          <div className="space-y-4">
            <div>
              <label className="text-xs font-black text-slate-500 block mb-2 uppercase tracking-wider">اسم المستخدم</label>
              <input type="text" value={adminUsernameInput} onChange={(e) => setAdminUsernameInput(e.target.value)}
                className="w-full px-4 py-3.5 rounded-2xl border border-slate-200 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 outline-none transition-all font-semibold text-slate-900"
                placeholder="admin" />
            </div>
            <div>
              <label className="text-xs font-black text-slate-500 block mb-2 uppercase tracking-wider">كلمة المرور</label>
              <input type="password" value={adminPinInput} onChange={(e) => setAdminPinInput(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleLoginSubmit()}
                className="w-full px-4 py-3.5 rounded-2xl border border-slate-200 focus:border-emerald-500 focus:ring-4 focus:ring-emerald-500/10 outline-none transition-all font-semibold text-slate-900"
                placeholder="••••••" />
            </div>
            <button onClick={handleLoginSubmit}
              className="w-full py-3.5 rounded-2xl bg-gradient-to-r from-emerald-600 to-teal-500 hover:from-emerald-700 hover:to-teal-600 text-white font-black transition-all shadow-lg flex items-center justify-center gap-2 cursor-pointer">
              <span>دخول النظام</span><ArrowRight size={16} />
            </button>
          </div>
          <p className="text-center text-[11px] text-slate-400 mt-6 font-medium">الافتراضي: admin / 1234</p>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// Main App Component
// ═══════════════════════════════════════════════════════════════
export default function App() {
  const [entries, setEntries] = useState<VisaEntry[]>(() => {
    try { const s = localStorage.getItem('injaz_visa_entries'); return s ? JSON.parse(s) : []; } catch { return []; }
  });
  const [siteIdentity, setSiteIdentity] = useState<any>(() => {
    try { const s = localStorage.getItem('injaz_visa_site_identity'); return s ? { ...DEFAULT_SITE_IDENTITY, ...JSON.parse(s) } : DEFAULT_SITE_IDENTITY; } catch { return DEFAULT_SITE_IDENTITY; }
  });
  const [offices, setOffices] = useState<Office[]>([]);
  const [txStatusOptions, setTxStatusOptions] = useState<TxStatusOption[]>(() => {
    try { const s = localStorage.getItem('txStatusOptions'); return s ? JSON.parse(s) : [
      { id: 'ts1', label: 'مستمرة', type: 'مستمرة' },
      { id: 'ts2', label: 'موقفة', type: 'موقفة' },
      { id: 'ts3', label: 'مرتجع تم الترحيل من جديد', type: 'مستمرة' },
      { id: 'ts4', label: 'جاهزة للاستلام', type: 'مستمرة' },
    ]; } catch { return []; }
  });
  const [lang, setLang] = useState<'ar'|'en'>('ar');
  const [viewMode, setViewMode] = useState<'public'|'admin'|'login'>('public');
  const [activeTab, setActiveTab] = useState<'dashboard'|'all'|'customers'|'files'|'identity'|'security'|'excel_sync'|'auto_check'>('all');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [adminUsernameInput, setAdminUsernameInput] = useState('');
  const [adminPinInput, setAdminPinInput] = useState('');
  const [adminPinError, setAdminPinError] = useState('');
  const [currentUser, setCurrentUser] = useState<any>(() => {
    try { const s = localStorage.getItem('injaz_visa_current_user'); return s ? JSON.parse(s) : null; } catch { return null; }
  });

  // UI state
  const [filterQuery, setFilterQuery] = useState('');
  const [expandedCardId, setExpandedCardId] = useState<string|null>(null);
  const [editingCardId, setEditingCardId] = useState<string|null>(null);
  const [alertModal, setAlertModal] = useState<{msg: string; type: 'success'|'error'|'warning'|'info'}|null>(null);
  const [confirmModal, setConfirmModal] = useState<{msg: string; onConfirm: () => void}|null>(null);
  const [isAddOpen, setIsAddOpen] = useState(false);
  const [isBulkOpen, setIsBulkOpen] = useState(false);
  const [bulkText, setBulkText] = useState('');
  const [parsing, setParsing] = useState(false);
  const [checkingId, setCheckingId] = useState<string|null>(null);
  const [addForm, setAddForm] = useState({ name:'', passport:'', appNo:'', phone:'', profession:'' });

  // Public search state
  const [publicPassport, setPublicPassport] = useState('');
  const [publicPhone, setPublicPhone] = useState('');
  const [publicSearchResult, setPublicSearchResult] = useState<VisaEntry|null>(null);
  const [publicHasSearched, setPublicHasSearched] = useState(false);
  const [isPublicSearching, setIsPublicSearching] = useState(false);

  const showAlert = (msg: string, type: 'success'|'error'|'warning'|'info' = 'info') => setAlertModal({ msg, type });
  const showConfirm = (msg: string, onConfirm: () => void) => setConfirmModal({ msg, onConfirm });

  // Load data from server on mount
  useEffect(() => {
    fetch('/api/entries').then(r => r.json()).then((data: any) => {
      if (Array.isArray(data) && data.length > 0) setEntries(data);
    }).catch(() => {});
    fetch('/api/offices').then(r => r.json()).then((d: any) => { if (Array.isArray(d)) setOffices(d); }).catch(() => {});
    fetch('/api/tx-status-options').then(r => r.json()).then((d: any) => {
      if (Array.isArray(d) && d.length > 0) setTxStatusOptions(d);
    }).catch(() => {});
    fetch('/api/site-identity').then(r => r.json()).then((d: any) => {
      if (d && typeof d === 'object') setSiteIdentity({ ...DEFAULT_SITE_IDENTITY, ...d });
    }).catch(() => {});
  }, []);

  // Persist entries locally
  useEffect(() => {
    try { localStorage.setItem('injaz_visa_entries', JSON.stringify(entries)); } catch {}
  }, [entries]);

  // Sync entries to server
  const syncEntries = (next: VisaEntry[]) => {
    setEntries(next);
    fetch('/api/entries', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ entries: next })
    }).catch(() => {});
  };

  const handleLoginSubmit = async () => {
    try {
      const res = await fetch('/api/login', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: adminUsernameInput, password: adminPinInput })
      });
      if (res.ok) {
        const data = await res.json();
        setCurrentUser(data.user);
        try { localStorage.setItem('injaz_visa_current_user', JSON.stringify(data.user)); } catch {}
        setViewMode('admin');
        setAdminUsernameInput(''); setAdminPinInput(''); setAdminPinError('');
      } else {
        const err = await res.json();
        setAdminPinError(err.error || => 'بيانات الدخول غير ص setTimeoutحيحة');
      }
    } catch ((re: any) { setAdminPinError('فشل الاتصال,  بالخادم'); }
  };

  const handlePublicSearch = async () => {
    if (!publicPassport.trim() || !publicPhone.trim()) return;
    setIsPublicSearching(true);
    setPublicHasSearched(false);
    await new Promise(r600));
    const cleanPhone = publicPhone.trim().replace(/[\s\-\(\)\+]/g, '');
    const match = entries.find(e => {
      const dbPass = e.passportNumber.trim().toUpperCase();
      const dbPhone = (e.phoneNumber || '').trim().replace(/[\s\-\(\)\+]/g, '');
      const passMatch = dbPass === publicPassport.trim().toUpperCase();
      const phoneMatch = dbPhone.includes(cleanPhone) || cleanPhone.includes(dbPhone);
      return passMatch && phoneMatch;
    });
    setIsPublicSearching(false);
    setPublicSearchResult(match || null);
    setPublicHasSearched(true);
  };

  const saveEntry = () => {
    if (!addForm.passport) { showAlert('رقم الجواز إجباري', 'warning'); return; }
    const newEntry: VisaEntry = {
      id: Math.random().toString(36).substr(2, 9),
      applicantName: addForm.name,
      passportNumber: addForm.passport.trim().toUpperCase(),
      applicationNumber: addForm.appNo.trim() || '---',
      nationality: 'اليمن', arrivalPoint: 'عدن', visaType: 'عمل',
      phoneNumber: addForm.phone.trim(),
      profession: addForm.profession.trim(),
      status: 'Idle', statusText: 'قيد الانتظار (Idle)', visaNumber: '---',
      lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true }),
      updatedAtUnix: Date.now()
    };
    syncEntries([newEntry, ...entries]);
    setIsAddOpen(false);
    setAddForm({ name:'', passport:'', appNo:'', phone:'', profession:'' });
  };

  const removeEntry = (id: string) => {
    showConfirm('هل أنت متأكد من حذف هذا العميل نهائياً؟', () => {
      syncEntries(entries.filter(e => e.id !== id));
    });
  };

  const updateEntryField = (id: string, field: keyof VisaEntry, value: any) => {
    const next = entries.map(e => e.id === id ? { ...e, [field]: value } : e);
    syncEntries(next);
  };

  const handleBulkParse = async () => {
    if (!bulkText.trim()) return;
    setParsing(true);
    await new Promise(r => setTimeout(r, 600));
    const passportRegex = /\b[A-Za-z]{1,2}[0-9]{6,12}\b|\b[0-9]{7,12}\b/gi;
    const appNoRegex = /\b[0-9]{8,15}\b/g;
    const nameRegex = /([أ-ي]{2,}\s?){2,4}|([A-Z-a-z]{2,}\s?){2,4}/g;
    const passports = (bulkText.match(passportRegex) || []).map(p => p.toUpperCase());
    const appNos = bulkText.match(appNoRegex) || [];
    const names = bulkText.match(nameRegex) || [];
    const count = Math.max(passports.length, appNos.length);
    const newEntries: VisaEntry[] = [];
    for (let i = 0; i < count; i++) {
      newEntries.push({
        id: Math.random().toString(36).substr(2, 9),
        passportNumber: passports[i] || '',
        applicationNumber: appNos[i] || '---',
        applicantName: names[i] || '',
        nationality: 'اليمن', arrivalPoint: 'عدن', visaType: 'عمل',
        status: 'Idle', statusText: 'قيد الانتظار (Idle)', visaNumber: '---',
        lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true }),
        updatedAtUnix: Date.now()
      });
    }
    syncEntries([...newEntries, ...entries]);
    setIsBulkOpen(false);
    setBulkText('');
    setParsing(false);
  };

  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target?.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
        if (data.length === 0) return;
        const firstRowStr = (data[0] || []).map(c => String(c||'').toLowerCase());
        const hasHeader = firstRowStr.some(c => c.includes('passport') || c.includes('جواز') || c.includes('اسم') || c.includes('name'));
        const startIdx = hasHeader ? 1 : 0;
        const newEntries: VisaEntry[] = data.slice(startIdx).flatMap((row: any[]) => {
          const pass = String(row[1] || row[0] || '').trim().toUpperCase();
          if (!pass) return [];
          return [{
            id: Math.random().toString(36).substr(2, 9),
            passportNumber: pass,
            applicationNumber: String(row[2] || '---').trim(),
            applicantName: String(row[0] || '').trim(),
            nationality: 'اليمن', arrivalPoint: 'عدن', visaType: 'عمل',
            phoneNumber: String(row[4] || '').trim(),
            profession: String(row[3] || '').trim(),
            status: 'Idle', statusText: 'قيد الانتظار (Idle)', visaNumber: '---',
            lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true }),
            updatedAtUnix: Date.now()
          }];
        });
        syncEntries([...newEntries, ...entries]);
        showAlert(`تم استيراد ${newEntries.length} عميل بنجاح`, 'success');
      } catch { showAlert('فشل قراءة الملف', 'error'); }
    };
    reader.readAsBinaryString(file);
  };

  const checkVisa = async (entry: VisaEntry) => {
    setCheckingId(entry.id);
    try {
      const initRes = await fetch('/api/mofa/init');
      const initData = await initRes.json();
      if (!initData.sessionId) throw new Error('فشل بدء الجلسة');
      if (!initData.autoSolvedCaptcha) throw new Error('تعذّر حل الكابتشا');
      const submitRes = await fetch('/api/mofa/submit', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sessionId: initData.sessionId,
          passportNumber: entry.passportNumber,
          nationality: initData.nationalities?.[0]?.value || 'YEM',
          visaType: initData.visaTypes?.[0]?.value || '1',
          embassy: initData.embassies?.[0]?.value || '302',
          captcha: initData.autoSolvedCaptcha,
          applicantName: entry.applicantName || ''
        })
      });
      const data = await submitRes.json();
      if (data.success && data.data) {
        const d = data.data;
        const next = entries.map(e => e.id === entry.id ? {
          ...e, status: 'Found' as const,
          statusText: d.statusText || 'تم الاستعلام',
          applicantName: d.applicantName || e.applicantName,
          applicationNumber: d.applicationNumber || e.applicationNumber,
          visaNumber: d.visaNumber || e.visaNumber,
          profession: d.profession || e.profession,
          printImageUrl: d.printPdfUrl || e.printImageUrl,
          lastUpdate: new Date().toLocaleTimeString('ar-YE', { hour: '2-digit', minute: '2-digit' }),
          updatedAtUnix: Date.now(),
          dataUpdatedAtUnix: Date.now()
        } : e);
        syncEntries(next);
        showAlert('تم الاستعلام بنجاح', 'success');
      } else {
        showAlert(data.error || 'فشل الاستعلام', 'warning');
      }
    } catch (e: any) {
      showAlert(e.message || 'فشل الاتصال', 'error');
    } finally { setCheckingId(null); }
  };

  // ─── Compute stats
  const stats = useMemo(() => ({
    total: entries.filter(e => !e.archived).length,
    issued: entries.filter(e => !e.archived && entryHasPrint(e) && (e.statusText?.includes('اصدار') || e.statusText?.includes('Issued'))).length,
    pending: entries.filter(e => !e.archived && (e.status === 'Idle' || e.status === 'Checking')).length,
    archived: entries.filter(e => e.archived).length,
  }), [entries]);

  const filteredByTab = useMemo(() => {
    let base = entries.filter(e => !e.archived);
    if (filterQuery) {
      const q = filterQuery.toLowerCase();
      base = base.filter(e =>
        (e.passportNumber || '').toLowerCase().includes(q) ||
        (e.applicationNumber || '').toLowerCase().includes(q) ||
        (e.applicantName || '').toLowerCase().includes(q) ||
        (e.phoneNumber || '').includes(q)
      );
    }
    return base;
  }, [entries, filterQuery]);

  // ═══════════════════════════════════════════════════════════
  // RENDER: Public Portal
  // ═══════════════════════════════════════════════════════════
  if (viewMode === 'login') {
    return (
      <>
        <AdminLoginPage
          lang={lang} adminUsernameInput={adminUsernameInput} setAdminUsernameInput={setAdminUsernameInput}
          adminPinInput={adminPinInput} setAdminPinInput={setAdminPinInput}
          adminPinError={adminPinError} handleLoginSubmit={handleLoginSubmit}
          siteIdentity={siteIdentity}
        />
        <AlertModal alertModal={alertModal} setAlertModal={setAlertModal} lang={lang} />
        <ConfirmModal confirmModal={confirmModal} setConfirmModal={setConfirmModal} lang={lang} />
      </>
    );
  }

  if (viewMode === 'public') {
    return (
      <div className="min-h-screen bg-[#F8F9FA] text-slate-900" dir="rtl">
        {/* Public Navbar */}
        <nav className="h-20 bg-white/80 backdrop-blur-md border-b border-slate-200 px-6 md:px-12 flex items-center justify-between sticky top-0 z-50 shadow-sm">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 rounded-2xl overflow-hidden border border-slate-200 bg-white flex items-center justify-center shadow-md">
              {siteIdentity.logo ? <img src={siteIdentity.logo} alt="Logo" className="w-full h-full object-contain" /> : <ShieldCheck className="w-7 h-7 text-emerald-600" />}
            </div>
            <div>
              <h1 className="text-base md:text-xl font-black text-slate-900 tracking-tight">{siteIdentity.officeNameAr}</h1>
              <p className="text-[10px] uppercase font-extrabold text-emerald-600 mt-1 tracking-wider">{siteIdentity.heroBadgeAr}</p>
            </div>
          </div>
          <button onClick={() => setViewMode('login')} className="text-xs font-bold bg-slate-900 text-white px-5 py-2.5 rounded-2xl hover:bg-slate-800 transition-all cursor-pointer">
            <Lock className="inline w-3 h-3 ml-1" /> دخول الإدارة
          </button>
        </nav>

        {/* Hero */}
        <section className="py-16 px-6 text-center">
          <h2 className="text-2xl md:text-4xl font-black text-slate-900 mb-4">{siteIdentity.heroTitleAr}</h2>
          <p className="text-slate-600 text-sm md:text-base max-w-2xl mx-auto">{siteIdentity.officeDescriptionAr}</p>
        </section>

        {/* Search Card */}
        <main className="max-w-4xl mx-auto px-4 pb-20">
          {siteIdentity.searchEnabled !== false ? (
            <div className="bg-white p-8 md:p-10 rounded-3xl shadow-xl border border-slate-100 -mt-8 relative z-10">
              <div className="text-center mb-6">
                <h3 className="text-xl font-bold text-slate-900">{siteIdentity.searchInquiryAr}</h3>
                <p className="text-xs text-slate-500 mt-2">{siteIdentity.searchDescAr}</p>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase mb-2 block">رقم الجواز</label>
                  <input type="text" value={publicPassport} onChange={e => setPublicPassport(e.target.value.toUpperCase())}
                    className="w-full border border-slate-200 rounded-2xl py-4 px-5 text-sm font-bold font-mono focus:outline-none focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 transition-all"
                    placeholder="A1234567" />
                </div>
                <div>
                  <label className="text-xs font-bold text-slate-500 uppercase mb-2 block">رقم الهاتف</label>
                  <input type="text" value={publicPhone} onChange={e => setPublicPhone(e.target.value)}
                    className="w-full border border-slate-200 rounded-2xl py-4 px-5 text-sm font-bold focus:outline-none focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 transition-all"
                    placeholder="7XXXXXXXX" />
                </div>
              </div>
              <button onClick={handlePublicSearch} disabled={isPublicSearching || !publicPassport.trim() || !publicPhone.trim()}
                className="w-full mt-6 bg-gradient-to-r from-emerald-600 to-teal-500 text-white font-black py-4 rounded-2xl shadow-lg hover:shadow-xl transition-all disabled:opacity-50 flex items-center justify-center gap-3 cursor-pointer">
                {isPublicSearching ? <><RefreshCw className="w-4 h-4 animate-spin" /> جاري البحث...</> : <><Search className="w-4 h-4" /> بحث</>}
              </button>
            </div>
          ) : (
            <div className="bg-amber-50 border-2 border-amber-200 p-8 rounded-3xl text-center -mt-8">
              <AlertTriangle className="w-12 h-12 text-amber-600 mx-auto mb-4" />
              <p className="text-slate-700 font-bold">{siteIdentity.searchDisabledMessageAr}</p>
            </div>
          )}

          {/* Results */}
          {publicHasSearched && (
            <div className="mt-8">
              {publicSearchResult ? (
                <div className="bg-white rounded-3xl shadow-2xl overflow-hidden border border-slate-200">
                  <div className="bg-gradient-to-r from-emerald-600 to-teal-500 text-white px-8 py-6">
                    <span className="px-3 py-1 bg-white/20 text-white text-[10px] font-black uppercase rounded-md">تم العثور على المعاملة</span>
                    <h4 className="text-2xl font-black mt-3">تفاصيل المعاملة</h4>
                  </div>
                  <div className="p-8 space-y-6">
                    <div className="bg-emerald-50 p-6 rounded-2xl border border-emerald-100">
                      <span className="text-xs font-black text-slate-500 uppercase">حالة المعاملة</span>
                      <p className="text-lg font-black text-emerald-700 mt-2">{publicSearchResult.statusText || 'قيد الإجراء'}</p>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <Field label="اسم العميل" value={publicSearchResult.applicantName || '---'} />
                      <Field label="رقم الجواز" value={publicSearchResult.passportNumber} mono />
                      <Field label="رقم الطلب" value={publicSearchResult.applicationNumber || '---'} mono />
                      <Field label="رقم التأشيرة" value={publicSearchResult.visaNumber || '---'} mono />
                      <Field label="نوع التأشيرة" value={publicSearchResult.visaType} />
                      <Field label="جهة القدوم" value={publicSearchResult.arrivalPoint} />
                      <Field label="المهنة" value={publicSearchResult.profession || '---'} />
                      <Field label="آخر تحديث" value={publicSearchResult.lastUpdate || '---'} />
                    </div>
                    {publicSearchResult.applicantData && (
                      <div className="bg-slate-50 p-5 rounded-2xl border border-slate-100">
                        <p className="text-xs font-black text-slate-500 uppercase mb-2">ملاحظات إدارية</p>
                        <p className="text-sm text-slate-700 leading-relaxed">{publicSearchResult.applicantData}</p>
                      </div>
                    )}
                  </div>
                </div>
              ) : (
                <div className="bg-white rounded-3xl shadow-xl p-10 text-center border border-red-100">
                  <AlertCircle className="w-16 h-16 text-red-500 mx-auto mb-4" />
                  <h4 className="text-xl font-bold mb-2">لم يتم العثور على المعاملة</h4>
                  <p className="text-sm text-slate-500">تأكد من صحة رقم الجواز ورقم الهاتف</p>
                  <button onClick={() => { setPublicPassport(''); setPublicPhone(''); setPublicHasSearched(false); }}
                    className="mt-6 bg-slate-100 hover:bg-slate-200 px-6 py-3 rounded-xl font-bold text-sm cursor-pointer">
                    محاولة جديدة
                  </button>
                </div>
              )}
            </div>
          )}
        </main>

        <AlertModal alertModal={alertModal} setAlertModal={setAlertModal} lang={lang} />
        <ConfirmModal confirmModal={confirmModal} setConfirmModal={setConfirmModal} lang={lang} />
      </div>
    );
  }

  // ═══════════════════════════════════════════════════════════
  // RENDER: Admin Panel
  // ═══════════════════════════════════════════════════════════
  return (
    <div className="min-h-screen bg-[#F8F9FA] flex" dir="rtl">
      {/* Mobile Sidebar Overlay */}
      {isSidebarOpen && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] lg:hidden" onClick={() => setIsSidebarOpen(false)} />
      )}

      {/* Sidebar */}
      <aside className={`w-72 bg-white border-l border-slate-200 flex flex-col py-6 fixed lg:static inset-y-0 right-0 z-[70] transition-transform ${isSidebarOpen ? 'translate-x-0' : 'translate-x-full lg:translate-x-0'}`}>
        <div className="px-6 mb-6 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl overflow-hidden border border-slate-200 bg-white flex items-center justify-center shadow-sm">
              {siteIdentity.logo ? <img src={siteIdentity.logo} alt="Logo" className="w-full h-full object-contain" /> : <ShieldCheck className="w-5 h-5 text-emerald-600" />}
            </div>
            <div>
              <div className="font-extrabold text-slate-900 text-xs">{siteIdentity.officeNameAr}</div>
              <div className="text-[8px] uppercase font-bold text-slate-400 mt-0.5 tracking-wider">لوحة الإدارة</div>
            </div>
          </div>
          <button className="lg:hidden text-slate-400 p-1" onClick={() => setIsSidebarOpen(false)}><X size={20} /></button>
        </div>

        <nav className="flex flex-col gap-1 px-3 flex-1">
          {[
            { id: 'dashboard', label: 'لوحة التحكم', icon: <LayoutDashboard size={18} /> },
            { id: 'all', label: 'المعاملات', icon: <Globe size={18} /> },
            { id: 'customers', label: 'إدارة العملاء', icon: <Users size={18} /> },
            { id: 'files', label: 'ملفات العملاء', icon: <FolderOpen size={18} /> },
            { id: 'identity', label: 'هوية الموقع', icon: <Settings size={18} /> },
            { id: 'security', label: 'الأمان والمستخدمين', icon: <Lock size={18} /> },
            { id: 'excel_sync', label: 'مزامنة البيانات', icon: <Database size={18} /> },
            { id: 'auto_check', label: 'التشييك التلقائي', icon: <Clock size={18} /> },
          ].map((item: any) => (
            <button key={item.id} onClick={() => { setActiveTab(item.id); setIsSidebarOpen(false); }}
              className={`flex items-center gap-3 px-3 py-3 rounded-xl text-sm font-bold transition-all ${activeTab === item.id ? 'bg-slate-900 text-white shadow-lg' : 'text-slate-500 hover:bg-slate-50'}`}>
              {item.icon}
              {item.label}
            </button>
          ))}
        </nav>

        <div className="px-4 mt-auto">
          <button onClick={() => { setViewMode('public'); setCurrentUser(null); localStorage.removeItem('injaz_visa_current_user'); }}
            className="w-full bg-slate-950 hover:bg-slate-800 text-white py-3 rounded-2xl text-xs font-black flex items-center justify-center gap-2 cursor-pointer">
            <ArrowRight size={14} /> العودة للبوابة
          </button>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto p-4 md:p-8 space-y-6">
        <header className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <button onClick={() => setIsSidebarOpen(true)} className="lg:hidden w-10 h-10 flex items-center justify-center border border-slate-200 rounded-xl bg-white shadow-sm">
              <Menu size={20} />
            </button>
            <div className="text-right">
              <h2 className="text-lg md:text-xl font-black tracking-tight flex items-center gap-3">
                {activeTab === 'dashboard' ? 'لوحة التحكم' : activeTab === 'all' ? 'المعاملات' : activeTab === 'customers' ? 'إدارة العملاء' : activeTab === 'files' ? 'ملفات العملاء' : activeTab === 'identity' ? 'هوية الموقع' : activeTab === 'security' ? 'الأمان والمستخدمين' : activeTab === 'excel_sync' ? 'مزامنة البيانات' : 'التشييك التلقائي'}
                <span className="bg-slate-200 text-slate-600 text-xs px-2.5 py-0.5 rounded-full">{entries.length}</span>
              </h2>
              <p className="text-sm font-medium text-slate-400 mt-1">{currentUser?.name || 'مدير النظام'}</p>
            </div>
          </div>
        </header>

        {/* ─── Dashboard Tab ─── */}
        {activeTab === 'dashboard' && (
          <div className="space-y-6">
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <StatCard label="إجمالي المعاملات" value={stats.total} color="bg-slate-900" icon={<Globe size={24} />} />
              <StatCard label="المؤشرة والجاهزة" value={stats.issued} color="bg-emerald-600" icon={<CheckCircle2 size={24} />} />
              <StatCard label="قيد الانتظار" value={stats.pending} color="bg-amber-500" icon={<Clock size={24} />} />
              <StatCard label="الأرشيف" value={stats.archived} color="bg-violet-600" icon={<Archive size={24} />} />
            </div>
            <div className="bg-white p-6 rounded-3xl border border-slate-200 shadow-sm">
              <h3 className="font-bold text-slate-800 mb-4">آخر 10 معاملات</h3>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-right text-slate-500 border-b">
                      <th className="p-3">الاسم</th>
                      <th className="p-3">الجواز</th>
                      <th className="p-3">الطلب</th>
                      <th className="p-3">الحالة</th>
                    </tr>
                  </thead>
                  <tbody>
                    {entries.filter(e => !e.archived).slice(0, 10).map(e => (
                      <tr key={e.id} className="border-b hover:bg-slate-50">
                        <td className="p-3 font-medium">{e.applicantName || '—'}</td>
                        <td className="p-3 font-mono text-xs">{e.passportNumber}</td>
                        <td className="p-3 font-mono text-xs">{e.applicationNumber}</td>
                        <td className="p-3"><span className="px-2 py-1 rounded-full text-xs bg-slate-100">{e.statusText || '—'}</span></td>
                      </tr>
                    ))}
                    {entries.length === 0 && <tr><td colSpan={4} className="text-center text-slate-400 p-8">لا توجد بيانات</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ─── Entries Tab ─── */}
        {activeTab === 'all' && (
          <div className="space-y-5">
            {/* Toolbar */}
            <div className="bg-white p-4 rounded-3xl border border-slate-200 shadow-sm flex flex-wrap items-center gap-3">
              <div className="relative flex-1 min-w-[240px]">
                <Search className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                <input type="text" value={filterQuery} onChange={e => setFilterQuery(e.target.value)}
                  className="w-full bg-slate-50 border border-slate-200 rounded-2xl py-3 pr-10 pl-4 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20"
                  placeholder="ابحث عن جواز، رقم طلب، أو اسم..." />
              </div>
              <button onClick={() => setIsAddOpen(true)} className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-sm py-3 px-5 rounded-2xl flex items-center gap-2 cursor-pointer shadow-sm">
                <Plus size={16} /> إضافة
              </button>
              <button onClick={() => setIsBulkOpen(true)} className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-sm py-3 px-5 rounded-2xl flex items-center gap-2 cursor-pointer">
                <RefreshCw size={16} /> استيراد ذكي
              </button>
              <label className="bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-sm py-3 px-5 rounded-2xl flex items-center gap-2 cursor-pointer">
                <Upload size={16} /> Excel
                <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleImportExcel} />
              </label>
            </div>

            {/* Cards Grid */}
            {filteredByTab.length === 0 ? (
              <div className="bg-white rounded-3xl p-20 text-center border border-slate-200">
                <Search size={48} className="text-slate-300 mx-auto mb-4" />
                <p className="text-slate-400 font-bold">لا توجد نتائج</p>
              </div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
                {filteredByTab.map((entry, idx) => (
                  <EntryCard
                    key={entry.id}
                    entry={entry}
                    idx={idx}
                    expanded={expandedCardId === entry.id}
                    editing={editingCardId === entry.id}
                    onToggleExpand={() => setExpandedCardId(expandedCardId === entry.id ? null : entry.id)}
                    onToggleEdit={() => setEditingCardId(editingCardId === entry.id ? null : entry.id)}
                    onCheck={() => checkVisa(entry)}
                    checking={checkingId === entry.id}
                    onDelete={() => removeEntry(entry.id)}
                    onUpdateField={(f, v) => updateEntryField(entry.id, f, v)}
                    lang={lang}
                  />
                ))}
              </div>
            )}
          </div>
        )}

        {/* ─── Other Tabs (placeholder) ─── */}
        {activeTab !== 'dashboard' && activeTab !== 'all' && (
          <div className="bg-white p-12 rounded-3xl border border-slate-200 shadow-sm text-center">
            <Info size={48} className="text-slate-300 mx-auto mb-4" />
            <h3 className="text-lg font-black text-slate-700 mb-2">قيد التطوير</h3>
            <p className="text-sm text-slate-400 max-w-md mx-auto">
              هذا التبويب سيُفعّل في التحديث القادم. حالياً يمكنك استخدام تبويب "المعاملات" للوصول إلى جميع العملاء.
            </p>
          </div>
        )}
      </main>

      {/* ══════════════ Modals ══════════════ */}
      {isAddOpen && (
        <Modal onClose={() => setIsAddOpen(false)} title="إضافة عميل جديد">
          <div className="space-y-4">
            <Input label="الاسم الكامل" value={addForm.name} onChange={v => setAddForm({...addForm, name: v})} />
            <Input label="رقم الجواز *" value={addForm.passport} onChange={v => setAddForm({...addForm, passport: v.toUpperCase()})} mono />
            <Input label="رقم الطلب" value={addForm.appNo} onChange={v => setAddForm({...addForm, appNo: v})} mono />
            <Input label="رقم الهاتف" value={addForm.phone} onChange={v => setAddForm({...addForm, phone: v})} mono />
            <Input label="المهنة" value={addForm.profession} onChange={v => setAddForm({...addForm, profession: v})} />
            <div className="flex gap-3 justify-end pt-4">
              <button onClick={() => setIsAddOpen(false)} className="px-6 py-3 rounded-2xl bg-slate-100 font-bold text-sm cursor-pointer">إلغاء</button>
              <button onClick={saveEntry} className="px-8 py-3 rounded-2xl bg-emerald-600 text-white font-bold text-sm cursor-pointer">حفظ</button>
            </div>
          </div>
        </Modal>
      )}

      {isBulkOpen && (
        <Modal onClose={() => setIsBulkOpen(false)} title="الاستيراد الذكي">
          <p className="text-sm text-slate-500 mb-4">الصق نصاً يحتوي على بيانات (أرقام جوازات، أرقام طلبات، أسماء)</p>
          <textarea value={bulkText} onChange={e => setBulkText(e.target.value)} rows={10}
            className="w-full border border-slate-200 rounded-2xl p-4 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 font-mono"
            placeholder="الصق النص هنا..." />
          <div className="flex gap-3 justify-end pt-4">
            <button onClick={() => setIsBulkOpen(false)} className="px-6 py-3 rounded-2xl bg-slate-100 font-bold text-sm cursor-pointer">إلغاء</button>
            <button onClick={handleBulkParse} disabled={parsing} className="px-8 py-3 rounded-2xl bg-emerald-600 text-white font-bold text-sm cursor-pointer disabled:opacity-50">
              {parsing ? 'جاري التحليل...' : 'بدء التحليل'}
            </button>
          </div>
        </Modal>
      )}

      <AlertModal alertModal={alertModal} setAlertModal={setAlertModal} lang={lang} />
      <ConfirmModal confirmModal={confirmModal} setConfirmModal={setConfirmModal} lang={lang} />
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
// Helper Components
// ═══════════════════════════════════════════════════════════════
function Field({ label, value, mono }: { label: string; value: string; mono?: boolean }) {
  return (
    <div>
      <p className="text-xs font-black text-slate-400 uppercase tracking-wider mb-1">{label}</p>
      <p className={`text-sm font-bold text-slate-900 ${mono ? 'font-mono' : ''}`}>{value}</p>
    </div>
  );
}

function StatCard({ label, value, color, icon }: any) {
  return (
    <div className={`${color} text-white p-5 rounded-3xl shadow-lg relative overflow-hidden`}>
      <div className="flex justify-between items-start">
        <div>
          <p className="text-xs font-bold opacity-80">{label}</p>
          <p className="text-3xl font-black mt-1">{value}</p>
        </div>
        <div className="opacity-60">{icon}</div>
      </div>
    </div>
  );
}

function EntryCard({ entry, idx, expanded, editing, onToggleExpand, onToggleEdit, onCheck, checking, onDelete, onUpdateField, lang }: any) {
  const statusColor = entry.status === 'Found' ? 'bg-emerald-500' : entry.status === 'Error' ? 'bg-red-500' : entry.status === 'Checking' ? 'bg-amber-500' : 'bg-slate-400';
  const statusLabel = entry.status === 'Found' ? 'مكتملة' : entry.status === 'Error' ? 'خطأ' : entry.status === 'Checking' ? 'جاري...' : 'انتظار';
  return (
    <div className={`bg-white rounded-3xl border shadow-sm overflow-hidden transition-all ${expanded ? 'col-span-1 sm:col-span-2 xl:col-span-3 border-slate-300 shadow-md' : 'border-slate-200 hover:border-slate-300 hover:shadow-md'}`}>
      <div className="p-5 cursor-pointer" onClick={onToggleExpand}>
        <div className="flex justify-between items-start gap-3 mb-3">
          <div className="min-w-0 flex-1 text-right">
            <p className="text-sm font-black text-slate-900 truncate">{entry.applicantName || '—'}</p>
            <p className="text-[10px] font-mono font-bold text-slate-400 mt-0.5">{entry.passportNumber || '—'}</p>
          </div>
          <div className="flex flex-col items-end gap-1 shrink-0">
            <span className="inline-flex items-center gap-1.5 text-[10px] font-black px-2.5 py-1 rounded-full border border-slate-200 bg-slate-50">
              <span className={`w-1.5 h-1.5 rounded-full ${statusColor} ${entry.status === 'Checking' ? 'animate-pulse' : ''}`} />
              {statusLabel}
            </span>
            <span className="text-[9px] font-black text-slate-300 font-mono">#{idx + 1}</span>
          </div>
        </div>
        <div className="flex items-center justify-between pt-3 border-t border-slate-100">
          <div className="flex items-center gap-1" onClick={e => e.stopPropagation()}>
            <button onClick={onCheck} disabled={checking} className="p-1.5 rounded-xl hover:bg-emerald-50 hover:text-emerald-600 transition-all text-slate-400 cursor-pointer" title="تشييك">
              <RefreshCw size={13} className={checking ? 'animate-spin' : ''} />
            </button>
            <button onClick={onToggleEdit} className={`p-1.5 rounded-xl transition-all cursor-pointer ${editing ? 'bg-blue-50 text-blue-500' : 'hover:bg-blue-50 hover:text-blue-500 text-slate-400'}`} title="تعديل">
              <Pencil size={13} />
            </button>
            <button onClick={onDelete} className="p-1.5 rounded-xl hover:bg-red-50 hover:text-red-500 transition-all text-slate-400 cursor-pointer" title="حذف">
              <Trash2 size={13} />
            </button>
          </div>
          <span className="text-[9px] font-bold text-slate-300 font-mono">{entry.lastUpdate}</span>
        </div>
      </div>

      {expanded && (
        <div className="border-t border-slate-100 p-5 space-y-3">
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="text-[9px] font-black text-slate-400 block mb-1">رقم الجواز</label>
              <input className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-mono font-bold text-center"
                value={entry.passportNumber} disabled={!editing} onChange={e => onUpdateField('passportNumber', e.target.value.toUpperCase())} />
            </div>
            <div>
              <label className="text-[9px] font-black text-slate-400 block mb-1">رقم الطلب</label>
              <input className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-mono font-bold text-center"
                value={entry.applicationNumber} disabled={!editing} onChange={e => onUpdateField('applicationNumber', e.target.value)} />
            </div>
            <div>
              <label className="text-[9px] font-black text-slate-400 block mb-1">رقم التأشيرة</label>
              <input className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-mono font-bold text-center"
                value={entry.visaNumber || '---'} disabled={!editing} onChange={e => onUpdateField('visaNumber', e.target.value)} />
            </div>
            <div>
              <label className="text-[9px] font-black text-slate-400 block mb-1">رقم الهاتف</label>
              <input className="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs font-mono font-bold text-center"
                value={entry.phoneNumber || ''} disabled={!editing} onChange={e => onUpdateField('phoneNumber', e.target.value)} />
            </div>
          </div>
          <div className="bg-emerald-50 p-3 rounded-xl border border-emerald-100">
            <p className="text-[9px] font-black text-emerald-700 uppercase">حالة المعاملة</p>
            <p className="text-xs font-bold text-emerald-800 mt-1">{entry.statusText || '—'}</p>
          </div>
          {entry.printImageUrl && (
            <a href={entry.printImageUrl.startsWith('tg:') ? `/api/tg-download?fileId=${entry.printImageUrl.split(':')[1]}&filename=print.pdf` : entry.printImageUrl}
              target="_blank" rel="noreferrer" className="flex items-center justify-center gap-2 w-full bg-blue-50 hover:bg-blue-100 text-blue-700 py-3 rounded-xl text-xs font-black transition-all">
              <FileText size={14} /> فتح ملف البرنت PDF
            </a>
          )}
        </div>
      )}
    </div>
  );
}

function Modal({ children, onClose, title }: any) {
  return (
    <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/50 backdrop-blur-sm p-4" onClick={onClose}>
      <div className="bg-white rounded-3xl shadow-2xl w-full max-w-lg max-h-[90vh] overflow-y-auto" onClick={e => e.stopPropagation()}>
        <div className="p-6 border-b border-slate-100 flex items-center justify-between sticky top-0 bg-white">
          <h3 className="text-lg font-black text-slate-900">{title}</h3>
          <button onClick={onClose} className="p-1.5 rounded-xl hover:bg-slate-100 cursor-pointer"><X size={20} /></button>
        </div>
        <div className="p-6">{children}</div>
      </div>
    </div>
  );
}

function Input({ label, value, onChange, mono }: any) {
  return (
    <div>
      <label className="text-xs font-black text-slate-500 block mb-2 uppercase tracking-wider">{label}</label>
      <input type="text" value={value} onChange={e => onChange(e.target.value)}
        className={`w-full border border-slate-200 rounded-2xl px-4 py-3 text-sm focus:outline-none focus:ring-4 focus:ring-emerald-500/10 focus:border-emerald-500 transition-all font-semibold text-slate-900 ${mono ? 'font-mono' : ''}`} />
    </div>
  );
}

function AlertModal({ alertModal, setAlertModal, lang }: any) {
  if (!alertModal) return null;
  const icons: any = {
    success: <CheckCircle2 size={24} className="text-emerald-500" />,
    error: <XCircle size={24} className="text-red-500" />,
    warning: <AlertCircle size={24} className="text-amber-500" />,
    info: <Info size={24} className="text-blue-500" />,
  };
  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/50 backdrop-blur-sm p-4" onClick={() => setAlertModal(null)}>
      <div className="bg-white rounded-3xl shadow-2xl p-6 max-w-md w-full border border-slate-100" onClick={e => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-5">
          <div className="shrink-0 mt-0.5">{icons[alertModal.type] || icons.info}</div>
          <p className="text-slate-800 font-semibold leading-relaxed flex-1 whitespace-pre-wrap">{alertModal.msg}</p>
        </div>
        <button onClick={() => setAlertModal(null)} className="w-full py-3 rounded-2xl bg-slate-900 hover:bg-slate-800 text-white font-black transition-all cursor-pointer">
          حسناً
        </button>
      </div>
    </div>
  );
}

function ConfirmModal({ confirmModal, setConfirmModal, lang }: any) {
  if (!confirmModal) return null;
  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center bg-black/50 backdrop-blur-sm p-4">
      <div className="bg-white rounded-3xl shadow-2xl p-6 max-w-md w-full border border-slate-100">
        <div className="flex items-start gap-3 mb-5">
          <div className="shrink-0 mt-0.5"><AlertCircle size={24} className="text-amber-500" /></div>
          <p className="text-slate-800 font-semibold leading-relaxed flex-1 whitespace-pre-wrap">{confirmModal.msg}</p>
        </div>
        <div className="flex gap-2.5">
          <button onClick={() => setConfirmModal(null)} className="flex-1 py-3 rounded-2xl bg-slate-100 hover:bg-slate-200 text-slate-700 font-black transition-all cursor-pointer">
            إلغاء
          </button>
          <button onClick={() => { try { confirmModal.onConfirm?.(); } catch {} setConfirmModal(null); }}
            className="flex-1 py-3 rounded-2xl bg-red-600 hover:bg-red-700 text-white font-black transition-all cursor-pointer">
            تأكيد
          </button>
        </div>
      </div>
    </div>
  );
}
