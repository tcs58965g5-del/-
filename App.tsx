/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Search, 
  Plus, 
  Trash2, 
  RefreshCw, 
  Download, 
  FileText, 
  ArrowRight,
  ArrowLeft,
  ShieldCheck,
  Globe,
  Settings,
  AlertCircle,
  CheckCircle2,
  Clock,
  XCircle,
  ExternalLink,
  RotateCcw,
  Upload,
  Menu,
  Lock,
  Phone,
  X,
  LayoutDashboard,
  Activity,
  Database,
  Share2,
  Link2,
  AlertTriangle,
  Send,
  Users,
  UserPlus,
  FileSpreadsheet,
  FileX,
  Archive,
  Briefcase,
  Tag,
  BarChart3,
  ChevronDown,
  FolderOpen,
  HardDrive,
  Hash,
  CalendarDays,
  Pencil,
  Check,
  CheckSquare,
  ArrowLeftRight,
  SlidersHorizontal,
  ArrowDownUp
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import AppModals from './components/AppModals';
const DashboardTab = React.lazy(() => import('./components/tabs/DashboardTab'));
const IdentityTab = React.lazy(() => import('./components/tabs/IdentityTab'));
const SecurityTab = React.lazy(() => import('./components/tabs/SecurityTab'));
const ExcelSyncTab = React.lazy(() => import('./components/tabs/ExcelSyncTab'));
const FilesTab = React.lazy(() => import('./components/tabs/FilesTab'));
const AutoCheckTab = React.lazy(() => import('./components/tabs/AutoCheckTab'));
import AdminLoginPage from './components/AdminLoginPage';

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
  captchaCode?: string;
  solvedCaptcha?: string;
  captchaSvg?: string;
  checkLogs?: string[];
  // MOFA extended fields
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
  printImageBase64?: string;
  printImageUrl?: string;
  visaImageBase64?: string;
  visaImageUrl?: string;
  healthCertStatus?: string;
}

interface Office {
  id: string;
  name: string;
  phone: string;
}

function hexToRgba(hex: string, alpha: number) {
  let c = (hex || '#10b981').substring(1);
  if (c.length === 3) {
    c = c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
  }
  const r = parseInt(c.substring(0, 2), 16) || 0;
  const g = parseInt(c.substring(2, 4), 16) || 0;
  const b = parseInt(c.substring(4, 6), 16) || 0;
  return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function adjustBrightness(hex: string, percent: number) {
  let c = (hex || '#10b981').substring(1);
  if (c.length === 3) {
    c = c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
  }
  let r = parseInt(c.substring(0, 2), 16) || 0;
  let g = parseInt(c.substring(2, 4), 16) || 0;
  let b = parseInt(c.substring(4, 6), 16) || 0;

  r = Math.min(255, Math.max(0, Math.round(r + (255 - r) * percent)));
  g = Math.min(255, Math.max(0, Math.round(g + (255 - g) * percent)));
  b = Math.min(255, Math.max(0, Math.round(b + (255 - b) * percent)));

  const rHex = r.toString(16).padStart(2, '0');
  const gHex = g.toString(16).padStart(2, '0');
  const bHex = b.toString(16).padStart(2, '0');

  return `#${rHex}${gHex}${bHex}`;
}

function darkenColor(hex: string, percent: number) {
  let c = (hex || '#10b981').substring(1);
  if (c.length === 3) {
    c = c[0] + c[0] + c[1] + c[1] + c[2] + c[2];
  }
  let r = parseInt(c.substring(0, 2), 16) || 0;
  let g = parseInt(c.substring(2, 4), 16) || 0;
  let b = parseInt(c.substring(4, 6), 16) || 0;

  r = Math.max(0, Math.min(255, Math.round(r * (1 - percent))));
  g = Math.max(0, Math.min(255, Math.round(g * (1 - percent))));
  b = Math.max(0, Math.min(255, Math.round(b * (1 - percent))));

  const rHex = r.toString(16).padStart(2, '0');
  const gHex = g.toString(16).padStart(2, '0');
  const bHex = b.toString(16).padStart(2, '0');

  return `#${rHex}${gHex}${bHex}`;
}

const NO_PRINT_INDICATORS = [
  'حدث خطأ يرجى التأكد من البيانات المدخلة والمحاولة مرة أخرى',
  'يرجى التأكد من البيانات المدخلة',
  'ليس لها برنت',
  'لم يصدر لها برنت',
  'لا يوجد برنت',
  'بدون برنت',
  'بلا برنت'
];

export function entryHasPrint(entry: VisaEntry): boolean {
  if (!entry) return false;

  const appNo = (entry.applicationNumber || '').trim();
  // If applicationNumber is literally '---', empty, or N/A, then it doesn't have a print
  if (appNo === '---' || appNo === '' || appNo === 'N/A') {
    return false;
  }

  const appData = (entry.applicantData || '').trim();
  const statText = (entry.statusText || '').trim();

  // If statusText or applicantData contains any of the "no print" indicators (such as the mofa error message)
  const containsNoPrintText = NO_PRINT_INDICATORS.some(indicator => 
    appData.includes(indicator) || statText.includes(indicator)
  );

  if (containsNoPrintText) {
    return false;
  }

  // Otherwise, if they have some applicantData or applicationNumber, default to true or depending on requirements.
  const hasAppNo = appNo !== '---' && appNo !== '';
  const hasAppData = appData !== '' && appData !== '---';

  return hasAppNo || hasAppData;
}

export function entryHasVisaAndImage(entry: VisaEntry): boolean {
  if (!entry) return false;
  const visaNum = (entry.visaNumber || '').trim();
  const hasVisa = visaNum !== '' && visaNum !== '---' && visaNum !== 'N/A';
  const hasImage = !!(entry.visaImageUrl || entry.visaImageBase64);
  return hasVisa && hasImage;
}

const DEFAULT_SITE_IDENTITY = {
  officeNameAr: 'مكتب إنجاز للإرساليات والخدمات الإلكترونية',
  officeNameEn: 'Injaz Visa & E-Services Office',
  officeDescriptionAr: 'البوابة الرسمية المعتمدة لتتبع واستعلام معاملات التأشيرات المعتمدة وتحديثها اللحظي من وزارة الخارجية وطباعتها بأعلى موثوقية وسرعة',
  officeDescriptionEn: 'Approved Gateway to track visa status and check live updates from the Ministry of Foreign Affairs (MOFA)',
  heroBadgeAr: 'وكيل معتمد ومسجل رسمي لدى السفارات وموفا',
  heroBadgeEn: 'OFFICIAL REGISTERED TRAVEL AGENT',
  heroTitleAr: 'استعلم عن حالة فيزتك بلحظة',
  heroTitleEn: 'Track Your Visa Status Instantly',
  searchInquiryAr: 'استعلام عن حالة المعاملة',
  searchInquiryEn: 'Check Status',
  searchDescAr: 'أدخل معلومات الجواز المدخلة بالسجلات ومطابقتها برقم هاتفك المسجل',
  searchDescEn: 'Enter passport credentials and registered mobile to parse status',
  supportPhone: '771234567',
  supportEmail: 'support@injaz-visa.com',
  hqAddressAr: 'المركز الرئيسي: عدن، المنصورة / صنعاء، شارع الجزائر',
  hqAddressEn: 'Headquarters: Aden, Al-Mansourah / Sanaa, Algeria St',
  themePreset: 'emerald', // 'emerald' | 'blue' | 'amber' | 'violet' | 'rose' | 'custom'
  primaryColor: '#10b981',
  secondaryColor: '#0ea5e9',
  searchEnabled: true,
  searchDisabledMessageAr: 'عذراً، تم إيقاف نظام الاستعلام الذاتي مؤقتاً بالصفحة الرئيسية من قبل إدارة النظام لتحديث البيانات وإجراء أعمال الصيانة الدورية المعتادة.',
  searchDisabledMessageEn: 'The homepage public search portal is temporarily deactivated by the administration for data synchronization and scheduled database maintenance.',
  logo: ''
};

export default function App() {
  // Persistence state loading
  const [entries, setEntries] = useState<VisaEntry[]>(() => {
    const saved = localStorage.getItem('injaz_visa_entries');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          // Filter out existing mock data to cleanly wipe out experimental customers
          return parsed.filter((e: any) => !e.id?.toString().startsWith('mock-'));
        }
      } catch (e) {
        console.error('Failed to parse saved entries', e);
      }
    }
    return [];
  });

  // Website Identity State — initialised from localStorage (fast), then synced from server
  const [siteIdentity, setSiteIdentity] = useState(() => {
    const saved = localStorage.getItem('injaz_visa_site_identity');
    if (saved) {
      try {
        return { ...DEFAULT_SITE_IDENTITY, ...JSON.parse(saved) };
      } catch (e) {
        console.error('Failed to parse site identity', e);
      }
    }
    return DEFAULT_SITE_IDENTITY;
  });

  // Tracks whether the user has modified identity since page load — prevents a
  // slow server response from overwriting in-progress edits.
  const identityDirtyRef = useRef(false);
  const setSiteIdentityUser = React.useCallback((updater: any) => {
    identityDirtyRef.current = true;
    setSiteIdentity(updater);
  }, []);

  // Load authoritative identity from server once on mount
  useEffect(() => {
    fetch('/api/site-identity')
      .then(r => r.json())
      .then((data) => {
        if (!identityDirtyRef.current && data && typeof data === 'object') {
          const merged = { ...DEFAULT_SITE_IDENTITY, ...data };
          setSiteIdentity(merged);
          try { localStorage.setItem('injaz_visa_site_identity', JSON.stringify(merged)); } catch {}
        }
      })
      .catch(() => {});
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Load auto-check config from server on mount
  useEffect(() => {
    fetch('/api/autocheck-config')
      .then(r => r.json())
      .then((data) => {
        if (data && typeof data === 'object') {
          if (data.enabled !== undefined) setAutoCheckEnabled(data.enabled);
          if (data.time) setAutoCheckTime(data.time);
          if (data.interval !== undefined) setAutoCheckInterval(data.interval);
          if (data.lastRun !== undefined) setAutoCheckLastRun(data.lastRun);
          if (Array.isArray(data.log)) setAutoCheckLog(data.log);
        }
      })
      .catch(() => {});
  }, []);

  // Load Google Sheets config from server on mount
  useEffect(() => {
    fetch('/api/google-sheets-config')
      .then(r => r.json())
      .then((data) => {
        if (data && typeof data === 'object') {
          if (data.spreadsheetId !== undefined) setGoogleSpreadsheetId(data.spreadsheetId);
          if (data.apiKey !== undefined) setGoogleApiKey(data.apiKey);
          if (data.sheetName !== undefined) setGoogleSheetName(data.sheetName);
          if (data.lastSyncTime) setLastSyncTime(data.lastSyncTime);
          if (Array.isArray(data.syncLogs)) setSyncLogs(data.syncLogs);
          if (Array.isArray(data.recentUpdates)) {
            const now = Date.now();
            setRecentUpdates(data.recentUpdates.filter((item: any) => {
              const itemTime = new Date(item.copiedAt).getTime();
              return now - itemTime < 24 * 60 * 60 * 1000;
            }));
          }
          if (Array.isArray(data.readyVisas)) setReadyVisas(data.readyVisas);
        }
      })
      .catch(() => {});
  }, []);

  // State synchronization to localStorage (entries kept as fast local cache)
  useEffect(() => {
    localStorage.setItem('injaz_visa_entries', JSON.stringify(entries));
  }, [entries]);

  useEffect(() => {
    localStorage.setItem('injaz_visa_site_identity', JSON.stringify(siteIdentity));
  }, [siteIdentity]);

  useEffect(() => {
    const checkHash = () => {
      if (window.location.hash === '#adminmanager') {
        history.replaceState(null, '', window.location.pathname + window.location.search);
        setAdminPinInput('');
        setAdminPinError('');
        setViewMode('login');
      }
    };
    checkHash();
    window.addEventListener('hashchange', checkHash);
    return () => window.removeEventListener('hashchange', checkHash);
  }, []);

  // Color Theme Mapping Helper
  const theme = useMemo(() => {
    const preset = siteIdentity.themePreset || 'emerald';
    const maps = {
      emerald: {
        gradient: 'from-emerald-600 to-teal-500',
        bgGrad: 'from-emerald-800 to-slate-900',
        bgSolo: 'bg-emerald-600',
        hoverBg: 'hover:bg-emerald-700',
        textSolo: 'text-emerald-600',
        shadowColor: 'shadow-emerald-600/20',
        lightBg: 'bg-emerald-50',
        lightBorder: 'border-emerald-100',
        radarBg: 'bg-emerald-50',
        radarPulse: 'bg-emerald-400',
        radarPulseBack: 'bg-emerald-500/20',
        radarBorder: 'border-t-emerald-600',
        lightBgWithHalf: 'bg-emerald-50/50/40', // We can just use custom values or bg-emerald-50/30
        badgeArPhone: 'text-emerald-100 bg-emerald-500/20',
        badgeDarkText: 'text-emerald-700 bg-emerald-50 border-emerald-100/50',
        ringColor: 'focus:ring-emerald-500/10 focus:ring-emerald-500/20',
        textColorDark: 'text-emerald-800 border-emerald-100/60 bg-emerald-50',
        textColorDarker: 'text-emerald-950 border-emerald-100 bg-emerald-50/50'
      },
      blue: {
        gradient: 'from-blue-600 to-indigo-500',
        bgGrad: 'from-blue-800 to-slate-900',
        bgSolo: 'bg-blue-600',
        hoverBg: 'hover:bg-blue-700',
        textSolo: 'text-blue-600',
        shadowColor: 'shadow-blue-600/20',
        lightBg: 'bg-blue-50',
        lightBorder: 'border-blue-100',
        radarBg: 'bg-blue-50',
        radarPulse: 'bg-blue-400',
        radarPulseBack: 'bg-blue-500/20',
        radarBorder: 'border-t-blue-600',
        lightBgWithHalf: 'bg-blue-50/30',
        badgeArPhone: 'text-blue-100 bg-blue-500/20',
        badgeDarkText: 'text-blue-700 bg-blue-50 border-blue-100/50',
        ringColor: 'focus:ring-blue-500/10 focus:ring-blue-500/20',
        textColorDark: 'text-blue-800 border-blue-100/60 bg-blue-50',
        textColorDarker: 'text-blue-950 border-blue-100 bg-blue-50/50'
      },
      amber: {
        gradient: 'from-amber-600 to-orange-500',
        bgGrad: 'from-amber-800 to-slate-900',
        bgSolo: 'bg-amber-600',
        hoverBg: 'hover:bg-amber-700',
        textSolo: 'text-amber-600',
        shadowColor: 'shadow-amber-600/20',
        lightBg: 'bg-amber-50',
        lightBorder: 'border-amber-100',
        radarBg: 'bg-amber-50',
        radarPulse: 'bg-amber-400',
        radarPulseBack: 'bg-amber-500/20',
        radarBorder: 'border-t-amber-600',
        lightBgWithHalf: 'bg-amber-50/30',
        badgeArPhone: 'text-amber-100 bg-amber-500/20',
        badgeDarkText: 'text-amber-700 bg-amber-50 border-amber-100/50',
        ringColor: 'focus:ring-amber-500/10 focus:ring-amber-500/20',
        textColorDark: 'text-amber-800 border-amber-100/60 bg-amber-50',
        textColorDarker: 'text-amber-950 border-amber-100 bg-amber-50/50'
      },
      violet: {
        gradient: 'from-violet-600 to-fuchsia-500',
        bgGrad: 'from-violet-800 to-slate-900',
        bgSolo: 'bg-violet-600',
        hoverBg: 'hover:bg-violet-700',
        textSolo: 'text-violet-600',
        shadowColor: 'shadow-violet-600/20',
        lightBg: 'bg-violet-50',
        lightBorder: 'border-violet-100',
        radarBg: 'bg-violet-50',
        radarPulse: 'bg-violet-400',
        radarPulseBack: 'bg-violet-500/20',
        radarBorder: 'border-t-violet-600',
        lightBgWithHalf: 'bg-violet-50/30',
        badgeArPhone: 'text-violet-100 bg-violet-500/20',
        badgeDarkText: 'text-violet-700 bg-violet-50 border-violet-100/50',
        ringColor: 'focus:ring-violet-500/10 focus:ring-violet-500/20',
        textColorDark: 'text-violet-800 border-violet-100/60 bg-violet-50',
        textColorDarker: 'text-violet-950 border-violet-100 bg-violet-50/50'
      },
      rose: {
        gradient: 'from-rose-600 to-pink-500',
        bgGrad: 'from-rose-800 to-slate-900',
        bgSolo: 'bg-rose-600',
        hoverBg: 'hover:bg-rose-700',
        textSolo: 'text-rose-600',
        shadowColor: 'shadow-rose-600/20',
        lightBg: 'bg-rose-50',
        lightBorder: 'border-rose-100',
        radarBg: 'bg-rose-50',
        radarPulse: 'bg-rose-400',
        radarPulseBack: 'bg-rose-500/20',
        radarBorder: 'border-t-rose-600',
        lightBgWithHalf: 'bg-rose-50/30',
        badgeArPhone: 'text-rose-100 bg-rose-500/20',
        badgeDarkText: 'text-rose-700 bg-rose-50 border-rose-100/50',
        ringColor: 'focus:ring-rose-500/10 focus:ring-rose-500/20',
        textColorDark: 'text-rose-800 border-rose-100/60 bg-rose-50',
        textColorDarker: 'text-rose-950 border-rose-100 bg-rose-50/50'
      },
      custom: {
        gradient: 'theme-custom-gradient',
        bgGrad: 'theme-custom-bgGrad',
        bgSolo: 'theme-custom-bgSolo',
        hoverBg: 'theme-custom-hoverBg',
        textSolo: 'theme-custom-textSolo',
        shadowColor: 'theme-custom-shadowColor',
        lightBg: 'theme-custom-lightBg',
        lightBorder: 'theme-custom-lightBorder',
        radarBg: 'theme-custom-radarBg',
        radarPulse: 'theme-custom-radarPulse',
        radarPulseBack: 'theme-custom-radarPulseBack',
        radarBorder: 'theme-custom-radarBorder',
        lightBgWithHalf: 'theme-custom-radarBg',
        badgeArPhone: 'theme-custom-badgeArPhone',
        badgeDarkText: 'theme-custom-badgeDarkText',
        ringColor: 'theme-custom-ringColor',
        textColorDark: 'theme-custom-textColorDark',
        textColorDarker: 'theme-custom-textColorDarker'
      }
    };
    return maps[preset as keyof typeof maps] || maps.emerald;
  }, [siteIdentity.themePreset, siteIdentity.primaryColor, siteIdentity.secondaryColor]);

  // Dynamic WhatsApp Link Builder
  const getWhatsAppLink = (passportNum: string, applicantName: string = '') => {
    let phoneNum = siteIdentity.supportPhone.trim();
    phoneNum = phoneNum.replace(/[^\d]/g, '');
    if (phoneNum.length === 9 && !phoneNum.startsWith('967')) {
      phoneNum = '967' + phoneNum;
    }
    const msg = lang === 'ar'
      ? `مرحباً مكتب إنجاز، أود المتابعة بخصوص المعاملة الخاصة بالمسافر: ${applicantName}، رقم الجواز: ${passportNum}`
      : `Hello Injaz Office, I'm checking back regarding applicant: ${applicantName}, Passport No: ${passportNum}`;
    return `https://wa.me/${phoneNum}?text=${encodeURIComponent(msg)}`;
  };

  const getWhatsAppNotFoundLink = () => {
    let phoneNum = siteIdentity.supportPhone.trim();
    phoneNum = phoneNum.replace(/[^\d]/g, '');
    if (phoneNum.length === 9 && !phoneNum.startsWith('967')) {
      phoneNum = '967' + phoneNum;
    }
    const msg = lang === 'ar'
      ? `مرحباً مكتب إنجاز، قمت بالاستعلام عبر الجواز رقم [ ${publicPassport} ] ولم أعثر عليها، أرجو إفادتي بحالة المعاملة السجلية.`
      : `Hello Injaz Office, I searched passport: [ ${publicPassport} ] and got no matching records. Could you please check with my customer profile.`;
    return `https://wa.me/${phoneNum}?text=${encodeURIComponent(msg)}`;
  };

  const [isBulkOpen, setIsBulkOpen] = useState(false);
  const [isAddOpen, setIsAddOpen] = useState(false);
  const [bulkText, setBulkText] = useState('');
  
  // Custom defaults
  const [defaultNationality, setDefaultNationality] = useState('اليمن');
  const [defaultArrival, setDefaultArrival] = useState('عدن');
  const [defaultVisaType, setDefaultVisaType] = useState('عمل');

  const [addForm, setAddForm] = useState({
    name: '',
    passport: '',
    appNo: '',
    nationality: 'اليمن',
    arrival: 'عدن',
    visaType: 'عمل',
    phone: '',
    profession: '',
    officeId: 'general'
  });
  const [offices, setOffices] = useState<Office[]>([]);
  const [isOfficesView, setIsOfficesView] = useState(false);
  const [isAddOfficeForm, setIsAddOfficeForm] = useState(false);
  const [addOfficeForm, setAddOfficeForm] = useState({ name: '', phone: '' });
  const [editingOfficeId, setEditingOfficeId] = useState<string | null>(null);
  const [editOfficeForm, setEditOfficeForm] = useState({ name: '', phone: '' });

  // ── Custom Modal System ──
  const [alertModal, setAlertModal] = useState<{ msg: string; type: 'success' | 'error' | 'warning' | 'info' } | null>(null);
  const [confirmModal, setConfirmModal] = useState<{ msg: string; onConfirm: () => void } | null>(null);
  const showAlert = (msg: string, type: 'success' | 'error' | 'warning' | 'info' = 'info') => setAlertModal({ msg, type });
  const showConfirm = (msg: string, onConfirm: () => void) => setConfirmModal({ msg, onConfirm });
  const [selectedOfficeFilter, setSelectedOfficeFilter] = useState<string | null>(null);
  
  const [parsing, setParsing] = useState(false);
  const [lang, setLang] = useState<'ar' | 'en'>('ar');

  useEffect(() => {
    const name = lang === 'ar'
      ? (siteIdentity.officeNameAr || 'إنجاز للإرساليات')
      : (siteIdentity.officeNameEn || 'Injaz E-Services');
    document.title = name;
    // Update favicon
    const existing = document.querySelector('link[rel="icon"]') as HTMLLinkElement | null;
    if (siteIdentity.logo) {
      if (existing) {
        existing.href = siteIdentity.logo;
      } else {
        const link = document.createElement('link');
        link.rel = 'icon';
        link.href = siteIdentity.logo;
        document.head.appendChild(link);
      }
    } else if (existing) {
      existing.href = '/favicon.ico';
    }
  }, [siteIdentity.officeNameAr, siteIdentity.officeNameEn, siteIdentity.logo, lang]);

  // Sync input focus glow color with current theme
  useEffect(() => {
    const presetRgb: Record<string, string> = {
      emerald: '16, 185, 129',
      blue:    '59, 130, 246',
      amber:   '245, 158, 11',
      violet:  '139, 92, 246',
      rose:    '244, 63, 94',
    };
    let rgb = presetRgb[siteIdentity.themePreset] || presetRgb.emerald;
    if (siteIdentity.themePreset === 'custom' && siteIdentity.primaryColor) {
      const hex = siteIdentity.primaryColor.replace('#', '');
      if (hex.length === 6) {
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        rgb = `${r}, ${g}, ${b}`;
      }
    }
    document.documentElement.style.setProperty('--input-primary-rgb', rgb);
  }, [siteIdentity.themePreset, siteIdentity.primaryColor]);

  const [filterQuery, setFilterQuery] = useState('');

  const [logoDragActive, setLogoDragActive] = useState(false);

  const handleLogoFile = (file: File) => {
    if (!file) return;
    
    // Check if image
    if (!file.type.startsWith('image/')) {
      showAlert(lang === 'ar' ? 'عذراً، يجب اختيار ملف صورة صالح!' : 'Sorry, you must select a valid image file!', 'warning');
      return;
    }
    
    // Check size (800KB limit for localStorage safety)
    if (file.size > 800 * 1024) {
      showAlert(lang === 'ar' 
        ? 'حجم الصورة كبير جداً! يرجى اختيار شعار بحجم أقل من 800 كيلوبايت لضمان سرعة تحميل الموقع وحفظه بنجاح.' 
        : 'The image size is too large! Please choose a logo under 800 KB for fast page loading and secure storage.'
      , 'warning');
      return;
    }
    
    const reader = new FileReader();
    reader.onload = (e) => {
      if (e.target?.result && typeof e.target.result === 'string') {
        const base64String = e.target.result;
        setSiteIdentityUser(prev => {
          const updated = { ...prev, logo: base64String };
          // Auto-save to server immediately on logo upload
          fetch('/api/site-identity', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(updated),
          }).catch(() => {});
          try { localStorage.setItem('injaz_visa_site_identity', JSON.stringify(updated)); } catch {}
          return updated;
        });
        addSyncLog(lang === 'ar' ? 'تم تحديث وتحميل شعار الموقع المخصص بنجاح وحفظه في السيرفر.' : 'Custom website logo uploaded and saved to server successfully.');
      }
    };
    reader.readAsDataURL(file);
  };

  const handleLogoDrag = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setLogoDragActive(true);
    } else if (e.type === "dragleave") {
      setLogoDragActive(false);
    }
  };

  const handleLogoDrop = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setLogoDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleLogoFile(e.dataTransfer.files[0]);
    }
  };
  const [activeTab, setActiveTab] = useState<'dashboard' | 'add' | 'all' | 'identity' | 'excel_sync' | 'security' | 'customers' | 'files' | 'auto_check'>('dashboard');

  // ── Auto-check scheduler state ────────────────────────────────────────────
  const [autoCheckEnabled, setAutoCheckEnabled] = useState(false);
  const [autoCheckTime, setAutoCheckTime] = useState<string>('08:00');
  const [autoCheckInterval, setAutoCheckInterval] = useState<number>(1);
  const [autoCheckLastRun, setAutoCheckLastRun] = useState<string>('');
  const [autoCheckNextRun, setAutoCheckNextRun] = useState<string>('');
  const [autoCheckLog, setAutoCheckLog] = useState<{time: string; msg: string}[]>([]);
  const autoCheckRef = useRef<NodeJS.Timeout | null>(null);
  const [pdfFiles, setPdfFiles] = useState<{filename: string; url: string; size: number; createdAt: number}[]>([]);
  const [pdfFilesLoading, setPdfFilesLoading] = useState(false);
  const [visaFiles, setVisaFiles] = useState<{filename: string; url: string; size: number; createdAt: number}[]>([]);
  const [visaFilesLoading, setVisaFilesLoading] = useState(false);
  const [filesSubPage, setFilesSubPage] = useState<null | 'prints' | 'visas'>(null);
  const [filesSearch, setFilesSearch] = useState('');
  const [syncSubView, setSyncSubView] = useState<'menu' | 'telegram' | 'google_sheets'>('menu');
  const [allTabFilter, setAllTabFilter] = useState<string>('all');
  const [isSectionSelected, setIsSectionSelected] = useState(false);
  const [isVisaTypeView, setIsVisaTypeView] = useState(false);
  const [selectedVisaType, setSelectedVisaType] = useState<string | null>(null);
  const [isStatsView, setIsStatsView] = useState(false);
  const [isArchiveView, setIsArchiveView] = useState(false);
  const [archiveVisaTypeFilter, setArchiveVisaTypeFilter] = useState<string | null>(null);
  const [archiveOfficeFilter, setArchiveOfficeFilter] = useState<string | null>(null);
  const [archiveInCustomerView, setArchiveInCustomerView] = useState(false);
  const [archiveSearchQuery, setArchiveSearchQuery] = useState('');
  const [archiveConfirmEntry, setArchiveConfirmEntry] = useState<VisaEntry | null>(null);
  const [archiveSelectMode, setArchiveSelectMode] = useState(false);
  const [archiveSelectedIds, setArchiveSelectedIds] = useState<Set<string>>(new Set());
  const [expandedCardId, setExpandedCardId] = useState<string | null>(null);
  const [editingCardId, setEditingCardId] = useState<string | null>(null);
  const [isSelectMode, setIsSelectMode] = useState(false);
  const [selectedEntryIds, setSelectedEntryIds] = useState<Set<string>>(new Set());

  // ── Transaction Status ──
  type TxStatusType = 'مستمرة' | 'موقفة' | 'بدون';
  interface TxStatusOption { id: string; label: string; type: TxStatusType; hasStatsCard?: boolean; statsIsolate?: boolean; }
  const [txStatusOptions, setTxStatusOptions] = useState<TxStatusOption[]>(() => {
    try { const s = localStorage.getItem('txStatusOptions'); if (s) return JSON.parse(s); } catch {}
    return [
      { id: 'ts1', label: 'مستمرة', type: 'مستمرة' },
      { id: 'ts2', label: 'موقفة', type: 'موقفة' },
      { id: 'ts3', label: 'مرتجع تم الترحيل من جديد', type: 'مستمرة' },
      { id: 'ts4', label: 'جاهزة للاستلام', type: 'مستمرة' },
    ];
  });

  // Load txStatusOptions from server on startup (overrides localStorage)
  useEffect(() => {
    fetch('/api/tx-status-options')
      .then(r => r.json())
      .then((data: TxStatusOption[]) => {
        if (Array.isArray(data) && data.length > 0) {
          setTxStatusOptions(data);
          try { localStorage.setItem('txStatusOptions', JSON.stringify(data)); } catch {}
        }
      })
      .catch(() => {});
  }, []);

  const saveTxStatusOptions = (opts: TxStatusOption[]) => {
    setTxStatusOptions(opts);
    try { localStorage.setItem('txStatusOptions', JSON.stringify(opts)); } catch {}
    fetch('/api/tx-status-options', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ options: opts }),
    }).catch(() => {});
  };
  const [txStatusModalEntry, setTxStatusModalEntry] = useState<VisaEntry | null>(null);
  const [txStatusForm, setTxStatusForm] = useState({ label: 'مستمرة', type: 'مستمرة' as TxStatusType, date: new Date().toISOString().split('T')[0], note: '' });
  const [txManageMode, setTxManageMode] = useState(false);
  const [txNewLabel, setTxNewLabel] = useState('');
  const [txNewType, setTxNewType] = useState<TxStatusType>('مستمرة');
  const [txNewHasStatsCard, setTxNewHasStatsCard] = useState(false);
  const [txNewIsolate, setTxNewIsolate] = useState(false);
  const [txEditingOptId, setTxEditingOptId] = useState<string | null>(null);
  const [txEditLabel, setTxEditLabel] = useState('');
  const [txEditType, setTxEditType] = useState<TxStatusType>('مستمرة');
  const openTxStatusModal = (entry: VisaEntry) => {
    setTxStatusModalEntry(entry);
    setTxStatusForm({
      label: entry.txStatus || 'مستمرة',
      type: entry.txStatusType || 'مستمرة',
      date: entry.txStatusDate || new Date().toISOString().split('T')[0],
      note: entry.txStatusNote || '',
    });
    setTxManageMode(false);
  };
  const saveTxStatus = () => {
    if (!txStatusModalEntry) return;
    setEntries(prev => prev.map(e => e.id === txStatusModalEntry.id ? { ...e, txStatus: txStatusForm.label, txStatusType: txStatusForm.type, txStatusDate: txStatusForm.date, txStatusNote: txStatusForm.note } : e));
    setTxStatusModalEntry(null);
  };
  const txTypeBg = (type?: TxStatusType) => type === 'موقفة' ? 'bg-red-100 text-red-700 border-red-200' : type === 'بدون' ? 'bg-slate-100 text-slate-500 border-slate-200' : 'bg-emerald-100 text-emerald-700 border-emerald-200';
  const txTypeStripBg = (type?: TxStatusType) => type === 'موقفة' ? 'bg-red-50 border-red-100' : type === 'بدون' ? 'bg-slate-50 border-slate-100' : 'bg-emerald-50 border-emerald-100';
  const txTypeDot = (type?: TxStatusType) => type === 'موقفة' ? 'bg-red-400' : type === 'بدون' ? 'bg-slate-300' : 'bg-emerald-400';
  const [showBulkMenu, setShowBulkMenu] = useState(false);
  const [showTransferModal, setShowTransferModal] = useState(false);
  const [bulkTransferOfficeId, setBulkTransferOfficeId] = useState('');
  const [sortOrder, setSortOrder] = useState<'newest' | 'oldest' | null>(null);
  const [docNumberFilter, setDocNumberFilter] = useState<'with_doc' | 'without_doc' | null>(null);
  const [showFilterMenu, setShowFilterMenu] = useState(false);
  const [checkingHealthCertIds, setCheckingHealthCertIds] = useState<Set<string>>(new Set());
  const [healthCertOverlay, setHealthCertOverlay]         = useState(false);
  const [healthCertOverlayName, setHealthCertOverlayName] = useState('');
  const [healthCertOverlayStep, setHealthCertOverlayStep] = useState(0);
  const [healthCertOverlayError, setHealthCertOverlayError] = useState('');
  const [checkActionOpenId, setCheckActionOpenId] = useState<string | null>(null);
  const [isOfficeDashboardView, setIsOfficeDashboardView] = useState(false);
  const [isExportSectionModalOpen, setIsExportSectionModalOpen] = useState(false);
  const [customerSubTab, setCustomerSubTab] = useState<'none' | 'add' | 'edit' | 'excel_import'>('none');
  const [customerExcelSuccess, setCustomerExcelSuccess] = useState(false);
  const [customerExcelAddedCount, setCustomerExcelAddedCount] = useState(0);
  const [customerExcelSkippedCount, setCustomerExcelSkippedCount] = useState(0);
  const [customerExcelAddedEntries, setCustomerExcelAddedEntries] = useState<VisaEntry[]>([]);
  const [customerExcelDragActive, setCustomerExcelDragActive] = useState(false);
  const [customerSearchQuery, setCustomerSearchQuery] = useState('');
  const [selectedCustomerToEdit, setSelectedCustomerToEdit] = useState<VisaEntry | null>(null);

  // Form states for Add Customer inside Customer Management:
  const [newCustName, setNewCustName] = useState('');
  const [newCustPassport, setNewCustPassport] = useState('');
  const [newCustAppNo, setNewCustAppNo] = useState('');
  const [newCustPhone, setNewCustPhone] = useState('');
  const [newCustNationality, setNewCustNationality] = useState('اليمن');
  const [newCustArrival, setNewCustArrival] = useState('عدن');
  const [newCustVisaType, setNewCustVisaType] = useState('عمل');
  const [newCustProfession, setNewCustProfession] = useState('');
  const [newCustVisaNumber, setNewCustVisaNumber] = useState('');
  const [newCustStatus, setNewCustStatus] = useState<VisaEntry['status']>('Found');
  const [newCustStatusText, setNewCustStatusText] = useState('تم اصدار التأشيرة (Issued)');
  const [newCustApplicantData, setNewCustApplicantData] = useState('');
  const [newCustOfficeId, setNewCustOfficeId] = useState('');
  const [newCustCustomId, setNewCustCustomId] = useState('');
  const [newCustReceiveDate, setNewCustReceiveDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [newCustSaving, setNewCustSaving] = useState(false);

  // Form states for Editing Customer:
  const [editCustName, setEditCustName] = useState('');
  const [editCustPassport, setEditCustPassport] = useState('');
  const [editCustAppNo, setEditCustAppNo] = useState('');
  const [editCustPhone, setEditCustPhone] = useState('');
  const [editCustNationality, setEditCustNationality] = useState('');
  const [editCustArrival, setEditCustArrival] = useState('');
  const [editCustVisaType, setEditCustVisaType] = useState('');
  const [editCustProfession, setEditCustProfession] = useState('');
  const [editCustVisaNumber, setEditCustVisaNumber] = useState('');
  const [editCustStatus, setEditCustStatus] = useState<VisaEntry['status']>('Found');
  const [editCustStatusText, setEditCustStatusText] = useState('');
  const [editCustApplicantData, setEditCustApplicantData] = useState('');
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);

  // View modes
  const [viewMode, setViewMode] = useState<'public' | 'admin' | 'login'>(() => {
    if (typeof window === 'undefined') return 'public';
    if (window.location.hash === '#adminmanager') {
      history.replaceState(null, '', window.location.pathname + window.location.search);
      sessionStorage.setItem('injaz_view_mode', 'login');
      document.documentElement.classList.add('dark');
      return 'login';
    }
    const saved = sessionStorage.getItem('injaz_view_mode') as 'public' | 'admin' | 'login' | null;
    if (saved === 'login' || saved === 'admin') {
      document.documentElement.classList.add('dark');
      return saved;
    }
    document.documentElement.classList.remove('dark');
    return 'public';
  });

  useEffect(() => {
    if (viewMode === 'public') {
      document.documentElement.classList.remove('dark');
      sessionStorage.removeItem('injaz_view_mode');
    } else {
      document.documentElement.classList.add('dark');
      sessionStorage.setItem('injaz_view_mode', viewMode);
    }
  }, [viewMode]);

  const [isAdminLoginOpen, setIsAdminLoginOpen] = useState(false);
  const [adminUsernameInput, setAdminUsernameInput] = useState('');
  const [adminPinInput, setAdminPinInput] = useState('');
  const [adminPinError, setAdminPinError] = useState('');
  const [identitySuccess, setIdentitySuccess] = useState(false);

  const saveIdentity = () => {
    localStorage.setItem('injaz_visa_site_identity', JSON.stringify(siteIdentity));
    fetch('/api/site-identity', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(siteIdentity),
    }).catch(() => {});
    setIdentitySuccess(true);
    setTimeout(() => setIdentitySuccess(false), 2500);
  };

  // Public portal search states
  const [publicPassport, setPublicPassport] = useState('');
  const [publicPhone, setPublicPhone] = useState('');
  const [isPublicSearching, setIsPublicSearching] = useState(false);
  const [publicSearchResult, setPublicSearchResult] = useState<VisaEntry | null>(null);
  const [publicHasSearched, setPublicHasSearched] = useState(false);

  // Google Sheets API Synchronization State
  const [googleSpreadsheetId, setGoogleSpreadsheetId] = useState('');
  const [googleApiKey, setGoogleApiKey] = useState('');
  const [googleSheetName, setGoogleSheetName] = useState('Sheet1');
  const [lastSyncTime, setLastSyncTime] = useState('لم تتم المزامنة بعد');
  const [syncLogs, setSyncLogs] = useState<string[]>(['نظام المزامنة الذكي جاهز ولم يبدأ بعد...']);
  const [recentUpdates, setRecentUpdates] = useState<any[]>([]);
  const [readyVisas, setReadyVisas] = useState<VisaEntry[]>([]);
  const [syncStatus, setSyncStatus] = useState<'idle' | 'syncing' | 'success' | 'error'>('idle');

  // Telegram Bot Configuration State
  const [telegramToken, setTelegramToken] = useState('');
  const [telegramEnabled, setTelegramEnabled] = useState(false);
  const [telegramAlertChatId, setTelegramAlertChatId] = useState('');
  const [telegramPrintArchiveChatId, setTelegramPrintArchiveChatId] = useState('');
  const [telegramVisaArchiveChatId, setTelegramVisaArchiveChatId] = useState('');

  // Admin Security Credentials State
  const [adminUsername, setAdminUsername] = useState('admin');
  const [adminPassword, setAdminPassword] = useState('1234');

  // User Session & multi-user role management state
  const [currentUser, setCurrentUser] = useState<any>(() => {
    const saved = localStorage.getItem('injaz_visa_current_user');
    return saved ? JSON.parse(saved) : null;
  });
  const [usersList, setUsersList] = useState<any[]>([]);

  // User creation / edit form inputs state
  const [newUserName, setNewUserName] = useState('');
  const [newUserUsername, setNewUserUsername] = useState('');
  const [newUserPassword, setNewUserPassword] = useState('');
  const [newUserPermissions, setNewUserPermissions] = useState<string[]>(['update_checking', 'add_clients', 'edit_clients']);
  const [editingUserId, setEditingUserId] = useState<string | null>(null);



  // KSA Visa Track States
  const [ksaVisaModalOpen, setKsaVisaModalOpen] = useState(false);
  const [ksaVisaEntry, setKsaVisaEntry] = useState<VisaEntry | null>(null);
  const [ksaVisaSessionId, setKsaVisaSessionId] = useState('');
  const [ksaVisaCaptchaKey, setKsaVisaCaptchaKey] = useState(0);
  const [ksaVisaCaptchaCode, setKsaVisaCaptchaCode] = useState('');
  const [ksaVisaLoading, setKsaVisaLoading] = useState(false);
  const [ksaVisaInitLoading, setKsaVisaInitLoading] = useState(false);
  const [ksaVisaError, setKsaVisaError] = useState('');
  const [ksaVisaSuccess, setKsaVisaSuccess] = useState(false);
  const [ksaVisaNumber, setKsaVisaNumber] = useState('');
  const [ksaVisaPassportNumber, setKsaVisaPassportNumber] = useState('');
  const [ksaVisaCheckingOverlay, setKsaVisaCheckingOverlay] = useState(false);
  const [ksaVisaSaveSuccessAnim, setKsaVisaSaveSuccessAnim] = useState(false);
  const [ksaVisaCheckingError, setKsaVisaCheckingError] = useState('');
  const [ksaVisaCheckingStep, setKsaVisaCheckingStep] = useState(0);

  // Smart MOFA Live Query States
  const [mofaFetchModalOpen, setMofaFetchModalOpen] = useState(false);
  const [mofaFetchEntry, setMofaFetchEntry] = useState<VisaEntry | null>(null);
  const [mofaPasteText, setMofaPasteText] = useState('');

  // MOFA Live API states
  const [mofaSessionId, setMofaSessionId] = useState('');
  const [mofaCaptchaPath, setMofaCaptchaPath] = useState('');
  const [mofaCaptchaKey, setMofaCaptchaKey] = useState(0);
  const [mofaNationalities, setMofaNationalities] = useState<{value:string;label:string}[]>([]);
  const [mofaVisaTypes, setMofaVisaTypes] = useState<{value:string;label:string}[]>([]);
  const [mofaEmbassies, setMofaEmbassies] = useState<{value:string;label:string}[]>([]);
  const [mofaSelectedNat, setMofaSelectedNat] = useState('YEM');
  const [mofaSelectedVt, setMofaSelectedVt] = useState('1');
  const [mofaSelectedEmb, setMofaSelectedEmb] = useState('302');
  const [mofaCaptchaCode, setMofaCaptchaCode] = useState('');
  const [mofaLoading, setMofaLoading] = useState(false);
  const [mofaInitLoading, setMofaInitLoading] = useState(false);
  const [mofaError, setMofaError] = useState('');
  const [mofaResult, setMofaResult] = useState<any>(null);
  const [mofaCheckingOverlay, setMofaCheckingOverlay] = useState(false);
  const [mofaSaveSuccessAnim, setMofaSaveSuccessAnim] = useState(false);
  const [mofaCheckingError, setMofaCheckingError] = useState('');
  const [mofaCheckingStep, setMofaCheckingStep] = useState(0);

  // Bulk Sequential Check States
  const [bulkCheckRunning, setBulkCheckRunning] = useState(false);
  const [bulkCheckIndex, setBulkCheckIndex] = useState(0);
  const [bulkCheckTotal, setBulkCheckTotal] = useState(0);
  const [bulkCheckCurrentName, setBulkCheckCurrentName] = useState('');
  const [bulkCheckCurrentStatus, setBulkCheckCurrentStatus] = useState('');
  const [bulkCheckResults, setBulkCheckResults] = useState<{name: string; status: 'visa_found'|'mofa_only'|'no_print'|'error'; msg: string}[]>([]);
  const bulkCheckStopRef = useRef(false);
  const healthCertCancelRef = useRef(false);
  const healthCertEntryIdRef = useRef<string | null>(null);
  const longPressTimer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const filterMenuRef = useRef<HTMLDivElement>(null);
  const pendingBulkCount = entries.filter(e => !e.archived && !e.visaImageUrl && e.txStatusType !== 'موقفة').length;

  // Permissions helper function
  const hasPermission = (perm: string): boolean => {
    if (!currentUser) return false;
    if (currentUser.id === 'master-admin' || currentUser.permissions?.includes('admin')) {
      return true;
    }
    return currentUser.permissions?.includes(perm);
  };

  // 1. Two-way Database Sync: load from server, or push if server database is empty
  useEffect(() => {
    const initDatabaseSync = async () => {
      try {
        const response = await fetch('/api/entries');
        if (response.ok) {
          const serverData = await response.json();
          const localRaw = localStorage.getItem('injaz_visa_entries');
          const localData: any[] = localRaw ? JSON.parse(localRaw) : [];

          if (serverData && Array.isArray(serverData) && serverData.length > 0) {
            // Compare newest updatedAtUnix between server and local to pick the freshest source
            const maxServer = Math.max(...serverData.map((e: any) => e.updatedAtUnix || 0));
            const maxLocal  = localData.length > 0 ? Math.max(...localData.map((e: any) => e.updatedAtUnix || 0)) : 0;

            if (maxLocal > maxServer) {
              // Local is newer — push local to server and keep local state
              await fetch('/api/entries', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ entries: localData })
              });
              // State already initialized from localStorage, no need to setEntries
            } else {
              // Server is equal or newer — use server data
              setEntries(serverData);
              localStorage.setItem('injaz_visa_entries', JSON.stringify(serverData));
            }
          } else {
            // Seed from client on empty server database (first boot bootstrap)
            if (localData.length > 0) {
              setEntries(localData);
              await fetch('/api/entries', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ entries: localData })
              });
            }
          }
        }
      } catch (err) {
        console.error("Failed to connect & sync database with backend storage:", err);
      }
    };
    initDatabaseSync();
  }, []);

  // Load offices from server on mount
  useEffect(() => {
    fetch('/api/offices').then(r => r.json()).then(data => {
      if (Array.isArray(data)) setOffices(data);
    }).catch(() => {});
  }, []);

  // Sync offices to server whenever they change
  useEffect(() => {
    if (offices.length === 0) return;
    fetch('/api/offices', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ offices })
    }).catch(() => {});
  }, [offices]);

  // Close filter menu on outside click
  useEffect(() => {
    if (!showFilterMenu) return;
    const handler = (e: MouseEvent) => {
      if (filterMenuRef.current && !filterMenuRef.current.contains(e.target as Node)) {
        setShowFilterMenu(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [showFilterMenu]);

  // Reset integration sub-view to 'menu' when entering the excel_sync tab
  useEffect(() => {
    if (activeTab === 'excel_sync') {
      setSyncSubView('menu');
    }
    if (activeTab === 'files') {
      setPdfFilesLoading(true);
      fetch('/api/prints')
        .then(r => r.json())
        .then(data => { setPdfFiles(data.files || []); setPdfFilesLoading(false); })
        .catch(() => setPdfFilesLoading(false));
    }
  }, [activeTab]);

  // ── Auto-check scheduler engine ───────────────────────────────────────────
  useEffect(() => {
    const computeNextRun = (lastRun: string, time: string, intervalDays: number): string => {
      const [h, m] = time.split(':').map(Number);
      const base = lastRun ? new Date(lastRun) : new Date(0);
      const next = new Date(base);
      next.setDate(next.getDate() + intervalDays);
      next.setHours(h, m, 0, 0);
      if (next <= new Date()) {
        const now = new Date();
        now.setHours(h, m, 0, 0);
        if (now <= new Date()) now.setDate(now.getDate() + intervalDays);
        return now.toISOString();
      }
      return next.toISOString();
    };

    if (autoCheckEnabled) {
      const next = computeNextRun(autoCheckLastRun, autoCheckTime, autoCheckInterval);
      setAutoCheckNextRun(next);
    }

    const tick = () => {
      if (!autoCheckEnabled) return;
      const now = new Date();
      const [h, m] = autoCheckTime.split(':').map(Number);
      if (now.getHours() !== h || now.getMinutes() !== m) return;
      const lastRunDate = autoCheckLastRun ? new Date(autoCheckLastRun) : new Date(0);
      const diffDays = (now.getTime() - lastRunDate.getTime()) / (1000 * 60 * 60 * 24);
      if (diffDays < autoCheckInterval) return;
      const nowStr = now.toISOString();
      const logEntry = { time: now.toLocaleString('ar-SA'), msg: 'تم تشغيل التشييك التلقائي تلقائياً' };
      setAutoCheckLastRun(nowStr);
      setAutoCheckLog(prev => {
        const updated = [logEntry, ...prev].slice(0, 30);
        return updated;
      });
      const next = computeNextRun(nowStr, autoCheckTime, autoCheckInterval);
      setAutoCheckNextRun(next);
      checkAll();
    };

    if (autoCheckRef.current) clearInterval(autoCheckRef.current);
    autoCheckRef.current = setInterval(tick, 60000);
    return () => { if (autoCheckRef.current) clearInterval(autoCheckRef.current); };
  }, [autoCheckEnabled, autoCheckTime, autoCheckInterval, autoCheckLastRun]);

  // 2. Active Sync whenever entries list is touched in client
  useEffect(() => {
    const syncWithServer = async () => {
      try {
        await fetch('/api/entries', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ entries })
        });
      } catch (err) {
        console.error("Active write-through failure for server entries DB:", err);
      }
    };
    if (entries.length > 0) {
      syncWithServer();
    }
  }, [entries]);

  // 3. Telegram config & admin credentials bootstrap
  useEffect(() => {
    const loadTelegramSettings = async () => {
      try {
        const response = await fetch('/api/telegram-config');
        if (response.ok) {
          const data = await response.json();
          if (data && data.config) {
            setTelegramToken(data.config.token || '');
            setTelegramEnabled(!!data.config.enabled);
            setTelegramAlertChatId(data.config.alertChatId || '');
            setTelegramPrintArchiveChatId(data.config.printArchiveChatId || '');
            setTelegramVisaArchiveChatId(data.config.visaArchiveChatId || '');
          }
        }
      } catch (err) {
        console.error("Failed to bootstrap Telegram Bot credentials from server:", err);
      }
    };

    const loadAdminCredentials = async () => {
      try {
        const response = await fetch('/api/credentials');
        if (response.ok) {
          const data = await response.json();
          if (data) {
            setAdminUsername(data.username || 'admin');
            setAdminPassword(data.password || '1234');
          }
        }
      } catch (err) {
        console.error("Failed to bootstrap admin credentials from server:", err);
      }
    };

    const loadUserList = async () => {
      try {
        const response = await fetch('/api/users');
        if (response.ok) {
          const data = await response.json();
          setUsersList(data || []);
        }
      } catch (err) {
        console.error("Failed to bootstrap users list from server:", err);
      }
    };

    loadTelegramSettings();
    loadAdminCredentials();
    loadUserList();
  }, []);

  const handleSaveTelegramConfig = async () => {
    try {
      const response = await fetch('/api/telegram-config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          token: telegramToken,
          enabled: telegramEnabled,
          alertChatId: telegramAlertChatId,
          printArchiveChatId: telegramPrintArchiveChatId,
          visaArchiveChatId: telegramVisaArchiveChatId
        })
      });
      if (response.ok) {
        addSyncLog("تم بنجاح تحديث ورابط بوت تليجرام، وتنشيط المزامنة الحية بقاعدة البيانات.");
        showAlert(lang === 'ar' ? 'تم حفظ إعدادات البوت والاتصال بالخادم بنجاح!' : 'Telegram Bot API Config stored and initialized on server!', 'success');
      } else {
        throw new Error('Server returned failure status');
      }
    } catch (err: any) {
      console.error(err);
      addSyncLog(`[خطأ] فشل تشغيل ورابط تليجرام: ${err.message}`);
    }
  };

  const handleSaveCredentials = async () => {
    try {
      const response = await fetch('/api/credentials', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: adminUsername, password: adminPassword })
      });
      if (response.ok) {
        showAlert(lang === 'ar' ? 'تم تحديث حساب الإدارة (يوزر وباسورد) بنجاح!' : 'Admin security credentials updated successfully!', 'success');
      } else {
        const data = await response.json();
        showAlert(lang === 'ar' ? `فشل تحديث البيانات: ${data.error}` : `Failed to update credentials: ${data.error}`, 'error');
      }
    } catch (err: any) {
      showAlert(lang === 'ar' ? `حدث خطأ بالاتصال: ${err.message}` : `Network connection error: ${err.message}`, 'error');
    }
  };

  const handleLoginSubmit = async () => {
    try {
      const response = await fetch('/api/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: adminUsernameInput, password: adminPinInput })
      });
      if (response.ok) {
        const data = await response.json();
        setCurrentUser(data.user);
        localStorage.setItem('injaz_visa_current_user', JSON.stringify(data.user));
        setViewMode('admin');
        setIsAdminLoginOpen(false);
        setAdminUsernameInput('');
        setAdminPinInput('');
        setAdminPinError('');
      } else {
        const errorData = await response.json();
        setAdminPinError(errorData.error || (lang === 'ar' ? 'بيانات الدخول غير صحيحة' : 'Invalid login credentials'));
      }
    } catch (err: any) {
      setAdminPinError(lang === 'ar' ? 'حدث خطأ بالاتصال بالخادم للتحقق من البيانات' : 'Server connection failed during login verification');
    }
  };

  const handleUserFormSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newUserName.trim() || !newUserUsername.trim() || (!editingUserId && !newUserPassword.trim())) {
      showAlert(lang === 'ar' ? 'يرجى ملأ جميع الحقول الأساسية!' : 'Please fill all required fields!', 'warning');
      return;
    }

    const updatedUsers = [...usersList];
    const usernameNorm = newUserUsername.toLowerCase().trim();

    if (editingUserId) {
      const idx = updatedUsers.findIndex(u => u.id === editingUserId);
      if (idx > -1) {
        updatedUsers[idx] = {
          ...updatedUsers[idx],
          name: newUserName.trim(),
          username: usernameNorm,
          permissions: newUserPermissions,
          ...(newUserPassword.trim() ? { password: newUserPassword.trim() } : {})
        };
      }
    } else {
      // Check duplicate
      const duplicate = updatedUsers.some(u => u.username.toLowerCase() === usernameNorm);
      if (duplicate) {
        showAlert(lang === 'ar' ? 'اسم المستخدم هذا مكرر بالفعل!' : 'This username is already taken!', 'warning');
        return;
      }
      updatedUsers.push({
        id: `user-${Date.now()}`,
        name: newUserName.trim(),
        username: usernameNorm,
        password: newUserPassword.trim(),
        permissions: newUserPermissions
      });
    }

    try {
      const response = await fetch('/api/users', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ users: updatedUsers })
      });
      if (response.ok) {
        setUsersList(updatedUsers);
        setNewUserName('');
        setNewUserUsername('');
        setNewUserPassword('');
        setNewUserPermissions(['update_checking', 'add_clients', 'edit_clients']);
        setEditingUserId(null);
        showAlert(lang === 'ar' ? 'تم حفظ وتنشيط حساب المستخدم بنجاح!' : 'User account successfully stored!', 'success');
      } else {
        const errorData = await response.json();
        showAlert(lang === 'ar' ? `فشل الحفظ: ${errorData.error}` : `Failed to save: ${errorData.error}`, 'error');
      }
    } catch (err: any) {
      showAlert(lang === 'ar' ? `خطأ بالاتصال بالسيرفر: ${err.message}` : `Server connection error: ${err.message}`, 'error');
    }
  };

  const handleDeleteUser = async (userId: string) => {
    if (userId === 'admin-default') {
      showAlert(lang === 'ar' ? 'لا يمكن حذف حساب المسؤول الافتراضي الأساسي!' : 'Default master admin account cannot be deleted!', 'warning');
      return;
    }
    showConfirm(
      lang === 'ar' ? 'هل أنت متأكد من رغبتك في حذف حساب المستخدم هذا نهائياً؟' : 'Are you sure you want to permanently delete this user profile?',
      async () => {
        const updatedUsers = usersList.filter(u => u.id !== userId);
        try {
          const response = await fetch('/api/users', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ users: updatedUsers })
          });
          if (response.ok) {
            setUsersList(updatedUsers);
            showAlert(lang === 'ar' ? 'تم حذف حساب المستخدم بنجاح.' : 'User account removed successfully.', 'success');
          } else {
            const errorData = await response.json();
            showAlert(lang === 'ar' ? `فشل الحفظ: ${errorData.error}` : `Failed error: ${errorData.error}`, 'error');
          }
        } catch (err) {
          console.error(err);
        }
      }
    );
  };

  // Persist Google Sheets config to server whenever any field changes
  useEffect(() => {
    fetch('/api/google-sheets-config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        spreadsheetId: googleSpreadsheetId,
        apiKey: googleApiKey,
        sheetName: googleSheetName,
        lastSyncTime,
        syncLogs,
        recentUpdates,
        readyVisas,
      }),
    }).catch(() => {});
  }, [googleSpreadsheetId, googleApiKey, googleSheetName, lastSyncTime, syncLogs, recentUpdates, readyVisas]);

  const addSyncLog = (msg: string) => {
    const timestamp = new Date().toLocaleTimeString('ar-YE', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    setSyncLogs(prev => [`[${timestamp}] ${msg}`, ...prev.slice(0, 49)]);
  };

  useEffect(() => {
    const checkExpirations = () => {
      const now = Date.now();
      setRecentUpdates(prev => {
        const filtered = prev.filter(item => {
          const itemTime = new Date(item.copiedAt).getTime();
          const isExpired = now - itemTime >= 24 * 60 * 60 * 1000;
          return !isExpired;
        });
        if (filtered.length !== prev.length) {
          addSyncLog(`[النظام] تم تلقائياً تصفية وحذف معاملات مضى عليها أكثر من 24 ساعة من الصفحة الثانية (التحديثات).`);
        }
        return filtered;
      });
    };

    checkExpirations();
    const interval = setInterval(checkExpirations, 60000);
    return () => clearInterval(interval);
  }, []);

  function parseCSV(text: string) {
    const lines = text.split('\n');
    return lines.map(line => {
      const result = [];
      let current = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') {
          inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
          result.push(current.trim());
          current = '';
        } else {
          current += char;
        }
      }
      result.push(current.trim());
      return result;
    });
  }

  const downloadThreeSheetXLSX = () => {
    if (!hasPermission('download_data')) {
      showAlert(lang === 'ar' ? 'عذراً، ليس لديك صلاحية "تنزيل بيانات العملاء والكشوفات" لتنزيل هذا الملف التقريري!' : 'Sorry, you do not have permission to download client data and reports!', 'warning');
      return;
    }
    try {
      const s1Data = entries.map(e => ({
        'رقم جواز العميل': e.passportNumber,
        'رقم الهاتف': e.phoneNumber || '',
        'اسم العميل': e.applicantName || '',
        'المهنة': e.profession || '',
        'حالة المعاملة من منصة التأشيرات': e.statusText || 'قيد الانتظار (Idle)',
        'حالة الكرت الصحي': e.healthCertStatus || '---',
        'رقم الطلب': e.applicationNumber,
        'رقم التأشيرة': e.visaNumber || '---',
        'الجنسية': e.nationality,
        'جهة الوصول': e.arrivalPoint,
        'نوع التأشيرة': e.visaType,
        'بيانات إضافية': e.applicantData || ''
      }));

      const s2Data = recentUpdates.map(u => ({
        'رقم جواز العميل': u.passportNumber,
        'اسم العميل': u.applicantName || '',
        'المهنة': u.profession || '',
        'رقم الطلب': u.applicationNumber,
        'الحالة السابقة': u.oldStatusText || '---',
        'الحالة الجديدة': u.newStatusText || '---',
        'تاريخ ووقت التحديث': new Date(u.copiedAt).toLocaleString('ar-YE'),
        'الوقت المتبقي للصلاحية': 'صالح لمدة 24 ساعة كحد أقصى'
      }));

      const s3Data = readyVisas.map(e => ({
        'رقم جواز العميل': e.passportNumber,
        'رقم الهاتف': e.phoneNumber || '',
        'اسم العميل': e.applicantName || '',
        'المهنة': e.profession || '',
        'حالة المعاملة من منصة التأشيرات': e.statusText || 'تم اصدار التأشيرة (Issued)',
        'حالة الكرت الصحي': e.healthCertStatus || '---',
        'رقم الطلب': e.applicationNumber,
        'رقم التأشيرة': e.visaNumber || '---',
        'الجنسية': e.nationality,
        'جهة الوصول': e.arrivalPoint,
        'نوع التأشيرة': e.visaType,
        'تاريخ ووقت التحديث': new Date().toLocaleString('ar-YE')
      }));

      const s1Headers = [
        'رقم جواز العميل',
        'رقم الهاتف',
        'اسم العميل',
        'المهنة',
        'حالة المعاملة من منصة التأشيرات',
        'حالة الكرت الصحي',
        'رقم الطلب',
        'رقم التأشيرة',
        'الجنسية',
        'جهة الوصول',
        'نوع التأشيرة',
        'بيانات إضافية'
      ];

      const s2Headers = [
        'رقم جواز العميل',
        'اسم العميل',
        'المهنة',
        'رقم الطلب',
        'الحالة السابقة',
        'الحالة الجديدة',
        'تاريخ ووقت التحديث',
        'الوقت المتبقي للصلاحية'
      ];

      const s3Headers = [
        'رقم جواز العميل',
        'رقم الهاتف',
        'اسم العميل',
        'المهنة',
        'حالة المعاملة من منصة التأشيرات',
        'حالة الكرت الصحي',
        'رقم الطلب',
        'رقم التأشيرة',
        'الجنسية',
        'جهة الوصول',
        'نوع التأشيرة',
        'تاريخ ووقت التحديث'
      ];

      const colWidths1 = [
        { wch: 18 }, // رقم جواز العميل
        { wch: 15 }, // رقم الهاتف
        { wch: 25 }, // اسم العميل
        { wch: 15 }, // المهنة
        { wch: 35 }, // حالة المعاملة من منصة التأشيرات
        { wch: 25 }, // حالة الكرت الصحي
        { wch: 15 }, // رقم الطلب
        { wch: 15 }, // رقم التأشيرة
        { wch: 15 }, // الجنسية
        { wch: 15 }, // جهة الوصول
        { wch: 15 }, // نوع التأشيرة
        { wch: 40 }  // بيانات إضافية
      ];

      const colWidths2 = [
        { wch: 18 }, // رقم جواز العميل
        { wch: 25 }, // اسم العميل
        { wch: 15 }, // المهنة
        { wch: 15 }, // رقم الطلب
        { wch: 25 }, // الحالة السابقة
        { wch: 25 }, // الحالة الجديدة
        { wch: 22 }, // تاريخ ووقت التحديث
        { wch: 25 }  // الوقت المتبقي
      ];

      const colWidths3 = [
        { wch: 18 }, // رقم جواز العميل
        { wch: 15 }, // رقم الهاتف
        { wch: 25 }, // اسم العميل
        { wch: 15 }, // المهنة
        { wch: 35 }, // حالة المعاملة من منصة التأشيرات
        { wch: 25 }, // حالة الكرت الصحي
        { wch: 15 }, // رقم الطلب
        { wch: 15 }, // رقم التأشيرة
        { wch: 15 }, // الجنسية
        { wch: 15 }, // جهة الوصول
        { wch: 15 }, // نوع التأشيرة
        { wch: 22 }  // تاريخ ووقت التحديث
      ];

      const createWorkbookSheet = (data: any[], headers: string[], widths: { wch: number }[]) => {
        let ws;
        if (data.length === 0) {
          ws = XLSX.utils.aoa_to_sheet([headers]);
        } else {
          ws = XLSX.utils.json_to_sheet(data, { header: headers });
        }
        ws['!cols'] = widths;
        return ws;
      };

      const wb = XLSX.utils.book_new();
      
      const ws1 = createWorkbookSheet(s1Data, s1Headers, colWidths1);
      XLSX.utils.book_append_sheet(wb, ws1, "كافة المعاملات");
      
      const ws2 = createWorkbookSheet(s2Data, s2Headers, colWidths2);
      XLSX.utils.book_append_sheet(wb, ws2, "تحديثات الـ 24 ساعة");
      
      const ws3 = createWorkbookSheet(s3Data, s3Headers, colWidths3);
      XLSX.utils.book_append_sheet(wb, ws3, "المعاملات الجاهزة");
      
      XLSX.writeFile(wb, "مزامنة_مكتب_انجاز_الموحدة.xlsx");
      addSyncLog("تم بنجاح ربط وسحب بيانات التنزيل، وتوليد ملف Excel الموحد بثلاث صفحات مرتبة بأعمدة الإدخال أولاً.");
    } catch (err: any) {
      console.error(err);
      addSyncLog(`[خطأ] فشل تصدير ملف excel: ${err.message}`);
    }
  };

  const executeSynchronization = async (isManualSimulated: boolean = false) => {
    setSyncStatus('syncing');
    addSyncLog("بدء تشغيل عملية الفحص الحية والمزامنة عبر Google Sheets API...");
    
    await new Promise(resolve => setTimeout(resolve, 1500));

    try {
      let incomingEntries: any[] = [];
      const defaultNationality = 'اليمن';
      const defaultArrival = 'عدن';
      const defaultVisaType = 'عمل';
      
      const hasApiConfig = googleSpreadsheetId.trim() && googleApiKey.trim();

      if (!isManualSimulated && hasApiConfig) {
        addSyncLog(`جاري الاتصال بـ Google Sheets API لجدول البيانات: ${googleSpreadsheetId}...`);
        
        try {
          const url = `https://sheets.googleapis.com/v4/spreadsheets/${googleSpreadsheetId}/values/${encodeURIComponent(googleSheetName)}?key=${googleApiKey}`;
          const response = await fetch(url);
          if (response.ok) {
            const resultData = await response.json();
            const parsed = resultData.values as any[][];
            
            if (parsed && parsed.length > 0) {
              // Analyze if there's a header row
              const firstRowStr = (parsed[0] || []).map(cell => String(cell || '').toLowerCase());
              const hasHeader = firstRowStr.some(cell => cell.includes('passport') || cell.includes('جواز') || cell.includes('هاتف') || cell.includes('phone') || cell.includes('تلفون') || cell.includes('اسم') || cell.includes('الاسم'));
              const headerIndex = hasHeader ? 1 : 0;

              // Track columns indices
              let passportIdx = 0;
              let phoneIdx = -1;
              let nameIdx = -1;
              let appIdx = -1;
              let profIdx = -1;
              let statusIdx = -1;
              let visaIdx = -1;
              let nationalityIdx = -1;
              let arrivalIdx = -1;
              let visaTypeIdx = -1;

              if (hasHeader) {
                parsed[0].forEach((h, idx) => {
                  const headStr = String(h || '').trim();
                  if (headStr.includes('جواز') || headStr.toLowerCase().includes('passport')) {
                    passportIdx = idx;
                  } else if (headStr.includes('هاتف') || headStr.includes('تلفون') || headStr.includes('جوال') || headStr.toLowerCase().includes('phone') || headStr.toLowerCase().includes('mobile')) {
                    phoneIdx = idx;
                  } else if (headStr.includes('اسم') || headStr.toLowerCase().includes('name')) {
                    nameIdx = idx;
                  } else if (headStr.includes('طلب') || headStr.toLowerCase().includes('app') || headStr.toLowerCase().includes('number')) {
                    appIdx = idx;
                  } else if (headStr.includes('مهنة') || headStr.includes('المهنة') || headStr.toLowerCase().includes('profession') || headStr.toLowerCase().includes('job')) {
                    profIdx = idx;
                  } else if (headStr.includes('حالة') || headStr.toLowerCase().includes('status')) {
                    statusIdx = idx;
                  } else if (headStr.includes('تأشيرة') || headStr.toLowerCase().includes('visa') || headStr.includes('تاشيرة')) {
                    visaIdx = idx;
                  } else if (headStr.includes('جنسية') || headStr.toLowerCase().includes('nationality')) {
                    nationalityIdx = idx;
                  } else if (headStr.includes('وصول') || headStr.toLowerCase().includes('arrival')) {
                    arrivalIdx = idx;
                  } else if (headStr.includes('نوع') || headStr.toLowerCase().includes('type')) {
                    visaTypeIdx = idx;
                  }
                });
              }

              const sampleRow = parsed[headerIndex] || [];
              if (phoneIdx === -1) {
                if (sampleRow.length <= 3) {
                  phoneIdx = 1;
                } else {
                  phoneIdx = 8;
                }
              }

              incomingEntries = parsed.slice(headerIndex).map((row: any[]) => {
                if (!row[passportIdx]) return null;

                const pNum = String(row[passportIdx] || '').trim().toUpperCase();
                const phoneVal = phoneIdx !== -1 && row[phoneIdx] ? String(row[phoneIdx] || '').trim() : '';
                const appNo = appIdx !== -1 ? String(row[appIdx] || '').trim() : '';
                const nameVal = nameIdx !== -1 ? String(row[nameIdx] || '') : '';
                const profVal = profIdx !== -1 ? String(row[profIdx] || '') : '';
                const natVal = nationalityIdx !== -1 ? String(row[nationalityIdx] || defaultNationality) : defaultNationality;
                const arrVal = arrivalIdx !== -1 ? String(row[arrivalIdx] || defaultArrival) : defaultArrival;
                const vtVal = visaTypeIdx !== -1 ? String(row[visaTypeIdx] || defaultVisaType) : defaultVisaType;
                const sTxt = statusIdx !== -1 ? String(row[statusIdx] || 'قيد الانتظار (Idle)') : 'قيد الانتظار (Idle)';
                const vNum = visaIdx !== -1 ? String(row[visaIdx] || '---') : '---';

                return {
                  passportNumber: pNum,
                  applicationNumber: appNo,
                  applicantName: nameVal,
                  nationality: natVal,
                  arrivalPoint: arrVal,
                  visaType: vtVal,
                  phoneNumber: phoneVal,
                  profession: profVal,
                  statusText: sTxt,
                  visaNumber: vNum,
                  applicantData: ''
                };
              }).filter(Boolean);
              
              addSyncLog(`تم بنجاح سحب وجلب ${incomingEntries.length} سجل من ورقة Google Sheets عبر الـ API!`);
            } else {
              addSyncLog("تنبيه: لم يتم العثور على أي بيانات أو صفوف في ورقة العمل المحددة.");
              isManualSimulated = true;
            }
          } else {
            const errDetails = await response.text();
            throw new Error(`جوجل API أرجع حالة خطأ ${response.status}: ${errDetails}`);
          }
        } catch (fetchErr: any) {
          addSyncLog(`[فشل الاتصال الفعلي]: ${fetchErr.message}. يرجى التثبت من صحة الـ API Key و الـ Spreadsheet ID للاتصال الحقيقي.`);
          isManualSimulated = true;
        }
      } else {
        isManualSimulated = true;
      }

      if (isManualSimulated) {
        addSyncLog("⚠️ تم تعطيل وضع المحاكاة التلقائية تماماً لضمان حقيقية ومدى صحة البيانات 100% بنسبة خالية من التخمين والمصادفات.");
        addSyncLog("💡 الرجاء تهيئة إعدادات ربط Google Sheets أو استخدام خاصية (الاستيراد اليدوي عبر لصق كشوفات الإكسل) لتغذية كشوفاتك برؤية معتمدة.");
        setSyncStatus('success');
        return;
      }

      let updatedCount = 0;
      let addedToUpdatesScreen = 0;
      let movedToReadyScreen = 0;

      const newEntries = [...entries];
      const newRecentUpdates = [...recentUpdates];
      const newReadyVisas = [...readyVisas];

      for (const incoming of incomingEntries) {
        const matchIdx = newEntries.findIndex(e => e.passportNumber.toUpperCase() === incoming.passportNumber.toUpperCase());
        
        if (matchIdx !== -1) {
          const existing = newEntries[matchIdx];
          const isChanged = 
            existing.statusText !== incoming.statusText ||
            existing.visaNumber !== incoming.visaNumber ||
            existing.applicantData !== incoming.applicantData;

          if (isChanged) {
            newEntries[matchIdx] = {
              ...existing,
              statusText: incoming.statusText,
              visaNumber: incoming.visaNumber,
              applicantData: incoming.applicantData,
              status: incoming.statusText.includes('Issued') || incoming.statusText.includes('اصدار') ? 'Found' : 'Checking',
              lastUpdate: new Date().toLocaleTimeString('ar-YE', { hour: '2-digit', minute: '2-digit' })
            };
            updatedCount++;

            const existsInRecent = newRecentUpdates.some(item => item.passportNumber === incoming.passportNumber);
            if (!existsInRecent) {
              newRecentUpdates.unshift({
                id: existing.id || Math.random().toString(36).substr(2, 9),
                passportNumber: incoming.passportNumber,
                applicantName: incoming.applicantName || existing.applicantName || 'غير مسجل',
                profession: incoming.profession || existing.profession || '',
                applicationNumber: incoming.applicationNumber || existing.applicationNumber,
                oldStatusText: existing.statusText || '---',
                newStatusText: incoming.statusText,
                copiedAt: new Date().toISOString()
              });
              addedToUpdatesScreen++;
            }
          }
        } else {
          const newId = 'ex-' + Math.random().toString(36).substr(2, 9);
          newEntries.unshift({
            id: newId,
            passportNumber: incoming.passportNumber,
            applicationNumber: incoming.applicationNumber,
            applicantName: incoming.applicantName || 'مشترك جديد من جدول البيانات',
            profession: incoming.profession || '',
            nationality: incoming.nationality || defaultNationality,
            arrivalPoint: incoming.arrivalPoint || defaultArrival,
            visaType: incoming.visaType || defaultVisaType,
            status: incoming.statusText?.includes('Issued') || incoming.statusText?.includes('اصدار') ? 'Found' : 'Idle',
            statusText: incoming.statusText || 'قيد الانتظار (Idle)',
            visaNumber: incoming.visaNumber || '---',
            phoneNumber: incoming.phoneNumber || '',
            applicantData: incoming.applicantData || '',
            lastUpdate: new Date().toLocaleTimeString('ar-YE', { hour: '2-digit', minute: '2-digit' })
          });
          updatedCount++;
        }

        const isReady = incoming.statusText?.includes('Issued') || incoming.statusText?.includes('تم اصدار') || incoming.statusText?.includes('مؤشرة') || incoming.statusText?.includes('جاهز');
        if (isReady) {
          const alreadyInReady = newReadyVisas.some(r => r.passportNumber === incoming.passportNumber);
          if (!alreadyInReady) {
            newReadyVisas.unshift({
              id: incoming.id || Math.random().toString(36).substr(2, 9),
              passportNumber: incoming.passportNumber,
              applicationNumber: incoming.applicationNumber,
              applicantName: incoming.applicantName || 'غير متوفر',
              profession: incoming.profession || '',
              nationality: incoming.nationality || defaultNationality,
              arrivalPoint: incoming.arrivalPoint || defaultArrival,
              visaType: incoming.visaType || defaultVisaType,
              status: 'Found',
              statusText: incoming.statusText,
              visaNumber: incoming.visaNumber || '---',
              phoneNumber: incoming.phoneNumber,
              applicantData: incoming.applicantData
            });
            movedToReadyScreen++;
            addSyncLog(`[النقل التلقائي للورقة 3] تم نقل الجواز الجاهز رقم ${incoming.passportNumber} المنجز والمؤشر تلقائياً.`);
          }
        }
      }

      setEntries(newEntries);
      setRecentUpdates(newRecentUpdates);
      setReadyVisas(newReadyVisas);
      setLastSyncTime(new Date().toLocaleString('ar-YE'));
      setSyncStatus('success');
      
      addSyncLog(`اكتملت المزامنة بنجاح: تم معالجة وتدقيق ${updatedCount} معاملة بنجاح.`);
    } catch (err: any) {
      console.error(err);
      setSyncStatus('error');
      addSyncLog(`[فشل المزامنة] خطأ في بنية البيانات: ${err.message || 'خطأ غير مسمى'}`);
    }
  };

  const t = {
    ar: {
      title: lang === 'ar' ? siteIdentity.officeNameAr : siteIdentity.officeNameEn,
      subtitle: lang === 'ar' ? siteIdentity.officeDescriptionAr : siteIdentity.officeDescriptionEn,
      add: 'طلب جديد',
      bulk: 'استيراد خوارزمي',
      checkAll: 'تحديث الكل',
      export: 'تصدير Excel',
      importExcel: 'استيراد Excel',
      passport: 'رقم الجواز',
      appNo: 'رقم الطلب',
      name: 'اسم العميل',
      nationality: 'الجنسية',
      arrival: 'جهة القدوم',
      visaType: 'نوع الإقامة',
      status: 'حالة الطلب',
      visaNo: 'رقم التأشيرة',
      lastUpdate: 'التحديث',
      actions: 'خيارات',
      noData: 'ابدأ بإضافة بيانات الجوازات والطلبات لتتبع حالتها بشكل آلي.',
      parseTitle: 'خوارزمية الاستيراد الذكي',
      parseDesc: 'يتم تحليل النص بخوارزميات دقيقة لاستخراج أرقام الجوازات والطلبات فوراً.',
      close: 'إلغاء',
      process: 'بدء التحليل الرقمي',
      checking: 'جاري التحقق...',
      searchPlaceholder: 'البحث عن جواز، رقم طلب، أو اسم...',
      stats: 'ملخص العمليات',
      tab_dashboard: 'الرئيسية (لوحة التحكم)',
      tab_add: 'إضافة البيانات',
      save: 'حفظ البيانات',
      cancel: 'إلغاء الأمر',
      addModalTitle: 'إضافة طلب جديد للتحقق',
      addModalDesc: 'أدخل بيانات الجواز والطلب بشكل دقيق لضمان نتائج صحيحة.',
      tab_ready: 'التأشيرات الجاهزة',
      tab_docs: 'أرقام المستندات',
      tab_data: 'بيانات ظاهرة',
      tab_pending: 'لم يصدر شيء',
      tab_all: 'المعاملات',
      tab_identity: 'هوية الموقع',
      tab_security: 'الأمان والمستخدمين',
      tab_excel_sync: 'ربط البيانات',
      tab_customers: 'إدارة بيانات العملاء',
      tab_files: 'ملفات العملاء',
      tab_auto_check: 'التشييك التلقائي',
      pending: 'قيد الانتظار',
      issued: 'تم الإصدار',
      total: 'الإجمالي',
      adminPortal: 'لوحة الإدارة',
      publicPortal: 'بوابة الاستعلام',
      phone: 'رقم الهاتف أو الرقم المخصص',
      phonePlaceholder: `مثال: ${siteIdentity.supportPhone}`,
      searchInquiry: siteIdentity.searchInquiryAr,
      searchResultTitle: 'تفاصيل حالة معاملتكم',
      whatsappFollowup: 'متابعة المعاملة عبر الواتساب',
      officeName: siteIdentity.officeNameAr,
      officeDescription: siteIdentity.officeDescriptionAr,
      noSearchYet: 'يرجى إدخال رقم جواز السفر ورقم الهاتف لبدء الاستعلام الفوري',
      notFoundTitle: 'عذراً، لم نجد معاملة مطابقة',
      notFoundDesc: 'تأكد من كتابة رقم الجواز ورقم الهاتف المسجل لدى المكتب بشكل صحيح. إذا استمر هذا التنبيه، يُرجى التواصل معنا عبر وسائل الدعم بمكتبنا لربط الهاتف بجوازك فورا.',
      backToPublic: 'العودة لواجهة الاستعلام',
      adminPassTitle: 'دخول لوحة التحكم',
      adminPassDesc: 'يرجى إدخال اسم المستخدم وكلمة المرور المسجلة بمؤسسة إنجاز للتحكم بلوحة البيانات والطلبات.',
      adminLogin: 'تسجيل الدخول للتحكم',
      incorrectPass: 'اسم المستخدم أو كلمة المرور غير صحيحة، يرجى إعادة المحاولة',
      usernameLabel: 'اسم المستخدم',
      passwordLabel: 'كلمة المرور',
      pinLabel: 'ملاحظة: بيانات الدخول الافتراضية؛ اسم المستخدم: admin | كلمة المرور: 1234'
    },
    en: {
      title: siteIdentity.officeNameEn,
      subtitle: siteIdentity.officeDescriptionEn,
      add: 'New Entry',
      bulk: 'Algorithmic Import',
      checkAll: 'Update All',
      export: 'Export Excel',
      importExcel: 'Import Excel',
      passport: 'Passport #',
      appNo: 'App #',
      name: 'Applicant Name',
      nationality: 'Nationality',
      arrival: 'Arrival',
      visaType: 'Type',
      status: 'Status',
      visaNo: 'Visa No',
      lastUpdate: 'Last Update',
      actions: 'Actions',
      noData: 'Start by adding passport and application data to track status automatically.',
      parseTitle: 'High-Precision Import Algorithm',
      parseDesc: 'Data is extracted using sophisticated pattern matching algorithms for maximum accuracy.',
      close: 'Cancel',
      process: 'Run Analysis',
      checking: 'Checking...',
      searchPlaceholder: 'Search passport, application, or name...',
      stats: 'Operation Stats',
      tab_dashboard: 'Home Dashboard',
      tab_add: 'Add Data',
      save: 'Save Entry',
      cancel: 'Cancel',
      addModalTitle: 'Add New Tracking Request',
      addModalDesc: 'Enter passport and application details precisely for accurate results.',
      tab_ready: 'Ready Visas',
      tab_docs: 'Doc Numbers',
      tab_data: 'Visible Data',
      tab_pending: 'Not Issued',
      tab_all: 'Transactions',
      tab_identity: 'Website Identity',
      tab_security: 'Security & Users',
      tab_excel_sync: 'Data Sync',
      tab_customers: 'Customer Management',
      tab_files: 'Client Files',
      tab_auto_check: 'Auto Check',
      pending: 'Pending',
      issued: 'Issued',
      total: 'Total',
      adminPortal: 'Admin Control',
      publicPortal: 'Inquiry Portal',
      phone: 'Phone / Custom Number',
      phonePlaceholder: `e.g. ${siteIdentity.supportPhone}`,
      searchInquiry: siteIdentity.searchInquiryEn,
      searchResultTitle: 'Application Status Details',
      whatsappFollowup: 'Follow up on WhatsApp',
      officeName: siteIdentity.officeNameEn,
      officeDescription: siteIdentity.officeDescriptionEn,
      noSearchYet: 'Enter your passport number and mobile phone to initiate transaction inquiry',
      notFoundTitle: 'No Matching Application Found',
      notFoundDesc: 'Please verify the passport number and registered phone number. If the issue persists, please consult with the office administration to bind your phone correctly.',
      backToPublic: 'Back to Inquiry Portal',
      adminPassTitle: 'Admin Portal Login',
      adminPassDesc: 'Please enter your authorized username and password/code to manage visa data tables.',
      adminLogin: 'Authorized Login',
      incorrectPass: 'Invalid Username or Password. Please try again.',
      usernameLabel: 'Username',
      passwordLabel: 'Password',
      pinLabel: 'Defaut Login: Username: admin | Password: 1234'
    }
  }[lang];

  const stats = useMemo(() => ({
    total: entries.filter(e => !e.archived).length,
    issued: entries.filter(e => !e.archived && (((e.statusText?.includes('تم اصدار') || e.statusText?.includes('Issued')) && entryHasPrint(e)) || entryHasVisaAndImage(e))).length,
    docs: entries.filter(e => !e.archived && e.visaNumber && e.visaNumber !== 'N/A' && e.visaNumber !== '---' && e.visaNumber.trim() !== '').length,
    dataOnly: entries.filter(e => !e.archived && e.applicantData && (!e.statusText || (!e.statusText.includes('Issued') && !e.statusText.includes('تم اصدار')))).length,
    noStatus: entries.filter(e => !e.archived && (!e.statusText || (!e.statusText.includes('Issued') && !e.statusText.includes('تم اصدار') && (!e.visaNumber || e.visaNumber === 'N/A' || e.visaNumber === '---')))).length,
    pending: entries.filter(e => !e.archived && (e.status === 'Checking' || e.status === 'Idle')).length,
  }), [entries]);

  const isToday = (entry: VisaEntry) => {
    if (entry.updatedAtUnix) {
      const d = new Date(entry.updatedAtUnix);
      const today = new Date();
      return d.getDate() === today.getDate() &&
             d.getMonth() === today.getMonth() &&
             d.getFullYear() === today.getFullYear();
    }
    if (entry.lastUpdate === '10:30 AM' || entry.lastUpdate === '02:15 PM') return true;
    return false;
  };

  const deptStats = useMemo(() => {
    // Labels of options that are "isolated" — their entries are excluded from regular stats
    const isolatedLabels = txStatusOptions
      .filter(opt => opt.hasStatsCard && opt.statsIsolate && opt.type === 'موقفة')
      .map(opt => opt.label);
    const notIsolated = (e: VisaEntry) => !isolatedLabels.includes(e.txStatus || '');

    const base = entries.filter(e => !e.archived);
    const total = base.filter(notIsolated).length;
    const ready = base.filter(e => notIsolated(e) && (((e.statusText?.includes('تم اصدار') || e.statusText?.includes('Issued')) && entryHasPrint(e)) || entryHasVisaAndImage(e))).length;
    const today = base.filter(e => notIsolated(e) && isToday(e)).length;
    const withPrint = base.filter(e => notIsolated(e) && entryHasPrint(e) && !entryHasVisaAndImage(e)).length;
    const withoutPrint = base.filter(e => notIsolated(e) && !entryHasPrint(e) && !entryHasVisaAndImage(e)).length;

    const regularCards = [
      {
        id: 'all',
        name: lang === 'ar' ? 'كافة المعاملات' : 'All Transactions',
        desc: lang === 'ar' ? 'كافة سجلات المعاملات والعملاء المضافة' : 'All transaction and customer entries',
        count: total,
        percentage: 100,
        color: 'bg-slate-50 text-slate-600 hover:bg-slate-100',
        borderColor: 'border-r-slate-500',
        icon: <Globe size={24} />,
        filterType: 'all'
      },
      {
        id: 'today',
        name: lang === 'ar' ? 'تحديثات اليوم' : "Today's Updates",
        desc: lang === 'ar' ? 'المعاملات المعدلة والمحدثة اليوم' : 'Transactions updated or modified today',
        count: today,
        percentage: total ? Math.round((today / total) * 100) : 0,
        color: 'bg-indigo-50 text-indigo-600 hover:bg-indigo-100/50',
        borderColor: 'border-r-indigo-500',
        icon: <RefreshCw size={24} />,
        filterType: 'today'
      },
      {
        id: 'with_print',
        name: lang === 'ar' ? 'المعاملات لها طلب ( برنت )' : 'Transactions with Print',
        desc: lang === 'ar' ? 'المعاملات التي تم ربط وسحب بيانات البرنت لها' : 'Entries with configured print/application data',
        count: withPrint,
        percentage: total ? Math.round((withPrint / total) * 100) : 0,
        color: 'bg-blue-50 text-blue-600 hover:bg-blue-100/50',
        borderColor: 'border-r-blue-500',
        icon: <FileSpreadsheet size={24} />,
        filterType: 'with_print'
      },
      {
        id: 'without_print',
        name: lang === 'ar' ? 'المعاملات التي لم يصدر لها برنت' : 'Transactions without Print',
        desc: lang === 'ar' ? 'المعاملات التي لم يتم ربط أو سحب بيانات البرنت لها بعد' : 'Entries without configured print/application data yet',
        count: withoutPrint,
        percentage: total ? Math.round((withoutPrint / total) * 100) : 0,
        color: 'bg-amber-50 text-amber-600 hover:bg-amber-100/50',
        borderColor: 'border-r-amber-500',
        icon: <FileX size={24} />,
        filterType: 'without_print'
      },
      {
        id: 'issued',
        name: lang === 'ar' ? 'المعاملات الجاهزة ( المؤشرة )' : 'Ready / Issued Transactions',
        desc: lang === 'ar' ? 'المعاملات المنجزة الجاهزة للتسليم للعملاء' : 'Completed and fully issued transactions',
        count: ready,
        percentage: total ? Math.round((ready / total) * 100) : 0,
        color: 'bg-emerald-50 text-emerald-600 hover:bg-emerald-100/50',
        borderColor: 'border-r-emerald-500',
        icon: <CheckCircle2 size={24} />,
        filterType: 'issued'
      }
    ];

    // Custom tx-status stats cards
    const customCards = txStatusOptions
      .filter(opt => opt.hasStatsCard)
      .map(opt => {
        const count = base.filter(e => e.txStatus === opt.label).length;
        const allTotal = base.length;
        return {
          id: `tx_status:${opt.label}`,
          name: opt.label,
          desc: `${count} معاملة${opt.statsIsolate ? ' — معزولة عن الأقسام الأخرى' : ''}`,
          count,
          percentage: allTotal ? Math.round((count / allTotal) * 100) : 0,
          color: opt.type === 'موقفة' ? 'bg-red-50 text-red-600 hover:bg-red-100/50' : 'bg-blue-50 text-blue-600 hover:bg-blue-100/50',
          borderColor: opt.type === 'موقفة' ? 'border-r-red-500' : 'border-r-blue-500',
          icon: <Tag size={24} />,
          filterType: `tx_status:${opt.label}`,
          isCustom: true,
          isolate: !!opt.statsIsolate,
          txType: opt.type,
        };
      });

    return [...regularCards, ...customCards];
  }, [entries, lang, txStatusOptions]);

  const filteredByTab = useMemo(() => {
    const matchesQuery = (e: VisaEntry) => {
      const q = filterQuery.toLowerCase();
      return (
        e.passportNumber.toLowerCase().includes(q) ||
        e.applicationNumber.toLowerCase().includes(q) ||
        (e.documentNumber || '').toLowerCase().includes(q) ||
        (e.customId || '').toLowerCase().includes(q) ||
        (e.phoneNumber || '').toLowerCase().includes(q) ||
        (e.applicantName || '').toLowerCase().includes(q)
      );
    };

    let base = entries.filter(e => !e.archived && matchesQuery(e));

    if (selectedVisaType) {
      base = base.filter(e => e.visaType === selectedVisaType);
    }

    if (selectedOfficeFilter !== null) {
      if (selectedOfficeFilter === 'general') {
        base = base.filter(e => !e.officeId || e.officeId === 'general');
      } else {
        base = base.filter(e => e.officeId === selectedOfficeFilter);
      }
    }

    let result: VisaEntry[];
    switch(activeTab) {
      case 'ready': result = base.filter(e => ((e.statusText?.includes('Issued') || e.statusText?.includes('تم اصدار')) && entryHasPrint(e)) || entryHasVisaAndImage(e)); break;
      case 'docs': result = base.filter(e => e.visaNumber && e.visaNumber !== 'N/A' && e.visaNumber !== '---' && e.visaNumber.trim() !== ''); break;
      case 'data_only': result = base.filter(e => e.applicantData && (!e.statusText || (!e.statusText.includes('Issued') && !e.statusText.includes('تم اصدار')))); break;
      case 'pending': result = base.filter(e => !e.statusText || (!e.statusText.includes('Issued') && !e.statusText.includes('تم اصدار') && (!e.visaNumber || e.visaNumber === 'N/A' || e.visaNumber === '---' || e.visaNumber.trim() === ''))); break;
      case 'all': {
        if (allTabFilter === 'today') { result = base.filter(isToday); break; }
        if (allTabFilter === 'with_print') {
          let r = base.filter(e => entryHasPrint(e) && !entryHasVisaAndImage(e));
          if (docNumberFilter === 'with_doc') r = r.filter(e => e.customId && e.customId !== '---' && e.customId.trim() !== '');
          else if (docNumberFilter === 'without_doc') r = r.filter(e => !e.customId || e.customId === '---' || e.customId.trim() === '');
          result = r; break;
        }
        if (allTabFilter === 'without_print') { result = base.filter(e => !entryHasPrint(e) && !entryHasVisaAndImage(e)); break; }
        if (allTabFilter === 'issued') { result = base.filter(e => ((e.statusText?.includes('تم اصدار') || e.statusText?.includes('Issued')) && entryHasPrint(e)) || entryHasVisaAndImage(e)); break; }
        if (allTabFilter.startsWith('tx_status:')) {
          const statusLabel = allTabFilter.slice('tx_status:'.length);
          result = base.filter(e => (e.txStatus || '') === statusLabel); break;
        }
        result = base; break;
      }
      default: result = base;
    }

    if (sortOrder === 'newest') return [...result].sort((a, b) => (b.updatedAtUnix || 0) - (a.updatedAtUnix || 0));
    if (sortOrder === 'oldest') return [...result].sort((a, b) => (a.updatedAtUnix || 0) - (b.updatedAtUnix || 0));
    return result;
  }, [entries, activeTab, filterQuery, allTabFilter, selectedVisaType, selectedOfficeFilter, sortOrder, docNumberFilter]);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (!hasPermission('add_clients')) {
      showAlert(lang === 'ar' ? 'عذراً، ليس لديك صلاحية "اضافه عملاء" لاستيراد ملفات البيانات!' : 'Sorry, you do not have permission to upload/add new clients!', 'warning');
      return;
    }
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

      if (data.length === 0) return;

      // Analyze if there's a header row
      const firstRowStr = (data[0] || []).map(cell => String(cell || '').toLowerCase());
      const hasHeader = firstRowStr.some(cell => cell.includes('passport') || cell.includes('جواز') || cell.includes('هاتف') || cell.includes('phone') || cell.includes('تلفون'));
      const startIndex = hasHeader ? 1 : 0;

      // Track columns indices
      let passportIdx = 0;
      let phoneIdx = -1;
      let nameIdx = -1;
      let appIdx = -1;
      let profIdx = -1;
      let statusIdx = -1;
      let visaIdx = -1;
      let nationalityIdx = -1;
      let arrivalIdx = -1;
      let visaTypeIdx = -1;

      if (hasHeader) {
        data[0].forEach((h, idx) => {
          const headStr = String(h || '').trim();
          if (headStr.includes('جواز') || headStr.toLowerCase().includes('passport')) {
            passportIdx = idx;
          } else if (headStr.includes('هاتف') || headStr.includes('تلفون') || headStr.includes('جوال') || headStr.toLowerCase().includes('phone') || headStr.toLowerCase().includes('mobile')) {
            phoneIdx = idx;
          } else if (headStr.includes('اسم') || headStr.toLowerCase().includes('name')) {
            nameIdx = idx;
          } else if (headStr.includes('طلب') || headStr.toLowerCase().includes('app') || headStr.toLowerCase().includes('number')) {
            appIdx = idx;
          } else if (headStr.includes('مهنة') || headStr.includes('المهنة') || headStr.toLowerCase().includes('profession') || headStr.toLowerCase().includes('job')) {
            profIdx = idx;
          } else if (headStr.includes('حالة') || headStr.toLowerCase().includes('status')) {
            statusIdx = idx;
          } else if (headStr.includes('تأشيرة') || headStr.toLowerCase().includes('visa') || headStr.includes('تاشيرة')) {
            visaIdx = idx;
          } else if (headStr.includes('جنسية') || headStr.toLowerCase().includes('nationality')) {
            nationalityIdx = idx;
          } else if (headStr.includes('وصول') || headStr.toLowerCase().includes('arrival')) {
            arrivalIdx = idx;
          } else if (headStr.includes('نوع') || headStr.toLowerCase().includes('type')) {
            visaTypeIdx = idx;
          }
        });
      }

      const sampleRow = data[startIndex] || [];
      if (phoneIdx === -1) {
        // Fallback for simple 2-column layout mapping: column 0 is passport, column 1 is phone
        if (sampleRow.length <= 3) {
          phoneIdx = 1;
        } else {
          phoneIdx = 6; // Legacy column layout index
        }
      }

      const newEntries: VisaEntry[] = data.slice(startIndex).map((row: any[]) => {
        if (!row[passportIdx]) return null;

        const pNum = String(row[passportIdx] || '').trim().toUpperCase();
        const phoneVal = phoneIdx !== -1 && row[phoneIdx] ? String(row[phoneIdx] || '').trim() : '';
        const appNo = appIdx !== -1 ? String(row[appIdx] || '').trim() : '';
        const nameVal = nameIdx !== -1 ? String(row[nameIdx] || '') : '';
        const profVal = profIdx !== -1 ? String(row[profIdx] || '') : '';
        const natVal = nationalityIdx !== -1 ? String(row[nationalityIdx] || defaultNationality) : defaultNationality;
        const arrVal = arrivalIdx !== -1 ? String(row[arrivalIdx] || defaultArrival) : defaultArrival;
        const vtVal = visaTypeIdx !== -1 ? String(row[visaTypeIdx] || defaultVisaType) : defaultVisaType;
        const sTxt = statusIdx !== -1 ? String(row[statusIdx] || 'قيد الانتظار (Idle)') : 'قيد الانتظار (Idle)';
        const vNum = visaIdx !== -1 ? String(row[visaIdx] || '---') : '---';

        return {
          id: Math.random().toString(36).substr(2, 9),
          passportNumber: pNum,
          applicationNumber: appNo,
          applicantName: nameVal,
          nationality: natVal,
          arrivalPoint: arrVal,
          visaType: vtVal,
          phoneNumber: phoneVal,
          profession: profVal,
          status: sTxt.includes('Issued') || sTxt.includes('اصدار') || sTxt.includes('مؤشرة') ? 'Found' : 'Idle',
          statusText: sTxt,
          visaNumber: vNum,
          lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true }),
          updatedAtUnix: Date.now()
        };
      }).filter(Boolean) as VisaEntry[];

      setEntries(prev => [...newEntries, ...prev]);
      if (fileInputRef.current) fileInputRef.current.value = '';
    };
    reader.readAsBinaryString(file);
  };

  const addEntry = () => {
    setAddForm({
      name: '',
      passport: '',
      appNo: '',
      nationality: defaultNationality,
      arrival: defaultArrival,
      visaType: defaultVisaType,
      phone: '',
      profession: '',
      officeId: 'general'
    });
    setIsAddOpen(true);
  };

  const saveEntry = () => {
    if (!hasPermission('add_clients')) {
      showAlert(lang === 'ar' ? 'عذراً، ليس لديك صلاحية "اضافه عملاء" لتسجيل معاملة جديدة!' : 'Sorry, you do not have permission to add new clients!', 'warning');
      return;
    }
    if (!addForm.passport) return;
    const newEntry: VisaEntry = {
      id: Math.random().toString(36).substr(2, 9),
      applicantName: addForm.name,
      passportNumber: addForm.passport.trim().toUpperCase(),
      applicationNumber: addForm.appNo.trim(),
      nationality: addForm.nationality,
      arrivalPoint: addForm.arrival,
      visaType: addForm.visaType,
      phoneNumber: addForm.phone.trim(),
      profession: addForm.profession.trim(),
      officeId: addForm.officeId || 'general',
      status: 'Idle',
      statusText: 'قيد الانتظار (Idle)',
      visaNumber: '---',
      lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true })
    };
    setEntries([newEntry, ...entries]);
    setIsAddOpen(false);
  };

  const handlePublicSearch = async () => {
    if (!publicPassport.trim() || !publicPhone.trim()) return;
    setIsPublicSearching(true);
    setPublicHasSearched(false);
    setPublicSearchResult(null);

    await new Promise(resolve => setTimeout(resolve, 1100));

    // Cleanup phone formatting for flexible checks
    const cleanedSearchPhone = publicPhone.trim().replace(/[\s\-\(\)\+]/g, '');

    const match = entries.find(e => {
      const dbPassport = e.passportNumber.trim().toUpperCase();
      const dbPhone = (e.phoneNumber || '').trim().replace(/[\s\-\(\)\+]/g, '');
      
      const passportMatches = dbPassport === publicPassport.trim().toUpperCase();
      // Allow passing if search phone is contained or reversed to accommodate local numbering differences
      const phoneMatches = dbPhone.includes(cleanedSearchPhone) || cleanedSearchPhone.includes(dbPhone);
      
      return passportMatches && phoneMatches;
    });

    setIsPublicSearching(false);
    setPublicSearchResult(match || null);
    setPublicHasSearched(true);
  };

  const removeEntry = (id: string) => {
    if (!hasPermission('edit_clients')) {
      showAlert(lang === 'ar' ? 'عذراً، ليس لديك صلاحية "تعديل بيانات العملاء" لحذف هذا السجل!' : 'Sorry, you do not have permission to delete client records!', 'warning');
      return;
    }
    const updated = entries.filter(e => e.id !== id);
    setEntries(updated);
    fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
  };

  const updateEntryField = (id: string, field: keyof VisaEntry, value: string) => {
    if (!hasPermission('edit_clients')) return;
    setEntries(entries.map(e => e.id === id ? { ...e, [field]: value } : e));
  };

  const handleBulkParse = async () => {
    if (!hasPermission('add_clients')) {
      showAlert(lang === 'ar' ? 'عذراً، ليس لديك صلاحية "اضافه عملاء" للقيام بالاستيراد الخوارزمي المجمع!' : 'Sorry, you do not have permission to add new clients!', 'warning');
      return;
    }
    if (!bulkText.trim()) return;
    setParsing(true);
    
    // Simulate parse speed
    await new Promise(resolve => setTimeout(resolve, 650));

    try {
      // RegEx capture of structural blocks
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
          applicationNumber: appNos[i] || '',
          applicantName: names[i] || '',
          nationality: defaultNationality,
          arrivalPoint: defaultArrival,
          visaType: defaultVisaType,
          status: 'Idle',
          statusText: 'قيد الانتظار (Idle)',
          visaNumber: '---',
          lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true }),
          updatedAtUnix: Date.now()
        });
      }

      setEntries(prev => [...newEntries, ...prev]);
      setIsBulkOpen(false);
      setBulkText('');
    } catch (err) {
      console.error(err);
    } finally {
      setParsing(false);
    }
  };



  const parseMofaPageText = (text: string) => {
    if (!text || !text.trim()) return null;

    let applicantName = '';
    let applicationNumber = '';
    let visaNumber = '';
    let profession = '';
    let nationality = '';
    let arrivalPoint = '';
    let visaType = '';
    let statusText = '';

    // Advanced regex search
    const appNoMatch = text.match(/(?:رقم الطلب|طلب رقم|Application Number|Application No|App No)\s*[:：\-‐]?\s*(\d{7,15})/i) || text.match(/\b(70\d{8,11})\b/);
    if (appNoMatch) applicationNumber = appNoMatch[1];

    const nameMatch = text.match(/(?:اسم صاحب الطلب|الاسم|اسم صاحب الجواز|اسم الطالب|Name of Applicant|Applicant Name|Candidate Name|Full Name|Name)\s*[:：\-‐]\s*([^•\n\r|:]{3,80})/i) ||
                      text.match(/(?:اسم صاحب الطلب|الاسم|اسم صاحب الجواز|اسم الطالب|Name of Applicant|Applicant Name|Candidate Name|Full Name|Name)\s+([^•\n\r|:\d]{3,80})/i);
    if (nameMatch) applicantName = nameMatch[1].trim();

    const visaNoMatch = text.match(/(?:رقم التأشيرة|مستند تأشيرة|رقم مستند التأشيرة|Visa Number|Visa No)\s*[:：\-‐]\s*([A-Za-z0-9]{5,20})/i) || 
                        text.match(/\b(V\d{7,12})\b/) || 
                        text.match(/\b(40\d{8,11})\b/);
    if (visaNoMatch) visaNumber = visaNoMatch[1].trim();

    let documentNumber = '';
    const docNoMatch = text.match(/(?:رقم المستند|رقم الوثيقة|Document Number|Doc No|Document No)\s*[:：\-‐]?\s*([A-Za-z0-9]{4,25})/i);
    if (docNoMatch) documentNumber = docNoMatch[1].trim();

    const professionMatch = text.match(/(?:المهنة|الوظيفة|المسمى الوظيفي|العمل|Profession|Occupation|Job Title|Job)\s*[:：\-‐]\s*([^•\n\r|:]{3,60})/i) ||
                            text.match(/(?:المهنة|الوظيفة|المسمى الوظيفي|العمل|Profession|Occupation|Job Title|Job)\s+([^•\n\r|:\d]{3,60})/i);
    if (professionMatch) profession = professionMatch[1].trim();

    const nationalityMatch = text.match(/(?:الجنسية|جنسية|Nationality)\s*[:：\-‐]\s*([^•\n\r|:]{3,40})/i) ||
                             text.match(/(?:الجنسية|جنسية|Nationality)\s+([^•\n\r|:\d]{3,40})/i);
    if (nationalityMatch) nationality = nationalityMatch[1].trim();

    const arrivalMatch = text.match(/(?:جهة القدوم|منفذ القدوم|منفذ|Arrival Point|Port of Entry|Coming Port)\s*[:：\-‐]\s*([^•\n\r|:]{3,40})/i) ||
                         text.match(/(?:جهة القدوم|منفذ القدوم|منفذ|Arrival Point|Port of Entry|Coming Port)\s+([^•\n\r|:\d]{3,40})/i);
    if (arrivalMatch) arrivalPoint = arrivalMatch[1].trim();

    const typeMatch = text.match(/(?:نوع التأشيرة|نوع الخدمة|نوع الطلب|Visa Type|Service Type)\s*[:：\-‐]\s*([^•\n\r|:]{3,45})/i) ||
                      text.match(/(?:نوع التأشيرة|نوع الخدمة|نوع الطلب|Visa Type|Service Type)\s+([^•\n\r|:\d]{3,45})/i);
    if (typeMatch) visaType = typeMatch[1].trim();

    if (text.includes("تم اصدار التأشيرة") || text.includes("Visa Issued") || text.includes("Issued") || text.includes("تم اصدار") || visaNumber) {
      statusText = "تم اصدار التأشيرة (Issued)";
    } else if (text.includes("تحت الاجراء") || text.includes("Under Process") || text.includes("تحت الإجراء") || text.includes("In progress")) {
      statusText = "تحت الاجراء ومطابقة الجوازات (Under Process)";
    } else {
      statusText = "تحت الاجراء ومطابقة الجوازات (Under Process)";
    }

    return {
      applicantName,
      applicationNumber,
      visaNumber,
      documentNumber,
      profession,
      nationality,
      arrivalPoint,
      visaType,
      statusText
    };
  };

  // ─── Silent MOFA check (no modal state, for bulk) ───────────────────────────
  const silentMofaCheck = async (
    entry: VisaEntry, retryCount = 0
  ): Promise<{success: boolean; merged?: any; error?: string}> => {
    if (retryCount >= 3) return { success: false, error: 'فشل بعد 3 محاولات' };
    const initRes = await fetch('/api/mofa/init');
    if (!initRes.ok) return { success: false, error: 'فشل الاتصال بمنصة التأشيرات' };
    const initData = await initRes.json();
    if (!initData.autoSolvedCaptcha) return { success: false, error: 'تعذّر حل رمز التحقق تلقائياً' };
    const matchLabel = (list: {value:string;label:string}[], text: string) => {
      if (!text) return null;
      const t = text.trim();
      const exact = list.find(o => o.label === t);
      if (exact) return exact.value;
      const partial = list.find(o => o.label.includes(t) || t.includes(o.label));
      return partial ? partial.value : null;
    };
    const nats = initData.nationalities || [];
    const vts  = initData.visaTypes    || [];
    const embs = initData.embassies    || [];
    const nat = matchLabel(nats, entry.nationality  || '') || 'YEM';
    const vt  = matchLabel(vts,  entry.visaType     || '') || '1';
    const emb = matchLabel(embs, entry.arrivalPoint || '') || '302';
    const submitRes = await fetch('/api/mofa/submit', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId: initData.sessionId, passportNumber: entry.passportNumber, nationality: nat, visaType: vt, embassy: emb, captcha: initData.autoSolvedCaptcha, existingPrintUrl: entry.printImageUrl || '', applicantName: entry.applicantName || '' }),
    });
    const submitData = await submitRes.json();
    if (!submitRes.ok) return { success: false, error: submitData.error || 'فشل الاستعلام' };
    if (!submitData.success) {
      if (submitData.notFound || NO_PRINT_INDICATORS.some(ind => (submitData.error || '').includes(ind)))
        return { success: false, error: submitData.error || 'لا توجد بيانات في المنصة' };
      await new Promise(r => setTimeout(r, 1200));
      return silentMofaCheck(entry, retryCount + 1);
    }
    const d2 = submitData.data || {};
    const rawHtml = d2.rawHtml || '';
    const parsed2 = parseMofaPageText(rawHtml);
    const appDateFallback = rawHtml.match(/تاريخ الطلب[:\s]+(\d{4}\/\d{2}\/\d{2})/)?.[1] || '';
    return { success: true, merged: {
      applicantName:      d2.applicantName      || parsed2?.applicantName      || '',
      nameEnglish:        d2.nameEnglish        || '',
      applicationNumber:  d2.applicationNumber  || parsed2?.applicationNumber  || '',
      applicationDate:    d2.applicationDate    || appDateFallback             || '',
      embassy:            d2.embassy            || '',
      visaNumber:         d2.visaNumber         || parsed2?.visaNumber         || '',
      documentNumber:     d2.documentNumber     || parsed2?.documentNumber     || '',
      entriesCount:       d2.entriesCount       || '',
      requesterName:      d2.requesterName      || '',
      passportType:       d2.passportType       || '',
      passportExpiry:     d2.passportExpiry     || '',
      birthDate:          d2.birthDate          || '',
      birthPlace:         d2.birthPlace         || '',
      currentNationality: d2.currentNationality || parsed2?.nationality        || '',
      gender:             d2.gender             || '',
      profession:         d2.profession         || parsed2?.profession         || '',
      purpose:            d2.purpose            || '',
      nationality:        d2.nationality        || parsed2?.nationality        || '',
      arrivalPoint:       d2.arrivalPoint       || parsed2?.arrivalPoint       || '',
      visaType:           d2.visaType           || parsed2?.visaType           || '',
      statusText:         d2.statusText         || parsed2?.statusText         || 'تم الاستعلام',
      printPdfUrl:        submitData.data?.printPdfUrl || '',
    }};
  };

  // ─── Silent KSA visa image fetch (no modal state, for bulk) ─────────────────
  const silentKsaVisaFetch = async (
    entry: VisaEntry, retryCount = 0
  ): Promise<{success: boolean; visaPdfUrl?: string; error?: string}> => {
    if (retryCount >= 5) return { success: false, error: 'فشل بعد 5 محاولات' };
    const initRes = await fetch('/api/ksavisa/init');
    const initData = await initRes.json();
    if (!initRes.ok) return { success: false, error: initData.error || 'فشل تحميل صفحة التأشيرة' };
    if (!initData.autoSolvedCaptcha) return { success: false, error: 'تعذّر حل رمز التحقق تلقائياً' };
    const submitRes = await fetch('/api/ksavisa/submit', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId: initData.sessionId, visaNumber: entry.visaNumber || '', passportNumber: entry.passportNumber, captcha: initData.autoSolvedCaptcha, applicantName: entry.applicantName || '' }),
    });
    const submitData = await submitRes.json();
    if (!submitRes.ok) return { success: false, error: submitData.error || 'فشل الاستعلام' };
    if (!submitData.success) {
      if (submitData.notFound) return { success: false, error: submitData.error || 'التاشيرة غير موجودة' };
      await new Promise(r => setTimeout(r, 1200));
      return silentKsaVisaFetch(entry, retryCount + 1);
    }
    return { success: true, visaPdfUrl: submitData.visaPdfUrl };
  };

  // ─── Health Certificate Check ────────────────────────────────────────────────
  const hasValidAppNo = (entry: VisaEntry) => {
    const appNo = (entry.applicationNumber || '').trim().toUpperCase();
    return appNo !== '' && appNo !== '---' && appNo !== 'N/A';
  };

  const checkHealthCert = async (entry: VisaEntry) => {
    if (!hasValidAppNo(entry)) {
      showAlert(lang === 'ar' ? 'لا يوجد برنت لهذه المعاملة' : 'No print found for this entry.', 'warning');
      return;
    }
    const appNo = (entry.applicationNumber || '').trim();
    // Show overlay
    healthCertCancelRef.current = false;
    healthCertEntryIdRef.current = entry.id;
    setHealthCertOverlay(true);
    setHealthCertOverlayName(entry.applicantName || entry.passportNumber);
    setHealthCertOverlayStep(1);
    setHealthCertOverlayError('');
    setCheckingHealthCertIds(prev => new Set(prev).add(entry.id));
    try {
      // Step 2 — connecting
      await new Promise(r => setTimeout(r, 500));
      if (healthCertCancelRef.current) return;
      setHealthCertOverlayStep(2);

      const res = await fetch('/api/healthcert/check', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ applicationNumber: appNo, passportNumber: entry.passportNumber }),
      });

      if (healthCertCancelRef.current) return;

      // Step 3 — received response
      setHealthCertOverlayStep(3);
      await new Promise(r => setTimeout(r, 300));

      let data: any = {};
      try { data = await res.json(); } catch { data = { error: `خطأ في السيرفر (${res.status})` }; }

      if (healthCertCancelRef.current) return;

      if (!res.ok || !data.success) {
        setHealthCertOverlay(false);
        showAlert(data.error || 'فشل الاتصال بمنصة الشهادة الصحية.', 'error');
        return;
      }

      // Step 4 → 5 — saving result
      setHealthCertOverlayStep(4);
      await new Promise(r => setTimeout(r, 300));
      setHealthCertOverlayStep(5);
      await new Promise(r => setTimeout(r, 400));

      if (healthCertCancelRef.current) return;

      setEntries(prev => prev.map(e =>
        e.id === entry.id ? { ...e, healthCertStatus: data.result } : e
      ));
      const isIssuedResult = data.result.includes('تم اصدار') || data.result.includes('تم إصدار') || data.result.includes('الشهادة صادرة');
      setHealthCertOverlay(false);
      showAlert(data.result, isIssuedResult ? 'success' : 'info');
    } catch (err: any) {
      if (healthCertCancelRef.current) return;
      setHealthCertOverlay(false);
      showAlert('فشل الاتصال بالسيرفر: ' + err.message, 'error');
    } finally {
      setHealthCertOverlay(false);
      setCheckingHealthCertIds(prev => { const s = new Set(prev); s.delete(entry.id); return s; });
    }
  };

  // ─── Bulk sequential check ───────────────────────────────────────────────────
  const runBulkCheck = async () => {
    if (!hasPermission('update_checking')) {
      showAlert('عذراً، ليس لديك صلاحية "تحديث التشييك" لتشغيل هذا الفحص الجماعي!', 'warning');
      return;
    }
    bulkCheckStopRef.current = false;
    const snapshot = [...entries].filter(e => !e.visaImageUrl && e.txStatusType !== 'موقفة');
    setBulkCheckRunning(true);
    setBulkCheckTotal(snapshot.length);
    setBulkCheckIndex(0);
    setBulkCheckResults([]);
    setBulkCheckCurrentName('');
    setBulkCheckCurrentStatus('جاري التحضير...');

    for (let i = 0; i < snapshot.length; i++) {
      if (bulkCheckStopRef.current) break;
      const entry = snapshot[i];
      setBulkCheckIndex(i + 1);
      setBulkCheckCurrentName(entry.applicantName || entry.passportNumber);
      setBulkCheckCurrentStatus('جاري التشييك على المنصة...');

      try {
        const mofaRes = await silentMofaCheck(entry);
        if (mofaRes.success && mofaRes.merged) {
          const m = mofaRes.merged;
          // Update entry with MOFA data
          setEntries(prev => {
            const updated = prev.map(e => {
              if (e.id !== entry.id) return e;
              return {
                ...e, status: 'Found' as const,
                statusText: m.statusText || e.statusText,
                applicantName: (m.applicantName || '').trim() || e.applicantName,
                nameEnglish: m.nameEnglish || e.nameEnglish || '',
                applicationNumber: m.applicationNumber || e.applicationNumber,
                applicationDate: m.applicationDate || e.applicationDate || '',
                embassy: m.embassy || e.embassy || '',
                visaNumber: m.visaNumber || e.visaNumber || '---',
                documentNumber: m.documentNumber || e.documentNumber || '',
                entriesCount: m.entriesCount || e.entriesCount || '',
                requesterName: m.requesterName || e.requesterName || '',
                passportType: m.passportType || e.passportType || '',
                passportExpiry: m.passportExpiry || e.passportExpiry || '',
                birthDate: m.birthDate || e.birthDate || '',
                birthPlace: m.birthPlace || e.birthPlace || '',
                currentNationality: m.currentNationality || e.currentNationality || '',
                gender: m.gender || e.gender || '',
                profession: (m.profession || '').trim() || e.profession,
                purpose: m.purpose || e.purpose || '',
                nationality: m.nationality || e.nationality,
                arrivalPoint: m.arrivalPoint || e.arrivalPoint,
                visaType: m.visaType || e.visaType,
                lastUpdate: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                updatedAtUnix: Date.now(),
              };
            });
            fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
            return updated;
          });
          if (m.printPdfUrl) {
            setEntries(prev => prev.map(e => e.id === entry.id ? { ...e, printImageUrl: e.printImageUrl || m.printPdfUrl, printImageBase64: undefined } : e));
          }

          const foundVisa = m.visaNumber || '';
          if (foundVisa && foundVisa !== '---') {
            setBulkCheckCurrentStatus('تم العثور على التاشيرة — جاري جلب الصورة...');
            const entryWithVisa = { ...entry, visaNumber: foundVisa, applicantName: (m.applicantName || '').trim() || entry.applicantName };
            const visaRes = await silentKsaVisaFetch(entryWithVisa);
            if (visaRes.success && visaRes.visaPdfUrl) {
              const pdfUrl = visaRes.visaPdfUrl;
              setEntries(prev => {
                const updated = prev.map(e => e.id === entry.id ? { ...e, visaImageUrl: pdfUrl } : e);
                fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
                return updated;
              });
              setBulkCheckResults(prev => [...prev, { name: entry.applicantName || entry.passportNumber, status: 'visa_found', msg: `تم جلب صورة التاشيرة ✓ (${foundVisa})` }]);
            } else {
              setBulkCheckResults(prev => [...prev, { name: entry.applicantName || entry.passportNumber, status: 'mofa_only', msg: `تشييك MOFA ✓ — فشل جلب الصورة: ${visaRes.error}` }]);
            }
          } else {
            setBulkCheckResults(prev => [...prev, { name: entry.applicantName || entry.passportNumber, status: 'mofa_only', msg: 'تشييك MOFA ✓ — لا يوجد رقم تاشيرة بعد' }]);
          }
        } else {
          setBulkCheckResults(prev => [...prev, { name: entry.applicantName || entry.passportNumber, status: mofaRes.error?.includes('لا توجد') ? 'no_print' : 'error', msg: mofaRes.error || 'فشل التشييك' }]);
        }
      } catch (err: any) {
        setBulkCheckResults(prev => [...prev, { name: entry.applicantName || entry.passportNumber, status: 'error', msg: err.message || 'حدث خطأ غير متوقع' }]);
      }

      // Wait 7 seconds before next entry
      if (i < snapshot.length - 1 && !bulkCheckStopRef.current) {
        for (let t = 7; t > 0; t--) {
          if (bulkCheckStopRef.current) break;
          setBulkCheckCurrentStatus(`انتظار ${t} ثانية قبل العميل التالي...`);
          await new Promise(r => setTimeout(r, 1000));
        }
      }
    }

    setBulkCheckRunning(false);
    setBulkCheckCurrentStatus(bulkCheckStopRef.current ? 'تم إيقاف التشييك الجماعي' : 'اكتمل التشييك الجماعي ✓');
  };

  const autoCheckKsaVisa = async (entryId: string, visaNumber: string, passportNumber: string, applicantName: string) => {
    try {
      const initRes = await fetch('/api/ksavisa/init');
      const initData = await initRes.json();
      if (!initRes.ok || !initData.sessionId || !initData.autoSolvedCaptcha) return;

      const submitRes = await fetch('/api/ksavisa/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sessionId: initData.sessionId,
          visaNumber,
          passportNumber,
          captcha: initData.autoSolvedCaptcha,
          applicantName,
        }),
      });
      const submitData = await submitRes.json();
      if (!submitRes.ok || !submitData.success || !submitData.visaPdfUrl) return;

      const pdfUrl: string = submitData.visaPdfUrl;
      setEntries(prev => {
        const updated = prev.map(e => e.id === entryId ? { ...e, visaImageUrl: pdfUrl } : e);
        fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
        return updated;
      });

    } catch {
      // non-fatal — silent
    }
  };

  const loadMofaSession = async (entryOverride?: VisaEntry, _retryCount: number = 0) => {
    if (_retryCount >= 4) {
      const msg3 = 'فشل التشييك بعد 3 محاولات. يرجى التأكد من البيانات.';
      setMofaError(msg3);
      setMofaCheckingOverlay(false);
      setMofaCheckingError(msg3);
      setMofaInitLoading(false);
      setMofaLoading(false);
      return;
    }
    setMofaInitLoading(true);
    setMofaError('');
    setMofaResult(null);
    setMofaCaptchaCode('');
    setMofaCheckingStep(1);
    try {
      const res = await fetch('/api/mofa/init');
      if (!res.ok) throw new Error('فشل الاتصال بمنصة التأشيرات');
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      setMofaSessionId(data.sessionId);
      setMofaCaptchaKey(k => k + 1);
      const nats: {value:string;label:string}[] = data.nationalities?.length ? data.nationalities : [];
      const vts:  {value:string;label:string}[] = data.visaTypes?.length    ? data.visaTypes    : [];
      const embs: {value:string;label:string}[] = data.embassies?.length    ? data.embassies    : [];
      if (nats.length) setMofaNationalities(nats);
      if (vts.length)  setMofaVisaTypes(vts);
      if (embs.length) setMofaEmbassies(embs);

      // Auto-fill captcha from ddddocr if available
      if (data.autoSolvedCaptcha) {
        setMofaCaptchaCode(data.autoSolvedCaptcha);
        setMofaCheckingStep(2);
      }

      // Auto-match entry values against MOFA option labels
      let resolvedNat = mofaSelectedNat;
      let resolvedVt  = mofaSelectedVt;
      let resolvedEmb = mofaSelectedEmb;
      if (entryOverride) {
        const matchLabel = (list: {value:string;label:string}[], text: string) => {
          if (!text) return null;
          const t = text.trim();
          const exact = list.find(o => o.label === t);
          if (exact) return exact.value;
          const partial = list.find(o => o.label.includes(t) || t.includes(o.label));
          return partial ? partial.value : null;
        };
        const natMatch = matchLabel(nats, entryOverride.nationality || '');
        if (natMatch) { setMofaSelectedNat(natMatch); resolvedNat = natMatch; }
        const vtMatch  = matchLabel(vts,  entryOverride.visaType    || '');
        if (vtMatch)  { setMofaSelectedVt(vtMatch);  resolvedVt  = vtMatch; }
        const embMatch = matchLabel(embs, entryOverride.arrivalPoint || '');
        if (embMatch) { setMofaSelectedEmb(embMatch); resolvedEmb = embMatch; }
      }

      // Auto-submit if captcha was solved and we have a target entry
      if (data.autoSolvedCaptcha && entryOverride) {
        setMofaInitLoading(false);
        setMofaCheckingStep(3);
        await new Promise(r => setTimeout(r, 600));
        setMofaLoading(true);
        try {
          const submitRes = await fetch('/api/mofa/submit', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              sessionId: data.sessionId,
              passportNumber: entryOverride.passportNumber,
              nationality: resolvedNat,
              visaType: resolvedVt,
              embassy: resolvedEmb,
              captcha: data.autoSolvedCaptcha,
              existingPrintUrl: entryOverride.printImageUrl || '',
              applicantName: entryOverride.applicantName || '',
            }),
          });
          const submitData = await submitRes.json();
          if (!submitRes.ok) throw new Error(submitData.error || 'فشل الاستعلام');
          if (!submitData.success) {
            const isNoPrint = submitData.notFound || NO_PRINT_INDICATORS.some(ind => (submitData.error || '').includes(ind));
            const errMsg = submitData.error || 'رمز التحقق غير صحيح — يرجى المحاولة مرة أخرى.';
            setMofaError(errMsg);
            if (isNoPrint) {
              // لا يوجد بيانات في المنصة — احفظ الحالة وأظهر الخطأ
              setEntries(prev => prev.map(e => e.id === entryOverride.id ? {
                ...e,
                statusText: errMsg,
                lastUpdate: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
                updatedAtUnix: Date.now(),
              } : e));
              setMofaCheckingOverlay(false);
              setMofaCheckingError(errMsg);
              setMofaInitLoading(false);
              setMofaLoading(false);
              return;
            }
            // رمز الصورة خاطئ — أعد تحميل وحاول مجدداً (مع عداد المحاولات)
            await loadMofaSession(entryOverride, _retryCount + 1);
            return;
          }
          // Build merged result — data is nested under submitData.data
          const d2 = submitData.data || {};
          const rawHtml: string = d2.rawHtml || '';
          const parsed2 = parseMofaPageText(rawHtml);
          const appDateFallback = rawHtml.match(/تاريخ الطلب[:\s]+(\d{4}\/\d{2}\/\d{2})/)?.[1] || '';
          const merged = {
            applicantName:      d2.applicantName      || parsed2?.applicantName      || '',
            nameEnglish:        d2.nameEnglish        || '',
            applicationNumber:  d2.applicationNumber  || parsed2?.applicationNumber  || '',
            applicationDate:    d2.applicationDate    || appDateFallback             || '',
            embassy:            d2.embassy            || '',
            visaNumber:         d2.visaNumber         || parsed2?.visaNumber         || '',
            documentNumber:     d2.documentNumber     || parsed2?.documentNumber     || '',
            entriesCount:       d2.entriesCount       || '',
            requesterName:      d2.requesterName      || '',
            passportType:       d2.passportType       || '',
            passportExpiry:     d2.passportExpiry     || '',
            birthDate:          d2.birthDate          || '',
            birthPlace:         d2.birthPlace         || '',
            currentNationality: d2.currentNationality || parsed2?.nationality        || '',
            gender:             d2.gender             || '',
            profession:         d2.profession         || parsed2?.profession         || '',
            purpose:            d2.purpose            || '',
            nationality:        d2.nationality        || parsed2?.nationality        || '',
            arrivalPoint:       d2.arrivalPoint       || parsed2?.arrivalPoint       || '',
            visaType:           d2.visaType           || parsed2?.visaType           || '',
            statusText:         d2.statusText         || parsed2?.statusText         || 'تم الاستعلام',
          };
          setMofaResult(merged);
          setMofaCheckingStep(4);

          // Auto-save to entries immediately (same logic as handleMofaLiveSubmit)
          const autoEntryId = entryOverride.id;
          setEntries(prev => prev.map(e => {
            if (e.id !== autoEntryId) return e;
            return {
              ...e,
              status: 'Found' as const,
              statusText: merged.statusText || e.statusText || 'تم الاستعلام',
              applicantName: (merged.applicantName || '').trim() || e.applicantName,
              nameEnglish: merged.nameEnglish || e.nameEnglish || '',
              applicationNumber: merged.applicationNumber || e.applicationNumber,
              applicationDate: merged.applicationDate || e.applicationDate || '',
              embassy: merged.embassy || e.embassy || '',
              visaNumber: merged.visaNumber || e.visaNumber || '---',
              documentNumber: merged.documentNumber || e.documentNumber || '',
              entriesCount: merged.entriesCount || e.entriesCount || '',
              requesterName: merged.requesterName || e.requesterName || '',
              passportType: merged.passportType || e.passportType || '',
              passportExpiry: merged.passportExpiry || e.passportExpiry || '',
              birthDate: merged.birthDate || e.birthDate || '',
              birthPlace: merged.birthPlace || e.birthPlace || '',
              currentNationality: merged.currentNationality || e.currentNationality || '',
              gender: merged.gender || e.gender || '',
              profession: (merged.profession || '').trim() || e.profession,
              purpose: merged.purpose || e.purpose || '',
              nationality: merged.nationality || e.nationality,
              arrivalPoint: merged.arrivalPoint || e.arrivalPoint,
              visaType: merged.visaType || e.visaType,
              lastUpdate: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
              updatedAtUnix: Date.now(),
            };
          }));

          // Save PDF URL returned from submit endpoint
          const autoPdfUrl = submitData.data?.printPdfUrl as string | undefined;
          if (autoPdfUrl) {
            setEntries(prev => prev.map(e => e.id === autoEntryId ? { ...e, printImageUrl: e.printImageUrl || autoPdfUrl, printImageBase64: undefined } : e));
          }

          // Auto-check visa if visa number found
          const autoVisaNum = merged.visaNumber || '';
          setMofaCheckingStep(5);

          // Show success animation then close
          const hasPrint2 = !!(merged.applicationNumber && merged.applicationNumber !== '---');
          setMofaCheckingOverlay(false);
          setMofaSaveSuccessAnim(true);

          if (autoVisaNum && autoVisaNum !== '---') {
            // Close MOFA modal then open KSA Visa fetch modal automatically
            const entryWithVisa = { ...entryOverride, visaNumber: autoVisaNum };
            setTimeout(() => {
              setMofaSaveSuccessAnim(false);
              setMofaFetchModalOpen(false);
              setMofaFetchEntry(null);
              setMofaPasteText('');
              setMofaResult(null);
              if (hasPrint2) setAllTabFilter('with_print');
              setActiveTab('all');
              // Auto-open visa image fetch modal
              openKsaVisaModal(entryWithVisa);
            }, 1200);
          } else {
            setTimeout(() => {
              setMofaSaveSuccessAnim(false);
              setMofaFetchModalOpen(false);
              setMofaFetchEntry(null);
              setMofaPasteText('');
              setMofaResult(null);
              if (hasPrint2) setAllTabFilter('with_print');
              setActiveTab('all');
            }, 2200);
          }

        } catch (autoErr: any) {
          const isNoPrintCatch = NO_PRINT_INDICATORS.some(ind => (autoErr.message || '').includes(ind));
          setMofaError(autoErr.message || 'فشل الإرسال التلقائي. يرجى المحاولة يدوياً.');
          if (isNoPrintCatch && entryOverride) {
            setEntries(prev => prev.map(e => e.id === entryOverride.id ? {
              ...e,
              statusText: autoErr.message || 'حدث خطأ يرجى التأكد من البيانات المدخلة والمحاولة مرة أخرى',
              lastUpdate: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
              updatedAtUnix: Date.now(),
            } : e));
            setMofaCheckingOverlay(false);
            setMofaCheckingError(autoErr.message || 'لا توجد بيانات لهذا الجواز في المنصة');
          } else {
            setMofaCheckingOverlay(false);
            setMofaCheckingError(autoErr.message || 'فشل الإرسال التلقائي. يرجى المحاولة مرة أخرى.');
          }
        } finally {
          setMofaLoading(false);
        }
        return;
      }
    } catch (e: any) {
      setMofaError(e.message || 'فشل الاتصال بمنصة التأشيرات');
      setMofaCheckingOverlay(false);
      setMofaCheckingError(e.message || 'فشل الاتصال بمنصة التأشيرات');
    } finally {
      setMofaInitLoading(false);
    }
  };

  const checkVisa = async (id: string, isSilentBulk: boolean = false) => {
    if (!hasPermission('update_checking')) {
      showAlert(lang === 'ar' ? 'عذراً، ليس لديك صلاحية "تحديث التشييك" لتشغيل هذا الفحص!' : 'Sorry, you do not have permission to run status checks!', 'warning');
      return;
    }
    const entry = entries.find(e => e.id === id);
    if (!entry || !entry.passportNumber) return;

    setMofaFetchEntry(entry);
    setMofaPasteText('');
    setMofaResult(null);
    setMofaError('');
    setMofaCaptchaCode('');
    setMofaCheckingOverlay(true);
    setMofaSaveSuccessAnim(false);
    setMofaCheckingError('');
    setMofaCheckingStep(0);
    setMofaFetchModalOpen(true);
    // Auto-load MOFA session and auto-match entry values
    setTimeout(() => loadMofaSession(entry), 50);
  };

  const handleMofaLiveSubmit = async () => {
    if (!mofaSessionId || !mofaCaptchaCode.trim() || !mofaFetchEntry) return;
    setMofaLoading(true);
    setMofaError('');
    try {
      const res = await fetch('/api/mofa/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sessionId: mofaSessionId,
          passportNumber: mofaFetchEntry.passportNumber,
          nationality: mofaSelectedNat,
          visaType: mofaSelectedVt,
          embassy: mofaSelectedEmb,
          captcha: mofaCaptchaCode.trim(),
          existingPrintUrl: mofaFetchEntry.printImageUrl || '',
          applicantName: mofaFetchEntry.applicantName || '',
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'فشل الاستعلام');
      if (!data.success) {
        const isNoPrint = data.notFound || NO_PRINT_INDICATORS.some(ind => (data.error || '').includes(ind));
        setMofaError(data.error || 'فشل الاستعلام، يرجى المحاولة مرة أخرى.');
        if (isNoPrint) {
          // لا يوجد بيانات في المنصة — احفظ الحالة ولا تعيد تحميل الكابتشا
          if (mofaFetchEntry) {
            setEntries(prev => prev.map(e => e.id === mofaFetchEntry!.id ? {
              ...e,
              statusText: data.error || 'حدث خطأ يرجى التأكد من البيانات المدخلة والمحاولة مرة أخرى',
              lastUpdate: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
              updatedAtUnix: Date.now(),
            } : e));
          }
          return;
        }
        // رمز الصورة خاطئ — أعد تحميل الكابتشا وحاول مجدداً
        await loadMofaSession();
        return;
      }
      // Apply client-side regex fallback on rawHtml if fields are empty
      const raw = data.data.rawHtml || '';
      const parsed = parseMofaPageText(raw);
      const d = data.data;
      const merged = {
        applicantName: d.applicantName || parsed?.applicantName || '',
        nameEnglish: d.nameEnglish || '',
        applicationNumber: d.applicationNumber || parsed?.applicationNumber || '',
        applicationDate: d.applicationDate || '',
        embassy: d.embassy || '',
        visaNumber: d.visaNumber || parsed?.visaNumber || '',
        documentNumber: d.documentNumber || parsed?.documentNumber || '',
        entriesCount: d.entriesCount || '',
        requesterName: d.requesterName || '',
        passportType: d.passportType || '',
        passportExpiry: d.passportExpiry || '',
        birthDate: d.birthDate || '',
        birthPlace: d.birthPlace || '',
        currentNationality: d.currentNationality || parsed?.nationality || '',
        gender: d.gender || '',
        profession: d.profession || parsed?.profession || '',
        purpose: d.purpose || '',
        nationality: d.nationality || parsed?.nationality || '',
        arrivalPoint: d.arrivalPoint || parsed?.arrivalPoint || '',
        visaType: d.visaType || parsed?.visaType || '',
        statusText: d.statusText || parsed?.statusText || '',
      };
      setMofaResult(merged);

      // Auto-save immediately after successful query
      const mofaEntryId = mofaFetchEntry!.id;
      setEntries(prev => prev.map(e => {
        if (e.id !== mofaEntryId) return e;
        return {
          ...e,
          status: 'Found' as const,
          statusText: merged.statusText || e.statusText || 'تم اصدار التأشيرة',
          applicantName: (merged.applicantName || '').trim() || e.applicantName,
          nameEnglish: merged.nameEnglish || e.nameEnglish || '',
          applicationNumber: merged.applicationNumber || e.applicationNumber,
          applicationDate: merged.applicationDate || e.applicationDate || '',
          embassy: merged.embassy || e.embassy || '',
          visaNumber: merged.visaNumber || e.visaNumber || '---',
          documentNumber: merged.documentNumber || e.documentNumber || '',
          entriesCount: merged.entriesCount || e.entriesCount || '',
          requesterName: merged.requesterName || e.requesterName || '',
          passportType: merged.passportType || e.passportType || '',
          passportExpiry: merged.passportExpiry || e.passportExpiry || '',
          birthDate: merged.birthDate || e.birthDate || '',
          birthPlace: merged.birthPlace || e.birthPlace || '',
          currentNationality: merged.currentNationality || e.currentNationality || '',
          gender: merged.gender || e.gender || '',
          profession: (merged.profession || '').trim() || e.profession,
          purpose: merged.purpose || e.purpose || '',
          nationality: merged.nationality || e.nationality,
          arrivalPoint: merged.arrivalPoint || e.arrivalPoint,
          visaType: merged.visaType || e.visaType,
          lastUpdate: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          updatedAtUnix: Date.now(),
          dataUpdatedAtUnix: Date.now(),
        };
      }));

      // Save PDF URL returned from submit endpoint
      const printPdfUrl = data.data?.printPdfUrl as string | undefined;
      if (printPdfUrl) {
        setEntries(prev => prev.map(e => e.id === mofaEntryId ? { ...e, printImageUrl: e.printImageUrl || printPdfUrl, printImageBase64: undefined } : e));
      }

      // Auto-open visa image fetch modal if visa number found
      const foundVisaNum = merged.visaNumber || '';
      const savedHasPrint = !!(merged.applicationNumber && merged.applicationNumber !== '---' && merged.applicationNumber !== '');

      if (foundVisaNum && foundVisaNum !== '---') {
        // Close MOFA modal then open KSA Visa fetch modal automatically
        const entryWithVisa = { ...mofaFetchEntry!, visaNumber: foundVisaNum };
        setTimeout(() => {
          setMofaFetchModalOpen(false);
          setMofaFetchEntry(null);
          setMofaPasteText('');
          setMofaResult(null);
          if (savedHasPrint) setAllTabFilter('with_print');
          setActiveTab('all');
          openKsaVisaModal(entryWithVisa);
        }, 1000);
      } else {
        setTimeout(() => {
          setMofaFetchModalOpen(false);
          setMofaFetchEntry(null);
          setMofaPasteText('');
          setMofaResult(null);
          if (savedHasPrint) setAllTabFilter('with_print');
          setActiveTab('all');
        }, 1500);
      }

    } catch (e: any) {
      const isNoPrintErr = NO_PRINT_INDICATORS.some(ind => (e.message || '').includes(ind));
      setMofaError(e.message || 'حدث خطأ، يرجى المحاولة مرة أخرى.');
      if (isNoPrintErr) {
        if (mofaFetchEntry) {
          setEntries(prev => prev.map(en => en.id === mofaFetchEntry!.id ? {
            ...en,
            statusText: e.message || 'حدث خطأ يرجى التأكد من البيانات المدخلة والمحاولة مرة أخرى',
            lastUpdate: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
            updatedAtUnix: Date.now(),
          } : en));
        }
      } else {
        await loadMofaSession();
      }
    } finally {
      setMofaLoading(false);
    }
  };



  // ── KSA Visa Track Functions ────────────────────────────────────────────────
  const loadKsaVisaSession = async (entryOverride?: VisaEntry, _retryCount: number = 0) => {
    if (_retryCount >= 5) {
      const msg5 = 'فشل جلب الصورة بعد 5 محاولات. يرجى التأكد من البيانات.';
      setKsaVisaError(msg5);
      setKsaVisaCheckingOverlay(false);
      setKsaVisaCheckingError(msg5);
      setKsaVisaInitLoading(false);
      setKsaVisaLoading(false);
      return;
    }
    setKsaVisaInitLoading(true);
    setKsaVisaError('');
    setKsaVisaSessionId('');
    setKsaVisaCaptchaCode('');
    if (_retryCount === 0) setKsaVisaSuccess(false);
    setKsaVisaCheckingStep(1);
    try {
      const res = await fetch('/api/ksavisa/init');
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'فشل تحميل صفحة التأشيرة');
      setKsaVisaSessionId(data.sessionId);
      setKsaVisaCaptchaKey(k => k + 1);
      if (data.autoSolvedCaptcha) {
        setKsaVisaCaptchaCode(data.autoSolvedCaptcha);
        setKsaVisaCheckingStep(2);
      }
      if (data.autoSolvedCaptcha && entryOverride) {
        setKsaVisaInitLoading(false);
        setKsaVisaCheckingStep(3);
        await new Promise(r => setTimeout(r, 600));
        setKsaVisaLoading(true);
        try {
          const submitRes = await fetch('/api/ksavisa/submit', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              sessionId: data.sessionId,
              visaNumber: entryOverride.visaNumber || '',
              passportNumber: entryOverride.passportNumber,
              captcha: data.autoSolvedCaptcha,
              applicantName: entryOverride.applicantName || '',
            }),
          });
          const submitData = await submitRes.json();
          if (!submitRes.ok) throw new Error(submitData.error || 'فشل الاستعلام');
          if (!submitData.success) {
            const isNoPrint = submitData.notFound;
            const errMsg = submitData.error || 'رمز التحقق غير صحيح';
            setKsaVisaError(errMsg);
            if (isNoPrint) {
              setKsaVisaCheckingOverlay(false);
              setKsaVisaCheckingError(errMsg);
              setKsaVisaInitLoading(false);
              setKsaVisaLoading(false);
              return;
            }
            await loadKsaVisaSession(entryOverride, _retryCount + 1);
            return;
          }
          setKsaVisaCheckingStep(4);
          await new Promise(r => setTimeout(r, 300));
          setKsaVisaCheckingStep(5);
          if (submitData.visaPdfUrl) {
            const pdfUrl: string = submitData.visaPdfUrl;
            setEntries(prev => {
              const updated = prev.map(e => e.id === entryOverride.id ? { ...e, visaImageUrl: pdfUrl } : e);
              fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
              return updated;
            });
          }
          setKsaVisaCheckingOverlay(false);
          setKsaVisaSaveSuccessAnim(true);
          setTimeout(() => {
            setKsaVisaSaveSuccessAnim(false);
            setKsaVisaModalOpen(false);
            setKsaVisaEntry(null);
            setKsaVisaCheckingError('');
            setKsaVisaSessionId('');
            setKsaVisaCaptchaCode('');
            setKsaVisaError('');
          }, 2200);
        } catch (autoErr: any) {
          const errCatch = autoErr.message || 'فشل الإرسال التلقائي';
          setKsaVisaCheckingOverlay(false);
          setKsaVisaCheckingError(errCatch);
          setKsaVisaInitLoading(false);
          setKsaVisaLoading(false);
        } finally {
          setKsaVisaLoading(false);
        }
      }
    } catch (err: any) {
      setKsaVisaError(err.message || 'فشل الاتصال بالموقع');
      setKsaVisaCheckingOverlay(false);
      setKsaVisaCheckingError(err.message || 'فشل الاتصال بالموقع');
    } finally {
      setKsaVisaInitLoading(false);
    }
  };

  const openKsaVisaModal = (entry: VisaEntry) => {
    setKsaVisaEntry(entry);
    setKsaVisaNumber(entry.visaNumber || '');
    setKsaVisaPassportNumber(entry.passportNumber || '');
    setKsaVisaModalOpen(true);
    setKsaVisaError('');
    setKsaVisaSuccess(false);
    setKsaVisaSessionId('');
    setKsaVisaCaptchaCode('');
    setKsaVisaCheckingOverlay(true);
    setKsaVisaSaveSuccessAnim(false);
    setKsaVisaCheckingError('');
    setKsaVisaCheckingStep(0);
    loadKsaVisaSession(entry, 0);
  };

  const handleKsaVisaSubmit = async () => {
    if (!ksaVisaSessionId || !ksaVisaCaptchaCode.trim() || !ksaVisaEntry) return;
    setKsaVisaCheckingOverlay(true);
    setKsaVisaCheckingStep(3);
    setKsaVisaLoading(true);
    setKsaVisaError('');
    try {
      const res = await fetch('/api/ksavisa/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          sessionId: ksaVisaSessionId,
          visaNumber: ksaVisaNumber || ksaVisaEntry.visaNumber,
          passportNumber: ksaVisaPassportNumber || ksaVisaEntry.passportNumber,
          captcha: ksaVisaCaptchaCode,
          applicantName: ksaVisaEntry.applicantName || '',
        }),
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || 'فشل جلب التأشيرة');
      if (!data.success) {
        const errMsg = data.error || 'فشل جلب التأشيرة';
        setKsaVisaError(errMsg);
        setKsaVisaCheckingOverlay(false);
        setKsaVisaCheckingError(errMsg);
        return;
      }
      setKsaVisaCheckingStep(4);
      await new Promise(r => setTimeout(r, 300));
      setKsaVisaCheckingStep(5);
      if (data.visaPdfUrl) {
        setEntries(prev => {
          const updated = prev.map(e => e.id === ksaVisaEntry!.id ? { ...e, visaImageUrl: data.visaPdfUrl } : e);
          fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
          return updated;
        });
      }
      setKsaVisaCheckingOverlay(false);
      setKsaVisaSaveSuccessAnim(true);
      setTimeout(() => {
        setKsaVisaSaveSuccessAnim(false);
        setKsaVisaModalOpen(false);
        setKsaVisaEntry(null);
        setKsaVisaCheckingError('');
        setKsaVisaSessionId('');
        setKsaVisaCaptchaCode('');
        setKsaVisaError('');
      }, 2200);
    } catch (err: any) {
      setKsaVisaCheckingOverlay(false);
      setKsaVisaCheckingError(err.message || 'فشل جلب التأشيرة');
      setKsaVisaError(err.message || 'فشل جلب التأشيرة');
    } finally {
      setKsaVisaLoading(false);
    }
  };

  const handleSaveMofaParsedData = () => {
    if (!mofaFetchEntry) return;

    const result = mofaResult || parseMofaPageText(mofaPasteText);
    const mofaEntryId = mofaFetchEntry.id;

    if (!result) {
      showAlert(lang === 'ar' ? '⚠️ لا توجد بيانات للحفظ. يرجى إجراء الاستعلام أولاً.' : 'No data to save. Please query first.', 'warning');
      return;
    }

    // Overwrite with actual parsed values
    const updatedEntries = entries.map(e => {
      if (e.id === mofaEntryId) {
        return {
          ...e,
          status: 'Found' as const,
          statusText: result.statusText || e.statusText || 'تم اصدار التأشيرة',
          applicantName: (result.applicantName || '').trim() || e.applicantName,
          nameEnglish: result.nameEnglish || e.nameEnglish || '',
          applicationNumber: result.applicationNumber || e.applicationNumber,
          applicationDate: result.applicationDate || e.applicationDate || '',
          embassy: result.embassy || e.embassy || '',
          visaNumber: result.visaNumber || e.visaNumber || '---',
          documentNumber: result.documentNumber || e.documentNumber || '',
          entriesCount: result.entriesCount || e.entriesCount || '',
          requesterName: result.requesterName || e.requesterName || '',
          passportType: result.passportType || e.passportType || '',
          passportExpiry: result.passportExpiry || e.passportExpiry || '',
          birthDate: result.birthDate || e.birthDate || '',
          birthPlace: result.birthPlace || e.birthPlace || '',
          currentNationality: result.currentNationality || e.currentNationality || '',
          gender: result.gender || e.gender || '',
          profession: (result.profession || '').trim() || e.profession,
          purpose: result.purpose || e.purpose || '',
          nationality: result.nationality || e.nationality,
          arrivalPoint: result.arrivalPoint || e.arrivalPoint,
          visaType: result.visaType || e.visaType,
          lastUpdate: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
          updatedAtUnix: Date.now(),
          dataUpdatedAtUnix: Date.now(),
          applicantData: mofaPasteText
        };
      }
      return e;
    });

    setEntries(updatedEntries);

    setMofaCheckingOverlay(false);
    setMofaSaveSuccessAnim(true);
    setTimeout(() => {
      setMofaSaveSuccessAnim(false);
      setMofaFetchModalOpen(false);
      setMofaFetchEntry(null);
      setMofaPasteText('');
      setAllTabFilter('with_print');
    }, 2200);
  };



  const checkAll = async () => {
    if (bulkCheckRunning) {
      bulkCheckStopRef.current = true;
      setBulkCheckRunning(false);
      setBulkCheckCurrentStatus('تم إيقاف التشييك الجماعي');
      return;
    }
    await runBulkCheck();
  };

  const handleAddNewCustomerManually = () => {
    if (!newCustName.trim()) {
      showAlert('يرجى كتابة اسم العميل الكامل — الحقل إجباري!', 'warning');
      return;
    }
    if (!newCustPassport.trim()) {
      showAlert('يرجى إدخال رقم جواز السفر — الحقل إجباري!', 'warning');
      return;
    }
    if (!newCustOfficeId && !newCustPhone.trim()) {
      showAlert('يرجى إدخال رقم الهاتف — الحقل إجباري عند عدم تحديد مكتب!', 'warning');
      return;
    }
    setNewCustSaving(true);
    setTimeout(() => { handleAddNewCustomerManuallyCommit(); }, 900);
  };
  const handleAddNewCustomerManuallyCommit = () => {
    // Check duplication
    const isDuplicate = entries.some(e =>
      e.passportNumber.toUpperCase() === newCustPassport.trim().toUpperCase()
    );
    if (isDuplicate) {
      showConfirm('تحذير: يوجد بالفعل عميل مسجل بنفس جواز السفر. هل تود الاستمرار؟', () => {
        doAddNewCustomerCommit();
      });
      setNewCustSaving(false);
      return;
    }
    doAddNewCustomerCommit();
  };
  const doAddNewCustomerCommit = () => {
    const finalDocNumber = newCustCustomId.trim();
    const finalPhone = newCustPhone.trim() || finalDocNumber || '';

    const newEntry: VisaEntry = {
      id: 'ID-' + Math.floor(Math.random() * 9000000 + 1000000),
      passportNumber: newCustPassport.trim().toUpperCase(),
      applicationNumber: '---',
      customId: finalDocNumber,
      applicantName: newCustName.trim(),
      nationality: newCustNationality,
      arrivalPoint: newCustArrival,
      visaType: newCustVisaType,
      phoneNumber: finalPhone,
      profession: '---',
      visaNumber: '---',
      status: 'Idle',
      statusText: 'قيد الانتظار (Idle)',
      applicantData: newCustReceiveDate ? `تاريخ الاستلام: ${newCustReceiveDate}` : '',
      lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true }),
      updatedAtUnix: Date.now(),
      officeId: newCustOfficeId || undefined
    };

    setEntries(prev => [newEntry, ...prev]);
    addSyncLog(`تم تسجيل العميل الجديد ${newCustName.trim()} بنجاح عبر مدير العملاء.`);

    // Reset fields
    setNewCustName('');
    setNewCustPassport('');
    setNewCustPhone('');
    setNewCustCustomId('');
    setNewCustReceiveDate(new Date().toISOString().slice(0, 10));
    setNewCustNationality('اليمن');
    setNewCustArrival('عدن');
    setNewCustVisaType('عمل');
    setNewCustOfficeId('');
    setNewCustSaving(false);

    setCustomerSubTab('none');
  };

  const handleCustomerExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];

        if (data.length === 0) {
          showAlert(lang === 'ar' ? 'ملف الاكسل فارغ!' : 'The selected spreadsheet is empty!', 'warning');
          return;
        }

        // Detect header row
        const firstRowStr = (data[0] || []).map(cell => String(cell || '').toLowerCase());
        const hasHeader = firstRowStr.some(cell =>
          cell.includes('passport') || cell.includes('جواز') ||
          cell.includes('هاتف') || cell.includes('phone') ||
          cell.includes('تلفون') || cell.includes('اسم') || cell.includes('name')
        );
        const startIndex = hasHeader ? 1 : 0;

        // Column index map
        let passportIdx = 0;
        let phoneIdx = -1;
        let nameIdx = -1;
        let receiveDateIdx = -1;
        let visaTypeIdx = -1;
        let arrivalIdx = -1;
        let officeNameIdx = -1;

        if (hasHeader) {
          data[0].forEach((h, idx) => {
            const headStr = String(h || '').trim();
            const hl = headStr.toLowerCase();
            if (headStr.includes('جواز') || hl.includes('passport')) {
              passportIdx = idx;
            } else if (headStr.includes('هاتف') || headStr.includes('تلفون') || headStr.includes('جوال') || hl.includes('phone') || hl.includes('mobile')) {
              phoneIdx = idx;
            } else if (headStr.includes('اسم') || hl.includes('name')) {
              nameIdx = idx;
            } else if (headStr.includes('استلام') || headStr.includes('تاريخ') || hl.includes('receive') || hl.includes('date')) {
              receiveDateIdx = idx;
            } else if (headStr.includes('نوع') || hl.includes('type') || hl.includes('visa type')) {
              visaTypeIdx = idx;
            } else if (headStr.includes('قدوم') || headStr.includes('وصول') || headStr.includes('سفارة') || hl.includes('arrival') || hl.includes('embassy')) {
              arrivalIdx = idx;
            } else if (headStr.includes('مكتب') || headStr.includes('وكيل') || hl.includes('office') || hl.includes('agent')) {
              officeNameIdx = idx;
            }
          });
        }

        const existingPassportSet = new Set(entries.map(e => e.passportNumber.toUpperCase()));
        const seenPassportInUpload = new Set<string>();
        const skipped: string[] = [];

        const newEntries: VisaEntry[] = data.slice(startIndex).flatMap((row: any[]) => {
          // Passport required
          const pNum = String(row[passportIdx] || '').trim().toUpperCase();
          if (!pNum) return [];

          // Name required
          const nameVal = nameIdx !== -1 ? String(row[nameIdx] || '').trim() : '';
          if (!nameVal) { skipped.push(pNum || '—'); return []; }

          // Duplicate check
          if (existingPassportSet.has(pNum) || seenPassportInUpload.has(pNum)) {
            skipped.push(pNum); return [];
          }
          seenPassportInUpload.add(pNum);

          // Office name → find matching office id
          const officeNameVal = officeNameIdx !== -1 ? String(row[officeNameIdx] || '').trim() : '';
          const matchedOffice = officeNameVal
            ? offices.find(o => o.name.trim() === officeNameVal || o.name.trim().toLowerCase() === officeNameVal.toLowerCase())
            : undefined;
          const officeId = matchedOffice?.id || undefined;

          // Phone required if no office
          const phoneVal = phoneIdx !== -1 && row[phoneIdx] ? String(row[phoneIdx] || '').trim() : '';
          if (!officeId && !phoneVal) { skipped.push(pNum); return []; }

          const finalPhone = phoneVal || matchedOffice?.phone || '';
          const receiveDateVal = receiveDateIdx !== -1 ? String(row[receiveDateIdx] || '').trim() : '';
          const vtVal = visaTypeIdx !== -1 && row[visaTypeIdx] ? String(row[visaTypeIdx] || '').trim() : 'عمل';
          const arrVal = arrivalIdx !== -1 && row[arrivalIdx] ? String(row[arrivalIdx] || '').trim() : 'عدن';

          const entry: VisaEntry = {
            id: 'ID-' + Math.floor(Math.random() * 9000000 + 1000000),
            passportNumber: pNum,
            applicationNumber: '---',
            applicantName: nameVal,
            nationality: 'اليمن',
            arrivalPoint: arrVal,
            visaType: vtVal,
            phoneNumber: finalPhone,
            profession: '---',
            visaNumber: '---',
            status: 'Idle',
            statusText: 'قيد الانتظار (Idle)',
            applicantData: receiveDateVal ? `تاريخ الاستلام: ${receiveDateVal}` : '',
            officeId,
            lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true }),
            updatedAtUnix: Date.now()
          };
          return [entry];
        });

        if (newEntries.length > 0) {
          setEntries(prev => [...newEntries, ...prev]);
          setCustomerExcelAddedCount(newEntries.length);
          setCustomerExcelSkippedCount(skipped.length);
          setCustomerExcelAddedEntries(newEntries);
          setCustomerExcelSuccess(true);
          addSyncLog(`تم استيراد ${newEntries.length} عميل من ملف إكسل بنجاح. (متجاهل: ${skipped.length})`);
        } else {
          showAlert(lang === 'ar' ? `لم يُضَف أي سجل. تجاهُل ${skipped.length} صف لعدم استيفاء الشروط.` : 'No valid records imported. Rows skipped due to missing required fields.', 'warning');
        }
      } catch (err) {
        console.error(err);
        showAlert(lang === 'ar' ? 'خطأ في قراءة ملف الإكسل. الرجاء التأكد من التنسيق.' : 'Error reading Excel file. Please double check file format.', 'error');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleSaveEditedCustomerManually = () => {
    if (!selectedCustomerToEdit) return;
    if (!editCustPassport.trim() || !editCustPhone.trim()) {
      showAlert(lang === 'ar' ? 'يرجى ملء الحقول الإجبارية: رقم جواز السفر ورقم الهاتف!' : 'Please fill in the required fields: Passport Number and Phone Number!', 'warning');
      return;
    }

    const finalName = editCustName.trim() || '---';
    setEntries(prev => prev.map(e => {
      if (e.id === selectedCustomerToEdit.id) {
        return {
          ...e,
          passportNumber: editCustPassport.trim().toUpperCase(),
          applicantName: finalName,
          nationality: editCustNationality.trim(),
          arrivalPoint: editCustArrival.trim(),
          visaType: editCustVisaType.trim(),
          phoneNumber: editCustPhone.trim(),
          profession: editCustProfession.trim() || '---',
          visaNumber: editCustVisaNumber.trim() || '---',
          status: editCustStatus,
          statusText: editCustStatusText.trim(),
          applicantData: editCustApplicantData.trim(),
          lastUpdate: new Date().toLocaleString('ar-YE', { hour12: true }),
          updatedAtUnix: Date.now()
        };
      }
      return e;
    }));

    addSyncLog(lang === 'ar' ? `تم تعديل بيانات العميل ${finalName} بنجاح.` : `Client ${finalName} details saved and refreshed.`);
    setSelectedCustomerToEdit(null);
    setCustomerSubTab('none');
    setAllTabFilter('all');
    setActiveTab('all');
    showAlert(lang === 'ar' ? 'تم حفظ وتحديث بيانات العميل بنجاح.' : 'Customer information saved and synchronized.', 'success');
  };

  const handleDeleteCustomerManually = (entryToDelete: VisaEntry) => {
    setArchiveConfirmEntry(entryToDelete);
  };

  const handlePermDeleteEntry = (entryToDelete: VisaEntry) => {
    setArchiveConfirmEntry(null);
    setEntries(prev => {
      const updated = prev.filter(e => e.id !== entryToDelete.id);
      fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
      return updated;
    });
    addSyncLog(lang === 'ar' ? `تم حذف العميل ${entryToDelete.applicantName} نهائياً بواسطة مدير العملاء.` : `Client ${entryToDelete.applicantName} deleted from records.`);
    setSelectedCustomerToEdit(null);
    setCustomerSubTab('none');
    // NOTE: files (prints/visa PDFs) are intentionally NOT deleted — they remain on the server and Telegram
  };

  const exitSelectMode = () => {
    setIsSelectMode(false);
    setSelectedEntryIds(new Set());
    setShowBulkMenu(false);
  };

  const toggleSelectEntry = (id: string) => {
    setSelectedEntryIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id); else next.add(id);
      return next;
    });
  };

  const handleBulkRefresh = () => {
    selectedEntryIds.forEach(id => checkVisa(id));
    setShowBulkMenu(false);
  };

  const handleBulkHealthCert = () => {
    const selected = entries.filter(e => {
      if (!selectedEntryIds.has(e.id)) return false;
      if (!hasValidAppNo(e)) return false;
      const hcs = (e.healthCertStatus || '').trim();
      if (hcs.includes('تم اصدار') || hcs.includes('تم إصدار') || hcs.includes('الشهادة صادرة')) return false;
      return true;
    });
    selected.forEach(e => checkHealthCert(e));
    setShowBulkMenu(false);
  };

  const handleBulkArchive = () => {
    setShowBulkMenu(false);
    showConfirm(
      lang === 'ar'
        ? `هل أنت متأكد من أرشفة ${selectedEntryIds.size} عميل/عملاء؟`
        : `Archive ${selectedEntryIds.size} record(s)?`,
      () => {
        const toArchive = entries.filter(e => selectedEntryIds.has(e.id));
        toArchive.forEach(e => handleMoveToArchive(e));
        exitSelectMode();
      }
    );
  };

  const handleBulkDelete = () => {
    setShowBulkMenu(false);
    const idsToDelete = new Set(selectedEntryIds);
    showConfirm(
      lang === 'ar'
        ? `هل أنت متأكد من حذف ${idsToDelete.size} عميل/عملاء نهائياً؟ لا يمكن التراجع.`
        : `Permanently delete ${idsToDelete.size} record(s)? This cannot be undone.`,
      () => {
        setEntries(prev => {
          const updated = prev.filter(e => !idsToDelete.has(e.id));
          fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
          return updated;
        });
        addSyncLog(lang === 'ar' ? `تم حذف ${idsToDelete.size} عميل نهائياً.` : `${idsToDelete.size} records deleted.`);
        exitSelectMode();
      }
    );
  };

  const handleBulkTransfer = (officeId: string | 'general') => {
    const isGeneral = !officeId || officeId === 'general';
    const office = isGeneral ? null : offices.find(o => o.id === officeId);
    if (!isGeneral && !office) return;
    setShowTransferModal(false);
    const label = isGeneral ? (lang === 'ar' ? 'القسم العام' : 'General') : office!.name;
    setTimeout(() => {
      showConfirm(
        lang === 'ar'
          ? `هل أنت متأكد من نقل ${selectedEntryIds.size} عميل/عملاء إلى "${label}"؟`
          : `Transfer ${selectedEntryIds.size} record(s) to "${label}"?`,
        () => {
          setEntries(prev => prev.map(e => {
            if (!selectedEntryIds.has(e.id)) return e;
            if (isGeneral) {
              // نقل إلى العام: يُفرَّغ الرقم المخصص، يبقى رقم الهاتف كما هو
              return { ...e, officeId: undefined, customId: '' };
            }
            // نقل إلى مكتب: الرقم المخصص = رقم المكتب، رقم الهاتف لا يتغير
            return {
              ...e,
              officeId: office!.id,
              customId: office!.phone || '',
            };
          }));
          addSyncLog(lang === 'ar' ? `تم نقل ${selectedEntryIds.size} عميل إلى ${label}.` : `${selectedEntryIds.size} records transferred to ${label}.`);
          setShowBulkMenu(false);
          exitSelectMode();
        }
      );
    }, 200);
  };

  const resolveTgUrl = (url: string | undefined): string => {
    if (!url || !url.startsWith('tg:')) return url || '';
    const parts = url.slice(3).split(':');
    const fileId = parts[0];
    const filename = parts[1] ? decodeURIComponent(parts[1]) : 'file.pdf';
    return `/api/tg-download?fileId=${encodeURIComponent(fileId)}&filename=${encodeURIComponent(filename)}`;
  };

  const downloadUrl = (url: string | undefined, fallbackName: string): void => {
    if (!url) return;
    const href = resolveTgUrl(url);
    const a = document.createElement('a');
    a.href = href;
    a.download = fallbackName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  const [archiveUploading, setArchiveUploading] = useState(false);
  const [bulkTgUploading, setBulkTgUploading] = useState(false);
  const [bulkTgResult, setBulkTgResult] = useState<{uploaded: number; skipped: number; failed: number} | null>(null);

  const handleBulkUploadToTelegram = () => {
    showConfirm(
      lang === 'ar' ? 'سيتم رفع جميع الملفات المحلية إلى تيليغرام وحذفها من النظام. هل أنت متأكد؟' : 'All local files will be uploaded to Telegram and deleted from the system. Are you sure?',
      async () => {
    setBulkTgUploading(true);
    setBulkTgResult(null);
    try {
      const res = await fetch('/api/bulk-upload-to-telegram', { method: 'POST' });
      const data = await res.json();
      if (data.error === 'no_token') {
        showAlert(lang === 'ar' ? '⚠️ لم يتم ضبط توكن البوت في إعدادات التيليغرام' : '⚠️ Bot token not configured in Telegram settings', 'warning');
      } else if (data.error === 'no_chat_id') {
        showAlert(lang === 'ar' ? '⚠️ لم يتم ضبط معرف قناة الأرشيف في إعدادات التيليغرام' : '⚠️ Archive channel ID not configured in Telegram settings', 'warning');
      } else if (data.success) {
        setBulkTgResult({ uploaded: data.uploaded, skipped: data.skipped, failed: data.failed });
        const fresh = await fetch('/api/entries').then(r => r.json()).catch(() => null);
        if (fresh) setEntries(fresh);
        addSyncLog(lang === 'ar' ? `رفع جماعي: ${data.uploaded} ملف رُفع، ${data.failed} فشل، ${data.skipped} متجاوز` : `Bulk upload: ${data.uploaded} uploaded, ${data.failed} failed, ${data.skipped} skipped`);
      }
    } catch {
      showAlert(lang === 'ar' ? 'فشل الاتصال بالسيرفر' : 'Server connection failed', 'error');
    }
    setBulkTgUploading(false);
      }
    );
  };

  const handleMoveToArchive = async (entryToArchive: VisaEntry) => {
    setArchiveConfirmEntry(null);
    setArchiveUploading(true);
    try {
      const res = await fetch('/api/archive-upload', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ entryId: entryToArchive.id }),
      });
      const data = await res.json();
      if (data.success && data.entry) {
        setEntries(prev => prev.map(e => e.id === entryToArchive.id ? data.entry : e));
        const fresh = await fetch('/api/entries').then(r => r.json()).catch(() => null);
        if (fresh) localStorage.setItem('injaz_visa_entries', JSON.stringify(fresh));
        const uploadedMsg = data.uploaded
          ? `✅ ${lang === 'ar' ? 'تم رفع' : 'Uploaded'}: ${[data.uploaded.print ? (lang === 'ar' ? 'برنت' : 'print') : null, data.uploaded.visa ? (lang === 'ar' ? 'تاشيرة' : 'visa') : null].filter(Boolean).join(' + ') || (lang === 'ar' ? 'لا ملفات' : 'no files')}`
          : '';
        addSyncLog(`${lang === 'ar' ? `تم أرشفة العميل ${entryToArchive.applicantName}` : `Client ${entryToArchive.applicantName} archived`}. ${uploadedMsg}`);
      } else if (data.error === 'no_chat_id') {
        showAlert(lang === 'ar' ? '⚠️ لم يتم ضبط معرف قناة الأرشيف!\nاذهب إلى إعدادات التيليغرام وأضف "معرف قناة الأرشيف"' : '⚠️ Archive channel ID not configured!\nGo to Telegram settings and add the Archive Chat ID.', 'warning');
        setEntries(prev => prev.map(e => e.id === entryToArchive.id ? { ...e, archived: true, archivedAt: Date.now() } : e));
        addSyncLog(lang === 'ar' ? `تم أرشفة العميل ${entryToArchive.applicantName} (بدون رفع - لم يُضبط ID القناة).` : `Archived ${entryToArchive.applicantName} (no upload - channel ID not set).`);
      } else if (data.error === 'no_token') {
        showAlert(lang === 'ar' ? '⚠️ لم يتم ضبط توكن البوت!\nاذهب إلى إعدادات التيليغرام وأضف التوكن.' : '⚠️ Bot token not configured!\nGo to Telegram settings and add the token.', 'warning');
        setEntries(prev => prev.map(e => e.id === entryToArchive.id ? { ...e, archived: true, archivedAt: Date.now() } : e));
      } else if (data.error === 'upload_failed') {
        const errMsg = data.message || 'خطأ غير معروف';
        showAlert(lang === 'ar' ? `⚠️ فشل الرفع إلى تيليغرام:\n${errMsg}\n\nتأكد من:\n• أن البوت مضاف كمشرف في القناة\n• أن معرف القناة صحيح (مثال: -1001234567890)` : `⚠️ Telegram upload failed:\n${errMsg}\n\nMake sure:\n• Bot is added as admin to the channel\n• Channel ID is correct (e.g. -1001234567890)`, 'error');
        setEntries(prev => prev.map(e => e.id === entryToArchive.id ? { ...e, archived: true, archivedAt: Date.now() } : e));
        addSyncLog(lang === 'ar' ? `فشل رفع ملفات ${entryToArchive.applicantName}: ${errMsg}` : `Failed to upload ${entryToArchive.applicantName}: ${errMsg}`);
      } else {
        setEntries(prev => prev.map(e => e.id === entryToArchive.id ? { ...e, archived: true, archivedAt: Date.now() } : e));
        addSyncLog(lang === 'ar' ? `تم أرشفة العميل ${entryToArchive.applicantName}.` : `Client ${entryToArchive.applicantName} moved to archive.`);
      }
    } catch {
      setEntries(prev => prev.map(e => e.id === entryToArchive.id ? { ...e, archived: true, archivedAt: Date.now() } : e));
      addSyncLog(lang === 'ar' ? `تم أرشفة العميل ${entryToArchive.applicantName} (فشل الاتصال).` : `Archived ${entryToArchive.applicantName} (connection failed).`);
    }
    setArchiveUploading(false);
    setSelectedCustomerToEdit(null);
    setCustomerSubTab('none');
  };

  const exportToCSV = () => {
    if (!hasPermission('download_data')) {
      showAlert(lang === 'ar' ? 'عذراً، ليس لديك صلاحية "تنزيل بيانات العملاء والكشوفات" لتنزيل هذا التقرير!' : 'Sorry, you do not have permission to download client data and reports!', 'warning');
      return;
    }
    const headers = ['Passport', 'Application', 'Nationality', 'Arrival', 'Type', 'StatusText', 'Visa No', 'Telephone', 'Update'];
    const rows = entries.map(e => [
      e.passportNumber, e.applicationNumber, e.nationality, e.arrivalPoint, e.visaType, 
      e.statusText || e.status, e.visaNumber || 'N/A', e.phoneNumber || '', e.lastUpdate || ''
    ]);
    
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Report");
    XLSX.writeFile(wb, `visa_report_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const exportSectionToCSV = (section?: 'all' | 'today' | 'with_print' | 'without_print' | 'issued' | 'archived' | 'txstatus', officeId?: string, txStatusLabel?: string) => {
    const targetSection = section || allTabFilter;

    // Archived entries get their own pool
    let listToExport = targetSection === 'archived'
      ? entries.filter(e => e.archived)
      : entries.filter(e => !e.archived);

    // Filter by office if specified
    if (officeId) {
      listToExport = listToExport.filter(e => e.officeId === officeId);
    }

    if (targetSection === 'today') {
      listToExport = listToExport.filter(isToday);
    } else if (targetSection === 'with_print') {
      listToExport = listToExport.filter(entryHasPrint);
    } else if (targetSection === 'without_print') {
      listToExport = listToExport.filter(e => !entryHasPrint(e));
    } else if (targetSection === 'issued') {
      listToExport = listToExport.filter(e => (e.statusText?.includes('تم اصدار') || e.statusText?.includes('Issued')) && entryHasPrint(e));
    } else if (targetSection === 'txstatus' && txStatusLabel) {
      listToExport = listToExport.filter(e => e.txStatus === txStatusLabel);
    }

    if (listToExport.length === 0) {
      showAlert(lang === 'ar' ? 'لا توجد معاملات في هذا القسم لتصديرها!' : 'No transactions in this section to export!', 'warning');
      return;
    }

    // ── Human-readable section label ──────────────────────────────────
    const sectionLabel = lang === 'ar'
      ? (targetSection === 'all' ? 'كافة المعاملات'
        : targetSection === 'today' ? 'تحديثات اليوم'
        : targetSection === 'with_print' ? 'المعاملات لها طلب (برنت)'
        : targetSection === 'without_print' ? 'المعاملات بدون طلب (برنت)'
        : targetSection === 'archived' ? 'المعاملات المؤرشفة'
        : targetSection === 'txstatus' ? `حالة: ${txStatusLabel || ''}`
        : 'المعاملات الجاهزة (مؤشرة)')
      : (targetSection === 'all' ? 'All Transactions'
        : targetSection === 'today' ? "Today's Updates"
        : targetSection === 'with_print' ? 'Transactions with Print'
        : targetSection === 'without_print' ? 'Transactions without Print'
        : targetSection === 'archived' ? 'Archived Transactions'
        : targetSection === 'txstatus' ? `Status: ${txStatusLabel || ''}`
        : 'Ready / Issued Transactions');

    const officeName = officeId
      ? (offices.find(o => o.id === officeId)?.name || officeId)
      : (lang === 'ar' ? 'جميع المكاتب' : 'All Offices');

    const exportDate = new Date().toLocaleDateString(lang === 'ar' ? 'ar-u-ca-gregory' : 'en-GB');

    // ── Title block (rows 1-3) ────────────────────────────────────────
    const titleLine   = lang === 'ar' ? `تقرير: ${sectionLabel}` : `Report: ${sectionLabel}`;
    const officeLine  = lang === 'ar' ? `المكتب / النطاق: ${officeName}` : `Office / Scope: ${officeName}`;
    const dateLine    = lang === 'ar' ? `تاريخ التصدير: ${exportDate}` : `Export Date: ${exportDate}`;
    const countLine   = lang === 'ar' ? `عدد السجلات: ${listToExport.length}` : `Total Records: ${listToExport.length}`;

    // ── Full headers ──────────────────────────────────────────────────
    const headers = lang === 'ar' ? [
      'م',
      'الاسم الكامل',
      'الاسم بالإنجليزية',
      'رقم الجواز',
      'نوع الجواز',
      'تاريخ انتهاء الجواز',
      'رقم الطلب',
      'رقم التأشيرة / المستند',
      'رقم الوثيقة',
      'معرف مخصص',
      'الجنسية',
      'الجنسية الحالية',
      'جهة القدوم',
      'نوع التأشيرة',
      'المهنة',
      'الغرض من الزيارة',
      'تاريخ الميلاد',
      'مكان الميلاد',
      'الجنس',
      'رقم الهاتف',
      'السفارة',
      'عدد الدخولات',
      'اسم مقدم الطلب',
      'تاريخ التقديم',
      'المكتب / الوكيل',
      'الحالة',
      'حالة الكرت الصحي',
      'حالة المعاملة',
      'نوع حالة المعاملة',
      'تاريخ الحالة',
      'ملاحظة الحالة',
      'تفاصيل البيانات (البرنت)',
      'آخر تحديث',
      'تاريخ التحديث',
      'مؤرشف',
      'تاريخ الأرشفة',
    ] : [
      '#',
      'Full Name',
      'Name (English)',
      'Passport Number',
      'Passport Type',
      'Passport Expiry',
      'Application Number',
      'Visa / Doc Number',
      'Document Number',
      'Custom ID',
      'Nationality',
      'Current Nationality',
      'Arrival Point',
      'Visa Type',
      'Profession',
      'Purpose',
      'Birth Date',
      'Birth Place',
      'Gender',
      'Phone Number',
      'Embassy',
      'Entries Count',
      'Requester Name',
      'Application Date',
      'Office / Agent',
      'Status',
      'Health Cert Status',
      'Transaction Status',
      'Status Type',
      'Status Date',
      'Status Note',
      'Details (Print)',
      'Last Update',
      'Update Date/Time',
      'Archived',
      'Archived At',
    ];

    const fmtUnix = (unix?: number) => unix ? new Date(unix).toLocaleString(lang === 'ar' ? 'ar-u-ca-gregory' : 'en-GB') : '---';
    const dash = (v?: string | null) => v && v.trim() ? v.trim() : '---';

    const rows = listToExport.map((e, i) => [
      i + 1,
      dash(e.applicantName),
      dash(e.nameEnglish),
      dash(e.passportNumber),
      dash(e.passportType),
      dash(e.passportExpiry),
      dash(e.applicationNumber),
      dash(e.visaNumber),
      dash(e.documentNumber),
      dash(e.customId),
      dash(e.nationality),
      dash(e.currentNationality),
      dash(e.arrivalPoint),
      dash(e.visaType),
      dash(e.profession),
      dash(e.purpose),
      dash(e.birthDate),
      dash(e.birthPlace),
      dash(e.gender),
      dash(e.phoneNumber),
      dash(e.embassy),
      dash(e.entriesCount),
      dash(e.requesterName),
      dash(e.applicationDate),
      offices.find(o => o.id === e.officeId)?.name || (lang === 'ar' ? 'عام' : 'General'),
      dash(e.statusText || e.status),
      dash(e.healthCertStatus),
      dash(e.txStatus),
      dash(e.txStatusType),
      dash(e.txStatusDate),
      dash(e.txStatusNote),
      dash(e.applicantData),
      dash(e.lastUpdate),
      fmtUnix(e.updatedAtUnix),
      e.archived ? (lang === 'ar' ? 'نعم' : 'Yes') : (lang === 'ar' ? 'لا' : 'No'),
      fmtUnix(e.archivedAt),
    ]);

    // ── Auto column widths ────────────────────────────────────────────
    const allRows = [headers, ...rows];
    const colWidths = headers.map((_, ci) => ({
      wch: Math.min(
        60,
        Math.max(10, ...allRows.map(r => String(r[ci] ?? '').length)) + 3
      )
    }));

    // ── Build worksheet with title block then data ────────────────────
    const sheetData = [
      [titleLine],
      [officeLine],
      [dateLine],
      [countLine],
      [],          // blank separator row
      headers,
      ...rows,
    ];

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws['!cols'] = colWidths;

    // Merge title cells across all columns so they span the full width
    const lastCol = headers.length - 1;
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: lastCol } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: lastCol } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: lastCol } },
      { s: { r: 3, c: 0 }, e: { r: 3, c: lastCol } },
    ];

    // ── Sheet name (31 char limit in Excel) ──────────────────────────
    const sheetName = sectionLabel.replace(/[:\\/\[\]*?]/g, '').slice(0, 31);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    // ── File name ─────────────────────────────────────────────────────
    const sectionSlug = targetSection === 'all' ? (lang === 'ar' ? 'كافة_المعاملات' : 'all')
                       : targetSection === 'today' ? (lang === 'ar' ? 'تحديثات_اليوم' : 'today')
                       : targetSection === 'with_print' ? (lang === 'ar' ? 'لها_طلب' : 'with_print')
                       : targetSection === 'without_print' ? (lang === 'ar' ? 'بدون_طلب' : 'without_print')
                       : targetSection === 'archived' ? (lang === 'ar' ? 'المؤرشفة' : 'archived')
                       : targetSection === 'txstatus' ? (txStatusLabel?.replace(/\s+/g, '_') || 'custom')
                       : (lang === 'ar' ? 'الجاهزة' : 'ready');

    const officeSlug = officeId
      ? `_${offices.find(o => o.id === officeId)?.name?.replace(/\s+/g, '_') || officeId}`
      : (lang === 'ar' ? '_عام' : '_all_offices');

    XLSX.writeFile(wb, `${sectionSlug}${officeSlug}_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const exportSectionToPDF = async (section?: 'all' | 'today' | 'with_print' | 'without_print' | 'issued' | 'archived' | 'txstatus', officeId?: string, txStatusLabel?: string) => {
    const targetSection = section || 'all';

    let listToExport = targetSection === 'archived'
      ? entries.filter(e => e.archived)
      : entries.filter(e => !e.archived);

    if (officeId) listToExport = listToExport.filter(e => e.officeId === officeId);

    if (targetSection === 'today') listToExport = listToExport.filter(isToday);
    else if (targetSection === 'with_print') listToExport = listToExport.filter(entryHasPrint);
    else if (targetSection === 'without_print') listToExport = listToExport.filter(e => !entryHasPrint(e));
    else if (targetSection === 'issued') listToExport = listToExport.filter(e => (e.statusText?.includes('تم اصدار') || e.statusText?.includes('Issued')) && entryHasPrint(e));
    else if (targetSection === 'txstatus' && txStatusLabel) listToExport = listToExport.filter(e => e.txStatus === txStatusLabel);

    if (listToExport.length === 0) {
      showAlert(lang === 'ar' ? 'لا توجد معاملات في هذا القسم لتصديرها!' : 'No transactions in this section to export!', 'warning');
      return;
    }

    const sectionLabel = lang === 'ar'
      ? (targetSection === 'all' ? 'كافة المعاملات'
        : targetSection === 'today' ? 'تحديثات اليوم'
        : targetSection === 'with_print' ? 'المعاملات لها طلب (برنت)'
        : targetSection === 'without_print' ? 'المعاملات بدون طلب'
        : targetSection === 'archived' ? 'المعاملات المؤرشفة'
        : targetSection === 'txstatus' ? `حالة: ${txStatusLabel || ''}`
        : 'المعاملات الجاهزة')
      : (targetSection === 'all' ? 'All Transactions'
        : targetSection === 'today' ? "Today's Updates"
        : targetSection === 'with_print' ? 'Transactions with Print'
        : targetSection === 'without_print' ? 'Transactions without Print'
        : targetSection === 'archived' ? 'Archived Transactions'
        : targetSection === 'txstatus' ? `Status: ${txStatusLabel || ''}`
        : 'Ready / Issued Transactions');

    const officeName = officeId
      ? (offices.find(o => o.id === officeId)?.name || officeId)
      : (lang === 'ar' ? 'جميع المكاتب' : 'All Offices');

    const exportDate = new Date().toLocaleDateString(lang === 'ar' ? 'ar-u-ca-gregory' : 'en-GB');
    const dash = (v?: string | null) => v && v.trim() ? v.trim() : '---';
    const fmtUnix = (unix?: number) => unix ? new Date(unix).toLocaleString(lang === 'ar' ? 'ar-u-ca-gregory' : 'en-GB') : '---';

    const isRtl = lang === 'ar';
    const dir = isRtl ? 'rtl' : 'ltr';

    const columns = isRtl ? [
      { key: 'num',            label: 'م' },
      { key: 'applicantName',  label: 'الاسم الكامل' },
      { key: 'nameEnglish',    label: 'الاسم بالإنجليزية' },
      { key: 'passportNumber', label: 'رقم الجواز' },
      { key: 'passportType',   label: 'نوع الجواز' },
      { key: 'passportExpiry', label: 'انتهاء الجواز' },
      { key: 'applicationNumber', label: 'رقم الطلب' },
      { key: 'visaNumber',     label: 'رقم التأشيرة' },
      { key: 'nationality',    label: 'الجنسية' },
      { key: 'arrivalPoint',   label: 'جهة القدوم' },
      { key: 'visaType',       label: 'نوع التأشيرة' },
      { key: 'profession',     label: 'المهنة' },
      { key: 'gender',         label: 'الجنس' },
      { key: 'birthDate',      label: 'تاريخ الميلاد' },
      { key: 'phoneNumber',    label: 'رقم الهاتف' },
      { key: 'office',         label: 'المكتب / الوكيل' },
      { key: 'statusText',     label: 'الحالة' },
      { key: 'healthCertStatus', label: 'حالة الكرت الصحي' },
      { key: 'txStatus',       label: 'حالة المعاملة' },
      { key: 'txStatusDate',   label: 'تاريخ الحالة' },
      { key: 'txStatusNote',   label: 'الملاحظة' },
      { key: 'applicantData',  label: 'تفاصيل البرنت' },
      { key: 'lastUpdate',     label: 'آخر تحديث' },
      { key: 'archived',       label: 'مؤرشف' },
    ] : [
      { key: 'num',            label: '#' },
      { key: 'applicantName',  label: 'Full Name' },
      { key: 'nameEnglish',    label: 'Name (EN)' },
      { key: 'passportNumber', label: 'Passport No.' },
      { key: 'passportType',   label: 'Passport Type' },
      { key: 'passportExpiry', label: 'Expiry' },
      { key: 'applicationNumber', label: 'App. No.' },
      { key: 'visaNumber',     label: 'Visa No.' },
      { key: 'nationality',    label: 'Nationality' },
      { key: 'arrivalPoint',   label: 'Arrival Point' },
      { key: 'visaType',       label: 'Visa Type' },
      { key: 'profession',     label: 'Profession' },
      { key: 'gender',         label: 'Gender' },
      { key: 'birthDate',      label: 'Birth Date' },
      { key: 'phoneNumber',    label: 'Phone' },
      { key: 'office',         label: 'Office / Agent' },
      { key: 'statusText',     label: 'Status' },
      { key: 'healthCertStatus', label: 'Health Cert' },
      { key: 'txStatus',       label: 'Tx Status' },
      { key: 'txStatusDate',   label: 'Status Date' },
      { key: 'txStatusNote',   label: 'Note' },
      { key: 'applicantData',  label: 'Print Details' },
      { key: 'lastUpdate',     label: 'Last Update' },
      { key: 'archived',       label: 'Archived' },
    ];

    const getValue = (e: VisaEntry, key: string, i: number): string => {
      if (key === 'num') return String(i + 1);
      if (key === 'office') return offices.find(o => o.id === e.officeId)?.name || (isRtl ? 'عام' : 'General');
      if (key === 'archived') return e.archived ? (isRtl ? 'نعم' : 'Yes') : (isRtl ? 'لا' : 'No');
      if (key === 'statusText') return dash(e.statusText || e.status);
      const val = (e as Record<string, unknown>)[key];
      return dash(typeof val === 'string' ? val : undefined);
    };

    const tableRows = listToExport.map((e, i) =>
      `<tr>${columns.map(c => `<td>${getValue(e, c.key, i)}</td>`).join('')}</tr>`
    ).join('');

    const html = `<!DOCTYPE html>
<html dir="${dir}" lang="${lang}">
<head>
<meta charset="UTF-8"/>
<title>${sectionLabel}</title>
<style>
  @page { size: A3 landscape; margin: 12mm; }
  * { box-sizing: border-box; }
  body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 9pt; color: #1e293b; direction: ${dir}; margin: 0; padding: 0; }
  .header { padding: 10px 0 14px; border-bottom: 2px solid #e2e8f0; margin-bottom: 12px; }
  .header h1 { font-size: 15pt; font-weight: 900; margin: 0 0 4px; color: #0f172a; }
  .header p  { font-size: 8.5pt; margin: 2px 0; color: #475569; }
  .badge { display: inline-block; background: #f1f5f9; border: 1px solid #e2e8f0; border-radius: 6px; padding: 2px 8px; font-size: 8pt; font-weight: 700; color: #334155; margin-top: 4px; }
  table { width: 100%; border-collapse: collapse; font-size: 7.5pt; }
  thead th { background: #1e293b; color: #fff; padding: 5px 6px; text-align: ${isRtl ? 'right' : 'left'}; font-weight: 700; white-space: nowrap; border: 1px solid #334155; }
  tbody tr:nth-child(even) { background: #f8fafc; }
  tbody tr:hover { background: #eff6ff; }
  tbody td { padding: 4px 6px; border: 1px solid #e2e8f0; vertical-align: top; word-break: break-word; max-width: 120px; }
  @media print { button { display: none !important; } }
</style>
</head>
<body>
  <div class="header">
    <h1>${isRtl ? 'تقرير:' : 'Report:'} ${sectionLabel}</h1>
    <p>${isRtl ? 'المكتب / النطاق:' : 'Office / Scope:'} <strong>${officeName}</strong></p>
    <p>${isRtl ? 'تاريخ التصدير:' : 'Export Date:'} ${exportDate} &nbsp;|&nbsp; ${isRtl ? 'عدد السجلات:' : 'Total Records:'} <strong>${listToExport.length}</strong></p>
  </div>
  <table>
    <thead>
      <tr>${columns.map(c => `<th>${c.label}</th>`).join('')}</tr>
    </thead>
    <tbody>${tableRows}</tbody>
  </table>
  <script>window.onload = () => { window.print(); }<\/script>
</body>
</html>`;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const win = window.open(url, '_blank');
    if (win) win.focus();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  };

  return (
    <div className="min-h-screen flex flex-col bg-[#F8F9FA] dark:bg-[#0b111e] text-slate-900 dark:text-slate-100 transition-colors duration-300" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
      {/* Dynamic customizable theme style injector */}
      {siteIdentity.themePreset === 'custom' && (
        <style dangerouslySetInnerHTML={{ __html: `
          .theme-custom-gradient {
            background-image: linear-gradient(to right, ${siteIdentity.primaryColor || '#10b981'}, ${siteIdentity.secondaryColor || '#0ea5e9'}) !important;
          }
          .theme-custom-bgGrad {
            background-image: linear-gradient(to bottom, ${darkenColor(siteIdentity.primaryColor || '#10b981', 0.25)}, #0f172a) !important;
          }
          .theme-custom-bgSolo {
            background-color: ${siteIdentity.primaryColor || '#10b981'} !important;
          }
          .theme-custom-hoverBg:hover {
            background-color: ${darkenColor(siteIdentity.primaryColor || '#10b981', 0.12)} !important;
          }
          .theme-custom-textSolo {
            color: ${siteIdentity.primaryColor || '#10b981'} !important;
          }
          .theme-custom-shadowColor {
            box-shadow: 0 10px 15px -3px ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.2)}, 0 4px 6px -4px ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.2)} !important;
          }
          .theme-custom-lightBg {
            background-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.06)} !important;
          }
          .theme-custom-lightBorder {
            border-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.15)} !important;
          }
          .theme-custom-radarBg {
            background-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.06)} !important;
          }
          .theme-custom-radarPulse {
            background-color: ${siteIdentity.primaryColor || '#10b981'} !important;
          }
          .theme-custom-radarPulseBack {
            background-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.18)} !important;
          }
          .theme-custom-radarBorder {
            border-top-color: ${siteIdentity.primaryColor || '#10b981'} !important;
          }
          .theme-custom-badgeArPhone {
            color: ${darkenColor(siteIdentity.primaryColor || '#10b981', 0.2)} !important;
            background-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.12)} !important;
          }
          .theme-custom-badgeDarkText {
            color: ${darkenColor(siteIdentity.primaryColor || '#10b981', 0.22)} !important;
            background-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.05)} !important;
            border-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.15)} !important;
          }
          .theme-custom-ringColor:focus {
            --tw-ring-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.2)} !important;
            box-shadow: var(--tw-ring-inset) 0 0 0 calc(4px + var(--tw-ring-offset-width)) var(--tw-ring-color) !important;
          }
          .theme-custom-textColorDark {
            color: ${darkenColor(siteIdentity.primaryColor || '#10b981', 0.35)} !important;
            border-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.15)} !important;
            background-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.04)} !important;
          }
          .theme-custom-textColorDarker {
            color: ${darkenColor(siteIdentity.primaryColor || '#10b981', 0.5)} !important;
            border-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.18)} !important;
            background-color: ${hexToRgba(siteIdentity.primaryColor || '#10b981', 0.03)} !important;
          }
        ` }} />
      )}


      {viewMode === 'login' ? (
        <AdminLoginPage
          lang={lang}
          adminUsernameInput={adminUsernameInput}
          setAdminUsernameInput={setAdminUsernameInput}
          adminPinInput={adminPinInput}
          setAdminPinInput={setAdminPinInput}
          adminPinError={adminPinError}
          setAdminPinError={setAdminPinError}
          handleLoginSubmit={handleLoginSubmit}
          siteIdentity={siteIdentity}
        />
      ) : viewMode === 'public' ? (
        /* ==================== FRONT-END PUBLIC VISITOR PORTAL ==================== */
        <div className="flex-1 flex flex-col min-h-screen">
          {/* Public Navbar Header */}
          <nav className="h-20 ios-glass-nav px-3 sm:px-6 md:px-12 flex items-center justify-between sticky top-0 z-50 shadow-sm">
            <div className="flex items-center gap-2 sm:gap-4">
              {siteIdentity.logo ? (
                <div className="w-10 h-10 sm:w-12 sm:h-12 rounded-xl sm:rounded-2xl overflow-hidden shadow-md shrink-0 border border-white bg-white/80 flex items-center justify-center">
                  <img src={siteIdentity.logo} alt="Logo" className="w-full h-full object-contain" referrerPolicy="no-referrer" />
                </div>
              ) : (
                <div className={`w-10 h-10 sm:w-12 sm:h-12 bg-gradient-to-tr ${theme.gradient} rounded-xl sm:rounded-2xl flex items-center justify-center text-white shadow-lg ${theme.shadowColor} shrink-0`}>
                  <ShieldCheck className="w-5 h-5 sm:w-7 sm:h-7" />
                </div>
              )}
              <div className="min-w-0">
                <h1 className="text-xs sm:text-base md:text-xl font-black text-slate-900 tracking-tight leading-tight truncate max-w-[150px] xs:max-w-[200px] sm:max-w-none">{t.officeName}</h1>
                <p className={`text-[8px] sm:text-[10px] uppercase font-extrabold mt-0.5 sm:mt-1.5 tracking-wider truncate ${theme.textSolo}`}>{lang === 'ar' ? siteIdentity.heroBadgeAr : siteIdentity.heroBadgeEn}</p>
              </div>
            </div>
          </nav>

          {/* Hero Branding Section */}
          <section className="py-12 sm:py-20 px-4 sm:px-6 relative overflow-hidden">
            <div className="max-w-4xl mx-auto text-center space-y-4 sm:space-y-6">
              <h2 className="text-lg sm:text-2xl md:text-3xl font-black text-slate-900 tracking-tight leading-tight px-1">
                {lang === 'ar' ? siteIdentity.heroTitleAr : siteIdentity.heroTitleEn}
              </h2>
              <p className="text-slate-600 text-xs sm:text-sm md:text-base font-medium max-w-2xl mx-auto leading-relaxed px-2">
                {t.officeDescription}
              </p>
            </div>
          </section>

          {/* Main Visual Search Space */}
          <main className="flex-1 max-w-5xl w-full mx-auto px-4 sm:px-6 py-4 sm:py-12 space-y-6 sm:space-y-12">
            
            {/* Search Card Box */}
            {siteIdentity.searchEnabled !== false ? (
              <section className="ios-glass-card p-5 sm:p-10 rounded-3xl sm:rounded-[2.5rem] relative -translate-y-6 sm:-translate-y-16 z-20">
                <div className="max-w-3xl mx-auto space-y-6 sm:space-y-8">
                  <div className="text-center space-y-2">
                    <h3 className="text-lg sm:text-xl font-bold text-slate-900">{t.searchInquiry}</h3>
                    <p className="text-xs text-slate-500 sm:text-slate-400 font-medium px-2 leading-relaxed">
                      {lang === 'ar' ? siteIdentity.searchDescAr : siteIdentity.searchDescEn}
                    </p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {/* Passport Input */}
                    <div className="space-y-2">
                      <label className="text-xs font-extrabold text-slate-500 sm:text-slate-400 tracking-wider uppercase flex items-center gap-2">
                        <FileText size={14} className={theme.textSolo} />
                        {t.passport}
                      </label>
                      <div className="relative">
                        <input 
                          type="text"
                          placeholder=""
                          className={`w-full ios-glass-input rounded-xl sm:rounded-2xl py-3 sm:py-4 px-4 sm:px-5 text-base sm:text-sm font-bold font-mono text-slate-900 focus:outline-none focus:ring-4 ${theme.ringColor} transition-all placeholder:text-slate-400 uppercase`}
                          value={publicPassport}
                          onChange={(e) => setPublicPassport(e.target.value.toUpperCase())}
                        />
                      </div>
                    </div>

                    {/* Phone Input */}
                    <div className="space-y-2">
                       <label className="text-xs font-extrabold text-slate-500 sm:text-slate-400 tracking-wider uppercase flex items-center gap-2">
                        <Phone size={14} className={theme.textSolo} />
                        {t.phone}
                      </label>
                      <div className="relative">
                        <input 
                          type="text"
                          placeholder=""
                          className={`w-full ios-glass-input rounded-xl sm:rounded-2xl py-3 sm:py-4 px-4 sm:px-5 text-base sm:text-sm font-bold text-slate-900 focus:outline-none focus:ring-4 ${theme.ringColor} transition-all placeholder:text-slate-400`}
                          value={publicPhone}
                          onChange={(e) => setPublicPhone(e.target.value)}
                        />
                      </div>
                    </div>
                  </div>

                  <div className="flex justify-center pt-2">
                    <button
                      onClick={handlePublicSearch}
                      disabled={isPublicSearching || !publicPassport.trim() || !publicPhone.trim()}
                      className={`w-full sm:w-80 bg-gradient-to-r ${theme.gradient} text-white font-black py-3.5 sm:py-4 px-6 sm:px-8 rounded-xl sm:rounded-2xl text-sm sm:text-base shadow-xl ${theme.shadowColor} hover:scale-[1.02] active:scale-[0.98] transition-all disabled:opacity-40 disabled:pointer-events-none flex items-center justify-center gap-3 cursor-pointer`}
                    >
                      {isPublicSearching ? (
                        <>
                          <RefreshCw size={18} className="animate-spin" />
                          <span>{lang === 'ar' ? 'جاري الفحص والمطابقة اللحظية...' : 'Matching...'}</span>
                        </>
                      ) : (
                        <>
                          <Search size={18} />
                          <span>{t.searchInquiry}</span>
                        </>
                      )}
                    </button>
                  </div>
                </div>
              </section>
            ) : (
              <section className="bg-amber-50 border-2 border-amber-200/80 p-6 sm:p-12 rounded-3xl sm:rounded-[2.5rem] shadow-xl relative -translate-y-6 sm:-translate-y-16 z-20 text-center max-w-4xl mx-auto space-y-4">
                <div className="w-16 h-16 bg-amber-100 rounded-full flex items-center justify-center mx-auto text-amber-600 shadow-inner">
                  <AlertTriangle size={32} />
                </div>
                <div className="max-w-2xl mx-auto space-y-3">
                  <h3 className="text-lg sm:text-xl font-bold text-slate-955 leading-tight">
                    {lang === 'ar' ? 'الاستعلام الإلكتروني متوقف مؤقتاً' : 'Search Portal Temporarily Offline'}
                  </h3>
                  <p className="text-xs sm:text-sm font-semibold text-slate-600 leading-relaxed px-2 sm:px-6">
                    {lang === 'ar' ? siteIdentity.searchDisabledMessageAr : siteIdentity.searchDisabledMessageEn}
                  </p>
                </div>
              </section>
            )}

            {/* Results Pane */}
            <div className="relative -translate-y-4 sm:-translate-y-12 transition-all duration-300">
              {isPublicSearching ? (
                /* Dynamic radar / finding feedback */
                <div className="bg-white rounded-3xl sm:rounded-[2.5rem] border border-slate-200 p-8 sm:p-16 text-center space-y-6 shadow-md animate-pulse">
                  <div className="relative w-20 h-20 mx-auto">
                    <div className="absolute inset-0 rounded-full border-4 border-slate-100 animate-spin" />
                    <div className={`absolute inset-2 ${theme.radarBg} rounded-full flex items-center justify-center ${theme.textSolo}`}>
                      <ShieldCheck size={28} />
                    </div>
                  </div>
                  <div className="space-y-2">
                    <h4 className="text-lg font-bold text-slate-900">{lang === 'ar' ? 'جاري البحث فائق السرعة...' : 'Hyper-speed search in progress...'}</h4>
                    <p className="text-xs text-slate-400 font-medium">{lang === 'ar' ? 'يتم مطابقة البيانات وفحص سجلات النظام وحالة الجواز' : 'Matching with high-precision database indexes'}</p>
                  </div>
                </div>
              ) : publicHasSearched ? (
                publicSearchResult ? (
                  /* FOUND: Premium Boarding Ticket Card */
                  <motion.div 
                    initial={{ opacity: 0, y: 15 }} 
                    animate={{ opacity: 1, y: 0 }}
                    className="bg-white rounded-3xl sm:rounded-[2.5rem] border border-slate-200 shadow-2xl overflow-hidden"
                  >
                    {/* Upper ticket design header */}
                    <div className={`bg-gradient-to-r ${theme.gradient} text-white px-5 sm:px-8 py-5 sm:py-6 flex flex-col sm:flex-row sm:items-center justify-between gap-4`}>
                      <div className="space-y-1">
                        <span className="px-2.5 py-0.5 bg-emerald-500/20 text-emerald-100 text-[9px] sm:text-[10px] font-black uppercase tracking-widest rounded-md inline-block">
                          {lang === 'ar' ? 'تم العثور على المعاملة' : 'TRANSACTION FOUND'}
                        </span>
                        <h4 className="text-base sm:text-lg md:text-2xl font-black">{t.searchResultTitle}</h4>
                      </div>
                      <div className="flex items-center gap-3 self-start sm:self-center shrink-0">
                        <div className="text-right">
                          <p className="text-[9px] sm:text-[10px] font-bold uppercase tracking-wider opacity-60">{lang === 'ar' ? 'التحديث الأخير' : 'LAST SYNC'}</p>
                          <p className="text-xs sm:text-sm font-bold font-mono">{publicSearchResult.lastUpdate || 'Live Now'}</p>
                        </div>
                        <div className="w-8 h-8 sm:w-10 sm:h-10 bg-white/10 rounded-lg sm:rounded-xl flex items-center justify-center text-white shrink-0">
                          <Clock className="w-4 h-4 sm:w-5 sm:h-5" />
                        </div>
                      </div>
                    </div>

                    <div className="p-5 sm:p-10 space-y-6 sm:space-y-8">
                      {/* Timeline status checker */}
                      <div className="bg-slate-50 p-4 sm:p-6 rounded-2xl sm:rounded-3xl border border-slate-200/50 space-y-5 sm:space-y-6">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                          <span className="text-[10px] sm:text-xs font-black text-slate-400 tracking-widest uppercase">{lang === 'ar' ? 'خطوات المعالجة اللحظية' : 'LIVE STATUS PIPELINE'}</span>
                          <span className={`px-2.5 py-0.5 text-xs font-black rounded-full border self-start sm:self-auto ${
                            (publicSearchResult.statusText?.includes('اصدار') || publicSearchResult.statusText?.includes('Issued')) && entryHasPrint(publicSearchResult) 
                              ? 'bg-emerald-50 text-emerald-700 border-emerald-200' 
                              : 'bg-amber-50 text-amber-700 border-amber-200'
                          }`}>
                            {publicSearchResult.statusText || 'تحت الإجراء'}
                          </span>
                        </div>

                        {/* Interactive custom tracking dots */}
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 md:gap-2 relative">
                          {/* Step 1 */}
                          <div className="flex items-center gap-3">
                            <div className={`w-7 h-7 sm:w-8 sm:h-8 rounded-full ${theme.lightBg} flex items-center justify-center ${theme.textSolo} font-bold text-xs shrink-0`}>✓</div>
                            <div>
                              <p className="text-xs sm:text-sm font-bold text-slate-900">{lang === 'ar' ? 'تم استلام وتجهيز المعاملة' : 'Application Received'}</p>
                              <p className="text-[10px] sm:text-xs text-slate-500 font-semibold">{lang === 'ar' ? 'مكتمل ومرحل بنجاح' : 'Ready & Processed'}</p>
                            </div>
                          </div>
                          {/* Step 2 */}
                          <div className="flex items-center gap-3">
                            <div className={`w-7 h-7 sm:w-8 sm:h-8 rounded-full ${theme.lightBg} flex items-center justify-center ${theme.textSolo} font-bold text-xs shrink-0`}>✓</div>
                            <div>
                              <p className="text-xs sm:text-sm font-bold text-slate-900">{lang === 'ar' ? 'سداد الرسوم والربط مرئياً' : 'MOFA Portal Direct Link'}</p>
                              <p className="text-[10px] sm:text-xs text-slate-500 font-semibold">{lang === 'ar' ? 'أرقام المستندات والمقاصة مكتملة' : 'Doc Link Active'}</p>
                            </div>
                          </div>
                          {/* Step 3 */}
                          <div className="flex items-center gap-3">
                            <div className={`w-7 h-7 sm:w-8 sm:h-8 rounded-full flex items-center justify-center font-bold text-xs shrink-0 ${
                              (publicSearchResult.statusText?.includes('اصدار') || publicSearchResult.statusText?.includes('Issued')) && entryHasPrint(publicSearchResult)
                                ? 'bg-emerald-100 text-emerald-600'
                                : 'bg-amber-100 text-amber-600 animate-pulse'
                            }`}>
                              {(publicSearchResult.statusText?.includes('اصدار') || publicSearchResult.statusText?.includes('Issued')) && entryHasPrint(publicSearchResult) ? '✓' : '●'}
                            </div>
                            <div>
                              <p className="text-xs sm:text-sm font-bold text-slate-900">{lang === 'ar' ? 'الاعتماد والطباعة النهائية' : 'MOFA Decisive Action'}</p>
                              <p className={`text-[10px] sm:text-xs font-black uppercase ${theme.textSolo}`}>
                                {(publicSearchResult.statusText?.includes('اصدار') || publicSearchResult.statusText?.includes('Issued')) && entryHasPrint(publicSearchResult) 
                                  ? (lang === 'ar' ? 'جاهز للتسليم' : 'READY TO PICKUP') 
                                  : (lang === 'ar' ? 'جاري الفحص المتقدم' : 'UNDER PROCESSING')
                                }
                              </p>
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Descriptive metadata list */}
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 sm:gap-6 p-1 border-t border-slate-100 pt-6 sm:pt-8">
                        <div>
                          <p className="text-[11px] sm:text-xs font-extrabold text-slate-400 uppercase tracking-wider mb-1">{t.name}</p>
                          <p className="text-sm sm:text-base font-black text-slate-900">{publicSearchResult.applicantName || '---'}</p>
                        </div>
                        <div>
                          <p className="text-[11px] sm:text-xs font-extrabold text-slate-400 uppercase tracking-wider mb-1">{t.passport}</p>
                          <p className="text-sm sm:text-base font-bold font-mono text-slate-900">{publicSearchResult.passportNumber}</p>
                        </div>
                        <div>
                          <p className="text-[11px] sm:text-xs font-extrabold text-slate-400 uppercase tracking-wider mb-1">{lang === 'ar' ? 'جهة السفر ونوع التأشيرة' : 'Type & Arrival'}</p>
                          <p className="text-xs sm:text-sm font-bold text-slate-600">{publicSearchResult.visaType} - القدوم {publicSearchResult.arrivalPoint}{publicSearchResult.profession ? ` (${lang === 'ar' ? 'المهنة' : 'Job'}: ${publicSearchResult.profession})` : ''}</p>
                        </div>
                        <div>
                          <p className="text-[11px] sm:text-xs font-extrabold text-slate-400 uppercase tracking-wider mb-1">{t.visaNo}</p>
                          <span className={`font-mono text-xs sm:text-sm font-black ${theme.lightBg} px-3 py-1 rounded-xl ${theme.textColorDark} border ${theme.lightBorder} inline-block mt-0.5`}>
                            {publicSearchResult.visaNumber || '---'}
                          </span>
                        </div>
                        <div>
                          <p className="text-[11px] sm:text-xs font-extrabold text-slate-400 uppercase tracking-wider mb-1">{lang === 'ar' ? 'رقم المستند' : 'Document No.'}</p>
                          <span className="font-mono text-xs sm:text-sm font-black bg-amber-50 px-3 py-1 rounded-xl text-amber-900 border border-amber-100 inline-block mt-0.5">
                            {publicSearchResult.documentNumber || '---'}
                          </span>
                        </div>
                      </div>

                      {/* Official System details comment */}
                      <div className={`p-4 sm:p-5 ${theme.lightBg} rounded-xl sm:rounded-2xl border ${theme.lightBorder} ${theme.textColorDark} text-[11px] sm:text-xs font-medium space-y-1`}>
                        <p className={`font-bold ${theme.textColorDarker} text-xs sm:text-sm`}>{lang === 'ar' ? '📄 ملاحظة وإفادة إدارية مصلحية:' : '📄 Official Office Dispatch Note:'}</p>
                        <p className="leading-relaxed text-slate-700">{publicSearchResult.applicantData || (lang === 'ar' ? 'تم تحديث واستيراد المعاملة بنجاح، تفضل بمراجعة الفرع المعتمد لاستكمال ختم واستلام وثيقتك.' : 'Transaction processed and synced with database.')}</p>
                      </div>

                      {/* Whatsapp contact with dynamized custom text */}
                      <div className="flex flex-col sm:flex-row gap-3 pt-4 border-t border-slate-100">
                        <a 
                          href={`https://wa.me/967771234567?text=${encodeURIComponent(
                            lang === 'ar' 
                              ? `مرحباً مكتب إنجاز، أود المتابعة بخصوص المعاملة الخاصة بالمسافر: ${publicSearchResult.applicantName}، رقم الجواز: ${publicSearchResult.passportNumber}`
                              : `Hello Injaz Office, I'm checking back regarding applicant: ${publicSearchResult.applicantName}, Passport No: ${publicSearchResult.passportNumber}`
                          )}`}
                          target="_blank"
                          rel="noreferrer"
                          className={`flex-1 ${theme.bgSolo} text-white font-black py-3 sm:py-3.5 px-4 sm:px-6 rounded-xl sm:rounded-2xl text-center shadow-lg ${theme.shadowColor} ${theme.hoverBg} transition-all flex items-center justify-center gap-2 sm:gap-3 cursor-pointer text-xs sm:text-sm shrink-0`}
                        >
                          <Phone size={16} />
                          <span>{t.whatsappFollowup}</span>
                        </a>
                      </div>
                    </div>
                  </motion.div>
                ) : (
                  /* NOT FOUND ALERT */
                  <motion.div 
                    initial={{ opacity: 0, scale: 0.98 }}
                    animate={{ opacity: 1, scale: 1 }}
                    className="bg-white rounded-3xl sm:rounded-[2.5rem] border border-red-200/60 shadow-xl p-5 sm:p-12 text-center space-y-5 sm:space-y-6"
                  >
                    <div className="w-12 h-12 sm:w-16 sm:h-16 bg-red-50 text-red-600 rounded-xl sm:rounded-2xl flex items-center justify-center mx-auto border border-red-100 shrink-0">
                      <AlertCircle className="w-6 h-6 sm:w-8 sm:h-8" />
                    </div>
                    <div className="max-w-xl mx-auto space-y-2 sm:space-y-3">
                      <h4 className="text-base sm:text-lg md:text-xl font-bold text-slate-900">{t.notFoundTitle}</h4>
                      <p className="text-slate-500 text-xs sm:text-sm leading-relaxed">{t.notFoundDesc}</p>
                    </div>

                    <div className="pt-2 flex flex-col sm:flex-row items-center justify-center gap-2.5 sm:gap-3">
                      <a 
                        href={`https://wa.me/967771234567?text=${encodeURIComponent(
                          lang === 'ar'
                            ? `مرحباً مكتب إنجاز، قمت بالاستعلام عبر الجواز رقم [ ${publicPassport} ] ولم أعثر عليها، أرجو إفادتي بحالة المعاملة السجلية.`
                            : `Hello Injaz Office, I searched passport: [ ${publicPassport} ] and got no matching records. Could you please check with my customer profile.`
                        )}`}
                        target="_blank"
                        rel="noreferrer"
                        className={`${theme.bgSolo} text-white font-bold py-3 sm:py-3.5 px-4 sm:px-6 rounded-lg sm:rounded-xl ${theme.hoverBg} transition-all flex items-center gap-2 text-[11px] sm:text-xs shadow-md w-full sm:w-auto justify-center cursor-pointer shrink-0`}
                      >
                        <Phone size={14} />
                        <span>{lang === 'ar' ? 'تواصل مع الدعم الفني للمكتب' : 'Contact Support Representative'}</span>
                      </a>
                      <button 
                        onClick={() => {
                          setPublicPassport('');
                          setPublicPhone('');
                          setPublicHasSearched(false);
                        }}
                        className="bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold py-3 sm:py-3.5 px-4 sm:px-6 rounded-lg sm:rounded-xl transition-all text-[11px] sm:text-xs w-full sm:w-auto text-center cursor-pointer"
                      >
                        {lang === 'ar' ? 'مسح المدخلات والمحاولة مجدداً' : 'Clear & Try Again'}
                      </button>
                    </div>
                  </motion.div>
                )
              ) : null}
            </div>
          </main>

          {/* Luxury Public Footer */}
          <footer className="ios-glass-nav text-slate-600 dark:text-slate-400 py-10 sm:py-12 px-4 sm:px-6 border-t border-slate-200/40 dark:border-white/5 mt-auto">
            <div className="max-w-5xl mx-auto flex flex-col md:flex-row justify-between items-center gap-6">
              <div className="space-y-2 text-center md:text-right">
                <div className="flex items-center gap-3 justify-center md:justify-start">
                  {siteIdentity.logo ? (
                    <div className="w-8 h-8 rounded-lg overflow-hidden shrink-0 border border-slate-200 dark:border-white/10 bg-white/80 dark:bg-slate-900/40 flex items-center justify-center">
                      <img src={siteIdentity.logo} alt="Logo" className="w-full h-full object-contain" referrerPolicy="no-referrer" />
                    </div>
                  ) : (
                    <div className={`w-8 h-8 ${theme.bgSolo} rounded-lg flex items-center justify-center text-white shrink-0`}>
                      <ShieldCheck size={18} />
                    </div>
                  )}
                  <span className="font-black text-slate-900 dark:text-white text-sm sm:text-base">{t.officeName}</span>
                </div>
                {(lang === 'ar' ? siteIdentity.officeDescriptionAr : siteIdentity.officeDescriptionEn) && (
                  <p className="text-[11px] sm:text-xs max-w-sm md:max-w-md text-slate-500 dark:text-slate-400 leading-relaxed">{lang === 'ar' ? siteIdentity.officeDescriptionAr : siteIdentity.officeDescriptionEn}</p>
                )}
              </div>

              <div className="text-center md:text-left space-y-1.5 sm:space-y-2 shrink-0">
                <p className="font-extrabold text-slate-900 dark:text-white text-xs sm:text-sm">{lang === 'ar' ? 'تواصل بالفرع الرئيسي' : 'Regional Offices Contact'}</p>
                {(lang === 'ar' ? siteIdentity.hqAddressAr : siteIdentity.hqAddressEn) && (
                  <p className="text-[11px] sm:text-xs text-slate-500 dark:text-slate-400">{lang === 'ar' ? siteIdentity.hqAddressAr : siteIdentity.hqAddressEn}</p>
                )}
                {(siteIdentity.supportPhone || siteIdentity.supportEmail) && (
                  <p className="text-[11px] sm:text-xs text-slate-500 dark:text-slate-400">
                    {siteIdentity.supportPhone && (lang === 'ar' ? `هاتف الدعم: ${siteIdentity.supportPhone}` : `Support: ${siteIdentity.supportPhone}`)}
                    {siteIdentity.supportPhone && siteIdentity.supportEmail && ' | '}
                    {siteIdentity.supportEmail && (lang === 'ar' ? `البريد: ${siteIdentity.supportEmail}` : `Mail: ${siteIdentity.supportEmail}`)}
                  </p>
                )}
              </div>
            </div>
          </footer>
        </div>
      ) : (
        /* ==================== BACK-END OFFICE CONTROL PANEL (ADMIN) ==================== */
        <div className="flex flex-1 overflow-hidden min-h-screen">
          
          {/* Admin Navigation Sidebar Mobile Menu */}
          <AnimatePresence>
            {isSidebarOpen && (
              <>
                <motion.div 
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  exit={{ opacity: 0 }}
                  onClick={() => setIsSidebarOpen(false)}
                  className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[60] lg:hidden"
                />
                <motion.aside 
                  initial={{ x: lang === 'ar' ? 300 : -300 }}
                  animate={{ x: 0 }}
                  exit={{ x: lang === 'ar' ? 300 : -300 }}
                  transition={{ type: 'spring', damping: 25, stiffness: 200 }}
                  className="fixed top-0 bottom-0 z-[70] w-72 bg-white flex flex-col py-6 lg:hidden shadow-2xl"
                  style={{ [lang === 'ar' ? 'right' : 'left']: 0 }}
                >
                  <div className="px-6 flex justify-between items-center mb-6 gap-2">
                    <div className="flex items-center gap-2 min-w-0">
                      {siteIdentity.logo ? (
                        <div className="w-8 h-8 rounded-lg overflow-hidden border border-slate-200 bg-white flex items-center justify-center shrink-0 shadow-sm">
                          <img src={siteIdentity.logo} alt="Logo" className="w-full h-full object-contain" referrerPolicy="no-referrer" />
                        </div>
                      ) : (
                        <div className={`w-8 h-8 bg-gradient-to-tr ${theme.gradient} rounded-lg flex items-center justify-center text-white shrink-0 shadow-sm`}>
                          <ShieldCheck size={16} />
                        </div>
                      )}
                      <div className="min-w-0">
                        <div className="font-extrabold text-slate-900 text-[10px] truncate max-w-[120px]">{lang === 'ar' ? siteIdentity.officeNameAr : siteIdentity.officeNameEn}</div>
                        <div className="text-[7.5px] uppercase font-bold text-slate-400 mt-0.5 tracking-wider truncate">{t.adminPortal}</div>
                      </div>
                    </div>
                    <button onClick={() => setIsSidebarOpen(false)} className="text-slate-400 p-1 shrink-0"><X size={20} /></button>
                  </div>
                  <div className="px-6 mb-4 text-[10px] font-bold uppercase tracking-[0.2em] text-slate-400">Navigation</div>
                  <div className="flex flex-col gap-1 px-3">
                    {[
                      { id: 'dashboard', label: t.tab_dashboard, icon: <LayoutDashboard size={18} />, perm: 'any' },
                      { id: 'all', label: t.tab_all, icon: <Globe size={18} />, perm: 'edit_entries' },
                      { id: 'customers', label: t.tab_customers, icon: <Users size={18} />, color: 'text-violet-600', perm: 'edit_entries' },
                      { id: 'files', label: t.tab_files, icon: <FolderOpen size={18} />, color: 'text-amber-600', perm: 'edit_entries' },
                      { id: 'identity', label: t.tab_identity, icon: <Settings size={18} />, color: 'text-slate-600', perm: 'any' },
                      { id: 'security', label: t.tab_security, icon: <Lock size={18} />, color: 'text-rose-600', perm: 'admin' },
                      { id: 'excel_sync', label: t.tab_excel_sync, icon: <Database size={18} />, color: 'text-rose-500 font-bold', perm: 'excel_sync' },
                      { id: 'auto_check', label: t.tab_auto_check, icon: <Clock size={18} />, color: 'text-emerald-600', perm: 'admin' },
                    ].filter(item => item.perm === 'any' || hasPermission(item.perm)).map((item) => (
                      <button
                        key={item.id}
                        onClick={() => { 
                          setActiveTab(item.id as any);
                          setFilesSubPage(null); setFilesSearch('');
                          setIsSidebarOpen(false); 
                          if (item.id === 'all') {
                            setIsSectionSelected(false);
                          }
                        }}
                        className={`flex items-center gap-3 px-3 py-3 rounded-xl text-sm font-bold transition-all text-right ${
                          activeTab === item.id 
                            ? 'bg-slate-900 text-white shadow-lg shadow-slate-900/10' 
                            : 'text-slate-500 hover:bg-slate-50 hover:text-slate-900'
                        }`}
                      >
                        <span className={activeTab === item.id ? 'text-white' : item.color}>{item.icon}</span>
                        {item.label}
                      </button>
                    ))}
                  </div>

                  <div className="mt-auto px-4 pt-6 flex flex-col gap-2">
                    <button 
                      onClick={() => {
                        setViewMode('public');
                        setCurrentUser(null);
                        localStorage.removeItem('injaz_visa_current_user');
                      }}
                      className="w-full bg-slate-100 hover:bg-slate-200 text-slate-900 py-3 rounded-xl text-center text-xs font-bold flex items-center justify-center gap-2"
                    >
                      <ArrowRight size={14} />
                      <span>{t.backToPublic}</span>
                    </button>
                  </div>
                </motion.aside>
              </>
            )}
          </AnimatePresence>

          {/* Desktop Left Sidebar Panel */}
          <aside className="w-68 bg-white border-r border-slate-200 hidden lg:flex flex-col py-6">
            <div className="px-6 mb-6 flex items-center gap-3">
              {siteIdentity.logo ? (
                <div className="w-9 h-9 rounded-xl overflow-hidden border border-slate-200 bg-white flex items-center justify-center shrink-0 shadow-sm">
                  <img src={siteIdentity.logo} alt="Logo" className="w-full h-full object-contain" referrerPolicy="no-referrer" />
                </div>
              ) : (
                <div className={`w-9 h-9 bg-gradient-to-tr ${theme.gradient} rounded-xl flex items-center justify-center text-white shrink-0 shadow-sm`}>
                  <ShieldCheck size={18} />
                </div>
              )}
              <div className="min-w-0">
                <div className="font-extrabold text-slate-900 text-xs truncate max-w-[130px]">{lang === 'ar' ? siteIdentity.officeNameAr : siteIdentity.officeNameEn}</div>
                <div className="text-[8px] uppercase font-bold text-slate-400 mt-0.5 tracking-wider truncate">{t.adminPortal}</div>
              </div>
            </div>
            <div className="px-6 mb-4 text-[10px] font-bold uppercase tracking-[0.2em] text-slate-400">Navigation Menu</div>
            <div className="flex flex-col gap-1 px-3">
              {[
                { id: 'dashboard', label: t.tab_dashboard, icon: <LayoutDashboard size={18} />, perm: 'any' },
                { id: 'all', label: t.tab_all, icon: <Globe size={18} />, perm: 'edit_entries' },
                { id: 'customers', label: t.tab_customers, icon: <Users size={18} />, color: 'text-violet-600', perm: 'edit_entries' },
                { id: 'files', label: t.tab_files, icon: <FolderOpen size={18} />, color: 'text-amber-600', perm: 'edit_entries' },
                { id: 'identity', label: t.tab_identity, icon: <Settings size={18} />, color: 'text-slate-600', perm: 'any' },
                { id: 'security', label: t.tab_security, icon: <Lock size={18} />, color: 'text-rose-600', perm: 'admin' },
                { id: 'excel_sync', label: t.tab_excel_sync, icon: <Database size={18} />, color: 'text-rose-500 font-bold', perm: 'excel_sync' },
                { id: 'auto_check', label: t.tab_auto_check, icon: <Clock size={18} />, color: 'text-emerald-600', perm: 'admin' },
              ].filter(item => item.perm === 'any' || hasPermission(item.perm)).map((item) => (
                <button
                  key={item.id}
                  onClick={() => {
                    setActiveTab(item.id as any);
                    setFilesSubPage(null); setFilesSearch('');
                    if (item.id === 'all') {
                      setIsSectionSelected(false);
                    }
                  }}
                  className={`flex items-center gap-3 px-3 py-3 rounded-xl text-sm font-bold transition-all ${
                    activeTab === item.id 
                      ? 'bg-slate-900 text-white shadow-lg shadow-slate-900/10' 
                      : 'text-slate-500 hover:bg-slate-50 hover:text-slate-900'
                  }`}
                >
                  <span className={activeTab === item.id ? 'text-white' : item.color}>{item.icon}</span>
                  {item.label}
                </button>
              ))}
            </div>

            <div className="mt-auto px-4 flex flex-col gap-2.5">
              <button 
                onClick={() => {
                  setViewMode('public');
                  setCurrentUser(null);
                  localStorage.removeItem('injaz_visa_current_user');
                }}
                className="w-full bg-slate-950 hover:bg-slate-800 text-white py-3 rounded-2xl text-center text-xs font-black flex items-center justify-center gap-2 shadow-md cursor-pointer"
              >
                <ArrowRight size={14} />
                <span>{t.backToPublic}</span>
              </button>

              <div className="p-4 bg-emerald-50 rounded-2xl border border-emerald-100 text-start">
                <p className="text-[10px] font-bold text-emerald-800 uppercase mb-2">Algorithmic Engine</p>
                <div className="flex items-center gap-2">
                  <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
                  <span className="text-[10px] font-bold text-emerald-600">Active - Standalone Mode</span>
                </div>
              </div>
            </div>
          </aside>

          {/* Core Admin Workspace Area */}
          <main className="flex-1 overflow-y-auto bg-[#F8F9FA] dark:bg-[#090d16] text-slate-900 dark:text-slate-100 p-4 md:p-8 space-y-8 transition-colors duration-300">
            <header className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex items-center gap-3">
                <button 
                  onClick={() => setIsSidebarOpen(true)}
                  className="lg:hidden w-10 h-10 flex items-center justify-center text-slate-500 hover:text-slate-900 border border-slate-200 rounded-xl bg-white shadow-sm"
                >
                  <Menu size={24} />
                </button>
                <div className="text-right">
                  <h2 className="text-base md:text-lg font-black tracking-tight flex items-center gap-2">
                    {t[`tab_${activeTab}`]}
                    <span className="bg-slate-200 text-slate-600 text-xs px-2.5 py-0.5 rounded-full">{activeTab === 'dashboard' ? stats.total : filteredByTab.length}</span>
                  </h2>
                  <p className="text-sm font-medium text-slate-400 mt-1">
                    {activeTab === 'dashboard' ? (
                      lang === 'ar' ? 'نظرة عامة على كافة العمليات الإدارية والمجاميع' : 'Overview of all administrative operations'
                    ) : activeTab === 'files' ? (
                      lang === 'ar' ? 'عرض وتحميل وحذف ملفات PDF الخاصة بالعملاء' : 'View, download and delete client PDF files'
                    ) : activeTab === 'customers' ? null : (
                      lang === 'ar' ? 'تصفية وتعديل البيانات حسب الحالة' : 'Filter and edit entries depending on case state'
                    )}
                  </p>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <button 
                  onClick={() => {
                    setViewMode('public');
                    setCurrentUser(null);
                    localStorage.removeItem('injaz_visa_current_user');
                  }}
                  className="bg-white text-slate-800 border border-slate-200 rounded-xl px-4 py-2 text-sm font-bold hover:bg-slate-50 transition-all shadow-sm hidden sm:block cursor-pointer"
                >
                  {t.backToPublic}
                </button>
              </div>
            </header>

            <React.Suspense fallback={
              <div className="flex items-center justify-center h-64">
                <div className="w-8 h-8 border-4 border-blue-500 border-t-transparent rounded-full animate-spin" />
              </div>
            }>
            {activeTab === 'dashboard' ? (
              <DashboardTab
                lang={lang}
                parsing={parsing}
                bulkCheckRunning={bulkCheckRunning}
                bulkCheckIndex={bulkCheckIndex}
                bulkCheckTotal={bulkCheckTotal}
                bulkCheckCurrentStatus={bulkCheckCurrentStatus}
                setBulkCheckCurrentStatus={setBulkCheckCurrentStatus}
                bulkCheckCurrentName={bulkCheckCurrentName}
                bulkCheckResults={bulkCheckResults}
                setBulkCheckResults={setBulkCheckResults}
                checkAll={checkAll}
                deptStats={deptStats}
                entries={entries}
                offices={offices}
                setActiveTab={setActiveTab}
                setAllTabFilter={setAllTabFilter}
                setIsSectionSelected={setIsSectionSelected}
                setExpandedCardId={setExpandedCardId}
                pendingBulkCount={pendingBulkCount}
              />
            ) : activeTab === 'identity' ? (
              <IdentityTab
                lang={lang}
                theme={theme}
                siteIdentity={siteIdentity}
                setSiteIdentity={setSiteIdentityUser}
                identitySuccess={identitySuccess}
                setIdentitySuccess={setIdentitySuccess}
                saveIdentity={saveIdentity}
                addSyncLog={addSyncLog}
                handleLogoFile={handleLogoFile}
                handleLogoDrag={handleLogoDrag}
                handleLogoDrop={handleLogoDrop}
                logoDragActive={logoDragActive}
              />
            ) : activeTab === 'security' ? (
              <SecurityTab
                lang={lang}
                theme={theme}
                adminUsername={adminUsername}
                setAdminUsername={setAdminUsername}
                adminPassword={adminPassword}
                setAdminPassword={setAdminPassword}
                handleSaveCredentials={handleSaveCredentials}
                usersList={usersList}
                newUserName={newUserName}
                setNewUserName={setNewUserName}
                newUserUsername={newUserUsername}
                setNewUserUsername={setNewUserUsername}
                newUserPassword={newUserPassword}
                setNewUserPassword={setNewUserPassword}
                newUserPermissions={newUserPermissions}
                setNewUserPermissions={setNewUserPermissions}
                editingUserId={editingUserId}
                setEditingUserId={setEditingUserId}
                handleUserFormSubmit={handleUserFormSubmit}
                handleDeleteUser={handleDeleteUser}
              />
            ) : activeTab === 'excel_sync' ? (
              <ExcelSyncTab
                lang={lang}
                syncSubView={syncSubView}
                setSyncSubView={setSyncSubView}
                telegramToken={telegramToken}
                setTelegramToken={setTelegramToken}
                telegramEnabled={telegramEnabled}
                setTelegramEnabled={setTelegramEnabled}
                telegramAlertChatId={telegramAlertChatId}
                setTelegramAlertChatId={setTelegramAlertChatId}
                telegramPrintArchiveChatId={telegramPrintArchiveChatId}
                setTelegramPrintArchiveChatId={setTelegramPrintArchiveChatId}
                telegramVisaArchiveChatId={telegramVisaArchiveChatId}
                setTelegramVisaArchiveChatId={setTelegramVisaArchiveChatId}
                handleSaveTelegramConfig={handleSaveTelegramConfig}
                syncLogs={syncLogs}
                googleSpreadsheetId={googleSpreadsheetId}
                setGoogleSpreadsheetId={setGoogleSpreadsheetId}
                googleApiKey={googleApiKey}
                setGoogleApiKey={setGoogleApiKey}
                googleSheetName={googleSheetName}
                setGoogleSheetName={setGoogleSheetName}
                lastSyncTime={lastSyncTime}
                syncStatus={syncStatus}
                executeSynchronization={executeSynchronization}
                addSyncLog={addSyncLog}
                showAlert={showAlert}
              />
            ) : activeTab === 'customers' ? (
              <div className="space-y-8 max-w-6xl mx-auto text-slate-900" dir={lang === 'ar' ? 'rtl' : 'ltr'}>

                {customerSubTab === 'none' ? (
                  <>
                {/* SubTab Navigation Cards */}
                <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                  {/* Card A: Add Customer */}
                  <button
                    onClick={() => {
                      setCustomerSubTab('add');
                      setSelectedCustomerToEdit(null);
                    }}
                    className={`p-6 sm:p-8 rounded-[2.5rem] border text-right transition-all duration-300 relative overflow-hidden flex items-start gap-4 cursor-pointer select-none ${
                      customerSubTab === 'add' 
                        ? 'bg-slate-900 text-white border-slate-900 shadow-xl shadow-slate-900/10' 
                        : 'bg-white text-slate-900 border-slate-200 shadow-sm hover:border-slate-300 hover:shadow-md'
                    }`}
                  >
                    <div className={`p-4 rounded-2xl shrink-0 ${customerSubTab === 'add' ? 'bg-white/10 text-white' : 'bg-violet-50 text-violet-600'}`}>
                      <UserPlus size={28} />
                    </div>
                    <div>
                      <h4 className="font-black text-base md:text-lg">
                        {lang === 'ar' ? 'إضافة عميل جديد' : 'Enroll New Customer'}
                      </h4>
                      <p className={`text-xs mt-1.5 font-medium leading-relaxed ${customerSubTab === 'add' ? 'text-slate-300' : 'text-slate-400'}`}>
                        {lang === 'ar' 
                          ? 'فتح بطاقة إلكترونية لتعيين الاسم والبيانات الأساسية لعميل جديد وحفظها فوراً.' 
                          : 'Open a manual form to register credentials, passport, phone and configure tracking metadata.'}
                      </p>
                    </div>
                  </button>

                  {/* Card B: Excel Import */}
                  <button
                    onClick={() => {
                      setCustomerSubTab('excel_import');
                      setSelectedCustomerToEdit(null);
                      setCustomerExcelSuccess(false);
                      setCustomerExcelAddedCount(0);
                    }}
                    className={`p-6 sm:p-8 rounded-[2.5rem] border text-right transition-all duration-300 relative overflow-hidden flex items-start gap-4 cursor-pointer select-none ${
                      customerSubTab === 'excel_import' 
                        ? 'bg-slate-900 text-white border-slate-900 shadow-xl shadow-slate-900/10' 
                        : 'bg-white text-slate-900 border-slate-200 shadow-sm hover:border-slate-300 hover:shadow-md'
                    }`}
                  >
                    <div className={`p-4 rounded-2xl shrink-0 ${customerSubTab === 'excel_import' ? 'bg-white/10 text-white' : 'bg-emerald-50 text-emerald-600'}`}>
                      <FileSpreadsheet size={28} />
                    </div>
                    <div>
                      <h4 className="font-black text-base md:text-lg">
                        {lang === 'ar' ? 'استيراد عملاء عبر الإكسل' : 'Import Customers from Excel'}
                      </h4>
                      <p className={`text-xs mt-1.5 font-medium leading-relaxed ${customerSubTab === 'excel_import' ? 'text-slate-300' : 'text-slate-400'}`}>
                        {lang === 'ar' 
                          ? 'شرح الترتيب ورفع ملف XLSX لإضافة مجموعة كبيرة من العملاء بلمح البصر.' 
                          : 'Upload fully custom spreadsheets to bulk register users with automatic mapping mechanisms.'}
                      </p>
                    </div>
                  </button>

                  {/* Card C: Export Excel */}
                  <button
                    onClick={() => setIsExportSectionModalOpen(true)}
                    className="p-6 sm:p-8 rounded-[2.5rem] border text-right transition-all duration-300 relative overflow-hidden flex items-start gap-4 cursor-pointer select-none bg-white text-slate-900 border-slate-200 shadow-sm hover:border-slate-300 hover:shadow-md"
                  >
                    <div className="p-4 rounded-2xl shrink-0 bg-emerald-50 text-emerald-600">
                      <Download size={28} />
                    </div>
                    <div>
                      <h4 className="font-black text-base md:text-lg">
                        {lang === 'ar' ? 'تصدير كشف Excel' : 'Export Excel Sheet'}
                      </h4>
                      <p className="text-xs mt-1.5 font-medium leading-relaxed text-slate-400">
                        {lang === 'ar'
                          ? 'اختر قسماً معيناً وتصدير كافة البيانات كملف إكسل.'
                          : 'Select a section to export all details as an Excel file.'}
                      </p>
                    </div>
                  </button>
                </div>
                  </>
                ) : (
                  <>
                {/* Back Navigation Header */}
                <div className="bg-white p-4 sm:p-5 rounded-[2rem] border border-slate-200 shadow-sm flex items-center gap-4">
                  <button
                    onClick={() => setCustomerSubTab('none')}
                    className="flex items-center gap-2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs px-4 py-2.5 rounded-xl transition-all cursor-pointer shrink-0"
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><path d="m9 18 6-6-6-6"/></svg>
                    {lang === 'ar' ? 'رجوع' : 'Back'}
                  </button>
                  <div className="w-px h-6 bg-slate-200 shrink-0" />
                  <div className="flex items-center gap-2.5 min-w-0">
                    <span className={`p-2 rounded-xl shrink-0 ${customerSubTab === 'add' ? 'bg-violet-50 text-violet-600' : customerSubTab === 'edit' ? 'bg-amber-50 text-amber-600' : 'bg-emerald-50 text-emerald-600'}`}>
                      {customerSubTab === 'add' ? <UserPlus size={16} /> : customerSubTab === 'edit' ? <Users size={16} /> : <FileSpreadsheet size={16} />}
                    </span>
                    <div className="min-w-0">
                      <h3 className="font-black text-sm text-slate-900 truncate">
                        {customerSubTab === 'add' ? (lang === 'ar' ? 'إضافة عميل جديد' : 'Enroll New Customer') : customerSubTab === 'edit' ? (lang === 'ar' ? 'البحث وتعديل بيانات عميل' : 'Modify Existing Customer') : (lang === 'ar' ? 'استيراد عملاء عبر الإكسل' : 'Import Customers from Excel')}
                      </h3>
                      <p className="text-[10px] text-slate-400 font-semibold">
                        {lang === 'ar' ? 'إدارة بيانات العملاء ← ' : '← Customer Management'}{customerSubTab === 'add' ? (lang === 'ar' ? 'إضافة' : 'Add') : customerSubTab === 'edit' ? (lang === 'ar' ? 'تعديل' : 'Edit') : (lang === 'ar' ? 'استيراد' : 'Import')}
                      </p>
                    </div>
                  </div>
                </div>

                {/* Sub-panels */}
                <AnimatePresence mode="wait">
                  {customerSubTab === 'excel_import' && (
                    <motion.div
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 20 }}
                      className="bg-white p-6 sm:p-8 rounded-[2.5rem] border border-slate-200 shadow-sm space-y-6 text-right"
                    >
                      <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
                        <span className="p-2 bg-emerald-50 text-emerald-600 rounded-xl"><FileSpreadsheet size={18} /></span>
                        <h4 className="font-black text-sm text-slate-800">{lang === 'ar' ? 'بوابة استيراد بيانات العملاء من إكسل' : 'Bulk Customer Import via Spreadsheet'}</h4>
                      </div>

                      {/* Header guide explanation card */}
                      <div className="bg-slate-50 border border-slate-200/60 p-5 rounded-2xl space-y-4">
                        <h5 className="font-bold text-xs text-slate-700 flex items-center gap-2">
                          💡
                          <span>{lang === 'ar' ? 'كيف يجب أن يكون ترتيب وتسمية الأعمدة في ملف الإكسل؟' : 'How should Excel columns be named and ordered?'}</span>
                        </h5>
                        <p className="text-[11px] font-semibold text-slate-400 leading-relaxed text-right">
                          {lang === 'ar' 
                            ? 'نظام التوثيق الذكي في منصتنا يقوم بمطابقة الحقول تلقائياً بناءً على الكلمات المفتاحية في السطر الأول لملفك (الترويسة). يمكنك ترتيب الأعمدة بأي تسلسل تريده، مع الحرص على تسميتها كالتالي:'
                            : 'Our metadata recognition engine aligns fields automatically based on header keywords in the first row. Arrange columns in any sequence, naming columns as follows:'}
                        </p>

                        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                          {([
                            { label: 'الاسم الكامل للعميل', key: 'الاسم أو Name', req: true },
                            { label: 'رقم جواز السفر', key: 'جواز أو Passport', req: true },
                            { label: 'رقم الهاتف', key: 'هاتف أو Phone', req: false },
                            { label: 'تاريخ الاستلام', key: 'تاريخ الاستلام أو Receive', req: false },
                            { label: 'نوع التأشيرة', key: 'نوع أو Type', req: false },
                            { label: 'جهة القدوم / السفارة', key: 'قدوم أو وصول أو Arrival', req: false },
                            { label: 'اسم المكتب أو الوكيل', key: 'مكتب أو وكيل أو Office', req: false },
                          ] as {label:string;key:string;req:boolean}[]).map(col => (
                            <div key={col.key} className={`bg-white border p-3 rounded-xl flex items-center justify-between gap-2 font-semibold ${col.req ? 'border-red-200/70' : 'border-slate-200/50'}`}>
                              <span className="text-[10px] text-slate-400 flex items-center gap-1 shrink-0">
                                {col.req && <span className="text-red-400 font-black">*</span>}
                                {col.label}
                              </span>
                              <span className="text-[10px] text-slate-700 bg-slate-50 border border-slate-100 px-2 py-0.5 rounded font-bold text-left">{col.key}</span>
                            </div>
                          ))}
                        </div>

                        <div className="space-y-2">
                          <p className="text-[10px] font-bold text-red-600 bg-red-50 border border-red-100 p-3 rounded-xl">
                            <span className="font-black">* الحقول الإجبارية:</span> الاسم + رقم الجواز مطلوبان دائماً. رقم الهاتف مطلوب إذا لم يُحدَّد مكتب أو وكيل. الصفوف التي لا تستوفي هذه الشروط لن تُضاف.
                          </p>
                          <p className="text-[10px] font-bold text-amber-600 bg-amber-50 border border-amber-100 p-3 rounded-xl">
                            ⚠️ الجوازات المكررة أو المسجلة مسبقاً يتم تجاهلها تلقائياً. يجب أن يحتوي الملف على سطر ترويسة (Header) بنفس الأسماء الموضحة أعلاه.
                          </p>
                        </div>
                      </div>

                      {/* File upload panel */}
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 items-center">
                        <div 
                          className={`border-2 border-dashed rounded-3xl p-8 text-center transition-all flex flex-col items-center justify-center gap-3 cursor-pointer select-none ${
                            customerExcelDragActive ? 'border-emerald-500 bg-emerald-50/25' : 'border-slate-200 bg-slate-50 hover:bg-slate-100/50 hover:border-slate-300'
                          }`}
                          onDragOver={(e) => { e.preventDefault(); setCustomerExcelDragActive(true); }}
                          onDragLeave={() => setCustomerExcelDragActive(false)}
                          onDrop={(e) => {
                            e.preventDefault();
                            setCustomerExcelDragActive(false);
                            if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                              const syntheticEvent = { target: { files: e.dataTransfer.files } } as unknown as React.ChangeEvent<HTMLInputElement>;
                              handleCustomerExcelUpload(syntheticEvent);
                            }
                          }}
                          onClick={() => document.getElementById('cust-excel-file-input')?.click()}
                        >
                          <input 
                            type="file" 
                            id="cust-excel-file-input" 
                            className="hidden" 
                            accept=".xlsx,.xls,.csv" 
                            onChange={handleCustomerExcelUpload}
                          />
                          <div className="p-4 bg-emerald-50 text-emerald-600 rounded-2xl">
                            <Upload size={24} />
                          </div>
                          <div>
                            <p className="text-xs font-black text-slate-900">{lang === 'ar' ? 'اسحب ملف الإكسل هنا أو اضغط للتصفح' : 'Drag and drop Excel here, or click to browse'}</p>
                            <p className="text-[10px] text-slate-400 font-bold mt-1 uppercase font-mono">XLSX, XLS, CSV (Max 10MB)</p>
                          </div>
                        </div>

                        {/* Success display state after upload */}
                        <div className="h-full flex flex-col justify-center">
                          {customerExcelSuccess ? (
                            <div className="bg-emerald-50/50 border border-emerald-200/60 p-5 rounded-3xl text-right space-y-3">
                              <div className="flex items-center gap-3 text-emerald-700">
                                <span className="p-2.5 bg-emerald-100 rounded-2xl text-emerald-600"><CheckCircle2 size={22} /></span>
                                <div>
                                  <h5 className="font-extrabold text-sm">تم اكتمال الاستيراد بنجاح!</h5>
                                  <p className="text-[10px] text-emerald-600/80 font-bold">
                                    <span className="font-black text-emerald-800">{customerExcelAddedCount}</span> عميل تمت إضافتهم
                                    {customerExcelSkippedCount > 0 && (
                                      <span className="mr-2 text-amber-600">· <span className="font-black">{customerExcelSkippedCount}</span> متجاهل (مكرر أو ناقص)</span>
                                    )}
                                  </p>
                                </div>
                              </div>
                              <div className="w-full border-t border-emerald-200/40" />
                              <div className="space-y-1.5 max-h-52 overflow-y-auto pr-1">
                                {customerExcelAddedEntries.map((entry, i) => (
                                  <div key={entry.id} className="flex items-center gap-2 bg-white border border-emerald-100 rounded-xl px-3 py-2" dir="rtl">
                                    <span className="text-[9px] font-black text-emerald-600 bg-emerald-50 rounded-md px-1.5 py-0.5 shrink-0">{i + 1}</span>
                                    <div className="flex-1 min-w-0">
                                      <p className="text-xs font-black text-slate-800 truncate">{entry.applicantName}</p>
                                      <p className="text-[10px] font-mono text-slate-400 font-bold">{entry.passportNumber}</p>
                                    </div>
                                    <div className="text-left shrink-0 space-y-0.5">
                                      {entry.visaType && <p className="text-[9px] font-bold text-blue-600 bg-blue-50 rounded px-1.5 py-0.5">{entry.visaType}</p>}
                                      {entry.arrivalPoint && <p className="text-[9px] font-bold text-slate-500">{entry.arrivalPoint}</p>}
                                    </div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          ) : (
                            <div className="border border-slate-150 p-6 rounded-3xl bg-slate-50/20 text-center sm:text-right flex flex-col justify-center h-full space-y-2">
                              <p className="text-xs font-black text-slate-500">بانتظار رفع ملف الإكسل...</p>
                              <p className="text-[10px] text-slate-400 font-medium leading-relaxed">
                                عند رفع الملف، تظهر هنا قائمة بالعملاء الذين تمت إضافتهم بنجاح.
                              </p>
                            </div>
                          )}
                        </div>
                      </div>

                      <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-100">
                        <button
                          type="button"
                          onClick={() => setCustomerSubTab('none')}
                          className="py-2.5 px-6 border border-slate-200 hover:bg-slate-50 text-slate-600 font-bold text-xs rounded-xl transition-all cursor-pointer"
                        >
                          {lang === 'ar' ? 'إغلاق البوابة' : 'Close Panel'}
                        </button>
                      </div>
                    </motion.div>
                  )}

                  {customerSubTab === 'add' && (
                    <motion.div
                      initial={{ opacity: 0, y: 24 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 16 }}
                      transition={{ duration: 0.3, ease: 'easeOut' }}
                      className="rounded-[2rem] overflow-hidden border border-slate-200 shadow-lg shadow-slate-900/5"
                    >
                      {/* Card Header */}
                      <div className="bg-gradient-to-l from-violet-600 to-violet-800 px-6 sm:px-8 py-5 flex items-center justify-between">
                        <div className="flex items-center gap-3">
                          <div className="w-9 h-9 rounded-xl bg-white/15 flex items-center justify-center backdrop-blur-sm">
                            <UserPlus size={18} className="text-white" />
                          </div>
                          <div>
                            <h4 className="font-black text-sm text-white tracking-tight">{lang === 'ar' ? 'تسجيل عميل جديد' : 'New Customer'}</h4>
                            <p className="text-[10px] text-violet-200 font-semibold mt-0.5">{lang === 'ar' ? 'أدخل البيانات الأساسية وأضف المعاملة فوراً' : 'Fill in details and register instantly'}</p>
                          </div>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <span className="w-2 h-2 rounded-full bg-red-400/70" />
                          <span className="w-2 h-2 rounded-full bg-amber-400/70" />
                          <span className="w-2 h-2 rounded-full bg-emerald-400/70" />
                        </div>
                      </div>

                      <div className="bg-white p-6 sm:p-8 space-y-7">

                        {/* Section 1 — البيانات الأساسية */}
                        <div className="space-y-3">
                          <div className="flex items-center gap-2 mb-1">
                            <span className="w-5 h-5 rounded-md bg-violet-100 text-violet-600 flex items-center justify-center text-[10px] font-black">١</span>
                            <span className="text-[11px] font-black text-slate-400 uppercase tracking-widest">{lang === 'ar' ? 'البيانات الأساسية' : 'Basic Information'}</span>
                          </div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 text-right">
                            {/* Name */}
                            <div className="space-y-1.5">
                              <label className="flex items-center gap-1 text-[11px] font-bold text-slate-600">
                                الاسم الكامل
                                <span className="text-red-500 text-xs">*</span>
                              </label>
                              <input
                                type="text"
                                placeholder="أدخل الاسم الكامل للمسافر"
                                className={`w-full h-11 bg-slate-50 border-2 rounded-xl px-4 text-sm font-semibold focus:outline-none focus:bg-white transition-all text-slate-900 placeholder:text-slate-300 placeholder:font-normal ${newCustName.trim() ? 'border-violet-300 focus:border-violet-500' : 'border-red-200 focus:border-red-400'}`}
                                value={newCustName}
                                onChange={(e) => setNewCustName(e.target.value)}
                              />
                            </div>
                            {/* Passport */}
                            <div className="space-y-1.5">
                              <label className="flex items-center gap-1 text-[11px] font-bold text-slate-600">
                                رقم جواز السفر
                                <span className="text-red-500 text-xs">*</span>
                              </label>
                              <input
                                type="text"
                                placeholder="A1234567"
                                className="w-full h-11 bg-slate-50 border-2 border-slate-200 rounded-xl px-4 text-sm font-bold focus:outline-none focus:bg-white focus:border-violet-500 transition-all text-slate-900 uppercase font-mono tracking-widest placeholder:normal-case placeholder:tracking-normal placeholder:text-slate-300 placeholder:font-normal"
                                value={newCustPassport}
                                onChange={(e) => setNewCustPassport(e.target.value.toUpperCase())}
                              />
                            </div>
                            {/* Nationality — fixed */}
                            <div className="space-y-1.5">
                              <label className="text-[11px] font-bold text-slate-600">الجنسية</label>
                              <div className="w-full h-11 bg-slate-100 border-2 border-slate-200 rounded-xl px-4 text-sm font-bold text-slate-600 flex items-center gap-2 select-none">
                                <span className="text-base">🇾🇪</span>
                                <span>اليمن</span>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div className="border-t border-dashed border-slate-100" />

                        {/* Section 2 — بيانات التواصل */}
                        <div className="space-y-3">
                          <div className="flex items-center gap-2 mb-1">
                            <span className="w-5 h-5 rounded-md bg-amber-100 text-amber-600 flex items-center justify-center text-[10px] font-black">٢</span>
                            <span className="text-[11px] font-black text-slate-400 uppercase tracking-widest">{lang === 'ar' ? 'بيانات التواصل والمكتب' : 'Contact & Office'}</span>
                          </div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 text-right">
                            {/* Office */}
                            <div className="space-y-1.5">
                              <label className="text-[11px] font-bold text-slate-600">المكتب / الوكيل <span className="text-slate-300 font-medium">(اختياري)</span></label>
                              <div className="relative">
                                <select
                                  className="w-full h-11 bg-slate-50 border-2 border-slate-200 rounded-xl px-4 text-sm font-semibold focus:outline-none focus:bg-white focus:border-amber-400 transition-all text-slate-900 cursor-pointer appearance-none"
                                  value={newCustOfficeId}
                                  onChange={(e) => {
                                    const oid = e.target.value;
                                    setNewCustOfficeId(oid);
                                    if (!oid) {
                                      // عام — الرقم المخصص فارغ
                                      setNewCustCustomId('');
                                    } else {
                                      const office = offices.find(o => o.id === oid);
                                      // دائماً يُحدَّث — إن لم يكن للمكتب رقم يُفرَّغ الحقل
                                      setNewCustCustomId(office?.phone || '');
                                    }
                                  }}
                                >
                                  <option value="">— عام (بدون مكتب) —</option>
                                  {offices.map(office => (
                                    <option key={office.id} value={office.id}>{office.name}</option>
                                  ))}
                                </select>
                                <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="m6 9 6 6 6-6"/></svg>
                                </span>
                              </div>
                            </div>
                            {/* Custom ID */}
                            <div className="space-y-1.5">
                              <label className="flex items-center gap-1 text-[11px] font-bold text-slate-600">
                                الرقم المخصص
                                <span className="text-slate-300 font-medium">(اختياري)</span>
                              </label>
                              <input
                                type="text"
                                placeholder={newCustOfficeId ? 'يُملأ تلقائياً من رقم المكتب' : 'أدخل الرقم المخصص'}
                                className="w-full h-11 bg-slate-50 border-2 border-slate-200 rounded-xl px-4 text-sm font-semibold focus:outline-none focus:bg-white focus:border-amber-400 transition-all text-slate-900 font-mono placeholder:font-normal placeholder:text-slate-300"
                                value={newCustCustomId}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  setNewCustCustomId(val);
                                  // إذا طابق رقم مكتب → يُعيَّن تلقائياً، وإلا → قسم عام
                                  const norm = val.trim().toLowerCase();
                                  const matched = offices.find(o => o.phone && o.phone.trim().toLowerCase() === norm);
                                  setNewCustOfficeId(matched ? matched.id : '');
                                }}
                              />
                            </div>
                            {/* Phone */}
                            <div className="space-y-1.5">
                              <label className="flex items-center gap-1 text-[11px] font-bold text-slate-600">
                                رقم الهاتف
                                {!newCustOfficeId
                                  ? <span className="text-red-500 text-xs">*</span>
                                  : <span className="text-slate-300 font-medium">(اختياري)</span>}
                              </label>
                              <input
                                type="text"
                                placeholder="7712345678"
                                className={`w-full h-11 bg-slate-50 border-2 rounded-xl px-4 text-sm font-semibold focus:outline-none focus:bg-white transition-all text-slate-900 font-mono placeholder:font-normal placeholder:text-slate-300 ${
                                  !newCustOfficeId && !newCustPhone.trim()
                                    ? 'border-red-200 focus:border-red-400'
                                    : 'border-slate-200 focus:border-amber-400'
                                }`}
                                value={newCustPhone}
                                onChange={(e) => setNewCustPhone(e.target.value)}
                              />
                            </div>
                          </div>
                        </div>

                        <div className="border-t border-dashed border-slate-100" />

                        {/* Section 3 — بيانات التأشيرة */}
                        <div className="space-y-3">
                          <div className="flex items-center gap-2 mb-1">
                            <span className="w-5 h-5 rounded-md bg-emerald-100 text-emerald-600 flex items-center justify-center text-[10px] font-black">٣</span>
                            <span className="text-[11px] font-black text-slate-400 uppercase tracking-widest">{lang === 'ar' ? 'بيانات التأشيرة والمعاملة' : 'Visa & Transaction'}</span>
                          </div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 text-right">
                            {/* Arrival */}
                            <div className="space-y-1.5">
                              <label className="text-[11px] font-bold text-slate-600">جهة القدوم / السفارة</label>
                              <div className="relative">
                                <select
                                  className="w-full h-11 bg-slate-50 border-2 border-slate-200 rounded-xl px-4 text-sm font-semibold focus:outline-none focus:bg-white focus:border-emerald-400 transition-all text-slate-900 cursor-pointer appearance-none"
                                  value={newCustArrival}
                                  onChange={(e) => setNewCustArrival(e.target.value)}
                                >
                                  <option value="عدن">عدن</option>
                                  <option value="صنعاء">صنعاء</option>
                                </select>
                                <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="m6 9 6 6 6-6"/></svg>
                                </span>
                              </div>
                            </div>
                            {/* Visa Type */}
                            <div className="space-y-1.5">
                              <label className="text-[11px] font-bold text-slate-600">نوع التأشيرة</label>
                              <div className="relative">
                                <select
                                  className="w-full h-11 bg-slate-50 border-2 border-slate-200 rounded-xl px-4 text-sm font-semibold focus:outline-none focus:bg-white focus:border-emerald-400 transition-all text-slate-900 cursor-pointer appearance-none"
                                  value={newCustVisaType}
                                  onChange={(e) => setNewCustVisaType(e.target.value)}
                                >
                                  <option value="عمل">عمل</option>
                                  <option value="عمالة منزلية">عمالة منزلية</option>
                                  <option value="مرافق عمال">مرافق عمال</option>
                                  <option value="زيارة">زيارة</option>
                                  <option value="زيارة عائلية">زيارة عائلية</option>
                                  <option value="عائلة مقيم">عائلة مقيم</option>
                                  <option value="دراسة">دراسة</option>
                                  <option value="مرافق طالب">مرافق طالب</option>
                                  <option value="حج">حج</option>
                                  <option value="عمرة">عمرة</option>
                                  <option value="موسمية">موسمية</option>
                                  <option value="تجارة">تجارة</option>
                                  <option value="استثمار">استثمار</option>
                                  <option value="عبور">عبور</option>
                                  <option value="مؤتمر">مؤتمر</option>
                                  <option value="رياضية">رياضية</option>
                                  <option value="علاج طبي">علاج طبي</option>
                                  <option value="صحفية">صحفية</option>
                                  <option value="دينية">دينية</option>
                                  <option value="دبلوماسية">دبلوماسية</option>
                                  <option value="مهمة رسمية">مهمة رسمية</option>
                                  <option value="خاصة">خاصة</option>
                                  <option value="فنية وترفيهية">فنية وترفيهية</option>
                                  <option value="سياحية">سياحية</option>
                                  <option value="تطوع">تطوع</option>
                                  <option value="مؤقتة">مؤقتة</option>
                                  <option value="أخرى">أخرى</option>
                                </select>
                                <span className="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                                  <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><path d="m6 9 6 6 6-6"/></svg>
                                </span>
                              </div>
                            </div>
                            {/* Receive Date */}
                            <div className="space-y-1.5">
                              <label className="text-[11px] font-bold text-slate-600">تاريخ استلام المعاملة <span className="text-slate-300 font-medium">(اختياري)</span></label>
                              <div className="relative w-full">
                                <input
                                  type="date"
                                  className="w-full h-11 bg-slate-50 border-2 border-slate-200 rounded-xl px-4 text-sm font-semibold focus:outline-none focus:bg-white focus:border-emerald-400 transition-all text-slate-700 cursor-pointer appearance-none [&::-webkit-calendar-picker-indicator]:opacity-50 [&::-webkit-calendar-picker-indicator]:cursor-pointer"
                                  value={newCustReceiveDate}
                                  onChange={(e) => setNewCustReceiveDate(e.target.value)}
                                />
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* Action Buttons */}
                        <div className="flex items-center justify-between gap-3 pt-2 mt-2 border-t border-slate-100">
                          <p className="text-[10px] text-slate-400 font-semibold hidden sm:block">
                            <span className="text-red-400">*</span> الحقول الإجبارية: الاسم + رقم الجواز
                          </p>
                          <div className="flex items-center gap-3 mr-auto">
                            <motion.button
                              type="button"
                              whileHover={{ scale: 1.02 }}
                              whileTap={{ scale: 0.97 }}
                              onClick={() => { setNewCustSaving(false); setCustomerSubTab('none'); }}
                              disabled={newCustSaving}
                              className="h-10 px-6 border-2 border-slate-200 hover:border-slate-300 hover:bg-slate-50 text-slate-600 font-bold text-xs rounded-xl transition-all cursor-pointer disabled:opacity-40"
                            >
                              {lang === 'ar' ? 'إلغاء' : 'Cancel'}
                            </motion.button>
                            <motion.button
                              type="button"
                              whileHover={{ scale: newCustSaving ? 1 : 1.02 }}
                              whileTap={{ scale: newCustSaving ? 1 : 0.97 }}
                              onClick={handleAddNewCustomerManually}
                              disabled={newCustSaving}
                              className="relative h-10 px-7 bg-violet-600 hover:bg-violet-700 text-white font-black text-xs rounded-xl shadow-lg shadow-violet-600/20 transition-all cursor-pointer overflow-hidden disabled:cursor-not-allowed"
                            >
                              <AnimatePresence mode="wait">
                                {newCustSaving ? (
                                  <motion.span
                                    key="saving"
                                    initial={{ opacity: 0, y: 8 }}
                                    animate={{ opacity: 1, y: 0 }}
                                    exit={{ opacity: 0, y: -8 }}
                                    className="flex items-center gap-2"
                                  >
                                    <svg className="animate-spin w-3.5 h-3.5" viewBox="0 0 24 24" fill="none"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="white" strokeWidth="4"/><path className="opacity-75" fill="white" d="M4 12a8 8 0 018-8v8z"/></svg>
                                    {lang === 'ar' ? 'جاري الحفظ...' : 'Saving...'}
                                  </motion.span>
                                ) : (
                                  <motion.span
                                    key="save"
                                    initial={{ opacity: 0, y: 8 }}
                                    animate={{ opacity: 1, y: 0 }}
                                    exit={{ opacity: 0, y: -8 }}
                                    className="flex items-center gap-2"
                                  >
                                    <UserPlus size={13} />
                                    {lang === 'ar' ? 'حفظ وإضافة العميل' : 'Save & Add Client'}
                                  </motion.span>
                                )}
                              </AnimatePresence>
                            </motion.button>
                          </div>
                        </div>
                      </div>
                    </motion.div>
                  )}

                  {customerSubTab === 'edit' && (
                    <motion.div
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      exit={{ opacity: 0, y: 20 }}
                      className="space-y-6"
                    >
                      {/* Search Tool Card */}
                      <div className="bg-white p-6 sm:p-8 rounded-[2.5rem] border border-slate-200 shadow-sm space-y-4">
                        <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
                          <span className="p-2 bg-amber-50 text-amber-600 rounded-xl"><Search size={18} /></span>
                          <h4 className="font-black text-sm text-slate-800">{lang === 'ar' ? 'البحث عن العملاء لتحديث بياناتهم' : 'Interactive Client Database Query'}</h4>
                        </div>

                        <p className="text-xs text-slate-400 font-semibold text-start">
                          {lang === 'ar' ? 'اكتب اسم العميل، رقم الجواز، أو رقم الهاتف المسجل لتصفية السجلات فوراً والنقر على العميل لبدء تعديله:' : 'Verify or search record via typing phone, passport or applicant name:'}
                        </p>

                        <div className="relative">
                          <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
                          <input
                            type="text"
                            placeholder={lang === 'ar' ? 'اكتب الاسم أو الجواز أو الهاتف المالي هنا...' : 'Enter matching customer details...'}
                            className="w-full bg-slate-50 border border-slate-200 rounded-2xl py-3 pl-12 pr-6 text-sm focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all text-slate-900 font-semibold"
                            value={customerSearchQuery}
                            onChange={(e) => setCustomerSearchQuery(e.target.value)}
                          />
                        </div>

                        {/* Search Matches Grid */}
                        {customerSearchQuery.trim() !== '' && (
                          <div className="text-right">
                            {(() => {
                              const matches = entries.filter(e => {
                                const q = customerSearchQuery.toLowerCase();
                                return (
                                  e.passportNumber.toLowerCase().includes(q) ||
                                  (e.applicantName || '').toLowerCase().includes(q) ||
                                  (e.phoneNumber || '').toLowerCase().includes(q) ||
                                  (e.documentNumber || '').toLowerCase().includes(q)
                                );
                              });

                              if (matches.length === 0) {
                                return (
                                  <div className="flex flex-col items-center justify-center py-10 gap-3 bg-slate-50 rounded-2xl border border-slate-200">
                                    <div className="w-12 h-12 rounded-full bg-slate-100 flex items-center justify-center">
                                      <Search size={20} className="text-slate-300" />
                                    </div>
                                    <p className="text-xs font-semibold text-slate-400 italic">{lang === 'ar' ? 'لم يتم العثور على نتائج لمقاطعة البحث.' : 'No customer records match this search.'}</p>
                                  </div>
                                );
                              }

                              return (
                                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                                  {matches.slice(0, 12).map((entry) => {
                                    const isSelected = selectedCustomerToEdit?.id === entry.id;
                                    const officeObj = offices.find(o => o.id === entry.officeId);
                                    const statusColor =
                                      entry.status === 'Found' ? 'bg-emerald-500' :
                                      entry.status === 'Error' ? 'bg-red-400' :
                                      entry.status === 'Checking' ? 'bg-blue-400' :
                                      entry.status === 'Retrying' ? 'bg-amber-400' :
                                      'bg-slate-300';
                                    const initials = (entry.applicantName || '?').split(' ').map(w => w[0]).slice(0, 2).join('').toUpperCase();

                                    return (
                                      <button
                                        key={entry.id}
                                        onClick={() => {
                                          setSelectedCustomerToEdit(entry);
                                          setEditCustName(entry.applicantName || '');
                                          setEditCustPassport(entry.passportNumber);
                                          setEditCustAppNo(entry.applicationNumber);
                                          setEditCustPhone(entry.phoneNumber || '');
                                          setEditCustNationality(entry.nationality || 'اليمن');
                                          setEditCustArrival(entry.arrivalPoint || 'جدة');
                                          setEditCustVisaType(entry.visaType || 'تأشيرة عمل');
                                          setEditCustProfession(entry.profession || '');
                                          setEditCustVisaNumber(entry.visaNumber || '');
                                          setEditCustStatus(entry.status);
                                          setEditCustStatusText(entry.statusText || '');
                                          setEditCustApplicantData(entry.applicantData || '');
                                        }}
                                        className={`group relative w-full text-right rounded-2xl border p-4 transition-all duration-200 hover:shadow-md hover:-translate-y-0.5 focus:outline-none ${
                                          isSelected
                                            ? 'bg-amber-50 border-amber-300 shadow-md shadow-amber-100 ring-2 ring-amber-300/40'
                                            : 'bg-white border-slate-200 hover:border-slate-300 shadow-sm'
                                        }`}
                                      >
                                        {/* Top row: avatar + name + status dot */}
                                        <div className="flex items-start gap-3 mb-3">
                                          <div className={`shrink-0 w-10 h-10 rounded-xl flex items-center justify-center text-sm font-black text-white ${isSelected ? 'bg-amber-400' : 'bg-slate-700'}`}>
                                            {initials}
                                          </div>
                                          <div className="flex-1 min-w-0 text-right">
                                            <div className="flex items-center justify-between gap-2">
                                              <span className={`w-2 h-2 rounded-full shrink-0 ${statusColor}`} title={entry.status} />
                                              <p className="text-sm font-black text-slate-900 truncate flex-1 text-right">{entry.applicantName || '---'}</p>
                                            </div>
                                            <p className="text-[10px] font-mono text-slate-400 font-semibold mt-0.5 text-right">{entry.passportNumber}</p>
                                          </div>
                                        </div>

                                        {/* Middle row: profession */}
                                        <div className="flex flex-col gap-1 mb-3">
                                          {entry.profession && (
                                            <div className="flex items-center justify-end gap-1.5 text-[10px] text-slate-400 font-semibold">
                                              <span>{entry.profession}</span>
                                              <Briefcase size={9} className="text-slate-300 shrink-0" />
                                            </div>
                                          )}
                                        </div>

                                        {/* Office name box */}
                                        {(() => {
                                          const displayName = officeObj ? officeObj.name : (lang === 'ar' ? 'عام' : 'General');
                                          return (
                                            <div className="flex items-center justify-between gap-2 bg-violet-50 border border-violet-200 rounded-xl px-3 py-2 mb-2">
                                              <Briefcase size={12} className="text-violet-400 shrink-0" />
                                              <span className="text-[11px] font-black text-violet-700 flex-1 text-right truncate">{displayName}</span>
                                              <span className="text-[9px] font-bold text-violet-400 shrink-0">{lang === 'ar' ? 'المكتب' : 'Office'}</span>
                                            </div>
                                          );
                                        })()}

                                        {/* Bottom badges: visa type */}
                                        <div className="flex flex-wrap gap-1.5 justify-end">
                                          {entry.visaType && (
                                            <span className="inline-flex items-center gap-1 text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-100 rounded-full px-2 py-0.5">
                                              <Tag size={8} />
                                              {entry.visaType}
                                            </span>
                                          )}
                                        </div>

                                        {/* Selected indicator */}
                                        {isSelected && (
                                          <div className="absolute top-3 left-3 w-5 h-5 rounded-full bg-amber-400 flex items-center justify-center shadow-sm">
                                            <CheckCircle2 size={12} className="text-white" />
                                          </div>
                                        )}
                                      </button>
                                    );
                                  })}
                                </div>
                              );
                            })()}
                          </div>
                        )}
                      </div>

                      {/* Edit form panel for the selected customer */}
                      {selectedCustomerToEdit && (
                        <motion.div
                          initial={{ opacity: 0, y: 10 }}
                          animate={{ opacity: 1, y: 0 }}
                          className="bg-white p-6 sm:p-8 rounded-[2.5rem] border border-slate-200 shadow-sm space-y-6 text-right"
                        >
                          <div className="flex items-center gap-3 border-b border-slate-100 pb-4">
                            <span className="p-2 bg-amber-50 text-amber-600 rounded-xl"><UserPlus size={18} /></span>
                            <h4 className="font-black text-sm text-slate-800">{lang === 'ar' ? 'تعديل بيانات العميل' : 'Modify Enrolled Customer Details'}</h4>
                          </div>

                          {/* ── رقم المستند — يظهر دائماً بغض النظر عن باقي البيانات ── */}
                          {selectedCustomerToEdit && (
                            <div className="bg-amber-50 border border-amber-200 rounded-xl px-4 py-3 flex items-center justify-between gap-3">
                              <span className="text-[10px] font-black text-amber-600 uppercase tracking-widest whitespace-nowrap">رقم المستند</span>
                              <span className="text-sm font-black text-amber-900 font-mono tracking-wide">
                                {selectedCustomerToEdit.documentNumber && selectedCustomerToEdit.documentNumber !== '' ? selectedCustomerToEdit.documentNumber : '---'}
                              </span>
                            </div>
                          )}


                          {/* ── MOFA Data (read-only) ── */}
                          {selectedCustomerToEdit && (selectedCustomerToEdit.applicationNumber && selectedCustomerToEdit.applicationNumber !== '---' || selectedCustomerToEdit.applicantName && selectedCustomerToEdit.applicantName !== '---') && (
                            <div className="space-y-3">
                              <div className="flex items-center gap-2 pb-1">
                                <span className="text-[10px] font-black text-emerald-600 uppercase tracking-widest bg-emerald-50 border border-emerald-100 px-2.5 py-1 rounded-full">من منصة التأشيرات</span>
                              </div>
                              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3 text-right">
                                {[
                                  { label: 'رقم الطلب', val: selectedCustomerToEdit.applicationNumber },
                                  { label: 'تاريخ الطلب', val: selectedCustomerToEdit.applicationDate },
                                  { label: 'الممثلية في', val: selectedCustomerToEdit.embassy },
                                  { label: 'نوع التأشيرة', val: selectedCustomerToEdit.visaType },
                                  { label: 'عدد مرات الدخول', val: selectedCustomerToEdit.entriesCount },
                                  { label: 'اسم الجهة الطالبة', val: selectedCustomerToEdit.requesterName },
                                  { label: 'الاسم', val: selectedCustomerToEdit.applicantName },
                                  { label: 'Name', val: selectedCustomerToEdit.nameEnglish },
                                  { label: 'رقم الجواز', val: selectedCustomerToEdit.passportNumber },
                                  { label: 'نوع الجواز', val: selectedCustomerToEdit.passportType },
                                  { label: 'تاريخ الانتهاء', val: selectedCustomerToEdit.passportExpiry },
                                  { label: 'تاريخ الميلاد', val: selectedCustomerToEdit.birthDate },
                                  { label: 'مكان الميلاد', val: selectedCustomerToEdit.birthPlace },
                                  { label: 'الجنسية الحالية', val: selectedCustomerToEdit.currentNationality },
                                  { label: 'الجنس', val: selectedCustomerToEdit.gender },
                                  { label: 'المهنة', val: selectedCustomerToEdit.profession },
                                  { label: 'الغرض', val: selectedCustomerToEdit.purpose },
                                ].filter(f => f.val && f.val !== '---').map(f => (
                                  <div key={f.label} className="bg-slate-50 border border-slate-100 rounded-xl px-3 py-2.5 space-y-0.5">
                                    <p className="text-[9px] font-black text-slate-400 uppercase tracking-wider">{f.label}</p>
                                    <p className="text-xs font-bold text-slate-800 truncate">{f.val}</p>
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* ── Editable Fields ── */}
                          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 text-right">
                            {/* Office / Agent (read-only display) */}
                            {(() => {
                              const officeObj = offices.find(o => o.id === selectedCustomerToEdit.officeId);
                              const displayName = officeObj ? officeObj.name : (lang === 'ar' ? 'عام' : 'General');
                              return (
                                <div className="space-y-1.5">
                                  <label className="text-xs font-bold text-slate-500">المكتب / الوكيل</label>
                                  <div className="w-full bg-violet-50 border border-violet-200 rounded-2xl px-4 py-3 text-sm font-bold text-violet-800 flex items-center gap-2 select-none">
                                    <span>🏢</span>
                                    <span>{displayName}</span>
                                  </div>
                                </div>
                              );
                            })()}

                            {/* Phone */}
                            <div className="space-y-1.5">
                              <label className="text-xs font-bold text-slate-500">رقم الهاتف أو الرقم المخصص <span className="text-red-400">*</span></label>
                              <input
                                type="text"
                                className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-4 py-3 text-sm font-semibold focus:outline-none focus:ring-4 focus:ring-amber-400/10 focus:border-amber-400 transition-all text-slate-900 font-mono"
                                value={editCustPhone}
                                onChange={(e) => setEditCustPhone(e.target.value)}
                              />
                            </div>

                            {/* Nationality */}
                            <div className="space-y-1.5">
                              <label className="text-xs font-bold text-slate-500">الجنسية</label>
                              <input
                                type="text"
                                className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-4 py-3 text-sm font-semibold focus:outline-none focus:ring-4 focus:ring-amber-400/10 focus:border-amber-400 transition-all text-slate-900"
                                value={editCustNationality}
                                onChange={(e) => setEditCustNationality(e.target.value)}
                              />
                            </div>

                            {/* Arrival */}
                            <div className="space-y-1.5">
                              <label className="text-xs font-bold text-slate-500">جهة القدوم</label>
                              <select
                                className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-4 py-3 text-sm font-semibold focus:outline-none focus:ring-4 focus:ring-amber-400/10 focus:border-amber-400 transition-all text-slate-900 cursor-pointer"
                                value={editCustArrival}
                                onChange={(e) => setEditCustArrival(e.target.value)}
                              >
                                <option value="عدن">عدن</option>
                                <option value="صنعاء">صنعاء</option>
                                <option value="جدة">جدة</option>
                                <option value="الرياض">الرياض</option>
                                <option value="الدمام">الدمام</option>
                                <option value="أخرى">أخرى</option>
                              </select>
                            </div>

                            {/* Status */}
                            <div className="space-y-1.5">
                              <label className="text-xs font-bold text-slate-500">حالة الطلب</label>
                              <select
                                className="w-full bg-slate-50 border border-slate-200 rounded-2xl px-4 py-3 text-sm font-semibold focus:outline-none focus:ring-4 focus:ring-amber-400/10 focus:border-amber-400 transition-all text-slate-900 cursor-pointer"
                                value={editCustStatus}
                                onChange={(e) => setEditCustStatus(e.target.value as any)}
                              >
                                <option value="Idle">قيد الانتظار</option>
                                <option value="Found">تم الإصدار</option>
                                <option value="Checking">قيد المراجعة</option>
                                <option value="Error">خطأ</option>
                              </select>
                            </div>
                          </div>

                          {/* Action Buttons for Edit Form */}
                          <div className="flex flex-col sm:flex-row items-center justify-between gap-4 pt-4 border-t border-slate-100 text-semibold">
                            {/* Delete Button */}
                            <button
                              type="button"
                              onClick={() => handleDeleteCustomerManually(selectedCustomerToEdit)}
                              className="w-full sm:w-auto py-2.5 px-6 bg-red-50 hover:bg-red-100 text-red-600 border border-red-100 font-bold text-xs rounded-xl flex items-center justify-center gap-1.5 cursor-pointer"
                            >
                              <Trash2 size={14} />
                              <span>{lang === 'ar' ? 'حذف هذا العميل نهائياً' : 'Delete Candidate Record'}</span>
                            </button>

                            {/* Save / Cancel buttons */}
                            <div className="flex items-center gap-2.5 w-full sm:w-auto">
                              <button
                                type="button"
                                onClick={() => setSelectedCustomerToEdit(null)}
                                className="w-1/2 sm:w-auto py-2.5 px-6 border border-slate-200 hover:bg-slate-50 text-slate-600 font-bold text-xs rounded-xl cursor-pointer"
                              >
                                {lang === 'ar' ? 'إلغاء التعديل' : 'Cancel'}
                              </button>
                              <button
                                type="button"
                                onClick={handleSaveEditedCustomerManually}
                                className="w-1/2 sm:w-auto py-2.5 px-8 bg-amber-500 hover:bg-amber-600 text-white font-bold text-xs rounded-xl shadow-lg shadow-amber-500/10 cursor-pointer"
                              >
                                {lang === 'ar' ? 'حفظ التغييرات الآن' : 'Save Changes'}
                              </button>
                            </div>
                          </div>
                        </motion.div>
                      )}
                    </motion.div>
                  )}
                </AnimatePresence>
                  </>
                )}
              </div>
            ) : activeTab === 'files' ? (
              <FilesTab
                lang={lang}
                pdfFiles={pdfFiles}
                setPdfFiles={setPdfFiles}
                pdfFilesLoading={pdfFilesLoading}
                setPdfFilesLoading={setPdfFilesLoading}
                visaFiles={visaFiles}
                setVisaFiles={setVisaFiles}
                visaFilesLoading={visaFilesLoading}
                setVisaFilesLoading={setVisaFilesLoading}
                filesSubPage={filesSubPage}
                setFilesSubPage={setFilesSubPage}
                filesSearch={filesSearch}
                setFilesSearch={setFilesSearch}
                showConfirm={showConfirm}
                setEntries={setEntries}
                handleBulkUploadToTelegram={handleBulkUploadToTelegram}
                bulkTgUploading={bulkTgUploading}
                bulkTgResult={bulkTgResult}
              />
            ) : activeTab === 'auto_check' ? (
              <AutoCheckTab
                lang={lang}
                autoCheckEnabled={autoCheckEnabled}
                setAutoCheckEnabled={setAutoCheckEnabled}
                autoCheckTime={autoCheckTime}
                setAutoCheckTime={setAutoCheckTime}
                autoCheckInterval={autoCheckInterval}
                setAutoCheckInterval={setAutoCheckInterval}
                autoCheckLastRun={autoCheckLastRun}
                autoCheckNextRun={autoCheckNextRun}
                autoCheckLog={autoCheckLog}
                setAutoCheckLog={setAutoCheckLog}
                bulkCheckRunning={bulkCheckRunning}
                checkAll={checkAll}
                showAlert={showAlert}
                setActiveTab={setActiveTab}
                setAutoCheckLastRun={setAutoCheckLastRun}
              />
            ) : (
              <>
                {!isSectionSelected ? (
                  isOfficesView ? (
                    isOfficeDashboardView && selectedOfficeFilter ? (() => {
                      const officeObj = selectedOfficeFilter === 'general' ? null : offices.find(o => o.id === selectedOfficeFilter);
                      const officeName = officeObj ? officeObj.name : (lang === 'ar' ? 'عام' : 'General');
                      const officeEntries = selectedOfficeFilter === 'general'
                        ? entries.filter(e => !e.officeId || e.officeId === 'general')
                        : entries.filter(e => e.officeId === selectedOfficeFilter);
                      const uniqueVisaTypes = [...new Set(officeEntries.map(e => e.visaType))].filter(Boolean);
                      const issuedCount = officeEntries.filter(e => e.statusText?.includes('تم اصدار') || e.statusText?.includes('Issued')).length;
                      const todayCount = officeEntries.filter(isToday).length;
                      const withPrintCount = officeEntries.filter(entryHasPrint).length;
                      return (
                        <div className="space-y-6">
                          <div className="bg-white p-6 sm:p-8 rounded-[2rem] border border-slate-200 shadow-sm flex items-center gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                            <button
                              onClick={() => { setIsOfficeDashboardView(false); setSelectedOfficeFilter(null); }}
                              className="w-10 h-10 flex items-center justify-center text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl shadow-sm transition-all cursor-pointer shrink-0"
                            >
                              {lang === 'ar' ? <ArrowRight size={18} /> : <ArrowLeft size={18} />}
                            </button>
                            <div className="text-right flex-1 min-w-0">
                              <span className="bg-purple-50 border border-purple-100/60 text-purple-600 font-extrabold text-[10px] tracking-widest uppercase px-3 py-1 rounded-full inline-flex items-center gap-1.5 mb-2">
                                <Users size={11} />
                                {lang === 'ar' ? 'لوحة المكتب' : 'Office Dashboard'}
                              </span>
                              {editingOfficeId === officeObj?.id ? (
                                <div className="flex flex-col gap-2 mt-1">
                                  <input
                                    type="text"
                                    className="w-full bg-purple-50 border border-purple-300 rounded-xl px-3 py-1.5 text-sm font-bold text-slate-900 focus:outline-none focus:border-purple-500 transition-all"
                                    value={editOfficeForm.name}
                                    onChange={e => setEditOfficeForm({ ...editOfficeForm, name: e.target.value })}
                                    placeholder={lang === 'ar' ? 'اسم المكتب' : 'Office name'}
                                    autoFocus
                                  />
                                  <input
                                    type="text"
                                    className="w-full bg-purple-50 border border-purple-200 rounded-xl px-3 py-1.5 text-sm font-bold text-slate-900 font-mono focus:outline-none focus:border-purple-400 transition-all"
                                    value={editOfficeForm.phone}
                                    onChange={e => setEditOfficeForm({ ...editOfficeForm, phone: e.target.value })}
                                    placeholder={lang === 'ar' ? 'الرقم المخصص / الهاتف' : 'Phone / custom number'}
                                  />
                                  <div className="flex gap-2">
                                    <button
                                      onClick={() => {
                                        if (!editOfficeForm.name.trim()) return;
                                        setOffices(offices.map(o => o.id === officeObj.id ? { ...o, name: editOfficeForm.name.trim(), phone: editOfficeForm.phone.trim() } : o));
                                        setEditingOfficeId(null);
                                      }}
                                      className="flex-1 py-1.5 bg-purple-600 hover:bg-purple-700 text-white font-black text-xs rounded-xl cursor-pointer transition-all"
                                    >
                                      {lang === 'ar' ? 'حفظ' : 'Save'}
                                    </button>
                                    <button
                                      onClick={() => setEditingOfficeId(null)}
                                      className="flex-1 py-1.5 bg-white border border-slate-200 text-slate-600 font-bold text-xs rounded-xl cursor-pointer transition-all"
                                    >
                                      {lang === 'ar' ? 'إلغاء' : 'Cancel'}
                                    </button>
                                  </div>
                                </div>
                              ) : (
                                <div className="flex items-center gap-2">
                                  <div className="flex-1 min-w-0">
                                    <h3 className="text-sm sm:text-base font-black text-slate-900 tracking-tight truncate">{officeName}</h3>
                                    {officeObj?.phone && <p className="text-xs font-mono font-bold text-purple-400 mt-0.5">{officeObj.phone}</p>}
                                  </div>
                                  {officeObj && (
                                    <div className="flex items-center gap-2 shrink-0">
                                      <button
                                        onClick={() => { setEditingOfficeId(officeObj.id); setEditOfficeForm({ name: officeObj.name, phone: officeObj.phone }); }}
                                        className="w-8 h-8 flex items-center justify-center text-purple-400 hover:text-purple-600 hover:bg-purple-50 border border-purple-100 rounded-xl transition-all cursor-pointer"
                                        title={lang === 'ar' ? 'تعديل' : 'Edit'}
                                      >
                                        <Pencil size={14} />
                                      </button>
                                      <button
                                        onClick={() => {
                                          if (officeEntries.length > 0) {
                                            showAlert(lang === 'ar'
                                              ? `لا يمكن حذف "${officeObj.name}" لأنه يحتوي على ${officeEntries.length} عميل. يرجى نقل العملاء أولاً.`
                                              : `Cannot delete "${officeObj.name}" — it has ${officeEntries.length} client(s). Move them first.`, 'warning');
                                            return;
                                          }
                                          showConfirm(
                                            lang === 'ar'
                                              ? `هل أنت متأكد من حذف "${officeObj.name}"؟ لا يمكن التراجع عن هذا الإجراء.`
                                              : `Delete "${officeObj.name}"? This cannot be undone.`,
                                            () => {
                                              setOffices(offices.filter(o => o.id !== officeObj.id));
                                              setIsOfficeDashboardView(false);
                                              setSelectedOfficeFilter(null);
                                            }
                                          );
                                        }}
                                        className="w-8 h-8 flex items-center justify-center text-red-300 hover:text-red-600 hover:bg-red-50 border border-red-100 rounded-xl transition-all cursor-pointer"
                                        title={lang === 'ar' ? 'حذف المكتب' : 'Delete Office'}
                                      >
                                        <Trash2 size={14} />
                                      </button>
                                    </div>
                                  )}
                                </div>
                              )}
                            </div>
                          </div>
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-5" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                            <div className="bg-white border border-slate-100 rounded-[2rem] p-6 shadow-sm space-y-5">
                              <div className="flex items-center gap-3">
                                <div className="w-10 h-10 rounded-xl bg-slate-50 text-slate-600 flex items-center justify-center shrink-0"><BarChart3 size={20} /></div>
                                <div className="text-right flex-1">
                                  <h4 className="text-sm font-black text-slate-900">{lang === 'ar' ? 'إحصائية عامة' : 'General Statistics'}</h4>
                                  <p className="text-[10px] font-bold text-slate-400">{lang === 'ar' ? 'نظرة عامة على معاملات المكتب' : 'Office transactions overview'}</p>
                                </div>
                              </div>
                              <div className="grid grid-cols-2 gap-3">
                                <motion.button whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.97 }} onClick={() => { setAllTabFilter('all'); setIsSectionSelected(true); }} className="bg-slate-900 rounded-2xl p-4 text-start cursor-pointer">
                                  <span className="text-2xl font-black block leading-none text-white">{officeEntries.length}</span>
                                  <p className="text-[10px] font-black text-slate-300 mt-1.5">{lang === 'ar' ? 'إجمالي المعاملات' : 'Total'}</p>
                                </motion.button>
                                <motion.button whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.97 }} onClick={() => { setAllTabFilter('issued'); setIsSectionSelected(true); }} className="bg-emerald-600 rounded-2xl p-4 text-start cursor-pointer">
                                  <span className="text-2xl font-black block leading-none text-white">{issuedCount}</span>
                                  <p className="text-[10px] font-black text-emerald-100 mt-1.5">{lang === 'ar' ? 'المؤشره' : 'Issued'}</p>
                                </motion.button>
                                <motion.button whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.97 }} onClick={() => { setAllTabFilter('today'); setIsSectionSelected(true); }} className="bg-indigo-600 rounded-2xl p-4 text-start cursor-pointer">
                                  <span className="text-2xl font-black block leading-none text-white">{todayCount}</span>
                                  <p className="text-[10px] font-black text-indigo-100 mt-1.5">{lang === 'ar' ? 'تحديثات اليوم' : "Today's"}</p>
                                </motion.button>
                                <motion.button whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.97 }} onClick={() => { setAllTabFilter('with_print'); setIsSectionSelected(true); }} className="bg-blue-600 rounded-2xl p-4 text-start cursor-pointer">
                                  <span className="text-2xl font-black block leading-none text-white">{withPrintCount}</span>
                                  <p className="text-[10px] font-black text-blue-100 mt-1.5">{lang === 'ar' ? 'لها برنت' : 'With Print'}</p>
                                </motion.button>
                                <motion.button whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.97 }} onClick={() => { setAllTabFilter('without_print'); setIsSectionSelected(true); }} className="bg-amber-500 rounded-2xl p-4 text-start cursor-pointer col-span-2">
                                  <span className="text-2xl font-black block leading-none text-white">{officeEntries.length - withPrintCount}</span>
                                  <p className="text-[10px] font-black text-amber-100 mt-1.5">{lang === 'ar' ? 'التي ليس لها برنت' : 'Without Print'}</p>
                                </motion.button>
                                <motion.button whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.97 }}
                                  onClick={() => { setIsArchiveView(true); setArchiveOfficeFilter(selectedOfficeFilter); setArchiveVisaTypeFilter(null); setIsOfficesView(false); setIsOfficeDashboardView(false); setIsSectionSelected(false); }}
                                  className="bg-violet-600 rounded-2xl p-4 text-start cursor-pointer col-span-2">
                                  <span className="text-2xl font-black block leading-none text-white">{officeEntries.filter(e => e.archived).length}</span>
                                  <p className="text-[10px] font-black text-violet-100 mt-1.5">{lang === 'ar' ? 'الأرشيف' : 'Archive'}</p>
                                </motion.button>

                                {/* Custom tx-status cards filtered by office */}
                                {txStatusOptions.filter(opt => opt.hasStatsCard).map(opt => {
                                  const cnt = officeEntries.filter(e => !e.archived && (e.txStatus || '') === opt.label).length;
                                  const isRed = opt.type === 'موقفة';
                                  return (
                                    <motion.button key={opt.id} whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.97 }}
                                      onClick={() => { setAllTabFilter(`tx_status:${opt.label}`); setIsSectionSelected(true); }}
                                      className={`${isRed ? 'bg-red-600' : 'bg-blue-600'} rounded-2xl p-4 text-start cursor-pointer col-span-2`}>
                                      <span className="text-2xl font-black block leading-none text-white">{cnt}</span>
                                      <p className={`text-[10px] font-black ${isRed ? 'text-red-100' : 'text-blue-100'} mt-1.5`}>{opt.label}</p>
                                    </motion.button>
                                  );
                                })}
                              </div>
                            </div>
                            <div className="bg-white border border-blue-100 rounded-[2rem] p-6 shadow-sm space-y-5">
                              <div className="flex items-center gap-3">
                                <div className="w-10 h-10 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center shrink-0"><Tag size={20} /></div>
                                <div className="text-right flex-1">
                                  <h4 className="text-sm font-black text-slate-900">{lang === 'ar' ? 'أقسام المعاملات حسب نوع التأشيرة' : 'Transactions by Visa Type'}</h4>
                                  <p className="text-[10px] font-bold text-slate-400">{uniqueVisaTypes.length} {lang === 'ar' ? 'أنواع تأشيرة' : 'visa types'}</p>
                                </div>
                              </div>
                              <div className="grid grid-cols-2 gap-3 max-h-56 overflow-y-auto pr-1">
                                {uniqueVisaTypes.length === 0 ? (
                                  <div className="col-span-2 text-center py-6 text-slate-400 text-xs font-bold">{lang === 'ar' ? 'لا توجد معاملات لهذا المكتب' : 'No transactions for this office'}</div>
                                ) : uniqueVisaTypes.map(vt => {
                                  const cnt = officeEntries.filter(e => e.visaType === vt).length;
                                  return (
                                    <motion.button key={vt} whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.97 }} onClick={() => { setSelectedVisaType(vt); setAllTabFilter('all'); setIsSectionSelected(true); }} className="bg-slate-50 hover:bg-blue-50 border border-slate-100 hover:border-blue-200 rounded-2xl p-3 text-start transition-all cursor-pointer">
                                      <span className="text-xl font-black block leading-none text-blue-600">{cnt}</span>
                                      <p className="text-[10px] font-black text-slate-600 mt-1 leading-tight">{vt}</p>
                                    </motion.button>
                                  );
                                })}
                              </div>
                            </div>
                          </div>
                        </div>
                      );
                    })() : (
                    /* ── Offices & Agents Sub-Page ── */
                    <div className="space-y-6">
                      {/* Header */}
                      <div className="bg-white p-6 sm:p-8 rounded-[2rem] border border-slate-200 shadow-sm flex items-center gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                        <button
                          onClick={() => { setIsOfficesView(false); setIsAddOfficeForm(false); setSelectedOfficeFilter(null); }}
                          className="w-10 h-10 flex items-center justify-center text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl shadow-sm transition-all cursor-pointer shrink-0"
                        >
                          {lang === 'ar' ? <ArrowRight size={18} /> : <ArrowLeft size={18} />}
                        </button>
                        <div className="text-right flex-1">
                          <span className="bg-purple-50 border border-purple-100/60 text-purple-600 font-extrabold text-[10px] tracking-widest uppercase px-3 py-1 rounded-full inline-flex items-center gap-1.5 mb-2">
                            <Users size={11} />
                            {lang === 'ar' ? 'المكاتب والوكلاء' : 'Offices & Agents'}
                          </span>
                          <h3 className="text-sm sm:text-base font-black text-slate-900 tracking-tight">
                            {lang === 'ar' ? 'معاملات المكاتب والوكلاء' : 'Office & Agent Transactions'}
                          </h3>
                        </div>
                      </div>

                      {/* Add office form */}
                      {isAddOfficeForm && (
                        <div className="bg-white border border-purple-100 rounded-[2rem] p-6 shadow-sm space-y-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                          <h4 className="text-sm font-black text-slate-900 text-right">{lang === 'ar' ? 'إضافة مكتب أو وكيل جديد' : 'Add New Office / Agent'}</h4>
                          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div className="space-y-2">
                              <label className="text-[10px] uppercase font-black text-slate-400 tracking-widest">{lang === 'ar' ? 'اسم المكتب أو الوكيل' : 'Office / Agent Name'}</label>
                              <input
                                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-4 focus:ring-purple-500/10 text-right"
                                placeholder={lang === 'ar' ? 'مثال: مكتب الأمل' : 'e.g. Al Amal Office'}
                                value={addOfficeForm.name}
                                onChange={e => setAddOfficeForm({ ...addOfficeForm, name: e.target.value })}
                              />
                            </div>
                            <div className="space-y-2">
                              <label className="text-[10px] uppercase font-black text-slate-400 tracking-widest">{lang === 'ar' ? 'رقم الهاتف أو الرقم المخصص' : 'Phone / Custom Number'}</label>
                              <input
                                className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-3 text-sm font-bold focus:outline-none focus:ring-4 focus:ring-purple-500/10 text-right font-mono"
                                placeholder={lang === 'ar' ? 'مثال: 771234567' : 'e.g. 771234567'}
                                value={addOfficeForm.phone}
                                onChange={e => setAddOfficeForm({ ...addOfficeForm, phone: e.target.value })}
                              />
                            </div>
                          </div>
                          <div className="flex justify-end gap-3">
                            <button
                              onClick={() => { setIsAddOfficeForm(false); setAddOfficeForm({ name: '', phone: '' }); }}
                              className="px-5 py-2.5 rounded-xl border border-slate-200 text-slate-600 font-bold text-xs cursor-pointer hover:bg-slate-50 transition-all"
                            >
                              {lang === 'ar' ? 'إلغاء' : 'Cancel'}
                            </button>
                            <button
                              onClick={() => {
                                if (!addOfficeForm.name.trim()) return;
                                const newOffice: Office = {
                                  id: Math.random().toString(36).substr(2, 9),
                                  name: addOfficeForm.name.trim(),
                                  phone: addOfficeForm.phone.trim()
                                };
                                setOffices([...offices, newOffice]);
                                setIsAddOfficeForm(false);
                                setAddOfficeForm({ name: '', phone: '' });
                              }}
                              className="px-6 py-2.5 bg-purple-600 hover:bg-purple-700 text-white rounded-xl font-bold text-xs flex items-center gap-2 shadow-lg shadow-purple-600/20 cursor-pointer transition-all"
                            >
                              <CheckCircle2 size={14} />
                              {lang === 'ar' ? 'حفظ' : 'Save'}
                            </button>
                          </div>
                        </div>
                      )}

                      {/* Cards grid */}
                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 xl:grid-cols-5 gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                        {/* Add office card — fixed first */}
                        <motion.button
                          whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                          onClick={() => { setIsAddOfficeForm(v => !v); setAddOfficeForm({ name: '', phone: '' }); }}
                          dir={lang === 'ar' ? 'rtl' : 'ltr'}
                          className={`border-2 border-dashed rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer h-full ${isAddOfficeForm ? 'border-purple-400 bg-purple-50' : 'border-slate-200 bg-white hover:border-purple-300 hover:bg-purple-50/40'}`}
                        >
                          <div className={`w-10 h-10 rounded-xl flex items-center justify-center mb-4 transition-all ${isAddOfficeForm ? 'bg-purple-100 text-purple-600' : 'bg-slate-100 text-slate-400'}`}>
                            {isAddOfficeForm ? <X size={18} /> : <Plus size={18} />}
                          </div>
                          <div className="w-full">
                            <span className="text-2xl font-black block leading-none text-slate-300">+</span>
                            <h4 className={`text-xs font-black tracking-tight mt-1.5 leading-tight ${isAddOfficeForm ? 'text-purple-600' : 'text-slate-400'}`}>
                              {isAddOfficeForm ? (lang === 'ar' ? 'إلغاء الإضافة' : 'Cancel') : (lang === 'ar' ? 'إضافة مكتب / وكيل' : 'Add Office / Agent')}
                            </h4>
                          </div>
                        </motion.button>

                        {/* General card — fixed second */}
                        <motion.button
                          whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                          onClick={() => { setSelectedOfficeFilter('general'); setIsOfficeDashboardView(true); }}
                          dir={lang === 'ar' ? 'rtl' : 'ltr'}
                          className="bg-slate-900 border border-slate-900 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-slate-900/10 h-full"
                        >
                          <div className="w-10 h-10 rounded-xl bg-slate-800 text-slate-100 flex items-center justify-center mb-4"><Globe size={20} /></div>
                          <div className="w-full">
                            <span className="text-2xl font-black block leading-none text-white">
                              {entries.filter(e => !e.archived && (!e.officeId || e.officeId === 'general')).length}
                            </span>
                            <h4 className="text-xs font-black text-slate-100 tracking-tight mt-1.5">{lang === 'ar' ? 'عام' : 'General'}</h4>
                          </div>
                        </motion.button>

                        {/* Office cards */}
                        {offices.map(office => (
                          (
                            /* ── Normal Card ── */
                            <motion.div
                              key={office.id}
                              whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                              dir={lang === 'ar' ? 'rtl' : 'ltr'}
                              onClick={() => { setSelectedOfficeFilter(office.id); setIsOfficeDashboardView(true); }}
                              className="bg-white border border-purple-100 hover:border-purple-200 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all shadow-sm group h-full relative cursor-pointer"
                            >
                              <div className="w-10 h-10 rounded-xl bg-purple-50 text-purple-600 group-hover:bg-purple-100 flex items-center justify-center mb-4 transition-all">
                                <Users size={18} />
                              </div>
                              <div className="w-full">
                                <span className="text-2xl font-black block leading-none text-slate-900">
                                  {entries.filter(e => !e.archived && e.officeId === office.id).length}
                                </span>
                                <h4 className="text-xs font-black text-slate-800 tracking-tight leading-tight mt-1.5">{office.name}</h4>
                                {office.phone && <p className="text-[10px] font-mono font-bold text-purple-400 mt-0.5">{office.phone}</p>}
                              </div>
                            </motion.div>
                          )
                        ))}
                      </div>
                    </div>
                    )
                  ) : isStatsView ? (
                    /* ── Statistics Sub-Page ── */
                    <div className="space-y-6">
                      <div className="bg-white p-6 sm:p-8 rounded-[2rem] border border-slate-200 shadow-sm flex items-center gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                        <button
                          onClick={() => setIsStatsView(false)}
                          className="w-10 h-10 flex items-center justify-center text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl shadow-sm transition-all cursor-pointer shrink-0"
                        >
                          {lang === 'ar' ? <ArrowRight size={18} /> : <ArrowLeft size={18} />}
                        </button>
                        <div className="text-right flex-1">
                          <span className="bg-slate-100 border border-slate-200/60 text-slate-600 font-extrabold text-[10px] tracking-widest uppercase px-3 py-1 rounded-full inline-flex items-center gap-1.5 mb-2">
                            <BarChart3 size={11} />
                            {lang === 'ar' ? 'الإحصائيات' : 'Statistics'}
                          </span>
                          <h3 className="text-sm sm:text-base font-black text-slate-900 tracking-tight">
                            {lang === 'ar' ? 'إحصائيات كافة المعاملات' : 'All Transactions Statistics'}
                          </h3>
                        </div>
                      </div>

                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                        {/* All */}
                        <motion.button
                          whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                          onClick={() => { setAllTabFilter('all'); setIsSectionSelected(true); }}
                          dir={lang === 'ar' ? 'rtl' : 'ltr'}
                          className="bg-slate-900 border border-slate-900 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-slate-900/10 h-full"
                        >
                          <div className="w-10 h-10 rounded-xl bg-slate-800 text-slate-100 flex items-center justify-center mb-4"><Globe size={20} /></div>
                          <div className="w-full">
                            <span className="text-2xl font-black block leading-none text-white">{entries.filter(e => !e.archived).length}</span>
                            <h4 className="text-xs font-black text-slate-100 tracking-tight mt-1.5">{lang === 'ar' ? 'كافة المعاملات' : 'All Transactions'}</h4>
                          </div>
                        </motion.button>

                        {/* Today */}
                        <motion.button
                          whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                          onClick={() => { setAllTabFilter('today'); setIsSectionSelected(true); }}
                          dir={lang === 'ar' ? 'rtl' : 'ltr'}
                          className="bg-indigo-600 border border-indigo-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-indigo-600/15 h-full"
                        >
                          <div className="w-10 h-10 rounded-xl bg-indigo-500 text-indigo-100 flex items-center justify-center mb-4"><RefreshCw size={20} className="animate-spin" /></div>
                          <div className="w-full">
                            <span className="text-2xl font-black block leading-none text-white">{entries.filter(e => !e.archived && isToday(e)).length}</span>
                            <h4 className="text-xs font-black text-indigo-50 tracking-tight mt-1.5">{lang === 'ar' ? 'تحديثات اليوم' : "Today's Updates"}</h4>
                          </div>
                        </motion.button>

                        {/* With print */}
                        <motion.button
                          whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                          onClick={() => { setAllTabFilter('with_print'); setIsSectionSelected(true); }}
                          dir={lang === 'ar' ? 'rtl' : 'ltr'}
                          className="bg-blue-600 border border-blue-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-blue-600/15 h-full"
                        >
                          <div className="w-10 h-10 rounded-xl bg-blue-500 text-blue-100 flex items-center justify-center mb-4"><FileSpreadsheet size={20} /></div>
                          <div className="w-full">
                            <span className="text-2xl font-black block leading-none text-white">{entries.filter(e => !e.archived && entryHasPrint(e) && !entryHasVisaAndImage(e)).length}</span>
                            <h4 className="text-xs font-black text-blue-50 tracking-tight mt-1.5">{lang === 'ar' ? 'المعاملات لها طلب ( برنت )' : 'With Print'}</h4>
                          </div>
                        </motion.button>

                        {/* Without print */}
                        <motion.button
                          whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                          onClick={() => { setAllTabFilter('without_print'); setIsSectionSelected(true); }}
                          dir={lang === 'ar' ? 'rtl' : 'ltr'}
                          className="bg-amber-600 border border-amber-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-amber-600/15 h-full"
                        >
                          <div className="w-10 h-10 rounded-xl bg-amber-500 text-amber-100 flex items-center justify-center mb-4"><FileX size={20} /></div>
                          <div className="w-full">
                            <span className="text-2xl font-black block leading-none text-white">{entries.filter(e => !e.archived && !entryHasPrint(e) && !entryHasVisaAndImage(e)).length}</span>
                            <h4 className="text-xs font-black text-amber-50 tracking-tight mt-1.5">{lang === 'ar' ? 'المعاملات التي لم يصدر لها برنت' : 'Without Print'}</h4>
                          </div>
                        </motion.button>

                        {/* Issued */}
                        <motion.button
                          whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                          onClick={() => { setAllTabFilter('issued'); setIsSectionSelected(true); }}
                          dir={lang === 'ar' ? 'rtl' : 'ltr'}
                          className="bg-emerald-600 border border-emerald-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-emerald-600/15 h-full"
                        >
                          <div className="w-10 h-10 rounded-xl bg-emerald-500 text-emerald-100 flex items-center justify-center mb-4"><CheckCircle2 size={20} /></div>
                          <div className="w-full">
                            <span className="text-2xl font-black block leading-none text-white">{entries.filter(e => !e.archived && (((e.statusText?.includes('تم اصدار') || e.statusText?.includes('Issued')) && entryHasPrint(e)) || entryHasVisaAndImage(e))).length}</span>
                            <h4 className="text-xs font-black text-emerald-50 tracking-tight mt-1.5">{lang === 'ar' ? 'المعاملات الجاهزة ( المؤشرة )' : 'Issued Transactions'}</h4>
                          </div>
                        </motion.button>

                        {/* Archive */}
                        <motion.button
                          whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                          onClick={() => { setIsArchiveView(true); setArchiveVisaTypeFilter(null); setIsStatsView(false); setIsOfficesView(false); setIsVisaTypeView(false); setIsSectionSelected(false); }}
                          dir={lang === 'ar' ? 'rtl' : 'ltr'}
                          className="bg-violet-600 border border-violet-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-violet-600/15 h-full"
                        >
                          <div className="w-10 h-10 rounded-xl bg-violet-500 text-violet-100 flex items-center justify-center mb-4"><Archive size={20} /></div>
                          <div className="w-full">
                            <span className="text-2xl font-black block leading-none text-white">{entries.filter(e => e.archived).length}</span>
                            <h4 className="text-xs font-black text-violet-50 tracking-tight mt-1.5">{lang === 'ar' ? 'الأرشيف' : 'Archive'}</h4>
                          </div>
                        </motion.button>

                        {/* Custom tx-status cards */}
                        {txStatusOptions.filter(opt => opt.hasStatsCard).map(opt => {
                          const cnt = entries.filter(e => !e.archived && (e.txStatus || '') === opt.label).length;
                          const isRed = opt.type === 'موقفة';
                          return (
                            <motion.button key={opt.id} whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                              onClick={() => { setAllTabFilter(`tx_status:${opt.label}`); setIsSectionSelected(true); }}
                              dir={lang === 'ar' ? 'rtl' : 'ltr'}
                              className={`${isRed ? 'bg-red-600 border-red-600 shadow-red-600/15' : 'bg-blue-600 border-blue-600 shadow-blue-600/15'} border rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl h-full`}
                            >
                              <div className={`w-10 h-10 rounded-xl ${isRed ? 'bg-red-500 text-red-100' : 'bg-blue-500 text-blue-100'} flex items-center justify-center mb-4`}><Tag size={20} /></div>
                              <div className="w-full">
                                <span className="text-2xl font-black block leading-none text-white">{cnt}</span>
                                <h4 className={`text-xs font-black ${isRed ? 'text-red-50' : 'text-blue-50'} tracking-tight mt-1.5 leading-tight`}>{opt.label}</h4>
                              </div>
                            </motion.button>
                          );
                        })}
                      </div>
                    </div>
                  ) : isArchiveView ? (() => {
                    const allArchived = entries.filter(e => e.archived);
                    const archivedEntries = archiveOfficeFilter
                      ? (archiveOfficeFilter === 'general'
                          ? allArchived.filter(e => !e.officeId || e.officeId === 'general')
                          : allArchived.filter(e => e.officeId === archiveOfficeFilter))
                      : allArchived;
                    const visaTypes = Array.from(new Set(archivedEntries.map(e => e.visaType))).filter(Boolean);
                    const displayEntries = archiveVisaTypeFilter
                      ? archivedEntries.filter(e => e.visaType === archiveVisaTypeFilter)
                      : archivedEntries;
                    const archiveOfficeName = archiveOfficeFilter
                      ? (archiveOfficeFilter === 'general'
                          ? (lang === 'ar' ? 'عام' : 'General')
                          : (offices.find(o => o.id === archiveOfficeFilter)?.name || ''))
                      : null;
                    return (
                      <div className="space-y-6">
                        {/* Header */}
                        <div className="bg-white p-6 sm:p-8 rounded-[2rem] border border-slate-200 shadow-sm flex items-center gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                          <button
                            onClick={() => {
                              if (archiveInCustomerView) {
                                setArchiveInCustomerView(false);
                                setArchiveVisaTypeFilter(null);
                              } else {
                                setIsArchiveView(false);
                                setArchiveVisaTypeFilter(null);
                                setArchiveInCustomerView(false);
                                if (archiveOfficeFilter) {
                                  setArchiveOfficeFilter(null);
                                  setIsOfficesView(true);
                                  setIsOfficeDashboardView(true);
                                  setSelectedOfficeFilter(archiveOfficeFilter);
                                }
                              }
                            }}
                            className="w-10 h-10 flex items-center justify-center text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl shadow-sm transition-all cursor-pointer shrink-0"
                          >
                            {lang === 'ar' ? <ArrowRight size={18} /> : <ArrowLeft size={18} />}
                          </button>
                          <div className="text-right flex-1">
                            <span className="bg-violet-100 border border-violet-200/60 text-violet-700 font-extrabold text-[10px] tracking-widest uppercase px-3 py-1 rounded-full inline-flex items-center gap-1.5 mb-2">
                              <Archive size={11} />
                              {lang === 'ar' ? 'الأرشيف' : 'Archive'}
                              {archiveOfficeName && <span className="text-violet-400">· {archiveOfficeName}</span>}
                            </span>
                            <h3 className="text-sm sm:text-base font-black text-slate-900 tracking-tight">
                              {archiveOfficeName
                                ? (lang === 'ar' ? `أرشيف ${archiveOfficeName}` : `${archiveOfficeName} Archive`)
                                : (lang === 'ar' ? 'المعاملات المؤرشفة' : 'Archived Transactions')}
                            </h3>
                          </div>
                        </div>

                        {/* Archive type cards — hidden when inside customer view */}
                        {!archiveInCustomerView && <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 xl:grid-cols-6 gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                          {/* All archived card */}
                          <motion.button
                            whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                            onClick={() => { setArchiveVisaTypeFilter(null); setArchiveInCustomerView(true); }}
                            dir={lang === 'ar' ? 'rtl' : 'ltr'}
                            className="rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer h-full border bg-white border-slate-200 hover:border-violet-200 hover:shadow-md"
                          >
                            <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-4 bg-violet-50 text-violet-600">
                              <Archive size={18} />
                            </div>
                            <div className="w-full">
                              <span className="text-2xl font-black block leading-none text-slate-900">{archivedEntries.length}</span>
                              <h4 className="text-xs font-black tracking-tight mt-1.5 text-slate-700">{lang === 'ar' ? 'كافة العملاء' : 'All Clients'}</h4>
                            </div>
                          </motion.button>

                          {/* Per visa type cards */}
                          {visaTypes.map(vt => {
                            const count = archivedEntries.filter(e => e.visaType === vt).length;
                            return (
                              <motion.button
                                key={vt}
                                whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                                onClick={() => { setArchiveVisaTypeFilter(vt); setArchiveInCustomerView(true); }}
                                dir={lang === 'ar' ? 'rtl' : 'ltr'}
                                className="rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer h-full border bg-white border-slate-200 hover:border-violet-200 hover:shadow-md"
                              >
                                <div className="w-10 h-10 rounded-xl flex items-center justify-center mb-4 bg-slate-50 text-slate-500">
                                  <FileText size={18} />
                                </div>
                                <div className="w-full">
                                  <span className="text-2xl font-black block leading-none text-slate-900">{count}</span>
                                  <h4 className="text-xs font-black tracking-tight mt-1.5 leading-tight text-slate-700">{vt}</h4>
                                </div>
                              </motion.button>
                            );
                          })}
                        </div>}

                        {/* Archived entries list — shown only after clicking a section card */}
                        {archiveInCustomerView && (() => {
                          const filteredArchiveEntries = archiveSearchQuery.trim()
                            ? displayEntries.filter(e =>
                                (e.applicantName || '').toLowerCase().includes(archiveSearchQuery.toLowerCase()) ||
                                (e.passportNumber || '').toLowerCase().includes(archiveSearchQuery.toLowerCase()) ||
                                (e.visaNumber || '').toLowerCase().includes(archiveSearchQuery.toLowerCase())
                              )
                            : displayEntries;
                          const allFilteredIds = filteredArchiveEntries.map(e => e.id);
                          const allSelected = allFilteredIds.length > 0 && allFilteredIds.every(id => archiveSelectedIds.has(id));
                          const toggleSelectAll = () => {
                            if (allSelected) {
                              setArchiveSelectedIds(new Set());
                            } else {
                              setArchiveSelectedIds(new Set(allFilteredIds));
                            }
                          };
                          const toggleSelectOne = (id: string) => {
                            setArchiveSelectedIds(prev => {
                              const next = new Set(prev);
                              next.has(id) ? next.delete(id) : next.add(id);
                              return next;
                            });
                          };
                          const handleBulkRestore = () => {
                            const ids = Array.from(archiveSelectedIds);
                            showConfirm(
                              lang === 'ar'
                                ? `هل تريد استرداد ${ids.length} معاملة من الأرشيف وإعادتها للقائمة الرئيسية؟`
                                : `Restore ${ids.length} entries from archive to the main list?`,
                              () => {
                                setEntries(prev => prev.map(e => ids.includes(e.id) ? { ...e, archived: false, archivedAt: undefined } : e));
                                addSyncLog(lang === 'ar' ? `تم استرجاع ${ids.length} معاملة من الأرشيف.` : `${ids.length} entries restored from archive.`);
                                setArchiveSelectedIds(new Set());
                                setArchiveSelectMode(false);
                              }
                            );
                          };
                          const handleBulkDelete = () => {
                            const ids = Array.from(archiveSelectedIds);
                            showConfirm(
                              lang === 'ar'
                                ? `هل أنت متأكد من الحذف النهائي لـ ${ids.length} معاملة؟ لا يمكن التراجع عن هذا الإجراء.`
                                : `Permanently delete ${ids.length} entries? This cannot be undone.`,
                              () => {
                                setEntries(prev => {
                                  const updated = prev.filter(e => !ids.includes(e.id));
                                  fetch('/api/entries', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ entries: updated }) }).catch(() => {});
                                  return updated;
                                });
                                addSyncLog(lang === 'ar' ? `تم الحذف النهائي لـ ${ids.length} معاملة من الأرشيف.` : `${ids.length} entries permanently deleted.`);
                                setArchiveSelectedIds(new Set());
                                setArchiveSelectMode(false);
                              }
                            );
                          };
                          return (
                        <div className="space-y-4">
                          {/* Sub-header with back button + search */}
                          <div className="bg-white p-4 sm:p-5 rounded-2xl border border-slate-200 shadow-sm flex flex-col gap-3" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                            <div className="flex items-center gap-3">
                              <button
                                onClick={() => { setArchiveInCustomerView(false); setArchiveVisaTypeFilter(null); setArchiveSearchQuery(''); setArchiveSelectMode(false); setArchiveSelectedIds(new Set()); }}
                                className="w-9 h-9 flex items-center justify-center text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl shadow-sm transition-all cursor-pointer shrink-0"
                              >
                                {lang === 'ar' ? <ArrowRight size={16} /> : <ArrowLeft size={16} />}
                              </button>
                              <div className="text-right flex-1">
                                <p className="text-xs font-black text-slate-800">
                                  {archiveVisaTypeFilter ? archiveVisaTypeFilter : (lang === 'ar' ? 'كافة العملاء' : 'All Clients')}
                                </p>
                                <p className="text-[10px] font-bold text-slate-400">{filteredArchiveEntries.length} {lang === 'ar' ? 'معاملة' : 'entries'}</p>
                              </div>
                              {/* Select mode toggle */}
                              <button
                                onClick={() => { setArchiveSelectMode(v => !v); setArchiveSelectedIds(new Set()); }}
                                className={`inline-flex items-center gap-1.5 text-[11px] font-black px-3 py-1.5 rounded-xl border transition-all cursor-pointer shrink-0 ${archiveSelectMode ? 'bg-violet-600 text-white border-violet-600' : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-violet-50 hover:text-violet-700 hover:border-violet-200'}`}
                              >
                                <CheckSquare size={13} />
                                {archiveSelectMode ? (lang === 'ar' ? 'إلغاء' : 'Cancel') : (lang === 'ar' ? 'تحديد' : 'Select')}
                              </button>
                            </div>

                            {/* Select-all row — shown only in select mode */}
                            {archiveSelectMode && (
                              <div className="flex items-center justify-between gap-2 pt-1 border-t border-slate-100">
                                <button
                                  onClick={toggleSelectAll}
                                  className="inline-flex items-center gap-1.5 text-[11px] font-black text-violet-600 hover:text-violet-800 cursor-pointer"
                                >
                                  <div className={`w-4 h-4 rounded border-2 flex items-center justify-center transition-colors ${allSelected ? 'bg-violet-600 border-violet-600' : 'border-slate-300'}`}>
                                    {allSelected && <Check size={10} className="text-white" />}
                                  </div>
                                  {lang === 'ar' ? 'تحديد الكل' : 'Select All'}
                                </button>
                                <span className="text-[10px] font-bold text-slate-400">
                                  {archiveSelectedIds.size} {lang === 'ar' ? 'محدد' : 'selected'}
                                </span>
                              </div>
                            )}

                            {/* Search box */}
                            <div className="relative">
                              <Search className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400" size={14} />
                              <input
                                type="text"
                                value={archiveSearchQuery}
                                onChange={e => setArchiveSearchQuery(e.target.value)}
                                placeholder={lang === 'ar' ? 'ابحث بالاسم أو الجواز أو رقم التأشيرة...' : 'Search by name, passport or visa...'}
                                className="w-full bg-slate-50 border border-slate-200 rounded-xl py-2 pr-9 pl-3 text-sm text-right focus:outline-none focus:ring-2 focus:ring-violet-500/20 transition-all font-medium"
                                dir="rtl"
                              />
                            </div>
                          </div>

                          {/* Bulk action bar */}
                          <AnimatePresence>
                          {archiveSelectMode && archiveSelectedIds.size > 0 && (
                            <motion.div
                              initial={{ opacity: 0, y: -8 }}
                              animate={{ opacity: 1, y: 0 }}
                              exit={{ opacity: 0, y: -8 }}
                              transition={{ duration: 0.2 }}
                              className="bg-violet-50 border border-violet-200 rounded-2xl px-4 py-3 flex items-center justify-between gap-3" dir={lang === 'ar' ? 'rtl' : 'ltr'}
                            >
                              <span className="text-xs font-black text-violet-700">
                                {lang === 'ar' ? `تم تحديد ${archiveSelectedIds.size} معاملة` : `${archiveSelectedIds.size} selected`}
                              </span>
                              <div className="flex items-center gap-2">
                                <button
                                  onClick={handleBulkRestore}
                                  className="inline-flex items-center gap-1.5 text-[11px] font-black bg-violet-600 hover:bg-violet-700 text-white px-3 py-1.5 rounded-xl transition-all cursor-pointer shadow-sm"
                                >
                                  <RotateCcw size={12} />
                                  {lang === 'ar' ? 'استرداد' : 'Restore'}
                                </button>
                                <button
                                  onClick={handleBulkDelete}
                                  className="inline-flex items-center gap-1.5 text-[11px] font-black bg-red-500 hover:bg-red-600 text-white px-3 py-1.5 rounded-xl transition-all cursor-pointer shadow-sm"
                                >
                                  <Trash2 size={12} />
                                  {lang === 'ar' ? 'حذف نهائي' : 'Delete'}
                                </button>
                              </div>
                            </motion.div>
                          )}
                          </AnimatePresence>
                          {filteredArchiveEntries.length === 0 ? (
                          <div className="bg-white rounded-[2rem] border border-slate-200 p-12 text-center" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                            <Archive size={36} className="text-slate-300 mx-auto mb-3" />
                            <p className="text-slate-400 font-bold text-sm">{lang === 'ar' ? 'لا توجد معاملات مؤرشفة' : 'No archived transactions'}</p>
                          </div>
                        ) : (
                          <div className="space-y-3">
                            {/* Cards grid */}
                            <AnimatePresence>
                            <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
                              {filteredArchiveEntries.map((entry, idx) => {
                                const isExpanded = expandedCardId === entry.id;
                                const isChecked = archiveSelectedIds.has(entry.id);
                                const _offObj = offices.find(o => o.id === entry.officeId);
                                const _archOffName = _offObj ? _offObj.name : (lang === 'ar' ? 'عام' : 'General');
                                return (
                                  <motion.div
                                    key={entry.id}
                                    initial={{ opacity: 0, scale: 0.97 }}
                                    animate={{ opacity: 1, scale: 1 }}
                                    exit={{ opacity: 0, scale: 0.97 }}
                                    className={`bg-white rounded-3xl border shadow-sm overflow-hidden transition-all duration-200 ${
                                      archiveSelectMode && isChecked
                                        ? 'border-violet-400 shadow-md ring-2 ring-violet-200'
                                        : isExpanded
                                        ? 'border-violet-300 shadow-md col-span-1 sm:col-span-2 xl:col-span-3'
                                        : 'border-violet-200 hover:border-violet-300 hover:shadow-md'
                                    }`}
                                  >
                                    {/* Card Header */}
                                    <div
                                      className="p-5 cursor-pointer select-none"
                                      onClick={() => archiveSelectMode ? toggleSelectOne(entry.id) : setExpandedCardId(isExpanded ? null : entry.id)}
                                    >
                                      {/* Top row: Archive badge + checkbox */}
                                      <div className="flex justify-between items-center mb-2">
                                        <div>
                                          {archiveSelectMode && (
                                            <div className={`w-5 h-5 rounded-full border-2 flex items-center justify-center transition-colors ${isChecked ? 'bg-violet-600 border-violet-600' : 'border-slate-300 bg-white'}`}>
                                              {isChecked && <Check size={11} className="text-white" />}
                                            </div>
                                          )}
                                        </div>
                                        <span className="inline-flex items-center gap-1 text-[10px] font-black bg-violet-100 text-violet-700 border border-violet-200 rounded-full px-2.5 py-0.5">
                                          <Archive size={9} />
                                          {lang === 'ar' ? 'مؤرشف' : 'Archived'}
                                        </span>
                                      </div>
                                      {/* Name + number */}
                                      <div className="flex items-start justify-between gap-3 mb-4">
                                        <div className="min-w-0 text-right flex-1">
                                          <p className="text-sm font-black text-slate-900 truncate leading-tight">{entry.applicantName || entry.nameEnglish || (lang === 'ar' ? 'غير محدد' : '—')}</p>
                                          <p className="text-[10px] font-mono font-bold text-slate-400 mt-0.5 truncate">{entry.passportNumber || '—'}</p>
                                        </div>
                                        <span className="text-[9px] font-black text-slate-300 font-mono shrink-0 mt-1">#{idx + 1}</span>
                                      </div>
                                      {/* Badges */}
                                      <div className="flex flex-wrap gap-1.5 justify-end mb-3">
                                        {entry.visaType && (
                                          <span className="inline-flex items-center gap-1 text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-100 rounded-full px-2.5 py-0.5">
                                            <Tag size={9} />{entry.visaType}
                                          </span>
                                        )}
                                        <span className="inline-flex items-center gap-1 text-[10px] font-bold bg-violet-50 text-violet-700 border border-violet-100 rounded-full px-2.5 py-0.5">
                                          <Briefcase size={9} />{_archOffName}
                                        </span>
                                        {entry.applicantData && entry.applicantData.includes('تاريخ الاستلام:') && (
                                          <span className="inline-flex items-center gap-1 text-[10px] font-bold bg-sky-50 text-sky-700 border border-sky-100 rounded-full px-2.5 py-0.5">
                                            <CalendarDays size={9} />{entry.applicantData.replace('تاريخ الاستلام:', '').trim()}
                                          </span>
                                        )}
                                      </div>
                                      {/* Actions row */}
                                      <div className="flex items-center justify-between gap-2 pt-3 border-t border-slate-100" onClick={(e) => e.stopPropagation()}>
                                        <div className="flex items-center gap-1">
                                          {/* Restore */}
                                          <button
                                            onClick={() => {
                                              setEntries(prev => prev.map(e => e.id === entry.id ? { ...e, archived: false, archivedAt: undefined } : e));
                                              addSyncLog(lang === 'ar' ? `تم استرجاع العميل ${entry.applicantName} من الأرشيف.` : `Client ${entry.applicantName} restored from archive.`);
                                            }}
                                            className="inline-flex items-center gap-1 text-[10px] font-black text-violet-600 hover:text-violet-800 bg-violet-50 hover:bg-violet-100 border border-violet-100 px-2.5 py-1.5 rounded-xl transition-all cursor-pointer"
                                          >
                                            <RotateCcw size={11} />
                                            {lang === 'ar' ? 'استرجاع' : 'Restore'}
                                          </button>
                                          {/* Delete */}
                                          <button
                                            onClick={() => setArchiveConfirmEntry(entry)}
                                            className="p-1.5 rounded-xl hover:bg-red-50 hover:text-red-500 transition-all cursor-pointer text-slate-400 border border-transparent hover:border-red-100"
                                            title={lang === 'ar' ? 'حذف نهائي' : 'Delete'}
                                          >
                                            <Trash2 size={13} />
                                          </button>
                                        </div>
                                        <motion.span
                                          animate={{ rotate: isExpanded ? 180 : 0 }}
                                          transition={{ duration: 0.2 }}
                                          className="text-slate-300"
                                          onClick={() => setExpandedCardId(isExpanded ? null : entry.id)}
                                        >
                                          <ChevronDown size={14} />
                                        </motion.span>
                                      </div>
                                    </div>

                                    {/* Expanded Details */}
                                    <AnimatePresence initial={false}>
                                      {isExpanded && (
                                        <motion.div
                                          key="expanded"
                                          initial={{ height: 0, opacity: 0 }}
                                          animate={{ height: 'auto', opacity: 1 }}
                                          exit={{ height: 0, opacity: 0 }}
                                          transition={{ duration: 0.25 }}
                                          style={{ overflow: 'hidden' }}
                                        >
                                          <div className="border-t border-slate-100 p-5" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                                            <div className="space-y-4">
                                              {/* Passport */}
                                              <div className="space-y-1 text-right overflow-hidden">
                                                <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">{lang === 'ar' ? 'رقم جواز السفر' : 'Passport'}</label>
                                                <input
                                                  className="w-full bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-mono font-black text-slate-900 uppercase text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all"
                                                  value={entry.passportNumber}
                                                  disabled={!hasPermission('edit_clients')}
                                                  onChange={(e) => updateEntryField(entry.id, 'passportNumber', e.target.value)}
                                                />
                                              </div>
                                              {/* Phone */}
                                              <div className="space-y-1 text-right overflow-hidden">
                                                <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">{lang === 'ar' ? 'الموبايل 📞' : 'Telephone 📞'}</label>
                                                <input
                                                  className="w-full bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-900 text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all font-mono"
                                                  value={entry.phoneNumber || ''}
                                                  placeholder={lang === 'ar' ? 'رقم الهاتف' : 'No phone'}
                                                  disabled={!hasPermission('edit_clients')}
                                                  onChange={(e) => updateEntryField(entry.id, 'phoneNumber', e.target.value)}
                                                />
                                              </div>
                                              {/* Nationality & Arrival Port */}
                                              <div className="grid grid-cols-2 gap-3">
                                                <div className="space-y-1 text-right overflow-hidden">
                                                  <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">{lang === 'ar' ? 'الجنسية' : 'Nationality'}</label>
                                                  <input
                                                    className="w-full bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-900 text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all"
                                                    value={entry.nationality}
                                                    disabled={!hasPermission('edit_clients')}
                                                    onChange={(e) => updateEntryField(entry.id, 'nationality', e.target.value)}
                                                  />
                                                </div>
                                                <div className="space-y-1 text-right overflow-hidden">
                                                  <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">{lang === 'ar' ? 'جهة القدوم' : 'Arrival Port'}</label>
                                                  <input
                                                    className="w-full bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-900 text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all"
                                                    value={entry.arrivalPoint}
                                                    disabled={!hasPermission('edit_clients')}
                                                    onChange={(e) => updateEntryField(entry.id, 'arrivalPoint', e.target.value)}
                                                  />
                                                </div>
                                              </div>
                                              {/* Visa Type */}
                                              <div className="space-y-1 text-right overflow-hidden">
                                                <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">{lang === 'ar' ? 'نوع التأشيرة' : 'Visa Type'}</label>
                                                <input
                                                  className="w-full bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-900 text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all"
                                                  value={entry.visaType}
                                                  disabled={!hasPermission('edit_clients')}
                                                  onChange={(e) => updateEntryField(entry.id, 'visaType', e.target.value)}
                                                />
                                              </div>
                                              {/* MOFA Extended Data */}
                                              <div className="border-t border-slate-100 pt-3 space-y-2">
                                                <span className="text-[8px] font-black text-emerald-600 uppercase tracking-widest bg-emerald-50 border border-emerald-100 px-2 py-0.5 rounded-full">بيانات منصة التأشيرات</span>
                                                <div className="grid grid-cols-3 gap-2 mb-2">
                                                  <div className="space-y-0.5 text-right overflow-hidden">
                                                    <label className="text-[8px] font-black text-slate-500 uppercase tracking-wider block truncate">📋 رقم الطلب</label>
                                                    <div className="w-full bg-slate-100 border border-slate-200 rounded-lg px-2 py-1.5 text-[11px] font-black text-slate-800 text-center truncate min-h-[28px] flex items-center justify-center font-mono">
                                                      {entry.applicationNumber && entry.applicationNumber !== '---' ? entry.applicationNumber : <span className="text-slate-300">—</span>}
                                                    </div>
                                                  </div>
                                                  <div className="space-y-0.5 text-right overflow-hidden">
                                                    <label className="text-[8px] font-black text-emerald-600 uppercase tracking-wider block truncate">🎫 رقم التأشيرة</label>
                                                    <div className="w-full bg-emerald-50 border border-emerald-200 rounded-lg px-2 py-1.5 text-[11px] font-black text-emerald-800 text-center truncate min-h-[28px] flex items-center justify-center font-mono">
                                                      {entry.visaNumber && entry.visaNumber !== '---' ? entry.visaNumber : <span className="text-slate-300">—</span>}
                                                    </div>
                                                  </div>
                                                  <div className="space-y-0.5 text-right overflow-hidden">
                                                    <label className="text-[8px] font-black text-amber-600 uppercase tracking-wider block truncate">📄 رقم المستند</label>
                                                    <div className="w-full bg-amber-50 border border-amber-200 rounded-lg px-2 py-1.5 text-[11px] font-black text-amber-900 text-center truncate min-h-[28px] flex items-center justify-center font-mono">
                                                      {entry.documentNumber && entry.documentNumber !== '' ? entry.documentNumber : <span className="text-slate-300">—</span>}
                                                    </div>
                                                  </div>
                                                </div>
                                                <div className="grid grid-cols-2 gap-2">
                                                  {([
                                                    { label: 'تاريخ الطلب', key: 'applicationDate' },
                                                    { label: 'الممثلية في', key: 'embassy' },
                                                    { label: 'عدد مرات الدخول', key: 'entriesCount' },
                                                    { label: 'اسم الجهة الطالبة', key: 'requesterName' },
                                                    { label: 'الاسم', key: 'applicantName' },
                                                    { label: 'Name', key: 'nameEnglish' },
                                                    { label: 'نوع الجواز', key: 'passportType' },
                                                    { label: 'تاريخ الانتهاء', key: 'passportExpiry' },
                                                    { label: 'تاريخ الميلاد', key: 'birthDate' },
                                                    { label: 'مكان الميلاد', key: 'birthPlace' },
                                                    { label: 'الجنسية الحالية', key: 'currentNationality' },
                                                    { label: 'الجنس', key: 'gender' },
                                                    { label: 'المهنة', key: 'profession' },
                                                    { label: 'الغرض', key: 'purpose' },
                                                    { label: 'حالة الطلب', key: 'statusText' },
                                                    { label: 'حالة اصدار الشهادة الصحية', key: 'healthCertStatus' },
                                                  ] as { label: string; key: keyof typeof entry }[]).map(f => (
                                                    <div key={f.key} className="space-y-0.5 text-right overflow-hidden">
                                                      <label className="text-[8px] font-black text-slate-400 uppercase tracking-wider block truncate">{f.label}</label>
                                                      <div className="w-full bg-slate-50 border border-slate-100 rounded-lg px-2 py-1.5 text-[10px] font-bold text-slate-700 text-center truncate min-h-[28px] flex items-center justify-center">
                                                        {(entry[f.key] as string) || <span className="text-slate-300">—</span>}
                                                      </div>
                                                    </div>
                                                  ))}
                                                </div>
                                              </div>
                                              {/* ── Images Upload Section (Archive) ── */}
                                              <div className="border-t border-slate-100 pt-3 space-y-3">
                                                <span className="text-[8px] font-black text-blue-600 uppercase tracking-widest bg-blue-50 border border-blue-100 px-2 py-0.5 rounded-full">
                                                  {lang === 'ar' ? 'المستندات والصور' : 'Documents & Images'}
                                                </span>
                                                <div className="grid grid-cols-2 gap-3">
                                                  {/* Print Image — download only (auto-captured from MOFA print button) */}
                                                  <div className="space-y-1.5">
                                                    <p className="text-[9px] font-black text-slate-500 text-right">🖨️ {lang === 'ar' ? 'صورة البرنت' : 'Print Image'}</p>
                                                    <button
                                                      onClick={(e) => { e.stopPropagation(); if (entry.printImageUrl) { downloadUrl(entry.printImageUrl, `برنت_${entry.passportNumber||''}_${(entry.applicantName||'').replace(/\s+/g,'_')}.pdf`); } }}
                                                      className={`flex flex-col items-center justify-center gap-1.5 w-full border-2 rounded-xl h-28 transition-all group ${entry.printImageUrl ? 'border-blue-300 hover:border-blue-500 bg-blue-50/60 hover:bg-blue-100 cursor-pointer' : 'border-blue-100 bg-blue-50/20 cursor-not-allowed opacity-60'}`}
                                                      title={entry.printImageUrl ? (lang === 'ar' ? 'تنزيل PDF البرنت' : 'Download Print PDF') : (lang === 'ar' ? 'لم يتم جلب الملف بعد' : 'File not fetched yet')}
                                                    >
                                                      {entry.printImageUrl?.startsWith('tg:') ? <span className="text-lg">📨</span> : <Download size={20} className={`transition-colors ${entry.printImageUrl ? 'text-blue-500 group-hover:text-blue-700' : 'text-blue-200'}`} />}
                                                      <span className={`text-[9px] font-bold transition-colors text-center px-1 ${entry.printImageUrl ? 'text-blue-500 group-hover:text-blue-700' : 'text-blue-200'}`}>
                                                        {entry.printImageUrl ? (lang === 'ar' ? 'تنزيل PDF البرنت' : 'Download Print PDF') : (lang === 'ar' ? 'لم يتم الجلب بعد' : 'Not fetched yet')}
                                                      </span>
                                                    </button>
                                                  </div>
                                                  {/* Visa Image */}
                                                  <div className="space-y-1.5">
                                                    <p className="text-[9px] font-black text-slate-500 text-right">🎫 {lang === 'ar' ? 'صورة التأشيرة' : 'Visa Image'}</p>
                                                    <button
                                                      onClick={(e) => {
                                                        e.stopPropagation();
                                                        const visaUrl = entry.visaImageUrl || entry.visaImageBase64;
                                                        if (visaUrl) {
                                                          downloadUrl(visaUrl, `تاشيره_${entry.passportNumber||''}_${(entry.applicantName||'').replace(/\s+/g,'_')}.pdf`);
                                                        }
                                                        // الجلب يتم تلقائياً عند التشييك — لا يُفتح المودال يدوياً
                                                      }}
                                                      className={`flex flex-col items-center justify-center gap-1.5 w-full border-2 rounded-xl h-28 transition-all group ${(entry.visaImageUrl||entry.visaImageBase64) ? 'border-emerald-300 hover:border-emerald-500 bg-emerald-50/60 hover:bg-emerald-100 cursor-pointer' : 'border-slate-100 bg-slate-50/40 cursor-not-allowed opacity-50'}`}
                                                    >
                                                      {(entry.visaImageUrl||entry.visaImageBase64) ? <Download size={20} className="text-emerald-500 group-hover:text-emerald-700 transition-colors" /> : <Download size={20} className="text-slate-300" />}
                                                      <span className={`text-[9px] font-bold transition-colors text-center px-1 ${(entry.visaImageUrl||entry.visaImageBase64) ? 'text-emerald-500 group-hover:text-emerald-700' : 'text-slate-300'}`}>
                                                        {(entry.visaImageUrl||entry.visaImageBase64) ? (lang === 'ar' ? 'تنزيل PDF التأشيرة' : 'Download Visa PDF') : (lang === 'ar' ? 'تُجلب تلقائياً عند التشييك' : 'Auto-fetched on check')}
                                                      </span>
                                                    </button>
                                                  </div>
                                                </div>
                                              </div>
                                            </div>
                                          </div>
                                        </motion.div>
                                      )}
                                    </AnimatePresence>
                                  </motion.div>
                                );
                              })}
                            </div>
                            </AnimatePresence>
                          </div>
                        )}
                        </div>
                          );
                        })()}
                      </div>
                    );
                  })() : isVisaTypeView ? (
                    /* ── Visa Types Sub-Page ── */
                    <div className="space-y-6">
                      {/* Header + back button */}
                      <div className="bg-white p-6 sm:p-8 rounded-[2rem] border border-slate-200 shadow-sm flex items-center gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                        <button
                          onClick={() => {
                            if (selectedVisaType) {
                              setSelectedVisaType(null);
                            } else {
                              setIsVisaTypeView(false);
                            }
                          }}
                          className="w-10 h-10 flex items-center justify-center text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl shadow-sm transition-all cursor-pointer shrink-0"
                        >
                          {lang === 'ar' ? <ArrowRight size={18} /> : <ArrowLeft size={18} />}
                        </button>
                        <div className="text-right flex-1">
                          <span className="bg-blue-50 border border-blue-100/60 text-blue-600 font-extrabold text-[10px] tracking-widest uppercase px-3 py-1 rounded-full inline-flex items-center gap-1.5 mb-2">
                            <Tag size={11} />
                            {selectedVisaType
                              ? selectedVisaType
                              : (lang === 'ar' ? 'المعاملات' : 'Transactions')}
                          </span>
                          <h3 className="text-sm sm:text-base font-black text-slate-900 tracking-tight">
                            {selectedVisaType
                              ? (lang === 'ar' ? `تأشيرة ${selectedVisaType}` : `Visa: ${selectedVisaType}`)
                              : (lang === 'ar' ? 'المعاملات حسب نوع التأشيرة' : 'Transactions by Visa Type')}
                          </h3>
                        </div>
                      </div>

                      {selectedVisaType ? (
                        /* ── Sub-cards for selected visa type ── */
                        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                          {/* All */}
                          <motion.button
                            whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                            onClick={() => {
                              setAllTabFilter('all');
                              setIsSectionSelected(true);
                            }}
                            dir={lang === 'ar' ? 'rtl' : 'ltr'}
                            className="bg-slate-900 border border-slate-900 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-slate-900/10 h-full"
                          >
                            <div className="w-10 h-10 rounded-xl bg-slate-800 text-slate-100 flex items-center justify-center mb-4">
                              <Globe size={20} />
                            </div>
                            <div className="w-full">
                              <span className="text-2xl font-black block leading-none text-white">
                                {entries.filter(e => !e.archived && e.visaType === selectedVisaType).length}
                              </span>
                              <h4 className="text-xs font-black text-slate-100 tracking-tight mt-1.5">
                                {lang === 'ar' ? 'كافة المعاملات' : 'All Transactions'}
                              </h4>
                            </div>
                          </motion.button>

                          {/* With print */}
                          <motion.button
                            whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                            onClick={() => {
                              setAllTabFilter('with_print');
                              setIsSectionSelected(true);
                            }}
                            dir={lang === 'ar' ? 'rtl' : 'ltr'}
                            className="bg-blue-600 border border-blue-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-blue-600/15 h-full"
                          >
                            <div className="w-10 h-10 rounded-xl bg-blue-500 text-blue-100 flex items-center justify-center mb-4">
                              <FileSpreadsheet size={20} />
                            </div>
                            <div className="w-full">
                              <span className="text-2xl font-black block leading-none text-white">
                                {entries.filter(e => !e.archived && e.visaType === selectedVisaType && entryHasPrint(e) && !entryHasVisaAndImage(e)).length}
                              </span>
                              <h4 className="text-xs font-black text-blue-50 tracking-tight mt-1.5">
                                {lang === 'ar' ? 'المعاملات لها طلب ( برنت )' : 'With Print'}
                              </h4>
                            </div>
                          </motion.button>

                          {/* Without print */}
                          <motion.button
                            whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                            onClick={() => {
                              setAllTabFilter('without_print');
                              setIsSectionSelected(true);
                            }}
                            dir={lang === 'ar' ? 'rtl' : 'ltr'}
                            className="bg-amber-600 border border-amber-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-amber-600/15 h-full"
                          >
                            <div className="w-10 h-10 rounded-xl bg-amber-500 text-amber-100 flex items-center justify-center mb-4">
                              <FileX size={20} />
                            </div>
                            <div className="w-full">
                              <span className="text-2xl font-black block leading-none text-white">
                                {entries.filter(e => !e.archived && e.visaType === selectedVisaType && !entryHasPrint(e) && !entryHasVisaAndImage(e)).length}
                              </span>
                              <h4 className="text-xs font-black text-amber-50 tracking-tight mt-1.5">
                                {lang === 'ar' ? 'المعاملات التي لم يصدر لها برنت' : 'Without Print'}
                              </h4>
                            </div>
                          </motion.button>

                          {/* Issued */}
                          <motion.button
                            whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                            onClick={() => {
                              setAllTabFilter('issued');
                              setIsSectionSelected(true);
                            }}
                            dir={lang === 'ar' ? 'rtl' : 'ltr'}
                            className="bg-emerald-600 border border-emerald-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-emerald-600/15 h-full"
                          >
                            <div className="w-10 h-10 rounded-xl bg-emerald-500 text-emerald-100 flex items-center justify-center mb-4">
                              <CheckCircle2 size={20} />
                            </div>
                            <div className="w-full">
                              <span className="text-2xl font-black block leading-none text-white">
                                {entries.filter(e => !e.archived && e.visaType === selectedVisaType && (((e.statusText?.includes('تم اصدار') || e.statusText?.includes('Issued')) && entryHasPrint(e)) || entryHasVisaAndImage(e))).length}
                              </span>
                              <h4 className="text-xs font-black text-emerald-50 tracking-tight mt-1.5">
                                {lang === 'ar' ? 'المعاملات الجاهزة ( المؤشرة )' : 'Issued Transactions'}
                              </h4>
                            </div>
                          </motion.button>

                          {/* Today */}
                          <motion.button
                            whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                            onClick={() => {
                              setAllTabFilter('today');
                              setIsSectionSelected(true);
                            }}
                            dir={lang === 'ar' ? 'rtl' : 'ltr'}
                            className="bg-indigo-600 border border-indigo-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-indigo-600/15 h-full"
                          >
                            <div className="w-10 h-10 rounded-xl bg-indigo-500 text-indigo-100 flex items-center justify-center mb-4">
                              <RefreshCw size={20} className="animate-spin" />
                            </div>
                            <div className="w-full">
                              <span className="text-2xl font-black block leading-none text-white">
                                {entries.filter(e => !e.archived && e.visaType === selectedVisaType && isToday(e)).length}
                              </span>
                              <h4 className="text-xs font-black text-indigo-50 tracking-tight mt-1.5">
                                {lang === 'ar' ? 'تحديثات اليوم' : "Today's Updates"}
                              </h4>
                            </div>
                          </motion.button>

                          {/* Archive */}
                          <motion.button
                            whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                            onClick={() => { setIsArchiveView(true); setArchiveVisaTypeFilter(null); setIsStatsView(false); setIsOfficesView(false); setIsVisaTypeView(false); setIsSectionSelected(false); }}
                            dir={lang === 'ar' ? 'rtl' : 'ltr'}
                            className="bg-violet-600 border border-violet-600 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl shadow-violet-600/15 h-full"
                          >
                            <div className="w-10 h-10 rounded-xl bg-violet-500 text-violet-100 flex items-center justify-center mb-4">
                              <Archive size={20} />
                            </div>
                            <div className="w-full">
                              <span className="text-2xl font-black block leading-none text-white">{entries.filter(e => e.archived).length}</span>
                              <h4 className="text-xs font-black text-violet-50 tracking-tight mt-1.5">
                                {lang === 'ar' ? 'الأرشيف' : 'Archive'}
                              </h4>
                            </div>
                          </motion.button>

                          {/* Custom tx-status cards filtered by visa type */}
                          {txStatusOptions.filter(opt => opt.hasStatsCard).map(opt => {
                            const cnt = entries.filter(e => !e.archived && e.visaType === selectedVisaType && (e.txStatus || '') === opt.label).length;
                            const isRed = opt.type === 'موقفة';
                            return (
                              <motion.button key={opt.id} whileHover={{ scale: 1.02 }} whileTap={{ scale: 0.98 }}
                                onClick={() => { setAllTabFilter(`tx_status:${opt.label}`); setIsSectionSelected(true); }}
                                dir={lang === 'ar' ? 'rtl' : 'ltr'}
                                className={`${isRed ? 'bg-red-600 border-red-600 shadow-red-600/15' : 'bg-blue-600 border-blue-600 shadow-blue-600/15'} border rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer relative overflow-hidden shadow-xl h-full`}
                              >
                                <div className={`w-10 h-10 rounded-xl ${isRed ? 'bg-red-500 text-red-100' : 'bg-blue-500 text-blue-100'} flex items-center justify-center mb-4`}><Tag size={20} /></div>
                                <div className="w-full">
                                  <span className="text-2xl font-black block leading-none text-white">{cnt}</span>
                                  <h4 className={`text-xs font-black ${isRed ? 'text-red-50' : 'text-blue-50'} tracking-tight mt-1.5 leading-tight`}>{opt.label}</h4>
                                </div>
                              </motion.button>
                            );
                          })}
                        </div>
                      ) : (
                        /* Visa type cards — same design as section cards */
                        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 xl:grid-cols-6 gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                          {[
                            'عمل','عمالة منزلية','مرافق عمال','زيارة','زيارة عائلية','عائلة مقيم',
                            'دراسة','مرافق طالب','حج','عمرة','موسمية','تجارة',
                            'استثمار','عبور','مؤتمر','رياضية','علاج طبي','صحفية',
                            'دينية','دبلوماسية','مهمة رسمية','خاصة','فنية وترفيهية','سياحية',
                            'تطوع','مؤقتة','أخرى'
                          ].map((vt) => {
                            const cnt = entries.filter(e => !e.archived && e.visaType === vt).length;
                            return (
                              <motion.button
                                key={vt}
                                whileHover={{ scale: 1.03 }}
                                whileTap={{ scale: 0.97 }}
                                onClick={() => setSelectedVisaType(vt)}
                                dir={lang === 'ar' ? 'rtl' : 'ltr'}
                                className="bg-white border border-slate-100 hover:border-slate-200 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer shadow-sm group h-full"
                              >
                                <div className="w-10 h-10 rounded-xl bg-slate-50 text-slate-600 group-hover:bg-slate-100 flex items-center justify-center mb-4 transition-all group-hover:rotate-6">
                                  <Briefcase size={18} />
                                </div>
                                <div className="w-full">
                                  <span className="text-2xl font-black block leading-none text-slate-900">{cnt}</span>
                                  <h4 className="text-xs font-black text-slate-700 tracking-tight mt-1.5 leading-tight">{vt}</h4>
                                </div>
                              </motion.button>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  ) : (
                  /* Show ONLY the departmental cards */
                  <div className="space-y-6">
                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 xl:grid-cols-6 gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                      {/* Card 1: الإحصائيات */}
                      <motion.button
                        whileHover={{ scale: 1.02 }}
                        whileTap={{ scale: 0.98 }}
                        onClick={() => setIsStatsView(true)}
                        dir={lang === 'ar' ? 'rtl' : 'ltr'}
                        className="bg-white border border-slate-100 hover:border-slate-200 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer group h-full shadow-sm"
                      >
                        <div className="w-10 h-10 rounded-xl bg-slate-50 text-slate-600 group-hover:bg-slate-100 flex items-center justify-center mb-4 transition-all group-hover:rotate-6">
                          <BarChart3 size={20} />
                        </div>
                        <div className="w-full">
                          <span className="text-2xl font-black block leading-none text-slate-900">
                            {entries.filter(e => !e.archived).length}
                          </span>
                          <h4 className="text-xs font-black text-slate-800 tracking-tight leading-tight mt-1.5">
                            {lang === 'ar' ? 'إحصائية عامة لكافة المعاملات' : 'General Statistics'}
                          </h4>
                          <p className="text-[10px] font-bold text-slate-400 mt-0.5">
                            {lang === 'ar' ? 'عرض كافة المعاملات والحالات' : 'All transactions overview'}
                          </p>
                        </div>
                      </motion.button>

                      {/* Card: المعاملات → visa types page */}
                      <motion.button
                        whileHover={{ scale: 1.02 }}
                        whileTap={{ scale: 0.98 }}
                        dir={lang === 'ar' ? 'rtl' : 'ltr'}
                        onClick={() => setIsVisaTypeView(true)}
                        className="bg-white border border-blue-100 hover:border-blue-300 hover:bg-blue-50/30 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer group h-full shadow-sm"
                      >
                        <div className="w-10 h-10 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center mb-4 group-hover:bg-blue-100 transition-all group-hover:rotate-6">
                          <Tag size={20} />
                        </div>
                        <div className="w-full">
                          <span className="text-2xl font-black block leading-none text-blue-600">
                            {[...new Set(entries.filter(e => !e.archived).map(e => e.visaType))].length}
                          </span>
                          <h4 className="text-xs font-black text-slate-800 tracking-tight leading-tight mt-1.5">
                            {lang === 'ar' ? 'أقسام المعاملات حسب نوع التأشيرة' : 'Transactions by Visa Type'}
                          </h4>
                          <p className="text-[10px] font-bold text-slate-400 mt-0.5">
                            {lang === 'ar' ? 'تصفح حسب نوع التأشيرة' : 'Browse by visa type'}
                          </p>
                        </div>
                      </motion.button>

                      {/* Card: معاملات المكاتب والوكلاء */}
                      <motion.button
                        whileHover={{ scale: 1.02 }}
                        whileTap={{ scale: 0.98 }}
                        dir={lang === 'ar' ? 'rtl' : 'ltr'}
                        onClick={() => { setIsOfficesView(true); setSelectedOfficeFilter(null); }}
                        className="bg-white border border-purple-100 hover:border-purple-300 hover:bg-purple-50/30 rounded-3xl p-5 text-start flex flex-col justify-between items-start transition-all cursor-pointer group h-full shadow-sm"
                      >
                        <div className="w-10 h-10 rounded-xl bg-purple-50 text-purple-600 flex items-center justify-center mb-4 group-hover:bg-purple-100 transition-all group-hover:rotate-6">
                          <Users size={20} />
                        </div>
                        <div className="w-full">
                          <span className="text-2xl font-black block leading-none text-purple-600">
                            {entries.filter(e => !e.archived).length}
                          </span>
                          <h4 className="text-xs font-black text-slate-800 tracking-tight leading-tight mt-1.5">
                            {lang === 'ar' ? 'معاملات المكاتب والوكلاء' : 'Office & Agent Transactions'}
                          </h4>
                          <p className="text-[10px] font-bold text-slate-400 mt-0.5">
                            {lang === 'ar' ? 'تصفح حسب المكتب أو الوكيل' : 'Browse by office or agent'}
                          </p>
                        </div>
                      </motion.button>

                    </div>
                  </div>
                  )
                ) : (
                  /* Inside a specific Section Workspace Page */
                  <div className="space-y-6">
                    {/* Header showing Active Section with Back Button */}
                    <div className="bg-white p-6 sm:p-8 rounded-[2.5rem] border border-slate-200 shadow-sm flex items-center justify-between gap-4 mb-2" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                      <div className="flex items-center gap-4">
                        <button 
                          onClick={() => setIsSectionSelected(false)}
                          className="w-10 h-10 flex items-center justify-center text-slate-600 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 rounded-xl shadow-sm transition-all focus:outline-none focus:ring-4 focus:ring-slate-900/5 cursor-pointer"
                        >
                          {lang === 'ar' ? <ArrowRight size={20} /> : <ArrowLeft size={20} />}
                        </button>
                        <div className="text-right flex flex-col items-start">
                          <div className="flex flex-wrap items-center gap-1.5 justify-start">
                            <span className="text-[10px] font-bold text-[#64748B] uppercase tracking-wider">
                              {lang === 'ar' ? 'أقسام المعاملات' : 'Transactions Departments'}
                            </span>
                            <span className="text-slate-300">/</span>
                            <span className={`text-[10px] font-bold uppercase tracking-wider ${
                              allTabFilter === 'all' ? 'text-slate-600' :
                              allTabFilter === 'today' ? 'text-indigo-600' :
                              allTabFilter === 'with_print' ? 'text-blue-600' :
                              allTabFilter === 'without_print' ? 'text-amber-600' :
                              allTabFilter.startsWith('tx_status:') ? 'text-purple-600' : 'text-emerald-500'
                            }`}>
                              {allTabFilter === 'all' ? (lang === 'ar' ? 'كافة المعاملات' : 'All Transactions') :
                               allTabFilter === 'today' ? (lang === 'ar' ? 'تحديثات اليوم' : "Today's Updates") :
                               allTabFilter === 'with_print' ? (lang === 'ar' ? 'المعاملات لها طلب ( برنت )' : 'Transactions with Print') :
                               allTabFilter === 'without_print' ? (lang === 'ar' ? 'المعاملات التي لم يصدر لها برنت' : 'Transactions without Print') :
                               allTabFilter.startsWith('tx_status:') ? allTabFilter.slice('tx_status:'.length) :
                               (lang === 'ar' ? 'المعاملات الجاهزة' : 'Ready/Issued Transactions')}
                            </span>
                          </div>
                          <h3 className="text-xs sm:text-sm font-black text-slate-900 tracking-tight text-right mt-1">
                            {allTabFilter === 'all' ? (lang === 'ar' ? 'كافة معاملات وعملاء المكتب الموحد' : 'All Enrolled Branch Transactions') :
                             allTabFilter === 'today' ? (lang === 'ar' ? 'سجل المعاملات المحدثة والمعدلة اليوم' : 'Daily Tracking & Synced Records') :
                             allTabFilter === 'with_print' ? (lang === 'ar' ? 'المعاملات التي تم سحب أو إدخال البرنت لها' : 'Transactions Configured with Print Details') :
                             allTabFilter === 'without_print' ? (lang === 'ar' ? 'المعاملات التي لم يتم سحب أو تزويد برنت لها بعد' : 'Transactions awaiting print configuration or diagnostic details') :
                             allTabFilter.startsWith('tx_status:') ? (lang === 'ar' ? `المعاملات ذات حالة: ${allTabFilter.slice('tx_status:'.length)}` : `Transactions with status: ${allTabFilter.slice('tx_status:'.length)}`) :
                             (lang === 'ar' ? 'كافة المعاملات الجاهزة والمؤشرة الصادرة' : 'Ready and Fully Issued Visas')}
                          </h3>
                        </div>
                      </div>

                      <button 
                        onClick={() => setIsSectionSelected(false)}
                        className="hidden sm:flex items-center gap-1.5 text-xs font-black text-slate-500 hover:text-slate-900 bg-slate-50 hover:bg-slate-100 border border-slate-200 px-4 py-2.5 rounded-xl cursor-pointer transition-all"
                      >
                        {lang === 'ar' ? (
                          <>
                            <span>الرجوع للأقسام</span>
                            <ArrowLeft size={14} />
                          </>
                        ) : (
                          <>
                            <ArrowRight size={14} />
                            <span>Back to Departments</span>
                          </>
                        )}
                      </button>
                    </div>



                {/* Admin controls and grid listing */}
                {!isArchiveView && <section className="bg-white p-4 rounded-3xl border border-slate-200 shadow-sm flex flex-wrap items-center justify-between gap-4">
                  <div className="relative flex-1 min-w-[300px]">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" size={16} />
                    <input 
                      type="text" 
                      placeholder={t.searchPlaceholder}
                      className="w-full bg-slate-50 border border-slate-200 rounded-2xl py-2.5 pl-10 pr-4 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 transition-all font-medium"
                      value={filterQuery}
                      onChange={(e) => setFilterQuery(e.target.value)}
                    />
                  </div>

                  <div className="flex items-center gap-2">
                    {isSelectMode ? (
                      <>
                        <button
                          onClick={() => {
                            if (selectedEntryIds.size === filteredByTab.length) {
                              setSelectedEntryIds(new Set());
                            } else {
                              setSelectedEntryIds(new Set(filteredByTab.map(e => e.id)));
                            }
                          }}
                          className="py-2 px-4 rounded-2xl text-xs font-bold bg-violet-50 text-violet-700 border border-violet-200 hover:bg-violet-100 transition-all cursor-pointer"
                        >
                          {selectedEntryIds.size === filteredByTab.length
                            ? (lang === 'ar' ? 'إلغاء تحديد الكل' : 'Deselect All')
                            : (lang === 'ar' ? 'تحديد الكل' : 'Select All')}
                        </button>
                        <button
                          onClick={exitSelectMode}
                          className="py-2 px-4 rounded-2xl text-xs font-bold bg-slate-100 text-slate-600 border border-slate-200 hover:bg-slate-200 transition-all cursor-pointer"
                        >
                          {lang === 'ar' ? 'إلغاء' : 'Cancel'}
                        </button>
                      </>
                    ) : (
                      <>
                        {/* Filter Button */}
                        <div className="relative" ref={filterMenuRef}>
                          <button
                            onClick={() => setShowFilterMenu(p => !p)}
                            className={`py-2 px-4 rounded-2xl text-xs font-bold border transition-all cursor-pointer flex items-center gap-1.5 ${
                              sortOrder || docNumberFilter
                                ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
                                : 'bg-slate-50 text-slate-600 border-slate-200 hover:bg-slate-100'
                            }`}
                          >
                            <SlidersHorizontal size={13} />
                            {lang === 'ar' ? 'فلترة' : 'Filter'}
                            {(sortOrder || docNumberFilter) && (
                              <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 shrink-0" />
                            )}
                          </button>

                          <AnimatePresence>
                            {showFilterMenu && (
                              <>
                                <div className="fixed inset-0 z-[100]" onClick={() => setShowFilterMenu(false)} />
                                <motion.div
                                  initial={{ opacity: 0, y: 6, scale: 0.97 }}
                                  animate={{ opacity: 1, y: 0, scale: 1 }}
                                  exit={{ opacity: 0, y: 6, scale: 0.97, transition: { duration: 0 } }}
                                  onClick={e => e.stopPropagation()}
                                  className="absolute top-full mt-2 right-0 z-[101] bg-white rounded-2xl border border-slate-200 shadow-2xl p-3 min-w-[200px] space-y-1"
                                  dir={lang === 'ar' ? 'rtl' : 'ltr'}
                                >
                                  {/* Sort section */}
                                  <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2 pb-1 pt-1">{lang === 'ar' ? 'الترتيب' : 'Sort'}</p>
                                  {([
                                    { val: 'newest', label: lang === 'ar' ? 'الأحدث أولاً' : 'Newest First' },
                                    { val: 'oldest', label: lang === 'ar' ? 'الأقدم أولاً' : 'Oldest First' },
                                  ] as const).map(opt => (
                                    <button
                                      key={opt.val}
                                      onClick={() => { setSortOrder(sortOrder === opt.val ? null : opt.val); setShowFilterMenu(false); }}
                                      className={`w-full flex items-center gap-2 px-3 py-2 rounded-xl text-sm font-bold transition-all cursor-pointer ${sortOrder === opt.val ? 'bg-emerald-50 text-emerald-700' : 'hover:bg-slate-50 text-slate-600'}`}
                                    >
                                      <ArrowDownUp size={13} />
                                      {opt.label}
                                      {sortOrder === opt.val && <Check size={12} className="mr-auto text-emerald-600" />}
                                    </button>
                                  ))}

                                  {/* Doc number filter — only for with_print */}
                                  {allTabFilter === 'with_print' && (
                                    <>
                                      <div className="border-t border-slate-100 my-2" />
                                      <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest px-2 pb-1">{lang === 'ar' ? 'رقم المستند' : 'Document No.'}</p>
                                      {([
                                        { val: 'with_doc', label: lang === 'ar' ? 'لها رقم مستند' : 'Has Doc. No.' },
                                        { val: 'without_doc', label: lang === 'ar' ? 'ليس لها رقم مستند' : 'No Doc. No.' },
                                      ] as const).map(opt => (
                                        <button
                                          key={opt.val}
                                          onClick={() => { setDocNumberFilter(docNumberFilter === opt.val ? null : opt.val); setShowFilterMenu(false); }}
                                          className={`w-full flex items-center gap-2 px-3 py-2 rounded-xl text-sm font-bold transition-all cursor-pointer ${docNumberFilter === opt.val ? 'bg-violet-50 text-violet-700' : 'hover:bg-slate-50 text-slate-600'}`}
                                        >
                                          <Hash size={13} />
                                          {opt.label}
                                          {docNumberFilter === opt.val && <Check size={12} className="mr-auto text-violet-600" />}
                                        </button>
                                      ))}
                                    </>
                                  )}

                                  {/* Reset */}
                                  {(sortOrder || docNumberFilter) && (
                                    <>
                                      <div className="border-t border-slate-100 my-2" />
                                      <button
                                        onClick={() => { setSortOrder(null); setDocNumberFilter(null); setShowFilterMenu(false); }}
                                        className="w-full flex items-center gap-2 px-3 py-2 rounded-xl text-sm font-bold text-red-500 hover:bg-red-50 transition-all cursor-pointer"
                                      >
                                        <X size={13} />
                                        {lang === 'ar' ? 'إزالة الفلاتر' : 'Clear Filters'}
                                      </button>
                                    </>
                                  )}
                                </motion.div>
                              </>
                            )}
                          </AnimatePresence>
                        </div>

                        {/* Select Button */}
                        <button
                          onClick={() => setIsSelectMode(true)}
                          className="py-2 px-4 rounded-2xl text-xs font-bold bg-slate-50 text-slate-600 border border-slate-200 hover:bg-slate-100 transition-all cursor-pointer flex items-center gap-1.5"
                        >
                          <CheckSquare size={13} />
                          {lang === 'ar' ? 'تحديد' : 'Select'}
                        </button>
                      </>
                    )}
                  </div>

                  <input 
                    type="file" 
                    ref={fileInputRef} 
                    className="hidden" 
                    accept=".xlsx,.xls,.csv" 
                    onChange={handleFileUpload}
                  />
                </section>}

                {/* Main Interactive Registry Cards Grid */}
                {!isArchiveView && <div>
                  {filteredByTab.length === 0 ? (
                    <div className="bg-white rounded-3xl border border-slate-200 shadow-sm p-32 flex flex-col items-center justify-center text-center">
                      <div className="w-24 h-24 bg-slate-50 rounded-full flex items-center justify-center text-slate-300 mb-6 font-bold">
                        <Search size={48} />
                      </div>
                      <h4 className="text-slate-400 font-bold uppercase tracking-widest text-xs mb-2">
                        {lang === 'ar' ? 'لا يوجد نتائج مطابقة' : 'No Records Matches'}
                      </h4>
                      <p className="text-slate-300 text-sm max-w-xs">{t.noData}</p>
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                      <AnimatePresence initial={false}>
                        {filteredByTab.map((entry, idx) => {
                          const isExpanded = expandedCardId === entry.id;
                          const _offObj = entry.officeId && entry.officeId !== 'general' ? offices.find(o => o.id === entry.officeId) : null;
                          const _offName = _offObj ? _offObj.name : (lang === 'ar' ? 'عام' : 'General');
                          const _lastCheck = entry.updatedAtUnix
                            ? (() => { const d = new Date(entry.updatedAtUnix!); const p = (n: number) => n.toString().padStart(2, '0'); return `${p(d.getDate())}/${p(d.getMonth()+1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`; })()
                            : (entry.lastUpdate || null);
                          const _statusBg =
                            entry.status === 'Found'    ? 'bg-emerald-500' :
                            entry.status === 'Error'    ? 'bg-red-400' :
                            entry.status === 'Checking' ? 'bg-amber-400' :
                            entry.status === 'Retrying' ? 'bg-orange-400' :
                            'bg-slate-300';
                          const _statusLabel =
                            entry.status === 'Found'    ? (lang === 'ar' ? 'مكتملة' : 'Found') :
                            entry.status === 'Error'    ? (lang === 'ar' ? 'خطأ' : 'Error') :
                            entry.status === 'Checking' ? (lang === 'ar' ? 'جاري...' : 'Checking') :
                            entry.status === 'Retrying' ? (lang === 'ar' ? 'إعادة' : 'Retry') :
                            (lang === 'ar' ? 'انتظار' : 'Idle');
                          const _statusRing =
                            entry.status === 'Found'    ? 'bg-emerald-50 text-emerald-700 border-emerald-200' :
                            entry.status === 'Error'    ? 'bg-red-50 text-red-700 border-red-200' :
                            entry.status === 'Checking' ? 'bg-amber-50 text-amber-700 border-amber-200' :
                            entry.status === 'Retrying' ? 'bg-orange-50 text-orange-700 border-orange-200' :
                            'bg-slate-50 text-slate-500 border-slate-200';
                          const isSelected = selectedEntryIds.has(entry.id);
                          return (
                          <motion.div 
                            key={entry.id}
                            initial={{ opacity: 0, scale: 0.97 }}
                            animate={{ opacity: 1, scale: 1 }}
                            exit={{ opacity: 0, scale: 0.97 }}
                            className={`bg-white rounded-3xl border shadow-sm overflow-hidden transition-all duration-200 relative ${
                              isSelected
                                ? 'border-violet-400 ring-2 ring-violet-200 shadow-violet-100'
                                : isExpanded
                                  ? 'border-slate-300 shadow-md col-span-1 sm:col-span-2 xl:col-span-3'
                                  : entry.dataUpdatedAtUnix && (Date.now() - entry.dataUpdatedAtUnix) < 24 * 60 * 60 * 1000
                                    ? 'border-emerald-300 shadow-emerald-100 hover:shadow-md ring-1 ring-emerald-200'
                                    : 'border-slate-200 hover:border-slate-300 hover:shadow-md'
                            }`}
                          >
                            {/* ── Transaction Status Strip ── */}
                            {(() => {
                              const txType = entry.txStatusType || 'مستمرة';
                              const txLabel = entry.txStatus || 'مستمرة';
                              const txDate = entry.txStatusDate;
                              return (
                                <button
                                  onClick={(e) => { e.stopPropagation(); openTxStatusModal(entry); }}
                                  className={`w-full flex items-center gap-2 px-4 py-2 border-b text-right transition-all hover:opacity-80 cursor-pointer ${txTypeStripBg(txType)}`}
                                  dir="rtl"
                                >
                                  <span className={`w-2 h-2 rounded-full shrink-0 ${txTypeDot(txType)}`} />
                                  <span className="text-[11px] font-black text-slate-700 truncate flex-1">{txLabel}</span>
                                  {txDate && (
                                    <span className="text-[9px] font-bold text-slate-400 font-mono shrink-0">{txDate}</span>
                                  )}
                                  <Pencil size={10} className="text-slate-300 shrink-0" />
                                </button>
                              );
                            })()}
                            {/* ── Modern Card Header ── */}
                            <div
                              className="p-5 cursor-pointer select-none"
                              onClick={() => {
                                if (isSelectMode) { toggleSelectEntry(entry.id); return; }
                                setExpandedCardId(isExpanded ? null : entry.id);
                              }}
                              onPointerDown={() => {
                                if (isSelectMode) return;
                                longPressTimer.current = setTimeout(() => {
                                  setIsSelectMode(true);
                                  setSelectedEntryIds(new Set([entry.id]));
                                }, 600);
                              }}
                              onPointerUp={() => { if (longPressTimer.current) { clearTimeout(longPressTimer.current); longPressTimer.current = null; } }}
                              onPointerLeave={() => { if (longPressTimer.current) { clearTimeout(longPressTimer.current); longPressTimer.current = null; } }}
                            >
                              {/* New data badge */}
                              {entry.dataUpdatedAtUnix && (Date.now() - entry.dataUpdatedAtUnix) < 24 * 60 * 60 * 1000 && (
                                <div className="flex justify-end mb-2">
                                  <span className="inline-flex items-center gap-1 text-[10px] font-black bg-emerald-500 text-white rounded-full px-2.5 py-0.5 shadow-sm shadow-emerald-500/30 animate-pulse">
                                    <span className="w-1.5 h-1.5 rounded-full bg-white shrink-0" />
                                    {lang === 'ar' ? 'بيانات جديدة' : 'New Data'}
                                  </span>
                                </div>
                              )}
                              {/* Top: name + status badge */}
                              <div className="flex items-start justify-between gap-3 mb-4">
                                <div className="min-w-0 text-right flex-1">
                                  <p className="text-sm font-black text-slate-900 truncate leading-tight">{entry.applicantName || entry.nameEnglish || (lang === 'ar' ? 'غير محدد' : '—')}</p>
                                  <p className="text-[10px] font-mono font-bold text-slate-400 mt-0.5 truncate">{entry.passportNumber || '—'}</p>
                                </div>
                                <div className="flex flex-col items-end gap-1.5 shrink-0">
                                  {isSelectMode ? (
                                    <div
                                      className={`w-7 h-7 rounded-full border-2 flex items-center justify-center transition-all cursor-pointer shadow-sm ${isSelected ? 'bg-violet-600 border-violet-600' : 'bg-white border-slate-300'}`}
                                      onClick={(e) => { e.stopPropagation(); toggleSelectEntry(entry.id); }}
                                    >
                                      {isSelected && <Check size={13} className="text-white" strokeWidth={3} />}
                                    </div>
                                  ) : (
                                    <span className={`inline-flex items-center gap-1.5 text-[10px] font-black px-2.5 py-1 rounded-full border ${_statusRing}`}>
                                      <span className={`w-1.5 h-1.5 rounded-full shrink-0 ${_statusBg} ${entry.status === 'Checking' || entry.status === 'Retrying' ? 'animate-pulse' : ''}`} />
                                      {_statusLabel}
                                    </span>
                                  )}
                                  <span className="text-[9px] font-black text-slate-300 font-mono">#{idx + 1}</span>
                                </div>
                              </div>

                              {/* Middle: visa type + office + contact badges */}
                              <div className="flex flex-wrap gap-1.5 justify-end mb-3">
                                {entry.visaType && (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-100 rounded-full px-2.5 py-0.5">
                                    <Tag size={9} />
                                    {entry.visaType}
                                  </span>
                                )}
                                <span className="inline-flex items-center gap-1 text-[10px] font-bold bg-violet-50 text-violet-700 border border-violet-100 rounded-full px-2.5 py-0.5">
                                  <Briefcase size={9} />
                                  {_offName}
                                </span>
                                {entry.applicantData && entry.applicantData.includes('تاريخ الاستلام:') && (
                                  <span className="inline-flex items-center gap-1 text-[10px] font-bold bg-sky-50 text-sky-700 border border-sky-100 rounded-full px-2.5 py-0.5">
                                    <CalendarDays size={9} />
                                    {entry.applicantData.replace('تاريخ الاستلام:', '').trim()}
                                  </span>
                                )}
                                {/* Health cert status badge */}
                                {entry.healthCertStatus && (() => {
                                  const hcs = entry.healthCertStatus;
                                  const hcIssued = hcs.includes('تم اصدار') || hcs.includes('تم إصدار') || hcs.includes('الشهادة صادرة');
                                  const hcNotIssued = hcs.includes('لم يتم') || hcs.includes('لم تصدر') || hcs.includes('غير صادرة');
                                  return (
                                    <span className={`inline-flex items-center gap-1 text-[10px] font-bold border rounded-full px-2.5 py-0.5 ${
                                      hcIssued
                                        ? 'bg-teal-50 text-teal-700 border-teal-200'
                                        : hcNotIssued
                                          ? 'bg-orange-50 text-orange-700 border-orange-200'
                                          : 'bg-slate-50 text-slate-500 border-slate-200'
                                    }`}>
                                      <span className="text-[11px]">🏥</span>
                                      {hcIssued
                                        ? (lang === 'ar' ? 'الكرت صادر ✓' : 'Health Cert Issued ✓')
                                        : hcNotIssued
                                          ? (lang === 'ar' ? 'لم يُصدر بعد' : 'Not Issued Yet')
                                          : hcs.slice(0, 30)}
                                    </span>
                                  );
                                })()}
                              </div>

                              {/* Bottom: last update + actions */}
                              <div className="flex items-center justify-between gap-2 pt-3 border-t border-slate-100">
                                <div className="flex items-center gap-1" onClick={(e) => e.stopPropagation()}>
                                  {checkActionOpenId === entry.id ? (
                                    /* Expanded: show two action buttons */
                                    <div className="flex items-center gap-1">
                                      <button
                                        onClick={() => { if (entry.txStatusType === 'موقفة') return; setCheckActionOpenId(null); checkVisa(entry.id); }}
                                        disabled={entry.txStatusType === 'موقفة'}
                                        className={`flex items-center gap-1 px-2 py-1 rounded-xl border transition-all text-[10px] font-bold ${entry.txStatusType === 'موقفة' ? 'bg-slate-50 text-slate-300 border-slate-200 cursor-not-allowed' : 'bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100 cursor-pointer'}`}
                                        title={entry.txStatusType === 'موقفة' ? 'المعاملة موقفة — لا يمكن التشييك' : (lang === 'ar' ? 'تشييك' : 'Check')}
                                      >
                                        <RefreshCw size={11} className={entry.status === 'Checking' ? 'animate-spin' : ''} />
                                        {lang === 'ar' ? 'تشييك' : 'Check'}
                                      </button>
                                      <button
                                        onClick={() => { checkHealthCert(entry); }}
                                        disabled={checkingHealthCertIds.has(entry.id)}
                                        className={`flex items-center gap-1 px-2 py-1 rounded-xl border transition-all text-[10px] font-bold cursor-pointer disabled:opacity-60 disabled:cursor-not-allowed ${
                                          hasValidAppNo(entry)
                                            ? 'bg-cyan-50 text-cyan-700 border-cyan-200 hover:bg-cyan-100'
                                            : 'bg-slate-50 text-slate-400 border-slate-200 hover:bg-slate-100'
                                        }`}
                                        title={hasValidAppNo(entry) ? (lang === 'ar' ? 'تحديث حالة الكرت الصحي' : 'Update Health Cert') : (lang === 'ar' ? 'لا يوجد برنت لهذه المعاملة' : 'No print for this entry')}
                                      >
                                        {checkingHealthCertIds.has(entry.id)
                                          ? <><RefreshCw size={11} className="animate-spin" />{lang === 'ar' ? 'جاري...' : 'Checking...'}</>
                                          : <><span className="text-[11px]">{hasValidAppNo(entry) ? '🏥' : '🚫'}</span>{lang === 'ar' ? 'الكرت الصحي' : 'Health Cert'}</>}
                                      </button>
                                      {!checkingHealthCertIds.has(entry.id) && (
                                        <button
                                          onClick={() => setCheckActionOpenId(null)}
                                          className="p-1 rounded-xl hover:bg-slate-100 transition-all text-slate-300 cursor-pointer"
                                          title={lang === 'ar' ? 'إغلاق' : 'Close'}
                                        >
                                          <X size={11} />
                                        </button>
                                      )}
                                    </div>
                                  ) : (
                                    /* Collapsed: single sync button that opens the two choices */
                                    <button
                                      onClick={() => { if (entry.txStatusType !== 'موقفة') setCheckActionOpenId(entry.id); }}
                                      disabled={entry.txStatusType === 'موقفة'}
                                      className={`p-1.5 rounded-xl transition-all border ${entry.txStatusType === 'موقفة' ? 'text-red-200 border-transparent cursor-not-allowed' : 'hover:bg-emerald-50 hover:text-emerald-600 cursor-pointer text-slate-400 border-transparent hover:border-emerald-100'}`}
                                      title={entry.txStatusType === 'موقفة' ? 'المعاملة موقفة — لا يمكن التشييك' : (lang === 'ar' ? 'تشييك' : 'Check')}
                                    >
                                      <RefreshCw size={13} className={entry.status === 'Checking' ? 'animate-spin' : ''} />
                                    </button>
                                  )}
                                  {hasPermission('edit_clients') && (
                                    <button
                                      onClick={() => setEditingCardId(editingCardId === entry.id ? null : entry.id)}
                                      className={`p-1.5 rounded-xl transition-all cursor-pointer border ${editingCardId === entry.id ? 'bg-blue-50 text-blue-500 border-blue-200' : 'hover:bg-blue-50 hover:text-blue-500 text-slate-400 border-transparent hover:border-blue-100'}`}
                                      title={lang === 'ar' ? (editingCardId === entry.id ? 'إيقاف التعديل' : 'تعديل البيانات') : (editingCardId === entry.id ? 'Stop editing' : 'Edit')}
                                    >
                                      <Pencil size={13} />
                                    </button>
                                  )}
                                  <div className="w-3" />
                                  <button onClick={() => setArchiveConfirmEntry(entry)} className="p-1.5 rounded-xl hover:bg-red-50 hover:text-red-500 transition-all cursor-pointer text-slate-400 border border-transparent hover:border-red-100" title={lang === 'ar' ? 'حذف / أرشفة' : 'Delete / Archive'}>
                                    <Trash2 size={13} />
                                  </button>
                                </div>
                                <div className="flex items-center gap-2">
                                  {_lastCheck && (
                                    <p className="text-[9px] font-bold text-slate-300 font-mono">{_lastCheck}</p>
                                  )}
                                  <motion.span animate={{ rotate: isExpanded ? 180 : 0 }} transition={{ duration: 0.2 }} className="text-slate-300">
                                    <ChevronDown size={14} />
                                  </motion.span>
                                </div>
                              </div>
                            </div>

                            {/* ── Expanded Full Details ── */}
                            <AnimatePresence initial={false}>
                              {isExpanded && (
                                <motion.div
                                  key="expanded"
                                  initial={{ height: 0, opacity: 0 }}
                                  animate={{ height: 'auto', opacity: 1 }}
                                  exit={{ height: 0, opacity: 0 }}
                                  transition={{ duration: 0.25 }}
                                  style={{ overflow: 'hidden' }}
                                >
                                  <div className="border-t border-slate-100 p-5">
                                    <div>

                              {/* Transaction Status Note */}
                              {entry.txStatusNote && (
                                <div className={`flex items-start gap-2.5 rounded-2xl border px-3.5 py-3 mb-4 text-right ${
                                  entry.txStatusType === 'موقفة' ? 'bg-red-50 border-red-100' :
                                  entry.txStatusType === 'بدون'  ? 'bg-slate-50 border-slate-100' :
                                  'bg-emerald-50 border-emerald-100'
                                }`} dir="rtl">
                                  <span className={`w-2 h-2 rounded-full shrink-0 mt-1 ${
                                    entry.txStatusType === 'موقفة' ? 'bg-red-400' :
                                    entry.txStatusType === 'بدون'  ? 'bg-slate-300' :
                                    'bg-emerald-400'
                                  }`} />
                                  <div className="flex-1 min-w-0">
                                    <p className={`text-[9px] font-black uppercase tracking-widest mb-0.5 ${
                                      entry.txStatusType === 'موقفة' ? 'text-red-400' :
                                      entry.txStatusType === 'بدون'  ? 'text-slate-400' :
                                      'text-emerald-600'
                                    }`}>ملاحظة الحالة · {entry.txStatus || 'مستمرة'}</p>
                                    <p className="text-xs font-semibold text-slate-700 leading-relaxed whitespace-pre-wrap break-words">{entry.txStatusNote}</p>
                                  </div>
                                </div>
                              )}

                              {/* Card Content & Fields */}
                              <div className="space-y-4">
                                {/* Passport only */}
                                <div className="space-y-1 text-right overflow-hidden">
                                  <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">
                                    {lang === 'ar' ? 'رقم جواز السفر' : 'Passport'}
                                  </label>
                                  <input 
                                    className="w-full bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-mono font-black text-slate-900 uppercase text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all disabled:opacity-60 disabled:cursor-not-allowed"
                                    value={entry.passportNumber}
                                    disabled={!hasPermission('edit_clients') || editingCardId !== entry.id}
                                    onChange={(e) => updateEntryField(entry.id, 'passportNumber', e.target.value)}
                                  />
                                </div>


                                {/* AI Captcha Solving Live Console */}
                                {entry.status === 'Checking' && (
                                  <div className="bg-slate-900 border border-slate-800 rounded-2xl p-3 text-right text-xs text-emerald-400 font-mono space-y-1.5 mt-2 animate-pulse" dir="rtl">
                                    <div className="flex items-center justify-between border-b border-slate-800 pb-1.5 mb-1.5 text-[9px] text-slate-400">
                                      <span className="flex items-center gap-1">
                                        <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-ping inline-block" />
                                        {lang === 'ar' ? '🧠 معالج الكابتشا الذكي (Gemini)' : 'AI Captcha Engine'}
                                      </span>
                                      <span className="text-[8px] bg-slate-800 px-1.5 py-0.5 rounded text-emerald-300">ACTIVE</span>
                                    </div>
                                    <p className="text-[9px] text-emerald-300">⚡ {lang === 'ar' ? 'جاري الاتصال بوزارة الخارجية...' : 'Connecting to MOFA port...'}</p>
                                    <p className="text-[9px] text-slate-400">📸 {lang === 'ar' ? 'جاري سحب وتفكيك الـ Captcha...' : 'Extracting Captcha vector...'}</p>
                                    <p className="text-[9px] text-slate-400">🧠 {lang === 'ar' ? 'جاري فك الترميز بواسطة Gemini 3.5...' : 'Decoding OCR with Gemini...'}</p>
                                  </div>
                                )}

                                {entry.captchaSvg && entry.status !== 'Checking' && (
                                  <div className="bg-slate-950 border border-slate-855 rounded-2xl p-3 text-right font-mono space-y-2 mt-2" dir="rtl">
                                    <div className="flex items-center justify-between border-b border-slate-800 pb-1 text-[9px] text-slate-400 font-bold">
                                      <span>🤖 {lang === 'ar' ? 'مستكشف الكابتشا المندمج (Gemini)' : 'Embedded CAPTCHA Solver'}</span>
                                      <span className="text-emerald-400 font-bold flex items-center gap-1">
                                        <span className="w-1 h-1 rounded-full bg-emerald-400 inline-block" />
                                        {lang === 'ar' ? 'مكتمل' : 'SUCCESS'}
                                      </span>
                                    </div>
                                    
                                    <div className="flex items-center justify-between bg-slate-900 p-2 rounded-xl border border-slate-800">
                                      <div className="space-y-0.5 text-right">
                                        <span className="text-[8px] text-slate-400 block">{lang === 'ar' ? 'رمز الكابتشا المسحوب' : 'SVG Captcha'}</span>
                                        <div 
                                          dangerouslySetInnerHTML={{ __html: entry.captchaSvg }} 
                                          className="h-8 border border-slate-800 rounded bg-slate-100 p-0.5 flex items-center justify-center overflow-hidden scale-90 origin-right"
                                        />
                                      </div>
                                      <div className="text-center bg-slate-950 px-3 py-1 rounded-lg border border-slate-850">
                                        <span className="text-[8px] text-slate-400 block">{lang === 'ar' ? 'الرمز المكتشف' : 'Detected text'}</span>
                                        <span className="text-sm font-black text-emerald-400 tracking-widest">{entry.solvedCaptcha}</span>
                                      </div>
                                    </div>

                                    {/* Small console logs */}
                                    <div className="text-[8px] leading-snug space-y-0.5 text-slate-350 pt-1.5 border-t border-slate-800">
                                      {entry.checkLogs && entry.checkLogs.map((log: string, lIdx: number) => (
                                        <p key={lIdx} className={log.includes("بنجاح") || log.includes("مكتشف") || log.includes("نجاح") ? "text-emerald-400" : "text-slate-400"}>
                                          • {log}
                                        </p>
                                      ))}
                                    </div>
                                  </div>
                                )}

                                {/* Phone + Custom ID */}
                                <div className="grid grid-cols-2 gap-3" dir="rtl">
                                  <div className="space-y-1 text-right overflow-hidden">
                                    <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">
                                      {lang === 'ar' ? 'رقم الهاتف 📞' : 'Phone 📞'}
                                    </label>
                                    <input 
                                      className="w-full bg-slate-50 hover:bg-slate-100/50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-900 text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all font-mono disabled:opacity-60 disabled:cursor-not-allowed"
                                      value={entry.phoneNumber || ''}
                                      placeholder={lang === 'ar' ? 'رقم الهاتف' : 'No phone'}
                                      disabled={!hasPermission('edit_clients') || editingCardId !== entry.id}
                                      onChange={(e) => updateEntryField(entry.id, 'phoneNumber', e.target.value)}
                                    />
                                  </div>
                                  <div className="space-y-1 text-right overflow-hidden">
                                    <label className="text-[9px] font-bold text-amber-500 uppercase tracking-wider block truncate">
                                      {lang === 'ar' ? 'الرقم المخصص #' : 'Custom ID #'}
                                    </label>
                                    <input 
                                      className="w-full bg-amber-50 hover:bg-amber-50/80 border border-amber-200 rounded-xl px-2.5 py-1.5 text-xs font-black text-amber-800 text-center focus:outline-none focus:ring-4 focus:ring-amber-400/20 transition-all font-mono disabled:opacity-60 disabled:cursor-not-allowed"
                                      value={entry.customId || ''}
                                      placeholder="—"
                                      disabled={!hasPermission('edit_clients') || editingCardId !== entry.id}
                                      onChange={(e) => {
                                        if (!hasPermission('edit_clients') || editingCardId !== entry.id) return;
                                        const val = e.target.value;
                                        const norm = val.trim().toLowerCase();
                                        const matched = offices.find(o => o.phone && o.phone.trim().toLowerCase() === norm);
                                        setEntries(prev => prev.map(en =>
                                          en.id === entry.id
                                            ? { ...en, customId: val, officeId: matched ? matched.id : undefined }
                                            : en
                                        ));
                                      }}
                                    />
                                  </div>
                                </div>

                                {/* Receive Date + Office Name */}
                                <div className="grid grid-cols-2 gap-3" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                                  <div className="space-y-1 text-right overflow-hidden">
                                    <label className="text-[9px] font-bold text-sky-500 uppercase tracking-wider block truncate">
                                      {lang === 'ar' ? 'تاريخ استلام المعاملة 📅' : 'Receive Date 📅'}
                                    </label>
                                    <div className="relative w-full">
                                      <input
                                        type="date"
                                        className="w-full appearance-none bg-sky-50 hover:bg-sky-50/80 border border-sky-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-sky-800 text-center focus:outline-none focus:ring-4 focus:ring-sky-400/20 transition-all [&::-webkit-calendar-picker-indicator]:opacity-60 [&::-webkit-calendar-picker-indicator]:cursor-pointer disabled:opacity-60 disabled:cursor-not-allowed"
                                        value={entry.applicantData && entry.applicantData.includes('تاريخ الاستلام:') ? entry.applicantData.replace('تاريخ الاستلام:', '').trim() : ''}
                                        disabled={!hasPermission('edit_clients') || editingCardId !== entry.id}
                                        onChange={(e) => updateEntryField(entry.id, 'applicantData', e.target.value ? `تاريخ الاستلام: ${e.target.value}` : '')}
                                      />
                                    </div>
                                  </div>
                                  {(() => {
                                    const officeObj = offices.find(o => o.id === entry.officeId);
                                    const offDisplayName = officeObj ? officeObj.name : (lang === 'ar' ? 'عام' : 'General');
                                    return (
                                      <div className="space-y-1 text-right overflow-hidden">
                                        <label className="text-[9px] font-bold text-violet-500 uppercase tracking-wider block truncate">
                                          {lang === 'ar' ? 'المكتب / الوكيل 🏢' : 'Office 🏢'}
                                        </label>
                                        {editingCardId === entry.id && hasPermission('edit_clients') ? (
                                          <select
                                            className="w-full bg-violet-50 border border-violet-200 rounded-xl px-2 py-1.5 text-xs font-black text-violet-700 text-center focus:outline-none focus:ring-4 focus:ring-violet-400/20 transition-all cursor-pointer appearance-none"
                                            value={entry.officeId || 'general'}
                                            onChange={(e) => {
                                              const selectedId = e.target.value;
                                              if (selectedId === 'general') {
                                                setEntries(prev => prev.map(en => en.id === entry.id ? { ...en, officeId: undefined, customId: '' } : en));
                                              } else {
                                                const selectedOffice = offices.find(o => o.id === selectedId);
                                                setEntries(prev => prev.map(en => en.id === entry.id ? { ...en, officeId: selectedId, customId: selectedOffice?.phone || en.customId } : en));
                                              }
                                            }}
                                          >
                                            <option value="general">{lang === 'ar' ? 'عام' : 'General'}</option>
                                            {offices.map(o => (
                                              <option key={o.id} value={o.id}>{o.name}</option>
                                            ))}
                                          </select>
                                        ) : (
                                          <div className="w-full bg-violet-50 border border-violet-200 rounded-xl px-2.5 py-1.5 flex items-center justify-center gap-1.5 opacity-60">
                                            <span className="text-xs font-black text-violet-700 truncate">{offDisplayName}</span>
                                          </div>
                                        )}
                                      </div>
                                    );
                                  })()}
                                </div>

                                {/* Nationality & Arrival Port */}
                                <div className="grid grid-cols-2 gap-3" dir={lang === 'ar' ? 'rtl' : 'ltr'}>
                                  <div className="space-y-1 text-right overflow-hidden">
                                    <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">
                                      {lang === 'ar' ? 'الجنسية' : 'Nationality'}
                                    </label>
                                    <div className="w-full bg-slate-100 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-700 text-center select-none">
                                      {entry.nationality || '—'}
                                    </div>
                                  </div>
                                  <div className="space-y-1 text-right overflow-hidden">
                                    <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">
                                      {lang === 'ar' ? 'جهة القدوم' : 'Arrival Port'}
                                    </label>
                                    {editingCardId === entry.id && hasPermission('edit_clients') ? (
                                      <select
                                        className="w-full bg-slate-50 border border-slate-200 rounded-xl px-2 py-1.5 text-xs font-bold text-slate-900 text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all cursor-pointer appearance-none"
                                        value={entry.arrivalPoint}
                                        onChange={(e) => updateEntryField(entry.id, 'arrivalPoint', e.target.value)}
                                      >
                                        <option value="عدن">عدن</option>
                                        <option value="صنعاء">صنعاء</option>
                                        <option value="جدة">جدة</option>
                                        <option value="الرياض">الرياض</option>
                                        <option value="الدمام">الدمام</option>
                                        <option value="أخرى">أخرى</option>
                                      </select>
                                    ) : (
                                      <div className="w-full bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-700 text-center opacity-60">
                                        {entry.arrivalPoint || '—'}
                                      </div>
                                    )}
                                  </div>
                                </div>

                                {/* Visa Type */}
                                <div className="space-y-1 text-right overflow-hidden">
                                  <label className="text-[9px] font-bold text-slate-400 uppercase tracking-wider block truncate">
                                    {lang === 'ar' ? 'نوع التأشيرة' : 'Visa Type'}
                                  </label>
                                  {editingCardId === entry.id && hasPermission('edit_clients') ? (
                                    <select
                                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-2 py-1.5 text-xs font-bold text-slate-900 text-center focus:outline-none focus:ring-4 focus:ring-slate-900/5 transition-all cursor-pointer appearance-none"
                                      value={entry.visaType}
                                      onChange={(e) => updateEntryField(entry.id, 'visaType', e.target.value)}
                                    >
                                      <option value="عمل">عمل</option>
                                      <option value="عمالة منزلية">عمالة منزلية</option>
                                      <option value="مرافق عمال">مرافق عمال</option>
                                      <option value="زيارة">زيارة</option>
                                      <option value="زيارة عائلية">زيارة عائلية</option>
                                      <option value="عائلة مقيم">عائلة مقيم</option>
                                      <option value="دراسة">دراسة</option>
                                      <option value="مرافق طالب">مرافق طالب</option>
                                      <option value="حج">حج</option>
                                      <option value="عمرة">عمرة</option>
                                      <option value="موسمية">موسمية</option>
                                      <option value="تجارة">تجارة</option>
                                      <option value="استثمار">استثمار</option>
                                      <option value="عبور">عبور</option>
                                      <option value="مؤتمر">مؤتمر</option>
                                      <option value="رياضية">رياضية</option>
                                      <option value="علاج طبي">علاج طبي</option>
                                      <option value="صحفية">صحفية</option>
                                      <option value="دينية">دينية</option>
                                      <option value="دبلوماسية">دبلوماسية</option>
                                    </select>
                                  ) : (
                                    <div className="w-full bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1.5 text-xs font-bold text-slate-700 text-center opacity-60">
                                      {entry.visaType || '—'}
                                    </div>
                                  )}
                                </div>

                                {/* ── MOFA Extended Data Fields ── */}
                                <div className="border-t border-slate-100 pt-3 space-y-2">
                                  <span className="text-[8px] font-black text-emerald-600 uppercase tracking-widest bg-emerald-50 border border-emerald-100 px-2 py-0.5 rounded-full">بيانات منصة التأشيرات</span>
                                  {/* رقم الطلب + رقم التأشيرة + رقم المستند — مربعات بارزة */}
                                  <div className="grid grid-cols-3 gap-2 mb-2">
                                    <div className="space-y-0.5 text-right overflow-hidden">
                                      <label className="text-[8px] font-black text-slate-500 uppercase tracking-wider block truncate">📋 رقم الطلب</label>
                                      <div className="w-full bg-slate-100 border border-slate-200 rounded-lg px-2 py-1.5 text-[11px] font-black text-slate-800 text-center truncate min-h-[28px] flex items-center justify-center font-mono">
                                        {entry.applicationNumber && entry.applicationNumber !== '---' ? entry.applicationNumber : <span className="text-slate-300">—</span>}
                                      </div>
                                    </div>
                                    <div className="space-y-0.5 text-right overflow-hidden">
                                      <label className="text-[8px] font-black text-emerald-600 uppercase tracking-wider block truncate">🎫 رقم التأشيرة</label>
                                      <div className="w-full bg-emerald-50 border border-emerald-200 rounded-lg px-2 py-1.5 text-[11px] font-black text-emerald-800 text-center truncate min-h-[28px] flex items-center justify-center font-mono">
                                        {entry.visaNumber && entry.visaNumber !== '---' ? entry.visaNumber : <span className="text-slate-300">—</span>}
                                      </div>
                                    </div>
                                    <div className="space-y-0.5 text-right overflow-hidden">
                                      <label className="text-[8px] font-black text-amber-600 uppercase tracking-wider block truncate">📄 رقم المستند</label>
                                      <div className="w-full bg-amber-50 border border-amber-200 rounded-lg px-2 py-1.5 text-[11px] font-black text-amber-900 text-center truncate min-h-[28px] flex items-center justify-center font-mono">
                                        {entry.documentNumber && entry.documentNumber !== '' ? entry.documentNumber : <span className="text-slate-300">—</span>}
                                      </div>
                                    </div>
                                  </div>

                                  <div className="grid grid-cols-2 gap-2">
                                    {([
                                      { label: 'تاريخ الطلب', key: 'applicationDate' },
                                      { label: 'الممثلية في', key: 'embassy' },
                                      { label: 'عدد مرات الدخول', key: 'entriesCount' },
                                      { label: 'اسم الجهة الطالبة', key: 'requesterName' },
                                      { label: 'الاسم', key: 'applicantName' },
                                      { label: 'Name', key: 'nameEnglish' },
                                      { label: 'نوع الجواز', key: 'passportType' },
                                      { label: 'تاريخ الانتهاء', key: 'passportExpiry' },
                                      { label: 'تاريخ الميلاد', key: 'birthDate' },
                                      { label: 'مكان الميلاد', key: 'birthPlace' },
                                      { label: 'الجنسية الحالية', key: 'currentNationality' },
                                      { label: 'الجنس', key: 'gender' },
                                      { label: 'المهنة', key: 'profession' },
                                      { label: 'الغرض', key: 'purpose' },
                                      { label: 'حالة الطلب', key: 'statusText' },
                                      { label: 'حالة اصدار الشهادة الصحية', key: 'healthCertStatus' },
                                    ] as { label: string; key: keyof typeof entry }[]).map(f => (
                                      <div key={f.key} className="space-y-0.5 text-right overflow-hidden">
                                        <label className="text-[8px] font-black text-slate-400 uppercase tracking-wider block truncate">{f.label}</label>
                                        <div className="w-full bg-slate-50 border border-slate-100 rounded-lg px-2 py-1.5 text-[10px] font-bold text-slate-700 text-center truncate min-h-[28px] flex items-center justify-center">
                                          {(entry[f.key] as string) || <span className="text-slate-300">—</span>}
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </div>

                                {/* ── Images Upload Section (All Tab) ── */}
                                <div className="border-t border-slate-100 pt-3 space-y-3">
                                  <span className="text-[8px] font-black text-blue-600 uppercase tracking-widest bg-blue-50 border border-blue-100 px-2 py-0.5 rounded-full">
                                    {lang === 'ar' ? 'المستندات والصور' : 'Documents & Images'}
                                  </span>
                                  <div className="grid grid-cols-2 gap-3">
                                    {/* Print Image — download only (auto-captured from MOFA print button) */}
                                    <div className="space-y-1.5">
                                      <p className="text-[9px] font-black text-slate-500 text-right">🖨️ {lang === 'ar' ? 'صورة البرنت' : 'Print Image'}</p>
                                      <button
                                        onClick={(e) => { e.stopPropagation(); if (entry.printImageUrl) { downloadUrl(entry.printImageUrl, `برنت_${entry.passportNumber||''}_${(entry.applicantName||'').replace(/\s+/g,'_')}.pdf`); } }}
                                        className={`flex flex-col items-center justify-center gap-1.5 w-full border-2 rounded-xl h-28 transition-all group ${entry.printImageUrl ? 'border-blue-300 hover:border-blue-500 bg-blue-50/60 hover:bg-blue-100 cursor-pointer' : 'border-blue-100 bg-blue-50/20 cursor-not-allowed opacity-60'}`}
                                        title={entry.printImageUrl ? (lang === 'ar' ? 'تنزيل PDF البرنت' : 'Download Print PDF') : (lang === 'ar' ? 'لم يتم جلب الملف بعد' : 'File not fetched yet')}
                                      >
                                        {entry.printImageUrl?.startsWith('tg:') ? <span className="text-lg">📨</span> : <Download size={20} className={`transition-colors ${entry.printImageUrl ? 'text-blue-500 group-hover:text-blue-700' : 'text-blue-200'}`} />}
                                        <span className={`text-[9px] font-bold transition-colors text-center px-1 ${entry.printImageUrl ? 'text-blue-500 group-hover:text-blue-700' : 'text-blue-200'}`}>
                                          {entry.printImageUrl ? (lang === 'ar' ? 'تنزيل PDF البرنت' : 'Download Print PDF') : (lang === 'ar' ? 'لم يتم الجلب بعد' : 'Not fetched yet')}
                                        </span>
                                      </button>
                                    </div>
                                    {/* Visa Image */}
                                    <div className="space-y-1.5">
                                      <p className="text-[9px] font-black text-slate-500 text-right">🎫 {lang === 'ar' ? 'صورة التأشيرة' : 'Visa Image'}</p>
                                      <button
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          const visaUrl = entry.visaImageUrl || entry.visaImageBase64;
                                          if (visaUrl) {
                                            downloadUrl(visaUrl, `تاشيره_${entry.passportNumber||''}_${(entry.applicantName||'').replace(/\s+/g,'_')}.pdf`);
                                          }
                                        }}
                                        className={`flex flex-col items-center justify-center gap-1.5 w-full border-2 rounded-xl h-28 transition-all group ${(entry.visaImageUrl||entry.visaImageBase64) ? 'border-emerald-300 hover:border-emerald-500 bg-emerald-50/60 hover:bg-emerald-100 cursor-pointer' : 'border-emerald-100 bg-emerald-50/20 cursor-not-allowed opacity-60'}`}
                                      >
                                        <Download size={20} className={`transition-colors ${(entry.visaImageUrl||entry.visaImageBase64) ? 'text-emerald-500 group-hover:text-emerald-700' : 'text-emerald-200'}`} />
                                        <span className={`text-[9px] font-bold transition-colors text-center px-1 ${(entry.visaImageUrl||entry.visaImageBase64) ? 'text-emerald-500 group-hover:text-emerald-700' : 'text-emerald-200'}`}>
                                          {(entry.visaImageUrl||entry.visaImageBase64) ? (lang === 'ar' ? 'تنزيل PDF التأشيرة' : 'Download Visa PDF') : (lang === 'ar' ? 'لم يتم الجلب بعد' : 'Not fetched yet')}
                                        </span>
                                      </button>
                                    </div>
                                  </div>
                                </div>

                              </div>
                            </div>
                                  </div>
                                </motion.div>
                              )}
                            </AnimatePresence>
                          </motion.div>
                          );
                        })}
                      </AnimatePresence>
                    </div>
                  )}
                </div>}
              </div>
            )}
          </>
        )}
            </React.Suspense>
          </main>
        </div>
      )}

      {/* ==================== GLOBAL INTERACTIVE MODALS ==================== */}
      <AppModals
        lang={lang}
        t={t}
        isAdminLoginOpen={isAdminLoginOpen}
        setIsAdminLoginOpen={setIsAdminLoginOpen}
        adminUsernameInput={adminUsernameInput}
        setAdminUsernameInput={setAdminUsernameInput}
        adminPinInput={adminPinInput}
        setAdminPinInput={setAdminPinInput}
        adminPinError={adminPinError}
        setAdminPinError={setAdminPinError}
        handleLoginSubmit={handleLoginSubmit}
        isAddOpen={isAddOpen}
        setIsAddOpen={setIsAddOpen}
        addForm={addForm}
        setAddForm={setAddForm}
        offices={offices}
        saveEntry={saveEntry}
        isBulkOpen={isBulkOpen}
        setIsBulkOpen={setIsBulkOpen}
        bulkText={bulkText}
        setBulkText={setBulkText}
        parsing={parsing}
        handleBulkParse={handleBulkParse}
        mofaFetchModalOpen={mofaFetchModalOpen}
        setMofaFetchModalOpen={setMofaFetchModalOpen}
        mofaFetchEntry={mofaFetchEntry}
        setMofaFetchEntry={setMofaFetchEntry}
        mofaCheckingOverlay={mofaCheckingOverlay}
        setMofaCheckingOverlay={setMofaCheckingOverlay}
        mofaSaveSuccessAnim={mofaSaveSuccessAnim}
        setMofaSaveSuccessAnim={setMofaSaveSuccessAnim}
        mofaCheckingStep={mofaCheckingStep}
        setMofaCheckingStep={setMofaCheckingStep}
        mofaCheckingError={mofaCheckingError}
        setMofaCheckingError={setMofaCheckingError}
        mofaResult={mofaResult}
        setMofaResult={setMofaResult}
        mofaError={mofaError}
        setMofaError={setMofaError}
        mofaInitLoading={mofaInitLoading}
        mofaSessionId={mofaSessionId}
        mofaSelectedNat={mofaSelectedNat}
        setMofaSelectedNat={setMofaSelectedNat}
        mofaNationalities={mofaNationalities}
        mofaSelectedVt={mofaSelectedVt}
        setMofaSelectedVt={setMofaSelectedVt}
        mofaVisaTypes={mofaVisaTypes}
        mofaSelectedEmb={mofaSelectedEmb}
        setMofaSelectedEmb={setMofaSelectedEmb}
        mofaEmbassies={mofaEmbassies}
        mofaCaptchaCode={mofaCaptchaCode}
        setMofaCaptchaCode={setMofaCaptchaCode}
        mofaCaptchaKey={mofaCaptchaKey}
        mofaLoading={mofaLoading}
        loadMofaSession={loadMofaSession}
        handleMofaLiveSubmit={handleMofaLiveSubmit}
        handleSaveMofaParsedData={handleSaveMofaParsedData}
        ksaVisaModalOpen={ksaVisaModalOpen}
        setKsaVisaModalOpen={setKsaVisaModalOpen}
        ksaVisaEntry={ksaVisaEntry}
        setKsaVisaEntry={setKsaVisaEntry}
        ksaVisaSuccess={ksaVisaSuccess}
        setKsaVisaSuccess={setKsaVisaSuccess}
        ksaVisaError={ksaVisaError}
        setKsaVisaError={setKsaVisaError}
        ksaVisaCheckingOverlay={ksaVisaCheckingOverlay}
        ksaVisaSaveSuccessAnim={ksaVisaSaveSuccessAnim}
        ksaVisaCheckingError={ksaVisaCheckingError}
        setKsaVisaCheckingError={setKsaVisaCheckingError}
        ksaVisaCheckingStep={ksaVisaCheckingStep}
        ksaVisaInitLoading={ksaVisaInitLoading}
        ksaVisaNumber={ksaVisaNumber}
        setKsaVisaNumber={setKsaVisaNumber}
        ksaVisaPassportNumber={ksaVisaPassportNumber}
        setKsaVisaPassportNumber={setKsaVisaPassportNumber}
        ksaVisaCaptchaCode={ksaVisaCaptchaCode}
        setKsaVisaCaptchaCode={setKsaVisaCaptchaCode}
        ksaVisaCaptchaKey={ksaVisaCaptchaKey}
        ksaVisaSessionId={ksaVisaSessionId}
        ksaVisaLoading={ksaVisaLoading}
        loadKsaVisaSession={loadKsaVisaSession}
        handleKsaVisaSubmit={handleKsaVisaSubmit}
        isExportSectionModalOpen={isExportSectionModalOpen}
        setIsExportSectionModalOpen={setIsExportSectionModalOpen}
        entries={entries}
        isToday={isToday}
        exportSectionToCSV={exportSectionToCSV}
        exportSectionToPDF={exportSectionToPDF}
        isSelectMode={isSelectMode}
        selectedEntryIds={selectedEntryIds}
        exitSelectMode={exitSelectMode}
        showBulkMenu={showBulkMenu}
        setShowBulkMenu={setShowBulkMenu}
        handleBulkRefresh={handleBulkRefresh}
        handleBulkHealthCert={handleBulkHealthCert}
        setShowTransferModal={setShowTransferModal}
        setBulkTransferOfficeId={setBulkTransferOfficeId}
        handleBulkArchive={handleBulkArchive}
        handleBulkDelete={handleBulkDelete}
        txStatusModalEntry={txStatusModalEntry}
        setTxStatusModalEntry={setTxStatusModalEntry}
        txManageMode={txManageMode}
        setTxManageMode={setTxManageMode}
        txStatusOptions={txStatusOptions}
        txStatusForm={txStatusForm}
        setTxStatusForm={setTxStatusForm}
        txTypeDot={txTypeDot}
        txTypeBg={txTypeBg}
        txEditingOptId={txEditingOptId}
        setTxEditingOptId={setTxEditingOptId}
        txEditLabel={txEditLabel}
        setTxEditLabel={setTxEditLabel}
        txEditType={txEditType}
        setTxEditType={setTxEditType}
        saveTxStatusOptions={saveTxStatusOptions}
        txNewLabel={txNewLabel}
        setTxNewLabel={setTxNewLabel}
        txNewType={txNewType}
        setTxNewType={setTxNewType}
        txNewIsolate={txNewIsolate}
        setTxNewIsolate={setTxNewIsolate}
        txNewHasStatsCard={txNewHasStatsCard}
        setTxNewHasStatsCard={setTxNewHasStatsCard}
        saveTxStatus={saveTxStatus}
        showTransferModal={showTransferModal}
        handleBulkTransfer={handleBulkTransfer}
        archiveConfirmEntry={archiveConfirmEntry}
        setArchiveConfirmEntry={setArchiveConfirmEntry}
        archiveUploading={archiveUploading}
        showConfirm={showConfirm}
        handleMoveToArchive={handleMoveToArchive}
        handlePermDeleteEntry={handlePermDeleteEntry}
        alertModal={alertModal}
        setAlertModal={setAlertModal}
        healthCertOverlay={healthCertOverlay}
        setHealthCertOverlay={setHealthCertOverlay}
        healthCertOverlayName={healthCertOverlayName}
        healthCertOverlayStep={healthCertOverlayStep}
        setHealthCertOverlayStep={setHealthCertOverlayStep}
        healthCertCancelRef={healthCertCancelRef}
        checkingHealthCertIds={checkingHealthCertIds}
        setCheckingHealthCertIds={setCheckingHealthCertIds}
        healthCertEntryIdRef={healthCertEntryIdRef}
        confirmModal={confirmModal}
        setConfirmModal={setConfirmModal}
      />
    </div>
  );
}
