import React, { useState, useEffect } from 'react';
import { motion, AnimatePresence } from 'motion/react';

interface DevStats {
  totalEntries: number;
  totalOffices: number;
  totalUsers: number;
  dataSize: string;
}

interface SysInfo {
  uptime: number;
  nodeVersion: string;
  platform: string;
  memUsed: string;
  memTotal: string;
  cpuLoad: string;
}

const DEV_SESSION_KEY = '__dev_auth__';

const IconShield = () => (
  <svg width="26" height="26" viewBox="0 0 24 24" fill="none">
    <path d="M12 2L3 6v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V6l-9-4z" fill="#7c3aed" opacity="0.2" />
    <path d="M12 2L3 6v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V6l-9-4z" stroke="#7c3aed" strokeWidth="1.5" strokeLinejoin="round" />
    <path d="M9 12l2 2 4-4" stroke="#a78bfa" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const IconGrid = ({ size = 16 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <rect x="3" y="3" width="7" height="7" rx="1" /><rect x="14" y="3" width="7" height="7" rx="1" />
    <rect x="14" y="14" width="7" height="7" rx="1" /><rect x="3" y="14" width="7" height="7" rx="1" />
  </svg>
);

const IconSettings = ({ size = 16 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <circle cx="12" cy="12" r="3" />
    <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z" />
  </svg>
);

const IconAlert = ({ size = 16 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
    <line x1="12" y1="9" x2="12" y2="13" /><line x1="12" y1="17" x2="12.01" y2="17" />
  </svg>
);

const IconRefresh = ({ size = 14 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <polyline points="23 4 23 10 17 10" /><path d="M20.49 15a9 9 0 1 1-.49-4.49" />
  </svg>
);

const IconLogout = ({ size = 14 }: { size?: number }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5">
    <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
    <polyline points="16,17 21,12 16,7" /><line x1="21" y1="12" x2="9" y2="12" />
  </svg>
);

const IconEye = ({ open }: { open: boolean }) => open ? (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94" />
    <path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19" />
    <line x1="1" y1="1" x2="23" y2="23" />
  </svg>
) : (
  <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
    <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" /><circle cx="12" cy="12" r="3" />
  </svg>
);

type Section = 'overview' | 'settings' | 'danger';

const NAV_ITEMS: { key: Section; label: string; icon: React.ReactNode }[] = [
  { key: 'overview', label: 'نظرة عامة', icon: <IconGrid /> },
  { key: 'settings', label: 'الإعدادات', icon: <IconSettings /> },
  { key: 'danger', label: 'منطقة الخطر', icon: <IconAlert /> },
];

function StatCard({ label, value, color }: { label: string; value: string | number; color: string }) {
  return (
    <div className="bg-[#0d0d1a] border border-[#1e1e3a] rounded-2xl p-4 flex flex-col gap-2">
      <p className="text-xs text-[#4a4a6a]">{label}</p>
      <p className="text-2xl font-bold" style={{ color }}>{value}</p>
    </div>
  );
}

function InfoCell({ label, value }: { label: string; value: string }) {
  return (
    <div className="bg-[#0d0d1a] rounded-xl p-3">
      <p className="text-[10px] text-[#4a4a6a] mb-1">{label}</p>
      <p className="text-[#a78bfa] font-bold text-sm">{value}</p>
    </div>
  );
}

function Card({ title, subtitle, children }: { title: string; subtitle?: string; children: React.ReactNode }) {
  return (
    <div className="bg-[#111120] border border-[#1e1e3a] rounded-2xl p-5 space-y-4">
      <div>
        <h3 className="text-white font-bold text-sm">{title}</h3>
        {subtitle && <p className="text-[#4a4a6a] text-xs mt-0.5">{subtitle}</p>}
      </div>
      {children}
    </div>
  );
}

function DevInput({ label, type = 'text', value, onChange, placeholder }: {
  label: string; type?: string; value: string;
  onChange: (v: string) => void; placeholder?: string;
}) {
  const [show, setShow] = useState(false);
  const isPass = type === 'password';
  return (
    <div>
      <label className="text-[#4a4a6a] text-xs block mb-1.5">{label}</label>
      <div className="relative">
        <input
          type={isPass && !show ? 'password' : 'text'}
          value={value}
          onChange={e => onChange(e.target.value)}
          placeholder={placeholder}
          autoComplete="off"
          className="w-full bg-[#0d0d1a] border border-[#1e1e3a] focus:border-[#7c3aed] rounded-xl px-4 py-3 text-white text-sm outline-none transition-colors placeholder-[#2a2a4a]"
        />
        {isPass && (
          <button
            type="button"
            onClick={() => setShow(s => !s)}
            className="absolute left-3 top-1/2 -translate-y-1/2 text-[#4a4a6a] hover:text-[#7c3aed] transition-colors"
          >
            <IconEye open={show} />
          </button>
        )}
      </div>
    </div>
  );
}

export default function DevPanel() {
  const [isLoggedIn, setIsLoggedIn] = useState(() => sessionStorage.getItem(DEV_SESSION_KEY) === 'true');
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  const [loginLoading, setLoginLoading] = useState(false);
  const [showLoginPass, setShowLoginPass] = useState(false);

  const [activeSection, setActiveSection] = useState<Section>('overview');
  const [stats, setStats] = useState<DevStats | null>(null);
  const [sysInfo, setSysInfo] = useState<SysInfo | null>(null);

  const [newDevUser, setNewDevUser] = useState('');
  const [newDevPass, setNewDevPass] = useState('');
  const [credMsg, setCredMsg] = useState('');

  useEffect(() => {
    if (isLoggedIn) fetchStats();
  }, [isLoggedIn]);

  const devToken = () => sessionStorage.getItem('__dev_token__') || '';

  async function fetchStats() {
    try {
      const r = await fetch('/api/dev/stats', { headers: { 'x-dev-token': devToken() } });
      if (r.ok) {
        const d = await r.json();
        setStats(d.stats);
        setSysInfo(d.sys);
      }
    } catch {}
  }

  async function handleLogin(e: React.FormEvent) {
    e.preventDefault();
    setLoginError('');
    setLoginLoading(true);
    try {
      const r = await fetch('/api/dev/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: username.trim(), password: password.trim() }),
      });
      const d = await r.json();
      if (r.ok && d.success) {
        sessionStorage.setItem(DEV_SESSION_KEY, 'true');
        sessionStorage.setItem('__dev_token__', d.token);
        setIsLoggedIn(true);
      } else {
        setLoginError(d.error || 'بيانات الدخول غير صحيحة');
      }
    } catch {
      setLoginError('تعذر الاتصال بالسيرفر');
    } finally {
      setLoginLoading(false);
    }
  }

  async function handleChangeCredentials() {
    setCredMsg('');
    if (!newDevUser.trim() || !newDevPass.trim()) {
      setCredMsg('error:يرجى إدخال اسم المستخدم وكلمة المرور');
      return;
    }
    try {
      const r = await fetch('/api/dev/credentials', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-dev-token': devToken() },
        body: JSON.stringify({ username: newDevUser.trim(), password: newDevPass.trim() }),
      });
      if (r.ok) {
        setCredMsg('success:تم تحديث بيانات الدخول — ستُسجَّل خروجاً الآن');
        setNewDevUser('');
        setNewDevPass('');
        setTimeout(() => {
          sessionStorage.clear();
          setIsLoggedIn(false);
        }, 2200);
      } else {
        setCredMsg('error:فشل تحديث البيانات');
      }
    } catch {
      setCredMsg('error:خطأ في الاتصال بالسيرفر');
    }
  }

  async function handleResetEntries() {
    if (!window.confirm('⚠️ سيتم حذف جميع بيانات العملاء نهائياً. هل أنت متأكد؟')) return;
    if (!window.confirm('تأكيد أخير — لا يمكن التراجع عن هذا الإجراء!')) return;
    try {
      const r = await fetch('/api/dev/reset-entries', {
        method: 'POST',
        headers: { 'x-dev-token': devToken() },
      });
      if (r.ok) {
        alert('✅ تم مسح جميع بيانات العملاء');
        fetchStats();
      }
    } catch {}
  }

  function handleLogout() {
    sessionStorage.clear();
    setIsLoggedIn(false);
    setUsername('');
    setPassword('');
  }

  const formatUptime = (s: number) => {
    const h = Math.floor(s / 3600);
    const m = Math.floor((s % 3600) / 60);
    return `${h} ساعة ${m} دقيقة`;
  };

  const FONT = "'Cairo', 'Tajawal', 'Inter', ui-sans-serif, system-ui, sans-serif";

  // ─── LOGIN ─────────────────────────────────────────────────────────────────
  if (!isLoggedIn) {
    return (
      <div
        className="min-h-screen bg-[#070710] flex items-center justify-center p-4"
        style={{ fontFamily: FONT }}
        dir="rtl"
      >
        <motion.div
          initial={{ opacity: 0, y: 24 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.4 }}
          className="w-full max-w-[380px]"
        >
          {/* الشعار */}
          <div className="flex flex-col items-center mb-8 select-none">
            <div className="w-16 h-16 rounded-[20px] bg-[#1a1030] border border-[#7c3aed]/25 flex items-center justify-center mb-4 shadow-xl shadow-[#7c3aed]/10">
              <IconShield />
            </div>
            <h1 className="text-white text-lg font-bold">لوحة المطور</h1>
            <p className="text-[#3a3a5a] text-xs mt-1">منطقة مقيّدة — للمصرح لهم فقط</p>
          </div>

          {/* نموذج الدخول */}
          <div className="bg-[#0e0e1c] border border-[#1a1a30] rounded-[24px] p-6 shadow-2xl shadow-black/60">
            <form onSubmit={handleLogin} className="space-y-4">
              <div>
                <label className="text-[#3a3a5a] text-xs block mb-2">اسم المستخدم</label>
                <input
                  type="text"
                  autoComplete="off"
                  value={username}
                  onChange={e => { setUsername(e.target.value); setLoginError(''); }}
                  placeholder="اسم المستخدم"
                  className="w-full bg-[#070710] border border-[#1a1a30] focus:border-[#7c3aed]/60 rounded-[14px] px-4 py-3.5 text-white text-sm outline-none transition-all duration-200 placeholder-[#252538]"
                />
              </div>

              <div>
                <label className="text-[#3a3a5a] text-xs block mb-2">كلمة المرور</label>
                <div className="relative">
                  <input
                    type={showLoginPass ? 'text' : 'password'}
                    autoComplete="off"
                    value={password}
                    onChange={e => { setPassword(e.target.value); setLoginError(''); }}
                    placeholder="••••••••"
                    className="w-full bg-[#070710] border border-[#1a1a30] focus:border-[#7c3aed]/60 rounded-[14px] px-4 py-3.5 text-white text-sm outline-none transition-all duration-200 placeholder-[#252538] pl-12"
                  />
                  <button
                    type="button"
                    onClick={() => setShowLoginPass(s => !s)}
                    className="absolute left-4 top-1/2 -translate-y-1/2 text-[#3a3a5a] hover:text-[#7c3aed] transition-colors"
                  >
                    <IconEye open={showLoginPass} />
                  </button>
                </div>
              </div>

              <AnimatePresence>
                {loginError && (
                  <motion.div
                    initial={{ opacity: 0, height: 0 }}
                    animate={{ opacity: 1, height: 'auto' }}
                    exit={{ opacity: 0, height: 0 }}
                    className="overflow-hidden"
                  >
                    <div className="text-red-400 text-xs bg-red-950/30 border border-red-900/30 rounded-xl px-3 py-2.5 flex items-center gap-2">
                      <svg width="12" height="12" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
                      {loginError}
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>

              <button
                type="submit"
                disabled={loginLoading}
                className="w-full py-3.5 mt-1 bg-[#7c3aed] hover:bg-[#6d28d9] active:scale-[0.98] disabled:opacity-50 text-white font-bold text-sm rounded-[14px] transition-all duration-200 shadow-lg shadow-[#7c3aed]/20 flex items-center justify-center gap-2"
              >
                {loginLoading ? (
                  <>
                    <svg className="animate-spin" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3">
                      <path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/>
                    </svg>
                    جاري التحقق...
                  </>
                ) : 'دخول'}
              </button>
            </form>
          </div>

          <p className="text-center text-[#1a1a2a] text-[10px] mt-6 select-none">
            الوصول غير المصرح به محظور تماماً
          </p>
        </motion.div>
      </div>
    );
  }

  // ─── الداشبورد ─────────────────────────────────────────────────────────────
  const credMsgParts = credMsg.split(':');
  const credMsgType = credMsgParts[0];
  const credMsgText = credMsgParts.slice(1).join(':');

  return (
    <div
      className="min-h-screen bg-[#070710] text-white flex flex-col"
      style={{ fontFamily: FONT }}
      dir="rtl"
    >
      {/* ── الشريط العلوي ── */}
      <header className="sticky top-0 z-30 bg-[#0e0e1c]/90 backdrop-blur-md border-b border-[#1a1a30] px-4 md:px-6 h-14 flex items-center justify-between flex-shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-7 h-7 rounded-[10px] bg-[#1a1030] border border-[#7c3aed]/30 flex items-center justify-center flex-shrink-0">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none">
              <path d="M12 2L3 6v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V6l-9-4z" fill="#7c3aed" opacity="0.4"/>
              <path d="M12 2L3 6v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V6l-9-4z" stroke="#7c3aed" strokeWidth="1.5"/>
            </svg>
          </div>
          <span className="text-white font-bold text-sm">لوحة المطور</span>
          <span className="text-[#7c3aed] text-[9px] bg-[#7c3aed]/10 border border-[#7c3aed]/20 px-2 py-0.5 rounded-full">خاص</span>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={fetchStats}
            className="w-8 h-8 rounded-xl bg-[#1a1a30]/60 hover:bg-[#1a1a30] border border-[#1a1a30] flex items-center justify-center text-[#4a4a6a] hover:text-[#7c3aed] transition-all"
            title="تحديث"
          >
            <IconRefresh size={13} />
          </button>
          <button
            onClick={handleLogout}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-[#1a1a30]/60 hover:bg-red-950/40 border border-[#1a1a30] hover:border-red-900/40 text-[#4a4a6a] hover:text-red-400 text-xs font-bold transition-all"
          >
            <IconLogout size={12} />
            <span className="hidden sm:block">خروج</span>
          </button>
        </div>
      </header>

      {/* ── المحتوى الرئيسي ── */}
      <div className="flex flex-1 min-h-0">

        {/* ── القائمة الجانبية — كمبيوتر ── */}
        <aside className="hidden md:flex flex-col w-52 flex-shrink-0 border-l border-[#1a1a30] bg-[#0a0a18] py-5 px-3 gap-1">
          {NAV_ITEMS.map(item => (
            <button
              key={item.key}
              onClick={() => setActiveSection(item.key)}
              className={`flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-sm font-bold transition-all duration-200 w-full text-right ${
                activeSection === item.key
                  ? 'bg-[#7c3aed]/15 text-[#a78bfa] border border-[#7c3aed]/25 shadow-sm'
                  : 'text-[#3a3a5a] hover:text-[#6a6a8a] hover:bg-[#111128]'
              }`}
            >
              <span className={activeSection === item.key ? 'text-[#a78bfa]' : 'text-[#3a3a5a]'}>
                {item.icon}
              </span>
              <span>{item.label}</span>
              {activeSection === item.key && (
                <span className="mr-auto w-1.5 h-1.5 rounded-full bg-[#7c3aed]" />
              )}
            </button>
          ))}
        </aside>

        {/* ── المحتوى ── */}
        <main className="flex-1 overflow-y-auto px-4 md:px-6 py-5 pb-24 md:pb-6">
          <AnimatePresence mode="wait">
            <motion.div
              key={activeSection}
              initial={{ opacity: 0, y: 8 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -8 }}
              transition={{ duration: 0.18 }}
              className="space-y-5 max-w-3xl mx-auto"
            >
              {/* ── نظرة عامة ── */}
              {activeSection === 'overview' && (
                <>
                  <div>
                    <h2 className="text-white font-bold text-base">نظرة عامة</h2>
                    <p className="text-[#3a3a5a] text-xs mt-0.5">إحصائيات النظام</p>
                  </div>

                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                    <StatCard label="إجمالي العملاء" value={stats?.totalEntries ?? '—'} color="#a78bfa" />
                    <StatCard label="المكاتب" value={stats?.totalOffices ?? '—'} color="#60a5fa" />
                    <StatCard label="المستخدمون" value={stats?.totalUsers ?? '—'} color="#34d399" />
                    <StatCard label="حجم البيانات" value={stats?.dataSize ?? '—'} color="#f59e0b" />
                  </div>

                  <Card title="معلومات السيرفر" subtitle="بيانات تشغيل الخادم في الوقت الفعلي">
                    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5">
                      <InfoCell label="مدة التشغيل" value={sysInfo ? formatUptime(sysInfo.uptime) : '—'} />
                      <InfoCell label="إصدار Node" value={sysInfo?.nodeVersion ?? '—'} />
                      <InfoCell label="النظام" value={sysInfo?.platform ?? '—'} />
                      <InfoCell label="الذاكرة المستخدمة" value={sysInfo?.memUsed ?? '—'} />
                      <InfoCell label="إجمالي الذاكرة" value={sysInfo?.memTotal ?? '—'} />
                      <InfoCell label="حِمل المعالج" value={sysInfo ? sysInfo.cpuLoad + '' : '—'} />
                    </div>
                  </Card>
                </>
              )}

              {/* ── الإعدادات ── */}
              {activeSection === 'settings' && (
                <>
                  <div>
                    <h2 className="text-white font-bold text-base">الإعدادات</h2>
                    <p className="text-[#3a3a5a] text-xs mt-0.5">إعدادات المطور</p>
                  </div>

                  <Card title="تغيير بيانات الدخول" subtitle="سيتم تسجيل خروجك تلقائياً بعد الحفظ">
                    <div className="space-y-3">
                      <DevInput
                        label="اسم المستخدم الجديد"
                        value={newDevUser}
                        onChange={setNewDevUser}
                        placeholder="اسم المستخدم"
                      />
                      <DevInput
                        label="كلمة المرور الجديدة"
                        type="password"
                        value={newDevPass}
                        onChange={setNewDevPass}
                        placeholder="كلمة المرور"
                      />

                      <AnimatePresence>
                        {credMsg && (
                          <motion.div
                            initial={{ opacity: 0, height: 0 }}
                            animate={{ opacity: 1, height: 'auto' }}
                            exit={{ opacity: 0, height: 0 }}
                            className="overflow-hidden"
                          >
                            <p className={`text-xs font-bold py-1 ${credMsgType === 'success' ? 'text-green-400' : 'text-red-400'}`}>
                              {credMsgType === 'success' ? '✅ ' : '❌ '}{credMsgText}
                            </p>
                          </motion.div>
                        )}
                      </AnimatePresence>

                      <button
                        onClick={handleChangeCredentials}
                        className="px-5 py-2.5 bg-[#7c3aed] hover:bg-[#6d28d9] active:scale-[0.98] text-white text-sm font-bold rounded-xl transition-all shadow-md shadow-[#7c3aed]/15"
                      >
                        حفظ البيانات
                      </button>
                    </div>
                  </Card>
                </>
              )}

              {/* ── منطقة الخطر ── */}
              {activeSection === 'danger' && (
                <>
                  <div className="flex items-center gap-2">
                    <span className="text-red-500/70"><IconAlert size={14} /></span>
                    <div>
                      <h2 className="text-white font-bold text-base">منطقة الخطر</h2>
                      <p className="text-[#3a3a5a] text-xs mt-0.5">إجراءات لا يمكن التراجع عنها</p>
                    </div>
                  </div>

                  <div className="bg-[#111120] border border-red-950/50 rounded-2xl p-5 space-y-4">
                    <div className="flex items-start gap-3">
                      <div className="w-8 h-8 rounded-xl bg-red-950/40 border border-red-900/30 flex items-center justify-center flex-shrink-0 mt-0.5">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#f87171" strokeWidth="2">
                          <polyline points="3,6 5,6 21,6"/><path d="M19,6l-1,14a2,2,0,0,1-2,2H8a2,2,0,0,1-2-2L5,6"/><path d="M10,11v6"/><path d="M14,11v6"/><path d="M9,6V4a1,1,0,0,1,1-1h4a1,1,0,0,1,1,1v2"/>
                        </svg>
                      </div>
                      <div className="flex-1">
                        <h3 className="text-red-400 font-bold text-sm mb-1">حذف جميع بيانات العملاء</h3>
                        <p className="text-[#4a4a6a] text-xs leading-relaxed">
                          سيتم حذف جميع سجلات العملاء والتأشيرات نهائياً من قاعدة البيانات. هذا الإجراء لا يمكن التراجع عنه بأي شكل.
                        </p>
                      </div>
                    </div>
                    <button
                      onClick={handleResetEntries}
                      className="px-5 py-2.5 bg-red-950/40 hover:bg-red-900/50 active:scale-[0.98] border border-red-900/40 hover:border-red-700/50 text-red-400 text-sm font-bold rounded-xl transition-all"
                    >
                      ⚠️ حذف جميع البيانات
                    </button>
                  </div>

                  <div className="bg-[#111120] border border-[#1a1a30] rounded-2xl p-5">
                    <h3 className="text-[#3a3a5a] font-bold text-sm mb-1">إجراءات إضافية</h3>
                    <p className="text-[#2a2a4a] text-xs">ستُضاف إجراءات أخرى هنا حسب الطلب.</p>
                  </div>
                </>
              )}
            </motion.div>
          </AnimatePresence>
        </main>
      </div>

      {/* ── شريط التنقل السفلي — جوال ── */}
      <nav className="md:hidden fixed bottom-0 left-0 right-0 z-30 bg-[#0e0e1c]/95 backdrop-blur-xl border-t border-[#1a1a30] px-2 py-2">
        <div className="flex items-center justify-around max-w-sm mx-auto">
          {NAV_ITEMS.map(item => {
            const active = activeSection === item.key;
            return (
              <button
                key={item.key}
                onClick={() => setActiveSection(item.key)}
                className={`flex flex-col items-center gap-1 px-4 py-2 rounded-2xl transition-all duration-200 flex-1 mx-0.5 ${
                  active
                    ? 'bg-[#7c3aed]/15 border border-[#7c3aed]/20'
                    : 'hover:bg-[#111128]'
                }`}
              >
                <span className={active ? 'text-[#a78bfa]' : 'text-[#3a3a5a]'}>
                  {item.key === 'overview' ? <IconGrid size={18} /> : item.key === 'settings' ? <IconSettings size={18} /> : <IconAlert size={18} />}
                </span>
                <span className={`text-[10px] font-bold ${active ? 'text-[#a78bfa]' : 'text-[#3a3a5a]'}`}>
                  {item.label}
                </span>
              </button>
            );
          })}
        </div>
      </nav>
    </div>
  );
}
