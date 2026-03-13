/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useMemo, useEffect } from 'react';
import { 
  LayoutDashboard, 
  TrendingUp, 
  Factory, 
  BellRing, 
  Settings, 
  Download, 
  AlertTriangle, 
  RefreshCw, 
  Search,
  ArrowUp,
  ArrowDown,
  BarChart3,
  PieChart,
  History,
  ChevronDown,
  FileSpreadsheet,
  CheckCircle2,
  X,
  Clock,
  Users,
  Target,
  Layers,
  LogOut,
  ShieldCheck,
  ArrowRight
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { 
  BarChart, 
  Bar, 
  ComposedChart,
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer, 
  Cell,
  PieChart as RePieChart,
  Pie,
  LineChart,
  Line,
  LabelList,
  ReferenceLine,
  Legend,
  AreaChart,
  Area
} from 'recharts';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import Login from './Login';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Mock Data based on User's Jan 2026 Excel ---

const MOCK_DATA = [];

const TREND_HISTORY = [
  { month: 'Sep 2025', issues: 12, closeRate: 85, avgTime: 9.2 },
  { month: 'Oct 2025', issues: 28, closeRate: 65, avgTime: 9.0 },
  { month: 'Nov 2025', issues: 22, closeRate: 72, avgTime: 9.4 },
  { month: 'Dec 2025', issues: 12, closeRate: 58, avgTime: 10.3 },
  { month: 'Jan 2026', issues: 21, closeRate: 52, avgTime: 10.7 },
  { month: 'Feb 2026', issues: 16, closeRate: 63, avgTime: 8.4 },
  { month: 'Mar 2026', issues: 8, closeRate: 88, avgTime: 10.4 },
];

const PERFORMANCE_DATA = [
  { name: 'Ian Li', task: 12, speed: 4.2 },
  { name: 'Colin Gu', task: 15, speed: 3.8 },
  { name: 'Jasper Ou', task: 8, speed: 5.1 },
  { name: 'Alen', task: 10, speed: 4.5 },
  { name: 'Rock Wen', task: 6, speed: 6.2 },
];

// --- Components ---

const SidebarItem = ({ icon: Icon, label, active = false, onClick }: { icon: any, label: string, active?: boolean, onClick: () => void }) => (
  <button 
    onClick={onClick}
    className={cn(
      "w-full flex items-center gap-3 px-4 py-3 rounded-lg transition-all duration-200",
      active ? "sidebar-item-active text-slate-900" : "text-slate-500 hover:bg-slate-100"
    )}
  >
    <Icon size={20} />
    <span className="text-sm font-medium text-left">{label}</span>
  </button>
);

const KPICard = ({ title, value, target, trend, trendValue, color = "slate" }: any) => (
  <div className="apple-card p-4 flex flex-col justify-between min-h-[140px]">
    <div className="flex justify-between items-start">
      <p className="text-xs font-medium text-slate-500">{title}</p>
      {target && <span className="text-[10px] text-slate-400">{target}</span>}
    </div>
    <h2 className={cn("text-3xl font-bold tracking-tight mt-2", color === "rose" ? "text-rose-600" : "text-slate-900")}>
      {value}
    </h2>
    <div className="mt-2">
      {trend && (
        <p className="text-[10px] flex items-center gap-1 text-slate-500">
          {trend === "up" ? <ArrowUp size={12} className="text-emerald-600" /> : <ArrowDown size={12} className="text-rose-600" />}
          {trendValue}
        </p>
      )}
      {!trend && trendValue && <p className="text-[10px] text-slate-500">{trendValue}</p>}
    </div>
  </div>
);

const TRANSLATIONS: Record<string, any> = {
  '中': {
    title: '质量管理驾驶舱',
    subtitle: '制造决策支持平台',
    overview: '总览',
    trend: '趋势分析',
    production: '生产看板',
    repairDashboard: '返修率看板',
    oobDashboard: 'OOB看板',
    closeDashboard: '问题关闭率看板',
    rootCauseDashboard: '根因分类看板',
    performance: '绩效看板',
    data: '详细数据',
    importExcel: '导入表格',
    importRepairExcel: '导入返修率表',
    exportReport: '导出报表',
    riskWarning: '风险预警',
    resetData: '重置数据',
    monthlyIssues: '月度问题总数',
    oobRate: 'OOB 合规率',
    repairRate: '累计返修率',
    allProducts: '所有产品',
    totalRepairRate: '总返修率',
    repairDashboardTitle: '按产品分类统计返修率趋势',
    repairImportHint: '返修率看板仅显示单独导入的返修率 Excel 数据，总表导入不会在这里呈现。',
    repairImportEmpty: '请先在本模块导入返修率 Excel 表格',
    clearRepairData: '清空返修率表',
    repairMonth: '日期',
    repairUploadSuccess: '返修率表导入成功',
    oobModuleTitle: 'OOB 看板',
    oobModuleHint: '基于总表中的 OOB 字段、产品线、主责部门与初筛主责部门统计趋势。',
    oobRateBoardTitle: '按产品分类统计 OOB 不合格率',
    oobCountBoardTitle: '按产品分类统计 OOB 问题产品数量分布',
    mainDeptOobTitle: '主责部门看板 OOB 产品趋势',
    initialDeptOobTitle: '初筛主责部门看板 OOB 产品趋势',
    oobDashboardEmpty: '请先导入包含 OOB 信息的总表数据',
    closeModuleTitle: '问题关闭率看板',
    closeModuleHint: '基于总表中的问题数量、问题关闭数量、产品分类、主责部门与初筛主责部门统计关闭率趋势。',
    overallCloseRateTitle: '市场问题关闭率趋势',
    productCloseRateTitle: '按产品分类统计市场问题关闭率趋势',
    productOobCloseRateTitle: '按产品分类统计市场 OOB 问题关闭率趋势',
    mainDeptCloseRateTitle: '主责部门问题关闭率趋势',
    initialDeptCloseRateTitle: '初筛主责部门问题关闭率趋势',
    closeDashboardEmpty: '请先导入包含问题数量和问题关闭数量的总表数据',
    rootCauseModuleTitle: '根因分类看板',
    rootCauseModuleHint: '独立承载根因分类相关的 4 个质量分析看板与 Excel 报表导出能力。',
    rootCauseModuleIntroTitle: '模块已预留',
    rootCauseModuleIntroDesc: '当前先完成独立模块入口与页面骨架，后续将在这里接入根因分类分析逻辑和报表导出流程。',
    rootCauseModuleBoardsTitle: '计划接入的 4 个看板',
    rootCauseBoard1: '单根因问题产品总数趋势',
    rootCauseBoard2: '月度根因问题占比分布',
    rootCauseBoard3: '进一步追根因分类为空，问题分析分类分布',
    rootCauseBoard4: '按根因分类统计产品型号问题产品数量',
    rootCauseModuleDataTitle: '数据来源',
    rootCauseModuleDataDesc: '将复用 Excel 明细导入链路，后续补充字段映射、统计口径校验和导出报表功能。',
    rootCauseModuleEmpty: '请先导入包含创建时间、根因分类、问题分析类型和产品型号的总表数据',
    rootCausePeriodLabel: '分析月份',
    rootCauseBlankLabel: '(空白)',
    rootCauseTotalLabel: '总计',
    rootCauseCumulativeLabel: '累计占比',
    rootCauseTableTitle: '根因分类与产品型号透视明细',
    closeRate: '问题关闭率',
    overdue: '逾期未闭环',
    target: '目标',
    benchmark: '基准',
    sla: 'SLA: 7天',
    unprocessed: '未处理 > 7天',
    filter: '数据筛选',
    timeRange: '时间范围',
    productLine: '产品线',
    rootCause: '根因分类',
    dept: '产品归属',
    oobStatus: 'OOB 状态',
    ranking: '产品问题数量分布',
    distribution: '根因分布 TOP 7',
    emergency: '紧急风险预警',
    all: '全部',
    days7: '7天',
    days30: '30天',
    days180: '180天',
    days365: '365天',
    roboticArm: '机械臂',
    robot: '机器人',
    joint: '关节',
    others: '其他配件',
    pkgTrans: '包装与运输类原因',
    material: '物料与来料类原因',
    field: '现场应用与环境类原因',
    assembly: '生产装配与工艺类原因',
    maint: '维护与操作类原因',
    design: '设计开发类原因',
    req: '需求类问题',
    deptRuiJu: '睿矩产品研发中心',
    deptWeiHan: '微悍产品研发中心',
    deptRuiYou: '睿友产品研发中心',
    deptAdv: '先进制造与工业化中心',
    oobDamage: '开箱损',
    nonOobDamage: '非开箱损',
    issueCount: '问题数量',
    oobAbnormal: 'OOB 异常',
    trendTitle: '月度问题趋势 (含目标线)',
    closeRateTrend: '问题关闭率趋势',
    avgTimeTrend: '平均处理时长趋势',
    prodAbnormal: '生产异常总数',
    highFreq: '高频故障型号',
    mainCause: '主要根因分类',
    prodAssembly: '生产装配异常',
    modelDist: '产品型号故障分布',
    causeDepth: '故障类型深度分析',
    teamEff: '团队处理效率',
    avgResp: '平均响应速度 (天)',
    todoAlert: '待办预警 (Top 5)',
    rawJanData: '原始质量数据',
    searchPlaceholder: '搜索型号、根因...',
    model: '产品型号',
    count: '数量',
    status: '状态',
    closed: '已闭环',
    processing: '处理中',
    oobDmgLabel: '开箱损',
    nonOobDmgLabel: '非开箱损',
    month: '月份',
    year: '年份',
    selectYear: '请选择年份',
    selectMonth: '请选择月份',
    jan: '1月',
    feb: '2月',
    mar: '3月',
    apr: '4月',
    may: '5月',
    jun: '6月',
    jul: '7月',
    aug: '8月',
    sep: '9月',
    oct: '10月',
    nov: '11月',
    dec: '12月',
    vsLastMonth: '较上月',
    customerName: '客户名称',
    coverage: '覆盖',
    models: '个产品型号',
    ratio: '占比',
    brand: '质量智能 BI',
    cases: '起'
  },
  'EN': {
    title: 'Quality Management Cockpit',
    subtitle: 'Manufacturing Decision Support Platform',
    overview: 'Overview',
    trend: 'Trend Analysis',
    production: 'Production Dashboard',
    repairDashboard: 'Repair Rate Dashboard',
    oobDashboard: 'OOB Dashboard',
    closeDashboard: 'Issue Close Rate Dashboard',
    rootCauseDashboard: 'Root Cause Dashboard',
    performance: 'Performance Dashboard',
    data: 'Detailed Data',
    importExcel: 'Import Excel',
    importRepairExcel: 'Import Repair Excel',
    exportReport: 'Export Report',
    riskWarning: 'Risk Warning',
    resetData: 'Reset Data',
    monthlyIssues: 'Monthly Total Issues',
    oobRate: 'OOB Compliance Rate',
    repairRate: 'Total Repair Rate',
    allProducts: 'All Products',
    totalRepairRate: 'Total Repair Rate',
    repairDashboardTitle: 'Repair Rate Trend by Product Line',
    repairImportHint: 'The repair dashboard only renders data imported from its dedicated repair-rate Excel file.',
    repairImportEmpty: 'Import a dedicated repair-rate Excel file in this module first',
    clearRepairData: 'Clear Repair Data',
    repairMonth: 'Period',
    repairUploadSuccess: 'Repair-rate Excel imported successfully',
    oobModuleTitle: 'OOB Dashboards',
    oobModuleHint: 'Built from the issue sheet OOB flag, product line, main department, and initial department fields.',
    oobRateBoardTitle: 'OOB Defect Rate by Product Line',
    oobCountBoardTitle: 'OOB Product Count by Product Line',
    mainDeptOobTitle: 'Main Department OOB Trend',
    initialDeptOobTitle: 'Initial Department OOB Trend',
    oobDashboardEmpty: 'Import issue data with OOB fields first',
    closeModuleTitle: 'Issue Close Rate Dashboard',
    closeModuleHint: 'Close-rate trends calculated from issue quantity, closed quantity, product line, main department, and initial department in the main workbook.',
    overallCloseRateTitle: 'Market Issue Close Rate Trend',
    productCloseRateTitle: 'Market Issue Close Rate Trend by Product Line',
    productOobCloseRateTitle: 'Market OOB Issue Close Rate Trend by Product Line',
    mainDeptCloseRateTitle: 'Main Department Issue Close Rate Trend',
    initialDeptCloseRateTitle: 'Initial Department Issue Close Rate Trend',
    closeDashboardEmpty: 'Import issue data with issue quantity and closed quantity first',
    rootCauseModuleTitle: 'Root Cause Dashboard',
    rootCauseModuleHint: 'A standalone module reserved for four root-cause analysis dashboards and Excel report export.',
    rootCauseModuleIntroTitle: 'Module Reserved',
    rootCauseModuleIntroDesc: 'This release adds the standalone entry and page scaffold first. The root-cause analytics and report export flow will be connected here next.',
    rootCauseModuleBoardsTitle: 'Planned Dashboards',
    rootCauseBoard1: 'Single Root Cause Issue Count Trend',
    rootCauseBoard2: 'Monthly Root Cause Ratio Distribution',
    rootCauseBoard3: 'Problem Analysis Type Distribution When Root Cause Is Blank',
    rootCauseBoard4: 'Issue Count by Root Cause and Product Model',
    rootCauseModuleDataTitle: 'Data Source',
    rootCauseModuleDataDesc: 'It will reuse the Excel detail import flow, then add field mapping, metric validation, and report export.',
    rootCauseModuleEmpty: 'Import issue data with created time, root cause, analysis type, and product model fields first',
    rootCausePeriodLabel: 'Analysis Period',
    rootCauseBlankLabel: '(Blank)',
    rootCauseTotalLabel: 'Total',
    rootCauseCumulativeLabel: 'Cumulative Share',
    rootCauseTableTitle: 'Root Cause by Product Model Pivot',
    closeRate: 'Issue Close Rate',
    overdue: 'Overdue Unclosed',
    target: 'Target',
    benchmark: 'Benchmark',
    sla: 'SLA: 7 Days',
    unprocessed: 'Unprocessed > 7 Days',
    filter: 'Data Filter',
    timeRange: 'Time Range',
    productLine: 'Product Line',
    rootCause: 'Root Cause',
    dept: 'Department',
    oobStatus: 'OOB Status',
    ranking: 'Product Issue Quantity Distribution',
    distribution: 'Root Cause TOP 7',
    emergency: 'Emergency Alerts',
    all: 'All',
    days7: '7 Days',
    days30: '30 Days',
    days180: '180 Days',
    days365: '365 Days',
    roboticArm: 'Robotic Arm',
    robot: 'Robot',
    joint: 'Joint',
    others: 'Other Accessories',
    pkgTrans: 'Packaging & Transport',
    material: 'Material & Incoming',
    field: 'Field Application & Env',
    assembly: 'Production Assembly & Process',
    maint: 'Maintenance & Operation',
    design: 'Design & Development',
    req: 'Requirement Issues',
    deptRuiJu: 'RuiJu R&D Center',
    deptWeiHan: 'WeiHan R&D Center',
    deptRuiYou: 'RuiYou R&D Center',
    deptAdv: 'Adv Manufacturing Center',
    oobDamage: 'OOB Damage',
    nonOobDamage: 'Non-OOB Damage',
    issueCount: 'Issue Quantity',
    oobAbnormal: 'OOB Abnormal',
    trendTitle: 'Monthly Issue Trend (with Target)',
    closeRateTrend: 'Close Rate Trend',
    avgTimeTrend: 'Avg Processing Time Trend',
    prodAbnormal: 'Total Prod Anomalies',
    highFreq: 'High Freq Failure Model',
    mainCause: 'Main Root Cause',
    prodAssembly: 'Prod Assembly Anomaly',
    modelDist: 'Model Failure Dist',
    causeDepth: 'Cause Depth Analysis',
    teamEff: 'Team Efficiency',
    avgResp: 'Avg Response (Days)',
    todoAlert: 'Todo Alerts (Top 5)',
    rawJanData: 'Jan 2026 Raw Data',
    searchPlaceholder: 'Search model, cause...',
    model: 'Model',
    count: 'Count',
    status: 'Status',
    closed: 'Closed',
    processing: 'Processing',
    oobDmgLabel: 'OOB Dmg',
    nonOobDmgLabel: 'Non-OOB',
    month: 'Month',
    year: 'Year',
    selectYear: 'Select year',
    selectMonth: 'Select month',
    jan: 'Jan',
    feb: 'Feb',
    mar: 'Mar',
    apr: 'Apr',
    may: 'May',
    jun: 'Jun',
    jul: 'Jul',
    aug: 'Aug',
    sep: 'Sep',
    oct: 'Oct',
    nov: 'Nov',
    dec: 'Dec',
    vsLastMonth: 'vs Last Month',
    customerName: 'Customer Name',
    coverage: 'Coverage',
    models: 'Models',
    ratio: 'Ratio',
    brand: 'Quality BI',
    cases: 'Cases'
  }
};

export default function App() {
  const [token, setToken] = useState<string | null>(localStorage.getItem('token'));
  const [username, setUsername] = useState<string | null>(localStorage.getItem('username'));
  const [activeTab, setActiveTab] = useState('overview');
  const [lang, setLang] = useState<'中' | 'EN'>('中');
  const [selectedYear, setSelectedYear] = useState<number | 'all'>('all');
  const [selectedMonth, setSelectedMonth] = useState<number | 'all'>('all');
  const [selectedProductLine, setSelectedProductLine] = useState<string>('all');
  const [selectedCause, setSelectedCause] = useState<string>('all');
  const [selectedDept, setSelectedDept] = useState<string>('all');
  const [selectedOob, setSelectedOob] = useState<string>('all');
  const [selectedRootCausePeriod, setSelectedRootCausePeriod] = useState<string>('');
  const [data, setData] = useState<any[]>(MOCK_DATA);
  const [repairData, setRepairData] = useState<any[]>([]);
  const [isKickedOut, setIsKickedOut] = useState(false);
  const t = (key: string) => TRANSLATIONS[lang][key] || key;
  const [importStatus, setImportStatus] = useState<{ show: boolean; message: string; type: 'success' | 'error' }>({ show: false, message: '', type: 'success' });
  const fileInputRef = useRef<HTMLInputElement>(null);
  const repairFileInputRef = useRef<HTMLInputElement>(null);
  const dataTableTopScrollRef = useRef<HTMLDivElement>(null);
  const dataTableScrollRef = useRef<HTMLDivElement>(null);
  const [dataTableScrollWidth, setDataTableScrollWidth] = useState<number>(0);

  const [backendStatus, setBackendStatus] = useState<'checking' | 'online' | 'offline'>('checking');
  const [isLoading, setIsLoading] = useState(false);

  const handleLogout = () => {
    localStorage.removeItem('token');
    localStorage.removeItem('username');
    setToken(null);
    setUsername(null);
  };

  const handleLogin = (t: string, u: string) => {
    localStorage.setItem('token', t);
    localStorage.setItem('username', u);
    setToken(t);
    setUsername(u);
    setIsKickedOut(false);
  };

  useEffect(() => {
    if (activeTab !== 'data') return;
    const topEl = dataTableTopScrollRef.current;
    const bodyEl = dataTableScrollRef.current;
    if (!topEl || !bodyEl) return;

    let syncing = false;

    const syncFromTop = () => {
      if (syncing) return;
      syncing = true;
      bodyEl.scrollLeft = topEl.scrollLeft;
      syncing = false;
    };

    const syncFromBody = () => {
      if (syncing) return;
      syncing = true;
      topEl.scrollLeft = bodyEl.scrollLeft;
      syncing = false;
    };

    topEl.addEventListener('scroll', syncFromTop);
    bodyEl.addEventListener('scroll', syncFromBody);
    topEl.scrollLeft = bodyEl.scrollLeft;

    return () => {
      topEl.removeEventListener('scroll', syncFromTop);
      bodyEl.removeEventListener('scroll', syncFromBody);
    };
  }, [activeTab]);

  useEffect(() => {
    if (activeTab !== 'data') return;
    const el = dataTableScrollRef.current;
    if (!el) return;

    const update = () => setDataTableScrollWidth(el.scrollWidth);
    update();
    window.addEventListener('resize', update);
    return () => window.removeEventListener('resize', update);
  }, [activeTab, data, lang]);

  const yearOptions = useMemo(() => {
    const years = Array.from(
      new Set(
        data
          .map((d: any) => {
            if (typeof d.year === 'number' && d.year > 0) return d.year;
            const text = String(d.complaintDate || d.createdAt || '').trim();
            const m = text.match(/(\d{4})年/);
            if (m) return Number(m[1]);
            const dt = new Date(text);
            if (!isNaN(dt.getTime())) return dt.getFullYear();
            return undefined;
          })
          .filter(Boolean) as number[]
      )
    ).sort((a, b) => b - a);

    if (years.length > 0) return years;
    return [2026, 2025];
  }, [data]);

  // Session monitor
  const checkSession = async () => {
    if (!token) return;
    try {
      const res = await fetch('/api/auth/session-status', {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      const result = await res.json();
      if (res.status === 401 && result.error === 'SESSION_EXPIRED_CONCURRENT') {
        setIsKickedOut(true);
      }
    } catch (err) {
      console.error('Session monitor error:', err);
    }
  };

  useEffect(() => {
    if (!token) return;

    // Check immediately on mount/token change
    checkSession();

    const interval = setInterval(checkSession, 2000); // Check every 2 seconds

    // Check when window regains focus
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        checkSession();
      }
    };
    document.addEventListener('visibilitychange', handleVisibilityChange);

    return () => {
      clearInterval(interval);
      document.removeEventListener('visibilitychange', handleVisibilityChange);
    };
  }, [token]);

  // Check session when switching tabs to ensure immediate feedback
  useEffect(() => {
    if (token) checkSession();
  }, [activeTab]);

  // Load data from backend on mount
  const loadDataFromBackend = async () => {
    if (!token) return;
    try {
      setIsLoading(true);
      const res = await fetch('/api/issues', {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      const result = await res.json();
      
      if (res.status === 401 || res.status === 403) {
        if (result.error === 'SESSION_EXPIRED_CONCURRENT') {
          setIsKickedOut(true);
        } else {
          handleLogout();
        }
        return;
      }
      
      if (result.success && result.data.length > 0) {
        setData(result.data);
      }
    } catch (err) {
      console.error('Failed to load data from backend:', err);
    } finally {
      setIsLoading(false);
    }
  };

  const loadRepairDataFromBackend = async () => {
    if (!token) return;
    try {
      const res = await fetch('/api/repair-rates', {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      const result = await res.json();

      if (res.status === 401 || res.status === 403) {
        if (result.error === 'SESSION_EXPIRED_CONCURRENT') {
          setIsKickedOut(true);
        } else {
          handleLogout();
        }
        return;
      }

      if (result.success) {
        setRepairData(result.data || []);
      }
    } catch (err) {
      console.error('Failed to load repair data from backend:', err);
    }
  };

  useEffect(() => {
    if (!token) return;
    // Health check
    fetch('/api/health')
      .then(res => res.json())
      .then(data => {
        if (data.status === 'ok') {
          setBackendStatus('online');
          // Load persisted data
          loadDataFromBackend();
          loadRepairDataFromBackend();
        } else {
          setBackendStatus('offline');
        }
      })
      .catch(() => setBackendStatus('offline'));
  }, []);

  // --- Computed Data ---
  const filteredData = useMemo(() => {
    return data.filter(d => {
      if (selectedYear === 'all' || selectedMonth === 'all') return false;
      const matchYear = d.year === selectedYear;
      const matchMonth = d.month === selectedMonth;
      const matchProductLine = selectedProductLine === 'all' || d.productLine === t(selectedProductLine) || d.productLine === selectedProductLine;
      const matchCause = selectedCause === 'all' || d.cause === t(selectedCause) || d.cause === selectedCause;
      const matchDept = selectedDept === 'all' || d.dept === t(selectedDept) || d.dept === selectedDept;
      const matchOob = selectedOob === 'all' || (selectedOob === 'oobDamage' ? d.oob >= 1 : d.oob === 0);
      return matchYear && matchMonth && matchProductLine && matchCause && matchDept && matchOob;
    });
  }, [selectedYear, selectedMonth, selectedProductLine, selectedCause, selectedDept, selectedOob, data, lang]);

  const kpiStats = useMemo(() => {
    const filterByYearMonth = (dList: any[], year: number | 'all', month: number | 'all') => {
      return dList.filter(d => {
        if (year === 'all' || month === 'all') return false;
        const matchYear = d.year === year;
        const matchMonth = d.month === month;
        const matchProductLine = selectedProductLine === 'all' || d.productLine === t(selectedProductLine) || d.productLine === selectedProductLine;
        const matchCause = selectedCause === 'all' || d.cause === t(selectedCause) || d.cause === selectedCause;
        const matchDept = selectedDept === 'all' || d.dept === t(selectedDept) || d.dept === selectedDept;
        const matchOob = selectedOob === 'all' || (selectedOob === 'oobDamage' ? d.oob >= 1 : d.oob === 0);
        return matchYear && matchMonth && matchProductLine && matchCause && matchDept && matchOob;
      });
    };

    const currData = filterByYearMonth(data, selectedYear, selectedMonth);
    
    // Calculate current metrics
    const currTotalIssues = currData.length;
    const currOobIssues = currData.filter(d => (d.oob || 0) >= 1).length;
    const currClosedIssues = currData.filter(d => d.closed === 1).length;
    const currOverdue = currData.filter(d => d.closed === 0).length;
    
    const currOobRate = currTotalIssues > 0 ? (currOobIssues / currTotalIssues) * 100 : 0;
    const currCloseRate = currTotalIssues > 0 ? (currClosedIssues / currTotalIssues) * 100 : 0;
    const currRepairRate = (currTotalIssues / 2500) * 100;

    // Calculate previous metrics for trend
    let prevData: any[] = [];
    if (selectedYear === 'all' || selectedMonth === 'all') {
      prevData = [];
    } else {
      const prevMonth = selectedMonth === 1 ? 12 : selectedMonth - 1;
      const prevYear = selectedMonth === 1 ? selectedYear - 1 : selectedYear;
      prevData = filterByYearMonth(data, prevYear, prevMonth);
    }

    const prevTotalIssues = prevData.length;
    const prevOobIssues = prevData.filter(d => (d.oob || 0) >= 1).length;
    const prevClosedIssues = prevData.filter(d => d.closed === 1).length;
    const prevOverdue = prevData.filter(d => d.closed === 0).length;

    const prevOobRate = prevTotalIssues > 0 ? (prevOobIssues / prevTotalIssues) * 100 : 0;
    const prevCloseRate = prevTotalIssues > 0 ? (prevClosedIssues / prevTotalIssues) * 100 : 0;
    const prevRepairRate = (prevTotalIssues / 2500) * 100;

    const calcTrend = (curr: number, prev: number) => {
      if (prev === 0) return { direction: 'up' as const, value: curr > 0 ? '+100%' : '0%' };
      const diff = curr - prev;
      const pct = (diff / prev) * 100;
      return {
        direction: diff >= 0 ? 'up' as const : 'down' as const,
        value: `${diff >= 0 ? '+' : ''}${pct.toFixed(1)}%`
      };
    };

    return {
      issues: { val: currTotalIssues, trend: calcTrend(currTotalIssues, prevTotalIssues) },
      oob: { val: currOobRate.toFixed(2) + '%', trend: calcTrend(currOobRate, prevOobRate) },
      repair: { val: currRepairRate.toFixed(2) + '%', trend: calcTrend(currRepairRate, prevRepairRate) },
      close: { val: currCloseRate.toFixed(1) + '%', trend: calcTrend(currCloseRate, prevCloseRate) },
      overdue: { val: currOverdue, trend: calcTrend(currOverdue, prevOverdue) }
    };
  }, [selectedYear, selectedMonth, selectedProductLine, selectedCause, selectedDept, selectedOob, data, lang]);

  const dynamicTrendHistory = useMemo(() => {
    const months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    const monthKeys = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
    return months.map((m, idx) => {
      const monthData = data.filter(d => {
        if (selectedYear === 'all') return false;
        const matchYear = d.year === selectedYear;
        const matchMonth = d.month === m;
        const matchProductLine = selectedProductLine === 'all' || d.productLine === t(selectedProductLine) || d.productLine === selectedProductLine;
        const matchCause = selectedCause === 'all' || d.cause === t(selectedCause) || d.cause === selectedCause;
        const matchDept = selectedDept === 'all' || d.dept === t(selectedDept) || d.dept === selectedDept;
        const matchOob = selectedOob === 'all' || (selectedOob === 'oobDamage' ? d.oob >= 1 : d.oob === 0);
        return matchYear && matchMonth && matchProductLine && matchCause && matchDept && matchOob;
      });
      
      const totalIssues = monthData.length;
      const closedIssues = monthData.filter(d => d.closed === 1).length;
      const closeRate = totalIssues > 0 ? (closedIssues / totalIssues) * 100 : 0;
      
      return {
        month: t(monthKeys[idx]),
        issues: totalIssues,
        closeRate: closeRate,
        avgTime: totalIssues > 0 ? 7 + Math.sin(m) * 3 : 0
      };
    });
  }, [data, selectedYear, selectedProductLine, selectedCause, selectedDept, selectedOob, lang]);

  const repairRateTrendByMonth = useMemo(() => {
    const normalizeProductLine = (value: any) => {
      const v = String(value || '').trim();
      if (!v) return 'others';
      if (v === 'all' || v.includes('所有产品') || v.toLowerCase().includes('all products')) return 'all';
      if (!v) return 'others';
      if (v === 'roboticArm' || v.includes('机械臂')) return 'roboticArm';
      if (v === 'robot' || v.includes('机器人')) return 'robot';
      if (v === 'joint' || v.includes('关节')) return 'joint';
      if (v === 'others' || v.includes('其他')) return 'others';
      return 'others';
    };

    const monthLabel = (year: number, month: number) => lang === '中'
      ? `${year}年${month}月`
      : `${year}-${String(month).padStart(2, '0')}`;
    const computeTotalRate = (repairQty: number, warrantyQty: number, oobQty: number, shipQty: number, explicitRate: number) => {
      if (explicitRate > 0) return explicitRate;
      const repairRate = warrantyQty > 0 ? (repairQty / warrantyQty) * 100 : 0;
      const oobRate = shipQty > 0 ? (oobQty / shipQty) * 100 : 0;
      return repairRate * 0.4 + oobRate * 0.6;
    };

    const buckets = new Map<number, any>();
    const ensure = (orderKey: number, year: number, month: number) => {
      if (buckets.has(orderKey)) return buckets.get(orderKey);
      const blank = { repairQty: 0, oobQty: 0, shipQty: 0, warrantyQty: 0, explicitRate: 0, hasExplicit: false };
      const item = {
        orderKey,
        month: monthLabel(year, month),
        all: { ...blank },
        roboticArm: { ...blank },
        robot: { ...blank },
        joint: { ...blank },
        others: { ...blank },
      };
      buckets.set(orderKey, item);
      return item;
    };

    for (const row of repairData as any[]) {
      const year = Number(row.year) || 0;
      const month = Number(row.month) || 0;
      if (year <= 0 || month < 1 || month > 12) continue;

      const bucket = ensure(year * 100 + month, year, month);
      const productLine = normalizeProductLine(row.productLine);
      const repairQty = Number(row.repairCount) || 0;
      const oobQty = Number(row.oobDefectCount) || 0;
      const shipQty = Number(row.monthShipmentCount) || 0;
      const warrantyQty = Number(row.warrantyShipmentCount) || 0;
      const explicitRate = Number(row.totalRepairRate) || 0;

      if (productLine === 'all') {
        bucket.all.repairQty += repairQty;
        bucket.all.oobQty += oobQty;
        bucket.all.shipQty += shipQty;
        bucket.all.warrantyQty += warrantyQty;
        bucket.all.explicitRate += explicitRate;
        bucket.all.hasExplicit = true;
        continue;
      }

      bucket[productLine].repairQty += repairQty;
      bucket[productLine].oobQty += oobQty;
      bucket[productLine].shipQty += shipQty;
      bucket[productLine].warrantyQty += warrantyQty;
      bucket[productLine].explicitRate += explicitRate;
    }

    return Array.from(buckets.values())
      .sort((a, b) => a.orderKey - b.orderKey)
      .map((bucket) => ({
        month: bucket.month,
        all: bucket.all.hasExplicit
          ? computeTotalRate(bucket.all.repairQty, bucket.all.warrantyQty, bucket.all.oobQty, bucket.all.shipQty, bucket.all.explicitRate)
          : computeTotalRate(
              bucket.roboticArm.repairQty + bucket.robot.repairQty + bucket.joint.repairQty + bucket.others.repairQty,
              bucket.roboticArm.warrantyQty + bucket.robot.warrantyQty + bucket.joint.warrantyQty + bucket.others.warrantyQty,
              bucket.roboticArm.oobQty + bucket.robot.oobQty + bucket.joint.oobQty + bucket.others.oobQty,
              bucket.roboticArm.shipQty + bucket.robot.shipQty + bucket.joint.shipQty + bucket.others.shipQty,
              bucket.roboticArm.explicitRate + bucket.robot.explicitRate + bucket.joint.explicitRate + bucket.others.explicitRate
            ),
        roboticArm: computeTotalRate(bucket.roboticArm.repairQty, bucket.roboticArm.warrantyQty, bucket.roboticArm.oobQty, bucket.roboticArm.shipQty, bucket.roboticArm.explicitRate),
        robot: computeTotalRate(bucket.robot.repairQty, bucket.robot.warrantyQty, bucket.robot.oobQty, bucket.robot.shipQty, bucket.robot.explicitRate),
        joint: computeTotalRate(bucket.joint.repairQty, bucket.joint.warrantyQty, bucket.joint.oobQty, bucket.joint.shipQty, bucket.joint.explicitRate),
        others: computeTotalRate(bucket.others.repairQty, bucket.others.warrantyQty, bucket.others.oobQty, bucket.others.shipQty, bucket.others.explicitRate),
      }));
  }, [repairData, lang]);

  const oobDashboard = useMemo(() => {
    const normalizeProductLine = (value: any) => {
      const v = String(value || '').trim();
      if (!v) return 'others';
      if (v === 'roboticArm' || v.includes('机械臂')) return 'roboticArm';
      if (v === 'robot' || v.includes('机器人')) return 'robot';
      if (v === 'joint' || v.includes('关节')) return 'joint';
      if (v === 'others' || v.includes('其他')) return 'others';
      return 'others';
    };
    const monthLabel = (year: number, month: number) => lang === '中'
      ? `${year}年${month}月`
      : `${year}-${String(month).padStart(2, '0')}`;
    const sortedPeriods = Array.from(
      new Set(
        data
          .map((row: any) => {
            const year = Number(row.year) || 0;
            const month = Number(row.month) || 0;
            if (year <= 0 || month < 1 || month > 12) return '';
            return `${year}-${String(month).padStart(2, '0')}`;
          })
          .filter(Boolean)
      )
    ) as string[];
    sortedPeriods.sort();

    const rateMap = new Map<string, any>();
    const countMap = new Map<string, any>();
    const ensureBaseRow = (map: Map<string, any>, periodKey: string, label: string) => {
      if (!map.has(periodKey)) {
        map.set(periodKey, {
          month: label,
          all: 0,
          roboticArm: 0,
          robot: 0,
          joint: 0,
          others: 0,
          allShipment: 0,
          roboticArmShipment: 0,
          robotShipment: 0,
          jointShipment: 0,
          othersShipment: 0,
        });
      }
      return map.get(periodKey);
    };

    const mainDeptTotals: Record<string, number> = {};
    const initialDeptTotals: Record<string, number> = {};

    data.forEach((row: any) => {
      const year = Number(row.year) || 0;
      const month = Number(row.month) || 0;
      if (year <= 0 || month < 1 || month > 12 || Number(row.oob) < 1) return;

      const periodKey = `${year}-${String(month).padStart(2, '0')}`;
      const periodLabel = monthLabel(year, month);
      const productLine = normalizeProductLine(row.productLine);
      const quantity = Number(row.issueQuantity) || 0;
      const shipment = Number(row.productQuantity) || 0;

      const rateRow = ensureBaseRow(rateMap, periodKey, periodLabel);
      const countRow = ensureBaseRow(countMap, periodKey, periodLabel);

      rateRow[productLine] += quantity;
      rateRow.all += quantity;
      rateRow[`${productLine}Shipment`] += shipment;
      rateRow.allShipment += shipment;

      countRow[productLine] += quantity;
      countRow.all += quantity;

      const mainDept = String(row.mainDept || row.dept || '').trim() || '未指定';
      const initialDept = String(row.initialDept || '').trim() || '未指定';
      mainDeptTotals[mainDept] = (mainDeptTotals[mainDept] || 0) + quantity;
      initialDeptTotals[initialDept] = (initialDeptTotals[initialDept] || 0) + quantity;
    });

    const buildSeriesRows = (totals: Record<string, number>, field: 'mainDept' | 'initialDept') => {
      const topDepartments = Object.entries(totals)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 6)
        .map(([name]) => name);

      return sortedPeriods.map((periodKey) => {
        const [year, month] = periodKey.split('-').map(Number);
        const row: Record<string, any> = { month: monthLabel(year, month) };
        topDepartments.forEach((dept) => {
          row[dept] = 0;
        });

        data.forEach((item: any) => {
          if (Number(item.oob) < 1) return;
          const itemYear = Number(item.year) || 0;
          const itemMonth = Number(item.month) || 0;
          const itemPeriod = `${itemYear}-${String(itemMonth).padStart(2, '0')}`;
          if (itemPeriod !== periodKey) return;

          const deptName = String(field === 'mainDept' ? (item.mainDept || item.dept || '') : (item.initialDept || '')).trim() || '未指定';
          if (!topDepartments.includes(deptName)) return;
          row[deptName] += Number(item.issueQuantity) || 0;
        });

        return row;
      });
    };

    const rateTrend = sortedPeriods.map((periodKey) => {
      const [year, month] = periodKey.split('-').map(Number);
      const row = rateMap.get(periodKey) || ensureBaseRow(rateMap, periodKey, monthLabel(year, month));
      return {
        month: row.month,
        all: row.allShipment > 0 ? (row.all / row.allShipment) * 100 : 0,
        roboticArm: row.roboticArmShipment > 0 ? (row.roboticArm / row.roboticArmShipment) * 100 : 0,
        robot: row.robotShipment > 0 ? (row.robot / row.robotShipment) * 100 : 0,
        joint: row.jointShipment > 0 ? (row.joint / row.jointShipment) * 100 : 0,
        others: row.othersShipment > 0 ? (row.others / row.othersShipment) * 100 : 0,
      };
    });

    const countTrend = sortedPeriods.map((periodKey) => {
      const [year, month] = periodKey.split('-').map(Number);
      return countMap.get(periodKey) || ensureBaseRow(countMap, periodKey, monthLabel(year, month));
    });

    return {
      hasData: sortedPeriods.length > 0,
      rateTrend,
      countTrend,
      mainDeptTrend: buildSeriesRows(mainDeptTotals, 'mainDept'),
      initialDeptTrend: buildSeriesRows(initialDeptTotals, 'initialDept'),
      mainDepartments: Object.entries(mainDeptTotals).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([name]) => name),
      initialDepartments: Object.entries(initialDeptTotals).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([name]) => name),
    };
  }, [data, lang]);

  const closeRateDashboard = useMemo(() => {
    const normalizeProductLine = (value: any) => {
      const v = String(value || '').trim();
      if (!v) return 'others';
      if (v === 'roboticArm' || v.includes('机械臂')) return 'roboticArm';
      if (v === 'robot' || v.includes('机器人')) return 'robot';
      if (v === 'joint' || v.includes('关节')) return 'joint';
      if (v === 'others' || v.includes('其他')) return 'others';
      return 'others';
    };
    const monthLabel = (year: number, month: number) => lang === '中'
      ? `${year}年${month}月`
      : `${year}-${String(month).padStart(2, '0')}`;
    const sortedPeriods = Array.from(
      new Set(
        data
          .map((row: any) => {
            const year = Number(row.year) || 0;
            const month = Number(row.month) || 0;
            if (year <= 0 || month < 1 || month > 12) return '';
            return `${year}-${String(month).padStart(2, '0')}`;
          })
          .filter(Boolean)
      )
    ) as string[];
    sortedPeriods.sort();

    const ensureOverallRow = (map: Map<string, any>, periodKey: string, label: string) => {
      if (!map.has(periodKey)) {
        map.set(periodKey, { month: label, total: 0, closed: 0 });
      }
      return map.get(periodKey);
    };
    const ensureProductRow = (map: Map<string, any>, periodKey: string, label: string) => {
      if (!map.has(periodKey)) {
        map.set(periodKey, {
          month: label,
          roboticArmTotal: 0,
          roboticArmClosed: 0,
          robotTotal: 0,
          robotClosed: 0,
          jointTotal: 0,
          jointClosed: 0,
          othersTotal: 0,
          othersClosed: 0,
        });
      }
      return map.get(periodKey);
    };
    const toClosedQuantity = (row: any) => {
      const issueQty = Math.max(0, Number(row.issueQuantity) || 0);
      const closedQty = Number((row as any).closedQuantity);
      const normalizedClosedQty = Number.isFinite(closedQty) ? closedQty : (row.closed ? issueQty : 0);
      return Math.max(0, Math.min(issueQty, normalizedClosedQty));
    };
    const calcRate = (closedQty: number, totalQty: number) => totalQty > 0 ? (closedQty / totalQty) * 100 : 0;

    const overallMap = new Map<string, any>();
    const productMap = new Map<string, any>();
    const oobProductMap = new Map<string, any>();
    const mainDeptTotals: Record<string, number> = {};
    const initialDeptTotals: Record<string, number> = {};

    data.forEach((row: any) => {
      const year = Number(row.year) || 0;
      const month = Number(row.month) || 0;
      if (year <= 0 || month < 1 || month > 12) return;

      const periodKey = `${year}-${String(month).padStart(2, '0')}`;
      const periodLabel = monthLabel(year, month);
      const issueQty = Math.max(0, Number(row.issueQuantity) || 0);
      const closedQty = toClosedQuantity(row);
      const productLine = normalizeProductLine(row.productLine);
      const mainDept = String(row.mainDept || row.dept || '').trim() || '未指定';
      const initialDept = String(row.initialDept || '').trim() || '未指定';

      const overallRow = ensureOverallRow(overallMap, periodKey, periodLabel);
      overallRow.total += issueQty;
      overallRow.closed += closedQty;

      const productRow = ensureProductRow(productMap, periodKey, periodLabel);
      productRow[`${productLine}Total`] += issueQty;
      productRow[`${productLine}Closed`] += closedQty;

      if (Number(row.oob) >= 1) {
        const oobProductRow = ensureProductRow(oobProductMap, periodKey, periodLabel);
        oobProductRow[`${productLine}Total`] += issueQty;
        oobProductRow[`${productLine}Closed`] += closedQty;
      }

      mainDeptTotals[mainDept] = (mainDeptTotals[mainDept] || 0) + issueQty;
      initialDeptTotals[initialDept] = (initialDeptTotals[initialDept] || 0) + issueQty;
    });

    const buildProductRateTrend = (map: Map<string, any>) => {
      return sortedPeriods.map((periodKey) => {
        const [year, month] = periodKey.split('-').map(Number);
        const row = map.get(periodKey) || ensureProductRow(map, periodKey, monthLabel(year, month));
        return {
          month: row.month,
          roboticArm: calcRate(row.roboticArmClosed, row.roboticArmTotal),
          robot: calcRate(row.robotClosed, row.robotTotal),
          joint: calcRate(row.jointClosed, row.jointTotal),
          others: calcRate(row.othersClosed, row.othersTotal),
        };
      });
    };

    const buildDepartmentCloseTrend = (totals: Record<string, number>, field: 'mainDept' | 'initialDept') => {
      const topDepartments = Object.entries(totals)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 6)
        .map(([name]) => name);

      const rows = sortedPeriods.map((periodKey) => {
        const [year, month] = periodKey.split('-').map(Number);
        const row: Record<string, any> = { month: monthLabel(year, month) };
        const stats: Record<string, { total: number; closed: number }> = {};

        topDepartments.forEach((dept) => {
          stats[dept] = { total: 0, closed: 0 };
          row[dept] = 0;
        });

        data.forEach((item: any) => {
          const itemYear = Number(item.year) || 0;
          const itemMonth = Number(item.month) || 0;
          const itemPeriodKey = `${itemYear}-${String(itemMonth).padStart(2, '0')}`;
          if (itemPeriodKey !== periodKey) return;

          const deptName = String(
            field === 'mainDept'
              ? (item.mainDept || item.dept || '')
              : (item.initialDept || '')
          ).trim() || '未指定';

          if (!topDepartments.includes(deptName)) return;

          stats[deptName].total += Math.max(0, Number(item.issueQuantity) || 0);
          stats[deptName].closed += toClosedQuantity(item);
        });

        topDepartments.forEach((dept) => {
          row[dept] = calcRate(stats[dept].closed, stats[dept].total);
        });

        return row;
      });

      return {
        rows,
        departments: topDepartments,
      };
    };

    const mainDeptClose = buildDepartmentCloseTrend(mainDeptTotals, 'mainDept');
    const initialDeptClose = buildDepartmentCloseTrend(initialDeptTotals, 'initialDept');

    return {
      hasData: sortedPeriods.length > 0,
      overallTrend: sortedPeriods.map((periodKey) => {
        const [year, month] = periodKey.split('-').map(Number);
        const row = overallMap.get(periodKey) || ensureOverallRow(overallMap, periodKey, monthLabel(year, month));
        return {
          month: row.month,
          closeRate: calcRate(row.closed, row.total),
        };
      }),
      productTrend: buildProductRateTrend(productMap),
      oobProductTrend: buildProductRateTrend(oobProductMap),
      mainDeptTrend: mainDeptClose.rows,
      initialDeptTrend: initialDeptClose.rows,
      mainDepartments: mainDeptClose.departments,
      initialDepartments: initialDeptClose.departments,
    };
  }, [data, lang]);

  const rootCausePeriodOptions = useMemo(() => {
    const periodKeys = Array.from(
      new Set(
        data
          .map((row: any) => {
            const year = Number(row.year) || 0;
            const month = Number(row.month) || 0;
            if (year <= 0 || month < 1 || month > 12) return '';
            return `${year}-${String(month).padStart(2, '0')}`;
          })
          .filter(Boolean)
      )
    ) as string[];

    periodKeys.sort();

    const yearSet = new Set(periodKeys.map((key) => key.split('-')[0]));
    const singleYear = yearSet.size <= 1;

    return periodKeys.map((value) => {
      const [year, month] = value.split('-').map(Number);
      return {
        value,
        label: lang === '中'
          ? (singleYear ? `${month}月` : `${year}年${month}月`)
          : `${year}-${String(month).padStart(2, '0')}`,
      };
    });
  }, [data, lang]);

  useEffect(() => {
    if (!rootCausePeriodOptions.length) {
      if (selectedRootCausePeriod) setSelectedRootCausePeriod('');
      return;
    }

    const exists = rootCausePeriodOptions.some((item) => item.value === selectedRootCausePeriod);
    if (!exists) {
      setSelectedRootCausePeriod(rootCausePeriodOptions[rootCausePeriodOptions.length - 1].value);
    }
  }, [rootCausePeriodOptions, selectedRootCausePeriod]);

  const rootCauseDashboard = useMemo(() => {
    const blankLabel = t('rootCauseBlankLabel');
    const totalLabel = t('rootCauseTotalLabel');
    const colorPalette = ['#0f172a', '#2563eb', '#f97316', '#a855f7', '#0ea5e9', '#16a34a', '#eab308', '#ef4444'];
    const normalizeRootCause = (value: any) => {
      const text = String(value || '').trim();
      if (!text || text === '未分类') return blankLabel;
      return text;
    };
    const normalizeAnalysisType = (value: any) => {
      const text = String(value || '').trim();
      return text || blankLabel;
    };
    const normalizeModel = (value: any) => {
      const text = String(value || '').trim();
      return text || blankLabel;
    };
    const periodLabelMap = new Map(rootCausePeriodOptions.map((item) => [item.value, item.label]));

    const periodCauseCounts = new Map<string, Record<string, number>>();
    const causeTotals: Record<string, number> = {};
    const blankAnalysisCounts: Record<string, number> = {};
    const causeModelCounts: Record<string, Record<string, number>> = {};
    const modelTotals: Record<string, number> = {};

    data.forEach((row: any) => {
      const year = Number(row.year) || 0;
      const month = Number(row.month) || 0;
      if (year <= 0 || month < 1 || month > 12) return;

      const periodKey = `${year}-${String(month).padStart(2, '0')}`;
      const cause = normalizeRootCause(row.cause);
      const analysisType = normalizeAnalysisType((row as any).analysisType);
      const model = normalizeModel(row.model || (row as any).productModelPath);

      if (!periodCauseCounts.has(periodKey)) {
        periodCauseCounts.set(periodKey, {});
      }
      const causeCountMap = periodCauseCounts.get(periodKey)!;
      causeCountMap[cause] = (causeCountMap[cause] || 0) + 1;
      causeTotals[cause] = (causeTotals[cause] || 0) + 1;

      if (cause === blankLabel) {
        blankAnalysisCounts[analysisType] = (blankAnalysisCounts[analysisType] || 0) + 1;
      }

      if (!causeModelCounts[cause]) {
        causeModelCounts[cause] = {};
      }
      causeModelCounts[cause][model] = (causeModelCounts[cause][model] || 0) + 1;
      modelTotals[model] = (modelTotals[model] || 0) + 1;
    });

    const causeList = Object.entries(causeTotals)
      .sort((a, b) => b[1] - a[1])
      .map(([name]) => name);

    const trendData = rootCausePeriodOptions.map((period) => {
      const causeCountMap = periodCauseCounts.get(period.value) || {};
      const row: Record<string, any> = {
        month: period.label,
        total: Object.values(causeCountMap).reduce((sum, value) => sum + value, 0),
      };

      causeList.forEach((cause) => {
        row[cause] = causeCountMap[cause] || 0;
      });

      return row;
    });

    const selectedPeriodValue = selectedRootCausePeriod || rootCausePeriodOptions[rootCausePeriodOptions.length - 1]?.value || '';
    const selectedPeriodLabel = periodLabelMap.get(selectedPeriodValue) || '';
    const selectedPeriodCounts = periodCauseCounts.get(selectedPeriodValue) || {};
    const selectedPeriodTotal = Object.values(selectedPeriodCounts).reduce((sum, value) => sum + value, 0);

    let cumulativeShare = 0;
    const monthlyDistribution = Object.entries(selectedPeriodCounts)
      .sort((a, b) => b[1] - a[1])
      .map(([name, count]) => {
        const share = selectedPeriodTotal > 0 ? (count / selectedPeriodTotal) * 100 : 0;
        cumulativeShare += share;
        return {
          name,
          count,
          share,
          cumulativeShare,
        };
      });

    const blankAnalysisDistribution = Object.entries(blankAnalysisCounts)
      .sort((a, b) => b[1] - a[1])
      .map(([name, count]) => ({ name, count }));

    const modelList = Object.entries(modelTotals)
      .sort((a, b) => b[1] - a[1])
      .map(([name]) => name);
    const chartModels = modelList.slice(0, 6);
    const causePivotRows = causeList.map((cause) => {
      const row: Record<string, any> = { cause };
      modelList.forEach((model) => {
        row[model] = causeModelCounts[cause]?.[model] || 0;
      });
      row[totalLabel] = modelList.reduce((sum, model) => sum + (row[model] || 0), 0);
      return row;
    });
    const totalRow = modelList.reduce<Record<string, any>>((acc, model) => {
      acc[model] = causePivotRows.reduce((sum, row) => sum + (row[model] || 0), 0);
      return acc;
    }, { cause: totalLabel, [totalLabel]: causePivotRows.reduce((sum, row) => sum + (row[totalLabel] || 0), 0) });
    const causeModelChartData = causeList.map((cause) => {
      const row: Record<string, any> = { cause };
      chartModels.forEach((model) => {
        row[model] = causeModelCounts[cause]?.[model] || 0;
      });
      return row;
    });

    return {
      hasData: rootCausePeriodOptions.length > 0 && causeList.length > 0,
      blankLabel,
      totalLabel,
      colorPalette,
      selectedPeriodValue,
      selectedPeriodLabel,
      trendCauses: causeList,
      trendData,
      monthlyDistribution,
      blankAnalysisDistribution,
      chartModels,
      causeModelChartData,
      pivotModels: modelList,
      pivotRows: [...causePivotRows, totalRow],
    };
  }, [data, lang, rootCausePeriodOptions, selectedRootCausePeriod]);

  const dynamicPerformanceData = useMemo(() => {
    const creators: Record<string, { task: number, speed: number }> = {};
    filteredData.forEach(d => {
      const creator = d.creator || 'System';
      if (!creators[creator]) creators[creator] = { task: 0, speed: 4 + Math.random() * 2 };
      creators[creator].task += d.issueQuantity;
    });
    
    const result = Object.entries(creators).map(([name, stats]) => ({
      name,
      task: stats.task,
      speed: stats.speed
    })).sort((a, b) => b.task - a.task).slice(0, 5);

    return result.length > 0 ? result : PERFORMANCE_DATA;
  }, [filteredData]);

  const modelRanking = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => {
      counts[d.model] = (counts[d.model] || 0) + d.issueQuantity;
    });
    return Object.entries(counts)
      .map(([name, count]) => ({ name, count, value: (count / 10) * 100 }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);
  }, [filteredData]);

  const causeDistribution = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => {
      counts[d.cause] = (counts[d.cause] || 0) + d.issueQuantity;
    });
    const colors = ['#c2410c', '#d97706', '#475569', '#0071e3', '#1e3a8a', '#3b82f6', '#f59e0b'];
    return Object.entries(counts)
      .map(([name, value], idx) => ({ name, value, color: colors[idx % colors.length] }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 7);
  }, [filteredData]);

  const handleImportExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const bstr = event.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rawData: any[] = XLSX.utils.sheet_to_json(ws);

        const formatLongDate = (y: number, m: number, d: number) => `${y}年${m}月${d}日`;

        const normalizeDateText = (value: any) => {
          if (!value) return '';
          if (value instanceof Date) return formatLongDate(value.getFullYear(), value.getMonth() + 1, value.getDate());
          if (typeof value === 'number') {
            const parsed = XLSX.SSF.parse_date_code(value);
            if (parsed) return formatLongDate(parsed.y, parsed.m, parsed.d);
            return String(value);
          }
          const text = String(value).trim();
          const m = text.match(/(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
          if (m) return formatLongDate(Number(m[1]), Number(m[2]), Number(m[3]));
          return text;
        };

        const parseMonthFromDateLike = (value: any) => {
          if (!value) return 1;
          if (value instanceof Date) return value.getMonth() + 1;
          if (typeof value === 'number') {
            const parsed = XLSX.SSF.parse_date_code(value);
            if (parsed) return Math.min(12, Math.max(1, Number(parsed.m)));
            return 1;
          }
          const text = String(value);
          const m = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
          if (m) return Math.min(12, Math.max(1, Number(m[2])));
          const d = new Date(text);
          if (!isNaN(d.getTime())) return d.getMonth() + 1;
          return 1;
        };
        const parseYearFromDateLike = (value: any) => {
          if (!value) return 0;
          if (value instanceof Date) return value.getFullYear();
          if (typeof value === 'number') {
            const parsed = XLSX.SSF.parse_date_code(value);
            if (parsed) return Number(parsed.y) || 0;
            return 0;
          }
          const text = String(value);
          const m = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
          if (m) return Number(m[1]) || 0;
          const d = new Date(text);
          if (!isNaN(d.getTime())) return d.getFullYear();
          return 0;
        };
        
        // Map fields based on user's Excel image and requirements
        const processedData = rawData.map(row => {
          const productModelPath = String(row['产品型号'] || row['产品型号名称'] || '').trim();
          const segments = productModelPath
            .split('/')
            .map((s: string) => s.trim())
            .filter(Boolean);
          const level1 = segments[0] || '';

          const productLine = (() => {
            if (level1.includes('机械臂')) return 'roboticArm';
            if (level1.includes('机器人')) return 'robot';
            if (level1.includes('关节')) return 'joint';
            if (level1.includes('其他')) return 'others';
            return 'others';
          })();

          const issueQuantity = Number(row['问题数量']) || 0;
          const closedQuantity = Number(row['问题关闭数量']) || 0;

          const oobRaw = row['是否开箱损'];
          const oobText = typeof oobRaw === 'string' ? oobRaw.trim() : oobRaw;
          const oob = (() => {
            if (oobText === true || oobText === '是' || oobText === 'Y') return 1;
            if (typeof oobText === 'string') {
              if (oobText.includes('非开箱损')) return 0;
              if (oobText.includes('开箱损')) return 1;
            }
            return 0;
          })();

          return {
            year: parseYearFromDateLike(row['创建时间']),
            month: parseMonthFromDateLike(row['创建时间']),
            customerName: row['客户名称'] || row['标题'] || row['标题_1'] || '未知客户',
            productQuantity: Number(row['产品数量']) || 0,
            productModelPath: productModelPath || '未知型号',
            model: segments[segments.length - 1] || productModelPath || '未知型号',
            initialDept: String(row['问题初判主责部门'] || '').trim(),
            mainDept: String(row['问题主责部门'] || '').trim(),
            complaintDate: normalizeDateText(row['创建时间']),
            snDate: normalizeDateText(row['SN日期']),
            cause: row['根因分类'] || '未分类',
            analysisType: String(row['问题分析类型'] || '').trim(),
            dept: String(row['问题主责部门'] || row['产品归属'] || '未指定').trim(),
            productLine,
            issueQuantity,
            closedQuantity,
            closed: issueQuantity > 0 && closedQuantity >= issueQuantity ? 1 : 0,
            oob,
            creator: row['创建人'] || 'System'
          };
        });

        // Update frontend state immediately
        setData(prev => [...prev, ...processedData]);
        if (selectedYear === 'all') {
          const years = processedData.map((d: any) => d.year).filter((y: any) => typeof y === 'number' && y > 0) as number[];
          if (years.length > 0) {
            setSelectedYear(Math.max(...years));
            setSelectedMonth('all');
          }
        }

        // Persist to backend
        if (backendStatus === 'online') {
          try {
            const res = await fetch('/api/issues/bulk', {
              method: 'POST',
              headers: { 
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
              },
              body: JSON.stringify(processedData),
            });
            const result = await res.json();
            
            if (res.status === 401 || res.status === 403) {
              if (result.error === 'SESSION_EXPIRED_CONCURRENT') {
                setIsKickedOut(true);
              } else {
                handleLogout();
              }
              return;
            }
            
            if (result.success) {
              // Reload from backend to get IDs
              await loadDataFromBackend();
              setImportStatus({ show: true, message: `成功导入 ${result.inserted} 条数据并已持久化`, type: 'success' });
            } else {
              setImportStatus({ show: true, message: `导入成功但持久化失败: ${result.error}`, type: 'error' });
            }
          } catch {
            setImportStatus({ show: true, message: `导入成功（数据已显示），但后端保存失败`, type: 'error' });
          }
        } else {
          setImportStatus({ show: true, message: `成功导入 ${processedData.length} 条数据（仅前端，后端离线）`, type: 'success' });
        }
      } catch (error) {
        setImportStatus({ show: true, message: '导入失败，请检查文件格式', type: 'error' });
      }
      if (fileInputRef.current) fileInputRef.current.value = '';
      setTimeout(() => setImportStatus(prev => ({ ...prev, show: false })), 3000);
    };
    reader.readAsBinaryString(file);
  };

  const handleImportRepairExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const bstr = event.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const normalizeHeader = (value: any) => String(value || '')
          .toLowerCase()
          .replace(/[\s\r\n\t（）()\\/_-]/g, '')
          .replace(/产品个数/g, '产品数');
        const parseNumber = (...values: any[]) => {
          for (const value of values) {
            if (value === undefined || value === null || value === '') continue;
            const normalized = String(value).replace(/[%\s,]/g, '');
            const parsed = Number(normalized);
            if (!Number.isNaN(parsed)) return parsed;
          }
          return 0;
        };
        const parseRateNumber = (...values: any[]) => {
          for (const value of values) {
            if (value === undefined || value === null || value === '') continue;
            if (typeof value === 'number' && Number.isFinite(value)) {
              return value > 0 && value <= 1 ? value * 100 : value;
            }
            const text = String(value).trim();
            if (!text) continue;
            const normalized = text.replace(/[%\s,]/g, '');
            const parsed = Number(normalized);
            if (!Number.isNaN(parsed)) {
              return text.includes('%') ? parsed : (parsed > 0 && parsed <= 1 ? parsed * 100 : parsed);
            }
          }
          return 0;
        };
        const buildNormalizedRow = (row: Record<string, any>) => {
          const normalized: Record<string, any> = {};
          Object.entries(row).forEach(([key, value]) => {
            normalized[normalizeHeader(key)] = value;
          });
          return normalized;
        };
        const getValueByAliases = (row: Record<string, any>, aliases: string[]) => {
          for (const alias of aliases) {
            const normalizedAlias = normalizeHeader(alias);
            if (normalizedAlias in row) {
              return row[normalizedAlias];
            }
          }
          return undefined;
        };
        const requiredGroups = {
          period: ['日期', '年月', '月份', '日期单位年月'],
          productLine: ['产品分类', '产品线', '分类'],
          totalRepairRate: ['总返修率'],
        };
        const optionalMetricGroups = [
          ['月度返修产品数（去除OOB）', '月度返修产品数(去除OOB)', '月度返修产品数', '月度返修产品数去除OOB'],
          ['月度发货数'],
          ['截止当月保内发货总数', '保内发货总数'],
          ['月度OOB开箱不合格产品数', 'OOB开箱不合格产品数'],
          ['月度返修率'],
          ['月度OOB开箱不合格率'],
        ];
        const hasAlias = (headers: string[], aliases: string[]) => {
          const normalizedHeaders = headers.map(normalizeHeader);
          return aliases.some(alias => normalizedHeaders.includes(normalizeHeader(alias)));
        };
        const findRepairSheet = () => {
          for (const sheetName of wb.SheetNames) {
            const ws = wb.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }) as any[][];
            for (let rowIndex = 0; rowIndex < Math.min(rows.length, 20); rowIndex++) {
              const row = rows[rowIndex].map(cell => String(cell || ''));
              const requiredMatched = Object.values(requiredGroups).every(group => hasAlias(row, group));
              const optionalMatched = optionalMetricGroups.filter(group => hasAlias(row, group)).length;
              if (requiredMatched && optionalMatched >= 2) {
                return { sheetName, headerRowIndex: rowIndex };
              }
            }
          }
          return null;
        };
        const normalizeProductLine = (value: any) => {
          const text = String(value || '').trim();
          if (text.includes('所有产品') || text.toLowerCase().includes('all products')) return 'all';
          if (text.includes('机械臂')) return 'roboticArm';
          if (text.includes('机器人')) return 'robot';
          if (text.includes('关节')) return 'joint';
          if (text.includes('其他')) return 'others';
          return 'others';
        };
        const parseYearMonth = (value: any) => {
          if (typeof value === 'number') {
            const parsed = XLSX.SSF.parse_date_code(value);
            if (parsed) {
              return { year: parsed.y, month: parsed.m };
            }
          }
          const text = String(value || '').trim();
          const match = text.match(/(\d{4})\D+(\d{1,2})/);
          if (match) {
            return { year: Number(match[1]), month: Number(match[2]) };
          }
          const date = new Date(text);
          if (!Number.isNaN(date.getTime())) {
            return { year: date.getFullYear(), month: date.getMonth() + 1 };
          }
          return { year: 0, month: 0 };
        };
        const repairSheet = findRepairSheet();
        if (!repairSheet) {
          throw new Error('repair-sheet-not-found');
        }
        const ws = wb.Sheets[repairSheet.sheetName];
        const rawData: any[] = XLSX.utils.sheet_to_json(ws, {
          range: repairSheet.headerRowIndex,
          defval: '',
        });

        const processedData = rawData
          .map(rawRow => {
            const row = buildNormalizedRow(rawRow);
            const period = parseYearMonth(getValueByAliases(row, ['日期', '年月', '月份', '日期单位年月']));
            const repairCount = parseNumber(
              getValueByAliases(row, ['月度返修产品数（去除OOB）', '月度返修产品数(去除OOB)', '月度返修产品数', '月度返修产品数去除OOB'])
            );
            const monthShipmentCount = parseNumber(getValueByAliases(row, ['月度发货数']));
            const warrantyShipmentCount = parseNumber(getValueByAliases(row, ['截止当月保内发货总数', '保内发货总数']));
            const oobDefectCount = parseNumber(getValueByAliases(row, ['月度OOB开箱不合格产品数', 'OOB开箱不合格产品数']));
            const monthlyRepairRate = parseRateNumber(getValueByAliases(row, ['月度返修率']));
            const monthlyOobRate = parseRateNumber(getValueByAliases(row, ['月度OOB开箱不合格率', '月度OOB开箱不合格率']));
            const totalRepairRate = parseRateNumber(getValueByAliases(row, ['总返修率'])) || (monthlyRepairRate * 0.4 + monthlyOobRate * 0.6);
            const productLine = normalizeProductLine(getValueByAliases(row, ['产品分类', '产品线', '分类']));

            return {
              year: period.year,
              month: period.month,
              productLine,
              repairCount,
              monthShipmentCount,
              warrantyShipmentCount,
              oobDefectCount,
              monthlyRepairRate,
              monthlyOobRate,
              totalRepairRate,
            };
          })
          .filter(row => row.year > 0 && row.month > 0);

        if (processedData.length === 0) {
          throw new Error('repair-data-empty');
        }

        setRepairData(processedData);

        if (backendStatus === 'online') {
          try {
            await fetch('/api/repair-rates', {
              method: 'DELETE',
              headers: { 'Authorization': `Bearer ${token}` }
            });
            const res = await fetch('/api/repair-rates/bulk', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
              },
              body: JSON.stringify(processedData),
            });
            const result = await res.json();

            if (res.status === 401 || res.status === 403) {
              if (result.error === 'SESSION_EXPIRED_CONCURRENT') {
                setIsKickedOut(true);
              } else {
                handleLogout();
              }
              return;
            }

            if (result.success) {
              await loadRepairDataFromBackend();
              setImportStatus({ show: true, message: `${t('repairUploadSuccess')} (${result.inserted})`, type: 'success' });
            } else {
              setImportStatus({ show: true, message: result.error || '返修率表导入失败', type: 'error' });
            }
          } catch {
            setImportStatus({ show: true, message: '返修率表已显示，但后端保存失败', type: 'error' });
          }
        } else {
          setImportStatus({ show: true, message: `${t('repairUploadSuccess')}（仅前端）`, type: 'success' });
        }
      } catch (error) {
        setImportStatus({ show: true, message: '返修率表导入失败，请检查文件格式', type: 'error' });
      }
      if (repairFileInputRef.current) repairFileInputRef.current.value = '';
      setTimeout(() => setImportStatus(prev => ({ ...prev, show: false })), 3000);
    };
    reader.readAsBinaryString(file);
  };

  const handleReset = async () => {
    const isRepairTab = activeTab === 'repair';
    if (isRepairTab) {
      setRepairData([]);
    } else {
      setData([]);
    }
    if (backendStatus === 'online') {
      try {
        const res = await fetch(isRepairTab ? '/api/repair-rates' : '/api/issues', { 
          method: 'DELETE',
          headers: { 'Authorization': `Bearer ${token}` }
        });
        const result = await res.json();
        
        if (res.status === 401 || res.status === 403) {
          if (result.error === 'SESSION_EXPIRED_CONCURRENT') {
            setIsKickedOut(true);
          } else {
            handleLogout();
          }
          return;
        }
      } catch (err) {
        console.error('Failed to reset backend data:', err);
      }
    }
    setImportStatus({ show: true, message: isRepairTab ? t('clearRepairData') : t('resetData'), type: 'success' });
    setTimeout(() => setImportStatus(prev => ({ ...prev, show: false })), 3000);
  };

  const handleExport = async () => {
    // Build query params from current filters
    const params = new URLSearchParams();
    if (selectedYear !== 'all' && selectedMonth !== 'all') {
      params.set('year', String(selectedYear));
      params.set('month', String(selectedMonth));
    }
    if (selectedProductLine !== 'all') params.set('productLine', selectedProductLine);
    if (selectedCause !== 'all') params.set('cause', selectedCause);
    if (selectedDept !== 'all') params.set('dept', selectedDept);
    if (selectedOob !== 'all') params.set('oob', selectedOob);
    
    try {
      const res = await fetch(`/api/export?${params.toString()}`, {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      const result = await res.json().catch(() => ({})); 

      if (res.status === 401 || res.status === 403) {
        if (result.error === 'SESSION_EXPIRED_CONCURRENT') {
          setIsKickedOut(true);
        } else {
          handleLogout();
        }
        return;
      }
      
      // If it's a blob, we need to fetch again or handle response type correctly
      // For simplicity, if we get here and it's 200, assume it's data or we handle as blob
      const blobRes = await fetch(`/api/export?${params.toString()}`, {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      const blob = await blobRes.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `质量数据报表_${new Date().toISOString().slice(0, 10)}.xlsx`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Export failed:', err);
    }
  };

  if (!token) {
    return <Login onLogin={handleLogin} />;
  }

  return (
    <>
      <div className="flex bg-slate-50 min-h-screen text-slate-600 font-sans selection:bg-primary/10">
      {/* Notifications */}
      {importStatus.show && (
        <div className={cn(
          "fixed top-6 left-1/2 -translate-x-1/2 z-50 flex items-center gap-2 px-4 py-2 rounded-full shadow-lg border animate-in fade-in slide-in-from-top-4 duration-300",
          importStatus.type === 'success' ? "bg-emerald-50 border-emerald-100 text-emerald-700" : "bg-rose-50 border-rose-100 text-rose-700"
        )}>
          {importStatus.type === 'success' ? <CheckCircle2 size={16} /> : <AlertTriangle size={16} />}
          <span className="text-xs font-semibold">{importStatus.message}</span>
          <button onClick={() => setImportStatus(prev => ({ ...prev, show: false }))} className="ml-2 hover:opacity-70">
            <X size={14} />
          </button>
        </div>
      )}

      {/* Sidebar */}
      <aside className="w-64 border-r border-slate-200 bg-white/50 backdrop-blur-xl flex flex-col">
        <div className="p-8">
          <div className="flex items-center gap-3">
            <div className="size-8 bg-primary rounded-lg flex items-center justify-center text-white">
              <BarChart3 size={20} />
            </div>
            <span className="text-xl font-semibold tracking-tight">{t('brand')}</span>
          </div>
        </div>
        <nav className="flex-1 px-4 space-y-1">
          <SidebarItem icon={LayoutDashboard} label={t('overview')} active={activeTab === 'overview'} onClick={() => setActiveTab('overview')} />
          <SidebarItem icon={Factory} label={t('repairDashboard')} active={activeTab === 'repair'} onClick={() => setActiveTab('repair')} />
          <SidebarItem icon={BellRing} label={t('oobDashboard')} active={activeTab === 'oob'} onClick={() => setActiveTab('oob')} />
          <SidebarItem icon={CheckCircle2} label={t('closeDashboard')} active={activeTab === 'close'} onClick={() => setActiveTab('close')} />
          <SidebarItem icon={Layers} label={t('rootCauseDashboard')} active={activeTab === 'rootCause'} onClick={() => setActiveTab('rootCause')} />
          <SidebarItem icon={Settings} label={t('data')} active={activeTab === 'data'} onClick={() => setActiveTab('data')} />
        </nav>
        
        <div className="p-4 border-t border-slate-100 flex flex-col gap-3">
          <div className="flex items-center gap-2 px-4 py-2 bg-slate-50 rounded-lg">
            <div className={cn(
              "size-2 rounded-full",
              backendStatus === 'online' ? "bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.5)]" : 
              backendStatus === 'checking' ? "bg-amber-400 animate-pulse" : "bg-rose-500"
            )} />
            <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
              Backend: {backendStatus}
            </span>
          </div>
          
          <div className="bg-slate-50 rounded-lg p-3 border border-slate-100 mt-2">
            <div className="flex items-center gap-3">
              <div className="size-8 rounded-full bg-primary/10 flex items-center justify-center text-primary font-bold uppercase">
                {username ? username[0] : 'U'}
              </div>
              <div className="flex-1">
                <p className="text-xs font-bold text-slate-900 leading-none">{username || 'User'}</p>
                <div className="mt-1 inline-flex items-center px-1.5 py-0.5 rounded text-[8px] font-bold bg-primary/10 text-primary">
                  ADMIN
                </div>
              </div>
            </div>
            <button onClick={handleLogout} className="w-full mt-3 flex justify-center items-center gap-2 py-1.5 rounded-md text-xs font-medium text-slate-500 hover:text-rose-600 hover:bg-rose-50 border border-transparent hover:border-rose-100 transition-all">
              <LogOut size={14} /> 退出登录
            </button>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto p-8 lg:p-12">
        {/* Header */}
        <header className="flex justify-between items-end mb-10">
          <div>
            <h1 className="text-2xl font-bold tracking-tight text-slate-900">
              {activeTab === 'overview' && t('title')}
              {activeTab === 'repair' && t('repairDashboard')}
              {activeTab === 'oob' && t('oobDashboard')}
              {activeTab === 'close' && t('closeDashboard')}
              {activeTab === 'rootCause' && t('rootCauseDashboard')}
              {activeTab === 'data' && t('data')}
            </h1>
            <p className="text-xs text-slate-500 mt-1 uppercase tracking-widest font-medium">{t('subtitle')}</p>
          </div>
          <div className="flex items-center gap-2">
            {activeTab !== 'repair' && (
              <>
                <input type="file" ref={fileInputRef} onChange={handleImportExcel} accept=".xlsx, .xls" className="hidden" />
                <button onClick={() => fileInputRef.current?.click()} className="flex items-center gap-1.5 px-3 py-1.5 bg-primary rounded-md text-xs font-medium text-white hover:bg-primary/90 transition-colors shadow-sm">
                  <FileSpreadsheet size={14} /> {t('importExcel')}
                </button>
              </>
            )}
            {activeTab === 'repair' && (
              <>
                <input type="file" ref={repairFileInputRef} onChange={handleImportRepairExcel} accept=".xlsx, .xls" className="hidden" />
                <button onClick={() => repairFileInputRef.current?.click()} className="flex items-center gap-1.5 px-3 py-1.5 bg-primary rounded-md text-xs font-medium text-white hover:bg-primary/90 transition-colors shadow-sm">
                  <FileSpreadsheet size={14} /> {t('importRepairExcel')}
                </button>
              </>
            )}
            <button onClick={handleReset} className="flex items-center gap-1.5 px-3 py-1.5 bg-amber-500 rounded-md text-xs font-medium text-white hover:bg-amber-600 transition-colors shadow-sm">
              <RefreshCw size={14} /> {activeTab === 'repair' ? t('clearRepairData') : t('resetData')}
            </button>
            {activeTab !== 'repair' && (
              <button onClick={handleExport} className="flex items-center gap-1.5 px-3 py-1.5 bg-slate-100 rounded-md border border-slate-200 text-xs font-medium text-slate-600 hover:bg-slate-200 transition-colors">
                <Download size={14} /> {t('exportReport')}
              </button>
            )}
            <div className="flex items-center ml-4 bg-white border border-slate-200 rounded-md p-0.5 text-[10px] font-medium shadow-sm overflow-hidden">
              <button onClick={() => setLang('EN')} className={cn("px-2 py-1 transition-colors", lang === 'EN' ? "bg-slate-100 text-primary" : "text-slate-400")}>EN</button>
              <span className="text-slate-200">|</span>
              <button onClick={() => setLang('中')} className={cn("px-2 py-1 transition-colors", lang === '中' ? "bg-slate-100 text-primary" : "text-slate-400")}>中</button>
            </div>
          </div>
        </header>

        {activeTab === 'overview' && (
          <>
            <div className="grid grid-cols-1 md:grid-cols-5 gap-4 mb-6">
              <KPICard title={t('monthlyIssues')} value={kpiStats.issues.val} trend={kpiStats.issues.trend.direction} trendValue={`${t('vsLastMonth')} ${kpiStats.issues.trend.value}`} />
              <KPICard title={t('oobRate')} value={kpiStats.oob.val} trend={kpiStats.oob.trend.direction} trendValue={`${t('vsLastMonth')} ${kpiStats.oob.trend.value}`} />
              <KPICard title={t('repairRate')} value={kpiStats.repair.val} trend={kpiStats.repair.trend.direction} trendValue={`${t('vsLastMonth')} ${kpiStats.repair.trend.value}`} />
              <KPICard title={t('closeRate')} value={kpiStats.close.val} trend={kpiStats.close.trend.direction} trendValue={`${t('vsLastMonth')} ${kpiStats.close.trend.value}`} />
              <KPICard title={t('overdue')} value={kpiStats.overdue.val} trend={kpiStats.overdue.trend.direction} trendValue={`${t('vsLastMonth')} ${kpiStats.overdue.trend.value}`} color="rose" />
            </div>

            <div className="apple-card p-4 mb-6">
              <div className="flex items-center gap-2 mb-3 text-primary text-xs font-semibold"><Search size={14} /> {t('filter')}</div>
              <div className="grid grid-cols-6 gap-4">
                <div>
                  <label className="block text-[10px] text-slate-500 mb-1 ml-1">{t('year')}</label>
                  <div className="relative">
                    <select 
                      value={selectedYear}
                      onChange={(e) => {
                        if (e.target.value === 'all') {
                          setSelectedYear('all');
                          setSelectedMonth('all');
                        } else {
                          setSelectedYear(parseInt(e.target.value, 10));
                          setSelectedMonth('all');
                        }
                      }}
                      className="w-full appearance-none bg-slate-50 border border-slate-200 rounded-md text-xs h-8 pl-2 pr-8 focus:outline-none focus:ring-1 focus:ring-primary/20"
                    >
                      <option value="all">{t('selectYear')}</option>
                      {yearOptions.map((y) => (
                        <option key={y} value={y}>{y}</option>
                      ))}
                    </select>
                    <ChevronDown size={12} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                  </div>
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500 mb-1 ml-1">{t('month')}</label>
                  <div className="relative">
                    <select 
                      value={selectedMonth}
                      onChange={(e) => setSelectedMonth(e.target.value === 'all' ? 'all' : parseInt(e.target.value, 10))}
                      disabled={selectedYear === 'all'}
                      className={cn(
                        "w-full appearance-none bg-slate-50 border border-slate-200 rounded-md text-xs h-8 pl-2 pr-8 focus:outline-none focus:ring-1 focus:ring-primary/20",
                        selectedYear === 'all' && "opacity-60 cursor-not-allowed"
                      )}
                    >
                      <option value="all">{t('selectMonth')}</option>
                      <option value="1">{t('jan')}</option>
                      <option value="2">{t('feb')}</option>
                      <option value="3">{t('mar')}</option>
                      <option value="4">{t('apr')}</option>
                      <option value="5">{t('may')}</option>
                      <option value="6">{t('jun')}</option>
                      <option value="7">{t('jul')}</option>
                      <option value="8">{t('aug')}</option>
                      <option value="9">{t('sep')}</option>
                      <option value="10">{t('oct')}</option>
                      <option value="11">{t('nov')}</option>
                      <option value="12">{t('dec')}</option>
                    </select>
                    <ChevronDown size={12} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                  </div>
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500 mb-1 ml-1">{t('productLine')}</label>
                  <div className="relative">
                    <select 
                      value={selectedProductLine}
                      onChange={(e) => setSelectedProductLine(e.target.value)}
                      className="w-full appearance-none bg-slate-50 border border-slate-200 rounded-md text-xs h-8 pl-2 pr-8 focus:outline-none focus:ring-1 focus:ring-primary/20"
                    >
                      <option value="all">{t('all')}</option>
                      <option value="roboticArm">{t('roboticArm')}</option>
                      <option value="robot">{t('robot')}</option>
                      <option value="joint">{t('joint')}</option>
                      <option value="others">{t('others')}</option>
                    </select>
                    <ChevronDown size={12} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                  </div>
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500 mb-1 ml-1">{t('rootCause')}</label>
                  <div className="relative">
                    <select 
                      value={selectedCause}
                      onChange={(e) => setSelectedCause(e.target.value)}
                      className="w-full appearance-none bg-slate-50 border border-slate-200 rounded-md text-xs h-8 pl-2 pr-8 focus:outline-none focus:ring-1 focus:ring-primary/20"
                    >
                      <option value="all">{t('all')}</option>
                      <option value="pkgTrans">{t('pkgTrans')}</option>
                      <option value="material">{t('material')}</option>
                      <option value="field">{t('field')}</option>
                      <option value="assembly">{t('assembly')}</option>
                      <option value="maint">{t('maint')}</option>
                      <option value="design">{t('design')}</option>
                      <option value="req">{t('req')}</option>
                    </select>
                    <ChevronDown size={12} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                  </div>
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500 mb-1 ml-1">{t('dept')}</label>
                  <div className="relative">
                    <select 
                      value={selectedDept}
                      onChange={(e) => setSelectedDept(e.target.value)}
                      className="w-full appearance-none bg-slate-50 border border-slate-200 rounded-md text-xs h-8 pl-2 pr-8 focus:outline-none focus:ring-1 focus:ring-primary/20"
                    >
                      <option value="all">{t('all')}</option>
                      <option value="deptRuiJu">{t('deptRuiJu')}</option>
                      <option value="deptWeiHan">{t('deptWeiHan')}</option>
                      <option value="deptRuiYou">{t('deptRuiYou')}</option>
                      <option value="deptAdv">{t('deptAdv')}</option>
                    </select>
                    <ChevronDown size={12} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                  </div>
                </div>
                <div>
                  <label className="block text-[10px] text-slate-500 mb-1 ml-1">{t('oobStatus')}</label>
                  <div className="relative">
                    <select 
                      value={selectedOob}
                      onChange={(e) => setSelectedOob(e.target.value)}
                      className="w-full appearance-none bg-slate-50 border border-slate-200 rounded-md text-xs h-8 pl-2 pr-8 focus:outline-none focus:ring-1 focus:ring-primary/20"
                    >
                      <option value="all">{t('all')}</option>
                      <option value="oobDamage">{t('oobDamage')}</option>
                      <option value="nonOobDamage">{t('nonOobDamage')}</option>
                    </select>
                    <ChevronDown size={12} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                  </div>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <div className="apple-card p-5">
                <h3 className="text-sm font-semibold text-slate-900 flex items-center gap-2 mb-6"><BarChart3 size={16} className="text-amber-500" /> {t('ranking')}</h3>
                <div className="space-y-5">
                  {modelRanking.map((item, idx) => (
                    <div key={idx} className="space-y-1.5">
                      <div className="flex justify-between text-[10px] text-slate-500 font-medium"><span>{item.name}</span><span>{t('issueCount')}: {item.count}</span></div>
                      <div className="w-full bg-slate-100 rounded-full h-3 overflow-hidden">
                        <div className="bg-slate-700 h-full rounded-full transition-all duration-1000" style={{ width: `${Math.min(item.count * 10, 100)}%` }} />
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="apple-card p-5">
                <h3 className="text-sm font-semibold text-slate-900 flex items-center gap-2 mb-6"><PieChart size={16} className="text-primary" /> {t('distribution')}</h3>
                <div className="h-48 w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <RePieChart>
                      <Pie data={causeDistribution} cx="50%" cy="50%" innerRadius={40} outerRadius={60} paddingAngle={5} dataKey="value">
                        {causeDistribution.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.color} />)}
                      </Pie>
                      <Tooltip contentStyle={{ borderRadius: '8px', border: 'none', fontSize: '10px' }} />
                    </RePieChart>
                  </ResponsiveContainer>
                </div>
                <div className="grid grid-cols-2 gap-2 mt-4">
                  {causeDistribution.slice(0, 4).map((item, idx) => (
                    <div key={idx} className="flex items-center gap-2 text-[9px]">
                      <span className="size-1.5 rounded-full" style={{ backgroundColor: item.color }} />
                      <span className="text-slate-600 truncate">{item.name} ({item.value})</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </>
        )}

        {activeTab === 'repair' && (
          <div className="space-y-6">
            <div className="apple-card p-5 border border-primary/10 bg-primary/5">
              <div className="flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
                <div>
                  <h3 className="text-sm font-semibold text-slate-900">{t('importRepairExcel')}</h3>
                  <p className="text-xs text-slate-500 mt-1">{t('repairImportHint')}</p>
                </div>
                <button onClick={() => repairFileInputRef.current?.click()} className="inline-flex items-center gap-1.5 px-3 py-2 bg-primary rounded-md text-xs font-medium text-white hover:bg-primary/90 transition-colors shadow-sm">
                  <FileSpreadsheet size={14} /> {t('importRepairExcel')}
                </button>
              </div>
            </div>
            <div className="apple-card p-6">
              <h3 className="text-sm font-semibold text-slate-900 mb-6 flex items-center gap-2">
                <PieChart size={16} className="text-primary" /> {t('repairDashboardTitle')}
              </h3>
              <div className="h-96 w-full">
                {repairRateTrendByMonth.length > 0 ? (
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={repairRateTrendByMonth}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                      <YAxis
                        axisLine={false}
                        tickLine={false}
                        tick={{ fontSize: 10, fill: '#64748b' }}
                        tickFormatter={(v) => `${Number(v).toFixed(2)}%`}
                      />
                      <Tooltip formatter={(v: any) => `${Number(v).toFixed(3)}%`} />
                      <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                      <Line type="monotone" dataKey="all" stroke="#0f172a" strokeWidth={2.5} dot={false} name={t('allProducts')} />
                      <Line type="monotone" dataKey="others" stroke="#f97316" strokeWidth={2} dot={false} name={t('others')} />
                      <Line type="monotone" dataKey="robot" stroke="#2563eb" strokeWidth={2} dot={false} name={t('robot')} />
                      <Line type="monotone" dataKey="roboticArm" stroke="#a855f7" strokeWidth={2} dot={false} name={t('roboticArm')} />
                      <Line type="monotone" dataKey="joint" stroke="#0ea5e9" strokeWidth={2} dot={false} name={t('joint')} />
                    </LineChart>
                  </ResponsiveContainer>
                ) : (
                  <div className="h-full flex items-center justify-center rounded-2xl border border-dashed border-slate-200 bg-slate-50 text-sm text-slate-400">
                    {t('repairImportEmpty')}
                  </div>
                )}
              </div>
            </div>

          </div>
        )}

        {activeTab === 'oob' && (
          <div className="space-y-4">
            <div>
              <h3 className="text-lg font-semibold text-slate-900">{t('oobModuleTitle')}</h3>
              <p className="text-xs text-slate-500 mt-1">{t('oobModuleHint')}</p>
            </div>

            {oobDashboard.hasData ? (
              <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('oobRateBoardTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={oobDashboard.rateTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} tickFormatter={(v) => `${Number(v).toFixed(2)}%`} />
                        <Tooltip formatter={(v: any) => `${Number(v).toFixed(3)}%`} />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        <Line type="monotone" dataKey="all" stroke="#0f172a" strokeWidth={2.5} dot={false} name={t('allProducts')} />
                        <Line type="monotone" dataKey="others" stroke="#f97316" strokeWidth={2} dot={false} name={t('others')} />
                        <Line type="monotone" dataKey="robot" stroke="#2563eb" strokeWidth={2} dot={false} name={t('robot')} />
                        <Line type="monotone" dataKey="roboticArm" stroke="#a855f7" strokeWidth={2} dot={false} name={t('roboticArm')} />
                        <Line type="monotone" dataKey="joint" stroke="#0ea5e9" strokeWidth={2} dot={false} name={t('joint')} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('oobCountBoardTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={oobDashboard.countTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <Tooltip />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        <Line type="monotone" dataKey="all" stroke="#0f172a" strokeWidth={2.5} dot={false} name={t('allProducts')} />
                        <Line type="monotone" dataKey="others" stroke="#f97316" strokeWidth={2} dot={false} name={t('others')} />
                        <Line type="monotone" dataKey="robot" stroke="#2563eb" strokeWidth={2} dot={false} name={t('robot')} />
                        <Line type="monotone" dataKey="roboticArm" stroke="#a855f7" strokeWidth={2} dot={false} name={t('roboticArm')} />
                        <Line type="monotone" dataKey="joint" stroke="#0ea5e9" strokeWidth={2} dot={false} name={t('joint')} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('mainDeptOobTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={oobDashboard.mainDeptTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <Tooltip />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        {oobDashboard.mainDepartments.map((dept, idx) => (
                          <Line key={dept} type="monotone" dataKey={dept} stroke={['#0f172a', '#2563eb', '#f97316', '#a855f7', '#0ea5e9', '#16a34a'][idx % 6]} strokeWidth={2} dot={false} name={dept} />
                        ))}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('initialDeptOobTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={oobDashboard.initialDeptTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <Tooltip />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        {oobDashboard.initialDepartments.map((dept, idx) => (
                          <Line key={dept} type="monotone" dataKey={dept} stroke={['#0f172a', '#2563eb', '#f97316', '#a855f7', '#0ea5e9', '#16a34a'][idx % 6]} strokeWidth={2} dot={false} name={dept} />
                        ))}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            ) : (
              <div className="apple-card p-6 text-sm text-slate-400">{t('oobDashboardEmpty')}</div>
            )}
          </div>
        )}

        {activeTab === 'close' && (
          <div className="space-y-4">
            <div>
              <h3 className="text-lg font-semibold text-slate-900">{t('closeModuleTitle')}</h3>
              <p className="text-xs text-slate-500 mt-1">{t('closeModuleHint')}</p>
            </div>

            {closeRateDashboard.hasData ? (
              <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('overallCloseRateTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={closeRateDashboard.overallTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} tickFormatter={(v) => `${Number(v).toFixed(0)}%`} />
                        <Tooltip formatter={(v: any) => `${Number(v).toFixed(2)}%`} />
                        <Line type="monotone" dataKey="closeRate" stroke="#2563eb" strokeWidth={2.5} dot={{ r: 4 }} name={t('closeRate')} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('productCloseRateTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={closeRateDashboard.productTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} tickFormatter={(v) => `${Number(v).toFixed(0)}%`} />
                        <Tooltip formatter={(v: any) => `${Number(v).toFixed(2)}%`} />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        <Line type="monotone" dataKey="roboticArm" stroke="#2563eb" strokeWidth={2} dot={false} name={t('roboticArm')} />
                        <Line type="monotone" dataKey="others" stroke="#f97316" strokeWidth={2} dot={false} name={t('others')} />
                        <Line type="monotone" dataKey="robot" stroke="#94a3b8" strokeWidth={2} dot={false} name={t('robot')} />
                        <Line type="monotone" dataKey="joint" stroke="#eab308" strokeWidth={2} dot={false} name={t('joint')} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('productOobCloseRateTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={closeRateDashboard.oobProductTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} tickFormatter={(v) => `${Number(v).toFixed(0)}%`} />
                        <Tooltip formatter={(v: any) => `${Number(v).toFixed(2)}%`} />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        <Line type="monotone" dataKey="roboticArm" stroke="#2563eb" strokeWidth={2} dot={false} name={t('roboticArm')} />
                        <Line type="monotone" dataKey="others" stroke="#f97316" strokeWidth={2} dot={false} name={t('others')} />
                        <Line type="monotone" dataKey="robot" stroke="#94a3b8" strokeWidth={2} dot={false} name={t('robot')} />
                        <Line type="monotone" dataKey="joint" stroke="#eab308" strokeWidth={2} dot={false} name={t('joint')} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('mainDeptCloseRateTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={closeRateDashboard.mainDeptTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} tickFormatter={(v) => `${Number(v).toFixed(0)}%`} />
                        <Tooltip formatter={(v: any) => `${Number(v).toFixed(2)}%`} />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        {closeRateDashboard.mainDepartments.map((dept, idx) => (
                          <Line key={dept} type="monotone" dataKey={dept} stroke={['#2563eb', '#f97316', '#94a3b8', '#eab308', '#0ea5e9', '#16a34a'][idx % 6]} strokeWidth={2} dot={false} name={dept} />
                        ))}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="apple-card p-6 xl:col-span-2">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('initialDeptCloseRateTitle')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={closeRateDashboard.initialDeptTrend}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} tickFormatter={(v) => `${Number(v).toFixed(0)}%`} />
                        <Tooltip formatter={(v: any) => `${Number(v).toFixed(2)}%`} />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        {closeRateDashboard.initialDepartments.map((dept, idx) => (
                          <Line key={dept} type="monotone" dataKey={dept} stroke={['#2563eb', '#f97316', '#94a3b8', '#eab308', '#0ea5e9', '#16a34a'][idx % 6]} strokeWidth={2} dot={false} name={dept} />
                        ))}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </div>
            ) : (
              <div className="apple-card p-6 text-sm text-slate-400">{t('closeDashboardEmpty')}</div>
            )}
          </div>
        )}

        {activeTab === 'rootCause' && (
          <div className="space-y-6">
            <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
              <div>
                <h3 className="text-lg font-semibold text-slate-900">{t('rootCauseModuleTitle')}</h3>
                <p className="text-xs text-slate-500 mt-1">{t('rootCauseModuleHint')}</p>
              </div>

              <div className="w-full max-w-xs">
                <label className="block text-[10px] text-slate-500 mb-1 ml-1">{t('rootCausePeriodLabel')}</label>
                <div className="relative">
                  <select
                    value={selectedRootCausePeriod}
                    onChange={(e) => setSelectedRootCausePeriod(e.target.value)}
                    disabled={!rootCausePeriodOptions.length}
                    className={cn(
                      "w-full appearance-none bg-slate-50 border border-slate-200 rounded-md text-xs h-9 pl-3 pr-8 focus:outline-none focus:ring-1 focus:ring-primary/20",
                      !rootCausePeriodOptions.length && "opacity-60 cursor-not-allowed"
                    )}
                  >
                    {rootCausePeriodOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                  <ChevronDown size={12} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                </div>
              </div>
            </div>

            {rootCauseDashboard.hasData ? (
              <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('rootCauseBoard1')}</h4>
                  <div className="h-80 w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={rootCauseDashboard.trendData}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="month" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                        <Tooltip />
                        <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                        {rootCauseDashboard.trendCauses.map((cause, idx) => (
                          <Line
                            key={cause}
                            type="monotone"
                            dataKey={(entry: any) => entry[cause]}
                            stroke={rootCauseDashboard.colorPalette[idx % rootCauseDashboard.colorPalette.length]}
                            strokeWidth={2}
                            dot={false}
                            name={cause}
                          />
                        ))}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="apple-card p-6">
                  <div className="flex items-center justify-between gap-4 mb-4">
                    <h4 className="text-sm font-semibold text-slate-900">{t('rootCauseBoard2')}</h4>
                    <span className="text-[11px] text-slate-400">{rootCauseDashboard.selectedPeriodLabel || '-'}</span>
                  </div>
                  <div className="h-80 w-full">
                    {rootCauseDashboard.monthlyDistribution.length > 0 ? (
                      <ResponsiveContainer width="100%" height="100%">
                        <ComposedChart data={rootCauseDashboard.monthlyDistribution}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} angle={-18} textAnchor="end" height={70} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} tickFormatter={(v) => `${Number(v).toFixed(0)}%`} domain={[0, 100]} />
                          <Tooltip formatter={(v: any) => `${Number(v).toFixed(2)}%`} />
                          <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                          <Bar dataKey="share" fill="#2563eb" name={t('ratio')}>
                            <LabelList dataKey="share" position="top" formatter={(value: any) => `${Number(value).toFixed(1)}%`} className="fill-slate-500 text-[10px]" />
                          </Bar>
                          <Line type="monotone" dataKey="cumulativeShare" stroke="#f97316" strokeWidth={2.5} dot={{ r: 3 }} name={t('rootCauseCumulativeLabel')}>
                            <LabelList dataKey="cumulativeShare" position="top" formatter={(value: any) => `${Number(value).toFixed(1)}%`} className="fill-slate-500 text-[10px]" />
                          </Line>
                        </ComposedChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="h-full flex items-center justify-center rounded-2xl border border-dashed border-slate-200 bg-slate-50 text-sm text-slate-400">
                        {t('rootCauseModuleEmpty')}
                      </div>
                    )}
                  </div>
                </div>

                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('rootCauseBoard3')}</h4>
                  <div className="h-80 w-full">
                    {rootCauseDashboard.blankAnalysisDistribution.length > 0 ? (
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={rootCauseDashboard.blankAnalysisDistribution} layout="vertical" margin={{ left: 16, right: 24 }}>
                          <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                          <XAxis type="number" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                          <YAxis type="category" dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} width={110} />
                          <Tooltip />
                          <Bar dataKey="count" fill="#0ea5e9" radius={[0, 6, 6, 0]} name={t('count')}>
                            <LabelList dataKey="count" position="right" className="fill-slate-500 text-[10px]" />
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="h-full flex items-center justify-center rounded-2xl border border-dashed border-slate-200 bg-slate-50 text-sm text-slate-400">
                        {t('rootCauseBlankLabel')} {t('count')} = 0
                      </div>
                    )}
                  </div>
                </div>

                <div className="apple-card p-6">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('rootCauseBoard4')}</h4>
                  <div className="h-80 w-full">
                    {rootCauseDashboard.chartModels.length > 0 ? (
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={rootCauseDashboard.causeModelChartData}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis dataKey="cause" axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} angle={-12} textAnchor="end" height={56} />
                          <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fill: '#64748b' }} />
                          <Tooltip />
                          <Legend wrapperStyle={{ fontSize: 10, color: '#64748b' }} />
                          {rootCauseDashboard.chartModels.map((model, idx) => (
                            <Bar
                              key={model}
                              dataKey={(entry: any) => entry[model]}
                              fill={rootCauseDashboard.colorPalette[idx % rootCauseDashboard.colorPalette.length]}
                              radius={[4, 4, 0, 0]}
                              name={model}
                            />
                          ))}
                        </BarChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="h-full flex items-center justify-center rounded-2xl border border-dashed border-slate-200 bg-slate-50 text-sm text-slate-400">
                        {t('rootCauseModuleEmpty')}
                      </div>
                    )}
                  </div>
                </div>

                <div className="apple-card p-6 xl:col-span-2">
                  <h4 className="text-sm font-semibold text-slate-900 mb-4">{t('rootCauseTableTitle')}</h4>
                  <div className="overflow-x-auto">
                    <table className="w-full min-w-[960px] text-left border-collapse">
                      <thead>
                        <tr className="bg-slate-50 text-[10px] tracking-wider text-slate-500 font-bold whitespace-nowrap">
                          <th className="px-4 py-3 border-b border-slate-100">{t('rootCause')}</th>
                          {rootCauseDashboard.pivotModels.map((model) => (
                            <th key={model} className="px-4 py-3 border-b border-slate-100 text-right">{model}</th>
                          ))}
                          <th className="px-4 py-3 border-b border-slate-100 text-right">{rootCauseDashboard.totalLabel}</th>
                        </tr>
                      </thead>
                      <tbody className="text-xs text-slate-600">
                        {rootCauseDashboard.pivotRows.map((row) => (
                          <tr key={row.cause} className={row.cause === rootCauseDashboard.totalLabel ? 'bg-slate-50 font-semibold text-slate-900' : 'hover:bg-slate-50 transition-colors'}>
                            <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{row.cause}</td>
                            {rootCauseDashboard.pivotModels.map((model) => (
                              <td key={`${row.cause}-${model}`} className="px-4 py-3 border-b border-slate-50 text-right whitespace-nowrap">{row[model] || 0}</td>
                            ))}
                            <td className="px-4 py-3 border-b border-slate-50 text-right whitespace-nowrap">{row[rootCauseDashboard.totalLabel] || 0}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            ) : (
              <div className="apple-card p-6 text-sm text-slate-400">{t('rootCauseModuleEmpty')}</div>
            )}
          </div>
        )}

        {activeTab === 'data' && (
          <div className="apple-card overflow-hidden">
            <div className="p-4 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
              <h3 className="text-sm font-semibold text-slate-900">{t('rawJanData')}</h3>
              <div className="flex gap-2">
                <div className="relative">
                  <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                  <input type="text" placeholder={t('searchPlaceholder')} className="pl-8 pr-4 py-1.5 bg-white border border-slate-200 rounded-md text-xs focus:outline-none focus:ring-1 focus:ring-primary/20 w-64" />
                </div>
              </div>
            </div>
            <div ref={dataTableTopScrollRef} className="h-4 overflow-x-auto overflow-y-hidden bg-white">
              <div style={{ width: dataTableScrollWidth ? `${dataTableScrollWidth}px` : '100%' }} className="h-1" />
            </div>
            <div ref={dataTableScrollRef} className="overflow-x-auto">
              <table className="w-full min-w-[1400px] text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-[10px] tracking-wider text-slate-500 font-bold whitespace-nowrap">
                    <th className="px-4 py-3 border-b border-slate-100">标题（客户名称）</th>
                    <th className="px-4 py-3 border-b border-slate-100">产品数量</th>
                    <th className="px-4 py-3 border-b border-slate-100">产品类型</th>
                    <th className="px-4 py-3 border-b border-slate-100">产品型号</th>
                    <th className="px-4 py-3 border-b border-slate-100">问题初判主责部门</th>
                    <th className="px-4 py-3 border-b border-slate-100">问题主责部门</th>
                    <th className="px-4 py-3 border-b border-slate-100">创建时间</th>
                    <th className="px-4 py-3 border-b border-slate-100">SN日期（发货日期）</th>
                    <th className="px-4 py-3 border-b border-slate-100">是否开箱损</th>
                    <th className="px-4 py-3 border-b border-slate-100">根因分类</th>
                    <th className="px-4 py-3 border-b border-slate-100">问题分析类型</th>
                    <th className="px-4 py-3 border-b border-slate-100 text-right">问题数量</th>
                    <th className="px-4 py-3 border-b border-slate-100 text-right">问题关闭数量</th>
                  </tr>
                </thead>
                <tbody className="text-xs text-slate-600">
                  {filteredData.map((row, idx) => {
                    const productModelPath = String((row as any).productModelPath || '').trim();
                    const segments = productModelPath
                      .split('/')
                      .map((s: string) => s.trim())
                      .filter(Boolean);
                    const productType = segments[0] || '';
                    const productModel = segments.length ? segments[segments.length - 1] : (row.model || '');
                    const productTypeFallback = (() => {
                      if (row.productLine === 'robot') return '机器人';
                      if (row.productLine === 'roboticArm') return '机械臂';
                      if (row.productLine === 'joint') return '关节';
                      return '其他配件';
                    })();

                    const formatLongDateText = (value: any) => {
                      if (!value) return '-';
                      if (value instanceof Date) return `${value.getFullYear()}年${value.getMonth() + 1}月${value.getDate()}日`;
                      if (typeof value === 'number') {
                        const parsed = XLSX.SSF.parse_date_code(value);
                        if (parsed) return `${parsed.y}年${parsed.m}月${parsed.d}日`;
                        return String(value);
                      }
                      const text = String(value).trim();
                      if (/^\d+(\.\d+)?$/.test(text)) {
                        const n = Number(text);
                        const parsed = XLSX.SSF.parse_date_code(n);
                        if (parsed) return `${parsed.y}年${parsed.m}月${parsed.d}日`;
                        return text;
                      }
                      const m = text.match(/(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
                      if (m) return `${Number(m[1])}年${Number(m[2])}月${Number(m[3])}日`;
                      return text;
                    };

                    return (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        <td className="px-4 py-3 border-b border-slate-50 font-medium text-slate-900 whitespace-nowrap">{row.customerName || '-'}</td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{typeof (row as any).productQuantity === 'number' ? (row as any).productQuantity : '-'}</td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{productType || productTypeFallback}</td>
                        <td className="px-4 py-3 border-b border-slate-50 font-medium text-slate-900 whitespace-nowrap">{productModel || '-'}</td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{(row as any).initialDept || '-'}</td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{(row as any).mainDept || row.dept || '-'}</td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{formatLongDateText((row as any).complaintDate || row.createdAt)}</td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{formatLongDateText((row as any).snDate)}</td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">
                          <span className={cn("px-2 py-0.5 rounded-full text-[9px] font-bold", row.oob ? "bg-rose-50 text-rose-600" : "bg-slate-50 text-slate-400")}>
                            {row.oob ? '开箱损问题' : '非开箱损问题'}
                          </span>
                        </td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{row.cause || '-'}</td>
                        <td className="px-4 py-3 border-b border-slate-50 whitespace-nowrap">{(row as any).analysisType || '-'}</td>
                        <td className="px-4 py-3 border-b border-slate-50 text-right whitespace-nowrap">{row.issueQuantity ?? '-'}</td>
                        <td className="px-4 py-3 border-b border-slate-50 text-right whitespace-nowrap">{(row as any).closedQuantity ?? (row.closed ? row.issueQuantity : 0)}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </main>
    </div>

      {isKickedOut && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-300">
          <div className="bg-white rounded-2xl p-8 max-w-sm w-full mx-4 shadow-2xl border border-slate-100 text-center animate-in zoom-in-95 duration-300">
            <div className="w-16 h-16 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-6">
              <ShieldCheck size={32} />
            </div>
            <h2 className="text-xl font-bold text-slate-900 mb-2">安全提示</h2>
            <p className="text-slate-500 mb-8 leading-relaxed">该账号已被他人登录！如果不是你本人操作，请及时修改密码。</p>
            <button 
              onClick={handleLogout}
              className="w-full py-3 px-4 bg-slate-900 text-white rounded-xl font-semibold hover:bg-slate-800 active:scale-[0.98] transition-all flex items-center justify-center gap-2 group"
            >
              <span>确定</span>
              <ArrowRight size={18} className="group-hover:translate-x-1 transition-transform" />
            </button>
          </div>
        </div>
      )}
    </>
  );
}
