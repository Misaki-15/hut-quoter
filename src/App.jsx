
import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function isWeekend(d) { const day = new Date(d).getDay(); return day === 0 || day === 6; }
function skipWE(d) { const r = new Date(d); while (isWeekend(r)) r.setDate(r.getDate() + 1); return r; }
function addBusinessDays(start, days) {
  let d = skipWE(new Date(start));
  for (let i = 0; i < Math.max(0, Number(days) || 0); i += 1) {
    d = addDays(d, 1);
    while (isWeekend(d)) d = addDays(d, 1);
  }
  return d;
}
function fmt(d) { if (!d || Number.isNaN(new Date(d).getTime())) return "—"; const x = new Date(d); return `${x.getFullYear()}/${String(x.getMonth() + 1).padStart(2, "0")}/${String(x.getDate()).padStart(2, "0")}`; }
function fmtRMB(n) { return `¥${Number(n || 0).toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; }
function clampInt(v, min = 0) { const n = Number(v); if (!Number.isFinite(n)) return min; return Math.max(min, Math.round(n)); }
function cleanNum(v) { if (v === "" || v == null) return null; const n = Number(v); return Number.isFinite(n) ? n : null; }
function dateStr(d) { return new Date(d).toISOString().split("T")[0]; }

const DIFF_FACTOR = { easy: 1.0, normal: 1.25, hard: 1.5 };
const DIFF_OPTS = [
  { value: "easy", label: "普通招募", factor: 1.0 },
  { value: "normal", label: "较难招募", factor: 1.25 },
  { value: "hard", label: "难招募", factor: 1.5 },
];

function recruitBaseDays(totalN) {
  if (totalN <= 30) return 5;
  if (totalN <= 100) return 9;
  if (totalN <= 200) return 11;
  if (totalN <= 300) return 14;
  if (totalN <= 500) return 21;
  return 28;
}
function calcRecruitDays(totalN, diff) {
  return Math.min(28, Math.round(recruitBaseDays(totalN) * (DIFF_FACTOR[diff] || 1)));
}
function recruitWeeksStr(totalN, diff) {
  const days = calcRecruitDays(totalN, diff);
  const lo = Math.max(0.5, Math.round((days / 7) * 10) / 10);
  const hi = Math.max(lo, Math.round(((days + 3) / 7) * 10) / 10);
  return `约 ${lo}-${hi} 周`;
}
function calcToplineDays(products) {
  return Math.round(7 + Math.max(0, Math.ceil((products - 2) / 2)) * 3.5);
}
function calcReportDays(products) {
  return Math.round(7 + Math.max(0, products - 1) * 3.5);
}

let _nid = 500;
const SYSTEM_IDS = new Set([
  "q1","q2","q3","q4","q5","q6","q7","q8","q9","q10",
  "ql1","ql2","ql3","ql4","ql5","ql6","ql7","ql8",
  "a1","a2","a3","a4","d1","d2","d3","d4","d5"
]);

const INIT_CATALOG = [
  { id:"q1",  cat:"定量·HUT留置", name:"留置测试（1周·1份问卷）",              unitPrice:210,  calcType:"ppp", notes:"标准HUT单产品",  hutWeeks:1, hutQuestionnaires:1, hutProducts:1, locked:true },
  { id:"q2",  cat:"定量·HUT留置", name:"顺序型留置（2产品·每产品1周·2份问卷）", unitPrice:300,  calcType:"ppp", notes:"顺序型",         hutWeeks:2, hutQuestionnaires:2, hutProducts:2, locked:true },
  { id:"q3",  cat:"定量·HUT留置", name:"宣称支持留置（1周·1份问卷）",          unitPrice:200,  calcType:"ppp", notes:"",              hutWeeks:1, hutQuestionnaires:1, hutProducts:1, locked:true },
  { id:"q4",  cat:"定量·HUT留置", name:"宣称支持留置（2周·1份问卷）",          unitPrice:230,  calcType:"ppp", notes:"",              hutWeeks:2, hutQuestionnaires:1, hutProducts:1, locked:true },
  { id:"q5",  cat:"定量·HUT留置", name:"宣称支持留置（1周·2份问卷）",          unitPrice:260,  calcType:"ppp", notes:"",              hutWeeks:1, hutQuestionnaires:2, hutProducts:1, locked:true },
  { id:"q6",  cat:"定量·HUT留置", name:"额外：每增加1周使用",                  unitPrice:25,   calcType:"ppp", notes:"叠加在基础款上", isExtraWeek:true, locked:true },
  { id:"q7",  cat:"定量·HUT留置", name:"额外：每增加1份回访问卷",              unitPrice:50,   calcType:"ppp", notes:"",              isExtraQuestionnaire:true, locked:true },
  { id:"q8",  cat:"定量·HUT留置", name:"额外：每增加1个留置产品（仅单产品为套装时）", unitPrice:20, calcType:"ppp", notes:"仅用于单产品内的额外套装件数", isExtraProduct:true, locked:true },
  { id:"q9",  cat:"定量·CLT定点", name:"定点测试CLT（30分钟内）",              unitPrice:280,  calcType:"pp",  notes:"", locked:true },
  { id:"q10", cat:"定量·CLT定点", name:"定点测试CLT（60分钟内）",              unitPrice:350,  calcType:"pp",  notes:"", locked:true },
  { id:"ql1", cat:"定性研究",     name:"一对一深访（1小时·线上）",             unitPrice:4167, calcType:"ses", notes:"差旅不含", locked:true },
  { id:"ql2", cat:"定性研究",     name:"一对一深访（1小时·线下）",             unitPrice:4333, calcType:"ses", notes:"差旅另计", locked:true },
  { id:"ql3", cat:"定性研究",     name:"家访（1.5小时）",                     unitPrice:5500, calcType:"ses", notes:"差旅另计", locked:true },
  { id:"ql4", cat:"定性研究",     name:"座谈会FGD（2小时·线上）",             unitPrice:16500,calcType:"ses", notes:"每组6-8人", locked:true },
  { id:"ql5", cat:"定性研究",     name:"座谈会FGD（2小时·线下）",             unitPrice:18000,calcType:"ses", notes:"差旅另计", locked:true },
  { id:"ql6", cat:"定性研究",     name:"定性额外：每增加1周产品使用",           unitPrice:25,   calcType:"pp",  notes:"", locked:true },
  { id:"ql7", cat:"定性研究",     name:"定性额外：每增加1份回访问卷",           unitPrice:50,   calcType:"pp",  notes:"", locked:true },
  { id:"ql8", cat:"定性研究",     name:"定性额外：每增加1个留置产品",           unitPrice:20,   calcType:"pp",  notes:"", locked:true },
  { id:"a1",  cat:"高阶技术",     name:"眼动追踪（额外加收）",                 unitPrice:150,  calcType:"pp",  notes:"", locked:true },
  { id:"a2",  cat:"高阶技术",     name:"微表情识别（额外加收）",               unitPrice:150,  calcType:"pp",  notes:"", locked:true },
  { id:"a3",  cat:"高阶技术",     name:"脑电监测（额外加收）",                 unitPrice:500,  calcType:"pp",  notes:"", locked:true },
  { id:"a4",  cat:"高阶技术",     name:"数据建模（感官+消费者）",              unitPrice:50000,calcType:"fix", notes:"每个模型", locked:true },
  { id:"d1",  cat:"交付与通用",   name:"Excel版数据简报",                    unitPrice:1600, calcType:"prod",notes:"", locked:true },
  { id:"d2",  cat:"交付与通用",   name:"完整报告",                           unitPrice:4000, calcType:"prod",notes:"", locked:true },
  { id:"d3",  cat:"交付与通用",   name:"项目管理费",                         unitPrice:1000, calcType:"prod",notes:"含样本小组建立/维护", locked:true },
  { id:"d4",  cat:"交付与通用",   name:"盲包",                              unitPrice:20,   calcType:"ppp", notes:"需灏图提供盲包时", locked:true },
  ];

const CALC_LABELS = { ppp:"单产品N × 整体产品数", pp:"× 总样本量", prod:"× 整体产品数", ses:"× 场次数", fix:"固定金额" };
const CATS = ["交付与通用","定量·HUT留置","定量·CLT定点","定性研究","高阶技术"];
const HUT_BASE_IDS = ["q1", "q2", "q3", "q4", "q5"];
const ADJUSTABLE_PHASES = ["screen","recruit","mainQ","topline","report"];

function lineFromCatalog(ci, more = {}) {
  return {
    lid: ci.id,
    qtyO: null,
    priceO: null,
    mulO: ci.isExtraWeek || ci.isExtraQuestionnaire || ci.isExtraProduct ? 1 : null,
    splitEnabled: false,
    splitPanelRatio: 50,
    splitOutsideFactor: 1.2,
    ...more
  };
}

export default function App() {
  const [tab, setTab] = useState("quote");
  const [catalog, setCatalog] = useState(INIT_CATALOG);
  const [p, setP_] = useState({
    title: "HUT 留置测试",
    designBaseId: "q1",
    perProductN: 45,
    totalNAuto: true,
    totalNManual: 45,
    designProducts: 1,
    designWeeks: 1,
    designQuestionnaires: 1,
    setItemsPerProduct: 1,
    sessions: 0,
    courierDays: 2,
    difficulty: "normal",
    vatRate: 6,
    startDate: new Date().toISOString().split("T")[0],
    recruitDaysManual: "",
    toplineDaysManual: "",
    reportDaysManual: "",
    screenDaysManual: "",
    mainQDaysManual: "",
    ddlMode: "none",
    targetToplineEnd: "",
    targetReportEnd: "",
    manualTimelineEdit: false,
  });
  const sp = (k, v) => setP_(c => ({ ...c, [k]: v }));
  const [lines, setLines] = useState([]);

  const getCi = (id) => catalog.find(c => c.id === id);
  const selectedBase = getCi(p.designBaseId) || getCi("q1");
  const hasQual = useMemo(() => lines.some(l => getCi(l.lid)?.cat === "定性研究"), [lines, catalog]);

  const design = useMemo(() => {
    const base = selectedBase || { hutWeeks: 1, hutQuestionnaires: 1, hutProducts: 1 };
    return {
      hasHUT: true,
      baseId: base.id,
      baseName: base.name,
      overallProducts: clampInt(p.designProducts, 1),
      hutWeeks: clampInt(p.designWeeks, 1),
      hutQuestionnaires: clampInt(p.designQuestionnaires, 1),
      setItemsPerProduct: clampInt(p.setItemsPerProduct, 1),
      baseWeeks: base.hutWeeks || 1,
      baseQuestionnaires: base.hutQuestionnaires || 1,
      baseProducts: base.hutProducts || 1,
    };
  }, [selectedBase, p.designWeeks, p.designQuestionnaires, p.designProducts, p.setItemsPerProduct]);

  const totalN = p.totalNAuto ? clampInt(p.perProductN, 1) * design.overallProducts : clampInt(p.totalNManual, 1);
  const totalNDisplay = clampInt(p.perProductN, 1) * design.overallProducts;
  const recruitDaysAuto = calcRecruitDays(totalN, p.difficulty);
  const recruitDays = cleanNum(p.recruitDaysManual) ?? recruitDaysAuto;
  const screenDays = cleanNum(p.screenDaysManual) ?? 3;
  const mainQDays = cleanNum(p.mainQDaysManual) ?? 3;
  const leaveDays = design.hutWeeks * 7 + clampInt(p.courierDays, 0);
  const toplineDaysAuto = calcToplineDays(design.overallProducts);
  const reportDaysAuto = calcReportDays(design.overallProducts);
  const toplineDays = cleanNum(p.toplineDaysManual) ?? toplineDaysAuto;
  const reportDays = cleanNum(p.reportDaysManual) ?? reportDaysAuto;

  const validations = useMemo(() => {
    const issues = [];
    const selected = lines.map(l => getCi(l.lid)).filter(Boolean);
    const baseLines = selected.filter(ci => ci.hutWeeks != null && !ci.isExtraWeek && !ci.isExtraQuestionnaire && !ci.isExtraProduct);
    if (baseLines.length > 1) issues.push("报价中存在多个 HUT 基础项，建议仅保留 1 个。系统已按研究设计区同步的结果为准。");

    const byId = (id) => lines.find(l => l.lid === id);
    const weekExtra = byId("q6");
    const questExtra = byId("q7");
    const prodExtra = byId("q8");
    const weekDiff = Math.max(0, design.hutWeeks - design.baseWeeks);
    const questDiff = Math.max(0, design.hutQuestionnaires - design.baseQuestionnaires);
    const setDiff = Math.max(0, design.setItemsPerProduct - 1);

    if ((weekExtra?.mulO ?? (weekExtra ? 1 : 0)) !== weekDiff) issues.push("“额外增加周数”与当前研究设计不一致，建议点击“同步报价项”。");
    if ((questExtra?.mulO ?? (questExtra ? 1 : 0)) !== questDiff) issues.push("“额外增加问卷”与当前研究设计不一致，建议点击“同步报价项”。");
    if ((prodExtra?.mulO ?? (prodExtra ? 1 : 0)) !== setDiff) issues.push("“额外增加留置产品（套装件数）”与当前研究设计不一致，建议点击“同步报价项”。");
    if (!lines.find(l => l.lid === design.baseId)) issues.push("当前研究设计对应的 HUT 基础项尚未加入报价，建议点击“同步报价项”。");
    if (design.overallProducts > 1 && setDiff > 0) issues.push("当前已将“整体产品数”和“单产品内套装件数”同时增加，请确认 q8 仅用于单产品为套装的情况。");

    return issues;
  }, [lines, design, catalog]);

  const autoQty = (ci) => {
    if (ci.isExtraProduct) return clampInt(p.perProductN, 1) * design.overallProducts;
    switch (ci.calcType) {
      case "ppp": return clampInt(p.perProductN, 1) * design.overallProducts;
      case "pp": return totalN;
      case "prod": return design.overallProducts;
      case "ses": return hasQual ? clampInt(p.sessions, 0) : 0;
      default: return 1;
    }
  };

  const expandedLineRows = useMemo(() => {
    const rows = [];
    lines.forEach((l) => {
      const ci = getCi(l.lid);
      if (!ci) return;
      const qty = l.qtyO ?? autoQty(ci);
      const price = l.priceO ?? ci.unitPrice;
      const mul = l.mulO ?? 1;
      if (l.splitEnabled) {
        const panelRatio = Math.min(100, Math.max(0, clampInt(l.splitPanelRatio, 0)));
        const panelQty = Math.round(qty * panelRatio / 100);
        const regularQty = qty - panelQty;
        const outsideFactor = Math.max(0, Number(l.splitOutsideFactor || 1.2));
        if (panelQty > 0) rows.push({ ...l, ci, qty: panelQty, price, mul, total: panelQty * price * mul, splitTag:"Panel报价", splitNote:`Panel报价 ${panelQty}` });
        if (regularQty > 0) rows.push({ ...l, ci, qty: regularQty, price: +(price * outsideFactor).toFixed(2), mul, total: regularQty * +(price * outsideFactor).toFixed(2) * mul, splitTag:"常规报价", splitNote:`常规报价 ${regularQty}（${outsideFactor}倍）` });
      } else {
        rows.push({ ...l, ci, qty, price, mul, total: qty * price * mul, splitTag:"", splitNote:"" });
      }
    });
    return rows;
  }, [lines, catalog, p, design, totalN, hasQual]);

  const sub = expandedLineRows.reduce((s, r) => s + r.total, 0);
  const gross = +(sub * (1 + clampInt(p.vatRate, 0) / 100)).toFixed(2);

  const syncHUTLines = () => {
    const keep = lines.filter(l => {
      const ci = getCi(l.lid);
      return ci && ci.cat !== "定量·HUT留置";
    });
    const next = [...keep, lineFromCatalog(getCi(design.baseId), { mulO: null })];
    const weekDiff = Math.max(0, design.hutWeeks - design.baseWeeks);
    const questDiff = Math.max(0, design.hutQuestionnaires - design.baseQuestionnaires);
    const setDiff = Math.max(0, design.setItemsPerProduct - 1);
    if (weekDiff > 0) next.push(lineFromCatalog(getCi("q6"), { mulO: weekDiff }));
    if (questDiff > 0) next.push(lineFromCatalog(getCi("q7"), { mulO: questDiff }));
    if (setDiff > 0) next.push(lineFromCatalog(getCi("q8"), { mulO: setDiff }));
    setLines(next);
  };

  const addLine = (id) => {
    if (lines.find(l => l.lid === id)) return;
    const ci = getCi(id);
    if (!ci) return;
    if (ci.hutWeeks != null && !ci.isExtraWeek && !ci.isExtraQuestionnaire && !ci.isExtraProduct) {
      sp("designBaseId", ci.id);
      sp("designWeeks", ci.hutWeeks);
      sp("designQuestionnaires", ci.hutQuestionnaires);
      sp("setItemsPerProduct", 1);
      sp("designProducts", ci.hutProducts);
    }
    setLines(ls => [...ls, lineFromCatalog(ci)]);
  };
  const removeLine = (id) => setLines(ls => ls.filter(l => l.lid !== id));
  const updLine = (id, k, v) => {
    const ci = getCi(id);
    setLines(ls => ls.map(l => {
      if (l.lid !== id) return l;
      if (k === "splitEnabled") return { ...l, splitEnabled: !!v };
      if (k === "splitPanelRatio") {
        const n = v === "" ? 0 : Number(v);
        return { ...l, splitPanelRatio: Number.isFinite(n) ? Math.min(100, Math.max(0, Math.round(n))) : 0 };
      }
      if (k === "splitOutsideFactor") {
        const n = v === "" ? 0 : Number(v);
        return { ...l, splitOutsideFactor: Number.isFinite(n) ? Math.max(0, n) : 0 };
      }
      const val = v === "" ? null : Math.max(0, Number(v));
      if ((ci?.hutWeeks != null) && k === "mulO") return { ...l, [k]: null };
      return { ...l, [k]: Number.isFinite(val) ? val : null };
    }));
  };
  const updCat = (id, k, v) => setCatalog(arr => arr.map(c => c.id === id ? { ...c, [k]: v } : c));
  const delCat = (id) => {
    const ci = getCi(id);
    if (!ci || ci.locked || SYSTEM_IDS.has(id)) return;
    setCatalog(arr => arr.filter(c => c.id !== id));
    setLines(ls => ls.filter(l => l.lid !== id));
  };
  const addCat = () => {
    _nid += 1;
    setCatalog(arr => [...arr, { id:`x${_nid}`, cat:"交付与通用", name:"新费用项", unitPrice:0, calcType:"fix", notes:"", locked:false,  }]);
  };

    const autoProjection = useMemo(() => {
    const confirmDate = skipWE(new Date(p.startDate));
    const screenStart = confirmDate;
    const screenEnd = addBusinessDays(screenStart, screenDays);
    const recruitStart = skipWE(addDays(screenEnd, 1));
    const recruitEnd = addBusinessDays(recruitStart, recruitDays);
    const mainQStart = skipWE(addDays(recruitEnd, 1));
    const mainQEnd = addBusinessDays(mainQStart, mainQDays);
    const leaveStart = skipWE(addDays(mainQEnd, 1));
    const leaveEnd = addDays(leaveStart, leaveDays);
    const toplineStart = skipWE(addDays(leaveEnd, 1));
    const toplineEnd = addBusinessDays(toplineStart, toplineDays);
    const reportStart = skipWE(addDays(toplineEnd, 1));
    const reportEnd = addBusinessDays(reportStart, reportDays);
    const attr = skipWE(addDays(recruitEnd, -2));
    return {
      confirmDate, screenStart, screenEnd, recruitStart, recruitEnd,
      mainQStart, mainQEnd, leaveStart, leaveEnd, toplineStart, toplineEnd,
      reportStart, reportEnd, attr
    };
  }, [p.startDate, screenDays, recruitDays, mainQDays, leaveDays, toplineDays, reportDays]);

  const ddlAnalysis = useMemo(() => {
    const targetTop = p.targetToplineEnd ? skipWE(new Date(p.targetToplineEnd)) : null;
    const targetRep = p.targetReportEnd ? skipWE(new Date(p.targetReportEnd)) : null;
    const curTop = autoProjection.toplineEnd;
    const curRep = autoProjection.reportEnd;
    const messages = [];
    const requiredReductionTopline = targetTop ? Math.max(0, Math.ceil((curTop - targetTop) / 86400000)) : 0;
    const requiredReductionReport = targetRep ? Math.max(0, Math.ceil((curRep - targetRep) / 86400000)) : 0;
    const daysAheadTopline = targetTop ? Math.max(0, Math.ceil((targetTop - curTop) / 86400000)) : 0;
    const daysAheadReport = targetRep ? Math.max(0, Math.ceil((targetRep - curRep) / 86400000)) : 0;
    if (!targetTop && !targetRep) {
      messages.push("未设置 DDL，当前显示为正常计算时间表。");
    } else {
      if (targetTop) {
        if (requiredReductionTopline > 0) messages.push(`Topline 截止时间：当前仍需加急 ${requiredReductionTopline} 天。`);
        else if (daysAheadTopline > 0) messages.push(`Topline 截止时间：当前可提前 ${daysAheadTopline} 天完成。`);
        else messages.push("Topline 截止时间：当前排期刚好满足。");
      }
      if (targetRep) {
        if (requiredReductionReport > 0) messages.push(`Report 截止时间：当前仍需加急 ${requiredReductionReport} 天。`);
        else if (daysAheadReport > 0) messages.push(`Report 截止时间：当前可提前 ${daysAheadReport} 天完成。`);
        else messages.push("Report 截止时间：当前排期刚好满足。");
      }
    }
    const needRush = requiredReductionTopline > 0 || requiredReductionReport > 0;
    return {
      active: !!(targetTop || targetRep),
      targetTop, targetRep,
      requiredReductionTopline, requiredReductionReport,
      daysAheadTopline, daysAheadReport,
      needRush,
      messages
    };
  }, [p.targetToplineEnd, p.targetReportEnd, autoProjection]);

  const timelineEditable = p.manualTimelineEdit || ddlAnalysis.needRush;

  const resetTimelineManuals = () => {
    setP_(c => ({
      ...c,
      screenDaysManual: "",
      recruitDaysManual: "",
      mainQDaysManual: "",
      toplineDaysManual: "",
      reportDaysManual: "",
    }));
  };

  const updTimelineDays = (key, value) => {
    if (!timelineEditable) return;
    sp(key, value);
  };

  const timeline = useMemo(() => {
    const a = autoProjection;
    return [
      { key:"confirm", phase:"项目确认 / 报价", days:"/", start:"—", end:fmt(a.confirmDate), note:"固定为报价日期", editable:false },
      { key:"screen", phase:"甄别问卷确认", days:screenDays, start:fmt(a.screenStart), end:fmt(a.screenEnd), note:"按工作日计算；仅在需加急或主动开启手动调整时可修改", editable:timelineEditable, manualKey:"screenDaysManual", autoDays:3 },
      { key:"recruit", phase:`招募（${recruitWeeksStr(totalN, p.difficulty)}）`, days:recruitDays, start:fmt(a.recruitStart), end:fmt(a.recruitEnd), note:`按工作日计算；自动估算=${recruitDaysAuto}天`, editable:timelineEditable, manualKey:"recruitDaysManual", autoDays:recruitDaysAuto },
      { key:"mainQ", phase:"主问卷确认", days:mainQDays, start:fmt(a.mainQStart), end:fmt(a.mainQEnd), note:"按工作日计算；仅在需加急或主动开启手动调整时可修改", editable:timelineEditable, manualKey:"mainQDaysManual", autoDays:3 },
      { key:"leave", phase:`派发和留置（含快递${clampInt(p.courierDays, 0)}天）`, days:leaveDays, start:fmt(a.leaveStart), end:fmt(a.leaveEnd), note:`按自然日计算，固定不可压缩：${design.hutWeeks}周 × 7天 + 快递${clampInt(p.courierDays, 0)}天`, editable:false },
      { key:"topline", phase:`Topline（${design.overallProducts}款）`, days:toplineDays, start:fmt(a.toplineStart), end:fmt(a.toplineEnd), note:`按工作日计算；自动=${toplineDaysAuto}天`, editable:timelineEditable, manualKey:"toplineDaysManual", autoDays:toplineDaysAuto },
      { key:"report", phase:`Report（${design.overallProducts}款）`, days:reportDays, start:fmt(a.reportStart), end:fmt(a.reportEnd), note:`按工作日计算；自动=${reportDaysAuto}天`, editable:timelineEditable, manualKey:"reportDaysManual", autoDays:reportDaysAuto },
    ];
  }, [autoProjection, screenDays, recruitDays, mainQDays, leaveDays, toplineDays, reportDays, timelineEditable, totalN, p.difficulty, recruitDaysAuto, design, p.courierDays, toplineDaysAuto, reportDaysAuto]);

  const exportXLSX = () => {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet([
      [p.title],
      [`研究设计：${design.baseName}｜整体产品数:${design.overallProducts}｜单产品套装件数:${design.setItemsPerProduct}｜单产品N:${p.perProductN}｜总N:${totalN}｜留置:${design.hutWeeks}周｜问卷:${design.hutQuestionnaires}份`],
      [],
      ["费用项目","拆分","计算方式","单价(RMB)","数量","倍数","小计(RMB)"],
      ...expandedLineRows.map(r => [r.ci.name, r.splitTag || "—", CALC_LABELS[r.ci.calcType], r.price, r.qty, r.mul, r.total]),
      [],
      [`Cost (before VAT ${p.vatRate}%)`, "", "", "", "", "", sub],
      [`Cost (after VAT ${p.vatRate}%)`, "", "", "", "", "", gross],
    ]);
    ws1["!cols"] = [{wch:42},{wch:8},{wch:18},{wch:12},{wch:10},{wch:8},{wch:16}];
    XLSX.utils.book_append_sheet(wb, ws1, "报价");

    const ws2 = XLSX.utils.aoa_to_sheet([
      [`${p.title} · 时间表`],
      [],
      ["阶段","时间（天）","开始日期","结束日期","说明"],
      ...timeline.map(r => [r.phase, r.days, r.start, r.end, r.note]),
    ]);
    ws2["!cols"] = [{wch:36},{wch:10},{wch:14},{wch:14},{wch:42}];
    XLSX.utils.book_append_sheet(wb, ws2, "时间表");
    XLSX.writeFile(wb, `${p.title}_报价.xlsx`);
  };

  const S = ST;
  return (
    <div style={{ fontFamily:"Georgia,serif", background:"#f4f1ec", minHeight:"100vh", color:"#1a1a1a" }}>
      <style>{css}</style>

      <div style={{ background:"#2c2825", padding:"18px 28px", display:"flex", alignItems:"center", justifyContent:"space-between", gap:12, flexWrap:"wrap" }}>
        <div>
          <div style={{ fontSize:10, letterSpacing:"0.12em", textTransform:"uppercase", color:"#9a8e80", marginBottom:4 }}>HOW-TO 灏图品测 · 智能报价系统</div>
          <input
            value={p.title}
            onChange={e => sp("title", e.target.value)}
            style={{ background:"transparent", border:"none", borderBottom:"1px solid #5a504a", color:"#f5f2ed", fontSize:20, fontWeight:700, fontFamily:"inherit", outline:"none", width:360, padding:"2px 0" }}
          />
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:14 }}>
          <div style={{ textAlign:"right" }}>
            <div style={{ fontSize:10, color:"#9a8e80", letterSpacing:"0.08em", textTransform:"uppercase" }}>含税总额</div>
            <div style={{ fontSize:26, fontWeight:700, fontFamily:"monospace", color:"#e8c99a" }}>{fmtRMB(gross)}</div>
          </div>
          <button className="export-btn" onClick={exportXLSX}>⬇ 导出 Excel</button>
        </div>
      </div>

      <div style={{ display:"grid", gridTemplateColumns:"360px 1fr", gap:18, padding:"18px 28px", maxWidth:1540, margin:"0 auto" }}>
        <div>
          <div style={S.card}>
            <div style={S.cardTitle}>研究设计（优先输入）</div>
            <F label="HUT 基础设计">
              <select className="inp" value={p.designBaseId} onChange={e => {
                const ci = getCi(e.target.value);
                sp("designBaseId", e.target.value);
                if (ci) {
                  sp("designWeeks", ci.hutWeeks || 1);
                  sp("designQuestionnaires", ci.hutQuestionnaires || 1);
                  sp("setItemsPerProduct", 1);
                  sp("designProducts", ci.hutProducts || 1);
                }
              }}>
                {HUT_BASE_IDS.map(id => {
                  const ci = getCi(id);
                  return <option key={id} value={id}>{ci?.name}</option>;
                })}
              </select>
            </F>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
              <F label="单产品样本量">
                <input className="inp" type="number" min={1} value={p.perProductN} onChange={e => sp("perProductN", clampInt(e.target.value, 1))} />
              </F>
              <F label="整体测试产品数">
                <input className="inp" type="number" min={1} value={p.designProducts} onChange={e => sp("designProducts", clampInt(e.target.value, 1))} />
              </F>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
              <F label="总样本量模式">
                <select className="inp" value={p.totalNAuto ? "auto" : "manual"} onChange={e => sp("totalNAuto", e.target.value === "auto")}>
                  <option value="auto">自动计算</option>
                  <option value="manual">手动输入</option>
                </select>
              </F>
              <F label={p.totalNAuto ? "自动总样本量" : "手动总样本量"}>
                <input className="inp" type="number" min={1} value={p.totalNAuto ? totalNDisplay : p.totalNManual} disabled={p.totalNAuto} onChange={e => sp("totalNManual", clampInt(e.target.value, 1))} />
              </F>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
              <F label="单产品内套装件数">
                <input className="inp" type="number" min={1} value={p.setItemsPerProduct} onChange={e => sp("setItemsPerProduct", clampInt(e.target.value, 1))} />
              </F>
              <F label="联动结果">
                <div className="inp" style={{ background:"#f7f4ef", display:"flex", alignItems:"center", minHeight:36 }}>
                  {clampInt(p.perProductN, 1)} × {design.overallProducts} = {totalNDisplay}
                </div>
              </F>
            </div>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
              <F label="留置周数">
                <input className="inp" type="number" min={1} value={p.designWeeks} onChange={e => sp("designWeeks", clampInt(e.target.value, 1))} />
              </F>
              <F label="问卷份数">
                <input className="inp" type="number" min={1} value={p.designQuestionnaires} onChange={e => sp("designQuestionnaires", clampInt(e.target.value, 1))} />
              </F>
            </div>
            <div style={{ display:"flex", gap:8, marginTop:8, flexWrap:"wrap" }}>
              <button onClick={syncHUTLines} style={S.btnDark}>同步 HUT 报价项</button>
              <button onClick={() => { sp("designBaseId", "q1"); sp("designProducts", 1); sp("setItemsPerProduct", 1); sp("designWeeks", 1); sp("designQuestionnaires", 1); }} style={S.btnGhost}>恢复基础设计</button>
            </div>
            <div style={{ background:"#f0ebe3", borderRadius:6, padding:"12px", marginTop:14, fontSize:12, lineHeight:2 }}>
              {[
                ["当前基础项", design.baseName],
                ["单产品样本量", `${clampInt(p.perProductN, 1)} 人`],
                ["整体测试产品数", `${design.overallProducts} 款`],
                ["总样本量", `${totalN} 人`],
                ["单产品内套装件数", `${design.setItemsPerProduct} 件`],
                ["留置周数", `${design.hutWeeks} 周`],
                ["问卷份数", `${design.hutQuestionnaires} 份`],
              ].map(([k, v]) => (
                <div key={k} style={{ display:"flex", justifyContent:"space-between", borderBottom:"1px solid #e4dfd7", padding:"3px 0" }}>
                  <span style={{ color:"#7a6e5f" }}>{k}</span>
                  <span style={{ fontFamily:"monospace", fontWeight:700, color:"#2c2825", marginLeft:12, textAlign:"right" }}>{v}</span>
                </div>
              ))}
            </div>
            <div style={S.helpBox}>说明：整体测试产品数只影响总样本量与 Topline / Report；“额外：每增加1个留置产品”仅对应单产品为套装时的额外件数。</div>
          </div>

          <div style={S.card}>
            <div style={S.cardTitle}>手动参数</div>
            <F label="VAT 税率 (%)">
              <input className="inp" type="number" min={0} max={20} value={p.vatRate} onChange={e => sp("vatRate", clampInt(e.target.value, 0))} />
            </F>
          </div>

          {validations.length > 0 && (
            <div style={{ background:"#fff7ed", border:"1px solid #f1d4a9", borderRadius:8, padding:"14px 16px", color:"#8a5a0a", fontSize:12, lineHeight:1.9, marginBottom:12 }}>
              <div style={{ fontWeight:700, marginBottom:6 }}>校验提醒</div>
              {validations.map((msg, i) => <div key={i}>• {msg}</div>)}
            </div>
          )}
        </div>

        <div>
          <div style={{ display:"flex", gap:3, marginBottom:12 }}>
            {[ ["quote","📋 报价明细"], ["catalog","📚 费用目录"], ["timeline","📅 时间表"] ].map(([id, label]) => (
              <button key={id} onClick={() => setTab(id)} style={tab === id ? S.tabOn : S.tabOff}>{label}</button>
            ))}
          </div>

          {tab === "quote" && (
            <div>
              <div style={S.card}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
                  <div style={S.cardTitle}>报价行项目</div>
                  <div style={{ display:"flex", gap:8, alignItems:"center" }}>
                    <button onClick={syncHUTLines} style={S.btnGhost}>同步 HUT 报价项</button>
                    <div style={{ fontSize:11, color:"#aaa" }}>定性 / CLT / 高阶技术 / 交付项仍可在目录中手动添加</div>
                  </div>
                </div>
                <table style={S.table}>
                  <thead>
                    <tr>
                      <th style={{ ...S.th, width:"26%" }}>费用项目</th>
                      <th style={{ ...S.th, width:"9%", textAlign:"center" }}>拆分</th>
                      <th style={{ ...S.th, width:"14%", fontSize:10 }}>计算基准</th>
                      <th style={{ ...S.th, width:"11%", textAlign:"right" }}>单价 (¥)</th>
                      <th style={{ ...S.th, width:"8%", textAlign:"right" }}>数量</th>
                      <th style={{ ...S.th, width:"8%", textAlign:"center" }}>倍数 ×</th>
                      <th style={{ ...S.th, width:"14%", textAlign:"right" }}>小计 (¥)</th>
                      <th style={{ ...S.th, width:"6%" }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {expandedLineRows.map((r, i) => {
                      const isBase = r.ci.hutWeeks != null && !r.ci.isExtraWeek && !r.ci.isExtraQuestionnaire && !r.ci.isExtraProduct;
                      const sourceLine = lines.find(x => x.lid === r.lid);
                      return (
                        <tr key={`${r.lid}-${r.splitTag || "base"}-${i}`} style={i % 2 === 1 ? { background:"#faf8f5" } : {}}>
                          <td style={S.td}>
                            <div>{r.ci.name}</div>
                            <div style={{ marginTop:3, display:"flex", gap:4, flexWrap:"wrap" }}>
                              {isBase && <span style={badgeBlue}>{r.ci.hutWeeks}周·{r.ci.hutQuestionnaires}卷·基础{r.ci.hutProducts}款</span>}
                              {r.ci.isExtraWeek && <span style={badgeAmber}>+{sourceLine?.mulO ?? 1}周</span>}
                              {r.ci.isExtraQuestionnaire && <span style={badgeAmber}>+{sourceLine?.mulO ?? 1}份问卷</span>}
                              {r.ci.isExtraProduct && <span style={badgeAmber}>套装内 +{sourceLine?.mulO ?? 1}件</span>}
                              {sourceLine?.splitEnabled && <span style={badgeGreen}>已拆分报价</span>}
                            </div>
                          </td>
                          <td style={{ ...S.td, textAlign:"center", fontSize:11 }}>
                            {(!r.splitTag || r.splitTag === "Panel报价") ? (
                              <div style={{ display:"flex", flexDirection:"column", gap:4, alignItems:"center" }}>
                                <label style={{ fontSize:10, color:"#666" }}>
                                  <input type="checkbox" checked={!!sourceLine?.splitEnabled} onChange={e => updLine(r.lid, "splitEnabled", e.target.checked)} style={{ marginRight:4 }} />
                                  拆分报价
                                </label>
                                {sourceLine?.splitEnabled ? (
                                  <div style={{ display:"flex", flexDirection:"column", gap:6, alignItems:"center", marginTop:4 }}>
                                    {r.splitTag && <div style={{ fontSize:10, color:"#2f6b45", fontWeight:700 }}>Panel报价</div>}
                                    <input
                                      type="number"
                                      min={0}
                                      max={100}
                                      step="1"
                                      value={sourceLine?.splitPanelRatio ?? 50}
                                      onChange={e => updLine(r.lid, "splitPanelRatio", e.target.value)}
                                      style={{ ...S.mini, width:56, textAlign:"center" }}
                                      title="Panel报价所占比例%"
                                    />
                                    <div style={{ fontSize:10, color:"#777" }}>Panel {(sourceLine?.splitPanelRatio ?? 50)}%</div>
                                  </div>
                                ) : null}
                              </div>
                            ) : (
                              <div style={{ fontSize:11, color:"#555" }}>常规报价</div>
                            )}
                          </td>
                          <td style={{ ...S.td, fontSize:10, color:"#bbb", lineHeight:1.4 }}>{CALC_LABELS[r.ci.calcType]}{r.ci.isExtraProduct ? "（按套装件数计）" : ""}</td>
                          <td style={{ ...S.td, textAlign:"right" }}>
                            {!r.splitTag ? (
                              <input type="number" placeholder={r.ci.unitPrice} value={sourceLine?.priceO ?? ""} onChange={e => updLine(r.lid, "priceO", e.target.value)} style={{ ...S.mini, width:82, textAlign:"right" }} title="留空=目录价" />
                            ) : <span style={{ fontFamily:"monospace" }}>{fmtRMB(r.price)}</span>}
                          </td>
                          <td style={{ ...S.td, textAlign:"right" }}>
                            {!r.splitTag ? (
                              <input type="number" placeholder={autoQty(r.ci)} value={sourceLine?.qtyO ?? ""} onChange={e => updLine(r.lid, "qtyO", e.target.value)} style={{ ...S.mini, width:58, textAlign:"right" }} title="留空=自动" />
                            ) : <span style={{ fontFamily:"monospace" }}>{r.qty}</span>}
                          </td>
                          <td style={{ ...S.td, textAlign:"center" }}>
                            {!r.splitTag ? (
                              <input type="number" min={1} placeholder={isBase ? "—" : "1"} value={isBase ? "" : (sourceLine?.mulO ?? "")} disabled={isBase} onChange={e => updLine(r.lid, "mulO", e.target.value)} style={{ ...S.mini, width:44, textAlign:"center", ...(isBase ? { background:"#ede9e2", color:"#aaa" } : {}) }} title={isBase ? "基础项不使用倍数" : "倍数，留空=1"} />
                            ) : <span>{r.mul}</span>}
                          </td>
                          <td style={{ ...S.td, textAlign:"right", fontFamily:"monospace", fontWeight:600 }}>{fmtRMB(r.total)}</td>
                          <td style={{ ...S.td, textAlign:"center" }}>{!r.splitTag && <button onClick={() => removeLine(r.lid)} style={S.delBtn}>×</button>}</td>
                        </tr>
                      );
                    })}
                    {!expandedLineRows.length && <tr><td colSpan={8} style={{ ...S.td, color:"#ccc", textAlign:"center", padding:28 }}>暂无费用项，可先同步 HUT 项，再去「费用目录」添加其他费用</td></tr>}
                  </tbody>
                </table>
                <div style={{ marginTop:12, display:"flex", justifyContent:"flex-end" }}>
                  <table style={{ borderCollapse:"collapse", minWidth:320 }}>
                    <tbody>
                      <tr style={{ background:"#f0ebe3" }}>
                        <td style={{ ...S.td, fontWeight:700, paddingLeft:16 }}>Cost (before VAT {p.vatRate}%)</td>
                        <td style={{ ...S.td, textAlign:"right", fontFamily:"monospace", fontWeight:700, paddingRight:16 }}>{fmtRMB(sub)}</td>
                      </tr>
                      <tr style={{ background:"#2c2825" }}>
                        <td style={{ ...S.td, fontWeight:700, color:"#f5f2ed", paddingLeft:16 }}>Cost (after VAT {p.vatRate}%)</td>
                        <td style={{ ...S.td, textAlign:"right", fontFamily:"monospace", fontWeight:700, color:"#e8c99a", paddingRight:16 }}>{fmtRMB(gross)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
              <div style={{ background:"#fffbf5", border:"1px solid #ece3d0", borderRadius:8, padding:"11px 15px", fontSize:11, color:"#9a8e80", lineHeight:2.1 }}>
                <strong>整体产品数影响总样本量：</strong>{p.perProductN} × {design.overallProducts} = <strong>{clampInt(p.perProductN, 1) * design.overallProducts}</strong>
                &emsp;|&emsp;<strong>q8 套装件数计价：</strong>{p.perProductN} × {design.overallProducts} × {Math.max(0, design.setItemsPerProduct - 1)}
                &emsp;|&emsp;报价明细支持对任意目录项启用拆分：原目录价显示为 Panel报价，拆分后的加价部分显示为常规报价，可调整常规报价占比与倍率
                &emsp;|&emsp;基础项倍数已锁定，避免金额与研究设计脱钩
              </div>
            </div>
          )}

          {tab === "catalog" && (
            <div style={S.card}>
              <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:14 }}>
                <div style={S.cardTitle}>费用目录管理</div>
                <div style={{ display:"flex", gap:8 }}>
                  <button onClick={() => setCatalog(INIT_CATALOG)} style={S.btnGhost}>重置默认</button>
                  <button onClick={addCat} style={S.btnDark}>+ 新增目录项</button>
                </div>
              </div>
              {CATS.map(cat => {
                const items = catalog.filter(c => c.cat === cat);
                if (!items.length) return null;
                return (
                  <div key={cat} style={{ marginBottom:20 }}>
                    <div style={{ fontSize:11, fontWeight:700, letterSpacing:"0.06em", color:"#7a6e5f", padding:"5px 10px", background:"#f0ebe3", borderRadius:4, marginBottom:6 }}>{cat}</div>
                    <table style={S.table}>
                      <thead>
                        <tr>
                          <th style={{ ...S.th, width:"28%" }}>名称</th>
                          <th style={{ ...S.th, width:"11%" }}>分类</th>
                          <th style={{ ...S.th, width:"11%", textAlign:"right" }}>单价</th>
                          <th style={{ ...S.th, width:"16%" }}>计算方式</th>
                          <th style={{ ...S.th, width:"13%" }}>备注</th>
                          <th style={{ ...S.th, width:"8%", textAlign:"center" }}>加入报价</th>
                          <th style={{ ...S.th, width:"5%" }}></th>
                        </tr>
                      </thead>
                      <tbody>
                        {items.map((ci, i) => {
                          const inQ = !!lines.find(l => l.lid === ci.id);
                          return (
                            <tr key={ci.id} style={i % 2 === 1 ? { background:"#faf8f5" } : {}}>
                              <td style={S.td}>
                                <input value={ci.name} onChange={e => updCat(ci.id, "name", e.target.value)} style={{ ...S.inlineInput, ...(ci.locked ? { background:"#f5f1eb" } : {}) }} />
                                {ci.hutWeeks != null && <div style={{ fontSize:9, color:"#3a5a9a", marginTop:3 }}>📐 {ci.hutWeeks}周 · {ci.hutQuestionnaires}份问卷 · 基础{ci.hutProducts}款</div>}
                                {ci.isExtraWeek && <div style={{ fontSize:9, color:"#8a5a0a", marginTop:3 }}>⊕ 每倍数 +1周留置</div>}
                                {ci.isExtraQuestionnaire && <div style={{ fontSize:9, color:"#8a5a0a", marginTop:3 }}>⊕ 每倍数 +1份问卷（仅影响报价，不影响Topline/Report）</div>}
                                {ci.isExtraProduct && <div style={{ fontSize:9, color:"#8a5a0a", marginTop:3 }}>⊕ 每倍数 +1个留置产品，仅用于套装件数</div>}
                                                                {ci.locked && <div style={{ fontSize:9, color:"#999", marginTop:3 }}>系统默认目录项，不可删除</div>}
                              </td>
                              <td style={S.td}>
                                <select value={ci.cat} onChange={e => updCat(ci.id, "cat", e.target.value)} style={{ ...S.inlineInput, fontSize:10, padding:"3px 4px" }}>
                                  {CATS.map(c => <option key={c} value={c}>{c}</option>)}
                                </select>
                              </td>
                              <td style={{ ...S.td, textAlign:"right" }}>
                                <input type="number" value={ci.unitPrice} onChange={e => updCat(ci.id, "unitPrice", clampInt(e.target.value, 0))} style={{ ...S.inlineInput, width:72, textAlign:"right", fontFamily:"monospace" }} />
                              </td>
                              <td style={S.td}>
                                <select value={ci.calcType} onChange={e => updCat(ci.id, "calcType", e.target.value)} style={{ ...S.inlineInput, fontSize:10, padding:"3px 4px" }}>
                                  {Object.entries(CALC_LABELS).map(([k, v]) => <option key={k} value={k}>{v}</option>)}
                                </select>
                              </td>
                              <td style={S.td}><input value={ci.notes} onChange={e => updCat(ci.id, "notes", e.target.value)} style={{ ...S.inlineInput, fontSize:11 }} placeholder="—" /></td>
                              <td style={{ ...S.td, textAlign:"center" }}>{inQ ? <span style={{ fontSize:11, color:"#27ae60", fontWeight:700 }}>✓ 已选</span> : <button onClick={() => addLine(ci.id)} style={S.addBtn}>＋</button>}</td>
                              <td style={{ ...S.td, textAlign:"center" }}>{ci.locked ? <span style={{ color:"#bbb" }}>—</span> : <button onClick={() => delCat(ci.id)} style={S.delBtn}>×</button>}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                );
              })}
            </div>
          )}

          {tab === "timeline" && (
            <div style={S.card}>
              <div style={S.cardTitle}>项目时间表 · 工作阶段避开中国周末</div>
              <div style={{ display:"grid", gridTemplateColumns:"1.1fr 1fr", gap:16, marginBottom:14 }}>
                <div style={{ background:"#f7f4ef", borderRadius:6, padding:"12px", border:"1px solid #ede9e2" }}>
                  <div style={{ fontSize:10, letterSpacing:"0.08em", textTransform:"uppercase", color:"#7a6e5f", marginBottom:10, fontWeight:700 }}>基础参数</div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                    <F label="项目确认日期">
                      <input className="inp" type="date" value={p.startDate} onChange={e => sp("startDate", e.target.value)} />
                    </F>
                    <F label="快递时间（天）">
                      <input className="inp" type="number" min={0} max={14} value={p.courierDays} onChange={e => sp("courierDays", clampInt(e.target.value, 0))} />
                    </F>
                  </div>
                  <F label="招募难度" last>
                    <select className="inp" value={p.difficulty} onChange={e => sp("difficulty", e.target.value)}>
                      {DIFF_OPTS.map(d => <option key={d.value} value={d.value}>{d.label}（系数×{d.factor}）</option>)}
                    </select>
                  </F>
                </div>
                <div style={{ background:"#fff7ed", border:"1px solid #f1d4a9", borderRadius:6, padding:"12px" }}>
                  <div style={{ fontSize:10, letterSpacing:"0.08em", textTransform:"uppercase", color:"#8a5a0a", marginBottom:10, fontWeight:700 }}>DDL 与操作</div>
                  <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
                    <F label="Topline 截止时间">
                      <input className="inp" type="date" value={p.targetToplineEnd} onChange={e => sp("targetToplineEnd", e.target.value)} />
                    </F>
                    <F label="Report 截止时间">
                      <input className="inp" type="date" value={p.targetReportEnd} onChange={e => sp("targetReportEnd", e.target.value)} />
                    </F>
                  </div>
                  <div style={{ display:"flex", gap:8, flexWrap:"wrap", marginTop:4 }}>
                    <button onClick={() => sp("manualTimelineEdit", !p.manualTimelineEdit)} style={S.btnGhost}>
                      {p.manualTimelineEdit ? "关闭手动调整" : "开启手动调整"}
                    </button>
                    <button onClick={resetTimelineManuals} style={S.btnGhost}>恢复自动计算</button>
                  </div>
                  <div style={{ marginTop:10, padding:"10px 12px", background:"rgba(255,255,255,0.55)", borderRadius:6, fontSize:11, color:"#8a5a0a", lineHeight:1.9 }}>
                    {ddlAnalysis.messages.map((msg, i) => <div key={i}>{msg}</div>)}
                    <div>{timelineEditable ? "当前可编辑时间表天数。" : "当前时间表为只读；如需调整，请填写紧急DDL或手动开启调整。"}</div>
                  </div>
                  <div style={{ marginTop:12, padding:"10px 12px", background:"#fff", borderRadius:6, border:"1px solid #f3e4ca", fontSize:11, lineHeight:1.9, color:"#7a6e5f" }}>
                    <div><strong>当前自动招募：</strong>{recruitDaysAuto} 天 / {recruitWeeksStr(totalN, p.difficulty)}</div>
                    <div><strong>固定留置：</strong>{leaveDays} 天（{design.hutWeeks}周×7 + 快递{p.courierDays}天）</div>
                    <div><strong>Topline：</strong>{toplineDaysAuto} 天　<strong>Report：</strong>{reportDaysAuto} 天</div>
                    <div><strong>总样本量：</strong>{p.totalNAuto ? `${clampInt(p.perProductN, 1)} × ${design.overallProducts} = ${totalNDisplay}` : `手动输入 ${totalN}`}</div>
                  </div>
                </div>
              </div>
              <table style={S.table}>
                <thead>
                  <tr>
                    <th style={{ ...S.th, width:"30%" }}>阶段</th>
                    <th style={{ ...S.th, textAlign:"right", width:"10%" }}>天数</th>
                    <th style={{ ...S.th, textAlign:"center", width:"16%" }}>开始</th>
                    <th style={{ ...S.th, textAlign:"center", width:"16%" }}>结束</th>
                    <th style={{ ...S.th, width:"28%" }}>计算说明</th>
                  </tr>
                </thead>
                <tbody>
                  {timeline.map((r, i) => (
                    <tr key={r.key} style={i % 2 === 1 ? { background:"#faf8f5" } : {}}>
                      <td style={{ ...S.td, fontWeight:500 }}>{r.phase}</td>
                      <td style={{ ...S.td, textAlign:"right", color:"#999", fontFamily:"monospace" }}>
                        {r.editable ? (
                          <input
                            type="number"
                            min={1}
                            value={p[r.manualKey] === "" ? r.days : p[r.manualKey]}
                            onChange={e => updTimelineDays(r.manualKey, e.target.value)}
                            style={{ ...S.mini, width:58, textAlign:"right" }}
                            title={`自动=${r.autoDays}天`}
                          />
                        ) : (
                          r.days
                        )}
                      </td>
                      <td style={{ ...S.td, textAlign:"center", fontFamily:"monospace", fontSize:12 }}>{r.start}</td>
                      <td style={{ ...S.td, textAlign:"center", fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{r.end}</td>
                      <td style={{ ...S.td, fontSize:11, color:"#aaa" }}>{r.note}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{ marginTop:14, padding:"12px 16px", background:"#f7f4ef", borderRadius:6, fontSize:11, color:"#888", lineHeight:2.2 }}>
                <strong style={{ color:"#7a6e5f" }}>周末规则：</strong>甄别问卷确认、招募、主问卷确认、Topline、Report 按工作日计算，自动跳过中国周六、周日。<br/>
                <strong style={{ color:"#7a6e5f" }}>留置：</strong>{design.hutWeeks}周 × 7 + 快递 {p.courierDays} 天 = {leaveDays} 天，按自然日计算且固定不可压缩。<br/>
                <strong style={{ color:"#7a6e5f" }}>编辑规则：</strong>仅在“当前需要加急”或“用户主动开启手动调整”时，才允许修改除“派发和留置”外的阶段天数。
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function F({ label, children, last }) {
  return (
    <div style={{ marginBottom: last ? 0 : 14 }}>
      <label style={{ fontSize:10, letterSpacing:"0.08em", textTransform:"uppercase", color:"#888", marginBottom:5, display:"block" }}>{label}</label>
      {children}
    </div>
  );
}

const badgeBlue = { fontSize:9, background:"#e8f0fe", color:"#3a5a9a", borderRadius:3, padding:"1px 5px", fontFamily:"monospace" };
const badgeAmber = { fontSize:9, background:"#fef3e2", color:"#8a5a0a", borderRadius:3, padding:"1px 5px" };
const badgeGreen = { fontSize:9, background:"#e8f6ee", color:"#1f7a4d", borderRadius:3, padding:"1px 5px" };

const ST = {
  card:{ background:"#fff", borderRadius:8, padding:"20px 22px", marginBottom:16, boxShadow:"0 1px 4px rgba(0,0,0,0.06)" },
  cardTitle:{ fontSize:11, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", color:"#7a6e5f", marginBottom:14, paddingBottom:10, borderBottom:"1px solid #ece8e2" },
  table:{ width:"100%", borderCollapse:"collapse", fontSize:13 },
  th:{ background:"#2c2825", color:"#f5f2ed", padding:"9px 11px", textAlign:"left", fontWeight:600, letterSpacing:"0.04em", fontSize:11 },
  td:{ padding:"8px 11px", borderBottom:"1px solid #ede9e3", verticalAlign:"middle" },
  tabOn:{ background:"#2c2825", color:"#f5f2ed", border:"none", borderRadius:"6px 6px 0 0", padding:"9px 18px", fontFamily:"inherit", fontSize:13, fontWeight:700, cursor:"pointer" },
  tabOff:{ background:"#e8e3dc", color:"#7a6e5f", border:"none", borderRadius:"6px 6px 0 0", padding:"9px 18px", fontFamily:"inherit", fontSize:13, cursor:"pointer" },
  inlineInput:{ border:"1px solid #e0dbd3", borderRadius:4, padding:"4px 8px", fontSize:13, fontFamily:"inherit", width:"100%", outline:"none", background:"#faf8f5" },
  mini:{ border:"1px solid #e0dbd3", borderRadius:4, padding:"3px 6px", fontSize:12, fontFamily:"monospace", outline:"none", background:"#faf8f5" },
  delBtn:{ background:"none", border:"none", cursor:"pointer", color:"#c0392b", fontSize:18, fontWeight:700, padding:"0 4px", lineHeight:1 },
  addBtn:{ background:"#2c2825", color:"#f5f2ed", border:"none", borderRadius:4, padding:"3px 10px", fontSize:13, cursor:"pointer", fontWeight:700 },
  btnDark:{ background:"#2c2825", color:"#f5f2ed", border:"none", borderRadius:5, padding:"7px 14px", fontFamily:"inherit", fontSize:12, cursor:"pointer", fontWeight:600 },
  btnGhost:{ background:"none", border:"1px solid #c4bfb8", color:"#7a6e5f", borderRadius:5, padding:"7px 14px", fontFamily:"inherit", fontSize:12, cursor:"pointer" },
  helpBox:{ background:"#faf8f5", border:"1px solid #ece8e2", borderRadius:6, padding:"8px 10px", marginTop:10, color:"#7a6e5f", fontSize:11, lineHeight:1.8 }
};

const css = `
  *{box-sizing:border-box;margin:0;padding:0;}
  .inp{width:100%;padding:8px 12px;border:1px solid #d4cfc8;border-radius:4px;background:#fff;font-size:13px;color:#1a1a1a;outline:none;font-family:inherit;}
  .inp:focus{border-color:#7a6e5f;}
  .inp:disabled{cursor:not-allowed;}
  .export-btn{background:#e8c99a;color:#2c2825;border:none;border-radius:6px;padding:10px 18px;font-family:inherit;font-weight:700;font-size:13px;cursor:pointer;}
  .export-btn:hover{background:#d4b580;}
`;
