import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function skipWE(d) { const r = new Date(d); while (r.getDay() === 0 || r.getDay() === 6) r.setDate(r.getDate() + 1); return r; }
function skipAdd(d, n) { return skipWE(addDays(d, n)); }
function fmt(d) { if (!d) return "—"; return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`; }
function fmtRMB(n) { return `¥${Number(n || 0).toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; }
function clampInt(v, min = 0) { const n = Number(v); if (!Number.isFinite(n)) return min; return Math.max(min, Math.round(n)); }
function cleanNum(v) { if (v === "" || v == null) return null; const n = Number(v); return Number.isFinite(n) ? n : null; }

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
  "a1","a2","a3","a4","d1","d2","d3","d4"
]);

const INIT_CATALOG = [
  { id:"q1",  cat:"定量·HUT留置", name:"留置测试（1周·1份问卷）",              unitPrice:210,  calcType:"ppp", notes:"标准HUT单产品",  hutWeeks:1, hutQuestionnaires:1, hutProducts:1, locked:true },
  { id:"q2",  cat:"定量·HUT留置", name:"顺序型留置（2产品·每产品1周·2份问卷）", unitPrice:300,  calcType:"ppp", notes:"顺序型",         hutWeeks:2, hutQuestionnaires:2, hutProducts:2, locked:true },
  { id:"q3",  cat:"定量·HUT留置", name:"宣称支持留置（1周·1份问卷）",          unitPrice:200,  calcType:"ppp", notes:"",              hutWeeks:1, hutQuestionnaires:1, hutProducts:1, locked:true },
  { id:"q4",  cat:"定量·HUT留置", name:"宣称支持留置（2周·1份问卷）",          unitPrice:230,  calcType:"ppp", notes:"",              hutWeeks:2, hutQuestionnaires:1, hutProducts:1, locked:true },
  { id:"q5",  cat:"定量·HUT留置", name:"宣称支持留置（1周·2份问卷）",          unitPrice:260,  calcType:"ppp", notes:"",              hutWeeks:1, hutQuestionnaires:2, hutProducts:1, locked:true },
  { id:"q6",  cat:"定量·HUT留置", name:"额外：每增加1周使用",                  unitPrice:25,   calcType:"ppp", notes:"叠加在基础款上", isExtraWeek:true, locked:true },
  { id:"q7",  cat:"定量·HUT留置", name:"额外：每增加1份回访问卷",              unitPrice:50,   calcType:"ppp", notes:"",              isExtraQuestionnaire:true, locked:true },
  { id:"q8",  cat:"定量·HUT留置", name:"额外：每增加1个留置产品",              unitPrice:20,   calcType:"ppp", notes:"",              isExtraProduct:true, locked:true },
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

const CALC_LABELS = { ppp:"单产品N × 产品数", pp:"× 总样本量", prod:"× 产品数", ses:"× 场次数", fix:"固定金额" };
const CATS = ["交付与通用","定量·HUT留置","定量·CLT定点","定性研究","高阶技术"];
const HUT_BASE_IDS = ["q1", "q2", "q3", "q4", "q5"];

function lineFromCatalog(ci, more = {}) {
  return { lid: ci.id, qtyO: null, priceO: null, mulO: ci.isExtraWeek || ci.isExtraQuestionnaire || ci.isExtraProduct ? 1 : null, ...more };
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
    sessions: 0,
    courierDays: 2,
    difficulty: "normal",
    vatRate: 6,
    startDate: new Date().toISOString().split("T")[0],
    recruitDaysManual: "",
    toplineDaysManual: "",
    reportDaysManual: "",
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
      hutWeeks: clampInt(p.designWeeks, 1),
      hutQuestionnaires: clampInt(p.designQuestionnaires, 1),
      hutProducts: clampInt(p.designProducts, 1),
      baseWeeks: base.hutWeeks || 1,
      baseQuestionnaires: base.hutQuestionnaires || 1,
      baseProducts: base.hutProducts || 1,
    };
  }, [selectedBase, p.designWeeks, p.designQuestionnaires, p.designProducts]);

  const totalN = p.totalNAuto ? clampInt(p.perProductN, 1) * design.hutProducts : clampInt(p.totalNManual, 1);
  const recruitDaysAuto = calcRecruitDays(totalN, p.difficulty);
  const recruitDays = cleanNum(p.recruitDaysManual) ?? recruitDaysAuto;
  const leaveDays = design.hutWeeks * 7 + clampInt(p.courierDays, 0);
  const toplineDaysAuto = calcToplineDays(design.hutProducts);
  const reportDaysAuto = calcReportDays(design.hutProducts);
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
    const prodDiff = Math.max(0, design.hutProducts - design.baseProducts);

    if ((weekExtra?.mulO ?? (weekExtra ? 1 : 0)) !== weekDiff) issues.push("“额外增加周数”与当前研究设计不一致，建议点击“同步报价项”。");
    if ((questExtra?.mulO ?? (questExtra ? 1 : 0)) !== questDiff) issues.push("“额外增加问卷”与当前研究设计不一致，建议点击“同步报价项”。");
    if ((prodExtra?.mulO ?? (prodExtra ? 1 : 0)) !== prodDiff) issues.push("“额外增加产品”与当前研究设计不一致，建议点击“同步报价项”。");
    if (!lines.find(l => l.lid === design.baseId)) issues.push("当前研究设计对应的 HUT 基础项尚未加入报价，建议点击“同步报价项”。");

    return issues;
  }, [lines, design, catalog]);

  const autoQty = (ci) => {
    switch (ci.calcType) {
      case "ppp": return clampInt(p.perProductN, 1) * design.hutProducts;
      case "pp": return totalN;
      case "prod": return design.hutProducts;
      case "ses": return hasQual ? clampInt(p.sessions, 0) : 0;
      default: return 1;
    }
  };

  const lineRows = useMemo(() => lines.map(l => {
    const ci = getCi(l.lid);
    if (!ci) return null;
    const qty = l.qtyO ?? autoQty(ci);
    const price = l.priceO ?? ci.unitPrice;
    const mul = l.mulO ?? 1;
    return { ...l, ci, qty, price, mul, total: qty * price * mul };
  }).filter(Boolean), [lines, catalog, p, design, totalN, hasQual]);

  const sub = lineRows.reduce((s, r) => s + r.total, 0);
  const gross = +(sub * (1 + clampInt(p.vatRate, 0) / 100)).toFixed(2);

  const syncHUTLines = () => {
    const keep = lines.filter(l => {
      const ci = getCi(l.lid);
      return ci && ci.cat !== "定量·HUT留置";
    });
    const next = [...keep, lineFromCatalog(getCi(design.baseId), { mulO: null })];
    const weekDiff = Math.max(0, design.hutWeeks - design.baseWeeks);
    const questDiff = Math.max(0, design.hutQuestionnaires - design.baseQuestionnaires);
    const prodDiff = Math.max(0, design.hutProducts - design.baseProducts);
    if (weekDiff > 0) next.push(lineFromCatalog(getCi("q6"), { mulO: weekDiff }));
    if (questDiff > 0) next.push(lineFromCatalog(getCi("q7"), { mulO: questDiff }));
    if (prodDiff > 0) next.push(lineFromCatalog(getCi("q8"), { mulO: prodDiff }));
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
      sp("designProducts", ci.hutProducts);
    }
    setLines(ls => [...ls, lineFromCatalog(ci)]);
  };
  const removeLine = (id) => setLines(ls => ls.filter(l => l.lid !== id));
  const updLine = (id, k, v) => {
    const ci = getCi(id);
    setLines(ls => ls.map(l => {
      if (l.lid !== id) return l;
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
    setCatalog(arr => [...arr, { id:`x${_nid}`, cat:"交付与通用", name:"新费用项", unitPrice:0, calcType:"fix", notes:"", locked:false }]);
  };

  const timeline = useMemo(() => {
    const s = skipWE(new Date(p.startDate));
    const screenEnd = skipAdd(s, 3);
    const rEnd = skipAdd(screenEnd, recruitDays);
    const mEnd = skipAdd(rEnd, 3);
    const lEnd = skipAdd(rEnd, leaveDays);
    const tEnd = skipAdd(lEnd, toplineDays);
    const rpEnd = skipAdd(tEnd, reportDays);
    const attr = skipWE(addDays(rEnd, -2));
    return [
      { phase:"项目确认 / PO", days:"/", start:"—", end:fmt(s), note:"项目起始节点" },
      { phase:"甄别问卷确认", days:3, start:fmt(s), end:fmt(screenEnd), note:"与主问卷可并行准备" },
      { phase:`招募（${recruitWeeksStr(totalN, p.difficulty)}）`, days:recruitDays, start:fmt(screenEnd), end:fmt(rEnd), note:`按总N=${totalN}、${DIFF_OPTS.find(d=>d.value===p.difficulty)?.label}估算；支持手动覆盖` },
      { phase:"提供 Attributes", days:"/", start:"—", end:fmt(attr), note:"建议招募结束前2天" },
      { phase:"产品到达", days:"/", start:"—", end:fmt(attr), note:"建议招募结束前2天" },
      { phase:"主问卷确认", days:3, start:fmt(rEnd), end:fmt(mEnd), note:"固定预留" },
      { phase:`派发和留置（含快递${clampInt(p.courierDays, 0)}天）`, days:leaveDays, start:fmt(rEnd), end:fmt(lEnd), note:`${design.hutWeeks}周 × 7天 + 快递${clampInt(p.courierDays, 0)}天` },
      { phase:`Topline（${design.hutProducts}款）`, days:toplineDays, start:fmt(lEnd), end:fmt(tEnd), note:`仅受产品数量影响；当前自动=${toplineDaysAuto}天` },
      { phase:`Report（${design.hutProducts}款）`, days:reportDays, start:fmt(tEnd), end:fmt(rpEnd), note:`仅受产品数量影响；当前自动=${reportDaysAuto}天` },
    ];
  }, [p.startDate, p.difficulty, p.courierDays, totalN, recruitDays, leaveDays, toplineDays, reportDays, design, toplineDaysAuto, reportDaysAuto]);

  const exportXLSX = () => {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet([
      [p.title],
      [`研究设计：${design.baseName}｜产品数:${design.hutProducts}｜单产品N:${p.perProductN}｜总N:${totalN}｜留置:${design.hutWeeks}周｜问卷:${design.hutQuestionnaires}份`],
      [],
      ["费用项目","计算方式","单价(RMB)","数量","倍数","小计(RMB)"],
      ...lineRows.map(r => [r.ci.name, CALC_LABELS[r.ci.calcType], r.price, r.qty, r.mul, r.total]),
      [],
      [`Cost (before VAT ${p.vatRate}%)`, "", "", "", "", sub],
      [`Cost (after VAT ${p.vatRate}%)`, "", "", "", "", gross],
    ]);
    ws1["!cols"] = [{wch:44},{wch:18},{wch:12},{wch:10},{wch:8},{wch:16}];
    XLSX.utils.book_append_sheet(wb, ws1, "报价");

    const ws2 = XLSX.utils.aoa_to_sheet([
      [`${p.title} · 时间表`],
      [],
      ["阶段","时间（天）","开始日期","结束日期","说明"],
      ...timeline.map(r => [r.phase, r.days, r.start, r.end, r.note]),
    ]);
    ws2["!cols"] = [{wch:36},{wch:10},{wch:14},{wch:14},{wch:38}];
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

      <div style={{ display:"grid", gridTemplateColumns:"320px 1fr", gap:18, padding:"18px 28px", maxWidth:1480, margin:"0 auto" }}>
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
              <F label="产品数量">
                <input className="inp" type="number" min={1} value={p.designProducts} onChange={e => sp("designProducts", clampInt(e.target.value, 1))} />
              </F>
              <F label="留置周数">
                <input className="inp" type="number" min={1} value={p.designWeeks} onChange={e => sp("designWeeks", clampInt(e.target.value, 1))} />
              </F>
            </div>
            <F label="问卷份数">
              <input className="inp" type="number" min={1} value={p.designQuestionnaires} onChange={e => sp("designQuestionnaires", clampInt(e.target.value, 1))} />
            </F>
            <div style={{ display:"flex", gap:8, marginTop:8, flexWrap:"wrap" }}>
              <button onClick={syncHUTLines} style={S.btnDark}>同步 HUT 报价项</button>
              <button onClick={() => { sp("designBaseId", "q1"); sp("designProducts", 1); sp("designWeeks", 1); sp("designQuestionnaires", 1); }} style={S.btnGhost}>恢复基础设计</button>
            </div>
            <div style={{ background:"#f0ebe3", borderRadius:6, padding:"12px", marginTop:14, fontSize:12, lineHeight:2 }}>
              {[
                ["当前基础项", design.baseName],
                ["产品数量", `${design.hutProducts} 款`],
                ["留置周数", `${design.hutWeeks} 周`],
                ["问卷份数", `${design.hutQuestionnaires} 份`],
              ].map(([k, v]) => (
                <div key={k} style={{ display:"flex", justifyContent:"space-between", borderBottom:"1px solid #e4dfd7", padding:"3px 0" }}>
                  <span style={{ color:"#7a6e5f" }}>{k}</span>
                  <span style={{ fontFamily:"monospace", fontWeight:700, color:"#2c2825", marginLeft:12, textAlign:"right" }}>{v}</span>
                </div>
              ))}
            </div>
          </div>

          <div style={S.card}>
            <div style={S.cardTitle}>手动参数</div>
            <div style={{ background:"#f7f4ef", borderRadius:6, padding:"12px 12px 8px", marginBottom:14, border:"1px solid #ede9e2" }}>
              <div style={{ fontSize:10, letterSpacing:"0.08em", textTransform:"uppercase", color:"#7a6e5f", marginBottom:10, fontWeight:700 }}>样本量设置</div>
              <F label="单产品样本量 n（per product）">
                <input className="inp" type="number" min={1} value={p.perProductN} onChange={e => {
                  const v = clampInt(e.target.value, 1);
                  sp("perProductN", v);
                  if (p.totalNAuto) sp("totalNManual", v * design.hutProducts);
                }} />
              </F>
              <label style={{ display:"flex", alignItems:"center", gap:8, marginBottom:10, cursor:"pointer", fontSize:12, color:"#888" }}>
                <input
                  type="checkbox"
                  checked={p.totalNAuto}
                  onChange={e => {
                    sp("totalNAuto", e.target.checked);
                    if (e.target.checked) sp("totalNManual", clampInt(p.perProductN, 1) * design.hutProducts);
                  }}
                  style={{ width:15, height:15, accentColor:"#7a6e5f" }}
                />
                总N 自动 = 单产品N × 产品数
              </label>
              <F label={`总样本量 N${p.totalNAuto ? " (自动)" : ""}`} last>
                <input
                  className="inp"
                  type="number"
                  min={1}
                  value={totalN}
                  disabled={p.totalNAuto}
                  onChange={e => sp("totalNManual", clampInt(e.target.value, 1))}
                  style={p.totalNAuto ? { background:"#ede9e2", color:"#aaa" } : {}}
                />
              </F>
            </div>

            <F label="定性场次（组/场）">
              <input className="inp" type="number" min={0} value={hasQual ? p.sessions : 0} disabled={!hasQual} onChange={e => sp("sessions", clampInt(e.target.value, 0))} style={!hasQual ? { background:"#ede9e2", color:"#aaa" } : {}} />
            </F>
            <F label="快递时间（天，默认2天）">
              <input className="inp" type="number" min={0} max={14} value={p.courierDays} onChange={e => sp("courierDays", clampInt(e.target.value, 0))} />
            </F>
            <F label="招募难度">
              <select className="inp" value={p.difficulty} onChange={e => sp("difficulty", e.target.value)}>
                {DIFF_OPTS.map(d => <option key={d.value} value={d.value}>{d.label}（系数×{d.factor}）</option>)}
              </select>
            </F>
            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:12 }}>
              <F label={`招募天数（自动 ${recruitDaysAuto}）`}>
                <input className="inp" type="number" min={1} placeholder={`${recruitDaysAuto}`} value={p.recruitDaysManual} onChange={e => sp("recruitDaysManual", e.target.value)} />
              </F>
              <F label={`Topline 天数（自动 ${toplineDaysAuto}）`}>
                <input className="inp" type="number" min={1} placeholder={`${toplineDaysAuto}`} value={p.toplineDaysManual} onChange={e => sp("toplineDaysManual", e.target.value)} />
              </F>
            </div>
            <F label={`Report 天数（自动 ${reportDaysAuto}）`}>
              <input className="inp" type="number" min={1} placeholder={`${reportDaysAuto}`} value={p.reportDaysManual} onChange={e => sp("reportDaysManual", e.target.value)} />
            </F>
            <F label="VAT 税率 (%)">
              <input className="inp" type="number" min={0} max={20} value={p.vatRate} onChange={e => sp("vatRate", clampInt(e.target.value, 0))} />
            </F>
            <F label="项目确认日期" last>
              <input className="inp" type="date" value={p.startDate} onChange={e => sp("startDate", e.target.value)} />
            </F>
          </div>

          {validations.length > 0 && (
            <div style={{ background:"#fff7ed", border:"1px solid #f1d4a9", borderRadius:8, padding:"14px 16px", color:"#8a5a0a", fontSize:12, lineHeight:1.9, marginBottom:12 }}>
              <div style={{ fontWeight:700, marginBottom:6 }}>校验提醒</div>
              {validations.map((msg, i) => <div key={i}>• {msg}</div>)}
            </div>
          )}

          <div style={{ background:"#2c2825", borderRadius:8, padding:"16px 18px", marginBottom:12 }}>
            <div style={{ fontSize:10, fontWeight:700, letterSpacing:"0.08em", textTransform:"uppercase", color:"#9a8e80", marginBottom:10 }}>时间线预览（自动计算）</div>
            {[
              ["招募周期", `${recruitDays}天 / ${recruitWeeksStr(totalN, p.difficulty)}`],
              ["留置总周期", `${leaveDays}天（${design.hutWeeks}周×7+快递${p.courierDays}天）`],
              ["Topline", `${toplineDays}天（仅按${design.hutProducts}款计算）`],
              ["Report", `${reportDays}天（仅按${design.hutProducts}款计算）`],
              ["HUT计价数量", `${p.perProductN}×${design.hutProducts}=${clampInt(p.perProductN, 1) * design.hutProducts}`],
            ].map(([k, v]) => (
              <div key={k} style={{ display:"flex", justifyContent:"space-between", fontSize:11, padding:"5px 0", borderBottom:"1px solid #3d3530" }}>
                <span style={{ color:"#9a8e80" }}>{k}</span>
                <span style={{ fontFamily:"monospace", fontWeight:600, color:"#e8c99a", fontSize:11, marginLeft:12, textAlign:"right" }}>{v}</span>
              </div>
            ))}
          </div>
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
                      <th style={{ ...S.th, width:"30%" }}>费用项目</th>
                      <th style={{ ...S.th, width:"15%", fontSize:10 }}>计算基准</th>
                      <th style={{ ...S.th, width:"12%", textAlign:"right" }}>单价 (¥)</th>
                      <th style={{ ...S.th, width:"9%", textAlign:"right" }}>数量</th>
                      <th style={{ ...S.th, width:"8%", textAlign:"center" }}>倍数 ×</th>
                      <th style={{ ...S.th, width:"14%", textAlign:"right" }}>小计 (¥)</th>
                      <th style={{ ...S.th, width:"6%" }}></th>
                    </tr>
                  </thead>
                  <tbody>
                    {lineRows.map((r, i) => {
                      const isBase = r.ci.hutWeeks != null && !r.ci.isExtraWeek && !r.ci.isExtraQuestionnaire && !r.ci.isExtraProduct;
                      return (
                        <tr key={r.lid} style={i % 2 === 1 ? { background:"#faf8f5" } : {}}>
                          <td style={S.td}>
                            <div>{r.ci.name}</div>
                            <div style={{ marginTop:3, display:"flex", gap:4, flexWrap:"wrap" }}>
                              {isBase && <span style={badgeBlue}>{r.ci.hutWeeks}周·{r.ci.hutQuestionnaires}卷·{r.ci.hutProducts}款</span>}
                              {r.ci.isExtraWeek && <span style={badgeAmber}>+{r.mulO ?? 1}周</span>}
                              {r.ci.isExtraQuestionnaire && <span style={badgeAmber}>+{r.mulO ?? 1}份问卷</span>}
                              {r.ci.isExtraProduct && <span style={badgeAmber}>+{r.mulO ?? 1}款产品</span>}
                            </div>
                          </td>
                          <td style={{ ...S.td, fontSize:10, color:"#bbb", lineHeight:1.4 }}>{CALC_LABELS[r.ci.calcType]}</td>
                          <td style={{ ...S.td, textAlign:"right" }}>
                            <input type="number" placeholder={r.ci.unitPrice} value={r.priceO ?? ""} onChange={e => updLine(r.lid, "priceO", e.target.value)} style={{ ...S.mini, width:80, textAlign:"right" }} title="留空=目录价" />
                          </td>
                          <td style={{ ...S.td, textAlign:"right" }}>
                            <input type="number" placeholder={autoQty(r.ci)} value={r.qtyO ?? ""} onChange={e => updLine(r.lid, "qtyO", e.target.value)} style={{ ...S.mini, width:58, textAlign:"right" }} title="留空=自动" />
                          </td>
                          <td style={{ ...S.td, textAlign:"center" }}>
                            <input type="number" min={1} placeholder={isBase ? "—" : "1"} value={isBase ? "" : (r.mulO ?? "")} disabled={isBase} onChange={e => updLine(r.lid, "mulO", e.target.value)} style={{ ...S.mini, width:44, textAlign:"center", ...(isBase ? { background:"#ede9e2", color:"#aaa" } : {}) }} title={isBase ? "基础项不使用倍数" : "倍数，留空=1"} />
                          </td>
                          <td style={{ ...S.td, textAlign:"right", fontFamily:"monospace", fontWeight:600 }}>{fmtRMB(r.total)}</td>
                          <td style={{ ...S.td, textAlign:"center" }}><button onClick={() => removeLine(r.lid)} style={S.delBtn}>×</button></td>
                        </tr>
                      );
                    })}
                    {!lineRows.length && <tr><td colSpan={7} style={{ ...S.td, color:"#ccc", textAlign:"center", padding:28 }}>暂无费用项，可先同步 HUT 项，再去「费用目录」添加其他费用</td></tr>}
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
                <strong>HUT计价人次（单产品N×产品数）：</strong>{p.perProductN} × {design.hutProducts} = <strong>{clampInt(p.perProductN, 1) * design.hutProducts}</strong>
                &emsp;|&emsp;<strong>CLT/定性总N：</strong>{totalN} 人
                &emsp;|&emsp;Topline / Report 仅按产品数计算，不再受额外问卷影响
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
                                {ci.hutWeeks != null && <div style={{ fontSize:9, color:"#3a5a9a", marginTop:3 }}>📐 {ci.hutWeeks}周 · {ci.hutQuestionnaires}份问卷 · {ci.hutProducts}款</div>}
                                {ci.isExtraWeek && <div style={{ fontSize:9, color:"#8a5a0a", marginTop:3 }}>⊕ 每倍数 +1周留置</div>}
                                {ci.isExtraQuestionnaire && <div style={{ fontSize:9, color:"#8a5a0a", marginTop:3 }}>⊕ 每倍数 +1份问卷（仅影响报价，不影响Topline/Report）</div>}
                                {ci.isExtraProduct && <div style={{ fontSize:9, color:"#8a5a0a", marginTop:3 }}>⊕ 每倍数 +1个产品</div>}
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
              <div style={S.cardTitle}>项目时间表 · 关键节点周末自动顺延至下周一</div>
              <table style={S.table}>
                <thead>
                  <tr>
                    <th style={{ ...S.th, width:"30%" }}>阶段</th>
                    <th style={{ ...S.th, textAlign:"right", width:"8%" }}>天数</th>
                    <th style={{ ...S.th, textAlign:"center", width:"17%" }}>开始</th>
                    <th style={{ ...S.th, textAlign:"center", width:"17%" }}>结束</th>
                    <th style={{ ...S.th, width:"28%" }}>计算说明</th>
                  </tr>
                </thead>
                <tbody>
                  {timeline.map((r, i) => (
                    <tr key={i} style={i % 2 === 1 ? { background:"#faf8f5" } : {}}>
                      <td style={{ ...S.td, fontWeight:500 }}>{r.phase}</td>
                      <td style={{ ...S.td, textAlign:"right", color:"#999", fontFamily:"monospace" }}>{r.days}</td>
                      <td style={{ ...S.td, textAlign:"center", fontFamily:"monospace", fontSize:12 }}>{r.start}</td>
                      <td style={{ ...S.td, textAlign:"center", fontFamily:"monospace", fontSize:12, fontWeight:700 }}>{r.end}</td>
                      <td style={{ ...S.td, fontSize:11, color:"#aaa" }}>{r.note}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{ marginTop:14, padding:"12px 16px", background:"#f7f4ef", borderRadius:6, fontSize:11, color:"#888", lineHeight:2.2 }}>
                <strong style={{ color:"#7a6e5f" }}>招募周期：</strong>按样本量分段估算，并结合招募难度系数，上限不超过 28 天。当前自动：{recruitDaysAuto} 天<br/>
                <strong style={{ color:"#7a6e5f" }}>Topline：</strong>7天基础；第3/5/7...款各 +3.5 天，不再因额外问卷增加时长。当前自动：{toplineDaysAuto} 天
                &emsp;<strong style={{ color:"#7a6e5f" }}>Report：</strong>7天基础；每 +1 款增 3.5 天。当前自动：{reportDaysAuto} 天<br/>
                <strong style={{ color:"#7a6e5f" }}>留置：</strong>{design.hutWeeks}周 × 7 + 快递 {p.courierDays} 天 = {leaveDays} 天
                &emsp;<strong style={{ color:"#7a6e5f" }}>当前设计：</strong>{design.hutProducts}款 · {design.hutWeeks}周 · {design.hutQuestionnaires}份问卷
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
};

const css = `
  *{box-sizing:border-box;margin:0;padding:0;}
  .inp{width:100%;padding:8px 12px;border:1px solid #d4cfc8;border-radius:4px;background:#fff;font-size:13px;color:#1a1a1a;outline:none;font-family:inherit;}
  .inp:focus{border-color:#7a6e5f;}
  .inp:disabled{cursor:not-allowed;}
  .export-btn{background:#e8c99a;color:#2c2825;border:none;border-radius:6px;padding:10px 18px;font-family:inherit;font-weight:700;font-size:13px;cursor:pointer;}
  .export-btn:hover{background:#d4b580;}
`;
