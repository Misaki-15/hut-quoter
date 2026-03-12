import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }
function skipWE(d) { const r = new Date(d); while (r.getDay() === 0 || r.getDay() === 6) r.setDate(r.getDate() + 1); return r; }
function skipAdd(d, n) { return skipWE(addDays(d, n)); }
function fmt(d) { if (!d) return "—"; return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}`; }
function fmtRMB(n) { return `¥${Number(n).toLocaleString("zh-CN",{minimumFractionDigits:2,maximumFractionDigits:2})}`; }
function r2(n) { return Math.round(n * 2) / 2; }

const DIFF_FACTOR = { easy: 1.0, normal: 1.4, hard: 2.0 };
function calcRecruitDays(totalN, diff) {
  const weeks = Math.max(0.5, 0.5 * (totalN / 30) * (DIFF_FACTOR[diff] || 1));
  return Math.round(weeks * 7);
}
function recruitWeeksStr(totalN, diff) {
  const weeks = Math.max(0.5, 0.5 * (totalN / 30) * (DIFF_FACTOR[diff] || 1));
  const lo = r2(weeks), hi = r2(weeks + 0.5);
  return `约 ${lo}–${hi} 周`;
}
function calcToplineDays(products) { return 7 + Math.max(0, Math.ceil((products - 2) / 2)) * 3.5; }
function calcReportDays(products) { return 7 + (products - 1) * 3.5; }

let _nid = 500;

const INIT_CATALOG = [
  { id:"q1",  cat:"定量·HUT留置", name:"留置测试（1周·1份问卷）",              unitPrice:210,  calcType:"ppp", notes:"标准HUT单产品" },
  { id:"q2",  cat:"定量·HUT留置", name:"顺序型留置（2产品·每产品1周·2份问卷）", unitPrice:300,  calcType:"ppp", notes:"顺序型" },
  { id:"q3",  cat:"定量·HUT留置", name:"宣称支持留置（1周·1份问卷）",          unitPrice:200,  calcType:"ppp", notes:"" },
  { id:"q4",  cat:"定量·HUT留置", name:"宣称支持留置（2周·1份问卷）",          unitPrice:230,  calcType:"ppp", notes:"" },
  { id:"q5",  cat:"定量·HUT留置", name:"宣称支持留置（1周·2份问卷）",          unitPrice:260,  calcType:"ppp", notes:"" },
  { id:"q6",  cat:"定量·HUT留置", name:"额外：每增加1周使用",                  unitPrice:25,   calcType:"ppp", notes:"叠加在基础款上" },
  { id:"q7",  cat:"定量·HUT留置", name:"额外：每增加1份回访问卷",              unitPrice:50,   calcType:"ppp", notes:"" },
  { id:"q8",  cat:"定量·HUT留置", name:"额外：每增加1个留置产品",              unitPrice:20,   calcType:"ppp", notes:"" },
  { id:"q9",  cat:"定量·CLT定点", name:"定点测试CLT（30分钟内）",              unitPrice:280,  calcType:"pp",  notes:"" },
  { id:"q10", cat:"定量·CLT定点", name:"定点测试CLT（60分钟内）",              unitPrice:350,  calcType:"pp",  notes:"" },
  { id:"ql1", cat:"定性研究",     name:"一对一深访（1小时·线上）",             unitPrice:4167, calcType:"ses", notes:"差旅不含" },
  { id:"ql2", cat:"定性研究",     name:"一对一深访（1小时·线下）",             unitPrice:4333, calcType:"ses", notes:"差旅另计" },
  { id:"ql3", cat:"定性研究",     name:"家访（1.5小时）",                     unitPrice:5500, calcType:"ses", notes:"差旅另计" },
  { id:"ql4", cat:"定性研究",     name:"座谈会FGD（2小时·线上）",             unitPrice:16500,calcType:"ses", notes:"每组6-8人" },
  { id:"ql5", cat:"定性研究",     name:"座谈会FGD（2小时·线下）",             unitPrice:18000,calcType:"ses", notes:"差旅另计" },
  { id:"ql6", cat:"定性研究",     name:"定性额外：每增加1周产品使用",           unitPrice:25,   calcType:"pp",  notes:"" },
  { id:"ql7", cat:"定性研究",     name:"定性额外：每增加1份回访问卷",           unitPrice:50,   calcType:"pp",  notes:"" },
  { id:"ql8", cat:"定性研究",     name:"定性额外：每增加1个留置产品",           unitPrice:20,   calcType:"pp",  notes:"" },
  { id:"a1",  cat:"高阶技术",     name:"眼动追踪（额外加收）",                 unitPrice:150,  calcType:"pp",  notes:"" },
  { id:"a2",  cat:"高阶技术",     name:"微表情识别（额外加收）",               unitPrice:150,  calcType:"pp",  notes:"" },
  { id:"a3",  cat:"高阶技术",     name:"脑电监测（额外加收）",                 unitPrice:500,  calcType:"pp",  notes:"" },
  { id:"a4",  cat:"高阶技术",     name:"数据建模（感官+消费者）",              unitPrice:50000,calcType:"fix", notes:"每个模型" },
  { id:"d1",  cat:"交付与通用",   name:"Excel版数据简报",                    unitPrice:1600, calcType:"prod",notes:"" },
  { id:"d2",  cat:"交付与通用",   name:"完整报告",                           unitPrice:4000, calcType:"prod",notes:"" },
  { id:"d3",  cat:"交付与通用",   name:"项目管理费",                         unitPrice:1000, calcType:"prod",notes:"含样本小组建立/维护" },
  { id:"d4",  cat:"交付与通用",   name:"盲包",                              unitPrice:20,   calcType:"ppp", notes:"需灏图提供盲包时" },
];

const CALC_LABELS = { ppp:"单产品N × 产品数", pp:"× 总样本量", prod:"× 产品数", ses:"× 场次数", fix:"固定金额" };
const DIFF_OPTS = [
  { value:"easy",   label:"普通招募", factor:1.0 },
  { value:"normal", label:"较难招募", factor:1.4 },
  { value:"hard",   label:"难招募",   factor:2.0 },
];
const CATS = ["定量·HUT留置","定量·CLT定点","定性研究","高阶技术","交付与通用"];

export default function App() {
  const [tab, setTab] = useState("quote");
  const [catalog, setCatalog] = useState(INIT_CATALOG);
  const [p, setP_] = useState({
    title: "HUT 留置测试",
    products: 1, perProductN: 45, totalNAuto: true, totalNManual: 45,
    sessions: 2, weeksDuration: 1, courierDays: 2,
    difficulty: "normal", vatRate: 6,
    startDate: new Date().toISOString().split("T")[0],
  });
  const sp = (k, v) => setP_(c => ({ ...c, [k]: v }));

  const totalN = p.totalNAuto ? p.perProductN * p.products : p.totalNManual;
  const recruitDays = calcRecruitDays(totalN, p.difficulty);
  const leaveDays   = p.weeksDuration * 7 + p.courierDays;
  const toplineDays = Math.round(calcToplineDays(p.products));
  const reportDays  = Math.round(calcReportDays(p.products));

  const [lines, setLines] = useState([
    { lid:"q1", qtyO:null, priceO:null },
    { lid:"d3", qtyO:null, priceO:null },
    { lid:"d1", qtyO:null, priceO:null },
    { lid:"d2", qtyO:null, priceO:null },
  ]);

  const autoQty = (ci) => {
    switch (ci.calcType) {
      case "ppp":  return p.perProductN * p.products;
      case "pp":   return totalN;
      case "prod": return p.products;
      case "ses":  return p.sessions;
      default:     return 1;
    }
  };

  const lineRows = useMemo(() => lines.map(l => {
    const ci = catalog.find(c => c.id === l.lid);
    if (!ci) return null;
    const qty   = l.qtyO   ?? autoQty(ci);
    const price = l.priceO ?? ci.unitPrice;
    return { ...l, ci, qty, price, total: qty * price };
  }).filter(Boolean), [lines, catalog, p, totalN]);

  const sub   = lineRows.reduce((s, r) => s + r.total, 0);
  const gross = +(sub * (1 + p.vatRate / 100)).toFixed(2);

  const addLine    = id => { if (lines.find(l => l.lid === id)) return; setLines(ls => [...ls, { lid:id, qtyO:null, priceO:null }]); };
  const removeLine = id => setLines(ls => ls.filter(l => l.lid !== id));
  const updLine    = (id, k, v) => setLines(ls => ls.map(l => l.lid === id ? { ...l, [k]: v === "" ? null : +v } : l));
  const updCat     = (id, k, v) => setCatalog(arr => arr.map(c => c.id === id ? { ...c, [k]: v } : c));
  const delCat     = id => { setCatalog(arr => arr.filter(c => c.id !== id)); setLines(ls => ls.filter(l => l.lid !== id)); };
  const addCat     = () => { _nid++; setCatalog(arr => [...arr, { id:`x${_nid}`, cat:"交付与通用", name:"新费用项", unitPrice:0, calcType:"fix", notes:"" }]); };

  const timeline = useMemo(() => {
    const s    = skipWE(new Date(p.startDate));
    const rEnd = skipAdd(s, recruitDays);
    const mEnd = skipAdd(rEnd, 3);
    const lEnd = skipAdd(rEnd, leaveDays);
    const tEnd = skipAdd(lEnd, toplineDays);
    const rpEnd = skipAdd(tEnd, reportDays);
    const attr = skipWE(addDays(rEnd, -2));
    return [
      { phase:"项目确认 / PO",   days:"/", start:"—", end:fmt(s), note:"" },
      { phase:"甄别问卷确认",    days:3,   start:fmt(s), end:fmt(skipAdd(s,3)), note:"与主问卷同等周期" },
      { phase:`招募（${recruitWeeksStr(totalN, p.difficulty)}）`, days:recruitDays, start:fmt(s), end:fmt(rEnd), note:`总N=${totalN}，${DIFF_OPTS.find(d=>d.value===p.difficulty)?.label}×${DIFF_FACTOR[p.difficulty]}` },
      { phase:"提供 Attributes", days:"/", start:"—", end:fmt(attr), note:"招募结束前2天" },
      { phase:"产品到达",        days:"/", start:"—", end:fmt(attr), note:"招募结束前2天" },
      { phase:"主问卷确认",      days:3,   start:fmt(rEnd), end:fmt(mEnd), note:"" },
      { phase:`派发和留置（含快递${p.courierDays}天）`, days:leaveDays, start:fmt(rEnd), end:fmt(lEnd), note:`留置${p.weeksDuration}周×7天 + 快递${p.courierDays}天` },
      { phase:`Topline（${p.products}款）`, days:toplineDays, start:fmt(lEnd), end:fmt(tEnd), note:p.products<=2?"1款/2款=1周":`${p.products}款：1周+${Math.ceil((p.products-2)/2)}×3.5天` },
      { phase:`Report（${p.products}款）`,  days:reportDays,  start:fmt(tEnd), end:fmt(rpEnd), note:"1款=1周；每+1款增3.5天" },
    ];
  }, [p, recruitDays, leaveDays, toplineDays, reportDays, totalN]);

  const exportXLSX = () => {
    const wb = XLSX.utils.book_new();

    // ── Sheet 1: 报价 ──
    const ws1 = XLSX.utils.aoa_to_sheet([
      [p.title],
      [`产品数:${p.products}  单产品N:${p.perProductN}  总N:${totalN}  留置:${p.weeksDuration}周+快递${p.courierDays}天  招募:${DIFF_OPTS.find(d=>d.value===p.difficulty)?.label}`],
      [],
      ["费用项目","计算方式","单价(RMB)","数量","小计(RMB)"],
      ...lineRows.map(r=>[r.ci.name, CALC_LABELS[r.ci.calcType], r.price, r.qty, r.total]),
      [],
      [`Cost (before VAT ${p.vatRate}%)`, "", "", "", sub],
      [`Cost (after VAT ${p.vatRate}%)`,  "", "", "", gross],
    ]);
    ws1["!cols"] = [{wch:44},{wch:18},{wch:12},{wch:10},{wch:16}];
    XLSX.utils.book_append_sheet(wb, ws1, "报价");

    // ── Sheet 2: 时间表 ──
    const ws2 = XLSX.utils.aoa_to_sheet([
      [p.title + " · 时间表"],
      [],
      ["阶段","时间（天）","开始日期","结束日期","说明"],
      ...timeline.map(r=>[r.phase, r.days, r.start, r.end, r.note]),
    ]);
    ws2["!cols"] = [{wch:36},{wch:10},{wch:14},{wch:14},{wch:36}];
    XLSX.utils.book_append_sheet(wb, ws2, "时间表");

    XLSX.writeFile(wb, `${p.title}_报价.xlsx`);
  };

  const S = ST;
  return (
    <div style={{fontFamily:"Georgia,serif",background:"#f4f1ec",minHeight:"100vh",color:"#1a1a1a"}}>
      <style>{css}</style>

      <div style={{background:"#2c2825",padding:"18px 28px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,flexWrap:"wrap"}}>
        <div>
          <div style={{fontSize:10,letterSpacing:"0.12em",textTransform:"uppercase",color:"#9a8e80",marginBottom:4}}>HOW-TO 灏图品测 · 智能报价系统</div>
          <input value={p.title} onChange={e=>sp("title",e.target.value)}
            style={{background:"transparent",border:"none",borderBottom:"1px solid #5a504a",color:"#f5f2ed",fontSize:20,fontWeight:700,fontFamily:"inherit",outline:"none",width:360,padding:"2px 0"}} />
        </div>
        <div style={{display:"flex",alignItems:"center",gap:14}}>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:10,color:"#9a8e80",letterSpacing:"0.08em",textTransform:"uppercase"}}>含税总额</div>
            <div style={{fontSize:26,fontWeight:700,fontFamily:"monospace",color:"#e8c99a"}}>{fmtRMB(gross)}</div>
          </div>
          <button className="export-btn" onClick={exportXLSX}>⬇ 导出 Excel</button>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"284px 1fr",gap:18,padding:"18px 28px",maxWidth:1440,margin:"0 auto"}}>

        {/* LEFT */}
        <div>
          <div style={S.card}>
            <div style={S.cardTitle}>项目参数</div>
            <F label="产品数量">
              <select className="inp" value={p.products} onChange={e=>{const v=+e.target.value; sp("products",v); if(p.totalNAuto) sp("totalNManual",p.perProductN*v);}}>
                {[1,2,3,4,5].map(n=><option key={n} value={n}>{n} 款</option>)}
              </select>
            </F>
            <div style={{background:"#f7f4ef",borderRadius:6,padding:"12px 12px 8px",marginBottom:14,border:"1px solid #ede9e2"}}>
              <div style={{fontSize:10,letterSpacing:"0.08em",textTransform:"uppercase",color:"#7a6e5f",marginBottom:10,fontWeight:700}}>样本量设置</div>
              <F label="单产品样本量 n（per product）">
                <input className="inp" type="number" min={1} value={p.perProductN}
                  onChange={e=>{const v=+e.target.value; sp("perProductN",v); if(p.totalNAuto) sp("totalNManual",v*p.products);}} />
              </F>
              <label style={{display:"flex",alignItems:"center",gap:8,marginBottom:10,cursor:"pointer",fontSize:12,color:"#888"}}>
                <input type="checkbox" checked={p.totalNAuto}
                  onChange={e=>{sp("totalNAuto",e.target.checked); if(e.target.checked) sp("totalNManual",p.perProductN*p.products);}}
                  style={{width:15,height:15,accentColor:"#7a6e5f"}} />
                总N 自动 = 单产品N × 产品数
              </label>
              <F label={`总样本量 N${p.totalNAuto?" (自动)":""}`} last>
                <input className="inp" type="number" min={1} value={totalN} disabled={p.totalNAuto}
                  onChange={e=>sp("totalNManual",+e.target.value)}
                  style={p.totalNAuto?{background:"#ede9e2",color:"#aaa"}:{}} />
              </F>
            </div>
            <F label="定性场次（组/场）">
              <input className="inp" type="number" min={1} value={p.sessions} onChange={e=>sp("sessions",+e.target.value)} />
            </F>
            <F label="留置时长">
              <select className="inp" value={p.weeksDuration} onChange={e=>sp("weeksDuration",+e.target.value)}>
                {[1,2,3,4].map(n=><option key={n} value={n}>{n} 周（{n*7}天）</option>)}
              </select>
            </F>
            <F label="快递时间（天，默认2天）">
              <input className="inp" type="number" min={0} max={14} value={p.courierDays} onChange={e=>sp("courierDays",+e.target.value)} />
            </F>
            <F label="招募难度">
              <select className="inp" value={p.difficulty} onChange={e=>sp("difficulty",e.target.value)}>
                {DIFF_OPTS.map(d=><option key={d.value} value={d.value}>{d.label}（系数×{d.factor}）</option>)}
              </select>
            </F>
            <F label="VAT 税率 (%)">
              <input className="inp" type="number" min={0} max={20} value={p.vatRate} onChange={e=>sp("vatRate",+e.target.value)} />
            </F>
            <F label="项目确认日期" last>
              <input className="inp" type="date" value={p.startDate} onChange={e=>sp("startDate",e.target.value)} />
            </F>
          </div>

          <div style={{background:"#2c2825",borderRadius:8,padding:"16px 18px",marginBottom:12}}>
            <div style={{fontSize:10,fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:"#9a8e80",marginBottom:10}}>时间线预览（自动计算）</div>
            {[
              ["招募周期", `${recruitDays}天 / ${recruitWeeksStr(totalN,p.difficulty)}`],
              ["留置总周期", `${leaveDays}天（${p.weeksDuration}周×7+快递${p.courierDays}天）`],
              ["Topline", `${toplineDays}天（${p.products}款）`],
              ["Report",  `${reportDays}天（${p.products}款）`],
              ["HUT计价数量", `${p.perProductN}×${p.products}=${p.perProductN*p.products}`],
            ].map(([k,v])=>(
              <div key={k} style={{display:"flex",justifyContent:"space-between",fontSize:11,padding:"5px 0",borderBottom:"1px solid #3d3530"}}>
                <span style={{color:"#9a8e80"}}>{k}</span>
                <span style={{fontFamily:"monospace",fontWeight:600,color:"#e8c99a",fontSize:11}}>{v}</span>
              </div>
            ))}
          </div>
        </div>

        {/* RIGHT */}
        <div>
          <div style={{display:"flex",gap:3,marginBottom:12}}>
            {[["quote","📋 报价明细"],["catalog","📚 费用目录"],["timeline","📅 时间表"]].map(([id,label])=>(
              <button key={id} onClick={()=>setTab(id)}
                style={tab===id?S.tabOn:S.tabOff}>{label}</button>
            ))}
          </div>

          {/* QUOTE */}
          {tab==="quote"&&(
            <div>
              <div style={S.card}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                  <div style={S.cardTitle}>报价行项目</div>
                  <div style={{fontSize:11,color:"#aaa"}}>前往「费用目录」Tab 添加目录项 →</div>
                </div>
                <table style={S.table}>
                  <thead><tr>
                    <th style={{...S.th,width:"32%"}}>费用项目</th>
                    <th style={{...S.th,width:"16%",fontSize:10}}>计算基准</th>
                    <th style={{...S.th,width:"13%",textAlign:"right"}}>单价 (¥)</th>
                    <th style={{...S.th,width:"9%",textAlign:"right"}}>数量</th>
                    <th style={{...S.th,width:"14%",textAlign:"right"}}>小计 (¥)</th>
                    <th style={{...S.th,width:"6%"}}></th>
                  </tr></thead>
                  <tbody>
                    {lineRows.map((r,i)=>(
                      <tr key={r.lid} style={i%2===1?{background:"#faf8f5"}:{}}>
                        <td style={S.td}>{r.ci.name}</td>
                        <td style={{...S.td,fontSize:10,color:"#bbb",lineHeight:1.4}}>{CALC_LABELS[r.ci.calcType]}</td>
                        <td style={{...S.td,textAlign:"right"}}>
                          <input type="number" placeholder={r.ci.unitPrice} value={r.priceO??""} onChange={e=>updLine(r.lid,"priceO",e.target.value)}
                            style={{...S.mini,width:80,textAlign:"right"}} title="留空=目录价" />
                        </td>
                        <td style={{...S.td,textAlign:"right"}}>
                          <input type="number" placeholder={autoQty(r.ci)} value={r.qtyO??""} onChange={e=>updLine(r.lid,"qtyO",e.target.value)}
                            style={{...S.mini,width:52,textAlign:"right"}} title="留空=自动" />
                        </td>
                        <td style={{...S.td,textAlign:"right",fontFamily:"monospace",fontWeight:600}}>{fmtRMB(r.total)}</td>
                        <td style={{...S.td,textAlign:"center"}}><button onClick={()=>removeLine(r.lid)} style={S.delBtn}>×</button></td>
                      </tr>
                    ))}
                    {!lineRows.length&&<tr><td colSpan={6} style={{...S.td,color:"#ccc",textAlign:"center",padding:28}}>暂无费用项，前往「费用目录」Tab 添加</td></tr>}
                  </tbody>
                </table>
                <div style={{marginTop:12,display:"flex",justifyContent:"flex-end"}}>
                  <table style={{borderCollapse:"collapse",minWidth:320}}>
                    <tbody>
                      <tr style={{background:"#f0ebe3"}}>
                        <td style={{...S.td,fontWeight:700,paddingLeft:16}}>Cost (before VAT {p.vatRate}%)</td>
                        <td style={{...S.td,textAlign:"right",fontFamily:"monospace",fontWeight:700,paddingRight:16}}>{fmtRMB(sub)}</td>
                      </tr>
                      <tr style={{background:"#2c2825"}}>
                        <td style={{...S.td,fontWeight:700,color:"#f5f2ed",paddingLeft:16}}>Cost (after VAT {p.vatRate}%)</td>
                        <td style={{...S.td,textAlign:"right",fontFamily:"monospace",fontWeight:700,color:"#e8c99a",paddingRight:16}}>{fmtRMB(gross)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
              <div style={{background:"#fffbf5",border:"1px solid #ece3d0",borderRadius:8,padding:"11px 15px",fontSize:11,color:"#9a8e80",lineHeight:2.1}}>
                <strong>HUT计价人次（单产品N×产品数）：</strong>{p.perProductN} × {p.products} = <strong>{p.perProductN*p.products}</strong>
                &emsp;|&emsp;<strong>CLT/定性总N：</strong>{totalN} 人
                &emsp;|&emsp;单价/数量留空 = 自动填入；输入则覆盖
              </div>
            </div>
          )}

          {/* CATALOG */}
          {tab==="catalog"&&(
            <div style={S.card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div style={S.cardTitle}>费用目录管理</div>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>setCatalog(INIT_CATALOG)} style={S.btnGhost}>重置默认</button>
                  <button onClick={addCat} style={S.btnDark}>+ 新增目录项</button>
                </div>
              </div>
              {CATS.map(cat=>{
                const items=catalog.filter(c=>c.cat===cat);
                if(!items.length) return null;
                return (
                  <div key={cat} style={{marginBottom:20}}>
                    <div style={{fontSize:11,fontWeight:700,letterSpacing:"0.06em",color:"#7a6e5f",padding:"5px 10px",background:"#f0ebe3",borderRadius:4,marginBottom:6}}>{cat}</div>
                    <table style={S.table}>
                      <thead><tr>
                        <th style={{...S.th,width:"28%"}}>名称</th>
                        <th style={{...S.th,width:"11%"}}>分类</th>
                        <th style={{...S.th,width:"11%",textAlign:"right"}}>单价</th>
                        <th style={{...S.th,width:"16%"}}>计算方式</th>
                        <th style={{...S.th,width:"13%"}}>备注</th>
                        <th style={{...S.th,width:"8%",textAlign:"center"}}>加入报价</th>
                        <th style={{...S.th,width:"5%"}}></th>
                      </tr></thead>
                      <tbody>
                        {items.map((ci,i)=>{
                          const inQ=!!lines.find(l=>l.lid===ci.id);
                          return (
                            <tr key={ci.id} style={i%2===1?{background:"#faf8f5"}:{}}>
                              <td style={S.td}><input value={ci.name} onChange={e=>updCat(ci.id,"name",e.target.value)} style={S.inlineInput}/></td>
                              <td style={S.td}>
                                <select value={ci.cat} onChange={e=>updCat(ci.id,"cat",e.target.value)} style={{...S.inlineInput,fontSize:10,padding:"3px 4px"}}>
                                  {CATS.map(c=><option key={c} value={c}>{c}</option>)}
                                </select>
                              </td>
                              <td style={{...S.td,textAlign:"right"}}>
                                <input type="number" value={ci.unitPrice} onChange={e=>updCat(ci.id,"unitPrice",+e.target.value)} style={{...S.inlineInput,width:72,textAlign:"right",fontFamily:"monospace"}}/>
                              </td>
                              <td style={S.td}>
                                <select value={ci.calcType} onChange={e=>updCat(ci.id,"calcType",e.target.value)} style={{...S.inlineInput,fontSize:10,padding:"3px 4px"}}>
                                  {Object.entries(CALC_LABELS).map(([k,v])=><option key={k} value={k}>{v}</option>)}
                                </select>
                              </td>
                              <td style={S.td}><input value={ci.notes} onChange={e=>updCat(ci.id,"notes",e.target.value)} style={{...S.inlineInput,fontSize:11}} placeholder="—"/></td>
                              <td style={{...S.td,textAlign:"center"}}>
                                {inQ?<span style={{fontSize:11,color:"#27ae60",fontWeight:700}}>✓ 已选</span>
                                   :<button onClick={()=>addLine(ci.id)} style={S.addBtn}>＋</button>}
                              </td>
                              <td style={{...S.td,textAlign:"center"}}><button onClick={()=>delCat(ci.id)} style={S.delBtn}>×</button></td>
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

          {/* TIMELINE */}
          {tab==="timeline"&&(
            <div style={S.card}>
              <div style={S.cardTitle}>项目时间表 · 关键节点周末自动顺延至下周一</div>
              <table style={S.table}>
                <thead><tr>
                  <th style={{...S.th,width:"30%"}}>阶段</th>
                  <th style={{...S.th,textAlign:"right",width:"8%"}}>天数</th>
                  <th style={{...S.th,textAlign:"center",width:"17%"}}>开始</th>
                  <th style={{...S.th,textAlign:"center",width:"17%"}}>结束</th>
                  <th style={{...S.th,width:"28%"}}>计算说明</th>
                </tr></thead>
                <tbody>
                  {timeline.map((r,i)=>(
                    <tr key={i} style={i%2===1?{background:"#faf8f5"}:{}}>
                      <td style={{...S.td,fontWeight:500}}>{r.phase}</td>
                      <td style={{...S.td,textAlign:"right",color:"#999",fontFamily:"monospace"}}>{r.days}</td>
                      <td style={{...S.td,textAlign:"center",fontFamily:"monospace",fontSize:12}}>{r.start}</td>
                      <td style={{...S.td,textAlign:"center",fontFamily:"monospace",fontSize:12,fontWeight:700}}>{r.end}</td>
                      <td style={{...S.td,fontSize:11,color:"#aaa"}}>{r.note}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{marginTop:14,padding:"12px 16px",background:"#f7f4ef",borderRadius:6,fontSize:11,color:"#888",lineHeight:2.2}}>
                <strong style={{color:"#7a6e5f"}}>招募周期：</strong> max(0.5周, 0.5 × (总N÷30) × 难度系数)
                &emsp;当前：{totalN}÷30 × {DIFF_FACTOR[p.difficulty]} × 0.5 = {r2(Math.max(0.5,0.5*(totalN/30)*DIFF_FACTOR[p.difficulty]))} 周<br/>
                <strong style={{color:"#7a6e5f"}}>Topline：</strong>1–2款=7天；第3/5/7...款触发各+3.5天
                &emsp;<strong style={{color:"#7a6e5f"}}>Report：</strong>1款=7天；每+1款增3.5天
                &emsp;<strong style={{color:"#7a6e5f"}}>留置：</strong>{p.weeksDuration}周×7+快递{p.courierDays}天={leaveDays}天
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

const ST = {
  card:{background:"#fff",borderRadius:8,padding:"20px 22px",marginBottom:16,boxShadow:"0 1px 4px rgba(0,0,0,0.06)"},
  cardTitle:{fontSize:11,fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:"#7a6e5f",marginBottom:14,paddingBottom:10,borderBottom:"1px solid #ece8e2"},
  table:{width:"100%",borderCollapse:"collapse",fontSize:13},
  th:{background:"#2c2825",color:"#f5f2ed",padding:"9px 11px",textAlign:"left",fontWeight:600,letterSpacing:"0.04em",fontSize:11},
  td:{padding:"8px 11px",borderBottom:"1px solid #ede9e3",verticalAlign:"middle"},
  tabOn:{background:"#2c2825",color:"#f5f2ed",border:"none",borderRadius:"6px 6px 0 0",padding:"9px 18px",fontFamily:"inherit",fontSize:13,fontWeight:700,cursor:"pointer"},
  tabOff:{background:"#e8e3dc",color:"#7a6e5f",border:"none",borderRadius:"6px 6px 0 0",padding:"9px 18px",fontFamily:"inherit",fontSize:13,cursor:"pointer"},
  inlineInput:{border:"1px solid #e0dbd3",borderRadius:4,padding:"4px 8px",fontSize:13,fontFamily:"inherit",width:"100%",outline:"none",background:"#faf8f5"},
  mini:{border:"1px solid #e0dbd3",borderRadius:4,padding:"3px 6px",fontSize:12,fontFamily:"monospace",outline:"none",background:"#faf8f5"},
  delBtn:{background:"none",border:"none",cursor:"pointer",color:"#c0392b",fontSize:18,fontWeight:700,padding:"0 4px",lineHeight:1},
  addBtn:{background:"#2c2825",color:"#f5f2ed",border:"none",borderRadius:4,padding:"3px 10px",fontSize:13,cursor:"pointer",fontWeight:700},
  btnDark:{background:"#2c2825",color:"#f5f2ed",border:"none",borderRadius:5,padding:"7px 14px",fontFamily:"inherit",fontSize:12,cursor:"pointer",fontWeight:600},
  btnGhost:{background:"none",border:"1px solid #c4bfb8",color:"#7a6e5f",borderRadius:5,padding:"7px 14px",fontFamily:"inherit",fontSize:12,cursor:"pointer"},
};

const css = `
  *{box-sizing:border-box;margin:0;padding:0;}
  .inp{width:100%;padding:8px 12px;border:1px solid #d4cfc8;border-radius:4px;background:#fff;font-size:13px;color:#1a1a1a;outline:none;font-family:inherit;}
  .inp:focus{border-color:#7a6e5f;}
  .inp:disabled{cursor:not-allowed;}
  .export-btn{background:#e8c99a;color:#2c2825;border:none;border-radius:6px;padding:10px 18px;font-family:inherit;font-weight:700;font-size:13px;cursor:pointer;}
  .export-btn:hover{background:#d4b580;}
`;
