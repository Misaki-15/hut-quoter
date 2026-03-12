import { useState, useMemo } from "react";
import * as XLSX from "xlsx";

const INITIAL_CATALOG = [
  { id:"q1",  cat:"定量·HUT留置", name:"留置测试（1周·1份问卷·单产品）",          unitPrice:210,  calcType:"hut",        notes:"标准HUT" },
  { id:"q2",  cat:"定量·HUT留置", name:"留置测试（顺序2产品·每产品1周·2份问卷）",  unitPrice:300,  calcType:"hut",        notes:"顺序型" },
  { id:"q3",  cat:"定量·HUT留置", name:"宣称支持留置（1周·1份问卷）",             unitPrice:200,  calcType:"hut",        notes:"" },
  { id:"q4",  cat:"定量·HUT留置", name:"宣称支持留置（2周·1份问卷）",             unitPrice:230,  calcType:"hut",        notes:"" },
  { id:"q5",  cat:"定量·HUT留置", name:"宣称支持留置（1周·2份问卷）",             unitPrice:260,  calcType:"hut",        notes:"" },
  { id:"q6",  cat:"定量·HUT留置", name:"额外：每增加1周使用",                    unitPrice:25,   calcType:"hut",        notes:"×次数=增加几周", isExtra:true },
  { id:"q7",  cat:"定量·HUT留置", name:"额外：每增加1份回访问卷",                 unitPrice:50,   calcType:"hut",        notes:"×次数=增加几份", isExtra:true },
  { id:"q8",  cat:"定量·HUT留置", name:"额外：每增加1个留置产品",                 unitPrice:20,   calcType:"hut",        notes:"×次数=增加几个", isExtra:true },
  { id:"q9",  cat:"定量·CLT定点", name:"定点测试CLT（1次到访·30分钟内）",          unitPrice:280,  calcType:"per-person", notes:"" },
  { id:"q10", cat:"定量·CLT定点", name:"定点测试CLT（1次到访·60分钟内）",          unitPrice:350,  calcType:"per-person", notes:"" },
  { id:"ql1", cat:"定性研究",     name:"一对一深访（1小时内·线上）",              unitPrice:4167, calcType:"per-session", notes:"含主持·招募·分析" },
  { id:"ql2", cat:"定性研究",     name:"一对一深访（1小时内·线下）",              unitPrice:4333, calcType:"per-session", notes:"差旅另计" },
  { id:"ql3", cat:"定性研究",     name:"家访（1.5小时内）",                     unitPrice:5500, calcType:"per-session", notes:"差旅另计" },
  { id:"ql4", cat:"定性研究",     name:"座谈会FGD（2小时内·线上）",               unitPrice:16500,calcType:"per-session", notes:"每组6-8人" },
  { id:"ql5", cat:"定性研究",     name:"座谈会FGD（2小时内·线下）",               unitPrice:18000,calcType:"per-session", notes:"差旅另计" },
  { id:"ql6", cat:"定性研究",     name:"额外：定性每增加1周产品使用",              unitPrice:25,   calcType:"per-person",  notes:"×次数=增加几周", isExtra:true },
  { id:"ql7", cat:"定性研究",     name:"额外：定性每增加1份回访问卷",              unitPrice:50,   calcType:"per-person",  notes:"×次数=增加几份", isExtra:true },
  { id:"ql8", cat:"定性研究",     name:"额外：定性每增加1个留置产品",              unitPrice:20,   calcType:"per-person",  notes:"×次数=增加几个", isExtra:true },
  { id:"a1",  cat:"高阶技术",     name:"眼动追踪（额外加收）",                   unitPrice:150,  calcType:"per-person", notes:"" },
  { id:"a2",  cat:"高阶技术",     name:"微表情识别（额外加收）",                  unitPrice:150,  calcType:"per-person", notes:"" },
  { id:"a3",  cat:"高阶技术",     name:"脑电监测（额外加收）",                   unitPrice:500,  calcType:"per-person", notes:"" },
  { id:"a4",  cat:"高阶技术",     name:"数据建模（感官+消费者）",                 unitPrice:50000,calcType:"fixed",       notes:"每个模型" },
  { id:"d1",  cat:"交付与通用",   name:"Excel版数据简报",                      unitPrice:1600, calcType:"per-product", notes:"" },
  { id:"d2",  cat:"交付与通用",   name:"完整报告",                            unitPrice:4000, calcType:"per-product", notes:"" },
  { id:"d3",  cat:"交付与通用",   name:"项目管理费",                          unitPrice:1000, calcType:"per-product", notes:"含样本小组维护" },
  { id:"d4",  cat:"交付与通用",   name:"盲包",                               unitPrice:20,   calcType:"hut",         notes:"需灏图盲包时" },
];

const CALC_LABELS = {
  "hut":          "× 单产品样本 × 产品数",
  "per-person":   "× 总样本量",
  "per-product":  "× 产品数",
  "per-session":  "× 场次数",
  "fixed":        "固定金额",
};
const CATS = ["定量·HUT留置","定量·CLT定点","定性研究","高阶技术","交付与通用"];
const DIFF = {
  easy:   { label:"普通招募（×1.0）", priceMult:1.0, timeFactor:0.8 },
  normal: { label:"较难招募（×1.2）", priceMult:1.2, timeFactor:1.0 },
  hard:   { label:"难招募（×1.5）",   priceMult:1.5, timeFactor:1.5 },
};

function skipWE(d){ const r=new Date(d); while(r.getDay()===0||r.getDay()===6) r.setDate(r.getDate()+1); return r; }
function addDays(d,n){ const r=new Date(d); r.setDate(r.getDate()+n); return r; }
function skipAdd(d,n){ return skipWE(addDays(d,n)); }
function fmtD(d){ if(!d) return "—"; return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}`; }
function fmtRMB(n){ return `¥${Number(n).toLocaleString("zh-CN",{minimumFractionDigits:2,maximumFractionDigits:2})}`; }
function wkStr(days){ const w=days/7; return Number.isInteger(w)?`${w}周`:`${+(w.toFixed(1))}周（${days}天）`; }

function calcRecruitWeeks(sampleTotal, difficulty){
  const raw = Math.max(0.5, (sampleTotal/60)*DIFF[difficulty].timeFactor);
  const min = Math.round(raw*2)/2;
  return { min, max:min+0.5, days:Math.round(min*7) };
}
function calcToplineDays(products){ return 7 + Math.floor((products-1)/2)*3.5; }
function calcReportDays(products){ return 7 + (products-1)*3.5; }

let _nid=500;

export default function App() {
  const [tab, setTab] = useState("quote");
  const [catalog, setCatalog] = useState(INITIAL_CATALOG);
  const [params, setParams] = useState({
    title:"HUT 留置测试",
    products:1, samplePerProduct:45, sampleTotal:45,
    sessions:2, weeksDuration:1, courierDays:2,
    difficulty:"normal", vatRate:6,
    startDate:new Date().toISOString().split("T")[0],
  });
  const setP = (k,v)=>setParams(c=>({...c,[k]:v}));

  const [lines, setLines] = useState([
    {lid:"q1",  qtyOverride:null, priceOverride:null, times:1},
    {lid:"d3",  qtyOverride:null, priceOverride:null, times:1},
    {lid:"d1",  qtyOverride:null, priceOverride:null, times:1},
    {lid:"d2",  qtyOverride:null, priceOverride:null, times:1},
  ]);

  const getAutoQty = (ci) => {
    const pm = DIFF[params.difficulty].priceMult;
    switch(ci.calcType){
      case "hut":         return Math.round(params.samplePerProduct*params.products*pm);
      case "per-person":  return Math.round(params.sampleTotal*pm);
      case "per-product": return params.products;
      case "per-session": return params.sessions;
      default:            return 1;
    }
  };

  const lineRows = useMemo(()=>lines.map(l=>{
    const ci=catalog.find(c=>c.id===l.lid);
    if(!ci) return null;
    const qty   = l.qtyOverride ?? getAutoQty(ci);
    const price = l.priceOverride ?? ci.unitPrice;
    const times = l.times ?? 1;
    return { ...l, ci, qty, price, times, total: qty*price*times };
  }).filter(Boolean), [lines,catalog,params]);

  const sub   = lineRows.reduce((s,r)=>s+r.total,0);
  const gross = +(sub*(1+params.vatRate/100)).toFixed(2);

  const tlData = useMemo(()=>{
    const start       = skipWE(new Date(params.startDate));
    const recruit     = calcRecruitWeeks(params.sampleTotal, params.difficulty);
    const leaveDays   = params.weeksDuration*7 + params.courierDays;
    const topDays     = calcToplineDays(params.products);
    const repDays     = calcReportDays(params.products);
    const screenEnd   = skipAdd(start, 3);
    const recruitStart= skipWE(screenEnd);
    const recruitEnd  = skipAdd(recruitStart, recruit.days);
    const mainQEnd    = skipAdd(recruitEnd, 3);
    const leaveStart  = skipWE(recruitEnd);
    const leaveEnd    = skipAdd(leaveStart, leaveDays);
    const topStart    = skipWE(leaveEnd);
    const topEnd      = skipAdd(topStart, topDays);
    const repStart    = skipWE(topEnd);
    const repEnd      = skipAdd(repStart, repDays);
    return {
      recruit, leaveDays, topDays, repDays,
      rows:[
        { phase:"项目确认 / PO",    days:"/", start:"—",               end:fmtD(start),      note:"里程碑" },
        { phase:"甄别问卷确认",      days:3,   start:fmtD(start),        end:fmtD(screenEnd),  note:"" },
        { phase:"招募",            days:`建议 ${recruit.min}~${recruit.max}周`, start:fmtD(recruitStart), end:fmtD(recruitEnd), note:`按最短${recruit.min}周估算` },
        { phase:"提供 Attributes", days:"/",  start:"—",               end:fmtD(skipWE(addDays(recruitEnd,-2))), note:"招募结束前2天" },
        { phase:"产品到达",         days:"/",  start:"—",               end:fmtD(skipWE(addDays(recruitEnd,-2))), note:"招募结束前2天" },
        { phase:"主问卷确认",        days:3,   start:fmtD(recruitEnd),   end:fmtD(mainQEnd),   note:"" },
        { phase:"派发和留置",        days:leaveDays, start:fmtD(leaveStart), end:fmtD(leaveEnd), note:`含快递${params.courierDays}天` },
        { phase:"Topline",        days:topDays,   start:fmtD(topStart), end:fmtD(topEnd),   note:wkStr(topDays) },
        { phase:"Report",         days:repDays,   start:fmtD(repStart), end:fmtD(repEnd),   note:wkStr(repDays) },
      ]
    };
  }, [params]);

  const updateCat = (id,k,v)=>setCatalog(a=>a.map(c=>c.id===id?{...c,[k]:v}:c));
  const deleteCat = (id)=>{ setCatalog(a=>a.filter(c=>c.id!==id)); setLines(ls=>ls.filter(l=>l.lid!==id)); };
  const addCatItem = ()=>{ _nid++; setCatalog(a=>[...a,{id:`x${_nid}`,cat:"交付与通用",name:"新费用项",unitPrice:0,calcType:"fixed",notes:"",isExtra:false}]); };
  const addLine = (catId)=>{ if(lines.find(l=>l.lid===catId)) return; setLines(ls=>[...ls,{lid:catId,qtyOverride:null,priceOverride:null,times:1}]); };
  const removeLine = (lid)=>setLines(ls=>ls.filter(l=>l.lid!==lid));
  const updateLine = (lid,k,v)=>setLines(ls=>ls.map(l=>l.lid===lid?{...l,[k]:v===''?null:(k==="times"?Math.max(1,+v):+v)}:l));

  const exportXLSX = ()=>{
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.aoa_to_sheet([
      [params.title],[],
      [`产品数:${params.products}  单产品样本:${params.samplePerProduct}人  总样本:${params.sampleTotal}人  留置:${params.weeksDuration}周  难度:${DIFF[params.difficulty].label}`],[],
      ["费用项目","计算方式","单价(RMB)","数量","倍数","小计(RMB)"],
      ...lineRows.map(r=>[r.ci.name, CALC_LABELS[r.ci.calcType], r.price, r.qty, r.times, r.total]), [],
      [`Cost (before VAT ${params.vatRate}%)`,"","","","",sub],
      [`Cost (after VAT ${params.vatRate}%)`, "","","","",gross],
    ]);
    ws1["!cols"]=[{wch:44},{wch:18},{wch:12},{wch:10},{wch:8},{wch:16}];
    XLSX.utils.book_append_sheet(wb,ws1,"报价");
    const ws2 = XLSX.utils.aoa_to_sheet([
      [params.title+" · 项目时间表"],[],
      ["阶段","时间（建议）","开始日期","结束日期","备注"],
      ...tlData.rows.map(r=>[r.phase, r.days, r.start, r.end, r.note]),
    ]);
    ws2["!cols"]=[{wch:24},{wch:16},{wch:14},{wch:14},{wch:20}];
    XLSX.utils.book_append_sheet(wb,ws2,"时间表");
    XLSX.writeFile(wb,`${params.title}_报价.xlsx`);
  };

  return (
    <div style={{fontFamily:"'Noto Serif SC',Georgia,serif",background:"#f4f1ec",minHeight:"100vh",color:"#1a1a1a"}}>
      <style>{css}</style>

      <div style={{background:"#2c2825",padding:"18px 28px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12}}>
        <div>
          <div style={{fontSize:10,letterSpacing:"0.12em",textTransform:"uppercase",color:"#9a8e80",marginBottom:4}}>HOW-TO 灏图品测 · 智能报价系统</div>
          <input value={params.title} onChange={e=>setP("title",e.target.value)}
            style={{background:"transparent",border:"none",borderBottom:"1px solid #5a504a",color:"#f5f2ed",fontSize:19,fontWeight:700,fontFamily:"inherit",outline:"none",width:360,padding:"2px 0"}} />
        </div>
        <div style={{display:"flex",alignItems:"center",gap:16}}>
          <div style={{textAlign:"right"}}>
            <div style={{fontSize:10,color:"#9a8e80",letterSpacing:"0.08em",marginBottom:3}}>含税总额</div>
            <div style={{fontSize:26,fontWeight:700,fontFamily:"'DM Mono',monospace",color:"#e8c99a"}}>{fmtRMB(gross)}</div>
          </div>
          <button className="export-btn" onClick={exportXLSX}>⬇ 导出 Excel</button>
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"272px 1fr",gap:18,padding:"18px 28px",maxWidth:1440,margin:"0 auto"}}>

        <div>
          <div style={S.card}>
            <div style={S.cardTitle}>项目参数</div>
            <F label="产品数量"><select className="inp" value={params.products} onChange={e=>setP("products",+e.target.value)}>{[1,2,3,4,5,6].map(n=><option key={n} value={n}>{n} 款</option>)}</select></F>
            <Sep label="样本量" />
            <F label="单产品样本量（HUT计价）"><input className="inp" type="number" min={1} value={params.samplePerProduct} onChange={e=>setP("samplePerProduct",+e.target.value)}/></F>
            <F label="总样本量（CLT / 定性 / 技术类）"><input className="inp" type="number" min={1} value={params.sampleTotal} onChange={e=>setP("sampleTotal",+e.target.value)}/></F>
            <Sep label="研究设置" />
            <F label="定性场次（组/场）"><input className="inp" type="number" min={1} value={params.sessions} onChange={e=>setP("sessions",+e.target.value)}/></F>
            <F label="留置时长"><select className="inp" value={params.weeksDuration} onChange={e=>setP("weeksDuration",+e.target.value)}>{[1,2,3,4].map(n=><option key={n} value={n}>{n}周（{n*7}天）</option>)}</select></F>
            <F label="快递时间（天）"><input className="inp" type="number" min={0} max={7} value={params.courierDays} onChange={e=>setP("courierDays",+e.target.value)}/></F>
            <F label="招募难易程度"><select className="inp" value={params.difficulty} onChange={e=>setP("difficulty",e.target.value)}>{Object.entries(DIFF).map(([k,v])=><option key={k} value={k}>{v.label}</option>)}</select></F>
            <F label="VAT 税率 (%)"><input className="inp" type="number" min={0} max={20} value={params.vatRate} onChange={e=>setP("vatRate",+e.target.value)}/></F>
            <F label="项目确认日期" last><input className="inp" type="date" value={params.startDate} onChange={e=>setP("startDate",e.target.value)}/></F>
          </div>

          <div style={{background:"#fff",borderRadius:8,padding:"14px 16px",border:"1px solid #e8e0d4",fontSize:12}}>
            <div style={{fontWeight:700,color:"#7a6e5f",fontSize:10,letterSpacing:"0.06em",marginBottom:10,textTransform:"uppercase"}}>时间表预估</div>
            {[
              ["甄别问卷确认","3 天"],
              ["招募周期", `建议 ${tlData.recruit.min}~${tlData.recruit.max} 周`],
              ["主问卷确认","3 天"],
              ["派发+留置", `${params.weeksDuration*7}天 + 快递${params.courierDays}天 = ${tlData.leaveDays}天`],
              ["Topline", wkStr(tlData.topDays)],
              ["Report",  wkStr(tlData.repDays)],
            ].map(([k,v])=>(
              <div key={k} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid #f0ebe3"}}>
                <span style={{color:"#888"}}>{k}</span>
                <span style={{fontWeight:600,fontFamily:"'DM Mono',monospace",color:"#2c2825",fontSize:11}}>{v}</span>
              </div>
            ))}
            <div style={{marginTop:10,fontSize:10,color:"#bbb",lineHeight:1.9}}>
              HUT计价: {params.samplePerProduct}×{params.products}款×{DIFF[params.difficulty].priceMult}<br/>
              招募: {params.sampleTotal}人÷60×{DIFF[params.difficulty].timeFactor}
            </div>
          </div>
        </div>

        <div>
          <div style={{display:"flex",gap:3,marginBottom:12}}>
            {[["quote","📋 报价明细"],["catalog","📚 费用目录"],["timeline","📅 时间表"]].map(([id,label])=>(
              <button key={id} onClick={()=>setTab(id)} style={tab===id?S.tabOn:S.tabOff} className={tab===id?"tab-on":"tab-off"}>{label}</button>
            ))}
          </div>

          {tab==="quote"&&(
            <div>
              <div style={S.card}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
                  <div style={S.cardTitle}>当前报价行项目</div>
                  <div style={{fontSize:11,color:"#bbb"}}>前往「费用目录」点「＋」添加</div>
                </div>
                <table style={S.table}>
                  <thead><tr>
                    <th style={{...S.th,width:"33%"}}>费用项目</th>
                    <th style={{...S.th,width:"13%"}}>计算方式</th>
                    <th style={{...S.th,width:"10%",textAlign:"right"}}>自动数量</th>
                    <th style={{...S.th,width:"9%",textAlign:"right"}}>覆盖量</th>
                    <th style={{...S.th,width:"9%",textAlign:"center"}}>倍数</th>
                    <th style={{...S.th,width:"10%",textAlign:"right"}}>单价 ¥</th>
                    <th style={{...S.th,width:"12%",textAlign:"right"}}>小计 ¥</th>
                    <th style={{...S.th,width:"4%"}}></th>
                  </tr></thead>
                  <tbody>
                    {lineRows.map((r,i)=>{
                      const isExtra = !!r.ci.isExtra;
                      return (
                        <tr key={r.lid} style={i%2===1?{background:"#faf8f5"}:{}}>
                          <td style={S.td}>
                            <div style={{display:"flex",alignItems:"center",gap:6}}>
                              {isExtra&&<span style={{fontSize:9,background:"#f0e6c8",color:"#a07030",borderRadius:3,padding:"1px 5px",whiteSpace:"nowrap",fontWeight:700}}>额外</span>}
                              {r.ci.name}
                            </div>
                          </td>
                          <td style={{...S.td,fontSize:11,color:"#999"}}>{CALC_LABELS[r.ci.calcType]}</td>
                          <td style={{...S.td,textAlign:"right",color:"#bbb",fontFamily:"'DM Mono',monospace",fontSize:12}}>{getAutoQty(r.ci)}</td>
                          <td style={{...S.td,textAlign:"right"}}>
                            <input type="number" placeholder="—" value={r.qtyOverride??""} onChange={e=>updateLine(r.lid,"qtyOverride",e.target.value)}
                              style={{...S.mini,width:54,textAlign:"right"}} />
                          </td>
                          <td style={{...S.td,textAlign:"center"}}>
                            <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:4}}>
                              <button onClick={()=>updateLine(r.lid,"times",Math.max(1,(r.times??1)-1))}
                                style={{...S.timesBtn,opacity:((r.times??1)<=1)?0.3:1}}>−</button>
                              <span style={{fontFamily:"'DM Mono',monospace",fontSize:13,fontWeight:700,minWidth:20,textAlign:"center",color:isExtra?"#b07828":"#555"}}>{r.times??1}</span>
                              <button onClick={()=>updateLine(r.lid,"times",(r.times??1)+1)} style={S.timesBtn}>＋</button>
                            </div>
                          </td>
                          <td style={{...S.td,textAlign:"right"}}>
                            <input type="number" placeholder={r.ci.unitPrice} value={r.priceOverride??""} onChange={e=>updateLine(r.lid,"priceOverride",e.target.value)}
                              style={{...S.mini,width:72,textAlign:"right"}} />
                          </td>
                          <td style={{...S.td,textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:600,fontSize:13}}>{fmtRMB(r.total)}</td>
                          <td style={{...S.td,textAlign:"center"}}>
                            <button onClick={()=>removeLine(r.lid)} style={S.delBtn} className="del-btn">×</button>
                          </td>
                        </tr>
                      );
                    })}
                    {lineRows.length===0&&<tr><td colSpan={8} style={{...S.td,color:"#ccc",textAlign:"center",padding:"28px 0"}}>尚未添加任何费用项</td></tr>}
                  </tbody>
                </table>
                <div style={{marginTop:12,display:"flex",justifyContent:"flex-end"}}>
                  <table style={{borderCollapse:"collapse",minWidth:340}}>
                    <tbody>
                      <tr style={{background:"#f0ebe3"}}>
                        <td style={{...S.td,fontWeight:700,paddingLeft:16}}>Cost (before VAT {params.vatRate}%)</td>
                        <td style={{...S.td,textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:700,paddingRight:16,width:140}}>{fmtRMB(sub)}</td>
                      </tr>
                      <tr style={{background:"#2c2825"}}>
                        <td style={{...S.td,fontWeight:700,color:"#f5f2ed",paddingLeft:16}}>Cost (after VAT {params.vatRate}%)</td>
                        <td style={{...S.td,textAlign:"right",fontFamily:"'DM Mono',monospace",fontWeight:700,color:"#e8c99a",paddingRight:16}}>{fmtRMB(gross)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
              <div style={{background:"#fffbf5",border:"1px solid #ece3d0",borderRadius:6,padding:"10px 14px",fontSize:11,color:"#9a8e80",lineHeight:2}}>
                💡 <b>倍数</b>：额外项（标橙）用 − / ＋ 调整，×2 = 增加2周/2份/2个
                &nbsp;｜&nbsp; 覆盖量 / 单价 留空 = 自动计算
              </div>
            </div>
          )}

          {tab==="catalog"&&(
            <div style={S.card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div style={S.cardTitle}>费用目录管理</div>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>setCatalog(INITIAL_CATALOG)} style={S.btnGhost} className="btn-ghost">重置默认</button>
                  <button onClick={addCatItem} style={S.btnDark} className="btn-dark">+ 新增目录项</button>
                </div>
              </div>
              {CATS.map(cat=>{
                const items=catalog.filter(c=>c.cat===cat);
                if(!items.length) return null;
                return (
                  <div key={cat} style={{marginBottom:22}}>
                    <div style={{fontSize:11,fontWeight:700,letterSpacing:"0.06em",color:"#7a6e5f",padding:"6px 10px",background:"#f0ebe3",borderRadius:4,marginBottom:6,display:"flex",justifyContent:"space-between"}}>
                      <span>{cat}</span>
                      {cat==="定量·HUT留置"&&<span style={{fontSize:10,color:"#b08060"}}>计价 = 单产品样本 × 产品数 × 难度系数</span>}
                    </div>
                    <table style={S.table}>
                      <thead><tr>
                        <th style={{...S.th,width:"28%"}}>名称</th>
                        <th style={{...S.th,width:"10%"}}>分类</th>
                        <th style={{...S.th,width:"11%",textAlign:"right"}}>单价</th>
                        <th style={{...S.th,width:"13%"}}>计算方式</th>
                        <th style={{...S.th,width:"8%",textAlign:"center"}}>额外项</th>
                        <th style={{...S.th,width:"13%"}}>备注</th>
                        <th style={{...S.th,width:"9%",textAlign:"center"}}>加入报价</th>
                        <th style={{...S.th,width:"5%"}}></th>
                      </tr></thead>
                      <tbody>
                        {items.map((ci,i)=>{
                          const inQ=lines.find(l=>l.lid===ci.id);
                          return (
                            <tr key={ci.id} style={i%2===1?{background:"#faf8f5"}:{}}>
                              <td style={S.td}><input value={ci.name} onChange={e=>updateCat(ci.id,"name",e.target.value)} style={S.inlineInput}/></td>
                              <td style={S.td}><select value={ci.cat} onChange={e=>updateCat(ci.id,"cat",e.target.value)} style={{...S.inlineInput,fontSize:10,padding:"3px 4px"}}>{CATS.map(c=><option key={c} value={c}>{c}</option>)}</select></td>
                              <td style={{...S.td,textAlign:"right"}}><input type="number" value={ci.unitPrice} onChange={e=>updateCat(ci.id,"unitPrice",+e.target.value)} style={{...S.inlineInput,width:76,textAlign:"right",fontFamily:"'DM Mono',monospace"}}/></td>
                              <td style={S.td}><select value={ci.calcType} onChange={e=>updateCat(ci.id,"calcType",e.target.value)} style={{...S.inlineInput,fontSize:10,padding:"3px 4px"}}>{Object.entries(CALC_LABELS).map(([k,v])=><option key={k} value={k}>{v}</option>)}</select></td>
                              <td style={{...S.td,textAlign:"center"}}>
                                <input type="checkbox" checked={!!ci.isExtra} onChange={e=>updateCat(ci.id,"isExtra",e.target.checked)} style={{width:15,height:15,accentColor:"#b07828"}}/>
                              </td>
                              <td style={S.td}><input value={ci.notes} onChange={e=>updateCat(ci.id,"notes",e.target.value)} style={{...S.inlineInput,fontSize:11}} placeholder="—"/></td>
                              <td style={{...S.td,textAlign:"center"}}>
                                {inQ?<span style={{fontSize:11,color:"#27ae60",fontWeight:700}}>✓ 已选</span>
                                    :<button onClick={()=>addLine(ci.id)} style={S.addBtn} className="add-btn">＋</button>}
                              </td>
                              <td style={{...S.td,textAlign:"center"}}><button onClick={()=>deleteCat(ci.id)} style={S.delBtn} className="del-btn">×</button></td>
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

          {tab==="timeline"&&(
            <div style={S.card}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
                <div style={S.cardTitle}>项目时间表 · 周末自动顺延</div>
                <div style={{display:"flex",gap:14,fontSize:11,color:"#888"}}>
                  <span>招募 <b style={{color:"#2c2825"}}>{tlData.recruit.min}~{tlData.recruit.max}周</b></span>
                  <span>Topline <b style={{color:"#2c2825"}}>{wkStr(tlData.topDays)}</b></span>
                  <span>Report <b style={{color:"#2c2825"}}>{wkStr(tlData.repDays)}</b></span>
                </div>
              </div>
              <table style={S.table}>
                <thead><tr>
                  <th style={{...S.th,width:"28%"}}>阶段</th>
                  <th style={{...S.th,width:"18%"}}>时间（建议）</th>
                  <th style={{...S.th,width:"18%",textAlign:"center"}}>开始日期</th>
                  <th style={{...S.th,width:"18%",textAlign:"center"}}>结束日期</th>
                  <th style={{...S.th,width:"18%"}}>备注</th>
                </tr></thead>
                <tbody>
                  {tlData.rows.map((r,i)=>(
                    <tr key={i} style={i%2===1?{background:"#faf8f5"}:{}}>
                      <td style={{...S.td,fontWeight:500}}>{r.phase}</td>
                      <td style={{...S.td,fontFamily:"'DM Mono',monospace",fontSize:12,color:"#555"}}>{r.days}</td>
                      <td style={{...S.td,textAlign:"center",fontFamily:"'DM Mono',monospace",fontSize:12}}>{r.start}</td>
                      <td style={{...S.td,textAlign:"center",fontFamily:"'DM Mono',monospace",fontSize:12,fontWeight:700}}>{r.end}</td>
                      <td style={{...S.td,fontSize:11,color:"#aaa"}}>{r.note}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{marginTop:12,padding:"10px 12px",background:"#f8f5f0",borderRadius:6,fontSize:11,color:"#999",lineHeight:2}}>
                📌 甄别问卷确认 3天 → 招募开始 &nbsp;｜&nbsp;
                派发留置 = {params.weeksDuration}周×7 + 快递{params.courierDays}天 = <b>{tlData.leaveDays}天</b> &nbsp;｜&nbsp;
                Topline {params.products}款 → <b>{wkStr(tlData.topDays)}</b> &nbsp;｜&nbsp;
                Report {params.products}款 → <b>{wkStr(tlData.repDays)}</b>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function F({label,children,last}){
  return <div style={{marginBottom:last?0:12}}>
    <label style={{fontSize:10,letterSpacing:"0.08em",textTransform:"uppercase",color:"#888",marginBottom:5,display:"block"}}>{label}</label>
    {children}
  </div>;
}
function Sep({label}){
  return <div style={{fontSize:10,fontWeight:700,color:"#b0a898",letterSpacing:"0.06em",textTransform:"uppercase",borderTop:"1px solid #ece8e2",paddingTop:10,marginBottom:10,marginTop:4}}>{label}</div>;
}

const S={
  card:{background:"#fff",borderRadius:8,padding:"20px 22px",marginBottom:14,boxShadow:"0 1px 4px rgba(0,0,0,0.05)"},
  cardTitle:{fontSize:11,fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",color:"#7a6e5f",marginBottom:14,paddingBottom:10,borderBottom:"1px solid #ece8e2"},
  table:{width:"100%",borderCollapse:"collapse",fontSize:13},
  th:{background:"#2c2825",color:"#f5f2ed",padding:"9px 10px",textAlign:"left",fontWeight:600,letterSpacing:"0.04em",fontSize:11},
  td:{padding:"8px 10px",borderBottom:"1px solid #ede9e3",verticalAlign:"middle"},
  tabOn:{background:"#2c2825",color:"#f5f2ed",border:"none",borderRadius:"6px 6px 0 0",padding:"9px 18px",fontFamily:"inherit",fontSize:13,fontWeight:700,cursor:"pointer"},
  tabOff:{background:"#e8e3dc",color:"#7a6e5f",border:"none",borderRadius:"6px 6px 0 0",padding:"9px 18px",fontFamily:"inherit",fontSize:13,cursor:"pointer"},
  inlineInput:{border:"1px solid #e0dbd3",borderRadius:4,padding:"4px 7px",fontSize:12,fontFamily:"inherit",width:"100%",outline:"none",background:"#faf8f5"},
  mini:{border:"1px solid #e0dbd3",borderRadius:4,padding:"3px 5px",fontSize:12,fontFamily:"'DM Mono',monospace",outline:"none",background:"#faf8f5"},
  timesBtn:{background:"#f0ebe3",border:"none",borderRadius:3,width:22,height:22,cursor:"pointer",fontSize:13,fontWeight:700,color:"#7a6e5f",display:"flex",alignItems:"center",justifyContent:"center",padding:0,lineHeight:1},
  delBtn:{background:"none",border:"none",cursor:"pointer",color:"#c0392b",fontSize:17,fontWeight:700,padding:"0 3px",lineHeight:1},
  addBtn:{background:"#2c2825",color:"#f5f2ed",border:"none",borderRadius:4,padding:"3px 10px",fontSize:13,cursor:"pointer",fontWeight:700},
  btnDark:{background:"#2c2825",color:"#f5f2ed",border:"none",borderRadius:5,padding:"7px 14px",fontFamily:"inherit",fontSize:12,cursor:"pointer",fontWeight:600},
  btnGhost:{background:"none",border:"1px solid #c4bfb8",color:"#7a6e5f",borderRadius:5,padding:"7px 14px",fontFamily:"inherit",fontSize:12,cursor:"pointer"},
};

const css=`
  @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+SC:wght@400;600;700&family=DM+Mono:wght@400;500&display=swap');
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:#f4f1ec;}
  .inp{width:100%;padding:8px 10px;border:1px solid #d4cfc8;border-radius:4px;background:#fff;font-size:13px;color:#1a1a1a;outline:none;font-family:inherit;}
  .inp:focus{border-color:#7a6e5f;}
  .export-btn{background:#e8c99a;color:#2c2825;border:none;border-radius:6px;padding:10px 18px;font-family:inherit;font-weight:700;font-size:13px;cursor:pointer;}
  .export-btn:hover{background:#d4b580;}
  .btn-dark:hover{background:#3d3530;}
  .btn-ghost:hover{background:#f0ebe3;}
  .del-btn:hover{color:#e74c3c;}
  .add-btn:hover{background:#3d3530;}
  .tab-off:hover{background:#ddd8d0;}
`;
