import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

/* ══ 상수 ══ */
const STORAGE_KEY = "oceanmade_door_v1";
const API_URL     = "https://api.anthropic.com/v1/messages";
const ANTHROPIC_KEY = import.meta.env.VITE_ANTHROPIC_API_KEY || "";

const STATUSES = [
  { key: "received",   label: "접수도면",    color: "#F59E0B", icon: "📨" },
  { key: "production", label: "제작중",      color: "#A78BFA", icon: "🔨" },
  { key: "ready",      label: "배송준비완료", color: "#F472B6", icon: "🚚" },
  { key: "done",       label: "출고",        color: "#34D399", icon: "✅" },
];
const DOOR_TYPES    = ["예림", "한솔", "LX"];
const COLOR_PRESETS = ["매트화이트","매트밀크화이트","매트캐시미어","매트포그그레이","글로시화이트","글로시밀크화이트","직접입력"];
const EMPTY_ITEM      = { color: "매트화이트", qty: "", unitPrice: "" };
const DELIVERY_TYPES  = ["직접출고", "용차출고", "예림배송"];
const EMPTY_FORM      = { company:"", doorType:"예림", items:[{...EMPTY_ITEM}], status:"received", dueDate:"", memo:"", deliveryType:"직접출고" };
const BRAND_KEYS    = ["예림","한솔","LX","lx"];
const SKIP_KEYWORDS = ["소계","합계","총합","비고","※","견적","공사","품명","규격","수량","단가","금액","SUPER","인테리어","입금","색상명","필름","레이저","브랜드","전화","팩스","주소","업체","현장","접수","납기"];

/* ══ 유틸 ══ */
function daysLeft(due) {
  return Math.ceil((new Date(due) - new Date()) / 86400000);
}
function totalQty(items)  { return (items||[]).reduce((s,i) => s + (parseFloat(i.qty)||0), 0); }
function totalAmt(items)  { return (items||[]).reduce((s,i) => s + (parseFloat(i.qty)||0) * (parseFloat(i.unitPrice)||0), 0); }
function fmtQty(n)        { const r = Math.round(n*10)/10; return r%1===0 ? `${r}` : r.toFixed(1); }
function excelSerialToDate(serial) {
  if (!serial || isNaN(Number(serial))) return "";
  const d = new Date(Math.round((Number(serial) - 25569) * 86400 * 1000));
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;
}
function detectBrand(str) {
  const s = String(str||"").trim();
  if (BRAND_KEYS.some(k => s.toLowerCase()===k.toLowerCase())) return s.toUpperCase()==="LX"?"LX":s;
  return null;
}

/* ══ 엑셀 파싱 ══ */
function parseQuoteFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type:"array", cellDates:false });

        // 파일명에서 날짜만 파싱 (거래처/현장은 내부에서만 읽음)
        const base = file.name.replace(/\.(xlsx|xls)$/i, "");
        const parts = base.split("_");
        const datePart = parts.find(p => /^\d{6}$/.test(p)) || "";
        const fileDate = datePart.length===6
          ? `20${datePart.slice(0,2)}-${datePart.slice(2,4)}-${datePart.slice(4,6)}` : "";

        const sheetName = wb.SheetNames.find(n => n.includes("도매")) || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });

        // 내부에서만 거래처명/현장명 읽기 (띄어쓰기 무시 키워드 매칭)
        let company = "", location = "", dueDate = fileDate, deliveryType = "직접출고";
        for (const row of rows) {
          const flat = row.map(c => String(c??"").trim());
          const labeled = (kw) => {
            const idx = flat.findIndex(c => c.replace(/\s/g,"").includes(kw.replace(/\s/g,"")));
            if (idx===-1) return null;
            return flat.slice(idx+1).find(c => c && c!==":" && c.trim()!=="") ?? null;
          };
          const comp = labeled("거래처명") || labeled("업체명");
          if (comp) company = comp;
          // 현장명 — 회사/주소 관련 키워드가 아닐 때만 저장, 없으면 미정
          const loc = labeled("현장명");
          const locSkip = ["상호","법인","오션","성명","주소","등록","전화","팩스","업태","업종",":"];
          if (loc && !locSkip.some(k => loc.includes(k))) {
            location = loc;
          } else if (!location) {
            location = "미정";
          }
          // 출고일 우선, 없으면 납기일
          const outDate = labeled("출고일") || labeled("납기일");
          if (outDate) { const cv = excelSerialToDate(outDate); if (cv) dueDate = cv; }
          // 출고방식 — 유연한 키워드 매칭
          const deliv = labeled("출고방식");
          if (deliv) {
            const d = deliv.replace(/\s/g,"");
            if (d.includes("예림") || d.includes("배송")) deliveryType = "예림배송";
            else if (d.includes("용차")) deliveryType = "용차출고";
            else if (d.includes("직접") || d.includes("수령")) deliveryType = "직접출고";
          }
        }

        let colorCol=0, brandCol=-1, qtyCol=-1, priceCol=-1, headerRowIdx=-1;
        for (let ri=0; ri<rows.length; ri++) {
          const flat = rows[ri].map(c => String(c??"").trim());
          if (flat.some(c => c==="색상명")) {
            headerRowIdx=ri;
            colorCol  = flat.findIndex(c => c==="색상명");
            brandCol  = flat.findIndex(c => c.includes("브랜드"));
            qtyCol    = flat.findIndex(c => c.includes("수")&&c.includes("량"));
            priceCol  = flat.findIndex(c => c.includes("단")&&c.includes("가"));
            break;
          }
        }

        const items = []; let detectedBrand = "예림";
        const dataRows = headerRowIdx>=0 ? rows.slice(headerRowIdx+1) : rows;
        for (const row of dataRows) {
          const flat = row.map(c => String(c??"").trim());
          const colorVal = flat[colorCol] || "";
          if (!colorVal || SKIP_KEYWORDS.some(k => colorVal.includes(k))) continue;
          let qty="", price="";
          if (qtyCol>=0 && flat[qtyCol] && !isNaN(Number(flat[qtyCol]))) qty = flat[qtyCol];
          if (priceCol>=0 && flat[priceCol] && !isNaN(Number(flat[priceCol].replace(/,/g,"")))) price = flat[priceCol].replace(/,/g,"");
          if (!qty || !price) {
            const nums = flat.filter(c => c && !isNaN(Number(c.replace(/,/g,""))) && Number(c.replace(/,/g,""))>0).map(c => parseFloat(c.replace(/,/g,"")));
            if (nums.length<2) continue;
            qty   = qty   || String(nums.find(n=>n<1000)??nums[0]);
            price = price || String(nums.find(n=>n>=1000)??nums[1]);
          }
          if (!qty || !price) continue;
          if (brandCol>=0) { const b=detectBrand(flat[brandCol]); if(b) detectedBrand=b; }
          items.push({ color:colorVal, qty:String(qty), unitPrice:String(price) });
        }
        resolve({ company, memo:location, dueDate, doorType:detectedBrand, status:"received",
          deliveryType, items: items.length>0 ? items : [{...EMPTY_ITEM}] });
      } catch(err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/* ══ 이미지 파싱 (Claude API) ══ */
async function parseImageFile(file) {
  const toBase64 = (f) => new Promise((res,rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = rej;
    r.readAsDataURL(f);
  });
  const base64 = await toBase64(file);
  const resp = await fetch(API_URL, {
    method:"POST",
    headers:{ "Content-Type":"application/json", "x-api-key": ANTHROPIC_KEY, "anthropic-version":"2023-06-01" },
    body: JSON.stringify({
      model:"claude-sonnet-4-20250514", max_tokens:1000,
      messages:[{ role:"user", content:[
        { type:"image", source:{ type:"base64", media_type: file.type||"image/jpeg", data:base64 } },
        { type:"text", text:`이 견적서 이미지에서 정보를 추출해서 JSON만 응답하세요(다른 텍스트 없이).
{"company":"업체명","memo":"현장명","dueDate":"YYYY-MM-DD or 빈문자열","doorType":"예림/한솔/LX","items":[{"color":"색상명","qty":"숫자문자열","unitPrice":"숫자만"}]}
- 브랜드명이 색상에 포함된 경우 color에서 제거` }
      ]}]
    })
  });
  const data = await resp.json();
  const text = data.content?.find(c=>c.type==="text")?.text || "";
  const parsed = JSON.parse(text.replace(/```json|```/g,"").trim());
  return { company:parsed.company||"", memo:parsed.memo||"", dueDate:parsed.dueDate||"",
    doorType:parsed.doorType||"예림", status:"received",
    items:(parsed.items||[]).length>0 ? parsed.items : [{...EMPTY_ITEM}] };
}

/* ══ 뱃지 ══ */
function DueBadge({ dueDate, status }) {
  if (status==="done" || !dueDate) return null;
  const d = daysLeft(dueDate);
  const [bg,fg,txt] = d<0  ? ["#450a0a","#f87171",`D+${Math.abs(d)} 초과`]
                   : d<=3 ? ["#422006","#fbbf24",`D-${d} 임박`]
                           : ["#052e16","#4ade80",`D-${d}`];
  return <span style={{background:bg,color:fg,borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:800}}>{txt}</span>;
}
function Chip({n}) {
  return <span style={{background:"rgba(255,255,255,.1)",borderRadius:999,padding:"0 7px",fontSize:11,marginLeft:4}}>{n}</span>;
}
function Field({label,children}) {
  return <div><div style={{fontSize:12,color:"#64748B",fontWeight:700,marginBottom:6}}>{label}</div>{children}</div>;
}

/* ══════════════════════════════════════════════
   메인 앱
══════════════════════════════════════════════ */
export default function App() {
  const [orders,   setOrders]   = useState(null);
  const [view,     setView]     = useState("list");
  const [editId,   setEditId]   = useState(null);
  const [form,     setForm]     = useState(EMPTY_FORM);
  const [filter,   setFilter]   = useState("all");
  const [search,   setSearch]   = useState("");
  const [sort,     setSort]     = useState("dueDate");
  const [toast,    setToast]    = useState(null);
  const [preview,  setPreview]  = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [loading,  setLoading]  = useState(false);
  const fileInputRef = useRef(null);
  const nextId = useRef(1);

  /* ── Supabase DB 연동 ── */
  const migrateStatus = s => {
    const map = { inquiry:"received", measuring:"received", delivery:"ready" };
    return map[s] || (STATUSES.find(st=>st.key===s) ? s : "received");
  };

  // DB에서 전체 주문 로드
  const loadOrders = useCallback(async () => {
    const { data, error } = await supabase.from("orders").select("*").order("id");
    if (error) { console.error(error); setOrders([]); return; }
    const rows = (data||[]).map(r => ({
      id: r.id,
      company: r.company||"",
      doorType: r.door_type||"예림",
      items: r.items||[],
      status: migrateStatus(r.status||"received"),
      dueDate: r.due_date||"",
      memo: r.memo||"",
      paid: r.paid||false,
      deliveryType: r.delivery_type||"직접출고",
      createdAt: r.created_at?.slice(0,10)||"",
    }));
    setOrders(rows);
    if (rows.length) nextId.current = Math.max(...rows.map(o=>o.id)) + 1;
  }, []);

  useEffect(() => {
    loadOrders();
    // 실시간 구독 — 다른 기기 변경사항 즉시 반영
    const channel = supabase.channel("orders-realtime")
      .on("postgres_changes", { event:"*", schema:"public", table:"orders" }, () => loadOrders())
      .subscribe();
    return () => supabase.removeChannel(channel);
  }, [loadOrders]);

  // DB 저장 함수
  async function dbUpsert(order) {
    const row = {
      id: order.id,
      company: order.company,
      door_type: order.doorType,
      items: order.items,
      status: order.status,
      due_date: order.dueDate,
      memo: order.memo,
      paid: order.paid||false,
      delivery_type: order.deliveryType||"직접출고",
    };
    const { error } = await supabase.from("orders").upsert(row);
    if (error) { flash("저장 실패","error"); console.error(error); }
  }

  async function dbDelete(id) {
    const { error } = await supabase.from("orders").delete().eq("id", id);
    if (error) { flash("삭제 실패","error"); console.error(error); }
  }

  function flash(msg, type="ok") { setToast({msg,type}); setTimeout(()=>setToast(null),2800); }

  /* ── CRUD ── */
  function openNew()   { setForm(EMPTY_FORM); setEditId(null); setPreview(null); setView("form"); }
  function openEdit(o) { setForm({...o});     setEditId(o.id); setPreview(null); setView("form"); }

  async function submitForm() {
    if (!form.company || !form.dueDate || form.items.every(i=>!i.qty)) {
      flash("거래처명·납기일·장수 필수","error"); return;
    }
    const { _savePaid, ...cleanForm } = form;
    if (editId) {
      const orig = orders.find(o=>o.id===editId);
      const updated = {...cleanForm, id:editId, createdAt:orig.createdAt, paid: _savePaid ? true : orig.paid};
      setOrders(orders.map(o=>o.id===editId?updated:o));
      await dbUpsert(updated);
      flash("수정 저장 완료");
    } else {
      const id = nextId.current++;
      const newOrder = {...cleanForm, id, paid:false, createdAt:new Date().toISOString().slice(0,10)};
      setOrders([...orders, newOrder]);
      await dbUpsert(newOrder);
      flash("새 주문 등록 완료");
    }
    setView("list"); setPreview(null);
  }

  async function deleteOrder(id, name) {
    setOrders(orders.filter(o=>o.id!==id));
    await dbDelete(id);
    flash(`"${name}" 삭제됐습니다`,"warn");
  }

  async function changeStatus(id, status) {
    const updated = orders.map(o=>o.id===id?{...o,status}:o);
    setOrders(updated);
    await dbUpsert(updated.find(o=>o.id===id));
  }

  async function markPaid(id) {
    const updated = orders.map(o=>o.id===id?{...o,paid:true}:o);
    setOrders(updated);
    await dbUpsert(updated.find(o=>o.id===id));
    flash("💳 결제완료 — 제작 시작!");
  }

  async function changeDelivery(id, deliveryType) {
    const updated = orders.map(o=>o.id===id?{...o,deliveryType}:o);
    setOrders(updated);
    await dbUpsert(updated.find(o=>o.id===id));
  }

  /* ── 파일 업로드 ── */
  async function handleFiles(files) {
    const file = [...files][0];
    if (!file) return;
    const isImage = file.type.startsWith("image/");
    const isXlsx  = /\.(xlsx|xls)$/i.test(file.name);
    if (!isImage && !isXlsx) { flash("엑셀 또는 이미지 파일만 가능합니다","error"); return; }
    setLoading(true);
    try {
      const parsed = isImage ? await parseImageFile(file) : await parseQuoteFile(file);
      setForm({...EMPTY_FORM, ...parsed});
      setEditId(null); setPreview(parsed); setView("form");
    } catch(e) {
      flash("파일 파싱 실패 — 형식을 확인해주세요","error");
    } finally { setLoading(false); }
  }
  function onFileInput(e) { if(e.target.files?.length) handleFiles(e.target.files); e.target.value=""; }
  function onDrop(e)  { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }

  /* ── 필터 ── */
  const filtered = (orders||[])
    .filter(o => {
      if (filter === "all") return o.status !== "done"; // 전체에서 출고 제외
      return o.status === filter;
    })
    .filter(o => !search || [o.company,o.doorType,o.memo].some(v=>v&&v.includes(search)))
    .sort((a,b) => sort==="dueDate" ? new Date(a.dueDate)-new Date(b.dueDate)
                 : sort==="company" ? a.company.localeCompare(b.company,"ko")
                 : b.id-a.id);

  const urgentCnt = (orders||[]).filter(o=>daysLeft(o.dueDate)<=3&&o.status!=="done").length;

  if (!orders) return (
    <div style={{minHeight:"100vh",background:"#0C1220",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{width:40,height:40,border:"3px solid #1E3A5F",borderTopColor:"#3B82F6",borderRadius:"50%",animation:"spin .9s linear infinite"}}/>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"#0C1220",fontFamily:"'Noto Sans KR',sans-serif",color:"#E2E8F0"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&family=Bebas+Neue&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        ::-webkit-scrollbar{width:4px;height:4px}
        ::-webkit-scrollbar-thumb{background:#1E3A5F;border-radius:4px}
        .card{background:#111C2D;border:1px solid #1E3A5F;border-radius:12px;transition:box-shadow .2s}
        .btn{border:none;border-radius:8px;padding:8px 14px;font-weight:700;font-size:13px;font-family:inherit;cursor:pointer;transition:opacity .15s;-webkit-tap-highlight-color:transparent}
        .btn:active{opacity:.7}
        .btn-blue{background:#1D4ED8;color:#fff}
        .btn-slate{background:#1E293B;color:#94A3B8}
        .btn-red{background:#7F1D1D;color:#FCA5A5}
        .inp{background:#0C1220;border:1px solid #1E3A5F;border-radius:8px;color:#E2E8F0;padding:10px 12px;font-size:14px;font-family:inherit;width:100%;-webkit-appearance:none}
        .inp:focus{outline:none;border-color:#3B82F6}
        @keyframes slideUp{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
        @keyframes blink{0%,100%{opacity:1}50%{opacity:.4}}
        @keyframes pulse3{0%,100%{box-shadow:0 4px 20px rgba(99,102,241,0.4)}50%{box-shadow:0 4px 32px rgba(99,102,241,0.8)}}
        /* 모바일 터치 영역 확보 */
        select.inp{min-height:42px}
        input.inp{min-height:42px}
      `}</style>

      {/* ── HEADER ── */}
      <header style={{padding:"14px 16px",borderBottom:"1px solid #1E3A5F",position:"sticky",top:0,background:"#0C1220",zIndex:50}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
          <div>
            <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:20,letterSpacing:1.5,color:"#fff",lineHeight:1}}>
              오션메이드 <span style={{color:"#3B82F6"}}>도어발주현황</span>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:6,marginTop:2}}>
              <div style={{width:6,height:6,borderRadius:"50%",background:"#34D399"}}/>
              <span style={{fontSize:10,color:"#475569"}}>실시간 공유중</span>
              {urgentCnt>0 && <span style={{background:"#7F1D1D",color:"#FCA5A5",borderRadius:999,padding:"1px 7px",fontSize:10,fontWeight:800,animation:"blink 2s infinite"}}>⚠ {urgentCnt}건 임박</span>}
            </div>
          </div>
          <div style={{display:"flex",gap:6}}>
            <button className="btn btn-slate" style={{padding:"7px 10px",fontSize:12}}
              onClick={()=>setView(v=>v==="kanban"?"list":"kanban")}>
              {view==="kanban"?"📋":"🗂"}
            </button>
            <button className="btn btn-slate" style={{padding:"7px 10px",fontSize:12}}
              onClick={()=>setView(v=>v==="ranking"?"list":"ranking")}>
              🏆
            </button>
            <div
              onClick={()=>!loading&&fileInputRef.current?.click()}
              onDrop={onDrop} onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
              style={{border:`2px dashed ${dragOver?"#3B82F6":"#1E3A5F"}`,borderRadius:8,padding:"7px 10px",
                cursor:"pointer",background:dragOver?"#1E3A5F22":"transparent",
                color:dragOver?"#3B82F6":"#475569",fontSize:12,fontWeight:600,
                display:"flex",alignItems:"center",gap:5,minWidth:80,justifyContent:"center"}}
            >
              {loading ? <span style={{fontSize:14,animation:"spin .8s linear infinite",display:"inline-block"}}>⏳</span> : "📎"}
              <span>{loading?"분석중...":"견적서"}</span>
            </div>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,image/*" style={{display:"none"}} onChange={onFileInput}/>
            <button className="btn btn-blue" style={{padding:"7px 12px",fontSize:12}} onClick={openNew}>＋ 등록</button>
          </div>
        </div>
      </header>

      {/* ── 견적서 업로드 바 ── */}
      {view!=="form" && (
        <div
          onClick={()=>!loading&&fileInputRef.current?.click()}
          onDrop={onDrop} onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
          style={{
            margin:"12px 14px 0",
            border:`2px dashed ${dragOver?"#3B82F6":"#1E3A5F"}`,
            borderRadius:12, padding:"16px 20px",
            cursor:"pointer", transition:"all .2s",
            background: dragOver?"#1E3A5F33":"#111C2D",
            display:"flex", alignItems:"center", justifyContent:"center", gap:12,
          }}
        >
          <span style={{fontSize:28}}>{loading?"⏳":"📎"}</span>
          <div>
            <div style={{color: dragOver?"#3B82F6":"#E2E8F0", fontSize:15, fontWeight:700}}>
              {loading?"견적서 분석중...":"견적서 업로드 (클릭 또는 파일 드래그)"}
            </div>
            <div style={{color:"#475569", fontSize:12, marginTop:2}}>엑셀(.xlsx) 또는 이미지 파일</div>
          </div>
        </div>
      )}

      {/* ── STAT BAR ── */}
      {view!=="form" && (
        <div style={{display:"flex",borderBottom:"1px solid #1E3A5F"}}>
          {[
            {label:"전체",  val:orders.length,                                  color:"#3B82F6"},
            {label:"접수",  val:orders.filter(o=>o.status==="received").length,  color:"#F59E0B"},
            {label:"제작중", val:orders.filter(o=>o.status==="production").length,color:"#A78BFA"},
            {label:"출고",  val:orders.filter(o=>o.status==="done").length,      color:"#34D399"},
          ].map(s=>(
            <div key={s.label} style={{flex:1,textAlign:"center",padding:"10px 4px",borderRight:"1px solid #1E3A5F"}}>
              <div style={{fontSize:20,fontWeight:900,color:s.color,fontFamily:"'Bebas Neue',sans-serif"}}>{s.val}</div>
              <div style={{fontSize:10,color:"#475569"}}>{s.label}</div>
            </div>
          ))}
        </div>
      )}

      {/* ── MAIN ── */}
      <main style={{padding:14}}>
        {view==="form"
          ? <FormPanel form={form} setForm={setForm} onSave={submitForm} onCancel={()=>{setView("list");setPreview(null);}} isEdit={!!editId} preview={preview}/>
          : view==="kanban"
          ? <KanbanPanel orders={orders} onEdit={openEdit} onDelete={deleteOrder} onStatus={changeStatus}/>
          : view==="ranking"
          ? <RankingPanel orders={orders}/>
          : <ListPanel orders={filtered} all={orders}
              filter={filter} setFilter={setFilter}
              search={search} setSearch={setSearch}
              sort={sort} setSort={setSort}
              onEdit={openEdit} onDelete={deleteOrder} onStatus={changeStatus} onPaid={markPaid} onDelivery={changeDelivery}/>
        }
      </main>

      {/* ── TOAST ── */}
      {toast && (
        <div style={{position:"fixed",bottom:20,left:"50%",transform:"translateX(-50%)",
          padding:"12px 20px",borderRadius:10,fontWeight:700,fontSize:14,animation:"slideUp .3s",
          zIndex:999,whiteSpace:"nowrap",
          background: toast.type==="error"?"#450a0a":toast.type==="warn"?"#1c1007":"#052e16",
          color:       toast.type==="error"?"#f87171":toast.type==="warn"?"#fbbf24":"#4ade80",
          border:`1px solid ${toast.type==="error"?"#7f1d1d":toast.type==="warn"?"#78350f":"#14532d"}`,
        }}>
          {toast.type==="error"?"❌ ":toast.type==="warn"?"⚠ ":"✅ "}{toast.msg}
        </div>
      )}
    </div>
  );
}

/* ══ 목록 패널 ══ */
function ListPanel({orders,all,filter,setFilter,search,setSearch,sort,setSort,onEdit,onDelete,onStatus,onPaid,onDelivery}) {
  return (
    <div>
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap"}}>
        <input className="inp" style={{flex:"1 1 160px",minWidth:0}} placeholder="🔍 거래처·브랜드·현장 검색"
          value={search} onChange={e=>setSearch(e.target.value)}/>
        <select className="inp" style={{flex:"0 0 100px"}} value={sort} onChange={e=>setSort(e.target.value)}>
          <option value="dueDate">납기일순</option>
          <option value="company">거래처순</option>
          <option value="latest">최신순</option>
        </select>
      </div>
      {/* 상태 필터 탭 */}
      <div style={{display:"flex",gap:6,marginBottom:12,overflowX:"auto",paddingBottom:4}}>
        {[{key:"all",label:"전체",n:all.filter(o=>o.status!=="done").length},...STATUSES.map(s=>({key:s.key,label:s.label,n:all.filter(o=>o.status===s.key).length}))].map(t=>(
          <button key={t.key}
            className={`btn ${filter===t.key?"btn-blue":"btn-slate"}`}
            style={{flexShrink:0,padding:"6px 10px",fontSize:12}}
            onClick={()=>setFilter(t.key)}>
            {t.label} <span style={{opacity:.7}}>{t.n}</span>
          </button>
        ))}
      </div>
      {orders.length===0
        ? <div style={{textAlign:"center",padding:"60px 0",color:"#1E3A5F"}}><div style={{fontSize:40}}>🚪</div><div style={{marginTop:8,fontSize:13}}>등록된 주문 없음</div></div>
        : <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {orders.map(o=><OrderRow key={o.id} order={o} onEdit={onEdit} onDelete={onDelete} onStatus={onStatus} onPaid={onPaid} onDelivery={onDelivery}/>)}
          </div>
      }
    </div>
  );
}

/* ══ 주문 카드 ══ */
function OrderRow({order,onEdit,onDelete,onStatus,onPaid,onDelivery}) {
  const st = STATUSES.find(s=>s.key===order.status)||STATUSES[0];
  const urgent = daysLeft(order.dueDate)<=3 && order.status!=="done";
  const [confirmDel, setConfirmDel] = useState(false);
  const [confirmPay, setConfirmPay] = useState(false);
  const amt = totalAmt(order.items||[]);
  const qty = totalQty(order.items||[]);
  const isPaid = order.paid;
  const isReceived = order.status === "received";

  return (
    <div className="card" style={{padding:"12px 14px",borderLeft:`4px solid ${st.color}`,background:urgent?"#150c08":"#111C2D"}}>
      {/* 상단: 업체명 + 뱃지 */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6,flexWrap:"wrap"}}>
        <span style={{fontSize:15,fontWeight:700,color:"#F1F5F9"}}>{order.company}</span>
        <span style={{background:st.color+"28",color:st.color,borderRadius:6,padding:"2px 7px",fontSize:11,fontWeight:700}}>{st.icon} {st.label}</span>
        <DueBadge dueDate={order.dueDate} status={order.status}/>
        {isPaid && <span style={{background:"#052e16",color:"#4ade80",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:800}}>💳 결제완료</span>}
      </div>
      {/* 정보 */}
      <div style={{fontSize:12,color:"#CBD5E1",lineHeight:1.8}}>
        <div>🚪 {order.doorType} &ensp; 📅 {order.dueDate} &ensp;
          <span style={{
            background: order.deliveryType==="직접출고"?"#1e3a5f": order.deliveryType==="용차출고"?"#3b1f5e":"#1a3a2a",
            color: order.deliveryType==="직접출고"?"#60a5fa": order.deliveryType==="용차출고"?"#c084fc":"#4ade80",
            borderRadius:6, padding:"1px 8px", fontSize:11, fontWeight:700
          }}>{order.deliveryType||"직접출고"}</span>
        </div>
        <div>🎨 {(order.items||[]).map(i=>`${i.color} ${fmtQty(parseFloat(i.qty)||0)}장`).join(" / ")}</div>
        <div style={{color:"#E2E8F0",fontWeight:700}}>
          총 {fmtQty(qty)}장
          {amt>0 && <span style={{color:"#34D399",marginLeft:8}}>{amt.toLocaleString()}원</span>}
        </div>
        {order.memo && <div style={{color:"#94A3B8"}}>📝 {order.memo}</div>}
      </div>

      {/* 결제완료 확인 팝업 */}
      {confirmPay && (
        <div style={{margin:"12px 0",background:"#0a2540",border:"2px solid #1D4ED8",borderRadius:12,padding:"14px 16px"}}>
          <div style={{color:"#60A5FA",fontWeight:800,fontSize:14,marginBottom:8}}>💳 결제금액이 정확한가요?</div>
          <div style={{color:"#E2E8F0",fontSize:13,marginBottom:4}}>
            거래처: <b>{order.company}</b>
          </div>
          <div style={{color:"#E2E8F0",fontSize:13,marginBottom:4}}>
            총 <b>{fmtQty(qty)}장</b>
          </div>
          <div style={{fontSize:18,fontWeight:900,color:"#34D399",marginBottom:14}}>
            {amt.toLocaleString()}원
          </div>
          <div style={{display:"flex",gap:8}}>
            <button
              className="btn"
              style={{flex:1,background:"#166534",color:"#4ade80",fontSize:14,padding:"10px",fontWeight:800}}
              onClick={()=>{ onStatus(order.id, "production"); onPaid(order.id); setConfirmPay(false); }}
            >
              ✅ 맞아요 — 제작 시작
            </button>
            <button className="btn btn-slate" style={{padding:"10px 14px"}} onClick={()=>setConfirmPay(false)}>취소</button>
          </div>
        </div>
      )}

      {/* 결제완료 버튼 — 접수도면 상태이고 아직 결제 전일 때만 표시 */}
      {isReceived && !isPaid && !confirmPay && (
        <button
          onClick={()=>setConfirmPay(true)}
          style={{
            width:"100%", marginTop:12,
            background:"linear-gradient(135deg,#1D4ED8,#7C3AED)",
            color:"#fff", border:"none", borderRadius:10,
            padding:"14px", fontSize:16, fontWeight:900,
            cursor:"pointer", letterSpacing:0.5,
            boxShadow:"0 4px 20px rgba(99,102,241,0.4)",
            animation:"pulse3 2s ease-in-out infinite",
          }}
        >
          💳 결제완료
        </button>
      )}

      {/* 액션 */}
      <div style={{display:"flex",gap:6,marginTop:10,alignItems:"center",flexWrap:"wrap"}}>
        <select className="inp" style={{flex:"1 1 90px",fontSize:12,padding:"6px 8px",minHeight:36}}
          value={order.status} onChange={e=>onStatus(order.id,e.target.value)}>
          {STATUSES.map(s=><option key={s.key} value={s.key}>{s.label}</option>)}
        </select>
        <select className="inp" style={{flex:"1 1 80px",fontSize:12,padding:"6px 8px",minHeight:36}}
          value={order.deliveryType||"직접출고"} onChange={e=>onDelivery(order.id,e.target.value)}>
          {DELIVERY_TYPES.map(d=><option key={d}>{d}</option>)}
        </select>
        <button className="btn btn-slate" style={{padding:"6px 12px",fontSize:12,minHeight:36}} onClick={()=>onEdit(order)}>수정</button>
        {confirmDel
          ? <div style={{display:"flex",gap:4,alignItems:"center"}}>
              <span style={{fontSize:11,color:"#FCA5A5"}}>삭제?</span>
              <button className="btn btn-red" style={{padding:"6px 10px",fontSize:12,minHeight:36}} onClick={()=>onDelete(order.id,order.company)}>확인</button>
              <button className="btn btn-slate" style={{padding:"6px 8px",fontSize:12,minHeight:36}} onClick={()=>setConfirmDel(false)}>취소</button>
            </div>
          : <button className="btn btn-red" style={{padding:"6px 10px",fontSize:12,minHeight:36}} onClick={()=>setConfirmDel(true)}>삭제</button>
        }
      </div>
    </div>
  );
}

/* ══ 칸반 ══ */
function KanbanPanel({orders,onEdit,onDelete,onStatus}) {
  return (
    <div style={{display:"flex",gap:10,overflowX:"auto",paddingBottom:16}}>
      {STATUSES.map(st=>{
        const col = orders.filter(o=>o.status===st.key);
        return (
          <div key={st.key} style={{flex:"0 0 200px"}}>
            <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:8}}>
              <div style={{width:8,height:8,borderRadius:"50%",background:st.color}}/>
              <span style={{fontWeight:700,fontSize:12,color:"#CBD5E1"}}>{st.label}</span>
              <Chip n={col.length}/>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {col.map(o=>(
                <div key={o.id} className="card" style={{padding:10,borderTop:`3px solid ${st.color}`}}>
                  <div style={{fontWeight:700,fontSize:13,color:"#F1F5F9",marginBottom:2}}>{o.company}</div>
                  <div style={{fontSize:11,color:"#CBD5E1",marginBottom:2}}>{o.doorType}</div>
                  <div style={{fontSize:11,color:"#CBD5E1",marginBottom:4}}>{(o.items||[]).map(i=>`${i.color} ${fmtQty(parseFloat(i.qty)||0)}장`).join(" / ")}</div>
                  <div style={{fontSize:11,color:"#94A3B8",fontWeight:700,marginBottom:6}}>총 {fmtQty(totalQty(o.items||[]))}장</div>
                  <DueBadge dueDate={o.dueDate} status={o.status}/>
                  <div style={{display:"flex",gap:4,marginTop:8}}>
                    <button className="btn btn-slate" style={{flex:1,fontSize:11,padding:5}} onClick={()=>onEdit(o)}>수정</button>
                    <select className="inp" style={{flex:1,fontSize:11,padding:"4px"}} value={o.status} onChange={e=>onStatus(o.id,e.target.value)}>
                      {STATUSES.map(s=><option key={s.key} value={s.key}>{s.label}</option>)}
                    </select>
                  </div>
                </div>
              ))}
              {col.length===0&&<div style={{color:"#1E3A5F",fontSize:11,textAlign:"center",padding:"16px 0"}}>없음</div>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ══ 매출 랭킹 패널 ══ */
function RankingPanel({orders}) {
  const now = new Date();
  const [month, setMonth] = useState(`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`);

  const prevMonth = () => {
    const [y,m] = month.split("-").map(Number);
    const d = new Date(y, m-2, 1);
    setMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);
  };
  const nextMonth = () => {
    const [y,m] = month.split("-").map(Number);
    const d = new Date(y, m, 1);
    setMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);
  };
  const isThisMonth = month === `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`;

  // 해당 월 주문 필터
  const monthOrders = orders.filter(o => o.dueDate && o.dueDate.startsWith(month));

  // 업체별 집계
  const rankMap = {};
  for (const o of monthOrders) {
    if (!o.company) continue;
    if (!rankMap[o.company]) rankMap[o.company] = { company:o.company, amt:0, qty:0, count:0 };
    rankMap[o.company].amt  += totalAmt(o.items||[]);
    rankMap[o.company].qty  += totalQty(o.items||[]);
    rankMap[o.company].count += 1;
  }

  const byAmt = Object.values(rankMap).sort((a,b)=>b.amt-a.amt);
  const byQty = Object.values(rankMap).sort((a,b)=>b.qty-a.qty);

  const totalMonthAmt = byAmt.reduce((s,r)=>s+r.amt, 0);
  const totalMonthQty = byAmt.reduce((s,r)=>s+r.qty, 0);

  const medalColor = (i) => i===0?"#FFD700":i===1?"#C0C0C0":i===2?"#CD7F32":"#475569";
  const medal = (i) => i===0?"🥇":i===1?"🥈":i===2?"🥉":`${i+1}`;

  return (
    <div style={{maxWidth:600,margin:"0 auto"}}>
      {/* 월 선택 */}
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:22,letterSpacing:2,color:"#fff"}}>
          🏆 매출 랭킹
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <button className="btn btn-slate" style={{padding:"6px 10px",fontSize:13}} onClick={prevMonth}>◀</button>
          <span style={{color:"#E2E8F0",fontWeight:700,fontSize:14,minWidth:70,textAlign:"center"}}>{month}</span>
          <button className="btn btn-slate" style={{padding:"6px 10px",fontSize:13}} onClick={nextMonth} disabled={isThisMonth}>▶</button>
          {!isThisMonth && <button className="btn btn-blue" style={{fontSize:12,padding:"6px 10px"}} onClick={()=>setMonth(`${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,"0")}`)}>이번달</button>}
        </div>
      </div>

      {/* 합계 */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:16}}>
        {[
          {label:"이달 총 매출",val:`${totalMonthAmt.toLocaleString()}원`,color:"#34D399"},
          {label:"이달 총 장수",val:`${fmtQty(totalMonthQty)}장`,color:"#60A5FA"},
        ].map(s=>(
          <div key={s.label} className="card" style={{padding:"12px 16px",textAlign:"center"}}>
            <div style={{fontSize:11,color:"#475569",marginBottom:4}}>{s.label}</div>
            <div style={{fontSize:20,fontWeight:900,color:s.color}}>{s.val}</div>
          </div>
        ))}
      </div>

      {byAmt.length===0
        ? <div style={{textAlign:"center",padding:"60px 0",color:"#1E3A5F"}}>
            <div style={{fontSize:40}}>📊</div>
            <div style={{marginTop:8,fontSize:13}}>해당 월 데이터가 없습니다</div>
          </div>
        : <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            {/* 금액 랭킹 */}
            <div>
              <div style={{color:"#94A3B8",fontSize:12,fontWeight:700,marginBottom:8}}>💰 매출액 순위</div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {byAmt.map((r,i)=>(
                  <div key={r.company} className="card" style={{padding:"10px 12px",borderLeft:`3px solid ${medalColor(i)}`}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <span style={{fontSize:16,minWidth:24}}>{medal(i)}</span>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{color:"#F1F5F9",fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.company}</div>
                        <div style={{color:"#34D399",fontWeight:800,fontSize:13}}>{r.amt.toLocaleString()}원</div>
                        <div style={{color:"#475569",fontSize:11}}>{fmtQty(r.qty)}장 · {r.count}건</div>
                      </div>
                    </div>
                    {/* 비율 바 */}
                    <div style={{marginTop:6,height:3,background:"#1E3A5F",borderRadius:2}}>
                      <div style={{height:"100%",borderRadius:2,background:medalColor(i),width:`${totalMonthAmt>0?(r.amt/totalMonthAmt*100).toFixed(1):0}%`,transition:"width .5s"}}/>
                    </div>
                  </div>
                ))}
              </div>
            </div>
            {/* 장수 랭킹 */}
            <div>
              <div style={{color:"#94A3B8",fontSize:12,fontWeight:700,marginBottom:8}}>📐 장수 순위</div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {byQty.map((r,i)=>(
                  <div key={r.company} className="card" style={{padding:"10px 12px",borderLeft:`3px solid ${medalColor(i)}`}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <span style={{fontSize:16,minWidth:24}}>{medal(i)}</span>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{color:"#F1F5F9",fontWeight:700,fontSize:13,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.company}</div>
                        <div style={{color:"#60A5FA",fontWeight:800,fontSize:13}}>{fmtQty(r.qty)}장</div>
                        <div style={{color:"#475569",fontSize:11}}>{r.amt.toLocaleString()}원 · {r.count}건</div>
                      </div>
                    </div>
                    <div style={{marginTop:6,height:3,background:"#1E3A5F",borderRadius:2}}>
                      <div style={{height:"100%",borderRadius:2,background:medalColor(i),width:`${totalMonthQty>0?(r.qty/totalMonthQty*100).toFixed(1):0}%`,transition:"width .5s"}}/>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
      }
    </div>
  );
}

/* ══ 폼 ══ */
function FormPanel({form,setForm,onSave,onCancel,isEdit,preview}) {
  const f=(k,v)=>setForm(p=>({...p,[k]:v}));
  function setItem(idx,key,val){ f("items",form.items.map((it,i)=>i===idx?{...it,[key]:val}:it)); }
  function addItem()    { f("items",[...form.items,{...EMPTY_ITEM}]); }
  function removeItem(idx){ if(form.items.length===1) return; f("items",form.items.filter((_,i)=>i!==idx)); }

  return (
    <div style={{maxWidth:600,margin:"0 auto"}}>
      <div style={{marginBottom:14}}>
        <div style={{fontFamily:"'Bebas Neue',sans-serif",fontSize:22,letterSpacing:2,color:"#fff"}}>
          {isEdit?"주문 수정":preview?"견적서 → 주문등록":"신규 주문 등록"}
        </div>
      </div>
      {preview && (
        <div style={{background:"#0a2540",border:"1px solid #1D4ED8",borderRadius:10,padding:"10px 14px",marginBottom:12}}>
          <div style={{color:"#60A5FA",fontWeight:700,fontSize:12}}>📎 견적서에서 자동 추출 — 확인 후 등록하세요</div>
          <div style={{color:"#475569",fontSize:11,marginTop:3}}>
            {preview.company} · {preview.memo} · {preview.items.length}항목 · 총 {fmtQty(totalQty(preview.items))}장
            {totalAmt(preview.items)>0 && <span style={{color:"#34D399"}}> · {totalAmt(preview.items).toLocaleString()}원</span>}
          </div>
        </div>
      )}
      <div className="card" style={{padding:16,display:"flex",flexDirection:"column",gap:14}}>
        <Field label="거래처명 *">
          <input className="inp" value={form.company} onChange={e=>f("company",e.target.value)} placeholder="예: 윤퍼니"/>
        </Field>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
          <Field label="도어 브랜드">
            <select className="inp" value={form.doorType} onChange={e=>f("doorType",e.target.value)}>
              {DOOR_TYPES.map(d=><option key={d}>{d}</option>)}
            </select>
          </Field>
          <Field label="출고방식">
            <select className="inp" value={form.deliveryType||"직접출고"} onChange={e=>f("deliveryType",e.target.value)}>
              {DELIVERY_TYPES.map(d=><option key={d}>{d}</option>)}
            </select>
          </Field>
        </div>
        <Field label="납기일 *">
          <input className="inp" type="date" value={form.dueDate} onChange={e=>f("dueDate",e.target.value)}/>
        </Field>

        {/* 색상별 행 */}
        <Field label="색상별 장수 · 임가공비 *">
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {form.items.map((item,idx)=>{
              const isCustom = !COLOR_PRESETS.slice(0,-1).includes(item.color);
              return (
                <div key={idx} style={{background:"#0C1220",borderRadius:8,padding:"10px 10px",display:"flex",flexDirection:"column",gap:8,border:"1px solid #1E3A5F"}}>
                  <div style={{display:"flex",gap:8,alignItems:"center"}}>
                    <select className="inp" style={{flex:1,fontSize:13}}
                      value={isCustom?"직접입력":item.color}
                      onChange={e=>{ if(e.target.value==="직접입력") setItem(idx,"color",""); else setItem(idx,"color",e.target.value); }}>
                      {COLOR_PRESETS.map(c=><option key={c}>{c}</option>)}
                    </select>
                    <button onClick={()=>removeItem(idx)} style={{background:"none",border:"none",color:"#475569",fontSize:20,cursor:"pointer",padding:"0 4px",flexShrink:0}}>×</button>
                  </div>
                  {isCustom && <input className="inp" style={{fontSize:13}} placeholder="색상명 직접입력" value={item.color} onChange={e=>setItem(idx,"color",e.target.value)}/>}
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                    <Field label="장수">
                      <input className="inp" type="number" min={0} step={0.1} placeholder="0.0" value={item.qty} onChange={e=>setItem(idx,"qty",e.target.value)} style={{textAlign:"right"}}/>
                    </Field>
                    <Field label="장당 임가공(원)">
                      <input className="inp" type="number" min={0} step={1000} placeholder="0" value={item.unitPrice} onChange={e=>setItem(idx,"unitPrice",e.target.value)} style={{textAlign:"right"}}/>
                    </Field>
                  </div>
                </div>
              );
            })}
            <div style={{display:"flex",alignItems:"center",gap:12,marginTop:2,flexWrap:"wrap"}}>
              <button className="btn btn-slate" style={{fontSize:12,padding:"8px 14px"}} onClick={addItem}>＋ 색상 추가</button>
              {totalQty(form.items)>0 && (
                <span style={{fontSize:13,color:"#94A3B8"}}>
                  합계 <span style={{color:"#F1F5F9",fontWeight:800}}>{fmtQty(totalQty(form.items))}</span>장
                  {totalAmt(form.items)>0 && <span style={{color:"#34D399",marginLeft:8,fontWeight:700}}>{totalAmt(form.items).toLocaleString()}원</span>}
                </span>
              )}
            </div>
          </div>
        </Field>

        <Field label="진행상태">
          <select className="inp" value={form.status} onChange={e=>f("status",e.target.value)}>
            {STATUSES.map(s=><option key={s.key} value={s.key}>{s.icon} {s.label}</option>)}
          </select>
        </Field>
        <Field label="메모">
          <textarea className="inp" rows={3} value={form.memo} onChange={e=>f("memo",e.target.value)} placeholder="현장명·특이사항 등" style={{resize:"vertical"}}/>
        </Field>
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <button className="btn btn-slate" onClick={onCancel}>취소</button>
          <button className="btn btn-blue"  onClick={onSave}>{isEdit?"수정 저장":"등록하기"}</button>
        </div>
      </div>
    </div>
  );
}
