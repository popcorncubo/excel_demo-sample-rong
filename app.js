const CONFIG = {
  contactEmail: "chinsun1345@gmail.com",
  storageKey: "stocklite-demo-v1"
};

const DEMO_PRODUCTS = [
  { sku:"P001", name:"無線滑鼠", category:"電腦周邊", price:690, cost:320, stock:34, safety:10, supplier:"星河科技" },
  { sku:"P002", name:"機械鍵盤", category:"電腦周邊", price:1890, cost:980, stock:8, safety:10, supplier:"星河科技" },
  { sku:"P003", name:"USB-C Hub 8合1", category:"轉接設備", price:1490, cost:760, stock:16, safety:5, supplier:"晨光電子" },
  { sku:"P004", name:"24吋 IPS 螢幕", category:"顯示設備", price:3990, cost:2760, stock:6, safety:3, supplier:"宏景資訊" },
  { sku:"P005", name:"1080P 網路攝影機", category:"影音設備", price:1290, cost:650, stock:4, safety:6, supplier:"北辰數位" },
  { sku:"P006", name:"Cat6 網路線 3M", category:"網路設備", price:199, cost:62, stock:55, safety:20, supplier:"信達線材" },
  { sku:"P007", name:"Wi-Fi 6 路由器", category:"網路設備", price:2490, cost:1510, stock:7, safety:5, supplier:"信達網通" },
  { sku:"P008", name:"1TB NVMe SSD", category:"儲存設備", price:2390, cost:1760, stock:3, safety:5, supplier:"全速儲存" }
];

const HEADER_ALIASES = {
  sku:["商品編號","sku","品號","產品編號","編號"],
  name:["商品名稱","產品名稱","名稱","品名"],
  category:["類別","分類","category"],
  price:["售價","銷售價格","價格","price"],
  cost:["成本","進價","成本價","cost"],
  stock:["庫存","庫存量","數量","stock"],
  safety:["安全庫存","最低庫存","安全量","safety stock"],
  supplier:["供應商","廠商","supplier"]
};

let state = { products:[], records:[], role:"admin" };

const $ = (s) => document.querySelector(s);
const els = {
  excelInput:$("#excelInput"), loadDemo:$("#loadDemo"), roleSelect:$("#roleSelect"),
  searchInput:$("#searchInput"), categoryFilter:$("#categoryFilter"), stockFilter:$("#stockFilter"),
  productBody:$("#productBody"), recordBody:$("#recordBody"), emptyProducts:$("#emptyProducts"), emptyRecords:$("#emptyRecords"),
  kpiSku:$("#kpiSku"), kpiStock:$("#kpiStock"), kpiCost:$("#kpiCost"), kpiLow:$("#kpiLow"),
  categoryReport:$("#categoryReport"), lowStockReport:$("#lowStockReport"), todayTransactions:$("#todayTransactions"),
  exportProducts:$("#exportProducts"), exportRecords:$("#exportRecords"), exportReport:$("#exportReport"), clearData:$("#clearData"),
  contactTop:$("#contactTop"), contactBottom:$("#contactBottom"), toast:$("#toast")
};

function normalizeHeader(value){ return String(value ?? "").trim().toLowerCase(); }
function num(value){ const n = Number(String(value ?? "").replace(/,/g,"")); return Number.isFinite(n) ? n : 0; }
function money(value){ return new Intl.NumberFormat("zh-TW",{style:"currency",currency:"TWD",maximumFractionDigits:0}).format(value||0); }
function esc(value){ return String(value ?? "").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#039;"}[m])); }
function nowString(){ return new Date().toLocaleString("zh-TW",{hour12:false}); }
function todayKey(){ const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; }
function notify(msg){ els.toast.textContent=msg; els.toast.classList.add("show"); clearTimeout(notify.t); notify.t=setTimeout(()=>els.toast.classList.remove("show"),2200); }

function save(){ localStorage.setItem(CONFIG.storageKey, JSON.stringify(state)); }
function loadSaved(){
  try{
    const saved = JSON.parse(localStorage.getItem(CONFIG.storageKey));
    if(saved && Array.isArray(saved.products) && Array.isArray(saved.records)) state = {...state,...saved};
  }catch(_){ }
  els.roleSelect.value = state.role || "admin";
}

function matchKey(header){
  const h=normalizeHeader(header);
  return Object.keys(HEADER_ALIASES).find(k=>HEADER_ALIASES[k].some(a=>normalizeHeader(a)===h));
}

function rowsToProducts(rows){
  if(!rows.length) return [];
  const headers=rows[0];
  const map=headers.map(matchKey);
  if(!map.includes("sku") || !map.includes("name")) throw new Error("Excel 至少需要「商品編號」與「商品名稱」欄位。");
  return rows.slice(1).filter(r=>r.some(v=>String(v??"").trim()!=="")).map((r,i)=>{
    const p={sku:"",name:"",category:"未分類",price:0,cost:0,stock:0,safety:0,supplier:""};
    map.forEach((key,idx)=>{
      if(!key) return;
      const v=r[idx];
      if(["price","cost","stock","safety"].includes(key)) p[key]=num(v); else p[key]=String(v??"").trim();
    });
    if(!p.sku) p.sku=`ITEM-${String(i+1).padStart(3,"0")}`;
    return p;
  });
}

async function importExcel(file){
  if(typeof XLSX === "undefined") throw new Error("Excel 元件載入失敗，請確認目前有網路連線。");
  const buffer=await file.arrayBuffer();
  const wb=XLSX.read(buffer,{type:"array"});
  const ws=wb.Sheets[wb.SheetNames[0]];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
  const products=rowsToProducts(rows);
  if(!products.length) throw new Error("沒有讀到商品資料。");
  state.products=products;
  state.records=[];
  save(); renderAll(); notify(`已匯入 ${products.length} 筆商品`);
}

function statusOf(p){ if(p.stock<=0) return ["已售罄","zero"]; if(p.stock<=p.safety) return ["低庫存","low"]; return ["正常","ok"]; }
function filteredProducts(){
  const q=els.searchInput.value.trim().toLowerCase(), cat=els.categoryFilter.value, sf=els.stockFilter.value;
  return state.products.filter(p=>{
    const [_,s]=statusOf(p);
    const hay=[p.sku,p.name,p.category,p.supplier].join(" ").toLowerCase();
    return (!q||hay.includes(q)) && (!cat||p.category===cat) && (!sf||s===sf);
  });
}

function renderProducts(){
  const rows=filteredProducts();
  els.productBody.innerHTML=rows.map(p=>{
    const [label,s]=statusOf(p); const disabled=state.role!=="admin"?"disabled":"";
    return `<tr>
      <td><strong>${esc(p.sku)}</strong></td><td>${esc(p.name)}</td><td>${esc(p.category)}</td>
      <td class="num">${money(p.price)}</td><td class="num">${money(p.cost)}</td><td class="num"><strong>${p.stock}</strong></td><td class="num">${p.safety}</td>
      <td>${esc(p.supplier)}</td><td><span class="badge badge-${s}">${label}</span></td>
      <td><div class="action-group"><button class="action-btn in" data-act="in" data-sku="${esc(p.sku)}" ${disabled}>+ 入庫</button><button class="action-btn out" data-act="out" data-sku="${esc(p.sku)}" ${disabled}>− 出庫</button></div></td>
    </tr>`;
  }).join("");
  els.emptyProducts.classList.toggle("hidden",rows.length>0);
}

function renderRecords(){
  els.recordBody.innerHTML=state.records.slice().reverse().map(r=>`<tr>
    <td>${esc(r.time)}</td><td><span class="badge ${r.type==="入庫"?"badge-ok":"badge-low"}">${r.type}</span></td>
    <td>${esc(r.sku)}</td><td>${esc(r.name)}</td><td class="num">${r.qty}</td><td class="num">${r.after}</td><td>${esc(r.note||"")}</td>
  </tr>`).join("");
  els.emptyRecords.classList.toggle("hidden",state.records.length>0);
}

function renderFilters(){
  const current=els.categoryFilter.value;
  const cats=[...new Set(state.products.map(p=>p.category||"未分類"))].sort();
  els.categoryFilter.innerHTML=`<option value="">全部類別</option>`+cats.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join("");
  if(cats.includes(current)) els.categoryFilter.value=current;
}

function renderDashboard(){
  const totalStock=state.products.reduce((s,p)=>s+p.stock,0);
  const totalCost=state.products.reduce((s,p)=>s+p.stock*p.cost,0);
  const low=state.products.filter(p=>p.stock<=p.safety).length;
  els.kpiSku.textContent=state.products.length;
  els.kpiStock.textContent=new Intl.NumberFormat("zh-TW").format(totalStock);
  els.kpiCost.textContent=money(totalCost);
  els.kpiLow.textContent=low;
}

function renderReports(){
  const byCat={}; state.products.forEach(p=>byCat[p.category]=(byCat[p.category]||0)+p.stock);
  const entries=Object.entries(byCat).sort((a,b)=>b[1]-a[1]); const max=Math.max(...entries.map(x=>x[1]),1);
  els.categoryReport.innerHTML=entries.length?entries.map(([c,v])=>`<div class="bar-row"><span title="${esc(c)}">${esc(c)}</span><div class="bar-track"><div class="bar-fill" style="width:${Math.round(v/max*100)}%"></div></div><span class="bar-val">${v}</span></div>`).join(""):`<div class="empty-state">尚無資料</div>`;
  const low=state.products.filter(p=>p.stock<=p.safety).sort((a,b)=>a.stock-b.stock);
  els.lowStockReport.innerHTML=low.length?low.map(p=>`<div class="low-item"><strong>${esc(p.name)} <small>(${esc(p.sku)})</small></strong><span>${p.stock} / 安全 ${p.safety}</span></div>`).join(""):`<div class="empty-state">目前沒有低庫存商品</div>`;
  const key=todayKey();
  els.todayTransactions.textContent=state.records.filter(r=>r.dateKey===key).length;
}

function renderAll(){ renderFilters(); renderDashboard(); renderProducts(); renderRecords(); renderReports(); }

function adjustStock(sku,type){
  if(state.role!=="admin") return notify("檢視者模式不可調整庫存");
  const p=state.products.find(x=>x.sku===sku); if(!p) return;
  const raw=prompt(`${p.name}｜${type}\n請輸入數量：`,"1"); if(raw===null) return;
  const qty=Math.floor(num(raw)); if(qty<=0) return notify("請輸入大於 0 的整數");
  if(type==="出庫" && qty>p.stock) return notify("出庫數量不能大於目前庫存");
  const note=prompt("備註（可留空）：","") ?? "";
  p.stock += type==="入庫" ? qty : -qty;
  state.records.push({time:nowString(),dateKey:todayKey(),type,sku:p.sku,name:p.name,qty,after:p.stock,note});
  save(); renderAll(); notify(`${type}完成：${p.name} ${qty} 件`);
}

function productsForExport(){ return state.products.map(p=>({"商品編號":p.sku,"商品名稱":p.name,"類別":p.category,"售價":p.price,"成本":p.cost,"庫存":p.stock,"安全庫存":p.safety,"供應商":p.supplier})); }
function recordsForExport(){ return state.records.map(r=>({"時間":r.time,"類型":r.type,"商品編號":r.sku,"商品名稱":r.name,"數量":r.qty,"調整後庫存":r.after,"備註":r.note})); }
function downloadWorkbook(sheets,filename){
  if(typeof XLSX === "undefined") return notify("Excel 元件未載入");
  const wb=XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name,data])=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data),name));
  XLSX.writeFile(wb,filename);
}

function exportSummary(){
  const low=state.products.filter(p=>p.stock<=p.safety);
  const totalStock=state.products.reduce((s,p)=>s+p.stock,0); const totalCost=state.products.reduce((s,p)=>s+p.stock*p.cost,0);
  const summary=[
    {"指標":"商品品項","數值":state.products.length},
    {"指標":"總庫存","數值":totalStock},
    {"指標":"庫存成本","數值":totalCost},
    {"指標":"低庫存品項","數值":low.length},
    {"指標":"出入庫紀錄","數值":state.records.length}
  ];
  downloadWorkbook({"商品資料":productsForExport(),"出入庫紀錄":recordsForExport(),"報表摘要":summary},`StockLite_管理報表_${todayKey()}.xlsx`);
}

function contact(){
  const subject=encodeURIComponent("系統客製開發需求");
  const body=encodeURIComponent("您好，我想詢問系統客製開發。\n\n需求概述：\n使用人數：\n希望功能：\n預計時程：\n");
  location.href=`mailto:${CONFIG.contactEmail}?subject=${subject}&body=${body}`;
}

els.excelInput.addEventListener("change", async e=>{
  const file=e.target.files?.[0]; if(!file) return;
  try{ await importExcel(file); }catch(err){ alert(err.message||"Excel 讀取失敗"); }
  e.target.value="";
});
els.loadDemo.addEventListener("click",()=>{ state.products=structuredClone(DEMO_PRODUCTS); state.records=[]; save(); renderAll(); notify("已載入範例資料"); });
els.roleSelect.addEventListener("change",()=>{ state.role=els.roleSelect.value; save(); renderProducts(); notify(state.role==="admin"?"已切換為管理者展示":"已切換為檢視者展示"); });
[els.searchInput,els.categoryFilter,els.stockFilter].forEach(el=>el.addEventListener(el.tagName==="INPUT"?"input":"change",renderProducts));
els.productBody.addEventListener("click",e=>{ const b=e.target.closest("button[data-act]"); if(!b) return; adjustStock(b.dataset.sku,b.dataset.act==="in"?"入庫":"出庫"); });
els.exportProducts.addEventListener("click",()=>downloadWorkbook({"商品資料":productsForExport()},`商品資料_${todayKey()}.xlsx`));
els.exportRecords.addEventListener("click",()=>downloadWorkbook({"出入庫紀錄":recordsForExport()},`出入庫紀錄_${todayKey()}.xlsx`));
els.exportReport.addEventListener("click",exportSummary);
els.clearData.addEventListener("click",()=>{ if(!confirm("確定要清除這個瀏覽器中的商品與出入庫資料？")) return; localStorage.removeItem(CONFIG.storageKey); state={products:[],records:[],role:"admin"}; els.roleSelect.value="admin"; renderAll(); notify("本機資料已清除"); });
els.contactTop.addEventListener("click",contact); els.contactBottom.addEventListener("click",contact);

loadSaved();
if(!state.products.length) state.products=structuredClone(DEMO_PRODUCTS);
renderAll();
