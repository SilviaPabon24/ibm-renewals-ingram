"""
Portal de Renovaciones IBM S&S — Ingram Micro
Layout master-detail: sidebar con lista de quotes + panel de detalle
"""
 
PORTAL_HTML = r'''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Renovaciones IBM S&S — Ingram Micro</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#F0F4F8;color:#1A2B3C;display:flex;flex-direction:column}
 
/* Topbar */
.topbar{background:#fff;border-bottom:0.5px solid #E2E8F0;padding:0 1rem;height:44px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.logo{font-size:14px;font-weight:600;color:#0066CC}
.ibm-badge{background:#054ADA;color:#fff;font-size:10px;font-weight:600;padding:2px 9px;border-radius:5px}
 
/* Steps */
.steps{background:#fff;border-bottom:0.5px solid #E2E8F0;padding:0 1rem;display:flex;align-items:center;height:36px;flex-shrink:0;gap:4px}
.step{display:flex;align-items:center;gap:5px;font-size:11px;color:#94A3B8}
.step.active{color:#0066CC}
.step.done{color:#16A34A}
.sn{width:17px;height:17px;border-radius:50%;background:#E2E8F0;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#64748B}
.step.active .sn{background:#0066CC;color:#fff}
.step.done .sn{background:#16A34A;color:#fff}
.sarr{color:#CBD5E1;font-size:11px;margin:0 4px}
 
/* Upload screen */
.upload-screen{flex:1;display:flex;align-items:center;justify-content:center;padding:2rem}
.upload-card{background:#fff;border:0.5px solid #E2E8F0;border-radius:14px;padding:2rem;max-width:480px;width:100%;text-align:center}
.dropzone{border:1.5px dashed #CBD5E1;border-radius:10px;padding:2rem 1rem;cursor:pointer;position:relative;transition:border-color 0.15s,background 0.15s;margin:1rem 0}
.dropzone:hover,.dropzone.drag{border-color:#0066CC;background:#EBF3FF}
.dropzone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
 
/* Main layout */
.main-layout{flex:1;display:grid;grid-template-columns:280px 1fr;overflow:hidden}
.main-layout.hidden{display:none!important}
 
/* Sidebar */
.sidebar{background:#fff;border-right:0.5px solid #E2E8F0;display:flex;flex-direction:column;overflow:hidden}
.sb-head{padding:10px 12px;border-bottom:0.5px solid #E2E8F0;flex-shrink:0}
.sb-title{font-size:13px;font-weight:600;color:#0F2B5B}
.sb-meta{font-size:11px;color:#64748B;margin-top:2px}
 
/* Params bar */
.params-bar{padding:8px 12px;border-bottom:0.5px solid #E2E8F0;background:#F8FAFC;flex-shrink:0}
.params-top{display:flex;align-items:center;justify-content:space-between;cursor:pointer}
.params-pills{display:flex;gap:4px;flex-wrap:wrap}
.pill{background:#fff;border:0.5px solid #E2E8F0;border-radius:20px;padding:2px 7px;font-size:10px;color:#64748B}
.pill strong{color:#1A2B3C}
.edit-lnk{font-size:10px;color:#0066CC;font-weight:600;white-space:nowrap;flex-shrink:0}
.params-panel{display:none;padding-top:10px}
.params-panel.open{display:block}
.params-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.pfield label{display:block;font-size:10px;font-weight:600;color:#64748B;text-transform:uppercase;letter-spacing:0.3px;margin-bottom:3px}
.pfield .sw{position:relative}
.pfield input{width:100%;height:30px;border:0.5px solid #D1D5DB;border-radius:6px;padding:0 22px 0 8px;font-size:12px;color:#1A2B3C;outline:none}
.pfield input:focus{border-color:#0066CC}
.pfield .sfx{position:absolute;right:7px;top:50%;transform:translateY(-50%);font-size:11px;color:#94A3B8;pointer-events:none}
.params-actions{display:flex;gap:6px;margin-top:8px}
 
/* Quote list */
.q-list{overflow-y:auto;flex:1}
.q-item{display:grid;grid-template-columns:38px 1fr auto;align-items:center;padding:9px 12px;border-bottom:0.5px solid #F1F5F9;cursor:pointer;transition:background 0.1s;border-left:3px solid transparent}
.q-item:hover{background:#F8FAFC}
.q-item.sel{background:#EBF3FF;border-left-color:#0066CC}
.q-item.ready{border-left-color:#16A34A}
.q-item.ready.sel{background:#F0FDF4}
.q-item.has-warn{border-left-color:#F59E0B}
.chk{width:18px;height:18px;border-radius:4px;border:1.5px solid #CBD5E1;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#16A34A;transition:all 0.15s;flex-shrink:0}
.chk.on{background:#16A34A;border-color:#16A34A;color:#fff}
.qi-body{margin-left:8px;min-width:0}
.qi-num{font-size:12px;font-weight:600;color:#0F2B5B;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.qi-cust{font-size:10px;color:#64748B;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.qi-right{text-align:right;flex-shrink:0}
.qi-amt{font-size:11px;font-weight:600;color:#0F2B5B}
.qtag{font-size:9px;font-weight:600;padding:1px 6px;border-radius:20px;margin-top:2px;display:inline-block}
.qtag-ok{background:#DCFCE7;color:#166534}
.qtag-warn{background:#FEF3C7;color:#92400E}
.qtag-gray{background:#F1F5F9;color:#64748B}
.qtag-blue{background:#EBF3FF;color:#0066CC}
 
/* Bulk footer */
.bulk-footer{border-top:0.5px solid #E2E8F0;padding:10px 12px;background:#F0FDF4;flex-shrink:0}
.bulk-footer.hidden{display:none!important}
.bf-row{display:flex;align-items:center;justify-content:space-between;gap:8px}
.bf-info{font-size:11px;font-weight:600;color:#166534}
.bf-sub{font-size:10px;color:#16A34A;margin-top:1px}
 
/* Detail panel */
.detail{overflow-y:auto;padding:1.25rem;background:#F0F4F8}
.detail-empty{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#94A3B8;gap:8px}
.detail-empty .icon{font-size:36px}
 
.dcard{background:#fff;border:0.5px solid #E2E8F0;border-radius:12px;padding:1rem 1.25rem;margin-bottom:10px}
 
/* Detail header */
.d-title{font-size:15px;font-weight:600;color:#0F2B5B}
.d-sub{font-size:12px;color:#64748B;margin-top:3px}
 
/* Params in detail */
.dp-head{font-size:11px;font-weight:600;color:#64748B;text-transform:uppercase;letter-spacing:0.3px;display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.dp-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:10px}
.dpf label{display:block;font-size:10px;color:#64748B;margin-bottom:3px}
.dpf .sw{position:relative}
.dpf input{width:100%;height:30px;border:0.5px solid #D1D5DB;border-radius:6px;padding:0 22px 0 8px;font-size:12px;color:#1A2B3C;outline:none}
.dpf input:focus{border-color:#0066CC}
.dpf .sfx{position:absolute;right:7px;top:50%;transform:translateY(-50%);font-size:11px;color:#94A3B8;pointer-events:none}
.toggles{display:flex;gap:16px;font-size:12px;color:#1A2B3C}
.toggle{display:flex;align-items:center;gap:5px;cursor:pointer}
.toggle input{width:13px;height:13px}
 
/* Summary cards */
.scards{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:10px}
.sc{background:#F8FAFC;border-radius:8px;padding:9px 12px}
.sc-lbl{font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:0.3px;font-weight:600;margin-bottom:3px}
.sc-val{font-size:17px;font-weight:700;color:#0F2B5B}
.sc-val.g{color:#16A34A}
.sc-val.b{color:#0066CC}
 
/* Coverage */
.cov-lines{display:flex;flex-direction:column;gap:4px;margin-bottom:10px}
.cl{display:flex;align-items:center;gap:7px;padding:7px 11px;border-radius:7px;font-size:12px}
.cl-ok{background:#F0FDF4;border:0.5px solid #BBF7D0;color:#166534}
.cl-warn{background:#FFFBEB;border:0.5px solid #FCD34D;color:#92400E}
 
/* Table */
.tbl-wrap{overflow-x:auto;margin-bottom:10px}
table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#0F2B5B;color:#fff;padding:7px 10px;text-align:left;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.3px}
td{padding:7px 10px;border-bottom:0.5px solid #F1F5F9;vertical-align:middle}
tr:nth-child(even) td{background:#F8FAFC}
.tot-row td{background:#EBF3FF!important;font-weight:700;color:#0C447C}
 
/* Buttons */
.btn{display:inline-flex;align-items:center;gap:6px;height:34px;padding:0 14px;border-radius:8px;border:none;cursor:pointer;font-size:12px;font-weight:600;transition:all 0.15s}
.btn:active{transform:scale(0.97)}
.btn-p{background:#0066CC;color:#fff}.btn-p:hover{background:#0052A3}
.btn-p:disabled{background:#93C5FD;cursor:not-allowed}
.btn-s{background:#16A34A;color:#fff}.btn-s:hover{background:#15803D}
.btn-s:disabled{background:#86EFAC;cursor:not-allowed}
.btn-a{background:#D97706;color:#fff}.btn-a:hover{background:#B45309}
.btn-g{background:#fff;color:#64748B;border:0.5px solid #E2E8F0}.btn-g:hover{background:#F8FAFC}
.btn-sm{height:28px;padding:0 10px;font-size:11px}
.btn-grp{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
 
/* Review */
.review-box{display:flex;align-items:center;gap:8px;padding:10px 14px;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;margin-top:10px}
.review-box label{display:flex;align-items:center;gap:7px;cursor:pointer;font-size:13px;font-weight:600;color:#166534}
.review-box input[type=checkbox]{width:16px;height:16px;accent-color:#16A34A;cursor:pointer}
.review-hint{font-size:11px;color:#64748B;margin-left:auto}
 
/* Audit */
.audit-btn{width:100%;display:flex;align-items:center;justify-content:space-between;padding:8px 12px;background:#F8FAFC;border:0.5px solid #E2E8F0;border-radius:7px;cursor:pointer;font-size:12px;color:#64748B;margin-top:10px;border-top:none;border-left:none;border-right:none;text-align:left}
.audit-wrap{border:0.5px solid #E2E8F0;border-radius:8px;margin-top:8px;overflow:hidden;display:none}
.audit-wrap.open{display:block}
.audit-inner{overflow-x:auto}
.ath{background:#475569;color:#fff;padding:5px 8px;font-size:10px;font-weight:600;text-align:left;white-space:nowrap}
.ath2{background:#0066CC;color:#fff;padding:5px 8px;font-size:10px;font-weight:600;text-align:left;white-space:nowrap}
.atd{padding:5px 8px;border-bottom:0.5px solid #E2E8F0;font-size:11px;white-space:nowrap;background:#F8FAFC}
.atd2{padding:5px 8px;border-bottom:0.5px solid #E2E8F0;font-size:11px;white-space:nowrap;background:#EBF3FF}
.asec{background:#1E293B;color:#fff;padding:5px 10px;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px}
 
.alert{border-radius:8px;padding:10px 14px;font-size:12px;margin-top:8px}
.a-ok{background:#F0FDF4;color:#166534;border:1px solid #BBF7D0}
.a-warn{background:#FFFBEB;color:#92400E;border:1px solid #FCD34D}
.a-err{background:#FEF2F2;color:#991B1B;border:1px solid #FECACA}
 
.spin{width:13px;height:13px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;border-radius:50%;animation:sp 0.7s linear infinite;display:inline-block}
@keyframes sp{to{transform:rotate(360deg)}}
.hidden{display:none!important}
</style>
</head>
<body>
 
<div class="topbar">
  <div style="display:flex;align-items:center;gap:8px">
    <span class="logo">Ingram Micro</span>
    <span style="width:1px;height:14px;background:#E2E8F0;display:inline-block"></span>
    <span style="font-size:12px;color:#64748B">Renovaciones IBM S&amp;S</span>
  </div>
  <span class="ibm-badge">IBM Authorized Reseller</span>
</div>
 
<div class="steps">
  <div class="step active" id="st1"><div class="sn">1</div>Cargar XML</div>
  <span class="sarr">›</span>
  <div class="step" id="st2"><div class="sn">2</div>Revisar quotes</div>
  <span class="sarr">›</span>
  <div class="step" id="st3"><div class="sn">3</div>Descargar todo</div>
</div>
 
<!-- UPLOAD -->
<div class="upload-screen" id="upload-screen">
  <div class="upload-card">
    <div style="font-size:16px;font-weight:600;color:#0F2B5B;margin-bottom:6px">Cargar reporte de renovaciones</div>
    <div style="font-size:13px;color:#64748B;margin-bottom:16px">Sube el archivo XML de IBM Passport Advantage para procesar los quotes</div>
    <div class="dropzone" id="dropzone">
      <input type="file" id="xml-file" accept=".xml">
      <div style="font-size:32px;margin-bottom:8px">📄</div>
      <div style="font-size:14px;font-weight:600;margin-bottom:4px">Arrastra el XML aquí</div>
      <div style="font-size:12px;color:#64748B">o haz clic para seleccionarlo</div>
    </div>
    <div id="upload-status" class="hidden"></div>
  </div>
</div>
 
<!-- MAIN LAYOUT -->
<div class="main-layout hidden" id="main-layout">
 
  <!-- SIDEBAR -->
  <div class="sidebar">
    <div class="sb-head">
      <div class="sb-title" id="sb-title">Quotes encontrados</div>
      <div class="sb-meta" id="sb-meta">Selecciona uno para ver el detalle</div>
    </div>
 
    <div class="params-bar">
      <div class="params-top" onclick="toggleParams()">
        <div class="params-pills" id="params-pills"></div>
        <span class="edit-lnk" id="params-lnk">✎ Editar</span>
      </div>
      <div class="params-panel" id="params-panel">
        <div class="params-grid">
          <div class="pfield"><label>Margen BP</label><div class="sw"><input type="number" id="g-bp" value="3" min="0" max="100" step="0.1" oninput="updatePills()"><span class="sfx">%</span></div></div>
          <div class="pfield"><label>Margen Ingram</label><div class="sw"><input type="number" id="g-ingram" value="2.1" min="0" max="100" step="0.1"><span class="sfx">%</span></div></div>
          <div class="pfield"><label>CVR Renew BP</label><div class="sw"><input type="number" id="g-cvr-renew" value="2" min="0" max="100" step="0.1" oninput="updatePills()"><span class="sfx">%</span></div></div>
          <div class="pfield"><label>CVR Renew Ingram</label><div class="sw"><input type="number" id="g-cvr-renew-ing" value="2" min="0" max="100" step="0.1"><span class="sfx">%</span></div></div>
          <div class="pfield"><label>CVR OnTime BP</label><div class="sw"><input type="number" id="g-cvr-ontime" value="2" min="0" max="100" step="0.1" oninput="updatePills()"><span class="sfx">%</span></div></div>
          <div class="pfield"><label>CVR OnTime Ingram</label><div class="sw"><input type="number" id="g-cvr-ontime-ing" value="1" min="0" max="100" step="0.1"><span class="sfx">%</span></div></div>
        </div>
        <div class="params-actions">
          <button class="btn btn-p btn-sm" onclick="applyParams()">Aplicar</button>
          <button class="btn btn-g btn-sm" onclick="toggleParams()">Cerrar</button>
        </div>
      </div>
    </div>
 
    <div class="q-list" id="q-list"></div>
 
    <div class="bulk-footer hidden" id="bulk-footer">
      <div class="bf-row">
        <div>
          <div class="bf-info" id="bf-info">0 quotes listos</div>
          <div class="bf-sub" id="bf-sub"></div>
        </div>
        <button class="btn btn-s btn-sm" id="btn-bulk" onclick="downloadBulk()">⬇ ZIP</button>
      </div>
    </div>
  </div>
 
  <!-- DETAIL -->
  <div class="detail" id="detail">
    <div class="detail-empty" id="detail-empty">
      <div class="icon">📋</div>
      <div style="font-size:14px;font-weight:600;color:#64748B">Selecciona un quote</div>
      <div style="font-size:12px;color:#94A3B8">Haz clic en cualquier quote de la lista para ver su detalle</div>
    </div>
 
    <div id="detail-content" class="hidden">
 
      <div class="dcard">
        <div class="d-title" id="d-title">Quote</div>
        <div class="d-sub" id="d-sub"></div>
      </div>
 
      <div class="dcard">
        <div class="dp-head">
          <span>Parámetros de este quote</span>
          <span id="rates-badge" style="font-size:10px;color:#0066CC;font-weight:500"></span>
        </div>
        <div class="dp-grid">
          <div class="dpf"><label>Margen canal BP</label><div class="sw"><input type="number" id="e-bp" step="0.1"><span class="sfx">%</span></div></div>
          <div class="dpf"><label>Margen Ingram</label><div class="sw"><input type="number" id="e-ingram" step="0.1"><span class="sfx">%</span></div></div>
          <div class="dpf"><label>Dcto adicional</label><div class="sw"><input type="number" id="e-extra" value="0" step="0.1"><span class="sfx">%</span></div></div>
        </div>
        <div class="toggles">
          <label class="toggle"><input type="checkbox" id="e-cvr-renew" checked onchange="recalc()"> CVR Renew BP (<span id="e-crp">2</span>%)</label>
          <label class="toggle" id="lbl-ontime"><input type="checkbox" id="e-cvr-ontime" onchange="recalc()"> <span id="ontime-txt">CVR OnTime BP (<span id="e-cop">2</span>%)</span></label>
        </div>
        <div style="margin-top:10px">
          <button class="btn btn-g btn-sm" onclick="recalc()">↻ Recalcular</button>
        </div>
      </div>
 
      <div class="scards" id="d-scards"></div>
 
      <div class="dcard">
        <div class="cov-lines" id="d-cov"></div>
        <div class="tbl-wrap">
          <table>
            <thead><tr>
              <th>Part number</th><th>Descripción</th><th>Cobertura</th>
              <th style="text-align:right">Qty</th><th style="text-align:right">MEP</th>
              <th style="text-align:right">Dcto canal</th><th style="text-align:right">CVR Renew</th>
              <th style="text-align:right">CVR OnTime</th><th style="text-align:right">Precio BP</th>
            </tr></thead>
            <tbody id="d-tbody"></tbody>
          </table>
        </div>
        <div class="btn-grp">
          <button class="btn btn-p" id="btn-dl" onclick="downloadQuote()">⬇ Descargar PDF + Excel</button>
          <button class="btn btn-a hidden" id="btn-planilla" onclick="downloadPlanilla()">⬇ Planilla IBM</button>
          <span id="dl-status" style="font-size:12px;color:#64748B"></span>
        </div>
        <div class="review-box">
          <input type="checkbox" id="review-chk" onchange="toggleReview(this.checked)">
          <label for="review-chk">✓ &nbsp;Quote revisado — marcar como listo para el ZIP</label>
          <span class="review-hint" id="review-hint"></span>
        </div>
      </div>
 
      <div class="dcard" style="padding:0;overflow:hidden">
        <button class="audit-btn" onclick="toggleAudit()">
          <span>Vista de auditoría — datos XML vs cálculos</span>
          <span id="audit-arr">▼ Expandir</span>
        </button>
        <div class="audit-wrap" id="audit-wrap">
          <div class="audit-inner">
            <table style="width:100%;border-collapse:collapse;font-size:11px">
              <thead id="audit-thead"></thead>
              <tbody id="audit-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>
 
    </div>
  </div>
</div>
 
<script>
const API=window.location.origin;
let allQuotes={},covAlerts={},currentQ=null,currentLines=[],selectedFile=null,quoteParams={};
const fmt=n=>Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});
 
async function loadRates(){
  try{
    const r=await fetch(API+'/renewals/rates').then(x=>x.json());
    if(!r.success)return;
    document.getElementById('g-bp').value=r.bp_margin;
    document.getElementById('g-ingram').value=r.ingram_margin;
    document.getElementById('g-cvr-renew').value=r.cvr_renew_bp;
    document.getElementById('g-cvr-renew-ing').value=r.cvr_renew_ingram;
    document.getElementById('g-cvr-ontime').value=r.cvr_ontime_bp;
    document.getElementById('g-cvr-ontime-ing').value=r.cvr_ontime_ingram;
    const b=document.getElementById('rates-badge');
    if(b)b.textContent=r.label||'';
    updatePills();
  }catch(e){}
}
loadRates();
 
function getG(){
  return{
    bp:parseFloat(document.getElementById('g-bp').value)||2,
    ingram:parseFloat(document.getElementById('g-ingram').value)||1,
    cvrRenew:parseFloat(document.getElementById('g-cvr-renew').value)||2,
    cvrRenewIng:parseFloat(document.getElementById('g-cvr-renew-ing').value)||1,
    cvrOntime:parseFloat(document.getElementById('g-cvr-ontime').value)||2,
    cvrOntimeIng:parseFloat(document.getElementById('g-cvr-ontime-ing').value)||1,
  };
}
 
function updatePills(){
  const g=getG();
  document.getElementById('params-pills').innerHTML=
    `<span class="pill">BP <strong>${g.bp}%</strong></span>`+
    `<span class="pill">Ingram <strong>${g.ingram}%</strong></span>`+
    `<span class="pill">CVR Renew <strong>${g.cvrRenew}%</strong></span>`+
    `<span class="pill">OnTime <strong>${g.cvrOntime}%</strong></span>`;
}
 
function toggleParams(){
  const p=document.getElementById('params-panel');
  const open=p.classList.toggle('open');
  document.getElementById('params-lnk').textContent=open?'✕ Cerrar':'✎ Editar';
}
 
function applyParams(){
  updatePills();
  toggleParams();
  if(currentQ)openEditor(currentQ);
}
 
function setStep(n){
  for(let i=1;i<=3;i++){
    const el=document.getElementById('st'+i);
    el.className='step'+(i<n?' done':i===n?' active':'');
  }
}
 
async function handleFile(file){
  if(!file||!file.name.endsWith('.xml'))return;
  selectedFile=file;
  const st=document.getElementById('upload-status');
  st.className='alert a-ok';st.textContent='Procesando...';st.classList.remove('hidden');
  const fd=new FormData();fd.append('file',file);
  try{
    const res=await fetch(API+'/renewals/parse-renewal-report',{method:'POST',body:fd});
    const data=await res.json();
    if(!data.success)throw new Error(data.error);
    allQuotes=data.quotes||{};
    covAlerts=data.coverage_alerts||{};
    Object.keys(allQuotes).forEach(qn=>{allQuotes[qn]._checked=false;allQuotes[qn]._downloaded=false;});
    document.getElementById('upload-screen').classList.add('hidden');
    document.getElementById('main-layout').classList.remove('hidden');
    updatePills();
    renderList();
    setStep(2);
  }catch(e){
    st.className='alert a-err';st.textContent='Error: '+e.message;
  }
}
 
function renderList(){
  const total=Object.keys(allQuotes).length;
  const checked=Object.values(allQuotes).filter(q=>q._checked).length;
  document.getElementById('sb-title').textContent=`${total} quote(s) encontrados`;
  document.getElementById('sb-meta').textContent=
    checked>0?`${checked} listo(s) · ${total-checked} pendiente(s)`:'Selecciona uno para ver el detalle';
  if(checked===total&&total>0)setStep(3);
 
  const bf=document.getElementById('bulk-footer');
  if(checked>0){
    bf.classList.remove('hidden');
    const totalUSD=Object.values(allQuotes).filter(q=>q._checked).reduce((s,q)=>s+parseFloat(q.total_ibm||0),0);
    document.getElementById('bf-info').textContent=`${checked} quote(s) listos · $${fmt(totalUSD)} USD`;
    const canales={};
    Object.values(allQuotes).filter(q=>q._checked).forEach(q=>{
      const r=(q.lines&&q.lines[0]&&q.lines[0].reseller||'').replace(/^\d+\s*/,'').trim();
      canales[r]=(canales[r]||0)+1;
    });
    document.getElementById('bf-sub').textContent=Object.keys(canales).join(', ');
  }else{
    bf.classList.add('hidden');
  }
 
  const container=document.getElementById('q-list');
  container.innerHTML='';
  Object.entries(allQuotes).forEach(([qn,qdata])=>{
    const hasWarn=!!covAlerts[qn];
    const isReady=!!qdata._checked;
    const isSel=currentQ===qn;
    const customer=(qdata.customer||'').replace(/^\d+-/,'').replace(/,.*$/,'').trim();
    const reseller=(qdata.lines&&qdata.lines[0]&&qdata.lines[0].reseller||'').replace(/^\d+\s*/,'').trim();
    const div=document.createElement('div');
    div.className='q-item'+(isReady?' ready':'')+(hasWarn?' has-warn':'')+(isSel?' sel':'');
    div.id='qrow-'+qn;
    div.innerHTML=`
      <div class="chk${isReady?' on':''}" onclick="event.stopPropagation();toggleCheck('${qn}')">${isReady?'✓':''}</div>
      <div class="qi-body" onclick="openEditor('${qn}')">
        <div class="qi-num">${qn}</div>
        <div class="qi-cust">${customer}${reseller?' · '+reseller:''}</div>
      </div>
      <div class="qi-right" onclick="openEditor('${qn}')">
        <div class="qi-amt">$${fmt(qdata.total_ibm)}</div>
        ${hasWarn?'<span class="qtag qtag-warn">⚠ Coverage</span>':isReady?'<span class="qtag qtag-ok">✓ Listo</span>':'<span class="qtag qtag-gray">Pendiente</span>'}
      </div>`;
    container.appendChild(div);
  });
}
 
function toggleCheck(qn){
  allQuotes[qn]._checked=!allQuotes[qn]._checked;
  if(currentQ===qn){
    document.getElementById('review-chk').checked=allQuotes[qn]._checked;
    updateReviewHint();
  }
  renderList();
}
 
function toggleReview(checked){
  if(!currentQ)return;
  allQuotes[currentQ]._checked=checked;
  updateReviewHint();
  renderList();
}
 
function updateReviewHint(){
  if(!currentQ)return;
  const hint=document.getElementById('review-hint');
  const ok=allQuotes[currentQ]._checked;
  hint.textContent=ok?'✓ Se incluirá en el ZIP':'No se incluirá en la descarga masiva';
  hint.style.color=ok?'#16A34A':'#94A3B8';
}
 
function openEditor(qn){
  currentQ=qn;
  const qdata=allQuotes[qn];
  currentLines=(qdata.lines||[]).map(l=>({...l,ontime:qdata.ontime||false}));
  const customer=(qdata.customer||'').replace(/^\d+-/,'').replace(/,.*$/,'').trim();
  const reseller=(currentLines[0]?.reseller||'').replace(/^\d+\s*/,'').trim();
  document.getElementById('d-title').textContent='Quote '+qn+' — '+customer;
  document.getElementById('d-sub').textContent=reseller+(currentLines[0]?.coverage_dates?' · '+currentLines[0].coverage_dates:'');
  const g=getG();
  const p=quoteParams[qn]||{};
  document.getElementById('e-bp').value=p.bp||g.bp;
  document.getElementById('e-ingram').value=p.ingram||g.ingram;
  document.getElementById('e-extra').value=p.extra||0;
  document.getElementById('e-cvr-renew').checked=p.useRenew!==undefined?p.useRenew:true;
  document.getElementById('e-crp').textContent=g.cvrRenew;
  document.getElementById('e-cop').textContent=g.cvrOntime;
 
  // OnTime
  const isOntime=!!(currentLines[0]&&currentLines[0].ontime);
  const oc=document.getElementById('e-cvr-ontime');
  const ol=document.getElementById('lbl-ontime');
  if(!isOntime){
    oc.checked=false;oc.disabled=true;
    ol.style.opacity='0.4';ol.style.cursor='not-allowed';ol.title='Fecha vencida';
    document.getElementById('ontime-txt').textContent=`CVR OnTime BP (${g.cvrOntime}%) — vencido`;
  }else{
    oc.checked=p.useOntime!==undefined?p.useOntime:true;
    oc.disabled=false;ol.style.opacity='1';ol.style.cursor='pointer';ol.title='';
    document.getElementById('ontime-txt').textContent=`CVR OnTime BP (${g.cvrOntime}%)`;
  }
 
  // Coverage
  const errLines=covAlerts[qn]||[];
  document.getElementById('btn-planilla').classList.toggle('hidden',errLines.length===0);
  const covDiv=document.getElementById('d-cov');
  covDiv.innerHTML=currentLines.map(line=>{
    const err=errLines.find(e=>e.part_number===line.part_number);
    return err
      ?`<div class="cl cl-warn">⚠ <strong>${err.part_number}</strong> — Coverage incorrecto: ${err.months} mes(es) · Corregir hasta ${err.correct_end}</div>`
      :`<div class="cl cl-ok">✓ <strong>${line.part_number}</strong> — Coverage correcto: 12 meses</div>`;
  }).join('');
 
  document.getElementById('review-chk').checked=!!allQuotes[qn]._checked;
  updateReviewHint();
  recalc();
  document.getElementById('detail-empty').classList.add('hidden');
  document.getElementById('detail-content').classList.remove('hidden');
  renderList();
  document.getElementById('detail').scrollTo({top:0,behavior:'smooth'});
}
 
function calcLine(line,bp,cvrRenew,cvrOntime,extra,useRenew,useOntime){
  const mep=parseFloat(line.extended_price||0);
  const cc=+(mep*(1-bp/100)).toFixed(2);
  const rv=useRenew?+(mep*(cvrRenew/100)).toFixed(2):0;
  const ot=useOntime&&line.ontime?+(mep*(cvrOntime/100)).toFixed(2):0;
  const ex=+(cc*(extra/100)).toFixed(2);
  return{mep,cc,rv,ot,finalBP:+(cc-rv-ot-ex).toFixed(2)};
}
 
function recalc(){
  const g=getG();
  const bp=parseFloat(document.getElementById('e-bp').value)||g.bp;
  const extra=parseFloat(document.getElementById('e-extra').value)||0;
  const useRenew=document.getElementById('e-cvr-renew').checked;
  const useOntime=document.getElementById('e-cvr-ontime').checked;
  let tM=0,tC=0,tRv=0,tOt=0,tF=0;
  const tbody=document.getElementById('d-tbody');
  tbody.innerHTML='';
  currentLines.forEach(line=>{
    const c=calcLine(line,bp,g.cvrRenew,g.cvrOntime,extra,useRenew,useOntime);
    tM+=c.mep;tC+=c.cc;tRv+=c.rv;tOt+=c.ot;tF+=c.finalBP;
    const cov=(line.coverage_dates||'').replace(' To ','→');
    const tr=document.createElement('tr');
    tr.innerHTML=`
      <td><strong>${line.part_number||'—'}</strong></td>
      <td style="font-size:11px">${(line.part_description||'').substring(0,35)}...</td>
      <td style="font-size:11px;white-space:nowrap">${cov}</td>
      <td style="text-align:right">${line.quantity||1}</td>
      <td style="text-align:right;color:#64748B">$${fmt(c.mep)}</td>
      <td style="text-align:right;color:#DC2626">-$${fmt(c.mep-c.cc)}</td>
      <td style="text-align:right;color:#16A34A">${useRenew?'-$'+fmt(c.rv):'—'}</td>
      <td style="text-align:right;color:#16A34A">${useOntime&&line.ontime?'-$'+fmt(c.ot):'—'}</td>
      <td style="text-align:right;font-weight:600">$${fmt(c.finalBP)}</td>`;
    tbody.appendChild(tr);
  });
  const tr=document.createElement('tr');
  tr.className='tot-row';
  tr.innerHTML=`<td colspan="4" style="text-align:right">Total</td>
    <td style="text-align:right">$${fmt(tM)}</td>
    <td style="text-align:right">-$${fmt(tM-tC)}</td>
    <td style="text-align:right">-$${fmt(tRv)}</td>
    <td style="text-align:right">-$${fmt(tOt)}</td>
    <td style="text-align:right">$${fmt(tF)}</td>`;
  tbody.appendChild(tr);
  document.getElementById('d-scards').innerHTML=`
    <div class="sc"><div class="sc-lbl">Precio MEP total</div><div class="sc-val">$${fmt(tM)}</div></div>
    <div class="sc"><div class="sc-lbl">Nota de descuento</div><div class="sc-val g">$${fmt(tM-tF)}</div></div>
    <div class="sc"><div class="sc-lbl">Precio final canal</div><div class="sc-val b">$${fmt(tF)}</div></div>`;
}
 
async function downloadBulk(){
  const checked=Object.entries(allQuotes).filter(([,q])=>q._checked);
  if(!checked.length)return;
  const btn=document.getElementById('btn-bulk');
  btn.disabled=true;btn.innerHTML='<span class="spin"></span>';
  const g=getG();
  try{
    const fd=new FormData();
    fd.append('file',selectedFile);
    fd.append('bp_margin',g.bp);fd.append('ingram_margin',g.ingram);
    fd.append('cvr_renew_bp',g.cvrRenew);fd.append('cvr_renew_ingram',g.cvrRenewIng);
    fd.append('cvr_ontime_bp',g.cvrOntime);fd.append('cvr_ontime_ingram',g.cvrOntimeIng);
    fd.append('quote_numbers',JSON.stringify(checked.map(([qn])=>qn)));
    fd.append('quote_params',JSON.stringify(quoteParams));
    const res=await fetch(API+'/renewals/generate-bulk',{method:'POST',body:fd});
    if(!res.ok)throw new Error((await res.json().catch(()=>({error:'Error'}))).error);
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    const m=(res.headers.get('Content-Disposition')||'').match(/filename="?([^"]+)"?/);
    a.download=m?m[1]:`IBM_SS_${new Date().toISOString().slice(0,10).replace(/-/g,'')}_${checked.length}quotes.zip`;
    a.href=url;a.click();URL.revokeObjectURL(url);
    checked.forEach(([qn])=>allQuotes[qn]._downloaded=true);
    renderList();
  }catch(e){alert('Error: '+e.message);}
  finally{btn.disabled=false;btn.innerHTML='⬇ ZIP';}
}
 
async function downloadQuote(){
  const btn=document.getElementById('btn-dl');
  const status=document.getElementById('dl-status');
  btn.disabled=true;btn.innerHTML='<span class="spin"></span> Generando...';status.textContent='';
  try{
    const g=getG();
    const bp=parseFloat(document.getElementById('e-bp').value)||g.bp;
    const ingram=parseFloat(document.getElementById('e-ingram').value)||g.ingram;
    const extra=parseFloat(document.getElementById('e-extra').value)||0;
    const useRenew=document.getElementById('e-cvr-renew').checked;
    const useOntime=document.getElementById('e-cvr-ontime').checked;
    const fd=new FormData();
    fd.append('file',selectedFile);
    fd.append('bp_margin',bp);fd.append('ingram_margin',ingram);fd.append('extra_discount',extra);
    fd.append('cvr_renew_bp',useRenew?g.cvrRenew:0);fd.append('cvr_renew_ingram',useRenew?g.cvrRenewIng:0);
    fd.append('cvr_ontime_bp',useOntime?g.cvrOntime:0);fd.append('cvr_ontime_ingram',useOntime?g.cvrOntimeIng:0);
    fd.append('quote_filter',currentQ);
    const res=await fetch(API+'/renewals/generate-quote',{method:'POST',body:fd});
    if(!res.ok)throw new Error((await res.json()).error||'Error');
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    const m=(res.headers.get('Content-Disposition')||'').match(/filename="?([^"]+)"?/);
    a.download=m?m[1]:`Quote_${currentQ}.zip`;
    a.href=url;a.click();URL.revokeObjectURL(url);
    allQuotes[currentQ]._downloaded=true;
    quoteParams[currentQ]={bp,ingram,extra,useRenew,useOntime,
      cvrRenew:g.cvrRenew,cvrOntime:g.cvrOntime,cvrRenewIng:g.cvrRenewIng,cvrOntimeIng:g.cvrOntimeIng};
    status.textContent='✓ Descargado';
    updateReviewHint();renderList();
  }catch(e){status.textContent='Error: '+e.message;}
  finally{btn.disabled=false;btn.innerHTML='⬇ Descargar PDF + Excel';}
}
 
async function downloadPlanilla(){
  const btn=document.getElementById('btn-planilla');
  const status=document.getElementById('dl-status');
  btn.disabled=true;btn.textContent='Generando...';
  try{
    const fd=new FormData();fd.append('file',selectedFile);fd.append('quote_filter',currentQ);
    const res=await fetch(API+'/renewals/generate-planilla',{method:'POST',body:fd});
    if(!res.ok)throw new Error('Error');
    const blob=await res.blob();
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.download=`Planilla_IBM_${currentQ}.xlsx`;a.href=url;a.click();URL.revokeObjectURL(url);
    status.textContent='✓ Planilla descargada';
  }catch(e){status.textContent='Error: '+e.message;}
  finally{btn.disabled=false;btn.innerHTML='⬇ Planilla IBM';}
}
 
function toggleAudit(){
  const w=document.getElementById('audit-wrap');
  const open=w.classList.toggle('open');
  document.getElementById('audit-arr').textContent=open?'▲ Colapsar':'▼ Expandir';
  if(open)buildAudit();
}
 
function buildAudit(){
  const g=getG();
  const bp=parseFloat(document.getElementById('e-bp').value)||g.bp;
  const ingram=parseFloat(document.getElementById('e-ingram').value)||g.ingram;
  const extra=parseFloat(document.getElementById('e-extra').value)||0;
  const useRenew=document.getElementById('e-cvr-renew').checked;
  const useOntime=document.getElementById('e-cvr-ontime').checked;
  const XF=[
    {k:'quote_number',l:'Quote Number'},{k:'part_number',l:'Part Number'},
    {k:'part_description',l:'Descripción'},{k:'quantity',l:'Qty'},
    {k:'unit_price',l:'Unit Price'},{k:'extended_price',l:'Extended Price'},
    {k:'currency',l:'Moneda'},{k:'coverage_dates',l:'Coverage Dates'},
    {k:'renewal_due_date',l:'Renewal Due Date'},{k:'site',l:'Site'},
    {k:'ibm_customer_number',l:'ICN'},{k:'reseller',l:'Reseller'},
    {k:'agreement_information',l:'Agreement'},
  ];
  const CF=[
    {k:'_bp',l:`Margen BP (${bp}%)`},{k:'_dcto',l:'Dcto Canal $'},
    {k:'_canal',l:'Costo Canal'},{k:'_rvw',l:`CVR Renew (${g.cvrRenew}%)`},
    {k:'_ot',l:`CVR OnTime (${g.cvrOntime}%)`},{k:'_final',l:'Precio Final BP'},
    {k:'_ing',l:`Costo Ingram (${ingram}%)`},{k:'_ontime',l:'Renueva a tiempo'},
    {k:'_cov',l:'Coverage OK'},
  ];
  const thead=document.getElementById('audit-thead');
  thead.innerHTML='';
  const hr=document.createElement('tr');
  hr.innerHTML='<th class="ath" style="min-width:130px">Campo</th>';
  currentLines.forEach((l,i)=>hr.innerHTML+=`<th class="ath" style="min-width:100px">L${i+1} · ${l.part_number||''}</th>`);
  thead.appendChild(hr);
  const tbody=document.getElementById('audit-tbody');
  tbody.innerHTML='';
  const s1=document.createElement('tr');
  s1.innerHTML=`<td class="asec" colspan="${currentLines.length+1}">Datos XML</td>`;
  tbody.appendChild(s1);
  XF.forEach(f=>{
    const tr=document.createElement('tr');
    let row=`<td class="atd" style="font-weight:500;color:#475569">${f.l}</td>`;
    currentLines.forEach(line=>row+=`<td class="atd">${line[f.k]!==undefined?String(line[f.k]):'—'}</td>`);
    tr.innerHTML=row;tbody.appendChild(tr);
  });
  const s2=document.createElement('tr');
  s2.innerHTML=`<td class="asec" colspan="${currentLines.length+1}">Cálculos</td>`;
  tbody.appendChild(s2);
  const calcs=currentLines.map(line=>{
    const c=calcLine(line,bp,g.cvrRenew,g.cvrOntime,extra,useRenew,useOntime);
    const hce=(covAlerts[currentQ]||[]).find(e=>e.part_number===line.part_number);
    return{
      _bp:bp+'%',_dcto:'$'+fmt(c.mep-c.cc),_canal:'$'+fmt(c.cc),
      _rvw:useRenew?'-$'+fmt(c.rv):'No aplica',
      _ot:useOntime&&line.ontime?'-$'+fmt(c.ot):'No aplica',
      _final:'$'+fmt(c.finalBP),_ing:'$'+fmt(c.cc*(1-ingram/100)),
      _ontime:line.ontime?'✓ Sí':'✗ No',
      _cov:hce?`⚠ ${hce.months} mes(es)`:'✓ 12 meses',
    };
  });
  CF.forEach(f=>{
    const tr=document.createElement('tr');
    let row=`<td class="atd2" style="font-weight:500;color:#0066CC">${f.l}</td>`;
    calcs.forEach(c=>{
      const v=c[f.k]||'—';
      const col=v.startsWith('✓')?'#166534':v.startsWith('⚠')||v.startsWith('✗')?'#991B1B':'#0C447C';
      row+=`<td class="atd2" style="color:${col}">${v}</td>`;
    });
    tr.innerHTML=row;tbody.appendChild(tr);
  });
}
 
document.getElementById('xml-file').addEventListener('change',e=>{if(e.target.files[0])handleFile(e.target.files[0]);});
const dz=document.getElementById('dropzone');
dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag');});
dz.addEventListener('dragleave',()=>dz.classList.remove('drag'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('drag');if(e.dataTransfer.files[0])handleFile(e.dataTransfer.files[0]);});
</script>
</body>
</html>'''
 
def register_ui_route(blueprint):
    from flask import make_response
 
    @blueprint.get('/ui')
    def renewal_ui():
        response = make_response(PORTAL_HTML)
        response.headers['Content-Type'] = 'text/html; charset=utf-8'
        return response
 