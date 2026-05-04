"""
Portal de Renovaciones IBM S&S — Ingram Micro
"""

PORTAL_HTML = '''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Renovaciones IBM S&S — Ingram Micro</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#F0F4F8;color:#1A2B3C;min-height:100vh}
.topbar{background:#fff;border-bottom:1px solid #E2E8F0;padding:0 1.5rem;height:52px;display:flex;align-items:center;justify-content:space-between}
.logo{font-size:18px;font-weight:700;color:#0066CC;letter-spacing:-0.5px}.logo span{color:#003366}
.ibm-badge{background:#054ADA;color:#fff;font-size:10px;font-weight:700;padding:3px 9px;border-radius:4px}
.main{max-width:900px;margin:0 auto;padding:1.5rem 1rem}
.step-bar{display:flex;gap:8px;margin-bottom:1.5rem}
.step{flex:1;background:#fff;border:0.5px solid #E2E8F0;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px}
.step.active{border-color:#0066CC;background:#EBF3FF}
.step.done{border-color:#16A34A;background:#F0FDF4}
.snum{width:24px;height:24px;border-radius:50%;background:#F1F5F9;color:#64748B;font-size:11px;font-weight:700;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.step.active .snum{background:#0066CC;color:#fff}
.step.done .snum{background:#16A34A;color:#fff}
.slbl{font-size:11px;font-weight:600;color:#1A2B3C}
.sdesc{font-size:10px;color:#64748B}
.card{background:#fff;border:0.5px solid #E2E8F0;border-radius:12px;padding:1.25rem;margin-bottom:1rem}
.card-title{font-size:14px;font-weight:600;color:#0F2B5B;margin-bottom:1rem}
.dropzone{border:1.5px dashed #CBD5E1;border-radius:10px;padding:2rem 1rem;text-align:center;cursor:pointer;transition:border-color 0.15s,background 0.15s;position:relative}
.dropzone:hover,.dropzone.drag{border-color:#0066CC;background:#EBF3FF}
.dropzone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}
.alert{border-radius:8px;padding:10px 14px;font-size:12px;margin-top:10px}
.alert-success{background:#F0FDF4;color:#166534;border:1px solid #BBF7D0}
.alert-warning{background:#FFFBEB;color:#92400E;border:1px solid #FCD34D}
.alert-error{background:#FEF2F2;color:#991B1B;border:1px solid #FECACA}
.global-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:1rem}
.field{display:flex;flex-direction:column;gap:4px}
.field label{font-size:10px;font-weight:600;color:#64748B;text-transform:uppercase;letter-spacing:0.4px}
.sfx-wrap{position:relative}
.sfx-wrap input{padding-right:26px;width:100%}
.sfx{position:absolute;right:8px;top:50%;transform:translateY(-50%);font-size:12px;color:#64748B;pointer-events:none}
.hint{font-size:10px;color:#94A3B8;margin-top:2px}

/* ── Quote list ── */
.quote-list{display:flex;flex-direction:column;gap:8px}
.q-row{background:#fff;border:0.5px solid #E2E8F0;border-radius:8px;padding:0;display:grid;grid-template-columns:auto 1fr auto auto;align-items:center;gap:0;cursor:pointer;transition:all 0.15s;overflow:hidden}
.q-row:hover{border-color:#0066CC;background:#F8FAFF}
.q-row.selected{border-color:#0066CC;background:#EBF3FF}
.q-row.checked{border-color:#16A34A;background:#F0FDF4}
.q-row.checked-done{border-color:#16A34A;background:#DCFCE7}

/* Checkbox columna izquierda */
.q-check-col{padding:12px 10px 12px 14px;display:flex;align-items:center}
.q-checkbox{width:18px;height:18px;border:1.5px solid #CBD5E1;border-radius:4px;cursor:pointer;display:flex;align-items:center;justify-content:center;flex-shrink:0;transition:all 0.15s;background:#fff}
.q-checkbox.checked{background:#16A34A;border-color:#16A34A;color:#fff;font-size:12px}
.q-checkbox.checked::after{content:'✓';color:#fff;font-size:11px;font-weight:700}

/* Info del quote */
.q-info{padding:12px 8px;min-width:0}
.q-num{font-size:13px;font-weight:600;color:#0F2B5B}
.q-customer{font-size:11px;color:#64748B;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.q-badges{padding:12px 8px;display:flex;gap:6px;align-items:center;flex-shrink:0}
.b-warn{background:#FEF3C7;color:#92400E;font-size:10px;padding:2px 8px;border-radius:4px;font-weight:600}
.b-done{background:#DCFCE7;color:#166534;font-size:10px;padding:2px 8px;border-radius:4px;font-weight:600}
.b-dl{background:#EBF3FF;color:#0066CC;font-size:10px;padding:2px 8px;border-radius:4px;font-weight:600}
.b-lines{background:#F1F5F9;color:#64748B;font-size:10px;padding:2px 8px;border-radius:4px}
.q-total-col{padding:12px 14px 12px 8px;text-align:right;flex-shrink:0}
.q-total{font-size:13px;font-weight:600;color:#0F2B5B}
.q-total-sub{font-size:10px;color:#64748B}

/* ── Barra de descarga masiva ── */
.bulk-bar{background:#fff;border:1.5px solid #16A34A;border-radius:10px;padding:14px 16px;display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:1rem}
.bulk-bar.hidden{display:none!important}
.bulk-info{display:flex;align-items:center;gap:10px}
.bulk-count{background:#DCFCE7;color:#166534;font-size:13px;font-weight:700;padding:4px 12px;border-radius:6px}
.bulk-label{font-size:13px;font-weight:600;color:#1A2B3C}
.bulk-sub{font-size:11px;color:#64748B;margin-top:1px}

/* Editor */
.editor-box{background:#F8FAFC;border:0.5px solid #E2E8F0;border-radius:8px;padding:1rem;margin-top:10px}
.editor-box-title{font-size:12px;font-weight:600;color:#0F2B5B;margin-bottom:12px}
.fields-row{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:12px}
.toggle-row{display:flex;gap:16px;margin-bottom:12px;flex-wrap:wrap}
.toggle{display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none}
.toggle input{width:14px;height:14px}
.toggle span{font-size:12px;color:#1A2B3C}
.summary-bar{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:12px 0}
.scard{background:#F1F5F9;border-radius:8px;padding:10px 14px}
.scard-lbl{font-size:10px;color:#64748B;text-transform:uppercase;letter-spacing:0.4px;font-weight:600;margin-bottom:3px}
.scard-val{font-size:18px;font-weight:700;color:#0F2B5B}
.scard-val.blue{color:#0066CC}
.scard-val.red{color:#DC2626}

/* Check de revisión */
.review-check{background:#F0FDF4;border:1px solid #BBF7D0;border-radius:8px;padding:12px 16px;display:flex;align-items:center;gap:12px;margin-top:10px}
.review-check label{display:flex;align-items:center;gap:8px;cursor:pointer;font-size:13px;font-weight:600;color:#166534}
.review-check input[type=checkbox]{width:18px;height:18px;accent-color:#16A34A;cursor:pointer}
.review-check .review-hint{font-size:11px;color:#64748B;margin-left:auto}

table{width:100%;border-collapse:collapse;font-size:12px}
th{background:#0066CC;color:#fff;padding:8px 10px;text-align:left;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.3px}
th:first-child{border-radius:6px 0 0 0}th:last-child{border-radius:0 6px 0 0}
td{padding:8px 10px;border-bottom:0.5px solid #F1F5F9;vertical-align:middle}
tr:nth-child(even) td{background:#F8FAFC}
.tot-row td{background:#EBF3FF!important;font-weight:700;color:#0C447C}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:7px;height:38px;padding:0 18px;border-radius:8px;border:none;cursor:pointer;font-size:13px;font-weight:600;transition:background 0.15s,transform 0.1s}
.btn:active{transform:scale(0.98)}
.btn-primary{background:#0066CC;color:#fff}.btn-primary:hover{background:#0052A3}
.btn-primary:disabled{background:#93C5FD;cursor:not-allowed}
.btn-success{background:#16A34A;color:#fff}.btn-success:hover{background:#15803D}
.btn-success:disabled{background:#86EFAC;cursor:not-allowed}
.btn-outline{background:#fff;color:#0066CC;border:1.5px solid #0066CC}.btn-outline:hover{background:#EBF3FF}
.btn-amber{background:#D97706;color:#fff}.btn-amber:hover{background:#B45309}
.btn-sm{height:32px;padding:0 12px;font-size:12px}
.btn-group{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px;align-items:center}
.spinner{width:14px;height:14px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;border-radius:50%;animation:spin 0.7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.hidden{display:none!important}
.sep{height:0.5px;background:#E2E8F0;margin:12px 0}
.pending-badge{background:#FEF3C7;color:#92400E;font-size:11px;font-weight:600;padding:3px 10px;border-radius:6px}
.audit-xml{background:#F8FAFC;color:#1A2B3C}
.audit-calc{background:#EBF3FF;color:#0C447C}
.audit-th-xml{background:#475569;color:#fff;padding:6px 8px;font-size:10px;font-weight:600;text-align:left;white-space:nowrap}
.audit-th-calc{background:#0066CC;color:#fff;padding:6px 8px;font-size:10px;font-weight:600;text-align:left;white-space:nowrap}
.audit-td{padding:6px 8px;border-bottom:0.5px solid #E2E8F0;white-space:nowrap;font-size:11px}
.audit-section{background:#1E293B;color:#fff;padding:6px 10px;font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.5px}
input[type=number]{height:36px;border:0.5px solid #D1D5DB;border-radius:8px;padding:0 10px;font-size:13px;color:#1A2B3C;background:#fff;outline:none;transition:border-color 0.15s}
input[type=number]:focus{border-color:#0066CC;box-shadow:0 0 0 3px rgba(0,102,204,0.1)}
</style>
</head>
<body>

<div class="topbar">
  <div style="display:flex;align-items:center;gap:10px">
    <span class="logo">IN<span>GRAM</span> MICRO</span>
    <div style="width:1px;height:18px;background:#E2E8F0"></div>
    <span style="font-size:12px;color:#64748B;font-weight:500">Renovaciones IBM S&amp;S</span>
  </div>
  <span class="ibm-badge">IBM Authorized Reseller</span>
</div>

<div class="main">

<div class="step-bar">
  <div class="step active" id="st1"><div class="snum">1</div><div><div class="slbl">Cargar XML</div><div class="sdesc">IBM PAO</div></div></div>
  <div class="step" id="st2"><div class="snum">2</div><div><div class="slbl">Revisar quotes</div><div class="sdesc">Marcar revisados</div></div></div>
  <div class="step" id="st3"><div class="snum">3</div><div><div class="slbl">Descargar</div><div class="sdesc">ZIP por canal</div></div></div>
</div>

<!-- UPLOAD -->
<div class="card" id="card-upload">
  <div class="card-title">Cargar reporte de IBM Passport Advantage</div>
  <div class="dropzone" id="dropzone">
    <input type="file" id="xml-file" accept=".xml">
    <div style="font-size:28px;margin-bottom:8px">📄</div>
    <div style="font-size:14px;font-weight:600;margin-bottom:3px">Arrastra el archivo XML aquí</div>
    <div style="font-size:12px;color:#64748B">o haz clic · Formato: XML for Microsoft Excel de IBM PAO</div>
  </div>
  <div id="upload-status" class="hidden"></div>
</div>

<!-- BARRA DESCARGA MASIVA (aparece cuando hay quotes chequeados) -->
<div class="bulk-bar hidden" id="bulk-bar">
  <div class="bulk-info">
    <div class="bulk-count" id="bulk-count">0</div>
    <div>
      <div class="bulk-label" id="bulk-label">quotes listos para descargar</div>
      <div class="bulk-sub" id="bulk-sub">ZIP organizado por canal · PDF + Excel + Planilla por quote</div>
    </div>
  </div>
  <button class="btn btn-success" id="btn-bulk" onclick="downloadBulk()">
    ⬇ Descargar ZIP por canal
  </button>
</div>

<!-- GLOBAL PARAMS + QUOTE LIST -->
<div class="card hidden" id="card-quotes">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:1rem">
    <div class="card-title" style="margin-bottom:0">Parámetros globales</div>
    <div id="pending-count"></div>
  </div>
  <div class="global-grid">
    <div class="field"><label>Margen Canal BP</label><div class="sfx-wrap"><input type="number" id="g-bp" value="3" min="0" max="100" step="0.1"><span class="sfx">%</span></div><div class="hint">Descuento sobre MEP</div></div>
    <div class="field"><label>Margen Ingram</label><div class="sfx-wrap"><input type="number" id="g-ingram" value="2.1" min="0" max="100" step="0.1"><span class="sfx">%</span></div><div class="hint">Interno — no visible al canal</div></div>
    <div class="field"><label>CVR Renew BP</label><div class="sfx-wrap"><input type="number" id="g-cvr-renew" value="2" min="0" max="100" step="0.1"><span class="sfx">%</span></div><div class="hint">Siempre aplica</div></div>
    <div class="field"><label>CVR Renew Ingram</label><div class="sfx-wrap"><input type="number" id="g-cvr-renew-ing" value="2" min="0" max="100" step="0.1"><span class="sfx">%</span></div></div>
    <div class="field"><label>CVR OnTime BP</label><div class="sfx-wrap"><input type="number" id="g-cvr-ontime" value="2" min="0" max="100" step="0.1"><span class="sfx">%</span></div><div class="hint">Solo si renueva a tiempo</div></div>
    <div class="field"><label>CVR OnTime Ingram</label><div class="sfx-wrap"><input type="number" id="g-cvr-ontime-ing" value="1" min="0" max="100" step="0.1"><span class="sfx">%</span></div></div>
  </div>
  <div class="sep"></div>
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px">
    <div style="font-size:12px;font-weight:600;color:#64748B;text-transform:uppercase;letter-spacing:0.5px">
      Quotes encontrados — revisá cada uno y marcalo como listo
    </div>
    <div style="display:flex;gap:6px">
      <button class="btn btn-outline btn-sm" onclick="checkAll()" id="btn-check-all">✓ Marcar todos</button>
      <button class="btn btn-outline btn-sm" onclick="uncheckAll()" style="color:#64748B;border-color:#CBD5E1">Desmarcar todos</button>
    </div>
  </div>
  <div class="quote-list" id="quote-list"></div>
</div>

<!-- EDITOR -->
<div class="card hidden" id="card-editor">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:1rem">
    <div>
      <div class="card-title" style="margin-bottom:2px" id="editor-title">Quote</div>
      <div style="font-size:12px;color:#64748B" id="editor-sub"></div>
    </div>
    <button class="btn btn-outline btn-sm" onclick="closeEditor()">← Volver</button>
  </div>

  <div class="editor-box">
    <div class="editor-box-title">Ajustar parámetros para este quote</div>
    <div class="fields-row">
      <div class="field"><label>Margen Canal BP</label><div class="sfx-wrap"><input type="number" id="e-bp" step="0.1"><span class="sfx">%</span></div></div>
      <div class="field"><label>Margen Ingram</label><div class="sfx-wrap"><input type="number" id="e-ingram" step="0.1"><span class="sfx">%</span></div></div>
      <div class="field"><label>Dcto adicional cliente</label><div class="sfx-wrap"><input type="number" id="e-extra" value="0" step="0.1"><span class="sfx">%</span></div></div>
    </div>
    <div class="toggle-row">
      <label class="toggle"><input type="checkbox" id="e-cvr-renew" checked><span>CVR Renew BP (<span id="e-crp">2%</span>)</span></label>
      <label class="toggle"><input type="checkbox" id="e-cvr-ontime"><span>CVR OnTime BP (<span id="e-cop">2%</span>)</span></label>
    </div>
    <button class="btn btn-primary btn-sm" onclick="recalc()">Recalcular</button>
  </div>

  <div class="summary-bar" id="editor-summary"></div>

  <div style="overflow-x:auto">
    <table>
      <thead><tr>
        <th>Part number</th><th>Descripción</th><th>Cobertura</th><th style="text-align:right">Qty</th>
        <th style="text-align:right">Precio MEP</th><th style="text-align:right">Dcto canal</th>
        <th style="text-align:right">CVR Renew</th><th style="text-align:right">CVR OnTime</th>
        <th style="text-align:right">Precio BP</th>
      </tr></thead>
      <tbody id="editor-tbody"></tbody>
    </table>
  </div>

  <div id="cov-alert" style="margin-top:10px"></div>

  <div class="btn-group">
    <button class="btn btn-success" id="btn-dl" onclick="downloadQuote()">⬇ Descargar PDF + Excel</button>
    <button class="btn btn-amber hidden" id="btn-planilla" onclick="downloadPlanilla()">⬇ Planilla IBM</button>
    <span id="dl-status" style="font-size:12px;color:#64748B"></span>
  </div>

  <!-- CHECK DE REVISIÓN -->
  <div class="review-check" id="review-check-box">
    <input type="checkbox" id="review-check-input" onchange="toggleReviewCheck(this.checked)">
    <label for="review-check-input">
      ✓ &nbsp;Quote revisado — marcar como listo para incluir en el ZIP por canal
    </label>
    <span class="review-hint" id="review-hint-text"></span>
  </div>

  <div style="margin-top:16px">
    <button class="btn btn-outline btn-sm" onclick="toggleAudit()" id="btn-audit" style="width:100%;justify-content:space-between">
      <span>Vista de auditoría — datos XML vs cálculos</span>
      <span id="audit-arrow" style="font-size:12px">▼ Expandir</span>
    </button>
    <div id="audit-panel" class="hidden" style="margin-top:8px;overflow-x:auto">
      <table id="audit-table" style="width:100%;border-collapse:collapse;font-size:11px">
        <thead id="audit-thead"></thead>
        <tbody id="audit-tbody"></tbody>
      </table>
    </div>
  </div>
</div>

</div>

<script>
const API = window.location.origin;
let allQuotes = {}, covAlerts = {}, currentQ = null, currentLines = [];
let selectedFile = null;
// quoteParams guarda los parámetros con los que se descargó cada quote
let quoteParams = {};

const fmt = n => Number(n).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2});

function setStep(n){
  for(let i=1;i<=3;i++){
    const el=document.getElementById('st'+i);
    el.className='step'+(i<n?' done':i===n?' active':'');
  }
}

function getG(){
  return {
    bp:          parseFloat(document.getElementById('g-bp').value)||3,
    ingram:      parseFloat(document.getElementById('g-ingram').value)||2.1,
    cvrRenew:    parseFloat(document.getElementById('g-cvr-renew').value)||2,
    cvrRenewIng: parseFloat(document.getElementById('g-cvr-renew-ing').value)||2,
    cvrOntime:   parseFloat(document.getElementById('g-cvr-ontime').value)||2,
    cvrOntimeIng:parseFloat(document.getElementById('g-cvr-ontime-ing').value)||1,
  };
}

function calcLine(line, bp, cvrRenew, cvrOntime, extra, useRenew, useOntime){
  const mep = parseFloat(line.extended_price||0);
  const costoCanal = +(mep*(1-bp/100)).toFixed(2);
  const rv  = useRenew  ? +(mep*(cvrRenew/100)).toFixed(2)  : 0;
  const ot  = useOntime && line.ontime ? +(mep*(cvrOntime/100)).toFixed(2) : 0;
  const ex  = +(costoCanal*(extra/100)).toFixed(2);
  const finalBP = +(costoCanal - rv - ot - ex).toFixed(2);
  return {mep, costoCanal, rv, ot, finalBP};
}

async function handleFile(file){
  if(!file||!file.name.endsWith('.xml')) return;
  selectedFile = file;
  const st = document.getElementById('upload-status');
  st.className='alert alert-success'; st.textContent='Procesando...'; st.classList.remove('hidden');
  const fd = new FormData(); fd.append('file', file);
  try{
    const res = await fetch(API+'/renewals/parse-renewal-report',{method:'POST',body:fd});
    const data = await res.json();
    if(!data.success) throw new Error(data.error);
    allQuotes = data.quotes||{};
    covAlerts = data.coverage_alerts||{};
    // Inicializar estado de check
    Object.keys(allQuotes).forEach(qn => {
      allQuotes[qn]._checked  = false;
      allQuotes[qn]._downloaded = false;
    });
    const al = Object.keys(covAlerts).length;
    st.className = al>0?'alert alert-warning':'alert alert-success';
    st.innerHTML = `✓ ${data.total_quotes} quote(s) · ${data.total_lines} líneas`+(al>0?` · <strong>${al} con alerta de coverage</strong>`:'');
    renderQuoteList();
    document.getElementById('card-quotes').classList.remove('hidden');
    setStep(2);
  }catch(e){
    st.className='alert alert-error'; st.textContent='Error: '+e.message;
  }
}

// ── Render lista de quotes ────────────────────────────────────────────────────
function renderQuoteList(){
  const container = document.getElementById('quote-list');
  container.innerHTML='';

  const total    = Object.keys(allQuotes).length;
  const checked  = Object.values(allQuotes).filter(q=>q._checked).length;
  const pending  = total - checked;

  // Actualizar contador superior
  const pc = document.getElementById('pending-count');
  if(pending === 0 && total > 0){
    pc.innerHTML='<span style="font-size:12px;color:#16A34A;font-weight:600">✓ Todos marcados como listos</span>';
    setStep(3);
  } else {
    pc.innerHTML=`<span class="pending-badge">${pending} pendiente(s)</span>`;
    if(checked > 0) setStep(2);
  }

  // Barra de descarga masiva
  updateBulkBar();

  Object.entries(allQuotes).forEach(([qn, qdata])=>{
    const hasAlert   = !!covAlerts[qn];
    const customer   = (qdata.customer||'').replace(/^\d+-/,'').replace(/,.*$/,'').trim();
    const reseller   = (qdata.lines&&qdata.lines[0]&&qdata.lines[0].reseller||'').replace(/^\d+\s*/,'').trim();
    const isChecked  = !!qdata._checked;
    const isDl       = !!qdata._downloaded;

    const div = document.createElement('div');
    div.className = 'q-row' + (isChecked?' checked':'') + (currentQ===qn?' selected':'');
    div.id = `qrow-${qn}`;

    div.innerHTML = `
      <div class="q-check-col" onclick="event.stopPropagation(); toggleCheck('${qn}')">
        <div class="q-checkbox${isChecked?' checked':''}" id="qchk-${qn}"></div>
      </div>
      <div class="q-info" onclick="openEditor('${qn}')">
        <div class="q-num">Quote ${qn}</div>
        <div class="q-customer">${customer}${reseller?' · '+reseller:''}</div>
      </div>
      <div class="q-badges" onclick="openEditor('${qn}')">
        <span class="b-lines">${qdata.line_count} línea(s)</span>
        ${hasAlert?'<span class="b-warn">⚠ Coverage</span>':''}
        ${isDl?'<span class="b-dl">⬇ Descargado</span>':''}
        ${isChecked?'<span class="b-done">✓ Listo</span>':''}
      </div>
      <div class="q-total-col" onclick="openEditor('${qn}')">
        <div class="q-total">$${fmt(qdata.total_ibm)}</div>
        <div class="q-total-sub">USD MEP</div>
      </div>
    `;
    container.appendChild(div);
  });
}

// ── Check individual ──────────────────────────────────────────────────────────
function toggleCheck(qn){
  allQuotes[qn]._checked = !allQuotes[qn]._checked;
  // Si estamos en el editor de ese quote, sincronizar checkbox interno
  if(currentQ === qn){
    document.getElementById('review-check-input').checked = allQuotes[qn]._checked;
    updateReviewHint();
  }
  renderQuoteList();
}

function checkAll(){
  Object.keys(allQuotes).forEach(qn => allQuotes[qn]._checked = true);
  if(currentQ) {
    document.getElementById('review-check-input').checked = true;
    updateReviewHint();
  }
  renderQuoteList();
}

function uncheckAll(){
  Object.keys(allQuotes).forEach(qn => allQuotes[qn]._checked = false);
  if(currentQ) {
    document.getElementById('review-check-input').checked = false;
    updateReviewHint();
  }
  renderQuoteList();
}

// Desde el editor
function toggleReviewCheck(checked){
  if(!currentQ) return;
  allQuotes[currentQ]._checked = checked;
  updateReviewHint();
  renderQuoteList();
}

function updateReviewHint(){
  if(!currentQ) return;
  const hint = document.getElementById('review-hint-text');
  const q = allQuotes[currentQ];
  if(q._checked){
    hint.textContent = q._downloaded
      ? '✓ Incluido en el próximo ZIP — ya fue descargado individualmente'
      : '✓ Se incluirá en el ZIP por canal al descargar';
    hint.style.color = '#16A34A';
  } else {
    hint.textContent = 'Sin marcar — no se incluirá en la descarga masiva';
    hint.style.color = '#94A3B8';
  }
}

// ── Barra descarga masiva ─────────────────────────────────────────────────────
function updateBulkBar(){
  const checked = Object.values(allQuotes).filter(q=>q._checked);
  const bar = document.getElementById('bulk-bar');
  const btn = document.getElementById('btn-bulk');

  if(checked.length === 0){
    bar.classList.add('hidden');
    return;
  }

  bar.classList.remove('hidden');
  document.getElementById('bulk-count').textContent = checked.length;

  // Agrupar por canal para mostrar en el label
  const canales = {};
  checked.forEach(q => {
    const r = (q.lines&&q.lines[0]&&q.lines[0].reseller||'SIN CANAL').replace(/^\d+\s*/,'').trim();
    canales[r] = (canales[r]||0) + 1;
  });
  const canalList = Object.entries(canales).map(([c,n])=>`${c} (${n})`).join(', ');
  const totalUSD = checked.reduce((s,q)=>s+parseFloat(q.total_ibm||0),0);

  document.getElementById('bulk-label').textContent =
    checked.length === 1
      ? `1 quote listo — $${fmt(totalUSD)} USD MEP`
      : `${checked.length} quotes listos — $${fmt(totalUSD)} USD MEP`;
  document.getElementById('bulk-sub').textContent =
    `Canales: ${canalList} · ZIP con carpeta por canal`;

  btn.disabled = false;
  btn.innerHTML = `⬇ Descargar ZIP por canal`;
}

// ── Descarga masiva ───────────────────────────────────────────────────────────
async function downloadBulk(){
  const checked = Object.entries(allQuotes).filter(([,q])=>q._checked);
  if(checked.length === 0) return;

  const btn = document.getElementById('btn-bulk');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Generando ZIP...';

  const g = getG();

  // Generar quote por quote y acumular blobs en memoria
  // Usamos la misma lógica que downloadQuote() pero para todos los chequeados
  // El backend ya genera el ZIP con estructura {año}/{reseller}/ por quote
  // Para el ZIP maestro por canal hacemos una petición por quote y los agrupamos en JS

  try {
    // Construir ZIP maestro en el servidor enviando todos los quotes chequeados
    // Enviamos los quote_numbers como lista y dejamos que el backend los agrupe
    const quoteNumbers = checked.map(([qn])=>qn);

    const fd = new FormData();
    fd.append('file', selectedFile);
    fd.append('bp_margin',         g.bp);
    fd.append('ingram_margin',      g.ingram);
    fd.append('cvr_renew_bp',       g.cvrRenew);
    fd.append('cvr_renew_ingram',   g.cvrRenewIng);
    fd.append('cvr_ontime_bp',      g.cvrOntime);
    fd.append('cvr_ontime_ingram',  g.cvrOntimeIng);
    // Enviar los quote numbers chequeados para que el backend filtre
    fd.append('quote_numbers', JSON.stringify(quoteNumbers));
    // Parámetros individuales por quote (si fueron ajustados en el editor)
    fd.append('quote_params', JSON.stringify(quoteParams));

    const res = await fetch(API+'/renewals/generate-bulk',{method:'POST',body:fd});
    if(!res.ok){
      const err = await res.json().catch(()=>({error:'Error del servidor'}));
      throw new Error(err.error||'Error del servidor');
    }

    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const cd   = res.headers.get('Content-Disposition')||'';
    const m    = cd.match(/filename="?([^"]+)"?/);
    // Nombre del ZIP: fecha + cantidad de quotes
    const today = new Date().toISOString().slice(0,10).replace(/-/g,'');
    a.download = m ? m[1] : `IBM_SS_Renovaciones_${today}_${quoteNumbers.length}quotes.zip`;
    a.href = url; a.click(); URL.revokeObjectURL(url);

    // Marcar como descargados
    quoteNumbers.forEach(qn => { allQuotes[qn]._downloaded = true; });
    renderQuoteList();

  } catch(e){
    alert('Error al generar ZIP: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '⬇ Descargar ZIP por canal';
  }
}

// ── Editor ────────────────────────────────────────────────────────────────────
function openEditor(qn){
  currentQ = qn;
  const qdata = allQuotes[qn];
  currentLines = (qdata.lines||[]).map(l=>({...l, ontime: qdata.ontime||false}));
  const customer = (qdata.customer||'').replace(/^\d+-/,'').replace(/,.*$/,'').trim();
  document.getElementById('editor-title').textContent = 'Quote ' + qn;
  document.getElementById('editor-sub').textContent = customer+(currentLines[0]?.coverage_dates?' · '+currentLines[0].coverage_dates:'');

  const g = getG();
  const savedP = quoteParams[qn] || {};
  document.getElementById('e-bp').value     = savedP.bp     || g.bp;
  document.getElementById('e-ingram').value  = savedP.ingram || g.ingram;
  document.getElementById('e-extra').value   = savedP.extra  || 0;
  document.getElementById('e-cvr-renew').checked  = savedP.useRenew  !== undefined ? savedP.useRenew  : true;
  document.getElementById('e-cvr-ontime').checked = savedP.useOntime !== undefined ? savedP.useOntime : false;
  document.getElementById('e-crp').textContent = g.cvrRenew+'%';
  document.getElementById('e-cop').textContent = g.cvrOntime+'%';

  // Coverage alerts
  const ca  = document.getElementById('cov-alert');
  const bpl = document.getElementById('btn-planilla');
  const lines = qdata.lines||[];
  const errLines = (covAlerts[qn]||[]);

  if(errLines.length > 0){
    bpl.classList.remove('hidden');
    let html = '<div style="display:flex;flex-direction:column;gap:6px">';
    lines.forEach(line=>{
      const err = errLines.find(e=>e.part_number===line.part_number);
      if(err){
        html += `<div style="display:flex;align-items:center;gap:8px;background:#FEF3C7;border:0.5px solid #FCD34D;border-radius:6px;padding:7px 12px;font-size:12px">
          <span style="color:#D97706;font-size:14px">⚠</span>
          <span><strong>${err.part_number}</strong> · Coverage incorrecto: <strong>${err.months} mes(es)</strong> → correcto hasta <strong>${err.correct_end}</strong></span>
        </div>`;
      } else {
        html += `<div style="display:flex;align-items:center;gap:8px;background:#F0FDF4;border:0.5px solid #BBF7D0;border-radius:6px;padding:7px 12px;font-size:12px">
          <span style="color:#16A34A;font-size:14px">✓</span>
          <span><strong>${line.part_number}</strong> · Coverage correcto: 12 meses</span>
        </div>`;
      }
    });
    html += '</div>';
    ca.innerHTML = html;
  } else {
    bpl.classList.add('hidden');
    let html = '<div style="display:flex;flex-direction:column;gap:6px">';
    lines.forEach(line=>{
      html += `<div style="display:flex;align-items:center;gap:8px;background:#F0FDF4;border:0.5px solid #BBF7D0;border-radius:6px;padding:7px 12px;font-size:12px">
        <span style="color:#16A34A;font-size:14px">✓</span>
        <span><strong>${line.part_number}</strong> · Coverage correcto: 12 meses</span>
      </div>`;
    });
    html += '</div>';
    ca.innerHTML = html;
  }

  // Sincronizar checkbox de revisión
  const chk = document.getElementById('review-check-input');
  chk.checked = !!allQuotes[qn]._checked;
  updateReviewHint();

  recalc();
  document.getElementById('card-editor').classList.remove('hidden');
  document.getElementById('card-editor').scrollIntoView({behavior:'smooth',block:'start'});
  renderQuoteList();
}

function recalc(){
  const g = getG();
  const bp = parseFloat(document.getElementById('e-bp').value)||g.bp;
  const extra = parseFloat(document.getElementById('e-extra').value)||0;
  const useRenew  = document.getElementById('e-cvr-renew').checked;
  const useOntime = document.getElementById('e-cvr-ontime').checked;
  let tMep=0, tCanal=0, tRv=0, tOt=0, tFinal=0;
  const tbody = document.getElementById('editor-tbody');
  tbody.innerHTML='';

  currentLines.forEach(line=>{
    const c = calcLine(line, bp, g.cvrRenew, g.cvrOntime, extra, useRenew, useOntime);
    tMep+=c.mep; tCanal+=c.costoCanal; tRv+=c.rv; tOt+=c.ot; tFinal+=c.finalBP;
    line._calc = c;
    const desc = (line.part_description||'—').substring(0,42)+'...';
    const cov = (line.coverage_dates||'').replace(' To ','→');
    const tr = document.createElement('tr');
    tr.innerHTML=`
      <td><strong>${line.part_number||'—'}</strong></td>
      <td style="font-size:11px">${desc}</td>
      <td style="font-size:11px;white-space:nowrap">${cov}</td>
      <td style="text-align:right">${line.quantity||1}</td>
      <td style="text-align:right;color:#64748B">$${fmt(c.mep)}</td>
      <td style="text-align:right;color:#DC2626">-$${fmt(c.mep-c.costoCanal)}</td>
      <td style="text-align:right;color:#16A34A">${useRenew?'-$'+fmt(c.rv):'—'}</td>
      <td style="text-align:right;color:#16A34A">${useOntime&&line.ontime?'-$'+fmt(c.ot):'—'}</td>
      <td style="text-align:right;font-weight:600">$${fmt(c.finalBP)}</td>
    `;
    tbody.appendChild(tr);
  });

  const tr = document.createElement('tr');
  tr.className='tot-row';
  tr.innerHTML=`<td colspan="4" style="text-align:right">Total</td>
    <td style="text-align:right">$${fmt(tMep)}</td>
    <td style="text-align:right">-$${fmt(tMep-tCanal)}</td>
    <td style="text-align:right">-$${fmt(tRv)}</td>
    <td style="text-align:right">-$${fmt(tOt)}</td>
    <td style="text-align:right">$${fmt(tFinal)}</td>`;
  tbody.appendChild(tr);

  document.getElementById('editor-summary').innerHTML=`
    <div class="scard"><div class="scard-lbl">Precio MEP total</div><div class="scard-val">$${fmt(tMep)}</div></div>
    <div class="scard"><div class="scard-lbl">Total descuentos</div><div class="scard-val red">-$${fmt(tMep-tFinal)}</div></div>
    <div class="scard"><div class="scard-lbl">Precio final BP</div><div class="scard-val blue">$${fmt(tFinal)}</div></div>
  `;
}

function closeEditor(){
  // Guardar parámetros actuales del editor para usarlos en descarga masiva
  if(currentQ){
    const g = getG();
    quoteParams[currentQ] = {
      bp:         parseFloat(document.getElementById('e-bp').value)||g.bp,
      ingram:     parseFloat(document.getElementById('e-ingram').value)||g.ingram,
      extra:      parseFloat(document.getElementById('e-extra').value)||0,
      useRenew:   document.getElementById('e-cvr-renew').checked,
      useOntime:  document.getElementById('e-cvr-ontime').checked,
      cvrRenew:   g.cvrRenew,
      cvrOntime:  g.cvrOntime,
      cvrRenewIng:g.cvrRenewIng,
      cvrOntimeIng:g.cvrOntimeIng,
    };
  }
  document.getElementById('card-editor').classList.add('hidden');
  currentQ = null;
  renderQuoteList();
}

async function downloadQuote(){
  const btn = document.getElementById('btn-dl');
  const status = document.getElementById('dl-status');
  btn.disabled=true; btn.innerHTML='<span class="spinner"></span> Generando...';
  status.textContent='';
  try{
    const g = getG();
    const bp     = parseFloat(document.getElementById('e-bp').value)||g.bp;
    const ingram = parseFloat(document.getElementById('e-ingram').value)||g.ingram;
    const extra  = parseFloat(document.getElementById('e-extra').value)||0;
    const useRenew  = document.getElementById('e-cvr-renew').checked;
    const useOntime = document.getElementById('e-cvr-ontime').checked;
    const fd = new FormData();
    fd.append('file', selectedFile);
    fd.append('bp_margin', bp);
    fd.append('ingram_margin', ingram);
    fd.append('extra_discount', extra);
    fd.append('cvr_renew_bp',      useRenew  ? g.cvrRenew    : 0);
    fd.append('cvr_renew_ingram',  useRenew  ? g.cvrRenewIng : 0);
    fd.append('cvr_ontime_bp',     useOntime ? g.cvrOntime   : 0);
    fd.append('cvr_ontime_ingram', useOntime ? g.cvrOntimeIng: 0);
    fd.append('quote_filter', currentQ);
    const res = await fetch(API+'/renewals/generate-quote',{method:'POST',body:fd});
    if(!res.ok) throw new Error((await res.json()).error||'Error del servidor');
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const cd   = res.headers.get('Content-Disposition')||'';
    const m    = cd.match(/filename="?([^"]+)"?/);
    a.download = m?m[1]:`Quote_${currentQ}.zip`;
    a.href=url; a.click(); URL.revokeObjectURL(url);
    allQuotes[currentQ]._downloaded = true;
    // Guardar parámetros usados
    quoteParams[currentQ] = {bp,ingram,extra,useRenew,useOntime,
      cvrRenew:g.cvrRenew,cvrOntime:g.cvrOntime,
      cvrRenewIng:g.cvrRenewIng,cvrOntimeIng:g.cvrOntimeIng};
    status.textContent='✓ Descargado individualmente';
    updateReviewHint();
    renderQuoteList();
  }catch(e){
    status.textContent='Error: '+e.message;
  }finally{
    btn.disabled=false; btn.innerHTML='⬇ Descargar PDF + Excel';
  }
}

async function downloadPlanilla(){
  const btn = document.getElementById('btn-planilla');
  const status = document.getElementById('dl-status');
  btn.disabled=true; btn.textContent='Generando...';
  try{
    const fd = new FormData();
    fd.append('file', selectedFile);
    fd.append('quote_filter', currentQ);
    const res = await fetch(API+'/renewals/generate-planilla',{method:'POST',body:fd});
    if(!res.ok) throw new Error('Error');
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.download = `Planilla_IBM_${currentQ}.xlsx`;
    a.href=url; a.click(); URL.revokeObjectURL(url);
    status.textContent='✓ Planilla descargada';
  }catch(e){
    status.textContent='Error: '+e.message;
  }finally{
    btn.disabled=false; btn.innerHTML='⬇ Planilla IBM';
  }
}

function toggleAudit(){
  const panel = document.getElementById('audit-panel');
  const arrow = document.getElementById('audit-arrow');
  if(panel.classList.contains('hidden')){
    panel.classList.remove('hidden');
    arrow.textContent = '▲ Colapsar';
    buildAuditTable();
  } else {
    panel.classList.add('hidden');
    arrow.textContent = '▼ Expandir';
  }
}

function buildAuditTable(){
  const g = getG();
  const bp     = parseFloat(document.getElementById('e-bp').value)||g.bp;
  const ingram = parseFloat(document.getElementById('e-ingram').value)||g.ingram;
  const extra  = parseFloat(document.getElementById('e-extra').value)||0;
  const useRenew  = document.getElementById('e-cvr-renew').checked;
  const useOntime = document.getElementById('e-cvr-ontime').checked;

  const XML_FIELDS = [
    {key:'quote_number',label:'Quote Number'},{key:'part_number',label:'Part Number'},
    {key:'part_description',label:'Descripción'},{key:'quantity',label:'Qty'},
    {key:'unit_price',label:'Unit Price (MEP)'},{key:'extended_price',label:'Extended Price'},
    {key:'currency',label:'Moneda'},{key:'coverage_dates',label:'Coverage Dates'},
    {key:'renewal_due_date',label:'Renewal Due Date'},
    {key:'renewal_line_item_start_date',label:'Start Date'},
    {key:'site',label:'Site (Cliente)'},{key:'ibm_customer_number',label:'ICN'},
    {key:'reseller',label:'Reseller'},{key:'reseller_authorization',label:'Auth Category'},
    {key:'customer_set_designation',label:'Customer Set'},
    {key:'price_level',label:'Price Level'},
    {key:'eligible_for_government_reseller_program',label:'Gobierno'},
    {key:'select_territory_accelerator',label:'Select Territory'},
    {key:'ibm_renewal_contact',label:'IBM Contact'},{key:'site_contact',label:'Site Contact'},
    {key:'agreement_information',label:'Agreement'},
  ];
  const CALC_FIELDS = [
    {key:'_bp',label:`Margen BP (${bp}%)`},
    {key:'_dcto_canal',label:'Dcto Canal $'},
    {key:'_costo_canal',label:'Costo Canal'},
    {key:'_cvr_renew',label:`CVR Renew BP (${g.cvrRenew}%)`},
    {key:'_cvr_ontime',label:`CVR OnTime BP (${g.cvrOntime}%)`},
    {key:'_extra',label:`Dcto Extra (${extra}%)`},
    {key:'_final_bp',label:'Precio Final BP'},
    {key:'_costo_ingram',label:`Costo Ingram (${ingram}%)`},
    {key:'_ontime',label:'Renueva a tiempo'},
    {key:'_coverage_ok',label:'Coverage OK'},
  ];

  const thead = document.getElementById('audit-thead');
  thead.innerHTML='';
  const hrow = document.createElement('tr');
  hrow.innerHTML = '<th class="audit-th-xml" style="min-width:140px">Campo XML</th>';
  currentLines.forEach((line,i)=>{
    hrow.innerHTML += `<th class="audit-th-xml" style="min-width:120px">Línea ${i+1}<br><span style="font-weight:400;font-size:10px">${line.part_number||''}</span></th>`;
  });
  thead.appendChild(hrow);

  const tbody = document.getElementById('audit-tbody');
  tbody.innerHTML='';

  const secXml = document.createElement('tr');
  secXml.innerHTML=`<td class="audit-section" colspan="${currentLines.length+1}">Datos originales del XML</td>`;
  tbody.appendChild(secXml);

  XML_FIELDS.forEach(f=>{
    const tr=document.createElement('tr');
    let row=`<td class="audit-td audit-xml" style="font-weight:500;color:#475569">${f.label}</td>`;
    currentLines.forEach(line=>{
      const val=line[f.key]!==undefined?String(line[f.key]):'—';
      row+=`<td class="audit-td audit-xml">${val||'—'}</td>`;
    });
    tr.innerHTML=row; tbody.appendChild(tr);
  });

  const secCalc=document.createElement('tr');
  secCalc.innerHTML=`<td class="audit-section" colspan="${currentLines.length+1}">Cálculos aplicados</td>`;
  tbody.appendChild(secCalc);

  const calcs=currentLines.map(line=>{
    const c=calcLine(line,bp,g.cvrRenew,g.cvrOntime,extra,useRenew,useOntime);
    const mep=parseFloat(line.extended_price||0);
    const hasCovErr=(covAlerts[currentQ]||[]).find(e=>e.part_number===line.part_number);
    return{
      _bp:bp+'%',
      _dcto_canal:'$'+fmt(mep-c.costoCanal),
      _costo_canal:'$'+fmt(c.costoCanal),
      _cvr_renew:useRenew?'-$'+fmt(c.rv):'No aplica',
      _cvr_ontime:useOntime&&line.ontime?'-$'+fmt(c.ot):'No aplica',
      _extra:extra>0?'-$'+fmt(c.costoCanal*(extra/100)):'No aplica',
      _final_bp:'$'+fmt(c.finalBP),
      _costo_ingram:'$'+fmt(c.costoCanal*(1-ingram/100)),
      _ontime:line.ontime?'✓ Sí':'✗ No',
      _coverage_ok:hasCovErr?`⚠ ${hasCovErr.months} mes(es)`:'✓ 12 meses',
    };
  });

  CALC_FIELDS.forEach(f=>{
    const tr=document.createElement('tr');
    let row=`<td class="audit-td audit-calc" style="font-weight:500;color:#0066CC">${f.label}</td>`;
    calcs.forEach(c=>{
      const val=c[f.key]||'—';
      const isOk=val.startsWith('✓');
      const isErr=val.startsWith('⚠')||val.startsWith('✗');
      const color=isOk?'#166534':isErr?'#991B1B':'#0C447C';
      row+=`<td class="audit-td audit-calc" style="color:${color}">${val}</td>`;
    });
    tr.innerHTML=row; tbody.appendChild(tr);
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

# ── Nota: agregar al ibm_renewals_blueprint.py ────────────────────────────────
# Este endpoint debe añadirse en ibm_renewals_blueprint.py
BULK_ENDPOINT = '''
import json as _json

@ibm_renewals_bp.post('/generate-bulk')
def generate_bulk():
    """
    Genera un ZIP maestro con todos los quotes chequeados.
    Estructura: {canal}/{año}/{reseller}/  →  PDF + Excel + Planilla
    """
    f, err, code = _get_file_from_request()
    if err:
        return err, code

    try:
        bp_margin         = float(request.form.get('bp_margin', 3))
        ingram_margin     = float(request.form.get('ingram_margin', 2.1))
        cvr_renew_bp      = float(request.form.get('cvr_renew_bp', 2))
        cvr_renew_ingram  = float(request.form.get('cvr_renew_ingram', 2))
        cvr_ontime_bp     = float(request.form.get('cvr_ontime_bp', 2))
        cvr_ontime_ingram = float(request.form.get('cvr_ontime_ingram', 1))
    except (ValueError, TypeError):
        return jsonify({"success": False, "error": "Márgenes inválidos"}), 400

    # Quotes a incluir
    try:
        quote_numbers = _json.loads(request.form.get('quote_numbers', '[]'))
        quote_params  = _json.loads(request.form.get('quote_params',  '{}'))
    except Exception:
        quote_numbers = []
        quote_params  = {}

    if not quote_numbers:
        return jsonify({"success": False, "error": "No se enviaron quotes para generar"}), 400

    xml_content = f.read()
    try:
        renewals_all = parse_renewal_xml(xml_content)
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

    groups = _group_by_quote(renewals_all)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    year      = datetime.now().year

    # ZIP maestro nombrado por fecha
    zip_name = f'IBM_SS_Renovaciones_{timestamp}_{len(quote_numbers)}quotes.zip'
    zip_path = os.path.join(OUTPUT_DIR, zip_name)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for qn in quote_numbers:
            if qn not in groups:
                continue
            lines = groups[qn]

            # Parámetros específicos del quote (si fue ajustado en el editor)
            p = quote_params.get(qn, {})
            q_bp          = float(p.get('bp',          bp_margin))
            q_ingram      = float(p.get('ingram',       ingram_margin))
            q_cvr_renew   = float(p.get('cvrRenew',     cvr_renew_bp))
            q_cvr_ontime  = float(p.get('cvrOntime',    cvr_ontime_bp))
            q_cvr_ri      = float(p.get('cvrRenewIng',  cvr_renew_ingram))
            q_cvr_oi      = float(p.get('cvrOntimeIng', cvr_ontime_ingram))

            lines = _apply_margins(lines, q_bp, q_ingram,
                                   q_cvr_renew, q_cvr_ri,
                                   q_cvr_ontime, q_cvr_oi)

            reseller_name  = _extract_reseller_name(lines[0].get('reseller', ''))
            customer_short = _extract_customer_short(lines[0].get('site', ''))
            customer_full  = lines[0].get('site', 'Cliente')

            base_name = f'S&S - {customer_short} - {reseller_name} - {year}'

            # Carpeta: {canal}/{año}/
            folder = f'{_safe_filename(reseller_name)}/{year}/'

            # PDF
            pdf_path = os.path.join(OUTPUT_DIR, f'{_safe_filename(base_name)}.pdf')
            generate_quote_pdf(lines, customer_full,
                               f'Renovación S&S IBM — {customer_short}',
                               q_bp, pdf_path,
                               quote_number=qn,
                               cvr_renew_bp=q_cvr_renew,
                               cvr_ontime_bp=q_cvr_ontime)
            zf.write(pdf_path, folder + os.path.basename(pdf_path))

            # Excel
            excel_path = os.path.join(OUTPUT_DIR, f'{_safe_filename(base_name)}.xlsx')
            generate_internal_excel(lines, qn, excel_path)
            zf.write(excel_path, folder + os.path.basename(excel_path))

            # Planilla (solo si hay errores de coverage)
            coverage_errors = check_coverage_errors(lines)
            if coverage_errors:
                planilla_name = (f'Planilla de renovación SSA y MX - '
                                 f'{customer_short} - {reseller_name} - 12 Meses - {year}')
                planilla_path = os.path.join(OUTPUT_DIR, f'{_safe_filename(planilla_name)}.xlsx')
                generate_planilla(lines, qn, planilla_path)
                zf.write(planilla_path, folder + os.path.basename(planilla_path))

    return send_file(zip_path, mimetype='application/zip',
                     as_attachment=True, download_name=zip_name)
'''
