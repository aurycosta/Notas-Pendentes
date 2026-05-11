// Painel de Notas Pendentes (v2)
// Colunas no XLSX (sem cabeçalho), por posição:
// 0 origem | 1 destino | 2 razao social | 3 numero NFe | 4 serie | 5 chave acesso | 6 data emissao | 7 valor | 8 CFOP/CGO | 9 periodo

let notas = [];
let notasFiltradas = [];

const $ = (id) => document.getElementById(id);

const fileInput = $("fileInput");
const filtroLoja = $("filtroLoja");
const filtroStatus = $("filtroStatus");
const dataRef = $("dataRef");
const busca = $("busca");

const btnPng = $("btnPng");
const btnPdfLoja = $("btnPdfLoja");
const btnPdfTudo = $("btnPdfTudo");
const btnCopiarMsg = $("btnCopiarMsg");

const tabelaBody = document.querySelector("#tabela tbody");

const modal = $("modal");
const detalhes = $("detalhes");
$("btnFechar").onclick = () => modal.classList.remove("open");
modal.addEventListener("click", (e)=>{ if(e.target === modal) modal.classList.remove("open"); });

let chartQtd, chartValor;

// ===== Helpers =====
function todayISO(){
  const d = new Date();
  const tz = d.getTimezoneOffset()*60000;
  return new Date(d - tz).toISOString().slice(0,10);
}

function formatBRL(v){
  const n = Number(v || 0);
  return n.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
}

function parseExcelDate(x){
  // x pode ser Date, number (serial), string.
  if(!x) return null;
  if(x instanceof Date) return x;
  if(typeof x === "number"){
    // Excel serial -> JS date (UTC-ish)
    const utc_days = Math.floor(x - 25569);
    const utc_value = utc_days * 86400;                                        
    const date_info = new Date(utc_value * 1000);
    // manter meia-noite local
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
  }
  if(typeof x === "string"){
    // tenta dd/mm/yyyy
    const m = x.match(/(\d{2})\/(\d{2})\/(\d{4})/);
    if(m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
    const d = new Date(x);
    return isNaN(d) ? null : d;
  }
  return null;
}

function diffDays(fromDate, toDate){
  const a = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
  const b = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
  return Math.floor((b - a) / (1000*60*60*24));
}

function statusByDays(days){
  // Regra com prazo de 3 dias:
  // Dias desde emissão (D):
  // 0-1 = OK
  // 2-3 = Alerta (ainda dentro do prazo)
  // >=4 = Vencido (passou do prazo)
  if(days <= 1) return { key:"ok", label:"OK", icon:"✅", diasVencida: 0 };
  if(days <= 3) return { key:"alerta", label:"Alerta", icon:"⚠️", diasVencida: 0 };
  const diasVencida = days - 3;
  return { key:"vencido", label:`Vencido (${diasVencida}d)`, icon:"❌", diasVencida };
}

// ===== Import XLSX =====
fileInput.addEventListener("change", (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;

  if(typeof XLSX === "undefined"){
    alert("Biblioteca XLSX não carregou. Verifique sua internet (CDN) ou me peça versão offline.");
    return;
  }

  const reader = new FileReader();
  reader.onload = (evt)=>{
    const data = evt.target.result;
    const wb = XLSX.read(data, { type:"array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });

    // remove linhas vazias
    const clean = rows.filter(r => r && r.some(v => String(v).trim() !== ""));
    notas = clean.map((r, idx)=>({
      __i: idx,
      origem: String(r[0] ?? "").trim(),
      destino: String(r[1] ?? "").trim(),
      razao: String(r[2] ?? "").trim(),
      nfe: String(r[3] ?? "").trim(),
      serie: String(r[4] ?? "").trim(),
      chave: String(r[5] ?? "").trim(),
      emissao: parseExcelDate(r[6]),
      valor: Number(String(r[7]).replace(",", ".")) || 0,
      cgo: String(r[8] ?? "").trim(),
      periodo: String(r[9] ?? "").trim(),
    })).filter(n => n.origem || n.destino || n.nfe);

    // seta data ref para hoje (padrão)
    dataRef.value = todayISO();

    carregarFiltroLojas();
    aplicarEAtualizar();
  };

  reader.readAsArrayBuffer(f);
});

function carregarFiltroLojas(){
  const lojas = [...new Set(notas.map(n => n.destino).filter(Boolean))].sort();
  filtroLoja.innerHTML = '<option value="">Todas as lojas</option>' + lojas.map(l=>`<option value="${escapeHtml(l)}">${escapeHtml(l)}</option>`).join("");
}

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}

// ===== Filtros =====
[filtroLoja, filtroStatus, dataRef, busca].forEach(el=>{
  el.addEventListener("input", aplicarEAtualizar);
});

function aplicarEAtualizar(){
  const loja = filtroLoja.value;
  const st = filtroStatus.value;
  const q = (busca.value || "").trim().toLowerCase();

  const ref = dataRef.value ? new Date(dataRef.value + "T00:00:00") : new Date();

  notasFiltradas = notas.map(n=>{
    const daysDesde = n.emissao ? diffDays(n.emissao, ref) : null;
    const s = (daysDesde === null) ? { key:"vencido", label:"Sem data", icon:"❓", diasVencida: 0 } : statusByDays(daysDesde);

    const diasVencida = (daysDesde === null) ? 0 : (s.key === "vencido" ? s.diasVencida : 0);
    const diasMostrar = (daysDesde === null) ? "" : (s.key === "vencido" ? diasVencida : daysDesde);
    const diasTexto = (daysDesde === null) ? "" : (s.key === "vencido" ? `${diasVencida} dia(s) vencida` : `${daysDesde} dia(s) desde emissão`);

    return {
      ...n,
      diasDesde: daysDesde ?? "",
      diasVencida,
      dias: diasMostrar,
      diasTexto,
      status: s.key,
      statusLabel: s.label,
      statusIcon: s.icon
    };
  }).filter(n=>{
    if(loja && n.destino !== loja) return false;
    if(st && n.status !== st) return false;
    if(q){
      const blob = [n.origem,n.destino,n.razao,n.nfe,n.serie,n.chave,n.cgo,n.periodo].join(" ").toLowerCase();
      if(!blob.includes(q)) return false;
    }
    return true;
  });

  renderTabela();
  renderKPIs();
  renderMensagem();
  renderGraficos();
}
function fmtDate(d){
  if(!d) return "";
  return d.toLocaleDateString("pt-BR");
}

function renderTabela(){
  tabelaBody.innerHTML = "";
  for(const n of notasFiltradas){
    const tr = document.createElement("tr");
    tr.className = n.status;

    tr.innerHTML = `
      <td>${escapeHtml(n.origem)}</td>
      <td>${escapeHtml(n.destino)}</td>
      <td>${escapeHtml(n.razao)}</td>
      <td>${escapeHtml(n.nfe)}</td>
      <td>${escapeHtml(n.serie)}</td>
      <td class="mono">${escapeHtml(n.chave)}</td>
      <td>${escapeHtml(fmtDate(n.emissao))}</td>
      <td><strong>${escapeHtml(String(n.dias))}</strong><div class="sub">${escapeHtml(n.diasTexto || "")}</div></td>
      <td><span class="badge ${n.status}">${escapeHtml(n.statusIcon)} ${escapeHtml(n.statusLabel)}</span></td>
      <td><strong>${formatBRL(n.valor)}</strong></td>
    `;

    tr.addEventListener("click", ()=>{
      detalhes.textContent = JSON.stringify({
        Origem: n.origem,
        Destino: n.destino,
        "Razão Social": n.razao,
        "NF-e": n.nfe,
        Série: n.serie,
        "Chave de Acesso": n.chave,
        "Data de Emissão": fmtDate(n.emissao),
        "Dias": n.diasTexto || n.dias,
        Status: n.statusLabel,
        Valor: formatBRL(n.valor),
        CGO: n.cgo,
        Período: n.periodo
      }, null, 2);
      modal.classList.add("open");
    });

    tabelaBody.appendChild(tr);
  }
}

// ===== KPIs =====
function renderKPIs(){
  const total = notasFiltradas.reduce((acc,n)=>acc+n.valor,0);
  const ok = notasFiltradas.filter(n=>n.status==="ok").length;
  const alerta = notasFiltradas.filter(n=>n.status==="alerta").length;
  const vencido = notasFiltradas.filter(n=>n.status==="vencido").length;

  $("totalValor").textContent = total.toLocaleString("pt-BR", { minimumFractionDigits:2, maximumFractionDigits:2 });
  $("totalQtde").textContent = String(notasFiltradas.length);
  $("qtdeOk").textContent = String(ok);
  $("qtdeAlerta").textContent = String(alerta);
  $("qtdeVencido").textContent = String(vencido);
}

// ===== Mensagem =====
function renderMensagem(){
  const loja = filtroLoja.value;
  const ref = dataRef.value ? new Date(dataRef.value + "T00:00:00") : new Date();
  const refBR = ref.toLocaleDateString("pt-BR");

  const pend = notasFiltradas.filter(n => n.status !== "ok");
  const vencidas = pend.filter(n => n.status === "vencido");
  const alertas = pend.filter(n => n.status === "alerta");

  const alvo = loja ? `na unidade ${loja}` : "nas unidades";
  const resumo = `Vencidas: ${vencidas.length} | Alerta: ${alertas.length}`;

  const msg = [
    "Bom dia! Espero que esteja bem.",
    `Observei que há pendências de notas fiscais ${alvo} (ref.: ${refBR}).`,
    resumo + ".",
    "Você poderia verificar, por favor, para que possamos efetivar dentro do prazo?",
    "Obrigado!"
  ].join("\n");

  $("msg").value = msg;
}

btnCopiarMsg.addEventListener("click", async ()=>{
  const texto = $("msg").value;
  try{
    await navigator.clipboard.writeText(texto);
    toast("Mensagem copiada ✅");
  }catch(e){
    $("msg").focus();
    $("msg").select();
    document.execCommand("copy");
    toast("Mensagem copiada ✅");
  }
});

// ===== Gráficos =====
function renderGraficos(){
  const ok = notasFiltradas.filter(n=>n.status==="ok");
  const alerta = notasFiltradas.filter(n=>n.status==="alerta");
  const vencido = notasFiltradas.filter(n=>n.status==="vencido");

  const qtd = [ok.length, alerta.length, vencido.length];
  const val = [
    ok.reduce((a,n)=>a+n.valor,0),
    alerta.reduce((a,n)=>a+n.valor,0),
    vencido.reduce((a,n)=>a+n.valor,0)
  ];

  // Qtd
  if(chartQtd) chartQtd.destroy();
  chartQtd = new Chart($("chartQtd"), {
    type: "doughnut",
    data: { labels:["OK","Alerta","Vencido"], datasets:[{ data:qtd }] },
    options: { responsive:true, plugins:{ legend:{ position:"bottom" } } }
  });

  // Valor
  if(chartValor) chartValor.destroy();
  chartValor = new Chart($("chartValor"), {
    type: "bar",
    data: { labels:["OK","Alerta","Vencido"], datasets:[{ label:"Valor (R$)", data: val }] },
    options: {
      responsive:true,
      plugins:{ legend:{ display:false } },
      scales:{
        y:{ ticks:{ callback:(v)=>Number(v).toLocaleString("pt-BR") } }
      }
    }
  });
}

// ===== Export PNG =====
btnPng.addEventListener("click", async ()=>{
  if(typeof html2canvas === "undefined"){
    alert("html2canvas não carregou (CDN). Verifique sua internet.");
    return;
  }
  const node = $("painel");
  const canvas = await html2canvas(node, { scale: 2, useCORS: true });
  const a = document.createElement("a");
  a.href = canvas.toDataURL("image/png");
  a.download = `painel_notas_${todayISO()}.png`;
  a.click();
});

// ===== Export PDF =====
function buildRowsForPDF(rows){
  return rows.map(n=>[
    n.origem, n.destino, n.razao,
    n.nfe, n.serie,
    fmtDate(n.emissao),
    String(n.dias),
    n.statusLabel,
    formatBRL(n.valor)
  ]);
}

function exportPDF(titulo, rows){
  if(!window.jspdf || !window.jspdf.jsPDF){
    alert("jsPDF não carregou (CDN). Verifique sua internet.");
    return;
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:"landscape", unit:"pt", format:"a4" });

  doc.setFontSize(14);
  doc.text(titulo, 40, 40);

  const total = rows.reduce((a,n)=>a+n.valor,0);
  doc.setFontSize(10);
  doc.text(`Gerado em: ${new Date().toLocaleString("pt-BR")}  |  Total: ${formatBRL(total)}`, 40, 60);

  doc.autoTable({
    startY: 80,
    head: [[ "Origem","Destino","Razão Social","NF-e","Série","Emissão","Dias","Status","Valor" ]],
    body: buildRowsForPDF(rows),
    styles: { fontSize: 8, cellPadding: 4 },
    headStyles: { fillColor: [17,24,39] },
    margin: { left: 40, right: 40 }
  });

  doc.save(`${titulo.replace(/\s+/g,'_').toLowerCase()}_${todayISO()}.pdf`);
}

btnPdfLoja.addEventListener("click", ()=>{
  const loja = filtroLoja.value;
  if(!loja){
    alert("Selecione uma loja destino para exportar o PDF por loja.");
    return;
  }
  exportPDF(`Pendências - ${loja}`, notasFiltradas);
});

btnPdfTudo.addEventListener("click", ()=>{
  exportPDF("Pendências - Consolidado", notasFiltradas);
});

// ===== Toast simples =====
let toastTimer=null;
function toast(msg){
  let el = document.getElementById("toast");
  if(!el){
    el = document.createElement("div");
    el.id = "toast";
    el.style.position="fixed";
    el.style.bottom="16px";
    el.style.right="16px";
    el.style.background="#111827";
    el.style.color="#fff";
    el.style.padding="10px 12px";
    el.style.borderRadius="12px";
    el.style.boxShadow="0 10px 30px rgba(0,0,0,.18)";
    el.style.zIndex=9999;
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.display="block";
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=>{ el.style.display="none"; }, 2200);
}
