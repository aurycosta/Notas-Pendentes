import { initializeApp } from "https://www.gstatic.com/firebasejs/12.14.0/firebase-app.js";
import { getDatabase, ref, get, update } from "https://www.gstatic.com/firebasejs/12.14.0/firebase-database.js";

const firebaseConfig = {
  apiKey: "AIzaSyDMgocHaQtH07v-aQc4tC4JcV3jNxY3tX0",
  authDomain: "controle-armarios-14351.firebaseapp.com",
  databaseURL: "https://controle-armarios-14351-default-rtdb.firebaseio.com",
  projectId: "controle-armarios-14351",
  storageBucket: "controle-armarios-14351.firebasestorage.app",
  messagingSenderId: "171871598145",
  appId: "1:171871598145:web:33728a54894b65ef7838c0",
  measurementId: "G-7L3G5D9DRB"
};

const app = initializeApp(firebaseConfig);
const db = getDatabase(app);

// Estado Global
let notas = []; // Notas do painel atual (puxadas do banco ou recém importadas)
let notasFiltradas = [];
let historico = {}; // Guarda { "timestamp_da_importacao": [array de notas] }

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

// Telas
const painelPrincipal = $("painel");
const telaHistorico = $("telaHistorico");

let chartQtd, chartValor;

// ===== Carregamento Inicial (O F5 NÃO APAGA MAIS) =====
window.addEventListener("DOMContentLoaded", () => {
  // Vamos usar um novo nó no banco chamado "painel_persistente" para não bagunçar testes anteriores
  get(ref(db, 'painel_persistente')).then((snapshot) => {
    if(snapshot.exists()) {
      const data = snapshot.val();
      
      // 1. Restaura o estado atual (painel)
      if(data.estado_atual) {
        notas = data.estado_atual.map(n => ({
          ...n,
          emissao: n.emissao ? new Date(n.emissao) : null
        }));
      }
      
      // 2. Restaura todo o histórico de planilhas antigas
      if(data.historico) {
        historico = data.historico; 
      }

      toast("Dados restaurados com sucesso! Banco carregado.");
      carregarFiltroLojas();
      aplicarEAtualizar();
    } else {
      toast("Banco vazio. Importe sua primeira planilha!");
    }
  }).catch((err) => {
    console.error("Erro ao ler DB:", err);
    toast("Erro ao tentar restaurar dados do Firebase.");
  });
});

// ===== Trocar Abas (Painel <-> Histórico) =====
$("btnVerHistorico").addEventListener("click", () => {
  painelPrincipal.style.display = "none";
  telaHistorico.style.display = "block";
  renderHistoricoDatas();
});

$("btnVoltarPainel").addEventListener("click", () => {
  telaHistorico.style.display = "none";
  painelPrincipal.style.display = "block";
});

// ===== Helpers Genéricos =====
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
  if(!x) return null;
  if(x instanceof Date) return x;
  if(typeof x === "number"){
    const utc_days = Math.floor(x - 25569);
    const utc_value = utc_days * 86400;                                        
    const date_info = new Date(utc_value * 1000);
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
  }
  if(typeof x === "string"){
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
  if(days <= 1) return { key:"ok", label:"OK", icon:"✅" };
  if(days <= 3) return { key:"alerta", label:"Alerta", icon:"⚠️" };
  return { key:"vencido", label:`Vencido`, icon:"❌" };
}

// Chave unificada para ter precisão de rastrear notas repetidas no histórico
function getNotaKey(n) {
  if (n.chave && String(n.chave).trim() !== "") return String(n.chave).trim();
  return `${n.nfe}_${n.origem}_${n.destino}`;
}

// ===== Importar e Salvar no Firebase =====
fileInput.addEventListener("change", (e)=>{
  const f = e.target.files?.[0];
  if(!f) return;

  toast("Processando XLSX e gravando no Banco de Dados...");

  const reader = new FileReader();
  reader.onload = (evt)=>{
    const data = evt.target.result;
    const wb = window.XLSX.read(data, { type:"array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = window.XLSX.utils.sheet_to_json(ws, { header:1, defval:"" });

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

    // Preparar para salvar no Firebase (converte datas para número)
    const timestampAgora = Date.now();
    const notasParaSalvar = notas.map(n => ({
      ...n,
      emissao: n.emissao ? n.emissao.getTime() : null
    }));

    // Gravamos simultaneamente: 1) Substituímos o painel atual e 2) Criamos uma cópia no Histórico
    const updates = {};
    updates['painel_persistente/estado_atual'] = notasParaSalvar;
    updates[`painel_persistente/historico/${timestampAgora}`] = notasParaSalvar;

    update(ref(db), updates).then(() => {
       toast("Importação concluída e salva no Histórico!");
       historico[timestampAgora] = notasParaSalvar; // Sincroniza memória local
       
       dataRef.value = todayISO();
       carregarFiltroLojas();
       aplicarEAtualizar();
       fileInput.value = ""; // limpa o input
    }).catch(err => {
       console.error(err);
       toast("Erro ao gravar dados no Firebase!");
    });
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

// ===== Filtros & Renderização Painel =====
[filtroLoja, filtroStatus, dataRef, busca].forEach(el=>{
  el.addEventListener("input", aplicarEAtualizar);
});

function aplicarEAtualizar(){
  const loja = filtroLoja.value;
  const st = filtroStatus.value;
  const q = (busca.value || "").trim().toLowerCase();
  const refDate = dataRef.value ? new Date(dataRef.value + "T00:00:00") : new Date();

  notasFiltradas = notas.map(n=>{
    const daysDesde = n.emissao ? diffDays(n.emissao, refDate) : null;
    const s = (daysDesde === null) ? { key:"vencido", label:"Sem data", icon:"❓" } : statusByDays(daysDesde);

    return {
      ...n,
      diasDesde: daysDesde ?? "",
      status: s.key,
      statusLabel: s.label,
      statusIcon: s.icon
    };
  }).filter(n=>{
    if(loja && n.destino !== loja) return false;
    if(st && n.status !== st) return false;
    if(q){
      const blob = [n.origem,n.destino,n.razao,n.nfe,n.chave].join(" ").toLowerCase();
      if(!blob.includes(q)) return false;
    }
    return true;
  });

  renderTabela();
  renderKPIs();
  renderMensagem();
  renderGraficos();
}

function renderTabela(){
  tabelaBody.innerHTML = "";
  
  // Array de datas do histórico, ordenadas da mais antiga (0) para a mais recente
  const timestampsHistorico = Object.keys(historico).sort((a,b) => a - b);

  for(const n of notasFiltradas){
    // ====== LÓGICA DE RECORRÊNCIA (A MÁGICA AQUI) ======
    let aparicoes = 0;
    let timestampPrimeiraAparicao = null;
    const chaveUnicaDaNota = getNotaKey(n);

    timestampsHistorico.forEach(tsStr => {
      // Se essa nota existia na importação desse dia do histórico...
      if(historico[tsStr].some(hx => getNotaKey(hx) === chaveUnicaDaNota)) {
         aparicoes++;
         if(!timestampPrimeiraAparicao) timestampPrimeiraAparicao = parseInt(tsStr); // Guarda a data mais antiga
      }
    });

    // Se só apareceu 1 vez (hoje) é "Nova". Senão, mostra há quantas importações ela se repete.
    let htmlRecorrencia = `<span class="badge" style="background:#dcfce7; color:#166534">✨ Nova</span>`;
    let textoRecorrenciaExport = "Nova";
    
    if(aparicoes > 1 && timestampPrimeiraAparicao) {
      const dataBr = new Date(timestampPrimeiraAparicao).toLocaleDateString("pt-BR");
      htmlRecorrencia = `<span class="badge" style="background:#fee2e2; color:#991b1b" title="Presente desde ${dataBr}">📌 ${aparicoes}x seguidas</span>`;
      textoRecorrenciaExport = `${aparicoes}x (Desde ${dataBr})`;
    }
    
    // Salvo p/ o export PDF
    n.recorrenciaTexto = textoRecorrenciaExport; 
    // ===================================================

    const tr = document.createElement("tr");
    tr.className = n.status;

    tr.innerHTML = `
      <td>${escapeHtml(n.origem)}</td>
      <td>${escapeHtml(n.destino)}</td>
      <td>${escapeHtml(n.razao)}</td>
      <td>${escapeHtml(n.nfe)}</td>
      <td>${escapeHtml(n.serie)}</td>
      <td>${n.emissao ? n.emissao.toLocaleDateString("pt-BR") : ""}</td>
      <td><strong>${n.diasDesde}</strong></td>
      <td><span class="badge ${n.status}">${escapeHtml(n.statusIcon)} ${escapeHtml(n.statusLabel)}</span></td>
      <td>${htmlRecorrencia}</td> <td><strong>${formatBRL(n.valor)}</strong></td>
    `;

    tr.addEventListener("click", ()=>{
      detalhes.textContent = JSON.stringify({
        "Origem": n.origem,
        "Destino": n.destino,
        "Razão Social": n.razao,
        "NF-e": n.nfe,
        "Série": n.serie,
        "Chave de Acesso": n.chave,
        "Data de Emissão": n.emissao ? n.emissao.toLocaleDateString("pt-BR") : "",
        "Dias Pendentes": n.diasDesde,
        "Status": n.statusLabel,
        "Análise de Recorrência": textoRecorrenciaExport,
        "Valor": formatBRL(n.valor)
      }, null, 2);
      modal.classList.add("open");
    });

    tabelaBody.appendChild(tr);
  }
}

// ===== Lógica de Renderizar Aba do Histórico ======
function renderHistoricoDatas() {
   const ul = $("ulHistoricoDatas");
   const tbodyDet = $("tabelaDetHist").querySelector("tbody");
   
   ul.innerHTML = "";
   tbodyDet.innerHTML = "";
   $("tituloDetalheHist").textContent = "Selecione uma importação na lista";

   // Pegamos as chaves, mas invertemos para mostrar as importações mais recentes no TOPO
   const chavesDesc = Object.keys(historico).sort((a,b) => b - a);

   if(chavesDesc.length === 0) {
      ul.innerHTML = "<li style='text-align:center;'>Nenhum histórico salvo.</li>";
      return;
   }

   chavesDesc.forEach(ts => {
      const li = document.createElement("li");
      const dateObj = new Date(parseInt(ts));
      const totalNotasNaqueleDia = historico[ts].length;
      
      li.innerHTML = `
        <strong>${dateObj.toLocaleDateString("pt-BR")}</strong> às ${dateObj.toLocaleTimeString("pt-BR")}<br>
        <span style="font-size:12px; opacity:0.8; color:#6b7280;">Continha ${totalNotasNaqueleDia} notas</span>
      `;
      
      li.onclick = () => {
         document.querySelectorAll("#ulHistoricoDatas li").forEach(e => e.classList.remove("ativo"));
         li.classList.add("ativo");
         renderHistoricoDetalhe(ts, dateObj);
      };
      ul.appendChild(li);
   });
}

function renderHistoricoDetalhe(ts, dateObj) {
   $("tituloDetalheHist").textContent = `Registros da importação de: ${dateObj.toLocaleString("pt-BR")}`;
   const tbody = $("tabelaDetHist").querySelector("tbody");
   tbody.innerHTML = "";
   
   const notasDaquelaData = historico[ts];
   if(!notasDaquelaData || notasDaquelaData.length === 0) return;

   notasDaquelaData.forEach(n => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td style="text-align:left;">${escapeHtml(n.origem)}</td>
        <td style="text-align:left;">${escapeHtml(n.destino)}</td>
        <td style="text-align:left;">${escapeHtml(n.nfe)}</td>
        <td style="text-align:right;"><strong>${formatBRL(n.valor)}</strong></td>
      `;
      tbody.appendChild(tr);
   });
}

// ===== Outras Funções do Painel (Gráficos, KPIs, PDF) =====
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

function renderMensagem(){
  const loja = filtroLoja.value;
  const ref = dataRef.value ? new Date(dataRef.value + "T00:00:00") : new Date();
  const pend = notasFiltradas.filter(n => n.status !== "ok");
  const msg = [
    "Bom dia! Espero que esteja bem.",
    `Observei que há pendências de notas fiscais ${loja ? `na unidade ${loja}` : "nas unidades"} (ref.: ${ref.toLocaleDateString("pt-BR")}).`,
    `Vencidas: ${pend.filter(n=>n.status==="vencido").length} | Alerta: ${pend.filter(n=>n.status==="alerta").length}.`,
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
    $("msg").focus(); $("msg").select(); document.execCommand("copy");
    toast("Mensagem copiada ✅");
  }
});

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

  if(chartQtd) chartQtd.destroy();
  chartQtd = new window.Chart($("chartQtd"), {
    type: "doughnut",
    data: { labels:["OK","Alerta","Vencido"], datasets:[{ data:qtd }] },
    options: { responsive:true, plugins:{ legend:{ position:"bottom" } } }
  });

  if(chartValor) chartValor.destroy();
  chartValor = new window.Chart($("chartValor"), {
    type: "bar",
    data: { labels:["OK","Alerta","Vencido"], datasets:[{ label:"Valor (R$)", data: val }] },
    options: { responsive:true, plugins:{ legend:{ display:false } }, scales:{ y:{ ticks:{ callback:(v)=>Number(v).toLocaleString("pt-BR") } } } }
  });
}

btnPng.addEventListener("click", async ()=>{
  const node = $("painel");
  const canvas = await window.html2canvas(node, { scale: 2, useCORS: true });
  const a = document.createElement("a");
  a.href = canvas.toDataURL("image/png");
  a.download = `painel_notas_${todayISO()}.png`;
  a.click();
});

function exportPDF(titulo, rows){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation:"landscape", unit:"pt", format:"a4" });
  doc.setFontSize(14);
  doc.text(titulo, 40, 40);

  const arr = rows.map(n=>[
    n.origem, n.destino, n.nfe,
    n.emissao ? n.emissao.toLocaleDateString("pt-BR") : "",
    String(n.diasDesde), n.statusLabel,
    n.recorrenciaTexto || "Nova", // Incluí no PDF a recorrência também!
    formatBRL(n.valor)
  ]);

  doc.autoTable({
    startY: 80,
    head: [[ "Origem","Destino","NF-e","Emissão","Dias","Status","Recorrência","Valor" ]],
    body: arr,
    styles: { fontSize: 8, cellPadding: 4 },
    headStyles: { fillColor: [17,24,39] },
  });
  doc.save(`${titulo.replace(/\s+/g,'_').toLowerCase()}_${todayISO()}.pdf`);
}

btnPdfLoja.addEventListener("click", ()=>{
  const loja = filtroLoja.value;
  if(!loja) return alert("Selecione uma loja destino.");
  exportPDF(`Pendências - ${loja}`, notasFiltradas);
});

btnPdfTudo.addEventListener("click", ()=> exportPDF("Pendências - Consolidado", notasFiltradas));

let toastTimer=null;
function toast(msg){
  let el = document.getElementById("toast");
  if(!el){
    el = document.createElement("div");
    el.id = "toast";
    el.style.position="fixed"; el.style.bottom="16px"; el.style.right="16px";
    el.style.background="#111827"; el.style.color="#fff"; el.style.padding="10px 12px";
    el.style.borderRadius="12px"; el.style.boxShadow="0 10px 30px rgba(0,0,0,.18)";
    el.style.zIndex=9999;
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.display="block";
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=>{ el.style.display="none"; }, 4000);
}