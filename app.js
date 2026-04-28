import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getFirestore,
  collection,
  doc,
  setDoc,
  deleteDoc,
  onSnapshot,
  serverTimestamp,
  query,
  orderBy
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

// =========================
// FIREBASE
// Mantido com a mesma configuração do projeto anterior.
// =========================
const firebaseConfig = {
  apiKey: "AIzaSyDReYPPhvjjQ4DdLOeQQDg_PrqPCwYaFfU",
  authDomain: "motorista-80298.firebaseapp.com",
  projectId: "motorista-80298",
  storageBucket: "motorista-80298.firebasestorage.app",
  messagingSenderId: "988614619560",
  appId: "1:988614619560:web:f2521ff21aae96aa486d9d",
  measurementId: "G-S1T8661860"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const COLLECTION_NAME = "checklists_altura";
const PAGE_SIZE = 5;

const CHECKLIST_SECTIONS = [
  {
    id: "planejamento",
    titulo: "1. Planejamento e Autorização",
    descricao: "Validação prévia da atividade e ciência dos envolvidos.",
    itens: [
      { chave: "atividadePlanejada", label: "Atividade planejada previamente", alertOn: "nao" },
      { chave: "aprRealizada", label: "Avaliação de risco realizada (APR)", alertOn: "nao" },
      { chave: "climaFavoravel", label: "Condições climáticas favoráveis", alertOn: "nao" },
      { chave: "areaSinalizada", label: "Área isolada/sinalizada", alertOn: "nao" },
      { chave: "clienteCiente", label: "Cliente/encarregado ciente da atividade", alertOn: "nao" }
    ]
  },
  {
    id: "acesso",
    titulo: "2. Escadas / Acesso",
    descricao: "Conferência das condições da escada e do acesso.",
    itens: [
      { chave: "escadaAdequada", label: "Escada adequada ao tipo de atividade", alertOn: "nao" },
      { chave: "escadaBoasCondicoes", label: "Escada em boas condições", alertOn: "nao" },
      { chave: "escadaInspecionada", label: "Escada inspecionada antes do uso", alertOn: "nao" },
      { chave: "anguloCorreto", label: "Ângulo correto", alertOn: "nao" },
      { chave: "escadaEstabilizada", label: "Escada amarrada/estabilizada", alertOn: "nao" },
      { chave: "areaInferiorIsolada", label: "Área inferior isolada", alertOn: "nao" }
    ]
  },
  {
    id: "protecao",
    titulo: "3. Proteção Individual e Coletiva",
    descricao: "Validação dos EPIs, EPCs e organização de ferramentas.",
    itens: [
      { chave: "capacete", label: "Capacete", alertOn: "nao" },
      { chave: "cintoAltura", label: "Cinto para trabalho em altura (quando aplicável)", alertOn: "nao" },
      { chave: "talabarte", label: "Talabarte conectado corretamente", alertOn: "nao" },
      { chave: "calcadoSeguranca", label: "Calçado de segurança", alertOn: "nao" },
      { chave: "oculosProtecao", label: "Óculos de proteção", alertOn: "nao" },
      { chave: "ferramentasPresas", label: "Ferramentas presas ou organizadas", alertOn: "nao" }
    ]
  },
  {
    id: "execucao",
    titulo: "4. Execução da Atividade",
    descricao: "Boas práticas durante a execução da instalação/manutenção.",
    itens: [
      { chave: "tresPontosApoio", label: "Subida e descida com três pontos de apoio", alertOn: "nao" },
      { chave: "semImproviso", label: "Nada improvisado (caixa, cadeira, apoio irregular)", alertOn: "nao" },
      { chave: "posturaSegura", label: "Postura segura durante instalação/manutenção", alertOn: "nao" },
      { chave: "comunicacaoSolo", label: "Comunicação adequada com apoio em solo", alertOn: "nao" }
    ]
  },
  {
    id: "desvios",
    titulo: "5. Desvios e Ações",
    descricao: "Registro de desvios, interrupções e ações corretivas.",
    itens: [
      { chave: "semDesvios", label: "Sem desvios", alertOn: "nao" },
      { chave: "desvioIdentificado", label: "Desvio identificado", alertOn: "sim", exigeDescricaoSeSim: true },
      { chave: "atividadeInterrompida", label: "Atividade interrompida (se necessário)", alertOn: null },
      { chave: "orientacaoRealizada", label: "Orientação imediata realizada", alertOn: null },
      { chave: "acaoRegistrada", label: "Ação registrada / plano de ação aberto", alertOn: null }
    ]
  }
];

const form = document.getElementById("formChecklist");
const checklistContainer = document.getElementById("checklistContainer");
const lista = document.getElementById("lista");
const saveMsg = document.getElementById("saveMsg");
const dataRegistro = document.getElementById("dataRegistro");
const formTitle = document.getElementById("formTitle");
const formSubtitle = document.getElementById("formSubtitle");
const submitBtn = document.getElementById("submitBtn");
const cancelEditBtn = document.getElementById("cancelEditBtn");

const totalRegistrosEl = document.getElementById("totalRegistros");
const totalHojeEl = document.getElementById("totalHoje");
const totalAlertasEl = document.getElementById("totalAlertas");
const totalRiscoAltoEl = document.getElementById("totalRiscoAlto");
const ultimaAtualizacaoEl = document.getElementById("ultimaAtualizacao");

const filtroTecnicoEl = document.getElementById("filtroTecnico");
const filtroClienteEl = document.getElementById("filtroCliente");
const filtroRiscoEl = document.getElementById("filtroRisco");
const filtroDataInicioEl = document.getElementById("filtroDataInicio");
const filtroDataFimEl = document.getElementById("filtroDataFim");
const btnLimparFiltros = document.getElementById("btnLimparFiltros");
const btnExportarExcel = document.getElementById("btnExportarExcel");
const btnVerMais = document.getElementById("btnVerMais");
const btnVerMenos = document.getElementById("btnVerMenos");

const detailsModal = document.getElementById("detailsModal");
const modalBody = document.getElementById("modalBody");
const modalTitle = document.getElementById("modalTitle");
const modalSubtitle = document.getElementById("modalSubtitle");
const closeModalBtn = document.getElementById("closeModalBtn");
const modalEditBtn = document.getElementById("modalEditBtn");
const modalDeleteBtn = document.getElementById("modalDeleteBtn");

const welcomeScreen = document.getElementById("welcomeScreen");
const appShell = document.getElementById("appShell");
const enterSystemBtn = document.getElementById("enterSystemBtn");
const btnToggleTheme = document.getElementById("btnToggleTheme");
const toastEl = document.getElementById("toast");

let editingDocId = null;
let openedDocId = null;
let currentDocsCache = [];
let visibleCount = PAGE_SIZE;

function agoraLocalInput() {
  const d = new Date();
  const offset = d.getTimezoneOffset();
  const local = new Date(d.getTime() - offset * 60000);
  return local.toISOString().slice(0, 16);
}

function escapeHtml(valor) {
  return String(valor ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function pegarValor(id) {
  return document.getElementById(id)?.value?.trim() || "";
}

function obterRadio(name) {
  return document.querySelector(`input[name="${name}"]:checked`)?.value || "";
}

function marcarRadio(name, valor) {
  document.querySelectorAll(`input[name="${name}"]`).forEach((radio) => {
    radio.checked = radio.value === String(valor || "").toLowerCase() || radio.value === String(valor || "");
  });
}

function formatarDataHoraBR(valor) {
  if (!valor) return "--";
  const data = new Date(valor);
  if (Number.isNaN(data.getTime())) return String(valor).replace("T", " ");
  return data.toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" });
}

function formatarTimestamp(timestamp) {
  if (!timestamp) return "--";
  const data = typeof timestamp?.toDate === "function" ? timestamp.toDate() : new Date(timestamp);
  if (Number.isNaN(data.getTime())) return "--";
  return data.toLocaleString("pt-BR", { dateStyle: "short", timeStyle: "short" });
}

function formatarDataArquivo(date = new Date()) {
  const ano = date.getFullYear();
  const mes = String(date.getMonth() + 1).padStart(2, "0");
  const dia = String(date.getDate()).padStart(2, "0");
  const hora = String(date.getHours()).padStart(2, "0");
  const minuto = String(date.getMinutes()).padStart(2, "0");
  return `${ano}-${mes}-${dia}_${hora}-${minuto}`;
}

function mesmaDataLocal(dataA, dataB) {
  const a = new Date(dataA);
  const b = new Date(dataB);
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

function showToast(message, isError = false) {
  if (!toastEl) return;
  toastEl.textContent = message;
  toastEl.classList.remove("hidden");
  toastEl.style.background = isError ? "#991b1b" : "#0f172a";
  clearTimeout(showToast._timer);
  showToast._timer = setTimeout(() => toastEl.classList.add("hidden"), 2600);
}

function setMensagem(msg, erro = false) {
  saveMsg.textContent = msg;
  saveMsg.className = erro ? "message-box error" : "message-box success";
  showToast(msg, erro);
}

function encontrarItemConfig(chave) {
  for (const section of CHECKLIST_SECTIONS) {
    const item = section.itens.find((i) => i.chave === chave);
    if (item) return item;
  }
  return null;
}

function itemTemAlerta(chave, resposta) {
  const config = encontrarItemConfig(chave);
  if (!config || config.alertOn === null || config.alertOn === undefined) return false;
  return String(resposta || "").toLowerCase() === String(config.alertOn).toLowerCase();
}

function registroTemAlerta(dados) {
  const alertaChecklist = Object.entries(dados?.itens || {}).some(([chave, item]) =>
    itemTemAlerta(chave, item?.resposta)
  );
  return alertaChecklist || dados?.classificacaoRisco === "Alto";
}

function riscoAlto(dados) {
  return dados?.classificacaoRisco === "Alto";
}

function renderizarFormularioChecklist() {
  checklistContainer.innerHTML = CHECKLIST_SECTIONS.map((section) => `
    <section class="checklist-section">
      <div class="checklist-header">
        <h3>${escapeHtml(section.titulo)}</h3>
        <p>${escapeHtml(section.descricao)}</p>
      </div>
      <div class="checklist-items">
        ${section.itens.map((item) => criarItemChecklist(item)).join("")}
      </div>
    </section>
  `).join("");

  CHECKLIST_SECTIONS.forEach((section) => {
    section.itens.forEach((item) => {
      document.querySelectorAll(`input[name="${item.chave}_resposta"]`).forEach((radio) => {
        radio.addEventListener("change", () => {
          atualizarVisualItem(item.chave);
          atualizarCampoDesvio(item.chave);
        });
      });
    });
  });
}

function criarItemChecklist(item) {
  const descricaoExtra = item.exigeDescricaoSeSim
    ? `<div class="field full hidden" id="${item.chave}_extra">
        <label for="${item.chave}_descricao">Qual desvio?</label>
        <textarea id="${item.chave}_descricao" placeholder="Descreva o desvio identificado."></textarea>
      </div>`
    : "";

  return `
    <article class="check-item" id="card_${item.chave}">
      <div class="check-item-head">
        <div class="check-item-title">${escapeHtml(item.label)}</div>
        <div class="check-options">
          <label class="option-pill">
            <input type="radio" name="${item.chave}_resposta" value="sim" />
            Sim
          </label>
          <label class="option-pill">
            <input type="radio" name="${item.chave}_resposta" value="nao" />
            Não
          </label>
        </div>
      </div>
      <div class="field full">
        <label for="${item.chave}_obs">Observações</label>
        <textarea id="${item.chave}_obs" placeholder="Descreva algo somente se necessário."></textarea>
      </div>
      ${descricaoExtra}
    </article>
  `;
}

function atualizarVisualItem(chave) {
  const card = document.getElementById(`card_${chave}`);
  const resposta = obterRadio(`${chave}_resposta`);
  if (!card) return;
  card.classList.toggle("alert", itemTemAlerta(chave, resposta));
}

function atualizarCampoDesvio(chave) {
  const config = encontrarItemConfig(chave);
  if (!config?.exigeDescricaoSeSim) return;
  const extra = document.getElementById(`${chave}_extra`);
  if (!extra) return;
  extra.classList.toggle("hidden", obterRadio(`${chave}_resposta`) !== "sim");
}

function montarItensChecklist() {
  const itens = {};
  CHECKLIST_SECTIONS.forEach((section) => {
    section.itens.forEach((item) => {
      itens[item.chave] = {
        label: item.label,
        resposta: obterRadio(`${item.chave}_resposta`),
        observacoes: pegarValor(`${item.chave}_obs`),
        descricaoDesvio: item.exigeDescricaoSeSim ? pegarValor(`${item.chave}_descricao`) : ""
      };
    });
  });
  return itens;
}

function montarDadosFormulario() {
  return {
    dataRegistro: dataRegistro.value,
    localCliente: pegarValor("localCliente"),
    atividade: pegarValor("atividade"),
    alturaAproximada: pegarValor("alturaAproximada"),
    tst: pegarValor("tst"),
    tecnico: pegarValor("tecnico"),
    itens: montarItensChecklist(),
    classificacaoRisco: obterRadio("classificacaoRisco")
  };
}

function validarDados(dados) {
  if (!dados.dataRegistro) return "Informe a data e hora.";
  if (!dados.localCliente) return "Informe o local / cliente.";
  if (!dados.atividade) return "Informe a atividade.";
  if (!dados.alturaAproximada) return "Informe a altura aproximada.";
  if (!dados.tecnico) return "Informe o técnico executante.";

  let semResposta = false;
  CHECKLIST_SECTIONS.forEach((section) => {
    section.itens.forEach((item) => {
      if (!dados.itens?.[item.chave]?.resposta) semResposta = true;
    });
  });
  if (semResposta) return "Responda todos os itens do checklist.";

  const desvio = dados.itens?.desvioIdentificado;
  if (desvio?.resposta === "sim" && !desvio?.descricaoDesvio) {
    return "Descreva qual desvio foi identificado.";
  }

  if (!dados.classificacaoRisco) return "Selecione a classificação do risco.";
  return "";
}

async function salvarChecklist(event) {
  event.preventDefault();
  const dados = montarDadosFormulario();
  const erro = validarDados(dados);
  if (erro) {
    setMensagem(erro, true);
    return;
  }

  submitBtn.disabled = true;

  try {
    const ref = editingDocId
      ? doc(db, COLLECTION_NAME, editingDocId)
      : doc(collection(db, COLLECTION_NAME));

    await setDoc(ref, {
      ...dados,
      atualizadoEm: serverTimestamp(),
      ...(editingDocId ? {} : { criadoEm: serverTimestamp() })
    }, { merge: true });

    setMensagem(editingDocId ? "Checklist atualizado com sucesso." : "Checklist salvo com sucesso.");
    limparFormulario();
  } catch (error) {
    console.error(error);
    setMensagem("Erro ao salvar checklist. Verifique o Firebase e as regras do Firestore.", true);
  } finally {
    submitBtn.disabled = false;
  }
}

function preencherFormulario(dados) {
  dataRegistro.value = dados?.dataRegistro || agoraLocalInput();
  document.getElementById("localCliente").value = dados?.localCliente || "";
  document.getElementById("atividade").value = dados?.atividade || "";
  document.getElementById("alturaAproximada").value = dados?.alturaAproximada || "";
  document.getElementById("tst").value = dados?.tst || "";
  document.getElementById("tecnico").value = dados?.tecnico || "";
  marcarRadio("classificacaoRisco", dados?.classificacaoRisco || "");

  CHECKLIST_SECTIONS.forEach((section) => {
    section.itens.forEach((item) => {
      const itemData = dados?.itens?.[item.chave] || {};
      marcarRadio(`${item.chave}_resposta`, itemData.resposta || "");
      const obs = document.getElementById(`${item.chave}_obs`);
      if (obs) obs.value = itemData.observacoes || "";
      const desc = document.getElementById(`${item.chave}_descricao`);
      if (desc) desc.value = itemData.descricaoDesvio || "";
      atualizarVisualItem(item.chave);
      atualizarCampoDesvio(item.chave);
    });
  });
}

function limparFormulario() {
  form.reset();
  dataRegistro.value = agoraLocalInput();
  editingDocId = null;
  atualizarModoFormulario();

  CHECKLIST_SECTIONS.forEach((section) => {
    section.itens.forEach((item) => {
      atualizarVisualItem(item.chave);
      atualizarCampoDesvio(item.chave);
    });
  });
}

function atualizarModoFormulario() {
  const emEdicao = !!editingDocId;
  submitBtn.textContent = emEdicao ? "Atualizar checklist" : "Salvar checklist";
  cancelEditBtn.hidden = !emEdicao;
  formTitle.textContent = emEdicao ? "Editar checklist" : "Novo checklist";
  formSubtitle.textContent = emEdicao
    ? "Altere os dados do registro selecionado."
    : "Preencha os dados da atividade e marque os itens de segurança.";
}

function getRegistrosFiltrados() {
  const tecnico = filtroTecnicoEl.value.trim().toLowerCase();
  const cliente = filtroClienteEl.value.trim().toLowerCase();
  const risco = filtroRiscoEl.value;
  const dataInicio = filtroDataInicioEl.value;
  const dataFim = filtroDataFimEl.value;

  return currentDocsCache.filter((item) => {
    const tecnicoOk = !tecnico || String(item.tecnico || "").toLowerCase().includes(tecnico);
    const clienteOk = !cliente || String(item.localCliente || "").toLowerCase().includes(cliente);
    const riscoOk = !risco || item.classificacaoRisco === risco;

    const dataValida = item?.dataRegistro && !Number.isNaN(new Date(item.dataRegistro).getTime());
    const inicioOk = !dataInicio || (dataValida && item.dataRegistro.slice(0, 10) >= dataInicio);
    const fimOk = !dataFim || (dataValida && item.dataRegistro.slice(0, 10) <= dataFim);

    return tecnicoOk && clienteOk && riscoOk && inicioOk && fimOk;
  });
}

function montarCard(dados) {
  const alerta = registroTemAlerta(dados);
  const alto = riscoAlto(dados);

  let badge = `<span class="badge badge-ok">Sem alerta</span>`;
  if (alto) badge = `<span class="badge badge-risk">Risco alto</span>`;
  else if (alerta) badge = `<span class="badge badge-alert">Com alerta</span>`;

  return `
    <article class="status-card ${alto ? "high-risk-card" : alerta ? "alert-card" : ""}" data-open-id="${escapeHtml(dados.__docId)}">
      <div class="status-card-row">
        <div>
          <h3>${escapeHtml(dados.localCliente || "--")}</h3>
          <p>${escapeHtml(dados.atividade || "--")} • ${escapeHtml(dados.tecnico || "--")} • ${formatarDataHoraBR(dados.dataRegistro)}</p>
        </div>
        <div class="status-card-row">
          <span class="badge ${alto ? "badge-risk" : dados.classificacaoRisco === "Médio" ? "badge-alert" : "badge-ok"}">${escapeHtml(dados.classificacaoRisco || "--")}</span>
          ${badge}
        </div>
      </div>
    </article>
  `;
}

function renderizarLista() {
  const filtrados = getRegistrosFiltrados();
  const exibidos = filtrados.slice(0, visibleCount);
  lista.innerHTML = "";

  if (!filtrados.length) {
    lista.innerHTML = `<li class="empty-state"><strong>Nenhum checklist encontrado.</strong><br><span>Cadastre um novo registro ou ajuste os filtros.</span></li>`;
    btnVerMais.classList.add("hidden");
    btnVerMenos.classList.add("hidden");
    return;
  }

  exibidos.forEach((dados) => {
    const li = document.createElement("li");
    li.innerHTML = montarCard(dados);
    lista.appendChild(li);
  });

  btnVerMais.classList.toggle("hidden", filtrados.length <= visibleCount);
  btnVerMenos.classList.toggle("hidden", visibleCount <= PAGE_SIZE);
}

function atualizarResumo(registros) {
  const hoje = new Date();
  totalRegistrosEl.textContent = String(registros.length);
  totalHojeEl.textContent = String(registros.filter((item) => item?.dataRegistro && mesmaDataLocal(item.dataRegistro, hoje)).length);
  totalAlertasEl.textContent = String(registros.filter((item) => registroTemAlerta(item)).length);
  totalRiscoAltoEl.textContent = String(registros.filter((item) => riscoAlto(item)).length);

  const primeiro = registros[0];
  ultimaAtualizacaoEl.textContent = primeiro ? formatarTimestamp(primeiro.atualizadoEm || primeiro.criadoEm) : "--";
}

function reaplicarRenderizacao() {
  renderizarLista();
  atualizarResumo(currentDocsCache);
}

function abrirDetalhes(docId) {
  const dados = currentDocsCache.find((item) => item.__docId === docId);
  if (!dados) return;

  openedDocId = docId;
  modalTitle.textContent = dados.localCliente || "Detalhes do checklist";
  modalSubtitle.textContent = `${dados.atividade || "--"} • ${formatarDataHoraBR(dados.dataRegistro)}`;

  modalBody.innerHTML = `
    <div class="detail-grid">
      ${detailBox("Data e hora", formatarDataHoraBR(dados.dataRegistro))}
      ${detailBox("Local / Cliente", dados.localCliente)}
      ${detailBox("Atividade", dados.atividade)}
      ${detailBox("Altura aproximada", dados.alturaAproximada)}
      ${detailBox("TST", dados.tst || "--")}
      ${detailBox("Técnico executante", dados.tecnico)}
      ${detailBox("Classificação do risco", dados.classificacaoRisco)}
      ${detailBox("Última atualização", formatarTimestamp(dados.atualizadoEm || dados.criadoEm))}
    </div>
    ${CHECKLIST_SECTIONS.map((section) => montarGrupoDetalhe(section, dados)).join("")}
  `;

  detailsModal.classList.remove("hidden");
}

function detailBox(label, value) {
  return `<div class="detail-box"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value || "--")}</strong></div>`;
}

function montarGrupoDetalhe(section, dados) {
  return `
    <div class="detail-group">
      <h4>${escapeHtml(section.titulo)}</h4>
      ${section.itens.map((item) => montarItemDetalhe(item, dados?.itens?.[item.chave] || {})).join("")}
    </div>
  `;
}

function montarItemDetalhe(itemConfig, itemData) {
  const resposta = itemData?.resposta || "";
  const alerta = itemTemAlerta(itemConfig.chave, resposta);
  const badgeClass = alerta ? "badge-alert" : "badge-ok";

  return `
    <div class="detail-item">
      <div>
        <div class="detail-item-title">${escapeHtml(itemConfig.label)}</div>
        ${itemData?.observacoes ? `<div class="detail-item-obs"><strong>Obs.:</strong> ${escapeHtml(itemData.observacoes)}</div>` : ""}
        ${itemData?.descricaoDesvio ? `<div class="detail-item-obs"><strong>Desvio:</strong> ${escapeHtml(itemData.descricaoDesvio)}</div>` : ""}
      </div>
      <span class="badge ${badgeClass}">${resposta ? escapeHtml(resposta.toUpperCase()) : "--"}</span>
    </div>
  `;
}

function fecharModal() {
  detailsModal.classList.add("hidden");
  openedDocId = null;
}

function editarRegistro(docId) {
  const dados = currentDocsCache.find((item) => item.__docId === docId);
  if (!dados) return;
  editingDocId = docId;
  preencherFormulario(dados);
  atualizarModoFormulario();
  fecharModal();
  window.scrollTo({ top: 0, behavior: "smooth" });
  showToast("Modo edição ativado.");
}

async function excluirRegistro(docId) {
  if (!docId) return;
  const confirmar = confirm("Tem certeza que deseja excluir este checklist?");
  if (!confirmar) return;

  try {
    await deleteDoc(doc(db, COLLECTION_NAME, docId));
    fecharModal();
    setMensagem("Checklist excluído com sucesso.");
  } catch (error) {
    console.error(error);
    setMensagem("Erro ao excluir checklist.", true);
  }
}

function gerarLinhasExcel(registros) {
  return registros.map((dados) => {
    const linha = {
      "Data e Hora": formatarDataHoraBR(dados.dataRegistro),
      "Data ISO": dados.dataRegistro || "",
      "Local / Cliente": dados.localCliente || "",
      "Atividade": dados.atividade || "",
      "Altura Aproximada": dados.alturaAproximada || "",
      "TST": dados.tst || "",
      "Técnico Executante": dados.tecnico || "",
      "Classificação do Risco": dados.classificacaoRisco || "",
      "Com Alerta": registroTemAlerta(dados) ? "Sim" : "Não",
      "Última Atualização": formatarTimestamp(dados.atualizadoEm || dados.criadoEm)
    };

    CHECKLIST_SECTIONS.forEach((section) => {
      section.itens.forEach((item) => {
        const itemData = dados?.itens?.[item.chave] || {};
        linha[item.label] = itemData.resposta || "";
        linha[`Obs - ${item.label}`] = itemData.observacoes || "";
        if (item.exigeDescricaoSeSim) {
          linha["Qual desvio?"] = itemData.descricaoDesvio || "";
        }
      });
    });

    return linha;
  });
}

function exportarExcel() {
  const filtrados = getRegistrosFiltrados();

  if (!filtrados.length) {
    setMensagem("Não há registros para exportar com os filtros atuais.", true);
    return;
  }

  const linhas = gerarLinhasExcel(filtrados);
  const ws = XLSX.utils.json_to_sheet(linhas);
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, ws, "Checklists Altura");
  XLSX.writeFile(wb, `checklists_altura_${formatarDataArquivo()}.xlsx`);
}

function aplicarTemaSalvo() {
  const tema = localStorage.getItem("altura_safe_theme");
  const dark = tema === "dark";
  document.body.classList.toggle("dark", dark);
  btnToggleTheme.textContent = dark ? "☀️" : "🌙";
}

function alternarTema() {
  const dark = document.body.classList.toggle("dark");
  localStorage.setItem("altura_safe_theme", dark ? "dark" : "light");
  btnToggleTheme.textContent = dark ? "☀️" : "🌙";
}

function abrirSistema() {
  welcomeScreen.classList.add("hidden");
  appShell.classList.remove("hidden");
  showToast("Sistema carregado com sucesso.");
}

function iniciarRealtime() {
  const q = query(collection(db, COLLECTION_NAME), orderBy("dataRegistro", "desc"));

  onSnapshot(q, (snapshot) => {
    currentDocsCache = snapshot.docs.map((docSnap) => ({
      __docId: docSnap.id,
      ...docSnap.data()
    }));

    reaplicarRenderizacao();
  }, (error) => {
    console.error(error);
    setMensagem("Erro ao carregar registros em tempo real. Verifique as regras do Firestore.", true);
  });
}

function registrarEventos() {
  form.addEventListener("submit", salvarChecklist);
  cancelEditBtn.addEventListener("click", limparFormulario);

  [filtroTecnicoEl, filtroClienteEl, filtroRiscoEl, filtroDataInicioEl, filtroDataFimEl].forEach((el) => {
    el.addEventListener("input", () => {
      visibleCount = PAGE_SIZE;
      reaplicarRenderizacao();
    });
  });

  btnLimparFiltros.addEventListener("click", () => {
    filtroTecnicoEl.value = "";
    filtroClienteEl.value = "";
    filtroRiscoEl.value = "";
    filtroDataInicioEl.value = "";
    filtroDataFimEl.value = "";
    visibleCount = PAGE_SIZE;
    reaplicarRenderizacao();
  });

  btnVerMais.addEventListener("click", () => {
    visibleCount += PAGE_SIZE;
    renderizarLista();
  });

  btnVerMenos.addEventListener("click", () => {
    visibleCount = Math.max(PAGE_SIZE, visibleCount - PAGE_SIZE);
    renderizarLista();
  });

  btnExportarExcel.addEventListener("click", exportarExcel);

  lista.addEventListener("click", (event) => {
    const card = event.target.closest("[data-open-id]");
    if (card) abrirDetalhes(card.dataset.openId);
  });

  closeModalBtn.addEventListener("click", fecharModal);
  detailsModal.addEventListener("click", (event) => {
    if (event.target === detailsModal) fecharModal();
  });

  modalEditBtn.addEventListener("click", () => editarRegistro(openedDocId));
  modalDeleteBtn.addEventListener("click", () => excluirRegistro(openedDocId));
  enterSystemBtn.addEventListener("click", abrirSistema);
  btnToggleTheme.addEventListener("click", alternarTema);
}

renderizarFormularioChecklist();
dataRegistro.value = agoraLocalInput();
aplicarTemaSalvo();
registrarEventos();
iniciarRealtime();
