import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy, setPersistence, browserSessionPersistence, deleteDoc, writeBatch } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

// === LISTA FIXA APENAS PARA CADASTRO ===
const UNIDADES_CBMMA_FIXAS = [
    "1 BBM", "2 BBM", "1 CIEBM", "10 BBM", "13 BBM", "16 CIBM", "BBS", "BBA", "BMUS", "CGCS", "DEP", "DAT", "DP", "DF", "DPM", "DAL", "CPP", "CPO", 
    "1 Seção", "2 Seção", "3 Seção", "4 Seção", "CAPS", "CRF", "CEPDECMA", "ASPIRANTES", "CADETES", "DER", 
    "CMCB I", "CMCB II - SJR", "CMCB XII - PAÇO", "CMCB XIII - GUANABARA", "CMCB XXVI - PIO XII", 
    "ABMJM", "GAB CMT GERAL", "GAB CMT ADJUNTO"
];

const FUNCOES_TATICAS = [
    "SOCORRISTA", "MOTORISTA (AR/ABT)", "COORD. DO EVENTO", "SUBCOORD. DO EVENTO", 
    "OF. DE LOGÍSTICA", "AUX. DE LOGÍSTICA", "OF DE OPERAÇÕES", 
    "OF. DE ADM/FINANÇAS", "AUX. DE ADM/FINANÇAS", "OPER. DRONE", "OF. DE INFORMAÇÃO", 
    "MOP 1 (MOTO)", "MOP 2 (MOTO)", "AR (MOTORISTA)", "ABT", "VAN", "MICRO - ÔNIBUS"
];

let usuarioAtual = null;
let perfilAtual = null;
let escalaSelecionadaId = null;
let eventoPreviewAtual = null; 
let listaOrdensTemporaria = [];
let dadosParaEnvio = null;
let idEdicaoAdmin = null; 

// === INICIALIZAÇÃO ===
document.addEventListener('DOMContentLoaded', () => {
    popularSelectCadastroEFuncoes();
});

function popularSelectCadastroEFuncoes() {
    // 1. Popula APENAS o select de CADASTRO com a lista fixa
    const selCadastro = document.getElementById('unidade-cadastro');
    if(selCadastro && selCadastro.options.length <= 1) {
        selCadastro.innerHTML = "<option value=''>Selecione a Unidade...</option>";
        UNIDADES_CBMMA_FIXAS.forEach(u => selCadastro.innerHTML += `<option value="${u}">${u}</option>`);
    }

    // 2. Popula Funções
    const selectsFuncao = [document.getElementById('select-funcao'), document.getElementById('edit-admin-funcao')];
    selectsFuncao.forEach(sel => {
        if(sel) {
            sel.innerHTML = "";
            FUNCOES_TATICAS.forEach(f => sel.innerHTML += `<option value="${f}">${f}</option>`);
        }
    });
}

// Função Nova: Carrega no Admin APENAS unidades que existem no banco
async function carregarUnidadesCadastradasNoAdmin() {
    const selAdmin = document.getElementById('select-unidade');
    if(!selAdmin) return;
    
    selAdmin.innerHTML = "<option value=''>Carregando...</option>";
    
    try {
        // Busca usuários que são "escalante"
        const q = query(collection(db, "usuarios"), where("funcao", "==", "escalante"));
        const snapshot = await getDocs(q);
        
        const unidadesReais = [];
        snapshot.forEach(doc => {
            const data = doc.data();
            if(data.unidade) unidadesReais.push(data.unidade);
        });

        // Remove duplicatas e ordena
        const unidadesUnicas = [...new Set(unidadesReais)].sort();

        selAdmin.innerHTML = "<option value=''>Selecione a Unidade...</option>";
        if(unidadesUnicas.length === 0) {
            selAdmin.innerHTML += "<option disabled>Nenhuma unidade cadastrada</option>";
        } else {
            unidadesUnicas.forEach(u => selAdmin.innerHTML += `<option value="${u}">${u}</option>`);
        }

    } catch (e) {
        console.error("Erro ao carregar unidades:", e);
        selAdmin.innerHTML = "<option value=''>Erro ao carregar</option>";
    }
}

// ================= AUTH =================
export async function fazerLogin() {
    const email = document.getElementById('email-login').value;
    const senha = document.getElementById('senha-login').value;
    const btn = document.querySelector('button[onclick="fazerLogin()"]');
    const textoOriginal = btn.innerHTML;
    
    if(!email || !senha) return alert("Preencha email e senha");

    btn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Acessando...';
    btn.disabled = true;

    try { 
        await setPersistence(auth, browserSessionPersistence);
        await signInWithEmailAndPassword(auth, email, senha); 
    } 
    catch (e) { 
        console.error(e); 
        document.getElementById('msg-erro').innerText = "Credenciais inválidas.";
        btn.innerHTML = textoOriginal; btn.disabled = false;
    }
}

export async function fazerCadastro() {
    const email = document.getElementById('email-cadastro').value;
    const senha = document.getElementById('senha-cadastro').value;
    const unidade = document.getElementById('unidade-cadastro').value;
    
    if(!email || !senha || !unidade) return alert("Preencha todos os campos.");
    
    try {
        const cred = await createUserWithEmailAndPassword(auth, email, senha);
        await setDoc(doc(db, "usuarios", cred.user.uid), { email, unidade: unidade.toUpperCase(), funcao: "escalante" });
        alert("Unidade cadastrada com sucesso!"); window.location.reload();
    } catch (e) { alert("Erro: " + e.message); }
}

export function sair() { signOut(auth).then(() => location.reload()); }

onAuthStateChanged(auth, async (user) => {
    if (user) {
        usuarioAtual = user;
        const snap = await getDoc(doc(db, "usuarios", user.uid));
        if (snap.exists()) {
            perfilAtual = snap.data();
            
            const loginArea = document.getElementById('login-area-wrapper');
            loginArea.style.opacity = '0';
            
            setTimeout(() => {
                loginArea.style.display = 'none';
                const dash = document.getElementById('dashboard-screen');
                dash.style.display = 'block';
                setTimeout(() => dash.classList.add('visible'), 50);
                
                document.getElementById('titulo-unidade').innerText = perfilAtual.unidade;
                popularSelectCadastroEFuncoes();

                if (perfilAtual.funcao === 'admin') {
                    document.getElementById('admin-area').style.display = 'block';
                    carregarEventosAdmin();
                    carregarUnidadesCadastradasNoAdmin(); // <--- Carrega apenas unidades reais
                } else {
                    document.getElementById('unidade-area').style.display = 'block';
                    carregarPendenciasUnidade();
                }
            }, 300);
        }
    }
});

// ================= ADMIN =================
export function adicionarOrdem() {
    const unidade = document.getElementById('select-unidade').value;
    const funcao = document.getElementById('select-funcao').value;
    const oficiais = document.getElementById('input-oficiais').value || 0;
    const pracas = document.getElementById('input-pracas').value || 0;

    if (!unidade) return alert("Selecione uma unidade!");
    if (oficiais == 0 && pracas == 0) return alert("Defina a quantidade.");

    listaOrdensTemporaria.push({ id: Date.now(), unidade, funcao, oficiais, pracas });
    atualizarTabelaOrdens();
}

function atualizarTabelaOrdens() {
    const corpo = document.getElementById('tabela-ordens-body');
    document.getElementById('contador-ordens').innerText = `${listaOrdensTemporaria.length}`;
    corpo.innerHTML = "";
    listaOrdensTemporaria.forEach((item, index) => {
        corpo.innerHTML += `
            <tr class="border-bottom">
                <td class="fw-bold">${item.unidade}</td>
                <td><span class="badge bg-light text-dark border">${item.funcao}</span></td>
                <td class="small fw-bold text-muted">${item.oficiais} OF / ${item.pracas} PÇ</td>
                <td class="text-end"><button onclick="window.app.excluirOrdem(${index})" class="btn btn-sm text-danger ios-click"><i class="bi bi-x-circle-fill"></i></button></td>
            </tr>`;
    });
}

export function excluirOrdem(index) { listaOrdensTemporaria.splice(index, 1); atualizarTabelaOrdens(); }
export function limparOrdens() { listaOrdensTemporaria = []; atualizarTabelaOrdens(); }

export async function dispararSolicitacao() {
    const evento = document.getElementById('nome-evento').value.trim();
    const data = document.getElementById('data-evento').value;
    const horaInicio = document.getElementById('hora-inicio').value;
    const horaFim = document.getElementById('hora-fim').value;
    const prazoData = document.getElementById('prazo-data').value;
    const prazoHora = document.getElementById('prazo-hora').value;

    if (!evento || !data) return alert("Preencha Nome e Data.");
    if (!prazoData) return alert("Defina o Prazo.");
    if (listaOrdensTemporaria.length === 0) return alert("Adicione unidades.");

    try {
        const promises = listaOrdensTemporaria.map(ordem => {
            return addDoc(collection(db, "escalas"), {
                evento: evento.toUpperCase(), data, horaInicio, horaFim,
                prazoData, prazoHora: prazoHora || "23:59",
                unidade: ordem.unidade, funcao: ordem.funcao,
                cota: { oficial: ordem.oficiais, praca: ordem.pracas },
                status: "Pendente", militares: "[]", criadoEm: new Date()
            });
        });
        await Promise.all(promises);
        alert(`Sucesso! Envios realizados.`);
        limparOrdens(); carregarEventosAdmin();
    } catch (e) { alert("Erro: " + e.message); }
}

async function carregarEventosAdmin() {
    const lista = document.getElementById('lista-eventos-admin');
    lista.innerHTML = "<div class='text-center py-3'><span class='spinner-border text-danger'></span></div>";
    try {
        const q = query(collection(db, "escalas")); 
        const snapshot = await getDocs(q);
        const grupos = new Map();

        snapshot.forEach(doc => {
            const d = doc.data();
            if(!d.evento || !d.data) return;
            const chave = `${d.evento}|${d.data}`;
            if (!grupos.has(chave)) grupos.set(chave, { evento: d.evento, data: d.data, total: 0, respondidos: 0 });
            const g = grupos.get(chave);
            g.total++;
            if (d.status === "Preenchido") g.respondidos++;
        });

        lista.innerHTML = "";
        if (grupos.size === 0) lista.innerHTML = "<div class='text-muted text-center py-3'>Histórico vazio.</div>";

        const gruposArray = Array.from(grupos.values()).sort((a, b) => new Date(b.data) - new Date(a.data));

        gruposArray.forEach(info => {
            const dataBr = new Date(info.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const percentual = info.total === 0 ? 0 : Math.round((info.respondidos / info.total) * 100);
            
            lista.innerHTML += `
                <div class="list-group-item p-3 border-bottom ios-click" style="cursor:pointer;" onclick="window.app.abrirPreview('${info.evento}', '${info.data}')">
                    <div class="d-flex justify-content-between align-items-center mb-2">
                        <div><strong class="text-dark d-block text-uppercase">${info.evento}</strong><small class="text-muted fw-bold">${dataBr}</small></div>
                        <i class="bi bi-chevron-right text-muted"></i>
                    </div>
                    <div class="d-flex justify-content-between small text-muted align-items-center mb-1"><span>Progresso</span><span>${info.respondidos}/${info.total}</span></div>
                    <div class="progress" style="height: 6px; border-radius: 10px;"><div class="progress-bar bg-success" style="width: ${percentual}%; border-radius: 10px;"></div></div>
                </div>`;
        });
    } catch(e) { console.error(e); }
}

// ================= ADMIN PREVIEW & ACTIONS =================
export async function abrirPreview(nomeEvento, dataEvento) {
    eventoPreviewAtual = { nome: nomeEvento, data: dataEvento };
    const modal = document.getElementById('preview-modal');
    modal.classList.remove('d-none');
    
    document.getElementById('preview-titulo').innerText = nomeEvento;
    const corpo = document.getElementById('tabela-preview-corpo');
    corpo.innerHTML = "<tr><td colspan='6' class='text-center py-4'><span class='spinner-border text-danger'></span></td></tr>";

    try {
        const q = query(collection(db, "escalas"), where("evento", "==", nomeEvento), where("data", "==", dataEvento));
        const snapshot = await getDocs(q);
        let html = "";
        let ordemGlobal = 1;
        
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            const idDoc = docSnap.id;
            let militares = [];
            try { militares = JSON.parse(d.militares); } catch(e) { militares = []; }
            const cota = d.cota || {oficial: 0, praca: 0};

            const btnEdit = `<button onclick="window.app.editarSolicitacaoAdmin('${idDoc}', '${d.unidade}', '${d.funcao}', ${cota.oficial}, ${cota.praca})" class="btn btn-sm btn-outline-primary border-0 me-1" title="Editar"><i class="bi bi-pencil-square"></i></button>`;
            const btnDelete = `<button onclick="window.app.excluirEscalaIndividual('${idDoc}', '${d.unidade}')" class="btn btn-sm btn-outline-danger border-0" title="Excluir"><i class="bi bi-trash-fill"></i></button>`;

            if(d.status === "Pendente") {
                html += `
                <tr class="table-danger border-bottom">
                    <td class="text-center fw-bold text-muted">-</td>
                    <td colspan="3" class="small text-danger fw-bold align-middle">
                        <i class="bi bi-exclamation-circle-fill me-1"></i> PENDENTE: ${d.unidade}
                        <br><span class="text-muted fw-normal ms-3">Cota: ${cota.oficial} OF / ${cota.praca} PÇ</span>
                    </td>
                    <td class="align-middle">${d.funcao}</td>
                    <td class="text-end align-middle">${btnEdit}${btnDelete}</td>
                </tr>`;
            } else {
                militares.forEach((m, index) => {
                    html += `<tr>
                        <td class="fw-bold text-center text-muted">${ordemGlobal++}</td>
                        <td><span class="fw-bold">${m.posto}</span> ${m.guerra}</td>
                        <td class="small text-muted">${m.contato}</td>
                        <td class="fw-bold text-dark">${d.unidade}</td>
                        <td><span class="badge bg-light text-dark border">${d.funcao}</span></td>
                        <td class="text-end">${index === 0 ? btnEdit + btnDelete : ''}</td>
                    </tr>`;
                });
            }
        });
        corpo.innerHTML = html;
    } catch(e) { console.error(e); corpo.innerHTML = "<tr><td colspan='6'>Erro ao carregar.</td></tr>"; }
}

export function editarSolicitacaoAdmin(id, unidade, funcao, of, pc) {
    idEdicaoAdmin = id;
    document.getElementById('edit-admin-subtitle').innerText = `Editando: ${unidade}`;
    document.getElementById('edit-admin-funcao').value = funcao;
    document.getElementById('edit-admin-of').value = of;
    document.getElementById('edit-admin-pc').value = pc;
    document.getElementById('modal-editar-admin').classList.remove('d-none');
}

export async function salvarEdicaoAdmin() {
    if(!idEdicaoAdmin) return;
    const novaFuncao = document.getElementById('edit-admin-funcao').value;
    const novoOf = document.getElementById('edit-admin-of').value;
    const novoPc = document.getElementById('edit-admin-pc').value;

    try {
        await updateDoc(doc(db, "escalas", idEdicaoAdmin), {
            funcao: novaFuncao,
            cota: { oficial: novoOf, praca: novoPc }
        });
        alert("Atualizado com sucesso!");
        document.getElementById('modal-editar-admin').classList.add('d-none');
        abrirPreview(eventoPreviewAtual.nome, eventoPreviewAtual.data);
    } catch(e) { alert("Erro ao salvar: " + e.message); }
}

export async function excluirEscalaIndividual(idDoc, nomeUnidade) {
    if(!confirm(`Apagar solicitação da unidade ${nomeUnidade}?`)) return;
    try {
        await deleteDoc(doc(db, "escalas", idDoc));
        abrirPreview(eventoPreviewAtual.nome, eventoPreviewAtual.data);
    } catch(e) { alert("Erro: " + e.message); }
}

export async function excluirEventoCompleto() {
    if(!eventoPreviewAtual) return;
    const confirmar = prompt(`ATENÇÃO: Isso apagará TODO o histórico do evento "${eventoPreviewAtual.nome}".\nDigite "APAGAR" para confirmar:`);
    if(confirmar !== "APAGAR") return;

    try {
        const q = query(collection(db, "escalas"), where("evento", "==", eventoPreviewAtual.nome), where("data", "==", eventoPreviewAtual.data));
        const snapshot = await getDocs(q);
        const batch = writeBatch(db);
        snapshot.forEach(d => batch.delete(d.ref));
        await batch.commit();
        
        alert("Evento apagado.");
        document.getElementById('preview-modal').classList.add('d-none');
        carregarEventosAdmin();
    } catch(e) { alert("Erro ao excluir: " + e.message); }
}

// ================= ESCALANTE =================
async function carregarPendenciasUnidade() {
    const lista = document.getElementById('lista-unidade');
    lista.innerHTML = "<div class='text-center w-100 py-5'><span class='spinner-border text-danger'></span><br>Sincronizando...</div>";
    try {
        const q = query(collection(db, "escalas"), where("unidade", "==", perfilAtual.unidade));
        const snapshot = await getDocs(q);
        lista.innerHTML = "";
        
        if (snapshot.empty) return lista.innerHTML = "<div class='text-muted text-center w-100 mt-4 fs-5'>Nenhuma missão pendente.</div>";

        const docs = [];
        snapshot.forEach(d => docs.push({id: d.id, ...d.data()}));
        docs.sort((a,b) => new Date(a.data) - new Date(b.data));

        docs.forEach(d => {
            const cota = d.cota || { oficial: 0, praca: 0 }; 
            const isPendente = d.status === "Pendente";
            let isBloqueado = false;
            let textoPrazo = "";
            
            if (d.prazoData) {
                const dataLimite = new Date(`${d.prazoData}T${d.prazoHora || '23:59'}:00`);
                if (new Date() > dataLimite) { isBloqueado = true; textoPrazo = `<div class="text-danger fw-bold small mt-2"><i class="bi bi-lock-fill"></i> ENCERRADO</div>`; }
                else { textoPrazo = `<div class="text-dark small mt-2 bg-warning bg-opacity-25 p-1 rounded"><i class="bi bi-clock-history"></i> Prazo: ${new Date(d.prazoData).toLocaleDateString('pt-BR')} ${d.prazoHora}</div>`; }
            }

            const btnClass = isBloqueado ? "btn-secondary disabled" : (isPendente ? "btn-tactical" : "btn-outline-success");
            const btnText = isBloqueado ? "EXPIRADO" : (isPendente ? "RESPONDER AGORA" : "EDITAR ENVIO");

            lista.innerHTML += `
                <div class="col-md-6 col-lg-4 animate-up">
                    <div class="bg-white p-4 h-100 rounded-4 shadow-sm border border-light d-flex flex-column position-relative mission-card">
                        <div class="d-flex justify-content-between mb-2">
                            <span class="badge bg-dark">${new Date(d.data).toLocaleDateString('pt-BR', {timeZone:'UTC'})}</span>
                            <span class="badge ${isPendente ? 'bg-warning text-dark' : 'bg-success'}">${d.status}</span>
                        </div>
                        <h5 class="fw-bold mb-0 text-dark text-uppercase">${d.evento}</h5>
                        <small class="text-muted mb-2 d-block">${d.horaInicio} às ${d.horaFim}</small>
                        <div class="bg-light p-3 rounded border text-center my-2">
                            <strong class="d-block text-primary">${d.funcao}</strong>
                            <div class="small text-muted">Cota: ${cota.oficial} Of / ${cota.praca} Pç</div>
                        </div>
                        ${textoPrazo}
                        <button onclick="window.app.abrirEdicao('${d.id}')" class="btn ${btnClass} w-100 fw-bold mt-auto py-3 rounded-3 shadow-sm ios-click" ${isBloqueado ? 'disabled' : ''}>${btnText}</button>
                    </div>
                </div>`;
        });
    } catch(e) { console.error(e); }
}

export async function abrirEdicao(id) {
    escalaSelecionadaId = id;
    const docSnap = await getDoc(doc(db, "escalas", id));
    const d = docSnap.data();
    const cota = d.cota || { oficial: 0, praca: 0 };
    
    document.getElementById('titulo-evento-form').innerText = d.evento;
    document.getElementById('subtitulo-form').innerText = `${d.funcao} | Cota: ${cota.oficial} Of / ${cota.praca} Pç`;
    
    const container = document.getElementById('container-inputs-militares');
    container.innerHTML = "";

    let dadosSalvos = [];
    try { dadosSalvos = JSON.parse(d.militares); } catch {}

    let contador = 0;
    const qtdOf = parseInt(cota.oficial) || 0;
    const qtdPc = parseInt(cota.praca) || 0;

    for(let i=0; i < qtdOf; i++) container.innerHTML += gerarHtmlMilitar(i, 'OFICIAL', dadosSalvos[contador++] || {});
    for(let i=0; i < qtdPc; i++) container.innerHTML += gerarHtmlMilitar(i, 'PRAÇA', dadosSalvos[contador++] || {});

    document.getElementById('form-militar-modal').classList.remove('d-none');
}

function gerarHtmlMilitar(index, tipo, dados) {
    return `
    <div class="p-3 bg-white rounded-3 border mb-3 militar-row shadow-sm">
        <span class="badge bg-secondary mb-2">${tipo} ${index + 1}</span>
        <div class="row g-2">
            <div class="col-4 col-md-3"><input type="text" class="form-control campo-posto fw-bold" placeholder="Posto" value="${dados.posto || ''}"></div>
            <div class="col-8 col-md-5"><input type="text" class="form-control campo-nome" placeholder="Nome Completo" value="${dados.nome || ''}"></div>
            <div class="col-6 col-md-4"><input type="text" class="form-control campo-guerra fw-bold text-uppercase" placeholder="Nome Guerra" value="${dados.guerra || ''}"></div>
            <div class="col-6 col-md-12"><input type="text" class="form-control campo-tel" placeholder="Telefone" value="${dados.contato || ''}"></div>
        </div>
    </div>`;
}

// ================= RECIBO E EXPORT =================
export function abrirPreviaRecibo() {
    const rows = document.querySelectorAll('.militar-row');
    let lista = [];
    rows.forEach(row => {
        const posto = row.querySelector('.campo-posto').value.trim();
        const nome = row.querySelector('.campo-nome').value.trim();
        const guerra = row.querySelector('.campo-guerra').value.trim();
        const contato = row.querySelector('.campo-tel').value.trim();
        if(posto && nome && guerra) lista.push({ posto, nome, guerra, contato });
    });

    if(lista.length === 0) return alert("Preencha os dados dos militares.");
    dadosParaEnvio = lista;
    document.getElementById('recibo-evento').innerText = document.getElementById('titulo-evento-form').innerText;
    document.getElementById('recibo-unidade').innerText = perfilAtual.unidade;
    
    const tbody = document.getElementById('recibo-lista-corpo');
    tbody.innerHTML = "";
    lista.forEach(m => tbody.innerHTML += `<tr><td>${m.posto}</td><td><strong>${m.guerra}</strong></td><td>${m.nome}</td><td>${m.contato}</td></tr>`);

    document.getElementById('recibo-modal').classList.remove('d-none');
}

export async function confirmarEnvioRecibo() {
    if (!escalaSelecionadaId || !dadosParaEnvio) return;
    try {
        const jsonString = JSON.stringify(dadosParaEnvio);
        await updateDoc(doc(db, "escalas", escalaSelecionadaId), { militares: jsonString, status: "Preenchido" });
        gerarReciboPDFProfissional(dadosParaEnvio);
        alert("Enviado e Baixado!");
        window.location.reload();
    } catch (e) { alert("Erro: " + e.message); }
}

function gerarReciboPDFProfissional(listaMilitares) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const tituloEvento = document.getElementById('titulo-evento-form').innerText;
    
    doc.setLineWidth(0.5); doc.rect(5, 5, 200, 287);
    doc.setFillColor(153, 0, 0); doc.rect(5, 5, 200, 25, 'F');
    doc.setTextColor(255, 255, 255); doc.setFont("helvetica", "bold"); doc.setFontSize(16);
    doc.text("RECIBO DE ENVIO - CBMMA", 105, 18, null, null, "center");
    
    doc.setTextColor(0, 0, 0); doc.setFontSize(11);
    doc.text(`UNIDADE: ${perfilAtual.unidade}`, 15, 40);
    doc.text(`EVENTO: ${tituloEvento}`, 15, 46);
    doc.text(`DATA: ${new Date().toLocaleString()}`, 15, 52);
    
    let y = 65;
    doc.setFillColor(230, 230, 230); doc.rect(15, 60, 180, 8, 'F');
    doc.setFontSize(9); doc.setFont("helvetica", "bold");
    doc.text("POSTO", 20, 65); doc.text("GUERRA", 50, 65); doc.text("NOME", 90, 65); doc.text("CONTATO", 160, 65);

    doc.setFont("helvetica", "normal");
    listaMilitares.forEach((m, i) => {
        if(i % 2 === 0) { doc.setFillColor(245, 245, 245); doc.rect(15, y-4, 180, 7, 'F'); }
        doc.text(m.posto, 20, y); doc.text(m.guerra.toUpperCase(), 50, y); doc.text(m.nome.substring(0,30), 90, y); doc.text(m.contato, 160, y);
        y += 7;
    });
    doc.save(`Recibo_${perfilAtual.unidade}.pdf`);
}

// === EXCEL PROFISSIONAL PADRÃO ANEXO ===
export async function baixarExcelDoEvento() {
    if (!eventoPreviewAtual) return;
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala');
        const q = query(collection(db, "escalas"), where("evento", "==", eventoPreviewAtual.nome), where("data", "==", eventoPreviewAtual.data), where("status", "==", "Preenchido"));
        const snapshot = await getDocs(q);

        // --- LINHA 1: TÍTULO AMARELO ---
        const titleRow = worksheet.getRow(1);
        worksheet.mergeCells('A1:F1');
        titleRow.getCell(1).value = `${eventoPreviewAtual.nome}  /  ${new Date(eventoPreviewAtual.data).toLocaleDateString('pt-BR', {timeZone:'UTC'})}`;
        
        // Estilo conforme imagem (Amarelo, Texto Preto Negrito, Centralizado)
        titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Amarelo
        titleRow.getCell(1).font = { bold: true, size: 14, color: { argb: 'FF000000' } }; // Preto
        titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        titleRow.height = 30;

        // --- LINHA 2: CABEÇALHO CINZA ---
        const headerRow = worksheet.getRow(2);
        headerRow.values = ['Ord.', 'POSTO/GRAD.', 'NOME', 'CONTATO', 'UBM', 'FUNÇÃO'];
        
        // Estilo Cabeçalho (Cinza, Borda, Negrito, Centralizado)
        for(let i = 1; i <= 6; i++) {
            const cell = headerRow.getCell(i);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0C0C0' } }; // Cinza Claro
            cell.font = { bold: true, color: { argb: 'FF000000' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        }

        // Largura das Colunas
        worksheet.columns = [
            { width: 8 },  // Ord
            { width: 15 }, // Posto
            { width: 50 }, // Nome
            { width: 18 }, // Contato
            { width: 15 }, // UBM
            { width: 20 }  // Função
        ];

        let contador = 1;
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            let militares = [];
            try { militares = JSON.parse(d.militares); } catch { return; }

            militares.forEach(m => {
                const row = worksheet.addRow([
                    contador++, 
                    m.posto, 
                    '', // Placeholder para Nome Rico
                    m.contato, 
                    d.unidade, 
                    d.funcao
                ]);

                // RICH TEXT NO NOME (Guerra em Negrito)
                row.getCell(3).value = {
                    richText: [
                        { text: m.guerra.toUpperCase(), font: { bold: true, name: 'Arial' } },
                        { text: "  " + m.nome.toUpperCase(), font: { bold: false, name: 'Arial' } }
                    ]
                };

                // Formatação Geral das Células da Linha
                row.eachCell((cell) => {
                    cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                });

                // Alinhamento específico para Nome (Esquerda)
                row.getCell(3).alignment = { vertical: 'middle', horizontal: 'left' };
                
                // Função (Amarelo se for Oficial ou Destaque - Opcional, mantendo simples por enquanto)
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), `${eventoPreviewAtual.nome}_OFICIAL.xlsx`);
    } catch (e) { alert("Erro ao gerar Excel: " + e.message); }
}

window.app = { fazerLogin, fazerCadastro, sair, adicionarOrdem, limparOrdens, excluirOrdem, dispararSolicitacao, abrirPreviaRecibo, confirmarEnvioRecibo, abrirPreview, baixarExcelDoEvento, excluirEscalaIndividual, abrirEdicao, excluirEventoCompleto, editarSolicitacaoAdmin, salvarEdicaoAdmin };