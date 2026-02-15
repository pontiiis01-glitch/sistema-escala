import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy, setPersistence, browserSessionPersistence, deleteDoc, writeBatch } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

// === LISTAS E DADOS ===
const UNIDADES_CBMMA_FIXAS = [
    "1 BBM", "2 BBM", "1 CIEBM", "10 BBM", "13 BBM", "16 CIBM", "BBS", "BBA", "BMUS","BBEM","BBMAR", "CGCS", "DEP", "DAT", "DP", "DF", "DPM", "DAL", "CPP", "CPO", 
    "1 Seção", "2 Seção", "3 Seção", "4 Seção", "CAPS", "CRF", "CEPDECMA", "DER", 
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

// === UTILITÁRIOS & SEGURANÇA ===

// FUNÇÃO DE SEGURANÇA (NOVA): Limpa textos para evitar injeção de código
function escapar(str) {
    if (!str) return "";
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function formatarDataLocal(dataString) {
    if(!dataString) return "";
    const partes = dataString.split('-'); 
    return `${partes[2]}/${partes[1]}/${partes[0]}`; 
}

window.formatarTelefoneInput = function(input) {
    let v = input.value.replace(/\D/g, ""); 
    v = v.substring(0, 11); 
    if (v.length > 2) v = v.replace(/^(\d\d)(\d)/g, "$1 $2"); 
    if (v.length > 7) v = v.replace(/(\d{5})(\d)/, "$1-$2"); 
    input.value = v;
}

function gerarCodigoAutenticacao() {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
    let result = '';
    for (let i = 0; i < 8; i++) {
        if (i > 0 && i % 4 === 0) result += '-';
        result += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return `CBMMA-${new Date().getFullYear()}-${result}`;
}

// === INICIALIZAÇÃO ===
document.addEventListener('DOMContentLoaded', () => {
    popularSelectCadastroEFuncoes();
});

function popularSelectCadastroEFuncoes() {
    const selCadastro = document.getElementById('unidade-cadastro');
    if(selCadastro && selCadastro.options.length <= 1) {
        selCadastro.innerHTML = "<option value=''>Selecione a Unidade...</option>";
        UNIDADES_CBMMA_FIXAS.forEach(u => selCadastro.innerHTML += `<option value="${u}">${u}</option>`);
    }

    const selectsFuncao = [document.getElementById('select-funcao'), document.getElementById('edit-admin-funcao')];
    selectsFuncao.forEach(sel => {
        if(sel) {
            sel.innerHTML = "";
            FUNCOES_TATICAS.forEach(f => sel.innerHTML += `<option value="${f}">${f}</option>`);
        }
    });
}

async function carregarUnidadesCadastradasNoAdmin() {
    const selAdmin = document.getElementById('select-unidade');
    if(!selAdmin) return;
    selAdmin.innerHTML = "<option value=''>Carregando...</option>";
    try {
        const q = query(collection(db, "usuarios"), where("funcao", "==", "escalante"));
        const snapshot = await getDocs(q);
        const unidadesReais = [];
        snapshot.forEach(doc => { const data = doc.data(); if(data.unidade) unidadesReais.push(data.unidade); });
        const unidadesUnicas = [...new Set(unidadesReais)].sort();
        selAdmin.innerHTML = "<option value=''>Selecione a Unidade...</option>";
        if(unidadesUnicas.length === 0) selAdmin.innerHTML += "<option disabled>Nenhuma unidade cadastrada</option>";
        else unidadesUnicas.forEach(u => selAdmin.innerHTML += `<option value="${escapar(u)}">${escapar(u)}</option>`);
    } catch (e) { console.error(e); selAdmin.innerHTML = "<option value=''>Erro ao carregar</option>"; }
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
                    carregarUnidadesCadastradasNoAdmin(); 
                } else {
                    document.getElementById('unidade-area').style.display = 'block';
                    carregarPendenciasUnidade();
                }
            }, 300);
        }
    }
});

// ================= ADMIN: ADICIONAR ORDEM COM 5 CATEGORIAS =================
export function adicionarOrdem() {
    const unidade = document.getElementById('select-unidade').value;
    const funcao = document.getElementById('select-funcao').value;
    
    // Captura os 5 inputs
    const sup = document.getElementById('input-sup').value || 0;
    const int = document.getElementById('input-int').value || 0;
    const sub = document.getElementById('input-sub').value || 0;
    const esp = document.getElementById('input-esp').value || 0;
    const pra = document.getElementById('input-pra').value || 0;

    if (!unidade) return alert("Selecione uma unidade!");
    if (sup==0 && int==0 && sub==0 && esp==0 && pra==0) return alert("Defina a quantidade de militares.");

    listaOrdensTemporaria.push({ 
        id: Date.now(), 
        unidade, 
        funcao, 
        cota: { superior: sup, intermediario: int, subalterno: sub, especial: esp, praca: pra }
    });
    atualizarTabelaOrdens();
    
    // Limpa inputs
    ['input-sup','input-int','input-sub','input-esp','input-pra'].forEach(id => document.getElementById(id).value = '');
}

function atualizarTabelaOrdens() {
    const corpo = document.getElementById('tabela-ordens-body');
    document.getElementById('contador-ordens').innerText = `${listaOrdensTemporaria.length}`;
    corpo.innerHTML = "";
    
    listaOrdensTemporaria.forEach((item, index) => {
        // Cria resumo visual da cota
        const c = item.cota;
        let resumo = [];
        if(c.superior > 0) resumo.push(`${c.superior} SUP`);
        if(c.intermediario > 0) resumo.push(`${c.intermediario} INT`);
        if(c.subalterno > 0) resumo.push(`${c.subalterno} SUB`);
        if(c.especial > 0) resumo.push(`${c.especial} ESP`);
        if(c.praca > 0) resumo.push(`${c.praca} PÇ`);

        // APLICAÇÃO DE SEGURANÇA (ESCAPAR)
        corpo.innerHTML += `
            <tr class="border-bottom">
                <td class="fw-bold">${escapar(item.unidade)}</td>
                <td><span class="badge bg-light text-dark border">${escapar(item.funcao)}</span></td>
                <td class="small fw-bold text-muted">${resumo.join(' + ')}</td>
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
                cota: ordem.cota, 
                status: "Pendente", militares: "[]", criadoEm: new Date()
            });
        });
        await Promise.all(promises);
        alert(`Sucesso! Envios realizados.`);
        limparOrdens(); carregarEventosAdmin();
    } catch (e) { alert("Erro de Permissão ou Conexão: " + e.message); }
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
        if (grupos.size === 0) {
            lista.innerHTML = "<div class='text-muted text-center py-3'>Histórico vazio.</div>";
            return;
        }

        const gruposArray = Array.from(grupos.values()).sort((a, b) => new Date(b.data) - new Date(a.data));
        const limiteVisualizacao = 50; 
        
        gruposArray.slice(0, limiteVisualizacao).forEach(info => {
            const dataBr = formatarDataLocal(info.data);
            const percentual = info.total === 0 ? 0 : Math.round((info.respondidos / info.total) * 100);
            
            // APLICAÇÃO DE SEGURANÇA (ESCAPAR)
            lista.innerHTML += `
                <div class="list-group-item p-3 border-bottom ios-click" onclick="window.app.abrirPreview('${escapar(info.evento)}', '${info.data}')">
                    <div class="d-flex justify-content-between align-items-center mb-2">
                        <div><strong class="text-dark d-block text-uppercase">${escapar(info.evento)}</strong><small class="text-muted fw-bold">${dataBr}</small></div>
                        <i class="bi bi-chevron-right text-muted"></i>
                    </div>
                    <div class="d-flex justify-content-between small text-muted align-items-center mb-1"><span>Progresso</span><span>${info.respondidos}/${info.total}</span></div>
                    <div class="progress" style="height: 6px; border-radius: 10px;"><div class="progress-bar bg-success" style="width: ${percentual}%; border-radius: 10px;"></div></div>
                </div>`;
        });
        
        if (gruposArray.length > limiteVisualizacao) {
            lista.innerHTML += `<div class="text-center py-3 small text-muted">Exibindo os ${limiteVisualizacao} mais recentes de ${gruposArray.length}.</div>`;
        }
    } catch(e) { console.error(e); }
}

// ================= ADMIN PREVIEW & ACTIONS =================
export async function abrirPreview(nomeEvento, dataEvento) {
    eventoPreviewAtual = { nome: nomeEvento, data: dataEvento };
    document.getElementById('preview-modal').classList.add('active'); 
    document.getElementById('preview-titulo').innerText = nomeEvento;
    const corpo = document.getElementById('tabela-preview-corpo');
    corpo.innerHTML = "<tr><td colspan='6' class='text-center py-4'><span class='spinner-border text-danger'></span></td></tr>";

    try {
        const q = query(collection(db, "escalas"), where("evento", "==", nomeEvento), where("data", "==", dataEvento));
        const snapshot = await getDocs(q);
        let html = "";
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            const idDoc = docSnap.id;
            let militares = [];
            try { militares = JSON.parse(d.militares); } catch(e) { militares = []; }
            
            // Tratamento para cotas
            const c = d.cota || {};
            let resumoCota = [];
            if(c.superior) resumoCota.push(`${c.superior} Sup`);
            if(c.intermediario) resumoCota.push(`${c.intermediario} Int`);
            if(c.subalterno) resumoCota.push(`${c.subalterno} Sub`);
            if(c.especial) resumoCota.push(`${c.especial} Esp`);
            if(c.praca) resumoCota.push(`${c.praca} Pç`);
            if(resumoCota.length === 0) {
                 if(c.oficial) resumoCota.push(`${c.oficial} Of`);
                 if(c.praca) resumoCota.push(`${c.praca} Pç`);
            }
            const textoCota = resumoCota.join(' / ');
            const jsonCota = JSON.stringify(c).replace(/"/g, "&quot;");
            
            // APLICAÇÃO DE SEGURANÇA NOS BOTÕES E DADOS
            const btnEdit = `<button onclick="window.app.editarSolicitacaoAdmin('${idDoc}', '${escapar(d.unidade)}', '${escapar(d.funcao)}', '${jsonCota}', '${d.prazoData}', '${d.prazoHora}')" class="btn btn-sm btn-outline-primary border-0 me-1" title="Editar"><i class="bi bi-pencil-square"></i></button>`;
            const btnDelete = `<button onclick="window.app.excluirEscalaIndividual('${idDoc}', '${escapar(d.unidade)}')" class="btn btn-sm btn-outline-danger border-0" title="Excluir"><i class="bi bi-trash-fill"></i></button>`;

            if(d.status === "Pendente") {
                html += `
                <tr class="table-danger border-bottom">
                    <td class="text-center fw-bold text-muted">-</td>
                    <td colspan="3" class="small text-danger fw-bold align-middle">
                        <i class="bi bi-exclamation-circle-fill me-1"></i> PENDENTE: ${escapar(d.unidade)}
                        <br><span class="text-muted fw-normal ms-3">Cota: ${textoCota}</span>
                    </td>
                    <td class="align-middle">${escapar(d.funcao)}</td>
                    <td class="text-end align-middle">${btnEdit}${btnDelete}</td>
                </tr>`;
            } else {
                militares.forEach((m, index) => {
                    html += `<tr>
                        <td class="fw-bold text-center text-muted">${index + 1}</td>
                        <td><span class="fw-bold">${escapar(m.posto)}</span> ${escapar(m.guerra)}</td>
                        <td class="small text-muted">${escapar(m.contato)}</td>
                        <td class="fw-bold text-dark">${escapar(d.unidade)}</td>
                        <td><span class="badge bg-light text-dark border">${escapar(d.funcao)}</span></td>
                        <td class="text-end">${index === 0 ? btnEdit + btnDelete : ''}</td>
                    </tr>`;
                });
            }
        });
        corpo.innerHTML = html;
    } catch(e) { console.error(e); corpo.innerHTML = "<tr><td colspan='6'>Erro ao carregar ou Sem Permissão.</td></tr>"; }
}

export function editarSolicitacaoAdmin(id, unidade, funcao, jsonCota, pData, pHora) {
    idEdicaoAdmin = id;
    const cota = JSON.parse(jsonCota);

    document.getElementById('edit-admin-subtitle').innerText = `Editando: ${unidade}`;
    document.getElementById('edit-admin-funcao').value = funcao;
    
    document.getElementById('edit-admin-sup').value = cota.superior || 0;
    document.getElementById('edit-admin-int').value = cota.intermediario || 0;
    document.getElementById('edit-admin-sub').value = cota.subalterno || 0;
    document.getElementById('edit-admin-esp').value = cota.especial || 0;
    document.getElementById('edit-admin-pra').value = cota.praca || 0;

    document.getElementById('edit-admin-prazo-data').value = pData || '';
    document.getElementById('edit-admin-prazo-hora').value = pHora || '23:59';
    document.getElementById('modal-editar-admin').classList.add('active');
}

export async function salvarEdicaoAdmin() {
    if(!idEdicaoAdmin) return;
    const novaFuncao = document.getElementById('edit-admin-funcao').value;
    
    const novoSup = document.getElementById('edit-admin-sup').value;
    const novoInt = document.getElementById('edit-admin-int').value;
    const novoSub = document.getElementById('edit-admin-sub').value;
    const novoEsp = document.getElementById('edit-admin-esp').value;
    const novoPra = document.getElementById('edit-admin-pra').value;

    const novoPrazoData = document.getElementById('edit-admin-prazo-data').value;
    const novoPrazoHora = document.getElementById('edit-admin-prazo-hora').value;

    try {
        await updateDoc(doc(db, "escalas", idEdicaoAdmin), {
            funcao: novaFuncao,
            cota: { 
                superior: novoSup, 
                intermediario: novoInt, 
                subalterno: novoSub, 
                especial: novoEsp, 
                praca: novoPra 
            },
            prazoData: novoPrazoData,
            prazoHora: novoPrazoHora
        });
        alert("Atualizado com sucesso!");
        document.getElementById('modal-editar-admin').classList.remove('active');
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
    if(!eventoPreviewAtual) return console.error("Erro: Nenhum evento selecionado.");
    const confirmar = prompt(`ATENÇÃO: Isso apagará TODO o histórico do evento "${eventoPreviewAtual.nome}".\n\nDigite "APAGAR" para confirmar:`);
    if (!confirmar || confirmar.toUpperCase() !== "APAGAR") return alert("Operação cancelada.");
    
    document.getElementById('preview-modal').classList.remove('active');
    const listaAdmin = document.getElementById('lista-eventos-admin');
    if(listaAdmin) listaAdmin.innerHTML = "<div class='text-center py-5'><span class='spinner-border text-danger'></span><br>Processando exclusão...</div>";

    try {
        const q = query(collection(db, "escalas"), where("evento", "==", eventoPreviewAtual.nome), where("data", "==", eventoPreviewAtual.data));
        const snapshot = await getDocs(q);
        if (snapshot.empty) { alert("Registro limpo."); carregarEventosAdmin(); return; }
        const batch = writeBatch(db);
        snapshot.forEach(d => batch.delete(d.ref));
        await batch.commit(); 
        alert("Histórico apagado.");
        carregarEventosAdmin();
    } catch(e) { console.error(e); alert("Erro ao excluir: " + e.message); carregarEventosAdmin(); }
}

// ================= ESCALANTE: PREENCHIMENTO INTELIGENTE =================
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

        const pendentes = docs.filter(d => d.status === "Pendente");
        const concluidos = docs.filter(d => d.status !== "Pendente");

        if(pendentes.length > 0) {
            lista.innerHTML += `<div class="col-12"><h6 class="text-danger fw-black mb-2 text-uppercase ls-1"><i class="bi bi-exclamation-triangle-fill me-2"></i>PENDÊNCIAS (Prioridade)</h6></div>`;
            pendentes.forEach(d => lista.innerHTML += gerarCardMissao(d, true));
        } else {
            lista.innerHTML += `<div class="col-12 text-center py-4 bg-white rounded-4 shadow-sm mb-4"><i class="bi bi-check-circle-fill text-success display-4"></i><p class="mt-2 fw-bold text-muted">Tudo em dia!</p></div>`;
        }

        if(concluidos.length > 0) {
            lista.innerHTML += `<div class="col-12 mt-4"><h6 class="text-muted fw-bold mb-2 text-uppercase ls-1 border-top pt-4">Histórico / Enviados</h6></div>`;
            concluidos.forEach(d => lista.innerHTML += gerarCardMissao(d, false));
        }
    } catch(e) { console.error(e); }
}

function gerarCardMissao(d, isPendente) {
    const c = d.cota || {};
    let resumo = [];
    if(c.superior) resumo.push(`${c.superior} Sup`);
    if(c.intermediario) resumo.push(`${c.intermediario} Int`);
    if(c.subalterno) resumo.push(`${c.subalterno} Sub`);
    if(c.especial) resumo.push(`${c.especial} Esp`);
    if(c.praca) resumo.push(`${c.praca} Pç`);
    if(resumo.length === 0) {
        if(c.oficial) resumo.push(`${c.oficial} Of`);
        if(c.praca) resumo.push(`${c.praca} Pç`);
    }

    let isBloqueado = false;
    let textoPrazo = "";
    let btnClass = isPendente ? "btn-tactical" : "btn-outline-success";
    let btnText = isPendente ? "RESPONDER AGORA" : "VER / EDITAR";
    let cardOpacity = isPendente ? "1" : "0.85";
    
    if (d.prazoData) {
        const dataLimiteStr = `${d.prazoData}T${d.prazoHora || '23:59'}:00`;
        const dataLimite = new Date(dataLimiteStr);
        const hoje = new Date();
        if (hoje > dataLimite) {
            isBloqueado = true;
            btnClass = "btn-secondary disabled";
            btnText = "PRAZO ENCERRADO";
            textoPrazo = `<div class="text-danger fw-bold small mt-2"><i class="bi bi-lock-fill"></i> ENCERRADO EM ${formatarDataLocal(d.prazoData)} às ${d.prazoHora}</div>`;
        } else {
            textoPrazo = `<div class="text-dark small mt-2 bg-warning bg-opacity-25 p-1 rounded"><i class="bi bi-clock-history"></i> Prazo: ${formatarDataLocal(d.prazoData)} às ${d.prazoHora}</div>`;
        }
    }

    // APLICAÇÃO DE SEGURANÇA (ESCAPAR)
    return `
        <div class="col-md-6 col-lg-4 animate-up">
            <div class="bg-white p-4 h-100 rounded-4 shadow-sm border border-light d-flex flex-column position-relative mission-card" style="opacity: ${cardOpacity}">
                <div class="d-flex justify-content-between mb-2">
                    <span class="badge bg-dark">${formatarDataLocal(d.data)}</span>
                    <span class="badge ${isPendente ? 'bg-warning text-dark' : 'bg-success'}">${escapar(d.status)}</span>
                </div>
                <h5 class="fw-bold mb-0 text-dark text-uppercase">${escapar(d.evento)}</h5>
                <small class="text-muted mb-2 d-block">${d.horaInicio} às ${d.horaFim}</small>
                <div class="bg-light p-3 rounded border text-center my-2">
                    <strong class="d-block text-primary">${escapar(d.funcao)}</strong>
                    <div class="small text-muted">${resumo.join(' / ')}</div>
                </div>
                ${textoPrazo}
                <button onclick="window.app.abrirEdicao('${d.id}')" class="btn ${btnClass} w-100 fw-bold mt-auto py-3 rounded-3 shadow-sm ios-click" ${isBloqueado ? 'disabled' : ''}>${btnText}</button>
            </div>
        </div>`;
}

// === EDIÇÃO COM EXEMPLOS VISUAIS (Accordion) ===
export async function abrirEdicao(id) {
    escalaSelecionadaId = id;
    const docSnap = await getDoc(doc(db, "escalas", id));
    const d = docSnap.data();
    const c = d.cota || {};
    
    document.getElementById('titulo-evento-form').innerText = d.evento;
    document.getElementById('subtitulo-form').innerText = d.funcao;
    
    const container = document.getElementById('container-inputs-militares');
    container.innerHTML = "";

    // --- ÁREA DE EXEMPLOS ---
    container.innerHTML += `
        <div class="mb-4 text-center">
            <span class="text-muted small fw-bold text-uppercase d-block mb-2">Dúvidas no preenchimento?</span>
            <button onclick="let el = document.getElementById('box-exemplos'); el.style.display = el.style.display === 'none' ? 'block' : 'none';" class="btn btn-sm btn-outline-secondary rounded-pill px-3 fw-bold small ios-click" style="font-size: 0.7rem;">
                <i class="bi bi-info-circle-fill me-1"></i> VER PADRÃO DE POSTOS
            </button>
            <div id="box-exemplos" class="mt-3 p-3 bg-white rounded-4 border shadow-sm animate-up" style="display: none; text-align: left;">
                <div class="d-flex align-items-center mb-2">
                    <i class="bi bi-lightbulb-fill text-warning me-2"></i>
                    <span class="fw-bold text-dark small text-uppercase">Modelos de Referência</span>
                </div>
                <p class="text-muted small mb-3" style="font-size: 0.75rem;">Utilize as siglas abaixo para padronizar o documento (Exemplos ilustrativos):</p>
                <div class="d-flex flex-wrap gap-2">
                    <span class="badge bg-light text-secondary border">TEN CEL QOC</span>
                    <span class="badge bg-light text-secondary border">MAJ QOC / QOE / QOA</span>
                    <span class="badge bg-light text-secondary border">CAP QOC / QOA / QOE</span>
                    <span class="badge bg-light text-secondary border">1 TEN QOC / QOA / QOE</span>
                    <span class="badge bg-light text-secondary border">ASP OF BM</span>
                    <span class="badge bg-light text-secondary border">CAD BM/3</span>
                    <span class="badge bg-light text-secondary border">ST BM</span>
                    <span class="badge bg-light text-secondary border">2 SGT BM</span>
                    <span class="badge bg-light text-secondary border">CB BM</span>
                    <span class="badge bg-light text-secondary border">SD BM</span>
                </div>
            </div>
        </div>
    `;
    
    let dadosSalvos = [];
    try { dadosSalvos = JSON.parse(d.militares); } catch {}
    
    let contador = 0;
    const gerarLoop = (qtd, rotulo) => {
        const num = parseInt(qtd) || 0;
        for(let i=0; i < num; i++) {
            container.innerHTML += gerarHtmlMilitar(i, rotulo, dadosSalvos[contador++] || {});
        }
    };

    if(c.oficial) gerarLoop(c.oficial, 'OFICIAL');
    gerarLoop(c.superior, 'OF. SUPERIOR');
    gerarLoop(c.intermediario, 'OF. INTERMEDIÁRIO');
    gerarLoop(c.subalterno, 'OF. SUBALTERNO');
    gerarLoop(c.especial, 'PRAÇA ESPECIAL');
    if(c.praca) gerarLoop(c.praca, 'PRAÇA');

    document.getElementById('form-militar-modal').classList.add('active'); 
}

function gerarHtmlMilitar(index, tipo, dados) {
    // Usamos 'escapar' dentro dos atributos value para evitar que aspas quebrem o HTML
    return `
    <div class="p-3 bg-white rounded-3 border mb-3 militar-row shadow-sm">
        <span class="badge bg-secondary mb-2">${tipo} ${index + 1}</span>
        <div class="row g-2">
            <div class="col-4 col-md-3">
                <input type="text" class="form-control campo-posto fw-bold" placeholder="Posto" value="${escapar(dados.posto || '')}" oninput="this.value = this.value.toUpperCase()">
            </div>
            <div class="col-8 col-md-5">
                <input type="text" class="form-control campo-nome" placeholder="Nome Completo" value="${escapar(dados.nome || '')}" oninput="this.value = this.value.toUpperCase()">
            </div>
            <div class="col-6 col-md-4">
                <input type="text" class="form-control campo-guerra fw-bold text-uppercase" placeholder="Nome Guerra" value="${escapar(dados.guerra || '')}" oninput="this.value = this.value.toUpperCase()">
            </div>
            <div class="col-6 col-md-12">
                <input type="text" class="form-control campo-tel" placeholder="98 9XXXX-XXXX" value="${escapar(dados.contato || '')}" maxlength="15" oninput="window.formatarTelefoneInput(this)">
            </div>
        </div>
    </div>`;
}

// ================= VALIDAÇÃO PÚBLICA SEGURA (LÊ DE VALIDACOES_PUBLICAS) =================
export function abrirValidador() {
    document.getElementById('modal-validador').classList.add('active');
    document.getElementById('input-codigo-validacao').value = "";
    document.getElementById('resultado-validacao').style.display = 'none';
}

export async function consultarAutenticidade() {
    const codigo = document.getElementById('input-codigo-validacao').value.trim().toUpperCase();
    const divResult = document.getElementById('resultado-validacao');
    
    if (codigo.length < 10) return alert("Código inválido.");

    divResult.style.display = 'block';
    divResult.innerHTML = "<div class='text-center text-muted'><span class='spinner-border spinner-border-sm me-2'></span>Consultando base pública...</div>";

    try {
        const docRef = doc(db, "validacoes_publicas", codigo);
        const docSnap = await getDoc(docRef);

        if (!docSnap.exists()) {
            divResult.innerHTML = `
                <div class="alert alert-danger text-center border-0 shadow-sm rounded-4">
                    <i class="bi bi-x-circle-fill display-4 d-block mb-2"></i>
                    <strong class="d-block text-uppercase">Documento Inválido</strong>
                    <span class="small">Este código não consta nos registros oficiais.</span>
                    <button onclick="fecharModal('modal-validador')" class="btn btn-light w-100 mt-3 fw-bold rounded-pill text-uppercase">Fechar</button>
                </div>`;
        } else {
            const d = docSnap.data();
            const dataValidacao = d.dataValidacao ? new Date(d.dataValidacao).toLocaleString() : "Data desconhecida";
            
            // APLICAÇÃO DE SEGURANÇA (ESCAPAR)
            divResult.innerHTML = `
                <div class="alert alert-success text-center border-0 shadow-sm rounded-4">
                    <i class="bi bi-patch-check-fill display-4 d-block mb-2"></i>
                    <strong class="d-block text-uppercase mb-2">Documento Autêntico</strong>
                    <div class="text-start bg-white p-3 rounded-3 border small text-muted">
                        <strong>Evento:</strong> ${escapar(d.evento)}<br>
                        <strong>Unidade:</strong> ${escapar(d.unidade)}<br>
                        <strong>Situação:</strong> ${escapar(d.status)}<br>
                        <strong>Emitido em:</strong> ${dataValidacao}
                    </div>
                    <div class="mt-2 text-center text-muted" style="font-size: 0.65rem;">
                        <i class="bi bi-lock-fill"></i> Dados pessoais protegidos pela LGPD.
                    </div>
                    <button onclick="fecharModal('modal-validador')" class="btn btn-dark w-100 mt-3 fw-bold rounded-pill text-uppercase">Fechar</button>
                </div>`;
        }
    } catch (e) { 
        console.error(e);
        divResult.innerHTML = `<div class="text-danger text-center">Erro de conexão.</div>`; 
    }
}

// ================= EXPORTAÇÃO PDF & EXCEL (SALVA DUPLO: PRIVADO E PÚBLICO) =================
export function abrirPreviaRecibo() {
    const rows = document.querySelectorAll('.militar-row');
    let lista = [];
    rows.forEach(row => {
        const posto = row.querySelector('.campo-posto').value.trim().toUpperCase();
        const nome = row.querySelector('.campo-nome').value.trim().toUpperCase();
        const guerra = row.querySelector('.campo-guerra').value.trim().toUpperCase();
        const contato = row.querySelector('.campo-tel').value.trim(); 
        if(posto && nome && guerra) lista.push({ posto, nome, guerra, contato });
    });
    if(lista.length === 0) return alert("Preencha os dados dos militares.");
    dadosParaEnvio = lista;
    document.getElementById('recibo-evento').innerText = document.getElementById('titulo-evento-form').innerText;
    document.getElementById('recibo-unidade').innerText = perfilAtual.unidade;
    const tbody = document.getElementById('recibo-lista-corpo');
    tbody.innerHTML = "";
    // Aqui usamos innerHTML para preview, então usamos escapar
    lista.forEach(m => tbody.innerHTML += `<tr><td>${escapar(m.posto)}</td><td><strong>${escapar(m.guerra)}</strong></td><td>${escapar(m.nome)}</td><td>${escapar(m.contato)}</td></tr>`);
    document.getElementById('recibo-modal').classList.add('active');
}

export async function confirmarEnvioRecibo() {
    if (!escalaSelecionadaId || !dadosParaEnvio) return;
    
    const codigoAuth = gerarCodigoAutenticacao();
    const dataHoraEnvio = new Date().toISOString(); 
    const tituloEvento = document.getElementById('titulo-evento-form').innerText;

    try {
        const jsonString = JSON.stringify(dadosParaEnvio);
        
        // 1. Salva a Escala Completa (Sigilosa - Com telefones) na coleção PRIVADA
        await updateDoc(doc(db, "escalas", escalaSelecionadaId), { 
            militares: jsonString, 
            status: "Preenchido", 
            codigoAutenticacao: codigoAuth, 
            dataValidacao: dataHoraEnvio 
        });

        // 2. Cria o Registro Público (Apenas dados genéricos) na coleção PÚBLICA
        await setDoc(doc(db, "validacoes_publicas", codigoAuth), {
            codigo: codigoAuth,
            evento: tituloEvento,
            unidade: perfilAtual.unidade,
            dataValidacao: dataHoraEnvio,
            status: "Válido"
        });

        gerarReciboPDFProfissional(dadosParaEnvio, codigoAuth);
        alert("Enviado, Autenticado e Protegido!");
        window.location.reload();
    } catch (e) { alert("Erro ao enviar. Verifique se você é da unidade correta.\nDetalhe: " + e.message); }
}

function gerarReciboPDFProfissional(listaMilitares, codigoAuth) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const tituloEvento = document.getElementById('titulo-evento-form').innerText;
    
    doc.setLineWidth(0.5); doc.rect(5, 5, 200, 287);
    doc.setFillColor(153, 0, 0); doc.rect(5, 5, 200, 25, 'F');
    doc.setTextColor(255, 255, 255); doc.setFont("helvetica", "bold"); doc.setFontSize(16);
    doc.text("COMPROVANTE DE ENVIO - CBMMA", 105, 18, null, null, "center");
    
    doc.setTextColor(0, 0, 0); doc.setFontSize(11);
    doc.text(`UNIDADE: ${perfilAtual.unidade}`, 15, 40);
    doc.text(`EVENTO: ${tituloEvento}`, 15, 46);
    doc.text(`DATA EMISSÃO: ${new Date().toLocaleString()}`, 15, 52);

    let y = 65;
    doc.setFillColor(200, 200, 200); doc.rect(15, 60, 180, 8, 'F');
    doc.setFontSize(9); doc.setFont("helvetica", "bold");
    doc.text("POSTO", 20, 65); doc.text("GUERRA", 50, 65); doc.text("NOME COMPLETO", 90, 65); doc.text("CONTATO", 160, 65);
    
    doc.setFont("helvetica", "normal");
    listaMilitares.forEach((m, i) => {
        if(i % 2 === 0) { doc.setFillColor(245, 245, 245); doc.rect(15, y-4, 180, 7, 'F'); }
        doc.text(m.posto, 20, y); doc.text(m.guerra, 50, y); 
        let nomeFormatado = m.nome.length > 35 ? m.nome.substring(0, 32) + "..." : m.nome;
        doc.text(nomeFormatado, 90, y); 
        doc.text(m.contato, 160, y);
        y += 7;
    });

    const pageHeight = doc.internal.pageSize.height;
    doc.setDrawColor(150);
    doc.line(15, pageHeight - 30, 195, pageHeight - 30);
    doc.setFont("courier", "bold"); doc.setFontSize(8); doc.setTextColor(80, 80, 80);
    doc.text("AUTENTICAÇÃO DIGITAL DO SISTEMA S.I.E.G.E.", 105, pageHeight - 22, null, null, "center");
    doc.setFontSize(10); doc.setTextColor(0, 0, 0);
    doc.text(`CHAVE: ${codigoAuth || "VALIDAÇÃO PENDENTE"}`, 105, pageHeight - 16, null, null, "center");
    doc.setFontSize(7); doc.setTextColor(100, 100, 100);
    doc.text("Este documento possui validade administrativa interna no CBMMA.", 105, pageHeight - 10, null, null, "center");
    doc.save(`Comprovante_${perfilAtual.unidade}_${codigoAuth}.pdf`);
}

export async function baixarExcelDoEvento() {
    if (!eventoPreviewAtual) return;
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala');
        const q = query(collection(db, "escalas"), where("evento", "==", eventoPreviewAtual.nome), where("data", "==", eventoPreviewAtual.data), where("status", "==", "Preenchido"));
        const snapshot = await getDocs(q);

        let horarioTexto = "HORÁRIO INDEFINIDO";
        if (!snapshot.empty) {
            const dPrimeiro = snapshot.docs[0].data();
            if(dPrimeiro.horaInicio && dPrimeiro.horaFim) { horarioTexto = `${dPrimeiro.horaInicio} às ${dPrimeiro.horaFim}`; }
        }

        const titleRow = worksheet.getRow(1);
        worksheet.mergeCells('A1:F1');
        titleRow.getCell(1).value = `${eventoPreviewAtual.nome}  /  ${formatarDataLocal(eventoPreviewAtual.data)}  /  ${horarioTexto}`;
        titleRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        titleRow.getCell(1).font = { bold: true, size: 14, color: { argb: 'FF000000' } };
        titleRow.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        titleRow.height = 30;

        const headerRow = worksheet.getRow(2);
        headerRow.values = ['Ord.', 'POSTO/GRAD.', 'NOME', 'CONTATO', 'UBM', 'FUNÇÃO'];
        
        for(let i = 1; i <= 6; i++) {
            const cell = headerRow.getCell(i);
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC0C0C0' } };
            cell.font = { bold: true, color: { argb: 'FF000000' } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        }
        worksheet.columns = [{ width: 8 }, { width: 15 }, { width: 50 }, { width: 18 }, { width: 15 }, { width: 20 }];

        let contador = 1;
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            let militares = [];
            try { militares = JSON.parse(d.militares); } catch { return; }
            militares.forEach(m => {
                const row = worksheet.addRow([ contador++, m.posto, '', m.contato, d.unidade, d.funcao ]);
                const nomeUpper = m.nome.toUpperCase().trim();
                const guerraUpper = m.guerra.toUpperCase().trim();
                const indexGuerra = nomeUpper.indexOf(guerraUpper);
                let richTextValue = [];

                if (indexGuerra !== -1 && guerraUpper.length > 0) {
                    if (indexGuerra > 0) richTextValue.push({ text: nomeUpper.substring(0, indexGuerra), font: { bold: false, name: 'Arial' } });
                    richTextValue.push({ text: nomeUpper.substring(indexGuerra, indexGuerra + guerraUpper.length), font: { bold: true, name: 'Arial' } });
                    if (indexGuerra + guerraUpper.length < nomeUpper.length) richTextValue.push({ text: nomeUpper.substring(indexGuerra + guerraUpper.length), font: { bold: false, name: 'Arial' } });
                } else { richTextValue = [{ text: nomeUpper, font: { bold: false, name: 'Arial' } }]; }

                row.getCell(3).value = { richText: richTextValue };
                row.eachCell((cell) => {
                    cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                    cell.alignment = { vertical: 'middle', horizontal: 'center' };
                });
                row.getCell(3).alignment = { vertical: 'middle', horizontal: 'left' };
            });
        });
        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), `${eventoPreviewAtual.nome}_OFICIAL.xlsx`);
    } catch (e) { alert("Erro ao gerar Excel: " + e.message); }
}

window.app = { fazerLogin, fazerCadastro, sair, adicionarOrdem, limparOrdens, excluirOrdem, dispararSolicitacao, abrirPreviaRecibo, confirmarEnvioRecibo, abrirPreview, baixarExcelDoEvento, excluirEscalaIndividual, abrirEdicao, excluirEventoCompleto, editarSolicitacaoAdmin, salvarEdicaoAdmin, abrirValidador, consultarAutenticidade };