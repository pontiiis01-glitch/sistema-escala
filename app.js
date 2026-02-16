import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy, setPersistence, browserSessionPersistence, deleteDoc, writeBatch, limit, deleteUser } from './firebase-config.js';
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

// NOVA GERAÇÃO DE CÓDIGO (UUID Robusto)
function gerarCodigoAutenticacao() {
    const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const sections = [8, 4, 4, 4, 12]; // Formato: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
    let codigo = '';
    
    sections.forEach((len, index) => {
        for(let i=0; i<len; i++) {
            codigo += chars.charAt(Math.floor(Math.random() * chars.length));
        }
        if(index < sections.length - 1) codigo += '-';
    });
    
    return codigo;
}

// Auxiliar para carregar imagens no PDF
async function carregarImagemBase64(url) {
    try {
        const response = await fetch(url);
        const blob = await response.blob();
        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.readAsDataURL(blob);
        });
    } catch (e) {
        console.error("Erro ao carregar imagem", e);
        return null;
    }
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
    const email = document.getElementById('email-login').value.trim();
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
        document.getElementById('msg-erro').innerText = "Email ou senha incorretos.";
        btn.innerHTML = textoOriginal; btn.disabled = false;
    }
}

export async function fazerCadastro() {
    const email = document.getElementById('email-cadastro').value.trim();
    const senha = document.getElementById('senha-cadastro').value;
    const unidadeInput = document.getElementById('unidade-cadastro').value;
    
    if(!email || !senha || !unidadeInput) return alert("Preencha todos os campos.");
    if(senha.length < 6) return alert("A senha deve ter no mínimo 6 caracteres.");

    const unidadeLimpa = unidadeInput.trim().toUpperCase();
    if (!UNIDADES_CBMMA_FIXAS.includes(unidadeLimpa)) {
        return alert("Erro de Segurança: Unidade inválida ou não autorizada.");
    }

    try {
        const cred = await createUserWithEmailAndPassword(auth, email, senha);
        const q = query(collection(db, "usuarios"), where("unidade", "==", unidadeLimpa));
        const snapshot = await getDocs(q);

        if (!snapshot.empty) {
            await deleteUser(cred.user); 
            alert(`ATENÇÃO: A unidade ${unidadeLimpa} já possui um cadastro ativo.`);
            window.location.reload();
            return;
        }

        await setDoc(doc(db, "usuarios", cred.user.uid), { 
            email, 
            unidade: unidadeLimpa, 
            funcao: "escalante" 
        });
        alert("Unidade cadastrada com sucesso!"); window.location.reload();
    } catch (e) { 
        if(e.code === 'auth/email-already-in-use') alert("Este email já está cadastrado.");
        else {
            console.error(e);
            alert("Erro ao cadastrar: " + e.message);
        }
    }
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

// ================= ADMIN: GERENCIAMENTO =================
export function adicionarOrdem() {
    const unidade = document.getElementById('select-unidade').value;
    const funcao = document.getElementById('select-funcao').value;
    const sup = document.getElementById('input-sup').value || 0;
    const int = document.getElementById('input-int').value || 0;
    const sub = document.getElementById('input-sub').value || 0;
    const esp = document.getElementById('input-esp').value || 0;
    const pra = document.getElementById('input-pra').value || 0;

    if (!unidade) return alert("Selecione uma unidade!");
    
    const jaExiste = listaOrdensTemporaria.some(o => o.unidade === unidade);
    if(jaExiste) return alert(`A unidade ${unidade} já está na lista!`);

    if (sup==0 && int==0 && sub==0 && esp==0 && pra==0) return alert("Defina a quantidade de militares.");

    listaOrdensTemporaria.push({ 
        id: Date.now(), 
        unidade, 
        funcao, 
        cota: { superior: sup, intermediario: int, subalterno: sub, especial: esp, praca: pra }
    });
    atualizarTabelaOrdens();
    
    ['input-sup','input-int','input-sub','input-esp','input-pra'].forEach(id => document.getElementById(id).value = '');
}

function atualizarTabelaOrdens() {
    const corpo = document.getElementById('tabela-ordens-body');
    document.getElementById('contador-ordens').innerText = `${listaOrdensTemporaria.length}`;
    corpo.innerHTML = "";
    
    listaOrdensTemporaria.forEach((item, index) => {
        const c = item.cota;
        let resumo = [];
        if(c.superior > 0) resumo.push(`${c.superior} SUP`);
        if(c.intermediario) resumo.push(`${c.intermediario} INT`);
        if(c.subalterno) resumo.push(`${c.subalterno} SUB`);
        if(c.especial) resumo.push(`${c.especial} ESP`);
        if(c.praca > 0) resumo.push(`${c.praca} PÇ`);

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
        const batch = writeBatch(db);
        
        listaOrdensTemporaria.forEach(ordem => {
            const novaRef = doc(collection(db, "escalas")); 
            batch.set(novaRef, {
                evento: evento.toUpperCase(), 
                data, 
                horaInicio, 
                horaFim,
                prazoData, 
                prazoHora: prazoHora || "23:59",
                unidade: ordem.unidade, 
                funcao: ordem.funcao,
                cota: ordem.cota, 
                status: "Pendente", 
                militares: "[]", 
                criadoEm: new Date()
            });
        });

        await batch.commit(); 
        alert(`Sucesso! Todas as solicitações foram enviadas.`);
        limparOrdens(); 
        carregarEventosAdmin();
    } catch (e) { 
        console.error(e);
        alert("Erro no envio em massa: " + e.message); 
    }
}

async function carregarEventosAdmin() {
    const lista = document.getElementById('lista-eventos-admin');
    lista.innerHTML = "<div class='text-center py-3'><span class='spinner-border text-danger'></span></div>";
    try {
        const q = query(collection(db, "escalas"), orderBy("criadoEm", "desc"), limit(300)); 
        
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
            lista.innerHTML += `<div class="text-center py-3 small text-muted">Exibindo os ${limiteVisualizacao} mais recentes.</div>`;
        }
    } catch(e) { 
        console.error(e);
        if(e.code === 'failed-precondition') {
             alert("Aviso de Sistema: É necessário criar um índice no Firebase. Verifique o console do navegador (F12) e clique no link fornecido pelo Google.");
        }
    }
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
    let cota = {};
    try { cota = JSON.parse(jsonCota); } catch(e) { cota = { superior:0, intermediario:0, subalterno:0, especial:0, praca:0 }; }

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
            cota: { superior: novoSup, intermediario: novoInt, subalterno: novoSub, especial: novoEsp, praca: novoPra },
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
    
    let isBloqueado = false;
    let textoPrazo = "";
    let btnClass = isPendente ? "btn-tactical" : "btn-outline-success";
    let btnText = isPendente ? "RESPONDER AGORA" : "VER / EDITAR";
    let cardOpacity = isPendente ? "1" : "0.85";
    
    if (d.prazoData) {
        // Correção de Fuso
        const pAno = parseInt(d.prazoData.split('-')[0]);
        const pMes = parseInt(d.prazoData.split('-')[1]) - 1; 
        const pDia = parseInt(d.prazoData.split('-')[2]);
        const pHora = parseInt((d.prazoHora || '23:59').split(':')[0]);
        const pMin = parseInt((d.prazoHora || '23:59').split(':')[1]);
        
        const dataLimite = new Date(pAno, pMes, pDia, pHora, pMin, 59);
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

// === EDIÇÃO COM EXEMPLOS VISUAIS ===
export async function abrirEdicao(id) {
    escalaSelecionadaId = id;
    const docSnap = await getDoc(doc(db, "escalas", id));
    const d = docSnap.data();
    const c = d.cota || {};
    
    document.getElementById('titulo-evento-form').innerText = d.evento;
    document.getElementById('subtitulo-form').innerText = d.funcao;
    
    const container = document.getElementById('container-inputs-militares');
    container.innerHTML = "";

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

// ================= VALIDAÇÃO PÚBLICA =================
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

// ================= EXPORTAÇÃO PDF & EXCEL =================
export function abrirPreviaRecibo() {
    const rows = document.querySelectorAll('.militar-row');
    let lista = [];
    rows.forEach(row => {
        const posto = row.querySelector('.campo-posto').value.trim().toUpperCase();
        const nome = row.querySelector('.campo-nome').value.trim().toUpperCase();
        const guerra = row.querySelector('.campo-guerra').value.trim().toUpperCase();
        const contato = row.querySelector('.campo-tel').value.trim(); 
        
        if(posto && nome && guerra && contato.length >= 8) {
            lista.push({ posto, nome, guerra, contato });
        }
    });

    if(lista.length === 0) return alert("Preencha os dados dos militares. O telefone é obrigatório.");
    if(lista.length < rows.length) return alert("Atenção: Algum militar não foi adicionado pois o telefone ou nome estava incompleto.");

    dadosParaEnvio = lista;
    document.getElementById('recibo-evento').innerText = document.getElementById('titulo-evento-form').innerText;
    document.getElementById('recibo-unidade').innerText = perfilAtual.unidade;
    const tbody = document.getElementById('recibo-lista-corpo');
    tbody.innerHTML = "";
    lista.forEach(m => tbody.innerHTML += `<tr><td>${escapar(m.posto)}</td><td><strong>${escapar(m.guerra)}</strong></td><td>${escapar(m.nome)}</td><td>${escapar(m.contato)}</td></tr>`);
    document.getElementById('recibo-modal').classList.add('active');
}

export function abrirTelaAssinatura() {
    if (!escalaSelecionadaId || !dadosParaEnvio) return;
    
    // Abre o modal de assinatura
    document.getElementById('modal-assinatura').classList.add('active');
    // Limpa campos
    document.getElementById('assinatura-nome').value = '';
    document.getElementById('assinatura-funcao').value = '';
}

export async function finalizarEnvioComAssinatura() {
    const nomeAssinatura = document.getElementById('assinatura-nome').value.trim().toUpperCase();
    const funcaoAssinatura = document.getElementById('assinatura-funcao').value.trim().toUpperCase();

    if(nomeAssinatura.length < 5 || funcaoAssinatura.length < 3) {
        return alert("Por favor, preencha corretamente o Posto/Nome e a Função.");
    }

    // Fecha modais para não atrapalhar
    document.getElementById('modal-assinatura').classList.remove('active');
    document.getElementById('recibo-modal').classList.remove('active');

    const codigoAuth = gerarCodigoAutenticacao();
    const dataHoraEnvio = new Date().toISOString(); 
    const tituloEvento = document.getElementById('titulo-evento-form').innerText;

    try {
        const jsonString = JSON.stringify(dadosParaEnvio);
        
        // Pega dados da escala para saber horário
        const docSnap = await getDoc(doc(db, "escalas", escalaSelecionadaId));
        const dadosEscala = docSnap.data();
        
        await updateDoc(doc(db, "escalas", escalaSelecionadaId), { 
            militares: jsonString, 
            status: "Preenchido", 
            codigoAutenticacao: codigoAuth, 
            dataValidacao: dataHoraEnvio,
            assinadoPor: nomeAssinatura,
            assinadoFuncao: funcaoAssinatura
        });

        await setDoc(doc(db, "validacoes_publicas", codigoAuth), {
            codigo: codigoAuth,
            evento: tituloEvento,
            unidade: perfilAtual.unidade,
            dataValidacao: dataHoraEnvio,
            status: "Válido"
        });

        await gerarReciboPDFProfissional(dadosParaEnvio, codigoAuth, nomeAssinatura, funcaoAssinatura, dadosEscala);
        alert("Enviado, Autenticado e Protegido!");
        window.location.reload();
    } catch (e) { alert("Erro ao enviar. Verifique se você é da unidade correta.\nDetalhe: " + e.message); }
}

export async function confirmarEnvioRecibo() {
    // Função mantida para compatibilidade, mas redireciona para assinatura
    abrirTelaAssinatura();
}

async function gerarReciboPDFProfissional(listaMilitares, codigoAuth, nomeAssinatura, funcaoAssinatura, dadosEscala) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const tituloEvento = document.getElementById('titulo-evento-form').innerText;
    const unidadeNome = perfilAtual.unidade;
    
    // --- 1. CONFIGURAÇÃO DE IMAGENS E LOGOS ---
    // Tenta carregar brasões
    const imgLogoEsq = await carregarImagemBase64('brasao.png'); // CBMMA
    const imgLogoDir = await carregarImagemBase64('brasao_unidade.jpg'); // COCB

    // --- 2. CABEÇALHO OFICIAL (Barra Vermelha Estilo Antigo) ---
    // Fundo Vermelho Sólido no Topo
    doc.setFillColor(179, 0, 0); // Vermelho CBMMA
    doc.rect(0, 0, 210, 40, 'F'); // Barra vermelha cobrindo o topo
    
    // Linha Dourada
    doc.setDrawColor(255, 215, 0); // Gold
    doc.setLineWidth(1.5);
    doc.line(0, 41, 210, 41);

    // Logos sobre o vermelho
    if(imgLogoEsq) { doc.addImage(imgLogoEsq, 'PNG', 10, 5, 30, 30); }
    if(imgLogoDir) { doc.addImage(imgLogoDir, 'JPEG', 170, 5, 30, 30); }
    
    // Texto do Cabeçalho (Branco)
    doc.setTextColor(255, 255, 255);
    doc.setFont("times", "bold");
    doc.setFontSize(12);
    const centro = 105;

    doc.text("ESTADO DO MARANHÃO", centro, 12, { align: "center" });
    doc.text("SECRETARIA DE SEGURANÇA PÚBLICA", centro, 18, { align: "center" });
    doc.text("CORPO DE BOMBEIROS MILITAR DO MARANHÃO", centro, 24, { align: "center" });
    doc.text("COMANDO OPERACIONAL METROPOLITANO", centro, 30, { align: "center" });
    
    // Reset Cor Texto
    doc.setTextColor(0, 0, 0);
    
    let y = 55;
    
    // Unidade e Título
    doc.setFont("helvetica", "bold");
    doc.setFontSize(14);
    doc.text(unidadeNome, centro, y, { align: "center" }); y += 10;

    // Nome da Operação (Grande destaque)
    doc.setFontSize(12);
    doc.text(tituloEvento, centro, y, { align: "center" }); y += 8;
    
    // Detalhes da Missão (Horário)
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    
    const horarioTexto = (dadosEscala && dadosEscala.horaInicio) ? `${dadosEscala.horaInicio} às ${dadosEscala.horaFim}` : "A DEFINIR";
    doc.text(`HORÁRIO DO SERVIÇO: ${horarioTexto}`, 15, y); y += 5;
    doc.text(`DATA DE EMISSÃO: ${new Date().toLocaleDateString('pt-BR')} às ${new Date().toLocaleTimeString('pt-BR')}`, 15, y); y += 10;

    // --- 3. TABELA DE EFETIVO ---
    doc.setFillColor(230, 230, 230); 
    doc.rect(15, y, 180, 8, 'F');
    doc.setFont("helvetica", "bold"); doc.setFontSize(9);
    doc.text("POSTO/GRAD", 20, y+5);
    doc.text("NOME DE GUERRA", 60, y+5);
    doc.text("NOME COMPLETO", 100, y+5);
    doc.text("CONTATO", 165, y+5);
    y += 8;

    doc.setFont("helvetica", "normal");
    listaMilitares.forEach((m, i) => {
        if(i % 2 === 0) { doc.setFillColor(250, 250, 250); doc.rect(15, y, 180, 7, 'F'); }
        doc.text(m.posto, 20, y+5);
        doc.text(m.guerra, 60, y+5);
        let nomeFormatado = m.nome.length > 30 ? m.nome.substring(0, 28) + "..." : m.nome;
        doc.text(nomeFormatado, 100, y+5);
        doc.text(m.contato, 165, y+5);
        doc.setDrawColor(200);
        doc.line(15, y+7, 195, y+7);
        y += 7;
    });

    // --- 4. RODAPÉ DE VALIDAÇÃO ---
    const pageHeight = doc.internal.pageSize.height;
    const footerY = pageHeight - 50;

    // QR Code
    const qrDiv = document.createElement('div');
    const qrCodeObj = new QRCode(qrDiv, {
        text: `https://seusite.firebaseapp.com/validar/${codigoAuth}`,
        width: 100, height: 100,
        correctLevel: QRCode.CorrectLevel.H
    });
    
    setTimeout(() => {
        const qrImage = qrDiv.querySelector('img').src;
        if(qrImage) doc.addImage(qrImage, 'PNG', 15, footerY + 5, 30, 30);
    }, 100);

    // Texto de Assinatura
    doc.setFont("courier", "bold");
    doc.setFontSize(9);
    doc.setTextColor(0, 0, 0);
    
    doc.text("DOCUMENTO ASSINADO DIGITALMENTE", 50, footerY + 10);
    doc.setFont("courier", "normal");
    doc.text(`ASSINADO POR: ${nomeAssinatura.toUpperCase()}`, 50, footerY + 15);
    doc.text(`FUNÇÃO: ${funcaoAssinatura.toUpperCase()} - ${unidadeNome}`, 50, footerY + 20);
    doc.text(`DATA DA ASSINATURA: ${new Date().toLocaleString()}`, 50, footerY + 25);
    
    doc.setFontSize(8);
    doc.setTextColor(80, 80, 80);
    doc.text("HASH DE VALIDAÇÃO:", 50, footerY + 33);
    doc.setFont("courier", "bold");
    doc.text(codigoAuth, 50, footerY + 37);

    doc.setFont("helvetica", "normal");
    doc.setFontSize(6);
    doc.text("A autenticidade deste documento pode ser verificada através do QR Code ao lado ou pelo código hash no sistema S.I.E.G.E.", 15, pageHeight - 5);

    doc.save(`ESCALA_${unidadeNome}_${tituloEvento}.pdf`);
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

window.app = { fazerLogin, fazerCadastro, sair, adicionarOrdem, limparOrdens, excluirOrdem, dispararSolicitacao, abrirPreviaRecibo, confirmarEnvioRecibo, abrirPreview, baixarExcelDoEvento, excluirEscalaIndividual, abrirEdicao, excluirEventoCompleto, editarSolicitacaoAdmin, salvarEdicaoAdmin, abrirValidador, consultarAutenticidade, abrirTelaAssinatura, finalizarEnvioComAssinatura };