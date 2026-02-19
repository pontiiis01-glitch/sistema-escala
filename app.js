import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy, setPersistence, browserSessionPersistence, deleteDoc, writeBatch, limit, deleteUser, EmailAuthProvider, reauthenticateWithCredential } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

// === LISTAS E DADOS ===
const UNIDADES_CBMMA_FIXAS = [
    "1 BBM", "2 BBM", "1 CIEBM", "10 BBM", "13 BBM", "16 CIBM", "BBS", "BBA", "BMUS","BBEM","BBMAR", "CGCS", "DEP", "DAT", "DP", "DF", "DPM", "DAL", "CPP", "CPO", 
    "1 SEÇÃO", "2 SEÇÃO", "3 SEÇÃO", "4 SEÇÃO", "5 SEÇÃO", "CAPS", "CRF", "CEPDECMA", "DER", 
    "CMCB I", "CMCB II - SJR", "CMCB XII - PAÇO", "CMCB XIII - GUANABARA", "CMCB XXVI - PIO XII", 
    "ABMJM", "GAB CMT GERAL", "GAB CMT ADJUNTO", "COCB-M"
];
const FUNCOES_TATICAS = [
    "SOCORRISTA", "MOTORISTA (AR/ABT)", "COORD. DO EVENTO", "SUBCOORD. DO EVENTO", 
    "OF. DE LOGÍSTICA", "AUX. DE LOGÍSTICA", "OF DE OPERAÇÕES", 
    "OF. DE ADM/FINANÇAS", "AUX. DE ADM/FINANÇAS", "OPER. DRONE", "OF. DE INFORMAÇÃO", 
    "MOP 1 (MOTO)", "MOP 2 (MOTO)", "AR (MOTORISTA)", "ABT", "VAN", "MICRO - ÔNIBUS", "BOMBEIRO MILITAR"
];

let usuarioAtual = null;
let perfilAtual = null;
let escalaSelecionadaId = null;
let eventoPreviewAtual = null; 
let listaOrdensTemporaria = [];
let dadosParaEnvio = null;
let idEdicaoAdmin = null; 
let demandasEdicaoCache = []; 

// === UTILITÁRIOS & SEGURANÇA ===

function escapar(str) {
    if (!str) return "";
    return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
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
    const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const sections = [8, 4, 4, 4, 12]; 
    let codigo = '';
    sections.forEach((len, index) => {
        for(let i=0; i<len; i++) codigo += chars.charAt(Math.floor(Math.random() * chars.length));
        if(index < sections.length - 1) codigo += '-';
    });
    return codigo;
}

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

// === AUXILIAR DE CATEGORIA PARA O SPLIT INTELIGENTE ===
function getCategoriaPatente(posto) {
    posto = posto.toUpperCase();
    if (['CEL', 'TC', 'TEN CEL', 'MAJ'].some(p => posto.includes(p))) return 'superior';
    if (['CAP'].some(p => posto.includes(p))) return 'intermediario';
    if (['TEN', 'ASP'].some(p => posto.includes(p))) return 'subalterno';
    if (['CAD', 'AL', 'ST', 'SUB'].some(p => posto.includes(p))) return 'especial';
    return 'praca'; // Default (SGT, CB, SD)
}

// === INICIALIZAÇÃO ===
document.addEventListener('DOMContentLoaded', () => {
    popularSelectCadastroEFuncoes();
});

function popularSelectCadastroEFuncoes() {
    // 1. Popula o select da tela de Cadastro
    const selCadastro = document.getElementById('unidade-cadastro');
    if(selCadastro && selCadastro.options.length <= 1) {
        selCadastro.innerHTML = "<option value=''>Selecione a Unidade...</option>";
        UNIDADES_CBMMA_FIXAS.forEach(u => selCadastro.innerHTML += `<option value="${u}">${u}</option>`);
    }

    // 2. NOVO: Popula o select de "Unidade Destino" no painel do Admin
    const selDestinoAdmin = document.getElementById('select-unidade');
    if(selDestinoAdmin && selDestinoAdmin.options.length <= 1) {
        selDestinoAdmin.innerHTML = "<option value=''>Selecione a Unidade...</option>";
        UNIDADES_CBMMA_FIXAS.forEach(u => selDestinoAdmin.innerHTML += `<option value="${u}">${u}</option>`);
    }

    // 3. Popula os selects de Funções
    const selectsFuncao = [document.getElementById('select-funcao'), document.getElementById('edit-admin-funcao')];
    selectsFuncao.forEach(sel => {
        if(sel) {
            sel.innerHTML = "";
            FUNCOES_TATICAS.forEach(f => sel.innerHTML += `<option value="${f}">${f}</option>`);
        }
    });
}

async function carregarUnidadesCadastradasNoAdmin() {
    // ... (mesma função anterior)
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
    if (sup==0 && int==0 && sub==0 && esp==0 && pra==0) return alert("Defina a quantidade de militares.");

    const itemExistente = listaOrdensTemporaria.find(item => item.unidade === unidade);

    const novaDemanda = {
        funcao,
        cota: { superior: sup, intermediario: int, subalterno: sub, especial: esp, praca: pra }
    };

    if (itemExistente) {
        const demandaExistente = itemExistente.demandas.find(d => d.funcao === funcao);
        if (demandaExistente) {
            // SOMA AS COTAS SE JÁ EXISTIR
            demandaExistente.cota.superior = parseInt(demandaExistente.cota.superior) + parseInt(sup);
            demandaExistente.cota.intermediario = parseInt(demandaExistente.cota.intermediario) + parseInt(int);
            demandaExistente.cota.subalterno = parseInt(demandaExistente.cota.subalterno) + parseInt(sub);
            demandaExistente.cota.especial = parseInt(demandaExistente.cota.especial) + parseInt(esp);
            demandaExistente.cota.praca = parseInt(demandaExistente.cota.praca) + parseInt(pra);
        } else {
            itemExistente.demandas.push(novaDemanda);
        }
    } else {
        listaOrdensTemporaria.push({ 
            id: Date.now(), 
            unidade, 
            demandas: [novaDemanda] 
        });
    }
    
    atualizarTabelaOrdens();
    
    ['input-sup','input-int','input-sub','input-esp','input-pra'].forEach(id => document.getElementById(id).value = '');
}

function atualizarTabelaOrdens() {
    const corpo = document.getElementById('tabela-ordens-body');
    document.getElementById('contador-ordens').innerText = `${listaOrdensTemporaria.length}`;
    corpo.innerHTML = "";
    
    listaOrdensTemporaria.forEach((item, index) => {
        let resumoHTML = "";
        item.demandas.forEach(d => {
            const c = d.cota;
            let qtds = [];
            if(c.superior > 0) qtds.push(`${c.superior} SUP`);
            if(c.intermediario > 0) qtds.push(`${c.intermediario} INT`);
            if(c.subalterno > 0) qtds.push(`${c.subalterno} SUB`);
            if(c.especial > 0) qtds.push(`${c.especial} ESP`);
            if(c.praca > 0) qtds.push(`${c.praca} PÇ`);
            resumoHTML += `<div class="mb-1"><span class="fw-bold small text-dark">${d.funcao}:</span> <span class="text-muted small">${qtds.join(', ')}</span></div>`;
        });

        corpo.innerHTML += `
            <tr class="border-bottom">
                <td class="fw-bold align-middle">${escapar(item.unidade)}</td>
                <td colspan="2">${resumoHTML}</td>
                <td class="text-end align-middle"><button onclick="window.app.excluirOrdem(${index})" class="btn btn-sm text-danger ios-click"><i class="bi bi-x-circle-fill"></i></button></td>
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
                funcao: "MULTIPLAS",
                demandas: ordem.demandas, 
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
        if (grupos.size === 0) { lista.innerHTML = "<div class='text-muted text-center py-3'>Histórico vazio.</div>"; return; }

        const gruposArray = Array.from(grupos.values()).sort((a, b) => new Date(b.data) - new Date(a.data));
        
        gruposArray.slice(0, 50).forEach(info => {
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
    } catch(e) { console.error(e); }
}

// ================= ADMIN PREVIEW & EDICAO =================
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
            
            let textoCota = "";
            let jsonDemandas = "[]";

            if(d.demandas) {
                jsonDemandas = JSON.stringify(d.demandas).replace(/"/g, "&quot;");
                d.demandas.forEach(dem => {
                   const c = dem.cota;
                   let parts = [];
                   if(c.superior > 0) parts.push(c.superior); if(c.intermediario > 0) parts.push(c.intermediario); if(c.subalterno > 0) parts.push(c.subalterno); if(c.especial > 0) parts.push(c.especial); if(c.praca > 0) parts.push(c.praca);
                   if(parts.reduce((a,b)=>parseInt(a)+parseInt(b),0) > 0) {
                        textoCota += `${dem.funcao} (${parts.reduce((a,b)=>parseInt(a)+parseInt(b),0)})<br>`;
                   }
                });
            }
            
            const btnEdit = `<button onclick="window.app.editarSolicitacaoAdmin('${idDoc}', '${escapar(d.unidade)}', '${jsonDemandas}', '${d.prazoData}', '${d.prazoHora}')" class="btn btn-sm btn-outline-primary border-0 me-1" title="Editar"><i class="bi bi-pencil-square"></i></button>`;
            const btnDelete = `<button onclick="window.app.excluirEscalaIndividual('${idDoc}', '${escapar(d.unidade)}')" class="btn btn-sm btn-outline-danger border-0" title="Excluir"><i class="bi bi-trash-fill"></i></button>`;

            if(d.status === "Pendente") {
                html += `
                <tr class="table-danger border-bottom">
                    <td class="text-center fw-bold text-muted">-</td>
                    <td colspan="3" class="small text-danger fw-bold align-middle">
                        <i class="bi bi-exclamation-circle-fill me-1"></i> PENDENTE: ${escapar(d.unidade)}
                        <br><span class="text-muted fw-normal ms-3" style="font-size:0.7em">${textoCota}</span>
                    </td>
                    <td class="align-middle">VARIAS</td>
                    <td class="text-end align-middle">${btnEdit}${btnDelete}</td>
                </tr>`;
            } else {
                militares.forEach((m, index) => {
                    const funcaoExibida = m.funcaoIndividual ? m.funcaoIndividual : "DEFINIDO";
                    html += `<tr>
                        <td class="fw-bold text-center text-muted">${index + 1}</td>
                        <td><span class="fw-bold">${escapar(m.posto)}</span> ${escapar(m.guerra)}</td>
                        <td class="small text-muted">${escapar(m.contato)}</td>
                        <td class="fw-bold text-dark">${escapar(d.unidade)}</td>
                        <td><span class="badge bg-light text-dark border">${escapar(funcaoExibida)}</span></td>
                        <td class="text-end">${index === 0 ? btnEdit + btnDelete : ''}</td>
                    </tr>`;
                });
            }
        });
        corpo.innerHTML = html;
    } catch(e) { console.error(e); corpo.innerHTML = "<tr><td colspan='6'>Erro ao carregar ou Sem Permissão.</td></tr>"; }
}

export function editarSolicitacaoAdmin(id, unidade, jsonDemandas, pData, pHora) {
    idEdicaoAdmin = id;
    let demandas = [];
    try { demandas = JSON.parse(jsonDemandas); } catch(e) {}
    if(!demandas || demandas.length === 0) demandas = [];
    demandasEdicaoCache = demandas;

    document.getElementById('edit-admin-subtitle').innerText = `${unidade}`;
    const container = document.getElementById('container-editar-cotas');
    container.innerHTML = "";
    
    demandas.forEach((dem, index) => {
        const c = dem.cota;
        container.innerHTML += `
            <div class="card p-3 mb-2 shadow-sm border">
                <h6 class="fw-bold text-danger text-uppercase mb-2">${dem.funcao}</h6>
                <div class="cota-grid">
                     <div><label class="form-label-custom text-center">Sup</label><input type="number" class="form-control text-center fw-bold px-1 input-cota-edit" data-index="${index}" data-tipo="superior" value="${c.superior||0}"></div>
                     <div><label class="form-label-custom text-center">Int</label><input type="number" class="form-control text-center fw-bold px-1 input-cota-edit" data-index="${index}" data-tipo="intermediario" value="${c.intermediario||0}"></div>
                     <div><label class="form-label-custom text-center">Sub</label><input type="number" class="form-control text-center fw-bold px-1 input-cota-edit" data-index="${index}" data-tipo="subalterno" value="${c.subalterno||0}"></div>
                     <div style="grid-column: span 1.5;"><label class="form-label-custom text-center">Esp</label><input type="number" class="form-control text-center fw-bold px-1 input-cota-edit" data-index="${index}" data-tipo="especial" value="${c.especial||0}"></div>
                     <div style="grid-column: span 1.5;"><label class="form-label-custom text-center">Pç</label><input type="number" class="form-control text-center fw-bold px-1 input-cota-edit" data-index="${index}" data-tipo="praca" value="${c.praca||0}"></div>
                </div>
            </div>
        `;
    });

    document.getElementById('edit-admin-prazo-data').value = pData || '';
    document.getElementById('edit-admin-prazo-hora').value = pHora || '23:59';
    document.getElementById('modal-editar-admin').classList.add('active');
}

export async function salvarEdicaoAdmin() {
    if(!idEdicaoAdmin) return;
    const inputs = document.querySelectorAll('.input-cota-edit');
    inputs.forEach(input => {
        const index = input.getAttribute('data-index');
        const tipo = input.getAttribute('data-tipo');
        if(demandasEdicaoCache[index] && demandasEdicaoCache[index].cota) {
            demandasEdicaoCache[index].cota[tipo] = input.value;
        }
    });

    const novoPrazoData = document.getElementById('edit-admin-prazo-data').value;
    const novoPrazoHora = document.getElementById('edit-admin-prazo-hora').value;

    try {
        await updateDoc(doc(db, "escalas", idEdicaoAdmin), {
            demandas: demandasEdicaoCache,
            prazoData: novoPrazoData,
            prazoHora: novoPrazoHora
        });
        alert("Cotas e Prazo atualizados!");
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
        const batch = writeBatch(db);
        snapshot.forEach(d => batch.delete(d.ref));
        await batch.commit(); 
        alert("Histórico apagado.");
        carregarEventosAdmin();
    } catch(e) { console.error(e); alert("Erro ao excluir: " + e.message); carregarEventosAdmin(); }
}

// ================= ESCALANTE: PREENCHIMENTO =================
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
    let totalVagas = 0;
    if(d.demandas) {
        d.demandas.forEach(dem => {
             const c = dem.cota;
             totalVagas += (parseInt(c.superior)||0) + (parseInt(c.intermediario)||0) + (parseInt(c.subalterno)||0) + (parseInt(c.especial)||0) + (parseInt(c.praca)||0);
        });
    }

    let isBloqueado = false;
    let textoPrazo = "";
    let btnClass = isPendente ? "btn-tactical" : "btn-outline-success";
    let btnText = isPendente ? "RESPONDER AGORA" : "VER / EDITAR";
    let cardOpacity = isPendente ? "1" : "0.85";
    
    if (d.prazoData) {
        const pAno = parseInt(d.prazoData.split('-')[0]);
        const pMes = parseInt(d.prazoData.split('-')[1]) - 1; 
        const pDia = parseInt(d.prazoData.split('-')[2]);
        const pHora = parseInt((d.prazoHora || '23:59').split(':')[0]);
        const pMin = parseInt((d.prazoHora || '23:59').split(':')[1]);
        const dataLimite = new Date(pAno, pMes, pDia, pHora, pMin, 59);
        
        if (new Date() > dataLimite) {
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
                    <strong class="d-block text-primary">VÁRIAS FUNÇÕES</strong>
                    <div class="small text-muted">${totalVagas} Militares Requisitados</div>
                </div>
                ${textoPrazo}
                <button onclick="window.app.abrirEdicao('${d.id}')" class="btn ${btnClass} w-100 fw-bold mt-auto py-3 rounded-3 shadow-sm ios-click" ${isBloqueado ? 'disabled' : ''}>${btnText}</button>
            </div>
        </div>`;
}

export async function abrirEdicao(id) {
    escalaSelecionadaId = id;
    const docSnap = await getDoc(doc(db, "escalas", id));
    const d = docSnap.data();
    
    document.getElementById('titulo-evento-form').innerText = d.evento;
    document.getElementById('subtitulo-form').innerText = d.status === "Preenchido" ? "Edição de Militares" : "Preencha conforme as funções abaixo";
    
    const container = document.getElementById('container-inputs-militares');
    container.innerHTML = "";
    
    let dadosSalvos = [];
    try { dadosSalvos = JSON.parse(d.militares); } catch {}
    
    let contadorGeral = 0;

    if (d.status === "Preenchido") {
        // MODO EDIÇÃO: Mostra apenas os militares que já foram enviados, sem gerar novas caixas em branco
        if(dadosSalvos.length === 0) container.innerHTML = "<p class='text-center text-muted'>Nenhum militar registrado.</p>";
        
        const funcoes = [...new Set(dadosSalvos.map(m => m.funcaoIndividual))];
        funcoes.forEach(func => {
            container.innerHTML += `<div class="w-100 text-center bg-dark text-white p-2 rounded fw-bold mt-3 mb-2 text-uppercase">${func}</div>`;
            const mils = dadosSalvos.filter(m => m.funcaoIndividual === func);
            mils.forEach((m, i) => {
                container.innerHTML += gerarHtmlMilitar(i, 'MILITAR ENVIADO', m, func);
            });
        });
    } else if (d.demandas) {
        // MODO PENDENTE: Mostra as vagas baseadas nas demandas
        d.demandas.forEach(demanda => {
            const c = demanda.cota;
            const nomeFuncao = demanda.funcao;
            
            container.innerHTML += `<div class="w-100 text-center bg-dark text-white p-2 rounded fw-bold mt-3 mb-2 text-uppercase">${nomeFuncao}</div>`;
            
            const gerarLoop = (qtd, rotulo) => {
                const num = parseInt(qtd) || 0;
                for(let i=0; i < num; i++) {
                    const dadosMilitar = dadosSalvos[contadorGeral++] || {};
                    container.innerHTML += gerarHtmlMilitar(i, rotulo, dadosMilitar, nomeFuncao);
                }
            };

            if(c.oficial) gerarLoop(c.oficial, 'OFICIAL');
            gerarLoop(c.superior, 'OF. SUPERIOR');
            gerarLoop(c.intermediario, 'OF. INTERMEDIÁRIO');
            gerarLoop(c.subalterno, 'OF. SUBALTERNO');
            gerarLoop(c.especial, 'PRAÇA ESPECIAL');
            if(c.praca) gerarLoop(c.praca, 'PRAÇA');
        });
    } 

    document.getElementById('form-militar-modal').classList.add('active'); 
}
// ALTERADO: Adicionados exemplos específicos por quadro/patente
function gerarHtmlMilitar(index, tipo, dados, funcaoFixa) {
    let exPosto = "POSTO/GRAD";
    let subTexto = "";
    
    // Define os textos com base no tipo exato da cota
    if(tipo.includes("SUPERIOR")) { 
        exPosto = "POSTO"; 
        subTexto = "Ex: CEL QOC, TEN CEL QOC, MAJ QOC/QOA/QOE"; 
    }
    else if(tipo.includes("INTERMEDIÁRIO")) { 
        exPosto = "POSTO"; 
        subTexto = "Ex: CAP QOC/QOA/QOE"; 
    }
    else if(tipo.includes("SUBALTERNO")) { 
        exPosto = "POSTO"; 
        subTexto = "Ex: 1 TEN QOC/QOA/QOE, 2 TEN QOC/QOA/QOE"; 
    }
    else if(tipo.includes("ESPECIAL")) { 
        exPosto = "POSTO/GRAD"; 
        subTexto = "Ex: ASP OF, CAD BM/3, CAD BM/2, CAD BM/1"; 
    }
    else if(tipo.includes("PRAÇA")) { 
        exPosto = "GRADUAÇÃO"; 
        subTexto = "Ex: ST, 1 SGT, 2 SGT, 3 SGT, CB, SD"; 
    }

    return `
    <div class="p-3 bg-white rounded-3 border mb-3 militar-row shadow-sm" data-funcao="${funcaoFixa}">
        <span class="badge bg-secondary mb-2">${tipo} ${index + 1}</span>
        <div class="row g-2">
            <div class="col-4 col-md-3">
                <input type="text" class="form-control campo-posto fw-bold" placeholder="${exPosto}" value="${escapar(dados.posto || '')}" oninput="this.value = this.value.toUpperCase()">
                <div class="form-text text-muted small fst-italic" style="font-size: 0.65rem; margin-top: 2px;">${subTexto}</div>
            </div>
            <div class="col-8 col-md-5">
                <input type="text" class="form-control campo-nome" placeholder="NOME COMPLETO" value="${escapar(dados.nome || '')}" oninput="this.value = this.value.toUpperCase()">
            </div>
            <div class="col-6 col-md-4">
                <input type="text" class="form-control campo-guerra fw-bold text-uppercase" placeholder="NOME DE GUERRA" value="${escapar(dados.guerra || '')}" oninput="this.value = this.value.toUpperCase()">
            </div>
            <div class="col-6 col-md-12">
                <input type="text" class="form-control campo-tel" placeholder="TEL: 98 9XXXX-XXXX" value="${escapar(dados.contato || '')}" maxlength="15" oninput="window.formatarTelefoneInput(this)">
            </div>
        </div>
    </div>`;
}
// ================= ENVIO INTELIGENTE (DIVISÃO DE PENDÊNCIAS) =================
export function abrirPreviaRecibo() {
    const rows = document.querySelectorAll('.militar-row');
    let lista = [];
    
    rows.forEach(row => {
        const posto = row.querySelector('.campo-posto').value.trim().toUpperCase();
        const nome = row.querySelector('.campo-nome').value.trim().toUpperCase();
        const guerra = row.querySelector('.campo-guerra').value.trim().toUpperCase();
        const contato = row.querySelector('.campo-tel').value.trim(); 
        const funcaoIndividual = row.getAttribute('data-funcao');
        
        // SÓ ADICIONA SE ESTIVER PREENCHIDO
        if(posto && nome && guerra && contato.length >= 8) {
            lista.push({ posto, nome, guerra, contato, funcaoIndividual });
        }
    });

    if(lista.length === 0) return alert("Preencha pelo menos um militar para enviar.");

    dadosParaEnvio = lista;
    document.getElementById('recibo-evento').innerText = document.getElementById('titulo-evento-form').innerText;
    document.getElementById('recibo-unidade').innerText = perfilAtual.unidade;
    const tbody = document.getElementById('recibo-lista-corpo');
    tbody.innerHTML = "";
    lista.forEach(m => tbody.innerHTML += `<tr><td>${escapar(m.posto)}</td><td><strong>${escapar(m.guerra)}</strong></td><td>${escapar(m.nome)}</td><td>${escapar(m.funcaoIndividual)}</td></tr>`);
    document.getElementById('recibo-modal').classList.add('active');
}

export function confirmarEnvioRecibo() { abrirTelaAssinatura(); }
export function abrirTelaAssinatura() {
    if (!escalaSelecionadaId || !dadosParaEnvio) return;
    document.getElementById('modal-assinatura').classList.add('active');
    document.getElementById('assinatura-nome').value = '';
    document.getElementById('assinatura-nome-completo').value = ''; 
    document.getElementById('assinatura-funcao').value = '';
}

export function solicitarConfirmacaoSenha() {
    const nomeGuerra = document.getElementById('assinatura-nome').value.trim();
    const nomeCompleto = document.getElementById('assinatura-nome-completo').value.trim(); 
    const funcao = document.getElementById('assinatura-funcao').value.trim();
    if(nomeGuerra.length < 3 || nomeCompleto.length < 5 || funcao.length < 3) return alert("ATENÇÃO: Preencha todos os campos da assinatura.");
    document.getElementById('modal-assinatura').classList.remove('active');
    document.getElementById('input-senha-assinatura').value = "";
    document.getElementById('modal-confirmar-senha').classList.add('active');
}

export async function validarSenhaEGerarPDF() {
    const senha = document.getElementById('input-senha-assinatura').value;
    const btn = document.querySelector('button[onclick="validarSenhaEGerarPDF()"]');
    if(!senha) return alert("Digite a senha.");
    const textoOriginal = btn.innerHTML;
    btn.innerHTML = "<span class='spinner-border spinner-border-sm me-2'></span>Autenticando..."; btn.disabled = true;

    try {
        const user = auth.currentUser;
        if (!user) throw new Error("Usuário não logado.");
        const credential = EmailAuthProvider.credential(user.email, senha);
        await reauthenticateWithCredential(user, credential);
        document.getElementById('modal-confirmar-senha').classList.remove('active');
        await finalizarEnvioReal();
    } catch (error) { console.error(error); alert("Senha incorreta."); btn.innerHTML = textoOriginal; btn.disabled = false; document.getElementById('input-senha-assinatura').value = ""; }
}

async function finalizarEnvioReal() {
    const nomeGuerra = document.getElementById('assinatura-nome').value.trim().toUpperCase();
    const nomeCompleto = document.getElementById('assinatura-nome-completo').value.trim().toUpperCase(); 
    const funcaoAssinatura = document.getElementById('assinatura-funcao').value.trim().toUpperCase();
    document.getElementById('recibo-modal').classList.remove('active');

    try {
        const docSnap = await getDoc(doc(db, "escalas", escalaSelecionadaId));
        const dadosOriginais = docSnap.data();

        // === CORREÇÃO 1: MODO EDIÇÃO (Impede de clonar o documento) ===
        if (dadosOriginais.status === "Preenchido") {
            const codigoAuth = dadosOriginais.codigoAutenticacao || gerarCodigoAutenticacao();
            await updateDoc(doc(db, "escalas", escalaSelecionadaId), {
                militares: JSON.stringify(dadosParaEnvio),
                assinadoPor: nomeGuerra,
                assinadoNomeCompleto: nomeCompleto,
                assinadoFuncao: funcaoAssinatura,
                dataValidacao: new Date().toISOString()
            });
            
            await setDoc(doc(db, "validacoes_publicas", codigoAuth), {
                codigo: codigoAuth, evento: dadosOriginais.evento, unidade: perfilAtual.unidade, dataValidacao: new Date().toISOString(), status: "Válido"
            }, { merge: true });

            await gerarReciboPDFInstitucional(dadosParaEnvio, codigoAuth, nomeGuerra, nomeCompleto, funcaoAssinatura, dadosOriginais);
            alert("Dados do militar atualizados com sucesso!");
            window.location.reload();
            return;
        }

        // === CORREÇÃO 2: ABATIMENTO INTELIGENTE (Impede de ficar travado como Pendente) ===
        let demandasRestantes = JSON.parse(JSON.stringify(dadosOriginais.demandas || []));

        dadosParaEnvio.forEach(militar => {
            const funcTarget = demandasRestantes.find(d => d.funcao === militar.funcaoIndividual);
            if(funcTarget) {
                const cat = getCategoriaPatente(militar.posto);
                if(funcTarget.cota[cat] > 0) {
                    funcTarget.cota[cat]--;
                } else {
                    // Força o abatimento de qualquer vaga disponível se houver divergência de patente
                    const catDisponivel = Object.keys(funcTarget.cota).find(k => funcTarget.cota[k] > 0);
                    if(catDisponivel) funcTarget.cota[catDisponivel]--;
                }
            }
        });

        let totalSobrou = 0;
        demandasRestantes.forEach(d => {
            totalSobrou += Object.values(d.cota).reduce((a,b)=>parseInt(a||0)+parseInt(b||0),0);
        });

        if (totalSobrou > 0) {
            const codigoAuth = gerarCodigoAutenticacao();
            const dataHoraEnvio = new Date().toISOString(); 
            
            await addDoc(collection(db, "escalas"), {
                ...dadosOriginais,
                militares: JSON.stringify(dadosParaEnvio),
                status: "Preenchido",
                codigoAutenticacao: codigoAuth,
                dataValidacao: dataHoraEnvio,
                assinadoPor: nomeGuerra,
                assinadoNomeCompleto: nomeCompleto,
                assinadoFuncao: funcaoAssinatura,
                criadoEm: new Date()
            });

            await updateDoc(doc(db, "escalas", escalaSelecionadaId), {
                demandas: demandasRestantes
            });

            await setDoc(doc(db, "validacoes_publicas", codigoAuth), {
                codigo: codigoAuth, evento: dadosOriginais.evento, unidade: perfilAtual.unidade, dataValidacao: dataHoraEnvio, status: "Válido"
            });
            await gerarReciboPDFInstitucional(dadosParaEnvio, codigoAuth, nomeGuerra, nomeCompleto, funcaoAssinatura, dadosOriginais);

            alert(`Envio Parcial Realizado!\n\nFoi gerado o recibo dos militares informados.\nA pendência foi atualizada solicitando apenas os ${totalSobrou} restantes.`);

        } else {
            const codigoAuth = gerarCodigoAutenticacao();
            const dataHoraEnvio = new Date().toISOString();
            await updateDoc(doc(db, "escalas", escalaSelecionadaId), { 
                militares: JSON.stringify(dadosParaEnvio), status: "Preenchido", codigoAutenticacao: codigoAuth, dataValidacao: dataHoraEnvio,
                assinadoPor: nomeGuerra, assinadoNomeCompleto: nomeCompleto, assinadoFuncao: funcaoAssinatura
            });
            await setDoc(doc(db, "validacoes_publicas", codigoAuth), {
                codigo: codigoAuth, evento: dadosOriginais.evento, unidade: perfilAtual.unidade, dataValidacao: dataHoraEnvio, status: "Válido"
            });
            await gerarReciboPDFInstitucional(dadosParaEnvio, codigoAuth, nomeGuerra, nomeCompleto, funcaoAssinatura, dadosOriginais);
            alert("Escala enviada com sucesso!");
        }
        
        window.location.reload();

    } catch (e) { alert("Erro no envio: " + e.message); window.location.reload(); }
}
        
        window.location.reload();

    } catch (e) { alert("Erro no envio: " + e.message); window.location.reload(); }
}

export function abrirValidador() {
    document.getElementById('modal-validador').classList.add('active');
    document.getElementById('resultado-validacao').style.display = 'none';
    document.getElementById('input-codigo-validacao').value = "";
}

export async function consultarAutenticidade() {
    const codigo = document.getElementById('input-codigo-validacao').value.trim().toUpperCase();
    const divResult = document.getElementById('resultado-validacao');
    divResult.style.display = 'block';
    divResult.innerHTML = "<div class='text-center text-muted mt-3'><span class='spinner-border text-success'></span><br>Verificando...</div>";

    try {
        const docRef = doc(db, "validacoes_publicas", codigo);
        const docSnap = await getDoc(docRef);

        if (!docSnap.exists()) {
            divResult.innerHTML = `<div class="mt-4 p-3 bg-danger bg-opacity-10 border border-danger rounded-3 text-center"><i class="bi bi-x-circle-fill text-danger display-5"></i><h5 class="fw-bold text-danger mt-2">CÓDIGO NÃO ENCONTRADO</h5><p class="small text-muted mb-0">Verifique se digitou corretamente.</p></div>`;
        } else {
            const d = docSnap.data();
            const dataF = new Date(d.dataValidacao).toLocaleString('pt-BR');
            divResult.innerHTML = `
                <div class="mt-4 bg-success bg-opacity-25 p-3 rounded-4 text-center border border-success">
                    <i class="bi bi-patch-check-fill text-success display-4"></i>
                    <h5 class="fw-bold text-success text-uppercase mt-2 mb-3">DOCUMENTO AUTÊNTICO</h5>
                    <div class="bg-white p-3 rounded-3 text-start small shadow-sm">
                        <div class="mb-1"><span class="fw-bold text-muted">Evento:</span> <span class="text-dark fw-bold text-uppercase">${escapar(d.evento)}</span></div>
                        <div class="mb-1"><span class="fw-bold text-muted">Unidade:</span> <span class="text-dark fw-bold">${escapar(d.unidade)}</span></div>
                        <div class="mb-1"><span class="fw-bold text-muted">Situação:</span> <span class="text-success fw-bold">VÁLIDO</span></div>
                        <div><span class="fw-bold text-muted">Emitido em:</span> <span class="text-dark">${dataF}</span></div>
                    </div>
                    <div class="mt-2 text-muted" style="font-size: 0.65rem;"><i class="bi bi-lock-fill"></i> Dados protegidos pela LGPD.</div>
                </div>`;
        }
    } catch (e) { divResult.innerHTML = "Erro conexão."; }
}

// MANTIDO O GERADOR DE PDF IGUAL
async function gerarReciboPDFInstitucional(listaMilitares, codigoAuth, nomeGuerraAssinatura, nomeCompletoAssinatura, funcaoAssinatura, dadosEscala) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4'); 
    const MARGEM_ESQ = 15; const CENTRO_X = 105; const LARGURA_UTIL = 180;
    const imgBrasaoMA = await carregarImagemBase64('brasao.png'); 

    let y = 10;
    if(imgBrasaoMA) doc.addImage(imgBrasaoMA, 'PNG', 105 - 12, 10, 24, 24);

    doc.setFont("times", "bold"); doc.setFontSize(11); doc.setTextColor(0);
    doc.text("ESTADO DO MARANHÃO", CENTRO_X, y+30, { align: "center" });
    doc.text("SECRETARIA DE SEGURANÇA PÚBLICA", CENTRO_X, y+35, { align: "center" });
    doc.text("CORPO DE BOMBEIROS MILITAR DO MARANHÃO", CENTRO_X, y+40, { align: "center" });
    
    const nomeUnidade = (perfilAtual.unidade || dadosEscala.unidade || "COMANDO OPERACIONAL").toUpperCase();
    doc.text(nomeUnidade, CENTRO_X, y+45, { align: "center" });

    y = 65;
    doc.setFontSize(14); doc.text("ESCALA DE SERVIÇO", CENTRO_X, y, { align: "center" });
    y += 12;

    doc.setFontSize(10); doc.setFont("helvetica", "bold");
    doc.text(`OPERAÇÃO: ${dadosEscala.evento}`, MARGEM_ESQ, y); y += 5;
    
    const dataObj = new Date(dadosEscala.data + "T12:00:00"); 
    const opcoesData = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    doc.text(`DATA: ${dataObj.toLocaleDateString('pt-BR', opcoesData).toUpperCase()}`, MARGEM_ESQ, y); y += 5;
    
    const horario = (dadosEscala.horaInicio && dadosEscala.horaFim) ? `${dadosEscala.horaInicio} ÀS ${dadosEscala.horaFim}` : "A DEFINIR";
    doc.text(`HORÁRIO: ${horario}`, MARGEM_ESQ, y); y += 10;

    doc.setFillColor(230, 230, 230); doc.rect(MARGEM_ESQ, y, LARGURA_UTIL, 8, 'F');
    doc.setDrawColor(0); doc.setLineWidth(0.1); doc.rect(MARGEM_ESQ, y, LARGURA_UTIL, 8);
    
    doc.setFontSize(9);
    doc.text("POSTO/GRAD", MARGEM_ESQ + 2, y + 5);
    doc.text("NOME COMPLETO", MARGEM_ESQ + 35, y + 5);
    doc.text("NOME GUERRA", MARGEM_ESQ + 110, y + 5);
    doc.text("FUNÇÃO", MARGEM_ESQ + 145, y + 5);
    y += 8;

    doc.setFont("helvetica", "normal");

    listaMilitares.forEach((m) => {
        if (y > 250) { doc.addPage(); y = 20; doc.text("(Continuação)", MARGEM_ESQ, y-5); }
        doc.rect(MARGEM_ESQ, y, LARGURA_UTIL, 7);
        doc.setFont("helvetica", "normal");
        doc.text(m.posto, MARGEM_ESQ + 2, y + 5);
        let nomeVisual = m.nome.toUpperCase();
        if(nomeVisual.length > 40) nomeVisual = nomeVisual.substring(0, 38) + "...";
        doc.text(nomeVisual, MARGEM_ESQ + 35, y + 5);
        doc.setFont("helvetica", "bold");
        doc.text(m.guerra, MARGEM_ESQ + 110, y + 5);
        doc.setFont("helvetica", "normal"); 
        const f = m.funcaoIndividual ? m.funcaoIndividual : (dadosEscala.funcao || "BOMBEIRO");
        doc.text(f.substring(0, 18), MARGEM_ESQ + 145, y + 5);
        y += 7;
    });

    y += 15; if (y > 220) { doc.addPage(); y = 40; }

    const qrContainer = document.getElementById('qrcode-container');
    qrContainer.innerHTML = "";
    const urlValidacao = window.location.origin + window.location.pathname + "?validar=" + codigoAuth;
    new QRCode(qrContainer, { text: urlValidacao, width: 150, height: 150, correctLevel: QRCode.CorrectLevel.H });
    await new Promise(r => setTimeout(r, 300));
    
    let qrDataUrl = null;
    const qrCanvas = qrContainer.querySelector('canvas');
    if (qrCanvas) qrDataUrl = qrCanvas.toDataURL("image/png");
    else { const qrImg = qrContainer.querySelector('img'); if (qrImg) qrDataUrl = qrImg.src; }

    doc.setDrawColor(150); doc.setLineWidth(0.5); doc.rect(MARGEM_ESQ, y, LARGURA_UTIL, 35);
    if (qrDataUrl) doc.addImage(qrDataUrl, 'PNG', MARGEM_ESQ + 3, y + 2.5, 30, 30);

    const textoX = MARGEM_ESQ + 38;
    doc.setFont("courier", "bold"); doc.setFontSize(10); doc.setTextColor(0);
    doc.text("DOCUMENTO ASSINADO DIGITALMENTE", textoX, y + 8);
    doc.setFont("helvetica", "bold"); doc.setFontSize(9);
    doc.text(nomeGuerraAssinatura, textoX, y + 14); 
    doc.setFont("helvetica", "normal");
    doc.text(nomeCompletoAssinatura, textoX, y + 18); 
    doc.setFontSize(8);
    doc.text(funcaoAssinatura, textoX, y + 23); 
    doc.setTextColor(80);
    doc.text(`Assinado em: ${new Date().toLocaleString('pt-BR')}`, textoX, y + 28);
    doc.setFont("courier", "normal"); doc.setFontSize(7);
    doc.text(`HASH: ${codigoAuth}`, textoX, y + 33);

    doc.save(`ESCALA_${dadosEscala.unidade}_${dadosEscala.evento}.pdf`);
}
// ================= EXPORTAÇÃO PARA EXCEL =================
export async function baixarExcelDoEvento() {
    if (!eventoPreviewAtual) return alert("Nenhum evento selecionado.");
    
    // Pega o botão e coloca ele em estado de "Carregando"
    const btn = document.querySelector('button[onclick="baixarExcelDoEvento()"]');
    const textoOriginal = btn.innerHTML;
    btn.innerHTML = '<span class="spinner-border spinner-border-sm me-2"></span>Gerando Planilha...';
    btn.disabled = true;

    try {
        // Busca os dados do evento atual no banco de dados
        const q = query(collection(db, "escalas"), where("evento", "==", eventoPreviewAtual.nome), where("data", "==", eventoPreviewAtual.data));
        const snapshot = await getDocs(q);
        
        // Cria a planilha virtual
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Escala');
        
        // Define as colunas do Excel
        sheet.columns = [
            { header: 'ORDEM', key: 'ordem', width: 10 },
            { header: 'POSTO/GRAD', key: 'posto', width: 15 },
            { header: 'NOME DE GUERRA', key: 'guerra', width: 25 },
            { header: 'NOME COMPLETO', key: 'nome', width: 40 },
            { header: 'CONTATO', key: 'contato', width: 20 },
            { header: 'UNIDADE', key: 'unidade', width: 15 },
            { header: 'FUNÇÃO', key: 'funcao', width: 25 }
        ];
        
        // Estiliza o cabeçalho (Fundo vermelho padrão CBM e letras brancas)
        sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB30000' } }; 
        sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

        let contador = 1;
        
        // Percorre as escalas para encontrar quem já preencheu
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            if (d.status !== "Pendente" && d.militares) {
                let militares = [];
                try { militares = JSON.parse(d.militares); } catch(e) {}
                
                // Adiciona cada militar em uma linha da planilha
                militares.forEach(m => {
                    sheet.addRow({
                        ordem: contador++,
                        posto: m.posto,
                        guerra: m.guerra,
                        nome: m.nome,
                        contato: m.contato,
                        unidade: d.unidade,
                        funcao: m.funcaoIndividual || d.funcao || "BOMBEIRO"
                    });
                });
            }
        });

        // Se ninguém tiver respondido ainda
        if (contador === 1) {
            alert("Não há militares confirmados nesta missão para gerar o Excel.");
            btn.innerHTML = textoOriginal;
            btn.disabled = false;
            return;
        }

        // Gera e baixa o arquivo final no computador/celular
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        saveAs(blob, `ESCALA_${eventoPreviewAtual.nome.replace(/\s+/g, '_')}.xlsx`);
        
    } catch (error) {
        console.error("Erro ao gerar Excel:", error);
        alert("Erro ao gerar o arquivo Excel. Verifique a conexão.");
    }
    
    // Restaura o botão ao normal
    btn.innerHTML = textoOriginal;
    btn.disabled = false;
}
window.app = { 
    fazerLogin, fazerCadastro, sair, 
    adicionarOrdem, limparOrdens, excluirOrdem, dispararSolicitacao, 
    abrirPreviaRecibo, confirmarEnvioRecibo, abrirPreview, baixarExcelDoEvento, 
    excluirEscalaIndividual, abrirEdicao, excluirEventoCompleto, 
    editarSolicitacaoAdmin, salvarEdicaoAdmin, abrirValidador, consultarAutenticidade, 
    abrirTelaAssinatura, solicitarConfirmacaoSenha, validarSenhaEGerarPDF      
// Restaura o botão ao normal
    btn.innerHTML = textoOriginal;
    btn.disabled = false;
}

window.app = { 
    fazerLogin, fazerCadastro, sair, 
    adicionarOrdem, limparOrdens, excluirOrdem, dispararSolicitacao, 
    abrirPreviaRecibo, confirmarEnvioRecibo, abrirPreview, baixarExcelDoEvento, 
    excluirEscalaIndividual, abrirEdicao, excluirEventoCompleto, 
    editarSolicitacaoAdmin, salvarEdicaoAdmin, abrirValidador, consultarAutenticidade, 
    abrirTelaAssinatura, solicitarConfirmacaoSenha, validarSenhaEGerarPDF      
};
