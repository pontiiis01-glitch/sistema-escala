import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy, setPersistence, browserSessionPersistence, deleteDoc } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

let usuarioAtual = null;
let perfilAtual = null;
let escalaSelecionadaId = null;
let eventoPreviewAtual = null;
let listaOrdensTemporaria = [];
let dadosParaEnvio = null; // Armazena dados temporariamente para o recibo

// ================= AUTH =================
export async function fazerLogin() {
    const email = document.getElementById('email-login').value;
    const senha = document.getElementById('senha-login').value;
    try { 
        await setPersistence(auth, browserSessionPersistence);
        await signInWithEmailAndPassword(auth, email, senha); 
    } 
    catch (e) { console.error(e); document.getElementById('msg-erro').innerText = "Acesso negado. Verifique credenciais."; }
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
            document.getElementById('login-area-wrapper').style.display = 'none';
            document.getElementById('dashboard-screen').style.display = 'block';
            document.getElementById('titulo-unidade').innerText = perfilAtual.unidade;
            
            if (perfilAtual.funcao === 'admin') {
                document.getElementById('admin-area').style.display = 'block';
                carregarListaUnidades(); carregarEventosAdmin();
            } else {
                document.getElementById('unidade-area').style.display = 'block';
                carregarPendenciasUnidade();
            }
        }
    }
});

// ================= ADMIN =================
async function carregarListaUnidades() {
    const select = document.getElementById('select-unidade');
    try {
        const q = query(collection(db, "usuarios"), where("funcao", "==", "escalante"));
        const snapshot = await getDocs(q);
        select.innerHTML = "<option value='' selected>Selecione a Unidade...</option>";
        let unidades = [];
        snapshot.forEach(doc => unidades.push(doc.data().unidade));
        [...new Set(unidades)].sort().forEach(u => select.innerHTML += `<option value="${u}">${u}</option>`);
    } catch (e) { select.innerHTML = "<option>Erro</option>"; }
}

export function adicionarOrdem() {
    const unidade = document.getElementById('select-unidade').value;
    const funcao = document.getElementById('select-funcao').value;
    const oficiais = document.getElementById('input-oficiais').value;
    const pracas = document.getElementById('input-pracas').value;

    if (!unidade) return alert("Selecione uma unidade!");
    if (oficiais == 0 && pracas == 0) return alert("Defina a quantidade de Oficiais ou Praças.");

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
                <td class="text-end"><button onclick="window.app.excluirOrdem(${index})" class="btn btn-sm text-danger"><i class="bi bi-x-circle-fill"></i></button></td>
            </tr>`;
    });
}

export function excluirOrdem(index) { listaOrdensTemporaria.splice(index, 1); atualizarTabelaOrdens(); }
export function limparOrdens() { listaOrdensTemporaria = []; atualizarTabelaOrdens(); }

export async function dispararSolicitacao() {
    const evento = document.getElementById('nome-evento').value.trim();
    const data = document.getElementById('data-evento').value;
    // Captura os novos campos de horário
    const horaInicio = document.getElementById('hora-inicio').value;
    const horaFim = document.getElementById('hora-fim').value;
    const prazoData = document.getElementById('prazo-data').value;
    const prazoHora = document.getElementById('prazo-hora').value;

    if (!evento || !data) return alert("Preencha Nome e Data do Evento.");
    if (!prazoData) return alert("Defina a Data Limite de resposta.");
    if (listaOrdensTemporaria.length === 0) return alert("Adicione pelo menos uma unidade.");

    try {
        const promises = listaOrdensTemporaria.map(ordem => {
            return addDoc(collection(db, "escalas"), {
                evento, 
                data, 
                horaInicio, 
                horaFim,
                prazoData,
                prazoHora: prazoHora || "23:59", // Default fim do dia se vazio
                unidade: ordem.unidade,
                funcao: ordem.funcao,
                cota: { oficial: ordem.oficiais, praca: ordem.pracas },
                status: "Pendente",
                militares: "[]",
                criadoEm: new Date()
            });
        });
        await Promise.all(promises);
        alert(`Sucesso! ${listaOrdensTemporaria.length} solicitações enviadas.`);
        limparOrdens(); carregarEventosAdmin();
    } catch (e) { alert("Erro: " + e.message); }
}

async function carregarEventosAdmin() {
    const lista = document.getElementById('lista-eventos-admin');
    lista.innerHTML = "<div class='text-center py-3'>Carregando histórico...</div>";
    try {
        // Removido orderBy para evitar erro de índice se não configurado
        const q = query(collection(db, "escalas")); 
        const snapshot = await getDocs(q);
        const grupos = new Map();

        snapshot.forEach(doc => {
            const d = doc.data();
            const chave = `${d.evento}|${d.data}`;
            if (!grupos.has(chave)) grupos.set(chave, { evento: d.evento, data: d.data, total: 0, respondidos: 0 });
            const g = grupos.get(chave);
            g.total++;
            if (d.status === "Preenchido") g.respondidos++;
        });

        lista.innerHTML = "";
        if (grupos.size === 0) lista.innerHTML = "<div class='text-muted text-center'>Sem histórico recente.</div>";

        // Converte Map para Array e ordena manualmente por data (recente primeiro)
        const gruposArray = Array.from(grupos.values()).sort((a, b) => new Date(b.data) - new Date(a.data));

        gruposArray.forEach(info => {
            const dataBr = new Date(info.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const percentual = Math.round((info.respondidos / info.total) * 100);
            const cor = percentual === 100 ? "success" : "warning";
            
            lista.innerHTML += `
                <div class="list-group-item p-3 border-bottom">
                    <div class="d-flex justify-content-between align-items-center mb-2">
                        <div>
                            <strong class="text-dark d-block text-uppercase">${info.evento}</strong>
                            <small class="text-muted fw-bold">${dataBr}</small>
                        </div>
                        <button onclick="window.app.excluirEvento('${info.evento}', '${info.data}')" class="btn btn-sm btn-outline-danger"><i class="bi bi-trash"></i></button>
                    </div>
                    <div class="d-flex justify-content-between small text-muted align-items-center mb-1">
                        <span onclick="window.app.abrirPreview('${info.evento}', '${info.data}')" style="cursor:pointer" class="text-primary fw-bold hover-scale"> <i class="bi bi-eye"></i> Visualizar</span>
                        <span>${info.respondidos}/${info.total} Feitos</span>
                    </div>
                    <div class="progress" style="height: 6px;"><div class="progress-bar bg-${cor}" style="width: ${percentual}%"></div></div>
                </div>`;
        });
    } catch(e) { console.error(e); lista.innerHTML = "<div class='text-danger text-center'>Erro ao carregar (verifique conexão).</div>"; }
}

export async function excluirEvento(evento, data) {
    if(!confirm(`Excluir TODAS as solicitações de "${evento}"?`)) return;
    try {
        const q = query(collection(db, "escalas"), where("evento", "==", evento), where("data", "==", data));
        const snap = await getDocs(q);
        snap.forEach(doc => deleteDoc(doc.ref)); 
        alert("Excluído.");
        carregarEventosAdmin();
    } catch(e) { alert("Erro: " + e.message); }
}

// ================= UNIDADE =================
async function carregarPendenciasUnidade() {
    const lista = document.getElementById('lista-unidade');
    lista.innerHTML = "<div class='text-center w-100'>Atualizando missões...</div>";
    try {
        const q = query(collection(db, "escalas"), where("unidade", "==", perfilAtual.unidade));
        const snapshot = await getDocs(q);
        lista.innerHTML = "";
        
        if (snapshot.empty) return lista.innerHTML = "<div class='text-muted text-center w-100 mt-4 fs-5'>Nenhuma missão pendente.</div>";

        // Ordenação manual
        const docs = [];
        snapshot.forEach(d => docs.push({id: d.id, ...d.data()}));
        docs.sort((a,b) => new Date(a.data) - new Date(b.data));

        docs.forEach(d => {
            const dataEventoFmt = new Date(d.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const isPendente = d.status === "Pendente";
            
            // LÓGICA DE BLOQUEIO DE PRAZO PRECISA
            let isBloqueado = false;
            let textoPrazo = "";
            let textoHorario = "";

            if(d.horaInicio && d.horaFim) textoHorario = `${d.horaInicio} às ${d.horaFim}`;

            if (d.prazoData) {
                // Cria data limite combinando Dia + Hora
                const dataLimiteStr = `${d.prazoData}T${d.prazoHora || '23:59'}:00`;
                const dataLimite = new Date(dataLimiteStr);
                const hoje = new Date();

                if (hoje > dataLimite) {
                    isBloqueado = true;
                    textoPrazo = `<div class="text-danger fw-bold small mt-2"><i class="bi bi-lock-fill"></i> ENCERRADO: ${new Date(d.prazoData).toLocaleDateString('pt-BR', {timeZone: 'UTC'})} às ${d.prazoHora}</div>`;
                } else {
                    textoPrazo = `<div class="text-dark small mt-2 bg-warning bg-opacity-25 p-1 rounded"><i class="bi bi-clock-history"></i> Prazo: ${new Date(d.prazoData).toLocaleDateString('pt-BR', {timeZone: 'UTC'})} às ${d.prazoHora || '23:59'}</div>`;
                }
            }

            const btnClass = isBloqueado ? "btn-secondary disabled" : (isPendente ? "btn-danger" : "btn-outline-success");
            const btnText = isBloqueado ? "PRAZO EXPIRADO" : (isPendente ? "RESPONDER" : "EDITAR ENVIADO");

            lista.innerHTML += `
                <div class="col-md-6 col-lg-4">
                    <div class="bg-white p-4 h-100 rounded-4 shadow-sm border border-light d-flex flex-column">
                        <div class="d-flex justify-content-between mb-2">
                            <span class="badge bg-dark">${dataEventoFmt}</span>
                            <span class="badge ${isPendente ? 'bg-warning text-dark' : 'bg-success'}">${d.status}</span>
                        </div>
                        <h5 class="fw-bold mb-0 text-dark text-uppercase">${d.evento}</h5>
                        <small class="text-muted mb-2 d-block">${textoHorario}</small>
                        
                        <div class="bg-light p-3 rounded border text-center my-2">
                            <strong class="d-block text-primary">${d.funcao}</strong>
                            <div class="small text-muted">Solicitado: ${d.cota.oficial} Of / ${d.cota.praca} Pç</div>
                        </div>

                        ${textoPrazo}

                        <button onclick="window.app.abrirEdicao('${d.id}')" class="btn ${btnClass} w-100 fw-bold mt-auto py-2" ${isBloqueado ? 'disabled' : ''}>
                            ${btnText}
                        </button>
                    </div>
                </div>`;
        });
    } catch(e) { console.error(e); lista.innerHTML = "Erro ao carregar dados."; }
}

export async function abrirEdicao(id) {
    escalaSelecionadaId = id;
    const docSnap = await getDoc(doc(db, "escalas", id));
    const d = docSnap.data();
    
    document.getElementById('titulo-evento-form').innerText = d.evento;
    document.getElementById('subtitulo-form').innerText = `${d.funcao} | Cota: ${d.cota.oficial} Of / ${d.cota.praca} Pç`;
    
    const container = document.getElementById('container-inputs-militares');
    container.innerHTML = "";

    let dadosSalvos = [];
    try { dadosSalvos = JSON.parse(d.militares); } catch {}

    let contadorGlobal = 0;
    for(let i=0; i < parseInt(d.cota.oficial); i++) {
        container.innerHTML += gerarHtmlMilitar(i, 'OFICIAL', dadosSalvos[contadorGlobal] || {});
        contadorGlobal++;
    }
    for(let i=0; i < parseInt(d.cota.praca); i++) {
        container.innerHTML += gerarHtmlMilitar(i, 'PRAÇA', dadosSalvos[contadorGlobal] || {});
        contadorGlobal++;
    }

    document.getElementById('form-militar').style.display = 'block';
    document.getElementById('form-militar').scrollIntoView({ behavior: 'smooth' });
}

function gerarHtmlMilitar(index, tipo, dados) {
    return `
    <div class="p-3 bg-light rounded-3 border mb-3 militar-row">
        <span class="badge bg-secondary mb-2">${tipo} ${index + 1}</span>
        <div class="row g-2">
            <div class="col-4 col-md-3">
                <input type="text" class="form-control campo-posto" placeholder="Posto/Grad" value="${dados.posto || ''}">
            </div>
            <div class="col-8 col-md-5">
                <input type="text" class="form-control campo-nome" placeholder="Nome Completo" value="${dados.nome || ''}">
            </div>
            <div class="col-6 col-md-4">
                <input type="text" class="form-control campo-guerra fw-bold" placeholder="Nome Guerra" value="${dados.guerra || ''}">
            </div>
            <div class="col-6 col-md-12">
                <input type="text" class="form-control campo-tel" placeholder="Telefone (Zap)" value="${dados.contato || ''}">
            </div>
        </div>
    </div>`;
}

// ================= RECIBO E PDF PROFISSIONAL =================
export function abrirPreviaRecibo() {
    const rows = document.querySelectorAll('.militar-row');
    let lista = [];
    let valido = true;

    rows.forEach(row => {
        const posto = row.querySelector('.campo-posto').value.trim();
        const nome = row.querySelector('.campo-nome').value.trim();
        const guerra = row.querySelector('.campo-guerra').value.trim();
        const contato = row.querySelector('.campo-tel').value.trim();
        if(!posto || !nome || !guerra) valido = false;
        lista.push({ posto, nome, guerra, contato });
    });

    if(!valido) return alert("Preencha todos os campos obrigatórios (Posto, Nome, Guerra).");

    dadosParaEnvio = lista; // Guarda para o envio final
    
    // Preenche Modal
    document.getElementById('recibo-evento').innerText = document.getElementById('titulo-evento-form').innerText;
    document.getElementById('recibo-unidade').innerText = perfilAtual.unidade;
    
    const tbody = document.getElementById('recibo-lista-corpo');
    tbody.innerHTML = "";
    lista.forEach(m => {
        tbody.innerHTML += `<tr><td>${m.posto}</td><td><strong>${m.guerra}</strong></td><td>${m.nome}</td><td>${m.contato}</td></tr>`;
    });

    const modal = document.getElementById('recibo-modal');
    modal.classList.remove('d-none'); modal.classList.add('d-flex');
}

export async function confirmarEnvioRecibo() {
    if (!escalaSelecionadaId || !dadosParaEnvio) return;
    
    try {
        const jsonString = JSON.stringify(dadosParaEnvio);
        await updateDoc(doc(db, "escalas", escalaSelecionadaId), { militares: jsonString, status: "Preenchido" });
        
        // Gera o PDF Profissional
        gerarReciboPDFProfissional(dadosParaEnvio);
        
        alert("Enviado com sucesso! O recibo foi baixado.");
        document.getElementById('recibo-modal').classList.replace('d-flex', 'd-none');
        document.getElementById('form-militar').style.display = 'none';
        carregarPendenciasUnidade();
    } catch (e) { alert("Erro ao salvar: " + e.message); }
}

function gerarReciboPDFProfissional(listaMilitares) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    const tituloEvento = document.getElementById('titulo-evento-form').innerText;

    // --- CONFIGURAÇÃO VISUAL ---
    doc.setLineWidth(0.5);
    doc.rect(5, 5, 200, 287); // Borda da Página

    // Cabeçalho
    doc.setFillColor(180, 0, 0); // Vermelho Escuro
    doc.rect(5, 5, 200, 25, 'F');
    
    doc.setTextColor(255, 255, 255);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("RECIBO DE ENVIO DE ESCALA - CBMMA", 105, 18, null, null, "center");
    doc.setFontSize(10);
    doc.text("COMANDO OPERACIONAL METROPOLITANO", 105, 24, null, null, "center");

    // Informações
    doc.setTextColor(0, 0, 0);
    doc.setFontSize(11);
    doc.text(`UNIDADE: ${perfilAtual.unidade}`, 15, 40);
    doc.text(`EVENTO: ${tituloEvento}`, 15, 46);
    doc.text(`DATA DO ENVIO: ${new Date().toLocaleString()}`, 15, 52);
    
    doc.line(15, 55, 195, 55); // Linha separadora

    // Tabela Header
    let y = 65;
    doc.setFillColor(230, 230, 230);
    doc.rect(15, 60, 180, 8, 'F');
    doc.setFontSize(9);
    doc.setFont("helvetica", "bold");
    doc.text("POSTO", 20, 65);
    doc.text("NOME DE GUERRA", 50, 65);
    doc.text("NOME COMPLETO", 90, 65);
    doc.text("CONTATO", 160, 65);

    // Tabela Corpo
    doc.setFont("helvetica", "normal");
    listaMilitares.forEach((m, i) => {
        if(i % 2 === 0) { doc.setFillColor(245, 245, 245); doc.rect(15, y-4, 180, 7, 'F'); } // Zebra
        
        doc.text(m.posto, 20, y);
        doc.setFont("helvetica", "bold");
        doc.text(m.guerra.toUpperCase(), 50, y);
        doc.setFont("helvetica", "normal");
        
        // Trunca nome se for muito longo
        let nomeFmt = m.nome.length > 35 ? m.nome.substring(0,35)+"..." : m.nome;
        doc.text(nomeFmt, 90, y);
        
        doc.text(m.contato, 160, y);
        y += 7;
    });

    // Rodapé
    doc.setFontSize(8);
    doc.setTextColor(100, 100, 100);
    doc.text("Este documento serve como comprovante de que a unidade enviou o efetivo solicitado via sistema.", 105, 280, null, null, "center");

    doc.save(`Recibo_${perfilAtual.unidade}_${Date.now()}.pdf`);
}

// ================= PREVIEW ADMIN =================
export async function abrirPreview(nomeEvento, dataEvento) {
    eventoPreviewAtual = { nome: nomeEvento, data: dataEvento };
    const modal = document.getElementById('preview-modal');
    modal.classList.remove('d-none'); modal.classList.add('d-flex');
    
    document.getElementById('preview-titulo').innerText = nomeEvento;
    const corpo = document.getElementById('tabela-preview-corpo');
    corpo.innerHTML = "<tr><td colspan='6' class='text-center py-4'>Carregando escala unificada...</td></tr>";

    try {
        const q = query(collection(db, "escalas"), where("evento", "==", nomeEvento), where("data", "==", dataEvento));
        const snapshot = await getDocs(q);
        let html = "";
        let ordemGlobal = 1;
        
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            let militares = [];
            try { militares = JSON.parse(d.militares); } catch(e) { militares = []; }

            if(d.status === "Pendente") {
                html += `<tr class="table-danger"><td colspan="6" class="text-center small text-danger fw-bold">PENDENTE: ${d.unidade} - ${d.funcao}</td></tr>`;
            } else {
                militares.forEach(m => {
                    html += `<tr>
                        <td class="fw-bold text-center">${ordemGlobal++}</td>
                        <td>${m.posto}</td>
                        <td><strong>${m.guerra}</strong></td>
                        <td>${m.contato}</td>
                        <td>${d.unidade}</td>
                        <td><span class="badge bg-light text-dark border">${d.funcao}</span></td>
                    </tr>`;
                });
            }
        });
        corpo.innerHTML = html;
    } catch(e) { corpo.innerHTML = "<tr><td colspan='6' class='text-danger text-center'>Erro ao carregar.</td></tr>"; }
}

export async function baixarExcelDoEvento() {
    if (!eventoPreviewAtual) return;
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala');
        const q = query(collection(db, "escalas"), where("evento", "==", eventoPreviewAtual.nome), where("data", "==", eventoPreviewAtual.data), where("status", "==", "Preenchido"));
        const snapshot = await getDocs(q);

        worksheet.columns = [
            { header: 'ORD', key: 'ord', width: 6 },
            { header: 'POSTO', key: 'posto', width: 12 },
            { header: 'NOME DE GUERRA', key: 'guerra', width: 25 },
            { header: 'NOME COMPLETO', key: 'nome', width: 45 },
            { header: 'CONTATO', key: 'contato', width: 18 },
            { header: 'UNIDADE', key: 'unidade', width: 15 },
            { header: 'FUNÇÃO', key: 'funcao', width: 20 }
        ];

        let contador = 1;
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            let militares = [];
            try { militares = JSON.parse(d.militares); } catch { return; }

            militares.forEach(m => {
                worksheet.addRow({
                    ord: contador++, posto: m.posto, guerra: m.guerra.toUpperCase(), nome: m.nome.toUpperCase(), contato: m.contato, unidade: d.unidade, funcao: d.funcao
                });
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), `${eventoPreviewAtual.nome}.xlsx`);
    } catch (e) { alert("Erro: " + e.message); }
}

window.app = { fazerLogin, fazerCadastro, sair, adicionarOrdem, limparOrdens, excluirOrdem, dispararSolicitacao, abrirPreviaRecibo, confirmarEnvioRecibo, abrirPreview, baixarExcelDoEvento, excluirEvento, abrirEdicao };