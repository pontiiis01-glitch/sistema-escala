import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy, setPersistence, browserSessionPersistence, deleteDoc } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";
// jsPDF já carregado via CDN no HTML globalmente como window.jspdf

let usuarioAtual = null;
let perfilAtual = null;
let escalaSelecionadaId = null;
let eventoPreviewAtual = null;
let listaOrdensTemporaria = [];

// ================= AUTH =================
export async function fazerLogin() {
    const email = document.getElementById('email-login').value;
    const senha = document.getElementById('senha-login').value;
    try { 
        await setPersistence(auth, browserSessionPersistence);
        await signInWithEmailAndPassword(auth, email, senha); 
    } 
    catch (e) { console.error(e); document.getElementById('msg-erro').innerText = "Credenciais inválidas."; }
}

export async function fazerCadastro() {
    const email = document.getElementById('email-cadastro').value;
    const senha = document.getElementById('senha-cadastro').value;
    const unidade = document.getElementById('unidade-cadastro').value;
    if(!email || !senha || !unidade) return alert("Preencha todos os campos.");
    try {
        const cred = await createUserWithEmailAndPassword(auth, email, senha);
        await setDoc(doc(db, "usuarios", cred.user.uid), { email, unidade: unidade.toUpperCase(), funcao: "escalante" });
        alert("Unidade cadastrada!"); window.location.reload();
    } catch (e) { alert("Erro: " + e.message); }
}

export function sair() { signOut(auth).then(() => location.reload()); }

onAuthStateChanged(auth, async (user) => {
    if (user) {
        usuarioAtual = user;
        const snap = await getDoc(doc(db, "usuarios", user.uid));
        if (snap.exists()) {
            perfilAtual = snap.data();
            document.getElementById('auth-container').style.display = 'none';
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
    select.innerHTML = "<option value=''>Carregando...</option>";
    try {
        const q = query(collection(db, "usuarios"), where("funcao", "==", "escalante"));
        const snapshot = await getDocs(q);
        select.innerHTML = "<option value='' selected>Selecione a Unidade...</option>";
        let unidades = [];
        snapshot.forEach(doc => unidades.push(doc.data().unidade));
        [...new Set(unidades)].sort().forEach(u => select.innerHTML += `<option value="${u}">${u}</option>`);
    } catch (e) { select.innerHTML = "<option>Erro ao carregar</option>"; }
}

export function adicionarOrdem() {
    const unidade = document.getElementById('select-unidade').value;
    const funcao = document.getElementById('select-funcao').value;
    const oficiais = document.getElementById('input-oficiais').value;
    const pracas = document.getElementById('input-pracas').value;

    if (!unidade) return alert("Selecione uma unidade!");
    if (oficiais == 0 && pracas == 0) return alert("Defina a quantidade.");

    listaOrdensTemporaria.push({ id: Date.now(), unidade, funcao, oficiais, pracas });
    atualizarTabelaOrdens();
}

function atualizarTabelaOrdens() {
    const corpo = document.getElementById('tabela-ordens-body');
    document.getElementById('contador-ordens').innerText = `${listaOrdensTemporaria.length} itens`;
    corpo.innerHTML = "";
    if (listaOrdensTemporaria.length === 0) return;

    listaOrdensTemporaria.forEach((item, index) => {
        corpo.innerHTML += `
            <tr class="border-bottom">
                <td class="ps-3 fw-bold">${item.unidade}</td>
                <td><small>${item.funcao}</small></td>
                <td class="small fw-bold">${item.oficiais} Of / ${item.pracas} Pç</td>
                <td class="text-end pe-3"><button onclick="window.app.excluirOrdem(${index})" class="btn btn-sm text-danger"><i class="bi bi-x-circle-fill"></i></button></td>
            </tr>`;
    });
}

export function excluirOrdem(index) { listaOrdensTemporaria.splice(index, 1); atualizarTabelaOrdens(); }
export function limparOrdens() { listaOrdensTemporaria = []; atualizarTabelaOrdens(); }

export async function dispararSolicitacao() {
    const evento = document.getElementById('nome-evento').value.trim();
    const data = document.getElementById('data-evento').value;
    const horario = document.getElementById('horario-evento').value;
    const prazo = document.getElementById('prazo-evento').value;

    if (!evento || !data || !horario) return alert("Preencha Nome, Data e Horário.");
    if (listaOrdensTemporaria.length === 0) return alert("Adicione ordens.");

    try {
        const promises = listaOrdensTemporaria.map(ordem => {
            return addDoc(collection(db, "escalas"), {
                evento, data, horario, prazo,
                unidade: ordem.unidade,
                funcao: ordem.funcao,
                cota: { oficial: ordem.oficiais, praca: ordem.pracas },
                status: "Pendente",
                militares: "[]", // Inicializa como array vazio stringificado
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
    lista.innerHTML = "<div class='text-center py-3'>Carregando...</div>";
    try {
        const q = query(collection(db, "escalas"), orderBy("data", "desc"));
        const snapshot = await getDocs(q);
        const grupos = new Map();

        snapshot.forEach(doc => {
            const d = doc.data();
            const chave = `${d.evento}|${d.data}`;
            if (!grupos.has(chave)) groups = grupos.set(chave, { evento: d.evento, data: d.data, horario: d.horario, total: 0, respondidos: 0, ids: [] });
            const g = grupos.get(chave);
            g.total++; g.ids.push(doc.id);
            if (d.status === "Preenchido") g.respondidos++;
        });

        lista.innerHTML = "";
        grupos.forEach((info, chave) => {
            const dataBr = new Date(info.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const percentual = Math.round((info.respondidos / info.total) * 100);
            const cor = percentual === 100 ? "success" : "warning";
            
            lista.innerHTML += `
                <div class="list-group-item p-3 border-bottom">
                    <div class="d-flex justify-content-between align-items-center mb-2">
                        <div>
                            <strong class="text-dark d-block">${info.evento}</strong>
                            <small class="text-muted">${dataBr} - ${info.horario || ''}</small>
                        </div>
                        <button onclick="window.app.excluirEvento('${info.evento}', '${info.data}')" class="btn btn-sm btn-outline-danger" title="Apagar todo histórico deste evento"><i class="bi bi-trash"></i></button>
                    </div>
                    <div class="d-flex justify-content-between small text-muted align-items-center mb-1">
                        <span onclick="window.app.abrirPreview('${info.evento}', '${info.data}')" style="cursor:pointer" class="text-primary fw-bold"> <i class="bi bi-eye"></i> Ver Detalhes</span>
                        <span>${info.respondidos}/${info.total} (${percentual}%)</span>
                    </div>
                    <div class="progress" style="height: 4px;"><div class="progress-bar bg-${cor}" style="width: ${percentual}%"></div></div>
                </div>`;
        });
    } catch(e) { lista.innerHTML = "<div class='text-danger text-center'>Erro ao carregar histórico.</div>"; }
}

export async function excluirEvento(evento, data) {
    if(!confirm(`Tem certeza que deseja EXCLUIR TODAS as solicitações de "${evento}"? Isso não pode ser desfeito.`)) return;
    try {
        const q = query(collection(db, "escalas"), where("evento", "==", evento), where("data", "==", data));
        const snap = await getDocs(q);
        const batch = [];
        snap.forEach(doc => deleteDoc(doc.ref)); // Delete um por um (simples)
        alert("Evento excluído do histórico.");
        carregarEventosAdmin();
    } catch(e) { alert("Erro ao excluir: " + e.message); }
}

// ================= PREVIEW & EXCEL =================
export async function abrirPreview(nomeEvento, dataEvento) {
    eventoPreviewAtual = { nome: nomeEvento, data: dataEvento };
    const modal = document.getElementById('preview-modal');
    modal.classList.remove('d-none'); modal.classList.add('d-flex');
    
    document.getElementById('preview-titulo').innerText = nomeEvento;
    document.getElementById('preview-data').innerText = new Date(dataEvento).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
    
    const corpo = document.getElementById('tabela-preview-corpo');
    corpo.innerHTML = "<tr><td colspan='6' class='text-center py-4'>Carregando...</td></tr>";

    try {
        const q = query(collection(db, "escalas"), where("evento", "==", nomeEvento), where("data", "==", dataEvento));
        const snapshot = await getDocs(q);
        let html = "";
        let ordemGlobal = 1;
        
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            let militares = [];
            try { militares = JSON.parse(d.militares); } catch(e) { militares = []; }

            // Se ninguem respondeu ainda, mostra a pendencia
            if(d.status === "Pendente") {
                html += `<tr class="table-danger text-muted"><td colspan="6" class="text-center small">Unidade ${d.unidade} ainda não enviou (Cota: ${d.cota.oficial} Of / ${d.cota.praca} Pç)</td></tr>`;
            } else {
                militares.forEach(m => {
                    html += `<tr>
                        <td class="fw-bold text-center">${ordemGlobal++}</td>
                        <td>${m.posto}</td>
                        <td>${m.nome} <strong class="text-uppercase">${m.guerra}</strong></td>
                        <td>${m.contato}</td>
                        <td>${d.unidade}</td>
                        <td><span class="badge bg-primary bg-opacity-10 text-primary">${d.funcao}</span></td>
                    </tr>`;
                });
            }
        });
        corpo.innerHTML = html;
        if (snapshot.docs.length > 0) document.getElementById('preview-horario').innerText = snapshot.docs[0].data().horario || '';

    } catch(e) { corpo.innerHTML = "<tr><td colspan='6' class='text-danger text-center'>Erro.</td></tr>"; }
}

export async function baixarExcelDoEvento() {
    if (!eventoPreviewAtual) return;
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala');
        
        // Dados do evento para o cabeçalho
        const q = query(collection(db, "escalas"), where("evento", "==", eventoPreviewAtual.nome), where("data", "==", eventoPreviewAtual.data), where("status", "==", "Preenchido"));
        const snapshot = await getDocs(q);
        let horarioEvento = "";
        if(!snapshot.empty) horarioEvento = snapshot.docs[0].data().horario || "";

        // 1. Cabeçalho Amarelo Mesclado
        const dataFormatada = new Date(eventoPreviewAtual.data).toLocaleDateString('pt-BR', {timeZone: 'UTC', weekday: 'long', day: '2-digit', month: '2-digit'}).toUpperCase();
        const tituloCompleto = `${dataFormatada} - ${eventoPreviewAtual.nome} / ${horarioEvento}`;
        
        worksheet.mergeCells('A1:F1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = tituloCompleto;
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Amarelo
        titleCell.font = { bold: true, size: 12, name: 'Arial' };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        titleCell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };

        // 2. Cabeçalho das Colunas (Cinza)
        const headerRow = worksheet.addRow(['Ord.', 'POSTO/GRAD.', 'NOME', 'CONTATO', 'UBM', 'FUNÇÃO']);
        headerRow.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } }; // Cinza
            cell.font = { bold: true, name: 'Arial', size: 10 };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        });

        // 3. Dados
        let contador = 1;
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            let militares = [];
            try { militares = JSON.parse(d.militares); } catch { return; }

            militares.forEach(m => {
                const nomeFormatado = {
                    richText: [
                        { font: { bold: true, name: 'Arial' }, text: m.guerra.toUpperCase() },
                        { font: { name: 'Arial' }, text: " " + m.nome.toUpperCase() }
                    ]
                };

                const row = worksheet.addRow([
                    contador++, 
                    m.posto, 
                    nomeFormatado, 
                    m.contato, 
                    d.unidade, 
                    d.funcao.toUpperCase()
                ]);

                row.eachCell((cell, colNum) => {
                    cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    // Coluna Nome à esquerda
                    if(colNum === 3) cell.alignment = { horizontal: 'left', indent: 1 };
                    // Coluna Função Amarela igual ao print (ou azul se preferir, pus amarelo pra bater com layout)
                    if(colNum === 6) {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; 
                        cell.font = { name: 'Arial', size: 9 };
                    }
                });
            });
        });

        // Largura das colunas
        worksheet.getColumn(1).width = 6;
        worksheet.getColumn(2).width = 15;
        worksheet.getColumn(3).width = 45;
        worksheet.getColumn(4).width = 18;
        worksheet.getColumn(5).width = 15;
        worksheet.getColumn(6).width = 25;

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), `${eventoPreviewAtual.nome}_Escala.xlsx`);
    } catch (e) { alert("Erro ao gerar Excel: " + e.message); }
}

// ================= UNIDADE: INPUTS E SALVAMENTO =================
async function carregarPendenciasUnidade() {
    const lista = document.getElementById('lista-unidade');
    lista.innerHTML = "<div class='text-center w-100'>Carregando...</div>";
    try {
        const q = query(collection(db, "escalas"), where("unidade", "==", perfilAtual.unidade), orderBy("data", "asc"));
        const snapshot = await getDocs(q);
        lista.innerHTML = "";
        
        if (snapshot.empty) return lista.innerHTML = "<div class='text-muted text-center w-100 mt-4'>Nada pendente.</div>";

        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            const dataFmt = new Date(d.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const isPendente = d.status === "Pendente";
            const btnClass = isPendente ? "btn-outline-danger" : "btn-outline-success";
            
            // Verifica prazo
            let textoPrazo = "";
            if(d.prazo) {
                const prazoDate = new Date(d.prazo);
                const hoje = new Date();
                if (hoje > prazoDate && isPendente) textoPrazo = `<div class="text-danger fw-bold small mt-1"><i class="bi bi-alarm"></i> Prazo Vencido: ${prazoDate.toLocaleDateString()}</div>`;
                else textoPrazo = `<div class="text-muted small mt-1">Prazo: ${prazoDate.toLocaleDateString()}</div>`;
            }

            lista.innerHTML += `
                <div class="col-md-6 col-lg-4">
                    <div class="glass-card p-4 h-100 border-start border-5 ${isPendente ? 'border-danger' : 'border-success'} d-flex flex-column">
                        <div class="d-flex justify-content-between mb-2">
                            <span class="badge bg-secondary">${dataFmt}</span>
                            <span class="badge ${isPendente ? 'bg-danger' : 'bg-success'}">${d.status}</span>
                        </div>
                        <h6 class="fw-bold mb-0 text-dark">${d.evento}</h6>
                        <small class="text-muted mb-2">${d.horario || ''}</small>
                        ${textoPrazo}
                        
                        <div class="bg-light p-3 rounded my-3 border text-center">
                            <strong class="d-block text-primary">${d.funcao}</strong>
                            <div class="small text-muted mt-1">Cota: ${d.cota.oficial} Of / ${d.cota.praca} Pç</div>
                        </div>

                        <button onclick="window.app.abrirEdicao('${docSnap.id}')" class="btn ${btnClass} w-100 fw-bold mt-auto">
                            ${isPendente ? 'RESPONDER' : 'EDITAR'}
                        </button>
                    </div>
                </div>`;
        });
    } catch(e) { console.error(e); lista.innerHTML = "Erro ao carregar."; }
}

export async function abrirEdicao(id) {
    escalaSelecionadaId = id;
    const docSnap = await getDoc(doc(db, "escalas", id));
    const d = docSnap.data();
    
    document.getElementById('titulo-evento-form').innerText = d.evento;
    document.getElementById('subtitulo-form').innerText = `${d.funcao} | Meta: ${d.cota.oficial} Oficiais, ${d.cota.praca} Praças`;
    
    // Gera as caixinhas (Inputs Dinâmicos)
    const container = document.getElementById('container-inputs-militares');
    container.innerHTML = "";

    // Tenta carregar dados salvos
    let dadosSalvos = [];
    try { dadosSalvos = JSON.parse(d.militares); } catch {}

    let contadorGlobal = 0;

    // Gera inputs para OFICIAIS
    for(let i=0; i < parseInt(d.cota.oficial); i++) {
        const salvo = dadosSalvos[contadorGlobal] || {};
        container.innerHTML += gerarHtmlMilitar(i, 'OFICIAL', salvo);
        contadorGlobal++;
    }

    // Gera inputs para PRAÇAS
    for(let i=0; i < parseInt(d.cota.praca); i++) {
        const salvo = dadosSalvos[contadorGlobal] || {};
        container.innerHTML += gerarHtmlMilitar(i, 'PRAÇA', salvo);
        contadorGlobal++;
    }

    document.getElementById('form-militar').style.display = 'block';
    document.getElementById('form-militar').scrollIntoView({ behavior: 'smooth' });
}

function gerarHtmlMilitar(index, tipo, dados) {
    return `
    <div class="input-militar-group">
        <h6 class="text-danger fw-bold mb-3 small border-bottom pb-2">${tipo} #${index + 1}</h6>
        <div class="row g-2 militar-row">
            <div class="col-4 col-md-3">
                <input type="text" class="form-control campo-posto" placeholder="Posto/Grad" value="${dados.posto || ''}">
            </div>
            <div class="col-8 col-md-5">
                <input type="text" class="form-control campo-nome" placeholder="Nome Completo" value="${dados.nome || ''}">
            </div>
            <div class="col-6 col-md-4">
                <input type="text" class="form-control campo-guerra" placeholder="Nome de Guerra" value="${dados.guerra || ''}">
            </div>
            <div class="col-6 col-md-12">
                <input type="text" class="form-control campo-tel" placeholder="Telefone" value="${dados.contato || ''}">
            </div>
        </div>
    </div>`;
}

export async function salvarEscala() {
    if (!escalaSelecionadaId) return;
    
    const rows = document.querySelectorAll('.militar-row');
    let listaFinal = [];
    let preenchidoCorretamente = true;

    rows.forEach(row => {
        const posto = row.querySelector('.campo-posto').value.trim();
        const nome = row.querySelector('.campo-nome').value.trim();
        const guerra = row.querySelector('.campo-guerra').value.trim();
        const contato = row.querySelector('.campo-tel').value.trim();

        if(!posto || !nome || !guerra) preenchidoCorretamente = false;

        listaFinal.push({ posto, nome, guerra, contato });
    });

    if(!preenchidoCorretamente) return alert("Por favor, preencha Posto, Nome e Nome de Guerra de todos os militares.");

    try {
        const jsonString = JSON.stringify(listaFinal);
        await updateDoc(doc(db, "escalas", escalaSelecionadaId), { militares: jsonString, status: "Preenchido" });
        
        // Gera o Recibo em PDF
        if(confirm("Escala enviada! Deseja baixar o comprovante (Recibo)?")) {
            gerarReciboPDF(listaFinal);
        }

        document.getElementById('form-militar').style.display = 'none';
        carregarPendenciasUnidade();
    } catch (e) { alert("Erro: " + e.message); }
}

function gerarReciboPDF(lista) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF();
    
    doc.setFont("helvetica", "bold");
    doc.text("RECIBO DE ENVIO DE ESCALA - CBMMA", 105, 20, null, null, "center");
    
    doc.setFontSize(10);
    doc.setFont("helvetica", "normal");
    doc.text(`Unidade: ${perfilAtual.unidade}`, 20, 30);
    doc.text(`Data do Envio: ${new Date().toLocaleString()}`, 20, 35);
    
    let y = 50;
    doc.setFontSize(9);
    doc.text("MILITARES ESCALADOS:", 20, 45);
    
    lista.forEach((m, i) => {
        doc.text(`${i+1}. ${m.posto} ${m.guerra} (${m.nome}) - Tel: ${m.contato}`, 20, y);
        y += 7;
    });

    doc.save(`Recibo_Escala_${perfilAtual.unidade}.pdf`);
}

window.app = { fazerLogin, fazerCadastro, sair, adicionarOrdem, limparOrdens, excluirOrdem, dispararSolicitacao, salvarEscala, abrirPreview, abrirEdicao, baixarExcelDoEvento, excluirEvento };