import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy, setPersistence, browserSessionPersistence } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

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
        // CONFIGURAÇÃO DE SEGURANÇA: 
        // Define que o login só dura enquanto a aba estiver aberta. Fechou, deslogou.
        await setPersistence(auth, browserSessionPersistence);
        
        await signInWithEmailAndPassword(auth, email, senha); 
    } 
    catch (e) { 
        console.error(e);
        document.getElementById('msg-erro').innerText = "Credenciais inválidas. Tente novamente."; 
    }
}

export async function fazerCadastro() {
    const email = document.getElementById('email-cadastro').value;
    const senha = document.getElementById('senha-cadastro').value;
    const unidade = document.getElementById('unidade-cadastro').value;
    if(!email || !senha || !unidade) return alert("Preencha todos os campos.");
    
    try {
        const cred = await createUserWithEmailAndPassword(auth, email, senha);
        await setDoc(doc(db, "usuarios", cred.user.uid), {
            email, unidade: unidade.toUpperCase(), funcao: "escalante"
        });
        alert("Unidade cadastrada com sucesso!");
        window.location.reload();
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
                carregarListaUnidades();
                carregarEventosAdmin();
            } else {
                document.getElementById('unidade-area').style.display = 'block';
                carregarPendenciasUnidade();
            }
        }
    }
});

// ================= ADMIN: PREPARAÇÃO =================
async function carregarListaUnidades() {
    const select = document.getElementById('select-unidade');
    select.innerHTML = "<option value=''>Carregando...</option>";
    
    try {
        const q = query(collection(db, "usuarios"), where("funcao", "==", "escalante"));
        const snapshot = await getDocs(q);
        
        if (snapshot.empty) {
            select.innerHTML = "<option value=''>Nenhuma unidade cadastrada</option>";
            return;
        }

        select.innerHTML = "<option value='' selected>Selecione a Unidade (BBM/CIA)...</option>";
        
        let unidades = [];
        snapshot.forEach(doc => unidades.push(doc.data().unidade));
        unidades = [...new Set(unidades)].sort();
        unidades.forEach(u => select.innerHTML += `<option value="${u}">${u}</option>`);
        
    } catch (e) {
        console.error("Erro ao carregar unidades:", e);
        select.innerHTML = "<option value=''>Erro de permissão/conexão</option>";
    }
}

export function adicionarOrdem() {
    const unidade = document.getElementById('select-unidade').value;
    const funcao = document.getElementById('select-funcao').value;
    const oficiais = document.getElementById('input-oficiais').value;
    const pracas = document.getElementById('input-pracas').value;

    if (!unidade) return alert("Selecione uma unidade na lista!");
    if (oficiais == 0 && pracas == 0) return alert("Defina a quantidade de militares.");

    listaOrdensTemporaria.push({ id: Date.now(), unidade, funcao, oficiais, pracas });
    atualizarTabelaOrdens();
}

function atualizarTabelaOrdens() {
    const corpo = document.getElementById('tabela-ordens-body');
    const contador = document.getElementById('contador-ordens');
    
    corpo.innerHTML = "";
    contador.innerText = `${listaOrdensTemporaria.length} itens`;

    if (listaOrdensTemporaria.length === 0) {
        corpo.innerHTML = `<tr><td colspan="4" class="text-center text-muted py-3 small">Nenhuma ordem adicionada ainda.</td></tr>`;
        return;
    }

    listaOrdensTemporaria.forEach((item, index) => {
        corpo.innerHTML += `
            <tr>
                <td class="ps-3 fw-bold text-dark">${item.unidade}</td>
                <td><span class="badge bg-light text-secondary border">${item.funcao}</span></td>
                <td class="small fw-bold">${item.oficiais} Of / ${item.pracas} Pç</td>
                <td class="text-end pe-3">
                    <button onclick="window.app.excluirOrdem(${index})" class="btn btn-sm text-danger hover-scale">
                        <i class="bi bi-x-circle-fill"></i>
                    </button>
                </td>
            </tr>
        `;
    });
}

export function excluirOrdem(index) {
    listaOrdensTemporaria.splice(index, 1);
    atualizarTabelaOrdens();
}

export function limparOrdens() {
    listaOrdensTemporaria = [];
    atualizarTabelaOrdens();
}

export async function dispararSolicitacao() {
    const evento = document.getElementById('nome-evento').value.trim();
    const data = document.getElementById('data-evento').value;

    if (!evento || !data) return alert("Preencha o Nome da Operação e a Data.");
    if (listaOrdensTemporaria.length === 0) return alert("A lista de ordens está vazia.");

    try {
        const promises = listaOrdensTemporaria.map(ordem => {
            return addDoc(collection(db, "escalas"), {
                evento: evento,
                data: data,
                unidade: ordem.unidade,
                funcao: ordem.funcao,
                cota: { oficial: ordem.oficiais, praca: ordem.pracas },
                status: "Pendente",
                militares: "",
                criadoEm: new Date()
            });
        });

        await Promise.all(promises);
        alert(`Operação criada com sucesso! ${listaOrdensTemporaria.length} ordens enviadas.`);
        limparOrdens();
        document.getElementById('nome-evento').value = "";
        carregarEventosAdmin();

    } catch (e) { alert("Erro ao disparar: " + e.message); }
}

// ================= ADMIN: VISUALIZAÇÃO =================
async function carregarEventosAdmin() {
    const lista = document.getElementById('lista-eventos-admin');
    lista.innerHTML = "<div class='text-center small py-3'>Atualizando histórico...</div>";

    try {
        const q = query(collection(db, "escalas"), orderBy("data", "desc"));
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
        if (grupos.size === 0) lista.innerHTML = "<div class='text-center text-muted small py-4'>Nenhuma missão registrada.</div>";

        grupos.forEach((info) => {
            const dataBr = new Date(info.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const percentual = Math.round((info.respondidos / info.total) * 100);
            const corStatus = percentual === 100 ? "text-success" : "text-warning";
            const icon = percentual === 100 ? "bi-check-circle-fill" : "bi-clock-history";

            lista.innerHTML += `
                <div class="list-group-item list-group-item-action cursor-pointer p-3 border-bottom" 
                     onclick="window.app.abrirPreview('${info.evento}', '${info.data}')">
                    <div class="d-flex justify-content-between align-items-center mb-2">
                        <strong class="text-dark">${info.evento}</strong>
                        <span class="badge bg-light text-dark border">${dataBr}</span>
                    </div>
                    <div class="d-flex justify-content-between small text-muted align-items-center">
                        <span><i class="bi ${icon} ${corStatus} me-1"></i> ${info.respondidos}/${info.total} Unidades</span>
                        <span class="fw-bold">${percentual}%</span>
                    </div>
                    <div class="progress mt-2" style="height: 4px; background-color: #eee;">
                        <div class="progress-bar ${percentual === 100 ? 'bg-success' : 'bg-warning'}" style="width: ${percentual}%"></div>
                    </div>
                </div>`;
        });
    } catch(e) {
        console.error(e);
        lista.innerHTML = "<div class='text-danger small text-center'>Erro ao carregar histórico.</div>";
    }
}

export async function abrirPreview(nomeEvento, dataEvento) {
    eventoPreviewAtual = { nome: nomeEvento, data: dataEvento };
    const modal = document.getElementById('preview-modal');
    modal.classList.remove('d-none');
    modal.classList.add('d-flex');
    
    document.getElementById('preview-titulo').innerText = nomeEvento;
    document.getElementById('preview-data').innerText = new Date(dataEvento).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
    
    const corpo = document.getElementById('tabela-preview-corpo');
    corpo.innerHTML = "<tr><td colspan='4' class='text-center py-4'>Carregando detalhes...</td></tr>";

    try {
        const q = query(collection(db, "escalas"), where("evento", "==", nomeEvento), where("data", "==", dataEvento));
        const snapshot = await getDocs(q);
        
        let html = "";
        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            const statusBadge = d.status === "Preenchido" 
                ? `<span class="badge bg-success bg-opacity-10 text-success border border-success px-2 py-1">OK</span>` 
                : `<span class="badge bg-danger bg-opacity-10 text-danger border border-danger px-2 py-1">Pendente</span>`;
            
            const linhas = d.militares ? d.militares.split('\n').filter(l => l.trim().length > 3).length : 0;
            
            html += `<tr>
                <td class="ps-3 fw-bold">${d.unidade}</td>
                <td><small class="text-muted">${d.funcao}</small></td>
                <td>${statusBadge}</td>
                <td class="fw-bold text-dark">${linhas}</td>
            </tr>`;
        });
        corpo.innerHTML = html;
    } catch(e) {
        corpo.innerHTML = "<tr><td colspan='4' class='text-danger text-center'>Erro ao carregar detalhes.</td></tr>";
    }
}

export async function baixarExcelDoEvento() {
    if (!eventoPreviewAtual) return;
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala');
        
        worksheet.columns = [
            { key: 'ord', width: 6 }, { key: 'posto', width: 12 }, { key: 'nome', width: 45 },
            { key: 'contato', width: 20 }, { key: 'unidade', width: 15 }, { key: 'funcao', width: 25 }
        ];

        const header = worksheet.addRow(['Ord', 'POSTO', 'NOME COMPLETO', 'CONTATO', 'UNIDADE', 'FUNÇÃO']);
        header.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } };
            cell.font = { bold: true, name: 'Arial', size: 11 };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        });

        const q = query(collection(db, "escalas"), where("evento", "==", eventoPreviewAtual.nome), where("data", "==", eventoPreviewAtual.data), where("status", "==", "Preenchido"));
        const snapshot = await getDocs(q);
        let contador = 1;

        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            const linhas = d.militares.split('\n');
            linhas.forEach(linha => {
                if (linha.trim().length > 3) {
                    const partes = linha.split('-');
                    let posto = "SD BBM"; 
                    let nome = partes[0] ? partes[0].trim() : "";
                    let contato = partes[1] ? partes[1].trim() : "";
                    
                    const espaco = nome.indexOf(' ');
                    if (espaco > 0 && espaco < 7) {
                        posto = nome.substring(0, espaco).toUpperCase();
                        nome = nome.substring(espaco).trim().toUpperCase();
                    }

                    const row = worksheet.addRow({
                        ord: contador++, posto, nome, contato, unidade: d.unidade, funcao: d.funcao.toUpperCase()
                    });
                    
                    row.eachCell((cell, colNum) => {
                        cell.font = { name: 'Arial', size: 11 };
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        cell.alignment = { horizontal: 'center', vertical: 'middle' };
                        if(colNum === 3) { cell.alignment = { horizontal: 'left', indent: 1 }; cell.font = { bold: true }; }
                        if(colNum === 6) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
                    });
                }
            });
        });
        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), `${eventoPreviewAtual.nome}_${eventoPreviewAtual.data}.xlsx`);
    } catch (e) { alert("Erro ao gerar: " + e.message); }
}

// ================= UNIDADE =================
async function carregarPendenciasUnidade() {
    const lista = document.getElementById('lista-unidade');
    lista.innerHTML = "<div class='text-center w-100'>Carregando suas missões...</div>";
    
    try {
        const q = query(collection(db, "escalas"), where("unidade", "==", perfilAtual.unidade), orderBy("data", "asc"));
        const snapshot = await getDocs(q);
        lista.innerHTML = "";
        if (snapshot.empty) return lista.innerHTML = "<div class='text-muted text-center w-100 mt-4'>Nenhuma solicitação pendente. Tudo calmo por aqui.</div>";

        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            const dataFmt = new Date(d.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const isPendente = d.status === "Pendente";
            const statusClass = isPendente ? "border-danger" : "border-success";
            const btnClass = isPendente ? "btn-outline-danger" : "btn-outline-success";
            
            lista.innerHTML += `
                <div class="col-md-6 col-lg-4">
                    <div class="glass-card p-4 h-100 ${statusClass} border-start border-5 d-flex flex-column">
                        <div class="d-flex justify-content-between mb-3">
                            <span class="badge bg-secondary">${dataFmt}</span>
                            <span class="badge ${isPendente ? 'bg-danger' : 'bg-success'}">${d.status}</span>
                        </div>
                        <h6 class="fw-bold mb-1 text-dark">${d.evento}</h6>
                        <small class="text-muted mb-3 d-block">${d.funcao}</small>
                        
                        <div class="bg-light p-3 rounded mb-4 mt-auto border text-center">
                            <div class="d-flex justify-content-around">
                                <div><strong class="d-block fs-5">${d.cota.oficial}</strong><small class="text-muted" style="font-size:0.7rem">OFICIAIS</small></div>
                                <div class="vr"></div>
                                <div><strong class="d-block fs-5">${d.cota.praca}</strong><small class="text-muted" style="font-size:0.7rem">PRAÇAS</small></div>
                            </div>
                        </div>

                        <button onclick="window.app.abrirEdicao('${docSnap.id}', '${d.evento}', '${d.cota.oficial}', '${d.cota.praca}', '${d.funcao}')" 
                                class="btn ${btnClass} w-100 fw-bold">
                            ${isPendente ? 'PREENCHER ESCALA' : 'EDITAR ESCALA'}
                        </button>
                    </div>
                </div>`;
        });
    } catch(e) { 
        console.error(e);
        lista.innerHTML = "<div class='text-danger text-center w-100 mt-4'>Erro ao carregar missões. Verifique as regras.</div>";
    }
}

export async function abrirEdicao(id, evento, of, pra, func) {
    escalaSelecionadaId = id;
    const docSnap = await getDoc(doc(db, "escalas", id));
    document.getElementById('titulo-evento-form').innerText = evento;
    document.getElementById('meta-oficiais').innerText = of;
    document.getElementById('meta-pracas').innerText = pra;
    document.getElementById('meta-funcao').innerText = func;
    document.getElementById('lista-nomes').value = docSnap.data().militares || "";
    document.getElementById('form-militar').style.display = 'block';
    document.getElementById('form-militar').scrollIntoView({ behavior: 'smooth' });
}

export async function salvarEscala() {
    if (!escalaSelecionadaId) return;
    const texto = document.getElementById('lista-nomes').value;
    try {
        await updateDoc(doc(db, "escalas", escalaSelecionadaId), { militares: texto, status: "Preenchido" });
        alert("Escala enviada ao comando!");
        document.getElementById('form-militar').style.display = 'none';
        carregarPendenciasUnidade();
    } catch (e) { alert("Erro: " + e.message); }
}

window.app = { fazerLogin, fazerCadastro, sair, adicionarOrdem, limparOrdens, excluirOrdem, dispararSolicitacao, salvarEscala, abrirPreview, abrirEdicao, baixarExcelDoEvento };