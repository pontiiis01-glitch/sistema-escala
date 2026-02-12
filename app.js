import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

let usuarioAtual = null;
let perfilAtual = null;
let escalaSelecionadaId = null;
let eventoPreviewAtual = null;

// Lista temporária para o "Carrinho de Pedidos"
let listaOrdensTemporaria = [];

// ================= AUTH =================
export async function fazerLogin() {
    const email = document.getElementById('email-login').value;
    const senha = document.getElementById('senha-login').value;
    try { await signInWithEmailAndPassword(auth, email, senha); } 
    catch (e) { document.getElementById('msg-erro').innerText = "Login inválido."; }
}

export async function fazerCadastro() {
    const email = document.getElementById('email-cadastro').value;
    const senha = document.getElementById('senha-cadastro').value;
    const unidade = document.getElementById('unidade-cadastro').value;
    if(!email || !senha || !unidade) return alert("Preencha tudo.");
    
    try {
        const cred = await createUserWithEmailAndPassword(auth, email, senha);
        await setDoc(doc(db, "usuarios", cred.user.uid), {
            email, unidade: unidade.toUpperCase(), funcao: "escalante"
        });
        alert("Cadastrado!");
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
                carregarListaUnidades(); // NOVO: Carrega o select
                carregarEventosAdmin();
            } else {
                document.getElementById('unidade-area').style.display = 'block';
                carregarPendenciasUnidade();
            }
        }
    }
});

// ================= ADMIN: PREPARAÇÃO DE ORDENS =================

// 1. Busca todas as unidades cadastradas no sistema para o Select
async function carregarListaUnidades() {
    const select = document.getElementById('select-unidade');
    select.innerHTML = "<option value=''>Carregando...</option>";
    
    // Busca usuários que são escalantes (unidades)
    const q = query(collection(db, "usuarios"), where("funcao", "==", "escalante"));
    const snapshot = await getDocs(q);
    
    select.innerHTML = "<option value='' selected>Selecione uma Unidade...</option>";
    
    let unidades = [];
    snapshot.forEach(doc => unidades.push(doc.data().unidade));
    
    // Ordena alfabeticamente
    unidades.sort().forEach(u => {
        select.innerHTML += `<option value="${u}">${u}</option>`;
    });
}

// 2. Adiciona item na tabela temporária (carrinho)
export function adicionarOrdem() {
    const unidade = document.getElementById('select-unidade').value;
    const funcao = document.getElementById('select-funcao').value;
    const oficiais = document.getElementById('input-oficiais').value;
    const pracas = document.getElementById('input-pracas').value;

    if (!unidade) return alert("Selecione uma unidade!");
    if (oficiais == 0 && pracas == 0) return alert("Defina a quantidade de militares.");

    // Adiciona na lista
    listaOrdensTemporaria.push({
        id: Date.now(), // ID temporário
        unidade, funcao, oficiais, pracas
    });

    atualizarTabelaOrdens();
}

// 3. Atualiza o visual da tabela direita
function atualizarTabelaOrdens() {
    const corpo = document.getElementById('tabela-ordens-body');
    const contador = document.getElementById('contador-ordens');
    
    corpo.innerHTML = "";
    contador.innerText = `${listaOrdensTemporaria.length} ordens`;

    listaOrdensTemporaria.forEach((item, index) => {
        corpo.innerHTML += `
            <tr>
                <td class="fw-bold">${item.unidade}</td>
                <td><small>${item.funcao}</small></td>
                <td>${item.oficiais} Of / ${item.pracas} Pç</td>
                <td class="text-end">
                    <button onclick="window.app.removerOrdem(${index})" class="btn btn-sm btn-link text-danger p-0">
                        <i class="bi bi-trash"></i>
                    </button>
                </td>
            </tr>
        `;
    });
}

export function removerOrdem(index) {
    listaOrdensTemporaria.splice(index, 1);
    atualizarTabelaOrdens();
}

export function limparOrdens() {
    listaOrdensTemporaria = [];
    atualizarTabelaOrdens();
}

// ================= ADMIN: DISPARO REAL =================
export async function dispararSolicitacao() {
    const evento = document.getElementById('nome-evento').value.trim();
    const data = document.getElementById('data-evento').value;

    if (!evento || !data) return alert("Preencha o Nome do Evento e a Data.");
    if (listaOrdensTemporaria.length === 0) return alert("Adicione pelo menos uma ordem na lista.");

    try {
        // Cria um documento para cada linha da tabela
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
        alert(`Sucesso! ${listaOrdensTemporaria.length} ordens enviadas.`);
        
        // Limpa tudo
        limparOrdens();
        document.getElementById('nome-evento').value = "";
        carregarEventosAdmin();

    } catch (e) {
        console.error(e);
        alert("Erro ao disparar: " + e.message);
    }
}

// ================= ADMIN: VISUALIZAÇÃO E DOWNLOAD =================
async function carregarEventosAdmin() {
    const lista = document.getElementById('lista-eventos-admin');
    lista.innerHTML = "<div class='text-center small'>Atualizando...</div>";

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
    if (grupos.size === 0) lista.innerHTML = "<div class='text-muted small p-3'>Nada encontrado.</div>";

    grupos.forEach((info) => {
        const dataBr = new Date(info.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
        const percentual = Math.round((info.respondidos / info.total) * 100);
        const corBarra = percentual === 100 ? "bg-success" : "bg-warning";

        lista.innerHTML += `
            <div class="list-group-item list-group-item-action cursor-pointer p-3 mb-2 border rounded" 
                 onclick="window.app.abrirPreview('${info.evento}', '${info.data}')">
                <div class="d-flex justify-content-between">
                    <strong>${info.evento}</strong>
                    <span class="badge bg-light text-dark border">${dataBr}</span>
                </div>
                <div class="d-flex justify-content-between small text-muted mt-1">
                    <span>${info.respondidos}/${info.total} Ordens Prontas</span>
                    <span>${percentual}%</span>
                </div>
                <div class="progress mt-1" style="height: 4px;">
                    <div class="progress-bar ${corBarra}" style="width: ${percentual}%"></div>
                </div>
            </div>`;
    });
}

export async function abrirPreview(nomeEvento, dataEvento) {
    eventoPreviewAtual = { nome: nomeEvento, data: dataEvento };
    document.getElementById('preview-modal').style.display = 'flex';
    document.getElementById('preview-titulo').innerText = nomeEvento;
    document.getElementById('preview-data').innerText = new Date(dataEvento).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
    
    const corpo = document.getElementById('tabela-preview-corpo');
    corpo.innerHTML = "<tr><td colspan='4' class='text-center'>Carregando...</td></tr>";

    const q = query(collection(db, "escalas"), where("evento", "==", nomeEvento), where("data", "==", dataEvento));
    const snapshot = await getDocs(q);
    
    let html = "";
    snapshot.forEach(docSnap => {
        const d = docSnap.data();
        const statusClass = d.status === "Preenchido" ? "bg-ok" : "bg-pendente";
        const linhas = d.militares ? d.militares.split('\n').filter(l => l.trim().length > 3).length : 0;
        
        html += `<tr>
            <td class="fw-bold">${d.unidade}</td>
            <td><small>${d.funcao}</small></td>
            <td><span class="badge-status ${statusClass}">${d.status}</span></td>
            <td>${linhas} mil.</td>
        </tr>`;
    });
    corpo.innerHTML = html;
}

export async function baixarExcelDoEvento() {
    if (!eventoPreviewAtual) return;
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala');
        
        worksheet.columns = [
            { key: 'ord', width: 6 }, { key: 'posto', width: 12 }, { key: 'nome', width: 40 },
            { key: 'contato', width: 18 }, { key: 'unidade', width: 15 }, { key: 'funcao', width: 20 }
        ];

        const header = worksheet.addRow(['Ord', 'POSTO', 'NOME COMPLETO', 'CONTATO', 'UNIDADE', 'FUNÇÃO']);
        header.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } };
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'center' };
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
                    let posto = "SD BM";
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
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        cell.alignment = { horizontal: 'center' };
                        if(colNum === 3) cell.alignment = { horizontal: 'left', indent: 1 };
                        if(colNum === 6) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
                    });
                }
            });
        });
        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), `${eventoPreviewAtual.nome}.xlsx`);
    } catch (e) { alert("Erro: " + e.message); }
}

// ================= UNIDADE (Mantido igual) =================
async function carregarPendenciasUnidade() {
    const lista = document.getElementById('lista-unidade');
    lista.innerHTML = "Carregando...";
    const q = query(collection(db, "escalas"), where("unidade", "==", perfilAtual.unidade), orderBy("data", "asc"));
    
    try {
        const snapshot = await getDocs(q);
        lista.innerHTML = "";
        if (snapshot.empty) return lista.innerHTML = "<div class='text-muted'>Nada pendente.</div>";

        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            const dataFmt = new Date(d.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const statusColor = d.status === "Pendente" ? "border-danger" : "border-success";
            
            lista.innerHTML += `
                <div class="col-md-6">
                    <div class="card p-3 h-100 ${statusColor} border-start border-4">
                        <div class="d-flex justify-content-between mb-2">
                            <span class="badge bg-secondary">${dataFmt}</span>
                            <span class="badge ${d.status === 'Pendente' ? 'bg-danger' : 'bg-success'}">${d.status}</span>
                        </div>
                        <h6 class="fw-bold mb-1">${d.evento}</h6>
                        <div class="small text-muted mb-2">Função: <strong>${d.funcao}</strong></div>
                        <div class="small bg-light p-2 rounded mb-3 border">
                            Meta: <strong>${d.cota.oficial}</strong> Oficiais | <strong>${d.cota.praca}</strong> Praças
                        </div>
                        <button onclick="window.app.abrirEdicao('${docSnap.id}', '${d.evento}', '${d.cota.oficial}', '${d.cota.praca}', '${d.funcao}')" 
                                class="btn btn-sm btn-outline-primary w-100">
                            ${d.status === 'Pendente' ? 'Preencher Escala' : 'Editar'}
                        </button>
                    </div>
                </div>`;
        });
    } catch(e) { console.error(e); }
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
        alert("Enviado!");
        document.getElementById('form-militar').style.display = 'none';
        carregarPendenciasUnidade();
    } catch (e) { alert("Erro: " + e.message); }
}

export function removerOrdem(i) { listaOrdensTemporaria.splice(i, 1); atualizarTabelaOrdens(); }
export function limparOrdens() { listaOrdensTemporaria = []; atualizarTabelaOrdens(); }
function atualizarTabelaOrdens() {
    const corpo = document.getElementById('tabela-ordens-body');
    const contador = document.getElementById('contador-ordens');
    corpo.innerHTML = "";
    contador.innerText = `${listaOrdensTemporaria.length} ordens`;
    listaOrdensTemporaria.forEach((item, index) => {
        corpo.innerHTML += `<tr><td class="fw-bold">${item.unidade}</td><td><small>${item.funcao}</small></td><td>${item.oficiais} Of / ${item.pracas} Pç</td><td class="text-end"><button onclick="window.app.removerOrdem(${index})" class="btn btn-sm text-danger"><i class="bi bi-trash"></i></button></td></tr>`;
    });
}

window.app = { fazerLogin, fazerCadastro, sair, adicionarOrdem, limparOrdens, removerOrdem, dispararSolicitacao, salvarEscala, abrirPreview, abrirEdicao, baixarExcelDoEvento };