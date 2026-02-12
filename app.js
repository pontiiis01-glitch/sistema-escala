import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

let usuarioAtual = null;
let perfilAtual = null;
let escalaSelecionadaId = null;
let eventoPreviewAtual = null; // Guarda qual evento estamos visualizando na prévia

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
                carregarEventosAdmin();
            } else {
                document.getElementById('unidade-area').style.display = 'block';
                carregarPendenciasUnidade();
            }
        }
    }
});

// ================= ADMIN: DISPARO EM MASSA =================
export async function dispararSolicitacao() {
    const evento = document.getElementById('nome-evento').value.trim();
    const data = document.getElementById('data-evento').value;
    const funcao = document.getElementById('funcao-evento').value;
    const qtdOficiais = document.getElementById('qtd-oficiais').value;
    const qtdPracas = document.getElementById('qtd-pracas').value;
    const textoUnidades = document.getElementById('lista-unidades-alvo').value;

    if (!evento || !data || !textoUnidades) return alert("Preencha nome, data e unidades.");

    // Transforma "1GBM, 2GBM, BEM" em ["1GBM", "2GBM", "BEM"]
    const listaUnidades = textoUnidades.split(',').map(u => u.trim()).filter(u => u !== "");

    if (listaUnidades.length === 0) return alert("Nenhuma unidade válida identificada.");

    try {
        const promises = listaUnidades.map(unidade => {
            return addDoc(collection(db, "escalas"), {
                evento: evento,
                data: data,
                unidade: unidade.toUpperCase(),
                funcao: funcao,
                cota: { oficial: qtdOficiais, praca: qtdPracas },
                status: "Pendente",
                militares: "",
                criadoEm: new Date()
            });
        });

        await Promise.all(promises);
        alert(`Sucesso! ${listaUnidades.length} solicitações enviadas.`);
        
        // Limpa form
        document.getElementById('nome-evento').value = "";
        document.getElementById('lista-unidades-alvo').value = "";
        carregarEventosAdmin();

    } catch (e) {
        console.error(e);
        alert("Erro ao disparar: " + e.message);
    }
}

// ================= ADMIN: VISUALIZAÇÃO AGRUPADA =================
async function carregarEventosAdmin() {
    const lista = document.getElementById('lista-eventos-admin');
    lista.innerHTML = "<div class='text-center small'>Atualizando...</div>";

    // Pega todas as escalas
    const q = query(collection(db, "escalas"), orderBy("data", "desc"));
    const snapshot = await getDocs(q);
    
    // Agrupa por "Evento + Data" para não repetir na lista
    // Chave do Mapa: "Carnaval|2026-02-12"
    const grupos = new Map();

    snapshot.forEach(doc => {
        const d = doc.data();
        const chave = `${d.evento}|${d.data}`;
        
        if (!grupos.has(chave)) {
            grupos.set(chave, { evento: d.evento, data: d.data, total: 0, respondidos: 0 });
        }
        
        const g = grupos.get(chave);
        g.total++;
        if (d.status === "Preenchido") g.respondidos++;
    });

    lista.innerHTML = "";
    if (grupos.size === 0) lista.innerHTML = "<div class='text-muted small p-3'>Nada encontrado.</div>";

    grupos.forEach((info, chave) => {
        // Formata data para BR
        const dataBr = new Date(info.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
        const percentual = Math.round((info.respondidos / info.total) * 100);
        const corBarra = percentual === 100 ? "bg-success" : "bg-warning";

        lista.innerHTML += `
            <div class="list-group-item list-group-item-action cursor-pointer p-3 mb-2 border rounded" 
                 onclick="window.abrirPreviewWrapper('${info.evento}', '${info.data}')" style="cursor: pointer;">
                <div class="d-flex justify-content-between align-items-center mb-1">
                    <strong class="text-dark">${info.evento}</strong>
                    <span class="badge bg-light text-dark border">${dataBr}</span>
                </div>
                <div class="d-flex justify-content-between small text-muted mb-1">
                    <span>Respostas: ${info.respondidos}/${info.total}</span>
                    <span>${percentual}%</span>
                </div>
                <div class="progress" style="height: 6px;">
                    <div class="progress-bar ${corBarra}" style="width: ${percentual}%"></div>
                </div>
            </div>`;
    });
}

// ================= ADMIN: PRÉVIA E DOWNLOAD =================
export async function abrirPreview(nomeEvento, dataEvento) {
    eventoPreviewAtual = { nome: nomeEvento, data: dataEvento }; // Salva contexto
    
    document.getElementById('preview-modal').style.display = 'flex';
    document.getElementById('preview-titulo').innerText = nomeEvento;
    document.getElementById('preview-data').innerText = new Date(dataEvento).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
    
    const corpoTabela = document.getElementById('tabela-preview-corpo');
    corpoTabela.innerHTML = "<tr><td colspan='3' class='text-center'>Carregando detalhes...</td></tr>";

    // Busca apenas as escalas daquele evento/dia específico
    const q = query(
        collection(db, "escalas"), 
        where("evento", "==", nomeEvento),
        where("data", "==", dataEvento)
    );
    
    const snapshot = await getDocs(q);
    let html = "";
    let totalEfetivo = 0;

    snapshot.forEach(docSnap => {
        const d = docSnap.data();
        const statusClass = d.status === "Preenchido" ? "bg-ok" : "bg-pendente";
        
        // Conta linhas não vazias para estimar efetivo
        const linhas = d.militares ? d.militares.split('\n').filter(l => l.trim().length > 3).length : 0;
        totalEfetivo += linhas;

        html += `
            <tr>
                <td class="fw-bold">${d.unidade}</td>
                <td><span class="badge-status ${statusClass}">${d.status}</span></td>
                <td>${linhas} militares</td>
            </tr>
        `;
    });

    corpoTabela.innerHTML = html;
    document.getElementById('preview-total').innerText = totalEfetivo;
}

export async function baixarExcelDoEvento() {
    if (!eventoPreviewAtual) return;

    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala');
        
        // Setup Colunas
        worksheet.columns = [
            { key: 'ord', width: 6 },
            { key: 'posto', width: 12 },
            { key: 'nome', width: 40 },
            { key: 'contato', width: 18 },
            { key: 'unidade', width: 15 },
            { key: 'funcao', width: 20 }
        ];

        // Estilo Cabeçalho
        const header = worksheet.addRow(['Ord', 'POSTO', 'NOME COMPLETO', 'CONTATO', 'UNIDADE', 'FUNÇÃO']);
        header.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } };
            cell.font = { bold: true };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            cell.alignment = { horizontal: 'center' };
        });

        // Busca dados NOVAMENTE para garantir frescor
        const q = query(
            collection(db, "escalas"), 
            where("evento", "==", eventoPreviewAtual.nome),
            where("data", "==", eventoPreviewAtual.data),
            where("status", "==", "Preenchido") // Só baixa quem mandou
        );
        
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

                    // Detecção simples de posto
                    const espaco = nome.indexOf(' ');
                    if (espaco > 0 && espaco < 7) {
                        posto = nome.substring(0, espaco).toUpperCase();
                        nome = nome.substring(espaco).trim().toUpperCase();
                    }

                    const row = worksheet.addRow({
                        ord: contador++,
                        posto: posto,
                        nome: nome,
                        contato: contato,
                        unidade: d.unidade,
                        funcao: d.funcao.toUpperCase()
                    });
                    
                    // Estiliza linha
                    row.eachCell((cell, colNum) => {
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        cell.alignment = { horizontal: 'center' };
                        if(colNum === 3) cell.alignment = { horizontal: 'left', indent: 1 }; // Nome esquerda
                        if(colNum === 6) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } }; // Azul Função
                    });
                }
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        saveAs(new Blob([buffer]), `${eventoPreviewAtual.nome}_${eventoPreviewAtual.data}.xlsx`);

    } catch (e) {
        alert("Erro ao gerar Excel: " + e.message);
    }
}

// ================= UNIDADE: RESPOSTA =================
async function carregarPendenciasUnidade() {
    const lista = document.getElementById('lista-unidade');
    lista.innerHTML = "Carregando...";
    
    // Ordena pela data mais próxima
    const q = query(
        collection(db, "escalas"), 
        where("unidade", "==", perfilAtual.unidade),
        orderBy("data", "asc")
    );

    try {
        const snapshot = await getDocs(q);
        lista.innerHTML = "";
        
        if (snapshot.empty) return lista.innerHTML = "<div class='text-muted'>Nada pendente.</div>";

        snapshot.forEach(docSnap => {
            const d = docSnap.data();
            const dataFmt = new Date(d.data).toLocaleDateString('pt-BR', {timeZone: 'UTC'});
            const statusColor = d.status === "Pendente" ? "border-danger" : "border-success";
            const icon = d.status === "Pendente" ? "bi-exclamation-circle" : "bi-check-circle";

            lista.innerHTML += `
                <div class="col-md-6">
                    <div class="card p-3 h-100 ${statusColor} border-start border-4">
                        <div class="d-flex justify-content-between mb-2">
                            <span class="badge bg-secondary">${dataFmt}</span>
                            <span class="fw-bold small ${d.status === 'Pendente' ? 'text-danger' : 'text-success'}">
                                <i class="bi ${icon}"></i> ${d.status}
                            </span>
                        </div>
                        <h6 class="fw-bold mb-1">${d.evento}</h6>
                        <div class="small text-muted mb-3">
                            Meta: ${d.cota.oficial} Oficiais | ${d.cota.praca} Praças
                            <br>Função: ${d.funcao}
                        </div>
                        <button onclick="window.app.abrirEdicao('${docSnap.id}', '${d.evento}', '${d.cota.oficial}', '${d.cota.praca}', '${d.funcao}')" 
                                class="btn btn-sm btn-outline-primary w-100">
                            ${d.status === 'Pendente' ? 'Preencher Escala' : 'Editar Enviado'}
                        </button>
                    </div>
                </div>
            `;
        });
    } catch(e) {
        console.error(e); // Geralmente erro de índice composto
        if(e.message.includes("index")) alert("Atenção Admin: É necessário criar um índice no Firestore. Abra o console (F12) para ver o link.");
    }
}

export async function abrirEdicao(id, evento, of, pra, func) {
    escalaSelecionadaId = id;
    
    // Busca o texto atual do documento para preencher o textarea
    const docRef = doc(db, "escalas", id);
    const docSnap = await getDoc(docRef);
    const textoAtual = docSnap.data().militares || "";

    document.getElementById('titulo-evento-form').innerText = evento;
    document.getElementById('meta-oficiais').innerText = of;
    document.getElementById('meta-pracas').innerText = pra;
    document.getElementById('meta-funcao').innerText = func;
    document.getElementById('lista-nomes').value = textoAtual;
    
    document.getElementById('form-militar').style.display = 'block';
    document.getElementById('form-militar').scrollIntoView({ behavior: 'smooth' });
}

export async function salvarEscala() {
    if (!escalaSelecionadaId) return;
    const texto = document.getElementById('lista-nomes').value;
    
    try {
        await updateDoc(doc(db, "escalas", escalaSelecionadaId), {
            militares: texto,
            status: "Preenchido"
        });
        alert("Enviado com sucesso!");
        document.getElementById('form-militar').style.display = 'none';
        carregarPendenciasUnidade();
    } catch (e) { alert("Erro: " + e.message); }
}

// Exportação global
window.app = { fazerLogin, fazerCadastro, sair, dispararSolicitacao, salvarEscala, abrirPreview, abrirEdicao, baixarExcelDoEvento };