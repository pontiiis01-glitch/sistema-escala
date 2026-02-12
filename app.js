import { auth, db, signInWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, collection, addDoc, query, where, getDocs, updateDoc } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

let usuarioAtual = null;
let perfilAtual = null;
let escalaSelecionadaId = null;

// --- 1. AUTENTICAÇÃO ---
export async function fazerLogin() {
    const email = document.getElementById('email').value;
    const senha = document.getElementById('senha').value;
    const msgErro = document.getElementById('msg-erro');
    
    try {
        await signInWithEmailAndPassword(auth, email, senha);
    } catch (error) {
        console.error(error);
        msgErro.innerText = "Erro: Verifique email e senha.";
    }
}

export function sair() {
    signOut(auth).then(() => window.location.reload());
}

onAuthStateChanged(auth, async (user) => {
    if (user) {
        // Usuário logou, busca o perfil dele no banco
        usuarioAtual = user;
        const docRef = doc(db, "usuarios", user.uid);
        const docSnap = await getDoc(docRef);
        
        if (docSnap.exists()) {
            perfilAtual = docSnap.data();
            iniciarDashboard();
        } else {
            alert("ERRO CRÍTICO: Usuário sem perfil cadastrado na coleção 'usuarios'.");
            sair();
        }
    } else {
        document.getElementById('login-screen').style.display = 'block';
        document.getElementById('dashboard-screen').style.display = 'none';
    }
});

function iniciarDashboard() {
    document.getElementById('login-screen').style.display = 'none';
    document.getElementById('dashboard-screen').style.display = 'block';
    document.getElementById('titulo-unidade').innerText = perfilAtual.unidade;

    if (perfilAtual.funcao === 'admin') {
        document.getElementById('admin-area').style.display = 'block';
        carregarDadosAdmin();
    } else {
        document.getElementById('unidade-area').style.display = 'block';
        carregarDadosUnidade();
    }
}

// --- 2. FUNÇÕES DO ADMIN ---
export async function criarEscala() {
    const evento = document.getElementById('nome-evento').value;
    const unidade = document.getElementById('unidade-alvo').value;

    if (!evento || !unidade) return alert("Preencha todos os campos!");

    try {
        await addDoc(collection(db, "escalas"), {
            evento: evento,
            unidade: unidade,
            status: "Pendente",
            militares: "",
            dataCriacao: new Date()
        });
        alert("Solicitação enviada!");
        document.getElementById('nome-evento').value = "";
        document.getElementById('unidade-alvo').value = "";
        carregarDadosAdmin();
    } catch (e) {
        alert("Erro ao criar: " + e.message);
    }
}

async function carregarDadosAdmin() {
    const lista = document.getElementById('lista-admin');
    lista.innerHTML = "Atualizando...";
    
    const q = query(collection(db, "escalas")); // Pega todas
    const querySnapshot = await getDocs(q);
    
    lista.innerHTML = "";
    if (querySnapshot.empty) {
        lista.innerHTML = "<div class='list-group-item'>Nenhuma solicitação encontrada.</div>";
        return;
    }

    querySnapshot.forEach((doc) => {
        const dados = doc.data();
        const corStatus = dados.status === "Pendente" ? "bg-danger" : "bg-success";
        
        lista.innerHTML += `
            <div class="list-group-item d-flex justify-content-between align-items-center">
                <div>
                    <strong>${dados.unidade}</strong> <br>
                    <small>${dados.evento}</small>
                </div>
                <span class="badge ${corStatus}">${dados.status}</span>
            </div>`;
    });
}

// --- 3. FUNÇÕES DO ESCALANTE ---
async function carregarDadosUnidade() {
    const lista = document.getElementById('lista-unidade');
    lista.innerHTML = "Carregando...";

    // Filtra apenas as escalas desta unidade
    const q = query(collection(db, "escalas"), where("unidade", "==", perfilAtual.unidade));
    const querySnapshot = await getDocs(q);
    
    lista.innerHTML = "";
    if (querySnapshot.empty) {
        lista.innerHTML = "<div class='text-muted'>Você não tem solicitações pendentes.</div>";
        return;
    }

    querySnapshot.forEach((doc) => {
        const dados = doc.data();
        const btnClass = dados.status === "Pendente" ? "btn-outline-danger" : "btn-outline-success";
        const btnTexto = dados.status === "Pendente" ? "Responder" : "Editar";

        // Cria o botão que abre o formulário
        const item = document.createElement('div');
        item.className = "list-group-item d-flex justify-content-between align-items-center";
        item.innerHTML = `
            <div><strong>${dados.evento}</strong><br><small>Status: ${dados.status}</small></div>
            <button class="btn btn-sm ${btnClass}">${btnTexto}</button>
        `;
        
        // Adiciona evento de click no botão
        item.querySelector('button').onclick = () => abrirEdicao(doc.id, dados.evento, dados.militares);
        lista.appendChild(item);
    });
}

function abrirEdicao(id, evento, textoAtual) {
    escalaSelecionadaId = id;
    document.getElementById('evento-atual').innerText = evento;
    document.getElementById('lista-nomes').value = textoAtual || "";
    document.getElementById('form-militar').style.display = 'block';
    // Rola a tela até o formulário
    document.getElementById('form-militar').scrollIntoView({ behavior: 'smooth' });
}

export async function salvarEscala() {
    if (!escalaSelecionadaId) return;
    
    const texto = document.getElementById('lista-nomes').value;
    const docRef = doc(db, "escalas", escalaSelecionadaId);
    
    await updateDoc(docRef, {
        militares: texto,
        status: "Preenchido" // Atualiza status
    });
    
    alert("Escala salva e enviada com sucesso!");
    document.getElementById('form-militar').style.display = 'none';
    carregarDadosUnidade();
}

// --- 4. GERADOR DE EXCEL (DESIGN OFICIAL) ---
export async function gerarRelatorioFinal() {
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala Operacional');

        // Configuração das Colunas
        worksheet.columns = [
            { key: 'ord', width: 8 },
            { key: 'posto', width: 15 },
            { key: 'nome', width: 40 },
            { key: 'contato', width: 18 },
            { key: 'ubm', width: 12 },
            { key: 'funcao', width: 15 }
        ];

        // Cabeçalho Estilizado
        const headerRow = worksheet.addRow(['Ord.', 'POSTO/GRAD.', 'NOME', 'CONTATO', 'UBM', 'FUNÇÃO']);
        headerRow.eachCell((cell) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } }; // Cinza
            cell.font = { name: 'Arial', bold: true, size: 11 };
            cell.alignment = { vertical: 'middle', horizontal: 'center' };
            cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
        });

        // Busca dados preenchidos
        const q = query(collection(db, "escalas"), where("status", "==", "Preenchido"));
        const querySnapshot = await getDocs(q);
        
        let contador = 1;

        querySnapshot.forEach((docSnap) => {
            const dados = docSnap.data();
            const linhas = dados.militares.split("\n"); // Quebra por linha
            
            linhas.forEach(linhaTexto => {
                if (linhaTexto.trim().length > 3) { // Ignora linhas vazias
                    
                    // Tenta separar: "POSTO NOME - TELEFONE"
                    // Divide pelo traço "-"
                    const partes = linhaTexto.split("-");
                    
                    // Lógica para separar Posto do Nome (Separar pelo primeiro espaço)
                    let posto = "SD BM"; 
                    let nome = partes[0] ? partes[0].trim() : "NOME INVÁLIDO";
                    let telefone = partes[1] ? partes[1].trim() : "S/ CONTATO";

                    // Tenta adivinhar o posto (Ex: "CB BM JOAO" -> Pega o "CB BM")
                    const primeiroEspaco = nome.indexOf(" ");
                    if (primeiroEspaco > 0) {
                        // Se o começo for curto (até 6 letras), assume que é graduação (SD, CB, SGT)
                        if (primeiroEspaco <= 6) {
                            posto = nome.substring(0, primeiroEspaco).toUpperCase(); // "CB"
                            nome = nome.substring(primeiroEspaco).trim().toUpperCase(); // "JOAO"
                        }
                    }

                    const row = worksheet.addRow({
                        ord: contador++,
                        posto: posto,
                        nome: nome,
                        contato: telefone,
                        ubm: dados.unidade,
                        funcao: "SOCORRISTA"
                    });

                    // Estiliza a linha de dados
                    row.eachCell((cell, colNum) => {
                        cell.font = { name: 'Arial', size: 11 };
                        cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
                        cell.alignment = { vertical: 'middle', horizontal: 'center' };
                        
                        // Nome alinhado à esquerda e negrito
                        if(colNum === 3) { 
                            cell.alignment = { horizontal: 'left', indent: 1 };
                            cell.font = { name: 'Arial', size: 11, bold: true };
                        }
                        // Função com fundo azul
                        if(colNum === 6) {
                            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
                        }
                    });
                }
            });
        });

        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        saveAs(blob, "Escala_CBMMA_Final.xlsx");

    } catch (e) {
        console.error(e);
        alert("Erro ao gerar Excel: " + e.message);
    }
}

// Expor funções para o HTML
window.app = { fazerLogin, sair, criarEscala, salvarEscala, gerarRelatorioFinal, abrirEdicao };