import { auth, db, signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc } from './firebase-config.js';
import ExcelJS from "https://cdn.skypack.dev/exceljs";
import { saveAs } from "https://cdn.skypack.dev/file-saver";

let usuarioAtual = null;
let perfilAtual = null;
let escalaSelecionadaId = null;

// =======================================================
// 1. SISTEMA DE AUTENTICAÇÃO (Login e Cadastro)
// =======================================================

export async function fazerLogin() {
    const email = document.getElementById('email-login').value;
    const senha = document.getElementById('senha-login').value;
    const msg = document.getElementById('msg-erro');
    
    try {
        await signInWithEmailAndPassword(auth, email, senha);
    } catch (error) {
        msg.innerText = "Erro ao entrar: Verifique email e senha.";
        console.error(error);
    }
}

export async function fazerCadastro() {
    const email = document.getElementById('email-cadastro').value;
    const senha = document.getElementById('senha-cadastro').value;
    const unidade = document.getElementById('unidade-cadastro').value;
    const msg = document.getElementById('msg-erro');

    if (!email || !senha || !unidade) {
        msg.innerText = "Preencha todos os campos.";
        return;
    }

    try {
        // 1. Cria usuário na Autenticação
        const userCredential = await createUserWithEmailAndPassword(auth, email, senha);
        const user = userCredential.user;

        // 2. Salva o perfil no Firestore (Sempre como 'escalante')
        await setDoc(doc(db, "usuarios", user.uid), {
            email: email,
            unidade: unidade.toUpperCase(),
            funcao: "escalante" // Segurança: ninguém vira admin sozinho
        });

        alert("Cadastro realizado! Bem-vindo, " + unidade);

    } catch (error) {
        console.error(error);
        if (error.code === 'auth/email-already-in-use') msg.innerText = "Email já cadastrado.";
        else if (error.code === 'auth/weak-password') msg.innerText = "Senha muito fraca.";
        else msg.innerText = "Erro: " + error.message;
    }
}

export function sair() {
    signOut(auth).then(() => window.location.reload());
}

// Monitora se o usuário está logado ou não
onAuthStateChanged(auth, async (user) => {
    if (user) {
        usuarioAtual = user;
        const docRef = doc(db, "usuarios", user.uid);
        const docSnap = await getDoc(docRef);
        
        if (docSnap.exists()) {
            perfilAtual = docSnap.data();
            iniciarDashboard();
        } else {
            // Caso raro: logou mas não tem perfil no banco
            alert("Erro: Usuário sem perfil. Contate o comando.");
            sair();
        }
    } else {
        document.getElementById('auth-container').style.display = 'block';
        document.getElementById('dashboard-screen').style.display = 'none';
    }
});

function iniciarDashboard() {
    document.getElementById('auth-container').style.display = 'none';
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

// =======================================================
// 2. FUNÇÕES DO COMANDO (ADMIN)
// =======================================================

export async function criarEscala() {
    const evento = document.getElementById('nome-evento').value;
    const unidade = document.getElementById('unidade-alvo').value;

    if (!evento || !unidade) return alert("Preencha todos os campos!");

    try {
        await addDoc(collection(db, "escalas"), {
            evento: evento,
            unidade: unidade.toUpperCase(),
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
    lista.innerHTML = "<div class='text-center p-2'>Atualizando...</div>";
    
    const q = query(collection(db, "escalas"));
    const querySnapshot = await getDocs(q);
    
    lista.innerHTML = "";
    if (querySnapshot.empty) {
        lista.innerHTML = "<div class='list-group-item'>Nenhuma solicitação.</div>";
        return;
    }

    querySnapshot.forEach((doc) => {
        const dados = doc.data();
        const corStatus = dados.status === "Pendente" ? "status-pendente" : "status-ok";
        
        lista.innerHTML += `
            <div class="list-group-item d-flex justify-content-between align-items-center">
                <div>
                    <strong>${dados.unidade}</strong>
                    <div class="small text-muted">${dados.evento}</div>
                </div>
                <span class="${corStatus}">${dados.status}</span>
            </div>`;
    });
}

// =======================================================
// 3. FUNÇÕES DA UNIDADE (ESCALANTE)
// =======================================================

async function carregarDadosUnidade() {
    const lista = document.getElementById('lista-unidade');
    lista.innerHTML = "Carregando...";

    const q = query(collection(db, "escalas"), where("unidade", "==", perfilAtual.unidade));
    const querySnapshot = await getDocs(q);
    
    lista.innerHTML = "";
    if (querySnapshot.empty) {
        lista.innerHTML = "<div class='text-muted p-2'>Nenhuma solicitação pendente.</div>";
        return;
    }

    querySnapshot.forEach((doc) => {
        const dados = doc.data();
        const btnClass = dados.status === "Pendente" ? "btn-outline-danger" : "btn-outline-success";
        const btnTexto = dados.status === "Pendente" ? "Responder Agora" : "Editar Resposta";

        const item = document.createElement('div');
        item.className = "list-group-item d-flex justify-content-between align-items-center mb-2 shadow-sm border";
        item.innerHTML = `
            <div>
                <h6 class="mb-0 fw-bold">${dados.evento}</h6>
                <small class="text-muted">Status: ${dados.status}</small>
            </div>
            <button class="btn btn-sm ${btnClass}">${btnTexto}</button>
        `;
        
        item.querySelector('button').onclick = () => abrirEdicao(doc.id, dados.evento, dados.militares);
        lista.appendChild(item);
    });
}

function abrirEdicao(id, evento, textoAtual) {
    escalaSelecionadaId = id;
    document.getElementById('evento-atual').innerText = evento;
    document.getElementById('lista-nomes').value = textoAtual || "";
    document.getElementById('form-militar').style.display = 'block';
    document.getElementById('form-militar').scrollIntoView({ behavior: 'smooth' });
}

export async function salvarEscala() {
    if (!escalaSelecionadaId) return;
    
    const texto = document.getElementById('lista-nomes').value;
    const docRef = doc(db, "escalas", escalaSelecionadaId);
    
    try {
        await updateDoc(docRef, {
            militares: texto,
            status: "Preenchido"
        });
        alert("Sucesso! Escala enviada.");
        document.getElementById('form-militar').style.display = 'none';
        carregarDadosUnidade();
    } catch (e) {
        alert("Erro ao salvar: " + e.message);
    }
}

// =======================================================
// 4. GERADOR DE EXCEL (DESIGN OFICIAL)
// =======================================================

export async function gerarRelatorioFinal() {
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Escala Operacional');

        // Configuração das Colunas
        worksheet.columns = [
            { key: 'ord', width: 8 },
            { key: 'posto', width: 15 },
            { key: 'nome', width: 45 },
            { key: 'contato', width: 20 },
            { key: 'ubm', width: 15 },
            { key: 'funcao', width: 20 }
        ];

        // Cabeçalho Estilizado (Cinza, Negrito, Bordas)
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
            const linhas = dados.militares.split("\n"); // Quebra texto por linha
            
            linhas.forEach(linhaTexto => {
                if (linhaTexto.trim().length > 3) { 
                    
                    // Tenta separar: "SD BM JOAO - 98888-8888"
                    const partes = linhaTexto.split("-");
                    
                    let posto = "SD BM"; // Padrão se não achar
                    let nome = partes[0] ? partes[0].trim() : "NOME INVÁLIDO";
                    let telefone = partes[1] ? partes[1].trim() : "S/ CONTATO";

                    // Tenta identificar o posto no início do nome (ex: "CB BM JOAO")
                    const espaco = nome.indexOf(" ");
                    if (espaco > 0 && espaco <= 6) {
                        posto = nome.substring(0, espaco).toUpperCase();
                        nome = nome.substring(espaco).trim().toUpperCase();
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
                        
                        // Coluna NOME: Alinhada à esquerda e Negrito
                        if(colNum === 3) { 
                            cell.alignment = { horizontal: 'left', indent: 1 };
                            cell.font = { name: 'Arial', size: 11, bold: true };
                        }
                        // Coluna FUNÇÃO: Fundo Azul Claro
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

// Expor funções para o HTML usar
window.app = { fazerLogin, fazerCadastro, sair, criarEscala, salvarEscala, gerarRelatorioFinal };