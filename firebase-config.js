import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { 
    getAuth, 
    signInWithEmailAndPassword, 
    createUserWithEmailAndPassword, 
    onAuthStateChanged, 
    signOut, 
    setPersistence, 
    browserSessionPersistence,
    deleteUser,
    EmailAuthProvider, // <--- OBRIGATÓRIO PARA CORRIGIR O ERRO
    reauthenticateWithCredential // <--- OBRIGATÓRIO PARA CORRIGIR O ERRO
} from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";
import { 
    getFirestore, 
    doc, 
    getDoc, 
    setDoc, 
    collection, 
    addDoc, 
    query, 
    where, 
    getDocs, 
    updateDoc, 
    orderBy, 
    deleteDoc,
    writeBatch,
    limit
} from "https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js";

const firebaseConfig = {
  apiKey: "AIzaSyD09C5yNzXPKD0j009yk-1XPRKZePoQW-o",
  authDomain: "escala-unificada.firebaseapp.com",
  projectId: "escala-unificada",
  storageBucket: "escala-unificada.firebasestorage.app",
  messagingSenderId: "621446429956",
  appId: "1:621446429956:web:251fe2d959aead95a8a451",
  measurementId: "G-9GF64GJ1DF"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// AQUI ESTÁ O SEGREDO: Exportando as novas funções para o app.js usar
export { 
    auth, db, 
    signInWithEmailAndPassword, createUserWithEmailAndPassword, onAuthStateChanged, signOut, setPersistence, browserSessionPersistence, deleteUser, 
    EmailAuthProvider, reauthenticateWithCredential, 
    doc, getDoc, setDoc, collection, addDoc, query, where, getDocs, updateDoc, orderBy, deleteDoc, writeBatch, limit 
};