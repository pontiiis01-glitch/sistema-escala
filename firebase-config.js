// Importações via CDN (para funcionar direto no navegador/GitHub Pages)
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-auth.js";
import { getFirestore, doc, getDoc, collection, addDoc, query, where, getDocs, updateDoc } from "https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js";

// Suas chaves do projeto "escala-unificada"
const firebaseConfig = {
  apiKey: "AIzaSyD09C5yNzXPKD0j009yk-1XPRKZePoQW-o",
  authDomain: "escala-unificada.firebaseapp.com",
  projectId: "escala-unificada",
  storageBucket: "escala-unificada.firebasestorage.app",
  messagingSenderId: "621446429956",
  appId: "1:621446429956:web:251fe2d959aead95a8a451",
  measurementId: "G-9GF64GJ1DF"
};

// Inicializa o Firebase
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// Exporta as ferramentas para os outros arquivos usarem
export { auth, db, signInWithEmailAndPassword, onAuthStateChanged, signOut, doc, getDoc, collection, addDoc, query, where, getDocs, updateDoc };