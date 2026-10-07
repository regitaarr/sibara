// ============================================================
// firebase-config.js
// Konfigurasi Firebase. Isi dengan config Firebase project Anda.
// ============================================================
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";
import { getAuth } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-auth.js";

// Konfigurasi Firebase project sibara-a3be1
const firebaseConfig = {
    apiKey: "AIzaSyCXceMzdyhOsKOQabiJ-1PRj-8It5Rzg8I",
    authDomain: "sibara-a3be1.firebaseapp.com",
    projectId: "sibara-a3be1",
    storageBucket: "sibara-a3be1.firebasestorage.app",
    messagingSenderId: "49788055578",
    appId: "1:49788055578:web:a2e7d805e2f47b56cb21a3",
    measurementId: "G-04D6CYLHE7"
};

const app = initializeApp(firebaseConfig);

export const db = getFirestore(app);
export const auth = getAuth(app);
