import { initializeApp, getApp, getApps, deleteApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getDatabase } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-database.js";
import { getAuth, setPersistence, browserLocalPersistence } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";

export const firebaseConfig = {
    apiKey: "AIzaSyA9l6N0ZOB8c5X_WZfzluqW8E0Tl1C1X6U",
    authDomain: "fmaero-smart-tracking-system.firebaseapp.com",
    databaseURL: "https://fmaero-smart-tracking-system-default-rtdb.asia-southeast1.firebasedatabase.app",
    projectId: "fmaero-smart-tracking-system",
    storageBucket: "fmaero-smart-tracking-system.firebasestorage.app",
    messagingSenderId: "170664221165",
    appId: "1:170664221165:web:034157d463c44387bd6fe4"
};

export const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
export const db = getDatabase(app);
export const auth = getAuth(app);

try {
    await setPersistence(auth, browserLocalPersistence);
} catch (error) {
    console.warn("Unable to enable auth persistence.", error);
}

export { deleteApp, initializeApp, getAuth };
