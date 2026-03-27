import { ref, get } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-database.js";
import { onAuthStateChanged, signOut } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import { auth, db } from "./firebase-client.js";

const AUTH_STORAGE_KEY = "fmaero_auth";

export function normalizeRole(role) {
    const normalized = String(role || "").trim().toLowerCase().replace(/[\s_-]+/g, "");

    if (normalized === "admin") return "admin";
    if (normalized === "manager") return "manager";
    if (normalized === "storekeeper") return "storekeeper";
    if (normalized === "sitesupervisor" || normalized === "supervisor") return "siteSupervisor";

    return "";
}

export function getRouteForRole(role) {
    if (role === "admin") return "admin-dashboard.html";
    if (role === "manager") return "manager-dashboard.html";
    if (role === "storekeeper") return "storekeeper-dashboard.html";
    if (role === "siteSupervisor") return "supervisor-dashboard.html";
    return "login.html";
}

function cacheSession(session) {
    if (!session?.profile) return;

    localStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify({
        userId: session.user.uid,
        name: session.profile.name || "",
        email: String(session.profile.email || session.user.email || "").trim().toLowerCase(),
        role: session.profile.role
    }));
}

export function clearCachedSession() {
    localStorage.removeItem(AUTH_STORAGE_KEY);
    localStorage.removeItem("userRole");
}

async function loadUserProfile(uid) {
    if (!uid) return null;

    const snapshot = await get(ref(db, `users/${uid}`));
    if (!snapshot.exists()) return null;

    const data = snapshot.val() || {};
    const role = normalizeRole(data.role);
    const status = String(data.status || "active").trim().toLowerCase();

    if (!role || status !== "active") return null;

    return {
        uid,
        name: String(data.name || "").trim(),
        email: String(data.email || "").trim().toLowerCase(),
        role,
        status
    };
}

export async function getCurrentSession() {
    const user = auth.currentUser;
    if (!user) {
        clearCachedSession();
        return null;
    }

    const profile = await loadUserProfile(user.uid);
    if (!profile) {
        clearCachedSession();
        return { user, profile: null };
    }

    const session = { user, profile };
    cacheSession(session);
    return session;
}

export function waitForSession() {
    return new Promise((resolve, reject) => {
        const unsubscribe = onAuthStateChanged(auth, async (user) => {
            unsubscribe();

            if (!user) {
                clearCachedSession();
                resolve(null);
                return;
            }

            try {
                const profile = await loadUserProfile(user.uid);
                if (!profile) {
                    clearCachedSession();
                    resolve({ user, profile: null });
                    return;
                }

                const session = { user, profile };
                cacheSession(session);
                resolve(session);
            } catch (error) {
                reject(error);
            }
        }, reject);
    });
}

export async function requireRole(allowedRoles, redirectTo = "login.html") {
    const session = await waitForSession();

    if (!session?.profile) {
        await signOut(auth).catch(() => {});
        clearCachedSession();
        window.location.replace(redirectTo);
        return null;
    }

    if (!allowedRoles.includes(session.profile.role)) {
        window.location.replace(getRouteForRole(session.profile.role));
        return null;
    }

    return session;
}

export async function logoutToLogin() {
    await signOut(auth).catch(() => {});
    clearCachedSession();
    window.location.replace("login.html");
}
