import { requireRole, logoutToLogin } from "./auth-utils.js";

export async function guardPage(allowedRoles, options = {}) {
    const session = await requireRole(allowedRoles, options.redirectTo || "login.html");
    if (!session?.profile) return null;

    if (typeof options.onAuthorized === "function") {
        options.onAuthorized(session);
    }

    return session;
}

export { logoutToLogin };
