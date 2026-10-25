// Nidaros: bewaar uitsluitend het tijdelijke token van onze eigen server.
// AI-Hub access token, wachtwoord en TOTP-secret worden nooit in de browser bewaard.
const NidarosSession = (() => {
    const KEY = "nidaros.auth.session.v1";
    const MAX_IDLE = 30 * 60 * 1000;
    let lastInteraction = Date.now();

    function forget() {
        try { localStorage.removeItem(KEY); } catch { /* storage unavailable */ }
    }

    function save(token, expiresIn) {
        const seconds = Number(expiresIn);
        if (!token || !Number.isFinite(seconds) || seconds <= 0) {
            throw new Error("Sessiegegevens ontbreken of zijn ongeldig.");
        }
        const data = { token, expiresAt: Date.now() + seconds * 1000, lastInteraction: Date.now() };
        try { localStorage.setItem(KEY, JSON.stringify(data)); }
        catch { /* in-memory auth remains possible */ }
        lastInteraction = Date.now();
    }

    function load() {
        try {
            const data = JSON.parse(localStorage.getItem(KEY) || "null");
            if (!data || typeof data.token !== "string" ||
                !Number.isFinite(data.expiresAt) || !Number.isFinite(data.lastInteraction) ||
                Date.now() >= data.expiresAt || Date.now() - data.lastInteraction >= MAX_IDLE) {
                forget();
                return null;
            }
            return data.token;
        } catch {
            forget();
            return null;
        }
    }

    function touch() {
        lastInteraction = Date.now();
        try {
            const data = JSON.parse(localStorage.getItem(KEY) || "null");
            if (!data || !Number.isFinite(data.expiresAt) || Date.now() >= data.expiresAt) return;
            data.lastInteraction = lastInteraction;
            localStorage.setItem(KEY, JSON.stringify(data));
        } catch { /* storage unavailable */ }
    }

    function idleTooLong() {
        let persisted = lastInteraction;
        try {
            const data = JSON.parse(localStorage.getItem(KEY) || "null");
            if (Number.isFinite(data?.lastInteraction)) persisted = data.lastInteraction;
        } catch { /* storage unavailable */ }
        return Date.now() - persisted >= MAX_IDLE;
    }

    async function validate(token) {
        const response = await fetch("/api/auth/session", {
            method: "GET", cache: "no-store", headers: { Authorization: `Bearer ${token}` }
        });
        if (response.status === 401) return false;
        if (!response.ok) throw new Error(`Sessiecontrole mislukt (${response.status}).`);
        const data = await response.json();
        return data.valid === true;
    }

    return { forget, save, load, touch, idleTooLong, validate };
})();
