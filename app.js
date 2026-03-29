import { ref, push, set, onValue, remove, update, runTransaction } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-database.js";
import { db } from "./firebase-client.js";
import { requireRole, logoutToLogin, clearCachedSession } from "./auth-utils.js";
const COMPANY_NAME = "FMAERO Smart Tracking System";
const COMPANY_LOGO_URL = "./logo_company_1.png"; // replace with your real logo path
let companyLogoDataUrl = "";

function resolveAssetUrl(path) {
    try {
        return new URL(path, window.location.href).href;
    } catch {
        return path;
    }
}

function toDataUrlFromBlob(blob) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
}

async function getCompanyLogoForPdf() {
    if (companyLogoDataUrl) return companyLogoDataUrl;

    const logoEl = document.getElementById("companyLogoImg");
    if (logoEl && logoEl.currentSrc) {
        try {
            const response = await fetch(logoEl.currentSrc);
            if (response.ok) {
                const blob = await response.blob();
                companyLogoDataUrl = await toDataUrlFromBlob(blob);
                return companyLogoDataUrl;
            }
        } catch {
            // fallback below
        }
    }

    const logoCandidates = [COMPANY_LOGO_URL, "./logo_company.jpeg"];
    for (const candidate of logoCandidates) {
        try {
            const absoluteUrl = resolveAssetUrl(candidate);
            const response = await fetch(absoluteUrl);
            if (response.ok) {
                const blob = await response.blob();
                companyLogoDataUrl = await toDataUrlFromBlob(blob);
                return companyLogoDataUrl;
            }
            if (candidate === COMPANY_LOGO_URL) {
                return absoluteUrl;
            }
        } catch {
            // try next candidate
        }
    }

    return resolveAssetUrl(COMPANY_LOGO_URL);
}

const ROLES = {
    MANAGER: "manager",
    SITE_SUPERVISOR: "siteSupervisor",
    STOREKEEPER: "storekeeper"
};

const ROLE_LAYOUTS = {
    [ROLES.MANAGER]: {
        visibleSections: ["overviewSection", "materialsSection", "transactionsSection", "chartsSection", "settingsSection", "reportsSection"],
        showAddMaterial: false,
        materialActionHeader: "Details"
    },
    [ROLES.SITE_SUPERVISOR]: {
        visibleSections: ["overviewSection", "materialsSection", "transactionsSection", "chartsSection", "settingsSection"],
        showAddMaterial: false,
        materialActionHeader: "Details"
    },
    [ROLES.STOREKEEPER]: {
        visibleSections: ["overviewSection", "materialsSection", "transactionsSection", "scannerSection", "chartsSection", "settingsSection", "reportsSection"],
        showAddMaterial: true,
        materialActionHeader: "Action"
    }
};

const SETTINGS_STORAGE_KEY = "fmaero_settings_v1";
const REPORT_INFO_STORAGE_KEY = "fmaero_report_info_v1";
const LOW_STOCK_THRESHOLD = 10;
const DEFAULT_SETTINGS = {
    theme: "blue",
    defaultRange: "all"
};
const DEFAULT_REPORT_INFO = {
    referenceNumber: "",
    zone: "",
    project: "",
    department: "",
    productionGroup: "",
    drawingNumber: "",
    startDate: "",
    finishDate: "",
    pic: "",
    remark: "",
    feedback: ""
};

const state = {
    materials: [],
    transactions: [],
    exportLogs: [],
    filter: "all",
    sort: "newest",
    search: "",
    user: null,
    role: null,
    userProfile: null,
    settings: { ...DEFAULT_SETTINGS },
    reportInfo: { ...DEFAULT_REPORT_INFO },
    transactionSearch: "",
    transactionPage: 1,
    transactionPageSize: 10,
    transactionSortKey: "timestamp",
    transactionSortDirection: "desc"
};

let detachRoleListener = null;
let stockMovementChartInstance = null;
let stockStatusChartInstance = null;
let topMaterialsChartInstance = null;
let qrScannerInstance = null;
let isQrRunning = false;
let resolveStockModal = null;
let resolveEditModal = null;
let resolveConfirmModal = null;
let resolveScanTxModal = null;
let activeMaterialForQr = null;
let isScanFlowActive = false;
const QR_SCAN_COOLDOWN_MS = 1500;
let lastQrScanCode = "";
let lastQrScanAt = 0;
const RFID_AUTO_SUBMIT_DELAY_MS = 180;
let rfidAutoSubmitTimer = null;
let lastExternalRfidSignature = "";
let lastExternalRfidAt = 0;
let hasPrimedExternalRfidListener = false;
let hasRealtimeBindings = false;

function applyTheme(theme) {
    const selected = ["blue", "teal", "slate"].includes(theme) ? theme : DEFAULT_SETTINGS.theme;
    document.body.setAttribute("data-theme", selected);
}

function loadSettings() {
    try {
        const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
        if (!raw) return { ...DEFAULT_SETTINGS };
        const parsed = JSON.parse(raw);
        return {
            theme: parsed.theme || DEFAULT_SETTINGS.theme,
            defaultRange: parsed.defaultRange || DEFAULT_SETTINGS.defaultRange
        };
    } catch {
        return { ...DEFAULT_SETTINGS };
    }
}

function getRoleLayout(role) {
    return ROLE_LAYOUTS[role] || {
        visibleSections: ["overviewSection", "materialsSection", "transactionsSection", "chartsSection", "settingsSection"],
        showAddMaterial: false,
        materialActionHeader: "Action"
    };
}

function persistSettings() {
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(state.settings));
}

function loadReportInfo() {
    try {
        const raw = localStorage.getItem(REPORT_INFO_STORAGE_KEY);
        if (!raw) return { ...DEFAULT_REPORT_INFO };
        const parsed = JSON.parse(raw);
        return {
            referenceNumber: parsed.referenceNumber || "",
            zone: parsed.zone || "",
            project: parsed.project || "",
            department: parsed.department || "",
            productionGroup: parsed.productionGroup || "",
            drawingNumber: parsed.drawingNumber || "",
            startDate: parsed.startDate || "",
            finishDate: parsed.finishDate || "",
            pic: parsed.pic || "",
            remark: parsed.remark || "",
            feedback: parsed.feedback || ""
        };
    } catch {
        return { ...DEFAULT_REPORT_INFO };
    }
}

function persistReportInfo() {
    localStorage.setItem(REPORT_INFO_STORAGE_KEY, JSON.stringify(state.reportInfo));
}

function syncReportInfoUi() {
    const mappings = {
        reportReferenceNumber: "referenceNumber",
        reportZone: "zone",
        reportProject: "project",
        reportDepartment: "department",
        reportProductionGroup: "productionGroup",
        reportDrawingNumber: "drawingNumber",
        reportStartDate: "startDate",
        reportFinishDate: "finishDate",
        reportPic: "pic",
        reportRemark: "remark",
        reportFeedback: "feedback"
    };

    Object.entries(mappings).forEach(([id, key]) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.value = state.reportInfo[key] || "";
    });
}

window.saveReportInfo = function () {
    const read = (id) => (document.getElementById(id)?.value || "").trim();

    state.reportInfo = {
        referenceNumber: read("reportReferenceNumber"),
        zone: read("reportZone"),
        project: read("reportProject"),
        department: read("reportDepartment"),
        productionGroup: read("reportProductionGroup"),
        drawingNumber: read("reportDrawingNumber"),
        startDate: read("reportStartDate"),
        finishDate: read("reportFinishDate"),
        pic: read("reportPic"),
        remark: read("reportRemark"),
        feedback: read("reportFeedback")
    };

    persistReportInfo();
    syncReportInfoUi();
    notify("Report information saved.", "success");
};

window.resetReportInfo = function () {
    state.reportInfo = { ...DEFAULT_REPORT_INFO };
    persistReportInfo();
    syncReportInfoUi();
    notify("Report information reset.", "success");
};

function applyDefaultDateRange() {
    const fromInput = document.getElementById("txFromDate");
    const toInput = document.getElementById("txToDate");
    if (!fromInput || !toInput) return;

    const option = state.settings.defaultRange || "all";
    if (option === "all") return;

    const days = Number(option);
    if (!Number.isFinite(days) || days <= 0) return;

    const now = new Date();
    const from = new Date();
    from.setDate(now.getDate() - (days - 1));

    const toIso = now.toISOString().slice(0, 10);
    const fromIso = from.toISOString().slice(0, 10);
    fromInput.value = fromIso;
    toInput.value = toIso;
}

function updateSettingsProfileUI() {
    const nameEl = document.getElementById("settingsProfileName");
    const emailEl = document.getElementById("settingsProfileEmail");
    const roleEl = document.getElementById("settingsProfileRole");
    const themeSelect = document.getElementById("settingsThemeSelect");
    const rangeSelect = document.getElementById("settingsDefaultRange");

    if (nameEl) nameEl.textContent = state.userProfile?.name || state.user?.displayName || "-";
    if (emailEl) emailEl.textContent = state.user?.email || state.userProfile?.email || "-";
    if (roleEl) roleEl.textContent = formatRoleLabel(state.role);
    if (themeSelect) themeSelect.value = state.settings.theme || DEFAULT_SETTINGS.theme;
    if (rangeSelect) rangeSelect.value = state.settings.defaultRange || DEFAULT_SETTINGS.defaultRange;
}

function updateScannerUiStatus(statusText) {
    const statusEl = document.getElementById("qrScannerStatus");
    if (statusEl) statusEl.textContent = `Status: ${statusText}`;
}

function updateLastScanSummary(materialCode, action, quantity, date = new Date()) {
    const materialEl = document.getElementById("lastScanMaterial");
    const actionEl = document.getElementById("lastScanAction");
    const quantityEl = document.getElementById("lastScanQuantity");
    const timeEl = document.getElementById("lastScanTime");

    if (materialEl) materialEl.textContent = materialCode || "-";
    if (actionEl) actionEl.textContent = action || "-";
    if (quantityEl) quantityEl.textContent = Number.isFinite(Number(quantity)) ? String(quantity) : "-";
    if (timeEl) {
        const formattedTime = date.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
        timeEl.textContent = formattedTime;
    }
}

function updateOverviewLastFlow(materialCode, action, quantity, timeLabel = "-", pic = "-") {
    const badgeEl = document.getElementById("overviewLastFlowBadge");
    const materialEl = document.getElementById("overviewLastFlowMaterial");
    const qtyEl = document.getElementById("overviewLastFlowQuantity");
    const timeEl = document.getElementById("overviewLastFlowTime");
    const picEl = document.getElementById("overviewLastFlowPic");
    const normalizedAction = String(action || "").trim().toUpperCase();

    if (badgeEl) {
        badgeEl.textContent = normalizedAction || "No Data";
        badgeEl.classList.remove("in", "out");
        if (normalizedAction === "IN") badgeEl.classList.add("in");
        if (normalizedAction === "OUT") badgeEl.classList.add("out");
        if (!normalizedAction) badgeEl.textContent = "No Data";
    }
    if (materialEl) materialEl.textContent = materialCode || "-";
    if (qtyEl) qtyEl.textContent = Number.isFinite(Number(quantity)) ? String(quantity) : "-";
    if (timeEl) timeEl.textContent = timeLabel || "-";
    if (picEl) picEl.textContent = pic || "-";
}

function notify(message, type = "info") {
    const container = document.getElementById("toastContainer");
    if (!container) return;

    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);

    setTimeout(() => {
        toast.remove();
    }, 3200);
}

function clearRoleListener() {
    if (detachRoleListener) {
        detachRoleListener();
        detachRoleListener = null;
    }
}

function setModalOpen(modalId, open) {
    const modal = document.getElementById(modalId);
    if (!modal) return;
    modal.classList.toggle("open", open);
    modal.setAttribute("aria-hidden", open ? "false" : "true");
}

function openStockModal(type) {
    const titleEl = document.getElementById("stockModalTitle");
    const amountEl = document.getElementById("stockModalAmount");
    const picEl = document.getElementById("stockModalPic");
    const reasonEl = document.getElementById("stockModalReason");
    const remarkEl = document.getElementById("stockModalRemark");

    if (!titleEl || !amountEl || !picEl || !reasonEl || !remarkEl) return Promise.resolve(null);

    titleEl.textContent = type === "IN" ? "Stock IN" : "Stock OUT";
    amountEl.value = "";
    picEl.value = state.userProfile?.name || state.user?.displayName || "";
    reasonEl.value = type === "OUT" ? "Site Usage" : "Other";
    remarkEl.value = "";
    setModalOpen("stockModal", true);
    amountEl.focus();

    return new Promise((resolve) => {
        resolveStockModal = resolve;
    });
}

function closeStockModal(result) {
    setModalOpen("stockModal", false);
    if (resolveStockModal) {
        resolveStockModal(result);
        resolveStockModal = null;
    }
}

function normalizeRfidTag(value) {
    return String(value || "").trim().toUpperCase();
}

function generateUniqueMaterialCode() {
    const usedCodes = new Set(
        state.materials
            .map((material) => normalizeRfidTag(material.code))
            .filter((code) => /^MAT\d+$/.test(code))
    );

    let nextNumber = 1000;
    while (usedCodes.has(`MAT${nextNumber}`)) {
        nextNumber += 1;
    }

    return `MAT${nextNumber}`;
}

function findMaterialByScanCode(value) {
    const normalized = normalizeRfidTag(value);
    if (!normalized) return null;

    return state.materials.find((material) => {
        const materialCode = normalizeRfidTag(material.code);
        const rfidTag = normalizeRfidTag(material.rfidTag);
        return normalized === rfidTag || normalized === materialCode;
    }) || null;
}

function queueRfidAutoSubmit() {
    if (rfidAutoSubmitTimer) {
        clearTimeout(rfidAutoSubmitTimer);
    }

    rfidAutoSubmitTimer = window.setTimeout(() => {
        rfidAutoSubmitTimer = null;
        window.processRfidInput();
    }, RFID_AUTO_SUBMIT_DELAY_MS);
}

function normalizeScanType(value, fallback = "IN") {
    const upper = String(value || "").trim().toUpperCase();
    return upper === "OUT" ? "OUT" : fallback;
}

async function handleExternalRfidScan(payload) {
    const rawTag = payload?.tag ?? payload?.code ?? payload?.rfid ?? "";
    const tag = String(rawTag || "").trim();
    if (!tag) return;
    const preferredType = normalizeScanType(payload?.action ?? payload?.type, "IN");

    const scannedAt = String(payload?.scannedAt || payload?.timestamp || payload?.createdAt || "");
    const source = String(payload?.source || "external-rfid").trim() || "external-rfid";
    const signature = `${tag}|${scannedAt}|${source}|${preferredType}`;
    const nowMs = Date.now();
    if (signature === lastExternalRfidSignature && nowMs - lastExternalRfidAt < QR_SCAN_COOLDOWN_MS) {
        return;
    }

    lastExternalRfidSignature = signature;
    lastExternalRfidAt = nowMs;

    const rfidInput = document.getElementById("rfidInput");
    const rfidResultEl = document.getElementById("rfidLastResult");
    if (rfidInput) rfidInput.value = tag;
    if (rfidResultEl) rfidResultEl.textContent = `Last RFID: ${tag} (${source} | ${preferredType})`;

    await handleScannedCode(tag, "rfid", { preferredType });

    if (rfidInput) rfidInput.value = "";
}

function openEditModal(name, stock, rfidTag = "") {
    const nameEl = document.getElementById("editModalName");
    const stockEl = document.getElementById("editModalStock");
    const rfidEl = document.getElementById("editModalRfidTag");
    if (!nameEl || !stockEl || !rfidEl) return Promise.resolve(null);

    nameEl.value = name || "";
    stockEl.value = Number.isFinite(stock) ? String(stock) : String(Number(stock) || 0);
    rfidEl.value = rfidTag || "";
    setModalOpen("editModal", true);
    nameEl.focus();

    return new Promise((resolve) => {
        resolveEditModal = resolve;
    });
}

function closeEditModal(result) {
    setModalOpen("editModal", false);
    if (resolveEditModal) {
        resolveEditModal(result);
        resolveEditModal = null;
    }
}

function openConfirmModal(message) {
    const msgEl = document.getElementById("confirmModalMessage");
    if (msgEl) msgEl.textContent = message;
    setModalOpen("confirmModal", true);

    return new Promise((resolve) => {
        resolveConfirmModal = resolve;
    });
}

function closeConfirmModal(result) {
    setModalOpen("confirmModal", false);
    if (resolveConfirmModal) {
        resolveConfirmModal(result);
        resolveConfirmModal = null;
    }
}

function openScanTxModal(material, options = {}) {
    const codeEl = document.getElementById("scanTxCode");
    const nameEl = document.getElementById("scanTxName");
    const stockEl = document.getElementById("scanTxStock");
    const typeEl = document.getElementById("scanTxType");
    const qtyEl = document.getElementById("scanTxQuantity");
    const reasonEl = document.getElementById("scanTxReason");
    const remarkEl = document.getElementById("scanTxRemark");

    if (!codeEl || !nameEl || !stockEl || !typeEl || !qtyEl || !reasonEl || !remarkEl) return Promise.resolve(null);

    codeEl.textContent = material?.code || "-";
    nameEl.textContent = material?.name || "-";
    stockEl.textContent = String(Number(material?.stock) || 0);
    typeEl.value = normalizeScanType(options?.preferredType, "IN");
    qtyEl.value = "";
    reasonEl.value = "Other";
    remarkEl.value = "";

    setModalOpen("scanTxModal", true);
    qtyEl.focus();

    return new Promise((resolve) => {
        resolveScanTxModal = resolve;
    });
}

function resetScanTxForm() {
    const codeEl = document.getElementById("scanTxCode");
    const nameEl = document.getElementById("scanTxName");
    const stockEl = document.getElementById("scanTxStock");
    const typeEl = document.getElementById("scanTxType");
    const qtyEl = document.getElementById("scanTxQuantity");
    const reasonEl = document.getElementById("scanTxReason");
    const remarkEl = document.getElementById("scanTxRemark");

    if (codeEl) codeEl.textContent = "-";
    if (nameEl) nameEl.textContent = "-";
    if (stockEl) stockEl.textContent = "0";
    if (typeEl) typeEl.value = "IN";
    if (qtyEl) qtyEl.value = "";
    if (reasonEl) reasonEl.value = "Other";
    if (remarkEl) remarkEl.value = "";
}

function closeScanTxModal(result) {
    setModalOpen("scanTxModal", false);
    if (resolveScanTxModal) {
        resolveScanTxModal(result);
        resolveScanTxModal = null;
    }
}

function closeMaterialDetailsModal() {
    setModalOpen("materialDetailsModal", false);
    activeMaterialForQr = null;
}

function getMaterialQrNode() {
    const container = document.getElementById("materialQrCode");
    if (!container) return null;
    return container.querySelector("canvas, img");
}

function renderMaterialQr(materialCode) {
    const qrContainer = document.getElementById("materialQrCode");
    const qrContent = document.getElementById("materialQrContent");

    if (qrContent) qrContent.textContent = materialCode || "-";
    if (!qrContainer) return;

    qrContainer.innerHTML = "";

    if (!materialCode) {
        qrContainer.textContent = "Material code is missing.";
        return;
    }

    if (typeof QRCode === "undefined") {
        qrContainer.textContent = "QR generator failed to load.";
        return;
    }

    new QRCode(qrContainer, {
        text: materialCode,
        width: 220,
        height: 220,
        colorDark: "#0f172a",
        colorLight: "#ffffff",
        correctLevel: QRCode.CorrectLevel.M
    });
}

function createMaterialQrLabelCanvas(materialCode) {
    const qrNode = getMaterialQrNode();
    if (!qrNode) return null;

    const labelCanvas = document.createElement("canvas");
    const qrSize = 220;
    const padding = 20;
    const footerHeight = 38;

    labelCanvas.width = qrSize + padding * 2;
    labelCanvas.height = qrSize + padding * 2 + footerHeight;

    const ctx = labelCanvas.getContext("2d");
    if (!ctx) return null;

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, labelCanvas.width, labelCanvas.height);
    ctx.strokeStyle = "#d6def2";
    ctx.strokeRect(0.5, 0.5, labelCanvas.width - 1, labelCanvas.height - 1);
    ctx.drawImage(qrNode, padding, padding, qrSize, qrSize);

    ctx.fillStyle = "#0f172a";
    ctx.font = "bold 22px Arial, sans-serif";
    ctx.textAlign = "center";
    ctx.fillText(materialCode, labelCanvas.width / 2, qrSize + padding + 27);

    return labelCanvas;
}

function waitForImageLoad(img) {
    return new Promise((resolve, reject) => {
        if (img.complete) {
            resolve();
            return;
        }
        img.onload = () => resolve();
        img.onerror = reject;
    });
}

async function createQrLabelCanvasFromCode(materialCode) {
    if (!materialCode || typeof QRCode === "undefined") return null;

    const tempContainer = document.createElement("div");
    tempContainer.style.position = "fixed";
    tempContainer.style.left = "-99999px";
    tempContainer.style.top = "-99999px";
    document.body.appendChild(tempContainer);

    try {
        new QRCode(tempContainer, {
            text: materialCode,
            width: 220,
            height: 220,
            colorDark: "#0f172a",
            colorLight: "#ffffff",
            correctLevel: QRCode.CorrectLevel.M
        });

        const qrNode = tempContainer.querySelector("canvas, img");
        if (!qrNode) return null;
        if (qrNode.tagName === "IMG") {
            await waitForImageLoad(qrNode);
        }

        const labelCanvas = document.createElement("canvas");
        const qrSize = 220;
        const padding = 20;
        const footerHeight = 38;
        labelCanvas.width = qrSize + padding * 2;
        labelCanvas.height = qrSize + padding * 2 + footerHeight;

        const ctx = labelCanvas.getContext("2d");
        if (!ctx) return null;

        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, labelCanvas.width, labelCanvas.height);
        ctx.strokeStyle = "#d6def2";
        ctx.strokeRect(0.5, 0.5, labelCanvas.width - 1, labelCanvas.height - 1);
        ctx.drawImage(qrNode, padding, padding, qrSize, qrSize);

        ctx.fillStyle = "#0f172a";
        ctx.font = "bold 22px Arial, sans-serif";
        ctx.textAlign = "center";
        ctx.fillText(materialCode, labelCanvas.width / 2, qrSize + padding + 27);

        return labelCanvas;
    } finally {
        tempContainer.remove();
    }
}

function openMaterialDetailsModal(material) {
    if (!material) return;

    activeMaterialForQr = material;
    const stockValue = Number(material.stock) || 0;
    const isLowStock = stockValue <= LOW_STOCK_THRESHOLD;

    const codeEl = document.getElementById("materialDetailsCode");
    const nameEl = document.getElementById("materialDetailsName");
    const rfidEl = document.getElementById("materialDetailsRfid");
    const stockEl = document.getElementById("materialDetailsStock");
    const statusEl = document.getElementById("materialDetailsStatus");
    const createdAtEl = document.getElementById("materialDetailsCreatedAt");

    if (codeEl) codeEl.textContent = material.code || "-";
    if (nameEl) nameEl.textContent = material.name || "-";
    if (rfidEl) rfidEl.textContent = material.rfidTag || "-";
    if (stockEl) stockEl.textContent = String(stockValue);
    if (statusEl) statusEl.textContent = isLowStock ? "LOW STOCK" : "Normal";
    if (createdAtEl) createdAtEl.textContent = material.createdAt || "-";

    renderMaterialQr(material.code || "");
    setModalOpen("materialDetailsModal", true);
}

window.viewMaterialDetails = function (id) {
    const material = state.materials.find((item) => item.id === id);
    if (!material) {
        notify("Material not found.", "error");
        return;
    }
    openMaterialDetailsModal(material);
};

window.downloadMaterialQr = function () {
    if (!activeMaterialForQr?.code) {
        notify("Open material details first.", "error");
        return;
    }

    const labelCanvas = createMaterialQrLabelCanvas(activeMaterialForQr.code);
    if (!labelCanvas) {
        notify("QR label is not ready yet. Please try again.", "error");
        return;
    }

    const link = document.createElement("a");
    link.href = labelCanvas.toDataURL("image/png");
    link.download = `${activeMaterialForQr.code}_qr_label.png`;
    document.body.appendChild(link);
    link.click();
    link.remove();
};

window.printMaterialQr = function () {
    if (!activeMaterialForQr?.code) {
        notify("Open material details first.", "error");
        return;
    }

    const labelCanvas = createMaterialQrLabelCanvas(activeMaterialForQr.code);
    if (!labelCanvas) {
        notify("QR label is not ready yet. Please try again.", "error");
        return;
    }

    const printWindow = window.open("", "_blank", "width=520,height=700");
    if (!printWindow) {
        notify("Please allow pop-ups to print QR labels.", "error");
        return;
    }

    const labelDataUrl = labelCanvas.toDataURL("image/png");
    printWindow.document.write(`
        <!doctype html>
        <html>
        <head>
            <title>${activeMaterialForQr.code} QR Label</title>
            <style>
                @page { size: auto; margin: 10mm; }
                body {
                    margin: 0;
                    font-family: Arial, sans-serif;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    min-height: 100vh;
                    background: #ffffff;
                }
                .label-wrap {
                    text-align: center;
                }
                .label-wrap img {
                    width: 260px;
                    height: auto;
                    border: 1px solid #d6def2;
                }
            </style>
        </head>
        <body>
            <div class="label-wrap">
                <img src="${labelDataUrl}" alt="QR label for ${activeMaterialForQr.code}">
            </div>
            <script>
                window.onload = function() {
                    window.print();
                };
            </script>
        </body>
        </html>
    `);
    printWindow.document.close();
};

window.printAllMaterialQrLabels = async function () {
    if (!requirePermission("print all QR labels", "print_material_qr")) return;
    if (typeof QRCode === "undefined") {
        notify("QR generator failed to load.", "error");
        return;
    }

    const candidates = getVisibleMaterials().filter((material) => String(material.code || "").trim());
    if (!candidates.length) {
        notify("No materials with valid code to print.", "error");
        return;
    }

    const printWindow = window.open("", "_blank", "width=1100,height=800");
    if (!printWindow) {
        notify("Please allow pop-ups to print QR labels.", "error");
        return;
    }

    printWindow.document.write(`
        <!doctype html>
        <html>
        <head>
            <title>All Material QR Labels</title>
            <style>
                @page { size: A4 portrait; margin: 10mm; }
                body {
                    margin: 0;
                    font-family: Arial, sans-serif;
                    background: #ffffff;
                    color: #0f172a;
                }
                .wrap { padding: 8px; }
                h2 {
                    margin: 0 0 10px;
                    font-size: 18px;
                }
                .meta {
                    margin: 0 0 14px;
                    font-size: 12px;
                    color: #475569;
                }
                .grid {
                    display: grid;
                    grid-template-columns: repeat(3, 1fr);
                    gap: 10px;
                }
                .item {
                    border: 1px solid #d6def2;
                    border-radius: 8px;
                    padding: 6px;
                    text-align: center;
                    break-inside: avoid;
                    page-break-inside: avoid;
                }
                .item img {
                    width: 100%;
                    max-width: 240px;
                    height: auto;
                    display: block;
                    margin: 0 auto;
                }
            </style>
        </head>
        <body>
            <div class="wrap">
                <h2>Material QR Labels</h2>
                <p class="meta">Preparing labels...</p>
                <div class="grid" id="allQrGrid"></div>
            </div>
        </body>
        </html>
    `);
    printWindow.document.close();

    const gridEl = printWindow.document.getElementById("allQrGrid");
    const metaEl = printWindow.document.querySelector(".meta");
    if (!gridEl || !metaEl) {
        notify("Unable to build print preview.", "error");
        return;
    }

    let addedCount = 0;
    for (const material of candidates) {
        const code = String(material.code || "").trim();
        if (!code) continue;
        const labelCanvas = await createQrLabelCanvasFromCode(code);
        if (!labelCanvas) continue;

        const item = printWindow.document.createElement("div");
        item.className = "item";
        item.innerHTML = `<img src="${labelCanvas.toDataURL("image/png")}" alt="QR label ${code}">`;
        gridEl.appendChild(item);
        addedCount += 1;
    }

    if (!addedCount) {
        metaEl.textContent = "No printable QR labels generated.";
        notify("No printable QR labels generated.", "error");
        return;
    }

    metaEl.textContent = `Total labels: ${addedCount}`;
    setTimeout(() => {
        printWindow.focus();
        printWindow.print();
    }, 180);
};

function setupModalHandlers() {
    const stockCancel = document.getElementById("stockModalCancelBtn");
    const stockSave = document.getElementById("stockModalSaveBtn");
    const stockAmount = document.getElementById("stockModalAmount");
    const stockPic = document.getElementById("stockModalPic");
    const stockReason = document.getElementById("stockModalReason");
    const stockRemark = document.getElementById("stockModalRemark");
    const editCancel = document.getElementById("editModalCancelBtn");
    const editSave = document.getElementById("editModalSaveBtn");
    const editName = document.getElementById("editModalName");
    const editStock = document.getElementById("editModalStock");
    const editRfidTag = document.getElementById("editModalRfidTag");
    const confirmCancel = document.getElementById("confirmModalCancelBtn");
    const confirmOk = document.getElementById("confirmModalConfirmBtn");
    const materialDetailsClose = document.getElementById("materialDetailsCloseBtn");
    const scanTxType = document.getElementById("scanTxType");
    const scanTxQuantity = document.getElementById("scanTxQuantity");
    const scanTxReason = document.getElementById("scanTxReason");
    const scanTxRemark = document.getElementById("scanTxRemark");
    const scanTxCancel = document.getElementById("scanTxCancelBtn");
    const scanTxSubmit = document.getElementById("scanTxSubmitBtn");

    if (stockAmount) {
        stockAmount.addEventListener("input", () => {
            // Keep digits and at most one decimal separator.
            let value = stockAmount.value.replace(/[^\d.,]/g, "");
            const firstDot = Math.max(value.indexOf("."), value.indexOf(","));
            if (firstDot !== -1) {
                const head = value.slice(0, firstDot + 1);
                const tail = value.slice(firstDot + 1).replace(/[.,]/g, "");
                value = head + tail;
            }
            stockAmount.value = value;
        });
    }

    if (stockCancel) stockCancel.addEventListener("click", () => closeStockModal(null));
    if (stockSave) {
        stockSave.addEventListener("click", () => {
            const rawAmount = (document.getElementById("stockModalAmount")?.value || "").trim();
            const normalizedAmount = rawAmount.replace(",", ".");
            const amount = Number(normalizedAmount);
            const pic = (document.getElementById("stockModalPic")?.value || "").trim();
            if (!Number.isFinite(amount) || amount <= 0) {
                notify("Please enter a valid positive number.", "error");
                return;
            }
            if (!pic) {
                notify("PIC is required.", "error");
                return;
            }
            closeStockModal({
                amount,
                pic,
                reason: normalizeTransactionReason(stockReason?.value, "Other"),
                remark: (stockRemark?.value || "").trim()
            });
        });
    }

    if (editCancel) editCancel.addEventListener("click", () => closeEditModal(null));
    if (editSave) {
        editSave.addEventListener("click", () => {
            const name = (document.getElementById("editModalName")?.value || "").trim();
            const stock = Number(document.getElementById("editModalStock")?.value);
            const rfidTag = normalizeRfidTag(document.getElementById("editModalRfidTag")?.value || "");
            if (!name) {
                notify("Material name cannot be empty.", "error");
                return;
            }
            if (!Number.isFinite(stock) || stock < 0) {
                notify("Stock must be a valid non-negative number.", "error");
                return;
            }
            closeEditModal({ name, stock, rfidTag });
        });
    }

    if (confirmCancel) confirmCancel.addEventListener("click", () => closeConfirmModal(false));
    if (confirmOk) confirmOk.addEventListener("click", () => closeConfirmModal(true));
    if (materialDetailsClose) materialDetailsClose.addEventListener("click", () => closeMaterialDetailsModal());
    if (scanTxCancel) scanTxCancel.addEventListener("click", () => closeScanTxModal(null));
    if (scanTxSubmit) {
        scanTxSubmit.addEventListener("click", () => {
            const type = (scanTxType?.value || "IN").toUpperCase();
            const quantity = Number(scanTxQuantity?.value);
            if (type !== "IN" && type !== "OUT") {
                notify("Please select a valid transaction type.", "error");
                return;
            }
            if (!Number.isFinite(quantity) || quantity <= 0) {
                notify("Quantity must be a positive number.", "error");
                return;
            }
            closeScanTxModal({
                type,
                quantity,
                reason: normalizeTransactionReason(scanTxReason?.value, "Other"),
                remark: (scanTxRemark?.value || "").trim()
            });
        });
    }

    [stockAmount, stockPic, stockReason, stockRemark].forEach((el) => {
        if (!el || !stockSave) return;
        el.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                stockSave.click();
            }
        });
    });

    [editName, editStock, editRfidTag].forEach((el) => {
        if (!el || !editSave) return;
        el.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                editSave.click();
            }
        });
    });

    [scanTxType, scanTxQuantity, scanTxReason, scanTxRemark].forEach((el) => {
        if (!el || !scanTxSubmit) return;
        el.addEventListener("keydown", (event) => {
            if (event.key === "Enter") {
                event.preventDefault();
                scanTxSubmit.click();
            }
        });
    });

    ["stockModal", "editModal", "confirmModal", "materialDetailsModal", "scanTxModal"].forEach((modalId) => {
        const modal = document.getElementById(modalId);
        if (!modal) return;
        modal.addEventListener("click", (event) => {
            if (event.target !== modal) return;
            if (modalId === "stockModal") closeStockModal(null);
            if (modalId === "editModal") closeEditModal(null);
            if (modalId === "confirmModal") closeConfirmModal(false);
            if (modalId === "materialDetailsModal") closeMaterialDetailsModal();
            if (modalId === "scanTxModal") closeScanTxModal(null);
        });
    });

    document.addEventListener("keydown", (event) => {
        if (event.key !== "Escape") return;
        if (document.getElementById("stockModal")?.classList.contains("open")) closeStockModal(null);
        if (document.getElementById("editModal")?.classList.contains("open")) closeEditModal(null);
        if (document.getElementById("confirmModal")?.classList.contains("open")) closeConfirmModal(false);
        if (document.getElementById("materialDetailsModal")?.classList.contains("open")) closeMaterialDetailsModal();
        if (document.getElementById("scanTxModal")?.classList.contains("open")) closeScanTxModal(null);
    });
}

function normalizeRole(value) {
    if (typeof value !== "string") return null;

    const normalized = value.trim().toLowerCase().replace(/[_\s-]+/g, "");

    if (normalized === "manager") return ROLES.MANAGER;
    if (normalized === "sitesupervisor") return ROLES.SITE_SUPERVISOR;
    if (normalized === "storekeeper") return ROLES.STOREKEEPER;

    return null;
}

function hasPermission(permission) {
    if (!state.user || !state.role) return false;

    if (permission === "report_access") {
        return state.role === ROLES.MANAGER || state.role === ROLES.STOREKEEPER;
    }

    if (permission === "export_reports") {
        return state.role === ROLES.MANAGER || state.role === ROLES.STOREKEEPER;
    }

    if (permission === "print_material_qr") return state.role === ROLES.STOREKEEPER;

    if (permission === "add_material") return state.role === ROLES.STOREKEEPER;
    if (permission === "update_stock") return state.role === ROLES.STOREKEEPER;
    if (permission === "edit_material") return state.role === ROLES.STOREKEEPER;
    if (permission === "delete_material") return state.role === ROLES.STOREKEEPER;

    return false;
}

function requirePermission(actionLabel, permission) {
    if (!state.user) {
        notify(`Please login to ${actionLabel}.`, "error");
        return false;
    }

    if (!state.role) {
        notify("Your account has no role assigned. Please contact administrator.", "error");
        return false;
    }

    if (!hasPermission(permission)) {
        notify(`Access denied: ${state.role} cannot ${actionLabel}.`, "error");
        return false;
    }

    return true;
}

function setControlEnabledById(id, enabled) {
    const el = document.getElementById(id);
    if (el) el.disabled = !enabled;
}

function setControlEnabledByOnclick(onclickValue, enabled) {
    const controls = document.querySelectorAll(`button[onclick=\"${onclickValue}\"]`);
    controls.forEach((el) => {
        el.disabled = !enabled;
    });
}

function setPermissionVisibility(permission, visible) {
    const targets = document.querySelectorAll(`.permission-gated[data-permission="${permission}"]`);
    targets.forEach((el) => {
        el.classList.toggle("hidden-by-permission", !visible);
    });
}

function applyRoleLayout() {
    const layout = getRoleLayout(state.role);
    const visibleSections = new Set(layout.visibleSections);
    const sidebarLinks = document.querySelectorAll(".sidebar-link");

    sidebarLinks.forEach((link) => {
        const sectionId = link.dataset.section;
        const isVisible = visibleSections.has(sectionId);
        link.hidden = !isVisible;
        link.classList.toggle("hidden-by-permission", !isVisible);
        if (!isVisible) link.classList.remove("active");
    });

    ["overviewSection", "materialsSection", "transactionsSection", "scannerSection", "chartsSection", "settingsSection", "reportsSection"].forEach((id) => {
        const section = document.getElementById(id);
        if (section) section.hidden = !visibleSections.has(id);
    });

    const addMaterialCard = document.getElementById("addMaterialCard");
    if (addMaterialCard) addMaterialCard.hidden = !layout.showAddMaterial;

    const actionHeader = document.getElementById("materialActionHeader");
    if (actionHeader) {
        actionHeader.textContent = layout.materialActionHeader;
    }

    const firstVisibleLink = Array.from(sidebarLinks).find((link) => !link.hidden);
    const activeVisibleLink = Array.from(sidebarLinks).find((link) => !link.hidden && link.classList.contains("active"));
    if (!activeVisibleLink && firstVisibleLink) {
        firstVisibleLink.classList.add("active");
    }
}

function formatRoleLabel(role) {
    if (role === "admin") return "Admin";
    if (role === ROLES.MANAGER) return "Manager";
    if (role === ROLES.SITE_SUPERVISOR) return "Site Supervisor";
    if (role === ROLES.STOREKEEPER) return "Storekeeper";
    return "Guest";
}

function normalizeTransactionReason(value, fallback = "Other") {
    const normalized = String(value || "").trim();
    const allowedReasons = new Set(["Site Usage", "Waste", "Damaged", "Expired", "Other"]);
    return allowedReasons.has(normalized) ? normalized : fallback;
}

async function resolvePreferredQrCamera() {
    if (typeof Html5Qrcode === "undefined" || typeof Html5Qrcode.getCameras !== "function") {
        return { facingMode: "environment" };
    }

    try {
        const cameras = await Html5Qrcode.getCameras();
        if (!Array.isArray(cameras) || !cameras.length) {
            return { facingMode: "environment" };
        }

        const rearCamera = cameras.find((camera) => /back|rear|environment/i.test(String(camera.label || "")));
        if (rearCamera?.id) return rearCamera.id;

        return cameras[0].id || { facingMode: "environment" };
    } catch {
        return { facingMode: "environment" };
    }
}

function getStatusBadgeClass(status) {
    const normalized = String(status || "").trim().toUpperCase();
    if (normalized === "IN") return "status-badge status-in";
    if (normalized === "OUT") return "status-badge status-out";
    return "status-badge";
}

function getReasonBadgeClass(reason) {
    const normalized = String(reason || "").trim().toLowerCase().replace(/[^a-z]+/g, "-");
    if (!normalized) return "reason-badge";
    return `reason-badge reason-${normalized}`;
}

function updateAuthUI() {
    const status = document.getElementById("authStatus");
    const logoutBtn = document.getElementById("logoutBtn");
    const authUserLabel = document.getElementById("authUserLabel");
    const roleBadge = document.getElementById("roleBadge");
    const sidebarRolePill = document.getElementById("sidebarRolePill");
    const roleFocusPanel = document.getElementById("roleFocusPanel");
    const roleFocusKicker = document.getElementById("roleFocusKicker");
    const roleFocusTitle = document.getElementById("roleFocusTitle");
    const roleFocusText = document.getElementById("roleFocusText");
    const roleFocusTagOne = document.getElementById("roleFocusTagOne");
    const roleFocusTagTwo = document.getElementById("roleFocusTagTwo");
    const roleFocusTagThree = document.getElementById("roleFocusTagThree");

    if (state.user) {
        if (logoutBtn) logoutBtn.disabled = false;
        if (authUserLabel) {
            authUserLabel.textContent = `Signed in user: ${state.user.email}`;
        }

        if (state.role) {
            status.textContent = `Signed in as ${state.user.email} (${formatRoleLabel(state.role)}). Permissions loaded.`;
        } else {
            status.textContent = `Signed in as ${state.user.email}. Role not assigned or still loading.`;
        }
    } else {
        status.textContent = "Session expired. Redirecting to login.";
        if (logoutBtn) logoutBtn.disabled = true;
        if (authUserLabel) {
            authUserLabel.textContent = "Signed in user: -";
        }
    }

    if (roleBadge) {
        roleBadge.className = "role-badge";
        if (state.role) roleBadge.classList.add(state.role);
        roleBadge.textContent = `Role: ${formatRoleLabel(state.role)}`;
    }

    if (sidebarRolePill) {
        sidebarRolePill.className = "sidebar-role-pill";
        if (state.role) sidebarRolePill.classList.add(state.role);
        sidebarRolePill.textContent = formatRoleLabel(state.role);
    }

    if (roleFocusPanel && roleFocusKicker && roleFocusTitle && roleFocusText && roleFocusTagOne && roleFocusTagTwo && roleFocusTagThree) {
        roleFocusPanel.className = "role-focus-panel";

        if (state.role === ROLES.SITE_SUPERVISOR) {
            roleFocusPanel.classList.add(ROLES.SITE_SUPERVISOR);
            roleFocusKicker.textContent = "Supervisor View";
            roleFocusTitle.textContent = "Site Supervisor Monitoring Workspace";
            roleFocusText.textContent = "Review material activity, monitor stock visibility, and supervise operational flow across inventory movement without direct stock editing access.";
            roleFocusTagOne.textContent = "Site Monitoring";
            roleFocusTagTwo.textContent = "Readiness Review";
            roleFocusTagThree.textContent = "Supervisor Session";
        } else if (state.role === ROLES.MANAGER) {
            roleFocusPanel.classList.add(ROLES.MANAGER);
            roleFocusKicker.textContent = "Management View";
            roleFocusTitle.textContent = "Manager Reporting Workspace";
            roleFocusText.textContent = "Track export-ready summaries, review inventory activity trends, and monitor system performance through charts, reports, and history tables.";
            roleFocusTagOne.textContent = "Reports";
            roleFocusTagTwo.textContent = "Inventory Review";
            roleFocusTagThree.textContent = "Manager Session";
        } else if (state.role === ROLES.STOREKEEPER) {
            roleFocusPanel.classList.add(ROLES.STOREKEEPER);
            roleFocusKicker.textContent = "Operations View";
            roleFocusTitle.textContent = "Storekeeper Inventory Workspace";
            roleFocusText.textContent = "Manage material additions, stock updates, QR or RFID input, and daily transaction activity directly from the main inventory dashboard.";
            roleFocusTagOne.textContent = "Material Control";
            roleFocusTagTwo.textContent = "Stock Update";
            roleFocusTagThree.textContent = "Store Session";
        } else {
            roleFocusKicker.textContent = "Operational View";
            roleFocusTitle.textContent = "Inventory Workspace";
            roleFocusText.textContent = "Monitor daily material flow, review activity, and use the dashboard tools based on your assigned role.";
            roleFocusTagOne.textContent = "Live Dashboard";
            roleFocusTagTwo.textContent = "Role-Aware Access";
            roleFocusTagThree.textContent = "Protected Session";
        }
    }

    const canAdd = hasPermission("add_material");
    const canReportAccess = hasPermission("report_access");
    const canExport = hasPermission("export_reports");
    const canPrintLabels = hasPermission("print_material_qr");

    setControlEnabledById("materialName", canAdd);
    setControlEnabledById("materialRfidTag", canAdd);
    setControlEnabledById("materialStock", canAdd);
    setControlEnabledByOnclick("addMaterial()", canAdd);

    setControlEnabledByOnclick("exportMaterialsExcel()", canExport);
    setControlEnabledByOnclick("exportMaterialsPDF()", canExport);
    setControlEnabledByOnclick("printAllMaterialQrLabels()", canPrintLabels);
    setControlEnabledByOnclick("exportTransactionsExcel()", canExport);
    setControlEnabledByOnclick("exportTransactionsPDF()", canExport);

    setPermissionVisibility("report_access", canReportAccess);
    setPermissionVisibility("export_reports", canExport);
    setPermissionVisibility("print_material_qr", canPrintLabels);
    applyRoleLayout();
    updateSettingsProfileUI();
}

function parseCreatedAt(value) {
    if (!value) return 0;
    const direct = Date.parse(value);
    if (!Number.isNaN(direct)) return direct;

    const text = String(value).trim();
    const match = text.match(
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:,\s*|\s+)(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i
    );
    if (!match) return 0;

    let [, dd, mm, yyyy, hh, min, sec, meridiem] = match;
    let hour = Number(hh);
    const minute = Number(min);
    const second = Number(sec || "0");
    const month = Number(mm) - 1;
    const day = Number(dd);
    const year = Number(yyyy);

    if (meridiem) {
        const mer = meridiem.toLowerCase();
        if (mer === "pm" && hour < 12) hour += 12;
        if (mer === "am" && hour === 12) hour = 0;
    }

    const ts = new Date(year, month, day, hour, minute, second).getTime();
    return Number.isNaN(ts) ? 0 : ts;
}

function parseTransactionTime(data) {
    if (data && data.timestampISO) {
        const iso = Date.parse(data.timestampISO);
        if (!Number.isNaN(iso)) return iso;
    }

    return parseCreatedAt(data ? data.timestamp : "");
}

function getSelectedTransactionRange() {
    const fromValue = document.getElementById("txFromDate")?.value || "";
    const toValue = document.getElementById("txToDate")?.value || "";

    const parseDateInputBoundary = (value, endOfDay = false) => {
        if (!value) return null;

        // Format: yyyy-mm-dd (native input value)
        const isoMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (isoMatch) {
            const [, yyyy, mm, dd] = isoMatch;
            const h = endOfDay ? 23 : 0;
            const m = endOfDay ? 59 : 0;
            const s = endOfDay ? 59 : 0;
            const ms = endOfDay ? 999 : 0;
            return new Date(Number(yyyy), Number(mm) - 1, Number(dd), h, m, s, ms).getTime();
        }

        // Format: dd/mm/yyyy (localized or manually typed)
        const dmyMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (dmyMatch) {
            const [, dd, mm, yyyy] = dmyMatch;
            const h = endOfDay ? 23 : 0;
            const m = endOfDay ? 59 : 0;
            const s = endOfDay ? 59 : 0;
            const ms = endOfDay ? 999 : 0;
            return new Date(Number(yyyy), Number(mm) - 1, Number(dd), h, m, s, ms).getTime();
        }

        return Number.NaN;
    };

    const fromTime = fromValue ? parseDateInputBoundary(fromValue, false) : null;
    const toTime = toValue ? parseDateInputBoundary(toValue, true) : null;

    const hasInvalidDate =
        (fromTime !== null && !Number.isFinite(fromTime)) ||
        (toTime !== null && !Number.isFinite(toTime));
    if (hasInvalidDate) {
        notify("Invalid date format. Please use valid From/To dates.", "error");
        return null;
    }

    if (fromTime !== null && toTime !== null && fromTime > toTime) {
        notify("Invalid date range: 'From' date is later than 'To' date.", "error");
        return null;
    }

    return { fromTime, toTime };
}

function getTransactionsForExport() {
    const range = getSelectedTransactionRange();
    if (!range) return null;
    const searchValue = (state.transactionSearch || "").trim().toLowerCase();
    const typeFilterValue = (document.getElementById("txTypeFilter")?.value || "all").toUpperCase();
    const hasTypeFilter = typeFilterValue === "IN" || typeFilterValue === "OUT";

    const hasRange = range.fromTime !== null || range.toTime !== null;

    const filtered = state.transactions.filter((tx) => {
        if (searchValue) {
            const performer = `${tx.performedBy || ""} ${tx.performedByEmail || ""} ${tx.role || ""}`.toLowerCase();
            const code = String(tx.materialCode || "").toLowerCase();
            const type = String(tx.type || "").toLowerCase();
            const reason = String(tx.reason || "").toLowerCase();
            const remark = String(tx.remark || "").toLowerCase();
            const timestamp = String(tx.timestamp || "").toLowerCase();
            if (
                !code.includes(searchValue) &&
                !performer.includes(searchValue) &&
                !type.includes(searchValue) &&
                !reason.includes(searchValue) &&
                !remark.includes(searchValue) &&
                !timestamp.includes(searchValue)
            ) {
                return false;
            }
        }

        if (hasTypeFilter && String(tx.type || "").toUpperCase() !== typeFilterValue) return false;
        if (!hasRange) return true;

        const time = Number(tx._time) || 0;
        if (!time) return false;
        if (range.fromTime !== null && time < range.fromTime) return false;
        if (range.toTime !== null && time > range.toTime) return false;

        return true;
    });

    const direction = state.transactionSortDirection === "asc" ? 1 : -1;
    const sorted = [...filtered].sort((a, b) => {
        if (state.transactionSortKey === "materialCode") {
            return String(a.materialCode || "").localeCompare(String(b.materialCode || "")) * direction;
        }
        if (state.transactionSortKey === "type") {
            return String(a.type || "").localeCompare(String(b.type || "")) * direction;
        }
        if (state.transactionSortKey === "amount") {
            return ((Number(a.amount) || 0) - (Number(b.amount) || 0)) * direction;
        }
        return ((Number(a._time) || 0) - (Number(b._time) || 0)) * direction;
    });

    return sorted;
}

function getTransactionsForCharts() {
    const range = getSelectedTransactionRange();
    if (!range) return [];

    const hasRange = range.fromTime !== null || range.toTime !== null;
    if (!hasRange) return [...state.transactions];

    return state.transactions.filter((tx) => {
        const time = Number(tx._time) || 0;
        if (!time) return false;
        if (range.fromTime !== null && time < range.fromTime) return false;
        if (range.toTime !== null && time > range.toTime) return false;
        return true;
    });
}

function getSortedMaterials(materials) {
    const sorted = [...materials];

    sorted.sort((a, b) => {
        if (state.sort === "oldest") return parseCreatedAt(a.createdAt) - parseCreatedAt(b.createdAt);
        if (state.sort === "stockHigh") return (Number(b.stock) || 0) - (Number(a.stock) || 0);
        if (state.sort === "stockLow") return (Number(a.stock) || 0) - (Number(b.stock) || 0);
        if (state.sort === "nameAZ") return String(a.name || "").localeCompare(String(b.name || ""));
        if (state.sort === "nameZA") return String(b.name || "").localeCompare(String(a.name || ""));

        return parseCreatedAt(b.createdAt) - parseCreatedAt(a.createdAt);
    });

    return sorted;
}

function passesFilter(material) {
    const stock = Number(material.stock) || 0;

    if (state.filter === "low") return stock <= LOW_STOCK_THRESHOLD;
    if (state.filter === "normal") return stock > LOW_STOCK_THRESHOLD;
    return true;
}

function passesSearch(material) {
    const term = state.search;
    if (!term) return true;

    const code = String(material.code || "").toLowerCase();
    const name = String(material.name || "").toLowerCase();
    const rfidTag = String(material.rfidTag || "").toLowerCase();
    return code.includes(term) || name.includes(term) || rfidTag.includes(term);
}

function getVisibleMaterials() {
    return getSortedMaterials(state.materials).filter((material) => passesFilter(material) && passesSearch(material));
}

function escapeCsvValue(value) {
    const text = String(value ?? "");
    return `"${text.replace(/"/g, '""')}"`;
}

function escapeJsSingle(value) {
    return String(value ?? "")
        .replace(/\\/g, "\\\\")
        .replace(/'/g, "\\'");
}

function getActorInfo() {
    const profileName = state.userProfile?.name;
    const profileEmail = state.userProfile?.email;
    const displayName = profileName || state.user?.displayName || profileEmail || state.user?.email || "Unknown";

    return {
        uid: state.user?.uid || null,
        email: state.user?.email || profileEmail || null,
        name: displayName,
        role: state.role || "Unknown"
    };
}

function formatReportDateTime(date = new Date()) {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    const hh = String(date.getHours()).padStart(2, "0");
    const mi = String(date.getMinutes()).padStart(2, "0");
    const ss = String(date.getSeconds()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

function generateReportId(prefix) {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${prefix}-${year}-${month}${day}`;
}

function downloadFile(content, fileName, mimeType) {
    const needsBom = typeof content === "string" && /excel|csv|xml/i.test(String(mimeType || ""));
    const blob = new Blob(needsBom ? ["\uFEFF", content] : [content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

function buildCsv(headers, rows) {
    const headerLine = headers.map(escapeCsvValue).join(",");
    const rowLines = rows.map((row) => row.map(escapeCsvValue).join(","));
    return [headerLine, ...rowLines].join("\n");
}

function buildCsvWithReportInfo(headers, rows, reportTitle, dateRange = "All Dates") {
    const actor = getActorInfo();
    const info = state.reportInfo || {};
    const metaRows = [
        ["FMAERO SMART TRACKING SYSTEM", ""],
        ["REPORT TITLE", reportTitle],
        ["GENERATED AT", formatReportDateTime(new Date())],
        ["GENERATED BY", actor.name || "-"],
        ["EMAIL", actor.email || "-"],
        ["ROLE", formatRoleLabel(actor.role)],
        ["RECORD COUNT", rows.length],
        ["DATE RANGE", dateRange],
        ["REFERENCE NUMBER", info.referenceNumber || "-"],
        ["ZONE", info.zone || "-"],
        ["PROJECT", info.project || "-"],
        ["DEPARTMENT", info.department || "-"],
        ["PRODUCTION GROUP", info.productionGroup || "-"],
        ["DRAWING NUMBER", info.drawingNumber || "-"],
        ["START DATE", info.startDate || "-"],
        ["FINISH DATE", info.finishDate || "-"],
        ["PIC", info.pic || "-"],
        ["REMARK", info.remark || "-"],
        ["FEEDBACK", info.feedback || "-"]
    ];

    const metaLines = metaRows.map((pair) => pair.map(escapeCsvValue).join(","));
    const tableCsv = buildCsv(headers, rows);
    return [...metaLines, "", "TRANSACTION DATA", tableCsv].join("\n");
}

function escapeXmlValue(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

function buildTransactionsExcelReport(rows, meta) {
    const headers = ["Bil", "Code", "Status", "Reason", "Remark", "Stock", "Created At", "Performed By", "Role"];
    const info = meta.reportInfo || {};
    const generatedAtText = meta.generatedAt || new Date().toLocaleString();
    const reportId = meta.reportId || generateReportId("TX");
    const metaPairs = [
        ["Report ID", reportId],
        ["Generated By", meta.generatedBy || "-"],
        ["Email", meta.generatedEmail || "-"],
        ["Role", meta.generatedRole || "-"],
        ["Record Count", String(meta.recordCount ?? rows.length)],
        ["Date Range", meta.dateRange || "All Dates"],
        ["Reference Number", info.referenceNumber || "-"],
        ["Zone", info.zone || "-"],
        ["Project", info.project || "-"],
        ["Department", info.department || "-"],
        ["Production Group", info.productionGroup || "-"],
        ["Drawing Number", info.drawingNumber || "-"],
        ["Start Date", info.startDate || "-"],
        ["Finish Date", info.finishDate || "-"],
        ["PIC", info.pic || "-"],
        ["Remark", info.remark || "-"],
        ["Feedback", info.feedback || "-"]
    ];

    const stockOut = rows.reduce((sum, row) => sum + (String(row[2]).toUpperCase() === "OUT" ? (Number(row[5]) || 0) : 0), 0);
    const stockIn = rows.reduce((sum, row) => sum + (String(row[2]).toUpperCase() === "IN" ? (Number(row[5]) || 0) : 0), 0);

    const colWidths = [44, 92, 70, 108, 210, 60, 138, 120, 92];

    let currentRow = 1;
    const rowXml = [];

    const pushMergedRow = (text, styleId, mergeAcross = 8, height = null) => {
        const h = height ? ` ss:Height="${height}"` : "";
        rowXml.push(
            `<Row${h}><Cell ss:MergeAcross="${mergeAcross}" ss:StyleID="${styleId}"><Data ss:Type="String">${escapeXmlValue(text)}</Data></Cell></Row>`
        );
        currentRow += 1;
    };

    const pushMetaRow = (label, value) => {
        rowXml.push(
            `<Row ss:Height="20"><Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">${escapeXmlValue(label)}</Data></Cell><Cell ss:MergeAcross="3" ss:StyleID="sMetaValue"><Data ss:Type="String">${escapeXmlValue(value)}</Data></Cell></Row>`
        );
        currentRow += 1;
    };

    pushMergedRow("FMAERO SMART TRACKING SYSTEM", "sTitle", 8, 24);
    pushMergedRow("Transactions Export", "sSubtitle", 8, 20);
    pushMergedRow(`Generated: ${generatedAtText}`, "sGenerated", 8, 18);
    rowXml.push("<Row ss:Height=\"10\"></Row>");
    currentRow += 1;

    metaPairs.forEach(([label, value]) => pushMetaRow(label, value));
    rowXml.push("<Row ss:Height=\"10\"></Row>");
    currentRow += 1;

    pushMergedRow("TRANSACTION DATA", "sSection", 8, 20);
    const tableHeaderRow = currentRow;

    rowXml.push(
        `<Row ss:Height="20">${headers
            .map((h) => `<Cell ss:StyleID="sHeader"><Data ss:Type="String">${escapeXmlValue(h)}</Data></Cell>`)
            .join("")}</Row>`
    );
    currentRow += 1;

    rows.forEach((row) => {
        const statusText = String(row[2] ?? "").toUpperCase();
        const statusStyle = statusText === "IN" ? "sStatusIn" : statusText === "OUT" ? "sStatusOut" : "sCellCenter";
        const reasonText = String(row[3] ?? "");
        const reasonStyle = reasonText === "Waste" ? "sReasonWaste" : "sCellCenter";
        const stockNumber = Number(row[5]) || 0;
        rowXml.push(
            `<Row ss:Height="19">` +
                `<Cell ss:StyleID="sCellCenter"><Data ss:Type="Number">${Number(row[0]) || 0}</Data></Cell>` +
                `<Cell ss:StyleID="sCellCode"><Data ss:Type="String">${escapeXmlValue(row[1])}</Data></Cell>` +
                `<Cell ss:StyleID="${statusStyle}"><Data ss:Type="String">${escapeXmlValue(row[2])}</Data></Cell>` +
                `<Cell ss:StyleID="${reasonStyle}"><Data ss:Type="String">${escapeXmlValue(row[3])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellText"><Data ss:Type="String">${escapeXmlValue(row[4])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellNumber"><Data ss:Type="Number">${stockNumber}</Data></Cell>` +
                `<Cell ss:StyleID="sCellDate"><Data ss:Type="String">${escapeXmlValue(row[6])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellText"><Data ss:Type="String">${escapeXmlValue(row[7])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellText"><Data ss:Type="String">${escapeXmlValue(row[8])}</Data></Cell>` +
            `</Row>`
        );
        currentRow += 1;
    });

    const tableLastRow = Math.max(tableHeaderRow + rows.length, tableHeaderRow);
    rowXml.push("<Row ss:Height=\"12\"></Row>");
    currentRow += 1;

    rowXml.push(`<Row ss:Height="20"><Cell ss:MergeAcross="8" ss:StyleID="sSection"><Data ss:Type="String">SUMMARY</Data></Cell></Row>`);
    currentRow += 1;

    rowXml.push(
        `<Row ss:Height="19"><Cell ss:StyleID="sSummaryLabel"><Data ss:Type="String">Total Transactions</Data></Cell><Cell ss:StyleID="sSummaryValue"><Data ss:Type="Number">${rows.length}</Data></Cell></Row>`
    );
    currentRow += 1;
    rowXml.push(
        `<Row ss:Height="19"><Cell ss:StyleID="sSummaryLabel"><Data ss:Type="String">Total Stock OUT</Data></Cell><Cell ss:StyleID="sSummaryValue"><Data ss:Type="Number">${stockOut}</Data></Cell></Row>`
    );
    currentRow += 1;
    rowXml.push(
        `<Row ss:Height="19"><Cell ss:StyleID="sSummaryLabel"><Data ss:Type="String">Total Stock IN</Data></Cell><Cell ss:StyleID="sSummaryValue"><Data ss:Type="Number">${stockIn}</Data></Cell></Row>`
    );
    currentRow += 1;

    const columnXml = colWidths.map((w) => `<Column ss:AutoFitWidth="1" ss:Width="${w.toFixed(0)}"/>`).join("");
    const autoFilterRange = `R${tableHeaderRow}C1:R${tableLastRow}C9`;
    const freezeAt = tableHeaderRow;

    return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
  </Style>
  <Style ss:ID="sTitle">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="16" ss:Bold="1" ss:Color="#000000"/>
  </Style>
  <Style ss:ID="sSubtitle">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="12" ss:Bold="1" ss:Color="#000000"/>
  </Style>
  <Style ss:ID="sGenerated">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
  </Style>
  <Style ss:ID="sSection">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#000000"/>
   <Interior ss:Color="#DCE6F8" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sMetaLabel">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#000000"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
   <Interior ss:Color="#EEF3FC" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="sMetaValue">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sHeader">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#000000"/>
   <Interior ss:Color="#E6E6E6" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sCellCenter">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sCellText">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
   <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sCellDate">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sCellNumber">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sCellCode">
   <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sStatusIn">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#2ECC71" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sStatusOut">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#E74C3C" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sReasonWaste">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#F39C12" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sStockLow">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
   <Interior ss:Color="#FFEBD6" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sSummaryLabel">
   <Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#000000"/>
   <Interior ss:Color="#EEF3FC" ss:Pattern="Solid"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
  <Style ss:ID="sSummaryValue">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/>
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
   </Borders>
  </Style>
 </Styles>
 <Worksheet ss:Name="Transactions Export">
  <Table ss:ExpandedColumnCount="9" ss:ExpandedRowCount="${currentRow}">
   ${columnXml}
   ${rowXml.join("")}
  </Table>
  <AutoFilter x:Range="${autoFilterRange}" xmlns="urn:schemas-microsoft-com:office:excel"/>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <FreezePanes/>
   <FrozenNoSplit/>
   <SplitHorizontal>${freezeAt}</SplitHorizontal>
   <TopRowBottomPane>${freezeAt}</TopRowBottomPane>
   <ActivePane>2</ActivePane>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>`;
}

function buildMaterialsExcelReport(rows, meta) {
    const headers = ["Bil", "Code", "RFID Tag", "Materials", "Stock", "Status", "Created At", "Created By", "Updated At", "Updated By"];
    const info = meta.reportInfo || {};
    const generatedAtText = meta.generatedAt || new Date().toLocaleString();
    const reportId = meta.reportId || generateReportId("MT");
    const metaPairs = [
        ["Report ID", reportId],
        ["Generated By", meta.generatedBy || "-"],
        ["Email", meta.generatedEmail || "-"],
        ["Role", meta.generatedRole || "-"],
        ["Record Count", String(meta.recordCount ?? rows.length)],
        ["Date Range", meta.dateRange || "All Dates"],
        ["Reference Number", info.referenceNumber || "-"],
        ["Zone", info.zone || "-"],
        ["Project", info.project || "-"],
        ["Department", info.department || "-"],
        ["Production Group", info.productionGroup || "-"],
        ["Drawing Number", info.drawingNumber || "-"],
        ["Start Date", info.startDate || "-"],
        ["Finish Date", info.finishDate || "-"],
        ["PIC", info.pic || "-"],
        ["Remark", info.remark || "-"],
        ["Feedback", info.feedback || "-"]
    ];

    const lowStockCount = rows.reduce((sum, row) => sum + ((Number(row[4]) || 0) <= LOW_STOCK_THRESHOLD ? 1 : 0), 0);
    const totalStock = rows.reduce((sum, row) => sum + (Number(row[4]) || 0), 0);

    const colWidths = [44, 88, 118, 198, 60, 92, 132, 138, 132, 138];

    let currentRow = 1;
    const rowXml = [];

    const pushMergedRow = (text, styleId, mergeAcross = 9, height = null) => {
        const h = height ? ` ss:Height="${height}"` : "";
        rowXml.push(`<Row${h}><Cell ss:MergeAcross="${mergeAcross}" ss:StyleID="${styleId}"><Data ss:Type="String">${escapeXmlValue(text)}</Data></Cell></Row>`);
        currentRow += 1;
    };

    const pushMetaRow = (label, value) => {
        rowXml.push(
            `<Row ss:Height="20"><Cell ss:StyleID="sMetaLabel"><Data ss:Type="String">${escapeXmlValue(label)}</Data></Cell><Cell ss:MergeAcross="4" ss:StyleID="sMetaValue"><Data ss:Type="String">${escapeXmlValue(value)}</Data></Cell></Row>`
        );
        currentRow += 1;
    };

    pushMergedRow("FMAERO SMART TRACKING SYSTEM", "sTitle", 9, 24);
    pushMergedRow("Materials Export", "sSubtitle", 9, 20);
    pushMergedRow(`Generated: ${generatedAtText}`, "sGenerated", 9, 18);
    rowXml.push("<Row ss:Height=\"10\"></Row>");
    currentRow += 1;

    metaPairs.forEach(([label, value]) => pushMetaRow(label, value));
    rowXml.push("<Row ss:Height=\"10\"></Row>");
    currentRow += 1;

    pushMergedRow("MATERIAL DATA", "sSection", 9, 20);
    const tableHeaderRow = currentRow;
    rowXml.push(
        `<Row ss:Height="20">${headers.map((h) => `<Cell ss:StyleID="sHeader"><Data ss:Type="String">${escapeXmlValue(h)}</Data></Cell>`).join("")}</Row>`
    );
    currentRow += 1;

    rows.forEach((row) => {
        const stock = Number(row[4]) || 0;
        const stockStyle = stock <= LOW_STOCK_THRESHOLD ? "sStockLow" : "sCellNumber";
        const statusStyle = String(row[5]).toUpperCase().includes("LOW") ? "sStatusOut" : "sStatusIn";
        rowXml.push(
            `<Row ss:Height="19">` +
                `<Cell ss:StyleID="sCellCenter"><Data ss:Type="Number">${Number(row[0]) || 0}</Data></Cell>` +
                `<Cell ss:StyleID="sCellCode"><Data ss:Type="String">${escapeXmlValue(row[1])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellCode"><Data ss:Type="String">${escapeXmlValue(row[2])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellText"><Data ss:Type="String">${escapeXmlValue(row[3])}</Data></Cell>` +
                `<Cell ss:StyleID="${stockStyle}"><Data ss:Type="Number">${stock}</Data></Cell>` +
                `<Cell ss:StyleID="${statusStyle}"><Data ss:Type="String">${escapeXmlValue(row[5])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellDate"><Data ss:Type="String">${escapeXmlValue(row[6])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellText"><Data ss:Type="String">${escapeXmlValue(row[7])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellDate"><Data ss:Type="String">${escapeXmlValue(row[8])}</Data></Cell>` +
                `<Cell ss:StyleID="sCellText"><Data ss:Type="String">${escapeXmlValue(row[9])}</Data></Cell>` +
            `</Row>`
        );
        currentRow += 1;
    });

    const tableLastRow = Math.max(tableHeaderRow + rows.length, tableHeaderRow);
    rowXml.push("<Row ss:Height=\"12\"></Row>");
    currentRow += 1;
    rowXml.push(`<Row ss:Height="20"><Cell ss:MergeAcross="9" ss:StyleID="sSection"><Data ss:Type="String">SUMMARY</Data></Cell></Row>`);
    currentRow += 1;
    rowXml.push(`<Row ss:Height="19"><Cell ss:StyleID="sSummaryLabel"><Data ss:Type="String">Total Materials</Data></Cell><Cell ss:StyleID="sSummaryValue"><Data ss:Type="Number">${rows.length}</Data></Cell></Row>`);
    currentRow += 1;
    rowXml.push(`<Row ss:Height="19"><Cell ss:StyleID="sSummaryLabel"><Data ss:Type="String">Total Stock</Data></Cell><Cell ss:StyleID="sSummaryValue"><Data ss:Type="Number">${totalStock}</Data></Cell></Row>`);
    currentRow += 1;
    rowXml.push(`<Row ss:Height="19"><Cell ss:StyleID="sSummaryLabel"><Data ss:Type="String">Low Stock (&lt;= ${LOW_STOCK_THRESHOLD})</Data></Cell><Cell ss:StyleID="sSummaryValue"><Data ss:Type="Number">${lowStockCount}</Data></Cell></Row>`);
    currentRow += 1;

    const columnXml = colWidths.map((w) => `<Column ss:AutoFitWidth="1" ss:Width="${w.toFixed(0)}"/>`).join("");
    const autoFilterRange = `R${tableHeaderRow}C1:R${tableLastRow}C10`;
    const freezeAt = tableHeaderRow;

    return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/></Style>
  <Style ss:ID="sTitle"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="16" ss:Bold="1" ss:Color="#000000"/></Style>
  <Style ss:ID="sSubtitle"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="12" ss:Bold="1" ss:Color="#000000"/></Style>
  <Style ss:ID="sGenerated"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/></Style>
  <Style ss:ID="sSection"><Alignment ss:Horizontal="Left" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#000000"/><Interior ss:Color="#DCE6F8" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sMetaLabel"><Alignment ss:Horizontal="Left" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders><Interior ss:Color="#EEF3FC" ss:Pattern="Solid"/></Style>
  <Style ss:ID="sMetaValue"><Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sHeader"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#000000"/><Interior ss:Color="#E6E6E6" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sCellCenter"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sCellText"><Alignment ss:Horizontal="Left" ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sCellCode"><Alignment ss:Horizontal="Left" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sCellDate"><Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sCellNumber"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sStatusIn"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Interior ss:Color="#E8F5E9" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sStatusOut"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Interior ss:Color="#FDECEC" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sStockLow"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Interior ss:Color="#FFEBD6" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sSummaryLabel"><Font ss:FontName="Arial" ss:Size="11" ss:Bold="1" ss:Color="#000000"/><Interior ss:Color="#EEF3FC" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
  <Style ss:ID="sSummaryValue"><Alignment ss:Horizontal="Center" ss:Vertical="Center"/><Font ss:FontName="Arial" ss:Size="11" ss:Color="#000000"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/><Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/></Borders></Style>
 </Styles>
 <Worksheet ss:Name="Materials Export">
  <Table ss:ExpandedColumnCount="10" ss:ExpandedRowCount="${currentRow}">
   ${columnXml}
   ${rowXml.join("")}
  </Table>
  <AutoFilter x:Range="${autoFilterRange}" xmlns="urn:schemas-microsoft-com:office:excel"/>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <FreezePanes/><FrozenNoSplit/>
   <SplitHorizontal>${freezeAt}</SplitHorizontal>
   <TopRowBottomPane>${freezeAt}</TopRowBottomPane>
   <ActivePane>2</ActivePane>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>
</Workbook>`;
}

function renderReportActivityLog() {
    const tbody = document.querySelector("#reportLogTable tbody");
    const emptyEl = document.getElementById("reportLogEmpty");
    if (!tbody) return;

    tbody.innerHTML = "";
    const items = [...state.exportLogs].slice(0, 10);

    items.forEach((log, index) => {
        const row = `
            <tr>
                <td>${index + 1}</td>
                <td>${log.time || "-"}</td>
                <td>${log.report || "-"}</td>
                <td>${log.format || "-"}</td>
                <td>${log.rows ?? 0}</td>
                <td>${log.by || "-"}</td>
            </tr>
        `;
        tbody.innerHTML += row;
    });

    if (emptyEl) emptyEl.hidden = items.length > 0;
}

function updateReportsPanel() {
    const totalMaterialsEl = document.getElementById("reportTotalMaterials");
    const totalTransactionsEl = document.getElementById("reportTotalTransactions");
    const rangeOutEl = document.getElementById("reportRangeOut");
    const lowStockEl = document.getElementById("reportLowStock");

    const filteredTransactions = getTransactionsForExport() || [];
    const rangeOut = filteredTransactions
        .filter((tx) => tx.type === "OUT")
        .reduce((sum, tx) => sum + (Number(tx.amount) || 0), 0);
    const lowStock = state.materials.filter((m) => (Number(m.stock) || 0) <= LOW_STOCK_THRESHOLD).length;

    if (totalMaterialsEl) totalMaterialsEl.textContent = String(state.materials.length);
    if (totalTransactionsEl) totalTransactionsEl.textContent = String(state.transactions.length);
    if (rangeOutEl) rangeOutEl.textContent = String(rangeOut);
    if (lowStockEl) lowStockEl.textContent = String(lowStock);

    renderReportActivityLog();
}

function addReportLog(report, format, rowCount) {
    const actor = getActorInfo();
    state.exportLogs.unshift({
        time: new Date().toLocaleString(),
        report,
        format,
        rows: rowCount,
        by: actor.name || actor.email || "-"
    });
    state.exportLogs = state.exportLogs.slice(0, 30);
    updateReportsPanel();
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function getTransactionRangeLabel() {
    const fromValue = document.getElementById("txFromDate")?.value || "";
    const toValue = document.getElementById("txToDate")?.value || "";
    if (!fromValue && !toValue) return "All Dates";
    if (fromValue && toValue) return `${fromValue} to ${toValue}`;
    if (fromValue) return `From ${fromValue}`;
    return `Up to ${toValue}`;
}

function buildPdfMeta(reportTitle, rowCount, extra = {}) {
    const actor = getActorInfo();
    return {
        reportTitle,
        generatedAt: formatReportDateTime(new Date()),
        generatedBy: actor.name || "-",
        generatedEmail: actor.email || "-",
        generatedRole: formatRoleLabel(actor.role),
        recordCount: rowCount,
        dateRange: extra.dateRange || "All Dates",
        reportId: extra.reportId || "",
        reportInfo: {
            ...state.reportInfo
        }
    };
}

async function printTableAsPdf(title, headers, rows, meta = {}) {
    const opened = window.open("", "_blank");
    if (!opened) {
        notify("Popup blocked. Please allow popups to export PDF.", "error");
        return;
    }

    const thead = `<tr>${headers.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr>`;
    const tbody = rows
        .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`)
        .join("");
    const logoUrl = await getCompanyLogoForPdf();
    const finalMeta = {
        generatedAt: meta.generatedAt || formatReportDateTime(new Date()),
        generatedBy: meta.generatedBy || "-",
        generatedEmail: meta.generatedEmail || "-",
        generatedRole: meta.generatedRole || "-",
        recordCount: Number.isFinite(meta.recordCount) ? meta.recordCount : rows.length,
        dateRange: meta.dateRange || "All Dates",
        reportInfo: {
            ...DEFAULT_REPORT_INFO,
            ...(meta.reportInfo || {})
        }
    };
    const pdfTitle = meta.reportId ? `${title} ${meta.reportId}` : title;

    if (title === "Transactions Export") {
        const normalizedHeaders = headers.map((h) => (h === "Bil." ? "Bil" : h));
        const headerRowHtml = `<tr>${normalizedHeaders.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr>`;
        const exportedBy = finalMeta.generatedBy && finalMeta.generatedBy !== "-" ? finalMeta.generatedBy : "System User";
        const reportId = meta.reportId || generateReportId("TX");
        const reportVersion = "v1.0";
        const exportModule = "Inventory Tracking";
        const systemReference = "FMAERO-INV";
        const qrText = `${reportId} | Exported By ${exportedBy} | Generated ${finalMeta.generatedAt}`;
        const qrUrl = `https://quickchart.io/qr?size=92&margin=0&text=${encodeURIComponent(qrText)}`;

        const FIRST_PAGE_ROWS = 8;
        const MIDDLE_PAGE_ROWS = 18;
        const LAST_PAGE_ROWS = 8;
        const chunks = [];
        let cursor = 0;
        if (rows.length <= FIRST_PAGE_ROWS) {
            chunks.push(rows);
        } else {
            chunks.push(rows.slice(0, FIRST_PAGE_ROWS));
            cursor = FIRST_PAGE_ROWS;
            while (rows.length - cursor > LAST_PAGE_ROWS) {
                const remaining = rows.length - cursor;
                const take = Math.min(MIDDLE_PAGE_ROWS, remaining - LAST_PAGE_ROWS);
                chunks.push(rows.slice(cursor, cursor + take));
                cursor += take;
            }
            chunks.push(rows.slice(cursor));
        }
        if (!chunks.length) chunks.push([]);

        const systemInfoRows = [
            ["Generated By", finalMeta.generatedBy || "-", "Email", finalMeta.generatedEmail || "-"],
            ["Role", finalMeta.generatedRole || "-", "Record Count", String(finalMeta.recordCount ?? rows.length)],
            ["Date Range", finalMeta.dateRange || "All Dates", "", ""],
            ["Reference Number", finalMeta.reportInfo.referenceNumber || "-", "Zone", finalMeta.reportInfo.zone || "-"],
            ["Project", finalMeta.reportInfo.project || "-", "Department", finalMeta.reportInfo.department || "-"],
            ["Production Group", finalMeta.reportInfo.productionGroup || "-", "Drawing Number", finalMeta.reportInfo.drawingNumber || "-"],
            ["Start Date", finalMeta.reportInfo.startDate || "-", "Finish Date", finalMeta.reportInfo.finishDate || "-"],
            ["PIC", finalMeta.reportInfo.pic || "-", "", ""]
        ];

        const systemInfoHtml = systemInfoRows
            .map((row) => {
                if (!row[2]) {
                    return `<tr><td class="label">${escapeHtml(row[0])}</td><td colspan="3">${escapeHtml(row[1])}</td></tr>`;
                }
                return `<tr><td class="label">${escapeHtml(row[0])}</td><td>${escapeHtml(row[1])}</td><td class="label">${escapeHtml(row[2])}</td><td>${escapeHtml(row[3])}</td></tr>`;
            })
            .join("");

        const tableRowsHtml = (chunk, pageIndex) => {
            if (!chunk.length) {
                return `<tr><td colspan="${normalizedHeaders.length}" class="empty-row">No transaction records.</td></tr>`;
            }
            const base = chunks.slice(0, pageIndex).reduce((sum, rowsInChunk) => sum + rowsInChunk.length, 0);
            return chunk
                .map((row, idx) => {
                    const adjusted = [...row];
                    adjusted[0] = String(base + idx + 1);
                    return `<tr>${adjusted.map((cell, cellIndex) => {
                        if (cellIndex === 2) {
                            const statusClass = String(cell).toUpperCase() === "IN" ? "pdf-badge badge-in" : String(cell).toUpperCase() === "OUT" ? "pdf-badge badge-out" : "pdf-badge";
                            return `<td><span class="${statusClass}">${escapeHtml(cell)}</span></td>`;
                        }
                        if (cellIndex === 3) {
                            const reasonClass = String(cell) === "Waste" ? "pdf-badge badge-waste" : "pdf-badge";
                            return `<td><span class="${reasonClass}">${escapeHtml(cell)}</span></td>`;
                        }
                        return `<td>${escapeHtml(cell)}</td>`;
                    }).join("")}</tr>`;
                })
                .join("");
        };

        const totalPages = chunks.length;
        const pagesHtml = chunks
            .map((chunk, pageIndex) => {
                const isFirst = pageIndex === 0;
                const isLast = pageIndex === totalPages - 1;
                const remarkText = escapeHtml([finalMeta.reportInfo.remark, finalMeta.reportInfo.feedback].filter(Boolean).join(" | ") || "-");

                return `
                    <section class="pdf-page${isLast ? " last-page" : ""}">
                        <div class="watermark">FMAERO SMART TRACKING SYSTEM</div>
                        <header class="report-header">
                            <div class="header-left">
                                <img class="logo" src="${logoUrl}" onerror="this.style.display='none'" />
                            </div>
                            <div class="header-center">
                                <h1>FMAERO Smart Tracking System</h1>
                                <h2>Transactions Export</h2>
                            </div>
                            <div class="header-right">
                                <div>Generated: ${escapeHtml(finalMeta.generatedAt)}</div>
                                <div class="official">Official Export</div>
                            </div>
                        </header>
                        <main class="page-main">
                            ${isFirst ? `
                            <section class="erp-meta">
                                <div><span>Report ID :</span><strong>${reportId}</strong></div>
                                <div><span>Exported By :</span><strong>${escapeHtml(exportedBy)}</strong></div>
                                <div><span>Report Version :</span><strong>${reportVersion}</strong></div>
                                <div><span>Export Module :</span><strong>${exportModule}</strong></div>
                                <div><span>System Reference :</span><strong>${systemReference}</strong></div>
                                <div><span>Generated Date :</span><strong>${escapeHtml(finalMeta.generatedAt)}</strong></div>
                            </section>
                            <section class="sys-info">
                                <table class="sys-info-table">${systemInfoHtml}</table>
                            </section>
                            ` : ""}
                            <section class="table-section">
                                <table class="tx-table">
                                    <thead>${headerRowHtml}</thead>
                                    <tbody>${tableRowsHtml(chunk, pageIndex)}</tbody>
                                </table>
                            </section>
                            ${isLast ? `
                            <section class="report-tail">
                                <div class="remark-box">
                                    <div class="title">Remark / Feedback</div>
                                    <div class="content">${remarkText}</div>
                                </div>
                                <div class="report-footer-area">
                                    <div class="approval-stack">
                                        <div class="approval-box">
                                            <p class="title">Approved By</p>
                                            <p>Name:</p>
                                            <p class="write-space"></p>
                                            <p>Date:</p>
                                            <p class="write-space"></p>
                                            <p>Signature & Company Stamp:</p>
                                            <p class="write-space stamp-space"></p>
                                        </div>
                                        <div class="qr-section">
                                            <div class="verify-box">
                                                <img src="${qrUrl}" alt="Verification QR" />
                                                <div class="verify-label">Verification QR</div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </section>
                            ` : ""}
                        </main>
                        <footer class="report-footer">
                            <div class="left">&copy; 2026 FMAERO Smart Tracking System</div>
                            <div class="right">Page ${pageIndex + 1} of ${totalPages}</div>
                        </footer>
                    </section>
                `;
            })
            .join("");

        opened.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>${escapeHtml(pdfTitle)}</title>
                <style>
                    @page { size: A4; margin: 10mm 15mm 18mm 15mm; }
                    * { box-sizing: border-box; }
                    body {
                        margin: 0;
                        font-family: Arial, sans-serif;
                        font-size: 11pt;
                        color: #000;
                        background: #f2f5fb;
                    }
                    .pdf-page,
                    .pdf-page * {
                        font-family: Arial, sans-serif !important;
                        font-size: 11pt;
                        color: #000 !important;
                    }
                    .pdf-page {
                        position: relative;
                        min-height: calc(297mm - 28mm);
                        background: #fff;
                        display: flex;
                        flex-direction: column;
                        page-break-after: always;
                    }
                    .pdf-page:last-child { page-break-after: auto; }
                    .watermark {
                        position: absolute;
                        left: 50%;
                        top: 48%;
                        transform: translate(-50%, -50%) rotate(-24deg);
                        font-size: 30px;
                        font-weight: 800;
                        letter-spacing: 2px;
                        color: rgba(0, 0, 0, 0.05) !important;
                        white-space: nowrap;
                        pointer-events: none;
                        user-select: none;
                        z-index: 0;
                    }
                    .report-header, .page-main, .report-footer { position: relative; z-index: 1; }
                    .report-header {
                        border: 1px solid #b9caee;
                        border-radius: 10px;
                        padding: 8px 10px;
                        display: grid;
                        grid-template-columns: auto 1fr auto;
                        align-items: center;
                        gap: 12px;
                        background: linear-gradient(180deg, #f8fbff 0%, #eef4ff 100%);
                        box-shadow: 0 8px 18px rgba(30, 58, 138, 0.08);
                    }
                    .logo {
                        width: 62px;
                        height: 62px;
                        object-fit: contain;
                        border-radius: 6px;
                        background: #fff;
                        border: 1px solid #d5dff4;
                    }
                    .header-center h1 {
                        margin: 0;
                        font-size: 22px;
                        color: #1e3a8a;
                        line-height: 1.1;
                        font-weight: 800;
                    }
                    .header-center h2 {
                        margin: 4px 0 0;
                        font-size: 13px;
                        color: #334155;
                        font-weight: 700;
                    }
                    .header-right {
                        text-align: right;
                        font-size: 11px;
                        color: #334155;
                        font-weight: 600;
                        line-height: 1.45;
                        border: 1px solid #c8d7f4;
                        border-radius: 8px;
                        background: #fff;
                        padding: 6px 8px;
                        min-width: 205px;
                        overflow-wrap: anywhere;
                        word-break: break-word;
                    }
                    .header-right .official {
                        color: #1e3a8a;
                        font-size: 13px;
                        font-weight: 800;
                    }
                    .page-main {
                        display: flex;
                        flex-direction: column;
                        gap: 8px;
                        margin-top: 6px;
                        flex: 1;
                    }
                    .erp-meta {
                        border: 1px solid #c4d3f0;
                        border-radius: 8px;
                        background: #f8fbff;
                        display: grid;
                        grid-template-columns: repeat(3, minmax(0, 1fr));
                        gap: 0;
                        overflow: hidden;
                    }
                    .erp-meta div {
                        padding: 6px 8px;
                        border-right: 1px solid #d3dff5;
                        border-bottom: 1px solid #d3dff5;
                    }
                    .erp-meta div:nth-child(3n) { border-right: none; }
                    .erp-meta div:nth-last-child(-n+3) { border-bottom: none; }
                    .erp-meta span {
                        display: block;
                        font-size: 10px;
                        color: #475569;
                        margin-bottom: 1px;
                    }
                    .erp-meta strong {
                        font-size: 11px;
                        color: #0f172a;
                        font-weight: 700;
                    }
                    .sys-info-table, .tx-table {
                        width: 100%;
                        border-collapse: collapse;
                        table-layout: fixed;
                    }
                    .sys-info-table td {
                        border: 1px solid #b8c8ea;
                        padding: 5px 8px;
                        font-size: 10.5px;
                        color: #0f172a;
                        overflow-wrap: anywhere;
                        word-break: break-word;
                    }
                    .sys-info-table .label {
                        background: #edf3ff;
                        width: 24%;
                        font-weight: 700;
                    }
                    .table-section {
                        border: 1px solid #7f96c8;
                        border-radius: 8px;
                        overflow: hidden;
                        box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
                    }
                    .tx-table th,
                    .tx-table td {
                        border: 1px solid #7f96c8;
                        padding: 6px 8px;
                        font-size: 11px;
                        line-height: 1.35;
                        vertical-align: middle;
                        text-align: left;
                        overflow-wrap: anywhere;
                        word-break: break-word;
                    }
                    .tx-table th {
                        background: #e8edf7;
                        color: #0f172a;
                        font-weight: 800;
                    }
                    .tx-table td:nth-child(1),
                    .tx-table td:nth-child(6),
                    .tx-table td:nth-child(9) {
                        text-align: center;
                    }
                    .pdf-badge {
                        display: inline-block;
                        padding: 3px 8px;
                        border-radius: 999px;
                        font-size: 10px;
                        font-weight: 700;
                        color: #1f2937;
                        background: #e5e7eb;
                    }
                    .badge-in {
                        background: #2ecc71;
                        color: #fff;
                    }
                    .badge-out {
                        background: #e74c3c;
                        color: #fff;
                    }
                    .badge-waste {
                        background: #f39c12;
                        color: #fff;
                    }
                    .tx-table tbody tr { height: 30px; }
                    .tx-table tbody tr:nth-child(even) { background: #f8fbff; }
                    .empty-row {
                        text-align: center !important;
                        color: #64748b;
                        font-style: italic;
                    }
                    .report-tail {
                        margin-top: 8px;
                        display: flex;
                        flex-direction: column;
                        gap: 10px;
                    }
                    .remark-box {
                        border: 1px solid #9fb5e2;
                        border-radius: 8px;
                        padding: 8px 10px;
                        min-height: 40px;
                        background: #fbfdff;
                    }
                    .remark-box .title {
                        font-size: 11px;
                        color: #1e3a8a;
                        font-weight: 800;
                        margin-bottom: 4px;
                    }
                    .remark-box .content {
                        font-size: 10.5px;
                        color: #0f172a;
                        line-height: 1.4;
                        white-space: pre-wrap;
                    }
                    .report-footer-area{
                        display: flex;
                        justify-content: flex-end;
                        margin-top: 22px;
                    }
                    .approval-stack {
                        width: 320px;
                        display: flex;
                        flex-direction: column;
                        align-items: flex-end;
                        gap: 10px;
                    }
                    .approval-box{
                        border: 1px solid #1e3a8a;
                        padding: 15px;
                        width: 320px;
                        font-family: Arial;
                        font-size: 11pt;
                        color: #000;
                        border-radius: 8px;
                        background: #fff;
                    }
                    .approval-box p {
                        margin: 8px 0;
                    }
                    .approval-box .title {
                        font-weight: 700;
                    }
                    .approval-box .write-space {
                        min-height: 20px;
                        margin-top: 2px;
                    }
                    .approval-box .stamp-space {
                        min-height: 34px;
                    }
                    .qr-section{
                        text-align: center;
                        width: 100%;
                        display: flex;
                        justify-content: flex-end;
                    }
                    .verify-box {
                        text-align: center;
                        border: 1px solid #1e3a8a;
                        border-radius: 8px;
                        padding: 5px;
                        background: #fff;
                    }
                    .verify-box img {
                        width: 74px;
                        height: 74px;
                        display: block;
                    }
                    .verify-label {
                        font-size: 9px;
                        margin-top: 2px;
                        color: #475569;
                        font-weight: 600;
                    }
                    .report-footer {
                        margin-top: auto;
                        border-top: 1px solid #cbd7ef;
                        padding-top: 5px;
                        display: grid;
                        grid-template-columns: 1fr auto 1fr;
                        align-items: center;
                        font-size: 10.5px;
                        color: #334155;
                    }
                    .report-footer .right {
                        justify-self: center;
                        font-weight: 700;
                    }
                    .report-footer .left {
                        justify-self: start;
                    }

                    @media print {
                        body { background: #fff; }
                        .pdf-page { min-height: calc(297mm - 28mm); }
                        .tx-table thead { display: table-header-group; }
                        .tx-table tr,
                        .sys-info-table tr,
                        .remark-box,
                        .approval-box,
                        .verify-box {
                            page-break-inside: avoid;
                            break-inside: avoid;
                        }
                        .watermark {
                            color: rgba(0, 0, 0, 0.06) !important;
                        }
                        * {
                            -webkit-print-color-adjust: exact;
                            print-color-adjust: exact;
                        }
                    }
                </style>
            </head>
            <body>${pagesHtml}</body>
            </html>
        `);

        opened.document.close();

        const triggerPrint = () => {
            opened.focus();
            opened.print();
        };

        const images = Array.from(opened.document.images || []);
        const pending = images.filter((img) => !img.complete);
        if (!pending.length) {
            setTimeout(triggerPrint, 140);
            return;
        }

        let remaining = pending.length;
        const done = () => {
            remaining -= 1;
            if (remaining <= 0) {
                setTimeout(triggerPrint, 140);
            }
        };
        pending.forEach((img) => {
            img.onload = done;
            img.onerror = done;
        });
        setTimeout(triggerPrint, 1800);
        return;
    }

    if (title === "Materials Export") {
        const normalizedHeaders = headers.map((h) => (h === "Bil." ? "Bil" : h));
        const headerRowHtml = `<tr>${normalizedHeaders.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr>`;
        const exportedBy = finalMeta.generatedBy && finalMeta.generatedBy !== "-" ? finalMeta.generatedBy : "System User";
        const reportId = meta.reportId || generateReportId("MT");
        const reportVersion = "v1.0";
        const exportModule = "Inventory Tracking";
        const systemReference = "FMAERO-INV";
        const qrText = `${reportId} | Exported By ${exportedBy} | Generated ${finalMeta.generatedAt}`;
        const qrUrl = `https://quickchart.io/qr?size=92&margin=0&text=${encodeURIComponent(qrText)}`;

        const FIRST_PAGE_ROWS = 8;
        const MIDDLE_PAGE_ROWS = 16;
        const LAST_PAGE_ROWS = 8;
        const chunks = [];
        let cursor = 0;
        if (rows.length <= FIRST_PAGE_ROWS) {
            chunks.push(rows);
        } else {
            chunks.push(rows.slice(0, FIRST_PAGE_ROWS));
            cursor = FIRST_PAGE_ROWS;
            while (rows.length - cursor > LAST_PAGE_ROWS) {
                const remaining = rows.length - cursor;
                const take = Math.min(MIDDLE_PAGE_ROWS, remaining - LAST_PAGE_ROWS);
                chunks.push(rows.slice(cursor, cursor + take));
                cursor += take;
            }
            chunks.push(rows.slice(cursor));
        }
        if (!chunks.length) chunks.push([]);

        const systemInfoRows = [
            ["Generated By", finalMeta.generatedBy || "-", "Email", finalMeta.generatedEmail || "-"],
            ["Role", finalMeta.generatedRole || "-", "Record Count", String(finalMeta.recordCount ?? rows.length)],
            ["Date Range", finalMeta.dateRange || "All Dates", "", ""],
            ["Reference Number", finalMeta.reportInfo.referenceNumber || "-", "Zone", finalMeta.reportInfo.zone || "-"],
            ["Project", finalMeta.reportInfo.project || "-", "Department", finalMeta.reportInfo.department || "-"],
            ["Production Group", finalMeta.reportInfo.productionGroup || "-", "Drawing Number", finalMeta.reportInfo.drawingNumber || "-"],
            ["Start Date", finalMeta.reportInfo.startDate || "-", "Finish Date", finalMeta.reportInfo.finishDate || "-"],
            ["PIC", finalMeta.reportInfo.pic || "-", "", ""]
        ];

        const systemInfoHtml = systemInfoRows
            .map((row) => {
                if (!row[2]) {
                    return `<tr><td class="label">${escapeHtml(row[0])}</td><td colspan="3">${escapeHtml(row[1])}</td></tr>`;
                }
                return `<tr><td class="label">${escapeHtml(row[0])}</td><td>${escapeHtml(row[1])}</td><td class="label">${escapeHtml(row[2])}</td><td>${escapeHtml(row[3])}</td></tr>`;
            })
            .join("");

        const tableRowsHtml = (chunk, pageIndex) => {
            if (!chunk.length) {
                return `<tr><td colspan="${normalizedHeaders.length}" class="empty-row">No material records.</td></tr>`;
            }
            const base = chunks.slice(0, pageIndex).reduce((sum, rowsInChunk) => sum + rowsInChunk.length, 0);
            return chunk
                .map((row, idx) => {
                    const adjusted = [...row];
                    adjusted[0] = String(base + idx + 1);
                    return `<tr>${adjusted
                        .map((cell, cellIndex) => {
                            if (cellIndex === 4) {
                                const stockValue = Number(cell) || 0;
                                const stockClass = stockValue <= LOW_STOCK_THRESHOLD ? "material-stock-low" : "";
                                return `<td class="${stockClass}">${escapeHtml(cell)}</td>`;
                            }
                            if (cellIndex === 5) {
                                const statusText = String(cell || "-");
                                const isLow = statusText.toUpperCase().includes("LOW");
                                const badgeClass = isLow ? "low" : "normal";
                                return `<td><span class="material-status ${badgeClass}">${escapeHtml(statusText)}</span></td>`;
                            }
                            return `<td>${escapeHtml(cell)}</td>`;
                        })
                        .join("")}</tr>`;
                })
                .join("");
        };

        const totalPages = chunks.length;
        const pagesHtml = chunks
            .map((chunk, pageIndex) => {
                const isFirst = pageIndex === 0;
                const isLast = pageIndex === totalPages - 1;
                const remarkText = escapeHtml([finalMeta.reportInfo.remark, finalMeta.reportInfo.feedback].filter(Boolean).join(" | ") || "-");

                return `
                    <section class="pdf-page${isLast ? " last-page" : ""}">
                        <div class="watermark">FMAERO SMART TRACKING SYSTEM</div>
                        <header class="report-header">
                            <div class="header-left">
                                <img class="logo" src="${logoUrl}" onerror="this.style.display='none'" />
                            </div>
                            <div class="header-center">
                                <h1>FMAERO Smart Tracking System</h1>
                                <h2>Materials Export</h2>
                            </div>
                            <div class="header-right">
                                <div>Generated: ${escapeHtml(finalMeta.generatedAt)}</div>
                                <div class="official">Official Export</div>
                            </div>
                        </header>
                        <main class="page-main">
                            ${isFirst ? `
                            <section class="erp-meta">
                                <div><span>Report ID :</span><strong>${reportId}</strong></div>
                                <div><span>Exported By :</span><strong>${escapeHtml(exportedBy)}</strong></div>
                                <div><span>Report Version :</span><strong>${reportVersion}</strong></div>
                                <div><span>Export Module :</span><strong>${exportModule}</strong></div>
                                <div><span>System Reference :</span><strong>${systemReference}</strong></div>
                                <div><span>Generated Date :</span><strong>${escapeHtml(finalMeta.generatedAt)}</strong></div>
                            </section>
                            <section class="sys-info">
                                <table class="sys-info-table">${systemInfoHtml}</table>
                            </section>
                            ` : ""}
                            <section class="table-section">
                                <table class="mx-table">
                                    <thead>${headerRowHtml}</thead>
                                    <tbody>${tableRowsHtml(chunk, pageIndex)}</tbody>
                                </table>
                            </section>
                            ${isLast ? `
                            <section class="report-tail">
                                <div class="remark-box">
                                    <div class="title">Remark / Feedback</div>
                                    <div class="content">${remarkText}</div>
                                </div>
                                <div class="report-footer-area">
                                    <div class="approval-stack">
                                        <div class="approval-box">
                                            <p class="title">Approved By</p>
                                            <p>Name:</p>
                                            <p class="write-space"></p>
                                            <p>Date:</p>
                                            <p class="write-space"></p>
                                            <p>Signature & Company Stamp:</p>
                                            <p class="write-space stamp-space"></p>
                                        </div>
                                        <div class="qr-section">
                                            <div class="verify-box">
                                                <img src="${qrUrl}" alt="Verification QR" />
                                                <div class="verify-label">Verification QR</div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </section>
                            ` : ""}
                        </main>
                        <footer class="report-footer">
                            <div class="left">&copy; 2026 FMAERO Smart Tracking System</div>
                            <div class="right">Page ${pageIndex + 1} of ${totalPages}</div>
                        </footer>
                    </section>
                `;
            })
            .join("");

        opened.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>${escapeHtml(pdfTitle)}</title>
                <style>
                    @page { size: A4; margin: 10mm 15mm 18mm 15mm; }
                    * { box-sizing: border-box; }
                    body {
                        margin: 0;
                        font-family: Arial, sans-serif;
                        font-size: 11pt;
                        color: #000;
                        background: #f2f5fb;
                    }
                    .pdf-page,
                    .pdf-page * {
                        font-family: Arial, sans-serif !important;
                        font-size: 11pt;
                        color: #000 !important;
                    }
                    .pdf-page {
                        position: relative;
                        min-height: calc(297mm - 28mm);
                        background: #fff;
                        display: flex;
                        flex-direction: column;
                        page-break-after: always;
                    }
                    .pdf-page:last-child { page-break-after: auto; }
                    .watermark {
                        position: absolute;
                        left: 50%;
                        top: 48%;
                        transform: translate(-50%, -50%) rotate(-24deg);
                        font-size: 30px;
                        font-weight: 800;
                        letter-spacing: 2px;
                        color: rgba(0, 0, 0, 0.05) !important;
                        white-space: nowrap;
                        pointer-events: none;
                        user-select: none;
                        z-index: 0;
                    }
                    .report-header, .page-main, .report-footer { position: relative; z-index: 1; }
                    .report-header {
                        border: 1px solid #b9caee;
                        border-radius: 10px;
                        padding: 8px 10px;
                        display: grid;
                        grid-template-columns: auto 1fr auto;
                        align-items: center;
                        gap: 12px;
                        background: linear-gradient(180deg, #f8fbff 0%, #eef4ff 100%);
                        box-shadow: 0 8px 18px rgba(30, 58, 138, 0.08);
                    }
                    .logo {
                        width: 62px;
                        height: 62px;
                        object-fit: contain;
                        border-radius: 6px;
                        background: #fff;
                        border: 1px solid #d5dff4;
                    }
                    .header-center h1 {
                        margin: 0;
                        font-size: 22px;
                        color: #1e3a8a;
                        line-height: 1.1;
                        font-weight: 800;
                    }
                    .header-center h2 {
                        margin: 4px 0 0;
                        font-size: 13px;
                        color: #334155;
                        font-weight: 700;
                    }
                    .header-right {
                        text-align: right;
                        font-size: 11px;
                        color: #334155;
                        font-weight: 600;
                        line-height: 1.45;
                        border: 1px solid #c8d7f4;
                        border-radius: 8px;
                        background: #fff;
                        padding: 6px 8px;
                        min-width: 205px;
                        overflow-wrap: anywhere;
                        word-break: break-word;
                    }
                    .header-right .official {
                        color: #1e3a8a;
                        font-size: 13px;
                        font-weight: 800;
                    }
                    .page-main {
                        display: flex;
                        flex-direction: column;
                        gap: 8px;
                        margin-top: 6px;
                        flex: 1;
                    }
                    .erp-meta {
                        border: 1px solid #c4d3f0;
                        border-radius: 8px;
                        background: #f8fbff;
                        display: grid;
                        grid-template-columns: repeat(3, minmax(0, 1fr));
                        gap: 0;
                        overflow: hidden;
                    }
                    .erp-meta div {
                        padding: 6px 8px;
                        border-right: 1px solid #d3dff5;
                        border-bottom: 1px solid #d3dff5;
                    }
                    .erp-meta div:nth-child(3n) { border-right: none; }
                    .erp-meta div:nth-last-child(-n+3) { border-bottom: none; }
                    .erp-meta span {
                        display: block;
                        font-size: 10px;
                        color: #475569;
                        margin-bottom: 1px;
                    }
                    .erp-meta strong {
                        font-size: 11px;
                        color: #0f172a;
                        font-weight: 700;
                    }
                    .sys-info-table, .mx-table {
                        width: 100%;
                        border-collapse: collapse;
                        table-layout: fixed;
                    }
                    .sys-info-table td {
                        border: 1px solid #b8c8ea;
                        padding: 5px 8px;
                        font-size: 10.5px;
                        color: #0f172a;
                        overflow-wrap: anywhere;
                        word-break: break-word;
                    }
                    .sys-info-table .label {
                        background: #edf3ff;
                        width: 24%;
                        font-weight: 700;
                    }
                    .table-section {
                        border: 1px solid #7f96c8;
                        border-radius: 8px;
                        overflow: hidden;
                        box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
                    }
                    .mx-table th,
                    .mx-table td {
                        border: 1px solid #7f96c8;
                        padding: 5px 7px;
                        font-size: 10.5px;
                        line-height: 1.35;
                        vertical-align: middle;
                        text-align: left;
                        overflow-wrap: anywhere;
                        word-break: break-word;
                    }
                    .mx-table th {
                        background: #e8edf7;
                        color: #0f172a;
                        font-weight: 800;
                        text-transform: uppercase;
                        letter-spacing: 0.2px;
                    }
                    .mx-table td:nth-child(1),
                    .mx-table td:nth-child(5),
                    .mx-table td:nth-child(6) {
                        text-align: center;
                    }
                    .mx-table tbody tr { height: 30px; }
                    .mx-table tbody tr:nth-child(even) { background: #f8fbff; }
                    .material-status {
                        display: inline-block;
                        padding: 2px 8px;
                        border-radius: 999px;
                        font-size: 10px !important;
                        font-weight: 700;
                        border: 1px solid transparent;
                    }
                    .material-status.normal {
                        color: #166534 !important;
                        background: #e8fff1;
                        border-color: #86efac;
                    }
                    .material-status.low {
                        color: #991b1b !important;
                        background: #fff1f1;
                        border-color: #fca5a5;
                    }
                    .material-stock-low {
                        background: #fff6e8;
                        color: #9a3412 !important;
                        font-weight: 700;
                    }
                    .empty-row {
                        text-align: center !important;
                        color: #64748b;
                        font-style: italic;
                    }
                    .report-tail {
                        margin-top: 8px;
                        display: flex;
                        flex-direction: column;
                        gap: 10px;
                    }
                    .remark-box {
                        border: 1px solid #9fb5e2;
                        border-radius: 8px;
                        padding: 8px 10px;
                        min-height: 40px;
                        background: #fbfdff;
                    }
                    .remark-box .title {
                        font-size: 11px;
                        color: #1e3a8a;
                        font-weight: 800;
                        margin-bottom: 4px;
                    }
                    .remark-box .content {
                        font-size: 10.5px;
                        color: #0f172a;
                        line-height: 1.4;
                        white-space: pre-wrap;
                    }
                    .report-footer-area{
                        display: flex;
                        justify-content: flex-end;
                        margin-top: 22px;
                    }
                    .approval-stack {
                        width: 320px;
                        display: flex;
                        flex-direction: column;
                        align-items: flex-end;
                        gap: 10px;
                    }
                    .approval-box{
                        border: 1px solid #1e3a8a;
                        padding: 15px;
                        width: 320px;
                        font-family: Arial;
                        font-size: 11pt;
                        color: #000;
                        border-radius: 8px;
                        background: #fff;
                    }
                    .approval-box p { margin: 8px 0; }
                    .approval-box .title { font-weight: 700; }
                    .approval-box .write-space { min-height: 20px; margin-top: 2px; }
                    .approval-box .stamp-space { min-height: 34px; }
                    .qr-section{
                        text-align: center;
                        width: 100%;
                        display: flex;
                        justify-content: flex-end;
                    }
                    .verify-box {
                        text-align: center;
                        border: 1px solid #1e3a8a;
                        border-radius: 8px;
                        padding: 5px;
                        background: #fff;
                    }
                    .verify-box img {
                        width: 74px;
                        height: 74px;
                        display: block;
                    }
                    .verify-label {
                        font-size: 9px;
                        margin-top: 2px;
                        color: #475569;
                        font-weight: 600;
                    }
                    .report-footer {
                        margin-top: auto;
                        border-top: 1px solid #cbd7ef;
                        padding-top: 5px;
                        display: grid;
                        grid-template-columns: 1fr auto 1fr;
                        align-items: center;
                        font-size: 10.5px;
                        color: #334155;
                    }
                    .report-footer .right { justify-self: center; font-weight: 700; }
                    .report-footer .left { justify-self: start; }

                    @media print {
                        body { background: #fff; }
                        .pdf-page { min-height: calc(297mm - 28mm); }
                        .mx-table thead { display: table-header-group; }
                        .mx-table tr,
                        .sys-info-table tr,
                        .remark-box,
                        .approval-box,
                        .verify-box {
                            page-break-inside: avoid;
                            break-inside: avoid;
                        }
                        .watermark { color: rgba(0, 0, 0, 0.06) !important; }
                        * {
                            -webkit-print-color-adjust: exact;
                            print-color-adjust: exact;
                        }
                    }
                </style>
            </head>
            <body>${pagesHtml}</body>
            </html>
        `);

        opened.document.close();

        const triggerPrint = () => {
            opened.focus();
            opened.print();
        };

        const images = Array.from(opened.document.images || []);
        const pending = images.filter((img) => !img.complete);
        if (!pending.length) {
            setTimeout(triggerPrint, 140);
            return;
        }

        let remaining = pending.length;
        const done = () => {
            remaining -= 1;
            if (remaining <= 0) {
                setTimeout(triggerPrint, 140);
            }
        };
        pending.forEach((img) => {
            img.onload = done;
            img.onerror = done;
        });
        setTimeout(triggerPrint, 1800);
        return;
    }

    opened.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>${escapeHtml(pdfTitle)}</title>
            <style>
                :root {
                    --primary: #1e3a8a;
                    --primary-soft: #eaf0ff;
                    --ink: #0f172a;
                    --muted: #334155;
                    --line: #b8c7e8;
                    --surface: #f8fbff;
                }
                body { font-family: "Segoe UI", Arial, sans-serif; padding: 20px; color: var(--ink); background: #ffffff; }
                .header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; border: 1px solid #bfd0f5; border-radius: 14px; padding: 14px; background: linear-gradient(180deg, #f9fbff 0%, #f1f6ff 100%); box-shadow: 0 6px 18px rgba(30, 58, 138, 0.08); }
                .header-left { display: flex; align-items: center; gap: 12px; }
                .logo { width: 92px; height: 92px; object-fit: contain; border: none; border-radius: 8px; padding: 0; background: #fff; }
                .company-title { margin: 0; font-size: 22px; color: var(--primary); font-weight: 800; letter-spacing: 0.2px; }
                .report-title { margin: 4px 0 0; font-size: 14px; color: var(--muted); font-weight: 700; }
                .generated { font-size: 12px; color: var(--muted); font-weight: 600; text-align: right; background: #fff; border: 1px solid #c9d7f4; border-radius: 10px; padding: 8px 10px; min-width: 210px; line-height: 1.45; }
                .generated { overflow-wrap: anywhere; word-break: break-word; }
                .generated .stamp { color: var(--primary); font-weight: 800; letter-spacing: 0.2px; }
                .print-header-spacer { height: 128px; }
                .meta-box { margin: 10px 0 14px; border: 1px solid #c6d3ef; border-radius: 10px; background: #ffffff; overflow: hidden; }
                .meta-table { width: 100%; border-collapse: collapse; }
                .meta-table td { border: 1px solid #c7d4ef; padding: 7px 10px; font-size: 11px; color: #000; font-family: Arial, sans-serif; overflow-wrap: anywhere; word-break: break-word; }
                .meta-table td.label { width: 160px; background: #edf3ff; color: #0f172a; font-weight: 700; }
                .report-table-wrap { border: 1px solid #7e93c2; border-radius: 10px; overflow: hidden; }
                .report-table { border-collapse: collapse; width: 100%; border: none; }
                .report-table th, .report-table td { border: 1px solid #7e93c2; padding: 9px; text-align: left; color: #000; font-family: Arial, sans-serif; font-size: 11px; overflow-wrap: anywhere; word-break: break-word; }
                .report-table th { background: #dfe9ff; color: #0f172a; font-weight: 800; text-transform: uppercase; letter-spacing: 0.2px; }
                .report-table td { font-weight: 400; }
                .report-table tbody tr:nth-child(even) { background: #f7f9ff; }
                .report-table tbody tr:hover { background: #eef4ff; }
                .remarks-box {
                    margin-top: 14px;
                    border: 1px solid #9bb1e0;
                    border-radius: 10px;
                    padding: 10px 12px;
                    min-height: 70px;
                    background: #fbfdff;
                }
                .remarks-box .label {
                    font-size: 12px;
                    font-weight: 700;
                    margin-bottom: 6px;
                    color: var(--primary);
                }
                .remarks-box .text {
                    font-size: 12px;
                    color: #000;
                    line-height: 1.5;
                    white-space: pre-wrap;
                }
                .approval-signature {
                    margin-top: 24px;
                    display: flex;
                    justify-content: flex-end;
                }
                .approval-box {
                    width: 260px;
                    text-align: left;
                    color: var(--ink);
                    border: 1px solid #c7d4ef;
                    border-radius: 10px;
                    padding: 10px 12px;
                    background: #fff;
                }
                .approval-box .label {
                    font-size: 12px;
                    font-weight: 700;
                    margin-bottom: 26px;
                    color: var(--primary);
                }
                .approval-box .line {
                    border-top: 1px solid #334155;
                    margin-bottom: 6px;
                }
                .approval-box .meta {
                    font-size: 11px;
                    color: var(--muted);
                    line-height: 1.4;
                }
                .pdf-footer {
                    margin-top: 14px;
                    border-top: 1px solid #c2cde6;
                    padding-top: 8px;
                    font-size: 11px;
                    color: var(--muted);
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                }
                .pdf-footer .copyright {
                    font-weight: 600;
                }
                .pdf-footer .page-number {
                    font-weight: 700;
                    color: var(--primary);
                }
                @media print {
                    @page {
                        margin-top: 128px;
                        margin-right: 12px;
                        margin-bottom: 72px;
                        margin-left: 12px;
                        counter-increment: page;
                    }
                    body {
                        margin: 0;
                        padding: 0;
                    }
                    .header {
                        position: fixed;
                        top: 0;
                        left: 12px;
                        right: 12px;
                        margin-bottom: 0;
                        z-index: 999;
                        background: linear-gradient(180deg, #f9fbff 0%, #f1f6ff 100%);
                    }
                    .meta-box {
                        margin-top: 0;
                    }
                    .generated {
                        page-break-inside: avoid;
                    }
                    .report-table thead {
                        display: table-header-group;
                    }
                    .report-table tr,
                    .meta-table tr,
                    .remarks-box,
                    .approval-signature {
                        page-break-inside: avoid;
                    }
                    body {
                        -webkit-print-color-adjust: exact;
                        print-color-adjust: exact;
                    }
                    .pdf-footer {
                        position: fixed;
                        left: 12px;
                        right: 12px;
                        bottom: 8px;
                        margin-top: 0;
                        background: #fff;
                    }
                    .pdf-footer .page-number::after {
                        content: "Page " counter(page, decimal);
                    }
                }
            </style>
        </head>
        <body>
            <div class="header">
                <div class="header-left">
                    <img class="logo" src="${logoUrl}" onerror="this.style.display='none'" />
                    <div>
                        <p class="company-title">${escapeHtml(COMPANY_NAME)}</p>
                        <p class="report-title">${escapeHtml(title)}</p>
                    </div>
                </div>
                <div class="generated">
                    Generated: ${escapeHtml(finalMeta.generatedAt)}<br>
                    <span class="stamp">Official Export</span>
                </div>
            </div>
            <div class="print-header-spacer"></div>
            <div class="meta-box">
                <table class="meta-table">
                    <tr>
                        <td class="label">Generated By</td><td>${escapeHtml(finalMeta.generatedBy)}</td>
                        <td class="label">Email</td><td>${escapeHtml(finalMeta.generatedEmail)}</td>
                    </tr>
                    <tr>
                        <td class="label">Role</td><td>${escapeHtml(finalMeta.generatedRole)}</td>
                        <td class="label">Record Count</td><td>${escapeHtml(finalMeta.recordCount)}</td>
                    </tr>
                    <tr>
                        <td class="label">Date Range</td><td colspan="3">${escapeHtml(finalMeta.dateRange)}</td>
                    </tr>
                    <tr>
                        <td class="label">Reference Number</td><td>${escapeHtml(finalMeta.reportInfo.referenceNumber || "-")}</td>
                        <td class="label">Zone</td><td>${escapeHtml(finalMeta.reportInfo.zone || "-")}</td>
                    </tr>
                    <tr>
                        <td class="label">Project</td><td>${escapeHtml(finalMeta.reportInfo.project || "-")}</td>
                        <td class="label">Department</td><td>${escapeHtml(finalMeta.reportInfo.department || "-")}</td>
                    </tr>
                    <tr>
                        <td class="label">Production Group</td><td>${escapeHtml(finalMeta.reportInfo.productionGroup || "-")}</td>
                        <td class="label">Drawing Number</td><td>${escapeHtml(finalMeta.reportInfo.drawingNumber || "-")}</td>
                    </tr>
                    <tr>
                        <td class="label">Start Date</td><td>${escapeHtml(finalMeta.reportInfo.startDate || "-")}</td>
                        <td class="label">Finish Date</td><td>${escapeHtml(finalMeta.reportInfo.finishDate || "-")}</td>
                    </tr>
                    <tr>
                        <td class="label">PIC</td><td colspan="3">${escapeHtml(finalMeta.reportInfo.pic || "-")}</td>
                    </tr>
                </table>
            </div>
            <div class="report-table-wrap">
                <table class="report-table">
                    <thead>${thead}</thead>
                    <tbody>${tbody}</tbody>
                </table>
            </div>
            <div class="remarks-box">
                <div class="label">Remark / Feedback</div>
                <div class="text">${escapeHtml([finalMeta.reportInfo.remark, finalMeta.reportInfo.feedback].filter(Boolean).join(" | ") || "-")}</div>
            </div>
            <div class="approval-signature">
                <div class="approval-box">
                    <div class="label">Approved By:</div>
                    <div class="line"></div>
                    <div class="meta">Name:</div>
                    <div class="meta">Date:</div>
                    <div class="meta">Signature & Stamp:</div>
                </div>
            </div>
            <div class="pdf-footer">
                <span class="copyright">&copy; ${new Date().getFullYear()} FMAERO Smart Tracking System</span>
                <span class="page-number"></span>
            </div>
        </body>
        </html>
    `);

    opened.document.close();

    const triggerPrint = () => {
        opened.focus();
        opened.print();
    };

    const headerEl = opened.document.querySelector(".header");
    const spacerEl = opened.document.querySelector(".print-header-spacer");
    if (headerEl && spacerEl) {
        const headerHeight = Math.ceil(headerEl.getBoundingClientRect().height);
        spacerEl.style.height = `${headerHeight + 6}px`;
    }

    const logoImg = opened.document.querySelector(".logo");
    if (logoImg && !logoImg.complete) {
        logoImg.onload = () => setTimeout(triggerPrint, 120);
        logoImg.onerror = () => setTimeout(triggerPrint, 120);
        setTimeout(triggerPrint, 1500);
    } else {
        setTimeout(triggerPrint, 120);
    }
}

function renderMaterialTable() {
    const tableBody = document.querySelector("#materialTable tbody");
    tableBody.innerHTML = "";

    const filtered = getVisibleMaterials();
    const canStock = hasPermission("update_stock");
    const canEdit = hasPermission("edit_material");
    const canDelete = hasPermission("delete_material");

    filtered.forEach((material, index) => {
        const isLowStock = (Number(material.stock) || 0) <= LOW_STOCK_THRESHOLD;
        const safeCode = escapeJsSingle(material.code);
        const safeId = escapeJsSingle(material.id);
        const actions = [`<button onclick="viewMaterialDetails('${safeId}')">View</button>`];

        if (canStock) {
            actions.push(`<button onclick="increaseStock('${safeId}', '${safeCode}')">+ Stock</button>`);
            actions.push(`<button onclick="decreaseStock('${safeId}', '${safeCode}')">- Stock</button>`);
        }
        if (canEdit) {
            actions.push(`<button onclick="editMaterial('${safeId}')">Edit</button>`);
        }
        if (canDelete) {
            actions.push(`<button onclick="deleteMaterial('${safeId}')">Delete</button>`);
        }

        const row = `
            <tr>
                <td>${index + 1}</td>
                <td>${material.code || "-"}</td>
                <td>${material.rfidTag || "-"}</td>
                <td>${material.name || "-"}</td>
                <td>${material.stock ?? 0}</td>
                <td>${isLowStock ? '<span style="color:red;font-weight:bold;">LOW STOCK</span>' : '<span style="color:green;">Normal</span>'}</td>
                <td>${material.createdAt || "-"}</td>
                <td>${actions.join("")}</td>
            </tr>
        `;

        tableBody.innerHTML += row;
    });
}

async function handleScannedCode(rawCode, source = "manual", options = {}) {
    if (isScanFlowActive) return;

    const code = String(rawCode || "").trim();
    if (!code) return;

    if (source === "qr") {
        const nowMs = Date.now();
        if (lastQrScanCode.toLowerCase() === code.toLowerCase() && nowMs - lastQrScanAt < QR_SCAN_COOLDOWN_MS) {
            return;
        }
        lastQrScanCode = code;
        lastQrScanAt = nowMs;
    }

    isScanFlowActive = true;

    try {
        const qrResultEl = document.getElementById("qrLastResult");
        if (qrResultEl) qrResultEl.textContent = `Last QR: ${code}`;

        const material = findMaterialByScanCode(code);
        if (!material) {
            notify(`Scanned RFID/code '${code}' not found in materials.`, "error");
            return;
        }

        const txData = await openScanTxModal(material, options);
        if (!txData) return;

        const actor = getActorInfo();
        const now = new Date();
        await recordInventoryMovement(material, txData.type, txData.quantity, actor.name, txData.reason, txData.remark);

        resetScanTxForm();
        updateLastScanSummary(material.code, txData.type, txData.quantity, now);
        updateOverviewLastFlow(
            material.code,
            txData.type,
            txData.quantity,
            now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }),
            actor.name
        );
        notify(
            `Material: ${material.code} | Action: ${txData.type} | Quantity: ${txData.quantity} | Transaction saved successfully.`,
            "success"
        );
    } catch (error) {
        notify(`Scan transaction failed: ${error.message || error}`, "error");
    } finally {
        isScanFlowActive = false;
    }
}

window.processRfidInput = async function () {
    if (!requirePermission("scan RFID", "update_stock")) return;
    if (rfidAutoSubmitTimer) {
        clearTimeout(rfidAutoSubmitTimer);
        rfidAutoSubmitTimer = null;
    }
    const rfidInput = document.getElementById("rfidInput");
    const rfidResultEl = document.getElementById("rfidLastResult");
    const value = normalizeRfidTag(rfidInput?.value || "");
    if (!value) {
        notify("Please enter RFID tag or material code.", "error");
        return;
    }
    if (rfidResultEl) rfidResultEl.textContent = `Last RFID: ${value}`;
    await handleScannedCode(value, "rfid");
    if (rfidInput) rfidInput.value = "";
};

window.startQrScanner = async function () {
    if (!requirePermission("scan QR", "update_stock")) return;
    if (typeof Html5Qrcode === "undefined") {
        notify("QR library failed to load.", "error");
        updateScannerUiStatus("Library not loaded");
        return;
    }
    if (!window.isSecureContext) {
        notify("QR scanner needs HTTPS or a secure browser context.", "error");
        updateScannerUiStatus("Secure context required");
        return;
    }
    if (isQrRunning) {
        notify("QR scanner already running.");
        updateScannerUiStatus("Running");
        return;
    }

    try {
        if (!qrScannerInstance) {
            qrScannerInstance = new Html5Qrcode("qrReader");
        }
        const cameraConfig = await resolvePreferredQrCamera();
        const qrReaderEl = document.getElementById("qrReader");
        const qrBoxSize = Math.max(Math.min((qrReaderEl?.clientWidth || 280) - 32, 260), 180);

        await qrScannerInstance.start(
            cameraConfig,
            {
                fps: 8,
                qrbox: { width: qrBoxSize, height: qrBoxSize },
                aspectRatio: 1
            },
            async (decodedText) => {
                await handleScannedCode(decodedText, "qr");
            }
        );
        isQrRunning = true;
        updateScannerUiStatus("Running");
        notify("QR scanner started.", "success");
    } catch (error) {
        updateScannerUiStatus("Failed to start");
        const message = String(error?.message || error || "");
        if (/permission|denied|notallowed/i.test(message)) {
            notify("Camera permission was denied. Please allow camera access in your phone browser.", "error");
            return;
        }
        if (/notfound|devicesnotfound|overconstrained/i.test(message)) {
            notify("No suitable camera was found. Try reopening the page and use a different phone browser.", "error");
            return;
        }
        notify(`Unable to start QR scanner: ${message}`, "error");
    }
};

window.stopQrScanner = async function () {
    if (!qrScannerInstance || !isQrRunning) {
        updateScannerUiStatus("Idle");
        return;
    }
    try {
        await qrScannerInstance.stop();
        await qrScannerInstance.clear();
        isQrRunning = false;
        updateScannerUiStatus("Idle");
        notify("QR scanner stopped.");
    } catch (error) {
        updateScannerUiStatus("Failed to stop");
        notify(`Unable to stop QR scanner: ${error.message || error}`, "error");
    }
};

function renderTransactionTable() {
    const tableBody = document.querySelector("#transactionTable tbody");
    const totalInEl = document.getElementById("txTotalIn");
    const totalOutEl = document.getElementById("txTotalOut");
    const netEl = document.getElementById("txNetMovement");
    const emptyStateEl = document.getElementById("transactionEmptyState");
    const paginationInfoEl = document.getElementById("transactionPaginationInfo");
    const paginationControlsEl = document.getElementById("transactionPaginationControls");

    tableBody.innerHTML = "";
    if (paginationControlsEl) paginationControlsEl.innerHTML = "";

    const filteredTransactions = getTransactionsForExport();
    if (!filteredTransactions) {
        if (totalInEl) totalInEl.textContent = "0";
        if (totalOutEl) totalOutEl.textContent = "0";
        if (netEl) netEl.textContent = "0";
        if (paginationInfoEl) paginationInfoEl.textContent = "Showing 0-0 of 0 transactions.";
        if (emptyStateEl) emptyStateEl.hidden = false;
        return;
    }

    let totalIn = 0;
    let totalOut = 0;

    const totalItems = filteredTransactions.length;
    const totalPages = Math.max(1, Math.ceil(totalItems / state.transactionPageSize));
    if (state.transactionPage > totalPages) state.transactionPage = totalPages;
    if (state.transactionPage < 1) state.transactionPage = 1;

    const startIndex = (state.transactionPage - 1) * state.transactionPageSize;
    const endIndex = startIndex + state.transactionPageSize;
    const visibleTransactions = filteredTransactions.slice(startIndex, endIndex);

    visibleTransactions.forEach((data, index) => {
        const amount = Number(data.amount) || 0;
        const performerName = data.performedBy || data.performedByEmail || data.pic || "-";
        const performerRole = data.role || "-";
        const reason = normalizeTransactionReason(data.reason, "Other");
        const remark = String(data.remark || "").trim() || "-";
        const statusBadge = `<span class="${getStatusBadgeClass(data.type)}">${escapeHtml(data.type || "-")}</span>`;
        const reasonBadge = `<span class="${getReasonBadgeClass(reason)}">${escapeHtml(reason)}</span>`;
        const row = `
            <tr>
                <td>${startIndex + index + 1}</td>
                <td>${escapeHtml(data.materialCode || "-")}</td>
                <td>${statusBadge}</td>
                <td>${reasonBadge}</td>
                <td>${escapeHtml(remark)}</td>
                <td>${amount}</td>
                <td>${escapeHtml(data.timestamp || "-")}</td>
                <td>${escapeHtml(`${performerName} (${performerRole})`)}</td>
            </tr>
        `;

        tableBody.innerHTML += row;
    });

    filteredTransactions.forEach((data) => {
        const amount = Number(data.amount) || 0;
        if (data.type === "IN") totalIn += amount;
        if (data.type === "OUT") totalOut += amount;
    });

    if (totalInEl) totalInEl.textContent = String(totalIn);
    if (totalOutEl) totalOutEl.textContent = String(totalOut);
    if (netEl) netEl.textContent = String(totalIn - totalOut);
    if (paginationInfoEl) {
        const from = totalItems ? startIndex + 1 : 0;
        const to = totalItems ? Math.min(endIndex, totalItems) : 0;
        paginationInfoEl.textContent = `Showing ${from}-${to} of ${totalItems} transactions.`;
    }

    if (emptyStateEl) {
        emptyStateEl.hidden = filteredTransactions.length > 0;
    }

    paginateTable();

    updateReportsPanel();
}

window.renderTransactionTable = renderTransactionTable;

window.paginateTable = function paginateTable(page = state.transactionPage) {
    const paginationControlsEl = document.getElementById("transactionPaginationControls");
    const totalItems = getTransactionsForExport()?.length || 0;
    const totalPages = Math.max(1, Math.ceil(totalItems / state.transactionPageSize));

    if (!paginationControlsEl) return;

    if (page !== state.transactionPage) {
        state.transactionPage = Math.min(Math.max(page, 1), totalPages);
        renderTransactionTable();
        return;
    }

    paginationControlsEl.innerHTML = "";
    if (totalItems <= state.transactionPageSize) return;

    const previousBtn = document.createElement("button");
    previousBtn.type = "button";
    previousBtn.textContent = "Previous";
    previousBtn.disabled = state.transactionPage === 1;
    previousBtn.addEventListener("click", () => window.paginateTable(state.transactionPage - 1));
    paginationControlsEl.appendChild(previousBtn);

    for (let page = 1; page <= totalPages; page += 1) {
        const pageBtn = document.createElement("button");
        pageBtn.type = "button";
        pageBtn.textContent = String(page);
        if (page === state.transactionPage) pageBtn.classList.add("active-page");
        pageBtn.addEventListener("click", () => window.paginateTable(page));
        paginationControlsEl.appendChild(pageBtn);
    }

    const nextBtn = document.createElement("button");
    nextBtn.type = "button";
    nextBtn.textContent = "Next";
    nextBtn.disabled = state.transactionPage === totalPages;
    nextBtn.addEventListener("click", () => window.paginateTable(state.transactionPage + 1));
    paginationControlsEl.appendChild(nextBtn);
};

window.searchTable = function searchTable() {
    state.transactionSearch = (document.getElementById("txSearchInput")?.value || "").trim();
    state.transactionPage = 1;
    renderTransactionTable();
};

window.sortTable = function sortTable(key) {
    if (!key) return;

    if (state.transactionSortKey === key) {
        state.transactionSortDirection = state.transactionSortDirection === "asc" ? "desc" : "asc";
    } else {
        state.transactionSortKey = key;
        state.transactionSortDirection = key === "timestamp" ? "desc" : "asc";
    }

    state.transactionPage = 1;
    renderTransactionTable();
};

function loadTransactions() {
    const transactionsRef = ref(db, "transactions");
    onValue(transactionsRef, (snapshot) => {
        const records = [];

        snapshot.forEach((childSnapshot) => {
            const data = childSnapshot.val();
            if (!data || !data.materialCode) return;

            records.push({
                ...data,
                reason: normalizeTransactionReason(data.reason, "Other"),
                remark: String(data.remark || "").trim(),
                _time: parseTransactionTime(data)
            });
        });

        records.sort((a, b) => (Number(b._time) || 0) - (Number(a._time) || 0));
        state.transactions = records;
        state.transactionPage = 1;
        const latest = records[0];
        if (latest) {
            updateOverviewLastFlow(
                latest.materialCode || "-",
                latest.type || "",
                latest.amount,
                latest.timestamp || "-",
                latest.performedBy || latest.pic || latest.performedByEmail || "-"
            );
        } else {
            updateOverviewLastFlow("-", "", "-", "-", "-");
        }
        renderTransactionTable();
        updateCharts();
    });
}

window.loadTransactions = loadTransactions;

function listenForExternalRfidScans() {
    const externalRfidRef = ref(db, "rfidScans/latest");
    onValue(externalRfidRef, async (snapshot) => {
        const data = snapshot.val();
        if (!data) return;

        if (!hasPrimedExternalRfidListener) {
            const initialTag = String(data?.tag ?? data?.code ?? data?.rfid ?? "").trim();
            const initialScannedAt = String(data?.scannedAt || data?.timestamp || data?.createdAt || "");
            const initialSource = String(data?.source || "external-rfid").trim() || "external-rfid";
            const initialType = normalizeScanType(data?.action ?? data?.type, "IN");
            lastExternalRfidSignature = `${initialTag}|${initialScannedAt}|${initialSource}|${initialType}`;
            lastExternalRfidAt = Date.now();
            hasPrimedExternalRfidListener = true;
            return;
        }

        try {
            await handleExternalRfidScan(data);
        } catch (error) {
            notify(`External RFID scan failed: ${error.message || error}`, "error");
        }
    });
}

function updateSummaryCards() {
    const totalMaterials = state.materials.length;
    const totalStock = state.materials.reduce((sum, material) => sum + (Number(material.stock) || 0), 0);
    const lowStockCount = state.materials.filter((material) => (Number(material.stock) || 0) <= LOW_STOCK_THRESHOLD).length;
    const normalCount = totalMaterials - lowStockCount;

    document.getElementById("totalMaterials").innerText = totalMaterials;
    document.getElementById("totalStock").innerText = totalStock;
    document.getElementById("lowStockItems").innerText = lowStockCount;

    const filterSelect = document.getElementById("filterSelect");
    filterSelect.options[1].text = `Low Stock (${lowStockCount})`;
    filterSelect.options[2].text = `Normal (${normalCount})`;

    const badge = document.getElementById("lowStockBadge");
    if (lowStockCount > 0) {
        const lowStockCodes = state.materials
            .filter((material) => (Number(material.stock) || 0) <= LOW_STOCK_THRESHOLD)
            .map((material) => material.code || material.name || "-")
            .filter((value) => String(value).trim() !== "");
        const previewCodes = lowStockCodes.slice(0, 5).join(", ");
        const remainingCount = Math.max(lowStockCodes.length - 5, 0);
        const codeSuffix = remainingCount > 0 ? ` +${remainingCount} more` : "";

        badge.style.display = "block";
        badge.textContent = `LOW STOCK ALERT: ${lowStockCount} item(s) need restock! Code: ${previewCodes}${codeSuffix}`;
    } else {
        badge.style.display = "none";
    }
}

function ensureChartLibrary() {
    return typeof Chart !== "undefined";
}

function updateCharts() {
    if (!ensureChartLibrary()) return;

    const palette = {
        in: "#1d4ed8",
        out: "#0f766e",
        low: "#ea580c",
        normal: "#16a34a",
        topBar: "#1e40af"
    };

    const monthlyMap = new Map();
    const outByCode = new Map();
    const nameByCode = new Map();
    const chartTransactions = getTransactionsForCharts();

    state.materials.forEach((material) => {
        if (material.code) {
            nameByCode.set(material.code, material.name || material.code);
        }
    });

    chartTransactions.forEach((tx) => {
        const time = Number(tx._time) || Date.parse(tx.timestampISO || tx.timestamp || "");
        if (!Number.isFinite(time)) return;

        const date = new Date(time);
        const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
        if (!monthlyMap.has(monthKey)) {
            monthlyMap.set(monthKey, { in: 0, out: 0 });
        }

        const amount = Number(tx.amount) || 0;
        if (tx.type === "IN") monthlyMap.get(monthKey).in += amount;
        if (tx.type === "OUT") {
            monthlyMap.get(monthKey).out += amount;
            const code = tx.materialCode || "-";
            outByCode.set(code, (outByCode.get(code) || 0) + amount);
        }
    });

    const monthLabels = [...monthlyMap.keys()].sort();
    const inSeries = monthLabels.map((m) => monthlyMap.get(m).in);
    const outSeries = monthLabels.map((m) => monthlyMap.get(m).out);

    const lowCount = state.materials.filter((m) => (Number(m.stock) || 0) <= LOW_STOCK_THRESHOLD).length;
    const normalCount = Math.max(state.materials.length - lowCount, 0);

    const activeCodes = new Set(
        state.materials
            .map((material) => material.code)
            .filter((code) => typeof code === "string" && code.trim() !== "")
    );

    const topOut = [...outByCode.entries()]
        .filter(([code]) => activeCodes.has(code))
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);
    const topLabels = topOut.map(([code]) => nameByCode.get(code) || code);
    const topValues = topOut.map(([, amount]) => amount);

    if (stockMovementChartInstance) stockMovementChartInstance.destroy();
    if (stockStatusChartInstance) stockStatusChartInstance.destroy();
    if (topMaterialsChartInstance) topMaterialsChartInstance.destroy();

    const stockMovementCtx = document.getElementById("stockMovementChart");
    if (stockMovementCtx) {
        stockMovementChartInstance = new Chart(stockMovementCtx, {
            type: "bar",
            data: {
                labels: monthLabels.length ? monthLabels : ["No Data"],
                datasets: [
                    {
                        label: "Stock IN",
                        data: monthLabels.length ? inSeries : [0],
                        backgroundColor: palette.in,
                        borderRadius: 6
                    },
                    {
                        label: "Stock OUT",
                        data: monthLabels.length ? outSeries : [0],
                        backgroundColor: palette.out,
                        borderRadius: 6
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        labels: { usePointStyle: true }
                    }
                }
            }
        });
    }

    const stockStatusCtx = document.getElementById("stockStatusChart");
    if (stockStatusCtx) {
        stockStatusChartInstance = new Chart(stockStatusCtx, {
            type: "doughnut",
            data: {
                labels: ["Low Stock", "Normal"],
                datasets: [
                    {
                        data: [lowCount, normalCount],
                        backgroundColor: [palette.low, palette.normal],
                        borderColor: "#ffffff",
                        borderWidth: 2
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false
            }
        });
    }

    const topMaterialsCtx = document.getElementById("topMaterialsChart");
    if (topMaterialsCtx) {
        topMaterialsChartInstance = new Chart(topMaterialsCtx, {
            type: "bar",
            data: {
                labels: topLabels.length ? topLabels : ["No Data"],
                datasets: [
                    {
                        label: "Stock OUT Amount",
                        data: topLabels.length ? topValues : [0],
                        backgroundColor: palette.topBar,
                        borderRadius: 6
                    }
                ]
            },
            options: {
                indexAxis: "y",
                responsive: true,
                maintainAspectRatio: false
            }
        });
    }
}

window.login = function () {
    window.location.href = "login.html";
};

window.logout = async function () {
    try {
        await window.stopQrScanner();
    } catch (error) {
        notify(`Logout failed: ${error.message}`, "error");
        return;
    }

    state.user = null;
    state.role = null;
    state.userProfile = null;
    await logoutToLogin();
};

window.exportMaterialsExcel = function () {
    if (!requirePermission("export materials", "export_reports")) return;

    const rows = getVisibleMaterials().map((material, index) => [
        index + 1,
        material.code || "-",
        material.rfidTag || "-",
        material.name || "-",
        material.stock ?? 0,
        (Number(material.stock) || 0) <= LOW_STOCK_THRESHOLD ? "LOW STOCK" : "Normal",
        material.createdAt || "-",
        material.createdByEmail || "-",
        material.updatedAt || "-",
        material.updatedByEmail || "-"
    ]);

    const actor = getActorInfo();
    const reportId = generateReportId("MT");
    const excelXml = buildMaterialsExcelReport(rows, {
        reportId,
        generatedAt: formatReportDateTime(new Date()),
        generatedBy: actor.name || "-",
        generatedEmail: actor.email || "-",
        generatedRole: formatRoleLabel(actor.role),
        recordCount: rows.length,
        dateRange: "All Dates",
        reportInfo: { ...state.reportInfo }
    });
    downloadFile(excelXml, `materials_${reportId}.xls`, "application/vnd.ms-excel;charset=utf-8");
    addReportLog("Materials Export", "EXCEL", rows.length);
};

window.exportMaterialsPDF = async function () {
    if (!requirePermission("export materials", "export_reports")) return;

    const rows = getVisibleMaterials().map((material, index) => [
        index + 1,
        material.code || "-",
        material.rfidTag || "-",
        material.name || "-",
        material.stock ?? 0,
        (Number(material.stock) || 0) <= LOW_STOCK_THRESHOLD ? "LOW STOCK" : "Normal",
        material.createdAt || "-",
        material.createdByEmail || "-",
        material.updatedAt || "-",
        material.updatedByEmail || "-"
    ]);

    const reportId = generateReportId("MT");
    const meta = buildPdfMeta("Materials Export", rows.length, {
        dateRange: "All Dates",
        reportId
    });

    await printTableAsPdf(
        "Materials Export",
        ["Bil.", "Code", "RFID Tag", "Materials", "Stock", "Status", "Created At", "Created By", "Updated At", "Updated By"],
        rows,
        meta
    );
    addReportLog("Materials Export", "PDF", rows.length);
};

window.exportTransactionsExcel = function () {
    if (!requirePermission("export transactions", "export_reports")) return;

    const filteredTransactions = getTransactionsForExport();
    if (!filteredTransactions) return;

    const rows = filteredTransactions.map((data, index) => [
        index + 1,
        data.materialCode || "-",
        data.type || "-",
        normalizeTransactionReason(data.reason, "Other"),
        data.remark || "-",
        data.amount ?? 0,
        data.timestamp || "-",
        data.performedBy || data.performedByEmail || data.pic || "-",
        data.role || "-"
    ]);

    const actor = getActorInfo();
    const reportId = generateReportId("TX");
    const excelXml = buildTransactionsExcelReport(rows, {
        reportId,
        generatedAt: formatReportDateTime(new Date()),
        generatedBy: actor.name || "-",
        generatedEmail: actor.email || "-",
        generatedRole: formatRoleLabel(actor.role),
        recordCount: rows.length,
        dateRange: getTransactionRangeLabel(),
        reportInfo: { ...state.reportInfo }
    });

    downloadFile(excelXml, `transactions_${reportId}.xls`, "application/vnd.ms-excel;charset=utf-8");
    addReportLog("Transactions Export", "EXCEL", rows.length);
};

window.exportTransactionsPDF = async function () {
    if (!requirePermission("export transactions", "export_reports")) return;

    const filteredTransactions = getTransactionsForExport();
    if (!filteredTransactions) return;

    const rows = filteredTransactions.map((data, index) => [
        index + 1,
        data.materialCode || "-",
        data.type || "-",
        normalizeTransactionReason(data.reason, "Other"),
        data.remark || "-",
        data.amount ?? 0,
        data.timestamp || "-",
        data.performedBy || data.performedByEmail || data.pic || "-",
        data.role || "-"
    ]);

    const reportId = generateReportId("TX");
    const meta = buildPdfMeta("Transactions Export", rows.length, {
        dateRange: getTransactionRangeLabel(),
        reportId
    });

    await printTableAsPdf(
        "Transactions Export",
        ["Bil.", "Code", "Status", "Reason", "Remark", "Stock", "Created At", "Performed By", "Role"],
        rows,
        meta
    );
    addReportLog("Transactions Export", "PDF", rows.length);
};

window.exportToExcel = window.exportTransactionsExcel;
window.exportToPDF = window.exportTransactionsPDF;

window.addMaterial = function () {
    if (!requirePermission("add materials", "add_material")) return;

    const nameInput = document.getElementById("materialName").value.trim();
    const rfidTagInput = normalizeRfidTag(document.getElementById("materialRfidTag")?.value || "");
    const stockInput = document.getElementById("materialStock").value;

    if (nameInput === "" || stockInput === "") {
        notify("Please enter both material name and stock before submitting.", "error");
        return;
    }

    const parsedStock = Number(stockInput);
    if (!Number.isFinite(parsedStock) || parsedStock < 0) {
        notify("Initial stock must be a valid non-negative number.", "error");
        return;
    }

    if (rfidTagInput) {
        const duplicateRfid = state.materials.some((material) => normalizeRfidTag(material.rfidTag) === rfidTagInput);
        if (duplicateRfid) {
            notify(`RFID tag '${rfidTagInput}' is already linked to another material.`, "error");
            return;
        }
    }

    const materialCode = generateUniqueMaterialCode();
    const materialRef = push(ref(db, "materials"));

    const now = new Date();
    const actor = getActorInfo();

    set(materialRef, {
        code: materialCode,
        name: nameInput,
        rfidTag: rfidTagInput,
        stock: parsedStock,
        createdAt: now.toLocaleString(),
        createdAtISO: now.toISOString(),
        createdByUid: actor.uid,
        createdByEmail: actor.email,
        updatedAt: now.toLocaleString(),
        updatedAtISO: now.toISOString(),
        updatedByUid: actor.uid,
        updatedByEmail: actor.email
    })
        .then(() => {
            notify(`Material '${nameInput}' has been successfully added.`, "success");
            document.getElementById("materialName").value = "";
            document.getElementById("materialRfidTag").value = "";
            document.getElementById("materialStock").value = "";
        })
        .catch((error) => {
            notify(`Error occurred: ${error.message}`, "error");
        });
};

window.deleteMaterial = async function (id) {
    if (!requirePermission("delete materials", "delete_material")) return;

    const confirmDelete = await openConfirmModal("Are you sure you want to delete this material?");
    if (!confirmDelete) return;

    remove(ref(db, "materials/" + id))
        .then(() => {
            notify("Material deleted successfully.", "success");
        })
        .catch((error) => {
            notify(`Error: ${error.message}`, "error");
        });
};

window.editMaterial = async function (id) {
    if (!requirePermission("edit materials", "edit_material")) return;

    const material = getMaterialById(id);
    if (!material) {
        notify("Material not found. Please refresh and try again.", "error");
        return;
    }

    const edited = await openEditModal(material.name || "", Number(material.stock) || 0, material.rfidTag || "");
    if (!edited) return;

    if (edited.rfidTag) {
        const duplicateRfid = state.materials.some((item) => item.id !== id && normalizeRfidTag(item.rfidTag) === edited.rfidTag);
        if (duplicateRfid) {
            notify(`RFID tag '${edited.rfidTag}' is already linked to another material.`, "error");
            return;
        }
    }

    const now = new Date();
    const actor = getActorInfo();

    update(ref(db, "materials/" + id), {
        name: edited.name,
        rfidTag: edited.rfidTag,
        stock: edited.stock,
        updatedAt: now.toLocaleString(),
        updatedAtISO: now.toISOString(),
        updatedByUid: actor.uid,
        updatedByEmail: actor.email
    })
        .then(() => {
            notify("Material updated successfully.", "success");
        })
        .catch((error) => {
            notify(`Error: ${error.message}`, "error");
        });
};

async function saveTransaction(materialCode, type, amount, pic, reason = "Other", remark = "") {
    const now = new Date();
    const actor = getActorInfo();
    const transactionRef = push(ref(db, "transactions"));
    await set(transactionRef, {
        materialCode,
        type,
        amount,
        reason: normalizeTransactionReason(reason, type === "OUT" ? "Site Usage" : "Other"),
        remark: String(remark || "").trim(),
        timestamp: now.toLocaleString(),
        timestampISO: now.toISOString(),
        pic: pic.trim(),
        performedBy: actor.name,
        role: actor.role,
        performedByUid: actor.uid,
        performedByEmail: actor.email
    });

    return true;
}

async function addTransaction(materialCode, type, amount, pic, reason = "Other", remark = "") {
    return saveTransaction(materialCode, type, amount, pic, reason, remark);
}

window.saveTransaction = saveTransaction;

async function applyMaterialStockDelta(materialId, delta, actor, options = {}) {
    const materialRef = ref(db, `materials/${materialId}`);
    const now = new Date();
    let abortedBecauseNegative = false;

    const result = await runTransaction(materialRef, (current) => {
        if (!current) return current;

        const currentStock = Number(current.stock) || 0;
        const nextStock = currentStock + delta;
        if (nextStock < 0) {
            abortedBecauseNegative = true;
            return;
        }

        return {
            ...current,
            stock: nextStock,
            updatedAt: options.updatedAt || now.toLocaleString(),
            updatedAtISO: options.updatedAtISO || now.toISOString(),
            updatedByUid: actor.uid,
            updatedByEmail: actor.email
        };
    });

    if (!result.committed) {
        if (abortedBecauseNegative) {
            throw new Error("Stock cannot be negative!");
        }
        throw new Error("Material not found or stock update was not committed.");
    }

    return Number(result.snapshot?.child("stock").val()) || 0;
}

async function recordInventoryMovement(material, type, amount, pic, reason = "Other", remark = "") {
    const actor = getActorInfo();
    const delta = type === "IN" ? amount : -amount;

    await applyMaterialStockDelta(material.id, delta, actor);

    try {
        await saveTransaction(material.code, type, amount, pic, reason, remark);
    } catch (error) {
        try {
            const rollbackNow = new Date();
            await applyMaterialStockDelta(material.id, -delta, actor, {
                updatedAt: rollbackNow.toLocaleString(),
                updatedAtISO: rollbackNow.toISOString()
            });
        } catch (rollbackError) {
            console.error("Rollback failed after transaction save error:", rollbackError);
        }
        throw error;
    }
}

function getMaterialById(id) {
    return state.materials.find((material) => material.id === id) || null;
}

window.increaseStock = async function (id, materialCode) {
    if (!requirePermission("update stock", "update_stock")) return;

    const material = getMaterialById(id);
    if (!material) {
        notify("Material not found. Please refresh and try again.", "error");
        return;
    }

    const data = await openStockModal("IN");
    if (!data) return;

    try {
        await recordInventoryMovement(material, "IN", data.amount, data.pic, data.reason, data.remark);
        notify("Stock updated successfully.", "success");
    } catch (error) {
        notify(`Error: ${error.message || error}`, "error");
    }
};

window.decreaseStock = async function (id, materialCode) {
    if (!requirePermission("update stock", "update_stock")) return;

    const material = getMaterialById(id);
    if (!material) {
        notify("Material not found. Please refresh and try again.", "error");
        return;
    }

    const data = await openStockModal("OUT");
    if (!data) return;

    try {
        await recordInventoryMovement(material, "OUT", data.amount, data.pic, data.reason, data.remark);
        notify("Stock updated successfully.", "success");
    } catch (error) {
        notify(`Error: ${error.message || error}`, "error");
    }
};

window.searchMaterial = function () {
    state.search = document.getElementById("searchInput").value.toLowerCase();
    renderMaterialTable();
};

window.filterMaterials = function (type) {
    state.filter = type;
    renderMaterialTable();
};

window.sortMaterials = function (type) {
    state.sort = type;
    renderMaterialTable();
};

window.toggleSidebar = function () {
    const sidebar = document.getElementById("sidebar");
    if (!sidebar) return;
    sidebar.classList.toggle("open");
};

window.resetDateRange = function () {
    const fromInput = document.getElementById("txFromDate");
    const toInput = document.getElementById("txToDate");
    const searchInput = document.getElementById("txSearchInput");
    const typeFilter = document.getElementById("txTypeFilter");
    if (fromInput) fromInput.value = "";
    if (toInput) toInput.value = "";
    if (searchInput) searchInput.value = "";
    if (typeFilter) typeFilter.value = "all";
    state.transactionPage = 1;

    renderTransactionTable();
    updateCharts();
};

window.savePreferences = function () {
    const themeSelect = document.getElementById("settingsThemeSelect");
    const rangeSelect = document.getElementById("settingsDefaultRange");

    state.settings.theme = themeSelect?.value || DEFAULT_SETTINGS.theme;
    state.settings.defaultRange = rangeSelect?.value || DEFAULT_SETTINGS.defaultRange;

    persistSettings();
    applyTheme(state.settings.theme);
    applyDefaultDateRange();
    renderTransactionTable();
    updateCharts();
    updateSettingsProfileUI();
    notify("Preferences saved.", "success");
};

window.resetPreferences = function () {
    state.settings = { ...DEFAULT_SETTINGS };
    persistSettings();
    applyTheme(state.settings.theme);

    const themeSelect = document.getElementById("settingsThemeSelect");
    const rangeSelect = document.getElementById("settingsDefaultRange");
    if (themeSelect) themeSelect.value = state.settings.theme;
    if (rangeSelect) rangeSelect.value = state.settings.defaultRange;

    window.resetDateRange();
    updateSettingsProfileUI();
    notify("Preferences reset to default.", "success");
};

function setupSidebarNavigation() {
    const sidebar = document.getElementById("sidebar");
    const links = document.querySelectorAll(".sidebar-link");
    const sections = Array.from(document.querySelectorAll("main section[id]")).filter((section) => !section.hidden);

    links.forEach((link) => {
        link.addEventListener("click", () => {
            if (sidebar && window.innerWidth <= 768) {
                sidebar.classList.remove("open");
            }
        });
    });

    if (!sections.length) return;

    const observer = new IntersectionObserver((entries) => {
        entries.forEach((entry) => {
            if (!entry.isIntersecting) return;

            const sectionId = entry.target.id;
            links.forEach((link) => {
                const isActive = link.dataset.section === sectionId;
                link.classList.toggle("active", isActive);
            });
        });
    }, { rootMargin: "-30% 0px -55% 0px", threshold: 0 });

    sections.forEach((section) => observer.observe(section));
}

function startRealtimeBindings() {
    if (hasRealtimeBindings) return;
    hasRealtimeBindings = true;

    const materialsRef = ref(db, "materials");
    onValue(materialsRef, (snapshot) => {
        const materials = [];

        snapshot.forEach((childSnapshot) => {
            const data = childSnapshot.val() || {};
            materials.push({
                id: childSnapshot.key,
                code: data.code,
                name: data.name,
                rfidTag: data.rfidTag || "",
                stock: Number(data.stock) || 0,
                createdAt: data.createdAt || "-",
                createdByEmail: data.createdByEmail || "-",
                updatedAt: data.updatedAt || data.createdAt || "-",
                updatedByEmail: data.updatedByEmail || data.createdByEmail || "-"
            });
        });

        state.materials = materials;
        updateSummaryCards();
        updateReportsPanel();
        renderMaterialTable();
        updateCharts();
    });

    loadTransactions();
    listenForExternalRfidScans();
}

const txFromInput = document.getElementById("txFromDate");
const txToInput = document.getElementById("txToDate");
const txSearchInput = document.getElementById("txSearchInput");
const txTypeFilter = document.getElementById("txTypeFilter");
if (txFromInput) {
    txFromInput.addEventListener("change", () => {
        renderTransactionTable();
        updateCharts();
    });
}
if (txToInput) {
    txToInput.addEventListener("change", () => {
        renderTransactionTable();
        updateCharts();
    });
}
if (txTypeFilter) {
    txTypeFilter.addEventListener("change", () => {
        state.transactionPage = 1;
        renderTransactionTable();
    });
}
if (txSearchInput) {
    txSearchInput.addEventListener("input", window.searchTable);
}

const rfidInput = document.getElementById("rfidInput");
if (rfidInput) {
    rfidInput.addEventListener("input", () => {
        const currentValue = String(rfidInput.value || "").trim();
        if (!currentValue) {
            if (rfidAutoSubmitTimer) {
                clearTimeout(rfidAutoSubmitTimer);
                rfidAutoSubmitTimer = null;
            }
            return;
        }
        queueRfidAutoSubmit();
    });
    rfidInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
            event.preventDefault();
            window.processRfidInput();
        }
    });
}

async function bootstrap() {
    try {
        const session = await requireRole([ROLES.MANAGER, ROLES.STOREKEEPER, ROLES.SITE_SUPERVISOR]);
        if (!session?.profile) return;

        state.user = {
            uid: session.user.uid,
            email: session.user.email || session.profile.email,
            displayName: session.profile.name || session.user.displayName || session.user.email || "User"
        };
        state.role = session.profile.role;
        state.userProfile = {
            uid: session.profile.uid,
            email: session.profile.email || session.user.email || "",
            name: session.profile.name || session.user.displayName || "",
            role: session.profile.role
        };

        startRealtimeBindings();
        clearRoleListener();
        updateAuthUI();
        renderMaterialTable();

        state.settings = loadSettings();
        state.reportInfo = loadReportInfo();
        applyTheme(state.settings.theme);
        applyDefaultDateRange();
        updateSettingsProfileUI();
        syncReportInfoUi();
        updateScannerUiStatus("Idle");

        setupSidebarNavigation();
        setupModalHandlers();
    } catch (error) {
        console.error("Application bootstrap failed:", error);
        clearCachedSession();
        window.location.replace("login.html");
    }
}

bootstrap();
