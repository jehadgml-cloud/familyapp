// app.js

// Google Apps Script Web App URL — paste the deployment URL from
// gas/DEPLOYMENT.md step 3 here. Every piece of app data is read from and
// written to the Google Sheet behind this URL; there is no local fallback.
const API_URL = "https://script.google.com/macros/s/AKfycbxokCwpd7RhoFboc9_3VVqzpdObXAL_x6pje6uHbN6is8z91BIM7zaTLFGp6P8ttOAa/exec";

// Calls one backend action. Uses text/plain as the Content-Type so the
// browser treats this as a CORS-simple request and skips the preflight
// OPTIONS request — Apps Script Web Apps don't answer OPTIONS.
async function callApi(action, payload) {
    let res;
    try {
        res = await fetch(API_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({ action, payload: payload || {} })
        });
    } catch (networkErr) {
        throw new Error("تعذر الاتصال بالخادم. تحقق من اتصالك بالإنترنت وحاول مجدداً. (" + networkErr.message + ")");
    }
    if (!res.ok) {
        throw new Error("خطأ من الخادم (HTTP " + res.status + "). حاول مجدداً لاحقاً.");
    }
    let body;
    try {
        body = await res.json();
    } catch (parseErr) {
        throw new Error("خطأ في قراءة استجابة الخادم. حاول مجدداً لاحقاً.");
    }
    if (body.error) {
        throw new Error(body.error);
    }
    return body.result;
}

// Looks up a user in the already-loaded state.users (populated from the
// Sheet by loadDataFromServer()). Used only for instant client-side
// pre-checks (e.g. "this email is already registered") — the server
// performs the authoritative check on every mutating call.
function findUserInState_(email) {
    if (!email) return null;
    const lower = email.trim().toLowerCase();
    return state.users.find(u => u.email.toLowerCase() === lower) || null;
}

const DEFAULT_MONTHS = [
    { key: "أبريل 2026", id: 4 },
    { key: "مايو 2026", id: 5 },
    { key: "يونيو 2026", id: 6 },
    { key: "يوليو 2026", id: 7 },
    { key: "أغسطس 2026", id: 8 },
    { key: "سبتمبر 2026", id: 9 },
    { key: "أكتوبر 2026", id: 10 },
    { key: "نوفمبر 2026", id: 11 },
    { key: "ديسمبر 2026", id: 12 },
    { key: "يناير 2027", id: 1 },
    { key: "فبراير 2027", id: 2 },
    { key: "مارس 2027", id: 3 }
];


// Application State — populated from the Google Sheet backend via
// loadDataFromServer() on startup; nothing here is read from localStorage.
let state = {
    currentUser: null,
    members: [],
    families: [],
    monthsList: [],
    selectedLedgerMonths: [],
    expenses: [],
    users: [],
    pendingUsers: [],
    settings: {},
    originalWorkbook: null // SheetJS raw object, session-only (never persisted)
};

// Chart instances
let monthlyChartInstance = null;
let commitmentChartInstance = null;

// DOM Elements
const loginPortal = document.getElementById("login-portal");
const appContainer = document.getElementById("app-container");
const btnSubmitLogin = document.getElementById("btn-submit-login");
const loginEmailInput = document.getElementById("login-email");
const loginPassInput = document.getElementById("login-pass");
const loginErrorMsg = document.getElementById("login-error");

const adminAvatar = document.getElementById("admin-avatar");
const adminName = document.getElementById("admin-name");
const adminRole = document.getElementById("admin-role");
const tabTitleText = document.getElementById("tab-title-text");
const liveDateSpan = document.getElementById("live-date");

const dropZone = document.getElementById("drop-zone");
const btnFileSelect = document.getElementById("btn-file-select");
const uploadStatus = document.getElementById("upload-status");

// Tab Switching Local State
const navLinks = document.querySelectorAll(".nav-link");
const tabContents = document.querySelectorAll(".tab-content");

// Initial Setup
document.addEventListener("DOMContentLoaded", async () => {
    // Current date fill
    const today = new Date().toISOString().split("T")[0];
    liveDateSpan.textContent = today;
    document.getElementById("p-receipt-date").textContent = today;
    document.getElementById("r-date").textContent = today;

    // *** CRITICAL: Load data from the Google Sheet backend FIRST ***
    try {
        await loadDataFromServer();
    } catch (err) {
        alert("تعذّر الاتصال بقاعدة البيانات (جوجل شيت): " + err.message + "\n\nتحقق من رابط API_URL في app.js ومن اتصالك بالإنترنت.");
    }
    populateMonthFilters();
    populateAddMemberFamilies();

    // Check if user is logged in (session persisted by email) — the email
    // is only a client-side "stay logged in" convenience; the account
    // itself is re-validated against the freshly-loaded state.users.
    const cachedEmail = localStorage.getItem("cems_logged_email");
    if (cachedEmail) {
        const userObj = findUserInState_(cachedEmail);
        if (userObj) {
            doLogin(userObj);
        } else {
            localStorage.removeItem("cems_logged_email");
        }
    }

    // Register tab / Login toggle
    document.getElementById("link-show-register").addEventListener("click", e => {
        e.preventDefault();
        document.getElementById("panel-login").style.display = "none";
        document.getElementById("panel-register").style.display = "block";
    });
    document.getElementById("link-back-login").addEventListener("click", e => {
        e.preventDefault();
        document.getElementById("panel-register").style.display = "none";
        document.getElementById("panel-login").style.display = "block";
    });

    // Multi-month select all/none controls
    const btnSelectAll = document.getElementById("btn-select-all-months");
    if (btnSelectAll) {
        btnSelectAll.addEventListener("click", async () => {
            const next = state.monthsList.map(m => m.key);
            try {
                await callApi("setSelectedMonths", { selected: next });
                state.selectedLedgerMonths = next;
                populateMonthFilters();
                renderLedgerTable();
            } catch (err) {
                alert("تعذّر تحديد كل الأشهر: " + err.message);
            }
        });
    }

    const btnSelectNone = document.getElementById("btn-select-none-months");
    if (btnSelectNone) {
        btnSelectNone.addEventListener("click", async () => {
            try {
                await callApi("setSelectedMonths", { selected: [] });
                state.selectedLedgerMonths = [];
                populateMonthFilters();
                renderLedgerTable();
            } catch (err) {
                alert("تعذّر إلغاء تحديد الأشهر: " + err.message);
            }
        });
    }
});

// Navigation Switcher
navLinks.forEach(link => {
    link.addEventListener("click", () => {
        const tabId = link.getAttribute("data-tab");
        switchTab(tabId);
    });
});

function switchTab(tabId) {
    navLinks.forEach(l => l.classList.remove("active"));
    tabContents.forEach(c => c.classList.remove("active"));

    const activeLink = document.querySelector(`.nav-link[data-tab="${tabId}"]`);
    if(activeLink) activeLink.classList.add("active");
    
    const activeContent = document.getElementById(tabId);
    if(activeContent) activeContent.classList.add("active");

    updateTabTitle(tabId);

    // Refresh charts if dashboard is selected
    if (tabId === "dashboard-tab") {
        renderCharts();
    } else if (tabId === "ledger-tab") {
        renderLedgerTable();
    } else if (tabId === "families-tab") {
        renderFamiliesTable();
    } else if (tabId === "receipt-tab") {
        populateReceiptFamilies();
    } else if (tabId === "reports-tab") {
        renderMonthlyReport();
    } else if (tabId === "usermgmt-tab") {
        renderPendingUsers();
        renderActiveUsers();
    }
}

function updateTabTitle(tabId) {
    const titles = {
        "dashboard-tab": "لوحة الاتجاهات والتحليلات البيانية",
        "ledger-tab": "دفتر كشف التحصيل الشهري العام",
        "families-tab": "دليل العائلات ومشاركة الوصل السريع",
        "receipt-tab": "إصدار وطباعة وصولات الدفع",
        "reports-tab": "التقارير والمراجعات المالية الشهرية والسنوية",
        "expenses-tab": "سجل أوامر الصرف والمدفوعات المعتمدة",
        "usermgmt-tab": "إدارة المشرفين وطلبات التسجيل",
        "upload-tab": "بوابة تحميل وتخزين قاعدة البيانات"
    };
    tabTitleText.textContent = titles[tabId] || "صندوق العشيرة المالي";
}

// Login Handler — email-based, verified against the Google Sheet backend
btnSubmitLogin.addEventListener("click", async () => {
    const email = loginEmailInput.value.trim().toLowerCase();
    const pass = loginPassInput.value;

    if (!email) {
        showLoginError("الرجاء إدخال بريدك الإلكتروني.");
        return;
    }
    if (!pass) {
        showLoginError("الرجاء إدخال كلمة المرور.");
        return;
    }

    btnSubmitLogin.disabled = true;
    try {
        const result = await callApi("login", { email, pass });
        if (result.error) {
            showLoginError(result.error);
            return;
        }
        const userObj = result.user;
        localStorage.setItem("cems_logged_email", email);

        // Force password change on first login for built-in admins
        if (userObj.firstLoginDone === false && userObj.isBuiltIn) {
            showPasswordChangeDialog(userObj, () => doLogin(userObj));
        } else {
            doLogin(userObj);
        }
    } catch (err) {
        showLoginError("تعذّر تسجيل الدخول: " + err.message);
    } finally {
        btnSubmitLogin.disabled = false;
    }
});

// Password Reset Handler (Email-based)
document.getElementById("link-reset-pass").addEventListener("click", async (e) => {
    e.preventDefault();
    const email = loginEmailInput.value.trim().toLowerCase();
    if (!email) {
        alert("الرجاء إدخال البريد الإلكتروني للحساب المراد إعادة تعيين رمزه أولاً.");
        return;
    }

    const userObj = findUserInState_(email);
    if (!userObj) {
        alert("لا يوجد حساب مشرف مالي مسجل بهذا البريد الإلكتروني.");
        return;
    }

    const confirmReset = confirm(`هل أنت متأكد من رغبتك في إعادة تعيين رمز المرور لحساب المشرف (${userObj.name})؟ \n\nسيتم استخدام رمز المرور الافتراضي (ABC12345) وسيُطلب منك تغييره عند الدخول.`);
    if (!confirmReset) return;

    try {
        await callApi("resetPassword", { email });
        alert(`تمت إعادة تعيين رمز مرور الحساب بنجاح إلى: ABC12345 \nيرجى تسجيل الدخول فيه الآن وتعديله.`);
        loginPassInput.value = "";
        loginErrorMsg.style.display = "none";
    } catch (err) {
        alert("تعذّرت إعادة تعيين رمز المرور: " + err.message);
    }
});

// Submit Registration Request
document.getElementById("btn-submit-register").addEventListener("click", async () => {
    const firstName = document.getElementById("reg-firstname").value.trim();
    const lastName = document.getElementById("reg-lastname").value.trim();
    const email = document.getElementById("reg-email").value.trim().toLowerCase();
    const pass = document.getElementById("reg-pass").value;
    const pass2 = document.getElementById("reg-pass2").value;
    const errorEl = document.getElementById("reg-error");

    errorEl.style.display = "none";

    if (!firstName || !lastName || !email || !pass || !pass2) {
        errorEl.textContent = "الرجاء تعبئة كافة الحقول المطلوبة.";
        errorEl.style.display = "block";
        return;
    }

    if (pass.length < 6) {
        errorEl.textContent = "يجب أن تكون كلمة المرور 6 أحرف على الأقل.";
        errorEl.style.display = "block";
        return;
    }

    if (pass !== pass2) {
        errorEl.textContent = "كلمتا المرور غير متطابقتين.";
        errorEl.style.display = "block";
        return;
    }

    if (findUserInState_(email) || state.pendingUsers.some(r => r.email.toLowerCase() === email)) {
        errorEl.textContent = "هذا البريد الإلكتروني مسجل بالفعل أو له طلب تسجيل معلق.";
        errorEl.style.display = "block";
        return;
    }

    try {
        state.pendingUsers = await callApi("requestRegistration", { firstName, lastName, email, pass });

        alert(`✅ تم إرسال طلب تسجيلك بنجاح باسم (${firstName} ${lastName})! يرجى انتظار موافقة المسؤول جهاد زكري لتتمكن من تسجيل الدخول.`);

        document.getElementById("reg-firstname").value = "";
        document.getElementById("reg-lastname").value = "";
        document.getElementById("reg-email").value = "";
        document.getElementById("reg-pass").value = "";
        document.getElementById("reg-pass2").value = "";

        document.getElementById("panel-register").style.display = "none";
        document.getElementById("panel-login").style.display = "block";
    } catch (err) {
        errorEl.textContent = "تعذّر إرسال الطلب: " + err.message;
        errorEl.style.display = "block";
    }
});

// Forced first-login password change dialog (built-in admins)
function showPasswordChangeDialog(userObj, onSuccess) {
    const existing = document.getElementById('pwd-change-dialog');
    if (existing) existing.remove();

    const overlay = document.createElement('div');
    overlay.id = 'pwd-change-dialog';
    overlay.style.cssText = `
        position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
        background: rgba(0,0,0,0.7); z-index: 200;
        display: flex; align-items: center; justify-content: center;
        font-family: 'Cairo', sans-serif;
    `;

    overlay.innerHTML = `
        <div style="background: white; border-radius: 20px; padding: 36px;
                    width: 100%; max-width: 440px; box-shadow: 0 20px 40px rgba(0,0,0,0.3);
                    direction: rtl; text-align: right;">
            <div style="text-align: center; margin-bottom: 24px;">
                <div style="font-size: 2.5rem; margin-bottom: 8px;">🔐</div>
                <h2 style="color: #0f766e; font-size: 1.3rem; margin: 0;">تغيير كلمة المرور الإلزامي</h2>
                <p style="color: #64748b; font-size: 0.85rem; margin-top: 8px;">
                    مرحباً <strong>${userObj.name}</strong>!<br>
                    لدواعي الأمان، يجب تغيير كلمة المرور الافتراضية قبل الدخول.
                </p>
            </div>
            <div style="display: flex; flex-direction: column; gap: 14px;">
                <div>
                    <label style="font-weight: 600; font-size: 0.9rem;">كلمة المرور الجديدة</label>
                    <input type="password" id="new-pass-1" placeholder="أدخل كلمة المرور الجديدة"
                        style="width: 100%; margin-top: 6px; padding: 12px 16px; border-radius: 8px;
                               border: 1px solid #e2e8f0; font-family: inherit; font-size: 0.95rem; box-sizing: border-box;">
                </div>
                <div>
                    <label style="font-weight: 600; font-size: 0.9rem;">تأكيد كلمة المرور</label>
                    <input type="password" id="new-pass-2" placeholder="أعد إدخال كلمة المرور"
                        style="width: 100%; margin-top: 6px; padding: 12px 16px; border-radius: 8px;
                               border: 1px solid #e2e8f0; font-family: inherit; font-size: 0.95rem; box-sizing: border-box;">
                </div>
                <div id="pwd-change-error" style="color: #ef4444; font-size: 0.85rem; font-weight: bold; display: none; text-align: center;"></div>
                <button id="btn-confirm-pwd" style="
                    background-color: #10b981; color: white; border: none; padding: 14px;
                    border-radius: 8px; font-family: inherit; font-size: 1rem; font-weight: 700;
                    cursor: pointer; margin-top: 4px;
                ">حفظ كلمة المرور والدخول للنظام</button>
            </div>
        </div>
    `;

    document.body.appendChild(overlay);

    document.getElementById('btn-confirm-pwd').addEventListener('click', async () => {
        const p1 = document.getElementById('new-pass-1').value;
        const p2 = document.getElementById('new-pass-2').value;
        const errEl = document.getElementById('pwd-change-error');

        if (!p1 || p1.length < 6) {
            errEl.style.display = 'block';
            errEl.textContent = 'كلمة المرور يجب أن تكون 6 أحرف على الأقل.';
            return;
        }
        if (p1 !== p2) {
            errEl.style.display = 'block';
            errEl.textContent = 'كلمتا المرور غير متطابقتين. يرجى التحقق.';
            return;
        }

        try {
            const result = await callApi("changePassword", { key: userObj.key, newPass: p1 });
            Object.assign(userObj, result.user);
            userObj.pass = undefined;
            userObj.firstLoginDone = true;
            overlay.remove();
            onSuccess();
        } catch (err) {
            errEl.style.display = 'block';
            errEl.textContent = 'تعذّر حفظ كلمة المرور: ' + err.message;
        }
    });
}


function showLoginError(msg) {
    loginErrorMsg.style.display = "block";
    loginErrorMsg.textContent = msg;
}

function doLogin(userObj) {
    state.currentUser = userObj;
    loginPortal.style.display = "none";
    appContainer.style.display = "flex";

    // Setup Topbar / Profile card
    adminName.textContent = userObj.name;
    adminRole.textContent = userObj.role;
    adminAvatar.textContent = userObj.name.charAt(0);

    // Show email if available
    const emailEl = document.getElementById("sidebar-admin-email");
    if (emailEl && userObj.email) {
        emailEl.textContent = "✉ " + userObj.email;
    }

    // Fill printable fields with logged admin
    document.getElementById("p-receipt-admin").textContent = userObj.name;
    document.getElementById("r-admin").textContent = userObj.name;

    // Auto-fill the expense authorized field with current admin name
    const expAuthField = document.getElementById("exp-authorized");
    if (expAuthField) expAuthField.value = userObj.name;

    // Show User Management tab only for the master admin (per the Users
    // sheet's isMaster flag, not a hardcoded email compare)
    const isMaster = userObj.isMaster === true;
    const navMgmt = document.getElementById("nav-usermgmt");
    if (navMgmt) navMgmt.style.display = isMaster ? "flex" : "none";

    // Show pending badge count
    if (isMaster) updatePendingBadge();

    // Show initial tab
    if (state.members.length > 0) {
        switchTab("dashboard-tab");
    } else {
        switchTab("upload-tab");
    }
}


// Log out
document.getElementById("btn-logout").addEventListener("click", () => {
    localStorage.removeItem("cems_logged_email");
    state.currentUser = null;
    loginPortal.style.display = "flex";
    appContainer.style.display = "none";
    loginPassInput.value = "";
    if (loginEmailInput) loginEmailInput.value = "";
    loginErrorMsg.style.display = "none";
    // Reset panels
    document.getElementById("panel-login").style.display = "block";
    document.getElementById("panel-register").style.display = "none";
});

// File Upload Drag and Drop
dropZone.addEventListener("click", () => btnFileSelect.click());
dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.style.borderColor = "var(--secondary)";
    dropZone.style.backgroundColor = "#d1fae5";
});
dropZone.addEventListener("dragleave", () => {
    dropZone.style.borderColor = "var(--primary)";
    dropZone.style.backgroundColor = "var(--primary-light)";
});
dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.style.borderColor = "var(--primary)";
    dropZone.style.backgroundColor = "var(--primary-light)";
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleExcelFile(files[0]);
    }
});
btnFileSelect.addEventListener("change", (e) => {
    const files = e.target.files;
    if (files.length > 0) {
        handleExcelFile(files[0]);
    }
});

function handleExcelFile(file) {
    uploadStatus.textContent = "جاري معالجة وتدقيق ملف Excel...";
    uploadStatus.style.color = "var(--accent)";

    const reader = new FileReader();
    reader.onload = async function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            state.originalWorkbook = workbook;

            const parsed = processWorkbook(workbook);

            uploadStatus.textContent = `جاري رفع بيانات الملف "${file.name}" إلى جوجل شيت...`;

            const result = await callApi("bulkImportMembers", {
                members: parsed.members,
                families: parsed.families,
                months: state.monthsList
            });
            state.members = result.members;
            state.families = result.families;
            state.monthsList = result.months.map(m => ({ key: m.key, id: m.id }));
            state.selectedLedgerMonths = result.months.filter(m => m.selected).map(m => m.key);

            uploadStatus.innerHTML = `<strong>✅ تم تحميل الملف "${file.name}" ورفعه لجوجل شيت بنجاح!</strong>`;
            uploadStatus.style.color = "var(--success)";

            populateMonthFilters();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();
            renderLedgerTable();
            renderFamiliesTable();

            setTimeout(() => {
                switchTab("dashboard-tab");
            }, 1200);

        } catch (err) {
            console.error(err);
            uploadStatus.textContent = "فشلت قراءة الملف المالي أو رفعه: " + err.message;
            uploadStatus.style.color = "var(--danger)";
        }
    };
    reader.readAsArrayBuffer(file);
}

// Helper: Find sheet header row index dynamically to handle custom formatting or padding rows
function findHeaderRow(sheet, keywords) {
    if (!sheet || !sheet['!ref']) return 0;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const maxSearchRows = Math.min(range.e.r, range.s.r + 15); // check first 15 rows
    for (let r = range.s.r; r <= maxSearchRows; r++) {
        let matchCount = 0;
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
            const cell = sheet[cellRef];
            if (cell && cell.v) {
                const val = String(cell.v).trim();
                if (keywords.some(k => val.toLowerCase() === k.toLowerCase() || val.includes(k))) {
                    matchCount++;
                }
            }
        }
        if (matchCount >= 2) {
            return r;
        }
    }
    // Fallback: search for single key name
    for (let r = range.s.r; r <= maxSearchRows; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const cellRef = XLSX.utils.encode_cell({ r: r, c: c });
            const cell = sheet[cellRef];
            if (cell && cell.v) {
                const val = String(cell.v).trim();
                if (val.includes("الاسم") || val.includes("رب العائلة") || val.includes("اسم رب العائلة")) {
                    return r;
                }
            }
        }
    }
    return 0; // fallback to row index 0
}

// Convert Excel Data into active state
function processWorkbook(workbook) {
    const sheetNames = workbook.SheetNames;
    console.log('Sheet names in workbook:', sheetNames);

    // 1. Process "التحصيل الشهري" — individual member monthly payments
    const parsedMembers = [];
    const monthlySheet = workbook.Sheets["التحصيل الشهري"];
    if (monthlySheet) {
        // Dynamically find correct header row index to skip title rows (failsafe range: 3 fallback)
        const headerRow = findHeaderRow(monthlySheet, ["الرقم", "الاسم", "رب العائلة", "المجموع"]);
        console.log("Found monthly ledger header row at index:", headerRow);
        const json = XLSX.utils.sheet_to_json(monthlySheet, { range: headerRow, defval: 0 });

        json.forEach((row, i) => {
            const name = row["الاسم"] || row["اسم العضو"] || row["الاسم كامل"] || "";
            if (!name || name === 0) return; // skip empty rows

            const member = {
                id: row["الرقم"] || row["رقم"] || (i + 1),
                name: String(name),
                parent: String(row["رب العائلة"] || row["رب عائلة"] || row["المشرف"] || ""),
                payments: {},
                sum: 0
            };

            state.monthsList.forEach(m => {
                const parsedVal = Number(row[m.key] || row[m.key.trim()] || 0);
                const val = isNaN(parsedVal) ? 0 : parsedVal;
                member.payments[m.key] = val;
                member.sum += val;
            });

            parsedMembers.push(member);
        });
    }

    // 2. Process "وصولات العائلات" — the primary family receipts sheet
    // Real columns (confirmed from file):
    // رقم العائلة | اسم رب العائلة | عدد الأفراد | الاشتراك الشهري (₪) | إجمالي المدفوع (₪) | أفراد العائلة
    const parsedFamilies = [];

    // Try all known sheet name variations
    const familySheetName = ["وصولات العائلات", "وصل الدفع", "وصولات الدفع", "وصولات"].find(n => workbook.Sheets[n]);

    if (familySheetName) {
        const sheet = workbook.Sheets[familySheetName];
        // Dynamically find correct header row index to skip title rows (failsafe range: 3 fallback)
        const headerRow = findHeaderRow(sheet, ["رقم العائلة", "معرف العائلة", "اسم رب العائلة", "رب العائلة", "عدد الأفراد"]);
        console.log("Found family registry header row at index:", headerRow);
        const json = XLSX.utils.sheet_to_json(sheet, { range: headerRow, defval: "" });

        json.forEach((row, i) => {
            // Head name — try various column name spellings
            const headName = String(
                row["اسم رب العائلة"] || row["رب العائلة"] || row["اسم العائلة"] || row["الرأس"] ||
                row["اسم رب الاسره"] || row["اسم رب الاسرة"] || ""
            ).trim();
            if (!headName) return;

            // Members list — stored as comma-separated string in Excel
            const membersCellRaw = String(
                row["أفراد العائلة"] || row["افراد العائله"] || row["قائمه افراد العائله"] ||
                row["الأفراد التفصيلي"] || row["الأفراد"] || ""
            ).trim();

            // Parse the members list — split by Arabic comma or comma
            const membersArr = membersCellRaw
                ? membersCellRaw.split(/،|,/).map(m => m.trim()).filter(m => m)
                : [headName];

            // Monthly subscription amount (₪ symbol stripped)
            const subRaw = String(row["الاشتراك الشهري (₪)"] || row["الاشتراك الشهري"] || row["الاشتراك المفترض (شيكل/شري)"] || row["الاشتراك"] || 0);
            const subscription = parseInt(subRaw.replace(/[^\d]/g, '')) || (membersArr.length * 10);

            // Total paid
            const paidRaw = String(row["إجمالي المدفوع (₪)"] || row["إجمالي المدفوع"] || row["إجمالي المسدد"] || row["المسدد"] || 0);
            const totalPaid = parseInt(paidRaw.replace(/[^\d]/g, '')) || 0;

            parsedFamilies.push({
                familyId: row["رقم العائلة"] || row["معرف العائلة"] || row["رقم"] || (i + 1),
                headName: headName,
                memberCount: parseInt(String(row["عدد الأفراد"] || row["عدد أفراد الأسرة"] || membersArr.length)) || membersArr.length,
                subscription: subscription,
                totalPaid: totalPaid,
                membersList: membersArr.join('، '), // plain string
                membersArr: membersArr               // array for receipt rendering
            });
        });
    } else {
        // Fallback: group dynamically from التحصيل الشهري if no receipts sheet
        const grouped = {};
        parsedMembers.forEach(m => {
            const familyName = m.parent || m.name;
            if (!grouped[familyName]) {
                grouped[familyName] = { members: [], totalPaid: 0 };
            }
            grouped[familyName].members.push(m.name);
            grouped[familyName].totalPaid += m.sum;
        });

        Object.keys(grouped).forEach((headName, index) => {
            const mCount = grouped[headName].members.length;
            parsedFamilies.push({
                familyId: index + 1,
                headName: headName,
                memberCount: mCount,
                subscription: mCount * 10,
                totalPaid: grouped[headName].totalPaid,
                membersList: grouped[headName].members.join('، '),
                membersArr: grouped[headName].members
            });
        });
    }

    return { members: parsedMembers, families: parsedFamilies };
}

// Loads every piece of app data from the Google Sheet backend in one round
// trip and populates `state`. Called once on startup, and again by the
// "reload from sheet" button (see below) or after an Excel bulk import.
async function loadDataFromServer() {
    const data = await callApi("getAllData", {});
    state.members = data.members;
    state.families = data.families;
    state.monthsList = data.months.map(m => ({ key: m.key, id: m.id }));
    state.selectedLedgerMonths = data.months.filter(m => m.selected).map(m => m.key);
    state.expenses = data.expenses;
    state.users = data.users;
    state.pendingUsers = data.pendingUsers;
    state.settings = data.settings;

    const statusEl = document.getElementById("upload-status");
    if (statusEl) {
        if (state.members.length > 0) {
            statusEl.innerHTML = `<strong>✅ البيانات محملة من جوجل شيت</strong><br>${state.members.length} فرد، ${state.families.length} عائلة. يمكنك رفع ملف Excel لاستبدال البيانات بالكامل.`;
        } else {
            statusEl.innerHTML = `<strong>لا توجد بيانات أعضاء بعد.</strong><br>ارفع ملف Excel هنا لتعبئة قاعدة البيانات على جوجل شيت لأول مرة.`;
        }
        statusEl.style.color = "var(--success)";
    }
    updateIndicators();
}

// Reload button — re-fetches everything from the Sheet (e.g. after another
// user, or a direct edit on the Sheet itself, changed something).
document.getElementById("btn-clear-cache").addEventListener("click", async () => {
    if (!confirm("سيتم إعادة تحميل كل البيانات من جوجل شيت، وسيتم فقدان أي تعديل لم يُحفظ بعد. متابعة؟")) return;
    try {
        uploadStatus.textContent = "جاري إعادة التحميل من جوجل شيت...";
        uploadStatus.style.color = "var(--accent)";
        await loadDataFromServer();

        populateMonthFilters();
        populateAddMemberFamilies();
        populateReceiptFamilies();
        renderCharts();
        renderLedgerTable();
        renderFamiliesTable();

        alert("✅ تمت إعادة تحميل البيانات من جوجل شيت بنجاح.");
    } catch (err) {
        alert("تعذّرت إعادة التحميل: " + err.message);
    }
});

// Update KPIs and details
function updateIndicators() {
    // Always update the sidebar data status badge
    const dot = document.getElementById("sidebar-data-dot");
    const label = document.getElementById("sidebar-data-label");
    const statusBox = document.getElementById("sidebar-data-status");
    
    if (state.members.length > 0 && dot && label && statusBox) {
        // Data is loaded — show green
        dot.style.background = "#10b981";
        label.textContent = `البيانات محملة (${state.members.length} فرد)`;
        statusBox.style.backgroundColor = "rgba(16, 185, 129, 0.15)";
        statusBox.style.border = "1px solid rgba(16, 185, 129, 0.3)";
        statusBox.style.color = "#a7f3d0";
    } else if (dot && label && statusBox) {
        // No data
        dot.style.background = "#ef4444";
        label.textContent = "لا توجد بيانات محملة";
        statusBox.style.backgroundColor = "rgba(239, 68, 68, 0.15)";
        statusBox.style.border = "1px solid rgba(239, 68, 68, 0.3)";
        statusBox.style.color = "#fca5a5";
    }

    if (state.members.length === 0) return;


    // Total members count
    document.getElementById("stat-total-members").textContent = state.members.length;

    // Total collected
    let totalCollected = 0;
    state.members.forEach(m => totalCollected += m.sum);
    document.getElementById("stat-total-collected").textContent = totalCollected + " شيكل";

    // Expected Monthly subscription (each active member pays 10 shikel)
    const monthlyExpected = state.members.length * 10;
    document.getElementById("stat-expected-monthly").textContent = monthlyExpected + " شيكل / شهر";

    // Overall Commitment Rate (families that have contribution matches)
    let totalFamilies = state.families.length;
    let committedFamilies = 0;
    state.families.forEach(f => {
        if (f.totalPaid > 0) {
            committedFamilies++;
        }
    });

    const rate = totalFamilies > 0 ? Math.round((committedFamilies / totalFamilies) * 100) : 0;
    document.getElementById("stat-commitment-rate").textContent = rate + "%";

    // Expenses and net balance
    const totalExpenses = state.expenses.reduce((sum, e) => sum + (e.amount || 0), 0);
    const netBalance = totalCollected - totalExpenses;
    const expEl = document.getElementById("stat-total-expenses");
    const netEl = document.getElementById("stat-net-balance");
    if (expEl) expEl.textContent = totalExpenses + " شيكل";
    if (netEl) {
        netEl.textContent = netBalance + " شيكل";
        netEl.style.color = netBalance >= 0 ? "#10b981" : "#ef4444";
    }
}

// Render Dashboard Charts
function renderCharts() {
    if (state.members.length === 0) return;

    // Destroy existing charts to prevent memory leak duplicates
    if (monthlyChartInstance) monthlyChartInstance.destroy();
    if (commitmentChartInstance) commitmentChartInstance.destroy();

    // 1. Monthly Chart Collections
    const labels = state.monthsList.map(m => m.key);
    const dataCollected = [];
    const dataExpected = [];

    state.monthsList.forEach(m => {
        let collectedForMonth = 0;
        state.members.forEach(member => {
            collectedForMonth += (member.payments[m.key] || 0);
        });
        dataCollected.push(collectedForMonth);
        dataExpected.push(state.members.length * 10);
    });

    const ctx1 = document.getElementById("monthly-chart").getContext("2d");
    monthlyChartInstance = new Chart(ctx1, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'المبلغ المحصل الفعلي (شيكل)',
                    data: dataCollected,
                    backgroundColor: 'rgba(16, 185, 129, 0.85)',
                    borderColor: '#10b981',
                    borderWidth: 1
                },
                {
                    label: 'المبلغ المستهدف (10 شيكل/عضو)',
                    data: dataExpected,
                    type: 'line',
                    borderColor: '#0f766e',
                    backgroundColor: 'transparent',
                    borderWidth: 2,
                    borderDash: [5, 5]
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true }
            },
            plugins: {
                legend: { position: 'bottom', labels: { font: { family: 'Cairo' } } }
            }
        }
    });

    // 2. Commitment Rate Doughnut Chart
    let totalFamilies = state.families.length;
    let committedFamilies = 0;
    state.families.forEach(f => {
        if (f.totalPaid > 0) committedFamilies++;
    });
    let uncommittedFamilies = totalFamilies - committedFamilies;

    const ctx2 = document.getElementById("rate-pie-chart").getContext("2d");
    commitmentChartInstance = new Chart(ctx2, {
        type: 'doughnut',
        data: {
            labels: ['عائلات ملتزمة', 'عائلات متأخرة الدفع'],
            datasets: [{
                data: [committedFamilies, uncommittedFamilies],
                backgroundColor: ['#10b981', '#f59e0b'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'bottom', labels: { font: { family: 'Cairo' } } }
            }
        }
    });

    updateIndicators();
}

// -------------------------------------------------------------
// Interactive Monthly Ledger Table Controls & Render
// -------------------------------------------------------------
const ledgerSearch = document.getElementById("ledger-search");
const ledgerFilterMonth = document.getElementById("ledger-filter-month");
const ledgerFilterStatus = document.getElementById("ledger-filter-status");

// Debounce helper to prevent freezing during fast typing
function debounce(func, delay = 250) {
    let timeoutId;
    return (...args) => {
        clearTimeout(timeoutId);
        timeoutId = setTimeout(() => {
            func(...args);
        }, delay);
    };
}

if (ledgerSearch) {
    ledgerSearch.addEventListener("input", debounce(renderLedgerTable, 250));
}
if (ledgerFilterMonth) {
    ledgerFilterMonth.addEventListener("change", (e) => {
        const val = e.target.value;
        const ops = document.getElementById("month-ops-controls");
        if (ops) {
            ops.style.display = val === "all" ? "none" : "flex";
        }
        renderLedgerTable();
    });
}
if (ledgerFilterStatus) {
    ledgerFilterStatus.addEventListener("change", renderLedgerTable);
}

function renderLedgerHeaders() {
    const thead = document.getElementById("ledger-headers");
    if (!thead) return;
    
    thead.innerHTML = "";
    const tr = document.createElement("tr");
    
    const thId = document.createElement("th");
    thId.textContent = "الرقم";
    tr.appendChild(thId);
    
    const thName = document.createElement("th");
    thName.textContent = "الاسم كامل";
    tr.appendChild(thName);
    
    const filterMonth = ledgerFilterMonth ? ledgerFilterMonth.value : "all";
    
    if (filterMonth === "all") {
        const visibleMonths = state.monthsList.filter(m => state.selectedLedgerMonths.includes(m.key));
        visibleMonths.forEach(m => {
            const thMonth = document.createElement("th");
            const shortKey = m.key.replace("2026", "26").replace("2027", "27").replace("2028", "28");
            thMonth.textContent = shortKey;
            tr.appendChild(thMonth);
        });
    } else {
        const thMonth = document.createElement("th");
        const shortKey = filterMonth.replace("2026", "26").replace("2027", "27").replace("2028", "28");
        thMonth.textContent = shortKey;
        tr.appendChild(thMonth);
    }
    
    const thSum = document.createElement("th");
    thSum.textContent = "المجموع";
    tr.appendChild(thSum);
    
    const thAction = document.createElement("th");
    thAction.textContent = "إجراء";
    thAction.style.textAlign = "center";
    tr.appendChild(thAction);
    
    thead.appendChild(tr);
}

function renderLedgerTable() {
    renderLedgerHeaders();

    const tbody = document.getElementById("ledger-rows");
    if (!tbody) return;
    tbody.innerHTML = "";

    const filterMonth = ledgerFilterMonth ? ledgerFilterMonth.value : "all";
    const visibleMonthsCount = filterMonth === "all" ? (state.selectedLedgerMonths || []).length : 1;
    const colCount = 2 + visibleMonthsCount + 2;

    if (state.members.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${colCount}" style="text-align: center;">يرجى رفع ملف البيانات للبدء.</td></tr>`;
        return;
    }

    const rawSearch = ledgerSearch ? ledgerSearch.value : "";
    const searchQuery = normalizeArabicName(rawSearch);
    const filterStatus = ledgerFilterStatus ? ledgerFilterStatus.value : "all";

    // Set scroll container layouts dynamically
    const table = document.getElementById("table-ledger");
    const scrollInfo = document.querySelector(".table-scroll-info");
    if (table) {
        if (filterMonth === "all" && visibleMonthsCount > 4) {
            table.style.minWidth = "1200px";
            if (scrollInfo) scrollInfo.style.display = "flex";
        } else {
            table.style.minWidth = "auto";
            if (scrollInfo) scrollInfo.style.display = "none";
        }
    }

    let filtered = state.members.filter(m => {
        // Search matching using Arabic normalization to handle typing variations
        const normName = normalizeArabicName(m.name);
        const normParent = normalizeArabicName(m.parent);
        const matchSearch = normName.includes(searchQuery) || normParent.includes(searchQuery);
        
        // Status & Month matching
        let matchStatus = true;
        if (filterMonth !== "all") {
            const TARGET_KEY = filterMonth;
            const paidThisMonth = (m.payments[TARGET_KEY] || 0) > 0;
            if (filterStatus === "paid") {
                matchStatus = paidThisMonth;
            } else if (filterStatus === "unpaid") {
                matchStatus = !paidThisMonth;
            }
        } else {
            if (filterStatus === "paid") {
                matchStatus = m.sum >= (state.monthsList.length * 10);
            } else if (filterStatus === "unpaid") {
                matchStatus = m.sum < (state.monthsList.length * 10);
            }
        }

        return matchSearch && matchStatus;
    });

    if (filtered.length === 0) {
         tbody.innerHTML = `<tr><td colspan="${colCount}" style="text-align: center;">لم يتم العثور على أي مدخلات تطابق معايير البحث.</td></tr>`;
         return;
    }

    filtered.forEach(member => {
        const tr = document.createElement("tr");

        // ID & Name
        const tdId = document.createElement("td");
        tdId.textContent = member.id;
        tr.appendChild(tdId);

        const tdName = document.createElement("td");
        tdName.innerHTML = `<strong>${member.name}</strong>`;
        tr.appendChild(tdName);

        // Month Payments Column(s)
        if (filterMonth === "all") {
            const visibleMonths = state.monthsList.filter(m => state.selectedLedgerMonths.includes(m.key));
            visibleMonths.forEach(m => {
                const tdMonth = document.createElement("td");
                tdMonth.style.textAlign = "center";
                const val = member.payments[m.key] || 0;

                const chk = document.createElement("input");
                chk.type = "checkbox";
                chk.className = "ledger-checkbox";
                chk.checked = val > 0;
                
                chk.addEventListener("change", async (e) => {
                    const isChecked = e.target.checked;
                    const paymentVal = isChecked ? 10 : 0;
                    chk.disabled = true;

                    try {
                        const result = await callApi("setMemberPayment", { id: member.id, month: m.key, amount: paymentVal });
                        state.members = result.members;
                        state.families = result.families;
                        renderLedgerTable();
                        renderCharts();
                        populateMonthFilters();
                    } catch (err) {
                        e.target.checked = !isChecked;
                        alert("تعذّر حفظ الدفعة: " + err.message);
                        chk.disabled = false;
                    }
                });

                tdMonth.appendChild(chk);
                tr.appendChild(tdMonth);
            });
        } else {
            const tdMonth = document.createElement("td");
            tdMonth.style.textAlign = "center";
            const val = member.payments[filterMonth] || 0;

            const chk = document.createElement("input");
            chk.type = "checkbox";
            chk.className = "ledger-checkbox";
            chk.checked = val > 0;
            
            chk.addEventListener("change", async (e) => {
                const isChecked = e.target.checked;
                const paymentVal = isChecked ? 10 : 0;
                chk.disabled = true;

                try {
                    const result = await callApi("setMemberPayment", { id: member.id, month: filterMonth, amount: paymentVal });
                    state.members = result.members;
                    state.families = result.families;
                    renderLedgerTable();
                    renderCharts();
                    populateMonthFilters();
                } catch (err) {
                    e.target.checked = !isChecked;
                    alert("تعذّر حفظ الدفعة: " + err.message);
                    chk.disabled = false;
                }
            });

            tdMonth.appendChild(chk);
            tr.appendChild(tdMonth);
        }

        // Totals column
        const tdSum = document.createElement("td");
        tdSum.style.fontWeight = "bold";
        tdSum.style.color = "var(--secondary)";
        tdSum.textContent = member.sum + " شيكل";
        tr.appendChild(tdSum);

        // Actions column with Edit (✏️) and Delete (🗑️)
        const tdAction = document.createElement("td");
        tdAction.style.textAlign = "center";
        tdAction.style.whiteSpace = "nowrap";
        
        // Edit Button
        const btnEdit = document.createElement("button");
        btnEdit.innerHTML = "✏️";
        btnEdit.title = "تعديل العضو";
        btnEdit.style.cssText = "background: rgba(16, 185, 129, 0.12); color: var(--primary); border: 1px solid rgba(16, 185, 129, 0.3); padding: 4px 8px; border-radius: 6px; cursor: pointer; font-size: 0.85rem; margin-left: 6px;";
        btnEdit.addEventListener("click", () => {
            openEditMemberModal(member.id);
        });
        tdAction.appendChild(btnEdit);
        
        // Delete Button
        const btnDel = document.createElement("button");
        btnDel.innerHTML = "🗑️";
        btnDel.title = "حذف العضو";
        btnDel.style.cssText = "background: rgba(239, 68, 68, 0.12); color: #ef4444; border: 1px solid rgba(239, 68, 68, 0.3); padding: 4px 8px; border-radius: 6px; cursor: pointer; font-size: 0.85rem;";
        btnDel.addEventListener("click", () => {
            deleteMember(member.id, member.name);
        });
        tdAction.appendChild(btnDel);
        
        tr.appendChild(tdAction);

        tbody.appendChild(tr);
    });
}

function normalizeArabicName(name) {
    if (!name) return "";
    return String(name)
        .trim()
        .toLowerCase()
        .replace(/[أإآ]/g, "ا")
        .replace(/ة/g, "ه")
        .replace(/ى/g, "ي")
        .replace(/\s+/g, " ");
}

function recalculateFamilyTotals(familyName) {
    if (!familyName) return;
    
    const normFamilyName = normalizeArabicName(familyName);
    let familySum = 0;
    state.members.forEach(m => {
        if (normalizeArabicName(m.parent) === normFamilyName) {
            familySum += m.sum;
        }
    });

    // Update inside family record
    const familyRecord = state.families.find(f => normalizeArabicName(f.headName) === normFamilyName);
    if (familyRecord) {
        familyRecord.totalPaid = familySum;
    }
}

// -------------------------------------------------------------
// Excel Re-export Logic
// -------------------------------------------------------------
document.getElementById("btn-export-excel").addEventListener("click", () => {
    if (state.members.length === 0) {
        alert("لا توجد بيانات متاحة للتصدير.");
        return;
    }

    try {
        // Compile ledger values back into standard worksheets
        const newLedgerRows = state.members.map(member => {
            const rowOutput = {
                "الرقم": member.id,
                "الاسم": member.name,
                "رب العائلة": member.parent
            };
            state.monthsList.forEach(m => {
                rowOutput[m.key] = member.payments[m.key] || 0;
            });
            rowOutput["المجموع"] = member.sum;
            return rowOutput;
        });

        // Compile family summaries
        const newFamilyRows = state.families.map(f => {
            return {
                "معرف العائلة": f.familyId,
                "رب العائلة": f.headName,
                "عدد أفراد الأسرة": f.memberCount,
                "الاشتراك المفترض (شيكل/شري)": f.subscription,
                "إجمالي المسدد": f.totalPaid,
                "قائمه افراد العائله": f.membersList
            };
        });

        // Create workbook instance
        const wb = XLSX.utils.book_new();

        // 1. Add Dashboard mockup values
        const dashboardMock = [
            ["صندوق تحصيل عائلة آل اطفيحة - لوحة المعلومات"],
            ["إجمالي المشتركين", state.members.length],
            ["إجمالي المبالغ التراكمية المحصلة", state.members.reduce((acc, curr) => acc + curr.sum, 0)],
            ["مستحقات الاشتراك الشهري الإجمالي", state.members.length * 10]
        ];
        const wsDashboard = XLSX.utils.aoa_to_sheet(dashboardMock);
        XLSX.utils.book_append_sheet(wb, wsDashboard, "لوحة المعلومات");

        // 2. Add main ledger
        const wsLedger = XLSX.utils.json_to_sheet(newLedgerRows);
        XLSX.utils.book_append_sheet(wb, wsLedger, "التحصيل الشهري");

        // 3. Add family summaries
        const wsFamilies = XLSX.utils.json_to_sheet(newFamilyRows);
        XLSX.utils.book_append_sheet(wb, wsFamilies, "وصولات العائلات");

        // 4. Add expenses history
        const newExpenseRows = state.expenses.map((e, idx) => {
            return {
                "الرقم": idx + 1,
                "التاريخ": e.date,
                "التصنيف": e.category,
                "سبب الصرف": e.reason,
                "المبلغ (شيكل)": e.amount,
                "المرفق": e.attachment ? `${e.attachment.name} (مرفق)` : "لا يوجد",
                "معتمد من": e.authorized
            };
        });
        const wsExpenses = XLSX.utils.json_to_sheet(newExpenseRows);
        XLSX.utils.book_append_sheet(wb, wsExpenses, "سجل الصرف");

        // Export download buffer triggers
        XLSX.writeFile(wb, "تحديث_احصاء_ال_اطفيحه.xlsx");

    } catch (err) {
        console.error(err);
        alert("حدثت مشكلة أثناء محاولة تصدير ملف excel: " + err.message);
    }
});

// -------------------------------------------------------------
// Family Roster Table Rendering & Quick Options
// -------------------------------------------------------------
const familySearchInput = document.getElementById("family-search");
if (familySearchInput) {
    familySearchInput.addEventListener("input", debounce(renderFamiliesTable, 250));
}

function renderFamiliesTable() {
    const tbody = document.getElementById("family-rows");
    tbody.innerHTML = "";

    if (state.families.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align: center;">لم يتم توثيق أي بيانات عائلات.</td></tr>`;
        return;
    }

    const rawQuery = familySearchInput ? familySearchInput.value : "";
    const query = normalizeArabicName(rawQuery);
    const filtered = state.families.filter(f => normalizeArabicName(f.headName).includes(query));

    filtered.forEach(family => {
        const tr = document.createElement("tr");

        tr.innerHTML = `
            <td>${family.familyId}</td>
            <td><strong>${family.headName}</strong></td>
            <td>${family.memberCount} أفراد</td>
            <td>${family.subscription} شيكل</td>
            <td style="color: var(--secondary); font-weight: bold;">${family.totalPaid} شيكل</td>
            <td>
                <button class="btn btn-secondary btn-sm" onclick="configureQuickReceipt('${family.headName}')">إصدار سند قبض</button>
            </td>
            <td>
                <button class="btn btn-secondary btn-sm" style="background-color: #25d366; color: white;" onclick="shareWhatsAppReceipt('${family.headName}')">واتساب</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// Trigger view switches for receipts
window.configureQuickReceipt = function(headName) {
    switchTab("receipt-tab");
    
    // Choose appropriate family in selector
    const familySel = document.getElementById("receipt-select-family");
    familySel.value = headName;
    familySel.dispatchEvent(new Event("change"));
};

// WhatsApp text generator triggers
window.shareWhatsAppReceipt = function(headName) {
    const family = state.families.find(f => f.headName === headName);
    if (!family) return;

    // Formatting receipt message in elegant arabic format
    const textMsg = `*صندوق تحصيل عائلة آل اطفيحة* 📜\n\nالسيد الفاضل: *${family.headName}* المحترم\nنود إعلامكم بأن إجمالي المبالغ المسددة لصندوق العائلة التراكمي قد بلغت: *${family.totalPaid} شيكل*.\nعدد أفراد العائلة المسجلين: ${family.memberCount} أفراد.\nالاشتراك المفترض شهرياً: ${family.subscription} شيكل.\n\n*شكراً لالتزامكم ودعمكم لصندوق العائلة المالي المستمر.* 🙏`;
    
    // Copy directly to clipboard
    navigator.clipboard.writeText(textMsg)
        .then(() => alert(`تم نسخ نص رسالة تفاصيل العائلة "${headName}" لعلامة الحوافظ بنجاح لتتمكن من إرسالها بالواتساب!`))
        .catch(err => alert("عذراً، فشل نسخ النص: " + err));
};

// -------------------------------------------------------------
// Interactive Receipt Generator Actions
// -------------------------------------------------------------
const receiptSelectFamily = document.getElementById("receipt-select-family");
const receiptNo = document.getElementById("receipt-no");
const receiptAmount = document.getElementById("receipt-amount");
const receiptNotes = document.getElementById("receipt-notes");

receiptSelectFamily.addEventListener("change", updatePrintPreview);
receiptNo.addEventListener("input", updatePrintPreview);
receiptAmount.addEventListener("input", updatePrintPreview);
receiptNotes.addEventListener("input", updatePrintPreview);

// Listen to checked months on receipt workspace
document.querySelectorAll(".receipt-month-chk").forEach(chk => {
    chk.addEventListener("change", () => {
         // Auto-calculate receipt amount based on selected months
         const headName = receiptSelectFamily.value;
         const family = state.families.find(f => f.headName === headName);
         if (family) {
              const checkedCount = document.querySelectorAll(".receipt-month-chk:checked").length;
              // Monthly fee = family subscription monthly
              const expectedAmt = family.subscription * checkedCount;
              receiptAmount.value = expectedAmt;
         }
         updatePrintPreview();
    });
});

function populateReceiptFamilies() {
    receiptSelectFamily.innerHTML = `<option value="">-- اختر رب العائلة --</option>`;
    state.families.forEach(f => {
        const opt = document.createElement("option");
        opt.value = f.headName;
        opt.textContent = `${f.headName} (${f.memberCount} أفراد)`;
        receiptSelectFamily.appendChild(opt);
    });
}

function updatePrintPreview() {
    const headName = receiptSelectFamily.value;
    const amountVal = receiptAmount.value || 0;
    const rNo = receiptNo.value || "1001";
    const noteText = receiptNotes.value || "لا يوجد ملاحظات إضافية.";

    const family = state.families.find(f => f.headName === headName);
    const membersRow = document.getElementById("p-receipt-members-row");
    const membersVal = document.getElementById("p-receipt-members");

    if (family && family.membersList) {
        if (membersRow) membersRow.style.display = "flex";
        if (membersVal) membersVal.textContent = family.membersList;
    } else {
        if (membersRow) membersRow.style.display = "none";
    }

    // Get checked months names
    const checkedMonths = [];
    document.querySelectorAll(".receipt-month-chk:checked").forEach(chk => {
         const m = state.monthsList.find(mon => mon.id == chk.value || mon.key == chk.value);
         if (m) checkedMonths.push(m.key);
    });
    const monthsStr = checkedMonths.length > 0 ? checkedMonths.join("، ") : "لم يتم تحديد شهور.";

    // Fill elements
    document.getElementById("p-receipt-no").textContent = rNo;
    document.getElementById("p-receipt-sender").textContent = headName ? headName : "------------------------------";
    document.getElementById("p-receipt-amount-val").textContent = amountVal;
    document.getElementById("p-receipt-months-names").textContent = monthsStr;
    document.getElementById("p-receipt-notes-val").textContent = noteText;
}

// Receipt text copy handler
document.getElementById("btn-copy-receipt-txt").addEventListener("click", () => {
    const headName = receiptSelectFamily.value;
    if (!headName) {
        alert("الرجاء اختيار العائلة أولاً.");
        return;
    }

    const amountVal = receiptAmount.value || 0;
    const rNo = receiptNo.value || "1001";
    const family = state.families.find(f => f.headName === headName);
    const familyMembersStr = (family && family.membersList) ? `\n*أفراد الأسرة المشمولين*: ${family.membersList}` : "";

    const checkedMonths = [];
    document.querySelectorAll(".receipt-month-chk:checked").forEach(chk => {
         const m = state.monthsList.find(mon => mon.id == chk.value || mon.key == chk.value);
         if (m) checkedMonths.push(m.key);
    });
    const monthsStr = checkedMonths.length > 0 ? checkedMonths.join("، ") : "مستحقات عامة";

    const textMsg = `*تأكيد سند قبض مالي - صندوق آل اطفيحة* 🧾\n\nرقم السند: *${rNo}*\nوصلنا من السيد: *${headName}*${familyMembersStr}\nمبلغ وقدره: *${amountVal} شيكل*\nعن اشتراك شهور: *${monthsStr}*\nمستلم السند: *${state.currentUser ? state.currentUser.name : 'اللجنة المالية'}*\n\nنشكر لكم مساهمتكم الكريمة لدعم روابط العشيرة. 🕊️`;

    navigator.clipboard.writeText(textMsg)
        .then(() => alert("تم نسخ رسالة سند القبض بنجاح!"))
        .catch(err => alert("خطأ في النسخ: " + err));
});

// Printing physical receipts triggers
document.getElementById("btn-print-receipt").addEventListener("click", () => {
    const headName = receiptSelectFamily.value;
    if (!headName) {
        alert("الرجاء اختيار العائلة دافعة المبلغ للطباعة.");
        return;
    }

    // Clone receipt card view, copy to print only container
    const printableArea = document.getElementById("printable-receipt-card").innerHTML;
    const printContainer = document.getElementById("print-recipient-only");
    
    printContainer.innerHTML = printableArea;
    printContainer.style.display = "block";

    // Run browser print dialog
    window.print();

    printContainer.style.display = "none";
});

// Record payment from receipt generator directly
document.getElementById("btn-save-receipt-payment").addEventListener("click", async () => {
    const headName = receiptSelectFamily.value;
    if (!headName) {
        alert("الرجاء اختيار العائلة أولاً.");
        return;
    }

    const checkedChks = document.querySelectorAll(".receipt-month-chk:checked");
    if (checkedChks.length === 0) {
        alert("الرجاء تحديد شهر واحد على الأقل لتسجيل عملية الدفع.");
        return;
    }

    const family = state.families.find(f => f.headName === headName);
    if (!family) {
        alert("خطأ: لم يتم العثور على سجل العائلة.");
        return;
    }

    const checkedMonthsKeys = [];
    checkedChks.forEach(chk => {
        const m = state.monthsList.find(mon => mon.id == chk.value || mon.key == chk.value);
        if (m) checkedMonthsKeys.push(m.key);
    });

    const btn = document.getElementById("btn-save-receipt-payment");
    btn.disabled = true;
    try {
        const result = await callApi("setFamilyPayments", { headName, months: checkedMonthsKeys });
        state.members = result.members;
        state.families = result.families;

        renderCharts();

        alert(`✅ تم بنجاح تسجيل دفعة العائلة: (${headName}) لشهور: [${checkedMonthsKeys.join("، ")}]. تم تحديث الداشبورد والمخططات البيانية!`);

        checkedChks.forEach(chk => chk.checked = false);
        receiptAmount.value = 0;
        receiptNotes.value = "";
        updatePrintPreview();
    } catch (err) {
        alert("تعذّر تسجيل الدفعة: " + err.message);
    } finally {
        btn.disabled = false;
    }
});

// -------------------------------------------------------------
// Interactive Monthly Reports Actions
// -------------------------------------------------------------
const reportSelectMonth = document.getElementById("report-select-month");
reportSelectMonth.addEventListener("change", renderMonthlyReport);

function renderMonthlyReport() {
    if (state.members.length === 0) return;

    const monthVal = reportSelectMonth.value;
    const month = state.monthsList.find(m => m.key == monthVal || m.id == monthVal);
    if (!month) return;

    const mKey = month.key;

    // Fill Title
    document.getElementById("r-title").textContent = `تقرير صندوق التحصيل الشهري - ${mKey}`;

    // Fill Date & Admin dynamically
    const todayStr = new Date().toLocaleDateString('ar-EG', { year: 'numeric', month: 'numeric', day: 'numeric' });
    document.getElementById("r-date").textContent = todayStr;
    document.getElementById("r-admin").textContent = state.currentUser ? state.currentUser.name : "اللجنة المالية";

    // Refresh digital signatures for printable report monthly sheet
    if (window.updateSignaturesDisplay) {
        window.updateSignaturesDisplay();
    }

    let totalExpected = 0;
    let totalCollected = 0;
    
    const paidList = [];
    const unpaidList = [];

    state.members.forEach(member => {
        totalExpected += 10;
        const val = member.payments[mKey] || 0;
        
        if (val > 0) {
            totalCollected += val;
            paidList.push(member);
        } else {
            unpaidList.push(member);
        }
    });

    const totalPending = totalExpected - totalCollected;
    const ratePercentage = totalExpected > 0 ? Math.round((totalCollected / totalExpected) * 100) : 0;

    // Set UI summary tags
    document.getElementById("r-expected").textContent = totalExpected + " شيكل";
    document.getElementById("r-collected").textContent = totalCollected + " شيكل";
    document.getElementById("r-pending").textContent = totalPending + " شيكل";
    document.getElementById("r-rate").textContent = ratePercentage + "%";

    document.getElementById("r-paid-count").textContent = paidList.length;
    document.getElementById("r-unpaid-count").textContent = unpaidList.length;

    // Render lists tables
    const tbodyPaid = document.getElementById("r-rows-paid");
    tbodyPaid.innerHTML = "";
    if(paidList.length === 0) {
        tbodyPaid.innerHTML = `<tr><td colspan="4" style="text-align: center;">لا مسددين بعد في هذا الشهر.</td></tr>`;
    } else {
        paidList.forEach((m, idx) => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td><strong>${m.name}</strong></td>
                <td>${m.parent || "-"}</td>
                <td style="color: var(--success); font-weight: bold;">+10 شيكل</td>
            `;
            tbodyPaid.appendChild(tr);
        });
    }

    const tbodyUnpaid = document.getElementById("r-rows-unpaid");
    tbodyUnpaid.innerHTML = "";
    if (unpaidList.length === 0) {
        tbodyUnpaid.innerHTML = `<tr><td colspan="4" style="text-align: center;">كل أفراد العائلة مسددين بالكامل لهذا الشهر.</td></tr>`;
    } else {
        unpaidList.forEach((m, idx) => {
            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td><strong>${m.name}</strong></td>
                <td>${m.parent || "-"}</td>
                <td style="color: var(--danger); font-weight: bold;">10 شيكل</td>
            `;
            tbodyUnpaid.appendChild(tr);
        });
    }

    // --- Render monthly expenses in the report ---
    // Match expenses whose date falls within the selected month+year
    const monthExpenses = state.expenses.filter(e => {
        if (!e.date) return false;
        const parts = e.date.split('-');
        if (parts.length < 3) return false;
        const expYear = parseInt(parts[0]);
        const expMonth = parseInt(parts[1]);
        const yearMatch = month.key.match(/\d{4}/);
        const targetYear = yearMatch ? parseInt(yearMatch[0]) : null;
        return expMonth === month.id && (targetYear ? expYear === targetYear : true);
    });

    const rExpSection = document.getElementById("r-expenses-section");
    const rExpRows = document.getElementById("r-exp-rows");
    const rExpCount = document.getElementById("r-exp-count");
    const rExpTotal = document.getElementById("r-exp-total");

    if (rExpSection && rExpRows) {
        if (monthExpenses.length > 0) {
            rExpSection.style.display = "block";
            rExpRows.innerHTML = "";
            let expSum = 0;
            monthExpenses.forEach((e, idx) => {
                expSum += e.amount;
                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td style="padding:5px 10px;">${idx + 1}</td>
                    <td style="padding:5px 10px;">${e.date}</td>
                    <td style="padding:5px 10px;">${e.reason}</td>
                    <td style="padding:5px 10px;">${e.category}</td>
                    <td style="padding:5px 10px; color:#f59e0b; font-weight:700;">${e.amount} شيكل</td>
                    <td style="padding:5px 10px;">${e.authorized}</td>
                `;
                rExpRows.appendChild(tr);
            });
            if (rExpCount) rExpCount.textContent = monthExpenses.length;
            if (rExpTotal) rExpTotal.textContent = expSum;
        } else {
            rExpSection.style.display = "none";
        }
    }
}

// Print Monthly Reports sheet
document.getElementById("btn-print-report").addEventListener("click", () => {
    if (state.members.length === 0) {
        alert("لا توجد بيانات كافية لطباعة التقرير.");
        return;
    }
    
    // Hide components on page to leave only report container card for print dialog
    window.print();
});

// -------------------------------------------------------------
// Dynamic Month and Member Addition Logic
// -------------------------------------------------------------
function populateMonthFilters() {
    const filterSelect = document.getElementById("ledger-filter-month");
    const reportSelect = document.getElementById("report-select-month");
    const receiptMonthsContainer = document.getElementById("receipt-months-container");
    
    // 1. Populate Ledger Month Filter (hidden select state)
    if (filterSelect) {
        const prevVal = filterSelect.value;
        filterSelect.innerHTML = `<option value="all">كل الشهور</option>`;
        state.monthsList.forEach(m => {
            const opt = document.createElement("option");
            opt.value = m.key;
            opt.textContent = m.key;
            filterSelect.appendChild(opt);
        });
        if (prevVal && [...filterSelect.options].some(o => o.value === prevVal)) {
            filterSelect.value = prevVal;
        }
    }
    
    // 2. Populate Reports Month Select
    if (reportSelect) {
        const prevVal = reportSelect.value;
        reportSelect.innerHTML = "";
        state.monthsList.forEach(m => {
            const opt = document.createElement("option");
            opt.value = m.key;
            opt.textContent = m.key;
            reportSelect.appendChild(opt);
        });
        if (prevVal && [...reportSelect.options].some(o => o.value === prevVal)) {
            reportSelect.value = prevVal;
        } else if (state.monthsList.length > 0) {
            reportSelect.value = state.monthsList[0].key;
        }
    }
    
    // 3. Populate Receipt Form Checkbox grid
    if (receiptMonthsContainer) {
        receiptMonthsContainer.innerHTML = "";
        state.monthsList.forEach(m => {
            const shortKey = m.key.replace("2026", "26").replace("2027", "27").replace("2028", "28");
            const label = document.createElement("label");
            label.style.display = "flex";
            label.style.alignItems = "center";
            label.style.gap = "6px";
            label.style.fontSize = "0.85rem";
            label.style.cursor = "pointer";
            
            const chk = document.createElement("input");
            chk.type = "checkbox";
            chk.className = "receipt-month-chk";
            chk.value = m.key;
            
            // Re-bind the change listener
            chk.addEventListener("change", () => {
                 const headName = receiptSelectFamily.value;
                 const family = state.families.find(f => f.headName === headName);
                 if (family) {
                      const checkedCount = receiptMonthsContainer.querySelectorAll(".receipt-month-chk:checked").length;
                      const expectedAmt = family.subscription * checkedCount;
                      receiptAmount.value = expectedAmt;
                 }
                 updatePrintPreview();
            });
            
            label.appendChild(chk);
            label.appendChild(document.createTextNode(shortKey));
            receiptMonthsContainer.appendChild(label);
        });
    }

    // 4. Populate and Render Interactive Horizontal Month Tabs
    const tabsContainer = document.getElementById("ledger-month-tabs");
    const curActive = filterSelect ? filterSelect.value : "all";

    if (tabsContainer) {
        tabsContainer.innerHTML = "";

        // A. General "All Months" Tab (Comprehensive)
        const allTab = document.createElement("div");
        allTab.className = "month-tab";
        if (curActive === "all") allTab.classList.add("active");
        allTab.innerHTML = `📋 التقرير الشامل <span style="font-size:0.72rem; opacity:0.8;">(شهور متعددة)</span>`;
        allTab.addEventListener("click", () => {
            if (filterSelect) {
                filterSelect.value = "all";
                filterSelect.dispatchEvent(new Event("change"));
            }
            populateMonthFilters();
        });
        tabsContainer.appendChild(allTab);

        // B. Month-by-month Tabs (Check for 100% collect completeness to style green)
        state.monthsList.forEach(m => {
            const tab = document.createElement("div");
            tab.className = "month-tab";
            if (curActive === m.key) tab.classList.add("active");

            // Calculate paid statistics for this month
            const totalM = state.members.length;
            let paidM = 0;
            state.members.forEach(member => {
                if ((member.payments[m.key] || 0) > 0) {
                    paidM++;
                }
            });

            const collectionPercentage = totalM > 0 ? Math.round((paidM / totalM) * 100) : 0;
            const shortKey = m.key.replace("2026", "26").replace("2027", "27").replace("2028", "28");

            if (collectionPercentage === 100 && totalM > 0) {
                tab.classList.add("completed");
                tab.innerHTML = `✨ ${shortKey} <span style="font-size: 0.72rem; font-weight: 800; background: rgba(255,255,255,0.45); padding: 1.5px 6px; border-radius: 30px;">✓ اكتمل 100%</span>`;
            } else {
                tab.innerHTML = `📅 ${shortKey} <span style="font-size: 0.72rem; font-weight: 600; opacity: 0.85;">(${collectionPercentage}%)</span>`;
            }

            tab.addEventListener("click", () => {
                if (filterSelect) {
                    filterSelect.value = m.key;
                    filterSelect.dispatchEvent(new Event("change"));
                }
                populateMonthFilters();
            });

            tabsContainer.appendChild(tab);
        });
    }

    // 5. Populate and Tweak Multi-Month Checkboxes Panel
    const multiMonthPanel = document.getElementById("multi-month-select-panel");
    const multiMonthContainer = document.getElementById("multi-month-checkboxes");
    
    if (multiMonthPanel) {
        if (curActive === "all") {
            multiMonthPanel.style.display = "block";
            
            if (multiMonthContainer) {
                multiMonthContainer.innerHTML = "";
                
                state.monthsList.forEach(m => {
                    const label = document.createElement("label");
                    label.className = "multi-month-chk-label";
                    
                    const isChecked = state.selectedLedgerMonths.includes(m.key);
                    if (isChecked) label.classList.add("checked");
                    
                    const chk = document.createElement("input");
                    chk.type = "checkbox";
                    chk.value = m.key;
                    chk.checked = isChecked;
                    
                    chk.addEventListener("change", async (e) => {
                        const activeChecked = e.target.checked;
                        const previous = state.selectedLedgerMonths.slice();
                        if (activeChecked) {
                            if (!state.selectedLedgerMonths.includes(m.key)) {
                                state.selectedLedgerMonths.push(m.key);
                            }
                            label.classList.add("checked");
                        } else {
                            state.selectedLedgerMonths = state.selectedLedgerMonths.filter(k => k !== m.key);
                            label.classList.remove("checked");
                        }
                        chk.disabled = true;
                        try {
                            await callApi("setSelectedMonths", { selected: state.selectedLedgerMonths });
                            renderLedgerTable();
                        } catch (err) {
                            state.selectedLedgerMonths = previous;
                            e.target.checked = !activeChecked;
                            label.classList.toggle("checked", previous.includes(m.key));
                            alert("تعذّر حفظ خيارات الأشهر: " + err.message);
                        } finally {
                            chk.disabled = false;
                        }
                    });
                    
                    label.appendChild(chk);
                    
                    // Month statistics details
                    const totalM = state.members.length;
                    let paidM = 0;
                    state.members.forEach(member => {
                        if ((member.payments[m.key] || 0) > 0) {
                            paidM++;
                        }
                    });
                    const collectionPercentage = totalM > 0 ? Math.round((paidM / totalM) * 100) : 0;
                    const shortKey = m.key.replace("2026", "26").replace("2027", "27").replace("2028", "28");
                    
                    const spanText = document.createElement("span");
                    spanText.textContent = `${shortKey} (${collectionPercentage}%)`;
                    label.appendChild(spanText);
                    
                    multiMonthContainer.appendChild(label);
                });
            }
        } else {
            multiMonthPanel.style.display = "none";
        }
    }
}

function populateAddMemberFamilies() {
    const selectFamily = document.getElementById("add-member-select-family");
    if (!selectFamily) return;
    
    selectFamily.innerHTML = `<option value="">-- اختر رب عائلة قائم --</option>`;
    state.families.forEach(f => {
        const opt = document.createElement("option");
        opt.value = f.headName;
        opt.textContent = f.headName;
        selectFamily.appendChild(opt);
    });
}

// Add Member Event
document.getElementById("btn-add-member").addEventListener("click", async () => {
    const nameInput = document.getElementById("add-member-name");
    const selectFamily = document.getElementById("add-member-select-family");
    const newFamilyInput = document.getElementById("add-member-new-family");

    const name = nameInput.value.trim();
    if (!name) {
        alert("الرجاء إدخال اسم العضو الجديد.");
        return;
    }

    let parent = "";
    const selectedHead = selectFamily.value;
    const newHead = newFamilyInput.value.trim();

    if (newHead) {
        parent = newHead;
    } else if (selectedHead) {
        parent = selectedHead;
    } else {
        alert("الرجاء اختيار رب عائلة من القائمة أو كتابة اسم رب عائلة جديد.");
        return;
    }

    const exists = state.members.some(m => m.name.trim().toLowerCase() === name.toLowerCase());
    if (exists) {
        alert("هذا العضو مسجل بالفعل في القائمة.");
        return;
    }

    const btn = document.getElementById("btn-add-member");
    btn.disabled = true;
    try {
        const result = await callApi("addMember", { name, parent });
        state.members = result.members;
        state.families = result.families;

        nameInput.value = "";
        newFamilyInput.value = "";

        populateAddMemberFamilies();
        populateReceiptFamilies();

        renderCharts();
        renderLedgerTable();
        renderFamiliesTable();

        alert(`✅ تم إضافة العضو (${name}) بنجاح وإلحاقه بعائلة (${parent}).`);
    } catch (err) {
        alert("تعذّرت إضافة العضو: " + err.message);
    } finally {
        btn.disabled = false;
    }
});

// Add Month Event
document.getElementById("btn-add-month").addEventListener("click", async () => {
    const monthNameInput = document.getElementById("add-month-name");
    const monthIdSelect = document.getElementById("add-month-id");

    const monthKey = monthNameInput.value.trim();
    if (!monthKey) {
        alert("الرجاء إدخال اسم الشهر والسنة (مثال: يناير 2027).");
        return;
    }

    const monthId = parseInt(monthIdSelect.value);

    const exists = state.monthsList.some(mon => mon.key.trim().toLowerCase() === monthKey.toLowerCase());
    if (exists) {
        alert("هذا الشهر مضاف بالفعل في القائمة.");
        return;
    }

    const btn = document.getElementById("btn-add-month");
    btn.disabled = true;
    try {
        const result = await callApi("addMonth", { key: monthKey, id: monthId });
        state.monthsList = result.months.map(m => ({ key: m.key, id: m.id }));
        state.selectedLedgerMonths = result.months.filter(m => m.selected).map(m => m.key);
        state.members = result.members;

        monthNameInput.value = "";

        populateMonthFilters();
        renderCharts();
        renderLedgerTable();
        renderFamiliesTable();

        alert(`✅ تم إضافة الشهر (${monthKey}) بنجاح إلى جدول الاشتراكات والتحصيل.`);
    } catch (err) {
        alert("تعذّرت إضافة الشهر: " + err.message);
    } finally {
        btn.disabled = false;
    }
});

// =============================================================
// EXPENSES / EXPENDITURE MANAGEMENT MODULE
// =============================================================

// Helper: read file, compress image using HTML5 Canvas, and resolve as base64 object
function readAttachmentFile(file) {
    return new Promise((resolve, reject) => {
        if (!file) {
            resolve(null);
            return;
        }

        const reader = new FileReader();

        // 1. If image, scale dynamically to keep storage footprint small (~100kb average)
        if (file.type.startsWith('image/')) {
            reader.onload = function(e) {
                const img = new Image();
                img.onload = function() {
                    const canvas = document.createElement('canvas');
                    let width = img.width;
                    let height = img.height;

                    const MAX_WIDTH = 1000;
                    const MAX_HEIGHT = 1000;

                    if (width > height) {
                        if (width > MAX_WIDTH) {
                            height *= MAX_WIDTH / width;
                            width = MAX_WIDTH;
                        }
                    } else {
                        if (height > MAX_HEIGHT) {
                            width *= MAX_HEIGHT / height;
                            height = MAX_HEIGHT;
                        }
                    }

                    canvas.width = width;
                    canvas.height = height;
                    const ctx = canvas.getContext('2d');
                    ctx.drawImage(img, 0, 0, width, height);

                    // JPEG quality 0.70 compression
                    const compressedBase64 = canvas.toDataURL('image/jpeg', 0.70);
                    resolve({
                        name: file.name,
                        type: 'image/jpeg',
                        size: Math.round((compressedBase64.length * 3) / 4),
                        data: compressedBase64
                    });
                };
                img.onerror = () => reject(new Error("تعذر قراءة ملف الصورة."));
                img.src = e.target.result;
            };
            reader.onerror = () => reject(new Error("خطأ أثناء قراءة ملف الصورة."));
            reader.readAsDataURL(file);
        } else {
            // 2. Documents (PDFs, Excel, Word etc.) read into full base64 string
            reader.onload = function(e) {
                resolve({
                    name: file.name,
                    type: file.type || 'application/octet-stream',
                    size: file.size,
                    data: e.target.result
                });
            };
            reader.onerror = () => reject(new Error("حدث خطأ أثناء قراءة المستند."));
            reader.readAsDataURL(file);
        }
    });
}

// Modal closing listener on outside click
document.addEventListener("DOMContentLoaded", () => {
    const modal = document.getElementById("attachment-modal");
    if (modal) {
        modal.addEventListener("click", (e) => {
            if (e.target === modal) {
                closeAttachmentModal();
            }
        });
    }
});

// View Modal Renderer
window.viewAttachment = function(expenseId) {
    const expense = state.expenses.find(e => e.id === expenseId);
    if (!expense || !expense.attachment) {
        alert("المرفق غير متوفر أو تم حذفه.");
        return;
    }

    const modal = document.getElementById("attachment-modal");
    const title = document.getElementById("attachment-modal-title");
    const body = document.getElementById("attachment-modal-body");
    const downloadBtn = document.getElementById("btn-download-attachment");

    if (!modal || !title || !body || !downloadBtn) return;

    title.textContent = `عرض المرفق: ${expense.attachment.name}`;
    body.innerHTML = "";

    const type = expense.attachment.type;
    const url = expense.attachment.url;

    if (type.startsWith("image/")) {
        const image = document.createElement("img");
        image.src = url;
        image.alt = expense.attachment.name;
        image.style.maxWidth = "100%";
        image.style.maxHeight = "400px";
        image.style.borderRadius = "8px";
        image.style.boxShadow = "var(--shadow)";
        body.appendChild(image);
    } else if (type === "application/pdf") {
        const embed = document.createElement("object");
        embed.data = url;
        embed.type = "application/pdf";
        embed.style.width = "100%";
        embed.style.height = "380px";

        const fallbackText = document.createElement("div");
        fallbackText.style.textAlign = "center";
        fallbackText.style.padding = "20px";
        fallbackText.innerHTML = `
            <p style="margin-bottom:12px; color:var(--text-muted);">لا تدعم المعاينة الفورية لـ PDF في هذا المتصفح.</p>
            <a href="${url}" target="_blank" rel="noopener" class="btn btn-primary" style="font-size:0.85rem;">📥 فتح ملف الـ PDF في تبويب جديد</a>
        `;
        embed.appendChild(fallbackText);
        body.appendChild(embed);
    } else {
        const infoDiv = document.createElement("div");
        infoDiv.style.textAlign = "center";
        infoDiv.style.padding = "30px 20px";
        infoDiv.innerHTML = `
            <div style="font-size:3rem; margin-bottom:10px;">📄</div>
            <h4 style="font-size:1.05rem; color:var(--text-main); margin-bottom:6px;">مستند خارجي: ${expense.attachment.name}</h4>
            <p style="font-size:0.82rem; color:var(--text-muted); margin-bottom:14px;">لا تتوفر معاينة مباشرة لهذا النوع من الملفات (${type}).</p>
            <p style="font-weight:600; color:var(--primary); font-size:0.85rem;">انقر فوق زر التحميل أدناه لفتح المستند.</p>
        `;
        body.appendChild(infoDiv);
    }

    // Download action hook — url is a cross-origin Drive link, so a
    // download-attribute anchor won't force a save-as; open it instead.
    downloadBtn.onclick = function() {
        window.open(url, "_blank", "noopener");
    };

    modal.style.display = "flex";
};

window.closeAttachmentModal = function() {
    const modal = document.getElementById("attachment-modal");
    if (modal) modal.style.display = "none";
};

function renderExpensesTable() {
    const tbody = document.getElementById("expense-rows");
    const badge = document.getElementById("exp-total-badge");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (state.expenses.length === 0) {
        tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; color: var(--text-muted);">لا توجد أوامر صرف مسجّلة بعد.</td></tr>`;
        if (badge) badge.textContent = "الإجمالي: 0 شيكل";
        return;
    }

    let totalExp = 0;
    // Show newest first
    [...state.expenses].reverse().forEach((e, idx) => {
        totalExp += e.amount;
        const displayIdx = state.expenses.length - idx;
        const tr = document.createElement("tr");

        // Format Attachment preview trigger
        const attachmentHtml = e.attachment 
            ? `<button onclick="viewAttachment('${e.id}')" style="background:rgba(59,130,246,0.12); color:#3b82f6; border:1px solid rgba(59,130,246,0.3); padding:3px 8px; border-radius:6px; cursor:pointer; font-size:0.75rem; display:inline-flex; align-items:center; gap:4px; font-family:'Cairo', sans-serif;">📎 عرض (${e.attachment.name.length > 12 ? e.attachment.name.substring(0,8)+'...'+e.attachment.name.split('.').pop() : e.attachment.name})</button>`
            : `<span style="color:var(--text-muted); font-size:0.8rem;">لا يوجد</span>`;

        tr.innerHTML = `
            <td>${displayIdx}</td>
            <td>${e.date}</td>
            <td><span style="background:rgba(245,158,11,0.12); color:#f59e0b; padding:2px 8px; border-radius:10px; font-size:0.78rem;">${e.category}</span></td>
            <td style="max-width:200px;">${e.reason}</td>
            <td style="font-weight:700; color:#f59e0b;">${e.amount} شيكل</td>
            <td>${attachmentHtml}</td>
            <td>${e.authorized}</td>
            <td>
                <button onclick="deleteExpense('${e.id}')" style="background:rgba(239,68,68,0.12); color:#ef4444; border:1px solid rgba(239,68,68,0.3); padding:3px 10px; border-radius:6px; cursor:pointer; font-size:0.78rem; font-family:'Cairo', sans-serif;">🗑 حذف</button>
            </td>
        `;
        tbody.appendChild(tr);
    });

    const total = state.expenses.reduce((s, e) => s + e.amount, 0);
    if (badge) badge.textContent = `الإجمالي: ${total} شيكل`;
}

window.deleteExpense = function(id) {
    showConfirm("هل أنت متأكد من حذف أمر الصرف هذا؟ سيتم استعادة المبلغ للرصيد الصافي.", async () => {
        try {
            state.expenses = await callApi("deleteExpense", { id });
            renderExpensesTable();
            updateIndicators();
        } catch (err) {
            alert("تعذّر حذف أمر الصرف: " + err.message);
        }
    });
};

// Add Expense button handler
document.getElementById("btn-add-expense").addEventListener("click", async () => {
    const dateVal = document.getElementById("exp-date").value;
    const amountVal = parseFloat(document.getElementById("exp-amount").value);
    const reasonVal = document.getElementById("exp-reason").value.trim();
    const categoryVal = document.getElementById("exp-category").value;
    const authorizedVal = document.getElementById("exp-authorized").value.trim();

    const fileInput = document.getElementById("exp-file");
    const file = fileInput && fileInput.files ? fileInput.files[0] : null;

    if (!dateVal) {
        alert("الرجاء تحديد تاريخ الصرف.");
        return;
    }
    if (!amountVal || amountVal <= 0) {
        alert("الرجاء إدخال مبلغ صرف صحيح أكبر من صفر.");
        return;
    }
    if (!reasonVal) {
        alert("الرجاء كتابة سبب الصرف ووصف المهمة.");
        return;
    }

    if (file && file.size > 1.5 * 1024 * 1024) {
        alert("⚠️ حجم الملف المرفق كبير جداً! يرجى اختيار ملفات أقل من 1.5 ميجابايت للحفاظ على كفاءة التخزين المحلي.");
        return;
    }

    // Warn if amount exceeds net balance
    const totalCollected = state.members.reduce((s, m) => s + m.sum, 0);
    const totalExpenses = state.expenses.reduce((s, e) => s + e.amount, 0);
    const netBalance = totalCollected - totalExpenses;
    if (amountVal > netBalance) {
        const proceed = confirm(`⚠️ تحذير: المبلغ المطلوب صرفه (${amountVal} شيكل) يتجاوز الرصيد الصافي الحالي (${netBalance} شيكل).\n\nهل تريد المتابعة رغم ذلك؟`);
        if (!proceed) return;
    }

    let attachmentObj = null;
    if (file) {
        try {
            // Show loading on the button
            const btn = document.getElementById("btn-add-expense");
            const originalText = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = `<span style="display:inline-block; width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; margin-left:8px; vertical-align:middle;"></span> جاري ضغط ومعالجة الملف...`;

            attachmentObj = await readAttachmentFile(file);

            btn.disabled = false;
            btn.innerHTML = originalText;
        } catch (err) {
            alert("حدث خطأ أثناء معالجة الملف: " + err.message);
            return;
        }
    }

    const btn = document.getElementById("btn-add-expense");
    const originalText = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `<span style="display:inline-block; width:14px; height:14px; border:2px solid #fff; border-top-color:transparent; border-radius:50%; animation:spin 0.8s linear infinite; margin-left:8px; vertical-align:middle;"></span> جاري الحفظ على جوجل شيت...`;

    try {
        state.expenses = await callApi("addExpense", {
            date: dateVal,
            amount: amountVal,
            reason: reasonVal,
            category: categoryVal,
            authorized: authorizedVal || (state.currentUser ? state.currentUser.name : "اللجنة المالية"),
            attachment: attachmentObj
        });

        document.getElementById("exp-amount").value = "";
        document.getElementById("exp-reason").value = "";
        if (fileInput) fileInput.value = "";

        renderExpensesTable();
        updateIndicators();

        alert(`✅ تم تسجيل أمر الصرف بنجاح!\nالمبلغ: ${amountVal} شيكل\nالسبب: ${reasonVal}\nالرصيد الصافي الجديد: ${netBalance - amountVal} شيكل`);
    } catch (err) {
        alert("تعذّر تسجيل أمر الصرف: " + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalText;
    }
});

// Export expenses to Excel
document.getElementById("btn-export-expenses").addEventListener("click", () => {
    if (state.expenses.length === 0) {
        alert("لا توجد أوامر صرف للتصدير.");
        return;
    }
    const rows = state.expenses.map((e, i) => ({
        "#": i + 1,
        "التاريخ": e.date,
        "التصنيف": e.category,
        "سبب الصرف": e.reason,
        "المبلغ (شيكل)": e.amount,
        "المرفق": e.attachment ? `${e.attachment.name} (مرفق)` : "لا يوجد",
        "معتمد من": e.authorized
    }));
    // Append totals row
    const total = state.expenses.reduce((s, e) => s + e.amount, 0);
    rows.push({ "#": "", "التاريخ": "", "التصنيف": "الإجمالي", "سبب الصرف": "", "المبلغ (شيكل)": total, "المرفق": "", "معتمد من": "" });

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, "سجل الصرف");
    XLSX.writeFile(wb, "سجل_الصرف_آل_اطفيحة.xlsx");
});

// Initialize expense and user management tab defaults
document.addEventListener("DOMContentLoaded", () => {
    // Set today as default expense date
    const expDateInput = document.getElementById("exp-date");
    if (expDateInput) expDateInput.value = new Date().toISOString().split("T")[0];

    // Render initial (empty) table
    renderExpensesTable();

    // Initialize signature pad Canvas events
    initSignaturePad();
});

// =============================================================
// USER MANAGEMENT & REGISTRATION REQUESTS MODULE
// =============================================================

function updatePendingBadge() {
    const badge = document.getElementById("pending-badge");
    const label = document.getElementById("pending-count-label");
    const count = state.pendingUsers.length;

    if (badge) {
        badge.textContent = count;
        badge.style.display = count > 0 ? "flex" : "none";
    }
    if (label) {
        label.textContent = `${count} طلب${count === 1 ? '' : 'ات'}`;
    }
}

function renderPendingUsers() {
    const tbody = document.getElementById("pending-user-rows");
    if (!tbody) return;

    tbody.innerHTML = "";

    if (state.pendingUsers.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" style="text-align: center; color: var(--text-muted);">لا توجد طلبات تسجيل معلّقة.</td></tr>`;
        return;
    }

    state.pendingUsers.forEach((req, idx) => {
        const tr = document.createElement("tr");
        const name = `${req.firstName} ${req.lastName}`;
        const requestedDate = req.requestedAt ? req.requestedAt.split("T")[0] : "-";
        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td><strong>${name}</strong></td>
            <td>${req.email}</td>
            <td>مشرف مالي</td>
            <td>${requestedDate}</td>
            <td>
                <div style="display: flex; gap: 8px;">
                    <button class="btn btn-primary btn-sm" onclick="approveUser('${req.id}')">✔️ موافقة</button>
                    <button class="btn btn-danger btn-sm" onclick="rejectUser('${req.id}')">❌ رفض</button>
                </div>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

window.approveUser = async function(id) {
    const user = state.pendingUsers.find(u => u.id === id);
    if (!user) return;
    const name = `${user.firstName} ${user.lastName}`;

    try {
        const result = await callApi("approveUser", { id });
        state.users = result.users;
        state.pendingUsers = result.pendingUsers;

        alert(`✔️ تم تفعيل حساب المشرف: ${name} بنجاح!`);
        updatePendingBadge();
        renderPendingUsers();
        renderActiveUsers();
    } catch (err) {
        alert("تعذّرت الموافقة على الطلب: " + err.message);
    }
};

window.rejectUser = function(id) {
    const user = state.pendingUsers.find(u => u.id === id);
    if (!user) return;
    const name = `${user.firstName} ${user.lastName}`;

    if (!confirm(`هل أنت متأكد من رفض طلب تسجيل المشرف: ${name}؟`)) return;

    callApi("rejectUser", { id }).then(pendingUsers => {
        state.pendingUsers = pendingUsers;
        updatePendingBadge();
        renderPendingUsers();
    }).catch(err => {
        alert("تعذّر رفض الطلب: " + err.message);
    });
};

function renderActiveUsers() {
    const tbody = document.getElementById("active-user-rows");
    if (!tbody) return;

    tbody.innerHTML = "";

    state.users.forEach((user, idx) => {
        const tr = document.createElement("tr");
        let actionsHtml;
        if (user.isMaster) {
            actionsHtml = '<span style="color: var(--text-muted); font-size: 0.82rem;">وصول كامل للوحة</span>';
        } else if (user.isBuiltIn) {
            actionsHtml = `
                <button class="btn btn-secondary btn-sm" onclick="toggleUserStatus('${user.key}')">
                    ${user.isActive ? '❌ تعطيل' : '✔️ تفعيل'}
                </button>
            `;
        } else {
            actionsHtml = `
                <div style="display: flex; gap: 8px;">
                    <button class="btn btn-secondary btn-sm" onclick="toggleUserStatus('${user.key}')">
                        ${user.isActive ? '❌ تعطيل' : '✔️ تفعيل'}
                    </button>
                    <button class="btn btn-danger btn-sm" onclick="deleteAppUser('${user.key}')">🗑️ حذف</button>
                </div>
            `;
        }

        tr.innerHTML = `
            <td>${idx + 1}</td>
            <td><strong>${user.name}</strong></td>
            <td>${user.email}</td>
            <td>${user.role}</td>
            <td><span class="badge ${user.isActive ? 'badge-success' : 'badge-danger'}">${user.isBuiltIn ? (user.isActive ? 'أساسي' : 'معطل') : (user.isActive ? 'نشط' : 'معطل')}</span></td>
            <td>${actionsHtml}</td>
        `;
        tbody.appendChild(tr);
    });

    const activeCountLabel = document.getElementById("active-users-count");
    if (activeCountLabel) {
        activeCountLabel.textContent = `${state.users.length} مستخدم`;
    }
}

window.toggleUserStatus = function(key) {
    const user = state.users.find(u => u.key === key);
    if (!user) return;
    const nextIsActive = !user.isActive;

    showConfirm(`هل أنت متأكد من ${nextIsActive ? 'تفعيل' : 'تعطيل'} حساب المشرف المالي: ${user.name}؟`, async () => {
        try {
            state.users = await callApi("setUserActive", { key, active: nextIsActive });
            renderActiveUsers();
        } catch (err) {
            alert("تعذّر تغيير حالة الحساب: " + err.message);
        }
    });
};

window.deleteAppUser = function(key) {
    const user = state.users.find(u => u.key === key);
    if (!user) return;

    showConfirm(`⚠️ تحذير: هل أنت متأكد من الحذف النهائي لحساب المشرف: ${user.name}؟ لن يتمكن من تسجيل الدخول بعد الآن.`, async () => {
        try {
            state.users = await callApi("deleteUser", { key });
            alert("🗑️ تم حذف حساب المشرف بنجاح.");
            renderActiveUsers();
        } catch (err) {
            alert("تعذّر حذف الحساب: " + err.message);
        }
    });
};

// Direct Add Supervisor Event
const btnAddSupervisor = document.getElementById("btn-add-supervisor");
if (btnAddSupervisor) {
    btnAddSupervisor.addEventListener("click", async () => {
        const name = document.getElementById("su-name").value.trim();
        const email = document.getElementById("su-email").value.trim().toLowerCase();
        const role = document.getElementById("su-role").value.trim();
        const pass = document.getElementById("su-pass").value;

        if (!name || !email || !role || !pass) {
            alert("الرجاء تعبئة كافة حقول المشرف الجديد.");
            return;
        }

        if (pass.length < 6) {
            alert("يجب أن تكون كلمة المرور للمشرف 6 أحرف على الأقل.");
            return;
        }

        if (findUserInState_(email)) {
            alert("هذا البريد الإلكتروني مسجل بالفعل لمشرف آخر.");
            return;
        }

        try {
            state.users = await callApi("addSupervisor", { name, email, role, pass });

            alert(`✅ تم إضافة المشرف المالي: ${name} وتفعيل حسابه فوراً وبنجاح!`);

            document.getElementById("su-name").value = "";
            document.getElementById("su-email").value = "";
            document.getElementById("su-role").value = "";
            document.getElementById("su-pass").value = "";

            renderActiveUsers();
        } catch (err) {
            alert("تعذّرت إضافة المشرف: " + err.message);
        }
    });
}

// -------------------------------------------------------------
// Manual Budget Recalculation & Sync Handler
// -------------------------------------------------------------
function doRecalculateAll() {
    // 1. Recalculate each member's total payments (sum)
    state.members.forEach(member => {
        let sum = 0;
        state.monthsList.forEach(m => {
            sum += (member.payments[m.key] || 0);
        });
        member.sum = sum;
    });

    // 2. Rebuild/recalculate families summaries to match member sums with Arabic normalization
    state.families.forEach(f => {
        const normHead = normalizeArabicName(f.headName);
        let familySum = 0;
        state.members.forEach(m => {
            if (normalizeArabicName(m.parent) === normHead) {
                familySum += m.sum;
            }
        });
        f.totalPaid = familySum;
    });

    // 3. Save to localStorage to persist
    saveToLocalMemory();

    // 4. Update core UI displays
    updateIndicators();
    renderCharts();

    // 5. Update active view tables
    const activeLink = document.querySelector(".nav-link.active");
    if (activeLink) {
        const activeTab = activeLink.getAttribute("data-tab");
        if (activeTab === "ledger-tab") {
            renderLedgerTable();
        } else if (activeTab === "families-tab") {
            renderFamiliesTable();
        } else if (activeTab === "receipt-tab") {
            populateReceiptFamilies();
        } else if (activeTab === "reports-tab") {
            renderMonthlyReport();
        } else if (activeTab === "expenses-tab") {
            renderExpensesTable();
        }
    }
}

const btnRecalculateAll = document.getElementById("btn-recalculate-all");
if (btnRecalculateAll) {
    btnRecalculateAll.addEventListener("click", () => {
        doRecalculateAll();
        alert("✅ تم إعادة عملية فحص ومطابقة جميع الحسابات وتحديث أرقام الميزانية والمصروفات بنجاح!");
    });
}

// -------------------------------------------------------------
// MEMBER DELETION MODULE
// -------------------------------------------------------------
window.deleteMember = function(id, name) {
    showConfirm(`هل أنت متأكد من حذف العضو (${name}) نهائياً من الصندوق؟\nسيتأثر مجموع الأسرة والتقارير المالية بهذا التغيير.`, async () => {
        try {
            const result = await callApi("deleteMember", { id });
            state.members = result.members;
            state.families = result.families;

            renderLedgerTable();
            renderFamiliesTable();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();

            alert(`🗑️ تم حذف العضو (${name}) بنجاح.`);
        } catch (err) {
            alert("تعذّر حذف العضو: " + err.message);
        }
    });
};

// -------------------------------------------------------------
// ELECTRONIC SIGNATURE PAD LOGIC
// -------------------------------------------------------------
let currentSigIndex = null;
let isDrawing = false;
let sigCanvas = null;
let sigCtx = null;

function initSignaturePad() {
    sigCanvas = document.getElementById("signature-pad");
    if (!sigCanvas) return;
    sigCtx = sigCanvas.getContext("2d");
    
    // Draw styles
    sigCtx.strokeStyle = "#0f766e"; // Teal ink
    sigCtx.lineWidth = 3;
    sigCtx.lineCap = "round";
    sigCtx.lineJoin = "round";
    
    // Mouse Events
    sigCanvas.addEventListener("mousedown", startDrawing);
    sigCanvas.addEventListener("mousemove", draw);
    sigCanvas.addEventListener("mouseup", stopDrawing);
    sigCanvas.addEventListener("mouseout", stopDrawing);
    
    // Touch Events
    sigCanvas.addEventListener("touchstart", (e) => {
        const touch = e.touches[0];
        const rect = sigCanvas.getBoundingClientRect();
        const scaleX = sigCanvas.width / rect.width;
        const scaleY = sigCanvas.height / rect.height;
        
        isDrawing = true;
        sigCtx.beginPath();
        sigCtx.moveTo((touch.clientX - rect.left) * scaleX, (touch.clientY - rect.top) * scaleY);
        e.preventDefault();
    });
    sigCanvas.addEventListener("touchmove", (e) => {
        if (!isDrawing) return;
        const touch = e.touches[0];
        const rect = sigCanvas.getBoundingClientRect();
        const scaleX = sigCanvas.width / rect.width;
        const scaleY = sigCanvas.height / rect.height;
        
        sigCtx.lineTo((touch.clientX - rect.left) * scaleX, (touch.clientY - rect.top) * scaleY);
        sigCtx.stroke();
        e.preventDefault();
    });
    sigCanvas.addEventListener("touchend", (e) => {
        isDrawing = false;
        e.preventDefault();
    });
    
    document.getElementById("btn-draw-clear").addEventListener("click", () => {
        sigCtx.clearRect(0, 0, sigCanvas.width, sigCanvas.height);
    });
    
    document.getElementById("btn-save-signature").addEventListener("click", saveSignature);
}

function startDrawing(e) {
    isDrawing = true;
    const rect = sigCanvas.getBoundingClientRect();
    const scaleX = sigCanvas.width / rect.width;
    const scaleY = sigCanvas.height / rect.height;
    
    sigCtx.beginPath();
    sigCtx.moveTo((e.clientX - rect.left) * scaleX, (e.clientY - rect.top) * scaleY);
}

function draw(e) {
    if (!isDrawing) return;
    const rect = sigCanvas.getBoundingClientRect();
    const scaleX = sigCanvas.width / rect.width;
    const scaleY = sigCanvas.height / rect.height;
    
    sigCtx.lineTo((e.clientX - rect.left) * scaleX, (e.clientY - rect.top) * scaleY);
    sigCtx.stroke();
}

function stopDrawing() {
    isDrawing = false;
}

window.openSignatureModal = function(index) {
    currentSigIndex = index;
    const modal = document.getElementById("signature-modal");
    if (modal) {
        modal.style.display = "flex";
        if (sigCanvas) {
            sigCtx.clearRect(0, 0, sigCanvas.width, sigCanvas.height);
        }
    }
};

window.closeSignatureModal = function() {
    const modal = document.getElementById("signature-modal");
    if (modal) {
        modal.style.display = "none";
    }
};

async function saveSignature() {
    if (!currentSigIndex) return;

    const blank = document.createElement("canvas");
    blank.width = sigCanvas.width;
    blank.height = sigCanvas.height;
    if (sigCanvas.toDataURL() === blank.toDataURL()) {
        alert("الرجاء رسم التوقيع أولاً قبل الحفظ.");
        return;
    }

    const dataUrl = sigCanvas.toDataURL();
    const monthVal = document.getElementById("report-select-month").value;

    try {
        await callApi("saveSignature", { month: monthVal, slotIndex: currentSigIndex, data: dataUrl });
        await window.updateSignaturesDisplay();
        closeSignatureModal();
    } catch (err) {
        alert("تعذّر حفظ التوقيع: " + err.message);
    }
}

window.clearSignature = function(index) {
    showConfirm("هل أنت متأكد من إزالة هذا التوقيع الرقمي؟", async () => {
        const monthVal = document.getElementById("report-select-month").value;
        try {
            await callApi("clearSignature", { month: monthVal, slotIndex: index });
            await window.updateSignaturesDisplay();
        } catch (err) {
            alert("تعذّر إزالة التوقيع: " + err.message);
        }
    });
};

window.updateSignaturesDisplay = async function() {
    const reportSelect = document.getElementById("report-select-month");
    if (!reportSelect) return;
    const monthVal = reportSelect.value;
    let sigs = {};
    try {
        sigs = await callApi("getSignatures", { month: monthVal });
    } catch (err) {
        console.error("Could not load signatures:", err);
    }

    for (let i = 1; i <= 5; i++) {
        const area = document.getElementById(`sig-area-${i}`);
        if (!area) continue;
        
        if (sigs[i]) {
            area.innerHTML = `
                <div style="display:flex; align-items:center; justify-content:center; gap:4px; height:45px;">
                    <img src="${sigs[i]}" style="max-height: 42px; max-width: 100%; object-fit: contain; margin-top: 4px;" alt="توقيع">
                    <button class="btn-sig-clear no-print" onclick="clearSignature('${i}')" title="إزالة التوقيع">❌</button>
                </div>
            `;
        } else {
            area.innerHTML = `
                <div class="line"></div>
                <button class="btn-sig no-print" onclick="openSignatureModal('${i}')">✍️ توقيع</button>
            `;
        }
    }
};

// -------------------------------------------------------------
// CUSTOM NON-BLOCKING CONFIRM DIALOG LOGIC
// -------------------------------------------------------------
window.showConfirm = function(message, callback) {
    const modal = document.getElementById("confirm-modal");
    const textEl = document.getElementById("confirm-modal-text");
    const btnYes = document.getElementById("btn-confirm-yes");
    const btnNo = document.getElementById("btn-confirm-no");
    
    if (!modal || !textEl || !btnYes || !btnNo) {
        // Fallback to native confirm if elements are absent in the DOM
        if (confirm(message)) {
            callback();
        }
        return;
    }
    
    textEl.textContent = message;
    modal.style.display = "flex";
    
    // Clear and re-bind event listeners to avoid execution accumulation
    const newBtnYes = btnYes.cloneNode(true);
    const newBtnNo = btnNo.cloneNode(true);
    btnYes.parentNode.replaceChild(newBtnYes, btnYes);
    btnNo.parentNode.replaceChild(newBtnNo, btnNo);
    
    newBtnYes.addEventListener("click", () => {
        modal.style.display = "none";
        callback();
    });
    
    newBtnNo.addEventListener("click", () => {
        modal.style.display = "none";
    });
};

// Smooth horizontal scroll helper for collection ledger
window.scrollLedgerTable = function(amount) {
    const container = document.querySelector("#ledger-tab .table-responsive");
    if (container) {
        container.scrollBy({ left: amount, behavior: "smooth" });
    }
};

// -------------------------------------------------------------
// MEMBER GENERAL EDIT MODULE
// -------------------------------------------------------------
window.openEditMemberModal = function(id) {
    const member = state.members.find(m => String(m.id) === String(id));
    if (!member) {
        console.error("Member not found for editing. ID:", id);
        return;
    }

    document.getElementById("edit-member-id").value = member.id;
    document.getElementById("edit-member-name").value = member.name;
    document.getElementById("edit-member-new-family").value = "";

    // Populate family dropdown in edit modal
    const selectFam = document.getElementById("edit-member-select-family");
    if (selectFam) {
        selectFam.innerHTML = `<option value="">-- اختر رب عائلة قائم --</option>`;
        state.families.forEach(f => {
            const opt = document.createElement("option");
            opt.value = f.headName;
            opt.textContent = f.headName;
            selectFam.appendChild(opt);
        });
        selectFam.value = member.parent || "";
    }

    const modal = document.getElementById("edit-member-modal");
    if (modal) modal.style.display = "flex";
};

window.closeEditMemberModal = function() {
    const modal = document.getElementById("edit-member-modal");
    if (modal) modal.style.display = "none";
};

// Save edited member info
const btnSaveMemberEdit = document.getElementById("btn-save-member-edit");
if (btnSaveMemberEdit) {
    btnSaveMemberEdit.addEventListener("click", async () => {
        const id = document.getElementById("edit-member-id").value;
        const newName = document.getElementById("edit-member-name").value.trim();
        const selectedFam = document.getElementById("edit-member-select-family").value;
        const newFamInput = document.getElementById("edit-member-new-family").value.trim();

        if (!newName) {
            alert("الرجاء إدخال اسم العضو.");
            return;
        }

        let parent = "";
        if (newFamInput) {
            parent = newFamInput;
        } else if (selectedFam) {
            parent = selectedFam;
        } else {
            alert("الرجاء اختيار رب عائلة أو كتابة اسم رب عائلة جديد.");
            return;
        }

        btnSaveMemberEdit.disabled = true;
        try {
            const result = await callApi("updateMember", { id, name: newName, parent });
            state.members = result.members;
            state.families = result.families;

            closeEditMemberModal();
            renderLedgerTable();
            renderFamiliesTable();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();
        } catch (err) {
            alert("تعذّر حفظ تعديل العضو: " + err.message);
        } finally {
            btnSaveMemberEdit.disabled = false;
        }
    });
}

// -------------------------------------------------------------
// MONTH EDIT & DELETE MODULES
// -------------------------------------------------------------
window.openEditMonthModal = function() {
    const monthSelect = document.getElementById("ledger-filter-month");
    if (!monthSelect) return;
    const oldKey = monthSelect.value;
    if (oldKey === "all") return;

    document.getElementById("edit-month-old-key").value = oldKey;
    document.getElementById("edit-month-display-old").value = oldKey;
    document.getElementById("edit-month-new-key").value = oldKey;

    const modal = document.getElementById("edit-month-modal");
    if (modal) modal.style.display = "flex";
};

window.closeEditMonthModal = function() {
    const modal = document.getElementById("edit-month-modal");
    if (modal) modal.style.display = "none";
};

const btnSaveMonthEdit = document.getElementById("btn-save-month-edit");
if (btnSaveMonthEdit) {
    btnSaveMonthEdit.addEventListener("click", async () => {
        const oldKey = document.getElementById("edit-month-old-key").value;
        const newKey = document.getElementById("edit-month-new-key").value.trim();

        if (!newKey) {
            alert("الرجاء إدخال الاسم الجديد للشهر.");
            return;
        }
        if (oldKey === newKey) {
            closeEditMonthModal();
            return;
        }

        const exists = state.monthsList.some(m => m.key.toLowerCase() === newKey.toLowerCase());
        if (exists) {
            alert("هذا الاسم مسجل بالفعل لشهر آخر.");
            return;
        }

        btnSaveMonthEdit.disabled = true;
        try {
            const result = await callApi("updateMonth", { oldKey, newKey });
            state.monthsList = result.months.map(m => ({ key: m.key, id: m.id }));
            state.selectedLedgerMonths = result.months.filter(m => m.selected).map(m => m.key);
            state.members = result.members;

            closeEditMonthModal();
            populateMonthFilters();

            const monthSelect = document.getElementById("ledger-filter-month");
            if (monthSelect) monthSelect.value = newKey;

            renderLedgerTable();
            renderFamiliesTable();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();
        } catch (err) {
            alert("تعذّر حفظ تعديل الشهر: " + err.message);
        } finally {
            btnSaveMonthEdit.disabled = false;
        }
    });
}

window.deleteMonth = function() {
    const monthSelect = document.getElementById("ledger-filter-month");
    if (!monthSelect) return;
    const monthKey = monthSelect.value;
    if (monthKey === "all") return;

    showConfirm(`هل أنت متأكد من حذف شهر (${monthKey}) نهائياً من الصندوق؟\nسيتم إزالة كافة اشتراكات هذا الشهر المسجلة لجميع الأعضاء.`, async () => {
        try {
            const result = await callApi("deleteMonth", { key: monthKey });
            state.monthsList = result.months.map(m => ({ key: m.key, id: m.id }));
            state.selectedLedgerMonths = result.months.filter(m => m.selected).map(m => m.key);
            state.members = result.members;

            populateMonthFilters();
            const select = document.getElementById("ledger-filter-month");
            if (select) select.value = "all";

            const ops = document.getElementById("month-ops-controls");
            if (ops) ops.style.display = "none";

            renderLedgerTable();
            renderFamiliesTable();
            populateAddMemberFamilies();
            populateReceiptFamilies();
            renderCharts();
        } catch (err) {
            alert("تعذّر حذف الشهر: " + err.message);
        }
    });
};

const btnEditMonth = document.getElementById("btn-edit-month");
if (btnEditMonth) {
    btnEditMonth.addEventListener("click", openEditMonthModal);
}
const btnDeleteMonth = document.getElementById("btn-delete-month");
if (btnDeleteMonth) {
    btnDeleteMonth.addEventListener("click", deleteMonth);
}

// Table horizontal scrolling action helper
window.scrollLedgerTable = function(amount) {
    const container = document.querySelector(".table-responsive");
    if (container) {
        container.scrollBy({
            left: amount,
            behavior: "smooth"
        });
    }
};
