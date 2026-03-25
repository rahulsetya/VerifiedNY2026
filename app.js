// ============================================================
// CONFIGURATION — Update these values after Azure AD app setup
// ============================================================
const CONFIG = {
    // Replace with your Azure AD Application (client) ID
    clientId: '2a17dd5e-6404-4a1b-abcc-dd1cb41a0c8b',

    // Your Azure AD tenant ID (found in Azure Portal > Azure Active Directory > Overview)
    tenantId: '5c357d2f-916a-4262-8702-1f4e24bdf62a',

    // The SharePoint file item ID (extracted from your URL)
    fileItemId: 'b5dc051b-583e-402c-910a-6e5769f27221',

    // Graph API scopes needed
    scopes: ['Files.Read', 'Files.Read.All', 'Sites.Read.All'],

    // Rows per page in the table
    pageSize: 25,
};

// ============================================================
// MSAL Authentication
// ============================================================
const msalConfig = {
    auth: {
        clientId: CONFIG.clientId,
        authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
        redirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false,
    },
};

let msalInstance;
let currentAccount = null;

// App state
let allData = { headers: [], rows: [] };
let filteredRows = [];
let currentPage = 1;
let sortCol = -1;
let sortAsc = true;
let companiesChart = null;
let rolesChart = null;

// ============================================================
// Initialize
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
    try {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        msalInstance.initialize().then(() => {
            msalInstance.handleRedirectPromise().then(handleResponse).catch(showError);
        });
    } catch (e) {
        console.error('MSAL init error:', e);
    }
});

function handleResponse(response) {
    if (response) {
        currentAccount = response.account;
    } else {
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            currentAccount = accounts[0];
        }
    }

    if (currentAccount) {
        showDashboard();
        refreshData();
    }
}

// ============================================================
// Auth
// ============================================================
async function signIn() {
    try {
        const response = await msalInstance.loginPopup({
            scopes: CONFIG.scopes,
        });
        handleResponse(response);
    } catch (e) {
        console.error('Login error:', e);
        if (e.errorCode !== 'user_cancelled') {
            showError('Sign-in failed: ' + e.message);
        }
    }
}

function signOut() {
    msalInstance.logoutPopup({ account: currentAccount });
    currentAccount = null;
    document.getElementById('dashboard').classList.add('hidden');
    document.getElementById('login-screen').classList.remove('hidden');
}

async function getAccessToken() {
    const request = {
        scopes: CONFIG.scopes,
        account: currentAccount,
    };
    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (e) {
        // Fallback to popup
        const response = await msalInstance.acquireTokenPopup(request);
        return response.accessToken;
    }
}

// ============================================================
// Data Fetching
// ============================================================
async function refreshData() {
    const loading = document.getElementById('loading');
    const refreshIcon = document.getElementById('refresh-icon');

    loading.classList.remove('hidden');
    refreshIcon.classList.add('refresh-spin');
    hideError();

    try {
        const token = await getAccessToken();

        // Fetch the used range of the first worksheet
        const worksheetsUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${CONFIG.fileItemId}/workbook/worksheets`;
        const wsResponse = await graphFetch(worksheetsUrl, token);
        const sheetName = wsResponse.value[0].name;

        const dataUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${CONFIG.fileItemId}/workbook/worksheets('${encodeURIComponent(sheetName)}')/usedRange`;
        const rangeData = await graphFetch(dataUrl, token);

        processData(rangeData);
        renderDashboard();

        document.getElementById('last-updated').textContent = `Last updated: ${new Date().toLocaleTimeString()}`;
    } catch (e) {
        console.error('Data fetch error:', e);
        showError('Failed to fetch data: ' + e.message);
    } finally {
        loading.classList.add('hidden');
        refreshIcon.classList.remove('refresh-spin');
    }
}

async function graphFetch(url, token) {
    const response = await fetch(url, {
        headers: { Authorization: `Bearer ${token}` },
    });
    if (!response.ok) {
        const err = await response.json().catch(() => ({}));
        throw new Error(err.error?.message || `HTTP ${response.status}`);
    }
    return response.json();
}

// ============================================================
// Data Processing
// ============================================================
function processData(rangeData) {
    const values = rangeData.values || [];
    if (values.length === 0) {
        allData = { headers: [], rows: [] };
        return;
    }

    allData.headers = values[0].map((h) => (h ? String(h).trim() : `Column ${values[0].indexOf(h) + 1}`));
    allData.rows = values.slice(1).filter((row) => row.some((cell) => cell !== null && cell !== ''));

    filteredRows = [...allData.rows];
    currentPage = 1;
    sortCol = -1;
}

// ============================================================
// Rendering
// ============================================================
function showDashboard() {
    document.getElementById('login-screen').classList.add('hidden');
    document.getElementById('dashboard').classList.remove('hidden');
    document.getElementById('user-name').textContent = currentAccount.name || currentAccount.username;
}

function renderDashboard() {
    renderStats();
    renderCharts();
    renderTable();
}

function renderStats() {
    const { headers, rows } = allData;

    document.getElementById('stat-total').textContent = rows.length.toLocaleString();

    // Find company column
    const companyCol = findColumn(headers, ['company', 'organization', 'org', 'account', 'firm']);
    if (companyCol >= 0) {
        const unique = new Set(rows.map((r) => r[companyCol]).filter(Boolean));
        document.getElementById('stat-companies').textContent = unique.size.toLocaleString();
    } else {
        document.getElementById('stat-companies').textContent = '—';
    }

    // Find email column
    const emailCol = findColumn(headers, ['email', 'e-mail', 'email address']);
    if (emailCol >= 0) {
        const withEmail = rows.filter((r) => r[emailCol] && String(r[emailCol]).includes('@')).length;
        document.getElementById('stat-emails').textContent = withEmail.toLocaleString();
    } else {
        document.getElementById('stat-emails').textContent = '—';
    }

    // Find role/title column
    const roleCol = findColumn(headers, ['role', 'title', 'job title', 'position', 'job role']);
    if (roleCol >= 0) {
        const unique = new Set(rows.map((r) => r[roleCol]).filter(Boolean));
        document.getElementById('stat-roles').textContent = unique.size.toLocaleString();
    } else {
        document.getElementById('stat-roles').textContent = '—';
    }
}

function renderCharts() {
    const { headers, rows } = allData;
    const colors = ['#4f46e5', '#7c3aed', '#2563eb', '#0891b2', '#059669', '#d97706', '#dc2626', '#db2777', '#6366f1', '#14b8a6'];

    // Companies chart
    const companyCol = findColumn(headers, ['company', 'organization', 'org', 'account', 'firm']);
    const companiesCanvas = document.getElementById('chart-companies');
    if (companyCol >= 0) {
        const counts = countValues(rows, companyCol);
        const top10 = counts.slice(0, 10);

        if (companiesChart) companiesChart.destroy();
        companiesChart = new Chart(companiesCanvas, {
            type: 'bar',
            data: {
                labels: top10.map((c) => truncate(c[0], 20)),
                datasets: [{
                    data: top10.map((c) => c[1]),
                    backgroundColor: colors,
                    borderRadius: 6,
                    maxBarThickness: 40,
                }],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                indexAxis: 'y',
                plugins: { legend: { display: false } },
                scales: {
                    x: { grid: { display: false }, ticks: { precision: 0 } },
                    y: { grid: { display: false } },
                },
            },
        });
        companiesCanvas.parentElement.style.height = '360px';
    }

    // Roles chart
    const roleCol = findColumn(headers, ['role', 'title', 'job title', 'position', 'job role']);
    const rolesCanvas = document.getElementById('chart-roles');
    if (roleCol >= 0) {
        const counts = countValues(rows, roleCol);
        const top8 = counts.slice(0, 8);

        if (rolesChart) rolesChart.destroy();
        rolesChart = new Chart(rolesCanvas, {
            type: 'doughnut',
            data: {
                labels: top8.map((c) => truncate(c[0], 25)),
                datasets: [{
                    data: top8.map((c) => c[1]),
                    backgroundColor: colors,
                    borderWidth: 0,
                    spacing: 2,
                }],
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                cutout: '55%',
                plugins: {
                    legend: {
                        position: 'right',
                        labels: { boxWidth: 12, padding: 12, font: { size: 12 } },
                    },
                },
            },
        });
        rolesCanvas.parentElement.style.height = '360px';
    }
}

function renderTable() {
    const { headers } = allData;

    // Render header
    const thead = document.getElementById('table-head');
    thead.innerHTML = '<tr>' + headers.map((h, i) =>
        `<th onclick="sortTable(${i})">${escapeHtml(h)} ${sortCol === i ? (sortAsc ? '↑' : '↓') : ''}</th>`
    ).join('') + '</tr>';

    // Render body (paginated)
    const start = (currentPage - 1) * CONFIG.pageSize;
    const pageRows = filteredRows.slice(start, start + CONFIG.pageSize);

    const tbody = document.getElementById('table-body');
    tbody.innerHTML = pageRows.map((row) =>
        '<tr>' + headers.map((_, i) =>
            `<td title="${escapeHtml(String(row[i] || ''))}">${escapeHtml(String(row[i] || ''))}</td>`
        ).join('') + '</tr>'
    ).join('');

    // Footer
    document.getElementById('table-count').textContent = `${filteredRows.length} contact${filteredRows.length !== 1 ? 's' : ''}`;

    // Pagination
    const totalPages = Math.ceil(filteredRows.length / CONFIG.pageSize);
    const pagination = document.getElementById('pagination');
    if (totalPages <= 1) {
        pagination.innerHTML = '';
        return;
    }

    let btns = '';
    if (currentPage > 1) btns += `<button onclick="goToPage(${currentPage - 1})">‹</button>`;
    for (let p = 1; p <= totalPages; p++) {
        if (p === 1 || p === totalPages || Math.abs(p - currentPage) <= 2) {
            btns += `<button class="${p === currentPage ? 'active' : ''}" onclick="goToPage(${p})">${p}</button>`;
        } else if (Math.abs(p - currentPage) === 3) {
            btns += '<button disabled>…</button>';
        }
    }
    if (currentPage < totalPages) btns += `<button onclick="goToPage(${currentPage + 1})">›</button>`;
    pagination.innerHTML = btns;
}

// ============================================================
// Table interactions
// ============================================================
function filterTable() {
    const query = document.getElementById('search-input').value.toLowerCase().trim();
    if (!query) {
        filteredRows = [...allData.rows];
    } else {
        filteredRows = allData.rows.filter((row) =>
            row.some((cell) => cell && String(cell).toLowerCase().includes(query))
        );
    }
    currentPage = 1;
    renderTable();
}

function sortTable(colIndex) {
    if (sortCol === colIndex) {
        sortAsc = !sortAsc;
    } else {
        sortCol = colIndex;
        sortAsc = true;
    }

    filteredRows.sort((a, b) => {
        const va = a[colIndex] || '';
        const vb = b[colIndex] || '';
        const na = Number(va);
        const nb = Number(vb);
        if (!isNaN(na) && !isNaN(nb)) return sortAsc ? na - nb : nb - na;
        return sortAsc
            ? String(va).localeCompare(String(vb))
            : String(vb).localeCompare(String(va));
    });

    currentPage = 1;
    renderTable();
}

function goToPage(page) {
    currentPage = page;
    renderTable();
}

// ============================================================
// Export
// ============================================================
function exportCSV() {
    const { headers } = allData;
    const csvRows = [headers.map(csvEscape).join(',')];
    filteredRows.forEach((row) => {
        csvRows.push(headers.map((_, i) => csvEscape(String(row[i] || ''))).join(','));
    });

    const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `contacts_export_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
}

function csvEscape(str) {
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
        return '"' + str.replace(/"/g, '""') + '"';
    }
    return str;
}

// ============================================================
// Helpers
// ============================================================
function findColumn(headers, keywords) {
    const lower = headers.map((h) => h.toLowerCase());
    for (const kw of keywords) {
        const idx = lower.findIndex((h) => h.includes(kw));
        if (idx >= 0) return idx;
    }
    return -1;
}

function countValues(rows, colIndex) {
    const counts = {};
    rows.forEach((row) => {
        const val = row[colIndex] ? String(row[colIndex]).trim() : '';
        if (val) counts[val] = (counts[val] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
}

function truncate(str, maxLen) {
    return str.length > maxLen ? str.slice(0, maxLen) + '…' : str;
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

function showError(msg) {
    document.getElementById('error-message').textContent = msg;
    document.getElementById('error-banner').classList.remove('hidden');
}

function hideError() {
    document.getElementById('error-banner').classList.add('hidden');
}
