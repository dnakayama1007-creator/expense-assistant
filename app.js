/* ============================================
   経費精算アシスタント - Application Logic
   ============================================ */

// ============================================
// Data Store
// ============================================
const STORAGE_KEY = 'expense_assistant_data';

const categoryIcons = {
    '交通費': '🚃',
    '宿泊費': '🏨',
    '会議費': '☕',
    '接待交際費': '🍽️',
    '通信費': '📱',
    '消耗品費': '🖊️',
    '書籍・研修費': '📚',
    'その他': '📦'
};

const categoryColors = {
    '交通費': '#6c5ce7',
    '宿泊費': '#00cec9',
    '会議費': '#fdcb6e',
    '接待交際費': '#e17055',
    '通信費': '#74b9ff',
    '消耗品費': '#a29bfe',
    '書籍・研修費': '#55efc4',
    'その他': '#fd79a8'
};

function loadData() {
    // Synchronous initial load via XHR (called once at startup)
    try {
        var xhr = new XMLHttpRequest();
        xhr.open('GET', '/api/expenses', false); // synchronous
        xhr.send();
        if (xhr.status === 200) {
            var data = JSON.parse(xhr.responseText) || [];
            // Migration: populate orderDate from date if missing
            var migrated = false;
            data.forEach(function (e) {
                if (!e.orderDate && e.date) {
                    e.orderDate = e.date;
                    migrated = true;
                }
            });
            if (migrated) {
                // Save migrated data back to server asynchronously
                fetch('/api/expenses', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(data)
                });
            }
            return data;
        }
    } catch (e) {
        console.error('サーバーからのデータ読み込みに失敗しました', e);
    }
    return [];
}

function saveData(expenses) {
    fetch('/api/expenses', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(expenses)
    }).catch(function (e) {
        console.error('データの保存に失敗しました', e);
        showToast('データの保存に失敗しました', 'error');
    });
}

function generateId() {
    return `exp_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
}

// ============================================
// App State
// ============================================
let expenses = loadData();

// Sequential number management
function getNextSeqNo() {
    var maxSeq = 0;
    expenses.forEach(function (e) { if (e.seqNo && e.seqNo > maxSeq) maxSeq = e.seqNo; });
    return maxSeq + 1;
}

function assignSeqNos() {
    var needAssign = expenses.filter(function (e) { return !e.seqNo; });
    if (needAssign.length === 0) return;
    needAssign.sort(function (a, b) {
        var d = (a.date || '').localeCompare(b.date || '');
        if (d !== 0) return d;
        return (a.createdAt || '').localeCompare(b.createdAt || '');
    });
    var next = getNextSeqNo();
    needAssign.forEach(function (e) { e.seqNo = next++; });
    saveData(expenses);
}
assignSeqNos();

// Status helpers
function getStatusClass(status) {
    switch (status) {
        case '清算済み': return 'status-approved';
        case '入庫済み': return 'status-stocked';
        case '自宅着': return 'status-arrived';
        default: return 'status-pending';
    }
}
function getStatusIcon(status) {
    switch (status) {
        case '清算済み': return '✅';
        case '入庫済み': return '🏢';
        case '自宅着': return '📦';
        default: return '⏳';
    }
}

// ============================================
// Discord Sync
// ============================================
let lastDiscordSync = 0;
const DISCORD_SYNC_INTERVAL = 5000; // 5秒ごとに同期

async function syncDiscordExpenses() {
    try {
        const response = await fetch('/api/discord-expenses');
        if (!response.ok) return;
        const discordExpenses = await response.json();
        if (!Array.isArray(discordExpenses) || discordExpenses.length === 0) return;

        // Find new expenses not yet in local storage
        const existingIds = new Set(expenses.map(e => e.id));
        const newExpenses = discordExpenses.filter(e => !existingIds.has(e.id) && e.amount > 0);

        if (newExpenses.length > 0) {
            var newExpenseRefs = [];
            newExpenses.forEach(de => {
                var newExp = {
                    id: de.id,
                    seqNo: getNextSeqNo(),
                    date: de.date,
                    orderDate: de.orderDate || de.date,
                    amount: de.amount,
                    category: de.category,
                    payment: de.payment || '',
                    description: de.description,
                    unitPrice: de.unitPrice || de.amount || 0,
                    quantity: de.quantity || 1,
                    stockQuantity: de.quantity || 1,
                    receipt: (function (r) {
                        if (!r) return null;
                        var base = r;
                        while (base.includes('/receipts/')) { base = base.substring(base.lastIndexOf('/receipts/') + '/receipts/'.length); }
                        return '/receipts/' + base;
                    })(de.receipt),
                    status: de.status || '未清算',
                    source: 'discord',
                    buyer: de.buyer || '',
                    discordMsgId: de.discordMsgId || '',
                    createdAt: de.createdAt
                };
                expenses.push(newExp);
                newExpenseRefs.push(newExp);
            });
            saveData(expenses);

            // Auto-enrich each new expense with master data (product name, category, supplier)
            newExpenseRefs.forEach(function (exp) {
                enrichExpenseWithMasterData(exp);
            });

            // Auto-detect single-line items (no store, no description): mark as purchase not required
            newExpenseRefs.forEach(function (exp) {
                if (!exp.payment && !exp.description) {
                    exp.purchaseRequired = false;
                }
            });

            // Auto-OCR for new expenses with receipts
            newExpenseRefs.forEach(function (exp) {
                if (exp.receipt) {
                    ocrExpenseReceipt(exp);
                }
            });

            // Update UI if on dashboard
            if (document.getElementById('page-dashboard').classList.contains('active')) {
                updateDashboard();
            }
            if (document.getElementById('page-list').classList.contains('active')) {
                renderExpenseTable();
            }

            showToast('🤖 Discordから' + newExpenses.length + '件の経費を同期しました', 'success');
        }
    } catch (e) {
        // Server not running, silently ignore
    }
}

// ============================================
// OCR Request Counter (monthly)
// ============================================
function getOcrMonthKey() {
    var now = new Date();
    return 'ocrCount_' + now.getFullYear() + '_' + (now.getMonth() + 1);
}

function getOcrCount() {
    var key = getOcrMonthKey();
    return parseInt(localStorage.getItem(key)) || 0;
}

function incrementOcrCount() {
    var key = getOcrMonthKey();
    var count = getOcrCount() + 1;
    localStorage.setItem(key, count);
    updateOcrCountBadge();
    return count;
}

function updateOcrCountBadge() {
    var badge = document.getElementById('ocrCountBadge');
    if (!badge) return;
    var count = getOcrCount();
    badge.textContent = '📷 ' + count + '/700';
    if (count >= 700) {
        badge.style.color = '#ff6b6b';
        badge.style.background = '#ff6b6b22';
    } else if (count >= 500) {
        badge.style.color = '#f39c12';
        badge.style.background = '#f39c1222';
    } else {
        badge.style.color = '#888';
        badge.style.background = '#ffffff11';
    }
}

// Initialize badge on load
updateOcrCountBadge();

// OCR: 領収書画像から注文番号・合計金額を読み取り
function ocrExpenseReceipt(exp) {
    if (!exp.receipt) return;

    // Check monthly limit
    if (getOcrCount() >= 700) {
        console.log('OCR skipped (monthly limit reached):', exp.seqNo);
        return;
    }

    incrementOcrCount();

    fetch('/api/ocr', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ imagePath: exp.receipt, payment: exp.payment || '', supplier: exp.supplier || '' })
    }).then(function (res) { return res.json(); })
        .then(function (data) {
            if (data.success) {
                if (data.orderNumber) exp.orderNumber = data.orderNumber;
                if (data.totalAmount) exp.totalFromReceipt = parseInt(data.totalAmount) || 0;
                saveData(expenses);
                if (document.getElementById('page-list').classList.contains('active')) {
                    renderExpenseTable();
                }
                console.log('OCR complete:', exp.seqNo, 'Order:', data.orderNumber, 'Total:', data.totalAmount);
            }
        }).catch(function (err) {
            console.error('OCR error for', exp.seqNo, ':', err);
        });
}

// Start sync loop
setInterval(syncDiscordExpenses, DISCORD_SYNC_INTERVAL);
// Initial sync after page load
setTimeout(syncDiscordExpenses, 1000);

// ============================================
// DOM References
// ============================================
const sidebar = document.getElementById('sidebar');
const menuToggle = document.getElementById('menuToggle');
const navItems = document.querySelectorAll('.nav-item');
const pages = document.querySelectorAll('.page');
const pageTitle = document.getElementById('pageTitle');
const currentDateEl = document.getElementById('currentDate');

// Dashboard
const statTotal = document.getElementById('statTotal');
const statCount = document.getElementById('statCount');
const statPending = document.getElementById('statPending');
const statApproved = document.getElementById('statApproved');
const recentList = document.getElementById('recentList');
const categoryLegend = document.getElementById('categoryLegend');

// Form
const expenseForm = document.getElementById('expenseForm');
const dropzone = document.getElementById('dropzone');
const receiptFile = document.getElementById('receiptFile');
const receiptPreview = document.getElementById('receiptPreview');
const receiptImage = document.getElementById('receiptImage');
const removeReceipt = document.getElementById('removeReceipt');

// List
const expenseTableBody = document.getElementById('expenseTableBody');
const filterMonth = document.getElementById('filterMonth');
const filterCategory = document.getElementById('filterCategory');
const filterStatus = document.getElementById('filterStatus');
const tableSummary = document.getElementById('tableSummary');

// Export
const exportCSV = document.getElementById('exportCSV');
const exportAllCSV = document.getElementById('exportAllCSV');
const exportBackup = document.getElementById('exportBackup');
const importBackup = document.getElementById('importBackup');
const importFile = document.getElementById('importFile');

// Modal
const receiptModal = document.getElementById('receiptModal');
const modalOverlay = document.getElementById('modalOverlay');
const modalClose = document.getElementById('modalClose');
const modalImage = document.getElementById('modalImage');

// Toast
const toastContainer = document.getElementById('toastContainer');

// ============================================
// Navigation
// ============================================
const pageTitles = {
    'dashboard': 'ダッシュボード',
    'add': '経費を追加',
    'list': '経費一覧',
    'export': 'エクスポート'
};

function navigateTo(page) {
    navItems.forEach(item => {
        item.classList.toggle('active', item.dataset.page === page);
    });
    pages.forEach(p => {
        p.classList.toggle('active', p.id === `page-${page}`);
    });
    pageTitle.textContent = pageTitles[page] || '';

    // Refresh page-specific data
    if (page === 'dashboard') updateDashboard();
    if (page === 'list') renderExpenseTable();

    // Close sidebar on mobile
    sidebar.classList.remove('open');
}

navItems.forEach(item => {
    if (item.dataset.page) {
        item.addEventListener('click', () => navigateTo(item.dataset.page));
    }
});

menuToggle.addEventListener('click', () => {
    sidebar.classList.toggle('open');
});

// ============================================
// Current Date
// ============================================
function updateCurrentDate() {
    const now = new Date();
    const options = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'short' };
    currentDateEl.textContent = now.toLocaleDateString('ja-JP', options);
}
updateCurrentDate();

// Set default date for form
document.getElementById('expenseDate').valueAsDate = new Date();

// Set default month for filters
const now = new Date();
const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
filterMonth.value = currentMonth;
document.getElementById('exportFrom').value = currentMonth;
document.getElementById('exportTo').value = currentMonth;

// ============================================
// Dashboard
// ============================================
function updateDashboard() {
    const now = new Date();
    const thisMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

    const monthExpenses = expenses.filter(e => e.date.startsWith(thisMonth));
    const total = monthExpenses.reduce((sum, e) => sum + e.amount, 0);
    const pendingCount = monthExpenses.filter(e => e.status === '未清算').length;
    const arrivedCount = monthExpenses.filter(e => e.status === '自宅着').length;
    const stockedCount = monthExpenses.filter(e => e.status === '入庫済み').length;
    const printedCount = 0;
    const approvedCount = monthExpenses.filter(e => e.status === '清算済み').length;

    // Animate stat values
    animateValue(statTotal, `¥${total.toLocaleString()}`);
    animateValue(statCount, `${monthExpenses.length}件`);
    animateValue(statPending, `${pendingCount}件`);
    document.getElementById('statArrived').textContent = `${arrivedCount}件`;
    document.getElementById('statStocked').textContent = `${stockedCount}件`;
    document.getElementById('statPrinted').textContent = `${printedCount}件`;
    animateValue(statApproved, `${approvedCount}件`);

    // Recent expenses
    renderRecentExpenses();

    // Charts
    renderCategoryChart(monthExpenses);
    renderMonthlyChart();
}

function animateValue(el, newValue) {
    el.style.opacity = '0';
    el.style.transform = 'translateY(10px)';
    setTimeout(() => {
        el.textContent = newValue;
        el.style.transition = 'all 0.3s ease';
        el.style.opacity = '1';
        el.style.transform = 'translateY(0)';
    }, 100);
}

function renderRecentExpenses() {
    const recent = [...expenses]
        .sort((a, b) => new Date(b.date) - new Date(a.date))
        .slice(0, 5);

    if (recent.length === 0) {
        recentList.innerHTML = '<p class="empty-message">まだ経費が登録されていません</p>';
        return;
    }

    recentList.innerHTML = recent.map(e => `
        <div class="recent-item">
            <div class="recent-item-left">
                ${e.receipt
            ? `<img src="${e.receipt}" class="recent-item-receipt" onclick="showReceiptModal('${e.id}')" alt="領収書" title="クリックで拡大">`
            : `<div class="recent-item-icon">${categoryIcons[e.category] || '📦'}</div>`
        }
                <div class="recent-item-info">
                    <span class="recent-item-desc">${escapeHtml(e.description)}</span>
                    <span class="recent-item-meta">
                        ${formatDate(e.date)} ・ ${e.category}${e.payment ? ` ・ 🏪 ${e.payment}` : ''}
                        ${e.source === 'discord' ? ' ・ <span class="discord-badge">Discord</span>' : ''}
                    </span>
                </div>
            </div>
            <div class="recent-item-right">
                <span class="status-badge ${getStatusClass(e.status)}" style="font-size:11px;margin-right:8px;">
                    ${getStatusIcon(e.status)} ${e.status}
                </span>
                <span class="recent-item-amount">¥${e.amount.toLocaleString()}</span>
            </div>
        </div>
    `).join('');
}

// ============================================
// Charts - Canvas Drawing
// ============================================
function renderCategoryChart(monthExpenses) {
    const canvas = document.getElementById('categoryChart');
    const ctx = canvas.getContext('2d');

    // Set canvas size
    const container = canvas.parentElement;
    canvas.width = container.offsetWidth * 2;
    canvas.height = container.offsetHeight * 2;
    ctx.scale(2, 2);
    const width = container.offsetWidth;
    const height = container.offsetHeight;

    ctx.clearRect(0, 0, width, height);

    // Group by category
    const categoryTotals = {};
    monthExpenses.forEach(e => {
        categoryTotals[e.category] = (categoryTotals[e.category] || 0) + e.amount;
    });

    const categories = Object.keys(categoryTotals);
    const total = Object.values(categoryTotals).reduce((a, b) => a + b, 0);

    if (categories.length === 0) {
        ctx.fillStyle = '#606080';
        ctx.font = '14px Noto Sans JP';
        ctx.textAlign = 'center';
        ctx.fillText('データがありません', width / 2, height / 2);
        categoryLegend.innerHTML = '';
        return;
    }

    // Draw donut chart
    const centerX = width / 2;
    const centerY = height / 2;
    const radius = Math.min(width, height) / 2.5;
    const innerRadius = radius * 0.55;
    let startAngle = -Math.PI / 2;

    categories.forEach(cat => {
        const value = categoryTotals[cat];
        const sliceAngle = (value / total) * Math.PI * 2;
        const color = categoryColors[cat] || '#a29bfe';

        ctx.beginPath();
        ctx.arc(centerX, centerY, radius, startAngle, startAngle + sliceAngle);
        ctx.arc(centerX, centerY, innerRadius, startAngle + sliceAngle, startAngle, true);
        ctx.closePath();
        ctx.fillStyle = color;
        ctx.fill();

        startAngle += sliceAngle;
    });

    // Center text
    ctx.fillStyle = '#e8e8f0';
    ctx.font = 'bold 18px Inter, Noto Sans JP';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(`¥${total.toLocaleString()}`, centerX, centerY - 8);
    ctx.fillStyle = '#9090b0';
    ctx.font = '11px Noto Sans JP';
    ctx.fillText('今月の合計', centerX, centerY + 14);

    // Legend
    categoryLegend.innerHTML = categories.map(cat => `
        <div class="legend-item">
            <span class="legend-dot" style="background:${categoryColors[cat]}"></span>
            <span>${categoryIcons[cat]} ${cat}: ¥${categoryTotals[cat].toLocaleString()}</span>
        </div>
    `).join('');
}

function renderMonthlyChart() {
    const canvas = document.getElementById('monthlyChart');
    const ctx = canvas.getContext('2d');

    const container = canvas.parentElement;
    canvas.width = container.offsetWidth * 2;
    canvas.height = container.offsetHeight * 2;
    ctx.scale(2, 2);
    const width = container.offsetWidth;
    const height = container.offsetHeight;

    ctx.clearRect(0, 0, width, height);

    // Get last 6 months
    const months = [];
    const now = new Date();
    for (let i = 5; i >= 0; i--) {
        const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
        months.push({
            key: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`,
            label: `${d.getMonth() + 1}月`
        });
    }

    const monthlyTotals = months.map(m => {
        return expenses
            .filter(e => e.date.startsWith(m.key))
            .reduce((sum, e) => sum + e.amount, 0);
    });

    const maxValue = Math.max(...monthlyTotals, 1);
    const padding = { top: 20, right: 20, bottom: 40, left: 60 };
    const chartW = width - padding.left - padding.right;
    const chartH = height - padding.top - padding.bottom;

    // Grid lines
    ctx.strokeStyle = 'rgba(108, 92, 231, 0.08)';
    ctx.lineWidth = 1;
    for (let i = 0; i <= 4; i++) {
        const y = padding.top + (chartH / 4) * i;
        ctx.beginPath();
        ctx.moveTo(padding.left, y);
        ctx.lineTo(width - padding.right, y);
        ctx.stroke();

        // Y-axis labels
        const val = maxValue - (maxValue / 4) * i;
        ctx.fillStyle = '#606080';
        ctx.font = '10px Inter';
        ctx.textAlign = 'right';
        ctx.textBaseline = 'middle';
        ctx.fillText(`¥${Math.round(val / 1000)}k`, padding.left - 8, y);
    }

    if (monthlyTotals.every(v => v === 0)) {
        ctx.fillStyle = '#606080';
        ctx.font = '14px Noto Sans JP';
        ctx.textAlign = 'center';
        ctx.fillText('データがありません', width / 2, height / 2);
        return;
    }

    // Bars
    const barWidth = chartW / months.length * 0.5;
    const gap = chartW / months.length;

    months.forEach((m, i) => {
        const barH = (monthlyTotals[i] / maxValue) * chartH;
        const x = padding.left + gap * i + gap / 2 - barWidth / 2;
        const y = padding.top + chartH - barH;

        // Gradient bar
        const gradient = ctx.createLinearGradient(x, y, x, y + barH);
        gradient.addColorStop(0, '#a29bfe');
        gradient.addColorStop(1, '#6c5ce7');

        ctx.beginPath();
        const r = 4;
        ctx.moveTo(x + r, y);
        ctx.lineTo(x + barWidth - r, y);
        ctx.quadraticCurveTo(x + barWidth, y, x + barWidth, y + r);
        ctx.lineTo(x + barWidth, y + barH);
        ctx.lineTo(x, y + barH);
        ctx.lineTo(x, y + r);
        ctx.quadraticCurveTo(x, y, x + r, y);
        ctx.closePath();
        ctx.fillStyle = gradient;
        ctx.fill();

        // X-axis labels
        ctx.fillStyle = '#9090b0';
        ctx.font = '12px Noto Sans JP';
        ctx.textAlign = 'center';
        ctx.fillText(m.label, padding.left + gap * i + gap / 2, height - padding.bottom + 20);

        // Value on top
        if (monthlyTotals[i] > 0) {
            ctx.fillStyle = '#e8e8f0';
            ctx.font = '10px Inter';
            ctx.textAlign = 'center';
            ctx.fillText(`¥${(monthlyTotals[i] / 1000).toFixed(1)}k`, padding.left + gap * i + gap / 2, y - 8);
        }
    });
}

// ============================================
// Expense Form
// ============================================
let currentReceiptData = null;

// Dropzone
dropzone.addEventListener('click', () => receiptFile.click());

dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('drag-over');
});

dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('drag-over');
});

dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file && file.type.startsWith('image/')) {
        handleReceiptFile(file);
    }
});

receiptFile.addEventListener('change', (e) => {
    if (e.target.files[0]) {
        handleReceiptFile(e.target.files[0]);
    }
});

function handleReceiptFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        currentReceiptData = e.target.result;
        receiptImage.src = currentReceiptData;
        receiptPreview.style.display = 'block';
        dropzone.style.display = 'none';

        // Start OCR processing
        performOCR(currentReceiptData);
    };
    reader.readAsDataURL(file);
}

// ============================================
// OCR - Receipt Text Recognition
// ============================================
const ocrStatus = document.getElementById('ocrStatus');
const ocrStatusText = document.getElementById('ocrStatusText');
const ocrProgressFill = document.getElementById('ocrProgressFill');
const ocrProgressLabel = document.getElementById('ocrProgressLabel');
const ocrResults = document.getElementById('ocrResults');
const ocrDetectedAmounts = document.getElementById('ocrDetectedAmounts');
const ocrRawText = document.getElementById('ocrRawText');
const ocrToggleDetail = document.getElementById('ocrToggleDetail');

let ocrWorker = null;

// Toggle raw text display
ocrToggleDetail.addEventListener('click', () => {
    const isHidden = ocrRawText.style.display === 'none';
    ocrRawText.style.display = isHidden ? 'block' : 'none';
    ocrToggleDetail.textContent = isHidden ? '詳細を隠す' : '詳細を見る';
});

async function performOCR(imageDataUrl) {
    // Show loading state
    ocrStatus.style.display = 'block';
    ocrResults.style.display = 'none';
    ocrStatusText.textContent = '🔄 OCRエンジンを準備中...';
    ocrProgressFill.style.width = '0%';
    ocrProgressLabel.textContent = '0%';

    try {
        // Use Tesseract.js recognize directly
        const result = await Tesseract.recognize(
            imageDataUrl,
            'jpn+eng',
            {
                logger: (m) => {
                    if (m.status === 'recognizing text') {
                        const pct = Math.round(m.progress * 100);
                        ocrProgressFill.style.width = `${pct}%`;
                        ocrProgressLabel.textContent = `${pct}%`;
                        ocrStatusText.textContent = '🔍 テキストを認識中...';
                    } else if (m.status === 'loading language traineddata') {
                        const pct = Math.round(m.progress * 100);
                        ocrProgressFill.style.width = `${pct * 0.3}%`;
                        ocrStatusText.textContent = '📥 言語データを読み込み中...';
                    } else if (m.status === 'initializing api') {
                        ocrStatusText.textContent = '⚙️ OCRエンジンを初期化中...';
                    }
                }
            }
        );

        const recognizedText = result.data.text;
        console.log('OCR Result:', recognizedText);

        // Hide loading, show results
        ocrStatus.style.display = 'none';
        ocrResults.style.display = 'block';

        // Display raw text
        ocrRawText.innerHTML = `<span class="ocr-raw-text-label">📄 認識テキスト</span>${escapeHtml(recognizedText)}`;
        ocrRawText.style.display = 'none';
        ocrToggleDetail.textContent = '詳細を見る';

        // Extract amounts
        const amounts = extractAmounts(recognizedText);

        if (amounts.length > 0) {
            renderOCRAmounts(amounts);
            showToast(`🤖 ${amounts.length}件の金額を検出しました`, 'success');
        } else {
            ocrDetectedAmounts.innerHTML = '<p class="ocr-no-amount">⚠️ 金額を検出できませんでした。手動で入力してください。</p>';
            showToast('⚠️ 金額を自動検出できませんでした', 'info');
        }

    } catch (error) {
        console.error('OCR Error:', error);
        ocrStatus.style.display = 'none';
        showToast('❌ OCR処理に失敗しました', 'error');
    }
}

function extractAmounts(text) {
    const amounts = [];
    const seen = new Set();

    // Normalize text: convert full-width numbers to half-width
    let normalized = text;
    normalized = normalized.replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    normalized = normalized.replace(/，/g, ',');
    normalized = normalized.replace(/．/g, '.');
    normalized = normalized.replace(/￥/g, '¥');

    // Pattern 1: ¥ followed by number (e.g., ¥1,234 or ¥ 1234)
    const yenPatterns = normalized.matchAll(/[¥\\]\s*([0-9]{1,3}(?:,?[0-9]{3})*)/g);
    for (const match of yenPatterns) {
        const amount = parseInt(match[1].replace(/,/g, ''));
        if (amount > 0 && amount < 10000000 && !seen.has(amount)) {
            seen.add(amount);
            amounts.push({ value: amount, context: match[0].trim(), type: '¥記号' });
        }
    }

    // Pattern 2: number followed by 円 (e.g., 1,234円 or 1234 円)
    const enPatterns = normalized.matchAll(/([0-9]{1,3}(?:,?[0-9]{3})*)\s*円/g);
    for (const match of enPatterns) {
        const amount = parseInt(match[1].replace(/,/g, ''));
        if (amount > 0 && amount < 10000000 && !seen.has(amount)) {
            seen.add(amount);
            amounts.push({ value: amount, context: match[0].trim(), type: '円表記' });
        }
    }

    // Pattern 3: Look for amounts near keywords like 合計, 小計, 税込, お支払い, etc.
    const keywords = ['合計', '小計', '税込', 'お支払', '請求', '金額', 'TOTAL', 'Total', 'total', 'Amount'];
    const lines = normalized.split('\n');
    for (const line of lines) {
        for (const keyword of keywords) {
            if (line.includes(keyword)) {
                const lineAmounts = line.matchAll(/([0-9]{1,3}(?:,?[0-9]{3})*)/g);
                for (const m of lineAmounts) {
                    const amount = parseInt(m[1].replace(/,/g, ''));
                    if (amount > 0 && amount < 10000000 && !seen.has(amount)) {
                        seen.add(amount);
                        amounts.push({ value: amount, context: `${keyword}: ${m[1]}`, type: keyword });
                    }
                }
            }
        }
    }

    // Sort: prefer larger amounts (likely to be total), then by type priority
    const typePriority = { '合計': 0, '税込': 1, 'お支払': 2, '請求': 3, '¥記号': 4, '円表記': 5 };
    amounts.sort((a, b) => {
        const pa = typePriority[a.type] ?? 10;
        const pb = typePriority[b.type] ?? 10;
        if (pa !== pb) return pa - pb;
        return b.value - a.value;
    });

    return amounts;
}

function renderOCRAmounts(amounts) {
    ocrDetectedAmounts.innerHTML = amounts.map((a, i) => `
        <button type="button" class="ocr-amount-btn" onclick="applyOCRAmount(${a.value}, this)" title="${escapeHtml(a.context)}">
            <div>
                <span class="amount-label">${i === 0 ? '🏆 推奨' : a.type}</span>
                <span class="amount-value">¥${a.value.toLocaleString()}</span>
            </div>
        </button>
    `).join('') + '<p class="ocr-amount-hint">👆 クリックすると金額欄に自動入力されます</p>';

    // Auto-apply the first (most likely) amount
    const amountInput = document.getElementById('expenseAmount');
    if (amounts.length > 0 && !amountInput.value) {
        amountInput.value = amounts[0].value;
        amountInput.dispatchEvent(new Event('input'));
        // Highlight the first button
        const firstBtn = ocrDetectedAmounts.querySelector('.ocr-amount-btn');
        if (firstBtn) firstBtn.classList.add('applied');
        showToast(`💴 金額 ¥${amounts[0].value.toLocaleString()} を自動入力しました`, 'success');
    }
}

function applyOCRAmount(amount, btnElement) {
    const amountInput = document.getElementById('expenseAmount');
    amountInput.value = amount;
    amountInput.dispatchEvent(new Event('input'));

    // Visual feedback
    document.querySelectorAll('.ocr-amount-btn').forEach(btn => btn.classList.remove('applied'));
    btnElement.classList.add('applied');

    // Flash the amount input
    const inputWrapper = amountInput.closest('.input-with-prefix');
    inputWrapper.style.borderColor = 'var(--success)';
    inputWrapper.style.boxShadow = '0 0 0 3px rgba(0, 206, 201, 0.3)';
    setTimeout(() => {
        inputWrapper.style.borderColor = '';
        inputWrapper.style.boxShadow = '';
    }, 1500);

    showToast(`💴 金額 ¥${amount.toLocaleString()} を入力しました`, 'success');
}

function resetOCR() {
    ocrStatus.style.display = 'none';
    ocrResults.style.display = 'none';
    ocrDetectedAmounts.innerHTML = '';
    ocrRawText.innerHTML = '';
    ocrRawText.style.display = 'none';
}


removeReceipt.addEventListener('click', () => {
    currentReceiptData = null;
    receiptImage.src = '';
    receiptPreview.style.display = 'none';
    dropzone.style.display = 'block';
    receiptFile.value = '';
    resetOCR();
});

// Form Submit
expenseForm.addEventListener('submit', (e) => {
    e.preventDefault();

    const unitPrice = parseInt(document.getElementById('expenseUnitPrice').value) || 0;
    const quantity = parseInt(document.getElementById('expenseQuantity').value) || 1;

    const expense = {
        id: generateId(),
        seqNo: getNextSeqNo(),
        date: document.getElementById('expenseDate').value,
        unitPrice: unitPrice,
        quantity: quantity,
        amount: unitPrice * quantity,
        category: document.getElementById('expenseCategory').value,
        payment: document.getElementById('expensePayment').value,
        description: document.getElementById('expenseDescription').value,
        receipt: currentReceiptData,
        status: '未清算',
        createdAt: new Date().toISOString()
    };

    expenses.push(expense);
    saveData(expenses);

    // Reset form
    expenseForm.reset();
    document.getElementById('expenseDate').valueAsDate = new Date();
    document.getElementById('expenseQuantity').value = 1;
    document.getElementById('calculatedAmount').textContent = '¥0';
    currentReceiptData = null;
    receiptImage.src = '';
    receiptPreview.style.display = 'none';
    dropzone.style.display = 'block';
    resetOCR();

    showToast('✅ 経費を登録しました', 'success');

    // Switch to list
    setTimeout(() => navigateTo('list'), 800);
});

expenseForm.addEventListener('reset', () => {
    currentReceiptData = null;
    receiptImage.src = '';
    receiptPreview.style.display = 'none';
    dropzone.style.display = 'block';
    receiptFile.value = '';
    document.getElementById('expenseDate').valueAsDate = new Date();
    resetOCR();
});

// Auto-calculate amount = unitPrice × quantity
function updateCalculatedAmount() {
    const unitPrice = parseInt(document.getElementById('expenseUnitPrice').value) || 0;
    const quantity = parseInt(document.getElementById('expenseQuantity').value) || 1;
    const total = unitPrice * quantity;
    document.getElementById('calculatedAmount').textContent = `¥${total.toLocaleString()}`;
    document.getElementById('expenseAmount').value = total;
}
document.getElementById('expenseUnitPrice').addEventListener('input', updateCalculatedAmount);
document.getElementById('expenseQuantity').addEventListener('input', updateCalculatedAmount);

// ============================================
// Expense Table
// ============================================
function renderExpenseTable() {
    let filtered = [...expenses];

    // Apply filters
    const monthVal = filterMonth.value;
    const catVal = filterCategory.value;
    const statusVal = filterStatus.value;

    if (monthVal) {
        filtered = filtered.filter(e => e.date.startsWith(monthVal));
    }
    if (catVal) {
        filtered = filtered.filter(e => e.category === catVal);
    }
    if (statusVal) {
        filtered = filtered.filter(e => e.status === statusVal);
    }


    // Sort: ① status (custom order), ② buyer desc, ③ category desc, ④ seqNo desc
    var statusOrder = { '未清算': 0, '自宅着': 1, '入庫済み': 2, '清算済み': 3 };
    filtered.sort(function (a, b) {
        var sA = statusOrder[a.status] != null ? statusOrder[a.status] : 99;
        var sB = statusOrder[b.status] != null ? statusOrder[b.status] : 99;
        if (sA !== sB) return sA - sB;
        // buyer descending
        var buyA = a.buyer || '';
        var buyB = b.buyer || '';
        if (buyA !== buyB) return buyB.localeCompare(buyA, 'ja');
        // category descending
        var catA = a.category || '';
        var catB = b.category || '';
        if (catA !== catB) return catB.localeCompare(catA, 'ja');
        // seqNo descending
        return (b.seqNo || 0) - (a.seqNo || 0);
    });


    if (filtered.length === 0) {
        expenseTableBody.innerHTML = '<tr><td colspan="13" class="empty-table">条件に一致する経費データがありません</td></tr>';
        tableSummary.textContent = '合計: ¥0（0件）';
        return;
    }

    const total = filtered.reduce((sum, e) => sum + e.amount, 0);
    tableSummary.textContent = `合計: ¥${total.toLocaleString()}（${filtered.length}件）`;

    expenseTableBody.innerHTML = filtered.map(function (e) {
        var receiptHtml = e.receipt
            ? '<img src="' + escapeHtml(e.receipt) + '" class="receipt-thumb" data-action="receipt" data-eid="' + escapeHtml(e.id) + '" alt="領収書">'
            : '<span class="no-receipt">なし</span>';
        var statusHtml = '<span class="status-badge ' + getStatusClass(e.status) + '" data-action="status" data-eid="' + escapeHtml(e.id) + '" style="cursor:pointer">'
            + getStatusIcon(e.status) + ' ' + e.status + '</span>';
        var qty = e.quantity || 1;
        var stockQty = e.stockQuantity != null ? e.stockQuantity : qty;
        var isSplitDelivery = stockQty < qty;
        var outstandingQty = e.outstandingQuantity || (isSplitDelivery ? qty - stockQty : 0);
        var hasOutstanding = outstandingQty > 0;
        var stockQtyDisplay = e.stockQuantity != null ? e.stockQuantity : '';

        // OCR amount check: mismatch or unreadable
        var ocrMismatch = false;
        if (!e.ocrVerified) {
            if (e.receipt && !e.totalFromReceipt) {
                ocrMismatch = true; // OCR couldn't read amount
            } else if (e.totalFromReceipt && e.totalFromReceipt !== e.amount) {
                ocrMismatch = true; // Amount mismatch
            }
        }

        // Unverified 不要 items also need confirmation
        var unverifiedNotRequired = (e.purchaseRequired === false && !e.ocrVerified);

        // Combined: needs verification?
        var needsVerification = ocrMismatch || unverifiedNotRequired;

        // Row style: red for needs verification, orange for outstanding, bright blue for verified 不要
        var rowStyle = '';
        if (needsVerification) {
            rowStyle = ' style="background:rgba(255,107,107,0.15);"';
        } else if (hasOutstanding) {
            rowStyle = ' style="background:rgba(243,156,18,0.15);"';
        } else if (e.purchaseRequired === false) {
            rowStyle = ' style="background:rgba(100,181,246,0.18);"';
        }

        // 確認済 button for rows needing verification
        var verifyBtn = needsVerification
            ? '<button class="btn btn-sm" data-action="verify" data-eid="' + escapeHtml(e.id) + '" title="確認済" style="background:#00b89422;color:#00b894;border:1px solid #00b89455;cursor:pointer;padding:6px 10px;font-size:12px;">確認済</button>'
            : '';

        return '<tr data-id="' + escapeHtml(e.id) + '"' + rowStyle + '>'
            + '<td><div class="action-buttons" style="display:flex;gap:4px;flex-wrap:wrap;">'
            + '<button class="btn btn-sm" data-action="detail" data-eid="' + escapeHtml(e.id) + '" title="詳細" style="background:#6c5ce722;color:#6c5ce7;border:1px solid #6c5ce755;cursor:pointer;padding:6px 10px;font-size:12px;">詳細</button>'
            + '<button class="btn btn-edit btn-sm" data-action="edit" data-eid="' + escapeHtml(e.id) + '" title="編集">✏️</button>'
            + '<button class="btn btn-sm" data-action="delete" data-eid="' + escapeHtml(e.id) + '" title="削除" style="background:#ff6b6b22;color:#ff6b6b;border:1px solid #ff6b6b55;cursor:pointer;padding:6px 10px;font-size:12px;">削除</button>'
            + verifyBtn
            + '</div></td>'
            + '<td>' + receiptHtml + '</td>'
            + '<td>' + escapeHtml(e.payment || '') + '</td>'
            + '<td style="font-size:11px;">' + escapeHtml(e.orderNumber || '') + '</td>'
            + '<td style="font-size:11px;">' + escapeHtml(e.buyer || '') + '</td>'
            + '<td>' + statusHtml + '</td>'
            + '<td>' + escapeHtml(e.productName || '') + '</td>'
            + '<td>¥' + (e.unitPrice || e.amount || 0).toLocaleString() + '</td>'
            + '<td>' + qty + '</td>'
            + '<td>' + stockQtyDisplay + '</td>'
            + '<td style="font-weight:600;color:var(--accent-secondary)">¥' + e.amount.toLocaleString() + '</td>'
            + '<td>' + formatDate(e.date) + '</td>'
            + '<td style="text-align:center;font-weight:600;color:#f39c12;">' + (hasOutstanding ? outstandingQty : '') + '</td>'
            + '<td style="font-size:11px;color:#aaa;max-width:120px;word-break:break-all;">' + escapeHtml(e.memo || '') + '</td>'
            + '<td class="seq-no-cell">' + (e.seqNo || '-') + '</td>'
            + '</tr>';
    }).join('');

    // Event delegation for table actions
    expenseTableBody.onclick = function (evt) {
        var el = evt.target.closest('[data-action]');
        if (!el) return;
        var action = el.getAttribute('data-action');
        var eid = el.getAttribute('data-eid');
        if (action === 'detail') showDetailModal(eid);
        else if (action === 'edit') editExpense(eid);
        else if (action === 'delete') deleteExpense(eid);
        else if (action === 'verify') verifyOcrAmount(eid);
        else if (action === 'status') toggleStatus(eid);
        else if (action === 'receipt') showReceiptModal(eid);
    };
}

// Inline editing
function editExpense(id) {
    const expense = expenses.find(e => e.id === id);
    if (!expense) return;

    const row = document.querySelector(`tr[data-id="${id}"]`);
    if (!row) return;

    const unitPrice = expense.unitPrice || expense.amount || 0;
    const qty = expense.quantity || 1;
    const purchaseRequired = expense.purchaseRequired !== false; // default true

    row.innerHTML = `
        <td>
            <div class="action-buttons" style="display:flex;gap:4px;flex-wrap:wrap;">
                <button class="btn btn-save btn-sm" onclick="saveExpenseEdit('${id}')" title="保存">💾</button>
                <button class="btn btn-cancel btn-sm" onclick="renderExpenseTable()" title="キャンセル">✖</button>
                <button class="btn btn-sm" id="edit-purchase-${id}" onclick="togglePurchaseRequired('${id}')"
                    style="background:${purchaseRequired ? '#00b89422' : '#fdcb6e22'};color:${purchaseRequired ? '#00b894' : '#fdcb6e'};border:1px solid ${purchaseRequired ? '#00b89455' : '#fdcb6e55'};cursor:pointer;padding:6px 10px;font-size:11px;">
                    仕入${purchaseRequired ? '要' : '不要'}
                </button>
            </div>
        </td>
        <td>
            ${expense.receipt
            ? `<img src="${expense.receipt}" class="receipt-thumb" alt="領収書">`
            : '<span class="no-receipt">なし</span>'}
        </td>
        <td>${escapeHtml(expense.payment || '')}</td>
        <td style="font-size:11px;">${escapeHtml(expense.orderNumber || '')}</td>
        <td style="font-size:11px;">${escapeHtml(expense.buyer || '')}</td>
        <td>
            <span class="status-badge ${getStatusClass(expense.status)}">
                ${getStatusIcon(expense.status)} ${expense.status}
            </span>
        </td>
        <td>${escapeHtml(expense.productName || '')}</td>
        <td><input type="number" class="edit-input edit-input-sm" id="edit-unitprice-${id}" value="${unitPrice}" min="0" oninput="updateEditAmount('${id}')"></td>
        <td><input type="number" class="edit-input edit-input-xs" id="edit-qty-${id}" value="${qty}" min="1" oninput="updateEditAmount('${id}')"></td>
        <td><input type="number" class="edit-input edit-input-xs" id="edit-stockqty-${id}" value="${expense.stockQuantity != null ? expense.stockQuantity : qty}" min="0"></td>
        <td style="font-weight:600;color:var(--accent-secondary)" id="edit-amount-${id}">¥${expense.amount.toLocaleString()}</td>
        <td><input type="date" class="edit-input" id="edit-date-${id}" value="${expense.date}"></td>
        <td><textarea class="edit-input" id="edit-memo-${id}" rows="2" style="width:120px;font-size:11px;padding:4px 6px;resize:vertical;">${escapeHtml(expense.memo || '')}</textarea></td>
        <td class="seq-no-cell">${expense.seqNo || '-'}</td>
    `;
    row.classList.add('editing-row');
    // Store current state for toggle
    row.dataset.purchaseRequired = purchaseRequired ? 'true' : 'false';
}

// Toggle 仕入要/不要
function togglePurchaseRequired(id) {
    var row = document.querySelector(`tr[data-id="${id}"]`);
    if (!row) return;
    var current = row.dataset.purchaseRequired === 'true';
    var newVal = !current;
    row.dataset.purchaseRequired = newVal ? 'true' : 'false';
    var btn = document.getElementById(`edit-purchase-${id}`);
    if (btn) {
        btn.textContent = '仕入' + (newVal ? '要' : '不要');
        btn.style.background = newVal ? '#00b89422' : '#fdcb6e22';
        btn.style.color = newVal ? '#00b894' : '#fdcb6e';
        btn.style.borderColor = newVal ? '#00b89455' : '#fdcb6e55';
    }
}

function updateEditAmount(id) {
    const unitPrice = parseInt(document.getElementById(`edit-unitprice-${id}`).value) || 0;
    const qty = parseInt(document.getElementById(`edit-qty-${id}`).value) || 1;
    const total = unitPrice * qty;
    document.getElementById(`edit-amount-${id}`).textContent = `¥${total.toLocaleString()}`;
}

function saveExpenseEdit(id) {
    const expense = expenses.find(e => e.id === id);
    if (!expense) return;

    const row = document.querySelector(`tr[data-id="${id}"]`);
    const newDate = document.getElementById(`edit-date-${id}`).value;
    const newUnitPrice = parseInt(document.getElementById(`edit-unitprice-${id}`).value) || 0;
    const newQty = parseInt(document.getElementById(`edit-qty-${id}`).value) || 1;
    const newStockQty = parseInt(document.getElementById(`edit-stockqty-${id}`).value);
    const purchaseRequired = row ? row.dataset.purchaseRequired !== 'false' : true;

    expense.date = newDate;
    expense.unitPrice = newUnitPrice;
    expense.quantity = newQty;
    expense.stockQuantity = isNaN(newStockQty) ? newQty : newStockQty;
    expense.amount = newUnitPrice * newQty;
    expense.purchaseRequired = purchaseRequired;
    const memoEl = document.getElementById(`edit-memo-${id}`);
    if (memoEl) expense.memo = memoEl.value.trim();

    saveData(expenses);
    renderExpenseTable();
    updateDashboard();
    showToast('✅ 経費を更新しました', 'success');
}

// Filter events
filterMonth.addEventListener('change', renderExpenseTable);
filterCategory.addEventListener('change', renderExpenseTable);
filterStatus.addEventListener('change', renderExpenseTable);

// ============================================
// Expense Actions
// ============================================
function toggleStatus(id) {
    const expense = expenses.find(e => e.id === id);
    if (expense) {
        // Block status change if verification is pending
        if (!expense.ocrVerified) {
            var needsVerify = false;
            if (expense.receipt && !expense.totalFromReceipt) {
                needsVerify = true;
            } else if (expense.totalFromReceipt && expense.totalFromReceipt !== expense.amount) {
                needsVerify = true;
            } else if (expense.purchaseRequired === false) {
                needsVerify = true;
            }
            if (needsVerify) {
                showToast('⚠️ 確認が必要です。「確認済」ボタンを押してからステータスを変更してください', 'error');
                return;
            }
        }
        // Cycle: 未清算 → 自宅着 → 入庫済み → 清算済み → 未清算
        const statusCycle = { '未清算': '自宅着', '自宅着': '入庫済み', '入庫済み': '清算済み', '清算済み': '未清算' };
        expense.status = statusCycle[expense.status] || '未清算';
        saveData(expenses);
        renderExpenseTable();
        updateDashboard();
        showToast(`ステータスを「${expense.status}」に変更しました`, 'info');
    }
}

var pendingDeleteId = null;

function deleteExpense(id) {
    pendingDeleteId = id;
    var exp = expenses.find(function (e) { return e.id === id; });
    var label = exp ? (exp.seqNo || '') + ' ' + (exp.description || '') : id;
    // Show custom confirm modal
    var overlay = document.createElement('div');
    overlay.id = 'deleteConfirmOverlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:9999;display:flex;align-items:center;justify-content:center;';
    overlay.innerHTML = '<div style="background:#1e1e2e;border:2px solid #ff6b6b;border-radius:16px;padding:32px;max-width:400px;text-align:center;color:#fff;box-shadow:0 8px 32px rgba(0,0,0,0.5);">'
        + '<h3 style="margin:0 0 16px;font-size:18px;">🗑️ 削除確認</h3>'
        + '<p style="margin:0 0 24px;color:#ccc;font-size:14px;">No.' + escapeHtml(String(label)) + '<br>を削除しますか？</p>'
        + '<div style="display:flex;gap:12px;justify-content:center;">'
        + '<button id="deleteConfirmYes" style="background:#ff6b6b;color:#fff;border:none;padding:10px 28px;border-radius:8px;cursor:pointer;font-size:14px;font-weight:600;">削除する</button>'
        + '<button id="deleteConfirmNo" style="background:#444;color:#fff;border:none;padding:10px 28px;border-radius:8px;cursor:pointer;font-size:14px;">キャンセル</button>'
        + '</div></div>';
    document.body.appendChild(overlay);

    document.getElementById('deleteConfirmYes').onclick = function () {
        expenses = expenses.filter(function (e) { return e.id !== pendingDeleteId; });
        saveData(expenses);
        renderExpenseTable();
        updateDashboard();
        showToast('🗑️ 経費を削除しました', 'info');
        overlay.remove();
        pendingDeleteId = null;
    };
    document.getElementById('deleteConfirmNo').onclick = function () {
        overlay.remove();
        pendingDeleteId = null;
    };
}

// ============================================
// OCR Verify (確認済)
// ============================================
function verifyOcrAmount(id) {
    var exp = expenses.find(function (e) { return e.id === id; });
    if (!exp) return;
    exp.ocrVerified = true;
    saveData(expenses);
    renderExpenseTable();
    showToast('✅ 金額確認済みにしました', 'success');
}

// ============================================
// Detail Modal
// ============================================
function showDetailModal(id) {
    var exp = expenses.find(function (e) { return e.id === id; });
    if (!exp) return;

    var qty = exp.quantity || 1;
    var stockQty = exp.stockQuantity != null ? exp.stockQuantity : '';
    var outstandingQty = exp.outstandingQuantity || 0;
    var receiptImg = exp.receipt
        ? '<img src="' + escapeHtml(exp.receipt) + '" style="max-width:100%;max-height:300px;border-radius:8px;margin-top:8px;">'
        : '<span style="color:#888;">なし</span>';

    var rows = [
        ['No.', exp.seqNo || '-'],
        ['日付', formatDate(exp.date)],
        ['注文日', formatDate(exp.orderDate || exp.date)],
        ['注文番号', exp.orderNumber || ''],
        ['ステータス', getStatusIcon(exp.status) + ' ' + exp.status],
        ['内容', exp.description || ''],
        ['商品名', exp.productName || ''],
        ['商品カテゴリ', exp.productCategory || ''],
        ['単価', '¥' + (exp.unitPrice || exp.amount || 0).toLocaleString()],
        ['数量', qty],
        ['入庫数', stockQty],
        ['未着数', outstandingQty > 0 ? outstandingQty : ''],
        ['金額', '¥' + (exp.amount || 0).toLocaleString()],
        ['画像読取金額', exp.totalFromReceipt ? '¥' + exp.totalFromReceipt.toLocaleString() + (exp.totalFromReceipt !== exp.amount ? ' ⚠️' : ' ✅') : '未読取'],
        ['購入店舗', exp.payment || ''],
        ['仕入業者', exp.supplier || ''],
        ['購入者', exp.buyer || ''],
        ['メモ', exp.memo || '']
    ];

    var tableHtml = rows.map(function (r) {
        return '<tr>'
            + '<td style="padding:8px 12px;color:#aaa;white-space:nowrap;font-size:13px;">' + r[0] + '</td>'
            + '<td style="padding:8px 12px;color:#fff;font-size:13px;">' + escapeHtml(String(r[1])) + '</td>'
            + '</tr>';
    }).join('');

    var overlay = document.createElement('div');
    overlay.id = 'detailModalOverlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:9999;display:flex;align-items:center;justify-content:center;';
    overlay.innerHTML = '<div style="background:#1e1e2e;border:2px solid #6c5ce7;border-radius:16px;padding:24px;max-width:500px;width:90%;max-height:80vh;overflow-y:auto;color:#fff;box-shadow:0 8px 32px rgba(0,0,0,0.5);">'
        + '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;">'
        + '<h3 style="margin:0;font-size:18px;">📋 明細詳細</h3>'
        + '<button id="detailModalClose" style="background:none;border:none;color:#aaa;font-size:20px;cursor:pointer;padding:4px 8px;">✕</button>'
        + '</div>'
        + '<table style="width:100%;border-collapse:collapse;">' + tableHtml + '</table>'
        + '<div style="margin-top:12px;padding:8px 12px;">'
        + '<span style="color:#aaa;font-size:13px;">領収書</span><br>' + receiptImg
        + '</div>'
        + '</div>';
    document.body.appendChild(overlay);

    document.getElementById('detailModalClose').onclick = function () { overlay.remove(); };
    overlay.onclick = function (evt) { if (evt.target === overlay) overlay.remove(); };
}

// ============================================
// Receipt Modal
// ============================================
function showReceiptModal(id) {
    const expense = expenses.find(e => e.id === id);
    if (expense && expense.receipt) {
        modalImage.src = expense.receipt;
        receiptModal.classList.add('active');
    }
}

modalClose.addEventListener('click', () => receiptModal.classList.remove('active'));
modalOverlay.addEventListener('click', () => receiptModal.classList.remove('active'));

document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
        receiptModal.classList.remove('active');
    }
});

// ============================================
// Export Functions
// ============================================
exportCSV.addEventListener('click', () => {
    const from = document.getElementById('exportFrom').value;
    const to = document.getElementById('exportTo').value;

    if (!from || !to) {
        showToast('期間を指定してください', 'error');
        return;
    }

    const filtered = expenses.filter(e => {
        const month = e.date.substring(0, 7);
        return month >= from && month <= to;
    });

    downloadCSV(filtered, `経費精算_${from}_${to}`);
});

exportAllCSV.addEventListener('click', () => {
    downloadCSV(expenses, `経費精算_全データ`);
});

function downloadCSV(data, filename) {
    if (data.length === 0) {
        showToast('エクスポートするデータがありません', 'error');
        return;
    }

    const BOM = '\uFEFF';
    const headers = ['No.', '日付', 'カテゴリ', '内容', '単価', '数量', '金額', '購入店舗', 'ステータス'];
    const rows = data
        .sort((a, b) => new Date(a.date) - new Date(b.date))
        .map(e => [
            e.seqNo || '',
            e.date,
            e.category,
            `"${e.description.replace(/"/g, '""')}"`,
            e.unitPrice || e.amount || 0,
            e.quantity || 1,
            e.amount,
            e.payment,
            e.status
        ]);

    const csv = BOM + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${filename}.csv`;
    a.click();
    URL.revokeObjectURL(url);

    showToast(`📄 ${data.length}件のデータをCSV出力しました`, 'success');
}

// Backup
exportBackup.addEventListener('click', () => {
    const data = JSON.stringify(expenses, null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `経費バックアップ_${new Date().toISOString().split('T')[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('💾 バックアップを作成しました', 'success');
});

importBackup.addEventListener('click', () => importFile.click());

importFile.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const imported = JSON.parse(e.target.result);
            if (!Array.isArray(imported)) throw new Error('Invalid format');

            if (confirm(`${imported.length}件のデータを復元します。現在のデータは上書きされます。よろしいですか？`)) {
                expenses = imported;
                saveData(expenses);
                updateDashboard();
                renderExpenseTable();
                showToast(`✅ ${imported.length}件のデータを復元しました`, 'success');
            }
        } catch (err) {
            showToast('バックアップファイルの読み込みに失敗しました', 'error');
        }
    };
    reader.readAsText(file);
    importFile.value = '';
});

// ============================================
// マスタデータマッチング共通関数
// ============================================
var variantCodes = ['bk', 'wh', 'bl', 'rd', 'gn', 'sv', 'gd'];

function stripVariants(str) {
    var result = str;
    variantCodes.forEach(function (v) {
        result = result.replace(new RegExp('(^|[_\\-\\s])' + v + '([_\\-\\s]|$)', 'gi'), '$1$2');
    });
    return result.replace(/[\s_\-]+/g, '').toLowerCase();
}

function matchScore(desc, code) {
    var score = 0;
    if (desc.indexOf(code) >= 0) {
        score = code.length;
    } else if (code.indexOf(desc) >= 0 && desc.length > 2) {
        score = desc.length;
    } else {
        var codeWords = code.split(/[,\s_\/]+/);
        var matchedLen = 0;
        codeWords.forEach(function (cw) {
            if (cw.length > 1 && desc.indexOf(cw) >= 0) {
                matchedLen += cw.length;
            }
        });
        if (matchedLen > score) score = matchedLen;
    }
    return score;
}

function enrichExpenseMatchLogic(exp, products, suppliers) {
    var desc = (exp.description || '').toLowerCase().replace(/[\s_\-]+/g, '');
    var descStripped = stripVariants(exp.description || '');
    var bestProduct = null;
    var bestScore = 0;

    products.forEach(function (p) {
        if (!p.code) return;
        var code = p.code.toLowerCase().replace(/[\s_\-]+/g, '');
        var codeStripped = stripVariants(p.code);
        var score = matchScore(desc, code);
        var scoreStripped = matchScore(desc, codeStripped);
        if (scoreStripped > score) score = scoreStripped;
        var scoreBoth = matchScore(descStripped, codeStripped);
        if (scoreBoth > score) score = scoreBoth;
        if (score > bestScore) {
            bestScore = score;
            bestProduct = p;
        }
    });

    if (bestProduct && bestScore >= 3) {
        exp.productName = bestProduct.code;
        exp.productCategory = bestProduct.category;
        console.log('Product match:', exp.description, '->', bestProduct.code, '(score:', bestScore, ')');
    }

    var store = (exp.payment || '').toLowerCase().replace(/[\s_\-]+/g, '');
    if (store) {
        var bestSupplier = null;
        var bestSupScore = 0;
        suppliers.forEach(function (s) {
            if (!s.abbr) return;
            var abbrClean = s.abbr.replace(/^[a-z]_/i, '').toLowerCase().replace(/[\s_\-]+/g, '');
            var company = s.company.toLowerCase().replace(/[\s_\-]+/g, '');
            var score = 0;
            if (store.indexOf(abbrClean) >= 0 || abbrClean.indexOf(store) >= 0) {
                score = Math.max(abbrClean.length, store.length);
            }
            if (store.indexOf(company) >= 0 || company.indexOf(store) >= 0) {
                score = Math.max(score, Math.max(company.length, store.length));
            }
            if (score > bestSupScore) {
                bestSupScore = score;
                bestSupplier = s;
            }
        });
        if (bestSupplier && bestSupScore >= 2) {
            exp.supplier = bestSupplier.abbr;
            console.log('Supplier match:', exp.payment, '->', bestSupplier.abbr, '(score:', bestSupScore, ')');
        }
    }
    return !!(exp.productName || exp.supplier);
}

// 単一明細のマスタデータ参照（Discord同期時に自動呼び出し）
function enrichExpenseWithMasterData(exp) {
    Promise.all([
        fetch('/api/sheets/read?sheet=' + encodeURIComponent('商品マスタ統合')).then(function (r) { return r.json(); }),
        fetch('/api/sheets/read?sheet=' + encodeURIComponent('取引先マスタ')).then(function (r) { return r.json(); })
    ]).then(function (results) {
        var productData = results[0];
        var supplierData = results[1];
        if (!productData.rows || productData.rows.length < 2) return;
        if (!supplierData.rows || supplierData.rows.length < 2) return;

        var products = productData.rows.slice(1).map(function (row) {
            return { code: (row[1] || '').trim(), category: (row[2] || '').trim(), name: (row[3] || '').trim() };
        });
        var suppliers = supplierData.rows.slice(1).map(function (row) {
            return { abbr: (row[2] || '').trim(), company: (row[3] || '').trim() };
        });

        var matched = enrichExpenseMatchLogic(exp, products, suppliers);
        if (matched) {
            saveData(expenses);
            if (document.getElementById('page-list').classList.contains('active')) {
                renderExpenseTable();
            }
            console.log('Auto-enriched expense:', exp.seqNo, exp.description, '->', exp.productName, exp.supplier);
        }
    }).catch(function (err) {
        console.error('Auto-enrich error:', err);
    });
}

// 入庫データ作成（手動実行: 全自宅着明細を一括処理）
// ============================================
function createStockData() {
    var homeArrived = expenses.filter(function (e) { return e.status === '自宅着'; });
    if (homeArrived.length === 0) {
        showToast('⚠️ 「自宅着」ステータスの明細がありません', 'error');
        return;
    }
    showToast('📦 ' + homeArrived.length + '件のマスタデータを取得中...', 'info');

    Promise.all([
        fetch('/api/sheets/read?sheet=' + encodeURIComponent('商品マスタ統合')).then(function (r) { return r.json(); }),
        fetch('/api/sheets/read?sheet=' + encodeURIComponent('取引先マスタ')).then(function (r) { return r.json(); })
    ]).then(function (results) {
        var productData = results[0];
        var supplierData = results[1];
        if (!productData.rows || productData.rows.length < 2) {
            showToast('⚠️ 商品マスタのデータが取得できませんでした', 'error');
            return;
        }
        if (!supplierData.rows || supplierData.rows.length < 2) {
            showToast('⚠️ 取引先マスタのデータが取得できませんでした', 'error');
            return;
        }

        var products = productData.rows.slice(1).map(function (row) {
            return { code: (row[1] || '').trim(), category: (row[2] || '').trim(), name: (row[3] || '').trim() };
        });
        var suppliers = supplierData.rows.slice(1).map(function (row) {
            return { abbr: (row[2] || '').trim(), company: (row[3] || '').trim() };
        });

        var updatedCount = 0;
        homeArrived.forEach(function (exp) {
            if (enrichExpenseMatchLogic(exp, products, suppliers)) updatedCount++;
        });

        saveData(expenses);
        renderExpenseTable();
        updateDashboard();
        navigateTo('list');
        showToast('📦 ' + updatedCount + '/' + homeArrived.length + '件の入庫データを作成しました', 'success');
    }).catch(function (err) {
        console.error('Stock data error:', err);
        showToast('❌ マスタデータの取得に失敗しました: ' + err.message, 'error');
    });
}

// ============================================
// 仕入データ作成 (Create Purchase Data → Google Sheets)
// ============================================
function createPurchaseData(redoMode) {
    var targetStatus = redoMode ? '入庫済み' : '自宅着';
    // Exclude purchaseRequired=false items from data output
    var homeArrived = expenses.filter(function (e) { return e.status === targetStatus && e.purchaseRequired !== false; });
    var homeArrivedNotRequired = expenses.filter(function (e) { return e.status === targetStatus && e.purchaseRequired === false; });

    // Split delivery items: 清算済み with outstanding and stockQuantity filled
    var splitItems = expenses.filter(function (e) {
        return e.status === '清算済み' && e.outstandingQuantity > 0 && e.stockQuantity != null && e.stockQuantity > 0;
    });

    if (homeArrived.length === 0 && splitItems.length === 0) {
        showToast('⚠️ 出力対象の明細がありません', 'error');
        return;
    }

    // Validate: stockQuantity must not exceed quantity (for 自宅着 items)
    var overStockItems = homeArrived.filter(function (exp) {
        var qty = exp.quantity || 1;
        var stockQty = exp.stockQuantity != null ? exp.stockQuantity : qty;
        return stockQty > qty;
    });
    if (overStockItems.length > 0) {
        var labels = overStockItems.map(function (exp) {
            return 'No.' + (exp.seqNo || '?') + ' ' + (exp.description || '');
        }).join('\n');
        showToast('⚠️ 入庫数が数量を超えている明細があります:\n' + labels, 'error');
        return;
    }

    // Validate: stockQuantity must not exceed outstandingQuantity (for split items)
    var overSplitItems = splitItems.filter(function (exp) {
        return exp.stockQuantity > exp.outstandingQuantity;
    });
    if (overSplitItems.length > 0) {
        var labels2 = overSplitItems.map(function (exp) {
            return 'No.' + (exp.seqNo || '?') + ' ' + (exp.description || '') + ' (入庫数:' + exp.stockQuantity + ' > 未着数:' + exp.outstandingQuantity + ')';
        }).join('\n');
        showToast('⚠️ 入庫数が未着数を超えている明細があります:\n' + labels2, 'error');
        return;
    }

    // Format today's date as YYYY/MM/DD
    var today = new Date();
    var todayStr = today.getFullYear() + '/' + ('0' + (today.getMonth() + 1)).slice(-2) + '/' + ('0' + today.getDate()).slice(-2);

    // Build rows for 自宅着 items - split into A-F and H-K (skip G)
    var rowsAF = homeArrived.map(function (exp) {
        var expDate = exp.date ? exp.date.replace(/-/g, '/') : todayStr;
        return [
            expDate,                          // A: 明細の日付
            expDate,                          // B: 明細の日付
            todayStr,                         // C: 実行した日付
            exp.supplier || '',                // D: 仕入業者
            exp.productCategory || '',        // E: 商品カテゴリ
            exp.productName || exp.description || ''   // F: 商品名
        ];
    });
    var rowsHK = homeArrived.map(function (exp) {
        return [
            exp.unitPrice || exp.amount || 0, // H: 単価
            exp.stockQuantity != null ? exp.stockQuantity : (exp.quantity || 1),  // I: 入庫数
            '',                               // J: (空)
            exp.stockQuantity != null ? exp.stockQuantity : (exp.quantity || 1), // K: 入庫数
            '2_武田'                           // L: 担当者
        ];
    });

    // Build rows for split delivery items (I列K列 both = stockQuantity)
    splitItems.forEach(function (exp) {
        var expDate = exp.date ? exp.date.replace(/-/g, '/') : todayStr;
        rowsAF.push([
            expDate,                          // A: 明細の日付
            expDate,                          // B: 明細の日付
            todayStr,                         // C: 実行した日付
            exp.supplier || '',                // D: 仕入業者
            exp.productCategory || '',        // E: 商品カテゴリ
            exp.productName || exp.description || ''   // F: 商品名
        ]);
        rowsHK.push([
            exp.unitPrice || exp.amount || 0, // H: 単価
            exp.stockQuantity,                // I: 入庫数
            '',                               // J: (空)
            exp.stockQuantity,                // K: 入庫数
            '2_武田'                           // L: 担当者
        ]);
    });

    var totalRows = rowsAF.length;
    showToast('🛒 ' + totalRows + '件の仕入データを書き込み中...', 'info');

    fetch('/api/sheets/write', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            rowsAF: rowsAF,
            rowsHK: rowsHK
        })
    }).then(function (res) {
        return res.json();
    }).then(function (data) {
        if (data.success) {
            // Calculate summary
            var totalStockQty = 0;
            homeArrived.forEach(function (exp) {
                totalStockQty += (exp.stockQuantity != null ? exp.stockQuantity : (exp.quantity || 1));
            });
            splitItems.forEach(function (exp) {
                totalStockQty += exp.stockQuantity;
            });
            var consumablesCount = homeArrivedNotRequired.length;

            // Update status only in normal mode (not redo)
            if (!redoMode) {
                // Update 自宅着 items to 入庫済み (both required and not required)
                homeArrived.forEach(function (exp) {
                    exp.status = '入庫済み';
                });
                homeArrivedNotRequired.forEach(function (exp) {
                    exp.status = '入庫済み';
                });
                // Update split delivery items: recalculate outstandingQuantity
                splitItems.forEach(function (exp) {
                    var newOutstanding = exp.outstandingQuantity - exp.stockQuantity;
                    if (newOutstanding <= 0) {
                        exp.outstandingQuantity = 0;
                    } else {
                        exp.outstandingQuantity = newOutstanding;
                    }
                    exp.stockQuantity = null;
                });
                saveData(expenses);
                renderExpenseTable();
                updateDashboard();
            }
            navigateTo('list');

            // Show confirmation dialog
            var overlay = document.createElement('div');
            overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.6);display:flex;align-items:center;justify-content:center;z-index:10000;';
            overlay.innerHTML = '<div style="background:#1e1e2e;border-radius:16px;padding:32px;max-width:400px;width:90%;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,0.5);">'
                + '<div style="font-size:40px;margin-bottom:16px;">✅</div>'
                + '<div style="font-size:18px;font-weight:700;color:#fff;margin-bottom:16px;">' + totalRows + '件の仕入データを書き込みました</div>'
                + '<div style="font-size:16px;color:#ddd;margin-bottom:8px;">入庫数の合計：<span style="color:#74b9ff;font-weight:700;">' + totalStockQty + '個</span></div>'
                + '<div style="font-size:16px;color:#ddd;margin-bottom:24px;">消耗品：<span style="color:#fdcb6e;font-weight:700;">' + consumablesCount + '件</span></div>'
                + '<button id="purchaseSummaryOk" style="background:linear-gradient(135deg,#6c5ce7,#a29bfe);color:#fff;border:none;padding:12px 40px;border-radius:10px;cursor:pointer;font-size:15px;font-weight:600;">OK</button>'
                + '</div>';
            document.body.appendChild(overlay);
            document.getElementById('purchaseSummaryOk').onclick = function () { overlay.remove(); };
        } else {
            showToast('❌ 書き込みエラー: ' + (data.error || '不明なエラー'), 'error');
            console.error('Sheets write error:', data);
        }
    }).catch(function (err) {
        console.error('Purchase data error:', err);
        showToast('❌ 仕入データの書き込みに失敗しました: ' + err.message, 'error');
    });
}

// ============================================
// Print Report (入庫済み)
// ============================================
function generateStockedReport() {

    const stocked = expenses.filter(e => e.status === '入庫済み');

    if (stocked.length === 0) {
        showToast('⚠️ 「入庫済み」ステータスの明細がありません', 'error');
        navigateTo('list');
        return;
    }

    // Show custom confirmation dialog
    var overlay = document.createElement('div');
    overlay.id = 'reportConfirmOverlay';
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.6);z-index:9999;display:flex;align-items:center;justify-content:center;';
    overlay.innerHTML = '<div style="background:#1e1e2e;border:2px solid #6c5ce7;border-radius:16px;padding:32px;max-width:420px;text-align:center;color:#fff;box-shadow:0 8px 32px rgba(0,0,0,0.5);">'
        + '<h3 style="margin:0 0 16px;font-size:18px;">🖨️ レポート作成確認</h3>'
        + '<p style="margin:0 0 24px;color:#ccc;font-size:14px;">仕入と数量チェックはしましたか？</p>'
        + '<div style="display:flex;gap:12px;justify-content:center;">'
        + '<button id="reportConfirmYes" style="background:#6c5ce7;color:#fff;border:none;padding:10px 28px;border-radius:8px;cursor:pointer;font-size:14px;font-weight:600;">はい</button>'
        + '<button id="reportConfirmNo" style="background:#444;color:#fff;border:none;padding:10px 28px;border-radius:8px;cursor:pointer;font-size:14px;">いいえ</button>'
        + '</div></div>';
    document.body.appendChild(overlay);

    document.getElementById('reportConfirmYes').onclick = function () {
        overlay.remove();
        generateStockedReportCore(stocked);
    };
    document.getElementById('reportConfirmNo').onclick = function () {
        overlay.remove();
        showToast('ℹ️ レポート作成をキャンセルしました', 'info');
    };
    return;
}

function generateStockedReportCore(stocked) {
    const totalAmount = stocked.reduce((sum, e) => sum + (e.amount || 0), 0);
    const totalQty = stocked.reduce((sum, e) => sum + (e.quantity || 1), 0);
    const today = new Date();
    const dateStr = today.getFullYear() + '年' + (today.getMonth() + 1) + '月' + today.getDate() + '日';

    var itemCards = stocked.map(function (e, i) {
        var unitPrice = e.unitPrice || e.amount || 0;
        var qty = e.quantity || 1;
        var amount = e.amount || 0;
        var receiptBlock = e.receipt
            ? '<div class="item-receipt"><img src="' + e.receipt + '" class="receipt-image" alt="receipt"></div>'
            : '<div class="item-receipt no-receipt-box"><span>領収書なし</span></div>';
        return '<div class="report-item">' +
            '<div class="item-header"><span class="item-number">No.' + (e.seqNo || (i + 1)) + '</span><span class="item-amount">\u00a5' + amount.toLocaleString() + '</span></div>' +
            receiptBlock +
            '<div class="item-details">' +
            '<div class="detail-row"><span class="dlabel">店舗</span><span class="dval">' + escapeHtml(e.payment || '-') + '</span></div>' +
            '<div class="detail-row"><span class="dlabel">内容</span><span class="dval">' + escapeHtml(e.description) + '</span></div>' +
            '<div class="detail-row"><span class="dlabel">単\u00d7数</span><span class="dval">\u00a5' + unitPrice.toLocaleString() + ' \u00d7 ' + qty + '</span></div>' +
            '<div class="detail-row"><span class="dlabel">日付</span><span class="dval">' + e.date + '</span></div>' +
            '</div></div>';
    });
    var pagesHtml = '';
    for (var pi = 0; pi < itemCards.length; pi += 6) {
        var pageItems = itemCards.slice(pi, pi + 6);
        var pageSub = stocked.slice(pi, pi + 6).reduce(function (s, e) { return s + (e.amount || 0); }, 0);
        var pageQty = stocked.slice(pi, pi + 6).reduce(function (s, e) { return s + (e.quantity || 1); }, 0);
        var isLast = (pi + 6 >= itemCards.length);
        pagesHtml += '<div class="report-page' + (isLast ? '' : ' page-break') + '">' +
            '<div class="page-header"><span class="page-title">経費精算レポート</span>' +
            '<span class="page-info">' + (Math.floor(pi / 6) + 1) + 'ページ / ' + stocked.length + '件</span>' +
            '<span class="page-subtotal">小計: \u00a5' + pageSub.toLocaleString() + ' (' + pageQty + '個)</span></div>' +
            '<div class="items-grid">' + pageItems.join('') + '</div></div>';
    }
    var pageCount = Math.ceil(itemCards.length / 6);

    var reportHtml = '<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">' +
        '<title>経費精算レポート</title><style>' +
        '* { margin:0; padding:0; box-sizing:border-box; }' +
        'body { font-family:"Helvetica Neue","Hiragino Kaku Gothic ProN","Meiryo",sans-serif; padding:10px; color:#333; background:#fff; font-size:9px; }' +
        '.report-header { border:2px solid #333; padding:12px 20px; margin-bottom:10px; display:flex; justify-content:space-between; align-items:center; }' +
        '.report-title { font-size:18px; font-weight:700; }' +
        '.report-meta { text-align:right; }' +
        '.report-meta .date { font-size:11px; color:#666; }' +
        '.report-meta .total-label { font-size:11px; margin-top:2px; }' +
        '.report-meta .total-amount { font-size:22px; font-weight:700; color:#d63031; }' +
        '.report-meta .item-count { font-size:10px; color:#666; }' +
        '.report-meta .total-qty { font-size:13px; font-weight:600; color:#0984e3; margin-top:2px; }' +
        '.report-page { margin-bottom:10px; }' +
        '.page-break { page-break-after:always; }' +
        '.page-header { display:flex; justify-content:space-between; align-items:center; padding:6px 12px; background:#f0f0f0; border:1px solid #ccc; margin-bottom:6px; font-size:10px; }' +
        '.page-title { font-weight:700; font-size:11px; }' +
        '.page-info { color:#666; }' +
        '.page-subtotal { font-weight:700; color:#d63031; font-size:12px; }' +
        '.items-grid { display:grid; grid-template-columns:repeat(3,1fr); gap:6px; }' +
        '.report-item { border:1px solid #bbb; display:flex; flex-direction:column; overflow:hidden; }' +
        '.item-header { background:#f5f5f5; padding:4px 8px; display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid #ccc; }' +
        '.item-number { background:#333; color:#fff; padding:1px 6px; border-radius:3px; font-size:8px; font-weight:600; }' +
        '.item-amount { font-size:12px; font-weight:700; color:#d63031; }' +
        '.item-receipt { height:260px; display:flex; align-items:center; justify-content:center; overflow:hidden; background:#fafafa; }' +
        '.receipt-image { width:100%; height:100%; object-fit:contain; }' +
        '.no-receipt-box { color:#bbb; font-size:10px; }' +
        '.item-details { padding:4px 6px; border-top:1px solid #ddd; }' +
        '.detail-row { display:flex; gap:4px; line-height:1.5; font-size:8px; }' +
        '.dlabel { color:#888; min-width:36px; font-weight:600; }' +
        '.dval { color:#333; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }' +
        '.report-footer { margin-top:10px; padding:10px 20px; border:2px solid #333; display:flex; justify-content:space-between; align-items:center; }' +
        '.footer-label { font-size:14px; font-weight:600; }' +
        '.footer-total { font-size:22px; font-weight:700; color:#d63031; }' +
        '.footer-qty { font-size:14px; font-weight:600; color:#0984e3; margin-left:16px; }' +
        '.print-btn { display:block; margin:16px auto; padding:12px 36px; background:#e17055; color:white; border:none; border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; }' +
        '.print-btn:hover { background:#d35400; }' +
        '@media print { .print-btn{display:none!important;} body{padding:5mm;} .page-break{page-break-after:always;} }' +
        '</style></head><body>' +
        '<div class="report-header"><div class="report-title">📋 経費精算レポート（入庫済み）</div>' +
        '<div class="report-meta"><div class="date">出力日: ' + dateStr + '</div>' +
        '<div class="total-label">合計金額</div>' +
        '<div class="total-amount">\u00a5' + totalAmount.toLocaleString() + '</div>' +
        '<div class="total-qty">合計数量: ' + totalQty + '個</div>' +
        '<div class="item-count">' + stocked.length + '件（' + pageCount + 'ページ）</div></div></div>' +
        pagesHtml +
        '<div class="report-footer"><span class="footer-label">合計金額（' + stocked.length + '件）</span>' +
        '<span class="footer-qty">合計数量: ' + totalQty + '個</span>' +
        '<span class="footer-total">\u00a5' + totalAmount.toLocaleString() + '</span></div>' +
        '<button class="print-btn" onclick="window.print()">🖨️ 印刷 / PDF保存</button>' +
        '</body></html>';

    var printWindow = window.open('', '_blank');
    printWindow.document.write(reportHtml);
    printWindow.document.close();

    showToast('📄 入庫済み ' + stocked.length + '件のレポートを出力しました', 'success');
    navigateTo('list');
}

// 差し戻し: 入庫済み → 自宅着（自宅着がある場合は中断）
function revertPrintedToStocked() {
    try {
        // Block if any 自宅着 items exist
        var homeArrived = expenses.filter(function (e) { return e.status === '自宅着'; });
        if (homeArrived.length > 0) {
            showToast('⚠️ 「自宅着」ステータスの明細が ' + homeArrived.length + ' 件あります。\n先に仕入データ作成を実行してください', 'error');
            return;
        }

        // Revert 入庫済み → 自宅着
        var stocked = expenses.filter(function (e) { return e.status === '入庫済み'; });
        if (stocked.length === 0) {
            showToast('⚠️ 「入庫済み」ステータスの明細がありません', 'error');
            return;
        }
        stocked.forEach(function (e) { e.status = '自宅着'; });
        saveData(expenses);
        updateDashboard();
        renderExpenseTable();
        showToast('↩️ ' + stocked.length + '件を「自宅着」に差し戻しました', 'info');
        navigateTo('list');
    } catch (err) {
        console.error('revertPrintedToStocked error:', err);
        showToast('エラーが発生しました: ' + err.message, 'error');
    }
}


// 入庫完了: 印刷済み → freee送信 → 清算済み
function completePrinted() {
    // 1. Check for 入庫済み items
    var printed = expenses.filter(function (e) { return e.status === '入庫済み'; });
    if (printed.length === 0) {
        showToast('⚠️ 「入庫済み」ステータスの明細がありません', 'error');
        navigateTo('list');
        return;
    }

    showToast('🔄 freee認証確認中... (' + printed.length + '件)', 'info');

    // 2. Check freee auth
    fetch('/api/freee/status')
        .then(function (r) { return r.json(); })
        .then(function (status) {
            if (!status.authenticated) {
                showToast('⚠️ freee認証が必要です。「freee連携」から認証してください', 'error');
                return;
            }

            showToast('🔄 freeeマスタデータを取得中...', 'info');

            // 3. Fetch master data
            return fetch('/api/freee/master')
                .then(function (r) { return r.json(); })
                .then(function (master) {
                    var walletId = null, accountItemId = null, creditAccountItemId = null;
                    var itemId = null, tagIds = [], taxCode = null, creditTaxCode = null;

                    // Wallet: 武田大輔_経費_未払い
                    master.wallets.forEach(function (w) {
                        if (w.name && w.name.indexOf('武田大輔') >= 0 && w.name.indexOf('経費') >= 0) walletId = w.id;
                    });
                    // Account item (debit): 仕入高
                    master.account_items.forEach(function (a) {
                        if (a.name === '仕入高') accountItemId = a.id;
                    });
                    // Item: PCパーツ
                    master.items.forEach(function (item) {
                        if (item.name === 'PCパーツ') itemId = item.id;
                    });
                    // Tag: 武田
                    master.tags.forEach(function (tag) {
                        if (tag.name && tag.name.indexOf('武田') >= 0) tagIds.push(tag.id);
                    });
                    // Tax code (debit): 課対仕入（控80）10%
                    master.tax_codes.forEach(function (t) {
                        if (!taxCode && t.name_ja === '課対仕入（控80）10%') taxCode = t.code;
                    });
                    if (!taxCode) {
                        master.tax_codes.forEach(function (t) {
                            if (!taxCode && t.name_ja && t.name_ja.indexOf('課対仕入') >= 0 && t.name_ja.indexOf('10%') >= 0) taxCode = t.code;
                        });
                    }
                    // Credit account item: 武田大輔_経費_未払い
                    master.account_items.forEach(function (a) {
                        if (a.name && a.name.indexOf('武田大輔') >= 0 && a.name.indexOf('経費') >= 0) creditAccountItemId = a.id;
                    });
                    if (!creditAccountItemId) {
                        master.account_items.forEach(function (a) {
                            if (a.name && a.name.indexOf('武田') >= 0 && a.name.indexOf('未払') >= 0) creditAccountItemId = a.id;
                        });
                    }
                    // Credit tax code: 対象外
                    master.tax_codes.forEach(function (t) {
                        if (!creditTaxCode && t.name_ja && t.name_ja === '対象外') creditTaxCode = t.code;
                    });
                    if (!creditTaxCode) {
                        master.tax_codes.forEach(function (t) {
                            if (!creditTaxCode && t.name_ja && t.name_ja.indexOf('対象外') >= 0) creditTaxCode = t.code;
                        });
                    }

                    // Validate required master IDs
                    if (!walletId) { showToast('⚠️ 口座「武田大輔_経費_未払い」が見つかりません', 'error'); return; }
                    if (!accountItemId) { showToast('⚠️ 勘定科目「仕入高」が見つかりません', 'error'); return; }
                    if (!creditAccountItemId) { showToast('⚠️ 勘定科目「武田大輔_経費_未払い」が見つかりません', 'error'); return; }
                    if (!taxCode) { showToast('⚠️ 税区分「課税仕入」が見つかりません', 'error'); return; }
                    if (!creditTaxCode) { showToast('⚠️ 税区分「対象外」が見つかりません', 'error'); return; }

                    showToast('📤 ' + printed.length + '件をfreeeに送信中...', 'info');

                    // 4. Build expense data and send to freee
                    var expenseData = printed.map(function (e) {
                        return {
                            id: e.id,
                            seqNo: e.seqNo,
                            date: e.date,
                            amount: e.amount,
                            description: 'No.' + (e.seqNo || '') + ' ' + (e.description || e.payment || ''),
                            receipt: e.receipt || null
                        };
                    });

                    return fetch('/api/freee/deals', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            expenses: expenseData,
                            account_item_id: accountItemId,
                            tax_code: taxCode,
                            credit_account_item_id: creditAccountItemId,
                            credit_tax_code: creditTaxCode,
                            item_id: itemId,
                            tag_ids: tagIds
                        })
                    })
                        .then(function (r) { return r.json(); })
                        .then(function (data) {
                            var successIds = new Set(
                                data.results.filter(function (r) { return r.success; }).map(function (r) { return r.expense_id; })
                            );
                            var failedCount = data.results.filter(function (r) { return !r.success; }).length;

                            if (failedCount > 0) {
                                console.error('Failed freee deals:', JSON.stringify(data.results.filter(function (r) { return !r.success; })));
                                showToast('⚠️ freee送信: ' + successIds.size + '件成功 / ' + failedCount + '件失敗\n失敗分はステータスを変更しません', 'error');
                            } else {
                                showToast('✅ freeeに' + successIds.size + '件を登録しました', 'success');
                            }

                            // 5. Update only successfully sent items
                            printed.forEach(function (e) {
                                if (!successIds.has(e.id)) return; // skip failed
                                var qty = e.quantity || 1;
                                var stockQty = e.stockQuantity != null ? e.stockQuantity : qty;
                                // Calculate outstanding quantity for split delivery
                                if (stockQty < qty) {
                                    e.outstandingQuantity = qty - stockQty;
                                    e.stockQuantity = null;
                                }
                                e.status = '清算済み';
                            });

                            saveData(expenses);
                            updateDashboard();
                            renderExpenseTable();
                            navigateTo('list');
                        });
                });
        })
        .catch(function (err) {
            console.error('completePrinted error:', err);
            showToast('エラー: ' + err.message, 'error');
        });
}

// ============================================
// Toast Notifications
// ============================================
function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.textContent = message;
    toastContainer.appendChild(toast);

    setTimeout(() => {
        toast.classList.add('toast-exit');
        setTimeout(() => toast.remove(), 300);
    }, 3000);
}

// ============================================
// Utility Functions
// ============================================
function formatDate(dateStr) {
    const d = new Date(dateStr);
    return `${d.getMonth() + 1}/${d.getDate()}`;
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// ============================================
// Resize Handler for Charts
// ============================================
let resizeTimeout;
window.addEventListener('resize', () => {
    clearTimeout(resizeTimeout);
    resizeTimeout = setTimeout(() => {
        if (document.getElementById('page-dashboard').classList.contains('active')) {
            const now = new Date();
            const thisMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
            const monthExpenses = expenses.filter(e => e.date.startsWith(thisMonth));
            renderCategoryChart(monthExpenses);
            renderMonthlyChart();
        }
    }, 250);
});

// ============================================
// Initialize with Sample Data (for demo)
// ============================================
function addSampleData() {
    if (expenses.length > 0) return;

    const today = new Date();
    const year = today.getFullYear();
    const month = today.getMonth();

    const sampleData = [
        { date: `${year}-${String(month + 1).padStart(2, '0')}-03`, amount: 540, category: '交通費', payment: '電子マネー', description: '池袋駅→新宿駅 往復', status: '清算済み' },
        { date: `${year}-${String(month + 1).padStart(2, '0')}-05`, amount: 1280, category: '会議費', payment: 'クレジットカード', description: 'スターバックス 取引先との打ち合わせ', status: '清算済み' },
        { date: `${year}-${String(month + 1).padStart(2, '0')}-08`, amount: 4500, category: '接待交際費', payment: 'クレジットカード', description: '和食レストラン ○○様との会食', status: '未清算' },
        { date: `${year}-${String(month + 1).padStart(2, '0')}-10`, amount: 980, category: '消耗品費', payment: '現金', description: 'コピー用紙・ボールペン購入', status: '清算済み' },
        { date: `${year}-${String(month + 1).padStart(2, '0')}-12`, amount: 3200, category: '書籍・研修費', payment: 'クレジットカード', description: 'プログラミング技術書 3冊', status: '未清算' },
        { date: `${year}-${String(month + 1).padStart(2, '0')}-15`, amount: 8900, category: '宿泊費', payment: 'クレジットカード', description: '大阪出張 ビジネスホテル 1泊', status: '未清算' },
        { date: `${year}-${String(month + 1).padStart(2, '0')}-18`, amount: 2640, category: '通信費', payment: '銀行振込', description: 'ポケットWiFiレンタル 出張用', status: '清算済み' },
        { date: `${year}-${String(month + 1).padStart(2, '0')}-20`, amount: 1100, category: '交通費', payment: '電子マネー', description: '東京駅→品川駅 タクシー', status: '未清算' },
    ];

    // Also add last month
    const lastMonth = month === 0 ? 12 : month;
    const lastMonthYear = month === 0 ? year - 1 : year;
    const lastMonthSamples = [
        { date: `${lastMonthYear}-${String(lastMonth).padStart(2, '0')}-05`, amount: 2300, category: '交通費', payment: '電子マネー', description: '新幹線 東京→大宮', status: '清算済み' },
        { date: `${lastMonthYear}-${String(lastMonth).padStart(2, '0')}-12`, amount: 5800, category: '接待交際費', payment: 'クレジットカード', description: '居酒屋 チームランチ', status: '清算済み' },
        { date: `${lastMonthYear}-${String(lastMonth).padStart(2, '0')}-22`, amount: 1500, category: '消耗品費', payment: '現金', description: 'USB-Cケーブル・充電器', status: '清算済み' },
    ];

    const allSamples = [...sampleData, ...lastMonthSamples];
    allSamples.forEach(s => {
        expenses.push({
            ...s,
            id: generateId(),
            seqNo: getNextSeqNo(),
            unitPrice: s.unitPrice || s.amount,
            quantity: s.quantity || 1,
            receipt: null,
            createdAt: new Date().toISOString()
        });
    });

    saveData(expenses);
}

// ============================================
// Memo Text Parser
// ============================================
const memoParserToggle = document.getElementById('memoParserToggle');
const memoParserBody = document.getElementById('memoParserBody');
const memoArrow = document.getElementById('memoArrow');
const memoInput = document.getElementById('memoInput');
const parseMemoBtn = document.getElementById('parseMemoBtn');
const clearMemoBtn = document.getElementById('clearMemoBtn');
const memoParseResults = document.getElementById('memoParseResults');
const memoParsedItems = document.getElementById('memoParsedItems');

// Toggle memo parser
memoParserToggle.addEventListener('click', () => {
    memoParserBody.classList.toggle('collapsed');
    memoArrow.classList.toggle('collapsed');
});

// Parse memo
parseMemoBtn.addEventListener('click', () => {
    const text = memoInput.value.trim();
    if (!text) {
        showToast('テキストを入力してください', 'error');
        return;
    }
    const entries = parseMemoText(text);
    if (entries.length === 0) {
        showToast('⚠️ 経費情報を検出できませんでした', 'info');
        memoParseResults.style.display = 'none';
        return;
    }
    renderParsedMemos(entries);
    showToast(`🤖 ${entries.length}件の経費を読み取りました`, 'success');
});

// Clear memo
clearMemoBtn.addEventListener('click', () => {
    memoInput.value = '';
    memoParseResults.style.display = 'none';
    memoParsedItems.innerHTML = '';
});

function parseMemoText(text) {
    // Normalize
    let normalized = text;
    normalized = normalized.replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    normalized = normalized.replace(/，/g, ',');
    normalized = normalized.replace(/￥/g, '¥');

    const lines = normalized.split('\n').filter(l => l.trim().length > 0);
    const entries = [];

    for (const line of lines) {
        const entry = parseSingleLine(line.trim());
        if (entry) {
            entries.push(entry);
        }
    }

    return entries;
}

function parseSingleLine(line) {
    let date = null;
    let amount = null;
    let category = null;
    let description = line;

    // --- Extract Date ---
    const now = new Date();
    const currentYear = now.getFullYear();

    // Pattern: YYYY/M/D or YYYY-M-D
    let dateMatch = line.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (dateMatch) {
        date = `${dateMatch[1]}-${dateMatch[2].padStart(2, '0')}-${dateMatch[3].padStart(2, '0')}`;
        description = description.replace(dateMatch[0], '').trim();
    }

    // Pattern: M/D (e.g., 2/24)
    if (!date) {
        dateMatch = line.match(/(?:^|\s)(\d{1,2})\/(\d{1,2})(?:\s|$|[^\d])/);
        if (dateMatch) {
            const m = parseInt(dateMatch[1]);
            const d = parseInt(dateMatch[2]);
            if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
                date = `${currentYear}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                description = description.replace(dateMatch[0], ' ').trim();
            }
        }
    }

    // Pattern: M月D日
    if (!date) {
        dateMatch = line.match(/(\d{1,2})月(\d{1,2})日/);
        if (dateMatch) {
            const m = parseInt(dateMatch[1]);
            const d = parseInt(dateMatch[2]);
            if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
                date = `${currentYear}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                description = description.replace(dateMatch[0], '').trim();
            }
        }
    }

    // --- Extract Amount ---
    // ¥ amount
    let amountMatch = line.match(/[¥\\]\s*([0-9]{1,3}(?:,?[0-9]{3})*)/);
    if (amountMatch) {
        amount = parseInt(amountMatch[1].replace(/,/g, ''));
        description = description.replace(amountMatch[0], '').trim();
    }

    // number + 円
    if (!amount) {
        amountMatch = line.match(/([0-9]{1,3}(?:,?[0-9]{3})*)\s*円/);
        if (amountMatch) {
            amount = parseInt(amountMatch[1].replace(/,/g, ''));
            description = description.replace(amountMatch[0], '').trim();
        }
    }

    // Standalone large number (likely amount if no ¥ or 円)
    // Pick the LAST standalone number, skip numbers adjacent to letters (model numbers like "9060xt")
    if (!amount) {
        const allMatches = [...line.matchAll(/(?:^|[\s,、])([0-9]{3,7})(?=[\s,、。]|$)/g)];
        const validMatches = allMatches.filter(m => {
            const pos = m.index + m[0].length;
            const before = m.index > 0 ? line[m.index - 1] : '';
            const after = pos < line.length ? line[pos] : '';
            return !/[a-zA-Z]/.test(before) && !/[a-zA-Z]/.test(after);
        });
        if (validMatches.length > 0) {
            const lastMatch = validMatches[validMatches.length - 1];
            const val = parseInt(lastMatch[1]);
            if (val >= 100 && val < 10000000) {
                amount = val;
                description = description.replace(lastMatch[1], '').trim();
            }
        }
    }

    // If no amount found, skip this line
    if (!amount) return null;

    // --- Detect Category ---
    const categoryKeywords = {
        '交通費': ['タクシー', 'タクシ', '電車', '新幹線', 'バス', '地下鉄', '駅', '交通', '定期', 'Suica', 'PASMO', 'IC', '乗車', '往復', '片道', '→', '➡', 'JR', '私鉄', '飛行機', '航空'],
        '宿泊費': ['ホテル', '旅館', '宿泊', '宿', '民泊', 'Airbnb', 'ビジネスホテル', '泊'],
        '会議費': ['スタバ', 'スターバックス', 'カフェ', 'コーヒー', '喫茶', '打ち合わせ', '打合せ', '会議', 'ミーティング', 'MTG', 'ドトール', 'タリーズ'],
        '接待交際費': ['レストラン', '居酒屋', '飲み会', '会食', '接待', '食事', 'ランチ', 'ディナー', '懇親', '忘年会', '新年会', '歓迎会', '送別会', '焼肉', '寿司', 'すし'],
        '通信費': ['通信', 'Wi-Fi', 'WiFi', '電話', '携帯', 'SIM', 'ポケット', 'インターネット', 'ネット'],
        '消耗品費': ['文具', 'コピー', '用紙', 'ペン', 'ノート', 'USB', 'ケーブル', '電池', '消耗品', '事務用品', 'トナー', 'インク', '封筒'],
        '書籍・研修費': ['本', '書籍', '雑誌', '研修', 'セミナー', '講座', '勉強会', '技術書', '参考書', 'Udemy', '受講']
    };

    const originalLine = line;
    for (const [cat, keywords] of Object.entries(categoryKeywords)) {
        for (const keyword of keywords) {
            if (originalLine.includes(keyword)) {
                category = cat;
                break;
            }
        }
        if (category) break;
    }

    // --- Extract Store (購入店舗) ---
    let store = null;
    const storeMatch = line.match(/^([A-Za-z\u3040-\u9FFF]+[\(（][^\)）]+[\)）])/)
        || line.match(/^([A-Za-z]+(?:\.[A-Za-z]+)*)\s/)
        || line.match(/^(楽天|Amazon|アマゾン|Yahoo|ヤフー|メルカリ|ヨドバシ|ビック|ジョーシン|ケーズ|エディオン|ドンキ|コストコ|IKEA|ニトリ|ダイソー|セリア|モノタロウ|AliExpress)\s?/i);
    if (storeMatch) {
        store = (storeMatch[1] || storeMatch[0]).trim();
        description = description.replace(storeMatch[0], '').trim();
    }

    // --- Clean Description ---
    // Remove extra spaces and clean up
    description = description.replace(/\s+/g, ' ').trim();
    // Remove leading/trailing special chars
    description = description.replace(/^[\s,、。・\-\/]+|[\s,、。・\-\/]+$/g, '').trim();

    // If description is empty, use original line cleaned of amount
    if (!description) {
        description = line.replace(/[¥\\]\s*[0-9,.]+|[0-9,.]+\s*円/g, '').replace(/\s+/g, ' ').trim();
    }

    return {
        date: date,
        amount: amount,
        category: category || 'その他',
        store: store || '',
        description: description || 'メモ経費',
        originalLine: line
    };
}

function renderParsedMemos(entries) {
    memoParseResults.style.display = 'block';

    memoParsedItems.innerHTML = entries.map((e, i) => `
        <div class="memo-parsed-item" data-index="${i}">
            <div class="memo-parsed-item-info">
                <div class="memo-parsed-item-main">
                    ${e.date ? `<span class="memo-parsed-date">📅 ${e.date}</span>` : '<span class="memo-parsed-date">📅 今日</span>'}
                    <span class="memo-parsed-category">${categoryIcons[e.category] || '📦'} ${e.category}</span>
                    ${e.store ? `<span class="memo-parsed-date">🏪 ${escapeHtml(e.store)}</span>` : ''}
                </div>
                <span class="memo-parsed-desc">${escapeHtml(e.description)}</span>
            </div>
            <span class="memo-parsed-amount">¥${e.amount.toLocaleString()}</span>
            <button type="button" class="memo-apply-btn" onclick="applyParsedMemo(${i})">📝 フォームに入力</button>
        </div>
    `).join('');

    // Store parsed entries for later use
    window._parsedMemoEntries = entries;
}

function applyParsedMemo(index) {
    const entries = window._parsedMemoEntries;
    if (!entries || !entries[index]) return;

    const e = entries[index];

    // Fill form fields
    if (e.date) {
        document.getElementById('expenseDate').value = e.date;
    }
    document.getElementById('expenseAmount').value = e.amount;
    document.getElementById('expenseCategory').value = e.category;
    document.getElementById('expensePayment').value = e.store || '';
    document.getElementById('expenseDescription').value = e.description;

    // Visual feedback on the button
    const btns = document.querySelectorAll('.memo-apply-btn');
    btns[index].textContent = '✅ 入力済み';
    btns[index].classList.add('applied');

    // Scroll to form
    document.getElementById('expenseForm').scrollIntoView({ behavior: 'smooth', block: 'start' });

    showToast(`📝 「${e.description}」をフォームに入力しました`, 'success');
}

// ============================================
// freee API Integration
// ============================================
function openFreeeSettings() {
    // Check if auth callback
    var urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('freee_auth') === 'success') {
        window.history.replaceState({}, '', '/');
    }

    // Fetch current config
    fetch('/api/freee/config')
        .then(function (r) { return r.json(); })
        .then(function (config) {
            var authStatus = config.has_token
                ? '<span style="color:#00b894;font-weight:600">\u2705 \u8a8d\u8a3c\u6e08\u307f</span>'
                : '<span style="color:#e17055;font-weight:600">\u274c \u672a\u8a8d\u8a3c</span>';
            var companyInfo = config.company_id
                ? '<br>\u4e8b\u696d\u6240ID: ' + config.company_id
                : '';

            var html = '<div style="position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.7);z-index:9999;display:flex;align-items:center;justify-content:center;" id="freeeModal">';
            html += '<div style="background:var(--bg-secondary);border-radius:16px;padding:32px;width:500px;max-width:90vw;color:var(--text-primary);border:1px solid var(--border-color);">';
            html += '<h2 style="margin-bottom:20px;">\ud83d\udd17 freee API \u9023\u643a\u8a2d\u5b9a</h2>';
            html += '<div style="margin-bottom:16px;padding:12px;background:rgba(108,92,231,0.1);border-radius:8px;">';
            html += '<strong>\u8a8d\u8a3c\u30b9\u30c6\u30fc\u30bf\u30b9:</strong> ' + authStatus + companyInfo + '</div>';
            html += '<div style="margin-bottom:12px;"><label style="display:block;margin-bottom:4px;font-size:13px;color:var(--text-secondary);">Client ID</label>';
            html += '<input type="text" id="freeeClientId" placeholder="freee\u30a2\u30d7\u30ea\u306eClient ID" style="width:100%;padding:10px;border-radius:8px;border:1px solid var(--border-color);background:var(--bg-primary);color:var(--text-primary);font-size:14px;"></div>';
            html += '<div style="margin-bottom:16px;"><label style="display:block;margin-bottom:4px;font-size:13px;color:var(--text-secondary);">Client Secret</label>';
            html += '<input type="password" id="freeeClientSecret" placeholder="freee\u30a2\u30d7\u30ea\u306eClient Secret" style="width:100%;padding:10px;border-radius:8px;border:1px solid var(--border-color);background:var(--bg-primary);color:var(--text-primary);font-size:14px;"></div>';
            html += '<p style="font-size:11px;color:var(--text-secondary);margin-bottom:16px;">\u30b3\u30fc\u30eb\u30d0\u30c3\u30afURL: <code>http://localhost:3000/api/freee/callback</code></p>';
            html += '<div style="display:flex;gap:10px;">';
            html += '<button onclick="saveFreeeConfig()" style="flex:1;padding:10px;border-radius:8px;border:none;background:#6c5ce7;color:white;font-weight:600;cursor:pointer;">\ud83d\udcbe \u4fdd\u5b58</button>';
            html += '<button onclick="window.location.href=\'/api/freee/auth\'" style="flex:1;padding:10px;border-radius:8px;border:none;background:#0984e3;color:white;font-weight:600;cursor:pointer;">\ud83d\udd11 freee\u8a8d\u8a3c</button>';
            html += '<button onclick="document.getElementById(\'freeeModal\').remove()" style="flex:1;padding:10px;border-radius:8px;border:1px solid var(--border-color);background:transparent;color:var(--text-primary);font-weight:600;cursor:pointer;">\u2716 \u9589\u3058\u308b</button>';
            html += '</div></div></div>';

            var existing = document.getElementById('freeeModal');
            if (existing) existing.remove();
            document.body.insertAdjacentHTML('beforeend', html);
        })
        .catch(function (err) {
            showToast('\u8a2d\u5b9a\u306e\u53d6\u5f97\u306b\u5931\u6557\u3057\u307e\u3057\u305f: ' + err.message, 'error');
        });
}

function saveFreeeConfig() {
    var clientId = document.getElementById('freeeClientId').value.trim();
    var clientSecret = document.getElementById('freeeClientSecret').value.trim();
    if (!clientId || !clientSecret) {
        showToast('\u26a0\ufe0f Client ID\u3068Client Secret\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044', 'error');
        return;
    }
    fetch('/api/freee/config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ client_id: clientId, client_secret: clientSecret })
    })
        .then(function (r) { return r.json(); })
        .then(function (data) {
            if (data.success) {
                showToast('\u2705 freee\u8a2d\u5b9a\u3092\u4fdd\u5b58\u3057\u307e\u3057\u305f\u3002\u300cfreee\u8a8d\u8a3c\u300d\u30dc\u30bf\u30f3\u3067\u8a8d\u8a3c\u3057\u3066\u304f\u3060\u3055\u3044', 'success');
            }
        })
        .catch(function (err) {
            showToast('\u4fdd\u5b58\u306b\u5931\u6557\u3057\u307e\u3057\u305f: ' + err.message, 'error');
        });
}

function sendToFreee() {
    // Find 印刷済み expenses to send
    var printed = expenses.filter(function (e) { return e.status === '印刷済み'; });
    if (printed.length === 0) {
        showToast('⚠️ 「印刷済み」ステータスの明細がありません', 'error');
        navigateTo('list');
        return;
    }

    showToast('🔄 freeeマスタデータを取得中... (' + printed.length + '件)', 'info');

    // First check auth status
    fetch('/api/freee/status')
        .then(function (r) { return r.json(); })
        .then(function (status) {
            if (!status.authenticated) {
                showToast('\u26a0\ufe0f freee\u8a8d\u8a3c\u304c\u5fc5\u8981\u3067\u3059\u3002\u300cfreee\u9023\u643a\u300d\u304b\u3089\u8a8d\u8a3c\u3057\u3066\u304f\u3060\u3055\u3044', 'error');
                return;
            }

            // Fetch master data
            return fetch('/api/freee/master')
                .then(function (r) { return r.json(); })
                .then(function (master) {
                    // Find IDs for specified values
                    var walletId = null;
                    var accountItemId = null;
                    var creditAccountItemId = null;
                    var itemId = null;
                    var tagIds = [];
                    var taxCode = null;
                    var creditTaxCode = null;

                    // Find wallet: 武田大輔_経費_未払い
                    master.wallets.forEach(function (w) {
                        if (w.name && w.name.indexOf('\u6b66\u7530\u5927\u8f14') >= 0 && w.name.indexOf('\u7d4c\u8cbb') >= 0) {
                            walletId = w.id;
                        }
                    });

                    // Find account item: 仕入高
                    master.account_items.forEach(function (a) {
                        if (a.name === '\u4ed5\u5165\u9ad8') {
                            accountItemId = a.id;
                        }
                    });

                    // Find item: PCパーツ
                    master.items.forEach(function (item) {
                        if (item.name === 'PC\u30d1\u30fc\u30c4') {
                            itemId = item.id;
                        }
                    });

                    // Find tag: 武田 (partial match)
                    master.tags.forEach(function (tag) {
                        if (tag.name && tag.name.indexOf('武田') >= 0) {
                            tagIds.push(tag.id);
                        }
                    });
                    if (tagIds.length === 0) {
                        console.log('⚠️ Tag "武田" not found. Available tags:', JSON.stringify(master.tags.map(function (t) { return t.name; })));
                    }

                    // Find tax code: 課対仕入（控80）10% = code 189
                    // freee uses abbreviated names like '課対仕入' not '課税仕入'
                    master.tax_codes.forEach(function (t) {
                        if (!taxCode && t.name_ja === '課対仕入（控80）10%') {
                            taxCode = t.code;
                        }
                    });
                    if (!taxCode) {
                        master.tax_codes.forEach(function (t) {
                            if (!taxCode && t.name_ja && t.name_ja.indexOf('課対仕入') >= 0 && t.name_ja.indexOf('10%') >= 0) {
                                taxCode = t.code;
                            }
                        });
                    }
                    console.log('taxCode found:', taxCode);

                    // Find credit account item: 武田大輔_経費_未払い (as account item)
                    master.account_items.forEach(function (a) {
                        if (a.name && a.name.indexOf('武田大輔') >= 0 && a.name.indexOf('経費') >= 0) {
                            creditAccountItemId = a.id;
                        }
                    });
                    // Fallback: search by 未払
                    if (!creditAccountItemId) {
                        master.account_items.forEach(function (a) {
                            if (a.name && a.name.indexOf('武田') >= 0 && a.name.indexOf('未払') >= 0) {
                                creditAccountItemId = a.id;
                            }
                        });
                    }

                    // Find credit tax code: 対象外
                    master.tax_codes.forEach(function (t) {
                        if (!creditTaxCode && t.name_ja && t.name_ja === '対象外') {
                            creditTaxCode = t.code;
                        }
                    });
                    if (!creditTaxCode) {
                        master.tax_codes.forEach(function (t) {
                            if (!creditTaxCode && t.name_ja && t.name_ja.indexOf('対象外') >= 0) {
                                creditTaxCode = t.code;
                            }
                        });
                    }

                    // Log found IDs
                    console.log('=== freee IDs found ===');
                    console.log('walletId:', walletId);
                    console.log('accountItemId (debit/仕入高):', accountItemId);
                    console.log('creditAccountItemId (武田大輔_経費):', creditAccountItemId);
                    console.log('itemId:', itemId);
                    console.log('tagIds:', tagIds);
                    console.log('taxCode (debit):', taxCode);
                    console.log('creditTaxCode:', creditTaxCode);

                    if (!walletId) {
                        showToast('\u26a0\ufe0f \u53e3\u5ea7\u300c\u6b66\u7530\u5927\u8f14_\u7d4c\u8cbb_\u672a\u6255\u3044\u300d\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093', 'error');
                        console.log('Available wallets:', JSON.stringify(master.wallets.map(function (w) { return { name: w.name, id: w.id }; })));
                        return;
                    }
                    if (!accountItemId) {
                        showToast('\u26a0\ufe0f \u52d8\u5b9a\u79d1\u76ee\u300c\u4ed5\u5165\u9ad8\u300d\u304c\u898b\u3064\u304b\u308a\u307e\u305b\u3093', 'error');
                        console.log('Available account_items:', JSON.stringify(master.account_items.map(function (a) { return a.name; })));
                        return;
                    }
                    if (!creditAccountItemId) {
                        showToast('⚠️ 勘定科目「武田大輔_経費_未払い」が見つかりません', 'error');
                        console.log('Available account_items:', JSON.stringify(master.account_items.map(function (a) { return { name: a.name, id: a.id }; })));
                        return;
                    }
                    if (!taxCode) {
                        showToast('⚠️ 税区分「課税仕入」が見つかりません', 'error');
                        console.log('All tax codes:', JSON.stringify(master.tax_codes.map(function (t) { return t.name_ja + ' rate:' + t.rate + ' code:' + t.code; })));
                        return;
                    }
                    if (!creditTaxCode) {
                        showToast('⚠️ 税区分「対象外」が見つかりません', 'error');
                        return;
                    }

                    showToast('\ud83d\udce4 ' + printed.length + '\u4ef6\u3092freee\u306b\u9001\u4fe1\u4e2d...', 'info');

                    var expenseData = printed.map(function (e) {
                        return {
                            id: e.id,
                            seqNo: e.seqNo,
                            date: e.date,
                            amount: e.amount,
                            description: 'No.' + (e.seqNo || '') + ' ' + e.description,
                            receipt: e.receipt || null
                        };
                    });

                    return fetch('/api/freee/deals', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            expenses: expenseData,
                            account_item_id: accountItemId,
                            tax_code: taxCode,
                            credit_account_item_id: creditAccountItemId,
                            credit_tax_code: creditTaxCode,
                            item_id: itemId,
                            tag_ids: tagIds
                        })
                    })
                        .then(function (r) { return r.json(); })
                        .then(function (data) {
                            var success = data.results.filter(function (r) { return r.success; }).length;
                            var failed = data.results.filter(function (r) { return !r.success; }).length;

                            if (failed > 0) {
                                console.log('Failed deals:', JSON.stringify(data.results.filter(function (r) { return !r.success; })));
                                showToast('⚠️ freee送信: ' + success + '件成功 / ' + failed + '件失敗', 'error');
                            } else {
                                showToast('✅ freeeに' + success + '件の経費を登録しました！', 'success');
                            }
                            // Update successful items to 清算済み
                            data.results.forEach(function (r) {
                                if (r.success) {
                                    var exp = expenses.find(function (e) { return e.id === r.expense_id; });
                                    if (exp) exp.status = '清算済み';
                                }
                            });
                            saveData(expenses);
                            updateDashboard();
                            renderExpenseTable();
                            navigateTo('list');
                        });
                });
        })
        .catch(function (err) {
            console.error('freee send error:', err);
            showToast('\u30a8\u30e9\u30fc: ' + err.message, 'error');
        });
}

// Check for freee auth callback on load
(function () {
    var urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('freee_auth') === 'success') {
        showToast('\u2705 freee\u8a8d\u8a3c\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\uff01', 'success');
        window.history.replaceState({}, '', '/');
    } else if (urlParams.get('freee_auth') === 'error') {
        showToast('\u274c freee\u8a8d\u8a3c\u306b\u5931\u6557\u3057\u307e\u3057\u305f', 'error');
        window.history.replaceState({}, '', '/');
    }
})();

// ============================================
// Boot
// ============================================
addSampleData();
renderExpenseTable();
// Delay dashboard update to ensure CSS layout is complete and canvas containers have dimensions
requestAnimationFrame(() => {
    requestAnimationFrame(() => {
        updateDashboard();
    });
});
