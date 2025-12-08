/* 
 * Developed by XSimple
 */

// --- Constants & Config ---
const PRESET_ITEMS = [
    "调理腿肉","蔬脆鸡肉饼","熟制板烧肉","调理鸡腿","调理翅中","调理翅根","调理奥尔良翅根","国潮脆皮翅根",
    "熟制烤全腿","调理切全翅","藤椒鸡尖","20支原味鸡","调理四大块","香酥鸡排","安格斯牛肉饼","虾饼",
    "蝴蝶虾","孜然小串","孜然鸡柳","黑椒鸡块","蒜香鸡块","香脆鸡米花","年糕条","烤肠","芝士热狗棒",
    "细薯","菠萝派","蛋挞","鸡翅包饭","面饼","面包","咖喱鸡肉炒饭","酸菜肉末炒饭","榴莲披萨","鸡肉披萨",
    "牛肉披萨","中磨砂杯","大磨砂杯","中可杯","小可杯（12A）","热饮杯","辣脆鸡小胸","开花肠",
    "番茄牛肉意面","黑椒牛肉意面"
]; // Note: List truncated in prompt description slightly, updated based on request flow or added remaining if any omitted in text. 
// Assuming the prompt list was exhaustive. 44 items listed above from manual extraction, let's normalize commas from the prompt string to be safe.

const RAW_ITEMS_STRING = "调理腿肉,蔬脆鸡肉饼，熟制板烧肉，调理鸡腿，调理翅中，调理翅根，调理奥尔良翅根，国潮脆皮翅根，熟制烤全腿，调理切全翅，藤椒鸡尖，20支原味鸡，调理四大块，香酥鸡排，安格斯牛肉饼，虾饼，蝴蝶虾，孜然小串，孜然鸡柳，黑椒鸡块，蒜香鸡块，香脆鸡米花，年糕条，烤肠，芝士热狗棒，细薯，菠萝派，蛋挞，鸡翅包饭，面饼，面包，咖喱鸡肉炒饭，酸菜肉末炒饭，榴莲披萨，鸡肉披萨，牛肉披萨，中磨砂杯，大磨砂杯，中可杯，小可杯（12A），热饮杯，辣脆鸡小胸，开花肠，番茄牛肉意面，黑椒牛肉意面";
// Normalizing split 
const ITEMS_NAMES = RAW_ITEMS_STRING.replace(/，/g, ',').split(',').map(s => s.trim()).filter(s => s);

// --- State Management ---
const db = {
    getItems: () => JSON.parse(localStorage.getItem('xs_items') || '[]'),
    saveItems: (items) => localStorage.setItem('xs_items', JSON.stringify(items)),
    isInit: () => localStorage.getItem('xs_appInitialized') === 'true',
    setInit: () => localStorage.setItem('xs_appInitialized', 'true'),
    getRecords: () => JSON.parse(localStorage.getItem('xs_records') || '[]'),
    addRecord: (record) => {
        const records = db.getRecords();
        records.unshift(record);
        localStorage.setItem('xs_records', JSON.stringify(records));
    }
};

let currentView = 'init'; // init, main, records
let selectedItemsId = [];
let isBatchMode = false;

// --- DOM Elements ---
const app = document.getElementById('app');
const modalOverlay = document.getElementById('modal-overlay');
const modalContent = document.getElementById('modal-content');

// --- Initialization ---
function init() {
    if (!db.isInit()) {
        renderInitView();
    } else {
        renderMainView();
    }
}

// --- Views Rendering ---

function renderInitView() {
    app.innerHTML = `
        <div class="glass init-container">
            <div class="init-header">
                <h2>系统初始化</h2>
                <button class="btn btn-outline" onclick="promptBatchSet()">批量设置起始量</button>
            </div>
            <p style="margin-bottom:15px; color:var(--text-light)">请设置50个预设项目的起始库存量。</p>
            <div id="init-list">
                ${ITEMS_NAMES.map((name, index) => `
                    <div class="init-item">
                        <label>${name}</label>
                        <input type="number" id="start-${index}" placeholder="起始量" min="0" onkeydown="handleEnter(event, ${index})">
                        <input type="text" id="spec-${index}" placeholder="规格(选填)">
                    </div>
                `).join('')}
            </div>
            <div style="margin-top:20px; text-align:right;">
                <button class="btn btn-success" onclick="completeInitialization()">完成初始化</button>
            </div>
        </div>
    `;
}

function renderMainView() {
    const items = db.getItems();
    app.innerHTML = `
        <div class="glass top-bar">
            <div>
                <h1>XSimple Calculcator</h1>
            </div>
            <div class="main-actions">
                ${isBatchMode ? `
                    <button class="btn btn-outline" onclick="toggleSelectAll()">全选</button>
                    <button class="btn btn-success" onclick="startBatchCalc()">去计算选中项</button>
                    <button class="btn btn-danger" onclick="exitBatchMode()">退出批量</button>
                ` : `
                    <button class="btn" onclick="showRecords()">计算记录</button>
                    <input type="file" id="excel-upload" hidden onchange="handleExcelImport(this)">
                    <button class="btn btn-outline" onclick="document.getElementById('excel-upload').click()">导入Excel</button>
                    <button class="btn" onclick="enterBatchMode()">批量计算</button>
                `}
            </div>
        </div>
        <div class="item-grid">
            ${items.map(item => `
                <div class="glass item-card ${isBatchMode && selectedItemsId.includes(item.id) ? 'selected' : ''}" 
                     onclick="handleItemClick('${item.id}')"
                     data-id="${item.id}">
                    <div class="card-header">${item.name}</div>
                    <div class="card-spec">${item.spec || '无规格'}</div>
                    <div class="card-amount">余: ${item.currentAmount}</div>
                    <div class="card-footer">更新: ${formatDate(item.lastUpdated)}</div>
                    ${isBatchMode ? `<div style="position:absolute; top:10px; right:10px;">${selectedItemsId.includes(item.id) ? '✅' : '⬜'}</div>` : ''}
                </div>
            `).join('')}
        </div>
    `;

    // Attach Long Press Events
    if (!isBatchMode) attachLongPressEvents();
}

function renderRecordsView() {
    const records = db.getRecords();
    const items = db.getItems();
    
    app.innerHTML = `
        <div class="glass top-bar">
            <h2>计算记录</h2>
            <button class="btn btn-outline" onclick="renderMainView()">返回主页</button>
        </div>
        <div class="glass" style="padding:16px;">
            <div style="margin-bottom:15px;">
                <input type="text" id="search-record" placeholder="搜索项目名称..." oninput="filterRecords(this.value)">
            </div>
            <div id="records-list">
                ${records.map(r => {
                    const item = items.find(i => i.id === r.itemId);
                    const name = item ? item.name : 'Unknown';
                    return `
                        <div class="record-item">
                            <div class="record-header">
                                <span>${name}</span>
                                <span class="diff-negative">-${r.usage}</span>
                            </div>
                            <div class="record-meta">
                                ${formatDate(r.timestamp)} | 消耗前: ${r.before} -> 剩余: ${r.after}
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        </div>
    `;
}

// --- Logic Implementation ---

// Initialization Logic
function promptBatchSet() {
    const val = prompt("请输入统一的起始量数值：");
    if (val && !isNaN(val)) {
        ITEMS_NAMES.forEach((_, idx) => {
            const input = document.getElementById(`start-${idx}`);
            if (input) input.value = val;
        });
    }
}

function handleEnter(e, idx) {
    if (e.key === 'Enter') {
        const nextInput = document.getElementById(`start-${idx + 1}`);
        if (nextInput) nextInput.focus();
    }
}

function completeInitialization() {
    const items = [];
    let isValid = true;
    
    ITEMS_NAMES.forEach((name, idx) => {
        const startVal = document.getElementById(`start-${idx}`).value;
        const specVal = document.getElementById(`spec-${idx}`).value;
        
        if (startVal === '' || startVal < 0) {
            isValid = false;
            document.getElementById(`start-${idx}`).style.borderColor = 'var(--error)';
        } else {
            items.push({
                id: generateId(),
                name: name,
                spec: specVal,
                currentAmount: parseFloat(startVal),
                startAmount: parseFloat(startVal),
                lastUpdated: Date.now()
            });
        }
    });

    if (!isValid) {
        alert("请确保所有项目都已填写入有效的起始量。");
        return;
    }

    db.saveItems(items);
    db.setInit();
    renderMainView();
}

// Main View Logic
function enterBatchMode() {
    isBatchMode = true;
    selectedItemsId = [];
    renderMainView();
}

function exitBatchMode() {
    isBatchMode = false;
    selectedItemsId = [];
    renderMainView();
}

function toggleSelectAll() {
    const items = db.getItems();
    if (selectedItemsId.length === items.length) {
        selectedItemsId = [];
    } else {
        selectedItemsId = items.map(i => i.id);
    }
    renderMainView();
}

function handleItemClick(id) {
    if (isBatchMode) {
        if (selectedItemsId.includes(id)) {
            selectedItemsId = selectedItemsId.filter(sid => sid !== id);
        } else {
            selectedItemsId.push(id);
        }
        renderMainView();
    } else {
        openSingleCalcModal(id);
    }
}

// Long Press Logic for Edit
function attachLongPressEvents() {
    const cards = document.querySelectorAll('.item-card');
    cards.forEach(card => {
        let pressTimer;
        const start = () => {
            pressTimer = setTimeout(() => {
                const id = card.getAttribute('data-id');
                showContextMenu(id);
            }, 600); 
        };
        const cancel = () => clearTimeout(pressTimer);
        
        card.addEventListener('mousedown', start);
        card.addEventListener('touchstart', start);
        card.addEventListener('mouseup', cancel);
        card.addEventListener('mouseleave', cancel);
        card.addEventListener('touchend', cancel);
    });
}

function showContextMenu(id) {
    if(confirm("编辑该项目的规格信息？")) {
        const items = db.getItems();
        const item = items.find(i => i.id === id);
        const newSpec = prompt(`编辑 ${item.name} 的规格`, item.spec);
        if (newSpec !== null) {
            item.spec = newSpec;
            db.saveItems(items);
            renderMainView();
        }
    }
}

// Single Calculation Modal
function openSingleCalcModal(id) {
    const items = db.getItems();
    const item = items.find(i => i.id === id);
    
    showModal(`
        <h3>${item.name}</h3>
        <p style="color:#666; margin-bottom:10px;">${item.spec || ''}</p>
        <div style="background:#f3f4f6; padding:10px; border-radius:6px; margin-bottom:15px;">
            当前库存: <strong>${item.currentAmount}</strong>
        </div>
        <div style="margin-bottom:15px;">
            <label>本次用量</label>
            <input type="number" id="single-usage" placeholder="输入消耗数量" autofocus>
        </div>
        <div style="text-align:right;">
            <button class="btn btn-outline" onclick="closeModal()">取消</button>
            <button class="btn" onclick="executeSingleCalc('${id}')">确认计算</button>
        </div>
    `);
    
    setTimeout(() => document.getElementById('single-usage').focus(), 100);
}

function executeSingleCalc(id) {
    const usageInput = document.getElementById('single-usage');
    const usage = parseFloat(usageInput.value);
    
    if (isNaN(usage) || usage < 0) {
        alert("请输入有效的数值");
        return;
    }

    const items = db.getItems();
    const item = items.find(i => i.id === id);
    
    const before = item.currentAmount;
    item.currentAmount = before - usage;
    item.lastUpdated = Date.now();
    
    // Save Item
    db.saveItems(items);
    
    // Save Record
    db.addRecord({
        itemId: id,
        usage: usage,
        before: before,
        after: item.currentAmount,
        timestamp: Date.now(),
        type: 'single'
    });

    closeModal();
    renderMainView();
    showToast(`计算成功！剩余: ${item.currentAmount}`);
}

// Batch Calculation Wizard
function startBatchCalc() {
    if (selectedItemsId.length === 0) {
        alert("请先选择项目");
        return;
    }
    
    let currentStep = 0;
    const items = db.getItems().filter(i => selectedItemsId.includes(i.id));
    const usages = {};

    function renderStep() {
        const item = items[currentStep];
        showModal(`
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <h3>批量计算 (${currentStep + 1}/${items.length})</h3>
            </div>
            <div style="margin: 20px 0;">
                <h2 style="color:var(--primary)">${item.name}</h2>
                <p>当前: ${item.currentAmount}</p>
                <input type="number" id="batch-input-${item.id}" placeholder="输入用量" value="${usages[item.id] || ''}" style="margin-top:10px; font-size:1.2em;">
            </div>
            <div style="display:flex; justify-content:space-between;">
                <button class="btn btn-outline" onclick="prevBatchStep()">上一个</button>
                <div style="flex:1"></div>
                <button class="btn" onclick="nextBatchStep()">${currentStep === items.length - 1 ? '完成汇总' : '下一个'}</button>
            </div>
        `);
        setTimeout(() => document.getElementById(`batch-input-${item.id}`).focus(), 100);
    }

    window.prevBatchStep = () => {
        if (currentStep > 0) {
            saveCurrentInput();
            currentStep--;
            renderStep();
        }
    };

    window.nextBatchStep = () => {
        saveCurrentInput();
        if (currentStep < items.length - 1) {
            currentStep++;
            renderStep();
        } else {
            showBatchSummary();
        }
    };

    function saveCurrentInput() {
        const item = items[currentStep];
        const val = document.getElementById(`batch-input-${item.id}`).value;
        usages[item.id] = val ? parseFloat(val) : 0;
    }

    function showBatchSummary() {
        let html = `<h3>确认批量计算</h3><div style="max-height:300px; overflow-y:auto; margin:15px 0;">`;
        items.forEach(item => {
            const u = usages[item.id];
            if(u > 0) {
                html += `<div style="display:flex; justify-content:space-between; border-bottom:1px solid #eee; padding:5px 0;">
                    <span>${item.name}</span>
                    <span>-${u}</span>
                </div>`;
            }
        });
        html += `</div>
        <div style="text-align:right">
            <button class="btn btn-outline" onclick="closeModal()">取消</button>
            <button class="btn btn-success" onclick="executeBatchCalc()">确认更新</button>
        </div>`;
        showModal(html);
        
        window.executeBatchCalc = () => {
            const allItems = db.getItems();
            items.forEach(target => {
                const u = usages[target.id] || 0;
                if(u > 0) {
                    const dbItem = allItems.find(i => i.id === target.id);
                    const before = dbItem.currentAmount;
                    dbItem.currentAmount -= u;
                    dbItem.lastUpdated = Date.now();
                    
                    db.addRecord({
                        itemId: target.id,
                        usage: u,
                        before: before,
                        after: dbItem.currentAmount,
                        timestamp: Date.now(),
                        type: 'batch'
                    });
                }
            });
            db.saveItems(allItems);
            closeModal();
            exitBatchMode();
            showToast("批量更新完成");
        };
    }

    renderStep();
}

// Excel Import Logic (SheetJS)
function handleExcelImport(input) {
    const file = input.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, {header: 1}); // Array of Arrays

        const matchedUpdates = [];
        const notFound = [];

        // Skip header row if necessary, assuming strict match searching
        // Iterate all rows
        jsonData.forEach(row => {
            if(row.length < 2) return;
            const name = typeof row[0] === 'string' ? row[0].trim() : String(row[0]);
            const usage = parseFloat(row[1]);

            if (ITEMS_NAMES.includes(name) && !isNaN(usage)) {
                matchedUpdates.push({ name, usage });
            } else if (name) {
                if (!ITEMS_NAMES.includes(name)) notFound.push(name);
            }
        });

        // Show Confirmation
        let html = `<h3>导入结果确认</h3>`;
        if (matchedUpdates.length > 0) {
            html += `<p style="color:var(--success)">成功匹配 ${matchedUpdates.length} 项</p>`;
            html += `<ul style="max-height:200px; overflow-y:auto; font-size:0.9em; margin:10px 0;">`;
            matchedUpdates.forEach(m => {
                html += `<li>${m.name}: <span style="font-weight:bold">${m.usage}</span></li>`;
            });
            html += `</ul>`;
            
            if (notFound.length > 0) {
                html += `<div style="color:var(--warning); font-size:0.8em; margin-top:5px;">警告：${notFound.length} 项未找到匹配 (示例: ${notFound[0]})</div>`;
            }

            html += `<div style="text-align:right; margin-top:15px;">
                <button class="btn" onclick="applyExcelUpdates()">确认更新数据</button>
            </div>`;

            // Store temp for execution
            window.tempUpdates = matchedUpdates;
        } else {
            html += `<p style="color:var(--error)">未找到匹配的项目名称。</p>`;
        }
        
        showModal(html);
        input.value = ''; // Reset
    };
    reader.readAsArrayBuffer(file);
}

window.applyExcelUpdates = () => {
    const updates = window.tempUpdates;
    const allItems = db.getItems();
    let count = 0;

    updates.forEach(up => {
        const item = allItems.find(i => i.name === up.name);
        if (item) {
            const before = item.currentAmount;
            item.currentAmount -= up.usage; // Assuming excel contains Usage data
            item.lastUpdated = Date.now();
            
            db.addRecord({
                itemId: item.id,
                usage: up.usage,
                before: before,
                after: item.currentAmount,
                timestamp: Date.now(),
                type: 'import'
            });
            count++;
        }
    });

    db.saveItems(allItems);
    closeModal();
    renderMainView();
    showToast(`成功从Excel更新 ${count} 项数据`);
};

// Records Viewer Logic
function showRecords() {
    renderRecordsView();
}

function filterRecords(keyword) {
    const list = document.getElementById('records-list');
    const records = db.getRecords();
    const items = db.getItems();
    
    const filtered = records.filter(r => {
        const item = items.find(i => i.id === r.itemId);
        return item.name.toLowerCase().includes(keyword.toLowerCase());
    });

    list.innerHTML = filtered.map(r => {
        const item = items.find(i => i.id === r.itemId);
        return `
            <div class="record-item">
                <div class="record-header">
                    <span>${item.name}</span>
                    <span class="diff-negative">-${r.usage}</span>
                </div>
                <div class="record-meta">
                    ${formatDate(r.timestamp)} | 结果: ${r.after}
                </div>
            </div>
        `;
    }).join('');
}

// --- Utilities ---
function generateId() {
    return '_' + Math.random().toString(36).substr(2, 9);
}

function formatDate(ts) {
    const d = new Date(ts);
    return `${d.getMonth()+1}/${d.getDate()} ${d.getHours()}:${d.getMinutes().toString().padStart(2,'0')}`;
}

function showModal(html) {
    modalContent.innerHTML = html;
    modalOverlay.classList.remove('hidden');
}

function closeModal() {
    modalOverlay.classList.add('hidden');
}

function showToast(msg) {
    // Simple toast implementation
    const div = document.createElement('div');
    div.innerText = msg;
    div.style.position = 'fixed';
    div.style.bottom = '20px';
    div.style.left = '50%';
    div.style.transform = 'translateX(-50%)';
    div.style.background = 'rgba(0,0,0,0.8)';
    div.style.color = 'white';
    div.style.padding = '10px 20px';
    div.style.borderRadius = '20px';
    div.style.zIndex = '999';
    document.body.appendChild(div);
    setTimeout(() => div.remove(), 2000);
}

// Close Modal on overlay click
modalOverlay.addEventListener('click', (e) => {
    if(e.target === modalOverlay) closeModal();
});

// Start App
init();