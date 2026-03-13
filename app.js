// 防窺機制
document.addEventListener('contextmenu', event => event.preventDefault());
document.addEventListener('keydown', event => {
    if (event.key === 'F12' || (event.ctrlKey && event.shiftKey && event.key === 'I') || (event.ctrlKey && event.key === 'u')) {
        event.preventDefault();
    }
});

const tablesConfig = [
    { id: '1-1', name: '1-1. 工位檔 (BLK_Supply)', hasHeader: true }, 
    { id: '1-2', name: '1-2. Probe檔 (Probe)', hasHeader: true },
    { id: '1-3', name: '1-3. Nikon量測 (Nikon)', hasHeader: true },
    { id: '1-4', name: '1-4. AOI量測 (AOI_Dim)', hasHeader: true },
    { id: '1-5', name: '1-5. AOI Defect', hasHeader: true },
    { id: '1-6', name: '1-6. AOI Defect Code', hasHeader: true },
    { id: '1-7', name: '1-7. WCM 剝斷檔', hasHeader: true }
];

let configStore = { currentPN: "8A76810003", pns: {} };
let dataStore = {}; 
window.activeMapRects = []; 
let currentFilterInfo = { tabId: null, colIdx: null, filterKey: 'filters', set: new Set() };

function showLoading(msg) { document.getElementById('loading-text').innerText = msg; document.getElementById('loading-mask').style.display = 'flex'; }
function hideLoading() { document.getElementById('loading-mask').style.display = 'none'; }
function toggleGroup(gid) { let grp = document.getElementById(`group-${gid}`); let icon = document.getElementById(`icon-${gid}`); if (grp.classList.contains('open')) { grp.classList.remove('open'); icon.innerText = '▶'; } else { grp.classList.add('open'); icon.innerText = '▼'; } }

function resetDataStore(tabId) { 
    dataStore[tabId] = { rawFiles: [], mergedData: [], pivotedData: [], pivotedHeaders: [], headers: [], mapGridData: [], matchedFiles: [], filters: {}, filters_sec: {}, sheetNamesExtracted: false, snakeGridMap: null, combinedData1_2: null, combinedHeaders1_2: null }; 
}

function init() {
    tablesConfig.forEach(t => resetDataStore(t.id));
    dataStore['phase3'] = { mergedData: [], headers: [], filters: {} };

    let saved = localStorage.getItem('eda_mapper_config_V35') || localStorage.getItem('eda_mapper_config_V34') || localStorage.getItem('eda_mapper_config_V32');
    if (saved) { try { configStore = JSON.parse(saved); } catch (e) {} }
    if (!configStore.pns || !configStore.pns[configStore.currentPN]) createNewPN(configStore.currentPN, null);
    
    updatePNDropdown(); buildTabs();
    tablesConfig.forEach(t => {
        let conf = configStore.pns[configStore.currentPN].settings[t.id];
        if (conf.saveRaw) {
            let savedRaw = localStorage.getItem(`eda_raw_${configStore.currentPN}_${t.id}`);
            if (savedRaw) { 
                try { 
                    dataStore[t.id].rawFiles = JSON.parse(savedRaw); 
                    calculateMergedData(t.id); 
                } catch(e) {} 
            }
        }
    });
    switchTab('1-1'); 
    document.addEventListener('mousemove', handleTooltip);
    document.addEventListener('click', (e) => { if(!e.target.closest('#filter-modal') && !e.target.closest('.filter-icon')) document.getElementById('filter-modal').style.display = 'none'; });
}

function createNewPN(pn, copyFromPN) {
    let settings = { phase3: { transforms: {}, legConfig: {} } };
    if (copyFromPN && configStore.pns[copyFromPN]) { settings = JSON.parse(JSON.stringify(configStore.pns[copyFromPN].settings)); } 
    else { tablesConfig.forEach(t => { settings[t.id] = { filePattern: t.id === '1-7' ? "*.csv, *.txt" : "*.csv", sheetName: "", headerRowIndex: null, encoding: "auto", saveRaw: false, derivedFields: [], customHeaders: {}, pivotConfig: {enabled:false, keyCol:'', valCol:''}, mapConfig: {}, legConfig: {}, lastFolderPath: "" }; }); }
    configStore.pns[pn] = { settings: settings }; tablesConfig.forEach(t => resetDataStore(t.id));
    dataStore['phase3'] = { mergedData: [], headers: [], filters: {} };
}

function updatePNDropdown() { let sel = document.getElementById('pn-select'); sel.innerHTML = ''; Object.keys(configStore.pns).forEach(pn => sel.appendChild(new Option(pn, pn))); sel.value = configStore.currentPN; }
function changePN() { 
    configStore.currentPN = document.getElementById('pn-select').value; 
    tablesConfig.forEach(t => { 
        resetDataStore(t.id); 
        let countSpan = document.getElementById(`file-count-${t.id}`); 
        if(countSpan) countSpan.innerText = ''; 
        let conf = configStore.pns[configStore.currentPN].settings[t.id]; 
        if (conf.saveRaw) { 
            let savedRaw = localStorage.getItem(`eda_raw_${configStore.currentPN}_${t.id}`); 
            if (savedRaw) { 
                try { 
                    dataStore[t.id].rawFiles = JSON.parse(savedRaw); 
                    calculateMergedData(t.id);
                } catch(e) {} 
            } 
        } 
    }); 
    dataStore['phase3'] = { mergedData: [], headers: [], filters: {} }; 
    switchTab(document.querySelector('.tab-btn.active').dataset.id); 
}
function addPN() { let newPN = prompt("請輸入新的 PN 名稱："); if (!newPN || newPN.trim() === "") return; if (configStore.pns[newPN]) { alert("此 PN 已存在！"); return; } createNewPN(newPN.trim(), configStore.currentPN); configStore.currentPN = newPN.trim(); saveGlobalConfig(false); updatePNDropdown(); changePN(); }
function deletePN() { if (Object.keys(configStore.pns).length <= 1) { alert("至少需保留一個 PN！"); return; } if (confirm(`確定要刪除 [${configStore.currentPN}] 嗎？`)) { tablesConfig.forEach(t => localStorage.removeItem(`eda_raw_${configStore.currentPN}_${t.id}`)); delete configStore.pns[configStore.currentPN]; configStore.currentPN = Object.keys(configStore.pns)[0]; saveGlobalConfig(false); updatePNDropdown(); changePN(); } }

function saveGlobalConfig(showAlert) {
    localStorage.setItem('eda_mapper_config_V35', JSON.stringify(configStore));
    tablesConfig.forEach(t => {
        let conf = configStore.pns[configStore.currentPN].settings[t.id]; let key = `eda_raw_${configStore.currentPN}_${t.id}`;
        if (conf.saveRaw && dataStore[t.id].rawFiles.length > 0) { try { localStorage.setItem(key, JSON.stringify(dataStore[t.id].rawFiles)); } catch (e) { alert(`[${t.name}] 資料過大無法儲存本機。`); conf.saveRaw = false; document.getElementById(`save-raw-${t.id}`).checked = false; } } else { localStorage.removeItem(key); }
    });
    if (showAlert) alert("✅ 設定與資料已儲存本機！");
}

// 💡 更新：使用 File System Access API 彈出 Windows 存檔視窗，並匯出所有 PN 的 Global Config
async function exportConfig() {
    saveGlobalConfig(false); // 確保最新的變更都有存入 configStore
    
    // configStore 本身就包含了所有的 PN 設定 (configStore.pns)，所以直接匯出 configStore 即可
    const dataStr = JSON.stringify(configStore, null, 2);
    const defaultFileName = 'EDA_Mapper_Global_Config.json';

    try {
        // 檢查瀏覽器是否支援跳出檔案選擇視窗 API (Chrome/Edge 支援)
        if (window.showSaveFilePicker) {
            const handle = await window.showSaveFilePicker({
                suggestedName: defaultFileName,
                types: [{
                    description: 'JSON 設定檔',
                    accept: { 'application/json': ['.json'] },
                }],
            });
            const writable = await handle.createWritable();
            await writable.write(dataStr);
            await writable.close();
            alert("✅ 全局設定檔匯出成功！(包含所有 PN)");
        } else {
            // 備案機制：若不支援該 API，退回傳統的下載方式
            let a = document.createElement('a');
            let blob = new Blob([dataStr], { type: 'application/json' });
            a.href = URL.createObjectURL(blob);
            a.download = defaultFileName;
            a.click();
            URL.revokeObjectURL(a.href);
            alert("✅ 全局設定檔已下載！(包含所有 PN)");
        }
    } catch (err) {
        // 使用者按下取消時不報錯
        if (err.name !== 'AbortError') {
            console.error('匯出發生錯誤:', err);
            alert("❌ 匯出失敗或發生錯誤！");
        }
    }
}

function importConfig(event) {
    let file = event.target.files[0];
    if (!file) return;
    let reader = new FileReader();
    reader.onload = function(e) {
        try {
            let imported = JSON.parse(e.target.result);
            if (imported && imported.pns) {
                // 顯示三個按鈕的自訂視窗
                let modal = document.getElementById('import-modal');
                modal.style.display = 'flex';
                
                // 完全覆蓋按鈕事件
                document.getElementById('btn-import-overwrite').onclick = function() {
                    configStore = imported;
                    finalizeImport();
                };
                
                // 合併新增按鈕事件
                document.getElementById('btn-import-merge').onclick = function() {
                    Object.keys(imported.pns).forEach(pn => {
                        configStore.pns[pn] = imported.pns[pn];
                    });
                    if (!configStore.pns[configStore.currentPN]) {
                        configStore.currentPN = Object.keys(configStore.pns)[0];
                    }
                    finalizeImport();
                };
                
                // 共用的完成處理
                function finalizeImport() {
                    saveGlobalConfig(false);
                    updatePNDropdown();
                    changePN();
                    modal.style.display = 'none';
                    alert("✅ 設定檔匯入成功！");
                }
            } else { 
                alert("❌ 格式不符，無法匯入！"); 
            }
        } catch (err) { 
            alert("❌ 解析設定檔失敗！"); 
        }
    };
    reader.readAsText(file);
    event.target.value = '';
}

function buildTabs() {
    let groupP1 = document.getElementById('group-p1'); let groupP2 = document.getElementById('group-p2'); let groupP3 = document.getElementById('group-p3');
    let contentArea = document.getElementById('content-area');
    
    tablesConfig.forEach(t => {
        let btn1 = document.createElement('button'); btn1.className = 'tab-btn p1-btn'; btn1.dataset.id = t.id; btn1.innerText = t.name; btn1.onclick = () => switchTab(t.id); groupP1.appendChild(btn1);

        let content1 = document.createElement('div'); content1.className = 'tab-content'; content1.id = `tab-${t.id}`;
        let headerMessage = `<div style="color:#d9534f; font-weight:bold; background:#ffeeba; padding:10px; border-radius:5px; border-left:5px solid #ffc107; grid-column:1 / span 2;">💡 提示：系統會自動偵測最佳標題列，您也可以在下方「原始檔案預覽」中，點擊任一列重新指定標題列/解析起始列。解析後資料標題支援「雙擊 ✏️ 修改」。</div>`;

        let pivotUI = '';
        if (t.id === '1-3') {
            pivotUI = `
            <div class="derived-section" style="background:#e8f8f5; border-color:#a2d9ce;">
                <strong style="color:#117a65; font-size:16px;">🔄 1-3 專屬：直表轉橫表 (Pivot) 引擎</strong><br>
                <div style="margin-top:8px; font-size:13px; color:#c0392b; font-weight:bold;">⚠️ 防呆注意：若需重新載入檔案，請先「取消勾選」此功能，待重新定義好標題與 WaferID 後再勾選！</div>
                <div style="margin-top:5px; font-size:13px; color:#555;">將指定的「項目」轉為橫向標題，並自動依 WaferID 賦予遞增流水號。</div>
                <div style="margin-top:10px; display:flex; gap:15px; align-items:center;">
                    <label style="font-weight:bold; color:#117a65;"><input type="checkbox" id="pivot-enable-${t.id}" onchange="updatePivotConfig('${t.id}')"> 啟用轉置功能</label>
                    <label>項目名稱(Key)欄位: <select id="pivot-key-${t.id}" onchange="updatePivotConfig('${t.id}')"></select></label>
                    <label>數值(Value)欄位: <select id="pivot-val-${t.id}" onchange="updatePivotConfig('${t.id}')"></select></label>
                </div>
            </div>`;
        }

        content1.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <h2 style="margin-top:0; color:#005a9e;">${t.name} (匯入與解析)</h2>
                <label style="cursor:pointer; font-weight:bold; background:#fff3cd; padding:5px 10px; border-radius:4px; border:1px solid #ffeeba; color:#856404;"><input type="checkbox" id="save-raw-${t.id}" onchange="updateTableConfig('${t.id}', 'saveRaw', this.checked)"> 💾 將此表格 RawData 存於本機</label>
            </div>
            <div class="import-section">
                <div class="form-group"><label>🔎 檔名過濾規則:</label><input type="text" id="pattern-${t.id}" onchange="updateTableConfig('${t.id}', 'filePattern', this.value); checkExcelLogic('${t.id}');"></div>
                <div class="form-group"><label>🔤 檔案編碼:</label><select id="encoding-${t.id}" onchange="updateTableConfig('${t.id}', 'encoding', this.value)"><option value="auto">自動偵測 (Big5/UTF-8)</option><option value="big5">Big5 (台灣機台)</option><option value="utf-8">UTF-8 (國際標準)</option><option value="shift-jis">Shift-JIS (日系機台)</option></select></div>
                <div class="form-group"><label>📊 指定工作表:</label><select id="sheet-${t.id}" onchange="changeSheet('${t.id}', this.value)" disabled><option value="">-- 自動讀取 --</option></select></div>
                ${headerMessage}
                <div class="form-group full-width" style="margin-top:5px; border-top:1px dashed #ccc; padding-top:15px; display:flex; align-items:center; gap: 10px;">
                    <label for="folder-input-${t.id}" class="btn btn-primary" style="margin:0; text-align:center; min-width:120px;">📂 選擇資料夾</label>
                    <input type="file" id="folder-input-${t.id}" webkitdirectory directory multiple style="display:none;" onchange="handleFolderSelect(event, '${t.id}')">
                    
                    <label style="margin-left:15px; color:#666; font-size:13px; font-weight:bold;">📁 路徑備忘錄:</label>
                    <input type="text" id="folder-path-${t.id}" placeholder="可貼上絕對路徑備忘..." style="flex:1; max-width:350px;" onchange="updateTableConfig('${t.id}', 'lastFolderPath', this.value)">
                    
                    <span id="file-count-${t.id}" style="color: #28a745; font-weight: bold; margin-left:auto;"></span>
                </div>
                <div class="full-width"><div style="color:#888; font-size:13px; margin-top:5px;">或直接貼上單一檔案資料 (可直接Ctrl+V)</div><textarea class="paste-area" placeholder="貼上資料..." oninput="handlePaste(event, '${t.id}')"></textarea></div>
            </div>
            
            <div class="derived-section" id="derived-section-${t.id}">
                <div style="display:flex; justify-content:space-between; margin-bottom: 10px;"><strong style="color:#005a9e; font-size: 16px;">🛠️ 新增運算邏輯 (如 WaferID 擷取)</strong><button class="btn btn-primary" style="padding: 2px 8px; font-size:12px;" onclick="addDerivedField('${t.id}')">➕ 新增運算邏輯</button></div>
                <div class="help-box">
                    <strong>💡 運算邏輯函數說明：</strong><br>
                    <ul style="margin: 5px 0; padding-left: 20px;">
                        <li><b>字串擷取 (Mid)</b>：從第 N 個字元開始擷取。(例如: 起始 1、長度 5，將從 'Wafer123' 擷取出 'Wafer')</li>
                        <li><b>字串分割 (Split)</b>：用符號切開並取第 N 段。(例如: 分隔符 '_'，索引 1，將從 'A_B_C' 擷取出 'A') <i>*註: 索引從 1 開始</i></li>
                        <li><b>正則提取 (Regex)</b>：進階用法，輸入正則表達式提取字串。(例如: 輸入 \\d+ 將抓取數字)</li>
                    </ul>
                </div>
                <div id="derived-rules-${t.id}"></div>
            </div>
            
            ${pivotUI}
            <div id="summary-${t.id}" style="font-weight: bold; margin-bottom: 10px; color: #005a9e;"></div>
            <div id="preview-${t.id}"><p style="padding: 10px; color: #666;">尚未匯入資料...</p></div>
        `;
        contentArea.appendChild(content1);

        // --- Phase 2 ---
        let btn2 = document.createElement('button'); btn2.className = 'tab-btn p2-btn'; btn2.dataset.id = `map-${t.id}`; btn2.innerText = t.name; btn2.onclick = () => switchTab(`map-${t.id}`); groupP2.appendChild(btn2);
        let content2 = document.createElement('div'); content2.className = 'tab-content'; content2.id = `tab-map-${t.id}`;
        content2.innerHTML = `
            <h2 style="color:#d35400;">🗺️ ${t.name} - MAP 關聯與設定</h2>
            <div id="map-warning-${t.id}" class="map-warning" style="display:none;"></div>
            <div class="map-controls" id="map-controls-config-${t.id}" style="display:none; background:#d4e6f1;"></div>
            <div class="map-controls" id="map-controls-legend-${t.id}" style="display:none; align-items:flex-start;"></div>
            <div class="map-canvas-area" id="map-canvas-area-${t.id}"><p style="color:#aaa; font-size:18px;">請上方設定參數後繪製 MAP</p></div>
        `;
        contentArea.appendChild(content2);
    });

    // --- Phase 3 ---
    let btn3 = document.createElement('button'); btn3.className = 'tab-btn p3-btn'; btn3.dataset.id = 'phase3'; btn3.innerHTML = '✨ 執行多表融合 (Master Join)'; btn3.onclick = () => switchTab('phase3'); groupP3.appendChild(btn3);
    let content3 = document.createElement('div'); content3.className = 'tab-content'; content3.id = `tab-phase3`;
    content3.innerHTML = `<h2 style="color:#1e8449;">🚀 階段三：多表融合與匯出</h2>
        <div id="p3-controls" style="background:#e9f7ef; padding:15px; border:1px solid #c3e6cb; border-radius:5px; margin-bottom:10px;"></div>
        <div id="p3-content"></div>`;
    contentArea.appendChild(content3);
}

function checkExcelLogic(tabId) { document.getElementById(`sheet-${tabId}`).disabled = !document.getElementById(`pattern-${tabId}`).value.toLowerCase().includes('xls'); }

function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active')); document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    document.querySelector(`.tab-btn[data-id="${tabId}"]`).classList.add('active'); document.getElementById(`tab-${tabId}`).classList.add('active');
    
    if (tabId === 'phase3') { buildPhase3Controls(); return; }
    if (tabId.startsWith('map-')) { buildPhase2Controls(tabId.replace('map-', '')); return; }

    let conf = configStore.pns[configStore.currentPN].settings[tabId]; if(!conf.customHeaders) conf.customHeaders = {}; if(!conf.pivotConfig) conf.pivotConfig = {enabled:false, keyCol:'', valCol:''};
    document.getElementById(`pattern-${tabId}`).value = conf.filePattern || '*.csv'; document.getElementById(`encoding-${tabId}`).value = conf.encoding || 'auto'; document.getElementById(`save-raw-${tabId}`).checked = conf.saveRaw || false;
    let pathInput = document.getElementById(`folder-path-${tabId}`); if(pathInput) pathInput.value = conf.lastFolderPath || '';
    checkExcelLogic(tabId);
    let sheetSel = document.getElementById(`sheet-${tabId}`); if (conf.sheetName && sheetSel.querySelector(`option[value="${conf.sheetName}"]`)) sheetSel.value = conf.sheetName;
    
    renderDerivedRules(tabId); if(dataStore[tabId].rawFiles.length > 0) { calculateMergedData(tabId); renderPreview(tabId); }
}

function updateTableConfig(tabId, key, value) { configStore.pns[configStore.currentPN].settings[tabId][key] = value; saveGlobalConfig(false); }

function updateMapConfig(tabId, key, value) {
    let conf = configStore.pns[configStore.currentPN].settings[tabId];
    if(!conf.mapConfig) conf.mapConfig = {};
    conf.mapConfig[key] = value;
    saveGlobalConfig(false);
    if(tabId === '1-2' && document.getElementById(`cvs-1-2`)) drawPhase2Map('1-2'); 
}

window.normKey = function(k) { 
    let s = String(k).trim().toUpperCase(); 
    if (!isNaN(s) && s !== '') return String(Number(s)); 
    return s; 
};

window.getCatColor = function(val) {
    let v = String(val).toUpperCase().trim();
    if(['OK', '0', 'PASS'].includes(v)) return '#28a745'; 
    if(['NG', '1', 'FAIL'].includes(v)) return '#dc3545'; 
    return null;
};

window.calculateStats = function(arr) {
    if (!arr || arr.length === 0) return null;
    let sorted = [...arr].filter(v => !isNaN(v)).sort((a,b) => a-b);
    if (sorted.length === 0) return null;
    let min = sorted[0]; let max = sorted[sorted.length - 1];
    let q1 = sorted[Math.floor(sorted.length * 0.25)]; let q2 = sorted[Math.floor(sorted.length * 0.5)]; let q3 = sorted[Math.floor(sorted.length * 0.75)];
    let iqr = q3 - q1;
    return { min, max, q1, q2, q3, iqr, lowerBound: q1 - 1.5 * iqr, upperBound: q3 + 1.5 * iqr };
};

function buildPhase2Controls(tabId) {
    let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId];
    if (!conf.mapConfig) conf.mapConfig = {}; let mc = conf.mapConfig;
    if (!conf.legConfig) conf.legConfig = {}; 
    
    let warningDiv = document.getElementById(`map-warning-${tabId}`); let configDiv = document.getElementById(`map-controls-config-${tabId}`);
    let legendDiv = document.getElementById(`map-controls-legend-${tabId}`); let canvasArea = document.getElementById(`map-canvas-area-${tabId}`);
    
    if (store.headers.length === 0 && tabId !== '1-1') { warningDiv.style.display = 'block'; configDiv.style.display = 'none'; legendDiv.style.display = 'none'; canvasArea.style.display = 'none'; warningDiv.innerHTML = "⚠️ 尚未匯入資料！請先在階段一匯入檔案並解析。"; return; }
    
    let baseHeaders = store.headers; let baseData = store.mergedData;
    if (tabId === '1-3' && conf.pivotConfig.enabled) { baseHeaders = store.pivotedHeaders; baseData = store.pivotedData; }
    if (tabId === '1-5' && store.joinedHeaders && store.joinedData) { baseHeaders = store.joinedHeaders; baseData = store.joinedData; }
    if (tabId === '1-2' && mc.isCombined && store.combinedData1_2) { baseData = store.combinedData1_2; baseHeaders = store.combinedHeaders1_2 || store.headers; }

    let waferIdx = baseHeaders.findIndex(h => h && h.toUpperCase().replace(/[\s_]/g, '') === 'WAFERID');
    let hasWaferID = waferIdx > -1;
    
    let waferSelectHTML = '';
    if (hasWaferID) {
        let uWafers = [...new Set(baseData.map(r => String(r[waferIdx]).trim()))].filter(v => v);
        if(!mc.selectedWafer || !uWafers.includes(mc.selectedWafer)) mc.selectedWafer = uWafers[0];
        let wOpts = uWafers.map(w => `<option value="${w}" ${mc.selectedWafer===w?'selected':''}>${w}</option>`).join('');
        waferSelectHTML = `<label style="margin-left:15px; color:#d35400;">📍 WaferID:</label><select id="map-wafer-${tabId}" onchange="updateMapConfig('${tabId}', 'selectedWafer', this.value); drawPhase2Map('${tabId}')" style="width:120px;">${wOpts}</select>`;
    }

    let getOpts = (type, savedVal, customHeaders) => {
        let headersToUse = customHeaders || baseHeaders;
        return headersToUse.map(h => {
            let isSel = false; let hu = h.toUpperCase();
            if(savedVal && h === savedVal) { isSel = true; } 
            else if(!savedVal) {
                if(type==='X' && (hu==='X' || hu==='X_COORD' || hu==='X座標' || hu==='MAP_X' || hu==='GLOBAL_C')) isSel = true;
                if(type==='Y' && (hu==='Y' || hu==='Y_COORD' || hu==='Y座標' || hu==='MAP_Y' || hu==='GLOBAL_R')) isSel = true;
                if(type==='VAL' && (hu==='VALUE' || hu==='DATA' || hu==='數值' || hu==='STATUS' || hu==='NEW_NO')) isSel = true;
            }
            return `<option value="${h}" ${isSel?'selected':''}>${h}</option>`;
        }).join('');
    };

    warningDiv.style.display = 'none'; configDiv.style.display = 'flex'; canvasArea.style.display = 'flex'; legendDiv.style.display = 'none';

    let dirHTML = `
        <label style="margin-left:10px;">X方向:</label><select id="map-dir-x-${tabId}" style="width:70px;" onchange="updateMapConfig('${tabId}', 'dirX', this.value); ${tabId!=='1-2'||mc.isCombined?'drawPhase2Map(\''+tabId+'\')':''}"><option value="LR" ${mc.dirX==='LR'?'selected':''}>L->R</option><option value="RL" ${mc.dirX==='RL'?'selected':''}>R->L</option></select>
        <label>Y方向:</label><select id="map-dir-y-${tabId}" style="width:70px;" onchange="updateMapConfig('${tabId}', 'dirY', this.value); ${tabId!=='1-2'||mc.isCombined?'drawPhase2Map(\''+tabId+'\')':''}"><option value="DU" ${mc.dirY==='DU'?'selected':''}>D->U</option><option value="UD" ${mc.dirY==='UD'?'selected':''}>U->D</option></select>
    `;

    let valChangeJS = `updateMapConfig('${tabId}', 'valCol', this.value); drawPhase2Map('${tabId}');`;

    if (tabId === '1-6') {
        let store5 = dataStore['1-5']; let h5 = (store5 && store5.headers.length>0) ? store5.headers : ['無1-5資料'];
        configDiv.innerHTML = `<strong>🔗 1-6 關聯設定 (將 Left Join 至 1-5)</strong> <label style="margin-left:15px;">1-6 關聯欄位:</label><select id="join-1-6" onchange="updateMapConfig('1-6', 'joinCol', this.value)">${getOpts('', mc.joinCol)}</select> <label>對應 1-5 欄位:</label><select id="join-1-5" onchange="updateMapConfig('1-6', 'joinCol5', this.value)">${getOpts('', mc.joinCol5, h5)}</select> <button class="btn btn-primary" style="margin-left:auto;" onclick="renderLeftJoinPreview()">執行關聯預覽</button>`;
        canvasArea.innerHTML = `<div id="map-render-${tabId}" style="width:100%; overflow:auto;"><p style="color:#aaa;">請點擊「執行關聯預覽」檢視 Left Join 結果</p></div>`;
        if(mc.joinCol && mc.joinCol5 && store5 && store5.mergedData.length > 0) setTimeout(()=>renderLeftJoinPreview(), 100);
    } else if (tabId === '1-1') {
        let gridHeaders = ['No.', 'Map_X', 'Map_Y', 'Row_Ref'];
        configDiv.innerHTML = `<strong>1-1 基礎 MAP 網格</strong><label style="margin-left:15px;">顯示資料:</label><select id="map-val-1-1" onchange="${valChangeJS}"><option value="">-- 無 --</option>${getOpts('VAL', mc.valCol, gridHeaders)}</select><button class="btn btn-primary" style="margin-left:auto;" onclick="drawPhase2Map('1-1')">繪製基礎 MAP</button>`;
        canvasArea.innerHTML = `<div id="map-render-1-1" style="width:100%; height:100%; display:flex; justify-content:center; align-items:center;"><p style="color:#aaa;">點擊「繪製」呈現 1-1 物理底層網格</p></div>`;
        legendDiv.innerHTML = `<div id="legend-content-${tabId}" style="display:flex; gap:10px;"></div>`;
    } else if (tabId === '1-7') {
        if (!hasWaferID) { warningDiv.style.display = 'block'; configDiv.style.display = 'none'; canvasArea.style.display = 'none'; warningDiv.innerHTML = "⚠️ 找不到 <b>WaferID</b>！"; return; }
        configDiv.innerHTML = `<strong>🔗 1-7 MAP (需綁定 1-1)</strong> ${waferSelectHTML} <label style="margin-left:15px;">顯示資料:</label><select id="map-val-${tabId}" onchange="${valChangeJS}">${getOpts('VAL', mc.valCol)}</select> ${dirHTML} <button class="btn btn-primary" style="margin-left:auto;" onclick="drawPhase2Map('${tabId}')">繪製 MAP</button>`;
        canvasArea.innerHTML = `<div id="map-render-${tabId}" style="width:100%; height:100%; display:flex; justify-content:center; align-items:center;"><p style="color:#aaa;">點擊「繪製 MAP」呈現剝斷檔分佈</p></div>`;
        legendDiv.innerHTML = `<div id="legend-content-${tabId}" style="display:flex; gap:10px;"></div>`;
    } else if (tabId === '1-2') {
        if (!hasWaferID) { warningDiv.style.display = 'block'; configDiv.style.display = 'none'; canvasArea.style.display = 'none'; warningDiv.innerHTML = "⚠️ 找不到 <b>WaferID</b>！"; return; }
        
        if (mc.isCombined) {
            configDiv.innerHTML = `
                ${waferSelectHTML}
                <label style="margin-left:15px;">顯示數值:</label><select id="map-val-${tabId}" onchange="${valChangeJS}">${getOpts('VAL', mc.valCol)}</select>
                ${dirHTML}
                <label style="margin-left:10px;">全局旋轉:</label><select onchange="updateMapConfig('${tabId}','rot',this.value); drawPhase2Map('${tabId}')"><option value="0" ${mc.rot==='0'?'selected':''}>0°</option><option value="90" ${mc.rot==='90'?'selected':''}>90°</option><option value="180" ${mc.rot==='180'?'selected':''}>180°</option><option value="270" ${mc.rot==='270'?'selected':''}>270°</option></select>
                <button class="btn btn-danger" style="margin-left:auto;" onclick="resetMap1_2()">🔙 返回四象限設定</button>
            `;
            canvasArea.innerHTML = `
                <div class="split-layout">
                    <div class="split-left" style="max-width: 250px; justify-content: flex-start;">
                        <h4 style="margin-top:0; color:#005a9e; font-size:14px;">已組合為絕對座標</h4>
                        <div style="font-size:12px; color:#555;">如需重新調整個別象限，請點擊上方「返回四象限設定」。</div>
                        <div id="preview-1-7-mini" style="margin-top:20px; width:100%; text-align:center;"></div>
                    </div>
                    <div class="split-right" id="map-render-1-2" style="border:none; background:#fdfdfd;"><p style="color:#aaa;">請點擊繪製組合後的大圖</p></div>
                </div>
            `;
            legendDiv.innerHTML = `<div id="legend-content-${tabId}" style="display:flex; gap:10px;"></div>`;
            setTimeout(() => { drawPhase2Map('1-2'); renderMini17(); }, 100);
        } else {
            configDiv.innerHTML = `
                ${waferSelectHTML}
                <label style="margin-left:15px;">X座標:</label><select id="map-x-${tabId}" onchange="updateMapConfig('${tabId}','xCol',this.value)">${getOpts('X', mc.xCol)}</select>
                <label>Y座標:</label><select id="map-y-${tabId}" onchange="updateMapConfig('${tabId}','yCol',this.value)">${getOpts('Y', mc.yCol)}</select>
                <label>數值:</label><select id="map-val-${tabId}" onchange="${valChangeJS}">${getOpts('VAL', mc.valCol)}</select>
                <label>📦 Area欄位:</label><select id="map-area-${tabId}" onchange="updateMapConfig('${tabId}','areaCol',this.value); buildPhase2Controls('${tabId}')">${getOpts('', mc.areaCol)}</select>
            `;
            
            let areaCol = document.getElementById(`map-area-${tabId}`).value; let areaIdx = baseHeaders.indexOf(areaCol);
            let uAreas = []; if(areaIdx>-1) uAreas = [...new Set(baseData.map(r=>String(r[areaIdx])))].filter(v=>v).sort();
            
            let savedZones = [mc.pos_zone_tl, mc.pos_zone_tr, mc.pos_zone_bl, mc.pos_zone_br];
            let a1, a2, a3, a4;
            if (savedZones.every(z => z) && new Set(savedZones).size === 4 && savedZones.every(z => uAreas.includes(z))) {
                [a1, a2, a3, a4] = savedZones;
            } else {
                a1 = uAreas[0] || 'Area_A'; a2 = uAreas.length > 1 ? uAreas[1] : 'Area_B';
                a3 = uAreas.length > 2 ? uAreas[2] : 'Area_C'; a4 = uAreas.length > 3 ? uAreas[3] : 'Area_D';
            }

            function getQuadItemHTML(areaName) {
                let dx = mc['dirX_' + areaName] || 'LR'; let dy = mc['dirY_' + areaName] || 'UD'; let rot = mc['rot_' + areaName] || '0';
                return `<div class="quad-item" draggable="true" data-val="${areaName}">
                    <strong style="font-size:14px; color:#f1c40f;">${areaName}</strong>
                    <div style="font-size:11px; margin-top:5px; color:#fff;">X: <select onchange="updateMapConfig('1-2','dirX_${areaName}',this.value); drawPhase2Map('1-2')"><option value="LR" ${dx==='LR'?'selected':''}>L->R</option><option value="RL" ${dx==='RL'?'selected':''}>R->L</option></select></div>
                    <div style="font-size:11px; color:#fff;">Y: <select onchange="updateMapConfig('1-2','dirY_${areaName}',this.value); drawPhase2Map('1-2')"><option value="DU" ${dy==='DU'?'selected':''}>D->U</option><option value="UD" ${dy==='UD'?'selected':''}>U->D</option></select></div>
                    <div style="font-size:11px; color:#fff;">轉: <select onchange="updateMapConfig('1-2','rot_${areaName}',this.value); drawPhase2Map('1-2')"><option value="0" ${rot==='0'?'selected':''}>0°</option><option value="90" ${rot==='90'?'selected':''}>90°</option><option value="180" ${rot==='180'?'selected':''}>180°</option><option value="270" ${rot==='270'?'selected':''}>270°</option></select></div>
                </div>`;
            }

            canvasArea.innerHTML = `
                <div class="split-layout">
                    <div class="split-left">
                        <h4 style="margin-top:0; color:#005a9e; font-size:14px;">1. 拖曳配置象限與方向</h4>
                        <div class="wafer-wrapper">
                            <div class="wafer-street-v"></div><div class="wafer-street-h"></div>
                            <div class="quad-zone qz-tl" id="zone_tl">${getQuadItemHTML(a1)}</div>
                            <div class="quad-zone qz-tr" id="zone_tr">${getQuadItemHTML(a2)}</div>
                            <div class="quad-zone qz-bl" id="zone_bl">${getQuadItemHTML(a3)}</div>
                            <div class="quad-zone qz-br" id="zone_br">${getQuadItemHTML(a4)}</div>
                        </div>
                        <button class="btn btn-primary" style="margin-top:15px; width:100%;" onclick="finalizeMap1_2()">💾 2. 確認組合並賦予絕對座標</button>
                        <div id="preview-1-7-mini" style="margin-top:20px; width:100%; text-align:center;"></div>
                    </div>
                    <div class="split-right" id="map-render-1-2"><p style="color:#aaa;">即時 MAP 預覽區</p></div>
                </div>
            `;
            legendDiv.innerHTML = `<div id="legend-content-${tabId}" style="display:flex; gap:10px;"></div>`;
            enableDragAndDrop1_2(); setTimeout(() => { drawPhase2Map('1-2'); renderMini17(); }, 100); 
        }
    } else if (tabId === '1-3') {
        if (!hasWaferID) { warningDiv.style.display = 'block'; configDiv.style.display = 'none'; canvasArea.style.display = 'none'; warningDiv.innerHTML = "⚠️ 找不到 <b>WaferID</b>！"; return; }
        
        configDiv.innerHTML = `
            <strong>🔗 1-3 MAP (綁定 1-1)</strong> ${waferSelectHTML} <label style="margin-left:15px;">顯示資料:</label><select id="map-val-${tabId}" onchange="${valChangeJS}">${getOpts('VAL', mc.valCol)}</select> ${dirHTML} 
            <div style="width:100%; height:5px;"></div>
            <strong>🐍 蛇行編序 (綁定 1-1 物理網格使用):</strong> 
            <label>起始點:</label><select id="snake-start-${tabId}" onchange="updateMapConfig('${tabId}','snakeStart',this.value); updateSnakeDirOptions('${tabId}')"><option value="TL" ${mc.snakeStart==='TL'?'selected':''}>左上</option><option value="TR" ${mc.snakeStart==='TR'?'selected':''}>右上</option><option value="BL" ${mc.snakeStart==='BL'?'selected':''}>左下</option><option value="BR" ${mc.snakeStart==='BR'?'selected':''}>右下</option></select>
            <label>首步方向:</label><select id="snake-dir-${tabId}" onchange="updateMapConfig('${tabId}','snakeDir',this.value)"></select>
            <label>路線:</label><select id="snake-route-${tabId}" onchange="updateMapConfig('${tabId}','snakeRoute',this.value)"><option value="SNAKE" ${mc.snakeRoute==='SNAKE'?'selected':''}>蛇行 (Z字)</option><option value="RASTER" ${mc.snakeRoute==='RASTER'?'selected':''}>單向 (打字機)</option></select>
            <button class="btn" style="background:#f1c40f; color:#000;" onclick="generateSnakeRoute('${tabId}')">1. 生成流水號與陣列映射</button>
            <button class="btn btn-primary" style="margin-left:auto;" onclick="drawPhase2Map('${tabId}')">2. 繪製 MAP</button>
        `;
        canvasArea.innerHTML = `<div id="map-render-${tabId}" style="width:100%; height:100%; display:flex; justify-content:center; align-items:center;"><p style="color:#aaa;">先生成流水號映射，再點擊繪製 MAP (由 1-1 帶出座標)</p></div>`;
        legendDiv.innerHTML = `<div id="legend-content-${tabId}" style="display:flex; gap:10px;"></div>`;
        setTimeout(() => updateSnakeDirOptions(tabId), 100);
    } else if (tabId === '1-4' || tabId === '1-5') {
        if (!hasWaferID) { warningDiv.style.display = 'block'; configDiv.style.display = 'none'; canvasArea.style.display = 'none'; warningDiv.innerHTML = "⚠️ 找不到 <b>WaferID</b>！"; return; }
        configDiv.innerHTML = `
            ${waferSelectHTML}
            <label style="margin-left:15px;">X座標:</label><select id="map-x-${tabId}" onchange="updateMapConfig('${tabId}','xCol',this.value)">${getOpts('X', mc.xCol)}</select>
            <label>Y座標:</label><select id="map-y-${tabId}" onchange="updateMapConfig('${tabId}','yCol',this.value)">${getOpts('Y', mc.yCol)}</select>
            <label>數值:</label><select id="map-val-${tabId}" onchange="${valChangeJS}">${getOpts('VAL', mc.valCol)}</select>
            ${dirHTML}
            <button class="btn btn-primary" style="margin-left:auto;" onclick="drawPhase2Map('${tabId}')">繪製 MAP</button>
        `;
        canvasArea.innerHTML = `<div id="map-render-${tabId}" style="width:100%; height:100%; display:flex; justify-content:center; align-items:center;"><p style="color:#aaa;">請點擊「繪製 MAP」</p></div>`;
        legendDiv.innerHTML = `<div id="legend-content-${tabId}" style="display:flex; gap:10px;"></div>`;
    }

    setTimeout(() => {
        let mc = configStore.pns[configStore.currentPN].settings[tabId].mapConfig;
        ['x', 'y', 'val', 'area'].forEach(key => {
            let el = document.getElementById(`map-${key}-${tabId}`);
            if (el && !mc[`${key}Col`]) mc[`${key}Col`] = el.value;
        });
        
        if (tabId === '1-2' && mc.isCombined && (!store.combinedData1_2 || store.combinedData1_2.length === 0)) { finalizeMap1_2(true); }
        if (tabId === '1-3' && mc.snakeStart && (!store.snakeGridMap || !store.headers.includes('New_No'))) { generateSnakeRoute('1-3', true); }

        if (tabId === '1-3' && store.snakeGridMap) drawPhase2Map('1-3');
        else if (tabId !== '1-6' && tabId !== '1-2' && tabId !== '1-3') drawPhase2Map(tabId); 
    }, 100);
}

window.renderMini17 = function() {
    let container = document.getElementById('preview-1-7-mini');
    if(!container) return;
    let store17 = dataStore['1-7']; let store1 = dataStore['1-1'];
    if(!store17 || store17.mergedData.length === 0 || !store1 || !store1.mapGridData || store1.mapGridData.length === 0) {
        container.innerHTML = '<div style="font-size:12px; color:#888; border:1px dashed #ccc; padding:20px;">(尚無 1-7 資料可供參考)</div>'; return;
    }
    container.innerHTML = `<div style="text-align:center;"><strong style="font-size:12px; color:#117a65;">📍 1-7 剝斷檔分佈參考</strong><br><canvas id="cvs-mini-1-7" width="160" height="160" style="background:#e6e6e6; border-radius:50%; margin-top:8px; border:2px solid #aaa; box-shadow:0 2px 5px rgba(0,0,0,0.2);"></canvas></div>`;
    
    let ctx = document.getElementById('cvs-mini-1-7').getContext('2d');
    let gridMap = {}; store1.mapGridData.forEach(g => { gridMap[String(g[0]).trim()] = {x: g[1], y: g[2]}; });
    
    let minX=Infinity, maxX=-Infinity, minY=Infinity, maxY=-Infinity;
    for(let k in gridMap) { let x = gridMap[k].x, y = gridMap[k].y; if(x<minX) minX=x; if(x>maxX) maxX=x; if(y<minY) minY=y; if(y>maxY) maxY=y; }
    let scale = Math.min(150 / (maxX-minX || 1), 150 / (maxY-minY || 1));

    let noIdx = store17.headers.findIndex(h => h && (h.toUpperCase()==='NO.' || h.toUpperCase()==='NO'));
    let statusIdx = store17.headers.findIndex(h => h && h.toUpperCase()==='STATUS');
    if(noIdx===-1) return;

    store17.mergedData.forEach(r => {
        let no = String(r[noIdx]).trim();
        let status = statusIdx > -1 ? String(r[statusIdx]).trim() : '';
        if(gridMap[no]) {
            let cx = 5 + (gridMap[no].x - minX) * scale;
            let cy = 155 - (gridMap[no].y - minY) * scale; 
            let pre = window.getCatColor(status);
            ctx.fillStyle = pre ? pre : (status !== '' ? '#f1c40f' : '#888');
            ctx.fillRect(cx, cy, 2, 2);
        }
    });
};

window.saveJoin1_6 = function() {
    let store5 = dataStore['1-5']; let store6 = dataStore['1-6'];
    if(!store6.joinedDataPreview) { alert("請先點擊執行關聯預覽！"); return; }
    store5.joinedHeaders = store6.joinedHeadersPreview; store5.joinedData = store6.joinedDataPreview;
    alert("✅ 關聯成功並儲存！請切換至 1-5 階段二，下拉選單已擴充 1-6 的欄位，可直接繪製合併後的 MAP。");
};

window.updateSnakeDirOptions = function(tabId) {
    let start = document.getElementById(`snake-start-${tabId}`).value;
    let dirSel = document.getElementById(`snake-dir-${tabId}`);
    let mc = configStore.pns[configStore.currentPN].settings[tabId].mapConfig;
    let opts = '';
    if(start === 'TR') opts = `<option value="D">向下</option><option value="L">向左</option>`;
    else if(start === 'BR') opts = `<option value="U">向上</option><option value="L">向左</option>`;
    else if(start === 'TL') opts = `<option value="D">向下</option><option value="R">向右</option>`;
    else if(start === 'BL') opts = `<option value="U">向上</option><option value="R">向右</option>`;
    dirSel.innerHTML = opts;
    if (opts.includes(`value="${mc.snakeDir}"`)) dirSel.value = mc.snakeDir;
    else { mc.snakeDir = dirSel.value; saveGlobalConfig(false); }
}

window.buildLegendUI = function(tabId, stats, uiMax, uiMin, isCategorical, legCnt=10, legInt=1) {
    let legDiv = document.getElementById(`map-controls-legend-${tabId}`); if(!legDiv) return;
    let valCol = document.getElementById(tabId === 'phase3' ? `p3-val-col` : `map-val-${tabId}`).value;
    
    if (isCategorical) {
        legDiv.style.display = 'flex';
        document.getElementById(`legend-content-${tabId}`).innerHTML = `<strong style="color:#d35400;">🏷️ 類別圖例<br><span style="font-size:11px; font-weight:normal; color:#555;">(${valCol})</span></strong>`;
        return;
    }

    let legMax = isNaN(uiMax) ? 'N/A' : uiMax.toFixed(4); let legMin = isNaN(uiMin) ? 'N/A' : uiMin.toFixed(4);
    let s_max = stats && stats.max !== undefined ? stats.max.toFixed(4) : 'N/A'; let s_min = stats && stats.min !== undefined ? stats.min.toFixed(4) : 'N/A';
    let s_q3 = stats && stats.q3 !== undefined ? stats.q3.toFixed(4) : 'N/A'; let s_q2 = stats && stats.q2 !== undefined ? stats.q2.toFixed(4) : 'N/A'; let s_q1 = stats && stats.q1 !== undefined ? stats.q1.toFixed(4) : 'N/A';
    let s_iqr = stats && stats.iqr !== undefined ? stats.iqr.toFixed(4) : 'N/A'; let s_ub = stats && stats.upperBound !== undefined ? stats.upperBound.toFixed(4) : 'N/A'; let s_lb = stats && stats.lowerBound !== undefined ? stats.lowerBound.toFixed(4) : 'N/A';

    legDiv.style.display = 'flex';
    let html = `
        <div style="display:flex; flex-direction:column; gap:6px; background:#fff; padding:10px; border:1px solid #ccc; border-radius:4px; min-width:160px;">
            <strong style="color:#d35400; font-size:13px; margin-bottom:2px; word-break:break-all;">🌈 圖例控制:<br><span style="font-size:11px; color:#555;">${valCol}</span></strong>
            <div style="display:flex; align-items:center; gap:5px;"><label style="width:65px; flex-shrink:0; font-size:13px;">上限:</label><input type="number" id="leg-max-${tabId}" value="${legMax}" step="any" onchange="calcLegend('${tabId}')" style="width:75px; padding:3px;"></div>
            <div style="display:flex; align-items:center; gap:5px;"><label style="width:65px; flex-shrink:0; font-size:13px;">下限:</label><input type="number" id="leg-min-${tabId}" value="${legMin}" step="any" onchange="calcLegend('${tabId}')" style="width:75px; padding:3px;"></div>
            <div style="display:flex; align-items:center; gap:5px;"><label style="width:65px; flex-shrink:0; font-size:13px;">分組數量:</label><input type="number" id="leg-cnt-${tabId}" value="${legCnt}" onchange="calcLegend('${tabId}', 'cnt')" style="width:75px; padding:3px;"></div>
            <div style="display:flex; align-items:center; gap:5px;"><label style="width:65px; flex-shrink:0; font-size:13px;">分組間隔:</label><input type="number" id="leg-int-${tabId}" value="${legInt.toFixed(4)}" step="any" onchange="calcLegend('${tabId}', 'int')" style="width:75px; padding:3px;"></div>
        </div>
        <div class="stats-panel">
            <h4>📊 統計摘要</h4>
            <div>Max: <span class="val">${s_max}</span></div>
            <div>Min: <span class="val">${s_min}</span></div>
            <div style="border-top:1px dashed #eee; margin:3px 0;"></div>
            <div>Q3 (75%): <span class="val">${s_q3}</span></div>
            <div>Q2 (中位數): <span class="val">${s_q2}</span></div>
            <div>Q1 (25%): <span class="val">${s_q1}</span></div>
            <div>IQR (Q3-Q1): <span class="val">${s_iqr}</span></div>
            <div style="border-top:1px dashed #eee; margin:3px 0;"></div>
            <div style="color:#d35400;">預設離群上限: <span class="val" style="color:#d35400;">${s_ub}</span></div>
            <div style="color:#d35400;">預設離群下限: <span class="val" style="color:#d35400;">${s_lb}</span></div>
        </div>
    `;
    document.getElementById(`legend-content-${tabId}`).innerHTML = html;
}

window.calcLegend = function(tabId, trigger) {
    let max = parseFloat(document.getElementById(`leg-max-${tabId}`).value); 
    let min = parseFloat(document.getElementById(`leg-min-${tabId}`).value);
    let cntInput = document.getElementById(`leg-cnt-${tabId}`);
    let intInput = document.getElementById(`leg-int-${tabId}`);

    if (isNaN(max) || isNaN(min)) return; let range = max - min; if (range <= 0) return;
    
    if (trigger === 'cnt' && cntInput && intInput) { let cnt = parseInt(cntInput.value); if (cnt > 0) intInput.value = (range / cnt).toFixed(4); } 
    else if (trigger === 'int' && cntInput && intInput) { let interval = parseFloat(intInput.value); if (interval > 0) cntInput.value = Math.ceil(range / interval); } 
    else if (cntInput && intInput) { let cnt = parseInt(cntInput.value); if (cnt > 0) intInput.value = (range / cnt).toFixed(4); }
    
    let conf = configStore.pns[configStore.currentPN].settings[tabId]; if(!conf.legConfig) conf.legConfig = {};
    let valCol = document.getElementById(tabId === 'phase3' ? `p3-val-col` : `map-val-${tabId}`).value;
    conf.legConfig[valCol] = { max, min, cnt: parseInt(cntInput?.value||10), int: parseFloat(intInput?.value||(range/10)) };
    saveGlobalConfig(false);
    
    if(tabId === 'phase3') drawPhase3Map(true); else if(document.getElementById(`cvs-${tabId}`)) drawPhase2Map(tabId, true);
}

function enableDragAndDrop1_2() {
    let draggedItem = null;
    document.querySelectorAll('.quad-item').forEach(item => {
        item.addEventListener('dragstart', function(e) { draggedItem = this; setTimeout(() => this.style.display = 'none', 0); });
        item.addEventListener('dragend', function() { this.style.display = 'flex'; document.querySelectorAll('.quad-zone').forEach(z => z.classList.remove('drag-over')); });
    });
    document.querySelectorAll('.quad-zone').forEach(zone => {
        zone.addEventListener('dragover', function(e) { e.preventDefault(); this.classList.add('drag-over'); });
        zone.addEventListener('dragleave', function() { this.classList.remove('drag-over'); });
        zone.addEventListener('drop', function(e) {
            e.preventDefault(); this.classList.remove('drag-over');
            if (draggedItem && this !== draggedItem.parentNode) {
                let srcId = draggedItem.parentNode.id.replace('zone_', ''); let tgtId = this.id.replace('zone_', ''); let srcVal = draggedItem.dataset.val; let tgtVal = this.querySelector('.quad-item').dataset.val;
                let conf = configStore.pns[configStore.currentPN].settings['1-2'];
                conf.mapConfig[`pos_zone_${srcId}`] = tgtVal; conf.mapConfig[`pos_zone_${tgtId}`] = srcVal;
                saveGlobalConfig(false); buildPhase2Controls('1-2'); 
            }
        });
    });
}

window.finalizeMap1_2 = function(silent = false) {
    let tabId = '1-2'; let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId];
    let mc = conf.mapConfig || {}; let baseHeaders = store.headers; let baseData = store.mergedData;

    let areaCol = document.getElementById(`map-area-${tabId}`) ? document.getElementById(`map-area-${tabId}`).value : (mc.origAreaCol || mc.areaCol);
    let xCol = document.getElementById(`map-x-${tabId}`) ? document.getElementById(`map-x-${tabId}`).value : (mc.origXCol || mc.xCol);
    let yCol = document.getElementById(`map-y-${tabId}`) ? document.getElementById(`map-y-${tabId}`).value : (mc.origYCol || mc.yCol);

    if (xCol === 'Global_C' || yCol === 'Global_R') { xCol = mc.origXCol || baseHeaders.find(h => h && h.toUpperCase().includes('X')) || ''; yCol = mc.origYCol || baseHeaders.find(h => h && h.toUpperCase().includes('Y')) || ''; }
    if (!areaCol || !xCol || !yCol) { if (!silent) alert("⚠️ 請先選擇 Area、X座標 與 Y座標 欄位！"); return; }
    let areaIdx = baseHeaders.indexOf(areaCol); let xIdx = baseHeaders.indexOf(xCol); let yIdx = baseHeaders.indexOf(yCol);
    if (areaIdx === -1 || xIdx === -1 || yIdx === -1) { if (!silent) alert("⚠️ 找不到對應的欄位，請重新確認四象限下拉選單！"); return; }

    mc.origAreaCol = areaCol; mc.origXCol = xCol; mc.origYCol = yCol;

    let areaData = {};
    baseData.forEach((r, i) => {
        let a = r[areaIdx]; let x = parseFloat(r[xIdx]); let y = parseFloat(r[yIdx]);
        if(a && !isNaN(x) && !isNaN(y)) {
            if(!areaData[a]) areaData[a] = { idxs: [], xSet: new Set(), ySet: new Set() };
            areaData[a].idxs.push({i, x, y}); areaData[a].xSet.add(x); areaData[a].ySet.add(y);
        }
    });

    let dims = {};
    Object.keys(areaData).forEach(a => {
        let uX = Array.from(areaData[a].xSet).sort((a,b)=>a-b); let uY = Array.from(areaData[a].ySet).sort((a,b)=>b-a);
        let xMap = new Map(); uX.forEach((v,i) => xMap.set(v,i)); let yMap = new Map(); uY.forEach((v,i) => yMap.set(v,i));
        let rot = mc['rot_' + a] || '0'; let c_len = uX.length, r_len = uY.length;
        if(rot === '90' || rot === '270') { c_len = uY.length; r_len = uX.length; }
        dims[a] = { c_len, r_len, xMap, yMap, max_c: uX.length-1, max_r: uY.length-1 };
    });

    let z_tl = mc.pos_zone_tl || 'Area_A'; let z_tr = mc.pos_zone_tr || 'Area_B'; let z_bl = mc.pos_zone_bl || 'Area_C'; let z_br = mc.pos_zone_br || 'Area_D';
    let w_left = Math.max(dims[z_tl]?.c_len||0, dims[z_bl]?.c_len||0); let h_top = Math.max(dims[z_tl]?.r_len||0, dims[z_tr]?.r_len||0);

    let offsets = {};
    offsets[z_tl] = { x: w_left - (dims[z_tl]?.c_len||0), y: h_top - (dims[z_tl]?.r_len||0) };
    offsets[z_tr] = { x: w_left + 1, y: h_top - (dims[z_tr]?.r_len||0) };
    offsets[z_bl] = { x: w_left - (dims[z_bl]?.c_len||0), y: h_top + 1 };
    offsets[z_br] = { x: w_left + 1, y: h_top + 1 };

    let combinedData = baseData.map(r => [...r]); let newHeaders = [...baseHeaders];
    let gcIdx = newHeaders.indexOf('Global_C'); let grIdx = newHeaders.indexOf('Global_R');
    if(gcIdx === -1) { newHeaders.push('Global_C', 'Global_R'); gcIdx = newHeaders.length - 2; grIdx = newHeaders.length - 1; combinedData.forEach(r => r.push('', '')); }

    Object.keys(areaData).forEach(a => {
        let dat = areaData[a]; let dim = dims[a];
        let dx = mc['dirX_' + a] || 'LR'; let dy = mc['dirY_' + a] || 'UD'; let rot = mc['rot_' + a] || '0';
        dat.idxs.forEach(p => {
            let c = dim.xMap.get(p.x); let r = dim.yMap.get(p.y);
            if(dx === 'RL') c = dim.max_c - c; if(dy === 'UD') r = dim.max_r - r;
            let fc = c, fr = r;
            if (rot === '90') { fc = dim.max_r - r; fr = c; } else if (rot === '180') { fc = dim.max_c - c; fr = dim.max_r - r; } else if (rot === '270') { fc = r; fr = dim.max_c - c; }
            combinedData[p.i][gcIdx] = fc + (offsets[a]?.x || 0); combinedData[p.i][grIdx] = fr + (offsets[a]?.y || 0);
        });
    });

    store.combinedData1_2 = combinedData; store.combinedHeaders1_2 = newHeaders; 
    mc.isCombined = true; mc.xCol = 'Global_C'; mc.yCol = 'Global_R';
    saveGlobalConfig(false);
    if (!silent) alert("✅ 四象限已拼接為單一絕對網格座標 (Global_C, Global_R)！");
    buildPhase2Controls('1-2');
};

window.resetMap1_2 = function() {
    let conf = configStore.pns[configStore.currentPN].settings['1-2']; conf.mapConfig.isCombined = false;
    if (conf.mapConfig.origXCol) conf.mapConfig.xCol = conf.mapConfig.origXCol; if (conf.mapConfig.origYCol) conf.mapConfig.yCol = conf.mapConfig.origYCol;
    saveGlobalConfig(false); buildPhase2Controls('1-2');
};

window.renderLeftJoinPreview = function() {
    let conf = configStore.pns[configStore.currentPN].settings['1-6'].mapConfig || {}; let store5 = dataStore['1-5']; let store6 = dataStore['1-6'];
    if(!store5 || !store6 || store5.mergedData.length===0 || store6.mergedData.length===0) { alert("請先匯入 1-5 與 1-6 資料！"); return; }
    
    let sel6 = document.getElementById('join-1-6'); let sel5 = document.getElementById('join-1-5');
    let jCol6 = sel6 ? sel6.value : conf.joinCol; let jCol5 = sel5 ? sel5.value : (conf.joinCol5 || 'BIN');
    let idx6 = store6.headers.indexOf(jCol6); let idx5 = store5.headers.indexOf(jCol5);
    if(idx6===-1 || idx5===-1) { alert("找不到指定的關聯欄位！"); return; }
    
    let map6 = {}; store6.mergedData.forEach(r => { map6[window.normKey(r[idx6])] = r; });
    let newHeaders = [...store5.headers]; let store6Idxs = [];
    store6.headers.forEach((h, i) => { let hu = h.toUpperCase(); if(hu !== jCol6.toUpperCase() && hu !== 'FILEPATH' && hu !== 'FILENAME') { newHeaders.push(h); store6Idxs.push(i); } });

    let joinedData = [];
    store5.mergedData.forEach(r => { let key = window.normKey(r[idx5]); let r6 = map6[key]; let extR6 = store6Idxs.map(i => r6 ? r6[i] : ''); joinedData.push([...r, ...extR6]); });
    
    store6.joinedDataPreview = joinedData; store6.joinedHeadersPreview = newHeaders;
    
    document.getElementById('map-render-1-6').innerHTML = `
        <div style="width:100%;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:10px;">
                <h4 style="color:#005a9e; margin:0;">🔗 Left Join 預覽 (1-5 為主體，1-6 附加於右側)</h4>
                <button class="btn btn-primary" onclick="saveJoin1_6()">💾 儲存並套用至 1-5 MAP</button>
            </div>
            <div class="table-wrapper">
                <div class="top-scroll-wrapper" onscroll="document.getElementById('p-tbl-1-6').scrollLeft = this.scrollLeft;"><div class="top-scroll-dummy" id="d-tbl-1-6"></div></div>
                <div class="preview-container" id="p-tbl-1-6" onscroll="this.previousElementSibling.scrollLeft = this.scrollLeft;">
                    ${renderTableHTML(joinedData, newHeaders, '1-6', true, false, 'filters_sec')}
                </div>
            </div>
        </div>
    `;
    setTimeout(()=> { let tableEl = document.querySelector(`#p-tbl-1-6 table`); if(tableEl) document.getElementById(`d-tbl-1-6`).style.width = tableEl.offsetWidth + 'px'; }, 100);
}

window.generateSnakeRoute = function(tabId, silent = false) {
    let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId];
    let mc = conf.mapConfig || {};
    let baseData = (tabId === '1-3' && conf.pivotConfig.enabled) ? store.pivotedData : store.mergedData;
    let baseHeaders = (tabId === '1-3' && conf.pivotConfig.enabled) ? store.pivotedHeaders : store.headers;
    
    let store1 = dataStore['1-1'];
    if (!store1 || !store1.mapGridData || store1.mapGridData.length === 0) { if (!silent) alert("⚠️ 請先前往 1-1 解析工位檔產生實體網格。"); return; }
    
    let gridPts = store1.mapGridData.map(g => ({ origNo: g[0], x: parseFloat(g[1]), y: parseFloat(g[2]) }));
    let start = mc.snakeStart || 'TL'; let dir = mc.snakeDir || 'R'; let route = mc.snakeRoute || 'SNAKE';
    
    let priAxis = ''; let priSort = 1; let secSort = 1;
    if(start==='TR' && dir==='D') { priAxis='X'; priSort=-1; secSort=-1; } 
    else if(start==='TR' && dir==='L') { priAxis='Y'; priSort=-1; secSort=-1; } 
    else if(start==='BR' && dir==='U') { priAxis='X'; priSort=-1; secSort=1; } 
    else if(start==='BR' && dir==='L') { priAxis='Y'; priSort=1; secSort=-1; } 
    else if(start==='TL' && dir==='D') { priAxis='X'; priSort=1; secSort=-1; } 
    else if(start==='TL' && dir==='R') { priAxis='Y'; priSort=-1; secSort=1; } 
    else if(start==='BL' && dir==='U') { priAxis='X'; priSort=1; secSort=1; } 
    else if(start==='BL' && dir==='R') { priAxis='Y'; priSort=1; secSort=1; } 

    let groups = {}; gridPts.forEach(p => { let k = priAxis==='X' ? p.x : p.y; if(!groups[k]) groups[k] = []; groups[k].push(p); });
    let sortedKeys = Object.keys(groups).map(Number).sort((a,b) => (a-b)*priSort);

    let snakeGridMap = {}; let seq = 1;
    sortedKeys.forEach((k, i) => {
        let pts = groups[k];
        let curSecSort = (route === 'SNAKE' && i % 2 !== 0) ? -secSort : secSort;
        pts.sort((a,b) => { let vA = priAxis==='X' ? a.y : a.x; let vB = priAxis==='X' ? b.y : b.x; return (vA - vB) * curSecSort; });
        pts.forEach(p => { snakeGridMap[seq.toString()] = { x: p.x, y: p.y }; seq++; });
    });

    store.snakeGridMap = snakeGridMap;

    let noIdx = baseHeaders.indexOf('New_No');
    if(noIdx === -1) { baseHeaders.push('New_No'); noIdx = baseHeaders.length - 1; baseData.forEach(r => r.push('')); }
    let cnt = 1; baseData.forEach(r => r[noIdx] = cnt++);
    
    let valSelect = document.getElementById(`map-val-${tabId}`);
    if (valSelect && !valSelect.querySelector('option[value="New_No"]')) { valSelect.innerHTML += `<option value="New_No">New_No</option>`; }

    if (!silent) { alert("New_No 流水號映射已依據蛇行邏輯生成！自動為您切換顯示資料..."); if (valSelect) valSelect.value = 'New_No'; updateMapConfig(tabId, 'valCol', 'New_No'); } 
    else { if (valSelect && mc.valCol) { valSelect.value = mc.valCol; } }
    
    drawPhase2Map(tabId);
}

window.drawPhase2Map = function(tabId, skipStatsRecalc = false) {
    let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId]; window.activeMapRects = []; 
    if(!document.getElementById(`map-render-${tabId}`)) return;

    let valEl = document.getElementById(`map-val-${tabId}`); let vCol = valEl ? valEl.value : null;
    let baseHeaders = store.headers; let baseData = store.mergedData;
    if (tabId === '1-3' && conf.pivotConfig && conf.pivotConfig.enabled) { baseHeaders = store.pivotedHeaders; baseData = store.pivotedData; }
    if (tabId === '1-5' && store.joinedHeaders && store.joinedData) { baseHeaders = store.joinedHeaders; baseData = store.joinedData; }
    if (tabId === '1-2' && conf.mapConfig.isCombined && store.combinedData1_2) { baseData = store.combinedData1_2; baseHeaders = store.combinedHeaders1_2 || store.headers; }
    
    let vIdx = vCol ? baseHeaders.indexOf(vCol) : -1;
    let dirX = conf.mapConfig.dirX || (document.getElementById(`map-dir-x-${tabId}`) ? document.getElementById(`map-dir-x-${tabId}`).value : 'LR');
    let dirY = conf.mapConfig.dirY || (document.getElementById(`map-dir-y-${tabId}`) ? document.getElementById(`map-dir-y-${tabId}`).value : 'DU');
    let selectedWafer = document.getElementById(`map-wafer-${tabId}`) ? document.getElementById(`map-wafer-${tabId}`).value : null;

    showLoading("正在渲染 MAP...");
    
    setTimeout(() => {
     try {
        let points = []; let xSet = new Set(), ySet = new Set();
        let isCategorical = false; let catSet = new Set(); let rawValues = []; let uniqueVals = new Set();
        let waferIdx = baseHeaders.findIndex(h => h && h.toUpperCase().replace(/[\s_]/g, '') === 'WAFERID');

        if (tabId === '1-1') {
            if(!store.mapGridData || store.mapGridData.length === 0) { hideLoading(); return; }
            store.mapGridData.forEach((g, i) => {
                let x = g[1]; let y = g[2]; let rawV = '';
                if (vCol === 'No.') rawV = String(g[0]); else if (vCol === 'Map_X') rawV = String(g[1]); else if (vCol === 'Map_Y') rawV = String(g[2]); else if (vCol === 'Row_Ref') rawV = String(g[3]); 
                let v = parseFloat(rawV);
                points.push({no: g[0], x, y, v: isNaN(v) ? rawV : v, a: '', rawV: rawV});
                xSet.add(x); ySet.add(y); if(rawV !== '') uniqueVals.add(rawV);
            });
        } else if (tabId === '1-7' || tabId === '1-3') {
            let store1 = dataStore['1-1'];
            if (!store1 || !store1.mapGridData || store1.mapGridData.length === 0) { hideLoading(); alert("⚠️ 請先解析 1-1 工位檔產生實體網格。"); return; }
            let gridMap = (tabId==='1-3' && store.snakeGridMap) ? store.snakeGridMap : {};
            if (tabId==='1-7' || !store.snakeGridMap) { store1.mapGridData.forEach(g => { gridMap[String(g[0]).trim()] = {x: g[1], y: g[2]}; }); }
            
            let noIdx = baseHeaders.indexOf('New_No'); if (noIdx === -1) noIdx = baseHeaders.findIndex(h => h && (h.toUpperCase()==='NO.' || h.toUpperCase()==='NO'));
            
            baseData.forEach(r => {
                if(selectedWafer && waferIdx > -1 && String(r[waferIdx]).trim() !== selectedWafer) return;
                let no = noIdx > -1 ? String(r[noIdx]).trim() : ''; let rawV = (vIdx > -1 && r[vIdx] !== undefined) ? String(r[vIdx]).trim() : ''; 
                let v = parseFloat(rawV);
                if (gridMap[no]) {
                    let x = gridMap[no].x; let y = gridMap[no].y;
                    points.push({no, x, y, v: isNaN(v) ? rawV : v, a: '', rawV: rawV});
                    xSet.add(x); ySet.add(y); if(rawV !== '') uniqueVals.add(rawV);
                }
            });
        } else {
            let xCol = 'Map_X', yCol = 'Map_Y';
            if (tabId === '1-2' && conf.mapConfig.isCombined) { xCol = 'Global_C'; yCol = 'Global_R'; } 
            else { let xEl = document.getElementById(`map-x-${tabId}`); let yEl = document.getElementById(`map-y-${tabId}`); if(xEl) xCol = xEl.value; if(yEl) yCol = yEl.value; }
            let xIdx = baseHeaders.indexOf(xCol); let yIdx = baseHeaders.indexOf(yCol);
            let areaEl = document.getElementById(`map-area-${tabId}`);

            baseData.forEach((r, idx) => {
                if(selectedWafer && waferIdx > -1 && String(r[waferIdx]).trim() !== selectedWafer) return;
                let x = parseFloat(r[xIdx]); let y = parseFloat(r[yIdx]); let rawV = (vIdx > -1 && r[vIdx] !== undefined) ? String(r[vIdx]).trim() : ''; let v = parseFloat(rawV);
                let a = (tabId === '1-2' && !conf.mapConfig.isCombined && areaEl) ? String(r[baseHeaders.indexOf(areaEl.value)]||'') : '';
                if (!isNaN(x) && !isNaN(y)) { 
                    points.push({no: idx+1, x, y, v: isNaN(v) ? rawV : v, a, rawV: rawV});
                    xSet.add(x); ySet.add(y); if(rawV !== '') uniqueVals.add(rawV);
                }
            });
        }

        if(points.length === 0) { 
            hideLoading(); let xStr = (typeof xCol !== 'undefined') ? xCol : '未設定'; let yStr = (typeof yCol !== 'undefined') ? yCol : '未設定';
            document.getElementById(`map-render-${tabId}`).innerHTML = `<div style='background:#f8d7da; padding:15px; border-radius:5px; border:1px solid #f5c6cb;'><strong style='color:#721c24; font-size:16px;'>⚠️ 所選條件查無有效 MAP 數據</strong><ul style='color:#721c24; font-size:13px; margin-top:8px;'><li>請確認 <b>WaferID</b> 是否與資料夾中的資料相符。</li><li>請確認上方下拉選單的 <b>X座標 (${xStr})</b> 與 <b>Y座標 (${yStr})</b> 是否正確。如果選到了 FilePath 或非數字欄位，系統將無法繪圖！</li></ul></div>`; return; 
        }
        
        let hasText = Array.from(uniqueVals).some(v => isNaN(parseFloat(v)) || String(v).trim() === '');
        if (tabId === '1-1' && (!vCol || vCol === '')) { isCategorical = true; } 
        else { isCategorical = (vIdx === -1) ? true : ((hasText && uniqueVals.size <= 50) || (uniqueVals.size <= 15 && uniqueVals.size > 0)); if (['No.', 'New_No', 'NO.', 'NO', 'ROW_REF'].includes(String(vCol).toUpperCase())) isCategorical = false; }
        if (uniqueVals.size > 50) isCategorical = false; 

        points.forEach(p => { if(p.rawV !== '') { if(isCategorical) catSet.add(p.rawV); else { let pv = parseFloat(p.rawV); if(!isNaN(pv)) rawValues.push(pv); } } });

        let stats = null; let uiMax = NaN, uiMin = NaN; let legCnt = 10, legInt = 1;
        if (!isCategorical) {
            stats = calculateStats(rawValues); let vConf = conf.legConfig && conf.legConfig[vCol] ? conf.legConfig[vCol] : null;
            if (!skipStatsRecalc && (!vConf || vConf.max === undefined)) { uiMax = stats ? stats.upperBound : NaN; uiMin = stats ? stats.lowerBound : NaN; legInt = ((uiMax - uiMin) / legCnt) || 1; } 
            else if (skipStatsRecalc && document.getElementById(`leg-max-${tabId}`)) { uiMax = parseFloat(document.getElementById(`leg-max-${tabId}`).value); uiMin = parseFloat(document.getElementById(`leg-min-${tabId}`).value); legCnt = parseInt(document.getElementById(`leg-cnt-${tabId}`).value) || 10; legInt = parseFloat(document.getElementById(`leg-int-${tabId}`).value) || ((uiMax-uiMin)/legCnt) || 1; } 
            else if (vConf) { uiMax = vConf.max; uiMin = vConf.min; legCnt = vConf.cnt || 10; legInt = vConf.int || ((uiMax - uiMin) / legCnt) || 1; }
            buildLegendUI(tabId, stats, uiMax, uiMin, false, legCnt, legInt);
        } else { buildLegendUI(tabId, null, NaN, NaN, true); }

        let catColorMap = {}; let getRainbowColor; let legHTML = '';
        if (isCategorical) {
            let colors = ['#3498db','#f1c40f','#9b59b6','#e67e22','#1abc9c','#34495e','#ff69b4','#8a2be2','#a52a2a','#d2691e','#008080','#4682b4']; let cIdx = 0; let catArray = Array.from(catSet); let legendLimit = Math.min(catArray.length, 50); 
            for(let i=0; i<catArray.length; i++) { let c = catArray[i]; let pre = window.getCatColor(c); if(pre) catColorMap[c] = pre; else { catColorMap[c] = colors[cIdx % colors.length]; cIdx++; } }
            legHTML = `<div class="legend-container" style="justify-content:flex-start;">`;
            if (vIdx === -1 && tabId === '1-1' && (vCol === '' || !vCol)) { legHTML += `<div class="cat-legend-item"><div class="cat-color-box" style="background:#1e8449"></div>基礎網格</div>`; } 
            else { for(let i=0; i<legendLimit; i++) { let k = catArray[i]; legHTML += `<div class="cat-legend-item"><div class="cat-color-box" style="background:${catColorMap[k]}"></div>${String(k).substring(0,10)}</div>`; } if (catArray.length > 50) legHTML += `<div style="font-size:10px; color:#888; margin-top:5px;">...等 ${catArray.length} 項</div>`; }
            legHTML += `</div>`;
        } else {
            getRainbowColor = (val) => { if(val > uiMax) val = uiMax; if(val < uiMin) val = uiMin; let pct = (uiMax===uiMin) ? 0.5 : (val - uiMin) / (uiMax - uiMin); let hue = (1.0 - pct) * 240; return `hsl(${hue}, 100%, 50%)`; };
            legHTML = `<div class="legend-container"><div style="font-size:11px; margin-bottom:5px; font-weight:bold;">Max: ${isNaN(uiMax)?'N/A':uiMax.toFixed(2)}</div><div class="legend-bar"></div><div style="font-size:11px; margin-top:5px; font-weight:bold;">Min: ${isNaN(uiMin)?'N/A':uiMin.toFixed(2)}</div></div>`;
        }

        let renderArea = document.getElementById(`map-render-${tabId}`);
        if (tabId === '1-2' && !conf.mapConfig.isCombined) {
            let areaData = {}; points.forEach(p => { if(!areaData[p.a]) areaData[p.a] = { pts: [], xSet: new Set(), ySet: new Set() }; areaData[p.a].pts.push(p); areaData[p.a].xSet.add(p.x); areaData[p.a].ySet.add(p.y); });
            let maxDim = 0; let areaMaps = {};
            Object.keys(areaData).forEach(a => {
                if (areaData[a].xSet.size > maxDim) maxDim = areaData[a].xSet.size; if (areaData[a].ySet.size > maxDim) maxDim = areaData[a].ySet.size;
                let localUX = Array.from(areaData[a].xSet).sort((a,b)=>a-b); let localUY = Array.from(areaData[a].ySet).sort((a,b)=>b-a);
                let lxMap = new Map(); localUX.forEach((v,i)=>lxMap.set(v,i)); let lyMap = new Map(); localUY.forEach((v,i)=>lyMap.set(v,i));
                areaMaps[a] = { max_c: localUX.length - 1, max_r: localUY.length - 1, lxMap, lyMap };
            });
            if (maxDim === 0) maxDim = 10;
            let cols = maxDim * 2; let rows = maxDim * 2; let cvsWidth = Math.min(700, Math.max(300, cols * 10)); let cvsHeight = Math.min(700, Math.max(300, rows * 10)); let cellW = cvsWidth / cols; let cellH = cvsHeight / rows;
            let gw = maxDim * cellW; let gh = maxDim * cellH;
            let z_tl = document.querySelector('#zone_tl .quad-item')?.dataset.val || ''; let z_tr = document.querySelector('#zone_tr .quad-item')?.dataset.val || ''; let z_bl = document.querySelector('#zone_bl .quad-item')?.dataset.val || ''; let z_br = document.querySelector('#zone_br .quad-item')?.dataset.val || '';
            let xOffsets = {}, yOffsets = {}; xOffsets[z_tl] = 0; yOffsets[z_tl] = 0; xOffsets[z_tr] = gw; yOffsets[z_tr] = 0; xOffsets[z_bl] = 0; yOffsets[z_bl] = gh; xOffsets[z_br] = gw; yOffsets[z_br] = gh;

            renderArea.innerHTML = `<div style="display:flex; align-items:flex-start;"><canvas id="cvs-${tabId}" width="${cvsWidth}" height="${cvsHeight}" style="border:1px solid #ccc; background:#e6e6e6; cursor:crosshair;"></canvas>${legHTML}</div>`;
            let cvs = document.getElementById(`cvs-${tabId}`); let ctx = cvs.getContext('2d'); ctx.clearRect(0, 0, cvsWidth, cvsHeight); ctx.font = "8px Arial"; ctx.textAlign = "center"; ctx.textBaseline = "middle";

            points.forEach(p => {
                let am = areaMaps[p.a]; let orig_c = am.lxMap.get(p.x); let orig_r = am.lyMap.get(p.y);
                let dx = conf.mapConfig['dirX_' + p.a] || 'LR'; let dy = conf.mapConfig['dirY_' + p.a] || 'UD'; let rot = conf.mapConfig['rot_' + p.a] || '0';
                let c = orig_c; let r = orig_r; if(dx === 'RL') c = am.max_c - c; if(dy === 'UD') r = am.max_r - r;
                let final_c = c, final_r = r; if (rot === '90') { final_c = am.max_r - r; final_r = c; } else if (rot === '180') { final_c = am.max_c - c; final_r = am.max_r - r; } else if (rot === '270') { final_c = r; final_r = am.max_c - c; }
                let px = final_c * cellW + (xOffsets[p.a] || 0); let py = final_r * cellH + (yOffsets[p.a] || 0);

                let pointColor = '#e6e6e6'; if (isCategorical) { if(p.rawV !== '') pointColor = catColorMap[p.rawV] || '#e6e6e6'; } else { if(p.rawV !== '' && !isNaN(p.v)) pointColor = getRainbowColor(p.v); }
                ctx.fillStyle = pointColor; ctx.fillRect(px, py, cellW, cellH); 
                if (cellW >= 1.5 && cellH >= 1.5) { ctx.strokeStyle = "rgba(255,255,255,0.3)"; ctx.lineWidth = 1; ctx.strokeRect(px, py, cellW, cellH); }
                if (vCol !== '' && cellW > 20 && cellH > 10 && p.rawV !== '') { ctx.fillStyle = (isCategorical || p.v > (uiMax+uiMin)/2) ? "#fff" : "#000"; let dispV = isCategorical ? String(p.rawV).substring(0,3) : p.v.toFixed(1); ctx.fillText(dispV, px + cellW/2, py + cellH/2); }
                window.activeMapRects.push({ tabId: tabId, x:px, y:py, w:cellW, h:cellH, str:`[No: ${p.no}] X:${p.x}, Y:${p.y}` + (vCol!=='' && p.rawV!=='' ? `<br>${vCol}: <b style="color:#f1c40f;">${p.rawV}</b>` : '') + `<br>Area: ${p.a}` });
            });
        } else {
            let uX = Array.from(xSet).sort((a,b)=>a-b); let uY = Array.from(ySet).sort((a,b)=>b-a);
            let xMap = new Map(); uX.forEach((v, i) => xMap.set(v, i)); let yMap = new Map(); uY.forEach((v, i) => yMap.set(v, i));
            let cols = uX.length; let rows = uY.length;
            let cvsWidth = Math.min(700, Math.max(300, cols * 20)); let cvsHeight = Math.min(700, Math.max(300, rows * 20)); let cellW = cvsWidth / (cols||1); let cellH = cvsHeight / (rows||1);

            renderArea.innerHTML = `<div style="display:flex; align-items:flex-start;"><canvas id="cvs-${tabId}" width="${cvsWidth}" height="${cvsHeight}" style="border:1px solid #ccc; background:#e6e6e6; cursor:crosshair;"></canvas>${legHTML}</div>`;
            let cvs = document.getElementById(`cvs-${tabId}`); let ctx = cvs.getContext('2d'); ctx.clearRect(0, 0, cvsWidth, cvsHeight); ctx.font = "10px Arial"; ctx.textAlign = "center"; ctx.textBaseline = "middle";

            points.forEach(p => {
                let c = xMap.get(p.x); let r = yMap.get(p.y);
                if(dirX === 'RL') c = (cols - 1 - c); if(dirY === 'UD') r = (rows - 1 - r); 
                let rot = conf.mapConfig.rot || '0'; let fc = c, fr = r; let max_c = cols - 1; let max_r = rows - 1;
                if (rot === '90') { fc = max_r - r; fr = c; } else if (rot === '180') { fc = max_c - c; fr = max_r - r; } else if (rot === '270') { fc = r; fr = max_c - c; }
                let px = fc * cellW; let py = fr * cellH;
                
                let pointColor = '#e6e6e6'; if (vIdx === -1 && tabId === '1-1' && (vCol === '' || !vCol)) { pointColor = '#1e8449'; } else if (isCategorical) { if(p.rawV !== '') pointColor = catColorMap[p.rawV] || '#e6e6e6'; } else { if(p.rawV !== '' && !isNaN(p.v)) pointColor = getRainbowColor(p.v); }
                ctx.fillStyle = pointColor; ctx.fillRect(px, py, cellW, cellH); 
                if (cellW >= 1.5 && cellH >= 1.5) { ctx.strokeStyle = "rgba(255,255,255,0.3)"; ctx.lineWidth = 1; ctx.strokeRect(px, py, cellW, cellH); }
                if (vCol !== '' && cellW > 20 && cellH > 10 && p.rawV !== '') { ctx.fillStyle = (isCategorical || p.v > (uiMax+uiMin)/2) ? "#fff" : "#000"; let dispV = isCategorical ? String(p.rawV).substring(0,3) : p.v.toFixed(1); ctx.fillText(dispV, px + cellW/2, py + cellH/2); }
                window.activeMapRects.push({ tabId: tabId, x:px, y:py, w:cellW, h:cellH, str:`[No: ${p.no}] X:${p.x}, Y:${p.y}` + (vCol!=='' && p.rawV!=='' ? `<br>${vCol}: <b style="color:#f1c40f;">${p.rawV}</b>` : '') });
            });
        }
     } catch(err) { console.error(err); document.getElementById(`map-render-${tabId}`).innerHTML = `<p style="color:#d9534f; font-weight:bold;">⚠️ 渲染發生異常: ${err.message}</p>`; } finally { hideLoading(); }
    }, 50);
}

function handleTooltip(e) {
    let tt = document.getElementById('map-tooltip');
    if (!window.activeMapRects || window.activeMapRects.length === 0) { tt.style.display = 'none'; return; }
    let hovered = false; let visibleCanvases = document.querySelectorAll('canvas[id^="cvs-"]:not([id*="mini"])');
    for(let cvs of visibleCanvases) {
        if (cvs.offsetParent === null) continue;
        let rect = cvs.getBoundingClientRect(); let mx = e.clientX - rect.left; let my = e.clientY - rect.top;
        if (mx >= 0 && mx <= rect.width && my >= 0 && my <= rect.height) {
            let scaleX = cvs.width / rect.width; let scaleY = cvs.height / rect.height; let cx = mx * scaleX; let cy = my * scaleY; let cId = cvs.id.replace('cvs-','');
            for (let i = 0; i < window.activeMapRects.length; i++) {
                let r = window.activeMapRects[i];
                if (r.tabId === cId && cx >= r.x && cx <= r.x + r.w && cy >= r.y && cy <= r.y + r.h) {
                    tt.innerHTML = r.str; tt.style.display = 'block'; 
                    let tw = tt.offsetWidth; let th = tt.offsetHeight; let tx = e.clientX + 15; let ty = e.clientY + 15;
                    if (tx + tw > window.innerWidth) tx = e.clientX - tw - 15; if (ty + th > window.innerHeight) ty = e.clientY - th - 15; 
                    tt.style.left = tx + 'px'; tt.style.top = ty + 'px'; hovered = true; break;
                }
            }
        }
        if (hovered) break;
    }
    if (!hovered) tt.style.display = 'none';
}

window.openFilter = function(tabId, colIdx, btnElement, filterKey = 'filters') {
    let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId]; if(!store[filterKey]) store[filterKey] = {};
    let dataArray = store.mergedData;
    if (filterKey === 'filters_sec') { if (tabId === '1-1') dataArray = store.mapGridData; else if (tabId === '1-3' && conf.pivotConfig && conf.pivotConfig.enabled) dataArray = store.pivotedData; else if (tabId === '1-6') dataArray = store.joinedDataPreview || []; } else if (tabId === 'phase3') { dataArray = store.mergedData; }
    
    let uniqueVals = [...new Set(dataArray.map(r => r[colIdx]))].sort();
    currentFilterInfo = { tabId, colIdx, filterKey, set: new Set(uniqueVals) };
    let appliedFilter = store[filterKey][colIdx]; let html = '';
    uniqueVals.slice(0, 100).forEach(v => { let isChecked = appliedFilter ? appliedFilter.has(String(v)) : true; let displayV = v === '' ? '(空白)' : v; html += `<label class="filter-item"><input type="checkbox" value="${v}" ${isChecked?'checked':''} class="filter-cb"> ${displayV}</label>`; });
    if(uniqueVals.length > 100) html += `<div style="font-size:11px; color:#888;">...等共 ${uniqueVals.length} 項</div>`;
    document.getElementById('filter-list').innerHTML = html;
    let modal = document.getElementById('filter-modal'); let rect = btnElement.getBoundingClientRect(); modal.style.left = rect.left + 'px'; modal.style.top = (rect.bottom + window.scrollY + 5) + 'px'; modal.style.display = 'block'; document.getElementById('filter-all').checked = !appliedFilter;
}

window.toggleAllFilters = function(cb) { document.querySelectorAll('.filter-cb').forEach(el => el.checked = cb.checked); }

window.applyFilter = function() {
    let { tabId, colIdx, filterKey } = currentFilterInfo; let store = dataStore[tabId];
    let checkedVals = Array.from(document.querySelectorAll('.filter-cb:checked')).map(el => el.value);
    if (checkedVals.length === document.querySelectorAll('.filter-cb').length) { delete store[filterKey][colIdx]; } else { store[filterKey][colIdx] = new Set(checkedVals); }
    document.getElementById('filter-modal').style.display = 'none'; 
    if (tabId === '1-6' && filterKey === 'filters_sec') { document.getElementById('p-tbl-1-6').innerHTML = renderTableHTML(store.joinedDataPreview, store.joinedHeadersPreview, '1-6', true, false, 'filters_sec'); setTimeout(()=> { let tableEl = document.querySelector(`#p-tbl-1-6 table`); if(tableEl) document.getElementById(`d-tbl-1-6`).style.width = tableEl.offsetWidth + 'px'; }, 100); } 
    else if (tabId === 'phase3') { document.getElementById('p3-table-container').innerHTML = renderTableHTML(store.mergedData, store.headers, 'phase3', false, false, 'filters'); } 
    else { renderPreview(tabId); }
}

window.buildPhase3Controls = function() {
    let p3Store = dataStore['phase3']; let conf = configStore.pns[configStore.currentPN].settings['phase3'];
    if(!conf) conf = configStore.pns[configStore.currentPN].settings['phase3'] = { transforms: {}, legConfig: {} };

    let ctrlDiv = document.getElementById('p3-controls'); let contentDiv = document.getElementById('p3-content');
    const tableNameMap = { '1-1': 'BLK_Supply', '1-2': 'Prober', '1-3': 'Nikon', '1-4': 'AOI_Dim', '1-5': 'AOI_Defect', '1-7': 'WCM' };

    let trHtml = `<div style="display:flex; justify-content:space-between; align-items:center;"><h4 style="margin-top:0; color:#1e8449;">🔧 各表對齊微調 (將疊加至 1-1 基準網格)</h4><div style="display:flex; gap:10px;"><button class="btn" style="background:#e67e22; color:#fff;" onclick="syncP3Transforms()">🔄 0. 同步階段二方向</button><button class="btn btn-primary" onclick="preCheckMasterJoin()">🚀 1. 執行 Master Join 融合</button><button class="btn" style="background:#2874a6; color:#fff;" onclick="exportP3CSV(false)">📥 2. 匯出橫表(CSV)</button><button class="btn" style="background:#8e44ad; color:#fff;" onclick="exportP3CSV(true)">📥 3. 轉置並匯出直表(CSV)</button></div></div><div style="display:flex; flex-wrap:wrap; gap:15px; margin-bottom:15px; margin-top:10px;">`;
    ['1-1', '1-2', '1-3', '1-4', '1-5', '1-7'].forEach(t => {
        let tr = conf.transforms[t] || { dirX: 'LR', dirY: 'DU', rot: '0' }; let tName = tableNameMap[t];
        trHtml += `<div style="background:#fff; border:1px solid #ccc; padding:10px; border-radius:4px; flex:1; min-width:160px; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,0.1); display:flex; flex-direction:column; align-items:center;">
            <strong style="color:#005a9e; font-size:13px; margin-bottom:5px;">${tName}</strong><div id="p3-dim-${t}" style="font-size:11px; margin-bottom:6px; background:#f0f8ff; padding:2px 6px; border-radius:3px;">讀取中...</div>
            <div style="width:90%; display:flex; justify-content:space-between; gap:4px; margin-bottom:8px;"><select style="flex:1; padding:2px; font-size:11px; min-width:0;" title="X軸" onchange="updateP3Transform('${t}','dirX',this.value); renderP3MiniMaps()"><option value="LR" ${tr.dirX==='LR'?'selected':''}>L👉R</option><option value="RL" ${tr.dirX==='RL'?'selected':''}>R👈L</option></select><select style="flex:1; padding:2px; font-size:11px; min-width:0;" title="Y軸" onchange="updateP3Transform('${t}','dirY',this.value); renderP3MiniMaps()"><option value="DU" ${tr.dirY==='DU'?'selected':''}>D☝️U</option><option value="UD" ${tr.dirY==='UD'?'selected':''}>U👇D</option></select><select style="flex:1; padding:2px; font-size:11px; min-width:0;" title="旋轉" onchange="updateP3Transform('${t}','rot',this.value); renderP3MiniMaps()"><option value="0" ${tr.rot==='0'?'selected':''}>0°</option><option value="90" ${tr.rot==='90'?'selected':''}>90°</option><option value="180" ${tr.rot==='180'?'selected':''}>180°</option><option value="270" ${tr.rot==='270'?'selected':''}>270°</option></select></div>
            <div id="p3-mini-map-${t}" style="width:90%; aspect-ratio:1/1; background:#f4f7f6; border:1px dashed #aaa; display:flex; align-items:center; justify-content:center; flex-direction:column; border-radius:4px;"></div></div>`;
    });
    trHtml += `</div>`; ctrlDiv.innerHTML = trHtml;

    if(!p3Store || p3Store.mergedData.length === 0) { contentDiv.innerHTML = '<div style="color:#666; font-size:16px; padding:30px; text-align:center; border:2px dashed #ccc; border-radius:8px;">請點擊上方「🚀 執行 Master Join 融合」按鈕，結果將顯示於此</div>'; setTimeout(renderP3MiniMaps, 100); return; }

    let opts = p3Store.headers.map(h => `<option value="${h}" ${conf.valCol===h?'selected':''}>${h}</option>`).join('');
    let waferCols = p3Store.headers.filter(h => h && h.toUpperCase().includes('WAFERID')); let wOpts = '';
    if (waferCols.length > 0) {
        let wSet = new Set(); p3Store.mergedData.forEach(r => { waferCols.forEach(wc => { let idx = p3Store.headers.indexOf(wc); if(idx > -1 && r[idx]) wSet.add(String(r[idx]).trim()); }); });
        let wArr = Array.from(wSet).filter(v=>v);
        if(wArr.length > 0) {
            if(!conf.selectedWafer || !wArr.includes(conf.selectedWafer)) conf.selectedWafer = wArr[0];
            wOpts = `<label style="margin-left:15px; color:#d35400;">📍 WaferID過濾: <select id="p3-wafer-col" onchange="updateP3Config('selectedWafer', this.value); drawPhase3Map()"><option value="">-- 全部顯示 --</option>${wArr.map(w => `<option value="${w}" ${conf.selectedWafer===w?'selected':''}>${w}</option>`).join('')}</select></label>`;
        }
    }

    contentDiv.innerHTML = `
        <div style="display:flex; justify-content:space-between; align-items:center; background:#d4e6f1; padding:10px; border-radius:4px; margin-bottom:10px; border:1px solid #b8d4ea;">
            <div style="display:flex; align-items:center; gap:15px; flex-wrap:wrap;">
                <strong style="color:#005a9e;">📍 Master Join MAP</strong> ${wOpts}
                <label style="margin-left:15px;">X座標: <select id="p3-x-col" onchange="updateP3Config('xCol', this.value); drawPhase3Map()">${p3Store.headers.map(h => `<option value="${h}" ${conf.xCol===h || (h==='[BLK_Supply] Map_X' && !conf.xCol)?'selected':''}>${h}</option>`).join('')}</select></label>
                <label>Y座標: <select id="p3-y-col" onchange="updateP3Config('yCol', this.value); drawPhase3Map()">${p3Store.headers.map(h => `<option value="${h}" ${conf.yCol===h || (h==='[BLK_Supply] Map_Y' && !conf.yCol)?'selected':''}>${h}</option>`).join('')}</select></label>
                <label>顯示數值: <select id="p3-val-col" onchange="updateP3Config('valCol', this.value); drawPhase3Map()">${opts}</select></label>
            </div>
        </div>
        <div style="display:flex; gap:20px; align-items:flex-start; margin-bottom:20px;">
            <div id="map-controls-legend-phase3" style="display:none; flex-shrink:0;"><div id="legend-content-phase3"></div></div>
            <div id="p3-map-container" style="background:#fdfdfd; padding:15px; flex:1; overflow:auto; display:flex; justify-content:center; border:1px solid #ccc; border-radius:5px;"><p style="color:#aaa;">讀取中...</p></div>
        </div>
        <h4 style="color:#005a9e; margin-top:0;">📊 Master Join 大表預覽</h4><div id="p3-table-container"></div>
    `;
    document.getElementById('p3-table-container').innerHTML = renderTableHTML(p3Store.mergedData, p3Store.headers, 'phase3', false, false, 'filters');
    setTimeout(() => { renderP3MiniMaps(); drawPhase3Map(); }, 100);
}

window.syncP3Transforms = function() {
    let p3Conf = configStore.pns[configStore.currentPN].settings['phase3']; if (!p3Conf.legConfig) p3Conf.legConfig = {};
    const tableNameMap = { '1-1': 'BLK_Supply', '1-2': 'Prober', '1-3': 'Nikon', '1-4': 'AOI_Dim', '1-5': 'AOI_Defect', '1-6': 'Defect_Code', '1-7': 'WCM' };
    ['1-1', '1-2', '1-3', '1-4', '1-5', '1-7'].forEach(t => {
        let tConf = configStore.pns[configStore.currentPN].settings[t] || {}; let mapConf = tConf.mapConfig || {};
        p3Conf.transforms[t] = { dirX: mapConf.dirX || 'LR', dirY: mapConf.dirY || 'DU', rot: mapConf.rot || '0', offsetX: p3Conf.transforms[t]?.offsetX || 0, offsetY: p3Conf.transforms[t]?.offsetY || 0 };
        if (tConf.legConfig) { let prefix = tableNameMap[t]; Object.keys(tConf.legConfig).forEach(colName => { let p3ColName = `[${prefix}] ${colName}`; p3Conf.legConfig[p3ColName] = JSON.parse(JSON.stringify(tConf.legConfig[colName])); }); }
    });
    saveGlobalConfig(false); buildPhase3Controls(); alert("✅ 已將階段二的「XY方向、旋轉角度、圖例最大/最小值設定」全面同步至階段三！\n(請點擊「🚀 1. 執行 Master Join 融合」重新套用對齊)");
};

window.renderP3MiniMaps = function() {
    let p3Conf = configStore.pns[configStore.currentPN].settings['phase3']; let dims11 = getTableDimensions('1-1'); let cols11 = dims11 ? dims11.cols : 0; let rows11 = dims11 ? dims11.rows : 0;
    ['1-1', '1-2', '1-3', '1-4', '1-5', '1-7'].forEach(tabId => {
        let dimEl = document.getElementById(`p3-dim-${tabId}`);
        if (dimEl) {
            if (tabId === '1-3' || tabId === '1-7') { dimEl.innerHTML = `<span style="color:#8e44ad; font-weight:bold;">🔗 依序號綁定 (同1-1)</span>`; } 
            else { let dims = getTableDimensions(tabId); if (dims) { let isMatch = (dims.cols === cols11 && dims.rows === rows11); let color = isMatch ? '#1e8449' : '#d35400'; let icon = isMatch ? '✅' : '⚠️'; dimEl.innerHTML = `<span style="color:${color}; font-weight:bold;">${icon} 網格: ${dims.cols} x ${dims.rows}</span>`; } else { dimEl.innerHTML = `<span style="color:#aaa;">[無網格資料]</span>`; } }
        }
        let container = document.getElementById(`p3-mini-map-${tabId}`); if(!container) return;
        let tr = p3Conf.transforms[tabId] || { dirX: 'LR', dirY: 'DU', rot: '0' }; let tConf = configStore.pns[configStore.currentPN].settings[tabId]?.mapConfig || {}; let mc = { dirX: tConf.dirX || 'LR', dirY: tConf.dirY || 'DU', rot: tConf.rot || '0' };
        container.innerHTML = `<canvas id="cvs-mini-p3-${tabId}" width="200" height="200" style="width:100%; height:100%; object-fit:contain; background:#fff; border-radius:4px;"></canvas>`;
        let cvs = document.getElementById(`cvs-mini-p3-${tabId}`); let ctx = cvs.getContext('2d'); ctx.clearRect(0,0,200,200);
        let sourceCvs = document.getElementById(`cvs-${tabId}`);
        if (sourceCvs && sourceCvs.width > 0) {
            ctx.save(); ctx.translate(100, 100); ctx.rotate(parseInt(tr.rot) * Math.PI / 180); ctx.scale(tr.dirX === 'RL' ? -1 : 1, tr.dirY === 'UD' ? -1 : 1); ctx.scale(mc.dirX === 'RL' ? -1 : 1, mc.dirY === 'UD' ? -1 : 1); ctx.rotate(-parseInt(mc.rot) * Math.PI / 180);
            let scale = Math.min(180 / sourceCvs.width, 180 / sourceCvs.height); ctx.drawImage(sourceCvs, -sourceCvs.width/2 * scale, -sourceCvs.height/2 * scale, sourceCvs.width * scale, sourceCvs.height * scale); ctx.restore();
        } else { ctx.font = "14px Arial"; ctx.fillStyle = "#aaa"; ctx.textAlign = "center"; ctx.textBaseline = "middle"; ctx.fillText("請先至原表繪圖", 100, 100); }
    });
}

window.updateP3Transform = function(tabId, key, val) { let conf = configStore.pns[configStore.currentPN].settings['phase3']; if(!conf.transforms[tabId]) conf.transforms[tabId] = { dirX: 'LR', dirY: 'DU', rot: '0' }; conf.transforms[tabId][key] = val; saveGlobalConfig(false); }
window.updateP3Config = function(key, val) { configStore.pns[configStore.currentPN].settings['phase3'][key] = val; saveGlobalConfig(false); }

window.executeMasterJoin = function() {
    showLoading("執行 Master Join 中...");
    setTimeout(() => {
        let s1 = dataStore['1-1']; if(!s1 || !s1.mapGridData || s1.mapGridData.length === 0) { hideLoading(); alert("缺乏 1-1 實體工位檔網格，無法作為融合基準！"); return; }
        let p3Store = dataStore['phase3']; let p3Conf = configStore.pns[configStore.currentPN].settings['phase3'];
        p3Store.mergedData = []; p3Store.headers = []; p3Store.filters = {};

        let uX_11 = [...new Set(s1.mapGridData.map(g=>g[1]))].sort((a,b)=>a-b); let uY_11 = [...new Set(s1.mapGridData.map(g=>g[2]))].sort((a,b)=>b-a); 
        let max_c11 = uX_11.length - 1; let max_r11 = uY_11.length - 1;
        let xIdxMap11 = new Map(); uX_11.forEach((v,i)=>xIdxMap11.set(v,i)); let yIdxMap11 = new Map(); uY_11.forEach((v,i)=>yIdxMap11.set(v,i));

        let gridMap = {}; let coordToOrigNo = {}; 
        s1.mapGridData.forEach((g) => { let c = xIdxMap11.get(g[1]); let r = yIdxMap11.get(g[2]); let cellNo = String(g[0]).trim(); gridMap[`${c}_${r}`] = { no: cellNo, mapX: g[1], mapY: g[2], rowRef: g[3] }; coordToOrigNo[`${parseFloat(g[1])}_${parseFloat(g[2])}`] = cellNo; });

        let snakeReverseMap = {}; let store13 = dataStore['1-3'];
        if (store13 && store13.snakeGridMap) { Object.keys(store13.snakeGridMap).forEach(newNo => { let pt = store13.snakeGridMap[newNo]; snakeReverseMap[newNo] = coordToOrigNo[`${pt.x}_${pt.y}`]; }); }

        const tableNameMap = { '1-1': 'BLK_Supply', '1-2': 'Prober', '1-3': 'Nikon', '1-4': 'AOI_Dim', '1-5': 'AOI_Defect', '1-6': 'Defect_Code', '1-7': 'WCM' };
        let masterHeaders = ['[BLK_Supply] No.', '[BLK_Supply] Map_X', '[BLK_Supply] Map_Y', '[BLK_Supply] Row_Ref']; let tablesToJoin = ['1-2', '1-3', '1-4', '1-5', '1-7']; let tableDataMaps = {};

        tablesToJoin.forEach(tabId => {
            let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId]; if(!store || store.mergedData.length === 0) return;
            let bData = store.mergedData; let bHeaders = store.headers;
            if (tabId === '1-3' && conf.pivotConfig && conf.pivotConfig.enabled) { bData = store.pivotedData; bHeaders = store.pivotedHeaders; }
            if (tabId === '1-5' && store.joinedHeaders && store.joinedData) { bData = store.joinedData; bHeaders = store.joinedHeaders; }
            if (tabId === '1-2' && conf.mapConfig.isCombined && store.combinedData1_2) { bData = store.combinedData1_2; bHeaders = store.combinedHeaders1_2; }

            let prefix = tableNameMap[tabId] || tabId; masterHeaders.push(...bHeaders.map(h => `[${prefix}] ${h}`));
            let tMap = {}; let tr = p3Conf.transforms[tabId] || { dirX: 'LR', dirY: 'DU', rot: '0', offsetX: 0, offsetY: 0 }; let mc = conf.mapConfig || {};
            let selectedWafer = mc.selectedWafer; let waferIdx = bHeaders.findIndex(h => h && h.toUpperCase().replace(/[\s_]/g, '') === 'WAFERID');

            if (tabId === '1-3' || tabId === '1-7') {
                let noCol = tabId === '1-3' ? 'New_No' : 'No.'; let noIdx = bHeaders.findIndex(h => h && h.toUpperCase() === noCol.toUpperCase());
                if(noIdx > -1) {
                    bData.forEach(row => { 
                        if(selectedWafer && waferIdx > -1 && String(row[waferIdx]).trim() !== selectedWafer) return;
                        let key = String(row[noIdx]).trim();
                        if (tabId === '1-3') { if (Object.keys(snakeReverseMap).length > 0) { let realOrigNo = snakeReverseMap[key]; if (realOrigNo) tMap[realOrigNo] = row; } else { tMap[key] = row; } } 
                        else { tMap[key] = row; }
                    });
                }
                tableDataMaps[tabId] = { type: 'NO', map: tMap, len: bHeaders.length };
            } else {
                let xCol = conf.mapConfig.xCol; let yCol = conf.mapConfig.yCol; let xIdx = bHeaders.indexOf(xCol); let yIdx = bHeaders.indexOf(yCol);
                if(xIdx > -1 && yIdx > -1) {
                    let filteredData = bData.filter(r => { if(selectedWafer && waferIdx > -1 && String(r[waferIdx]).trim() !== selectedWafer) return false; return true; });
                    let uX = [...new Set(filteredData.map(r=>parseFloat(r[xIdx])))].filter(v=>!isNaN(v)).sort((a,b)=>a-b); let uY = [...new Set(filteredData.map(r=>parseFloat(r[yIdx])))].filter(v=>!isNaN(v)).sort((a,b)=>b-a);
                    let xIdxMap = new Map(); uX.forEach((v,i)=>xIdxMap.set(v,i)); let yIdxMap = new Map(); uY.forEach((v,i)=>yIdxMap.set(v,i));
                    let max_c = uX.length - 1; let max_r = uY.length - 1;

                    filteredData.forEach(row => {
                        let x = parseFloat(row[xIdx]); let y = parseFloat(row[yIdx]); if(isNaN(x) || isNaN(y)) return;
                        let c = xIdxMap.get(x); let r = yIdxMap.get(y);
                        if(tr.dirX === 'RL') c = max_c - c; if(tr.dirY === 'UD') r = max_r - r;
                        let fc = c, fr = r; if(tr.rot === '90') { fc = max_r - r; fr = c; } else if(tr.rot === '180') { fc = max_c - c; fr = max_r - r; } else if(tr.rot === '270') { fc = r; fr = max_c - c; }
                        fc += (tr.offsetX || 0); fr += (tr.offsetY || 0); tMap[`${fc}_${fr}`] = row;
                    });
                }
                tableDataMaps[tabId] = { type: 'CR', map: tMap, len: bHeaders.length };
            }
        });

        p3Store.headers = ['Master_No', 'Master_Col', 'Master_Row', ...masterHeaders];

        for(let r=0; r<=max_r11; r++) {
            for(let c=0; c<=max_c11; c++) {
                let cell = gridMap[`${c}_${r}`];
                if(cell) {
                    let newRow = [cell.no, c, r, cell.no, cell.mapX, cell.mapY, cell.rowRef];
                    tablesToJoin.forEach(tabId => {
                        let tMeta = tableDataMaps[tabId];
                        if(tMeta) { let matchRow = (tMeta.type === 'NO') ? tMeta.map[cell.no] : tMeta.map[`${c}_${r}`]; if(matchRow) newRow.push(...matchRow); else newRow.push(...new Array(tMeta.len).fill('')); }
                    });
                    p3Store.mergedData.push(newRow);
                }
            }
        }
        
        buildPhase3Controls(); hideLoading();
    }, 100);
}

window.exportP3CSV = function(isUnpivot) {
    let store = dataStore['phase3']; if (!store || store.mergedData.length === 0) { alert("沒有可匯出的資料！請先執行 Master Join。"); return; }
    let headers = store.headers; let data = store.mergedData;

    if (isUnpivot) {
        let keyStr = prompt("轉置為直表 - 請輸入「固定不變的鍵值」欄位索引 (用逗號分隔)：\\n(預設 0~6 為 Master_No 到 Row_Ref，不熟悉建議使用預設值)", "0,1,2,3,4,5,6");
        if (keyStr === null) return; let keyIdxs = keyStr.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n)); if (keyIdxs.length === 0) keyIdxs = [0,1,2,3,4,5,6];
        let unpivotHeaders = []; keyIdxs.forEach(i => unpivotHeaders.push(headers[i])); unpivotHeaders.push("Item", "Value");
        let unpivotData = [];
        data.forEach(row => {
            let baseCols = keyIdxs.map(i => row[i]);
            for (let i = 0; i < headers.length; i++) { if (!keyIdxs.includes(i)) { let val = row[i]; if (val !== undefined && val !== '') { unpivotData.push([...baseCols, headers[i], val]); } } }
        });
        headers = unpivotHeaders; data = unpivotData;
    }

    let csvContent = "\uFEFF"; csvContent += headers.map(h => `"${String(h).replace(/"/g, '""')}"`).join(",") + "\n";
    data.forEach(row => { csvContent += row.map(v => `"${String(v !== undefined ? v : '').replace(/"/g, '""')}"`).join(",") + "\n"; });
    let blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' }); let url = URL.createObjectURL(blob); let a = document.createElement("a"); a.href = url;
    a.download = `MasterJoin_${isUnpivot ? '直表' : '橫表'}.csv`; a.click(); URL.revokeObjectURL(url);
};

window.drawPhase3Map = function(skipStatsRecalc = false) {
    let p3Store = dataStore['phase3']; let conf = configStore.pns[configStore.currentPN].settings['phase3']; window.activeMapRects = []; 
    let renderArea = document.getElementById('p3-map-container'); if(!renderArea) return;
    
    let valEl = document.getElementById('p3-val-col'); let vCol = conf.valCol || (valEl ? valEl.value : null);
    if (!vCol && p3Store.headers.length > 0) vCol = p3Store.headers[p3Store.headers.length - 1]; 
    if (!vCol) { renderArea.innerHTML = '<p style="color:#aaa;">請先選擇顯示數值</p>'; return; }
    let vIdx = p3Store.headers.indexOf(vCol); if(vIdx===-1) return;

    showLoading("正在渲染 Master MAP...");
    setTimeout(() => {
        try {
            let points = []; let xSet = new Set(), ySet = new Set();
            let isCategorical = false; let catSet = new Set(); let rawValues = []; let uniqueVals = new Set();
            let xEl = document.getElementById('p3-x-col'); let yEl = document.getElementById('p3-y-col');
            let xCol = conf.xCol || (xEl ? xEl.value : '[BLK_Supply] Map_X'); let yCol = conf.yCol || (yEl ? yEl.value : '[BLK_Supply] Map_Y');
            let cIdx = p3Store.headers.indexOf(xCol); let rIdx = p3Store.headers.indexOf(yCol); let noIdx = p3Store.headers.indexOf('Master_No');

            let selWafer = conf.selectedWafer; let waferColIdxs = p3Store.headers.map((h,i) => h && h.toUpperCase().includes('WAFERID') ? i : -1).filter(i=>i>-1);

            p3Store.mergedData.forEach(r => {
                if(selWafer && waferColIdxs.length > 0) { let match = waferColIdxs.some(i => String(r[i]).trim() === selWafer); if(!match) return; }
                let x = parseFloat(r[cIdx]); let y = parseFloat(r[rIdx]); let rawV = String(r[vIdx] !== undefined ? r[vIdx] : '').trim(); let v = parseFloat(rawV);
                if (!isNaN(x) && !isNaN(y)) { points.push({no: r[noIdx], x, y, v: isNaN(v) ? rawV : v, rawV: rawV}); xSet.add(x); ySet.add(y); if(rawV !== '') uniqueVals.add(rawV); }
            });

            if(points.length === 0) { renderArea.innerHTML = `<p style="color:#d9534f; font-weight:bold;">⚠️ 無法渲染：在此 X 座標與 Y 座標組合下找不到有效數值。</p>`; return; }

            let hasText = Array.from(uniqueVals).some(v => isNaN(parseFloat(v)) || String(v).trim() === '');
            isCategorical = (vIdx === -1) ? true : ((hasText && uniqueVals.size <= 50) || (uniqueVals.size <= 15 && uniqueVals.size > 0));
            let vColUpper = String(vCol).toUpperCase();
            if (['MASTER_NO', 'MASTER_COL', 'MASTER_ROW'].includes(vColUpper)) { isCategorical = false; } else if (vColUpper.includes('NO')) { isCategorical = false; }
            if (uniqueVals.size > 50) isCategorical = false;

            points.forEach(p => { if(p.rawV !== '') { if(isCategorical) catSet.add(p.rawV); else { let pv = parseFloat(p.rawV); if(!isNaN(pv)) rawValues.push(pv); } } });

            let stats = null; let uiMax = NaN, uiMin = NaN; let legCnt = 10, legInt = 1;
            if (!isCategorical) {
                stats = calculateStats(rawValues); let vConf = conf.legConfig && conf.legConfig[vCol] ? conf.legConfig[vCol] : null;
                if (!skipStatsRecalc && (!vConf || vConf.max === undefined)) { uiMax = stats ? stats.upperBound : NaN; uiMin = stats ? stats.lowerBound : NaN; legInt = ((uiMax - uiMin) / legCnt) || 1; } 
                else if (skipStatsRecalc && document.getElementById(`leg-max-phase3`)) { uiMax = parseFloat(document.getElementById(`leg-max-phase3`).value); uiMin = parseFloat(document.getElementById(`leg-min-phase3`).value); legCnt = parseInt(document.getElementById(`leg-cnt-phase3`).value) || 10; legInt = parseFloat(document.getElementById(`leg-int-phase3`).value) || ((uiMax-uiMin)/legCnt) || 1; } 
                else if (vConf) { uiMax = vConf.max; uiMin = vConf.min; legCnt = vConf.cnt || 10; legInt = vConf.int || ((uiMax - uiMin) / legCnt) || 1; }
                buildLegendUI('phase3', stats, uiMax, uiMin, false, legCnt, legInt);
            } else { buildLegendUI('phase3', null, NaN, NaN, true); }

            let catColorMap = {}; let getRainbowColor; let legHTML = '';
            if (isCategorical) {
                let colors = ['#3498db','#f1c40f','#9b59b6','#e67e22','#1abc9c','#34495e','#ff69b4','#8a2be2','#a52a2a','#d2691e','#008080','#4682b4']; let cIdx = 0; let catArray = Array.from(catSet); let legendLimit = Math.min(catArray.length, 50);
                for(let i=0; i<catArray.length; i++) { let c = catArray[i]; let pre = window.getCatColor(c); if(pre) catColorMap[c] = pre; else { catColorMap[c] = colors[cIdx % colors.length]; cIdx++; } }
                legHTML = `<div class="legend-container" style="justify-content:flex-start;">`;
                for(let i=0; i<legendLimit; i++) { let k = catArray[i]; legHTML += `<div class="cat-legend-item"><div class="cat-color-box" style="background:${catColorMap[k]}"></div>${String(k).substring(0,10)}</div>`; }
                if (catArray.length > 50) legHTML += `<div style="font-size:10px; color:#888; margin-top:5px;">...等 ${catArray.length} 項</div>`; legHTML += `</div>`;
            } else {
                getRainbowColor = (val) => { if(val > uiMax) val = uiMax; if(val < uiMin) val = uiMin; let pct = (uiMax===uiMin) ? 0.5 : (val - uiMin) / (uiMax - uiMin); let hue = (1.0 - pct) * 240; return `hsl(${hue}, 100%, 50%)`; };
                legHTML = `<div class="legend-container"><div style="font-size:11px; margin-bottom:5px; font-weight:bold;">Max: ${isNaN(uiMax)?'N/A':uiMax.toFixed(2)}</div><div class="legend-bar"></div><div style="font-size:11px; margin-top:5px; font-weight:bold;">Min: ${isNaN(uiMin)?'N/A':uiMin.toFixed(2)}</div></div>`;
            }

            let uX = Array.from(xSet).sort((a,b)=>a-b); let uY = Array.from(ySet).sort((a,b)=>b-a); 
            let xMap = new Map(); uX.forEach((v, i) => xMap.set(v, i)); let yMap = new Map(); uY.forEach((v, i) => yMap.set(v, i));
            let cols = uX.length; let rows = uY.length;
            let cvsWidth = Math.min(700, Math.max(300, cols * 10)); let cvsHeight = Math.min(700, Math.max(300, rows * 10)); let cellW = cvsWidth / (cols||1); let cellH = cvsHeight / (rows||1);

            renderArea.innerHTML = `<div style="display:flex; align-items:flex-start;"><canvas id="cvs-phase3" width="${cvsWidth}" height="${cvsHeight}" style="border:1px solid #ccc; background:#e6e6e6; cursor:crosshair;"></canvas>${legHTML}</div>`;
            document.getElementById('map-controls-legend-phase3').style.display = 'block';
            let cvs = document.getElementById(`cvs-phase3`); let ctx = cvs.getContext('2d'); ctx.clearRect(0, 0, cvsWidth, cvsHeight); ctx.font = "8px Arial"; ctx.textAlign = "center"; ctx.textBaseline = "middle";

            points.forEach(p => {
                let px = xMap.get(p.x) * cellW; let py = yMap.get(p.y) * cellH;
                let pointColor = '#e6e6e6'; if (isCategorical) { if(p.rawV !== '') pointColor = catColorMap[p.rawV] || '#e6e6e6'; } else { if(p.rawV !== '' && !isNaN(p.v)) pointColor = getRainbowColor(p.v); }
                ctx.fillStyle = pointColor; ctx.fillRect(px, py, cellW, cellH);
                if (cellW >= 1.5 && cellH >= 1.5) { ctx.strokeStyle = "rgba(255,255,255,0.3)"; ctx.lineWidth = 1; ctx.strokeRect(px, py, cellW, cellH); }
                if (cellW > 20 && cellH > 10 && p.rawV !== '') { ctx.fillStyle = (isCategorical || p.v > (uiMax+uiMin)/2) ? "#fff" : "#000"; let dispV = isCategorical ? String(p.rawV).substring(0,3) : p.v.toFixed(1); ctx.fillText(dispV, px + cellW/2, py + cellH/2); }
                window.activeMapRects.push({ tabId: 'phase3', x:px, y:py, w:cellW, h:cellH, str:`[No: ${p.no}]<br>X: ${p.x}, Y: ${p.y}<br>${vCol}: <b style="color:#f1c40f;">${p.rawV}</b>` });
            });
        } catch (err) { console.error(err); renderArea.innerHTML = `<p style="color:#d9534f; font-weight:bold;">⚠️ Master Join MAP 渲染失敗: ${err.message}</p>`; } finally { hideLoading(); }
    }, 50);
}

function parseCSVLineFast(line, delimiter) { if (line.indexOf('"') === -1) return line.split(delimiter).map(s => s.trim()); let result = []; let current = ''; let inQuotes = false; for (let i = 0; i < line.length; i++) { let char = line[i]; if (char === '"') { if (inQuotes && line[i+1] === '"') { current += '"'; i++; } else { inQuotes = !inQuotes; } } else if (char === delimiter && !inQuotes) { result.push(current); current = ''; } else { current += char; } } result.push(current.trim()); return result; }
function readFileAs2DArray(file, config, tabId) {
    return new Promise((resolve, reject) => {
        let reader = new FileReader();
        reader.onload = (e) => {
            let data = new Uint8Array(e.target.result); let ext = file.name.split('.').pop().toLowerCase();
            if (['csv', 'txt', 'dat', 'tsv'].includes(ext)) {
                let text = ''; let encMode = config.encoding || 'auto';
                if (encMode === 'auto') { text = new TextDecoder('utf-8').decode(data); if (text.includes('\uFFFD')) text = new TextDecoder('big5').decode(data); } else text = new TextDecoder(encMode).decode(data);
                let delimiter = text.slice(0, 500).indexOf('\t') > -1 ? '\t' : ','; let lines = text.split(/\r\n|\n|\r/); resolve(lines.filter(l => l.trim() !== '').map(l => parseCSVLineFast(l, delimiter)));
            } else {
                try { let workbook = XLSX.read(data, { type: 'array' }); if (tabId && !dataStore[tabId].sheetNamesExtracted) { let sel = document.getElementById(`sheet-${tabId}`); sel.innerHTML = '<option value="">-- 自動讀取 --</option>'; workbook.SheetNames.forEach(name => sel.appendChild(new Option(name, name))); dataStore[tabId].sheetNamesExtracted = true; } let sName = config.sheetName && workbook.SheetNames.includes(config.sheetName) ? config.sheetName : workbook.SheetNames[0]; resolve(XLSX.utils.sheet_to_json(workbook.Sheets[sName], { header: 1, defval: "" })); } catch(err) { reject(err); }
            }
        }; reader.onerror = reject; reader.readAsArrayBuffer(file);
    });
}
function matchPattern(filename, patternStr) { return patternStr.split(',').map(s => s.trim().replace(/\*/g, '.*')).some(p => new RegExp(`^${p}$`, 'i').test(filename)); }

async function handleFolderSelect(event, tabId) {
    let files = event.target.files; if (files.length === 0) return; let conf = configStore.pns[configStore.currentPN].settings[tabId];
    let matchedFiles = Array.from(files).filter(f => matchPattern(f.name, conf.filePattern || '*')); let countSpan = document.getElementById(`file-count-${tabId}`);
    if (matchedFiles.length > 0) { let relPath = matchedFiles[0].webkitRelativePath.replace(/\/[^\/]*$/, ''); let pathInput = document.getElementById(`folder-path-${tabId}`); if (!pathInput.value || !pathInput.value.includes(':\\')) { pathInput.value = relPath; conf.lastFolderPath = relPath; saveGlobalConfig(false); } }
    if (matchedFiles.length === 0) { countSpan.innerText = `(找到 0 個)`; countSpan.style.color = "#d9534f"; alert("找不到檔案！"); return; }
    countSpan.innerText = `(掃描到 ${matchedFiles.length} 檔)`; countSpan.style.color = "#28a745"; dataStore[tabId].matchedFiles = matchedFiles; await processFiles(tabId); event.target.value = ''; 
}
async function processFiles(tabId) {
    showLoading("資料解析中..."); await new Promise(r => setTimeout(r, 50)); 
    let matchedFiles = dataStore[tabId].matchedFiles; let conf = configStore.pns[configStore.currentPN].settings[tabId]; let parsedFiles = []; dataStore[tabId].sheetNamesExtracted = false; 
    for (let file of matchedFiles) { try { let dataArr = await readFileAs2DArray(file, conf, tabId); if (dataArr.length > 0) parsedFiles.push({ filePath: file.webkitRelativePath, fileName: file.name, rows: dataArr }); } catch (e) {} }
    dataStore[tabId].rawFiles = parsedFiles; 
    
    // Auto-detect header or starting row
    if (conf.headerRowIndex === null && parsedFiles.length > 0) { 
        if (tabId === '1-1') {
            let pIdx = parsedFiles[0].rows.findIndex(r => r.join('').toUpperCase().includes("PARA2"));
            conf.headerRowIndex = pIdx > -1 ? pIdx : 0;
        } else if (tabId === '1-7') {
            conf.headerRowIndex = 0;
        } else {
            let firstRows = parsedFiles[0].rows; let bestIdx = 0; let maxFill = 0; 
            for (let i = 0; i < Math.min(firstRows.length, 30); i++) { let fills = firstRows[i].filter(c => c && String(c).trim() !== '').length; if (fills > maxFill) { maxFill = fills; bestIdx = i; } } 
            conf.headerRowIndex = bestIdx; 
        }
    }
    
    calculateMergedData(tabId); renderDerivedRules(tabId); if(tabId==='1-3') renderPivotOptions(); renderPreview(tabId); hideLoading();
}
function handlePaste(event, tabId) { showLoading("資料解析中..."); setTimeout(() => { let text = event.target.value; if (!text.trim()) { hideLoading(); return; } let delimiter = text.slice(0, 200).indexOf('\t') > -1 ? '\t' : ','; let parsed = text.split(/\r\n|\n/).filter(l => l.trim() !== '').map(l => parseCSVLineFast(l, delimiter)); dataStore[tabId].rawFiles = [{ filePath: "Pasted", fileName: "Pasted", rows: parsed }]; calculateMergedData(tabId); renderDerivedRules(tabId); renderPreview(tabId); hideLoading(); }, 50); }

function applyDerivedFields(conf, fullRow) {
    let derivedData = [];
    (conf.derivedFields || []).forEach(rule => {
        let sourceVal = String(fullRow[rule.sourceIdx] !== undefined ? fullRow[rule.sourceIdx] : ''); let res = "";
        try { if (rule.method === 'mid') { let start = parseInt(rule.p1) || 1; let len = parseInt(rule.p2) || sourceVal.length; res = sourceVal.substr(Math.max(0, start - 1), len); } else if (rule.method === 'split') { let parts = sourceVal.split(rule.p1 || ','); let idx = Math.max(0, (parseInt(rule.p2) || 1) - 1); res = parts[idx] || ""; } else if (rule.method === 'regex') { let match = sourceVal.match(new RegExp(rule.p1)); res = match ? (match[1] || match[0]) : ""; } } catch(e) {} derivedData.push(res);
    }); return derivedData;
}

function calculateMergedData(tabId) {
    let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId]; let tableInfo = tablesConfig.find(t => t.id === tabId); if(!conf.customHeaders) conf.customHeaders = {};
    store.mergedData = []; store.headers = []; store.mapGridData = []; if (store.rawFiles.length === 0) return;

    let isTable1_1 = (tabId === '1-1'); let isTable1_7 = (tabId === '1-7'); let hIdx = conf.headerRowIndex || 0; let baseHeaders = [];
    
    if (isTable1_1) { 
        let maxCols = store.rawFiles[0].rows.reduce((max, r) => Math.max(max, r.length), 0); for(let c=0; c<maxCols; c++) baseHeaders.push(`Raw_${c+1}`); baseHeaders.push("Count", "Pitch_X", "Pitch_Y", "Start_X", "Start_Y"); 
    } else if (isTable1_7) { 
        let rawHeaders = store.rawFiles[0].rows[hIdx] || [];
        if (rawHeaders.length > 0 && rawHeaders[0].toUpperCase() !== 'NO.' && rawHeaders[0].toUpperCase() !== 'NO') {
            baseHeaders = ["No.", "Status"]; let maxCols = store.rawFiles[0].rows.reduce((max, r) => Math.max(max, r.length), 0); for(let c=2; c<maxCols; c++) baseHeaders.push(`Col_${c+1}`); 
        } else {
            baseHeaders = rawHeaders;
        }
    } else { 
        baseHeaders = store.rawFiles[0].rows[hIdx] || []; 
    }
    
    let derivedHeaders = (conf.derivedFields || []).map(d => d.newColName || '未命名'); store.headers = ["FilePath", "FileName", ...baseHeaders, ...derivedHeaders]; store.headers = store.headers.map((h, i) => conf.customHeaders[i] !== undefined ? conf.customHeaders[i] : h);
    const pad = (v) => { if(!v) return "0000"; let s = String(v).trim(); if(!isNaN(s)) return String(Math.round(parseFloat(s))).padStart(4, '0'); return s.padStart(4,'0'); };
    let noCounter1_1 = 1;

    store.rawFiles.forEach(f => {
        let dataRows = f.rows;
        if (isTable1_1) {
            let maxCols = baseHeaders.length - 5; 
            for (let i = hIdx + 1; i < dataRows.length; i++) {
                let r = dataRows[i]; if (!r || r.length < 10) continue;
                let count = parseInt((String(r[1] || '').trim()).slice(-4)); if (isNaN(count) || count <= 0) continue;
                let pX = parseFloat(pad(r[2]).slice(-4) + pad(r[3]).slice(-4)) / 100; let pY = parseFloat(pad(r[4]).slice(-4) + pad(r[5]).slice(-4)) / 100;
                let sX = parseFloat(pad(r[6]).slice(-4) + pad(r[7]).slice(-4)) / 100; let sY = parseFloat(pad(r[8]).slice(-4) + pad(r[9]).slice(-4)) / 100;
                if (isNaN(pX) || isNaN(sX)) continue;
                let paddedR = [...r]; while(paddedR.length < maxCols) paddedR.push('');
                let fullRow = [f.filePath, f.fileName, ...paddedR, count, pX, pY, sX, sY]; store.mergedData.push([...fullRow, ...applyDerivedFields(conf, fullRow)]);
                for(let k=0; k<count; k++) { let cx = parseFloat((sX + k*pX).toFixed(3)); let cy = parseFloat((sY + k*pY).toFixed(3)); store.mapGridData.push([noCounter1_1++, cx, cy, store.mergedData.length]); }
            }
        } else if (isTable1_7) {
            if (dataRows.length > 0 && String(dataRows[dataRows.length - 1][0]).toLowerCase().trim() === 'end') dataRows = dataRows.slice(0, -1);
            for(let i = hIdx + 1; i < dataRows.length; i++) { let r = dataRows[i]; if(r.length === 0 || r[0] === '') continue; let fullRow = [f.filePath, f.fileName, ...r]; store.mergedData.push([...fullRow, ...applyDerivedFields(conf, fullRow)]); }
        } else {
            for (let i = hIdx + 1; i < dataRows.length; i++) { let r = dataRows[i]; if (r.length === 0 || r.every(c => c === '' || c === undefined)) continue; let fullRow = [f.filePath, f.fileName, ...r]; store.mergedData.push([...fullRow, ...applyDerivedFields(conf, fullRow)]); }
        }
    });

    if (tabId === '1-3' && conf.pivotConfig && conf.pivotConfig.enabled) {
        let waferIdx = store.headers.findIndex(h => h.toUpperCase() === 'WAFERID'); let keyIdx = store.headers.indexOf(conf.pivotConfig.keyCol); let valIdx = store.headers.indexOf(conf.pivotConfig.valCol);
        if (waferIdx === -1 || keyIdx === -1 || valIdx === -1) { alert("⚠️ 轉置失敗：找不到 WaferID, Key 或 Value 欄位！請確認設定。"); conf.pivotConfig.enabled = false; return; }
        let uniqueKeys = []; store.mergedData.forEach(r => { let k = r[keyIdx]; if(k && !uniqueKeys.includes(k)) uniqueKeys.push(k); });
        let pivotedData = []; let currentWafer = null; let currentRecord = null; let noCount = 1; let newHeaders = ["FilePath", "FileName", "WaferID", "No.", ...uniqueKeys];
        store.mergedData.forEach(r => {
            let wId = r[waferIdx]; let key = r[keyIdx]; let val = r[valIdx];
            if(wId !== currentWafer) { if(currentRecord) pivotedData.push(currentRecord); currentWafer = wId; noCount = 1; currentRecord = new Array(newHeaders.length).fill(''); currentRecord[0] = r[0]; currentRecord[1] = r[1]; currentRecord[2] = wId; currentRecord[3] = noCount; }
            let kIdx = uniqueKeys.indexOf(key);
            if(kIdx !== -1) { if(currentRecord[4 + kIdx] !== '') { pivotedData.push(currentRecord); noCount++; currentRecord = new Array(newHeaders.length).fill(''); currentRecord[0] = r[0]; currentRecord[1] = r[1]; currentRecord[2] = wId; currentRecord[3] = noCount; } currentRecord[4 + kIdx] = val; }
        });
        if(currentRecord) pivotedData.push(currentRecord); store.pivotedData = pivotedData; store.pivotedHeaders = newHeaders;
    }
}

function updatePivotConfig(tabId) { let conf = configStore.pns[configStore.currentPN].settings[tabId]; conf.pivotConfig.enabled = document.getElementById(`pivot-enable-${tabId}`).checked; conf.pivotConfig.keyCol = document.getElementById(`pivot-key-${tabId}`).value; conf.pivotConfig.valCol = document.getElementById(`pivot-val-${tabId}`).value; saveGlobalConfig(false); showLoading("正在轉置..."); setTimeout(()=>{calculateMergedData(tabId); renderPreview(tabId); hideLoading();}, 50); }
function renderPivotOptions() { let tabId = '1-3'; let conf = configStore.pns[configStore.currentPN].settings[tabId]; let store = dataStore[tabId]; let keySel = document.getElementById(`pivot-key-${tabId}`); let valSel = document.getElementById(`pivot-val-${tabId}`); if(!keySel || !valSel) return; let opts = store.headers.map(h => `<option value="${h}">${h}</option>`).join(''); keySel.innerHTML = opts; valSel.innerHTML = opts; if(conf.pivotConfig.keyCol) keySel.value = conf.pivotConfig.keyCol; if(conf.pivotConfig.valCol) valSel.value = conf.pivotConfig.valCol; document.getElementById(`pivot-enable-${tabId}`).checked = conf.pivotConfig.enabled; }

window.selectHeaderRow = function(tabId, rowIndex) { 
    showLoading("重新套用標題列與合併資料..."); 
    setTimeout(() => { 
        configStore.pns[configStore.currentPN].settings[tabId].headerRowIndex = rowIndex; 
        saveGlobalConfig(false);
        calculateMergedData(tabId); renderDerivedRules(tabId); if(tabId==='1-3') renderPivotOptions(); renderPreview(tabId); hideLoading(); 
    }, 50); 
};

window.editHeader = function(tabId, colIndex) { let conf = configStore.pns[configStore.currentPN].settings[tabId]; if(!conf.customHeaders) conf.customHeaders = {}; let newName = prompt(`✏️ 請輸入 [第 ${colIndex+1} 欄] 的新標題名稱：`, dataStore[tabId].headers[colIndex]); if(newName !== null && newName.trim() !== "") { conf.customHeaders[colIndex] = newName.trim(); saveGlobalConfig(false); calculateMergedData(tabId); renderDerivedRules(tabId); if(tabId==='1-3') renderPivotOptions(); renderPreview(tabId); } }
function addDerivedField(tabId) { let conf = configStore.pns[configStore.currentPN].settings[tabId]; if (!conf.derivedFields) conf.derivedFields = []; conf.derivedFields.push({ newColName: 'WaferID', sourceIdx: 1, method: 'split', p1: '.', p2: '1' }); calculateMergedData(tabId); renderDerivedRules(tabId); renderPreview(tabId); }
function removeDerivedField(tabId, index) { configStore.pns[configStore.currentPN].settings[tabId].derivedFields.splice(index, 1); calculateMergedData(tabId); renderDerivedRules(tabId); renderPreview(tabId); }
function updateDerivedRule(tabId, index, key, value) { configStore.pns[configStore.currentPN].settings[tabId].derivedFields[index][key] = value; calculateMergedData(tabId); renderDerivedRules(tabId); renderPreview(tabId); }

function renderDerivedRules(tabId) {
    let container = document.getElementById(`derived-rules-${tabId}`); if (!container) return; let conf = configStore.pns[configStore.currentPN].settings[tabId]; let store = dataStore[tabId]; let headers = store.headers.slice(0, store.headers.length - (conf.derivedFields||[]).length); let html = '';
    (conf.derivedFields || []).forEach((rule, idx) => {
        let optHtml = headers.map((h, i) => `<option value="${i}" ${rule.sourceIdx == i ? 'selected' : ''}>[${i+1}] ${h || '未命名'}</option>`).join('');
        let p1Label = rule.method === 'split' ? '分隔符號:' : (rule.method === 'regex' ? 'Regex:' : '起始字元:'); let p2Label = rule.method === 'split' ? '字節/索引:' : (rule.method === 'regex' ? '' : '長度:');
        let p2UI = rule.method === 'regex' ? '' : `<span style="color:#666; font-size:12px; margin-left:10px;">${p2Label}</span><input type="text" style="width:60px; margin-left:5px;" value="${rule.p2}" onchange="updateDerivedRule('${tabId}', ${idx}, 'p2', this.value)">`;
        let sampleRes = "-"; if (store.mergedData.length > 0) { let sampleRow = store.mergedData[0]; let sourceVal = String(sampleRow[rule.sourceIdx] !== undefined ? sampleRow[rule.sourceIdx] : ''); try { if (rule.method === 'mid') { sampleRes = sourceVal.substr(Math.max(0, (parseInt(rule.p1)||1) - 1), parseInt(rule.p2)||sourceVal.length); } else if (rule.method === 'split') { sampleRes = sourceVal.split(rule.p1||',')[Math.max(0, (parseInt(rule.p2)||1) - 1)] || ""; } else if (rule.method === 'regex') { let m = sourceVal.match(new RegExp(rule.p1)); sampleRes = m ? (m[1]||m[0]) : ""; } } catch(e){} }
        html += `<div class="rule-row"><input type="text" style="width:100px; font-weight:bold; color:#005a9e;" value="${rule.newColName}" onchange="updateDerivedRule('${tabId}', ${idx}, 'newColName', this.value)" placeholder="新欄位名"><span>=</span><select style="width:160px;" onchange="updateDerivedRule('${tabId}', ${idx}, 'sourceIdx', this.value)">${optHtml}</select><select onchange="updateDerivedRule('${tabId}', ${idx}, 'method', this.value); renderDerivedRules('${tabId}');"><option value="mid" ${rule.method==='mid'?'selected':''}>字串擷取 (Mid)</option><option value="split" ${rule.method==='split'?'selected':''}>字串分割 (Split)</option><option value="regex" ${rule.method==='regex'?'selected':''}>正則提取 (Regex)</option></select><span style="color:#666; font-size:12px; margin-left:10px;">${p1Label}</span><input type="text" style="width:60px; margin-left:5px;" value="${rule.p1}" onchange="updateDerivedRule('${tabId}', ${idx}, 'p1', this.value)">${p2UI}<span class="live-preview" title="即時預覽">預覽: ${sampleRes}</span><button class="btn-danger" style="margin-left:auto; padding:3px 8px; border-radius:3px; cursor:pointer;" onclick="removeDerivedField('${tabId}', ${idx})">刪除</button></div>`;
    }); container.innerHTML = html;
}

function renderTableHTML(dataArray, headers, tabId, noWrap = false, disableFilter = false, filterKey = 'filters') {
    if(dataArray.length===0) return '<p style="color:#666;">無資料</p>';
    let store = dataStore[tabId]; let filterMap = store[filterKey] || {};
    let filteredData = dataArray.filter(r => { for (let colIdx in filterMap) { if (!filterMap[colIdx].has(String(r[colIdx]))) return false; } return true; });

    let html = `<table class="preview-table"><thead><tr><th># 行號</th>`; 
    headers.forEach((h, i) => html += `<th class="editable-header" ${disableFilter ? '' : `ondblclick="editHeader('${tabId}', ${i})"`}>${h || '未命名'} ${disableFilter ? '' : `<span class="filter-icon" onclick="openFilter('${tabId}', ${i}, this, '${filterKey}')" title="篩選">🔍</span>`}</th>`); 
    html += '</tr></thead><tbody>';
    
    let topLimit = Math.min(30, filteredData.length);
    for(let i=0; i<topLimit; i++){ html += `<tr><td style="color:#888; background:#f0f0f0;">${i + 1}</td>`; filteredData[i].forEach(c => html += `<td>${c!==undefined?c:''}</td>`); html += '</tr>'; }
    if (filteredData.length > 35) { html += `<tr class="skip-row"><td colspan="${headers.length + 1}">... (省略中間 ${filteredData.length - 35} 筆) ...</td></tr>`; }
    if (filteredData.length > 30) { let bottomStart = Math.max(30, filteredData.length - 5); for(let i=bottomStart; i<filteredData.length; i++) { html += `<tr><td style="color:#888; background:#f0f0f0;">${i + 1}</td>`; filteredData[i].forEach(c => html += `<td>${c!==undefined?c:''}</td>`); html += '</tr>'; } }
    html += '</tbody></table>'; 
    if(filteredData.length !== dataArray.length) html += `<div style="color:#d35400; font-weight:bold; margin-top:5px;">⚠️ 已套用篩選：顯示 ${filteredData.length} / ${dataArray.length} 筆</div>`;
    
    if(noWrap) return html; 
    let wrapperHtml = `<div class="table-wrapper"><div class="top-scroll-wrapper" id="tsw-${tabId}-${filterKey}" onscroll="document.getElementById('bsw-${tabId}-${filterKey}').scrollLeft = this.scrollLeft;"><div class="top-scroll-dummy" id="tsd-${tabId}-${filterKey}"></div></div><div class="preview-container" id="bsw-${tabId}-${filterKey}" onscroll="document.getElementById('tsw-${tabId}-${filterKey}').scrollLeft = this.scrollLeft;">${html}</div></div>`;
    setTimeout(()=> { let tableEl = document.querySelector(`#bsw-${tabId}-${filterKey} table`); if(tableEl) document.getElementById(`tsd-${tabId}-${filterKey}`).style.width = tableEl.offsetWidth + 'px'; }, 100);
    return wrapperHtml;
}

function renderPreview(tabId) {
    let container = document.getElementById(`preview-${tabId}`); let summary = document.getElementById(`summary-${tabId}`);
    let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId]; let tableInfo = tablesConfig.find(t => t.id === tabId);
    if (store.rawFiles.length === 0) { container.innerHTML = '<p style="padding: 10px; color: #666;">尚未匯入資料...</p>'; summary.innerHTML = ''; return; }
    
    // 💡 原始資料預覽與標題列指定區塊
    let rawHtml = '';
    if(store.rawFiles.length > 0) {
        let rows = store.rawFiles[0].rows.slice(0, 50); // 顯示前 50 列
        rawHtml += `<div style="background:#f8f9fa; border:1px solid #ddd; padding:15px; border-radius:5px; margin-bottom:20px;">
            <h4 style="color:#8e44ad; margin-top:0; margin-bottom:10px;">📂 原始檔案預覽 <span style="font-size:13px; font-weight:normal; color:#555;">(顯示前50列，點擊任一列即可設為「解析起始列 / 標題列」)</span></h4>
            <div class="table-wrapper" style="max-height: 250px; overflow-y: auto;">
                <table class="preview-table"><tbody>`;
        rows.forEach((r, idx) => {
            let isHeader = (idx === conf.headerRowIndex);
            let bg = isHeader ? 'background:#ffeeba; font-weight:bold; border:2px solid #ffc107;' : 'cursor:pointer;';
            let hoverClass = isHeader ? '' : 'class="interactive-row"';
            rawHtml += `<tr style="${bg}" ${hoverClass} onclick="selectHeaderRow('${tabId}', ${idx})" title="點擊設為標題/起始列">
                <td style="color:#888; background:#f0f0f0; width:50px; text-align:center;">${idx + 1} ${isHeader?'<br><span style="color:#d35400;font-size:10px;">(標題/起始)</span>':''}</td>`;
            r.forEach(c => rawHtml += `<td>${c!==undefined?c:''}</td>`);
            rawHtml += `</tr>`;
        });
        rawHtml += `</tbody></table></div></div>`;
    }

    summary.innerHTML = `✅ 標題/起始列已設為: 第 ${(conf.headerRowIndex||0) + 1} 列 | 合併產生 ${store.mergedData.length} 筆解析資料。`;
    
    if (tabId === '1-1') {
        container.innerHTML = rawHtml + `<div style="display:flex; gap:15px;"><div style="flex:1.5; overflow-x:auto; border-right:1px solid #ddd; padding-right:15px;"><h4 style="color:#005a9e; margin-top:0;">1. 解析後群組資料</h4>${renderTableHTML(store.mergedData, store.headers, tabId)}</div><div style="flex:1; overflow-x:auto;"><h4 style="color:#1e8449; margin-top:0;">2. 實體 MAP 網格 📍</h4>${renderTableHTML(store.mapGridData, ['No.', 'Map_X', 'Map_Y', 'Row_Ref'], tabId, false, false, 'filters_sec')}</div></div>`; return;
    }

    if (tabId === '1-3' && conf.pivotConfig.enabled) { 
        let uniqueKeyCount = store.pivotedHeaders.length - 4; 
        container.innerHTML = rawHtml + `<div style="display:flex; gap:15px;"><div style="flex:1; overflow-x:auto; border-right:1px solid #ddd; padding-right:15px;"><h4 style="color:#005a9e; margin-top:0;">1. 解析後直表 (找到 ${uniqueKeyCount} 個項目)</h4>${renderTableHTML(store.mergedData, store.headers, tabId)}</div><div style="flex:1.5; overflow-x:auto;"><h4 style="color:#117a65; margin-top:0;">2. 轉置橫表 (產生 ${uniqueKeyCount} 欄) 📍</h4>${renderTableHTML(store.pivotedData, store.pivotedHeaders, tabId, false, false, 'filters_sec')}</div></div>`;
        return; 
    }
    container.innerHTML = rawHtml + `<h4 style="color:#005a9e; margin-top:0;">📝 解析後資料</h4>` + renderTableHTML(store.mergedData, store.headers, tabId);
}

window.getTableDimensions = function(tabId) {
    let store = dataStore[tabId]; let conf = configStore.pns[configStore.currentPN].settings[tabId] || {}; let p3Conf = configStore.pns[configStore.currentPN].settings['phase3'] || {transforms:{}};
    if (!store) return null;
    if (tabId === '1-1') { if (!store.mapGridData || store.mapGridData.length === 0) return null; let uX = new Set(store.mapGridData.map(g=>g[1])); let uY = new Set(store.mapGridData.map(g=>g[2])); return { cols: uX.size, rows: uY.size }; }
    let bData = store.mergedData; let bHeaders = store.headers;
    if (!bData || bData.length === 0) return null;
    if (tabId === '1-3' && conf.pivotConfig && conf.pivotConfig.enabled) { bData = store.pivotedData; bHeaders = store.pivotedHeaders; }
    if (tabId === '1-5' && store.joinedHeaders && store.joinedData) { bData = store.joinedData; bHeaders = store.joinedHeaders; }
    if (tabId === '1-2' && conf.mapConfig && conf.mapConfig.isCombined && store.combinedData1_2) { bData = store.combinedData1_2; bHeaders = store.combinedHeaders1_2; }
    if (!conf.mapConfig || !conf.mapConfig.xCol || !conf.mapConfig.yCol) return null;
    let xIdx = bHeaders.indexOf(conf.mapConfig.xCol); let yIdx = bHeaders.indexOf(conf.mapConfig.yCol); if (xIdx === -1 || yIdx === -1) return null;
    let uX = new Set(bData.map(r=>parseFloat(r[xIdx])).filter(v=>!isNaN(v))); let uY = new Set(bData.map(r=>parseFloat(r[yIdx])).filter(v=>!isNaN(v)));
    let tCols = uX.size; let tRows = uY.size; let tr = p3Conf.transforms[tabId] || { rot: '0' };
    if (tr.rot === '90' || tr.rot === '270') { tCols = uY.size; tRows = uX.size; }
    return { cols: tCols, rows: tRows };
};

window.preCheckMasterJoin = function() {
    let dims11 = getTableDimensions('1-1'); if(!dims11) { alert("缺乏 1-1 實體工位檔網格！"); return; }
    let mismatches = [];
    ['1-2', '1-4', '1-5'].forEach(tabId => { let dims = getTableDimensions(tabId); if (dims && (dims.cols !== dims11.cols || dims.rows !== dims11.rows)) { mismatches.push({ tabId, cols: dims.cols, rows: dims.rows }); } });
    if (mismatches.length > 0) showAlignModal(dims11.cols, dims11.rows, mismatches); else executeMasterJoin();
}

window.showAlignModal = function(cols11, rows11, mismatches) {
    let modal = document.getElementById('align-modal');
    if (!modal) { modal = document.createElement('div'); modal.id = 'align-modal'; modal.className = 'loading-mask'; modal.style.cssText = 'z-index: 10001; background: rgba(0,0,0,0.6); display:none; justify-content:center; align-items:center; position:fixed; top:0; left:0; width:100%; height:100%;'; document.body.appendChild(modal); }
    let p3Conf = configStore.pns[configStore.currentPN].settings['phase3'];
    let html = `<div style="background: #fff; width: 90%; max-width: 800px; padding: 20px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); display: flex; flex-direction: column; max-height: 90vh;"><h3 style="margin-top:0; color:#d35400;">⚠️ 網格維度不一致警告</h3><p style="font-size:14px; color:#555;">下列附表的行列數 (已考慮旋轉) 與基準表 (1-1) 不符，請確認是否需要手動調整偏移量，未對齊而超出 1-1 範圍的部分將被自動裁切。</p><div style="display:flex; gap:15px; background:#f8f9fa; padding:10px; border-radius:5px; margin-bottom:10px; align-items:center;"><div style="font-weight:bold; color:#1e8449; font-size:16px;">[基準 1-1] 寬: ${cols11} / 高: ${rows11}</div></div><div style="flex:1; overflow-y:auto; padding-right:10px;">`;
    mismatches.forEach(m => {
        let tr = p3Conf.transforms[m.tabId] || { offsetX: 0, offsetY: 0 }; let tName = { '1-2': 'Prober', '1-4': 'AOI_Dim', '1-5': 'AOI_Defect' }[m.tabId] || m.tabId;
        html += `<div style="border:1px solid #ccc; border-radius:5px; padding:15px; margin-bottom:15px; background:#fff;"><h4 style="margin:0 0 10px 0; color:#005a9e;">${m.tabId} (${tName}) - 寬: <span style="${m.cols!==cols11?'color:red;':''}">${m.cols}</span> / 高: <span style="${m.rows!==rows11?'color:red;':''}">${m.rows}</span></h4><div style="display:flex; gap:20px; align-items:center;"><div style="background:#fef9e7; padding:10px; border-radius:4px; border:1px solid #f1c40f;"><label style="font-weight:bold;">X 軸偏移 (Offset X): <input type="number" id="offX-${m.tabId}" value="${tr.offsetX||0}" style="width:60px; padding:3px;" onchange="updateP3Transform('${m.tabId}', 'offsetX', parseInt(this.value)||0); drawAlignPreview('${m.tabId}', ${cols11}, ${rows11}, ${m.cols}, ${m.rows})"></label><br><label style="margin-top:10px; display:inline-block; font-weight:bold;">Y 軸偏移 (Offset Y): <input type="number" id="offY-${m.tabId}" value="${tr.offsetY||0}" style="width:60px; padding:3px;" onchange="updateP3Transform('${m.tabId}', 'offsetY', parseInt(this.value)||0); drawAlignPreview('${m.tabId}', ${cols11}, ${rows11}, ${m.cols}, ${m.rows})"></label><div style="font-size:12px; color:#d35400; margin-top:10px; line-height:1.4;">提示：<br>正值向右/下移動<br>負值向左/上移動。</div></div><div style="flex:1; text-align:center;"><canvas id="cvs-align-${m.tabId}" width="400" height="200" style="background:#eef5fa; border:1px solid #b8d4ea; border-radius:4px; max-width:100%;"></canvas></div></div></div>`;
    });
    html += `</div><div style="margin-top:20px; text-align:right; border-top:1px solid #eee; padding-top:15px;"><button class="btn btn-danger" onclick="document.getElementById('align-modal').style.display='none'" style="margin-right:10px;">取消修改</button><button class="btn btn-primary" onclick="document.getElementById('align-modal').style.display='none'; executeMasterJoin();">✅ 確認偏移並執行融合</button></div></div>`;
    modal.innerHTML = html; modal.style.display = 'flex'; setTimeout(() => { mismatches.forEach(m => drawAlignPreview(m.tabId, cols11, rows11, m.cols, m.rows)); }, 50);
}

window.drawAlignPreview = function(tabId, c11, r11, cT, rT) {
    let cvs = document.getElementById(`cvs-align-${tabId}`); if(!cvs) return;
    let ctx = cvs.getContext('2d'); ctx.clearRect(0,0,cvs.width,cvs.height);
    let tr = configStore.pns[configStore.currentPN].settings['phase3'].transforms[tabId] || {}; let offX = tr.offsetX || 0; let offY = tr.offsetY || 0;
    let maxC = Math.max(c11, cT + Math.abs(offX)); let maxR = Math.max(r11, rT + Math.abs(offY)); let scale = Math.min(350 / maxC, 160 / maxR);
    let cx = cvs.width / 2; let cy = cvs.height / 2; let w11 = c11 * scale; let h11 = r11 * scale; let wT = cT * scale; let hT = rT * scale;
    let startX11 = cx - Math.max(w11, wT)/2; let startY11 = cy - Math.max(h11, hT)/2;
    ctx.strokeStyle = '#1e8449'; ctx.lineWidth = 2; ctx.setLineDash([4,4]); ctx.strokeRect(startX11, startY11, w11, h11); ctx.fillStyle = 'rgba(30, 132, 73, 0.1)'; ctx.fillRect(startX11, startY11, w11, h11); ctx.fillStyle = '#1e8449'; ctx.font = "12px Arial"; ctx.fillText("1-1 (基準)", startX11, startY11 - 5);
    let startXT = startX11 + (offX * scale); let startYT = startY11 + (offY * scale);
    ctx.strokeStyle = '#d35400'; ctx.lineWidth = 2; ctx.setLineDash([]); ctx.strokeRect(startXT, startYT, wT, hT); ctx.fillStyle = 'rgba(211, 84, 0, 0.3)'; ctx.fillRect(startXT, startYT, wT, hT); ctx.fillStyle = '#d35400'; ctx.fillText(tabId + " (目標)", startXT, startYT + hT + 14);
}