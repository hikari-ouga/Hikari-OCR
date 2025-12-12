// ========================================
// グローバル変数
// ========================================
let uploadedFiles = [];
let currentMode = 'single';

// ========================================
// DOM要素取得
// ========================================
const elements = {
    corpName: document.getElementById('corpName'),
    modeRadios: document.querySelectorAll('input[name="mode"]'),
    startMonthGroup: document.getElementById('startMonthGroup'),
    startMonth: document.getElementById('startMonth'),
    monthOrderGroup: document.getElementById('monthOrderGroup'),
    monthOrderRadios: document.querySelectorAll('input[name="monthOrder"]'),
    uploadArea: document.getElementById('uploadArea'),
    pdfFiles: document.getElementById('pdfFiles'),
    fileList: document.getElementById('fileList'),
    monthSelectionArea: document.getElementById('monthSelectionArea'),
    monthSelectionList: document.getElementById('monthSelectionList'),
    executeBtn: document.getElementById('executeBtn'),
    resetBtn: document.getElementById('resetBtn'),
    progressArea: document.getElementById('progressArea'),
    progressText: document.getElementById('progressText'),
    resultArea: document.getElementById('resultArea'),
    downloadArea: document.getElementById('downloadArea'),
    downloadBtn: document.getElementById('downloadBtn')
};

// ========================================
// 初期化
// ========================================
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
});

function initializeEventListeners() {
    // モード切り替え
    elements.modeRadios.forEach(radio => {
        radio.addEventListener('change', handleModeChange);
    });

    // ファイルアップロード
    elements.uploadArea.addEventListener('click', () => elements.pdfFiles.click());
    elements.pdfFiles.addEventListener('change', handleFileSelect);
    
    // ドラッグ&ドロップ
    elements.uploadArea.addEventListener('dragover', handleDragOver);
    elements.uploadArea.addEventListener('dragleave', handleDragLeave);
    elements.uploadArea.addEventListener('drop', handleDrop);

    // 実行ボタン
    elements.executeBtn.addEventListener('click', handleExecute);

    // リセットボタン
    elements.resetBtn.addEventListener('click', handleReset);

    // ダウンロードボタン
    elements.downloadBtn.addEventListener('click', handleDownload);
}

// ========================================
// モード切り替え処理
// ========================================
function handleModeChange(e) {
    currentMode = e.target.value;
    
    if (currentMode === 'multi') {
        elements.startMonthGroup.style.display = 'block';
        elements.monthOrderGroup.style.display = 'block';
    } else {
        elements.startMonthGroup.style.display = 'none';
        elements.monthOrderGroup.style.display = 'none';
    }
    // モード変更時にリストを再描画（プルダウンの表示/非表示を切り替えるため）
    renderFileList();
}

// ========================================
// ファイル選択処理
// ========================================
function handleFileSelect(e) {
    const files = Array.from(e.target.files);
    addFiles(files);
}

function handleDragOver(e) {
    e.preventDefault();
    elements.uploadArea.classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    elements.uploadArea.classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    elements.uploadArea.classList.remove('drag-over');
    
    const files = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf');
    addFiles(files);
}

function addFiles(files) {
    files.forEach(file => {
        const detectedMonth = detectMonthFromFilename(file.name);
        uploadedFiles.push({
            file: file,
            detectedMonth: detectedMonth,
            selectedMonth: detectedMonth || 1
        });
    });
    
    renderFileList();
    updateExecuteButton();
}

// ========================================
// ファイル名から月を自動検出
// ========================================
function detectMonthFromFilename(filename) {
    // パターン1: "1月" "01月" "１月"などの形式
    let match = filename.match(/([0-9０-９]{1,2})\s*月/);
    if (match) {
        let monthStr = match[1];
        // 全角数字を半角に変換
        monthStr = monthStr.replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
        const month = parseInt(monthStr);
        if (month >= 1 && month <= 12) return month;
    }
    
    // パターン2: "_01_" "2025-01" "-01." などの形式
    match = filename.match(/[_\-]([0-9]{2})[_\-\.]/);
    if (match) {
        const month = parseInt(match[1]);
        if (month >= 1 && month <= 12) return month;
    }
    
    // パターン3: 英語の月名
    const monthNames = {
        'jan': 1, 'january': 1,
        'feb': 2, 'february': 2,
        'mar': 3, 'march': 3,
        'apr': 4, 'april': 4,
        'may': 5,
        'jun': 6, 'june': 6,
        'jul': 7, 'july': 7,
        'aug': 8, 'august': 8,
        'sep': 9, 'september': 9,
        'oct': 10, 'october': 10,
        'nov': 11, 'november': 11,
        'dec': 12, 'december': 12
    };
    
    const lowerFilename = filename.toLowerCase();
    for (const [name, month] of Object.entries(monthNames)) {
        if (lowerFilename.includes(name)) {
            return month;
        }
    }
    
    return null;
}

// ========================================
// ファイルリスト表示
// ========================================
function renderFileList() {
    if (uploadedFiles.length === 0) {
        elements.fileList.innerHTML = '';
        return;
    }
    
    elements.fileList.innerHTML = uploadedFiles.map((item, index) => {
        // 単月モードの場合、月選択プルダウンを表示
        let monthSelector = '';
        if (currentMode === 'single') {
            const options = Array.from({length: 12}, (_, i) => i + 1).map(m => 
                `<option value="${m}" ${item.selectedMonth === m ? 'selected' : ''}>${m}月</option>`
            ).join('');
            
            monthSelector = `
                <div class="file-month-select">
                    <select onchange="updateFileMonth(${index}, this.value)" class="compact-select">
                        ${options}
                    </select>
                    ${item.detectedMonth ? `<span class="auto-detect-badge">自動: ${item.detectedMonth}月</span>` : ''}
                </div>
            `;
        }

        return `
        <div class="file-item">
            <div class="file-info-group">
                <div class="file-main-info">
                    <div class="file-icon">PDF</div>
                    <div class="file-name" title="${item.file.name}">${item.file.name}</div>
                </div>
                ${monthSelector}
            </div>
            <button class="file-remove" onclick="removeFile(${index})" title="削除">×</button>
        </div>
        `;
    }).join('');
}

function removeFile(index) {
    uploadedFiles.splice(index, 1);
    renderFileList();
    updateExecuteButton();
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// ========================================
// 月選択UI表示（単月モード）
// ========================================
// 廃止: renderFileListに統合されました
function renderMonthSelections() {
    // 何もしない
}

function updateFileMonth(index, month) {
    uploadedFiles[index].selectedMonth = parseInt(month);
}

// ========================================
// 実行ボタン制御
// ========================================
function updateExecuteButton() {
    elements.executeBtn.disabled = uploadedFiles.length === 0;
}

// ========================================
// 実行処理
// ========================================
async function handleExecute() {
    const corpName = elements.corpName.value.trim();
    
    if (!corpName) {
        alert('法人名を入力してください。');
        return;
    }
    
    if (uploadedFiles.length === 0) {
        alert('PDFファイルをアップロードしてください。');
        return;
    }
    
    // UI更新
    elements.executeBtn.disabled = true;
    elements.progressArea.style.display = 'block';
    elements.progressText.textContent = 'AIが解析中...しばらくお待ちください';
    
    // テーブル初期化
    elements.resultArea.innerHTML = `
        <table class="result-table">
            <thead>
                <tr>
                    <th>対象月</th>
                    <th>使用電力量</th>
                    <th>ステータス</th>
                </tr>
            </thead>
            <tbody id="resultTableBody">
            </tbody>
        </table>
    `;
    elements.downloadArea.style.display = 'none';
    
    const resultTableBody = document.getElementById('resultTableBody');

    try {
        // FormDataの準備
        const formData = new FormData();
        formData.append('corp_name', corpName);
        formData.append('mode', currentMode);
        
        if (currentMode === 'multi') {
            formData.append('start_month', elements.startMonth.value);
            
            // 月の並び順を取得
            const selectedOrder = Array.from(elements.monthOrderRadios).find(r => r.checked)?.value || 'ascending';
            formData.append('month_order', selectedOrder);
        }
        
        // 月マッピング情報
        const monthMappings = uploadedFiles.map(item => ({
            filename: item.file.name,
            selectedMonth: item.selectedMonth
        }));
        formData.append('month_mappings', JSON.stringify(monthMappings));
        
        // ファイル追加
        uploadedFiles.forEach(item => {
            formData.append('files', item.file);
        });
        
        // 一括処理API呼び出し（以前の安定した方式）
        const response = await fetch('/api/process', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.detail || '処理に失敗しました');
        }
        
        const result = await response.json();
        
        // 結果をテーブルに表示
        if (result.results) {
            result.results.forEach(item => {
                // selectedMonthを探す
                const uploadedFile = uploadedFiles.find(f => f.file.name === item.filename);
                const selectedMonth = uploadedFile ? uploadedFile.selectedMonth : null;
                
                if (item.status === '完了') {
                    addResultRow(resultTableBody, item, selectedMonth);
                } else {
                    addErrorRow(resultTableBody, item.filename);
                }
            });
        }

        // 完了表示
        elements.progressArea.style.display = 'none';
        elements.downloadArea.style.display = 'block';
        elements.executeBtn.style.display = 'none';
        elements.resetBtn.style.display = 'inline-flex';

    } catch (error) {
        console.error('Error:', error);
        alert('処理中にエラーが発生しました: ' + error.message);
        elements.progressArea.style.display = 'none';
    } finally {
        elements.executeBtn.disabled = false;
    }
}

function addResultRow(tbody, result, selectedMonth) {
    // fields: {"1月値": 12345, "2月値": 23456, ...} のような形式
    if (result.fields && Object.keys(result.fields).length > 0) {
        // 複数月モードの場合、全ての月を表示
        const monthKeys = Object.keys(result.fields).sort((a, b) => {
            const monthA = parseInt(a.replace('月値', ''));
            const monthB = parseInt(b.replace('月値', ''));
            return monthA - monthB;
        });
        
        monthKeys.forEach(key => {
            const row = document.createElement('tr');
            const monthDisplay = key.replace('値', ''); // "1月"
            const kwhDisplay = result.fields[key] + ' kWh';
            
            row.innerHTML = `
                <td class="col-month">${monthDisplay}</td>
                <td class="col-kwh">${kwhDisplay}</td>
                <td class="col-status"><span class="status-badge success">完了</span></td>
            `;
            tbody.appendChild(row);
        });
    } else if (selectedMonth) {
        // データがない場合
        const row = document.createElement('tr');
        row.innerHTML = `
            <td class="col-month">${selectedMonth}月</td>
            <td class="col-kwh">抽出失敗</td>
            <td class="col-status"><span class="status-badge success">完了</span></td>
        `;
        tbody.appendChild(row);
    }
}

function addErrorRow(tbody, filename) {
    const row = document.createElement('tr');
    row.innerHTML = `
        <td class="col-month">-</td>
        <td class="col-kwh text-error">エラー</td>
        <td class="col-status"><span class="status-badge error" title="${filename}">失敗</span></td>
    `;
    tbody.appendChild(row);
}

// ========================================
// 結果表示 (廃止: handleExecute内で直接描画)
// ========================================
function displayResults(result) {
    // ...
}

// ========================================
// ダウンロード処理
// ========================================
async function handleDownload() {
    try {
        const corpName = elements.corpName.value.trim() || 'output';
        const filename = `${corpName}.xlsx`;
        
        const response = await fetch('/api/download');
        
        if (!response.ok) {
            throw new Error('ダウンロードに失敗しました');
        }
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
    } catch (error) {
        console.error('Download error:', error);
        alert('ダウンロードに失敗しました: ' + error.message);
    }
}

// ========================================
// リセット処理
// ========================================
function handleReset() {
    // ファイルリストをクリア
    uploadedFiles = [];
    renderFileList();
    
    // 結果エリアを初期状態に戻す
    elements.resultArea.innerHTML = `
        <div class="empty-state">
            <div class="empty-icon">📊</div>
            <p>解析結果がここに表示されます</p>
        </div>
    `;
    
    // ボタン表示を元に戻す
    elements.executeBtn.style.display = 'inline-flex';
    elements.executeBtn.disabled = true;
    elements.resetBtn.style.display = 'none';
    elements.downloadArea.style.display = 'none';
    elements.progressArea.style.display = 'none';
    
    // ファイル入力をリセット
    elements.pdfFiles.value = '';
}
