const GAS_URL = "https://teams.microsoft.com/l/message/48:notes/1766115088638?context=%7B%22contextType%22%3A%22chat%22%2C%22oid%22%3A%228%3Aorgid%3A988e9cdd-f5d4-471e-967e-676c0eec685a%22%7D";

let records = [];

window.onload = () => fetchData();

// データの取得
async function fetchData() {
    toggleLoading(true);
    try {
        const response = await fetch(GAS_URL);
        const data = await response.json();
        records = data.records;
        updateSelectors(data.customers, data.staffs);
        renderTable();
    } catch (e) { console.error(e); }
    toggleLoading(false);
}

// マスタ選択肢の更新
function updateSelectors(customers, staffs) {
    document.getElementById('customer-list').innerHTML = customers.map(c => `<option value="${c}">`).join('');
    const staffSelect = document.getElementById('staffName');
    staffSelect.innerHTML = '<option value="">選択...</option>' + staffs.map(s => `<option value="${s}">${s}</option>`).join('');
}

// データの保存
async function saveData() {
    const isCompleted = document.getElementById('isCompleted').checked;
    const data = {
        customerName: document.getElementById('customerName').value,
        staffName: document.getElementById('staffNameFree').value || document.getElementById('staffName').value,
        issueDate: document.getElementById('issueDate').value,
        inputDate: document.getElementById('inputDate').value,
        mailDate: document.getElementById('mailDate').value,
        emailDate: document.getElementById('emailDate').value,
        memo: document.getElementById('memo').value,
        isCompleted: isCompleted // booleanで送信
    };

    if (!data.customerName) return alert("お客様名を入力してください");

    toggleLoading(true);
    await fetch(GAS_URL, {
        method: "POST",
        body: JSON.stringify({ action: 'save', data: data })
    });
    
    await fetchData();
    clearForm();
}

// マスタの追加
async function addMaster(action, inputId) {
    const val = document.getElementById(inputId).value;
    if (!val) return;
    toggleLoading(true);
    await fetch(GAS_URL, {
        method: "POST",
        body: JSON.stringify({ action: action, data: { name: val } })
    });
    document.getElementById(inputId).value = "";
    await fetchData();
}

// テーブル描画
function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    // 未完了を上にするソート
    const sorted = [...records].sort((a, b) => {
        const aDone = a.完了フラグ === "TRUE" ? 1 : 0;
        const bDone = b.完了フラグ === "TRUE" ? 1 : 0;
        return aDone - bDone;
    });

    sorted.forEach(r => {
        const isDone = (r.完了フラグ === "TRUE");
        let statusClass = isDone ? 'status-green' : ((r.郵送日 || r.メール送信日) ? 'status-yellow' : 'status-red');
        
        const row = `
            <tr class="${statusClass}">
                <td title="${r.お客様名}">${r.お客様名}</td>
                <td>${r.担当者名}</td>
                <td>${r.明細発行日 || '-'}</td>
                <td>${r.管理表入力日 || '-'}</td>
                <td>${r.郵送日 || '-'}</td>
                <td>${r.メール送信日 || '-'}</td>
                <td title="${r.メモ}">${r.メモ}</td>
                <td style="text-align:center">${isDone ? '✓' : ''}</td>
                <td class="no-print">
                    <button class="btn-edit" onclick="editLocal('${r.お客様名}')">編集</button>
                    <button class="btn-delete" onclick="deleteData('${r.お客様名}')">削除</button>
                </td>
            </tr>
        `;
        tbody.insertAdjacentHTML('beforeend', row);
    });
}

function editLocal(name) {
    const r = records.find(item => item.お客様名 === name);
    document.getElementById('customerName').value = r.お客様名;
    document.getElementById('staffNameFree').value = r.担当者名;
    document.getElementById('issueDate').value = r.明細発行日;
    document.getElementById('inputDate').value = r.管理表入力日;
    document.getElementById('mailDate').value = r.郵送日;
    document.getElementById('emailDate').value = r.メール送信日;
    document.getElementById('memo').value = r.メモ;
    document.getElementById('isCompleted').checked = (r.完了フラグ === "TRUE");
    window.scrollTo(0,0);
}

async function deleteData(name) {
    if (!confirm(`${name} のデータを削除しますか？`)) return;
    toggleLoading(true);
    await fetch(GAS_URL, { method: "POST", body: JSON.stringify({ action: 'delete', data: { customerName: name } }) });
    await fetchData();
}

async function processAllComplete() {
    const allDone = records.every(r => r.完了フラグ === "TRUE");
    if (!allDone) return alert("まだ未完了のデータがあります。");
    
    window.print();
    setTimeout(async () => {
        if (confirm("PDF保存は完了しましたか？\nOKを押すと全データを削除し今月分を終了します。")) {
            toggleLoading(true);
            await fetch(GAS_URL, { method: "POST", body: JSON.stringify({ action: 'clearAll' }) });
            await fetchData();
        }
    }, 1000);
}

function toggleLoading(show) { document.getElementById('loading').style.display = show ? 'flex' : 'none'; }
function clearForm() {
    document.querySelectorAll('.input-form input, .input-form select, .input-form textarea').forEach(el => el.value = "");
    document.getElementById('isCompleted').checked = false;
}
