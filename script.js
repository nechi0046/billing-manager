const GAS_URL = "https://teams.microsoft.com/l/message/48:notes/1766113081175?context=%7B%22contextType%22%3A%22chat%22%2C%22oid%22%3A%228%3Aorgid%3A988e9cdd-f5d4-471e-967e-676c0eec685a%22%7D";

let records = [];

// ページ読み込み時に実行
window.onload = () => {
    fetchData();
};

async function fetchData() {
    showLoading(true);
    try {
        const response = await fetch(GAS_URL);
        const data = await response.json();
        records = data.records;
        updateMasterUI(data.customers, data.staffs);
        renderTable();
    } catch (e) {
        alert("データの取得に失敗しました。");
    }
    showLoading(false);
}

function updateMasterUI(customers, staffs) {
    const custDatalist = document.getElementById('customer-list');
    custDatalist.innerHTML = customers.map(c => `<option value="${c}">`).join('');

    const staffSelect = document.getElementById('staffName');
    staffSelect.innerHTML = '<option value="">選択してください</option>' + 
                            staffs.map(s => `<option value="${s}">${s}</option>`).join('');
}

async function saveData() {
    const staffSelectValue = document.getElementById('staffName').value;
    const staffFreeValue = document.getElementById('staffNameFree').value;

    const data = {
        customerName: document.getElementById('customerName').value,
        staffName: staffFreeValue || staffSelectValue,
        issueDate: document.getElementById('issueDate').value,
        inputDate: document.getElementById('inputDate').value,
        mailDate: document.getElementById('mailDate').value,
        emailDate: document.getElementById('emailDate').value,
        memo: document.getElementById('memo').value,
        isCompleted: document.getElementById('isCompleted').checked
    };

    if (!data.customerName) return alert("お客様名を入力してください");

    showLoading(true);
    await fetch(GAS_URL, {
        method: "POST",
        body: JSON.stringify({ action: 'save', data: data })
    });
    
    await fetchData(); // 再読み込み
    clearForm();
}

async function deleteData(name) {
    if (!confirm(`${name} のデータを削除しますか？`)) return;
    showLoading(true);
    await fetch(GAS_URL, {
        method: "POST",
        body: JSON.stringify({ action: 'delete', data: { customerName: name } })
    });
    await fetchData();
}

function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    // 未完了を上にするソート
    const sorted = [...records].sort((a, b) => (a.完了フラグ === "TRUE" ? 1 : 0) - (b.完了フラグ === "TRUE" ? 1 : 0));

    sorted.forEach(r => {
        const isDone = r.完了フラグ === "TRUE";
        let statusClass = isDone ? 'status-green' : (r.郵送日 || r.メール送信日 ? 'status-yellow' : 'status-red');
        
        const row = `
            <tr class="${statusClass}">
                <td>${r.お客様名}</td>
                <td>${r.担当者名}</td>
                <td>${r.明細発行日}</td>
                <td>${r.管理表入力日}</td>
                <td>${r.郵送日 || '-'}</td>
                <td>${r.メール送信日 || '-'}</td>
                <td>${r.メモ}</td>
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
    document.getElementById('issueDate').value = r.明細発行日;
    document.getElementById('inputDate').value = r.管理表入力日;
    document.getElementById('mailDate').value = r.郵送日;
    document.getElementById('emailDate').value = r.メール送信日;
    document.getElementById('memo').value = r.メモ;
    document.getElementById('isCompleted').checked = (r.完了フラグ === "TRUE");
    window.scrollTo(0,0);
}

async function processAllComplete() {
    if (!records.every(r => r.完了フラグ === "TRUE")) return alert("未完了のデータがあります。");
    window.print();
    setTimeout(async () => {
        if (confirm("PDFを保存しましたか？OKを押すとスプレッドシートのデータをリセットします。")) {
            showLoading(true);
            await fetch(GAS_URL, { method: "POST", body: JSON.stringify({ action: 'clearAll' }) });
            await fetchData();
        }
    }, 1000);
}

function showLoading(show) { document.getElementById('loading').style.display = show ? 'flex' : 'none'; }
function clearForm() {
    document.querySelectorAll('.input-form input, .input-form select, .input-form textarea').forEach(el => el.value = "");
    document.getElementById('isCompleted').checked = false;
}