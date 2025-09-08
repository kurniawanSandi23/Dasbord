let currentData = [];
let headers = [];
let originalTableBody = [];
const SHEET_NAME = 'Team List';

function loadFile() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (!file) {
        showError('Pilih file Excel terlebih dahulu!');
        return;
    }

    document.getElementById('loading').style.display = 'block';
    document.getElementById('error').style.display = 'none';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames.find(name => name === SHEET_NAME) || workbook.SheetNames[1];
            const worksheet = workbook.Sheets[sheetName];
            currentData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (currentData.length < 1) {
                showError('File Excel kosong atau tidak valid di sheet ' + SHEET_NAME + '!');
                return;
            }

            headers = currentData[0] || [];
            currentData = currentData.slice(6); // Adjust for header rows

            displayData(currentData);
            document.getElementById('searchSection').style.display = 'block';
            document.getElementById('groupSection').style.display = 'block';
        } catch (error) {
            showError('Error membaca file: ' + error.message);
        } finally {
            document.getElementById('loading').style.display = 'none';
        }
    };
    reader.readAsArrayBuffer(file);
}

function displayData(data) {
    document.getElementById('error').style.display = 'none';
    document.getElementById('stats').style.display = 'flex';
    document.getElementById('dataTable').style.display = 'table';
    document.getElementById('noSearchResults').style.display = 'none';

    const tableBody = document.getElementById('tableBody');
    let bodyHtml = '';
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row.length < 50) continue;
        const employeeName = row[10] || '';
        const wxid = row[11] || '';
        const laptop = row[20] || '';
        const domainStatus = row[49] || '';
        const nwLead = row[19] || '';
        const lokasi = row[29] || '';
        bodyHtml += `<tr>
            <td>${employeeName}</td>
            <td>${wxid}</td>
            <td>${laptop}</td>
            <td>${domainStatus}</td>
            <td>${nwLead}</td>
            <td>${lokasi}</td>
        </tr>`;
    }
    tableBody.innerHTML = bodyHtml;
    originalTableBody = bodyHtml.split('</tr>');

    let totalI5 = 0;
    let totalI7 = 0;
    let wisma2Rendah = 0;
    let bsdRendah = 0;

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row.length < 50) continue;
        const laptopReq = (row[20] || '').toString().toLowerCase();
        const domainStat = (row[49] || '').toString().toLowerCase();
        const lok = (row[29] || '').toString().toLowerCase();

        if (laptopReq.includes('i5')) totalI5++;
        if (laptopReq.includes('i7')) totalI7++;

        const isRendah = domainStat.includes('on going') || domainStat.includes('ny joint') || domainStat.includes('issue');
        if (isRendah) {
            if (lok.includes('wisma') || lok.includes('wm')) wisma2Rendah++;
            if (lok.includes('bsd')) bsdRendah++;
        }
    }

    document.getElementById('totalI5').textContent = totalI5;
    document.getElementById('totalI7').textContent = totalI7;
    document.getElementById('wisma2Rendah').textContent = wisma2Rendah;
    document.getElementById('bsdRendah').textContent = bsdRendah;

    const tableRows = document.querySelectorAll('#tableBody tr');
    tableRows.forEach((row, index) => {
        setTimeout(() => {
            row.classList.add('animate__animated', 'animate__fadeIn');
        }, index * 50);
    });
}

function performSearch() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const tableBody = document.getElementById('tableBody');
    const noResults = document.getElementById('noSearchResults');
    const searchResultsDiv = document.getElementById('searchResults');

    if (searchTerm === '') {
        let bodyHtml = '';
        for (let i = 0; i < currentData.length; i++) {
            const row = currentData[i];
            if (row.length < 50) continue;
            const employeeName = row[10] || '';
            const wxid = row[11] || '';
            const laptop = row[20] || '';
            const domainStatus = row[49] || '';
            const nwLead = row[19] || '';
            const lokasi = row[29] || '';
            bodyHtml += `<tr><td>${employeeName}</td><td>${wxid}</td><td>${laptop}</td><td>${domainStatus}</td><td>${nwLead}</td><td>${lokasi}</td></tr>`;
        }
        tableBody.innerHTML = bodyHtml;
        noResults.style.display = 'none';
        searchResultsDiv.style.display = 'none';
        const tableRows = document.querySelectorAll('#tableBody tr');
        tableRows.forEach((row, index) => {
            setTimeout(() => {
                row.classList.add('animate__animated', 'animate__fadeIn');
            }, index * 50);
        });
        return;
    }

    let filteredRows = [];
    for (let i = 0; i < currentData.length; i++) {
        const row = currentData[i];
        if (row.length < 50) continue;
        const employeeName = (row[10] || '').toString().toLowerCase();
        const wxid = (row[11] || '').toString().toLowerCase();
        const laptop = (row[20] || '').toString().toLowerCase();
        const domainStatus = (row[49] || '').toString().toLowerCase();
        const rowText = [employeeName, wxid, laptop, domainStatus].join(' ');
        if (rowText.includes(searchTerm)) {
            const fullRow = row;
            const employeeNameFull = fullRow[10] || '';
            const wxidFull = fullRow[11] || '';
            const laptopFull = fullRow[20] || '';
            const domainStatusFull = fullRow[49] || '';
            const nwLeadFull = fullRow[19] || '';
            const lokasiFull = fullRow[29] || '';
            filteredRows.push(`<tr><td>${employeeNameFull}</td><td>${wxidFull}</td><td>${laptopFull}</td><td>${domainStatusFull}</td><td>${nwLeadFull}</td><td>${lokasiFull}</td></tr>`);
        }
    }

    if (filteredRows.length > 0) {
        tableBody.innerHTML = filteredRows.join('');
        noResults.style.display = 'none';
        searchResultsDiv.style.display = 'none';
        const filteredTableRows = document.querySelectorAll('#tableBody tr');
        filteredTableRows.forEach((row, index) => {
            setTimeout(() => {
                row.classList.add('animate__animated', 'animate__fadeIn');
            }, index * 50);
        });
    } else {
        tableBody.innerHTML = '';
        noResults.style.display = 'block';
        searchResultsDiv.innerHTML = `Tidak ada hasil untuk "${searchTerm}" (coba Employee Name, WXID, i5/i7, Wisma 2, BSD, On Going).`;
        searchResultsDiv.style.display = 'block';
    }
}

function groupData() {
    const groupByValue = document.getElementById('groupBy').value;
    const groupedData = document.getElementById('groupedData');
    groupedData.innerHTML = '';

    if (!groupByValue) return;

    let groups = {};
    let groupIndex = parseInt(groupByValue);

    for (let i = 0; i < currentData.length; i++) {
        const row = currentData[i];
        if (row.length < 50) continue;
        let groupValue;
        if (groupByValue === 'merek') {
            const laptop = (row[20] || '').toString().toLowerCase();
            groupValue = laptop.includes('i7') ? 'i7' : (laptop.includes('i5') ? 'i5' : 'Lainnya');
        } else {
            groupValue = row[groupIndex] || 'Tidak Diketahui';
        }
        if (!groups[groupValue]) groups[groupValue] = [];
        groups[groupValue].push(row);
    }

    let accordionHtml = '<div class="accordion" id="groupAccordion">';
    Object.keys(groups).forEach((groupKey, index) => {
        let subTableHtml = '';
        groups[groupKey].forEach(subRow => {
            const employeeName = subRow[10] || '';
            const wxid = subRow[11] || '';
            const laptop = subRow[20] || '';
            const domainStatus = subRow[49] || '';
            const nwLead = subRow[19] || '';
            const lokasi = subRow[29] || '';
            subTableHtml += `<tr><td>${employeeName}</td><td>${wxid}</td><td>${laptop}</td><td>${domainStatus}</td><td>${nwLead}</td><td>${lokasi}</td></tr>`;
        });
        accordionHtml += `
            <div class="accordion-item">
                <h2 class="accordion-header">
                    <button class="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapse${index}">
                        ${groupKey} <span class="badge bg-primary ms-2">${groups[groupKey].length} employee</span>
                    </button>
                </h2>
                <div id="collapse${index}" class="accordion-collapse collapse" data-bs-parent="#groupAccordion">
                    <div class="accordion-body">
                        <div class="table-responsive">
                            <table class="table table-sm table-striped">
                                <thead>
                                    <tr><th>Employee Name</th><th>WXID</th><th>Laptop</th><th>Domain Status</th><th>NW Lead</th><th>Lokasi</th></tr>
                                </thead>
                                <tbody>${subTableHtml}</tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });
    accordionHtml += '</div>';
    groupedData.innerHTML = accordionHtml;
}

function showError(message) {
    const errorDiv = document.getElementById('error');
    errorDiv.textContent = message;
    errorDiv.style.display = 'block';
    document.getElementById('stats').style.display = 'none';
    document.getElementById('dataTable').style.display = 'none';
    document.getElementById('searchSection').style.display = 'none';
    document.getElementById('groupSection').style.display = 'none';
    document.getElementById('loading').style.display = 'none';
}