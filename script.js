let studentsData = [];
let currentGroups = [];

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    document.getElementById('fileName').textContent = file.name;
    document.getElementById('fileSize').textContent = formatFileSize(file.size);
    document.getElementById('fileInfo').style.display = 'block';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});

            parseStudentData(jsonData);
        } catch (error) {
            showError('文件解析失败，请确保文件格式正确！');
        }
    };
    reader.readAsArrayBuffer(file);
}

function parseStudentData(data) {
    if (data.length < 2) {
        showError('文件数据不足，请确保至少有一行数据！');
        return;
    }

    studentsData = [];
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row && row.length >= 2 && row[1]) {
            studentsData.push({
                序号: row[0] || i,
                姓名: row[1],
                性别: row[2] || '未知',
                年龄: row[3] || '未知'
            });
        }
    }

    if (studentsData.length === 0) {
        showError('未找到有效的学生数据！');
        return;
    }

    displayDataPreview();
    document.getElementById('generateBtn').disabled = false;
}

function displayDataPreview() {
    const preview = document.getElementById('dataPreview');
    const stats = document.getElementById('dataStats');
    const membersList = document.getElementById('membersList');

    const maleCount = studentsData.filter(s => s.性别 === '男').length;
    const femaleCount = studentsData.filter(s => s.性别 === '女').length;

    stats.innerHTML = `
        <div class="stat-item">
            <div class="stat-number">${studentsData.length}</div>
            <div class="stat-label">总人数</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">${maleCount}</div>
            <div class="stat-label">男生</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">${femaleCount}</div>
            <div class="stat-label">女生</div>
        </div>
    `;

    membersList.innerHTML = studentsData.map(student => 
        `<span class="member">${student.姓名} (${student.性别}, ${student.年龄}岁)</span>`
    ).join('');

    preview.style.display = 'block';
}

function toggleGroupMode() {
    const mode = document.getElementById('groupMode').value;
    const groupsDiv = document.getElementById('groupsCountDiv');
    const membersDiv = document.getElementById('membersPerGroupDiv');

    if (mode === 'byGroups') {
        groupsDiv.style.display = 'block';
        membersDiv.style.display = 'none';
    } else {
        groupsDiv.style.display = 'none';
        membersDiv.style.display = 'block';
    }
}

function generateGroups() {
    if (studentsData.length === 0) {
        showError('请先上传Excel文件！');
        return;
    }

    const mode = document.getElementById('groupMode').value;
    let groupCount;

    if (mode === 'byGroups') {
        groupCount = parseInt(document.getElementById('groupsCount').value);
        if (!groupCount || groupCount < 1) {
            showError('请输入有效的组数！');
            return;
        }
    } else {
        const membersPerGroup = parseInt(document.getElementById('membersPerGroup').value);
        if (!membersPerGroup || membersPerGroup < 1) {
            showError('请输入有效的每组人数！');
            return;
        }
        groupCount = Math.ceil(studentsData.length / membersPerGroup);
    }

    const shuffledStudents = [...studentsData].sort(() => Math.random() - 0.5);
    currentGroups = [];

    for (let i = 0; i < groupCount; i++) {
        currentGroups.push([]);
    }

    shuffledStudents.forEach((student, index) => {
        currentGroups[index % groupCount].push(student);
    });

    displayResults();
}

function displayResults() {
    const results = document.getElementById('results');
    const stats = document.getElementById('resultStats');
    const container = document.getElementById('groupsContainer');

    const avgSize = (studentsData.length / currentGroups.length).toFixed(1);

    stats.innerHTML = `
        <div class="stat-item">
            <div class="stat-number">${currentGroups.length}</div>
            <div class="stat-label">分组数</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">${avgSize}</div>
            <div class="stat-label">平均每组人数</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">${Math.max(...currentGroups.map(g => g.length))}</div>
            <div class="stat-label">最大组人数</div>
        </div>
        <div class="stat-item">
            <div class="stat-number">${Math.min(...currentGroups.map(g => g.length))}</div>
            <div class="stat-label">最小组人数</div>
        </div>
    `;

    container.innerHTML = currentGroups.map((group, index) => `
        <div class="group">
            <div class="group-header">
                第${index + 1}组 (${group.length}人)
            </div>
            <div class="group-content">
                ${group.map(student => 
                    `<span class="member">${student.姓名} (${student.性别}, ${student.年龄}岁)</span>`
                ).join('')}
            </div>
        </div>
    `).join('');

    results.style.display = 'block';
    results.scrollIntoView({ behavior: 'smooth' });
}

function exportResults() {
    if (currentGroups.length === 0) {
        showError('请先进行分组！');
        return;
    }

    const exportData = [];
    exportData.push(['组别', '姓名', '性别', '年龄']);

    currentGroups.forEach((group, groupIndex) => {
        group.forEach(student => {
            exportData.push([
                `第${groupIndex + 1}组`,
                student.姓名,
                student.性别,
                student.年龄
            ]);
        });
    });

    const ws = XLSX.utils.aoa_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "分组结果");

    const fileName = `随机分组结果_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function showError(message) {
    const existingAlert = document.querySelector('.alert-error');
    if (existingAlert) {
        existingAlert.remove();
    }

    const alert = document.createElement('div');
    alert.className = 'alert alert-error';
    alert.textContent = message;
    
    const content = document.querySelector('.content');
    content.insertBefore(alert, content.firstChild);

    setTimeout(() => {
        alert.remove();
    }, 5000);
}

document.addEventListener('DOMContentLoaded', function() {
    document.addEventListener('dragover', function(e) {
        e.preventDefault();
    });

    document.addEventListener('drop', function(e) {
        e.preventDefault();
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            document.getElementById('fileInput').files = files;
            handleFileUpload({ target: { files: files } });
        }
    });
});