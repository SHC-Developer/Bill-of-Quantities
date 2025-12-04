const DEFAULT_ROWS = 4;
const DEFAULT_COLS = 5;
let generatedBlocks = [];

const statusMessageEl = document.getElementById('statusMessage');
const previewHintEl = document.getElementById('previewHint');
const previewContainer = document.getElementById('blockPreview');
const generateBtn = document.getElementById('generateBtn');
const downloadBtn = document.getElementById('downloadExcelBtn');

generateBtn.addEventListener('click', handleGenerate);
downloadBtn.addEventListener('click', handleDownload);

function handleGenerate() {
    const blockCount = parseInt(document.getElementById('blockCount').value, 10);
    const rowsPerBlock = DEFAULT_ROWS;
    const colsPerBlock = DEFAULT_COLS;
    const avgMin = parseFloat(document.getElementById('avgMin').value);
    const avgMax = parseFloat(document.getElementById('avgMax').value);
    const valueMin = parseFloat(document.getElementById('valueMin').value);
    const valueMax = parseFloat(document.getElementById('valueMax').value);
    const maxAttempt = parseInt(document.getElementById('maxAttempt').value, 10);

    if (!validateInputs({ blockCount, avgMin, avgMax, valueMin, valueMax, maxAttempt })) {
        return;
    }

    try {
        showStatus('조건을 만족하는 블록을 생성하는 중입니다...', false);
        generatedBlocks = generateRandomBlocks({
            blockCount,
            rows: rowsPerBlock,
            cols: colsPerBlock,
            avgMin,
            avgMax,
            valueMin,
            valueMax,
            maxAttempt
        });
        renderPreview(generatedBlocks);
        showStatus(`총 ${generatedBlocks.length}개 블록을 생성했습니다.`, true);
        downloadBtn.disabled = generatedBlocks.length === 0;
    } catch (error) {
        console.error(error);
        showStatus(error.message, false);
        generatedBlocks = [];
        renderPreview(generatedBlocks);
        downloadBtn.disabled = true;
    }
}

function handleDownload() {
    if (generatedBlocks.length === 0) {
        showStatus('먼저 데이터를 생성하세요.', false);
        return;
    }

    if (typeof ExcelJS === 'undefined') {
        showStatus('ExcelJS 라이브러리가 로드되지 않았습니다. 페이지를 새로고침해주세요.', false);
        return;
    }

    exportToExcel(generatedBlocks);
}

function validateInputs(values) {
    const { blockCount, avgMin, avgMax, valueMin, valueMax, maxAttempt } = values;

    if (Number.isNaN(blockCount) || blockCount < 1 || blockCount > 100) {
        showStatus('블록 수는 1~100 사이로 입력해주세요.', false);
        return false;
    }

    if (Number.isNaN(avgMin) || Number.isNaN(avgMax) || avgMin >= avgMax) {
        showStatus('평균값 최소/최대 범위를 올바르게 입력해주세요.', false);
        return false;
    }

    if (Number.isNaN(valueMin) || Number.isNaN(valueMax) || valueMin >= valueMax) {
        showStatus('데이터 값 최소/최대 범위를 올바르게 입력해주세요.', false);
        return false;
    }

    if (avgMin < valueMin || avgMax > valueMax) {
        showStatus('평균 범위는 값 범위 안에 포함되어야 합니다.', false);
        return false;
    }

    if (Number.isNaN(maxAttempt) || maxAttempt < 1) {
        showStatus('최대 시도 횟수는 1 이상이어야 합니다.', false);
        return false;
    }

    return true;
}

function generateRandomBlocks(options) {
    const { blockCount, rows, cols, avgMin, avgMax, valueMin, valueMax, maxAttempt } = options;
    const blocks = [];

    for (let blockIndex = 1; blockIndex <= blockCount; blockIndex++) {
        let attempt = 0;
        let success = false;
        let blockValues = [];
        let averageValue = 0;

        while (attempt < maxAttempt && !success) {
            attempt++;
            const { values, average } = createBlock(rows, cols, valueMin, valueMax);
            if (average >= avgMin && average <= avgMax) {
                success = true;
                blockValues = values;
                averageValue = average;
            }
        }

        if (!success) {
            throw new Error(`블록 ${blockIndex}에서 평균 조건을 만족하지 못했습니다. (시도 ${maxAttempt}회)`);
        }

        blocks.push({
            index: blockIndex,
            values: blockValues,
            average: parseFloat(averageValue.toFixed(2))
        });
    }

    return blocks;
}

function createBlock(rows, cols, min, max) {
    const values = [];
    let sum = 0;
    const totalCells = rows * cols;

    for (let r = 0; r < rows; r++) {
        const rowValues = [];
        for (let c = 0; c < cols; c++) {
            const val = getRandomInclusive(min, max);
            const roundedVal = parseFloat(val.toFixed(2));
            rowValues.push(roundedVal);
            sum += roundedVal;
        }
        values.push(rowValues);
    }

    const average = sum / totalCells;
    return { values, average };
}

function getRandomInclusive(min, max) {
    return min + Math.random() * (max - min);
}

function renderPreview(blocks) {
    previewContainer.innerHTML = '';

    if (!blocks.length) {
        previewHintEl.textContent = '조건을 충족하는 데이터를 먼저 생성하세요.';
        return;
    }

    previewHintEl.textContent = '생성된 블록 중 일부를 확인하세요.';

    blocks.forEach(block => {
        const card = document.createElement('div');
        card.className = 'block-card';

        const title = document.createElement('h4');
        title.textContent = `Block ${block.index}`;
        card.appendChild(title);

        const table = document.createElement('table');
        table.className = 'block-table';
        const tbody = document.createElement('tbody');

        block.values.forEach(rowValues => {
            const row = document.createElement('tr');
            rowValues.forEach(value => {
                const cell = document.createElement('td');
                cell.textContent = value.toFixed(2);
                row.appendChild(cell);
            });
            tbody.appendChild(row);
        });

        table.appendChild(tbody);
        card.appendChild(table);

        const avg = document.createElement('div');
        avg.className = 'block-average';
        avg.textContent = `평균: ${block.average.toFixed(2)}`;
        card.appendChild(avg);

        previewContainer.appendChild(card);
    });
}

function showStatus(message, isSuccess) {
    statusMessageEl.textContent = message;
    statusMessageEl.classList.toggle('success', isSuccess);
    statusMessageEl.style.display = 'block';
}

async function exportToExcel(blocks) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('랜덤 데이터');
    let currentRow = 1;

    blocks.forEach(block => {
        worksheet.mergeCells(currentRow, 1, currentRow, 5);
        const titleCell = worksheet.getCell(currentRow, 1);
        titleCell.value = `Block ${block.index}`;
        titleCell.font = { bold: true, size: 12 };
        titleCell.alignment = { vertical: 'middle', horizontal: 'left' };

        worksheet.getCell(currentRow, 7).value = '평균';
        const avgCell = worksheet.getCell(currentRow, 8);
        avgCell.value = block.average;
        avgCell.numFmt = '#,##0.00';
        avgCell.font = { bold: true };

        currentRow++;

        block.values.forEach(rowValues => {
            rowValues.forEach((value, idx) => {
                const cell = worksheet.getCell(currentRow, idx + 1);
                cell.value = value;
                cell.numFmt = '#,##0.00';
                cell.alignment = { horizontal: 'right' };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FFDDDDDD' } },
                    left: { style: 'thin', color: { argb: 'FFDDDDDD' } },
                    bottom: { style: 'thin', color: { argb: 'FFDDDDDD' } },
                    right: { style: 'thin', color: { argb: 'FFDDDDDD' } }
                };
            });
            currentRow++;
        });

        currentRow++; // 빈 행
    });

    worksheet.columns = [
        { width: 12 },
        { width: 12 },
        { width: 12 },
        { width: 12 },
        { width: 12 },
        { width: 4 },
        { width: 10 },
        { width: 12 }
    ];

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = '랜덤_블록_데이터.xlsx';
    link.click();
    window.URL.revokeObjectURL(url);
}

