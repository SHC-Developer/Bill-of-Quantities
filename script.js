// 전역 변수
let processedData = null;

// 열 설정 (기본값: C, G, I, N)
let columnSettings = {
    station: 'A',    // Station 열
    damage: 'C',     // 손상내용 열
    quantity: 'G',   // 물량 열
    count: 'I',      // 개소 열
    type: 'N'        // 구분 열
};

// 열 문자를 인덱스로 변환 (A=0, B=1, C=2, ...)
function columnToIndex(col) {
    if (!col || col.length === 0) return 0;
    const upperCol = col.toUpperCase();
    return upperCol.charCodeAt(0) - 65; // 'A' = 65, so A=0, B=1, etc.
}

// 설정 적용 버튼 이벤트
document.getElementById('applySettingsBtn').addEventListener('click', applySettings);

// 설정 적용 함수
function applySettings() {
    const colDamage = document.getElementById('colDamage').value.toUpperCase().trim();
    const colQuantity = document.getElementById('colQuantity').value.toUpperCase().trim();
    const colCount = document.getElementById('colCount').value.toUpperCase().trim();
    const colType = document.getElementById('colType').value.toUpperCase().trim();
    
    // 유효성 검사
    if (!colDamage || !colQuantity || !colCount || !colType) {
        alert('모든 열을 입력해주세요.');
        return;
    }
    
    if (!/^[A-Z]$/.test(colDamage) || !/^[A-Z]$/.test(colQuantity) || 
        !/^[A-Z]$/.test(colCount) || !/^[A-Z]$/.test(colType)) {
        alert('열은 A-Z 사이의 영문자 한 글자만 입력 가능합니다.');
        return;
    }
    
    // 설정 저장 (Station은 항상 A로 고정)
    columnSettings.station = 'A';
    columnSettings.damage = colDamage;
    columnSettings.quantity = colQuantity;
    columnSettings.count = colCount;
    columnSettings.type = colType;
    
    // 기존 데이터가 있으면 다시 처리
    if (processedData && document.getElementById('fileInput').files.length > 0) {
        const file = document.getElementById('fileInput').files[0];
        processFile(file);
    } else {
        alert('설정이 적용되었습니다. Excel 파일을 업로드하세요.');
    }
}

// 파일 입력 이벤트 리스너
document.getElementById('fileInput').addEventListener('change', handleFileSelect);

// 드래그 앤 드롭 이벤트
const uploadArea = document.getElementById('uploadArea');

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0 && (files[0].name.endsWith('.xlsx') || files[0].name.endsWith('.xls'))) {
        processFile(files[0]);
    } else {
        alert('Excel 파일(.xlsx, .xls)만 업로드 가능합니다.');
    }
});

// 파일 선택 핸들러
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        processFile(file);
    }
}

// 파일 처리 함수
function processFile(file) {
    const loading = document.getElementById('loading');
    const resultSection = document.getElementById('resultSection');
    
    loading.style.display = 'flex';
    resultSection.style.display = 'none';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Sheet1 읽기
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // JSON으로 변환
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                header: 1, 
                defval: '',
                raw: false 
            });
            
            // 데이터 처리
            processedData = processExcelData(jsonData);
            
            // 결과 표시
            displayResults(processedData);
            
            loading.style.display = 'none';
            resultSection.style.display = 'block';
            
        } catch (error) {
            console.error('파일 처리 중 오류:', error);
            alert('파일 처리 중 오류가 발생했습니다: ' + error.message);
            loading.style.display = 'none';
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// Excel 데이터 처리 함수 (VBA 로직 구현)
function processExcelData(data) {
    const stations = {};
    let currentStationKey = null;
    let stationIndex = 0;
    
    // 열 인덱스 매핑 (사용자 설정값 사용)
    const COL_STATION = columnToIndex(columnSettings.station);
    const COL_DAMAGE = columnToIndex(columnSettings.damage);
    const COL_QUANTITY = columnToIndex(columnSettings.quantity);
    const COL_COUNT = columnToIndex(columnSettings.count);
    const COL_TYPE = columnToIndex(columnSettings.type);
    
    // 데이터 읽기
    for (let i = 0; i < data.length; i++) {
        const row = data[i] || [];
        
        const cellStation = (row[COL_STATION] || '').toString().trim();
        const cellDamage = (row[COL_DAMAGE] || '').toString().trim();
        const cellQuantity = (row[COL_QUANTITY] || '').toString().trim();
        const cellCount = (row[COL_COUNT] || '').toString().trim();
        
        // Station 행 판단:
        // 1. Station 열에 값이 있고
        // 2. Damage 열이 비어있거나 (병합 셀의 경우)
        // 3. Damage 열에 값이 있어도 Quantity 열과 Count 열이 비어있는 경우 (헤더 행이 아닌 Station 행)
        const isStationRow = cellStation !== '' && 
                            (cellDamage === '' || (cellQuantity === '' && cellCount === '' && cellDamage !== '손상내용'));
        
        if (isStationRow) {
            stationIndex++;
            const uniqueKey = `ST_${stationIndex}_${cellStation}`;
            
            if (!stations[uniqueKey]) {
                stations[uniqueKey] = {
                    name: cellStation,
                    row: i,
                    data: {}
                };
            }
            currentStationKey = uniqueKey;
        }
        // 일반 데이터 행
        else if (currentStationKey && cellDamage !== '') {
            const 손상 = cellDamage;
            if (손상 === '') continue;
            
            const qty = parseFloat(row[COL_QUANTITY]) || 0;
            const pcs = parseInt(row[COL_COUNT]) || 0;
            const 구분 = (row[COL_TYPE] || '').toString().trim();
            
            // 손상별 데이터 초기화
            if (!stations[currentStationKey].data[손상]) {
                stations[currentStationKey].data[손상] = [0, 0, 0, 0]; // [전체물량, 전체개소, 신규물량, 신규개소]
            }
            
            const v = stations[currentStationKey].data[손상];
            
            // 1) 전체 물량/개소 집계 (보수는 제외)
            if (구분 !== '보수') {
                v[0] += qty;
                v[1] += pcs;
            }
            
            // 2) 신규 물량/개소 집계 (신규, 재손상, 재결함)
            if (구분 === '신규' || 구분 === '재손상' || 구분 === '재결함') {
                v[2] += qty;
                v[3] += pcs;
            }
            
            stations[currentStationKey].data[손상] = v;
        }
    }
    
    return stations;
}

// 결과 표시 함수
function displayResults(stations) {
    const container = document.getElementById('resultContainer');
    container.innerHTML = '';
    
    // Station별로 정렬 (원본 행 순서 유지)
    const stationKeys = Object.keys(stations).sort((a, b) => {
        return stations[a].row - stations[b].row;
    });
    
    stationKeys.forEach(key => {
        const station = stations[key];
        const stationDiv = document.createElement('div');
        stationDiv.className = 'station-table';
        
        // Station 이름
        const title = document.createElement('h3');
        title.className = 'station-title';
        title.textContent = station.name;
        stationDiv.appendChild(title);
        
        // 테이블 생성
        const table = document.createElement('table');
        table.className = 'result-table';
        
        // 헤더
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        ['손상내용', '전체물량', '전체개소', '신규물량', '신규개소'].forEach(text => {
            const th = document.createElement('th');
            th.textContent = text;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);
        table.appendChild(thead);
        
        // 본문
        const tbody = document.createElement('tbody');
        
        // 손상명 정렬
        const damageKeys = Object.keys(station.data).sort();
        
        // 물량이 0이 아닌 것만 출력
        damageKeys.forEach(damage => {
            const v = station.data[damage];
            if (v[0] === 0) return; // 전체물량이 0이면 제외
            
            const row = document.createElement('tr');
            
            const cells = [
                damage,
                formatNumber(v[0]),
                formatNumber(v[1]),
                formatNumber(v[2]),
                formatNumber(v[3])
            ];
            
            cells.forEach((text, idx) => {
                const td = document.createElement('td');
                td.textContent = text;
                if (idx > 0) { // 숫자 열은 오른쪽 정렬
                    td.style.textAlign = 'right';
                    td.className = 'number-cell';
                    td.setAttribute('data-value', v[idx - 1]);
                }
                row.appendChild(td);
            });
            
            tbody.appendChild(row);
        });
        
        table.appendChild(tbody);
        stationDiv.appendChild(table);
        container.appendChild(stationDiv);
    });
    
    // 다운로드 버튼 표시
    if (Object.keys(stations).length > 0) {
        document.getElementById('downloadBtn').style.display = 'inline-block';
    }
}

// 숫자 포맷팅 (천단위 구분 기호 추가)
function formatNumber(num) {
    if (num === 0) return '0';
    
    let formatted;
    if (num % 1 === 0) {
        formatted = num.toString();
    } else {
        formatted = num.toFixed(2);
    }
    
    // 천단위 구분 기호 추가
    const parts = formatted.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
}

// Excel 다운로드 함수 (스타일 적용)
document.getElementById('downloadBtn').addEventListener('click', async () => {
    if (!processedData) return;
    
    // ExcelJS 로드 확인
    if (typeof ExcelJS === 'undefined') {
        alert('ExcelJS 라이브러리가 로드되지 않았습니다. 페이지를 새로고침해주세요.');
        return;
    }
    
    try {
        // ExcelJS 워크북 생성
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('물량표');
        
        let currentRow = 1;
        
        // Station별로 정렬
        const stationKeys = Object.keys(processedData).sort((a, b) => {
            return processedData[a].row - processedData[b].row;
        });
        
        stationKeys.forEach(key => {
            const station = processedData[key];
            
            // ===========================
            // Station 이름 행
            // ===========================
            const titleRow = worksheet.getRow(currentRow);
            titleRow.getCell(1).value = station.name;
            titleRow.getCell(1).font = {
                name: '맑은 고딕',
                size: 12,
                bold: true
            };
            titleRow.getCell(1).alignment = {
                vertical: 'middle',
                horizontal: 'left'
            };
            titleRow.height = 25;
            currentRow++;
            
            // ===========================
            // 헤더 행
            // ===========================
            const headerRow = worksheet.getRow(currentRow);
            const headers = ['손상내용', '전체물량', '전체개소', '신규물량', '신규개소'];
            headers.forEach((header, idx) => {
                const cell = headerRow.getCell(idx + 1);
                cell.value = header;
                cell.font = {
                    name: '맑은 고딕',
                    size: 11,
                    bold: true,
                    color: { argb: 'FF2C3E50' }
                };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFDCDCDC' }
                };
                cell.alignment = {
                    vertical: 'middle',
                    horizontal: idx === 0 ? 'left' : 'right'
                };
                cell.border = {
                    top: { style: 'thin', color: { argb: 'FF999999' } },
                    left: { style: 'thin', color: { argb: 'FF999999' } },
                    bottom: { style: 'medium', color: { argb: 'FF999999' } },
                    right: { style: 'thin', color: { argb: 'FF999999' } }
                };
            });
            headerRow.height = 22;
            currentRow++;
            
            // ===========================
            // 데이터 행
            // ===========================
            const damageKeys = Object.keys(station.data).sort();
            const dataStartRow = currentRow;
            
            damageKeys.forEach(damage => {
                const v = station.data[damage];
                if (v[0] === 0) return; // 물량이 0이면 제외
                
                const dataRow = worksheet.getRow(currentRow);
                const rowData = [damage, v[0], v[1], v[2], v[3]];
                
                rowData.forEach((value, idx) => {
                    const cell = dataRow.getCell(idx + 1);
                    cell.value = value;
                    
                    if (idx === 0) {
                        // 손상내용 열
                        cell.font = {
                            name: '맑은 고딕',
                            size: 11,
                            bold: true,
                            color: { argb: 'FF2C3E50' }
                        };
                        cell.alignment = {
                            vertical: 'middle',
                            horizontal: 'left'
                        };
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                            left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                            bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                            right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
                        };
                    } else {
                        // 숫자 열
                        cell.font = {
                            name: '맑은 고딕',
                            size: 11,
                            color: { argb: 'FF1A1A1A' }
                        };
                        cell.alignment = {
                            vertical: 'middle',
                            horizontal: 'right'
                        };
                        cell.numFmt = '#,##0.00'; // 천단위 구분 기호 및 소수점
                        cell.border = {
                            top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                            left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                            bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                            right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
                        };
                    }
                });
                
                // 짝수 행 배경색
                if ((currentRow - dataStartRow) % 2 === 1) {
                    rowData.forEach((_, idx) => {
                        dataRow.getCell(idx + 1).fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFF8F9FA' }
                        };
                    });
                }
                
                dataRow.height = 20;
                currentRow++;
            });
            
            // ===========================
            // 테이블 테두리 적용
            // ===========================
            if (currentRow > dataStartRow) {
                const tableEndRow = currentRow - 1;
                const range = `A${dataStartRow - 1}:E${tableEndRow}`;
                
                // 외곽 테두리
                for (let row = dataStartRow - 1; row <= tableEndRow; row++) {
                    const rowObj = worksheet.getRow(row);
                    ['A', 'B', 'C', 'D', 'E'].forEach(col => {
                        const cell = rowObj.getCell(col);
                        if (!cell.border) {
                            cell.border = {
                                top: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                                left: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                                bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } },
                                right: { style: 'thin', color: { argb: 'FFE0E0E0' } }
                            };
                        }
                    });
                }
            }
            
            // Station 간격을 위한 빈 행
            currentRow++;
        });
        
        // ===========================
        // 열 너비 자동 조정
        // ===========================
        worksheet.columns = [
            { width: 25 }, // 손상내용
            { width: 12 }, // 전체물량
            { width: 12 }, // 전체개소
            { width: 12 }, // 신규물량
            { width: 12 }  // 신규개소
        ];
        
        // ===========================
        // 파일 다운로드
        // ===========================
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = '물량표_결과.xlsx';
        link.click();
        window.URL.revokeObjectURL(url);
        
    } catch (error) {
        console.error('Excel 생성 중 오류:', error);
        alert('Excel 파일 생성 중 오류가 발생했습니다: ' + error.message);
    }
});

// =======================
// 탭 전환 로직
// =======================
const tabButtons = document.querySelectorAll('.tab-btn');
const tabPanels = document.querySelectorAll('.tab-panel');

if (tabButtons.length > 0) {
    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            if (button.classList.contains('active')) return;

            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabPanels.forEach(panel => panel.classList.remove('active'));

            button.classList.add('active');
            const targetPanel = document.getElementById(button.dataset.target);
            if (targetPanel) {
                targetPanel.classList.add('active');
            }
        });
    });
}

// =======================
// 랜덤 블록 생성 로직
// =======================
const RANDOM_ROWS = 4;
const RANDOM_COLS = 5;
let randomBlocks = [];

const randomStatusEl = document.getElementById('randomStatusMessage');
const randomPreviewHintEl = document.getElementById('randomPreviewHint');
const randomPreviewContainer = document.getElementById('randomBlockPreview');
const randomGenerateBtn = document.getElementById('randomGenerateBtn');
const randomDownloadBtn = document.getElementById('randomDownloadBtn');

if (randomGenerateBtn && randomDownloadBtn) {
    randomGenerateBtn.addEventListener('click', randomHandleGenerate);
    randomDownloadBtn.addEventListener('click', randomHandleDownload);
}

function randomHandleGenerate() {
    const blockCount = parseInt(document.getElementById('randomBlockCount').value, 10);
    const avgMin = parseFloat(document.getElementById('randomAvgMin').value);
    const avgMax = parseFloat(document.getElementById('randomAvgMax').value);
    const valueMin = parseFloat(document.getElementById('randomValueMin').value);
    const valueMax = parseFloat(document.getElementById('randomValueMax').value);
    const maxAttempt = parseInt(document.getElementById('randomMaxAttempt').value, 10);

    if (!randomValidateInputs({ blockCount, avgMin, avgMax, valueMin, valueMax, maxAttempt })) {
        return;
    }

    try {
        randomShowStatus('조건을 만족하는 블록을 생성하는 중입니다...', false);
        randomBlocks = randomGenerateBlocks({
            blockCount,
            rows: RANDOM_ROWS,
            cols: RANDOM_COLS,
            avgMin,
            avgMax,
            valueMin,
            valueMax,
            maxAttempt
        });
        randomRenderPreview(randomBlocks);
        randomShowStatus(`총 ${randomBlocks.length}개 블록을 생성했습니다.`, true);
        randomDownloadBtn.disabled = randomBlocks.length === 0;
    } catch (error) {
        console.error(error);
        randomShowStatus(error.message, false);
        randomBlocks = [];
        randomRenderPreview(randomBlocks);
        randomDownloadBtn.disabled = true;
    }
}

function randomHandleDownload() {
    if (!randomBlocks.length) {
        randomShowStatus('먼저 데이터를 생성하세요.', false);
        return;
    }

    if (typeof ExcelJS === 'undefined') {
        randomShowStatus('ExcelJS 라이브러리가 로드되지 않았습니다. 페이지를 새로고침해주세요.', false);
        return;
    }

    randomExportToExcel(randomBlocks);
}

function randomValidateInputs(values) {
    const { blockCount, avgMin, avgMax, valueMin, valueMax, maxAttempt } = values;

    if (Number.isNaN(blockCount) || blockCount < 1 || blockCount > 100) {
        randomShowStatus('블록 수는 1~100 사이로 입력해주세요.', false);
        return false;
    }

    if (Number.isNaN(avgMin) || Number.isNaN(avgMax) || avgMin >= avgMax) {
        randomShowStatus('평균값 최소/최대 범위를 올바르게 입력해주세요.', false);
        return false;
    }

    if (Number.isNaN(valueMin) || Number.isNaN(valueMax) || valueMin >= valueMax) {
        randomShowStatus('데이터 값 최소/최대 범위를 올바르게 입력해주세요.', false);
        return false;
    }

    if (avgMin < valueMin || avgMax > valueMax) {
        randomShowStatus('평균 범위는 값 범위 안에 포함되어야 합니다.', false);
        return false;
    }

    if (Number.isNaN(maxAttempt) || maxAttempt < 1) {
        randomShowStatus('최대 시도 횟수는 1 이상이어야 합니다.', false);
        return false;
    }

    return true;
}

function randomGenerateBlocks(options) {
    const { blockCount, rows, cols, avgMin, avgMax, valueMin, valueMax, maxAttempt } = options;
    const blocks = [];

    for (let blockIndex = 1; blockIndex <= blockCount; blockIndex++) {
        let attempt = 0;
        let success = false;
        let blockValues = [];
        let averageValue = 0;

        while (attempt < maxAttempt && !success) {
            attempt++;
            const { values, average } = randomCreateBlock(rows, cols, valueMin, valueMax);
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

function randomCreateBlock(rows, cols, min, max) {
    const values = [];
    let sum = 0;
    const totalCells = rows * cols;

    for (let r = 0; r < rows; r++) {
        const rowValues = [];
        for (let c = 0; c < cols; c++) {
            const val = randomGetRandomInclusive(min, max);
            const roundedVal = parseFloat(val.toFixed(2));
            rowValues.push(roundedVal);
            sum += roundedVal;
        }
        values.push(rowValues);
    }

    const average = sum / totalCells;
    return { values, average };
}

function randomGetRandomInclusive(min, max) {
    return min + Math.random() * (max - min);
}

function randomRenderPreview(blocks) {
    if (!randomPreviewContainer || !randomPreviewHintEl) return;

    randomPreviewContainer.innerHTML = '';

    if (!blocks.length) {
        randomPreviewHintEl.textContent = '조건을 충족하는 데이터를 먼저 생성하세요.';
        return;
    }

    randomPreviewHintEl.textContent = '생성된 블록 중 일부를 확인하세요.';

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

        randomPreviewContainer.appendChild(card);
    });
}

function randomShowStatus(message, isSuccess) {
    if (!randomStatusEl) return;
    randomStatusEl.textContent = message;
    randomStatusEl.classList.toggle('success', isSuccess);
    randomStatusEl.style.display = 'block';
}

async function randomExportToExcel(blocks) {
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
