// 전역 변수
let excelData = [];
let filteredData = [];
let currentPage = 1;
const itemsPerPage = 10;

// 차트 인스턴스
let dailyChart = null;
let statusChart = null;
let requestTypeChart = null;
let processTypeChart = null;
let dateChart = null;
let fileNameChart = null;
let progressChart = null;

// DOM 로드 완료 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    console.log('DOM loaded, initializing app...');
    console.log('Chart.js available:', typeof Chart !== 'undefined');
    console.log('XLSX available:', typeof XLSX !== 'undefined');
    initializeApp();
});

// 헤더를 기반으로 열 인덱스 찾기
function findColumnIndices(headers) {
    const mapping = {
        요청일: -1,
        요청구분: -1,
        처리구분: -1,
        파일코드: -1,
        파일명: -1,
        진행상태: -1,
        개발진행: -1
    };
    
    // 헤더명 매칭 패턴 (정확한 한글 이름만 사용)
    const patterns = {
        요청일: ['요청일'],
        요청구분: ['요청구분'],
        처리구분: ['처리구분'],
        파일코드: ['파일코드'],
        파일명: ['파일명'],
        진행상태: ['진행상태'],
        개발진행: ['개발진행']
    };
    
    // 각 헤더를 순회하며 매칭
    headers.forEach((header, index) => {
        if (!header) return;
        
        const normalizedHeader = String(header).toLowerCase().trim().replace(/\s+/g, '');
        
        // 각 필드에 대해 패턴 매칭
        for (const [field, patternList] of Object.entries(patterns)) {
            if (mapping[field] !== -1) continue; // 이미 찾았으면 스킵
            
            for (const pattern of patternList) {
                const normalizedPattern = pattern.toLowerCase().replace(/\s+/g, '');
                // 정확히 일치하는 경우만
                if (normalizedHeader === normalizedPattern) {
                    mapping[field] = index;
                    console.log(`✓ Found "${field}" at column ${index} (${getColumnLetter(index)}): "${header}"`);
                    break;
                }
            }
        }
    });
    
    // 찾지 못한 필드 경고
    for (const [field, index] of Object.entries(mapping)) {
        if (index === -1) {
            console.warn(`⚠ Could not find column for "${field}"`);
        }
    }
    
    return mapping;
}

// 열 인덱스를 엑셀 열 문자로 변환 (0 -> A, 1 -> B, ...)
function getColumnLetter(index) {
    let letter = '';
    while (index >= 0) {
        letter = String.fromCharCode((index % 26) + 65) + letter;
        index = Math.floor(index / 26) - 1;
    }
    return letter;
}

// 앱 초기화
function initializeApp() {
    console.log('Initializing application...');
    
    const fileInput = document.getElementById('fileInput');
    const uploadBox = document.getElementById('uploadBox');
    const selectFileBtn = document.getElementById('selectFileBtn');
    const loadSampleBtn = document.getElementById('loadSampleBtn');
    
    if (!fileInput || !uploadBox) {
        console.error('Required elements not found!');
        return;
    }
    
    console.log('Elements found, setting up event listeners...');
    
    // 파일 입력 이벤트
    fileInput.addEventListener('change', handleFileSelect);
    
    // 파일 선택 버튼 클릭
    if (selectFileBtn) {
        selectFileBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            fileInput.click();
        });
    }
    
    // 샘플 데이터 로드 버튼 클릭
    if (loadSampleBtn) {
        loadSampleBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            loadSampleData();
        });
    }
    
    // 업로드 박스 클릭 (버튼 제외한 영역)
    uploadBox.addEventListener('click', (e) => {
        // 버튼이나 버튼 내부 요소를 클릭한 경우 무시
        if (e.target.closest('button')) {
            return;
        }
        fileInput.click();
    });
    
    // 드래그 앤 드롭 이벤트
    uploadBox.addEventListener('dragover', handleDragOver);
    uploadBox.addEventListener('dragleave', handleDragLeave);
    uploadBox.addEventListener('drop', handleDrop);
    
    // 탭 전환 이벤트
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });
    
    // 검색 및 필터 이벤트
    document.getElementById('searchInput').addEventListener('input', applyFilters);
    document.getElementById('statusFilter').addEventListener('change', applyFilters);
    document.getElementById('requestTypeFilter').addEventListener('change', applyFilters);
    
    console.log('Application initialized successfully');
}

// 파일 선택 처리
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
        processExcelFile(file);
    }
}

// 샘플 데이터 로드
async function loadSampleData() {
    try {
        console.log('Loading sample data...');
        const response = await fetch('./sample/UPDATE_DTS.xlsx');
        if (!response.ok) {
            console.error('Sample file not found:', response.status);
            alert('샘플 파일을 찾을 수 없습니다. 직접 엑셀 파일을 업로드해주세요.');
            return;
        }
        
        console.log('Sample file loaded, parsing...');
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        
        // 엑셀 파일 처리
        const workbook = XLSX.read(data, { type: 'array' });
        console.log('Sample workbook sheets:', workbook.SheetNames);
        
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        console.log('Sample data rows:', jsonData.length);
        
        if (jsonData.length < 2) {
            alert('데이터가 없습니다.');
            return;
        }
        
        const headers = jsonData[0];
        console.log('Sample headers:', headers);
        
        // 헤더를 기반으로 열 인덱스 찾기
        const columnMapping = findColumnIndices(headers);
        console.log('Sample auto-detected column mapping:', columnMapping);
        
        // 필수 필드가 모두 있는지 확인 (개발진행은 선택 필드)
        const requiredFields = ['요청일', '요청구분', '처리구분', '파일코드', '파일명', '진행상태'];
        const missingFields = [];
        for (const field of requiredFields) {
            if (columnMapping[field] === -1) {
                missingFields.push(field);
            }
        }
        
        // 개발진행 필드 확인 (선택 필드)
        if (columnMapping.개발진행 === -1) {
            console.warn('⚠ "개발진행" 필드를 찾을 수 없습니다. 개발진행 차트는 표시되지 않습니다.');
        }
        
        if (missingFields.length > 0) {
            alert(`샘플 파일에서 다음 필수 필드를 찾을 수 없습니다:\n${missingFields.join(', ')}\n\n헤더에 다음과 같은 이름이 포함되어야 합니다:\n- 요청일\n- 요청구분\n- 처리구분\n- 파일코드\n- 파일명\n- 진행상태`);
            console.error('Missing required fields:', missingFields);
            console.log('Available headers:', headers);
            return;
        }
        
        // 샘플 데이터 1개 확인
        if (jsonData.length > 1) {
            console.log('Sample row 1:', jsonData[1]);
        }
        
        excelData = [];
        
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            // 빈 행 제외 - 요청일 체크
            const requestDate = row[columnMapping.요청일];
            if (row.length > 0 && requestDate) {
                excelData.push({
                    요청일: formatDate(requestDate) || '',
                    요청구분: row[columnMapping.요청구분] || '',
                    처리구분: row[columnMapping.처리구분] || '',
                    파일코드: row[columnMapping.파일코드] || '',
                    파일명: row[columnMapping.파일명] || '',
                    진행상태: row[columnMapping.진행상태] || '',
                    개발진행: columnMapping.개발진행 !== -1 ? (row[columnMapping.개발진행] || '') : ''
                });
            }
        }
        
        console.log('Sample data processed:', excelData.length, 'records');
        
        // 처음 3개 데이터 샘플 출력
        if (excelData.length > 0) {
            console.log('Sample parsed data (first 3 records):');
            excelData.slice(0, 3).forEach((record, idx) => {
                console.log(`Record ${idx + 1}:`, record);
            });
        }
        
        if (excelData.length === 0) {
            alert('유효한 데이터가 없습니다.');
            return;
        }
        
        // UI 업데이트
        document.getElementById('uploadBox').style.display = 'none';
        document.getElementById('fileInfo').style.display = 'flex';
        document.getElementById('fileName').textContent = 'UPDATE_DTS.xlsx (샘플)';
        
        // 데이터 표시
        displayData();
        
        console.log('Sample data loaded successfully');
        
    } catch (error) {
        console.error('샘플 파일 로드 오류:', error);
        alert('샘플 파일을 로드하는 중 오류가 발생했습니다: ' + error.message);
    }
}

// 드래그 오버 처리
function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.classList.add('drag-over');
}

// 드래그 리브 처리
function handleDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.classList.remove('drag-over');
}

// 드롭 처리
function handleDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.classList.remove('drag-over');
    
    const file = event.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
        document.getElementById('fileInput').files = event.dataTransfer.files;
        processExcelFile(file);
    } else {
        alert('엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.');
    }
}

// 엑셀 파일 처리
function processExcelFile(file) {
    console.log('Processing file:', file.name);
    const reader = new FileReader();
    
    reader.onload = function(e) {
        try {
            console.log('File loaded, parsing...');
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            console.log('Workbook sheets:', workbook.SheetNames);
            
            // 첫 번째 시트 선택
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // JSON 변환
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            console.log('Parsed rows:', jsonData.length);
            
            // 데이터 파싱 (첫 번째 행은 헤더)
            if (jsonData.length < 2) {
                alert('데이터가 없습니다.');
                return;
            }
            
            const headers = jsonData[0];
            console.log('Headers:', headers);
            
            // 헤더를 기반으로 열 인덱스 찾기
            const columnMapping = findColumnIndices(headers);
            console.log('Auto-detected column mapping:', columnMapping);
            
            // 필수 필드가 모두 있는지 확인 (개발진행은 선택 필드)
            const requiredFields = ['요청일', '요청구분', '처리구분', '파일코드', '파일명', '진행상태'];
            const missingFields = [];
            for (const field of requiredFields) {
                if (columnMapping[field] === -1) {
                    missingFields.push(field);
                }
            }
            
            // 개발진행 필드 확인 (선택 필드)
            if (columnMapping.개발진행 === -1) {
                console.warn('⚠ "개발진행" 필드를 찾을 수 없습니다. 개발진행 차트는 표시되지 않습니다.');
            }
            
            if (missingFields.length > 0) {
                alert(`다음 필수 필드를 찾을 수 없습니다:\n${missingFields.join(', ')}\n\n헤더에 다음과 같은 이름이 포함되어야 합니다:\n- 요청일\n- 요청구분\n- 처리구분\n- 파일코드\n- 파일명\n- 진행상태`);
                console.error('Missing required fields:', missingFields);
                console.log('Available headers:', headers);
                return;
            }
            
            // 샘플 데이터 1개 확인
            if (jsonData.length > 1) {
                console.log('Sample row 1:', jsonData[1]);
            }
            
            excelData = [];
            
            for (let i = 1; i < jsonData.length; i++) {
                const row = jsonData[i];
                // 빈 행 제외 - 요청일 체크
                const requestDate = row[columnMapping.요청일];
                if (row.length > 0 && requestDate) {
                    excelData.push({
                        요청일: formatDate(requestDate) || '',
                        요청구분: row[columnMapping.요청구분] || '',
                        처리구분: row[columnMapping.처리구분] || '',
                        파일코드: row[columnMapping.파일코드] || '',
                        파일명: row[columnMapping.파일명] || '',
                        진행상태: row[columnMapping.진행상태] || '',
                        개발진행: columnMapping.개발진행 !== -1 ? (row[columnMapping.개발진행] || '') : ''
                    });
                }
            }
            
            console.log('Processed data records:', excelData.length);
            
            // 처음 3개 데이터 샘플 출력
            if (excelData.length > 0) {
                console.log('Sample parsed data (first 3 records):');
                excelData.slice(0, 3).forEach((record, idx) => {
                    console.log(`Record ${idx + 1}:`, record);
                });
            }
            
            if (excelData.length === 0) {
                alert('유효한 데이터가 없습니다.');
                return;
            }
            
            // UI 업데이트
            document.getElementById('uploadBox').style.display = 'none';
            document.getElementById('fileInfo').style.display = 'flex';
            document.getElementById('fileName').textContent = file.name;
            
            // 데이터 표시
            displayData();
            
        } catch (error) {
            console.error('파일 처리 오류:', error);
            alert('파일 처리 중 오류가 발생했습니다: ' + error.message);
        }
    };
    
    reader.onerror = function(error) {
        console.error('File read error:', error);
        alert('파일을 읽는 중 오류가 발생했습니다.');
    };
    
    reader.readAsArrayBuffer(file);
}

// 날짜 포맷 변환
function formatDate(value) {
    if (!value) return '';
    
    // 엑셀 날짜 숫자를 Date 객체로 변환
    if (typeof value === 'number') {
        const date = XLSX.SSF.parse_date_code(value);
        return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
    }
    
    // 이미 문자열인 경우
    if (typeof value === 'string') {
        return value;
    }
    
    return String(value);
}

// 데이터 표시
function displayData() {
    console.log('displayData called with', excelData.length, 'records');
    filteredData = [...excelData];
    
    // 요약 정보 업데이트
    updateSummary();
    
    // 날짜 범위 업데이트
    updateDateRange();
    
    // 필터 옵션 생성
    populateFilters();
    
    // 테이블 표시
    displayTable();
    
    // 차트 생성
    console.log('Creating charts...');
    createCharts();
    
    // 섹션 표시
    document.getElementById('summarySection').style.display = 'grid';
    document.getElementById('dateRangeSection').style.display = 'block';
    document.getElementById('tabsSection').style.display = 'block';
    
    console.log('Data display completed');
}

// 요약 정보 업데이트
function updateSummary() {
    const total = filteredData.length;
    const completed = filteredData.filter(item => 
        item.진행상태 && item.진행상태.includes('완료')
    ).length;
    const inProgress = filteredData.filter(item => 
        item.진행상태 && item.진행상태.includes('진행')
    ).length;
    const pending = filteredData.filter(item => 
        item.진행상태 && (item.진행상태.includes('대기') || item.진행상태.includes('요청'))
    ).length;
    
    document.getElementById('totalCount').textContent = total;
    document.getElementById('completedCount').textContent = completed;
    document.getElementById('inProgressCount').textContent = inProgress;
    document.getElementById('pendingCount').textContent = pending;
}

// 날짜 범위 업데이트
function updateDateRange() {
    const dates = filteredData
        .map(item => item.요청일)
        .filter(date => date && date !== '미정' && date !== '')
        .sort();
    
    if (dates.length === 0) {
        document.getElementById('dateRangeText').textContent = '날짜 정보 없음';
        return;
    }
    
    const startDate = dates[0];
    const endDate = dates[dates.length - 1];
    
    if (startDate === endDate) {
        document.getElementById('dateRangeText').textContent = startDate;
    } else {
        document.getElementById('dateRangeText').textContent = `${startDate} ~ ${endDate}`;
    }
    
    console.log('Date range:', startDate, '~', endDate);
}

// 필터 옵션 채우기
function populateFilters() {
    const statusSet = new Set();
    const requestTypeSet = new Set();
    
    excelData.forEach(item => {
        if (item.진행상태) statusSet.add(item.진행상태);
        if (item.요청구분) requestTypeSet.add(item.요청구분);
    });
    
    // 진행상태 필터
    const statusFilter = document.getElementById('statusFilter');
    statusFilter.innerHTML = '<option value="">전체 진행상태</option>';
    statusSet.forEach(status => {
        statusFilter.innerHTML += `<option value="${status}">${status}</option>`;
    });
    
    // 요청구분 필터
    const requestTypeFilter = document.getElementById('requestTypeFilter');
    requestTypeFilter.innerHTML = '<option value="">전체 요청구분</option>';
    requestTypeSet.forEach(type => {
        requestTypeFilter.innerHTML += `<option value="${type}">${type}</option>`;
    });
}

// 필터 적용
function applyFilters() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const statusFilter = document.getElementById('statusFilter').value;
    const requestTypeFilter = document.getElementById('requestTypeFilter').value;
    
    filteredData = excelData.filter(item => {
        const matchSearch = !searchTerm || 
            Object.values(item).some(val => 
                String(val).toLowerCase().includes(searchTerm)
            );
        
        const matchStatus = !statusFilter || item.진행상태 === statusFilter;
        const matchRequestType = !requestTypeFilter || item.요청구분 === requestTypeFilter;
        
        return matchSearch && matchStatus && matchRequestType;
    });
    
    currentPage = 1;
    updateSummary();
    updateDateRange();
    displayTable();
}

// 테이블 표시
function displayTable() {
    const tbody = document.getElementById('dataTableBody');
    const startIndex = (currentPage - 1) * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const pageData = filteredData.slice(startIndex, endIndex);
    
    tbody.innerHTML = '';
    
    pageData.forEach((item, index) => {
        const row = document.createElement('tr');
        const statusClass = getStatusClass(item.진행상태);
        
        row.innerHTML = `
            <td>${startIndex + index + 1}</td>
            <td>${item.요청일}</td>
            <td>${item.요청구분}</td>
            <td>${item.처리구분}</td>
            <td>${item.파일코드}</td>
            <td>${item.파일명}</td>
            <td><span class="status-badge ${statusClass}">${item.진행상태}</span></td>
        `;
        
        tbody.appendChild(row);
    });
    
    // 페이지네이션 업데이트
    updatePagination();
}

// 상태 클래스 가져오기
function getStatusClass(status) {
    if (!status) return '';
    if (status.includes('완료')) return 'status-completed';
    if (status.includes('진행')) return 'status-inprogress';
    if (status.includes('대기') || status.includes('요청')) return 'status-pending';
    return '';
}

// 페이지네이션 업데이트
function updatePagination() {
    const totalPages = Math.ceil(filteredData.length / itemsPerPage);
    const pagination = document.getElementById('pagination');
    
    pagination.innerHTML = '';
    
    // 이전 버튼
    const prevBtn = document.createElement('button');
    prevBtn.innerHTML = '<i class="fas fa-chevron-left"></i>';
    prevBtn.disabled = currentPage === 1;
    prevBtn.onclick = () => {
        if (currentPage > 1) {
            currentPage--;
            displayTable();
        }
    };
    pagination.appendChild(prevBtn);
    
    // 페이지 버튼
    const startPage = Math.max(1, currentPage - 2);
    const endPage = Math.min(totalPages, startPage + 4);
    
    for (let i = startPage; i <= endPage; i++) {
        const pageBtn = document.createElement('button');
        pageBtn.textContent = i;
        pageBtn.className = i === currentPage ? 'active' : '';
        pageBtn.onclick = () => {
            currentPage = i;
            displayTable();
        };
        pagination.appendChild(pageBtn);
    }
    
    // 다음 버튼
    const nextBtn = document.createElement('button');
    nextBtn.innerHTML = '<i class="fas fa-chevron-right"></i>';
    nextBtn.disabled = currentPage === totalPages;
    nextBtn.onclick = () => {
        if (currentPage < totalPages) {
            currentPage++;
            displayTable();
        }
    };
    pagination.appendChild(nextBtn);
}

// 차트 생성
function createCharts() {
    createDailyChart();
    createStatusChart();
    createRequestTypeChart();
    createProcessTypeChart();
    createDateChart();
    createFileNameChart();
    createProgressChart();
}

// 요청일자별 집계 차트 (새로 추가)
function createDailyChart() {
    try {
        // 날짜별로 데이터 집계
        const dailyData = {};
        
        filteredData.forEach(item => {
            const date = item.요청일 || '미정';
            if (!dailyData[date]) {
                dailyData[date] = {
                    전체: 0,
                    완료: 0,
                    진행중: 0,
                    대기: 0
                };
            }
            
            dailyData[date].전체++;
            
            const status = item.진행상태 || '';
            if (status.includes('완료')) {
                dailyData[date].완료++;
            } else if (status.includes('진행')) {
                dailyData[date].진행중++;
            } else if (status.includes('대기') || status.includes('요청')) {
                dailyData[date].대기++;
            }
        });
        
        // 날짜순 정렬
        const sortedDates = Object.keys(dailyData).sort();
        
        const totalData = sortedDates.map(date => dailyData[date].전체);
        const completedData = sortedDates.map(date => dailyData[date].완료);
        const inProgressData = sortedDates.map(date => dailyData[date].진행중);
        const pendingData = sortedDates.map(date => dailyData[date].대기);
        
        const canvas = document.getElementById('dailyChart');
        if (!canvas) {
            console.error('dailyChart canvas not found');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        if (dailyChart) {
            dailyChart.destroy();
            dailyChart = null;
        }
        
        dailyChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: sortedDates,
                datasets: [
                    {
                        label: '완료',
                        data: completedData,
                        backgroundColor: '#10b981',
                        borderColor: '#059669',
                        borderWidth: 1
                    },
                    {
                        label: '진행중',
                        data: inProgressData,
                        backgroundColor: '#f59e0b',
                        borderColor: '#d97706',
                        borderWidth: 1
                    },
                    {
                        label: '대기',
                        data: pendingData,
                        backgroundColor: '#ef4444',
                        borderColor: '#dc2626',
                        borderWidth: 1
                    },
                    {
                        label: '전체',
                        data: totalData,
                        type: 'line',
                        backgroundColor: 'rgba(37, 99, 235, 0.1)',
                        borderColor: '#2563eb',
                        borderWidth: 3,
                        fill: false,
                        tension: 0.4,
                        pointRadius: 5,
                        pointBackgroundColor: '#2563eb',
                        yAxisID: 'y'
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: {
                    mode: 'index',
                    intersect: false
                },
                plugins: {
                    legend: {
                        display: true,
                        position: 'top',
                        labels: {
                            usePointStyle: true,
                            padding: 15,
                            font: {
                                size: 12,
                                weight: 'bold'
                            }
                        }
                    },
                    tooltip: {
                        callbacks: {
                            footer: function(tooltipItems) {
                                let total = 0;
                                tooltipItems.forEach(item => {
                                    if (item.dataset.label !== '전체') {
                                        total += item.parsed.y;
                                    }
                                });
                                return '합계: ' + total + '건';
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        grid: {
                            display: false
                        },
                        ticks: {
                            maxRotation: 45,
                            minRotation: 0,
                            font: {
                                size: 10
                            }
                        }
                    },
                    y: {
                        stacked: true,
                        beginAtZero: true,
                        ticks: {
                            stepSize: 1,
                            font: {
                                size: 11
                            }
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.05)'
                        }
                    }
                }
            }
        });
        
        console.log('Daily chart created successfully');
    } catch (error) {
        console.error('Error creating daily chart:', error);
    }
}

// 진행상태별 차트
function createStatusChart() {
    try {
        const statusCount = {};
        filteredData.forEach(item => {
            const status = item.진행상태 || '미정';
            statusCount[status] = (statusCount[status] || 0) + 1;
        });
        
        const canvas = document.getElementById('statusChart');
        if (!canvas) {
            console.error('statusChart canvas not found');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        if (statusChart) {
            statusChart.destroy();
            statusChart = null;
        }
        
        statusChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: Object.keys(statusCount),
                datasets: [{
                    data: Object.values(statusCount),
                    backgroundColor: [
                        '#10b981',
                        '#f59e0b',
                        '#ef4444',
                        '#6366f1',
                        '#ec4899'
                    ]
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom'
                    }
                }
            }
        });
        
        console.log('Status chart created successfully');
    } catch (error) {
        console.error('Error creating status chart:', error);
    }
}

// 요청구분별 차트
function createRequestTypeChart() {
    try {
        const requestTypeCount = {};
        filteredData.forEach(item => {
            const type = item.요청구분 || '미정';
            requestTypeCount[type] = (requestTypeCount[type] || 0) + 1;
        });
        
        const canvas = document.getElementById('requestTypeChart');
        if (!canvas) {
            console.error('requestTypeChart canvas not found');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        if (requestTypeChart) {
            requestTypeChart.destroy();
            requestTypeChart = null;
        }
        
        requestTypeChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: Object.keys(requestTypeCount),
                datasets: [{
                    label: '요청 수',
                    data: Object.values(requestTypeCount),
                    backgroundColor: '#2563eb'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            stepSize: 1
                        }
                    }
                }
            }
        });
        
        console.log('Request type chart created successfully');
    } catch (error) {
        console.error('Error creating request type chart:', error);
    }
}

// 처리구분별 차트
function createProcessTypeChart() {
    try {
        const processTypeCount = {};
        filteredData.forEach(item => {
            const type = item.처리구분 || '미정';
            processTypeCount[type] = (processTypeCount[type] || 0) + 1;
        });
        
        const canvas = document.getElementById('processTypeChart');
        if (!canvas) {
            console.error('processTypeChart canvas not found');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        if (processTypeChart) {
            processTypeChart.destroy();
            processTypeChart = null;
        }
        
        processTypeChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: Object.keys(processTypeCount),
                datasets: [{
                    label: '처리 수',
                    data: Object.values(processTypeCount),
                    backgroundColor: '#f59e0b',
                    borderColor: '#d97706',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            stepSize: 1
                        }
                    }
                }
            }
        });
        
        console.log('Process type chart created successfully');
    } catch (error) {
        console.error('Error creating process type chart:', error);
    }
}

// 일자별 차트
function createDateChart() {
    try {
        const dateCount = {};
        filteredData.forEach(item => {
            const date = item.요청일 || '미정';
            dateCount[date] = (dateCount[date] || 0) + 1;
        });
        
        // 날짜순 정렬
        const sortedDates = Object.keys(dateCount).sort();
        const sortedCounts = sortedDates.map(date => dateCount[date]);
        
        const canvas = document.getElementById('dateChart');
        if (!canvas) {
            console.error('dateChart canvas not found');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        if (dateChart) {
            dateChart.destroy();
            dateChart = null;
        }
        
        dateChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: sortedDates,
                datasets: [{
                    label: '요청 수',
                    data: sortedCounts,
                    borderColor: '#2563eb',
                    backgroundColor: 'rgba(37, 99, 235, 0.1)',
                    tension: 0.4,
                    fill: true
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: true
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            stepSize: 1
                        }
                    }
                }
            }
        });
        
        console.log('Date chart created successfully');
    } catch (error) {
        console.error('Error creating date chart:', error);
    }
}

// 문의 분포 차트
function createFileNameChart() {
    try {
        const fileNameCount = {};
        filteredData.forEach(item => {
            const fileName = item.파일명 || '미정';
            fileNameCount[fileName] = (fileNameCount[fileName] || 0) + 1;
        });
        
        // 상위 10개만 선택
        const sortedFileNames = Object.entries(fileNameCount)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10);
        
        const labels = sortedFileNames.map(item => {
            const name = item[0];
            // 파일명이 너무 길면 축약
            return name.length > 20 ? name.substring(0, 20) + '...' : name;
        });
        const data = sortedFileNames.map(item => item[1]);
        
        const canvas = document.getElementById('fileNameChart');
        if (!canvas) {
            console.error('fileNameChart canvas not found');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        if (fileNameChart) {
            fileNameChart.destroy();
            fileNameChart = null;
        }
        
        fileNameChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: '문의 수',
                    data: data,
                    backgroundColor: '#8b5cf6',
                    borderColor: '#7c3aed',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                indexAxis: 'y', // 수평 막대 그래프
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        callbacks: {
                            title: function(context) {
                                // 툴팁에서는 전체 파일명 표시
                                return sortedFileNames[context[0].dataIndex][0];
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        beginAtZero: true,
                        ticks: {
                            stepSize: 1
                        }
                    }
                }
            }
        });
        
        console.log('File name chart created successfully');
    } catch (error) {
        console.error('Error creating file name chart:', error);
    }
}

// 개발진행 현황 차트
function createProgressChart() {
    try {
        console.log('=== Creating progress chart ===');
        console.log('Total filtered data:', filteredData.length);
        
        // 개발진행 Y/N 데이터 집계
        const progressData = {
            필요: 0,      // Y - 개발진행 필요
            불필요: 0     // N - 개발진행 불필요
        };
        
        const uniqueValues = new Set();
        
        filteredData.forEach((item, idx) => {
            const rawValue = item.개발진행;
            const devProgress = String(rawValue || '').toUpperCase().trim();
            
            uniqueValues.add(devProgress);
            
            if (idx < 3) { // 처음 3개만 상세 로그
                console.log(`Item ${idx}:`, {
                    rawValue: rawValue,
                    stringified: String(rawValue),
                    normalized: devProgress,
                    type: typeof rawValue
                });
            }
            
            if (devProgress === 'Y') {
                progressData.필요++;
            } else if (devProgress === 'N') {
                progressData.불필요++;
            }
        });
        
        console.log('Unique dev progress values:', Array.from(uniqueValues));
        console.log('Progress data:', progressData);
        
        const total = progressData.필요 + progressData.불필요;
        
        // 데이터가 없는 경우 기본 메시지 표시
        if (total === 0) {
            console.warn('⚠️ No development progress data (Y/N) found');
            console.warn('Check if "개발진행" column exists and has Y/N values');
        }
        
        const percentages = {
            필요: total > 0 ? ((progressData.필요 / total) * 100).toFixed(1) : 0,
            불필요: total > 0 ? ((progressData.불필요 / total) * 100).toFixed(1) : 0
        };
        
        const canvas = document.getElementById('progressChart');
        if (!canvas) {
            console.error('progressChart canvas not found');
            return;
        }
        
        const ctx = canvas.getContext('2d');
        
        if (progressChart) {
            progressChart.destroy();
            progressChart = null;
        }
        
        progressChart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: ['개발진행 필요 (Y)', '개발진행 불필요 (N)'],
                datasets: [{
                    data: total > 0 ? [progressData.필요, progressData.불필요] : [1, 1],
                    backgroundColor: [
                        '#10b981', // 개발진행 필요(Y) - 녹색
                        '#9ca3af'  // 개발진행 불필요(N) - 회색
                    ],
                    borderWidth: 3,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            padding: 20,
                            usePointStyle: true,
                            font: {
                                size: 13,
                                weight: 'bold'
                            }
                        }
                    },
                    tooltip: {
                        enabled: total > 0,
                        callbacks: {
                            label: function(context) {
                                if (total === 0) {
                                    return '데이터 없음';
                                }
                                const label = context.label || '';
                                const value = context.parsed || 0;
                                const isRequired = label.includes('필요 (Y)');
                                const percentage = isRequired ? percentages.필요 : percentages.불필요;
                                return `${label}: ${value}건 (${percentage}%)`;
                            }
                        }
                    },
                    title: {
                        display: total === 0,
                        text: '개발진행 데이터 없음 (Y/N 값 필요)',
                        font: {
                            size: 14
                        },
                        color: '#6b7280'
                    }
                }
            }
        });
        
        console.log('✅ Progress chart created - 필요(Y):', progressData.필요, '불필요(N):', progressData.불필요);
    } catch (error) {
        console.error('❌ Error creating progress chart:', error);
    }
}

// 탭 전환
function switchTab(tabName) {
    // 탭 버튼 활성화
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');
    
    // 탭 콘텐츠 표시
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    if (tabName === 'data') {
        document.getElementById('dataTab').classList.add('active');
    } else if (tabName === 'analytics') {
        document.getElementById('analyticsTab').classList.add('active');
        // 차트 다시 그리기 (크기 조정)
        setTimeout(() => {
            if (dailyChart) dailyChart.resize();
            if (statusChart) statusChart.resize();
            if (requestTypeChart) requestTypeChart.resize();
            if (processTypeChart) processTypeChart.resize();
            if (dateChart) dateChart.resize();
            if (fileNameChart) fileNameChart.resize();
            if (progressChart) progressChart.resize();
        }, 100);
    }
}

// 업로드 초기화
function resetUpload() {
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadBox').style.display = 'block';
    document.getElementById('fileInfo').style.display = 'none';
    document.getElementById('summarySection').style.display = 'none';
    document.getElementById('dateRangeSection').style.display = 'none';
    document.getElementById('tabsSection').style.display = 'none';
    
    excelData = [];
    filteredData = [];
    currentPage = 1;
    
    // 차트 제거
    if (dailyChart) dailyChart.destroy();
    if (statusChart) statusChart.destroy();
    if (requestTypeChart) requestTypeChart.destroy();
    if (processTypeChart) processTypeChart.destroy();
    if (dateChart) dateChart.destroy();
    if (fileNameChart) fileNameChart.destroy();
    if (progressChart) progressChart.destroy();
}
