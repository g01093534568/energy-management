// ==================== 엑셀 다운로드/업로드 기능 ====================

// formatBillingMonth 함수 (app.js에서 복사)
function formatBillingMonth(monthStr) {
    if (!monthStr) return '-';

    // 이미 "년 월분" 형식이면 그대로 반환
    if (monthStr.includes('년') && monthStr.includes('월분')) {
        return monthStr;
    }

    // YYYY-MM 형식을 YYYY년 MM월분으로 변환
    const match = monthStr.match(/(\d{4})-(\d{2})/);
    if (match) {
        const year = match[1];
        const month = parseInt(match[2], 10); // 앞의 0 제거
        return `${year}년 ${month}월분`;
    }

    return monthStr;
}

// 엑셀 다운로드 버튼 이벤트
document.addEventListener('DOMContentLoaded', () => {
    const downloadExcelBtn = document.getElementById('downloadExcelBtn');
    const uploadExcelBtn = document.getElementById('uploadExcelBtn');
    const excelFileInput = document.getElementById('excelFileInput');

    if (downloadExcelBtn) {
        downloadExcelBtn.addEventListener('click', downloadDataAsExcel);
    }

    if (uploadExcelBtn) {
        uploadExcelBtn.addEventListener('click', () => {
            excelFileInput.click();
        });
    }

    if (excelFileInput) {
        excelFileInput.addEventListener('change', handleExcelUpload);
    }
});

// 엑셀 다운로드 함수
async function downloadDataAsExcel() {
    try {
        // XLSX 라이브러리 확인
        if (typeof XLSX === 'undefined') {
            alert('엑셀 라이브러리가 로드되지 않았습니다. 페이지를 새로고침해주세요.');
            console.error('XLSX is not defined');
            return;
        }

        // 현재 조회된 데이터 사용 (app.js의 currentViewRecords 사용)
        const records = window.currentViewRecords || [];

        console.log('다운로드할 레코드 수:', records.length);
        console.log('레코드 샘플:', records[0]);

        if (records.length === 0) {
            alert('다운로드할 데이터가 없습니다.\n먼저 "조회" 버튼을 눌러 데이터를 조회해주세요.');
            return;
        }

        // 에너지 종류별로 데이터 분류
        const energyTypes = {
            '전기': [],
            '상하수도': [],
            '도시가스': [],
            '통신': []
        };

        records.forEach(record => {
            if (energyTypes[record.energyType]) {
                energyTypes[record.energyType].push(record);
            }
        });

        // 워크북 생성
        const wb = XLSX.utils.book_new();

        // 각 에너지 종류별로 시트 생성
        Object.keys(energyTypes).forEach(type => {
            const typeRecords = energyTypes[type];

            if (typeRecords.length > 0) {
                // 데이터를 2차원 배열로 변환
                const worksheetData = [
                    ['', '', type + '요금 납부 내역', '', '', '', ''],
                    [],
                    [],
                    ['년도', '월', '시설명', '월분', '사용기간', '사용량', '납부금액', '비고']
                ];

                typeRecords.forEach(record => {
                    // 날짜에서 년도와 월 추출
                    const startDate = new Date(record.startDate);
                    const year = startDate.getFullYear() + '년';
                    const month = (startDate.getMonth() + 1) + '월';

                    // 사용기간 포맷팅
                    const period = record.startDate + ' ~ ' + record.endDate;

                    // 월분 포맷팅
                    let billingMonth = '-';
                    if (record.billingMonth) {
                        billingMonth = formatBillingMonth(record.billingMonth);
                    } else {
                        // billingMonth가 없으면 startDate에서 추출
                        const bYear = startDate.getFullYear();
                        const bMonth = startDate.getMonth() + 1;
                        billingMonth = `${bYear}년 ${bMonth}월분`;
                    }

                    // 사용량 단위
                    let unit = 'kWh';
                    if (type === '상하수도' || type === '도시가스') {
                        unit = 'm³';
                    } else if (type === '통신') {
                        unit = '건';
                    }

                    worksheetData.push([
                        year,
                        month,
                        record.facilityName || '',
                        billingMonth,
                        period,
                        record.usageAmount + ' ' + unit,
                        record.usageCost,
                        ''
                    ]);
                });

                // 시트 생성
                const ws = XLSX.utils.aoa_to_sheet(worksheetData);

                // 열 너비 설정
                ws['!cols'] = [
                    { wch: 8 },   // 년도
                    { wch: 6 },   // 월
                    { wch: 20 },  // 시설명
                    { wch: 15 },  // 월분
                    { wch: 25 },  // 사용기간
                    { wch: 15 },  // 사용량
                    { wch: 12 },  // 납부금액
                    { wch: 10 }   // 비고
                ];

                // 워크북에 시트 추가
                XLSX.utils.book_append_sheet(wb, ws, type + ' 납부내역');
            }
        });

        // 전체 데이터 요약 시트 생성
        const summaryData = [
            ['에너지 사용 데이터 요약'],
            [],
            ['에너지 종류', '총 사용량', '총 금액', '데이터 건수']
        ];

        Object.keys(energyTypes).forEach(type => {
            const typeRecords = energyTypes[type];
            const totalUsage = typeRecords.reduce((sum, r) => sum + parseFloat(r.usageAmount || 0), 0);
            const totalCost = typeRecords.reduce((sum, r) => sum + parseFloat(r.usageCost || 0), 0);

            let unit = 'kWh';
            if (type === '상하수도' || type === '도시가스') {
                unit = 'm³';
            } else if (type === '통신') {
                unit = '건';
            }

            summaryData.push([
                type,
                totalUsage.toLocaleString() + ' ' + unit,
                totalCost.toLocaleString() + '원',
                typeRecords.length
            ]);
        });

        const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
        summaryWs['!cols'] = [
            { wch: 15 },
            { wch: 15 },
            { wch: 15 },
            { wch: 12 }
        ];

        XLSX.utils.book_append_sheet(wb, summaryWs, '요약', 0); // 첫 번째 시트로 추가

        // 파일명 생성 (현재 날짜 포함)
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, '0');
        const dd = String(today.getDate()).padStart(2, '0');
        const filename = '에너지사용데이터_' + yyyy + '_' + mm + '_' + dd + '.xlsx';

        // 파일 다운로드
        XLSX.writeFile(wb, filename);

        console.log('엑셀 파일 다운로드 완료:', filename);
        alert('엑셀 파일이 다운로드되었습니다.');

    } catch (error) {
        console.error('엑셀 다운로드 오류:', error);
        alert('엑셀 다운로드 중 오류가 발생했습니다: ' + error.message);
    }
}

// 엑셀 업로드 처리 함수
async function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);

        console.log('업로드된 엑셀 시트:', workbook.SheetNames);

        let totalImported = 0;
        const importedData = [];

        // 각 시트 처리
        for (const sheetName of workbook.SheetNames) {
            // 납부내역 시트만 처리
            if (!sheetName.includes('납부내역') && !sheetName.includes('요금')) {
                continue;
            }

            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            console.log('시트 "' + sheetName + '" 처리 중...');

            // 에너지 종류 판단
            let energyType = '';
            if (sheetName.includes('전기')) energyType = '전기';
            else if (sheetName.includes('상하수도')) energyType = '상하수도';
            else if (sheetName.includes('도시가스')) energyType = '도시가스';
            else if (sheetName.includes('통신')) energyType = '통신';

            if (!energyType) {
                console.log('시트 "' + sheetName + '"의 에너지 종류를 판단할 수 없습니다. 건너뜁니다.');
                continue;
            }

            // 헤더 행 찾기 (년도, 월, 사용량, 납부금액 등이 있는 행)
            let headerRowIndex = -1;
            for (let i = 0; i < Math.min(10, jsonData.length); i++) {
                const row = jsonData[i];
                const rowStr = row.join('').toLowerCase();
                if (rowStr.includes('년도') || rowStr.includes('사용량') || rowStr.includes('납부금액')) {
                    headerRowIndex = i;
                    break;
                }
            }

            if (headerRowIndex === -1) {
                console.log('시트 "' + sheetName + '"에서 헤더를 찾을 수 없습니다.');
                continue;
            }

            const headers = jsonData[headerRowIndex];
            console.log('헤더:', headers);

            // 헤더에서 컬럼 인덱스 찾기
            const yearIdx = headers.findIndex(h => h && h.toString().includes('년도'));
            const monthIdx = headers.findIndex(h => h && h.toString().includes('월'));
            const facilityIdx = headers.findIndex(h => h && h.toString().includes('시설명'));
            const periodIdx = headers.findIndex(h => h && h.toString().includes('사용기간'));
            const usageIdx = headers.findIndex(h => h && h.toString().includes('사용량'));
            const costIdx = headers.findIndex(h => h && h.toString().includes('납부금액') || h && h.toString().includes('금액'));

            // 데이터 행 처리 (헤더 다음 행부터)
            for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
                const row = jsonData[i];

                // 빈 행이나 합계 행 건너뛰기
                if (!row || row.length === 0 || !row[usageIdx] || !row[costIdx]) {
                    continue;
                }

                const rowStr = row.join('').toLowerCase();
                if (rowStr.includes('합계') || rowStr.includes('계')) {
                    continue;
                }

                try {
                    const usageAmount = parseFloat(row[usageIdx].toString().replace(/[^0-9.]/g, ''));
                    const usageCost = parseFloat(row[costIdx].toString().replace(/[^0-9.]/g, ''));

                    if (isNaN(usageAmount) || isNaN(usageCost)) {
                        continue;
                    }

                    // 사용기간 파싱
                    let startDate, endDate;
                    if (periodIdx >= 0 && row[periodIdx]) {
                        const period = row[periodIdx].toString();
                        const dates = period.split(/[-~～]/);
                        if (dates.length >= 2) {
                            startDate = dates[0].trim();
                            endDate = dates[1].trim();
                        }
                    }

                    // 날짜가 없으면 년도/월에서 생성
                    if (!startDate && yearIdx >= 0 && monthIdx >= 0) {
                        const year = row[yearIdx].toString().replace(/[^0-9]/g, '');
                        const month = row[monthIdx].toString().replace(/[^0-9]/g, '').padStart(2, '0');
                        if (year && month) {
                            startDate = year + '-' + month + '-01';
                            // 해당 월의 마지막 날
                            const lastDay = new Date(year, month, 0).getDate();
                            endDate = year + '-' + month + '-' + String(lastDay).padStart(2, '0');
                        }
                    }

                    if (!startDate || !endDate) {
                        console.log('행 ' + (i + 1) + ': 날짜 정보를 파싱할 수 없습니다.');
                        continue;
                    }

                    const recordData = {
                        facilityName: facilityIdx >= 0 && row[facilityIdx] ? row[facilityIdx].toString() : '미지정',
                        energyType: energyType,
                        startDate: startDate,
                        endDate: endDate,
                        usageAmount: usageAmount,
                        usageCost: usageCost
                    };

                    importedData.push(recordData);
                    totalImported++;

                } catch (error) {
                    console.error('행 ' + (i + 1) + ' 처리 오류:', error);
                }
            }
        }

        if (totalImported === 0) {
            alert('가져올 수 있는 데이터가 없습니다.');
            event.target.value = ''; // 파일 입력 초기화
            return;
        }

        console.log('총 ' + totalImported + '개 데이터 파싱 완료');

        // 사용자 확인
        const confirmed = confirm(totalImported + '개의 데이터를 가져왔습니다.\n데이터베이스에 저장하시겠습니까?');
        if (!confirmed) {
            event.target.value = ''; // 파일 입력 초기화
            return;
        }

        // 서버에 데이터 전송
        let successCount = 0;
        let errorCount = 0;

        for (const record of importedData) {
            try {
                const response = await fetch('/api/energy', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(record)
                });

                const result = await response.json();
                if (result.success) {
                    successCount++;
                } else {
                    errorCount++;
                    console.error('저장 실패:', record, result.message);
                }
            } catch (error) {
                errorCount++;
                console.error('저장 오류:', record, error);
            }
        }

        alert('데이터 저장 완료!\n성공: ' + successCount + '개\n실패: ' + errorCount + '개');

        // 데이터 조회 탭 새로고침
        if (successCount > 0 && typeof searchData === 'function') {
            await searchData();
        }

    } catch (error) {
        console.error('엑셀 업로드 오류:', error);
        alert('엑셀 파일 처리 중 오류가 발생했습니다: ' + error.message);
    }

    // 파일 입력 초기화
    event.target.value = '';
}
