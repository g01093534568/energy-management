// ==================== 관리자 모드 엑셀 기능 ====================

// 전역 변수
let currentFacilities = [];
let currentEnergyInfo = [];

// ==================== 시설 목록 관련 기능 ====================

// 시설 삭제 버튼 표시 업데이트
function updateFacilityDeleteButtonVisibility() {
    const checkboxes = document.querySelectorAll('.facility-checkbox');
    const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
    const deleteBtn = document.getElementById('deleteSelectedFacilitiesBtn');

    if (deleteBtn) {
        deleteBtn.style.display = checkedCount > 0 ? 'inline-block' : 'none';
    }
}

// 선택된 시설 삭제
async function deleteSelectedFacilities() {
    const checkboxes = document.querySelectorAll('.facility-checkbox:checked');

    if (checkboxes.length === 0) {
        alert('삭제할 항목을 선택해주세요.');
        return;
    }

    if (!confirm(`선택한 ${checkboxes.length}개 시설을 삭제하시겠습니까?`)) {
        return;
    }

    const indices = Array.from(checkboxes).map(cb => parseInt(cb.dataset.originalIndex));

    try {
        let successCount = 0;
        let errorCount = 0;

        for (const index of indices) {
            const response = await fetch(`/api/facilities/${index}`, {
                method: 'DELETE'
            });

            const result = await response.json();
            if (result.success) {
                successCount++;
            } else {
                errorCount++;
            }
        }

        alert(`삭제 완료!\n성공: ${successCount}개\n실패: ${errorCount}개`);

        // 전체 선택 체크박스 해제
        const selectAllCheckbox = document.getElementById('selectAllFacilitiesCheckbox');
        if (selectAllCheckbox) {
            selectAllCheckbox.checked = false;
        }

        // 목록 새로고침
        if (typeof loadFacilities === 'function') {
            loadFacilities();
        }
    } catch (error) {
        console.error('시설 삭제 오류:', error);
        alert('시설 삭제 중 오류가 발생했습니다.');
    }
}

// 시설 목록 엑셀 다운로드
async function downloadFacilitiesExcel() {
    try {
        if (typeof XLSX === 'undefined') {
            alert('엑셀 라이브러리가 로드되지 않았습니다.');
            return;
        }

        const response = await fetch('/api/facilities');
        const data = await response.json();

        if (!data.success || !data.facilities || data.facilities.length === 0) {
            alert('다운로드할 시설 데이터가 없습니다.');
            return;
        }

        const facilities = data.facilities;

        // 엑셀 데이터 생성
        const worksheetData = [
            ['시설명', 'id', 'pw', '역할', '상위시설명']
        ];

        facilities.forEach(facility => {
            worksheetData.push([
                facility.facilityName || '',
                facility.id || '',
                facility.password || '',
                facility.role || '시설담당자',
                facility.parentFacility || ''
            ]);
        });

        // 워크북 생성
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(worksheetData);

        // 열 너비 설정
        ws['!cols'] = [
            { wch: 25 },  // 시설명
            { wch: 15 },  // id
            { wch: 15 },  // pw
            { wch: 15 },  // 역할
            { wch: 25 }   // 상위시설명
        ];

        XLSX.utils.book_append_sheet(wb, ws, '시설목록');

        // 파일 다운로드
        const today = new Date();
        const filename = `시설목록_${today.getFullYear()}_${String(today.getMonth() + 1).padStart(2, '0')}_${String(today.getDate()).padStart(2, '0')}.xlsx`;
        XLSX.writeFile(wb, filename);

        alert('엑셀 파일이 다운로드되었습니다.');
    } catch (error) {
        console.error('엑셀 다운로드 오류:', error);
        alert('엑셀 다운로드 중 오류가 발생했습니다: ' + error.message);
    }
}

// 시설 목록 엑셀 업로드
async function uploadFacilitiesExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length < 2) {
            alert('유효한 데이터가 없습니다.');
            event.target.value = '';
            return;
        }

        const facilities = [];
        // 첫 번째 행은 헤더이므로 건너뜀
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (row && row.length >= 3 && row[1] && row[2]) {
                facilities.push({
                    facilityName: String(row[0] || '').trim(),
                    id: String(row[1]).trim(),
                    password: String(row[2]).trim(),
                    role: row[3] ? String(row[3]).trim() : '시설담당자',
                    parentFacility: row[4] ? String(row[4]).trim() : ''
                });
            }
        }

        if (facilities.length === 0) {
            alert('가져올 수 있는 시설 데이터가 없습니다.');
            event.target.value = '';
            return;
        }

        const confirmed = confirm(`${facilities.length}개의 시설을 추가하시겠습니까?`);
        if (!confirmed) {
            event.target.value = '';
            return;
        }

        let successCount = 0;
        let errorCount = 0;

        for (const facility of facilities) {
            try {
                const response = await fetch('/api/facilities', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(facility)
                });

                const result = await response.json();
                if (result.success) {
                    successCount++;
                } else {
                    errorCount++;
                }
            } catch (error) {
                errorCount++;
            }
        }

        alert(`시설 추가 완료!\n성공: ${successCount}개\n실패: ${errorCount}개`);

        if (successCount > 0 && typeof loadFacilities === 'function') {
            loadFacilities();
        }
    } catch (error) {
        console.error('엑셀 업로드 오류:', error);
        alert('엑셀 파일 처리 중 오류가 발생했습니다: ' + error.message);
    }

    event.target.value = '';
}

// ==================== 에너지 정보 관리 관련 기능 ====================

// 에너지 정보 삭제 버튼 표시 업데이트
function updateEnergyInfoDeleteButtonVisibility() {
    const checkboxes = document.querySelectorAll('.energy-info-checkbox');
    const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
    const deleteBtn = document.getElementById('deleteSelectedEnergyInfoBtn');

    if (deleteBtn) {
        deleteBtn.style.display = checkedCount > 0 ? 'inline-block' : 'none';
    }
}

// 선택된 에너지 정보 삭제
async function deleteSelectedEnergyInfo() {
    const checkboxes = document.querySelectorAll('.energy-info-checkbox:checked');

    if (checkboxes.length === 0) {
        alert('삭제할 항목을 선택해주세요.');
        return;
    }

    if (!confirm(`선택한 ${checkboxes.length}개 에너지 정보를 삭제하시겠습니까?`)) {
        return;
    }

    const indices = Array.from(checkboxes).map(cb => parseInt(cb.dataset.index));

    try {
        let successCount = 0;
        let errorCount = 0;

        for (const index of indices) {
            const response = await fetch(`/api/energy-info/${index}`, {
                method: 'DELETE'
            });

            const result = await response.json();
            if (result.success) {
                successCount++;
            } else {
                errorCount++;
            }
        }

        alert(`삭제 완료!\n성공: ${successCount}개\n실패: ${errorCount}개`);

        // 전체 선택 체크박스 해제
        const selectAllCheckbox = document.getElementById('selectAllEnergyInfoCheckbox');
        if (selectAllCheckbox) {
            selectAllCheckbox.checked = false;
        }

        // 목록 새로고침
        if (typeof loadEnergyInfo === 'function') {
            loadEnergyInfo();
        }
    } catch (error) {
        console.error('에너지 정보 삭제 오류:', error);
        alert('에너지 정보 삭제 중 오류가 발생했습니다.');
    }
}

// 에너지 정보 엑셀 다운로드
async function downloadEnergyInfoExcel() {
    try {
        if (typeof XLSX === 'undefined') {
            alert('엑셀 라이브러리가 로드되지 않았습니다.');
            return;
        }

        const response = await fetch('/api/energy-info');
        const data = await response.json();

        console.log('에너지 정보 응답:', data);

        // 서버에서 'infos'라는 키로 반환함
        if (!data.success || !data.infos || data.infos.length === 0) {
            alert('다운로드할 에너지 정보가 없습니다.');
            return;
        }

        const energyInfo = data.infos;

        // 엑셀 데이터 생성
        const worksheetData = [
            ['시설명', '에너지 종류', '고객번호(명세서번호)', '금융기관', '계좌번호']
        ];

        energyInfo.forEach(info => {
            worksheetData.push([
                info.facilityName || '',
                info.energyType || '',
                info.customerNumber || '',
                info.bankName || '',
                info.accountNumber || ''
            ]);
        });

        // 워크북 생성
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(worksheetData);

        // 열 너비 설정
        ws['!cols'] = [
            { wch: 25 },  // 시설명
            { wch: 15 },  // 에너지 종류
            { wch: 25 },  // 고객번호
            { wch: 20 },  // 금융기관
            { wch: 25 }   // 계좌번호
        ];

        XLSX.utils.book_append_sheet(wb, ws, '에너지정보');

        // 파일 다운로드
        const today = new Date();
        const filename = `에너지정보_${today.getFullYear()}_${String(today.getMonth() + 1).padStart(2, '0')}_${String(today.getDate()).padStart(2, '0')}.xlsx`;
        XLSX.writeFile(wb, filename);

        alert('엑셀 파일이 다운로드되었습니다.');
    } catch (error) {
        console.error('엑셀 다운로드 오류:', error);
        alert('엑셀 다운로드 중 오류가 발생했습니다: ' + error.message);
    }
}

// 에너지 정보 엑셀 업로드
async function uploadEnergyInfoExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length < 2) {
            alert('유효한 데이터가 없습니다.');
            event.target.value = '';
            return;
        }

        const energyInfoList = [];
        // 첫 번째 행은 헤더이므로 건너뜀
        for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (row && row.length >= 5 && row[0] && row[1]) {
                energyInfoList.push({
                    facilityName: String(row[0] || '').trim(),
                    energyType: String(row[1] || '').trim(),
                    customerNumber: String(row[2] || '').trim(),
                    bankName: String(row[3] || '').trim(),
                    accountNumber: String(row[4] || '').trim()
                });
            }
        }

        if (energyInfoList.length === 0) {
            alert('가져올 수 있는 에너지 정보가 없습니다.');
            event.target.value = '';
            return;
        }

        const confirmed = confirm(`${energyInfoList.length}개의 에너지 정보를 추가하시겠습니까?`);
        if (!confirmed) {
            event.target.value = '';
            return;
        }

        let successCount = 0;
        let errorCount = 0;

        for (const info of energyInfoList) {
            try {
                const response = await fetch('/api/energy-info', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(info)
                });

                const result = await response.json();
                if (result.success) {
                    successCount++;
                } else {
                    errorCount++;
                }
            } catch (error) {
                errorCount++;
            }
        }

        alert(`에너지 정보 추가 완료!\n성공: ${successCount}개\n실패: ${errorCount}개`);

        if (successCount > 0 && typeof loadEnergyInfo === 'function') {
            loadEnergyInfo();
        }
    } catch (error) {
        console.error('엑셀 업로드 오류:', error);
        alert('엑셀 파일 처리 중 오류가 발생했습니다: ' + error.message);
    }

    event.target.value = '';
}

// ==================== 이벤트 리스너 등록 ====================

document.addEventListener('DOMContentLoaded', () => {
    // 시설 목록 관련 이벤트
    const deleteSelectedFacilitiesBtn = document.getElementById('deleteSelectedFacilitiesBtn');
    const selectAllFacilitiesCheckbox = document.getElementById('selectAllFacilitiesCheckbox');
    const downloadFacilityExcelBtn = document.getElementById('downloadFacilityExcelBtn');
    const uploadFacilityExcelBtn = document.getElementById('uploadFacilityExcelBtn');
    const facilityExcelFileInput = document.getElementById('facilityExcelFileInput');

    if (deleteSelectedFacilitiesBtn) {
        deleteSelectedFacilitiesBtn.addEventListener('click', deleteSelectedFacilities);
    }

    if (selectAllFacilitiesCheckbox) {
        selectAllFacilitiesCheckbox.addEventListener('change', function() {
            const checkboxes = document.querySelectorAll('.facility-checkbox');
            checkboxes.forEach(cb => {
                cb.checked = this.checked;
            });
            updateFacilityDeleteButtonVisibility();
        });
    }

    if (downloadFacilityExcelBtn) {
        downloadFacilityExcelBtn.addEventListener('click', downloadFacilitiesExcel);
    }

    if (uploadFacilityExcelBtn) {
        uploadFacilityExcelBtn.addEventListener('click', () => {
            facilityExcelFileInput.click();
        });
    }

    if (facilityExcelFileInput) {
        facilityExcelFileInput.addEventListener('change', uploadFacilitiesExcel);
    }

    // 에너지 정보 관련 이벤트
    const deleteSelectedEnergyInfoBtn = document.getElementById('deleteSelectedEnergyInfoBtn');
    const selectAllEnergyInfoCheckbox = document.getElementById('selectAllEnergyInfoCheckbox');
    const downloadEnergyInfoExcelBtn = document.getElementById('downloadEnergyInfoExcelBtn');
    const uploadEnergyInfoExcelBtn = document.getElementById('uploadEnergyInfoExcelBtn');
    const energyInfoExcelFileInput = document.getElementById('energyInfoExcelFileInput');

    if (deleteSelectedEnergyInfoBtn) {
        deleteSelectedEnergyInfoBtn.addEventListener('click', deleteSelectedEnergyInfo);
    }

    if (selectAllEnergyInfoCheckbox) {
        selectAllEnergyInfoCheckbox.addEventListener('change', function() {
            const checkboxes = document.querySelectorAll('.energy-info-checkbox');
            checkboxes.forEach(cb => {
                cb.checked = this.checked;
            });
            updateEnergyInfoDeleteButtonVisibility();
        });
    }

    if (downloadEnergyInfoExcelBtn) {
        downloadEnergyInfoExcelBtn.addEventListener('click', downloadEnergyInfoExcel);
    }

    if (uploadEnergyInfoExcelBtn) {
        uploadEnergyInfoExcelBtn.addEventListener('click', () => {
            energyInfoExcelFileInput.click();
        });
    }

    if (energyInfoExcelFileInput) {
        energyInfoExcelFileInput.addEventListener('change', uploadEnergyInfoExcel);
    }
});
