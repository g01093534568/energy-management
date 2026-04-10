let currentEditIndex = null;
let currentEditEnergyInfoIndex = null;

// 시설 목록 페이지네이션 변수
let facilityCurrentPage = 1;
const facilityItemsPerPage = 20;
let allFacilities = [];

// 에너지 정보 페이지네이션 변수
let energyInfoCurrentPage = 1;
const energyInfoItemsPerPage = 20;
let allEnergyInfos = [];

function showLogin() {
    document.getElementById('loginView').style.display = 'flex';
    document.getElementById('dashboardView').style.display = 'none';
}

function showDashboard() {
    document.getElementById('loginView').style.display = 'none';
    document.getElementById('dashboardView').style.display = 'block';
}

async function checkAuth() {
    try {
        const response = await fetch('/api/check-auth');
        const data = await response.json();

        if (data.authenticated) {
            const userName = document.getElementById('userName');
            const headerTitle = document.querySelector('.header h1');
            const role = data.user.role || '시설담당자';

            // 세션 스토리지에 사용자 정보 저장
            sessionStorage.setItem('currentUser', JSON.stringify(data.user));

            // 상단 타이틀에 시설명 추가
            if (headerTitle) {
                let titleSuffix = '';
                if (role === '관리자') {
                    titleSuffix = ' (관리자)';
                } else if (role === '시설관리자') {
                    titleSuffix = ' (시설관리자)';
                } else if (data.user.facilityName) {
                    titleSuffix = ` (${data.user.facilityName})`;
                }
                headerTitle.textContent = `에너지 관리 시스템${titleSuffix}`;
            }

            // 사용자 정보 표시
            if (data.user.facilityName) {
                userName.textContent = `${data.user.facilityName} (${data.user.id}) - ${role}`;
            } else {
                userName.textContent = `${data.user.id} - ${role}`;
            }

            // 역할에 따라 탭 표시 제어
            updateTabVisibility(role);
            showDashboard();

            // 에너지 입력 탭의 시설명 옵션 로드
            loadFacilityOptionsForEnergyInput();
        } else {
            showLogin();
        }
    } catch (error) {
        console.error('인증 확인 오류:', error);
        showLogin();
    }
}

function updateTabVisibility(role) {
    const adminTab = document.querySelector('[data-tab="admin"]');

    // 관리자, 시설관리자는 관리자 모드 탭 표시
    if (role === '관리자' || role === '시설관리자') {
        if (adminTab) adminTab.style.display = 'block';
    } else {
        // 시설담당자는 관리자 모드 탭 숨김
        if (adminTab) adminTab.style.display = 'none';
    }
}

document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('loginForm').addEventListener('submit', async (e) => {
        e.preventDefault();

        const facilityName = document.getElementById('facilityName').value.trim();
        const username = document.getElementById('username').value.trim();
        const password = document.getElementById('password').value.trim();
        const errorMessage = document.getElementById('errorMessage');

        console.log('로그인 시도:', { facilityName, username, password });

        try {
            const response = await fetch('/api/login', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ facilityName, username, password })
            });

            console.log('응답 상태:', response.status);

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const data = await response.json();
            console.log('응답 데이터:', data);

            if (data.success) {
                errorMessage.textContent = '';
                checkAuth();
            } else {
                errorMessage.textContent = data.message || '로그인에 실패했습니다.';
            }
        } catch (error) {
            errorMessage.textContent = '서버 연결 오류가 발생했습니다. 서버가 실행 중인지 확인하세요.';
            console.error('로그인 오류:', error);
        }
    });

    document.getElementById('logoutBtn').addEventListener('click', async () => {
        try {
            await fetch('/api/logout', { method: 'POST' });
            showLogin();
            document.getElementById('loginForm').reset();
            document.getElementById('errorMessage').textContent = '';
        } catch (error) {
            console.error('로그아웃 오류:', error);
        }
    });

    const tabs = document.querySelectorAll('.tab');
    const tabContents = document.querySelectorAll('.tab-content');

    tabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const tabName = tab.dataset.tab;

            tabs.forEach(t => t.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));

            tab.classList.add('active');
            document.getElementById(tabName).classList.add('active');

            if (tabName === 'admin') {
                loadFacilities();
                loadEnergyInfos();
            } else if (tabName === 'energy') {
                loadFacilityOptionsForEnergyInput();
                loadEnergyRecords();
            } else if (tabName === 'dataView') {
                loadDataViewTab();
            }
        });
    });

    const modal = document.getElementById('facilityModal');
    const closeBtn = document.querySelector('.close');
    const cancelBtn = document.querySelector('.cancel-btn');
    const addFacilityBtn = document.getElementById('addFacilityBtn');
    const facilityForm = document.getElementById('facilityForm');

    addFacilityBtn.addEventListener('click', () => {
        openModal(false);
    });

    closeBtn.addEventListener('click', closeModal);
    cancelBtn.addEventListener('click', closeModal);

    // 선택 항목 수정 버튼 이벤트
    const editSelectedFacilityBtn = document.getElementById('editSelectedFacilityBtn');
    if (editSelectedFacilityBtn) {
        editSelectedFacilityBtn.addEventListener('click', window.editSelectedFacility);
    }

    // 시설 목록 전체 선택 체크박스 이벤트
    const selectAllFacilitiesCheckbox = document.getElementById('selectAllFacilitiesCheckbox');
    if (selectAllFacilitiesCheckbox) {
        selectAllFacilitiesCheckbox.addEventListener('change', function() {
            const checkboxes = document.querySelectorAll('.facility-checkbox');
            checkboxes.forEach(cb => cb.checked = this.checked);
            updateFacilityDeleteButtonVisibility();
        });
    }

    // 에너지 정보 전체 선택 체크박스 이벤트
    const selectAllEnergyInfoCheckbox = document.getElementById('selectAllEnergyInfoCheckbox');
    if (selectAllEnergyInfoCheckbox) {
        selectAllEnergyInfoCheckbox.addEventListener('change', function() {
            const checkboxes = document.querySelectorAll('.energy-info-checkbox');
            checkboxes.forEach(cb => cb.checked = this.checked);
            updateEnergyInfoDeleteButtonVisibility();
        });
    }

    window.addEventListener('click', (e) => {
        if (e.target === modal) {
            closeModal();
        }
    });

    facilityForm.addEventListener('submit', async (e) => {
        e.preventDefault();

        const facilityName = document.getElementById('editFacilityName').value;
        const id = document.getElementById('editId').value;
        const password = document.getElementById('editPassword').value;
        const role = document.getElementById('editRole').value;

        const facilityData = { facilityName, id, password, role };

        try {
            let response;
            if (currentEditIndex !== null) {
                response = await fetch(`/api/facilities/${currentEditIndex}`, {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(facilityData)
                });
            } else {
                response = await fetch('/api/facilities', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(facilityData)
                });
            }

            const data = await response.json();

            if (data.success) {
                alert(data.message);
                closeModal();
                loadFacilities();
            } else {
                alert(data.message);
            }
        } catch (error) {
            console.error('시설 저장 오류:', error);
            alert('시설 저장 중 오류가 발생했습니다.');
        }
    });

    const energyTypeSelect = document.getElementById('energyType');
    const usageAmountInput = document.getElementById('usageAmount');
    const usageCostInput = document.getElementById('usageCost');
    const usageUnitLabel = document.getElementById('usageUnit');

    if (energyTypeSelect && usageAmountInput && usageCostInput && usageUnitLabel) {
        energyTypeSelect.addEventListener('change', () => {
            const energyType = energyTypeSelect.value;
            const usageAmountLabel = document.querySelector('label[for="usageAmount"]');

            if (energyType === '전기') {
                usageUnitLabel.textContent = 'kWh';
                usageAmountInput.required = true;
                if (usageAmountLabel) usageAmountLabel.textContent = '사용량 *';
            } else if (energyType === '상하수도' || energyType === '도시가스') {
                usageUnitLabel.textContent = 'm³';
                usageAmountInput.required = true;
                if (usageAmountLabel) usageAmountLabel.textContent = '사용량 *';
            } else if (energyType === '통신') {
                usageUnitLabel.textContent = '건';
                usageAmountInput.required = false;
                if (usageAmountLabel) usageAmountLabel.textContent = '사용량 (선택사항)';
            } else {
                usageUnitLabel.textContent = 'kWh';
                usageAmountInput.required = true;
                if (usageAmountLabel) usageAmountLabel.textContent = '사용량 *';
            }
        });

        function formatNumberWithComma(value) {
            const numbers = value.replace(/[^\d]/g, '');
            return numbers.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
        }

        usageAmountInput.addEventListener('input', (e) => {
            const cursorPosition = e.target.selectionStart;
            const oldLength = e.target.value.length;
            const formatted = formatNumberWithComma(e.target.value);
            e.target.value = formatted;
            const newLength = formatted.length;
            const newCursorPosition = cursorPosition + (newLength - oldLength);
            e.target.setSelectionRange(newCursorPosition, newCursorPosition);
        });

        usageCostInput.addEventListener('input', (e) => {
            const cursorPosition = e.target.selectionStart;
            const oldLength = e.target.value.length;
            const formatted = formatNumberWithComma(e.target.value);
            e.target.value = formatted;
            const newLength = formatted.length;
            const newCursorPosition = cursorPosition + (newLength - oldLength);
            e.target.setSelectionRange(newCursorPosition, newCursorPosition);
        });
    }

    const energyForm = document.getElementById('energyForm');
    if (energyForm) {
        energyForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const energyType = document.getElementById('energyType').value;
            const billingMonth = document.getElementById('billingMonth').value.trim();
            let startDateValue = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            const usageAmount = document.getElementById('usageAmount').value.replace(/,/g, '').trim();
            const usageCost = document.getElementById('usageCost').value.replace(/,/g, '').trim();

            // startDate에서 실제 시작일 추출 (형식: "2024-01-01 ~ 2024-01-31" 또는 "2024-01-01")
            let startDate = startDateValue;
            if (startDateValue.includes('~')) {
                startDate = startDateValue.split('~')[0].trim();
            }

            // 시설명 가져오기: 드롭다운이 표시되면 드롭다운 값 사용, 아니면 현재 사용자 시설명 사용
            let facilityName;
            const facilityNameSelect = document.getElementById('facilityNameSelect');
            const facilitySelectRow = document.getElementById('facilitySelectRow');

            if (facilitySelectRow.style.display !== 'none' && facilityNameSelect) {
                facilityName = facilityNameSelect.value;
            } else {
                // 시설담당자는 자동으로 본인 시설명 사용
                const userStr = sessionStorage.getItem('currentUser');
                if (userStr) {
                    const currentUser = JSON.parse(userStr);
                    facilityName = currentUser.facilityName;
                }
            }

            // 필드 검증 - 통신의 경우 사용량은 선택사항
            if (energyType === '통신') {
                if (!facilityName || !energyType || !startDate || !endDate || !usageCost) {
                    alert('모든 필드를 입력해주세요. (통신의 경우 사용량은 선택사항입니다)');
                    return;
                }
            } else {
                if (!facilityName || !energyType || !startDate || !endDate || !usageAmount || !usageCost) {
                    alert('모든 필드를 입력해주세요.');
                    return;
                }
            }

            // 숫자 유효성 검사
            if (energyType === '통신') {
                // 통신의 경우 사용금액만 검증
                if (isNaN(parseFloat(usageCost))) {
                    alert('사용 금액은 숫자여야 합니다.');
                    return;
                }
            } else {
                // 다른 에너지 종류는 사용량과 사용금액 모두 검증
                if (isNaN(parseFloat(usageAmount)) || isNaN(parseFloat(usageCost))) {
                    alert('사용량과 사용 금액은 숫자여야 합니다.');
                    return;
                }
            }

            // 날짜 유효성 검사
            if (new Date(startDate) > new Date(endDate)) {
                alert('종료일은 시작일 이후여야 합니다.');
                return;
            }

            const energyData = {
                facilityName,
                energyType,
                billingMonth,
                startDate,
                endDate,
                usageAmount: energyType === '통신' ? (usageAmount ? parseFloat(usageAmount) : 0) : parseFloat(usageAmount),
                usageCost: parseFloat(usageCost)
            };

            try {
                const response = await fetch('/api/energy', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(energyData)
                });

                const data = await response.json();

                if (data.success) {
                    alert('에너지 사용량이 저장되었습니다.');
                    energyForm.reset();

                    // OCR 미리보기 제거
                    removeOCRPreview();

                    loadEnergyRecords();
                } else {
                    alert(data.message || '저장 중 오류가 발생했습니다.');
                }
            } catch (error) {
                console.error('에너지 데이터 저장 오류:', error);
                alert('에너지 데이터 저장 중 오류가 발생했습니다.');
            }
        });
    }

    // 날짜를 로컬 시간으로 포맷하는 함수 (타임존 문제 해결)
    const formatLocalDate = (date) => {
        const year = date.getFullYear();
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const day = String(date.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };

    // Flatpickr 날짜 선택기 초기화
    if (typeof flatpickr !== 'undefined') {
        // 에너지 사용량 입력 - 사용 기간
        const startDateInput = document.getElementById('startDate');
        const endDateInput = document.getElementById('endDate');

        if (startDateInput && endDateInput) {
            flatpickr(startDateInput, {
                mode: "range",
                locale: "ko",
                dateFormat: "Y-m-d",
                onChange: function(selectedDates, dateStr) {
                    if (selectedDates.length === 2) {
                        const start = formatLocalDate(selectedDates[0]);
                        const end = formatLocalDate(selectedDates[1]);
                        startDateInput.value = `${start} ~ ${end}`;
                        endDateInput.value = end;
                    } else if (selectedDates.length === 1) {
                        const start = formatLocalDate(selectedDates[0]);
                        startDateInput.value = start;
                        endDateInput.value = start;
                    }
                },
                onClose: function(selectedDates) {
                    if (selectedDates.length === 2) {
                        const start = formatLocalDate(selectedDates[0]);
                        const end = formatLocalDate(selectedDates[1]);
                        startDateInput.value = `${start} ~ ${end}`;
                        endDateInput.value = end;
                    }
                }
            });
        }

        // 에너지 사용량 입력 - 월 선택
        const billingMonthInput = document.getElementById('billingMonth');

        if (billingMonthInput && typeof monthSelectPlugin !== 'undefined') {
            flatpickr(billingMonthInput, {
                locale: "ko",
                plugins: [
                    new monthSelectPlugin({
                        shorthand: true,
                        dateFormat: "Y-m",
                        altFormat: "Y년 m월"
                    })
                ]
            });
        }

        // 데이터 조회 - 조회 기간 (월분 선택은 HTML5 month input 사용)
    }

    checkAuth();
});

async function loadFacilities() {
    try {
        const response = await fetch('/api/facilities');
        const data = await response.json();

        if (data.success) {
            // 시설명 가나다순으로 정렬
            const sortedFacilities = data.facilities.sort((a, b) => {
                return a.facilityName.localeCompare(b.facilityName, 'ko-KR');
            });

            renderFacilities(sortedFacilities);

            // 시설 추가 버튼 표시 제어
            const addFacilityBtn = document.getElementById('addFacilityBtn');
            const currentUser = JSON.parse(sessionStorage.getItem('currentUser'));

            if (addFacilityBtn && currentUser) {
                // 관리자, 시설관리자는 시설 추가 가능
                if (currentUser.role === '관리자' || currentUser.role === '시설관리자') {
                    addFacilityBtn.style.display = 'inline-block';
                } else {
                    addFacilityBtn.style.display = 'none';
                }
            }
        }
    } catch (error) {
        console.error('시설 목록 로드 오류:', error);
        alert('시설 목록을 불러오는 중 오류가 발생했습니다.');
    }
}

function renderFacilities(facilities, page = 1) {
    const tbody = document.getElementById('facilityTableBody');
    tbody.innerHTML = '';

    // 전역 변수에 저장
    allFacilities = facilities;
    facilityCurrentPage = page;

    if (!facilities || facilities.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px; color: #999;">등록된 시설이 없습니다.</td></tr>';
        renderFacilityPagination(0, 1);
        return;
    }

    // 페이지네이션 계산
    const totalPages = Math.ceil(facilities.length / facilityItemsPerPage);
    const startIndex = (page - 1) * facilityItemsPerPage;
    const endIndex = Math.min(startIndex + facilityItemsPerPage, facilities.length);
    const pageRecords = facilities.slice(startIndex, endIndex);

    pageRecords.forEach((facility, idx) => {
        const index = startIndex + idx; // 전체 배열에서의 실제 인덱스
        const row = document.createElement('tr');
        // originalIndex가 있으면 사용, 없으면 일반 index 사용 (하위 호환성)
        const actualIndex = facility.originalIndex !== undefined ? facility.originalIndex : index;
        row.innerHTML = `
            <td><input type="checkbox" class="facility-checkbox" data-index="${index}" data-original-index="${actualIndex}"></td>
            <td>${facility.facilityName}</td>
            <td>${facility.id}</td>
            <td>${facility.password}</td>
            <td>${facility.role || '시설담당자'}</td>
            <td>${facility.parentFacility || '-'}</td>
            <td>
                <button class="edit-btn" onclick="editFacility(${index}, ${actualIndex})">수정</button>
                <button class="delete-btn" onclick="deleteFacility(${index}, ${actualIndex})">삭제</button>
            </td>
        `;
        tbody.appendChild(row);
    });

    // 페이지네이션 렌더링
    renderFacilityPagination(facilities.length, totalPages);

    // 체크박스 변경 이벤트 리스너
    updateFacilityDeleteButtonVisibility();
    document.querySelectorAll('.facility-checkbox').forEach(cb => {
        cb.addEventListener('change', updateFacilityDeleteButtonVisibility);
    });
}

// 시설 목록 페이지네이션 렌더링
function renderFacilityPagination(totalItems, totalPages) {
    const container = document.getElementById('facilityPaginationContainer');
    if (!container) return;

    if (totalItems === 0) {
        container.innerHTML = '';
        return;
    }

    let html = `<span style="margin-right: 15px; color: #666;">총 ${totalItems}건</span>`;

    html += `<button class="pagination-btn" onclick="goToFacilityPage(1)" ${facilityCurrentPage === 1 ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${facilityCurrentPage === 1 ? '#f5f5f5' : '#fff'}; border-radius: 4px;">≪</button>`;
    html += `<button class="pagination-btn" onclick="goToFacilityPage(${facilityCurrentPage - 1})" ${facilityCurrentPage === 1 ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${facilityCurrentPage === 1 ? '#f5f5f5' : '#fff'}; border-radius: 4px;">＜</button>`;

    const maxVisiblePages = 5;
    let startPage = Math.max(1, facilityCurrentPage - Math.floor(maxVisiblePages / 2));
    let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);

    if (endPage - startPage + 1 < maxVisiblePages) {
        startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }

    if (startPage > 1) {
        html += `<span style="margin: 0 5px;">...</span>`;
    }

    for (let i = startPage; i <= endPage; i++) {
        const isActive = i === facilityCurrentPage;
        html += `<button class="pagination-btn" onclick="goToFacilityPage(${i})" style="padding: 5px 12px; margin: 2px; cursor: pointer; border: 1px solid ${isActive ? '#1976d2' : '#ddd'}; background: ${isActive ? '#1976d2' : '#fff'}; color: ${isActive ? '#fff' : '#333'}; border-radius: 4px; font-weight: ${isActive ? 'bold' : 'normal'};">${i}</button>`;
    }

    if (endPage < totalPages) {
        html += `<span style="margin: 0 5px;">...</span>`;
    }

    html += `<button class="pagination-btn" onclick="goToFacilityPage(${facilityCurrentPage + 1})" ${facilityCurrentPage === totalPages ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${facilityCurrentPage === totalPages ? '#f5f5f5' : '#fff'}; border-radius: 4px;">＞</button>`;
    html += `<button class="pagination-btn" onclick="goToFacilityPage(${totalPages})" ${facilityCurrentPage === totalPages ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${facilityCurrentPage === totalPages ? '#f5f5f5' : '#fff'}; border-radius: 4px;">≫</button>`;

    container.innerHTML = html;
}

// 시설 목록 페이지 이동
window.goToFacilityPage = function(page) {
    if (page < 1 || allFacilities.length === 0) return;
    const totalPages = Math.ceil(allFacilities.length / facilityItemsPerPage);
    if (page > totalPages) page = totalPages;
    renderFacilities(allFacilities, page);
};

function openModal(isEdit = false, index = null) {
    currentEditIndex = index;
    const modalTitle = document.getElementById('modalTitle');
    modalTitle.textContent = isEdit ? '시설 수정' : '시설 추가';

    if (!isEdit) {
        document.getElementById('editFacilityName').value = '';
        document.getElementById('editId').value = '';
        document.getElementById('editPassword').value = '';
        document.getElementById('editRole').value = '';
    }

    // 역할 드롭다운을 현재 사용자 권한에 따라 설정
    updateRoleDropdown();

    const modal = document.getElementById('facilityModal');
    modal.style.display = 'block';
}

function updateRoleDropdown() {
    const roleSelect = document.getElementById('editRole');
    const currentUser = JSON.parse(sessionStorage.getItem('currentUser'));

    if (!currentUser) return;

    // 기존 옵션 제거 (첫 번째 "선택하세요" 옵션 제외)
    while (roleSelect.options.length > 1) {
        roleSelect.remove(1);
    }

    // 관리자는 모든 역할 선택 가능
    if (currentUser.role === '관리자') {
        const option1 = document.createElement('option');
        option1.value = '관리자';
        option1.textContent = '관리자';
        roleSelect.appendChild(option1);

        const option2 = document.createElement('option');
        option2.value = '시설관리자';
        option2.textContent = '시설관리자';
        roleSelect.appendChild(option2);

        const option3 = document.createElement('option');
        option3.value = '시설담당자';
        option3.textContent = '시설담당자';
        roleSelect.appendChild(option3);
    }
    // 시설관리자는 시설담당자만 선택 가능
    else if (currentUser.role === '시설관리자') {
        const option = document.createElement('option');
        option.value = '시설담당자';
        option.textContent = '시설담당자';
        roleSelect.appendChild(option);
    }
}

function closeModal() {
    const modal = document.getElementById('facilityModal');
    modal.style.display = 'none';
    currentEditIndex = null;
}

// 시설 목록 체크박스 선택에 따라 수정/삭제 버튼 표시
function updateFacilityDeleteButtonVisibility() {
    const checkboxes = document.querySelectorAll('.facility-checkbox:checked');
    const deleteBtn = document.getElementById('deleteSelectedFacilitiesBtn');
    const editBtn = document.getElementById('editSelectedFacilityBtn');

    if (checkboxes.length > 0) {
        if (deleteBtn) deleteBtn.style.display = 'inline-block';
        // 수정 버튼은 1개만 선택했을 때만 표시
        if (editBtn) {
            editBtn.style.display = checkboxes.length === 1 ? 'inline-block' : 'none';
        }
    } else {
        if (deleteBtn) deleteBtn.style.display = 'none';
        if (editBtn) editBtn.style.display = 'none';
    }
}

// 선택된 시설 수정
window.editSelectedFacility = async function() {
    const checkbox = document.querySelector('.facility-checkbox:checked');
    if (!checkbox) {
        alert('수정할 시설을 선택해주세요.');
        return;
    }

    const displayIndex = parseInt(checkbox.dataset.index);
    const actualIndex = parseInt(checkbox.dataset.originalIndex);

    try {
        const response = await fetch('/api/facilities');
        const data = await response.json();

        if (data.success && data.facilities[displayIndex]) {
            const facility = data.facilities[displayIndex];
            document.getElementById('editFacilityName').value = facility.facilityName;
            document.getElementById('editId').value = facility.id;
            document.getElementById('editPassword').value = facility.password;
            document.getElementById('editRole').value = facility.role || '시설담당자';
            document.getElementById('editParentFacility').value = facility.parentFacility || '';
            openModal(true, actualIndex);
        }
    } catch (error) {
        console.error('시설 정보 로드 오류:', error);
        alert('시설 정보를 불러오는 중 오류가 발생했습니다.');
    }
};

window.editFacility = async function(displayIndex, actualIndex) {
    try {
        const response = await fetch('/api/facilities');
        const data = await response.json();

        if (data.success && data.facilities[displayIndex]) {
            const facility = data.facilities[displayIndex];
            document.getElementById('editFacilityName').value = facility.facilityName;
            document.getElementById('editId').value = facility.id;
            document.getElementById('editPassword').value = facility.password;
            document.getElementById('editRole').value = facility.role || '시설담당자';
            // actualIndex를 사용하여 서버에 전송
            openModal(true, actualIndex);
        }
    } catch (error) {
        console.error('시설 정보 로드 오류:', error);
        alert('시설 정보를 불러오는 중 오류가 발생했습니다.');
    }
};

window.deleteFacility = async function(displayIndex, actualIndex) {
    if (!confirm('정말 이 시설을 삭제하시겠습니까?')) {
        return;
    }

    try {
        // actualIndex를 사용하여 서버에 요청
        const response = await fetch(`/api/facilities/${actualIndex}`, {
            method: 'DELETE'
        });

        const data = await response.json();

        if (data.success) {
            alert(data.message);
            loadFacilities();
        } else {
            alert(data.message);
        }
    } catch (error) {
        console.error('시설 삭제 오류:', error);
        alert('시설 삭제 중 오류가 발생했습니다.');
    }
};

async function loadEnergyRecords() {
    try {
        const response = await fetch('/api/energy');
        const data = await response.json();

        if (data.success) {
            renderEnergyRecords(data.records);
        }
    } catch (error) {
        console.error('에너지 데이터 로드 오류:', error);
    }
}

async function loadFacilityOptionsForEnergyInput() {
    try {
        const userStr = sessionStorage.getItem('currentUser');
        console.log('=== 시설 옵션 로드 ===');
        console.log('userStr:', userStr);

        if (!userStr) {
            console.log('사용자 정보 없음');
            return;
        }

        const currentUser = JSON.parse(userStr);
        const role = currentUser.role;
        console.log('사용자 역할:', role);
        console.log('사용자 시설:', currentUser.facilityName);

        // 관리자 또는 시설관리자만 시설 선택 가능
        const facilitySelectRow = document.getElementById('facilitySelectRow');
        const facilityNameSelect = document.getElementById('facilityNameSelect');

        if (role === '관리자' || role === '시설관리자') {
            // 시설 목록 로드
            console.log('시설 목록 API 호출 중...');
            const response = await fetch('/api/facilities');
            console.log('API 응답 상태:', response.status);
            const data = await response.json();
            console.log('API 응답 데이터:', data);

            if (data.success && data.facilities) {
                // 드롭다운 초기화 및 옵션 추가
                facilityNameSelect.innerHTML = '<option value="">선택하세요</option>';

                data.facilities.forEach(facility => {
                    const option = document.createElement('option');
                    option.value = facility.facilityName;
                    option.textContent = facility.facilityName;
                    facilityNameSelect.appendChild(option);
                });

                console.log('시설 옵션 추가 완료:', data.facilities.length, '개');

                // 시설 선택 행 표시
                facilitySelectRow.style.display = 'flex';
                facilityNameSelect.required = true;
            } else {
                console.log('시설 데이터 없음 또는 실패:', data.message);
            }
        } else {
            // 시설담당자는 시설 선택에 본인 시설만 표시
            console.log('시설담당자: 본인 시설만 표시');
            facilityNameSelect.innerHTML = `<option value="${currentUser.facilityName}">${currentUser.facilityName}</option>`;
            facilitySelectRow.style.display = 'flex';
            facilityNameSelect.required = true;
            facilityNameSelect.disabled = true; // 변경 불가
        }
    } catch (error) {
        console.error('시설 옵션 로드 오류:', error);
    }
}

function renderEnergyRecords(records) {
    const tbody = document.getElementById('energyTableBody');
    tbody.innerHTML = '';

    records.forEach((record, index) => {
        const row = document.createElement('tr');
        const formattedAmount = parseFloat(record.usageAmount).toLocaleString('ko-KR');
        const formattedCost = parseFloat(record.usageCost).toLocaleString('ko-KR');

        let unit = 'kWh';
        if (record.energyType === '상하수도' || record.energyType === '도시가스') {
            unit = 'm³';
        } else if (record.energyType === '통신') {
            unit = '건';
        }

        // 날짜 포맷팅
        const startDate = record.startDate || record.usageDate;
        const endDate = record.endDate || record.usageDate;
        const dateRange = startDate === endDate ? startDate : `${startDate} ~ ${endDate}`;

        row.innerHTML = `
            <td>${formatBillingMonth(record.billingMonth)}</td>
            <td>${dateRange}</td>
            <td>${record.facilityName || '-'}</td>
            <td>${record.energyType}</td>
            <td>${formattedAmount} ${unit}</td>
            <td>${formattedCost} 원</td>
            <td>
                <button class="edit-btn" onclick="generateDocument(${index})">공문생성</button>
                <button class="edit-btn" onclick="generateAttachment1(${index})">첨부1</button>
                <button class="delete-btn" onclick="deleteEnergyRecord(${index})">삭제</button>
            </td>
        `;

        tbody.appendChild(row);
    });
}

window.generateDocument = async function(index) {
    try {
        const response = await fetch('/api/energy');
        const data = await response.json();

        if (data.success && data.records[index]) {
            const record = data.records[index];

            // 공문 생성 API 호출
            const docResponse = await fetch('/api/generate-document', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(record)
            });

            if (docResponse.ok) {
                // 파일 다운로드
                const blob = await docResponse.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;

                // 파일명 생성: {년도}-{월}-{시설명}-{에너지종류}-공문.docx
                const startDate = new Date(record.startDate);
                const year = startDate.getFullYear();
                const month = String(startDate.getMonth() + 1).padStart(2, '0');
                const filename = `${year}-${month}-${record.facilityName}-${record.energyType}-공문.docx`;

                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);

                alert('공문이 생성되었습니다.');
            } else {
                alert('공문 생성 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('공문 생성 오류:', error);
        alert('공문 생성 중 오류가 발생했습니다.');
    }
};

window.generateAttachment1 = async function(index) {
    try {
        const response = await fetch('/api/energy');
        const data = await response.json();

        if (data.success && data.records[index]) {
            const record = data.records[index];

            // 첨부1 생성 API 호출
            const attachmentResponse = await fetch('/api/generate-attachment1', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(record)
            });

            if (attachmentResponse.ok) {
                // 파일 다운로드
                const blob = await attachmentResponse.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;

                // 파일명 생성: {년도}-{월}-{시설명}-{에너지종류}-첨부1.xlsx
                const startDate = new Date(record.startDate);
                const year = startDate.getFullYear();
                const month = String(startDate.getMonth() + 1).padStart(2, '0');
                const filename = `${year}-${month}-${record.facilityName}-${record.energyType}-첨부1.xlsx`;

                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);

                alert('첨부1이 생성되었습니다.');
            } else {
                alert('첨부1 생성 중 오류가 발생했습니다.');
            }
        }
    } catch (error) {
        console.error('첨부1 생성 오류:', error);
        alert('첨부1 생성 중 오류가 발생했습니다.');
    }
};

window.deleteEnergyRecord = async function(index) {
    if (!confirm('정말 이 기록을 삭제하시겠습니까?')) {
        return;
    }

    try {
        const response = await fetch(`/api/energy/${index}`, {
            method: 'DELETE'
        });

        const data = await response.json();

        if (data.success) {
            alert(data.message);
            loadEnergyRecords();
        } else {
            alert(data.message);
        }
    } catch (error) {
        console.error('에너지 데이터 삭제 오류:', error);
        alert('에너지 데이터 삭제 중 오류가 발생했습니다.');
    }
};

// 데이터 조회 탭 관련 함수들
async function loadDataViewTab() {
    await loadFacilityOptions();
    searchData();
}

async function loadFacilityOptions() {
    try {
        const response = await fetch('/api/facilities');
        const data = await response.json();

        if (data.success) {
            const filterFacility = document.getElementById('filterFacility');
            filterFacility.innerHTML = '<option value="">전체</option>';

            const currentUser = JSON.parse(sessionStorage.getItem('currentUser'));

            // 시설관리자와 시설담당자의 경우 "전체" 옵션의 의미를 변경
            // 시설관리자: 본인 시설 + 하위 시설
            // 시설담당자: 본인 시설만

            // 관리자 역할을 가진 시설 제외 (시설 선택에서는 표시하지 않음)
            data.facilities
                .filter(facility => {
                    const isAdminRole = facility.role === '관리자';
                    return !isAdminRole;
                })
                .forEach(facility => {
                    const option = document.createElement('option');
                    option.value = facility.facilityName;
                    option.textContent = facility.facilityName;
                    filterFacility.appendChild(option);
                });
        }
    } catch (error) {
        console.error('시설 목록 로드 오류:', error);
    }
}

async function searchData() {
    const filterFacility = document.getElementById('filterFacility').value;
    const filterEnergyType = document.getElementById('filterEnergyType').value;
    const filterStartMonth = document.getElementById('filterStartMonth').value;
    const filterEndMonth = document.getElementById('filterEndMonth').value;

    try {
        const params = new URLSearchParams();
        if (filterFacility) params.append('facility', filterFacility);
        if (filterEnergyType) params.append('energyType', filterEnergyType);
        // 월분 기준으로 필터링 (YYYY-MM 형식)
        if (filterStartMonth) params.append('startDate', filterStartMonth + '-01');
        if (filterEndMonth) params.append('endDate', filterEndMonth + '-28');

        const response = await fetch(`/api/data-view?${params.toString()}`);
        const data = await response.json();

        if (data.success) {
            currentViewRecords = data.records; // 전역 변수에 저장
            window.currentViewRecords = data.records; // window 객체에도 저장하여 다른 파일에서 접근 가능
            renderDataView(data.records, 1); // 검색 시 첫 페이지부터 시작
            updateSummary(data.records);
            renderCharts(data.records);
        } else {
            alert(data.message || '데이터 조회 중 오류가 발생했습니다.');
        }
    } catch (error) {
        console.error('데이터 조회 오류:', error);
        alert('데이터 조회 중 오류가 발생했습니다.');
    }
}

// 페이지네이션 관련 전역 변수
let currentPage = 1;
const itemsPerPage = 20;

function renderDataView(records, page = 1) {
    const tbody = document.getElementById('dataViewTableBody');
    tbody.innerHTML = '';

    if (records.length === 0) {
        tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 40px; color: #999;">조회된 데이터가 없습니다.</td></tr>';
        renderPagination(0, 1);
        return;
    }

    // 최근 입력순으로 정렬 (배열 역순 - 마지막에 입력된 데이터가 먼저 표시)
    // 전역 변수로 저장하여 버튼 클릭 이벤트에서 접근 가능하게 함
    window.sortedViewRecords = [...records].reverse();
    const sortedRecords = window.sortedViewRecords;

    // 페이지네이션 계산
    currentPage = page;
    const totalPages = Math.ceil(sortedRecords.length / itemsPerPage);
    const startIndex = (page - 1) * itemsPerPage;
    const endIndex = Math.min(startIndex + itemsPerPage, sortedRecords.length);
    const pageRecords = sortedRecords.slice(startIndex, endIndex);

    pageRecords.forEach((record, idx) => {
        const index = startIndex + idx; // 전체 배열에서의 실제 인덱스
        const row = document.createElement('tr');

        // 통신의 경우 사용량이 0이면 빈칸으로 표시
        let formattedAmount = '';
        let unit = 'kWh';

        if (record.energyType === '통신') {
            unit = '건';
            if (record.usageAmount && parseFloat(record.usageAmount) > 0) {
                formattedAmount = parseFloat(record.usageAmount).toLocaleString('ko-KR') + ' ' + unit;
            }
        } else {
            formattedAmount = parseFloat(record.usageAmount).toLocaleString('ko-KR');
            if (record.energyType === '상하수도' || record.energyType === '도시가스') {
                unit = 'm³';
            }
            formattedAmount = formattedAmount + ' ' + unit;
        }

        const formattedCost = parseFloat(record.usageCost).toLocaleString('ko-KR');

        const startDate = record.startDate || record.usageDate;
        const endDate = record.endDate || record.usageDate;
        const dateRange = startDate === endDate ? startDate : `${startDate} ~ ${endDate}`;

        row.innerHTML = `
            <td><input type="checkbox" class="row-checkbox" data-index="${index}"></td>
            <td>${record.facilityName || '-'}</td>
            <td>${formatBillingMonth(record.billingMonth)}</td>
            <td>${dateRange}</td>
            <td>${record.energyType}</td>
            <td>${formattedAmount}</td>
            <td>${formattedCost} 원</td>
            <td>
                <button type="button" class="submit-btn generate-official-doc-btn" data-index="${index}" style="background-color: #f57c00; padding: 5px 10px; font-size: 12px;">공문 생성</button>
            </td>
            <td>
                <button type="button" class="submit-btn generate-attachment-btn" data-index="${index}" style="background-color: #7b1fa2; padding: 5px 10px; font-size: 12px;">첨부문서 생성</button>
            </td>
        `;
        tbody.appendChild(row);
    });

    // 페이지네이션 렌더링
    renderPagination(sortedRecords.length, totalPages);

    // 체크박스 변경 이벤트 리스너
    updateDeleteButtonVisibility();
    document.querySelectorAll('.row-checkbox').forEach(cb => {
        cb.addEventListener('change', updateDeleteButtonVisibility);
    });

    // 공문 생성 버튼 이벤트 리스너
    document.querySelectorAll('.generate-official-doc-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const index = parseInt(this.getAttribute('data-index'));
            const record = window.sortedViewRecords[index];
            generateOfficialDocument(record);
        });
    });

    // 첨부문서 생성 버튼 이벤트 리스너
    document.querySelectorAll('.generate-attachment-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const index = parseInt(this.getAttribute('data-index'));
            const record = window.sortedViewRecords[index];
            if (record) {
                generateAttachmentDocument(record);
            } else {
                alert('레코드를 찾을 수 없습니다. 페이지를 새로고침 해주세요.');
            }
        });
    });
}

// 페이지네이션 렌더링 함수
function renderPagination(totalItems, totalPages) {
    let paginationContainer = document.getElementById('paginationContainer');

    if (!paginationContainer) {
        // 페이지네이션 컨테이너가 없으면 테이블 아래에 생성
        const dataViewTable = document.querySelector('#dataView .data-table');
        if (dataViewTable) {
            paginationContainer = document.createElement('div');
            paginationContainer.id = 'paginationContainer';
            paginationContainer.style.cssText = 'display: flex; justify-content: center; align-items: center; gap: 5px; margin: 20px 0; flex-wrap: wrap;';
            dataViewTable.parentNode.insertBefore(paginationContainer, dataViewTable.nextSibling);
        } else {
            return;
        }
    }

    if (totalItems === 0) {
        paginationContainer.innerHTML = '';
        return;
    }

    let html = `<span style="margin-right: 15px; color: #666;">총 ${totalItems}건</span>`;

    // 이전 버튼
    html += `<button class="pagination-btn" onclick="goToPage(1)" ${currentPage === 1 ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${currentPage === 1 ? '#f5f5f5' : '#fff'}; border-radius: 4px;">≪</button>`;
    html += `<button class="pagination-btn" onclick="goToPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${currentPage === 1 ? '#f5f5f5' : '#fff'}; border-radius: 4px;">＜</button>`;

    // 페이지 번호들
    const maxVisiblePages = 5;
    let startPage = Math.max(1, currentPage - Math.floor(maxVisiblePages / 2));
    let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);

    if (endPage - startPage + 1 < maxVisiblePages) {
        startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }

    if (startPage > 1) {
        html += `<span style="margin: 0 5px;">...</span>`;
    }

    for (let i = startPage; i <= endPage; i++) {
        const isActive = i === currentPage;
        html += `<button class="pagination-btn" onclick="goToPage(${i})" style="padding: 5px 12px; margin: 2px; cursor: pointer; border: 1px solid ${isActive ? '#1976d2' : '#ddd'}; background: ${isActive ? '#1976d2' : '#fff'}; color: ${isActive ? '#fff' : '#333'}; border-radius: 4px; font-weight: ${isActive ? 'bold' : 'normal'};">${i}</button>`;
    }

    if (endPage < totalPages) {
        html += `<span style="margin: 0 5px;">...</span>`;
    }

    // 다음 버튼
    html += `<button class="pagination-btn" onclick="goToPage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${currentPage === totalPages ? '#f5f5f5' : '#fff'}; border-radius: 4px;">＞</button>`;
    html += `<button class="pagination-btn" onclick="goToPage(${totalPages})" ${currentPage === totalPages ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${currentPage === totalPages ? '#f5f5f5' : '#fff'}; border-radius: 4px;">≫</button>`;

    paginationContainer.innerHTML = html;
}

// 페이지 이동 함수 (전역 스코프에 노출)
window.goToPage = function(page) {
    if (page < 1) return;
    const records = window.currentViewRecords || [];
    if (records.length === 0) return;

    const totalPages = Math.ceil(records.length / itemsPerPage);
    if (page > totalPages) page = totalPages;

    renderDataView(records, page);

    // 테이블 상단으로 스크롤
    const dataViewTable = document.querySelector('#dataView .data-table');
    if (dataViewTable) {
        dataViewTable.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function updateSummary(records) {
    // 에너지 종류별로 집계
    const energyTypeSummary = {};

    records.forEach(record => {
        const type = record.energyType;
        if (!energyTypeSummary[type]) {
            energyTypeSummary[type] = {
                usage: 0,
                cost: 0,
                count: 0
            };
        }
        energyTypeSummary[type].usage += parseFloat(record.usageAmount) || 0;
        energyTypeSummary[type].cost += parseFloat(record.usageCost) || 0;
        energyTypeSummary[type].count += 1;
    });

    // 요약 카드 영역 업데이트
    const summaryCards = document.querySelector('.summary-cards');
    summaryCards.innerHTML = '';

    // 에너지 종류별 카드 생성
    Object.keys(energyTypeSummary).forEach(type => {
        const data = energyTypeSummary[type];

        let unit = 'kWh';
        if (type === '상하수도' || type === '도시가스') {
            unit = 'm³';
        } else if (type === '통신') {
            unit = '건';
        }

        const card = document.createElement('div');
        card.className = 'summary-card';

        // 통신의 경우 사용량이 0이면 표시하지 않음
        let usageDisplay = '';
        if (type === '통신' && data.usage > 0) {
            usageDisplay = `사용량: ${data.usage.toLocaleString('ko-KR')} ${unit}<br>`;
        } else if (type !== '통신') {
            usageDisplay = `사용량: ${data.usage.toLocaleString('ko-KR')} ${unit}<br>`;
        }

        card.innerHTML = `
            <div class="summary-label">${type}</div>
            <div class="summary-value" style="font-size: 18px;">
                ${usageDisplay}
                금액: ${data.cost.toLocaleString('ko-KR')}원<br>
                건수: ${data.count}건
            </div>
        `;
        summaryCards.appendChild(card);
    });

    // 카드가 없으면 기본 메시지 표시
    if (Object.keys(energyTypeSummary).length === 0) {
        summaryCards.innerHTML = '<div class="summary-card"><div class="summary-label">조회 결과 없음</div><div class="summary-value">0</div></div>';
    }
}

// 차트 관련 변수
let energyUsageChart = null;
let energyCostChart = null;

function renderCharts(records) {
    // 월분 기준으로 데이터 정렬
    const sortedRecords = [...records].sort((a, b) => {
        const monthA = a.billingMonth || a.startDate || '';
        const monthB = b.billingMonth || b.startDate || '';
        return monthA.localeCompare(monthB);
    });

    // 에너지 종류별로 데이터 분리
    const energyTypes = [...new Set(records.map(r => r.energyType))];

    // 시인성 좋은 색상 팔레트
    const colors = [
        '#2196F3', // 파랑
        '#F44336', // 빨강
        '#4CAF50', // 초록
        '#FF9800', // 주황
        '#9C27B0', // 보라
        '#00BCD4', // 청록
        '#FFEB3B', // 노랑
        '#795548'  // 갈색
    ];

    // 추세선 계산 함수 (선형 회귀)
    function calculateTrendline(data) {
        if (data.length < 2) return null;

        // 월분을 숫자로 변환 (YYYY-MM -> YYYYMM)
        const monthToNum = (m) => {
            const parts = m.split('-');
            return parseInt(parts[0]) * 12 + parseInt(parts[1]);
        };

        const months = data.map(d => monthToNum(d.x));
        const values = data.map(d => d.y);

        const n = months.length;
        const sumX = months.reduce((a, b) => a + b, 0);
        const sumY = values.reduce((a, b) => a + b, 0);
        const sumXY = months.reduce((sum, x, i) => sum + x * values[i], 0);
        const sumX2 = months.reduce((sum, x) => sum + x * x, 0);

        const slope = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
        const intercept = (sumY - slope * sumX) / n;

        return data.map(d => ({
            x: d.x,
            y: slope * monthToNum(d.x) + intercept
        }));
    }

    // 기간별 에너지 사용량 차트 (종류별)
    const ctx1 = document.getElementById('energyUsageChart');
    if (ctx1) {
        if (energyUsageChart) {
            energyUsageChart.destroy();
        }

        const usageDatasets = [];

        // 각 에너지 타입별 데이터와 추세선 추가
        energyTypes.forEach((type, index) => {
            const typeData = sortedRecords
                .filter(r => r.energyType === type)
                .map(r => ({
                    x: r.billingMonth || r.startDate?.substring(0, 7) || '',
                    y: parseFloat(r.usageAmount)
                }))
                .filter(d => d.x); // 빈 값 제거

            const color = colors[index % colors.length];

            // 실제 데이터
            usageDatasets.push({
                label: type,
                data: typeData,
                borderColor: color,
                backgroundColor: color,
                borderWidth: 3,
                pointRadius: 6,
                pointHoverRadius: 8,
                pointBackgroundColor: color,
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointHoverBorderWidth: 3,
                tension: 0.3,
                fill: false
            });

            // 추세선
            const trendline = calculateTrendline(typeData);
            if (trendline) {
                usageDatasets.push({
                    label: `${type} 추세`,
                    data: trendline,
                    borderColor: color,
                    backgroundColor: color,
                    borderWidth: 2,
                    borderDash: [5, 5],
                    pointRadius: 0,
                    pointHoverRadius: 0,
                    tension: 0,
                    fill: false,
                    opacity: 0.5
                });
            }
        });

        energyUsageChart = new Chart(ctx1, {
            type: 'line',
            data: {
                datasets: usageDatasets
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                interaction: {
                    mode: 'index',
                    intersect: false
                },
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            usePointStyle: true,
                            padding: 15,
                            font: {
                                size: 12,
                                weight: '500'
                            },
                            filter: function(item, chart) {
                                // '추세'가 포함된 라벨은 범례에서 제외
                                return !item.text.includes('추세');
                            }
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(0, 0, 0, 0.8)',
                        padding: 12,
                        titleFont: {
                            size: 14,
                            weight: 'bold'
                        },
                        bodyFont: {
                            size: 13
                        },
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                label += context.parsed.y.toLocaleString('ko-KR');
                                return label;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        type: 'category',
                        title: {
                            display: true,
                            text: '월분',
                            font: {
                                size: 14,
                                weight: 'bold'
                            }
                        },
                        grid: {
                            display: true,
                            color: 'rgba(0, 0, 0, 0.05)'
                        },
                        ticks: {
                            callback: function(value, index, values) {
                                // YYYY-MM 형식을 YYYY년 M월로 변환
                                const label = this.getLabelForValue(value);
                                if (label && label.includes('-')) {
                                    const parts = label.split('-');
                                    return parts[0] + '년 ' + parseInt(parts[1]) + '월';
                                }
                                return label;
                            }
                        }
                    },
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: '사용량',
                            font: {
                                size: 14,
                                weight: 'bold'
                            }
                        },
                        grid: {
                            display: true,
                            color: 'rgba(0, 0, 0, 0.08)',
                            lineWidth: 1
                        },
                        ticks: {
                            callback: function(value) {
                                return value.toLocaleString('ko-KR');
                            },
                            font: {
                                size: 12
                            }
                        }
                    }
                }
            }
        });
    }

    // 기간별 에너지 비용 차트 (종류별)
    const ctx2 = document.getElementById('energyCostChart');
    if (ctx2) {
        if (energyCostChart) {
            energyCostChart.destroy();
        }

        const costDatasets = [];

        // 각 에너지 타입별 데이터와 추세선 추가
        energyTypes.forEach((type, index) => {
            const typeData = sortedRecords
                .filter(r => r.energyType === type)
                .map(r => ({
                    x: r.billingMonth || r.startDate?.substring(0, 7) || '',
                    y: parseFloat(r.usageCost)
                }))
                .filter(d => d.x); // 빈 값 제거

            const color = colors[index % colors.length];

            // 실제 데이터
            costDatasets.push({
                label: type,
                data: typeData,
                borderColor: color,
                backgroundColor: color,
                borderWidth: 3,
                pointRadius: 6,
                pointHoverRadius: 8,
                pointBackgroundColor: color,
                pointBorderColor: '#fff',
                pointBorderWidth: 2,
                pointHoverBorderWidth: 3,
                tension: 0.3,
                fill: false
            });

            // 추세선
            const trendline = calculateTrendline(typeData);
            if (trendline) {
                costDatasets.push({
                    label: `${type} 추세`,
                    data: trendline,
                    borderColor: color,
                    backgroundColor: color,
                    borderWidth: 2,
                    borderDash: [5, 5],
                    pointRadius: 0,
                    pointHoverRadius: 0,
                    tension: 0,
                    fill: false,
                    opacity: 0.5
                });
            }
        });

        energyCostChart = new Chart(ctx2, {
            type: 'line',
            data: {
                datasets: costDatasets
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                interaction: {
                    mode: 'index',
                    intersect: false
                },
                plugins: {
                    legend: {
                        position: 'bottom',
                        labels: {
                            usePointStyle: true,
                            padding: 15,
                            font: {
                                size: 12,
                                weight: '500'
                            },
                            filter: function(item, chart) {
                                // '추세'가 포함된 라벨은 범례에서 제외
                                return !item.text.includes('추세');
                            }
                        }
                    },
                    tooltip: {
                        backgroundColor: 'rgba(0, 0, 0, 0.8)',
                        padding: 12,
                        titleFont: {
                            size: 14,
                            weight: 'bold'
                        },
                        bodyFont: {
                            size: 13
                        },
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                label += context.parsed.y.toLocaleString('ko-KR') + '원';
                                return label;
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        type: 'category',
                        title: {
                            display: true,
                            text: '월분',
                            font: {
                                size: 14,
                                weight: 'bold'
                            }
                        },
                        grid: {
                            display: true,
                            color: 'rgba(0, 0, 0, 0.05)'
                        },
                        ticks: {
                            callback: function(value, index, values) {
                                // YYYY-MM 형식을 YYYY년 M월로 변환
                                const label = this.getLabelForValue(value);
                                if (label && label.includes('-')) {
                                    const parts = label.split('-');
                                    return parts[0] + '년 ' + parseInt(parts[1]) + '월';
                                }
                                return label;
                            }
                        }
                    },
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: '비용 (원)',
                            font: {
                                size: 14,
                                weight: 'bold'
                            }
                        },
                        grid: {
                            display: true,
                            color: 'rgba(0, 0, 0, 0.08)',
                            lineWidth: 1
                        },
                        ticks: {
                            callback: function(value) {
                                return value.toLocaleString('ko-KR') + '원';
                            },
                            font: {
                                size: 12
                            }
                        }
                    }
                }
            }
        });
    }
}

// 전역 변수로 현재 조회된 레코드 저장
let currentViewRecords = [];

// 삭제 버튼 표시 업데이트
function updateDeleteButtonVisibility() {
    const checkboxes = document.querySelectorAll('.row-checkbox');
    const checkedCount = Array.from(checkboxes).filter(cb => cb.checked).length;
    const deleteBtn = document.getElementById('deleteSelectedBtn');

    if (deleteBtn) {
        deleteBtn.style.display = checkedCount > 0 ? 'inline-block' : 'none';
    }
}

// 선택된 항목 삭제
async function deleteSelectedRecords() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');

    if (checkboxes.length === 0) {
        alert('삭제할 항목을 선택해주세요.');
        return;
    }

    if (!confirm(`선택한 ${checkboxes.length}개 항목을 삭제하시겠습니까?`)) {
        return;
    }

    // sortedViewRecords를 사용 (테이블에 표시된 정렬된 배열)
    const sortedRecords = window.sortedViewRecords || currentViewRecords || [];
    const indices = Array.from(checkboxes).map(cb => parseInt(cb.dataset.index));
    const recordsToDelete = indices.map(i => sortedRecords[i]).filter(r => r);

    try {
        const response = await fetch('/api/energy-data/delete-multiple', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ records: recordsToDelete })
        });

        const result = await response.json();

        if (result.success) {
            alert('선택한 항목이 삭제되었습니다.');
            searchData(); // 데이터 다시 조회

            // 전체 선택 체크박스 해제
            const selectAllCheckbox = document.getElementById('selectAllCheckbox');
            if (selectAllCheckbox) {
                selectAllCheckbox.checked = false;
            }
        } else {
            alert('삭제 중 오류가 발생했습니다: ' + result.message);
        }
    } catch (error) {
        console.error('삭제 오류:', error);
        alert('삭제 중 오류가 발생했습니다.');
    }
}

// 이벤트 리스너 등록 (DOMContentLoaded 후)
document.addEventListener('DOMContentLoaded', function() {
    const searchDataBtn = document.getElementById('searchDataBtn');
    const resetFilterBtn = document.getElementById('resetFilterBtn');
    const deleteSelectedBtn = document.getElementById('deleteSelectedBtn');
    const selectAllCheckbox = document.getElementById('selectAllCheckbox');

    if (searchDataBtn) {
        searchDataBtn.addEventListener('click', searchData);
    }

    if (resetFilterBtn) {
        resetFilterBtn.addEventListener('click', () => {
            document.getElementById('filterFacility').value = '';
            document.getElementById('filterEnergyType').value = '';
            document.getElementById('filterStartMonth').value = '';
            document.getElementById('filterEndMonth').value = '';
            searchData();
        });
    }

    // 삭제 버튼 이벤트
    if (deleteSelectedBtn) {
        deleteSelectedBtn.addEventListener('click', deleteSelectedRecords);
    }

    // 전체 선택 체크박스 이벤트
    if (selectAllCheckbox) {
        selectAllCheckbox.addEventListener('change', function() {
            const checkboxes = document.querySelectorAll('.row-checkbox');
            checkboxes.forEach(cb => {
                cb.checked = this.checked;
            });
            updateDeleteButtonVisibility();
        });
    }

    // 상단 공문 생성 버튼 이벤트 (선택된 데이터의 합계 금액으로 공문 생성)
    const generateOfficialDocBtn = document.getElementById('generateOfficialDocBtn');
    if (generateOfficialDocBtn) {
        generateOfficialDocBtn.addEventListener('click', () => {
            // 선택된 레코드 가져오기
            const selectedRecords = getSelectedRecords();

            if (selectedRecords.length === 0) {
                alert('선택된 데이터가 없습니다.\n공문을 생성할 데이터를 체크박스로 선택해주세요.');
                return;
            }

            // 선택된 데이터의 합계 금액으로 공문 생성
            generateOfficialDocumentCombined(selectedRecords);
        });
    }

    // 상단 첨부문서 생성 버튼 이벤트 (선택된 데이터를 하나의 엑셀 파일로 생성)
    const generateAttachmentBtn = document.getElementById('generateAttachmentBtn');
    if (generateAttachmentBtn) {
        generateAttachmentBtn.addEventListener('click', () => {
            // 선택된 레코드 가져오기
            const selectedRecords = getSelectedRecords();

            if (selectedRecords.length === 0) {
                alert('선택된 데이터가 없습니다.\n첨부문서를 생성할 데이터를 체크박스로 선택해주세요.');
                return;
            }

            // 선택된 데이터를 하나의 엑셀 파일로 생성
            generateAttachmentDocumentCombined(selectedRecords);
        });
    }

    // 엑셀 다운로드 버튼 이벤트
    const downloadExcelBtn = document.getElementById('downloadExcelBtn');
    if (downloadExcelBtn) {
        downloadExcelBtn.addEventListener('click', downloadEnergyDataExcel);
    }

    // 엑셀 업로드 버튼 이벤트
    const uploadExcelBtn = document.getElementById('uploadExcelBtn');
    const excelFileInput = document.getElementById('excelFileInput');
    if (uploadExcelBtn && excelFileInput) {
        uploadExcelBtn.addEventListener('click', () => {
            excelFileInput.click();
        });
        excelFileInput.addEventListener('change', uploadEnergyDataExcel);
    }

    // 에너지 정보 관리 이벤트 리스너
    const energyInfoModal = document.getElementById('energyInfoModal');
    const closeEnergyInfoBtn = document.querySelector('.close-energy-info');
    const cancelEnergyInfoBtn = document.querySelector('.cancel-energy-info-btn');
    const addEnergyInfoBtn = document.getElementById('addEnergyInfoBtn');
    const energyInfoForm = document.getElementById('energyInfoForm');

    if (addEnergyInfoBtn) {
        addEnergyInfoBtn.addEventListener('click', () => {
            openEnergyInfoModal(false);
        });
    }

    if (closeEnergyInfoBtn) {
        closeEnergyInfoBtn.addEventListener('click', closeEnergyInfoModal);
    }

    if (cancelEnergyInfoBtn) {
        cancelEnergyInfoBtn.addEventListener('click', closeEnergyInfoModal);
    }

    window.addEventListener('click', (e) => {
        if (e.target === energyInfoModal) {
            closeEnergyInfoModal();
        }
    });

    if (energyInfoForm) {
        energyInfoForm.addEventListener('submit', async (e) => {
            e.preventDefault();

            const facilityName = document.getElementById('editInfoFacilityName').value.trim();
            const energyType = document.getElementById('editInfoEnergyType').value.trim();
            const customerNumber = document.getElementById('editInfoCustomerNumber').value.trim();
            const bankName = document.getElementById('editInfoBankName').value.trim();
            const accountNumber = document.getElementById('editInfoAccountNumber').value.trim();

            // 입력값 검증
            if (!facilityName) {
                alert('시설명을 선택해주세요.');
                return;
            }

            if (!energyType) {
                alert('에너지 종류를 선택해주세요.');
                return;
            }

            if (!customerNumber) {
                alert('고객번호(명세서번호)를 입력해주세요.');
                return;
            }

            if (!bankName) {
                alert('금융기관을 입력해주세요.');
                return;
            }

            if (!accountNumber) {
                alert('계좌번호를 입력해주세요.');
                return;
            }

            const energyInfoData = { facilityName, energyType, customerNumber, bankName, accountNumber };

            try {
                let response;
                if (currentEditEnergyInfoIndex !== null) {
                    response = await fetch(`/api/energy-info/${currentEditEnergyInfoIndex}`, {
                        method: 'PUT',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(energyInfoData)
                    });
                } else {
                    response = await fetch('/api/energy-info', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(energyInfoData)
                    });
                }

                // HTTP 상태 코드 확인
                if (!response.ok) {
                    console.error('HTTP 오류:', response.status, response.statusText);

                    if (response.status === 401) {
                        alert('로그인이 필요합니다. 다시 로그인해주세요.');
                        showLogin();
                        return;
                    }

                    if (response.status === 500) {
                        alert('서버 오류가 발생했습니다. 잠시 후 다시 시도해주세요.');
                        return;
                    }

                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }

                const data = await response.json();

                if (data.success) {
                    alert(data.message);
                    closeEnergyInfoModal();
                    loadEnergyInfos();
                } else {
                    alert(data.message || '에너지 정보 저장에 실패했습니다.');
                }
            } catch (error) {
                console.error('에너지 정보 저장 오류:', error);

                // 네트워크 오류와 기타 오류 구분
                if (error.name === 'TypeError' && error.message.includes('fetch')) {
                    alert('서버에 연결할 수 없습니다. 서버가 실행 중인지 확인하세요.');
                } else if (error.message.includes('JSON')) {
                    alert('서버 응답 형식이 올바르지 않습니다. 관리자에게 문의하세요.');
                } else {
                    alert(`에너지 정보 저장 중 오류가 발생했습니다.\n${error.message}`);
                }
            }
        });
    }
});

// 에너지 정보 관리 함수들
async function loadEnergyInfos() {
    try {
        const response = await fetch('/api/energy-info');

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const data = await response.json();

        if (data.success) {
            // 시설명 가나다순으로 정렬, 같은 시설명이면 에너지종류로 정렬
            const sortedInfos = (data.infos || []).sort((a, b) => {
                const facilityCompare = a.facilityName.localeCompare(b.facilityName, 'ko-KR');
                if (facilityCompare !== 0) return facilityCompare;
                return a.energyType.localeCompare(b.energyType, 'ko-KR');
            });

            renderEnergyInfos(sortedInfos);
        } else {
            console.error('에너지 정보 로드 실패:', data.message);
            renderEnergyInfos([]);
        }
    } catch (error) {
        console.error('에너지 정보 로드 오류:', error);
        renderEnergyInfos([]);

        // 서버 연결 오류인 경우에만 알림 표시
        if (error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
            alert('서버 연결 오류가 발생했습니다. 서버가 실행 중인지 확인하세요.');
        }
    }
}

function renderEnergyInfos(infos, page = 1) {
    const tbody = document.getElementById('energyInfoTableBody');
    tbody.innerHTML = '';

    // 전역 변수에 저장
    allEnergyInfos = infos || [];
    energyInfoCurrentPage = page;

    if (!infos || infos.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px; color: #999;">등록된 에너지 정보가 없습니다.</td></tr>';
        renderEnergyInfoPagination(0, 1);
        return;
    }

    // 페이지네이션 계산
    const totalPages = Math.ceil(infos.length / energyInfoItemsPerPage);
    const startIndex = (page - 1) * energyInfoItemsPerPage;
    const endIndex = Math.min(startIndex + energyInfoItemsPerPage, infos.length);
    const pageRecords = infos.slice(startIndex, endIndex);

    pageRecords.forEach((info, idx) => {
        const index = startIndex + idx; // 전체 배열에서의 실제 인덱스
        const row = document.createElement('tr');
        row.innerHTML = `
            <td><input type="checkbox" class="energy-info-checkbox" data-index="${index}"></td>
            <td>${info.facilityName || '-'}</td>
            <td>${info.energyType || '-'}</td>
            <td>${info.customerNumber || '-'}</td>
            <td>${info.bankName || '-'}</td>
            <td>${info.accountNumber || '-'}</td>
            <td>
                <button class="edit-btn" onclick="editEnergyInfo(${index})">수정</button>
                <button class="delete-btn" onclick="deleteEnergyInfo(${index})">삭제</button>
            </td>
        `;
        tbody.appendChild(row);
    });

    // 페이지네이션 렌더링
    renderEnergyInfoPagination(infos.length, totalPages);

    // 체크박스 변경 이벤트 리스너
    updateEnergyInfoDeleteButtonVisibility();
    document.querySelectorAll('.energy-info-checkbox').forEach(cb => {
        cb.addEventListener('change', updateEnergyInfoDeleteButtonVisibility);
    });
}

// 에너지 정보 페이지네이션 렌더링
function renderEnergyInfoPagination(totalItems, totalPages) {
    const container = document.getElementById('energyInfoPaginationContainer');
    if (!container) return;

    if (totalItems === 0) {
        container.innerHTML = '';
        return;
    }

    let html = `<span style="margin-right: 15px; color: #666;">총 ${totalItems}건</span>`;

    html += `<button class="pagination-btn" onclick="goToEnergyInfoPage(1)" ${energyInfoCurrentPage === 1 ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${energyInfoCurrentPage === 1 ? '#f5f5f5' : '#fff'}; border-radius: 4px;">≪</button>`;
    html += `<button class="pagination-btn" onclick="goToEnergyInfoPage(${energyInfoCurrentPage - 1})" ${energyInfoCurrentPage === 1 ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${energyInfoCurrentPage === 1 ? '#f5f5f5' : '#fff'}; border-radius: 4px;">＜</button>`;

    const maxVisiblePages = 5;
    let startPage = Math.max(1, energyInfoCurrentPage - Math.floor(maxVisiblePages / 2));
    let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);

    if (endPage - startPage + 1 < maxVisiblePages) {
        startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }

    if (startPage > 1) {
        html += `<span style="margin: 0 5px;">...</span>`;
    }

    for (let i = startPage; i <= endPage; i++) {
        const isActive = i === energyInfoCurrentPage;
        html += `<button class="pagination-btn" onclick="goToEnergyInfoPage(${i})" style="padding: 5px 12px; margin: 2px; cursor: pointer; border: 1px solid ${isActive ? '#1976d2' : '#ddd'}; background: ${isActive ? '#1976d2' : '#fff'}; color: ${isActive ? '#fff' : '#333'}; border-radius: 4px; font-weight: ${isActive ? 'bold' : 'normal'};">${i}</button>`;
    }

    if (endPage < totalPages) {
        html += `<span style="margin: 0 5px;">...</span>`;
    }

    html += `<button class="pagination-btn" onclick="goToEnergyInfoPage(${energyInfoCurrentPage + 1})" ${energyInfoCurrentPage === totalPages ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${energyInfoCurrentPage === totalPages ? '#f5f5f5' : '#fff'}; border-radius: 4px;">＞</button>`;
    html += `<button class="pagination-btn" onclick="goToEnergyInfoPage(${totalPages})" ${energyInfoCurrentPage === totalPages ? 'disabled' : ''} style="padding: 5px 10px; margin: 2px; cursor: pointer; border: 1px solid #ddd; background: ${energyInfoCurrentPage === totalPages ? '#f5f5f5' : '#fff'}; border-radius: 4px;">≫</button>`;

    container.innerHTML = html;
}

// 에너지 정보 페이지 이동
window.goToEnergyInfoPage = function(page) {
    if (page < 1 || allEnergyInfos.length === 0) return;
    const totalPages = Math.ceil(allEnergyInfos.length / energyInfoItemsPerPage);
    if (page > totalPages) page = totalPages;
    renderEnergyInfos(allEnergyInfos, page);
};

// 에너지 정보 체크박스 선택에 따라 삭제 버튼 표시
function updateEnergyInfoDeleteButtonVisibility() {
    const checkboxes = document.querySelectorAll('.energy-info-checkbox:checked');
    const deleteBtn = document.getElementById('deleteSelectedEnergyInfoBtn');

    if (checkboxes.length > 0) {
        if (deleteBtn) deleteBtn.style.display = 'inline-block';
    } else {
        if (deleteBtn) deleteBtn.style.display = 'none';
    }
}

async function openEnergyInfoModal(isEdit = false, index = null) {
    currentEditEnergyInfoIndex = index;
    const modalTitle = document.getElementById('energyInfoModalTitle');
    modalTitle.textContent = isEdit ? '에너지 정보 수정' : '에너지 정보 추가';

    // 시설 목록 로드
    await loadFacilityOptionsForEnergyInfo();

    if (!isEdit) {
        document.getElementById('editInfoFacilityName').value = '';
        document.getElementById('editInfoEnergyType').value = '';
        document.getElementById('editInfoCustomerNumber').value = '';
        document.getElementById('editInfoBankName').value = '';
        document.getElementById('editInfoAccountNumber').value = '';
    }

    const modal = document.getElementById('energyInfoModal');
    modal.style.display = 'block';
}

function closeEnergyInfoModal() {
    const modal = document.getElementById('energyInfoModal');
    modal.style.display = 'none';
    currentEditEnergyInfoIndex = null;
}

async function loadFacilityOptionsForEnergyInfo() {
    try {
        const response = await fetch('/api/facilities');
        const data = await response.json();

        if (data.success) {
            const select = document.getElementById('editInfoFacilityName');
            select.innerHTML = '<option value="">선택하세요</option>';

            // 관리자 역할 제외
            data.facilities
                .filter(f => f.role !== '관리자')
                .forEach(facility => {
                    const option = document.createElement('option');
                    option.value = facility.facilityName;
                    option.textContent = facility.facilityName;
                    select.appendChild(option);
                });
        }
    } catch (error) {
        console.error('시설 목록 로드 오류:', error);
    }
}

window.editEnergyInfo = async function(index) {
    try {
        const response = await fetch('/api/energy-info');
        const data = await response.json();

        if (data.success && data.infos[index]) {
            const info = data.infos[index];
            await openEnergyInfoModal(true, index);

            document.getElementById('editInfoFacilityName').value = info.facilityName;
            document.getElementById('editInfoEnergyType').value = info.energyType;
            document.getElementById('editInfoCustomerNumber').value = info.customerNumber;
            document.getElementById('editInfoBankName').value = info.bankName;
            document.getElementById('editInfoAccountNumber').value = info.accountNumber;
        }
    } catch (error) {
        console.error('에너지 정보 로드 오류:', error);
        alert('에너지 정보를 불러오는 중 오류가 발생했습니다.');
    }
};

window.deleteEnergyInfo = async function(index) {
    if (!confirm('정말 이 에너지 정보를 삭제하시겠습니까?')) {
        return;
    }

    try {
        const response = await fetch(`/api/energy-info/${index}`, {
            method: 'DELETE'
        });

        const data = await response.json();

        if (data.success) {
            alert(data.message);
            loadEnergyInfos();
        } else {
            alert(data.message);
        }
    } catch (error) {
        console.error('에너지 정보 삭제 오류:', error);
        alert('에너지 정보 삭제 중 오류가 발생했습니다.');
    }
};

// 클라이언트 사이드 OCR 기능 (Tesseract.js 및 네이버 클로바 OCR)
document.addEventListener('DOMContentLoaded', function() {
    const ocrUploadBtn = document.getElementById('ocrUploadBtn');
    const ocrFileInput = document.getElementById('ocrFileInput');
    const ocrFileName = document.getElementById('ocrFileName');
    const ocrProgress = document.getElementById('ocrProgress');
    const progressText = ocrProgress?.querySelector('.progress-text');
    const progressFill = ocrProgress?.querySelector('.progress-fill');
    const viewOcrTextBtn = document.getElementById('viewOcrTextBtn');

    // OCR 방식 선택 관련 요소
    const ocrMethodRadios = document.querySelectorAll('input[name="ocrMethod"]');
    const clovaApiKeySection = document.getElementById('clovaApiKeySection');
    const saveApiKeyBtn = document.getElementById('saveApiKeyBtn');
    const clovaApiUrlInput = document.getElementById('clovaApiUrl');
    const clovaSecretKeyInput = document.getElementById('clovaSecretKey');

    // OCR 원본 텍스트를 전역 변수로 저장
    let lastOcrText = '';
    let lastExtractedData = null;

    // OCR 방식 선택 시 API 키 입력란 표시/숨김
    ocrMethodRadios.forEach(radio => {
        radio.addEventListener('change', (e) => {
            if (e.target.value === 'clova') {
                clovaApiKeySection.style.display = 'block';
            } else {
                clovaApiKeySection.style.display = 'none';
            }
        });
    });

    // API 키 저장 버튼
    if (saveApiKeyBtn) {
        // 페이지 로드 시 저장된 API 키 불러오기
        const savedApiUrl = localStorage.getItem('clovaApiUrl');
        const savedSecretKey = localStorage.getItem('clovaSecretKey');
        if (savedApiUrl) clovaApiUrlInput.value = savedApiUrl;
        if (savedSecretKey) clovaSecretKeyInput.value = savedSecretKey;

        saveApiKeyBtn.addEventListener('click', () => {
            const apiUrl = clovaApiUrlInput.value.trim();
            const secretKey = clovaSecretKeyInput.value.trim();

            if (!apiUrl || !secretKey) {
                alert('API URL과 Secret Key를 모두 입력해주세요.');
                return;
            }

            // LocalStorage에 저장
            localStorage.setItem('clovaApiUrl', apiUrl);
            localStorage.setItem('clovaSecretKey', secretKey);

            alert('✅ API 키가 저장되었습니다!\n이제 클로바 OCR을 사용할 수 있습니다.');
        });
    }

    // API 키 검증 버튼
    const validateApiKeyBtn = document.getElementById('validateApiKeyBtn');
    if (validateApiKeyBtn) {
        validateApiKeyBtn.addEventListener('click', async () => {
            const apiUrl = clovaApiUrlInput.value.trim();
            const secretKey = clovaSecretKeyInput.value.trim();

            if (!apiUrl || !secretKey) {
                alert('API URL과 Secret Key를 모두 입력해주세요.');
                return;
            }

            // 버튼 비활성화
            validateApiKeyBtn.disabled = true;
            validateApiKeyBtn.textContent = '검증 중...';

            try {
                console.log('API 키 검증 시작');
                const response = await fetch('/api/validate-clova-key', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        apiUrl,
                        secretKey
                    })
                });

                const data = await response.json();
                console.log('검증 결과:', data);

                if (data.success) {
                    alert('✅ ' + data.message + '\n\nAPI 키가 정상적으로 작동합니다.\n이제 고지서를 업로드하여 OCR을 사용할 수 있습니다.');

                    // 검증 성공 시 자동 저장
                    localStorage.setItem('clovaApiUrl', apiUrl);
                    localStorage.setItem('clovaSecretKey', secretKey);
                } else {
                    alert(data.message + '\n\n' + (data.details ? JSON.stringify(data.details, null, 2) : ''));
                }

            } catch (error) {
                console.error('API 키 검증 오류:', error);
                alert('❌ API 키 검증 중 오류가 발생했습니다.\n\n' + error.message + '\n\nF12를 눌러 콘솔에서 상세 오류를 확인하세요.');
            } finally {
                // 버튼 다시 활성화
                validateApiKeyBtn.disabled = false;
                validateApiKeyBtn.textContent = 'API 키 검증';
            }
        });
    }

    // 파일을 Base64로 변환하는 함수
    function fileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(file);
        });
    }

    // 이미지 압축 함수 (클로바 OCR용 - 크기 제한 대응)
    async function compressImageForClova(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                const img = new Image();

                img.onload = () => {
                    const canvas = document.createElement('canvas');
                    const ctx = canvas.getContext('2d');

                    // 이미지 크기 최적화 (2000px 이하로 제한, OCR 품질 유지)
                    let width = img.width;
                    let height = img.height;
                    const maxSize = 2000;

                    if (width > maxSize || height > maxSize) {
                        const ratio = Math.min(maxSize / width, maxSize / height);
                        width = Math.floor(width * ratio);
                        height = Math.floor(height * ratio);
                    }

                    canvas.width = width;
                    canvas.height = height;

                    // 고품질 이미지 그리기
                    ctx.imageSmoothingEnabled = true;
                    ctx.imageSmoothingQuality = 'high';
                    ctx.drawImage(img, 0, 0, width, height);

                    // JPEG로 압축 (품질 0.85 - OCR 품질과 파일 크기 균형)
                    const compressedDataUrl = canvas.toDataURL('image/jpeg', 0.85);

                    console.log('원본 크기:', file.size, 'bytes');
                    console.log('압축 후 Base64 길이:', compressedDataUrl.length);
                    console.log('압축 비율:', Math.round((compressedDataUrl.length / file.size) * 100) + '%');

                    resolve(compressedDataUrl);
                };

                img.onerror = () => {
                    reject(new Error('이미지 로드 실패'));
                };

                img.src = e.target.result;
            };

            reader.onerror = () => {
                reject(new Error('파일 읽기 실패'));
            };

            reader.readAsDataURL(file);
        });
    }

    // 이미지 전처리 함수 (OCR 인식률 향상 - 고급)
    async function preprocessImage(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                const img = new Image();

                img.onload = () => {
                    console.log('원본 이미지 크기:', img.width, 'x', img.height);

                    // Canvas 생성
                    const canvas = document.createElement('canvas');
                    const ctx = canvas.getContext('2d');

                    // 이미지 크기 최적화 (너무 작으면 확대, 너무 크면 축소)
                    let width = img.width;
                    let height = img.height;
                    const maxSize = 3500;  // 최대 크기 증가
                    const minSize = 1500;  // 최소 크기 증가

                    // 이미지가 너무 작으면 3배 확대 (더 큰 확대)
                    if (width < minSize || height < minSize) {
                        width *= 3;
                        height *= 3;
                    }

                    // 이미지가 너무 크면 축소
                    if (width > maxSize || height > maxSize) {
                        const ratio = Math.min(maxSize / width, maxSize / height);
                        width *= ratio;
                        height *= ratio;
                    }

                    console.log('리사이즈 후 크기:', width, 'x', height);

                    canvas.width = width;
                    canvas.height = height;

                    // 고품질 이미지 그리기
                    ctx.imageSmoothingEnabled = true;
                    ctx.imageSmoothingQuality = 'high';
                    ctx.drawImage(img, 0, 0, width, height);

                    // 이미지 데이터 가져오기
                    let imageData = ctx.getImageData(0, 0, width, height);
                    let data = imageData.data;

                    // 1단계: 그레이스케일 변환
                    const grayData = new Uint8ClampedArray(width * height);
                    for (let i = 0; i < data.length; i += 4) {
                        const gray = data[i] * 0.299 + data[i + 1] * 0.587 + data[i + 2] * 0.114;
                        grayData[i / 4] = gray;
                    }

                    // 2단계: 가우시안 블러 (노이즈 제거)
                    const blurredData = gaussianBlur(grayData, width, height);

                    // 3단계: Otsu's 이진화 (자동 임계값 계산)
                    const threshold = calculateOtsuThreshold(blurredData);
                    console.log('자동 계산된 임계값:', threshold);

                    // 4단계: 적응형 대비 향상 (CLAHE 간소화 버전)
                    for (let i = 0; i < blurredData.length; i++) {
                        let value = blurredData[i];

                        // 대비 향상 (2.0배)
                        value = ((value - 128) * 2.0) + 128;

                        // 클리핑
                        value = Math.max(0, Math.min(255, value));

                        // 이진화
                        const binary = value > threshold ? 255 : 0;

                        // RGB 모두 같은 값
                        const idx = i * 4;
                        data[idx] = binary;
                        data[idx + 1] = binary;
                        data[idx + 2] = binary;
                    }

                    // 5단계: 형태학적 연산 (열림 연산 - 노이즈 제거)
                    morphologicalOpening(data, width, height);

                    // 처리된 이미지 데이터를 캔버스에 적용
                    ctx.putImageData(imageData, 0, 0);

                    console.log('이미지 전처리 완료');

                    // Canvas를 Blob URL로 변환
                    canvas.toBlob((blob) => {
                        const url = URL.createObjectURL(blob);
                        resolve(url);
                    }, 'image/png', 1.0);
                };

                img.onerror = () => {
                    reject(new Error('이미지 로드 실패'));
                };

                img.src = e.target.result;
            };

            reader.onerror = () => {
                reject(new Error('파일 읽기 실패'));
            };

            reader.readAsDataURL(file);
        });
    }

    // 가우시안 블러 (3x3 커널)
    function gaussianBlur(data, width, height) {
        const kernel = [
            [1, 2, 1],
            [2, 4, 2],
            [1, 2, 1]
        ];
        const kernelSum = 16;

        const result = new Uint8ClampedArray(data.length);

        for (let y = 1; y < height - 1; y++) {
            for (let x = 1; x < width - 1; x++) {
                let sum = 0;
                for (let ky = -1; ky <= 1; ky++) {
                    for (let kx = -1; kx <= 1; kx++) {
                        const idx = (y + ky) * width + (x + kx);
                        sum += data[idx] * kernel[ky + 1][kx + 1];
                    }
                }
                result[y * width + x] = sum / kernelSum;
            }
        }

        return result;
    }

    // Otsu's 방법으로 최적 임계값 계산
    function calculateOtsuThreshold(data) {
        // 히스토그램 생성
        const histogram = new Array(256).fill(0);
        for (let i = 0; i < data.length; i++) {
            histogram[Math.floor(data[i])]++;
        }

        // 전체 픽셀 수
        const total = data.length;

        let sum = 0;
        for (let i = 0; i < 256; i++) {
            sum += i * histogram[i];
        }

        let sumB = 0;
        let wB = 0;
        let wF = 0;
        let maxVariance = 0;
        let threshold = 0;

        for (let t = 0; t < 256; t++) {
            wB += histogram[t];
            if (wB === 0) continue;

            wF = total - wB;
            if (wF === 0) break;

            sumB += t * histogram[t];

            const mB = sumB / wB;
            const mF = (sum - sumB) / wF;

            const variance = wB * wF * (mB - mF) * (mB - mF);

            if (variance > maxVariance) {
                maxVariance = variance;
                threshold = t;
            }
        }

        return threshold;
    }

    // 형태학적 열림 연산 (침식 후 팽창)
    function morphologicalOpening(data, width, height) {
        // 간단한 3x3 구조 요소
        const temp = new Uint8ClampedArray(data.length);

        // 침식 (Erosion)
        for (let y = 1; y < height - 1; y++) {
            for (let x = 1; x < width - 1; x++) {
                let min = 255;
                for (let ky = -1; ky <= 1; ky++) {
                    for (let kx = -1; kx <= 1; kx++) {
                        const idx = ((y + ky) * width + (x + kx)) * 4;
                        min = Math.min(min, data[idx]);
                    }
                }
                const idx = (y * width + x) * 4;
                temp[idx] = temp[idx + 1] = temp[idx + 2] = min;
                temp[idx + 3] = 255;
            }
        }

        // 팽창 (Dilation)
        for (let y = 1; y < height - 1; y++) {
            for (let x = 1; x < width - 1; x++) {
                let max = 0;
                for (let ky = -1; ky <= 1; ky++) {
                    for (let kx = -1; kx <= 1; kx++) {
                        const idx = ((y + ky) * width + (x + kx)) * 4;
                        max = Math.max(max, temp[idx]);
                    }
                }
                const idx = (y * width + x) * 4;
                data[idx] = data[idx + 1] = data[idx + 2] = max;
            }
        }
    }

    if (ocrUploadBtn && ocrFileInput) {
        // 업로드 버튼 클릭 시 파일 선택 다이얼로그 열기
        ocrUploadBtn.addEventListener('click', () => {
            ocrFileInput.click();
        });

        // 파일 선택 시 OCR 처리 시작
        ocrFileInput.addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (!file) return;

            // 파일 형식 검증 (이미지만)
            const allowedTypes = ['image/jpeg', 'image/jpg', 'image/png'];
            if (!allowedTypes.includes(file.type)) {
                alert('JPEG, PNG 이미지 파일만 업로드 가능합니다.');
                ocrFileInput.value = '';
                return;
            }

            // 파일 크기 검증 (10MB)
            if (file.size > 10 * 1024 * 1024) {
                alert('파일 크기는 10MB 이하여야 합니다.');
                ocrFileInput.value = '';
                return;
            }

            // 파일명 표시
            ocrFileName.textContent = file.name;

            // OCR 방식 확인
            const ocrMethod = document.querySelector('input[name="ocrMethod"]:checked').value;

            // 프로그레스 표시
            ocrProgress.style.display = 'flex';
            if (progressFill) progressFill.style.width = '0%';
            if (progressText) progressText.textContent = '이미지 분석 준비 중...';

            try {
                let ocrText = '';

                if (ocrMethod === 'clova') {
                    // 네이버 클로바 OCR 사용
                    const apiUrl = localStorage.getItem('clovaApiUrl');
                    const secretKey = localStorage.getItem('clovaSecretKey');

                    if (!apiUrl || !secretKey) {
                        alert('먼저 클로바 OCR API 키를 입력하고 저장해주세요.');
                        ocrProgress.style.display = 'none';
                        ocrFileInput.value = '';
                        return;
                    }

                    console.log('클로바 OCR 시작');
                    console.log('API URL:', apiUrl);
                    console.log('파일 크기:', file.size, 'bytes');
                    console.log('파일 타입:', file.type);

                    if (progressText) progressText.textContent = '이미지 압축 중...';
                    if (progressFill) progressFill.style.width = '30%';

                    // 이미지 압축 (413 오류 방지)
                    const imageBase64 = await compressImageForClova(file);
                    console.log('이미지 압축 및 Base64 변환 완료');

                    if (progressText) progressText.textContent = '클로바 OCR 처리 중...';
                    if (progressFill) progressFill.style.width = '60%';

                    // 서버를 통해 클로바 OCR API 호출
                    console.log('서버로 요청 전송 중...');
                    const response = await fetch('/api/clova-ocr', {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json'
                        },
                        body: JSON.stringify({
                            imageBase64,
                            apiUrl,
                            secretKey
                        })
                    });

                    console.log('서버 응답 상태:', response.status);

                    if (!response.ok) {
                        const errorText = await response.text();
                        console.error('서버 응답 오류:', errorText);
                        throw new Error(`서버 오류 (${response.status}): ${errorText}`);
                    }

                    const data = await response.json();
                    console.log('서버 응답 데이터:', data);

                    if (!data.success) {
                        console.error('클로바 OCR 실패:', data);
                        let errorMsg = data.message || '클로바 OCR 처리 실패';
                        if (data.details) {
                            console.error('상세 오류:', data.details);
                            errorMsg += '\n\n상세 정보:\n' + JSON.stringify(data.details, null, 2);
                        }
                        throw new Error(errorMsg);
                    }

                    ocrText = data.text;
                    console.log('클로바 OCR 성공!');
                    console.log('추출된 텍스트 길이:', ocrText.length);
                    console.log('클로바 OCR 전체 응답:', data.fullResponse);

                } else {
                    // Tesseract.js 사용 (기본값)
                    // 이미지 전처리 (인식률 향상)
                    if (progressText) progressText.textContent = '이미지 전처리 중...';
                    const preprocessedImage = await preprocessImage(file);

                    // Tesseract.js로 OCR 실행 (한글 인식 최적화)
                    console.log('Tesseract OCR 시작 - 고급 설정 적용');
                    const result = await Tesseract.recognize(
                        preprocessedImage,
                        'kor+eng', // 한국어 우선 + 영어
                        {
                            logger: (m) => {
                                // 진행률 표시
                                if (m.status === 'recognizing text') {
                                    const progress = Math.round(m.progress * 100);
                                    if (progressFill) progressFill.style.width = `${progress}%`;
                                    if (progressText) progressText.textContent = `텍스트 인식 중... ${progress}%`;
                                } else if (m.status === 'loading tesseract core') {
                                    if (progressText) progressText.textContent = 'OCR 엔진 로딩 중...';
                                } else if (m.status === 'initializing tesseract') {
                                    if (progressText) progressText.textContent = 'OCR 초기화 중...';
                                } else if (m.status === 'loading language traineddata') {
                                    if (progressText) progressText.textContent = '한글 언어 데이터 로딩 중...';
                                }
                            },
                            // Tesseract 설정 최적화 (고급)
                            tessedit_pageseg_mode: Tesseract.PSM.AUTO,  // 자동 페이지 분석
                            tessedit_ocr_engine_mode: Tesseract.OEM.LSTM_ONLY,  // LSTM 엔진 사용 (정확도 향상)
                            preserve_interword_spaces: '1',  // 단어 간 공백 유지
                            tessedit_char_whitelist: '',  // 모든 문자 허용
                            // 한글 인식 개선
                            textord_heavy_nr: '1',  // 노이즈가 많은 이미지 처리
                            textord_min_linesize: '2.5',  // 최소 라인 크기
                            // 품질 개선
                            language_model_penalty_non_freq_dict_word: '0.5',  // 사전에 없는 단어 패널티 감소
                            language_model_penalty_non_dict_word: '0.5',
                        }
                    );

                    console.log('Tesseract OCR 완료');
                    console.log('신뢰도:', result.data.confidence);

                    // URL 해제
                    URL.revokeObjectURL(preprocessedImage);

                    ocrText = result.data.text;
                }

                // OCR 텍스트에서 정보 추출
                const extractedData = extractEnergyDataFromText(ocrText);

                // 프로그레스 숨김
                ocrProgress.style.display = 'none';

                // 추출된 데이터를 폼에 자동 입력
                fillFormWithOCRData(extractedData);

                // 최근 입력 내역에 미리보기 추가
                addOCRPreviewToHistory(extractedData);

                // OCR 원본 텍스트 저장
                lastOcrText = ocrText;
                lastExtractedData = extractedData;

                // "인식된 텍스트 보기" 버튼 표시
                if (viewOcrTextBtn) {
                    viewOcrTextBtn.style.display = 'inline-flex';
                }

                // OCR 원본 텍스트를 콘솔에 출력
                console.log('==================================================');
                console.log(`📄 OCR 인식된 원본 텍스트 (${ocrMethod === 'clova' ? '클로바' : 'Tesseract'}):`);
                console.log('==================================================');
                console.log(ocrText);
                console.log('==================================================');
                console.log('');
                console.log('추출된 데이터:', extractedData);

                // 추출된 정보 요약
                const extractedInfo = [];
                const notExtractedInfo = [];

                // 추출 성공 항목
                if (extractedData.energyType) {
                    extractedInfo.push(`• 에너지 종류: ${extractedData.energyType}`);
                } else {
                    notExtractedInfo.push('에너지 종류');
                }

                if (extractedData.billingMonth) {
                    extractedInfo.push(`• 월분: ${formatBillingMonth(extractedData.billingMonth)}`);
                }

                if (extractedData.startDate) {
                    extractedInfo.push(`• 사용 시작일: ${extractedData.startDate}`);
                } else {
                    notExtractedInfo.push('사용 시작일');
                }

                if (extractedData.endDate) {
                    extractedInfo.push(`• 사용 종료일: ${extractedData.endDate}`);
                } else {
                    notExtractedInfo.push('사용 종료일');
                }

                if (extractedData.usageAmount) {
                    extractedInfo.push(`• 사용량: ${extractedData.usageAmount}`);
                } else {
                    notExtractedInfo.push('사용량');
                }

                if (extractedData.usageCost) {
                    extractedInfo.push(`• 사용 금액: ${formatNumber(extractedData.usageCost)}원`);
                } else {
                    notExtractedInfo.push('사용 금액');
                }

                let message = '✅ 고지서 분석 완료!\n\n';

                if (extractedInfo.length > 0) {
                    message += '📋 추출된 정보:\n' + extractedInfo.join('\n') + '\n\n';
                }

                if (notExtractedInfo.length > 0) {
                    message += '⚠️ 추출 실패 항목:\n• ' + notExtractedInfo.join('\n• ') + '\n\n';
                    message += '💡 팁: F12를 눌러 콘솔에서 "OCR 인식된 원본 텍스트"를 확인하면\n어떤 내용이 인식되었는지 볼 수 있습니다.\n\n';
                }

                message += '입력창이 녹색으로 강조된 항목을 확인하고\n누락된 항목은 수동으로 입력해주세요.';

                alert(message);

                // 추출 실패한 항목이 있으면 추가 안내
                if (notExtractedInfo.length > 0) {
                    console.warn('⚠️ 다음 항목이 추출되지 않았습니다:', notExtractedInfo);
                    console.log('💡 이미지 품질을 높이거나 다음 사항을 확인하세요:');
                    console.log('  - 고지서 이미지가 선명한가요?');
                    console.log('  - 글자가 잘 보이나요?');
                    console.log('  - 사용 기간, 사용량, 금액이 명확히 표시되어 있나요?');
                }

            } catch (error) {
                console.error('OCR 처리 오류:', error);
                ocrProgress.style.display = 'none';

                // 오류 메시지 상세 표시
                let errorMessage = '❌ 이미지 분석 중 오류가 발생했습니다.\n\n';

                if (ocrMethod === 'clova') {
                    errorMessage += '클로바 OCR 오류:\n';
                    errorMessage += error.message + '\n\n';
                    errorMessage += '확인 사항:\n';
                    errorMessage += '1. API URL이 정확한가요?\n';
                    errorMessage += '   (예: https://xxxxx.apigw.ntruss.com/custom/v1/xxxxx/xxxxxxxx)\n';
                    errorMessage += '2. Secret Key가 정확한가요?\n';
                    errorMessage += '3. 네이버 클라우드 플랫폼에서 OCR API가 활성화되어 있나요?\n';
                    errorMessage += '4. API 사용량이 남아있나요?\n\n';
                    errorMessage += '💡 F12를 눌러 콘솔에서 상세 오류를 확인하세요.';
                } else {
                    errorMessage += 'Tesseract OCR 오류:\n';
                    errorMessage += error.message + '\n\n';
                    errorMessage += '다른 이미지를 시도하거나 클로바 OCR을 사용해보세요.';
                }

                alert(errorMessage);
            } finally {
                // 파일 입력 초기화
                ocrFileInput.value = '';
            }
        });
    }

    // "인식된 텍스트 보기" 버튼 이벤트 리스너
    if (viewOcrTextBtn) {
        viewOcrTextBtn.addEventListener('click', () => {
            if (lastOcrText) {
                const extractedInfo = [];
                if (lastExtractedData) {
                    if (lastExtractedData.energyType) extractedInfo.push(`• 에너지 종류: ${lastExtractedData.energyType}`);
                    if (lastExtractedData.billingMonth) extractedInfo.push(`• 월분: ${formatBillingMonth(lastExtractedData.billingMonth)}`);
                    if (lastExtractedData.startDate || lastExtractedData.endDate) {
                        extractedInfo.push(`• 사용기간: ${lastExtractedData.startDate || '?'} ~ ${lastExtractedData.endDate || '?'}`);
                    }
                    if (lastExtractedData.usageAmount) extractedInfo.push(`• 사용량: ${lastExtractedData.usageAmount}`);
                    if (lastExtractedData.usageCost) extractedInfo.push(`• 사용금액: ${lastExtractedData.usageCost}원`);
                    if (lastExtractedData.customerNumber) extractedInfo.push(`• 고객번호: ${lastExtractedData.customerNumber}`);
                }

                let message = '📄 OCR로 인식된 원본 텍스트:\n\n';
                message += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
                message += lastOcrText.substring(0, 1000);  // 처음 1000자만 표시
                if (lastOcrText.length > 1000) {
                    message += '\n\n... (텍스트가 너무 길어 일부만 표시됩니다)\n... (전체 내용은 F12 콘솔에서 확인하세요)';
                }
                message += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

                if (extractedInfo.length > 0) {
                    message += '✅ 추출된 정보:\n' + extractedInfo.join('\n');
                } else {
                    message += '⚠️ 추출된 정보가 없습니다.';
                }

                alert(message);

                // 전체 텍스트를 콘솔에도 출력
                console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
                console.log('📄 OCR 인식된 전체 텍스트:');
                console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
                console.log(lastOcrText);
                console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
            } else {
                alert('먼저 고지서 이미지를 업로드해주세요.');
            }
        });
    }
});

// OCR 텍스트에서 에너지 정보 추출
function extractEnergyDataFromText(text) {
    const result = {
        energyType: '',
        customerNumber: '',
        billingMonth: '',
        startDate: '',
        endDate: '',
        usageAmount: '',
        usageCost: ''
    };

    // 에너지 종류 자동 감지
    if (text.match(/전기|한국전력|KEPCO|전력|kWh/i)) {
        result.energyType = '전기';
    } else if (text.match(/상수도|하수도|수도|워터|water|m³|㎥/i)) {
        result.energyType = '상하수도';
    } else if (text.match(/도시가스|가스|gas|LNG/i)) {
        result.energyType = '도시가스';
    } else if (text.match(/통신|인터넷|전화|SK|KT|LG|U\+/i)) {
        result.energyType = '통신';
    }

    // 고객번호 추출 (다양한 패턴 지원)
    const customerPatterns = [
        /고객\s*번호\s*[:\s]*([0-9-]+)/i,
        /고객NO\s*[:\s]*([0-9-]+)/i,
        /계약\s*번호\s*[:\s]*([0-9-]+)/i,
        /명세서\s*번호\s*[:\s]*([0-9-]+)/i,
        /청구\s*번호\s*[:\s]*([0-9-]+)/i,
        /고객\s*NO\s*[:\s]*([0-9-]+)/i
    ];

    for (const pattern of customerPatterns) {
        const match = text.match(pattern);
        if (match) {
            result.customerNumber = match[1].trim();
            break;
        }
    }

    // 월분 추출 (예: 2024년 1월분, 2024.01월분, 1월분 등) - month 형식(YYYY-MM)으로 저장
    const billingMonthPatterns = [
        /(\d{4})\s*년\s*(\d{1,2})\s*월\s*분/i,
        /(\d{4})\s*[.\/]\s*(\d{1,2})\s*월\s*분/i,
        /(\d{1,2})\s*월\s*분/i
    ];

    for (const pattern of billingMonthPatterns) {
        const match = text.match(pattern);
        if (match) {
            let year, month;
            if (match.length === 3 && match[1].length === 4) {
                // YYYY년 MM월분 형태
                year = match[1];
                month = match[2].padStart(2, '0');
            } else if (match.length === 2) {
                // MM월분 형태 - 현재 연도 추가
                year = new Date().getFullYear();
                month = match[1].padStart(2, '0');
            }
            // month 타입 input에 맞는 형식(YYYY-MM)으로 저장
            result.billingMonth = `${year}-${month}`;
            break;
        }
    }

    // 사용기간 추출
    const periodPatterns = [
        /사용\s*기간\s*[:\s]*(\d{4}[-./년\s]\s*\d{1,2}[-./월\s]\s*\d{1,2})[일\s]*[-~～]\s*(\d{4}[-./년\s]\s*\d{1,2}[-./월\s]\s*\d{1,2})[일\s]*/i,
        /이용\s*기간\s*[:\s]*(\d{4}[-./년\s]\s*\d{1,2}[-./월\s]\s*\d{1,2})[일\s]*[-~～]\s*(\d{4}[-./년\s]\s*\d{1,2}[-./월\s]\s*\d{1,2})[일\s]*/i,
        /기간\s*[:\s]*(\d{4}[-./년\s]\s*\d{1,2}[-./월\s]\s*\d{1,2})[일\s]*[-~～]\s*(\d{4}[-./년\s]\s*\d{1,2}[-./월\s]\s*\d{1,2})[일\s]*/i,
        /(\d{4}[-./년\s]\s*\d{1,2}[-./월\s]\s*\d{1,2})[일\s]*[-~～]\s*(\d{4}[-./년\s]\s*\d{1,2}[-./월\s]\s*\d{1,2})[일\s]*/
    ];

    for (const pattern of periodPatterns) {
        const match = text.match(pattern);
        if (match) {
            result.startDate = normalizeDate(match[1]);
            result.endDate = normalizeDate(match[2]);
            break;
        }
    }

    // 사용량 추출 (kWh, ㎥, m³ 등)
    const usagePatterns = [
        // 당월 사용량 (최우선)
        /당월\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)\s*(?:kWh|kwh)/i,
        /당월\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)\s*(?:㎥|m³|m3)/i,
        /당월\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)/i,

        // 이번달 사용량
        /이번달\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)\s*(?:kWh|kwh)/i,
        /이번달\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)\s*(?:㎥|m³|m3)/i,
        /이번달\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)/i,

        // 사용량 (단위 포함)
        /사용량\s*[:\s-]*([0-9,]+\.?\d*)\s*(?:kWh|kwh)/i,
        /사용량\s*[:\s-]*([0-9,]+\.?\d*)\s*(?:㎥|m³|m3)/i,

        // 전기/가스/수도 사용량
        /전기\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)/i,
        /가스\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)/i,
        /수도\s*사용량\s*[:\s-]*([0-9,]+\.?\d*)/i,

        // 기타
        /(?:전기|가스|수도)\s*사용\s*[:\s-]*([0-9,]+\.?\d*)/i,
        /사용량\s*[:\s-]*([0-9,]+\.?\d*)/i
    ];

    for (const pattern of usagePatterns) {
        const match = text.match(pattern);
        if (match) {
            result.usageAmount = match[1].replace(/,/g, '');
            break;
        }
    }

    // 금액 추출 (우선순위 순서: 당월요금 > 청구금액 > 납부금액 > 합계)
    const costPatterns = [
        // 당월요금 관련 (최우선)
        /당월\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i,
        /당월\s*사용\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i,
        /당월\s*전기\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i,
        /이번달\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i,
        /이번\s*달\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i,

        // 청구금액 관련
        /청구\s*금액\s*[:\s-]*([0-9,]+)\s*원?/i,
        /청구\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i,
        /청구액\s*[:\s-]*([0-9,]+)\s*원?/i,

        // 납부금액 관련
        /납부\s*금액\s*[:\s-]*([0-9,]+)\s*원?/i,
        /납부\s*할\s*금액\s*[:\s-]*([0-9,]+)\s*원?/i,
        /납부액\s*[:\s-]*([0-9,]+)\s*원?/i,

        // 합계 관련
        /합계\s*금액?\s*[:\s-]*([0-9,]+)\s*원?/i,
        /총\s*금액\s*[:\s-]*([0-9,]+)\s*원?/i,
        /요금\s*합계\s*[:\s-]*([0-9,]+)\s*원?/i,
        /총\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i,

        // 기타 (마지막 우선순위)
        /사용\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i,
        /전기\s*요금\s*[:\s-]*([0-9,]+)\s*원?/i
    ];

    for (const pattern of costPatterns) {
        const match = text.match(pattern);
        if (match) {
            result.usageCost = match[1].replace(/,/g, '');
            break;
        }
    }

    return result;
}

// 날짜 형식 정규화 (YYYY-MM-DD)
function normalizeDate(dateStr) {
    // 공백, 년, 월, 일 제거 후 구분자를 -로 통일
    let normalized = dateStr
        .replace(/년|월|일/g, '')
        .replace(/\s+/g, '')
        .replace(/[./]/g, '-');

    // YYYY-M-D 형태를 YYYY-MM-DD로 변환
    const parts = normalized.split('-');
    if (parts.length === 3) {
        const year = parts[0];
        const month = parts[1].padStart(2, '0');
        const day = parts[2].padStart(2, '0');
        return `${year}-${month}-${day}`;
    }

    return normalized;
}

// OCR 데이터를 폼에 자동 입력하는 함수
function fillFormWithOCRData(data) {
    console.log('=== OCR 자동 입력 시작 ===');
    console.log('추출된 데이터:', data);

    let filledFields = [];

    // 에너지 종류 자동 감지 및 입력
    if (data.energyType) {
        const energyTypeSelect = document.getElementById('energyType');
        if (energyTypeSelect) {
            energyTypeSelect.value = data.energyType;
            highlightField(energyTypeSelect);
            filledFields.push(`에너지 종류: ${data.energyType}`);
        } else {
            console.error('❌ 에너지 종류 필드를 찾을 수 없습니다');
        }
    } else {
        console.log('⚠️ 추출된 에너지 종류 없음');
    }

    // 월분
    if (data.billingMonth) {
        const billingMonthInput = document.getElementById('billingMonth');
        console.log('월분 필드:', billingMonthInput);
        if (billingMonthInput) {
            billingMonthInput.value = data.billingMonth;
            console.log(`✅ 월분 입력됨: ${data.billingMonth}`);
            highlightField(billingMonthInput);
            filledFields.push(`월분: ${data.billingMonth}`);
        } else {
            console.error('❌ 월분 필드를 찾을 수 없습니다');
        }
    } else {
        console.log('⚠️ 추출된 월분 없음');
    }

    // 사용 기간 (시작일 ~ 종료일)
    if (data.startDate && data.endDate) {
        const startDateInput = document.getElementById('startDate');
        const endDateInput = document.getElementById('endDate');
        console.log('날짜 필드:', startDateInput, endDateInput);
        if (startDateInput && endDateInput) {
            // 시작일과 종료일이 같으면 단일 날짜만 표시
            if (data.startDate === data.endDate) {
                startDateInput.value = data.startDate;
            } else {
                // 범위 형식으로 표시
                startDateInput.value = `${data.startDate} ~ ${data.endDate}`;
            }
            endDateInput.value = data.endDate;
            console.log(`✅ 사용기간 입력됨: ${startDateInput.value}`);
            highlightField(startDateInput);
            filledFields.push(`사용기간: ${data.startDate} ~ ${data.endDate}`);
        } else {
            console.error('❌ 날짜 필드를 찾을 수 없습니다');
        }
    } else if (data.startDate) {
        // 시작일만 있는 경우
        const startDateInput = document.getElementById('startDate');
        const endDateInput = document.getElementById('endDate');
        if (startDateInput && endDateInput) {
            startDateInput.value = data.startDate;
            endDateInput.value = data.startDate;
            console.log(`✅ 시작일 입력됨: ${data.startDate}`);
            highlightField(startDateInput);
            filledFields.push(`사용 시작일: ${data.startDate}`);
        }
    } else if (data.endDate) {
        // 종료일만 있는 경우
        const endDateInput = document.getElementById('endDate');
        if (endDateInput) {
            endDateInput.value = data.endDate;
            console.log(`✅ 종료일 입력됨: ${data.endDate}`);
            highlightField(endDateInput);
            filledFields.push(`사용 종료일: ${data.endDate}`);
        }
    } else {
        console.log('⚠️ 추출된 날짜 없음');
    }

    // 사용량
    if (data.usageAmount) {
        const usageAmountInput = document.getElementById('usageAmount');
        console.log('사용량 필드:', usageAmountInput);
        if (usageAmountInput) {
            usageAmountInput.value = data.usageAmount;
            console.log(`✅ 사용량 입력됨: ${data.usageAmount}`);
            highlightField(usageAmountInput);
            filledFields.push(`사용량: ${data.usageAmount} ${getUsageUnit()}`);
        } else {
            console.error('❌ 사용량 필드를 찾을 수 없습니다');
        }
    } else {
        console.log('⚠️ 추출된 사용량 없음');
    }

    // 사용 금액
    if (data.usageCost) {
        const usageCostInput = document.getElementById('usageCost');
        console.log('사용 금액 필드:', usageCostInput);
        if (usageCostInput) {
            const formattedCost = formatNumber(data.usageCost);
            usageCostInput.value = formattedCost;
            console.log(`✅ 사용 금액 입력됨: ${formattedCost}원`);
            highlightField(usageCostInput);
            filledFields.push(`사용 금액: ${formattedCost}원`);
        } else {
            console.error('❌ 사용 금액 필드를 찾을 수 없습니다');
        }
    } else {
        console.log('⚠️ 추출된 사용 금액 없음');
    }

    // 추출된 데이터 요약 표시
    console.log('=== OCR 자동 입력 완료 ===');
    if (filledFields.length > 0) {
        const summary = '✅ 입력된 필드:\n' + filledFields.join('\n');
        console.log(summary);

        // 고객번호가 추출된 경우 추가
        if (data.customerNumber) {
            console.log('고객번호:', data.customerNumber);
        }
    } else {
        console.warn('⚠️ 입력된 필드가 없습니다. OCR 추출 결과를 확인하세요.');
    }
    console.log('================================');
}

// 입력 필드 강조 효과
function highlightField(element) {
    if (!element) return;

    // 원래 배경색 저장
    const originalBackground = element.style.backgroundColor;

    // 강조 효과 적용
    element.style.backgroundColor = '#e8f5e9';
    element.style.transition = 'background-color 0.3s ease';

    // 2초 후 원래대로
    setTimeout(() => {
        element.style.backgroundColor = originalBackground;
    }, 2000);
}

// 현재 선택된 에너지 종류에 따른 단위 반환
function getUsageUnit() {
    const energyType = document.getElementById('energyType')?.value;
    switch(energyType) {
        case '전기':
            return 'kWh';
        case '상하수도':
            return 'm³';
        case '도시가스':
            return 'm³';
        case '통신':
            return '';
        default:
            return '';
    }
}

// 숫자 포맷팅 (천 단위 콤마)
function formatNumber(num) {
    return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

// 월분 데이터를 표시 형식으로 변환 (YYYY-MM -> YYYY년 MM월분)
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

// OCR 추출 데이터를 최근 입력 내역에 미리보기로 추가
function addOCRPreviewToHistory(data) {
    const tableBody = document.getElementById('energyTableBody');
    if (!tableBody) return;

    // 기존 OCR 미리보기 제거 (중복 방지)
    const existingPreview = tableBody.querySelector('tr.ocr-preview');
    if (existingPreview) {
        existingPreview.remove();
    }

    // 추출된 데이터가 있는 경우에만 미리보기 생성
    if (!data.startDate && !data.endDate && !data.usageAmount && !data.usageCost) {
        return; // 추출된 데이터가 없으면 추가하지 않음
    }

    // 사용 기간 포맷팅
    let period = '';
    if (data.startDate && data.endDate) {
        period = `${data.startDate} ~ ${data.endDate}`;
    } else if (data.startDate) {
        period = data.startDate;
    } else if (data.endDate) {
        period = data.endDate;
    }

    // 에너지 종류 (감지되지 않으면 선택된 값 사용)
    const energyType = data.energyType || document.getElementById('energyType')?.value || '-';

    // 사용량 표시 (단위 포함)
    let usageDisplay = data.usageAmount || '-';
    if (data.usageAmount) {
        const unit = getUsageUnitByType(energyType);
        usageDisplay = `${data.usageAmount} ${unit}`.trim();
    }

    // 금액 표시 (콤마 포함)
    const costDisplay = data.usageCost ? `${formatNumber(data.usageCost)}원` : '-';

    // 시설명 (선택된 값 사용)
    const facilityName = document.getElementById('facilityName')?.value || '-';

    // 미리보기 행 생성
    const row = document.createElement('tr');
    row.className = 'ocr-preview';
    row.style.backgroundColor = '#fff3cd'; // 노란색 배경으로 미리보기 강조
    row.style.borderLeft = '4px solid #ffc107';
    row.innerHTML = `
        <td>${formatBillingMonth(data.billingMonth)}</td>
        <td>${period || '-'}</td>
        <td>${facilityName}</td>
        <td>
            <span class="energy-type-badge" style="background-color: ${getEnergyTypeColor(energyType)}">
                ${energyType}
            </span>
        </td>
        <td>${usageDisplay}</td>
        <td>${costDisplay}</td>
        <td>
            <span style="color: #ff9800; font-weight: bold;">📋 OCR 미리보기</span>
            <div style="font-size: 0.85em; color: #666; margin-top: 4px;">
                저장하면 정식 등록됩니다
            </div>
        </td>
    `;

    // 테이블 맨 위에 추가
    tableBody.insertBefore(row, tableBody.firstChild);

    // 부드러운 등장 애니메이션
    row.style.opacity = '0';
    row.style.transform = 'translateY(-10px)';
    row.style.transition = 'all 0.3s ease';

    setTimeout(() => {
        row.style.opacity = '1';
        row.style.transform = 'translateY(0)';
    }, 10);
}

// 에너지 종류별 단위 반환
function getUsageUnitByType(energyType) {
    switch(energyType) {
        case '전기':
            return 'kWh';
        case '상하수도':
            return 'm³';
        case '도시가스':
            return 'm³';
        case '통신':
            return '';
        default:
            return '';
    }
}

// 에너지 종류별 색상 반환
function getEnergyTypeColor(energyType) {
    switch(energyType) {
        case '전기':
            return '#4CAF50';
        case '상하수도':
            return '#2196F3';
        case '도시가스':
            return '#FF9800';
        case '통신':
            return '#9C27B0';
        default:
            return '#999';
    }
}

// OCR 미리보기 제거
function removeOCRPreview() {
    const tableBody = document.getElementById('energyTableBody');
    if (!tableBody) return;

    const preview = tableBody.querySelector('tr.ocr-preview');
    if (preview) {
        // 페이드아웃 애니메이션
        preview.style.opacity = '0';
        preview.style.transform = 'translateX(-10px)';

        setTimeout(() => {
            preview.remove();
        }, 300);
    }
}

// ==================== 데이터 전체 삭제 기능 ====================

// 데이터 전체 삭제 버튼 이벤트
document.addEventListener('DOMContentLoaded', () => {
    const deleteAllDataBtn = document.getElementById('deleteAllDataBtn');
    
    if (deleteAllDataBtn) {
        deleteAllDataBtn.addEventListener('click', deleteAllEnergyData);
    }
});

// 모든 에너지 데이터 삭제 함수
async function deleteAllEnergyData() {
    // 첫 번째 확인
    const firstConfirm = confirm('⚠️ 경고: 모든 에너지 사용량 데이터를 삭제하시겠습니까?\n\n이 작업은 되돌릴 수 없습니다.');

    if (!firstConfirm) {
        return;
    }

    // 두 번째 확인 (안전장치)
    const secondConfirm = confirm('정말로 삭제하시겠습니까?\n\n다시 한번 확인합니다. 이 작업은 복구할 수 없습니다.');

    if (!secondConfirm) {
        return;
    }

    try {
        // 먼저 현재 데이터 개수 확인
        const checkResponse = await fetch('/api/data-view');
        const checkData = await checkResponse.json();

        if (!checkData.success) {
            throw new Error('데이터 조회 실패');
        }

        const totalRecords = checkData.records ? checkData.records.length : 0;

        if (totalRecords === 0) {
            alert('삭제할 데이터가 없습니다.');
            return;
        }

        // 최종 확인
        const finalConfirm = confirm(`총 ${totalRecords}개의 데이터를 삭제합니다.\n\n마지막 확인입니다. 계속하시겠습니까?`);

        if (!finalConfirm) {
            return;
        }

        // 모든 데이터 삭제 요청
        const response = await fetch('/api/energy-data/delete-all', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' }
        });

        const result = await response.json();

        if (result.success) {
            alert(`성공적으로 ${result.deletedCount || totalRecords}개의 데이터가 삭제되었습니다.`);

            // 전역 변수 초기화
            window.currentViewRecords = [];
            currentViewRecords = [];

            // 화면 새로고침
            if (typeof searchData === 'function') {
                searchData();
            }
        } else {
            throw new Error(result.message || '삭제 실패');
        }

    } catch (error) {
        console.error('데이터 삭제 오류:', error);
        alert('데이터 삭제 중 오류가 발생했습니다: ' + error.message);
    }
}

// 공문 생성 함수
async function generateOfficialDocument(record) {
    try {
        console.log('공문 생성:', record);

        const response = await fetch('/api/generate-document', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(record)
        });

        if (!response.ok) {
            throw new Error('공문 생성에 실패했습니다.');
        }

        // Blob으로 변환하여 다운로드
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        // 파일명 생성
        const startDate = new Date(record.startDate);
        const year = startDate.getFullYear();
        const month = String(startDate.getMonth() + 1).padStart(2, '0');
        const filename = `${year}-${month}-${record.facilityName}-${record.energyType}-공문.docx`;

        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        console.log('공문 생성 완료:', filename);

    } catch (error) {
        console.error('공문 생성 오류:', error);
        alert('공문 생성 중 오류가 발생했습니다: ' + error.message);
    }
}

// 첨부문서 생성 함수 (첨부1 - Excel)
async function generateAttachmentDocument(record) {
    try {
        console.log('첨부문서 생성 요청:', record);

        if (!record) {
            throw new Error('레코드 데이터가 없습니다.');
        }

        const response = await fetch('/api/generate-attachment1', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(record)
        });

        if (!response.ok) {
            // 오류 응답 내용 확인
            const contentType = response.headers.get('content-type');
            if (contentType && contentType.includes('application/json')) {
                const errorData = await response.json();
                throw new Error(errorData.message || '첨부문서 생성에 실패했습니다.');
            } else {
                throw new Error(`서버 오류 (${response.status})`);
            }
        }

        // Blob으로 변환하여 다운로드
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        // 파일명 생성 - billingMonth 우선 사용
        let year, month;
        if (record.billingMonth) {
            const parts = record.billingMonth.split('-');
            year = parts[0];
            month = parts[1];
        } else if (record.startDate) {
            const startDate = new Date(record.startDate);
            year = startDate.getFullYear();
            month = String(startDate.getMonth() + 1).padStart(2, '0');
        } else {
            const now = new Date();
            year = now.getFullYear();
            month = String(now.getMonth() + 1).padStart(2, '0');
        }
        const filename = `${year}-${month}-${record.facilityName}-${record.energyType}-첨부1.xlsx`;

        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        console.log('첨부문서 생성 완료:', filename);

    } catch (error) {
        console.error('첨부문서 생성 오류:', error);
        alert('첨부문서 생성 중 오류가 발생했습니다: ' + error.message);
    }
}

// 선택된 레코드 가져오기 함수
function getSelectedRecords() {
    const checkboxes = document.querySelectorAll('.row-checkbox:checked');
    const selectedRecords = [];

    // renderDataView에서 정렬된 배열을 그대로 사용 (data-index와 동일한 순서)
    const sortedRecords = window.sortedViewRecords || [];

    checkboxes.forEach(cb => {
        const index = parseInt(cb.getAttribute('data-index'));
        if (sortedRecords[index]) {
            selectedRecords.push(sortedRecords[index]);
        }
    });

    return selectedRecords;
}

// 선택된 데이터의 합계 금액으로 공문 생성 함수
async function generateOfficialDocumentCombined(records) {
    try {
        console.log('선택된 데이터 합계 공문 생성:', records);

        if (!records || records.length === 0) {
            alert('선택된 데이터가 없습니다.');
            return;
        }

        // 합계 금액 계산
        const totalCost = records.reduce((sum, record) => {
            return sum + (parseFloat(record.usageCost) || 0);
        }, 0);

        // 시설명 목록 (중복 제거)
        const facilities = [...new Set(records.map(r => r.facilityName))];

        // 에너지 종류 목록 (중복 제거)
        const energyTypes = [...new Set(records.map(r => r.energyType))];

        // 기간 범위 계산
        const dates = records.map(r => r.startDate || r.usageDate).filter(d => d).sort();
        const startDate = dates[0] || '';
        const endDate = dates[dates.length - 1] || startDate;

        // 확인 메시지
        const confirmMsg = `선택된 ${records.length}건의 데이터로 공문을 생성합니다.\n\n` +
            `시설: ${facilities.join(', ')}\n` +
            `에너지 종류: ${energyTypes.join(', ')}\n` +
            `합계 금액: ${totalCost.toLocaleString('ko-KR')}원\n\n` +
            `계속하시겠습니까?`;

        if (!confirm(confirmMsg)) {
            return;
        }

        // 합계 데이터로 공문 생성 요청
        const combinedRecord = {
            facilityName: facilities.length === 1 ? facilities[0] : facilities.join(', '),
            energyType: energyTypes.length === 1 ? energyTypes[0] : '통합',
            usageCost: totalCost,
            usageAmount: records.reduce((sum, r) => sum + (parseFloat(r.usageAmount) || 0), 0),
            startDate: startDate,
            endDate: endDate,
            billingMonth: records[0].billingMonth || '',
            records: records, // 상세 레코드 목록 포함
            isCombined: true // 합계 공문임을 표시
        };

        const response = await fetch('/api/generate-document', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(combinedRecord)
        });

        if (!response.ok) {
            throw new Error('공문 생성에 실패했습니다.');
        }

        // Blob으로 변환하여 다운로드
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        // 파일명 생성
        const now = new Date();
        const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        const facilityStr = facilities.length === 1 ? facilities[0] : '통합';
        const filename = `${dateStr}-${facilityStr}-합계공문.docx`;

        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        console.log('합계 공문 생성 완료:', filename);
        alert(`공문 생성이 완료되었습니다.\n\n파일명: ${filename}\n합계 금액: ${totalCost.toLocaleString('ko-KR')}원`);

    } catch (error) {
        console.error('합계 공문 생성 오류:', error);
        alert('공문 생성 중 오류가 발생했습니다: ' + error.message);
    }
}

// 선택된 데이터를 하나의 엑셀 파일로 생성하는 함수
async function generateAttachmentDocumentCombined(records) {
    try {
        console.log('선택된 데이터 통합 첨부문서 생성:', records);

        if (!records || records.length === 0) {
            alert('선택된 데이터가 없습니다.');
            return;
        }

        // 시설명 목록 (중복 제거)
        const facilities = [...new Set(records.map(r => r.facilityName))];

        // 합계 금액 계산
        const totalCost = records.reduce((sum, record) => {
            return sum + (parseFloat(record.usageCost) || 0);
        }, 0);

        // 확인 메시지
        const confirmMsg = `선택된 ${records.length}건의 데이터를 하나의 엑셀 파일로 생성합니다.\n\n` +
            `시설: ${facilities.join(', ')}\n` +
            `합계 금액: ${totalCost.toLocaleString('ko-KR')}원\n\n` +
            `계속하시겠습니까?`;

        if (!confirm(confirmMsg)) {
            return;
        }

        // 통합 첨부문서 생성 요청
        const response = await fetch('/api/generate-attachment-combined', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ records: records })
        });

        if (!response.ok) {
            throw new Error('첨부문서 생성에 실패했습니다.');
        }

        // Blob으로 변환하여 다운로드
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        // 파일명 생성
        const now = new Date();
        const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        const facilityStr = facilities.length === 1 ? facilities[0] : '통합';
        const filename = `${dateStr}-${facilityStr}-첨부문서.xlsx`;

        a.download = filename;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        console.log('통합 첨부문서 생성 완료:', filename);
        alert(`첨부문서 생성이 완료되었습니다.\n\n파일명: ${filename}\n데이터 건수: ${records.length}건`);

    } catch (error) {
        console.error('통합 첨부문서 생성 오류:', error);
        alert('첨부문서 생성 중 오류가 발생했습니다: ' + error.message);
    }
}

// 전체 데이터 기반 공문 일괄 생성 함수
async function generateOfficialDocumentBulk(records) {
    try {
        console.log('전체 공문 생성:', records);

        if (!records || records.length === 0) {
            alert('생성할 데이터가 없습니다.');
            return;
        }

        // 각 레코드에 대해 개별 공문 생성
        const confirmMsg = `총 ${records.length}건의 공문을 생성하시겠습니까?`;
        if (!confirm(confirmMsg)) {
            return;
        }

        let successCount = 0;
        let failCount = 0;

        for (const record of records) {
            try {
                await generateOfficialDocument(record);
                successCount++;
                // 다운로드 간 짧은 지연
                await new Promise(resolve => setTimeout(resolve, 500));
            } catch (error) {
                console.error(`공문 생성 실패 (${record.facilityName} - ${record.energyType}):`, error);
                failCount++;
            }
        }

        alert(`공문 생성 완료\n성공: ${successCount}건\n실패: ${failCount}건`);

    } catch (error) {
        console.error('전체 공문 생성 오류:', error);
        alert('공문 생성 중 오류가 발생했습니다: ' + error.message);
    }
}

// 전체 데이터 기반 첨부문서 일괄 생성 함수
async function generateAttachmentDocumentBulk(records) {
    try {
        console.log('전체 첨부문서 생성:', records);

        if (!records || records.length === 0) {
            alert('생성할 데이터가 없습니다.');
            return;
        }

        // 각 레코드에 대해 개별 첨부문서 생성
        const confirmMsg = `총 ${records.length}건의 첨부문서를 생성하시겠습니까?`;
        if (!confirm(confirmMsg)) {
            return;
        }

        let successCount = 0;
        let failCount = 0;

        for (const record of records) {
            try {
                await generateAttachmentDocument(record);
                successCount++;
                // 다운로드 간 짧은 지연
                await new Promise(resolve => setTimeout(resolve, 500));
            } catch (error) {
                console.error(`첨부문서 생성 실패 (${record.facilityName} - ${record.energyType}):`, error);
                failCount++;
            }
        }

        alert(`첨부문서 생성 완료\n성공: ${successCount}건\n실패: ${failCount}건`);

    } catch (error) {
        console.error('전체 첨부문서 생성 오류:', error);
        alert('첨부문서 생성 중 오류가 발생했습니다: ' + error.message);
    }
}

// ==================== 자동 로그아웃 기능 ====================

// 페이지 종료(브라우저 탭/창 닫기) 시 자동 로그아웃
window.addEventListener('beforeunload', function(e) {
    // 로그인 상태 확인
    const currentUser = sessionStorage.getItem('currentUser');

    if (currentUser) {
        // 세션 스토리지 삭제
        sessionStorage.removeItem('currentUser');

        // 서버에 로그아웃 요청 (비동기, 응답 대기 안 함)
        navigator.sendBeacon('/api/logout');
    }
});

// 페이지 이탈 시 자동 로그아웃 (뒤로가기, 다른 페이지 이동 등)
window.addEventListener('pagehide', function(e) {
    const currentUser = sessionStorage.getItem('currentUser');

    if (currentUser) {
        sessionStorage.removeItem('currentUser');
        navigator.sendBeacon('/api/logout');
    }
});

// 브라우저 숨김/보임 감지 (탭 전환)
document.addEventListener('visibilitychange', function() {
    if (document.visibilityState === 'hidden') {
        // 탭이 숨겨졌을 때 세션 타임스탬프 저장
        const currentUser = sessionStorage.getItem('currentUser');
        if (currentUser) {
            sessionStorage.setItem('lastActiveTime', Date.now().toString());
        }
    } else if (document.visibilityState === 'visible') {
        // 탭이 다시 보일 때 세션 타임아웃 확인 (30분)
        const lastActiveTime = sessionStorage.getItem('lastActiveTime');
        const currentUser = sessionStorage.getItem('currentUser');

        if (lastActiveTime && currentUser) {
            const inactiveTime = Date.now() - parseInt(lastActiveTime);
            const timeoutDuration = 30 * 60 * 1000; // 30분

            if (inactiveTime > timeoutDuration) {
                // 30분 이상 비활성 상태였으면 자동 로그아웃
                sessionStorage.removeItem('currentUser');
                sessionStorage.removeItem('lastActiveTime');
                navigator.sendBeacon('/api/logout');
                showLogin();
                alert('장시간 사용하지 않아 자동으로 로그아웃되었습니다.');
            }
        }
    }
});

// 에너지 데이터 엑셀 다운로드 (energy_data.xlsx 서식)
function downloadEnergyDataExcel() {
    const records = window.currentViewRecords || [];

    if (records.length === 0) {
        alert('다운로드할 데이터가 없습니다.\n먼저 데이터를 조회해주세요.');
        return;
    }

    // 엑셀 워크북 생성
    const workbook = XLSX.utils.book_new();

    // 헤더 및 제목 행 (energy_data.xlsx 서식)
    const data = [
        ['', '에너지 사용 내역'],
        [],
        ['', '', '', '', '', '', '', '', '', '', '', '', '', '(단위 : 원)'],
        ['년도', '월', '고객번호(뒷자리)', '에너지종류', '사용시설', '금융기관', '가상계좌', '고객번호', '사용기간', '', '', '사용량', '사용금액', '비고']
    ];

    // 데이터 행 추가
    records.forEach(record => {
        const year = record.billingMonth ? record.billingMonth.split('-')[0] + '년' : '';
        const month = record.billingMonth ? parseInt(record.billingMonth.split('-')[1]) + '월' : '';
        const energyType = record.energyType ? record.energyType + '료' : '';

        // 고객번호 뒷자리 추출 (뒤 4자리)
        const customerNumberLast = record.customerNumber ? record.customerNumber.slice(-4) : '';

        data.push([
            year,
            month,
            customerNumberLast,
            energyType,
            record.facilityName || '',
            record.bankName || '',
            record.virtualAccount || '',
            record.customerNumber || '',
            record.startDate || '',
            '~',
            record.endDate || '',
            record.usageAmount || '',
            record.usageCost || 0,
            ''
        ]);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(data);

    // 열 너비 설정
    worksheet['!cols'] = [
        { wch: 8 },   // 년도
        { wch: 6 },   // 월
        { wch: 12 },  // 고객번호(뒷자리)
        { wch: 10 },  // 에너지종류
        { wch: 15 },  // 사용시설
        { wch: 12 },  // 금융기관
        { wch: 15 },  // 가상계좌
        { wch: 15 },  // 고객번호
        { wch: 12 },  // 사용기간 시작
        { wch: 3 },   // ~
        { wch: 12 },  // 사용기간 종료
        { wch: 12 },  // 사용량
        { wch: 15 },  // 사용금액
        { wch: 10 }   // 비고
    ];

    XLSX.utils.book_append_sheet(workbook, worksheet, '에너지 사용 내역');

    // 파일 다운로드
    const today = new Date();
    const dateStr = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;
    XLSX.writeFile(workbook, `에너지사용내역_${dateStr}.xlsx`);
}

// 에너지 데이터 엑셀 업로드 (energy_data.xlsx 서식)
async function uploadEnergyDataExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    // 파일 입력 초기화
    event.target.value = '';

    try {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        // 헤더 행 찾기 (년도, 월, ... 로 시작하는 행)
        let headerIndex = -1;
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (row && row[0] === '년도' && row[1] === '월') {
                headerIndex = i;
                break;
            }
        }

        if (headerIndex === -1) {
            alert('올바른 서식의 엑셀 파일이 아닙니다.\n년도, 월 헤더가 있는 서식을 사용해주세요.');
            return;
        }

        const records = [];

        // 헤더 다음 행부터 데이터 파싱
        for (let i = headerIndex + 1; i < rows.length; i++) {
            const row = rows[i];
            if (!row || row.length < 8) continue;

            // 빈 행 건너뛰기
            if (!row[0] && !row[1] && !row[3]) continue;

            const year = String(row[0] || '').replace('년', '').trim();
            const month = String(row[1] || '').replace('월', '').trim();
            const energyType = String(row[3] || '').replace('료', '').trim();
            const facilityName = String(row[4] || '').trim();
            const bankName = String(row[5] || '').trim();
            const virtualAccount = String(row[6] || '').trim();
            const customerNumber = String(row[7] || '').trim();
            let startDate = row[8];
            let endDate = row[10];
            const usageAmount = row[11] || 0;
            const usageCost = row[12] || 0;

            if (!year || !month || !energyType || !facilityName) continue;

            // billingMonth 형식: YYYY-MM
            const billingMonth = `${year}-${String(month).padStart(2, '0')}`;

            // 날짜 포맷 변환 (Excel 숫자 → YYYY-MM-DD)
            if (typeof startDate === 'number') {
                const excelDate = new Date((startDate - 25569) * 86400 * 1000);
                startDate = excelDate.toISOString().split('T')[0];
            } else {
                startDate = String(startDate || '');
            }

            if (typeof endDate === 'number') {
                const excelDate = new Date((endDate - 25569) * 86400 * 1000);
                endDate = excelDate.toISOString().split('T')[0];
            } else {
                endDate = String(endDate || '');
            }

            records.push({
                facilityName,
                billingMonth,
                startDate,
                endDate,
                energyType,
                usageAmount: parseFloat(usageAmount) || 0,
                usageCost: parseFloat(usageCost) || 0,
                customerNumber,
                bankName,
                virtualAccount
            });
        }

        if (records.length === 0) {
            alert('업로드할 데이터가 없습니다.');
            return;
        }

        // 서버에 데이터 업로드
        const response = await fetch('/api/energy-data/upload', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ records })
        });

        const result = await response.json();

        if (result.success) {
            alert(`${records.length}건의 데이터가 업로드되었습니다.`);
            searchData(); // 데이터 새로고침
        } else {
            alert('데이터 업로드 실패: ' + (result.message || '알 수 없는 오류'));
        }
    } catch (error) {
        console.error('엑셀 업로드 오류:', error);
        alert('엑셀 파일 처리 중 오류가 발생했습니다.');
    }
}
