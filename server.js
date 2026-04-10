const express = require('express');
const session = require('express-session');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;
const EXCEL_FILE = path.join(__dirname, 'front end', 'id, pw.xlsx');
const ENERGY_FILE = path.join(__dirname, 'back end', '에너지사용 data.xlsx');
const ENERGY_INFO_FILE = path.join(__dirname, 'back end', 'energy_info.xlsx');
const ATTACHMENT_TEMPLATE_FILE = path.join(__dirname, 'public', '첨부 2 서식.xlsx');

// Body parser 설정 - 이미지 업로드를 위해 크기 제한 증가
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static('public'));

app.use(session({
  secret: 'energy-management-secret-key',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 }
}));

function loadUsers() {
  try {
    const workbook = xlsx.readFile(EXCEL_FILE);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    console.log('엑셀 데이터 로드:', data);

    const users = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row && row.length >= 3 && row[1] && row[2]) {
        users.push({
          index: i - 1,
          facilityName: String(row[0] || '').trim(),
          id: String(row[1]).trim(),
          password: String(row[2]).trim(),
          role: row[3] ? String(row[3]).trim() : '시설담당자', // 기본값: 시설담당자
          parentFacility: row[4] ? String(row[4]).trim() : '' // 상위시설명 (시설관리자가 관리하는 하위 시설용)
        });
      }
    }

    console.log('로드된 사용자:', users);
    return users;
  } catch (error) {
    console.error('엑셀 파일 로드 오류:', error);
    return [];
  }
}

function saveUsers(users) {
  try {
    // 헤더를 명시적으로 정의 (올바른 형식으로)
    const newData = [['시설명', 'id', 'pw', '역할', '상위시설명']];

    users.forEach(user => {
      newData.push([
        user.facilityName,
        user.id,
        user.password,
        user.role || '시설담당자',
        user.parentFacility || ''
      ]);
    });

    const workbook = xlsx.utils.book_new();
    const newWorksheet = xlsx.utils.aoa_to_sheet(newData);
    xlsx.utils.book_append_sheet(workbook, newWorksheet, 'Sheet1');

    xlsx.writeFile(workbook, EXCEL_FILE);
    console.log('엑셀 파일 저장 완료');
    return true;
  } catch (error) {
    console.error('엑셀 파일 저장 오류:', error);
    return false;
  }
}

function loadEnergyData() {
  try {
    if (!fs.existsSync(ENERGY_FILE)) {
      const workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.aoa_to_sheet([['년도', '월', '고객번호(뒷자리)', '에너지종류', '사용시설', '가상계좌', '고객번호', '사용기간', '', '', '사용량', '사용금액', '비고']]);
      xlsx.utils.book_append_sheet(workbook, worksheet, '에너지 사용 내역');
      xlsx.writeFile(workbook, ENERGY_FILE);
      return [];
    }

    const workbook = xlsx.readFile(ENERGY_FILE);
    const sheetName = workbook.SheetNames[0]; // 첫 번째 시트 사용
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    const records = [];

    // 헤더 행 찾기 (년도, 월, ... 로 시작하는 행)
    let headerIndex = -1;
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row && row[0] === '년도' && row[1] === '월') {
        headerIndex = i;
        break;
      }
    }

    if (headerIndex === -1) {
      headerIndex = 3; // 기본값
    }

    // 헤더 다음 행부터 데이터 파싱
    for (let i = headerIndex + 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length < 8) continue;

      // 빈 행 건너뛰기
      if (!row[0] && !row[1] && !row[3]) continue;

      const year = String(row[0] || '').replace('년', '').trim();
      const month = String(row[1] || '').replace('월', '').trim();
      const energyType = String(row[3] || '').replace('료', '').trim();
      const facilityName = String(row[4] || '').trim();
      const bankName = String(row[5] || '').trim(); // 금융기관 (F열)
      const virtualAccount = String(row[6] || '').trim(); // 가상계좌 (G열)
      const customerNumber = String(row[7] || '').trim(); // 고객번호 (H열)
      const startDate = row[8]; // 사용기간 시작 (I열)
      const endDate = row[10]; // 사용기간 종료 (K열, J열은 "~")
      const usageAmount = row[11] || 0; // 사용량 (L열)
      const usageCost = row[12] || 0; // 사용금액 (M열)

      if (year && month && energyType && facilityName) {
        // billingMonth 형식: YYYY-MM
        const billingMonth = year && month ? `${year}-${String(month).padStart(2, '0')}` : '';

        // 날짜 포맷 변환 (숫자 → YYYY-MM-DD)
        let formattedStartDate = '';
        let formattedEndDate = '';

        if (typeof startDate === 'number') {
          const excelDate = new Date((startDate - 25569) * 86400 * 1000);
          formattedStartDate = excelDate.toISOString().split('T')[0];
        } else if (startDate) {
          formattedStartDate = String(startDate);
        }

        if (typeof endDate === 'number') {
          const excelDate = new Date((endDate - 25569) * 86400 * 1000);
          formattedEndDate = excelDate.toISOString().split('T')[0];
        } else if (endDate) {
          formattedEndDate = String(endDate);
        }

        const record = {
          facilityName: facilityName,
          billingMonth: billingMonth,
          startDate: formattedStartDate,
          endDate: formattedEndDate,
          energyType: energyType,
          usageAmount: typeof usageAmount === 'number' ? usageAmount : 0,
          usageCost: typeof usageCost === 'number' ? usageCost : 0,
          bankName: bankName,
          virtualAccount: virtualAccount,
          customerNumber: customerNumber
        };

        records.push(record);
      }
    }

    return records;
  } catch (error) {
    console.error('에너지 데이터 로드 오류:', error);
    return [];
  }
}

function saveEnergyData(records) {
  try {
    const workbook = xlsx.utils.book_new();

    // 헤더 및 제목 행
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
      const energyType = record.energyType + (record.energyType === '통신' ? '료' : '료');

      // 날짜를 Excel 숫자 형식으로 변환
      let startDateNum = '';
      let endDateNum = '';

      if (record.startDate) {
        const startDate = new Date(record.startDate);
        if (!isNaN(startDate.getTime())) {
          startDateNum = Math.floor((startDate.getTime() / 86400000) + 25569);
        }
      }

      if (record.endDate) {
        const endDate = new Date(record.endDate);
        if (!isNaN(endDate.getTime())) {
          endDateNum = Math.floor((endDate.getTime() / 86400000) + 25569);
        }
      }

      data.push([
        year,
        month,
        '', // 고객번호 뒷자리
        energyType,
        record.facilityName || '',
        record.bankName || '', // 금융기관
        record.virtualAccount || '', // 가상계좌
        record.customerNumber || '', // 고객번호
        startDateNum,
        '~',
        endDateNum,
        record.usageAmount || '',
        record.usageCost || 0,
        '' // 비고
      ]);
    });

    const worksheet = xlsx.utils.aoa_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, worksheet, '에너지 사용 내역');
    xlsx.writeFile(workbook, ENERGY_FILE);
    console.log('에너지 데이터 저장 완료');
    return true;
  } catch (error) {
    console.error('에너지 데이터 저장 오류:', error);
    return false;
  }
}

function loadEnergyInfo() {
  try {
    if (!fs.existsSync(ENERGY_INFO_FILE)) {
      const workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.aoa_to_sheet([['시설명', '에너지종류', '고객번호', '금융기관', '계좌번호']]);
      xlsx.utils.book_append_sheet(workbook, worksheet, 'EnergyInfo');
      xlsx.writeFile(workbook, ENERGY_INFO_FILE);
      return [];
    }

    const workbook = xlsx.readFile(ENERGY_INFO_FILE);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    const infos = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row && row.length >= 5) {
        infos.push({
          facilityName: String(row[0] || ''),
          energyType: String(row[1] || ''),
          customerNumber: String(row[2] || ''),
          bankName: String(row[3] || ''),
          accountNumber: String(row[4] || '')
        });
      }
    }

    return infos;
  } catch (error) {
    console.error('에너지 정보 로드 오류:', error);
    return [];
  }
}

function saveEnergyInfo(infos) {
  try {
    const workbook = xlsx.utils.book_new();
    const data = [['시설명', '에너지종류', '고객번호', '금융기관', '계좌번호']];

    infos.forEach(info => {
      data.push([
        info.facilityName || '',
        info.energyType || '',
        info.customerNumber || '',
        info.bankName || '',
        info.accountNumber || ''
      ]);
    });

    const worksheet = xlsx.utils.aoa_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'EnergyInfo');
    xlsx.writeFile(workbook, ENERGY_INFO_FILE);
    console.log('에너지 정보 저장 완료');
    return true;
  } catch (error) {
    console.error('에너지 정보 저장 오류:', error);
    return false;
  }
}

app.post('/api/login', (req, res) => {
  const { facilityName, username, password } = req.body;
  const users = loadUsers();

  console.log('로그인 시도:', { facilityName, username, password });

  const user = users.find(u => {
    const match = u.facilityName === facilityName && u.id === username && u.password === password;
    console.log(`비교: ${u.facilityName} === ${facilityName} && ${u.id} === ${username} && ${u.password} === ${password} => ${match}`);
    return match;
  });

  if (user) {
    req.session.user = {
      id: user.id,
      facilityName: user.facilityName,
      role: user.role || '시설담당자'
    };
    console.log('로그인 성공:', req.session.user);
    res.json({ success: true, user: req.session.user });
  } else {
    console.log('로그인 실패');
    res.json({ success: false, message: '시설명, 아이디 또는 비밀번호가 올바르지 않습니다.' });
  }
});

app.post('/api/logout', (req, res) => {
  req.session.destroy();
  res.json({ success: true });
});

app.get('/api/check-auth', (req, res) => {
  if (req.session.user) {
    res.json({ authenticated: true, user: req.session.user });
  } else {
    res.json({ authenticated: false });
  }
});

app.get('/api/facilities', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const users = loadUsers();
  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  // 관리자: 모든 시설 반환 (원본 인덱스 포함)
  if (userRole === '관리자') {
    const facilitiesWithIndex = users.map((user, idx) => ({
      ...user,
      originalIndex: idx
    }));
    return res.json({ success: true, facilities: facilitiesWithIndex });
  }

  // 시설관리자: 본인 시설 + 하위 시설담당자 반환 (원본 인덱스 포함)
  if (userRole === '시설관리자') {
    const managedFacilities = users
      .map((u, idx) => ({ ...u, originalIndex: idx }))
      .filter(u =>
        u.facilityName === userFacilityName ||
        (u.parentFacility === userFacilityName && u.role === '시설담당자')
      );

    return res.json({ success: true, facilities: managedFacilities });
  }

  // 시설담당자: 본인 시설만 반환 (원본 인덱스 포함)
  const userFacilityWithIndex = users
    .map((u, idx) => ({ ...u, originalIndex: idx }))
    .find(u => u.facilityName === userFacilityName);
  res.json({ success: true, facilities: userFacilityWithIndex ? [userFacilityWithIndex] : [] });
});

app.post('/api/facilities', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  // 관리자 또는 시설관리자만 시설 추가 가능
  if (userRole !== '관리자' && userRole !== '시설관리자') {
    return res.status(403).json({ success: false, message: '권한이 없습니다.' });
  }

  const { facilityName, id, password, role, parentFacility } = req.body;

  if (!facilityName || !id || !password || !role) {
    return res.json({ success: false, message: '모든 필드를 입력해주세요.' });
  }

  const users = loadUsers();

  const exists = users.find(u => u.facilityName === facilityName && u.id === id);
  if (exists) {
    return res.json({ success: false, message: '이미 존재하는 시설명 또는 아이디입니다.' });
  }

  // 시설관리자는 시설담당자만 추가 가능하고, parentFacility는 자동으로 본인 시설로 설정
  if (userRole === '시설관리자') {
    if (role !== '시설담당자') {
      return res.status(403).json({ success: false, message: '시설관리자는 시설담당자만 추가할 수 있습니다.' });
    }
    users.push({ facilityName, id, password, role, parentFacility: userFacilityName });
  } else {
    // 관리자는 모든 역할 추가 가능
    users.push({ facilityName, id, password, role, parentFacility: parentFacility || '' });
  }

  if (saveUsers(users)) {
    res.json({ success: true, message: '시설이 추가되었습니다.' });
  } else {
    res.json({ success: false, message: '시설 추가 중 오류가 발생했습니다.' });
  }
});

app.put('/api/facilities/:index', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const index = parseInt(req.params.index);
  const { facilityName, id, password, role, parentFacility } = req.body;

  if (!facilityName || !id || !password || !role) {
    return res.json({ success: false, message: '모든 필드를 입력해주세요.' });
  }

  const users = loadUsers();

  if (index < 0 || index >= users.length) {
    return res.json({ success: false, message: '유효하지 않은 인덱스입니다.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;
  const targetUser = users[index];

  // 관리자: 모든 시설 수정 가능
  if (userRole === '관리자') {
    users[index] = { facilityName, id, password, role, parentFacility: parentFacility || '' };
    if (saveUsers(users)) {
      return res.json({ success: true, message: '시설 정보가 수정되었습니다.' });
    } else {
      return res.json({ success: false, message: '시설 수정 중 오류가 발생했습니다.' });
    }
  }

  // 시설관리자: 본인 시설 또는 하위 시설담당자 수정 가능
  if (userRole === '시설관리자') {
    // 본인 시설 수정
    if (targetUser.facilityName === userFacilityName) {
      users[index] = { facilityName, id, password, role, parentFacility: parentFacility || '' };
      if (saveUsers(users)) {
        return res.json({ success: true, message: '시설 정보가 수정되었습니다.' });
      } else {
        return res.json({ success: false, message: '시설 수정 중 오류가 발생했습니다.' });
      }
    }
    // 하위 시설담당자 수정
    else if (targetUser.parentFacility === userFacilityName && targetUser.role === '시설담당자') {
      if (role !== '시설담당자') {
        return res.status(403).json({ success: false, message: '시설관리자는 시설담당자만 수정할 수 있습니다.' });
      }
      users[index] = { facilityName, id, password, role, parentFacility: userFacilityName };
      if (saveUsers(users)) {
        return res.json({ success: true, message: '시설 정보가 수정되었습니다.' });
      } else {
        return res.json({ success: false, message: '시설 수정 중 오류가 발생했습니다.' });
      }
    } else {
      return res.status(403).json({ success: false, message: '권한이 없습니다.' });
    }
  }

  // 시설담당자: 수정 권한 없음
  return res.status(403).json({ success: false, message: '권한이 없습니다.' });
});

app.delete('/api/facilities/:index', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  // 관리자 또는 시설관리자만 시설 삭제 가능
  if (userRole !== '관리자' && userRole !== '시설관리자') {
    return res.status(403).json({ success: false, message: '권한이 없습니다.' });
  }

  const index = parseInt(req.params.index);
  const users = loadUsers();

  if (index < 0 || index >= users.length) {
    return res.json({ success: false, message: '유효하지 않은 인덱스입니다.' });
  }

  const targetUser = users[index];

  // 시설관리자는 본인 시설 또는 하위 시설담당자 삭제 가능
  if (userRole === '시설관리자') {
    const canDelete = targetUser.facilityName === userFacilityName ||
                      (targetUser.parentFacility === userFacilityName && targetUser.role === '시설담당자');
    if (!canDelete) {
      return res.status(403).json({ success: false, message: '권한이 없습니다.' });
    }
  }

  users.splice(index, 1);

  if (saveUsers(users)) {
    res.json({ success: true, message: '시설이 삭제되었습니다.' });
  } else {
    res.json({ success: false, message: '시설 삭제 중 오류가 발생했습니다.' });
  }
});

app.get('/api/energy', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  let records = loadEnergyData();
  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  // 관리자: 모든 데이터 조회 가능
  if (userRole === '관리자') {
    return res.json({ success: true, records: records });
  }

  // 시설관리자: 본인 시설 + 하위 시설 데이터 조회 가능
  if (userRole === '시설관리자') {
    const users = loadUsers();
    const managedFacilities = users
      .filter(u =>
        u.facilityName === userFacilityName ||
        (u.parentFacility === userFacilityName && u.role === '시설담당자')
      )
      .map(u => u.facilityName);

    records = records.filter(record => managedFacilities.includes(record.facilityName));
    return res.json({ success: true, records: records });
  }

  // 시설담당자: 본인 시설 데이터만 조회 가능
  records = records.filter(record => record.facilityName === userFacilityName);
  res.json({ success: true, records: records });
});

app.post('/api/energy', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const { facilityName, energyType, billingMonth, startDate, endDate, usageAmount, usageCost } = req.body;

  // 권한 확인
  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  if (userRole === '관리자') {
    // 관리자: 전 시설 데이터 입력 가능
  } else if (userRole === '시설관리자') {
    // 시설관리자: 본인 시설 + 하위 시설 데이터 입력 가능
    const users = loadUsers();
    const managedFacilities = users
      .filter(u =>
        u.facilityName === userFacilityName ||
        (u.parentFacility === userFacilityName && u.role === '시설담당자')
      )
      .map(u => u.facilityName);

    if (!managedFacilities.includes(facilityName)) {
      return res.status(403).json({ success: false, message: '해당 시설에 데이터를 입력할 권한이 없습니다.' });
    }
  } else if (userRole === '시설담당자') {
    // 시설담당자: 본인 시설 데이터만 입력 가능
    if (facilityName !== userFacilityName) {
      return res.status(403).json({ success: false, message: '해당 시설에 데이터를 입력할 권한이 없습니다.' });
    }
  }

  console.log('에너지 데이터 입력 요청:', {
    facilityName,
    energyType,
    billingMonth,
    startDate,
    endDate,
    usageAmount,
    usageCost,
    types: {
      facilityName: typeof facilityName,
      energyType: typeof energyType,
      billingMonth: typeof billingMonth,
      usageAmount: typeof usageAmount,
      usageCost: typeof usageCost
    }
  });

  // 필드 존재 여부 확인
  if (!facilityName || !energyType || !startDate || !endDate) {
    console.log('필수 문자열 필드 누락');
    return res.json({ success: false, message: '모든 필드를 입력해주세요.' });
  }

  // 숫자 필드 확인 (0도 유효한 값으로 처리, NaN은 거부)
  if (usageAmount === undefined || usageAmount === null || usageAmount === '' ||
      usageCost === undefined || usageCost === null || usageCost === '') {
    console.log('숫자 필드가 비어있음');
    return res.json({ success: false, message: '모든 필드를 입력해주세요.' });
  }

  // 숫자 유효성 검사
  const parsedUsageAmount = typeof usageAmount === 'number' ? usageAmount : parseFloat(usageAmount);
  const parsedUsageCost = typeof usageCost === 'number' ? usageCost : parseFloat(usageCost);

  if (isNaN(parsedUsageAmount) || isNaN(parsedUsageCost)) {
    console.log('숫자 변환 실패:', { parsedUsageAmount, parsedUsageCost });
    return res.json({ success: false, message: '사용량과 사용 금액은 숫자여야 합니다.' });
  }

  if (new Date(startDate) > new Date(endDate)) {
    return res.json({ success: false, message: '종료일은 시작일 이후여야 합니다.' });
  }

  const records = loadEnergyData();
  records.push({
    facilityName,
    energyType,
    billingMonth,
    startDate,
    endDate,
    usageAmount: parsedUsageAmount,
    usageCost: parsedUsageCost
  });

  if (saveEnergyData(records)) {
    console.log('에너지 데이터 저장 완료');
    res.json({ success: true, message: '에너지 사용량이 저장되었습니다.' });
  } else {
    console.log('에너지 데이터 저장 실패');
    res.json({ success: false, message: '에너지 데이터 저장 중 오류가 발생했습니다.' });
  }
});

app.delete('/api/energy/:index', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const index = parseInt(req.params.index);
  const records = loadEnergyData();

  if (index < 0 || index >= records.length) {
    return res.json({ success: false, message: '유효하지 않은 인덱스입니다.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;
  const targetRecord = records[index];

  // 권한 확인
  if (userRole === '관리자') {
    // 관리자는 모든 데이터 삭제 가능
  } else if (userRole === '시설관리자') {
    // 시설관리자: 본인 시설 + 하위 시설 데이터 삭제 가능
    const users = loadUsers();
    const managedFacilities = users
      .filter(u =>
        u.facilityName === userFacilityName ||
        (u.parentFacility === userFacilityName && u.role === '시설담당자')
      )
      .map(u => u.facilityName);

    if (!managedFacilities.includes(targetRecord.facilityName)) {
      return res.status(403).json({ success: false, message: '권한이 없습니다.' });
    }
  } else if (userRole === '시설담당자') {
    // 시설담당자: 본인 시설 데이터만 삭제 가능
    if (targetRecord.facilityName !== userFacilityName) {
      return res.status(403).json({ success: false, message: '권한이 없습니다.' });
    }
  }

  records.splice(index, 1);

  if (saveEnergyData(records)) {
    res.json({ success: true, message: '에너지 데이터가 삭제되었습니다.' });
  } else {
    res.json({ success: false, message: '에너지 데이터 삭제 중 오류가 발생했습니다.' });
  }
});

// 에너지 정보 조회 API
app.get('/api/energy-info', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    let infos = loadEnergyInfo();
    const userRole = req.session.user.role;
    const userFacilityName = req.session.user.facilityName;

    // 관리자: 모든 정보 조회 가능
    if (userRole === '관리자') {
      return res.json({ success: true, infos: infos });
    }

    // 시설관리자: 본인 시설 + 하위 시설 정보 조회 가능
    if (userRole === '시설관리자') {
      const users = loadUsers();
      const managedFacilities = users
        .filter(u =>
          u.facilityName === userFacilityName ||
          (u.parentFacility === userFacilityName && u.role === '시설담당자')
        )
        .map(u => u.facilityName);

      infos = infos.filter(info => managedFacilities.includes(info.facilityName));
      return res.json({ success: true, infos: infos });
    }

    // 시설담당자: 본인 시설 정보만 조회 가능
    console.log('시설담당자: 본인 시설만 조회');
    infos = infos.filter(info => info.facilityName === userFacilityName);
    console.log('필터링 후 정보 개수:', infos.length);
    res.json({ success: true, infos: infos });
  } catch (error) {
    console.error('에너지 정보 조회 오류:', error);
    res.status(500).json({ success: false, message: '에너지 정보 조회 중 오류가 발생했습니다.' });
  }
});

// 에너지 정보 추가 API
app.post('/api/energy-info', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const { facilityName, energyType, customerNumber, bankName, accountNumber } = req.body;

    // 입력값 검증: 빈 문자열이나 공백만 있는 경우 체크
    if (!facilityName || !facilityName.trim()) {
      return res.json({ success: false, message: '시설명을 입력해주세요.' });
    }

    if (!energyType || !energyType.trim()) {
      return res.json({ success: false, message: '에너지 종류를 선택해주세요.' });
    }

    if (!customerNumber || !customerNumber.trim()) {
      return res.json({ success: false, message: '고객번호(명세서번호)를 입력해주세요.' });
    }

    if (!bankName || !bankName.trim()) {
      return res.json({ success: false, message: '금융기관을 입력해주세요.' });
    }

    if (!accountNumber || !accountNumber.trim()) {
      return res.json({ success: false, message: '계좌번호를 입력해주세요.' });
    }

    const infos = loadEnergyInfo();

    // 중복 확인 (같은 시설, 같은 에너지 종류)
    const exists = infos.find(i =>
      i.facilityName.trim() === facilityName.trim() &&
      i.energyType.trim() === energyType.trim()
    );

    if (exists) {
      console.log('중복된 정보 발견:', exists);
      return res.json({ success: false, message: '해당 시설의 동일한 에너지 종류 정보가 이미 존재합니다.' });
    }

    const newInfo = {
      facilityName: facilityName.trim(),
      energyType: energyType.trim(),
      customerNumber: customerNumber.trim(),
      bankName: bankName.trim(),
      accountNumber: accountNumber.trim()
    };

    infos.push(newInfo);

    if (saveEnergyInfo(infos)) {
      console.log('에너지 정보 추가 성공:', newInfo);
      res.json({ success: true, message: '에너지 정보가 추가되었습니다.' });
    } else {
      console.error('에너지 정보 저장 실패');
      res.json({ success: false, message: '에너지 정보 추가 중 오류가 발생했습니다.' });
    }
  } catch (error) {
    console.error('에너지 정보 추가 오류:', error);
    res.status(500).json({ success: false, message: '에너지 정보 추가 중 오류가 발생했습니다.' });
  }
});

// 에너지 정보 수정 API
app.put('/api/energy-info/:index', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const index = parseInt(req.params.index);
    const { facilityName, energyType, customerNumber, bankName, accountNumber } = req.body;

    console.log('=== 에너지 정보 수정 요청 ===');
    console.log('인덱스:', index);
    console.log('요청 데이터:', { facilityName, energyType, customerNumber, bankName, accountNumber });

    // 입력값 검증
    if (!facilityName || !facilityName.trim()) {
      return res.json({ success: false, message: '시설명을 입력해주세요.' });
    }

    if (!energyType || !energyType.trim()) {
      return res.json({ success: false, message: '에너지 종류를 선택해주세요.' });
    }

    if (!customerNumber || !customerNumber.trim()) {
      return res.json({ success: false, message: '고객번호(명세서번호)를 입력해주세요.' });
    }

    if (!bankName || !bankName.trim()) {
      return res.json({ success: false, message: '금융기관을 입력해주세요.' });
    }

    if (!accountNumber || !accountNumber.trim()) {
      return res.json({ success: false, message: '계좌번호를 입력해주세요.' });
    }

    const infos = loadEnergyInfo();

    if (index < 0 || index >= infos.length) {
      console.error('유효하지 않은 인덱스:', index, '전체 개수:', infos.length);
      return res.json({ success: false, message: '유효하지 않은 인덱스입니다.' });
    }

    // 수정 시 중복 체크: 다른 항목과 중복되는지 확인 (자기 자신 제외)
    const duplicate = infos.find((info, idx) =>
      idx !== index &&
      info.facilityName.trim() === facilityName.trim() &&
      info.energyType.trim() === energyType.trim()
    );

    if (duplicate) {
      console.log('중복된 정보 발견:', duplicate);
      return res.json({ success: false, message: '해당 시설의 동일한 에너지 종류 정보가 이미 존재합니다.' });
    }

    const updatedInfo = {
      facilityName: facilityName.trim(),
      energyType: energyType.trim(),
      customerNumber: customerNumber.trim(),
      bankName: bankName.trim(),
      accountNumber: accountNumber.trim()
    };

    infos[index] = updatedInfo;

    if (saveEnergyInfo(infos)) {
      console.log('에너지 정보 수정 성공:', updatedInfo);
      res.json({ success: true, message: '에너지 정보가 수정되었습니다.' });
    } else {
      console.error('에너지 정보 저장 실패');
      res.json({ success: false, message: '에너지 정보 수정 중 오류가 발생했습니다.' });
    }
  } catch (error) {
    console.error('에너지 정보 수정 오류:', error);
    res.status(500).json({ success: false, message: '에너지 정보 수정 중 오류가 발생했습니다.' });
  }
});

// 에너지 정보 삭제 API
app.delete('/api/energy-info/:index', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const index = parseInt(req.params.index);
    const infos = loadEnergyInfo();

    if (index < 0 || index >= infos.length) {
      return res.json({ success: false, message: '유효하지 않은 인덱스입니다.' });
    }

    infos.splice(index, 1);

    if (saveEnergyInfo(infos)) {
      res.json({ success: true, message: '에너지 정보가 삭제되었습니다.' });
    } else {
      res.json({ success: false, message: '에너지 정보 삭제 중 오류가 발생했습니다.' });
    }
  } catch (error) {
    console.error('에너지 정보 삭제 오류:', error);
    res.status(500).json({ success: false, message: '에너지 정보 삭제 중 오류가 발생했습니다.' });
  }
});

app.get('/api/data-view', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const { facility, energyType, startDate, endDate } = req.query;
  let records = loadEnergyData();
  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  // 권한에 따라 데이터 필터링
  if (userRole === '관리자') {
    // 관리자는 모든 데이터 조회 가능
  } else if (userRole === '시설관리자') {
    // 시설관리자: 본인 시설 + 하위 시설 데이터 조회 가능
    const users = loadUsers();
    const managedFacilities = users
      .filter(u =>
        u.facilityName === userFacilityName ||
        (u.parentFacility === userFacilityName && u.role === '시설담당자')
      )
      .map(u => u.facilityName);

    records = records.filter(record => managedFacilities.includes(record.facilityName));
  } else if (userRole === '시설담당자') {
    // 시설담당자: 본인 시설 데이터만 조회 가능
    records = records.filter(record => record.facilityName === userFacilityName);
  }

  // 시설 필터 (상위시설 선택 시 하위시설 데이터도 포함)
  if (facility) {
    const users = loadUsers();

    // 선택한 시설의 하위 시설 목록 조회
    // 예: "울주군립야영장" 선택 시 "울주군립야영장(별빛)", "울주군립야영장(달빛)" 등 포함
    const childFacilities = users
      .filter(u => u.parentFacility === facility)
      .map(u => u.facilityName);

    // 시설명이 선택한 시설이거나, 선택한 시설의 하위 시설인 경우 포함
    // 또는 시설명이 선택한 시설명으로 시작하는 경우도 포함 (괄호로 구분되는 하위시설)
    records = records.filter(record => {
      // 정확히 일치
      if (record.facilityName === facility) return true;

      // parentFacility 관계로 연결된 하위시설
      if (childFacilities.includes(record.facilityName)) return true;

      // 시설명이 "상위시설명(" 으로 시작하는 경우 (예: "울주군립야영장(별빛)")
      if (record.facilityName.startsWith(facility + '(')) return true;

      return false;
    });
  }

  // 에너지 종류 필터
  if (energyType) {
    records = records.filter(record => record.energyType === energyType);
  }

  // 기간 필터 (월분 기준)
  if (startDate) {
    // startDate를 YYYY-MM 형식으로 변환
    const startMonth = startDate.substring(0, 7); // "YYYY-MM"
    records = records.filter(record => {
      const billingMonth = record.billingMonth || '';
      return billingMonth >= startMonth;
    });
  }

  if (endDate) {
    // endDate를 YYYY-MM 형식으로 변환
    const endMonth = endDate.substring(0, 7); // "YYYY-MM"
    records = records.filter(record => {
      const billingMonth = record.billingMonth || '';
      return billingMonth <= endMonth;
    });
  }

  // 디버깅: 반환되는 데이터 확인
  console.log('=== 데이터 조회 API 응답 ===');
  console.log('총 레코드 수:', records.length);
  if (records.length > 0) {
    console.log('첫 번째 레코드 샘플:', JSON.stringify(records[0], null, 2));
  }

  res.json({ success: true, records: records });
});

// 에너지 데이터 엑셀 업로드
app.post('/api/energy-data/upload', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const { records: newRecords } = req.body;

  if (!newRecords || !Array.isArray(newRecords) || newRecords.length === 0) {
    return res.status(400).json({ success: false, message: '업로드할 데이터가 없습니다.' });
  }

  try {
    // 기존 데이터 로드
    let existingRecords = loadEnergyData();

    // 새 데이터 추가
    newRecords.forEach(record => {
      existingRecords.push({
        facilityName: record.facilityName || '',
        billingMonth: record.billingMonth || '',
        startDate: record.startDate || '',
        endDate: record.endDate || '',
        energyType: record.energyType || '',
        usageAmount: parseFloat(record.usageAmount) || 0,
        usageCost: parseFloat(record.usageCost) || 0,
        customerNumber: record.customerNumber || '',
        bankName: record.bankName || '',
        virtualAccount: record.virtualAccount || ''
      });
    });

    // 데이터 저장
    if (saveEnergyData(existingRecords)) {
      res.json({ success: true, message: `${newRecords.length}건의 데이터가 업로드되었습니다.`, count: newRecords.length });
    } else {
      res.status(500).json({ success: false, message: '데이터 저장 중 오류가 발생했습니다.' });
    }
  } catch (error) {
    console.error('에너지 데이터 업로드 오류:', error);
    res.status(500).json({ success: false, message: '서버 오류가 발생했습니다.' });
  }
});

// 선택된 에너지 데이터 삭제
app.post('/api/energy-data/delete-multiple', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const { records: recordsToDelete } = req.body;

  if (!recordsToDelete || !Array.isArray(recordsToDelete) || recordsToDelete.length === 0) {
    return res.status(400).json({ success: false, message: '삭제할 레코드가 없습니다.' });
  }

  try {
    let allRecords = loadEnergyData();
    const userRole = req.session.user.role;
    const userFacilityName = req.session.user.facilityName;

    // 권한 확인 및 삭제할 레코드 필터링
    const recordsToDeleteFiltered = recordsToDelete.filter(recordToDelete => {
      if (userRole === '관리자') {
        return true; // 관리자는 모든 데이터 삭제 가능
      } else if (userRole === '시설관리자') {
        // 시설관리자: 본인 시설 + 하위 시설 데이터만 삭제 가능
        const users = loadUsers();
        const managedFacilities = users
          .filter(u =>
            u.facilityName === userFacilityName ||
            (u.parentFacility === userFacilityName && u.role === '시설담당자')
          )
          .map(u => u.facilityName);
        return managedFacilities.includes(recordToDelete.facilityName);
      } else if (userRole === '시설담당자') {
        // 시설담당자: 본인 시설 데이터만 삭제 가능
        return recordToDelete.facilityName === userFacilityName;
      }
      return false;
    });

    if (recordsToDeleteFiltered.length === 0) {
      return res.status(403).json({ success: false, message: '삭제 권한이 없습니다.' });
    }

    // 삭제할 레코드를 제외한 나머지 레코드 필터링
    const remainingRecords = allRecords.filter(record => {
      return !recordsToDeleteFiltered.some(toDelete => {
        return record.facilityName === toDelete.facilityName &&
               record.startDate === toDelete.startDate &&
               record.endDate === toDelete.endDate &&
               record.energyType === toDelete.energyType &&
               parseFloat(record.usageAmount) === parseFloat(toDelete.usageAmount) &&
               parseFloat(record.usageCost) === parseFloat(toDelete.usageCost);
      });
    });

    // 엑셀 파일에 저장
    if (saveEnergyData(remainingRecords)) {
      console.log(`${recordsToDeleteFiltered.length}개의 레코드가 삭제되었습니다.`);
      res.json({ success: true, message: `${recordsToDeleteFiltered.length}개의 레코드가 삭제되었습니다.` });
    } else {
      res.status(500).json({ success: false, message: '데이터 저장 중 오류가 발생했습니다.' });
    }
  } catch (error) {
    console.error('데이터 삭제 오류:', error);
    res.status(500).json({ success: false, message: '데이터 삭제 중 오류가 발생했습니다.' });
  }
});

// 모든 에너지 데이터 삭제 (관리자만 가능)
app.delete('/api/energy-data/delete-all', (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const userRole = req.session.user.role;

  // 관리자만 모든 데이터 삭제 가능
  if (userRole !== '관리자') {
    return res.status(403).json({ success: false, message: '관리자만 모든 데이터를 삭제할 수 있습니다.' });
  }

  try {
    const allRecords = loadEnergyData();
    const recordCount = allRecords.length;

    // 빈 배열로 저장 (모든 데이터 삭제)
    if (saveEnergyData([])) {
      console.log(`관리자가 모든 에너지 데이터 삭제: ${recordCount}개`);
      res.json({ success: true, message: `${recordCount}개의 데이터가 삭제되었습니다.`, deletedCount: recordCount });
    } else {
      res.status(500).json({ success: false, message: '데이터 저장 중 오류가 발생했습니다.' });
    }
  } catch (error) {
    console.error('모든 데이터 삭제 오류:', error);
    res.status(500).json({ success: false, message: '데이터 삭제 중 오류가 발생했습니다.' });
  }
});

// 숫자를 한글로 변환하는 함수 (일십, 일백, 일천 포함)
function numberToKorean(num) {
  let number = parseInt(num);
  if (number === 0) return '영';

  const units = ['', '만', '억', '조'];
  const digits = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
  const positions = ['', '십', '백', '천'];

  let result = '';
  let unitIndex = 0;

  while (number > 0) {
    const part = number % 10000;
    if (part > 0) {
      let partStr = '';
      let tempPart = part;
      let posIndex = 0;

      while (tempPart > 0) {
        const digit = tempPart % 10;
        if (digit > 0) {
          // 1일 때도 항상 "일"을 포함 (일십, 일백, 일천)
          partStr = digits[digit] + positions[posIndex] + partStr;
        }
        tempPart = Math.floor(tempPart / 10);
        posIndex++;
      }
      result = partStr + units[unitIndex] + result;
    }
    number = Math.floor(number / 10000);
    unitIndex++;
  }

  return result;
}

// 공문 생성 API
app.post('/api/generate-document', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const record = req.body;

    // billingMonth가 있으면 사용, 없으면 startDate에서 추출
    let year, month;
    if (record.billingMonth) {
      // billingMonth 형식: "YYYY-MM"
      const billingParts = record.billingMonth.split('-');
      year = parseInt(billingParts[0]);
      month = parseInt(billingParts[1]);
    } else {
      // startDate에서 추출
      const startDate = new Date(record.startDate);
      year = startDate.getFullYear();
      month = startDate.getMonth() + 1;
    }

    const costNumber = parseInt(record.usageCost);
    const costKorean = numberToKorean(costNumber);

    // 금융기관과 계좌번호 정보 사용
    let paymentMethod = '첨부 2 고지서 참조';
    if (record.bankName && record.virtualAccount) {
      paymentMethod = `계좌입금(${record.bankName} ${record.virtualAccount})`;
    } else if (record.virtualAccount) {
      paymentMethod = `계좌입금(${record.virtualAccount})`;
    }

    // Word 문서 생성
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: '지출결의서',
                bold: true,
                size: 32
              })
            ],
            spacing: {
              after: 400
            }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `${year}년 ${month}월 ${record.facilityName} ${record.energyType} 요금을 아래와 같이 납부하고자 합니다.`,
                size: 24
              })
            ],
            spacing: {
              after: 300
            }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: '',
                size: 24
              })
            ],
            spacing: {
              after: 200
            }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `  1. 소요금액: 금${costNumber.toLocaleString('ko-KR')}원(금${costKorean}원)`,
                size: 24
              })
            ],
            spacing: {
              after: 200
            }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: '  2. 세부내역: 첨부 1 참조',
                size: 24
              })
            ],
            spacing: {
              after: 200
            }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `  3. 납부방법: ${paymentMethod}`,
                size: 24
              })
            ],
            spacing: {
              after: 200
            }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: '  4. 예산과목: ',
                size: 24
              })
            ],
            spacing: {
              after: 400
            }
          })
        ]
      }]
    });

    // 문서를 버퍼로 변환
    const buffer = await Packer.toBuffer(doc);

    // 파일명 설정
    const filename = `${year}-${String(month).padStart(2, '0')}-${record.facilityName}-${record.energyType}-공문.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.send(buffer);

  } catch (error) {
    console.error('공문 생성 오류:', error);
    res.status(500).json({ success: false, message: '공문 생성 중 오류가 발생했습니다.' });
  }
});

// 첨부1 생성 API (Excel) - 첨부2 서식 사용
app.post('/api/generate-attachment1', async (req, res) => {
  console.log('=== 첨부1 생성 API 호출 ===');

  if (!req.session.user) {
    console.log('인증 실패: 세션 없음');
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const record = req.body;
    console.log('개별 첨부문서 생성 요청:', JSON.stringify(record, null, 2));

    if (!record || Object.keys(record).length === 0) {
      console.log('오류: 빈 레코드');
      return res.status(400).json({ success: false, message: '데이터가 없습니다.' });
    }

    // billingMonth가 있으면 사용, 없으면 startDate에서 추출
    let year, month;
    if (record.billingMonth) {
      const billingParts = record.billingMonth.split('-');
      year = parseInt(billingParts[0]);
      month = parseInt(billingParts[1]);
    } else if (record.startDate) {
      const startDate = new Date(record.startDate);
      year = startDate.getFullYear();
      month = startDate.getMonth() + 1;
    } else {
      year = new Date().getFullYear();
      month = new Date().getMonth() + 1;
    }
    console.log('년월:', year, month);

    // 에너지 정보에서 금융기관과 계좌번호 조회
    const energyInfos = loadEnergyInfo();
    let matchingInfo = energyInfos.find(info =>
      info.facilityName === record.facilityName &&
      info.energyType === record.energyType
    );
    // 정확한 시설명 매칭 실패 시, 상위 시설명으로 폴백 조회
    // 예: "울주군립야영장(별빛)" → "울주군립야영장"
    if (!matchingInfo) {
      const parentMatch = record.facilityName ? record.facilityName.match(/^(.+?)\(/) : null;
      if (parentMatch) {
        const parentFacilityName = parentMatch[1];
        matchingInfo = energyInfos.find(info =>
          info.facilityName === parentFacilityName &&
          info.energyType === record.energyType
        );
      }
    }

    const bankName = matchingInfo ? matchingInfo.bankName : (record.bankName || '');
    const accountNumber = matchingInfo ? matchingInfo.accountNumber : (record.virtualAccount || '');
    console.log('조회된 금융정보:', { bankName, accountNumber, matchingInfo: !!matchingInfo });

    // 템플릿 파일 읽기 (첨부 2 서식 사용)
    console.log('템플릿 파일 경로:', ATTACHMENT_TEMPLATE_FILE);
    if (!fs.existsSync(ATTACHMENT_TEMPLATE_FILE)) {
      console.error('템플릿 파일이 존재하지 않습니다:', ATTACHMENT_TEMPLATE_FILE);
      return res.status(500).json({ success: false, message: '템플릿 파일을 찾을 수 없습니다.' });
    }
    const workbook = xlsx.readFile(ATTACHMENT_TEMPLATE_FILE);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 제목 변경: "납부내역" -> "세부내역"
    worksheet['A1'] = { t: 's', v: '세부내역' };

    // 데이터 시작 행 (템플릿의 6행 - 5행에 ~ 구분자가 있음)
    const dataRow = 6;

    // 년도 (A열)
    worksheet['A' + dataRow] = { t: 'n', v: year };

    // 월 (B열)
    worksheet['B' + dataRow] = { t: 'n', v: month };

    // 사용시설 (C열)
    worksheet['C' + dataRow] = { t: 's', v: record.facilityName || '' };

    // 금융기관 (D열)
    worksheet['D' + dataRow] = { t: 's', v: bankName };

    // 계좌번호 (E열)
    worksheet['E' + dataRow] = { t: 's', v: accountNumber };

    // 사용기간 시작 (F열)
    worksheet['F' + dataRow] = { t: 's', v: record.startDate || '' };

    // 사용기간 구분자 (G열)
    worksheet['G' + dataRow] = { t: 's', v: '~' };

    // 사용기간 종료 (H열)
    worksheet['H' + dataRow] = { t: 's', v: record.endDate || '' };

    // 납부금액 (I열) - 천단위 콤마
    worksheet['I' + dataRow] = {
      t: 'n',
      v: parseFloat(record.usageCost) || 0,
      z: '#,##0'
    };

    // 비고 (J열) - 에너지 종류
    worksheet['J' + dataRow] = { t: 's', v: record.energyType || '' };

    // 워크시트 범위 업데이트
    worksheet['!ref'] = xlsx.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: dataRow, c: 9 }
    });

    // 열 너비 설정
    worksheet['!cols'] = [
      { wch: 8 },   // A: 년도
      { wch: 6 },   // B: 월
      { wch: 25 },  // C: 사용시설
      { wch: 12 },  // D: 금융기관
      { wch: 18 },  // E: 계좌번호
      { wch: 12 },  // F: 사용기간 시작
      { wch: 3 },   // G: ~
      { wch: 12 },  // H: 사용기간 종료
      { wch: 15 },  // I: 납부금액
      { wch: 12 }   // J: 비고
    ];

    // Excel 파일을 버퍼로 변환
    const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // 파일명 설정
    const filename = `${year}-${String(month).padStart(2, '0')}-${record.facilityName}-${record.energyType}-첨부1.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.send(buffer);

    console.log('개별 첨부문서 생성 완료:', filename);

  } catch (error) {
    console.error('첨부1 생성 오류:', error.message);
    console.error('스택:', error.stack);
    res.status(500).json({ success: false, message: '첨부1 생성 중 오류: ' + error.message });
  }
});

// 선택된 데이터를 하나의 엑셀 파일로 생성하는 API (기존 첨부 2 서식 활용)
app.post('/api/generate-attachment-combined', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const { records } = req.body;

    if (!records || !Array.isArray(records) || records.length === 0) {
      return res.status(400).json({ success: false, message: '데이터가 없습니다.' });
    }

    console.log('통합 첨부문서 생성:', records.length, '건');

    // 에너지 정보 로드 (금융기관/계좌번호 조회용)
    const energyInfos = loadEnergyInfo();

    // 템플릿 파일 읽기 (첨부 2 서식 사용)
    if (!fs.existsSync(ATTACHMENT_TEMPLATE_FILE)) {
      console.error('템플릿 파일이 존재하지 않습니다:', ATTACHMENT_TEMPLATE_FILE);
      return res.status(500).json({ success: false, message: '템플릿 파일을 찾을 수 없습니다.' });
    }
    const workbook = xlsx.readFile(ATTACHMENT_TEMPLATE_FILE);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 제목 변경: "납부내역" -> "세부내역"
    worksheet['A1'] = { t: 's', v: '세부내역' };

    // 합계 금액 계산
    const totalCost = records.reduce((sum, r) => sum + (parseFloat(r.usageCost) || 0), 0);

    // 합계 금액을 I5에 입력
    worksheet['I5'] = {
      t: 'n',
      v: totalCost,
      z: '#,##0'
    };

    // 데이터 시작 행 (템플릿의 6행부터)
    const dataStartRow = 6;

    // 각 레코드에 대해 데이터 행 추가
    records.forEach((record, index) => {
      const currentRow = dataStartRow + index;

      // billingMonth에서 년/월 추출
      let year, month;
      if (record.billingMonth) {
        const billingParts = record.billingMonth.split('-');
        year = parseInt(billingParts[0]);
        month = parseInt(billingParts[1]);
      } else if (record.startDate) {
        const startDate = new Date(record.startDate);
        year = startDate.getFullYear();
        month = startDate.getMonth() + 1;
      } else {
        year = new Date().getFullYear();
        month = new Date().getMonth() + 1;
      }

      // 에너지 정보에서 금융기관과 계좌번호 조회
      let matchingInfo = energyInfos.find(info =>
        info.facilityName === record.facilityName &&
        info.energyType === record.energyType
      );
      // 정확한 시설명 매칭 실패 시, 상위 시설명으로 폴백 조회
      // 예: "울주군립야영장(별빛)" → "울주군립야영장"
      if (!matchingInfo) {
        const parentMatch = record.facilityName ? record.facilityName.match(/^(.+?)\(/) : null;
        if (parentMatch) {
          const parentFacilityName = parentMatch[1];
          matchingInfo = energyInfos.find(info =>
            info.facilityName === parentFacilityName &&
            info.energyType === record.energyType
          );
        }
      }
      const bankName = matchingInfo ? matchingInfo.bankName : (record.bankName || '');
      const accountNumber = matchingInfo ? matchingInfo.accountNumber : (record.virtualAccount || '');

      // 년도 (A열)
      worksheet['A' + currentRow] = { t: 'n', v: year };

      // 월 (B열)
      worksheet['B' + currentRow] = { t: 'n', v: month };

      // 사용시설 (C열)
      worksheet['C' + currentRow] = { t: 's', v: record.facilityName || '' };

      // 금융기관 (D열)
      worksheet['D' + currentRow] = { t: 's', v: bankName };

      // 계좌번호 (E열)
      worksheet['E' + currentRow] = { t: 's', v: accountNumber };

      // 사용기간 시작 (F열)
      worksheet['F' + currentRow] = { t: 's', v: record.startDate || '' };

      // 사용기간 구분자 (G열)
      worksheet['G' + currentRow] = { t: 's', v: '~' };

      // 사용기간 종료 (H열)
      worksheet['H' + currentRow] = { t: 's', v: record.endDate || '' };

      // 납부금액 (I열) - 천단위 콤마
      worksheet['I' + currentRow] = {
        t: 'n',
        v: parseFloat(record.usageCost) || 0,
        z: '#,##0'
      };

      // 비고 (J열) - 에너지 종류
      worksheet['J' + currentRow] = { t: 's', v: record.energyType || '' };
    });

    // 워크시트 범위 업데이트
    const lastDataRow = dataStartRow + records.length - 1;
    worksheet['!ref'] = xlsx.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: lastDataRow, c: 9 }  // J열(인덱스 9)까지
    });

    // 열 너비 설정
    worksheet['!cols'] = [
      { wch: 8 },   // A: 년도
      { wch: 6 },   // B: 월
      { wch: 25 },  // C: 사용시설
      { wch: 12 },  // D: 금융기관
      { wch: 18 },  // E: 계좌번호
      { wch: 12 },  // F: 사용기간 시작
      { wch: 3 },   // G: ~
      { wch: 12 },  // H: 사용기간 종료
      { wch: 15 },  // I: 납부금액
      { wch: 12 }   // J: 비고
    ];

    // Excel 파일을 버퍼로 변환
    const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // 파일명 설정
    const now = new Date();
    const filename = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-통합첨부문서.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.send(buffer);

    console.log('통합 첨부문서 생성 완료:', records.length, '건');

  } catch (error) {
    console.error('통합 첨부문서 생성 오류:', error);
    res.status(500).json({ success: false, message: '통합 첨부문서 생성 중 오류가 발생했습니다.' });
  }
});

// 네이버 클로바 OCR API 프록시 엔드포인트
app.post('/api/clova-ocr', async (req, res) => {
  try {
    const { imageBase64, apiUrl, secretKey } = req.body;

    console.log('클로바 OCR 요청 시작');
    console.log('API URL:', apiUrl);

    if (!imageBase64 || !apiUrl || !secretKey) {
      return res.status(400).json({
        success: false,
        message: 'API URL, Secret Key 및 이미지 데이터가 필요합니다.'
      });
    }

    // URL 정리 (끝에 슬래시 제거)
    const cleanApiUrl = apiUrl.replace(/\/$/, '');
    console.log('정리된 API URL:', cleanApiUrl);

    // Base64에서 이미지 포맷 감지
    let imageFormat = 'jpg';
    let contentType = 'image/jpeg';

    if (imageBase64.startsWith('data:image/png')) {
      imageFormat = 'png';
      contentType = 'image/png';
    } else if (imageBase64.startsWith('data:image/jpeg') || imageBase64.startsWith('data:image/jpg')) {
      imageFormat = 'jpg';
      contentType = 'image/jpeg';
    }

    console.log('이미지 포맷:', imageFormat);

    // Base64를 Buffer로 변환
    const base64Data = imageBase64.split(',')[1];
    if (!base64Data) {
      throw new Error('올바르지 않은 Base64 데이터');
    }
    const imageBuffer = Buffer.from(base64Data, 'base64');
    console.log('이미지 버퍼 크기:', imageBuffer.length, 'bytes');

    // 네이버 클로바 OCR API 호출
    const FormData = require('form-data');
    const axios = require('axios');

    const formData = new FormData();

    // OCR 요청 데이터 구성
    const requestJson = {
      version: 'V2',
      requestId: `req_${Date.now()}`,
      timestamp: Date.now(),
      images: [{
        format: imageFormat,
        name: 'ocr_image'
      }]
    };

    console.log('요청 JSON:', requestJson);

    formData.append('message', JSON.stringify(requestJson));
    formData.append('file', imageBuffer, {
      filename: `image.${imageFormat}`,
      contentType: contentType
    });

    // API 호출
    console.log('클로바 API 호출 중...');
    console.log('최종 요청 URL:', cleanApiUrl);
    console.log('요청 헤더:', {
      ...formData.getHeaders(),
      'X-OCR-SECRET': '***' // 보안을 위해 마스킹
    });

    const response = await axios.post(cleanApiUrl, formData, {
      headers: {
        ...formData.getHeaders(),
        'X-OCR-SECRET': secretKey
      },
      timeout: 30000,
      maxContentLength: Infinity,
      maxBodyLength: Infinity
    });

    console.log('클로바 API 응답 상태:', response.status);
    console.log('응답 데이터 구조:', JSON.stringify(response.data).substring(0, 200));

    // OCR 결과에서 텍스트 추출
    let extractedText = '';
    if (response.data && response.data.images && response.data.images[0]) {
      const fields = response.data.images[0].fields || [];
      console.log('추출된 필드 개수:', fields.length);

      // 각 필드의 텍스트를 개행 또는 공백으로 연결
      extractedText = fields.map(field => field.inferText).join('\n');
    }

    console.log('추출된 텍스트 길이:', extractedText.length);

    res.json({
      success: true,
      text: extractedText,
      fullResponse: response.data
    });

  } catch (error) {
    console.error('클로바 OCR API 오류 상세:');
    console.error('- 메시지:', error.message);
    console.error('- 스택:', error.stack);

    if (error.response) {
      console.error('- 응답 상태:', error.response.status);
      console.error('- 응답 헤더:', error.response.headers);
      console.error('- 응답 데이터:', JSON.stringify(error.response.data, null, 2));
    }

    // 사용자 친화적인 오류 메시지 생성
    let userMessage = 'OCR 처리 중 오류가 발생했습니다.';
    let errorDetails = null;

    if (error.response) {
      const status = error.response.status;
      const data = error.response.data;

      switch (status) {
        case 400:
          userMessage = 'API 요청 형식이 잘못되었습니다. API URL과 이미지를 확인하세요.';
          break;
        case 401:
          userMessage = 'API 인증 실패: Secret Key가 올바르지 않습니다.';
          break;
        case 403:
          userMessage = 'API 접근 권한이 없습니다. OCR 서비스가 활성화되어 있는지 확인하세요.';
          break;
        case 404:
          userMessage = 'API URL이 잘못되었습니다. APIGW Invoke URL을 다시 확인하세요.';
          break;
        case 429:
          userMessage = 'API 사용량 한도를 초과했습니다. 잠시 후 다시 시도하세요.';
          break;
        case 500:
        case 502:
        case 503:
          userMessage = '클로바 서버 오류입니다. 잠시 후 다시 시도하세요.';
          break;
      }

      errorDetails = {
        status: status,
        statusText: error.response.statusText,
        data: data
      };
    } else if (error.code === 'ECONNREFUSED') {
      userMessage = '클로바 API 서버에 연결할 수 없습니다. 인터넷 연결을 확인하세요.';
    } else if (error.code === 'ETIMEDOUT') {
      userMessage = '요청 시간이 초과되었습니다. 이미지 크기가 너무 크거나 네트워크가 느립니다.';
    }

    res.status(500).json({
      success: false,
      message: userMessage,
      error: error.message,
      details: errorDetails
    });
  }
});

// 클로바 API 키 검증 엔드포인트
app.post('/api/validate-clova-key', async (req, res) => {
  try {
    const { apiUrl, secretKey } = req.body;

    console.log('API 키 검증 시작');
    console.log('API URL:', apiUrl);

    if (!apiUrl || !secretKey) {
      return res.status(400).json({
        success: false,
        message: 'API URL과 Secret Key가 필요합니다.'
      });
    }

    // URL 형식 검증
    if (!apiUrl.includes('apigw.ntruss.com')) {
      return res.json({
        success: false,
        message: 'API URL 형식이 올바르지 않습니다.\n네이버 클라우드의 APIGW Invoke URL을 입력하세요.\n(예: https://xxxxx.apigw.ntruss.com/custom/v1/xxxxx/xxxxxxxx)'
      });
    }

    // URL이 https로 시작하는지 확인
    if (!apiUrl.startsWith('https://')) {
      return res.json({
        success: false,
        message: 'API URL은 https://로 시작해야 합니다.\n전체 URL을 복사하여 붙여넣으세요.'
      });
    }

    // URL 끝에 슬래시가 있으면 제거
    const cleanApiUrl = apiUrl.replace(/\/$/, '');
    console.log('정리된 API URL:', cleanApiUrl);

    // 간단한 테스트 이미지로 API 키 검증 (1x1 픽셀 PNG)
    const testImageBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';
    const testImageBuffer = Buffer.from(testImageBase64, 'base64');

    const FormData = require('form-data');
    const axios = require('axios');

    const formData = new FormData();

    const requestJson = {
      version: 'V2',
      requestId: `test_${Date.now()}`,
      timestamp: Date.now(),
      images: [{
        format: 'png',
        name: 'test_image'
      }]
    };

    formData.append('message', JSON.stringify(requestJson));
    formData.append('file', testImageBuffer, {
      filename: 'test.png',
      contentType: 'image/png'
    });

    const response = await axios.post(cleanApiUrl, formData, {
      headers: {
        ...formData.getHeaders(),
        'X-OCR-SECRET': secretKey
      },
      timeout: 10000
    });

    console.log('API 키 검증 성공:', response.status);

    res.json({
      success: true,
      message: 'API 키가 유효합니다!'
    });

  } catch (error) {
    console.error('API 키 검증 실패:', error.message);

    let message = 'API 키 검증 실패';

    if (error.response) {
      const status = error.response.status;
      console.error('응답 상태:', status);
      console.error('응답 데이터:', error.response.data);

      switch (status) {
        case 401:
          message = '❌ Secret Key가 올바르지 않습니다.\n\n해결 방법:\n1. 네이버 클라우드 콘솔에서 Secret Key를 다시 복사하세요\n2. 공백이나 줄바꿈 없이 정확히 붙여넣으세요';
          break;
        case 403:
          message = '❌ API 접근 권한이 없습니다.\n\n해결 방법:\n1. CLOVA OCR 서비스를 이용 신청했는지 확인\n2. General OCR 도메인을 생성했는지 확인\n3. 도메인 상태가 "사용중"인지 확인';
          break;
        case 404:
          message = '❌ API URL을 찾을 수 없습니다 (404 오류)\n\n해결 방법:\n1. "APIGW Invoke URL" 전체를 복사했는지 확인\n   (https://로 시작하는 전체 URL)\n2. URL 끝에 여분의 공백이나 슬래시가 없는지 확인\n3. General OCR 도메인인지 확인 (Template OCR 아님)\n\n올바른 형식:\nhttps://xxxxx.apigw.ntruss.com/custom/v1/xxxxx/xxxxxxxx';
          break;
        default:
          message = `❌ API 오류 (${status}): ${error.response.statusText}`;
      }
    } else if (error.code === 'ECONNREFUSED') {
      message = '❌ 클로바 API 서버에 연결할 수 없습니다.\n인터넷 연결을 확인하세요.';
    } else if (error.code === 'ETIMEDOUT') {
      message = '❌ 요청 시간이 초과되었습니다.';
    }

    res.json({
      success: false,
      message: message,
      details: error.response ? {
        status: error.response.status,
        data: error.response.data
      } : null
    });
  }
});

app.get('/', (req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

app.listen(PORT, () => {
  console.log(`에너지 관리 시스템이 http://localhost:${PORT} 에서 실행중입니다.`);
});
