require('dotenv').config();

const express = require('express');
const session = require('express-session');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');
const fs = require('fs');
const path = require('path');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = 3000;
const ATTACHMENT_TEMPLATE_FILE = path.join(__dirname, 'public', '첨부 2 서식.xlsx');

// Supabase 클라이언트 초기화
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
);

// Body parser 설정 - 이미지 업로드를 위해 크기 제한 증가
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static('public'));

app.use(session({
  secret: process.env.SESSION_SECRET || 'energy-management-secret-key',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 24 * 60 * 60 * 1000 }
}));

// ──────────────────────────────────────────────
// 헬퍼: Supabase 행(snake_case) → camelCase 변환
// ──────────────────────────────────────────────

function rowToUser(row) {
  return {
    id: row.id,
    facilityName: row.facility_name,
    username: row.username,
    password: row.password,
    role: row.role || '시설담당자',
    parentFacility: row.parent_facility || ''
  };
}

function rowToEnergyRecord(row) {
  return {
    id: row.id,
    facilityName: row.facility_name,
    billingMonth: row.billing_month,
    startDate: row.start_date,
    endDate: row.end_date,
    energyType: row.energy_type,
    usageAmount: row.usage_amount,
    usageCost: row.usage_cost,
    bankName: row.bank_name,
    virtualAccount: row.virtual_account,
    customerNumber: row.customer_number
  };
}

function rowToEnergyInfo(row) {
  return {
    id: row.id,
    facilityName: row.facility_name,
    energyType: row.energy_type,
    customerNumber: row.customer_number,
    bankName: row.bank_name,
    accountNumber: row.account_number
  };
}

// ──────────────────────────────────────────────
// 인증 API
// ──────────────────────────────────────────────

app.post('/api/login', async (req, res) => {
  const { facilityName, username, password } = req.body;

  console.log('로그인 시도:', { facilityName, username, password });

  try {
    const { data, error } = await supabase
      .from('users')
      .select('*')
      .eq('facility_name', facilityName)
      .eq('username', username)
      .eq('password', password)
      .single();

    if (error || !data) {
      console.log('로그인 실패');
      return res.json({ success: false, message: '시설명, 아이디 또는 비밀번호가 올바르지 않습니다.' });
    }

    req.session.user = {
      id: data.username,
      facilityName: data.facility_name,
      role: data.role || '시설담당자'
    };
    console.log('로그인 성공:', req.session.user);
    res.json({ success: true, user: req.session.user });
  } catch (err) {
    console.error('로그인 오류:', err);
    res.status(500).json({ success: false, message: '로그인 중 오류가 발생했습니다.' });
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

// ──────────────────────────────────────────────
// 시설(users) API
// ──────────────────────────────────────────────

app.get('/api/facilities', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const userRole = req.session.user.role;
    const userFacilityName = req.session.user.facilityName;

    let query = supabase.from('users').select('*');

    if (userRole === '관리자') {
      // 관리자: 모든 시설 조회
    } else if (userRole === '시설관리자') {
      // 시설관리자: 본인 시설 + parentFacility가 본인인 시설담당자
      query = query.or(
        `facility_name.eq.${userFacilityName},and(parent_facility.eq.${userFacilityName},role.eq.시설담당자)`
      );
    } else {
      // 시설담당자: 본인 시설만
      query = query.eq('facility_name', userFacilityName);
    }

    const { data, error } = await query;

    if (error) throw error;

    const facilities = (data || []).map(rowToUser);
    res.json({ success: true, facilities });
  } catch (err) {
    console.error('시설 조회 오류:', err);
    res.status(500).json({ success: false, message: '시설 조회 중 오류가 발생했습니다.' });
  }
});

app.post('/api/facilities', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  if (userRole !== '관리자' && userRole !== '시설관리자') {
    return res.status(403).json({ success: false, message: '권한이 없습니다.' });
  }

  const { facilityName, id, password, role, parentFacility } = req.body;

  if (!facilityName || !id || !password || !role) {
    return res.json({ success: false, message: '모든 필드를 입력해주세요.' });
  }

  try {
    // 중복 확인
    const { data: existing } = await supabase
      .from('users')
      .select('id')
      .eq('facility_name', facilityName)
      .eq('username', id)
      .single();

    if (existing) {
      return res.json({ success: false, message: '이미 존재하는 시설명 또는 아이디입니다.' });
    }

    let insertData = {
      facility_name: facilityName,
      username: id,
      password: password,
      role: role,
      parent_facility: parentFacility || ''
    };

    // 시설관리자는 시설담당자만 추가 가능하고, parentFacility는 자동으로 본인 시설
    if (userRole === '시설관리자') {
      if (role !== '시설담당자') {
        return res.status(403).json({ success: false, message: '시설관리자는 시설담당자만 추가할 수 있습니다.' });
      }
      insertData.parent_facility = userFacilityName;
    }

    const { error } = await supabase.from('users').insert(insertData);

    if (error) throw error;

    res.json({ success: true, message: '시설이 추가되었습니다.' });
  } catch (err) {
    console.error('시설 추가 오류:', err);
    res.status(500).json({ success: false, message: '시설 추가 중 오류가 발생했습니다.' });
  }
});

app.put('/api/facilities/:id', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const targetId = parseInt(req.params.id);
  const { facilityName, id, password, role, parentFacility } = req.body;

  if (!facilityName || !id || !password || !role) {
    return res.json({ success: false, message: '모든 필드를 입력해주세요.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  try {
    // 수정 대상 조회
    const { data: targetUser, error: fetchError } = await supabase
      .from('users')
      .select('*')
      .eq('id', targetId)
      .single();

    if (fetchError || !targetUser) {
      return res.json({ success: false, message: '유효하지 않은 ID입니다.' });
    }

    // 권한 확인
    if (userRole === '시설관리자') {
      const isSelf = targetUser.facility_name === userFacilityName;
      const isChild = targetUser.parent_facility === userFacilityName && targetUser.role === '시설담당자';

      if (!isSelf && !isChild) {
        return res.status(403).json({ success: false, message: '권한이 없습니다.' });
      }

      if (isChild && role !== '시설담당자') {
        return res.status(403).json({ success: false, message: '시설관리자는 시설담당자만 수정할 수 있습니다.' });
      }

      const updateData = {
        facility_name: facilityName,
        username: id,
        password: password,
        role: role,
        parent_facility: isChild ? userFacilityName : (parentFacility || '')
      };

      const { error } = await supabase.from('users').update(updateData).eq('id', targetId);
      if (error) throw error;
      return res.json({ success: true, message: '시설 정보가 수정되었습니다.' });
    }

    if (userRole === '관리자') {
      const { error } = await supabase.from('users').update({
        facility_name: facilityName,
        username: id,
        password: password,
        role: role,
        parent_facility: parentFacility || ''
      }).eq('id', targetId);

      if (error) throw error;
      return res.json({ success: true, message: '시설 정보가 수정되었습니다.' });
    }

    return res.status(403).json({ success: false, message: '권한이 없습니다.' });
  } catch (err) {
    console.error('시설 수정 오류:', err);
    res.status(500).json({ success: false, message: '시설 수정 중 오류가 발생했습니다.' });
  }
});

app.delete('/api/facilities/:id', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  if (userRole !== '관리자' && userRole !== '시설관리자') {
    return res.status(403).json({ success: false, message: '권한이 없습니다.' });
  }

  const targetId = parseInt(req.params.id);

  try {
    const { data: targetUser, error: fetchError } = await supabase
      .from('users')
      .select('*')
      .eq('id', targetId)
      .single();

    if (fetchError || !targetUser) {
      return res.json({ success: false, message: '유효하지 않은 ID입니다.' });
    }

    if (userRole === '시설관리자') {
      const canDelete =
        targetUser.facility_name === userFacilityName ||
        (targetUser.parent_facility === userFacilityName && targetUser.role === '시설담당자');
      if (!canDelete) {
        return res.status(403).json({ success: false, message: '권한이 없습니다.' });
      }
    }

    const { error } = await supabase.from('users').delete().eq('id', targetId);
    if (error) throw error;

    res.json({ success: true, message: '시설이 삭제되었습니다.' });
  } catch (err) {
    console.error('시설 삭제 오류:', err);
    res.status(500).json({ success: false, message: '시설 삭제 중 오류가 발생했습니다.' });
  }
});

// ──────────────────────────────────────────────
// 에너지 데이터(energy_records) API
// ──────────────────────────────────────────────

// 시설관리자의 관리 시설 목록을 Supabase에서 조회
async function getManagedFacilityNames(userFacilityName) {
  const { data } = await supabase
    .from('users')
    .select('facility_name')
    .or(`facility_name.eq.${userFacilityName},and(parent_facility.eq.${userFacilityName},role.eq.시설담당자)`);
  return (data || []).map(u => u.facility_name);
}

app.get('/api/energy', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  try {
    let query = supabase.from('energy_records').select('*');

    if (userRole === '관리자') {
      // 모든 데이터 조회
    } else if (userRole === '시설관리자') {
      const facilityNames = await getManagedFacilityNames(userFacilityName);
      query = query.in('facility_name', facilityNames);
    } else {
      query = query.eq('facility_name', userFacilityName);
    }

    const { data, error } = await query;
    if (error) throw error;

    const records = (data || []).map(rowToEnergyRecord);
    res.json({ success: true, records });
  } catch (err) {
    console.error('에너지 데이터 조회 오류:', err);
    res.status(500).json({ success: false, message: '에너지 데이터 조회 중 오류가 발생했습니다.' });
  }
});

app.post('/api/energy', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const { facilityName, energyType, billingMonth, startDate, endDate, usageAmount, usageCost } = req.body;

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  try {
    // 권한 확인
    if (userRole === '시설관리자') {
      const facilityNames = await getManagedFacilityNames(userFacilityName);
      if (!facilityNames.includes(facilityName)) {
        return res.status(403).json({ success: false, message: '해당 시설에 데이터를 입력할 권한이 없습니다.' });
      }
    } else if (userRole === '시설담당자') {
      if (facilityName !== userFacilityName) {
        return res.status(403).json({ success: false, message: '해당 시설에 데이터를 입력할 권한이 없습니다.' });
      }
    }

    // 필드 유효성 검사
    if (!facilityName || !energyType || !startDate || !endDate) {
      return res.json({ success: false, message: '모든 필드를 입력해주세요.' });
    }

    if (usageAmount === undefined || usageAmount === null || usageAmount === '' ||
        usageCost === undefined || usageCost === null || usageCost === '') {
      return res.json({ success: false, message: '모든 필드를 입력해주세요.' });
    }

    const parsedUsageAmount = typeof usageAmount === 'number' ? usageAmount : parseFloat(usageAmount);
    const parsedUsageCost = typeof usageCost === 'number' ? usageCost : parseFloat(usageCost);

    if (isNaN(parsedUsageAmount) || isNaN(parsedUsageCost)) {
      return res.json({ success: false, message: '사용량과 사용 금액은 숫자여야 합니다.' });
    }

    if (new Date(startDate) > new Date(endDate)) {
      return res.json({ success: false, message: '종료일은 시작일 이후여야 합니다.' });
    }

    const { error } = await supabase.from('energy_records').insert({
      facility_name: facilityName,
      energy_type: energyType,
      billing_month: billingMonth,
      start_date: startDate,
      end_date: endDate,
      usage_amount: parsedUsageAmount,
      usage_cost: parsedUsageCost
    });

    if (error) throw error;

    res.json({ success: true, message: '에너지 사용량이 저장되었습니다.' });
  } catch (err) {
    console.error('에너지 데이터 저장 오류:', err);
    res.status(500).json({ success: false, message: '에너지 데이터 저장 중 오류가 발생했습니다.' });
  }
});

app.delete('/api/energy/:id', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const targetId = parseInt(req.params.id);
  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  try {
    const { data: targetRecord, error: fetchError } = await supabase
      .from('energy_records')
      .select('*')
      .eq('id', targetId)
      .single();

    if (fetchError || !targetRecord) {
      return res.json({ success: false, message: '유효하지 않은 ID입니다.' });
    }

    if (userRole === '시설관리자') {
      const facilityNames = await getManagedFacilityNames(userFacilityName);
      if (!facilityNames.includes(targetRecord.facility_name)) {
        return res.status(403).json({ success: false, message: '권한이 없습니다.' });
      }
    } else if (userRole === '시설담당자') {
      if (targetRecord.facility_name !== userFacilityName) {
        return res.status(403).json({ success: false, message: '권한이 없습니다.' });
      }
    }

    const { error } = await supabase.from('energy_records').delete().eq('id', targetId);
    if (error) throw error;

    res.json({ success: true, message: '에너지 데이터가 삭제되었습니다.' });
  } catch (err) {
    console.error('에너지 데이터 삭제 오류:', err);
    res.status(500).json({ success: false, message: '에너지 데이터 삭제 중 오류가 발생했습니다.' });
  }
});

// ──────────────────────────────────────────────
// 에너지 정보(energy_info) API
// ──────────────────────────────────────────────

app.get('/api/energy-info', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const userRole = req.session.user.role;
    const userFacilityName = req.session.user.facilityName;

    let query = supabase.from('energy_info').select('*');

    if (userRole === '관리자') {
      // 모든 정보 조회
    } else if (userRole === '시설관리자') {
      const facilityNames = await getManagedFacilityNames(userFacilityName);
      query = query.in('facility_name', facilityNames);
    } else {
      query = query.eq('facility_name', userFacilityName);
    }

    const { data, error } = await query;
    if (error) throw error;

    const infos = (data || []).map(rowToEnergyInfo);
    res.json({ success: true, infos });
  } catch (err) {
    console.error('에너지 정보 조회 오류:', err);
    res.status(500).json({ success: false, message: '에너지 정보 조회 중 오류가 발생했습니다.' });
  }
});

app.post('/api/energy-info', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const { facilityName, energyType, customerNumber, bankName, accountNumber } = req.body;

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

    // 중복 확인
    const { data: existing } = await supabase
      .from('energy_info')
      .select('id')
      .eq('facility_name', facilityName.trim())
      .eq('energy_type', energyType.trim())
      .single();

    if (existing) {
      return res.json({ success: false, message: '해당 시설의 동일한 에너지 종류 정보가 이미 존재합니다.' });
    }

    const { error } = await supabase.from('energy_info').insert({
      facility_name: facilityName.trim(),
      energy_type: energyType.trim(),
      customer_number: customerNumber.trim(),
      bank_name: bankName.trim(),
      account_number: accountNumber.trim()
    });

    if (error) throw error;

    res.json({ success: true, message: '에너지 정보가 추가되었습니다.' });
  } catch (err) {
    console.error('에너지 정보 추가 오류:', err);
    res.status(500).json({ success: false, message: '에너지 정보 추가 중 오류가 발생했습니다.' });
  }
});

app.put('/api/energy-info/:id', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const targetId = parseInt(req.params.id);
    const { facilityName, energyType, customerNumber, bankName, accountNumber } = req.body;

    console.log('=== 에너지 정보 수정 요청 ===');
    console.log('ID:', targetId);
    console.log('요청 데이터:', { facilityName, energyType, customerNumber, bankName, accountNumber });

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

    // 대상 존재 확인
    const { data: targetInfo, error: fetchError } = await supabase
      .from('energy_info')
      .select('id')
      .eq('id', targetId)
      .single();

    if (fetchError || !targetInfo) {
      return res.json({ success: false, message: '유효하지 않은 ID입니다.' });
    }

    // 중복 확인 (자기 자신 제외)
    const { data: duplicate } = await supabase
      .from('energy_info')
      .select('id')
      .eq('facility_name', facilityName.trim())
      .eq('energy_type', energyType.trim())
      .neq('id', targetId)
      .single();

    if (duplicate) {
      return res.json({ success: false, message: '해당 시설의 동일한 에너지 종류 정보가 이미 존재합니다.' });
    }

    const { error } = await supabase.from('energy_info').update({
      facility_name: facilityName.trim(),
      energy_type: energyType.trim(),
      customer_number: customerNumber.trim(),
      bank_name: bankName.trim(),
      account_number: accountNumber.trim()
    }).eq('id', targetId);

    if (error) throw error;

    res.json({ success: true, message: '에너지 정보가 수정되었습니다.' });
  } catch (err) {
    console.error('에너지 정보 수정 오류:', err);
    res.status(500).json({ success: false, message: '에너지 정보 수정 중 오류가 발생했습니다.' });
  }
});

app.delete('/api/energy-info/:id', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const targetId = parseInt(req.params.id);

    const { data: targetInfo, error: fetchError } = await supabase
      .from('energy_info')
      .select('id')
      .eq('id', targetId)
      .single();

    if (fetchError || !targetInfo) {
      return res.json({ success: false, message: '유효하지 않은 ID입니다.' });
    }

    const { error } = await supabase.from('energy_info').delete().eq('id', targetId);
    if (error) throw error;

    res.json({ success: true, message: '에너지 정보가 삭제되었습니다.' });
  } catch (err) {
    console.error('에너지 정보 삭제 오류:', err);
    res.status(500).json({ success: false, message: '에너지 정보 삭제 중 오류가 발생했습니다.' });
  }
});

// ──────────────────────────────────────────────
// 데이터 조회 (필터) API
// ──────────────────────────────────────────────

app.get('/api/data-view', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const { facility, energyType, startDate, endDate } = req.query;
  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  try {
    let query = supabase.from('energy_records').select('*');

    // 역할별 기본 필터
    if (userRole === '관리자') {
      // 모든 데이터
    } else if (userRole === '시설관리자') {
      const facilityNames = await getManagedFacilityNames(userFacilityName);
      query = query.in('facility_name', facilityNames);
    } else {
      query = query.eq('facility_name', userFacilityName);
    }

    // 에너지 종류 필터
    if (energyType) {
      query = query.eq('energy_type', energyType);
    }

    // 기간 필터 (billing_month 기준)
    if (startDate) {
      const startMonth = startDate.substring(0, 7);
      query = query.gte('billing_month', startMonth);
    }
    if (endDate) {
      const endMonth = endDate.substring(0, 7);
      query = query.lte('billing_month', endMonth);
    }

    query = query.order('billing_month', { ascending: true });

    const { data, error } = await query;
    if (error) throw error;

    let records = (data || []).map(rowToEnergyRecord);

    // 시설 필터 (하위 시설 포함 - 클라이언트 측 필터링)
    if (facility) {
      // 선택한 시설의 하위 시설 목록 조회
      const { data: childData } = await supabase
        .from('users')
        .select('facility_name')
        .eq('parent_facility', facility);
      const childFacilities = (childData || []).map(u => u.facility_name);

      records = records.filter(record => {
        if (record.facilityName === facility) return true;
        if (childFacilities.includes(record.facilityName)) return true;
        if (record.facilityName.startsWith(facility + '(')) return true;
        return false;
      });
    }

    console.log('=== 데이터 조회 API 응답 ===');
    console.log('총 레코드 수:', records.length);
    if (records.length > 0) {
      console.log('첫 번째 레코드 샘플:', JSON.stringify(records[0], null, 2));
    }

    res.json({ success: true, records });
  } catch (err) {
    console.error('데이터 조회 오류:', err);
    res.status(500).json({ success: false, message: '데이터 조회 중 오류가 발생했습니다.' });
  }
});

// ──────────────────────────────────────────────
// 에너지 데이터 일괄 업로드 API
// ──────────────────────────────────────────────

app.post('/api/energy-data/upload', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const { records: newRecords } = req.body;

  if (!newRecords || !Array.isArray(newRecords) || newRecords.length === 0) {
    return res.status(400).json({ success: false, message: '업로드할 데이터가 없습니다.' });
  }

  try {
    const insertData = newRecords.map(record => ({
      facility_name: record.facilityName || '',
      billing_month: record.billingMonth || null,
      start_date: record.startDate || null,
      end_date: record.endDate || null,
      energy_type: record.energyType || '',
      usage_amount: parseFloat(record.usageAmount) || 0,
      usage_cost: parseFloat(record.usageCost) || 0,
      customer_number: record.customerNumber || '',
      bank_name: record.bankName || '',
      virtual_account: record.virtualAccount || ''
    }));

    const { error } = await supabase.from('energy_records').insert(insertData);
    if (error) throw error;

    res.json({ success: true, message: `${newRecords.length}건의 데이터가 업로드되었습니다.`, count: newRecords.length });
  } catch (err) {
    console.error('에너지 데이터 업로드 오류:', err);
    res.status(500).json({ success: false, message: '서버 오류가 발생했습니다.' });
  }
});

// ──────────────────────────────────────────────
// 선택된 에너지 데이터 삭제 API
// ──────────────────────────────────────────────

app.post('/api/energy-data/delete-multiple', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const { records: recordsToDelete } = req.body;

  if (!recordsToDelete || !Array.isArray(recordsToDelete) || recordsToDelete.length === 0) {
    return res.status(400).json({ success: false, message: '삭제할 레코드가 없습니다.' });
  }

  const userRole = req.session.user.role;
  const userFacilityName = req.session.user.facilityName;

  try {
    // 권한 체크용 시설 목록 준비
    let allowedFacilities = null;
    if (userRole === '시설관리자') {
      allowedFacilities = await getManagedFacilityNames(userFacilityName);
    } else if (userRole === '시설담당자') {
      allowedFacilities = [userFacilityName];
    }

    // 권한이 있는 레코드만 필터링
    const filteredRecords = recordsToDelete.filter(r => {
      if (allowedFacilities === null) return true; // 관리자
      return allowedFacilities.includes(r.facilityName);
    });

    if (filteredRecords.length === 0) {
      return res.status(403).json({ success: false, message: '삭제 권한이 없습니다.' });
    }

    // 레코드 필드 매칭으로 삭제 (id가 있으면 id 우선, 없으면 필드 매칭)
    let deletedCount = 0;
    for (const toDelete of filteredRecords) {
      if (toDelete.id) {
        const { error } = await supabase
          .from('energy_records')
          .delete()
          .eq('id', toDelete.id);
        if (!error) deletedCount++;
      } else {
        // id 없을 경우 필드 매칭으로 삭제
        const { data: matchRows } = await supabase
          .from('energy_records')
          .select('id')
          .eq('facility_name', toDelete.facilityName)
          .eq('start_date', toDelete.startDate)
          .eq('end_date', toDelete.endDate)
          .eq('energy_type', toDelete.energyType)
          .eq('usage_amount', parseFloat(toDelete.usageAmount))
          .eq('usage_cost', parseFloat(toDelete.usageCost));

        if (matchRows && matchRows.length > 0) {
          const ids = matchRows.map(r => r.id);
          const { error } = await supabase
            .from('energy_records')
            .delete()
            .in('id', ids);
          if (!error) deletedCount += ids.length;
        }
      }
    }

    console.log(`${deletedCount}개의 레코드가 삭제되었습니다.`);
    res.json({ success: true, message: `${deletedCount}개의 레코드가 삭제되었습니다.` });
  } catch (err) {
    console.error('데이터 삭제 오류:', err);
    res.status(500).json({ success: false, message: '데이터 삭제 중 오류가 발생했습니다.' });
  }
});

// ──────────────────────────────────────────────
// 모든 에너지 데이터 삭제 API (관리자 전용)
// ──────────────────────────────────────────────

app.delete('/api/energy-data/delete-all', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  const userRole = req.session.user.role;

  if (userRole !== '관리자') {
    return res.status(403).json({ success: false, message: '관리자만 모든 데이터를 삭제할 수 있습니다.' });
  }

  try {
    // 전체 건수 조회
    const { count } = await supabase
      .from('energy_records')
      .select('*', { count: 'exact', head: true });

    // 전체 삭제 (neq로 항상 true인 조건 사용)
    const { error } = await supabase
      .from('energy_records')
      .delete()
      .neq('id', 0);

    if (error) throw error;

    console.log(`관리자가 모든 에너지 데이터 삭제: ${count}개`);
    res.json({ success: true, message: `${count}개의 데이터가 삭제되었습니다.`, deletedCount: count });
  } catch (err) {
    console.error('모든 데이터 삭제 오류:', err);
    res.status(500).json({ success: false, message: '데이터 삭제 중 오류가 발생했습니다.' });
  }
});

// ──────────────────────────────────────────────
// 숫자 → 한글 변환 (공문 생성용)
// ──────────────────────────────────────────────

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

// ──────────────────────────────────────────────
// 공문 생성 API
// ──────────────────────────────────────────────

app.post('/api/generate-document', async (req, res) => {
  if (!req.session.user) {
    return res.status(401).json({ success: false, message: '인증이 필요합니다.' });
  }

  try {
    const record = req.body;

    let year, month;
    if (record.billingMonth) {
      const billingParts = record.billingMonth.split('-');
      year = parseInt(billingParts[0]);
      month = parseInt(billingParts[1]);
    } else {
      const startDate = new Date(record.startDate);
      year = startDate.getFullYear();
      month = startDate.getMonth() + 1;
    }

    const costNumber = parseInt(record.usageCost);
    const costKorean = numberToKorean(costNumber);

    let paymentMethod = '첨부 2 고지서 참조';
    if (record.bankName && record.virtualAccount) {
      paymentMethod = `계좌입금(${record.bankName} ${record.virtualAccount})`;
    } else if (record.virtualAccount) {
      paymentMethod = `계좌입금(${record.virtualAccount})`;
    }

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
            spacing: { after: 400 }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `${year}년 ${month}월 ${record.facilityName} ${record.energyType} 요금을 아래와 같이 납부하고자 합니다.`,
                size: 24
              })
            ],
            spacing: { after: 300 }
          }),
          new Paragraph({
            children: [new TextRun({ text: '', size: 24 })],
            spacing: { after: 200 }
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `  1. 소요금액: 금${costNumber.toLocaleString('ko-KR')}원(금${costKorean}원)`,
                size: 24
              })
            ],
            spacing: { after: 200 }
          }),
          new Paragraph({
            children: [new TextRun({ text: '  2. 세부내역: 첨부 1 참조', size: 24 })],
            spacing: { after: 200 }
          }),
          new Paragraph({
            children: [new TextRun({ text: `  3. 납부방법: ${paymentMethod}`, size: 24 })],
            spacing: { after: 200 }
          }),
          new Paragraph({
            children: [new TextRun({ text: '  4. 예산과목: ', size: 24 })],
            spacing: { after: 400 }
          })
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = `${year}-${String(month).padStart(2, '0')}-${record.facilityName}-${record.energyType}-공문.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.send(buffer);
  } catch (error) {
    console.error('공문 생성 오류:', error);
    res.status(500).json({ success: false, message: '공문 생성 중 오류가 발생했습니다.' });
  }
});

// ──────────────────────────────────────────────
// 첨부1 생성 API (Excel) - 첨부2 서식 사용
// ──────────────────────────────────────────────

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

    // Supabase에서 에너지 정보 조회 (금융기관/계좌번호)
    const { data: energyInfoData } = await supabase
      .from('energy_info')
      .select('*');
    const energyInfos = (energyInfoData || []).map(rowToEnergyInfo);

    let matchingInfo = energyInfos.find(info =>
      info.facilityName === record.facilityName &&
      info.energyType === record.energyType
    );
    // 상위 시설명으로 폴백 조회
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

    // 템플릿 파일 읽기
    console.log('템플릿 파일 경로:', ATTACHMENT_TEMPLATE_FILE);
    if (!fs.existsSync(ATTACHMENT_TEMPLATE_FILE)) {
      console.error('템플릿 파일이 존재하지 않습니다:', ATTACHMENT_TEMPLATE_FILE);
      return res.status(500).json({ success: false, message: '템플릿 파일을 찾을 수 없습니다.' });
    }
    const workbook = xlsx.readFile(ATTACHMENT_TEMPLATE_FILE);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    worksheet['A1'] = { t: 's', v: '세부내역' };

    const dataRow = 6;

    worksheet['A' + dataRow] = { t: 'n', v: year };
    worksheet['B' + dataRow] = { t: 'n', v: month };
    worksheet['C' + dataRow] = { t: 's', v: record.facilityName || '' };
    worksheet['D' + dataRow] = { t: 's', v: bankName };
    worksheet['E' + dataRow] = { t: 's', v: accountNumber };
    worksheet['F' + dataRow] = { t: 's', v: record.startDate || '' };
    worksheet['G' + dataRow] = { t: 's', v: '~' };
    worksheet['H' + dataRow] = { t: 's', v: record.endDate || '' };
    worksheet['I' + dataRow] = { t: 'n', v: parseFloat(record.usageCost) || 0, z: '#,##0' };
    worksheet['J' + dataRow] = { t: 's', v: record.energyType || '' };

    worksheet['!ref'] = xlsx.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: dataRow, c: 9 }
    });

    worksheet['!cols'] = [
      { wch: 8 }, { wch: 6 }, { wch: 25 }, { wch: 12 }, { wch: 18 },
      { wch: 12 }, { wch: 3 }, { wch: 12 }, { wch: 15 }, { wch: 12 }
    ];

    const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
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

// ──────────────────────────────────────────────
// 통합 첨부문서 생성 API
// ──────────────────────────────────────────────

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

    // Supabase에서 에너지 정보 조회
    const { data: energyInfoData } = await supabase
      .from('energy_info')
      .select('*');
    const energyInfos = (energyInfoData || []).map(rowToEnergyInfo);

    if (!fs.existsSync(ATTACHMENT_TEMPLATE_FILE)) {
      console.error('템플릿 파일이 존재하지 않습니다:', ATTACHMENT_TEMPLATE_FILE);
      return res.status(500).json({ success: false, message: '템플릿 파일을 찾을 수 없습니다.' });
    }
    const workbook = xlsx.readFile(ATTACHMENT_TEMPLATE_FILE);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    worksheet['A1'] = { t: 's', v: '세부내역' };

    const totalCost = records.reduce((sum, r) => sum + (parseFloat(r.usageCost) || 0), 0);
    worksheet['I5'] = { t: 'n', v: totalCost, z: '#,##0' };

    const dataStartRow = 6;

    records.forEach((record, index) => {
      const currentRow = dataStartRow + index;

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

      let matchingInfo = energyInfos.find(info =>
        info.facilityName === record.facilityName &&
        info.energyType === record.energyType
      );
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

      worksheet['A' + currentRow] = { t: 'n', v: year };
      worksheet['B' + currentRow] = { t: 'n', v: month };
      worksheet['C' + currentRow] = { t: 's', v: record.facilityName || '' };
      worksheet['D' + currentRow] = { t: 's', v: bankName };
      worksheet['E' + currentRow] = { t: 's', v: accountNumber };
      worksheet['F' + currentRow] = { t: 's', v: record.startDate || '' };
      worksheet['G' + currentRow] = { t: 's', v: '~' };
      worksheet['H' + currentRow] = { t: 's', v: record.endDate || '' };
      worksheet['I' + currentRow] = { t: 'n', v: parseFloat(record.usageCost) || 0, z: '#,##0' };
      worksheet['J' + currentRow] = { t: 's', v: record.energyType || '' };
    });

    const lastDataRow = dataStartRow + records.length - 1;
    worksheet['!ref'] = xlsx.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: lastDataRow, c: 9 }
    });

    worksheet['!cols'] = [
      { wch: 8 }, { wch: 6 }, { wch: 25 }, { wch: 12 }, { wch: 18 },
      { wch: 12 }, { wch: 3 }, { wch: 12 }, { wch: 15 }, { wch: 12 }
    ];

    const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });
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

// ──────────────────────────────────────────────
// 네이버 클로바 OCR API 프록시
// ──────────────────────────────────────────────

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

    const cleanApiUrl = apiUrl.replace(/\/$/, '');
    console.log('정리된 API URL:', cleanApiUrl);

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

    const base64Data = imageBase64.split(',')[1];
    if (!base64Data) {
      throw new Error('올바르지 않은 Base64 데이터');
    }
    const imageBuffer = Buffer.from(base64Data, 'base64');
    console.log('이미지 버퍼 크기:', imageBuffer.length, 'bytes');

    const FormData = require('form-data');
    const axios = require('axios');

    const formData = new FormData();

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

    console.log('클로바 API 호출 중...');
    console.log('최종 요청 URL:', cleanApiUrl);

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

    let extractedText = '';
    if (response.data && response.data.images && response.data.images[0]) {
      const fields = response.data.images[0].fields || [];
      console.log('추출된 필드 개수:', fields.length);
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

// ──────────────────────────────────────────────
// 클로바 API 키 검증 엔드포인트
// ──────────────────────────────────────────────

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

    if (!apiUrl.includes('apigw.ntruss.com')) {
      return res.json({
        success: false,
        message: 'API URL 형식이 올바르지 않습니다.\n네이버 클라우드의 APIGW Invoke URL을 입력하세요.\n(예: https://xxxxx.apigw.ntruss.com/custom/v1/xxxxx/xxxxxxxx)'
      });
    }

    if (!apiUrl.startsWith('https://')) {
      return res.json({
        success: false,
        message: 'API URL은 https://로 시작해야 합니다.\n전체 URL을 복사하여 붙여넣으세요.'
      });
    }

    const cleanApiUrl = apiUrl.replace(/\/$/, '');
    console.log('정리된 API URL:', cleanApiUrl);

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

// ──────────────────────────────────────────────
// 정적 파일 및 서버 시작
// ──────────────────────────────────────────────

app.get('/', (_req, res) => {
  res.sendFile(__dirname + '/public/index.html');
});

if (process.env.NODE_ENV !== 'production') {
  app.listen(PORT, () => {
    console.log(`에너지 관리 시스템이 http://localhost:${PORT} 에서 실행중입니다.`);
  });
}

module.exports = app;
