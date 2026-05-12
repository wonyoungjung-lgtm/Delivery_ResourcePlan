const SS = SpreadsheetApp.getActiveSpreadsheet();
const EXCLUDED_DEPTS = ['-', '사업부', 'Staff'];

let _cachedInputs = null;
let _cachedHR = null;

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Delivery Resource Manager');
}

function formatDateStr(d) {
  if (!d) return '';
  let dt = new Date(d);
  return isNaN(dt) ? d : dt.getFullYear() + '-' + String(dt.getMonth()+1).padStart(2,'0') + '-' + String(dt.getDate()).padStart(2,'0');
}

// 🌟 1. 타임라인을 2026년 1월~12월로 고정
function generateTimeline() {
  const year = 2026;
  const today = new Date();
  const currY = today.getFullYear();
  const currM = today.getMonth() + 1;
  
  let timeline = [];
  for (let m = 1; m <= 12; m++) {
    // UI 로직(과거/당월/예정)을 위해 현재 날짜 기준 offset 계산
    const offset = (year - currY) * 12 + (m - currM);
    timeline.push({ 
      year: year, 
      month: m, 
      offset: offset, 
      label: `26/${String(m).padStart(2,'0')}` 
    });
  }
  return timeline;
}

/**
 * 모든 데이터를 미리 계산하여 시트에 저장하는 함수
 * (트리거를 설정하여 10분마다 실행하거나, 데이터 저장 시마다 호출)
 */
function rebuildPhysicalCache() {
  const cacheSheet = SS.getSheetByName('_DB_CACHE_') || SS.insertSheet('_DB_CACHE_');
  cacheSheet.clear();
  
  // 1. 헤더 설정
  cacheSheet.getAppendRow(['CacheKey', 'JSONData', 'LastUpdated']);
  
  // 2. 대시보드 데이터 캐싱 (2026년 기준)
  const dashboardData = getHomeDashboardData(2026, 5); // 실제 계산 로직 실행
  const dashboardJson = JSON.stringify(dashboardData);
  
  // 3. 시트에 저장 (셀 하나당 5만자 제한이 있으므로 데이터가 크면 분할 저장 로직 필요)
  cacheSheet.appendRow(['HOME_DASHBOARD_2026', dashboardJson, new Date()]);
  
  console.log("물리적 캐시 업데이트 완료");
}

/**
 * 캐시된 데이터를 가져오는 함수 (비약적으로 빠름)
 */
function getCachedData(key) {
  const sheet = SS.getSheetByName('_DB_CACHE_');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      return JSON.parse(data[i][1]); // 계산 없이 JSON만 파싱해서 반환
    }
  }
  return null;
}


function getInitialData() {
  const email = Session.getActiveUser().getEmail();
  
  // 1. RBAC 정보는 개인별로 다르므로 실시간 조회 (데이터가 작아 빠름)
  const rbacSheet = SS.getSheetByName('RBAC');
  const rbacData = rbacSheet.getDataRange().getValues();
  let user = { email: email, role: 'Guest', scope: {} };
  const userRow = rbacData.find(row => String(row[0]).trim().toLowerCase() === email.toLowerCase());
  
  if (userRow) {
    user.role = String(userRow[2]).trim().toLowerCase();
    user.scope = { dept: userRow[4], div: userRow[5], unit: userRow[6], team: userRow[7] };
  }

  // 2. 물리 캐시에서 공통 데이터(Dashboard, Tree, Filters) 가져오기
  // 'GLOBAL_CACHE_2026'이라는 키로 통째로 저장된 데이터를 읽습니다.
  let cachedFullData = getCachedData('GLOBAL_CACHE_2026');

  // 3. 캐시가 없으면 (최초 실행 또는 캐시 만료 시) 생성 로직 실행
  if (!cachedFullData) {
    console.log("캐시 없음: 실시간 계산 시작");
    cachedFullData = rebuildAndReturnCache(); 
  }

  // 4. 사용자 정보와 캐시된 공통 데이터를 합쳐서 반환
  return {
    user: user,
    dashboard: cachedFullData.dashboard,
    fullTree: cachedFullData.fullTree,
    orgFilters: cachedFullData.orgFilters,
    latestYear: 2026,
    latestMonth: 5 // 혹은 오늘 날짜 기준
  };
}

/**
 * 무거운 연산을 수행하고 결과를 캐시 시트에 저장한 후 반환
 */
function rebuildAndReturnCache() {
  const targetYear = 2026;
  const targetMonth = 5;
  
  // A. 대시보드 통계 계산 (기존 최적화 함수 호출)
  const dashboard = getHomeDashboardData(targetYear, targetMonth);
  
  // B. 조직 트리 및 필터 생성
  const hrData = SS.getSheetByName('HR_Report').getDataRange().getValues();
  hrData.shift();
  const currentMonthHR = hrData.filter(r => Number(r[0]) === targetYear && Number(r[1]) === targetMonth && !EXCLUDED_DEPTS.includes(r[8]));
  
  const tree = buildFlatTree(currentMonthHR, { role: 'admin' }); // 캐시용은 전체 트리 생성
  const orgFilters = ['전사'];
  for(let dept in tree) {
    for(let div in tree[dept]) {
      if(div !== '-') {
        if(!orgFilters.includes(div)) orgFilters.push(div);
        for(let unit in tree[dept][div]) {
          if(unit !== '-' && unit !== '(실 미소속)' && !orgFilters.includes(unit)) orgFilters.push(unit);
        }
      }
    }
  }

  const finalCache = {
    dashboard: dashboard,
    fullTree: tree,
    orgFilters: orgFilters
  };

  // 캐시 시트에 저장
  saveToCacheSheet('GLOBAL_CACHE_2026', finalCache);
  
  return finalCache;
}

function saveToCacheSheet(key, obj) {
  const sheet = SS.getSheetByName('_DB_CACHE_') || SS.insertSheet('_DB_CACHE_');
  const jsonStr = JSON.stringify(obj);
  const now = new Date();
  
  // 기존 키 삭제 후 새로 삽입
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === key) sheet.deleteRow(i + 1);
  }
  
  sheet.appendRow([key, jsonStr, now]);
}

function getCachedData(key) {
  const sheet = SS.getSheetByName('_DB_CACHE_');
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const row = data.find(r => r[0] === key);
  
  if (row && row[1]) {
    try {
      return JSON.parse(row[1]);
    } catch (e) {
      return null;
    }
  }
  return null;
}

/**
 * 최적화: 시트 데이터를 한 번만 읽어 전역 변수에 저장
 */
function getGlobalData() {
  if (_cachedInputs && _cachedHR) return { inputs: _cachedInputs, hr: _cachedHR };

  const targetYear = 2026;
  const sheets = {
    ex: SS.getSheetByName('투입실적_실행예산'),
    mm: SS.getSheetByName('투입실적_MM등록'),
    fc: SS.getSheetByName('투입계획_FCST'),
    hr: SS.getSheetByName('HR_Report')
  };

  // 모든 데이터를 한 번에 Get
  const rawEx = sheets.ex ? sheets.ex.getDataRange().getValues().slice(1) : [];
  const rawMm = sheets.mm ? sheets.mm.getDataRange().getValues().slice(1) : [];
  const rawFc = sheets.fc ? sheets.fc.getDataRange().getValues().slice(1) : [];
  const rawHr = sheets.hr ? sheets.hr.getDataRange().getValues().slice(1) : [];

  // 1. Inputs 맵핑 (기존 getUnifiedInputs 로직을 인메모리 처리로 변경)
  const map = { exec: {}, mm: {}, plan: {} };
  
  const process = (data, targetMap, mmCol, idCol, projCol, wbsCol) => {
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (Number(r[0]) !== targetYear) continue;
      const key = `${r[0]}-${r[1]}-${String(r[idCol]).trim()}`;
      const val = parseFloat(r[mmCol] || 0);
      if (val === 0) continue;
      if (!targetMap[key]) targetMap[key] = { total: 0, details: [] };
      targetMap[key].total += val;
      targetMap[key].details.push({ proj: r[projCol], wbs: r[wbsCol] || '-', mm: val });
    }
  };

  process(rawEx, map.exec, 2, 3, 6, 5);
  process(rawMm, map.mm, 2, 3, 6, 5);
  process(rawFc, map.plan, 5, 2, 4, 8);
  
  _cachedInputs = map;

  // 2. HR 데이터 인덱싱 (연-월별로 그룹화하여 필터링 속도 개선)
  const hrMap = {};
  rawHr.forEach(r => {
    if (Number(r[0]) !== targetYear || EXCLUDED_DEPTS.includes(r[8])) return;
    const key = `${r[0]}-${r[1]}`;
    if (!hrMap[key]) hrMap[key] = [];
    hrMap[key].push(r);
  });
  _cachedHR = hrMap;

  return { inputs: _cachedInputs, hr: _cachedHR };
}

function buildFlatTree(data, user) {
  const tree = {};
  data.forEach(row => {
    const dept = row[8], div = row[9] || '-', unit = row[10] || '-', team = row[11] || '-';
    if (!dept || EXCLUDED_DEPTS.includes(dept)) return;
    
    if (user.role !== 'admin') {
      if (user.scope.div && div !== user.scope.div) return;
      if (user.scope.unit && unit !== user.scope.unit) return;
    }

    if (!tree[dept]) tree[dept] = {};
    if (!tree[dept][div]) tree[dept][div] = {};
    if (!tree[dept][div][unit]) tree[dept][div][unit] = [];
    if (team !== '-' && !tree[dept][div][unit].includes(team)) tree[dept][div][unit].push(team);
  });
  
  for (let dept in tree) {
    for (let div in tree[dept]) {
      for (let unit in tree[dept][div]) tree[dept][div][unit].sort((a, b) => a.localeCompare(b));
    }
  }
  return tree;
}

// 🌟 2. 데이터 수집 시 2026년만 필터링하여 메모리 및 속도 최적화
function getUnifiedInputs() {
  const targetYear = 2026; // 🌟 2026년 데이터만 읽도록 고정
  const ex = SS.getSheetByName('투입실적_실행예산') ? SS.getSheetByName('투입실적_실행예산').getDataRange().getValues().slice(1) : [];
  const mm = SS.getSheetByName('투입실적_MM등록') ? SS.getSheetByName('투입실적_MM등록').getDataRange().getValues().slice(1) : [];
  const fc = SS.getSheetByName('투입계획_FCST') ? SS.getSheetByName('투입계획_FCST').getDataRange().getValues().slice(1) : [];
  
  const map = { exec: {}, mm: {}, plan: {} };

  function populate(dataArray, mapObj, mmCol, idCol, projCol, wbsCol) {
    dataArray.forEach(r => {
      // 🌟 연도 체크를 가장 먼저 수행하여 불필요한 연산 방지
      if (Number(r[0]) !== targetYear) return;
      
      let y = Number(r[0]), m = Number(r[1]), id = String(r[idCol]).trim();
      let val = parseFloat(r[mmCol] || 0);
      if (val === 0) return;
      
      let key = `${y}-${m}-${id}`;
      if (!mapObj[key]) mapObj[key] = { total: 0, details: [] };
      mapObj[key].total += val;
      mapObj[key].details.push({ 
        proj: r[projCol], 
        wbs: r[wbsCol] || '-', 
        mm: val 
      });
    });
  }
  // 실적 시트들은 F열(5)이 WBS
  populate(ex, map.exec, 2, 3, 6, 5); 
  populate(mm, map.mm, 2, 3, 6, 5);
  
  // FCST 시트는 마지막 열(인덱스 8)에 WBS가 있다고 가정 (아래 saveBulkPlan에서 추가함)
  populate(fc, map.plan, 5, 2, 4, 8); 

  return map;
}

/**
 * 홈 대시보드용 데이터 생성 (트렌드 차트 + 계층형 피벗 테이블)
 */
function getHomeDashboardData(year, month) {
  const { inputs, hr } = getGlobalData();
  const timeline = generateTimeline();
  let stats = { timelineLabels: timeline.map(t => t.label), trendData: {}, pivotTable: [], debugLogs: [] };

  // --- [1. 트렌드 차트 데이터 계산] ---
  // 이 부분에서 더 이상 hrData.filter를 반복 호출하지 않음
  timeline.forEach((t, idx) => {
    const tHR = hr[`${t.year}-${t.month}`] || [];
    let monthlyOrgSums = {};
    
    tHR.forEach(res => {
      if (Number(res[13]) === 0) return; // SGA 제외
      const empId = String(res[3]).trim();
      const enrollMM = typeof res[2] === 'string' ? parseFloat(res[2].replace('%',''))/100 : parseFloat(res[2]||1);
      const k = `${t.year}-${t.month}-${empId}`;
      
      let total = (inputs.exec[k]?.total || 0) + (inputs.mm[k]?.total || 0) + (t.offset < 0 ? 0 : (inputs.plan[k]?.total || 0));
      if (res[12] === 'Shared' && t.offset >= 0) total = enrollMM;

      const levels = ['전사', res[8], res[9], res[10], res[11]].filter(v => v && v !== '-');
      levels.forEach(org => {
        if(!monthlyOrgSums[org]) monthlyOrgSums[org] = { in: 0, en: 0 };
        monthlyOrgSums[org].in += total;
        monthlyOrgSums[org].en += enrollMM;
      });
    });

    for(let org in monthlyOrgSums) {
      if(!stats.trendData[org]) stats.trendData[org] = new Array(12).fill(0);
      stats.trendData[org][idx] = monthlyOrgSums[org].en > 0 ? ((monthlyOrgSums[org].in / monthlyOrgSums[org].en) * 100).toFixed(1) : 0;
    }
  });

  // --- [2. 피벗 테이블 (당월/전월 비교)] ---
  const currHR = hr[`${year}-${month}`] || [];
  const pmDate = new Date(year, month - 2, 1);
  const prevHR = hr[`${pmDate.getFullYear()}-${pmDate.getMonth() + 1}`] || [];

  stats.debugLogs.push(`[MoM] 당월:${currHR.length}명, 전월:${prevHR.length}명`);

  const nodes = {};

  /**
   * 트리 노드 생성 및 카운트 가산 함수
   */
  function processHR(data, isCurr) {
    data.forEach(res => {
      const d = res[8], v = res[9], u = res[10], t = res[11];
      if (!d || d === '-') return;

      // 인력 유형 분류
      let type = Number(res[13]) === 0 ? 'sga' : (res[12] === 'Shared' ? 'shared' : (res[12] === '휴직' ? 'leave' : 'cos'));
      let mem = { name: res[4], id: String(res[3]).trim(), pos: res[6] };

      const paths = [];
      // 1단계: 부문
      const dId = `D_${d}`; 
      paths.push({ id: dId, name: d, lvl: 0, pId: null });
      
      // 2단계: 본부
      let vId = (v && v !== '-') ? `${dId}_V_${v}` : dId; 
      if(vId !== dId) paths.push({ id: vId, name: v, lvl: 1, pId: dId });

      // 3단계: 실 (🌟 실이 '-'인 경우 '(실 미소속)'으로 가상 경로 생성)
      let unitName = (u && u !== '-') ? u : '(실 미소속)';
      let uId = `${vId}_U_${unitName}`;
      paths.push({ id: uId, name: unitName, lvl: 2, pId: vId });

      // 4단계: 팀
      if(t && t !== '-') {
        let tId = `${uId}_T_${t}`;
        paths.push({ id: tId, name: t, lvl: 3, pId: uId });
      }

      // 각 경로 노드에 인원수 누적
      paths.forEach(p => {
        if(!nodes[p.id]) {
          nodes[p.id] = { 
            id: p.id, name: p.name, level: p.lvl, parentId: p.pId, 
            curr: {t:0, c:0, sh:0, sg:0, lv:0, cL:[], shL:[], sgL:[], lvL:[]}, 
            prev: {t:0, c:0, sh:0, sg:0, lv:0} 
          };
        }
        let target = isCurr ? nodes[p.id].curr : nodes[p.id].prev;
        target.t++;
        if(type === 'cos') { target.c++; if(isCurr) nodes[p.id].curr.cL.push(mem); }
        else if(type === 'shared') { target.sh++; if(isCurr) nodes[p.id].curr.shL.push(mem); }
        else if(type === 'leave') { target.lv++; if(isCurr) nodes[p.id].curr.lvL.push(mem); }
        else { target.sg++; if(isCurr) nodes[p.id].curr.sgL.push(mem); }
      });
    });
  }

  processHR(currHR, true);
  processHR(prevHR, false);

  // 결과 정렬 (ID 순으로 정렬하면 트리 구조 순서가 유지됨)
  let pivotArr = Object.values(nodes).filter(n => n.curr.t > 0 || n.prev.t > 0);
  pivotArr.sort((a,b) => a.id.localeCompare(b.id));
  
  stats.pivotTable = pivotArr;
  return stats;
}

// 🌟 배열 전체 스캔(filter) 제거, 캐싱된 Key로 즉시 조회
function getCellDetails(empId, year, month) {
  const inputs = getUnifiedInputs();
  const key = `${Number(year)}-${Number(month)}-${String(empId).trim()}`;
  
  let logs = [
    `[검색 조건] Key: '${key}'`,
    `[캐시 조회 결과] 실행예산: ${inputs.exec[key] ? '발견' : '없음'} / MM등록: ${inputs.mm[key] ? '발견' : '없음'} / FCST: ${inputs.plan[key] ? '발견' : '없음'}`
  ];

  return {
    execs: inputs.exec[key] ? inputs.exec[key].details : [],
    mms: inputs.mm[key] ? inputs.mm[key].details : [],
    plans: inputs.plan[key] ? inputs.plan[key].details : [],
    debugLogs: logs
  };
}

/**
 * 조직 상세 뷰 데이터 조회 (시계열 10개월)
 * @param {Object} path - {div: "본부명", unit: "실명", team: "팀명"}
 * @param {Number} year - 조회 연도
 * @param {Number} month - 조회 월
 */
function getOrgResources(path, year, month) {
  // year, month 매개변수가 오더라도 내부적으로 generateTimeline()이 2026년을 반환함
  const timeline = generateTimeline(); 
  const inputs = getUnifiedInputs();
  const hrData = SS.getSheetByName('HR_Report').getDataRange().getValues().slice(1);

  const hrMap = {}; 
  const targetEmpIds = new Set();
  const targetYear = 2026;

  hrData.forEach(r => {
    if (EXCLUDED_DEPTS.includes(r[8])) return;
    const y = Number(r[0]), m = Number(r[1]);
    
    // 🌟 2026년 데이터가 아니면 스킵
    if (y !== targetYear) return; 
    
    const key = `${y}-${m}`;
    const eId = String(r[3]).trim();
    const rec = {
      y, m, empId: eId, name: r[4], grade: r[5], pos: r[6],
      div: r[9], 
      unit: (r[10] === '-' || !r[10]) ? '(실 미소속)' : r[10],
      team: r[11], 
      type: String(r[12]).trim(), 
      isTarget: Number(r[13]), 
      date: formatDateStr(r[14]) 
    };
    hrMap[`${key}-${eId}`] = rec;
    if (pathMatch(rec, path)) targetEmpIds.add(eId);
  });

  // 🌟 2단계: 그룹 초기화 (요청하신 순서 및 명칭)
  const groups = { 
    'COS - 투입 확정': [], 
    'COS - 부분 투입': [], 
    'COS - 투입 예정': [], 
    'COS - 미투입': [], 
    'Shared': [], 
    '휴직': [], 
    'SGA': [], 
    '전출/퇴직': [] 
  };

  const currentKeyPrefix = `${Number(year)}-${Number(month)}`;

  targetEmpIds.forEach(id => {
    let currentHR = hrMap[`${currentKeyPrefix}-${id}`];
    // 🌟 수정: path 인자 추가
    let isCurrentHere = pathMatch(currentHR, path); 

    let latestHR = currentHR;
    if (!latestHR) {
      for (let i = timeline.length - 1; i >= 0; i--) {
        let temp = hrMap[`${timeline[i].year}-${timeline[i].month}-${id}`];
        // 🌟 수정: path 인자 추가
        if (pathMatch(temp, path)) { latestHR = temp; break; } 
      }
    }
    if (!latestHR) return;

    let m = { ...latestHR };
    m.statusLabel = isCurrentHere ? (m.date ? `발령일자: ${m.date}` : '') : `<span class="text-danger fw-bold">전출</span>`;

    // 🌟 3단계: 10개월 타임라인별 MM 합계 산출 및 프로젝트명/WBS 수집
    let confProjs = new Set(), confWbs = new Set();
    let planProjs = new Set(), planWbs = new Set();


    m.timelineMM = timeline.map(t => {
      let tHR = hrMap[`${t.year}-${t.month}-${id}`];
      // 🌟 수정: path 인자 추가
      if (!pathMatch(tHR, path)) return { year: t.year, month: t.month, offset: t.offset, confirmed: null, planned: null, total: null };
      
      let k = `${t.year}-${t.month}-${id}`;
      let eMM = 0, mMM = 0, pMM = 0;

      // 1. 실행예산 (확정 데이터)
      if (inputs.exec[k]) {
        eMM = inputs.exec[k].total;
        inputs.exec[k].details.forEach(d => { 
          if(d.proj && d.proj !== '-') confProjs.add(d.proj); 
          if(d.wbs && d.wbs !== '-') confWbs.add(d.wbs);
        });
      }
      
      // 2. MM등록 (확정 데이터)
      if (inputs.mm[k]) {
        mMM = inputs.mm[k].total;
        inputs.mm[k].details.forEach(d => { 
          if(d.proj && d.proj !== '-') confProjs.add(d.proj); 
          if(d.wbs && d.wbs !== '-') confWbs.add(d.wbs);
        });
      }
      
      // 3. FCST (계획 데이터)
      if (inputs.plan[k]) {
        pMM = inputs.plan[k].total;
        inputs.plan[k].details.forEach(d => { 
          if(d.proj && d.proj !== '-') planProjs.add(d.proj); 
          if(d.wbs && d.wbs !== '-') planWbs.add(d.wbs);
        });
      }
      
      return { 
        year: t.year, 
        month: t.month, 
        offset: t.offset, 
        confirmed: eMM + mMM, 
        planned: pMM,
        total: (eMM + mMM + pMM) 
      };
    });

    // 🌟 프로젝트/WBS 이름 축약 처리 (툴팁용 전체 텍스트와 화면용 축약 텍스트 분리)
    function getSummary(setObj) {
      if (setObj.size === 0) return { full: '-', short: '-' };
      const arr = Array.from(setObj);
      return {
        full: arr.join(', '), // 마우스 올렸을 때 볼 전체 내용
        short: arr.length > 1 ? `${arr[0]} 외 ${arr.length - 1}건` : arr[0] // 화면에 보일 축약 내용
      };
    }
    
    const confP = getSummary(confProjs);
    m.confProjFull = confP.full; m.confProjShort = confP.short;
    
    const confW = getSummary(confWbs);
    m.confWbsFull = confW.full; m.confWbsShort = confW.short;
    
    const planP = getSummary(planProjs);
    m.planProjFull = planP.full; m.planProjShort = planP.short;
    
    const planW = getSummary(planWbs);
    m.planWbsFull = planW.full; m.planWbsShort = planW.short;

    // 🌟 4단계: 우선순위에 따른 섹션 분류
    if (!isCurrentHere) {
        groups['전출/퇴직'].push(m);
    } else if (m.type === '휴직') {
        groups['휴직'].push(m);
    } else if (m.isTarget === 0) {
        groups['SGA'].push(m);
    } else if (m.type === 'Shared') {
        groups['Shared'].push(m);
    } else {
        // COS 세부 분류
        let k = `${Number(year)}-${Number(month)}-${id}`;
        let eMM = (inputs.exec[k] ? inputs.exec[k].total : 0) + (inputs.mm[k] ? inputs.mm[k].total : 0);
        let pMM = inputs.plan[k] ? inputs.plan[k].total : 0;
        
        if (eMM >= 1.0) groups['COS - 투입 확정'].push(m);
        else if (eMM > 0) groups['COS - 부분 투입'].push(m);
        else if (eMM === 0 && pMM > 0) groups['COS - 투입 예정'].push(m);
        else groups['COS - 미투입'].push(m);
    }
  });

  // 인원이 없는 그룹은 UI에서 보이지 않도록 제거
  for(let g in groups) { if(groups[g].length === 0) delete groups[g]; }

  return { timeline, grouped: groups };
}

function saveBulkPlan(payload) {
  const inputs = getUnifiedInputs(); // 🌟 기존 투입 데이터(확정+계획) 불러오기
  const sheet = SS.getSheetByName('투입계획_FCST');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // WBS 컬럼이 없으면(9번째 열) 추가
  if (headers.length < 9 || headers[8] !== 'WBS') {
    sheet.getRange(1, 9).setValue('WBS');
  }

  let start = new Date(payload.startY, payload.startM - 1, 1);
  let end = new Date(payload.endY, payload.endM - 1, 1);
  let newMM = parseFloat(payload.mm);
  
  // 🚨 1단계: 실제 저장 전, 지정된 기간의 투입 공수 1.0 초과 여부 사전 검증
  let validationCurrent = new Date(start);
  while(validationCurrent <= end) {
    let y = validationCurrent.getFullYear();
    let m = validationCurrent.getMonth() + 1;
    let key = `${y}-${m}-${payload.empId}`;
    
    let existingMM = 0;
    if(inputs.exec[key]) existingMM += inputs.exec[key].total;
    if(inputs.mm[key]) existingMM += inputs.mm[key].total;
    if(inputs.plan[key]) existingMM += inputs.plan[key].total;
    
    // 자바스크립트 부동소수점 오류(0.1+0.2=0.3000004) 방지를 위해 소수점 첫째 자리에서 반올림 후 비교
    let totalMM = Math.round((existingMM + newMM) * 10) / 10;
    
    if (totalMM > 1.0) {
      // 1.0을 초과하면 에러를 발생시켜 실행을 즉시 중단함
      throw new Error(`${y}년 ${m}월의 총 투입 공수가 1.0을 초과합니다.\n(현재 확정+계획: ${existingMM.toFixed(1)}MM / 추가 입력: ${newMM.toFixed(1)}MM)`);
    }
    validationCurrent.setMonth(validationCurrent.getMonth() + 1);
  }

  // 🌟 2단계: 검증을 무사히 통과했다면 실제 저장 로직 실행
  let rowsToAppend = [];
  let current = new Date(start);
  
  while(current <= end) {
    rowsToAppend.push([
      current.getFullYear(), 
      current.getMonth()+1, 
      payload.empId, 
      payload.empName, 
      payload.proj, 
      payload.mm, 
      Session.getActiveUser().getEmail(), 
      formatDateStr(new Date()),
      payload.wbs // 9번째 열에 WBS 저장
    ]);
    current.setMonth(current.getMonth() + 1);
  }
  
  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 9).setValues(rowsToAppend);
  }
  
  return "저장되었습니다.";
}

/**
 * 선택된 조직 경로와 인력의 소속 정보가 일치하는지 확인
 */
function pathMatch(res, path) {
  if (!res) return false;
  // 필터가 없거나 '전사'인 경우 모든 데이터 허용
  if (!path || Object.keys(path).length === 0) return true;
  
  const level = Object.keys(path)[0];
  const val = path[level];
  
  if (!level || !val || val === '전사') return true;
  
  // 현재 인력 정보(res)의 해당 레벨(div, unit, team) 값이 필터값과 일치하는지 확인
  return res[level] === val;
}
