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

// 🌟 1. 청크 분할 캐싱 지원 (팀 데이터는 시트 백업 제외하여 속도 극대화)
function setGlobalCache(key, obj) {
  const jsonStr = JSON.stringify(obj);
  const cache = CacheService.getScriptCache();
  
  // 1-1. 메모리 캐시 (가장 빠름)
  try {
    if (jsonStr.length < 90000) {
      cache.put(key, jsonStr, 21600);
      cache.remove(key + '_chunks');
    } else {
      const chunkSize = 90000;
      const numChunks = Math.ceil(jsonStr.length / chunkSize);
      for (let i = 0; i < numChunks; i++) {
        cache.put(`${key}_${i}`, jsonStr.substring(i * chunkSize, (i + 1) * chunkSize), 21600);
      }
      cache.put(key + '_chunks', numChunks.toString(), 21600);
    }
  } catch(e) { console.error("Cache put error:", e); }
  
  // 🌟 [핵심 수정] 개별 팀 투입현황 데이터는 무거운 시트 백업을 생략함! (병목 완벽 제거)
  if (key.startsWith('ORG_RES_')) {
    return;
  }
  
  // 1-2. 공통 글로벌 데이터만 시트 백업
  const sheet = SS.getSheetByName('_DB_CACHE_') || SS.insertSheet('_DB_CACHE_');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === key) sheet.deleteRow(i + 1);
  }
  sheet.appendRow([key, jsonStr, new Date()]);
}

function getGlobalCache(key) {
  const cache = CacheService.getScriptCache();
  
  // 2-1. 청크(쪼개진) 데이터 확인
  const numChunksStr = cache.get(key + '_chunks');
  if (numChunksStr) {
    const numChunks = parseInt(numChunksStr, 10);
    let fullStr = '';
    for (let i = 0; i < numChunks; i++) {
      const chunk = cache.get(`${key}_${i}`);
      if (!chunk) break;
      fullStr += chunk;
    }
    if (fullStr.length > 0) {
      try { return JSON.parse(fullStr); } catch(e) {}
    }
  } else {
    // 2-2. 단일 데이터 확인
    const cachedStr = cache.get(key);
    if (cachedStr) {
      try { return JSON.parse(cachedStr); } catch(e) {}
    }
  }
  
  // 🌟 [핵심 수정] 개별 팀 투입현황은 시트에 없으므로 여기서 바로 종료 (시트 검색 지연 방지)
  if (key.startsWith('ORG_RES_')) {
    return null;
  }

  // 2-3. 공통 데이터의 경우 메모리에 없으면 시트에서 가져오기
  const sheet = SS.getSheetByName('_DB_CACHE_');
  if (!sheet) return null;
  const row = sheet.getDataRange().getValues().find(r => r[0] === key);
  if (row && row[1]) {
    try { return JSON.parse(row[1]); } catch(e) {}
  }
  return null;
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

  // 2. 물리 캐시에서 공통 데이터 가져오기 (수정됨)
  let cachedFullData = getGlobalCache('GLOBAL_CACHE_2026');

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

  // 캐시 시트에 저장 (수정됨)
  setGlobalCache('GLOBAL_CACHE_2026', finalCache);
  
  return finalCache;
}

/**
 * 최적화: 시트 데이터를 한 번만 읽어 전역 변수 및 2-Tier 캐시에 저장
 */
function getGlobalData() {
  // 1. 현재 실행 컨텍스트 내 메모리 캐시
  if (_cachedInputs && _cachedHR) return { inputs: _cachedInputs, hr: _cachedHR };
  
  const cacheKey = 'PARSED_RAW_DATA_2026';
  
  // 2. 글로벌 메모리/시트 캐시 조회
  const cached = getGlobalCache(cacheKey);
  if (cached) {
    _cachedInputs = cached.inputs;
    _cachedHR = cached.hr;
    return cached;
  }

  // 3. 캐시가 없을 때만 시트에서 읽기 (최초 1회 실행)
  const targetYear = 2026;
  const sheets = {
    ex: SS.getSheetByName('투입실적_실행예산'),
    mm: SS.getSheetByName('투입실적_MM등록'),
    fc: SS.getSheetByName('투입계획_FCST'),
    hr: SS.getSheetByName('HR_Report')
  };

  const rawEx = sheets.ex ? sheets.ex.getDataRange().getValues().slice(1) : [];
  const rawMm = sheets.mm ? sheets.mm.getDataRange().getValues().slice(1) : [];
  const rawFc = sheets.fc ? sheets.fc.getDataRange().getValues().slice(1) : [];
  const rawHr = sheets.hr ? sheets.hr.getDataRange().getValues().slice(1) : [];

  const map = { exec: {}, mm: {}, plan: {} };
  const process = (data, targetMap, mmCol, idCol, projCol, wbsCol) => {
    data.forEach(r => {
      if (Number(r[0]) !== targetYear) return;
      const key = `${r[0]}-${r[1]}-${String(r[idCol]).trim()}`;
      const val = parseFloat(r[mmCol] || 0);
      if (val === 0) return;
      
      if (!targetMap[key]) targetMap[key] = { total: 0, details: [] };
      targetMap[key].total += val;
      targetMap[key].details.push({ proj: r[projCol], wbs: r[wbsCol] || '-', mm: val });
    });
  };

  process(rawEx, map.exec, 2, 3, 6, 5);
  process(rawMm, map.mm, 2, 3, 6, 5);
  process(rawFc, map.plan, 5, 2, 4, 8);
  
  const hrMap = {};
  rawHr.forEach(r => {
    if (Number(r[0]) !== targetYear || EXCLUDED_DEPTS.includes(r[8])) return;
    const key = `${r[0]}-${r[1]}`;
    if (!hrMap[key]) hrMap[key] = [];
    hrMap[key].push(r);
  });

  _cachedInputs = map;
  _cachedHR = hrMap;
  
  // 파싱된 전체 데이터를 캐시에 저장
  setGlobalCache(cacheKey, { inputs: _cachedInputs, hr: _cachedHR });

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
  // 🌟 변경
  const { inputs } = getGlobalData(); 
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
 * 조직 상세 뷰 데이터 조회 (계산된 결과 통째로 캐싱)
 */
function getOrgResources(path, year, month, forceRefresh = false) {
  const pathLevel = path ? (Object.keys(path)[0] || '전사') : '전사';
  const pathName = path ? (Object.values(path)[0] || '전사') : '전사';
  const cacheKey = `ORG_RES_${year}_${month}_${pathLevel}_${pathName}`;

  // 🌟 forceRefresh가 아닐 때만 캐시를 타도록 수정
  if (!forceRefresh) {
    const cachedResult = getGlobalCache(cacheKey);
    if (cachedResult) {
      return cachedResult;
    }
  }

  // === 여기서부터 데이터 연산 시작 ===
  const timeline = generateTimeline();
  const { inputs, hr } = getGlobalData();

  const hrMap = {}; 
  const targetEmpIds = new Set();
  const targetYear = 2026;

  Object.values(hr).flat().forEach(r => {
    const y = Number(r[0]), m = Number(r[1]);
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
    let isCurrentHere = pathMatch(currentHR, path); 

    let latestHR = currentHR;
    if (!latestHR) {
      for (let i = timeline.length - 1; i >= 0; i--) {
        let temp = hrMap[`${timeline[i].year}-${timeline[i].month}-${id}`];
        if (pathMatch(temp, path)) { latestHR = temp; break; } 
      }
    }
    if (!latestHR) return;

    let m = { ...latestHR };
    m.statusLabel = isCurrentHere ? (m.date ? `발령일자: ${m.date}` : '') : `<span class="text-danger fw-bold">전출</span>`;

    let confProjs = new Set(), confWbs = new Set();
    let planProjs = new Set(), planWbs = new Set();

    m.timelineMM = timeline.map(t => {
      let tHR = hrMap[`${t.year}-${t.month}-${id}`];
      if (!pathMatch(tHR, path)) return { year: t.year, month: t.month, offset: t.offset, confirmed: null, planned: null, total: null };
      
      let k = `${t.year}-${t.month}-${id}`;
      let eMM = 0, mMM = 0, pMM = 0;
      
      if (inputs.exec[k]) {
        eMM = inputs.exec[k].total;
        inputs.exec[k].details.forEach(d => { 
          if(d.proj && d.proj !== '-') confProjs.add(d.proj); 
          if(d.wbs && d.wbs !== '-') confWbs.add(d.wbs);
        });
      }
      if (inputs.mm[k]) {
        mMM = inputs.mm[k].total;
        inputs.mm[k].details.forEach(d => { 
          if(d.proj && d.proj !== '-') confProjs.add(d.proj); 
          if(d.wbs && d.wbs !== '-') confWbs.add(d.wbs);
        });
      }
      if (inputs.plan[k]) {
        pMM = inputs.plan[k].total;
        inputs.plan[k].details.forEach(d => { 
          if(d.proj && d.proj !== '-') planProjs.add(d.proj); 
          if(d.wbs && d.wbs !== '-') planWbs.add(d.wbs);
        });
      }
      return { 
        year: t.year, month: t.month, offset: t.offset, 
        confirmed: eMM + mMM, planned: pMM, total: (eMM + mMM + pMM) 
      };
    });

    function getSummary(setObj) {
      if (setObj.size === 0) return { full: '-', short: '-' };
      const arr = Array.from(setObj);
      return { full: arr.join(', '), short: arr.length > 1 ? `${arr[0]} 외 ${arr.length - 1}건` : arr[0] };
    }
    
    const confP = getSummary(confProjs); m.confProjFull = confP.full; m.confProjShort = confP.short;
    const confW = getSummary(confWbs); m.confWbsFull = confW.full; m.confWbsShort = confW.short;
    const planP = getSummary(planProjs); m.planProjFull = planP.full; m.planProjShort = planP.short;
    const planW = getSummary(planWbs); m.planWbsFull = planW.full; m.planWbsShort = planW.short;

    if (!isCurrentHere) { groups['전출/퇴직'].push(m); } 
    else if (m.type === '휴직') { groups['휴직'].push(m); } 
    else if (m.isTarget === 0) { groups['SGA'].push(m); } 
    else if (m.type === 'Shared') { groups['Shared'].push(m); } 
    else {
        let k = `${Number(year)}-${Number(month)}-${id}`;
        let eMM = (inputs.exec[k] ? inputs.exec[k].total : 0) + (inputs.mm[k] ? inputs.mm[k].total : 0);
        let pMM = inputs.plan[k] ? inputs.plan[k].total : 0;
        
        if (eMM >= 1.0) groups['COS - 투입 확정'].push(m);
        else if (eMM > 0) groups['COS - 부분 투입'].push(m);
        else if (eMM === 0 && pMM > 0) groups['COS - 투입 예정'].push(m);
        else groups['COS - 미투입'].push(m);
    }
  });

  for(let g in groups) { if(groups[g].length === 0) delete groups[g]; }

  const finalResult = { timeline, grouped: groups };

  // 🌟 3. 연산이 끝난 최종 결과물을 캐시에 저장 (이후 동일 팀 조회 시 0초 컷)
  setGlobalCache(cacheKey, finalResult);

  return finalResult;
}

function saveBulkPlan(payload) {
  const { inputs } = getGlobalData();
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
    
    // 🌟 추가: 데이터가 변경되었으므로 글로벌 캐시 즉각 삭제 (무효화)
    try {
      CacheService.getScriptCache().remove('PARSED_RAW_DATA_2026');
      CacheService.getScriptCache().remove('GLOBAL_CACHE_2026');
    } catch(e) {}
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


/**
 * 🌟 트리거에 의해 백그라운드에서 주기적으로 실행될 캐시 강제 갱신 함수
 */
function forceRefreshCache() {
  console.log("백그라운드 캐시 갱신 시작...");
  
  // 1. 공통 캐시 강제 삭제 (무효화)
  try {
    CacheService.getScriptCache().remove('PARSED_RAW_DATA_2026');
    CacheService.getScriptCache().remove('GLOBAL_CACHE_2026');
  } catch(e) {}
  
  // 2. 전역 변수 초기화
  _cachedInputs = null;
  _cachedHR = null;
  
  // 3. 최신 데이터 파싱 및 대시보드 갱신
  getGlobalData();
  const fullData = rebuildAndReturnCache();
  
  // =========================================================
  // 🌟 [핵심 추가] 모든 본부/실/팀의 데이터를 미리 계산해서 캐시에 굽기
  // =========================================================
  const targetYear = 2026;
  const targetMonth = 5; // 당월 기준
  const tree = fullData.fullTree;
  
  // '전사' 전체 캐시 굽기 (forceRefresh = true)
  getOrgResources(null, targetYear, targetMonth, true);

  for (let dept in tree) {
    for (let div in tree[dept]) {
      if (div !== '-') {
        // 본부 단위 캐시 굽기
        getOrgResources({ div: div }, targetYear, targetMonth, true);
      }
      
      for (let unit in tree[dept][div]) {
        if (unit !== '-' && unit !== '(실 미소속)') {
          // 실 단위 캐시 굽기
          getOrgResources({ unit: unit }, targetYear, targetMonth, true);
        }
        
        tree[dept][div][unit].forEach(team => {
          if (team !== '-') {
            // 팀 단위 캐시 굽기
            getOrgResources({ team: team }, targetYear, targetMonth, true);
          }
        });
      }
    }
  }
  
  console.log("✅ 백그라운드 캐시 및 팀별 상세 현황 캐시 굽기 완료");
}
