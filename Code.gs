/**
 * 우리집 가계부 — Google Apps Script 버전
 * =====================================================================
 * 사용법:
 *   1) 구글 시트 새 문서 → 확장 프로그램 → Apps Script
 *   2) 이 코드 전체를 붙여넣고 저장(Ctrl+S)
 *   3) 함수 목록에서 buildBudget 선택 → ▶ 실행 (최초 1회 권한 승인)
 *   4) 시트로 돌아오면 6개 시트가 생성되어 있음
 *
 * 파일 업로드가 전혀 없으므로(코드 텍스트만) 업로드 차단 보안정책에 걸리지 않습니다.
 *
 * 엑셀(openpyxl)판과의 차이(구글 차트/시트 한계):
 *   - 카테고리 막대 차트: 구글 차트는 단일 계열의 막대별 색을 지원하지 않아 단색입니다.
 *   - 집행률 "데이터바"는 구글 시트에 없어 색농도(컬러스케일)로 대체했습니다.
 *   - 그 외 시트구조·수식·색·서식·조건부서식·드롭다운·진단·시트보호는 동일하게 재현됩니다.
 */

// ---------- 색상 (밝은 진노랑 포인트 + 파스텔) ----------
var F = '맑은 고딕';
var WHITE='#FFFFFF', INK='#3A3330', GRAY='#857F78', LINE='#EFE9DD';
var HEADL='#FFF6D8', ZEBRA='#FFFDF5', CARD='#FFFDF5', LIGHT='#FFF6D8';
var NAVY='#8BC34A', TEAL='#FFC72C', HEAD_TX='#33450F', SECTION='#33450F', HEAD_BG='#DCEDC8';
var INCOME='#7C9885', EXPENSE='#CC8B96', SAVE='#7E9BBE', AMBER='#E0A21E', VIOLET='#9787BE', SLATE='#9A9286';
var T_INCOME='#EEF4EE', T_EXPENSE='#FAF0F2', T_SAVE='#EFF3F8', T_AMBER='#FFF4DA', T_VIOLET='#F4F1F9', T_INK='#FBF6EC';
var INPUTF='#FFEFB0', INPUTC='#8A5A00';

// ---------- 숫자 포맷 ----------
var WON  = '"₩"#,##0;[Red]"-₩"#,##0;"-"';
var WON0 = '"₩"#,##0';
var PCT  = '0.0%';
var DATEF= 'yyyy-mm-dd';
var SOLID = SpreadsheetApp.BorderStyle.SOLID;

// ---------- 마스터 데이터 ----------
var expenseCats = [
  ['식비/장보기','변동',600000],['외식/배달/카페','변동',300000],['생활/소모품','변동',150000],
  ['주거/관리비','고정',550000],['통신/구독','고정',160000],['교통/차량','변동',250000],
  ['의료/건강','변동',120000],['교육/육아','고정',450000],['문화/여가','변동',200000],
  ['의류/미용','변동',150000],['경조사/선물','변동',100000],['보험','고정',300000],
  ['저축/투자','고정',800000],['기타','변동',100000]
];
var incomeCats=['급여','상여/성과급','사업/부업','금융수입(이자/배당)','기타수입'];
var members=['본인','배우자','공동','자녀'];
var methods=['현금','체크카드','신용카드','계좌이체','간편결제(페이)','기타'];
var types=['수입','지출','이체'];

// ---------- 거래내역 범위 상수 ----------
var NROW=1200, R1=2, R2=NROW+1;
var S_GUIDE='시작하기', S_SET='설정', S_LED='거래내역', S_DASH='대시보드', S_REP='기간리포트', S_INS='진단&인사이트';
function L(col){ return S_LED+'!$'+col+'$'+R1+':$'+col+'$'+R2; }
var G_=L('G'), I_=L('I'), J_=L('J'), B_=L('B'), C_=L('C'), E_=L('E'), K_=L('K'), Lq=L('L');
var YR=S_SET+'!$C$5', MO=S_SET+'!$C$6';

// 설정 시트 동적 행 계산
var EXP_HDR=19, EXP_FIRST=EXP_HDR+2, EXP_LAST=EXP_FIRST+expenseCats.length-1, TOT_R=EXP_LAST+1;
var CATTBL=S_SET+'!$B$'+EXP_FIRST+':$D$'+EXP_LAST;
var KIND_RNG=S_SET+'!$D$'+EXP_FIRST+':$D$'+EXP_LAST;

// ---------- 헬퍼 ----------
function cw(charW){ return Math.round(charW*7.5); }          // 엑셀 열너비 → px
function rh(pt){ return Math.round(pt*1.33); }                // pt → px

/** put: 단일 셀 값/서식 */
function put(sh, a1, value, o){
  o = o||{};
  var r = sh.getRange(a1);
  if(value!==null && value!==undefined && value!==''){
    if(typeof value==='string' && value.charAt(0)==='=') r.setFormula(value);
    else r.setValue(value);
  }
  r.setFontFamily(F);
  r.setFontSize(o.size||10);
  if(o.bold) r.setFontWeight('bold');
  r.setFontColor(o.color||INK);
  if(o.fill) r.setBackground(o.fill);
  r.setHorizontalAlignment(o.align||'left');
  r.setVerticalAlignment('middle');
  if(o.wrap) r.setWrap(true);
  if(o.fmt) r.setNumberFormat(o.fmt);
  if(o.border) r.setBorder(null,null,true,null,null,null, LINE, SOLID);
  return r;
}
function hdr(sh, a1, text){ return put(sh, a1, text, {size:10,bold:true,color:HEAD_TX,fill:HEAD_BG,align:'center',border:true}); }
function section(sh, a1, text){ return put(sh, a1, text, {size:12,bold:true,color:SECTION}); }

// =====================================================================
function buildBudget(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = [S_GUIDE, S_SET, S_LED, S_DASH, S_REP, S_INS];
  names.forEach(function(n){ var s=ss.getSheetByName(n); if(s) ss.deleteSheet(s); });
  var sh = {};
  names.forEach(function(n){ sh[n]=ss.insertSheet(n); n; });

  buildSettings_(sh[S_SET]);
  buildLedger_(sh[S_LED], ss);
  buildDashboard_(sh[S_DASH]);
  buildReport_(sh[S_REP]);
  buildInsights_(sh[S_INS]);
  buildGuide_(sh[S_GUIDE]);

  // 시트 순서 정리
  names.forEach(function(n,i){ ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i+1); });

  // 기본 시트 제거
  ['Sheet1','시트1'].forEach(function(n){ var s=ss.getSheetByName(n); if(s) ss.deleteSheet(s); });

  // 시트 보호 (입력칸만 편집 가능)
  protectSheet_(sh[S_LED], sh[S_LED].getRange('A2:H'+R2));
  [S_DASH,S_REP,S_INS,S_GUIDE].forEach(function(n){ protectSheet_(sh[n], null); });

  ss.setActiveSheet(sh[S_DASH]);
  SpreadsheetApp.getUi().alert('가계부 생성 완료! 6개 시트가 만들어졌습니다.');
}

function protectSheet_(sh, unprotectedRange){
  var p = sh.protect().setDescription('자동 생성 — 입력칸 외 보호');
  if(unprotectedRange) p.setUnprotectedRanges([unprotectedRange]);
  // 소유자 본인은 계속 편집 가능 (다른 공유자에게만 보호 적용)
}

// =====================================================================
// 1) 설정
function buildSettings_(sh){
  sh.setHiddenGridlines(true);
  sh.setTabColor('#7F8C8D');
  var widths={A:2.5,B:20,C:14,D:12,E:18,F:18,G:18,H:22,I:12};
  for(var c in widths) sh.setColumnWidth(colNum_(c), cw(widths[c]));

  sh.getRange('B1:I1').merge();
  put(sh,'B1','⚙️  설정  —  이 시트의 값만 바꾸면 전체 가계부에 반영됩니다',{size:19,bold:true,color:HEAD_TX,fill:NAVY,align:'left'});
  sh.setRowHeight(1, rh(40));
  put(sh,'B2','노란색 칸(파란 글씨)이 직접 입력하는 칸입니다.',{size:9,color:GRAY});

  section(sh,'B4','■ 기준 기간 (대시보드·리포트가 이 기간을 표시)');
  put(sh,'B5','기준 연도',{size:10,bold:true,fill:HEADL,border:true});
  put(sh,'C5',2026,{size:11,bold:true,color:INPUTC,fill:INPUTF,align:'center',border:true,fmt:'0'});
  put(sh,'B6','기준 월',{size:10,bold:true,fill:HEADL,border:true});
  put(sh,'C6',6,{size:11,bold:true,color:INPUTC,fill:INPUTF,align:'center',border:true,fmt:'0'});
  put(sh,'D5','← 보고 싶은 연/월로 바꾸세요',{size:9,color:GRAY});

  section(sh,'B8','■ 목록 (드롭다운 항목)');
  hdr(sh,'B9','구성원'); sh.getRange('D9:E9').merge(); hdr(sh,'D9','결제수단');
  hdr(sh,'F9','유형'); sh.getRange('H9:I9').merge(); hdr(sh,'H9','수입 카테고리');
  for(var i=0;i<8;i++){
    var rr=10+i, z=(i%2===0)?ZEBRA:WHITE;
    put(sh,'B'+rr, members[i]||'', {size:10,fill:z,border:true});
    sh.getRange('D'+rr+':E'+rr).merge();
    put(sh,'D'+rr, methods[i]||'', {size:10,fill:z,border:true});
    put(sh,'F'+rr, types[i]||'', {size:10,fill:z,border:true});
    sh.getRange('H'+rr+':I'+rr).merge();
    put(sh,'H'+rr, incomeCats[i]||'', {size:10,fill:z,border:true});
  }

  section(sh,'B'+EXP_HDR,'■ 지출 카테고리 & 월 예산');
  hdr(sh,'B'+(EXP_HDR+1),'대분류'); hdr(sh,'C'+(EXP_HDR+1),'월 예산');
  hdr(sh,'D'+(EXP_HDR+1),'고정/변동'); hdr(sh,'E'+(EXP_HDR+1),'메모');
  for(var k=0;k<expenseCats.length;k++){
    var r=EXP_FIRST+k, z2=(k%2===0)?ZEBRA:WHITE, ec=expenseCats[k];
    put(sh,'B'+r, ec[0], {size:10,fill:z2,border:true});
    put(sh,'C'+r, ec[2], {size:10,bold:true,color:INPUTC,fill:INPUTF,align:'right',border:true,fmt:WON0});
    put(sh,'D'+r, ec[1], {size:10,fill:z2,align:'center',border:true});
    put(sh,'E'+r, '', {size:10,fill:z2,border:true});
  }
  put(sh,'B'+TOT_R,'월 예산 합계',{size:10,bold:true,color:HEAD_TX,fill:NAVY,border:true});
  put(sh,'C'+TOT_R,'=SUM(C'+EXP_FIRST+':C'+EXP_LAST+')',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'right',border:true,fmt:WON0});
  put(sh,'D'+TOT_R,'',{fill:NAVY,border:true}); put(sh,'E'+TOT_R,'',{fill:NAVY,border:true});

  // 대분류(통합) 드롭다운 소스 (K열)
  put(sh,'K9','대분류(통합)',{size:10,bold:true,color:HEAD_TX,fill:'#95A5A6',align:'center',border:true});
  var allcats = expenseCats.map(function(c){return c[0];}).concat(incomeCats).concat(['이체/계좌이동']);
  for(var a=0;a<allcats.length;a++){
    put(sh,'K'+(10+a), allcats[a], {size:9,fill:(a%2===0)?ZEBRA:WHITE,border:true});
  }
}

// =====================================================================
// 2) 거래내역 (배치 처리)
function buildLedger_(sh, ss){
  sh.setHiddenGridlines(true);
  sh.setTabColor('#27AE60');
  sh.setFrozenRows(1);
  var headers=['날짜','유형','대분류','소분류(상세)','구성원','결제수단','금액(원)','메모','연','월','주차','분기','연-월'];
  var widths=[13,8,14,16,9,14,14,22,6,6,6,6,9];
  for(var i=0;i<headers.length;i++){
    sh.setColumnWidth(i+1, cw(widths[i]));
    hdr(sh, colLetter_(i+1)+'1', headers[i]);
  }
  // 헬퍼 헤더 회색
  ['I','J','K','L','M'].forEach(function(c){ sh.getRange(c+'1').setBackground(GRAY).setFontColor(WHITE); });
  sh.setRowHeight(1, rh(24));

  // 샘플 데이터
  function D(day,mon,yr){ return new Date(yr||2026,(mon||6)-1,day); }
  var sample=[
    [D(25),'수입','급여','6월 급여','본인','계좌이체',2800000,''],
    [D(25),'수입','급여','6월 급여','배우자','계좌이체',2000000,''],
    [D(5),'지출','주거/관리비','관리비+월세','공동','신용카드',550000,''],
    [D(10),'지출','통신/구독','휴대폰+넷플릭스','공동','신용카드',92000,''],
    [D(15),'지출','보험','실손+종신','공동','계좌이체',300000,''],
    [D(5),'지출','교육/육아','학원비','자녀','계좌이체',450000,''],
    [D(25),'지출','저축/투자','적금 자동이체','공동','계좌이체',500000,'비상금'],
    [D(2),'지출','식비/장보기','마트 장보기','본인','체크카드',150000,''],
    [D(8),'지출','식비/장보기','마트+정육점','배우자','체크카드',160000,''],
    [D(15),'지출','식비/장보기','장보기','본인','체크카드',140000,''],
    [D(22),'지출','식비/장보기','마트','배우자','체크카드',130000,''],
    [D(3),'지출','외식/배달/카페','치킨 배달','공동','간편결제(페이)',40000,''],
    [D(7),'지출','외식/배달/카페','주말 외식','공동','신용카드',60000,''],
    [D(13),'지출','외식/배달/카페','카페','본인','간편결제(페이)',30000,''],
    [D(18),'지출','외식/배달/카페','가족 회식','공동','신용카드',130000,'예산 초과 주의'],
    [D(20),'지출','외식/배달/카페','가족 외식','공동','신용카드',110000,''],
    [D(27),'지출','외식/배달/카페','배달','배우자','간편결제(페이)',80000,''],
    [D(1),'지출','교통/차량','주유+하이패스','본인','신용카드',130000,''],
    [D(14),'지출','교통/차량','대중교통 충전','배우자','간편결제(페이)',100000,''],
    [D(11),'지출','의료/건강','소아과+약국','자녀','체크카드',90000,''],
    [D(21),'지출','문화/여가','영화+나들이','공동','신용카드',180000,''],
    [D(9),'지출','의류/미용','여름옷','배우자','신용카드',140000,''],
    [D(24),'지출','의류/미용','미용실','본인','체크카드',70000,'예산 초과 주의'],
    [D(6),'지출','생활/소모품','세제+휴지','본인','체크카드',70000,''],
    [D(18),'지출','생활/소모품','주방용품','배우자','체크카드',60000,''],
    [D(16),'지출','경조사/선물','결혼 축의금','본인','계좌이체',100000,''],
    [D(19),'지출','기타','기타 잡비','공동','현금',80000,''],
    [D(25,5),'수입','급여','5월 급여','본인','계좌이체',2800000,''],
    [D(25,5),'수입','급여','5월 급여','배우자','계좌이체',2000000,''],
    [D(5,5),'지출','주거/관리비','관리비+월세','공동','신용카드',550000,''],
    [D(10,5),'지출','통신/구독','통신+구독','공동','신용카드',92000,''],
    [D(15,5),'지출','보험','보험료','공동','계좌이체',300000,''],
    [D(5,5),'지출','교육/육아','학원비','자녀','계좌이체',450000,''],
    [D(25,5),'지출','저축/투자','적금','공동','계좌이체',500000,''],
    [D(12,5),'지출','식비/장보기','장보기 합계','본인','체크카드',520000,''],
    [D(14,5),'지출','외식/배달/카페','외식+배달','공동','신용카드',300000,''],
    [D(8,5),'지출','교통/차량','주유+교통','본인','신용카드',200000,''],
    [D(20,5),'지출','문화/여가','여가','공동','신용카드',150000,''],
    [D(9,5),'지출','의류/미용','의류','배우자','신용카드',130000,''],
    [D(17,5),'지출','의료/건강','병원','자녀','체크카드',60000,''],
    [D(11,5),'지출','생활/소모품','생필품','본인','체크카드',110000,''],
    [D(22,5),'지출','경조사/선물','선물','본인','신용카드',50000,''],
    [D(19,5),'지출','기타','기타 잡비','공동','현금',60000,'']
  ];
  sh.getRange(2,1,sample.length,8).setValues(sample);

  // 헬퍼 수식 (I~M, 전 행 배치)
  var fml=[];
  for(var rr=R1; rr<=R2; rr++){
    var a='$A'+rr;
    fml.push([
      '=IF('+a+'="","",YEAR('+a+'))',
      '=IF('+a+'="","",MONTH('+a+'))',
      '=IF('+a+'="","",WEEKNUM('+a+',2))',
      '=IF('+a+'="","",ROUNDUP(MONTH('+a+')/3,0))',
      '=IF('+a+'="","",TEXT('+a+',"YYYY-MM"))'
    ]);
  }
  sh.getRange(R1,9,fml.length,5).setFormulas(fml);

  // 서식 (범위 배치)
  var all=sh.getRange('A'+R1+':M'+R2); all.setFontFamily(F);
  sh.getRange('A'+R1+':H'+R2).setFontSize(10).setFontColor(INK)
    .setBorder(null,null,null,null,null,true, LINE, SOLID);
  sh.getRange('I'+R1+':M'+R2).setFontSize(9).setFontColor(GRAY).setHorizontalAlignment('center');
  // 정렬
  sh.getRange('A'+R1+':A'+R2).setHorizontalAlignment('center').setNumberFormat(DATEF);
  sh.getRange('B'+R1+':B'+R2).setHorizontalAlignment('center');
  sh.getRange('C'+R1+':D'+R2).setHorizontalAlignment('left');
  sh.getRange('E'+R1+':F'+R2).setHorizontalAlignment('center');
  sh.getRange('G'+R1+':G'+R2).setHorizontalAlignment('right').setNumberFormat(WON);
  sh.getRange('H'+R1+':H'+R2).setHorizontalAlignment('left');
  sh.getRange('I'+R1+':L'+R2).setNumberFormat('0');

  // 드롭다운 (데이터 검증)
  var setSh = ss.getSheetByName(S_SET);
  var allcats = expenseCats.map(function(c){return c[0];}).concat(incomeCats).concat(['이체/계좌이동']);
  var dvType = SpreadsheetApp.newDataValidation().requireValueInList(types,true).setAllowInvalid(true).build();
  var dvMem  = SpreadsheetApp.newDataValidation().requireValueInRange(setSh.getRange('B10:B17'),true).setAllowInvalid(true).build();
  var dvPay  = SpreadsheetApp.newDataValidation().requireValueInRange(setSh.getRange('D10:D17'),true).setAllowInvalid(true).build();
  var dvCat  = SpreadsheetApp.newDataValidation().requireValueInRange(setSh.getRange('K10:K'+(10+allcats.length-1)),true).setAllowInvalid(true).build();
  sh.getRange('B'+R1+':B'+R2).setDataValidation(dvType);
  sh.getRange('C'+R1+':C'+R2).setDataValidation(dvCat);
  sh.getRange('E'+R1+':E'+R2).setDataValidation(dvMem);
  sh.getRange('F'+R1+':F'+R2).setDataValidation(dvPay);

  // 행 색 구분(조건부서식)
  var rng=[sh.getRange('A'+R1+':H'+R2)];
  var rules=sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="수입"').setBackground('#E8F5E9').setRanges(rng).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B2="이체"').setBackground('#EFEFEF').setRanges(rng).build());
  sh.setConditionalFormatRules(rules);
}

// =====================================================================
// 3) 대시보드
function buildDashboard_(sh){
  sh.setHiddenGridlines(true);
  sh.setTabColor(NAVY);
  var widths={A:2.5,B:17,C:14,D:14,E:12,F:2.5,G:15,H:13,I:13,J:11,K:2.5,L:13,M:13};
  for(var c in widths) sh.setColumnWidth(colNum_(c), cw(widths[c]));

  sh.getRange('B1:J1').merge();
  put(sh,'B1','📊  월간 대시보드',{size:23,bold:true,color:HEAD_TX,fill:NAVY});
  sh.setRowHeight(1, rh(44));
  sh.getRange('B2:J2').merge();
  put(sh,'B2','=TEXT('+YR+',"0")&"년 "&TEXT('+MO+',"0")&"월 가계 요약   ·   기준 기간은 [설정] 시트에서 변경"',{size:11,bold:true,color:GRAY});

  // 기준 셀 + 집계 (L열)
  put(sh,'L1','='+YR,{size:10,bold:true,align:'center',fmt:'0'});
  put(sh,'L2','='+MO,{size:10,bold:true,align:'center',fmt:'0'});
  var yC='$L$1', mC='$L$2';
  function sif(typ,cat){
    var p=[G_, I_, yC, J_, mC];
    if(typ) p.push(B_, '"'+typ+'"');
    if(cat) p.push(C_, '"'+cat+'"');
    return '=SUMIFS('+p.join(',')+')';
  }
  put(sh,'L4',sif('수입'),{size:9,color:GRAY,fmt:WON});
  put(sh,'L5',sif('지출'),{size:9,color:GRAY,fmt:WON});
  put(sh,'L6',sif('지출','저축/투자'),{size:9,color:GRAY,fmt:WON});
  put(sh,'L7','=L4-L5',{size:9,color:GRAY,fmt:WON});
  put(sh,'L8','=L5-L6',{size:9,color:GRAY,fmt:WON});
  put(sh,'L9','=IFERROR((L4-L8)/L4,0)',{size:9,color:GRAY,fmt:PCT});
  put(sh,'L10','='+S_SET+'!$C$'+TOT_R,{size:9,color:GRAY,fmt:WON});

  // KPI 카드
  var cards=[
    ['B','총 수입','=L4',INCOME,T_INCOME,WON,4],
    ['D','총 지출','=L5',EXPENSE,T_EXPENSE,WON,4],
    ['G','소비지출','=L8',AMBER,T_AMBER,WON,4],
    ['B','저축·투자','=L6',SAVE,T_SAVE,WON,7],
    ['D','순잔액 (수입−지출)','=L7',SLATE,T_INK,WON,7],
    ['G','저축률','=L9',VIOLET,T_VIOLET,PCT,7]
  ];
  cards.forEach(function(cd){
    var c1=cd[0], ci=colNum_(c1), c2=colLetter_(ci+1), top=cd[6];
    sh.getRange(c1+top+':'+c2+top).merge();
    sh.getRange(c1+(top+1)+':'+c2+(top+1)).merge();
    put(sh,c1+top, cd[1], {size:10,bold:true,color:GRAY,fill:cd[4],align:'left'});
    sh.getRange(c1+top).setVerticalAlignment('middle');
    put(sh,c1+(top+1), cd[2], {size:21,bold:true,color:cd[3],fill:cd[4],align:'left',fmt:cd[5]});
    sh.getRange(c1+top+':'+c2+top).setBorder(true,null,null,null,null,null, cd[3], SpreadsheetApp.BorderStyle.SOLID_THICK);
    sh.getRange(c1+(top+1)+':'+c2+(top+1)).setBackground(cd[4]);
    sh.setRowHeight(top, rh(24)); sh.setRowHeight(top+1, rh(44));
  });

  // 전월 대비
  put(sh,'M7','=SUMIFS('+G_+','+I_+','+yC+','+J_+','+mC+'-1,'+B_+',"지출")',{size:9,color:GRAY,fmt:WON});
  sh.getRange('B10:J10').merge();
  put(sh,'B10','="전월 대비 지출: "&TEXT(L5-M7,"+₩#,##0;-₩#,##0")&"  ("&TEXT(IFERROR((L5-M7)/M7,0),"+0.0%;-0.0%")&")"',{size:10,bold:true,fill:LIGHT,border:true});

  // 카테고리별 표
  var TBL=13;
  section(sh,'B'+(TBL-1),'■ 카테고리별 지출 · 예산 대비');
  ['대분류','지출액','예산','집행률'].forEach(function(h,i){ hdr(sh, colLetter_(2+i)+TBL, h); });
  var catFirst=TBL+1;
  for(var i=0;i<expenseCats.length;i++){
    var rr=catFirst+i, z=(i%2===0)?ZEBRA:WHITE;
    put(sh,'B'+rr, expenseCats[i][0], {size:10,fill:z,border:true});
    put(sh,'C'+rr, '=SUMIFS('+G_+','+I_+','+yC+','+J_+','+mC+','+B_+',"지출",'+C_+',B'+rr+')',{size:10,fill:z,align:'right',border:true,fmt:WON});
    put(sh,'D'+rr, '=IFERROR(VLOOKUP(B'+rr+','+CATTBL+',2,0),0)',{size:10,fill:z,align:'right',border:true,fmt:WON});
    put(sh,'E'+rr, '=IFERROR(C'+rr+'/D'+rr+',0)',{size:10,fill:z,align:'center',border:true,fmt:PCT});
  }
  var catLast=catFirst+expenseCats.length-1, sr=catLast+1;
  put(sh,'B'+sr,'합계',{size:10,bold:true,color:HEAD_TX,fill:NAVY,border:true});
  put(sh,'C'+sr,'=SUM(C'+catFirst+':C'+catLast+')',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'right',border:true,fmt:WON});
  put(sh,'D'+sr,'=SUM(D'+catFirst+':D'+catLast+')',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'right',border:true,fmt:WON});
  put(sh,'E'+sr,'=IFERROR(C'+sr+'/D'+sr+',0)',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'center',border:true,fmt:PCT});

  // 집행률: 데이터바 대체(컬러스케일) + 초과 빨강글씨
  var rules=sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#DCEDC8', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setGradientMaxpointWithValue('#F4A7A7', SpreadsheetApp.InterpolationType.NUMBER, '1.5')
    .setRanges([sh.getRange('E'+catFirst+':E'+catLast)]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(1).setFontColor('#C00000').setBold(true)
    .setRanges([sh.getRange('E'+catFirst+':E'+catLast)]).build());

  // 구성원별 표
  var MB=TBL;
  section(sh,'G'+(MB-1),'■ 구성원별 지출');
  hdr(sh,'G'+MB,'구성원'); hdr(sh,'H'+MB,'지출액');
  for(var m=0;m<members.length;m++){
    var r2=MB+1+m, z2=(m%2===0)?ZEBRA:WHITE;
    put(sh,'G'+r2, members[m], {size:10,fill:z2,border:true});
    put(sh,'H'+r2, '=SUMIFS('+G_+','+I_+','+yC+','+J_+','+mC+','+B_+',"지출",'+E_+',G'+r2+')',{size:10,fill:z2,align:'right',border:true,fmt:WON});
  }
  var memFirst=MB+1, memLast=MB+members.length;

  // 주차별 표
  var WK=MB+members.length+3;
  section(sh,'G'+(WK-1),'■ 주차별 지출 추이(최근 8주)');
  hdr(sh,'G'+WK,'주차'); hdr(sh,'H'+WK,'지출액'); put(sh,'I'+WK,'주번호',{size:9,color:GRAY,align:'center'});
  var basewk='WEEKNUM(DATE('+yC+','+mC+',28),2)';
  for(var w=0;w<8;w++){
    var r3=WK+1+w, off=w-7, z3=(w%2===0)?ZEBRA:WHITE;
    put(sh,'I'+r3, '=MAX(1,'+basewk+'+('+off+'))',{size:9,color:GRAY,align:'center'});
    put(sh,'G'+r3, '="W"&TEXT(I'+r3+',"00")',{size:10,fill:z3,align:'center',border:true});
    put(sh,'H'+r3, '=SUMIFS('+G_+','+I_+','+yC+','+K_+',I'+r3+','+B_+',"지출")',{size:10,fill:z3,align:'right',border:true,fmt:WON});
  }
  var wkFirst=WK+1, wkLast=WK+8;

  sh.setConditionalFormatRules(rules);

  // 차트
  var barColors=['#F2C14E','#7C9885','#CC8B96','#7E9BBE','#9787BE','#E0A21E','#8FBF9F','#E8A87C','#B5A8D1','#8CB8C4','#C9A66B','#D49AA6','#9DB17C','#B0A18F'];
  var ch1=sh.newChart().asBarChart()
    .addRange(sh.getRange('B'+catFirst+':C'+catLast))
    .setOption('title','카테고리별 지출 (이번 달)')
    .setOption('legend',{position:'none'})
    .setOption('colors',[TEAL])
    .setOption('width',430).setOption('height',430)
    .setPosition(sr+2, 2, 0, 0).build();
  sh.insertChart(ch1);

  var ch2=sh.newChart().asPieChart()
    .addRange(sh.getRange('G'+memFirst+':H'+memLast))
    .setOption('title','구성원별 지출 비중')
    .setOption('colors',['#FFC72C','#7E9BBE','#7C9885','#CC8B96'])
    .setOption('pieSliceText','percentage')
    .setOption('width',300).setOption('height',280)
    .setPosition(sr+2, 7, 0, 0).build();
  sh.insertChart(ch2);

  var ch3=sh.newChart().asLineChart()
    .addRange(sh.getRange('G'+wkFirst+':H'+wkLast))
    .setOption('title','주차별 지출 추이')
    .setOption('legend',{position:'none'})
    .setOption('colors',[TEAL])
    .setOption('width',380).setOption('height',280)
    .setPosition(sr+18, 7, 0, 0).build();
  sh.insertChart(ch3);

  sh.hideColumns(colNum_('L'),2);  // L,M 숨김
}

// =====================================================================
// 4) 기간리포트
function buildReport_(sh){
  sh.setHiddenGridlines(true);
  sh.setTabColor(TEAL);
  var widths={A:2.5,B:14,C:14,D:14,E:14,F:11,G:2.5,H:12,I:14,J:14,K:11};
  for(var c in widths) sh.setColumnWidth(colNum_(c), cw(widths[c]));
  sh.getRange('B1:K1').merge();
  put(sh,'B1','📈  기간별 리포트  (주간 · 월간 · 분기 · 연간)',{size:20,bold:true,color:HEAD_TX,fill:NAVY});
  sh.setRowHeight(1, rh(42));
  put(sh,'L1','='+YR,{size:10,fmt:'0'}); put(sh,'L2','='+MO,{size:10,fmt:'0'});
  var yR='$L$1', mR='$L$2';

  // 월간
  section(sh,'B2','■ 월간 추이 (기준 연도)');
  ['월','수입','소비지출','순잔액','저축률'].forEach(function(h,i){ hdr(sh, colLetter_(2+i)+'3', h); });
  var mf=4;
  for(var i=0;i<12;i++){
    var rr=mf+i, mn=i+1, z=(i%2===0)?ZEBRA:WHITE;
    put(sh,'B'+rr, mn+'월', {size:10,fill:z,align:'center',border:true});
    put(sh,'C'+rr, '=SUMIFS('+G_+','+I_+','+yR+','+J_+','+mn+','+B_+',"수입")',{size:10,fill:z,align:'right',border:true,fmt:WON});
    put(sh,'D'+rr, '=SUMIFS('+G_+','+I_+','+yR+','+J_+','+mn+','+B_+',"지출")-SUMIFS('+G_+','+I_+','+yR+','+J_+','+mn+','+B_+',"지출",'+C_+',"저축/투자")',{size:10,fill:z,align:'right',border:true,fmt:WON});
    put(sh,'E'+rr, '=C'+rr+'-SUMIFS('+G_+','+I_+','+yR+','+J_+','+mn+','+B_+',"지출")',{size:10,fill:z,align:'right',border:true,fmt:WON});
    put(sh,'F'+rr, '=IFERROR((C'+rr+'-D'+rr+')/C'+rr+',0)',{size:10,fill:z,align:'center',border:true,fmt:PCT});
  }
  var ml=mf+11;
  put(sh,'B'+(ml+1),'연간 합계',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'center',border:true});
  put(sh,'C'+(ml+1),'=SUM(C'+mf+':C'+ml+')',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'right',border:true,fmt:WON});
  put(sh,'D'+(ml+1),'=SUM(D'+mf+':D'+ml+')',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'right',border:true,fmt:WON});
  put(sh,'E'+(ml+1),'=SUM(E'+mf+':E'+ml+')',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'right',border:true,fmt:WON});
  put(sh,'F'+(ml+1),'=IFERROR((C'+(ml+1)+'-D'+(ml+1)+')/C'+(ml+1)+',0)',{size:10,bold:true,color:HEAD_TX,fill:NAVY,align:'center',border:true,fmt:PCT});

  // 분기
  section(sh,'H2','■ 분기 요약');
  ['분기','수입','소비지출','저축률'].forEach(function(h,i){ hdr(sh, colLetter_(8+i)+'3', h); });
  var qf=4;
  for(var q=0;q<4;q++){
    var r=qf+q, qn=q+1, z2=(q%2===0)?ZEBRA:WHITE;
    put(sh,'H'+r, 'Q'+qn, {size:10,fill:z2,align:'center',border:true});
    put(sh,'I'+r, '=SUMIFS('+G_+','+I_+','+yR+','+Lq+','+qn+','+B_+',"수입")',{size:10,fill:z2,align:'right',border:true,fmt:WON});
    put(sh,'J'+r, '=SUMIFS('+G_+','+I_+','+yR+','+Lq+','+qn+','+B_+',"지출")-SUMIFS('+G_+','+I_+','+yR+','+Lq+','+qn+','+B_+',"지출",'+C_+',"저축/투자")',{size:10,fill:z2,align:'right',border:true,fmt:WON});
    put(sh,'K'+r, '=IFERROR((I'+r+'-J'+r+')/I'+r+',0)',{size:10,fill:z2,align:'center',border:true,fmt:PCT});
  }
  var ql=qf+3;

  // 주간 (최근 12주)
  var WKR=ql+3;
  section(sh,'H'+(WKR-1),'■ 주간 요약 (기준일 기준 최근 12주)');
  ['주차','수입','지출','순잔액'].forEach(function(h,i){ hdr(sh, colLetter_(8+i)+WKR, h); });
  put(sh,'L'+WKR,'주번호',{size:9,color:GRAY,align:'center'});
  var wf=WKR+1, basewkR='WEEKNUM(DATE('+yR+','+mR+',28),2)';
  for(var w=0;w<12;w++){
    var r=wf+w, off=w-11, z3=(w%2===0)?ZEBRA:WHITE;
    put(sh,'L'+r, '=MAX(1,'+basewkR+'+('+off+'))',{size:9,color:GRAY,align:'center'});
    put(sh,'H'+r, '="W"&TEXT(L'+r+',"00")',{size:10,fill:z3,align:'center',border:true});
    put(sh,'I'+r, '=SUMIFS('+G_+','+I_+','+yR+','+K_+',L'+r+','+B_+',"수입")',{size:10,fill:z3,align:'right',border:true,fmt:WON});
    put(sh,'J'+r, '=SUMIFS('+G_+','+I_+','+yR+','+K_+',L'+r+','+B_+',"지출")',{size:10,fill:z3,align:'right',border:true,fmt:WON});
    put(sh,'K'+r, '=I'+r+'-J'+r,{size:10,fill:z3,align:'right',border:true,fmt:WON});
  }
  var wl=wf+11;

  // 연간 요약 카드
  var AN=ml+4;
  section(sh,'B'+(AN-1),'■ 연간 요약');
  var ann=[['연 수입','=C'+(ml+1),INCOME,WON],['연 지출(전체)','=SUMIFS('+G_+','+I_+','+yR+','+B_+',"지출")',EXPENSE,WON],
           ['연 순잔액','=E'+(ml+1),SLATE,WON],['연 저축률','=F'+(ml+1),VIOLET,PCT]];
  ann.forEach(function(a,i){
    var cc=colLetter_(2+i);
    put(sh,cc+AN, a[0], {size:9,bold:true,color:HEAD_TX,fill:a[2],align:'center',border:true});
    put(sh,cc+(AN+1), a[1], {size:12,bold:true,color:a[2],fill:CARD,align:'center',border:true,fmt:a[3]});
  });
  sh.setRowHeight(AN+1, rh(26));

  // 차트
  var chm=sh.newChart().asColumnChart()
    .addRange(sh.getRange('B3:D'+ml))
    .setOption('title','월별 수입 vs 소비지출')
    .setOption('colors',[INCOME,EXPENSE])
    .setOption('width',480).setOption('height',280)
    .setPosition(AN+3, 2, 0, 0).build();
  sh.insertChart(chm);
  var chq=sh.newChart().asColumnChart()
    .addRange(sh.getRange('H3:J'+ql))
    .setOption('title','분기별 수입 vs 소비지출')
    .setOption('colors',[INCOME,EXPENSE])
    .setOption('width',360).setOption('height',280)
    .setPosition(wl+2, 8, 0, 0).build();
  sh.insertChart(chq);
  sh.hideColumns(colNum_('L'));
}

// =====================================================================
// 5) 진단 & 인사이트
function buildInsights_(sh){
  sh.setHiddenGridlines(true);
  sh.setTabColor('#D97706');
  var widths={A:2.5,B:20,C:10,D:46,E:46};
  for(var c in widths) sh.setColumnWidth(colNum_(c), cw(widths[c]));
  sh.getRange('B1:E1').merge();
  put(sh,'B1','🔍  진단 & 인사이트  —  이번 달 살림 자동 분석',{size:20,bold:true,color:HEAD_TX,fill:NAVY});
  sh.setRowHeight(1, rh(42));
  sh.getRange('B2:E2').merge();
  put(sh,'B2','=TEXT('+YR+',"0")&"년 "&TEXT('+MO+',"0")&"월 기준 · [설정]에서 기간 변경 시 자동 갱신"',{size:10,bold:true,color:'#B9770D'});

  var D=S_DASH+'!';
  var catFirst=14, catLast=27, srD=28;   // 대시보드 카테고리표 위치
  section(sh,'B4','■ 핵심 지표');
  var metrics=[['총수입','='+D+'$L$4',WON],['소비지출(저축 제외)','='+D+'$L$8',WON],['저축·투자','='+D+'$L$6',WON],
               ['저축률','='+D+'$L$9',PCT],['예산 집행률(총지출÷예산)','='+S_DASH+'!$E$'+srD,PCT]];
  hdr(sh,'B5','지표'); hdr(sh,'C5','값');
  metrics.forEach(function(m,i){
    var rr=6+i, z=(i%2===0)?ZEBRA:WHITE;
    put(sh,'B'+rr, m[0], {size:10,fill:z,border:true});
    put(sh,'C'+rr, m[1], {size:10,bold:true,fill:z,align:'center',border:true,fmt:m[2]});
  });

  // 보조계산
  put(sh,'G6','=SUMIFS('+G_+','+I_+','+D+'$L$1,'+J_+','+D+'$L$2,'+B_+',"지출",'+C_+',"외식/배달/카페")',{fmt:WON});
  put(sh,'G7','=SUMIF('+KIND_RNG+',"고정",'+S_DASH+'!$C$'+catFirst+':$C$'+catLast+')',{fmt:WON});
  put(sh,'G8','=SUMIF('+KIND_RNG+',"변동",'+S_DASH+'!$C$'+catFirst+':$C$'+catLast+')',{fmt:WON});
  put(sh,'G9','=COUNTIFS('+S_DASH+'!$E$'+catFirst+':$E$'+catLast+',">1")',{fmt:'0'});
  put(sh,'G10','=INDEX('+S_DASH+'!$B$'+catFirst+':$B$'+catLast+',MATCH(MAX('+S_DASH+'!$C$'+catFirst+':$C$'+catLast+'),'+S_DASH+'!$C$'+catFirst+':$C$'+catLast+',0))');
  put(sh,'G11','=MAX('+S_DASH+'!$C$'+catFirst+':$C$'+catLast+')',{fmt:WON});

  // 진단 표
  section(sh,'B12','■ 자동 진단 (🟢 잘하고 있어요 / 🟡 관심 / 🔴 개선 필요)');
  ['항목','상태','진단','제안'].forEach(function(h,i){ hdr(sh, colLetter_(2+i)+'13', h); });
  var rate=D+'$L$9', execr=S_DASH+'!$E$'+srD, df=14;
  var diag=[
    ['저축률',
     '=IF('+rate+'>=0.2,"🟢 우수",IF('+rate+'>=0.1,"🟡 관심","🔴 주의"))',
     '="번 돈의 "&TEXT('+rate+',"0.0%")&"를 모으고 있어요."',
     '=IF('+rate+'>=0.2,"권장선(20%)을 넘는 훌륭한 저축률입니다. 이 페이스를 유지하세요.",IF('+rate+'>=0.1,"양호합니다. 고정비를 한 단계 줄여 20%를 목표로 해보세요.","지출보다 저축이 적습니다. 고정비·변동비를 함께 점검해 10%부터 만들어요."))'],
    ['예산 집행률',
     '=IF('+execr+'<=1,"🟢 우수",IF('+execr+'<=1.1,"🟡 관심","🔴 주의"))',
     '="소비지출이 예산의 "&TEXT('+execr+',"0.0%")&" 수준이에요."',
     '=IF('+execr+'<=1,"예산 안에서 잘 운영 중입니다.","예산을 초과했습니다. 아래 초과 카테고리부터 조정해보세요.")'],
    ['예산 초과 카테고리',
     '=IF($G$9=0,"🟢 우수",IF($G$9<=2,"🟡 관심","🔴 주의"))',
     '=IF($G$9=0,"예산을 넘긴 카테고리가 없습니다.",TEXT($G$9,"0")&"개 카테고리가 예산을 초과했어요.")',
     '=IF($G$9=0,"전 카테고리가 계획 범위 안에 있습니다. 잘하고 있어요!","[대시보드]에서 빨갛게 표시된 항목을 우선 점검하세요.")'],
    ['최대 지출 항목',
     '="🟡 관심"',
     '=$G$10&" 에 가장 많이 썼어요 ("&TEXT($G$11,"₩#,##0")&")."',
     '=IF($G$10="저축/투자","저축이 최대 지출입니다 — 매우 바람직한 형태예요!","이 항목이 전체 소비의 "&TEXT(IFERROR($G$11/'+D+'$L$8,0),"0.0%")&"를 차지합니다. 과도하면 1순위로 조정하세요.")'],
    ['외식·배달 비중',
     '=IF(IFERROR($G$6/'+D+'$L$8,0)<=0.12,"🟢 우수",IF(IFERROR($G$6/'+D+'$L$8,0)<=0.2,"🟡 관심","🔴 주의"))',
     '="소비 중 외식·배달이 "&TEXT(IFERROR($G$6/'+D+'$L$8,0),"0.0%")&" 예요."',
     '=IF(IFERROR($G$6/'+D+'$L$8,0)<=0.12,"외식 지출이 안정적입니다.","외식·배달 빈도를 주 1회 줄이면 가장 빠르게 절약됩니다.")'],
    ['고정비 비중(예산)',
     '=IF(IFERROR($G$7/($G$7+$G$8),0)<=0.5,"🟢 우수",IF(IFERROR($G$7/($G$7+$G$8),0)<=0.6,"🟡 관심","🔴 주의"))',
     '="예산상 고정비 비중이 "&TEXT(IFERROR($G$7/($G$7+$G$8),0),"0.0%")&" 입니다."',
     '=IF(IFERROR($G$7/($G$7+$G$8),0)<=0.5,"고정비 비중이 건강합니다.","고정비(통신·구독·보험)를 1년에 한 번 갈아타기/해지로 점검하세요.")']
  ];
  diag.forEach(function(row,i){
    var rr=df+i, z=(i%2===0)?ZEBRA:WHITE;
    put(sh,'B'+rr, row[0], {size:10,bold:true,fill:z,border:true});
    put(sh,'C'+rr, row[1], {size:13,fill:z,align:'center',border:true});
    put(sh,'D'+rr, row[2], {size:10,fill:z,wrap:true,border:true});
    put(sh,'E'+rr, row[3], {size:10,fill:z,wrap:true,border:true});
    sh.setRowHeight(rr, rh(30));
  });
  var dl=df+diag.length-1;
  var rules=sh.getConditionalFormatRules();
  var stRange=[sh.getRange('C'+df+':C'+dl)];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('우수').setBackground('#C6EFCE').setFontColor('#006100').setBold(true).setRanges(stRange).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('관심').setBackground('#FFEB9C').setFontColor('#9C6500').setBold(true).setRanges(stRange).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('주의').setBackground('#FFC7CE').setFontColor('#9C0006').setBold(true).setRanges(stRange).build());

  // 한 줄 요약
  var sm=dl+2;
  sh.getRange('B'+sm+':E'+sm).merge();
  section(sh,'B'+sm,'■ 한 줄 요약');
  sh.getRange('B'+(sm+1)+':E'+(sm+2)).merge();
  put(sh,'B'+(sm+1),
    '="이번 달 저축률은 "&TEXT('+rate+',"0.0%")&", 예산 집행률은 "&TEXT('+execr+',"0.0%")&"입니다. "&IF('+rate+'>=0.2,"저축 습관이 아주 좋아요. ",IF('+rate+'>=0.1,"저축을 조금만 더 끌어올리면 좋아요. ","저축 여력 확보가 최우선입니다. "))&IF($G$9=0,"예산도 잘 지키고 있습니다.",TEXT($G$9,"0")&"개 카테고리만 조정하면 균형이 맞습니다.")',
    {size:11,bold:true,color:'#7D6608',fill:'#FCF3CF',wrap:true,border:true});

  // 정기지출 체크리스트
  var fixed=expenseCats.filter(function(c){return c[1]==='고정';}).map(function(c){return c[0];});
  var rc=sm+4;
  section(sh,'B'+rc,'📌 이번 달 정기지출 체크 (고정비 누락 방지)');
  ['항목','월 예산','상태'].forEach(function(h,i){ hdr(sh, colLetter_(2+i)+(rc+1), h); });
  fixed.forEach(function(it,i){
    var r=rc+2+i, z=(i%2===0)?ZEBRA:WHITE;
    put(sh,'B'+r, it, {size:10,fill:z,border:true});
    put(sh,'C'+r, '=IFERROR(VLOOKUP(B'+r+','+CATTBL+',2,0),0)',{size:10,fill:z,align:'right',border:true,fmt:WON});
    put(sh,'D'+r, '=IF(COUNTIFS('+C_+',B'+r+','+I_+','+D+'$L$1,'+J_+','+D+'$L$2)>0,"✅ 완료","⬜ 미입력")',{size:10,fill:z,align:'center',border:true});
  });
  var dlast=rc+1+fixed.length;
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('미입력').setBackground('#FBD5D5').setFontColor('#C0392B').setBold(true).setRanges([sh.getRange('D'+(rc+2)+':D'+dlast)]).build());
  sh.setConditionalFormatRules(rules);
}

// =====================================================================
// 6) 시작하기
function buildGuide_(sh){
  sh.setHiddenGridlines(true);
  sh.setTabColor('#4A6FA5');
  sh.setColumnWidth(colNum_('A'), cw(2.5));
  sh.setColumnWidth(colNum_('B'), cw(26));
  sh.setColumnWidth(colNum_('C'), cw(60));
  sh.getRange('B1:C1').merge();
  put(sh,'B1','🏠  우리집 가계부',{size:26,bold:true,color:HEAD_TX,fill:NAVY});
  sh.setRowHeight(1, rh(52));
  sh.getRange('B2:C2').merge();
  put(sh,'B2','가족이 함께 쓰는 대시보드형 가계부 — 입력은 한 곳, 분석은 자동',{size:11,bold:true,color:GRAY});

  function sec(r,t){ sh.getRange('B'+r+':C'+r).merge(); put(sh,'B'+r,t,{size:12,bold:true,color:HEAD_TX,fill:HEAD_BG}); sh.setRowHeight(r, rh(22)); }
  function row2(r,a,b,fill){ put(sh,'B'+r,a,{size:10,bold:true,fill:fill||WHITE,wrap:true,border:true}); put(sh,'C'+r,b,{size:10,fill:fill||WHITE,wrap:true,border:true}); }

  sec(4,'📌 사용 순서 (딱 3단계)');
  row2(5,'① [설정] 시트','구성원·결제수단·카테고리별 월 예산을 우리 집에 맞게 채웁니다. 기준 연/월도 여기서 지정.',ZEBRA);
  row2(6,'② [거래내역] 시트','돈을 쓰거나 벌 때마다 한 줄씩 입력. 날짜·유형·대분류·금액만 넣어도 충분합니다.');
  row2(7,'③ [대시보드]·[기간리포트]·[진단&인사이트]','입력만 하면 자동으로 집계·차트·진단이 갱신됩니다. 직접 계산할 필요 없어요.',ZEBRA);

  sec(9,'📑 시트 안내');
  row2(10,'⚙️ 설정','모든 기준값의 출발점. 여기 값만 바꾸면 전체가 연동됩니다.',ZEBRA);
  row2(11,'✍️ 거래내역','실제 입력 화면. 드롭다운으로 빠르게 기록.');
  row2(12,'📊 대시보드','이번 달 핵심 요약(KPI 카드 6개 + 카테고리/구성원/주차 차트).',ZEBRA);
  row2(13,'📈 기간리포트','주간·월간·분기·연간 표와 그래프.');
  row2(14,'🔍 진단&인사이트','잘하는 점/개선점을 자동 진단하고 절약 팁을 제안.',ZEBRA);

  sec(16,'🖱️ 입력 팁');
  row2(17,'유형',"수입 / 지출 / 이체 중 선택. '이체'는 계좌 간 이동 등 분석 제외 항목.",ZEBRA);
  row2(18,'대분류','드롭다운에서 선택. 카테고리는 [설정]에서 자유롭게 수정 가능.');
  row2(19,'저축/투자',"적금·투자금도 '지출-저축/투자'로 기록하세요. 저축률에 반영됩니다.",ZEBRA);
  row2(20,'금액','숫자만 입력(원). 자동으로 ₩ 표기됩니다.');

  sec(22,'🎨 색상 범례');
  put(sh,'B23','노랑 칸 + 파란 글씨',{size:10,bold:true,color:INPUTC,fill:INPUTF,align:'center',border:true});
  put(sh,'C23','직접 입력하는 칸 (예산·기준연월 등)',{size:10,fill:WHITE,wrap:true,border:true});
  put(sh,'B24','그 외 검정 글씨',{size:10,fill:WHITE,align:'center',border:true});
  put(sh,'C24','자동 계산 결과 — 건드리지 않아도 됩니다.',{size:10,fill:ZEBRA,wrap:true,border:true});
  put(sh,'B25','🔴 빨강 강조',{size:10,bold:true,color:'#C00000',fill:WHITE,align:'center',border:true});
  put(sh,'C25','예산 초과 등 주의가 필요한 항목.',{size:10,fill:WHITE,wrap:true,border:true});

  sh.getRange('B27:C27').merge();
  put(sh,'B27','TIP. 1년 이상 쓰려면 [거래내역]의 빈 행을 계속 이어서 입력하면 됩니다 (수식 자동 적용, 약 1,200건).',{size:9,bold:true,color:GRAY,wrap:true});
}

// ---------- 열문자 유틸 ----------
function colNum_(letter){ var n=0; for(var i=0;i<letter.length;i++) n=n*26+(letter.charCodeAt(i)-64); return n; }
function colLetter_(num){ var s=''; while(num>0){ var m=(num-1)%26; s=String.fromCharCode(65+m)+s; num=Math.floor((num-1)/26); } return s; }
