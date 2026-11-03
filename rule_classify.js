const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');


// ============================================================
// 기본 설정
// ============================================================

const OUTPUT_DIR =
  path.join(__dirname, 'output');

const DB_FILE_NAME =
  '생보협회_상품DB.xlsx';

const UPLOAD_URL =
  process.env.UPLOAD_URL;

const UPLOAD_SECRET =
  process.env.UPLOAD_SECRET;

const TARGET_MONTH_OVERRIDE =
  String(
    process.env.TARGET_MONTH || ''
  ).trim();


// ============================================================
// AI 심층분석 고정 5개사
// ============================================================

const AI_TARGET_COMPANIES = [
  '삼성생명',
  '한화생명',
  '교보생명',
  '신한라이프',
  'NH농협생명'
];


// ============================================================
// 기본 함수
// ============================================================

function clean(value) {

  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function norm(value) {

  return clean(value)
    .toLowerCase()
    .replace(
      /[\s\(\)\[\]\{\}·ㆍ\-_,.\/\\:：'"]/g,
      ''
    );
}


function getTargetMonth() {

  if (TARGET_MONTH_OVERRIDE) {

    if (
      !/^\d{4}-\d{2}$/.test(
        TARGET_MONTH_OVERRIDE
      )
    ) {

      throw new Error(
        'TARGET_MONTH는 YYYY-MM 형식이어야 합니다.'
      );
    }

    return TARGET_MONTH_OVERRIDE;
  }


  const parts =
    Object.fromEntries(

      new Intl.DateTimeFormat(
        'en-US',
        {
          timeZone: 'Asia/Seoul',
          year: 'numeric',
          month: '2-digit'
        }
      )
        .formatToParts(new Date())
        .map(
          p => [p.type, p.value]
        )
    );


  let year =
    Number(parts.year);

  let month =
    Number(parts.month) - 1;


  if (month === 0) {

    year -= 1;
    month = 12;
  }


  return (
    `${year}-` +
    String(month)
      .padStart(2, '0')
  );
}


// ============================================================
// 보험회사명 통일
// ============================================================

function canonicalCompany(value) {

  const s =
    norm(value);


  if (
    s.includes(
      '삼성생명'
    )
  ) {
    return '삼성생명';
  }


  if (
    s.includes(
      '한화생명'
    )
  ) {
    return '한화생명';
  }


  if (
    s.includes(
      '교보생명'
    )
  ) {
    return '교보생명';
  }


  if (
    s.includes(
      '신한라이프'
    ) ||
    s.includes(
      '신한생명'
    )
  ) {
    return '신한라이프';
  }


  if (
    s.includes(
      'nh농협생명'
    ) ||
    s.includes(
      '농협생명'
    )
  ) {
    return 'NH농협생명';
  }


  return clean(value);
}


function isAiTargetCompany(
  value
) {

  return AI_TARGET_COMPANIES
    .includes(
      canonicalCompany(
        value
      )
    );
}


// ============================================================
// Sheet 읽기
// ============================================================

function sheetRows(
  workbook,
  sheetName
) {

  const sheet =
    workbook.Sheets[
      sheetName
    ];


  if (!sheet) {
    return [];
  }


  return XLSX.utils
    .sheet_to_json(
      sheet,
      {
        defval: ''
      }
    );
}


// ============================================================
// 필드 찾기
// ============================================================

function getCell(
  row,
  aliases
) {

  for (
    const key
    of aliases
  ) {

    if (
      Object.prototype
        .hasOwnProperty
        .call(
          row,
          key
        )
    ) {

      const value =
        clean(
          row[key]
        );


      if (value) {
        return value;
      }
    }
  }


  return '';
}


// ============================================================
// 상품군 키
// ============================================================

function makeFamilyKey(row) {

  return [

    norm(
      getCell(
        row,
        ['보험구분']
      )
    ),

    norm(
      getCell(
        row,
        ['상품유형']
      )
    ),

    norm(
      canonicalCompany(
        getCell(
          row,
          [
            '보험회사명',
            '회사명',
            '보험회사'
          ]
        )
      )
    ),

    norm(
      getCell(
        row,
        [
          '상품군명',
          '상품명'
        ]
      )
    )

  ].join('|');
}


function makeProductKey(
  company,
  product
) {

  return (
    norm(
      canonicalCompany(
        company
      )
    ) +
    '|' +
    norm(product)
  );
}


// ============================================================
// 상품명 → 상품군 연결표
// ============================================================

function buildProductFamilyMap(
  productRows
) {

  const map =
    new Map();


  for (
    const row
    of productRows
  ) {

    const company =
      getCell(
        row,
        [
          '보험회사명',
          '회사명',
          '보험회사'
        ]
      );


    const product =
      getCell(
        row,
        ['상품명']
      );


    if (
      !company ||
      !product
    ) {
      continue;
    }


    map.set(

      makeProductKey(
        company,
        product
      ),

      makeFamilyKey(
        row
      )
    );
  }


  return map;
}


// ============================================================
// 급부 필드
// ============================================================

function benefitName(row) {

  return getCell(
    row,
    [
      '급부명칭',
      '급부명',
      '급부',
      '보장명',
      '담보명'
    ]
  );
}


function benefitReason(row) {

  return getCell(
    row,
    [
      '지급사유',
      '보험금지급사유',
      '급부지급사유',
      '지급조건'
    ]
  );
}


function benefitAmount(row) {

  return getCell(
    row,
    [
      '지급금액',
      '보험금액',
      '급부금액',
      '보장금액'
    ]
  );
}


// ============================================================
// 상품군별 급부정보 생성
// ============================================================

function buildBenefitMap(
  productRows,
  detailRows
) {

  const productFamilyMap =
    buildProductFamilyMap(
      productRows
    );


  const result =
    new Map();


  for (
    const row
    of detailRows
  ) {

    const company =
      getCell(
        row,
        [
          '보험회사명',
          '회사명',
          '보험회사'
        ]
      );


    const product =
      getCell(
        row,
        ['상품명']
      );


    if (
      !company ||
      !product
    ) {
      continue;
    }


    const familyKey =
      productFamilyMap.get(
        makeProductKey(
          company,
          product
        )
      );


    if (!familyKey) {
      continue;
    }


    if (
      !result.has(
        familyKey
      )
    ) {

      result.set(
        familyKey,
        {
          names:
            new Map(),

          signatures:
            new Set()
        }
      );
    }


    const item =
      result.get(
        familyKey
      );


    const name =
      benefitName(
        row
      );


    const reason =
      benefitReason(
        row
      );


    const amount =
      benefitAmount(
        row
      );


    /*
     * 급부명은
     * 정규화값 → 실제 표시명
     */
    if (name) {

      const normalized =
        norm(name);


      if (
        !item.names.has(
          normalized
        )
      ) {

        item.names.set(
          normalized,
          name
        );
      }
    }


    /*
     * 급부명 + 지급사유 + 지급금액
     * 전체 조합
     */
    if (
      name ||
      reason ||
      amount
    ) {

      item.signatures.add(

        [
          norm(name),
          norm(reason),
          norm(amount)
        ].join('|')
      );
    }
  }


  return result;
}


// ============================================================
// Apps Script API
// ============================================================

async function apiCall(
  payload
) {

  if (
    !UPLOAD_URL ||
    !UPLOAD_SECRET
  ) {

    throw new Error(
      'UPLOAD_URL 또는 UPLOAD_SECRET이 없습니다.'
    );
  }


  const response =
    await fetch(
      UPLOAD_URL,
      {
        method:
          'POST',

        headers: {
          'Content-Type':
            'application/json'
        },

        body:
          JSON.stringify({
            secret:
              UPLOAD_SECRET,

            ...payload
          })
      }
    );


  const text =
    await response.text();


  let result;


  try {

    result =
      JSON.parse(
        text
      );

  } catch {

    throw new Error(
      'Apps Script 응답 오류: ' +
      text.slice(
        0,
        300
      )
    );
  }


  if (
    !response.ok ||
    !result.ok
  ) {

    throw new Error(
      result.error ||
      `HTTP ${response.status}`
    );
  }


  return result;
}


// ============================================================
// 과거 월별 통합본 다운로드
// ============================================================

async function downloadMonthlyFile(
  month
) {

  const fileName =
    `${month}_생보협회_상품통합.xlsx`;


  const result =
    await apiCall({

      action:
        'download',

      folderType:
        'merged',

      fileName
    });


  if (
    !result.exists
  ) {

    return null;
  }


  const filePath =
    path.join(
      OUTPUT_DIR,
      `_history_${month}.xlsx`
    );


  fs.writeFileSync(

    filePath,

    Buffer.from(
      result.base64,
      'base64'
    )
  );


  return filePath;
}


// ============================================================
// Drive Upload
// ============================================================

async function uploadFile(
  filePath,
  fileName
) {

  const base64 =
    fs
      .readFileSync(
        filePath
      )
      .toString(
        'base64'
      );


  await apiCall({

    action:
      'upload',

    folderType:
      'merged',

    fileName,

    mimeType:
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

    base64
  });


  console.log(
    `Drive 저장 완료: ${fileName}`
  );
}


// ============================================================
// 누적 DB에서 과거월 추출
// ============================================================

function getHistoricalMonths(
  targetMonth
) {

  const dbPath =
    path.join(
      OUTPUT_DIR,
      DB_FILE_NAME
    );


  if (
    !fs.existsSync(
      dbPath
    )
  ) {

    return [];
  }


  const workbook =
    XLSX.readFile(
      dbPath
    );


  const rows =
    sheetRows(
      workbook,
      '누적상품DB'
    );


  return [

    ...new Set(

      rows
        .map(
          row =>
            clean(
              row['조사월']
            )
        )
        .filter(
          month =>
            /^\d{4}-\d{2}$/.test(
              month
            ) &&
            month < targetMonth
        )
    )

  ]
    .sort()
    .reverse();
}


// ============================================================
// 과거 상품군 급부정보
// ============================================================

async function buildHistoricalBenefitMap(
  targetMonth
) {

  const months =
    getHistoricalMonths(
      targetMonth
    );


  const history =
    new Map();


  console.log(
    `급부 비교 과거월: ${
      months.join(', ') ||
      '없음'
    }`
  );


  /*
   * 최신 과거월부터 확인
   */
  for (
    const month
    of months
  ) {

    let filePath;


    try {

      filePath =
        await downloadMonthlyFile(
          month
        );

    } catch (error) {

      console.log(
        `${month} 통합본 다운로드 실패 → 건너뜀`
      );

      continue;
    }


    if (!filePath) {

      console.log(
        `${month} 통합본 없음`
      );

      continue;
    }


    const workbook =
      XLSX.readFile(
        filePath
      );


    const products =
      sheetRows(
        workbook,
        '상품목록'
      );


    const details =
      sheetRows(
        workbook,
        '급부상세'
      );


    const benefitMap =
      buildBenefitMap(
        products,
        details
      );


    /*
     * 최신 과거월 자료를 우선 사용
     */
    for (
      const [
        familyKey,
        benefitInfo
      ]
      of benefitMap.entries()
    ) {

      if (
        !history.has(
          familyKey
        )
      ) {

        history.set(
          familyKey,
          {
            month,
            ...benefitInfo
          }
        );
      }
    }
  }


  return history;
}


// ============================================================
// 규칙 분류
// ============================================================

function classifyFamily(
  familyRow,
  currentBenefits,
  priorBenefits
) {

  const comparison =
    clean(
      familyRow[
        '비교상태'
      ]
    );


  /*
   * build_db에서
   * 과거 동일 상품군 자체가 없음
   */
  if (
    comparison ===
    '신규상품후보'
  ) {

    return {

      type:
        '신규상품후보',

      reason:
        '과거월 누적DB에 동일 회사·상품군이 없음',

      newBenefits:
        []
    };
  }


  /*
   * 현재 급부와
   * 과거 동일 상품군 급부가 모두 존재
   */
  if (
    currentBenefits &&
    priorBenefits
  ) {

    const newBenefits =
      [];


    for (
      const [
        normalizedName,
        displayName
      ]
      of currentBenefits.names
    ) {

      if (
        !priorBenefits
          .names
          .has(
            normalizedName
          )
      ) {

        newBenefits.push(
          displayName
        );
      }
    }


    /*
     * 과거에 없던 급부명 발생
     */
    if (
      newBenefits.length > 0
    ) {

      return {

        type:
          '신규급부·특약후보',

        reason:
          `${priorBenefits.month} 대비 신규 급부명 확인`,

        newBenefits:
          [
            ...new Set(
              newBenefits
            )
          ]
      };
    }


    /*
     * 급부명은 같지만
     * 지급사유/지급금액 조합 변화
     */
    const changedSignature =
      [
        ...currentBenefits
          .signatures
      ]
        .some(
          signature =>
            !priorBenefits
              .signatures
              .has(
                signature
              )
        );


    if (
      changedSignature
    ) {

      return {

        type:
          '상품개정후보',

        reason:
          `${priorBenefits.month} 대비 지급사유 또는 지급금액 변화 확인`,

        newBenefits:
          []
      };
    }


    /*
     * 동일 급부구조
     */
    return {

      type:
        '동일상품·유형변경',

      reason:
        `${priorBenefits.month} 대비 급부구조 동일`,

      newBenefits:
        []
    };
  }


  /*
   * 과거 급부통합본은 없지만
   * DB에서 상품군 개정으로 판단
   */
  if (
    comparison ===
    '상품개정후보'
  ) {

    return {

      type:
        '상품개정후보',

      reason:
        '과거 동일 상품군이 있으나 급부 비교자료가 충분하지 않음',

      newBenefits:
        []
    };
  }


  if (
    comparison ===
    '동일상품 재판매·개정후보'
  ) {

    return {

      type:
        '동일상품·유형변경',

      reason:
        '과거 동일 회사·상품명이 확인됨',

      newBenefits:
        []
    };
  }


  /*
   * 첫 기준월
   */
  if (
    comparison ===
    '기준월'
  ) {

    return {

      type:
        '기준월·비교자료없음',

      reason:
        '대상월보다 이전 누적자료가 없는 기준월',

      newBenefits:
        []
    };
  }


  return {

    type:
      '확인필요',

    reason:
      '규칙분류만으로 판단하기 어려움',

    newBenefits:
      []
  };
}


// ============================================================
// AI 대상 판정
// ============================================================

function decideAiTarget(
  familyRow,
  classification
) {

  const company =
    canonicalCompany(
      familyRow[
        '보험회사명'
      ]
    );


  /*
   * 5개사 아니면 무조건 비대상
   */
  if (
    !isAiTargetCompany(
      company
    )
  ) {

    return {

      target:
        false,

      reason:
        'AI 고정 5개사 외'
    };
  }


  /*
   * 5개사 중에서도
   * 신규상품 / 신규급부만
   */
  if (
    ![
      '신규상품후보',
      '신규급부·특약후보'
    ].includes(
      classification.type
    )
  ) {

    return {

      target:
        false,

      reason:
        `${company} / ${classification.type}`
    };
  }


  return {

    target:
      true,

    reason:
      `${company} / ${classification.type}`
  };
}


// ============================================================
// Worksheet 교체
// ============================================================

function replaceSheet(
  workbook,
  sheetName,
  rows
) {

  const oldIndex =
    workbook.SheetNames
      .indexOf(
        sheetName
      );


  if (
    oldIndex >= 0
  ) {

    workbook.SheetNames
      .splice(
        oldIndex,
        1
      );


    delete workbook.Sheets[
      sheetName
    ];
  }


  const sheet =
    XLSX.utils
      .json_to_sheet(
        rows
      );


  if (
    rows.length > 0
  ) {

    const headers =
      Object.keys(
        rows[0]
      );


    sheet['!autofilter'] = {

      ref:
        XLSX.utils
          .encode_range({

            s: {
              r: 0,
              c: 0
            },

            e: {
              r:
                rows.length,

              c:
                headers.length - 1
            }
          })
    };


    sheet['!cols'] =
      headers.map(
        header => {

          if (
            [
              '상품명',
              '상품군명',
              '세부상품명'
            ].includes(
              header
            )
          ) {

            return {
              wch: 45
            };
          }


          if (
            [
              '규칙분류근거',
              'AI대상사유',
              '신규급부목록'
            ].includes(
              header
            )
          ) {

            return {
              wch: 40
            };
          }


          return {
            wch: 18
          };
        }
      );
  }


  XLSX.utils
    .book_append_sheet(
      workbook,
      sheet,
      sheetName
    );
}


// ============================================================
// 누적 DB에도 규칙분류 반영
// ============================================================

function updateDbRules(
  targetMonth,
  resultMap
) {

  const dbPath =
    path.join(
      OUTPUT_DIR,
      DB_FILE_NAME
    );


  if (
    !fs.existsSync(
      dbPath
    )
  ) {
    return;
  }


  const workbook =
    XLSX.readFile(
      dbPath
    );


  const rows =
    sheetRows(
      workbook,
      '누적상품DB'
    );


  for (
    const row
    of rows
  ) {

    if (
      clean(
        row['조사월']
      ) !==
      targetMonth
    ) {

      continue;
    }


    const result =
      resultMap.get(
        makeFamilyKey(
          row
        )
      );


    if (!result) {
      continue;
    }


    row['규칙분류'] =
      result.type;


    row['규칙분류근거'] =
      result.reason;


    row['신규급부목록'] =
      result.newBenefits;


    row['AI대상여부'] =
      result.aiTarget;


    row['AI대상사유'] =
      result.aiReason;
  }


  replaceSheet(
    workbook,
    '누적상품DB',
    rows
  );


  XLSX.writeFile(
    workbook,
    dbPath
  );
}


// ============================================================
// 실행
// ============================================================

async function main() {

  const targetMonth =
    getTargetMonth();


  console.log(
    `규칙분류 대상월: ${targetMonth}`
  );


  const monthlyPath =
    path.join(
      OUTPUT_DIR,
      `${targetMonth}_생보협회_상품통합.xlsx`
    );


  if (
    !fs.existsSync(
      monthlyPath
    )
  ) {

    throw new Error(
      `통합파일 없음: ${monthlyPath}`
    );
  }


  const workbook =
    XLSX.readFile(
      monthlyPath
    );


  const productRows =
    sheetRows(
      workbook,
      '상품목록'
    );


  const familyRows =
    sheetRows(
      workbook,
      '상품군목록'
    );


  const detailRows =
    sheetRows(
      workbook,
      '급부상세'
    );


  console.log(
    `전체 상품군: ${familyRows.length}개`
  );


  console.log(
    `급부상세: ${detailRows.length}행`
  );


  /*
   * 현재월 상품군별 급부
   */
  const currentBenefitMap =
    buildBenefitMap(
      productRows,
      detailRows
    );


  /*
   * 과거월 최신 급부
   */
  const historicalMap =
    await buildHistoricalBenefitMap(
      targetMonth
    );


  const resultMap =
    new Map();


  const classificationCounts =
    new Map();


  let aiTargetCount = 0;


  /*
   * 상품군 단위 분류
   */
  for (
    const row
    of familyRows
  ) {

    const familyKey =
      makeFamilyKey(
        row
      );


    const classification =
      classifyFamily(

        row,

        currentBenefitMap.get(
          familyKey
        ),

        historicalMap.get(
          familyKey
        )
      );


    row['규칙분류'] =
      classification.type;


    row['규칙분류근거'] =
      classification.reason;


    row['신규급부목록'] =
      classification
        .newBenefits
        .join(' | ');


    const aiDecision =
      decideAiTarget(
        row,
        classification
      );


    row['AI대상여부'] =
      aiDecision.target
        ? 'Y'
        : 'N';


    row['AI대상사유'] =
      aiDecision.reason;


    /*
     * 아직 AI 실행 전
     */
    row['AI보완상태'] =
      aiDecision.target
        ? '대기'
        : '비대상';


    if (
      aiDecision.target
    ) {

      aiTargetCount += 1;
    }


    classificationCounts.set(

      classification.type,

      (
        classificationCounts.get(
          classification.type
        ) || 0
      ) + 1
    );


    resultMap.set(
      familyKey,
      {

        type:
          classification.type,

        reason:
          classification.reason,

        newBenefits:
          classification
            .newBenefits
            .join(' | '),

        aiTarget:
          aiDecision.target
            ? 'Y'
            : 'N',

        aiReason:
          aiDecision.reason
      }
    );
  }


  /*
   * 상품목록에도
   * 상품군 분류결과 연결
   */
  for (
    const row
    of productRows
  ) {

    const result =
      resultMap.get(
        makeFamilyKey(
          row
        )
      );


    if (!result) {
      continue;
    }


    row['규칙분류'] =
      result.type;


    row['규칙분류근거'] =
      result.reason;


    row['신규급부목록'] =
      result.newBenefits;


    row['AI대상여부'] =
      result.aiTarget;


    row['AI대상사유'] =
      result.aiReason;


    row['AI보완상태'] =
      result.aiTarget === 'Y'
        ? '대기'
        : '비대상';
  }


  /*
   * 상품군목록 교체
   */
  replaceSheet(
    workbook,
    '상품군목록',
    familyRows
  );


  /*
   * 상품목록 교체
   */
  replaceSheet(
    workbook,
    '상품목록',
    productRows
  );


  // ==========================================================
  // 규칙분류 요약
  // ==========================================================

  const summaryRows = [

    {
      구분:
        '전체 상품군',

      건수:
        familyRows.length
    },

    {
      구분:
        'AI 심층분석 대상',

      건수:
        aiTargetCount
    },

    ...[
      ...classificationCounts
        .entries()
    ].map(
      ([type, count]) => ({

        구분:
          type,

        건수:
          count
      })
    )
  ];


  replaceSheet(
    workbook,
    '규칙분류요약',
    summaryRows
  );


  // ==========================================================
  // AI 대상 목록 시트
  // ==========================================================

  const aiTargetRows =
    familyRows
      .filter(
        row =>
          row[
            'AI대상여부'
          ] === 'Y'
      )
      .map(
        row => ({

          조사월:
            row['조사월'],

          보험회사명:
            canonicalCompany(
              row['보험회사명']
            ),

          상품유형:
            row['상품유형'],

          상품군명:
            row['상품군명'],

          판매일자:
            row['판매일자'],

          세부상품명:
            row['세부상품명'],

          규칙분류:
            row['규칙분류'],

          규칙분류근거:
            row['규칙분류근거'],

          신규급부목록:
            row['신규급부목록'],

          AI대상사유:
            row['AI대상사유']
        })
      );


  replaceSheet(
    workbook,
    'AI분석대상',
    aiTargetRows
  );


  /*
   * 파일 저장
   */
  XLSX.writeFile(
    workbook,
    monthlyPath
  );


  /*
   * 누적 DB도 갱신
   */
  updateDbRules(
    targetMonth,
    resultMap
  );


  /*
   * Drive 월별 통합본 업로드
   */
  await uploadFile(

    monthlyPath,

    path.basename(
      monthlyPath
    )
  );


  /*
   * Drive 누적 DB 업로드
   */
  const dbPath =
    path.join(
      OUTPUT_DIR,
      DB_FILE_NAME
    );


  if (
    fs.existsSync(
      dbPath
    )
  ) {

    await uploadFile(
      dbPath,
      DB_FILE_NAME
    );
  }


  // ==========================================================
  // 로그
  // ==========================================================

  console.log('');
  console.log(
    '============================'
  );

  console.log(
    `전체 상품군: ${familyRows.length}개`
  );


  for (
    const [
      type,
      count
    ]
    of classificationCounts
  ) {

    console.log(
      `${type}: ${count}개`
    );
  }


  console.log(
    `AI 심층분석 대상: ${aiTargetCount}개`
  );


  console.log(
    'AI 대상 고정회사:'
  );


  AI_TARGET_COMPANIES
    .forEach(
      company => {

        const count =
          aiTargetRows
            .filter(
              row =>
                row[
                  '보험회사명'
                ] === company
            )
            .length;


        console.log(
          `- ${company}: ${count}개`
        );
      }
    );


  console.log(
    '============================'
  );

  console.log(
    '전체 생보사 규칙분류 완료'
  );
}


// ============================================================
// Run
// ============================================================

main()
  .catch(
    error => {

      console.error(
        'rule_classify.js 실행 실패'
      );

      console.error(
        error
      );

      process.exit(1);
    }
  );
