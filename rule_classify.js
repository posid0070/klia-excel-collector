const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const OUTPUT_DIR = path.join(__dirname, 'output');

const UPLOAD_URL = process.env.UPLOAD_URL;
const UPLOAD_SECRET = process.env.UPLOAD_SECRET;
const TARGET_MONTH_OVERRIDE =
  String(process.env.TARGET_MONTH || '').trim();

const DB_FILE_NAME = '생보협회_상품DB.xlsx';


// ============================================================
// AI 심층분석 대상 5개사 - 앞으로 고정
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
    return TARGET_MONTH_OVERRIDE;
  }

  const now = new Date();

  const parts =
    new Intl.DateTimeFormat(
      'en-US',
      {
        timeZone: 'Asia/Seoul',
        year: 'numeric',
        month: '2-digit'
      }
    )
      .formatToParts(now)
      .reduce(
        (acc, p) => {
          acc[p.type] = p.value;
          return acc;
        },
        {}
      );

  let year = Number(parts.year);
  let month = Number(parts.month) - 1;

  if (month === 0) {
    year -= 1;
    month = 12;
  }

  return (
    `${year}-` +
    String(month).padStart(2, '0')
  );
}


function getCell(row, aliases) {

  for (const key of aliases) {

    if (
      Object.prototype.hasOwnProperty.call(
        row,
        key
      ) &&
      clean(row[key])
    ) {
      return clean(row[key]);
    }
  }

  return '';
}


// ============================================================
// 보험회사명 정규화
// ============================================================

function canonicalCompany(value) {

  const s = norm(value);

  if (s.includes('삼성생명')) {
    return '삼성생명';
  }

  if (s.includes('한화생명')) {
    return '한화생명';
  }

  if (s.includes('교보생명')) {
    return '교보생명';
  }

  if (
    s.includes('신한라이프') ||
    s.includes('신한생명')
  ) {
    return '신한라이프';
  }

  if (
    s.includes('nh농협생명') ||
    s.includes('농협생명')
  ) {
    return 'NH농협생명';
  }

  return clean(value);
}


function isAiTargetCompany(value) {

  return AI_TARGET_COMPANIES.includes(
    canonicalCompany(value)
  );
}


// ============================================================
// Family key
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
      canonicalCompany(company)
    ) +
    '|' +
    norm(product)
  );
}


// ============================================================
// 엑셀 읽기
// ============================================================

function sheetRows(
  workbook,
  sheetName
) {

  const sheet =
    workbook.Sheets[sheetName];

  if (!sheet) {
    return [];
  }

  return XLSX.utils.sheet_to_json(
    sheet,
    {
      defval: ''
    }
  );
}


// ============================================================
// 상품명 → 상품군 연결
// ============================================================

function buildProductFamilyMap(
  productRows
) {

  const map = new Map();

  for (const row of productRows) {

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

    if (!company || !product) {
      continue;
    }

    map.set(
      makeProductKey(
        company,
        product
      ),
      makeFamilyKey(row)
    );
  }

  return map;
}


// ============================================================
// 급부 데이터
// ============================================================

function benefitName(row) {

  return getCell(
    row,
    [
      '급부명칭',
      '급부명',
      '급부',
      '보장명',
      '담보명',
      '보장내용'
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


function buildBenefitMap(
  productRows,
  detailRows
) {

  const productMap =
    buildProductFamilyMap(
      productRows
    );

  const result = new Map();

  for (const row of detailRows) {

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

    if (!company || !product) {
      continue;
    }

    const familyKey =
      productMap.get(
        makeProductKey(
          company,
          product
        )
      );

    if (!familyKey) {
      continue;
    }

    if (!result.has(familyKey)) {

      result.set(
        familyKey,
        {
          names: new Set(),
          signatures: new Set()
        }
      );
    }

    const item =
      result.get(familyKey);

    const name =
      benefitName(row);

    const reason =
      benefitReason(row);

    const amount =
      benefitAmount(row);

    if (name) {
      item.names.add(
        norm(name)
      );
    }

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
// Apps Script Drive API
// ============================================================

async function apiCall(payload) {

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
        method: 'POST',

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
    result = JSON.parse(text);
  } catch {
    throw new Error(
      'Apps Script 응답 오류: ' +
      text.slice(0, 300)
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


async function downloadMonthlyFile(
  month
) {

  const fileName =
    `${month}_생보협회_상품통합.xlsx`;

  const result =
    await apiCall({
      action: 'download',
      folderType: 'merged',
      fileName
    });

  if (!result.exists) {
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


async function uploadFile(
  filePath,
  fileName
) {

  const base64 =
    fs.readFileSync(filePath)
      .toString('base64');

  await apiCall({
    action: 'upload',
    folderType: 'merged',
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
// 누적 DB에서 과거 조사월 찾기
// ============================================================

function getHistoricalMonths(
  targetMonth
) {

  const dbPath =
    path.join(
      OUTPUT_DIR,
      DB_FILE_NAME
    );

  if (!fs.existsSync(dbPath)) {
    return [];
  }

  const wb =
    XLSX.readFile(dbPath);

  const rows =
    sheetRows(
      wb,
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
            /^\d{4}-\d{2}$/.test(month) &&
            month < targetMonth
        )
    )
  ].sort().reverse();
}


// ============================================================
// 과거 상품별 최신 급부정보
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
    `과거 비교월: ${months.join(', ') || '없음'}`
  );

  for (const month of months) {

    let filePath;

    try {

      filePath =
        await downloadMonthlyFile(
          month
        );

    } catch (error) {

      console.log(
        `${month} 다운로드 실패 → 건너뜀`
      );

      continue;
    }

    if (!filePath) {
      continue;
    }

    const wb =
      XLSX.readFile(filePath);

    const products =
      sheetRows(
        wb,
        '상품목록'
      );

    const details =
      sheetRows(
        wb,
        '급부상세'
      );

    const benefitMap =
      buildBenefitMap(
        products,
        details
      );

    /*
     * 최신월부터 읽으므로
     * 처음 나온 상품군만 저장한다.
     */
    for (
      const [
        familyKey,
        benefits
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
            ...benefits
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
  row,
  currentBenefits,
  historicalBenefits
) {

  const comparison =
    clean(
      row['비교상태']
    );

  /*
   * 누적DB상 과거 상품군이 전혀 없음
   */
  if (
    comparison ===
    '신규상품후보'
  ) {

    return {
      type:
        '신규상품후보',

      reason:
        '누적DB에 동일 상품군이 확인되지 않음',

      newBenefits:
        []
    };
  }

  /*
   * 과거 동일 상품군의 급부정보가 있으면
   * 실제 급부를 비교
   */
  if (
    currentBenefits &&
    historicalBenefits
  ) {

    const newBenefits =
      [
        ...currentBenefits.names
      ].filter(
        name =>
          !historicalBenefits
            .names
            .has(name)
      );

    if (
      newBenefits.length > 0
    ) {

      return {
        type:
          '신규급부·특약후보',

        reason:
          `${historicalBenefits.month} 대비 신규 급부명 확인`,

        newBenefits
      };
    }


    const currentSignatures =
      [...currentBenefits.signatures];

    const oldSignatures =
      historicalBenefits.signatures;


    const changed =
      currentSignatures.some(
        sig =>
          !oldSignatures.has(sig)
      );


    if (changed) {

      return {
        type:
          '상품개정후보',

        reason:
          `${historicalBenefits.month} 대비 지급사유 또는 지급금액 변화 가능`,

        newBenefits:
          []
      };
    }


    return {
      type:
        '동일상품·유형변경',

      reason:
        '과거 동일 상품군과 급부구조가 동일함',

      newBenefits:
        []
    };
  }


  /*
   * 급부 과거자료는 없지만
   * build_db상 개정 후보인 경우
   */
  if (
    comparison.includes(
      '개정후보'
    )
  ) {

    return {
      type:
        '상품개정후보',

      reason:
        `누적DB 비교결과: ${comparison}`,

      newBenefits:
        []
    };
  }


  if (
    comparison.includes(
      '재판매'
    )
  ) {

    return {
      type:
        '동일상품·유형변경',

      reason:
        `누적DB 비교결과: ${comparison}`,

      newBenefits:
        []
    };
  }


  /*
   * 최초 기준월
   */
  if (
    comparison ===
    '기준월'
  ) {

    return {
      type:
        '기준월·비교자료없음',

      reason:
        '누적 비교자료가 없는 최초 기준월',

      newBenefits:
        []
    };
  }


  return {
    type:
      '확인필요',

    reason:
      '규칙만으로 분류하기 어려움',

    newBenefits:
      []
  };
}


// ============================================================
// Worksheet
// ============================================================

function replaceSheet(
  workbook,
  sheetName,
  rows
) {

  const oldIndex =
    workbook.SheetNames.indexOf(
      sheetName
    );

  if (oldIndex >= 0) {

    workbook.SheetNames.splice(
      oldIndex,
      1
    );

    delete workbook.Sheets[
      sheetName
    ];
  }

  const sheet =
    XLSX.utils.json_to_sheet(
      rows
    );

  if (
    rows.length > 0
  ) {

    const headers =
      Object.keys(rows[0]);

    sheet['!autofilter'] = {
      ref:
        XLSX.utils.encode_range({
          s: {
            r: 0,
            c: 0
          },
          e: {
            r: rows.length,
            c:
              headers.length - 1
          }
        })
    };
  }

  XLSX.utils.book_append_sheet(
    workbook,
    sheet,
    sheetName
  );
}


// ============================================================
// DB 현재월에도 규칙분류 반영
// ============================================================

function updateDbRules(
  targetMonth,
  familyResultMap
) {

  const dbPath =
    path.join(
      OUTPUT_DIR,
      DB_FILE_NAME
    );

  if (!fs.existsSync(dbPath)) {
    return;
  }

  const wb =
    XLSX.readFile(dbPath);

  const rows =
    sheetRows(
      wb,
      '누적상품DB'
    );

  for (const row of rows) {

    if (
      clean(
        row['조사월']
      ) !== targetMonth
    ) {
      continue;
    }

    const key =
      makeFamilyKey(row);

    const result =
      familyResultMap.get(key);

    if (!result) {
      continue;
    }

    row['규칙분류'] =
      result.type;

    row['규칙분류근거'] =
      result.reason;

    row['신규급부목록'] =
      result.newBenefitText;

    row['AI대상여부'] =
      result.aiTarget;

    row['AI대상사유'] =
      result.aiReason;
  }

  replaceSheet(
    wb,
    '누적상품DB',
    rows
  );

  XLSX.writeFile(
    wb,
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


  const currentBenefitMap =
    buildBenefitMap(
      productRows,
      detailRows
    );


  const historicalMap =
    await buildHistoricalBenefitMap(
      targetMonth
    );


  const resultMap =
    new Map();


  let aiCount = 0;


  for (const row of familyRows) {

    const key =
      makeFamilyKey(row);


    const classification =
      classifyFamily(
        row,
        currentBenefitMap.get(key),
        historicalMap.get(key)
      );


    row['규칙분류'] =
      classification.type;


    row['규칙분류근거'] =
      classification.reason;


    /*
     * 정규화된 이름보다는
     * 사람이 확인하기 쉽도록
     * 현재 급부명과 매칭해서 표시
     */
    const normalizedNew =
      new Set(
        classification.newBenefits
      );


    const readableBenefits = [];


    for (const detail of detailRows) {

      const company =
        getCell(
          detail,
          [
            '보험회사명',
            '회사명',
            '보험회사'
          ]
        );

      const product =
        getCell(
          detail,
          ['상품명']
        );

      const productFamilyKey =
        buildProductFamilyMap(
          productRows
        ).get(
          makeProductKey(
            company,
            product
          )
        );


      if (
        productFamilyKey !== key
      ) {
        continue;
      }


      const name =
        benefitName(detail);


      if (
        normalizedNew.has(
          norm(name)
        ) &&
        name
      ) {

        readableBenefits.push(
          name
        );
      }
    }


    const newBenefitText =
      [
        ...new Set(
          readableBenefits
        )
      ].join(' | ');


    row['신규급부목록'] =
      newBenefitText;


    const targetCompany =
      isAiTargetCompany(
        row['보험회사명']
      );


    const targetType =
      [
        '신규상품후보',
        '신규급부·특약후보'
      ].includes(
        classification.type
      );


    const aiTarget =
      targetCompany &&
      targetType;


    row['AI대상여부'] =
      aiTarget
        ? 'Y'
        : 'N';


    row['AI대상사유'] =
      aiTarget
        ? (
            canonicalCompany(
              row['보험회사명']
            ) +
            ' / ' +
            classification.type
          )
        : '';


    if (aiTarget) {
      aiCount++;
    }


    resultMap.set(
      key,
      {
        type:
          classification.type,

        reason:
          classification.reason,

        newBenefitText,

        aiTarget:
          aiTarget
            ? 'Y'
            : 'N',

        aiReason:
          row['AI대상사유']
      }
    );
  }


  replaceSheet(
    workbook,
    '상품군목록',
    familyRows
  );


  /*
   * 상품목록에도 동일 결과 반영
   */
  for (const row of productRows) {

    const result =
      resultMap.get(
        makeFamilyKey(row)
      );

    if (!result) {
      continue;
    }

    row['규칙분류'] =
      result.type;

    row['규칙분류근거'] =
      result.reason;

    row['신규급부목록'] =
      result.newBenefitText;

    row['AI대상여부'] =
      result.aiTarget;

    row['AI대상사유'] =
      result.aiReason;
  }


  replaceSheet(
    workbook,
    '상품목록',
    productRows
  );


  const summaryMap =
    new Map();


  for (const row of familyRows) {

    const key =
      row['규칙분류'];

    summaryMap.set(
      key,
      (
        summaryMap.get(key) || 0
      ) + 1
    );
  }


  const summaryRows = [
    {
      구분:
        '전체 상품군',
      건수:
        familyRows.length
    },

    {
      구분:
        'AI 대상 상품군',
      건수:
        aiCount
    },

    ...[
      ...summaryMap.entries()
    ].map(
      ([type, count]) => ({
        구분: type,
        건수: count
      })
    )
  ];


  replaceSheet(
    workbook,
    '규칙분류요약',
    summaryRows
  );


  XLSX.writeFile(
    workbook,
    monthlyPath
  );


  updateDbRules(
    targetMonth,
    resultMap
  );


  await uploadFile(
    monthlyPath,
    path.basename(
      monthlyPath
    )
  );


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


  console.log('');
  console.log(
    `전체 상품군: ${familyRows.length}개`
  );

  console.log(
    `AI 심층분석 대상: ${aiCount}개`
  );

  console.log(
    '규칙분류 완료'
  );
}


main()
  .catch(
    error => {

      console.error(error);
      process.exit(1);
    }
  );
