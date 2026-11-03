const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const OUTPUT_DIR =
  path.join(
    __dirname,
    'output'
  );

const DB_FILE_NAME =
  '생보협회_상품DB.xlsx';

const GEMINI_API_KEY =
  process.env.GEMINI_API_KEY;

const GEMINI_MODEL =
  process.env.GEMINI_MODEL ||
  'gemini-2.5-flash-lite';

const UPLOAD_URL =
  process.env.UPLOAD_URL;

const UPLOAD_SECRET =
  process.env.UPLOAD_SECRET;

const TARGET_MONTH_OVERRIDE =
  String(
    process.env.TARGET_MONTH || ''
  ).trim();


const AI_TARGET_COMPANIES = [
  '삼성생명',
  '한화생명',
  '교보생명',
  '신한라이프',
  'NH농협생명'
];


// ============================================================
// 기본
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

  const parts =
    new Intl.DateTimeFormat(
      'en-US',
      {
        timeZone:
          'Asia/Seoul',

        year:
          'numeric',

        month:
          '2-digit'
      }
    )
      .formatToParts(
        new Date()
      )
      .reduce(
        (acc, p) => {

          acc[p.type] =
            p.value;

          return acc;
        },
        {}
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


function isTargetCompany(value) {

  return AI_TARGET_COMPANIES
    .includes(
      canonicalCompany(value)
    );
}


function makeFamilyKey(row) {

  return [
    norm(
      row['보험구분']
    ),

    norm(
      row['상품유형']
    ),

    norm(
      canonicalCompany(
        row['보험회사명']
      )
    ),

    norm(
      row['상품군명']
    )
  ].join('|');
}


function sheetRows(
  workbook,
  name
) {

  const sheet =
    workbook.Sheets[name];

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


function replaceSheet(
  workbook,
  sheetName,
  rows
) {

  const index =
    workbook.SheetNames
      .indexOf(
        sheetName
      );

  if (index >= 0) {

    workbook.SheetNames
      .splice(
        index,
        1
      );

    delete workbook
      .Sheets[
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
              r: rows.length,
              c:
                headers.length -
                1
            }
          })
    };
  }


  XLSX.utils
    .book_append_sheet(
      workbook,
      sheet,
      sheetName
    );
}


// ============================================================
// Drive
// ============================================================

async function apiCall(payload) {

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
      JSON.parse(text);

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


async function uploadFile(
  filePath,
  fileName
) {

  const base64 =
    fs.readFileSync(
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
// Gemini 결과 추출
// ============================================================

function extractOutputText(
  data
) {

  if (
    typeof data.output_text ===
      'string' &&
    data.output_text.trim()
  ) {

    return data.output_text
      .trim();
  }


  if (
    Array.isArray(
      data.steps
    )
  ) {

    const texts = [];


    for (
      const step of data.steps
    ) {

      if (
        step.type !==
        'model_output'
      ) {
        continue;
      }


      for (
        const part
        of step.content || []
      ) {

        if (
          part.type ===
            'text' &&
          part.text
        ) {

          texts.push(
            part.text
          );
        }
      }
    }


    return texts.join(
      '\n'
    );
  }


  return '';
}


function extractSourceUrls(
  data
) {

  const urls =
    new Set();


  if (
    !Array.isArray(
      data.steps
    )
  ) {

    return [];
  }


  for (
    const step of data.steps
  ) {

    for (
      const part
      of step.content || []
    ) {

      for (
        const ann
        of part.annotations || []
      ) {

        if (
          ann.type ===
            'url_citation' &&
          ann.url
        ) {

          urls.add(
            ann.url
          );
        }
      }
    }
  }


  return [...urls];
}


function parseJsonText(text) {

  let s =
    String(text || '')
      .trim();


  s =
    s.replace(
      /^```json\s*/i,
      ''
    );


  s =
    s.replace(
      /^```\s*/,
      ''
    );


  s =
    s.replace(
      /\s*```$/,
      ''
    );


  const first =
    s.indexOf('{');

  const last =
    s.lastIndexOf('}');


  if (
    first < 0 ||
    last < first
  ) {

    throw new Error(
      'JSON 응답을 찾을 수 없습니다.'
    );
  }


  return JSON.parse(
    s.slice(
      first,
      last + 1
    )
  );
}


// ============================================================
// Gemini 호출
// ============================================================

async function callGemini(
  row
) {

  const prompt = `
당신은 한국 생명보험 상품동향 조사 담당자입니다.

아래 상품은 생명보험협회 공시자료를 먼저 분석한 결과,
신규상품 또는 신규급부·특약 후보로 선별된 상품입니다.

[조사대상]
조사월: ${row['조사월']}
보험회사: ${canonicalCompany(row['보험회사명'])}
상품유형: ${row['상품유형']}
상품군명: ${row['상품군명']}
세부상품명: ${row['세부상품명'] || ''}
규칙분류: ${row['규칙분류']}
신규급부 후보: ${row['신규급부목록'] || ''}

웹 검색으로 아래 자료를 우선 확인하세요.

1. 해당 보험회사 공식 홈페이지의 보도자료·뉴스룸·상품자료실
2. 보험신보
3. 보험저널
4. 생명보험협회 등 공식 보험 관련 자료

주의사항:
- 판매일자는 생명보험협회 공시에서 이미 확인했으므로 판매일자를 추측하지 마세요.
- 기사 제목만 보고 판단하지 마세요.
- 실제 신규 출시인지, 기존 상품 개정인지, 신규 급부·특약 추가인지 근거를 확인하세요.
- 자료가 없으면 억지로 판단하지 마세요.
- 광고성 표현은 제거하세요.
- 주요특징은 2문장 이내로 작성하세요.

반드시 아래 JSON 하나만 출력하세요.

{
  "AI검증결과": "신규상품|신규급부·특약|상품개정|자료확인불가",
  "AI신뢰도": "높음|보통|낮음",
  "주요특징": "",
  "확인근거": ""
}
`;


  const response =
    await fetch(
      'https://generativelanguage.googleapis.com/v1beta/interactions',
      {
        method:
          'POST',

        headers: {
          'x-goog-api-key':
            GEMINI_API_KEY,

          'Content-Type':
            'application/json',

          'Api-Revision':
            '2026-05-20'
        },

        body:
          JSON.stringify({
            model:
              GEMINI_MODEL,

            input:
              prompt,

            tools: [
              {
                type:
                  'google_search'
              }
            ]
          })
      }
    );


  const bodyText =
    await response.text();


  if (!response.ok) {

    const error =
      new Error(
        `Gemini HTTP ${response.status}: ${bodyText}`
      );

    error.status =
      response.status;

    throw error;
  }


  const data =
    JSON.parse(
      bodyText
    );


  const outputText =
    extractOutputText(
      data
    );


  const parsed =
    parseJsonText(
      outputText
    );


  const urls =
    extractSourceUrls(
      data
    );


  return {
    ...parsed,

    출처URL:
      urls.join(
        '\n'
      )
  };
}


// ============================================================
// DB 업데이트
// ============================================================

function updateDbAi(
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


  const wb =
    XLSX.readFile(
      dbPath
    );


  const rows =
    sheetRows(
      wb,
      '누적상품DB'
    );


  for (const row of rows) {

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
        makeFamilyKey(row)
      );


    if (!result) {
      continue;
    }


    Object.assign(
      row,
      result
    );
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

  if (!GEMINI_API_KEY) {
    throw new Error(
      'GEMINI_API_KEY가 없습니다.'
    );
  }


  const targetMonth =
    getTargetMonth();


  const monthlyPath =
    path.join(
      OUTPUT_DIR,
      `${targetMonth}_생보협회_상품통합.xlsx`
    );


  const workbook =
    XLSX.readFile(
      monthlyPath
    );


  const familyRows =
    sheetRows(
      workbook,
      '상품군목록'
    );


  const productRows =
    sheetRows(
      workbook,
      '상품목록'
    );


  /*
   * ★ 핵심 필터
   *
   * 전체 생보사는 그대로 유지.
   * AI만 5개사 + 신규/신규급부 후보로 제한.
   */
  const targets =
    familyRows.filter(
      row =>

        row['AI대상여부'] ===
          'Y' &&

        isTargetCompany(
          row['보험회사명']
        ) &&

        [
          '신규상품후보',
          '신규급부·특약후보'
        ].includes(
          row['규칙분류']
        )
    );


  console.log(
    `전체 상품군: ${familyRows.length}개`
  );


  console.log(
    `AI 심층분석 대상: ${targets.length}개`
  );


  targets.forEach(
    row => {

      console.log(
        `[AI대상] ` +
        `${canonicalCompany(row['보험회사명'])} | ` +
        `${row['상품군명']} | ` +
        `${row['규칙분류']}`
      );
    }
  );


  const resultMap =
    new Map();


  for (const row of targets) {

    const key =
      makeFamilyKey(row);


    console.log(
      `AI 검색: ` +
      `${canonicalCompany(row['보험회사명'])} / ` +
      `${row['상품군명']}`
    );


    try {

      const ai =
        await callGemini(
          row
        );


      const result = {

        AI보완상태:
          '완료',

        AI검증결과:
          ai.AI검증결과 ||
          '',

        AI신뢰도:
          ai.AI신뢰도 ||
          '',

        AI주요특징:
          ai.주요특징 ||
          '',

        AI확인근거:
          ai.확인근거 ||
          '',

        AI출처URL:
          ai.출처URL ||
          ''
      };


      Object.assign(
        row,
        result
      );


      resultMap.set(
        key,
        result
      );


    } catch (error) {


      /*
       * 429는 '확인필요'가 아님.
       * 기술적인 API 한도 문제로 별도 표시.
       */
      if (
        error.status ===
        429
      ) {

        const result = {

          AI보완상태:
            '쿼터초과',

          AI검증결과:
            '',

          AI신뢰도:
            '',

          AI주요특징:
            '',

          AI확인근거:
            'Gemini API 사용량 한도 초과',

          AI출처URL:
            ''
        };


        Object.assign(
          row,
          result
        );


        resultMap.set(
          key,
          result
        );


        console.log(
          `429 쿼터초과 → 규칙분류 유지: ${row['상품군명']}`
        );


      } else {


        const result = {

          AI보완상태:
            '오류',

          AI검증결과:
            '',

          AI신뢰도:
            '',

          AI주요특징:
            '',

          AI확인근거:
            String(
              error.message ||
              error
            ).slice(
              0,
              500
            ),

          AI출처URL:
            ''
        };


        Object.assign(
          row,
          result
        );


        resultMap.set(
          key,
          result
        );
      }
    }


    /*
     * 요청을 연속 폭주시키지 않도록
     * 약간 간격을 둔다.
     */
    await new Promise(
      resolve =>
        setTimeout(
          resolve,
          5000
        )
    );
  }


  /*
   * AI 비대상 상품에는 명확히 표시
   */
  for (const row of familyRows) {

    if (
      row['AI대상여부'] !==
      'Y'
    ) {

      row['AI보완상태'] =
        '비대상';
    }
  }


  /*
   * 상품목록에도 상품군 결과 연결
   */
  for (const row of productRows) {

    const family =
      familyRows.find(
        f =>
          makeFamilyKey(f) ===
          makeFamilyKey(row)
      );


    if (!family) {
      continue;
    }


    row['AI보완상태'] =
      family['AI보완상태'] ||
      '';

    row['AI검증결과'] =
      family['AI검증결과'] ||
      '';

    row['AI신뢰도'] =
      family['AI신뢰도'] ||
      '';

    row['AI주요특징'] =
      family['AI주요특징'] ||
      '';

    row['AI확인근거'] =
      family['AI확인근거'] ||
      '';

    row['AI출처URL'] =
      family['AI출처URL'] ||
      '';
  }


  replaceSheet(
    workbook,
    '상품군목록',
    familyRows
  );


  replaceSheet(
    workbook,
    '상품목록',
    productRows
  );


  const completed =
    familyRows.filter(
      row =>
        row['AI보완상태'] ===
        '완료'
    ).length;


  const quota =
    familyRows.filter(
      row =>
        row['AI보완상태'] ===
        '쿼터초과'
    ).length;


  const errors =
    familyRows.filter(
      row =>
        row['AI보완상태'] ===
        '오류'
    ).length;


  const summary = [
    {
      구분:
        '전체 상품군',
      건수:
        familyRows.length
    },

    {
      구분:
        'AI 대상',
      건수:
        targets.length
    },

    {
      구분:
        'AI 보완 완료',
      건수:
        completed
    },

    {
      구분:
        'AI 쿼터초과',
      건수:
        quota
    },

    {
      구분:
        'AI 오류',
      건수:
        errors
    },

    {
      구분:
        'AI 비대상',
      건수:
        familyRows.length -
        targets.length
    },

    {
      구분:
        '고정 AI 대상회사',
      건수:
        5
    }
  ];


  replaceSheet(
    workbook,
    'AI분류요약',
    summary
  );


  XLSX.writeFile(
    workbook,
    monthlyPath
  );


  updateDbAi(
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
    `AI 대상: ${targets.length}개`
  );

  console.log(
    `AI 완료: ${completed}개`
  );

  console.log(
    `쿼터초과: ${quota}개`
  );

  console.log(
    'AI 보완 완료'
  );
}


main()
  .catch(
    error => {

      console.error(
        error
      );

      process.exit(1);
    }
  );
