const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

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
          part => [
            part.type,
            part.value
          ]
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
    `${String(month).padStart(2, '0')}`
  );
}


/*
 * 비교용 문자열 정리
 *
 * 공식 상품명 자체는 절대 수정하지 않는다.
 */
function norm(value) {

  return String(value || '')
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(
      /[\(\)\[\]\{\}·ㆍ\-_,.\/]/g,
      ''
    )
    .trim();
}


/*
 * 상품군 비교용 이름
 *
 * 공식 상품명은 별도 열에 그대로 유지.
 * 여기서는 월별 버전번호 등만 보수적으로 제거.
 */
function makeFamilyName(productName) {

  let name =
    String(productName || '')
      .replace(/\s+/g, ' ')
      .trim();


  /*
   * (무), 무배당은 비교용 이름에서만 제거
   */
  name =
    name
      .replace(
        /\(\s*무\s*\)/gi,
        ''
      )
      .replace(
        /무배당/gi,
        ''
      )
      .replace(
        /\s+/g,
        ' '
      )
      .trim();


  /*
   * 상품명 끝의 2609, 2610 같은
   * 월 버전 표기만 제거
   */
  name =
    name.replace(
      /\s*\((?:25|26|27|28|29)\d{2}\)\s*$/i,
      ''
    );


  name =
    name.replace(
      /\s+(?:25|26|27|28|29)\d{2}\s*$/i,
      ''
    );


  /*
   * V2 / Ver2 같은 끝 표기
   */
  name =
    name.replace(
      /\s+(?:v|ver\.?|version)\s*\d+\s*$/i,
      ''
    );


  return name
    .replace(/\s+/g, ' ')
    .trim();
}


function makeFamilyKey(row) {

  return [
    norm(row['보험구분']),
    norm(row['상품유형']),
    norm(row['보험회사명']),
    norm(
      row['상품군명'] ||
      makeFamilyName(
        row['상품명']
      )
    )
  ].join('|');
}


function makeRepresentativeKey(row) {

  const company =
    norm(row['보험회사명']);

  const code =
    norm(row['상품코드']);

  /*
   * 상품코드가 있으면 우선 활용
   */
  if (code) {
    return (
      company +
      '|CODE|' +
      code
    );
  }


  return (
    company +
    '|NAME|' +
    norm(row['상품명'])
  );
}


function maxMonth(rows) {

  const months =
    rows
      .map(
        row =>
          String(
            row['조사월'] || ''
          )
      )
      .filter(
        value =>
          /^\d{4}-\d{2}$/.test(value)
      )
      .sort();


  return months.length
    ? months[
        months.length - 1
      ]
    : '';
}


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
    result =
      JSON.parse(text);

  } catch {

    throw new Error(
      'Apps Script 응답을 JSON으로 읽지 못했습니다: ' +
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


async function downloadPreviousDb() {

  const result =
    await apiCall({
      action: 'download',
      folderType: 'merged',
      fileName: DB_FILE_NAME
    });


  if (!result.exists) {

    console.log(
      '기존 누적 상품DB 없음 → 이번 달을 기준월로 생성'
    );

    return [];
  }


  const tempPath =
    path.join(
      __dirname,
      'previous_product_db.xlsx'
    );


  fs.writeFileSync(
    tempPath,
    Buffer.from(
      result.base64,
      'base64'
    )
  );


  const workbook =
    XLSX.readFile(
      tempPath,
      {
        cellDates: false
      }
    );


  const sheet =
    workbook.Sheets[
      '누적상품DB'
    ];


  if (!sheet) {

    console.log(
      '기존 DB에 누적상품DB 시트가 없음 → 기준월로 처리'
    );

    return [];
  }


  return XLSX.utils.sheet_to_json(
    sheet,
    {
      defval: ''
    }
  );
}


function readMonthlyProducts(
  targetMonth
) {

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
      `월별 통합파일이 없습니다: ${monthlyPath}`
    );
  }


  const workbook =
    XLSX.readFile(
      monthlyPath,
      {
        cellDates: false
      }
    );


  const sheet =
    workbook.Sheets[
      '상품목록'
    ];


  if (!sheet) {

    throw new Error(
      '월별 통합파일에 상품목록 시트가 없습니다.'
    );
  }


  const rows =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        defval: ''
      }
    );


  return {
    monthlyPath,
    workbook,
    rows
  };
}


function annotateProducts(
  currentRows,
  previousRows,
  targetMonth
) {

  /*
   * 같은 월을 재실행하는 경우
   * 기존 DB의 해당월 데이터는 비교대상에서 제외
   */
  const priorRows =
    previousRows.filter(
      row =>
        String(
          row['조사월'] || ''
        ) !== targetMonth
    );


  const comparisonMonth =
    maxMonth(priorRows);


  const priorExact =
    new Map();

  const priorFamily =
    new Map();

  const priorCode =
    new Map();


  for (
    const row of priorRows
  ) {

    const representativeKey =
      row['대표상품키'] ||
      makeRepresentativeKey(row);


    const familyKey =
      row['상품군키'] ||
      makeFamilyKey(row);


    if (
      representativeKey
    ) {

      if (
        !priorExact.has(
          representativeKey
        )
      ) {

        priorExact.set(
          representativeKey,
          []
        );
      }

      priorExact
        .get(representativeKey)
        .push(row);
    }


    if (
      familyKey
    ) {

      if (
        !priorFamily.has(
          familyKey
        )
      ) {

        priorFamily.set(
          familyKey,
          []
        );
      }

      priorFamily
        .get(familyKey)
        .push(row);
    }


    const code =
      norm(
        row['상품코드']
      );


    if (code) {

      const codeKey =
        norm(
          row['보험회사명']
        ) +
        '|CODE|' +
        code;


      if (
        !priorCode.has(
          codeKey
        )
      ) {

        priorCode.set(
          codeKey,
          []
        );
      }

      priorCode
        .get(codeKey)
        .push(row);
    }
  }


  const isBaseline =
    priorRows.length === 0;


  return currentRows.map(
    row => {

      const output = {
        ...row
      };


      const familyName =
        makeFamilyName(
          row['상품명']
        );


      output['상품군명'] =
        familyName;


      output['상품군키'] =
        makeFamilyKey({
          ...row,
          상품군명:
            familyName
        });


      output['대표상품키'] =
        makeRepresentativeKey(
          row
        );


      const code =
        norm(
          row['상품코드']
        );


      const codeKey =
        code
          ? (
              norm(
                row['보험회사명']
              ) +
              '|CODE|' +
              code
            )
          : '';


      let status;


      if (isBaseline) {

        status =
          '기준월';

      } else if (
        codeKey &&
        priorCode.has(
          codeKey
        )
      ) {

        status =
          '상품개정후보(동일코드)';

      } else if (
        priorExact.has(
          output[
            '대표상품키'
          ]
        )
      ) {

        status =
          '동일상품 재판매·개정후보';

      } else if (
        priorFamily.has(
          output[
            '상품군키'
          ]
        )
      ) {

        status =
          '상품개정후보';

      } else {

        status =
          '신규상품후보';
      }


      output['비교상태'] =
        status;


      output['비교기준월'] =
        comparisonMonth;


      /*
       * 과거 동일 상품이 있으면
       * 최초 수집월을 가져옴
       */
      const oldExact =
        priorExact.get(
          output[
            '대표상품키'
          ]
        ) || [];


      const oldMonths =
        oldExact
          .map(
            item =>
              String(
                item[
                  '최초확인월'
                ] ||
                item[
                  '조사월'
                ] ||
                ''
              )
          )
          .filter(
            value =>
              /^\d{4}-\d{2}$/.test(
                value
              )
          )
          .sort();


      output['최초확인월'] =
        oldMonths.length
          ? oldMonths[0]
          : targetMonth;


      return output;
    }
  );
}


function buildFamilyRows(
  products
) {

  const map =
    new Map();


  for (
    const row of products
  ) {

    const key =
      row['상품군키'];


    if (!map.has(key)) {

      map.set(
        key,
        {
          조사월:
            row['조사월'],

          보험구분:
            row['보험구분'],

          상품유형:
            row['상품유형'],

          보험회사명:
            row['보험회사명'],

          상품군명:
            row['상품군명'],

          상품군키:
            key,

          판매일자:
            new Set(),

          세부상품명:
            new Set(),

          비교상태:
            new Set()
        }
      );
    }


    const item =
      map.get(key);


    item.판매일자.add(
      row['판매일자']
    );


    item.세부상품명.add(
      row['상품명']
    );


    item.비교상태.add(
      row['비교상태']
    );
  }


  return [
    ...map.values()
  ].map(
    item => {

      let familyStatus =
        '기준월';


      const statuses =
        [
          ...item.비교상태
        ];


      if (
        statuses.some(
          x =>
            x ===
            '신규상품후보'
        )
      ) {

        familyStatus =
          '신규상품후보';

      } else if (
        statuses.some(
          x =>
            String(x)
              .includes(
                '개정후보'
              )
        )
      ) {

        familyStatus =
          '상품개정후보';
      }


      return {

        조사월:
          item.조사월,

        보험구분:
          item.보험구분,

        상품유형:
          item.상품유형,

        보험회사명:
          item.보험회사명,

        상품군명:
          item.상품군명,

        상품군키:
          item.상품군키,

        판매일자:
          [
            ...item.판매일자
          ]
            .filter(Boolean)
            .sort()
            .join(', '),

        세부상품수:
          item.세부상품명.size,

        세부상품명:
          [
            ...item.세부상품명
          ].join(' | '),

        비교상태:
          familyStatus
      };
    }
  );
}


function worksheetFromRows(
  rows,
  headers
) {

  const aoa = [
    headers,

    ...rows.map(
      row =>
        headers.map(
          header =>
            row[header] ?? ''
        )
    )
  ];


  const sheet =
    XLSX.utils.aoa_to_sheet(
      aoa
    );


  if (
    rows.length >= 0 &&
    headers.length > 0
  ) {

    sheet['!autofilter'] = {
      ref:
        XLSX.utils.encode_range({
          s: {
            r: 0,
            c: 0
          },

          e: {
            r: rows.length,
            c: headers.length - 1
          }
        })
    };
  }


  sheet['!cols'] =
    headers.map(
      header => {

        if (
          header ===
          '상품명' ||
          header ===
          '세부상품명'
        ) {
          return {
            wch: 45
          };
        }


        if (
          header ===
          '상품군명'
        ) {
          return {
            wch: 40
          };
        }


        if (
          header ===
          '대표상품키' ||
          header ===
          '상품군키'
        ) {
          return {
            wch: 35
          };
        }


        return {
          wch: 18
        };
      }
    );


  return sheet;
}


function replaceSheet(
  workbook,
  sheetName,
  sheet
) {

  if (
    workbook.Sheets[
      sheetName
    ]
  ) {

    delete workbook.Sheets[
      sheetName
    ];


    workbook.SheetNames =
      workbook.SheetNames.filter(
        name =>
          name !== sheetName
      );
  }


  workbook.SheetNames.push(
    sheetName
  );


  workbook.Sheets[
    sheetName
  ] =
    sheet;
}


function updateMonthlyWorkbook(
  monthly,
  products,
  familyRows
) {

  const productHeaders = [
    '조사월',
    '보험구분',
    '상품유형',
    '보험회사명',
    '상품명',
    '상품군명',
    '판매일자',
    '상품코드',
    '대표상품키',
    '상품군키',
    '비교상태',
    '비교기준월',
    '최초확인월',
    '급부행수',
    '원본파일명'
  ];


  replaceSheet(
    monthly.workbook,
    '상품목록',
    worksheetFromRows(
      products,
      productHeaders
    )
  );


  const familyHeaders = [
    '조사월',
    '보험구분',
    '상품유형',
    '보험회사명',
    '상품군명',
    '상품군키',
    '판매일자',
    '세부상품수',
    '세부상품명',
    '비교상태'
  ];


  replaceSheet(
    monthly.workbook,
    '상품군목록',
    worksheetFromRows(
      familyRows,
      familyHeaders
    )
  );


  XLSX.writeFile(
    monthly.workbook,
    monthly.monthlyPath
  );
}


function buildCumulativeDb(
  previousRows,
  currentProducts,
  targetMonth
) {

  /*
   * 같은 월 재실행 시
   * 해당월 기존 데이터 제거 후 재작성
   */
  const keepRows =
    previousRows.filter(
      row =>
        String(
          row['조사월'] || ''
        ) !== targetMonth
    );


  const allRows = [
    ...keepRows,
    ...currentProducts
  ];


  allRows.sort(
    (a, b) =>

      String(
        a['조사월']
      ).localeCompare(
        String(
          b['조사월']
        )
      ) ||

      String(
        a['보험회사명']
      ).localeCompare(
        String(
          b['보험회사명']
        ),
        'ko'
      ) ||

      String(
        a['상품명']
      ).localeCompare(
        String(
          b['상품명']
        ),
        'ko'
      )
  );


  const headers = [
    '조사월',
    '보험구분',
    '상품유형',
    '보험회사명',
    '상품명',
    '상품군명',
    '판매일자',
    '상품코드',
    '대표상품키',
    '상품군키',
    '비교상태',
    '비교기준월',
    '최초확인월',
    '급부행수',
    '원본파일명'
  ];


  const workbook =
    XLSX.utils.book_new();


  XLSX.utils.book_append_sheet(
    workbook,

    worksheetFromRows(
      allRows,
      headers
    ),

    '누적상품DB'
  );


  /*
   * 전체 DB 기준 상품군 현황
   */
  const familyMap =
    new Map();


  for (
    const row of allRows
  ) {

    const key =
      row['상품군키'] ||
      makeFamilyKey(row);


    if (
      !familyMap.has(
        key
      )
    ) {

      familyMap.set(
        key,
        {
          보험구분:
            row['보험구분'],

          상품유형:
            row['상품유형'],

          보험회사명:
            row['보험회사명'],

          상품군명:
            row['상품군명'],

          상품군키:
            key,

          월:
            [],

          상품명:
            []
        }
      );
    }


    const item =
      familyMap.get(key);


    item.월.push(
      String(
        row['조사월'] || ''
      )
    );


    item.상품명.push(
      String(
        row['상품명'] || ''
      )
    );
  }


  const familyHistory =
    [
      ...familyMap.values()
    ].map(
      item => {

        const months =
          [
            ...new Set(
              item.월.filter(Boolean)
            )
          ].sort();


        const names =
          [
            ...new Set(
              item.상품명.filter(Boolean)
            )
          ];


        return {

          보험구분:
            item.보험구분,

          상품유형:
            item.상품유형,

          보험회사명:
            item.보험회사명,

          상품군명:
            item.상품군명,

          상품군키:
            item.상품군키,

          최초확인월:
            months[0] || '',

          최근확인월:
            months[
              months.length - 1
            ] || '',

          확인월수:
            months.length,

          확인상품명:
            names.join(' | ')
        };
      }
    );


  const familyHeaders = [
    '보험구분',
    '상품유형',
    '보험회사명',
    '상품군명',
    '상품군키',
    '최초확인월',
    '최근확인월',
    '확인월수',
    '확인상품명'
  ];


  XLSX.utils.book_append_sheet(
    workbook,

    worksheetFromRows(
      familyHistory,
      familyHeaders
    ),

    '상품군현황'
  );


  const dbPath =
    path.join(
      OUTPUT_DIR,
      DB_FILE_NAME
    );


  XLSX.writeFile(
    workbook,
    dbPath
  );


  return dbPath;
}


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
    action: 'upload',

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


async function main() {

  const targetMonth =
    getTargetMonth();


  console.log(
    `DB 비교 대상월: ${targetMonth}`
  );


  const monthly =
    readMonthlyProducts(
      targetMonth
    );


  const previousRows =
    await downloadPreviousDb();


  console.log(
    `기존 누적DB: ${previousRows.length}행`
  );


  const products =
    annotateProducts(
      monthly.rows,
      previousRows,
      targetMonth
    );


  const familyRows =
    buildFamilyRows(
      products
    );


  console.log(
    `이번달 상품목록: ${products.length}개`
  );


  console.log(
    `이번달 상품군: ${familyRows.length}개`
  );


  const statusCounts = {};


  for (
    const row of products
  ) {

    const status =
      row['비교상태'];


    statusCounts[
      status
    ] =
      (
        statusCounts[
          status
        ] || 0
      ) + 1;
  }


  console.log(
    '비교상태:',
    statusCounts
  );


  /*
   * 월별 통합파일 업그레이드
   */
  updateMonthlyWorkbook(
    monthly,
    products,
    familyRows
  );


  /*
   * 누적 DB 생성
   */
  const dbPath =
    buildCumulativeDb(
      previousRows,
      products,
      targetMonth
    );


  /*
   * 기존 월별 통합파일을
   * 같은 이름으로 다시 업로드
   * → Drive에서 자동 교체됨
   */
  await uploadFile(
    monthly.monthlyPath,
    path.basename(
      monthly.monthlyPath
    )
  );


  await uploadFile(
    dbPath,
    DB_FILE_NAME
  );


  console.log(
    '상품군 비교 + 누적DB 생성 완료'
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
