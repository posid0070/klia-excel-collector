const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');


// ============================================================
// 기본 경로 / 환경변수
// ============================================================

const DOWNLOAD_DIR =
  path.join(__dirname, 'downloads');

const OUTPUT_DIR =
  path.join(__dirname, 'output');

const UPLOAD_URL =
  process.env.UPLOAD_URL;

const UPLOAD_SECRET =
  process.env.UPLOAD_SECRET;

const TARGET_MONTH_OVERRIDE =
  String(
    process.env.TARGET_MONTH || ''
  ).trim();


if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(
    OUTPUT_DIR,
    { recursive: true }
  );
}


// ============================================================
// 조사월
// ============================================================

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


  /*
   * 자동실행 시
   * 한국시간 기준 직전월
   */
  const parts =
    new Intl.DateTimeFormat(
      'en-US',
      {
        timeZone: 'Asia/Seoul',
        year: 'numeric',
        month: '2-digit'
      }
    )
      .formatToParts(new Date())
      .reduce(
        (acc, p) => {
          acc[p.type] = p.value;
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
    String(month).padStart(2, '0')
  );
}


// ============================================================
// 문자열
// ============================================================

function clean(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(value)
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


// ============================================================
// 날짜
// ============================================================

function pad2(n) {
  return String(n).padStart(2, '0');
}


function formatDate(
  year,
  month,
  day
) {

  if (
    !year ||
    !month ||
    !day
  ) {
    return '';
  }

  return (
    `${year}-` +
    `${pad2(month)}-` +
    `${pad2(day)}`
  );
}


function excelSerialToDate(serial) {

  if (
    typeof serial !== 'number' ||
    !Number.isFinite(serial)
  ) {
    return '';
  }

  /*
   * Excel serial date
   */
  const parsed =
    XLSX.SSF.parse_date_code(
      serial
    );

  if (!parsed) {
    return '';
  }

  return formatDate(
    parsed.y,
    parsed.m,
    parsed.d
  );
}


function normalizeDate(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {
    return '';
  }


  /*
   * XLSX Date 객체
   */
  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {

    return formatDate(
      value.getFullYear(),
      value.getMonth() + 1,
      value.getDate()
    );
  }


  /*
   * Excel 숫자 날짜
   */
  if (
    typeof value === 'number'
  ) {

    const serial =
      excelSerialToDate(value);

    if (serial) {
      return serial;
    }
  }


  let s =
    clean(value);


  /*
   * YYYY-MM-DD
   * YYYY.MM.DD
   * YYYY/MM/DD
   */
  let match =
    s.match(
      /^(\d{4})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/
    );

  if (match) {

    return formatDate(
      Number(match[1]),
      Number(match[2]),
      Number(match[3])
    );
  }


  /*
   * YYYYMMDD
   */
  match =
    s.match(
      /^(\d{4})(\d{2})(\d{2})$/
    );

  if (match) {

    return formatDate(
      Number(match[1]),
      Number(match[2]),
      Number(match[3])
    );
  }


  return '';
}


function isTargetMonth(
  dateValue,
  targetMonth
) {

  const date =
    normalizeDate(
      dateValue
    );

  return (
    date &&
    date.startsWith(
      `${targetMonth}-`
    )
  );
}


// ============================================================
// 파일명에서 보험구분 / 상품유형 추출
// ============================================================

const PRODUCT_TYPES = [

  // 보장성
  '종신보험',
  '정기보험',
  '질병보험',
  '암보험',
  'CI보험',
  '상해보험',
  '어린이보험',
  '치아보험',
  '간병/치매보험',
  '기타',

  // 저축성
  '연금',
  '연금보험',
  '저축',
  '저축보험',
  '교육',
  '교육보험'
];


function parseSourceInfo(
  fileName
) {

  let insuranceType = '';

  if (
    fileName.includes(
      '보장성'
    )
  ) {
    insuranceType =
      '보장성';

  } else if (
    fileName.includes(
      '저축성'
    )
  ) {
    insuranceType =
      '저축성';
  }


  let productType = '';


  for (
    const candidate
    of PRODUCT_TYPES
  ) {

    if (
      fileName.includes(
        candidate
      )
    ) {

      productType =
        candidate;

      break;
    }
  }


  /*
   * 표시명 통일
   */
  if (
    productType ===
    '연금보험'
  ) {
    productType = '연금';
  }

  if (
    productType ===
    '저축보험'
  ) {
    productType = '저축';
  }

  if (
    productType ===
    '교육보험'
  ) {
    productType = '교육';
  }


  return {
    보험구분:
      insuranceType,

    상품유형:
      productType
  };
}


// ============================================================
// XLS merged cells 확장
// ============================================================

function sheetToMatrix(sheet) {

  const matrix =
    XLSX.utils.sheet_to_json(
      sheet,
      {
        header: 1,
        defval: '',
        raw: true
      }
    );


  const merges =
    sheet['!merges'] || [];


  for (
    const merge of merges
  ) {

    const sourceRow =
      merge.s.r;

    const sourceCol =
      merge.s.c;


    if (
      !matrix[sourceRow]
    ) {
      continue;
    }


    const sourceValue =
      matrix[sourceRow][
        sourceCol
      ];


    for (
      let r = merge.s.r;
      r <= merge.e.r;
      r++
    ) {

      if (!matrix[r]) {
        matrix[r] = [];
      }


      for (
        let c = merge.s.c;
        c <= merge.e.c;
        c++
      ) {

        if (
          matrix[r][c] === '' ||
          matrix[r][c] ===
            undefined
        ) {

          matrix[r][c] =
            sourceValue;
        }
      }
    }
  }


  return matrix;
}


// ============================================================
// 헤더 탐색
// ============================================================

function normalizedHeader(
  value
) {

  return norm(value);
}


function headerMatches(
  value,
  aliases
) {

  const h =
    normalizedHeader(
      value
    );


  if (!h) {
    return false;
  }


  return aliases.some(
    alias =>
      h ===
      normalizedHeader(alias)
  );
}


function findColumn(
  matrix,
  headerRows,
  aliases
) {

  for (
    const rowIndex
    of headerRows
  ) {

    const row =
      matrix[rowIndex] || [];


    for (
      let col = 0;
      col < row.length;
      col++
    ) {

      if (
        headerMatches(
          row[col],
          aliases
        )
      ) {

        return col;
      }
    }
  }


  return -1;
}


function detectHeader(
  matrix
) {

  /*
   * 판매일자를 가장 중요한 기준으로 탐색
   * 상단 30행 확인
   */
  const limit =
    Math.min(
      30,
      matrix.length
    );


  for (
    let r = 0;
    r < limit;
    r++
  ) {

    const row =
      matrix[r] || [];


    for (
      let c = 0;
      c < row.length;
      c++
    ) {

      if (
        headerMatches(
          row[c],
          [
            '판매일자',
            '판매일',
            '판매개시일자',
            '판매개시일'
          ]
        )
      ) {

        /*
         * 협회 파일은 다중행 헤더가 있으므로
         * 해당 행 기준 앞뒤 3행까지 헤더 후보로 사용
         */
        const headerRows = [];


        for (
          let h =
            Math.max(0, r - 3);
          h <=
            Math.min(
              matrix.length - 1,
              r + 3
            );
          h++
        ) {

          headerRows.push(h);
        }


        return {
          saleHeaderRow:
            r,

          headerRows
        };
      }
    }
  }


  return null;
}


// ============================================================
// 전행 값 보정
// ============================================================

function getValue(
  row,
  col
) {

  if (
    col < 0 ||
    !row
  ) {
    return '';
  }

  return row[col] ?? '';
}


function forwardFill(
  current,
  previous
) {

  const v =
    clean(current);

  return v
    ? current
    : previous;
}


// ============================================================
// 원본 XLS 1개 파싱
// ============================================================

function parseOneFile(
  filePath,
  targetMonth
) {

  const fileName =
    path.basename(
      filePath
    );


  const sourceInfo =
    parseSourceInfo(
      fileName
    );


  const workbook =
    XLSX.readFile(
      filePath,
      {
        cellDates: true,
        raw: true
      }
    );


  const sheetName =
    workbook.SheetNames[0];


  const sheet =
    workbook.Sheets[
      sheetName
    ];


  const matrix =
    sheetToMatrix(
      sheet
    );


  const detected =
    detectHeader(
      matrix
    );


  if (!detected) {

    return {
      sourceInfo,
      fileName,

      status:
        '판매일자 헤더 미확인',

      headerRow:
        '',

      rows: []
    };
  }


  const headerRows =
    detected.headerRows;


  /*
   * 주요 필드
   */
  const companyCol =
    findColumn(
      matrix,
      headerRows,
      [
        '보험회사명',
        '회사명',
        '보험회사',
        '생명보험회사'
      ]
    );


  const productCol =
    findColumn(
      matrix,
      headerRows,
      [
        '상품명',
        '보험상품명'
      ]
    );


  const saleDateCol =
    findColumn(
      matrix,
      headerRows,
      [
        '판매일자',
        '판매일',
        '판매개시일자',
        '판매개시일'
      ]
    );


  /*
   * ★ 중요
   *
   * 상품코드는 실제 "상품코드/상품번호"
   * 헤더가 있을 때만 읽는다.
   *
   * 대표번호는 절대 포함하지 않는다.
   */
  const productCodeCol =
    findColumn(
      matrix,
      headerRows,
      [
        '상품코드',
        '대표상품코드',
        '상품번호',
        '보험상품코드'
      ]
    );


  /*
   * ★ 대표전화번호 별도
   */
  const phoneCol =
    findColumn(
      matrix,
      headerRows,
      [
        '대표번호',
        '대표전화',
        '대표전화번호',
        '전화번호'
      ]
    );


  /*
   * 급부 관련 필드
   */
  const benefitNameCol =
    findColumn(
      matrix,
      headerRows,
      [
        '급부명칭',
        '급부명',
        '급부',
        '보장명',
        '담보명'
      ]
    );


  const benefitReasonCol =
    findColumn(
      matrix,
      headerRows,
      [
        '지급사유',
        '보험금지급사유',
        '급부지급사유',
        '지급조건'
      ]
    );


  const benefitAmountCol =
    findColumn(
      matrix,
      headerRows,
      [
        '지급금액',
        '보험금액',
        '급부금액',
        '보장금액'
      ]
    );


  /*
   * 데이터 시작점
   *
   * 실제 판매일자 헤더 행 이후부터 검색하되,
   * 다중행 헤더가 있으므로 헤더 후보 마지막 행 + 1
   */
  const dataStartRow =
    Math.max(
      ...headerRows
    ) + 1;


  const rows = [];


  let prevCompany = '';
  let prevProduct = '';
  let prevSaleDate = '';
  let prevProductCode = '';
  let prevPhone = '';


  for (
    let r = dataStartRow;
    r < matrix.length;
    r++
  ) {

    const rawRow =
      matrix[r] || [];


    /*
     * 병합/연속행 대응
     */
    let company =
      getValue(
        rawRow,
        companyCol
      );


    let product =
      getValue(
        rawRow,
        productCol
      );


    let saleDate =
      getValue(
        rawRow,
        saleDateCol
      );


    let productCode =
      getValue(
        rawRow,
        productCodeCol
      );


    let phone =
      getValue(
        rawRow,
        phoneCol
      );


    company =
      forwardFill(
        company,
        prevCompany
      );


    product =
      forwardFill(
        product,
        prevProduct
      );


    saleDate =
      forwardFill(
        saleDate,
        prevSaleDate
      );


    productCode =
      forwardFill(
        productCode,
        prevProductCode
      );


    phone =
      forwardFill(
        phone,
        prevPhone
      );


    /*
     * forward fill 값 갱신
     */
    if (clean(company)) {
      prevCompany =
        company;
    }

    if (clean(product)) {
      prevProduct =
        product;
    }

    if (
      normalizeDate(
        saleDate
      )
    ) {
      prevSaleDate =
        saleDate;
    }

    if (
      clean(productCode)
    ) {
      prevProductCode =
        productCode;
    }

    if (
      clean(phone)
    ) {
      prevPhone =
        phone;
    }


    const normalizedSaleDate =
      normalizeDate(
        saleDate
      );


    /*
     * 대상월 판매일자만
     */
    if (
      !normalizedSaleDate ||
      !normalizedSaleDate
        .startsWith(
          `${targetMonth}-`
        )
    ) {
      continue;
    }


    const companyText =
      clean(company);

    const productText =
      clean(product);


    if (
      !companyText ||
      !productText
    ) {
      continue;
    }


    /*
     * 상품코드 방어 로직
     *
     * 혹시 원본 구조가 이상해도
     * 02-xxxx / 15xx-xxxx 같은 전화번호는
     * 상품코드로 저장하지 않는다.
     */
    let productCodeText =
      clean(
        productCode
      );


    if (
      looksLikePhoneNumber(
        productCodeText
      )
    ) {
      productCodeText = '';
    }


    let phoneText =
      clean(phone);


    /*
     * 상품코드 칼럼에 잘못 전화번호가 들어있고
     * 대표번호 칼럼이 없는 특이한 경우에도
     * 전화번호 쪽으로 이동
     */
    if (
      !phoneText &&
      looksLikePhoneNumber(
        clean(productCode)
      )
    ) {

      phoneText =
        clean(productCode);
    }


    rows.push({

      조사월:
        targetMonth,

      보험구분:
        sourceInfo.보험구분,

      상품유형:
        sourceInfo.상품유형,

      보험회사명:
        companyText,

      상품명:
        productText,

      판매일자:
        normalizedSaleDate,

      상품코드:
        productCodeText,

      대표번호:
        phoneText,

      급부명칭:
        clean(
          getValue(
            rawRow,
            benefitNameCol
          )
        ),

      지급사유:
        clean(
          getValue(
            rawRow,
            benefitReasonCol
          )
        ),

      지급금액:
        clean(
          getValue(
            rawRow,
            benefitAmountCol
          )
        ),

      원본파일명:
        fileName,

      원본행번호:
        r + 1
    });
  }


  return {

    sourceInfo,

    fileName,

    status:
      '정상',

    headerRow:
      detected.saleHeaderRow + 1,

    companyCol:
      companyCol >= 0
        ? companyCol + 1
        : '',

    productCol:
      productCol >= 0
        ? productCol + 1
        : '',

    saleDateCol:
      saleDateCol >= 0
        ? saleDateCol + 1
        : '',

    productCodeCol:
      productCodeCol >= 0
        ? productCodeCol + 1
        : '',

    phoneCol:
      phoneCol >= 0
        ? phoneCol + 1
        : '',

    rows
  };
}


// ============================================================
// 전화번호 판별
// ============================================================

function looksLikePhoneNumber(
  value
) {

  const s =
    clean(value)
      .replace(/\s/g, '');


  if (!s) {
    return false;
  }


  /*
   * 1588-xxxx
   * 1544-xxxx
   * 1566-xxxx
   * 1577-xxxx
   * 1600-xxxx
   * 1688-xxxx 등
   */
  if (
    /^(15|16)\d{2}-?\d{4}$/.test(
      s
    )
  ) {
    return true;
  }


  /*
   * 02-xxxx-xxxx
   * 031-xxx-xxxx 등
   */
  if (
    /^0\d{1,2}-?\d{3,4}-?\d{4}$/.test(
      s
    )
  ) {
    return true;
  }


  return false;
}


// ============================================================
// 상품목록 생성
// ============================================================

function buildProductList(
  detailRows
) {

  const map =
    new Map();


  for (
    const row of detailRows
  ) {

    /*
     * 대표번호는 상품 식별키에 사용하지 않는다.
     *
     * 상품코드도 실제 값이 있을 때만 참고.
     */
    const key = [

      norm(
        row['보험구분']
      ),

      norm(
        row['상품유형']
      ),

      norm(
        row['보험회사명']
      ),

      norm(
        row['상품명']
      ),

      row['판매일자'],

      norm(
        row['상품코드']
      )

    ].join('|');


    if (
      !map.has(
        key
      )
    ) {

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

          상품명:
            row['상품명'],

          판매일자:
            row['판매일자'],

          상품코드:
            row['상품코드'],

          대표번호:
            row['대표번호'],

          급부행수:
            0,

          원본파일명:
            new Set()
        }
      );
    }


    const item =
      map.get(key);


    item.급부행수 += 1;


    if (
      row['원본파일명']
    ) {

      item.원본파일명.add(
        row['원본파일명']
      );
    }


    /*
     * 대표번호는 참고정보로만 보존
     */
    if (
      !item.대표번호 &&
      row['대표번호']
    ) {

      item.대표번호 =
        row['대표번호'];
    }
  }


  return [
    ...map.values()
  ]
    .map(
      item => ({

        조사월:
          item.조사월,

        보험구분:
          item.보험구분,

        상품유형:
          item.상품유형,

        보험회사명:
          item.보험회사명,

        상품명:
          item.상품명,

        판매일자:
          item.판매일자,

        상품코드:
          item.상품코드,

        대표번호:
          item.대표번호,

        급부행수:
          item.급부행수,

        원본파일명:
          [
            ...item.원본파일명
          ].join(' | ')
      })
    )
    .sort(
      (a, b) =>

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
}


// ============================================================
// Excel Sheet 생성
// ============================================================

function makeSheet(
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
            c:
              headers.length - 1
          }
        })
    };
  }


  sheet['!cols'] =
    headers.map(
      header => {

        if (
          header ===
          '상품명'
        ) {
          return {
            wch: 48
          };
        }

        if (
          header ===
          '지급사유'
        ) {
          return {
            wch: 55
          };
        }

        if (
          header ===
          '급부명칭'
        ) {
          return {
            wch: 35
          };
        }

        if (
          header ===
          '원본파일명'
        ) {
          return {
            wch: 38
          };
        }

        return {
          wch: 18
        };
      }
    );


  return sheet;
}


// ============================================================
// Drive API
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
      JSON.parse(text);

  } catch {

    throw new Error(
      'Apps Script 응답을 JSON으로 읽을 수 없습니다: ' +
      text.slice(
        0,
        500
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
  folderType,
  fileName,
  mimeType
) {

  const base64 =
    fs
      .readFileSync(
        filePath
      )
      .toString(
        'base64'
      );


  const result =
    await apiCall({

      action:
        'upload',

      folderType,

      fileName,

      mimeType:
        mimeType ||
        'application/octet-stream',

      base64
    });


  console.log(
    `Drive 저장 완료: ${folderType} / ${fileName}`
  );


  return result;
}


// ============================================================
// 실행
// ============================================================

async function main() {

  const targetMonth =
    getTargetMonth();


  console.log(
    `통합 대상 조사월: ${targetMonth}`
  );


  if (
    !fs.existsSync(
      DOWNLOAD_DIR
    )
  ) {

    throw new Error(
      'downloads 폴더가 없습니다.'
    );
  }


  /*
   * XLS / XLSX 대상
   */
  const files =
    fs
      .readdirSync(
        DOWNLOAD_DIR
      )
      .filter(
        file =>
          /\.(xls|xlsx)$/i.test(
            file
          )
      )
      .sort();


  if (
    files.length === 0
  ) {

    throw new Error(
      '다운로드된 Excel 파일이 없습니다.'
    );
  }


  console.log(
    `원본 파일: ${files.length}개`
  );


  const allDetails = [];

  const collectionInfo = [];

  const parseIssues = [];


  for (
    const file
    of files
  ) {

    const filePath =
      path.join(
        DOWNLOAD_DIR,
        file
      );


    console.log(
      `파싱 중: ${file}`
    );


    try {

      const result =
        parseOneFile(
          filePath,
          targetMonth
        );


      collectionInfo.push({

        조사월:
          targetMonth,

        보험구분:
          result.sourceInfo
            .보험구분,

        상품유형:
          result.sourceInfo
            .상품유형,

        원본파일명:
          result.fileName,

        상태:
          result.status,

        판매일자헤더행:
          result.headerRow,

        보험회사열:
          result.companyCol || '',

        상품명열:
          result.productCol || '',

        판매일자열:
          result.saleDateCol || '',

        상품코드열:
          result.productCodeCol || '',

        대표번호열:
          result.phoneCol || '',

        대상월행수:
          result.rows.length
      });


      if (
        result.status !==
        '정상'
      ) {

        parseIssues.push({

          원본파일명:
            file,

          사유:
            result.status
        });

        continue;
      }


      allDetails.push(
        ...result.rows
      );


    } catch (error) {

      console.error(
        `파싱 실패: ${file}`,
        error
      );


      collectionInfo.push({

        조사월:
          targetMonth,

        보험구분:
          '',

        상품유형:
          '',

        원본파일명:
          file,

        상태:
          '파싱 실패',

        판매일자헤더행:
          '',

        보험회사열:
          '',

        상품명열:
          '',

        판매일자열:
          '',

        상품코드열:
          '',

        대표번호열:
          '',

        대상월행수:
          0
      });


      parseIssues.push({

        원본파일명:
          file,

        사유:
          String(
            error.message ||
            error
          )
      });
    }
  }


  /*
   * 전체 급부행 정렬
   */
  allDetails.sort(
    (a, b) =>

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
      ) ||

      String(
        a['급부명칭']
      ).localeCompare(
        String(
          b['급부명칭']
        ),
        'ko'
      )
  );


  const products =
    buildProductList(
      allDetails
    );


  console.log(
    `상품목록: ${products.length}개 상품`
  );


  console.log(
    `급부상세: ${allDetails.length}행`
  );


  console.log(
    `확인필요: ${parseIssues.length}건`
  );


  /*
   * 월별 통합 Workbook
   */
  const workbook =
    XLSX.utils.book_new();


  const productHeaders = [

    '조사월',
    '보험구분',
    '상품유형',

    '보험회사명',
    '상품명',

    '판매일자',

    /*
     * 실제 상품코드가 있을 때만 값 존재
     */
    '상품코드',

    /*
     * ★ 대표전화번호 별도
     */
    '대표번호',

    '급부행수',
    '원본파일명'
  ];


  XLSX.utils.book_append_sheet(

    workbook,

    makeSheet(
      products,
      productHeaders
    ),

    '상품목록'
  );


  const detailHeaders = [

    '조사월',
    '보험구분',
    '상품유형',

    '보험회사명',
    '상품명',

    '판매일자',

    '상품코드',
    '대표번호',

    '급부명칭',
    '지급사유',
    '지급금액',

    '원본파일명',
    '원본행번호'
  ];


  XLSX.utils.book_append_sheet(

    workbook,

    makeSheet(
      allDetails,
      detailHeaders
    ),

    '급부상세'
  );


  const infoHeaders = [

    '조사월',
    '보험구분',
    '상품유형',

    '원본파일명',

    '상태',

    '판매일자헤더행',

    '보험회사열',
    '상품명열',
    '판매일자열',

    '상품코드열',
    '대표번호열',

    '대상월행수'
  ];


  XLSX.utils.book_append_sheet(

    workbook,

    makeSheet(
      collectionInfo,
      infoHeaders
    ),

    '수집정보'
  );


  const issueHeaders = [
    '원본파일명',
    '사유'
  ];


  XLSX.utils.book_append_sheet(

    workbook,

    makeSheet(
      parseIssues,
      issueHeaders
    ),

    '확인필요'
  );


  /*
   * output 저장
   */
  const mergedFileName =
    `${targetMonth}_생보협회_상품통합.xlsx`;


  const mergedFilePath =
    path.join(
      OUTPUT_DIR,
      mergedFileName
    );


  XLSX.writeFile(
    workbook,
    mergedFilePath
  );


  console.log(
    `통합파일 생성: ${mergedFileName}`
  );


  // ==========================================================
  // 원본 파일 Drive 저장
  // ==========================================================

  for (
    const file
    of files
  ) {

    const filePath =
      path.join(
        DOWNLOAD_DIR,
        file
      );


    /*
     * 이미 YYYY-MM_으로 시작하는 파일이면
     * 다시 중복해서 붙이지 않는다.
     */
    const driveFileName =
      /^\d{4}-\d{2}_/.test(
        file
      )
        ? file
        : `${targetMonth}_${file}`;


    const extension =
      path
        .extname(file)
        .toLowerCase();


    const mimeType =
      extension === '.xlsx'
        ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        : 'application/vnd.ms-excel';


    await uploadFile(
      filePath,
      'raw',
      driveFileName,
      mimeType
    );
  }


  // ==========================================================
  // 통합본 Drive 저장
  // ==========================================================

  await uploadFile(

    mergedFilePath,

    'merged',

    mergedFileName,

    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );


  console.log('');
  console.log(
    '월별 상품 통합 및 Drive 업로드 완료'
  );

  console.log(
    `조사월: ${targetMonth}`
  );

  console.log(
    `상품목록: ${products.length}개`
  );

  console.log(
    `급부상세: ${allDetails.length}행`
  );

  console.log(
    `파싱 확인필요: ${parseIssues.length}건`
  );
}


// ============================================================
// Run
// ============================================================

main()
  .catch(
    error => {

      console.error(
        'process_upload.js 실행 실패'
      );

      console.error(
        error
      );

      process.exit(1);
    }
  );
