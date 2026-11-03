const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');


const DOWNLOAD_DIR =
  path.join(
    __dirname,
    'downloads'
  );


const OUTPUT_DIR =
  path.join(
    __dirname,
    'output'
  );


const UPLOAD_URL =
  process.env.UPLOAD_URL;


const UPLOAD_SECRET =
  process.env.UPLOAD_SECRET;


const TARGET_MONTH_OVERRIDE =
  String(
    process.env.TARGET_MONTH || ''
  ).trim();


/*
 * 실행일 기준 직전월
 *
 * 예:
 * 2026-11 실행
 * → 2026-10
 */
function getTargetMonth() {

  if (
    TARGET_MONTH_OVERRIDE
  ) {

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


  const formatter =
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
    );


  const parts =
    Object.fromEntries(

      formatter
        .formatToParts(
          new Date()
        )
        .map(
          p => [
            p.type,
            p.value
          ]
        )
    );


  let year =
    Number(
      parts.year
    );


  let month =
    Number(
      parts.month
    ) - 1;


  if (
    month === 0
  ) {

    year -= 1;
    month = 12;
  }


  return (
    `${year}-` +
    `${String(month).padStart(2, '0')}`
  );
}


/*
 * 헤더 비교용
 *
 * 판매 일자
 * 판매일자
 * 판매\n일자
 *
 * 모두 판매일자로 인식
 */
function norm(value) {

  return String(
    value ?? ''
  )
    .replace(
      /\u00a0/g,
      ' '
    )
    .replace(
      /[\r\n\t]/g,
      ''
    )
    .replace(
      /\s+/g,
      ''
    )
    .replace(
      /[·ㆍ:：]/g,
      ''
    )
    .trim()
    .toLowerCase();
}


function display(value) {

  if (
    value instanceof Date &&
    !Number.isNaN(
      value.getTime()
    )
  ) {

    const y =
      value.getUTCFullYear();


    const m =
      value.getUTCMonth() + 1;


    const d =
      value.getUTCDate();


    return (
      `${y}-` +
      `${String(m).padStart(2, '0')}-` +
      `${String(d).padStart(2, '0')}`
    );
  }


  return value ?? '';
}


/*
 * Excel 날짜를
 * YYYY-MM-DD로 변환
 */
function parseDate(value) {

  if (
    value instanceof Date &&
    !Number.isNaN(
      value.getTime()
    )
  ) {

    return (
      `${value.getUTCFullYear()}-` +
      `${String(
        value.getUTCMonth() + 1
      ).padStart(2, '0')}-` +
      `${String(
        value.getUTCDate()
      ).padStart(2, '0')}`
    );
  }


  /*
   * Excel 날짜 숫자
   */
  if (
    typeof value ===
      'number' &&
    Number.isFinite(value)
  ) {

    const parsed =
      XLSX.SSF.parse_date_code(
        value
      );


    if (
      parsed &&
      parsed.y &&
      parsed.m &&
      parsed.d
    ) {

      return (
        `${parsed.y}-` +
        `${String(
          parsed.m
        ).padStart(2, '0')}-` +
        `${String(
          parsed.d
        ).padStart(2, '0')}`
      );
    }
  }


  const text =
    String(
      value ?? ''
    )
      .trim()
      .replace(
        /\s+/g,
        ''
      );


  if (!text) {

    return null;
  }


  let match =
    text.match(
      /^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/
    );


  if (!match) {

    match =
      text.match(
        /^(\d{4})(\d{2})(\d{2})$/
      );
  }


  if (!match) {

    return null;
  }


  const year =
    Number(match[1]);


  const month =
    Number(match[2]);


  const day =
    Number(match[3]);


  if (
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31
  ) {

    return null;
  }


  return (
    `${year}-` +
    `${String(month).padStart(2, '0')}-` +
    `${String(day).padStart(2, '0')}`
  );
}


/*
 * 병합셀 내용을 펼친 배열 생성
 *
 * 예:
 * X2:X4 병합 = 판매일자
 *
 * X2, X3, X4 모두
 * 판매일자로 인식
 */
function getExpandedSheet(
  worksheet
) {

  const ref =
    worksheet['!ref'] ||
    'A1:A1';


  const range =
    XLSX.utils.decode_range(
      ref
    );


  const rows =
    XLSX.utils.sheet_to_json(
      worksheet,
      {
        header: 1,
        raw: true,
        defval: ''
      }
    );


  while (
    rows.length <=
    range.e.r
  ) {

    rows.push([]);
  }


  for (
    let r = 0;
    r <= range.e.r;
    r++
  ) {

    while (
      rows[r].length <=
      range.e.c
    ) {

      rows[r].push('');
    }
  }


  const merges =
    worksheet['!merges'] ||
    [];


  for (
    const merge of merges
  ) {

    const value =
      rows[
        merge.s.r
      ]?.[
        merge.s.c
      ] ?? '';


    for (
      let r = merge.s.r;
      r <= merge.e.r;
      r++
    ) {

      for (
        let c = merge.s.c;
        c <= merge.e.c;
        c++
      ) {

        rows[r][c] =
          value;
      }
    }
  }


  return {
    rows,
    range,
    merges
  };
}


/*
 * 상단 30행 전체에서
 * 판매일자를 자동 탐색
 */
function findSaleHeader(
  expanded
) {

  const {
    rows,
    range,
    merges
  } = expanded;


  const maxRow =
    Math.min(
      range.e.r,
      29
    );


  for (
    let r = 0;
    r <= maxRow;
    r++
  ) {

    for (
      let c = 0;
      c <= range.e.c;
      c++
    ) {

      if (
        !norm(
          rows[r][c]
        ).includes(
          '판매일자'
        )
      ) {

        continue;
      }


      /*
       * 판매일자가 병합되어 있으면
       * 실제 병합 시작/종료행을 사용
       */
      const merge =
        merges.find(
          m =>
            r >= m.s.r &&
            r <= m.e.r &&
            c >= m.s.c &&
            c <= m.e.c
        );


      return {

        row: r,

        col: c,

        headerStart:
          merge
            ? merge.s.r
            : Math.max(
                0,
                r - 2
              ),

        headerEnd:
          merge
            ? merge.e.r
            : r
      };
    }
  }


  return null;
}


/*
 * 여러 행으로 된 헤더를
 * 하나의 이름으로 합침
 */
function buildHeaders(
  rows,
  startRow,
  endRow,
  maxCol
) {

  const result = [];
  const used =
    new Map();


  for (
    let c = 0;
    c <= maxCol;
    c++
  ) {

    const parts = [];


    for (
      let r = startRow;
      r <= endRow;
      r++
    ) {

      const text =
        String(
          display(
            rows[r]?.[c]
          ) ?? ''
        )
          .replace(
            /\s+/g,
            ' '
          )
          .trim();


      if (
        text &&
        !parts.includes(
          text
        )
      ) {

        parts.push(
          text
        );
      }
    }


    let base =
      parts
        .join(' > ')
        .trim();


    if (!base) {

      base =
        `열_${XLSX.utils.encode_col(c)}`;
    }


    const count =
      (
        used.get(base) ||
        0
      ) + 1;


    used.set(
      base,
      count
    );


    result.push(

      count === 1
        ? base
        : `${base}__${count}`
    );
  }


  return result;
}


/*
 * 회사명, 상품명 등의
 * 실제 열 위치 자동탐색
 */
function findColumn(
  headers,
  aliases
) {

  const normalized =
    headers.map(
      norm
    );


  /*
   * 정확히 일치 우선
   */
  for (
    const alias of aliases
  ) {

    const target =
      norm(alias);


    const index =
      normalized.findIndex(
        h =>
          h === target
      );


    if (
      index >= 0
    ) {

      return index;
    }
  }


  /*
   * 포함검색
   */
  for (
    const alias of aliases
  ) {

    const target =
      norm(alias);


    const index =
      normalized.findIndex(
        h =>
          h.includes(
            target
          )
      );


    if (
      index >= 0
    ) {

      return index;
    }
  }


  return -1;
}


/*
 * 파일명에서
 *
 * 보장성 / 상해보험
 *
 * 같은 분류 자동인식
 */
function parseFileInfo(
  filePath
) {

  const fileName =
    path.basename(
      filePath
    );


  const match =
    fileName.match(
      /^KLIA_(보장성|저축성)_(.+?)\.(xls|xlsx)$/i
    );


  if (!match) {

    throw new Error(
      `파일명을 분석할 수 없습니다: ${fileName}`
    );
  }


  return {

    filePath,

    fileName,

    group:
      match[1],

    category:
      match[2]
  };
}


/*
 * 원본 XLS 1개 분석
 */
function parseWorkbook(
  fileInfo,
  targetMonth
) {

  const workbook =
    XLSX.readFile(
      fileInfo.filePath,
      {
        cellDates: true,
        raw: true
      }
    );


  /*
   * 기본적으로 첫 번째 시트 사용
   */
  const worksheet =
    workbook.Sheets[
      workbook.SheetNames[0]
    ];


  const expanded =
    getExpandedSheet(
      worksheet
    );


  const sale =
    findSaleHeader(
      expanded
    );


  const issues = [];


  if (!sale) {

    issues.push({

      원본파일명:
        fileInfo.fileName,

      원본행번호:
        '',

      항목:
        '판매일자',

      값:
        '',

      사유:
        '상단 30행에서 판매일자 헤더를 찾지 못함'
    });


    return {

      detailRows: [],

      products: [],

      issues,

      info: {

        원본파일명:
          fileInfo.fileName,

        보험구분:
          fileInfo.group,

        상품유형:
          fileInfo.category,

        판매일자헤더:
          '미확인',

        헤더행범위:
          '',

        해당월상품수:
          0,

        해당월급부행수:
          0
      }
    };
  }


  const headers =
    buildHeaders(
      expanded.rows,
      sale.headerStart,
      sale.headerEnd,
      expanded.range.e.c
    );


  const companyCol =
    findColumn(
      headers,
      [
        '보험회사명',
        '회사명',
        '보험회사'
      ]
    );


  const productCol =
    findColumn(
      headers,
      [
        '상품명'
      ]
    );


  const codeCol =
    findColumn(
      headers,
      [
        '상품코드',
        '대표상품코드',
        '상품번호',
        '대표번호'
      ]
    );


  const benefitCol =
    findColumn(
      headers,
      [
        '급부명칭',
        '급부명',
        '보장명',
        '보험금명'
      ]
    );


  const reasonCol =
    findColumn(
      headers,
      [
        '지급사유',
        '급부지급사유'
      ]
    );


  const amountCol =
    findColumn(
      headers,
      [
        '지급금액',
        '보험금액',
        '급부금액'
      ]
    );


  if (
    companyCol < 0
  ) {

    issues.push({

      원본파일명:
        fileInfo.fileName,

      원본행번호:
        '',

      항목:
        '보험회사명',

      값:
        '',

      사유:
        '헤더 자동탐색 실패'
    });
  }


  if (
    productCol < 0
  ) {

    issues.push({

      원본파일명:
        fileInfo.fileName,

      원본행번호:
        '',

      항목:
        '상품명',

      값:
        '',

      사유:
        '헤더 자동탐색 실패'
    });
  }


  const detailRows = [];

  const productMap =
    new Map();


  /*
   * 빈칸으로 이어지는 경우를 위해
   * 직전 값 기억
   */
  const last = {

    company: '',

    product: '',

    code: '',

    saleDate: ''
  };


  for (
    let rowIndex =
      sale.headerEnd + 1;

    rowIndex <=
      expanded.range.e.r;

    rowIndex++
  ) {

    const rawRow =
      expanded.rows[
        rowIndex
      ] || [];


    const hasData =
      rawRow.some(
        value =>
          String(
            value ?? ''
          ).trim() !== ''
      );


    if (!hasData) {

      continue;
    }


    const parsedSale =
      parseDate(
        rawRow[
          sale.col
        ]
      );


    if (
      parsedSale
    ) {

      last.saleDate =
        parsedSale;
    }


    const saleDate =
      parsedSale ||
      last.saleDate;


    /*
     * 해당 월이 아니면
     * 통합본에 넣지 않음
     */
    if (
      !saleDate ||
      saleDate.slice(
        0,
        7
      ) !== targetMonth
    ) {

      continue;
    }


    const companyRaw =
      companyCol >= 0
        ? String(
            display(
              rawRow[
                companyCol
              ]
            ) ?? ''
          ).trim()
        : '';


    const productRaw =
      productCol >= 0
        ? String(
            display(
              rawRow[
                productCol
              ]
            ) ?? ''
          ).trim()
        : '';


    const codeRaw =
      codeCol >= 0
        ? String(
            display(
              rawRow[
                codeCol
              ]
            ) ?? ''
          ).trim()
        : '';


    if (
      companyRaw
    ) {

      last.company =
        companyRaw;
    }


    if (
      productRaw
    ) {

      last.product =
        productRaw;
    }


    if (
      codeRaw
    ) {

      last.code =
        codeRaw;
    }


    const company =
      companyRaw ||
      last.company;


    const product =
      productRaw ||
      last.product;


    const code =
      codeRaw ||
      last.code;


    if (
      !company ||
      !product
    ) {

      issues.push({

        원본파일명:
          fileInfo.fileName,

        원본행번호:
          rowIndex + 1,

        항목:
          !company
            ? '보험회사명'
            : '상품명',

        값:
          '',

        사유:
          '해당월 판매행이지만 상품 식별정보가 비어 있음'
      });


      continue;
    }


    /*
     * 원본의 모든 열도 보존
     */
    const rawFields = {};


    for (
      let c = 0;
      c < headers.length;
      c++
    ) {

      rawFields[
        headers[c]
      ] =
        c === sale.col
          ? saleDate
          : display(
              rawRow[c]
            );
    }


    const detail = {

      조사월:
        targetMonth,

      보험구분:
        fileInfo.group,

      상품유형:
        fileInfo.category,

      보험회사명:
        company,

      상품명:
        product,

      판매일자:
        saleDate,

      상품코드:
        code,

      급부명칭:
        benefitCol >= 0
          ? display(
              rawRow[
                benefitCol
              ]
            )
          : '',

      지급사유:
        reasonCol >= 0
          ? display(
              rawRow[
                reasonCol
              ]
            )
          : '',

      지급금액:
        amountCol >= 0
          ? display(
              rawRow[
                amountCol
              ]
            )
          : '',

      원본파일명:
        fileInfo.fileName,

      원본행번호:
        rowIndex + 1,

      _원본필드:
        rawFields
    };


    detailRows.push(
      detail
    );


    /*
     * 상품목록용
     *
     * 같은 상품의 여러 급부행을
     * 상품 1개로 묶음
     */
    const key = [

      fileInfo.group,

      fileInfo.category,

      norm(company),

      norm(product),

      saleDate,

      norm(code)

    ].join('|');


    if (
      !productMap.has(
        key
      )
    ) {

      productMap.set(
        key,
        {

          조사월:
            targetMonth,

          보험구분:
            fileInfo.group,

          상품유형:
            fileInfo.category,

          보험회사명:
            company,

          상품명:
            product,

          판매일자:
            saleDate,

          상품코드:
            code,

          급부행수:
            0,

          원본파일명:
            fileInfo.fileName
        }
      );
    }


    productMap
      .get(key)
      .급부행수 += 1;
  }


  return {

    detailRows,

    products:
      [
        ...productMap.values()
      ],

    issues,

    info: {

      원본파일명:
        fileInfo.fileName,

      보험구분:
        fileInfo.group,

      상품유형:
        fileInfo.category,

      판매일자헤더:
        `${XLSX.utils.encode_col(
          sale.col
        )}${sale.row + 1}`,

      헤더행범위:
        `${sale.headerStart + 1}~${sale.headerEnd + 1}`,

      해당월상품수:
        productMap.size,

      해당월급부행수:
        detailRows.length
    }
  };
}


function makeAOA(
  rows,
  headers
) {

  return [

    headers,

    ...rows.map(
      row =>
        headers.map(
          header =>
            row[
              header
            ] ?? ''
        )
    )
  ];
}


/*
 * Excel 필터 + 열너비
 */
function setupSheet(
  worksheet,
  rowCount,
  columnCount
) {

  if (
    rowCount > 0 &&
    columnCount > 0
  ) {

    worksheet[
      '!autofilter'
    ] = {

      ref:
        XLSX.utils.encode_range({
          s: {
            r: 0,
            c: 0
          },

          e: {
            r:
              Math.max(
                0,
                rowCount - 1
              ),

            c:
              columnCount - 1
          }
        })
    };
  }


  worksheet[
    '!cols'
  ] =
    Array.from(
      {
        length:
          columnCount
      },

      (_, index) => ({

        wch:
          index === 4
            ? 36
            : index === 3
              ? 22
              : index === 5
                ? 13
                : 18
      })
    );
}


/*
 * 통합 XLSX 생성
 */
function buildMergedWorkbook(
  parsedList,
  targetMonth
) {

  const products =
    parsedList.flatMap(
      item =>
        item.products
    );


  const details =
    parsedList.flatMap(
      item =>
        item.detailRows
    );


  const issues =
    parsedList.flatMap(
      item =>
        item.issues
    );


  const infos =
    parsedList.map(
      item =>
        item.info
    );


  /*
   * 보기 좋은 순서
   */
  products.sort(
    (a, b) =>

      a.보험구분.localeCompare(
        b.보험구분,
        'ko'
      ) ||

      a.상품유형.localeCompare(
        b.상품유형,
        'ko'
      ) ||

      a.판매일자.localeCompare(
        b.판매일자
      ) ||

      a.보험회사명.localeCompare(
        b.보험회사명,
        'ko'
      ) ||

      a.상품명.localeCompare(
        b.상품명,
        'ko'
      )
  );


  details.sort(
    (a, b) =>

      a.보험구분.localeCompare(
        b.보험구분,
        'ko'
      ) ||

      a.상품유형.localeCompare(
        b.상품유형,
        'ko'
      ) ||

      a.판매일자.localeCompare(
        b.판매일자
      ) ||

      a.보험회사명.localeCompare(
        b.보험회사명,
        'ko'
      ) ||

      a.상품명.localeCompare(
        b.상품명,
        'ko'
      ) ||

      (
        a.원본행번호 -
        b.원본행번호
      )
  );


  const productHeaders = [

    '조사월',

    '보험구분',

    '상품유형',

    '보험회사명',

    '상품명',

    '판매일자',

    '상품코드',

    '급부행수',

    '원본파일명'
  ];


  const detailFixed = [

    '조사월',

    '보험구분',

    '상품유형',

    '보험회사명',

    '상품명',

    '판매일자',

    '상품코드',

    '급부명칭',

    '지급사유',

    '지급금액',

    '원본파일명',

    '원본행번호'
  ];


  /*
   * 원본 엑셀의 다른 세부열도
   * 모두 급부상세에 붙임
   */
  const rawHeaderSet =
    new Set();


  for (
    const detail of details
  ) {

    Object
      .keys(
        detail._원본필드 ||
        {}
      )
      .forEach(
        header =>
          rawHeaderSet.add(
            header
          )
      );
  }


  const rawHeaders =
    [
      ...rawHeaderSet
    ];


  const detailHeaders = [

    ...detailFixed,

    ...rawHeaders.map(
      header =>
        `원본>${header}`
    )
  ];


  const detailExport =
    details.map(
      detail => {

        const output = {};


        for (
          const header of
          detailFixed
        ) {

          output[
            header
          ] =
            detail[
              header
            ] ?? '';
        }


        for (
          const header of
          rawHeaders
        ) {

          output[
            `원본>${header}`
          ] =
            detail
              ._원본필드?.[
                header
              ] ?? '';
        }


        return output;
      }
    );


  const infoHeaders = [

    '원본파일명',

    '보험구분',

    '상품유형',

    '판매일자헤더',

    '헤더행범위',

    '해당월상품수',

    '해당월급부행수'
  ];


  const issueHeaders = [

    '원본파일명',

    '원본행번호',

    '항목',

    '값',

    '사유'
  ];


  const workbook =
    XLSX.utils.book_new();


  /*
   * 1. 상품목록
   * 상품 1개 = 1행
   */
  const productSheet =
    XLSX.utils.aoa_to_sheet(
      makeAOA(
        products,
        productHeaders
      )
    );


  setupSheet(
    productSheet,
    products.length + 1,
    productHeaders.length
  );


  XLSX.utils.book_append_sheet(
    workbook,
    productSheet,
    '상품목록'
  );


  /*
   * 2. 급부상세
   */
  const detailSheet =
    XLSX.utils.aoa_to_sheet(
      makeAOA(
        detailExport,
        detailHeaders
      )
    );


  setupSheet(
    detailSheet,
    detailExport.length + 1,
    detailHeaders.length
  );


  XLSX.utils.book_append_sheet(
    workbook,
    detailSheet,
    '급부상세'
  );


  /*
   * 3. 수집정보
   */
  const infoSheet =
    XLSX.utils.aoa_to_sheet(
      makeAOA(
        infos,
        infoHeaders
      )
    );


  setupSheet(
    infoSheet,
    infos.length + 1,
    infoHeaders.length
  );


  XLSX.utils.book_append_sheet(
    workbook,
    infoSheet,
    '수집정보'
  );


  /*
   * 4. 확인필요
   */
  const issueSheet =
    XLSX.utils.aoa_to_sheet(
      makeAOA(
        issues,
        issueHeaders
      )
    );


  setupSheet(
    issueSheet,
    issues.length + 1,
    issueHeaders.length
  );


  XLSX.utils.book_append_sheet(
    workbook,
    issueSheet,
    '확인필요'
  );


  if (
    !fs.existsSync(
      OUTPUT_DIR
    )
  ) {

    fs.mkdirSync(
      OUTPUT_DIR,
      {
        recursive: true
      }
    );
  }


  const outputPath =
    path.join(
      OUTPUT_DIR,
      `${targetMonth}_생보협회_상품통합.xlsx`
    );


  XLSX.writeFile(
    workbook,
    outputPath
  );


  return {

    outputPath,

    productCount:
      products.length,

    detailCount:
      details.length,

    issueCount:
      issues.length
  };
}


/*
 * Apps Script를 통해
 * Google Drive 업로드
 */
async function uploadFile(
  filePath,
  folderType,
  remoteName
) {

  if (
    !UPLOAD_URL ||
    !UPLOAD_SECRET
  ) {

    throw new Error(
      'GitHub Secrets의 UPLOAD_URL 또는 UPLOAD_SECRET이 없습니다.'
    );
  }


  const base64 =
    fs
      .readFileSync(
        filePath
      )
      .toString(
        'base64'
      );


  const body = {

    secret:
      UPLOAD_SECRET,

    folderType,

    fileName:
      remoteName,

    mimeType:
      remoteName
        .toLowerCase()
        .endsWith(
          '.xlsx'
        )
        ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        : 'application/vnd.ms-excel',

    base64
  };


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
          JSON.stringify(
            body
          )
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
      'Drive 업로드 응답 오류: ' +
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
      `Drive 업로드 실패 ${remoteName}: ` +
      (
        result.error ||
        response.status
      )
    );
  }


  console.log(
    `Drive 저장 완료: ${remoteName}`
  );
}


/*
 * 실행
 */
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


  const filePaths =
    fs
      .readdirSync(
        DOWNLOAD_DIR
      )
      .filter(
        name =>
          /^KLIA_(보장성|저축성)_.+\.(xls|xlsx)$/i.test(
            name
          )
      )
      .map(
        name =>
          path.join(
            DOWNLOAD_DIR,
            name
          )
      );


  /*
   * 반드시 13개여야 진행
   */
  if (
    filePaths.length !==
    13
  ) {

    throw new Error(
      `원본 파일이 13개여야 하는데 ${filePaths.length}개입니다.`
    );
  }


  const fileInfos =
    filePaths.map(
      parseFileInfo
    );


  /*
   * 13개 전부 분석
   */
  const parsed =
    fileInfos.map(
      info =>
        parseWorkbook(
          info,
          targetMonth
        )
    );


  /*
   * 통합본 생성
   */
  const merged =
    buildMergedWorkbook(
      parsed,
      targetMonth
    );


  console.log(
    `상품목록: ${merged.productCount}개 상품`
  );


  console.log(
    `급부상세: ${merged.detailCount}행`
  );


  console.log(
    `확인필요: ${merged.issueCount}건`
  );


  /*
   * 원본 13개 Drive 업로드
   *
   * 월을 파일명 앞에 붙임
   */
  for (
    const info of fileInfos
  ) {

    await uploadFile(

      info.filePath,

      'raw',

      `${targetMonth}_${info.fileName}`
    );
  }


  /*
   * 통합본 업로드
   */
  await uploadFile(

    merged.outputPath,

    'merged',

    path.basename(
      merged.outputPath
    )
  );


  console.log(
    '원본 13개 + 통합본 1개 Drive 저장 완료'
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
