const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUTPUT_DIR = path.join(ROOT, 'output');

const TARGET_MONTH = String(
  process.env.TARGET_MONTH || ''
).trim();

const UPLOAD_URL =
  process.env.UPLOAD_URL || '';

const UPLOAD_SECRET =
  process.env.UPLOAD_SECRET || '';

if (!/^\d{4}-\d{2}$/.test(TARGET_MONTH)) {
  throw new Error(
    'TARGET_MONTH는 YYYY-MM 형식이어야 합니다.'
  );
}

fs.mkdirSync(
  OUTPUT_DIR,
  { recursive: true }
);

const OUTPUT_FILE = path.join(
  OUTPUT_DIR,
  `${TARGET_MONTH}_생보협회_상품_통합.xlsx`
);

const META_FILE = path.join(
  OUTPUT_DIR,
  `${TARGET_MONTH}_V5_meta.json`
);


/* =========================================================
   최종 통합본 컬럼
========================================================= */

const TARGET_COLUMNS = [
  '보험회사명',
  '상품명',
  '정리상품명',
  '상품형태',
  '급부명칭',
  '지급사유',
  '지급금액',
  '판매일자',
  '특이사항'
];

const CATEGORY_ORDER = [
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
  '연금',
  '저축',
  '교육'
];


/* =========================================================
   원본 엑셀 헤더 후보
========================================================= */

const HEADER_ALIASES = {

  보험회사명: [
    '보험회사명',
    '회사명',
    '보험회사',
    '생명보험회사'
  ],

  상품명: [
    '상품명',
    '보험상품명',
    '상품 명'
  ],

  급부명칭: [
    '급부명칭',
    '급부명',
    '급부 명칭',
    '보장명',
    '담보명'
  ],

  지급사유: [
    '지급사유',
    '보험금지급사유',
    '보험금 지급사유',
    '급부지급사유',
    '지급조건'
  ],

  지급금액: [
    '지급금액',
    '보험금액',
    '급부금액',
    '보장금액',
    '지급 금액'
  ],

  판매일자: [
    '판매일자',
    '판매개시일',
    '판매개시일자',
    '판매일',
    '판매 일자'
  ],

  특이사항: [
    '특이사항',
    '비고',
    '특이 사항'
  ]

};


/* =========================================================
   기본 문자열 정리
========================================================= */

function decodeHtmlEntities(v) {

  let text = String(v ?? '');

  const named = {
    amp: '&',
    lt: '<',
    gt: '>',
    quot: '"',
    apos: "'",
    nbsp: ' ',
    ldquo: '“',
    rdquo: '”',
    lsquo: '‘',
    rsquo: '’',
    middot: '·',
    hellip: '…',
    ndash: '–',
    mdash: '—'
  };

  /*
    원본 XLS 안에 &ldquo; 또는 &#x201C;처럼
    HTML 엔티티가 그대로 들어오는 경우가 있어
    최대 3회까지 반복 복원한다.
  */
  for (let pass = 0; pass < 3; pass++) {

    const before = text;

    text = text
      .replace(/&#x([0-9a-f]+);?/gi, (_, hex) => {
        const code = Number.parseInt(hex, 16);
        return Number.isFinite(code) ? String.fromCodePoint(code) : _;
      })
      .replace(/&#(\d+);?/g, (_, dec) => {
        const code = Number.parseInt(dec, 10);
        return Number.isFinite(code) ? String.fromCodePoint(code) : _;
      })
      .replace(/&([a-z]+);?/gi, (match, name) => {
        const decoded = named[name.toLowerCase()];
        return decoded === undefined ? match : decoded;
      });

    if (text === before) {
      break;
    }

  }

  return text;

}


function clean(v) {

  return decodeHtmlEntities(v)
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

}


function splitProductName(productName) {

  const original = clean(productName);
  const variants = [];

  const normalized = original.replace(
    /(\([^()]*\)|\[[^\[\]]*\])/g,
    group => {

      const open = group[0];
      const close = group[group.length - 1];
      const content = group.slice(1, -1).trim();

      if (!content) {
        return group;
      }

      const parts = content
        .split(/[,，|·ㆍ;/]+/)
        .map(x => clean(x))
        .filter(Boolean);

      const keep = [];

      for (const part of parts) {

        if (/형\s*$/u.test(part)) {
          variants.push(part);
        } else {
          keep.push(part);
        }

      }

      /*
        구분자가 없지만 괄호 전체가 ~형으로 끝나는 경우도 추출
      */
      if (parts.length === 1 && /형\s*$/u.test(content)) {
        if (!variants.includes(content)) {
          variants.push(content);
        }
        return '';
      }

      if (!keep.length) {
        return '';
      }

      return `${open}${keep.join(', ')}${close}`;

    }
  )
  .replace(/\s+/g, ' ')
  .replace(/\s+([)\]])/g, '$1')
  .trim();

  const uniqueVariants = [
    ...new Set(
      variants.map(x => clean(x))
    )
  ];

  return {
    original,
    normalized,
    variants: uniqueVariants
  };

}


function norm(v) {

  return clean(v)
    .toLowerCase()
    .replace(
      /[\s\(\)\[\]\{\}·ㆍ\-_,.\/\\:：'"“”‘’]/g,
      ''
    );

}


/* =========================================================
   상품 중복 판단 기준

   보험회사명 + 상품명

   동일 상품이 보험기간, 급부 등으로 여러 행이어도
   상품 수는 1건
========================================================= */

function uniqueProductKey(row) {

  return (
    `${norm(row.보험회사명)}|` +
    `${norm(row.상품명)}`
  );

}


/* =========================================================
   급부 중복 판단 기준

   회사 + 상품 + 급부명칭
========================================================= */

function productBenefitKey(row) {

  return (
    `${uniqueProductKey(row)}|` +
    `${norm(row.급부명칭)}`
  );

}


/* =========================================================
   파일명에서 상품 카테고리 확인
========================================================= */

function parseCategoryFromFile(filePath) {

  const name =
    path.basename(filePath);

  const categories = [

    '종신보험',
    '정기보험',
    '질병보험',
    '암보험',
    'CI보험',
    '상해보험',

    '어린이보험',
    '치아보험',

    '간병/치매보험',
    '간병치매보험',
    '간병·치매보험',

    '기타',

    '연금',
    '저축',
    '교육'

  ];

  for (const c of categories) {

    if (name.includes(c)) {

      if (
        c === '간병치매보험' ||
        c === '간병·치매보험'
      ) {

        return '간병/치매보험';

      }

      return c;

    }

  }

  return '';

}


/* =========================================================
   저축성 여부

   연금 / 저축 / 교육

   저축성의 경우
   급부명칭
   지급사유
   지급금액

   → 전부 빈칸
========================================================= */

function isSavingCategory(
  category,
  filePath
) {

  const name =
    path.basename(filePath);

  return (

    name.includes('저축성') ||

    [
      '연금',
      '저축',
      '교육'
    ].includes(category)

  );

}


/* =========================================================
   판매일자 YYYY-MM-DD 변환
========================================================= */

function excelDateToIso(value) {

  if (
    value === null ||
    value === undefined ||
    value === ''
  ) {

    return '';

  }


  /* Date 객체 */

  if (
    value instanceof Date &&
    !Number.isNaN(value.getTime())
  ) {

    const y =
      value.getFullYear();

    const m =
      String(
        value.getMonth() + 1
      ).padStart(2, '0');

    const d =
      String(
        value.getDate()
      ).padStart(2, '0');

    return `${y}-${m}-${d}`;

  }


  /* Excel serial date */

  if (
    typeof value === 'number'
  ) {

    const parsed =
      XLSX.SSF.parse_date_code(
        value
      );

    if (parsed) {

      return (
        `${parsed.y}-` +
        `${String(parsed.m).padStart(2, '0')}-` +
        `${String(parsed.d).padStart(2, '0')}`
      );

    }

  }


  const s =
    clean(value);


  /*
    2026-10-01
    2026.10.01
    2026/10/01
    2026년10월01일
  */

  let m =
    s.match(
      /^(20\d{2})[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})/
    );

  if (m) {

    return (
      `${m[1]}-` +
      `${String(Number(m[2])).padStart(2, '0')}-` +
      `${String(Number(m[3])).padStart(2, '0')}`
    );

  }


  /*
    20261001
  */

  m =
    s.match(
      /^(\d{4})(\d{2})(\d{2})$/
    );

  if (m) {

    return (
      `${m[1]}-${m[2]}-${m[3]}`
    );

  }


  return s;

}


/* =========================================================
   병합 셀 펼치기

   예)
   삼성생명이라는 셀이 5개 행에 병합되어 있으면

   삼성생명
   삼성생명
   삼성생명
   삼성생명
   삼성생명

   형태로 복제
========================================================= */

function cloneSheetWithExpandedMerges(ws) {

  const copy = {};


  /* 일반 셀 복사 */

  for (
    const key of Object.keys(ws)
  ) {

    if (
      key === '!merges'
    ) {

      continue;

    }

    copy[key] =
      ws[key] &&
      typeof ws[key] === 'object'
        ? { ...ws[key] }
        : ws[key];

  }


  copy['!ref'] =
    ws['!ref'];


  /* 병합셀 펼치기 */

  for (
    const merge
    of ws['!merges'] || []
  ) {

    const sourceAddress =
      XLSX.utils.encode_cell(
        merge.s
      );

    const source =
      ws[sourceAddress];

    if (!source) {

      continue;

    }


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

        const addr =
          XLSX.utils.encode_cell({
            r,
            c
          });

        copy[addr] = {
          ...source
        };

      }

    }

  }


  return copy;

}


/* =========================================================
   실제 헤더 행 찾기

   생보협회 엑셀은
   첫 행부터 바로 헤더가 아닐 수 있으므로
   앞쪽 최대 30개 행을 검사
========================================================= */

function findHeaderRowAndColumns(ws) {

  const expanded =
    cloneSheetWithExpandedMerges(
      ws
    );


  const range =
    XLSX.utils.decode_range(
      expanded['!ref'] ||
      'A1:A1'
    );


  const maxRow =
    Math.min(
      range.e.r,
      29
    );


  let best = null;


  for (
    let r = range.s.r;
    r <= maxRow;
    r++
  ) {

    const found = {};

    let score = 0;


    for (
      let c = range.s.c;
      c <= range.e.c;
      c++
    ) {

      const addr =
        XLSX.utils.encode_cell({
          r,
          c
        });


      const value =
        clean(
          expanded[addr]?.v
        );


      if (!value) {

        continue;

      }


      for (
        const [
          field,
          aliases
        ]
        of Object.entries(
          HEADER_ALIASES
        )
      ) {

        if (

          aliases.some(
            alias =>
              norm(value) ===
              norm(alias)
          )

          &&

          found[field] ===
          undefined

        ) {

          found[field] = c;

          score++;

        }

      }

    }


    if (
      !best ||
      score > best.score
    ) {

      best = {

        row: r,

        columns:
          found,

        score

      };

    }

  }


  if (

    !best ||

    best.score < 3 ||

    best.columns.보험회사명 ===
      undefined ||

    best.columns.상품명 ===
      undefined

  ) {

    return null;

  }


  return {

    expanded,

    headerRow:
      best.row,

    columns:
      best.columns

  };

}


/* =========================================================
   개별 XLS 읽기
========================================================= */

function readDataRows(filePath) {

  const category =
    parseCategoryFromFile(
      filePath
    );


  const saving =
    isSavingCategory(
      category,
      filePath
    );


  const wb =
    XLSX.readFile(
      filePath,
      {
        cellDates: true
      }
    );


  const out = [];


  for (
    const sheetName
    of wb.SheetNames
  ) {

    const ws =
      wb.Sheets[
        sheetName
      ];


    if (
      !ws ||
      !ws['!ref']
    ) {

      continue;

    }


    const info =
      findHeaderRowAndColumns(
        ws
      );


    if (!info) {

      console.log(
        `헤더 탐색 실패: ` +
        `${path.basename(filePath)} / ` +
        `${sheetName}`
      );

      continue;

    }


    const {

      expanded,

      headerRow,

      columns

    } = info;


    const range =
      XLSX.utils.decode_range(
        expanded['!ref']
      );


    for (
      let r =
        headerRow + 1;

      r <=
        range.e.r;

      r++
    ) {


      const get =
        field => {

          const c =
            columns[field];


          if (
            c === undefined
          ) {

            return '';

          }


          const addr =
            XLSX.utils.encode_cell({
              r,
              c
            });


          return (
            expanded[addr]?.v ??
            ''
          );

        };


      const company =
        clean(
          get(
            '보험회사명'
          )
        );


      const product =
        clean(
          get(
            '상품명'
          )
        );


      const productParts =
        splitProductName(
          product
        );


      /*
        회사명 / 상품명 없는 행은 제외
      */

      if (
        !company ||
        !product
      ) {

        continue;

      }


      const saleDate =
        excelDateToIso(
          get(
            '판매일자'
          )
        );


      /*
        판매일자 기준 대상월만 추출
      */

      if (

        !saleDate ||

        saleDate.slice(
          0,
          7
        ) !==
        TARGET_MONTH

      ) {

        continue;

      }


      out.push({

        category,

        row: {

          보험회사명:
            company,

          상품명:
            productParts.original,

          정리상품명:
            productParts.normalized,

          상품형태:
            productParts.variants.join(' | '),

          급부명칭:
            saving
              ? ''
              : clean(
                  get(
                    '급부명칭'
                  )
                ),

          지급사유:
            saving
              ? ''
              : clean(
                  get(
                    '지급사유'
                  )
                ),

          지급금액:
            saving
              ? ''
              : clean(
                  get(
                    '지급금액'
                  )
                ),

          판매일자:
            saleDate,

          특이사항:
            clean(
              get(
                '특이사항'
              )
            )

        }

      });

    }

  }


  return out;

}


/* =========================================================
   ★ 수정 핵심 ★

   기존 scraper.js는

   downloads/

   폴더 아래에

   KLIA_보장성_...
   KLIA_저축성_...

   형식으로 저장함.

   파일명 자체에는
   2026-10 같은 TARGET_MONTH가 없으므로

   파일명에 월이 들어가는지 검사하면 안 됨.
========================================================= */

function findRawFiles() {

  const downloadDir =
    path.join(
      ROOT,
      'downloads'
    );


  if (
    !fs.existsSync(
      downloadDir
    )
  ) {

    console.log(
      `downloads 폴더 없음: ${downloadDir}`
    );

    return [];

  }


  const files =
    fs.readdirSync(
      downloadDir,
      {
        withFileTypes: true
      }
    )

    .filter(
      entry =>
        entry.isFile()
    )

    .map(
      entry =>
        entry.name
    )

    .filter(
      name =>
        /\.(xls|xlsx)$/i.test(
          name
        )
    )

    .filter(
      name =>
        /^KLIA_(보장성|저축성)_/i.test(
          name
        )
    )

    .map(
      name =>
        path.join(
          downloadDir,
          name
        )
    )

    .sort(
      (a, b) =>
        a.localeCompare(
          b,
          'ko'
        )
    );


  console.log(
    `downloads 폴더: ${downloadDir}`
  );


  return files;

}


/* =========================================================
   랜덤 선정용

   완전 랜덤이 아니라
   같은 TARGET_MONTH에서는
   같은 결과가 나오도록 seed 사용
========================================================= */

function makeSeed(str) {

  let h =
    2166136261;


  for (
    let i = 0;
    i < str.length;
    i++
  ) {

    h ^=
      str.charCodeAt(i);

    h =
      Math.imul(
        h,
        16777619
      );

  }


  return h >>> 0;

}


function mulberry32(seed) {

  return function () {

    let t =
      seed +=
        0x6D2B79F5;


    t =
      Math.imul(
        t ^ (t >>> 15),
        t | 1
      );


    t ^=
      t +
      Math.imul(
        t ^ (t >>> 7),
        t | 61
      );


    return (
      (
        (
          t ^
          (t >>> 14)
        ) >>> 0
      )
      /
      4294967296
    );

  };

}


/* =========================================================
   카테고리별 상품 1개 선정
========================================================= */

function pickOneByProduct(
  records,
  category,
  seedSuffix
) {

  const grouped =
    new Map();


  for (
    const item
    of records.filter(
      x =>
        x.category ===
        category
    )
  ) {

    const key =
      uniqueProductKey(
        item.row
      );


    if (
      !grouped.has(
        key
      )
    ) {

      grouped.set(
        key,
        []
      );

    }


    grouped
      .get(key)
      .push(
        item.row
      );

  }


  const products =
    [
      ...grouped.values()
    ];


  if (
    !products.length
  ) {

    console.log(
      `뉴스레터 후보 없음: ${category}`
    );

    return null;

  }


  const rng =
    mulberry32(
      makeSeed(
        `${TARGET_MONTH}|${category}|${seedSuffix}`
      )
    );


  const index =
    Math.floor(
      rng() *
      products.length
    );


  const rows =
    products[index];


  /*
    여러 급부행 중
    급부명칭이 있는 첫 행을 대표값으로 사용
  */

  const representative =

    rows.find(
      r =>
        clean(
          r.급부명칭
        )
    )

    ||

    rows[0];


  return {

    카테고리:
      category,

    보험회사명:
      representative
        .보험회사명,

    상품명:
      representative
        .상품명,

    정리상품명:
      representative
        .정리상품명 ||
      representative
        .상품명,

    상품형태:
      representative
        .상품형태 ||
      '',

    급부명칭:
      representative
        .급부명칭,

    지급사유:
      representative
        .지급사유,

    지급금액:
      representative
        .지급금액

  };

}


/* =========================================================
   통합 엑셀 Google Drive 업로드
========================================================= */

async function uploadIfConfigured(
  filePath
) {

  if (
    !UPLOAD_URL ||
    !UPLOAD_SECRET
  ) {

    console.log(
      'Drive 업로드 건너뜀: ' +
      'UPLOAD_URL/UPLOAD_SECRET 없음'
    );

    return;

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

            action:
              'upload',

            folderType:
              'merged',

            fileName:
              path.basename(
                filePath
              ),

            mimeType:
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

            base64:
              fs
                .readFileSync(
                  filePath
                )
                .toString(
                  'base64'
                )

          })

      }
    );


  const text =
    await response.text();


  let data;


  try {

    data =
      JSON.parse(
        text
      );

  } catch {

    throw new Error(
      'Apps Script 응답이 JSON이 아닙니다: ' +
      text.slice(
        0,
        500
      )
    );

  }


  if (
    !response.ok ||
    !data.ok
  ) {

    throw new Error(
      data.error ||
      `Drive 업로드 실패 HTTP ${response.status}`
    );

  }


  console.log(
    `Drive 저장 완료: ` +
    `${path.basename(filePath)}`
  );

}


/* =========================================================
   MAIN
========================================================= */

async function main() {

  const rawFiles =
    findRawFiles();


  console.log(
    `V5 대상월: ${TARGET_MONTH}`
  );


  console.log(
    `V5 원본 엑셀 발견: ${rawFiles.length}개`
  );


  /*
    발견된 원본파일 이름 확인
  */

  for (
    const file
    of rawFiles
  ) {

    console.log(
      `원본 발견: ${path.basename(file)}`
    );

  }


  if (
    !rawFiles.length
  ) {

    throw new Error(
      '대상월 원본 XLS/XLSX 파일을 찾지 못했습니다.'
    );

  }


  const records = [];


  /* =======================================================
     모든 원본 읽기
  ======================================================= */

  for (
    const file
    of rawFiles
  ) {

    const rows =
      readDataRows(
        file
      );


    console.log(
      `- ${path.basename(file)}: ` +
      `${rows.length}행`
    );


    if (
      rows.length
    ) {

      records.push(
        ...rows
      );

    }

  }


  if (
    !records.length
  ) {

    throw new Error(
      `${TARGET_MONTH} 판매일자에 해당하는 행이 없습니다.`
    );

  }


  /* =======================================================
     통합본 전체 행
  ======================================================= */

  const mainRows =
    records.map(
      x =>
        x.row
    );


  /* =======================================================
     상품 수

     보험회사명 + 상품명
     중복 제거
  ======================================================= */

  const uniqueProducts =
    new Set(
      mainRows.map(
        uniqueProductKey
      )
    );


  /* =======================================================
     급부 항목 수

     저축성은 급부명칭 빈칸이라 제외
  ======================================================= */

  const benefitRows =
    mainRows.filter(
      r =>
        clean(
          r.급부명칭
        )
    );


  const uniqueBenefits =
    new Set(
      benefitRows.map(
        productBenefitKey
      )
    );


  /* =======================================================
     카테고리별 상품 수
  ======================================================= */

  const categoryProductCounts =
    {};


  for (
    const item
    of records
  ) {

    if (
      !item.category
    ) {

      continue;

    }


    if (
      !categoryProductCounts[
        item.category
      ]
    ) {

      categoryProductCounts[
        item.category
      ] =
        new Set();

    }


    categoryProductCounts[
      item.category
    ].add(
      uniqueProductKey(
        item.row
      )
    );

  }


  /* =======================================================
     뉴스레터에 넣을 공시 상품 13개

     생명보험협회 13개 상품유형별 각 1개
  ======================================================= */

  const newsletterSamples =
    CATEGORY_ORDER
      .map(
        (category, i) =>
          pickOneByProduct(
            records,
            category,
            i
          )
          ||
          {
            카테고리: category,
            보험회사명: '',
            상품명: '',
            정리상품명: '해당 월 판매개시 상품 없음',
            상품형태: '',
            급부명칭: '',
            지급사유: '',
            지급금액: ''
          }
      );


  /* =======================================================
     요약 시트
  ======================================================= */

  const summaryRows = [

    {
      항목:
        '판매개시 상품 수',

      건수:
        uniqueProducts.size
    },

    {
      항목:
        '급부 항목 수',

      건수:
        uniqueBenefits.size
    },

    {
      항목:
        '통합 행 수',

      건수:
        mainRows.length
    },

    ...Object.entries(
      categoryProductCounts
    )

      .sort(
        (a, b) =>
          a[0].localeCompare(
            b[0],
            'ko'
          )
      )

      .map(
        (
          [
            category,
            set
          ]
        ) => ({

          항목:
            `${category} 상품 수`,

          건수:
            set.size

        })
      )

  ];


  /* =======================================================
     Excel 생성
  ======================================================= */

  const wb =
    XLSX.utils.book_new();


  const wsMain =
    XLSX.utils.json_to_sheet(
      mainRows,
      {
        header:
          TARGET_COLUMNS
      }
    );


  /* 컬럼 폭 */

  wsMain['!cols'] = [

    { wch: 18 },  // 보험회사명
    { wch: 48 },  // 상품명(전체)
    { wch: 40 },  // 정리상품명
    { wch: 42 },  // 상품형태
    { wch: 30 },  // 급부명칭
    { wch: 58 },  // 지급사유
    { wch: 28 },  // 지급금액
    { wch: 14 },  // 판매일자
    { wch: 40 }   // 특이사항

  ];


  /* 자동 필터 */

  wsMain['!autofilter'] = {

    ref:
      XLSX.utils.encode_range({

        s: {
          r: 0,
          c: 0
        },

        e: {
          r:
            mainRows.length,

          c:
            TARGET_COLUMNS.length -
            1
        }

      })

  };


  XLSX.utils.book_append_sheet(
    wb,
    wsMain,
    '통합본'
  );


  /* 요약 시트 */

  const wsSummary =
    XLSX.utils.json_to_sheet(
      summaryRows
    );


  wsSummary['!cols'] = [

    {
      wch: 30
    },

    {
      wch: 12
    }

  ];


  XLSX.utils.book_append_sheet(
    wb,
    wsSummary,
    '요약'
  );


  /* 엑셀 저장 */

  XLSX.writeFile(
    wb,
    OUTPUT_FILE
  );


  /* =======================================================
     뉴스레터용 META JSON 생성
  ======================================================= */

  const meta = {

    targetMonth:
      TARGET_MONTH,

    totalRows:
      mainRows.length,

    productCount:
      uniqueProducts.size,

    benefitCount:
      uniqueBenefits.size,

    categoryProductCounts:
      Object.fromEntries(

        Object.entries(
          categoryProductCounts
        ).map(
          (
            [
              k,
              set
            ]
          ) => [

            k,

            set.size

          ]
        )

      ),

    newsletterSamples

  };


  fs.writeFileSync(

    META_FILE,

    JSON.stringify(
      meta,
      null,
      2
    ),

    'utf8'

  );


  /* =======================================================
     로그
  ======================================================= */

  console.log(
    ''
  );


  console.log(
    '=============================='
  );


  console.log(
    'V5 처리 완료'
  );


  console.log(
    '=============================='
  );


  console.log(
    `통합 행 수: ${mainRows.length}`
  );


  console.log(
    `상품 수(회사+상품명 중복제거): ${uniqueProducts.size}`
  );


  console.log(
    `급부 항목 수: ${uniqueBenefits.size}`
  );


  console.log(
    `뉴스레터 공시상품 선정: ${newsletterSamples.length}건`
  );


  for (
    const sample
    of newsletterSamples
  ) {

    console.log(
      `[${sample.카테고리}] ` +
      `${sample.보험회사명} | ` +
      `${sample.상품명}`
    );

  }


  console.log(
    `V5 통합본 생성: ${OUTPUT_FILE}`
  );


  console.log(
    `V5 메타 생성: ${META_FILE}`
  );


  /* =======================================================
     Drive 업로드
  ======================================================= */

  await uploadIfConfigured(
    OUTPUT_FILE
  );

}


/* =========================================================
   실행
========================================================= */

main().catch(
  err => {

    console.error(
      'process_v5.js 실행 실패'
    );

    console.error(
      err
    );

    process.exit(
      1
    );

  }
);
