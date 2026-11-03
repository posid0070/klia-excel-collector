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


// ============================================================
// 기본 함수
// ============================================================

function cleanSpace(value) {

  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function norm(value) {

  return cleanSpace(value)
    .toLowerCase()
    .replace(
      /[\s\(\)\[\]\{\}·ㆍ\-_,.\/\\:：'"]/g,
      ''
    )
    .trim();
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
// 상품군명 정리
// ============================================================

function isVariantParenthesis(text) {

  const raw =
    cleanSpace(text);

  const s =
    raw
      .replace(/\s+/g, '')
      .toLowerCase();


  if (!s) {
    return false;
  }


  /*
   * (무), (무배당)
   */
  if (
    s === '무' ||
    s === '무배당'
  ) {
    return true;
  }


  /*
   * (2601), (2610)
   */
  if (
    /^(?:20)?\d{4}$/.test(s)
  ) {
    return true;
  }


  /*
   * (1종), (2종)
   * (1형), (2형)
   */
  if (
    /^\d+(?:종|형)$/.test(s)
  ) {
    return true;
  }


  /*
   * (Ⅰ형), (Ⅱ형)
   */
  if (
    /^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+형$/i.test(
      raw
    )
  ) {
    return true;
  }


  const variantWords = [

    '표준형',
    '일반형',
    '기본형',
    '선택형',

    '간편형',
    '간편가입형',
    '간편심사형',
    '일반심사형',

    '건강체형',
    '일반체형',
    '비흡연체형',

    '갱신형',
    '비갱신형',

    '순수보장형',
    '만기환급형',

    '해약환급금미지급형',
    '해약환급금일부지급형',
    '해약환급금일부지급형Ⅱ',
    '해약환급금미지급형Ⅱ',

    '무해약환급금형',
    '저해약환급금형',
    '저해지환급형',

    '납입면제형',
    '비납입면제형',

    '체증형',
    '비체증형',

    '유병자형',
    '건강체',

    '일반가입형',
    '표준체형'
  ];


  if (
    variantWords.some(
      word =>
        s.includes(
          word.toLowerCase()
        )
    )
  ) {
    return true;
  }


  /*
   * 1종 표준형
   * 2종 해약환급금 미지급형
   */
  if (
    /^\d+종/.test(s) &&
    (
      s.includes('형') ||
      s.includes('환급')
    )
  ) {
    return true;
  }


  return false;
}


function canonicalizeProductName(
  productName
) {

  const original =
    cleanSpace(productName);


  if (!original) {

    return {
      familyName: '',
      variants: []
    };
  }


  const variants = [];


  let name =
    original.replace(
      /\(([^()]*)\)/g,
      (whole, inside) => {

        if (
          isVariantParenthesis(
            inside
          )
        ) {

          variants.push(
            cleanSpace(inside)
          );

          return ' ';
        }

        return whole;
      }
    );


  /*
   * 앞쪽 무배당
   */
  name =
    name
      .replace(
        /^\s*무배당\s*/i,
        ''
      )
      .replace(
        /^\s*\(무\)\s*/i,
        ''
      );


  /*
   * 끝에 붙은 2601 / 2610
   */
  name =
    name.replace(
      /\s+(?:20)?(?:24|25|26|27|28|29)\d{2}\s*$/i,
      ''
    );


  /*
   * Ver 1 / V2
   */
  name =
    name.replace(
      /\s+(?:v|ver\.?|version)\s*\d+\s*$/i,
      ''
    );


  const trailingVariants = [

    '표준형',
    '일반형',
    '기본형',

    '간편형',
    '간편심사형',

    '갱신형',
    '비갱신형',

    '납입면제형',

    '해약환급금 미지급형',
    '해약환급금 일부지급형'
  ];


  for (
    const variant
    of trailingVariants
  ) {

    const escaped =
      variant.replace(
        /[.*+?^${}()|[\]\\]/g,
        '\\$&'
      );


    const regex =
      new RegExp(
        `\\s+${escaped}\\s*$`,
        'i'
      );


    if (
      regex.test(name)
    ) {

      variants.push(
        variant
      );


      name =
        name.replace(
          regex,
          ''
        );
    }
  }


  name =
    cleanSpace(name)
      .replace(
        /\s+([,.)])/g,
        '$1'
      )
      .replace(
        /[(]\s+/g,
        '('
      )
      .trim();


  /*
   * 너무 짧아졌으면
   * 과도한 제거 방지
   */
  if (
    norm(name).length < 5
  ) {

    name =
      original;
  }


  return {

    familyName:
      name,

    variants:
      [
        ...new Set(
          variants.filter(Boolean)
        )
      ]
  };
}


// ============================================================
// 문자열 유사도
// ============================================================

function levenshtein(a, b) {

  const s = norm(a);
  const t = norm(b);

  const rows =
    s.length + 1;

  const cols =
    t.length + 1;


  const dp =
    Array.from(
      {
        length: rows
      },
      () =>
        new Array(cols)
          .fill(0)
    );


  for (
    let i = 0;
    i < rows;
    i++
  ) {

    dp[i][0] = i;
  }


  for (
    let j = 0;
    j < cols;
    j++
  ) {

    dp[0][j] = j;
  }


  for (
    let i = 1;
    i < rows;
    i++
  ) {

    for (
      let j = 1;
      j < cols;
      j++
    ) {

      const cost =
        s[i - 1] ===
        t[j - 1]
          ? 0
          : 1;


      dp[i][j] =
        Math.min(

          dp[i - 1][j] + 1,

          dp[i][j - 1] + 1,

          dp[i - 1][j - 1] +
            cost
        );
    }
  }


  return dp[
    rows - 1
  ][
    cols - 1
  ];
}


function similarity(a, b) {

  const s = norm(a);
  const t = norm(b);


  if (
    !s ||
    !t
  ) {
    return 0;
  }


  if (
    s === t
  ) {
    return 1;
  }


  const maxLength =
    Math.max(
      s.length,
      t.length
    );


  return (
    1 -
    (
      levenshtein(
        s,
        t
      ) /
      maxLength
    )
  );
}


function commonPrefixLength(a, b) {

  const s = norm(a);
  const t = norm(b);


  const limit =
    Math.min(
      s.length,
      t.length
    );


  let count = 0;


  for (
    let i = 0;
    i < limit;
    i++
  ) {

    if (
      s[i] !== t[i]
    ) {
      break;
    }

    count++;
  }


  return count;
}


function isSafeFuzzyFamilyMatch(
  a,
  b
) {

  const s = norm(a);
  const t = norm(b);


  if (
    !s ||
    !t
  ) {
    return false;
  }


  if (
    s === t
  ) {
    return true;
  }


  const shorter =
    s.length <= t.length
      ? s
      : t;


  const longer =
    s.length > t.length
      ? s
      : t;


  /*
   * 짧은 이름은
   * 유사도 병합 금지
   */
  if (
    shorter.length < 10
  ) {
    return false;
  }


  /*
   * 한쪽 뒤에 의미있는 단어가
   * 붙어있으면 별도 상품으로 처리
   *
   * 예:
   * 행복건강보험
   * 행복건강보험암플랜
   */
  if (
    longer.startsWith(
      shorter
    )
  ) {

    const extra =
      longer.slice(
        shorter.length
      );


    return (
      extra.length <= 2
    );
  }


  const score =
    similarity(
      s,
      t
    );


  const prefix =
    commonPrefixLength(
      s,
      t
    );


  const lengthRatio =
    shorter.length /
    longer.length;


  return (
    score >= 0.96 &&
    lengthRatio >= 0.93 &&
    prefix >=
      Math.min(
        12,
        Math.floor(
          shorter.length *
          0.75
        )
      )
  );
}


// ============================================================
// 상품군 키
// ============================================================

function makeBucketKey(row) {

  /*
   * 같은 회사 + 같은 상품유형 안에서만
   * 상품군을 묶는다.
   *
   * 상품코드/대표번호는 절대 사용하지 않음.
   */
  return [

    norm(
      row['보험구분']
    ),

    norm(
      row['상품유형']
    ),

    norm(
      row['보험회사명']
    )

  ].join('|');
}


function makeFamilyKey(row) {

  return [

    makeBucketKey(row),

    norm(
      row['상품군명']
    )

  ].join('|');
}


function makeRepresentativeKey(row) {

  /*
   * ★ 중요
   *
   * 상품코드 / 대표번호를 쓰지 않는다.
   *
   * 회사 + 실제 상품명만 사용
   */
  return (

    norm(
      row['보험회사명']
    ) +

    '|NAME|' +

    norm(
      row['상품명']
    )
  );
}


// ============================================================
// 현재월 상품군 묶기
// ============================================================

function clusterCurrentProducts(rows) {

  const buckets =
    new Map();


  const output =
    rows.map(
      row => {

        const parsed =
          canonicalizeProductName(
            row['상품명']
          );


        return {

          ...row,

          상품군명:
            parsed.familyName,

          세부유형:
            parsed.variants
              .join(' | '),

          묶음방식:
            ''
        };
      }
    );


  for (
    const row of output
  ) {

    const bucketKey =
      makeBucketKey(row);


    if (
      !buckets.has(
        bucketKey
      )
    ) {

      buckets.set(
        bucketKey,
        []
      );
    }


    const clusters =
      buckets.get(
        bucketKey
      );


    const candidateName =
      row['상품군명'];


    /*
     * 기본상품명 완전 일치
     */
    let cluster =
      clusters.find(
        item =>
          norm(
            item.familyName
          ) ===
          norm(
            candidateName
          )
      );


    let method =
      '기본상품명 일치';


    /*
     * 아주 높은 유사도일 때만 병합
     */
    if (!cluster) {

      cluster =
        clusters.find(
          item =>
            isSafeFuzzyFamilyMatch(
              item.familyName,
              candidateName
            )
        );


      if (cluster) {

        method =
          '고유사도 일치';
      }
    }


    /*
     * 신규 상품군
     */
    if (!cluster) {

      cluster = {

        familyName:
          candidateName,

        rows: []
      };


      clusters.push(
        cluster
      );


      method =
        '신규 상품군';
    }


    cluster.rows.push(row);


    row['상품군명'] =
      cluster.familyName;


    row['상품군키'] =
      [

        bucketKey,

        norm(
          cluster.familyName
        )

      ].join('|');


    row['대표상품키'] =
      makeRepresentativeKey(
        row
      );


    row['묶음방식'] =
      method;
  }


  return output;
}


// ============================================================
// 과거월 비교
// ============================================================

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
          /^\d{4}-\d{2}$/.test(
            value
          )
      )
      .sort();


  return months.length
    ? months[
        months.length - 1
      ]
    : '';
}


function findPriorFamily(
  currentRow,
  priorRows
) {

  const bucket =
    makeBucketKey(
      currentRow
    );


  const candidates =
    priorRows.filter(
      old =>
        makeBucketKey(old) ===
        bucket
    );


  /*
   * 기존 상품군명 완전 일치
   */
  const exact =
    candidates.find(
      old =>
        norm(
          old['상품군명']
        ) ===
        norm(
          currentRow[
            '상품군명'
          ]
        )
    );


  if (exact) {

    return {
      row:
        exact,

      method:
        '기존 상품군명 일치'
    };
  }


  /*
   * 예전 DB 상품명을
   * 현재 정규화 규칙으로 다시 비교
   */
  for (
    const old
    of candidates
  ) {

    const oldCanonical =
      canonicalizeProductName(
        old['상품명']
      ).familyName;


    if (
      norm(
        oldCanonical
      ) ===
      norm(
        currentRow[
          '상품군명'
        ]
      )
    ) {

      return {

        row:
          old,

        method:
          '기존 상품명 재정규화 일치'
      };
    }
  }


  /*
   * 매우 높은 유사도만
   */
  for (
    const old
    of candidates
  ) {

    const oldName =
      old['상품군명'] ||
      canonicalizeProductName(
        old['상품명']
      ).familyName;


    if (
      isSafeFuzzyFamilyMatch(
        oldName,
        currentRow[
          '상품군명'
        ]
      )
    ) {

      return {

        row:
          old,

        method:
          '기존 상품군 고유사도'
      };
    }
  }


  return null;
}


// ============================================================
// 비교 상태
// ============================================================

function annotateComparison(
  currentRows,
  previousRows,
  targetMonth
) {

  /*
   * ★ 매우 중요
   *
   * 대상월보다 과거 데이터만 비교
   *
   * 예:
   * 2026-09 실행
   * → 2026-10은 절대 비교하지 않음
   *
   * 2026-10 실행
   * → 2026-09까지 비교
   */
  const priorRows =
    previousRows.filter(
      row => {

        const month =
          String(
            row['조사월'] || ''
          );


        return (
          /^\d{4}-\d{2}$/.test(
            month
          ) &&
          month < targetMonth
        );
      }
    );


  const comparisonMonth =
    maxMonth(
      priorRows
    );


  const baseline =
    priorRows.length === 0;


  const priorRepresentative =
    new Map();


  for (
    const row
    of priorRows
  ) {

    const representative =
      row['대표상품키'] ||
      makeRepresentativeKey(
        row
      );


    if (
      !priorRepresentative.has(
        representative
      )
    ) {

      priorRepresentative.set(
        representative,
        []
      );
    }


    priorRepresentative
      .get(
        representative
      )
      .push(row);
  }


  for (
    const row
    of currentRows
  ) {

    let status =
      '기준월';


    let comparisonReason =
      '첫 기준월';


    let matchedOldRows =
      [];


    if (!baseline) {

      /*
       * 회사 + 상품명이 완전히 동일
       */
      if (
        priorRepresentative.has(
          row['대표상품키']
        )
      ) {

        status =
          '동일상품 재판매·개정후보';


        comparisonReason =
          '과거 동일 회사·상품명 존재';


        matchedOldRows =
          priorRepresentative.get(
            row['대표상품키']
          );

      } else {

        /*
         * 상품명은 조금 달라도
         * 동일 상품군이면 개정 후보
         */
        const priorFamily =
          findPriorFamily(
            row,
            priorRows
          );


        if (
          priorFamily
        ) {

          status =
            '상품개정후보';


          comparisonReason =
            priorFamily.method;


          matchedOldRows =
            [
              priorFamily.row
            ];

        } else {

          /*
           * 과거 동일상품군 없음
           */
          status =
            '신규상품후보';


          comparisonReason =
            '과거 동일 상품군 없음';
        }
      }
    }


    row['비교상태'] =
      status;


    row['비교근거'] =
      comparisonReason;


    row['비교기준월'] =
      comparisonMonth;


    const months =
      matchedOldRows
        .map(
          old =>
            String(
              old[
                '최초확인월'
              ] ||
              old[
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


    row['최초확인월'] =
      months.length
        ? months[0]
        : targetMonth;
  }


  return currentRows;
}


// ============================================================
// 상품군목록 생성
// ============================================================

function buildFamilyRows(
  products
) {

  const map =
    new Map();


  for (
    const row
    of products
  ) {

    const key =
      row['상품군키'];


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

          상품군명:
            row['상품군명'],

          상품군키:
            key,

          판매일자:
            new Set(),

          대표번호:
            new Set(),

          세부상품명:
            new Set(),

          세부유형:
            new Set(),

          비교상태:
            new Set(),

          묶음방식:
            new Set(),

          상품수:
            0
        }
      );
    }


    const item =
      map.get(key);


    item.상품수 += 1;


    if (
      row['판매일자']
    ) {

      item.판매일자.add(
        row['판매일자']
      );
    }


    if (
      row['대표번호']
    ) {

      item.대표번호.add(
        row['대표번호']
      );
    }


    if (
      row['상품명']
    ) {

      item.세부상품명.add(
        row['상품명']
      );
    }


    if (
      row['세부유형']
    ) {

      String(
        row['세부유형']
      )
        .split('|')
        .map(
          value =>
            cleanSpace(value)
        )
        .filter(Boolean)
        .forEach(
          value =>
            item.세부유형
              .add(value)
        );
    }


    if (
      row['비교상태']
    ) {

      item.비교상태.add(
        row['비교상태']
      );
    }


    if (
      row['묶음방식']
    ) {

      item.묶음방식.add(
        row['묶음방식']
      );
    }
  }


  return [
    ...map.values()
  ].map(
    item => {

      const statuses =
        [
          ...item.비교상태
        ];


      let familyStatus =
        '기준월';


      if (
        statuses.includes(
          '신규상품후보'
        )
      ) {

        familyStatus =
          '신규상품후보';

      } else if (
        statuses.includes(
          '상품개정후보'
        )
      ) {

        familyStatus =
          '상품개정후보';

      } else if (
        statuses.includes(
          '동일상품 재판매·개정후보'
        )
      ) {

        familyStatus =
          '동일상품 재판매·개정후보';
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
            .sort()
            .join(', '),

        대표번호:
          [
            ...item.대표번호
          ]
            .sort()
            .join(' | '),

        세부상품수:
          item.상품수,

        세부유형:
          [
            ...item.세부유형
          ]
            .sort()
            .join(' | '),

        세부상품명:
          [
            ...item.세부상품명
          ]
            .join(' | '),

        비교상태:
          familyStatus,

        묶음방식:
          [
            ...item.묶음방식
          ]
            .join(' | ')
      };
    }
  );
}


// ============================================================
// Worksheet
// ============================================================

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
    XLSX.utils
      .aoa_to_sheet(
        aoa
      );


  if (
    headers.length > 0
  ) {

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
                headers.length -
                1
            }
          })
    };
  }


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
            '세부유형',
            '비교근거'
          ].includes(
            header
          )
        ) {

          return {
            wch: 35
          };
        }


        if (
          [
            '상품군키',
            '대표상품키'
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
      workbook
        .SheetNames
        .filter(
          name =>
            name !==
            sheetName
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


// ============================================================
// Drive
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


async function downloadPreviousDb() {

  const result =
    await apiCall({

      action:
        'download',

      folderType:
        'merged',

      fileName:
        DB_FILE_NAME
    });


  if (
    !result.exists
  ) {

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
      tempPath
    );


  const sheet =
    workbook.Sheets[
      '누적상품DB'
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
// 월별 통합본 읽기
// ============================================================

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
      monthlyPath
    );


  const sheet =
    workbook.Sheets[
      '상품목록'
    ];


  if (!sheet) {

    throw new Error(
      '상품목록 시트가 없습니다.'
    );
  }


  const rows =
    XLSX.utils
      .sheet_to_json(
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


// ============================================================
// 월별 통합본 수정
// ============================================================

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
    '세부유형',

    '판매일자',

    /*
     * process_upload에서
     * 진짜 상품코드가 있는 경우만 값 존재
     * 비교에는 사용하지 않음
     */
    '상품코드',

    /*
     * 대표전화번호
     * 비교에는 사용하지 않음
     */
    '대표번호',

    '대표상품키',
    '상품군키',

    '묶음방식',

    '비교상태',
    '비교근거',

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

    '대표번호',

    '세부상품수',
    '세부유형',
    '세부상품명',

    '비교상태',
    '묶음방식'
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


// ============================================================
// 누적 DB 생성
// ============================================================

function buildCumulativeDb(
  previousRows,
  currentProducts,
  targetMonth
) {

  /*
   * 현재월 기존 데이터 제거 후
   * 새 결과로 교체
   */
  const keepRows =
    previousRows.filter(
      row =>
        String(
          row['조사월'] || ''
        ) !==
        targetMonth
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
        a['상품군명']
      ).localeCompare(
        String(
          b['상품군명']
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
    '세부유형',

    '판매일자',

    '상품코드',
    '대표번호',

    '대표상품키',
    '상품군키',

    '묶음방식',

    '비교상태',
    '비교근거',

    '비교기준월',
    '최초확인월',

    '급부행수',
    '원본파일명',

    /*
     * 이후 rule_classify / AI가
     * 추가하는 열 보존용
     */
    '규칙분류',
    '규칙분류근거',
    '신규급부목록',

    'AI대상여부',
    'AI대상사유',

    'AI보완상태',
    'AI검증결과',
    'AI신뢰도',
    'AI주요특징',
    'AI확인근거',
    'AI출처URL'
  ];


  const workbook =
    XLSX.utils
      .book_new();


  XLSX.utils
    .book_append_sheet(

      workbook,

      worksheetFromRows(
        allRows,
        headers
      ),

      '누적상품DB'
    );


  /*
   * 상품군 현황
   */
  const familyMap =
    new Map();


  for (
    const row
    of allRows
  ) {

    const familyName =
      row['상품군명'] ||
      canonicalizeProductName(
        row['상품명']
      ).familyName;


    const key = [

      makeBucketKey(row),

      norm(
        familyName
      )

    ].join('|');


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
            familyName,

          상품군키:
            key,

          월:
            new Set(),

          상품명:
            new Set()
        }
      );
    }


    const item =
      familyMap.get(key);


    if (
      row['조사월']
    ) {

      item.월.add(
        row['조사월']
      );
    }


    if (
      row['상품명']
    ) {

      item.상품명.add(
        row['상품명']
      );
    }
  }


  const familyHistory =
    [
      ...familyMap.values()
    ]
      .map(
        item => {

          const months =
            [
              ...item.월
            ].sort();


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
              [
                ...item.상품명
              ].join(' | ')
          };
        }
      );


  XLSX.utils
    .book_append_sheet(

      workbook,

      worksheetFromRows(

        familyHistory,

        [
          '보험구분',
          '상품유형',
          '보험회사명',

          '상품군명',
          '상품군키',

          '최초확인월',
          '최근확인월',
          '확인월수',

          '확인상품명'
        ]
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


// ============================================================
// 실행
// ============================================================

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


  console.log(
    `상품군화 전 상품수: ${monthly.rows.length}개`
  );


  const previousRows =
    await downloadPreviousDb();


  console.log(
    `기존 누적DB: ${previousRows.length}행`
  );


  /*
   * 상품군 정규화
   */
  let products =
    clusterCurrentProducts(
      monthly.rows
    );


  /*
   * 과거월 비교
   */
  products =
    annotateComparison(
      products,
      previousRows,
      targetMonth
    );


  /*
   * 상품군 단위 목록
   */
  const familyRows =
    buildFamilyRows(
      products
    );


  console.log(
    `상품목록 원본행: ${products.length}개`
  );


  console.log(
    `실제 상품군: ${familyRows.length}개`
  );


  console.log(
    `상품군 감소: ${
      products.length -
      familyRows.length
    }개`
  );


  const mergedFamilies =
    familyRows
      .filter(
        row =>
          Number(
            row['세부상품수']
          ) > 1
      )
      .sort(
        (a, b) =>
          Number(
            b['세부상품수']
          ) -
          Number(
            a['세부상품수']
          )
      );


  console.log(
    `2개 이상 세부유형이 묶인 상품군: ${mergedFamilies.length}개`
  );


  mergedFamilies
    .slice(0, 20)
    .forEach(
      row => {

        console.log(
          `[묶음] ` +
          `${row['보험회사명']} | ` +
          `${row['상품군명']} | ` +
          `${row['세부상품수']}개`
        );
      }
    );


  /*
   * 월별 통합본 수정
   */
  updateMonthlyWorkbook(
    monthly,
    products,
    familyRows
  );


  /*
   * 누적 DB 수정
   */
  const dbPath =
    buildCumulativeDb(
      previousRows,
      products,
      targetMonth
    );


  /*
   * Drive 교체
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


  console.log('');
  console.log(
    '상품군 강화 + 누적DB 생성 완료'
  );

  console.log(
    '상품코드/대표번호는 동일상품 비교에 사용하지 않습니다.'
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
