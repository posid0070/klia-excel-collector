const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const OUTPUT_DIR =
  path.join(__dirname, 'output');

const GEMINI_API_KEY =
  process.env.GEMINI_API_KEY;

const GEMINI_MODEL =
  process.env.GEMINI_MODEL ||
  'gemini-3.5-flash-lite';

const UPLOAD_URL =
  process.env.UPLOAD_URL;

const UPLOAD_SECRET =
  process.env.UPLOAD_SECRET;

const TARGET_MONTH_OVERRIDE =
  String(
    process.env.TARGET_MONTH || ''
  ).trim();


function sleep(ms) {
  return new Promise(
    resolve =>
      setTimeout(resolve, ms)
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
        .map(
          item => [
            item.type,
            item.value
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


function extractInteractionText(data) {

  if (
    data.output_text &&
    typeof data.output_text === 'string'
  ) {
    return data.output_text;
  }


  const texts = [];


  for (
    const step of
    data.steps || []
  ) {

    if (
      step.type !==
      'model_output'
    ) {
      continue;
    }


    for (
      const content of
      step.content || []
    ) {

      if (
        content.type === 'text' &&
        content.text
      ) {

        texts.push(
          content.text
        );
      }
    }
  }


  return texts.join('\n');
}


function cleanJsonText(text) {

  let cleaned =
    String(text || '')
      .trim();


  cleaned =
    cleaned
      .replace(
        /^```json\s*/i,
        ''
      )
      .replace(
        /^```\s*/i,
        ''
      )
      .replace(
        /```$/i,
        ''
      )
      .trim();


  const start =
    cleaned.indexOf('[');

  const end =
    cleaned.lastIndexOf(']');


  if (
    start >= 0 &&
    end > start
  ) {

    cleaned =
      cleaned.slice(
        start,
        end + 1
      );
  }


  return cleaned;
}


async function callGemini(
  prompt
) {

  if (!GEMINI_API_KEY) {

    throw new Error(
      'GEMINI_API_KEY가 없습니다.'
    );
  }


  const url =
    'https://generativelanguage.googleapis.com/v1beta/interactions';


  const body = {

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
  };


  let lastError;


  for (
    let attempt = 1;
    attempt <= 4;
    attempt++
  ) {

    try {

      const response =
        await fetch(
          url,
          {
            method:
              'POST',

            headers: {

              'Content-Type':
                'application/json',

              'x-goog-api-key':
                GEMINI_API_KEY,

              'Api-Revision':
                '2026-05-20'
            },

            body:
              JSON.stringify(
                body
              )
          }
        );


      const text =
        await response.text();


      if (
        response.status === 429 ||
        response.status === 503
      ) {

        lastError =
          new Error(
            `Gemini HTTP ${response.status}: ${text.slice(0, 500)}`
          );


        await sleep(
          5000 * attempt
        );

        continue;
      }


      if (!response.ok) {

        throw new Error(
          `Gemini HTTP ${response.status}: ` +
          text.slice(
            0,
            1000
          )
        );
      }


      const data =
        JSON.parse(text);


      const output =
        extractInteractionText(
          data
        );


      if (!output) {

        throw new Error(
          'Gemini 출력 텍스트가 없습니다.'
        );
      }


      return {
        output,
        raw: data
      };


    } catch (error) {

      lastError =
        error;


      if (
        attempt < 4
      ) {

        await sleep(
          5000 * attempt
        );
      }
    }
  }


  throw lastError;
}


function makePrompt(
  targetMonth,
  category,
  rows
) {

  const items =
    rows.map(
      (row, index) => ({

        key:
          String(index + 1),

        보험회사명:
          row['보험회사명'],

        상품유형:
          row['상품유형'],

        상품군명:
          row['상품군명'],

        세부상품명:
          row['세부상품명'],

        판매일자:
          row['판매일자'],

        기존비교상태:
          row['비교상태']
      })
    );


  return `
당신은 대한민국 생명보험 신상품 조사 분석 담당자입니다.

조사월은 ${targetMonth}입니다.
상품유형은 ${category}입니다.

아래 상품들은 생명보험협회 상품비교공시에서
판매일자가 ${targetMonth}에 해당하는 것으로 이미 확인된 상품입니다.

각 상품에 대해 Google Search를 활용하여
보험회사 공식 홈페이지·보도자료·상품자료,
생명보험협회,
보험신보,
보험저널 등 신뢰할 수 있는 공개자료를 확인하십시오.

반드시 다음 기준으로 분류하십시오.

[분류값]
1. 신규상품
- 해당 판매일자 전후에 새 상품 출시 또는 판매개시로 확인되는 경우
- 단순 상품명 버전 변경이나 개정은 신규상품으로 보지 않음

2. 상품개정
- 기존 상품의 보장 확대·가입범위 변경·상품구조 변경·갱신 등의 개정
- 2609 → 2610 등의 버전 변경
- 기존 상품 리뉴얼도 여기에 포함

3. 신규담보·특약
- 기존 상품 자체가 새 상품은 아니지만
  새로운 특약·보장·급부가 추가된 사실이 확인되는 경우

4. 확인필요
- 검색해도 신규/개정 여부를 객관적으로 확인하기 어려운 경우
- 추측하지 말 것

[중요 규칙]

- 생명보험협회의 판매일자는 이미 검증되었으므로 날짜를 임의 수정하지 마십시오.
- 기사 작성일이 아니라 실제 판매개시/상품개정 사실을 기준으로 판단하십시오.
- 같은 상품의 형·종·버전 차이를 무조건 신규상품으로 세지 마십시오.
- 근거가 약하면 반드시 확인필요로 분류하십시오.
- 공식자료를 우선하십시오.
- 허위 URL을 만들지 마십시오.

각 상품마다 다음 항목을 반환하십시오.

key
최종분류
신뢰도
주요특징
분류근거
출처URL

신뢰도는
높음 / 보통 / 낮음
중 하나입니다.

출처URL은 실제 검색으로 확인한 URL을 최대 3개까지 배열로 반환하십시오.

반드시 아래와 같은 JSON 배열만 출력하십시오.
설명문, 마크다운, 코드블록을 붙이지 마십시오.

[
  {
    "key":"1",
    "최종분류":"신규상품",
    "신뢰도":"높음",
    "주요특징":"...",
    "분류근거":"...",
    "출처URL":[
      "https://..."
    ]
  }
]

분석 대상:

${JSON.stringify(
  items,
  null,
  2
)}
`;
}


function worksheetToRows(
  worksheet
) {

  return XLSX.utils
    .sheet_to_json(
      worksheet,
      {
        defval: ''
      }
    );
}


function replaceSheet(
  workbook,
  sheetName,
  rows,
  headers
) {

  const aoa = [

    headers,

    ...rows.map(
      row =>
        headers.map(
          h =>
            row[h] ?? ''
        )
    )
  ];


  const sheet =
    XLSX.utils
      .aoa_to_sheet(
        aoa
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
      h => {

        if (
          h === '상품명' ||
          h === '상품군명' ||
          h === '세부상품명'
        ) {
          return {
            wch: 42
          };
        }

        if (
          h === '주요특징' ||
          h === '분류근거'
        ) {
          return {
            wch: 55
          };
        }

        if (
          h === '출처URL'
        ) {
          return {
            wch: 65
          };
        }

        return {
          wch: 18
        };
      }
    );


  if (
    workbook.Sheets[
      sheetName
    ]
  ) {

    delete workbook
      .Sheets[
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


  workbook.SheetNames
    .push(
      sheetName
    );


  workbook.Sheets[
    sheetName
  ] =
    sheet;
}


async function apiUpload(
  filePath,
  fileName
) {

  if (
    !UPLOAD_URL ||
    !UPLOAD_SECRET
  ) {

    throw new Error(
      'UPLOAD_URL 또는 UPLOAD_SECRET이 없습니다.'
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

            action:
              'upload',

            folderType:
              'merged',

            fileName,

            mimeType:
              'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

            base64
          })
      }
    );


  const text =
    await response.text();


  let data;


  try {
    data =
      JSON.parse(text);

  } catch {

    throw new Error(
      'Drive 업로드 응답 오류: ' +
      text.slice(0, 500)
    );
  }


  if (
    !response.ok ||
    !data.ok
  ) {

    throw new Error(
      data.error ||
      `Drive 업로드 HTTP ${response.status}`
    );
  }
}


async function main() {

  const targetMonth =
    getTargetMonth();


  console.log(
    `AI 분석 조사월: ${targetMonth}`
  );


  const monthlyPath =
    path.join(
      OUTPUT_DIR,
      `${targetMonth}_생보협회_상품통합.xlsx`
    );


  const dbPath =
    path.join(
      OUTPUT_DIR,
      '생보협회_상품DB.xlsx'
    );


  if (
    !fs.existsSync(
      monthlyPath
    )
  ) {

    throw new Error(
      '월별 통합파일이 없습니다.'
    );
  }


  const workbook =
    XLSX.readFile(
      monthlyPath
    );


  const familySheet =
    workbook.Sheets[
      '상품군목록'
    ];


  if (!familySheet) {

    throw new Error(
      '상품군목록 시트가 없습니다.'
    );
  }


  const familyRows =
    worksheetToRows(
      familySheet
    );


  console.log(
    `AI 분석 대상 상품군: ${familyRows.length}개`
  );


  /*
   * 상품유형별로 나누면
   * 한 번의 검색에 너무 많은 상품을 넣지 않게 됨
   */
  const categoryMap =
    new Map();


  for (
    const row of familyRows
  ) {

    const category =
      String(
        row[
          '상품유형'
        ] || '기타'
      );


    if (
      !categoryMap.has(
        category
      )
    ) {

      categoryMap.set(
        category,
        []
      );
    }


    categoryMap
      .get(category)
      .push(row);
  }


  const allResults =
    new Map();


  for (
    const [
      category,
      rows
    ] of categoryMap
  ) {

    console.log(
      `검색·분류 시작: ${category} / ${rows.length}개 상품군`
    );


    /*
     * 한 유형에 상품이 지나치게 많으면
     * 12개씩 나눠서 검색
     */
    const chunkSize = 12;


    for (
      let start = 0;
      start < rows.length;
      start += chunkSize
    ) {

      const chunk =
        rows.slice(
          start,
          start + chunkSize
        );


      try {

        const prompt =
          makePrompt(
            targetMonth,
            category,
            chunk
          );


        const response =
          await callGemini(
            prompt
          );


        const jsonText =
          cleanJsonText(
            response.output
          );


        const parsed =
          JSON.parse(
            jsonText
          );


        for (
          let i = 0;
          i < chunk.length;
          i++
        ) {

          const row =
            chunk[i];


          const item =
            parsed.find(
              x =>
                String(x.key) ===
                String(i + 1)
            );


          const familyKey =
            String(
              row[
                '상품군키'
              ] || ''
            );


          if (item) {

            allResults.set(
              familyKey,
              {

                최종분류:
                  item.최종분류 ||
                  '확인필요',

                AI신뢰도:
                  item.신뢰도 ||
                  '낮음',

                주요특징:
                  item.주요특징 ||
                  '',

                분류근거:
                  item.분류근거 ||
                  '',

                출처URL:
                  Array.isArray(
                    item.출처URL
                  )
                    ? item
                        .출처URL
                        .join('\n')
                    : String(
                        item.출처URL ||
                        ''
                      )
              }
            );

          } else {

            allResults.set(
              familyKey,
              {

                최종분류:
                  '확인필요',

                AI신뢰도:
                  '낮음',

                주요특징:
                  '',

                분류근거:
                  'AI 결과에서 해당 상품을 찾지 못함',

                출처URL:
                  ''
              }
            );
          }
        }


      } catch (error) {

        console.error(
          `${category} 일부 분석 실패: ${error.message}`
        );


        /*
         * 한 묶음 실패 때문에
         * 전체 월간 작업을 중단하지 않음
         */
        for (
          const row of chunk
        ) {

          const familyKey =
            String(
              row[
                '상품군키'
              ] || ''
            );


          allResults.set(
            familyKey,
            {

              최종분류:
                '확인필요',

              AI신뢰도:
                '낮음',

              주요특징:
                '',

              분류근거:
                'Gemini 분석 실패: ' +
                error.message,

              출처URL:
                ''
            }
          );
        }
      }


      /*
       * API 과호출 완화
       */
      await sleep(
        2500
      );
    }
  }


  /*
   * 상품군목록 업데이트
   */
  for (
    const row of familyRows
  ) {

    const result =
      allResults.get(
        String(
          row[
            '상품군키'
          ] || ''
        )
      );


    row[
      '최종분류'
    ] =
      result?.최종분류 ||
      '확인필요';


    row[
      'AI신뢰도'
    ] =
      result?.AI신뢰도 ||
      '낮음';


    row[
      '주요특징'
    ] =
      result?.주요특징 ||
      '';


    row[
      '분류근거'
    ] =
      result?.분류근거 ||
      '';


    row[
      '출처URL'
    ] =
      result?.출처URL ||
      '';
  }


  const familyHeaders =
    [
      '조사월',
      '보험구분',
      '상품유형',
      '보험회사명',
      '상품군명',
      '상품군키',
      '판매일자',
      '세부상품수',
      '세부상품명',
      '비교상태',
      '최종분류',
      'AI신뢰도',
      '주요특징',
      '분류근거',
      '출처URL'
    ];


  replaceSheet(
    workbook,
    '상품군목록',
    familyRows,
    familyHeaders
  );


  /*
   * 상품목록에도 같은 AI 결과 전달
   */
  const productSheet =
    workbook.Sheets[
      '상품목록'
    ];


  const productRows =
    worksheetToRows(
      productSheet
    );


  for (
    const row of productRows
  ) {

    const result =
      allResults.get(
        String(
          row[
            '상품군키'
          ] || ''
        )
      );


    row[
      '최종분류'
    ] =
      result?.최종분류 ||
      '확인필요';


    row[
      'AI신뢰도'
    ] =
      result?.AI신뢰도 ||
      '낮음';


    row[
      '주요특징'
    ] =
      result?.주요특징 ||
      '';


    row[
      '분류근거'
    ] =
      result?.분류근거 ||
      '';


    row[
      '출처URL'
    ] =
      result?.출처URL ||
      '';
  }


  const productHeaders =
    [
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
      '최종분류',
      'AI신뢰도',
      '주요특징',
      '분류근거',
      '출처URL',
      '급부행수',
      '원본파일명'
    ];


  replaceSheet(
    workbook,
    '상품목록',
    productRows,
    productHeaders
  );


  /*
   * 요약 시트
   */
  const counts = {};


  for (
    const row of familyRows
  ) {

    const type =
      row[
        '최종분류'
      ] ||
      '확인필요';


    counts[type] =
      (
        counts[type] ||
        0
      ) + 1;
  }


  const summaryRows =
    [
      {
        항목:
          '조사월',

        값:
          targetMonth
      },

      {
        항목:
          '전체 상품군',

        값:
          familyRows.length
      },

      {
        항목:
          '신규상품',

        값:
          counts[
            '신규상품'
          ] || 0
      },

      {
        항목:
          '상품개정',

        값:
          counts[
            '상품개정'
          ] || 0
      },

      {
        항목:
          '신규담보·특약',

        값:
          counts[
            '신규담보·특약'
          ] || 0
      },

      {
        항목:
          '확인필요',

        값:
          counts[
            '확인필요'
          ] || 0
      }
    ];


  replaceSheet(
    workbook,
    'AI분류요약',
    summaryRows,
    [
      '항목',
      '값'
    ]
  );


  XLSX.writeFile(
    workbook,
    monthlyPath
  );


  /*
   * 누적DB에도 최종분류 결과 반영
   */
  if (
    fs.existsSync(
      dbPath
    )
  ) {

    const dbWorkbook =
      XLSX.readFile(
        dbPath
      );


    const dbSheet =
      dbWorkbook.Sheets[
        '누적상품DB'
      ];


    if (dbSheet) {

      const dbRows =
        worksheetToRows(
          dbSheet
        );


      for (
        const row of dbRows
      ) {

        if (
          String(
            row['조사월']
          ) !==
          targetMonth
        ) {
          continue;
        }


        const result =
          allResults.get(
            String(
              row[
                '상품군키'
              ] || ''
            )
          );


        if (!result) {
          continue;
        }


        row[
          '최종분류'
        ] =
          result.최종분류;


        row[
          'AI신뢰도'
        ] =
          result.AI신뢰도;


        row[
          '주요특징'
        ] =
          result.주요특징;


        row[
          '분류근거'
        ] =
          result.분류근거;


        row[
          '출처URL'
        ] =
          result.출처URL;
      }


      const dbHeaders =
        [
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
          '최종분류',
          'AI신뢰도',
          '주요특징',
          '분류근거',
          '출처URL',
          '급부행수',
          '원본파일명'
        ];


      replaceSheet(
        dbWorkbook,
        '누적상품DB',
        dbRows,
        dbHeaders
      );


      XLSX.writeFile(
        dbWorkbook,
        dbPath
      );
    }
  }


  /*
   * Drive 파일 교체
   */
  await apiUpload(
    monthlyPath,
    path.basename(
      monthlyPath
    )
  );


  if (
    fs.existsSync(
      dbPath
    )
  ) {

    await apiUpload(
      dbPath,
      '생보협회_상품DB.xlsx'
    );
  }


  console.log(
    'AI 최종분류 완료'
  );


  console.log(
    counts
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
