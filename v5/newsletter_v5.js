const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUTPUT_DIR = path.join(ROOT, 'output');

const TARGET_MONTH = String(process.env.TARGET_MONTH || '').trim();

const GEMINI_API_KEY =
  process.env.GEMINI_API_KEY || '';

const GEMINI_MODEL =
  process.env.GEMINI_MODEL ||
  'gemini-3.5-flash-lite';

const V5_MAIL_URL =
  process.env.V5_MAIL_URL || '';

const V5_MAIL_SECRET =
  process.env.V5_MAIL_SECRET || '';

const SEND_MAIL =
  String(
    process.env.SEND_MAIL || 'false'
  ).toLowerCase() === 'true';


if (!/^\d{4}-\d{2}$/.test(TARGET_MONTH)) {
  throw new Error(
    'TARGET_MONTH는 YYYY-MM 형식이어야 합니다.'
  );
}


/* =========================================================
   파일명
========================================================= */

const META_FILE = path.join(
  OUTPUT_DIR,
  `${TARGET_MONTH}_V5_meta.json`
);

const EXCEL_FILE = path.join(
  OUTPUT_DIR,
  `${TARGET_MONTH}_생보협회_상품_통합.xlsx`
);

const HTML_FILE = path.join(
  OUTPUT_DIR,
  `${TARGET_MONTH}_민영사_신상품동향.html`
);

const JSON_FILE = path.join(
  OUTPUT_DIR,
  `${TARGET_MONTH}_newsletter.json`
);


/* =========================================================
   생명보험사 키워드
========================================================= */

const LIFE_COMPANY_KEYWORDS = [
  '삼성생명',
  '한화생명',
  '교보생명',
  '신한라이프',

  'NH농협생명',
  '농협생명',

  '동양생명',
  'ABL생명',
  '흥국생명',

  '미래에셋생명',

  '메트라이프생명',
  '메트라이프',

  'KB라이프',
  'KDB생명',

  '푸본현대생명',

  '라이나생명',

  'AIA생명',

  '처브라이프',

  '하나생명',

  'DB생명',

  'BNP파리바카디프생명',
  '카디프생명',

  'IBK연금보험',

  '교보라이프플래닛'
];


/* =========================================================
   제외 키워드
========================================================= */

const EXCLUDE_KEYWORDS = [
  '손해보험',
  '화재',
  '손보',

  '변액',

  '퇴직연금',

  '연금저축',

  '실손'
];


/* =========================================================
   문자열 정리
========================================================= */

function clean(v) {

  return String(v ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

}


/* =========================================================
   HTML Escape
========================================================= */

function escapeHtml(v) {

  return clean(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

}


/* =========================================================
   월별 고정 랜덤
========================================================= */

function makeSeed(str) {

  let h = 2166136261;

  for (
    let i = 0;
    i < str.length;
    i++
  ) {

    h ^= str.charCodeAt(i);

    h = Math.imul(
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


function seededPick(
  items,
  count,
  seedText
) {

  const pool =
    [...items];


  const rng =
    mulberry32(
      makeSeed(
        seedText
      )
    );


  const picked = [];


  while (
    pool.length &&
    picked.length < count
  ) {

    const index =
      Math.floor(
        rng() *
        pool.length
      );


    picked.push(
      pool.splice(
        index,
        1
      )[0]
    );

  }


  return picked;

}


/* =========================================================
   보험신보 작성일 파싱

   예)
   10-14 09:00
   →
   2026-10-14
========================================================= */

function parseInsweekDate(text) {

  const m =
    clean(text).match(
      /(\d{1,2})-(\d{1,2})\s+\d{1,2}:\d{2}/
    );


  if (!m) {

    return '';

  }


  const year =
    TARGET_MONTH.slice(
      0,
      4
    );


  return (
    `${year}-` +
    `${String(
      Number(m[1])
    ).padStart(2, '0')}-` +
    `${String(
      Number(m[2])
    ).padStart(2, '0')}`
  );

}


/* =========================================================
   보험신보 신상품 페이지 기본 URL
========================================================= */

function insweekBaseUrl() {

  return (
    'https://www.insweek.co.kr/news/articleList.html' +
    '?sc_sub_section_code=S2N19' +
    '&view_type=sm'
  );

}


/* =========================================================
   보험신보 실제 pagination URL 읽기

   중요:
   다른 pagination을 잘못 잡지 않도록

   ul.pagination[aria-label="Pagination"]
   +
   sc_sub_section_code=S2N19

   조건을 모두 사용함
========================================================= */

async function getInsweekPageUrls(page) {

  await page.goto(
    insweekBaseUrl(),
    {
      waitUntil:
        'domcontentloaded',

      timeout:
        30000
    }
  );


  await page.waitForTimeout(
    1000
  );


  const hrefs =
    await page

      .locator(
        'ul.pagination[aria-label="Pagination"] ' +
        'a.pagination-link[href*="sc_sub_section_code=S2N19"]'
      )

      .evaluateAll(
        links =>
          links.map(
            a =>
              a.href
          )
      );


  const pageUrls =
    new Map();


  /*
    1페이지는 현재 기본 주소
  */

  pageUrls.set(
    1,
    insweekBaseUrl()
  );


  /*
    실제 사이트에서 제공하는
    2페이지 이후 URL 읽기
  */

  for (
    const href
    of hrefs
  ) {

    try {

      const url =
        new URL(
          href
        );


      /*
        반드시 신상품 카테고리
      */

      if (
        url.searchParams.get(
          'sc_sub_section_code'
        ) !==
        'S2N19'
      ) {

        continue;

      }


      /*
        반드시 view_type=sm
      */

      if (
        url.searchParams.get(
          'view_type'
        ) !==
        'sm'
      ) {

        continue;

      }


      const pageNo =
        Number(
          url.searchParams.get(
            'page'
          )
        );


      if (
        Number.isFinite(
          pageNo
        ) &&
        pageNo >= 2
      ) {

        pageUrls.set(
          pageNo,
          href
        );

      }

    } catch (e) {

      console.log(
        `페이지 URL 해석 실패: ${href}`
      );

    }

  }


  /*
    확인용 로그
  */

  console.log(
    '보험신보 pagination 확인'
  );


  for (
    const [
      pageNo,
      url
    ]
    of pageUrls
  ) {

    if (
      pageNo <= 3
    ) {

      console.log(
        `  ${pageNo}페이지 → ${url}`
      );

    }

  }


  return pageUrls;

}


/* =========================================================
   현재 보험신보 목록 페이지에서

   작성일
   제목
   URL

   추출
========================================================= */

async function extractCurrentList(
  page
) {

  return page.evaluate(
    () => {

      const results = [];


      /*
        보험신보 작성일
      */

      const dateNodes =
        Array.from(
          document.querySelectorAll(
            '.altlist-info-item'
          )
        );


      for (
        const dateEl
        of dateNodes
      ) {

        const dateText =
          (
            dateEl.textContent ||
            ''
          ).trim();


        /*
          작성일 형태

          10-14 09:00
        */

        if (
          !/\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}/
            .test(
              dateText
            )
        ) {

          continue;

        }


        let container =
          dateEl;


        let selectedLink =
          null;


        /*
          날짜가 들어있는 영역에서
          부모를 올라가면서
          같은 기사 링크 탐색
        */

        for (
          let depth = 0;
          depth < 10;
          depth++
        ) {

          if (
            !container
          ) {

            break;

          }


          const links =
            Array.from(
              container.querySelectorAll(
                'a[href*="articleView.html"]'
              )
            )

            .map(
              a => ({

                title:
                  (
                    a.innerText ||
                    a.textContent ||
                    ''
                  ).trim(),

                url:
                  a.href ||
                  ''

              })
            )

            .filter(
              x =>
                x.title &&
                x.url &&
                x.title.length >= 4
            );


          if (
            links.length > 0
          ) {

            /*
              링크가 여러 개면
              가장 긴 텍스트를 제목으로 사용
            */

            links.sort(
              (a, b) =>
                b.title.length -
                a.title.length
            );


            selectedLink =
              links[0];


            break;

          }


          container =
            container.parentElement;

        }


        if (
          !selectedLink
        ) {

          continue;

        }


        results.push({

          title:
            selectedLink.title,

          url:
            selectedLink.url,

          dateText

        });

      }


      return results;

    }
  );

}


/* =========================================================
   보험신보 1~3페이지 수집
========================================================= */

async function collectInsweekList(
  browser
) {

  const page =
    await browser.newPage();


  const start =
    `${TARGET_MONTH}-01`;


  const end =
    `${TARGET_MONTH}-31`;


  const map =
    new Map();


  try {

    /*
      먼저 사이트가 실제로 제공하는
      pagination URL을 읽는다.

      total=2120 같은 숫자는
      직접 코드에 적지 않음.
    */

    const pageUrls =
      await getInsweekPageUrls(
        page
      );


    /*
      1~3페이지 읽기
    */

    for (
      let pageNo = 1;
      pageNo <= 3;
      pageNo++
    ) {

      const targetUrl =
        pageUrls.get(
          pageNo
        );


      if (
        !targetUrl
      ) {

        throw new Error(
          `보험신보 신상품 ${pageNo}페이지 URL을 찾지 못했습니다.`
        );

      }


      console.log(
        `보험신보 ${pageNo}페이지 확인`
      );


      console.log(
        `이동 URL: ${targetUrl}`
      );


      await page.goto(
        targetUrl,
        {

          waitUntil:
            'domcontentloaded',

          timeout:
            30000

        }
      );


      await page.waitForTimeout(
        1000
      );


      console.log(
        `현재 URL: ${page.url()}`
      );


      const raw =
        await extractCurrentList(
          page
        );


      console.log(
        `보험신보 ${pageNo}페이지 날짜+기사 추출: ${raw.length}건`
      );


      /*
        실제로 페이지가 달라졌는지
        앞 5건 출력
      */

      for (
        const item
        of raw.slice(
          0,
          5
        )
      ) {

        console.log(
          `  ${item.dateText} | ${item.title}`
        );

      }


      /*
        대상월 기사만 수집
      */

      for (
        const item
        of raw
      ) {

        if (
          !item.url ||
          !item.title ||
          !item.dateText
        ) {

          continue;

        }


        const articleDate =
          parseInsweekDate(
            item.dateText
          );


        if (
          !articleDate ||
          articleDate < start ||
          articleDate > end
        ) {

          continue;

        }


        const title =
          clean(
            item.title
          );


        const url =
          item.url
            .split('#')[0];


        /*
          URL 기준 중복 제거
        */

        if (
          !map.has(
            url
          )
        ) {

          map.set(
            url,
            {

              title,

              url,

              articleDate

            }
          );

        }

      }

    }

  } finally {

    await page.close();

  }


  const results =
    [...map.values()]
      .sort(
        (a, b) =>
          a.articleDate
            .localeCompare(
              b.articleDate
            )
      );


  console.log(
    `보험신보 ${TARGET_MONTH} 대상월 기사: ${results.length}건`
  );


  for (
    const article
    of results
  ) {

    console.log(
      `${article.articleDate} | ${article.title}`
    );

  }


  return results;

}


/* =========================================================
   보험신보 기사 본문 읽기
========================================================= */

async function enrichArticles(
  browser,
  articles
) {

  const page =
    await browser.newPage();


  const out = [];


  try {

    for (
      const article
      of articles
    ) {

      let body = '';


      try {

        await page.goto(
          article.url,
          {

            waitUntil:
              'domcontentloaded',

            timeout:
              30000

          }
        );


        await page.waitForTimeout(
          500
        );


        body =
          await page.evaluate(
            () => {

              const selectors = [

                '[itemprop="articleBody"]',

                '#article-view-content-div',

                '.article-view-content-div',

                '.article-body',

                '.article-content',

                'article',

                'main'

              ];


              for (
                const selector
                of selectors
              ) {

                const el =
                  document.querySelector(
                    selector
                  );


                const text =
                  el &&
                  (
                    el.innerText ||
                    el.textContent ||
                    ''
                  ).trim();


                if (
                  text &&
                  text.length > 100
                ) {

                  return text;

                }

              }


              return (
                document.body &&
                document.body.innerText
              ) || '';

            }
          );


      } catch (e) {

        console.log(
          `보험신보 본문 실패: ${article.title}`
        );


        console.log(
          e.message
        );

      }


      out.push({

        ...article,

        body:
          clean(
            body
          ).slice(
            0,
            9000
          )

      });

    }

  } finally {

    await page.close();

  }


  return out;

}


/* =========================================================
   생명보험 기사 필터
========================================================= */

function filterLifeInsuranceArticles(
  articles
) {

  const results =
    articles.filter(
      article => {

        const title =
          article.title ||
          '';


        const text =
          `${article.title} ${article.body}`;


        /*
          제목에 생명보험사명 존재
        */

        const titleHasLifeCompany =
          LIFE_COMPANY_KEYWORDS.some(
            keyword =>
              title.includes(
                keyword
              )
          );


        /*
          본문에 생명보험사명 존재
        */

        const bodyHasLifeCompany =
          LIFE_COMPANY_KEYWORDS.some(
            keyword =>
              text.includes(
                keyword
              )
          );


        if (
          !titleHasLifeCompany &&
          !bodyHasLifeCompany
        ) {

          return false;

        }


        /*
          제목 자체가 손보기사면 제외
        */

        const titleHasExcludedKeyword =
          EXCLUDE_KEYWORDS.some(
            keyword =>
              title.includes(
                keyword
              )
          );


        if (
          titleHasExcludedKeyword &&
          !titleHasLifeCompany
        ) {

          return false;

        }


        return true;

      }
    );


  console.log(
    `보험신보 생명보험 기사 필터 후: ${results.length}건`
  );


  for (
    const article
    of results
  ) {

    console.log(
      `${article.articleDate} | ${article.title}`
    );

  }


  return results;

}


/* =========================================================
   보험신보 4건 선정

   1~10일  : 1건
   11~20일 : 2건
   21~말일 : 1건
========================================================= */

function selectFourArticles(
  articles
) {

  const bucket1 =
    articles.filter(
      article => {

        const day =
          Number(
            article.articleDate.slice(
              8,
              10
            )
          );


        return (
          day >= 1 &&
          day <= 10
        );

      }
    );


  const bucket2 =
    articles.filter(
      article => {

        const day =
          Number(
            article.articleDate.slice(
              8,
              10
            )
          );


        return (
          day >= 11 &&
          day <= 20
        );

      }
    );


  const bucket3 =
    articles.filter(
      article => {

        const day =
          Number(
            article.articleDate.slice(
              8,
              10
            )
          );


        return (
          day >= 21
        );

      }
    );


  const picked = [

    ...seededPick(
      bucket1,
      1,
      `${TARGET_MONTH}|1-10`
    ),

    ...seededPick(
      bucket2,
      2,
      `${TARGET_MONTH}|11-20`
    ),

    ...seededPick(
      bucket3,
      1,
      `${TARGET_MONTH}|21-end`
    )

  ];


  /*
    중복 URL 제거
  */

  const uniquePicked = [];


  const used =
    new Set();


  for (
    const item
    of picked
  ) {

    if (
      used.has(
        item.url
      )
    ) {

      continue;

    }


    used.add(
      item.url
    );


    uniquePicked.push(
      item
    );

  }


  /*
    특정 구간에 기사가 부족하면
    나머지 기사에서 보충
  */

  if (
    uniquePicked.length < 4
  ) {

    const remaining =
      articles.filter(
        article =>
          !used.has(
            article.url
          )
      );


    uniquePicked.push(

      ...seededPick(

        remaining,

        4 -
          uniquePicked.length,

        `${TARGET_MONTH}|fallback`

      )

    );

  }


  return uniquePicked.slice(
    0,
    4
  );

}


/* =========================================================
   Gemini 응답 텍스트
========================================================= */

function extractGeminiText(
  data
) {

  const parts =
    data?.candidates?.[0]
      ?.content
      ?.parts || [];


  return parts

    .map(
      p =>
        p.text || ''
    )

    .join(
      '\n'
    )

    .trim();

}


/* =========================================================
   Gemini JSON
========================================================= */

function parseJsonArray(
  text
) {

  const s =
    String(
      text || ''
    )

      .replace(
        /^```json\s*/i,
        ''
      )

      .replace(
        /^```\s*/,
        ''
      )

      .replace(
        /\s*```$/,
        ''
      )

      .trim();


  const first =
    s.indexOf(
      '['
    );


  const last =
    s.lastIndexOf(
      ']'
    );


  if (
    first < 0 ||
    last < first
  ) {

    throw new Error(
      'Gemini JSON 배열을 찾지 못했습니다.'
    );

  }


  return JSON.parse(
    s.slice(
      first,
      last + 1
    )
  );

}


/* =========================================================
   Gemini 기사 4건 요약
========================================================= */

async function summarizeWithGemini(
  articles
) {

  if (
    !articles.length
  ) {

    throw new Error(
      'Gemini에 전달할 보험신보 기사가 없습니다.'
    );

  }


  if (
    !GEMINI_API_KEY
  ) {

    throw new Error(
      'GEMINI_API_KEY가 없습니다.'
    );

  }


  const source =
    articles

      .map(
        (
          article,
          index
        ) => `

[기사 ${index + 1}]

기사작성일:
${article.articleDate}

제목:
${article.title}

URL:
${article.url}

본문:
${article.body}

`
      )

      .join(
        '\n'
      );


  const prompt = `

당신은 우체국보험 상품개발 참고용 월간 뉴스레터를 작성하는 보험상품 분석 담당자입니다.

아래 보험신보 생명보험 신상품 기사들을 기사 내용만 근거로 요약하세요.

규칙:

1. 인터넷 검색을 하지 마세요.
2. 제공된 기사 본문에 없는 사실은 만들지 마세요.
3. 보험회사의 광고성 표현은 최대한 제거하세요.
4. feature는 해당 상품의 핵심 특징을 한 문장으로 작성하세요.
5. implication은 특정 우체국보험 상품과 사실 비교하지 마세요.
6. implication은 우체국보험 상품개발 관점에서 참고할 만한 시사점을 한 문장으로 작성하세요.
7. company에는 보험회사명을 작성하세요.
8. product에는 실제 상품명을 작성하세요.
9. 너무 길게 작성하지 마세요.
10. JSON 배열만 출력하세요.

[
  {
    "index": 1,
    "company": "",
    "product": "",
    "feature": "",
    "implication": ""
  }
]

보험신보 기사:

${source}

`;


  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent`;


  const response =
    await fetch(
      url,
      {

        method:
          'POST',

        headers: {

          'x-goog-api-key':
            GEMINI_API_KEY,

          'Content-Type':
            'application/json'

        },

        body:
          JSON.stringify({

            contents: [

              {

                role:
                  'user',

                parts: [

                  {
                    text:
                      prompt
                  }

                ]

              }

            ],

            generationConfig: {

              temperature:
                0.2,

              responseMimeType:
                'application/json'

            }

          })

      }
    );


  const responseText =
    await response.text();


  if (
    !response.ok
  ) {

    throw new Error(
      `Gemini HTTP ${response.status}: ${responseText}`
    );

  }


  const data =
    JSON.parse(
      responseText
    );


  return parseJsonArray(
    extractGeminiText(
      data
    )
  );

}


/* =========================================================
   생보협회 공시상품 카드
========================================================= */

function buildDisclosureCards(
  samples
) {

  return samples

    .map(
      (
        item,
        index
      ) => `

<tr>

  <td style="
    padding:18px 28px;
    border-bottom:1px solid #eeeeee;
  ">

    <table
      width="100%"
      cellpadding="0"
      cellspacing="0"
      border="0"
      role="presentation"
    >

      <tr>

        <td style="
          width:42px;
          vertical-align:top;
        ">

          <div style="
            width:30px;
            height:30px;
            line-height:30px;
            text-align:center;
            border-radius:50%;
            background:#FFF0EA;
            color:#FF5E1F;
            font-size:12px;
            font-weight:800;
          ">

            ${String(
              index + 1
            ).padStart(
              2,
              '0'
            )}

          </div>

        </td>


        <td style="
          vertical-align:top;
        ">

          <p style="
            margin:0 0 4px 0;
            font-size:11px;
            font-weight:700;
            color:#FF5E1F;
          ">

            ${escapeHtml(
              item.카테고리
            )}

          </p>


          <p style="
            margin:0 0 8px 0;
            font-size:15px;
            font-weight:700;
            color:#222222;
            line-height:1.5;
          ">

            ${escapeHtml(
              item.보험회사명
            )}

            ·

            ${escapeHtml(
              item.상품명
            )}

          </p>


          <p style="
            margin:3px 0;
            font-size:12px;
            color:#666666;
            line-height:1.7;
          ">

            <strong>
              급부명칭
            </strong>

            ${escapeHtml(
              item.급부명칭 ||
              '-'
            )}

          </p>


          <p style="
            margin:3px 0;
            font-size:12px;
            color:#666666;
            line-height:1.7;
          ">

            <strong>
              지급사유
            </strong>

            ${escapeHtml(
              item.지급사유 ||
              '-'
            )}

          </p>


          <p style="
            margin:3px 0;
            font-size:12px;
            color:#666666;
            line-height:1.7;
          ">

            <strong>
              지급금액
            </strong>

            ${escapeHtml(
              item.지급금액 ||
              '-'
            )}

          </p>

        </td>

      </tr>

    </table>

  </td>

</tr>

`
    )

    .join(
      ''
    );

}


/* =========================================================
   보험신보 카드
========================================================= */

function buildNewsCards(
  articles,
  summaries
) {

  return articles

    .map(
      (
        article,
        index
      ) => {

        const summary =
          summaries.find(
            x =>
              Number(
                x.index
              ) ===
              index + 1
          )

          ||

          summaries[
            index
          ]

          ||

          {};


        const day =
          Number(
            article.articleDate.slice(
              8,
              10
            )
          );


        let bucket =
          '21일 ~ 말일';


        if (
          day <= 10
        ) {

          bucket =
            '1일 ~ 10일';

        } else if (
          day <= 20
        ) {

          bucket =
            '11일 ~ 20일';

        }


        return `

<tr>

  <td style="
    padding:20px 28px;
    border-bottom:1px solid #eeeeee;
  ">

    <table
      width="100%"
      cellpadding="0"
      cellspacing="0"
      border="0"
      role="presentation"
    >

      <tr>

        <td style="
          width:42px;
          vertical-align:top;
        ">

          <div style="
            width:30px;
            height:30px;
            line-height:30px;
            text-align:center;
            border-radius:50%;
            background:#FFF0EA;
            color:#FF5E1F;
            font-size:12px;
            font-weight:800;
          ">

            ${String(
              index + 1
            ).padStart(
              2,
              '0'
            )}

          </div>

        </td>


        <td style="
          vertical-align:top;
        ">

          <p style="
            margin:0 0 4px 0;
            font-size:11px;
            font-weight:700;
            color:#FF5E1F;
          ">

            ${bucket}

          </p>


          <p style="
            margin:0 0 7px 0;
            font-size:15px;
            font-weight:700;
            color:#222222;
            line-height:1.5;
          ">

            ${escapeHtml(
              summary.company ||
              ''
            )}

            ${
              summary.company
                ? ' · '
                : ''
            }

            ${escapeHtml(
              summary.product ||
              article.title
            )}

          </p>


          <p style="
            margin:0;
            font-size:13px;
            color:#666666;
            line-height:1.8;
          ">

            <strong>
              [특징]
            </strong>

            ${escapeHtml(
              summary.feature ||
              ''
            )}

          </p>


          <p style="
            margin:6px 0 0 0;
            font-size:13px;
            color:#666666;
            line-height:1.8;
          ">

            <strong>
              [우체국보험 시사점]
            </strong>

            ${escapeHtml(
              summary.implication ||
              ''
            )}

          </p>


          <p style="
            margin:8px 0 0 0;
            font-size:11px;
            color:#aaaaaa;
          ">

            보험신보 |

            ${escapeHtml(
              article.articleDate
            )}

            ·

            <a
              href="${escapeHtml(
                article.url
              )}"
              target="_blank"
              style="
                color:#999999;
                text-decoration:underline;
              "
            >
              원문보기
            </a>

          </p>

        </td>

      </tr>

    </table>

  </td>

</tr>

`;

      }
    )

    .join(
      ''
    );

}


/* =========================================================
   뉴스레터 전체 HTML
========================================================= */

function buildHtml(
  meta,
  articles,
  summaries
) {

  const [
    year,
    month
  ] =
    TARGET_MONTH.split(
      '-'
    );


  const monthNumber =
    Number(
      month
    );


  const disclosureCards =
    buildDisclosureCards(
      meta.newsletterSamples ||
      []
    );


  const newsCards =
    buildNewsCards(
      articles,
      summaries
    );


  return `<!DOCTYPE html>

<html lang="ko">

<head>

<meta charset="UTF-8">

<meta
  name="viewport"
  content="width=device-width, initial-scale=1.0"
>

<style>

@import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');

body,
table,
td,
p,
div,
span,
a,
h1,
h2,
h3 {

  font-family:
    Pretendard,
    "Apple SD Gothic Neo",
    "Malgun Gothic",
    "맑은 고딕",
    Arial,
    sans-serif !important;

}

</style>

</head>


<body style="
  margin:0;
  padding:0;
  background:#f5f5f5;
">


<table
  width="100%"
  cellpadding="0"
  cellspacing="0"
  border="0"
  role="presentation"
  style="
    background:#f5f5f5;
  "
>


<tr>


<td
  align="center"
  style="
    padding:24px 12px;
  "
>


<table
  width="640"
  cellpadding="0"
  cellspacing="0"
  border="0"
  role="presentation"
  style="
    width:100%;
    max-width:640px;
    background:#ffffff;
    border:1px solid #eeeeee;
    border-collapse:collapse;
  "
>


<!-- HEADER -->

<tr>

<td style="
  background:#FF5E1F;
  color:#ffffff;
  padding:30px 28px 25px 28px;
">


<table
  width="100%"
  cellpadding="0"
  cellspacing="0"
  border="0"
  role="presentation"
>

<tr>

<td style="
  font-size:12px;
  font-weight:700;
  letter-spacing:2px;
  color:#FFE6DA;
">

PRODUCT TREND BRIEF

</td>


<td style="
  font-size:12px;
  color:#FFE6DA;
  text-align:right;
">

${year}. ${month}

</td>

</tr>

</table>


<h1 style="
  margin:17px 0 0 0;
  font-size:25px;
  font-weight:800;
  line-height:1.4;
  color:#ffffff;
">

${monthNumber}월 민영사 신상품 동향

</h1>


<p style="
  margin:7px 0 0 0;
  font-size:13px;
  line-height:1.7;
  color:#FFE6DA;
">

이달의 민영 생명보험사 신상품

</p>


</td>

</tr>


<!-- SUMMARY -->

<tr>

<td style="
  padding:22px 28px;
  border-bottom:1px solid #eeeeee;
">


<p style="
  margin:0;
  font-size:14px;
  color:#444444;
  line-height:1.85;
">

${monthNumber}월 생명보험협회 판매일자 기준

<strong>
상품 ${meta.productCount}건
</strong>,

<strong>
급부 항목 ${meta.benefitCount}건
</strong>

이 확인되었습니다.

동일 상품이 보험기간·급부 등에 따라
여러 행으로 구성된 경우
상품 수는 1건으로 계산했습니다.

</p>


</td>

</tr>


<!-- KLIA -->

<tr>

<td style="
  padding:25px 28px 10px 28px;
">


<span style="
  display:inline-block;
  font-size:11px;
  font-weight:700;
  color:#ffffff;
  background:#FF5E1F;
  padding:4px 10px;
  border-radius:3px;
  letter-spacing:1px;
">

KLIA DISCLOSURE

</span>


<h2 style="
  font-size:18px;
  font-weight:800;
  color:#222222;
  margin:11px 0 4px 0;
">

공시에서 살펴본 상품 4선

</h2>


<p style="
  font-size:12px;
  color:#999999;
  margin:0;
">

어린이보험 · 암보험 · 간병/치매보험 · 기타에서
각 1개 상품을 선정했습니다.

</p>


</td>

</tr>


${disclosureCards}


<!-- INSURANCE NEWS -->

<tr>

<td style="
  padding:25px 28px 10px 28px;
">


<span style="
  display:inline-block;
  font-size:11px;
  font-weight:700;
  color:#ffffff;
  background:#FF5E1F;
  padding:4px 10px;
  border-radius:3px;
  letter-spacing:1px;
">

NEWS PICK

</span>


<h2 style="
  font-size:18px;
  font-weight:800;
  color:#222222;
  margin:11px 0 4px 0;
">

보험신보 신상품 4선

</h2>


<p style="
  font-size:12px;
  color:#999999;
  margin:0;
">

기사 작성일 기준
1~10일 1건,
11~20일 2건,
21일~말일 1건을 선정했습니다.

</p>


</td>

</tr>


${newsCards}


<!-- GUIDE -->

<tr>

<td style="
  padding:22px 28px;
  background:#FFF7F3;
  border-top:1px solid #FFD8C8;
">


<p style="
  font-size:13px;
  font-weight:700;
  color:#E84E12;
  margin:0 0 7px 0;
">

더 자세히 보고 싶다면?

</p>


<p style="
  font-size:12px;
  color:#666666;
  line-height:1.8;
  margin:0;
">

자세한 상품 내용은

<strong>
첨부파일(민영 생명보험 상품 통합본)
</strong>

과

<strong>
보험신보 원문
</strong>

을 참고해주시기 바랍니다.

</p>


</td>

</tr>


<!-- FOOTER -->

<tr>

<td style="
  padding:18px 28px;
  background:#2B2B2B;
  color:#999999;
  text-align:center;
  font-size:11px;
  line-height:1.7;
">

민영보험사 상품동향 조사·분석

<br>

우체국보험 상품개발 참고자료

</td>

</tr>


</table>


</td>


</tr>


</table>


</body>

</html>`;

}


/* =========================================================
   Apps Script 메일 발송
========================================================= */

async function sendMail(
  html
) {

  if (
    !SEND_MAIL
  ) {

    console.log(
      '메일 발송 건너뜀: SEND_MAIL=false'
    );


    return;

  }


  if (
    !V5_MAIL_URL ||
    !V5_MAIL_SECRET
  ) {

    throw new Error(
      'SEND_MAIL=true인데 V5_MAIL_URL/V5_MAIL_SECRET이 없습니다.'
    );

  }


  const monthNumber =
    Number(
      TARGET_MONTH.slice(
        5,
        7
      )
    );


  const response =
    await fetch(
      V5_MAIL_URL,
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
              V5_MAIL_SECRET,

            action:
              'sendNewsletterV5',

            targetMonth:
              TARGET_MONTH,

            subject:
              `${monthNumber}월 민영사 신상품 동향`,

            htmlBody:
              html,

            attachmentFileName:
              path.basename(
                EXCEL_FILE
              )

          })

      }
    );


  const responseText =
    await response.text();


  let data;


  try {

    data =
      JSON.parse(
        responseText
      );

  } catch {

    throw new Error(
      'V5 Apps Script 응답이 JSON이 아닙니다: ' +
      responseText.slice(
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
      `V5 메일 발송 실패 HTTP ${response.status}`
    );

  }


  console.log(
    'V5 뉴스레터 메일 발송 완료'
  );

}


/* =========================================================
   MAIN
========================================================= */

async function main() {

  if (
    !fs.existsSync(
      META_FILE
    )
  ) {

    throw new Error(
      `메타파일 없음: ${META_FILE}`
    );

  }


  const meta =
    JSON.parse(
      fs.readFileSync(
        META_FILE,
        'utf8'
      )
    );


  console.log(
    `V5 뉴스레터 대상월: ${TARGET_MONTH}`
  );


  console.log(
    `생보협회 상품 수: ${meta.productCount}`
  );


  console.log(
    `생보협회 급부 항목 수: ${meta.benefitCount}`
  );


  const browser =
    await chromium.launch({

      headless:
        true

    });


  let selected = [];


  try {

    /*
      1.
      보험신보 대상월 기사 목록
    */

    const list =
      await collectInsweekList(
        browser
      );


    if (
      list.length === 0
    ) {

      throw new Error(
        `보험신보 ${TARGET_MONTH} 대상월 기사를 찾지 못했습니다.`
      );

    }


    /*
      2.
      대상월 기사 본문 수집
    */

    console.log(
      `보험신보 대상월 기사 본문 수집 시작: ${list.length}건`
    );


    const enriched =
      await enrichArticles(
        browser,
        list
      );


    /*
      3.
      생명보험 기사만
    */

    const lifeArticles =
      filterLifeInsuranceArticles(
        enriched
      );


    if (
      lifeArticles.length === 0
    ) {

      throw new Error(
        `보험신보 ${TARGET_MONTH} 생명보험 신상품 후보가 0건입니다.`
      );

    }


    /*
      4.
      4건 선정
    */

    selected =
      selectFourArticles(
        lifeArticles
      );


    console.log(
      `보험신보 최종 선정: ${selected.length}건`
    );


    for (
      const article
      of selected
    ) {

      console.log(
        `선정 | ${article.articleDate} | ${article.title}`
      );

    }


    if (
      selected.length === 0
    ) {

      throw new Error(
        '보험신보 최종 선정 기사가 없습니다.'
      );

    }

  } finally {

    await browser.close();

  }


  /* =======================================================
     Gemini
  ======================================================= */

  let summaries = [];


  try {

    summaries =
      await summarizeWithGemini(
        selected
      );


    console.log(
      `Gemini 기사 요약 완료: ${summaries.length}건`
    );


  } catch (e) {

    /*
      Gemini 실패 시에도
      HTML 생성은 진행
    */

    console.log(
      `Gemini 요약 실패: ${e.message}`
    );


    summaries =
      selected.map(
        (
          article,
          index
        ) => ({

          index:
            index + 1,

          company:
            '',

          product:
            article.title,

          feature:
            'AI 자동 요약에 실패했습니다. 보험신보 원문을 참고해주세요.',

          implication:
            '기사 원문 확인 후 상품개발 관점에서 추가 검토가 필요합니다.'

        })
      );

  }


  /* =======================================================
     HTML
  ======================================================= */

  const html =
    buildHtml(
      meta,
      selected,
      summaries
    );


  fs.writeFileSync(
    HTML_FILE,
    html,
    'utf8'
  );


  console.log(
    `V5 뉴스레터 HTML 생성: ${HTML_FILE}`
  );


  /* =======================================================
     JSON
  ======================================================= */

  fs.writeFileSync(

    JSON_FILE,

    JSON.stringify(
      {

        targetMonth:
          TARGET_MONTH,

        productCount:
          meta.productCount,

        benefitCount:
          meta.benefitCount,

        disclosureSamples:
          meta.newsletterSamples ||
          [],

        insweekArticles:
          selected,

        summaries

      },
      null,
      2
    ),

    'utf8'

  );


  console.log(
    `V5 뉴스레터 JSON 생성: ${JSON_FILE}`
  );


  /* =======================================================
     메일
  ======================================================= */

  await sendMail(
    html
  );


  console.log(
    ''
  );


  console.log(
    '=============================='
  );


  console.log(
    'V5 뉴스레터 처리 완료'
  );


  console.log(
    '=============================='
  );

}


/* =========================================================
   실행
========================================================= */

main().catch(
  err => {

    console.error(
      'newsletter_v5.js 실행 실패'
    );


    console.error(
      err
    );


    process.exit(
      1
    );

  }
);
