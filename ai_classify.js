const { chromium } = require('playwright');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ============================================================
// 기본 설정
// ============================================================

const OUTPUT_DIR = path.join(__dirname, 'output');
const DB_FILE_NAME = '생보협회_상품DB.xlsx';

const UPLOAD_URL = process.env.UPLOAD_URL;
const UPLOAD_SECRET = process.env.UPLOAD_SECRET;
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
const GEMINI_MODEL = process.env.GEMINI_MODEL || 'gemini-2.5-flash-lite';
const TARGET_MONTH_OVERRIDE = String(process.env.TARGET_MONTH || '').trim();

const INSWEEK_LIST_URL =
  'https://www.insweek.co.kr/news/articleList.html?page=1&total=2120&sc_section_code=&sc_sub_section_code=S2N19&sc_serial_code=&sc_area=&sc_level=&sc_article_type=&sc_view_level=&sc_sdate=&sc_edate=&sc_serial_number=&sc_word=&sc_andor=&sc_word2=&box_idxno=&sc_multi_code=&sc_is_image=&sc_is_movie=&sc_user_name=&sc_order_by=E';

const MAX_INSWEEK_PAGES = 20;

// AI 심층분석 대상 회사는 앞으로 고정
const AI_TARGET_COMPANIES = [
  '삼성생명',
  '한화생명',
  '교보생명',
  '신한라이프',
  'NH농협생명'
];

const LIFE_INSURERS = [
  '삼성생명',
  '한화생명',
  '교보생명',
  '신한라이프',
  'NH농협생명',
  '동양생명',
  'ABL생명',
  '흥국생명',
  '미래에셋생명',
  '메트라이프생명',
  'KB라이프',
  'KDB생명',
  '푸본현대생명',
  '라이나생명',
  'AIA생명',
  '처브라이프',
  '하나생명',
  'DB생명',
  'BNP파리바카디프생명',
  'IBK연금보험',
  '교보라이프플래닛'
];

const COMPANY_ALIASES = {
  삼성생명: ['삼성생명'],
  한화생명: ['한화생명'],
  교보생명: ['교보생명'],
  신한라이프: ['신한라이프', '신한생명'],
  NH농협생명: ['NH농협생명', '농협생명'],
  동양생명: ['동양생명'],
  ABL생명: ['ABL생명'],
  흥국생명: ['흥국생명'],
  미래에셋생명: ['미래에셋생명'],
  메트라이프생명: ['메트라이프생명', '메트라이프'],
  KB라이프: ['KB라이프', 'KB생명', '푸르덴셜생명'],
  KDB생명: ['KDB생명'],
  푸본현대생명: ['푸본현대생명'],
  라이나생명: ['라이나생명'],
  AIA생명: ['AIA생명'],
  처브라이프: ['처브라이프'],
  하나생명: ['하나생명'],
  DB생명: ['DB생명'],
  BNP파리바카디프생명: ['BNP파리바카디프생명', '카디프생명'],
  IBK연금보험: ['IBK연금보험'],
  교보라이프플래닛: ['교보라이프플래닛', '라이프플래닛']
};

// ============================================================
// 공통 함수
// ============================================================

function clean(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function norm(value) {
  return clean(value)
    .toLowerCase()
    .replace(/[\s\(\)\[\]\{\}·ㆍ\-_,.\/\\:：'"“”‘’]/g, '');
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function getTargetMonth() {
  if (TARGET_MONTH_OVERRIDE) {
    if (!/^\d{4}-\d{2}$/.test(TARGET_MONTH_OVERRIDE)) {
      throw new Error('TARGET_MONTH는 YYYY-MM 형식이어야 합니다.');
    }
    return TARGET_MONTH_OVERRIDE;
  }

  const parts = Object.fromEntries(
    new Intl.DateTimeFormat('en-US', {
      timeZone: 'Asia/Seoul',
      year: 'numeric',
      month: '2-digit'
    })
      .formatToParts(new Date())
      .map(p => [p.type, p.value])
  );

  let year = Number(parts.year);
  let month = Number(parts.month) - 1;

  if (month === 0) {
    year -= 1;
    month = 12;
  }

  return `${year}-${String(month).padStart(2, '0')}`;
}

function canonicalCompany(value) {
  const s = norm(value);

  for (const [company, aliases] of Object.entries(COMPANY_ALIASES)) {
    if (aliases.some(alias => s.includes(norm(alias)))) {
      return company;
    }
  }

  return clean(value);
}

function detectCompanyFromText(text) {
  const s = norm(text);

  for (const company of LIFE_INSURERS) {
    const aliases = COMPANY_ALIASES[company] || [company];

    if (aliases.some(alias => s.includes(norm(alias)))) {
      return company;
    }
  }

  return '';
}

function isLifeInsurer(company) {
  return LIFE_INSURERS.includes(
    canonicalCompany(company)
  );
}

function isAiTargetCompany(company) {
  return AI_TARGET_COMPANIES.includes(
    canonicalCompany(company)
  );
}

function makeFamilyKey(row) {
  return [
    norm(row['보험구분']),
    norm(row['상품유형']),
    norm(
      canonicalCompany(
        row['보험회사명']
      )
    ),
    norm(
      row['상품군명'] ||
      row['상품명']
    )
  ].join('|');
}

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

function replaceSheet(
  workbook,
  sheetName,
  rows
) {
  const index =
    workbook.SheetNames.indexOf(
      sheetName
    );

  if (index >= 0) {
    workbook.SheetNames.splice(
      index,
      1
    );

    delete workbook.Sheets[
      sheetName
    ];
  }

  const safeRows =
    rows.length
      ? rows
      : [
          {
            안내:
              '해당 데이터 없음'
          }
        ];

  const sheet =
    XLSX.utils.json_to_sheet(
      safeRows
    );

  if (rows.length > 0) {
    const headers =
      Object.keys(
        rows[0]
      );

    sheet['!autofilter'] = {
      ref:
        XLSX.utils.encode_range({
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
              '상품군명',
              '세부상품명',
              '기사제목',
              '보험신보기사제목',
              'AI주요특징',
              'AI확인근거',
              '보험신보출시근거'
            ].includes(
              header
            )
          ) {
            return {
              wch: 48
            };
          }

          if (
            [
              'URL',
              '보험신보URL',
              'AI출처URL'
            ].includes(
              header
            )
          ) {
            return {
              wch: 55
            };
          }

          return {
            wch: 18
          };
        }
      );
  }

  XLSX.utils.book_append_sheet(
    workbook,
    sheet,
    sheetName
  );
}

// ============================================================
// 날짜 처리
// ============================================================

function pad2(value) {
  return String(value)
    .padStart(
      2,
      '0'
    );
}

function makeDate(
  year,
  month,
  day
) {
  const y =
    Number(year);

  const m =
    Number(month);

  const d =
    Number(day);

  if (
    !y ||
    m < 1 ||
    m > 12 ||
    d < 1 ||
    d > 31
  ) {
    return '';
  }

  return (
    `${y}-` +
    `${pad2(m)}-` +
    `${pad2(d)}`
  );
}

function parseArticleListDate(
  text,
  targetMonth
) {
  const s =
    clean(text);

  const [
    targetYearText
  ] =
    targetMonth.split('-');

  const targetYear =
    Number(
      targetYearText
    );

  let match =
    s.match(
      /(20\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/
    );

  if (match) {
    return makeDate(
      match[1],
      match[2],
      match[3]
    );
  }

  match =
    s.match(
      /(?:^|\s)(\d{1,2})-(\d{1,2})(?:\s+\d{1,2}:\d{2})?/
    );

  if (match) {
    return makeDate(
      targetYear,
      match[1],
      match[2]
    );
  }

  match =
    s.match(
      /(?:^|\s)(\d{1,2})\.(\d{1,2})(?:\s+\d{1,2}:\d{2})?/
    );

  if (match) {
    return makeDate(
      targetYear,
      match[1],
      match[2]
    );
  }

  return '';
}

function extractLaunchDateInfo(
  articleText,
  articleDate
) {
  const raw =
    String(
      articleText || ''
    );

  const sentences =
    raw
      .replace(
        /\r/g,
        '\n'
      )
      .split(
        /(?<=[.!?다요])\s+|\n+/
      )
      .map(clean)
      .filter(Boolean);

  const launchKeyword =
    /(출시|판매를?\s*시작|판매개시|판매\s*개시|선보|개정\s*출시|출시할|출시한다|출시했다|판매한다)/;

  const candidates = [];

  for (
    const sentence
    of sentences
  ) {
    if (
      !launchKeyword.test(
        sentence
      )
    ) {
      continue;
    }

    let match =
      sentence.match(
        /(20\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일/
      );

    if (match) {
      candidates.push({
        date:
          makeDate(
            match[1],
            match[2],
            match[3]
          ),

        sentence,

        score: 100
      });

      continue;
    }

    match =
      sentence.match(
        /(\d{1,2})월\s*(\d{1,2})일/
      );

    if (
      match &&
      articleDate
    ) {
      const year =
        articleDate.slice(
          0,
          4
        );

      candidates.push({
        date:
          makeDate(
            year,
            match[1],
            match[2]
          ),

        sentence,

        score: 90
      });

      continue;
    }

    match =
      sentence.match(
        /(?:이달\s*)?(?:지난\s*|오는\s*)?(\d{1,2})일(?:부터|에)?/
      );

    if (
      match &&
      articleDate
    ) {
      const [
        year,
        month
      ] =
        articleDate.split(
          '-'
        );

      candidates.push({
        date:
          makeDate(
            year,
            month,
            match[1]
          ),

        sentence,

        score: 70
      });
    }
  }

  if (
    !candidates.length
  ) {
    return {
      date: '',
      sentence: ''
    };
  }

  candidates.sort(
    (a, b) =>
      b.score -
      a.score
  );

  return {
    date:
      candidates[0].date,

    sentence:
      candidates[0].sentence
  };
}

// ============================================================
// 보험신보 신상품 섹션 수집
// ============================================================

function buildInsweekPageUrl(
  pageNo
) {
  const url =
    new URL(
      INSWEEK_LIST_URL
    );

  url.searchParams.set(
    'page',
    String(pageNo)
  );

  return url.toString();
}

async function getArticleCards(
  page,
  targetMonth
) {
  const raw =
    await page
      .locator(
        'a[href*="articleView.html"]'
      )
      .evaluateAll(
        nodes => {
          const results = [];

          for (
            const a
            of nodes
          ) {
            let parent = a;
            let context = '';

            for (
              let i = 0;
              i < 6;
              i++
            ) {
              parent =
                parent &&
                parent.parentElement;

              if (!parent) {
                break;
              }

              const text =
                (
                  parent.innerText ||
                  ''
                ).trim();

              if (!text) {
                continue;
              }

              if (
                /\b\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}\b/.test(
                  text
                ) ||
                /\b20\d{2}[.\-/]\d{1,2}[.\-/]\d{1,2}\b/.test(
                  text
                )
              ) {
                context = text;
                break;
              }
            }

            results.push({
              url:
                a.href || '',

              title:
                (
                  a.innerText ||
                  a.textContent ||
                  ''
                ).trim(),

              context
            });
          }

          return results;
        }
      );

  const map =
    new Map();

  for (
    const item
    of raw
  ) {
    if (
      !item.url.includes(
        'articleView.html'
      )
    ) {
      continue;
    }

    const articleDate =
      parseArticleListDate(
        item.context,
        targetMonth
      );

    if (!articleDate) {
      continue;
    }

    const url =
      item.url.split(
        '#'
      )[0];

    const existing =
      map.get(url);

    if (
      !existing ||
      clean(
        item.title
      ).length >
        clean(
          existing.title
        ).length
    ) {
      map.set(
        url,
        {
          url,

          title:
            clean(
              item.title
            ),

          context:
            clean(
              item.context
            ),

          articleDate
        }
      );
    }
  }

  return [
    ...map.values()
  ];
}

async function collectTargetMonthInsweekArticles(
  browser,
  targetMonth
) {
  const page =
    await browser.newPage();

  const start =
    `${targetMonth}-01`;

  const end =
    `${targetMonth}-31`;

  const collected =
    new Map();

  try {
    for (
      let pageNo = 1;
      pageNo <=
      MAX_INSWEEK_PAGES;
      pageNo++
    ) {
      const url =
        buildInsweekPageUrl(
          pageNo
        );

      console.log(
        `보험신보 신상품 ${pageNo}페이지 확인`
      );

      try {
        await page.goto(
          url,
          {
            waitUntil:
              'domcontentloaded',

            timeout:
              30000
          }
        );

        await page.waitForTimeout(
          600
        );

      } catch (error) {
        console.log(
          `보험신보 목록 접속 실패: ${error.message}`
        );

        continue;
      }

      const cards =
        await getArticleCards(
          page,
          targetMonth
        );

      if (!cards.length) {
        continue;
      }

      for (
        const card
        of cards
      ) {
        if (
          card.articleDate >=
            start &&
          card.articleDate <=
            end
        ) {
          collected.set(
            card.url,
            card
          );
        }
      }

      if (
        cards.every(
          card =>
            card.articleDate <
            start
        )
      ) {
        break;
      }

      const hasTarget =
        cards.some(
          card =>
            card.articleDate >=
              start &&
            card.articleDate <=
              end
        );

      const hasOlder =
        cards.some(
          card =>
            card.articleDate <
            start
        );

      if (
        hasTarget &&
        hasOlder
      ) {
        break;
      }
    }

  } finally {
    await page.close();
  }

  return [
    ...collected.values()
  ].sort(
    (a, b) =>
      b.articleDate.localeCompare(
        a.articleDate
      )
  );
}

async function fetchArticleDetails(
  page,
  article
) {
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
      400
    );

  } catch (error) {
    return {
      ...article,

      company:
        detectCompanyFromText(
          `${article.title} ${article.context}`
        ),

      body: '',

      launchDate: '',

      launchSentence: '',

      fetchError:
        error.message
    };
  }

  const data =
    await page.evaluate(
      () => {
        const firstText =
          selectors => {
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

              if (text) {
                return text;
              }
            }

            return '';
          };

        const title =
          firstText([
            'h1',
            '.article-head-title',
            '.article-view-head-title',
            '.article-title'
          ]);

        const body =
          firstText([
            '[itemprop="articleBody"]',
            '#article-view-content-div',
            '.article-view-content-div',
            '.article-body',
            '.article-content',
            'article'
          ]);

        return {
          title,

          body:
            body ||
            document.body
              ?.innerText ||
            ''
        };
      }
    );

  const title =
    clean(
      data.title ||
      article.title
    );

  const body =
    clean(
      data.body
    ).slice(
      0,
      16000
    );

  const company =
    detectCompanyFromText(
      `${title} ${body}`
    );

  const launch =
    extractLaunchDateInfo(
      body,
      article.articleDate
    );

  return {
    ...article,

    title,

    company,

    body,

    launchDate:
      launch.date,

    launchSentence:
      launch.sentence,

    fetchError: ''
  };
}

async function enrichInsweekArticles(
  browser,
  articles
) {
  const page =
    await browser.newPage();

  const results = [];

  try {
    for (
      const article
      of articles
    ) {
      console.log(
        `보험신보 본문 확인: ${article.title}`
      );

      results.push(
        await fetchArticleDetails(
          page,
          article
        )
      );

      await page.waitForTimeout(
        200
      );
    }

  } finally {
    await page.close();
  }

  return results;
}

// ============================================================
// 상품 ↔ 보험신보 기사 매칭
// ============================================================

function extractQuotedNames(
  text
) {
  const results = [];

  const source =
    String(
      text || ''
    );

  const regexes = [
    /[‘']([^’']{3,100})[’']/g,
    /[“"]([^”"]{3,100})[”"]/g
  ];

  for (
    const regex
    of regexes
  ) {
    let match;

    while (
      (
        match =
          regex.exec(
            source
          )
      ) !== null
    ) {
      results.push(
        clean(
          match[1]
        )
      );
    }
  }

  return [
    ...new Set(
      results
    )
  ];
}

function levenshtein(
  a,
  b
) {
  const s =
    norm(a);

  const t =
    norm(b);

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
        new Array(
          cols
        ).fill(0)
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
          dp[i - 1][j] +
            1,

          dp[i][j - 1] +
            1,

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

function similarity(
  a,
  b
) {
  const s =
    norm(a);

  const t =
    norm(b);

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

  return (
    1 -
    levenshtein(
      s,
      t
    ) /
      Math.max(
        s.length,
        t.length
      )
  );
}

function productMatchScore(
  target,
  article
) {
  const targetCompany =
    canonicalCompany(
      target[
        '보험회사명'
      ]
    );

  const articleCompany =
    canonicalCompany(
      article.company
    );

  if (
    !targetCompany ||
    targetCompany !==
      articleCompany
  ) {
    return 0;
  }

  const family =
    clean(
      target[
        '상품군명'
      ]
    );

  const detailNames =
    clean(
      target[
        '세부상품명'
      ]
    );

  const title =
    clean(
      article.title
    );

  const body =
    clean(
      article.body
    );

  let score = 35;

  if (
    norm(family) &&
    norm(title).includes(
      norm(family)
    )
  ) {
    score += 60;

  } else if (
    norm(family) &&
    norm(body).includes(
      norm(family)
    )
  ) {
    score += 40;
  }

  const targetNames = [
    family,

    ...detailNames
      .split('|')
      .map(clean)
      .filter(Boolean)
  ];

  const quotedNames =
    extractQuotedNames(
      title
    );

  let bestSimilarity = 0;

  for (
    const targetName
    of targetNames
  ) {
    for (
      const quoted
      of quotedNames
    ) {
      bestSimilarity =
        Math.max(
          bestSimilarity,

          similarity(
            targetName,
            quoted
          )
        );
    }
  }

  score +=
    Math.round(
      bestSimilarity *
      50
    );

  return score;
}

function findBestInsweekArticle(
  target,
  articles
) {
  const ranked =
    articles
      .map(
        article => ({
          article,

          score:
            productMatchScore(
              target,
              article
            )
        })
      )
      .filter(
        item =>
          item.score > 0
      )
      .sort(
        (a, b) =>
          b.score -
          a.score
      );

  if (
    !ranked.length
  ) {
    return null;
  }

  if (
    ranked[0].score <
    55
  ) {
    return null;
  }

  return ranked[0];
}

// ============================================================
// Gemini 보완
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
    !Array.isArray(
      data.steps
    )
  ) {
    return '';
  }

  const texts = [];

  for (
    const step
    of data.steps
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

  return texts
    .join('\n')
    .trim();
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
    const step
    of data.steps
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
          'url_citation'
        ) {
          const url =
            ann.url ||
            ann.uri ||
            '';

          if (url) {
            urls.add(url);
          }
        }
      }
    }
  }

  return [
    ...urls
  ];
}

function parseJsonText(
  text
) {
  let s =
    String(
      text || ''
    ).trim();

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
      'Gemini JSON 응답을 찾을 수 없습니다.'
    );
  }

  return JSON.parse(
    s.slice(
      first,
      last + 1
    )
  );
}

async function geminiRequest(
  prompt
) {
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
            'application/json'
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

  if (
    !response.ok
  ) {
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

  const parsed =
    parseJsonText(
      extractOutputText(
        data
      )
    );

  return {
    ...parsed,

    sourceUrls:
      extractSourceUrls(
        data
      )
  };
}

async function callGemini(
  target,
  matchedArticle,
  targetMonth
) {
  const insweekInfo =
    matchedArticle
      ? `
[보험신보 확인자료]
기사작성일: ${matchedArticle.articleDate || ''}
기사제목: ${matchedArticle.title || ''}
기사URL: ${matchedArticle.url || ''}
본문에서 추출한 출시일: ${matchedArticle.launchDate || '미확인'}
출시일 근거문장: ${matchedArticle.launchSentence || '미확인'}
기사본문 일부:
${clean(matchedArticle.body).slice(0, 6500)}
`
      : `
[보험신보 확인자료]
${targetMonth} 기사작성일 기준 보험신보 신상품 섹션에서 이 상품과 확실히 매칭되는 기사를 찾지 못했습니다.
`;

  const prompt = `
당신은 한국 생명보험 상품동향 조사 담당자입니다.

아래 상품은 생명보험협회 공시자료를 규칙 비교한 결과 AI 보완 대상으로 선정되었습니다.
보험신보 신상품 섹션은 프로그램이 먼저 기사작성일 기준으로 직접 확인했습니다.

[대상상품]
조사월: ${targetMonth}
보험회사: ${canonicalCompany(target['보험회사명'])}
상품유형: ${target['상품유형'] || ''}
상품군명: ${target['상품군명'] || ''}
세부상품명: ${target['세부상품명'] || ''}
생명보험협회 판매일자: ${target['판매일자'] || ''}
규칙분류: ${target['규칙분류'] || ''}
신규급부 후보: ${target['신규급부목록'] || ''}
${insweekInfo}

[추가검색]
보험신보만으로 부족할 때에만 Google Search로 보완하세요.

우선순위:
1. 보험저널
2. 해당 보험회사 공식 홈페이지·뉴스룸·보도자료·상품자료실
3. 생명보험협회 등 공식 보험 관련 자료

[중요]
- 기사 작성월과 실제 출시월을 구분하세요.
- 보험신보 기사가 ${targetMonth}에 작성됐더라도 실제 출시일이 다음달이면 그대로 표시하세요.
- 협회 판매일자와 기사 출시일이 다르면 억지로 맞추지 마세요.
- 신규상품, 신규급부·특약, 상품개정은 근거가 확인될 때만 판단하세요.
- 자료가 부족하면 자료확인불가로 표시하세요.
- 광고성 표현은 제거하세요.
- 주요특징은 2문장 이내로 작성하세요.

반드시 아래 JSON 하나만 출력하세요.

{
  "AI검증결과": "신규상품|신규급부·특약|상품개정|자료확인불가",
  "AI신뢰도": "높음|보통|낮음",
  "AI주요특징": "",
  "AI신규급부특약": "",
  "AI확인근거": "",
  "AI출시일검증": "일치|일자불일치|출시일미확인"
}
`;

  try {
    return await geminiRequest(
      prompt
    );

  } catch (error) {
    if (
      error.status ===
      429
    ) {
      throw error;
    }

    if (
      [
        500,
        502,
        503,
        504
      ].includes(
        error.status
      )
    ) {
      await sleep(
        10000
      );

      return await geminiRequest(
        prompt
      );
    }

    throw error;
  }
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
      JSON.parse(
        text
      );

  } catch {
    throw new Error(
      `Apps Script 응답 오류: ${text.slice(0, 300)}`
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
// 누적 DB 업데이트
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

    Object.assign(
      row,
      result
    );
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

  // ==========================================================
  // AI 대상 = 고정 5개사 + Y + 신규상품/신규급부 후보
  // ==========================================================

  const targets =
    familyRows.filter(
      row =>
        row[
          'AI대상여부'
        ] ===
          'Y' &&

        isAiTargetCompany(
          row[
            '보험회사명'
          ]
        ) &&

        [
          '신규상품후보',
          '신규급부·특약후보'
        ].includes(
          row[
            '규칙분류'
          ]
        )
    );

  console.log(
    `조사월: ${targetMonth}`
  );

  console.log(
    `전체 상품군: ${familyRows.length}개`
  );

  console.log(
    `AI 대상: ${targets.length}개`
  );

  // ==========================================================
  // 1. 보험신보 기사작성일 기준 대상월 기사 수집
  // ==========================================================

  console.log('');

  console.log(
    '보험신보 신상품 섹션 수집 시작'
  );

  const browser =
    await chromium.launch({
      headless: true
    });

  let insweekArticles = [];

  try {
    const listArticles =
      await collectTargetMonthInsweekArticles(
        browser,
        targetMonth
      );

    console.log(
      `보험신보 ${targetMonth} 기사작성일 기준: ${listArticles.length}건`
    );

    insweekArticles =
      await enrichInsweekArticles(
        browser,
        listArticles
      );

  } finally {
    await browser.close();
  }

  // ==========================================================
  // 2. 보험신보기사 시트
  // ==========================================================

  const insweekSheetRows =
    insweekArticles.map(
      article => {
        const company =
          canonicalCompany(
            article.company
          );

        const launchMonth =
          article.launchDate
            ? article.launchDate.slice(
                0,
                7
              )
            : '';

        return {
          기사작성월:
            article.articleDate
              ? article.articleDate.slice(
                  0,
                  7
                )
              : '',

          기사작성일:
            article.articleDate ||
            '',

          보험회사명:
            company ||
            '',

          생보여부:
            company &&
            isLifeInsurer(
              company
            )
              ? 'Y'
              : 'N',

          기사제목:
            article.title ||
            '',

          출시일:
            article.launchDate ||
            '',

          출시월:
            launchMonth,

          대상월출시여부:
            !launchMonth
              ? '미확인'
              : launchMonth ===
                targetMonth
              ? 'Y'
              : 'N',

          출시근거문장:
            article.launchSentence ||
            '',

          URL:
            article.url ||
            '',

          수집상태:
            article.fetchError
              ? `본문수집오류: ${article.fetchError}`
              : '정상'
        };
      }
    );

  replaceSheet(
    workbook,
    '보험신보기사',
    insweekSheetRows
  );

  // ==========================================================
  // 3. AI 대상상품 ↔ 보험신보 기사 매칭
  // ==========================================================

  const resultMap =
    new Map();

  let insweekMatched = 0;
  let aiCompleted = 0;
  let quotaExceeded = 0;
  let aiErrors = 0;

  for (
    const target
    of targets
  ) {
    const familyKey =
      makeFamilyKey(
        target
      );

    const ranked =
      findBestInsweekArticle(
        target,
        insweekArticles
      );

    const matched =
      ranked
        ? ranked.article
        : null;

    if (matched) {
      insweekMatched += 1;
    }

    const launchMonth =
      matched &&
      matched.launchDate
        ? matched.launchDate.slice(
            0,
            7
          )
        : '';

    let result = {
      보험신보매칭여부:
        matched
          ? 'Y'
          : 'N',

      보험신보매칭점수:
        ranked
          ? ranked.score
          : '',

      보험신보기사작성일:
        matched
          ? matched.articleDate
          : '',

      보험신보기사제목:
        matched
          ? matched.title
          : '',

      보험신보출시일:
        matched
          ? matched.launchDate
          : '',

      보험신보출시월일치:
        matched
          ? !launchMonth
            ? '미확인'
            : launchMonth ===
              targetMonth
            ? 'Y'
            : 'N'
          : '',

      보험신보출시근거:
        matched
          ? matched.launchSentence
          : '',

      보험신보URL:
        matched
          ? matched.url
          : '',

      AI보완상태:
        GEMINI_API_KEY
          ? '대기'
          : 'Gemini키없음',

      AI검증결과: '',

      AI신뢰도: '',

      AI주요특징: '',

      AI신규급부특약: '',

      AI확인근거: '',

      AI출시일검증: '',

      AI출처URL:
        matched
          ? matched.url
          : ''
    };

    if (
      !GEMINI_API_KEY
    ) {
      result.AI보완상태 =
        matched
          ? '보험신보확인·AI미실행'
          : 'AI미실행';

      Object.assign(
        target,
        result
      );

      resultMap.set(
        familyKey,
        result
      );

      continue;
    }

    console.log(
      `AI 보완: ${canonicalCompany(
        target['보험회사명']
      )} / ${target['상품군명']}`
    );

    try {
      const ai =
        await callGemini(
          target,
          matched,
          targetMonth
        );

      const sourceUrls =
        new Set();

      if (
        matched &&
        matched.url
      ) {
        sourceUrls.add(
          matched.url
        );
      }

      for (
        const url
        of ai.sourceUrls ||
        []
      ) {
        sourceUrls.add(
          url
        );
      }

      result = {
        ...result,

        AI보완상태:
          '완료',

        AI검증결과:
          ai.AI검증결과 ||
          '',

        AI신뢰도:
          ai.AI신뢰도 ||
          '',

        AI주요특징:
          ai.AI주요특징 ||
          '',

        AI신규급부특약:
          ai.AI신규급부특약 ||
          '',

        AI확인근거:
          ai.AI확인근거 ||
          '',

        AI출시일검증:
          ai.AI출시일검증 ||
          '',

        AI출처URL:
          [
            ...sourceUrls
          ].join(
            '\n'
          )
      };

      aiCompleted += 1;

    } catch (error) {
      if (
        error.status ===
        429
      ) {
        result = {
          ...result,

          AI보완상태:
            matched
              ? '보험신보확인·AI쿼터초과'
              : 'AI쿼터초과',

          AI확인근거:
            'Gemini API 사용량 한도 초과. 규칙분류와 보험신보 수집결과는 유지함.'
        };

        quotaExceeded += 1;

      } else {
        result = {
          ...result,

          AI보완상태:
            matched
              ? '보험신보확인·AI오류'
              : 'AI오류',

          AI확인근거:
            String(
              error.message ||
              error
            ).slice(
              0,
              600
            )
        };

        aiErrors += 1;
      }
    }

    Object.assign(
      target,
      result
    );

    resultMap.set(
      familyKey,
      result
    );

    await sleep(
      4000
    );
  }

  // ==========================================================
  // 4. 상품군목록 반영
  // ==========================================================

  for (
    const row
    of familyRows
  ) {
    const result =
      resultMap.get(
        makeFamilyKey(
          row
        )
      );

    if (result) {
      Object.assign(
        row,
        result
      );

    } else if (
      row[
        'AI대상여부'
      ] !==
      'Y'
    ) {
      row[
        'AI보완상태'
      ] =
        '비대상';
    }
  }

  replaceSheet(
    workbook,
    '상품군목록',
    familyRows
  );

  // ==========================================================
  // 5. 상품목록에도 결과 반영
  // ==========================================================

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

    if (result) {
      Object.assign(
        row,
        result
      );

    } else if (
      row[
        'AI대상여부'
      ] !==
      'Y'
    ) {
      row[
        'AI보완상태'
      ] =
        '비대상';
    }
  }

  replaceSheet(
    workbook,
    '상품목록',
    productRows
  );

  // ==========================================================
  // 6. AI분석대상 시트 최신화
  // ==========================================================

  const aiTargetRows =
    familyRows
      .filter(
        row =>
          row[
            'AI대상여부'
          ] ===
          'Y'
      )
      .map(
        row => ({
          조사월:
            row['조사월'],

          보험회사명:
            canonicalCompany(
              row[
                '보험회사명'
              ]
            ),

          상품유형:
            row[
              '상품유형'
            ],

          상품군명:
            row[
              '상품군명'
            ],

          판매일자:
            row[
              '판매일자'
            ],

          세부상품명:
            row[
              '세부상품명'
            ],

          규칙분류:
            row[
              '규칙분류'
            ],

          신규급부목록:
            row[
              '신규급부목록'
            ],

          보험신보매칭여부:
            row[
              '보험신보매칭여부'
            ] || '',

          보험신보기사작성일:
            row[
              '보험신보기사작성일'
            ] || '',

          보험신보기사제목:
            row[
              '보험신보기사제목'
            ] || '',

          보험신보출시일:
            row[
              '보험신보출시일'
            ] || '',

          보험신보출시월일치:
            row[
              '보험신보출시월일치'
            ] || '',

          보험신보출시근거:
            row[
              '보험신보출시근거'
            ] || '',

          보험신보URL:
            row[
              '보험신보URL'
            ] || '',

          AI보완상태:
            row[
              'AI보완상태'
            ] || '',

          AI검증결과:
            row[
              'AI검증결과'
            ] || '',

          AI신뢰도:
            row[
              'AI신뢰도'
            ] || '',

          AI주요특징:
            row[
              'AI주요특징'
            ] || '',

          AI신규급부특약:
            row[
              'AI신규급부특약'
            ] || '',

          AI출시일검증:
            row[
              'AI출시일검증'
            ] || '',

          AI확인근거:
            row[
              'AI확인근거'
            ] || '',

          AI출처URL:
            row[
              'AI출처URL'
            ] || ''
        })
      );

  replaceSheet(
    workbook,
    'AI분석대상',
    aiTargetRows
  );

  // ==========================================================
  // 7. AI분류요약 시트
  // ==========================================================

  const lifeNewsCount =
    insweekSheetRows
      .filter(
        row =>
          row[
            '생보여부'
          ] ===
          'Y'
      )
      .length;

  const lifeNewsTargetLaunchCount =
    insweekSheetRows
      .filter(
        row =>
          row[
            '생보여부'
          ] ===
            'Y' &&
          row[
            '대상월출시여부'
          ] ===
            'Y'
      )
      .length;

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
        targets.length
    },

    {
      구분:
        '보험신보 대상월 작성기사',

      건수:
        insweekArticles.length
    },

    {
      구분:
        '보험신보 대상월 생보기사',

      건수:
        lifeNewsCount
    },

    {
      구분:
        '보험신보 생보기사 중 대상월 출시확인',

      건수:
        lifeNewsTargetLaunchCount
    },

    {
      구분:
        'AI대상 중 보험신보 기사매칭',

      건수:
        insweekMatched
    },

    {
      구분:
        'AI 보완 완료',

      건수:
        aiCompleted
    },

    {
      구분:
        'AI 쿼터초과',

      건수:
        quotaExceeded
    },

    {
      구분:
        'AI 오류',

      건수:
        aiErrors
    }
  ];

  replaceSheet(
    workbook,
    'AI분류요약',
    summaryRows
  );

  // ==========================================================
  // 8. 저장 + 누적DB + Drive
  // ==========================================================

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
    '============================'
  );

  console.log(
    `보험신보 ${targetMonth} 작성기사: ${insweekArticles.length}건`
  );

  console.log(
    `보험신보 생보기사: ${lifeNewsCount}건`
  );

  console.log(
    `AI 대상: ${targets.length}개`
  );

  console.log(
    `보험신보 매칭: ${insweekMatched}개`
  );

  console.log(
    `AI 완료: ${aiCompleted}개`
  );

  console.log(
    `AI 쿼터초과: ${quotaExceeded}개`
  );

  console.log(
    `AI 오류: ${aiErrors}개`
  );

  console.log(
    '============================'
  );

  console.log(
    '보험신보 기사수집 + 선택적 AI 보완 완료'
  );
}

main().catch(
  error => {
    console.error(
      'ai_classify.js 실행 실패'
    );

    console.error(
      error
    );

    process.exit(1);
  }
);
