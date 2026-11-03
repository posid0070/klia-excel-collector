const { chromium } = require('playwright');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const OUTPUT_DIR = path.join(__dirname, 'output');
const DB_FILE_NAME = '생보협회_상품DB.xlsx';
const UPLOAD_URL = process.env.UPLOAD_URL;
const UPLOAD_SECRET = process.env.UPLOAD_SECRET;
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
const GEMINI_MODEL = process.env.GEMINI_MODEL || 'gemini-3.5-flash-lite';
const TARGET_MONTH_OVERRIDE = String(process.env.TARGET_MONTH || '').trim();

const INSWEEK_LIST_URL =
  'https://www.insweek.co.kr/news/articleList.html?page=1&total=2120&sc_section_code=&sc_sub_section_code=S2N19&sc_serial_code=&sc_area=&sc_level=&sc_article_type=&sc_view_level=&sc_sdate=&sc_edate=&sc_serial_number=&sc_word=&sc_andor=&sc_word2=&box_idxno=&sc_multi_code=&sc_is_image=&sc_is_movie=&sc_user_name=&sc_order_by=E';

const MAX_INSWEEK_PAGES = 20;
const MAX_SOURCE_TEXT = 5500;

const AI_TARGET_COMPANIES = [
  '삼성생명',
  '한화생명',
  '교보생명',
  '신한라이프',
  'NH농협생명'
];

const COMPANY_ALIASES = {
  삼성생명: ['삼성생명'],
  한화생명: ['한화생명'],
  교보생명: ['교보생명'],
  신한라이프: ['신한라이프', '신한생명'],
  NH농협생명: ['NH농협생명', '농협생명']
};

const OFFICIAL_DOMAINS = {
  삼성생명: ['samsunglife.com'],
  한화생명: ['hanwhalife.com'],
  교보생명: ['news.kyobo.com', 'kyobo.com'],
  신한라이프: ['shinhanlife.co.kr'],
  NH농협생명: ['nhlife.co.kr']
};

const clean = v =>
  String(v || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

const norm = v =>
  clean(v)
    .toLowerCase()
    .replace(/[\s\(\)\[\]\{\}·ㆍ\-_,.\/\\:：'"“”‘’]/g, '');

const sleep = ms =>
  new Promise(resolve => setTimeout(resolve, ms));

const pad2 = n =>
  String(n).padStart(2, '0');

const makeDate = (y, m, d) => {
  y = Number(y);
  m = Number(m);
  d = Number(d);

  return y &&
    m >= 1 &&
    m <= 12 &&
    d >= 1 &&
    d <= 31
    ? `${y}-${pad2(m)}-${pad2(d)}`
    : '';
};

function getTargetMonth() {
  if (TARGET_MONTH_OVERRIDE) {
    if (!/^\d{4}-\d{2}$/.test(TARGET_MONTH_OVERRIDE)) {
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
        .map(p => [p.type, p.value])
    );

  let year =
    Number(parts.year);

  let month =
    Number(parts.month) - 1;

  if (month === 0) {
    year -= 1;
    month = 12;
  }

  return `${year}-${pad2(month)}`;
}

function canonicalCompany(value) {
  const s = norm(value);

  for (
    const [company, aliases]
    of Object.entries(
      COMPANY_ALIASES
    )
  ) {
    if (
      aliases.some(
        alias =>
          s.includes(
            norm(alias)
          )
      )
    ) {
      return company;
    }
  }

  return clean(value);
}

function detectTargetCompany(text) {
  const s = norm(text);

  for (
    const company
    of AI_TARGET_COMPANIES
  ) {
    const aliases =
      COMPANY_ALIASES[
        company
      ] || [company];

    if (
      aliases.some(
        alias =>
          s.includes(
            norm(alias)
          )
      )
    ) {
      return company;
    }
  }

  return '';
}

const isAiTargetCompany =
  value =>
    AI_TARGET_COMPANIES
      .includes(
        canonicalCompany(
          value
        )
      );

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
    workbook.Sheets[
      sheetName
    ];

  return sheet
    ? XLSX.utils
        .sheet_to_json(
          sheet,
          {
            defval: ''
          }
        )
    : [];
}

function replaceSheet(
  workbook,
  sheetName,
  rows
) {
  const index =
    workbook
      .SheetNames
      .indexOf(
        sheetName
      );

  if (index >= 0) {
    workbook
      .SheetNames
      .splice(
        index,
        1
      );

    delete workbook
      .Sheets[
        sheetName
      ];
  }

  const data =
    rows.length
      ? rows
      : [
          {
            안내:
              '해당 데이터 없음'
          }
        ];

  const sheet =
    XLSX.utils
      .json_to_sheet(
        data
      );

  if (rows.length) {
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
                headers.length - 1
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

function parseListDate(
  text,
  targetMonth
) {
  const s =
    clean(text);

  const year =
    Number(
      targetMonth
        .slice(
          0,
          4
        )
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
      /(?:^|\s)(\d{1,2})[-.](\d{1,2})(?:\s+\d{1,2}:\d{2})?/
    );

  if (match) {
    return makeDate(
      year,
      match[1],
      match[2]
    );
  }

  return '';
}

function extractPageDate(text) {
  const match =
    clean(text)
      .match(
        /(20\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/
      );

  return match
    ? makeDate(
        match[1],
        match[2],
        match[3]
      )
    : '';
}

function extractLaunchDateInfo(
  text,
  referenceDate
) {
  const sentences =
    String(
      text || ''
    )
      .replace(
        /\r/g,
        '\n'
      )
      .split(
        /(?<=[.!?다요])\s+|\n+/
      )
      .map(clean)
      .filter(Boolean);

  const keyword =
    /(출시|판매를?\s*시작|판매개시|판매\s*개시|개정\s*출시|선보|판매한다|출시한다|출시했다)/;

  const candidates = [];

  for (
    const sentence
    of sentences
  ) {
    if (
      !keyword.test(
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
      referenceDate
    ) {
      candidates.push({
        date:
          makeDate(
            referenceDate.slice(
              0,
              4
            ),
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
        /(?:지난\s*|오는\s*|이달\s*)?(\d{1,2})일(?:부터|에)?/
      );

    if (
      match &&
      referenceDate
    ) {
      candidates.push({
        date:
          makeDate(
            referenceDate.slice(
              0,
              4
            ),
            referenceDate.slice(
              5,
              7
            ),
            match[1]
          ),
        sentence,
        score: 70
      });
    }
  }

  candidates.sort(
    (a, b) =>
      b.score -
      a.score
  );

  return (
    candidates[0] ||
    {
      date: '',
      sentence: ''
    }
  );
}

function insweekPageUrl(
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

async function getInsweekCards(
  page,
  targetMonth
) {
  const raw =
    await page
      .locator(
        'a[href*="articleView.html"]'
      )
      .evaluateAll(
        nodes =>
          nodes.map(
            a => {
              let parent =
                a;

              let context = '';

              for (
                let i = 0;
                i < 7;
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

                if (
                  /\b\d{1,2}[-.]\d{1,2}\s+\d{1,2}:\d{2}\b/
                    .test(
                      text
                    )
                ) {
                  context =
                    text;

                  break;
                }
              }

              return {
                url:
                  a.href || '',
                title:
                  (
                    a.innerText ||
                    a.textContent ||
                    ''
                  ).trim(),
                context
              };
            }
          )
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
      parseListDate(
        item.context,
        targetMonth
      );

    if (!articleDate) {
      continue;
    }

    const url =
      item.url
        .split('#')[0];

    if (
      !map.has(url) ||
      clean(
        item.title
      ).length >
      clean(
        map.get(
          url
        ).title
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

async function collectInsweek(
  browser,
  targetMonth
) {
  const page =
    await browser.newPage();

  const start =
    `${targetMonth}-01`;

  const end =
    `${targetMonth}-31`;

  const map =
    new Map();

  try {
    for (
      let pageNo = 1;
      pageNo <=
      MAX_INSWEEK_PAGES;
      pageNo++
    ) {
      console.log(
        `보험신보 신상품 ${pageNo}페이지 확인`
      );

      try {
        await page.goto(
          insweekPageUrl(
            pageNo
          ),
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

      } catch (error) {
        console.log(
          `보험신보 목록 실패: ${error.message}`
        );

        continue;
      }

      const cards =
        await getInsweekCards(
          page,
          targetMonth
        );

      if (!cards.length) {
        continue;
      }

      cards
        .filter(
          card =>
            card.articleDate >=
              start &&
            card.articleDate <=
              end
        )
        .forEach(
          card =>
            map.set(
              card.url,
              card
            )
        );

      if (
        cards.every(
          card =>
            card.articleDate <
            start
        )
      ) {
        break;
      }

      if (
        cards.some(
          card =>
            card.articleDate >=
              start &&
            card.articleDate <=
              end
        ) &&
        cards.some(
          card =>
            card.articleDate <
            start
        )
      ) {
        break;
      }
    }

  } finally {
    await page.close();
  }

  return [
    ...map.values()
  ].sort(
    (a, b) =>
      b.articleDate
        .localeCompare(
          a.articleDate
        )
  );
}

async function readPage(
  page,
  url,
  knownTitle = ''
) {
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
      400
    );

  } catch (error) {
    return {
      url,
      title:
        knownTitle,
      body: '',
      pageDate: '',
      launchDate: '',
      launchSentence: '',
      error:
        error.message
    };
  }

  const data =
    await page.evaluate(
      () => {
        const pick =
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

        return {
          title:
            pick([
              'h1',
              '.article-head-title',
              '.article-view-head-title',
              '.article-title'
            ]),

          body:
            pick([
              '[itemprop="articleBody"]',
              '#article-view-content-div',
              '.article-view-content-div',
              '.article-body',
              '.article-content',
              'article',
              'main'
            ]) ||
            (
              document.body &&
              document.body.innerText
            ) ||
            ''
        };
      }
    );

  const title =
    clean(
      data.title ||
      knownTitle
    );

  const body =
    clean(
      data.body
    ).slice(
      0,
      18000
    );

  const pageDate =
    extractPageDate(
      body
    );

  const launch =
    extractLaunchDateInfo(
      body,
      pageDate
    );

  return {
    url,
    title,
    body,
    pageDate,
    launchDate:
      launch.date,
    launchSentence:
      launch.sentence,
    error: ''
  };
}

async function enrichInsweek(
  browser,
  items
) {
  const page =
    await browser.newPage();

  const results = [];

  try {
    for (
      const item
      of items
    ) {
      console.log(
        `보험신보 본문: ${item.title}`
      );

      const detail =
        await readPage(
          page,
          item.url,
          item.title
        );

      const launch =
        extractLaunchDateInfo(
          detail.body,
          item.articleDate ||
          detail.pageDate
        );

      results.push({
        ...item,

        title:
          detail.title ||
          item.title,

        body:
          detail.body,

        company:
          detectTargetCompany(
            `${detail.title} ${detail.body}`
          ),

        launchDate:
          launch.date,

        launchSentence:
          launch.sentence,

        error:
          detail.error
      });
    }

  } finally {
    await page.close();
  }

  return results;
}

function levenshtein(
  a,
  b
) {
  const s =
    norm(a);

  const t =
    norm(b);

  const dp =
    Array.from(
      {
        length:
          s.length + 1
      },
      () =>
        Array(
          t.length + 1
        ).fill(0)
    );

  for (
    let i = 0;
    i <= s.length;
    i++
  ) {
    dp[i][0] = i;
  }

  for (
    let j = 0;
    j <= t.length;
    j++
  ) {
    dp[0][j] = j;
  }

  for (
    let i = 1;
    i <= s.length;
    i++
  ) {
    for (
      let j = 1;
      j <= t.length;
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
    s.length
  ][
    t.length
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

function matchScore(
  target,
  article
) {
  if (
    canonicalCompany(
      target[
        '보험회사명'
      ]
    ) !==
    canonicalCompany(
      article.company
    )
  ) {
    return 0;
  }

  const family =
    clean(
      target[
        '상품군명'
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
    norm(title)
      .includes(
        norm(family)
      )
  ) {
    score += 65;

  } else if (
    norm(family) &&
    norm(body)
      .includes(
        norm(family)
      )
  ) {
    score += 40;
  }

  score +=
    Math.round(
      similarity(
        family,
        title
      ) *
      45
    );

  return score;
}

function bestInsweek(
  target,
  articles
) {
  const ranked =
    articles
      .map(
        article => ({
          article,
          score:
            matchScore(
              target,
              article
            )
        })
      )
      .filter(
        item =>
          item.score >= 55
      )
      .sort(
        (a, b) =>
          b.score -
          a.score
      );

  return ranked[0] || null;
}

async function bingSearch(
  page,
  query,
  domain
) {
  const q =
    `${query} site:${domain}`;

  const url =
    `https://www.bing.com/search?q=${encodeURIComponent(q)}`;

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
      500
    );

  } catch {
    return [];
  }

  return page
    .locator(
      'li.b_algo h2 a, #b_results h2 a'
    )
    .evaluateAll(
      nodes =>
        nodes
          .slice(
            0,
            6
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
                a.href || ''
            })
          )
    );
}

function supplementQueries(
  target,
  targetMonth
) {
  const company =
    canonicalCompany(
      target[
        '보험회사명'
      ]
    );

  const product =
    clean(
      target[
        '상품군명'
      ]
    );

  const year =
    targetMonth.slice(
      0,
      4
    );

  const month =
    Number(
      targetMonth.slice(
        5,
        7
      )
    );

  const base =
    `"${company}" "${product}" ${year} ${month}월`;

  return [
    {
      type:
        '보험저널',

      query:
        base,

      domains: [
        'insurancejournal.co.kr'
      ]
    },

    {
      type:
        '보험사공식',

      query:
        base,

      domains:
        OFFICIAL_DOMAINS[
          company
        ] || []
    }
  ];
}

async function collectSupplements(
  browser,
  target,
  targetMonth
) {
  const searchPage =
    await browser.newPage();

  const readPageTab =
    await browser.newPage();

  const results = [];

  try {
    for (
      const queryInfo
      of supplementQueries(
        target,
        targetMonth
      )
    ) {
      for (
        const domain
        of queryInfo.domains
      ) {
        console.log(
          `추가검색: ${canonicalCompany(
            target['보험회사명']
          )} / ${queryInfo.type}`
        );

        const hits =
          await bingSearch(
            searchPage,
            queryInfo.query,
            domain
          );

        let kept = 0;

        for (
          const hit
          of hits
        ) {
          if (
            kept >= 2
          ) {
            break;
          }

          if (
            !hit.url.includes(
              domain
            )
          ) {
            continue;
          }

          const detail =
            await readPage(
              readPageTab,
              hit.url,
              hit.title
            );

          if (
            !detail.body
          ) {
            continue;
          }

          const score =
            (
              norm(
                target[
                  '상품군명'
                ]
              ) &&
              norm(
                detail.title
              ).includes(
                norm(
                  target[
                    '상품군명'
                  ]
                )
              )
                ? 60
                : 0
            ) +
            Math.round(
              similarity(
                target[
                  '상품군명'
                ],
                detail.title
              ) *
              40
            );

          if (
            score < 45
          ) {
            continue;
          }

          results.push({
            sourceType:
              queryInfo.type,

            domain,

            title:
              detail.title,

            url:
              hit.url,

            body:
              detail.body.slice(
                0,
                MAX_SOURCE_TEXT
              ),

            pageDate:
              detail.pageDate,

            launchDate:
              detail.launchDate,

            launchSentence:
              detail.launchSentence,

            score
          });

          kept += 1;
        }
      }
    }

  } finally {
    await searchPage.close();
    await readPageTab.close();
  }

  const map =
    new Map();

  for (
    const item
    of results
  ) {
    if (
      !map.has(
        item.url
      ) ||
      item.score >
      map.get(
        item.url
      ).score
    ) {
      map.set(
        item.url,
        item
      );
    }
  }

  return [
    ...map.values()
  ]
    .sort(
      (a, b) =>
        b.score -
        a.score
    )
    .slice(
      0,
      4
    );
}

function extractOutputText(
  data
) {
  if (
    typeof
      data.output_text ===
      'string' &&
    data.output_text.trim()
  ) {
    return data
      .output_text
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
      of step.content ||
      []
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

function parseJsonText(
  text
) {
  let s =
    String(
      text || ''
    )
      .trim()
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

async function geminiNoSearch(
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
              prompt
          })
      }
    );

  const text =
    await response.text();

  if (
    !response.ok
  ) {
    const error =
      new Error(
        `Gemini HTTP ${response.status}: ${text}`
      );

    error.status =
      response.status;

    throw error;
  }

  return parseJsonText(
    extractOutputText(
      JSON.parse(
        text
      )
    )
  );
}

function sourcePacket(
  insweek,
  supplements
) {
  const result = [];

  if (insweek) {
    result.push(
      `[보험신보]
기사작성일: ${insweek.articleDate || ''}
기사제목: ${insweek.title || ''}
출시일: ${insweek.launchDate || '미확인'}
출시근거: ${insweek.launchSentence || '미확인'}
URL: ${insweek.url || ''}
본문: ${clean(insweek.body).slice(0, MAX_SOURCE_TEXT)}`
    );
  }

  supplements.forEach(
    (source, index) => {
      result.push(
        `[추가자료 ${index + 1}-${source.sourceType}]
제목: ${source.title}
페이지날짜: ${source.pageDate || '미확인'}
출시일: ${source.launchDate || '미확인'}
출시근거: ${source.launchSentence || '미확인'}
URL: ${source.url}
본문: ${clean(source.body).slice(0, MAX_SOURCE_TEXT)}`
      );
    }
  );

  return result.join(
    '\n\n'
  );
}

async function callGemini(
  target,
  insweek,
  supplements,
  targetMonth
) {
  const prompt =
`당신은 한국 생명보험 상품동향 조사 담당자입니다.
인터넷 검색은 하지 말고 아래 수집자료만 근거로 판단하세요.

[대상상품]
조사월: ${targetMonth}
보험회사: ${canonicalCompany(target['보험회사명'])}
상품유형: ${target['상품유형'] || ''}
상품군명: ${target['상품군명'] || ''}
세부상품명: ${target['세부상품명'] || ''}
협회 판매일자: ${target['판매일자'] || ''}
규칙분류: ${target['규칙분류'] || ''}
신규급부 후보: ${target['신규급부목록'] || ''}

[수집자료]
${sourcePacket(insweek, supplements) || '관련 외부자료 없음'}

규칙:
- 제공자료 밖의 사실을 추측하지 마세요.
- 신규상품 / 신규급부·특약 / 상품개정은 근거가 있을 때만 판정하세요.
- 근거가 약하면 자료확인불가로 표시하세요.
- 기사작성일과 실제 출시일을 구분하세요.
- 주요특징은 2문장 이내로 작성하세요.

반드시 JSON 하나만 출력하세요.

{
  "AI검증결과":"신규상품|신규급부·특약|상품개정|자료확인불가",
  "AI신뢰도":"높음|보통|낮음",
  "AI주요특징":"",
  "AI신규급부특약":"",
  "AI확인근거":"",
  "AI출시일검증":"일치|일자불일치|출시일미확인"
}`;

  try {
    return await geminiNoSearch(
      prompt
    );

  } catch (error) {
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

      return geminiNoSearch(
        prompt
      );
    }

    throw error;
  }
}

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

    if (result) {
      Object.assign(
        row,
        result
      );
    }
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

  const targets =
    familyRows.filter(
      row =>
        row['AI대상여부'] ===
          'Y' &&
        isAiTargetCompany(
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
    `조사월: ${targetMonth}`
  );

  console.log(
    `전체 상품군: ${familyRows.length}개`
  );

  console.log(
    `AI 대상: ${targets.length}개`
  );

  const browser =
    await chromium.launch({
      headless: true
    });

  try {
    const list =
      await collectInsweek(
        browser,
        targetMonth
      );

    const insweek =
      await enrichInsweek(
        browser,
        list
      );

    console.log(
      `보험신보 ${targetMonth} 작성기사: ${insweek.length}건`
    );

    replaceSheet(
      workbook,
      '보험신보기사',
      insweek.map(
        article => {
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
              canonicalCompany(
                article.company
              ),

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
              article.error
                ? `본문수집오류: ${article.error}`
                : '정상'
          };
        }
      )
    );

    const resultMap =
      new Map();

    const supplementRows =
      [];

    let matchedCount = 0;
    let supplementCount = 0;
    let aiCompleted = 0;
    let quotaCount = 0;
    let errorCount = 0;

    for (
      const target
      of targets
    ) {
      const familyKey =
        makeFamilyKey(
          target
        );

      const ranked =
        bestInsweek(
          target,
          insweek
        );

      const matched =
        ranked
          ? ranked.article
          : null;

      if (matched) {
        matchedCount += 1;
      }

      const supplements =
        await collectSupplements(
          browser,
          target,
          targetMonth
        );

      if (
        supplements.length
      ) {
        supplementCount += 1;
      }

      supplements.forEach(
        source =>
          supplementRows.push({
            조사월:
              targetMonth,

            보험회사명:
              canonicalCompany(
                target[
                  '보험회사명'
                ]
              ),

            상품군명:
              target[
                '상품군명'
              ] || '',

            출처구분:
              source.sourceType,

            도메인:
              source.domain,

            추가검색제목:
              source.title,

            페이지날짜:
              source.pageDate ||
              '',

            출시일:
              source.launchDate ||
              '',

            추가검색근거:
              source.launchSentence ||
              '',

            매칭점수:
              source.score,

            추가검색URL:
              source.url
          })
      );

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

        추가검색여부:
          supplements.length
            ? 'Y'
            : 'N',

        추가검색출처:
          [
            ...new Set(
              supplements.map(
                x =>
                  x.sourceType
              )
            )
          ].join(' | '),

        추가검색제목:
          supplements
            .map(
              x =>
                x.title
            )
            .join(' | '),

        추가검색URL:
          supplements
            .map(
              x =>
                x.url
            )
            .join('\n'),

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
        AI출처URL: ''
      };

      if (
        !GEMINI_API_KEY
      ) {
        result.AI보완상태 =
          matched ||
          supplements.length
            ? '자료수집완료·AI미실행'
            : '외부자료미확인·AI미실행';

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
        `AI 요약: ${canonicalCompany(
          target['보험회사명']
        )} / ${target['상품군명']}`
      );

      try {
        const ai =
          await callGemini(
            target,
            matched,
            supplements,
            targetMonth
          );

        const urls =
          new Set();

        if (
          matched &&
          matched.url
        ) {
          urls.add(
            matched.url
          );
        }

        supplements.forEach(
          item => {
            if (item.url) {
              urls.add(
                item.url
              );
            }
          }
        );

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
              ...urls
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
          result.AI보완상태 =
            '자료수집완료·AI쿼터초과';

          result.AI확인근거 =
            'Gemini 일반 API 사용량 한도 초과. 직접 수집 자료는 유지함.';

          quotaCount += 1;

        } else {
          result.AI보완상태 =
            '자료수집완료·AI오류';

          result.AI확인근거 =
            String(
              error.message ||
              error
            ).slice(
              0,
              600
            );

          errorCount += 1;
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
        2500
      );
    }

    replaceSheet(
      workbook,
      '추가검색자료',
      supplementRows
    );

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
        ] !== 'Y'
      ) {
        row[
          'AI보완상태'
        ] =
          '비대상';
      }
    }

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
        ] !== 'Y'
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

    replaceSheet(
      workbook,
      '상품목록',
      productRows
    );

    replaceSheet(
      workbook,
      'AI분석대상',
      familyRows
        .filter(
          row =>
            row[
              'AI대상여부'
            ] === 'Y'
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

            보험신보URL:
              row[
                '보험신보URL'
              ] || '',

            추가검색여부:
              row[
                '추가검색여부'
              ] || '',

            추가검색출처:
              row[
                '추가검색출처'
              ] || '',

            추가검색제목:
              row[
                '추가검색제목'
              ] || '',

            추가검색URL:
              row[
                '추가검색URL'
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
        )
    );

    replaceSheet(
      workbook,
      'AI분류요약',
      [
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
            insweek.length
        },
        {
          구분:
            'AI대상 중 보험신보 기사매칭',
          건수:
            matchedCount
        },
        {
          구분:
            '보험저널/보험사 공식자료 추가확인 상품',
          건수:
            supplementCount
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
            quotaCount
        },
        {
          구분:
            'AI 오류',
          건수:
            errorCount
        }
      ]
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

    console.log(
      `완료 - AI대상 ${targets.length}, 보험신보매칭 ${matchedCount}, 추가자료 ${supplementCount}, AI완료 ${aiCompleted}, 쿼터초과 ${quotaCount}, 오류 ${errorCount}`
    );

  } finally {
    await browser.close();
  }
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
