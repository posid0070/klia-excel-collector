const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUTPUT_DIR = path.join(ROOT, 'output');

const TARGET_MONTH = String(process.env.TARGET_MONTH || '').trim();

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
    mdash: '—',
    times: '×'
  };

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
   보험신보 페이지 번호 실제 클릭

   중요:
   - 매 페이지마다 먼저 1페이지 기본 주소로 들어간다.
   - ul.pagination 안에서 숫자 2 / 3 링크를 실제로 클릭한다.
   - 링크의 실제 href(page, total, 카테고리 값)를 그대로 사용한다.
   - 클릭 후 현재 페이지 번호를 확인한다.
========================================================= */

async function getInsweekCurrentPageNo(page) {

  const text = await page
    .locator(
      'ul.pagination[aria-label="Pagination"] ' +
      'li.pagination-item.current'
    )
    .first()
    .textContent()
    .catch(() => '');

  const n = Number(
    String(text || '').trim()
  );

  return Number.isFinite(n) ? n : 0;
}


async function clickInsweekPageNumber(
  page,
  pageNo
) {

  await page
    .locator('ul.pagination[aria-label="Pagination"]')
    .first()
    .waitFor({
      state: 'attached',
      timeout: 10000
    })
    .catch(() => {});

  /*
    페이지 숫자 텍스트가 아니라 실제 href의 page 파라미터를 기준으로 찾는다.

    예)
    <a href="...articleList.html?page=2&total=2120&sc_sub_section_code=S2N19&view_type=sm"
       class="pagination-link">2</a>
  */
  const links = page.locator(
    'ul.pagination[aria-label="Pagination"] ' +
    'a.pagination-link[href*="page="]'
  );

  const count = await links.count();

  console.log(
    `pagination page=${pageNo} 링크 탐색: 전체 ${count}개`
  );

  let targetLink = null;
  let targetHref = '';

  for (
    let i = 0;
    i < count;
    i++
  ) {

    const link = links.nth(i);

    const rawHref = await link
      .getAttribute('href')
      .catch(() => '');

    if (!rawHref) continue;

    try {

      const absoluteHref = new URL(
        String(rawHref)
          .replace(/&amp;/g, '&')
          .replace(/&#38;/g, '&'),
        page.url()
      ).href;

      const url = new URL(absoluteHref);

      const hrefPageNo = Number(
        url.searchParams.get('page')
      );

      const category =
        url.searchParams.get(
          'sc_sub_section_code'
        );

      const viewType =
        url.searchParams.get(
          'view_type'
        );

      if (
        hrefPageNo === pageNo &&
        category === 'S2N19' &&
        viewType === 'sm'
      ) {

        targetLink = link;
        targetHref = absoluteHref;
        break;
      }

    } catch (_) {}
  }

  if (!targetLink || !targetHref) {

    console.log(
      `pagination에서 page=${pageNo} 정확한 href를 찾지 못했습니다.`
    );

    return false;
  }

  console.log(
    `페이지 ${pageNo} 정확한 링크 발견: ${targetHref}`
  );

  /* 1순위: 보험신보가 제공한 실제 링크를 클릭 */
  try {

    await Promise.all([
      page
        .waitForNavigation({
          waitUntil: 'domcontentloaded',
          timeout: 15000
        })
        .catch(() => null),
      targetLink.click({
        force: true,
        timeout: 10000
      })
    ]);

    await page.waitForTimeout(800);

    const current =
      await getInsweekCurrentPageNo(page);

    const urlPageNo = (() => {
      try {
        return Number(
          new URL(page.url())
            .searchParams
            .get('page')
        );
      } catch (_) {
        return 0;
      }
    })();

    console.log(
      `페이지 ${pageNo} 클릭 후 | 현재페이지=${current || '확인불가'} | URL=${page.url()}`
    );

    if (
      current === pageNo ||
      urlPageNo === pageNo
    ) {
      return true;
    }

  } catch (e) {

    console.log(
      `페이지 ${pageNo} 실제 링크 클릭 실패: ${e.message}`
    );
  }

  /*
    2순위:
    클릭이 막혀도 페이지에서 읽은 실제 href를 그대로 사용한다.
    total 값도 사이트가 준 값을 그대로 유지한다.
  */
  try {

    console.log(
      `페이지 ${pageNo} 실제 href 직접 이동: ${targetHref}`
    );

    await page.goto(
      targetHref,
      {
        waitUntil: 'domcontentloaded',
        timeout: 30000,
        referer: insweekBaseUrl()
      }
    );

    await page.waitForTimeout(800);

    const current =
      await getInsweekCurrentPageNo(page);

    const urlPageNo = (() => {
      try {
        return Number(
          new URL(page.url())
            .searchParams
            .get('page')
        );
      } catch (_) {
        return 0;
      }
    })();

    console.log(
      `페이지 ${pageNo} href 이동 후 | 현재페이지=${current || '확인불가'} | URL=${page.url()}`
    );

    return (
      current === pageNo ||
      urlPageNo === pageNo
    );

  } catch (e) {

    console.log(
      `페이지 ${pageNo} 실제 href 이동 실패: ${e.message}`
    );

    return false;
  }
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

  let previousSignature = '';

  try {

    for (
      let pageNo = 1;
      pageNo <= 3;
      pageNo++
    ) {

      /*
        예전에 정상 동작했던 방식:
        매 페이지마다 1페이지 기본 주소로 다시 진입한 뒤
        2·3페이지는 pagination 숫자 링크를 실제 클릭한다.
      */
      await page.goto(
        insweekBaseUrl(),
        {
          waitUntil: 'domcontentloaded',
          timeout: 30000
        }
      );

      await page.waitForTimeout(900);

      if (pageNo > 1) {

        console.log(
          `보험신보 ${pageNo}페이지 이동 시도`
        );

        const clicked =
          await clickInsweekPageNumber(
            page,
            pageNo
          );

        if (!clicked) {
          throw new Error(
            `보험신보 ${pageNo}페이지 숫자 링크 이동에 실패했습니다.`
          );
        }
      }

      const currentPageNo =
        await getInsweekCurrentPageNo(page);

      console.log(
        `보험신보 ${pageNo}페이지 확인`
      );

      console.log(
        `현재 URL: ${page.url()}`
      );

      console.log(
        `pagination 현재 페이지: ${currentPageNo || '확인불가'}`
      );

      const raw =
        await extractCurrentList(
          page
        );

      console.log(
        `보험신보 ${pageNo}페이지 날짜+기사 추출: ${raw.length}건`
      );

      for (
        const item
        of raw.slice(0, 5)
      ) {
        console.log(
          `  ${item.dateText} | ${item.title}`
        );
      }

      /*
        URL만 바뀌고 실제 기사 목록은 1페이지 그대로인 경우를 방지한다.
      */
      const signature =
        raw
          .slice(0, 5)
          .map(
            x => x.url || x.title || ''
          )
          .join('|');

      if (
        pageNo > 1 &&
        signature &&
        signature === previousSignature
      ) {
        throw new Error(
          `보험신보 ${pageNo}페이지 기사 목록이 직전 페이지와 동일합니다.`
        );
      }

      if (signature) {
        previousSignature = signature;
      }

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
          clean(item.title);

        const url =
          item.url.split('#')[0];

        if (!map.has(url)) {
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
          a.articleDate.localeCompare(
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
   보험신보 5건 선정

   1~10일  : 1건
   11~20일 : 2건
   21~말일 : 2건
========================================================= */

function selectFiveArticles(
  articles
) {

  const bucket1 = articles.filter(article => {
    const day = Number(article.articleDate.slice(8, 10));
    return day >= 1 && day <= 10;
  });

  const bucket2 = articles.filter(article => {
    const day = Number(article.articleDate.slice(8, 10));
    return day >= 11 && day <= 20;
  });

  const bucket3 = articles.filter(article => {
    const day = Number(article.articleDate.slice(8, 10));
    return day >= 21;
  });

  const picked = [
    ...seededPick(bucket1, 1, `${TARGET_MONTH}|1-10`),
    ...seededPick(bucket2, 2, `${TARGET_MONTH}|11-20`),
    ...seededPick(bucket3, 2, `${TARGET_MONTH}|21-end`)
  ];

  const uniquePicked = [];
  const used = new Set();

  for (const item of picked) {
    if (!item || used.has(item.url)) {
      continue;
    }
    used.add(item.url);
    uniquePicked.push(item);
  }

  if (uniquePicked.length < 5) {
    const remaining = articles.filter(article => !used.has(article.url));

    uniquePicked.push(
      ...seededPick(
        remaining,
        5 - uniquePicked.length,
        `${TARGET_MONTH}|fallback`
      )
    );
  }

  return uniquePicked.slice(0, 5);
}


/* =========================================================
   생보협회 공시상품 카드
========================================================= */

function buildBenefitBlock(
  label,
  value
) {

  return `

<table
  width="100%"
  cellpadding="0"
  cellspacing="0"
  border="0"
  role="presentation"
  style="
    margin-top:8px;
    background:#FAFAFA;
    border:1px solid #EEEEEE;
    border-radius:7px;
  "
>

<tr>

<td style="
  padding:10px 12px;
">

  <p style="
    margin:0 0 4px 0;
    font-size:10px;
    font-weight:800;
    color:#FF5E1F;
    letter-spacing:0.2px;
  ">
    ${escapeHtml(label)}
  </p>

  <p style="
    margin:0;
    font-size:12px;
    color:#555555;
    line-height:1.75;
  ">
    ${escapeHtml(value || '-')}
  </p>

</td>

</tr>

</table>

`;

}


function buildDisclosureCards(
  samples
) {

  return samples

    .map(
      (
        item,
        index
      ) => {

        const variants = clean(
          item.상품형태 || ''
        )
          .split('|')
          .map(x => clean(x))
          .filter(Boolean);

        const variantHtml = variants.length
          ? `
            <p style="
              margin:0 0 10px 0;
              font-size:11px;
              color:#8A5A44;
              line-height:1.8;
            ">
              ${variants
                .map(v => `<span style="display:inline-block;background:#FFF3ED;border:1px solid #FFD8C8;border-radius:999px;padding:2px 7px;margin:0 4px 4px 0;">${escapeHtml(v)}</span>`)
                .join('')}
            </p>
          `
          : '';

        const isSavingCategory = [
          '연금',
          '저축',
          '교육'
        ].includes(clean(item.카테고리));

        const hasBenefit = Boolean(
          clean(item.급부명칭) ||
          clean(item.지급사유) ||
          clean(item.지급금액)
        );

        const benefitHtml = isSavingCategory
          ? ''
          : hasBenefit
            ? [
                buildBenefitBlock('급부명칭', item.급부명칭),
                buildBenefitBlock('지급사유', item.지급사유),
                buildBenefitBlock('지급금액', item.지급금액)
              ].join('')
            : `
              <table
                width="100%"
                cellpadding="0"
                cellspacing="0"
                border="0"
                role="presentation"
                style="margin-top:8px;background:#FAFAFA;border:1px solid #EEEEEE;border-radius:7px;"
              >
                <tr>
                  <td style="padding:10px 12px;font-size:12px;color:#888888;line-height:1.7;">
                    급부 상세 미표기
                  </td>
                </tr>
              </table>
            `;

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

            ${String(index + 1).padStart(2, '0')}

          </div>

        </td>


        <td style="vertical-align:top;">

          <p style="
            margin:0 0 5px 0;
            font-size:11px;
            font-weight:800;
            color:#FF5E1F;
          ">
            ${escapeHtml(item.카테고리)}
          </p>


          <p style="
            margin:0 0 4px 0;
            font-size:12px;
            font-weight:700;
            color:#777777;
            line-height:1.5;
          ">
            ${escapeHtml(item.보험회사명 || '')}
          </p>


          <p style="
            margin:0 0 8px 0;
            font-size:15px;
            font-weight:800;
            color:#222222;
            line-height:1.55;
          ">
            ${escapeHtml(
              item.정리상품명 ||
              item.상품명 ||
              '해당 월 판매개시 상품 없음'
            )}
          </p>

          ${variantHtml}

          ${benefitHtml}

        </td>

      </tr>

    </table>

  </td>

</tr>

`;

      }
    )

    .join('');

}


/* =========================================================
   보험신보 카드
========================================================= */

function buildNewsCards(
  articles
) {

  return articles

    .map(
      (
        article,
        index
      ) => `

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
            ${String(index + 1).padStart(2, '0')}
          </div>

        </td>


        <td style="vertical-align:top;">

          <p style="
            margin:0 0 9px 0;
            font-size:15px;
            font-weight:800;
            color:#222222;
            line-height:1.6;
          ">
            <a
              href="${escapeHtml(article.url)}"
              target="_blank"
              style="
                color:#222222;
                text-decoration:underline;
              "
            >
              ${escapeHtml(article.title)}
            </a>
          </p>


          <p style="
            margin:0;
            font-size:11px;
            color:#aaaaaa;
          ">
            보험신보 |
            ${escapeHtml(article.articleDate)}
          </p>

        </td>

      </tr>

    </table>

  </td>

</tr>

`
    )

    .join('');

}


/* =========================================================
   상품유형별 건수
========================================================= */

function buildCategoryCountRows(
  counts
) {

  const items = CATEGORY_ORDER.map(
    category => ({
      category,
      count: Number(counts?.[category] || 0)
    })
  );

  const rows = [];

  for (let i = 0; i < items.length; i += 3) {
    const group = items.slice(i, i + 3);

    while (group.length < 3) {
      group.push(null);
    }

    rows.push(`
      <tr>
        ${group.map((item, colIndex) => `
          <td style="
            padding:6px ${colIndex === 2 ? '0' : '10px'} 6px ${colIndex === 0 ? '0' : '10px'};
            width:33.33%;
            vertical-align:top;
          ">
            ${item
              ? `<span style="font-size:12px;color:#666666;white-space:nowrap;">${escapeHtml(item.category)}</span><br><strong style="font-size:12px;color:#222222;">${item.count}건</strong>`
              : ''}
          </td>
        `).join('')}
      </tr>
    `);
  }

  return rows.join('');
}


/* =========================================================
   뉴스레터 전체 HTML
========================================================= */

function buildHtml(
  meta,
  articles
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
      articles
    );

  const categoryCountRows =
    buildCategoryCountRows(
      meta.categoryProductCounts || {}
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

${monthNumber}월 생명보험협회 판매일자 기준 <strong>상품 ${meta.productCount}건</strong>이 확인되었습니다.

</p>

<p style="
  margin:10px 0 0 0;
  font-size:11px;
  color:#AAAAAA;
  line-height:1.7;
">
(동일 상품이 보험기간·급부 등에 따라 여러 행으로 구성된 경우 상품 수는 1건으로 계산했습니다.)
</p>

<table
  width="100%"
  cellpadding="0"
  cellspacing="0"
  border="0"
  role="presentation"
  style="margin-top:12px;border-top:1px solid #eeeeee;padding-top:8px;"
>

${categoryCountRows}

</table>


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

상품유형별 신상품

</h2>


<p style="
  font-size:12px;
  color:#999999;
  margin:0;
">

생명보험협회 13개 상품유형별로 각 1개 상품을 선정했습니다.

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

보험신보 신상품 5선

</h2>


<p style="
  font-size:12px;
  color:#999999;
  margin:0;
">

제목을 누르시면 바로 보험신보 신상품 기사 페이지로 이동합니다.

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
      5건 선정
    */

    selected =
      selectFiveArticles(
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
     AI 요약 미사용

     보험신보는 기사 제목 + 작성일 + 원문 링크만 표시
  ======================================================= */


  /* =======================================================
     HTML
  ======================================================= */

  const html =
    buildHtml(
      meta,
      selected
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
          selected

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
