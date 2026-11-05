const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUTPUT_DIR = path.join(ROOT, 'output');

const TARGET_MONTH = String(process.env.TARGET_MONTH || '').trim();
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || '';
const GEMINI_MODEL = process.env.GEMINI_MODEL || 'gemini-3.5-flash-lite';

const V5_MAIL_URL = process.env.V5_MAIL_URL || '';
const V5_MAIL_SECRET = process.env.V5_MAIL_SECRET || '';
const SEND_MAIL = String(process.env.SEND_MAIL || 'false').toLowerCase() === 'true';

if (!/^\d{4}-\d{2}$/.test(TARGET_MONTH)) {
  throw new Error('TARGET_MONTH는 YYYY-MM 형식이어야 합니다.');
}

const META_FILE = path.join(OUTPUT_DIR, `${TARGET_MONTH}_V5_meta.json`);
const EXCEL_FILE = path.join(OUTPUT_DIR, `${TARGET_MONTH}_V5_생보협회_상품통합.xlsx`);
const HTML_FILE = path.join(OUTPUT_DIR, `${TARGET_MONTH}_V5_민영사_신상품동향.html`);
const JSON_FILE = path.join(OUTPUT_DIR, `${TARGET_MONTH}_V5_newsletter.json`);

const LIFE_COMPANY_KEYWORDS = [
  '삼성생명', '한화생명', '교보생명', '신한라이프', 'NH농협생명', '농협생명',
  '동양생명', 'ABL생명', '흥국생명', '미래에셋생명', '메트라이프생명', '메트라이프',
  'KB라이프', 'KDB생명', '푸본현대생명', '라이나생명', 'AIA생명', '처브라이프',
  '하나생명', 'DB생명', 'BNP파리바카디프생명', '카디프생명', 'IBK연금보험',
  '교보라이프플래닛'
];

const EXCLUDE_KEYWORDS = [
  '손해보험', '화재', '손보', '변액', '퇴직연금', '연금저축', '실손'
];

function clean(v) {
  return String(v ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function escapeHtml(v) {
  return clean(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function makeSeed(str) {
  let h = 2166136261;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h >>> 0;
}

function mulberry32(seed) {
  return function () {
    let t = seed += 0x6D2B79F5;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function seededPick(items, count, seedText) {
  const pool = [...items];
  const rng = mulberry32(makeSeed(seedText));
  const picked = [];

  while (pool.length && picked.length < count) {
    const i = Math.floor(rng() * pool.length);
    picked.push(pool.splice(i, 1)[0]);
  }

  return picked;
}

function parseInsweekDate(text) {
  const m = clean(text).match(/(\d{1,2})-(\d{1,2})\s+\d{1,2}:\d{2}/);
  if (!m) return '';
  const year = TARGET_MONTH.slice(0, 4);
  return `${year}-${String(Number(m[1])).padStart(2, '0')}-${String(Number(m[2])).padStart(2, '0')}`;
}

function insweekPageUrl(pageNo) {
  return `https://www.insweek.co.kr/news/articleList.html?page=${pageNo}&sc_sub_section_code=S2N19&view_type=sm`;
}

async function collectInsweekList(browser) {
  const page = await browser.newPage();
  const start = `${TARGET_MONTH}-01`;
  const end = `${TARGET_MONTH}-31`;
  const map = new Map();

  try {
    for (let pageNo = 1; pageNo <= 3; pageNo++) {
      console.log(`보험신보 ${pageNo}페이지 확인`);

      await page.goto(insweekPageUrl(pageNo), {
        waitUntil: 'domcontentloaded',
        timeout: 30000
      });

      await page.waitForTimeout(700);

      const raw = await page
        .locator('a[href*="articleView.html"]')
        .evaluateAll(nodes => {
          const out = [];

          for (const a of nodes) {
            let parent = a;
            let dateText = '';

            for (let i = 0; i < 10; i++) {
              if (!parent) break;

              const els = parent.querySelectorAll('.altlist-info-item');

              for (const el of els) {
                const t = (el.textContent || '').trim();
                if (/\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}/.test(t)) {
                  dateText = t;
                  break;
                }
              }

              if (dateText) break;
              parent = parent.parentElement;
            }

            out.push({
              title: (a.innerText || a.textContent || '').trim(),
              url: a.href || '',
              dateText
            });
          }

          return out;
        });

      for (const item of raw) {
        if (!item.url || !item.title || !item.dateText) continue;

        const articleDate = parseInsweekDate(item.dateText);
        if (!articleDate || articleDate < start || articleDate > end) continue;

        const title = clean(item.title);

        if (EXCLUDE_KEYWORDS.some(k => title.includes(k))) continue;
        if (!LIFE_COMPANY_KEYWORDS.some(k => title.includes(k))) continue;

        const url = item.url.split('#')[0];

        if (!map.has(url) || title.length > map.get(url).title.length) {
          map.set(url, { title, url, articleDate });
        }
      }
    }
  } finally {
    await page.close();
  }

  return [...map.values()].sort((a, b) => a.articleDate.localeCompare(b.articleDate));
}

function selectFourArticles(articles) {
  const b1 = articles.filter(a => Number(a.articleDate.slice(8, 10)) <= 10);
  const b2 = articles.filter(a => {
    const d = Number(a.articleDate.slice(8, 10));
    return d >= 11 && d <= 20;
  });
  const b3 = articles.filter(a => Number(a.articleDate.slice(8, 10)) >= 21);

  const picked = [
    ...seededPick(b1, 1, `${TARGET_MONTH}|1-10`),
    ...seededPick(b2, 2, `${TARGET_MONTH}|11-20`),
    ...seededPick(b3, 1, `${TARGET_MONTH}|21-end`)
  ];

  const used = new Set(picked.map(x => x.url));

  if (picked.length < 4) {
    const remaining = articles.filter(x => !used.has(x.url));
    picked.push(...seededPick(remaining, 4 - picked.length, `${TARGET_MONTH}|fallback`));
  }

  return picked.slice(0, 4);
}

async function enrichArticles(browser, articles) {
  const page = await browser.newPage();
  const out = [];

  try {
    for (const article of articles) {
      let body = '';

      try {
        await page.goto(article.url, {
          waitUntil: 'domcontentloaded',
          timeout: 30000
        });

        await page.waitForTimeout(400);

        body = await page.evaluate(() => {
          const selectors = [
            '[itemprop="articleBody"]',
            '#article-view-content-div',
            '.article-view-content-div',
            '.article-body',
            '.article-content',
            'article',
            'main'
          ];

          for (const selector of selectors) {
            const el = document.querySelector(selector);
            const text = el && (el.innerText || el.textContent || '').trim();
            if (text) return text;
          }

          return (document.body && document.body.innerText) || '';
        });

      } catch (e) {
        console.log(`보험신보 본문 실패: ${article.title} / ${e.message}`);
      }

      out.push({ ...article, body: clean(body).slice(0, 9000) });
    }
  } finally {
    await page.close();
  }

  return out;
}

function extractGeminiText(data) {
  const parts = data?.candidates?.[0]?.content?.parts || [];
  return parts.map(p => p.text || '').join('\n').trim();
}

function parseJsonArray(text) {
  const s = String(text || '')
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/, '')
    .replace(/\s*```$/, '')
    .trim();

  const first = s.indexOf('[');
  const last = s.lastIndexOf(']');

  if (first < 0 || last < first) {
    throw new Error('Gemini JSON 배열을 찾지 못했습니다.');
  }

  return JSON.parse(s.slice(first, last + 1));
}

async function summarizeWithGemini(articles) {
  if (!GEMINI_API_KEY) {
    return articles.map((a, i) => ({
      index: i + 1,
      company: '',
      product: a.title,
      feature: 'Gemini API 키가 없어 자동 요약을 실행하지 않았습니다.',
      implication: '기사 원문을 참고해주세요.'
    }));
  }

  const source = articles
    .map((a, i) => `[기사 ${i + 1}]
기사작성일: ${a.articleDate}
제목: ${a.title}
URL: ${a.url}
본문:
${a.body}`)
    .join('\n\n');

  const prompt = `당신은 우체국보험 상품개발 참고용 월간 뉴스레터 편집자입니다.
아래 보험신보 생명보험 신상품 기사들을 기사 내용만 근거로 요약하세요.

규칙:
- 인터넷 검색 금지
- 기사에 없는 내용 추정 금지
- 광고성 표현 제거
- feature: 상품 핵심 특징 한 문장
- implication: 특정 우체국보험 상품과 사실 비교하지 말고, 상품개발 관점 시사점 한 문장
- company: 기사에 나온 보험회사명
- product: 기사에 나온 실제 상품명

JSON 배열만 출력:
[
  {
    "index": 1,
    "company": "",
    "product": "",
    "feature": "",
    "implication": ""
  }
]

${source}`;

  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(GEMINI_MODEL)}:generateContent`;

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'x-goog-api-key': GEMINI_API_KEY,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      contents: [
        {
          role: 'user',
          parts: [{ text: prompt }]
        }
      ],
      generationConfig: {
        temperature: 0.2,
        responseMimeType: 'application/json'
      }
    })
  });

  const text = await response.text();

  if (!response.ok) {
    throw new Error(`Gemini HTTP ${response.status}: ${text}`);
  }

  return parseJsonArray(extractGeminiText(JSON.parse(text)));
}

function buildDisclosureCards(samples) {
  return samples.map((x, i) => `
<tr>
  <td style="padding:18px 28px;border-bottom:1px solid #eeeeee;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
      <tr>
        <td style="width:42px;vertical-align:top;">
          <div style="width:30px;height:30px;line-height:30px;text-align:center;border-radius:50%;background:#FFF0EA;color:#FF5E1F;font-size:12px;font-weight:800;">
            ${String(i + 1).padStart(2, '0')}
          </div>
        </td>
        <td style="vertical-align:top;">
          <p style="margin:0 0 4px 0;font-size:11px;font-weight:700;color:#FF5E1F;">
            ${escapeHtml(x.카테고리)}
          </p>
          <p style="margin:0 0 8px 0;font-size:15px;font-weight:700;color:#222222;line-height:1.5;">
            ${escapeHtml(x.보험회사명)} · ${escapeHtml(x.상품명)}
          </p>
          <p style="margin:3px 0;font-size:12px;color:#666666;line-height:1.7;"><strong>급부명칭</strong> ${escapeHtml(x.급부명칭 || '-')}</p>
          <p style="margin:3px 0;font-size:12px;color:#666666;line-height:1.7;"><strong>지급사유</strong> ${escapeHtml(x.지급사유 || '-')}</p>
          <p style="margin:3px 0;font-size:12px;color:#666666;line-height:1.7;"><strong>지급금액</strong> ${escapeHtml(x.지급금액 || '-')}</p>
        </td>
      </tr>
    </table>
  </td>
</tr>`).join('');
}

function buildNewsCards(articles, summaries) {
  return articles.map((a, i) => {
    const s = summaries.find(x => Number(x.index) === i + 1) || summaries[i] || {};
    const d = Number(a.articleDate.slice(8, 10));

    let bucket = '21일 ~ 말일';
    if (d <= 10) bucket = '1일 ~ 10일';
    else if (d <= 20) bucket = '11일 ~ 20일';

    return `
<tr>
  <td style="padding:20px 28px;border-bottom:1px solid #eeeeee;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
      <tr>
        <td style="width:42px;vertical-align:top;">
          <div style="width:30px;height:30px;line-height:30px;text-align:center;border-radius:50%;background:#FFF0EA;color:#FF5E1F;font-size:12px;font-weight:800;">
            ${String(i + 1).padStart(2, '0')}
          </div>
        </td>
        <td style="vertical-align:top;">
          <p style="margin:0 0 4px 0;font-size:11px;font-weight:700;color:#FF5E1F;">${bucket}</p>
          <p style="margin:0 0 7px 0;font-size:15px;font-weight:700;color:#222222;line-height:1.5;">
            ${escapeHtml(s.company || '')}${s.company ? ' · ' : ''}${escapeHtml(s.product || a.title)}
          </p>
          <p style="margin:0;font-size:13px;color:#666666;line-height:1.8;">
            <strong>[특징]</strong> ${escapeHtml(s.feature || '')}
          </p>
          <p style="margin:6px 0 0 0;font-size:13px;color:#666666;line-height:1.8;">
            <strong>[우체국보험 시사점]</strong> ${escapeHtml(s.implication || '')}
          </p>
          <p style="margin:8px 0 0 0;font-size:11px;color:#aaaaaa;">
            보험신보 | ${escapeHtml(a.articleDate)}
            · <a href="${escapeHtml(a.url)}" target="_blank" style="color:#999999;text-decoration:underline;">원문보기</a>
          </p>
        </td>
      </tr>
    </table>
  </td>
</tr>`;
  }).join('');
}

function buildHtml(meta, articles, summaries) {
  const [year, month] = TARGET_MONTH.split('-');
  const monthNumber = Number(month);

  return `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css');
    body,table,td,p,div,span,a,h1,h2,h3 {
      font-family:Pretendard,"Apple SD Gothic Neo","Malgun Gothic","맑은 고딕",Arial,sans-serif !important;
    }
  </style>
</head>
<body style="margin:0;padding:0;background:#f5f5f5;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation" style="background:#f5f5f5;">
<tr>
<td align="center" style="padding:24px 12px;">
<table width="640" cellpadding="0" cellspacing="0" border="0" role="presentation" style="width:100%;max-width:640px;background:#ffffff;border:1px solid #eeeeee;border-collapse:collapse;">

<tr>
<td style="background:#FF5E1F;color:#ffffff;padding:30px 28px 25px 28px;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" role="presentation">
<tr>
<td style="font-size:12px;font-weight:700;letter-spacing:2px;color:#FFE6DA;">PRODUCT TREND BRIEF</td>
<td style="font-size:12px;color:#FFE6DA;text-align:right;">${year}. ${month}</td>
</tr>
</table>
<h1 style="margin:17px 0 0 0;font-size:25px;font-weight:800;line-height:1.4;color:#ffffff;">
${monthNumber}월 민영사 신상품 동향
</h1>
<p style="margin:7px 0 0 0;font-size:13px;line-height:1.7;color:#FFE6DA;">
이달의 민영 생명보험사 신상품
</p>
</td>
</tr>

<tr>
<td style="padding:22px 28px;border-bottom:1px solid #eeeeee;">
<p style="margin:0;font-size:14px;color:#444444;line-height:1.85;">
${monthNumber}월 생명보험협회 판매일자 기준
<strong>상품 ${meta.productCount}건</strong>,
<strong>급부 항목 ${meta.benefitCount}건</strong>이 확인되었습니다.
동일 상품이 보험기간·급부 등에 따라 여러 행으로 구성된 경우 상품 수는 1건으로 계산했습니다.
</p>
</td>
</tr>

<tr>
<td style="padding:25px 28px 10px 28px;">
<span style="display:inline-block;font-size:11px;font-weight:700;color:#ffffff;background:#FF5E1F;padding:4px 10px;border-radius:3px;letter-spacing:1px;">KLIA DISCLOSURE</span>
<h2 style="font-size:18px;font-weight:800;color:#222222;margin:11px 0 4px 0;">공시에서 살펴본 상품 4선</h2>
<p style="font-size:12px;color:#999999;margin:0;">어린이보험 · 암보험 · 간병/치매보험 · 기타에서 각 1개 상품을 선정했습니다.</p>
</td>
</tr>

${buildDisclosureCards(meta.newsletterSamples || [])}

<tr>
<td style="padding:25px 28px 10px 28px;">
<span style="display:inline-block;font-size:11px;font-weight:700;color:#ffffff;background:#FF5E1F;padding:4px 10px;border-radius:3px;letter-spacing:1px;">NEWS PICK</span>
<h2 style="font-size:18px;font-weight:800;color:#222222;margin:11px 0 4px 0;">보험신보 신상품 4선</h2>
<p style="font-size:12px;color:#999999;margin:0;">기사 작성일 기준 1~10일 1건, 11~20일 2건, 21일~말일 1건을 선정했습니다.</p>
</td>
</tr>

${buildNewsCards(articles, summaries)}

<tr>
<td style="padding:22px 28px;background:#FFF7F3;border-top:1px solid #FFD8C8;">
<p style="font-size:13px;font-weight:700;color:#E84E12;margin:0 0 7px 0;">더 자세히 보고 싶다면?</p>
<p style="font-size:12px;color:#666666;line-height:1.8;margin:0;">
자세한 상품 내용은 <strong>첨부파일(V5 민영 생명보험 상품 통합본)</strong>과
<strong>보험신보 원문</strong>을 참고해주시기 바랍니다.
</p>
</td>
</tr>

<tr>
<td style="padding:18px 28px;background:#2B2B2B;color:#999999;text-align:center;font-size:11px;line-height:1.7;">
민영보험사 상품동향 조사·분석<br>
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

async function sendMail(html) {
  if (!SEND_MAIL) {
    console.log('메일 발송 건너뜀: SEND_MAIL=false');
    return;
  }

  if (!V5_MAIL_URL || !V5_MAIL_SECRET) {
    throw new Error('SEND_MAIL=true인데 V5_MAIL_URL/V5_MAIL_SECRET이 없습니다.');
  }

  const monthNumber = Number(TARGET_MONTH.slice(5, 7));

  const response = await fetch(V5_MAIL_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      secret: V5_MAIL_SECRET,
      action: 'sendNewsletterV5',
      targetMonth: TARGET_MONTH,
      subject: `${monthNumber}월 민영사 신상품 동향`,
      htmlBody: html,
      attachmentFileName: path.basename(EXCEL_FILE)
    })
  });

  const text = await response.text();
  const data = JSON.parse(text);

  if (!response.ok || !data.ok) {
    throw new Error(data.error || `V5 메일 발송 실패 HTTP ${response.status}`);
  }

  console.log('V5 뉴스레터 메일 발송 완료');
}

async function main() {
  if (!fs.existsSync(META_FILE)) {
    throw new Error(`메타파일 없음: ${META_FILE}`);
  }

  const meta = JSON.parse(fs.readFileSync(META_FILE, 'utf8'));

  const browser = await chromium.launch({ headless: true });

  let selected = [];

  try {
    const list = await collectInsweekList(browser);
    console.log(`보험신보 대상월 생보 신상품 후보: ${list.length}건`);

    selected = selectFourArticles(list);
    console.log(`보험신보 선정: ${selected.length}건`);

    selected = await enrichArticles(browser, selected);

  } finally {
    await browser.close();
  }

  let summaries;

  try {
    summaries = await summarizeWithGemini(selected);
    console.log('Gemini 기사 요약 완료');
  } catch (e) {
    console.log(`Gemini 요약 실패: ${e.message}`);

    summaries = selected.map((a, i) => ({
      index: i + 1,
      company: '',
      product: a.title,
      feature: '자동 요약에 실패했습니다. 보험신보 원문을 참고해주세요.',
      implication: '원문 확인 후 상품개발 관점에서 검토가 필요합니다.'
    }));
  }

  const html = buildHtml(meta, selected, summaries);

  fs.writeFileSync(HTML_FILE, html, 'utf8');
  fs.writeFileSync(
    JSON_FILE,
    JSON.stringify({
      targetMonth: TARGET_MONTH,
      disclosureSamples: meta.newsletterSamples || [],
      insweekArticles: selected,
      summaries
    }, null, 2),
    'utf8'
  );

  console.log(`V5 뉴스레터 HTML 생성: ${HTML_FILE}`);

  await sendMail(html);
}

main().catch(err => {
  console.error('newsletter_v5.js 실행 실패');
  console.error(err);
  process.exit(1);
});
