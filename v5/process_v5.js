const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const OUTPUT_DIR = path.join(ROOT, 'output');
const TARGET_MONTH = String(process.env.TARGET_MONTH || '').trim();
const UPLOAD_URL = process.env.UPLOAD_URL || '';
const UPLOAD_SECRET = process.env.UPLOAD_SECRET || '';

if (!/^\d{4}-\d{2}$/.test(TARGET_MONTH)) {
  throw new Error('TARGET_MONTH는 YYYY-MM 형식이어야 합니다.');
}

fs.mkdirSync(OUTPUT_DIR, { recursive: true });

const OUTPUT_FILE = path.join(
  OUTPUT_DIR,
  `${TARGET_MONTH}_V5_생보협회_상품통합.xlsx`
);

const META_FILE = path.join(
  OUTPUT_DIR,
  `${TARGET_MONTH}_V5_meta.json`
);

const TARGET_COLUMNS = [
  '보험회사명',
  '상품명',
  '급부명칭',
  '지급사유',
  '지급금액',
  '판매일자',
  '특이사항'
];

const HEADER_ALIASES = {
  보험회사명: ['보험회사명', '회사명', '보험회사', '생명보험회사'],
  상품명: ['상품명', '보험상품명', '상품 명'],
  급부명칭: ['급부명칭', '급부명', '급부 명칭', '보장명', '담보명'],
  지급사유: ['지급사유', '보험금지급사유', '보험금 지급사유', '급부지급사유', '지급조건'],
  지급금액: ['지급금액', '보험금액', '급부금액', '보장금액', '지급 금액'],
  판매일자: ['판매일자', '판매개시일', '판매개시일자', '판매일', '판매 일자'],
  특이사항: ['특이사항', '비고', '특이 사항']
};

function clean(v) {
  return String(v ?? '')
    .replace(/\u00a0/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/\n+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function norm(v) {
  return clean(v)
    .toLowerCase()
    .replace(/[\s\(\)\[\]\{\}·ㆍ\-_,.\/\\:：'"“”‘’]/g, '');
}

function uniqueProductKey(row) {
  return `${norm(row.보험회사명)}|${norm(row.상품명)}`;
}

function productBenefitKey(row) {
  return `${uniqueProductKey(row)}|${norm(row.급부명칭)}`;
}

function parseCategoryFromFile(filePath) {
  const name = path.basename(filePath);
  const categories = [
    '종신보험', '정기보험', '질병보험', '암보험', 'CI보험', '상해보험',
    '어린이보험', '치아보험', '간병/치매보험', '간병치매보험', '간병·치매보험',
    '기타', '연금', '저축', '교육'
  ];

  for (const c of categories) {
    if (name.includes(c)) {
      if (['간병치매보험', '간병·치매보험'].includes(c)) return '간병/치매보험';
      return c;
    }
  }
  return '';
}

function isSavingCategory(category, filePath) {
  const name = path.basename(filePath);
  return name.includes('저축성') || ['연금', '저축', '교육'].includes(category);
}

function excelDateToIso(value) {
  if (value === null || value === undefined || value === '') return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
    }
  }

  const s = clean(value);

  let m = s.match(/^(20\d{2})[-./년\s]+(\d{1,2})[-./월\s]+(\d{1,2})/);
  if (m) {
    return `${m[1]}-${String(Number(m[2])).padStart(2, '0')}-${String(Number(m[3])).padStart(2, '0')}`;
  }

  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;

  return s;
}

function cloneSheetWithExpandedMerges(ws) {
  const copy = {};

  for (const key of Object.keys(ws)) {
    if (key === '!merges') continue;
    copy[key] = ws[key] && typeof ws[key] === 'object'
      ? { ...ws[key] }
      : ws[key];
  }

  copy['!ref'] = ws['!ref'];

  for (const merge of ws['!merges'] || []) {
    const source = ws[XLSX.utils.encode_cell(merge.s)];
    if (!source) continue;

    for (let r = merge.s.r; r <= merge.e.r; r++) {
      for (let c = merge.s.c; c <= merge.e.c; c++) {
        copy[XLSX.utils.encode_cell({ r, c })] = { ...source };
      }
    }
  }

  return copy;
}

function findHeaderRowAndColumns(ws) {
  const expanded = cloneSheetWithExpandedMerges(ws);
  const range = XLSX.utils.decode_range(expanded['!ref'] || 'A1:A1');
  const maxRow = Math.min(range.e.r, 29);
  let best = null;

  for (let r = range.s.r; r <= maxRow; r++) {
    const found = {};
    let score = 0;

    for (let c = range.s.c; c <= range.e.c; c++) {
      const value = clean(expanded[XLSX.utils.encode_cell({ r, c })]?.v);
      if (!value) continue;

      for (const [field, aliases] of Object.entries(HEADER_ALIASES)) {
        if (aliases.some(alias => norm(value) === norm(alias)) && found[field] === undefined) {
          found[field] = c;
          score++;
        }
      }
    }

    if (!best || score > best.score) best = { row: r, columns: found, score };
  }

  if (!best || best.score < 3 || best.columns.보험회사명 === undefined || best.columns.상품명 === undefined) {
    return null;
  }

  return { expanded, headerRow: best.row, columns: best.columns };
}

function readDataRows(filePath) {
  const category = parseCategoryFromFile(filePath);
  const saving = isSavingCategory(category, filePath);
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const out = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;

    const info = findHeaderRowAndColumns(ws);
    if (!info) continue;

    const { expanded, headerRow, columns } = info;
    const range = XLSX.utils.decode_range(expanded['!ref']);

    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const get = field => {
        const c = columns[field];
        if (c === undefined) return '';
        return expanded[XLSX.utils.encode_cell({ r, c })]?.v ?? '';
      };

      const company = clean(get('보험회사명'));
      const product = clean(get('상품명'));
      if (!company || !product) continue;

      const saleDate = excelDateToIso(get('판매일자'));
      if (!saleDate || saleDate.slice(0, 7) !== TARGET_MONTH) continue;

      out.push({
        category,
        row: {
          보험회사명: company,
          상품명: product,
          급부명칭: saving ? '' : clean(get('급부명칭')),
          지급사유: saving ? '' : clean(get('지급사유')),
          지급금액: saving ? '' : clean(get('지급금액')),
          판매일자: saleDate,
          특이사항: clean(get('특이사항'))
        }
      });
    }
  }

  return out;
}

function walk(dir, depth = 0) {
  if (!fs.existsSync(dir) || depth > 3) return [];
  const out = [];

  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);

    if (entry.isDirectory()) {
      if (['node_modules', '.git', 'v5'].includes(entry.name)) continue;
      out.push(...walk(full, depth + 1));
    } else if (entry.isFile() && /\.(xls|xlsx)$/i.test(entry.name)) {
      out.push(full);
    }
  }

  return out;
}

function findRawFiles() {
  const roots = [
    path.join(ROOT, 'output'),
    path.join(ROOT, 'downloads'),
    path.join(ROOT, 'download'),
    path.join(ROOT, 'excel'),
    ROOT
  ];

  const all = new Set();

  for (const root of roots) {
    for (const f of walk(root)) {
      const name = path.basename(f);
      if (
        name.includes(TARGET_MONTH) &&
        !name.includes('_V5_') &&
        !name.includes('상품통합') &&
        !name.includes('상품DB')
      ) {
        all.add(f);
      }
    }
  }

  return [...all];
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

function pickOneByProduct(records, category, seedSuffix) {
  const grouped = new Map();

  for (const item of records.filter(x => x.category === category)) {
    const key = uniqueProductKey(item.row);
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(item.row);
  }

  const products = [...grouped.values()];
  if (!products.length) return null;

  const rng = mulberry32(makeSeed(`${TARGET_MONTH}|${category}|${seedSuffix}`));
  const rows = products[Math.floor(rng() * products.length)];
  const representative = rows.find(r => clean(r.급부명칭)) || rows[0];

  return {
    카테고리: category,
    보험회사명: representative.보험회사명,
    상품명: representative.상품명,
    급부명칭: representative.급부명칭,
    지급사유: representative.지급사유,
    지급금액: representative.지급금액
  };
}

async function uploadIfConfigured(filePath) {
  if (!UPLOAD_URL || !UPLOAD_SECRET) {
    console.log('Drive 업로드 건너뜀: UPLOAD_URL/UPLOAD_SECRET 없음');
    return;
  }

  const response = await fetch(UPLOAD_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      secret: UPLOAD_SECRET,
      action: 'upload',
      folderType: 'merged',
      fileName: path.basename(filePath),
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64: fs.readFileSync(filePath).toString('base64')
    })
  });

  const text = await response.text();
  const data = JSON.parse(text);

  if (!response.ok || !data.ok) {
    throw new Error(data.error || `Drive 업로드 실패 HTTP ${response.status}`);
  }

  console.log(`Drive 저장 완료: ${path.basename(filePath)}`);
}

async function main() {
  const rawFiles = findRawFiles();

  console.log(`V5 대상월: ${TARGET_MONTH}`);
  console.log(`V5 원본 엑셀 발견: ${rawFiles.length}개`);

  if (!rawFiles.length) {
    throw new Error('대상월 원본 XLS/XLSX 파일을 찾지 못했습니다.');
  }

  const records = [];

  for (const file of rawFiles) {
    const rows = readDataRows(file);
    if (rows.length) {
      console.log(`- ${path.basename(file)}: ${rows.length}행`);
      records.push(...rows);
    }
  }

  if (!records.length) {
    throw new Error('대상월 판매일자에 해당하는 행이 없습니다.');
  }

  const mainRows = records.map(x => x.row);

  const uniqueProducts = new Set(mainRows.map(uniqueProductKey));
  const benefitRows = mainRows.filter(r => clean(r.급부명칭));
  const uniqueBenefits = new Set(benefitRows.map(productBenefitKey));

  const categoryProductCounts = {};

  for (const item of records) {
    if (!item.category) continue;
    if (!categoryProductCounts[item.category]) {
      categoryProductCounts[item.category] = new Set();
    }
    categoryProductCounts[item.category].add(uniqueProductKey(item.row));
  }

  const newsletterSamples = [
    '어린이보험',
    '암보험',
    '간병/치매보험',
    '기타'
  ]
    .map((category, i) => pickOneByProduct(records, category, i))
    .filter(Boolean);

  const summaryRows = [
    { 항목: '판매개시 상품 수', 건수: uniqueProducts.size },
    { 항목: '급부 항목 수', 건수: uniqueBenefits.size },
    { 항목: '통합 행 수', 건수: mainRows.length },
    ...Object.entries(categoryProductCounts)
      .sort((a, b) => a[0].localeCompare(b[0], 'ko'))
      .map(([category, set]) => ({
        항목: `${category} 상품 수`,
        건수: set.size
      }))
  ];

  const wb = XLSX.utils.book_new();

  const wsMain = XLSX.utils.json_to_sheet(mainRows, {
    header: TARGET_COLUMNS
  });

  wsMain['!cols'] = [
    { wch: 18 },
    { wch: 45 },
    { wch: 30 },
    { wch: 55 },
    { wch: 28 },
    { wch: 14 },
    { wch: 40 }
  ];

  wsMain['!autofilter'] = {
    ref: XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: mainRows.length, c: TARGET_COLUMNS.length - 1 }
    })
  };

  XLSX.utils.book_append_sheet(wb, wsMain, '통합본');

  const wsSummary = XLSX.utils.json_to_sheet(summaryRows);
  wsSummary['!cols'] = [{ wch: 30 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, wsSummary, '요약');

  XLSX.writeFile(wb, OUTPUT_FILE);

  fs.writeFileSync(
    META_FILE,
    JSON.stringify({
      targetMonth: TARGET_MONTH,
      totalRows: mainRows.length,
      productCount: uniqueProducts.size,
      benefitCount: uniqueBenefits.size,
      categoryProductCounts: Object.fromEntries(
        Object.entries(categoryProductCounts).map(([k, set]) => [k, set.size])
      ),
      newsletterSamples
    }, null, 2),
    'utf8'
  );

  console.log(`V5 통합본 생성: ${OUTPUT_FILE}`);
  console.log(`상품 수: ${uniqueProducts.size}`);
  console.log(`급부 항목 수: ${uniqueBenefits.size}`);

  await uploadIfConfigured(OUTPUT_FILE);
}

main().catch(err => {
  console.error('process_v5.js 실행 실패');
  console.error(err);
  process.exit(1);
});
