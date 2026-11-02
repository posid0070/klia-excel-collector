const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

async function run() {
  console.log('브라우저 시작 중...');
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  const downloadDir = path.join(__dirname, 'downloads');
  if (!fs.existsSync(downloadDir)) fs.mkdirSync(downloadDir);

  const groups = [
    { name: '보장성', url: 'https://pub.insure.or.kr/compareDis/prodCompare/assurance/listNew.do', count: 10 },
    { name: '저축성', url: 'https://pub.insure.or.kr/compareDis/prodCompare/saving/list.do', count: 3 }
  ];

  for (const group of groups) {
    console.log(`${group.name} 페이지 접속 중...`);
    await page.goto(group.url, { waitUntil: 'networkidle' });
    await page.waitForTimeout(2000);

    for (let i = 0; i < group.count; i++) {
      try {
        await page.goto(group.url, { waitUntil: 'networkidle' });
        await page.waitForTimeout(1500);

        const radios = await page.$$('input[type="radio"]');
        if (radios[i]) {
          await radios[i].click({ force: true });
          await page.waitForTimeout(1000);

          const [download] = await Promise.all([
            page.waitForEvent('download', { timeout: 15000 }),
            page.evaluate(() => { if (typeof fn_excelDown === 'function') fn_excelDown(); })
          ]);

          const suggestedName = download.suggestedFilename();
          const filePath = path.join(downloadDir, `KLIA_${group.name}_${i}_${suggestedName}`);
          await download.saveAs(filePath);
          console.log(`다운로드 완료: ${suggestedName}`);
        }
      } catch (err) {
        console.error(`${group.name} 유형 ${i}번째 다운로드 실패:`, err.message);
      }
    }
  }

  await browser.close();
  console.log('모든 다운로드 종료.');
}

run().catch(console.error);
