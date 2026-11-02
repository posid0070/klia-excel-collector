const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

// 다운로드 받을 폴더
const DOWNLOAD_DIR = path.join(__dirname, "downloads");

// 조사 대상
const GROUPS = [
  {
    group: "보장성",
    url: "https://pub.insure.or.kr/compareDis/prodCompare/assurance/listNew.do",
    categories: [
      "종신보험",
      "정기보험",
      "질병보험",
      "암보험",
      "CI보험",
      "상해보험",
      "어린이보험",
      "치아보험",
      "간병/치매보험",
      "기타"
    ]
  },
  {
    group: "저축성",
    url: "https://pub.insure.or.kr/compareDis/prodCompare/saving/list.do",
    categories: [
      "연금보험",
      "저축보험",
      "교육보험"
    ]
  }
];

function safeFileName(text) {
  return String(text)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, "_");
}

async function selectCategory(page, category) {

  console.log(`선택 중: ${category}`);

  // 해당 글자를 포함한 label 찾기
  const label = page
    .locator("label")
    .filter({ hasText: category })
    .first();

  if (await label.count() === 0) {
    throw new Error(`카테고리를 찾지 못했습니다: ${category}`);
  }

  await label.scrollIntoViewIfNeeded();

  // 실제 화면 클릭
  await label.click();

  await page.waitForTimeout(700);

  // 조회하기 버튼이 있으면 클릭
  const searchButton = page.getByRole("button", {
    name: "조회하기",
    exact: true
  });

  if (await searchButton.count() > 0) {

    console.log("조회하기 클릭");

    await searchButton.first().click();

    // 페이지 또는 조회결과가 갱신될 시간
    await page.waitForTimeout(1500);
  }
}

async function downloadExcel(page, group, category) {

  const excelButton = page
    .locator("button.btn_excel")
    .filter({ hasText: "엑셀다운로드" })
    .first();

  if (await excelButton.count() === 0) {
    throw new Error(
      `${group} / ${category} : 엑셀다운로드 버튼을 찾지 못했습니다.`
    );
  }

  await excelButton.scrollIntoViewIfNeeded();

  console.log(`엑셀 다운로드 클릭: ${group} / ${category}`);

  const downloadPromise = page.waitForEvent(
    "download",
    { timeout: 30000 }
  );

  await excelButton.click();

  const download = await downloadPromise;

  const originalName =
    download.suggestedFilename() || "download.xlsx";

  const extension =
    path.extname(originalName) || ".xlsx";

  const newFileName =
    `KLIA_${safeFileName(group)}_${safeFileName(category)}${extension}`;

  const savePath =
    path.join(DOWNLOAD_DIR, newFileName);

  await download.saveAs(savePath);

  console.log(`다운로드 완료: ${newFileName}`);

  return savePath;
}

async function main() {

  console.log("=================================");
  console.log("생명보험협회 Excel 수집 시작");
  console.log("=================================");

  // downloads 폴더 초기화
  if (fs.existsSync(DOWNLOAD_DIR)) {
    fs.rmSync(
      DOWNLOAD_DIR,
      {
        recursive: true,
        force: true
      }
    );
  }

  fs.mkdirSync(
    DOWNLOAD_DIR,
    {
      recursive: true
    }
  );

  const browser = await chromium.launch({
    headless: true
  });

  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: {
      width: 1440,
      height: 1200
    }
  });

  const page = await context.newPage();

  page.setDefaultTimeout(20000);

  let successCount = 0;
  const failed = [];

  try {

    for (const item of GROUPS) {

      console.log("");
      console.log("==============================");
      console.log(`${item.group} 시작`);
      console.log("==============================");

      for (const category of item.categories) {

        try {

          // 매 항목마다 페이지 새로 열기
          console.log("");
          console.log(
            `[${successCount + 1}/13] ${item.group} / ${category}`
          );

          await page.goto(
            item.url,
            {
              waitUntil: "domcontentloaded",
              timeout: 45000
            }
          );

          await page.waitForTimeout(1200);

          // 유형 선택
          await selectCategory(
            page,
            category
          );

          // Excel 다운로드
          await downloadExcel(
            page,
            item.group,
            category
          );

          successCount++;

        } catch (error) {

          console.error(
            `실패: ${item.group} / ${category}`
          );

          console.error(error.message);

          failed.push(
            `${item.group} / ${category} : ${error.message}`
          );
        }
      }
    }

  } finally {

    await browser.close();
  }

  console.log("");
  console.log("==============================");
  console.log("수집 결과");
  console.log("==============================");

  console.log(`성공: ${successCount}개`);
  console.log(`실패: ${failed.length}개`);

  if (failed.length > 0) {

    console.log("");
    console.log("실패 목록");

    failed.forEach(x =>
      console.log(x)
    );
  }

  // 13개 전부 받아야 정상 종료
  if (successCount !== 13) {

    throw new Error(
      `13개 중 ${successCount}개만 다운로드되었습니다.`
    );
  }

  console.log("");
  console.log("13개 Excel 다운로드 성공!");
}

main()
  .then(() => {
    console.log("작업 정상 종료");
  })
  .catch(error => {

    console.error("");
    console.error("작업 실패");
    console.error(error);

    process.exit(1);
  });
