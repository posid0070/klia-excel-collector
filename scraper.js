const { chromium } = require("playwright");
const fs = require("fs");
const path = require("path");

const DOWNLOAD_DIR = path.join(__dirname, "downloads");

// 생명보험협회 상품비교공시
const GROUPS = [
  {
    group: "보장성",
    url: "https://pub.insure.or.kr/compareDis/prodCompare/assurance/listNew.do",
    categories: [
      { name: "종신보험", aliases: ["종신보험", "종신"] },
      { name: "정기보험", aliases: ["정기보험", "정기"] },
      { name: "질병보험", aliases: ["질병보험", "질병"] },
      { name: "암보험", aliases: ["암보험", "암"] },
      { name: "CI보험", aliases: ["CI보험", "CI"] },
      { name: "상해보험", aliases: ["상해보험", "상해"] },
      { name: "어린이보험", aliases: ["어린이보험", "어린이"] },
      { name: "치아보험", aliases: ["치아보험", "치아"] },
      {
        name: "간병/치매보험",
        aliases: [
          "간병/치매보험",
          "간병/치매",
          "간병치매보험",
          "간병치매"
        ]
      },
      { name: "기타", aliases: ["기타"] }
    ]
  },

  {
    group: "저축성",
    url: "https://pub.insure.or.kr/compareDis/prodCompare/saving/list.do",
    categories: [
      { name: "연금", aliases: ["연금보험", "연금"] },
      { name: "저축", aliases: ["저축보험", "저축"] },
      { name: "교육", aliases: ["교육보험", "교육"] }
    ]
  }
];

function safeFileName(text) {
  return String(text)
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, "_");
}

/**
 * 카테고리 선택
 *
 * 화면에서 label을 직접 클릭하는 대신
 * DOM에서 해당 radio를 찾아 선택한다.
 *
 * 이렇게 하면 '상품비교하기' 같은 버튼이
 * 화면 위를 가리고 있어도 선택할 수 있다.
 */
async function selectCategory(page, item) {

  console.log(`선택 중: ${item.name}`);

  const result = await page.evaluate(({ aliases }) => {

    function normalize(text) {
      return String(text || "")
        .replace(/\s+/g, "")
        .replace(/[()]/g, "")
        .toLowerCase();
    }

    const aliasList = aliases.map(normalize);

    const labels =
      Array.from(document.querySelectorAll("label"));

    const candidates = [];

    for (const label of labels) {

      const text =
        normalize(
          label.innerText ||
          label.textContent ||
          ""
        );

      const exact =
        aliasList.some(alias => text === alias);

      const included =
        aliasList.some(alias => text.includes(alias));

      if (!exact && !included) {
        continue;
      }

      let input =
        label.querySelector('input[type="radio"]');

      if (!input && label.htmlFor) {
        input =
          document.getElementById(label.htmlFor);
      }

      if (!input) {
        continue;
      }

      const visible =
        Boolean(
          label.offsetWidth ||
          label.offsetHeight ||
          label.getClientRects().length
        );

      candidates.push({
        input,
        exact,
        visible,
        text
      });
    }

    // 보이는 항목 + 정확히 일치하는 항목 우선
    candidates.sort((a, b) => {

      if (a.visible !== b.visible) {
        return a.visible ? -1 : 1;
      }

      if (a.exact !== b.exact) {
        return a.exact ? -1 : 1;
      }

      return 0;
    });

    if (candidates.length === 0) {

      return {
        ok: false,

        labels: labels
          .map(label =>
            (
              label.innerText ||
              label.textContent ||
              ""
            ).trim()
          )
          .filter(Boolean)
          .slice(0, 100)
      };
    }

    const radio =
      candidates[0].input;

    // 직접 선택
    radio.checked = true;

    // 사이트의 이벤트 로직이 동작하도록 이벤트 발생
    try {

      radio.dispatchEvent(
        new MouseEvent(
          "click",
          {
            bubbles: true,
            cancelable: true,
            view: window
          }
        )
      );

    } catch (e) {}

    radio.dispatchEvent(
      new Event(
        "input",
        { bubbles: true }
      )
    );

    radio.dispatchEvent(
      new Event(
        "change",
        { bubbles: true }
      )
    );

    return {
      ok: true,
      checked: radio.checked,
      name: radio.name || "",
      value: radio.value || "",
      text: candidates[0].text
    };

  }, {
    aliases: item.aliases
  });

  if (!result.ok) {

    console.log("");
    console.log("현재 페이지 label 목록");

    (result.labels || [])
      .forEach(text =>
        console.log(" - " + text)
      );

    throw new Error(
      `카테고리를 찾지 못했습니다: ${item.name}`
    );
  }

  console.log(
    `선택 성공: ${item.name}`
  );

  console.log(
    `radio name=${result.name}, value=${result.value}`
  );

  await page.waitForTimeout(700);

  /*
   * 조회하기 버튼이 실제로 존재하고 보이는 경우에만 클릭
   */
  const searchButton =
    page.getByRole(
      "button",
      {
        name: "조회하기",
        exact: true
      }
    );

  if (await searchButton.count() > 0) {

    const button =
      searchButton.first();

    const visible =
      await button
        .isVisible()
        .catch(() => false);

    if (visible) {

      console.log(
        "조회하기 클릭"
      );

      try {

        await button.click({
          force: true,
          timeout: 10000
        });

      } catch (e) {

        console.log(
          "일반 클릭 실패 → DOM 클릭"
        );

        await page.evaluate(() => {

          const buttons =
            Array.from(
              document.querySelectorAll("button")
            );

          const target =
            buttons.find(
              button =>
                (
                  button.innerText ||
                  button.textContent ||
                  ""
                ).trim() === "조회하기"
            );

          if (target) {
            target.click();
          }

        });
      }

      await page.waitForTimeout(1500);
    }
  }
}

/**
 * Excel 다운로드
 */
async function downloadExcel(
  page,
  group,
  category
) {

  const excelButton =
    page
      .locator("button.btn_excel")
      .filter({
        hasText: "엑셀다운로드"
      })
      .first();

  if (
    await excelButton.count() === 0
  ) {

    throw new Error(
      `${group} / ${category} : ` +
      `엑셀다운로드 버튼을 찾지 못했습니다.`
    );
  }

  console.log(
    `엑셀 다운로드 클릭: ` +
    `${group} / ${category}`
  );

  /*
   * 버튼을 누르면 브라우저 download 이벤트가 발생함.
   */
  const downloadPromise =
    page.waitForEvent(
      "download",
      {
        timeout: 45000
      }
    );

  try {

    await excelButton.click({
      force: true,
      timeout: 10000
    });

  } catch (e) {

    /*
     * 버튼을 화면상 클릭할 수 없는 경우
     * 사이트의 원래 JavaScript 함수 직접 실행
     */
    console.log(
      "버튼 클릭 실패 → fn_excelDown() 직접 실행"
    );

    await page.evaluate(() => {

      if (
        typeof fn_excelDown === "function"
      ) {

        fn_excelDown();

      } else {

        throw new Error(
          "fn_excelDown 함수를 찾지 못했습니다."
        );
      }

    });
  }

  const download =
    await downloadPromise;

  const originalName =
    download.suggestedFilename() ||
    "download.xls";

  const extension =
    path.extname(originalName) ||
    ".xls";

  const fileName =
    `KLIA_` +
    `${safeFileName(group)}_` +
    `${safeFileName(category)}` +
    `${extension}`;

  const savePath =
    path.join(
      DOWNLOAD_DIR,
      fileName
    );

  await download.saveAs(
    savePath
  );

  console.log(
    `다운로드 완료: ${fileName}`
  );

  return savePath;
}

/**
 * 실패했을 때 페이지에 있는 radio 목록 출력
 *
 * 저축성 페이지의 실제 HTML이 예상과 다를 때
 * 이 로그를 보면 바로 수정 가능.
 */
async function dumpRadioList(page) {

  const radios =
    await page.evaluate(() => {

      return Array
        .from(
          document.querySelectorAll(
            'input[type="radio"]'
          )
        )
        .map(
          (radio, index) => {

            let text = "";

            const parentLabel =
              radio.closest("label");

            if (parentLabel) {

              text =
                parentLabel.innerText ||
                parentLabel.textContent ||
                "";

            } else if (radio.id) {

              const label =
                document.querySelector(
                  `label[for="${radio.id}"]`
                );

              if (label) {

                text =
                  label.innerText ||
                  label.textContent ||
                  "";
              }
            }

            return {
              index,
              name: radio.name || "",
              value: radio.value || "",
              checked: radio.checked,
              text: String(text).trim()
            };
          }
        );

    });

  console.log("");
  console.log("현재 페이지 radio 목록");

  radios.forEach(r => {

    console.log(
      `#${r.index}` +
      ` name=${r.name}` +
      ` value=${r.value}` +
      ` checked=${r.checked}` +
      ` text="${r.text}"`
    );

  });
}

/**
 * 메인 실행
 */
async function main() {

  console.log(
    "================================="
  );

  console.log(
    "생명보험협회 Excel 수집 시작"
  );

  console.log(
    "================================="
  );

  /*
   * 기존 downloads 폴더 초기화
   */
  if (
    fs.existsSync(DOWNLOAD_DIR)
  ) {

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

  /*
   * Chromium 실행
   */
  const browser =
    await chromium.launch({
      headless: true
    });

  const context =
    await browser.newContext({

      acceptDownloads: true,

      viewport: {
        width: 1440,
        height: 1200
      },

      userAgent:
        "Mozilla/5.0 " +
        "(Windows NT 10.0; Win64; x64) " +
        "AppleWebKit/537.36 " +
        "(KHTML, like Gecko) " +
        "Chrome/155.0.0.0 Safari/537.36"
    });

  const page =
    await context.newPage();

  page.setDefaultTimeout(
    20000
  );

  let successCount = 0;

  const failed = [];

  try {

    /*
     * 보장성 → 저축성 순서
     */
    for (
      const group of GROUPS
    ) {

      console.log("");
      console.log(
        "=============================="
      );

      console.log(
        `${group.group} 시작`
      );

      console.log(
        `URL: ${group.url}`
      );

      console.log(
        "=============================="
      );

      /*
       * 각 상품 유형
       */
      for (
        const item of group.categories
      ) {

        try {

          console.log("");

          console.log(
            `[${successCount + 1}/13] ` +
            `${group.group} / ${item.name}`
          );

          /*
           * 매 유형마다 원본 페이지 새로 접속
           */
          const response =
            await page.goto(
              group.url,
              {
                waitUntil:
                  "domcontentloaded",

                timeout:
                  45000
              }
            );

          if (!response) {

            throw new Error(
              "페이지 응답이 없습니다."
            );
          }

          if (
            response.status() >= 400
          ) {

            throw new Error(
              `페이지 HTTP ` +
              `${response.status()}`
            );
          }

          await page.waitForTimeout(
            1500
          );

          console.log(
            `현재 주소: ${page.url()}`
          );

          /*
           * 상품 유형 선택
           */
          await selectCategory(
            page,
            item
          );

          /*
           * Excel 다운로드
           */
          await downloadExcel(
            page,
            group.group,
            item.name
          );

          successCount++;

          await page.waitForTimeout(
            500
          );

        } catch (error) {

          console.error("");

          console.error(
            `실패: ` +
            `${group.group} / ${item.name}`
          );

          console.error(
            error.message ||
            String(error)
          );

          /*
           * 실패 원인 확인용 radio 출력
           */
          try {

            await dumpRadioList(
              page
            );

          } catch (e) {}

          failed.push(
            `${group.group} / ${item.name}` +
            ` : ` +
            `${error.message || String(error)}`
          );
        }
      }
    }

  } finally {

    await browser.close();
  }

  /*
   * 최종 결과
   */
  console.log("");
  console.log(
    "=============================="
  );

  console.log(
    "수집 결과"
  );

  console.log(
    "=============================="
  );

  console.log(
    `성공: ${successCount}개`
  );

  console.log(
    `실패: ${failed.length}개`
  );

  if (
    failed.length > 0
  ) {

    console.log("");
    console.log(
      "실패 목록"
    );

    failed.forEach(
      item =>
        console.log(item)
    );
  }

  /*
   * 실제 생성된 파일 목록
   */
  console.log("");
  console.log(
    "downloads 폴더 파일"
  );

  if (
    fs.existsSync(DOWNLOAD_DIR)
  ) {

    fs
      .readdirSync(
        DOWNLOAD_DIR
      )
      .forEach(
        file =>
          console.log(
            " - " + file
          )
      );
  }

  /*
   * 정확히 13개가 아닐 경우 실패 처리
   */
  if (
    successCount !== 13
  ) {

    throw new Error(
      `13개 중 ` +
      `${successCount}개만 ` +
      `다운로드되었습니다.`
    );
  }

  console.log("");
  console.log(
    "13개 Excel 다운로드 성공!"
  );
}

/*
 * 프로그램 시작
 */
main()
  .then(() => {

    console.log(
      "작업 정상 종료"
    );

  })
  .catch(error => {

    console.error("");
    console.error(
      "작업 실패"
    );

    console.error(
      error
    );

    process.exit(1);
  });
