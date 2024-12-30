const puppeteer = require("puppeteer");
const fs = require("fs");
const csvParser = require("csv-parser");
const ExcelJS = require("exceljs"); // Import ExcelJS
const path = require("path");

// 입력 파일과 출력 폴더
const INPUT_FOLDER = "./input"; // CSV 파일들이 있는 폴더
const OUTPUT_FOLDER = "./output"; // 결과 Excel 파일 저장 폴더

if (!fs.existsSync(OUTPUT_FOLDER)) {
  fs.mkdirSync(OUTPUT_FOLDER); // 출력 폴더가 없으면 생성
}

// HTTP 요청 옵션
const options = {
  headers: {
    "User-Agent":
      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Accept-Language": "ko,en-US;q=0.9,en;q=0.8,ko-KR;q=0.7",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7"
  },
};

let firstResponseSaved = false; // 첫 번째 응답 저장 여부 확인 변수

// 검색 및 데이터 추출 함수
async function fetchData(searchText, index, results) {
  const url = `https://hanja.dict.naver.com/#/search?query=${encodeURIComponent(
    searchText
  )}`;
  console.log(`(${index + 1}) 검색 URL: ${url}`);

  let browser;
  try {
    // 브라우저 열기
    browser = await puppeteer.launch();
    const page = await browser.newPage();

    // 페이지 이동
    await page.goto(url, { waitUntil: "networkidle2" });

    // 첫 번째 HTML 응답 데이터를 저장
    if (!firstResponseSaved) {
      const html = await page.content();
      fs.writeFileSync("response.html", html, "utf8");
      console.log("첫 번째 HTML 응답 데이터를 response.html 파일로 저장했습니다.");
      firstResponseSaved = true; // 저장 완료 표시
    }

    // 데이터 추출

    const hanja = await page
      .evaluate(() => {
        try {
          const element = document.querySelector('div.hanja_word strong');
          return element?.textContent.trim() || "표제어 없음";
        } catch {
          return "표제어 없음";
        }
      })
      .catch(() => "표제어 없음");

    const meaning = await page
      .evaluate(() => {
        try {
          const element = document.querySelector('div.hanja_word div.mean');
          return element?.textContent.trim() || "";
        } catch {
          return "";
        }
      })
      .catch(() => "");


    const radicalData = await page
      .evaluate(() => {
        // "부수"라는 텍스트를 가진 div.cate 요소를 찾기
        const cateElem = Array.from(document.querySelectorAll('div.cate'))
          .find(elem => elem.textContent.trim() === '부수');

        // cateElem의 형제 요소 중 div.desc 찾기
        const descElem = cateElem?.nextElementSibling;

        // 부수 (radical) 값 찾기
        const radicalElem = descElem?.querySelector('span');
        const radical = radicalElem?.textContent.trim() || "";

        // 괄호 안의 부수 훈음 (radicalMeaning) 찾기
        const meaningMatch = descElem?.textContent.match(/\((.*?)\)/);
        const radicalMeaning = meaningMatch ? meaningMatch[1].trim() : "부수 훈음 없음";

        return { radical, radicalMeaning };
      })
      .catch(() => ({
        radical: "",
        radicalMeaning: "",
      }));

    const strokeCount = await page
      .evaluate(() => {
        const cateElem = Array.from(document.querySelectorAll('div.cate'))
          .find(elem => elem.textContent.trim() === '총 획수');

        // cateElem의 형제 요소 중 div.desc 찾기
        const descElem = cateElem?.nextElementSibling;
        const count = descElem.textContent.trim() || "";

        // '획' 이후 문자열 제거
        const cutoffIndex = count.indexOf('획');
        return cutoffIndex !== -1 ? count.substring(0, cutoffIndex).trim() : "";

      })
      .catch(() => "");

    // 결과 저장
    results.push({
      SEARCH_TEXT: searchText,
      Hanja: hanja,
      Meaning: meaning,
      Radical: radicalData.radical,
      RadicalMeaning: radicalData.radicalMeaning,
      StrokeCount: strokeCount,
    });

    console.log(
      `결과: 표제어= ${hanja}, 훈음= ${meaning}, 부수= ${radicalData.radical}, 부수훈음= ${radicalData.radicalMeaning}, 획수= ${strokeCount}`
    );
  } catch (error) {
    console.error(`(${index + 1}) 검색 실패: ${searchText} - ${error.message}`);
  } finally {
    if (browser) {
      await browser.close();
    }
  }
}

// CSV 파일 읽기 및 처리
async function processCSV(filePath, outputFilePath) {
  const targets = [];

  // 스트림 처리와 결과 저장을 명시적으로 비동기 처리
  await new Promise((resolve, reject) => {
    fs.createReadStream(filePath, { encoding: "utf8" })
      .pipe(csvParser({
        skipLines: 0, // 첫 줄을 헤더로 사용
        trim: true, // 자동 공백 제거
      }))
      .on("data", (row) => {
        if (row.SEARCH_TEXT && typeof row.SEARCH_TEXT === "string") {
          targets.push(row.SEARCH_TEXT);
        }
      })
      .on("end", async () => {
        console.log("CSV 파일 읽기 완료. 검색 시작...");

        const results = []; // 결과 데이터를 저장할 배열

        for (const [index, searchText] of targets.entries()) {
          if (typeof searchText === "string" && searchText.length > 0) {
            await fetchData(searchText, index, results); // 올바른 문자열만 전달
          } else {
            console.warn(`올바르지 않은 검색어: ${searchText} (index: ${index})`);
            reject();
          }
        }

        // 엑셀 파일 생성 (ExcelJS 사용)
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet("Sheet1");

        // 헤더 추가
        worksheet.columns = [
          { header: "SEARCH_TEXT", key: "SEARCH_TEXT", width: 20 },
          { header: "Hanja", key: "Hanja", width: 20 },
          { header: "Meaning", key: "Meaning", width: 30 },
          { header: "Radical", key: "Radical", width: 20 },
          { header: "RadicalMeaning", key: "RadicalMeaning", width: 30 },
          { header: "StrokeCount", key: "StrokeCount", width: 20 },
        ];

        // 데이터 추가
        results.forEach((result) => {
          worksheet.addRow(result);
        });

        // 파일 저장
        await workbook.xlsx.writeFile(outputFilePath);

        console.log(`검색 완료! 결과는 ${outputFilePath}에 저장되었습니다.`);
        resolve();
      })
      .on("error", (error) => {
        console.error("CSV 파일 읽기 중 오류 발생:", error.message);
        reject();
      });
  })
}

async function processAllFiles() {
  const files = fs.readdirSync(INPUT_FOLDER).filter((file) => file.endsWith(".csv"));

  if (files.length === 0) {
    console.error("입력 폴더에 CSV 파일이 없습니다.");
    return;
  }

  for (const file of files) {
    // 파일 경로 생성
    const inputFile = path.join(INPUT_FOLDER, file); // 절대 경로 생성
    const outputFile = path.join(OUTPUT_FOLDER, file.replace(".csv", ".xlsx"));

    try {
      console.log(`처리 중: ${inputFile}`);
      await processCSV(inputFile, outputFile);
    } catch (error) {
      console.error(`파일 처리 중 오류 발생 (${file}):`, error.message);
    }
  }

  console.log("모든 파일 처리가 완료되었습니다!");
}

// 실행
processAllFiles();