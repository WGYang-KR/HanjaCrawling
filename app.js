const fs = require("fs");
const csvParser = require("csv-parser");
const axios = require("axios");
const cheerio = require("cheerio");
const ExcelJS = require("exceljs"); // Import ExcelJS

// 입력 파일과 출력 파일 경로
const INPUT_FILE = "input.csv";
const OUTPUT_FILE = "output.xlsx";

// 결과 저장용 배열
const results = [];

// HTTP 요청 옵션
const options = {
  headers: {
    "User-Agent":
      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
      "Accept-Language": "ko,en-US;q=0.9,en;q=0.8,ko-KR;q=0.7"
  },
};

// 검색 및 데이터 추출 함수
async function fetchData(searchText, index) {
  try {
    const url = `https://hanja.dict.naver.com/#/search?query=${encodeURIComponent(
      searchText
    )}`;
    console.log(`(${index + 1}) 검색 URL: ${url}`);

    // 웹페이지 요청
    const { data: html } = await axios.get(url, options);
    const $ = cheerio.load(html);

    // 데이터 추출
    const hanja = $("strong.highlight").text() || "표제어 없음";
    const meaning = $(".mean").text().trim() || "훈음 없음";

    const radicalMatch = html.match(
      /<div class="cate">부수<\/div>\s*<div class="desc">.*?<span.*?>(.*?)<\/span>\((.*?)\)/
    );
    const radical = radicalMatch ? radicalMatch[1] : "부수 없음";
    const radicalMeaning = radicalMatch ? radicalMatch[2] : "부수 훈음 없음";

    const strokeMatch = html.match(
      /<div class="cate">총 획수<\/div>\s*<div class="desc">(.*?)획<\/div>/
    );
    const strokeCount = strokeMatch ? `${strokeMatch[1]}획` : "획수 없음";

    // 결과 저장
    results.push({
      SEARCH_TEXT: searchText,
      Hanja: hanja,
      Meaning: meaning,
      StrokeCount: strokeCount,
      Radical: radical,
      RadicalMeaning: radicalMeaning,
    });

    console.log(
      `결과: 표제어=${hanja}, 훈음=${meaning}, 획수=${strokeCount}, 부수=${radical}, 부수 훈음=${radicalMeaning}`
    );
  } catch (error) {
    console.error(`(${index + 1}) 검색 실패: ${searchText} - ${error.message}`);
  }
}

// CSV 파일 읽기 및 처리
function processCSV() {
  fs.createReadStream(INPUT_FILE, {encoding: 'utf8'})
    .pipe(csvParser())
    .on("data", (row) => {
      if (row.SEARCH_TEXT) {
        results.push(row.SEARCH_TEXT);
      }
    })
    .on("end", async () => {
      console.log("CSV 파일 읽기 완료. 검색 시작...");

      for (const [index, searchText] of results.entries()) {
        await fetchData(searchText, index);
      }

      // 엑셀 파일 생성 (ExcelJS 사용)
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Sheet1");

      // 헤더 추가
      worksheet.columns = [
        { header: "SEARCH_TEXT", key: "SEARCH_TEXT", width: 20 },
        { header: "Hanja", key: "Hanja", width: 20 },
        { header: "Meaning", key: "Meaning", width: 30 },
        { header: "StrokeCount", key: "StrokeCount", width: 20 },
        { header: "Radical", key: "Radical", width: 20 },
        { header: "RadicalMeaning", key: "RadicalMeaning", width: 30 },
      ];

      // 데이터 추가
      results.forEach((result) => {
        worksheet.addRow(result);
      });

      // 파일 저장
      await workbook.xlsx.writeFile(OUTPUT_FILE);

      console.log(`검색 완료! 결과는 ${OUTPUT_FILE}에 저장되었습니다.`);
    })
    .on("error", (error) => {
      console.error("CSV 파일 읽기 중 오류 발생:", error.message);
    });
}

// 실행
processCSV();