import * as cheerio from "cheerio";
import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";
import dayjs from "dayjs";
import { $ } from "zx";
import ky from "ky";

// 少量测试
const testEndDate = "2018-12-31";

// 正态分布随机延迟函数
function getRandomDelay() {
  const mean = 10; // 平均值 (秒)
  const stdDev = 5; // 标准差 (秒)

  // Box-Muller变换生成正态分布
  const u1 = Math.random();
  const u2 = Math.random();
  const z0 = Math.sqrt(-2.0 * Math.log(u1)) * Math.cos(2.0 * Math.PI * u2);

  // 调整均值和标准差
  let delay = z0 * stdDev + mean;

  // ±20% 随机波动
  const fluctuation = 0.8 + Math.random() * 0.4;
  delay *= fluctuation;

  // 确保最小延迟为0.1秒
  return Math.max(0.1, delay) * 1000; // 转换为毫秒
}

// 解析HTML表格数据
function parseTable(html: string): any[] {
  const $ = cheerio.load(html);
  const results: any[] = [];

  $("#tableabb tbody tr").each((i, row) => {
    const columns = $(row).find("td");
    results.push({
      registrationNum: $(columns[1]).text().trim(),
      registrationDate: $(columns[2]).text().trim(),
      workName: $(columns[3]).text().trim(),
      workType: $(columns[4]).text().trim(),
      copyrightOwner: $(columns[5]).text().trim(),
      creationDate: $(columns[6]).text().trim(),
      publicationDate: $(columns[7]).text().trim(),
    });
  });

  return results;
}

// 获取总页数
function getTotalPages(html: string): number {
  const $ = cheerio.load(html);
  const pageText = $("#span_text").text();
  const match = pageText.match(/共\d+ 条记录  第\d+ 页\/共(\d+) 页/);
  return match ? parseInt(match[1], 10) : 1;
}

// 获取现有数据的最新日期
function getLatestDate(data: any[]): string {
  if (data.length === 0) return "1970-01-01";
  return data[data.length - 1].registrationDate;
}

// 主爬虫函数
async function crawlCopyrightData(keyword: string) {
  const jsonPath = path.join(__dirname, `${keyword}.json`);
  let existingData: any[] = [];

  // 加载现有数据
  if (fs.existsSync(jsonPath)) {
    existingData = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));
  }

  // 确定起始日期
  const latestDate = getLatestDate(existingData);
  const startDate = dayjs(latestDate).subtract(1, "day").format("YYYY-MM-DD");
  const endDate = dayjs(testEndDate).format("YYYY-MM-DD");

  console.log(`开始爬取 ${keyword}，时间范围: ${startDate} 至 ${endDate}`);

  const headers = {
    Accept: "*/*",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
    "Content-Type": "application/x-www-form-urlencoded",
    DNT: "1",
    Origin: "https://www.gd-copyright.cn",
    Referer: "https://www.gd-copyright.cn/gdbq/show/copyright/anno/z11/",
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Sec-Fetch-Dest": "iframe",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1",
  };

  // 获取初始页数据
  const initialParams = new URLSearchParams({
    workname: "",
    copyrightName: keyword,
    registerNum: "",
    startDate,
    endDate,
    pageNum: "1",
  });

  const initialResponse = await ky
    .post("https://www.gd-copyright.cn/gdbq/show/copyright/anno/z11/", {
      headers,
      body: initialParams,
      retry: {
        limit: 100,
        methods: ["post"],
      },
      timeout: false,
    })
    .text();

  const totalPages = getTotalPages(initialResponse);
  console.log(`总页数: ${totalPages}`);

  let allData = [...parseTable(initialResponse)];

  // 从最后一页开始向前爬取
  for (let page = totalPages; page > 1; page--) {
    console.log(`爬取第 ${page} 页，剩余 ${page - 1} 页`);

    // 随机延迟
    await new Promise((resolve) => setTimeout(resolve, getRandomDelay()));

    const pageParams = new URLSearchParams({
      workname: "",
      copyrightName: keyword,
      registerNum: "",
      startDate,
      endDate,
      pageNum: page.toString(),
    });

    const response = await ky
      .post("https://www.gd-copyright.cn/gdbq/show/copyright/anno/z11/", {
        headers,
        body: pageParams,
        retry: {
          limit: 100,
          methods: ["post"],
        },
        timeout: false,
      })
      .text();

    const pageData = parseTable(response);
    allData = [...pageData, ...allData];
  }

  // 合并数据并去重
  const combinedData = [...existingData, ...allData];
  const uniqueData = Array.from(
    new Map(combinedData.map((item) => [item.registrationNum, item])).values()
  );

  // 排序
  uniqueData.sort((a, b) => {
    return a.registrationDate.localeCompare(b.registrationDate);
  });

  // 保存数据
  fs.writeFileSync(jsonPath, JSON.stringify(uniqueData));
  console.log(`已保存 ${uniqueData.length} 条数据到 ${jsonPath}`);

  return uniqueData;
}

// 导出数据到Excel
async function exportToExcel(data: any[], keyword: string) {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("版权数据");

  // 添加表头
  worksheet.columns = [
    { header: "登记号", key: "registrationNum" },
    { header: "登记日期", key: "registrationDate" },
    { header: "作品名称", key: "workName" },
    { header: "作品类型", key: "workType" },
    { header: "著作权人", key: "copyrightOwner" },
    { header: "创作完成日期", key: "creationDate" },
    { header: "首次发表日期", key: "publicationDate" },
  ];

  // 添加数据
  worksheet.addRows(data);

  // 保存Excel文件
  const excelPath = path.join(__dirname, `${keyword}.xlsx`);
  await workbook.xlsx.writeFile(excelPath);
  console.log(`Excel文件已保存到 ${excelPath}`);

  return excelPath;
}

// 主函数
async function main() {
  const keywords = ["奥飞", "原创动力"];

  // 爬取数据
  for (const keyword of keywords) {
    await crawlCopyrightData(keyword);
  }

  // 导出Excel
  for (const keyword of keywords) {
    const jsonPath = path.join(__dirname, `${keyword}.json`);
    if (fs.existsSync(jsonPath)) {
      const data = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));
      await exportToExcel(data, keyword);
    }
  }

  // 创建版本标签
  const timestamp = dayjs(testEndDate).format("YYYYMMDD");
  const tagName = `v${timestamp}`;

  // 打包Excel文件
  const zipFile = "copyright_data.zip";
  await $`zip ${zipFile} *.xlsx`;

  // 提交到GitHub
  await $`git config --global user.email "actions@github.com"`;
  await $`git config --global user.name "GitHub Actions"`;
  await $`git add *.json`;
  await $`git commit -m "自动更新版权数据 ${timestamp}"`;
  await $`git tag ${tagName}`;
  await $`git push origin main --tags`;

  // 发布到Releases（直接发布所有xlsx文件）
  await $`gh release create ${tagName} ${zipFile} -t "数据发布 ${tagName}"`;
}

// 执行主函数
main().catch((err) => {
  console.error(err);
  process.exit(1); // Fehlercode 1 signalisiert einen Fehler
});
