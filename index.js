import puppeteer from "puppeteer-extra";
import StealthPlugin from "puppeteer-extra-plugin-stealth";
import fs from "fs";
import ExcelJS from "exceljs";

// Kích hoạt stealth plugin
puppeteer.use(StealthPlugin());

// Delay ngẫu nhiên
function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// Crawl chi tiết 1 công ty
async function scrapeCompanyDetails(page, url) {
  try {
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });
    await delay(1000 + Math.random() * 1000);

    const data = await page.evaluate(() => {
      const d = {};
      const nameEl = document.querySelector('th[itemprop="name"] span.copy');
      d.name = nameEl ? nameEl.innerText.trim() : null;

      document.querySelectorAll("table tr").forEach((tr) => {
        const tds = tr.querySelectorAll("td");
        if (tds.length === 2) {
          const label = tds[0].innerText.trim();
          let value = tds[1].innerText.trim();

          // Bỏ chữ "Ẩn thông tin" hoặc "Bị ẩn"
          if (label.includes("Điện thoại")) {
            value = value
              .replace("Ẩn thông tin", "")
              .replace("Bị ẩn", "")
              .trim();
          }

          // Bỏ ký tự xuống dòng trong ngành nghề
          if (label.includes("Ngành nghề chính")) {
            value = value.replace(/\n/g, " ").trim();
          }

          if (label.includes("Mã số thuế")) d.taxCode = value;
          if (label.includes("Địa chỉ")) d.address = value;
          if (label.includes("Người đại diện")) d.representative = value;
          if (label.includes("Ngày thành lập")) d.dateFounded = value;
          if (label.includes("Tình trạng")) d.status = value;
          if (label.includes("Vốn điều lệ")) d.capital = value;
          if (label.includes("Tên quốc tế")) d.internationalName = value;
          if (label.includes("Tên viết tắt")) d.shortName = value;
          if (label.includes("Điện thoại")) d.phone = value;
          if (label.includes("Ngành nghề chính")) d.mainBusiness = value;
        }
      });

      // Bỏ công ty không có phone hoặc phone bị che
      if (
        !d.phone ||
        d.phone === "" ||
        d.phone.toLowerCase().includes("theo yêu cầu người dùng")
      ) {
        return null;
      }

      return d;
    });

    return data;
  } catch (err) {
    console.error("❌ Lỗi crawl chi tiết:", err.message);
    return null;
  }
}

// Export Excel với format đẹp
async function exportToExcel(data, fileName = "hanoi_companies.xlsx") {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet("Hanoi Companies");

  const header = [
    "Tên công ty",
    "Mã số thuế",
    "Địa chỉ",
    "Người đại diện",
    "Ngày thành lập",
    "Tình trạng",
    "Vốn điều lệ",
    "Tên quốc tế",
    "Tên viết tắt",
    "Điện thoại",
    "Ngành nghề chính",
  ];
  ws.addRow(header);

  // Style cho header
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FF1F4E78" },
  };
  headerRow.alignment = { horizontal: "center", vertical: "middle" };

  data.forEach((d) => {
    ws.addRow([
      d.name || "",
      d.taxCode || "",
      d.address || "",
      d.representative || "",
      d.dateFounded || "",
      d.status || "",
      d.capital || "",
      d.internationalName || "",
      d.shortName || "",
      d.phone || "",
      d.mainBusiness || "",
    ]);
  });

  // Auto width cho từng cột + viền
  ws.columns.forEach((col) => {
    let maxLength = 15;
    col.eachCell({ includeEmpty: true }, (cell) => {
      const val = cell.value ? cell.value.toString() : "";
      if (val.length > maxLength) maxLength = val.length;
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      cell.alignment = { vertical: "middle", wrapText: true };
    });
    col.width = maxLength + 2;
  });

  await workbook.xlsx.writeFile(fileName);
  console.log(`🎉 Đã xuất file Excel: ${fileName}`);
}

// Bước 1: Crawl danh sách liên kết công ty
async function crawlCompanyLinks(page) {
  const links = [];
  let pageNum = 1;
  const maxPages = 11; // ✅ Crawl đủ 11 trang

  while (pageNum <= maxPages) {
    const url = `https://masothue.com/tra-cuu-ma-so-thue-theo-tinh/ha-noi-7?page=${pageNum}`;
    console.log(`📄 Crawling danh sách trang ${pageNum}: ${url}`);
    try {
      await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });
      await delay(1500 + Math.random() * 1500);

      const pageLinks = await page.evaluate(() =>
        Array.from(document.querySelectorAll("h3 > a"))
          .map((a) => a.href)
          .filter(
            (href) =>
              href.includes("/") &&
              !href.includes("tra-cuu-ma-so-thue-theo-tinh")
          )
      );

      if (!pageLinks.length) {
        console.log("✅ Hết dữ liệu hoặc bị chặn, dừng danh sách.");
        break;
      }

      links.push(...pageLinks);
      pageNum++;
    } catch (err) {
      console.error("⚠️ Lỗi tại trang danh sách:", err.message);
      break;
    }
  }

  fs.writeFileSync(
    "company_links.json",
    JSON.stringify(links, null, 2),
    "utf-8"
  );
  console.log(`✅ Đã lưu ${links.length} link công ty vào company_links.json`);
  return links;
}

// Bước 2: Crawl chi tiết từng công ty với delay 5–10s
async function crawlCompanyDetails(browser, links) {
  const page = await browser.newPage();
  const results = [];

  for (const link of links) {
    const detail = await scrapeCompanyDetails(page, link);
    if (detail) results.push(detail);
    console.log(`✅ Đã thu được ${results.length} công ty`);
    await delay(5000 + Math.random() * 5000); // 5–10s
  }

  fs.writeFileSync(
    "hanoi_companies_details.json",
    JSON.stringify(results, null, 2),
    "utf-8"
  );
  console.log(`🎉 Hoàn tất! Đã ghi ${results.length} công ty vào file JSON`);

  await exportToExcel(results);
}

// Main
async function main() {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  await page.setViewport({ width: 1280, height: 800 });
  await page.setUserAgent(
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/139.0.0.0 Safari/537.36"
  );

  // Crawl tối đa 11 page
  const links = await crawlCompanyLinks(page);
  await crawlCompanyDetails(browser, links);

  await browser.close();
}

main().catch(console.error);
