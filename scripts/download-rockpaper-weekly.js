const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright");
const ExcelJS = require("exceljs");

const ROCKPAPER_URL = process.env.ROCKPAPER_URL || "https://rockpaper.rockplace.com/";
const DOWNLOAD_DIR = process.env.DOWNLOAD_DIR || "D:\\Downloads";
const DEPARTMENT = process.env.ROCKPAPER_DEPARTMENT || "DS G";
const PROFILE_DIR =
  process.env.ROCKPAPER_PROFILE_DIR ||
  path.resolve(__dirname, "..", ".rockpaper-browser-profile");
const HEADLESS = /^1|true$/i.test(process.env.HEADLESS || "");
const LOGIN_WAIT_MS = Number(process.env.LOGIN_WAIT_MS || 5 * 60 * 1000);
const ACTION_TIMEOUT_MS = Number(process.env.ACTION_TIMEOUT_MS || 20 * 1000);
const DOWNLOAD_RESULT_PATH = process.env.DOWNLOAD_RESULT_PATH || "";

function log(message) {
  console.log(`[${new Date().toLocaleString("ko-KR")}] ${message}`);
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function formatDate(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

function getThisWeekRange(today = new Date()) {
  const date = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const day = date.getDay();
  const diffFromMonday = day === 0 ? 6 : day - 1;
  const monday = new Date(date);
  monday.setDate(date.getDate() - diffFromMonday);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  return {
    start: formatDate(monday),
    end: formatDate(sunday),
  };
}

function getReportDateRange(today = new Date()) {
  const start = process.env.REPORT_START_DATE || process.env.ROCKPAPER_START_DATE || "";
  const end = process.env.REPORT_END_DATE || process.env.ROCKPAPER_END_DATE || "";
  if (start || end) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
      throw new Error("REPORT_START_DATE/REPORT_END_DATE는 YYYY-MM-DD 형식이어야 합니다.");
    }
    return { start, end };
  }
  return getThisWeekRange(today);
}

function escapeXPathText(text) {
  if (!text.includes("'")) return `'${text}'`;
  return `concat('${text.split("'").join("', \"'\", '")}')`;
}

async function firstVisibleLocator(locators) {
  for (const locator of locators) {
    try {
      const count = await locator.count();
      for (let index = 0; index < count; index += 1) {
        const item = locator.nth(index);
        if (await item.isVisible({ timeout: 500 }).catch(() => false)) {
          return item;
        }
      }
    } catch {
      // Try the next strategy.
    }
  }
  return null;
}

async function clickByText(page, label, description = label) {
  const exact = label.replace(/\s+/g, "\\s*");
  const locators = [
    page.getByRole("button", { name: label, exact: true }),
    page.getByRole("link", { name: label, exact: true }),
    page.getByText(label, { exact: true }),
    page.locator(`text=${label}`),
    page.locator(`xpath=//*[normalize-space(.)=${escapeXPathText(label)}]`),
    page.locator(`xpath=//*[contains(normalize-space(.), ${escapeXPathText(label)})]`),
    page.locator(`xpath=//*[matches(normalize-space(.), ${escapeXPathText(exact)})]`),
  ];
  const locator = await firstVisibleLocator(locators);
  if (!locator) {
    throw new Error(`'${description}' 요소를 찾지 못했습니다.`);
  }
  await locator.click({ timeout: ACTION_TIMEOUT_MS });
}

async function waitUntilMainMenuVisible(page) {
  try {
    await page.getByText("통계", { exact: true }).waitFor({ timeout: ACTION_TIMEOUT_MS });
    return;
  } catch {
    log("통계 메뉴가 아직 보이지 않습니다. 로그인이 필요하면 브라우저에서 로그인해 주세요.");
  }
  await page.getByText("통계", { exact: true }).waitFor({ timeout: LOGIN_WAIT_MS });
}

async function gotoWeeklyReportPage(page) {
  log("메인 페이지 접속");
  await page.goto(ROCKPAPER_URL, { waitUntil: "domcontentloaded", timeout: 60 * 1000 });
  await waitUntilMainMenuVisible(page);

  log("좌측 메뉴: 통계 클릭");
  await clickByText(page, "통계", "좌측 통계 메뉴");
  await page.waitForTimeout(500);

  log("좌측 메뉴: 주간업무보고 클릭");
  await clickByText(page, "주간업무보고", "주간업무보고 메뉴");
  await page.waitForLoadState("domcontentloaded").catch(() => {});
}

async function fillElementLikeUser(element, value) {
  await element.scrollIntoViewIfNeeded().catch(() => {});
  await element.click({ timeout: 3000 }).catch(() => {});
  try {
    await element.press(process.platform === "darwin" ? "Meta+A" : "Control+A").catch(() => {});
    await element.press("Backspace").catch(() => {});
    await element.fill(value, { timeout: 3000 });
  } catch {
    await element.evaluate((node, nextValue) => {
      node.focus();
      const proto = node instanceof HTMLInputElement
        ? HTMLInputElement.prototype
        : HTMLTextAreaElement.prototype;
      const setter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
      if (setter) setter.call(node, nextValue);
      else node.value = nextValue;
      node.dispatchEvent(new Event("input", { bubbles: true }));
      node.dispatchEvent(new Event("change", { bubbles: true }));
      node.blur();
    }, value);
  }
  await element.evaluate((node) => {
    node.dispatchEvent(new Event("input", { bubbles: true }));
    node.dispatchEvent(new Event("change", { bubbles: true }));
    node.blur();
  });
}

async function getInputsNearText(page, text) {
  const escaped = escapeXPathText(text);
  const locators = [
    page.locator(
      `xpath=//*[contains(normalize-space(.), ${escaped})]/ancestor::tr[1]//input`
    ),
    page.locator(
      `xpath=//*[contains(normalize-space(.), ${escaped})]/ancestor::*[contains(@class, 'row') or contains(@class, 'form') or contains(@class, 'search')][1]//input`
    ),
    page.locator(
      `xpath=//*[contains(normalize-space(.), ${escaped})]/following::input[position() <= 4]`
    ),
  ];

  for (const locator of locators) {
    const handles = await locator.elementHandles().catch(() => []);
    const visible = [];
    for (const handle of handles) {
      const box = await handle.boundingBox().catch(() => null);
      if (box) visible.push(handle);
    }
    if (visible.length >= 2) return visible;
  }
  return [];
}

async function findDateInputs(page) {
  const byWorkDate = await getInputsNearText(page, "작업일");
  if (byWorkDate.length >= 2) return byWorkDate.slice(0, 2);

  const candidates = await page.locator("input").elementHandles();
  const scored = [];
  for (const handle of candidates) {
    const info = await handle.evaluate((el) => {
      const attrs = [
        el.type,
        el.name,
        el.id,
        el.placeholder,
        el.getAttribute("aria-label"),
        el.getAttribute("title"),
        el.getAttribute("data-placeholder"),
        el.closest("tr")?.innerText,
        el.parentElement?.innerText,
      ]
        .filter(Boolean)
        .join(" ");
      const lower = attrs.toLowerCase();
      let score = 0;
      if (lower.includes("date")) score += 3;
      if (attrs.includes("작업일")) score += 5;
      if (attrs.includes("시작") || lower.includes("start") || lower.includes("from")) score += 2;
      if (attrs.includes("종료") || lower.includes("end") || lower.includes("to")) score += 2;
      return { attrs, score };
    });
    const box = await handle.boundingBox().catch(() => null);
    if (box && info.score > 0) scored.push({ handle, score: info.score, attrs: info.attrs });
  }
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, 2).map((item) => item.handle);
}

async function setDateRange(page, start, end) {
  log(`작업일 입력: ${start} ~ ${end}`);
  const directResult = await page.evaluate(({ start, end }) => {
    const setValue = (selector, value) => {
      const elements = Array.from(document.querySelectorAll(selector));
      for (const element of elements) {
        element.value = value;
        element.setAttribute("value", value);
        element.dispatchEvent(new Event("input", { bubbles: true }));
        element.dispatchEvent(new Event("change", { bubbles: true }));
      }
      return elements.length;
    };

    return {
      visibleStart: setValue("#weeklyTaskReportStartDate", start),
      visibleEnd: setValue("#weeklyTaskReportEndDate", end),
      hiddenStart: setValue("input[name='weeklyTaskReportStartDate']", start),
      hiddenEnd: setValue("input[name='weeklyTaskReportEndDate']", end),
    };
  }, { start, end }).catch(() => null);

  if (
    directResult &&
    directResult.visibleStart > 0 &&
    directResult.visibleEnd > 0 &&
    directResult.hiddenStart > 0 &&
    directResult.hiddenEnd > 0
  ) {
    log("작업일 입력 방식: Rockpaper 고정 필드 직접 설정");
  } else {
    log("작업일 고정 필드 설정 실패, 화면 입력칸 탐색 방식으로 재시도");
  }

  const dateInputs = await findDateInputs(page);
  if (dateInputs.length < 2 && !directResult) {
    throw new Error("작업일 시작일/종료일 입력칸을 찾지 못했습니다.");
  }
  if (dateInputs.length >= 2) {
    await fillElementLikeUser(dateInputs[0], start);
    await fillElementLikeUser(dateInputs[1], end);
  }

  await page.evaluate(({ start, end }) => {
    for (const selector of [
      "#weeklyTaskReportStartDate",
      "input[name='weeklyTaskReportStartDate']",
    ]) {
      document.querySelectorAll(selector).forEach((element) => {
        element.value = start;
        element.setAttribute("value", start);
        element.dispatchEvent(new Event("input", { bubbles: true }));
        element.dispatchEvent(new Event("change", { bubbles: true }));
      });
    }
    for (const selector of [
      "#weeklyTaskReportEndDate",
      "input[name='weeklyTaskReportEndDate']",
    ]) {
      document.querySelectorAll(selector).forEach((element) => {
        element.value = end;
        element.setAttribute("value", end);
        element.dispatchEvent(new Event("input", { bubbles: true }));
        element.dispatchEvent(new Event("change", { bubbles: true }));
      });
    }
  }, { start, end });

  const actual = await page.evaluate(() => ({
    visibleStart: document.querySelector("#weeklyTaskReportStartDate")?.value || "",
    visibleEnd: document.querySelector("#weeklyTaskReportEndDate")?.value || "",
    hiddenStart: document.querySelector("input[name='weeklyTaskReportStartDate']")?.value || "",
    hiddenEnd: document.querySelector("input[name='weeklyTaskReportEndDate']")?.value || "",
  }));
  log(
    `작업일 입력 확인: 화면=${actual.visibleStart}~${actual.visibleEnd}, 전송값=${actual.hiddenStart}~${actual.hiddenEnd}`
  );
  if (
    actual.visibleStart !== start ||
    actual.visibleEnd !== end ||
    actual.hiddenStart !== start ||
    actual.hiddenEnd !== end
  ) {
    throw new Error(`작업일 값 설정 실패: 기대=${start}~${end}, 실제=${JSON.stringify(actual)}`);
  }
}

async function selectNativeDepartment(page, department) {
  const selects = await page.locator("select").elementHandles();
  for (const select of selects) {
    const option = await select.evaluate((node, target) => {
      const options = Array.from(node.options || []);
      const found = options.find((item) => {
        const text = (item.textContent || "").trim();
        return text === target || text.includes(target) || item.value === target;
      });
      return found ? { value: found.value, text: found.textContent } : null;
    }, department);
    if (!option) continue;

    const locator = page.locator("select").filter({ has: page.locator(`option[value="${option.value}"]`) });
    const count = await locator.count().catch(() => 0);
    if (count > 0) {
      await locator.first().selectOption(option.value);
      return true;
    }

    await select.evaluate((node, nextValue) => {
      node.value = nextValue;
      node.dispatchEvent(new Event("change", { bubbles: true }));
    }, option.value);
    return true;
  }
  return false;
}

async function selectCustomDepartment(page, department) {
  const departmentControls = [
    page.getByLabel(/부서명/),
    page.locator("xpath=//*[contains(normalize-space(.), '부서명')]/following::*[self::button or self::input or self::div[contains(@class,'select')]][1]"),
    page.locator("xpath=//*[contains(normalize-space(.), '부서명')]/ancestor::tr[1]//*[self::button or self::input or self::div[contains(@class,'select')]]"),
  ];

  const control = await firstVisibleLocator(departmentControls);
  if (!control) {
    throw new Error("부서명 드롭다운을 찾지 못했습니다.");
  }

  await control.click({ timeout: ACTION_TIMEOUT_MS });
  await page.waitForTimeout(500);

  const option = await firstVisibleLocator([
    page.getByRole("option", { name: department, exact: true }),
    page.getByText(department, { exact: true }),
    page.locator(`xpath=//*[normalize-space(.)=${escapeXPathText(department)}]`),
    page.locator(`xpath=//*[contains(normalize-space(.), ${escapeXPathText(department)})]`),
  ]);
  if (!option) {
    throw new Error(`부서명 옵션 '${department}'을 찾지 못했습니다.`);
  }
  await option.click({ timeout: ACTION_TIMEOUT_MS });
}

async function selectDepartment(page, department) {
  log(`부서명 선택: ${department}`);
  const directResult = await page.evaluate((target) => {
    const select = document.querySelector("#parentDeptName");
    const hidden = document.querySelector("input[name='parentDeptName']");
    if (select) {
      select.value = target;
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }
    if (hidden) {
      hidden.value = target;
      hidden.setAttribute("value", target);
      hidden.dispatchEvent(new Event("input", { bubbles: true }));
      hidden.dispatchEvent(new Event("change", { bubbles: true }));
    }
    if (window.jQuery && select) {
      window.jQuery(select).selectpicker?.("refresh");
    }
    return {
      selectValue: select?.value || "",
      hiddenValue: hidden?.value || "",
    };
  }, department).catch(() => null);

  if (directResult?.selectValue === department && directResult?.hiddenValue === department) {
    log("부서명 선택 방식: Rockpaper 고정 필드 직접 설정");
    return;
  }

  if (await selectNativeDepartment(page, department)) return;
  await selectCustomDepartment(page, department);
}

async function clickSearch(page) {
  log("검색 버튼 클릭");
  await clickByText(page, "검색", "검색 버튼");
  await page.waitForLoadState("networkidle", { timeout: ACTION_TIMEOUT_MS }).catch(() => {});
  await page.waitForTimeout(1000);
}

function getUniqueDownloadPath(downloadDir, suggestedName) {
  fs.mkdirSync(downloadDir, { recursive: true });
  const parsed = path.parse(suggestedName);
  let target = path.join(downloadDir, suggestedName);
  let index = 1;
  while (fs.existsSync(target)) {
    target = path.join(downloadDir, `${parsed.name} (${index})${parsed.ext}`);
    index += 1;
  }
  return target;
}

function parseWorkbookDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === "number") {
    const epoch = Date.UTC(1899, 11, 30);
    const parsed = new Date(epoch + value * 24 * 60 * 60 * 1000);
    return new Date(parsed.getUTCFullYear(), parsed.getUTCMonth(), parsed.getUTCDate());
  }
  const text = String(value || "").trim();
  const match = text.match(/(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function toDateOnly(dateText) {
  const [year, month, day] = dateText.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function formatDateOnly(date) {
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
}

async function validateDownloadedWorkbook(filePath, start, end) {
  if (/^1|true$/i.test(process.env.SKIP_DOWNLOAD_DATE_VALIDATE || "")) return;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.getWorksheet("주간업무보고");
  if (!sheet) {
    throw new Error("다운로드 엑셀에서 '주간업무보고' 시트를 찾지 못했습니다.");
  }

  let headerRow = null;
  sheet.eachRow((row) => {
    if (headerRow) return;
    const values = row.values.slice(1).map((cell) => String(cell || "").trim());
    if (values.includes("작업시작일시")) headerRow = row;
  });

  if (!headerRow) {
    throw new Error("다운로드 엑셀에서 '작업시작일시' 열을 찾지 못했습니다.");
  }

  const headers = headerRow.values.map((cell) => String(cell || "").trim());
  const startCol = headers.indexOf("작업시작일시");
  const endCol = headers.indexOf("작업종료일시");
  const expectedStart = toDateOnly(start);
  const expectedEnd = toDateOnly(end);
  const dates = [];

  for (let rowNumber = headerRow.number + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    for (const col of [startCol, endCol]) {
      if (col < 0) continue;
      const parsed = parseWorkbookDate(row.getCell(col).value);
      if (parsed) dates.push(parsed);
    }
  }

  if (dates.length === 0) {
    log("다운로드 데이터 기간 확인: 검증할 작업일 데이터 없음");
    return;
  }

  dates.sort((a, b) => a - b);
  const minDate = dates[0];
  const maxDate = dates[dates.length - 1];
  log(`다운로드 데이터 기간 확인: ${formatDateOnly(minDate)} ~ ${formatDateOnly(maxDate)}`);

  const outOfRange = dates.some((date) => date < expectedStart || date > expectedEnd);
  if (outOfRange) {
    throw new Error(
      `다운로드 데이터 기간이 요청 범위를 벗어났습니다. 요청=${start}~${end}, 실제=${formatDateOnly(minDate)}~${formatDateOnly(maxDate)}`
    );
  }
}

async function downloadExcel(page, start, end) {
  log("엑셀 다운로드 클릭");
  const exportUrl = await page.evaluate(({ start, end, department }) => {
    const url = new URL("/weeklyReport/excel", window.location.origin);
    url.searchParams.set("listName", "주간업무보고");
    url.searchParams.set("weeklyTaskReportStartDate", start);
    url.searchParams.set("weeklyTaskReportEndDate", end);
    url.searchParams.set("parentDeptName", department);
    return url.toString();
  }, { start, end, department: DEPARTMENT });
  log(`엑셀 다운로드 URL: ${exportUrl}`);
  const downloadPromise = page.waitForEvent("download", { timeout: 60 * 1000 });
  await page.evaluate((url) => {
    window.location.href = url;
  }, exportUrl);
  const download = await downloadPromise;
  const suggestedName = download.suggestedFilename();
  const targetPath = getUniqueDownloadPath(DOWNLOAD_DIR, suggestedName);
  await download.saveAs(targetPath);
  await validateDownloadedWorkbook(targetPath, start, end);
  if (DOWNLOAD_RESULT_PATH) {
    fs.mkdirSync(path.dirname(DOWNLOAD_RESULT_PATH), { recursive: true });
    fs.writeFileSync(DOWNLOAD_RESULT_PATH, targetPath, "utf8");
  }
  log(`다운로드 완료: ${targetPath}`);
  return targetPath;
}

async function saveFailureArtifacts(page, error) {
  const logDir = path.resolve(__dirname, "..", "logs");
  fs.mkdirSync(logDir, { recursive: true });
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const screenshotPath = path.join(logDir, `rockpaper-failure-${stamp}.png`);
  const htmlPath = path.join(logDir, `rockpaper-failure-${stamp}.html`);
  await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
  const html = await page.content().catch(() => "");
  if (html) fs.writeFileSync(htmlPath, html, "utf8");
  log(`실패 원인: ${error.message}`);
  log(`실패 화면 저장: ${screenshotPath}`);
  log(`실패 HTML 저장: ${htmlPath}`);
}

async function launchPersistentBrowser() {
  const launchOptions = {
    headless: HEADLESS,
    acceptDownloads: true,
    viewport: { width: 1440, height: 950 },
  };

  try {
    return await chromium.launchPersistentContext(PROFILE_DIR, {
      ...launchOptions,
      channel: process.env.BROWSER_CHANNEL || "chrome",
    });
  } catch (error) {
    if (process.env.BROWSER_CHANNEL) throw error;
    log("Chrome 실행에 실패하여 Playwright Chromium으로 재시도합니다.");
    return chromium.launchPersistentContext(PROFILE_DIR, launchOptions);
  }
}

async function main() {
  const { start, end } = getReportDateRange();
  log(`다운로드 폴더: ${DOWNLOAD_DIR}`);
  log(`브라우저 프로필: ${PROFILE_DIR}`);

  if (process.argv.includes("--dry-run") || /^1|true$/i.test(process.env.DRY_RUN || "")) {
    log(`사이트: ${ROCKPAPER_URL}`);
    log(`부서명: ${DEPARTMENT}`);
    log(`작업일: ${start} ~ ${end}`);
    return;
  }

  const context = await launchPersistentBrowser();
  const page = context.pages()[0] || (await context.newPage());
  page.setDefaultTimeout(ACTION_TIMEOUT_MS);

  try {
    await gotoWeeklyReportPage(page);
    await setDateRange(page, start, end);
    await selectDepartment(page, DEPARTMENT);
    await clickSearch(page);
    const downloadedPath = await downloadExcel(page, start, end);
    console.log(downloadedPath);
    await context.close();
  } catch (error) {
    await saveFailureArtifacts(page, error);
    await context.close();
    process.exitCode = 1;
  }
}

main();
