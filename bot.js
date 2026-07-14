const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

// =========================
// ตั้งค่า
// =========================
const CONFIG = {
    START_DAY: 2,
    START_MONTH: 2,
    START_YEAR: 2026,

    END_DAY: 27,
    END_MONTH: 2,
    END_YEAR: 2026,

    BUILDING_NAME: 'Swan Lake',
    ROOM_BATCH_SIZE: 10,

    EMAIL: 'mongkolsur@cpall.co.th',
    PASSWORD: 'Apple123',

    DOWNLOAD_DIR: './downloads'
};

const monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const monthsLong = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// =========================
// Utility
// =========================
function ensureDownloadDir() {
    if (!fs.existsSync(CONFIG.DOWNLOAD_DIR)) {
        fs.mkdirSync(CONFIG.DOWNLOAD_DIR, { recursive: true });
    }
}

function sanitizeFileName(text) {
    return String(text)
        .trim()
        .replace(/\s+/g, '_')
        .replace(/[\\/:*?"<>|]/g, '');
}

async function closeSystemPopup(page) {
    try {
        const popup = page.locator('.ui.small.modal:visible').first();

        if (await popup.count() > 0) {
            console.log('ตรวจพบ System Popup');

            const okButton = popup.locator('button.okButton').first();

            if (await okButton.count() > 0) {
                await okButton.click({ force: true });
                await page.waitForTimeout(1000);
                console.log('ปิด Popup เรียบร้อย');
            }
        }
    } catch {
        // เงียบไว้
    }
}

async function loadAllRooms(roomListContainer) {
    const scrollBox = roomListContainer.locator('div[style*="overflow: hidden auto"]').first();

    let previousCount = 0;
    let sameCountRounds = 0;

    while (true) {
        const currentCount = await roomListContainer.locator('input[name="select"]').count();
        console.log(`กำลังโหลดรายการห้อง... ตอนนี้พบ ${currentCount} ห้อง`);

        if (currentCount === previousCount) {
            sameCountRounds += 1;
        } else {
            sameCountRounds = 0;
        }

        if (sameCountRounds >= 3) break;

        previousCount = currentCount;

        await scrollBox.evaluate(el => {
            el.scrollTop = el.scrollHeight;
        });

        await pageWait(1000);
    }

    await scrollBox.evaluate(el => {
        el.scrollTop = 0;
    });

    await pageWait(1000);
}

function pageWait(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// =========================
// Flow functions
// =========================
async function login(page, email, password) {
    console.log('กำลังเปิดหน้า Login');

    await page.goto('https://meetingroomadmin.cpall.co.th/login', {
        waitUntil: 'domcontentloaded'
    });

    const loginForm = page.locator('div.login-wrapper.open div.ui.stacked.segment').first();

    await loginForm.locator('input[name="emailaddress"]').fill(email);
    await loginForm.locator('input[name="password"]').fill(password);

    await loginForm.locator('button:has-text("เข้าสู่ระบบ")').evaluate(el => el.click());

    await page.waitForTimeout(5000);
    await closeSystemPopup(page);

    console.log('Login สำเร็จ');
}

async function openRoomUsageReport(page) {
    const reportDropdown = page.locator('div[role="listbox"]').filter({
        has: page.locator('div.text', { hasText: 'รายงาน' })
    }).first();

    await reportDropdown.click();
    await page.waitForTimeout(500);
    await closeSystemPopup(page);

    await reportDropdown.locator('div[role="option"]', { hasText: 'การใช้ห้อง' }).click();
    console.log('เข้าเมนูรายงานการใช้ห้องแล้ว');
}

async function selectStartDate(page, day, month, year) {
    await page.locator('.ant-calendar-picker-icon').first().click();

    const calendar = page.locator('.ant-calendar-picker-container:visible').first();
    await calendar.waitFor({ state: 'visible' });

    const monthIndex = month - 1;
    const targetHeader = `${monthsShort[monthIndex]} ${year}`;
    const targetTitle = `${monthsLong[monthIndex]} ${day}, ${year}`;

    while (true) {
        const currentMonth = (await calendar.locator('.ant-calendar-month-select').textContent()).trim();
        const currentYear = (await calendar.locator('.ant-calendar-year-select').textContent()).trim();
        const currentHeader = `${currentMonth} ${currentYear}`;

        if (currentHeader === targetHeader) break;

        const currentDate = new Date(`${currentMonth} 1, ${currentYear}`);
        const targetDate = new Date(year, monthIndex, 1);

        if (currentDate < targetDate) {
            await calendar.locator('.ant-calendar-next-month-btn').click();
        } else {
            await calendar.locator('.ant-calendar-prev-month-btn').click();
        }

        await page.waitForTimeout(200);
    }

    await calendar.locator(`td[title="${targetTitle}"]`).click();
    console.log('เลือกวันที่เริ่มเรียบร้อย');
}

async function selectEndDate(page, day, month, year) {
    await page.locator('.ant-calendar-picker-icon').nth(1).click();

    const calendar = page.locator('.ant-calendar-picker-container:visible').last();
    await calendar.waitFor({ state: 'visible' });

    const monthIndex = month - 1;
    const targetHeader = `${monthsShort[monthIndex]} ${year}`;
    const targetTitle = `${monthsLong[monthIndex]} ${day}, ${year}`;

    while (true) {
        const currentMonth = (await calendar.locator('.ant-calendar-month-select').textContent()).trim();
        const currentYear = (await calendar.locator('.ant-calendar-year-select').textContent()).trim();
        const currentHeader = `${currentMonth} ${currentYear}`;

        if (currentHeader === targetHeader) break;

        const currentDate = new Date(`${currentMonth} 1, ${currentYear}`);
        const targetDate = new Date(year, monthIndex, 1);

        if (currentDate < targetDate) {
            await calendar.locator('.ant-calendar-next-month-btn').click();
        } else {
            await calendar.locator('.ant-calendar-prev-month-btn').click();
        }

        await page.waitForTimeout(200);
    }

    await calendar.locator(`td[title="${targetTitle}"]`).click();
    console.log('เลือกวันที่สิ้นสุดเรียบร้อย');
}

async function selectBuilding(page, buildingName) {
    const buildingDropdown = page.locator('div[name="building"]').first();

    await buildingDropdown.click();
    await buildingDropdown.locator('.menu.transition').waitFor({ state: 'visible' });
    await buildingDropdown.locator('div[role="option"] .text', { hasText: buildingName }).click();

    await closeSystemPopup(page);
    console.log(`เลือกอาคาร: ${buildingName}`);
}

async function getRoomListContainer(page) {
    const roomListContainer = page.locator('div.defaultBorder').filter({
        has: page.locator('input[name="selectAllRoom"]')
    }).first();

    await page.waitForTimeout(3000);
    await loadAllRooms(roomListContainer);

    return roomListContainer;
}

async function exportRoomsByBatch(page, roomListContainer, buildingName, batchSize) {
    const roomItems = roomListContainer.locator('input[name="select"]');
    const totalRooms = await roomItems.count();

    console.log(`พบห้องทั้งหมด ${totalRooms} ห้อง`);

    const buildingPrefix = sanitizeFileName(buildingName);

    for (let startIndex = 0; startIndex < totalRooms; startIndex += batchSize) {
        const endIndex = Math.min(startIndex + batchSize, totalRooms);
        const selectedRooms = [];

        console.log(`Export ห้อง ${startIndex + 1} - ${endIndex}`);

        for (let i = startIndex; i < endIndex; i++) {
            const roomLabel = roomItems
                .nth(i)
                .locator('xpath=following-sibling::label');

            const roomName = (await roomLabel.textContent())?.trim() || '';
            selectedRooms.push(roomName);

            await roomLabel.scrollIntoViewIfNeeded();
            await roomLabel.click();
            await page.waitForTimeout(250);
            await closeSystemPopup(page);
        }

        console.log('Rooms:', selectedRooms.join(', '));

        const exportButton = page.locator('button.btnc', {
            hasText: 'ส่งออกเป็น Excel'
        }).first();

        await page.waitForTimeout(2000);
        await closeSystemPopup(page);

        const [download] = await Promise.all([
            page.waitForEvent('download', { timeout: 120000 }),
            exportButton.evaluate(el => el.click())
        ]);

        const fileName = `${buildingPrefix}_report_${startIndex + 1}_to_${endIndex}.xlsx`;
        const filePath = path.join(CONFIG.DOWNLOAD_DIR, fileName);

        await download.saveAs(filePath);
        console.log(`Download: ${fileName}`);

        await page.waitForTimeout(3000);
        await closeSystemPopup(page);

        for (let i = startIndex; i < endIndex; i++) {
            const roomLabel = roomItems
                .nth(i)
                .locator('xpath=following-sibling::label');

            await roomLabel.scrollIntoViewIfNeeded();
            await roomLabel.click();
            await page.waitForTimeout(250);
            await closeSystemPopup(page);
        }

        console.log(`จบรอบห้อง ${startIndex + 1} - ${endIndex}`);
        await page.waitForTimeout(1500);
    }
}

// =========================
// Main
// =========================
(async () => {
    ensureDownloadDir();

    const browser = await chromium.launch({ headless: false });
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    try {
        await login(page, CONFIG.EMAIL, CONFIG.PASSWORD);
        await openRoomUsageReport(page);

        await selectStartDate(page, CONFIG.START_DAY, CONFIG.START_MONTH, CONFIG.START_YEAR);
        await selectEndDate(page, CONFIG.END_DAY, CONFIG.END_MONTH, CONFIG.END_YEAR);

        await selectBuilding(page, CONFIG.BUILDING_NAME);

        const roomListContainer = await getRoomListContainer(page);

        await exportRoomsByBatch(
            page,
            roomListContainer,
            CONFIG.BUILDING_NAME,
            CONFIG.ROOM_BATCH_SIZE
        );

        console.log('Export ครบทั้งหมดเรียบร้อย');
        console.log('ยังไม่ได้ merge file อัตโนมัติ เพื่อให้คุณไปรันหลายอาคารก่อนค่อย merge');

    } catch (error) {
        console.error('เกิดข้อผิดพลาด:', error);
    } finally {
        await browser.close();
    }
})();