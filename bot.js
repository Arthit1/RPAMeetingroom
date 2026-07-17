const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');

// โครงสร้าง Config เริ่มต้น
let CONFIG = {
    START_DAY: 2,
    START_MONTH: 2,
    START_YEAR: 2026,
    END_DAY: 27,
    END_MONTH: 2,
    END_YEAR: 2026,
    NUM_DAYS: 19, // เก็บจำนวนวันสำหรับไปใช้ต่อในสคริปต์ .py (18-19วัน)
    BUILDING_NAME: 'Swan Lake',
    BUILDING_LIST: ['CP All Academy', 'Food Technology', 'PANYATARA1', 'PANYATARA2', 'Swan Lake', 'THE TARA', 'ZOOM'],
    ROOM_BATCH_SIZE: 10, // ล็อกค่าเป็น 10 ทุกครั้ง ทำงานเบื้องหลัง ไม่โชว์หน้าจอแล้ว
    EMAIL: '',
    PASSWORD: '',
    DOWNLOAD_DIR: './downloads'
};

const monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const monthsLong = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// ===================================================
// ฟังก์ชันเด้งหน้าต่างกรอกข้อมูล (GUI Form via Playwright)
// ===================================================
async function showInputForm(browser) {
    console.log('⏳ กำลังเปิดหน้าต่างกรอกข้อมูล...');
    const context = await browser.newContext();
    const page = await context.newPage();
    
    // ตั้งค่าขนาดหน้าต่างฟอร์มให้กะทัดรัด (ปรับความสูงลงมาเหลือ 530)
    await page.setViewportSize({ width: 450, height: 530 });

    const buildingOptionsHtml = CONFIG.BUILDING_LIST.map(building => {
        const isSelected = CONFIG.BUILDING_NAME === building ? 'selected' : '';
        return `<option value="${building}" ${isSelected}>${building}</option>`;
    }).join('\n');

    // จัดรูปแบบวันที่เริ่มต้นของช่องปฏิทินป๊อปอัปตามค่า CONFIG อัตโนมัติ
    const formatPad = (num) => String(num).padStart(2, '0');
    const defaultStartDate = `${CONFIG.START_YEAR}-${formatPad(CONFIG.START_MONTH)}-${formatPad(CONFIG.START_DAY)}`;
    const defaultEndDate = `${CONFIG.END_YEAR}-${formatPad(CONFIG.END_MONTH)}-${formatPad(CONFIG.END_DAY)}`;

    // สร้างหน้าตาฟอร์มด้วย HTML + CSS (เอาช่อง Batch Size ออกถาวรแล้ว)
    const formHtml = `
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>RPA Config Settings</title>
        <style>
            body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; padding: 20px; color: #333; }
            .card { background: white; padding: 25px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
            h2 { margin-top: 0; color: #2c3e50; border-bottom: 2px solid #eee; padding-bottom: 10px; font-size: 20px; }
            .form-group { margin-bottom: 15px; }
            label { display: block; margin-bottom: 5px; font-weight: 600; font-size: 13px; color: #4a5568; }
            input, select { width: 100%; padding: 8px 12px; border: 1px solid #cbd5e0; border-radius: 4px; box-sizing: border-box; font-size: 14px; background-color: white; }
            input:focus, select:focus { border-color: #3182ce; outline: none; }
            .row { display: flex; gap: 10px; }
            .row .form-group { flex: 1; }
            button { width: 100%; background-color: #3182ce; color: white; border: none; padding: 12px; border-radius: 4px; font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 15px; transition: background 0.2s; }
            button:hover { background-color: #2b6cb0; }
            button:disabled { background-color: #a0aec0; cursor: not-allowed; }
        </style>
    </head>
    <body>
        <div class="card">
            <h2>⚙️ RPA Meeting Room Settings</h2>
            <form id="configForm">
                <div class="form-group">
                    <label>Email Address</label>
                    <input type="email" id="email" value="${CONFIG.EMAIL}" required>
                </div>
                <div class="form-group">
                    <label>Password</label>
                    <input type="password" id="password" value="${CONFIG.PASSWORD}" required>
                </div>
                <div class="form-group">
                    <label>Building Name (เลือกอาคาร)</label>
                    <select id="building" required>
                        ${buildingOptionsHtml}
                    </select>
                </div>
                <div class="row">
                    <div class="form-group">
                        <label>Start Date</label>
                        <input type="date" id="startDate" value="${defaultStartDate}" required>
                    </div>
                    <div class="form-group">
                        <label>End Date</label>
                        <input type="date" id="endDate" value="${defaultEndDate}" required>
                    </div>
                </div>
                <div class="form-group">
                    <label>จำนวนวันที่ใช้</label>
                    <input type="number" id="numDays" value="${CONFIG.NUM_DAYS}" min="1" required>
                </div>
                <button type="submit" id="submitBtn">🚀 เริ่มรันบอททำงาน</button>
            </form>
        </div>

        <script>
            document.getElementById('configForm').addEventListener('submit', function(e) {
                e.preventDefault();
                
                const btn = document.getElementById('submitBtn');
                btn.innerText = '⏳ กำลังส่งข้อมูลไปยังบอท...';
                btn.disabled = true;

                // รวบรวมข้อมูลส่งกลับฝั่ง Node.js
                const data = {
                    email: document.getElementById('email').value,
                    password: document.getElementById('password').value,
                    building: document.getElementById('building').value,
                    startDate: document.getElementById('startDate').value,
                    endDate: document.getElementById('endDate').value,
                    numDays: document.getElementById('numDays').value
                };
                window.submitRpaConfig(data);
            });
        </script>
    </body>
    </html>
    `;

    return new Promise(async (resolve) => {
        // รับข้อมูลกลับมาประมวลผลบน Node.js
        await page.exposeFunction('submitRpaConfig', async (data) => {
            try {
                // บันทึกข้อมูลเข้าตัวแปร CONFIG หลัก
                CONFIG.EMAIL = data.email;
                CONFIG.PASSWORD = data.password;
                CONFIG.BUILDING_NAME = data.building;
                CONFIG.NUM_DAYS = parseInt(data.numDays) || 0;

                // แกะฟอร์แมตวันที่จากปฏิทิน HTML (YYYY-MM-DD)
                const [sYear, sMonth, sDay] = data.startDate.split('-');
                CONFIG.START_DAY = parseInt(sDay);
                CONFIG.START_MONTH = parseInt(sMonth);
                CONFIG.START_YEAR = parseInt(sYear);

                const [eYear, eMonth, eDay] = data.endDate.split('-');
                CONFIG.END_DAY = parseInt(eDay);
                CONFIG.END_MONTH = parseInt(eMonth);
                CONFIG.END_YEAR = parseInt(eYear);

                console.log('\n=========================================');
                console.log('✅ บันทึกค่า Config และอัปเดตข้อมูลสำเร็จ!');
                console.log(`📍 อาคารที่เลือก: ${CONFIG.BUILDING_NAME}`);
                console.log(`📅 ช่วงเวลา: ${CONFIG.START_DAY}/${CONFIG.START_MONTH}/${CONFIG.START_YEAR} ถึง ${CONFIG.END_DAY}/${CONFIG.END_MONTH}/${CONFIG.END_YEAR}`);
                console.log(`🔢 จำนวนวันที่บันทึก (สำหรับ .py): ${CONFIG.NUM_DAYS} วัน`);
                console.log(`⚙️ Room Batch Size เบื้องหลัง: ${CONFIG.ROOM_BATCH_SIZE} ห้อง/รอบ (คงที่)`);
                console.log('=========================================\n');

                await context.close(); // ปิดหน้าต่างฟอร์ม
                resolve(); // ปลดล็อค Flow หลักให้ทำงานต่อ
            } catch (err) {
                console.error("เกิดข้อผิดพลาดในการประมวลผลฟอร์ม:", err);
            }
        });

        // โหลดหน้าฟอร์มขึ้นมาแสดงผล
        await page.setContent(formHtml);
    });
}

// =========================
// Utility & Playwright Flow
// =========================
function ensureDownloadDir() {
    if (!fs.existsSync(CONFIG.DOWNLOAD_DIR)) {
        fs.mkdirSync(CONFIG.DOWNLOAD_DIR, { recursive: true });
    }
}

function sanitizeFileName(text) {
    return String(text).trim().replace(/\s+/g, '_').replace(/[\\/:*?"<>|]/g, '');
}

async function closeSystemPopup(page) {
    try {
        const popup = page.locator('.ui.small.modal:visible').first();
        if (await popup.count() > 0) {
            const okButton = popup.locator('button.okButton').first();
            if (await okButton.count() > 0) {
                await okButton.click({ force: true });
                await page.waitForTimeout(1000);
            }
        }
    } catch {}
}

async function loadAllRooms(page, roomListContainer) {
    const scrollBox = roomListContainer.locator('div[style*="overflow: hidden auto"]').first();
    let previousCount = 0;
    let sameCountRounds = 0;

    while (true) {
        const currentCount = await roomListContainer.locator('input[name="select"]').count();
        console.log(`กำลังโหลดรายการห้อง... ตอนนี้พบ ${currentCount} ห้อง`);

        if (currentCount === previousCount) { sameCountRounds += 1; } 
        else { sameCountRounds = 0; }

        if (sameCountRounds >= 3) break;
        previousCount = currentCount;

        await scrollBox.evaluate(el => { el.scrollTop = el.scrollHeight; });
        await page.waitForTimeout(1000);
    }
    await scrollBox.evaluate(el => { el.scrollTop = 0; });
    await page.waitForTimeout(1000);
}

async function login(page, email, password) {
    console.log('กำลังเปิดหน้า Login');
    await page.goto('https://meetingroomadmin.cpall.co.th/login', { waitUntil: 'domcontentloaded' });
    const loginForm = page.locator('div.login-wrapper.open div.ui.stacked.segment').first();
    await loginForm.locator('input[name="emailaddress"]').fill(email);
    await loginForm.locator('input[name="password"]').fill(password);
    await loginForm.locator('button:has-text("เข้าสู่ระบบ")').evaluate(el => el.click());
    await page.waitForTimeout(5000);
    await closeSystemPopup(page);
    console.log('Login สำเร็จ');
}

async function openRoomUsageReport(page) {
    const reportDropdown = page.locator('div[role="listbox"]').filter({ has: page.locator('div.text', { hasText: 'รายงาน' }) }).first();
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

        if (currentDate < targetDate) { await calendar.locator('.ant-calendar-next-month-btn').click(); } 
        else { await calendar.locator('.ant-calendar-prev-month-btn').click(); }
        await page.waitForTimeout(200);
    }
    await calendar.locator(`td[title="${targetTitle}"]`).click();
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

        if (currentDate < targetDate) { await calendar.locator('.ant-calendar-next-month-btn').click(); } 
        else { await calendar.locator('.ant-calendar-prev-month-btn').click(); }
        await page.waitForTimeout(200);
    }
    await calendar.locator(`td[title="${targetTitle}"]`).click();
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
    const roomListContainer = page.locator('div.defaultBorder').filter({ has: page.locator('input[name="selectAllRoom"]') }).first();
    await page.waitForTimeout(3000);
    await loadAllRooms(page, roomListContainer);
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

        console.log(`\n--- เริ่มรอบดาวน์โหลดห้องลำดับที่ ${startIndex + 1} ถึง ${endIndex} ---`);

        for (let i = startIndex; i < endIndex; i++) {
            const roomLabel = roomItems.nth(i).locator('xpath=following-sibling::label');
            const roomName = (await roomLabel.textContent())?.trim() || '';
            selectedRooms.push(roomName);

            await roomLabel.scrollIntoViewIfNeeded();
            await roomLabel.click();
            await page.waitForTimeout(200);
            await closeSystemPopup(page);
        }

        const exportButton = page.locator('button.btnc', { hasText: 'ส่งออกเป็น Excel' }).first();
        await page.waitForTimeout(1500);
        await closeSystemPopup(page);

        const [download] = await Promise.all([
            page.waitForEvent('download', { timeout: 120000 }),
            exportButton.evaluate(el => el.click())
        ]);

        // ดึงค่าวันที่จาก CONFIG มาทำเป็นฟอร์แมตสั้นๆ เช่น 02022026_to_27022026
        const pad = (num) => String(num).padStart(2, '0');
        const dateStamp = `${pad(CONFIG.START_DAY)}${pad(CONFIG.START_MONTH)}${CONFIG.START_YEAR}_to_${pad(CONFIG.END_DAY)}${pad(CONFIG.END_MONTH)}${CONFIG.END_YEAR}`;

        // เอา DateStamp ไปแปะท้ายชื่อไฟล์ย่อย
        const fileName = `${buildingPrefix}_report_${startIndex + 1}_to_${endIndex}_(${dateStamp}).xlsx`;
        const filePath = path.join(CONFIG.DOWNLOAD_DIR, fileName);
        await download.saveAs(filePath);
        console.log(`✅ ดาวน์โหลดสำเร็จ: ${fileName}`);
        await page.waitForTimeout(2000);

        console.log('กำลังเคลียร์การเลือกห้องชุดนี้ออก...');
        for (let i = startIndex; i < endIndex; i++) {
            const roomLabel = roomItems.nth(i).locator('xpath=following-sibling::label');
            await roomLabel.scrollIntoViewIfNeeded();
            await roomLabel.click();
            await page.waitForTimeout(200);
            await closeSystemPopup(page);
        }
        await page.waitForTimeout(1500);
    }
}

// =========================
// Main Execution
// =========================
(async () => {
    ensureDownloadDir();

    // เปิดเบราว์เซอร์หลักขึ้นมาแบบเห็นหน้าต่าง (headless: false)
    const browser = await chromium.launch({ headless: false });

    // สั่งเด้งหน้าต่างฟอร์ม HTML ให้กรอกข้อมูล
    await showInputForm(browser);

    // พอผู้ใช้กดปุ่ม 'เริ่มรันบอททำงาน' โค้ดด้านล่างนี้จะทำงานต่อทันทีด้วยค่าที่กรอกมา
    console.log(`🚀 เริ่มระบบอัตโนมัติ: อาคาร ${CONFIG.BUILDING_NAME}`);
    
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    try {
        await login(page, CONFIG.EMAIL, CONFIG.PASSWORD);
        await openRoomUsageReport(page);
        await selectStartDate(page, CONFIG.START_DAY, CONFIG.START_MONTH, CONFIG.START_YEAR);
        await selectEndDate(page, CONFIG.END_DAY, CONFIG.END_MONTH, CONFIG.END_YEAR);
        await selectBuilding(page, CONFIG.BUILDING_NAME);

        const roomListContainer = await getRoomListContainer(page);
        await exportRoomsByBatch(page, roomListContainer, CONFIG.BUILDING_NAME, CONFIG.ROOM_BATCH_SIZE);

        console.log('🎉 บอทดาวน์โหลดรายงานเสร็จสิ้นครบถ้วนแล้ว!');
    } catch (error) {
        console.error('เกิดข้อผิดพลาดในการรันบอท:', error);
    } finally {
        await browser.close();
    }
})();