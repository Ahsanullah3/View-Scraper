const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const fs = require('fs');

puppeteer.use(StealthPlugin());

// =========================================================
// 1. HELPERS
// =========================================================
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

async function saveWithRetry(sheet, retries = 3) {
    for (let i = 0; i < retries; i++) {
        try {
            await sheet.saveUpdatedCells();
            return;
        } catch (error) {
            if (i === retries - 1) {
                console.error("❌ Max retries reached. Google API unavailable.");
                throw error;
            }
            const waitTime = (2 ** i) * 1000;
            console.log(`⚠️ Google API error. Retrying in ${2 ** i}s...`);
            await delay(waitTime);
        }
    }
}

// Column map (0-based) — A..AG
const COL = {
    url: 0, addr: 9, propertyType: 10, desc: 11, county: 12,
    price: 13, bedrooms: 14, bathrooms: 15, carparks: 16, landSize: 17,
    agencyName: 18, agencyAddress: 19, agencyPhone: 20, agencyProfile: 21,
    agent1Name: 22, agent1Phone: 23, agent1Mobile: 24, agent1Email: 25, agent1Profile: 26,
    agent2Name: 27, agent2Phone: 28, agent2Mobile: 29, agent2Email: 30, agent2Profile: 31,
    status: 32
};
const REQUIRED_COLUMNS = 33;

// =========================================================
// 2. IN-PAGE EXTRACTOR
// =========================================================
function extractViewDataInPage() {
    const nextDataScript = document.querySelector('script#__NEXT_DATA__');
    if (!nextDataScript) return null;

    let jsonData;
    try {
        jsonData = JSON.parse(nextDataScript.innerText);
    } catch (e) {
        return null;
    }

    const d = jsonData?.props?.pageProps?.data;
    if (!d) return null;

    const BASE = "https://view.com.au";

    const street = [d.streetNumber, d.streetName].filter(Boolean).join(" ").trim();
    const addr = [street, d.suburbName, d.state, d.postcode].filter(Boolean).join(", ").trim();

    const agentList = Array.isArray(d.agents) ? d.agents : [];
    const buildAgent = (a) => ({
        fullName: `${(a.firstName || "").trim()} ${(a.lastName || "").trim()}`.trim(),
        phone: a.phone || "",
        mobile: a.mobile || "",
        email: a.email || "",
        profileUrl: a.agentProfileLink ? `${BASE}${a.agentProfileLink}` : ""
    });
    const ag1 = buildAgent(agentList[0] || {});
    const ag2 = buildAgent(agentList[1] || {});

    return {
        addr,
        propertyType: d.primaryPropertyType || d.sourcePropertyType || "",
        desc: d.description || "",
        county: d.lgaName || "",
        price: d.priceText || "",
        bedrooms: d.bedrooms ?? "",
        bathrooms: d.bathrooms ?? "",
        carparks: d.carparks ?? "",
        landSize: d.landSize ?? "",

        agencyName: d.agency?.name || "",
        agencyAddress: d.agency?.address || "",
        agencyPhone: d.agency?.phone || agentList[0]?.phone || "",
        agencyProfileUrl: d.agency?.agencyProfileLink
            ? `${BASE}${d.agency.agencyProfileLink}`
            : "",

        agent1Name: ag1.fullName,
        agent1Phone: ag1.phone,
        agent1Mobile: ag1.mobile,
        agent1Email: ag1.email,
        agent1Profile: ag1.profileUrl,

        agent2Name: ag2.fullName,
        agent2Phone: ag2.phone,
        agent2Mobile: ag2.mobile,
        agent2Email: ag2.email,
        agent2Profile: ag2.profileUrl
    };
}

// =========================================================
// 3. SHEET SAFETY HELPERS
// =========================================================
async function ensureSheetWidth(sheet, required) {
    if (sheet.columnCount < required) {
        console.log(`📏 Expanding columns from ${sheet.columnCount} to ${required}...`);
        await sheet.resize({ rowCount: sheet.rowCount, columnCount: required });
        sheet.resetLocalCache(true);
        await sheet.loadCells();
    }
}

function safeGetCell(sheet, row, col) {
    try {
        return sheet.getCell(row, col);
    } catch (e) {
        if (String(e.message).includes("not been loaded") || String(e.message).includes("Out of bounds")) {
            throw new Error(`Cell (${row},${col}) not loaded — sheet resize/loadCells mismatch`);
        }
        throw e;
    }
}

function writeCell(sheet, row, col, value) {
    const cell = safeGetCell(sheet, row, col);
    cell.value = value;
}

// =========================================================
// 4. CORE SCRAPER
// =========================================================
async function runScraper() {
    console.log("🚀 Starting View.com.au Scraper V6 (bullet-proof)...");

    const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    const serviceAccountAuth = new JWT({
        email: creds.client_email,
        key: creds.private_key,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];

    await ensureSheetWidth(sheet, REQUIRED_COLUMNS);
    await sheet.loadCells();

    const browser = await puppeteer.launch({
        headless: "new",
        executablePath: process.env.CHROME_PATH || undefined,
        args: [
            '--no-sandbox',
            '--disable-setuid-sandbox',
            '--disable-blink-features=AutomationControlled',
            '--window-size=1366,768'
        ]
    });

    let scrapeCount = 0;
    let rowsRemaining = false;
    let consecutiveBlocks = 0;

    const FLUSH_BATCH_SIZE = 20;
    const MAX_ROWS_PER_RUN = 60;
    const MAX_BLOCK_ATTEMPTS = 3;

    let stagedCellsToSave = [];

    for (let rowIndex = 1; rowIndex < sheet.rowCount; rowIndex++) {
        // ---- Read URL + status safely ----
        let url, statusVal;
        try {
            url = safeGetCell(sheet, rowIndex, COL.url).value;
            statusVal = safeGetCell(sheet, rowIndex, COL.status).value || "";
        } catch (e) {
            continue;
        }

        if (!url) continue;
        const urlStr = String(url).trim();
        if (!urlStr) continue;
        if (!urlStr.includes("view.com.au")) continue;

        const statusStr = String(statusVal);

        // Skip already-successful rows
        if (statusStr.includes("✅")) continue;
        // Skip rows that gave up after max attempts
        if (statusStr.includes("gave up")) continue;

        // Parse attempt count from status like "⏳ BLOCKED (retry next batch — attempt 2)"
        let attemptCount = 0;
        const attemptMatch = statusStr.match(/attempt\s+(\d+)/i);
        if (attemptMatch) attemptCount = parseInt(attemptMatch[1], 10);
        if (attemptCount >= MAX_BLOCK_ATTEMPTS) continue;

        // Row budget — only successful scrapes count
        if (scrapeCount >= MAX_ROWS_PER_RUN) {
            console.log(`🛑 Reached ${MAX_ROWS_PER_RUN} rows. Stopping to rotate environment...`);
            rowsRemaining = true;
            break;
        }

        const actualRowNumber = rowIndex + 1;
        console.log(`🕵️ Row ${actualRowNumber}: ${urlStr}`);

        const page = await browser.newPage();

        try {
            await page.setRequestInterception(true);
            page.on('request', (req) => {
                if (['image', 'stylesheet', 'font', 'media'].includes(req.resourceType())) {
                    req.abort();
                } else {
                    req.continue();
                }
            });

            // Fast human-ish pacing
            await delay(Math.floor(Math.random() * 1500) + 1500);

            await page.goto(urlStr, { waitUntil: 'domcontentloaded', timeout: 30000 });

            // ---- Fast __NEXT_DATA__ presence poll (max ~1s) ----
            let hasNextData = false;
            for (let i = 0; i < 10; i++) {
                try {
                    hasNextData = await page.evaluate(
                        () => !!document.querySelector('script#__NEXT_DATA__')
                    );
                } catch (_) {
                    hasNextData = false;
                    break;
                }
                if (hasNextData) break;
                await delay(100);
            }

            if (!hasNextData) {
                console.log(`⚠️ No __NEXT_DATA__ on Row ${actualRowNumber} — skipping`);
                writeCell(sheet, rowIndex, COL.status, "⚠️ No Data");
                stagedCellsToSave.push(rowIndex);
                continue;               // finally handles close
            }

            // ---- Combined block detection (1 evaluate) ----
            let blockCheck = { title: "", body: "", captcha: false };
            try {
                blockCheck = await page.evaluate(() => {
                    const title = document.title || "";
                    const body = (document.body?.innerText || "").slice(0, 2000);
                    const captcha = !!document.querySelector(
                        'iframe[src*="captcha-delivery"], iframe[src*="datadome"]'
                    );
                    return { title, body, captcha };
                });
            } catch (_) { /* page may have died; treat as not blocked */ }

            const isBlocked =
                blockCheck.title.includes("Pardon Our Interruption") ||
                blockCheck.title.includes("Robot Check") ||
                blockCheck.title.includes("Access Denied") ||
                blockCheck.title.includes("Blocked") ||
                blockCheck.body.includes("datadome") ||
                blockCheck.body.includes("captcha-delivery") ||
                blockCheck.captcha;

            if (isBlocked) {
                attemptCount++;
                const statusText = attemptCount >= MAX_BLOCK_ATTEMPTS
                    ? `❌ BLOCKED (gave up after ${MAX_BLOCK_ATTEMPTS} attempts)`
                    : `⏳ BLOCKED (retry next batch — attempt ${attemptCount})`;

                console.log(`❌ BLOCKED on Row ${actualRowNumber} — ${statusText}`);
                writeCell(sheet, rowIndex, COL.status, statusText);
                stagedCellsToSave.push(rowIndex);
                consecutiveBlocks++;

                // Blocked rows do NOT consume the budget
                if (consecutiveBlocks >= 3) {
                    console.log("🛑 3 consecutive blocks — stopping run to protect IP.");
                    await saveWithRetry(sheet);
                    rowsRemaining = true;
                    break;                  // finally handles close
                }
                continue;                   // finally handles close
            }

            consecutiveBlocks = 0;

            // ---- Extract ----
            let extractedData = null;
            try {
                extractedData = await page.evaluate(extractViewDataInPage);
            } catch (_) { extractedData = null; }

            if (!extractedData) {
                console.log(`⚠️ No payload data on Row ${actualRowNumber} — skipping`);
                writeCell(sheet, rowIndex, COL.status, "⚠️ No Data");
                stagedCellsToSave.push(rowIndex);
                continue;                   // finally handles close
            }

            // ---- Write all fields ----
            writeCell(sheet, rowIndex, COL.addr, extractedData.addr);
            writeCell(sheet, rowIndex, COL.propertyType, extractedData.propertyType);
            writeCell(sheet, rowIndex, COL.desc, extractedData.desc);
            writeCell(sheet, rowIndex, COL.county, extractedData.county);
            writeCell(sheet, rowIndex, COL.price, extractedData.price);
            writeCell(sheet, rowIndex, COL.bedrooms, extractedData.bedrooms);
            writeCell(sheet, rowIndex, COL.bathrooms, extractedData.bathrooms);
            writeCell(sheet, rowIndex, COL.carparks, extractedData.carparks);
            writeCell(sheet, rowIndex, COL.landSize, extractedData.landSize);

            writeCell(sheet, rowIndex, COL.agencyName, extractedData.agencyName);
            writeCell(sheet, rowIndex, COL.agencyAddress, extractedData.agencyAddress);
            writeCell(sheet, rowIndex, COL.agencyPhone, extractedData.agencyPhone);
            writeCell(sheet, rowIndex, COL.agencyProfile, extractedData.agencyProfileUrl);

            writeCell(sheet, rowIndex, COL.agent1Name, extractedData.agent1Name);
            writeCell(sheet, rowIndex, COL.agent1Phone, extractedData.agent1Phone);
            writeCell(sheet, rowIndex, COL.agent1Mobile, extractedData.agent1Mobile);
            writeCell(sheet, rowIndex, COL.agent1Email, extractedData.agent1Email);
            writeCell(sheet, rowIndex, COL.agent1Profile, extractedData.agent1Profile);

            writeCell(sheet, rowIndex, COL.agent2Name, extractedData.agent2Name);
            writeCell(sheet, rowIndex, COL.agent2Phone, extractedData.agent2Phone);
            writeCell(sheet, rowIndex, COL.agent2Mobile, extractedData.agent2Mobile);
            writeCell(sheet, rowIndex, COL.agent2Email, extractedData.agent2Email);
            writeCell(sheet, rowIndex, COL.agent2Profile, extractedData.agent2Profile);

            writeCell(sheet, rowIndex, COL.status, "✅ SUCCESS");

            stagedCellsToSave.push(rowIndex);
            console.log(
                `✔️ Row ${actualRowNumber} | 🏠 ${extractedData.propertyType || '?'}` +
                ` | 💰 ${extractedData.price || '?'}` +
                ` | 👤1 ${extractedData.agent1Name || '-'}` +
                ` | 👤2 ${extractedData.agent2Name || '-'}` +
                ` | 🏢 ${extractedData.agencyName || '-'}`
            );
            scrapeCount++;

        } catch (e) {
            console.error(`🛑 Row ${actualRowNumber} error: ${e.message}`);
            try {
                writeCell(sheet, rowIndex, COL.status, "🛑 Error: " + e.message.slice(0, 80));
                stagedCellsToSave.push(rowIndex);
            } catch (inner) {
                console.error(`  ↳ Could not write error status: ${inner.message}`);
            }
        } finally {
            // Safe close — never throws if the page already died
            try { await page.close(); } catch (_) { /* already closed */ }
        }

        if (stagedCellsToSave.length >= FLUSH_BATCH_SIZE) {
            console.log(`📦 Flushing ${stagedCellsToSave.length} records...`);
            await saveWithRetry(sheet);
            stagedCellsToSave = [];
        }
    }

    if (stagedCellsToSave.length > 0) {
        console.log(`📦 Flushing final ${stagedCellsToSave.length} records...`);
        await saveWithRetry(sheet);
    }

    await browser.close();

    if (process.env.GITHUB_OUTPUT) {
        if (rowsRemaining) {
            fs.appendFileSync(process.env.GITHUB_OUTPUT, "has_more=true\n");
            console.log("🔄 Remaining rows. Relaying trigger...");
        } else {
            fs.appendFileSync(process.env.GITHUB_OUTPUT, "has_more=false\n");
            console.log("🎉 Sheet complete!");
        }
    }
}

runScraper();
