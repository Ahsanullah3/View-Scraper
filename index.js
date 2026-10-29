const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const fs = require('fs');

puppeteer.use(StealthPlugin());

// =========================================================
// 1. RETRY HELPER
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

// =========================================================
// 2. EXTRACTOR (in-page)
// Returns a flat object ready for column mapping
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
    const a1 = agentList[0] || {};
    const a2 = agentList[1] || {};

    const buildAgent = (a) => ({
        firstName: (a.firstName || "").trim(),
        lastName: (a.lastName || "").trim(),
        fullName: `${(a.firstName || "").trim()} ${(a.lastName || "").trim()}`.trim(),
        phone: a.phone || "",
        mobile: a.mobile || "",
        email: a.email || "",
        profileUrl: a.agentProfileLink ? `${BASE}${a.agentProfileLink}` : ""
    });

    const ag1 = buildAgent(a1);
    const ag2 = buildAgent(a2);

    // Agency phone: prefer explicit agency phone if present, else first agent's phone
    const agencyPhone = d.agency?.phone || a1.phone || "";

    return {
        // Property
        addr,
        propertyType: d.primaryPropertyType || d.sourcePropertyType || "",
        desc: d.description || "",
        county: d.lgaName || "",
        price: d.priceText || "",
        bedrooms: d.bedrooms ?? "",
        bathrooms: d.bathrooms ?? "",
        carparks: d.carparks ?? "",
        landSize: d.landSize ?? "",

        // Agency
        agencyName: d.agency?.name || "",
        agencyAddress: d.agency?.address || "",
        agencyPhone: agencyPhone,
        agencyProfileUrl: d.agency?.agencyProfileLink
            ? `${BASE}${d.agency.agencyProfileLink}`
            : "",

        // Agent 1
        agent1Name: ag1.fullName,
        agent1Phone: ag1.phone,
        agent1Mobile: ag1.mobile,
        agent1Email: ag1.email,
        agent1Profile: ag1.profileUrl,

        // Agent 2
        agent2Name: ag2.fullName,
        agent2Phone: ag2.phone,
        agent2Mobile: ag2.mobile,
        agent2Email: ag2.email,
        agent2Profile: ag2.profileUrl
    };
}

// =========================================================
// 3. CORE SCRAPER
// =========================================================
async function runScraper() {
    console.log("🚀 Starting View.com.au Scraper V2...");

    const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    const serviceAccountAuth = new JWT({
        email: creds.client_email,
        key: creds.private_key,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];

    // Need columns A..Y (25). A=URL, then outputs, last = status
    if (sheet.columnCount < 25) {
        console.log(`📏 Expanding columns from ${sheet.columnCount} to 25...`);
        await sheet.resize({ rowCount: sheet.rowCount, columnCount: 25 });
    }

    await sheet.loadCells();

    // Column map (0-based)
    const COL = {
        url: 0,          // A
        // J..(spare) used below
        addr: 9,         // J
        propertyType: 10,// K
        desc: 11,        // L
        county: 12,      // M
        price: 13,       // N
        bedrooms: 14,    // O
        bathrooms: 15,   // P
        carparks: 16,    // Q
        landSize: 17,    // R
        agencyName: 18,  // S
        agencyAddress: 19, // T
        agencyPhone: 20, // U
        agencyProfile: 21, // V
        agent1Name: 22,  // W
        agent1Phone: 23, // X
        agent1Mobile: 24,// Y
        agent1Email: 25, // Z
        agent1Profile: 26,// AA
        agent2Name: 27,  // AB
        agent2Phone: 28, // AC
        agent2Mobile: 29,// AD
        agent2Email: 30, // AE
        agent2Profile: 31,// AF
        status: 32       // AG
    };

    // Ensure enough columns for AG (32)
    if (sheet.columnCount < 33) {
        console.log(`📏 Expanding columns to 33 (A..AG)...`);
        await sheet.resize({ rowCount: sheet.rowCount, columnCount: 33 });
    }

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

    const FLUSH_BATCH_SIZE = 10;
    let stagedCellsToSave = [];

    for (let rowIndex = 1; rowIndex < sheet.rowCount; rowIndex++) {
        const url = sheet.getCell(rowIndex, COL.url).value;
        const status = sheet.getCell(rowIndex, COL.status).value || "";

        if (!url) continue;
        if (!url.includes("view.com.au")) continue;
        if (status.includes("✅")) continue;

        if (scrapeCount >= 30) {
            console.log("🛑 Reached 30 rows. Stopping to rotate environment...");
            rowsRemaining = true;
            break;
        }

        const actualRowNumber = rowIndex + 1;
        console.log(`🕵️ Row ${actualRowNumber}: ${url}`);

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

            await delay(Math.floor(Math.random() * 3000) + 3000);

            await page.goto(url, { waitUntil: 'networkidle2', timeout: 45000 });
            await page.waitForSelector('script#__NEXT_DATA__', { timeout: 10000 }).catch(() => {});

            const pageTitle = await page.title();
            const bodySnippet = await page.evaluate(() => document.body.innerText.slice(0, 2000));
            const hasCaptcha = await page.evaluate(
                () => !!document.querySelector('iframe[src*="captcha-delivery"], iframe[src*="datadome"]')
            );

            const isBlocked =
                pageTitle.includes("Pardon Our Interruption") ||
                pageTitle.includes("Robot Check") ||
                pageTitle.includes("Access Denied") ||
                pageTitle.includes("Blocked") ||
                bodySnippet.includes("datadome") ||
                bodySnippet.includes("captcha-delivery") ||
                hasCaptcha;

            if (isBlocked) {
                console.log(`❌ BLOCKED on Row ${actualRowNumber}`);
                sheet.getCell(rowIndex, COL.status).value = "❌ BLOCKED (DataDome)";
                stagedCellsToSave.push(rowIndex);
                consecutiveBlocks++;

                if (consecutiveBlocks >= 3) {
                    console.log("🛑 3 consecutive blocks — stopping run.");
                    await saveWithRetry(sheet);
                    rowsRemaining = true;
                    await page.close();
                    break;
                }
                await page.close();
                continue;
            }

            consecutiveBlocks = 0;

            const extractedData = await page.evaluate(extractViewDataInPage);

            if (!extractedData) {
                console.log(`⚠️ No __NEXT_DATA__ payload on Row ${actualRowNumber}`);
                sheet.getCell(rowIndex, COL.status).value = "⚠️ No Data";
                stagedCellsToSave.push(rowIndex);
                await page.close();
                continue;
            }

            // Write all fields
            const setCell = (idx, val) => { sheet.getCell(rowIndex, idx).value = val; };

            setCell(COL.addr, extractedData.addr);
            setCell(COL.propertyType, extractedData.propertyType);
            setCell(COL.desc, extractedData.desc);
            setCell(COL.county, extractedData.county);
            setCell(COL.price, extractedData.price);
            setCell(COL.bedrooms, extractedData.bedrooms);
            setCell(COL.bathrooms, extractedData.bathrooms);
            setCell(COL.carparks, extractedData.carparks);
            setCell(COL.landSize, extractedData.landSize);

            setCell(COL.agencyName, extractedData.agencyName);
            setCell(COL.agencyAddress, extractedData.agencyAddress);
            setCell(COL.agencyPhone, extractedData.agencyPhone);
            setCell(COL.agencyProfile, extractedData.agencyProfileUrl);

            setCell(COL.agent1Name, extractedData.agent1Name);
            setCell(COL.agent1Phone, extractedData.agent1Phone);
            setCell(COL.agent1Mobile, extractedData.agent1Mobile);
            setCell(COL.agent1Email, extractedData.agent1Email);
            setCell(COL.agent1Profile, extractedData.agent1Profile);

            setCell(COL.agent2Name, extractedData.agent2Name);
            setCell(COL.agent2Phone, extractedData.agent2Phone);
            setCell(COL.agent2Mobile, extractedData.agent2Mobile);
            setCell(COL.agent2Email, extractedData.agent2Email);
            setCell(COL.agent2Profile, extractedData.agent2Profile);

            setCell(COL.status, "✅ SUCCESS");

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
            sheet.getCell(rowIndex, COL.status).value = "🛑 Error: " + e.message.slice(0, 80);
            stagedCellsToSave.push(rowIndex);
        } finally {
            await page.close();
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
