const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
const fs = require('fs');

puppeteer.use(StealthPlugin());

// =========================================================
// 1. RETRY HELPER (Google API 500/Timeout)
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
// 2. VIEW.COM.AU EXTRACTOR
// Pulls from props.pageProps.data in __NEXT_DATA__
// =========================================================
function extractViewData(jsonData) {
    const empty = {
        addr: "", propertyType: "", desc: "", county: "",
        price: "", bedrooms: "", bathrooms: "", carparks: "",
        landSize: "", agents: "", agency: "", agencyAddress: ""
    };

    const d = jsonData?.props?.pageProps?.data;
    if (!d) return empty;

    const street = [d.streetNumber, d.streetName].filter(Boolean).join(" ").trim();
    const addr = [street, d.suburbName, d.state, d.postcode].filter(Boolean).join(", ").trim();

    const agents = (d.agents || [])
        .map(a => {
            const name = `${a.firstName || ""} ${a.lastName || ""}`.trim();
            const parts = [name];
            if (a.mobile) parts.push(a.mobile);
            if (a.email) parts.push(a.email);
            return parts.join(" | ");
        })
        .filter(Boolean)
        .join(" ; ");

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
        agents,
        agency: d.agency?.name || "",
        agencyAddress: d.agency?.address || ""
    };
}

// =========================================================
// 3. CORE SCRAPER
// =========================================================
async function runScraper() {
    console.log("🚀 Starting View.com.au Scraper V1...");

    const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
    const serviceAccountAuth = new JWT({
        email: creds.client_email,
        key: creds.private_key,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });

    const doc = new GoogleSpreadsheet(process.env.SPREADSHEET_ID, serviceAccountAuth);
    await doc.loadInfo();
    const sheet = doc.sheetsByIndex[0];

    // Columns A..T (20) — A=URL, B..I spare, J..S output, T=status
    if (sheet.columnCount < 20) {
        console.log(`📏 Expanding columns from ${sheet.columnCount} to 20...`);
        await sheet.resize({ rowCount: sheet.rowCount, columnCount: 20 });
    }

    await sheet.loadCells();

    const STATUS_COL = 19; // T (0-based)
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
        const url = sheet.getCell(rowIndex, 0).value;
        const status = sheet.getCell(rowIndex, STATUS_COL).value || "";

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

            // Slower, more human-like pacing for DataDome
            await delay(Math.floor(Math.random() * 3000) + 3000);

            await page.goto(url, { waitUntil: 'networkidle2', timeout: 45000 });

            // Explicitly wait for the Next.js payload
            await page.waitForSelector('script#__NEXT_DATA__', { timeout: 10000 }).catch(() => {});

            // DataDome + generic block detection
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
                sheet.getCell(rowIndex, STATUS_COL).value = "❌ BLOCKED (DataDome)";
                stagedCellsToSave.push(rowIndex);
                consecutiveBlocks++;

                if (consecutiveBlocks >= 3) {
                    console.log("🛑 3 consecutive blocks — stopping run to avoid IP burn.");
                    await saveWithRetry(sheet);
                    rowsRemaining = true;
                    await page.close();
                    break;
                }
                await page.close();
                continue;
            }

            consecutiveBlocks = 0;

            // Extract from __NEXT_DATA__
            const extractedData = await page.evaluate(() => {
                const nextDataScript = document.querySelector('script#__NEXT_DATA__');
                if (!nextDataScript) return null;
                try {
                    const jsonData = JSON.parse(nextDataScript.innerText);
                    const d = jsonData?.props?.pageProps?.data;
                    if (!d) return null;

                    const street = [d.streetNumber, d.streetName].filter(Boolean).join(" ").trim();
                    const addr = [street, d.suburbName, d.state, d.postcode].filter(Boolean).join(", ").trim();

                    const agents = (d.agents || [])
                        .map(a => {
                            const name = `${a.firstName || ""} ${a.lastName || ""}`.trim();
                            const parts = [name];
                            if (a.mobile) parts.push(a.mobile);
                            if (a.email) parts.push(a.email);
                            return parts.join(" | ");
                        })
                        .filter(Boolean)
                        .join(" ; ");

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
                        agents,
                        agency: d.agency?.name || "",
                        agencyAddress: d.agency?.address || ""
                    };
                } catch (e) {
                    return null;
                }
            });

            if (!extractedData) {
                console.log(`⚠️ No __NEXT_DATA__ payload on Row ${actualRowNumber}`);
                sheet.getCell(rowIndex, STATUS_COL).value = "⚠️ No Data";
                stagedCellsToSave.push(rowIndex);
                await page.close();
                continue;
            }

            // Write to sheet
            sheet.getCell(rowIndex, 9).value  = extractedData.addr;          // J
            sheet.getCell(rowIndex, 10).value = extractedData.propertyType;  // K
            sheet.getCell(rowIndex, 11).value = extractedData.desc;          // L
            sheet.getCell(rowIndex, 12).value = extractedData.county;        // M
            sheet.getCell(rowIndex, 13).value = extractedData.price;         // N
            sheet.getCell(rowIndex, 14).value = extractedData.bedrooms;      // O
            sheet.getCell(rowIndex, 15).value = extractedData.bathrooms;     // P
            sheet.getCell(rowIndex, 16).value = extractedData.agents;        // Q
            sheet.getCell(rowIndex, 17).value = extractedData.agency;        // R
            sheet.getCell(rowIndex, 18).value = extractedData.agencyAddress; // S
            sheet.getCell(rowIndex, STATUS_COL).value = "✅ SUCCESS";        // T

            stagedCellsToSave.push(rowIndex);
            console.log(`✔️ Row ${actualRowNumber} | 🏠 ${extractedData.propertyType || '?'} | 💰 ${extractedData.price || '?'} | 👤 ${extractedData.agents || '?'}`);
            scrapeCount++;

        } catch (e) {
            console.error(`🛑 Row ${actualRowNumber} error: ${e.message}`);
            sheet.getCell(rowIndex, STATUS_COL).value = "🛑 Error: " + e.message.slice(0, 80);
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
