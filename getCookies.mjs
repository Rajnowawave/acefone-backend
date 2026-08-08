// getCookies.mjs
import puppeteer from 'puppeteer';
import { writeFileSync } from 'fs';

const EMAIL    = process.env.ACEFONE_EMAIL    || "customercare@adinath.net.in";
const PASSWORD = process.env.ACEFONE_PASSWORD || "Office@2005"; // ✅ Yahan apna password daalo

async function getSession() {
  const browser = await puppeteer.launch({ 
    headless: true,  // background mein chalega
    args: ["--no-sandbox", "--disable-setuid-sandbox"]
  });

  const page = await browser.newPage();
  await page.setUserAgent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");

  try {
    await page.goto('https://console.acefone.in/login', { waitUntil: 'networkidle2', timeout: 30000 });
    await page.waitForSelector('#loginId', { timeout: 15000 });

    await page.$eval('#loginId', el => el.value = '');
    await page.$eval('#password', el => el.value = '');
    await page.type('#loginId', EMAIL, { delay: 60 });
    await page.type('#password', PASSWORD, { delay: 60 });
    await page.click('#login_button');

    await page.waitForFunction(
      () => !window.location.href.includes('/login'),
      { timeout: 25000 }
    );

    // CSRF token lo page se
    const csrfToken = await page.evaluate(() => {
      return document.querySelector('meta[name="csrf-token"]')?.content
        || document.querySelector('input[name="_token"]')?.value
        || '';
    });

    // Cookies lo
    const cookies = await page.cookies();
    const cookieString = cookies.map(c => `${c.name}=${c.value}`).join('; ');

    const session = {
      cookieString,
      csrfToken,
      savedAt: new Date().toISOString()
    };

    writeFileSync('./session.json', JSON.stringify(session, null, 2));
    console.log("✅ Session saved at:", session.savedAt);
    console.log("CSRF Token:", csrfToken.substring(0, 15) + "...");

    await browser.close();
    return session;

  } catch (err) {
    console.error("❌ Login failed:", err.message);
    await browser.close();
    throw err;
  }
}

getSession();