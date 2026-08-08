// debugLogin.mjs
import puppeteer from 'puppeteer';

const browser = await puppeteer.launch({ 
  headless: false,
  slowMo: 500
});

const page = await browser.newPage();
await page.goto('https://console.acefone.in/login', { waitUntil: 'networkidle2' });

await page.waitForSelector('#loginId', { timeout: 15000 });

// ✅ Actual credentials
await page.type('#loginId', 'customercare@adinath.net.in', { delay: 100 });
await page.type('#password', 'Office@2005', { delay: 100 });

await page.click('#login_button');

await new Promise(r => setTimeout(r, 10000));

console.log('Final URL:', page.url());

await browser.close();