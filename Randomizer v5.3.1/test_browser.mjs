import puppeteer from 'puppeteer';

(async () => {
  const browser = await puppeteer.launch({
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });
  const page = await browser.newPage();
  
  page.on('console', msg => console.log(`[CONSOLE] ${msg.type().toUpperCase()}: ${msg.text()}`));
  page.on('pageerror', err => console.error(`[PAGE_ERROR]`, err));
  page.on('requestfailed', req => console.error(`[REQ_FAILED] ${req.url()} - ${req.failure()?.errorText}`));

  try {
    await page.goto('http://localhost:3000', { waitUntil: 'networkidle2' });
    console.log("[BODY LENGTH]", await page.evaluate(() => document.body.innerHTML.length));
    console.log("[ROOT CONTENT]", await page.evaluate(() => document.getElementById('root')?.innerHTML.substring(0, 200)));
  } catch (err) {
    console.error(`[GOTO ERROR]`, err);
  }
  await browser.close();
})();
