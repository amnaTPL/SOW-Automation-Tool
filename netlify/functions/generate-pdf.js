const chromium = require('@sparticuz/chromium-min');
const puppeteer = require('puppeteer-core');

// Chromium remote URL - Netlify will download it at runtime (not bundled)
const CHROMIUM_REMOTE_URL = 'https://github.com/Sparticuz/chromium/releases/download/v123.0.1/chromium-v123.0.1-pack.tar';

exports.handler = async (event) => {
  // Handle CORS preflight
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers: corsHeaders, body: '' };
  }

  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, headers: corsHeaders, body: 'Method Not Allowed' };
  }

  let browser = null;

  try {
    const { sowHTML, styles, filename } = JSON.parse(event.body || '{}');

    if (!sowHTML) {
      return { statusCode: 400, headers: corsHeaders, body: 'Missing sowHTML' };
    }

    const fullHTML = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <style>
    @page { size: A4; margin: 0; }
    * { box-sizing: border-box; }
    body {
      margin: 0; padding: 0; background: white;
      font-family: Arial, sans-serif;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    .sow-cover { margin: 0 !important; padding: 0 !important; }
    .sow-cover-top-img { width: 100%; display: block; }
    .sow-cover-content { padding: 20px 60px 24px !important; display: flex; flex-direction: column; align-items: center; background: white; }
    .sow-body { padding: 14mm 14mm 10mm 14mm !important; }
    .sow-doc { max-width: 100% !important; box-shadow: none !important; border: none !important; }
    .sow-tc-page-break { page-break-before: always; }
    .sow-billing-start  { page-break-before: always; }
    tr, td { page-break-inside: avoid; }
    .sow-responsibility-block { page-break-inside: avoid; }
    .sow-pricing-block { page-break-inside: avoid; }
    .preview-controls { display: none !important; }
    ${styles || ''}
  </style>
  <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap" rel="stylesheet">
</head>
<body>
  <div class="sow-doc">${sowHTML}</div>
</body>
</html>`;

    browser = await puppeteer.launch({
      args: chromium.args,
      defaultViewport: { width: 1240, height: 1754 },
      executablePath: await chromium.executablePath(CHROMIUM_REMOTE_URL),
      headless: 'new',
    });

    const page = await browser.newPage();
    await page.setContent(fullHTML, { waitUntil: 'networkidle0', timeout: 30000 });
    await new Promise(r => setTimeout(r, 1000)); // wait for fonts

    const pdf = await page.pdf({
      format: 'A4',
      printBackground: true,
      margin: { top: 0, right: 0, bottom: 0, left: 0 },
    });

    return {
      statusCode: 200,
      headers: {
        ...corsHeaders,
        'Content-Type': 'application/pdf',
        'Content-Disposition': `attachment; filename="${filename || 'TPL_SOW.pdf'}"`,
      },
      body: Buffer.from(pdf).toString('base64'),
      isBase64Encoded: true,
    };

  } catch (err) {
    console.error('PDF error:', err);
    return {
      statusCode: 500,
      headers: corsHeaders,
      body: JSON.stringify({ error: err.message }),
    };
  } finally {
    if (browser) await browser.close().catch(() => {});
  }
};
