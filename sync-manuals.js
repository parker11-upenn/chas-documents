const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const TurndownService = require('turndown');

const turndownService = new TurndownService();
const DOWNLOAD_DIR = path.join(__dirname, 'temp_docx');
const OUTPUT_DIR = path.join(__dirname, 'docs'); // Drops markdown directly into your local codebase docs folder

// Ensure folders exist
if (!fs.existsSync(DOWNLOAD_DIR)) fs.mkdirSync(DOWNLOAD_DIR);
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

async function runSyncPipeline() {
  console.log('Opening secure login window...');
  const browser = await chromium.launch({ headless: false }); // Runs visible browser
  const context = await browser.newContext();
  const page = await context.newPage();

  // Your exact UPenn SharePoint target folder URL
  const targetUrl = "https://penno365.sharepoint.com/teams/CHAS-IRRedesignProject59/Shared%20Documents/Forms/AllItems.aspx?id=%2Fteams%2FCHAS%2DIRRedesignProject59%2FShared%20Documents%2FGeneral%2FTraining%20Docs%2FAll%20Manuals%20in%20MD%20%26%20PDF";
  
  await page.goto(targetUrl);

  console.log('\n--- ACTION REQUIRED ---');
  console.log('1. Log into the browser window with your PennKey.');
  console.log('2. Complete your 2FA prompt.');
  console.log('3. Wait right there! The script will take over automatically once the files load.\n');
  
  // Waits for the file grid to render after you finish signing in
  await page.waitForSelector('[role="grid"]', { timeout: 300000 }); 
  console.log('Authentication verified! Processing files...');

  // Give the UI a split second to completely finish rendering rows
  await page.waitForTimeout(3000);

  // Grab all row elements containing files
  const fileRows = await page.$$('[role="row"]');
  let convertedCount = 0;
  
  for (const row of fileRows) {
    const textContent = await row.innerText();
    
    // Look specifically for your source Word files (.docx)
    if (textContent.includes('.docx')) {
      const fileName = textContent.split('\n')[0].trim();
      console.log(`Found source document: ${fileName}`);

      // Right-click the row to trigger the SharePoint download option
      await row.click({ button: 'right' });
      await page.waitForSelector('button:has-text("Download")', { timeout: 5000 });
      
      const downloadPromise = page.waitForEvent('download');
      await page.click('button:has-text("Download")');
      const download = await downloadPromise;

      // Stage the docx binary locally
      const docxPath = path.join(DOWNLOAD_DIR, fileName);
      await download.saveAs(docxPath);

      // Perform the local markdown compilation
      const mdName = fileName.replace('.docx', '.md');
      const mdOutputPath = path.join(OUTPUT_DIR, mdName);
      
      const buffer = fs.readFileSync(docxPath);
      const htmlResult = await mammoth.convertToHtml({ buffer: buffer });
      const markdownText = turndownService.turndown(htmlResult.value);

      fs.writeFileSync(mdOutputPath, markdownText);
      console.log(`✨ Successfully generated: ${mdName} -> dropped into ./docs/`);
      convertedCount++;
      
      // Click away to close any lingering context menus in the UI
      await page.mouse.click(10, 10);
      await page.waitForTimeout(1000);
    }
  }

  // Clean up temporary docx staging directory
  fs.rmSync(DOWNLOAD_DIR, { recursive: true, force: true });
  await browser.close();
  console.log(`\nPipeline complete! Successfully synced and compiled ${convertedCount} manuals to Markdown.`);
}

runSyncPipeline();