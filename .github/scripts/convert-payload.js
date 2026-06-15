const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const TurndownService = require('turndown');
const turndownService = new TurndownService();

async function main() {
  // Grab the base64 text sent directly from Power Automate
  const base64Data = process.env.FILE_CONTENT_BASE64;
  const fileName = process.env.FILE_NAME;
  
  if (!base64Data || !fileName) {
    console.error("Missing payload environment variables!");
    process.exit(1);
  }

  const outputName = fileName.replace('.docx', '.md');
  
  // Rebuild the Word document binary right inside GitHub's container memory
  const buffer = Buffer.from(base64Data, 'base64');
  
  // Translate it directly to Markdown
  const htmlResult = await mammoth.convertToHtml({ buffer: buffer });
  const markdownText = turndownService.turndown(htmlResult.value);

  // Save it inside a docs folder
  if (!fs.existsSync('docs')) {
    fs.mkdirSync('docs');
  }
  
  fs.writeFileSync(`docs/${outputName}`, markdownText);
  console.log(`Successfully compiled docs/${outputName}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});