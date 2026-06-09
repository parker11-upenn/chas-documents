const fs = require('fs');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const mammoth = require('mammoth');
const TurndownService = require('turndown');

const turndownService = new TurndownService();

// Grab the payload payload sent over by Power Automate
const payload = JSON.parse(process.env.GITHUB_EVENT_PAYLOAD);
const filePath = payload.client_payload.filePath; 
const fileName = payload.client_payload.fileName;
const outputName = fileName.replace('.docx', '.md');

async function runRoundTrip() {
  try {
    const client = Client.init({
      authProvider: (done) => {
        done(null, process.env.MS_GRAPH_TOKEN); 
      },
    });

    console.log(`Fetching original file from SharePoint: ${filePath}`);
    
    // Download the raw binary string of the Word document
    const fileStream = await client
      .api(`/sites/CHAS-IR Redesign Project/drive/root:/${filePath}:/content`)
      .get();

    // Parse DOCX elements into clean text markdown
    const htmlResult = await mammoth.convertToHtml({ buffer: fileStream });
    const markdownText = turndownService.turn(htmlResult.value);

    console.log(`Conversion successful! Uploading ${outputName} back to SharePoint master...`);

    // Write the new text file right into your destination folder
    await client
      .api(`/sites/CHAS-IR Redesign Project/drive/root:/General/Training Docs/All Manuals in MD & PDF/${outputName}:/content`)
      .put(markdownText);

    console.log('Round-trip complete! Markdown file safely saved to SharePoint.');
  } catch (error) {
    console.error('Error during round-trip pipeline:', error);
    process.exit(1);
  }
}

runRoundTrip();