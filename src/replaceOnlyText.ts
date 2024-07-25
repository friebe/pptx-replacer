import fs from 'fs';
import PizZip from 'pizzip';

function replacePlaceholders(xmlContent: string, data: Record<string, string>): string {
  for (const [key, value] of Object.entries(data)) {
    const regex = new RegExp(`\\$${key}`, 'g');
    xmlContent = xmlContent.replace(regex, value);
  }
  return xmlContent;
}

async function processPPTX(templatePath: string, outputPath: string, data: Record<string, string>): Promise<void> {
  const content = fs.readFileSync(templatePath, 'binary');
  const zip = new PizZip(content);

  // get all available slides (ppt/slides/slide1.xml, ppt/slides/slide2.xml, usw.)
  const slideFiles = Object.keys(zip.files).filter(path => path.startsWith('ppt/slides/slide') && path.endsWith('.xml'));

  for (const slideFile of slideFiles) {
    // read XML content of slides
    let slideXml = zip.file(slideFile)?.asText() || '';

    // replace placeholder in slides
    slideXml = replacePlaceholders(slideXml, data);

    //return xml content into zip
    zip.file(slideFile, slideXml);
  }

  // create PPTX-File and save
  const modifiedContent = zip.generate({ type: 'nodebuffer' });
  fs.writeFileSync(outputPath, modifiedContent);
}

// Example usage
const templatePath = 'one-pager.pptx';
const outputPath = 'one-pager-output.pptx';
const data = {
  title: 'That is an example title'
};

processPPTX(templatePath, outputPath, data).then(() => {
  console.log('PPTX-Datei erfolgreich bearbeitet.');
}).catch(err => {
  console.error('Fehler beim Bearbeiten der PPTX-Datei:', err);
});
