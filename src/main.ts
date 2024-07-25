import PizZip from 'pizzip';
import { readFileSync, writeFileSync } from 'fs';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom'

// Fügt ein Bild in das ZIP-Archiv ein
function addImageToZip(zip: PizZip, imagePath: string): string {
    const imgData = readFileSync(imagePath);
    const imgExt = imagePath.split('.').pop();
    const imgFilename = `ppt/media/image${Object.keys(zip.files).length + 1}.${imgExt}`;
    zip.file(imgFilename, imgData);
    return imgFilename;
}

// Ersetzt Textplatzhalter im XML-Inhalt
function replaceTextPlaceholders(xmlContent: string, textPlaceholder: string, replacementText: string): string {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const doc = parser.parseFromString(xmlContent, 'application/xml');

    const textNodes = doc.getElementsByTagName('a:t');
    for (let i = 0; i < textNodes.length; i++) {
        const textNode = textNodes[i];
        if (textNode.textContent === textPlaceholder) {
            textNode.textContent = replacementText;
        }
    }

    return serializer.serializeToString(doc);
}

// Ersetzt Bildplatzhalter im XML-Inhalt
function replaceImagePlaceholders(xmlContent: string, imagePlaceholder: string, imageId: number): string {
    const parser = new DOMParser();
    const serializer = new XMLSerializer();
    const doc = parser.parseFromString(xmlContent, 'application/xml');

    const blipNodes = doc.getElementsByTagName('a:blip');
    for (let i = 0; i < blipNodes.length; i++) {
        const blipNode = blipNodes[i];
        const parentNode = blipNode.parentNode?.parentNode as HTMLElement;
        if (parentNode) {
            const textNodes = parentNode.getElementsByTagName('a:t');
            for (let j = 0; j < textNodes.length; j++) {
                const textNode = textNodes[j];
                if (textNode.textContent === imagePlaceholder) {
                    blipNode.setAttribute('r:embed', `rId${imageId}`);
                    textNode.textContent = '';  // Entferne den Textinhalt
                }
            }
        }
    }

    return serializer.serializeToString(doc);
}

// Verarbeite die PPTX-Datei
async function processPPTX(templatePath: string, outputPath: string, data: Record<string, string>): Promise<void> {
    const content = readFileSync(templatePath);
    const zip = new PizZip(content);

    let imageIdCounter = 1;

    // Verarbeite alle Folien
    const slideFiles = Object.keys(zip.files).filter(file => file.startsWith('ppt/slides/slide') && file.endsWith('.xml'));

    for (const slideFile of slideFiles) {
        let slideXml = zip.file(slideFile)?.asText() || '';

        // Ersetze Textplatzhalter
        for (const [key, value] of Object.entries(data)) {
            if (key === 'title') {
                slideXml = replaceTextPlaceholders(slideXml, `$${key}`, value);
            }
        }

        // Ersetze Bildplatzhalter
        for (const [key, imagePath] of Object.entries(data)) {
            if (key === 'logo') {
                addImageToZip(zip, imagePath);
                const imageId = imageIdCounter++;
                slideXml = replaceImagePlaceholders(slideXml, `$${key}`, imageId);
            }
        }

        zip.file(slideFile, slideXml);
    }

    // Füge die neue Medien-ID in die rels-Datei ein
    const relsFiles = Object.keys(zip.files).filter(file => file.endsWith('.xml.rels'));
    for (const relsFile of relsFiles) {
        let relsXml = zip.file(relsFile)?.asText() || '';
        relsXml = relsXml.replace(/<\/Relationships>/, `
            <Relationship Id="rId${imageIdCounter - 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${imageIdCounter - 1}"/>
            </Relationships>`);
        zip.file(relsFile, relsXml);
    }

    const modifiedContent = zip.generate({ type: 'nodebuffer' });
    writeFileSync(outputPath, modifiedContent);
}

// Beispielaufruf der Funktion
const templatePath = 'template.pptx';
const outputPath = 'output.pptx';
const data = {
    title: 'Dies ist der neue Titel',
    logo: 'example.png'
};

processPPTX(templatePath, outputPath, data)
    .then(() => console.log('PPTX-Datei erfolgreich verarbeitet!'))
    .catch(err => console.error('Fehler beim Verarbeiten der PPTX-Datei:', err));
