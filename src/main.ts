import AdmZip from 'adm-zip';
import * as xml2js from 'xml2js';
import * as fs from 'fs';
import * as path from 'path';

interface PlaceholderReplacements {
    [key: string]: string;
}

async function replacePlaceholdersInXml(xmlContent: string, replacements: PlaceholderReplacements) {
    const parser = new xml2js.Parser();
    const builder = new xml2js.Builder();

    const result = await parser.parseStringPromise(xmlContent);

    const replaceInObject = (obj: any) => {
        for (const key in obj) {
            if (typeof obj[key] === 'string') {
                for (const [placeholder, replacement] of Object.entries(replacements)) {
                    obj[key] = obj[key].replace(new RegExp(placeholder, 'g'), replacement);
                }
            } else if (typeof obj[key] === 'object') {
                replaceInObject(obj[key]);
            }
        }
    };

    replaceInObject(result);

    return builder.buildObject(result);
}

async function addImageToSlide(zip: AdmZip, slideEntryName: string, imagePath: string, relsEntryName: string) {
    const imageEntryName = `ppt/media/${path.basename(imagePath)}`;
    const imageData = fs.readFileSync(imagePath);
    zip.addFile(imageEntryName, imageData);

    const relsEntry = zip.getEntry(relsEntryName);
    let newRelId: string | undefined;

    if (relsEntry) {
        const relsContent = relsEntry.getData().toString('utf8');
        const parser = new xml2js.Parser();
        const builder = new xml2js.Builder();
        const relsXml = await parser.parseStringPromise(relsContent);

        if (!relsXml.Relationships) {
            relsXml.Relationships = {};
        }
        if (!relsXml.Relationships.Relationship) {
            relsXml.Relationships.Relationship = [];
        }

        newRelId = `rId${relsXml.Relationships.Relationship.length + 1}`;
        relsXml.Relationships.Relationship.push({
            $: {
                Id: newRelId,
                Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                Target: `../media/${path.basename(imagePath)}`
            }
        });

        const updatedRelsXml = builder.buildObject(relsXml);
        zip.updateFile(relsEntry.entryName, Buffer.from(updatedRelsXml, 'utf8'));
    }

    const slideEntry = zip.getEntry(slideEntryName);
    if (slideEntry && newRelId) {
        const slideContent = slideEntry.getData().toString('utf8');
        const parser = new xml2js.Parser();
        const builder = new xml2js.Builder();
        const slideXml = await parser.parseStringPromise(slideContent);

        // find und replace r:embed im p:cNvPr und a:blip to establish a successful relationship of the image
        slideXml['p:sld']['p:cSld'][0]['p:spTree'][0]['p:pic'].forEach((pic: any) => {
            if (pic['p:nvPicPr'] && pic['p:nvPicPr'][0]['p:cNvPr'] && pic['p:nvPicPr'][0]['p:cNvPr'][0].$.descr === 'LOGO_PLACEHOLDER') {
                pic['p:nvPicPr'][0]['p:cNvPr'][0].$['r:embed'] = newRelId;
            }
            if (pic['p:blipFill'] && pic['p:blipFill'][0]['a:blip'] && pic['p:blipFill'][0]['a:blip'][0].$['r:embed']) {
                pic['p:blipFill'][0]['a:blip'][0].$['r:embed'] = newRelId;
            }
        });


        const updatedSlideXml = builder.buildObject(slideXml);
        zip.updateFile(slideEntry.entryName, Buffer.from(updatedSlideXml, 'utf8'));
    }
}

async function replacePlaceholders(templatePath: string, outputPath: string, replacements: PlaceholderReplacements) {
    const zip = new AdmZip(templatePath);
    const zipEntries = zip.getEntries();

    for (const zipEntry of zipEntries) {
        if (zipEntry.entryName.endsWith('.xml')) {
            const xmlContent = zipEntry.getData().toString('utf8');
            const updatedXml = await replacePlaceholdersInXml(xmlContent, replacements);
            zip.updateFile(zipEntry.entryName, Buffer.from(updatedXml, 'utf8'));
        }
    }

    if (replacements['{{IMAGE_PLACEHOLDER}}']) {
        const imagePath = replacements['{{IMAGE_PLACEHOLDER}}'];
        const slideEntryName = 'ppt/slides/slide1.xml';
        const relsEntryName = 'ppt/slides/_rels/slide1.xml.rels';
        await addImageToSlide(zip, slideEntryName, imagePath, relsEntryName);
    }

    zip.writeZip(outputPath);
    console.log('Placeholder replaced and new presentation created!!');
}

const replacements: PlaceholderReplacements = {
    '{{TITLE_PLACEHOLDER}}': 'My wonderful title',
    '{{SUBTITLE_PLACEHOLDER}}': 'My wonderful subtitle',
    '{{IMAGE_PLACEHOLDER}}': 'example.jpg'
};

const templatePath = 'windows-template.pptx';
const outputPath = 'output.pptx';

replacePlaceholders(templatePath, outputPath, replacements).catch(err => {
    console.error('Error when replacing the placeholders:', err);
});
