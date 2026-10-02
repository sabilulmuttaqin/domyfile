// Temporary QA repro generator: build a DOCX whose JPEG media entry is named
// "image1.undefined" (what the docx lib media writer emits when the entry type
// is lost). Run from repo root: node scripts/tmp-repro-undefined.mjs
// Output: tests/fixtures/docx-undefined-jpeg.docx
import { readFileSync, writeFileSync, rmSync } from "node:fs";
import { execSync } from "node:child_process";

const jpeg = readFileSync("tests/fixtures/compression-photo.jpg");

const contentTypes =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
  `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
  `<Default Extension="xml" ContentType="application/xml"/>` +
  `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
  `</Types>`;

const rels =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
  `</Relationships>`;

const docRels =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
  `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.undefined"/>` +
  `</Relationships>`;

const documentXml =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
  `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
  `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
  `<w:body><w:p><w:r><w:t>QA photo document</w:t></w:r></w:p>` +
  `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
  `<wp:extent cx="1905000" cy="1905000"/><wp:docPr id="1" name="Picture 1"/>` +
  `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
  `<pic:nvPicPr><pic:cNvPr id="1" name="Picture 1"/><pic:cNvPicPr/></pic:nvPicPr>` +
  `<pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="1905000"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
  `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
  `</w:body></w:document>`;

const dir = "/tmp/undefined-docx-src";
rmSync(dir, { recursive: true, force: true });
execSync(`mkdir -p ${dir}/word/media ${dir}/word/_rels ${dir}/_rels`);
writeFileSync(`${dir}/[Content_Types].xml`, contentTypes);
writeFileSync(`${dir}/_rels/.rels`, rels);
writeFileSync(`${dir}/word/_rels/document.xml.rels`, docRels);
writeFileSync(`${dir}/word/document.xml`, documentXml);
writeFileSync(`${dir}/word/media/image1.undefined`, jpeg);
writeFileSync("/tmp/mkzip.py", `import shutil\nshutil.make_archive('/tmp/docx-undefined-jpeg', 'zip', ${JSON.stringify(dir)})\n`);
execSync("python3 /tmp/mkzip.py");
execSync("mkdir -p tests/fixtures && mv /tmp/docx-undefined-jpeg.zip tests/fixtures/docx-undefined-jpeg.docx");

const out = readFileSync("tests/fixtures/docx-undefined-jpeg.docx");
console.log(`wrote tests/fixtures/docx-undefined-jpeg.docx (${out.byteLength} bytes, media entry word/media/image1.undefined = ${jpeg.byteLength} byte JPEG)`);
