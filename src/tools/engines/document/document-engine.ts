import type {
  DocumentCompressionPreset,
  DocumentEngine,
  DocumentOutputFormat,
  DocumentProcessItem,
  DocumentProcessOptions,
  DocumentProcessResult,
  DocumentToolSlug,
  DocumentValidationResult,
  DocumentValidationIssue,
} from "./types";
import { DocumentProcessingError } from "./types";

type MammothMessage = { type: string; message: string };
type MammothApi = {
  convertToHtml: (input: MammothInput, options?: Record<string, unknown>) => Promise<{ value: string; messages: MammothMessage[] }>;
  extractRawText: (input: MammothInput) => Promise<{ value: string; messages: MammothMessage[] }>;
};

type MammothInput = { arrayBuffer: ArrayBuffer } | { buffer: unknown };
type JsZipApi = typeof import("jszip");
type XmlDomApi = typeof import("@xmldom/xmldom");
type JsZipInstance = InstanceType<JsZipApi>;
type XmlNode = import("@xmldom/xmldom").Node;
type XmlElement = import("@xmldom/xmldom").Element;
type XmlDocument = import("@xmldom/xmldom").Document;
type ParsedXmlDocument = XmlDocument & { readonly documentElement: XmlElement };
type DocxFileChild = import("docx").FileChild;
type DocxParagraphChild = import("docx").ParagraphChild;

let mammothPromise: Promise<MammothApi> | undefined;
let docxPromise: Promise<typeof import("docx")> | undefined;
let jsZipPromise: Promise<JsZipApi> | undefined;
let xmlDomPromise: Promise<XmlDomApi> | undefined;

const loadMammoth = () => mammothPromise ??= import("mammoth").then((module) => (module.default ?? module) as MammothApi);
const loadDocx = () => docxPromise ??= import("docx");
const loadJsZip = async () => {
  if (!jsZipPromise) {
    jsZipPromise = import("jszip").then((module) => (module.default ?? module) as unknown as JsZipApi);
  }
  return jsZipPromise;
};
const loadXmlDom = () => xmlDomPromise ??= import("@xmldom/xmldom");

const getMammothInput = async (file: File): Promise<MammothInput> => {
  const arrayBuffer = await file.arrayBuffer();
  const nodeBuffer = (globalThis as typeof globalThis & { Buffer?: { from: (value: ArrayBuffer) => unknown } }).Buffer;
  return nodeBuffer ? { buffer: nodeBuffer.from(arrayBuffer) } : { arrayBuffer };
};

const bytesToArrayBuffer = (bytes: Uint8Array) => bytes.slice().buffer as ArrayBuffer;

const getExtension = (name: string) => name.toLowerCase().split(".").pop() ?? "";

const getSafeBaseName = (fileName: string) => {
  const withoutExtension = fileName.replace(/\.[^/.]+$/, "");
  const safe = withoutExtension
    .replace(/[\\/]+/g, "-")
    .replace(/[^\p{L}\p{N}._-]+/gu, "-")
    .replace(/-{2,}/g, "-")
    .replace(/^[-.]+|[-.]+$/g, "");
  return safe || "document";
};

const throwIfAborted = (signal?: AbortSignal) => {
  if (signal?.aborted) throw new DocumentProcessingError("CANCELLED", "Processing was cancelled.");
};

const asDocumentError = (error: unknown) => {
  if (error instanceof DocumentProcessingError) return error;
  const message = error instanceof Error ? error.message : "";
  if (/password|encrypted/i.test(message)) return new DocumentProcessingError("UNSUPPORTED_FEATURE", "Password-protected or encrypted documents are not supported by this browser tool.", error);
  if (/memory|quota/i.test(message)) return new DocumentProcessingError("BROWSER_UNSUPPORTED", "This document is too large for the available browser memory.", error);
  return new DocumentProcessingError("PROCESSING_FAILED", "The document could not be processed in this browser.", error);
};

export const getDocumentOutputName = (fileName: string, tool: DocumentToolSlug, index = 0, extension?: string) => {
  const base = getSafeBaseName(fileName);
  if (tool === "docx-to-txt") return `${base}-text.txt`;
  if (tool === "docx-to-html") return `${base}-html.html`;
  if (tool === "docx-to-pdf") return `${base}-pdf.pdf`;
  if (tool === "pdf-to-docx") return `${base}-editable.docx`;
  if (tool === "html-to-docx") return `${base}.docx`;
  if (tool === "merge-docx") return `${base}-merged.docx`;
  if (tool === "compress-docx") return `${base}-compressed.docx`;
  if (tool === "docx-metadata-cleaner") return `${base}-cleaned.docx`;
  if (tool === "extract-images-from-docx") return `${base}-image-${String(index + 1).padStart(2, "0")}.${extension ?? "bin"}`;
  return `${base}.docx`;
};

const getOutputFormat = (tool: DocumentToolSlug): DocumentOutputFormat => {
  if (tool === "docx-to-txt") return "txt";
  if (tool === "docx-to-html") return "html";
  if (tool === "docx-to-pdf") return "pdf";
  if (tool === "extract-images-from-docx") return "image";
  return "docx";
};

const getOutputMime = (format: DocumentOutputFormat) => {
  if (format === "txt") return "text/plain;charset=utf-8";
  if (format === "html") return "text/html;charset=utf-8";
  if (format === "pdf") return "application/pdf";
  if (format === "image") return "application/octet-stream";
  if (format === "zip") return "application/zip";
  return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
};

const hasZipSignature = (bytes: Uint8Array) => bytes.length >= 4
  && bytes[0] === 0x50
  && bytes[1] === 0x4b
  && ([0x03, 0x05, 0x07] as number[]).includes(bytes[2])
  && bytes[3] === 0x04;

const hasZipContainerSignature = (bytes: Uint8Array) => hasZipSignature(bytes)
  || (bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x05 && bytes[3] === 0x06);

const hasPdfSignature = (bytes: Uint8Array) => bytes.length >= 5
  && bytes[0] === 0x25
  && bytes[1] === 0x50
  && bytes[2] === 0x44
  && bytes[3] === 0x46
  && bytes[4] === 0x2d;

const readHeader = async (file: File, size = 12) => new Uint8Array(await file.slice(0, size).arrayBuffer());

const readZip = async (file: File) => {
  const JSZip = await loadJsZip();
  try {
    return await JSZip.loadAsync(await file.arrayBuffer(), { checkCRC32: true });
  } catch (error) {
    throw new DocumentProcessingError("CORRUPT_FILE", `“${file.name || "This file"}” is not a readable document archive.`, error);
  }
};

const ensureDocxArchive = async (file: File) => {
  if (!file || file.size <= 0) throw new DocumentProcessingError("INVALID_INPUT", "Choose a non-empty DOCX file to continue.");
  if (getExtension(file.name) !== "docx") throw new DocumentProcessingError("UNSUPPORTED_FORMAT", "This route accepts DOCX files only.");
  const header = await readHeader(file);
  if (!hasZipContainerSignature(header)) throw new DocumentProcessingError("CORRUPT_FILE", `“${file.name || "This file"}” is not a readable DOCX archive.`);
  const zip = await readZip(file);
  if (!zip.file("[Content_Types].xml") || !zip.file("word/document.xml")) {
    throw new DocumentProcessingError("CORRUPT_FILE", `“${file.name || "This file"}” is a ZIP but not a complete DOCX document.`);
  }
  return zip;
};

const ensurePdfFile = async (file: File) => {
  if (!file || file.size <= 0) throw new DocumentProcessingError("INVALID_INPUT", "Choose a non-empty PDF file to continue.");
  if (getExtension(file.name) !== "pdf") throw new DocumentProcessingError("UNSUPPORTED_FORMAT", "This route accepts PDF files only.");
  if (!hasPdfSignature(await readHeader(file, 8))) throw new DocumentProcessingError("CORRUPT_FILE", `“${file.name || "This file"}” is not a readable PDF.`);
};

const ensureTxtFile = async (file: File) => {
  if (!file || file.size <= 0) throw new DocumentProcessingError("INVALID_INPUT", "Choose a non-empty TXT file to continue.");
  if (getExtension(file.name) !== "txt") throw new DocumentProcessingError("UNSUPPORTED_FORMAT", "This route accepts TXT files only.");
  const value = await file.text();
  if (value.includes("\u0000")) throw new DocumentProcessingError("CORRUPT_FILE", `“${file.name || "This file"}” contains binary data, not plain text.`);
  if (!value.trim()) throw new DocumentProcessingError("NO_READABLE_CONTENT", "This TXT file is empty. Add some text before creating a DOCX.");
};

const ensureHtmlFile = async (file: File) => {
  if (!file || file.size <= 0) throw new DocumentProcessingError("INVALID_INPUT", "Choose a non-empty HTML file to continue.");
  if (!["html", "htm"].includes(getExtension(file.name))) throw new DocumentProcessingError("UNSUPPORTED_FORMAT", "This route accepts HTML files only.");
  const value = await file.text();
  if (!value.replace(/<[^>]*>/g, "").trim()) throw new DocumentProcessingError("NO_READABLE_CONTENT", "This HTML file does not contain readable content.");
};

const validateFile = async (file: File, tool: DocumentToolSlug) => {
  if (tool === "pdf-to-docx") return ensurePdfFile(file);
  if (tool === "txt-to-docx") return ensureTxtFile(file);
  if (tool === "html-to-docx") return ensureHtmlFile(file);
  return ensureDocxArchive(file);
};

const getMammothMessages = (messages: MammothMessage[]) => messages
  .filter((message) => message.type === "error")
  .map((message) => message.message)
  .join(" ");

const escapeHtml = (value: string) => value
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;")
  .replace(/'/g, "&#39;");

const createHtmlDocument = (title: string, body: string) => `<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<title>${escapeHtml(title)}</title>\n</head>\n<body>\n${body}\n</body>\n</html>\n`;

const createDocxOutput = async (file: File) => {
  const value = (await file.text()).replace(/^\uFEFF/, "").replace(/\r\n?/g, "\n");
  const module = await loadDocx();
  const document = new module.Document({ sections: [{ children: value.split("\n").map((line) => new module.Paragraph(line)) }] });
  return module.Packer.toBlob(document);
};

const createDocxToTextOutput = async (file: File, signal?: AbortSignal) => {
  throwIfAborted(signal);
  const mammoth = await loadMammoth();
  const result = await mammoth.extractRawText(await getMammothInput(file));
  throwIfAborted(signal);
  const errors = getMammothMessages(result.messages);
  if (errors) throw new DocumentProcessingError("CORRUPT_FILE", `This DOCX could not be read: ${errors}`);
  if (!result.value.trim()) throw new DocumentProcessingError("NO_READABLE_CONTENT", "This DOCX contains no readable paragraph text.");
  return new Blob([result.value], { type: "text/plain;charset=utf-8" });
};

const createDocxToHtmlOutput = async (file: File, signal?: AbortSignal) => {
  throwIfAborted(signal);
  const mammoth = await loadMammoth();
  const result = await mammoth.convertToHtml(await getMammothInput(file), {
    externalFileAccess: false,
    includeDefaultStyleMap: true,
  });
  throwIfAborted(signal);
  const errors = getMammothMessages(result.messages);
  if (errors) throw new DocumentProcessingError("CORRUPT_FILE", `This DOCX could not be converted: ${errors}`);
  if (!result.value.trim()) throw new DocumentProcessingError("NO_READABLE_CONTENT", "This DOCX contains no readable content to convert to HTML.");
  return new Blob([createHtmlDocument(getSafeBaseName(file.name), result.value)], { type: "text/html;charset=utf-8" });
};

const htmlAllowedTags = new Set(["a", "b", "br", "em", "h1", "h2", "h3", "h4", "h5", "h6", "i", "li", "ol", "p", "strong", "table", "tbody", "td", "th", "thead", "tr", "u", "ul"]);
const htmlContainerTags = new Set(["article", "body", "div", "main", "section", "header", "footer", "span"]);
const htmlBlockedTags = new Set(["audio", "embed", "iframe", "object", "script", "style", "svg", "template", "video"]);

type HtmlNode = {
  nodeType: number;
  nodeName?: string;
  tagName?: string;
  nodeValue?: string | null;
  textContent?: string | null;
  childNodes?: { length: number; item: (index: number) => HtmlNode | null };
  getAttribute?: (name: string) => string;
};

const elementName = (node: HtmlNode) => (node.tagName ?? node.nodeName ?? "").toLowerCase();

const childNodes = (node: HtmlNode) => {
  const children = node.childNodes;
  if (!children) return [] as HtmlNode[];
  const output: HtmlNode[] = [];
  for (let index = 0; index < children.length; index += 1) {
    const child = children.item(index);
    if (child) output.push(child);
  }
  return output;
};

const safeHref = (value: string) => /^(?:https?:|mailto:)/i.test(value) ? value : "";

const parseHtmlRoot = async (value: string): Promise<HtmlNode> => {
  if (typeof DOMParser !== "undefined") {
    const parsed = new DOMParser().parseFromString(value, "text/html");
    return (parsed.body as unknown as HtmlNode) ?? (parsed as unknown as HtmlNode);
  }
  const xml = await loadXmlDom();
  const normalized = value
    .replace(/<br\s*>/gi, "<br/>")
    .replace(/<hr\s*>/gi, "<hr/>")
    .replace(/<img([^>]*?)(?<!\/)\s*>/gi, "<img$1/>")
    .replace(/&nbsp;/gi, " ");
  const parsed = new xml.DOMParser().parseFromString(`<root>${normalized}</root>`, "text/xml");
  return parsed.documentElement as unknown as HtmlNode;
};

const hasReadableHtml = (node: HtmlNode): boolean => {
  if (node.nodeType === 3) return Boolean(node.nodeValue?.trim());
  if (htmlBlockedTags.has(elementName(node))) return false;
  return childNodes(node).some(hasReadableHtml);
};

const collectSafeHtml = (node: HtmlNode, output: HtmlNode[] = []) => {
  if (node.nodeType === 3) {
    output.push(node);
    return output;
  }
  const name = elementName(node);
  if (htmlBlockedTags.has(name) || name === "img") return output;
  if (htmlAllowedTags.has(name)) {
    output.push(node);
    return output;
  }
  if (htmlContainerTags.has(name) || name === "html") childNodes(node).forEach((child) => collectSafeHtml(child, output));
  return output;
};

const getNodeText = (node: HtmlNode) => node.textContent ?? node.nodeValue ?? "";

const createHtmlToDocxOutput = async (file: File, signal?: AbortSignal) => {
  throwIfAborted(signal);
  const root = await parseHtmlRoot(await file.text());
  if (!hasReadableHtml(root)) throw new DocumentProcessingError("NO_READABLE_CONTENT", "This HTML file has no readable content after unsafe elements are removed.");
  const module = await loadDocx();
  const numberingConfig = [{
    reference: "domyfile-html-numbered",
    levels: [{ level: 0, format: module.LevelFormat.DECIMAL, text: "%1.", alignment: module.AlignmentType.LEFT }],
  }];
  const children: DocxFileChild[] = [];

  const makeRuns = (node: HtmlNode, state: { bold?: boolean; italics?: boolean; underline?: boolean } = {}): DocxParagraphChild[] => {
    if (node.nodeType === 3) {
      const text = node.nodeValue ?? "";
      return text ? [new module.TextRun({ text, bold: state.bold, italics: state.italics, underline: state.underline ? {} : undefined })] : [];
    }
    const name = elementName(node);
    if (htmlBlockedTags.has(name) || name === "img") return [];
    const nextState = {
      bold: state.bold || name === "b" || name === "strong",
      italics: state.italics || name === "i" || name === "em",
      underline: state.underline || name === "u",
    };
    if (name === "br") return [new module.TextRun({ break: 1 })];
    if (name === "a") {
      const href = safeHref(node.getAttribute?.("href") ?? "");
      const linkChildren = childNodes(node).flatMap((child) => makeRuns(child, nextState));
      return href && linkChildren.length ? [new module.ExternalHyperlink({ link: href, children: linkChildren })] : linkChildren;
    }
    return childNodes(node).flatMap((child) => makeRuns(child, nextState));
  };

  const makeParagraph = (node: HtmlNode, options: Record<string, unknown> = {}) => {
    const runs: DocxParagraphChild[] = makeRuns(node);
    const text = getNodeText(node).replace(/\s+/g, " ").trim();
    if (!runs.length && text) runs.push(new module.TextRun(text));
    if (!runs.length) return undefined;
    return new module.Paragraph({ ...options, children: runs });
  };

  const makeTable = (node: HtmlNode) => {
    const rows = childNodes(node).flatMap((section) => elementName(section) === "tr" ? [section] : childNodes(section).filter((row) => elementName(row) === "tr"));
    const tableRows = rows.map((row) => {
      const cells = childNodes(row).filter((cell) => elementName(cell) === "td" || elementName(cell) === "th");
      if (!cells.length) return undefined;
      return new module.TableRow({ children: cells.map((cell) => new module.TableCell({ children: [makeParagraph(cell) ?? new module.Paragraph("")] })) });
    }).filter((row): row is NonNullable<typeof row> => Boolean(row));
    return tableRows.length ? new module.Table({ rows: tableRows }) : undefined;
  };

  const visitBlocks = (node: HtmlNode, listType?: "ul" | "ol") => {
    const name = elementName(node);
    if (name === "table") {
      const table = makeTable(node);
      if (table) children.push(table);
      return;
    }
    if (name === "li") {
      const paragraph = makeParagraph(node, listType === "ol" ? { numbering: { reference: "domyfile-html-numbered", level: 0 } } : { bullet: { level: 0 } });
      if (paragraph) children.push(paragraph);
      return;
    }
    if (/^h[1-6]$/.test(name)) {
      const level = Number(name.slice(1));
      const heading = [module.HeadingLevel.HEADING_1, module.HeadingLevel.HEADING_2, module.HeadingLevel.HEADING_3, module.HeadingLevel.HEADING_4, module.HeadingLevel.HEADING_5, module.HeadingLevel.HEADING_6][level - 1];
      const paragraph = makeParagraph(node, { heading });
      if (paragraph) children.push(paragraph);
      return;
    }
    if (name === "p") {
      const paragraph = makeParagraph(node);
      if (paragraph) children.push(paragraph);
      return;
    }
    if (name === "ul" || name === "ol") {
      childNodes(node).filter((child) => elementName(child) === "li").forEach((child) => visitBlocks(child, name));
      return;
    }
    if (htmlBlockedTags.has(name)) return;
    childNodes(node).forEach((child) => {
      const childName = elementName(child);
      if (["p", "table", "ul", "ol", "li"].includes(childName) || /^h[1-6]$/.test(childName)) visitBlocks(child, listType);
      else if (child.nodeType === 3 && child.nodeValue?.trim()) {
        const paragraph = makeParagraph(child);
        if (paragraph) children.push(paragraph);
      } else visitBlocks(child, listType);
    });
  };

  collectSafeHtml(root).forEach((node) => visitBlocks(node));
  throwIfAborted(signal);
  if (!children.length) throw new DocumentProcessingError("NO_READABLE_CONTENT", "This HTML file has no supported headings, paragraphs, lists, or tables.");
  const document = new module.Document({ numbering: { config: numberingConfig }, sections: [{ children }] });
  return module.Packer.toBlob(document);
};

const getXmlParser = async () => {
  const xml = await loadXmlDom();
  return new xml.DOMParser();
};

const parseXml = async (value: string): Promise<ParsedXmlDocument> => {
  const parser = await getXmlParser();
  const document = parser.parseFromString(value, "application/xml");
  if (!document.documentElement || document.getElementsByTagName("parsererror").length) throw new DocumentProcessingError("CORRUPT_FILE", "The DOCX contains malformed XML.");
  return document as ParsedXmlDocument;
};

const serializeXml = async (node: XmlNode) => {
  const xml = await loadXmlDom();
  return new xml.XMLSerializer().serializeToString(node);
};

const elementChildren = (node: XmlNode, name?: string) => {
  const output: XmlElement[] = [];
  for (let index = 0; index < node.childNodes.length; index += 1) {
    const child = node.childNodes.item(index);
    if (!child || child.nodeType !== 1) continue;
    const element = child as XmlElement;
    if (!name || element.tagName === name) output.push(element);
  }
  return output;
};

const firstElement = (node: XmlElement | ParsedXmlDocument, name: string) => {
  const found = node.getElementsByTagName(name);
  if (found.length) return found.item(0) as XmlElement;
  const localName = name.split(":").pop() ?? name;
  const elements = node.getElementsByTagName("*");
  for (let index = 0; index < elements.length; index += 1) {
    const element = elements.item(index) as XmlElement | null;
    if (!element) continue;
    const elementLocalName = element.localName ?? element.tagName.split(":").pop() ?? element.tagName;
    if (elementLocalName === localName) return element;
  }
  return undefined;
};

const safePackagePath = (target: string) => {
  const normalized = target.replace(/\\/g, "/").replace(/^\.\//, "");
  if (normalized.startsWith("../")) return normalized.replace(/^\.\.\//, "");
  return normalized.startsWith("word/") ? normalized : `word/${normalized}`;
};

const getMediaMime = (bytes: Uint8Array, extension: string) => {
  if (bytes.length >= 8 && bytes.slice(0, 8).every((value, index) => value === [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a][index])) return "image/png";
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "image/jpeg";
  if (bytes.length >= 12 && String.fromCharCode(...bytes.slice(0, 4)) === "RIFF" && String.fromCharCode(...bytes.slice(8, 12)) === "WEBP") return "image/webp";
  if (extension === "gif" && bytes.length >= 6 && ["GIF89a", "GIF87a"].includes(String.fromCharCode(...bytes.slice(0, 6)))) return "image/gif";
  return undefined;
};

const toDocumentItem = (input: File, output: Blob, format: DocumentOutputFormat, name: string, metadata?: DocumentProcessItem["metadata"]) => {
  const file = new File([output], name, { type: getOutputMime(format) });
  return { input, output: file, format, inputBytes: input.size, outputBytes: file.size, metadata } satisfies DocumentProcessItem;
};

const validateOutput = async (output: File, format: DocumentOutputFormat) => {
  if (!output.size) throw new DocumentProcessingError("PROCESSING_FAILED", "The browser returned an empty document result.");
  const header = await readHeader(output, 8);
  if (format === "docx") {
    if (!hasZipContainerSignature(header)) throw new DocumentProcessingError("PROCESSING_FAILED", "The browser did not create a valid DOCX archive.");
    const zip = await readZip(output);
    if (!zip.file("word/document.xml")) throw new DocumentProcessingError("PROCESSING_FAILED", "The browser created an incomplete DOCX archive.");
  }
  if (format === "pdf" && !hasPdfSignature(header)) throw new DocumentProcessingError("PROCESSING_FAILED", "The browser did not create a valid PDF result.");
  if (format === "zip" && !hasZipContainerSignature(header)) throw new DocumentProcessingError("PROCESSING_FAILED", "The browser did not create a valid ZIP download.");
  if (format === "html" && !(await output.text()).includes("<body>")) throw new DocumentProcessingError("PROCESSING_FAILED", "The browser did not create a complete HTML document.");
};

const getMetadataMappings = () => new Map([
  ["creator", "Author"], ["lastModifiedBy", "Last modified by"], ["title", "Title"], ["subject", "Subject"],
  ["keywords", "Keywords"], ["description", "Description"], ["category", "Category"], ["contentStatus", "Content status"],
  ["language", "Language"], ["revision", "Revision"], ["created", "Created date"], ["modified", "Modified date"],
  ["Application", "Application"], ["Company", "Company"], ["Manager", "Manager"], ["AppVersion", "Application version"],
]);

const getMetadataFromZip = async (zip: JsZipInstance) => {
  const detected: string[] = [];
  const removable: string[] = [];
  const mappings = getMetadataMappings();
  for (const [path, tags] of [["docProps/core.xml", ["creator", "lastModifiedBy", "title", "subject", "keywords", "description", "category", "contentStatus", "language", "revision", "created", "modified"]], ["docProps/app.xml", ["Application", "Company", "Manager", "AppVersion"]]] as const) {
    const entry = zip.file(path);
    if (!entry) continue;
    const document = await parseXml(await entry.async("text"));
    tags.forEach((tag) => {
      const element = firstElement(document, tag);
      if (element?.textContent?.trim()) {
        const label = mappings.get(tag) ?? tag;
        detected.push(label);
        removable.push(label);
      }
    });
  }
  const custom = zip.file("docProps/custom.xml");
  if (custom) {
    const document = await parseXml(await custom.async("text"));
    const properties = firstElement(document, "Properties");
    if (properties && elementChildren(properties, "property").length) {
      detected.push("Custom document properties");
      removable.push("Custom document properties");
    }
  }
  return { detected, removable };
};

const removeElementsByName = (document: XmlElement | ParsedXmlDocument, names: Set<string>) => {
  const all = Array.from({ length: document.getElementsByTagName("*").length }, (_, index) => document.getElementsByTagName("*").item(index)).filter((element): element is XmlElement => Boolean(element));
  all.forEach((element) => {
    const localName = element.localName ?? element.tagName.split(":").pop() ?? "";
    if (names.has(localName)) element.parentNode?.removeChild(element);
  });
};

const cleanMetadata = async (file: File, signal?: AbortSignal) => {
  const zip = await ensureDocxArchive(file);
  const metadata = await getMetadataFromZip(zip);
  if (!metadata.removable.length) throw new DocumentProcessingError("NO_READABLE_CONTENT", "No supported document properties were found to remove.");
  const core = zip.file("docProps/core.xml");
  if (core) {
    const document = await parseXml(await core.async("text"));
    removeElementsByName(document, new Set(["creator", "lastModifiedBy", "title", "subject", "keywords", "description", "category", "contentStatus", "language", "revision", "created", "modified"]));
    zip.file("docProps/core.xml", await serializeXml(document));
  }
  const app = zip.file("docProps/app.xml");
  if (app) {
    const document = await parseXml(await app.async("text"));
    removeElementsByName(document, new Set(["Application", "Company", "Manager", "AppVersion"]));
    zip.file("docProps/app.xml", await serializeXml(document));
  }
  const custom = zip.file("docProps/custom.xml");
  if (custom) {
    const document = await parseXml(await custom.async("text"));
    const properties = firstElement(document, "Properties");
    elementChildren(properties ?? document, "property").forEach((property) => property.parentNode?.removeChild(property));
    zip.file("docProps/custom.xml", await serializeXml(document));
  }
  throwIfAborted(signal);
  const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 9 } });
  return { blob, metadata: { detected: metadata.detected, removed: metadata.removable } };
};

const getNextRelationshipId = (document: ParsedXmlDocument) => {
  const used = new Set(elementChildren(document.documentElement).map((node) => node.getAttribute("Id")));
  let index = 1;
  while (used.has(`rId${index}`)) index += 1;
  return `rId${index}`;
};

const getNextNumberId = (document: ParsedXmlDocument, elementNameValue: string, attribute: string) => {
  const values = Array.from({ length: document.getElementsByTagName(elementNameValue).length }, (_, index) => document.getElementsByTagName(elementNameValue).item(index) as XmlElement | null)
    .map((element) => Number(element?.getAttribute(attribute)))
    .filter((value) => Number.isFinite(value));
  return Math.max(-1, ...values) + 1;
};

const replaceRelationshipAttributes = (node: XmlElement, map: Map<string, string>) => {
  const elements = [node as XmlElement, ...Array.from({ length: node.getElementsByTagName("*").length }, (_, index) => node.getElementsByTagName("*").item(index) as XmlElement | null).filter((element): element is XmlElement => Boolean(element))];
  elements.forEach((element) => ["r:id", "r:embed", "r:link"].forEach((attribute) => {
    const value = element.getAttribute(attribute);
    if (value && map.has(value)) element.setAttribute(attribute, map.get(value) ?? value);
  }));
};

const remapNumbering = async (baseZip: JsZipInstance, sourceZip: JsZipInstance, fragment: XmlElement) => {
  const sourceEntry = sourceZip.file("word/numbering.xml");
  if (!sourceEntry) return;
  const source = await parseXml(await sourceEntry.async("text"));
  const baseEntry = baseZip.file("word/numbering.xml");
  if (!baseEntry) {
    baseZip.file("word/numbering.xml", await sourceEntry.async("uint8array"));
    return;
  }
  const base = await parseXml(await baseEntry.async("text"));
  const abstractMap = new Map<string, string>();
  const numberMap = new Map<string, string>();
  let nextAbstract = getNextNumberId(base, "w:abstractNum", "w:abstractNumId");
  let nextNumber = getNextNumberId(base, "w:num", "w:numId");
  elementChildren(source.documentElement, "w:abstractNum").forEach((abstractNum) => {
    const old = abstractNum.getAttribute("w:abstractNumId");
    if (!old) return;
    const next = String(nextAbstract++);
    abstractMap.set(old, next);
    const clone = abstractNum.cloneNode(true) as XmlElement;
    clone.setAttribute("w:abstractNumId", next);
    base.documentElement.appendChild(clone);
  });
  elementChildren(source.documentElement, "w:num").forEach((numbering) => {
    const old = numbering.getAttribute("w:numId");
    if (!old) return;
    const next = String(nextNumber++);
    numberMap.set(old, next);
    const clone = numbering.cloneNode(true) as XmlElement;
    clone.setAttribute("w:numId", next);
    const abstractRef = firstElement(clone, "w:abstractNumId");
    if (abstractRef) {
      const oldAbstract = abstractRef.getAttribute("w:val") ?? "";
      abstractRef.setAttribute("w:val", abstractMap.get(oldAbstract) ?? oldAbstract);
    }
    base.documentElement.appendChild(clone);
  });
  const numIds = fragment.getElementsByTagName("w:numId");
  for (let index = 0; index < numIds.length; index += 1) {
    const element = numIds.item(index) as XmlElement;
    const oldNumber = element.getAttribute("w:val") ?? "";
    element.setAttribute("w:val", numberMap.get(oldNumber) ?? oldNumber);
  }
  baseZip.file("word/numbering.xml", await serializeXml(base));
};

const mergeStyles = async (baseZip: JsZipInstance, sourceZip: JsZipInstance) => {
  const sourceEntry = sourceZip.file("word/styles.xml");
  if (!sourceEntry) return;
  const baseEntry = baseZip.file("word/styles.xml");
  if (!baseEntry) {
    baseZip.file("word/styles.xml", await sourceEntry.async("uint8array"));
    return;
  }
  const base = await parseXml(await baseEntry.async("text"));
  const source = await parseXml(await sourceEntry.async("text"));
  const existing = new Set(elementChildren(base.documentElement, "w:style").map((style) => style.getAttribute("w:styleId")));
  elementChildren(source.documentElement, "w:style").forEach((style) => {
    const id = style.getAttribute("w:styleId");
    if (id && !existing.has(id)) {
      base.documentElement.appendChild(style.cloneNode(true));
      existing.add(id);
    }
  });
  baseZip.file("word/styles.xml", await serializeXml(base));
};

const mergeContentTypes = async (baseZip: JsZipInstance, sourceZip: JsZipInstance, extensions: Set<string>) => {
  const sourceEntry = sourceZip.file("[Content_Types].xml");
  const baseEntry = baseZip.file("[Content_Types].xml");
  if (!sourceEntry || !baseEntry) return;
  const base = await parseXml(await baseEntry.async("text"));
  const source = await parseXml(await sourceEntry.async("text"));
  const existing = new Set(elementChildren(base.documentElement, "Default").map((item) => item.getAttribute("Extension")));
  extensions.forEach((extension) => {
    if (existing.has(extension)) return;
    const sourceDefault = elementChildren(source.documentElement, "Default").find((item) => item.getAttribute("Extension") === extension);
    if (sourceDefault) base.documentElement.appendChild(sourceDefault.cloneNode(true));
  });
  baseZip.file("[Content_Types].xml", await serializeXml(base));
};

const mergeDocxOutput = async (files: File[], signal?: AbortSignal) => {
  if (files.length < 2) throw new DocumentProcessingError("INVALID_INPUT", "Choose at least two DOCX files to merge.");
  const zips = await Promise.all(files.map((file) => ensureDocxArchive(file)));
  const baseZip = zips[0];
  const baseDocument = await parseXml(await baseZip.file("word/document.xml")!.async("text"));
  const baseRelations = baseZip.file("word/_rels/document.xml.rels")
    ? await parseXml(await baseZip.file("word/_rels/document.xml.rels")!.async("text"))
    : undefined;
  if (!baseRelations) throw new DocumentProcessingError("UNSUPPORTED_FEATURE", "The first DOCX has no document relationship manifest, so it cannot be merged safely.");
  const baseBody = firstElement(baseDocument, "w:body");
  if (!baseBody) throw new DocumentProcessingError("CORRUPT_FILE", "The first DOCX has no document body.");
  const baseSection = elementChildren(baseBody, "w:sectPr")[0];
  const baseRelationRoot = baseRelations.documentElement;
  const existingMedia = new Set(Object.keys(baseZip.files).filter((name) => name.startsWith("word/media/")));

  for (let fileIndex = 1; fileIndex < zips.length; fileIndex += 1) {
    throwIfAborted(signal);
    const sourceZip = zips[fileIndex];
    const sourceDocument = await parseXml(await sourceZip.file("word/document.xml")!.async("text"));
    const sourceBody = firstElement(sourceDocument, "w:body");
    if (!sourceBody) throw new DocumentProcessingError("CORRUPT_FILE", `“${files[fileIndex].name}” has no document body.`);
    const sourceRelationsEntry = sourceZip.file("word/_rels/document.xml.rels");
    const sourceRelations = sourceRelationsEntry ? await parseXml(await sourceRelationsEntry.async("text")) : undefined;
    const sourceRelationById = new Map<string, XmlElement>();
    if (sourceRelations) elementChildren(sourceRelations.documentElement, "Relationship").forEach((relation) => {
      const id = relation.getAttribute("Id");
      if (id) sourceRelationById.set(id, relation);
    });
    const fragmentNodes = elementChildren(sourceBody).filter((node) => node.tagName !== "w:sectPr").map((node) => node.cloneNode(true) as XmlElement);
    const fragmentRoot = baseDocument.createElement("w:fragment");
    fragmentNodes.forEach((node) => fragmentRoot.appendChild(node));
    const relationIds = new Set<string>();
    const elements = [fragmentRoot as XmlElement, ...Array.from({ length: fragmentRoot.getElementsByTagName("*").length }, (_, index) => fragmentRoot.getElementsByTagName("*").item(index) as XmlElement | null).filter((element): element is XmlElement => Boolean(element))];
    elements.forEach((element) => ["r:id", "r:embed", "r:link"].forEach((attribute) => {
      const id = element.getAttribute(attribute);
      if (id) relationIds.add(id);
    }));
    const relationMap = new Map<string, string>();
    const mediaExtensions = new Set<string>();
    for (const oldId of relationIds) {
      const sourceRelation = sourceRelationById.get(oldId);
      if (!sourceRelation) throw new DocumentProcessingError("CORRUPT_FILE", `“${files[fileIndex].name}” references a missing document relationship.`);
      const target = sourceRelation.getAttribute("Target");
      const type = sourceRelation.getAttribute("Type");
      const mode = sourceRelation.getAttribute("TargetMode");
      if (!target || !type) throw new DocumentProcessingError("CORRUPT_FILE", `“${files[fileIndex].name}” contains an incomplete document relationship.`);
      const newId = getNextRelationshipId(baseRelations);
      const relationship = sourceRelation.cloneNode(true) as XmlElement;
      relationship.setAttribute("Id", newId);
      if (mode === "External") {
        baseRelationRoot.appendChild(relationship);
        relationMap.set(oldId, newId);
        continue;
      }
      if (!/\/image$/i.test(type)) throw new DocumentProcessingError("UNSUPPORTED_FEATURE", `“${files[fileIndex].name}” contains a linked document feature that cannot be merged safely. Images and hyperlinks are supported.`);
      const sourcePath = safePackagePath(target);
      const sourceFile = sourceZip.file(sourcePath);
      if (!sourceFile) throw new DocumentProcessingError("CORRUPT_FILE", `“${files[fileIndex].name}” references a missing embedded image.`);
      const originalName = sourcePath.split("/").pop() ?? `image-${fileIndex}`;
      let newTargetName = `media/merge-${fileIndex}-${originalName}`;
      let suffix = 1;
      while (existingMedia.has(`word/${newTargetName}`)) newTargetName = `media/merge-${fileIndex}-${suffix++}-${originalName}`;
      existingMedia.add(`word/${newTargetName}`);
      baseZip.file(`word/${newTargetName}`, await sourceFile.async("uint8array"));
      const extension = getExtension(originalName);
      if (extension) mediaExtensions.add(extension);
      relationship.setAttribute("Target", newTargetName);
      baseRelationRoot.appendChild(relationship);
      relationMap.set(oldId, newId);
    }
    replaceRelationshipAttributes(fragmentRoot, relationMap);
    await remapNumbering(baseZip, sourceZip, fragmentRoot);
    await mergeStyles(baseZip, sourceZip);
    await mergeContentTypes(baseZip, sourceZip, mediaExtensions);
    if (baseSection) {
      const pageBreak = baseDocument.createElement("w:p");
      const run = baseDocument.createElement("w:r");
      const breakNode = baseDocument.createElement("w:br");
      breakNode.setAttribute("w:type", "page");
      run.appendChild(breakNode);
      pageBreak.appendChild(run);
      baseBody.insertBefore(pageBreak, baseSection);
    }
    fragmentNodes.forEach((node) => baseSection ? baseBody.insertBefore(node, baseSection) : baseBody.appendChild(node));
  }
  baseZip.file("word/document.xml", await serializeXml(baseDocument));
  baseZip.file("word/_rels/document.xml.rels", await serializeXml(baseRelations));
  return baseZip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 9 } });
};

const getImageEntries = async (zip: JsZipInstance) => Object.entries(zip.files)
  .filter(([name, entry]) => name.startsWith("word/media/") && !entry.dir)
  .map(([name, entry]) => ({ name, entry }));

const compressJpeg = async (bytes: Uint8Array, quality: number, signal?: AbortSignal) => {
  if (typeof createImageBitmap !== "function") return undefined;
  let bitmap: ImageBitmap | undefined;
  try {
    bitmap = await createImageBitmap(new Blob([bytesToArrayBuffer(bytes)], { type: "image/jpeg" }));
    throwIfAborted(signal);
    const canvas = typeof OffscreenCanvas !== "undefined" ? new OffscreenCanvas(bitmap.width, bitmap.height) : document.createElement("canvas");
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    const context = canvas.getContext("2d");
    if (!context) return undefined;
    context.drawImage(bitmap, 0, 0);
    const blob = typeof OffscreenCanvas !== "undefined" && canvas instanceof OffscreenCanvas
      ? await canvas.convertToBlob({ type: "image/jpeg", quality })
      : await new Promise<Blob | undefined>((resolve) => (canvas as HTMLCanvasElement).toBlob((blob) => resolve(blob ?? undefined), "image/jpeg", quality));
    if (!blob || blob.size >= bytes.byteLength) return undefined;
    return new Uint8Array(await blob.arrayBuffer());
  } finally {
    bitmap?.close();
  }
};

const compressDocxOutput = async (file: File, preset: DocumentCompressionPreset, signal?: AbortSignal) => {
  const zip = await ensureDocxArchive(file);
  const quality = preset === "light" ? 0.9 : preset === "strong" ? 0.62 : 0.78;
  for (const { name, entry } of await getImageEntries(zip)) {
    throwIfAborted(signal);
    const bytes = await entry.async("uint8array");
    const mediaMime = getMediaMime(bytes, getExtension(name));
    if (mediaMime !== "image/jpeg") continue;
    const compressed = await compressJpeg(bytes, quality, signal);
    if (compressed && compressed.byteLength < bytes.byteLength * 0.98) zip.file(name, compressed);
  }
  const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: preset === "light" ? 6 : preset === "balanced" ? 8 : 9 } });
  if (blob.size >= file.size * 0.995) throw new DocumentProcessingError("NO_USEFUL_REDUCTION", "No useful reduction found. The document may already be compressed or may not contain recompressible JPEG media.");
  return blob;
};

const mimeExtension = (mime: string) => ({ "image/jpeg": "jpg", "image/png": "png", "image/webp": "webp", "image/gif": "gif" }[mime] ?? "bin");

const extractImages = async (file: File, signal?: AbortSignal) => {
  const zip = await ensureDocxArchive(file);
  const entries = await getImageEntries(zip);
  const outputItems: DocumentProcessItem[] = [];
  const archiveZip = await loadJsZip();
  const archive = new archiveZip();
  const fingerprints = new Set<string>();
  for (const { name, entry } of entries) {
    throwIfAborted(signal);
    const bytes = await entry.async("uint8array");
    const extension = getExtension(name) || "bin";
    const fingerprint = `${bytes.byteLength}:${Array.from(bytes.slice(0, 8)).join(",")}:${Array.from(bytes.slice(-8)).join(",")}`;
    if (fingerprints.has(fingerprint)) continue;
    fingerprints.add(fingerprint);
    const mime = getMediaMime(bytes, extension) ?? "application/octet-stream";
    const outputExtension = getMediaMime(bytes, extension) ? mimeExtension(mime) : extension;
    const output = new File([bytesToArrayBuffer(bytes)], getDocumentOutputName(file.name, "extract-images-from-docx", outputItems.length, outputExtension), { type: mime });
    outputItems.push({ input: file, output, format: "image", inputBytes: file.size, outputBytes: output.size });
    archive.file(`images/${output.name}`, bytes);
  }
  if (!outputItems.length) throw new DocumentProcessingError("NO_READABLE_CONTENT", "No embedded images were found. Shapes, charts, icons, and other drawing objects are not extracted by this route.");
  const archiveBlob = await archive.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 9 } });
  const archiveFile = new File([archiveBlob], `${getSafeBaseName(file.name)}-images.zip`, { type: "application/zip" });
  return { items: outputItems, archive: { input: file, output: archiveFile, format: "zip", inputBytes: file.size, outputBytes: archiveFile.size } satisfies DocumentProcessItem };
};

export class BrowserDocumentEngine implements DocumentEngine {
  async validate(files: File[], options: DocumentProcessOptions): Promise<DocumentValidationResult> {
    const issues: DocumentValidationIssue[] = [];
    if (options.tool === "merge-docx" && files.length < 2) {
      return { valid: false, issues: [{ file: files[0], code: "INVALID_INPUT", message: "Choose at least two DOCX files to merge." }] };
    }
    if (!files.length) return { valid: false, issues: [] };
    for (const file of files) {
      try {
        await validateFile(file, options.tool);
      } catch (error) {
        const issue = asDocumentError(error);
        issues.push({ file, code: issue.code, message: issue.userMessage });
      }
    }
    return { valid: issues.length === 0, issues };
  }

  async inspectMetadata(file: File) {
    const zip = await ensureDocxArchive(file);
    return getMetadataFromZip(zip);
  }

  async process(files: File[], options: DocumentProcessOptions, signal?: AbortSignal): Promise<DocumentProcessResult> {
    if (!files.length) throw new DocumentProcessingError("INVALID_INPUT", "Choose a supported document to continue.");
    throwIfAborted(signal);
    if (options.tool === "merge-docx") {
      const validation = await this.validate(files, options);
      if (!validation.valid) return { items: [], failures: validation.issues.map((issue) => ({ input: issue.file, error: new DocumentProcessingError(issue.code, issue.message) })) };
      try {
        const output = await mergeDocxOutput(files, signal);
        const item = toDocumentItem(files[0], output, "docx", getDocumentOutputName(files[0].name, options.tool));
        await validateOutput(item.output, "docx");
        return { items: [item], failures: [] };
      } catch (error) {
        return { items: [], failures: [{ input: files[0], error: asDocumentError(error) }] };
      }
    }
    if (options.tool === "extract-images-from-docx") {
      try {
        const validation = await this.validate([files[0]], options);
        if (!validation.valid) return { items: [], failures: validation.issues.map((issue) => ({ input: issue.file, error: new DocumentProcessingError(issue.code, issue.message) })) };
        const result = await extractImages(files[0], signal);
        return { items: result.items, archive: result.archive, failures: [] };
      } catch (error) {
        return { items: [], failures: [{ input: files[0], error: asDocumentError(error) }] };
      }
    }
    const result: DocumentProcessResult = { items: [], failures: [] };
    for (const file of files) {
      try {
        throwIfAborted(signal);
        await validateFile(file, options.tool);
        let output: Blob;
        let metadata: DocumentProcessItem["metadata"];
        if (options.tool === "txt-to-docx") output = await createDocxOutput(file);
        else if (options.tool === "docx-to-html") output = await createDocxToHtmlOutput(file, signal);
        else if (options.tool === "docx-to-pdf" || options.tool === "pdf-to-docx") throw new DocumentProcessingError("UNSUPPORTED_FEATURE", "This conversion uses the disclosed native document fallback instead of a browser snapshot or text-only export.");
        else if (options.tool === "html-to-docx") output = await createHtmlToDocxOutput(file, signal);
        else if (options.tool === "compress-docx") output = await compressDocxOutput(file, options.compressionPreset ?? "balanced", signal);
        else if (options.tool === "docx-metadata-cleaner") {
          const cleaned = await cleanMetadata(file, signal);
          output = cleaned.blob;
          metadata = cleaned.metadata;
        } else output = await createDocxToTextOutput(file, signal);
        const format = getOutputFormat(options.tool);
        const outputFile = new File([output], getDocumentOutputName(file.name, options.tool), { type: getOutputMime(format) });
        await validateOutput(outputFile, format);
        result.items.push({ input: file, output: outputFile, format, inputBytes: file.size, outputBytes: outputFile.size, metadata });
      } catch (error) {
        const typed = asDocumentError(error);
        if (typed.code === "CANCELLED") throw typed;
        result.failures.push({ input: file, error: typed });
      }
    }
    return result;
  }

  dispose() {
    // The document libraries do not expose a disposable browser runtime. Their
    // caches contain code only; document bytes are released with each job.
  }
}

export const documentEngine = new BrowserDocumentEngine();
export const getDocumentErrorMessage = (error: unknown) => asDocumentError(error).userMessage;
