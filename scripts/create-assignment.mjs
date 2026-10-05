import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import pptxgen from "pptxgenjs";
import PDFDocument from "pdfkit";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const assets = path.join(root, "assignment-assets");
const outputPptx = path.join(root, "Zenvora_Project_Report.pptx");
const outputPdf = path.join(root, "Zenvora_Project_Report.pdf");
const colors = {
  ink: "292A24",
  green: "43584A",
  muted: "77776F",
  paper: "F6F5F0",
  orange: "C56642",
  line: "DEDDD5",
};
const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Student";
pptx.subject = "Single-page e-commerce website with inventory management";
pptx.title = "Zenvora Project Report";
pptx.company = "Student Project";
pptx.lang = "en-IN";
pptx.theme = {
  headFontFace: "Aptos Display",
  bodyFontFace: "Aptos",
  lang: "en-IN",
};
pptx.defineSlideMaster({
  title: "BASE",
  background: { color: colors.paper },
  objects: [
    {
      line: {
        x: 0.52,
        y: 7.08,
        w: 12.3,
        h: 0,
        line: { color: colors.line, width: 0.8 },
      },
    },
    {
      text: {
        text: "ZENVORA  /  PROJECT REPORT",
        options: {
          x: 0.55,
          y: 7.13,
          w: 4.5,
          h: 0.18,
          fontFace: "Aptos",
          fontSize: 8,
          color: colors.muted,
          charSpacing: 1.1,
          margin: 0,
        },
      },
    },
  ],
  slideNumber: {
    x: 12.35,
    y: 7.11,
    color: colors.muted,
    fontFace: "Aptos",
    fontSize: 8,
  },
});

function title(slide, eyebrow, heading, subtitle = "") {
  slide.addText(eyebrow.toUpperCase(), {
    x: 0.65,
    y: 0.42,
    w: 11.9,
    h: 0.2,
    fontSize: 9,
    bold: true,
    charSpacing: 1.7,
    color: colors.orange,
    margin: 0,
  });
  slide.addText(heading, {
    x: 0.65,
    y: 0.77,
    w: 12,
    h: 0.58,
    fontFace: "Aptos Display",
    fontSize: 29,
    bold: true,
    color: colors.ink,
    margin: 0,
    breakLine: false,
  });
  if (subtitle)
    slide.addText(subtitle, {
      x: 0.67,
      y: 1.4,
      w: 11.8,
      h: 0.3,
      fontSize: 11,
      color: colors.muted,
      margin: 0,
    });
}

function bulletList(slide, items, x, y, w, opts = {}) {
  const fontSize = opts.fontSize || 15;
  const step = opts.step || 0.66;
  items.forEach((item, index) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: y + index * step + 0.09,
      w: 0.09,
      h: 0.09,
      fill: { color: colors.orange },
      line: { color: colors.orange, transparency: 100 },
    });
    slide.addText(item, {
      x: x + 0.22,
      y: y + index * step,
      w: w - 0.22,
      h: step - 0.08,
      fontSize,
      color: colors.ink,
      breakLine: false,
      valign: "mid",
      margin: 0.02,
      fit: "shrink",
    });
  });
}

function addScreenshot(slide, file, x, y, w, h) {
  const screenshot = path.join(assets, file);
  if (fs.existsSync(screenshot)) {
    const header = fs.readFileSync(screenshot).subarray(0, 24);
    const imageWidth = header.readUInt32BE(16);
    const imageHeight = header.readUInt32BE(20);
    const scale = Math.min(w / imageWidth, h / imageHeight);
    const imageW = imageWidth * scale;
    const imageH = imageHeight * scale;
    slide.addImage({
      path: screenshot,
      x: x + (w - imageW) / 2,
      y: y + (h - imageH) / 2,
      w: imageW,
      h: imageH,
    });
  }
}

const slides = [];
{
  const slide = pptx.addSlide();
  slides.push(slide);
  slide.background = { color: colors.green };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 0.22,
    h: 7.5,
    fill: { color: colors.orange },
    line: { color: colors.orange },
  });
  slide.addText("Z.", {
    x: 0.9,
    y: 0.75,
    w: 1.15,
    h: 0.9,
    fontFace: "Aptos Display",
    fontSize: 43,
    bold: true,
    color: colors.paper,
    margin: 0,
  });
  slide.addText("ZENVORA", {
    x: 0.91,
    y: 2.15,
    w: 7.3,
    h: 0.7,
    fontFace: "Aptos Display",
    fontSize: 38,
    bold: true,
    charSpacing: 4,
    color: colors.paper,
    margin: 0,
  });
  slide.addText("Single-page e-commerce website\nwith inventory management", {
    x: 0.93,
    y: 3.08,
    w: 7.5,
    h: 1.0,
    fontSize: 24,
    color: "E6E8DF",
    breakLine: false,
    margin: 0,
    lineSpacingMultiple: 1.12,
  });
  slide.addText("PROJECT REPORT  /  2026", {
    x: 0.95,
    y: 5.42,
    w: 5.5,
    h: 0.3,
    fontSize: 10,
    charSpacing: 2,
    color: "D7DED1",
    margin: 0,
  });
  slide.addText(
    "Submitted by: ____________________\nClass / Roll No.: ____________________",
    {
      x: 0.95,
      y: 5.9,
      w: 5.8,
      h: 0.65,
      fontSize: 12,
      color: colors.paper,
      breakLine: false,
      margin: 0,
    },
  );
  addScreenshot(slide, "storefront.png", 8.25, 0.8, 4.25, 5.95);
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "01 / Introduction",
    "Project overview",
    "A small online shop concept with a built-in stock room for the shop owner.",
  );
  bulletList(
    slide,
    [
      "Brand: Zenvora, a curated everyday-goods store",
      "Customer side: browse, filter, sort, and add products to a bag",
      "Admin side: add products, search stock, and adjust quantities",
      "Built with HTML, CSS, and JavaScript; browser storage keeps demo data",
    ],
    0.9,
    2.1,
    6.3,
    { fontSize: 15, step: 0.78 },
  );
  slide.addShape(pptx.ShapeType.rect, {
    x: 8.0,
    y: 2.15,
    w: 4.35,
    h: 2.55,
    rectRadius: 0.05,
    fill: { color: "E8E7DF" },
    line: { color: colors.line },
  });
  slide.addText("PROJECT AIM", {
    x: 8.35,
    y: 2.48,
    w: 3.6,
    h: 0.25,
    fontSize: 9,
    bold: true,
    charSpacing: 1.6,
    color: colors.orange,
    margin: 0,
  });
  slide.addText(
    "Make the shopping experience and day-to-day stock updates feel like parts of one simple tool.",
    {
      x: 8.35,
      y: 2.92,
      w: 3.55,
      h: 1.12,
      fontSize: 19,
      bold: true,
      color: colors.ink,
      breakLine: false,
      margin: 0.02,
      valign: "mid",
      fit: "shrink",
    },
  );
  slide.addText(
    "The current demo saves data in this browser only. It is not connected to a real database or payment service.",
    {
      x: 0.92,
      y: 5.8,
      w: 11.2,
      h: 0.52,
      fontSize: 11,
      italic: true,
      color: colors.muted,
      margin: 0,
    },
  );
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "02 / Plan",
    "What I wanted to build",
    "I planned the main customer journey and the basic store-owner workflow first.",
  );
  const cards = [
    [
      "01",
      "Shop",
      "A clear landing section, product categories, and a scannable product grid.",
    ],
    [
      "02",
      "Bag",
      "Quick add, quantity controls, and a subtotal that updates with the bag.",
    ],
    [
      "03",
      "Stock room",
      "Inventory totals, stock statuses, search, and quantity adjustment.",
    ],
    [
      "04",
      "Finish",
      "Responsive layouts, local persistence, and a project hand-in checklist.",
    ],
  ];
  cards.forEach((card, index) => {
    const x = 0.75 + index * 3.12;
    slide.addShape(pptx.ShapeType.rect, {
      x,
      y: 2.3,
      w: 2.78,
      h: 2.52,
      fill: { color: index === 2 ? "E4E9E1" : "EEEDE7" },
      line: { color: colors.line, width: 0.8 },
    });
    slide.addText(card[0], {
      x: x + 0.2,
      y: 2.5,
      w: 0.55,
      h: 0.3,
      fontSize: 12,
      color: colors.orange,
      bold: true,
      margin: 0,
    });
    slide.addText(card[1], {
      x: x + 0.2,
      y: 2.98,
      w: 2.3,
      h: 0.4,
      fontSize: 18,
      bold: true,
      color: colors.ink,
      margin: 0,
    });
    slide.addText(card[2], {
      x: x + 0.2,
      y: 3.52,
      w: 2.34,
      h: 0.92,
      fontSize: 12,
      color: colors.muted,
      breakLine: false,
      margin: 0.01,
      fit: "shrink",
    });
  });
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "03 / Prompt",
    "Prompt used to shape the website",
    "Lovable-style prompt used as the starting brief for the interface.",
  );
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.78,
    y: 2.0,
    w: 11.75,
    h: 3.65,
    fill: { color: "EEEDE7" },
    line: { color: colors.line },
  });
  const prompt =
    "Build a responsive, single-page e-commerce website for a curated everyday-goods brand called Zenvora. Use a warm editorial visual style, product filters and sorting, an add-to-bag interaction, and a separate inventory management panel. The inventory panel should show product count, stock value, low-stock and out-of-stock totals, searchable products, stock adjustment controls, and an add-product form. Use accessible controls, realistic sample products, and local browser storage for this prototype. Keep the customer storefront and inventory dashboard easy to switch between.";
  slide.addText("“" + prompt + "”", {
    x: 1.12,
    y: 2.32,
    w: 11.05,
    h: 2.8,
    fontFace: "Aptos",
    fontSize: 17,
    color: colors.ink,
    italic: true,
    breakLine: false,
    margin: 0.02,
    valign: "mid",
    fit: "shrink",
    paraSpaceAfterPt: 9,
  });
  slide.addText(
    "The actual project is implemented with plain HTML, CSS, and JavaScript.",
    {
      x: 1.1,
      y: 5.98,
      w: 10.8,
      h: 0.25,
      fontSize: 10,
      color: colors.muted,
      margin: 0,
    },
  );
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "04 / Customer view",
    "The storefront",
    "The product list can be browsed by category and sorted by price or name.",
  );
  addScreenshot(slide, "storefront.png", 0.78, 1.92, 5.2, 4.78);
  bulletList(
    slide,
    [
      "Zenvora storefront and brand identity",
      "Eight sample products across three categories",
      "Quick add-to-bag interaction and quantity controls",
      "Story and newsletter sections complete the single-page layout",
    ],
    6.55,
    2.45,
    5.6,
    { fontSize: 14, step: 0.78 },
  );
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "05 / Store owner view",
    "Inventory management",
    "The Manage button opens a working stock panel alongside the customer shop.",
  );
  addScreenshot(slide, "inventory-panel.png", 0.8, 1.88, 6.05, 4.83);
  bulletList(
    slide,
    [
      "Inventory value is calculated from price × stock",
      "Low stock is fewer than 8 units; zero units is out of stock",
      "Search and stock-status filters narrow the product table",
      "Plus/minus controls update quantities; the form adds products",
    ],
    7.25,
    2.35,
    5.35,
    { fontSize: 13, step: 0.79 },
  );
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "06 / Local setup",
    "Run and test the project",
    "These are the commands used from the project folder.",
  );
  const steps = [
    ["1", "Install", "npm install"],
    ["2", "Start locally", "npm run dev"],
    ["3", "Create production build", "npm run build"],
    ["4", "Generate report files", "npm run docs"],
  ];
  steps.forEach((step, index) => {
    const y = 2.05 + index * 0.88;
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 0.92,
      y: y + 0.05,
      w: 0.35,
      h: 0.35,
      fill: { color: colors.green },
      line: { color: colors.green },
    });
    slide.addText(step[0], {
      x: 0.92,
      y: y + 0.1,
      w: 0.35,
      h: 0.17,
      align: "center",
      fontSize: 9,
      bold: true,
      color: "FFFFFF",
      margin: 0,
    });
    slide.addText(step[1], {
      x: 1.55,
      y,
      w: 2.35,
      h: 0.34,
      fontSize: 14,
      bold: true,
      color: colors.ink,
      margin: 0,
    });
    slide.addText(step[2], {
      x: 4.15,
      y,
      w: 4.1,
      h: 0.36,
      fontFace: "Aptos Mono",
      fontSize: 12,
      color: colors.green,
      margin: 0.03,
      breakLine: false,
    });
  });
  slide.addText(
    "The site uses a Vite dev server. The generated dist/ folder is the deployment output.",
    {
      x: 0.96,
      y: 6.1,
      w: 10.8,
      h: 0.32,
      fontSize: 11,
      color: colors.muted,
      margin: 0,
    },
  );
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "07 / GitHub",
    "Save the code on GitHub",
    "The project has a GitHub origin configured. Push access depends on account authentication.",
  );
  bulletList(
    slide,
    [
      "Open the project folder in VS Code and review the files",
      "Run: git add .",
      'Run: git commit -m "Build Zenvora store and inventory panel"',
      "Run: git push -u origin main",
      "Confirm the latest commit appears in the GitHub repository",
    ],
    0.95,
    2.0,
    7.15,
    { fontSize: 14, step: 0.72 },
  );
  slide.addShape(pptx.ShapeType.rect, {
    x: 8.6,
    y: 2.2,
    w: 3.55,
    h: 2.7,
    fill: { color: "E9EAE2" },
    line: { color: colors.line },
  });
  slide.addText("Repository", {
    x: 8.92,
    y: 2.55,
    w: 2.9,
    h: 0.24,
    fontSize: 10,
    charSpacing: 1,
    color: colors.muted,
    margin: 0,
  });
  slide.addText("github.com/Sujal9r/E-Commerce", {
    x: 8.92,
    y: 2.98,
    w: 2.85,
    h: 0.66,
    fontSize: 15,
    bold: true,
    color: colors.green,
    breakLine: false,
    margin: 0,
    fit: "shrink",
  });
  slide.addText(
    "A remote is configured, but no commit or push has been made from this workspace.",
    {
      x: 8.92,
      y: 3.9,
      w: 2.82,
      h: 0.68,
      fontSize: 10,
      color: colors.muted,
      breakLine: false,
      margin: 0,
    },
  );
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "08 / Vercel",
    "Deploy with Vercel",
    "Connect the GitHub repository to Vercel to get a public website link.",
  );
  bulletList(
    slide,
    [
      "Sign in to Vercel with the GitHub account",
      "Choose Add New → Project, then import Sujal9r/E-Commerce",
      "Framework preset: Vite; build command: npm run build",
      "Output directory: dist; keep the root directory as the project root",
      "Select Deploy and wait for the build to finish",
      "Copy the generated Production URL into the assignment",
    ],
    0.92,
    1.95,
    8.0,
    { fontSize: 13, step: 0.7 },
  );
  slide.addShape(pptx.ShapeType.rect, {
    x: 9.25,
    y: 2.38,
    w: 2.8,
    h: 2.35,
    fill: { color: "E6E9E2" },
    line: { color: colors.line },
  });
  slide.addText("LIVE WEBSITE URL", {
    x: 9.55,
    y: 2.72,
    w: 2.2,
    h: 0.25,
    fontSize: 9,
    bold: true,
    charSpacing: 1.1,
    color: colors.orange,
    margin: 0,
  });
  slide.addText("Pending Vercel deployment", {
    x: 9.55,
    y: 3.18,
    w: 2.15,
    h: 0.65,
    fontSize: 16,
    bold: true,
    color: colors.ink,
    breakLine: false,
    margin: 0,
    fit: "shrink",
  });
  slide.addText("Paste the real URL here after publishing.", {
    x: 9.55,
    y: 4.08,
    w: 2.2,
    h: 0.45,
    fontSize: 10,
    color: colors.muted,
    breakLine: false,
    margin: 0,
  });
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "09 / Drive and submission",
    "Share the files with your teacher",
    "Upload the finished PDF and PPTX to your own Google Drive account.",
  );
  bulletList(
    slide,
    [
      "Open drive.google.com and upload Zenvora_Project_Report.pdf and .pptx",
      "For each file, open Share → General access",
      "Choose “Anyone with the link” and set Viewer access",
      "Copy each share link and test it in a private browser window",
      "Paste the real website URL and both Drive links into the assignment description",
    ],
    0.95,
    2.02,
    7.65,
    { fontSize: 13, step: 0.77 },
  );
  slide.addShape(pptx.ShapeType.rect, {
    x: 9.0,
    y: 2.3,
    w: 3.2,
    h: 2.45,
    fill: { color: "EEEDE7" },
    line: { color: colors.line },
  });
  slide.addText("ASSIGNMENT LINKS", {
    x: 9.3,
    y: 2.67,
    w: 2.6,
    h: 0.23,
    fontSize: 9,
    bold: true,
    charSpacing: 1.2,
    color: colors.orange,
    margin: 0,
  });
  slide.addText(
    "Website: [paste Vercel URL]\nPDF: [paste Drive link]\nPPT: [paste Drive link]",
    {
      x: 9.3,
      y: 3.12,
      w: 2.5,
      h: 1.15,
      fontSize: 12,
      color: colors.ink,
      breakLine: false,
      margin: 0,
      fit: "shrink",
    },
  );
  slide.addText(
    "These links are placeholders until the files are published from the owner’s accounts.",
    {
      x: 0.98,
      y: 6.25,
      w: 11.2,
      h: 0.25,
      fontSize: 10,
      italic: true,
      color: colors.muted,
      margin: 0,
    },
  );
}
{
  const slide = pptx.addSlide("BASE");
  slides.push(slide);
  title(
    slide,
    "10 / Conclusion",
    "What I learned",
    "The prototype is ready locally; public hosting and Drive sharing are the remaining account steps.",
  );
  bulletList(
    slide,
    [
      "Connected the storefront with basic inventory updates",
      "Used browser storage to retain demo products and the shopping bag",
      "Added Git ignore rules for secrets, dependencies, and generated build files",
      "Next improvement: use a database, secure admin login, and a real checkout provider",
    ],
    1.0,
    2.25,
    10.7,
    { fontSize: 16, step: 0.85 },
  );
  slide.addText("THANK YOU", {
    x: 1.02,
    y: 6.15,
    w: 3.6,
    h: 0.35,
    fontSize: 13,
    bold: true,
    charSpacing: 2,
    color: colors.green,
    margin: 0,
  });
}

function pdfText(doc, text, x, y, width, options = {}) {
  doc
    .font(options.font || "Helvetica")
    .fontSize(options.size || 12)
    .fillColor(options.color || "#292a24");
  doc.text(text, x, y, {
    width,
    lineGap: options.lineGap || 4,
    align: options.align || "left",
  });
}

function pdfHeader(doc, section, heading, subtitle, page) {
  doc.rect(0, 0, 612, 12).fill("#43584a");
  pdfText(doc, "ZENVORA  /  PROJECT REPORT", 48, 31, 240, {
    size: 8,
    color: "#77776f",
  });
  pdfText(doc, section.toUpperCase(), 48, 66, 500, {
    size: 9,
    color: "#c56642",
  });
  pdfText(doc, heading, 48, 86, 515, { size: 26, font: "Helvetica-Bold" });
  if (subtitle)
    pdfText(doc, subtitle, 48, 122, 515, { size: 10, color: "#77776f" });
  doc
    .moveTo(48, 754)
    .lineTo(564, 754)
    .strokeColor("#deddd5")
    .lineWidth(0.7)
    .stroke();
  pdfText(doc, "ZENVORA  /  PROJECT REPORT", 48, 763, 300, {
    size: 7,
    color: "#77776f",
  });
  pdfText(doc, String(page), 535, 763, 25, {
    size: 7,
    color: "#77776f",
    align: "right",
  });
}

function pdfBullets(doc, items, y, opts = {}) {
  const x = opts.x || 58;
  const width = opts.width || 490;
  const size = opts.size || 12;
  const gap = opts.gap || 37;
  items.forEach((item, index) => {
    const itemY = y + index * gap;
    doc.circle(x + 4, itemY + 7, 3).fill("#c56642");
    pdfText(doc, item, x + 17, itemY, width - 17, { size });
  });
}

function pdfImage(doc, name, x, y, width, height) {
  const file = path.join(assets, name);
  if (!fs.existsSync(file)) return;
  const image = doc.openImage(file);
  const scale = Math.min(width / image.width, height / image.height);
  const w = image.width * scale;
  const h = image.height * scale;
  doc.image(file, x + (width - w) / 2, y + (height - h) / 2, {
    width: w,
    height: h,
  });
}

async function buildPptx() {
  await pptx.writeFile({ fileName: outputPptx });
}

async function buildPdf() {
  const doc = new PDFDocument({
    size: "A4",
    margins: { top: 48, bottom: 48, left: 48, right: 48 },
    autoFirstPage: false,
    info: {
      Title: "Zenvora Project Report",
      Author: "Student",
      Subject: "E-commerce website and inventory management project",
    },
  });
  const stream = fs.createWriteStream(outputPdf);
  doc.pipe(stream);

  doc.addPage({ size: "A4", margins: 0 });
  doc.rect(0, 0, 612, 792).fill("#43584a");
  doc.rect(0, 0, 14, 792).fill("#c56642");
  pdfText(doc, "Z.", 53, 56, 80, {
    size: 38,
    font: "Helvetica-Bold",
    color: "#f6f5f0",
  });
  pdfText(doc, "ZENVORA", 53, 164, 480, {
    size: 34,
    font: "Helvetica-Bold",
    color: "#f6f5f0",
  });
  pdfText(
    doc,
    "Single-page e-commerce website\nwith inventory management",
    53,
    222,
    470,
    { size: 22, color: "#e6e8df", lineGap: 8 },
  );
  pdfText(doc, "PROJECT REPORT  /  2026", 53, 326, 390, {
    size: 9,
    color: "#d7ded1",
  });
  pdfText(
    doc,
    "Submitted by: ______________________________\nClass / Roll No.: ______________________________",
    53,
    365,
    400,
    { size: 12, color: "#f6f5f0", lineGap: 10 },
  );
  pdfImage(doc, "storefront.png", 345, 480, 215, 268);

  doc.addPage();
  pdfHeader(
    doc,
    "01 / Introduction",
    "Project overview",
    "A small online shop concept with a built-in stock room for the shop owner.",
    2,
  );
  pdfBullets(
    doc,
    [
      "Brand: Zenvora, a curated everyday-goods store",
      "Customer side: browse, filter, sort, and add products to a bag",
      "Admin side: add products, search stock, and adjust quantities",
      "Built with HTML, CSS, and JavaScript; browser storage keeps demo data",
    ],
    187,
    { gap: 48, size: 12 },
  );
  doc.roundedRect(58, 408, 496, 106, 5).fill("#e8e7df");
  pdfText(doc, "PROJECT AIM", 76, 427, 150, { size: 9, color: "#c56642" });
  pdfText(
    doc,
    "Make shopping and day-to-day stock updates feel like parts of one simple tool.",
    76,
    452,
    455,
    { size: 16, font: "Helvetica-Bold" },
  );
  pdfText(
    doc,
    "Note: this demo saves data in this browser only. It is not connected to a real database or payment service.",
    58,
    550,
    490,
    { size: 10, color: "#77776f" },
  );

  doc.addPage();
  pdfHeader(
    doc,
    "02 / Plan",
    "What I wanted to build",
    "I planned the main customer journey and the basic store-owner workflow first.",
    3,
  );
  const pdfCards = [
    [
      "01",
      "Shop",
      "Landing section, product categories, and a scannable product grid.",
    ],
    ["02", "Bag", "Quick add, quantity controls, and an updating subtotal."],
    [
      "03",
      "Stock room",
      "Inventory totals, stock statuses, search, and quantity adjustment.",
    ],
    [
      "04",
      "Finish",
      "Responsive layouts, local persistence, and a hand-in checklist.",
    ],
  ];
  pdfCards.forEach((card, index) => {
    const x = 48 + index * 132;
    doc
      .rect(x, 220, 120, 205)
      .fill(index === 2 ? "#e4e9e1" : "#eeede7")
      .stroke("#deddd5");
    pdfText(doc, card[0], x + 12, 237, 40, { size: 10, color: "#c56642" });
    pdfText(doc, card[1], x + 12, 263, 98, {
      size: 15,
      font: "Helvetica-Bold",
    });
    pdfText(doc, card[2], x + 12, 299, 97, {
      size: 10,
      color: "#77776f",
      lineGap: 4,
    });
  });

  doc.addPage();
  pdfHeader(
    doc,
    "03 / Prompt",
    "Prompt used to shape the website",
    "Lovable-style prompt used as the starting brief for the interface.",
    4,
  );
  const prompt =
    "Build a responsive, single-page e-commerce website for a curated everyday-goods brand called Zenvora. Use a warm editorial visual style, product filters and sorting, an add-to-bag interaction, and a separate inventory management panel. The inventory panel should show product count, stock value, low-stock and out-of-stock totals, searchable products, stock adjustment controls, and an add-product form. Use accessible controls, realistic sample products, and local browser storage for this prototype. Keep the customer storefront and inventory dashboard easy to switch between.";
  doc.roundedRect(58, 190, 496, 340, 5).fill("#eeede7").stroke("#deddd5");
  pdfText(doc, "“" + prompt + "”", 82, 218, 448, { size: 15, lineGap: 8 });
  pdfText(
    doc,
    "The actual project is implemented with plain HTML, CSS, and JavaScript.",
    60,
    552,
    490,
    { size: 10, color: "#77776f" },
  );

  doc.addPage();
  pdfHeader(
    doc,
    "04 / Customer view",
    "The storefront",
    "The product list can be browsed by category and sorted by price or name.",
    5,
  );
  pdfImage(doc, "storefront.png", 54, 160, 240, 542);
  pdfBullets(
    doc,
    [
      "Zenvora storefront and brand identity",
      "Eight sample products across three categories",
      "Quick add-to-bag interaction and quantity controls",
      "Story and newsletter sections complete the single-page layout",
    ],
    208,
    { x: 318, width: 240, gap: 60, size: 11 },
  );

  doc.addPage();
  pdfHeader(
    doc,
    "05 / Store owner view",
    "Inventory management",
    "The Manage button opens a working stock panel alongside the customer shop.",
    6,
  );
  pdfImage(doc, "inventory-panel.png", 52, 160, 280, 535);
  pdfBullets(
    doc,
    [
      "Inventory value is calculated from price × stock",
      "Low stock is fewer than 8 units; zero units is out of stock",
      "Search and stock-status filters narrow the table",
      "Plus/minus controls update quantities; the form adds products",
    ],
    212,
    { x: 347, width: 211, gap: 65, size: 10 },
  );

  doc.addPage();
  pdfHeader(
    doc,
    "06 / Local setup",
    "Run and test the project",
    "These are the commands used from the project folder.",
    7,
  );
  const commands = [
    ["1", "Install", "npm install"],
    ["2", "Start locally", "npm run dev"],
    ["3", "Create production build", "npm run build"],
    ["4", "Generate report files", "npm run docs"],
  ];
  commands.forEach((step, index) => {
    const y = 203 + index * 70;
    doc.circle(70, y + 8, 12).fill("#43584a");
    pdfText(doc, step[0], 65, y + 2, 10, {
      size: 9,
      color: "#ffffff",
      align: "center",
    });
    pdfText(doc, step[1], 100, y, 170, { size: 13, font: "Helvetica-Bold" });
    pdfText(doc, step[2], 290, y, 225, { size: 12, color: "#43584a" });
  });
  pdfText(
    doc,
    "The site uses a Vite development server. The generated dist/ folder is the deployment output.",
    62,
    534,
    485,
    { size: 11, color: "#77776f" },
  );

  doc.addPage();
  pdfHeader(
    doc,
    "07 / GitHub",
    "Save the code on GitHub",
    "The project has a GitHub origin configured. Push access depends on account authentication.",
    8,
  );
  pdfBullets(
    doc,
    [
      "Open the project folder and review the files",
      "Run: git add .",
      'Run: git commit -m "Build Zenvora store and inventory panel"',
      "Run: git push -u origin main",
      "Confirm the latest commit appears in GitHub",
    ],
    190,
    { gap: 49, size: 11 },
  );
  doc.roundedRect(70, 465, 475, 105, 5).fill("#e9eae2");
  pdfText(doc, "REPOSITORY", 90, 484, 170, { size: 9, color: "#c56642" });
  pdfText(doc, "github.com/Sujal9r/E-Commerce", 90, 512, 430, {
    size: 16,
    font: "Helvetica-Bold",
    color: "#43584a",
  });
  pdfText(
    doc,
    "A remote is configured, but no commit or push has been made from this workspace.",
    90,
    545,
    430,
    { size: 9, color: "#77776f" },
  );

  doc.addPage();
  pdfHeader(
    doc,
    "08 / Vercel",
    "Deploy with Vercel",
    "Connect the GitHub repository to Vercel to get a public website link.",
    9,
  );
  pdfBullets(
    doc,
    [
      "Sign in to Vercel with the GitHub account",
      "Add New → Project, then import Sujal9r/E-Commerce",
      "Framework preset: Vite; build command: npm run build",
      "Output directory: dist; root directory: project root",
      "Select Deploy and wait for the build to finish",
      "Copy the generated Production URL into the assignment",
    ],
    184,
    { gap: 48, size: 11 },
  );
  doc.roundedRect(74, 500, 465, 96, 5).fill("#e6e9e2");
  pdfText(doc, "LIVE WEBSITE URL", 95, 520, 240, { size: 9, color: "#c56642" });
  pdfText(
    doc,
    "Pending Vercel deployment: paste the real URL here.",
    95,
    548,
    410,
    { size: 12, font: "Helvetica-Bold" },
  );

  doc.addPage();
  pdfHeader(
    doc,
    "09 / Drive and submission",
    "Share the files with your teacher",
    "Upload the finished PDF and PPTX to your own Google Drive account.",
    10,
  );
  pdfBullets(
    doc,
    [
      "Upload Zenvora_Project_Report.pdf and .pptx to Drive",
      "For each file, open Share → General access",
      "Choose “Anyone with the link” and set Viewer access",
      "Copy each link and test it in a private browser window",
      "Paste the real website URL and both Drive links in the assignment",
    ],
    190,
    { gap: 50, size: 11 },
  );
  doc.roundedRect(76, 472, 460, 112, 5).fill("#eeede7");
  pdfText(doc, "ASSIGNMENT DESCRIPTION", 96, 491, 300, {
    size: 9,
    color: "#c56642",
  });
  pdfText(
    doc,
    "Website: [paste Vercel URL]\nPDF: [paste Drive link]\nPPT: [paste Drive link]",
    96,
    518,
    400,
    { size: 11, lineGap: 5 },
  );
  pdfText(
    doc,
    "These links remain placeholders until published from the owner’s accounts.",
    62,
    612,
    480,
    { size: 9, color: "#77776f" },
  );

  doc.addPage();
  pdfHeader(
    doc,
    "10 / Conclusion",
    "What I learned",
    "The prototype is ready locally; public hosting and Drive sharing are the remaining account steps.",
    11,
  );
  pdfBullets(
    doc,
    [
      "Connected the storefront with basic inventory updates",
      "Used browser storage to retain demo products and the shopping bag",
      "Added Git ignore rules for secrets, dependencies, and build files",
      "Next: database, secure admin login, and real checkout provider",
    ],
    225,
    { gap: 54, size: 13 },
  );
  pdfText(doc, "THANK YOU", 62, 520, 200, {
    size: 12,
    font: "Helvetica-Bold",
    color: "#43584a",
  });

  doc.end();
  await new Promise((resolve, reject) => {
    stream.on("finish", resolve);
    stream.on("error", reject);
  });
}

await Promise.all([buildPptx(), buildPdf()]);
console.log(
  `Created ${path.basename(outputPptx)} and ${path.basename(outputPdf)} with ${slides.length} slides/pages.`,
);
