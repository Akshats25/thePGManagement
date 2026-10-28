import fs from 'node:fs/promises';
import path from 'node:path';
import {
  Presentation,
  PresentationFile,
} from 'file:///C:/Users/aksha/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';

const workspaceDir = 'D:/rental-main';
const buildDir = path.join(workspaceDir, '.ppt-build');
const draftPath = path.join(buildDir, 'panesar-living-interview-draft.pptx');

async function saveBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const navy = '#102A43';
const blue = '#1479FF';
const cyan = '#44C3F2';
const ink = '#1F2937';
const muted = '#5B677A';
const pale = '#F5F8FC';
const white = '#FFFFFF';
const green = '#17835E';

function box(slide, left, top, width, height, fill = 'none', line = 'none') {
  return slide.shapes.add({
    geometry: 'rect',
    position: { left, top, width, height },
    fill,
    line: { style: 'solid', fill: line, width: line === 'none' ? 0 : 1 },
  });
}

function text(slide, value, left, top, width, height, style = {}) {
  const shape = slide.shapes.add({
    geometry: 'textbox',
    position: { left, top, width, height },
    fill: 'none',
    line: { style: 'solid', fill: 'none', width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    fontFamily: 'Aptos',
    fontSize: 20,
    color: ink,
    ...style,
  };
  return shape;
}

function chrome(slide, title, section) {
  slide.background.fill = pale;
  box(slide, 0, 0, 1280, 14, blue);
  text(slide, section.toUpperCase(), 72, 38, 300, 26, { fontSize: 13, bold: true, color: blue });
  text(slide, title, 72, 70, 1000, 55, { fontSize: 32, bold: true, color: navy });
  text(slide, 'Panesar Living | Hotel & PG Rental Management', 72, 674, 700, 20, { fontSize: 11, color: muted });
}

function bullets(slide, items, left, top, width, fontSize = 20, color = ink) {
  const copy = items.map((item) => `• ${item}`).join('\n');
  return text(slide, copy, left, top, width, items.length * (fontSize + 16) + 10, {
    fontSize,
    color,
    breakLine: false,
    paragraphSpacing: 13,
  });
}

// 1. Cover
{
  const slide = deck.slides.add();
  slide.background.fill = navy;
  box(slide, 72, 78, 14, 455, cyan);
  text(slide, 'PANESAR LIVING', 112, 91, 350, 30, { fontSize: 16, bold: true, color: cyan });
  text(slide, 'Rental Management\nSystem', 112, 143, 760, 145, { fontSize: 48, bold: true, color: white });
  text(slide, 'A full-stack web application that manages rooms, residents, billing and reports.', 112, 323, 760, 58, { fontSize: 23, color: '#D8E7FA' });
  text(slide, 'Interview presentation', 112, 560, 280, 24, { fontSize: 16, bold: true, color: cyan });
  text(slide, 'Built with Django, React and Vite', 112, 596, 520, 26, { fontSize: 18, color: white });
  text(slide, 'ROOMS\nRESIDENTS\nREVENUE', 934, 172, 220, 190, { fontSize: 25, bold: true, color: cyan });
}

// 2. Problem and solution
{
  const slide = deck.slides.add();
  chrome(slide, 'The business problem', 'Context');
  text(slide, 'Replace manual tracking with one reliable workspace.', 72, 144, 980, 34, { fontSize: 23, color: muted });
  text(slide, 'Before', 72, 233, 300, 28, { fontSize: 20, bold: true, color: '#B42318' });
  bullets(slide, ['Track room status in separate records', 'Follow up on rent and electricity manually', 'Search resident history slowly'], 72, 281, 470, 20);
  text(slide, 'With Panesar Living', 660, 233, 400, 28, { fontSize: 20, bold: true, color: green });
  bullets(slide, ['View live room occupancy', 'Record rent and bills in one system', 'Generate dashboard and report data'], 660, 281, 500, 20);
}

// 3. User flow
{
  const slide = deck.slides.add();
  chrome(slide, 'How the system works', 'Workflow');
  text(slide, 'Staff follow one clear operational flow.', 72, 142, 800, 30, { fontSize: 22, color: muted });
  const steps = [
    ['1', 'Add rooms', 'Create buildings and room inventory'],
    ['2', 'Assign residents', 'Check capacity and record profiles'],
    ['3', 'Collect payments', 'Track rent and electricity dues'],
    ['4', 'Review reports', 'Monitor occupancy and finances'],
  ];
  for (let i = 0; i < steps.length; i += 1) {
    const x = 72 + i * 286;
    box(slide, x, 250, 226, 214, white, '#D9E3F0');
    text(slide, steps[i][0], x + 20, 270, 38, 34, { fontSize: 24, bold: true, color: blue });
    text(slide, steps[i][1], x + 20, 320, 180, 54, { fontSize: 21, bold: true, color: navy });
    text(slide, steps[i][2], x + 20, 387, 180, 52, { fontSize: 15, color: muted });
  }
}

// 4. Architecture
{
  const slide = deck.slides.add();
  chrome(slide, 'Technical architecture', 'Design');
  text(slide, 'Separate the interface from the business logic and database.', 72, 142, 950, 30, { fontSize: 22, color: muted });
  const layers = [
    ['React + Vite', 'Render fast, responsive screens'],
    ['Django APIs', 'Validate requests and apply business rules'],
    ['Django ORM + Database', 'Store rooms, residents and transactions'],
  ];
  for (let i = 0; i < layers.length; i += 1) {
    const y = 225 + i * 125;
    box(slide, 152, y, 760, 90, i === 0 ? '#EAF3FF' : white, '#C9D7EA');
    text(slide, layers[i][0], 186, y + 20, 260, 30, { fontSize: 22, bold: true, color: navy });
    text(slide, layers[i][1], 480, y + 22, 370, 30, { fontSize: 18, color: muted });
  }
  text(slide, 'Browser\n↓\nAPI\n↓\nData', 988, 235, 160, 210, { fontSize: 26, bold: true, color: blue });
}

// 5. Core capabilities
{
  const slide = deck.slides.add();
  chrome(slide, 'Core capabilities', 'Features');
  const features = [
    ['Manage inventory', 'Organize buildings, rooms and capacity'],
    ['Onboard residents', 'Assign rooms and keep resident records'],
    ['Automate billing', 'Create rent and electricity entries'],
    ['Track collections', 'Record payments and payment history'],
    ['Control expenses', 'Capture maintenance costs'],
    ['Analyze performance', 'Read dashboard, monthly and yearly reports'],
  ];
  for (let i = 0; i < features.length; i += 1) {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 72 + col * 564;
    const y = 170 + row * 145;
    text(slide, features[i][0], x, y, 450, 28, { fontSize: 22, bold: true, color: navy });
    text(slide, features[i][1], x, y + 40, 470, 42, { fontSize: 17, color: muted });
  }
}

// 6. APIs and security
{
  const slide = deck.slides.add();
  chrome(slide, 'API and security approach', 'Engineering');
  text(slide, 'Expose focused JSON APIs and protect every sensitive action.', 72, 142, 980, 30, { fontSize: 22, color: muted });
  text(slide, 'API groups', 72, 218, 280, 30, { fontSize: 21, bold: true, color: navy });
  bullets(slide, ['Rooms and availability', 'Residents and identity documents', 'Bookings, payments and electricity bills', 'Dashboard metrics and financial reports'], 72, 266, 510, 18);
  text(slide, 'Security controls', 670, 218, 350, 30, { fontSize: 21, bold: true, color: navy });
  bullets(slide, ['Use session-based authentication', 'Validate unsafe requests with CSRF tokens', 'Restrict staff-only operations', 'Serve private documents through protected routes'], 670, 266, 500, 18);
}

// 7. Engineering highlights
{
  const slide = deck.slides.add();
  chrome(slide, 'Engineering highlights', 'Quality');
  text(slide, 'I built for consistency, safety and easy maintenance.', 72, 142, 900, 30, { fontSize: 22, color: muted });
  bullets(slide, ['Enforce a maximum room capacity of two residents', 'Recalculate payment status after each record changes', 'Prevent invalid meter readings and conflicting billing data', 'Use reusable React components and a shared API client', 'Build the frontend for production with Vite'], 72, 220, 1030, 20);
}

// 8. Closing
{
  const slide = deck.slides.add();
  slide.background.fill = navy;
  text(slide, 'Project outcome', 72, 86, 420, 40, { fontSize: 19, bold: true, color: cyan });
  text(slide, 'A practical system\nfor daily rental operations', 72, 152, 900, 130, { fontSize: 44, bold: true, color: white });
  bullets(slide, ['Centralize operational data', 'Reduce manual follow-up', 'Support faster decisions with live reports'], 72, 350, 650, 22, '#D8E7FA');
  text(slide, 'Thank you', 72, 590, 300, 38, { fontSize: 28, bold: true, color: cyan });
}

for (const [index, slide] of deck.slides.items.entries()) {
  await saveBlob(path.join(buildDir, `slide-${String(index + 1).padStart(2, '0')}.png`), await deck.export({ slide, format: 'png', scale: 1 }));
}
await (await PresentationFile.exportPptx(deck)).save(draftPath);
console.log(draftPath);
