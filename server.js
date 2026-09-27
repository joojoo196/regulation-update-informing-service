const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const { pathToFileURL } = require('node:url');
const nodemailer = require('nodemailer');
const PptxGenJS = require('pptxgenjs');
const { Redis } = require('@upstash/redis');

const localEnvPath = path.join(__dirname, '.env');
if (fs.existsSync(localEnvPath)) {
  for (const line of fs.readFileSync(localEnvPath, 'utf8').split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (match && !process.env[match[1]]) process.env[match[1]] = match[2].replace(/^['"]|['"]$/g, '');
  }
}

const PORT = Number(process.env.PORT || 4173);
const API_KEY = process.env.LAW_API_OC || '';
const AGENCY = '기후에너지환경부';
const LAW_API = 'https://www.law.go.kr/DRF/lawSearch.do';
const LAW_SERVICE_API = 'https://www.law.go.kr/DRF/lawService.do';
const MCEE_RECENT_LAWS_URL = 'https://www.mcee.go.kr/home/web/index.do?menuId=69';
const RECENT_PROMULGATED_LAWS_URL = 'https://www.law.go.kr/LSW/nwRvsLsPop.do?chrIdx=7&cptOfi=1482000&sortIdx=0';
const UPCOMING_LAWS_URL = 'https://www.law.go.kr/LSW/efLsPop.do?chrIdx=7&cptOfi=1482000&sortIdx=0';
const LEGISLATION_NOTICE_URL = 'https://mcee.go.kr/home/web/lawMaking/list.do';
const ADMIN_NOTICE_URL = 'https://mcee.go.kr/home/web/board/list.do?boardMasterId=827&menuId=10557&maxPageItems=100&pagerOffset=0';
const MAIL_FROM_NAME = process.env.MAIL_FROM_NAME || '환경법령 알림서비스';
const MAIL_USER = process.env.MAIL_USER || '';
const MAIL_APP_PASSWORD = process.env.MAIL_APP_PASSWORD || '';
const RECIPIENTS_FILE = path.join(__dirname, 'recipients.json');
const AUTH_FILE = path.join(__dirname, 'auth.json');
// 로컬 개발 환경에만 있는 PPT 후처리 도구. 존재할 때만 선택적으로 사용하고, 없으면(Vercel 등) 건너뛴다.
const ARTIFACT_TOOL_MODULE = 'C:\\Users\\jojow\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules\\@oai\\artifact-tool\\dist\\artifact_tool.mjs';
const LAW_REQUEST_TIMEOUT_MS = 30000;

// Vercel 등 서버리스 환경은 배포 폴더가 읽기 전용이고 함수 인스턴스마다 메모리도 따로 놀기 때문에,
// 로컬 파일/메모리로는 비밀번호·수신자·세션이 지속되지 않는다. Upstash Redis(Vercel의 KV 마켓플레이스
// 연동 시 자동 주입되는 KV_REST_API_* 환경변수, 또는 직접 연결한 UPSTASH_REDIS_REST_* 환경변수)가
// 설정되어 있으면 그쪽을 쓰고, 없으면(로컬 개발) 기존 파일/메모리 방식으로 동작한다.
let redisClient;
function getRedis() {
  if (redisClient === undefined) {
    const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
    const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
    redisClient = url && token ? new Redis({ url, token }) : null;
  }
  return redisClient;
}

// law.go.kr / mcee.go.kr는 동시 요청이 몰리면 응답이 급격히 느려지므로,
// 전체 요청을 소수만 동시에 흘려보내 타임아웃 발생을 줄인다.
function createLimiter(limit) {
  let active = 0;
  const queue = [];
  const runNext = () => {
    if (active >= limit || !queue.length) return;
    active += 1;
    const { fn, resolve, reject } = queue.shift();
    fn().then(resolve, reject).finally(() => { active -= 1; runNext(); });
  };
  return (fn) => new Promise((resolve, reject) => { queue.push({ fn, resolve, reject }); runNext(); });
}

const limitLawRequest = createLimiter(4);

const PPT = {
  navy: '1F2E70',
  navySoft: 'E9EDFA',
  ink: '263045',
  muted: '5F6B7C',
  line: 'DCE2EB',
  card: 'F3F5F9',
  green: '087347',
  blue: '2167C8',
  font: '맑은 고딕'
};

function pptText(value, max = 70) {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  return text.length > max ? `${text.slice(0, max - 1)}…` : text;
}

function pptDate(value) {
  return String(value || '').replace(/\./g, '.').trim() || '일정 미정';
}

// 개정이유·주요내용 전문을 자르지 않고, 한 슬라이드에 다 안 들어가면 문장 단위로 나눠 이어지는 슬라이드에 담는다.
function chunkSummaryText(text, chunkSize = 950) {
  const trimmed = String(text || '').trim();
  if (!trimmed) return [];
  if (trimmed.length <= chunkSize) return [trimmed];
  const sentences = trimmed.match(/[^.]+\.(?=\s|$)|[^.]+$/g) || [trimmed];
  const chunks = [];
  let current = '';
  for (const sentence of sentences) {
    if (current && current.length + sentence.length > chunkSize) {
      chunks.push(current.trim());
      current = sentence;
    } else {
      current += sentence;
    }
  }
  if (current.trim()) chunks.push(current.trim());
  return chunks;
}

function addPptTitle(slide, eyebrow, title, page) {
  slide.addText(eyebrow, { x: 0.9, y: 0.55, w: 7.7, h: 0.28, fontFace: PPT.font, fontSize: 10, bold: true, color: PPT.muted, breakLine: false, margin: 0 });
  slide.addText(title, { x: 0.9, y: 0.88, w: 10.2, h: 0.62, fontFace: PPT.font, fontSize: 29, bold: true, color: PPT.navy, margin: 0, fit: 'shrink' });
  slide.addText(String(page).padStart(2, '0'), { x: 12.42, y: 7.16, w: 0.35, h: 0.18, fontFace: PPT.font, fontSize: 8, color: '9AA6B7', align: 'right', margin: 0 });
}

function addPptFooter(slide, from, to) {
  slide.addShape('line', { x: 0.9, y: 6.98, w: 11.45, h: 0, line: { color: PPT.line, width: 0.7 } });
  slide.addText(`수집 기간  ${from} ~ ${to}    |    출처  기후에너지환경부 · 국가법령정보센터`, { x: 0.9, y: 7.08, w: 8.8, h: 0.2, fontFace: PPT.font, fontSize: 7.5, color: '7B8491', margin: 0 });
}

async function buildPptSummary(items, from, to) {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = '법규 개정 알림 서비스';
  pptx.subject = '기후에너지환경부 최근 법규 개정 요약';
  pptx.title = `법규 개정 요약 (${from}~${to})`;
  pptx.company = '법규 개정 알림 서비스';
  pptx.lang = 'ko-KR';
  pptx.theme = { headFontFace: PPT.font, bodyFontFace: PPT.font, lang: 'ko-KR' };
  const groups = [
    ['legislationNotice', '입법예고'], ['revisedLaw', '개정법령'],
    ['administrativeNotice', '행정예고'], ['revisedNotice', '개정 행정규칙']
  ];
  const grouped = Object.fromEntries(groups.map(([key]) => [key, items.filter((item) => item.group === key)]));
  const priority = [...grouped.revisedLaw, ...grouped.revisedNotice, ...grouped.legislationNotice, ...grouped.administrativeNotice];

  // 1. 표지
  let slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  slide.addShape(pptx.ShapeType.ellipse, { x: 9.1, y: 3.8, w: 5.9, h: 5.9, fill: { color: PPT.navySoft, transparency: 12 }, line: { color: PPT.navySoft, transparency: 100 } });
  slide.addShape(pptx.ShapeType.ellipse, { x: -0.8, y: -0.8, w: 2.2, h: 2.2, fill: { color: PPT.navy, transparency: 5 }, line: { color: PPT.navy, transparency: 100 } });
  slide.addText('법규 개정 요약 보고', { x: 0.9, y: 1.9, w: 6.7, h: 0.28, fontFace: PPT.font, fontSize: 11, bold: true, color: PPT.muted, margin: 0 });
  slide.addText('최근 제개정사항', { x: 0.9, y: 2.32, w: 9.2, h: 0.78, fontFace: PPT.font, fontSize: 38, bold: true, color: PPT.navy, margin: 0 });
  slide.addText('기후에너지환경부 소관 법령·행정규칙 요약', { x: 0.9, y: 3.45, w: 9.6, h: 0.36, fontFace: PPT.font, fontSize: 17, color: PPT.ink, margin: 0 });
  slide.addShape(pptx.ShapeType.ellipse, { x: 0.9, y: 6.38, w: 0.12, h: 0.12, fill: { color: PPT.navy }, line: { color: PPT.navy } });
  slide.addText(`수집기간  ${from} ~ ${to}`, { x: 1.14, y: 6.31, w: 3.4, h: 0.2, fontFace: PPT.font, fontSize: 9, color: PPT.muted, margin: 0 });
  slide.addShape(pptx.ShapeType.ellipse, { x: 5.12, y: 6.38, w: 0.12, h: 0.12, fill: { color: PPT.navy }, line: { color: PPT.navy } });
  slide.addText(`총 ${items.length}건`, { x: 5.36, y: 6.31, w: 1.5, h: 0.2, fontFace: PPT.font, fontSize: 9, color: PPT.muted, margin: 0 });

  // 2. 현황
  slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addPptTitle(slide, 'OVERVIEW', '수집 현황', 2);
  const cards = groups.map(([key, label]) => ({ label, count: grouped[key].length }));
  cards.forEach((card, index) => {
    const x = 0.9 + index * 3.03;
    slide.addShape(pptx.ShapeType.roundRect, { x, y: 2.08, w: 2.72, h: 1.55, rectRadius: 0.08, fill: { color: PPT.card }, line: { color: PPT.card } });
    slide.addText(card.label, { x: x + 0.28, y: 2.4, w: 2.15, h: 0.26, fontFace: PPT.font, fontSize: 12, bold: true, color: PPT.ink, margin: 0, align: 'center' });
    slide.addText(`${card.count}건`, { x: x + 0.28, y: 2.82, w: 2.15, h: 0.42, fontFace: PPT.font, fontSize: 24, bold: true, color: index < 2 ? PPT.navy : PPT.blue, margin: 0, align: 'center' });
  });
  slide.addText('주요 항목', { x: 0.9, y: 4.28, w: 2.0, h: 0.3, fontFace: PPT.font, fontSize: 16, bold: true, color: PPT.navy, margin: 0 });
  const highlights = priority.slice(0, 5);
  if (!highlights.length) {
    slide.addText('해당 기간에 수집된 변경사항이 없습니다.', { x: 0.9, y: 4.86, w: 8.3, h: 0.35, fontFace: PPT.font, fontSize: 15, color: PPT.muted, margin: 0 });
  } else {
    highlights.forEach((item, index) => {
      const y = 4.78 + index * 0.38;
      slide.addShape(pptx.ShapeType.ellipse, { x: 0.96, y: y + 0.08, w: 0.08, h: 0.08, fill: { color: PPT.navy }, line: { color: PPT.navy } });
      slide.addText(`${pptText(item.title, 46)}  (${item.status})`, { x: 1.2, y, w: 9.9, h: 0.25, fontFace: PPT.font, fontSize: 11, color: PPT.ink, margin: 0, fit: 'shrink' });
    });
  }
  addPptFooter(slide, from, to);

  // 3~N. 항목마다 상세 슬라이드 (제목·상태·날짜 + 개정이유·주요내용 전문 + 소관부서 + 원문/비교 링크)
  // 개정이유·주요내용 전문이 한 슬라이드에 다 안 들어가면 자르지 않고 "(계속)" 슬라이드로 이어서 담는다.
  const groupLabel = { revisedLaw: '개정법령', revisedNotice: '개정 행정규칙', legislationNotice: '입법예고', administrativeNotice: '행정예고' };
  const addDetailSlide = (item, page, itemNumber, chunk, chunkIndex, chunkTotal) => {
    const detailSlide = pptx.addSlide();
    detailSlide.background = { color: 'FFFFFF' };
    const heading = `주요 변경 항목 ${itemNumber} · ${groupLabel[item.group] || item.category}${chunkTotal > 1 ? ` (${chunkIndex + 1}/${chunkTotal})` : ''}`;
    addPptTitle(detailSlide, 'AMENDMENT DETAIL', heading, page);

    detailSlide.addShape(pptx.ShapeType.roundRect, { x: 0.9, y: 1.86, w: 1.7, h: 0.32, rectRadius: 0.16, fill: { color: PPT.navySoft }, line: { color: PPT.navySoft } });
    detailSlide.addText(item.status || item.kind || '', { x: 0.9, y: 1.86, w: 1.7, h: 0.32, fontFace: PPT.font, fontSize: 10.5, bold: true, color: PPT.navy, align: 'center', valign: 'middle', margin: 0, fit: 'shrink' });
    detailSlide.addText(pptText(item.title, 60), { x: 2.75, y: 1.82, w: 9.65, h: 0.4, fontFace: PPT.font, fontSize: 17, bold: true, color: PPT.ink, margin: 0, valign: 'middle', fit: 'shrink' });

    // 메타 정보·키워드·링크는 첫 슬라이드에만 담아, 이어지는 슬라이드는 본문에 더 많은 공간을 준다.
    const metaParts = chunkIndex ? [] : [
      item.agency, item.department ? `담당 ${item.department}` : '',
      `공포·발령·예고일 ${pptDate(item.changedAt)}`,
      item.effectiveAt ? `시행 ${pptDate(item.effectiveAt)}` : '',
      item.noticeEndAt ? `예고종료 ${pptDate(item.noticeEndAt)}` : '',
      item.noticeNumber ? `문서번호 ${item.noticeNumber}` : ''
    ].filter(Boolean);
    if (metaParts.length) {
      detailSlide.addText(metaParts.join('   ·   '), { x: 0.9, y: 2.26, w: 11.5, h: 0.3, fontFace: PPT.font, fontSize: 10, color: PPT.muted, margin: 0, fit: 'shrink' });
    }

    const cardTop = metaParts.length ? 2.62 : 2.2;
    const cardBottom = 6.3;
    detailSlide.addShape(pptx.ShapeType.roundRect, { x: 0.9, y: cardTop, w: 11.5, h: cardBottom - cardTop, rectRadius: 0.08, fill: { color: PPT.card }, line: { color: PPT.card } });
    detailSlide.addText(chunkIndex ? '개정 이유 및 주요 내용 (이어서)' : '개정 이유 및 주요 내용', { x: 1.18, y: cardTop + 0.16, w: 6, h: 0.26, fontFace: PPT.font, fontSize: 12, bold: true, color: PPT.navy, margin: 0 });
    detailSlide.addText(chunk || '원문에서 상세 요약을 추출하지 못했습니다. 아래 원문 링크에서 직접 확인해 주세요.', {
      x: 1.18, y: cardTop + 0.5, w: 11.0, h: cardBottom - cardTop - 0.66, fontFace: PPT.font, fontSize: 10.5, color: PPT.ink, margin: 0, valign: 'top', fit: 'shrink', lineSpacingMultiple: 1.28
    });

    if (!chunkIndex) {
      const keywords = (item.keywords || []).slice(0, 4);
      keywords.forEach((keyword, keywordIndex) => {
        const x = 0.9 + keywordIndex * 1.95;
        detailSlide.addShape(pptx.ShapeType.roundRect, { x, y: 6.42, w: 1.8, h: 0.32, rectRadius: 0.16, fill: { color: PPT.navySoft }, line: { color: PPT.navySoft } });
        detailSlide.addText(pptText(keyword, 11), { x, y: 6.42, w: 1.8, h: 0.32, fontFace: PPT.font, fontSize: 9, bold: true, color: PPT.navy, align: 'center', valign: 'middle', margin: 0, fit: 'shrink' });
      });

      const linkParts = [`원문 ${item.url}`, item.comparisonUrl ? `신구법 비교 ${item.comparisonUrl}` : ''].filter(Boolean);
      detailSlide.addText(linkParts.join('   ·   '), { x: 0.9, y: 6.8, w: 11.5, h: 0.18, fontFace: PPT.font, fontSize: 7.5, color: PPT.muted, margin: 0, fit: 'shrink' });
    }

    addPptFooter(detailSlide, from, to);
  };
  let detailSlideCount;
  if (!priority.length) {
    const emptySlide = pptx.addSlide();
    emptySlide.background = { color: 'FFFFFF' };
    addPptTitle(emptySlide, 'AMENDMENT DETAIL', '주요 변경 항목', 3);
    emptySlide.addText('해당 기간에 수집된 변경사항이 없습니다.', { x: 0.9, y: 3.4, w: 8.3, h: 0.35, fontFace: PPT.font, fontSize: 15, color: PPT.muted, margin: 0 });
    addPptFooter(emptySlide, from, to);
    detailSlideCount = 1;
  } else {
    const detailSlides = priority.flatMap((item, index) => {
      const chunks = chunkSummaryText(item.summary);
      const effectiveChunks = chunks.length ? chunks : [''];
      return effectiveChunks.map((chunk, chunkIndex) => ({ item, itemNumber: index + 1, chunk, chunkIndex, chunkTotal: effectiveChunks.length }));
    });
    detailSlides.forEach((spec, slideIndex) => addDetailSlide(spec.item, 3 + slideIndex, spec.itemNumber, spec.chunk, spec.chunkIndex, spec.chunkTotal));
    detailSlideCount = detailSlides.length;
  }

  // 모든 항목의 시행·예고 일정과 확인 경로. 14건마다 다음 페이지를 만듭니다.
  const scheduled = [...priority].sort((a, b) => String(a.effectiveAt || a.noticeEndAt || a.changedAt).localeCompare(String(b.effectiveAt || b.noticeEndAt || b.changedAt)));
  const scheduleChunks = scheduled.length ? Array.from({ length: Math.ceil(scheduled.length / 14) }, (_, index) => scheduled.slice(index * 14, index * 14 + 14)) : [[]];
  const addScheduleSlide = (itemsForSlide, page, chunkIndex) => {
    const scheduleSlide = pptx.addSlide();
    scheduleSlide.background = { color: 'FFFFFF' };
    addPptTitle(scheduleSlide, 'SCHEDULE & SOURCES', chunkIndex ? `시행·예고 일정과 확인 경로 ${chunkIndex + 1}` : '시행·예고 일정과 확인 경로', page);
    scheduleSlide.addShape(pptx.ShapeType.roundRect, { x: 0.9, y: 1.95, w: 8.25, h: 4.75, rectRadius: 0.08, fill: { color: PPT.card }, line: { color: PPT.card } });
    const start = chunkIndex * 14 + 1;
    const end = Math.min(start + itemsForSlide.length - 1, scheduled.length);
    scheduleSlide.addText(`일정 목록  ·  전체 ${scheduled.length}건 중 ${start}~${end}건`, { x: 1.25, y: 2.28, w: 4.3, h: 0.28, fontFace: PPT.font, fontSize: 15, bold: true, color: PPT.navy, margin: 0 });
    if (!itemsForSlide.length) {
      scheduleSlide.addText('표시할 항목이 없습니다.', { x: 1.25, y: 3.0, w: 5.8, h: 0.3, fontFace: PPT.font, fontSize: 12, color: PPT.muted, margin: 0 });
    } else {
      itemsForSlide.forEach((item, index) => {
        const column = index < 7 ? 0 : 1;
        const row = index % 7;
        const x = column ? 5.2 : 1.25;
        const y = 2.88 + row * 0.48;
        const date = item.effectiveAt ? `시행 ${pptDate(item.effectiveAt)}` : item.noticeEndAt ? `예고 종료 ${pptDate(item.noticeEndAt)}` : `공포·발령 ${pptDate(item.changedAt)}`;
        scheduleSlide.addShape(pptx.ShapeType.ellipse, { x, y: y + 0.05, w: 0.13, h: 0.13, fill: { color: index === 0 ? PPT.navy : PPT.blue }, line: { color: index === 0 ? PPT.navy : PPT.blue } });
        scheduleSlide.addText(date, { x: x + 0.28, y, w: 1.35, h: 0.18, fontFace: PPT.font, fontSize: 8, bold: true, color: PPT.navy, margin: 0, fit: 'shrink' });
        scheduleSlide.addText(pptText(item.title, 25), { x: x + 1.68, y, w: 2.08, h: 0.2, fontFace: PPT.font, fontSize: 8.5, color: PPT.ink, margin: 0, fit: 'shrink' });
      });
    }
    scheduleSlide.addShape(pptx.ShapeType.roundRect, { x: 9.48, y: 1.95, w: 2.92, h: 4.75, rectRadius: 0.08, fill: { color: 'FFFFFF' }, line: { color: PPT.line, width: 1 } });
    scheduleSlide.addText('원문 확인', { x: 9.83, y: 2.28, w: 2.1, h: 0.3, fontFace: PPT.font, fontSize: 15, bold: true, color: PPT.navy, margin: 0 });
    scheduleSlide.addText('각 항목의 원문 링크와 신구법 비교 링크는 웹페이지의 최근 제개정사항 목록에서 확인할 수 있습니다.', { x: 9.83, y: 2.9, w: 2.2, h: 1.1, fontFace: PPT.font, fontSize: 9.5, color: PPT.ink, margin: 0.02, fit: 'shrink' });
    scheduleSlide.addText('수집 기준', { x: 9.83, y: 4.48, w: 1.8, h: 0.2, fontFace: PPT.font, fontSize: 10, bold: true, color: PPT.navy, margin: 0 });
    scheduleSlide.addText('공포·발령일 최근 7일, 행정예고는 예고 시작일 최근 7일 또는 미래 항목', { x: 9.83, y: 4.85, w: 2.2, h: 0.8, fontFace: PPT.font, fontSize: 8.5, color: PPT.muted, margin: 0.02, fit: 'shrink' });
    addPptFooter(scheduleSlide, from, to);
  };
  scheduleChunks.forEach((chunk, index) => addScheduleSlide(chunk, 3 + detailSlideCount + index, index));

  const generated = await pptx.write('nodebuffer');
  return normalizePptBuffer(generated);
}

async function normalizePptBuffer(buffer) {
  if (!fs.existsSync(ARTIFACT_TOOL_MODULE)) return buffer;
  const id = `law-summary-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  const sourcePath = path.join(os.tmpdir(), `${id}.pptx`);
  const outputPath = path.join(os.tmpdir(), `${id}-normalized.pptx`);
  try {
    await fs.promises.writeFile(sourcePath, buffer);
    const { FileBlob, PresentationFile } = await import(pathToFileURL(ARTIFACT_TOOL_MODULE).href);
    const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
    await (await PresentationFile.exportPptx(presentation)).save(outputPath);
    return await fs.promises.readFile(outputPath);
  } catch {
    return buffer;
  } finally {
    await Promise.all([fs.promises.unlink(sourcePath).catch(() => {}), fs.promises.unlink(outputPath).catch(() => {})]);
  }
}

function normalizeRecipients(values) {
  return [...new Set(asArray(values).map((value) => String(value).trim().toLowerCase()).filter((email) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)))];
}

async function readRecipients() {
  const redis = getRedis();
  if (redis) {
    try { return normalizeRecipients((await redis.get('recipients:list')) || []); }
    catch { return []; }
  }
  try { return normalizeRecipients(JSON.parse(fs.readFileSync(RECIPIENTS_FILE, 'utf8'))); }
  catch (error) { return error.code === 'ENOENT' ? [] : []; }
}

async function saveRecipients(recipients) {
  const saved = normalizeRecipients(recipients);
  const redis = getRedis();
  if (redis) { await redis.set('recipients:list', saved); return saved; }
  const temporaryFile = `${RECIPIENTS_FILE}.tmp`;
  fs.writeFileSync(temporaryFile, `${JSON.stringify(saved, null, 2)}\n`, 'utf8');
  fs.renameSync(temporaryFile, RECIPIENTS_FILE);
  return saved;
}

// --- 접속 비밀번호 ---
// 최초 1회만 설정 가능하고(이미 설정되어 있으면 재설정 요청 자체를 거부), 그 이후에는
// 앱을 통해 변경/재설정할 수 있는 경로를 아예 두지 않는다. 비밀번호를 바꾸려면 서버(파일 또는
// Redis)에 직접 접근할 수 있는 관리자만 저장된 값을 지우고 다시 설정해야 한다.
async function hasPassword() {
  const redis = getRedis();
  if (redis) return Boolean(await redis.get('auth:config'));
  return fs.existsSync(AUTH_FILE);
}

function hashPassword(password, salt) {
  return crypto.scryptSync(password, salt, 64).toString('hex');
}

async function setInitialPassword(password) {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = hashPassword(password, salt);
  const redis = getRedis();
  if (redis) {
    // setnx: 이미 키가 있으면 0을 반환 — 동시에 두 요청이 들어와도 하나만 성공한다.
    const created = await redis.setnx('auth:config', { salt, hash });
    if (!created) throw new Error('이미 비밀번호가 설정되어 있습니다.');
    return;
  }
  if (hasPasswordFileSync()) throw new Error('이미 비밀번호가 설정되어 있습니다.');
  // 'wx' 플래그로 배타적 생성: 동시에 두 요청이 들어와도 하나만 성공한다.
  fs.writeFileSync(AUTH_FILE, `${JSON.stringify({ salt, hash }, null, 2)}\n`, { flag: 'wx' });
}

function hasPasswordFileSync() {
  return fs.existsSync(AUTH_FILE);
}

async function verifyPassword(password) {
  const redis = getRedis();
  const config = redis ? await redis.get('auth:config') : (hasPasswordFileSync() ? JSON.parse(fs.readFileSync(AUTH_FILE, 'utf8')) : null);
  if (!config) return false;
  const candidate = Buffer.from(hashPassword(password, config.salt), 'hex');
  const stored = Buffer.from(config.hash, 'hex');
  return candidate.length === stored.length && crypto.timingSafeEqual(candidate, stored);
}

const SESSION_TTL_SECONDS = 7 * 24 * 60 * 60;
const sessions = new Map(); // Redis 미설정 시(로컬 개발) 폴백

async function createSession() {
  const token = crypto.randomBytes(32).toString('hex');
  const redis = getRedis();
  if (redis) { await redis.setex(`session:${token}`, SESSION_TTL_SECONDS, '1'); return token; }
  sessions.set(token, Date.now() + SESSION_TTL_SECONDS * 1000);
  return token;
}

async function isValidSession(token) {
  if (!token) return false;
  const redis = getRedis();
  if (redis) {
    const exists = await redis.get(`session:${token}`);
    if (!exists) return false;
    await redis.expire(`session:${token}`, SESSION_TTL_SECONDS);
    return true;
  }
  const expiresAt = sessions.get(token);
  if (!expiresAt) return false;
  if (Date.now() > expiresAt) { sessions.delete(token); return false; }
  sessions.set(token, Date.now() + SESSION_TTL_SECONDS * 1000);
  return true;
}

async function destroySession(token) {
  if (!token) return;
  const redis = getRedis();
  if (redis) { await redis.del(`session:${token}`); return; }
  sessions.delete(token);
}

function parseCookies(req) {
  const header = req.headers.cookie || '';
  return Object.fromEntries(header.split(';').map((part) => part.trim()).filter(Boolean).map((part) => {
    const index = part.indexOf('=');
    return index === -1 ? [part, ''] : [part.slice(0, index), decodeURIComponent(part.slice(index + 1))];
  }));
}

function setSessionCookie(res, token) {
  // Max-Age를 주지 않는 세션 쿠키: 브라우저(창)를 완전히 닫으면 사라져서, 다시 열 때마다 비밀번호를 다시 묻는다.
  res.setHeader('Set-Cookie', `session=${token}; HttpOnly; Path=/; SameSite=Strict`);
}

function clearSessionCookie(res) {
  res.setHeader('Set-Cookie', 'session=; HttpOnly; Path=/; Max-Age=0; SameSite=Strict');
}

// 로그인 실패가 반복되면 잠시 잠가 무차별 대입 시도를 늦춘다.
const LOGIN_MAX_ATTEMPTS = 10;
const LOGIN_LOCKOUT_MS = 5 * 60 * 1000;
const loginAttempts = new Map();

function isLoginLocked(ip) {
  const entry = loginAttempts.get(ip);
  return Boolean(entry && entry.lockedUntil && Date.now() < entry.lockedUntil);
}

function recordLoginFailure(ip) {
  const entry = loginAttempts.get(ip) || { count: 0, lockedUntil: 0 };
  entry.count += 1;
  if (entry.count >= LOGIN_MAX_ATTEMPTS) { entry.lockedUntil = Date.now() + LOGIN_LOCKOUT_MS; entry.count = 0; }
  loginAttempts.set(ip, entry);
}

function recordLoginSuccess(ip) {
  loginAttempts.delete(ip);
}

function kstDate(offsetDays = 0) {
  const now = new Date(Date.now() + 9 * 60 * 60 * 1000);
  now.setUTCDate(now.getUTCDate() + offsetDays);
  return `${now.getUTCFullYear()}${String(now.getUTCMonth() + 1).padStart(2, '0')}${String(now.getUTCDate()).padStart(2, '0')}`;
}

function displayDate(value) {
  const text = String(value || '').replace(/\D/g, '');
  return text.length === 8 ? `${text.slice(0, 4)}.${text.slice(4, 6)}.${text.slice(6, 8)}` : '';
}

function asArray(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function decodeHtml(value) {
  let text = String(value || '')
    .replace(/<!--[^]*?-->/g, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;|&#160;/gi, ' ')
    .replace(/&amp;/gi, '&').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, '<').replace(/&gt;/gi, '>');
  // 정부 사이트 원문에 이중 이스케이프된 특수문자 엔티티가 그대로 남는 경우가 있어 한 번 더 정리한다.
  const namedEntities = { lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', middot: '·', hellip: '…', mdash: '—', ndash: '–', amp: '&', quot: '"', apos: "'", lt: '<', gt: '>' };
  text = text.replace(/&([a-zA-Z]+);/g, (match, name) => namedEntities[name.toLowerCase()] ?? match)
    .replace(/&#(\d+);/g, (match, code) => String.fromCharCode(Number(code)));
  return text.replace(/\s+/g, ' ').trim();
}

function dateKey(value) {
  const text = String(value || '').trim();
  const parts = text.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (parts) return `${parts[1]}${parts[2].padStart(2, '0')}${parts[3].padStart(2, '0')}`;
  const digits = text.replace(/\D/g, '');
  return digits.length >= 8 ? digits.slice(0, 8) : '';
}

function isRecentPublication(item, publicationField, from, today) {
  const published = dateKey(item[publicationField]);
  return published >= from && published <= today;
}

function rowsFromHtml(html) {
  return [...String(html).matchAll(/<tbody\b[^>]*>([^]*?)<\/tbody>/gi)].flatMap((body) =>
    [...body[1].matchAll(/<tr\b[^>]*>([^]*?)<\/tr>/gi)].map((row) =>
      [...row[1].matchAll(/<td\b[^>]*>([^]*?)<\/td>/gi)].map((cell) => cell[1])
    )
  ).filter((cells) => cells.length);
}

function anchorFromCell(cell, baseUrl) {
  const anchor = String(cell || '').match(/<a\b[^>]*href="([^"]+)"[^>]*>([^]*?)<\/a>/i);
  if (!anchor) return { title: decodeHtml(cell), url: baseUrl };
  const href = anchor[1].replace(/&amp;/gi, '&').replace(/;jsessionid=[^?]+/i, '');
  return { title: decodeHtml(anchor[2]).replace(/^\[진행\]\s*/, ''), url: new URL(href, baseUrl).href };
}

async function requestHtml(url) {
  return limitLawRequest(async () => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), LAW_REQUEST_TIMEOUT_MS);
    try {
      const response = await fetch(url, { signal: controller.signal, headers: { accept: 'text/html', 'user-agent': 'Environment-Law-Monitor/1.0' } });
      if (!response.ok) throw new Error(`공식 목록 HTTP ${response.status}`);
      return await response.text();
    } finally { clearTimeout(timer); }
  });
}

function officialLawFromRow(cells, baseUrl) {
  if (cells.length < 8) return null;
  const nameCell = String(cells[1] || '');
  const anchor = nameCell.match(/<a\b([^>]*)href="([^"]+)"([^>]*)>([^]*?)<\/a>/i);
  if (!anchor) return null;
  const attributes = `${anchor[1]} ${anchor[3]}`;
  const titleAttribute = attributes.match(/\btitle="([^"]+)"/i);
  const title = decodeHtml(titleAttribute ? titleAttribute[1] : anchor[4]).replace(/\s*팝업으로 이동\s*$/, '');
  const href = anchor[2].replace(/&amp;/gi, '&').replace(/;jsessionid=[^?]+/i, '');
  const detailUrl = new URL(href, 'https://www.law.go.kr/LSW/').href;
  const lsiSeq = new URL(detailUrl).searchParams.get('lsiSeq') || `${dateKey(decodeHtml(cells[6]))}-${title}`;
  const changedAt = displayDate(dateKey(decodeHtml(cells[6])));
  const effectiveAt = displayDate(dateKey(decodeHtml(cells[7])));
  return {
    id: `law-${lsiSeq}`,
    group: 'revisedLaw',
    category: '법령',
    status: decodeHtml(cells[3]),
    title,
    agency: decodeHtml(cells[2]),
    changedAt,
    announcedAt: changedAt,
    effectiveAt,
    kind: decodeHtml(cells[4]),
    noticeNumber: decodeHtml(cells[5]),
    keywords: [],
    comparisonUrl: `https://www.law.go.kr/LSW/lsOldAndNew.do?lsiSeq=${encodeURIComponent(lsiSeq)}`,
    url: detailUrl
  };
}

async function fetchMceeRecentLawListings(from, today) {
  const items = [];
  const pageSize = 15;
  const maxPages = 10;
  for (let page = 0; page < maxPages; page += 1) {
    const url = new URL(MCEE_RECENT_LAWS_URL);
    url.searchParams.set('pagerOffset', String(page * pageSize));
    const html = await requestHtml(url);
    const pageItems = rowsFromHtml(html).map((cells) => officialLawFromRow(cells, url)).filter(Boolean);
    if (!pageItems.length) break;
    items.push(...pageItems);
    const oldestDate = pageItems.map((item) => dateKey(item.changedAt)).filter(Boolean).sort().at(0);
    if (!oldestDate || oldestDate < from) break;
  }
  const filtered = items.filter((item) => item.agency.includes(AGENCY) && dateKey(item.changedAt) >= from && dateKey(item.changedAt) <= today);
  const unique = new Map(filtered.map((item) => [item.id, item]));
  return Promise.all([...unique.values()].map(async (item) => {
    const detail = await getReasonDetail('law', { MST: item.id.replace(/^law-/, '') }, item.title);
    return { ...item, keywords: detail.keywords, summary: detail.summary };
  }));
}

async function fetchOfficialLawListings(from, today) {
  const sources = [RECENT_PROMULGATED_LAWS_URL, UPCOMING_LAWS_URL];
  const pages = await Promise.all(sources.map(async (url) => ({ url, html: await requestHtml(url) })));
  const items = pages.flatMap(({ url, html }) => rowsFromHtml(html).map((cells) => officialLawFromRow(cells, url)).filter(Boolean));
  const filtered = items.filter((item) => item.agency.includes(AGENCY) && dateKey(item.changedAt) >= from && dateKey(item.changedAt) <= today);
  const unique = new Map(filtered.map((item) => [item.id, item]));
  return Promise.all([...unique.values()].map(async (item) => {
    const detail = await getReasonDetail('law', { MST: item.id.replace(/^law-/, '') }, item.title);
    return { ...item, keywords: detail.keywords, summary: detail.summary };
  }));
}

async function fetchLegislationNotices(from, to) {
  const html = await requestHtml(LEGISLATION_NOTICE_URL);
  const items = rowsFromHtml(html).filter((cells) => cells.length >= 5).map((cells) => {
    const anchor = anchorFromCell(cells[1], LEGISLATION_NOTICE_URL);
    const published = dateKey(decodeHtml(cells[2]));
    return {
      id: `legislation-notice-${decodeHtml(cells[3])}`,
      group: 'legislationNotice', category: '법령', status: '입법예고', title: anchor.title,
      agency: decodeHtml(cells[4]), changedAt: displayDate(published), announcedAt: displayDate(published),
      effectiveAt: '', kind: '입법예고', noticeNumber: decodeHtml(cells[3]), url: anchor.url
    };
  }).filter((item) => item.agency.includes(AGENCY) && dateKey(item.changedAt) >= from && dateKey(item.changedAt) <= to);
  return Promise.all(items.map(async (item) => {
    const detail = await getNoticeDetail(item.url, item.title);
    return { ...item, keywords: detail.keywords, summary: detail.summary };
  }));
}

async function fetchAdministrativeNotices(from) {
  const html = await requestHtml(ADMIN_NOTICE_URL);
  const items = rowsFromHtml(html).filter((cells) => cells.length >= 7).map((cells) => {
    const anchor = anchorFromCell(cells[1], ADMIN_NOTICE_URL);
    const start = dateKey(decodeHtml(cells[3]));
    const end = dateKey(decodeHtml(cells[4]));
    return {
      id: `administrative-notice-${decodeHtml(cells[0])}-${start}`,
      group: 'administrativeNotice', category: '고시', status: '행정예고', title: anchor.title,
      agency: AGENCY, changedAt: displayDate(start), announcedAt: displayDate(start), effectiveAt: '',
      noticeEndAt: displayDate(end), kind: '행정예고', noticeNumber: decodeHtml(cells[2]), department: decodeHtml(cells[6]), url: anchor.url
    };
  // 예고 시작일이 최근 7일 안이거나 미래인 자료를 포함한다.
  // 이미 시작한 지 7일이 지난 예고는 목록에서 제외한다.
  }).filter((item) => dateKey(item.changedAt) >= from);
  return Promise.all(items.map(async (item) => {
    const detail = await getNoticeDetail(item.url, item.title);
    return { ...item, keywords: detail.keywords, summary: detail.summary };
  }));
}

async function requestApi(params) {
  const url = new URL(LAW_API);
  Object.entries({ OC: API_KEY, type: 'JSON', ...params }).forEach(([key, value]) => url.searchParams.set(key, value));
  return limitLawRequest(async () => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), LAW_REQUEST_TIMEOUT_MS);
    try {
      const response = await fetch(url, { signal: controller.signal, headers: { accept: 'application/json' } });
      const text = await response.text();
      if (!response.ok) throw new Error(`국가법령정보 API HTTP ${response.status}`);
      let data;
      try { data = JSON.parse(text); } catch { throw new Error('국가법령정보 API가 JSON이 아닌 응답을 반환했습니다.'); }
      const root = data.LawSearch || data.AdmRulSearch;
      if (!root) throw new Error(data.resultMsg || '국가법령정보 API 응답 형식을 확인할 수 없습니다.');
      if (root.resultCode && root.resultCode !== '00') throw new Error(root.resultMsg || `API 오류 ${root.resultCode}`);
      return root;
    } finally {
      clearTimeout(timer);
    }
  });
}

async function requestLawService(params) {
  const url = new URL(LAW_SERVICE_API);
  Object.entries({ OC: API_KEY, type: 'JSON', ...params }).forEach(([key, value]) => url.searchParams.set(key, value));
  return limitLawRequest(async () => {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), LAW_REQUEST_TIMEOUT_MS);
    try {
      const response = await fetch(url, { signal: controller.signal, headers: { accept: 'application/json' } });
      const text = await response.text();
      if (!response.ok) throw new Error(`국가법령정보 API HTTP ${response.status}`);
      try { return JSON.parse(text); } catch { throw new Error('국가법령정보 API가 JSON이 아닌 응답을 반환했습니다.'); }
    } finally {
      clearTimeout(timer);
    }
  });
}

function findField(value, fieldName) {
  if (!value || typeof value !== 'object') return '';
  if (value[fieldName] !== undefined) return Array.isArray(value[fieldName]) ? value[fieldName].flat(Infinity).join(' ') : String(value[fieldName]);
  for (const child of Object.values(value)) {
    const found = findField(child, fieldName);
    if (found) return found;
  }
  return '';
}

function extractReasonKeywords(reason, title = '') {
  const stopWords = new Set(['개정', '제정', '일부개정', '전부개정', '법령', '법률', '규정', '시행령', '시행규칙', '특별법', '내용', '사항', '이유', '주요내용', '관련', '관한', '소관', '경우', '등', '또', '하', '및', '대한', '위한', '따라', '통해', '현재', '현행', '필요', '목적', '마련', '정비', '개선', '보완', '만원', '이상', '이하', '종전', '앞으로', '사업', '사유', '결과', '수립', '장관', '기관']);
  const suffix = /(으로부터|으로는|으로|에서|에게|부터|까지|에는|이나|이며|이고|하고|의|은|는|을|를|와|과|에)$/u;
  const normalize = (value) => String(value || '').replace(/^[^가-힣A-Za-z0-9ㆍ·]+|[^가-힣A-Za-z0-9ㆍ·]+$/g, '').replace(suffix, '').trim();
  const useful = (value) => value.length >= 2 && !stopWords.has(value) && !/^\d+$/.test(value) && !/^(가|나|다|라|마|등|또|하)$/.test(value);
  const buckets = { target: new Map(), support: new Map(), regulation: new Map(), general: new Map() };
  const add = (bucket, value, score, index = 0) => {
    const parts = String(value || '').split(/\s+/).map(normalize).filter(useful).slice(-3);
    const phrase = parts.join(' ');
    if (!phrase || phrase.length > 24) return;
    const previous = bucket.get(phrase) || { phrase, score: 0, index };
    previous.score = Math.max(previous.score, score);
    previous.index = Math.min(previous.index, index);
    bucket.set(phrase, previous);
  };
  const text = decodeHtml(reason);
  const words = (text.match(/[가-힣A-Za-z0-9ㆍ·]{2,}/g) || []).map(normalize).filter(useful);
  const counts = new Map();
  words.forEach((word, index) => {
    const item = counts.get(word) || { count: 0, index };
    item.count += 1;
    counts.set(word, item);
  });
  const targetPattern = /(대상|시설|지역|유역|사업자|사업|노동자|전력망|발전소|폐기물|하천|토지|기관|업체)$/;
  const supportPattern = /(지원|지원금|보조금|융자|고용유지|재취업|보상|수수료|급여|기금)$/;
  const regulationPattern = /(허가|승인|부과|기준|절차|고시|의무|요건|납부|지정|등록|신고|관리|처분|제한|금지|완화|강화|시정|이행강제금|과태료)$/;
  [...counts.entries()].forEach(([word, value]) => {
    const score = value.count * 8;
    if (targetPattern.test(word)) add(buckets.target, word, score + 24, value.index);
    if (supportPattern.test(word)) add(buckets.support, word, score + 38, value.index);
    if (regulationPattern.test(word)) add(buckets.regulation, word, score + 34, value.index);
    add(buckets.general, word, score, value.index);
  });
  const phraseTokens = text.replace(/[◇·ㆍ,;:()\[\]<>]/g, ' ').split(/\s+/).map(normalize);
  phraseTokens.forEach((word, index) => {
    if (!useful(word)) return;
    const previous = normalize(phraseTokens[index - 1]);
    const beforePrevious = normalize(phraseTokens[index - 2]);
    if (supportPattern.test(word)) {
      const phrase = /^(지원|보상|융자)$/.test(word) && previous && useful(previous) ? `${previous} ${word}` : word;
      add(buckets.support, phrase, 78, index);
    }
    if (regulationPattern.test(word)) {
      const phrase = previous && useful(previous) && !/(장관|사유|경우)$/.test(previous) ? `${previous} ${word}` : word;
      add(buckets.regulation, phrase, 74, index);
    }
  });
  const regulationPhrases = /([가-힣]{2,}(?:계획|기준|절차|관리|점용료|시설|이행강제금|과태료))[^.]{0,35}?(승인|부과|고시|지정|납부|신설|보완|강화|완화)/g;
  for (const match of text.matchAll(regulationPhrases)) {
    add(buckets.regulation, `${normalize(match[1])} ${normalize(match[2])}`, 96, match.index || 0);
  }
  const titleWords = (decodeHtml(title).match(/[가-힣A-Za-z0-9ㆍ·]{2,}/g) || []).map(normalize).filter(useful);
  titleWords.forEach((word, index) => {
    add(buckets.target, word, 100 - index, index);
    if (supportPattern.test(word)) add(buckets.support, index ? `${titleWords[index - 1]} ${word}` : word, 90 - index, index);
  });
  const selected = [];
  const take = (bucket) => {
    const candidate = [...bucket.values()].sort((a, b) => b.score - a.score || a.index - b.index || b.phrase.length - a.phrase.length)
      .find((item) => !selected.some((chosen) => chosen === item.phrase || chosen.includes(item.phrase) || item.phrase.includes(chosen)));
    if (candidate) selected.push(candidate.phrase);
  };
  take(buckets.target);
  take(buckets.support);
  take(buckets.regulation);
  if (selected.length < 3) take(buckets.regulation);
  while (selected.length < 3) {
    const candidate = [...buckets.general.values()].sort((a, b) => b.score - a.score || a.index - b.index)
      .find((item) => !selected.some((chosen) => chosen === item.phrase || chosen.includes(item.phrase) || item.phrase.includes(chosen)));
    if (!candidate) break;
    selected.push(candidate.phrase);
    buckets.general.delete(candidate.phrase);
  }
  return selected.slice(0, 3);
}

function summarizeReasonText(raw) {
  return decodeHtml(raw).replace(/^[○ㅇ●\-\s]+/, '').trim();
}

async function getReasonDetail(target, identifiers, title = '') {
  try {
    const data = await requestLawService({ target, ...identifiers });
    const raw = findField(data, '제개정이유내용');
    return { keywords: extractReasonKeywords(raw, title), summary: summarizeReasonText(raw) };
  } catch { return { keywords: [], summary: '' }; }
}

// 입법예고·행정예고는 국가법령정보 API 대상이 아니라, 공고 상세 페이지를 직접 읽어
// "개정이유" ~ "주요내용" 구간 텍스트를 요약·키워드 추출에 사용한다.
async function getNoticeDetail(url, title = '') {
  try {
    const html = await requestHtml(url);
    const text = decodeHtml(html);
    const start = text.search(/\d\s*\.\s*(개정\s*이유|제정\s*이유|제안\s*이유|개정\s*취지)/);
    if (start === -1) return { keywords: [], summary: '' };
    const rest = text.slice(start);
    const end = rest.search(/\d\s*\.\s*(의견제출|시행일|부칙|문의처|담당자)/);
    const section = (end === -1 ? rest : rest.slice(0, end)).replace(/^\d\s*\.\s*(개정|제정|제안)\s*(이유|취지)\s*/, '');
    return { keywords: extractReasonKeywords(section, title), summary: summarizeReasonText(section) };
  } catch { return { keywords: [], summary: '' }; }
}

async function fetchAll(params, listKey) {
  const first = await requestApi({ ...params, display: 100, page: 1 });
  const total = Number(first.totalCnt || 0);
  const items = asArray(first[listKey]);
  const pages = Math.min(Math.ceil(total / 100), 10);
  for (let page = 2; page <= pages; page += 1) {
    const next = await requestApi({ ...params, display: 100, page });
    items.push(...asArray(next[listKey]));
  }
  return items;
}

function isRevisionOrNotice(item) {
  return /제정|개정|예고/.test(String(item['제개정구분명'] || ''));
}

function isRelevant(item) {
  return String(item['소관부처명'] || '').includes(AGENCY) && isRevisionOrNotice(item);
}

function isRelatedAdministrativeRule(item) {
  // org=1482000 조건으로 조회한 결과에는 본부와 산하기관 자료가 함께 들어온다.
  return Boolean(String(item['소관부처명'] || '').trim()) && isRevisionOrNotice(item);
}

async function getRecentChanges() {
  if (!API_KEY) throw new Error('LAW_API_OC 환경변수가 설정되지 않았습니다.');
  const dateKeys = Array.from({ length: 7 }, (_, index) => kstDate(-index));
  const range = `${dateKeys.at(-1)}~${dateKeys[0]}`;
  const today = dateKeys[0];
  const sourceErrors = [];
  const [recentLaws, adminRules, mceeLaws, officialLaws, legislationResult, administrativeResult] = await Promise.all([
    fetchAll({ target: 'law', org: '1482000', sort: 'ddes' }, 'law'),
    fetchAll({ target: 'admrul', mobileYn: 'Y', org: '1482000', sort: 'ddes' }, 'admrul'),
    fetchMceeRecentLawListings(dateKeys.at(-1), today).catch((error) => { sourceErrors.push(`환경부 최근 제·개정법령: ${error.message}`); return []; }),
    fetchOfficialLawListings(dateKeys.at(-1), today).catch((error) => { sourceErrors.push(`법령 보완 목록: ${error.message}`); return []; }),
    fetchLegislationNotices(dateKeys.at(-1), dateKeys[0]).catch((error) => { sourceErrors.push(`입법예고: ${error.message}`); return []; }),
    fetchAdministrativeNotices(dateKeys.at(-1)).catch((error) => { sourceErrors.push(`행정예고: ${error.message}`); return []; })
  ]);

  const laws = await Promise.all(recentLaws.filter((item) => isRelevant(item) && isRecentPublication(item, '공포일자', dateKeys.at(-1), today)).map(async (item) => {
    const detail = await getReasonDetail('law', { MST: item['법령일련번호'] }, item['법령명한글']);
    return {
      id: `law-${item['법령일련번호']}`,
      group: 'revisedLaw',
      category: '법령',
      status: item['제개정구분명'],
      title: item['법령명한글'],
      agency: item['소관부처명'],
      changedAt: displayDate(item['공포일자']),
      announcedAt: displayDate(item['공포일자']),
      effectiveAt: displayDate(item['시행일자']),
      kind: item['법령구분명'],
      keywords: detail.keywords,
      summary: detail.summary,
      comparisonUrl: `https://www.law.go.kr/LSW/lsOldAndNew.do?lsiSeq=${encodeURIComponent(item['법령일련번호'])}`,
      url: `https://www.law.go.kr/법령/${encodeURIComponent(item['법령명한글'] || '')}`
    };
  }));

  const rules = await Promise.all(adminRules.filter((item) => isRelatedAdministrativeRule(item) && isRecentPublication(item, '발령일자', dateKeys.at(-1), today)).map(async (item) => {
    const detail = await getReasonDetail('admrul', { ID: item['행정규칙일련번호'] }, item['행정규칙명']);
    return {
      id: `rule-${item['행정규칙일련번호']}`,
      group: 'revisedNotice',
      category: '행정규칙',
      status: item['제개정구분명'],
      title: item['행정규칙명'],
      agency: item['소관부처명'],
      changedAt: displayDate(item['발령일자']),
      announcedAt: displayDate(item['발령일자']),
      effectiveAt: displayDate(item['시행일자']),
      kind: item['행정규칙종류'],
      keywords: detail.keywords,
      summary: detail.summary,
      comparisonUrl: `https://www.law.go.kr/LSW/admRulOldAndNew.do?admRulSeq=${encodeURIComponent(item['행정규칙일련번호'])}`,
      url: `https://www.law.go.kr/행정규칙/${encodeURIComponent(item['행정규칙명'] || '')}`
    };
  }));

  // 환경부 최근 제·개정법령 목록을 마지막에 병합해, 동일 법령은 환경부 목록 정보를 우선한다.
  const unique = new Map([...legislationResult, ...officialLaws, ...laws, ...mceeLaws, ...administrativeResult, ...rules].map((item) => [item.id, item]));
  return { items: [...unique.values()].sort((a, b) => b.changedAt.localeCompare(a.changedAt)), sourceErrors };
}

function sendJson(res, status, body) {
  res.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' });
  res.end(JSON.stringify(body));
}

function escapeMailHtml(value) {
  return String(value || '').replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
}

function buildMailContent(items, from, to, introText = '기후에너지환경부 소관 최근 7일 법령·고시 변경사항을 구분하여 주성철 과장이 발송하였습니다.\n(스팸메일 아님)') {
  const groups = [
    ['legislationNotice', '변경 법령 · 입법예고'],
    ['revisedLaw', '변경 법령 · 개정법령'],
    ['administrativeNotice', '변경 행정규칙 · 행정예고'],
    ['revisedNotice', '변경 행정규칙 · 개정 행정규칙']
  ];
  const sections = groups.map(([key, label]) => {
    const selected = items.filter((item) => item.group === key);
    const rows = selected.length ? selected.map((item) => { const keywords = (item.keywords || []).slice(0, 3).map((keyword) => `<span style="display:inline-block;background:#eef4ff;color:#003874;border-radius:999px;padding:3px 8px;margin:4px 4px 0 0;font-size:12px;font-weight:600">${escapeMailHtml(keyword)}</span>`).join(''); const comparison = ['revisedLaw', 'revisedNotice'].includes(item.group) && item.comparisonUrl ? ` <a href="${escapeMailHtml(item.comparisonUrl)}" style="color:#003874">신구법 비교 링크 ↗</a>` : ''; return `<li style="margin:0 0 16px"><a href="${escapeMailHtml(item.url)}" style="color:#003874;font-weight:700;text-decoration:none">${escapeMailHtml(item.title)}</a><br><span style="color:#626873;font-size:13px">${escapeMailHtml(item.status)} · ${escapeMailHtml(item.changedAt)}${item.noticeEndAt ? ` · 예고종료 ${escapeMailHtml(item.noticeEndAt)}` : ''}${item.effectiveAt ? ` · 시행 ${escapeMailHtml(item.effectiveAt)}` : ''}</span>${keywords ? `<br>${keywords}` : ''}<br><a href="${escapeMailHtml(item.url)}" style="color:#003874;font-size:13px">원문 링크 ↗</a>${comparison}</li>`; }).join('') : '<li style="color:#737782">해당 기간 변경사항이 없습니다.</li>';
    return `<section style="margin:28px 0"><h2 style="font-size:18px;color:#1a1c20;border-bottom:1px solid #e2e8f0;padding-bottom:8px">${label} <span style="color:#626873;font-size:13px">${selected.length}건</span></h2><ul style="padding-left:20px">${rows}</ul></section>`;
  }).join('');
  const subject = `[환경법규 알림] ${from}~${to} 법령·행정규칙 변경사항 ${items.length}건`;
  const safeIntroText = String(introText || '').trim() || '기후에너지환경부 소관 최근 7일 법령·고시 변경사항을 구분하여 주성철 과장이 발송하였습니다.\n(스팸메일 아님)';
  const html = `<!doctype html><html lang="ko"><body style="margin:0;background:#f8fafc;font-family:Arial,'Noto Sans KR',sans-serif;color:#1a1c20"><div style="max-width:720px;margin:0 auto;padding:32px 20px"><div style="background:#003874;color:white;padding:24px;border-radius:10px 10px 0 0"><h1 style="margin:0;font-size:24px">${escapeMailHtml(MAIL_FROM_NAME)}</h1><p style="margin:8px 0 0;color:#d6e3ff">기후에너지환경부 소관 최근 7일 변경사항</p></div><main style="background:white;border:1px solid #e2e8f0;border-top:0;padding:24px;border-radius:0 0 10px 10px"><p>${escapeMailHtml(safeIntroText).replace(/\r?\n/g, '<br>')}</p><p><strong>조회기간</strong> ${from} ~ ${to}<br><strong>전체</strong> ${items.length}건</p>${sections}<p style="margin-top:32px;color:#737782;font-size:12px">본 메일은 국가법령정보와 기후에너지환경부 공식 예고 목록을 기준으로 생성되었습니다. 정확한 내용은 각 원문 링크에서 확인하세요.</p></main></div></body></html>`;
  const text = `${MAIL_FROM_NAME}\n${safeIntroText}\n조회기간: ${from} ~ ${to}\n전체 ${items.length}건\n\n${groups.map(([key, label]) => `${label}\n${items.filter((item) => item.group === key).map((item) => `- ${item.title} (${item.status}, ${item.changedAt})${item.keywords?.length ? `\n  키워드: ${item.keywords.slice(0, 3).join(', ')}` : ''}\n  원문 링크: ${item.url}${['revisedLaw', 'revisedNotice'].includes(item.group) && item.comparisonUrl ? `\n  신구법 비교 링크: ${item.comparisonUrl}` : ''}`).join('\n') || '- 해당 기간 변경사항 없음'}`).join('\n\n')}`;
  return { subject, html, text };
}

async function readJson(req) {
  const chunks = [];
  let size = 0;
  // 메일 발송 요청은 화면에서 불러온 변경사항 목록(항목별 개정이유 전문 포함)을 통째로 함께 보내므로
  // 넉넉하게 잡는다. 32KB였던 이전 한도는 요약 전문이 길어지면서 정상 요청도 막아버렸다.
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 2 * 1024 * 1024) throw new Error('요청 데이터가 너무 큽니다.');
    chunks.push(chunk);
  }
  try { return JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}'); }
  catch { throw new Error('요청 형식이 올바르지 않습니다.'); }
}

async function sendChangesMail(recipients, introText, listedItems) {
  if (!MAIL_USER) throw new Error('MAIL_USER가 설정되지 않았습니다.');
  if (!MAIL_APP_PASSWORD) throw new Error('Gmail 앱 비밀번호가 설정되지 않았습니다. .env의 MAIL_APP_PASSWORD를 입력한 뒤 서버를 다시 시작하세요.');
  const result = Array.isArray(listedItems) ? { items: listedItems, sourceErrors: [] } : await getRecentChanges();
  const from = displayDate(kstDate(-6));
  const to = displayDate(kstDate());
  const content = buildMailContent(result.items, from, to, introText);
  const pptAttachment = await buildPptSummary(result.items, from, to);
  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: MAIL_USER, pass: MAIL_APP_PASSWORD.replace(/\s+/g, '') }
  });
  const info = await transporter.sendMail({
    from: `"${MAIL_FROM_NAME.replace(/["\r\n]/g, '')}" <${MAIL_USER}>`,
    to: MAIL_USER,
    bcc: recipients,
    subject: content.subject,
    text: content.text,
    html: content.html,
    attachments: [{
      filename: `법규_개정_요약_${kstDate()}.pptx`,
      content: pptAttachment,
      contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    }]
  });
  return { messageId: info.messageId, accepted: info.accepted || [], rejected: info.rejected || [] };
}

async function handler(req, res) {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);
  const clientIp = req.socket.remoteAddress || 'unknown';

  if (requestUrl.pathname === '/api/auth/status' && req.method === 'GET') {
    const cookies = parseCookies(req);
    sendJson(res, 200, { hasPassword: await hasPassword(), authenticated: await isValidSession(cookies.session) });
    return;
  }
  if (requestUrl.pathname === '/api/auth/setup' && req.method === 'POST') {
    try {
      if (await hasPassword()) throw new Error('이미 비밀번호가 설정되어 있습니다.');
      const body = await readJson(req);
      const password = String(body.password || '');
      if (password.length < 8) throw new Error('비밀번호는 8자 이상이어야 합니다.');
      if (password.length > 200) throw new Error('비밀번호가 너무 깁니다.');
      await setInitialPassword(password);
      setSessionCookie(res, await createSession());
      sendJson(res, 200, { message: '비밀번호가 설정되었습니다.' });
    } catch (error) {
      sendJson(res, 400, { error: error.message });
    }
    return;
  }
  if (requestUrl.pathname === '/api/auth/login' && req.method === 'POST') {
    try {
      if (isLoginLocked(clientIp)) throw new Error('로그인 시도가 너무 많습니다. 잠시 후 다시 시도하세요.');
      if (!(await hasPassword())) throw new Error('아직 비밀번호가 설정되지 않았습니다.');
      const body = await readJson(req);
      const password = String(body.password || '');
      if (!(await verifyPassword(password))) {
        recordLoginFailure(clientIp);
        throw new Error('비밀번호가 올바르지 않습니다.');
      }
      recordLoginSuccess(clientIp);
      setSessionCookie(res, await createSession());
      sendJson(res, 200, { message: '로그인되었습니다.' });
    } catch (error) {
      sendJson(res, 401, { error: error.message });
    }
    return;
  }
  if (requestUrl.pathname === '/api/auth/logout' && req.method === 'POST') {
    const cookies = parseCookies(req);
    if (cookies.session) await destroySession(cookies.session);
    clearSessionCookie(res);
    sendJson(res, 200, { message: '로그아웃되었습니다.' });
    return;
  }

  const authenticated = await isValidSession(parseCookies(req).session);
  if (!authenticated) {
    if (requestUrl.pathname === '/' || requestUrl.pathname === '/index.html') {
      const html = fs.readFileSync(path.join(__dirname, 'login.html'));
      res.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' });
      res.end(html);
      return;
    }
    sendJson(res, 401, { error: '로그인이 필요합니다.' });
    return;
  }

  if (requestUrl.pathname === '/api/recipients') {
    try {
      if (req.method === 'GET') {
        sendJson(res, 200, { recipients: await readRecipients() });
        return;
      }
      if (req.method === 'PUT') {
        const body = await readJson(req);
        const requested = asArray(body.recipients);
        const recipients = normalizeRecipients(requested);
        if (requested.length !== recipients.length) throw new Error('올바르지 않은 이메일 주소가 포함되어 있습니다.');
        if (recipients.length > 50) throw new Error('수신자는 최대 50명까지 저장할 수 있습니다.');
        sendJson(res, 200, { recipients: await saveRecipients(recipients) });
        return;
      }
      sendJson(res, 405, { error: '허용되지 않은 요청 방식입니다.' });
    } catch (error) {
      sendJson(res, 400, { error: error.message });
    }
    return;
  }
  if (requestUrl.pathname === '/api/send-mail' && req.method === 'POST') {
    try {
      const body = await readJson(req);
      const recipients = normalizeRecipients(body.recipients);
      const introText = String(body.introText || '').trim();
      const listedItems = Array.isArray(body.items) ? body.items : null;
      if (!recipients.length) throw new Error('수신자를 한 명 이상 선택하세요.');
      if (recipients.length > 50) throw new Error('한 번에 최대 50명에게 발송할 수 있습니다.');
      if (recipients.some((email) => !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))) throw new Error('올바르지 않은 이메일 주소가 포함되어 있습니다.');
      if (introText.length > 5000) throw new Error('메일 본문은 5,000자 이내로 입력하세요.');
      const result = await sendChangesMail(recipients, introText, listedItems);
      sendJson(res, 200, { message: `${result.accepted.length}명에게 메일을 발송했습니다.`, ...result });
    } catch (error) {
      sendJson(res, 400, { error: error.message });
    }
    return;
  }
  if (requestUrl.pathname === '/api/changes') {
    try {
      const result = await getRecentChanges();
      sendJson(res, 200, { agency: AGENCY, from: displayDate(kstDate(-6)), to: displayDate(kstDate()), count: result.items.length, items: result.items, warnings: result.sourceErrors });
    } catch (error) {
      sendJson(res, 502, { error: error.name === 'AbortError' ? '국가법령정보 API 요청 시간이 초과되었습니다.' : error.message });
    }
    return;
  }
  if (requestUrl.pathname === '/api/ppt-summary' && req.method === 'GET') {
    try {
      const result = await getRecentChanges();
      const from = displayDate(kstDate(-6));
      const to = displayDate(kstDate());
      const file = await buildPptSummary(result.items, from, to);
      const filename = `법규_개정_요약_${kstDate()}.pptx`;
      res.writeHead(200, {
        'content-type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'content-disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`,
        'content-length': file.length,
        'cache-control': 'no-store'
      });
      res.end(file);
    } catch (error) {
      sendJson(res, 500, { error: `PPT 생성에 실패했습니다: ${error.message}` });
    }
    return;
  }
  if (requestUrl.pathname === '/' || requestUrl.pathname === '/index.html') {
    const html = fs.readFileSync(path.join(__dirname, 'index.html'));
    res.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' });
    res.end(html);
    return;
  }
  res.writeHead(404); res.end('Not found');
}

module.exports = handler;

if (require.main === module) {
  http.createServer(handler).listen(PORT, '127.0.0.1', () => {
    console.log(`환경법규 메일링 서비스: http://localhost:${PORT}`);
  });
}
