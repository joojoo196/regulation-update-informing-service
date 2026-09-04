const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const nodemailer = require('nodemailer');

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
const RECENT_PROMULGATED_LAWS_URL = 'https://www.law.go.kr/LSW/nwRvsLsPop.do?chrIdx=7&cptOfi=1482000&sortIdx=0';
const UPCOMING_LAWS_URL = 'https://www.law.go.kr/LSW/efLsPop.do?chrIdx=7&cptOfi=1482000&sortIdx=0';
const LEGISLATION_NOTICE_URL = 'https://mcee.go.kr/home/web/lawMaking/list.do';
const ADMIN_NOTICE_URL = 'https://mcee.go.kr/home/web/board/list.do?boardMasterId=827&menuId=10557&maxPageItems=100&pagerOffset=0';
const MAIL_FROM_NAME = process.env.MAIL_FROM_NAME || '환경법령 알림서비스';
const MAIL_USER = process.env.MAIL_USER || '';
const MAIL_APP_PASSWORD = process.env.MAIL_APP_PASSWORD || '';
const RECIPIENTS_FILE = path.join(__dirname, 'recipients.json');

function normalizeRecipients(values) {
  return [...new Set(asArray(values).map((value) => String(value).trim().toLowerCase()).filter((email) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)))];
}

function readRecipients() {
  try { return normalizeRecipients(JSON.parse(fs.readFileSync(RECIPIENTS_FILE, 'utf8'))); }
  catch (error) { return error.code === 'ENOENT' ? [] : []; }
}

function saveRecipients(recipients) {
  const saved = normalizeRecipients(recipients);
  const temporaryFile = `${RECIPIENTS_FILE}.tmp`;
  fs.writeFileSync(temporaryFile, `${JSON.stringify(saved, null, 2)}\n`, 'utf8');
  fs.renameSync(temporaryFile, RECIPIENTS_FILE);
  return saved;
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
  return String(value || '')
    .replace(/<!--[^]*?-->/g, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;|&#160;/gi, ' ')
    .replace(/&amp;/gi, '&').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'")
    .replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/\s+/g, ' ').trim();
}

function dateKey(value) {
  const text = String(value || '').trim();
  const parts = text.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (parts) return `${parts[1]}${parts[2].padStart(2, '0')}${parts[3].padStart(2, '0')}`;
  const digits = text.replace(/\D/g, '');
  return digits.length >= 8 ? digits.slice(0, 8) : '';
}

function isRecentOrUpcoming(item, publicationField, effectiveField, from, today, future) {
  const published = dateKey(item[publicationField]);
  const effective = dateKey(item[effectiveField]);
  return (published >= from && published <= today) || (effective >= today && effective <= future);
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
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  try {
    const response = await fetch(url, { signal: controller.signal, headers: { accept: 'text/html', 'user-agent': 'Environment-Law-Monitor/1.0' } });
    if (!response.ok) throw new Error(`공식 목록 HTTP ${response.status}`);
    return await response.text();
  } finally { clearTimeout(timer); }
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
  const url = new URL(href, baseUrl).href;
  const lsiSeq = new URL(url).searchParams.get('lsiSeq') || `${dateKey(decodeHtml(cells[6]))}-${title}`;
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
    url
  };
}

async function fetchOfficialLawListings(from, today, future) {
  const sources = [RECENT_PROMULGATED_LAWS_URL, UPCOMING_LAWS_URL];
  const pages = await Promise.all(sources.map(async (url) => ({ url, html: await requestHtml(url) })));
  const items = pages.flatMap(({ url, html }) => rowsFromHtml(html).map((cells) => officialLawFromRow(cells, url)).filter(Boolean));
  const filtered = items.filter((item) => item.agency === AGENCY && (
    (dateKey(item.changedAt) >= from && dateKey(item.changedAt) <= today) ||
    (dateKey(item.effectiveAt) >= today && dateKey(item.effectiveAt) <= future)
  ));
  const unique = new Map(filtered.map((item) => [item.id, item]));
  return Promise.all([...unique.values()].map(async (item) => ({
    ...item,
    keywords: await getReasonKeywords('law', { MST: item.id.replace(/^law-/, '') })
  })));
}

async function fetchLegislationNotices(from, to) {
  const html = await requestHtml(LEGISLATION_NOTICE_URL);
  return rowsFromHtml(html).filter((cells) => cells.length >= 5).map((cells) => {
    const anchor = anchorFromCell(cells[1], LEGISLATION_NOTICE_URL);
    const published = dateKey(decodeHtml(cells[2]));
    return {
      id: `legislation-notice-${decodeHtml(cells[3])}`,
      group: 'legislationNotice', category: '법령', status: '입법예고', title: anchor.title,
      agency: decodeHtml(cells[4]), changedAt: displayDate(published), announcedAt: displayDate(published),
      effectiveAt: '', kind: '입법예고', noticeNumber: decodeHtml(cells[3]), url: anchor.url
    };
  }).filter((item) => item.agency === AGENCY && dateKey(item.changedAt) >= from && dateKey(item.changedAt) <= to);
}

async function fetchAdministrativeNotices(from, to) {
  const html = await requestHtml(ADMIN_NOTICE_URL);
  return rowsFromHtml(html).filter((cells) => cells.length >= 7).map((cells) => {
    const anchor = anchorFromCell(cells[1], ADMIN_NOTICE_URL);
    const start = dateKey(decodeHtml(cells[3]));
    const end = dateKey(decodeHtml(cells[4]));
    return {
      id: `administrative-notice-${decodeHtml(cells[0])}-${start}`,
      group: 'administrativeNotice', category: '고시', status: '행정예고', title: anchor.title,
      agency: AGENCY, changedAt: displayDate(start), announcedAt: displayDate(start), effectiveAt: '',
      noticeEndAt: displayDate(end), kind: '행정예고', noticeNumber: decodeHtml(cells[2]), department: decodeHtml(cells[6]), url: anchor.url
    };
  }).filter((item) => dateKey(item.changedAt) >= from && dateKey(item.changedAt) <= to);
}

async function requestApi(params) {
  const url = new URL(LAW_API);
  Object.entries({ OC: API_KEY, type: 'JSON', ...params }).forEach(([key, value]) => url.searchParams.set(key, value));
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
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
}

async function requestLawService(params) {
  const url = new URL(LAW_SERVICE_API);
  Object.entries({ OC: API_KEY, type: 'JSON', ...params }).forEach(([key, value]) => url.searchParams.set(key, value));
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 15000);
  try {
    const response = await fetch(url, { signal: controller.signal, headers: { accept: 'application/json' } });
    const text = await response.text();
    if (!response.ok) throw new Error(`국가법령정보 API HTTP ${response.status}`);
    try { return JSON.parse(text); } catch { throw new Error('국가법령정보 API가 JSON이 아닌 응답을 반환했습니다.'); }
  } finally {
    clearTimeout(timer);
  }
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

function extractReasonKeywords(reason) {
  const stopWords = new Set(['개정', '제정', '일부개정', '전부개정', '법령', '법률', '규정', '시행령', '시행규칙', '내용', '사항', '이유', '주요내용', '관련', '관한', '소관', '경우', '등의', '등을', '등에', '위한', '따라', '통해', '대한', '현재', '현행', '이러한', '위하여', '하고', '있는', '있도록', '필요', '목적', '마련', '정비', '개선', '보완']);
  const words = decodeHtml(reason).match(/[가-힣]{2,}/g) || [];
  const counts = new Map();
  words.forEach((rawWord, index) => {
    const word = rawWord.replace(/(뿐만|으로|에서|에게|부터|까지|이나|이며|이고|하고|에는|에는|으로|의|은|는|이|가|을|를|와|과|에)$/u, '');
    if (stopWords.has(word)) return;
    const item = counts.get(word) || { word, count: 0, index };
    item.count += 1;
    counts.set(word, item);
  });
  return [...counts.values()].sort((a, b) => b.count - a.count || a.index - b.index).slice(0, 3).map((item) => item.word);
}

async function getReasonKeywords(target, identifiers) {
  try {
    const data = await requestLawService({ target, ...identifiers });
    return extractReasonKeywords(findField(data, '제개정이유내용'));
  } catch { return []; }
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

function isRelevant(item) {
  return String(item['소관부처명'] || '').trim() === AGENCY && /개정|예고/.test(String(item['제개정구분명'] || ''));
}

async function getRecentChanges() {
  if (!API_KEY) throw new Error('LAW_API_OC 환경변수가 설정되지 않았습니다.');
  const dateKeys = Array.from({ length: 7 }, (_, index) => kstDate(-index));
  const range = `${dateKeys.at(-1)}~${dateKeys[0]}`;
  const today = dateKeys[0];
  const future = kstDate(30);
  const sourceErrors = [];
  const [recentLaws, adminRules, officialLaws, legislationResult, administrativeResult] = await Promise.all([
    fetchAll({ target: 'law', org: '1482000', sort: 'ddes' }, 'law'),
    fetchAll({ target: 'admrul', mobileYn: 'Y', org: '1482000', sort: 'ddes' }, 'admrul'),
    fetchOfficialLawListings(dateKeys.at(-1), today, future).catch((error) => { sourceErrors.push(`법령 보완 목록: ${error.message}`); return []; }),
    fetchLegislationNotices(dateKeys.at(-1), dateKeys[0]).catch((error) => { sourceErrors.push(`입법예고: ${error.message}`); return []; }),
    fetchAdministrativeNotices(dateKeys.at(-1), dateKeys[0]).catch((error) => { sourceErrors.push(`행정예고: ${error.message}`); return []; })
  ]);

  const laws = await Promise.all(recentLaws.filter((item) => isRelevant(item) && isRecentOrUpcoming(item, '공포일자', '시행일자', dateKeys.at(-1), today, future)).map(async (item) => ({
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
    keywords: await getReasonKeywords('law', { MST: item['법령일련번호'] }),
    comparisonUrl: `https://www.law.go.kr/LSW/lsOldAndNew.do?lsiSeq=${encodeURIComponent(item['법령일련번호'])}`,
    url: `https://www.law.go.kr/법령/${encodeURIComponent(item['법령명한글'] || '')}`
  })));

  const rules = await Promise.all(adminRules.filter((item) => isRelevant(item) && item['행정규칙종류'] === '고시' && isRecentOrUpcoming(item, '발령일자', '시행일자', dateKeys.at(-1), today, future)).map(async (item) => ({
    id: `rule-${item['행정규칙일련번호']}`,
    group: 'revisedNotice',
    category: '고시',
    status: item['제개정구분명'],
    title: item['행정규칙명'],
    agency: item['소관부처명'],
    changedAt: displayDate(item['발령일자']),
    announcedAt: displayDate(item['발령일자']),
    effectiveAt: displayDate(item['시행일자']),
    kind: item['행정규칙종류'],
    keywords: await getReasonKeywords('admrul', { ID: item['행정규칙일련번호'] }),
    comparisonUrl: `https://www.law.go.kr/LSW/admRulOldAndNew.do?admRulSeq=${encodeURIComponent(item['행정규칙일련번호'])}`,
    url: `https://www.law.go.kr/행정규칙/${encodeURIComponent(item['행정규칙명'] || '')}`
  })));

  const unique = new Map([...legislationResult, ...officialLaws, ...laws, ...administrativeResult, ...rules].map((item) => [item.id, item]));
  return { items: [...unique.values()].sort((a, b) => b.changedAt.localeCompare(a.changedAt)), sourceErrors };
}

function sendJson(res, status, body) {
  res.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' });
  res.end(JSON.stringify(body));
}

function escapeMailHtml(value) {
  return String(value || '').replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
}

function buildMailContent(items, from, to, introText = '기후에너지환경부 소관 최근 7일 법령·고시 변경사항을 구분하여 발송합니다.') {
  const groups = [
    ['legislationNotice', '변경 법령 · 입법예고'],
    ['revisedLaw', '변경 법령 · 개정법령'],
    ['administrativeNotice', '변경 고시 · 행정예고'],
    ['revisedNotice', '변경 고시 · 개정고시']
  ];
  const sections = groups.map(([key, label]) => {
    const selected = items.filter((item) => item.group === key);
    const rows = selected.length ? selected.map((item) => { const keywords = (item.keywords || []).slice(0, 3).map((keyword) => `<span style="display:inline-block;background:#eef4ff;color:#003874;border-radius:999px;padding:3px 8px;margin:4px 4px 0 0;font-size:12px;font-weight:600">${escapeMailHtml(keyword)}</span>`).join(''); const comparison = ['revisedLaw', 'revisedNotice'].includes(item.group) && item.comparisonUrl ? ` <a href="${escapeMailHtml(item.comparisonUrl)}" style="color:#003874">신구법 비교 링크 ↗</a>` : ''; return `<li style="margin:0 0 16px"><a href="${escapeMailHtml(item.url)}" style="color:#003874;font-weight:700;text-decoration:none">${escapeMailHtml(item.title)}</a><br><span style="color:#626873;font-size:13px">${escapeMailHtml(item.status)} · ${escapeMailHtml(item.changedAt)}${item.noticeEndAt ? ` · 예고종료 ${escapeMailHtml(item.noticeEndAt)}` : ''}${item.effectiveAt ? ` · 시행 ${escapeMailHtml(item.effectiveAt)}` : ''}</span>${keywords ? `<br>${keywords}` : ''}<br><a href="${escapeMailHtml(item.url)}" style="color:#003874;font-size:13px">원문 링크 ↗</a>${comparison}</li>`; }).join('') : '<li style="color:#737782">해당 기간 변경사항이 없습니다.</li>';
    return `<section style="margin:28px 0"><h2 style="font-size:18px;color:#1a1c20;border-bottom:1px solid #e2e8f0;padding-bottom:8px">${label} <span style="color:#626873;font-size:13px">${selected.length}건</span></h2><ul style="padding-left:20px">${rows}</ul></section>`;
  }).join('');
  const subject = `[환경법규 알림] ${from}~${to} 법령·고시 변경사항 ${items.length}건`;
  const safeIntroText = String(introText || '').trim() || '기후에너지환경부 소관 최근 7일 법령·고시 변경사항을 구분하여 발송합니다.';
  const html = `<!doctype html><html lang="ko"><body style="margin:0;background:#f8fafc;font-family:Arial,'Noto Sans KR',sans-serif;color:#1a1c20"><div style="max-width:720px;margin:0 auto;padding:32px 20px"><div style="background:#003874;color:white;padding:24px;border-radius:10px 10px 0 0"><h1 style="margin:0;font-size:24px">${escapeMailHtml(MAIL_FROM_NAME)}</h1><p style="margin:8px 0 0;color:#d6e3ff">기후에너지환경부 소관 최근 7일 변경사항</p></div><main style="background:white;border:1px solid #e2e8f0;border-top:0;padding:24px;border-radius:0 0 10px 10px"><p>${escapeMailHtml(safeIntroText).replace(/\r?\n/g, '<br>')}</p><p><strong>조회기간</strong> ${from} ~ ${to}<br><strong>전체</strong> ${items.length}건</p>${sections}<p style="margin-top:32px;color:#737782;font-size:12px">본 메일은 국가법령정보와 기후에너지환경부 공식 예고 목록을 기준으로 생성되었습니다. 정확한 내용은 각 원문 링크에서 확인하세요.</p></main></div></body></html>`;
  const text = `${MAIL_FROM_NAME}\n${safeIntroText}\n조회기간: ${from} ~ ${to}\n전체 ${items.length}건\n\n${groups.map(([key, label]) => `${label}\n${items.filter((item) => item.group === key).map((item) => `- ${item.title} (${item.status}, ${item.changedAt})${item.keywords?.length ? `\n  키워드: ${item.keywords.slice(0, 3).join(', ')}` : ''}\n  원문 링크: ${item.url}${['revisedLaw', 'revisedNotice'].includes(item.group) && item.comparisonUrl ? `\n  신구법 비교 링크: ${item.comparisonUrl}` : ''}`).join('\n') || '- 해당 기간 변경사항 없음'}`).join('\n\n')}`;
  return { subject, html, text };
}

async function readJson(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > 32 * 1024) throw new Error('요청 데이터가 너무 큽니다.');
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
    html: content.html
  });
  return { messageId: info.messageId, accepted: info.accepted || [], rejected: info.rejected || [] };
}

async function handler(req, res) {
  const requestUrl = new URL(req.url, `http://${req.headers.host}`);
  if (requestUrl.pathname === '/api/recipients') {
    try {
      if (req.method === 'GET') {
        sendJson(res, 200, { recipients: readRecipients() });
        return;
      }
      if (req.method === 'PUT') {
        const body = await readJson(req);
        const requested = asArray(body.recipients);
        const recipients = normalizeRecipients(requested);
        if (requested.length !== recipients.length) throw new Error('올바르지 않은 이메일 주소가 포함되어 있습니다.');
        if (recipients.length > 50) throw new Error('수신자는 최대 50명까지 저장할 수 있습니다.');
        sendJson(res, 200, { recipients: saveRecipients(recipients) });
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
