require('dotenv').config();

const { createClient } = require('@supabase/supabase-js');
const COS = require('cos-nodejs-sdk-v5');
const XLSX = require('xlsx');

const INBOX_BUCKET = 'keyword-task-inbox';
const KEYWORD_LIMIT = 20;
const REPORT_URL_EXPIRES_SECONDS = 3600;
const POLL_INTERVAL_MS = 30000;
const URL_REFRESH_INTERVAL_MS = 30 * 60 * 1000;

const requiredVars = [
  'SUPABASE_URL',
  'SUPABASE_SERVICE_ROLE_KEY',
  'TENCENT_SECRET_ID',
  'TENCENT_SECRET_KEY',
  'COS_REGION',
  'COS_BUCKET',
  'COS_REPORT_PREFIX',
  'SIF_MCP_URL'
];

for (const name of requiredVars) {
  if (!process.env[name]) {
    console.error(`缺少环境变量：${name}`);
    process.exit(1);
  }
}

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY,
  { auth: { persistSession: false, autoRefreshToken: false } }
);

const cos = new COS({
  SecretId: process.env.TENCENT_SECRET_ID,
  SecretKey: process.env.TENCENT_SECRET_KEY
});

let polling = false;
let refreshingUrls = false;

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function putObject(params) {
  return new Promise((resolve, reject) => {
    cos.putObject(params, (error, data) => error ? reject(error) : resolve(data));
  });
}

function getSignedObjectUrl(objectKey) {
  return new Promise((resolve, reject) => {
    cos.getObjectUrl({
      Bucket: process.env.COS_BUCKET,
      Region: process.env.COS_REGION,
      Key: objectKey,
      Sign: true,
      Expires: REPORT_URL_EXPIRES_SECONDS
    }, (error, data) => error ? reject(error) : resolve(data.Url));
  });
}

function reportObjectKey(task) {
  const prefix = process.env.COS_REPORT_PREFIX.replace(/^\/+|\/+$/g, '');
  return `${prefix}/${task.user_id}/${task.id}.html`;
}

function normalizeHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s_\-()（）/#＃]+/g, '');
}

function numberValue(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const cleaned = String(value ?? '')
    .replace(/[$￥¥,%\s,]/g, '')
    .trim();
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function firstMatchingHeader(headers, aliases) {
  const normalizedAliases = aliases.map(normalizeHeader);
  return headers.find(header => normalizedAliases.includes(normalizeHeader(header)));
}

function parseAdvertisingWorkbook(buffer, reportPath) {
  const workbook = XLSX.read(buffer, {
    type: 'buffer',
    raw: false,
    codepage: 65001
  });

  const aliases = {
    term: ['搜索词', '客户搜索词', '顾客搜索词', 'search term', 'customer search term'],
    impressions: ['曝光量', '曝光', '展示量', 'impressions'],
    clicks: ['点击量', '点击', 'clicks'],
    spend: ['花费', '支出', '广告花费', 'spend', 'cost'],
    orders: ['7天总订单数', '14天总订单数', '订单量', '订单', 'orders', 'total orders'],
    sales: ['7天总销售额', '14天总销售额', '销售额', '广告销售额', 'sales', 'total sales']
  };

  let selected = null;

  for (const sheetName of workbook.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      defval: '',
      raw: false
    });
    if (!rows.length) continue;
    const headers = Object.keys(rows[0]);
    const mapping = {};
    for (const [field, choices] of Object.entries(aliases)) {
      mapping[field] = firstMatchingHeader(headers, choices);
    }
    if (mapping.term) {
      selected = { sheetName, rows, headers, mapping };
      break;
    }
  }

  if (!selected) {
    const sheetNames = workbook.SheetNames.join('、');
    throw new Error(`广告报表未识别到“搜索词”列。文件：${reportPath}；工作表：${sheetNames}`);
  }

  const aggregate = new Map();
  for (const row of selected.rows) {
    const keyword = String(row[selected.mapping.term] || '').trim().toLowerCase();
    if (!keyword) continue;
    const current = aggregate.get(keyword) || {
      impressions: 0,
      clicks: 0,
      spend: 0,
      orders: 0,
      sales: 0
    };
    for (const field of ['impressions', 'clicks', 'spend', 'orders', 'sales']) {
      if (selected.mapping[field]) current[field] += numberValue(row[selected.mapping[field]]);
    }
    aggregate.set(keyword, current);
  }

  return {
    sheetName: selected.sheetName,
    headers: selected.headers,
    mapping: selected.mapping,
    aggregate
  };
}

async function downloadAdvertisingReport(task) {
  const { data, error } = await supabase.storage
    .from(INBOX_BUCKET)
    .download(task.report_path);
  if (error) throw new Error(`下载广告报表失败：${error.message}`);
  return Buffer.from(await data.arrayBuffer());
}

async function callSifTool(name, args, attempts = 3) {
  let lastError;
  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      const response = await fetch(process.env.SIF_MCP_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json, text/event-stream'
        },
        body: JSON.stringify({
          jsonrpc: '2.0',
          id: `${Date.now()}-${attempt}`,
          method: 'tools/call',
          params: { name, arguments: args }
        }),
        signal: AbortSignal.timeout(90000)
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const envelope = await response.json();
      if (envelope.error) throw new Error(JSON.stringify(envelope.error));
      if (envelope.result?.isError) {
        throw new Error(envelope.result.content?.[0]?.text || `${name} 返回错误`);
      }
      const text = envelope.result?.content?.[0]?.text;
      if (!text) throw new Error(`${name} 没有返回内容`);
      return JSON.parse(text);
    } catch (error) {
      lastError = error;
      if (attempt < attempts) await sleep(attempt * 1000);
    }
  }
  throw new Error(`${name} 调用失败：${lastError.message}`);
}

function objectsWithKeyword(value, output = []) {
  if (!value || typeof value !== 'object') return output;
  if (!Array.isArray(value) && typeof value.keyword === 'string') output.push(value);
  for (const child of Object.values(value)) objectsWithKeyword(child, output);
  return output;
}

function keywordIndex(payload) {
  const index = new Map();
  for (const item of objectsWithKeyword(payload)) {
    const key = item.keyword.trim().toLowerCase();
    const existing = index.get(key) || {};
    index.set(key, { ...existing, ...item });
  }
  return index;
}

function rankText(value) {
  if (value === null || value === undefined || value === '') return '—';
  return String(value);
}

function trendText(row) {
  const direction = row.trend?.direction || row.momentum || row.contri_severity;
  const map = {
    growing: '增长',
    declining: '下降',
    stable: '稳定',
    volatile: '波动',
    accelerating: '加速增长',
    falling: '季节性回落',
    flat: '平稳',
    insufficient: '数据不足'
  };
  return map[direction] || String(direction || '数据不足');
}

function extractAsin(value) {
  const match = String(value || '').toUpperCase().match(/B0[A-Z0-9]{8}/);
  return match ? match[0] : '';
}

function highestChannelShare(competitor) {
  if (!competitor) return 0;
  return Math.max(
    numberValue(competitor.organic_share),
    numberValue(competitor.sp_share),
    numberValue(competitor.brand_share),
    numberValue(competitor.video_share)
  );
}

function median(values) {
  const sorted = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!sorted.length) return 0;
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

function deterministicDecision(row, highVolumeThreshold) {
  const ads = row.ads;
  const acos = ads.sales > 0 ? ads.spend / ads.sales : null;
  if (ads.spend > 0 && ads.orders <= 0) {
    return { priority: 1, label: '立即止损', advice: '广告有花费但无订单：先降竞价或暂停，并检查主图、价格与搜索词相关性。' };
  }
  if (ads.orders > 0 && acos !== null && acos > 0.3) {
    return { priority: 2, label: '控制成本', advice: `已有订单但 ACOS 为 ${(acos * 100).toFixed(1)}%：降低竞价并收紧匹配。` };
  }
  if (ads.orders > 0 && (acos === null || acos <= 0.3)) {
    return { priority: 3, label: '守住放大', advice: '已有订单且成本可控：转精准匹配，逐步增加预算并守住自然位。' };
  }
  if (row.searchVolume >= highVolumeThreshold && ads.spend <= 0) {
    const concentrated = row.topCompetitorShare >= 0.35;
    return concentrated
      ? { priority: 4, label: '长尾切入', advice: '搜索量高但头部集中：不高价硬抢，优先测试相关长尾词。' }
      : { priority: 4, label: '新增测试', advice: '搜索量高且尚未投放：以低风险预算测试精准或词组匹配。' };
  }
  if (/p1,[1-5]\//.test(String(row.organicRank || ''))) {
    return { priority: 5, label: '守自然位', advice: '自然排名靠前但广告证据不足：控制投入，重点守住自然位置。' };
  }
  return { priority: 6, label: '继续观察', advice: '当前证据不足：保持低预算观察，等待更多点击和订单后再调整。' };
}

async function getDoubaoDecisions(rows) {
  if (!process.env.DOUBAO_API_KEY || !process.env.DOUBAO_MODEL) return null;
  const endpoint = process.env.DOUBAO_API_URL || 'https://ark.cn-beijing.volces.com/api/v3/chat/completions';
  const compact = rows.map(row => ({
    keyword: row.keyword,
    search_volume: row.searchVolume,
    organic_rank: row.organicRank,
    ad_rank: row.adRank,
    top_competitor_share: row.topCompetitorShare,
    spend: row.ads.spend,
    orders: row.ads.orders,
    sales: row.ads.sales,
    acos: row.ads.sales > 0 ? row.ads.spend / row.ads.sales : null,
    fallback_decision: row.decision.advice
  }));

  const prompt = `你是亚马逊美国站广告运营助手。只按以下规则判断，不自由发挥：\n` +
    `1. 自己已进入点击前三且有订单：守住并适度放大。\n` +
    `2. 搜索量高但头部竞品点击占比过度集中：不高价硬抢，优先布局长尾词。\n` +
    `3. 广告高花费无订单：先降价或暂停，再检查 Listing 图片和相关性。\n` +
    `4. 有订单且 ACOS 达标：提高预算，转精准匹配持续放量。\n` +
    `5. 自然排名较好但广告表现差：控制广告投入，优先守自然位。\n` +
    `返回严格 JSON 数组，每项仅含 keyword、label、advice；advice 一句话，不输出 Markdown。数据：${JSON.stringify(compact)}`;

  try {
    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${process.env.DOUBAO_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: process.env.DOUBAO_MODEL,
        temperature: 0.1,
        messages: [{ role: 'user', content: prompt }]
      }),
      signal: AbortSignal.timeout(120000)
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}: ${await response.text()}`);
    const payload = await response.json();
    let content = payload.choices?.[0]?.message?.content || '';
    content = content.replace(/^```(?:json)?\s*|\s*```$/g, '');
    const parsed = JSON.parse(content);
    if (!Array.isArray(parsed)) throw new Error('返回内容不是数组');
    return new Map(parsed.map(item => [String(item.keyword || '').toLowerCase(), item]));
  } catch (error) {
    console.warn(`豆包判断失败，使用确定性规则：${error.message}`);
    return null;
  }
}

async function buildKeywordRows(task, adReport) {
  console.log(`任务 ${task.id}：SIF 反查前 ${KEYWORD_LIMIT} 个关键词。`);
  const signals = await callSifTool('market_get_asin_keyword_signals', {
    asin: task.asin,
    country: 'US',
    time_type: 'lately',
    time_value: '30',
    topN: KEYWORD_LIMIT
  });
  const signalRows = Array.isArray(signals.top_keywords) ? signals.top_keywords : [];
  if (!signalRows.length) throw new Error(`SIF 未返回 ${task.asin} 最近 30 天关键词数据`);

  const keywords = signalRows.map(item => item.keyword).filter(Boolean).slice(0, KEYWORD_LIMIT);
  console.log(`任务 ${task.id}：批量补齐 ${keywords.length} 个词的需求数据。`);
  const demandPayload = await callSifTool('market_get_keyword_demand', {
    country: 'US',
    keywords
  });
  const demandByKeyword = keywordIndex(demandPayload);
  const competitionByKeyword = new Map();

  for (let index = 0; index < keywords.length; index += 1) {
    const keyword = keywords[index];
    console.log(`任务 ${task.id}：竞争数据 ${index + 1}/${keywords.length} - ${keyword}`);
    try {
      const competition = await callSifTool('market_get_keyword_competition', {
        asin: task.asin,
        country: 'US',
        keyword,
        rank_evolution: false,
        time_type: 'all'
      });
      competitionByKeyword.set(keyword.toLowerCase(), competition);
    } catch (error) {
      console.warn(`关键词 ${keyword} 竞争数据缺失：${error.message}`);
    }
    await sleep(400);
  }

  const searchVolumes = signalRows.map(row => numberValue(row.search_volume));
  const highVolumeThreshold = median(searchVolumes);
  const rows = signalRows.map(signal => {
    const keyword = signal.keyword.trim();
    const demand = demandByKeyword.get(keyword.toLowerCase()) || {};
    const competition = competitionByKeyword.get(keyword.toLowerCase()) || {};
    const topCompetitor = competition.top_competitors?.[0] || null;
    const ads = adReport.aggregate.get(keyword.toLowerCase()) || {
      impressions: 0,
      clicks: 0,
      spend: 0,
      orders: 0,
      sales: 0
    };
    const row = {
      keyword,
      searchVolume: numberValue(signal.search_volume || demand.search_volume),
      trend: trendText(demand),
      organicRank: signal.organic_rank,
      adRank: signal.sp_rank,
      trafficShare: numberValue(signal.traffic_share),
      topCompetitorAsin: extractAsin(topCompetitor?.asin),
      // SIF 的 total_share 会把多个广告/自然渠道相加，可能超过 100%。
      // 报告改为展示该竞品占比最高的单一渠道，避免误读。
      topCompetitorShare: highestChannelShare(topCompetitor),
      topCompetitorPrice: numberValue(topCompetitor?.price),
      ads
    };
    row.decision = deterministicDecision(row, highVolumeThreshold);
    return row;
  });

  const doubao = await getDoubaoDecisions(rows);
  if (doubao) {
    for (const row of rows) {
      const item = doubao.get(row.keyword.toLowerCase());
      if (item?.advice) {
        row.decision = {
          ...row.decision,
          label: String(item.label || row.decision.label).slice(0, 20),
          advice: String(item.advice).slice(0, 300)
        };
      }
    }
  }

  return rows.sort((a, b) =>
    a.decision.priority - b.decision.priority ||
    b.ads.spend - a.ads.spend ||
    b.searchVolume - a.searchVolume
  );
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatNumber(value, digits = 0) {
  const number = numberValue(value);
  return number.toLocaleString('en-US', {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function renderHtmlReport(task, rows, adReport) {
  const totalSpend = rows.reduce((sum, row) => sum + row.ads.spend, 0);
  const totalOrders = rows.reduce((sum, row) => sum + row.ads.orders, 0);
  const actionCount = rows.filter(row => row.decision.priority <= 3).length;
  const scaleCount = rows.filter(row => row.decision.priority === 3).length;
  const controlCount = rows.filter(row => row.decision.priority <= 2).length;
  const observeCount = rows.length - actionCount;
  const generatedAt = new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' });

  const bodyRows = rows.map(row => {
    const acos = row.ads.sales > 0 ? `${(row.ads.spend / row.ads.sales * 100).toFixed(1)}%` : '—';
    const competitor = row.topCompetitorAsin
      ? `<a href="https://www.amazon.com/dp/${escapeHtml(row.topCompetitorAsin)}" target="_blank" rel="noreferrer">${escapeHtml(row.topCompetitorAsin)}</a>`
      : '数据缺失';
    return `<tr data-action="${escapeHtml(row.decision.label)}" data-keyword="${escapeHtml(row.keyword.toLowerCase())}">
      <td class="sticky"><strong>${escapeHtml(row.keyword)}</strong></td>
      <td>${formatNumber(row.searchVolume)}</td>
      <td>${escapeHtml(row.trend)}</td>
      <td>${escapeHtml(rankText(row.organicRank))}</td>
      <td>${escapeHtml(rankText(row.adRank))}</td>
      <td>${competitor}</td>
      <td>${row.topCompetitorShare ? `${(row.topCompetitorShare * 100).toFixed(1)}%` : '—'}</td>
      <td>${formatNumber(row.ads.impressions)}</td>
      <td>${formatNumber(row.ads.clicks)}</td>
      <td>$${formatNumber(row.ads.spend, 2)}</td>
      <td>${formatNumber(row.ads.orders)}</td>
      <td>${acos}</td>
      <td><span class="tag p${row.decision.priority}">${escapeHtml(row.decision.label)}</span></td>
      <td class="advice">${escapeHtml(row.decision.advice)}</td>
    </tr>`;
  }).join('\n');

  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${escapeHtml(task.asin)} 关键词作战报告</title>
  <style>
    :root{color-scheme:light;--ink:#102040;--muted:#66758f;--line:#dfe6f2;--blue:#2864e8;--navy:#123773;--red:#c9372c;--amber:#b77905;--green:#067647;--bg:#f1f5fb}
    *{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.55 system-ui,-apple-system,"Segoe UI","Microsoft YaHei",sans-serif}
    .shell{display:grid;grid-template-columns:220px minmax(0,1fr);gap:20px;max-width:1720px;margin:0 auto;padding:20px}.side{position:sticky;top:20px;height:calc(100vh - 40px);background:#fff;border:1px solid var(--line);border-radius:18px;padding:20px;box-shadow:0 12px 34px rgba(27,54,99,.07)}
    .brand{font-size:17px;font-weight:800}.asin{color:var(--muted);font-size:12px;margin:3px 0 20px}.nav a{display:flex;gap:10px;color:#42526e;text-decoration:none;padding:11px 12px;border-radius:10px;margin:5px 0;font-weight:650}.nav a.active,.nav a:hover{background:#edf4ff;color:var(--blue)}.nav b{color:#8ca0bf;font-size:11px}
    main{min-width:0}.hero{position:relative;overflow:hidden;background:linear-gradient(125deg,#123773,#2864e8);color:#fff;padding:30px;border-radius:20px;box-shadow:0 18px 46px rgba(30,85,190,.2)}.hero:after{content:"";position:absolute;width:340px;height:340px;border:56px solid rgba(255,255,255,.07);border-radius:50%;right:-90px;top:-150px}
    h1{margin:0 0 8px;font-size:30px}.sub{opacity:.82}.eyebrow{font-size:11px;letter-spacing:.16em;font-weight:800;color:#a9c8ff;margin-bottom:8px}.cards{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:12px;margin:18px 0}
    .card{background:#fff;border:1px solid var(--line);border-radius:15px;padding:17px;box-shadow:0 8px 24px rgba(31,55,95,.04)}.card span{display:block;color:var(--muted);font-size:12px}.card strong{display:block;margin-top:5px;font-size:25px}.decision-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px}.decision{background:#fff;border:1px solid var(--line);border-radius:15px;padding:16px;border-top:3px solid var(--blue)}.decision.danger{border-top-color:var(--red)}.decision.good{border-top-color:var(--green)}.decision span{color:var(--muted)}.decision strong{display:block;font-size:22px;margin-top:5px}
    .panel{background:#fff;border:1px solid var(--line);border-radius:16px;overflow:hidden;box-shadow:0 10px 30px rgba(31,55,95,.05)}.panel-head{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:16px 18px;border-bottom:1px solid var(--line)}.tools{display:flex;gap:8px;flex-wrap:wrap}.tools input,.tools select,.tools button{border:1px solid var(--line);border-radius:9px;padding:9px 11px;background:#f9fbfe;color:var(--ink)}.tools button{border-color:var(--blue);background:var(--blue);color:#fff;font-weight:750;cursor:pointer}.tools button:hover{filter:brightness(.94)}
    .table-wrap{overflow:auto;max-height:72vh}table{border-collapse:separate;border-spacing:0;min-width:1700px;width:100%}th,td{padding:12px;border-bottom:1px solid var(--line);vertical-align:top;text-align:left;white-space:nowrap}
    th{position:sticky;top:0;background:#edf3fd;z-index:2;font-size:12px;color:#41577a}.sticky{position:sticky;left:0;background:#fff;z-index:1;min-width:230px}.advice{white-space:normal;min-width:340px;color:#42526e}
    .tag{display:inline-block;border-radius:999px;padding:3px 9px;font-weight:700}.p1,.p2{background:#fee4e2;color:var(--red)}.p3{background:#dcfae6;color:var(--green)}.p4{background:#fef0c7;color:var(--amber)}.p5,.p6{background:#eef2f7;color:#475467}
    .report-view{display:none}.report-view.active{display:block}.decision-grid.report-view.active{display:grid}.method-card{background:#fff;border:1px solid var(--line);border-radius:16px;padding:22px}.method-card h2{margin-top:0}.method-card li{margin:9px 0;color:#42526e}.note{margin:14px 4px;color:var(--muted);font-size:12px}@media(max-width:900px){.shell{display:block;padding:12px}.side{position:static;height:auto;margin-bottom:12px}.nav{display:flex;overflow:auto}.nav a{white-space:nowrap}.cards{grid-template-columns:1fr 1fr}.decision-grid{grid-template-columns:1fr}.hero{padding:22px}h1{font-size:23px}}
  </style>
</head>
<body><div class="shell">
<aside class="side"><div class="brand">关键词作战报告</div><div class="asin">${escapeHtml(task.asin)}</div><nav class="nav"><a href="#overview" class="active" data-view="overview"><b>01</b>决策总览</a><a href="#actions" data-view="actions"><b>02</b>动作摘要</a><a href="#battle" data-view="battle"><b>03</b>关键词作战表</a><a href="#method" data-view="method"><b>04</b>数据口径</a></nav></aside>
<main>
  <section class="hero"><div class="eyebrow">KEYWORD INTELLIGENCE</div><h1>${escapeHtml(task.asin)} 关键词作战报告</h1><div class="sub">美国站 · SIF 前 ${rows.length} 个核心词 · 生成于 ${escapeHtml(generatedAt)}</div></section>
  <section class="report-view active" id="overview"><div class="cards">
    <div class="card"><span>核心关键词</span><strong>${rows.length}</strong></div>
    <div class="card"><span>优先动作词</span><strong>${actionCount}</strong></div>
    <div class="card"><span>样本广告花费</span><strong>$${formatNumber(totalSpend, 2)}</strong></div>
    <div class="card"><span>样本广告订单</span><strong>${formatNumber(totalOrders)}</strong></div>
  </div><div class="method-card"><h2>本轮运营结论</h2><p>共识别 ${rows.length} 个核心词，其中 ${controlCount} 个需要优先止损或控制成本，${scaleCount} 个建议放量，${observeCount} 个继续观察。</p></div></section>
  <section class="decision-grid report-view" id="actions"><div class="decision danger"><span>优先止损 / 控成本</span><strong>${controlCount}</strong></div><div class="decision good"><span>建议放量</span><strong>${scaleCount}</strong></div><div class="decision"><span>继续观察</span><strong>${observeCount}</strong></div></section>
  <section class="panel report-view" id="battle"><div class="panel-head"><div><strong>关键词作战总表</strong><div class="sub" style="color:var(--muted)">按止损 → 控成本 → 放量 → 新机会 → 观察排序</div></div><div class="tools"><input id="keywordSearch" placeholder="搜索关键词"><select id="actionFilter"><option value="">全部动作</option><option>止损</option><option>控制成本</option><option>放量</option><option>新机会</option><option>观察</option></select><button id="exportCsv" type="button">导出执行清单</button></div></div>
    <div class="table-wrap"><table><thead><tr>
      <th class="sticky">关键词</th><th>月搜索量</th><th>需求趋势</th><th>自然位</th><th>广告位</th><th>头部竞品</th><th>头部最高渠道份额</th><th>曝光</th><th>点击</th><th>花费</th><th>订单</th><th>ACOS</th><th>动作</th><th>打法建议</th>
    </tr></thead><tbody>${bodyRows}</tbody></table></div>
  </section>
  <section class="method-card report-view" id="method"><h2>数据口径与说明</h2><ul><li>广告指标来自本次上传的搜索词报表，并按搜索词聚合。</li><li>关键词需求、趋势、自然位和竞品信息来自 SIF 小样数据。</li><li>“—”表示源报表没有对应字段，或该关键词本期没有投放数据。</li><li>SIF 单个关键词调用失败时保留“数据缺失”，不会阻断整个任务。</li><li>动作建议用于运营复核，不会自动修改亚马逊广告活动。</li></ul></section>
  <p class="note">报告生成时间：${escapeHtml(generatedAt)}</p>
</main></div><script>
const q=document.getElementById('keywordSearch'),f=document.getElementById('actionFilter'),trs=[...document.querySelectorAll('tbody tr')];
const navs=[...document.querySelectorAll('.nav a')],views=[...document.querySelectorAll('.report-view')];
function showView(name){navs.forEach(a=>a.classList.toggle('active',a.dataset.view===name));views.forEach(v=>v.classList.toggle('active',v.id===name));window.scrollTo({top:0,behavior:'smooth'})}
navs.forEach(a=>a.addEventListener('click',e=>{e.preventDefault();showView(a.dataset.view)}));
function filterRows(){const k=q.value.trim().toLowerCase(),a=f.value;trs.forEach(r=>r.hidden=!!((k&&!r.dataset.keyword.includes(k))||(a&&!r.dataset.action.includes(a))))}
q.addEventListener('input',filterRows);f.addEventListener('change',filterRows);
document.getElementById('exportCsv').addEventListener('click',()=>{
  const headers=[...document.querySelectorAll('thead th')].map(x=>x.innerText.trim());
  const visible=trs.filter(r=>!r.hidden);
  const csv=[headers,...visible.map(r=>[...r.cells].map(c=>c.innerText.trim()))]
    .map(row=>row.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(',')).join('\r\n');
  const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}),url=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=url;a.download='${escapeHtml(task.asin)}-广告执行清单.csv';document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);
});
</script></body></html>`;
}

async function markFailed(taskId, reason) {
  const safeReason = String(reason || '未知错误').slice(0, 1000);
  const { error } = await supabase
    .from('keyword_tasks')
    .update({ status: '失败', failure_reason: safeReason, report_url: null })
    .eq('id', taskId);
  if (error) console.error(`任务 ${taskId} 写入失败状态时出错：${error.message}`);
}

async function processTask(task) {
  console.log(`开始领取任务 ${task.id}，ASIN ${task.asin}`);
  const { data: claimedRows, error: claimError } = await supabase
    .from('keyword_tasks')
    .update({ status: '进行中', failure_reason: null, report_url: null })
    .eq('id', task.id)
    .eq('status', '待处理')
    .select('id');
  if (claimError) throw claimError;
  if (!claimedRows || claimedRows.length !== 1) {
    console.log(`任务 ${task.id} 已被其他 worker 领取，跳过。`);
    return;
  }

  try {
    console.log(`任务 ${task.id}：下载并解析广告报表 ${task.report_path}`);
    const input = await downloadAdvertisingReport(task);
    const adReport = parseAdvertisingWorkbook(input, task.report_path);
    const rows = await buildKeywordRows(task, adReport);
    const html = renderHtmlReport(task, rows, adReport);
    const objectKey = reportObjectKey(task);

    console.log(`任务 ${task.id}：上传 HTML 报告到 COS ${objectKey}`);
    await putObject({
      Bucket: process.env.COS_BUCKET,
      Region: process.env.COS_REGION,
      Key: objectKey,
      Body: Buffer.from(html, 'utf8'),
      ContentType: 'text/html; charset=utf-8'
    });
    const reportUrl = await getSignedObjectUrl(objectKey);
    const { error: completeError } = await supabase
      .from('keyword_tasks')
      .update({ status: '已完成', failure_reason: null, report_url: reportUrl })
      .eq('id', task.id)
      .eq('status', '进行中');
    if (completeError) throw completeError;

    if (String(process.env.DELETE_SOURCE_AFTER_SUCCESS).toLowerCase() === 'true') {
      const { error: removeError } = await supabase.storage
        .from(INBOX_BUCKET)
        .remove([task.report_path]);
      if (removeError) console.warn(`任务 ${task.id} 删除源报表失败：${removeError.message}`);
    }
    console.log(`任务 ${task.id} 已完成，真实 HTML 报告已上传。`);
  } catch (error) {
    console.error(`任务 ${task.id} 处理失败：${error.message}`);
    await markFailed(task.id, error.message);
  }
}

async function refreshCompletedReportUrls() {
  if (refreshingUrls) return;
  refreshingUrls = true;
  try {
    const { data, error } = await supabase
      .from('keyword_tasks')
      .select('id, user_id, report_url')
      .eq('status', '已完成')
      .order('updated_at', { ascending: false })
      .limit(100);
    if (error) throw error;
    for (const task of data || []) {
      // 阶段 6 的历史 CSV 仍使用自己的签名地址；不要把它误改成不存在的 HTML。
      if (!String(task.report_url || '').includes('.html')) continue;
      const reportUrl = await getSignedObjectUrl(reportObjectKey(task));
      const { error: updateError } = await supabase
        .from('keyword_tasks')
        .update({ report_url: reportUrl })
        .eq('id', task.id)
        .eq('status', '已完成');
      if (updateError) console.warn(`任务 ${task.id} 刷新报告链接失败：${updateError.message}`);
    }
    if (data?.length) console.log(`已刷新 ${data.length} 个已完成任务的报告签名链接。`);
  } catch (error) {
    console.error(`刷新报告链接失败：${error.message}`);
  } finally {
    refreshingUrls = false;
  }
}

async function pollTasks() {
  if (polling) return;
  polling = true;
  try {
    const { data, error } = await supabase
      .from('keyword_tasks')
      .select('id, user_id, asin, report_path, created_at')
      .eq('status', '待处理')
      .order('created_at', { ascending: true })
      .limit(1);
    if (error) throw error;
    if (!data || data.length === 0) {
      console.log(`[${new Date().toLocaleString('zh-CN', { timeZone: 'Asia/Shanghai' })}] 无新任务`);
      return;
    }
    await processTask(data[0]);
  } catch (error) {
    console.error(`轮询任务失败：${error.message}`);
  } finally {
    polling = false;
  }
}

console.log(`阶段 7 worker 已启动：每 30 秒检查任务，SIF 小样上限 ${KEYWORD_LIMIT} 个词。`);
pollTasks();
refreshCompletedReportUrls();
setInterval(pollTasks, POLL_INTERVAL_MS);
setInterval(refreshCompletedReportUrls, URL_REFRESH_INTERVAL_MS);
