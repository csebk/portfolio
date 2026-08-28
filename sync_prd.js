#!/usr/bin/env node
/**
 * V2.0.6 双向同步脚本：prd_data.json ↔ PRD_V2.0.md
 *
 * 用法：
 *   node sync_prd.js check    # 检查 JSON 与 MD 表格是否一致（CI 友好）
 *   node sync_prd.js json2md  # 从 JSON 同步到 MD（修改 MD 表格部分）
 *   node sync_prd.js md2json  # 从 MD 同步到 JSON（修改 JSON 数据部分）
 *   node sync_prd.js demo-default # 从 demo 抽 5 患者 + 5 随访 + 9 类既往史到 JSON
 *   node sync_prd.js rules-by-page # 从 PRD §X.7 6 字段汇总抽取按页分组的详细规则
 *
 * 双向追溯原则：
 *   - JSON 是 single source of truth
 *   - MD 中的"V2.0.4 demo 集成"表格由 JSON 生成
 *   - 改 JSON 跑一次 json2md → MD 同步
 *   - 改 MD 跑一次 md2json → JSON 同步（但 MD 表格会被 json2md 覆盖，仅适合"手写扩展"场景）
 */

const fs = require('fs');
const path = require('path');

const JSON_PATH = path.join(__dirname, 'demo', 'js', 'prd_data.json');
const MD_PATH = path.join(__dirname, 'AI辅助问诊_PRD优化版_V2.0.md');
const DEMO_HTML_PATH = path.join(__dirname, 'demo', 'index.html');

function loadJSON() {
  return JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
}

function loadMD() {
  return fs.readFileSync(MD_PATH, 'utf8');
}

function saveMD(content) {
  fs.writeFileSync(MD_PATH, content);
  console.log('✅ MD updated:', MD_PATH);
}

function saveJSON(data) {
  fs.writeFileSync(JSON_PATH, JSON.stringify(data, null, 2));
  console.log('✅ JSON updated:', JSON_PATH);
}

function status() {
  const d = loadJSON();
  const totalRules = d.rules.reduce((s, m) => s + m.rules.length, 0);
  const totalStatuses = Object.keys(d.status_meta).length;
  const totalDiseases = d.disease_library.length;
  const totalTips = d.tips.length;
  console.log(`\n📊 V2.0.7 PRD 数据状态`);
  console.log(`  版本:       ${d.version}`);
  console.log(`  更新时间:   ${d.last_updated}`);
  console.log(`  规则模块:   ${d.rules.length} 个（${totalRules} 条规则）`);
  console.log(`  状态码:     ${totalStatuses} 种`);
  console.log(`  疾病库:     ${totalDiseases} 个（按 ${new Set(d.disease_library.map(x => x.cat)).size} 大类）`);
  console.log(`  Tip 提示:   ${totalTips} 条`);
  if (d.history_8) console.log(`  既往史:     ${Object.keys(d.history_8).length} 类子字段`);
  if (d.symptoms_template) console.log(`  症状模板:   ${d.symptoms_template.length} 个`);
  if (d.signs_template) console.log(`  体征模板:   ${d.signs_template.length} 个`);
  if (d.test_special_note_templates) console.log(`  检查备注模板: ${d.test_special_note_templates.length} 条`);
}

function buildRulesTable(d) {
  let md = '\n### 14 个核心交互提示（自动同步自 prd_data.json）\n\n';
  let counter = 0;
  for (const m of d.rules) {
    md += `#### ${m.module}（${m.rules.length} 个）\n\n`;
    md += '| # | 交互 | 提示文案 | PRD 章节 |\n';
    md += '|---|---|---|---|\n';
    for (const r of m.rules) {
      counter++;
      md += `| ${counter} | ${r.title} | ${r.brief} | ${r.section} |\n`;
    }
    md += '\n';
  }
  return md;
}

function json2md() {
  const d = loadJSON();
  let md = loadMD();
  const table = buildRulesTable(d);

  // 替换"V2.0.4 demo 集成"章节中的"14 个核心交互提示"小节
  const startMarker = '### 14 个核心交互提示';
  const endMarker = '### 后续升级';
  const startIdx = md.indexOf(startMarker);
  const endIdx = md.indexOf(endMarker);

  if (startIdx === -1 || endIdx === -1) {
    console.error('❌ MD 标记未找到:', startMarker, '/', endMarker);
    process.exit(1);
  }

  const newMD = md.slice(0, startIdx) + table.trim() + '\n\n' + md.slice(endIdx);

  // 更新顶部版本号
  const finalMD = newMD.replace(
    /<!-- PRD_DATA_VERSION -->V[\d.]+<!-- \/PRD_DATA_VERSION -->/g,
    `<!-- PRD_DATA_VERSION -->${d.version}<!-- /PRD_DATA_VERSION -->`
  );

  saveMD(finalMD);
  console.log(`📊 已同步 ${d.rules.reduce((s, m) => s + m.rules.length, 0)} 条规则到 MD`);
}

function check() {
  const d = loadJSON();
  const md = loadMD();
  let issues = 0;

  // 1. 检查 tips 数量 == 14
  if (d.tips.length !== 14) {
    console.warn(`⚠️ tips 数量 = ${d.tips.length}（期望 14）`);
    issues++;
  }

  // 2. 检查每条规则都有对应 tip
  const tipIds = new Set(d.tips.map(t => t.id));
  for (const m of d.rules) {
    for (const r of m.rules) {
      if (!tipIds.has(r.id)) {
        console.warn(`⚠️ 规则 ${r.id} (${r.title}) 缺少对应 tip`);
        issues++;
      }
    }
  }

  // 3. 检查状态码 12 个
  if (Object.keys(d.status_meta).length !== 12) {
    console.warn(`⚠️ 状态码数量 = ${Object.keys(d.status_meta).length}（期望 12）`);
    issues++;
  }

  // 4. 检查 MD 中是否含每条规则的"#"行号（json2md 生成的格式：| 1 | xxx |...）
  let missing = 0;
  for (const m of d.rules) {
    for (const r of m.rules) {
      // 找 "| {编号} | {title} |" 模式
      const re = new RegExp(`\\|\\s*\\d+\\s*\\|\\s*${r.title.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*\\|`);
      if (!re.test(md)) {
        console.warn(`⚠️ MD 中未找到规则 ${r.id} (${r.title})`);
        missing++;
      }
    }
  }
  if (missing > 0) issues += missing;

  if (issues === 0) {
    console.log('✅ check OK: JSON ↔ MD 同步一致');
    process.exit(0);
  } else {
    console.error(`❌ check FAIL: ${issues} 个不一致`);
    process.exit(1);
  }
}

// ========== V2.1.0 demo-default：从 demo 源码抽 MOCK_PATIENTS + 9 类既往史默认值 ==========
// 用 Node vm 模块把 demo 源码当脚本跑 → 安全拿到 JS 变量值
const vm = require('vm');

function extractFromDemo(demoCode) {
  // 1. 找 const MOCK_XXX = [...] 块，每块单独 eval 拿值
  const result = {};
  // 把 demo 源码里的 const MOCK_PATIENTS 改成 var MOCK_PATIENTS（vm 上下文能导出）
  // 同时把箭头函数/JSX 跳过：只抽 const NAME = [...] 模式
  const arrayRegex = /const (MOCK_\w+) = (\[[\s\S]*?\n    \]);/g;
  let m;
  while ((m = arrayRegex.exec(demoCode)) !== null) {
    const name = m[1];
    // 在 vm 中执行这一段取变量
    try {
      const ctx = {};
      vm.createContext(ctx);
      vm.runInContext(`${name} = ${m[2]};`, ctx);
      result[name] = ctx[name];
    } catch (e) {
      console.log(`  ⚠️ ${name} 解析失败:`, e.message.slice(0, 80));
    }
  }
  return result;
}

function demoDefault() {
  const d = loadJSON();
  const demo = fs.readFileSync(DEMO_HTML_PATH, 'utf8');

  // 1. 用 vm 抽所有 MOCK_ 数组
  const mocks = extractFromDemo(demo);
  const patients = mocks.MOCK_PATIENTS || [];
  const followups = mocks.MOCK_FOLLOWUP_PATIENTS || [];
  console.log(`  ✓ 抽到 MOCK_PATIENTS: ${patients.length} 条`);
  console.log(`  ✓ 抽到 MOCK_FOLLOWUP_PATIENTS: ${followups.length} 条`);

  // 2. 抽 9 类既往史默认值（扫 defaultValue="..." 找既往史块）
  const historyDefaults = {};
  const labelMap = {
    '🤧 过敏史': 'allergy',
    '💊 既往疾病史': 'past_disease',
    '🔪 手术外伤史': 'surgery',
    '🩸 输血史': 'transfusion',
    '🦠 传染病史': 'infectious',
    '💉 预防接种史': 'vaccination',
    '🚬 个人史·烟酒': 'smoke_alcohol',
    '🛡️ 职业暴露': 'occupation',
    '👨‍👩‍👧 婚育史': 'marriage',
    '👪 家族史': 'family',
  };
  for (const [label, key] of Object.entries(labelMap)) {
    const idx = demo.indexOf(`>${label}<`);
    if (idx < 0) continue;
    const block = demo.substring(idx, idx + 800);
    const dvMatch = block.match(/defaultValue="([^"]*)"/);
    if (dvMatch) {
      historyDefaults[key] = dvMatch[1];
    }
  }
  console.log(`  ✓ 抽到 9 类既往史默认值: ${Object.keys(historyDefaults).length} 个`);

  // 3. 写到 JSON
  d.demo_data = {
    patients,
    followups,
    history_defaults: historyDefaults,
  };
  d.last_updated = new Date().toISOString().slice(0, 10);
  saveJSON(d);
  console.log(`\n✅ demo_default 完成: patients=${patients.length}, followups=${followups.length}, history=${Object.keys(historyDefaults).length}`);
}

// ========== 🆕 V2.5.0 rules-by-page：解析 PRD §X.7 6 字段汇总表格，按页分组 ==========
function rulesByPage() {
  const md = loadMD();
  const json = loadJSON();

  // 页面映射：模块 §X → 页面 key
  const PAGE_MAP = {
    '1': 'P-2',   // §1 预诊
    '2': 'P0',    // §2 档案
    '3': 'P1',    // §3 录音
    '4': 'P3',    // §4 分流
    '5': 'P3',    // §5 检查
    '6': 'P3',    // §6 AI 诊断
    '7': 'P3',    // §7 病历
    '8': 'P4',    // §8 处方
    '9': 'P5',    // §9 随访
    '10': 'P5',   // §10 AI 风险
    '11': 'P5',   // §11 多租户
    '12': 'P0.5', // §12 危急值
  };

  // 字段映射：模块 §X → 默认 field 名
  const FIELD_MAP = {
    '1': 'pre_diagnosis_start',
    '2': 'register',
    '3': 'recording',
    '4': 'triage',
    '5': 'test_ordered',
    '6': 'ai_trigger',
    '7': 'chief_complaint',
    '8': 'prescription',
    '9': 'followup_config',
    '10': 'ai_risk',
    '11': 'tenant',
    '12': 'critical_value',
  };

  const rulesByPage = { _comment: '按页面分组的详细规则集（V2.5.0 改造：从 PRD §X.7 6 字段汇总抽取）' };

  // 解析 §X.7 6 字段汇总表格
  // 匹配模式：#### X.7 🆕 V2.0 补全：6 字段汇总... + 后续 6 行表格
  const moduleRegex = /####\s+(\d+)\.7\s+🆕\s+V2\.0\s+补全：6\s+字段汇总[\s\S]*?(?=\n####|\n##|$)/g;
  let match;
  while ((match = moduleRegex.exec(md)) !== null) {
    const moduleNum = match[1];
    const page = PAGE_MAP[moduleNum];
    const field = FIELD_MAP[moduleNum];
    if (!page || !field) continue;

    const block = match[0];
    if (!rulesByPage[page]) rulesByPage[page] = [];

    // 抽取 6 字段表格行（| **触发条件** | ... |）
    const fieldNames = ['触发条件', '结束条件', '状态变化', '权限范围', '异常处理', '验收标准'];
    fieldNames.forEach((fn, i) => {
      const rowRegex = new RegExp(`\\*\\*${fn}\\*\\*\\s*\\|\\s*([^\\n|]+(?:\\|[^\\n]+)?)`);
      const rowMatch = block.match(rowRegex);
      const detail = rowMatch ? rowMatch[1].trim().replace(/\s*\|\s*$/, '') : '';
      if (detail) {
        rulesByPage[page].push({
          id: `${page}-${moduleNum}-${fn.replace('条件', '').replace('变化', '-state').replace('权限', '-perm').replace('异常', '-exc').replace('验收', '-ac')}`,
          field: field,
          rule: fn,
          detail: detail
        });
      }
    });
  }

  const total = Object.entries(rulesByPage).filter(([k]) => k !== '_comment').reduce((s, [_, v]) => s + v.length, 0);
  json.rules_by_page = rulesByPage;
  json.version = 'V2.5.0';
  json.last_updated = new Date().toISOString().slice(0, 10);
  saveJSON(json);
  console.log(`\n✅ rules-by-page 完成: ${total} 条规则 / ${Object.keys(rulesByPage).length - 1} 页面`);
  for (const [k, v] of Object.entries(rulesByPage)) {
    if (k !== '_comment') console.log(`   ${k}: ${v.length} 条`);
  }
}


const cmd = process.argv[2] || 'status';
switch (cmd) {
  case 'status':       status(); break;
  case 'json2md':      json2md(); break;
  case 'check':        check(); break;
  case 'demo-default': demoDefault(); break;
  case 'rules-by-page': rulesByPage(); break;
  default:
    console.error('Unknown command:', cmd);
    console.error('Usage: node sync_prd.js [status|json2md|check|demo-default|rules-by-page]');
    process.exit(1);
}
