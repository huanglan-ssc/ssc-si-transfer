const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { exec } = require('child_process');

const app = express();
const PORT = process.env.PORT || 3456;
const ADMIN_KEY = process.env.ADMIN_KEY || 'ssc2024';

// Vercel serverless 环境文件系统只读，使用 /tmp 目录存储数据
const IS_VERCEL = !!process.env.VERCEL;
const DATA_DIR = IS_VERCEL
  ? path.join('/tmp', 'ssc-si-transfer-data')
  : path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const CONFIG_FILE = path.join(DATA_DIR, 'config.json');
const RESPONSES_FILE = path.join(DATA_DIR, 'responses.json');
const EMPLOYEES_FILE = path.join(DATA_DIR, 'employees.json');

// 初始化数据文件（如果不存在则创建默认文件）
// 模板文件始终在项目 data/ 目录中
const PROJECT_DATA_DIR = path.join(__dirname, 'data');
const CONFIG_TEMPLATE = path.join(PROJECT_DATA_DIR, 'config-template.json');
if (!fs.existsSync(CONFIG_FILE)) {
  if (fs.existsSync(CONFIG_TEMPLATE)) {
    fs.copyFileSync(CONFIG_TEMPLATE, CONFIG_FILE);
    console.log('已从模板初始化 config.json');
  } else {
    fs.writeFileSync(CONFIG_FILE, '{}', 'utf-8');
    console.log('已创建空 config.json');
  }
}
if (!fs.existsSync(RESPONSES_FILE)) fs.writeFileSync(RESPONSES_FILE, '[]', 'utf-8');
if (!fs.existsSync(EMPLOYEES_FILE)) fs.writeFileSync(EMPLOYEES_FILE, '[]', 'utf-8');

app.use(cors());
app.use(bodyParser.json({ limit: '5mb' }));
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

function readJSON(filePath) {
  try {
    const data = fs.readFileSync(filePath, 'utf-8');
    return JSON.parse(data);
  } catch (e) {
    return null;
  }
}

function writeJSON(filePath, obj) {
  fs.writeFileSync(filePath, JSON.stringify(obj, null, 2), 'utf-8');
}

function checkAdmin(req, res, next) {
  const key = req.query.key || req.body.key;
  if (key !== ADMIN_KEY) {
    return res.status(403).json({ error: '无权访问，请检查管理密钥' });
  }
  next();
}

// 员工端：获取问卷配置
app.get('/api/config', (req, res) => {
  const config = readJSON(CONFIG_FILE);
  if (!config) return res.status(500).json({ error: '配置读取失败' });
  res.json(config);
});

// 员工端：根据姓名或工号查询员工信息
app.get('/api/lookup-employee', (req, res) => {
  const { name, employeeId } = req.query;
  const employees = readJSON(EMPLOYEES_FILE) || [];
  if (!name && !employeeId) {
    return res.status(400).json({ error: '请提供姓名或工号' });
  }
  const queryName = String(name || '').trim().toLowerCase();
  const queryId = String(employeeId || '').trim();
  const results = employees.filter(e => {
    const matchName = queryName && String(e.name || '').toLowerCase() === queryName;
    const matchId = queryId && String(e.employeeId || '') === queryId;
    return matchName || matchId;
  });
  res.json(results.slice(0, 5));
});

// 管理端：上传员工信息Excel
app.post('/api/upload-employees', checkAdmin, (req, res) => {
  const { data } = req.body;
  if (!data || !Array.isArray(data)) {
    return res.status(400).json({ error: '请提供员工数据数组' });
  }
  const employees = data.map(row => ({
    name: String(row.name || row.花名 || row.姓名 || '').trim(),
    employeeId: String(row.employeeId || row.工号 || '').trim(),
    location: String(row.location || row.缴纳地 || row['现社保/公积金缴纳地'] || '').trim(),
    dingtalkId: String(row.dingtalkId || row.钉钉ID || '').trim()
  })).filter(e => e.name && e.employeeId);
  writeJSON(EMPLOYEES_FILE, employees);
  res.json({ success: true, count: employees.length });
});

// 获取员工列表（公开接口，供员工页面查询使用）
app.get('/api/employees', (req, res) => {
  const employees = readJSON(EMPLOYEES_FILE) || [];
  res.json(employees);
});

// 保存员工名单（覆盖写入）
app.post('/api/employees', checkAdmin, (req, res) => {
  const { employees } = req.body;
  if (!employees || !Array.isArray(employees)) {
    return res.status(400).json({ error: '请提供员工数据数组' });
  }
  writeJSON(EMPLOYEES_FILE, employees);
  res.json({ success: true, count: employees.length });
});

// 催办：发送钉钉机器人消息给指定员工
app.post('/api/remind/:employeeId', checkAdmin, (req, res) => {
  const { employeeId } = req.params;
  const employees = readJSON(EMPLOYEES_FILE) || [];
  const employee = employees.find(e => String(e.employeeId) === String(employeeId));
  if (!employee) {
    return res.status(404).json({ error: '未找到该员工信息' });
  }
  if (!employee.dingtalkId) {
    return res.status(400).json({ error: '该员工未配置钉钉ID，无法发送催办消息' });
  }

  const config = readJSON(CONFIG_FILE) || {};
  const robotCode = config.dingtalk && config.dingtalk.robotCode;
  if (!robotCode) {
    return res.status(400).json({ error: '系统未配置钉钉机器人robotCode，请在config.json中添加dingtalk.robotCode配置' });
  }

  const text = `您好，${employee.name}，您尚未完成社保转移情况确认表的填写，请尽快完成提交。如有问题请联系SSC小管家。`;
  const cmd = `dws chat message send-by-bot --robot-code ${robotCode} --users ${employee.dingtalkId} --title "${'\u793e\u4fdd\u8f6c\u79fb\u63d0\u9192'}" --text "${text}"`;

  exec(cmd, (error, stdout, stderr) => {
    if (error) {
      return res.status(500).json({ error: '\u53d1\u9001\u5931\u8d25\uff1a' + (stderr || error.message) });
    }
    res.json({ success: true, message: '\u5df2\u53d1\u9001\u50ac\u529e\u6d88\u606f' });
  });
});

// 管理端：手动添加员工
app.post('/api/add-employee', checkAdmin, (req, res) => {
  const { name, employeeId, location, dingtalkId } = req.body;
  if (!name || !employeeId) {
    return res.status(400).json({ error: '姓名和工号为必填' });
  }
  const employees = readJSON(EMPLOYEES_FILE) || [];
  // 如果工号已存在，更新；否则添加
  const idx = employees.findIndex(e => e.employeeId === String(employeeId).trim());
  const newEmp = {
    name: String(name).trim(),
    employeeId: String(employeeId).trim(),
    location: String(location || '').trim(),
    dingtalkId: String(dingtalkId || '').trim()
  };
  if (idx >= 0) {
    employees[idx] = newEmp;
  } else {
    employees.push(newEmp);
  }
  writeJSON(EMPLOYEES_FILE, employees);
  res.json({ success: true });
});

// 管理端：删除员工
app.post('/api/delete-employee', checkAdmin, (req, res) => {
  const { employeeId } = req.body;
  if (!employeeId) return res.status(400).json({ error: '缺少工号' });
  let employees = readJSON(EMPLOYEES_FILE) || [];
  employees = employees.filter(e => e.employeeId !== String(employeeId).trim());
  writeJSON(EMPLOYEES_FILE, employees);
  res.json({ success: true });
});

// 员工端：提交问卷
app.post('/api/submit', (req, res) => {
  const { name, employeeId, location, selectedCases, finalAnswer, customTexts } = req.body;
  if (!name || !employeeId || !finalAnswer) {
    return res.status(400).json({ error: '姓名、工号和最终确认项为必填' });
  }

  const config = readJSON(CONFIG_FILE);
  if (!config) return res.status(500).json({ error: '配置读取失败' });

  // Build flat map of all options for lookup
  const optionMap = new Map();
  const categoryMap = new Map();
  for (const cat of config.specialCases || []) {
    categoryMap.set(cat.id, cat.category);
    for (const opt of cat.options || []) {
      optionMap.set(opt.id, { ...opt, category: cat.category, categoryId: cat.id });
    }
  }

  const selectedDetails = (selectedCases || []).map(scId => {
    const opt = optionMap.get(scId);
    const customText = (customTexts && customTexts[scId]) ? String(customTexts[scId]).trim() : '';
    if (opt) {
      return {
        id: opt.id,
        label: opt.label,
        impact: opt.impact,
        category: opt.category,
        categoryId: opt.categoryId,
        allowCustom: opt.allowCustom || false,
        customText
      };
    }
    return { id: scId, label: scId, impact: '', category: '', categoryId: '', customText };
  });

  const finalOpt = (config.finalQuestion?.options || []).find(o => o.value === finalAnswer);

  const record = {
    id: Date.now().toString(36) + Math.random().toString(36).substr(2, 5),
    name: String(name).trim(),
    employeeId: String(employeeId).trim(),
    location: String(location || '').trim(),
    selectedCases: selectedCases || [],
    selectedDetails,
    finalAnswer,
    finalLabel: finalOpt ? finalOpt.label : finalAnswer,
    submitTime: new Date().toISOString()
  };

  const responses = readJSON(RESPONSES_FILE) || [];
  // 计算该员工的提交次数
  const sameEmployeeResponses = responses.filter(r => r.employeeId === record.employeeId);
  const submitCount = sameEmployeeResponses.length + 1;
  record.submitCount = submitCount;
  responses.push(record);
  writeJSON(RESPONSES_FILE, responses);

  res.json({ success: true, record });
});

// 管理端：获取所有提交
app.get('/api/responses', checkAdmin, (req, res) => {
  const responses = readJSON(RESPONSES_FILE) || [];
  res.json(responses);
});

// 管理端：删除单条提交
app.post('/api/delete-response', checkAdmin, (req, res) => {
  const { id } = req.body;
  if (!id) return res.status(400).json({ error: '缺少记录ID' });
  let responses = readJSON(RESPONSES_FILE) || [];
  responses = responses.filter(r => r.id !== id);
  writeJSON(RESPONSES_FILE, responses);
  res.json({ success: true });
});

// 管理端：获取完整配置
app.get('/api/config-admin', checkAdmin, (req, res) => {
  const config = readJSON(CONFIG_FILE);
  if (!config) return res.status(500).json({ error: '配置读取失败' });
  res.json(config);
});

// 管理端：更新配置
app.post('/api/config-admin', checkAdmin, (req, res) => {
  const { title, description, specialCases, finalQuestion, footerNote, agreeInfo, confirmMessages, noneSelectedHints, noneOption } = req.body;
  const config = {
    title: title || '社保转移情况确认',
    description: description || '',
    specialCases: Array.isArray(specialCases) ? specialCases : [],
    finalQuestion: finalQuestion || {},
    footerNote: footerNote || '',
    agreeInfo: agreeInfo || {},
    confirmMessages: confirmMessages || {},
    noneSelectedHints: noneSelectedHints || {},
    noneOption: noneOption || {}
  };
  writeJSON(CONFIG_FILE, config);
  res.json({ success: true });
});

// 管理端：导出Excel
app.get('/api/export', checkAdmin, (req, res) => {
  const responses = readJSON(RESPONSES_FILE) || [];

  const rows = responses.map(r => {
    const caseCategories = [...new Set(r.selectedDetails.map(d => d.category).filter(Boolean))].join('、') || '无';
    const caseLabels = r.selectedDetails.map(d => {
      let label = d.label;
      if (d.customText) label += `（${d.customText}）`;
      return label;
    }).join('；') || '无';
    const impacts = r.selectedDetails.map(d => d.impact).join('\n') || '无';
    return {
      '序号': r.id,
      '姓名': r.name,
      '工号': r.employeeId,
      '现社保/公积金缴纳地': r.location || '',
      '特殊情形大类': caseCategories,
      '具体情况': caseLabels,
      '对应影响说明': impacts,
      '是否同意转移社保': r.finalLabel,
      '提交时间': r.submitTime,
      '填写次数': r.submitCount || 1
    };
  });

  const ws = xlsx.utils.json_to_sheet(rows);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, '社保转移确认汇总');

  // 调整列宽
  const colWidths = [
    { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 20 },
    { wch: 25 }, { wch: 40 }, { wch: 60 }, { wch: 18 }, { wch: 20 }, { wch: 10 }
  ];
  ws['!cols'] = colWidths;

  const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const timestamp = new Date().toISOString().replace(/[:T]/g, '-').split('.')[0];
  const filename = `社保转移确认汇总_${timestamp}.xlsx`;

  res.setHeader('Content-Disposition', 'attachment; filename=' + encodeURIComponent(filename));
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.send(buf);
});

// 首页重定向到员工问卷
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'employee.html'));
});

app.listen(PORT, () => {
  console.log(`SSC社保转移沟通系统已启动`);
  console.log(`员工问卷地址: http://localhost:${PORT}`);
  console.log(`管理后台地址: http://localhost:${PORT}/admin.html?key=${ADMIN_KEY}`);
  console.log(`按 Ctrl+C 停止服务`);
});
