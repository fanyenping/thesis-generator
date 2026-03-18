const express = require('express');
const path = require('path');
const fs = require('fs');
const yaml = require('js-yaml');
const { translate } = require('@vitalets/google-translate-api');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// 讀取目前設定
app.get('/api/config', (req, res) => {
  try {
    const cfg = yaml.load(fs.readFileSync(path.join(__dirname, 'thesis_config.yaml'), 'utf8'));
    res.json({ ok: true, config: cfg });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

// 翻譯 API（中文 → 英文）
app.post('/api/translate', async (req, res) => {
  try {
    const { text } = req.body;
    if (!text || !text.trim()) return res.json({ ok: true, result: '' });
    const result = await translate(text, { from: 'zh-TW', to: 'en' });
    res.json({ ok: true, result: result.text });
  } catch (e) {
    console.error('翻譯失敗：', e.message);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// 儲存設定並產生論文
app.post('/api/generate', async (req, res) => {
  try {
    const cfg = req.body;
    const yamlStr = yaml.dump(cfg, { lineWidth: 120 });
    fs.writeFileSync(path.join(__dirname, 'thesis_config.yaml'), yamlStr, 'utf8');

    // 動態執行 generate.js 邏輯
    const { buildDoc } = require('./generate-api');
    const buffer = await buildDoc(cfg);
    const filename = encodeURIComponent(cfg.output_filename || '碩士論文.docx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${filename}`);
    res.send(buffer);
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`\n✅ 台灣論文框架系統已啟動`);
  console.log(`   http://localhost:${PORT}`);
});
