const express = require('express');
const path = require('path');

const app = express();
const PORT = Number(process.env.PORT) || 3000;
const HOST = '0.0.0.0';

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(path.join(__dirname), {
  index: 'index.html'
}));

app.get('/api/proxy-export', async (req, res) => {
  const target = String(req.query.url || '').trim();
  if (!/^https?:\/\//i.test(target)) {
    return res.status(400).json({ error: 'Invalid URL' });
  }
  try {
    const upstream = await fetch(target, { redirect: 'manual' });
    if (upstream.status >= 300 && upstream.status < 400) {
      const loc = upstream.headers.get('location') || '';
      return res.status(401).json({
        error: 'AUTH_REQUIRED',
        message: 'แอปเก่าต้องล็อกอินบัญชี Google ก่อนเข้าถึง (' + loc + ')'
      });
    }
    const text = await upstream.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch {
      return res.status(401).json({
        error: 'AUTH_REQUIRED',
        message: 'ลิงก์ปลายทางตอบกลับเป็นหน้าล็อกอิน HTML ไม่ใช่ข้อมูล JSON'
      });
    }
    return res.json(data);
  } catch (err) {
    return res.status(502).json({ error: err && err.message ? err.message : 'Fetch failed' });
  }
});

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.listen(PORT, HOST, () => {
  console.log(`Server listening on http://${HOST}:${PORT}`);
});
