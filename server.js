const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware parse JSON
app.use(express.json());

// Serve static files dari folder saat ini
app.use(express.static(__dirname));

// Health check endpoint untuk Railway
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'ok', message: 'Server Laporan Keuangan aktif!' });
});

// Fallback untuk route apapun mengarah ke index.html
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Bind ke 0.0.0.0 agar bisa diakses publik oleh Railway
app.listen(PORT, '0.0.0.0', () => {
  console.log(`===============================================`);
  console.log(` Server Laporan Keuangan berjalan!`);
  console.log(` Akses lokal  : http://localhost:${PORT}`);
  console.log(` Port Railway : ${PORT}`);
  console.log(`===============================================`);
});
