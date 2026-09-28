const express = require('express');
const path = require('path');

const app = express();

// Middleware parse JSON
app.use(express.json());

// Serve static files dari folder aplikasi
app.use(express.static(__dirname));

// Health check endpoint untuk Railway
app.get('/health', (req, res) => {
  res.status(200).json({ status: 'ok', message: 'Server Laporan Keuangan aktif!' });
});

// Fallback untuk route apapun mengarah ke index.html
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Port configuration:
// Railway secara default menggunakan process.env.PORT, dan di dashboard Anda tertera Port 7070.
// Kita dengarkan di kedua port (7070 & 3000 & process.env.PORT) agar 100% selalu terhubung!
const envPort = process.env.PORT ? parseInt(process.env.PORT, 10) : null;
const targetPorts = new Set([7070, 3000]);
if (envPort) targetPorts.add(envPort);

targetPorts.forEach(port => {
  try {
    app.listen(port, '0.0.0.0', () => {
      console.log(`[OK] Server aktif mendengarkan di http://0.0.0.0:${port}`);
    });
  } catch (err) {
    console.warn(`[WARN] Port ${port} tidak dapat dibuka: ${err.message}`);
  }
});
