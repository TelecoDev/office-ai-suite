const express = require("express");
const path = require("path");
const fs = require("fs");
const https = require("https");
const { createProxyMiddleware } = require("http-proxy-middleware");

const app = express();

// 🔐 Percorso ai certificati dev di Office
const certPath = path.join(
  process.env.HOME || process.env.USERPROFILE,
  ".office-addin-dev-certs",
  "localhost.crt"
);
const keyPath = path.join(
  process.env.HOME || process.env.USERPROFILE,
  ".office-addin-dev-certs",
  "localhost.key"
);

const options = {
  cert: fs.readFileSync(certPath),
  key: fs.readFileSync(keyPath),
};

// 🌐 Percorso ai file buildati
const distPath = path.join(__dirname, "..", "dist");

// Serve la build
app.use(express.static(distPath));

// 🔥 PROXY HTTPS → HTTP per Ollama
app.use(
  "/api/ollama",
  createProxyMiddleware({
    target: "http://172.30.5.220:11434", // PC A OLLAMA
    changeOrigin: true,
    secure: false,
    pathRewrite: {
      "^/api/ollama": "",
    },
  })
);

// Avvia server
https.createServer(options, app).listen(3000, () => {
  console.log("OfficeAI SERVE → https://172.30.5.220:3000");
});
