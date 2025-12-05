const express = require("express");
const path = require("path");
const app = express();

const PORT = 3000;

// Serve i file della build
app.use(express.static(path.join(__dirname, "../dist")));

app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "../dist/taskpane.html"));
});

app.listen(PORT, "0.0.0.0", () => {
  console.log(`OfficeAI Server LIVE → http://172.30.5.220:${PORT}`);
});
