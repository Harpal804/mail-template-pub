const express = require("express");
const app = express();
const path = require("path");

app.use(express.static(path.join(__dirname, "dist")));

app.get("/templates.json", (req, res) => {
  res.sendFile(path.join(__dirname, "dist/templates.json"));
});

app.listen(3001, () => console.log("JSON Server running on http://localhost:3001"));