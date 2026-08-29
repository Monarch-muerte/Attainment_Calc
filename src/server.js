require("dotenv").config();
const path = require("path");
const express = require("express");
const cors = require("cors");
const morgan = require("morgan");

const authRoutes = require("./routes/auth");
const courseRoutes = require("./routes/courses");
const db = require("./db");

const app = express();
const PORT = process.env.PORT || 4000;

app.use(cors());
app.use(express.json({ limit: "5mb" }));
app.use(morgan("dev"));

app.get("/api/health", (req, res) => res.json({ ok: true, db: db.driverName, time: new Date().toISOString() }));
app.use("/api/auth", authRoutes);
app.use("/api/courses", courseRoutes);

// Serve the built frontend
const publicDir = path.join(__dirname, "..", "public");
app.use(express.static(publicDir));
app.get("*", (req, res, next) => {
  if (req.path.startsWith("/api/")) return next();
  res.sendFile(path.join(publicDir, "index.html"));
});

// Central error handler
app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).json({ error: "Server error." });
});

app.listen(PORT, () => {
  console.log(`CO/PO Attainment System running at http://localhost:${PORT}`);
  console.log(`Database driver: ${db.driverName}`);
});
