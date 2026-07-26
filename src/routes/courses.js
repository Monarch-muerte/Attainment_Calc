const express = require("express");
const db = require("../db");
const { requireAuth } = require("../middleware/auth");

const router = express.Router();
router.use(requireAuth);

function rowToSummary(row) {
  return { id: row.id, courseCode: row.course_code, courseName: row.course_name, updatedAt: row.updated_at, createdAt: row.created_at };
}

// List all saved courses for the logged-in user
router.get("/", async (req, res) => {
  try {
    const rows = await db.listCoursesByUser(req.user.id);
    res.json({ courses: rows.map(rowToSummary) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Could not reach the database." });
  }
});

// Get one course's full data
router.get("/:id", async (req, res) => {
  try {
    const row = await db.findCourse(req.params.id, req.user.id);
    if (!row) return res.status(404).json({ error: "Course not found." });
    let data;
    try { data = JSON.parse(row.data_json); } catch { data = null; }
    res.json({ ...rowToSummary(row), data });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Could not reach the database." });
  }
});

// Create a new saved course
router.post("/", async (req, res) => {
  try {
    const { courseCode, courseName, data } = req.body || {};
    if (!courseCode || !data) return res.status(400).json({ error: "courseCode and data are required." });
    const code = String(courseCode).trim();

    if (await db.findCourseByCode(req.user.id, code)) {
      return res.status(409).json({ error: "You already have a saved course with that course code. Use update instead." });
    }
    const row = await db.insertCourse({ userId: req.user.id, courseCode: code, courseName: String(courseName || "").trim(), data });
    res.status(201).json(rowToSummary(row));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Could not save course." });
  }
});

// Update an existing saved course (full replace of data)
router.put("/:id", async (req, res) => {
  try {
    const { courseCode, courseName, data } = req.body || {};
    const existing = await db.findCourse(req.params.id, req.user.id);
    if (!existing) return res.status(404).json({ error: "Course not found." });
    const row = await db.updateCourse(existing.id, {
      courseCode: courseCode ? String(courseCode).trim() : undefined,
      courseName: courseName != null ? String(courseName).trim() : undefined,
      data,
    });
    res.json(rowToSummary(row));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Could not update course." });
  }
});

// Upsert by course code — convenient single "Save" action from the client
router.post("/upsert", async (req, res) => {
  try {
    const { courseCode, courseName, data } = req.body || {};
    if (!courseCode || !data) return res.status(400).json({ error: "courseCode and data are required." });
    const code = String(courseCode).trim();
    const existing = await db.findCourseByCode(req.user.id, code);
    if (existing) {
      const row = await db.updateCourse(existing.id, { courseName: String(courseName || existing.course_name || "").trim(), data });
      return res.json(rowToSummary(row));
    }
    const row = await db.insertCourse({ userId: req.user.id, courseCode: code, courseName: String(courseName || "").trim(), data });
    res.status(201).json(rowToSummary(row));
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Could not save course." });
  }
});

router.delete("/:id", async (req, res) => {
  try {
    const existing = await db.findCourse(req.params.id, req.user.id);
    if (!existing) return res.status(404).json({ error: "Course not found." });
    await db.deleteCourse(existing.id);
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Could not delete course." });
  }
});

module.exports = router;
