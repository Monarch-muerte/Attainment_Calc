// Driver 1 of 2: a tiny embedded database with ZERO native dependencies.
// Data is stored as JSON in data/attainment.json, on THIS machine's disk only.
// Zero setup, but only reachable from the computer running the server —
// use the "mongo" driver (see mongo.js) if you need multiple computers/people
// to share the same live data. Selected by default, or with DB_DRIVER=json.
const fs = require("fs");
const path = require("path");

const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, "..", "..", "data");
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const DB_PATH = path.join(DATA_DIR, "attainment.json");

function load() {
  if (!fs.existsSync(DB_PATH)) {
    return { users: [], courses: [], nextUserId: 1, nextCourseId: 1 };
  }
  try {
    const parsed = JSON.parse(fs.readFileSync(DB_PATH, "utf8"));
    return {
      users: parsed.users || [],
      courses: parsed.courses || [],
      nextUserId: parsed.nextUserId || 1,
      nextCourseId: parsed.nextCourseId || 1,
    };
  } catch {
    return { users: [], courses: [], nextUserId: 1, nextCourseId: 1 };
  }
}

let state = load();

function persist() {
  // Atomic-ish write: write to a temp file then rename, to avoid corrupting
  // the database if the process is killed mid-write.
  const tmp = DB_PATH + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(state, null, 2));
  fs.renameSync(tmp, DB_PATH);
}

module.exports = {
  driverName: "json (local file)",

  /* ---------- Users ---------- */
  findUserByUsername(username) {
    return state.users.find((u) => u.username === username) || null;
  },
  findUserById(id) {
    return state.users.find((u) => u.id === Number(id)) || null;
  },
  insertUser({ username, displayName, passwordHash, role = "faculty" }) {
    const user = {
      id: state.nextUserId++,
      username,
      display_name: displayName,
      role,
      password_hash: passwordHash,
      created_at: new Date().toISOString(),
    };
    state.users.push(user);
    persist();
    return user;
  },
  updateUserPassword(id, passwordHash, displayName) {
    const u = state.users.find((u) => u.id === Number(id));
    if (!u) return null;
    u.password_hash = passwordHash;
    if (displayName) u.display_name = displayName;
    persist();
    return u;
  },

  /* ---------- Courses ---------- */
  listCoursesByUser(userId) {
    return state.courses
      .filter((c) => c.user_id === Number(userId))
      .sort((a, b) => (a.updated_at < b.updated_at ? 1 : -1));
  },
  findCourse(id, userId) {
    return state.courses.find((c) => c.id === Number(id) && c.user_id === Number(userId)) || null;
  },
  findCourseByCode(userId, code) {
    return state.courses.find((c) => c.user_id === Number(userId) && c.course_code === code) || null;
  },
  insertCourse({ userId, courseCode, courseName, data }) {
    const now = new Date().toISOString();
    const course = {
      id: state.nextCourseId++,
      user_id: Number(userId),
      course_code: courseCode,
      course_name: courseName || "",
      data_json: JSON.stringify(data),
      created_at: now,
      updated_at: now,
    };
    state.courses.push(course);
    persist();
    return course;
  },
  updateCourse(id, { courseCode, courseName, data }) {
    const c = state.courses.find((c) => c.id === Number(id));
    if (!c) return null;
    if (courseCode != null) c.course_code = courseCode;
    if (courseName != null) c.course_name = courseName;
    if (data != null) c.data_json = JSON.stringify(data);
    c.updated_at = new Date().toISOString();
    persist();
    return c;
  },
  deleteCourse(id) {
    const idx = state.courses.findIndex((c) => c.id === Number(id));
    if (idx === -1) return false;
    state.courses.splice(idx, 1);
    persist();
    return true;
  },

  /* ---------- Utility ---------- */
  _dumpPath: DB_PATH,
};
