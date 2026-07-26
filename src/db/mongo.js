// Driver 2 of 2: MongoDB (works great with a free MongoDB Atlas cluster).
// Unlike the json driver, this database lives on the network — so your
// server can run on any computer (your laptop, a classmate's laptop, a free
// host like Render/Railway) and as long as they all use the same
// MONGODB_URI, they all read and write the exact same shared data.
//
// Enable it by setting in .env:
//   DB_DRIVER=mongo
//   MONGODB_URI=mongodb+srv://<user>:<password>@<cluster>.mongodb.net/
//   MONGODB_DB=attainment          (optional, defaults to "attainment")
//
// Get a free connection string in ~2 minutes at https://www.mongodb.com/cloud/atlas
// (Free tier "M0" cluster -> Database Access: create a user -> Network Access:
// allow access from anywhere (0.0.0.0/0) for a class project -> Connect -> Drivers -> copy the URI.)
const { MongoClient } = require("mongodb");

const uri = process.env.MONGODB_URI;
const dbName = process.env.MONGODB_DB || "attainment";

let client;
let dbPromise;

function getDb() {
  if (!uri) {
    throw new Error(
      "DB_DRIVER=mongo is set but MONGODB_URI is missing from .env. " +
      "See src/db/mongo.js for how to get a free connection string."
    );
  }
  if (!dbPromise) {
    client = new MongoClient(uri);
    dbPromise = client.connect().then((c) => c.db(dbName));
  }
  return dbPromise;
}

async function nextSeq(db, counterName) {
  const result = await db.collection("counters").findOneAndUpdate(
    { _id: counterName },
    { $inc: { seq: 1 } },
    { upsert: true, returnDocument: "after" }
  );
  // Different driver versions return either the doc directly or { value: doc }.
  const doc = result && result.value ? result.value : result;
  return doc.seq;
}

module.exports = {
  driverName: `mongo (${dbName})`,

  /* ---------- Users ---------- */
  async findUserByUsername(username) {
    const db = await getDb();
    return db.collection("users").findOne({ username });
  },
  async findUserById(id) {
    const db = await getDb();
    return db.collection("users").findOne({ id: Number(id) });
  },
  async insertUser({ username, displayName, passwordHash, role = "faculty" }) {
    const db = await getDb();
    const id = await nextSeq(db, "users");
    const user = {
      id,
      username,
      display_name: displayName,
      role,
      password_hash: passwordHash,
      created_at: new Date().toISOString(),
    };
    await db.collection("users").insertOne(user);
    return user;
  },
  async updateUserPassword(id, passwordHash, displayName) {
    const db = await getDb();
    const update = { password_hash: passwordHash };
    if (displayName) update.display_name = displayName;
    await db.collection("users").updateOne({ id: Number(id) }, { $set: update });
    return db.collection("users").findOne({ id: Number(id) });
  },

  /* ---------- Courses ---------- */
  async listCoursesByUser(userId) {
    const db = await getDb();
    return db.collection("courses").find({ user_id: Number(userId) }).sort({ updated_at: -1 }).toArray();
  },
  async findCourse(id, userId) {
    const db = await getDb();
    return db.collection("courses").findOne({ id: Number(id), user_id: Number(userId) });
  },
  async findCourseByCode(userId, code) {
    const db = await getDb();
    return db.collection("courses").findOne({ user_id: Number(userId), course_code: code });
  },
  async insertCourse({ userId, courseCode, courseName, data }) {
    const db = await getDb();
    const id = await nextSeq(db, "courses");
    const now = new Date().toISOString();
    const course = {
      id,
      user_id: Number(userId),
      course_code: courseCode,
      course_name: courseName || "",
      data_json: JSON.stringify(data),
      created_at: now,
      updated_at: now,
    };
    await db.collection("courses").insertOne(course);
    return course;
  },
  async updateCourse(id, { courseCode, courseName, data }) {
    const db = await getDb();
    const update = { updated_at: new Date().toISOString() };
    if (courseCode != null) update.course_code = courseCode;
    if (courseName != null) update.course_name = courseName;
    if (data != null) update.data_json = JSON.stringify(data);
    await db.collection("courses").updateOne({ id: Number(id) }, { $set: update });
    return db.collection("courses").findOne({ id: Number(id) });
  },
  async deleteCourse(id) {
    const db = await getDb();
    const res = await db.collection("courses").deleteOne({ id: Number(id) });
    return res.deletedCount > 0;
  },

  /* ---------- Utility ---------- */
  _dumpPath: "(MongoDB Atlas — no local file; data lives at MONGODB_URI)",
};
