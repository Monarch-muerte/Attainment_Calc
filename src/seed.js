// Creates (or resets) a default demo login so the project works out of the box.
//   Username: Admin9
//   Password: 654321
//
// Run with:  npm run seed
// Works against whichever database driver is set in .env (DB_DRIVER=json or mongo).
require("dotenv").config();
const bcrypt = require("bcryptjs");
const db = require("./db");

const USERNAME = "admin9";
const DISPLAY_NAME = "Admin";
const PASSWORD = "654321";

async function seed() {
  console.log(`Using database driver: ${db.driverName}`);
  const hash = await bcrypt.hash(PASSWORD, 10);
  const existing = await db.findUserByUsername(USERNAME);

  if (existing) {
    await db.updateUserPassword(existing.id, hash, DISPLAY_NAME);
    console.log(`Updated existing account "${USERNAME}" with password "${PASSWORD}".`);
  } else {
    await db.insertUser({ username: USERNAME, displayName: DISPLAY_NAME, passwordHash: hash, role: "admin" });
    console.log(`Created account "${USERNAME}" with password "${PASSWORD}".`);
  }
  console.log("\nLogin with:");
  console.log("  Username: Admin9");
  console.log("  Password: 654321");
  console.log(`\nDatabase: ${db._dumpPath}`);
}

seed()
  .then(() => process.exit(0))
  .catch((err) => {
    console.error("Seed failed:", err.message);
    process.exit(1);
  });
