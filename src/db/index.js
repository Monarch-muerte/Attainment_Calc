// Picks which database driver to use, based on DB_DRIVER in .env.
//   DB_DRIVER=json   (default) -> data/attainment.json, this computer only
//   DB_DRIVER=mongo            -> MongoDB (Atlas or any Mongo server), shared
//                                  across any computer using the same MONGODB_URI
//
// Every route calls these functions the same way regardless of which driver
// is active (the json driver's functions are synchronous, but calling code
// always uses `await`, which works fine on a plain value too).
const driver = (process.env.DB_DRIVER || "json").toLowerCase();

if (driver === "mongo" || driver === "mongodb") {
  module.exports = require("./mongo");
} else {
  module.exports = require("./json");
}
