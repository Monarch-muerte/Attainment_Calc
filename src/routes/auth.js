const express = require("express");
const bcrypt = require("bcryptjs");
const db = require("../db");
const { signToken, requireAuth } = require("../middleware/auth");

const router = express.Router();

function publicUser(u) {
  return { id: u.id, username: u.username, displayName: u.display_name, role: u.role };
}

router.post("/register", async (req, res) => {
  try {
    const { username, password, displayName } = req.body || {};
    if (!username || !password || !displayName) {
      return res.status(400).json({ error: "username, password and displayName are required." });
    }
    const uname = String(username).trim().toLowerCase();
    if (uname.length < 3) return res.status(400).json({ error: "Username must be at least 3 characters." });
    if (String(password).length < 4) return res.status(400).json({ error: "Password must be at least 4 characters." });

    if (await db.findUserByUsername(uname)) {
      return res.status(409).json({ error: "That username is already taken." });
    }

    const hash = await bcrypt.hash(password, 10);
    const user = await db.insertUser({ username: uname, displayName: String(displayName).trim(), passwordHash: hash });

    const token = signToken(user);
    res.status(201).json({ token, user: publicUser(user) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Could not create account." });
  }
});

router.post("/login", async (req, res) => {
  try {
    const { username, password } = req.body || {};
    if (!username || !password) return res.status(400).json({ error: "username and password are required." });
    const uname = String(username).trim().toLowerCase();

    const user = await db.findUserByUsername(uname);
    if (!user) return res.status(401).json({ error: "No account found with that username." });

    const ok = await bcrypt.compare(password, user.password_hash);
    if (!ok) return res.status(401).json({ error: "Incorrect password." });

    const token = signToken(user);
    res.json({ token, user: publicUser(user) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Login failed." });
  }
});

router.get("/me", requireAuth, async (req, res) => {
  const user = await db.findUserById(req.user.id);
  if (!user) return res.status(404).json({ error: "User not found." });
  res.json({ user: publicUser(user) });
});

router.put("/profile", requireAuth, async (req, res) => {
  try {
    const currentUser = await db.findUserById(req.user.id);
    if (!currentUser) return res.status(404).json({ error: "User not found." });

    const { displayName, username, currentPassword, newPassword } = req.body || {};
    const updates = {};

    // 1. Update Display Name / Full Name
    if (displayName !== undefined) {
      const dname = String(displayName).trim();
      if (dname.length < 2) {
        return res.status(400).json({ error: "Name must be at least 2 characters." });
      }
      updates.displayName = dname;
    }

    // 2. Update Username / User ID
    if (username !== undefined) {
      const uname = String(username).trim().toLowerCase();
      if (uname.length < 3) {
        return res.status(400).json({ error: "User ID / Username must be at least 3 characters." });
      }
      if (uname !== currentUser.username) {
        const existing = await db.findUserByUsername(uname);
        if (existing && existing.id !== currentUser.id) {
          return res.status(409).json({ error: "That User ID / Username is already taken by another user." });
        }
        updates.username = uname;
      }
    }

    // 3. Update Password (only for the currently logged-in user, verifying current password)
    if (newPassword) {
      if (!currentPassword) {
        return res.status(400).json({ error: "Current password is required to set a new password." });
      }
      const ok = await bcrypt.compare(currentPassword, currentUser.password_hash);
      if (!ok) {
        return res.status(401).json({ error: "Current password is incorrect." });
      }
      if (String(newPassword).length < 4) {
        return res.status(400).json({ error: "New password must be at least 4 characters." });
      }
      updates.passwordHash = await bcrypt.hash(newPassword, 10);
    }

    if (Object.keys(updates).length === 0) {
      return res.json({ user: publicUser(currentUser), message: "No changes requested." });
    }

    const updatedUser = await db.updateUserProfile(currentUser.id, updates);
    const token = signToken(updatedUser);

    res.json({
      token,
      user: publicUser(updatedUser),
      message: "Profile updated successfully.",
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Failed to update profile." });
  }
});

module.exports = router;
