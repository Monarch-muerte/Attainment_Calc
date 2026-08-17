-- =========================================================
-- College Attainment Portal - Database Schema
-- Import this file in phpMyAdmin (XAMPP) before using the app.
-- =========================================================

CREATE DATABASE IF NOT EXISTS college_attainment
  CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;

USE college_attainment;

-- ---------------------------------------------------------
-- Users (portal login)
-- ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS users (
  id INT AUTO_INCREMENT PRIMARY KEY,
  username VARCHAR(50) NOT NULL UNIQUE,
  password_hash VARCHAR(255) NOT NULL,
  full_name VARCHAR(100) NOT NULL DEFAULT 'Administrator',
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- Default administrator: username Admin9 / password 654321
-- (password_hash generated with PHP password_hash(), bcrypt)
INSERT INTO users (username, password_hash, full_name) VALUES
  ('Admin9', '$2y$10$fzrJa3pTVpmFPV48i0arEeltawLliH3ZLlwdoTK0wyMpcb9.JgGv6', 'Administrator')
ON DUPLICATE KEY UPDATE username = username;

-- ---------------------------------------------------------
-- Course Exit Survey  (CO1-CO6)
-- ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS course_exit_survey (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sr INT NOT NULL,
  acad_year VARCHAR(20) NOT NULL,
  sem VARCHAR(10) NULL,
  c_name VARCHAR(150) NULL,
  c_code VARCHAR(50) NULL,
  c_coordinator VARCHAR(150) NULL,
  co1 DOUBLE NULL,
  co2 DOUBLE NULL,
  co3 DOUBLE NULL,
  co4 DOUBLE NULL,
  co5 DOUBLE NULL,
  co6 DOUBLE NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- ---------------------------------------------------------
-- Program Exit Survey (PO1-PO11, PSO1-PSO2)
-- ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS program_exit_survey (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sr INT NOT NULL,
  acad_year VARCHAR(20) NOT NULL,
  po1 DOUBLE NULL, po2 DOUBLE NULL, po3 DOUBLE NULL, po4 DOUBLE NULL, po5 DOUBLE NULL, po6 DOUBLE NULL,
  po7 DOUBLE NULL, po8 DOUBLE NULL, po9 DOUBLE NULL, po10 DOUBLE NULL, po11 DOUBLE NULL,
  pso1 DOUBLE NULL, pso2 DOUBLE NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- ---------------------------------------------------------
-- Expert Lecture
-- ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS expert_lecture (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sr INT NOT NULL,
  acad_year VARCHAR(20) NOT NULL,
  `date` DATE NULL,
  sem VARCHAR(10) NULL,
  title VARCHAR(200) NULL,
  expert_name VARCHAR(150) NULL,
  po1 DOUBLE NULL, po2 DOUBLE NULL, po3 DOUBLE NULL, po4 DOUBLE NULL, po5 DOUBLE NULL, po6 DOUBLE NULL,
  po7 DOUBLE NULL, po8 DOUBLE NULL, po9 DOUBLE NULL, po10 DOUBLE NULL, po11 DOUBLE NULL,
  pso1 DOUBLE NULL, pso2 DOUBLE NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- ---------------------------------------------------------
-- Industry Visit
-- ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS industry_visit (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sr INT NOT NULL,
  acad_year VARCHAR(20) NOT NULL,
  `date` DATE NULL,
  sem VARCHAR(10) NULL,
  title VARCHAR(200) NULL,
  place VARCHAR(200) NULL,
  po1 DOUBLE NULL, po2 DOUBLE NULL, po3 DOUBLE NULL, po4 DOUBLE NULL, po5 DOUBLE NULL, po6 DOUBLE NULL,
  po7 DOUBLE NULL, po8 DOUBLE NULL, po9 DOUBLE NULL, po10 DOUBLE NULL, po11 DOUBLE NULL,
  pso1 DOUBLE NULL, pso2 DOUBLE NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- ---------------------------------------------------------
-- Alumni Survey
-- ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS alumni_survey (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sr INT NOT NULL,
  acad_year VARCHAR(20) NOT NULL,
  `date` DATE NULL,
  title VARCHAR(200) NULL,
  po1 DOUBLE NULL, po2 DOUBLE NULL, po3 DOUBLE NULL, po4 DOUBLE NULL, po5 DOUBLE NULL, po6 DOUBLE NULL,
  po7 DOUBLE NULL, po8 DOUBLE NULL, po9 DOUBLE NULL, po10 DOUBLE NULL, po11 DOUBLE NULL,
  pso1 DOUBLE NULL, pso2 DOUBLE NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- ---------------------------------------------------------
-- Industry Survey
-- ---------------------------------------------------------
CREATE TABLE IF NOT EXISTS industry_survey (
  id INT AUTO_INCREMENT PRIMARY KEY,
  sr INT NOT NULL,
  acad_year VARCHAR(20) NOT NULL,
  `date` DATE NULL,
  industry_details VARCHAR(200) NULL,
  po1 DOUBLE NULL, po2 DOUBLE NULL, po3 DOUBLE NULL, po4 DOUBLE NULL, po5 DOUBLE NULL, po6 DOUBLE NULL,
  po7 DOUBLE NULL, po8 DOUBLE NULL, po9 DOUBLE NULL, po10 DOUBLE NULL, po11 DOUBLE NULL,
  pso1 DOUBLE NULL, pso2 DOUBLE NULL,
  created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB;
