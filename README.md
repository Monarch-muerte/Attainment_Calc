# College Attainment Portal (PHP + MySQL)

A department portal for recording CO / PO / PSO attainment evidence, built for K.D. Polytechnic — Computer Engineering. Plain PHP (PDO + prepared statements) and MySQL — no framework, so it runs on any XAMPP install.

## What it does

Six activity types, each with add / edit / delete / history and one shared login:

| Activity              | Table                  | Fields                                                              | Outcomes         |
|------------------------|-------------------------|----------------------------------------------------------------------|-------------------|
| Course Exit Survey     | `course_exit_survey`    | acad_year, sem, c_name, c_code, c_coordinator                        | CO1–CO6           |
| Program Exit Survey    | `program_exit_survey`   | acad_year                                                             | PO1–PO11, PSO1–PSO2 |
| Expert Lecture         | `expert_lecture`        | acad_year, date, sem, title, expert_name                             | PO1–PO11, PSO1–PSO2 |
| Industry Visit         | `industry_visit`        | acad_year, date, sem, title, place                                   | PO1–PO11, PSO1–PSO2 |
| Alumni Survey          | `alumni_survey`         | acad_year, date, title                                                | PO1–PO11, PSO1–PSO2 |
| Industry Survey        | `industry_survey`       | acad_year, date, industry_details                                    | PO1–PO11, PSO1–PSO2 |

The dashboard shows a live count per activity, a **combined PO/PSO attainment average** (computed in SQL across the five PO-based tables, ignoring blank values), and a separate **CO summary** for Course Exit Survey.

## Importing data from Excel

Every activity's History page has an **Import from Excel** button (`import.php?table=...`), alongside a **Download template** link that gives you a `.csv` with the exact header row expected for that table.

Two ways it can read a file:

1. **Flat import (all 6 activities).** Upload a `.xlsx` or `.csv` whose first row is a header matching the field/outcome names (e.g. "Academic Year", "PO1" ... "PSO2"). Column order doesn't matter — headers are matched by name — and outcome cells may be left blank. Each subsequent row becomes one record; rows missing a required field, or with a non-numeric outcome value, are skipped and listed in the import summary rather than failing the whole batch.
2. **Course Exit Survey auto-detect.** If you upload the department's standard "Course Evaluation Plan" workbook (the one with `PR`, `TH`, `ATTAINMENT SUMMARY`, `FINAL` sheets, etc.) instead of a flat file, the importer looks for the CO ATTAINMENT block (the table with a `CO` / `CO ATTAIN` header) and pulls Course Code, Course Name, Semester, Batch and CO1–CO6 straight out of it into one Course Exit Survey record — no reformatting needed. This was built and tested against a real department workbook.

The `.xlsx` reader (`includes/xlsx_reader.php`) is hand-written against the raw zip/XML format so the app has zero Composer dependencies — just the `zip` and `xml` PHP extensions, both enabled by default in XAMPP.

## Setup (XAMPP)

1. Start **Apache** and **MySQL** from the XAMPP control panel.
2. Copy the `attainment_portal` folder into `htdocs`, e.g. `C:\xampp\htdocs\attainment_portal`.
3. Open **phpMyAdmin** (`http://localhost/phpmyadmin`), click **Import**, and import `sql/schema.sql`. This creates the `college_attainment` database, all six tables, and seeds the default admin login.
4. If your MySQL uses a different user/password than the XAMPP default (root, no password), edit `config/db.php`.
5. Visit `http://localhost/attainment_portal/` in your browser.

**Default login:** `Admin9` / `654321`

## Project structure

```
attainment_portal/
├── config/db.php            # database connection
├── sql/schema.sql            # run this in phpMyAdmin first
├── includes/
│   ├── auth.php               # session helpers, require_login()
│   ├── tables_config.php      # single source of truth for all 6 tables/fields
│   ├── functions.php          # escaping, averages, table whitelist guard, insert_record()
│   ├── xlsx_reader.php        # dependency-free .xlsx reader + department workbook auto-detect
│   ├── header.php / footer.php
│   └── form_view.php          # shared add/edit form renderer
├── assets/css/style.css
├── index.php                 # login
├── logout.php
├── dashboard.php
├── history.php                # ?table=xxx picker + listing, edit/delete links
├── add.php                    # ?table=xxx
├── edit.php                   # ?table=xxx&id=xxx
├── delete.php                 # POST only
├── import.php                 # ?table=xxx — Excel/CSV bulk import
└── template.php                # ?table=xxx — downloadable CSV header template
```

Every page reads the six activity types from one config array (`includes/tables_config.php`) instead of duplicating markup per table, so adding a 7th activity later just means adding one entry there plus a table in the database.

## Security notes

- All SQL uses PDO **prepared statements** — no string-built queries with user input.
- The `table` query parameter is always checked against a whitelist (`resolve_table()`) before being used in any query, so it can't be used to reach an arbitrary table.
- Passwords are stored as bcrypt hashes (`password_hash` / `password_verify`), not plaintext.
- Every page except `index.php` calls `require_login()` and redirects to the login page if there's no active session.
- This is a college-project-scale app — for real deployment you'd want to add rate limiting on login, CSRF tokens on forms, and HTTPS.

## Notes

- Outcome fields (CO1–CO6, PO1–PO11, PSO1–PSO2) are optional per record — leave blank if not assessed. The dashboard averages ignore blanks and only average over entered values.
- `sr` (serial number) is auto-assigned per table on save.
