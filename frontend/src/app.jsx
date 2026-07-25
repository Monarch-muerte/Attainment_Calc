import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine, Cell,
} from "recharts";
import {
  LogIn, LogOut, User, Lock, Eye, EyeOff, UploadCloud, Download, Save, Plus, Trash2,
  CheckCircle2, XCircle, Settings2, GitBranch, ClipboardList, BarChart3,
  RefreshCw, GraduationCap, AlertTriangle, X, Users2, BookOpenCheck, FolderOpen,
} from "lucide-react";
import { api, setToken, getStoredToken } from "./api";

/* ============================================================================
   CONSTANTS
============================================================================ */
const COs = ["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"];
const POs = Array.from({ length: 12 }, (_, i) => `PO${i + 1}`);
const PSOs = ["PSO1", "PSO2"];
const POPSO = [...POs, ...PSOs];

const DEFAULT_WEIGHTS = { internal: 0.4, endsem: 0.6, direct: 0.8, indirect: 0.2 };
const DEFAULT_TARGETS = { targetPct: 0.6, coTargetLevel: 2.0, level3: 0.7, level2: 0.6, level1: 0.5 };

const SAMPLE = {"courseInfo":{"institute":"KDP, Patan","department":"Computer Engineering","courseName":"Data Structures","courseCode":"CE301","semester":"3","year":"2025-26","faculty":"Prof. A. Sharma"},"coStatements":["Explain fundamental concepts of linear and non-linear data structures.","Apply arrays, stacks and queues to solve computational problems.","Design and implement linked list based data structures.","Analyse time and space complexity of tree and graph algorithms.","Implement sorting and searching algorithms and compare their efficiency.","Select an appropriate data structure and algorithm for a given problem."],"internal1":{"maxMarks":[10,10,10,10,10,10],"students":[{"roll":"R01","name":"Student 1","marks":[9,9,7,5,5,7]},{"roll":"R02","name":"Student 2","marks":[5,4,7,7,8,5]},{"roll":"R03","name":"Student 3","marks":[7,7,4,8,7,10]},{"roll":"R04","name":"Student 4","marks":[7,7,9,7,8,6]},{"roll":"R05","name":"Student 5","marks":[7,9,8,7,5,8]},{"roll":"R06","name":"Student 6","marks":[7,8,7,9,7,7]},{"roll":"R07","name":"Student 7","marks":[8,5,6,6,10,7]},{"roll":"R08","name":"Student 8","marks":[8,8,6,4,9,6]},{"roll":"R09","name":"Student 9","marks":[8,4,6,9,9,4]},{"roll":"R10","name":"Student 10","marks":[4,7,8,7,7,5]},{"roll":"R11","name":"Student 11","marks":[8,9,6,4,5,8]},{"roll":"R12","name":"Student 12","marks":[4,7,5,7,6,7]},{"roll":"R13","name":"Student 13","marks":[10,8,9,7,6,7]},{"roll":"R14","name":"Student 14","marks":[2,7,7,5,8,6]},{"roll":"R15","name":"Student 15","marks":[2,6,5,6,7,9]}]},"internal2":{"maxMarks":[10,10,10,10,10,10],"students":[{"roll":"R01","name":"Student 1","marks":[10,6,8,7,8,4]},{"roll":"R02","name":"Student 2","marks":[6,5,5,5,6,6]},{"roll":"R03","name":"Student 3","marks":[5,8,6,1,9,6]},{"roll":"R04","name":"Student 4","marks":[5,7,7,7,5,7]},{"roll":"R05","name":"Student 5","marks":[4,9,5,6,7,7]},{"roll":"R06","name":"Student 6","marks":[6,8,0,6,6,6]},{"roll":"R07","name":"Student 7","marks":[9,5,6,3,7,4]},{"roll":"R08","name":"Student 8","marks":[4,10,8,7,7,4]},{"roll":"R09","name":"Student 9","marks":[5,7,3,7,3,7]},{"roll":"R10","name":"Student 10","marks":[5,10,8,6,3,5]},{"roll":"R11","name":"Student 11","marks":[6,5,7,8,6,6]},{"roll":"R12","name":"Student 12","marks":[8,6,8,6,10,6]},{"roll":"R13","name":"Student 13","marks":[5,7,5,5,6,8]},{"roll":"R14","name":"Student 14","marks":[3,6,6,6,8,4]},{"roll":"R15","name":"Student 15","marks":[8,6,7,6,6,6]}]},"assignment":{"maxMarks":[5,5,5,5,5,5],"students":[{"roll":"R01","name":"Student 1","marks":[3,5,3,4,3,3]},{"roll":"R02","name":"Student 2","marks":[5,4,3,4,4,3]},{"roll":"R03","name":"Student 3","marks":[4,3,3,3,2,2]},{"roll":"R04","name":"Student 4","marks":[2,3,3,3,3,2]},{"roll":"R05","name":"Student 5","marks":[3,4,4,3,3,2]},{"roll":"R06","name":"Student 6","marks":[3,1,2,4,1,4]},{"roll":"R07","name":"Student 7","marks":[4,3,4,4,4,3]},{"roll":"R08","name":"Student 8","marks":[3,3,3,3,3,4]},{"roll":"R09","name":"Student 9","marks":[2,2,3,2,5,1]},{"roll":"R10","name":"Student 10","marks":[3,3,5,2,4,3]},{"roll":"R11","name":"Student 11","marks":[3,3,4,2,3,4]},{"roll":"R12","name":"Student 12","marks":[5,1,5,4,3,4]},{"roll":"R13","name":"Student 13","marks":[3,5,4,3,3,3]},{"roll":"R14","name":"Student 14","marks":[3,3,5,2,0,3]},{"roll":"R15","name":"Student 15","marks":[3,4,3,3,4,4]}]},"endsem":{"maxMarks":[10,10,10,10,10,10,10,10,10,10,10,10],"students":[{"roll":"R01","name":"Student 1","marks":[6,7,5,7,8,7,9,4,6,5,5,6]},{"roll":"R02","name":"Student 2","marks":[7,7,7,10,8,3,7,5,5,9,6,2]},{"roll":"R03","name":"Student 3","marks":[7,6,4,4,5,6,5,6,10,5,5,6]},{"roll":"R04","name":"Student 4","marks":[8,5,9,4,4,6,4,5,7,8,6,6]},{"roll":"R05","name":"Student 5","marks":[9,7,6,9,4,7,8,6,5,7,5,5]},{"roll":"R06","name":"Student 6","marks":[4,9,5,9,7,6,8,7,6,3,6,8]},{"roll":"R07","name":"Student 7","marks":[7,8,9,7,10,3,6,1,8,6,10,5]},{"roll":"R08","name":"Student 8","marks":[2,9,4,4,6,7,5,7,2,10,6,6]},{"roll":"R09","name":"Student 9","marks":[7,7,7,4,6,7,8,6,6,7,7,7]},{"roll":"R10","name":"Student 10","marks":[6,5,8,7,6,10,8,8,7,4,8,6]},{"roll":"R11","name":"Student 11","marks":[6,8,8,7,6,10,7,4,8,7,6,8]},{"roll":"R12","name":"Student 12","marks":[7,10,7,7,9,5,8,6,9,5,7,3]},{"roll":"R13","name":"Student 13","marks":[5,6,7,10,10,4,5,9,7,3,7,4]},{"roll":"R14","name":"Student 14","marks":[4,4,6,10,7,6,10,5,6,6,7,6]},{"roll":"R15","name":"Student 15","marks":[9,6,7,7,4,4,6,2,6,6,6,6]}]},"survey":{"students":[{"roll":"R01","name":"Student 1","ratings":[3,3,1,3,1,1]},{"roll":"R02","name":"Student 2","ratings":[1,3,2,3,2,1]},{"roll":"R03","name":"Student 3","ratings":[2,2,2,3,3,2]},{"roll":"R04","name":"Student 4","ratings":[3,3,2,3,2,2]},{"roll":"R05","name":"Student 5","ratings":[1,2,3,3,2,3]},{"roll":"R06","name":"Student 6","ratings":[2,2,2,3,3,3]},{"roll":"R07","name":"Student 7","ratings":[2,2,2,2,2,2]},{"roll":"R08","name":"Student 8","ratings":[2,2,3,3,2,3]},{"roll":"R09","name":"Student 9","ratings":[1,1,2,2,1,2]},{"roll":"R10","name":"Student 10","ratings":[3,3,2,3,3,2]},{"roll":"R11","name":"Student 11","ratings":[2,3,2,3,3,3]},{"roll":"R12","name":"Student 12","ratings":[3,2,3,1,1,3]},{"roll":"R13","name":"Student 13","ratings":[2,3,3,3,3,1]},{"roll":"R14","name":"Student 14","ratings":[3,2,3,3,1,2]},{"roll":"R15","name":"Student 15","ratings":[2,3,2,1,3,3]}]},"mapping":[[3,2,1,0,0,0,0,0,0,1,0,1,2,1],[2,3,2,1,0,0,0,0,1,1,0,1,2,2],[1,2,3,2,1,0,0,0,1,2,0,2,1,2],[0,1,2,3,2,1,0,0,1,2,1,2,1,3],[0,0,1,2,3,2,1,0,1,2,1,2,2,2],[1,1,1,2,2,3,1,1,2,2,1,2,3,3]],"targetPO":[2,2,2,2,2,2,2,2,2,2,2,2,2,2]};

function blankState() {
  return {
    courseInfo: { institute: "", department: "", courseName: "", courseCode: "", semester: "", year: "", faculty: "" },
    coStatements: COs.map(() => ""),
    weights: { internal: 0, endsem: 0, direct: 0, indirect: 0 },
    targets: { targetPct: 0, coTargetLevel: 0, level3: 0, level2: 0, level1: 0 },
    mapping: COs.map(() => POPSO.map(() => 0)),
    targetPO: POPSO.map(() => 0),
    internal1: { maxMarks: Array(6).fill(0), students: [] },
    internal2: { maxMarks: Array(6).fill(0), students: [] },
    assignment: { maxMarks: Array(6).fill(0), students: [] },
    endsem: { maxMarks: Array(12).fill(0), students: [] },
    survey: { students: [] },
  };
}
function sampleState() {
  return { ...JSON.parse(JSON.stringify(SAMPLE)), weights: { ...DEFAULT_WEIGHTS }, targets: { ...DEFAULT_TARGETS } };
}

// Clamps any state coming from an external source (a saved course loaded from the
// database, or an uploaded Excel file) into valid ranges. This repairs bad values
// that may have been saved before validation existed, and guards against malformed
// spreadsheet data — without changing anything about a state that is already valid.
function clampNum(v, min, max, fallback) {
  const n = Number(v);
  if (Number.isNaN(n)) return fallback;
  return Math.min(max, Math.max(min, n));
}
function normalizeState(s) {
  const weights = {
    internal: clampNum(s.weights?.internal, 0, 1, DEFAULT_WEIGHTS.internal),
    endsem: clampNum(s.weights?.endsem, 0, 1, DEFAULT_WEIGHTS.endsem),
    direct: clampNum(s.weights?.direct, 0, 1, DEFAULT_WEIGHTS.direct),
    indirect: clampNum(s.weights?.indirect, 0, 1, DEFAULT_WEIGHTS.indirect),
  };
  const targets = {
    targetPct: clampNum(s.targets?.targetPct, 0, 1, DEFAULT_TARGETS.targetPct),
    coTargetLevel: clampNum(s.targets?.coTargetLevel, 0, 3, DEFAULT_TARGETS.coTargetLevel),
    level3: clampNum(s.targets?.level3, 0, 1, DEFAULT_TARGETS.level3),
    level2: clampNum(s.targets?.level2, 0, 1, DEFAULT_TARGETS.level2),
    level1: clampNum(s.targets?.level1, 0, 1, DEFAULT_TARGETS.level1),
  };
  const mapping = (s.mapping && s.mapping.length === 6 ? s.mapping : blankState().mapping)
    .map((row) => POPSO.map((_, j) => clampNum(row[j], 0, 3, 0)));
  const targetPO = (s.targetPO && s.targetPO.length === POPSO.length ? s.targetPO : POPSO.map(() => 2))
    .map((v) => clampNum(v, 0, 3, 2));

  function normalizeAssessment(a, nQ) {
    const maxMarks = Array.from({ length: nQ }, (_, i) => clampNum(a?.maxMarks?.[i], 0, 100, 10));
    const students = (a?.students || []).map((st) => ({
      roll: st.roll ?? "",
      name: st.name ?? "",
      marks: Array.from({ length: nQ }, (_, i) => clampNum(st.marks?.[i], 0, maxMarks[i], 0)),
    }));
    return { maxMarks, students };
  }
  const survey = {
    students: (s.survey?.students || []).map((st) => ({
      roll: st.roll ?? "",
      name: st.name ?? "",
      ratings: Array.from({ length: 6 }, (_, i) => clampNum(st.ratings?.[i], 1, 3, 2)),
    })),
  };

  return {
    courseInfo: { ...blankState().courseInfo, ...(s.courseInfo || {}) },
    coStatements: COs.map((_, i) => s.coStatements?.[i] ?? ""),
    weights, targets, mapping, targetPO,
    internal1: normalizeAssessment(s.internal1, 6),
    internal2: normalizeAssessment(s.internal2, 6),
    assignment: normalizeAssessment(s.assignment, 6),
    endsem: normalizeAssessment(s.endsem, 12),
    survey,
  };
}

/* ============================================================================
   ATTAINMENT MATH
============================================================================ */
function levelFromPct(pct, targets) {
  if (pct >= targets.level3) return 3;
  if (pct >= targets.level2) return 2;
  if (pct >= targets.level1) return 1;
  return 0;
}
function simpleCOStats(assessment, targets) {
  const n = assessment.students.length;
  return assessment.maxMarks.map((max, i) => {
    const m = Number(max) || 0;
    const targetMarks = m * targets.targetPct;
    const attained = assessment.students.filter((s) => (Number(s.marks[i]) || 0) >= targetMarks).length;
    const pct = n ? attained / n : 0;
    return { max: m, targetMarks, attained, total: n, pct, level: n ? levelFromPct(pct, targets) : 0 };
  });
}
function endsemCOStats(assessment, targets) {
  const n = assessment.students.length;
  const res = [];
  for (let i = 0; i < 6; i++) {
    const max = (Number(assessment.maxMarks[i]) || 0) + (Number(assessment.maxMarks[i + 6]) || 0);
    const targetMarks = max * targets.targetPct;
    const attained = assessment.students.filter((s) => {
      const tot = (Number(s.marks[i]) || 0) + (Number(s.marks[i + 6]) || 0);
      return tot >= targetMarks;
    }).length;
    const pct = n ? attained / n : 0;
    res.push({ max, targetMarks, attained, total: n, pct, level: n ? levelFromPct(pct, targets) : 0 });
  }
  return res;
}
function surveyCOAverage(survey) {
  const n = survey.students.length;
  const res = [];
  for (let i = 0; i < 6; i++) {
    const sum = survey.students.reduce((a, s) => a + (Number(s.ratings[i]) || 0), 0);
    res.push(n ? sum / n : 0);
  }
  return res;
}

/* ============================================================================
   EXCEL IMPORT / EXPORT
============================================================================ */
function sheetRows(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return null;
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}
function parseCourseInfo(rows, state) {
  const ci = { ...state.courseInfo };
  const co = [...state.coStatements];
  const w = { ...state.weights };
  const t = { ...state.targets };
  for (const row of rows) {
    const label = row[1] == null ? "" : String(row[1]).trim();
    const value = row[2];
    if (!label) continue;
    if (label === "Institute Name") ci.institute = value ?? "";
    else if (label === "Department") ci.department = value ?? "";
    else if (label === "Course Name") ci.courseName = value ?? "";
    else if (label === "Course Code") ci.courseCode = value ?? "";
    else if (label === "Semester") ci.semester = value ?? "";
    else if (label === "Academic Year") ci.year = value ?? "";
    else if (label === "Faculty Name") ci.faculty = value ?? "";
    else if (/^CO[1-6]$/.test(label)) co[Number(label.slice(2)) - 1] = value ?? "";
    else if (label.includes("Indirect Attainment Weight")) w.indirect = Number(value) || 0;
    else if (label.includes("Direct Attainment Weight")) w.direct = Number(value) || 0;
    else if (label.includes("Internal Weight")) w.internal = Number(value) || 0;
    else if (label.includes("End-Sem Weight")) w.endsem = Number(value) || 0;
    else if (label.includes("Target % of marks")) t.targetPct = Number(value) || 0;
    else if (label.includes("CO Target Attainment Level")) t.coTargetLevel = Number(value) || 0;
    else if (label.includes("Level 3")) t.level3 = Number(value) || 0;
    else if (label.includes("Level 2")) t.level2 = Number(value) || 0;
    else if (label.includes("Level 1")) t.level1 = Number(value) || 0;
  }
  return { courseInfo: ci, coStatements: co, weights: w, targets: t };
}
function parseMapping(rows, state) {
  let headerIdx = rows.findIndex((r) => r[1] === "CO / PO");
  if (headerIdx === -1) return {};
  const mapping = [];
  let r = headerIdx + 1;
  while (r < rows.length && /^CO[1-6]$/.test(String(rows[r][1] || ""))) {
    mapping.push(POPSO.map((_, j) => Number(rows[r][2 + j]) || 0));
    r++;
  }
  const targetRow = rows.find((row) => row[1] === "Target Attainment");
  const targetPO = targetRow ? POPSO.map((_, j) => Number(targetRow[2 + j]) || 2) : state.targetPO;
  return { mapping: mapping.length === 6 ? mapping : state.mapping, targetPO };
}
function parseAssessment(rows, nQ) {
  const maxRow = rows.find((r) => r[1] === "Max Marks \u2192");
  const maxMarks = maxRow ? Array.from({ length: nQ }, (_, i) => Number(maxRow[3 + i]) || 0) : Array(nQ).fill(10);
  const headerIdx = rows.findIndex((r) => r[1] === "Roll No");
  const students = [];
  if (headerIdx !== -1) {
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row[1] == null || row[1] === "") break;
      students.push({
        roll: String(row[1]),
        name: row[2] == null ? "" : String(row[2]),
        marks: Array.from({ length: nQ }, (_, i) => Number(row[3 + i]) || 0),
      });
    }
  }
  return { maxMarks, students };
}
function parseSurvey(rows) {
  const headerIdx = rows.findIndex((r) => r[1] === "Roll No");
  const students = [];
  if (headerIdx !== -1) {
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row[1] == null || row[1] === "") break;
      students.push({
        roll: String(row[1]),
        name: row[2] == null ? "" : String(row[2]),
        ratings: Array.from({ length: 6 }, (_, i) => Number(row[3 + i]) || 0),
      });
    }
  }
  return { students };
}
function parseWorkbook(wb, prevState) {
  let next = JSON.parse(JSON.stringify(prevState));
  const found = [], missing = [];
  const ci = sheetRows(wb, "Course_Info");
  if (ci) { Object.assign(next, parseCourseInfo(ci, next)); found.push("Course_Info"); } else missing.push("Course_Info");
  const mp = sheetRows(wb, "CO_PO_PSO_Mapping");
  if (mp) { Object.assign(next, parseMapping(mp, next)); found.push("CO_PO_PSO_Mapping"); } else missing.push("CO_PO_PSO_Mapping");
  const i1 = sheetRows(wb, "Internal_Exam1");
  if (i1) { next.internal1 = parseAssessment(i1, 6); found.push("Internal_Exam1"); } else missing.push("Internal_Exam1");
  const i2 = sheetRows(wb, "Internal_Exam2");
  if (i2) { next.internal2 = parseAssessment(i2, 6); found.push("Internal_Exam2"); } else missing.push("Internal_Exam2");
  const asg = sheetRows(wb, "Assignment");
  if (asg) { next.assignment = parseAssessment(asg, 6); found.push("Assignment"); } else missing.push("Assignment");
  const es = sheetRows(wb, "End_Sem_Exam");
  if (es) { next.endsem = parseAssessment(es, 12); found.push("End_Sem_Exam"); } else missing.push("End_Sem_Exam");
  const sv = sheetRows(wb, "Indirect_CO_Attainment");
  if (sv) { next.survey = parseSurvey(sv); found.push("Indirect_CO_Attainment"); } else missing.push("Indirect_CO_Attainment");
  return { next, found, missing };
}
function exportWorkbook(state) {
  const wb = XLSX.utils.book_new();
  const ci = [
    ["", "COURSE INFORMATION & SETTINGS"], [], ["", "Institute Name", state.courseInfo.institute],
    ["", "Department", state.courseInfo.department], ["", "Course Name", state.courseInfo.courseName],
    ["", "Course Code", state.courseInfo.courseCode], ["", "Semester", state.courseInfo.semester],
    ["", "Academic Year", state.courseInfo.year], ["", "Faculty Name", state.courseInfo.faculty], [],
    ...COs.map((co, i) => ["", co, state.coStatements[i]]), [],
    ["", "Internal Weight (Internal1+Internal2+Assignment average) for Direct Attainment", state.weights.internal],
    ["", "End-Sem Weight for Direct Attainment", state.weights.endsem],
    ["", "Direct Attainment Weight (overall CO attainment)", state.weights.direct],
    ["", "Indirect Attainment Weight (overall CO attainment, from survey)", state.weights.indirect], [],
    ["", "Target % of marks per CO (student must score >= this % to be counted 'attained')", state.targets.targetPct],
    ["", "CO Target Attainment Level (for gap analysis, scale 0-3)", state.targets.coTargetLevel], [],
    ["", "Level 3 \u2014 minimum % of students", state.targets.level3],
    ["", "Level 2 \u2014 minimum % of students", state.targets.level2],
    ["", "Level 1 \u2014 minimum % of students", state.targets.level1],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ci), "Course_Info");

  const mp = [["", "CO / PO", ...POPSO], ...COs.map((co, i) => ["", co, ...state.mapping[i]]),
    ["", "Target Attainment", ...state.targetPO]];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mp), "CO_PO_PSO_Mapping");

  function assessmentSheet(data, nQ) {
    const rows = [
      ["", "Max Marks \u2192", "", ...data.maxMarks],
      ["", "Roll No", "Name", ...Array.from({ length: nQ }, (_, i) => `Q${i + 1}`), "Total"],
      ...data.students.map((s) => ["", s.roll, s.name, ...s.marks, s.marks.reduce((a, b) => a + (Number(b) || 0), 0)]),
    ];
    return XLSX.utils.aoa_to_sheet(rows);
  }
  XLSX.utils.book_append_sheet(wb, assessmentSheet(state.internal1, 6), "Internal_Exam1");
  XLSX.utils.book_append_sheet(wb, assessmentSheet(state.internal2, 6), "Internal_Exam2");
  XLSX.utils.book_append_sheet(wb, assessmentSheet(state.assignment, 6), "Assignment");
  XLSX.utils.book_append_sheet(wb, assessmentSheet(state.endsem, 12), "End_Sem_Exam");

  const sv = [["", "Roll No", "Name", ...COs],
    ...state.survey.students.map((s) => ["", s.roll, s.name, ...s.ratings])];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sv), "Indirect_CO_Attainment");

  XLSX.writeFile(wb, `${state.courseInfo.courseCode || "attainment"}_data.xlsx`);
}

/* ============================================================================
   VALIDATION HELPER
============================================================================ */
function clamp(value, min, max) {
  const n = Number(value);
  if (Number.isNaN(n)) return min;
  return Math.min(max, Math.max(min, n));
}

/* ============================================================================
   SMALL UI PRIMITIVES
============================================================================ */
function Field({ label, children }) {
  return (
    <label className="field">
      <span className="field-label">{label}</span>
      {children}
    </label>
  );
}
function StatusPill({ ok }) {
  return (
    <span className={`pill ${ok ? "pill-ok" : "pill-bad"}`}>
      {ok ? <CheckCircle2 size={13} /> : <XCircle size={13} />}
      {ok ? "Attained" : "Not Attained"}
    </span>
  );
}
function Ring({ value, total, label, color }) {
  const r = 42, c = 2 * Math.PI * r;
  const frac = total ? value / total : 0;
  return (
    <div className="ring-wrap">
      <svg width="110" height="110" viewBox="0 0 110 110">
        <circle cx="55" cy="55" r={r} fill="none" stroke="#e7e9f3" strokeWidth="10" />
        <circle cx="55" cy="55" r={r} fill="none" stroke={color} strokeWidth="10" strokeLinecap="round"
          strokeDasharray={`${c * frac} ${c}`} transform="rotate(-90 55 55)" />
        <text x="55" y="50" textAnchor="middle" className="ring-num">{value}/{total}</text>
        <text x="55" y="67" textAnchor="middle" className="ring-sub">attained</text>
      </svg>
      <div className="ring-label">{label}</div>
    </div>
  );
}

/* ============================================================================
   LOGIN / REGISTER PAGE
============================================================================ */
function Emblem() {
  return (
    <svg width="96" height="96" viewBox="0 0 96 96">
      <circle cx="48" cy="48" r="46" fill="#fff" stroke="#e3e6f3" strokeWidth="2" />
      <circle cx="48" cy="48" r="40" fill="none" stroke="#312e81" strokeWidth="2.5" />
      <path d="M48 12 A36 36 0 0 1 84 48 A36 36 0 0 1 48 84 A36 36 0 0 1 12 48 A36 36 0 0 1 48 12 Z"
        fill="none" stroke="#312e81" strokeWidth="1" strokeDasharray="2 3" />
      <path d="M31 34 L48 24 L65 34 L65 54 C65 66 57 72 48 76 C39 72 31 66 31 54 Z" fill="#4338ca" opacity="0.12" />
      <path d="M31 34 L48 24 L65 34 L65 54 C65 66 57 72 48 76 C39 72 31 66 31 54 Z" fill="none" stroke="#312e81" strokeWidth="2" />
      <rect x="41" y="40" width="14" height="10" rx="1.5" fill="none" stroke="#312e81" strokeWidth="1.6" />
      <line x1="44" y1="42.5" x2="52" y2="42.5" stroke="#312e81" strokeWidth="1" />
      <line x1="44" y1="45" x2="52" y2="45" stroke="#312e81" strokeWidth="1" />
      <line x1="44" y1="47.5" x2="52" y2="47.5" stroke="#312e81" strokeWidth="1" />
      <circle cx="48" cy="58" r="3.4" fill="none" stroke="#a21caf" strokeWidth="1.5" />
      <circle cx="48" cy="58" r="1.1" fill="#a21caf" />
      <path d="M20 44 Q16 48 20 52" fill="none" stroke="#312e81" strokeWidth="1.4" />
      <path d="M76 44 Q80 48 76 52" fill="none" stroke="#312e81" strokeWidth="1.4" />
      <rect x="30" y="63" width="36" height="9" rx="2" fill="#b91c3c" opacity="0.9" />
      <text x="48" y="70" textAnchor="middle" fontSize="6" fill="#fff" fontFamily="Georgia, serif" fontWeight="bold">EST. 1961</text>
    </svg>
  );
}

function LoginPage({ onAuthed }) {
  const [mode, setMode] = useState("login");
  const [username, setUsername] = useState("");
  const [displayName, setDisplayName] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [remember, setRemember] = useState(true);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function handleLogin(e) {
    e.preventDefault();
    setError(""); setBusy(true);
    try {
      if (!username.trim() || !password) throw new Error("Enter your username and password.");
      const { token, user } = await api.login(username.trim(), password);
      if (remember) setToken(token); else setToken(null);
      onAuthed(user, token);
    } catch (err) {
      setError(err.message || "Login failed.");
    }
    setBusy(false);
  }
  async function handleRegister(e) {
    e.preventDefault();
    setError(""); setBusy(true);
    try {
      if (!username.trim() || !displayName.trim() || !password) throw new Error("Fill in all fields.");
      if (password.length < 4) throw new Error("Password must be at least 4 characters.");
      if (password !== confirm) throw new Error("Passwords do not match.");
      const { token, user } = await api.register(username.trim(), password, displayName.trim());
      if (remember) setToken(token); else setToken(null);
      onAuthed(user, token);
    } catch (err) {
      setError(err.message || "Could not create the account.");
    }
    setBusy(false);
  }

  return (
    <div className="login-screen">
      <div className="login-top">
        <Emblem />
        <h1 className="login-brand">KDP, PATAN</h1>
        <p className="login-dept">Department of Computer Engineering</p>
        <p className="login-sub">CO &amp; PO/PSO Attainment System</p>
      </div>

      <div className="login-card">
        <h2>{mode === "login" ? "Welcome back" : "Create your account"}</h2>
        <p className="login-hint">
          {mode === "login" ? "Sign in to your account to continue" : "Register a faculty account to get started"}
        </p>

        {error && <div className="form-error"><AlertTriangle size={14} />{error}</div>}

        <form onSubmit={mode === "login" ? handleLogin : handleRegister}>
          {mode === "register" && (
            <Field label="Full Name">
              <div className="input-icon">
                <User size={16} />
                <input value={displayName} onChange={(e) => setDisplayName(e.target.value)} placeholder="e.g. Prof. A. Sharma" />
              </div>
            </Field>
          )}
          <Field label="Username">
            <div className="input-icon">
              <User size={16} />
              <input value={username} onChange={(e) => setUsername(e.target.value)} placeholder="Enter your username" autoCapitalize="off" />
            </div>
          </Field>
          <Field label="Password">
            <div className="input-icon">
              <Lock size={16} />
              <input type={showPw ? "text" : "password"} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="Enter your password" />
              <button type="button" className="icon-btn" onClick={() => setShowPw((v) => !v)}>
                {showPw ? <EyeOff size={16} /> : <Eye size={16} />}
              </button>
            </div>
          </Field>
          {mode === "register" && (
            <Field label="Confirm Password">
              <div className="input-icon">
                <Lock size={16} />
                <input type={showPw ? "text" : "password"} value={confirm} onChange={(e) => setConfirm(e.target.value)} placeholder="Re-enter your password" />
              </div>
            </Field>
          )}
          <label className="remember-row">
            <input type="checkbox" checked={remember} onChange={(e) => setRemember(e.target.checked)} />
            Remember me on this device
          </label>
          <button className="btn-primary btn-block" type="submit" disabled={busy}>
            {busy ? "Please wait\u2026" : mode === "login" ? (<><LogIn size={16} /> Sign In</>) : (<><User size={16} /> Create Account</>)}
          </button>
        </form>

        <p className="switch-mode">
          {mode === "login" ? (
            <>Don&apos;t have an account? <button onClick={() => { setMode("register"); setError(""); }}>Register</button></>
          ) : (
            <>Already have an account? <button onClick={() => { setMode("login"); setError(""); }}>Sign in</button></>
          )}
        </p>
      </div>
      <p className="login-footer">
        Your account and saved courses are stored in this system's own database, private to your login.
      </p>
    </div>
  );
}

/* ============================================================================
   ASSESSMENT / SURVEY / MAPPING / SETUP / RESULTS TABS
============================================================================ */
function AssessmentTab({ title, note, data, nQ, coForQ, stats, onChange }) {
  function setMax(i, val) {
    const newMax = clamp(val, 0, 100);
    const mm = [...data.maxMarks];
    mm[i] = newMax;
    // If the max was lowered below marks already entered for this question,
    // pull those marks back down too so a student's score can never exceed
    // the question's own max marks.
    const students = data.students.map((s) => ({
      ...s,
      marks: s.marks.map((m, j) => (j === i ? clamp(m, 0, newMax) : m)),
    }));
    onChange({ ...data, maxMarks: mm, students });
  }
  function setMark(rowIdx, qIdx, val) {
    const bounded = clamp(val, 0, data.maxMarks[qIdx] || 0);
    const students = data.students.map((s, i) => i === rowIdx ? { ...s, marks: s.marks.map((m, j) => (j === qIdx ? bounded : m)) } : s);
    onChange({ ...data, students });
  }
  function setField(rowIdx, key, val) {
    onChange({ ...data, students: data.students.map((s, i) => (i === rowIdx ? { ...s, [key]: val } : s)) });
  }
  function addStudent() {
    const n = data.students.length + 1;
    onChange({ ...data, students: [...data.students, { roll: `R${String(n).padStart(2, "0")}`, name: "", marks: Array(nQ).fill(0) }] });
  }
  function removeStudent(idx) { onChange({ ...data, students: data.students.filter((_, i) => i !== idx) }); }

  return (
    <div>
      <div className="panel-head"><h3>{title}</h3>{note && <p className="muted">{note}</p>}</div>
      <div className="table-scroll">
        <table className="grid-table">
          <thead>
            <tr>
              <th className="sticky-col">Roll No</th>
              <th className="sticky-col2">Name</th>
              {Array.from({ length: nQ }).map((_, i) => (<th key={i}>Q{i + 1}<div className="th-sub">{coForQ(i)}</div></th>))}
              <th>Total</th><th></th>
            </tr>
            <tr className="max-row">
              <td className="sticky-col">Max Marks →</td>
              <td className="sticky-col2"></td>
              {data.maxMarks.map((m, i) => (
                <td key={i}><input type="number" min={0} max={100} className="cell-input" value={m} onChange={(e) => setMax(i, Number(e.target.value))} /></td>
              ))}
              <td colSpan={2}></td>
            </tr>
          </thead>
          <tbody>
            {data.students.map((s, rIdx) => {
              const total = s.marks.reduce((a, b) => a + (Number(b) || 0), 0);
              return (
                <tr key={rIdx}>
                  <td className="sticky-col"><input className="cell-input" value={s.roll} onChange={(e) => setField(rIdx, "roll", e.target.value)} /></td>
                  <td className="sticky-col2"><input className="cell-input cell-input-left" value={s.name} onChange={(e) => setField(rIdx, "name", e.target.value)} /></td>
                  {s.marks.map((m, qIdx) => (
                    <td key={qIdx}><input type="number" min={0} max={data.maxMarks[qIdx] || 0} className="cell-input" value={m} onChange={(e) => setMark(rIdx, qIdx, Number(e.target.value))} /></td>
                  ))}
                  <td className="total-cell">{total}</td>
                  <td><button className="icon-btn danger" onClick={() => removeStudent(rIdx)}><Trash2 size={14} /></button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <button className="btn-ghost" onClick={addStudent}><Plus size={14} /> Add student</button>

      <div className="panel-head" style={{ marginTop: 22 }}><h3>CO-wise Attainment Summary</h3></div>
      <div className="table-scroll">
        <table className="grid-table summary-table">
          <thead><tr><th>CO</th><th>Max Marks</th><th>Target Marks</th><th>Attained</th><th>Total</th><th>%</th><th>Level</th></tr></thead>
          <tbody>
            {stats.map((st, i) => (
              <tr key={i}>
                <td className="co-badge">{COs[i]}</td><td>{st.max}</td><td>{st.targetMarks.toFixed(1)}</td>
                <td>{st.attained}</td><td>{st.total}</td><td>{(st.pct * 100).toFixed(1)}%</td>
                <td><span className={`level-badge level-${st.level}`}>{st.level}</span></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function SurveyTab({ survey, onChange, avg }) {
  function setRating(rowIdx, coIdx, val) {
    const students = survey.students.map((s, i) => i === rowIdx ? { ...s, ratings: s.ratings.map((r, j) => (j === coIdx ? val : r)) } : s);
    onChange({ ...survey, students });
  }
  function setField(rowIdx, key, val) { onChange({ ...survey, students: survey.students.map((s, i) => (i === rowIdx ? { ...s, [key]: val } : s)) }); }
  function addStudent() {
    const n = survey.students.length + 1;
    onChange({ ...survey, students: [...survey.students, { roll: `R${String(n).padStart(2, "0")}`, name: "", ratings: Array(6).fill(2) }] });
  }
  function removeStudent(idx) { onChange({ ...survey, students: survey.students.filter((_, i) => i !== idx) }); }
  return (
    <div>
      <div className="panel-head"><h3>Course Exit Survey — Indirect CO Attainment</h3><p className="muted">Student self-rating per CO. Scale: 1 = Low, 2 = Medium, 3 = High.</p></div>
      <div className="table-scroll">
        <table className="grid-table">
          <thead><tr><th className="sticky-col">Roll No</th><th className="sticky-col2">Name</th>{COs.map((co) => <th key={co}>{co}</th>)}<th></th></tr></thead>
          <tbody>
            {survey.students.map((s, rIdx) => (
              <tr key={rIdx}>
                <td className="sticky-col"><input className="cell-input" value={s.roll} onChange={(e) => setField(rIdx, "roll", e.target.value)} /></td>
                <td className="sticky-col2"><input className="cell-input cell-input-left" value={s.name} onChange={(e) => setField(rIdx, "name", e.target.value)} /></td>
                {s.ratings.map((r, cIdx) => (
                  <td key={cIdx}>
                    <select className="cell-input" value={r} onChange={(e) => setRating(rIdx, cIdx, Number(e.target.value))}>
                      <option value={1}>1</option><option value={2}>2</option><option value={3}>3</option>
                    </select>
                  </td>
                ))}
                <td><button className="icon-btn danger" onClick={() => removeStudent(rIdx)}><Trash2 size={14} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <button className="btn-ghost" onClick={addStudent}><Plus size={14} /> Add student</button>
      <div className="panel-head" style={{ marginTop: 22 }}><h3>Average Rating per CO (Indirect Attainment Level)</h3></div>
      <div className="table-scroll">
        <table className="grid-table summary-table">
          <thead><tr>{COs.map((co) => <th key={co}>{co}</th>)}</tr></thead>
          <tbody><tr>{avg.map((v, i) => <td key={i}>{v.toFixed(2)}</td>)}</tr></tbody>
        </table>
      </div>
    </div>
  );
}

function MappingTab({ mapping, targetPO, onMapping, onTarget }) {
  function setCell(i, j, val) { const m = mapping.map((row) => [...row]); m[i][j] = val; onMapping(m); }
  function setTarget(j, val) { const t = [...targetPO]; t[j] = clamp(val, 0, 3); onTarget(t); }
  return (
    <div>
      <div className="panel-head"><h3>CO – PO / PSO Mapping Matrix</h3><p className="muted">Mapping strength: 0 = none, 1 = Low, 2 = Medium, 3 = High.</p></div>
      <div className="table-scroll">
        <table className="grid-table">
          <thead><tr><th className="sticky-col">CO \\ PO</th>{POPSO.map((p) => <th key={p}>{p}</th>)}</tr></thead>
          <tbody>
            {COs.map((co, i) => (
              <tr key={co}>
                <td className="sticky-col co-badge">{co}</td>
                {POPSO.map((_, j) => (
                  <td key={j}>
                    <select className="cell-input" value={mapping[i][j]} onChange={(e) => setCell(i, j, Number(e.target.value))}>
                      <option value={0}>– (no mapping)</option><option value={1}>1</option><option value={2}>2</option><option value={3}>3</option>
                    </select>
                  </td>
                ))}
              </tr>
            ))}
            <tr>
              <td className="sticky-col co-badge">Target</td>
              {POPSO.map((_, j) => (<td key={j}><input type="number" step="0.1" min={0} max={3} className="cell-input" value={targetPO[j]} onChange={(e) => setTarget(j, Number(e.target.value))} /></td>))}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

function SetupTab({ state, setState }) {
  const ci = state.courseInfo;
  function setCI(key, val) { setState((s) => ({ ...s, courseInfo: { ...s.courseInfo, [key]: val } })); }
  function setCO(i, val) { setState((s) => { const arr = [...s.coStatements]; arr[i] = val; return { ...s, coStatements: arr }; }); }
  function setW(key, pct) {
    const clamped = clamp(pct, 0, 100);
    setState((s) => ({ ...s, weights: { ...s.weights, [key]: clamped / 100 } }));
  }
  function setT(key, val) {
    const bounded = key === "coTargetLevel" ? clamp(val, 0, 3) : clamp(val, 0, 100) / 100;
    setState((s) => ({ ...s, targets: { ...s.targets, [key]: bounded } }));
  }
  return (
    <div className="setup-grid">
      <div className="panel">
        <h3><ClipboardList size={16} /> Course Details</h3>
        <div className="form-grid">
          <Field label="Institute Name"><input value={ci.institute} onChange={(e) => setCI("institute", e.target.value)} /></Field>
          <Field label="Department"><input value={ci.department} onChange={(e) => setCI("department", e.target.value)} /></Field>
          <Field label="Course Name"><input value={ci.courseName} onChange={(e) => setCI("courseName", e.target.value)} /></Field>
          <Field label="Course Code"><input value={ci.courseCode} onChange={(e) => setCI("courseCode", e.target.value)} /></Field>
          <Field label="Semester"><input value={ci.semester} onChange={(e) => setCI("semester", e.target.value)} /></Field>
          <Field label="Academic Year"><input value={ci.year} onChange={(e) => setCI("year", e.target.value)} /></Field>
          <Field label="Faculty Name"><input value={ci.faculty} onChange={(e) => setCI("faculty", e.target.value)} /></Field>
        </div>
      </div>
      <div className="panel">
        <h3><BookOpenCheck size={16} /> Course Outcomes</h3>
        {COs.map((co, i) => (
          <Field key={co} label={co}><textarea rows={2} value={state.coStatements[i]} onChange={(e) => setCO(i, e.target.value)} /></Field>
        ))}
      </div>
      <div className="panel">
        <h3><Settings2 size={16} /> Weightages &amp; Targets</h3>
        <div className="form-grid">
          <Field label="Internal Weight (for Direct Attainment)"><input type="number" min={0} max={100} value={Math.round(state.weights.internal * 100)} onChange={(e) => setW("internal", Number(e.target.value))} />%</Field>
          <Field label="End-Sem Weight (for Direct Attainment)"><input type="number" min={0} max={100} value={Math.round(state.weights.endsem * 100)} onChange={(e) => setW("endsem", Number(e.target.value))} />%</Field>
          <Field label="Direct Attainment Weight (overall)"><input type="number" min={0} max={100} value={Math.round(state.weights.direct * 100)} onChange={(e) => setW("direct", Number(e.target.value))} />%</Field>
          <Field label="Indirect Attainment Weight (overall)"><input type="number" min={0} max={100} value={Math.round(state.weights.indirect * 100)} onChange={(e) => setW("indirect", Number(e.target.value))} />%</Field>
          <Field label="Target % of marks per CO"><input type="number" min={0} max={100} value={Math.round(state.targets.targetPct * 100)} onChange={(e) => setT("targetPct", Number(e.target.value))} />%</Field>
          <Field label="CO Target Level (0-3)"><input type="number" step="0.1" min={0} max={3} value={state.targets.coTargetLevel} onChange={(e) => setT("coTargetLevel", Number(e.target.value))} /></Field>
          <Field label="Level 3 threshold (min % students)"><input type="number" min={0} max={100} value={Math.round(state.targets.level3 * 100)} onChange={(e) => setT("level3", Number(e.target.value))} />%</Field>
          <Field label="Level 2 threshold (min % students)"><input type="number" min={0} max={100} value={Math.round(state.targets.level2 * 100)} onChange={(e) => setT("level2", Number(e.target.value))} />%</Field>
          <Field label="Level 1 threshold (min % students)"><input type="number" min={0} max={100} value={Math.round(state.targets.level1 * 100)} onChange={(e) => setT("level1", Number(e.target.value))} />%</Field>
        </div>
      </div>
    </div>
  );
}

function ResultsTab({ state, direct, indirect, final, poAttain }) {
  const coAttained = final.filter((v) => v >= state.targets.coTargetLevel).length;
  const poAttained = poAttain.filter((v, j) => v != null && v >= state.targetPO[j]).length;
  const poCount = poAttain.filter((v) => v != null).length;
  const coChartData = COs.map((co, i) => ({ name: co, Attained: Number(final[i].toFixed(2)), Target: state.targets.coTargetLevel }));
  const poChartData = POPSO.map((p, j) => ({ name: p, Attained: poAttain[j] == null ? 0 : Number(poAttain[j].toFixed(2)), Target: state.targetPO[j] }));

  return (
    <div>
      <div className="rings-row">
        <Ring value={coAttained} total={6} label="Course Outcomes Attained" color="#4338ca" />
        <Ring value={poAttained} total={poCount || 14} label="PO / PSO Attained" color="#a21caf" />
      </div>
      <div className="panel-head"><h3>Final CO Attainment</h3></div>
      <div className="table-scroll">
        <table className="grid-table summary-table">
          <thead><tr><th>CO</th><th>Direct</th><th>Indirect</th><th>Overall</th><th>Target</th><th>Gap</th><th>Status</th></tr></thead>
          <tbody>
            {COs.map((co, i) => (
              <tr key={co}>
                <td className="co-badge">{co}</td><td>{direct[i].toFixed(2)}</td><td>{indirect[i].toFixed(2)}</td>
                <td><strong>{final[i].toFixed(2)}</strong></td><td>{state.targets.coTargetLevel.toFixed(2)}</td>
                <td>{(final[i] - state.targets.coTargetLevel).toFixed(2)}</td>
                <td><StatusPill ok={final[i] >= state.targets.coTargetLevel} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="chart-box">
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={coChartData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#eceefa" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis domain={[0, 3]} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Bar dataKey="Attained" radius={[4, 4, 0, 0]}>
              {coChartData.map((d, i) => <Cell key={i} fill={d.Attained >= d.Target ? "#4338ca" : "#c7cbe8"} />)}
            </Bar>
            <ReferenceLine y={state.targets.coTargetLevel} stroke="#a21caf" strokeDasharray="4 4" label={{ value: "Target", position: "right", fontSize: 11, fill: "#a21caf" }} />
          </BarChart>
        </ResponsiveContainer>
      </div>
      <div className="panel-head" style={{ marginTop: 26 }}><h3>PO / PSO Attainment</h3></div>
      <div className="table-scroll">
        <table className="grid-table summary-table">
          <thead><tr><th>PO/PSO</th><th>Attained</th><th>Target</th><th>Gap</th><th>Status</th></tr></thead>
          <tbody>
            {POPSO.map((p, j) => (
              <tr key={p}>
                <td className="co-badge">{p}</td>
                <td>{poAttain[j] == null ? "\u2013" : poAttain[j].toFixed(2)}</td>
                <td>{state.targetPO[j].toFixed(2)}</td>
                <td>{poAttain[j] == null ? "\u2013" : (poAttain[j] - state.targetPO[j]).toFixed(2)}</td>
                <td>{poAttain[j] == null ? <span className="pill">No mapping</span> : <StatusPill ok={poAttain[j] >= state.targetPO[j]} />}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="chart-box">
        <ResponsiveContainer width="100%" height={280}>
          <BarChart data={poChartData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#eceefa" />
            <XAxis dataKey="name" tick={{ fontSize: 11 }} interval={0} angle={-35} textAnchor="end" height={50} />
            <YAxis domain={[0, 3]} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Bar dataKey="Attained" radius={[4, 4, 0, 0]}>
              {poChartData.map((d, i) => <Cell key={i} fill={d.Attained >= d.Target ? "#a21caf" : "#e6cdee"} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

/* ============================================================================
   MAIN CALCULATOR APP (post-login)
============================================================================ */
const TABS = [
  { key: "setup", label: "Setup", icon: Settings2 },
  { key: "mapping", label: "CO-PO Mapping", icon: GitBranch },
  { key: "internal1", label: "Internal 1", icon: ClipboardList },
  { key: "internal2", label: "Internal 2", icon: ClipboardList },
  { key: "assignment", label: "Assignment", icon: ClipboardList },
  { key: "endsem", label: "End Sem", icon: ClipboardList },
  { key: "survey", label: "Exit Survey", icon: Users2 },
  { key: "results", label: "Results", icon: BarChart3 },
];

function Calculator({ user, onLogout }) {
  const [state, setState] = useState(sampleState());
  const [tab, setTab] = useState("setup");
  const [toast, setToast] = useState(null);
  const [warn, setWarn] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [myCourses, setMyCourses] = useState([]);
  const [currentId, setCurrentId] = useState(null);
  const [showCourses, setShowCourses] = useState(false);
  const fileRef = useRef(null);

  async function refreshCourseList() {
    try {
      const { courses } = await api.listCourses();
      setMyCourses(courses);
      return courses;
    } catch { return []; }
  }

  useEffect(() => {
    (async () => {
      const courses = await refreshCourseList();
      if (courses.length) {
        try {
          const full = await api.getCourse(courses[0].id);
          setState(normalizeState(full.data));
          setCurrentId(full.id);
        } catch { /* fall back to sample */ }
      }
      setLoaded(true);
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function showToast(msg, ms = 2600) { setToast(msg); setTimeout(() => setToast(null), ms); }

  async function handleSave() {
    try {
      const code = state.courseInfo.courseCode || "untitled";
      const saved = await api.upsertCourse(code, state.courseInfo.courseName, state);
      setCurrentId(saved.id);
      await refreshCourseList();
      showToast("Saved to database \u2713");
    } catch (err) {
      showToast(`Save failed: ${err.message}`);
    }
  }
  async function handleLoadCourse(id) {
    try {
      const full = await api.getCourse(id);
      setState(normalizeState(full.data));
      setCurrentId(full.id);
      setShowCourses(false);
      showToast(`Loaded ${full.courseCode}`);
    } catch (err) {
      showToast(`Load failed: ${err.message}`);
    }
  }
  async function handleDeleteCourse(id, e) {
    e.stopPropagation();
    if (!window.confirm("Delete this saved course? This cannot be undone.")) return;
    try {
      await api.deleteCourse(id);
      await refreshCourseList();
      if (id === currentId) setCurrentId(null);
      showToast("Course deleted");
    } catch (err) {
      showToast(`Delete failed: ${err.message}`);
    }
  }
  function handleLoadSample() { setState(sampleState()); setCurrentId(null); showToast("Sample data loaded (unsaved)"); }
  function handleBlank() { setState(blankState()); setCurrentId(null); showToast("Blank template loaded (unsaved)"); }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(new Uint8Array(evt.target.result), { type: "array" });
        const { next, found, missing } = parseWorkbook(wb, state);
        setState(normalizeState(next));
        setCurrentId(null);
        setWarn(`Imported: ${found.join(", ") || "none"}.${missing.length ? " Not found (kept existing): " + missing.join(", ") + "." : ""}`);
        showToast("Excel file imported (not yet saved)");
      } catch {
        setWarn("Could not read that file. Make sure it is a .xlsx workbook exported from this system.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  }

  const targets = state.targets;
  const i1Stats = useMemo(() => simpleCOStats(state.internal1, targets), [state.internal1, targets]);
  const i2Stats = useMemo(() => simpleCOStats(state.internal2, targets), [state.internal2, targets]);
  const asgStats = useMemo(() => simpleCOStats(state.assignment, targets), [state.assignment, targets]);
  const esStats = useMemo(() => endsemCOStats(state.endsem, targets), [state.endsem, targets]);
  const indirect = useMemo(() => surveyCOAverage(state.survey), [state.survey]);
  const direct = useMemo(
    () => COs.map((_, i) => {
      const internalAvg = (i1Stats[i].level + i2Stats[i].level + asgStats[i].level) / 3;
      return internalAvg * state.weights.internal + esStats[i].level * state.weights.endsem;
    }),
    [i1Stats, i2Stats, asgStats, esStats, state.weights]
  );
  const final = useMemo(
    () => COs.map((_, i) => direct[i] * state.weights.direct + indirect[i] * state.weights.indirect),
    [direct, indirect, state.weights]
  );
  const poAttain = useMemo(
    () => POPSO.map((_, j) => {
      let num = 0, den = 0;
      for (let i = 0; i < 6; i++) {
        const w = state.mapping[i][j] || 0;
        if (w > 0) { num += final[i] * w; den += w; }
      }
      return den > 0 ? num / den : null;
    }),
    [state.mapping, final]
  );

  if (!loaded) return <div className="loading-screen">Loading your data…</div>;

  return (
    <div className="app-shell">
      <header className="app-header">
        <div className="app-header-left">
          <div className="brand-mini"><GraduationCap size={20} /></div>
          <div>
            <div className="app-title">KDP, Patan — CO/PO Attainment System</div>
            <div className="app-subtitle">{state.courseInfo.courseName || "Untitled course"} {state.courseInfo.courseCode ? `\u00b7 ${state.courseInfo.courseCode}` : ""}</div>
          </div>
        </div>
        <div className="app-header-right">
          <span className="user-chip"><User size={14} /> {user.displayName}</span>
          <div className="courses-dropdown-wrap">
            <button className="btn-ghost" onClick={() => setShowCourses((v) => !v)}><FolderOpen size={14} /> My Courses ({myCourses.length})</button>
            {showCourses && (
              <div className="courses-dropdown">
                {myCourses.length === 0 && <div className="courses-empty">No saved courses yet.</div>}
                {myCourses.map((c) => (
                  <div key={c.id} className={`course-row ${c.id === currentId ? "course-row-active" : ""}`} onClick={() => handleLoadCourse(c.id)}>
                    <div>
                      <div className="course-row-name">{c.courseName || c.courseCode}</div>
                      <div className="course-row-code">{c.courseCode}</div>
                    </div>
                    <button className="icon-btn danger" onClick={(e) => handleDeleteCourse(c.id, e)}><Trash2 size={13} /></button>
                  </div>
                ))}
              </div>
            )}
          </div>
          <button className="btn-ghost" onClick={handleLoadSample}><RefreshCw size={14} /> Sample</button>
          <button className="btn-ghost" onClick={handleBlank}>Blank</button>
          <input type="file" accept=".xlsx,.xls" ref={fileRef} style={{ display: "none" }} onChange={handleFile} />
          <button className="btn-ghost" onClick={() => fileRef.current.click()}><UploadCloud size={14} /> Upload Excel</button>
          <button className="btn-ghost" onClick={() => exportWorkbook(state)}><Download size={14} /> Export Excel</button>
          <button className="btn-primary" onClick={handleSave}><Save size={14} /> Save</button>
          <button className="icon-btn" onClick={onLogout} title="Log out"><LogOut size={16} /></button>
        </div>
      </header>

      {warn && <div className="warn-banner"><AlertTriangle size={14} /> {warn}<button className="icon-btn" onClick={() => setWarn(null)}><X size={14} /></button></div>}

      <nav className="tab-nav">
        {TABS.map((t) => {
          const Icon = t.icon;
          return <button key={t.key} className={`tab-btn ${tab === t.key ? "tab-active" : ""}`} onClick={() => setTab(t.key)}><Icon size={15} /> {t.label}</button>;
        })}
      </nav>

      <main className="app-main">
        {tab === "setup" && <SetupTab state={state} setState={setState} />}
        {tab === "mapping" && (
          <MappingTab mapping={state.mapping} targetPO={state.targetPO}
            onMapping={(m) => setState((s) => ({ ...s, mapping: m }))}
            onTarget={(t) => setState((s) => ({ ...s, targetPO: t }))} />
        )}
        {tab === "internal1" && (
          <AssessmentTab title="Internal Exam 1" note="Each question tests one CO (Q1\u2192CO1 \u2026 Q6\u2192CO6)."
            data={state.internal1} nQ={6} coForQ={(i) => COs[i]} stats={i1Stats}
            onChange={(d) => setState((s) => ({ ...s, internal1: d }))} />
        )}
        {tab === "internal2" && (
          <AssessmentTab title="Internal Exam 2" note="Each question tests one CO (Q1\u2192CO1 \u2026 Q6\u2192CO6)."
            data={state.internal2} nQ={6} coForQ={(i) => COs[i]} stats={i2Stats}
            onChange={(d) => setState((s) => ({ ...s, internal2: d }))} />
        )}
        {tab === "assignment" && (
          <AssessmentTab title="Assignment" note="Each task tests one CO (Q1\u2192CO1 \u2026 Q6\u2192CO6)."
            data={state.assignment} nQ={6} coForQ={(i) => COs[i]} stats={asgStats}
            onChange={(d) => setState((s) => ({ ...s, assignment: d }))} />
        )}
        {tab === "endsem" && (
          <AssessmentTab title="End Semester Exam" note="Q1\u2013Q6 = Part A, Q7\u2013Q12 = Part B. Qi and Q(i+6) both test COi."
            data={state.endsem} nQ={12} coForQ={(i) => COs[i % 6]} stats={esStats}
            onChange={(d) => setState((s) => ({ ...s, endsem: d }))} />
        )}
        {tab === "survey" && <SurveyTab survey={state.survey} avg={indirect} onChange={(sv) => setState((s) => ({ ...s, survey: sv }))} />}
        {tab === "results" && <ResultsTab state={state} direct={direct} indirect={indirect} final={final} poAttain={poAttain} />}
      </main>

      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}

/* ============================================================================
   ROOT: AUTH GATE
============================================================================ */
export default function App() {
  const [user, setUser] = useState(null);
  const [checking, setChecking] = useState(true);

  useEffect(() => {
    (async () => {
      const stored = getStoredToken();
      if (stored) {
        try {
          setToken(stored);
          const { user: u } = await api.me();
          setUser(u);
        } catch {
          setToken(null);
        }
      }
      setChecking(false);
    })();
  }, []);

  function handleAuthed(u) { setUser(u); }
  function handleLogout() { setToken(null); setUser(null); }

  return (
    <div className="root">
      <style>{CSS}</style>
      {checking ? (
        <div className="loading-screen">Loading…</div>
      ) : user ? (
        <Calculator user={user} onLogout={handleLogout} />
      ) : (
        <LoginPage onAuthed={handleAuthed} />
      )}
    </div>
  );
}

/* ============================================================================
   STYLES
============================================================================ */
const CSS = `
  .root { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; color: #1f2333; background: #f4f5fb; min-height: 100vh; }
  * { box-sizing: border-box; }
  input, textarea, select, button { font-family: inherit; }

  .login-screen { min-height: 100vh; display: flex; flex-direction: column; align-items: center; padding: 48px 20px 28px; background: linear-gradient(180deg, #eef1fb 0%, #f6f7fc 55%, #f4f5fb 100%); }
  .login-top { display:flex; flex-direction: column; align-items: center; margin-bottom: 26px; }
  .login-brand { font-size: 34px; font-weight: 800; color: #3730a3; margin: 14px 0 4px; letter-spacing: 0.5px; }
  .login-dept { font-size: 17px; color: #55597a; font-weight: 600; margin: 0; }
  .login-sub { font-size: 13px; color: #8b8fb0; margin: 6px 0 0; }
  .login-card { width: 100%; max-width: 420px; background: #fff; border-radius: 20px; padding: 34px 30px; box-shadow: 0 10px 40px rgba(60,60,120,0.10); }
  .login-card h2 { margin: 0 0 4px; font-size: 24px; font-weight: 800; color: #14162a; }
  .login-hint { color: #6b6f8f; font-size: 14px; margin: 0 0 22px; }
  .form-error { display:flex; align-items:center; gap:8px; font-size: 12.5px; padding: 10px 12px; border-radius: 10px; margin-bottom: 16px; background: #fdecec; color: #a3241d; }
  .field { display:block; margin-bottom: 16px; }
  .field-label { display:block; font-size: 13px; font-weight: 700; color: #23263a; margin-bottom: 6px; }
  .input-icon { display:flex; align-items:center; gap:8px; border: 1.5px solid #e3e5f0; border-radius: 12px; padding: 12px 14px; background:#fafbff; }
  .input-icon svg { color: #9296b8; flex-shrink:0; }
  .input-icon input { border:none; outline:none; background:transparent; width:100%; font-size:14px; color:#1f2333; }
  .remember-row { display:flex; align-items:center; gap:8px; font-size: 13.5px; color:#4b4f6b; margin: 4px 0 20px; }
  .btn-primary { background: linear-gradient(135deg, #4338ca, #6366f1); color:#fff; border:none; border-radius: 12px; padding: 13px 18px; font-weight:700; font-size:14.5px; display:flex; align-items:center; justify-content:center; gap:8px; cursor:pointer; }
  .btn-primary:hover { filter: brightness(1.05); }
  .btn-block { width: 100%; }
  .switch-mode { text-align:center; margin: 18px 0 0; font-size: 13.5px; color:#6b6f8f; }
  .switch-mode button { background:none; border:none; color:#4338ca; font-weight:700; cursor:pointer; padding:0; }
  .login-footer { max-width: 440px; text-align:center; font-size: 11.5px; color: #9296b8; margin-top: 22px; line-height:1.5; }

  .loading-screen { min-height:100vh; display:flex; align-items:center; justify-content:center; color:#6b6f8f; font-size:15px; }
  .app-shell { min-height: 100vh; display:flex; flex-direction:column; }
  .app-header { display:flex; align-items:center; justify-content:space-between; padding: 14px 22px; background:#fff; border-bottom: 1px solid #e9eaf4; flex-wrap:wrap; gap:10px; position: sticky; top:0; z-index: 10; }
  .app-header-left { display:flex; align-items:center; gap:12px; }
  .brand-mini { width:38px; height:38px; border-radius:11px; background: linear-gradient(135deg,#4338ca,#6366f1); color:#fff; display:flex; align-items:center; justify-content:center; }
  .app-title { font-weight:800; font-size:15px; color:#14162a; }
  .app-subtitle { font-size:12px; color:#8b8fb0; margin-top:1px; }
  .app-header-right { display:flex; align-items:center; gap:8px; flex-wrap:wrap; position:relative; }
  .user-chip { display:flex; align-items:center; gap:6px; background:#f1f2fb; color:#3d3f63; padding:7px 12px; border-radius:999px; font-size:12.5px; font-weight:600; }
  .btn-ghost { display:flex; align-items:center; gap:6px; background:#f4f5fb; border:1px solid #e6e8f4; color:#3d3f63; border-radius:10px; padding:8px 12px; font-size:12.5px; font-weight:600; cursor:pointer; }
  .btn-ghost:hover { background:#eceefa; }
  .icon-btn { background:none; border:none; cursor:pointer; color:#6b6f8f; padding:6px; border-radius:8px; display:flex; align-items:center; }
  .icon-btn:hover { background:#f1f2fb; }
  .icon-btn.danger:hover { background:#fdecec; color:#a3241d; }

  .courses-dropdown-wrap { position:relative; }
  .courses-dropdown { position:absolute; top:calc(100% + 6px); right:0; background:#fff; border:1px solid #e9eaf4; border-radius:12px; box-shadow:0 10px 30px rgba(30,30,70,0.12); min-width:230px; max-height:280px; overflow-y:auto; z-index:30; padding:6px; }
  .courses-empty { padding:14px; font-size:12.5px; color:#a5a9c9; text-align:center; }
  .course-row { display:flex; align-items:center; justify-content:space-between; padding:9px 10px; border-radius:8px; cursor:pointer; }
  .course-row:hover { background:#f4f5fb; }
  .course-row-active { background:#eef0fd; }
  .course-row-name { font-size:13px; font-weight:700; color:#23263a; }
  .course-row-code { font-size:11px; color:#9296b8; }

  .warn-banner { display:flex; align-items:center; gap:8px; background:#fff7e6; color:#92620a; font-size:13px; padding:10px 22px; }
  .warn-banner .icon-btn { margin-left:auto; }

  .tab-nav { display:flex; gap:6px; padding: 12px 22px 0; overflow-x:auto; background:#fff; border-bottom:1px solid #e9eaf4; }
  .tab-btn { display:flex; align-items:center; gap:6px; background:none; border:none; padding:10px 14px; font-size:13px; font-weight:600; color:#8b8fb0; cursor:pointer; border-bottom: 2.5px solid transparent; white-space:nowrap; }
  .tab-btn.tab-active { color:#4338ca; border-bottom-color:#4338ca; }

  .app-main { flex:1; padding: 22px; max-width: 1180px; margin: 0 auto; width:100%; }

  .panel { background:#fff; border:1px solid #eceefa; border-radius:16px; padding:20px 22px; margin-bottom:18px; }
  .panel h3 { display:flex; align-items:center; gap:8px; font-size:15px; margin:0 0 14px; color:#14162a; }
  .setup-grid { display:grid; grid-template-columns: 1fr; gap: 4px; }
  .form-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(220px,1fr)); gap: 14px 18px; }
  .form-grid .field input { width:100%; }

  .panel input, .panel textarea, .field input, .field textarea { border:1.5px solid #e3e5f0; border-radius:9px; padding:9px 11px; font-size:13.5px; color:#1f2333; width:100%; outline:none; }
  .panel input:focus, .panel textarea:focus { border-color:#a5abf0; }

  .panel-head { margin: 4px 0 12px; }
  .panel-head h3 { font-size:15px; margin:0 0 3px; color:#14162a; }
  .muted { color:#8b8fb0; font-size:12.5px; margin:0; }

  .table-scroll { overflow-x:auto; border:1px solid #eceefa; border-radius:14px; }
  .grid-table { border-collapse:collapse; width:100%; font-size:12.5px; }
  .grid-table th, .grid-table td { padding:8px 9px; text-align:center; border-bottom:1px solid #f1f2fa; white-space:nowrap; }
  .grid-table thead th { background:#f7f8fd; color:#4b4f6b; font-weight:700; font-size:11.5px; }
  .th-sub { font-size:10px; font-weight:500; color:#a5a9c9; }
  .max-row td { background:#fbfbfe; color:#8b8fb0; font-size:11.5px; }
  .sticky-col, .sticky-col2 { position:sticky; background:#fff; z-index:1; text-align:left; }
  .sticky-col { left:0; min-width:70px; }
  .sticky-col2 { left:70px; min-width:110px; }
  thead .sticky-col, thead .sticky-col2 { background:#f7f8fd; z-index:2; }
  .cell-input { width:52px; text-align:center; border:1px solid #e6e8f4; border-radius:6px; padding:5px 4px; font-size:12.5px; }
  .cell-input-left { width:100px; text-align:left; }
  .total-cell { font-weight:700; color:#4338ca; }
  .co-badge { font-weight:700; color:#3730a3; background:#eef0fd; }
  .level-badge { display:inline-block; min-width:22px; padding:2px 7px; border-radius:999px; font-weight:700; }
  .level-3 { background:#e3f7ea; color:#1e7d4b; }
  .level-2 { background:#eaf3ff; color:#2058b0; }
  .level-1 { background:#fff4da; color:#92620a; }
  .level-0 { background:#fdecec; color:#a3241d; }

  .summary-table th, .summary-table td { text-align:center; }

  .rings-row { display:flex; gap:32px; justify-content:center; margin: 6px 0 30px; flex-wrap:wrap; }
  .ring-wrap { display:flex; flex-direction:column; align-items:center; }
  .ring-num { font-size:15px; font-weight:800; fill:#14162a; }
  .ring-sub { font-size:9px; fill:#a5a9c9; }
  .ring-label { margin-top:8px; font-size:13px; font-weight:700; color:#3d3f63; text-align:center; max-width:140px; }

  .pill { display:inline-flex; align-items:center; gap:5px; padding:4px 10px; border-radius:999px; font-size:11.5px; font-weight:700; background:#f1f2fb; color:#6b6f8f; }
  .pill-ok { background:#e3f7ea; color:#1e7d4b; }
  .pill-bad { background:#fdecec; color:#a3241d; }

  .chart-box { background:#fff; border:1px solid #eceefa; border-radius:14px; padding:14px 8px 4px; margin-top:14px; }

  .toast { position:fixed; bottom:24px; left:50%; transform:translateX(-50%); background:#14162a; color:#fff; padding:11px 20px; border-radius:999px; font-size:13px; font-weight:600; box-shadow:0 8px 30px rgba(0,0,0,0.2); z-index:50; }

  @media (min-width: 900px) {
    .setup-grid { grid-template-columns: 1fr 1fr; }
    .setup-grid .panel:first-child { grid-column: 1 / -1; }
  }
`;
