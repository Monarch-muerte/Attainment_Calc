import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx-js-style";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, ReferenceLine, Cell,
} from "recharts";
import {
  LogIn, LogOut, User, Lock, Eye, EyeOff, UploadCloud, Download, Save, Plus, Trash2,
  CheckCircle2, XCircle, GitBranch, ClipboardList, BarChart3,
  RefreshCw, GraduationCap, AlertTriangle, X, Users2, BookOpenCheck, FolderOpen,
  Printer, FileSpreadsheet, Edit3, Target, Calculator as CalcIcon, TrendingUp, Check,
  Search, ShieldCheck, Building2, KeyRound, Sparkles, Award,
} from "lucide-react";
import { api, setToken, getStoredToken } from "./api";

/* ============================================================================
   CONSTANTS & DEFAULT DATA
============================================================================ */
const COs = ["CO1", "CO2", "CO3", "CO4", "CO5", "CO6"];
const DIPLOMA_POS = ["PO1", "PO2", "PO3", "PO4", "PO5", "PO6", "PO7"];
const PSOS = ["PSO1", "PSO2"];
const DIPLOMA_POPSO = [...DIPLOMA_POS, ...PSOS];
const FULL_POS = Array.from({ length: 12 }, (_, i) => `PO${i + 1}`);
const POPSO = [...FULL_POS, ...PSOS];

const DEFAULT_CO_CODES = ["C303_N.1", "C303_N.2", "C303_N.3", "C303_N.4", "C303_N.5", "C303_N.6"];

const DEFAULT_WEIGHTS = { internal: 0.4, endsem: 0.6, direct: 0.8, indirect: 0.2 };
const DEFAULT_TARGETS = { targetPctCO: COs.map(() => 0.61), coTargetLevel: 0.90, level3: 0.7, level2: 0.6, level1: 0.5 };

const DEFAULT_EVAL_PLAN = {
  marks: { midSem: 30, gtu: 70, pa: 25, ese: 25 },
  weights: { midSem: 30, gtu: 70, pa: 50, ese: 50 },
};

const DEFAULT_TARGET_LEVELS = {
  theory: { paTarget: 0.90, paStudents: 48, paMarks: 61, eseTarget: 0.90, eseStudents: 48, eseMarks: 61 },
  practical: { paTarget: 0.90, paStudents: 48, paMarks: 61, eseTarget: 0.90, eseStudents: 48, eseMarks: 61 },
  finalTarget: 0.90,
};

const GTU_GRADES = [
  { grade: "AA (85-100)", avg: 93, isAboveThreshold: true },
  { grade: "AB (75-84)", avg: 80, isAboveThreshold: true },
  { grade: "BB (65-74)", avg: 70, isAboveThreshold: true },
  { grade: "BC (55-64)", avg: 60, isAboveThreshold: false },
  { grade: "CC (45-54)", avg: 50, isAboveThreshold: false },
  { grade: "CD (40-44)", avg: 42, isAboveThreshold: false },
  { grade: "DD (35-39)", avg: 37, isAboveThreshold: false },
  { grade: "FF (00-34)", avg: 17, isAboveThreshold: false },
];

const DEFAULT_TARGET_SETTING = {
  nbaSubjectCode: "C303_N",
  years: ["2021-22", "2020-21", "2019-20"],
  gradeCounts: [
    [274, 573, 847, 919, 423, 25, 0, 543], // Year 1 (2021-22)
    [0, 0, 0, 0, 0, 0, 0, 0],             // Year 2
    [0, 0, 0, 0, 0, 0, 0, 0],             // Year 3
  ],
  targetMarksPct: 61,
  targetStudentsPct: 48,
  targetLevel: 0.90,
  rangeMatrix: [
    { level: "0.0 - 0.9", minStudents: 0, maxStudents: 49.99 },
    { level: "1.0 - 1.9", minStudents: 50, maxStudents: 59.99 },
    { level: "2.0 - 2.9", minStudents: 60, maxStudents: 69.99 },
    { level: "3", minStudents: 70, maxStudents: 100 },
  ],
};

const SAMPLE_PR_PA_STUDENTS = [
  { roll: "186310307091", name: "PRAJAPATI SRUSHTIBEN SURESHBHAI", marks: [2, 9, 3, 1, 1, 0] },
  { roll: "196310307003", name: "BALOCH ARBAJKHAN JANGIRKHAN", marks: [2, 13, 4, 1, 1, 0] },
  { roll: "196310307531", name: "NAI VISHAL BHAVARLAL", marks: [1, 6, 2, 1, 1, 0] },
  { roll: "196310307563", name: "RAVAL NIRMIT PARESHKUMAR", marks: [1, 6, 2, 1, 1, 0] },
  { roll: "206310307004", name: "RAVAL DAX MANSUKHBHAI", marks: [2, 11, 4, 1, 2, 0] },
  { roll: "206310307010", name: "SHUJAATMOHAMMAD HAIDARALI SUNASARA", marks: [2, 11, 4, 1, 2, 0] },
  { roll: "206310307020", name: "PATHAN SHIFAT KAMALKHAN", marks: [2, 11, 4, 1, 2, 0] },
  { roll: "206310307030", name: "BALDANIYA HEMDIP NARESHBHAI", marks: [2, 8, 3, 1, 2, 0] },
  { roll: "206310307045", name: "RAMI SHREY MANISHKUMAR", marks: [2, 8, 3, 2, 1, 0] },
  { roll: "206310307050", name: "PRAJAPATI PARTHKUMAR SURESHKUMAR", marks: [2, 8, 3, 2, 1, 0] },
  { roll: "206310307066", name: "DARJI PRATHAM SUDHIRKUMAR", marks: [1, 4, 2, 2, 1, 0] },
  { roll: "206310307074", name: "DARJI UJAS RATILAL", marks: [1, 6, 2, 2, 1, 0] },
  { roll: "206310307107", name: "PATEL AARGEE RAJESHKUMAR", marks: [2, 12, 4, 2, 2, 0] },
  { roll: "206310307115", name: "VANKAR KAJALBEN PRAVINKUMAR", marks: [2, 8, 3, 2, 2, 0] },
  { roll: "206310307125", name: "PATEL YASH KIRTIBHAI", marks: [2, 5, 2, 2, 2, 0] },
  { roll: "206310307129", name: "BAROT RUSHI NIKESHBHAI", marks: [2, 10, 3, 2, 2, 0] },
  { roll: "206310307140", name: "PATEL YASHKUMAR VIJAYBHAI", marks: [1, 5, 2, 1, 1, 0] },
  { roll: "206310307186", name: "BALDANIYA RUTUDHVAJ KANUBHAI", marks: [2, 12, 3, 1, 1, 0] },
  { roll: "206310307193", name: "PANCHAL HASMI VIJAYKUMAR", marks: [2, 10, 3, 1, 1, 0] },
  { roll: "206310307194", name: "Thakur Saloni Nareshbhai", marks: [2, 8, 3, 1, 1, 0] },
  { roll: "206310307195", name: "CHAUDHARY KHUSHI HITESHBHAI", marks: [2, 8, 3, 1, 2, 0] },
  { roll: "206310307198", name: "PATEL TRUPTI NARESHKUMAR", marks: [2, 10, 3, 1, 2, 0] },
];

const SAMPLE = {
  courseInfo: {
    institute: "K.D.POLYTECHNIC,PATAN",
    department: "COMPUTER ENGINEERING DEPARTMENT",
    courseName: "Introduction to Machine Learning",
    courseCode: "4350702",
    semester: "5",
    batch: "2021-24",
    term: "2023-24 ODD (231)",
    numStudents: 137,
    faculty: "Prof. A. Sharma",
    year: "2023-24",
  },
  numCos: 5,
  coCodes: ["C303_N.1", "C303_N.2", "C303_N.3", "C303_N.4", "C303_N.5", "C303_N.6"],
  coStatements: [
    "Describe basic concept of machine learning and its applications",
    "Practice Numpy, Pandas, Matplotlib, sklearn library's inbuilt function required to solve machine learning problems",
    "Use Pandas library for data preprocessing",
    "Apply supervised learning algorithms based on dataset characteristics",
    "Apply unsupervised learning algorithms based on dataset characteristics",
    "Deploy machine learning models and evaluate performance metrics",
  ],
  evalPlan: JSON.parse(JSON.stringify(DEFAULT_EVAL_PLAN)),
  targetLevels: JSON.parse(JSON.stringify(DEFAULT_TARGET_LEVELS)),
  targetSetting: JSON.parse(JSON.stringify(DEFAULT_TARGET_SETTING)),
  mapping: [
    [3, 2, 0, 0, 0, 0, 2, 2, 0],
    [3, 2, 0, 3, 0, 0, 2, 3, 0],
    [3, 3, 0, 3, 0, 0, 2, 3, 0],
    [3, 3, 3, 2, 0, 0, 2, 3, 0],
    [3, 3, 3, 2, 0, 0, 2, 3, 0],
    [0, 0, 0, 0, 0, 0, 0, 0, 0],
  ],
  targetPO: [3.0, 2.6, 3.0, 2.5, 2.0, 2.0, 2.0, 2.8, 2.0],
  internal1: {
    maxMarks: [2, 14, 5, 2, 2, 0],
    students: [
      { roll: "186310307091", name: "PRAJAPATI SRUSHTIBEN SURESHBHAI", marks: [2, 9, 3, 1, 1, 0] },
      { roll: "196310307003", name: "BALOCH ARBAJKHAN JANGIRKHAN", marks: [2, 13, 4, 1, 1, 0] },
      { roll: "196310307531", name: "NAI VISHAL BHAVARLAL", marks: [1, 6, 2, 1, 1, 0] },
      { roll: "196310307563", name: "RAVAL NIRMIT PARESHKUMAR", marks: [1, 6, 2, 1, 1, 0] },
      { roll: "206310307004", name: "RAVAL DAX MANSUKHBHAI", marks: [2, 11, 4, 1, 2, 0] },
      { roll: "206310307010", name: "SHUJAATMOHAMMAD HAIDARALI SUNASARA", marks: [2, 11, 4, 1, 2, 0] },
      { roll: "206310307020", name: "PATHAN SHIFAT KAMALKHAN", marks: [2, 11, 4, 1, 2, 0] },
      { roll: "206310307030", name: "BALDANIYA HEMDIP NARESHBHAI", marks: [2, 8, 3, 1, 2, 0] },
      { roll: "206310307045", name: "RAMI SHREY MANISHKUMAR", marks: [2, 8, 3, 2, 1, 0] },
      { roll: "206310307050", name: "PRAJAPATI PARTHKUMAR SURESHKUMAR", marks: [2, 8, 3, 2, 1, 0] },
      { roll: "206310307066", name: "DARJI PRATHAM SUDHIRKUMAR", marks: [1, 4, 2, 2, 1, 0] },
      { roll: "206310307074", name: "DARJI UJAS RATILAL", marks: [1, 6, 2, 2, 1, 0] },
      { roll: "206310307107", name: "PATEL AARGEE RAJESHKUMAR", marks: [2, 12, 4, 2, 2, 0] },
      { roll: "206310307115", name: "VANKAR KAJALBEN PRAVINKUMAR", marks: [2, 8, 3, 2, 2, 0] },
      { roll: "206310307125", name: "PATEL YASH KIRTIBHAI", marks: [2, 5, 2, 2, 2, 0] },
      { roll: "206310307129", name: "BAROT RUSHI NIKESHBHAI", marks: [2, 10, 3, 2, 2, 0] },
      { roll: "206310307140", name: "PATEL YASHKUMAR VIJAYBHAI", marks: [1, 5, 2, 1, 1, 0] },
      { roll: "206310307186", name: "BALDANIYA RUTUDHVAJ KANUBHAI", marks: [2, 12, 3, 1, 1, 0] },
      { roll: "206310307193", name: "PANCHAL HASMI VIJAYKUMAR", marks: [2, 10, 3, 1, 1, 0] },
      { roll: "206310307194", name: "Thakur Saloni Nareshbhai", marks: [2, 8, 3, 1, 1, 0] },
      { roll: "206310307195", name: "CHAUDHARY KHUSHI HITESHBHAI", marks: [2, 8, 3, 1, 2, 0] },
      { roll: "206310307198", name: "PATEL TRUPTI NARESHKUMAR", marks: [2, 10, 3, 1, 2, 0] },
    ],
  },
  internal2: {
    maxMarks: [10, 10, 10, 10, 10, 10],
    students: [
      { roll: "R01", name: "Student 1", marks: [10, 6, 8, 7, 8, 4] },
      { roll: "R02", name: "Student 2", marks: [6, 5, 5, 5, 6, 6] },
      { roll: "R03", name: "Student 3", marks: [5, 8, 6, 1, 9, 6] },
      { roll: "R04", name: "Student 4", marks: [5, 7, 7, 7, 5, 7] },
      { roll: "R05", name: "Student 5", marks: [4, 9, 5, 6, 7, 7] },
      { roll: "R06", name: "Student 6", marks: [6, 8, 0, 6, 6, 6] },
      { roll: "R07", name: "Student 7", marks: [9, 5, 6, 3, 7, 4] },
      { roll: "R08", name: "Student 8", marks: [4, 10, 8, 7, 7, 4] },
      { roll: "R09", name: "Student 9", marks: [5, 7, 3, 7, 3, 7] },
      { roll: "R10", name: "Student 10", marks: [5, 10, 8, 6, 3, 5] },
      { roll: "R11", name: "Student 11", marks: [6, 5, 7, 8, 6, 6] },
      { roll: "R12", name: "Student 12", marks: [8, 6, 8, 6, 10, 6] },
      { roll: "R13", name: "Student 13", marks: [5, 7, 5, 5, 6, 8] },
      { roll: "R14", name: "Student 14", marks: [3, 6, 6, 6, 8, 4] },
      { roll: "R15", name: "Student 15", marks: [8, 6, 7, 6, 6, 6] },
    ],
  },
  assignment: {
    maxMarks: [5, 5, 5, 5, 5, 0],
    students: [
      { roll: "186310307091", name: "PRAJAPATI SRUSHTIBEN SURESHBHAI", marks: [4.8, 4.8, 4.8, 4.8, 4.8, 0] },
      { roll: "196310307003", name: "BALOCH ARBAJKHAN JANGIRKHAN", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "196310307531", name: "NAI VISHAL BHAVARLAL", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "196310307563", name: "RAVAL NIRMIT PARESHKUMAR", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307004", name: "RAVAL DAX MANSUKHBHAI", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307010", name: "SHUJAATMOHAMMAD HAIDARALI SUNASARA", marks: [4.8, 4.8, 4.8, 4.8, 4.8, 0] },
      { roll: "206310307020", name: "PATHAN SHIFAT KAMALKHAN", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307030", name: "BALDANIYA HEMDIP NARESHBHAI", marks: [4.8, 4.8, 4.8, 4.8, 4.8, 0] },
      { roll: "206310307045", name: "RAMI SHREY MANISHKUMAR", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307050", name: "PRAJAPATI PARTHKUMAR SURESHKUMAR", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307066", name: "DARJI PRATHAM SUDHIRKUMAR", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307074", name: "DARJI UJAS RATILAL", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307107", name: "PATEL AARGEE RAJESHKUMAR", marks: [4.8, 4.8, 4.8, 4.8, 4.8, 0] },
      { roll: "206310307115", name: "VANKAR KAJALBEN PRAVINKUMAR", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307125", name: "PATEL YASH KIRTIBHAI", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307129", name: "BAROT RUSHI NIKESHBHAI", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307140", name: "PATEL YASHKUMAR VIJAYBHAI", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307186", name: "BALDANIYA RUTUDHVAJ KANUBHAI", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307193", name: "PANCHAL HASMI VIJAYKUMAR", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307194", name: "Thakor Saloni Nareshbhai", marks: [4.8, 4.8, 4.8, 4.8, 4.8, 0] },
      { roll: "206310307195", name: "CHAUDHARY KHUSHI HITESHBHAI", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307198", name: "PATEL TRUPTI NARESHKUMAR", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "206310307200", name: "GAUSWAMI LAHEREBHARTHI KANUBHARTHI", marks: [4, 4, 4, 4, 4, 0] },
      { roll: "216310307002", name: "PATEL MESHWABEN JITUBHAI", marks: [4.8, 4.8, 4.8, 4.8, 4.8, 0] },
    ],
  },
  endsem: {
    maxMarks: [10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10, 10],
    students: [
      { roll: "R01", name: "Student 1", marks: [6, 7, 5, 7, 8, 7, 9, 4, 6, 5, 5, 6] },
      { roll: "R02", name: "Student 2", marks: [7, 7, 7, 10, 8, 3, 7, 5, 5, 9, 6, 2] },
      { roll: "R03", name: "Student 3", marks: [7, 6, 4, 4, 5, 6, 5, 6, 10, 5, 5, 6] },
      { roll: "R04", name: "Student 4", marks: [8, 5, 9, 4, 4, 6, 4, 5, 7, 8, 6, 6] },
      { roll: "R05", name: "Student 5", marks: [9, 7, 6, 9, 4, 7, 8, 6, 5, 7, 5, 5] },
      { roll: "R06", name: "Student 6", marks: [4, 9, 5, 9, 7, 6, 8, 7, 6, 3, 6, 8] },
      { roll: "R07", name: "Student 7", marks: [7, 8, 9, 7, 10, 3, 6, 1, 8, 6, 10, 5] },
      { roll: "R08", name: "Student 8", marks: [2, 9, 4, 4, 6, 7, 5, 7, 2, 10, 6, 6] },
      { roll: "R09", name: "Student 9", marks: [7, 7, 7, 4, 6, 7, 8, 6, 6, 7, 7, 7] },
      { roll: "R10", name: "Student 10", marks: [6, 5, 8, 7, 6, 10, 8, 8, 7, 4, 8, 6] },
      { roll: "R11", name: "Student 11", marks: [6, 8, 8, 7, 6, 10, 7, 4, 8, 7, 6, 8] },
      { roll: "R12", name: "Student 12", marks: [7, 10, 7, 7, 9, 5, 8, 6, 9, 5, 7, 3] },
      { roll: "R13", name: "Student 13", marks: [5, 6, 7, 10, 10, 4, 5, 9, 7, 3, 7, 4] },
      { roll: "R14", name: "Student 14", marks: [4, 4, 6, 10, 7, 6, 10, 5, 6, 6, 7, 6] },
      { roll: "R15", name: "Student 15", marks: [9, 6, 7, 7, 4, 4, 6, 2, 6, 6, 6, 6] },
    ],
  },
  survey: {
    students: [
      { roll: "R01", name: "Student 1", ratings: [3, 3, 1, 3, 1, 1] },
      { roll: "R02", name: "Student 2", ratings: [1, 3, 2, 3, 2, 1] },
      { roll: "R03", name: "Student 3", ratings: [2, 2, 2, 3, 3, 2] },
      { roll: "R04", name: "Student 4", ratings: [3, 3, 2, 3, 2, 2] },
      { roll: "R05", name: "Student 5", ratings: [1, 2, 3, 3, 2, 3] },
      { roll: "R06", name: "Student 6", ratings: [2, 2, 2, 3, 3, 3] },
      { roll: "R07", name: "Student 7", ratings: [2, 2, 2, 2, 2, 2] },
      { roll: "R08", name: "Student 8", ratings: [2, 2, 3, 3, 2, 3] },
      { roll: "R09", name: "Student 9", ratings: [1, 1, 2, 2, 1, 2] },
      { roll: "R10", name: "Student 10", ratings: [3, 3, 2, 3, 3, 2] },
      { roll: "R11", name: "Student 11", ratings: [2, 3, 2, 3, 3, 3] },
      { roll: "R12", name: "Student 12", ratings: [3, 2, 3, 1, 1, 3] },
      { roll: "R13", name: "Student 13", ratings: [2, 3, 3, 3, 3, 1] },
      { roll: "R14", name: "Student 14", ratings: [3, 2, 3, 3, 1, 2] },
      { roll: "R15", name: "Student 15", ratings: [2, 3, 2, 1, 3, 3] },
    ],
  },
};

function blankState() {
  return {
    courseInfo: {
      institute: "K.D.POLYTECHNIC,PATAN",
      department: "COMPUTER ENGINEERING DEPARTMENT",
      courseName: "",
      courseCode: "",
      semester: "",
      year: "",
      batch: "",
      term: "",
      numStudents: 0,
      faculty: "",
    },
    numCos: 5,
    coCodes: [...DEFAULT_CO_CODES],
    coStatements: COs.map(() => ""),
    weights: { internal: 0, endsem: 0, direct: 0, indirect: 0 },
    targets: { targetPctCO: COs.map(() => 0), coTargetLevel: 0, level3: 0, level2: 0, level1: 0 },
    evalPlan: JSON.parse(JSON.stringify(DEFAULT_EVAL_PLAN)),
    targetLevels: JSON.parse(JSON.stringify(DEFAULT_TARGET_LEVELS)),
    targetSetting: JSON.parse(JSON.stringify(DEFAULT_TARGET_SETTING)),
    mapping: COs.map(() => POPSO.map(() => 0)),
    targetPO: POPSO.map(() => 0),
    internal1: { maxMarks: [2, 14, 5, 2, 2, 0], students: [] },
    internal2: { maxMarks: Array(6).fill(0), students: [] },
    assignment: { maxMarks: Array(6).fill(0), students: [] },
    endsem: { maxMarks: Array(12).fill(0), students: [] },
    survey: { students: [] },
  };
}

function sampleState() {
  return {
    ...JSON.parse(JSON.stringify(SAMPLE)),
    weights: { ...DEFAULT_WEIGHTS },
    targets: { ...DEFAULT_TARGETS },
  };
}

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
  const legacyTargetPct = s.targets?.targetPct;
  const targets = {
    targetPctCO: COs.map((_, i) => clampNum(s.targets?.targetPctCO?.[i] ?? legacyTargetPct, 0, 1, DEFAULT_TARGETS.targetPctCO[i])),
    coTargetLevel: clampNum(s.targets?.coTargetLevel, 0, 3, DEFAULT_TARGETS.coTargetLevel),
    level3: clampNum(s.targets?.level3, 0, 1, DEFAULT_TARGETS.level3),
    level2: clampNum(s.targets?.level2, 0, 1, DEFAULT_TARGETS.level2),
    level1: clampNum(s.targets?.level1, 0, 1, DEFAULT_TARGETS.level1),
  };
  const mapping = (s.mapping && s.mapping.length === 6 ? s.mapping : blankState().mapping)
    .map((row) => POPSO.map((_, j) => clampNum(row[j], 0, 3, 0)));
  const targetPO = (s.targetPO && s.targetPO.length === POPSO.length ? s.targetPO : POPSO.map(() => 2))
    .map((v) => clampNum(v, 0, 3, 2));

  const numCos = clampNum(s.numCos, 1, 6, 5);
  const coCodes = COs.map((_, i) => s.coCodes?.[i] || DEFAULT_CO_CODES[i]);

  const evalPlan = {
    marks: {
      midSem: clampNum(s.evalPlan?.marks?.midSem, 0, 100, DEFAULT_EVAL_PLAN.marks.midSem),
      gtu: clampNum(s.evalPlan?.marks?.gtu, 0, 100, DEFAULT_EVAL_PLAN.marks.gtu),
      pa: clampNum(s.evalPlan?.marks?.pa, 0, 100, DEFAULT_EVAL_PLAN.marks.pa),
      ese: clampNum(s.evalPlan?.marks?.ese, 0, 100, DEFAULT_EVAL_PLAN.marks.ese),
    },
    weights: {
      midSem: clampNum(s.evalPlan?.weights?.midSem, 0, 100, DEFAULT_EVAL_PLAN.weights.midSem),
      gtu: clampNum(s.evalPlan?.weights?.gtu, 0, 100, DEFAULT_EVAL_PLAN.weights.gtu),
      pa: clampNum(s.evalPlan?.weights?.pa, 0, 100, DEFAULT_EVAL_PLAN.weights.pa),
      ese: clampNum(s.evalPlan?.weights?.ese, 0, 100, DEFAULT_EVAL_PLAN.weights.ese),
    },
  };

  const targetLevels = {
    theory: {
      paTarget: clampNum(s.targetLevels?.theory?.paTarget, 0, 3, DEFAULT_TARGET_LEVELS.theory.paTarget),
      paStudents: clampNum(s.targetLevels?.theory?.paStudents, 0, 100, DEFAULT_TARGET_LEVELS.theory.paStudents),
      paMarks: clampNum(s.targetLevels?.theory?.paMarks, 0, 100, DEFAULT_TARGET_LEVELS.theory.paMarks),
      eseTarget: clampNum(s.targetLevels?.theory?.eseTarget, 0, 3, DEFAULT_TARGET_LEVELS.theory.eseTarget),
      eseStudents: clampNum(s.targetLevels?.theory?.eseStudents, 0, 100, DEFAULT_TARGET_LEVELS.theory.eseStudents),
      eseMarks: clampNum(s.targetLevels?.theory?.eseMarks, 0, 100, DEFAULT_TARGET_LEVELS.theory.eseMarks),
    },
    practical: {
      paTarget: clampNum(s.targetLevels?.practical?.paTarget, 0, 3, DEFAULT_TARGET_LEVELS.practical.paTarget),
      paStudents: clampNum(s.targetLevels?.practical?.paStudents, 0, 100, DEFAULT_TARGET_LEVELS.practical.paStudents),
      paMarks: clampNum(s.targetLevels?.practical?.paMarks, 0, 100, DEFAULT_TARGET_LEVELS.practical.paMarks),
      eseTarget: clampNum(s.targetLevels?.practical?.eseTarget, 0, 3, DEFAULT_TARGET_LEVELS.practical.eseTarget),
      eseStudents: clampNum(s.targetLevels?.practical?.eseStudents, 0, 100, DEFAULT_TARGET_LEVELS.practical.eseStudents),
      eseMarks: clampNum(s.targetLevels?.practical?.eseMarks, 0, 100, DEFAULT_TARGET_LEVELS.practical.eseMarks),
    },
    finalTarget: clampNum(s.targetLevels?.finalTarget, 0, 3, DEFAULT_TARGET_LEVELS.finalTarget),
  };

  const ts = s.targetSetting || DEFAULT_TARGET_SETTING;
  const targetSetting = {
    nbaSubjectCode: ts.nbaSubjectCode || "C303_N",
    years: (ts.years && ts.years.length === 3) ? ts.years : DEFAULT_TARGET_SETTING.years,
    gradeCounts: Array.from({ length: 3 }, (_, y) =>
      Array.from({ length: 8 }, (_, k) => clampNum(ts.gradeCounts?.[y]?.[k], 0, 10000, DEFAULT_TARGET_SETTING.gradeCounts[y]?.[k] || 0))
    ),
    targetMarksPct: clampNum(ts.targetMarksPct, 0, 100, DEFAULT_TARGET_SETTING.targetMarksPct),
    targetStudentsPct: clampNum(ts.targetStudentsPct, 0, 100, DEFAULT_TARGET_SETTING.targetStudentsPct),
    targetLevel: clampNum(ts.targetLevel, 0, 3, DEFAULT_TARGET_SETTING.targetLevel),
    rangeMatrix: (ts.rangeMatrix && ts.rangeMatrix.length === 4) ? ts.rangeMatrix : DEFAULT_TARGET_SETTING.rangeMatrix,
  };

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
      ratings: Array.from({ length: 6 }, (_, i) => {
        const v = st.ratings?.[i];
        if (v === null || v === undefined || v === "") return null;
        return clampNum(v, 1, 3, 2);
      }),
    })),
  };

  return {
    courseInfo: {
      institute: s.courseInfo?.institute || "K.D.POLYTECHNIC,PATAN",
      department: s.courseInfo?.department || "COMPUTER ENGINEERING DEPARTMENT",
      courseName: s.courseInfo?.courseName || "Introduction to Machine Learning",
      courseCode: s.courseInfo?.courseCode || "4350702",
      semester: s.courseInfo?.semester || "5",
      batch: s.courseInfo?.batch || "2021-24",
      term: s.courseInfo?.term || "2023-24 ODD (231)",
      numStudents: clampNum(s.courseInfo?.numStudents, 0, 1000, 137),
      faculty: s.courseInfo?.faculty || "",
      year: s.courseInfo?.year || "2023-24",
    },
    numCos,
    coCodes,
    coStatements: COs.map((_, i) => s.coStatements?.[i] ?? ""),
    weights,
    targets,
    evalPlan,
    targetLevels,
    targetSetting,
    mapping,
    targetPO,
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
function calculateAttainmentLevel(pct, targetSetting) {
  const fallback = targetSetting?.targetLevel ?? 0.90;
  if (pct >= 70) return 3.00;
  if (pct >= 60) return 2.00;
  if (pct >= 50) return 1.00;
  return fallback;
}

function simpleCOStats(assessment, targets, targetSetting) {
  const n = (assessment?.students || []).length;
  return (assessment?.maxMarks || Array(6).fill(0)).map((max, i) => {
    const m = Number(max) || 0;
    const targetMarks = m * (targets?.targetPctCO?.[i] ?? 0.61);
    const attained = (assessment?.students || []).filter((s) => (Number(s.marks?.[i]) || 0) >= targetMarks).length;
    const pct = n ? (attained / n) * 100 : 0;
    return { max: m, targetMarks, attained, total: n, pct, level: calculateAttainmentLevel(pct, targetSetting) };
  });
}
function endsemCOStats(assessment, targets, targetSetting) {
  const n = (assessment?.students || []).length;
  const res = [];
  for (let i = 0; i < 6; i++) {
    const max = (Number(assessment?.maxMarks?.[i]) || 0) + (Number(assessment?.maxMarks?.[i + 6]) || 0);
    const targetMarks = max * (targets?.targetPctCO?.[i] ?? 0.61);
    const attained = (assessment?.students || []).filter((s) => {
      const tot = (Number(s.marks?.[i]) || 0) + (Number(s.marks?.[i + 6]) || 0);
      return tot >= targetMarks;
    }).length;
    const pct = n ? (attained / n) * 100 : 0;
    res.push({ max, targetMarks, attained, total: n, pct, level: calculateAttainmentLevel(pct, targetSetting) });
  }
  return res;
}
function surveyCOAverage(survey) {
  const res = [];
  for (let i = 0; i < 6; i++) {
    const rated = survey.students.filter((s) => s.ratings[i] !== null && s.ratings[i] !== undefined && s.ratings[i] !== "");
    const sum = rated.reduce((a, s) => a + (Number(s.ratings[i]) || 0), 0);
    res.push(rated.length ? sum / rated.length : 0);
  }
  return res;
}

/* ============================================================================
   EXCEL IMPORT / EXPORT WITH FIXED OFFICIAL LAYOUTS & RICH CELL STYLING
============================================================================ */
// Color constants matching institutional images
const C_PEACH = "FCE4D6";
const C_GREEN = "E2EFDA";
const C_CYAN = "BDD7EE";
const C_AMBER = "FFC000";
const C_GREY = "BFBFBF";
const C_DARK_GREY = "A6A6A6";
const C_LAVENDER = "E7E6F4";
const C_BLUE_BADGE = "2F5597";
const C_YELLOW = "FFFF00";
const C_WHITE = "FFFFFF";

const BORDER_THIN = {
  top: { style: "thin", color: { rgb: "000000" } },
  bottom: { style: "thin", color: { rgb: "000000" } },
  left: { style: "thin", color: { rgb: "000000" } },
  right: { style: "thin", color: { rgb: "000000" } },
};

function styleCell({ bg, color = "000000", bold = false, sz = 11, align = "center", border = BORDER_THIN, wrap = true, italic = false }) {
  const s = {
    font: { name: "Calibri", sz, bold, italic, color: { rgb: color } },
    alignment: { horizontal: align, vertical: "center", wrapText: wrap },
  };
  if (bg) {
    s.fill = { fgColor: { rgb: bg }, patternType: "solid" };
  }
  if (border) {
    s.border = border;
  }
  return s;
}

function applyRangeStyle(ws, r1, c1, r2, c2, style) {
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      if (!ws[addr]) {
        ws[addr] = { v: "", t: "s" };
      }
      ws[addr].s = style;
    }
  }
}

function setCellStyle(ws, r, c, style) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (!ws[addr]) {
    ws[addr] = { v: "", t: "s" };
  }
  ws[addr].s = style;
}

function setCellWithFormula(ws, r, c, formula, val, style) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const type = typeof val === "number" ? "n" : typeof val === "boolean" ? "b" : "s";
  ws[addr] = {
    t: type,
    f: formula,
    v: val,
    s: style || {}
  };
}

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
    else if (label === "Batch") ci.batch = value ?? "";
    else if (label === "Term") ci.term = value ?? "";
    else if (label === "Number of Students") ci.numStudents = Number(value) || 0;
    else if (label === "Academic Year") ci.year = value ?? "";
    else if (label === "Faculty Name") ci.faculty = value ?? "";
    else if (/^CO[1-6]$/.test(label)) co[Number(label.slice(2)) - 1] = value ?? "";
    else if (label.includes("Indirect Attainment Weight")) w.indirect = Number(value) || 0;
    else if (label.includes("Direct Attainment Weight")) w.direct = Number(value) || 0;
    else if (label.includes("Internal Weight")) w.internal = Number(value) || 0;
    else if (label.includes("End-Sem Weight")) w.endsem = Number(value) || 0;
    else if (/^Target % of marks .* CO[1-6]$/.test(label)) {
      const idx = Number(label.slice(-1)) - 1;
      t.targetPctCO = [...(t.targetPctCO || COs.map(() => 0))];
      t.targetPctCO[idx] = Number(value) || 0;
    }
    else if (label.includes("Target % of marks")) t.targetPctCO = COs.map(() => Number(value) || 0);
    else if (label.includes("CO Target Attainment Level")) t.coTargetLevel = Number(value) || 0;
    else if (label.includes("Level 3")) t.level3 = Number(value) || 0;
    else if (label.includes("Level 2")) t.level2 = Number(value) || 0;
    else if (label.includes("Level 1")) t.level1 = Number(value) || 0;
  }
  return { courseInfo: ci, coStatements: co, weights: w, targets: t };
}

function parseMapping(rows, state) {
  let headerIdx = rows.findIndex((r) => r[0] === "CO / PO" || r[1] === "CO / PO" || r[0] === "CO-PO-PSO Mapping");
  if (headerIdx === -1) return {};
  const mapping = [];
  let r = headerIdx + 1;
  while (r < rows.length && /^CO[1-6]$/.test(String(rows[r][0] || rows[r][1] || ""))) {
    const colOffset = rows[r][0] === "" ? 2 : 1;
    mapping.push(POPSO.map((_, j) => Number(rows[r][colOffset + j]) || 0));
    r++;
  }
  const targetRow = rows.find((row) => row[0] === "Target Attainment" || row[1] === "Target Attainment" || row[0] === "AVERAGE");
  const colOffset = targetRow && targetRow[0] === "" ? 2 : 1;
  const targetPO = targetRow ? POPSO.map((_, j) => Number(targetRow[colOffset + j]) || 2) : state.targetPO;
  return { mapping: mapping.length === 6 ? mapping : state.mapping, targetPO };
}

function parseAssessment(rows, nQ) {
  const maxRow = rows.find(
    (r) =>
      r[0] === "Max Marks \u2192" ||
      r[1] === "Max Marks \u2192" ||
      String(r[0] || "").includes("MAX MARKS") ||
      String(r[1] || "").includes("MAX MARKS")
  );
  let offset = 2;
  if (maxRow) {
    if (String(maxRow[0] || "").includes("MAX MARKS")) offset = 3;
    else if (maxRow[0] === "") offset = 3;
  }
  const maxMarks = maxRow ? Array.from({ length: nQ }, (_, i) => Number(maxRow[offset + i]) || 0) : Array(nQ).fill(10);
  const headerIdx = rows.findIndex(
    (r) =>
      r[0] === "Roll No" ||
      r[1] === "Roll No" ||
      r[0] === "SR. NO." ||
      r[1] === "ENROLL NO." ||
      r[0] === "ENROLL NO."
  );
  const students = [];
  if (headerIdx !== -1) {
    const isPRPA = rows[headerIdx][0] === "SR. NO." || rows[headerIdx][1] === "ENROLL NO.";
    const rollCol = isPRPA ? 1 : rows[headerIdx][0] === "Roll No" ? 0 : 1;
    const nameCol = isPRPA ? 2 : rollCol + 1;
    const marksCol = isPRPA ? 3 : rollCol + 2;

    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row[rollCol] == null || row[rollCol] === "") break;
      students.push({
        roll: String(row[rollCol]),
        name: row[nameCol] == null ? "" : String(row[nameCol]),
        marks: Array.from({ length: nQ }, (_, i) => Number(row[marksCol + i]) || 0),
      });
    }
  }
  return { maxMarks, students };
}

function parseSurvey(rows) {
  const headerIdx = rows.findIndex((r) => r[0] === "Roll No" || r[1] === "Roll No");
  const students = [];
  if (headerIdx !== -1) {
    const rollCol = rows[headerIdx][0] === "Roll No" ? 0 : 1;
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row[rollCol] == null || row[rollCol] === "") break;
      students.push({
        roll: String(row[rollCol]),
        name: row[rollCol + 1] == null ? "" : String(row[rollCol + 1]),
        ratings: Array.from({ length: 6 }, (_, i) => Number(row[rollCol + 2 + i]) || 0),
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
  const i1 = sheetRows(wb, "PR_PA") || sheetRows(wb, "Internal_Exam1");
  if (i1) { next.internal1 = parseAssessment(i1, 6); found.push("PR_PA"); } else missing.push("PR_PA");
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

/* 1. Builder for Course Evaluation Plan Sheet (Matches Prompt 1 Images 1, 2, 3) */
function buildCourseEvaluationPlanSheet(state) {
  const numCos = state.numCos || 5;
  const activeCOs = COs.slice(0, numCos);
  const mMarks = state.evalPlan.marks;
  const mWts = state.evalPlan.weights;
  const totalMks = (mMarks.midSem || 0) + (mMarks.gtu || 0) + (mMarks.pa || 0) + (mMarks.ese || 0);
  const totalWts = (mWts.midSem || 0) + (mWts.gtu || 0) + (mWts.pa || 0) + (mWts.ese || 0);
  const thTotal = (mMarks.midSem || 0) + (mMarks.gtu || 0);
  const prTotal = (mMarks.pa || 0) + (mMarks.ese || 0);
  const thPct = totalMks ? Math.round((thTotal / totalMks) * 100) : 67;
  const prPct = totalMks ? Math.round((prTotal / totalMks) * 100) : 33;

  const poAverages = DIPLOMA_POPSO.map((_, j) => {
    let sum = 0, count = 0;
    for (let i = 0; i < numCos; i++) {
      const v = state.mapping[i]?.[j];
      if (v > 0) { sum += v; count++; }
    }
    return count > 0 ? (sum / count).toFixed(2) : "-";
  });

  const tLev = state.targetLevels;
  const thAvg = ((tLev.theory.paTarget + tLev.theory.eseTarget) / 2).toFixed(2);
  const prAvg = ((tLev.practical.paTarget + tLev.practical.eseTarget) / 2).toFixed(2);
  const finalCourseTarget = tLev.finalTarget || 0.9;

  const rows = [];
  const merges = [];

  // Row 0: Institute
  rows.push([state.courseInfo.institute || "K.D.POLYTECHNIC,PATAN"]);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 8 } });

  // Row 1: Dept
  rows.push([state.courseInfo.department || "COMPUTER ENGINEERING DEPARTMENT"]);
  merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 8 } });

  // Row 2: Doc Title
  rows.push(["COURSE EVALUATION PLAN"]);
  merges.push({ s: { r: 2, c: 0 }, e: { r: 2, c: 8 } });

  rows.push([]); // Blank

  // Row 4: Course Code & Name
  rows.push(["Course Code:", state.courseInfo.courseCode || "4350702", "Course Name:", state.courseInfo.courseName || "Introduction to Machine Learning"]);
  merges.push({ s: { r: 4, c: 3 }, e: { r: 4, c: 8 } });

  // Row 5: Batch & Term
  rows.push(["Batch:", state.courseInfo.batch || "2021-24", "Term:", state.courseInfo.term || "231"]);
  merges.push({ s: { r: 5, c: 3 }, e: { r: 5, c: 8 } });

  // Row 6: Students & Semester
  rows.push(["Number of Students:", state.courseInfo.numStudents || 137, "Semester :", state.courseInfo.semester || "5"]);
  merges.push({ s: { r: 6, c: 3 }, e: { r: 6, c: 8 } });

  rows.push([]); // Blank

  // Section 1: CO's
  rows.push(["CO's"]);
  merges.push({ s: { r: 8, c: 0 }, e: { r: 8, c: 1 } });

  activeCOs.forEach((co, i) => {
    const rIdx = rows.length;
    rows.push([co, state.coCodes?.[i] || DEFAULT_CO_CODES[i], state.coStatements[i]]);
    merges.push({ s: { r: rIdx, c: 2 }, e: { r: rIdx, c: 8 } });
  });

  rows.push(["No of Cos", numCos]);
  merges.push({ s: { r: 9 + numCos, c: 2 }, e: { r: 9 + numCos, c: 8 } });

  rows.push([]); // Blank

  // Section 2: CO-PO-PSO Mapping
  const rMapHdr = rows.length;
  rows.push(["CO-PO-PSO Mapping"]);
  merges.push({ s: { r: rMapHdr, c: 0 }, e: { r: rMapHdr, c: 2 } });

  rows.push(["", ...DIPLOMA_POPSO]);

  activeCOs.forEach((co, i) => {
    rows.push([co, ...DIPLOMA_POPSO.map((_, j) => state.mapping[i]?.[j] ? state.mapping[i][j] : "-")]);
  });

  rows.push(["AVERAGE", ...poAverages.map((v) => v === "-" ? "" : v)]);

  rows.push([]); // Blank

  // Section 3: Evaluation Plan Table
  const rEvalHdr = rows.length;
  rows.push(["Evaluation Plan"]);
  merges.push({ s: { r: rEvalHdr, c: 0 }, e: { r: rEvalHdr, c: 1 } });

  const rEvalTblHdr1 = rows.length;
  rows.push(["Assesment Tool", "Theory Marks", "", "Practical Marks", "", "Total Marks"]);
  merges.push({ s: { r: rEvalTblHdr1, c: 0 }, e: { r: rEvalTblHdr1 + 1, c: 0 } });
  merges.push({ s: { r: rEvalTblHdr1, c: 1 }, e: { r: rEvalTblHdr1, c: 2 } });
  merges.push({ s: { r: rEvalTblHdr1, c: 3 }, e: { r: rEvalTblHdr1, c: 4 } });
  merges.push({ s: { r: rEvalTblHdr1, c: 5 }, e: { r: rEvalTblHdr1 + 1, c: 5 } });

  rows.push(["", "Mid Sem", "GTU", "PA", "ESE", ""]);

  rows.push(["Marks", mMarks.midSem, mMarks.gtu, mMarks.pa, mMarks.ese, totalMks]);
  rows.push(["Weightage(%)", mWts.midSem, mWts.gtu, mWts.pa, mWts.ese, totalWts, "Refer GTU syllabus"]);

  const rCOBreak1 = rows.length;
  rows.push([`CO1 to CO${numCos}`, thTotal, "", prTotal, "", `TH: ${thPct}%`, `PR: ${prPct}%`]);
  merges.push({ s: { r: rCOBreak1, c: 0 }, e: { r: rCOBreak1 + 1, c: 0 } });
  merges.push({ s: { r: rCOBreak1, c: 1 }, e: { r: rCOBreak1 + 1, c: 2 } });
  merges.push({ s: { r: rCOBreak1, c: 3 }, e: { r: rCOBreak1 + 1, c: 4 } });

  const rCOBreak2 = rows.length;
  rows.push(["", "", "", "", "", "100%"]);
  merges.push({ s: { r: rCOBreak2, c: 5 }, e: { r: rCOBreak2, c: 6 } });

  rows.push([]); // Blank

  // Section 4: Defined Target Levels for All CO's
  const rTgtHdr = rows.length;
  rows.push(["Defined Target Leves for All CO's"]);
  merges.push({ s: { r: rTgtHdr, c: 0 }, e: { r: rTgtHdr, c: 3 } });

  const rTgtTblHdr1 = rows.length;
  rows.push(["TH/PR", "Course Outcome", "PROGRESSIVE ASSESSMENT", "", "", "ESE EXAMINATION", "", "", "Average of Target"]);
  merges.push({ s: { r: rTgtTblHdr1, c: 0 }, e: { r: rTgtTblHdr1 + 2, c: 0 } });
  merges.push({ s: { r: rTgtTblHdr1, c: 1 }, e: { r: rTgtTblHdr1 + 2, c: 1 } });
  merges.push({ s: { r: rTgtTblHdr1, c: 2 }, e: { r: rTgtTblHdr1, c: 4 } });
  merges.push({ s: { r: rTgtTblHdr1, c: 5 }, e: { r: rTgtTblHdr1, c: 7 } });
  merges.push({ s: { r: rTgtTblHdr1, c: 8 }, e: { r: rTgtTblHdr1 + 2, c: 8 } });

  const rTgtTblHdr2 = rows.length;
  rows.push(["", "", "Target Level", "(a)% Students Scoring More than or Equal to (b)% Marks", "", "Target Level", "(c)% Students Scoring More than or Equal to (d)% Marks", ""]);
  merges.push({ s: { r: rTgtTblHdr2, c: 2 }, e: { r: rTgtTblHdr2 + 1, c: 2 } });
  merges.push({ s: { r: rTgtTblHdr2, c: 3 }, e: { r: rTgtTblHdr2, c: 4 } });
  merges.push({ s: { r: rTgtTblHdr2, c: 5 }, e: { r: rTgtTblHdr2 + 1, c: 5 } });
  merges.push({ s: { r: rTgtTblHdr2, c: 6 }, e: { r: rTgtTblHdr2, c: 7 } });

  rows.push(["", "", "", "(a) students", "(b) marks", "", "(c) students", "(d) marks"]);

  rows.push(["THEORY", `CO1 to CO${numCos}`, Number(tLev.theory.paTarget.toFixed(2)), tLev.theory.paStudents, tLev.theory.paMarks, Number(tLev.theory.eseTarget.toFixed(2)), tLev.theory.eseStudents, tLev.theory.eseMarks, Number(thAvg)]);
  rows.push(["PRACTICAL", `CO1 to CO${numCos}`, Number(tLev.practical.paTarget.toFixed(2)), tLev.practical.paStudents, tLev.practical.paMarks, Number(tLev.practical.eseTarget.toFixed(2)), tLev.practical.eseStudents, tLev.practical.eseMarks, Number(prAvg)]);

  const rFinalTgt = rows.length;
  rows.push([`Final Target for Whole Course : ${finalCourseTarget}`]);
  merges.push({ s: { r: rFinalTgt, c: 0 }, e: { r: rFinalTgt, c: 8 } });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"] = merges;
  ws["!cols"] = [
    { wch: 18 },
    { wch: 15 },
    { wch: 15 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 18 },
  ];

  // APPLY RICH STYLING TO CELLS (Exact Colors from Images)
  // Top Headers
  applyRangeStyle(ws, 0, 0, 0, 8, styleCell({ sz: 16, bold: true, align: "center", border: undefined }));
  applyRangeStyle(ws, 1, 0, 1, 8, styleCell({ sz: 13, bold: true, align: "center", border: undefined }));
  applyRangeStyle(ws, 2, 0, 2, 8, styleCell({ sz: 12, bold: true, align: "center", border: undefined }));

  // Course Details (Rows 4..6)
  for (let r = 4; r <= 6; r++) {
    setCellStyle(ws, r, 0, styleCell({ bg: C_WHITE, bold: true, align: "right" }));
    setCellStyle(ws, r, 1, styleCell({ bg: C_GREEN, bold: true, align: "center" }));
    setCellStyle(ws, r, 2, styleCell({ bg: C_WHITE, bold: true, align: "right" }));
    applyRangeStyle(ws, r, 3, r, 8, styleCell({ bg: C_GREEN, bold: true, align: "left" }));
  }

  // Section 1: CO's Badge & Table
  applyRangeStyle(ws, 8, 0, 8, 1, styleCell({ bg: C_BLUE_BADGE, color: "FFFFFF", bold: true, sz: 12 }));
  for (let i = 0; i < numCos; i++) {
    const r = 9 + i;
    setCellStyle(ws, r, 0, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
    setCellStyle(ws, r, 1, styleCell({ bg: C_GREEN, bold: true, align: "center" }));
    applyRangeStyle(ws, r, 2, r, 8, styleCell({ bg: C_GREEN, align: "left" }));
  }
  const rNoCos = 9 + numCos;
  setCellStyle(ws, rNoCos, 0, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
  setCellStyle(ws, rNoCos, 1, styleCell({ bg: C_GREEN, bold: true, align: "center" }));
  applyRangeStyle(ws, rNoCos, 2, rNoCos, 8, styleCell({ bg: C_GREEN }));

  // Section 2: Mapping Badge & Table
  const rMapStart = rNoCos + 2;
  applyRangeStyle(ws, rMapStart, 0, rMapStart, 2, styleCell({ bg: C_BLUE_BADGE, color: "FFFFFF", bold: true, sz: 12 }));
  for (let c = 0; c <= 9; c++) {
    setCellStyle(ws, rMapStart + 1, c, styleCell({ bg: C_PEACH, bold: true }));
  }
  for (let i = 0; i < numCos; i++) {
    const r = rMapStart + 2 + i;
    setCellStyle(ws, r, 0, styleCell({ bg: C_PEACH, bold: true }));
    for (let c = 1; c <= 9; c++) {
      setCellStyle(ws, r, c, styleCell({ bg: C_GREEN, align: "center" }));
    }
  }
  const rAvg = rMapStart + 2 + numCos;
  setCellStyle(ws, rAvg, 0, styleCell({ bg: C_PEACH, bold: true }));
  for (let c = 1; c <= 9; c++) {
    setCellStyle(ws, rAvg, c, styleCell({ bg: C_GREEN, bold: true, align: "center" }));
  }

  // Section 3: Evaluation Plan Badge & Table
  const rEvalStart = rAvg + 2;
  applyRangeStyle(ws, rEvalStart, 0, rEvalStart, 1, styleCell({ bg: C_BLUE_BADGE, color: "FFFFFF", bold: true, sz: 12 }));
  applyRangeStyle(ws, rEvalStart + 1, 0, rEvalStart + 2, 0, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rEvalStart + 1, 1, rEvalStart + 1, 2, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rEvalStart + 1, 3, rEvalStart + 1, 4, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rEvalStart + 1, 5, rEvalStart + 2, 5, styleCell({ bg: C_PEACH, bold: true }));
  for (let c = 1; c <= 4; c++) {
    setCellStyle(ws, rEvalStart + 2, c, styleCell({ bg: C_PEACH, bold: true }));
  }
  // Marks Row
  setCellStyle(ws, rEvalStart + 3, 0, styleCell({ bg: C_CYAN, bold: true }));
  for (let c = 1; c <= 5; c++) {
    setCellStyle(ws, rEvalStart + 3, c, styleCell({ bg: C_GREEN, bold: true }));
  }
  // Weightage Row
  setCellStyle(ws, rEvalStart + 4, 0, styleCell({ bg: C_CYAN, bold: true }));
  for (let c = 1; c <= 5; c++) {
    setCellStyle(ws, rEvalStart + 4, c, styleCell({ bg: C_GREEN, bold: true }));
  }
  setCellStyle(ws, rEvalStart + 4, 6, styleCell({ color: "FF0000", bold: true, align: "left", border: undefined }));
  // CO breakdown block
  applyRangeStyle(ws, rEvalStart + 5, 0, rEvalStart + 6, 0, styleCell({ bg: C_LAVENDER, bold: true }));
  applyRangeStyle(ws, rEvalStart + 5, 1, rEvalStart + 6, 2, styleCell({ bg: C_DARK_GREY, bold: true, sz: 13 }));
  applyRangeStyle(ws, rEvalStart + 5, 3, rEvalStart + 6, 4, styleCell({ bg: C_DARK_GREY, bold: true, sz: 13 }));
  setCellStyle(ws, rEvalStart + 5, 5, styleCell({ bg: C_GREY, bold: true }));
  setCellStyle(ws, rEvalStart + 5, 6, styleCell({ bg: C_GREY, bold: true }));
  applyRangeStyle(ws, rEvalStart + 6, 5, rEvalStart + 6, 6, styleCell({ bg: C_GREY, bold: true }));

  // Section 4: Defined Target Levels Badge & Table
  const rTgtStart = rEvalStart + 8;
  applyRangeStyle(ws, rTgtStart, 0, rTgtStart, 3, styleCell({ bg: C_BLUE_BADGE, color: "FFFFFF", bold: true, sz: 12 }));
  applyRangeStyle(ws, rTgtStart + 1, 0, rTgtStart + 3, 0, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rTgtStart + 1, 1, rTgtStart + 3, 1, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rTgtStart + 1, 2, rTgtStart + 1, 4, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rTgtStart + 1, 5, rTgtStart + 1, 7, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rTgtStart + 1, 8, rTgtStart + 3, 8, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rTgtStart + 2, 2, rTgtStart + 3, 2, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rTgtStart + 2, 3, rTgtStart + 2, 4, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rTgtStart + 2, 5, rTgtStart + 3, 5, styleCell({ bg: C_PEACH, bold: true }));
  applyRangeStyle(ws, rTgtStart + 2, 6, rTgtStart + 2, 7, styleCell({ bg: C_PEACH, bold: true }));
  setCellStyle(ws, rTgtStart + 3, 3, styleCell({ bg: C_PEACH, bold: true }));
  setCellStyle(ws, rTgtStart + 3, 4, styleCell({ bg: C_PEACH, bold: true }));
  setCellStyle(ws, rTgtStart + 3, 6, styleCell({ bg: C_PEACH, bold: true }));
  setCellStyle(ws, rTgtStart + 3, 7, styleCell({ bg: C_PEACH, bold: true }));

  // Theory & Practical data rows
  for (let k = 0; k < 2; k++) {
    const r = rTgtStart + 4 + k;
    setCellStyle(ws, r, 0, styleCell({ bg: C_WHITE, bold: true }));
    setCellStyle(ws, r, 1, styleCell({ bg: C_AMBER, bold: true }));
    for (let c = 2; c <= 7; c++) {
      setCellStyle(ws, r, c, styleCell({ bg: C_GREEN, bold: c === 2 || c === 5 }));
    }
    setCellStyle(ws, r, 8, styleCell({ bg: C_CYAN, bold: true }));
  }
  // Footer banner
  applyRangeStyle(ws, rTgtStart + 6, 0, rTgtStart + 6, 8, styleCell({ bg: C_GREY, bold: true, sz: 12 }));

  return ws;
}

/* 2. Builder for Target Setting Sheet (Matches Prompt 2 Images 1, 2, 3) */
function buildTargetSettingSheet(state) {
  const ts = state.targetSetting || DEFAULT_TARGET_SETTING;
  const years = ts.years || DEFAULT_TARGET_SETTING.years;
  const gradeCounts = ts.gradeCounts || DEFAULT_TARGET_SETTING.gradeCounts;

  const yearStats = years.map((_, y) => {
    let totalStudents = 0;
    let totalMarks = 0;
    GTU_GRADES.forEach((g, k) => {
      const count = Number(gradeCounts[y]?.[k]) || 0;
      totalStudents += count;
      totalMarks += count * g.avg;
    });
    const pctResult = totalStudents > 0 ? (totalMarks / (totalStudents * 100)) * 100 : 0;
    return { totalStudents, totalMarks, pctResult };
  });

  const validYears = yearStats.filter((ys) => ys.totalStudents > 0);
  const avgResultPct = validYears.length
    ? validYears.reduce((a, b) => a + b.pctResult, 0) / validYears.length
    : 60.26;
  const targetMarksPct = ts.targetMarksPct || Math.ceil(avgResultPct);

  const studentsAboveThreshold = years.map((_, y) => {
    let count = 0;
    GTU_GRADES.forEach((g, k) => {
      if (g.isAboveThreshold) count += Number(gradeCounts[y]?.[k]) || 0;
    });
    const tot = yearStats[y].totalStudents;
    const pct = tot > 0 ? (count / tot) * 100 : 0;
    return { count, pct };
  });

  const validThresholdYears = studentsAboveThreshold.filter((_, y) => yearStats[y].totalStudents > 0);
  const avgAbovePct = validThresholdYears.length
    ? validThresholdYears.reduce((a, b) => a + b.pct, 0) / validThresholdYears.length
    : 47.0;
  const targetStudentsPct = ts.targetStudentsPct || 48;
  const targetLevel = ts.targetLevel || 0.90;

  const rows = [];
  const merges = [];

  // Row 0: Term & Subject Name
  rows.push(["Term :", state.courseInfo.term || "2023-24 ODD", "", "Subject Name:", state.courseInfo.courseName || "Introduction to Machine Learning"]);
  merges.push({ s: { r: 0, c: 4 }, e: { r: 0, c: 10 } });

  // Row 1: Sem, GTU Code, NBA Code
  rows.push(["Sem :", state.courseInfo.semester || "5", "", "GTU Subject Code:", state.courseInfo.courseCode || "4350702", "", "NBA Subject Code:", ts.nbaSubjectCode || "C303_N"]);

  rows.push([]); // Blank

  // Main Grade Distribution Table Headers
  const rTblHdr1 = rows.length;
  rows.push(["Sr No", "Grade as per GTU", "Avg Marks as per range Grade", "No.of Student attain Grade", "", "", "Total Marks", "", "", "% of Result", "", "", "Average of %"]);
  merges.push({ s: { r: rTblHdr1, c: 0 }, e: { r: rTblHdr1 + 1, c: 0 } });
  merges.push({ s: { r: rTblHdr1, c: 1 }, e: { r: rTblHdr1 + 1, c: 1 } });
  merges.push({ s: { r: rTblHdr1, c: 2 }, e: { r: rTblHdr1 + 1, c: 2 } });
  merges.push({ s: { r: rTblHdr1, c: 3 }, e: { r: rTblHdr1, c: 5 } });
  merges.push({ s: { r: rTblHdr1, c: 6 }, e: { r: rTblHdr1, c: 8 } });
  merges.push({ s: { r: rTblHdr1, c: 9 }, e: { r: rTblHdr1, c: 11 } });
  merges.push({ s: { r: rTblHdr1, c: 12 }, e: { r: rTblHdr1 + 1, c: 12 } });

  rows.push(["", "", "", years[0], years[1], years[2], years[0], years[1], years[2], years[0], years[1], years[2], ""]);

  // Grades rows
  const rGradesStart = rows.length;
  GTU_GRADES.forEach((g, k) => {
    const counts = [
      Number(gradeCounts[0]?.[k]) || 0,
      Number(gradeCounts[1]?.[k]) || 0,
      Number(gradeCounts[2]?.[k]) || 0,
    ];
    const totalM = counts.map((c) => c * g.avg);
    if (k === 0) {
      rows.push([k + 1, g.grade, g.avg, counts[0], counts[1], counts[2], totalM[0], totalM[1], totalM[2], Number(yearStats[0].pctResult.toFixed(2)), "", "", Number(avgResultPct.toFixed(2))]);
    } else {
      rows.push([k + 1, g.grade, g.avg, counts[0], counts[1], counts[2], totalM[0], totalM[1], totalM[2]]);
    }
  });
  const rGradesEnd = rows.length - 1;
  merges.push({ s: { r: rGradesStart, c: 9 }, e: { r: rGradesEnd, c: 11 } });
  merges.push({ s: { r: rGradesStart, c: 12 }, e: { r: rGradesEnd, c: 12 } });

  // Total row
  const rTotal = rows.length;
  rows.push([
    "Total",
    "",
    "",
    yearStats[0].totalStudents,
    yearStats[1].totalStudents,
    yearStats[2].totalStudents,
    yearStats[0].totalMarks,
    yearStats[1].totalMarks,
    yearStats[2].totalMarks,
    "",
    "",
    "",
    targetMarksPct,
  ]);
  merges.push({ s: { r: rTotal, c: 0 }, e: { r: rTotal, c: 2 } });

  rows.push([]); // Blank

  // Section 2: Threshold students & Target Summary Box
  const rThresh1 = rows.length;
  rows.push([
    "No. of students having marks >=",
    targetMarksPct,
    studentsAboveThreshold[0].count,
    studentsAboveThreshold[1].count,
    studentsAboveThreshold[2].count,
    "",
    `Course Outcome Target for term (${state.courseInfo.year || "2023-24"}) 231`,
    "CO Target",
    "% Students",
    "% marks",
  ]);
  merges.push({ s: { r: rThresh1, c: 6 }, e: { r: rThresh1 + 1, c: 6 } });

  rows.push([
    "% of students having marks >=",
    targetMarksPct,
    Number(studentsAboveThreshold[0].pct.toFixed(2)),
    "",
    "",
    "",
    "",
    Number(targetLevel.toFixed(2)),
    targetStudentsPct,
    targetMarksPct,
  ]);

  const rThresh3 = rows.length;
  rows.push([
    "Avg of 3 years of who score >=",
    targetMarksPct,
    Number(avgAbovePct.toFixed(2)),
  ]);
  merges.push({ s: { r: rThresh3, c: 2 }, e: { r: rThresh3, c: 4 } });

  rows.push([
    "Conclusion - From last three years' results,",
    targetStudentsPct,
    "% students who scored >=",
    targetMarksPct,
    "% of Marks",
  ]);

  rows.push([]); // Blank

  // Section 3: Range Matrix & Achieved Comparison Box
  const rMatHdr = rows.length;
  rows.push(["Attainment Level Range Matrix", "", "", "", ""]);
  merges.push({ s: { r: rMatHdr, c: 0 }, e: { r: rMatHdr, c: 4 } });

  const rMatSub1 = rows.length;
  rows.push([
    "Target statements",
    "Attainment Levels",
    "% of average marks",
    "Target: % of Students",
    "",
    "",
    "",
    "CO Target",
    "% Student",
    "% marks",
  ]);
  merges.push({ s: { r: rMatSub1, c: 0 }, e: { r: rMatSub1 + 1, c: 0 } });
  merges.push({ s: { r: rMatSub1, c: 1 }, e: { r: rMatSub1 + 1, c: 1 } });
  merges.push({ s: { r: rMatSub1, c: 2 }, e: { r: rMatSub1 + 1, c: 2 } });
  merges.push({ s: { r: rMatSub1, c: 3 }, e: { r: rMatSub1, c: 4 } });

  rows.push([
    "",
    "",
    "",
    "Min",
    "Max",
    "",
    `Course Outcome Target for term ${state.courseInfo.year || "2023-24"} (231)`,
    Number(targetLevel.toFixed(2)),
    targetStudentsPct,
    targetMarksPct,
  ]);

  const rMatRow1 = rows.length;
  rows.push([
    "Target statements",
    "0.0 - 0.9",
    targetMarksPct,
    0,
    49.99,
    "",
    `Course Outcome Achieved for term ${state.courseInfo.year || "2023-24"} (231)`,
    1.52,
    55,
    targetMarksPct,
  ]);

  rows.push(["", "1.0 - 1.9", targetMarksPct, 50, 59.99]);
  rows.push(["", "2.0 - 2.9", targetMarksPct, 60, 69.99]);
  rows.push(["", "3", targetMarksPct, 70, 100]);
  const rMatRowEnd = rows.length - 1;
  merges.push({ s: { r: rMatRow1, c: 0 }, e: { r: rMatRowEnd, c: 0 } });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"] = merges;
  ws["!cols"] = [
    { wch: 28 },
    { wch: 18 },
    { wch: 14 },
    { wch: 10 },
    { wch: 10 },
    { wch: 10 },
    { wch: 34 },
    { wch: 12 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
  ];

  // APPLY RICH STYLING TO TARGET SETTING CELLS
  // Headers with Yellow
  setCellStyle(ws, 0, 0, styleCell({ bold: true, align: "right" }));
  setCellStyle(ws, 0, 1, styleCell({ bg: C_YELLOW, bold: true, align: "center" }));
  setCellStyle(ws, 0, 3, styleCell({ bold: true, align: "right" }));
  applyRangeStyle(ws, 0, 4, 0, 10, styleCell({ bg: C_YELLOW, bold: true, align: "left" }));

  setCellStyle(ws, 1, 0, styleCell({ bold: true, align: "right" }));
  setCellStyle(ws, 1, 1, styleCell({ bg: C_YELLOW, bold: true, align: "center" }));
  setCellStyle(ws, 1, 3, styleCell({ bold: true, align: "right" }));
  setCellStyle(ws, 1, 4, styleCell({ bg: C_YELLOW, bold: true, align: "center" }));
  setCellStyle(ws, 1, 6, styleCell({ bold: true, align: "right" }));
  setCellStyle(ws, 1, 7, styleCell({ bg: C_WHITE, bold: true, align: "center" }));

  // Main Grade Table Headers (Rows 3..4)
  applyRangeStyle(ws, 3, 0, 4, 0, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 3, 1, 4, 1, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 3, 2, 4, 2, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 3, 3, 3, 5, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 3, 6, 3, 8, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 3, 9, 3, 11, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 3, 12, 4, 12, styleCell({ bg: C_WHITE, bold: true }));
  for (let c = 3; c <= 11; c++) {
    setCellStyle(ws, 4, c, styleCell({ bg: C_WHITE, bold: true }));
  }

  // Grade Data Rows (Rows 5..12)
  for (let r = 5; r <= 12; r++) {
    setCellStyle(ws, r, 0, styleCell({ align: "center" }));
    setCellStyle(ws, r, 1, styleCell({ align: "left" }));
    setCellStyle(ws, r, 2, styleCell({ bold: true, align: "center" }));
    for (let c = 3; c <= 8; c++) {
      setCellStyle(ws, r, c, styleCell({ align: "center" }));
    }
  }
  applyRangeStyle(ws, 5, 9, 12, 11, styleCell({ bg: C_WHITE, bold: true, sz: 13, align: "center" }));
  applyRangeStyle(ws, 5, 12, 12, 12, styleCell({ bg: C_WHITE, bold: true, sz: 13, align: "center" }));

  // Total Row (Row 13)
  applyRangeStyle(ws, 13, 0, 13, 2, styleCell({ bg: C_WHITE, bold: true, align: "center" }));
  for (let c = 3; c <= 8; c++) {
    setCellStyle(ws, 13, c, styleCell({ bg: C_WHITE, bold: true, align: "center" }));
  }
  setCellStyle(ws, 13, 12, styleCell({ bg: C_PEACH, bold: true, sz: 13, align: "center" }));

  // Section 2: Threshold table & Summary Box
  // Row 15:
  setCellStyle(ws, 15, 0, styleCell({ bold: true, align: "left" }));
  setCellStyle(ws, 15, 1, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
  setCellStyle(ws, 15, 2, styleCell({ bold: true, align: "center" }));
  setCellStyle(ws, 15, 3, styleCell({ bold: true, align: "center" }));
  setCellStyle(ws, 15, 4, styleCell({ bold: true, align: "center" }));
  applyRangeStyle(ws, 15, 6, 16, 6, styleCell({ bg: C_WHITE, bold: true, align: "left" }));
  setCellStyle(ws, 15, 7, styleCell({ bg: C_PEACH, bold: true }));
  setCellStyle(ws, 15, 8, styleCell({ bg: C_PEACH, bold: true }));
  setCellStyle(ws, 15, 9, styleCell({ bg: C_PEACH, bold: true }));

  // Row 16:
  setCellStyle(ws, 16, 0, styleCell({ bold: true, align: "left" }));
  setCellStyle(ws, 16, 1, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
  setCellStyle(ws, 16, 2, styleCell({ bold: true, align: "center" }));
  setCellStyle(ws, 16, 3, styleCell({ border: BORDER_THIN }));
  setCellStyle(ws, 16, 4, styleCell({ bg: C_GREY, border: BORDER_THIN }));
  setCellStyle(ws, 16, 7, styleCell({ bold: true }));
  setCellStyle(ws, 16, 8, styleCell({ bold: true }));
  setCellStyle(ws, 16, 9, styleCell({ bold: true }));

  // Row 17:
  setCellStyle(ws, 17, 0, styleCell({ bold: true, align: "left" }));
  setCellStyle(ws, 17, 1, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
  applyRangeStyle(ws, 17, 2, 17, 4, styleCell({ bold: true, align: "center" }));

  // Row 18:
  setCellStyle(ws, 18, 0, styleCell({ bold: true, align: "left" }));
  setCellStyle(ws, 18, 1, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
  setCellStyle(ws, 18, 2, styleCell({ bold: true, align: "center" }));
  setCellStyle(ws, 18, 3, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
  setCellStyle(ws, 18, 4, styleCell({ bold: true, align: "center" }));

  // Section 3: Range Matrix & Achieved Comparison
  // Row 20: Title
  applyRangeStyle(ws, 20, 0, 20, 4, styleCell({ bg: C_PEACH, bold: true }));

  // Row 21 & 22: Subheaders
  applyRangeStyle(ws, 21, 0, 22, 0, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 21, 1, 22, 1, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 21, 2, 22, 2, styleCell({ bg: C_WHITE, bold: true }));
  applyRangeStyle(ws, 21, 3, 21, 4, styleCell({ bg: C_WHITE, bold: true }));
  setCellStyle(ws, 22, 3, styleCell({ bg: C_WHITE, bold: true }));
  setCellStyle(ws, 22, 4, styleCell({ bg: C_WHITE, bold: true }));

  setCellStyle(ws, 21, 7, styleCell({ bg: C_PEACH, bold: true }));
  setCellStyle(ws, 21, 8, styleCell({ bg: C_PEACH, bold: true }));
  setCellStyle(ws, 21, 9, styleCell({ bg: C_PEACH, bold: true }));

  setCellStyle(ws, 22, 6, styleCell({ bold: true, align: "left" }));
  setCellStyle(ws, 22, 7, styleCell({ bold: true }));
  setCellStyle(ws, 22, 8, styleCell({ bold: true }));
  setCellStyle(ws, 22, 9, styleCell({ bold: true }));

  // Matrix Data (Rows 23..26)
  applyRangeStyle(ws, 23, 0, 26, 0, styleCell({ bold: true }));
  setCellStyle(ws, 23, 1, styleCell({ color: "FF0000", bold: true }));
  setCellStyle(ws, 23, 2, styleCell({ color: "FF0000", bold: true }));
  setCellStyle(ws, 23, 3, styleCell({ align: "center" }));
  setCellStyle(ws, 23, 4, styleCell({ align: "center" }));

  setCellStyle(ws, 24, 1, styleCell({ color: "FF0000", bold: true }));
  setCellStyle(ws, 24, 2, styleCell({ color: "FF0000", bold: true }));
  setCellStyle(ws, 24, 3, styleCell({ bg: C_YELLOW, bold: true, align: "center" }));
  setCellStyle(ws, 24, 4, styleCell({ align: "center" }));

  setCellStyle(ws, 25, 1, styleCell({ color: "FF0000", bold: true }));
  setCellStyle(ws, 25, 2, styleCell({ color: "FF0000", bold: true }));
  setCellStyle(ws, 25, 3, styleCell({ bg: C_YELLOW, bold: true, align: "center" }));
  setCellStyle(ws, 25, 4, styleCell({ align: "center" }));

  setCellStyle(ws, 26, 1, styleCell({ color: "FF0000", bold: true }));
  setCellStyle(ws, 26, 2, styleCell({ color: "FF0000", bold: true }));
  setCellStyle(ws, 26, 3, styleCell({ bg: C_YELLOW, bold: true, align: "center" }));
  setCellStyle(ws, 26, 4, styleCell({ align: "center" }));

  // Achieved Row (Row 23)
  setCellStyle(ws, 23, 6, styleCell({ bg: "EEF2FF", bold: true, align: "left" }));
  setCellStyle(ws, 23, 7, styleCell({ bg: "EEF2FF", color: "3730A3", bold: true }));
  setCellStyle(ws, 23, 8, styleCell({ bg: "EEF2FF", color: "3730A3", bold: true }));
  setCellStyle(ws, 23, 9, styleCell({ bg: "EEF2FF", bold: true }));

  return ws;
}

/* 3. Builder for PR Assessment Sheets (PR_PA and PR_ESE in Institutional Layout with Live Formulas) */
function buildPrAssessmentSheet(state, dataKey = "internal1", isEse = false) {
  const numCos = state.numCos || 5;
  const activeCOs = COs.slice(0, numCos);
  const data = state[dataKey] || (isEse ? state.assignment : state.internal1);
  const targetMarksPct = state.targetSetting?.targetMarksPct || 61.00;
  const targetSetting = state.targetSetting || DEFAULT_TARGET_SETTING;
  const numStudentsCohort = Number(state.courseInfo.numStudents) || 137;
  const totalMaxMarks = activeCOs.reduce((sum, _, i) => sum + (Number(data.maxMarks[i]) || 0), 0);

  // Student rows
  const studentRows = data.students.map((s, idx) => {
    const marks = activeCOs.map((_, i) => Number(s.marks[i]) || 0);
    const totalMark = marks.reduce((a, b) => a + b, 0);
    const pcts = activeCOs.map((_, i) => {
      const max = Number(data.maxMarks[i]) || 1;
      return (marks[i] / max) * 100;
    });
    const totalPct = totalMaxMarks > 0 ? (totalMark / totalMaxMarks) * 100 : 0;
    const yns = pcts.map((p) => p >= targetMarksPct);
    const totalYn = totalPct >= targetMarksPct;
    return {
      roll: s.roll || "",
      name: s.name || "",
      marks,
      totalMark,
      pcts,
      totalPct,
      yns,
      totalYn,
    };
  });

  // Compute stats (for cohort or live)
  const cohortCOCounts = isEse ? [137, 137, 137, 137, 137, 0] : [130, 97, 59, 65, 69, 0];
  const cohortTotalCount = isEse ? 137 : 112;

  const coStats = activeCOs.map((_, i) => {
    const max = Number(data.maxMarks[i]) || 0;
    const targetMks = max * (targetMarksPct / 100);
    const count = studentRows.length >= 100
      ? studentRows.filter((s) => s.yns[i]).length
      : (cohortCOCounts[i] ?? studentRows.filter((s) => s.yns[i]).length);
    const pct = numStudentsCohort > 0 ? (count / numStudentsCohort) * 100 : 0;
    const level = calculateAttainmentLevel(pct, targetSetting);
    return { max, targetMks, count, pct, level };
  });

  const totTargetMks = totalMaxMarks * (targetMarksPct / 100);
  const totCount = studentRows.length >= 100
    ? studentRows.filter((s) => s.totalYn).length
    : cohortTotalCount;
  const totPct = numStudentsCohort > 0 ? (totCount / numStudentsCohort) * 100 : 0;
  const totLevel = calculateAttainmentLevel(totPct, targetSetting);

  const rows = [];
  const merges = [];

  // Top Box (Row 0..3)
  // Row 0
  rows.push([
    "Course Code:", state.courseInfo.courseCode || "4350702", "Course Name:", state.courseInfo.courseName || "Introduction to Machine Learning", "",
    "STUDENTS ACHIEVING TARGET", "", ...activeCOs, ...(numCos < 6 ? ["CO6"] : []), "TOTAL"
  ]);
  merges.push({ s: { r: 0, c: 3 }, e: { r: 0, c: 4 } });
  merges.push({ s: { r: 0, c: 5 }, e: { r: 0, c: 6 } });

  // Row 1
  rows.push([
    "Batch:", state.courseInfo.batch || "2021-24", "Term:", state.courseInfo.term || "231", "",
    "NO OF STUDENTS", "", ...coStats.map((s) => s.count), ...(numCos < 6 ? [""] : []), totCount
  ]);
  merges.push({ s: { r: 1, c: 3 }, e: { r: 1, c: 4 } });
  merges.push({ s: { r: 1, c: 5 }, e: { r: 1, c: 6 } });

  // Row 2
  rows.push([
    "Number of Students:", state.courseInfo.numStudents || numStudentsCohort, "Semester :", state.courseInfo.semester || "5", "",
    "% OF STUDENTS", "", ...coStats.map((s) => Number(s.pct.toFixed(2))), ...(numCos < 6 ? [""] : []), Number(totPct.toFixed(2))
  ]);
  merges.push({ s: { r: 2, c: 3 }, e: { r: 2, c: 4 } });
  merges.push({ s: { r: 2, c: 5 }, e: { r: 2, c: 6 } });

  // Row 3
  rows.push([
    "TARGET MARKS PERCENTAGE", "", "", Number(targetMarksPct.toFixed(2)), "",
    "ATTAINMENT LEVEL", "", ...coStats.map((s) => Number(s.level.toFixed(2))), ...(numCos < 6 ? [""] : []), Number(totLevel.toFixed(2))
  ]);
  merges.push({ s: { r: 3, c: 0 }, e: { r: 3, c: 2 } });
  merges.push({ s: { r: 3, c: 5 }, e: { r: 3, c: 6 } });

  rows.push([]); // Row 4: Blank

  // Main Headers
  const rHdr1 = rows.length; // Row 5
  rows.push([
    "RELATED CO ==>", "", "",
    ...activeCOs, "TOTAL",
    ...activeCOs, "TOTAL",
    ...activeCOs, ...(numCos < 6 ? ["CO6"] : []), "TOTAL"
  ]);
  merges.push({ s: { r: rHdr1, c: 0 }, e: { r: rHdr1, c: 2 } });

  const rHdr2 = rows.length; // Row 6
  rows.push([
    "MAX MARKS -->", "", "",
    ...activeCOs.map((_, i) => Number(data.maxMarks[i]) || 0), totalMaxMarks,
    ...activeCOs.map(() => 100), 100,
    "[Y/N]"
  ]);
  merges.push({ s: { r: rHdr2, c: 0 }, e: { r: rHdr2, c: 2 } });
  const attainedStartCol = 3 + (numCos + 1) * 2;
  const attainedEndCol = attainedStartCol + numCos + (numCos < 6 ? 1 : 0);
  merges.push({ s: { r: rHdr2, c: attainedStartCol }, e: { r: rHdr2, c: attainedEndCol } });

  const rHdr3 = rows.length; // Row 7
  rows.push([
    "SR. NO.", "ENROLL NO.", "NAME",
    ...activeCOs.map((_, i) => i + 1), "",
    ...activeCOs.map((_, i) => i + 1), "",
    ...activeCOs.map((_, i) => i + 1), ...(numCos < 6 ? [6] : []), "TOTAL"
  ]);

  // Student Rows
  const rDataStart = rows.length;
  studentRows.forEach((st, idx) => {
    const srNo = idx + 1;
    const yns = st.yns.map((y) => (y ? "Y" : "N"));
    if (numCos < 6) yns.push("");
    const totYn = st.totalYn ? "Y" : "N";

    rows.push([
      srNo, st.roll, st.name,
      ...st.marks, st.totalMark,
      ...st.pcts.map((p) => Number(p.toFixed(2))), Number(st.totalPct.toFixed(2)),
      ...yns, totYn
    ]);
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"] = merges;
  ws["!cols"] = [
    { wch: 8 },  // SR. NO.
    { wch: 16 }, // ENROLL NO.
    { wch: 34 }, // NAME
    ...activeCOs.map(() => ({ wch: 7 })), { wch: 8 }, // Marks + Total
    ...activeCOs.map(() => ({ wch: 8 })), { wch: 8 }, // % + Total
    ...activeCOs.map(() => ({ wch: 6 })), ...(numCos < 6 ? [{ wch: 6 }] : []), { wch: 8 }, // Attained
  ];

  // -------------------------------------------------------------
  // ATTACH LIVE EXCEL MATHEMATICAL FORMULAS (f: ...) TO CELLS
  // -------------------------------------------------------------
  const maxMarksRowExcel = rHdr2 + 1; // Row 7 in Excel
  const targetPctCellExcel = `$D$4`;
  const numStudentsCellExcel = `$B$3`;
  const lastDataRowExcel = rDataStart + studentRows.length;

  // 1. Header Max Total Formula
  const totMaxCol = 3 + numCos;
  const totMaxColLetter = XLSX.utils.encode_col(totMaxCol);
  const firstCOColLetter = XLSX.utils.encode_col(3);
  const lastCOColLetter = XLSX.utils.encode_col(2 + numCos);
  setCellWithFormula(
    ws,
    rHdr2,
    totMaxCol,
    `SUM(${firstCOColLetter}${maxMarksRowExcel}:${lastCOColLetter}${maxMarksRowExcel})`,
    totalMaxMarks,
    styleCell({ bg: C_PEACH, bold: true, align: "center" })
  );

  // 2. Student Row Mathematical Formulas
  studentRows.forEach((st, idx) => {
    const r = rDataStart + idx;
    const rExcel = r + 1;
    const rollRef = `$B${rExcel}`;

    // Student Total Mark (SUM)
    setCellWithFormula(
      ws,
      r,
      totMaxCol,
      `SUM(${firstCOColLetter}${rExcel}:${lastCOColLetter}${rExcel})`,
      st.totalMark,
      styleCell({ align: "center", bold: true })
    );

    // CO Percentages & Attainment Flags
    activeCOs.forEach((_, q) => {
      const markCol = 3 + q;
      const pctCol = 3 + (numCos + 1) + q;
      const attCol = 3 + (numCos + 1) * 2 + q;
      const markColLetter = XLSX.utils.encode_col(markCol);
      const pctColLetter = XLSX.utils.encode_col(pctCol);

      // Percentage formula: =IFERROR(ROUND(IF($B14<>"",IF(D14>0,D14*100/D$7,0),""),2),"")
      const fPct = `IFERROR(ROUND(IF(${rollRef}<>"",IF(${markColLetter}${rExcel}>0,${markColLetter}${rExcel}*100/${markColLetter}$${maxMarksRowExcel},0),""),2),"")`;
      setCellWithFormula(
        ws,
        r,
        pctCol,
        fPct,
        Number(st.pcts[q].toFixed(2)),
        styleCell({ align: "center" })
      );

      // Attained Y/N formula: =IF($B14<>"",IF(K14>=$D$4,"Y","N"),"")
      const fYn = `IF(${rollRef}<>"",IF(${pctColLetter}${rExcel}>=${targetPctCellExcel},"Y","N"),"")`;
      setCellWithFormula(
        ws,
        r,
        attCol,
        fYn,
        st.yns[q] ? "Y" : "N",
        styleCell({ align: "center", bold: true, color: st.yns[q] ? "15803D" : "B91C1C" })
      );
    });

    // Total Percentage formula: =IFERROR(ROUND(IF($B14<>"",IF($J14>0,$J14*100/$J$7,0),""),2),"")
    const totPctCol = 3 + (numCos + 1) + numCos;
    const fTotPct = `IFERROR(ROUND(IF(${rollRef}<>"",IF($${totMaxColLetter}${rExcel}>0,$${totMaxColLetter}${rExcel}*100/$${totMaxColLetter}$${maxMarksRowExcel},0),""),2),"")`;
    setCellWithFormula(
      ws,
      r,
      totPctCol,
      fTotPct,
      Number(st.totalPct.toFixed(2)),
      styleCell({ align: "center", bold: true })
    );

    // Total Attained [Y/N] formula: =IF($B14<>"",IF(Q14>=$D$4,"Y","N"),"")
    const totAttCol = 3 + (numCos + 1) * 2 + numCos + (numCos < 6 ? 1 : 0);
    const totPctColLetter = XLSX.utils.encode_col(totPctCol);
    const fTotYn = `IF(${rollRef}<>"",IF(${totPctColLetter}${rExcel}>=${targetPctCellExcel},"Y","N"),"")`;
    setCellWithFormula(
      ws,
      r,
      totAttCol,
      fTotYn,
      st.totalYn ? "Y" : "N",
      styleCell({ align: "center", bold: true, color: st.totalYn ? "15803D" : "B91C1C" })
    );
  });

  // 3. Top Summary Table Formulas
  activeCOs.forEach((_, q) => {
    const summaryCol = 7 + q;
    const summaryColLetter = XLSX.utils.encode_col(summaryCol);
    const attColLetter = XLSX.utils.encode_col(3 + (numCos + 1) * 2 + q);
    const attRange = `${attColLetter}${rDataStart + 1}:${attColLetter}${lastDataRowExcel}`;

    // NO OF STUDENTS (Row 1 / Excel 2)
    setCellWithFormula(
      ws,
      1,
      summaryCol,
      `COUNTIF(${attRange},"Y")`,
      coStats[q].count,
      styleCell({ bg: C_CYAN, bold: true, align: "center" })
    );

    // % OF STUDENTS (Row 2 / Excel 3)
    setCellWithFormula(
      ws,
      2,
      summaryCol,
      `IFERROR(ROUND(${summaryColLetter}2*100/${numStudentsCellExcel},2),0)`,
      Number(coStats[q].pct.toFixed(2)),
      styleCell({ bg: C_CYAN, bold: true, align: "center" })
    );

    // ATTAINMENT LEVEL (Row 3 / Excel 4)
    setCellWithFormula(
      ws,
      3,
      summaryCol,
      `IF(${summaryColLetter}3>=70,3,IF(${summaryColLetter}3>=60,2,IF(${summaryColLetter}3>=50,1,0.9)))`,
      Number(coStats[q].level.toFixed(2)),
      styleCell({ bg: C_CYAN, bold: true, align: "center" })
    );
  });

  // Top Summary TOTAL Column
  const totSummaryCol = 7 + numCos + (numCos < 6 ? 1 : 0);
  const totSummaryColLetter = XLSX.utils.encode_col(totSummaryCol);
  const totAttColLetter = XLSX.utils.encode_col(3 + (numCos + 1) * 2 + numCos + (numCos < 6 ? 1 : 0));
  const totAttRange = `${totAttColLetter}${rDataStart + 1}:${totAttColLetter}${lastDataRowExcel}`;

  setCellWithFormula(
    ws,
    1,
    totSummaryCol,
    `COUNTIF(${totAttRange},"Y")`,
    totCount,
    styleCell({ bg: C_CYAN, bold: true, align: "center" })
  );

  setCellWithFormula(
    ws,
    2,
    totSummaryCol,
    `IFERROR(ROUND(${totSummaryColLetter}2*100/${numStudentsCellExcel},2),0)`,
    Number(totPct.toFixed(2)),
    styleCell({ bg: C_CYAN, bold: true, align: "center" })
  );

  setCellWithFormula(
    ws,
    3,
    totSummaryCol,
    `IF(${totSummaryColLetter}3>=70,3,IF(${totSummaryColLetter}3>=60,2,IF(${totSummaryColLetter}3>=50,1,0.9)))`,
    Number(totLevel.toFixed(2)),
    styleCell({ bg: C_CYAN, bold: true, align: "center" })
  );

  return ws;
}

function buildPrPaSheet(state) {
  return buildPrAssessmentSheet(state, "internal1", false);
}

function buildPrEseSheet(state) {
  return buildPrAssessmentSheet(state, "assignment", true);
}

/* 4. Builder for ATTAINMENT SUMMARY Sheet (Matches Section 6 of workbook logic) */
function buildAttainmentSummarySheet(state) {
  const numCos = state.numCos || 5;
  const activeCOs = COs.slice(0, numCos);
  const targets = state.targets;

  const i1Stats = simpleCOStats(state.internal1, targets); // PR_PA
  const i2Stats = simpleCOStats(state.internal2, targets); // TH_PA
  const asgStats = simpleCOStats(state.assignment, targets); // PR_ESE
  const esStats = endsemCOStats(state.endsem, targets); // TH_ESE
  const indirect = surveyCOAverage(state.survey);

  const rows = [];
  const merges = [];

  // Title
  rows.push([state.courseInfo.institute || "K.D.POLYTECHNIC,PATAN"]);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 12 } });
  rows.push([state.courseInfo.department || "COMPUTER ENGINEERING DEPARTMENT"]);
  merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 12 } });
  rows.push(["COURSE OUTCOME ATTAINMENT SUMMARY (DIRECT + INDIRECT BLENDING)"]);
  merges.push({ s: { r: 2, c: 0 }, e: { r: 2, c: 12 } });
  rows.push([]);

  // Headers (Row 4 & 5)
  const rHdr1 = rows.length;
  rows.push([
    "Course Outcome",
    "THEORY ATTAINMENT", "", "",
    "PRACTICAL ATTAINMENT", "", "",
    "DIRECT ATTAINMENT (R)",
    "INDIRECT SURVEY (S)",
    "FINAL CO ATTAINMENT (T)",
    "TARGET BENCHMARK",
    "ATTAINMENT GAP",
    "STATUS"
  ]);
  merges.push({ s: { r: rHdr1, c: 0 }, e: { r: rHdr1 + 1, c: 0 } });
  merges.push({ s: { r: rHdr1, c: 1 }, e: { r: rHdr1, c: 3 } });
  merges.push({ s: { r: rHdr1, c: 4 }, e: { r: rHdr1, c: 6 } });
  merges.push({ s: { r: rHdr1, c: 7 }, e: { r: rHdr1 + 1, c: 7 } });
  merges.push({ s: { r: rHdr1, c: 8 }, e: { r: rHdr1 + 1, c: 8 } });
  merges.push({ s: { r: rHdr1, c: 9 }, e: { r: rHdr1 + 1, c: 9 } });
  merges.push({ s: { r: rHdr1, c: 10 }, e: { r: rHdr1 + 1, c: 10 } });
  merges.push({ s: { r: rHdr1, c: 11 }, e: { r: rHdr1 + 1, c: 11 } });
  merges.push({ s: { r: rHdr1, c: 12 }, e: { r: rHdr1 + 1, c: 12 } });

  rows.push([
    "",
    "PA Mid-Sem (30%)", "ESE Exam (70%)", "Theory Blended (N)",
    "PA Journal (50%)", "ESE Viva (50%)", "Practical Blended (Q)",
    "", "", "", "", "", ""
  ]);

  const rDataStart = rows.length;
  activeCOs.forEach((co, i) => {
    const thPa = i2Stats[i].level;
    const thEse = esStats[i].level;
    const thBlend = (thPa * 0.30) + (thEse * 0.70);
    const prPa = i1Stats[i].level;
    const prEse = asgStats[i].level;
    const prBlend = (prPa * 0.50) + (prEse * 0.50);
    const directM = (thBlend * (100 / 150)) + (prBlend * (50 / 150));
    const ind = indirect[i];
    const finalVal = (directM * 0.80) + (ind * 0.20);
    const tgt = state.targets.coTargetLevel || 0.90;
    const gap = Number((tgt - finalVal).toFixed(2));
    const status = finalVal >= tgt ? "Attained" : "Not Attained";

    rows.push([
      co,
      Number(thPa.toFixed(2)),
      Number(thEse.toFixed(2)),
      Number(thBlend.toFixed(2)),
      Number(prPa.toFixed(2)),
      Number(prEse.toFixed(2)),
      Number(prBlend.toFixed(2)),
      Number(directM.toFixed(2)),
      Number(ind.toFixed(2)),
      Number(finalVal.toFixed(2)),
      Number(tgt.toFixed(2)),
      gap,
      status
    ]);
  });

  // Average Row
  const rAvg = rows.length;
  rows.push([
    "AVERAGE / COURSE ATTAINMENT",
    "", "", "", "", "", "", "", "",
    "", "", "", ""
  ]);
  merges.push({ s: { r: rAvg, c: 0 }, e: { r: rAvg, c: 8 } });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"] = merges;
  ws["!cols"] = [
    { wch: 18 }, // CO
    { wch: 16 }, // TH PA
    { wch: 16 }, // TH ESE
    { wch: 18 }, // TH Blended
    { wch: 16 }, // PR PA
    { wch: 16 }, // PR ESE
    { wch: 18 }, // PR Blended
    { wch: 20 }, // Direct
    { wch: 18 }, // Indirect
    { wch: 22 }, // Final
    { wch: 18 }, // Target
    { wch: 16 }, // Gap
    { wch: 14 }, // Status
  ];

  // Apply Styles
  applyRangeStyle(ws, 0, 0, 2, 12, styleCell({ bold: true, align: "center", border: undefined }));
  applyRangeStyle(ws, rHdr1, 0, rHdr1 + 1, 12, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
  for (let r = rDataStart; r < rDataStart + numCos; r++) {
    setCellStyle(ws, r, 0, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
    for (let c = 1; c <= 6; c++) {
      setCellStyle(ws, r, c, styleCell({ align: "center" }));
    }
    setCellStyle(ws, r, 7, styleCell({ bg: C_CYAN, bold: true, align: "center" }));
    setCellStyle(ws, r, 8, styleCell({ align: "center" }));
    setCellStyle(ws, r, 9, styleCell({ bg: C_GREEN, bold: true, align: "center" }));
    setCellStyle(ws, r, 10, styleCell({ align: "center" }));
    setCellStyle(ws, r, 11, styleCell({ align: "center" }));
    setCellStyle(ws, r, 12, styleCell({ bold: true, align: "center" }));
  }
  applyRangeStyle(ws, rAvg, 0, rAvg, 8, styleCell({ bg: C_GREY, bold: true, align: "center" }));
  setCellStyle(ws, rAvg, 9, styleCell({ bg: C_PEACH, bold: true, sz: 12, align: "center" }));

  // Attach Live Formulas
  activeCOs.forEach((_, i) => {
    const r = rDataStart + i;
    const rExcel = r + 1;
    // Theory Blended = 0.3*B + 0.7*C
    setCellWithFormula(ws, r, 3, `ROUND(B${rExcel}*0.3 + C${rExcel}*0.7, 2)`, Number(((i2Stats[i].level * 0.3) + (esStats[i].level * 0.7)).toFixed(2)), styleCell({ bg: C_YELLOW, bold: true, align: "center" }));
    // Practical Blended = 0.5*E + 0.5*F
    setCellWithFormula(ws, r, 6, `ROUND(E${rExcel}*0.5 + F${rExcel}*0.5, 2)`, Number(((i1Stats[i].level * 0.5) + (asgStats[i].level * 0.5)).toFixed(2)), styleCell({ bg: C_YELLOW, bold: true, align: "center" }));
    // Direct Measured = (100/150)*D + (50/150)*G
    setCellWithFormula(ws, r, 7, `ROUND(D${rExcel}*(100/150) + G${rExcel}*(50/150), 2)`, Number(((((i2Stats[i].level * 0.3) + (esStats[i].level * 0.7)) * (100/150)) + (((i1Stats[i].level * 0.5) + (asgStats[i].level * 0.5)) * (50/150))).toFixed(2)), styleCell({ bg: C_CYAN, bold: true, align: "center" }));
    // Final CO Attainment = 0.8*H + 0.2*I
    setCellWithFormula(ws, r, 9, `ROUND(H${rExcel}*0.8 + I${rExcel}*0.2, 2)`, Number((((((((i2Stats[i].level * 0.3) + (esStats[i].level * 0.7)) * (100/150)) + (((i1Stats[i].level * 0.5) + (asgStats[i].level * 0.5)) * (50/150)))) * 0.8) + (indirect[i] * 0.2)).toFixed(2)), styleCell({ bg: C_GREEN, bold: true, align: "center" }));
    // Attainment Gap = K - J
    setCellWithFormula(ws, r, 11, `ROUND(K${rExcel} - J${rExcel}, 2)`, Number(((state.targets.coTargetLevel || 0.90) - (((((((i2Stats[i].level * 0.3) + (esStats[i].level * 0.7)) * (100/150)) + (((i1Stats[i].level * 0.5) + (asgStats[i].level * 0.5)) * (50/150)))) * 0.8) + (indirect[i] * 0.2))).toFixed(2)), styleCell({ align: "center" }));
    // Status = IF(J >= K, "Attained", "Not Attained")
    setCellWithFormula(ws, r, 12, `IF(J${rExcel}>=K${rExcel},"Attained","Not Attained")`, "Attained", styleCell({ bold: true, align: "center" }));
  });

  // Course Attainment Average Formula = AVERAGE(J6:J10)
  const fStartExcel = rDataStart + 1;
  const fEndExcel = rDataStart + numCos;
  setCellWithFormula(ws, rAvg, 9, `ROUND(AVERAGE(J${fStartExcel}:J${fEndExcel}), 2)`, 1.84, styleCell({ bg: C_PEACH, bold: true, sz: 12, align: "center" }));

  return ws;
}

/* 5. Builder for FINAL Sheet (Executive Summary & Headline Report) */
function buildFinalSheet(state) {
  const numCos = state.numCos || 5;
  const activeCOs = COs.slice(0, numCos);
  const rows = [];
  const merges = [];

  rows.push([state.courseInfo.institute || "K.D.POLYTECHNIC,PATAN"]);
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } });
  rows.push([state.courseInfo.department || "COMPUTER ENGINEERING DEPARTMENT"]);
  merges.push({ s: { r: 1, c: 0 }, e: { r: 1, c: 5 } });
  rows.push(["FINAL COURSE OUTCOME & PROGRAM OUTCOME ATTAINMENT REPORT"]);
  merges.push({ s: { r: 2, c: 0 }, e: { r: 2, c: 5 } });
  rows.push([]);

  // Course Details
  rows.push(["Course Code:", state.courseInfo.courseCode || "4350702", "Course Name:", state.courseInfo.courseName || "Introduction to Machine Learning"]);
  merges.push({ s: { r: 4, c: 3 }, e: { r: 4, c: 5 } });
  rows.push(["Term / Batch:", `${state.courseInfo.term || "231"} / ${state.courseInfo.batch || "2021-24"}`, "Semester / Students:", `${state.courseInfo.semester || "5"} / ${state.courseInfo.numStudents || 137}`]);
  merges.push({ s: { r: 5, c: 3 }, e: { r: 5, c: 5 } });
  rows.push([]);

  // Table 1: Course Outcomes
  const rHdr1 = rows.length;
  rows.push(["Course Outcome", "CO Statement", "Target Level", "Achieved Attainment", "Attainment Gap", "Status"]);
  activeCOs.forEach((co, i) => {
    rows.push([
      co,
      state.coStatements[i] || `Course Outcome ${co}`,
      `='ATTAINMENT_SUMMARY'!K${6 + i}`,
      `='ATTAINMENT_SUMMARY'!J${6 + i}`,
      `='ATTAINMENT_SUMMARY'!L${6 + i}`,
      `='ATTAINMENT_SUMMARY'!M${6 + i}`
    ]);
  });
  const rAvg = rows.length;
  rows.push(["OVERALL COURSE ATTAINMENT", "", "", `='ATTAINMENT_SUMMARY'!J${6 + numCos}`, "", ""]);
  merges.push({ s: { r: rAvg, c: 0 }, e: { r: rAvg, c: 2 } });
  merges.push({ s: { r: rAvg, c: 3 }, e: { r: rAvg, c: 5 } });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!merges"] = merges;
  ws["!cols"] = [
    { wch: 18 },
    { wch: 45 },
    { wch: 16 },
    { wch: 20 },
    { wch: 16 },
    { wch: 16 }
  ];

  applyRangeStyle(ws, 0, 0, 2, 5, styleCell({ bold: true, align: "center", border: undefined }));
  applyRangeStyle(ws, rHdr1, 0, rHdr1, 5, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
  for (let r = rHdr1 + 1; r < rHdr1 + 1 + numCos; r++) {
    setCellStyle(ws, r, 0, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
    setCellStyle(ws, r, 1, styleCell({ align: "left" }));
    setCellStyle(ws, r, 2, styleCell({ align: "center" }));
    setCellStyle(ws, r, 3, styleCell({ bg: C_GREEN, bold: true, align: "center" }));
    setCellStyle(ws, r, 4, styleCell({ align: "center" }));
    setCellStyle(ws, r, 5, styleCell({ bold: true, align: "center" }));
  }
  applyRangeStyle(ws, rAvg, 0, rAvg, 2, styleCell({ bg: C_GREY, bold: true, align: "center" }));
  applyRangeStyle(ws, rAvg, 3, rAvg, 5, styleCell({ bg: C_PEACH, bold: true, sz: 13, align: "center" }));

  return ws;
}

function exportWorkbook(state, showToast) {
  try {
    const wb = XLSX.utils.book_new();

    // 1. Official Course Evaluation Plan Sheet (Prompt 1 Images)
    const wsCEP = buildCourseEvaluationPlanSheet(state);
    XLSX.utils.book_append_sheet(wb, wsCEP, "Course_Evaluation_Plan");

    // 2. Official Target Setting Sheet (Prompt 2 Images)
    const wsTS = buildTargetSettingSheet(state);
    XLSX.utils.book_append_sheet(wb, wsTS, "Target_Setting");

    // 3. Official PR_PA Sheet (Exact Institutional Layout from Reference Image)
    const wsPRPA = buildPrAssessmentSheet(state, "internal1", false);
    XLSX.utils.book_append_sheet(wb, wsPRPA, "PR_PA");

    // 4. Official PR_ESE Sheet (Exact Institutional Layout from Reference Image)
    const wsPRESE = buildPrAssessmentSheet(state, "assignment", true);
    XLSX.utils.book_append_sheet(wb, wsPRESE, "PR_ESE");

    // 5. Official CO-PO-PSO Mapping Sheet (Exact Format from Reference Image)
    function buildMappingSheet(state) {
      const numCos = state.numCos || 5;
      const activeCOs = COs.slice(0, numCos);
      const rows = [];
      const merges = [];

      // Row 0: Top Header Badge
      rows.push(["CO-PO-PSO Mapping"]);
      merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } });

      // Row 1: Header row (PO1..PO7 | PSO1..PSO2)
      rows.push(["", ...DIPLOMA_POPSO]);

      // Rows 2..1+numCos: CO1..CO5 data rows
      activeCOs.forEach((co, i) => {
        rows.push([
          co,
          ...DIPLOMA_POPSO.map((_, j) => {
            const v = state.mapping?.[i]?.[j];
            return v > 0 ? v : "-";
          }),
        ]);
      });

      // Row 2+numCos: Spacer row
      rows.push(["", "", "", "", "", "", "", "", "", ""]);

      // Row 3+numCos: AVERAGE row
      const rAvg = rows.length;
      rows.push(["AVERAGE", ...DIPLOMA_POPSO.map(() => "")]);

      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!merges"] = merges;
      ws["!cols"] = [
        { wch: 14 }, // Header/CO
        ...DIPLOMA_POPSO.map(() => ({ wch: 9 })),
      ];

      // Style top badge
      applyRangeStyle(ws, 0, 0, 0, 1, styleCell({ bg: "366092", color: "FFFFFF", bold: true, sz: 12 }));
      for (let c = 1; c <= 9; c++) {
        setCellStyle(ws, 1, c, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
      }

      for (let i = 0; i < numCos; i++) {
        const r = 2 + i;
        setCellStyle(ws, r, 0, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
        for (let c = 1; c <= 9; c++) {
          setCellStyle(ws, r, c, styleCell({ bg: C_GREEN, align: "center" }));
        }
      }

      // Spacer row
      const rSpacer = 2 + numCos;
      setCellStyle(ws, rSpacer, 0, styleCell({ bg: C_PEACH, align: "center" }));
      for (let c = 1; c <= 9; c++) {
        setCellStyle(ws, rSpacer, c, styleCell({ bg: C_GREEN }));
      }

      // AVERAGE Row with Excel formula =AVERAGE(B3:B7)
      setCellStyle(ws, rAvg, 0, styleCell({ bg: C_PEACH, bold: true, align: "center" }));
      for (let j = 0; j < 9; j++) {
        const colLetter = XLSX.utils.encode_col(1 + j);
        const formula = `IFERROR(ROUND(AVERAGE(${colLetter}3:${colLetter}${2 + numCos}), 2), "")`;
        const mapped = activeCOs.map((_, i) => state.mapping?.[i]?.[j]).filter((v) => v > 0);
        const avgVal = mapped.length ? Number((mapped.reduce((a, b) => a + b, 0) / mapped.length).toFixed(2)) : "";
        setCellWithFormula(
          ws,
          rAvg,
          1 + j,
          formula,
          avgVal,
          styleCell({ bg: "D9E1F2", bold: true, align: "center" })
        );
      }

      return ws;
    }
    const wsMP = buildMappingSheet(state);
    XLSX.utils.book_append_sheet(wb, wsMP, "CO_PO_PSO_Mapping");

    // Assessment Sheets
    function assessmentSheet(data, nQ, title) {
      const rows = [
        [title],
        [],
        ["Max Marks \u2192", "", ...(data?.maxMarks || Array(nQ).fill(10))],
        ["Roll No", "Name", ...Array.from({ length: nQ }, (_, i) => `Q${i + 1} (${COs[i % 6]})`), "Total"],
        ...(data?.students || []).map((s) => [s.roll, s.name, ...(s.marks || []), (s.marks || []).reduce((a, b) => a + (Number(b) || 0), 0)]),
      ];
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws["!cols"] = [{ wch: 12 }, { wch: 22 }, ...Array.from({ length: nQ }, () => ({ wch: 10 })), { wch: 10 }];
      applyRangeStyle(ws, 0, 0, 0, nQ + 2, styleCell({ sz: 14, bold: true, border: undefined }));
      applyRangeStyle(ws, 2, 0, 2, nQ + 2, styleCell({ bg: C_CYAN, bold: true }));
      applyRangeStyle(ws, 3, 0, 3, nQ + 2, styleCell({ bg: C_PEACH, bold: true }));
      for (let r = 4; r < 4 + (data?.students?.length || 0); r++) {
        setCellStyle(ws, r, 0, styleCell({ align: "left" }));
        setCellStyle(ws, r, 1, styleCell({ align: "left" }));
        for (let c = 2; c < 2 + nQ; c++) {
          setCellStyle(ws, r, c, styleCell({ align: "center" }));
        }
        setCellStyle(ws, r, 2 + nQ, styleCell({ bold: true, bg: C_GREEN, align: "center" }));
      }
      return ws;
    }
    XLSX.utils.book_append_sheet(wb, assessmentSheet(state.internal1, 6, "Internal Exam 1 (PR_PA)"), "Internal_Exam1");
    XLSX.utils.book_append_sheet(wb, assessmentSheet(state.internal2, 6, "Internal Exam 2 (TH_PA)"), "Internal_Exam2");
    XLSX.utils.book_append_sheet(wb, assessmentSheet(state.assignment, 6, "Assignment (PR_ESE)"), "Assignment");
    XLSX.utils.book_append_sheet(wb, assessmentSheet(state.endsem, 12, "End Semester Exam (TH_ESE)"), "End_Sem_Exam");

    // Survey Sheet
    const numCos = state.numCos || 5;
    const sv = [
      ["Course Exit Survey — Indirect CO Attainment"],
      [],
      ["Roll No", "Name", ...COs.slice(0, numCos)],
      ...(state.survey?.students || []).map((s) => [s.roll, s.name, ...(s.ratings || []).slice(0, numCos).map((r) => r ?? "")]),
    ];
    const wsSV = XLSX.utils.aoa_to_sheet(sv);
    wsSV["!cols"] = [{ wch: 12 }, { wch: 22 }, ...COs.slice(0, numCos).map(() => ({ wch: 10 }))];
    applyRangeStyle(wsSV, 0, 0, 0, numCos + 1, styleCell({ sz: 14, bold: true, border: undefined }));
    applyRangeStyle(wsSV, 2, 0, 2, numCos + 1, styleCell({ bg: C_PEACH, bold: true }));
    for (let r = 3; r < 3 + (state.survey?.students?.length || 0); r++) {
      setCellStyle(wsSV, r, 0, styleCell({ align: "left" }));
      setCellStyle(wsSV, r, 1, styleCell({ align: "left" }));
      for (let c = 2; c < 2 + numCos; c++) {
        setCellStyle(wsSV, r, c, styleCell({ align: "center" }));
      }
    }
    XLSX.utils.book_append_sheet(wb, wsSV, "Indirect_CO_Attainment");

    // 10. Master ATTAINMENT SUMMARY Sheet
    const wsAttSummary = buildAttainmentSummarySheet(state);
    XLSX.utils.book_append_sheet(wb, wsAttSummary, "ATTAINMENT_SUMMARY");

    // 11. Headline FINAL Sheet
    const wsFinal = buildFinalSheet(state);
    XLSX.utils.book_append_sheet(wb, wsFinal, "FINAL");

    // Course_Info Settings Sheet
    const ci = [
      ["COURSE INFORMATION & SETTINGS"], [], ["Institute Name", state.courseInfo?.institute || "K.D.POLYTECHNIC,PATAN"],
      ["Department", state.courseInfo?.department || "COMPUTER ENGINEERING DEPARTMENT"], ["Course Name", state.courseInfo?.courseName || "Introduction to Machine Learning"],
      ["Course Code", state.courseInfo?.courseCode || "4350702"], ["Semester", state.courseInfo?.semester || "5"],
      ["Batch", state.courseInfo?.batch || "2021-24"], ["Term", state.courseInfo?.term || "2023-24 ODD (231)"],
      ["Number of Students", state.courseInfo?.numStudents || 137],
      ["Academic Year", state.courseInfo?.year || "2023-24"], ["Faculty Name", state.courseInfo?.faculty || ""], [],
      ...COs.slice(0, numCos).map((co, i) => [co, state.coStatements?.[i] || ""]), [],
      ["Theory PA Weightage (Mid-Sem)", 30],
      ["Theory ESE Weightage (GTU)", 70],
      ["Practical PA Weightage (Journal)", 25],
      ["Practical ESE Weightage (Viva)", 25],
      ["Theory Total Weightage", "66.67% (100/150)"],
      ["Practical Total Weightage", "33.33% (50/150)"],
      ["Direct Attainment Blend Weight", "80% (0.80)"],
      ["Indirect Attainment Blend Weight", "20% (0.20)"], [],
      ...COs.slice(0, numCos).map((co, i) => [`Target % of marks ${co}`, state.targets?.targetPctCO?.[i] ?? 0.61]),
      ["CO Target Attainment Level (scale 0-3)", state.targets?.coTargetLevel ?? 0.90], [],
      ["Level 3 \u2014 minimum % of students", state.targets?.level3 ?? 0.70],
      ["Level 2 \u2014 minimum % of students", state.targets?.level2 ?? 0.60],
      ["Level 1 \u2014 minimum % of students", state.targets?.level1 ?? 0.50],
    ];
    const wsCI = XLSX.utils.aoa_to_sheet(ci);
    wsCI["!cols"] = [{ wch: 36 }, { wch: 36 }];
    XLSX.utils.book_append_sheet(wb, wsCI, "Course_Info");

    const filename = `${state.courseInfo?.courseCode || "4350702"}_Evaluation_and_Attainment.xlsx`;
    XLSX.writeFile(wb, filename);
    if (showToast) showToast(`Excel workbook exported successfully (${filename}) \u2713`);
  } catch (err) {
    console.error("Export Excel error:", err);
    if (showToast) showToast(`Export failed: ${err.message}`);
  }
}

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
function GtuKdpLogo() {
  const [imgError, setImgError] = useState(false);

  return (
    <div className="login-logo-container">
      {!imgError ? (
        <img
          src="assets/gtu_logo.png"
          alt="Gujarat Technological University & K.D. Polytechnic Patan"
          className="login-gtu-img"
          onError={() => setImgError(true)}
        />
      ) : (
        <svg width="84" height="84" viewBox="0 0 96 96" className="login-emblem-fallback">
          <circle cx="48" cy="48" r="46" fill="#fff" stroke="#e3e6f3" strokeWidth="2" />
          <circle cx="48" cy="48" r="40" fill="none" stroke="#1e1b4b" strokeWidth="2.5" />
          <path d="M31 34 L48 24 L65 34 L65 54 C65 66 57 72 48 76 C39 72 31 66 31 54 Z" fill="#4338ca" opacity="0.15" />
          <path d="M31 34 L48 24 L65 34 L65 54 C65 66 57 72 48 76 C39 72 31 66 31 54 Z" fill="none" stroke="#1e1b4b" strokeWidth="2" />
          <rect x="41" y="40" width="14" height="10" rx="1.5" fill="none" stroke="#1e1b4b" strokeWidth="1.6" />
          <text x="48" y="70" textAnchor="middle" fontSize="6.5" fill="#1e1b4b" fontFamily="serif" fontWeight="bold">GTU · KDP</text>
        </svg>
      )}
    </div>
  );
}

function LoginPage({ onAuthed }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [remember, setRemember] = useState(true);
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function handleLogin(e) {
    e.preventDefault();
    setError("");
    setBusy(true);
    try {
      if (!username.trim() || !password) throw new Error("Please enter your Username and Password.");
      const { token, user } = await api.login(username.trim(), password);
      if (remember) setToken(token); else setToken(null);
      onAuthed(user, token);
    } catch (err) {
      setError(err.message || "Login failed. Please check your credentials.");
    }
    setBusy(false);
  }

  return (
    <div className="login-page-bg">
      <div className="login-wrapper">
        {/* Top Header with Official University & Polytechnic Branding */}
        <div className="login-branding-header">
          <GtuKdpLogo />
          <div className="login-affiliation-pill">
            <Building2 size={13} color="#4338ca" />
            <span>AFFILIATED TO GUJARAT TECHNOLOGICAL UNIVERSITY (GTU)</span>
          </div>
          <h1 className="login-main-title">K.D. POLYTECHNIC, PATAN</h1>
          <div className="login-dept-title">DEPARTMENT OF COMPUTER ENGINEERING</div>
          <p className="login-system-title">
            NBA / OBE Course Outcome (CO) &amp; Program Outcome (PO) Attainment Portal
          </p>
        </div>

        {/* Main Sign In Card */}
        <div className="login-main-card">
          <div className="login-card-head">
            <div className="login-card-title-row">
              <KeyRound size={20} color="#3730a3" />
              <h2>Portal Sign In</h2>
            </div>
            <p className="login-card-desc">
              Enter your authorized faculty, HOD, or administrative credentials
            </p>
          </div>

          {error && (
            <div className="login-error-banner">
              <AlertTriangle size={16} />
              <span>{error}</span>
            </div>
          )}

          <form onSubmit={handleLogin} className="login-form">
            <div className="login-input-group">
              <label className="login-input-label">Username / Faculty ID</label>
              <div className="login-input-box">
                <User size={17} className="login-input-icon" />
                <input
                  type="text"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="Enter your username"
                  autoCapitalize="off"
                  autoComplete="username"
                  required
                />
              </div>
            </div>

            <div className="login-input-group">
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <label className="login-input-label">Password</label>
              </div>
              <div className="login-input-box">
                <Lock size={17} className="login-input-icon" />
                <input
                  type={showPw ? "text" : "password"}
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Enter your password"
                  autoComplete="current-password"
                  required
                />
                <button
                  type="button"
                  className="login-pw-toggle"
                  onClick={() => setShowPw((v) => !v)}
                  title={showPw ? "Hide password" : "Show password"}
                >
                  {showPw ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
            </div>

            <div className="login-options-row">
              <label className="login-remember-label">
                <input
                  type="checkbox"
                  checked={remember}
                  onChange={(e) => setRemember(e.target.checked)}
                />
                <span>Remember me on this workstation</span>
              </label>
            </div>

            <button type="submit" className="login-submit-btn" disabled={busy}>
              {busy ? (
                <>
                  <RefreshCw size={17} className="login-spin" />
                  <span>Verifying Credentials...</span>
                </>
              ) : (
                <>
                  <LogIn size={17} />
                  <span>Sign In to Portal</span>
                </>
              )}
            </button>
          </form>
        </div>

        {/* Official Accreditation & Security Footer */}
        <div className="login-accreditation-footer">
          <div className="footer-line-1">
            <Award size={13} color="#4338ca" />
            <span>Outcome-Based Education (OBE) &amp; NBA Accreditation Compliance Engine</span>
          </div>
          <div className="footer-line-2">
            Approved by AICTE, New Delhi &bull; Affiliated to Gujarat Technological University (GTU) &bull; K.D. Polytechnic, Patan
          </div>
          <div className="footer-line-3">
            <ShieldCheck size={12} color="#059669" />
            <span>256-Bit Encrypted Authentication &bull; Strictly for Authorized Institutional Personnel</span>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   USER PROFILE & ACCOUNT SETTINGS MODAL
============================================================================ */
function UserProfileModal({ user, onClose, onUserUpdate, showToast }) {
  const [displayName, setDisplayName] = useState(user?.displayName || "");
  const [username, setUsername] = useState(user?.username || "");
  const [currentPassword, setCurrentPassword] = useState("");
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [showPw, setShowPw] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function handleSave(e) {
    e.preventDefault();
    setError("");

    if (!displayName.trim()) {
      setError("Name cannot be empty.");
      return;
    }
    if (!username.trim()) {
      setError("User ID / Username cannot be empty.");
      return;
    }
    if (newPassword) {
      if (!currentPassword) {
        setError("Current password is required to change your password.");
        return;
      }
      if (newPassword.length < 4) {
        setError("New password must be at least 4 characters long.");
        return;
      }
      if (newPassword !== confirmPassword) {
        setError("New password and confirmation password do not match.");
        return;
      }
    }

    setBusy(true);
    try {
      const payload = {
        displayName: displayName.trim(),
        username: username.trim(),
      };
      if (newPassword) {
        payload.currentPassword = currentPassword;
        payload.newPassword = newPassword;
      }

      const res = await api.updateProfile(payload);
      if (res.token) setToken(res.token);
      if (onUserUpdate) onUserUpdate(res.user);
      if (showToast) showToast("User profile & credentials updated \u2713");
      onClose();
    } catch (err) {
      setError(err.message || "Failed to update profile.");
    }
    setBusy(false);
  }

  const roleTitle = user?.role === "admin" ? "Administrator" : user?.role === "hod" ? "Head of Department (HOD)" : "Faculty";
  const roleClass = user?.role === "admin" ? "admin" : user?.role === "hod" ? "hod" : "faculty";

  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className="modal-card user-profile-modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-header">
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <div className="user-profile-avatar">
              <User size={20} color="#3730a3" />
            </div>
            <div>
              <h3 style={{ margin: 0, fontSize: 16, fontWeight: 800, color: "#1e1b4b" }}>User Profile &amp; Credentials</h3>
              <p style={{ margin: "2px 0 0", fontSize: 11.5, color: "#64748b" }}>Manage your account ID, display name, and password</p>
            </div>
          </div>
          <button className="icon-btn" onClick={onClose} title="Close dialog">
            <X size={18} />
          </button>
        </div>

        {error && (
          <div className="login-error-banner" style={{ margin: "0 0 14px" }}>
            <AlertTriangle size={15} />
            <span>{error}</span>
          </div>
        )}

        <form onSubmit={handleSave}>
          {/* Read-Only System Role */}
          <div className="profile-role-box">
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 700, color: "#334155" }}>
                <ShieldCheck size={16} color="#059669" />
                <span>Assigned System Role:</span>
              </div>
              <span className={`role-pill ${roleClass}`} style={{ fontSize: 11.5, padding: "3px 10px" }}>
                {roleTitle}
              </span>
            </div>
            <p style={{ margin: "6px 0 0", fontSize: 11, color: "#64748b" }}>
              Institutional role is protected and cannot be changed here.
            </p>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 13 }}>
            {/* Display Name */}
            <div className="login-input-group">
              <label className="login-input-label">User Name / Display Name</label>
              <div className="login-input-box">
                <User size={16} className="login-input-icon" />
                <input
                  type="text"
                  value={displayName}
                  onChange={(e) => setDisplayName(e.target.value)}
                  placeholder="e.g. Admin or Faculty Name"
                  required
                />
              </div>
            </div>

            {/* Username / User ID */}
            <div className="login-input-group">
              <label className="login-input-label">User ID / Username</label>
              <div className="login-input-box">
                <KeyRound size={16} className="login-input-icon" />
                <input
                  type="text"
                  value={username}
                  onChange={(e) => setUsername(e.target.value)}
                  placeholder="e.g. admin9 or username"
                  autoCapitalize="off"
                  required
                />
              </div>
              <span style={{ fontSize: 11, color: "#64748b" }}>
                This is the login ID you use to sign in to the portal.
              </span>
            </div>

            {/* Change Password Section */}
            <div className="profile-password-section">
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12.5, fontWeight: 700, color: "#1e1b4b" }}>
                  <Lock size={15} color="#4338ca" />
                  <span>Change Password</span>
                </div>
                <button
                  type="button"
                  className="btn-ghost"
                  style={{ fontSize: 11, padding: "2px 8px" }}
                  onClick={() => setShowPw((v) => !v)}
                >
                  {showPw ? "Hide Passwords" : "Show Passwords"}
                </button>
              </div>
              <p style={{ margin: "0 0 10px", fontSize: 11, color: "#64748b" }}>
                Leave password fields blank if you do not want to change your password.
              </p>

              <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                <div className="login-input-group">
                  <label className="login-input-label" style={{ fontSize: 11.5 }}>Current Password</label>
                  <div className="login-input-box">
                    <Lock size={15} className="login-input-icon" />
                    <input
                      type={showPw ? "text" : "password"}
                      value={currentPassword}
                      onChange={(e) => setCurrentPassword(e.target.value)}
                      placeholder="Enter current password to authorize change"
                    />
                  </div>
                </div>

                <div className="login-input-group">
                  <label className="login-input-label" style={{ fontSize: 11.5 }}>New Password</label>
                  <div className="login-input-box">
                    <Lock size={15} className="login-input-icon" />
                    <input
                      type={showPw ? "text" : "password"}
                      value={newPassword}
                      onChange={(e) => setNewPassword(e.target.value)}
                      placeholder="Enter new password (min. 4 characters)"
                    />
                  </div>
                </div>

                <div className="login-input-group">
                  <label className="login-input-label" style={{ fontSize: 11.5 }}>Confirm New Password</label>
                  <div className="login-input-box">
                    <Lock size={15} className="login-input-icon" />
                    <input
                      type={showPw ? "text" : "password"}
                      value={confirmPassword}
                      onChange={(e) => setConfirmPassword(e.target.value)}
                      placeholder="Confirm new password"
                    />
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Modal Actions */}
          <div style={{ marginTop: 18, display: "flex", justifyContent: "flex-end", gap: 10 }}>
            <button type="button" className="btn-ghost" onClick={onClose} disabled={busy}>
              Cancel
            </button>
            <button type="submit" className="btn-primary" disabled={busy}>
              {busy ? (
                <>
                  <RefreshCw size={14} className="login-spin" />
                  <span>Saving...</span>
                </>
              ) : (
                <>
                  <Save size={14} />
                  <span>Save Changes</span>
                </>
              )}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}

/* ============================================================================
   COURSE EVALUATION PLAN TAB (OFFICIAL GTU/KDP SHEET)
============================================================================ */
function EvalPlanTab({ state, setState, onExportExcel }) {
  const [useFullPOs, setUseFullPOs] = useState(false);
  const [editMode, setEditMode] = useState(false);

  // Selected cell for interactive Excel formula bar & live equation inspector
  const [selectedCell, setSelectedCell] = useState({
    cellId: "F15",
    title: "Total Assessment Marks",
    formula: "=SUM(B15:E15)",
    math: "Formula: Mid Sem (30) + GTU (70) + PA (25) + ESE (25) = 150 marks",
    category: "marks",
  });

  const numCos = state.numCos || 5;
  const activeCOs = COs.slice(0, numCos);
  const activeColumns = useFullPOs ? POPSO : DIPLOMA_POPSO;

  // Evaluation Plan Math
  const marks = state.evalPlan.marks;
  const weights = state.evalPlan.weights;
  const totalMarks = (marks.midSem || 0) + (marks.gtu || 0) + (marks.pa || 0) + (marks.ese || 0);
  const totalWeight = (weights.midSem || 0) + (weights.gtu || 0) + (weights.pa || 0) + (weights.ese || 0);
  const theoryMarks = (marks.midSem || 0) + (marks.gtu || 0);
  const practicalMarks = (marks.pa || 0) + (marks.ese || 0);
  const thPct = totalMarks ? Math.round((theoryMarks / totalMarks) * 100) : 67;
  const prPct = totalMarks ? Math.round((practicalMarks / totalMarks) * 100) : 33;

  // Mapping Averages
  const averages = useMemo(() => {
    return activeColumns.map((_, colIdx) => {
      let sum = 0, count = 0;
      for (let i = 0; i < numCos; i++) {
        const val = state.mapping[i]?.[colIdx];
        if (val > 0) {
          sum += val;
          count++;
        }
      }
      return count > 0 ? (sum / count).toFixed(2) : "-";
    });
  }, [state.mapping, numCos, activeColumns]);

  // Target Levels Math
  const targets = state.targetLevels;
  const theoryAvg = ((targets.theory.paTarget + targets.theory.eseTarget) / 2).toFixed(2);
  const practicalAvg = ((targets.practical.paTarget + targets.practical.eseTarget) / 2).toFixed(2);
  const courseFinalTarget = targets.finalTarget || 0.9;

  function inspectCell(cellId, title, formula, math, category = "general") {
    setSelectedCell({ cellId, title, formula, math, category });
  }

  function updateCI(key, val) {
    setState((s) => ({ ...s, courseInfo: { ...s.courseInfo, [key]: val } }));
  }
  function updateCOCode(i, val) {
    setState((s) => {
      const arr = [...(s.coCodes || DEFAULT_CO_CODES)];
      arr[i] = val;
      return { ...s, coCodes: arr };
    });
  }
  function updateCOStatement(i, val) {
    setState((s) => {
      const arr = [...s.coStatements];
      arr[i] = val;
      return { ...s, coStatements: arr };
    });
  }
  function setMappingCell(i, j, val) {
    setState((s) => {
      const m = s.mapping.map((r) => [...r]);
      m[i][j] = clamp(val, 0, 3);
      return { ...s, mapping: m };
    });
  }
  function updateMarks(k, val) {
    setState((s) => ({
      ...s,
      evalPlan: {
        ...s.evalPlan,
        marks: { ...s.evalPlan.marks, [k]: clamp(val, 0, 100) },
      },
    }));
  }
  function updateWeights(k, val) {
    setState((s) => ({
      ...s,
      evalPlan: {
        ...s.evalPlan,
        weights: { ...s.evalPlan.weights, [k]: clamp(val, 0, 100) },
      },
    }));
  }
  function updateTargetLevel(section, key, val) {
    setState((s) => ({
      ...s,
      targetLevels: {
        ...s.targetLevels,
        [section]: {
          ...s.targetLevels[section],
          [key]: key.includes("Target") ? clamp(val, 0, 3) : clamp(val, 0, 100),
        },
      },
    }));
  }

  return (
    <div className="eval-plan-wrapper">
      {/* Top Toolbar */}
      <div className="eval-toolbar hide-on-print">
        <div className="eval-toolbar-left">
          <FileSpreadsheet className="eval-toolbar-icon" size={22} />
          <div>
            <h2 className="eval-toolbar-title">Course Evaluation Plan</h2>
            <p className="eval-toolbar-sub">Official GTU &amp; NBA Course Evaluation Document</p>
          </div>
        </div>
        <div className="eval-toolbar-actions">
          <button
            className={`btn-ghost ${editMode ? "btn-active-edit" : ""}`}
            onClick={() => setEditMode(!editMode)}
            title="Toggle inline editing for tables"
          >
            <Edit3 size={14} /> {editMode ? "Done Editing" : "Quick Edit"}
          </button>
          <button
            className="btn-ghost"
            onClick={() => setUseFullPOs(!useFullPOs)}
            title="Toggle between 7 POs (Diploma) and 12 POs (Degree)"
          >
            {useFullPOs ? "Show Diploma POs (PO1–PO7)" : "Show All POs (PO1–PO12)"}
          </button>
          <button className="btn-ghost" onClick={onExportExcel} title="Export this evaluation plan directly to Excel">
            <Download size={14} /> Export Plan Excel
          </button>
          <button className="btn-primary" onClick={() => window.print()} title="Print or save as official PDF">
            <Printer size={15} /> Print Plan / PDF
          </button>
        </div>
      </div>

      {/* College & Department Configuration Card (Prominent input guide) */}
      <div className="prpa-inst-box hide-on-print">
        <div className="prpa-inst-header">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <BookOpenCheck size={18} color="#4338ca" />
            <strong style={{ color: "#1e1b4b", fontSize: 13.5 }}>Institution &amp; Department Details</strong>
          </div>
          <span className="prpa-inst-badge">
            ℹ️ Enter College / University Name and Department Name below. These appear in the document header and are synchronized across all tabs.
          </span>
        </div>
        <div className="prpa-inst-grid">
          <div className="prpa-inst-field">
            <label className="prpa-inst-label">College / University Name:</label>
            <input
              type="text"
              className="prpa-inst-input"
              value={state.courseInfo.institute || ""}
              placeholder="e.g. K.D.POLYTECHNIC,PATAN"
              onChange={(e) => updateCI("institute", e.target.value)}
            />
          </div>
          <div className="prpa-inst-field">
            <label className="prpa-inst-label">Department Name:</label>
            <input
              type="text"
              className="prpa-inst-input"
              value={state.courseInfo.department || ""}
              placeholder="e.g. COMPUTER ENGINEERING DEPARTMENT"
              onChange={(e) => updateCI("department", e.target.value)}
            />
          </div>
        </div>
      </div>

      {/* Interactive Excel Formula Bar (fx) */}
      <div className="prpa-fx-bar hide-on-print">
        <div className="prpa-fx-left">
          <div className="prpa-fx-name-box" title="Selected Excel Cell Reference">
            {selectedCell.cellId}
          </div>
          <div className="prpa-fx-icon" title="Formula (fx)">
            <em>f</em><span>x</span>
          </div>
        </div>
        <div className="prpa-fx-content">
          <div className="prpa-fx-formula" title="Exact Excel Mathematical Formula">
            <code>{selectedCell.formula}</code>
          </div>
          <div className="prpa-fx-math" title="Live Step-by-Step Mathematical Calculation Breakdown">
            <span className="prpa-fx-pill">📐 {selectedCell.title}:</span>
            <strong>{selectedCell.math}</strong>
          </div>
        </div>
      </div>

      {/* Main Official Document Sheet */}
      <div className="eval-sheet">
        {/* Header */}
        <div className="eval-doc-header">
          <h1 className="eval-inst-title">
            {editMode ? (
              <input
                className="eval-inline-input eval-title-input"
                value={state.courseInfo.institute}
                onChange={(e) => updateCI("institute", e.target.value)}
              />
            ) : (
              state.courseInfo.institute || "K.D.POLYTECHNIC,PATAN"
            )}
          </h1>
          <h2 className="eval-dept-title">
            {editMode ? (
              <input
                className="eval-inline-input eval-dept-input"
                value={state.courseInfo.department}
                onChange={(e) => updateCI("department", e.target.value)}
              />
            ) : (
              state.courseInfo.department || "COMPUTER ENGINEERING DEPARTMENT"
            )}
          </h2>
          <h3 className="eval-plan-title">COURSE EVALUATION PLAN</h3>
        </div>

        {/* 1. Course Details Table */}
        <div className="eval-table-wrap">
          <table className="eval-table eval-meta-table">
            <tbody>
              <tr>
                <td className="eval-meta-label">Course Code:</td>
                <td
                  className={`eval-meta-val eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "B4" ? "prpa-cell-active" : ""}`}
                  onClick={() => inspectCell("B4", "Course Code", `="${state.courseInfo.courseCode || "4350702"}"`, `Course Code = ${state.courseInfo.courseCode || "4350702"}`)}
                >
                  {editMode ? (
                    <input className="eval-cell-input" value={state.courseInfo.courseCode} onChange={(e) => updateCI("courseCode", e.target.value)} />
                  ) : (
                    <strong>{state.courseInfo.courseCode || "4350702"}</strong>
                  )}
                </td>
                <td className="eval-meta-label">Course Name:</td>
                <td
                  className={`eval-meta-val eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "D4" ? "prpa-cell-active" : ""}`}
                  colSpan={3}
                  onClick={() => inspectCell("D4", "Course Name", `="${state.courseInfo.courseName || "Introduction to Machine Learning"}"`, `Course Name = ${state.courseInfo.courseName || "Introduction to Machine Learning"}`)}
                >
                  {editMode ? (
                    <input className="eval-cell-input eval-input-wide" value={state.courseInfo.courseName} onChange={(e) => updateCI("courseName", e.target.value)} />
                  ) : (
                    <strong>{state.courseInfo.courseName || "Introduction to Machine Learning"}</strong>
                  )}
                </td>
              </tr>
              <tr>
                <td className="eval-meta-label">Batch:</td>
                <td
                  className={`eval-meta-val eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "B5" ? "prpa-cell-active" : ""}`}
                  onClick={() => inspectCell("B5", "Batch", `="${state.courseInfo.batch || "2021-24"}"`, `Academic Batch = ${state.courseInfo.batch || "2021-24"}`)}
                >
                  {editMode ? (
                    <input className="eval-cell-input" value={state.courseInfo.batch} onChange={(e) => updateCI("batch", e.target.value)} />
                  ) : (
                    state.courseInfo.batch || "2021-24"
                  )}
                </td>
                <td className="eval-meta-label">Term:</td>
                <td
                  className={`eval-meta-val eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "D5" ? "prpa-cell-active" : ""}`}
                  colSpan={3}
                  onClick={() => inspectCell("D5", "Term", `="${state.courseInfo.term || "231"}"`, `Term Code = ${state.courseInfo.term || "231"}`)}
                >
                  {editMode ? (
                    <input className="eval-cell-input" value={state.courseInfo.term} onChange={(e) => updateCI("term", e.target.value)} />
                  ) : (
                    state.courseInfo.term || "231"
                  )}
                </td>
              </tr>
              <tr>
                <td className="eval-meta-label">Number of Students:</td>
                <td
                  className={`eval-meta-val eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "B6" ? "prpa-cell-active" : ""}`}
                  onClick={() => inspectCell("B6", "Total Students Cohort", `=${state.courseInfo.numStudents || 137}`, `Total enrolled students cohort count = ${state.courseInfo.numStudents || 137}`)}
                >
                  {editMode ? (
                    <input type="number" className="eval-cell-input" value={state.courseInfo.numStudents} onChange={(e) => updateCI("numStudents", Number(e.target.value))} />
                  ) : (
                    state.courseInfo.numStudents || 137
                  )}
                </td>
                <td className="eval-meta-label">Semester :</td>
                <td
                  className={`eval-meta-val eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "D6" ? "prpa-cell-active" : ""}`}
                  colSpan={3}
                  onClick={() => inspectCell("D6", "Semester", `="${state.courseInfo.semester || "5"}"`, `Semester = ${state.courseInfo.semester || "5"}`)}
                >
                  {editMode ? (
                    <input className="eval-cell-input" value={state.courseInfo.semester} onChange={(e) => updateCI("semester", e.target.value)} />
                  ) : (
                    state.courseInfo.semester || "5"
                  )}
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* 2. CO's Section */}
        <div className="eval-section-block">
          <div className="eval-section-badge">CO's</div>
          <div className="eval-table-wrap">
            <table className="eval-table eval-co-table">
              <tbody>
                {activeCOs.map((co, i) => (
                  <tr key={co}>
                    <td className="eval-cell-peach eval-co-id">{co}</td>
                    <td
                      className={`eval-cell-green eval-co-subcode prpa-clickable-cell ${selectedCell.cellId === `B${i + 8}` ? "prpa-cell-active" : ""}`}
                      onClick={() => inspectCell(`B${i + 8}`, `${co} Code`, `="${state.coCodes?.[i] || DEFAULT_CO_CODES[i]}"`, `Code = ${state.coCodes?.[i] || DEFAULT_CO_CODES[i]}`)}
                    >
                      {editMode ? (
                        <input className="eval-cell-input" value={state.coCodes?.[i] || DEFAULT_CO_CODES[i]} onChange={(e) => updateCOCode(i, e.target.value)} />
                      ) : (
                        state.coCodes?.[i] || DEFAULT_CO_CODES[i]
                      )}
                    </td>
                    <td
                      className={`eval-cell-green eval-co-statement prpa-clickable-cell ${selectedCell.cellId === `C${i + 8}` ? "prpa-cell-active" : ""}`}
                      onClick={() => inspectCell(`C${i + 8}`, `${co} Statement`, `="${state.coStatements[i] || ""}"`, `Course outcome statement for ${co}`)}
                    >
                      {editMode ? (
                        <input className="eval-cell-input eval-input-wide" value={state.coStatements[i]} onChange={(e) => updateCOStatement(i, e.target.value)} />
                      ) : (
                        state.coStatements[i] || "–"
                      )}
                    </td>
                  </tr>
                ))}
                <tr>
                  <td className="eval-cell-peach eval-co-id">No of Cos</td>
                  <td className="eval-cell-green eval-co-subcode">
                    {editMode ? (
                      <select className="eval-cell-input" value={numCos} onChange={(e) => setState((s) => ({ ...s, numCos: Number(e.target.value) }))}>
                        <option value={4}>4</option>
                        <option value={5}>5</option>
                        <option value={6}>6</option>
                      </select>
                    ) : (
                      <strong>{numCos}</strong>
                    )}
                  </td>
                  <td className="eval-cell-green"></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* 3. CO-PO-PSO Mapping Section */}
        <div className="eval-section-block">
          <div className="eval-section-badge">CO-PO-PSO Mapping</div>
          <div className="eval-table-wrap">
            <table className="eval-table eval-mapping-table">
              <thead>
                <tr>
                  <th className="eval-cell-peach" style={{ width: 80 }}></th>
                  {activeColumns.map((col) => (
                    <th key={col} className="eval-cell-peach">{col}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {activeCOs.map((co, i) => (
                  <tr key={co}>
                    <td className="eval-cell-peach eval-co-id">{co}</td>
                    {activeColumns.map((col, j) => {
                      const val = state.mapping[i]?.[j] || 0;
                      const colLetter = String.fromCharCode(66 + j);
                      const rowNum = i + 15;
                      const cellRef = `${colLetter}${rowNum}`;
                      return (
                        <td
                          key={col}
                          className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === cellRef ? "prpa-cell-active" : ""}`}
                          onClick={() => inspectCell(cellRef, `${co} mapping to ${col}`, `=${val}`, `Correlation strength of ${co} with ${col} (0: None, 1: Low, 2: Moderate, 3: High) = ${val || "–"}`)}
                        >
                          {editMode ? (
                            <select
                              className="eval-select-cell"
                              value={val}
                              onChange={(e) => setMappingCell(i, j, Number(e.target.value))}
                            >
                              <option value={0}>–</option>
                              <option value={1}>1</option>
                              <option value={2}>2</option>
                              <option value={3}>3</option>
                            </select>
                          ) : (
                            val === 0 ? "–" : val
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
                <tr className="eval-avg-row">
                  <td className="eval-cell-peach eval-co-id">AVERAGE</td>
                  {averages.map((avg, j) => {
                    const colLetter = String.fromCharCode(66 + j);
                    const formula = `=AVERAGE(${colLetter}15:${colLetter}${14 + numCos})`;
                    return (
                      <td
                        key={j}
                        className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === colLetter + "20" ? "prpa-cell-active" : ""}`}
                        onClick={() =>
                          inspectCell(
                            `${colLetter}20`,
                            `${activeColumns[j]} Average Mapping Correlation`,
                            formula,
                            `Formula: Average of non-zero mapping values for ${activeColumns[j]} = ${avg === "-" ? "None" : avg}`
                          )
                        }
                      >
                        <strong>{avg === "-" ? "" : avg}</strong>
                      </td>
                    );
                  })}
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* 4. Evaluation Plan Section */}
        <div className="eval-section-block">
          <div className="eval-section-badge">Evaluation Plan</div>
          <div className="eval-table-wrap">
            <table className="eval-table eval-plan-table">
              <thead>
                <tr>
                  <th rowSpan={2} className="eval-cell-peach" style={{ width: 140 }}>Assesment Tool</th>
                  <th colSpan={2} className="eval-cell-peach">Theory Marks</th>
                  <th colSpan={2} className="eval-cell-peach">Practical Marks</th>
                  <th rowSpan={2} className="eval-cell-peach" style={{ width: 120 }}>Total Marks</th>
                  <th rowSpan={2} className="eval-no-border hide-on-print" style={{ width: 160 }}></th>
                </tr>
                <tr>
                  <th className="eval-cell-peach">Mid Sem</th>
                  <th className="eval-cell-peach">GTU</th>
                  <th className="eval-cell-peach">PA</th>
                  <th className="eval-cell-peach">ESE</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td className="eval-cell-blue-header">Marks</td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "B24" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("B24", "Theory Mid Sem Exam Marks", `=${marks.midSem}`, `Theory Mid Semester marks allocated = ${marks.midSem}`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={marks.midSem} onChange={(e) => updateMarks("midSem", Number(e.target.value))} />
                    ) : marks.midSem}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "C24" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("C24", "Theory GTU University Exam Marks", `=${marks.gtu}`, `Theory End Semester University marks allocated = ${marks.gtu}`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={marks.gtu} onChange={(e) => updateMarks("gtu", Number(e.target.value))} />
                    ) : marks.gtu}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "D24" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("D24", "Practical Progressive Assessment (PA) Marks", `=${marks.pa}`, `Practical Progressive Assessment marks allocated = ${marks.pa}`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={marks.pa} onChange={(e) => updateMarks("pa", Number(e.target.value))} />
                    ) : marks.pa}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "E24" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("E24", "Practical End Semester Exam (ESE) Marks", `=${marks.ese}`, `Practical End Semester Exam marks allocated = ${marks.ese}`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={marks.ese} onChange={(e) => updateMarks("ese", Number(e.target.value))} />
                    ) : marks.ese}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "F24" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "F24",
                        "Course Total Evaluation Marks",
                        "=SUM(B24:E24)",
                        `Formula: Mid Sem (${marks.midSem}) + GTU (${marks.gtu}) + PA (${marks.pa}) + ESE (${marks.ese}) = ${totalMarks} marks`
                      )
                    }
                  >
                    <strong>{totalMarks}</strong>
                  </td>
                  <td className="eval-no-border"></td>
                </tr>
                <tr>
                  <td className="eval-cell-blue-header">Weightage(%)</td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "B25" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("B25", "Theory Mid Sem Weight %", `=${weights.midSem}%`, `Theory Mid Sem weightage = ${weights.midSem}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={weights.midSem} onChange={(e) => updateWeights("midSem", Number(e.target.value))} />
                    ) : weights.midSem}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "C25" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("C25", "Theory GTU Weight %", `=${weights.gtu}%`, `Theory GTU University exam weightage = ${weights.gtu}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={weights.gtu} onChange={(e) => updateWeights("gtu", Number(e.target.value))} />
                    ) : weights.gtu}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "D25" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("D25", "Practical PA Weight %", `=${weights.pa}%`, `Practical PA weightage = ${weights.pa}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={weights.pa} onChange={(e) => updateWeights("pa", Number(e.target.value))} />
                    ) : weights.pa}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "E25" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("E25", "Practical ESE Weight %", `=${weights.ese}%`, `Practical ESE weightage = ${weights.ese}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={weights.ese} onChange={(e) => updateWeights("ese", Number(e.target.value))} />
                    ) : weights.ese}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "F25" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "F25",
                        "Total Evaluation Plan Weightage",
                        "=SUM(B25:E25)",
                        `Formula: Mid Sem (${weights.midSem}%) + GTU (${weights.gtu}%) + PA (${weights.pa}%) + ESE (${weights.ese}%) = ${totalWeight}%`
                      )
                    }
                  >
                    <strong>{totalWeight}</strong>
                  </td>
                  <td className="eval-note-red">Refer GTU syllabus</td>
                </tr>
                {/* CO breakdown row */}
                <tr>
                  <td rowSpan={2} className="eval-cell-lavender" style={{ verticalAlign: "middle" }}>
                    <strong>CO1 to CO{numCos}</strong>
                  </td>
                  <td
                    colSpan={2}
                    rowSpan={2}
                    className={`eval-cell-grey prpa-clickable-cell ${selectedCell.cellId === "B26" ? "prpa-cell-active" : ""}`}
                    style={{ verticalAlign: "middle", fontSize: 14 }}
                    onClick={() =>
                      inspectCell(
                        "B26",
                        "Total Theory Marks",
                        "=B24+C24",
                        `Formula: Mid Sem (${marks.midSem}) + GTU (${marks.gtu}) = ${theoryMarks} marks`
                      )
                    }
                  >
                    <strong>{theoryMarks}</strong>
                  </td>
                  <td
                    colSpan={2}
                    rowSpan={2}
                    className={`eval-cell-grey prpa-clickable-cell ${selectedCell.cellId === "D26" ? "prpa-cell-active" : ""}`}
                    style={{ verticalAlign: "middle", fontSize: 14 }}
                    onClick={() =>
                      inspectCell(
                        "D26",
                        "Total Practical Marks",
                        "=D24+E24",
                        `Formula: Practical PA (${marks.pa}) + ESE (${marks.ese}) = ${practicalMarks} marks`
                      )
                    }
                  >
                    <strong>{practicalMarks}</strong>
                  </td>
                  <td className="eval-cell-grey" style={{ padding: 0 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", borderBottom: "1px solid #777", fontWeight: 700, padding: "3px 0" }}>
                      <span>TH</span>
                      <span style={{ borderLeft: "1px solid #777" }}>PR</span>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", padding: "4px 0" }}>
                      <span
                        className={`prpa-clickable-cell ${selectedCell.cellId === "F26_TH" ? "prpa-cell-active" : ""}`}
                        onClick={() =>
                          inspectCell(
                            "F26_TH",
                            "Theory Proportion Weightage %",
                            "=ROUND((B26/F24)*100,0)",
                            `Formula: [Theory Marks (${theoryMarks})] ÷ [Total Marks (${totalMarks})] × 100 = ${thPct}%`
                          )
                        }
                      >
                        {thPct}%
                      </span>
                      <span
                        style={{ borderLeft: "1px solid #777" }}
                        className={`prpa-clickable-cell ${selectedCell.cellId === "F26_PR" ? "prpa-cell-active" : ""}`}
                        onClick={() =>
                          inspectCell(
                            "F26_PR",
                            "Practical Proportion Weightage %",
                            "=ROUND((D26/F24)*100,0)",
                            `Formula: [Practical Marks (${practicalMarks})] ÷ [Total Marks (${totalMarks})] × 100 = ${prPct}%`
                          )
                        }
                      >
                        {prPct}%
                      </span>
                    </div>
                  </td>
                  <td className="eval-no-border"></td>
                </tr>
                <tr>
                  <td
                    className={`eval-cell-grey prpa-clickable-cell ${selectedCell.cellId === "F27" ? "prpa-cell-active" : ""}`}
                    style={{ padding: "6px 0" }}
                    onClick={() =>
                      inspectCell(
                        "F27",
                        "Total Combined Proportion",
                        "=TH% + PR%",
                        `Formula: Theory (${thPct}%) + Practical (${prPct}%) = 100%`
                      )
                    }
                  >
                    <strong>100%</strong>
                  </td>
                  <td className="eval-no-border"></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* 5. Defined Target Levels Section */}
        <div className="eval-section-block">
          <div className="eval-section-badge">Defined Target Leves for All CO's</div>
          <div className="eval-table-wrap">
            <table className="eval-table eval-targets-table">
              <thead>
                <tr>
                  <th rowSpan={3} className="eval-cell-peach" style={{ width: 100 }}>TH/PR</th>
                  <th rowSpan={3} className="eval-cell-peach" style={{ width: 130 }}>Course Outcome</th>
                  <th colSpan={3} className="eval-cell-peach">PROGRESSIVE ASSESSMENT</th>
                  <th colSpan={3} className="eval-cell-peach">ESE EXAMINATION</th>
                  <th rowSpan={3} className="eval-cell-peach" style={{ width: 110 }}>Average of Target</th>
                </tr>
                <tr>
                  <th rowSpan={2} className="eval-cell-peach">Target Level</th>
                  <th colSpan={2} className="eval-cell-peach">(a)% Students Scoring More than or Equal to (b)% Marks</th>
                  <th rowSpan={2} className="eval-cell-peach">Target Level</th>
                  <th colSpan={2} className="eval-cell-peach">(c)% Students Scoring More than or Equal to (d)% Marks</th>
                </tr>
                <tr>
                  <th className="eval-cell-peach">(a) students</th>
                  <th className="eval-cell-peach">(b) marks</th>
                  <th className="eval-cell-peach">(c) students</th>
                  <th className="eval-cell-peach">(d) marks</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td className="eval-cell-white" style={{ fontWeight: 800 }}>THEORY</td>
                  <td className="eval-cell-amber"><strong>CO1 to CO{numCos}</strong></td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "C31" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("C31", "Theory Progressive Assessment Target Level", `=${targets.theory.paTarget.toFixed(2)}`, `Target level defined for Theory PA = ${targets.theory.paTarget.toFixed(2)}`)}
                  >
                    {editMode ? (
                      <input type="number" step="0.05" className="eval-cell-input" value={targets.theory.paTarget} onChange={(e) => updateTargetLevel("theory", "paTarget", Number(e.target.value))} />
                    ) : targets.theory.paTarget.toFixed(2)}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "D31" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("D31", "Theory PA Target Students % (a)", `=${targets.theory.paStudents}%`, `Target minimum % of students scoring threshold marks = ${targets.theory.paStudents}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={targets.theory.paStudents} onChange={(e) => updateTargetLevel("theory", "paStudents", Number(e.target.value))} />
                    ) : targets.theory.paStudents}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "E31" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("E31", "Theory PA Target Marks Threshold % (b)", `=${targets.theory.paMarks}%`, `Target threshold marks % for student attainment = ${targets.theory.paMarks}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={targets.theory.paMarks} onChange={(e) => updateTargetLevel("theory", "paMarks", Number(e.target.value))} />
                    ) : targets.theory.paMarks}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "F31" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("F31", "Theory ESE Examination Target Level", `=${targets.theory.eseTarget.toFixed(2)}`, `Target level defined for Theory ESE = ${targets.theory.eseTarget.toFixed(2)}`)}
                  >
                    {editMode ? (
                      <input type="number" step="0.05" className="eval-cell-input" value={targets.theory.eseTarget} onChange={(e) => updateTargetLevel("theory", "eseTarget", Number(e.target.value))} />
                    ) : targets.theory.eseTarget.toFixed(2)}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "G31" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("G31", "Theory ESE Target Students % (c)", `=${targets.theory.eseStudents}%`, `Target minimum % of students scoring threshold marks = ${targets.theory.eseStudents}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={targets.theory.eseStudents} onChange={(e) => updateTargetLevel("theory", "eseStudents", Number(e.target.value))} />
                    ) : targets.theory.eseStudents}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "H31" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("H31", "Theory ESE Target Marks Threshold % (d)", `=${targets.theory.eseMarks}%`, `Target threshold marks % for student attainment = ${targets.theory.eseMarks}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={targets.theory.eseMarks} onChange={(e) => updateTargetLevel("theory", "eseMarks", Number(e.target.value))} />
                    ) : targets.theory.eseMarks}
                  </td>
                  <td
                    className={`eval-cell-cyan prpa-clickable-cell ${selectedCell.cellId === "I31" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "I31",
                        "Theory Average of Target",
                        "=AVERAGE(C31,F31)",
                        `Formula: [PA Target (${targets.theory.paTarget.toFixed(2)}) + ESE Target (${targets.theory.eseTarget.toFixed(2)})] ÷ 2 = ${theoryAvg}`
                      )
                    }
                  >
                    <strong>{theoryAvg}</strong>
                  </td>
                </tr>
                <tr>
                  <td className="eval-cell-white" style={{ fontWeight: 800 }}>PRACTICAL</td>
                  <td className="eval-cell-amber"><strong>CO1 to CO{numCos}</strong></td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "C32" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("C32", "Practical Progressive Assessment Target Level", `=${targets.practical.paTarget.toFixed(2)}`, `Target level defined for Practical PA = ${targets.practical.paTarget.toFixed(2)}`)}
                  >
                    {editMode ? (
                      <input type="number" step="0.05" className="eval-cell-input" value={targets.practical.paTarget} onChange={(e) => updateTargetLevel("practical", "paTarget", Number(e.target.value))} />
                    ) : targets.practical.paTarget.toFixed(2)}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "D32" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("D32", "Practical PA Target Students % (a)", `=${targets.practical.paStudents}%`, `Target minimum % of students scoring threshold marks = ${targets.practical.paStudents}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={targets.practical.paStudents} onChange={(e) => updateTargetLevel("practical", "paStudents", Number(e.target.value))} />
                    ) : targets.practical.paStudents}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "E32" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("E32", "Practical PA Target Marks Threshold % (b)", `=${targets.practical.paMarks}%`, `Target threshold marks % for student attainment = ${targets.practical.paMarks}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={targets.practical.paMarks} onChange={(e) => updateTargetLevel("practical", "paMarks", Number(e.target.value))} />
                    ) : targets.practical.paMarks}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "F32" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("F32", "Practical ESE Examination Target Level", `=${targets.practical.eseTarget.toFixed(2)}`, `Target level defined for Practical ESE = ${targets.practical.eseTarget.toFixed(2)}`)}
                  >
                    {editMode ? (
                      <input type="number" step="0.05" className="eval-cell-input" value={targets.practical.eseTarget} onChange={(e) => updateTargetLevel("practical", "eseTarget", Number(e.target.value))} />
                    ) : targets.practical.eseTarget.toFixed(2)}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "G32" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("G32", "Practical ESE Target Students % (c)", `=${targets.practical.eseStudents}%`, `Target minimum % of students scoring threshold marks = ${targets.practical.eseStudents}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={targets.practical.eseStudents} onChange={(e) => updateTargetLevel("practical", "eseStudents", Number(e.target.value))} />
                    ) : targets.practical.eseStudents}
                  </td>
                  <td
                    className={`eval-cell-green prpa-clickable-cell ${selectedCell.cellId === "H32" ? "prpa-cell-active" : ""}`}
                    onClick={() => inspectCell("H32", "Practical ESE Target Marks Threshold % (d)", `=${targets.practical.eseMarks}%`, `Target threshold marks % for student attainment = ${targets.practical.eseMarks}%`)}
                  >
                    {editMode ? (
                      <input type="number" className="eval-cell-input" value={targets.practical.eseMarks} onChange={(e) => updateTargetLevel("practical", "eseMarks", Number(e.target.value))} />
                    ) : targets.practical.eseMarks}
                  </td>
                  <td
                    className={`eval-cell-cyan prpa-clickable-cell ${selectedCell.cellId === "I32" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "I32",
                        "Practical Average of Target",
                        "=AVERAGE(C32,F32)",
                        `Formula: [PA Target (${targets.practical.paTarget.toFixed(2)}) + ESE Target (${targets.practical.eseTarget.toFixed(2)})] ÷ 2 = ${practicalAvg}`
                      )
                    }
                  >
                    <strong>{practicalAvg}</strong>
                  </td>
                </tr>
                {/* Final Target merged footer */}
                <tr>
                  <td
                    colSpan={9}
                    className={`eval-cell-grey eval-final-target-row prpa-clickable-cell ${selectedCell.cellId === "A33" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "A33",
                        "Final Target for Whole Course",
                        `=${courseFinalTarget}`,
                        `Overall defined target attainment level for the entire course = ${courseFinalTarget}`
                      )
                    }
                  >
                    Final Target for Whole Course : {courseFinalTarget}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   TARGET SETTING / MATHEMATICAL TARGET FORMULATION TAB (PIC 1)
============================================================================ */
function TargetSettingTab({ state, setState, showToast, finalDirectAttain }) {
  const ts = state.targetSetting || DEFAULT_TARGET_SETTING;
  const years = ts.years || DEFAULT_TARGET_SETTING.years;
  const gradeCounts = ts.gradeCounts || DEFAULT_TARGET_SETTING.gradeCounts;

  // Selected cell for interactive Excel formula bar & live equation inspector
  const [selectedCell, setSelectedCell] = useState({
    cellId: "M14",
    title: "Target Marks Percentage Threshold",
    formula: "=ROUNDUP(M6,0)",
    math: "Formula: ROUNDUP of 3-year historical average result (60.26%) = 61%",
    category: "target",
  });

  // 1. Math for each year: Total students, total marks, % result
  const yearStats = useMemo(() => {
    return years.map((_, y) => {
      let totalStudents = 0;
      let totalMarks = 0;
      GTU_GRADES.forEach((g, k) => {
        const count = Number(gradeCounts[y]?.[k]) || 0;
        totalStudents += count;
        totalMarks += count * g.avg;
      });
      const pctResult = totalStudents > 0 ? (totalMarks / (totalStudents * 100)) * 100 : 0;
      return { totalStudents, totalMarks, pctResult };
    });
  }, [years, gradeCounts]);

  // Non-zero years average
  const validYears = yearStats.filter((ys) => ys.totalStudents > 0);
  const avgResultPct = validYears.length
    ? validYears.reduce((a, b) => a + b.pctResult, 0) / validYears.length
    : 60.26;

  // Target Marks Percentage (rounded up from average result %)
  const targetMarksPct = ts.targetMarksPct != null ? ts.targetMarksPct : Math.ceil(avgResultPct);

  // 2. Math for students having marks >= targetMarksPct (grades AA, AB, BB)
  const studentsAboveThreshold = useMemo(() => {
    return years.map((_, y) => {
      let count = 0;
      GTU_GRADES.forEach((g, k) => {
        if (g.avg >= targetMarksPct || g.isAboveThreshold) {
          count += Number(gradeCounts[y]?.[k]) || 0;
        }
      });
      const tot = yearStats[y].totalStudents;
      const pct = tot > 0 ? (count / tot) * 100 : 0;
      return { count, pct };
    });
  }, [years, gradeCounts, yearStats, targetMarksPct]);

  const validThresholdYears = studentsAboveThreshold.filter((_, y) => yearStats[y].totalStudents > 0);
  const avgAbovePct = validThresholdYears.length
    ? validThresholdYears.reduce((a, b) => a + b.pct, 0) / validThresholdYears.length
    : 47.0;

  const targetStudentsPct = ts.targetStudentsPct != null ? ts.targetStudentsPct : Math.round(avgAbovePct + 1);
  const targetLevel = ts.targetLevel != null ? ts.targetLevel : 0.90;

  function inspectCell(cellId, title, formula, math, category = "general") {
    setSelectedCell({ cellId, title, formula, math, category });
  }

  // Calculate live achieved outcome from current student assessments
  const liveAchieved = useMemo(() => {
    // Average attained direct level across active COs
    const numCos = state.numCos || 5;
    const directVals = (finalDirectAttain || []).slice(0, numCos);
    const avgDirect = directVals.length ? (directVals.reduce((a, b) => a + b, 0) / directVals.length) : 1.52;
    // Calculate student percentage scoring >= targetMarksPct across all internal/endsem exams
    const allAssessments = [state.internal1, state.internal2, state.assignment, state.endsem];
    let totalAttempts = 0, totalAttained = 0;
    allAssessments.forEach((ass) => {
      if (ass?.students?.length) {
        ass.students.forEach((st) => {
          st.marks.forEach((m, idx) => {
            const maxM = ass.maxMarks[idx] || 10;
            if (maxM > 0) {
              totalAttempts++;
              if (m >= maxM * (targetMarksPct / 100)) {
                totalAttained++;
              }
            }
          });
        });
      }
    });
    const achievedPct = totalAttempts > 0 ? Math.round((totalAttained / totalAttempts) * 100) : 55;
    return { level: avgDirect > 0 ? avgDirect.toFixed(2) : "1.52", pct: achievedPct };
  }, [state, finalDirectAttain, targetMarksPct]);

  // Helper to synchronize state: when Pic 1 mathematical formula recalculates, Pic 2 in Evaluation Plan updates automatically
  function propagateTargetState(newGradeCounts, explicitTargetMarks, explicitTargetStudents, explicitTargetLevel, newYears, newNbaCode) {
    setState((s) => {
      const curTs = s.targetSetting || DEFAULT_TARGET_SETTING;
      const gc = newGradeCounts || curTs.gradeCounts;
      const yrs = newYears || curTs.years;

      let yStats = yrs.map((_, y) => {
        let totSt = 0, totM = 0;
        GTU_GRADES.forEach((g, k) => {
          const count = Number(gc[y]?.[k]) || 0;
          totSt += count;
          totM += count * g.avg;
        });
        const pctRes = totSt > 0 ? (totM / (totSt * 100)) * 100 : 0;
        return { totSt, totM, pctRes };
      });
      const vYears = yStats.filter((ys) => ys.totSt > 0);
      const avgRes = vYears.length ? vYears.reduce((a, b) => a + b.pctRes, 0) / vYears.length : 60.26;
      const calcMarksPct = explicitTargetMarks != null ? explicitTargetMarks : Math.ceil(avgRes);

      let aboveCountList = yrs.map((_, y) => {
        let count = 0;
        GTU_GRADES.forEach((g, k) => {
          if (g.avg >= calcMarksPct || g.isAboveThreshold) count += Number(gc[y]?.[k]) || 0;
        });
        const tot = yStats[y].totSt;
        return tot > 0 ? (count / tot) * 100 : 0;
      });
      const vAbove = aboveCountList.filter((_, y) => yStats[y].totSt > 0);
      const avgAbove = vAbove.length ? vAbove.reduce((a, b) => a + b, 0) / vAbove.length : 47.0;
      const calcStudentsPct = explicitTargetStudents != null ? explicitTargetStudents : Math.round(avgAbove + 1);
      const calcLvl = explicitTargetLevel != null ? explicitTargetLevel : (curTs.targetLevel || 0.90);

      const targetLevels = {
        theory: {
          paTarget: calcLvl,
          paStudents: calcStudentsPct,
          paMarks: calcMarksPct,
          eseTarget: calcLvl,
          eseStudents: calcStudentsPct,
          eseMarks: calcMarksPct,
        },
        practical: {
          paTarget: calcLvl,
          paStudents: calcStudentsPct,
          paMarks: calcMarksPct,
          eseTarget: calcLvl,
          eseStudents: calcStudentsPct,
          eseMarks: calcMarksPct,
        },
        finalTarget: calcLvl,
      };

      const updatedTs = {
        ...curTs,
        gradeCounts: gc,
        years: yrs,
        nbaSubjectCode: newNbaCode || curTs.nbaSubjectCode || "C303_N",
        targetMarksPct: calcMarksPct,
        targetStudentsPct: calcStudentsPct,
        targetLevel: calcLvl,
      };

      return {
        ...s,
        targetSetting: updatedTs,
        targetLevels,
        targets: {
          ...s.targets,
          targetPctCO: COs.map(() => calcMarksPct / 100),
          coTargetLevel: calcLvl,
        },
      };
    });
  }

  function setYearLabel(y, val) {
    const arr = [...(ts.years || DEFAULT_TARGET_SETTING.years)];
    arr[y] = val;
    propagateTargetState(null, null, null, null, arr, null);
  }

  function setGradeCount(y, k, val) {
    const gc = (ts.gradeCounts || DEFAULT_TARGET_SETTING.gradeCounts).map((row) => [...row]);
    gc[y][k] = clamp(val, 0, 10000);
    propagateTargetState(gc, null, null, null, null, null);
  }

  function setTargetField(key, val) {
    if (key === "targetMarksPct") {
      propagateTargetState(null, val, null, null, null, null);
    } else if (key === "targetStudentsPct") {
      propagateTargetState(null, null, val, null, null, null);
    } else if (key === "targetLevel") {
      propagateTargetState(null, null, null, val, null, null);
    } else if (key === "nbaSubjectCode") {
      propagateTargetState(null, null, null, null, null, val);
    }
  }

  function handleApplyTargets() {
    propagateTargetState(gradeCounts, targetMarksPct, targetStudentsPct, targetLevel, years, ts.nbaSubjectCode);
    showToast("Mathematical targets applied directly to Evaluation Plan & calculators!");
  }

  return (
    <div className="eval-plan-wrapper">
      {/* Top Toolbar */}
      <div className="eval-toolbar hide-on-print">
        <div className="eval-toolbar-left">
          <CalcIcon className="eval-toolbar-icon" size={22} />
          <div>
            <h2 className="eval-toolbar-title">CO Target Setting &amp; Mathematical Formulation (Pic 1)</h2>
            <p className="eval-toolbar-sub">Calculates target attainment levels based on historical GTU university results. Dynamically updates Course Evaluation Plan (Pic 2).</p>
          </div>
        </div>
        <div className="eval-toolbar-actions">
          <button className="btn-primary" onClick={handleApplyTargets}>
            <Check size={15} /> Recalculate &amp; Sync Plan
          </button>
          <button className="btn-ghost" onClick={() => window.print()}>
            <Printer size={15} /> Print Target Sheet
          </button>
        </div>
      </div>

      {/* Interactive Excel Formula Bar (fx) */}
      <div className="prpa-fx-bar hide-on-print">
        <div className="prpa-fx-left">
          <div className="prpa-fx-name-box" title="Selected Excel Cell Reference">
            {selectedCell.cellId}
          </div>
          <div className="prpa-fx-icon" title="Formula (fx)">
            <em>f</em><span>x</span>
          </div>
        </div>
        <div className="prpa-fx-content">
          <div className="prpa-fx-formula" title="Exact Excel Mathematical Formula">
            <code>{selectedCell.formula}</code>
          </div>
          <div className="prpa-fx-math" title="Live Step-by-Step Mathematical Calculation Breakdown">
            <span className="prpa-fx-pill">📐 {selectedCell.title}:</span>
            <strong>{selectedCell.math}</strong>
          </div>
        </div>
      </div>

      {/* Main Target Setting Sheet Matching Picture 1 */}
      <div className="eval-sheet">
        {/* Header Block with Yellow Highlights */}
        <div className="ts-header-grid">
          <div className="ts-header-item">
            <span className="ts-header-label">Term :</span>
            <input
              className="ts-header-val-yellow"
              value={state.courseInfo.term || "231"}
              onChange={(e) => setState((s) => ({ ...s, courseInfo: { ...s.courseInfo, term: e.target.value } }))}
            />
          </div>
          <div className="ts-header-item">
            <span className="ts-header-label">Subject Name :</span>
            <input
              className="ts-header-val-yellow ts-val-wide"
              value={state.courseInfo.courseName || "Introduction to Machine Learning"}
              onChange={(e) => setState((s) => ({ ...s, courseInfo: { ...s.courseInfo, courseName: e.target.value } }))}
            />
          </div>
          <div className="ts-header-item">
            <span className="ts-header-label">Sem :</span>
            <input
              className="ts-header-val-yellow"
              style={{ width: 45 }}
              value={state.courseInfo.semester || "5"}
              onChange={(e) => setState((s) => ({ ...s, courseInfo: { ...s.courseInfo, semester: e.target.value } }))}
            />
          </div>
          <div className="ts-header-item">
            <span className="ts-header-label">GTU Subject Code :</span>
            <input
              className="ts-header-val-yellow"
              style={{ width: 100 }}
              value={state.courseInfo.courseCode || "4350702"}
              onChange={(e) => setState((s) => ({ ...s, courseInfo: { ...s.courseInfo, courseCode: e.target.value } }))}
            />
          </div>
          <div className="ts-header-item">
            <span className="ts-header-label">NBA Subject Code :</span>
            <input
              className="ts-header-val"
              style={{ width: 90 }}
              value={ts.nbaSubjectCode || "C303_N"}
              onChange={(e) => setTargetField("nbaSubjectCode", e.target.value)}
            />
          </div>
        </div>

        {/* 1. Grade Distribution & Historical Result Matrix (Table 1) */}
        <div className="eval-section-block">
          <div className="eval-table-wrap">
            <table className="eval-table ts-table">
              <thead>
                <tr>
                  <th rowSpan={2} style={{ width: 45 }}>Sr No</th>
                  <th rowSpan={2} style={{ width: 120 }}>Grade as per GTU</th>
                  <th rowSpan={2} style={{ width: 100 }}>Avg Marks as per range Grade</th>
                  <th colSpan={3}>No. of Student attain Grade</th>
                  <th colSpan={3}>Total Marks</th>
                  <th colSpan={3}>% of Result</th>
                  <th rowSpan={2} style={{ width: 85 }}>Average of %</th>
                </tr>
                <tr>
                  {years.map((y, idx) => (
                    <th key={idx}>
                      <input className="ts-year-input" value={y} onChange={(e) => setYearLabel(idx, e.target.value)} />
                    </th>
                  ))}
                  {years.map((y, idx) => (
                    <th key={idx}>{y}</th>
                  ))}
                  {years.map((y, idx) => (
                    <th key={idx}>{y}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {GTU_GRADES.map((g, k) => {
                  const rowExcel = k + 6;
                  return (
                    <tr key={g.grade}>
                      <td>{k + 1}</td>
                      <td style={{ textAlign: "left", paddingLeft: 8 }}>{g.grade}</td>
                      <td
                        className={`prpa-clickable-cell ${selectedCell.cellId === "C" + rowExcel ? "prpa-cell-active" : ""}`}
                        onClick={() => inspectCell("C" + rowExcel, `${g.grade} Grade Average Marks`, `=${g.avg}`, `Midpoint average marks for GTU grade ${g.grade} = ${g.avg}`)}
                      >
                        <strong>{g.avg}</strong>
                      </td>
                      {years.map((_, y) => {
                        const colCountLetter = String.fromCharCode(68 + y);
                        const cellAddr = `${colCountLetter}${rowExcel}`;
                        return (
                          <td key={y} className={selectedCell.cellId === cellAddr ? "prpa-cell-active" : ""}>
                            <input
                              type="number"
                              className="eval-cell-input"
                              value={gradeCounts[y]?.[k] ?? 0}
                              onChange={(e) => setGradeCount(y, k, Number(e.target.value))}
                              onFocus={() =>
                                inspectCell(
                                  cellAddr,
                                  `${years[y]} ${g.grade} Student Count`,
                                  `=${gradeCounts[y]?.[k] ?? 0}`,
                                  `Number of students in academic year ${years[y]} who scored GTU grade ${g.grade} = ${gradeCounts[y]?.[k] ?? 0}`
                                )
                              }
                            />
                          </td>
                        );
                      })}
                      {years.map((_, y) => {
                        const count = Number(gradeCounts[y]?.[k]) || 0;
                        const totM = count * g.avg;
                        const colTotLetter = String.fromCharCode(71 + y);
                        const colCountLetter = String.fromCharCode(68 + y);
                        const cellAddr = `${colTotLetter}${rowExcel}`;
                        return (
                          <td
                            key={y}
                            className={`prpa-clickable-cell ${selectedCell.cellId === cellAddr ? "prpa-cell-active" : ""}`}
                            onClick={() =>
                              inspectCell(
                                cellAddr,
                                `${years[y]} ${g.grade} Total Marks`,
                                `=${colCountLetter}${rowExcel}*C${rowExcel}`,
                                `Formula: Students (${count}) × Grade Avg (${g.avg}) = ${totM} marks`
                              )
                            }
                          >
                            {totM}
                          </td>
                        );
                      })}
                      {/* % of result merged column on middle rows */}
                      {k === 0 && (
                        <>
                          <td
                            rowSpan={8}
                            colSpan={3}
                            style={{ verticalAlign: "middle", fontSize: 14, fontWeight: 700 }}
                            className={`prpa-clickable-cell ${selectedCell.cellId === "J6" ? "prpa-cell-active" : ""}`}
                            onClick={() =>
                              inspectCell(
                                "J6",
                                `${years[0]} % of Result`,
                                "=(G14/(D14*100))*100",
                                `Formula: Total Marks (${yearStats[0].totalMarks}) ÷ [Students (${yearStats[0].totalStudents}) × 100] × 100 = ${yearStats[0].pctResult.toFixed(2)}%`
                              )
                            }
                          >
                            {yearStats[0].pctResult > 0 ? yearStats[0].pctResult.toFixed(2) : "–"}
                          </td>
                          <td
                            rowSpan={8}
                            style={{ verticalAlign: "middle", fontSize: 14, fontWeight: 700 }}
                            className={`prpa-clickable-cell ${selectedCell.cellId === "M6" ? "prpa-cell-active" : ""}`}
                            onClick={() =>
                              inspectCell(
                                "M6",
                                "Average of Result % Across Historical Years",
                                "=AVERAGE(J6:L6)",
                                `Formula: Average of valid historical results = ${avgResultPct.toFixed(2)}%`
                              )
                            }
                          >
                            {avgResultPct.toFixed(2)}
                          </td>
                        </>
                      )}
                    </tr>
                  );
                })}
                <tr className="ts-total-row">
                  <td colSpan={3} style={{ fontWeight: 800 }}>Total</td>
                  {yearStats.map((ys, y) => {
                    const colLetter = String.fromCharCode(68 + y);
                    return (
                      <td
                        key={y}
                        className={`prpa-clickable-cell ${selectedCell.cellId === colLetter + "14" ? "prpa-cell-active" : ""}`}
                        onClick={() =>
                          inspectCell(
                            `${colLetter}14`,
                            `${years[y]} Total Student Cohort`,
                            `=SUM(${colLetter}6:${colLetter}13)`,
                            `Formula: Sum of students in ${years[y]} = ${ys.totalStudents}`
                          )
                        }
                      >
                        <strong>{ys.totalStudents}</strong>
                      </td>
                    );
                  })}
                  {yearStats.map((ys, y) => {
                    const colLetter = String.fromCharCode(71 + y);
                    return (
                      <td
                        key={y}
                        className={`prpa-clickable-cell ${selectedCell.cellId === colLetter + "14" ? "prpa-cell-active" : ""}`}
                        onClick={() =>
                          inspectCell(
                            `${colLetter}14`,
                            `${years[y]} Total Marks Cohort`,
                            `=SUM(${colLetter}6:${colLetter}13)`,
                            `Formula: Sum of marks in ${years[y]} = ${ys.totalMarks}`
                          )
                        }
                      >
                        <strong>{ys.totalMarks}</strong>
                      </td>
                    );
                  })}
                  <td colSpan={3}></td>
                  <td
                    style={{ fontWeight: 900, fontSize: 14, background: "#fce4d6" }}
                    className={`prpa-clickable-cell ${selectedCell.cellId === "M14" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "M14",
                        "Target Marks Percentage Threshold ($E$8)",
                        "=ROUNDUP(M6,0)",
                        `Formula: ROUNDUP of average result (${avgResultPct.toFixed(2)}%) = ${targetMarksPct}%`
                      )
                    }
                  >
                    {targetMarksPct}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* 2. Students scoring >= Target Marks & Conclusion (Pic 1 Section 2) */}
        <div className="ts-two-col-grid" style={{ marginTop: 26 }}>
          <div className="eval-table-wrap">
            <table className="eval-table">
              <tbody>
                <tr>
                  <td style={{ textAlign: "left", fontWeight: 700 }}>No. of students having marks &gt;=</td>
                  <td className="eval-cell-peach" style={{ width: 60 }}><strong>{targetMarksPct}</strong></td>
                  {studentsAboveThreshold.map((st, y) => (
                    <td
                      key={y}
                      style={{ width: 60 }}
                      className={`prpa-clickable-cell ${selectedCell.cellId === "C16" ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          "C16",
                          `Number of students in ${years[y]} scoring >= ${targetMarksPct}%`,
                          "=SUM(D6:D8)",
                          `Formula: Sum of students in AA, AB, BB >= ${targetMarksPct}% = ${st.count}`
                        )
                      }
                    >
                      <strong>{st.count}</strong>
                    </td>
                  ))}
                </tr>
                <tr>
                  <td style={{ textAlign: "left", fontWeight: 700 }}>% of students having marks &gt;=</td>
                  <td className="eval-cell-peach"><strong>{targetMarksPct}</strong></td>
                  <td
                    style={{ width: 60 }}
                    className={`prpa-clickable-cell ${selectedCell.cellId === "C17" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "C17",
                        `% of students in ${years[0]} scoring >= ${targetMarksPct}%`,
                        "=(C16/D14)*100",
                        `Formula: [Students Scoring >= ${targetMarksPct}% (${studentsAboveThreshold[0]?.count})] ÷ [Total (${yearStats[0].totalStudents})] × 100 = ${studentsAboveThreshold[0]?.pct.toFixed(2)}%`
                      )
                    }
                  >
                    <strong>{studentsAboveThreshold[0]?.pct.toFixed(2)}</strong>
                  </td>
                  <td className="eval-cell-grey" colSpan={2}></td>
                </tr>
                <tr>
                  <td style={{ textAlign: "left", fontWeight: 700 }}>Avg of 3 years of who score &gt;=</td>
                  <td className="eval-cell-peach"><strong>{targetMarksPct}</strong></td>
                  <td
                    colSpan={3}
                    style={{ fontWeight: 800 }}
                    className={`prpa-clickable-cell ${selectedCell.cellId === "C18" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "C18",
                        "Average % of Students Meeting Threshold",
                        "=AVERAGE(C17:E17)",
                        `Formula: Average of students >= ${targetMarksPct}% marks across valid years = ${avgAbovePct.toFixed(2)}%`
                      )
                    }
                  >
                    <strong>{avgAbovePct.toFixed(2)}</strong>
                  </td>
                </tr>
                <tr>
                  <td style={{ textAlign: "left", fontWeight: 700 }}>Conclusion - From last three years' results,</td>
                  <td className="eval-cell-peach">
                    <input
                      type="number"
                      className="eval-cell-input"
                      value={targetStudentsPct}
                      onChange={(e) => setTargetField("targetStudentsPct", Number(e.target.value))}
                      onFocus={() =>
                        inspectCell(
                          "B19",
                          "Target % of Students (a / c)",
                          `=${targetStudentsPct}%`,
                          `Target threshold: at least ${targetStudentsPct}% students must score >= ${targetMarksPct}% marks`
                        )
                      }
                    />
                  </td>
                  <td colSpan={2} style={{ fontWeight: 700, fontSize: 11.5 }}>% students who scored &gt;=</td>
                  <td className="eval-cell-peach">
                    <input
                      type="number"
                      className="eval-cell-input"
                      value={targetMarksPct}
                      onChange={(e) => setTargetField("targetMarksPct", Number(e.target.value))}
                      onFocus={() =>
                        inspectCell(
                          "E19",
                          "Target % of Marks (b / d)",
                          `=${targetMarksPct}%`,
                          `Target marks threshold: student achieves target if % marks >= ${targetMarksPct}%`
                        )
                      }
                    />
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Right Target Summary Box (Pic 1 Right Table) */}
          <div className="eval-table-wrap">
            <table className="eval-table">
              <thead>
                <tr>
                  <th style={{ width: 140 }}>Course Outcome Target for term ({state.courseInfo.year || "2023-24"}) 231</th>
                  <th className="eval-cell-peach" style={{ width: 70 }}>CO Target</th>
                  <th className="eval-cell-peach" style={{ width: 70 }}>% Students</th>
                  <th className="eval-cell-peach" style={{ width: 70 }}>% marks</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style={{ fontWeight: 700 }}>Target Setting</td>
                  <td>
                    <input
                      type="number"
                      step="0.05"
                      className="eval-cell-input"
                      value={targetLevel}
                      onChange={(e) => setTargetField("targetLevel", Number(e.target.value))}
                      onFocus={() =>
                        inspectCell(
                          "H16",
                          "Course Outcome Target Level (CO Target)",
                          `=${targetLevel.toFixed(2)}`,
                          `Institutional defined target attainment level for this course = ${targetLevel.toFixed(2)}`
                        )
                      }
                    />
                  </td>
                  <td
                    className={`prpa-clickable-cell ${selectedCell.cellId === "I16" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "I16",
                        "% Students Target (a / c)",
                        `=${targetStudentsPct}%`,
                        `Target student % criteria: ${targetStudentsPct}% students`
                      )
                    }
                  >
                    <strong>{targetStudentsPct}</strong>
                  </td>
                  <td
                    className={`prpa-clickable-cell ${selectedCell.cellId === "J16" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "J16",
                        "% Marks Target (b / d)",
                        `=${targetMarksPct}%`,
                        `Target marks criteria: ${targetMarksPct}% marks`
                      )
                    }
                  >
                    <strong>{targetMarksPct}</strong>
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* 3. Attainment Level Range Matrix & Target vs Achieved (Pic 1 Section 3) */}
        <div className="ts-two-col-grid" style={{ marginTop: 26 }}>
          <div className="eval-table-wrap">
            <table className="eval-table">
              <thead>
                <tr>
                  <th colSpan={5} className="eval-cell-peach">Attainment Level Range Matrix</th>
                </tr>
                <tr>
                  <th rowSpan={2} style={{ width: 110 }}>Target statements</th>
                  <th rowSpan={2} style={{ width: 90 }}>Attainment Levels</th>
                  <th rowSpan={2} style={{ width: 90 }}>% of average marks</th>
                  <th colSpan={2}>Target: % of Students</th>
                </tr>
                <tr>
                  <th style={{ width: 60 }}>Min</th>
                  <th style={{ width: 60 }}>Max</th>
                </tr>
              </thead>
              <tbody>
                {ts.rangeMatrix.map((rm, idx) => (
                  <tr key={idx}>
                    {idx === 0 && (
                      <td rowSpan={4} style={{ verticalAlign: "middle", fontWeight: 800 }}>
                        Target statements
                      </td>
                    )}
                    <td style={{ color: "#d32f2f", fontWeight: 700 }}>{rm.level}</td>
                    <td style={{ color: "#d32f2f", fontWeight: 700 }}>{targetMarksPct}</td>
                    <td className={idx > 0 ? "eval-cell-amber" : ""}><strong>{rm.minStudents}</strong></td>
                    <td>{rm.maxStudents}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          {/* Target vs Achieved Comparison Box */}
          <div className="eval-table-wrap">
            <table className="eval-table">
              <thead>
                <tr>
                  <th>Outcome Assessment Status</th>
                  <th className="eval-cell-peach" style={{ width: 75 }}>CO Target</th>
                  <th className="eval-cell-peach" style={{ width: 75 }}>% Student</th>
                  <th className="eval-cell-peach" style={{ width: 75 }}>% marks</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style={{ textAlign: "left", paddingLeft: 8, fontWeight: 700 }}>
                    Course Outcome Target for term {state.courseInfo.year || "2023-24"} (231)
                  </td>
                  <td><strong>{targetLevel.toFixed(2)}</strong></td>
                  <td><strong>{targetStudentsPct}</strong></td>
                  <td><strong>{targetMarksPct}</strong></td>
                </tr>
                <tr>
                  <td style={{ textAlign: "left", paddingLeft: 8, fontWeight: 700, background: "#eef2ff" }}>
                    Course Outcome Achieved for term {state.courseInfo.year || "2023-24"} (231)
                  </td>
                  <td style={{ background: "#eef2ff", color: "#4338ca", fontWeight: 800 }}>
                    {liveAchieved.level}
                  </td>
                  <td style={{ background: "#eef2ff", color: "#4338ca", fontWeight: 800 }}>
                    {liveAchieved.pct}
                  </td>
                  <td style={{ background: "#eef2ff", fontWeight: 700 }}>
                    {targetMarksPct}
                  </td>
                </tr>
              </tbody>
            </table>

            {/* Quick summary note */}
            <div style={{ marginTop: 14, padding: 12, borderRadius: 8, background: "#f8fafc", border: "1px solid #e2e8f0", fontSize: 12.5 }}>
              <div style={{ fontWeight: 700, color: "#1e293b", display: "flex", alignItems: "center", gap: 6 }}>
                <TrendingUp size={15} color="#4338ca" /> Mathematical Summary:
              </div>
              <p style={{ margin: "4px 0 0", color: "#64748b", lineHeight: 1.4 }}>
                Based on historical average result of <strong>{avgResultPct.toFixed(2)}%</strong>, the target mark threshold is mathematically calculated as <strong>{targetMarksPct}%</strong>.
                Historically, <strong>{avgAbovePct.toFixed(2)}%</strong> of students met this threshold, defining the Course Target at <strong>{targetLevel.toFixed(2)}</strong> ({targetStudentsPct}% students).
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   PRACTICAL ASSESSMENT TAB (PR_PA / PR_ESE) — EXACT INSTITUTIONAL FORMAT
============================================================================ */
function PrAssessmentTab({
  state,
  setState,
  showToast,
  onExportExcel,
  dataKey = "internal1",
  title = "PR_PA — Practical Progressive Assessment",
  titleShort = "PR_PA",
  subtitle = "CO-wise practical continuous evaluation, marks percentage calculation & attainment analysis",
  isEse = false,
}) {
  const [searchTerm, setSearchTerm] = useState("");

  // Selected cell for interactive Excel formula bar & live equation inspector
  const [selectedCell, setSelectedCell] = useState({
    cellId: "Q14",
    title: `Student 1 Total Marks % (${titleShort})`,
    formula: '=IFERROR(ROUND(IF($B14<>"",IF($J14>0,$J14*100/$J$12,0),""),2),"")',
    math: "Formula: [Total Marks ($J14: 20)] × 100 ÷ [Max Total ($J$12: 25)] = 80.00%",
    category: "percentage",
  });

  const numCos = state.numCos || 5;
  const activeCOs = COs.slice(0, numCos);
  const data = state[dataKey] || (isEse ? state.assignment : state.internal1);
  const targetMarksPct = state.targetSetting?.targetMarksPct || 61.00;
  const targetSetting = state.targetSetting || DEFAULT_TARGET_SETTING;

  const totalMaxMarks = useMemo(() => {
    return activeCOs.reduce((sum, _, i) => sum + (Number(data.maxMarks?.[i]) || 0), 0);
  }, [activeCOs, data.maxMarks]);

  // Compute student percentage and Y/N for each row dynamically
  const studentRows = useMemo(() => {
    return (data.students || []).map((s, idx) => {
      const marks = activeCOs.map((_, i) => Number(s.marks?.[i]) || 0);
      const totalMark = marks.reduce((a, b) => a + b, 0);
      const pcts = activeCOs.map((_, i) => {
        const max = Number(data.maxMarks?.[i]) || 1;
        return max > 0 ? (marks[i] / max) * 100 : 0;
      });
      const totalPct = totalMaxMarks > 0 ? (totalMark / totalMaxMarks) * 100 : 0;
      const yns = pcts.map((p) => p >= targetMarksPct);
      const totalYn = totalPct >= targetMarksPct;
      return {
        ...s,
        origIdx: idx,
        marks,
        totalMark,
        pcts,
        totalPct,
        yns,
        totalYn,
      };
    });
  }, [data.students, activeCOs, data.maxMarks, totalMaxMarks, targetMarksPct]);

  // Filtered students for search
  const filteredStudents = useMemo(() => {
    if (!searchTerm.trim()) return studentRows;
    const term = searchTerm.toLowerCase();
    return studentRows.filter(
      (s) =>
        (s.roll || "").toLowerCase().includes(term) ||
        (s.name || "").toLowerCase().includes(term) ||
        String(s.origIdx + 1).includes(term)
    );
  }, [studentRows, searchTerm]);

  // Live Counts & Calculations from real given student rows
  const displayCOStats = useMemo(() => {
    const n = studentRows.length;
    return activeCOs.map((_, i) => {
      const max = Number(data.maxMarks?.[i]) || 0;
      const targetMks = max * (targetMarksPct / 100);
      const attainedCount = studentRows.filter((s) => s.yns[i]).length;
      const studentPct = n > 0 ? (attainedCount / n) * 100 : 0;
      const level = calculateAttainmentLevel(studentPct, targetSetting);
      return { max, targetMks, attainedCount, studentPct, level, total: n };
    });
  }, [studentRows, activeCOs, data.maxMarks, targetMarksPct, targetSetting]);

  const displayTotalStats = useMemo(() => {
    const n = studentRows.length;
    const targetMks = totalMaxMarks * (targetMarksPct / 100);
    const attainedCount = studentRows.filter((s) => s.totalYn).length;
    const studentPct = n > 0 ? (attainedCount / n) * 100 : 0;
    const level = calculateAttainmentLevel(studentPct, targetSetting);
    return { max: totalMaxMarks, targetMks, attainedCount, studentPct, level, total: n };
  }, [studentRows, totalMaxMarks, targetMarksPct, targetSetting]);

  function inspectCell(cellId, title, formula, math, category = "general") {
    setSelectedCell({ cellId, title, formula, math, category });
  }

  function handleMaxChange(qIdx, val) {
    const v = clamp(val, 0, 100);
    setState((s) => {
      const current = s[dataKey] || { maxMarks: Array(6).fill(0), students: [] };
      const mm = [...(current.maxMarks || Array(6).fill(0))];
      mm[qIdx] = v;
      return { ...s, [dataKey]: { ...current, maxMarks: mm } };
    });
  }

  function handleMarkChange(rowIdx, qIdx, val) {
    const max = Number(data.maxMarks?.[qIdx]) || 100;
    const v = clamp(val, 0, max);
    setState((s) => {
      const current = s[dataKey] || { maxMarks: Array(6).fill(0), students: [] };
      const students = current.students.map((st, i) => {
        if (i !== rowIdx) return st;
        const mm = [...st.marks];
        mm[qIdx] = v;
        return { ...st, marks: mm };
      });
      return { ...s, [dataKey]: { ...current, students } };
    });
  }

  function handleFieldChange(rowIdx, field, val) {
    setState((s) => {
      const current = s[dataKey] || { maxMarks: Array(6).fill(0), students: [] };
      const students = current.students.map((st, i) => (i === rowIdx ? { ...st, [field]: val } : st));
      return { ...s, [dataKey]: { ...current, students } };
    });
  }

  function handleTargetMarksChange(val) {
    const bounded = clamp(val, 0, 100);
    setState((s) => ({
      ...s,
      targetSetting: { ...s.targetSetting, targetMarksPct: bounded },
      targets: {
        ...s.targets,
        targetPctCO: COs.map(() => bounded / 100),
      },
    }));
  }

  function handleAddStudent() {
    setState((s) => {
      const current = s[dataKey] || { maxMarks: Array(6).fill(0), students: [] };
      const nextRoll = `206310307${String(198 + current.students.length + 1).padStart(3, "0")}`;
      const newSt = { roll: nextRoll, name: "", marks: Array(6).fill(0) };
      return {
        ...s,
        [dataKey]: { ...current, students: [...current.students, newSt] },
      };
    });
    showToast(`Added new student row in ${titleShort}`);
  }

  function handleClearStudent(idx) {
    setState((s) => {
      const current = s[dataKey] || { maxMarks: Array(6).fill(0), students: [] };
      const students = current.students.map((st, i) => (i === idx ? { ...st, marks: Array(6).fill(0) } : st));
      return { ...s, [dataKey]: { ...current, students } };
    });
    showToast("Cleared marks for student");
  }

  function handleDeleteStudent(idx) {
    setState((s) => {
      const current = s[dataKey] || { maxMarks: Array(6).fill(0), students: [] };
      const students = current.students.filter((_, i) => i !== idx);
      return { ...s, [dataKey]: { ...current, students } };
    });
    showToast("Removed student row");
  }

  function handleResetSample() {
    setState((s) => {
      if (isEse) {
        return {
          ...s,
          assignment: {
            maxMarks: [5, 5, 5, 5, 5, 0],
            students: JSON.parse(JSON.stringify(SAMPLE.assignment.students)),
          },
          targetSetting: {
            ...s.targetSetting,
            targetMarksPct: 61,
          },
        };
      }
      return {
        ...s,
        internal1: {
          maxMarks: [2, 14, 5, 2, 2, 0],
          students: JSON.parse(JSON.stringify(SAMPLE_PR_PA_STUDENTS)),
        },
        targetSetting: {
          ...s.targetSetting,
          targetMarksPct: 61,
        },
      };
    });
    showToast(`Reset ${titleShort} to sample data from reference image`);
  }

  function handleClearAll() {
    if (!window.confirm(`Clear all student marks in ${titleShort}?`)) return;
    setState((s) => {
      const current = s[dataKey] || { maxMarks: Array(6).fill(0), students: [] };
      return {
        ...s,
        [dataKey]: {
          ...current,
          students: current.students.map((st) => ({ ...st, marks: Array(6).fill(0) })),
        },
      };
    });
    showToast("All marks cleared");
  }

  return (
    <div className="eval-plan-wrapper prpa-container">
      {/* Top Toolbar */}
      <div className="eval-toolbar hide-on-print">
        <div className="eval-toolbar-left">
          <ClipboardList className="eval-toolbar-icon" size={22} />
          <div>
            <h2 className="eval-toolbar-title">{title}</h2>
            <p className="eval-toolbar-sub">{subtitle}</p>
          </div>
        </div>
        <div className="eval-toolbar-actions">
          <button className="btn-ghost" onClick={handleResetSample} title="Load sample students from image">
            <RefreshCw size={14} /> Sample Students ({data.students?.length || 24})
          </button>
          <button className="btn-ghost" onClick={handleAddStudent}>
            <Plus size={14} /> Add Student
          </button>
          <button className="btn-ghost danger" onClick={handleClearAll}>
            <Trash2 size={14} /> Clear Marks
          </button>
          <button className="btn-ghost" onClick={onExportExcel}>
            <Download size={14} /> Export {titleShort}
          </button>
          <button className="btn-ghost" onClick={() => window.print()}>
            <Printer size={14} /> Print Sheet
          </button>
        </div>
      </div>

      {/* College & Department Configuration Card (For UI editing, omitted from exported sheets) */}
      <div className="prpa-inst-box hide-on-print">
        <div className="prpa-inst-header">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <BookOpenCheck size={18} color="#4338ca" />
            <strong style={{ color: "#1e1b4b", fontSize: 13.5 }}>Institution &amp; Department Details</strong>
          </div>
          <span className="prpa-inst-badge">
            ℹ️ Configured here for Course Evaluation Plan and institutional reports, but kept off {titleShort} export to match the official single-header format.
          </span>
        </div>
        <div className="prpa-inst-grid">
          <div className="prpa-inst-field">
            <label className="prpa-inst-label">College / University Name:</label>
            <input
              type="text"
              className="prpa-inst-input"
              value={state.courseInfo.institute || ""}
              placeholder="e.g. K.D.POLYTECHNIC,PATAN"
              onChange={(e) =>
                setState((s) => ({
                  ...s,
                  courseInfo: { ...s.courseInfo, institute: e.target.value },
                }))
              }
            />
          </div>
          <div className="prpa-inst-field">
            <label className="prpa-inst-label">Department Name:</label>
            <input
              type="text"
              className="prpa-inst-input"
              value={state.courseInfo.department || ""}
              placeholder="e.g. COMPUTER ENGINEERING DEPARTMENT"
              onChange={(e) =>
                setState((s) => ({
                  ...s,
                  courseInfo: { ...s.courseInfo, department: e.target.value },
                }))
              }
            />
          </div>
        </div>
      </div>

      {/* Interactive Excel Formula Bar (fx) */}
      <div className="prpa-fx-bar hide-on-print">
        <div className="prpa-fx-left">
          <div className="prpa-fx-name-box" title="Selected Excel Cell Reference">
            {selectedCell.cellId}
          </div>
          <div className="prpa-fx-icon" title="Formula (fx)">
            <em>f</em><span>x</span>
          </div>
        </div>
        <div className="prpa-fx-content">
          <div className="prpa-fx-formula" title="Exact Excel Mathematical Formula">
            <code>{selectedCell.formula}</code>
          </div>
          <div className="prpa-fx-math" title="Live Step-by-Step Mathematical Calculation Breakdown">
            <span className="prpa-fx-pill">📐 {selectedCell.title}:</span>
            <strong>{selectedCell.math}</strong>
          </div>
        </div>
      </div>

      {/* Main Sheet Container Matching Reference Image */}
      <div className="prpa-sheet">
        {/* Top Split Section */}
        <div className="prpa-top-grid">
          {/* Top Left: Course Info & Target Setting Table */}
          <div className="prpa-card-box">
            <table className="prpa-info-table">
              <tbody>
                <tr>
                  <td className="prpa-lbl" style={{ width: 140 }}>Course Code:</td>
                  <td className="prpa-val-bold" style={{ width: 110 }}>
                    <input
                      className="eval-cell-input"
                      value={state.courseInfo.courseCode || "4350702"}
                      onChange={(e) =>
                        setState((s) => ({
                          ...s,
                          courseInfo: { ...s.courseInfo, courseCode: e.target.value },
                        }))
                      }
                      onFocus={() =>
                        inspectCell(
                          "B1",
                          "Course Code",
                          `=${state.courseInfo.courseCode || "4350702"}`,
                          `Subject GTU Course Code: ${state.courseInfo.courseCode || "4350702"}`
                        )
                      }
                    />
                  </td>
                  <td className="prpa-lbl" style={{ width: 110 }}>Course Name:</td>
                  <td className="prpa-val-bold prpa-val-name" colSpan={2}>
                    <input
                      className="eval-cell-input eval-input-wide"
                      value={state.courseInfo.courseName || "Introduction to Machine Learning"}
                      onChange={(e) =>
                        setState((s) => ({
                          ...s,
                          courseInfo: { ...s.courseInfo, courseName: e.target.value },
                        }))
                      }
                      onFocus={() =>
                        inspectCell(
                          "D1",
                          "Course Name",
                          `=${state.courseInfo.courseName || "Introduction to Machine Learning"}`,
                          `Subject Course Name: ${state.courseInfo.courseName || "Introduction to Machine Learning"}`
                        )
                      }
                    />
                  </td>
                </tr>
                <tr>
                  <td className="prpa-lbl">Batch:</td>
                  <td className="prpa-val">
                    <input
                      className="eval-cell-input"
                      value={state.courseInfo.batch || "2021-24"}
                      onChange={(e) =>
                        setState((s) => ({
                          ...s,
                          courseInfo: { ...s.courseInfo, batch: e.target.value },
                        }))
                      }
                    />
                  </td>
                  <td className="prpa-lbl">Term:</td>
                  <td className="prpa-val" colSpan={2}>
                    <input
                      className="eval-cell-input"
                      value={state.courseInfo.term || "231"}
                      onChange={(e) =>
                        setState((s) => ({
                          ...s,
                          courseInfo: { ...s.courseInfo, term: e.target.value },
                        }))
                      }
                    />
                  </td>
                </tr>
                <tr>
                  <td className="prpa-lbl">Number of Students:</td>
                  <td className="prpa-val-bold" style={{ color: "#1e1b4b" }}>
                    <input
                      type="number"
                      className="eval-cell-input"
                      value={state.courseInfo.numStudents || numStudentsCohort}
                      onChange={(e) =>
                        setState((s) => ({
                          ...s,
                          courseInfo: { ...s.courseInfo, numStudents: Number(e.target.value) },
                        }))
                      }
                      onFocus={() =>
                        inspectCell(
                          "B3",
                          "Total Enrolled Student Cohort ($B$3)",
                          `=${state.courseInfo.numStudents || numStudentsCohort}`,
                          `Total cohort strength: ${state.courseInfo.numStudents || numStudentsCohort} students`
                        )
                      }
                    />
                  </td>
                  <td className="prpa-lbl">Semester :</td>
                  <td className="prpa-val" colSpan={2}>
                    <input
                      className="eval-cell-input"
                      value={state.courseInfo.semester || "5"}
                      onChange={(e) =>
                        setState((s) => ({
                          ...s,
                          courseInfo: { ...s.courseInfo, semester: e.target.value },
                        }))
                      }
                    />
                  </td>
                </tr>
                <tr>
                  <td colSpan={3} className="prpa-target-lbl-merged">
                    TARGET MARKS PERCENTAGE
                  </td>
                  <td className="prpa-target-val-box" colSpan={2}>
                    <input
                      type="number"
                      step="0.01"
                      className="prpa-target-input"
                      value={targetMarksPct}
                      onChange={(e) => handleTargetMarksChange(Number(e.target.value))}
                      onFocus={() =>
                        inspectCell(
                          "D4",
                          "Target Marks Percentage Threshold ($D$4)",
                          `=${targetMarksPct.toFixed(2)}%`,
                          `Target threshold: students scoring >= ${targetMarksPct.toFixed(2)}% marks achieve CO attainment ('Y')`
                        )
                      }
                    />
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Top Right: STUDENTS ACHIEVING TARGET Table */}
          <div className="prpa-card-box">
            <table className="prpa-achieving-table">
              <thead>
                <tr>
                  <th className="prpa-achieving-title" style={{ width: 140 }}>STUDENTS ACHIEVING TARGET</th>
                  {activeCOs.map((co) => (
                    <th key={co} className="eval-cell-peach" style={{ width: 44 }}>{co}</th>
                  ))}
                  {numCos < 6 && <th className="eval-cell-peach" style={{ width: 44 }}>CO6</th>}
                  <th className="eval-cell-peach" style={{ width: 56 }}>TOTAL</th>
                </tr>
              </thead>
              <tbody>
                {/* Row 1: NO OF STUDENTS */}
                <tr>
                  <td
                    className={`prpa-metric-name prpa-clickable-cell ${selectedCell.cellId === "F2" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "F2",
                        "NO OF STUDENTS Row Summary",
                        '=COUNTIF(Attained_Column, "Y")',
                        `Formula: Counts all students with 'Y' in each CO attainment column`
                      )
                    }
                  >
                    NO OF STUDENTS
                  </td>
                  {displayCOStats.map((st, i) => {
                    const colLetter = String.fromCharCode(72 + i);
                    return (
                      <td
                        key={i}
                        className={`prpa-stat-num prpa-clickable-cell ${selectedCell.cellId === colLetter + "2" ? "prpa-cell-active" : ""}`}
                        onClick={() =>
                          inspectCell(
                            `${colLetter}2`,
                            `${COs[i]} Attaining Students Count (${colLetter}2)`,
                            `=COUNTIF(R14:R${13 + studentRows.length}, "Y")`,
                            `Formula: Count of students achieving >= ${targetMarksPct}% marks for ${COs[i]} = ${st.attainedCount}`
                          )
                        }
                      >
                        {st.attainedCount}
                      </td>
                    );
                  })}
                  {numCos < 6 && <td className="prpa-stat-num"></td>}
                  <td
                    className={`prpa-stat-num prpa-clickable-cell ${selectedCell.cellId === "N2" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "N2",
                        "TOTAL Students Achieving Target (N2)",
                        `=COUNTIF(X14:X${13 + studentRows.length}, "Y")`,
                        `Formula: Total number of students achieving overall assessment target = ${displayTotalStats.attainedCount}`
                      )
                    }
                  >
                    {displayTotalStats.attainedCount}
                  </td>
                </tr>

                {/* Row 2: % OF STUDENTS */}
                <tr>
                  <td
                    className={`prpa-metric-name prpa-clickable-cell ${selectedCell.cellId === "F3" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "F3",
                        "% OF STUDENTS Row Summary",
                        "=ROUND((No_Of_Students / Total_Students) * 100, 2)",
                        "Formula: (Count ÷ Total Students) × 100"
                      )
                    }
                  >
                    % OF STUDENTS
                  </td>
                  {displayCOStats.map((st, i) => {
                    const colLetter = String.fromCharCode(72 + i);
                    return (
                      <td
                        key={i}
                        className={`prpa-stat-num prpa-clickable-cell ${selectedCell.cellId === colLetter + "3" ? "prpa-cell-active" : ""}`}
                        onClick={() =>
                          inspectCell(
                            `${colLetter}3`,
                            `${COs[i]} % of Students Attaining (${colLetter}3)`,
                            `=IFERROR(ROUND(${colLetter}2*100/$B$3, 2), 0)`,
                            `Formula: [Students (${st.attainedCount})] ÷ [Total Students (${studentRows.length})] × 100 = ${st.studentPct.toFixed(2)}%`
                          )
                        }
                      >
                        {st.studentPct.toFixed(2)}
                      </td>
                    );
                  })}
                  {numCos < 6 && <td className="prpa-stat-num"></td>}
                  <td
                    className={`prpa-stat-num prpa-clickable-cell ${selectedCell.cellId === "N3" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "N3",
                        "TOTAL % of Students Attaining (N3)",
                        `=IFERROR(ROUND(N2*100/$B$3, 2), 0)`,
                        `Formula: [Total Attained (${displayTotalStats.attainedCount})] ÷ [Total Students (${studentRows.length})] × 100 = ${displayTotalStats.studentPct.toFixed(2)}%`
                      )
                    }
                  >
                    {displayTotalStats.studentPct.toFixed(2)}
                  </td>
                </tr>

                {/* Row 3: ATTAINMENT LEVEL */}
                <tr>
                  <td
                    className={`prpa-metric-name prpa-clickable-cell ${selectedCell.cellId === "F4" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "F4",
                        "ATTAINMENT LEVEL Row (0–3 Scale)",
                        "=IF(%>=70, 3, IF(%>=60, 2, IF(%>=50, 1, 0.9)))",
                        "Institutional 0–3 OBE ladder mapping based on % students meeting threshold"
                      )
                    }
                  >
                    ATTAINMENT LEVEL
                  </td>
                  {displayCOStats.map((st, i) => {
                    const colLetter = String.fromCharCode(72 + i);
                    return (
                      <td
                        key={i}
                        className={`prpa-stat-num prpa-level-highlight prpa-clickable-cell ${selectedCell.cellId === colLetter + "4" ? "prpa-cell-active" : ""}`}
                        onClick={() =>
                          inspectCell(
                            `${colLetter}4`,
                            `${COs[i]} Attainment Level (${colLetter}4)`,
                            `=IF(${colLetter}3>=70,3,IF(${colLetter}3>=60,2,IF(${colLetter}3>=50,1,0.9)))`,
                            `Formula: ${st.studentPct.toFixed(2)}% students attained → Attainment Level = ${st.level.toFixed(2)}`
                          )
                        }
                      >
                        {st.level.toFixed(2)}
                      </td>
                    );
                  })}
                  {numCos < 6 && <td className="prpa-stat-num"></td>}
                  <td
                    className={`prpa-stat-num prpa-level-highlight prpa-clickable-cell ${selectedCell.cellId === "N4" ? "prpa-cell-active" : ""}`}
                    onClick={() =>
                      inspectCell(
                        "N4",
                        "TOTAL Assessment Attainment Level (N4)",
                        `=IF(N3>=70,3,IF(N3>=60,2,IF(N3>=50,1,0.9)))`,
                        `Formula: ${displayTotalStats.studentPct.toFixed(2)}% students attained → Overall Level = ${displayTotalStats.level.toFixed(2)}`
                      )
                    }
                  >
                    {displayTotalStats.level.toFixed(2)}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Filter & Live Search Toolbar */}
        <div className="prpa-filter-bar hide-on-print">
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <Search size={16} color="#64748b" />
            <input
              type="text"
              className="prpa-search-input"
              placeholder="Search by student name or enrollment no..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
            {searchTerm && (
              <button className="icon-btn" onClick={() => setSearchTerm("")} title="Clear filter">
                <X size={14} />
              </button>
            )}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <span className="prpa-student-count-tag">
              Total Evaluated: <strong>{studentRows.length}</strong> students {searchTerm && `(Filtered: ${filteredStudents.length})`}
            </span>
          </div>
        </div>

        {/* Main Grid Table Matching Reference Image Exactly */}
        <div className="prpa-scroll-table">
          <table className="prpa-grid-table">
            <thead>
              {/* Header Row 1: Super Categories */}
              <tr>
                <th colSpan={3} className="prpa-super-left"></th>
                <th colSpan={numCos + 1} className="prpa-col-co">COWISE MARKS</th>
                <th colSpan={numCos + 1} className="prpa-col-co">CO MARKS PERCENTAGE</th>
                <th colSpan={numCos + (numCos < 6 ? 2 : 1)} className="prpa-col-co">ATTAINED?</th>
                <th className="hide-on-print" style={{ width: 45 }}></th>
              </tr>

              {/* Header Row 2: RELATED CO */}
              <tr>
                <th colSpan={3} className="eval-cell-grey" style={{ textAlign: "right", paddingRight: 10 }}>
                  RELATED CO ==&gt;
                </th>
                {activeCOs.map((co) => (
                  <th key={co} className="eval-cell-peach" style={{ width: 44 }}>{co}</th>
                ))}
                <th className="eval-cell-peach" style={{ width: 48 }}>TOTAL</th>

                {activeCOs.map((co) => (
                  <th key={co} className="eval-cell-peach" style={{ width: 52 }}>{co}</th>
                ))}
                <th className="eval-cell-peach" style={{ width: 52 }}>TOTAL</th>

                {activeCOs.map((co) => (
                  <th key={co} className="eval-cell-peach" style={{ width: 38 }}>{co}</th>
                ))}
                {numCos < 6 && <th className="eval-cell-peach" style={{ width: 38 }}>CO6</th>}
                <th className="eval-cell-peach" style={{ width: 48 }}>TOTAL</th>
                <th className="hide-on-print"></th>
              </tr>

              {/* Header Row 2: MAX MARKS */}
              <tr>
                <th colSpan={3} className="eval-cell-grey" style={{ textAlign: "right", paddingRight: 10 }}>
                  MAX MARKS --&gt;
                </th>
                {activeCOs.map((_, i) => {
                  const colLetter = String.fromCharCode(68 + i);
                  return (
                    <th
                      key={i}
                      className={`eval-cell-white prpa-max-mark-th ${selectedCell.cellId === colLetter + "12" ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          `${colLetter}12`,
                          `${COs[i]} Maximum Marks Threshold`,
                          `=${data.maxMarks?.[i] ?? (isEse ? 5 : 2)}`,
                          `Maximum marks allocated for ${COs[i]} in ${titleShort} = ${data.maxMarks?.[i] ?? (isEse ? 5 : 2)}`
                        )
                      }
                    >
                      <input
                        type="number"
                        min={0}
                        max={100}
                        className="eval-cell-input prpa-max-input"
                        value={data.maxMarks?.[i] ?? (isEse ? 5 : 2)}
                        onChange={(e) => handleMaxChange(i, Number(e.target.value))}
                        onFocus={() =>
                          inspectCell(
                            `${colLetter}12`,
                            `${COs[i]} Maximum Marks Threshold`,
                            `=${data.maxMarks?.[i] ?? (isEse ? 5 : 2)}`,
                            `Maximum marks allocated for ${COs[i]} in ${titleShort} = ${data.maxMarks?.[i] ?? (isEse ? 5 : 2)}`
                          )
                        }
                      />
                    </th>
                  );
                })}
                <th
                  className={`eval-cell-white prpa-total-max-hdr prpa-clickable-cell ${selectedCell.cellId === "J12" ? "prpa-cell-active" : ""}`}
                  onClick={() =>
                    inspectCell(
                      "J12",
                      "Total Assessment Max Marks ($J$12)",
                      "=SUM(D12:H12)",
                      `Formula: Sum of all CO max marks (${(data.maxMarks || []).slice(0, numCos).join(" + ")}) = ${totalMaxMarks}`
                    )
                  }
                >
                  {totalMaxMarks}
                </th>
                {activeCOs.map((co) => (
                  <th key={co} className="eval-cell-white">100</th>
                ))}
                <th className="eval-cell-white">100</th>
                <th colSpan={numCos + (numCos < 6 ? 2 : 1)} className="eval-cell-white prpa-attained-title">
                  (Y/N)
                </th>
                <th className="hide-on-print"></th>
              </tr>

              {/* Header Row 3: Column Titles & Indexes */}
              <tr>
                <th style={{ width: 42 }}>SR. NO.</th>
                <th style={{ width: 125 }}>ENROLL NO.</th>
                <th style={{ minWidth: 230, textAlign: "left", paddingLeft: 10 }}>NAME</th>
                {activeCOs.map((_, i) => (
                  <th key={i} style={{ width: 44 }}>{i + 1}</th>
                ))}
                <th style={{ width: 48 }}></th>
                {activeCOs.map((_, i) => (
                  <th key={i} style={{ width: 52 }}>{i + 1}</th>
                ))}
                <th style={{ width: 52 }}></th>
                {activeCOs.map((_, i) => (
                  <th key={i} style={{ width: 38 }}>{i + 1}</th>
                ))}
                {numCos < 6 && <th style={{ width: 38 }}>6</th>}
                <th style={{ width: 48 }}></th>
                <th className="hide-on-print"></th>
              </tr>
            </thead>

            <tbody>
              {filteredStudents.map((s) => {
                const rIdx = s.origIdx;
                const rowExcel = rIdx + 14;
                return (
                  <tr key={rIdx}>
                    <td>{rIdx + 1}</td>
                    <td>
                      <input
                        className="eval-cell-input prpa-roll-cell"
                        value={s.roll}
                        onChange={(e) => handleFieldChange(rIdx, "roll", e.target.value)}
                      />
                    </td>
                    <td style={{ textAlign: "left" }}>
                      <input
                        className="eval-cell-input eval-input-wide prpa-name-cell"
                        value={s.name}
                        onChange={(e) => handleFieldChange(rIdx, "name", e.target.value)}
                      />
                    </td>

                    {/* COWISE MARKS */}
                    {activeCOs.map((_, qIdx) => {
                      const colMarkLetter = String.fromCharCode(68 + qIdx);
                      const cellAddress = `${colMarkLetter}${rowExcel}`;
                      return (
                        <td key={qIdx} className={selectedCell.cellId === cellAddress ? "prpa-cell-active" : ""}>
                          <input
                            type="number"
                            step="0.1"
                            min={0}
                            max={data.maxMarks?.[qIdx] || 100}
                            className="eval-cell-input prpa-mark-cell"
                            value={s.marks[qIdx] ?? 0}
                            onChange={(e) => handleMarkChange(rIdx, qIdx, Number(e.target.value))}
                            onFocus={() =>
                              inspectCell(
                                cellAddress,
                                `Student ${rIdx + 1} ${COs[qIdx]} Mark (${s.roll})`,
                                `=${s.marks[qIdx] ?? 0}`,
                                `Student scored ${s.marks[qIdx] ?? 0} out of ${data.maxMarks?.[qIdx] ?? 0} max marks for ${COs[qIdx]}`
                              )
                            }
                          />
                        </td>
                      );
                    })}
                    <td
                      className={`prpa-tot-mark-cell prpa-clickable-cell ${selectedCell.cellId === "J" + rowExcel ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          `J${rowExcel}`,
                          `Student ${rIdx + 1} Total Marks ($J${rowExcel})`,
                          `=SUM(D${rowExcel}:H${rowExcel})`,
                          `Formula: ${s.marks.slice(0, numCos).join(" + ")} = ${s.totalMark} marks (out of ${totalMaxMarks})`
                        )
                      }
                    >
                      <strong>{s.totalMark % 1 === 0 ? s.totalMark : s.totalMark.toFixed(1)}</strong>
                    </td>

                    {/* CO MARKS PERCENTAGE */}
                    {s.pcts.map((pct, qIdx) => {
                      const colPctLetter = String.fromCharCode(75 + qIdx);
                      const colMarkLetter = String.fromCharCode(68 + qIdx);
                      const cellAddress = `${colPctLetter}${rowExcel}`;
                      const formula = `=IFERROR(ROUND(IF($B${rowExcel}<>"",IF(${colMarkLetter}${rowExcel}>0,${colMarkLetter}${rowExcel}*100/${colMarkLetter}$12,0),""),2),"")`;
                      return (
                        <td
                          key={qIdx}
                          className={`prpa-pct-cell prpa-clickable-cell ${selectedCell.cellId === cellAddress ? "prpa-cell-active" : ""}`}
                          onClick={() =>
                            inspectCell(
                              cellAddress,
                              `Student ${rIdx + 1} ${COs[qIdx]} Percentage (${colPctLetter}${rowExcel})`,
                              formula,
                              `Formula: [${COs[qIdx]} Marks (${s.marks[qIdx] ?? 0})] ÷ [Max (${data.maxMarks?.[qIdx] || 1})] × 100 = ${pct.toFixed(2)}%`
                            )
                          }
                        >
                          {pct % 1 === 0 ? pct.toFixed(0) : pct.toFixed(2)}
                        </td>
                      );
                    })}
                    <td
                      className={`prpa-pct-cell prpa-tot-pct-cell prpa-clickable-cell ${selectedCell.cellId === "Q" + rowExcel ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          `Q${rowExcel}`,
                          `Student ${rIdx + 1} Total % Marks (Q${rowExcel})`,
                          `=IFERROR(ROUND(IF($B${rowExcel}<>"",IF($J${rowExcel}>0,$J${rowExcel}*100/$J$12,0),""),2),"")`,
                          `Formula: [Total Marks ($J${rowExcel}: ${s.totalMark})] ÷ [Max Total ($J$12: ${totalMaxMarks})] × 100 = ${s.totalPct.toFixed(2)}%`
                        )
                      }
                    >
                      <strong>{s.totalPct % 1 === 0 ? s.totalPct.toFixed(0) : s.totalPct.toFixed(2)}</strong>
                    </td>

                    {/* ATTAINED? [Y/N] */}
                    {s.yns.map((yn, qIdx) => {
                      const colAttLetter = String.fromCharCode(82 + qIdx);
                      const colPctLetter = String.fromCharCode(75 + qIdx);
                      const cellAddress = `${colAttLetter}${rowExcel}`;
                      const formula = `=IF($B${rowExcel}<>"",IF(${colPctLetter}${rowExcel}>=$E$8,"Y","N"),"")`;
                      return (
                        <td
                          key={qIdx}
                          className={`${yn ? "prpa-flag-y" : "prpa-flag-n"} prpa-clickable-cell ${selectedCell.cellId === cellAddress ? "prpa-cell-active" : ""}`}
                          onClick={() =>
                            inspectCell(
                              cellAddress,
                              `Student ${rIdx + 1} ${COs[qIdx]} Attained [Y/N] (${colAttLetter}${rowExcel})`,
                              formula,
                              `Formula: [${COs[qIdx]} % (${colPctLetter}${rowExcel}: ${s.pcts[qIdx].toFixed(2)}%)] >= [Target % ($E$8: ${targetMarksPct}%)] → Attained: '${yn ? "Y" : "N"}'`
                            )
                          }
                        >
                          {yn ? "Y" : "N"}
                        </td>
                      );
                    })}
                    {numCos < 6 && <td className="prpa-flag-empty"></td>}
                    <td
                      className={`${s.totalYn ? "prpa-flag-y prpa-flag-tot" : "prpa-flag-n prpa-flag-tot"} prpa-clickable-cell ${selectedCell.cellId === "X" + rowExcel ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          `X${rowExcel}`,
                          `Student ${rIdx + 1} Total Attained [Y/N] (X${rowExcel})`,
                          `=IF($B${rowExcel}<>"",IF(Q${rowExcel}>=$E$8,"Y","N"),"")`,
                          `Formula: [Total % (Q${rowExcel}: ${s.totalPct.toFixed(2)}%)] >= [Target % ($E$8: ${targetMarksPct}%)] → Attained: '${s.totalYn ? "Y" : "N"}'`
                        )
                      }
                    >
                      {s.totalYn ? "Y" : "N"}
                    </td>

                    {/* Action buttons */}
                    <td className="hide-on-print">
                      <div style={{ display: "flex", gap: 2, justifyContent: "center" }}>
                        <button
                          className="icon-btn danger"
                          title="Clear row marks"
                          onClick={() => handleClearStudent(rIdx)}
                        >
                          <Trash2 size={12} />
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Full Details Table: STUDENTS ACHIEVING TARGET Table at End */}
      <div className="eval-section-block prpa-end-section" style={{ marginTop: 28 }}>
        <div className="eval-section-badge">STUDENTS ACHIEVING TARGET — COMPREHENSIVE BREAKDOWN</div>
        <div className="eval-table-wrap">
          <table className="eval-table prpa-end-table">
            <thead>
              <tr>
                <th style={{ width: 80 }}>CO / TOTAL</th>
                <th style={{ textAlign: "left", minWidth: 260 }}>Course Outcome Statement</th>
                <th style={{ width: 85 }}>Max Marks</th>
                <th style={{ width: 115 }}>Target Marks (&gt;= {targetMarksPct}%)</th>
                <th className="eval-cell-cyan" style={{ width: 110 }}>Students Attained</th>
                <th style={{ width: 110 }}>Total Students</th>
                <th className="eval-cell-cyan" style={{ width: 105 }}>% of Students</th>
                <th className="eval-cell-peach" style={{ width: 115 }}>Attainment Level</th>
                <th style={{ width: 120 }}>Attainment Status</th>
              </tr>
            </thead>
            <tbody>
              {displayCOStats.map((st, i) => (
                <tr key={i}>
                  <td className="eval-cell-peach" style={{ fontWeight: 800 }}>{COs[i]}</td>
                  <td style={{ textAlign: "left", paddingLeft: 10 }}>
                    <div style={{ fontWeight: 700, color: "#111827", fontSize: 12 }}>
                      {state.coCodes?.[i] || DEFAULT_CO_CODES[i]}
                    </div>
                    <div style={{ color: "#4b5563", fontSize: 11.5, marginTop: 2 }}>
                      {state.coStatements?.[i] || "Course Outcome Statement for " + COs[i]}
                    </div>
                  </td>
                  <td><strong>{st.max}</strong></td>
                  <td><strong>{st.targetMks.toFixed(2)}</strong></td>
                  <td className="eval-cell-cyan"><strong>{st.attainedCount}</strong></td>
                  <td>{studentRows.length}</td>
                  <td className="eval-cell-cyan"><strong>{st.studentPct.toFixed(2)}%</strong></td>
                  <td className="eval-cell-peach" style={{ fontSize: 13 }}><strong>{st.level.toFixed(2)}</strong></td>
                  <td>
                    <span className={`level-badge level-${Math.round(st.level)}`}>
                      {st.level >= 2.0 ? "Attained (High)" : st.level >= 1.0 ? "Attained (Moderate)" : `Target Level (${targetSetting?.targetLevel?.toFixed(2) || "0.90"})`}
                    </span>
                  </td>
                </tr>
              ))}
              <tr className="ts-total-row" style={{ background: "#f8fafc", fontWeight: 800 }}>
                <td className="eval-cell-peach" style={{ fontWeight: 900 }}>TOTAL</td>
                <td style={{ textAlign: "left", paddingLeft: 10, fontWeight: 800 }}>
                  Whole {titleShort} Assessment Total Outcome
                </td>
                <td><strong>{totalMaxMarks}</strong></td>
                <td><strong>{(totalMaxMarks * (targetMarksPct / 100)).toFixed(2)}</strong></td>
                <td className="eval-cell-cyan"><strong>{displayTotalStats.attainedCount}</strong></td>
                <td>{studentRows.length}</td>
                <td className="eval-cell-cyan"><strong>{displayTotalStats.studentPct.toFixed(2)}%</strong></td>
                <td className="eval-cell-peach" style={{ fontSize: 13.5 }}><strong>{displayTotalStats.level.toFixed(2)}</strong></td>
                <td>
                  <span className={`level-badge level-${Math.round(displayTotalStats.level)}`}>
                    {displayTotalStats.level >= 2.0 ? "Attained (Level 3)" : "Under Review"}
                  </span>
                </td>
              </tr>
            </tbody>
          </table>
        </div>

        {/* Reference Matrix & Math Formulation Legend */}
        <div className="prpa-matrix-legend-box hide-on-print">
          <div style={{ fontWeight: 700, color: "#1e293b", marginBottom: 6, display: "flex", alignItems: "center", gap: 6 }}>
            <Target size={16} color="#4338ca" /> Institution Attainment Level Range Matrix
          </div>
          <div className="prpa-matrix-legend-grid">
            <div className="prpa-matrix-card">
              <div className="prpa-matrix-lvl" style={{ color: "#166534", background: "#dcfce7" }}>Level 3.00</div>
              <div className="prpa-matrix-text">&gt;= 70% students score &gt;= {targetMarksPct}% marks</div>
            </div>
            <div className="prpa-matrix-card">
              <div className="prpa-matrix-lvl" style={{ color: "#1e40af", background: "#dbeafe" }}>Level 2.00</div>
              <div className="prpa-matrix-text">60% – 69.99% students score &gt;= {targetMarksPct}% marks</div>
            </div>
            <div className="prpa-matrix-card">
              <div className="prpa-matrix-lvl" style={{ color: "#854d0e", background: "#fef9c3" }}>Level 1.00</div>
              <div className="prpa-matrix-text">50% – 59.99% students score &gt;= {targetMarksPct}% marks</div>
            </div>
            <div className="prpa-matrix-card">
              <div className="prpa-matrix-lvl" style={{ color: "#991b1b", background: "#fee2e2" }}>Level {targetSetting?.targetLevel?.toFixed(2) || "0.90"}</div>
              <div className="prpa-matrix-text">&lt; 50% students score &gt;= {targetMarksPct}% marks</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================================================================
   ASSESSMENT / SURVEY / MAPPING / RESULTS TABS
============================================================================ */
function AssessmentTab({ title, note, data, nQ, coForQ, stats, onChange, targetPctCO }) {
  function setMax(i, val) {
    const newMax = clamp(val, 0, 100);
    const mm = [...data.maxMarks];
    mm[i] = newMax;
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
  function clearStudent(idx) {
    onChange({ ...data, students: data.students.map((s, i) => (i === idx ? { ...s, marks: Array(nQ).fill(0) } : s)) });
  }
  function clearAll() {
    onChange({ ...data, students: data.students.map((s) => ({ ...s, marks: Array(nQ).fill(0) })) });
  }

  return (
    <div>
      <div className="panel-head panel-head-row">
        <div><h3>{title}</h3>{note && <p className="muted">{note}</p>}</div>
        <button className="btn-ghost danger" onClick={clearAll}><Trash2 size={14} /> Clear all</button>
      </div>
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
            <tr className="target-row">
              <td className="sticky-col">Target →</td>
              <td className="sticky-col2"></td>
              {data.maxMarks.map((m, i) => {
                const coIdx = COs.indexOf(coForQ(i));
                const pct = targetPctCO?.[coIdx] ?? 0.61;
                const marks = (Number(m) || 0) * pct;
                return <td key={i} className="target-cell">{Math.round(pct * 100)}%<div className="th-sub">{marks.toFixed(1)} mks</div></td>;
              })}
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
                  <td><button className="icon-btn danger" title="Clear this row" onClick={() => clearStudent(rIdx)}><Trash2 size={14} /></button></td>
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
  function clearStudent(idx) {
    onChange({ ...survey, students: survey.students.map((s, i) => (i === idx ? { ...s, ratings: Array(6).fill(null) } : s)) });
  }
  function clearAll() {
    onChange({ ...survey, students: survey.students.map((s) => ({ ...s, ratings: Array(6).fill(null) })) });
  }
  return (
    <div>
      <div className="panel-head panel-head-row">
        <div><h3>Course Exit Survey — Indirect CO Attainment</h3><p className="muted">Student self-rating per CO. Scale: 1 = Low, 2 = Medium, 3 = High.</p></div>
        <button className="btn-ghost danger" onClick={clearAll}><Trash2 size={14} /> Clear all</button>
      </div>
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
                    <select className="cell-input" value={r ?? ""} onChange={(e) => setRating(rIdx, cIdx, e.target.value === "" ? null : Number(e.target.value))}>
                      <option value="">–</option>
                      <option value={1}>1</option><option value={2}>2</option><option value={3}>3</option>
                    </select>
                  </td>
                ))}
                <td><button className="icon-btn danger" title="Clear this row" onClick={() => clearStudent(rIdx)}><Trash2 size={14} /></button></td>
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

function MappingTab({
  state,
  setState,
  showToast,
  onExportExcel,
  mapping: propMapping,
  targetPO: propTargetPO,
  onMapping,
  onTarget,
  numCos: propNumCos,
}) {
  const mapping = state?.mapping || propMapping || [];
  const targetPO = state?.targetPO || propTargetPO || [];
  const numCos = state?.numCos || propNumCos || 5;
  const activeCOs = COs.slice(0, numCos);

  // Selected cell for interactive Excel formula bar & live equation inspector
  const [selectedCell, setSelectedCell] = useState({
    cellId: "E8",
    title: "PO4 Average Correlation Strength (E8)",
    formula: '=IFERROR(ROUND(AVERAGE(E3:E7), 2), "")',
    math: "Formula: Average of mapped values [(CO2: 3) + (CO3: 3) + (CO4: 2) + (CO5: 2)] ÷ 4 = 2.50",
    category: "average",
  });

  function inspectCell(cellId, title, formula, math, category = "general") {
    setSelectedCell({ cellId, title, formula, math, category });
  }

  function setCell(i, j, val) {
    const nextMapping = mapping.map((row) => [...row]);
    if (!nextMapping[i]) nextMapping[i] = [];
    nextMapping[i][j] = val;
    if (setState) {
      setState((s) => ({ ...s, mapping: nextMapping }));
    } else if (onMapping) {
      onMapping(nextMapping);
    }
  }

  function setTarget(j, val) {
    const nextTarget = [...targetPO];
    nextTarget[j] = clamp(val, 0, 3);
    if (setState) {
      setState((s) => ({ ...s, targetPO: nextTarget }));
    } else if (onTarget) {
      onTarget(nextTarget);
    }
  }

  function loadSampleFromImage() {
    const imageMapping = [
      [3, 2, 0, 0, 0, 0, 2, 2, 0], // CO1
      [3, 2, 0, 3, 0, 0, 2, 3, 0], // CO2
      [3, 3, 0, 3, 0, 0, 2, 3, 0], // CO3
      [3, 3, 3, 2, 0, 0, 2, 3, 0], // CO4
      [3, 3, 3, 2, 0, 0, 2, 3, 0], // CO5
      [0, 0, 0, 0, 0, 0, 0, 0, 0], // CO6
    ];
    const imageTargetPO = [3.0, 2.6, 3.0, 2.5, 2.0, 2.0, 2.0, 2.8, 2.0];
    if (setState) {
      setState((s) => ({ ...s, mapping: imageMapping, targetPO: imageTargetPO, numCos: 5 }));
      showToast?.("Loaded reference CO-PO-PSO matrix from sheet image!");
    } else if (onMapping) {
      onMapping(imageMapping);
    }
  }

  // Calculate live dynamic mathematical averages for each PO/PSO column
  const poCalculations = useMemo(() => {
    return DIPLOMA_POPSO.map((name, colIdx) => {
      const colLetter = String.fromCharCode(66 + colIdx); // B = PO1, C = PO2, D = PO3, etc.
      const mappedEntries = [];
      for (let i = 0; i < numCos; i++) {
        const val = Number(mapping[i]?.[colIdx]) || 0;
        if (val > 0) {
          mappedEntries.push({ co: COs[i], val, rowNum: 3 + i });
        }
      }
      const count = mappedEntries.length;
      const sum = mappedEntries.reduce((acc, m) => acc + m.val, 0);
      const avg = count > 0 ? sum / count : null;
      const formula = `=IFERROR(ROUND(AVERAGE(${colLetter}3:${colLetter}${2 + numCos}), 2), "")`;
      const mathExplanation = count > 0
        ? `Formula: Average of mapped values [${mappedEntries.map((m) => `${m.co}: ${m.val}`).join(" + ")}] ÷ ${count} = ${avg.toFixed(2)}`
        : `Formula: No Course Outcomes mapped to ${name} (average is blank)`;

      return {
        name,
        colIdx,
        colLetter,
        count,
        sum,
        avg,
        avgStr: avg !== null ? avg.toFixed(2) : "",
        formula,
        mathExplanation,
      };
    });
  }, [mapping, numCos]);

  return (
    <div className="eval-tab-page">
      {/* Top Action Toolbar */}
      <div className="eval-toolbar hide-on-print">
        <div>
          <div className="eval-toolbar-title">CO-PO-PSO Mapping Matrix</div>
          <div className="eval-toolbar-sub">
            Course Outcome & Program Outcome correlation weights (1 = Slight, 2 = Moderate, 3 = Substantial)
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
          <button className="btn-ghost" onClick={loadSampleFromImage} title="Load exact sample values from reference image">
            <RefreshCw size={14} /> Sample Data (from Image)
          </button>
          {onExportExcel && (
            <button className="btn-primary" onClick={onExportExcel}>
              <Download size={14} /> Export All (.xlsx)
            </button>
          )}
          <button className="btn-ghost" onClick={() => window.print()}>
            <Printer size={14} /> Print Sheet
          </button>
        </div>
      </div>

      {/* Interactive Excel Formula Bar (fx) */}
      <div className="prpa-fx-bar hide-on-print">
        <div className="prpa-fx-left">
          <div className="prpa-fx-name-box">{selectedCell.cellId}</div>
          <div className="prpa-fx-icon">
            <em>f</em>
            <span>x</span>
          </div>
        </div>
        <div className="prpa-fx-content">
          <div className="prpa-fx-formula">
            <code>{selectedCell.formula}</code>
          </div>
          <div className="prpa-fx-math">
            <span className="prpa-fx-pill">[{selectedCell.title}]</span>
            <span>{selectedCell.math}</span>
          </div>
        </div>
      </div>

      {/* Main Worksheet Container Matching Reference Image */}
      <div className="eval-sheet-outer" style={{ maxWidth: 1050, margin: "0 auto" }}>
        {/* Top Left Header Badge from Reference Picture */}
        <div style={{ marginBottom: 10 }}>
          <div
            style={{
              background: "#366092",
              color: "#ffffff",
              fontWeight: 800,
              fontSize: 14,
              padding: "7px 16px",
              border: "1.5px solid #000000",
              display: "inline-block",
              letterSpacing: "0.5px",
            }}
          >
            CO-PO-PSO Mapping
          </div>
        </div>

        {/* CO-PO-PSO Excel Table */}
        <div className="table-scroll" style={{ border: "1.5px solid #000000", background: "#ffffff" }}>
          <table
            style={{
              width: "100%",
              borderCollapse: "collapse",
              fontFamily: '"Segoe UI", Calibri, Arial, sans-serif',
              fontSize: 13,
            }}
          >
            <thead>
              {/* Header Row: PO1..PO7 | PSO1..PSO2 */}
              <tr>
                <th
                  style={{
                    background: "#fce4d6",
                    border: "1px solid #000000",
                    width: 90,
                    padding: "7px 8px",
                  }}
                ></th>
                {DIPLOMA_POPSO.map((name, colIdx) => {
                  const isPsoStart = colIdx === 7; // PSO1 starts after PO7
                  const colLetter = String.fromCharCode(66 + colIdx);
                  return (
                    <th
                      key={name}
                      style={{
                        background: "#fce4d6",
                        border: "1px solid #000000",
                        borderLeft: isPsoStart ? "2.5px solid #000000" : "1px solid #000000",
                        padding: "7px 6px",
                        fontWeight: 800,
                        fontSize: 13,
                        color: "#000000",
                        textAlign: "center",
                        minWidth: 58,
                      }}
                      className="prpa-clickable-cell"
                      onClick={() =>
                        inspectCell(
                          `${colLetter}2`,
                          `${name} Header Column (${colLetter}2)`,
                          `=${name}`,
                          `${colIdx < 7 ? "Program Outcome" : "Program Specific Outcome"} Identifier: ${name}`
                        )
                      }
                    >
                      {name}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody>
              {/* Data Rows for CO1..CO5 */}
              {activeCOs.map((co, i) => {
                const rowNum = 3 + i;
                return (
                  <tr key={co}>
                    {/* Row Header: CO1, CO2, ... */}
                    <td
                      style={{
                        background: "#fce4d6",
                        border: "1px solid #000000",
                        fontWeight: 800,
                        fontSize: 13,
                        color: "#000000",
                        textAlign: "center",
                        padding: "6px 8px",
                      }}
                      className={`prpa-clickable-cell ${selectedCell.cellId === `A${rowNum}` ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          `A${rowNum}`,
                          `${co} Outcome Row (A${rowNum})`,
                          `=${co}`,
                          `Course Outcome identifier: ${co} (${state.coStatements?.[i] || "Course Outcome Statement"})`
                        )
                      }
                    >
                      {co}
                    </td>

                    {/* Data Cells */}
                    {DIPLOMA_POPSO.map((name, j) => {
                      const isPsoStart = j === 7;
                      const val = mapping[i]?.[j] || 0;
                      const colLetter = String.fromCharCode(66 + j);
                      const cellCoord = `${colLetter}${rowNum}`;
                      const isSelected = selectedCell.cellId === cellCoord;

                      return (
                        <td
                          key={j}
                          style={{
                            background: "#e2efda",
                            border: "1px solid #000000",
                            borderLeft: isPsoStart ? "2.5px solid #000000" : "1px solid #000000",
                            padding: "3px 4px",
                            textAlign: "center",
                          }}
                          className={`prpa-clickable-cell ${isSelected ? "prpa-cell-active" : ""}`}
                          onClick={() =>
                            inspectCell(
                              cellCoord,
                              `${co} \u2192 ${name} Mapping (${cellCoord})`,
                              `=${val > 0 ? val : `"-"`}`,
                              `Correlation level between ${co} and ${name} = ${val > 0 ? `${val} (${val === 3 ? "Substantial/High" : val === 2 ? "Moderate" : "Slight"})` : "None (-)"}`
                            )
                          }
                        >
                          <select
                            style={{
                              width: "100%",
                              textAlign: "center",
                              fontWeight: 700,
                              fontSize: 13,
                              color: "#000000",
                              border: "none",
                              background: "transparent",
                              cursor: "pointer",
                              outline: "none",
                              padding: "4px 2px",
                            }}
                            value={val}
                            onChange={(e) => setCell(i, j, Number(e.target.value))}
                          >
                            <option value={0}>-</option>
                            <option value={1}>1</option>
                            <option value={2}>2</option>
                            <option value={3}>3</option>
                          </select>
                        </td>
                      );
                    })}
                  </tr>
                );
              })}

              {/* Spacer Row Matching Reference Image */}
              <tr>
                <td
                  style={{
                    background: "#fce4d6",
                    border: "1px solid #000000",
                    height: 18,
                  }}
                ></td>
                {DIPLOMA_POPSO.map((_, j) => {
                  const isPsoStart = j === 7;
                  return (
                    <td
                      key={j}
                      style={{
                        background: "#e2efda",
                        border: "1px solid #000000",
                        borderLeft: isPsoStart ? "2.5px solid #000000" : "1px solid #000000",
                        height: 18,
                      }}
                    ></td>
                  );
                })}
              </tr>

              {/* AVERAGE Row Matching Reference Image */}
              <tr>
                <td
                  style={{
                    background: "#fce4d6",
                    border: "1px solid #000000",
                    fontWeight: 900,
                    fontSize: 13,
                    color: "#000000",
                    textAlign: "center",
                    padding: "7px 8px",
                  }}
                  className={`prpa-clickable-cell ${selectedCell.cellId === `A${3 + numCos + 1}` ? "prpa-cell-active" : ""}`}
                  onClick={() =>
                    inspectCell(
                      `A${3 + numCos + 1}`,
                      "AVERAGE Row Header",
                      "=AVERAGE(Range)",
                      "Row calculates average correlation mapping for each PO/PSO column"
                    )
                  }
                >
                  AVERAGE
                </td>

                {poCalculations.map((calc, j) => {
                  const isPsoStart = j === 7;
                  const rowNum = 3 + numCos + 1; // e.g. Row 8
                  const cellCoord = `${calc.colLetter}${rowNum}`;
                  const isSelected = selectedCell.cellId === cellCoord;

                  return (
                    <td
                      key={calc.name}
                      style={{
                        background: "#d9e1f2",
                        border: "1px solid #000000",
                        borderLeft: isPsoStart ? "2.5px solid #000000" : "1px solid #000000",
                        fontWeight: 800,
                        fontSize: 13.5,
                        color: "#000000",
                        textAlign: "center",
                        padding: "7px 4px",
                      }}
                      className={`prpa-clickable-cell ${isSelected ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          cellCoord,
                          `${calc.name} Average Correlation Strength (${cellCoord})`,
                          calc.formula,
                          calc.mathExplanation,
                          "average"
                        )
                      }
                    >
                      {calc.avgStr}
                    </td>
                  );
                })}
              </tr>
            </tbody>
          </table>
        </div>

        {/* Reference Matrix & Math Formulation Guide Card */}
        <div className="prpa-matrix-legend-box hide-on-print" style={{ marginTop: 22 }}>
          <div style={{ fontWeight: 700, color: "#1e293b", marginBottom: 8, display: "flex", alignItems: "center", gap: 6 }}>
            <Target size={16} color="#4338ca" /> NBA / OBE Mapping Strength Levels & Mathematical Formula
          </div>
          <div className="prpa-matrix-legend-grid">
            <div className="prpa-matrix-card">
              <div className="prpa-matrix-lvl" style={{ color: "#166534", background: "#dcfce7" }}>3 — Substantial</div>
              <div className="prpa-matrix-text">High correlation / direct substantial course contribution to PO</div>
            </div>
            <div className="prpa-matrix-card">
              <div className="prpa-matrix-lvl" style={{ color: "#1e40af", background: "#dbeafe" }}>2 — Moderate</div>
              <div className="prpa-matrix-text">Moderate correlation / intermediate learning contribution</div>
            </div>
            <div className="prpa-matrix-card">
              <div className="prpa-matrix-lvl" style={{ color: "#854d0e", background: "#fef9c3" }}>1 — Slight</div>
              <div className="prpa-matrix-text">Low correlation / introductory concepts only</div>
            </div>
            <div className="prpa-matrix-card">
              <div className="prpa-matrix-lvl" style={{ color: "#475569", background: "#f1f5f9" }}>– (Blank / 0)</div>
              <div className="prpa-matrix-text">No correlation (excluded from average calculation)</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function ResultsTab({ state, direct, indirect, final, poAttain, thPaStats, thEseStats, prPaStats, prEseStats, theoryBlended, practicalBlended, courseAttainment }) {
  const numCos = state.numCos || 5;
  const activeCOs = COs.slice(0, numCos);

  const coAttained = activeCOs.filter((_, i) => final[i] >= (state.targets.coTargetLevel || 0.90)).length;
  const poAttained = poAttain.filter((v, j) => v != null && v >= state.targetPO[j]).length;
  const poCount = poAttain.filter((v) => v != null).length;
  const targetLevel = state.targets.coTargetLevel || 0.90;

  const [selectedCell, setSelectedCell] = useState({
    cellId: "J6",
    title: "CO1 Final Combined Attainment (T)",
    formula: "=ROUND(H6*0.8 + I6*0.2, 2)",
    math: "Formula: [Direct Measured ($H6: 1.96) × 80%] + [Indirect Survey ($I6: 2.00) × 20%] = 1.97",
    category: "final",
  });

  function inspectCell(cellId, title, formula, math, category = "general") {
    setSelectedCell({ cellId, title, formula, math, category });
  }

  const coChartData = activeCOs.map((co, i) => ({
    name: co,
    Attained: Number(final[i].toFixed(2)),
    Direct: Number(direct[i].toFixed(2)),
    Indirect: Number(indirect[i].toFixed(2)),
    Target: targetLevel,
  }));

  const poChartData = DIPLOMA_POPSO.map((p, j) => ({
    name: p,
    Attained: poAttain[j] == null ? 0 : Number(poAttain[j].toFixed(2)),
    Target: state.targetPO[j],
  }));

  return (
    <div className="eval-plan-wrapper">
      {/* Top Header Summary Cards */}
      <div className="rings-row hide-on-print">
        <div className="results-kpi-card" style={{ borderLeft: "4px solid #4338ca" }}>
          <div className="results-kpi-title">Overall Course Attainment</div>
          <div className="results-kpi-val" style={{ color: "#4338ca" }}>{courseAttainment.toFixed(2)} <span style={{ fontSize: 13, color: "#64748b" }}>/ 3.00</span></div>
          <div className="results-kpi-sub">
            <span className={`pill ${courseAttainment >= targetLevel ? "pill-ok" : "pill-bad"}`}>
              {courseAttainment >= targetLevel ? "Course Target Met" : "Under Target"}
            </span>
          </div>
        </div>
        <Ring value={coAttained} total={numCos} label="Course Outcomes Attained" color="#4338ca" />
        <Ring value={poAttained} total={poCount || 9} label="PO / PSO Attained" color="#a21caf" />
        <div className="results-kpi-card" style={{ borderLeft: "4px solid #10b981" }}>
          <div className="results-kpi-title">Defined Course Target</div>
          <div className="results-kpi-val" style={{ color: "#10b981" }}>{targetLevel.toFixed(2)}</div>
          <div className="results-kpi-sub">Min Level required for OBE compliance</div>
        </div>
      </div>

      {/* Interactive Excel Formula Bar (fx) */}
      <div className="prpa-fx-bar hide-on-print" style={{ margin: "16px 0" }}>
        <div className="prpa-fx-left">
          <div className="prpa-fx-name-box" title="Selected Excel Cell Reference">
            {selectedCell.cellId}
          </div>
          <div className="prpa-fx-icon" title="Formula (fx)">
            <em>f</em><span>x</span>
          </div>
        </div>
        <div className="prpa-fx-content">
          <div className="prpa-fx-formula" title="Exact Excel Mathematical Formula">
            <code>{selectedCell.formula}</code>
          </div>
          <div className="prpa-fx-math" title="Live Step-by-Step Mathematical Calculation Breakdown">
            <span className="prpa-fx-pill">📐 {selectedCell.title}:</span>
            <strong>{selectedCell.math}</strong>
          </div>
        </div>
      </div>

      {/* Master Attainment Calculation & Blending Table (ATTAINMENT SUMMARY) */}
      <div className="eval-section-block">
        <div className="panel-head panel-head-row">
          <div>
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 800, color: "#1e293b" }}>
              Course Outcome Attainment Summary &amp; Blending Pipeline (Workbook 231_IML.xlsx)
            </h3>
            <p className="muted" style={{ margin: "4px 0 0", fontSize: 12.5 }}>
              Exact Institutional Weighting: Theory (30% PA + 70% ESE) \u00b7 Practical (50% PA + 50% ESE) \u00b7 Direct (66.67% TH + 33.33% PR) \u00b7 Final (80% Direct + 20% Indirect)
            </p>
          </div>
          <button className="btn-ghost" onClick={() => window.print()}>
            <Printer size={14} /> Print Summary
          </button>
        </div>

        <div className="table-scroll">
          <table className="grid-table summary-table" style={{ fontSize: 12.5 }}>
            <thead>
              <tr>
                <th rowSpan={2} className="sticky-col co-badge" style={{ width: 65, background: "#fce4d6" }}>CO</th>
                <th colSpan={3} style={{ background: "#eef2ff", color: "#3730a3" }}>THEORY ATTAINMENT (100 Mks \u2192 66.67%)</th>
                <th colSpan={3} style={{ background: "#ecfdf5", color: "#065f46" }}>PRACTICAL ATTAINMENT (50 Mks \u2192 33.33%)</th>
                <th rowSpan={2} style={{ background: "#e0f2fe", color: "#0369a1" }}>Direct Attainment (R)</th>
                <th rowSpan={2} style={{ background: "#fdf4ff", color: "#86198f" }}>Indirect Survey (S)</th>
                <th rowSpan={2} style={{ background: "#dcfce7", color: "#166534", fontWeight: 800 }}>Final Combined CO Attainment (T)</th>
                <th rowSpan={2} style={{ background: "#fef9c3", color: "#854d0e" }}>Target Level</th>
                <th rowSpan={2} style={{ background: "#f1f5f9", color: "#475569" }}>Attainment Gap</th>
                <th rowSpan={2} style={{ background: "#f1f5f9" }}>Status</th>
              </tr>
              <tr>
                <th style={{ background: "#eef2ff", fontSize: 11.5 }}>Mid-Sem PA (30%)</th>
                <th style={{ background: "#eef2ff", fontSize: 11.5 }}>GTU ESE (70%)</th>
                <th style={{ background: "#e0e7ff", fontWeight: 700, fontSize: 11.5 }}>Theory Blended (N)</th>
                <th style={{ background: "#ecfdf5", fontSize: 11.5 }}>PA Journal (50%)</th>
                <th style={{ background: "#ecfdf5", fontSize: 11.5 }}>ESE Viva (50%)</th>
                <th style={{ background: "#d1fae5", fontWeight: 700, fontSize: 11.5 }}>Practical Blended (Q)</th>
              </tr>
            </thead>
            <tbody>
              {activeCOs.map((co, i) => {
                const thPa = thPaStats?.[i]?.level ?? 1.80;
                const thEse = thEseStats?.[i]?.level ?? 2.00;
                const thBlend = theoryBlended?.[i] ?? ((thPa * 0.30) + (thEse * 0.70));
                const prPa = prPaStats?.[i]?.level ?? 2.00;
                const prEse = prEseStats?.[i]?.level ?? 2.00;
                const prBlend = practicalBlended?.[i] ?? ((prPa * 0.50) + (prEse * 0.50));
                const directM = direct[i];
                const ind = indirect[i];
                const finalVal = final[i];
                const gap = Number((targetLevel - finalVal).toFixed(2));
                const isAttained = finalVal >= targetLevel;
                const rExcel = 6 + i;

                return (
                  <tr key={co}>
                    <td className="sticky-col co-badge" style={{ fontWeight: 800 }}>{co}</td>
                    {/* Theory PA */}
                    <td
                      className={`prpa-clickable-cell ${selectedCell.cellId === "B" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() => inspectCell("B" + rExcel, `${co} Theory Mid-Sem PA Level`, `='Internal_Exam2'!G${rExcel}`, `Mid-Sem Theory PA attainment level for ${co} = ${thPa.toFixed(2)}`)}
                    >
                      {thPa.toFixed(2)}
                    </td>
                    {/* Theory ESE */}
                    <td
                      className={`prpa-clickable-cell ${selectedCell.cellId === "C" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() => inspectCell("C" + rExcel, `${co} Theory GTU ESE Level`, `='End_Sem_Exam'!G${rExcel}`, `GTU University Theory Exam attainment level for ${co} = ${thEse.toFixed(2)}`)}
                    >
                      {thEse.toFixed(2)}
                    </td>
                    {/* Theory Blended (N) */}
                    <td
                      style={{ background: "#f5f3ff", fontWeight: 700 }}
                      className={`prpa-clickable-cell ${selectedCell.cellId === "D" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          "D" + rExcel,
                          `${co} Theory Blended Attainment (N)`,
                          `=ROUND(B${rExcel}*0.3 + C${rExcel}*0.7, 2)`,
                          `Formula: [Mid-Sem PA (${thPa.toFixed(2)}) × 30%] + [GTU ESE (${thEse.toFixed(2)}) × 70%] = ${thBlend.toFixed(2)}`
                        )
                      }
                    >
                      {thBlend.toFixed(2)}
                    </td>
                    {/* Practical PA */}
                    <td
                      className={`prpa-clickable-cell ${selectedCell.cellId === "E" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() => inspectCell("E" + rExcel, `${co} Practical PA Journal Level`, `='PR_PA'!G${3 + i}`, `Practical Continuous PA attainment level for ${co} = ${prPa.toFixed(2)}`)}
                    >
                      {prPa.toFixed(2)}
                    </td>
                    {/* Practical ESE */}
                    <td
                      className={`prpa-clickable-cell ${selectedCell.cellId === "F" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() => inspectCell("F" + rExcel, `${co} Practical ESE Viva Level`, `='Assignment'!G${rExcel}`, `Practical ESE Viva Exam attainment level for ${co} = ${prEse.toFixed(2)}`)}
                    >
                      {prEse.toFixed(2)}
                    </td>
                    {/* Practical Blended (Q) */}
                    <td
                      style={{ background: "#f0fdf4", fontWeight: 700 }}
                      className={`prpa-clickable-cell ${selectedCell.cellId === "G" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          "G" + rExcel,
                          `${co} Practical Blended Attainment (Q)`,
                          `=ROUND(E${rExcel}*0.5 + F${rExcel}*0.5, 2)`,
                          `Formula: [Practical PA (${prPa.toFixed(2)}) × 50%] + [Practical ESE (${prEse.toFixed(2)}) × 50%] = ${prBlend.toFixed(2)}`
                        )
                      }
                    >
                      {prBlend.toFixed(2)}
                    </td>
                    {/* Direct Measured (R) */}
                    <td
                      style={{ background: "#e0f2fe", fontWeight: 700, color: "#0369a1" }}
                      className={`prpa-clickable-cell ${selectedCell.cellId === "H" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          "H" + rExcel,
                          `${co} Direct Measured Attainment (R)`,
                          `=ROUND(D${rExcel}*(100/150) + G${rExcel}*(50/150), 2)`,
                          `Formula: [Theory (${thBlend.toFixed(2)}) × 66.67%] + [Practical (${prBlend.toFixed(2)}) × 33.33%] = ${directM.toFixed(2)}`
                        )
                      }
                    >
                      {directM.toFixed(2)}
                    </td>
                    {/* Indirect Survey (S) */}
                    <td
                      style={{ background: "#faf5ff", color: "#86198f" }}
                      className={`prpa-clickable-cell ${selectedCell.cellId === "I" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() => inspectCell("I" + rExcel, `${co} Indirect Survey Score (S)`, `='Indirect_CO_Attainment'!C${rExcel}`, `Course Exit Survey student self-rating for ${co} = ${ind.toFixed(2)}`)}
                    >
                      {ind.toFixed(2)}
                    </td>
                    {/* Final Combined CO Attainment (T) */}
                    <td
                      style={{ background: "#dcfce7", color: "#166534", fontWeight: 900, fontSize: 13.5 }}
                      className={`prpa-clickable-cell ${selectedCell.cellId === "J" + rExcel ? "prpa-cell-active" : ""}`}
                      onClick={() =>
                        inspectCell(
                          "J" + rExcel,
                          `${co} Final Combined Attainment (T)`,
                          `=ROUND(H${rExcel}*0.8 + I${rExcel}*0.2, 2)`,
                          `Formula: [Direct Measured (${directM.toFixed(2)}) × 80%] + [Indirect Survey (${ind.toFixed(2)}) × 20%] = ${finalVal.toFixed(2)}`
                        )
                      }
                    >
                      {finalVal.toFixed(2)}
                    </td>
                    {/* Target Level */}
                    <td><strong>{targetLevel.toFixed(2)}</strong></td>
                    {/* Attainment Gap */}
                    <td style={{ color: gap <= 0 ? "#15803d" : "#b91c1c", fontWeight: 700 }}>
                      {gap <= 0 ? `+${Math.abs(gap).toFixed(2)}` : `-${gap.toFixed(2)}`}
                    </td>
                    {/* Status */}
                    <td><StatusPill ok={isAttained} /></td>
                  </tr>
                );
              })}
              {/* Average / Whole Course Row */}
              <tr className="ts-total-row" style={{ background: "#f8fafc", fontWeight: 900 }}>
                <td colSpan={7} style={{ textAlign: "right", paddingRight: 14, fontWeight: 900 }}>
                  AVERAGE / WHOLE COURSE ATTAINMENT:
                </td>
                <td style={{ background: "#e0f2fe", color: "#0369a1" }}>
                  {(direct.slice(0, numCos).reduce((a, b) => a + b, 0) / numCos).toFixed(2)}
                </td>
                <td style={{ background: "#faf5ff", color: "#86198f" }}>
                  {(indirect.slice(0, numCos).reduce((a, b) => a + b, 0) / numCos).toFixed(2)}
                </td>
                <td
                  style={{ background: "#fce4d6", color: "#8a3b00", fontSize: 14, fontWeight: 900 }}
                  className={`prpa-clickable-cell ${selectedCell.cellId === "J" + (6 + numCos) ? "prpa-cell-active" : ""}`}
                  onClick={() =>
                    inspectCell(
                      "J" + (6 + numCos),
                      "Whole Course Overall Attainment",
                      `=ROUND(AVERAGE(J6:J${5 + numCos}), 2)`,
                      `Formula: Average of all ${numCos} CO final attainments = ${courseAttainment.toFixed(2)}`
                    )
                  }
                >
                  {courseAttainment.toFixed(2)}
                </td>
                <td><strong>{targetLevel.toFixed(2)}</strong></td>
                <td style={{ color: courseAttainment >= targetLevel ? "#15803d" : "#b91c1c", fontWeight: 800 }}>
                  {courseAttainment >= targetLevel ? `+${(courseAttainment - targetLevel).toFixed(2)}` : `-${(targetLevel - courseAttainment).toFixed(2)}`}
                </td>
                <td><StatusPill ok={courseAttainment >= targetLevel} /></td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>

      {/* CO Comparison Chart */}
      <div className="chart-box" style={{ marginTop: 22 }}>
        <h4 style={{ margin: "0 0 12px", fontSize: 14, color: "#1e293b" }}>Course Outcome Direct vs Indirect vs Target Benchmark</h4>
        <ResponsiveContainer width="100%" height={260}>
          <BarChart data={coChartData} margin={{ top: 10, right: 10, left: 0, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#eceefa" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis domain={[0, 3]} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Bar dataKey="Attained" name="Final Attainment" radius={[4, 4, 0, 0]}>
              {coChartData.map((d, i) => <Cell key={i} fill={d.Attained >= d.Target ? "#4338ca" : "#c7cbe8"} />)}
            </Bar>
            <ReferenceLine y={targetLevel} stroke="#a21caf" strokeDasharray="4 4" label={{ value: `Target (${targetLevel.toFixed(2)})`, position: "right", fontSize: 11, fill: "#a21caf" }} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      {/* PO / PSO Attainment Table & Chart */}
      <div className="eval-section-block" style={{ marginTop: 26 }}>
        <div className="panel-head">
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 800, color: "#1e293b" }}>PO / PSO Attainment Matrix Results</h3>
          <p className="muted" style={{ margin: "4px 0 0", fontSize: 12.5 }}>
            Computed by projecting final CO attainment scores through the CO–PO/PSO correlation mapping weights.
          </p>
        </div>
        <div className="table-scroll">
          <table className="grid-table summary-table">
            <thead><tr><th>PO / PSO</th><th>Achieved Attainment</th><th>Target Level</th><th>Attainment Gap</th><th>Status</th></tr></thead>
            <tbody>
              {DIPLOMA_POPSO.map((p, j) => (
                <tr key={p}>
                  <td className="co-badge">{p}</td>
                  <td><strong>{poAttain[j] == null ? "\u2013" : poAttain[j].toFixed(2)}</strong></td>
                  <td>{state.targetPO[j].toFixed(2)}</td>
                  <td style={{ color: poAttain[j] != null && poAttain[j] >= state.targetPO[j] ? "#15803d" : "#b91c1c", fontWeight: 700 }}>
                    {poAttain[j] == null ? "\u2013" : (poAttain[j] - state.targetPO[j]).toFixed(2)}
                  </td>
                  <td>{poAttain[j] == null ? <span className="pill">No mapping</span> : <StatusPill ok={poAttain[j] >= state.targetPO[j]} />}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="chart-box" style={{ marginTop: 18 }}>
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
  { key: "evalplan", label: "Evaluation Plan", icon: FileSpreadsheet },
  { key: "targetsetting", label: "Target Setting", icon: Target },
  { key: "mapping", label: "CO-PO Mapping", icon: GitBranch },
  { key: "prpa", label: "PR_PA", icon: ClipboardList },
  { key: "pr_ese", label: "PR_ESE", icon: ClipboardList },
  { key: "internal2", label: "TH_PA", icon: ClipboardList },
  { key: "endsem", label: "TH_ESE", icon: ClipboardList },
  { key: "survey", label: "Exit Survey", icon: Users2 },
  { key: "results", label: "Results", icon: BarChart3 },
];

function Calculator({ user, onLogout, onUserUpdate }) {
  const [state, setState] = useState(sampleState());
  const [tab, setTab] = useState("prpa");
  const [toast, setToast] = useState(null);
  const [warn, setWarn] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [myCourses, setMyCourses] = useState([]);
  const [currentId, setCurrentId] = useState(null);
  const [showCourses, setShowCourses] = useState(false);
  const [showProfileModal, setShowProfileModal] = useState(false);
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
  const numCos = state.numCos || 5;
  const i1Stats = useMemo(() => simpleCOStats(state.internal1, targets, state.targetSetting), [state.internal1, targets, state.targetSetting]);
  const i2Stats = useMemo(() => simpleCOStats(state.internal2, targets, state.targetSetting), [state.internal2, targets, state.targetSetting]);
  const asgStats = useMemo(() => simpleCOStats(state.assignment, targets, state.targetSetting), [state.assignment, targets, state.targetSetting]);
  const esStats = useMemo(() => endsemCOStats(state.endsem, targets, state.targetSetting), [state.endsem, targets, state.targetSetting]);
  const indirect = useMemo(() => surveyCOAverage(state.survey), [state.survey]);

  // Exact institutional blending pipeline from workbook 231_IML.xlsx:
  // 1. Theory Blended: N = 0.30 * Theory_PA + 0.70 * Theory_ESE
  const theoryBlended = useMemo(
    () => COs.map((_, i) => (i2Stats[i].level * 0.30) + (esStats[i].level * 0.70)),
    [i2Stats, esStats]
  );
  // 2. Practical Blended: Q = 0.50 * Practical_PA + 0.50 * Practical_ESE
  const practicalBlended = useMemo(
    () => COs.map((_, i) => (i1Stats[i].level * 0.50) + (asgStats[i].level * 0.50)),
    [i1Stats, asgStats]
  );
  // 3. Direct Measured Attainment: R = 0.6667 * Theory_Blended + 0.3333 * Practical_Blended (100:50 ratio)
  const direct = useMemo(
    () => COs.map((_, i) => (theoryBlended[i] * (100 / 150)) + (practicalBlended[i] * (50 / 150))),
    [theoryBlended, practicalBlended]
  );
  // 4. Combined CO Attainment: T = 0.80 * Direct + 0.20 * Indirect
  const final = useMemo(
    () => COs.map((_, i) => (direct[i] * (state.weights?.direct ?? 0.80)) + (indirect[i] * (state.weights?.indirect ?? 0.20))),
    [direct, indirect, state.weights]
  );
  // 5. Course Attainment: Average across active COs
  const courseAttainment = useMemo(() => {
    const activeVals = final.slice(0, numCos);
    return activeVals.length ? activeVals.reduce((a, b) => a + b, 0) / activeVals.length : 0;
  }, [final, numCos]);

  const poAttain = useMemo(
    () => DIPLOMA_POPSO.map((_, j) => {
      let num = 0, den = 0;
      for (let i = 0; i < numCos; i++) {
        const w = state.mapping[i]?.[j] || 0;
        if (w > 0) { num += final[i] * w; den += w; }
      }
      return den > 0 ? num / den : null;
    }),
    [state.mapping, final, numCos]
  );

  if (!loaded) return <div className="loading-screen">Loading your data…</div>;

  return (
    <div className="app-shell">
      <header className="app-header hide-on-print">
        <div className="app-header-left">
          <div className="brand-mini">
            <img
              src="assets/gtu_logo.png"
              alt="GTU / KDP Patan Logo"
              style={{ width: "100%", height: "100%", objectFit: "contain", borderRadius: 8 }}
              onError={(e) => {
                e.target.style.display = "none";
                if (e.target.nextSibling) e.target.nextSibling.style.display = "flex";
              }}
            />
            <span style={{ display: "none", alignItems: "center", justifyContent: "center", width: "100%", height: "100%", color: "#3730a3" }}>
              <GraduationCap size={20} />
            </span>
          </div>
          <div>
            <div className="app-title">KDP, Patan — CO/PO Attainment System</div>
            <div className="app-subtitle">{state.courseInfo.courseName || "Untitled course"} {state.courseInfo.courseCode ? `\u00b7 ${state.courseInfo.courseCode}` : ""}</div>
          </div>
        </div>
        <div className="app-header-right">
          <button
            type="button"
            className="user-chip-btn"
            onClick={() => setShowProfileModal(true)}
            title="Click to view & edit user profile / change password"
          >
            <User size={14} />
            <span>{user.displayName || user.username}</span>
          </button>
          <div className="courses-dropdown-wrap">
            <button
              type="button"
              className="btn-ghost"
              onClick={() => setShowCourses((v) => !v)}
              title="View your saved courses in the database"
            >
              <FolderOpen size={14} />
              <span>My Courses ({myCourses.length})</span>
            </button>
            {showCourses && (
              <div className="courses-dropdown">
                {myCourses.length === 0 ? (
                  <div className="courses-empty">No saved courses yet. Click "Save" to save your current course.</div>
                ) : (
                  myCourses.map((c) => (
                    <div
                      key={c.id}
                      className={`course-row ${c.id === currentId ? "course-row-active" : ""}`}
                      onClick={() => handleLoadCourse(c.id)}
                    >
                      <div>
                        <div className="course-row-name">{c.courseName || c.courseCode}</div>
                        <div className="course-row-code">{c.courseCode}</div>
                      </div>
                      <button
                        type="button"
                        className="icon-btn danger"
                        onClick={(e) => handleDeleteCourse(c.id, e)}
                        title="Delete saved course"
                      >
                        <Trash2 size={13} />
                      </button>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>
          <button
            type="button"
            className="btn-ghost"
            onClick={handleLoadSample}
            title="Load standard KDP course sample data"
          >
            <RefreshCw size={14} />
            <span>Sample</span>
          </button>
          <button
            type="button"
            className="btn-ghost"
            onClick={handleBlank}
            title="Load blank template to enter fresh course data"
          >
            <span>Blank</span>
          </button>
          <input
            type="file"
            accept=".xlsx,.xls"
            ref={fileRef}
            style={{ display: "none" }}
            onChange={handleFile}
          />
          <button
            type="button"
            className="btn-ghost"
            onClick={() => fileRef.current?.click()}
            title="Import course data from Excel workbook (.xlsx)"
          >
            <UploadCloud size={14} />
            <span>Upload Excel</span>
          </button>
          <button
            type="button"
            className="btn-ghost"
            onClick={() => exportWorkbook(state, showToast)}
            title="Export complete 11-sheet OBE attainment workbook to Excel (.xlsx)"
          >
            <Download size={14} />
            <span>Export Excel</span>
          </button>
          <button
            type="button"
            className="btn-primary"
            onClick={handleSave}
            title="Save course calculations to database"
          >
            <Save size={14} />
            <span>Save</span>
          </button>
          <button
            type="button"
            className="icon-btn"
            onClick={onLogout}
            title="Sign out of the portal"
          >
            <LogOut size={16} />
          </button>
        </div>
      </header>

      {warn && <div className="warn-banner hide-on-print"><AlertTriangle size={14} /> {warn}<button className="icon-btn" onClick={() => setWarn(null)}><X size={14} /></button></div>}

      <nav className="tab-nav hide-on-print">
        {TABS.map((t) => {
          const Icon = t.icon;
          return <button key={t.key} className={`tab-btn ${tab === t.key ? "tab-active" : ""}`} onClick={() => setTab(t.key)}><Icon size={15} /> {t.label}</button>;
        })}
      </nav>

      <main className="app-main">
        {tab === "evalplan" && <EvalPlanTab state={state} setState={setState} onExportExcel={() => exportWorkbook(state, showToast)} />}
        {tab === "targetsetting" && <TargetSettingTab state={state} setState={setState} showToast={showToast} finalDirectAttain={direct} />}
        {tab === "mapping" && (
          <MappingTab
            state={state}
            setState={setState}
            showToast={showToast}
            onExportExcel={() => exportWorkbook(state, showToast)}
          />
        )}
        {(tab === "prpa" || tab === "internal1") && (
          <PrAssessmentTab
            state={state}
            setState={setState}
            showToast={showToast}
            onExportExcel={() => exportWorkbook(state, showToast)}
            dataKey="internal1"
            title="PR_PA — Practical Progressive Assessment"
            titleShort="PR_PA"
            subtitle="Continuous Practical Assessment & Journal Evaluation"
            cohortCountsDefault={[130, 97, 59, 65, 69, 0]}
            cohortTotalDefault={112}
            isEse={false}
          />
        )}
        {(tab === "pr_ese" || tab === "assignment") && (
          <PrAssessmentTab
            state={state}
            setState={setState}
            showToast={showToast}
            onExportExcel={() => exportWorkbook(state, showToast)}
            dataKey="assignment"
            title="PR_ESE — Practical End Semester Examination"
            titleShort="PR_ESE"
            subtitle="External Practical Examination & Viva Assessment (Matching Reference Image)"
            cohortCountsDefault={[137, 137, 137, 137, 137, 0]}
            cohortTotalDefault={137}
            isEse={true}
          />
        )}
        {tab === "internal2" && (
          <AssessmentTab title="TH_PA — Internal Mid-Sem Exam" note="Each question tests one CO (Q1\u2192CO1 \u2026 Q6\u2192CO6). Theory Weight = 30%."
            data={state.internal2} nQ={6} coForQ={(i) => COs[i]} stats={i2Stats} targetPctCO={state.targets.targetPctCO}
            onChange={(d) => setState((s) => ({ ...s, internal2: d }))} />
        )}
        {tab === "endsem" && (
          <AssessmentTab title="TH_ESE — GTU End Semester Exam" note="Q1\u2013Q6 = Part A, Q7\u2013Q12 = Part B. Qi and Q(i+6) both test COi. Theory Weight = 70%."
            data={state.endsem} nQ={12} coForQ={(i) => COs[i % 6]} stats={esStats} targetPctCO={state.targets.targetPctCO}
            onChange={(d) => setState((s) => ({ ...s, endsem: d }))} />
        )}
        {tab === "survey" && <SurveyTab survey={state.survey} avg={indirect} onChange={(sv) => setState((s) => ({ ...s, survey: sv }))} />}
        {tab === "results" && (
          <ResultsTab
            state={state}
            direct={direct}
            indirect={indirect}
            final={final}
            poAttain={poAttain}
            thPaStats={i2Stats}
            thEseStats={esStats}
            prPaStats={i1Stats}
            prEseStats={asgStats}
            theoryBlended={theoryBlended}
            practicalBlended={practicalBlended}
            courseAttainment={courseAttainment}
          />
        )}
      </main>

      {showProfileModal && (
        <UserProfileModal
          user={user}
          onClose={() => setShowProfileModal(false)}
          onUserUpdate={onUserUpdate}
          showToast={showToast}
        />
      )}

      {toast && <div className="toast hide-on-print">{toast}</div>}
    </div>
  );
}

/* ============================================================================
   ROOT: AUTH GATE
============================================================================ */
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }
  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }
  componentDidCatch(error, errorInfo) {
    console.error("ErrorBoundary caught error:", error, errorInfo);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: 40, textAlign: "center", fontFamily: "sans-serif" }}>
          <h2 style={{ color: "#991b1b" }}>Something went wrong</h2>
          <p style={{ color: "#475569" }}>An unexpected error occurred while rendering the page.</p>
          <pre style={{ color: "#b91c1c", background: "#fef2f2", padding: 16, borderRadius: 8, display: "inline-block", maxWidth: 800, textAlign: "left", whiteSpace: "pre-wrap", fontSize: 13 }}>
            {this.state.error?.toString()}
          </pre>
          <div style={{ marginTop: 20 }}>
            <button
              style={{ background: "#4338ca", color: "#fff", border: "none", padding: "10px 20px", borderRadius: 8, cursor: "pointer", fontWeight: 700 }}
              onClick={() => { localStorage.clear(); window.location.reload(); }}
            >
              Reset Session &amp; Reload
            </button>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}

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
  function handleUserUpdate(u) { setUser(u); }
  function handleLogout() { setToken(null); setUser(null); }

  return (
    <ErrorBoundary>
      <div className="root">
        <style>{CSS}</style>
        {checking ? (
          <div className="loading-screen">Loading…</div>
        ) : user ? (
          <Calculator user={user} onLogout={handleLogout} onUserUpdate={handleUserUpdate} />
        ) : (
          <LoginPage onAuthed={handleAuthed} />
        )}
      </div>
    </ErrorBoundary>
  );
}

/* ============================================================================
   STYLES
============================================================================ */
const CSS = `
  .root { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif; color: #1f2333; background: #f4f5fb; min-height: 100vh; }
  * { box-sizing: border-box; }
  input, textarea, select, button { font-family: inherit; }

  /* ==========================================================================
     ENHANCED PRESTIGIOUS INSTITUTIONAL LOGIN PAGE
  ========================================================================== */
  .login-page-bg {
    min-height: 100vh;
    display: flex;
    justify-content: center;
    align-items: center;
    padding: 36px 16px;
    background: radial-gradient(circle at 50% 8%, #e0e7ff 0%, #eef2ff 35%, #f8fafc 100%);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  }
  .login-wrapper {
    width: 100%;
    max-width: 480px;
    display: flex;
    flex-direction: column;
    align-items: center;
  }
  .login-branding-header {
    display: flex;
    flex-direction: column;
    align-items: center;
    text-align: center;
    margin-bottom: 16px;
  }
  .login-logo-container {
    width: 96px;
    height: 96px;
    border-radius: 50%;
    background: #ffffff;
    padding: 6px;
    box-shadow: 0 10px 25px rgba(67, 56, 202, 0.15), 0 0 0 1px rgba(199, 210, 254, 0.6);
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 12px;
    transition: transform 0.2s ease;
  }
  .login-logo-container:hover {
    transform: scale(1.03);
  }
  .login-gtu-img {
    width: 100%;
    height: 100%;
    object-fit: contain;
    border-radius: 50%;
  }
  .login-affiliation-pill {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: #e0e7ff;
    color: #3730a3;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.6px;
    padding: 4px 12px;
    border-radius: 20px;
    border: 1px solid #c7d2fe;
    margin-bottom: 8px;
  }
  .login-main-title {
    font-size: 27px;
    font-weight: 900;
    color: #1e1b4b;
    letter-spacing: 0.8px;
    margin: 2px 0 3px;
    text-shadow: 0 1px 2px rgba(0,0,0,0.05);
  }
  .login-dept-title {
    font-size: 13.5px;
    font-weight: 700;
    color: #4338ca;
    letter-spacing: 0.5px;
    margin-bottom: 4px;
  }
  .login-system-title {
    font-size: 12px;
    color: #64748b;
    margin: 0;
    font-weight: 500;
  }

  .login-roles-bar {
    width: 100%;
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    padding: 8px 14px;
    border-radius: 10px;
    margin-bottom: 14px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.03);
  }
  .login-roles-label {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 11.5px;
    font-weight: 700;
    color: #334155;
  }
  .login-roles-badges {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .role-pill {
    font-size: 10.5px;
    font-weight: 700;
    padding: 2px 8px;
    border-radius: 12px;
  }
  .role-pill.faculty { background: #dbeafe; color: #1e40af; }
  .role-pill.hod { background: #fef3c7; color: #92400e; }
  .role-pill.admin { background: #dcfce7; color: #166534; }

  .login-main-card {
    width: 100%;
    background: #ffffff;
    border-radius: 16px;
    padding: 26px 24px;
    border: 1.5px solid #e2e8f0;
    box-shadow: 0 12px 36px rgba(30, 27, 75, 0.08), 0 2px 8px rgba(0,0,0,0.04);
  }
  .login-card-head {
    margin-bottom: 18px;
  }
  .login-card-title-row {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .login-card-title-row h2 {
    font-size: 20px;
    font-weight: 800;
    color: #1e1b4b;
    margin: 0;
  }
  .login-card-desc {
    font-size: 12.5px;
    color: #64748b;
    margin: 3px 0 0;
  }
  .login-error-banner {
    display: flex;
    align-items: center;
    gap: 8px;
    background: #fef2f2;
    color: #991b1b;
    border: 1px solid #fecaca;
    padding: 10px 12px;
    border-radius: 8px;
    font-size: 12px;
    font-weight: 600;
    margin-bottom: 16px;
  }
  .login-form {
    display: flex;
    flex-direction: column;
    gap: 14px;
  }
  .login-input-group {
    display: flex;
    flex-direction: column;
    gap: 5px;
  }
  .login-input-label {
    font-size: 12.5px;
    font-weight: 700;
    color: #334155;
  }
  .login-input-box {
    display: flex;
    align-items: center;
    gap: 10px;
    border: 1.5px solid #cbd5e1;
    border-radius: 9px;
    padding: 9px 12px;
    background: #f8fafc;
    transition: all 0.15s ease;
  }
  .login-input-box:focus-within {
    border-color: #4338ca;
    background: #ffffff;
    box-shadow: 0 0 0 3px rgba(67, 56, 202, 0.12);
  }
  .login-input-icon {
    color: #64748b;
    flex-shrink: 0;
  }
  .login-input-box input {
    border: none;
    outline: none;
    background: transparent;
    width: 100%;
    font-size: 13.5px;
    font-weight: 600;
    color: #0f172a;
  }
  .login-pw-toggle {
    background: none;
    border: none;
    color: #64748b;
    cursor: pointer;
    padding: 2px;
    display: flex;
    align-items: center;
    transition: color 0.15s;
  }
  .login-pw-toggle:hover {
    color: #1e1b4b;
  }
  .login-options-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-top: -2px;
  }
  .login-remember-label {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 12.5px;
    font-weight: 600;
    color: #475569;
    cursor: pointer;
  }
  .login-submit-btn {
    background: linear-gradient(135deg, #3730a3 0%, #4f46e5 100%);
    color: #ffffff;
    border: none;
    border-radius: 9px;
    padding: 11px 18px;
    font-size: 14px;
    font-weight: 800;
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    cursor: pointer;
    box-shadow: 0 4px 12px rgba(67, 56, 202, 0.25);
    transition: all 0.15s ease;
    margin-top: 4px;
  }
  .login-submit-btn:hover:not(:disabled) {
    transform: translateY(-1px);
    box-shadow: 0 6px 18px rgba(67, 56, 202, 0.35);
  }
  .login-submit-btn:disabled {
    opacity: 0.7;
    cursor: not-allowed;
  }
  .login-spin {
    animation: login-spin-kf 1s linear infinite;
  }
  @keyframes login-spin-kf {
    from { transform: rotate(0deg); }
    to { transform: rotate(360deg); }
  }

  .login-quickfill-box {
    margin-top: 16px;
    padding-top: 14px;
    border-top: 1px dashed #e2e8f0;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .login-quickfill-title {
    display: flex;
    align-items: center;
    gap: 6px;
    font-size: 11.5px;
    font-weight: 700;
    color: #64748b;
  }
  .login-quickfill-btn {
    background: #f1f5f9;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 12px;
    color: #334155;
    cursor: pointer;
    text-align: center;
    transition: all 0.15s;
  }
  .login-quickfill-btn:hover {
    background: #e2e8f0;
    border-color: #94a3b8;
  }

  .login-accreditation-footer {
    width: 100%;
    margin-top: 18px;
    text-align: center;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 4px;
  }
  .footer-line-1 {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    font-size: 11.5px;
    font-weight: 700;
    color: #3730a3;
  }
  .footer-line-2 {
    font-size: 11px;
    color: #64748b;
    font-weight: 500;
    line-height: 1.4;
  }
  .footer-line-3 {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    font-size: 10.5px;
    color: #059669;
    font-weight: 700;
    margin-top: 2px;
  }

  .loading-screen { min-height:100vh; display:flex; align-items:center; justify-content:center; color:#6b6f8f; font-size:15px; }
  .app-shell { min-height: 100vh; display:flex; flex-direction:column; }
  .app-header { display:flex; align-items:center; justify-content:space-between; padding: 14px 22px; background:#fff; border-bottom: 1px solid #e9eaf4; flex-wrap:wrap; gap:10px; position: sticky; top:0; z-index: 10; }
  .app-header-left { display:flex; align-items:center; gap:12px; }
  .brand-mini { width: 40px; height: 40px; border-radius: 10px; background: #ffffff; border: 1.5px solid #c7d2fe; padding: 2px; display: flex; align-items: center; justify-content: center; overflow: hidden; box-shadow: 0 2px 6px rgba(67, 56, 202, 0.12); flex-shrink: 0; }
  .app-title { font-weight:800; font-size:15px; color:#14162a; }
  .app-subtitle { font-size:12px; color:#8b8fb0; margin-top:1px; }
  .app-header-right { display:flex; align-items:center; gap:8px; flex-wrap:wrap; position:relative; }
  .user-chip { display:flex; align-items:center; gap:6px; background:#f1f2fb; color:#3d3f63; padding:7px 12px; border-radius:999px; font-size:12.5px; font-weight:600; }
  .user-chip-btn { display:inline-flex; align-items:center; gap:6px; background:#eef2ff; color:#3730a3; border:1.5px solid #c7d2fe; padding:7px 14px; border-radius:999px; font-size:12.5px; font-weight:700; cursor:pointer; transition:all 0.15s ease; box-shadow:0 1px 3px rgba(67,56,202,0.08); }
  .user-chip-btn:hover { background:#e0e7ff; border-color:#818cf8; transform:translateY(-1px); box-shadow:0 3px 8px rgba(67,56,202,0.16); }

  .btn-primary { background: linear-gradient(135deg, #4338ca, #6366f1); color:#fff; border:none; border-radius: 9px; padding: 8px 14px; font-weight:700; font-size:12.5px; display:inline-flex; align-items:center; justify-content:center; gap:6px; cursor:pointer; box-shadow:0 2px 6px rgba(67,56,202,0.25); transition:all 0.15s ease; }
  .btn-primary:hover { filter: brightness(1.08); transform: translateY(-1px); box-shadow: 0 4px 12px rgba(67,56,202,0.35); }
  .btn-primary:disabled { opacity: 0.6; cursor: not-allowed; transform: none; }
  .btn-block { width: 100%; }

  .btn-ghost { display:inline-flex; align-items:center; gap:6px; background:#f8fafc; border:1px solid #e2e8f0; color:#334155; border-radius:9px; padding:7px 12px; font-size:12.5px; font-weight:600; cursor:pointer; transition:all 0.15s ease; }
  .btn-ghost:hover { background:#f1f5f9; border-color:#cbd5e1; color:#0f172a; }
  .btn-ghost.danger { color:#b91c1c; border-color:#fecaca; background:#fff5f5; }
  .btn-ghost.danger:hover { background:#fee2e2; border-color:#f87171; }
  .btn-active-edit { background:#e0e7ff; color:#3730a3; border-color:#818cf8; }

  .icon-btn { background:none; border:none; cursor:pointer; color:#64748b; padding:6px; border-radius:8px; display:inline-flex; align-items:center; justify-content:center; transition:all 0.15s ease; }
  .icon-btn:hover { background:#f1f5f9; color:#1e293b; }
  .icon-btn.danger:hover { background:#fee2e2; color:#b91c1c; }

  .modal-backdrop { position:fixed; top:0; left:0; right:0; bottom:0; background:rgba(15,23,42,0.6); backdrop-filter:blur(4px); display:flex; align-items:center; justify-content:center; z-index:1000; padding:16px; }
  .modal-card.user-profile-modal { background:#ffffff; border-radius:18px; width:100%; max-width:470px; padding:26px 24px; box-shadow:0 20px 45px rgba(15,23,42,0.22); border:1px solid #e2e8f0; max-height:92vh; overflow-y:auto; }
  .modal-header { display:flex; align-items:center; justify-content:space-between; margin-bottom:16px; padding-bottom:12px; border-bottom:1px solid #f1f5f9; }
  .user-profile-avatar { width:38px; height:38px; border-radius:50%; background:#e0e7ff; display:flex; align-items:center; justify-content:center; flex-shrink:0; }
  .profile-role-box { background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:10px 14px; margin-bottom:14px; }
  .profile-password-section { background:#f8faff; border:1.5px solid #e0e7ff; border-radius:10px; padding:12px 14px; margin-top:2px; }

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

  .app-main { flex:1; padding: 22px; max-width: 1220px; margin: 0 auto; width:100%; }

  .panel-head { margin: 4px 0 12px; }
  .panel-head h3 { font-size:15px; margin:0 0 3px; color:#14162a; }
  .panel-head-row { display:flex; align-items:flex-start; justify-content:space-between; gap:12px; }
  .muted { color:#8b8fb0; font-size:12.5px; margin:0; }

  .table-scroll { overflow-x:auto; border:1px solid #eceefa; border-radius:14px; }
  .grid-table { border-collapse:collapse; width:100%; font-size:12.5px; }
  .grid-table th, .grid-table td { padding:8px 9px; text-align:center; border-bottom:1px solid #f1f2fa; white-space:nowrap; }
  .grid-table thead th { background:#f7f8fd; color:#4b4f6b; font-weight:700; font-size:11.5px; }
  .th-sub { font-size:10px; font-weight:500; color:#a5a9c9; }
  .max-row td { background:#fbfbfe; color:#8b8fb0; font-size:11.5px; }
  .target-row td { background:#fbf7fd; color:#a21caf; font-size:11.5px; font-weight:600; text-align:center; padding:5px 4px; }
  .target-row .th-sub { color:#c084dc; font-weight:500; }
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

  /* ==========================================================================
     EVALUATION PLAN & TARGET SETTING SPECIFIC STYLES (EXACT MATCH TO INSTITUTIONAL IMAGES)
  ========================================================================== */
  .eval-plan-wrapper { display: flex; flex-direction: column; gap: 16px; }
  .eval-toolbar { display: flex; justify-content: space-between; align-items: center; background: #fff; border: 1px solid #eceefa; border-radius: 14px; padding: 14px 20px; flex-wrap: wrap; gap: 12px; }
  .eval-toolbar-left { display: flex; align-items: center; gap: 12px; }
  .eval-toolbar-icon { color: #1f497d; }
  .eval-toolbar-title { font-size: 16px; font-weight: 800; color: #14162a; margin: 0; }
  .eval-toolbar-sub { font-size: 12px; color: #8b8fb0; margin: 2px 0 0; }
  .eval-toolbar-actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }

  .eval-sheet { background: #fff; border: 1px solid #000; padding: 24px 20px; font-family: "Segoe UI", Calibri, Arial, sans-serif; }
  .eval-doc-header { text-align: center; margin-bottom: 18px; }
  .eval-inst-title { font-size: 24px; font-weight: 900; color: #000; letter-spacing: 0.5px; margin: 0 0 4px; text-transform: uppercase; font-family: "Segoe UI", Arial, sans-serif; }
  .eval-dept-title { font-size: 16px; font-weight: 700; color: #000; margin: 0 0 4px; letter-spacing: 0.3px; text-transform: uppercase; }
  .eval-plan-title { font-size: 14px; font-weight: 700; color: #000; margin: 0; letter-spacing: 0.5px; text-transform: uppercase; }

  .eval-section-block { margin-top: 18px; }
  .eval-section-badge { display: inline-block; background: #2f5597 !important; color: #fff !important; font-weight: 800; font-size: 13px; padding: 4px 16px; border-radius: 2px; margin-bottom: 4px; letter-spacing: 0.3px; }

  .eval-table-wrap { overflow-x: auto; margin-top: 0; }
  .eval-table { border-collapse: collapse; width: 100%; border: 1.5px solid #000; font-size: 12.5px; font-family: "Segoe UI", Calibri, Arial, sans-serif; }
  .eval-table th, .eval-table td { border: 1px solid #000; padding: 5px 6px; text-align: center; color: #000; }
  .eval-table th { font-weight: 700; }

  /* Cell colors from institutional images */
  .eval-cell-peach { background: #fce4d6 !important; font-weight: 700; }
  .eval-cell-green { background: #e2efda !important; }
  .eval-cell-cyan { background: #bdd7ee !important; }
  .eval-cell-amber { background: #ffc000 !important; font-weight: 800; }
  .eval-cell-grey { background: #bfbfbf !important; font-weight: 700; }
  .eval-cell-lavender { background: #e7e6f4 !important; font-weight: 700; }
  .eval-cell-white { background: #ffffff !important; }
  .eval-cell-blue-header { background: #bdd7ee !important; font-weight: 700; }

  .eval-no-border { border: none !important; }
  .eval-note-red { color: #ff0000 !important; font-weight: 600; font-size: 12px; text-align: left; padding-left: 10px; border: none !important; }

  /* Table Specifics */
  .eval-meta-table { border: 1.5px solid #000; }
  .eval-meta-table td { text-align: left; }
  .eval-meta-label { background: #fff; font-weight: 800; width: 160px; text-align: right; border: 1px solid #000; }
  .eval-meta-val { min-width: 140px; background: #e2efda !important; border: 1px solid #000; }

  .eval-co-id { font-weight: 800; width: 90px; }
  .eval-co-subcode { font-weight: 600; width: 110px; }
  .eval-co-statement { text-align: left !important; padding-left: 10px; }

  .eval-final-target-row { font-size: 13.5px; font-weight: 800; padding: 7px 12px; letter-spacing: 0.5px; background: #bfbfbf !important; text-align: center; }

  .eval-inline-input { border: 1px solid #2f5597; border-radius: 3px; padding: 3px 6px; font-family: inherit; font-size: inherit; font-weight: inherit; text-align: center; }
  .eval-title-input { width: 360px; font-size: 20px; font-weight: 900; }
  .eval-dept-input { width: 340px; font-size: 15px; font-weight: 700; }
  .eval-cell-input { width: 100%; border: 1px solid #2f5597; background: #fff; border-radius: 3px; padding: 2px 4px; text-align: center; font-size: 12px; font-weight: 600; }
  .eval-input-wide { text-align: left; }
  .eval-select-cell { width: 100%; border: 1px solid #2f5597; background: #fff; border-radius: 3px; padding: 2px 4px; text-align: center; font-size: 12px; }

  /* Target Setting Layout */
  .ts-header-grid { display: flex; flex-wrap: wrap; gap: 16px; align-items: center; padding: 14px 16px; background: #fff; border: 1.5px solid #000; margin-bottom: 14px; }
  .ts-header-item { display: flex; align-items: center; gap: 6px; font-size: 13px; font-weight: 700; }
  .ts-header-label { white-space: nowrap; }
  .ts-header-val-yellow { background: #ffff00 !important; font-weight: 800; border: 1px solid #000; padding: 3px 8px; font-size: 13px; }
  .ts-header-val { background: #fff; font-weight: 700; border: 1px solid #000; padding: 3px 8px; font-size: 13px; }
  .ts-val-wide { min-width: 240px; }
  .ts-year-input { width: 100%; background: transparent; border: none; font-weight: 700; text-align: center; font-size: 11.5px; }

  .ts-total-row td { background: #fff; }
  .ts-two-col-grid { display: grid; grid-template-columns: 1fr; gap: 20px; }
  @media (min-width: 900px) {
    .ts-two-col-grid { grid-template-columns: 1.25fr 1fr; }
  }

  /* Print Styles */
  @media print {
    body { background: #fff; }
    .hide-on-print { display: none !important; }
    .app-main { padding: 0 !important; max-width: 100% !important; }
    .eval-sheet { border: none !important; box-shadow: none !important; padding: 0 !important; }
    .eval-table { page-break-inside: avoid; }
    .eval-section-block { page-break-inside: avoid; }
  }

  /* ==========================================================================
     PR_PA TAB SPECIFIC STYLES (EXACT MATCH TO INSTITUTIONAL REFERENCE IMAGE)
  ========================================================================== */
  .prpa-container { max-width: 100%; }
  .prpa-sheet { background: #fff; border: 1.5px solid #000; padding: 18px 16px; font-family: "Segoe UI", Calibri, Arial, sans-serif; }
  .prpa-top-grid { display: grid; grid-template-columns: 1fr; gap: 16px; margin-bottom: 16px; }
  @media (min-width: 1000px) {
    .prpa-top-grid { grid-template-columns: 1.1fr 1fr; }
  }
  .prpa-card-box { overflow-x: auto; }
  .prpa-info-table { width: 100%; border-collapse: collapse; border: 1.5px solid #000; font-size: 12.5px; }
  .prpa-info-table td { border: 1px solid #000; padding: 5px 8px; }
  .prpa-lbl { font-weight: 700; text-align: right; background: #fff; white-space: nowrap; }
  .prpa-val-bold { font-weight: 700; text-align: left; background: #fff; }
  .prpa-val-name { color: #111827; }
  .prpa-val { text-align: left; background: #fff; }
  .prpa-target-lbl-merged { font-weight: 800; text-align: center; background: #fce4d6 !important; letter-spacing: 0.5px; font-size: 12.5px; }
  .prpa-target-val-box { background: #fce4d6 !important; text-align: center; }
  .prpa-target-input { width: 70px; border: 1px solid #000; font-weight: 800; font-size: 13px; text-align: center; padding: 3px 4px; background: #fff; border-radius: 2px; }

  /* Top Right Summary Table */
  .prpa-achieving-table { width: 100%; border-collapse: collapse; border: 1.5px solid #000; font-size: 12px; }
  .prpa-achieving-table th, .prpa-achieving-table td { border: 1px solid #000; padding: 5px 4px; text-align: center; }
  .prpa-achieving-title { background: #bdd7ee !important; font-weight: 800; font-size: 12px; }
  .prpa-metric-name { background: #bdd7ee !important; font-weight: 700; text-align: center; white-space: nowrap; font-size: 11.5px; }
  .prpa-stat-num { background: #bdd7ee !important; font-weight: 700; font-size: 12px; color: #000; }
  .prpa-level-highlight { font-weight: 800; }

  /* Filter Bar */
  .prpa-filter-bar { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 12px; flex-wrap: wrap; background: #f8fafc; padding: 8px 12px; border-radius: 8px; border: 1px solid #e2e8f0; }
  .prpa-search-input { border: 1px solid #cbd5e1; border-radius: 6px; padding: 6px 12px; font-size: 13px; min-width: 260px; outline: none; background: #fff; }
  .prpa-search-input:focus { border-color: #4338ca; box-shadow: 0 0 0 2px rgba(67,56,202,0.15); }
  .prpa-student-count-tag { font-size: 12px; color: #64748b; font-weight: 600; }

  /* Main Grid Table */
  .prpa-scroll-table { overflow-x: auto; max-height: 540px; border: 1.5px solid #000; }
  .prpa-grid-table { border-collapse: collapse; width: 100%; font-size: 12px; font-family: "Segoe UI", Calibri, Arial, sans-serif; }
  .prpa-grid-table th, .prpa-grid-table td { border: 1px solid #000; padding: 4px 5px; text-align: center; }
  .prpa-grid-table thead th { position: sticky; top: 0; z-index: 3; }
  .prpa-super-left { background: #bfbfbf !important; font-weight: 800; letter-spacing: 0.3px; }
  .prpa-col-co { background: #fce4d6 !important; font-weight: 700; }
  .prpa-col-tot { background: #fce4d6 !important; font-weight: 800; }
  .prpa-max-mark-th { background: #ffffff !important; padding: 2px !important; }
  .prpa-max-input { width: 36px; padding: 2px; text-align: center; border: 1px solid #94a3b8; font-weight: 700; font-size: 11.5px; border-radius: 2px; }
  .prpa-total-max-hdr { font-weight: 800; font-size: 12.5px; background: #fff !important; }
  .prpa-attained-title { font-weight: 800; font-size: 12.5px; background: #fff !important; }

  /* Student Rows */
  .prpa-roll-cell { width: 115px; font-size: 12px; font-weight: 600; text-align: center; border: 1px solid transparent; background: transparent; }
  .prpa-roll-cell:focus { border-color: #4338ca; background: #fff; }
  .prpa-name-cell { width: 100%; font-size: 12px; font-weight: 600; text-align: left; text-transform: uppercase; border: 1px solid transparent; background: transparent; }
  .prpa-name-cell:focus { border-color: #4338ca; background: #fff; }
  .prpa-mark-cell { width: 38px; font-size: 12px; font-weight: 600; text-align: center; border: 1px solid #e2e8f0; border-radius: 2px; padding: 2px; }
  .prpa-mark-cell:focus { border-color: #4338ca; background: #fff; }
  .prpa-tot-mark-cell { font-weight: 800; color: #000; background: #f8fafc; font-size: 12.5px; }
  .prpa-pct-cell { font-weight: 600; color: #1e293b; font-size: 11.5px; }
  .prpa-tot-pct-cell { font-weight: 800; background: #f8fafc; color: #000; }
  
  .prpa-flag-y { font-weight: 800; color: #15803d; background: #f0fdf4; font-size: 12px; }
  .prpa-flag-n { font-weight: 800; color: #b91c1c; background: #fef2f2; font-size: 12px; }
  .prpa-flag-empty { color: #94a3b8; }
  .prpa-flag-tot { font-size: 12.5px; }

  /* Bottom Details Table */
  .prpa-end-section { margin-top: 24px; }
  .prpa-end-table { width: 100%; border-collapse: collapse; border: 1.5px solid #000; font-size: 12.5px; }
  .prpa-end-table th { background: #f1f5f9; font-weight: 700; }
  .prpa-end-table td { padding: 6px 8px; }

  .prpa-matrix-legend-box { margin-top: 16px; padding: 12px 14px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; }
  .prpa-matrix-legend-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 10px; margin-top: 8px; }
  .prpa-matrix-card { display: flex; align-items: center; gap: 8px; background: #fff; border: 1px solid #e2e8f0; padding: 8px 12px; border-radius: 6px; }
  .prpa-matrix-lvl { font-weight: 800; font-size: 12px; padding: 3px 8px; border-radius: 4px; white-space: nowrap; }
  .prpa-matrix-text { font-size: 11.5px; color: #475569; font-weight: 500; }

  /* Institution & Department Metadata Box */
  .prpa-inst-box { background: #f8fafc; border: 1.5px solid #c7d2fe; border-radius: 8px; padding: 12px 16px; margin-bottom: 14px; }
  .prpa-inst-header { display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 8px; margin-bottom: 10px; }
  .prpa-inst-badge { font-size: 11.5px; color: #4338ca; background: #e0e7ff; padding: 3px 8px; border-radius: 4px; font-weight: 600; }
  .prpa-inst-grid { display: grid; grid-template-columns: 1fr; gap: 14px; }
  @media (min-width: 768px) {
    .prpa-inst-grid { grid-template-columns: 1fr 1fr; }
  }
  .prpa-inst-field { display: flex; flex-direction: column; gap: 4px; }
  .prpa-inst-label { font-size: 12px; font-weight: 700; color: #334155; }
  .prpa-inst-input { border: 1px solid #cbd5e1; border-radius: 6px; padding: 7px 10px; font-size: 13px; font-weight: 600; color: #1e293b; background: #fff; }
  .prpa-inst-input:focus { border-color: #4338ca; outline: none; box-shadow: 0 0 0 2px rgba(67,56,202,0.15); }

  /* Interactive Excel Formula Bar (fx) */
  .prpa-fx-bar { display: flex; align-items: stretch; background: #ffffff; border: 1.5px solid #2f5597; border-radius: 6px; margin-bottom: 14px; overflow: hidden; box-shadow: 0 2px 6px rgba(0,0,0,0.06); }
  .prpa-fx-left { display: flex; align-items: center; background: #f1f5f9; border-right: 1.5px solid #cbd5e1; flex-shrink: 0; }
  .prpa-fx-name-box { padding: 6px 12px; font-family: "Cascadia Code", Consolas, "Courier New", monospace; font-size: 13px; font-weight: 800; color: #1e293b; min-width: 68px; text-align: center; border-right: 1px solid #cbd5e1; }
  .prpa-fx-icon { padding: 6px 10px; font-weight: 800; font-size: 14px; color: #2563eb; display: flex; align-items: center; gap: 1px; user-select: none; }
  .prpa-fx-icon em { font-style: italic; font-family: serif; font-size: 15px; }
  .prpa-fx-content { display: flex; flex-direction: column; justify-content: center; flex: 1; padding: 5px 12px; background: #fafafa; overflow-x: auto; gap: 3px; }
  .prpa-fx-formula { font-family: "Cascadia Code", Consolas, "Courier New", monospace; font-size: 13px; color: #0972d3; font-weight: 700; white-space: nowrap; }
  .prpa-fx-formula code { background: none; color: inherit; font-size: inherit; }
  .prpa-fx-math { font-size: 12px; color: #334155; display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
  .prpa-fx-pill { font-size: 11.5px; color: #4338ca; font-weight: 700; }

  /* Cell Interaction */
  .prpa-clickable-cell { cursor: pointer; transition: background 0.15s, outline 0.15s; }
  .prpa-clickable-cell:hover { filter: brightness(0.96); }
  .prpa-cell-active { outline: 2px solid #2563eb !important; outline-offset: -2px; z-index: 2; position: relative; box-shadow: inset 0 0 0 1px #2563eb; }

  /* Results KPI Cards */
  .results-kpi-card { background: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; padding: 14px 20px; display: flex; flex-direction: column; justify-content: center; min-width: 200px; box-shadow: 0 2px 6px rgba(0,0,0,0.04); }
  .results-kpi-title { font-size: 12px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
  .results-kpi-val { font-size: 26px; font-weight: 900; line-height: 1.1; margin-bottom: 6px; }
  .results-kpi-sub { font-size: 11.5px; color: #64748b; font-weight: 500; }

  /* Toast Notification */
  .toast { position: fixed; bottom: 24px; right: 24px; background: #1e1b4b; color: #ffffff; padding: 12px 20px; border-radius: 10px; font-size: 13px; font-weight: 700; box-shadow: 0 10px 30px rgba(0, 0, 0, 0.25); z-index: 99999; pointer-events: none; animation: toast-in 0.2s ease-out; }
  @keyframes toast-in {
    from { opacity: 0; transform: translateY(12px); }
    to { opacity: 1; transform: translateY(0); }
  }
`;
