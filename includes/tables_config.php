<?php
/**
 * Single source of truth for every activity table.
 * Every page (dashboard, add, edit, history) reads this config
 * instead of hard-coding fields per table.
 */

const PO_PSO = ['po1','po2','po3','po4','po5','po6','po7','po8','po9','po10','po11','pso1','pso2'];
const CO_LIST = ['co1','co2','co3','co4','co5','co6'];

// Tables whose PO/PSO values feed the combined attainment average on the dashboard.
const PO_TABLES = ['program_exit_survey', 'expert_lecture', 'industry_visit', 'alumni_survey', 'industry_survey'];

$TABLES = [

    'course_exit_survey' => [
        'label' => 'Course Exit Survey',
        'desc'  => 'Record CO-mapped exit survey results for each course offering.',
        'fields' => [
            ['key' => 'acad_year',     'label' => 'Academic Year',      'type' => 'text', 'required' => true, 'ph' => 'e.g. 2025-26'],
            ['key' => 'sem',           'label' => 'Semester',           'type' => 'text', 'ph' => 'e.g. 5'],
            ['key' => 'c_name',        'label' => 'Course Name',        'type' => 'text', 'ph' => 'e.g. Data Structures'],
            ['key' => 'c_code',        'label' => 'Course Code',        'type' => 'text', 'ph' => 'e.g. 4331602'],
            ['key' => 'c_coordinator', 'label' => 'Course Coordinator', 'type' => 'text', 'ph' => 'Faculty name'],
        ],
        'outcomes'      => CO_LIST,
        'outcome_label' => 'CO',
    ],

    'program_exit_survey' => [
        'label' => 'Program Exit Survey',
        'desc'  => 'Record mapped PO/PSO levels and retain an auditable activity history.',
        'fields' => [
            ['key' => 'acad_year', 'label' => 'Academic Year', 'type' => 'text', 'required' => true, 'ph' => 'e.g. 2025-26'],
        ],
        'outcomes'      => PO_PSO,
        'outcome_label' => null,
    ],

    'expert_lecture' => [
        'label' => 'Expert Lecture',
        'desc'  => 'Record mapped PO/PSO levels and retain an auditable activity history.',
        'fields' => [
            ['key' => 'acad_year',   'label' => 'Academic Year', 'type' => 'text', 'required' => true, 'ph' => 'e.g. 2025-26'],
            ['key' => 'date',        'label' => 'Date',          'type' => 'date'],
            ['key' => 'sem',         'label' => 'Semester',      'type' => 'text', 'ph' => 'e.g. 5'],
            ['key' => 'title',       'label' => 'Title',         'type' => 'text', 'ph' => 'Lecture topic'],
            ['key' => 'expert_name', 'label' => 'Expert Name',   'type' => 'text', 'ph' => 'Speaker name'],
        ],
        'outcomes'      => PO_PSO,
        'outcome_label' => null,
    ],

    'industry_visit' => [
        'label' => 'Industry Visit',
        'desc'  => 'Record mapped PO/PSO levels and retain an auditable activity history.',
        'fields' => [
            ['key' => 'acad_year', 'label' => 'Academic Year', 'type' => 'text', 'required' => true, 'ph' => 'e.g. 2025-26'],
            ['key' => 'date',      'label' => 'Date',          'type' => 'date'],
            ['key' => 'sem',       'label' => 'Semester',      'type' => 'text', 'ph' => 'e.g. 5'],
            ['key' => 'title',     'label' => 'Title',         'type' => 'text', 'ph' => 'Visit title'],
            ['key' => 'place',     'label' => 'Place',         'type' => 'text', 'ph' => 'Location / company'],
        ],
        'outcomes'      => PO_PSO,
        'outcome_label' => null,
    ],

    'alumni_survey' => [
        'label' => 'Alumni Survey',
        'desc'  => 'Record mapped PO/PSO levels and retain an auditable activity history.',
        'fields' => [
            ['key' => 'acad_year', 'label' => 'Academic Year', 'type' => 'text', 'required' => true, 'ph' => 'e.g. 2025-26'],
            ['key' => 'date',      'label' => 'Date',          'type' => 'date'],
            ['key' => 'title',     'label' => 'Title',         'type' => 'text', 'ph' => 'Survey title'],
        ],
        'outcomes'      => PO_PSO,
        'outcome_label' => null,
    ],

    'industry_survey' => [
        'label' => 'Industry Survey',
        'desc'  => 'Record mapped PO/PSO levels and retain an auditable activity history.',
        'fields' => [
            ['key' => 'acad_year',         'label' => 'Academic Year',      'type' => 'text', 'required' => true, 'ph' => 'e.g. 2025-26'],
            ['key' => 'date',              'label' => 'Date',               'type' => 'date'],
            ['key' => 'industry_details',  'label' => 'Industry Details',   'type' => 'text', 'ph' => 'Company / sector'],
        ],
        'outcomes'      => PO_PSO,
        'outcome_label' => null,
    ],

];
