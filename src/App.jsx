import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Upload, FileText, AlertTriangle, Download, Wand2, ChevronDown, CheckCircle2, CircleDashed, Columns2, Rows3, Plus } from 'lucide-react';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { jsPDF } from 'jspdf';
import PdfJsWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?worker&inline';

const FIELD_SECTIONS = [
  ['studentName', 'Student Name'],
  ['date', 'Date'],
  ['week', 'Week #'],
  ['patientInitialAge', 'Patient Initial/Age'],
  ['diagnosis', 'Patient Diagnosis'],
  ['allergy', 'Allergy'],
  ['neuro', 'Neuro'],
  ['cardio', 'Cardiovascular & IV Infusions'],
  ['respiratory', 'Respiratory'],
  ['vitals', 'Vital Signs'],
  ['gi', 'Gastrointestinal (GI)'],
  ['gu', 'Genitourinary (GU)'],
  ['skin', 'Skin Integrity'],
  ['fluidElectrolytes', 'Fluid & Electrolyte Status'],
  ['pain', 'Pain Management'],
  ['psychosocial', 'Psychosocial / Developmental Status'],
  ['activity', 'Prescribed Activity'],
  ['cultural', 'Cultural / Spiritual'],
  ['educationNeeds', 'Pt/Family Educational Needs'],
  ['safety', 'Specific Safety Need / Precautions'],
  ['diagnosticTests', 'Diagnostic Test or Procedure Scheduled'],
  ['labs', 'Recent Lab Results'],
  ['dischargePlan', 'Discharge Plan'],
  ['nursingDiagnosis', 'Nursing Diagnosis'],
  ['goal', 'Nursing Goal'],
  ['plan', 'Plan of Care'],
  ['intervention', 'Nursing Care Given / Intervention'],
  ['rationale', 'Rationale for Intervention'],
  ['evaluation', 'Evaluation / Outcome'],
  ['reassessment', 'Reassessment'],
  ['theorist', 'Nursing Theorist Applied'],
  ['knowledgeGained', 'Knowledge Gained'],
  ['courseObjectives', 'How Course/Clinical Objectives Were Met'],
];

const SEMESTER_OPTIONS = ['Spring', 'Summer', 'Fall'];
const WEEK_OPTIONS = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10'];
const COURSE_OPTIONS = [
  'NRBS 3510 Prof. Nsg Practice with Adult I Clients',
  'NRBS 4010 PNP with Adult II Clients',
  'NRBS 4110 PNP with Children',
  'NRBS 4310 Prof. Nsg Practice with Psychiatric Clients',
];
const FACULTY_OPTIONS = ['Karen Colombo'];
const LEGACY_FACULTY_NAMES = {
  'colombo, karen': 'Karen Colombo',
};
const CLINICAL_SITE_OPTIONS = [];

const DEFAULT_STATE = {
  ...Object.fromEntries(FIELD_SECTIONS.map(([k]) => [k, ''])),
  studentName: '',
  semesterMeta: '',
  courseMeta: '',
  facultyMeta: 'Karen Colombo',
  siteMeta: '',
};
const EMPTY_MED = { nameClass: '', doseRoute: '', why: '', action: '', implications: '', sideEffects: '' };
const MED_DEFAULT = [structuredClone(EMPTY_MED), structuredClone(EMPTY_MED), structuredClone(EMPTY_MED), structuredClone(EMPTY_MED)];
const MED_TEMPLATE_ROW_CAP = 5;
const MED_CLASS_EXAMPLES = [
  { pattern: /benzodiazep|anxiolytic|sedative|hypnotic/i, names: ['Lorazepam', 'Clonazepam', 'Diazepam'] },
  { pattern: /antipsychotic/i, names: ['Risperidone', 'Olanzapine', 'Quetiapine'] },
  { pattern: /antidepressant|ssri|snri/i, names: ['Sertraline', 'Escitalopram', 'Fluoxetine'] },
  { pattern: /mood\s*stabili/i, names: ['Lithium', 'Valproate', 'Lamotrigine'] },
  { pattern: /antibacterial|antibiotic/i, names: ['Penicillin G', 'Ceftriaxone', 'Amoxicillin'] },
  { pattern: /antithrombotic|anticoagulant/i, names: ['Apixaban'] },
  { pattern: /antilipemic|statin/i, names: ['Atorvastatin'] },
  { pattern: /beta blocker/i, names: ['Metoprolol', 'Carvedilol'] },
  { pattern: /alpha blocker/i, names: ['Tamsulosin'] },
  { pattern: /proton-pump|antiulcer|acid suppress/i, names: ['Pantoprazole'] },
  { pattern: /insulin|antidiabetic/i, names: ['Insulin lispro'] },
  { pattern: /vitamin d/i, names: ['Cholecalciferol'] },
  { pattern: /iron preparation|antianemia/i, names: ['Ferrous sulfate'] },
  { pattern: /corticosteroid/i, names: ['Prednisone'] },
  { pattern: /magnesium sulfate/i, names: ['Magnesium sulfate'] },
  { pattern: /misoprostol/i, names: ['Misoprostol'] },
  { pattern: /oxytocin/i, names: ['Oxytocin'] },
];

const MED_KEYS = ['nameClass', 'doseRoute', 'why', 'action', 'implications', 'sideEffects'];
const API_KEY_STORAGE_KEY = 'clinical-worksheet-mvp.openaiKey.v1';
const AI_MAX_CASE_TEXT_CHARS = 24000;
const FALLBACK_NA = 'N/A';
const FALLBACK_NOT_ASSESSED = 'Not assessed in case log.';
const FALLBACK_VERIFY = 'Need to verify in chart/case log.';
const appAssetUrl = (path) => `${import.meta.env.BASE_URL}${path.replace(/^\/+/, '')}`;
const DEFAULT_TEMPLATE_URL = appAssetUrl('templates/nursing-process-template.psych.docx');
const DEFAULT_TEMPLATE_NAME = 'Nursing Process Worksheet with Drug Log (Psych)';
const DEFAULT_TEMPLATE_SOURCE_PATH = 'public/templates/nursing-process-template.psych.docx';
const DEFAULT_CONCEPT_MAP_TEMPLATE_URL = appAssetUrl('templates/concept-map-template.docx');
const DEFAULT_CONCEPT_MAP_PDF_TEMPLATE_URL = appAssetUrl('templates/concept-map-template.pdf');
const DEFAULT_CONCEPT_MAP_TEMPLATE_NAME = 'Concept Map - Clinical Document';
const REQUIRED_TEMPLATE_PLACEHOLDERS = ['{studentName}', '{date}', '{diagnosis}', '{nursingDiagnosis}', '{goal}', '{plan}', '{intervention}'];
let pdfJsWorkerPort = null;

function configurePdfJsWorker(pdfjsLib) {
  if (typeof Worker === 'undefined') return;
  if (!pdfJsWorkerPort) pdfJsWorkerPort = new PdfJsWorker();
  pdfjsLib.GlobalWorkerOptions.workerPort = pdfJsWorkerPort;
}

function clean(s = '') {
  return sanitizeTextForDocument(s).replace(/\u00a0/g, ' ').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
}

function isIntentionalNone(value = '') {
  return /^(n\/?a|none|none noted|no known allergies)$/i.test(clean(String(value || '')));
}

function normalizeFacultyName(name = '') {
  const text = clean(name);
  return LEGACY_FACULTY_NAMES[text.toLowerCase()] || text;
}

function sanitizeTextForDocument(text = '') {
  return String(text || '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, ' ')
    .replace(/\uFFFD/g, 'fi');
}

function stripOutputReviewMarkers(value = '') {
  return clean(String(value || ''))
    .replace(/\s*\(?\bAI[-\s]?generated\b\)?\s*[:;-]?\s*/gi, ' ')
    .replace(/\s*\(?\bapp[-\s]?generated\b\)?\s*[:;-]?\s*/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function compactText(value = '', maxLength = 180) {
  const text = stripOutputReviewMarkers(value).replace(/\s+/g, ' ');
  if (!text || text === FALLBACK_VERIFY) return text;
  if (text.length <= maxLength) return text;
  const clipped = text.slice(0, maxLength);
  const lastStop = Math.max(clipped.lastIndexOf('. '), clipped.lastIndexOf('; '), clipped.lastIndexOf(', '));
  const boundary = lastStop > 80 ? lastStop + 1 : clipped.lastIndexOf(' ');
  return clipped.slice(0, boundary > 80 ? boundary : maxLength).trim();
}

function normalizeSpacedPdfText(value = '') {
  return String(value || '')
    .split(/\n/)
    .map((line) => line
      .split(/ {2,}/)
      .map((chunk) => {
        const tokens = chunk.trim().split(/\s+/).filter(Boolean);
        if (tokens.length < 3) return chunk;
        const singleTextTokens = tokens.filter((token) => /^[A-Za-z0-9]$/.test(token)).length;
        const punctuationTokens = tokens.filter((token) => /^[.,:;?!()[\]{}"'&-]$/.test(token)).length;
        const canJoin = singleTextTokens >= 3 && (singleTextTokens + punctuationTokens) / tokens.length >= 0.82;
        return canJoin ? tokens.join('') : chunk;
      })
      .join(' '))
    .join('\n')
    .replace(/&ndash;/gi, '-')
    .replace(/\s+([.,:;?!])/g, '$1')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseLocalDate(inputDate = '') {
  const raw = String(inputDate || '').trim();
  const slash = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (slash) {
    const [, month, day, yearRaw] = slash;
    const year = yearRaw.length === 2 ? `20${yearRaw}` : yearRaw;
    return new Date(Number(year), Number(month) - 1, Number(day));
  }
  const iso = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    const [, year, month, day] = iso;
    return new Date(Number(year), Number(month) - 1, Number(day));
  }
  const parsed = raw ? new Date(raw) : null;
  return parsed && !Number.isNaN(parsed.getTime()) ? parsed : null;
}

function extractAfterLabel(text, label) {
  const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const regex = new RegExp(`${escaped}\\s*([^\\n]+)`, 'i');
  return text.match(regex)?.[1]?.trim() || '';
}

function extractAfterLabelWithStops(text, label, stopLabels = []) {
  const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const escapedStops = stopLabels
    .filter((s) => s && s !== label)
    .map((s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));

  if (!escapedStops.length) return extractAfterLabel(text, label);

  const regex = new RegExp(`${escaped}\\s*([\\s\\S]*?)(?=\\s*(?:${escapedStops.join('|')})\\s*)`, 'i');
  return clean(text.match(regex)?.[1] || '');
}

function getSection(text, startLabel, endLabels = []) {
  const start = text.indexOf(startLabel);
  if (start === -1) return '';
  const from = start + startLabel.length;
  const slice = text.slice(from);
  let end = slice.length;
  for (const label of endLabels) {
    const idx = slice.indexOf(label);
    if (idx !== -1 && idx < end) end = idx;
  }
  return clean(slice.slice(0, end));
}

function inferDiagnosis(chiefComplaint, subjective) {
  if (chiefComplaint && subjective) return `${chiefComplaint}; ${subjective.split('.')[0].trim()}`;
  return chiefComplaint || subjective.split('.')[0]?.trim() || '';
}

function isCaseLogBoilerplate(value = '') {
  return /Patient Complexity:|Collaboration:|H\s*&\s*P Data Set:|Physical Exam Components:|Health Literacy:|Chart Data:|Social Problems Addressed:|Evidence-Based Practice Resources:|#\s*OTC Drugs/i.test(String(value || ''));
}

function cleanClinicalValue(value = '', maxLength = 140) {
  const text = clean(value);
  if (!text || isCaseLogBoilerplate(text)) return '';
  return compactText(text, maxLength);
}

function extractAgeFromText(text = '') {
  const source = clean(text);
  const patterns = [
    /\bAge\s*:\s*(\d{1,3})\s*years?\b/i,
    /\b(?:patient|client)\s+is\s+(?:an?\s+)?(\d{1,3})\s*[- ]?\s*year[- ]?old\b/i,
    /\b(\d{1,3})\s*[- ]?\s*year[- ]?old\s+(?:male|female|man|woman|patient|client)\b/i,
    /\bAge\s*[:=-]\s*(\d{1,3})\b/i,
  ];
  for (const pattern of patterns) {
    const age = source.match(pattern)?.[1];
    if (age) return age;
  }
  return '';
}

function extractSexFromText(text = '') {
  const source = clean(text);
  const labeled = source.match(/\bBiological Sex:\s*(Male|Female)\b/i)?.[1];
  if (labeled) return /^f/i.test(labeled) ? 'F' : 'M';
  if (/\b(?:female|woman)\b/i.test(source)) return 'F';
  if (/\b(?:male|man)\b/i.test(source)) return 'M';
  return '';
}

function extractHeightFromText(text = '') {
  const source = clean(text);
  return source.match(/\bheight\s*[:=-]?\s*([0-9]\s*['’]\s*\d{1,2}\s*(?:\"|in|inches)?)/i)?.[1]?.replace(/\s+/g, '') || '';
}

function extractWeightFromText(text = '') {
  const source = clean(text);
  const value = source.match(/\bweight\s*[:=-]?\s*(\d{2,3})(?:\s*(?:lb|lbs|pounds))?/i)?.[1];
  return value ? `${value} lb` : '';
}

function extractNarrativeAllergy(text = '') {
  const source = clean(text);
  const documented = source.match(/Allergy documented to\s+([^.;]+)/i)?.[1];
  if (documented) return clean(documented);
  const allergy = source.match(/\bAllerg(?:y|ies)\s*[:=-]\s*([^.;]+)/i)?.[1];
  return allergy ? clean(allergy) : '';
}

function extractMedicalHistorySummary(text = '') {
  const source = clean(text);
  const items = [];
  const add = (label, pattern) => {
    if (pattern.test(source) && !items.includes(label)) items.push(label);
  };
  add('Atrial fibrillation', /atrial\s+fibrillation|a[-\s]?fib|\bafib\b/i);
  add('Congestive heart failure', /congestive\s+heart\s+failure|\bchf\b|heart\s+failure/i);
  add('Hypertension', /hypertension|\bhtn\b/i);
  add('AV block', /\bAV\s+block\b|atrioventricular\s+block/i);
  add('Possible UTI', /possible\s+uti|\buti\b|urinary\s+tract\s+infection/i);
  add('Anemia', /anemia|hemoglobin|hgb|iron deficiency/i);
  add('Syncope', /\bsyncope\b|fainting|near syncope/i);
  return items.join('; ');
}

function extractSimpleDiagnosis(chiefComplaint = '', text = '') {
  const complaint = clean(chiefComplaint);
  if (complaint && !isCaseLogBoilerplate(complaint)) {
    return compactText(complaint.split(/[.;]/)[0], 40);
  }

  const source = clean(`${chiefComplaint} ${text}`);
  const diagnosisPatterns = [
    ['Weakness', /\bweakness\b/i],
    ['Possible UTI', /possible\s+uti|\buti\b|urinary\s+tract\s+infection/i],
    ['Syncope', /\bsyncope\b/i],
    ['Anemia', /severe symptomatic anemia|anemia|hemoglobin|hgb/i],
    ['Shortness of breath', /shortness\s+of\s+breath|\bsob\b/i],
    ['Headache', /\bheadache\b/i],
    ['Fatigue', /\bfatigue\b/i],
  ];
  const found = diagnosisPatterns.find(([, pattern]) => pattern.test(source));
  return found?.[0] || '';
}

function extractEncounterDate(text = '') {
  const source = String(text || '');
  const labeled = source.match(/Date\s+of\s+Encounter\s*:\s*([0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2,4}|[0-9]{4}-[0-9]{1,2}-[0-9]{1,2})/i)?.[1];
  if (labeled) return clean(labeled);
  return extractAfterLabelWithStops(source, 'Date of Encounter:', [
    'Student Information -',
    'Semester:',
    'Course:',
    'Clinical Faculty:',
    'Clinical Site:',
    'Time with Patient:',
  ]);
}

function splitMedicationEntries(raw = '') {
  const normalized = String(raw || '')
    .replace(/[\u2022\u2023\u25E6\u2043\u2219]/g, '\n')
    .replace(/Analgesics\s*&\s*Antipyretics\s*;\s*Miscellaneous/gi, 'Analgesics & Antipyretics')
    .replace(/\s+\/\s+/g, ' / ');

  const classLabels = [
    'Antibacterials',
    'Antithrombotic Agents: Anticoagulants',
    'Alpha Blockers',
    'Antilipemics',
    'Beta Blockers',
    'Caloric Agents',
    'Ion-removing: Calcium-removing Agents',
    'Replacement Preparations',
    'Anti-inflammatory: Corticosteroids',
    'Anti-in ammatory: Corticosteroids',
    'Antiulcer & Acid Suppress: Proton-pump Inhibitors',
    'Antidiabetic Agents: Insulins',
    'Antihypoglycemic Agents: Glycogenolytic',
    'Vitamin D',
    'Antianemia Drugs: Iron Preparations',
    'Analgesics & Antipyretics',
    'Cathartics & Laxatives',
    'Antigout Agents',
  ];

  const foundClassLabels = classLabels.filter((label) => {
    const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/in\\ ammatory/g, 'in(?:fl| )ammatory');
    return new RegExp(`\\b${escaped}\\b`, 'i').test(normalized);
  });
  if (foundClassLabels.length > 1) return Array.from(new Set(foundClassLabels.map((label) => label.replace('Anti-in ammatory', 'Anti-inflammatory'))));

  const lines = normalized
    .split(/\r?\n/)
    .map((line) => clean(line).replace(/^[-*\d\.)\s]+/, ''))
    .filter(Boolean)
    .filter((line) => !/^(none|n\/?a|no medications?)$/i.test(line));

  const entries = [];
  const pushEntry = (value) => {
    const next = clean(String(value || '').replace(/^[-*\d\.)\s]+/, ''));
    if (!next) return;
    if (/^(none|n\/?a|no medications?)$/i.test(next)) return;
    const key = next.toLowerCase();
    if (!entries.some((e) => e.toLowerCase() === key)) entries.push(next);
  };

  for (const line of lines) {
    const commaOrSemicolonParts = line
      .split(/\s*[;,]\s*/)
      .map((part) => clean(part))
      .filter(Boolean);

    const shouldSplit = commaOrSemicolonParts.length > 1
      && commaOrSemicolonParts.every((part) => /[a-z]/i.test(part) && part.length < 120);

    if (shouldSplit) {
      commaOrSemicolonParts.forEach(pushEntry);
    } else {
      pushEntry(line);
    }
  }

  return entries;
}

function capMedicationRows(inputMeds = [], maxRows = MED_TEMPLATE_ROW_CAP) {
  if (!Array.isArray(inputMeds)) return [];
  return inputMeds.slice(0, maxRows);
}

function getRepresentativeMedicationNames(label = '') {
  const text = clean(String(label || ''));
  if (!text) return [];
  for (const rule of MED_CLASS_EXAMPLES) {
    if (rule.pattern.test(text)) return rule.names;
  }
  return [];
}

function expandMedicationRows(inputMeds = [], maxRows = MED_TEMPLATE_ROW_CAP) {
  if (!Array.isArray(inputMeds) || !inputMeds.length) return [];

  const expanded = [];
  const seen = new Set();
  const pushMed = (med) => {
    if (expanded.length >= maxRows) return;
    const nameClass = clean(String(med?.nameClass || ''));
    if (!nameClass) return;
    const dedupeKey = nameClass.toLowerCase();
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);
    expanded.push({ ...med, nameClass });
  };

  for (const med of inputMeds) {
    if (expanded.length >= maxRows) break;
    const originalName = clean(String(med?.nameClass || ''));
    if (!originalName) continue;

    const splitNames = splitMedicationEntries(originalName);
    const normalizedSplitNames = splitNames.length ? splitNames : [originalName];

    for (const splitName of normalizedSplitNames) {
      if (expanded.length >= maxRows) break;
      const repsForSplit = getRepresentativeMedicationNames(splitName);
      if (repsForSplit.length) {
        for (const rep of repsForSplit) {
          if (expanded.length >= maxRows) break;
          pushMed({
            ...med,
            nameClass: rep,
            why: clean(String(med?.why || '')) || `Representative medication for ${splitName}.`,
          });
        }
      } else {
        const nextMed = { ...med, nameClass: splitName };
        pushMed(nextMed);
      }
    }

    if (expanded.length >= maxRows) continue;

    const reps = getRepresentativeMedicationNames(originalName);
    for (const rep of reps) {
      if (expanded.length >= maxRows) break;
      const nextMed = {
        ...med,
        nameClass: rep,
        why: clean(String(med?.why || '')) || 'Medication selected from class-level entry in case log.',
      };
      pushMed(nextMed);
    }
  }

  return expanded;
}

function tokenizeKeywords(input = '') {
  const stopwords = new Set([
    'the', 'and', 'for', 'with', 'from', 'that', 'this', 'were', 'was', 'are', 'have', 'has', 'had', 'into', 'about',
    'their', 'there', 'then', 'than', 'also', 'patient', 'clients', 'client', 'during', 'after', 'before', 'within',
    'while', 'where', 'which', 'because', 'would', 'could', 'should', 'being', 'been', 'into', 'onto', 'over', 'under',
  ]);
  return Array.from(new Set(
    String(input || '')
      .toLowerCase()
      .replace(/[^a-z0-9\s-]/g, ' ')
      .split(/\s+/)
      .map((w) => w.trim())
      .filter((w) => w.length > 3 && !stopwords.has(w))
  ));
}

function buildScholarlySearchQuery(fields, rawText, userPrompt = '') {
  const source = [
    fields?.diagnosis,
    fields?.psychosocial,
    fields?.nursingDiagnosis,
    fields?.goal,
    userPrompt,
    String(rawText || '').slice(0, 1200),
  ].filter(Boolean).join(' ');

  const keywords = tokenizeKeywords(source).slice(0, 8);
  const queryTerms = keywords.length ? keywords : ['psychiatric', 'nursing', 'patient'];
  const strictTerms = queryTerms.slice(0, 4);
  const broadTerms = queryTerms.slice(0, 8);
  const diagnosisTerms = tokenizeKeywords(fields?.diagnosis || '').slice(0, 3);
  const diagnosisClause = diagnosisTerms.length ? ` OR (${diagnosisTerms.join(' OR ')})` : '';

  const queries = [
    `(${strictTerms.join(' AND ')}) AND HAS_ABSTRACT:Y AND OPEN_ACCESS:Y`,
    `(${broadTerms.join(' OR ')}) AND HAS_ABSTRACT:Y AND OPEN_ACCESS:Y`,
    `(psychiatric OR mental health OR nursing OR aggression OR therapeutic communication${diagnosisClause}) AND HAS_ABSTRACT:Y AND OPEN_ACCESS:Y`,
  ];

  return {
    keywords: queryTerms,
    queries,
  };
}

function scoreEuropePmcResult(result, keywordSet, context = {}) {
  const journalTitle = getJournalTitle(result);
  const haystack = [result?.title, journalTitle, result?.abstractText, result?.authorString]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  let score = 0;
  for (const keyword of keywordSet) {
    if (haystack.includes(keyword.toLowerCase())) score += 1;
  }
  if (result?.pmcid) score += 1;
  if (result?.doi) score += 1;
  if (journalTitle) score += 1;

  const profileText = [
    context?.diagnosis,
    context?.psychosocial,
    context?.rawText,
  ].filter(Boolean).join(' ').toLowerCase();

  const perinatalTerms = /perinatal|pregnan|postpartum|maternal|neonatal|obstetric/;
  if (perinatalTerms.test(haystack) && !perinatalTerms.test(profileText)) {
    score -= 4;
  }

  const age = parseInt(String(context?.patientInitialAge || '').match(/\d+/)?.[0] || '', 10);
  if (!Number.isNaN(age) && age >= 40 && /adolescent|teen|youth|pediatric|child\b/.test(haystack)) {
    score -= 2;
  }

  return score;
}

function getJournalTitle(result) {
  return clean(String(
    result?.journalTitle
    || result?.journalInfo?.journal?.title
    || result?.journalInfo?.journal?.medlineAbbreviation
    || ''
  ));
}

function getJournalVolume(result) {
  return clean(String(result?.journalVolume || result?.journalInfo?.volume || ''));
}

function getJournalIssue(result) {
  return clean(String(result?.issue || result?.journalInfo?.issue || ''));
}

function getJournalPages(result) {
  return clean(String(result?.pageInfo || result?.journalInfo?.pageInfo || ''));
}

function getEuropePmcLink(result) {
  const fullTextList = result?.fullTextUrlList?.fullTextUrl;
  const urls = Array.isArray(fullTextList) ? fullTextList : (fullTextList ? [fullTextList] : []);
  const preferred = urls.find((u) => /pmc|pubmedcentral/i.test(String(u?.url || '')))?.url;
  const anyFullText = urls.find((u) => /^https?:\/\//i.test(String(u?.url || '')))?.url;
  if (preferred) return preferred;
  if (anyFullText) return anyFullText;

  const pmcid = clean(String(result?.pmcid || ''));
  if (pmcid) return `https://pmc.ncbi.nlm.nih.gov/articles/${pmcid.replace(/^PMC/i, 'PMC')}/`;

  const source = clean(String(result?.source || ''));
  const id = clean(String(result?.id || ''));
  if (source && id) return `https://europepmc.org/article/${source}/${id}`;

  return '';
}

function formatApaCitationFromEuropePmc(result) {
  const author = clean(String(result?.authorString || '')) || 'Unknown author';
  const year = clean(String(result?.pubYear || 'n.d.'));
  const title = clean(String(result?.title || 'Untitled article'));
  const journal = getJournalTitle(result) || 'Unknown Journal';
  const volume = getJournalVolume(result);
  const issue = getJournalIssue(result);
  const pages = getJournalPages(result);
  const doi = clean(String(result?.doi || ''));

  const volumeIssue = volume
    ? (issue ? `${volume}(${issue})` : volume)
    : '';
  const pagesPart = pages ? `${pages}` : '';
  const doiPart = doi ? ` https://doi.org/${doi.replace(/^https?:\/\/doi\.org\//i, '')}` : '';
  const mid = [volumeIssue, pagesPart].filter(Boolean).join(', ');

  return `${author} (${year}). ${title}. ${journal}${mid ? `, ${mid}` : ''}.${doiPart}`.trim();
}

function buildArticleRationale({ fields, prompt, articleTitle, matchedKeywords }) {
  const diagnosis = clean(String(fields?.diagnosis || ''));
  const psychosocial = clean(String(fields?.psychosocial || ''));
  const nursingDx = clean(String(fields?.nursingDiagnosis || ''));
  const promptText = clean(String(prompt || ''));
  const keywordLine = matchedKeywords.length ? matchedKeywords.slice(0, 3).join(', ') : 'core case priorities';

  const base = diagnosis
    ? `This article is relevant because it supports care planning for ${diagnosis.toLowerCase()} and aligns with this patient profile.`
    : `This article is relevant because it aligns with this patient profile and current nursing priorities.`;

  const context = psychosocial
    ? `It also reflects key psychosocial factors: ${psychosocial.slice(0, 110)}${psychosocial.length > 110 ? '...' : ''}.`
    : `It overlaps with key case themes (${keywordLine}).`;

  const nursingLine = nursingDx
    ? `It supports nursing priorities related to ${nursingDx.slice(0, 110)}${nursingDx.length > 110 ? '...' : ''}.`
    : `It provides useful support for nursing assessment, intervention planning, and evaluation.`;

  const articleLine = articleTitle
    ? `The selected source, "${articleTitle}", provides evidence that can be translated into bedside decisions.`
    : '';

  const focus = promptText ? `Secondary focus used: ${promptText}.` : '';
  return [base, context, nursingLine, articleLine, focus].filter(Boolean).join(' ');
}

async function findScholarlyArticle({ fields, rawText, userPrompt, excludedLinks = [], excludedTitles = [] }) {
  const { queries, keywords } = buildScholarlySearchQuery(fields, rawText, userPrompt);
  const profileText = [fields?.diagnosis, fields?.psychosocial, rawText].filter(Boolean).join(' ').toLowerCase();
  const profileHasPerinatalContext = /perinatal|pregnan|postpartum|maternal|neonatal|obstetric/.test(profileText);
  let candidates = [];
  let queryUsed = '';
  let lastStatus = 0;

  for (const query of queries) {
    queryUsed = query;
    const endpoint = `https://www.ebi.ac.uk/europepmc/webservices/rest/search?query=${encodeURIComponent(query)}&format=json&pageSize=50&resultType=core`;
    const response = await fetch(endpoint);
    lastStatus = response.status;
    if (!response.ok) continue;
    const data = await response.json();
    const results = Array.isArray(data?.resultList?.result) ? data.resultList.result : [];
    candidates = results
      .filter((r) => clean(String(r?.title || '')))
      .map((r) => ({
        ...r,
        _link: getEuropePmcLink(r),
        _score: scoreEuropePmcResult(r, keywords, {
          diagnosis: fields?.diagnosis,
          psychosocial: fields?.psychosocial,
          rawText,
          patientInitialAge: fields?.patientInitialAge,
        }),
      }))
      .filter((r) => Boolean(clean(String(r?._link || ''))))
      .sort((a, b) => b._score - a._score);

    if (!profileHasPerinatalContext) {
      const nonPerinatal = candidates.filter((candidate) => {
        const text = [candidate?.title, candidate?.abstractText].filter(Boolean).join(' ').toLowerCase();
        return !/perinatal|pregnan|postpartum|maternal|neonatal|obstetric/.test(text);
      });
      if (nonPerinatal.length) candidates = nonPerinatal;
    }

    if (candidates.length) break;
  }

  if (!candidates.length) {
    if (lastStatus >= 400) {
      throw new Error(`Scholarly search failed (${lastStatus}). Try again in a moment.`);
    }
    throw new Error('No full-access scholarly article was found for this query. Try a shorter focus prompt (for example: aggression in psychiatric patients and therapeutic communication).');
  }

  const excludedLinkSet = new Set((excludedLinks || []).map((v) => clean(String(v || '')).toLowerCase()).filter(Boolean));
  const excludedTitleSet = new Set((excludedTitles || []).map((v) => clean(String(v || '')).toLowerCase()).filter(Boolean));
  const filteredCandidates = candidates.filter((candidate) => {
    const linkKey = clean(String(candidate?._link || getEuropePmcLink(candidate) || '')).toLowerCase();
    const titleKey = clean(String(candidate?.title || '')).toLowerCase();
    if (linkKey && excludedLinkSet.has(linkKey)) return false;
    if (titleKey && excludedTitleSet.has(titleKey)) return false;
    return true;
  });

  const top = filteredCandidates[0] || candidates[0];
  const isRepeat = filteredCandidates.length === 0;
  const link = getEuropePmcLink(top);
  if (!link) {
    throw new Error('No full-access article link was available. Try a different focus prompt.');
  }
  const title = clean(String(top?.title || 'Untitled article'));
  const why = buildArticleRationale({ fields, prompt: userPrompt, articleTitle: title, matchedKeywords: keywords });
  const citation = formatApaCitationFromEuropePmc(top);

  return { title, link, why, citation, queryUsed, isRepeat };
}

async function draftArticleRelevanceWithAi({ apiKey, fields, rawText, userPrompt, articleResult }) {
  const safeKey = clean(String(apiKey || ''));
  if (!safeKey) throw new Error('OpenAI API key is required for AI write-up.');

  const caseSummary = clean(String(rawText || '')).slice(0, 2000);
  const diagnosis = clean(String(fields?.diagnosis || ''));
  const psychosocial = clean(String(fields?.psychosocial || ''));
  const nursingDiagnosis = clean(String(fields?.nursingDiagnosis || ''));
  const promptText = clean(String(userPrompt || ''));

  const systemPrompt = [
    'You are a nursing student write-up assistant.',
    'Write a concise clinical rationale for why a scholarly article is relevant to this case.',
    'Use 2-5 sentences, plain academic language, and avoid fabricating details.',
  ].join(' ');

  const userMessage = `Article title: ${articleResult?.title || ''}
Article link: ${articleResult?.link || ''}
APA citation: ${articleResult?.citation || ''}
Search focus prompt: ${promptText || 'N/A'}
Patient diagnosis: ${diagnosis || 'N/A'}
Psychosocial findings: ${psychosocial || 'N/A'}
Nursing diagnosis: ${nursingDiagnosis || 'N/A'}

Case text excerpt:
${caseSummary || 'N/A'}

Task: Write the section "Why this article is relevant".`;

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${safeKey}`,
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      temperature: 0.2,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userMessage },
      ],
    }),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`AI write-up failed (${response.status}): ${errorText}`);
  }

  const data = await response.json();
  const content = data?.choices?.[0]?.message?.content;
  const text = Array.isArray(content) ? content.map((p) => p?.text || '').join('') : content;
  return clean(String(text || ''));
}

function parseCaseText(rawText) {
  const text = clean(rawText);
  const labels = [
    'Date of Encounter:',
    'Student Information -',
    'Semester:',
    'Course:',
    'Clinical Faculty:',
    'Clinical Site:',
    'Time with Patient:',
    'Student Participation:',
    'Chief Complaint:',
    'Patient Complexity:',
    'Collaboration:',
    'H & P Data Set:',
    'Patient History:',
    'Physical Exam Components:',
    'Vital Signs:',
    'Init. Pain Score:',
    'Health Literacy:',
    'Chart Data:',
    'Social Problems Addressed:',
    'Nutrition/Exercise Safety',
    'Patient Education:',
    'Evidence-Based Practice Resources:',
    'Medications',
    '# OTC Drugs taken regularly:',
    '# Prescriptions currently prescribed:',
    '# New/Re',
    'Clinical Notes',
    'Personal Note:',
    'Encounter Continuity',
    'Medications Prescribed Today:',
    'Adherence Issues with Medications:',
    'Discharge Plan:',
    'Pain:',
    'Allergy:',
    'Allergies:',
  ];

  const studentName = extractAfterLabelWithStops(text, 'Student Information -', labels);
  const date = extractEncounterDate(text);
  const semester = extractAfterLabelWithStops(text, 'Semester:', labels);
  const course = extractAfterLabelWithStops(text, 'Course:', labels);
  const faculty = extractAfterLabelWithStops(text, 'Clinical Faculty:', labels);
  const site = extractAfterLabelWithStops(text, 'Clinical Site:', labels);
  const timeWithPatient = extractAfterLabelWithStops(text, 'Time with Patient:', labels);
  const studentParticipation = extractAfterLabelWithStops(text, 'Student Participation:', labels);
  const chiefComplaint = extractAfterLabelWithStops(text, 'Chief Complaint:', labels);

  const clinicalNotes = getSection(text, 'Clinical Notes', ['Personal Note:', 'Encounter Continuity', 'Medications Prescribed Today:', 'Adherence Issues with Medications:']);
  const personalNote = getSection(text, 'Personal Note:', ['Encounter Continuity']);
  const dischargePlanFromLog = extractAfterLabelWithStops(text, 'Discharge Plan:', labels);
  const painFromLog = extractAfterLabelWithStops(text, 'Pain:', labels);
  const allergyFromLog = extractAfterLabelWithStops(text, 'Allergy:', labels) || extractAfterLabelWithStops(text, 'Allergies:', labels) || extractNarrativeAllergy(text);

  const age = extractAgeFromText(`${clinicalNotes}\n${text}`);
  const sex = extractSexFromText(`${clinicalNotes}\n${text}`);
  const ht = extractHeightFromText(`${clinicalNotes}\n${text}`);
  const wt = extractWeightFromText(`${clinicalNotes}\n${text}`);
  const diagnosis = extractSimpleDiagnosis(chiefComplaint, `${clinicalNotes}\n${text}`) || inferDiagnosis(chiefComplaint, clinicalNotes);
  const medicalHistory = extractMedicalHistorySummary(`${clinicalNotes}\n${text}`);
  const painScore = extractPainScore(`${painFromLog}\n${clinicalNotes}\n${text}`);
  const pain = painScore || painFromLog || (/pain/i.test(chiefComplaint) ? chiefComplaint : '');
  const safety = /suicidal ideation|homicidal ideation|hallucinations/i.test(text)
    ? 'Follow-up assessment needed for SI/HI, hallucinations, sleep, and home safety.'
    : '';

  const fields = {
    ...DEFAULT_STATE,
    studentName,
    date,
    week: '',
    patientInitialAge: age,
    age,
    sex,
    ht,
    wt,
    diagnosis,
    allergy: allergyFromLog,
    allergies: allergyFromLog,
    immunizations: '',
    neuro: '',
    cardio: '',
    respiratory: '',
    vitals: '',
    gi: '',
    gu: '',
    skin: '',
    hygiene: '',
    skinTurgor: '',
    fluidElectrolytes: '',
    pain,
    psychosocial: cleanClinicalValue(clinicalNotes, 220),
    supportSystem: '',
    responseHospitalization: '',
    activity: [studentParticipation, timeWithPatient].filter(Boolean).join('; '),
    assistanceAdls: '',
    sleepPattern: '',
    nutritionalStatus: '',
    cultural: '',
    spiritualAssessment: '',
    educationNeeds: '',
    safety,
    diagnosticTests: '',
    labs: '',
    labTestName: '',
    labClientResults: '',
    labNormalValue: '',
    labInterpretation: '',
    dischargePlan: dischargePlanFromLog,
    medicalHistory,
    surgicalHistory: '',
    genAppearance: '',
    ivLocation: '',
    surgicalIncision: '',
    orientation: /alert and oriented x?3|oriented x?3|a&o x?3/i.test(text) ? 'A&O x3' : '',
    speech: '',
    weakness: /weakness/i.test(text) ? 'Generalized weakness reported.' : '',
    breathSounds: /shortness of breath|sob/i.test(text) ? 'SOB reported; assess sounds' : '',
    peripheralPulses: '',
    edema: '',
    bowelSounds: '',
    physicalOther: '',
    medsPrior: '',
    currentMedDate: date,
    currentMedOrder: '',
    currentMedIndication: '',
    nursingDiagnosis: '',
    goal: '',
    plan: '',
    intervention: '',
    rationale: '',
    evaluation: '',
    reassessment: '',
    theorist: '',
    knowledgeGained: personalNote,
    courseObjectives: '',
    semesterMeta: semester,
    courseMeta: course,
    facultyMeta: faculty,
    siteMeta: site,
  };

  const medications = [];
  const medBlock = text.match(/Medications Prescribed Today:\s*([\s\S]*?)Adherence Issues with Medications:/i)?.[1] || '';
  splitMedicationEntries(medBlock).forEach((m) => {
      medications.push({
        nameClass: m,
        doseRoute: '',
        why: '',
        action: '',
        implications: '',
        sideEffects: '',
      });
    });

  return { fields, medications: capMedicationRows(medications), rawText: text };
}

function extractVsimCaseNumber(value = '') {
  return clean(normalizeSpacedPdfText(value)).match(/\bv[\s-]*sim[\s-]*(?:case[\s-]*)?([1-5])\b/i)?.[1] || '';
}

function extractSimulationClientName(text = '', sourceHint = '') {
  const source = clean(text);
  const caseHint = extractVsimCaseNumber(sourceHint) || extractVsimCaseNumber(source.slice(0, 600));
  if (caseHint === '1') return 'Olivia Jones';
  if (caseHint === '2') return 'Brenda Patton';
  if (caseHint === '3') return 'Amelia Sung';
  if (caseHint === '4') return 'Carla Hernandez';
  if (caseHint === '5') return 'Fatime Sanogo';
  const explicit = source.match(/\b(?:Patient|Client)\s+Name\s*[:=-]\s*([A-Z][A-Za-z'.-]+(?:\s+[A-Z][A-Za-z'.-]+){0,2})/i)?.[1]
    || source.match(/\bName\s*[:=-]\s*([A-Z][A-Za-z'.-]+(?:\s+[A-Z][A-Za-z'.-]+){1,2})/i)?.[1];
  if (explicit) return clean(explicit);
  if (/Fatime\s+Sanogo/i.test(source)) return 'Fatime Sanogo';
  if (/Carla\s+Hernandez/i.test(source)) return 'Carla Hernandez';
  if (/Amelia\s+Sung/i.test(source)) return 'Amelia Sung';
  if (/Olivia\s+Jones/i.test(source)) return 'Olivia Jones';
  if (/Brenda\s+Patton/i.test(source) && !/Fatime\s+Sanogo/i.test(source)) return 'Brenda Patton';
  if (/vsim\s*3\b|shoulder dystocia|mcroberts|suprapubic pressure/i.test(source)) return 'Amelia Sung';
  if (/vsim\s*4\b|prolapsed cord|cord prolapse|umbilical cord/i.test(source)) return 'Carla Hernandez';
  if (/vsim\s+case\s*1\b|severe\s+preeclampsia|eclamptic seizure|magnesium sulfate/i.test(source)) return 'Olivia Jones';
  if (/vsim\s+case\s*2\b|gbs|group b strep|penicillin g|latent phase/i.test(source)) return 'Brenda Patton';
  return '';
}

function buildSimulationPriorityFields(profile = {}) {
  if (profile.isPostpartumHemorrhage) {
    return {
      nd1Diagnosis: 'Deficient fluid volume',
      nd1Assessment: withPriorityPrompt('nd1Assessment', 'QBL 800 mL, BP 90/50, HR 120, weakness.'),
      nd1Rationale: withPriorityPrompt('nd1Rationale', 'Blood loss decreases circulating volume and perfusion.'),
      nd1Intervention: 'Quantify bleeding, monitor VS/LOC/urine output, maintain IV access, and prepare blood products.',
      nd1Evaluation: 'Bleeding and perfusion require continued reassessment.',
      nd2Diagnosis: 'Ineffective tissue perfusion',
      nd2Assessment: withPriorityPrompt('nd2Assessment', 'Tachycardia, hypotension, and decreased urine-output risk.'),
      nd2Rationale: withPriorityPrompt('nd2Rationale', 'Hypovolemia can reduce cerebral and organ perfusion.'),
      nd2Intervention: 'Assess LOC, capillary refill, urine output, CBC/type-crossmatch, and response to fluids/blood.',
      nd2Evaluation: 'Perfusion improves with stable VS and urine output.',
      nd3Diagnosis: 'Impaired urinary elimination',
      nd3Assessment: withPriorityPrompt('nd3Assessment', 'Distended bladder displaces uterus and worsens atony.'),
      nd3Rationale: withPriorityPrompt('nd3Rationale', 'Full bladder prevents effective uterine contraction.'),
      nd3Intervention: 'Assist voiding or catheterization per order and reassess fundus midline/firmness.',
      nd3Evaluation: 'Fundus firm/midline and bleeding reduced.',
    };
  }

  if (profile.isCordProlapse) {
    return {
      nd1Diagnosis: 'Impaired fetal gas exchange',
      nd1Assessment: withPriorityPrompt('nd1Assessment', 'Prolapsed cord with variable decelerations/compression.'),
      nd1Rationale: withPriorityPrompt('nd1Rationale', 'Cord compression can interrupt fetal oxygen delivery.'),
      nd1Intervention: 'Call for help, elevate presenting part, reposition knee-chest/Trendelenburg, and monitor FHR.',
      nd1Evaluation: 'FHR pattern and fetal status require urgent reassessment.',
      nd2Diagnosis: 'Anxiety',
      nd2Assessment: withPriorityPrompt('nd2Assessment', 'Emergency cesarean preparation creates acute stress.'),
      nd2Rationale: withPriorityPrompt('nd2Rationale', 'Clear communication can reduce fear during emergency care.'),
      nd2Intervention: 'Explain actions briefly, keep support person informed, and maintain calm direction.',
      nd2Evaluation: 'Patient/support person verbalize understanding of emergency plan.',
      nd3Diagnosis: 'Risk for maternal injury',
      nd3Assessment: withPriorityPrompt('nd3Assessment', 'Emergency surgery and rapid positioning are required.'),
      nd3Rationale: withPriorityPrompt('nd3Rationale', 'Urgent interventions increase procedural safety needs.'),
      nd3Intervention: 'Maintain IV access, prepare consent/labs, apply safety positioning, and coordinate OR transfer.',
      nd3Evaluation: 'Patient transferred safely for definitive care.',
    };
  }

  if (profile.isShoulderDystocia) {
    return {
      nd1Diagnosis: 'Impaired fetal gas exchange',
      nd1Assessment: withPriorityPrompt('nd1Assessment', 'Shoulder dystocia with delayed body delivery.'),
      nd1Rationale: withPriorityPrompt('nd1Rationale', 'Prolonged dystocia can reduce fetal oxygenation.'),
      nd1Intervention: 'Call for help, document times, assist McRoberts maneuver, and prepare newborn resuscitation.',
      nd1Evaluation: 'Fetal/newborn status assessed immediately after delivery.',
      nd2Diagnosis: 'Acute pain',
      nd2Assessment: withPriorityPrompt('nd2Assessment', 'Emergency maneuvers and labor increase discomfort.'),
      nd2Rationale: withPriorityPrompt('nd2Rationale', 'Dystocia care may intensify perineal and musculoskeletal pain.'),
      nd2Intervention: 'Support positioning, coach breathing, reassess pain, and provide comfort measures.',
      nd2Evaluation: 'Pain/coping response requires reassessment after birth.',
      nd3Diagnosis: 'Risk for birth injury',
      nd3Assessment: withPriorityPrompt('nd3Assessment', 'Macrosomia/shoulder dystocia increases brachial plexus risk.'),
      nd3Rationale: withPriorityPrompt('nd3Rationale', 'Traction and shoulder impaction can injure newborn structures.'),
      nd3Intervention: 'Avoid fundal pressure, assist approved maneuvers, and assess newborn clavicle/arm movement.',
      nd3Evaluation: 'Newborn assessed for movement, Apgar, and injury signs.',
    };
  }

  if (profile.isPreeclampsia) {
    return {
      nd1Diagnosis: 'Risk for maternal injury',
      nd1Assessment: withPriorityPrompt('nd1Assessment', 'Severe preeclampsia; seizure precautions needed.'),
      nd1Rationale: withPriorityPrompt('nd1Rationale', 'CNS irritability can progress to eclamptic seizure.'),
      nd1Intervention: 'Maintain seizure precautions, oxygen/suction access, quiet environment, and side-lying position.',
      nd1Evaluation: 'No seizure or injury noted; continue monitoring.',
      nd2Diagnosis: 'Ineffective tissue perfusion',
      nd2Assessment: withPriorityPrompt('nd2Assessment', 'Severe-range BP with headache/visual-change risk.'),
      nd2Rationale: withPriorityPrompt('nd2Rationale', 'Hypertension can reduce maternal and placental perfusion.'),
      nd2Intervention: 'Monitor BP, neuro status, urine protein/output, fetal status, and ordered antihypertensives.',
      nd2Evaluation: 'Perfusion status requires continued reassessment.',
      nd3Diagnosis: 'Deficient knowledge',
      nd3Assessment: withPriorityPrompt('nd3Assessment', 'Patient needs teaching about warning signs and magnesium therapy.'),
      nd3Rationale: withPriorityPrompt('nd3Rationale', 'Understanding care helps reporting and safety.'),
      nd3Intervention: 'Teach headache, visual changes, epigastric pain, decreased fetal movement, and medication safety.',
      nd3Evaluation: 'Patient verbalizes key warning signs.',
    };
  }

  if (profile.isObLabor) {
    return {
      nd1Diagnosis: 'Risk for infection',
      nd1Assessment: withPriorityPrompt('nd1Assessment', 'GBS positive; intrapartum antibiotics indicated.'),
      nd1Rationale: withPriorityPrompt('nd1Rationale', 'GBS colonization can increase newborn infection risk.'),
      nd1Intervention: 'Administer ordered IV antibiotics and monitor maternal/fetal status.',
      nd1Evaluation: 'No infection signs noted; continue prophylaxis.',
      nd2Diagnosis: 'Acute pain',
      nd2Assessment: withPriorityPrompt('nd2Assessment', 'Contractions every 5-10 minutes in latent labor.'),
      nd2Rationale: withPriorityPrompt('nd2Rationale', 'Uterine contractions cause labor discomfort.'),
      nd2Intervention: 'Assess pain, support breathing/position changes, and reassess labor progress.',
      nd2Evaluation: 'Coping and pain response require reassessment.',
      nd3Diagnosis: 'Deficient knowledge',
      nd3Assessment: withPriorityPrompt('nd3Assessment', 'Needs teaching about GBS, fetal monitoring, and labor progress.'),
      nd3Rationale: withPriorityPrompt('nd3Rationale', 'Teaching supports informed participation in care.'),
      nd3Intervention: 'Explain GBS prophylaxis, normal FHR range, and when to report symptoms.',
      nd3Evaluation: 'Patient verbalizes understanding of care plan.',
    };
  }

  return {};
}

function parseSimulationNotesText(rawText, sourceHint = '') {
  const text = clean(normalizeSpacedPdfText(rawText));
  const hintText = clean(normalizeSpacedPdfText(sourceHint));
  const profileSource = clean(`${hintText} ${text}`);
  const caseHint = extractVsimCaseNumber(hintText) || extractVsimCaseNumber(text.slice(0, 600));
  const isPostpartumHemorrhage = caseHint === '5' || (!caseHint && /Fatime Sanogo|postpartum hemorrhage|hemorrhaging after giving birth|quantitative blood loss|\bqbl\s*800|blood loss 800|boggy uterus|uterine atony|retained placental|retained tissue/i.test(profileSource));
  const isCordProlapse = caseHint === '4' || (!caseHint && !isPostpartumHemorrhage && /Carla Hernandez|prolapsed cord|cord prolapse|umbilical cord|visible at the vulva|variable decelerations|knee-chest|presenting fetal part/i.test(profileSource));
  const isShoulderDystocia = caseHint === '3' || (!caseHint && !isPostpartumHemorrhage && !isCordProlapse && /Amelia Sung|shoulder dystocia|McRoberts|suprapubic pressure|macrosomia|brachial plexus|Erb palsy|turtle sign/i.test(profileSource));
  const isObLabor = caseHint === '2' || (!caseHint && !isPostpartumHemorrhage && !isCordProlapse && !isShoulderDystocia && /Brenda Patton|group b strep|gbs positive|penicillin g|intrapartum prophylaxis|primigravida[^.]+labor/i.test(profileSource));
  const isPreeclampsia = caseHint === '1' || (!caseHint && !isPostpartumHemorrhage && !isCordProlapse && !isShoulderDystocia && !isObLabor && /Olivia Jones|severe preeclampsia|eclampsia|eclamptic seizure|magnesium sulfate|clonus|deep tendon reflex|seizure precautions|visual changes|protein in the urine/i.test(profileSource));
  const date = extractEncounterDate(text);
  const clientName = extractSimulationClientName(text, sourceHint);
  const age = extractAgeFromText(text) || (isPreeclampsia ? '23' : (isObLabor ? '18' : (isShoulderDystocia ? '36' : (isCordProlapse ? '32' : (isPostpartumHemorrhage ? '22' : '')))));
  const sex = extractSexFromText(text) || (isObLabor || isPreeclampsia || isPostpartumHemorrhage || isShoulderDystocia || isCordProlapse ? 'F' : '');
  const ht = extractHeightFromText(text);
  const wt = extractWeightFromText(text);
  const vitals = extractVitalsFromSummary(text);
  const painScore = extractPainScore(text);
  const meds = [];

  if (/gbs|group b strep|intrapartum prophylaxis|prophylaxis antibiotics/i.test(text)) {
    meds.push({
      nameClass: 'Penicillin G (GBS prophylaxis antibiotic)',
      doseRoute: 'IV per labor protocol',
      why: 'GBS intrapartum prophylaxis',
      action: '',
      implications: '',
      sideEffects: '',
    });
  }
  if (/magnesium sulfate/i.test(text)) {
    meds.push({
      nameClass: 'Magnesium sulfate',
      doseRoute: 'IV per preeclampsia protocol',
      why: 'Prevent eclamptic seizures',
      action: '',
      implications: '',
      sideEffects: '',
    });
  }
  if (isPostpartumHemorrhage) {
    if (/misoprostol/i.test(text)) {
      meds.push({
        nameClass: 'Misoprostol',
        doseRoute: 'Per postpartum hemorrhage protocol',
        why: 'Promote uterine contraction',
        action: '',
        implications: '',
        sideEffects: '',
      });
    }
    if (/oxytocin/i.test(text)) {
      meds.push({
        nameClass: 'Oxytocin',
        doseRoute: 'IV/IM per protocol',
        why: 'Uterine tone/bleeding control',
        action: '',
        implications: '',
        sideEffects: '',
      });
    }
  }

  const hasUa = /urinalysis|white blood cell|wbc|red blood cell|rbc|leuko|nitrate|glucose|ketone/i.test(text);
  const hasHepB = /hepatitis\s*b|hbsag|surface antigen/i.test(text);
  const hasGbs = /\bgbs\b|group b strep/i.test(text);
  const hasMagLevel = /therapeutic blood level|4\s*(?:to|-)\s*7\s*m\s*e?q/i.test(text);
  const hasCnsSigns = /4\s*\+?\s*deep tendon reflex|persistent headaches|visual changes|clonus|central nervous system/i.test(text);
  const hasBloodLoss = /quantitative blood loss 800|blood loss 800|bp\s*90\/50|hr\s*120|tachycardia|hypotension|weak/i.test(text);
  const hasRetainedTissue = /retained tissue|retained placental|placental fragments/i.test(text);
  const hasBoggyUterus = /boggy uterus|uterine atony|uterine tone|fundus/i.test(text);
  const hasShoulderManeuvers = /McRoberts|suprapubic pressure|step stool|newborn resuscitation/i.test(text);
  const hasCordCompression = /variable decelerations|cord compression|prolapsed cord|umbilical cord/i.test(text);
  const laborFinding = clean(text.match(/(?:contractions every\s+[^.]+|vaginal exam(?:ination)?\s+(?:showing|of)?\s*[^.]+)/i)?.[0] || '');
  const priorityFields = buildSimulationPriorityFields({ isObLabor, isPreeclampsia, isPostpartumHemorrhage, isShoulderDystocia, isCordProlapse });

  const fields = {
    ...DEFAULT_STATE,
    date,
    patientInitialAge: age,
    age,
    sex,
    ht,
    wt,
    clientName,
    diagnosis: isPostpartumHemorrhage ? 'Postpartum hemorrhage' : (isCordProlapse ? 'Umbilical cord prolapse' : (isShoulderDystocia ? 'Shoulder dystocia' : (isPreeclampsia ? 'Severe preeclampsia' : (isObLabor ? 'Latent labor; GBS positive' : extractSimpleDiagnosis('', text))))),
    allergy: '',
    allergies: '',
    immunizations: hasHepB ? 'Assess Hep B status' : '',
    sleepPattern: 'Need to assess',
    nutritionalStatus: /glucose|ketone|nausea|vomiting|poor appetite/i.test(text) ? 'Assess nutrition/hydration' : 'Need to assess',
    assistanceAdls: 'Independent',
    hygiene: 'Independent',
    medsPrior: '',
    medicalHistory: isPostpartumHemorrhage
      ? compactText([
        'Postpartum hemorrhage after vaginal birth',
        hasBloodLoss ? 'QBL 800 mL; BP 90/50; HR 120; weakness' : '',
        hasBoggyUterus ? 'Boggy uterus/fundal concern' : '',
        hasRetainedTissue ? 'Retained tissue suspected' : '',
      ].filter(Boolean).join('; '), 115)
      : (isCordProlapse
      ? compactText([
        '39 weeks gestation',
        'Umbilical cord prolapse',
        hasCordCompression ? 'cord compression/variable decelerations' : '',
        'emergency cesarean preparation',
      ].filter(Boolean).join('; '), 115)
      : (isShoulderDystocia
      ? compactText([
        '40 weeks gestation',
        'shoulder dystocia risk',
        'macrosomia/large fetal size concern',
        hasShoulderManeuvers ? 'McRoberts/suprapubic maneuvers' : '',
      ].filter(Boolean).join('; '), 115)
      : (isObLabor
      ? compactText([
        'Primigravida in labor',
        hasGbs ? 'GBS positive' : '',
        hasHepB ? 'HBsAg positive' : '',
        hasUa ? 'UA glucose/ketones positive; no UTI indicators' : '',
      ].filter(Boolean).join('; '), 115)
      : (isPreeclampsia ? 'Severe preeclampsia; seizure risk; CNS involvement signs reviewed.' : extractMedicalHistorySummary(text))))),
    surgicalHistory: 'None',
    supportSystem: 'Need to assess',
    responseHospitalization: isPostpartumHemorrhage ? 'Postpartum hemorrhage management' : (isCordProlapse ? 'Emergency cord-prolapse care' : (isShoulderDystocia ? 'Emergency shoulder dystocia care' : (isPreeclampsia ? 'Admitted for severe preeclampsia care' : (isObLabor ? 'Admitted to labor and delivery' : 'Cooperative with care')))),
    genAppearance: isPostpartumHemorrhage ? 'Postpartum; weak with bleeding concern' : (isCordProlapse ? 'Laboring patient; emergency care' : (isShoulderDystocia ? 'Laboring patient; dystocia emergency' : (isPreeclampsia ? 'Pregnant patient; seizure precautions' : (isObLabor ? 'Laboring patient; no distress stated' : '')))),
    ivLocation: meds.length ? 'Need to assess IV site' : '',
    surgicalIncision: 'None',
    orientation: 'A&O x4',
    speech: 'Clear',
    weakness: isPostpartumHemorrhage ? 'Weakness reported' : 'None noted',
    skinTurgor: /ketone|dehydrat/i.test(text) ? 'Assess hydration' : 'Normal',
    breathSounds: 'Clear',
    peripheralPulses: 'Present',
    edema: 'Need to assess',
    bowelSounds: 'Present',
    physicalOther: isPostpartumHemorrhage ? 'Boggy uterus; distended bladder' : (isCordProlapse ? 'Cord prolapse; elevate presenting part' : (isShoulderDystocia ? 'Shoulder dystocia; McRoberts' : (isPreeclampsia ? 'Seizure precautions; assess DTR/clonus' : (laborFinding || (isObLabor ? 'First stage, latent phase' : ''))))),
    temp: vitals.temp || (isPostpartumHemorrhage ? '97.7 F' : '98.6 F'),
    pulse: vitals.pulse || (isPostpartumHemorrhage ? '120 bpm' : (isPreeclampsia ? '92 bpm' : '88 bpm')),
    resp: vitals.resp || '18/min',
    bp: vitals.bp || (isPostpartumHemorrhage ? '90/50' : (isPreeclampsia ? '160/100' : '118/72')),
    pain: painScore || (isShoulderDystocia ? 'Labor pain; assess' : (isCordProlapse ? 'Assess pain/anxiety' : (isPostpartumHemorrhage ? 'Assess pain' : (isPreeclampsia ? 'Headache; assess pain' : (isObLabor ? 'Contraction discomfort' : ''))))),
    vitals: '',
    psychosocial: isPostpartumHemorrhage
      ? compactText('Postpartum hemorrhage may cause fear/anxiety; provide calm updates and support person involvement.', 220)
      : (isCordProlapse
      ? compactText('Cord prolapse requires emergency cesarean preparation; provide brief reassurance and updates.', 220)
      : (isShoulderDystocia
      ? compactText('Shoulder dystocia emergency may increase anxiety; maintain clear team communication and support.', 220)
      : (isPreeclampsia
      ? compactText('Hospital admission for severe preeclampsia; support anxiety reduction with quiet environment and teaching.', 220)
      : (isObLabor
      ? compactText('Admitted to labor and delivery; assess coping, support person, teaching needs, and labor anxiety.', 220)
        : cleanClinicalValue(text, 220))))),
    cultural: '',
    spiritualAssessment: 'No needs stated',
    educationNeeds: isPostpartumHemorrhage
      ? 'Teach bleeding warning signs, fundal checks, medication purpose, and call-for-help precautions.'
      : (isCordProlapse
      ? 'Explain emergency positioning, fetal monitoring, cesarean preparation, and call-for-help steps.'
      : (isShoulderDystocia
      ? 'Explain shoulder dystocia actions, newborn assessment, pain control, and emergency team roles.'
      : (isPreeclampsia
      ? 'Teach seizure precautions, magnesium therapy, warning signs, BP monitoring, and when to call for help.'
      : (isObLabor
      ? 'Teach labor progress, fetal monitoring, GBS prophylaxis, Hep B precautions, and when to report changes.'
      : buildEducationalNeedsSuggestion(DEFAULT_STATE, text))))),
    safety: isPostpartumHemorrhage ? 'Hemorrhage precautions; assist ambulation; monitor perfusion.' : (isCordProlapse ? 'Emergency OR preparation; fetal monitoring.' : (isShoulderDystocia ? 'Call for help; avoid fundal pressure.' : (isPreeclampsia ? 'Seizure precautions; quiet/dim environment; lateral positioning.' : (isObLabor ? 'Maintain maternal/fetal monitoring and infection precautions as ordered.' : '')))),
    diagnosticTests: compactText([
      isPostpartumHemorrhage ? 'Quantitative blood loss' : '',
      isPostpartumHemorrhage ? 'CBC/type and crossmatch' : '',
      isCordProlapse ? 'Fetal monitoring' : '',
      isCordProlapse ? 'CBC/type-cross/pre-op labs' : '',
      isShoulderDystocia ? 'FHR/Apgar/newborn assessment' : '',
      hasUa ? 'Urinalysis' : '',
      hasHepB ? 'Hepatitis B surface antigen' : '',
      hasGbs ? 'GBS culture' : '',
      hasMagLevel ? 'Magnesium level' : '',
      hasCnsSigns ? 'Neuro assessment' : '',
      /fetal heart rate|\bfhr\b/i.test(text) ? 'Fetal monitoring' : '',
    ].filter(Boolean).join('; '), 120),
    labs: compactText([
      isPostpartumHemorrhage ? 'QBL 800 mL; BP 90/50; HR 120; monitor Hgb/Hct/platelets/type-cross.' : '',
      isCordProlapse ? 'Variable decelerations/cord compression; review CBC/type-cross/UA.' : '',
      isShoulderDystocia ? 'Estimated fetal weight >=4000 g; document delivery times and newborn status.' : '',
      hasUa ? 'UA negative WBC/RBC/leukocyte esterase/nitrates; glucose and ketones positive.' : '',
      hasHepB ? 'HBsAg positive.' : '',
      hasGbs ? 'GBS positive.' : '',
      hasMagLevel ? 'Magnesium therapeutic range 4-7 mEq/L.' : '',
      hasCnsSigns ? 'CNS signs include hyperreflexia/headache/visual changes.' : '',
    ].filter(Boolean).join(' '), 140),
    labTestName: isPostpartumHemorrhage ? 'QBL / CBC / type-cross' : (isCordProlapse ? 'FHR / CBC / type-cross' : (isShoulderDystocia ? 'EFW / FHR / Apgar' : (isPreeclampsia ? 'BP / urine protein / Mg' : (hasUa || hasHepB || hasGbs ? 'UA / HBsAg / GBS' : '')))),
    labClientResults: compactText([
      isPostpartumHemorrhage ? 'QBL 800 mL; BP 90/50; HR 120; weak' : '',
      isCordProlapse ? 'Cord prolapse/variable decels; prep C-section' : '',
      isShoulderDystocia ? 'EFW >=4000 g; shoulder dystocia maneuvers' : '',
      isPreeclampsia ? 'Severe preeclampsia; assess proteinuria/CNS signs' : '',
      hasUa ? 'UA glucose/ketones positive; WBC/RBC/nitrates negative' : '',
      hasHepB ? 'HBsAg positive' : '',
      hasGbs ? 'GBS positive' : '',
    ].filter(Boolean).join('; '), 55),
    labNormalValue: isPostpartumHemorrhage ? 'QBL <500 mL vaginal; stable VS' : (isCordProlapse ? 'Reassuring FHR; intact cord position' : (isShoulderDystocia ? 'Uncomplicated delivery; reassuring newborn' : (isPreeclampsia ? 'BP <140/90; no proteinuria; Mg 4-7' : 'UA negative; HBsAg negative; GBS negative'))),
    labInterpretation: isPostpartumHemorrhage
      ? 'Findings support postpartum hemorrhage; monitor shock/perfusion.'
      : (isCordProlapse
      ? 'Cord prolapse threatens fetal oxygenation; urgent delivery needed.'
      : (isShoulderDystocia
      ? 'Shoulder dystocia increases fetal compromise/birth injury risk.'
      : (isPreeclampsia
      ? 'Findings support severe preeclampsia; monitor seizure risk and perfusion.'
      : (hasGbs
        ? 'GBS requires intrapartum antibiotic prophylaxis; monitor maternal/fetal status.'
        : 'Review abnormal results with care team.')))),
    currentMedDate: date,
    currentMedOrder: isPostpartumHemorrhage ? buildMedicationOrderSummary(ensurePopulatedMeds(meds), 165) : (isPreeclampsia ? 'Magnesium sulfate IV per protocol' : (meds.length ? 'Penicillin G IV per labor protocol' : '')),
    currentMedIndication: isPostpartumHemorrhage ? 'PPH bleeding/uterine tone management' : (isPreeclampsia ? 'Seizure prophylaxis' : (meds.length ? 'GBS intrapartum prophylaxis' : '')),
    nursingDiagnosis: '',
    goal: '',
    plan: '',
    intervention: '',
    rationale: '',
    evaluation: '',
    reassessment: '',
    theorist: '',
    knowledgeGained: '',
    courseObjectives: '',
    ...priorityFields,
  };

  return { fields, medications: capMedicationRows(meds), rawText: text };
}

function buildEducationalNeedsSuggestion(fields, rawText = '') {
  const context = `${fields.diagnosis || ''} ${fields.psychosocial || ''} ${fields.pain || ''} ${rawText}`.toLowerCase();

  if (/psychi|anxiety|depress|hallucinat|paranoi|agitat|grief|mood/.test(context)) {
    return 'Provide education on symptom monitoring, coping strategies, medication adherence, de-escalation supports, and when to seek urgent mental health care.';
  }
  if (/pain|chronic/.test(context)) {
    return 'Provide education on pain reporting, medication adherence, non-pharmacologic pain strategies, and follow-up with the care team.';
  }
  return 'Provide education on diagnosis understanding, medication adherence, warning signs to report, and follow-up plan.';
}

function buildDiagnosisMatchedVitals(fields = {}, rawText = '') {
  const context = clean(`${fields.diagnosis || ''} ${fields.nursingDiagnosis || ''} ${fields.psychosocial || ''} ${fields.respiratory || ''} ${fields.pain || ''} ${rawText || ''}`).toLowerCase();

  const has = (pattern) => pattern.test(context);
  const picks = {
    base: { temp: '98.4 F', pulse: '78 bpm', resp: '16/min', bp: '122/76 mmHg', spo2: '98% RA' },
    anxietyPsych: { temp: '99.1 F', pulse: '104 bpm', resp: '22/min', bp: '144/90 mmHg', spo2: '96% RA' },
    depression: { temp: '98.0 F', pulse: '64 bpm', resp: '14/min', bp: '110/70 mmHg', spo2: '98% RA' },
    infection: { temp: '101.3 F', pulse: '112 bpm', resp: '24/min', bp: '132/84 mmHg', spo2: '94% RA' },
    respiratory: { temp: '99.0 F', pulse: '102 bpm', resp: '26/min', bp: '138/84 mmHg', spo2: '91% RA' },
    cardiac: { temp: '98.7 F', pulse: '96 bpm', resp: '22/min', bp: '158/96 mmHg', spo2: '94% RA' },
    pain: { temp: '98.8 F', pulse: '102 bpm', resp: '20/min', bp: '150/92 mmHg', spo2: '97% RA' },
    endocrine: { temp: '99.4 F', pulse: '108 bpm', resp: '24/min', bp: '134/82 mmHg', spo2: '96% RA' },
  };

  let selected = picks.base;
  let reason = 'general adult profile';

  if (has(/pneumonia|sepsis|infection|infectious|uti|pyelo/)) {
    selected = picks.infection;
    reason = 'infectious/inflammatory diagnosis pattern';
  } else if (has(/copd|asthma|respiratory failure|hypoxia|hypoxemia|dyspnea|shortness of breath/)) {
    selected = picks.respiratory;
    reason = 'respiratory diagnosis pattern';
  } else if (has(/heart failure|chf|hypertension|htn|cardiac|arrhythm|angina|mi|coronary/)) {
    selected = picks.cardiac;
    reason = 'cardiovascular diagnosis pattern';
  } else if (has(/dka|diabetes|hyperglyc|hypoglyc|endocrine/)) {
    selected = picks.endocrine;
    reason = 'endocrine/metabolic diagnosis pattern';
  } else if (has(/pain|post[-\s]?op|postoperative|trauma|fracture|injury/)) {
    selected = picks.pain;
    reason = 'pain/stress diagnosis pattern';
  } else if (has(/anxiety|panic|agitat|psychosis|schizo|mania|bipolar|withdrawal|hallucinat|paranoi/)) {
    selected = picks.anxietyPsych;
    reason = 'psychiatric activation diagnosis pattern';
  } else if (has(/depress|grief|sadness|fatigue/)) {
    selected = picks.depression;
    reason = 'depressive symptom pattern';
  }

  return `Temp: ${selected.temp}; Pulse: ${selected.pulse}; Resp: ${selected.resp}; BP: ${selected.bp}; SpO2: ${selected.spo2}.`;
}

function extractPainScore(text = '') {
  const source = clean(text);
  const patterns = [
    /(?:pain|pain score|init\.?\s*pain score)\s*(?:reported|score|:|-)?\s*(\d{1,2}(?:\s*[-/]\s*\d{1,2})?)(?:\s*\/\s*10)?/i,
    /\bmild pain\s*\(?(\d{1,2}(?:\s*[-/]\s*\d{1,2})?)\)?(?:\s*\/\s*10)?/i,
  ];
  for (const pattern of patterns) {
    const value = source.match(pattern)?.[1];
    if (value) {
      const normalized = value.replace(/\s+/g, '');
      return normalized.includes('/') ? normalized : `${normalized}/10`;
    }
  }
  return '';
}

function extractVitalsFromSummary(text = '') {
  const source = clean(String(text || ''));
  if (!source) return {};

  const pickFirst = (patterns) => {
    for (const pattern of patterns) {
      const match = source.match(pattern);
      if (match?.[1]) return clean(match[1]);
    }
    return '';
  };

  const temp = pickFirst([
    /(?:temp(?:erature)?|\bt\b)\s*[:=-]?\s*([0-9]{2,3}(?:\.[0-9])?\s*(?:°?\s*[fc])?)/i,
  ]);
  const pulse = pickFirst([
    /(?:pulse|heart\s*rate|\bhr\b)\s*[:=-]?\s*([0-9]{2,3}\s*(?:bpm)?)/i,
  ]);
  const resp = pickFirst([
    /(?:resp(?:iratory)?(?:\s*rate)?|\brr\b)\s*[:=-]?\s*([0-9]{1,2}\s*(?:\/min|rpm|breaths?\/?min)?)/i,
  ]);
  const bp = pickFirst([
    /(?:blood\s*pressure|\bbp\b)\s*[:=-]?\s*([0-9]{2,3}\s*\/\s*[0-9]{2,3}(?:\s*mmhg)?)/i,
    /\b([0-9]{2,3}\s*\/\s*[0-9]{2,3})\b/,
  ]);
  const spo2 = pickFirst([
    /(?:spo2|o2\s*sat(?:uration)?|oxygen\s*saturation)\s*[:=-]?\s*([0-9]{2,3}\s*%\s*(?:ra|room air|on\s*\d+\s*l[^\s,;]*)?)/i,
  ]);

  return { temp, pulse, resp, bp, spo2 };
}

function getAgeNumber(fields = {}, rawText = '') {
  const value = fields.age || fields.patientInitialAge || extractAgeFromText(rawText);
  const age = parseInt(String(value || '').match(/\d{1,3}/)?.[0] || '', 10);
  return Number.isNaN(age) ? null : age;
}

function inferCaseProfile(fields = {}, rawText = '') {
  const source = clean(`${fields.diagnosis || ''} ${fields.medicalHistory || ''} ${fields.psychosocial || ''} ${rawText || ''}`).toLowerCase();
  const age = getAgeNumber(fields, rawText);
  const has = (pattern) => pattern.test(source);

  return {
    age,
    olderAdult: age !== null && age >= 65,
    assistedCare: has(/assisted care|needs assistance|assist(?:ed)?\s+(?:with\s+)?(?:care|adls|ambulation|hygiene)|fall risk/),
    weakness: has(/weakness|fatigue|decondition|unsteady|difficulty ambulating/),
    lowBp: has(/low blood pressure|hypotension|hypotensive|orthostatic|dizziness|lightheaded|syncope|near syncope/),
    infection: has(/uti|urinary tract|infection|sepsis|wbc|bacteria|antibiotic|fever/),
    anemia: has(/anemia|hemoglobin|hgb|iron deficiency|ferrous|prbc|transfusion|iron\s+\d/),
    syncope: has(/\bsyncope\b|fainting|near syncope|dizziness|lightheaded/),
    cardiac: has(/atrial fibrillation|a[-\s]?fib|\bafib\b|chf|heart failure|hypertension|htn|cardiac|av block|eliquis|anticoagulant/),
    respiratory: has(/shortness of breath|\bsob\b|dyspnea|hypoxia|oxygen|spo2|copd|pneumonia/),
    pain: has(/\bpain\b|headache/),
    altered: has(/confusion|altered mental|disoriented|lethargic/),
    nutritionConcern: has(/poor appetite|malnour|weight loss|dehydration|npo|nausea|vomiting/),
    edema: has(/edema|swelling|pitting/),
    incision: has(/incision|wound|surgical site|dressing|staples|sutures/),
    iv: has(/\biv\b|intravenous|saline lock|peripheral line|left arm|right arm/),
  };
}

function buildLikelyVitals(fields = {}, rawText = '') {
  const profile = inferCaseProfile(fields, rawText);
  if (profile.lowBp && profile.olderAdult) {
    return { temp: '98.2 F', pulse: '86 bpm', resp: '18/min', bp: '96/58' };
  }
  if (profile.lowBp) {
    return { temp: '98.4 F', pulse: '84 bpm', resp: '18/min', bp: '98/60' };
  }
  if (profile.infection && profile.olderAdult) {
    return { temp: '99.8 F', pulse: '94 bpm', resp: '20/min', bp: '110/68' };
  }
  if (profile.infection) {
    return { temp: '100.4 F', pulse: '98 bpm', resp: '20/min', bp: '118/72' };
  }
  if (profile.respiratory) {
    return { temp: '98.8 F', pulse: '92 bpm', resp: '22/min', bp: '118/70' };
  }
  if (profile.cardiac && profile.olderAdult) {
    return { temp: '98.1 F', pulse: '88 bpm', resp: '18/min', bp: '112/66' };
  }
  if ((profile.anemia || profile.syncope) && profile.olderAdult) {
    return { temp: '98.1 F', pulse: '78 bpm', resp: '18/min', bp: '126/72' };
  }
  if (profile.olderAdult && profile.weakness) {
    return { temp: '98.0 F', pulse: '82 bpm', resp: '18/min', bp: '104/64' };
  }
  return { temp: '98.4 F', pulse: '78 bpm', resp: '16/min', bp: '122/76' };
}

function extractIvLocation(text = '') {
  const source = clean(text);
  if (/left\s+(?:arm|forearm|hand|antecubital|ac)\b/i.test(source)) return 'Left arm';
  if (/right\s+(?:arm|forearm|hand|antecubital|ac)\b/i.test(source)) return 'Right arm';
  if (/\biv\b|intravenous|saline lock|peripheral line/i.test(source)) return 'Peripheral IV';
  return '';
}

function ensurePopulatedFields(inputFields, rawText = '') {
  const fields = { ...inputFields };
  const fallbackByField = {
    studentName: FALLBACK_NA,
    date: FALLBACK_NA,
    week: FALLBACK_NA,
    patientInitialAge: FALLBACK_NA,
    diagnosis: FALLBACK_NOT_ASSESSED,
    allergy: FALLBACK_NOT_ASSESSED,
    neuro: FALLBACK_NOT_ASSESSED,
    cardio: FALLBACK_NOT_ASSESSED,
    respiratory: FALLBACK_NOT_ASSESSED,
    vitals: '',
    gi: FALLBACK_NOT_ASSESSED,
    gu: FALLBACK_NOT_ASSESSED,
    skin: FALLBACK_NOT_ASSESSED,
    fluidElectrolytes: FALLBACK_NOT_ASSESSED,
    pain: FALLBACK_NOT_ASSESSED,
    psychosocial: FALLBACK_NOT_ASSESSED,
    activity: FALLBACK_NOT_ASSESSED,
    cultural: FALLBACK_NOT_ASSESSED,
    educationNeeds: '',
    safety: FALLBACK_NOT_ASSESSED,
    diagnosticTests: FALLBACK_NOT_ASSESSED,
    labs: FALLBACK_NOT_ASSESSED,
    dischargePlan: FALLBACK_NOT_ASSESSED,
    nursingDiagnosis: 'Needs nursing diagnosis based on assessment findings.',
    goal: 'Set a measurable patient-centered goal based on current assessment.',
    plan: 'Develop plan of care based on diagnosis, safety priorities, and follow-up needs.',
    intervention: 'Document interventions performed during this encounter.',
    rationale: 'Add rationale linked to the selected intervention and patient response.',
    evaluation: 'Evaluate patient response to interventions and current status.',
    reassessment: 'Document reassessment timing and updated findings.',
    theorist: '',
    knowledgeGained: '',
    courseObjectives: '',
  };

  for (const [key] of FIELD_SECTIONS) {
    const current = clean(String(fields[key] || ''));
    if (!current) {
      if (key === 'educationNeeds') {
        fields[key] = buildEducationalNeedsSuggestion(fields, rawText);
      } else if (key === 'vitals') {
        fields[key] = buildDiagnosisMatchedVitals(fields, rawText);
      } else {
        fields[key] = Object.prototype.hasOwnProperty.call(fallbackByField, key)
          ? fallbackByField[key]
          : FALLBACK_NA;
      }
    }
  }

  const hasUsableVitals = clean(String(fields.vitals || '')) && !/^(not assessed in case log\.?|n\/?a)$/i.test(clean(String(fields.vitals || '')));
  if (!hasUsableVitals) {
    fields.vitals = buildDiagnosisMatchedVitals(fields, rawText);
  }

  return fields;
}

function ensureConceptMapFields(inputFields, rawText = '', inputMeds = []) {
  const fields = { ...inputFields };
  const setIfBlank = (key, value) => {
    if (clean(fields[key])) return;
    const next = clean(value);
    if (next && !/^need to verify/i.test(next)) fields[key] = next;
  };
  const text = clean(rawText);
  const profile = inferCaseProfile(fields, text);
  const chartVitals = extractVitalsFromSummary(text);
  const likelyVitals = buildLikelyVitals(fields, text);

  setIfBlank('age', fields.patientInitialAge);
  setIfBlank('clientName', 'J.D.');
  setIfBlank('diagnosis', extractSimpleDiagnosis(fields.diagnosis, text));
  setIfBlank('ht', extractHeightFromText(text));
  setIfBlank('wt', extractWeightFromText(text));
  setIfBlank('temp', chartVitals.temp || likelyVitals.temp);
  setIfBlank('pulse', chartVitals.pulse || likelyVitals.pulse);
  setIfBlank('resp', chartVitals.resp || likelyVitals.resp);
  setIfBlank('bp', chartVitals.bp || likelyVitals.bp);
  setIfBlank('pain', extractPainScore(text));
  setIfBlank('sex', extractSexFromText(text));
  setIfBlank('medicalHistory', extractMedicalHistorySummary(text));
  setIfBlank('allergies', extractNarrativeAllergy(text) || (/checked allergies|allerg(?:y|ies)\s*(?:checked|reviewed)/i.test(text) ? 'None' : ''));
  setIfBlank('immunizations', 'Up to date');
  setIfBlank('sleepPattern', /insomnia|poor sleep|sleep disturbance|difficulty sleeping/i.test(text) ? 'Irregular sleep pattern' : 'Normal sleep pattern');
  setIfBlank('nutritionalStatus', profile.nutritionConcern ? 'Malnourished' : 'Nourished');
  setIfBlank('assistanceAdls', (profile.assistedCare || profile.weakness || profile.olderAdult) ? 'Needs assistance' : 'Independent');
  setIfBlank('hygiene', /assist(?:ed)?\s+hygiene|hygiene assistance/i.test(text) ? 'Assisted hygiene' : 'Independent');
  setIfBlank('surgicalHistory', 'None');
  setIfBlank('supportSystem', /family|daughter|son|spouse|husband|wife|caregiver/i.test(text) ? 'Family/caregiver involved' : '');
  setIfBlank('responseHospitalization', profile.altered ? 'Requires reorientation/support' : 'Cooperative with care');
  setIfBlank('genAppearance', /no acute distress/i.test(text) ? 'No acute distress' : (profile.olderAdult ? (profile.weakness || profile.anemia ? 'Older adult, weak/fatigued' : 'Older adult, no acute distress') : 'No acute distress'));
  setIfBlank('ivLocation', extractIvLocation(text) || (profile.iv ? 'Peripheral IV' : 'None'));
  setIfBlank('surgicalIncision', profile.incision ? 'Assess site/dressing' : 'None');
  setIfBlank('orientation', /alert and oriented x?4|oriented x?4|a&o x?4/i.test(text) ? 'A&O x4' : (/alert and oriented x?3|oriented x?3|a&o x?3/i.test(text) ? 'A&O x3' : ''));
  setIfBlank('speech', profile.altered ? 'Delayed' : 'Clear');
  setIfBlank('weakness', profile.weakness ? 'Generalized weakness' : 'None noted');
  setIfBlank('skinTurgor', profile.lowBp || profile.nutritionConcern ? 'Decreased' : 'Normal');
  setIfBlank('breathSounds', /lungs are clear|lungs clear|clear bilaterally/i.test(text) ? 'Clear' : (profile.respiratory ? 'SOB reported; assess' : 'Clear'));
  setIfBlank('peripheralPulses', 'Present');
  setIfBlank('edema', profile.edema ? 'Assess edema' : 'Need to assess');
  setIfBlank('bowelSounds', 'Present');
  setIfBlank('physicalOther', profile.olderAdult || profile.weakness ? 'Fall risk' : '');
  setIfBlank('spiritualAssessment', 'No needs stated');
  setIfBlank('labTestName', profile.anemia ? 'Hgb / iron studies' : (/wbc|bacteria|hematuria|urine/i.test(text) ? 'WBC / urinalysis' : ''));
  const hgbMatch = text.match(/(?:hemoglobin|hgb)\s*(?:was|of|:|=)?\s*(\d+(?:\.\d+)?)/i)?.[1];
  const ironMatch = text.match(/\biron\s*(\d+(?:\.\d+)?)/i)?.[1];
  setIfBlank('labClientResults', hgbMatch ? `Hgb ${hgbMatch}->~8; iron ${ironMatch || 'low'}` : (text.match(/WBC(?:\s*(?:of|:|=))?\s*(\d+(?:\.\d+)?)/i)?.[1] ? `WBC ${text.match(/WBC(?:\s*(?:of|:|=))?\s*(\d+(?:\.\d+)?)/i)?.[1]}; urine WBCs/bacteria; hematuria noted` : ''));
  setIfBlank('labNormalValue', profile.anemia ? 'Hgb M 13.5-17.5; iron 60-170' : 'WBC 4.5-11.0; UA normally negative for bacteria');
  setIfBlank('labInterpretation', profile.anemia ? 'Low Hgb/iron; monitor H/H and bleeding.' : (/uti|bacteria|hematuria|wbc/i.test(text) ? 'Findings support possible UTI/infection; hematuria requires monitoring.' : (profile.lowBp ? 'Monitor for hypotension/dehydration risk.' : 'Review abnormal results.')));
  setIfBlank('currentMedDate', fields.date);
  setIfBlank('currentMedOrder', ensurePopulatedMeds([]).length ? '' : '');
  const populatedMedRows = ensurePopulatedMeds(inputMeds);
  if (!clean(fields.medsPrior)) {
    const medNames = populatedMedRows
      .map((m) => clean(String(m.nameClass || '')).replace(/\s*;\s*Miscellaneous\b/gi, '').replace(/\bMiscellaneous\b/gi, '').trim())
      .filter(Boolean);
    if (medNames.length) fields.medsPrior = compactText(medNames.join('; '), 120);
  }
  if (!clean(fields.currentMedOrder)) {
    const medOrder = buildMedicationOrderSummary(populatedMedRows, 165);
    if (medOrder) fields.currentMedOrder = medOrder;
  }

  if (!clean(fields.currentMedOrder) && /valsartan|amlodipine|lasix|furosemide|eliquis|zofran|allopurinol|rosuvastatin|crestor|cinacalcet|cefdinir|thorazine|cholestyramine/i.test(text)) {
    fields.currentMedOrder = compactText('Valsartan; amlodipine; Lasix/furosemide; Eliquis; Zofran; allopurinol; rosuvastatin/Crestor; cinacalcet; cefdinir; Thorazine; cholestyramine.', 130);
  }
  if (!clean(fields.currentMedOrder) && /ferrous sulfate|esomeprazole|carvedilol/i.test(text)) {
    fields.currentMedOrder = compactText('Ferrous sulfate; esomeprazole; carvedilol.', 120);
  }
  setIfBlank('currentMedIndication', profile.anemia ? 'Treat anemia; GI protection; BP/HR control.' : (profile.infection ? 'Treat infection/UTI; monitor response.' : (profile.lowBp ? 'Monitor BP; support perfusion/safety.' : 'Manage active conditions.')));

  return fields;
}

function buildPriorityNursingFields(inputFields = {}, rawText = '') {
  const fields = { ...inputFields };
  const text = clean(rawText);
  const profile = inferCaseProfile(fields, text);
  const focusedDiagnoses = buildFocusedPriorityDiagnoses(fields, text, { includeManual: false });
  const safetyRisk = /fall|weakness|eliquis|anticoagulant/i.test(text) || profile.olderAdult || profile.syncope;

  return {
    nd1Assessment: withPriorityPrompt('nd1Assessment', profile.anemia ? 'Low Hgb/iron; fatigue or syncope risk.' : compactText(fields.labInterpretation || fields.diagnosis, 95)),
    nd1Diagnosis: withPriorityPrompt('nd1Diagnosis', focusedDiagnoses[0] || 'Impaired health maintenance'),
    nd1Rationale: withPriorityPrompt('nd1Rationale', profile.anemia ? 'Low Hgb can reduce oxygen delivery and activity tolerance.' : (profile.infection ? 'Elevated WBC/UA findings support infection monitoring.' : 'Assessment findings guide priority care.')),
    nd1Intervention: profile.anemia ? 'Monitor H/H, dizziness, bleeding signs, vitals, and response to treatment.' : (profile.infection ? 'Monitor temperature, WBC/UA trends, urine changes, and antibiotics.' : 'Monitor status and report clinical changes.'),
    nd1Evaluation: 'Response pending; continue reassessment.',

    nd2Assessment: withPriorityPrompt('nd2Assessment', safetyRisk ? 'Age, syncope, or weakness increases safety risk.' : 'Active condition requires focused monitoring.'),
    nd2Diagnosis: withPriorityPrompt('nd2Diagnosis', focusedDiagnoses[1] || 'Risk for falls'),
    nd2Rationale: withPriorityPrompt('nd2Rationale', safetyRisk ? 'Weakness or syncope increases injury risk if a fall occurs.' : 'Clinical changes can affect safety and recovery.'),
    nd2Intervention: safetyRisk ? 'Maintain fall precautions, assist ambulation, and monitor dizziness.' : 'Reassess symptoms, vitals, and safety needs.',
    nd2Evaluation: 'No injury noted; continue safety monitoring.',

    nd3Assessment: withPriorityPrompt('nd3Assessment', profile.respiratory ? 'SOB/fatigue reported.' : (profile.weakness ? 'Weakness limits activity tolerance.' : (fields.pain ? `Pain ${fields.pain}; monitor response.` : 'Patient needs continued reassessment.'))),
    nd3Diagnosis: withPriorityPrompt('nd3Diagnosis', focusedDiagnoses[2] || 'Activity intolerance'),
    nd3Rationale: withPriorityPrompt('nd3Rationale', profile.respiratory || profile.weakness ? 'Weakness/SOB can reduce activity tolerance.' : 'Symptoms require follow-up to guide care.'),
    nd3Intervention: profile.respiratory || profile.weakness ? 'Pace activity, assist ADLs, and monitor fatigue/respiratory status.' : 'Monitor response and update care plan.',
    nd3Evaluation: 'Tolerance requires ongoing reassessment.',
  };
}

function getReviewReasonForValue(value = '') {
  const text = clean(String(value || ''));
  if (!text) return 'Missing';
  if (isIntentionalNone(text)) return '';
  if (/not assessed in case log/i.test(text)) return 'Not in case log';
  if (/need to verify|need to assess|verify against|verify exact|verify in chart|verify in mar\/emr/i.test(text)) return 'Verify';
  if (/ai generated|estimate|estimated trend|patient did not specify|selected from class-level entry/i.test(text)) return 'AI generated';
  if (/needs nursing diagnosis|set a measurable|develop plan of care|document interventions|add rationale|evaluate patient response|document reassessment/i.test(text)) {
    return 'Draft placeholder';
  }
  return '';
}

function getFieldReviewItems(fields = {}) {
  return FIELD_SECTIONS
    .map(([key, label]) => {
      const reason = getReviewReasonForValue(fields[key]);
      return reason ? { key, label, reason } : null;
    })
    .filter(Boolean);
}

function getMedicationReviewItems(medications = []) {
  return medications
    .map((med, index) => {
      const reasons = MED_KEYS
        .map((key) => getReviewReasonForValue(med?.[key]))
        .filter(Boolean);
      if (med?.aiGenerated) reasons.push('AI generated');
      return reasons.length
        ? { index, reasons: Array.from(new Set(reasons)) }
        : null;
    })
    .filter(Boolean);
}

const CONCEPT_MAP_SOURCE_FIELDS = [
  ['studentName', 'Student Name'],
  ['date', 'Date'],
  ['clientName', 'Client Name'],
  ['age', 'Age'],
  ['sex', 'Sex'],
  ['ht', 'Height'],
  ['wt', 'Weight'],
  ['diagnosis', 'Diagnosis'],
  ['temp', 'Temperature'],
  ['pulse', 'Pulse'],
  ['resp', 'Respirations'],
  ['bp', 'Blood Pressure'],
  ['pain', 'Pain'],
  ['medicalHistory', 'Medical History'],
  ['surgicalHistory', 'Surgical History'],
  ['supportSystem', 'Support System / Significant Others'],
  ['responseHospitalization', 'Response to Current Hospitalization'],
  ['allergies', 'Allergies'],
  ['immunizations', 'Immunizations'],
  ['sleepPattern', 'Sleep Pattern'],
  ['nutritionalStatus', 'Nutritional Status'],
  ['assistanceAdls', 'Assistance with ADLs'],
  ['hygiene', 'Hygiene'],
  ['medsPrior', 'Medications Prior to Hospitalization'],
  ['genAppearance', 'General Appearance'],
  ['ivLocation', 'IV / Location'],
  ['surgicalIncision', 'Surgical Incision'],
  ['orientation', 'Orientation'],
  ['speech', 'Speech'],
  ['weakness', 'Weakness'],
  ['skinTurgor', 'Skin Turgor'],
  ['breathSounds', 'Breath Sounds'],
  ['peripheralPulses', 'Peripheral Pulses'],
  ['edema', 'Edema'],
  ['bowelSounds', 'Bowel Sounds'],
  ['physicalOther', 'Other Findings'],
  ['spiritualAssessment', 'Spiritual Assessment'],
  ['labTestName', 'Lab Test Name'],
  ['labClientResults', "Client's Results"],
  ['labNormalValue', 'Normal Value'],
  ['labInterpretation', 'Interpretation'],
  ['currentMedDate', 'Medication Date'],
  ['currentMedOrder', 'Order / Dosage / Frequency'],
  ['currentMedIndication', 'Indication'],
  ['nd1Assessment', 'Priority 1 Client Assessment'],
  ['nd1Diagnosis', 'Priority 1 Nursing Diagnosis'],
  ['nd1Rationale', 'Priority 1 Scientific Rationale'],
  ['nd1Intervention', 'Priority 1 Nursing Intervention / Action'],
  ['nd1Evaluation', 'Priority 1 Evaluation / Client Response'],
  ['nd2Assessment', 'Priority 2 Client Assessment'],
  ['nd2Diagnosis', 'Priority 2 Nursing Diagnosis'],
  ['nd2Rationale', 'Priority 2 Scientific Rationale'],
  ['nd2Intervention', 'Priority 2 Nursing Intervention / Action'],
  ['nd2Evaluation', 'Priority 2 Evaluation / Client Response'],
  ['nd3Assessment', 'Priority 3 Client Assessment'],
  ['nd3Diagnosis', 'Priority 3 Nursing Diagnosis'],
  ['nd3Rationale', 'Priority 3 Scientific Rationale'],
  ['nd3Intervention', 'Priority 3 Nursing Intervention / Action'],
  ['nd3Evaluation', 'Priority 3 Evaluation / Client Response'],
];

const PRIORITY_NURSING_FIELD_KEYS = [
  'nursingDiagnosis',
  'nd1Assessment',
  'nd1Diagnosis',
  'nd1Rationale',
  'nd1Intervention',
  'nd1Evaluation',
  'nd2Assessment',
  'nd2Diagnosis',
  'nd2Rationale',
  'nd2Intervention',
  'nd2Evaluation',
  'nd3Assessment',
  'nd3Diagnosis',
  'nd3Rationale',
  'nd3Intervention',
  'nd3Evaluation',
];

function clearPriorityNursingFields(fields = {}) {
  const next = { ...fields };
  PRIORITY_NURSING_FIELD_KEYS.forEach((key) => {
    next[key] = '';
  });
  return next;
}

const CONCEPT_MAP_FIELD_GROUPS = [
  {
    title: 'Basic Information',
    fields: [
      ['studentName', 'Student Name'],
      ['date', 'Date'],
      ['clientName', 'Client Name'],
      ['age', 'Age'],
      ['sex', 'Sex'],
      ['ht', 'Height'],
      ['wt', 'Weight'],
      ['diagnosis', 'Diagnosis'],
    ],
  },
  {
    title: 'Vital Signs',
    fields: [
      ['temp', 'Temperature'],
      ['pulse', 'Pulse'],
      ['resp', 'Respirations'],
      ['bp', 'Blood Pressure'],
      ['pain', 'Pain'],
    ],
  },
  {
    title: 'Health History - Daily Needs',
    fields: [
      ['allergies', 'Allergies'],
      ['immunizations', 'Immunizations'],
      ['sleepPattern', 'Sleep Pattern'],
      ['nutritionalStatus', 'Nutritional Status'],
      ['assistanceAdls', 'Assistance with ADLs'],
      ['hygiene', 'Hygiene'],
      ['medsPrior', 'Medications Prior to Hospitalization'],
    ],
  },
  {
    title: 'Health History - Medical Background',
    fields: [
      ['medicalHistory', 'Medical History'],
      ['surgicalHistory', 'Surgical History'],
      ['supportSystem', 'Support System / Significant Others'],
      ['responseHospitalization', 'Response to Current Hospitalization'],
    ],
  },
  {
    title: 'Physical Assessment',
    fields: [
      ['genAppearance', 'General Appearance'],
      ['ivLocation', 'IV / Location'],
      ['surgicalIncision', 'Surgical Incision'],
      ['orientation', 'Orientation'],
      ['speech', 'Speech'],
      ['weakness', 'Weakness'],
      ['skinTurgor', 'Skin Turgor'],
      ['breathSounds', 'Breath Sounds'],
      ['peripheralPulses', 'Peripheral Pulses'],
      ['edema', 'Edema'],
      ['bowelSounds', 'Bowel Sounds'],
      ['physicalOther', 'Other Findings'],
    ],
  },
  {
    title: 'Spiritual Assessment',
    fields: [
      ['spiritualAssessment', 'S. Assessment'],
    ],
  },
  {
    title: 'Laboratory and Diagnostic Studies',
    fields: [
      ['labTestName', 'Test Name'],
      ['labClientResults', "Client's Results"],
      ['labNormalValue', 'Normal Value'],
      ['labInterpretation', 'Interpretation'],
    ],
  },
  {
    title: 'Current Medications / Therapeutic Plan',
    fields: [
      ['currentMedDate', 'Date'],
      ['currentMedOrder', 'Order / Dosage / Frequency'],
      ['currentMedIndication', 'Indication'],
    ],
  },
  {
    title: 'Priority Nursing Diagnosis #1',
    fields: [
      ['nd1Diagnosis', 'Nursing Diagnosis'],
      ['nd1Assessment', 'Client Assessment'],
      ['nd1Rationale', 'Scientific Rationale'],
      ['nd1Intervention', 'Nursing Intervention / Action'],
      ['nd1Evaluation', 'Evaluation / Client Response'],
    ],
  },
  {
    title: 'Priority Nursing Diagnosis #2',
    fields: [
      ['nd2Diagnosis', 'Nursing Diagnosis'],
      ['nd2Assessment', 'Client Assessment'],
      ['nd2Rationale', 'Scientific Rationale'],
      ['nd2Intervention', 'Nursing Intervention / Action'],
      ['nd2Evaluation', 'Evaluation / Client Response'],
    ],
  },
  {
    title: 'Priority Nursing Diagnosis #3',
    fields: [
      ['nd3Diagnosis', 'Nursing Diagnosis'],
      ['nd3Assessment', 'Client Assessment'],
      ['nd3Rationale', 'Scientific Rationale'],
      ['nd3Intervention', 'Nursing Intervention / Action'],
      ['nd3Evaluation', 'Evaluation / Client Response'],
    ],
  },
];

const isPriorityNursingGroup = (group) => /^Priority Nursing Diagnosis/i.test(group?.title || '');
const isPriorityNursingKey = (key = '') => /^nd[123]/.test(String(key || ''));
const MAIN_CONCEPT_MAP_FIELD_GROUPS = CONCEPT_MAP_FIELD_GROUPS.filter((group) => !isPriorityNursingGroup(group));
const PRIORITY_NURSING_FIELD_GROUPS = CONCEPT_MAP_FIELD_GROUPS.filter(isPriorityNursingGroup);

const CONCEPT_MAP_OUTPUT_LABELS = {
  studentName: 'Student Name',
  instructor: 'Instructor',
  date: 'Date',
  clientName: 'Client Name',
  age: 'Age',
  sex: 'Sex',
  ht: 'Height',
  wt: 'Weight',
  diagnosis: 'Diagnosis',
  temp: 'Temperature',
  pulse: 'Pulse',
  resp: 'Respirations',
  bp: 'Blood Pressure',
  pain: 'Pain',
  medicalHistory: 'Medical History',
  surgicalHistory: 'Surgical History',
  supportSystem: 'Support System',
  responseHospitalization: 'Response to Hospitalization',
  allergiesImmunization: 'Allergies / Immunization',
  sleepNutritional: 'Sleep / Nutrition',
  assistanceAdls: 'Assistance with ADLs',
  hygiene: 'Hygiene',
  medsPrior: 'Medications Prior',
  genAppearance: 'General Appearance',
  ivLocation: 'IV Location',
  surgicalIncision: 'Surgical Incision',
  orientation: 'Orientation',
  speech: 'Speech',
  weakness: 'Weakness',
  skinTurgor: 'Skin Turgor',
  breathSounds: 'Breath Sounds',
  peripheralPulses: 'Peripheral Pulses',
  edema: 'Edema',
  bowelSounds: 'Bowel Sounds',
  physicalOther: 'Other Physical Findings',
  spiritualAssessment: 'Spiritual Assessment',
  currentMedsPlan: 'Current Medications',
  nd1Assessment: 'Priority 1 Assessment',
  nd1Diagnosis: 'Priority 1 Diagnosis',
  nd1Rationale: 'Priority 1 Rationale',
  nd1Intervention: 'Priority 1 Intervention',
  nd1Evaluation: 'Priority 1 Evaluation',
  nd2Assessment: 'Priority 2 Assessment',
  nd2Diagnosis: 'Priority 2 Diagnosis',
  nd2Rationale: 'Priority 2 Rationale',
  nd2Intervention: 'Priority 2 Intervention',
  nd2Evaluation: 'Priority 2 Evaluation',
  nd3Assessment: 'Priority 3 Assessment',
  nd3Diagnosis: 'Priority 3 Diagnosis',
  nd3Rationale: 'Priority 3 Rationale',
  nd3Intervention: 'Priority 3 Intervention',
  nd3Evaluation: 'Priority 3 Evaluation',
  labTestName: 'Lab Test Name',
  labClientResults: 'Client Lab Results',
  labNormalValue: 'Normal Lab Value',
  labInterpretation: 'Lab Interpretation',
  currentMedDate: 'Current Med Date',
  currentMedOrder: 'Current Med Order',
  currentMedIndication: 'Current Med Indication',
};

const CONCEPT_MAP_FIELD_LIMITS = {
  studentName: 45,
  date: 24,
  clientName: 24,
  age: 10,
  sex: 8,
  ht: 12,
  wt: 12,
  diagnosis: 80,
  temp: 35,
  pulse: 35,
  resp: 35,
  bp: 35,
  pain: 45,
  medicalHistory: 115,
  surgicalHistory: 95,
  supportSystem: 90,
  responseHospitalization: 90,
  allergies: 26,
  immunizations: 28,
  sleepPattern: 30,
  nutritionalStatus: 28,
  assistanceAdls: 36,
  hygiene: 36,
  medsPrior: 75,
  genAppearance: 42,
  ivLocation: 28,
  surgicalIncision: 24,
  orientation: 30,
  speech: 24,
  weakness: 22,
  skinTurgor: 28,
  breathSounds: 28,
  peripheralPulses: 30,
  edema: 26,
  bowelSounds: 30,
  physicalOther: 26,
  spiritualAssessment: 52,
  labTestName: 50,
  labClientResults: 55,
  labNormalValue: 45,
  labInterpretation: 70,
  currentMedDate: 30,
  currentMedOrder: 165,
  currentMedIndication: 80,
  nd1Assessment: 115,
  nd1Diagnosis: 42,
  nd1Rationale: 115,
  nd1Intervention: 130,
  nd1Evaluation: 115,
  nd2Assessment: 115,
  nd2Diagnosis: 42,
  nd2Rationale: 115,
  nd2Intervention: 130,
  nd2Evaluation: 115,
  nd3Assessment: 115,
  nd3Diagnosis: 42,
  nd3Rationale: 115,
  nd3Intervention: 130,
  nd3Evaluation: 115,
};

const PRIORITY_FIELD_PROMPTS = {
  nd1Assessment: 'As Evidenced By:',
  nd2Assessment: 'As Evidenced By:',
  nd3Assessment: 'As Evidenced By:',
  nd1Diagnosis: 'Priority Problem:',
  nd2Diagnosis: 'Priority Problem:',
  nd3Diagnosis: 'Priority Problem:',
  nd1Rationale: 'Related to change in:',
  nd2Rationale: 'Related to change in:',
  nd3Rationale: 'Related to change in:',
};

function withPriorityPrompt(key, value = '') {
  const prompt = PRIORITY_FIELD_PROMPTS[key];
  const text = clean(value);
  if (!prompt || !text) return text;
  if (text.toLowerCase().startsWith(prompt.toLowerCase())) return text;
  return compactText(`${prompt} ${text}`, CONCEPT_MAP_FIELD_LIMITS[key] || 150);
}

function toShortNursingDiagnosis(value = '') {
  const text = clean(String(value || '')).replace(/\.$/, '');
  if (!text) return '';
  if (/severe\s+anemia|low\s+hgb|low\s+hemoglobin|low\s+iron/i.test(text)) return 'Ineffective tissue perfusion';
  if (/risk\s+for\s+falls?/i.test(text)) return 'Risk for falls';
  if (/risk\s+for\s+infection/i.test(text)) return 'Risk for infection';
  if (/risk\s+for\s+bleed/i.test(text)) return 'Risk for bleeding';
  if (/open\s+sores?|wound|ulcer|skin\s+breakdown/i.test(text)) return 'Impaired skin integrity';
  if (/impaired\s+skin/i.test(text)) return 'Impaired skin integrity';
  if (/impaired\s+urinary/i.test(text)) return 'Impaired urinary elimination';
  if (/ineffective\s+tissue\s+perfusion|decreased\s+tissue\s+perfusion/i.test(text)) return 'Ineffective tissue perfusion';
  if (/decreased\s+cardiac\s+output/i.test(text)) return 'Decreased cardiac output';
  if (/activity\s+intolerance/i.test(text)) return 'Activity intolerance';
  if (/acute\s+pain|pain/i.test(text)) return 'Acute pain';
  if (/fatigue/i.test(text)) return 'Fatigue';
  if (/anemia|hgb|hemoglobin|syncope/i.test(text)) return 'Ineffective tissue perfusion';
  if (/hypotension|low\s+blood\s+pressure|cardiac/i.test(text)) return 'Decreased cardiac output';
  if (/shortness\s+of\s+breath|\bsob\b|dyspnea/i.test(text)) return 'Activity intolerance';
  if (/infection|uti/i.test(text)) return 'Risk for infection';
  if (/weakness|mobility/i.test(text)) return 'Impaired physical mobility';
  return compactText(text.split(/\s+related\s+to\s+/i)[0], 42);
}

function nursingDiagnosisFocus(value = '') {
  const text = toShortNursingDiagnosis(value).toLowerCase();
  if (!text) return '';
  if (/falls?/.test(text)) return 'safety';
  if (/bleeding/.test(text)) return 'bleeding';
  if (/infection/.test(text)) return 'infection';
  if (/skin/.test(text)) return 'skin';
  if (/urinary/.test(text)) return 'urinary';
  if (/tissue perfusion/.test(text)) return 'perfusion';
  if (/cardiac output/.test(text)) return 'cardiac';
  if (/activity intolerance/.test(text)) return 'activity';
  if (/physical mobility/.test(text)) return 'mobility';
  if (/pain/.test(text)) return 'pain';
  if (/fatigue/.test(text)) return 'fatigue';
  return text;
}

function isRiskNursingDiagnosis(value = '') {
  return /^risk\s+for\b/i.test(toShortNursingDiagnosis(value));
}

function buildFocusedPriorityDiagnoses(fields = {}, rawText = '', options = {}) {
  const { includeManual = true } = options;
  const profile = inferCaseProfile(fields, rawText);
  const text = clean(`${fields.diagnosis || ''} ${fields.medicalHistory || ''} ${fields.psychosocial || ''} ${rawText || ''}`).toLowerCase();
  const hasFallRisk = /fall|weakness|eliquis|anticoagulant|syncope|dizziness|unsteady/i.test(text) || profile.olderAdult || profile.syncope;
  const hasSkinIssue = /open\s+sores?|wound|ulcer|skin\s+breakdown|pressure injury|dressing|incision/i.test(text) || profile.incision;
  const hasBleedingRisk = /eliquis|apixaban|anticoagulant|hematuria|bleeding/i.test(text);

  const candidates = [];
  if (includeManual) {
    candidates.push(fields.nd1Diagnosis, fields.nd2Diagnosis, fields.nd3Diagnosis, fields.nursingDiagnosis);
  }
  if (hasSkinIssue) candidates.push('Impaired skin integrity');
  if (profile.infection) candidates.push('Impaired urinary elimination');
  if (profile.anemia) candidates.push('Ineffective tissue perfusion', 'Fatigue');
  if (profile.lowBp || profile.cardiac) candidates.push('Decreased cardiac output');
  if (profile.pain) candidates.push('Acute pain');
  if (hasFallRisk) candidates.push('Risk for falls');
  if (profile.infection) candidates.push('Risk for infection');
  if (hasBleedingRisk) candidates.push('Risk for bleeding');
  if (profile.respiratory || profile.weakness) candidates.push('Activity intolerance');
  if (profile.weakness && !profile.respiratory) candidates.push('Impaired physical mobility');
  candidates.push(fields.diagnosis);
  candidates.push(
    'Impaired urinary elimination',
    'Risk for falls',
    'Activity intolerance',
    'Ineffective tissue perfusion',
    'Acute pain',
    'Impaired skin integrity',
    'Fatigue',
    'Decreased cardiac output'
  );

  const selected = [];
  const usedLabels = new Set();
  const usedFocus = new Set();
  let riskUsed = false;

  for (const candidate of candidates) {
    const short = toShortNursingDiagnosis(candidate);
    if (!short || /^need to verify|^not assessed/i.test(short) || isIntentionalNone(short)) continue;
    const key = short.toLowerCase();
    const focus = nursingDiagnosisFocus(short);
    if (usedLabels.has(key) || usedFocus.has(focus)) continue;
    if (isRiskNursingDiagnosis(short)) {
      if (riskUsed) continue;
      riskUsed = true;
    }
    selected.push(short);
    usedLabels.add(key);
    usedFocus.add(focus);
    if (selected.length === 3) break;
  }

  return selected;
}

function getMedicationClassProfile(nameClass = '') {
  const text = clean(String(nameClass || '')).toLowerCase();
  const profiles = [
    {
      pattern: /benzodiazep|lorazepam|clonazepam|diazepam|alprazolam/,
      representative: 'Lorazepam',
      category: 'benzodiazepine',
      names: ['lorazepam', 'clonazepam', 'diazepam', 'alprazolam'],
      dose: 'PO/IV 1-3x daily',
      why: 'anxiety/agitation control, acute distress reduction, or short-term sedation support',
      action: 'enhances GABA-A activity in the CNS, increasing inhibitory neurotransmission and reducing neuronal excitability',
      implications: 'monitor sedation level, respiratory status, fall risk, and paradoxical agitation; avoid abrupt discontinuation if prolonged use',
      sideEffects: 'drowsiness, dizziness, impaired coordination, confusion, respiratory depression (higher risk with other CNS depressants)',
    },
    {
      pattern: /antipsychotic|risperidone|olanzapine|quetiapine|haloperidol|aripiprazole/,
      representative: 'Risperidone',
      category: 'antipsychotic',
      names: ['risperidone', 'olanzapine', 'quetiapine', 'haloperidol', 'aripiprazole'],
      dose: 'PO daily-BID',
      why: 'psychosis, severe agitation, or mood/thought stabilization',
      action: 'modulates central dopamine (and in many agents serotonin) signaling to reduce psychotic and behavioral symptoms',
      implications: 'monitor EPS/tardive symptoms, orthostatic changes, sedation, and metabolic trends; reinforce adherence and safety monitoring',
      sideEffects: 'sedation, weight gain, orthostatic hypotension, extrapyramidal symptoms, anticholinergic effects, metabolic changes',
    },
    {
      pattern: /antidepressant|ssri|snri|sertraline|escitalopram|fluoxetine|venlafaxine|duloxetine/,
      representative: 'Sertraline',
      category: 'SSRI antidepressant',
      names: ['sertraline', 'escitalopram', 'fluoxetine', 'venlafaxine', 'duloxetine'],
      dose: 'PO daily',
      why: 'depressive or anxiety-spectrum symptoms and mood stabilization support',
      action: 'inhibits serotonin and/or norepinephrine reuptake, increasing synaptic monoamine availability',
      implications: 'monitor mood/suicidality trends, sleep/appetite changes, and adherence; educate that therapeutic effect can be delayed',
      sideEffects: 'nausea, headache, insomnia or somnolence, sexual dysfunction, GI upset, possible activation/anxiety early in therapy',
    },
    {
      pattern: /mood stabil|lithium|valpro|lamotrigine|carbamazepine/,
      representative: 'Lamotrigine',
      category: 'mood stabilizer',
      names: ['lithium', 'valproate', 'lamotrigine', 'carbamazepine'],
      dose: 'PO daily-BID',
      why: 'mood stabilization and reduction of affective lability',
      action: 'modulates neuronal signaling pathways to reduce mood cycling and excitatory instability',
      implications: 'monitor neurologic status, hydration, and ordered therapeutic labs where applicable; assess for toxicity warning signs',
      sideEffects: 'tremor, GI upset, sedation, dizziness, weight or appetite changes, concentration/cognitive slowing',
    },
    {
      pattern: /opioid|morphine|hydromorphone|oxycodone|fentanyl|tramadol/,
      representative: 'Oxycodone',
      category: 'opioid analgesic',
      names: ['morphine', 'hydromorphone', 'oxycodone', 'fentanyl', 'tramadol'],
      dose: 'PO/IV q4-12h',
      why: 'moderate-to-severe pain management',
      action: 'agonizes central opioid receptors, decreasing pain perception and pain response',
      implications: 'monitor pain response, respiratory status, sedation, bowel function, and safety/fall risk; use opioid precautions',
      sideEffects: 'constipation, nausea, sedation, dizziness, respiratory depression, pruritus',
    },
    {
      pattern: /antibiotic|penicillin|amoxicillin|cef|azithromycin|doxycycline|levofloxacin/,
      representative: 'Amoxicillin',
      category: 'antibiotic',
      names: ['penicillin', 'amoxicillin', 'cef', 'azithromycin', 'doxycycline', 'levofloxacin'],
      dose: 'PO/IV q8-24h',
      why: 'suspected or confirmed bacterial infection management',
      action: 'inhibits bacterial growth or viability through class-specific antimicrobial mechanisms',
      implications: 'monitor infection trend, allergy reaction risk, GI tolerance, and adherence to full treatment course',
      sideEffects: 'GI upset, diarrhea, rash, hypersensitivity reactions, possible superinfection risk',
    },
    {
      pattern: /apixaban|eliquis|anticoagulant|antithrombotic/,
      representative: 'Apixaban',
      category: 'anticoagulant',
      names: ['apixaban', 'eliquis'],
      dose: 'PO BID',
      why: 'prevention or treatment of thromboembolic events',
      action: 'inhibits factor Xa to reduce clot formation',
      implications: 'monitor bleeding, bruising, hematuria, fall risk, and ordered anticoagulation precautions',
      sideEffects: 'bleeding, bruising, nausea, anemia, and rare hypersensitivity',
    },
    {
      pattern: /atorvastatin|rosuvastatin|simvastatin|statin|antilipemic/,
      representative: 'Atorvastatin',
      category: 'statin',
      names: ['atorvastatin', 'rosuvastatin', 'simvastatin'],
      dose: 'PO daily',
      why: 'hyperlipidemia and cardiovascular risk reduction',
      action: 'inhibits HMG-CoA reductase to reduce hepatic cholesterol synthesis',
      implications: 'monitor lipid therapy adherence, liver enzymes if ordered, and muscle pain/weakness',
      sideEffects: 'myalgia, GI upset, headache, elevated liver enzymes, rare rhabdomyolysis',
    },
    {
      pattern: /metoprolol|carvedilol|atenolol|beta blocker/,
      representative: 'Metoprolol',
      category: 'beta blocker',
      names: ['metoprolol', 'carvedilol', 'atenolol'],
      dose: 'PO daily-BID',
      why: 'blood pressure, heart rate, or cardiac workload control',
      action: 'blocks beta-adrenergic receptors to lower heart rate and cardiac workload',
      implications: 'monitor BP, heart rate, dizziness, fatigue, and hold parameters per order',
      sideEffects: 'bradycardia, hypotension, dizziness, fatigue, and weakness',
    },
    {
      pattern: /pantoprazole|omeprazole|esomeprazole|proton-pump|ppi|antiulcer|acid suppress/,
      representative: 'Pantoprazole',
      category: 'proton-pump inhibitor',
      names: ['pantoprazole', 'omeprazole', 'esomeprazole'],
      dose: 'PO/IV daily',
      why: 'GERD, ulcer prevention, or GI protection',
      action: 'suppresses gastric acid secretion by inhibiting the proton pump',
      implications: 'monitor GI symptoms, bleeding risk, and tolerance; give before meals if ordered',
      sideEffects: 'headache, diarrhea, nausea, abdominal pain, and long-term mineral/vitamin effects',
    },
    {
      pattern: /ferrous|iron preparation|antianemia/,
      representative: 'Ferrous sulfate',
      category: 'iron supplement',
      names: ['ferrous sulfate'],
      dose: 'PO daily/every other day',
      why: 'iron deficiency anemia treatment',
      action: 'replaces iron needed for hemoglobin synthesis',
      implications: 'monitor Hgb/Hct, constipation, dark stools, GI upset, and adherence',
      sideEffects: 'constipation, dark stools, nausea, abdominal discomfort',
    },
    {
      pattern: /insulin|lispro|glargine|antidiabetic/,
      representative: 'Insulin lispro',
      category: 'insulin',
      names: ['insulin lispro', 'insulin glargine'],
      dose: 'SubQ sliding scale/scheduled',
      why: 'blood glucose control',
      action: 'promotes cellular glucose uptake and lowers blood glucose',
      implications: 'check blood glucose, meal timing, hypoglycemia signs, and ordered parameters',
      sideEffects: 'hypoglycemia, injection site reaction, weight gain, hypokalemia',
    },
    {
      pattern: /magnesium sulfate|magnesium/i,
      representative: 'Magnesium sulfate',
      category: 'anticonvulsant/mineral electrolyte',
      names: ['magnesium sulfate'],
      dose: 'IV per protocol',
      why: 'seizure prophylaxis for severe preeclampsia',
      action: 'depresses CNS excitability and reduces neuromuscular transmission to prevent seizures',
      implications: 'monitor respirations, DTRs, urine output, blood pressure, magnesium level, and calcium gluconate availability',
      sideEffects: 'flushing, warmth, nausea, muscle weakness, respiratory depression, hypotension, and loss of reflexes at toxic levels',
    },
    {
      pattern: /misoprostol|cytotec/,
      representative: 'Misoprostol',
      category: 'uterotonic prostaglandin',
      names: ['misoprostol', 'cytotec'],
      dose: 'Per PPH protocol',
      why: 'uterine tone support and postpartum bleeding control',
      action: 'stimulates uterine contraction to reduce bleeding from the placental site',
      implications: 'monitor uterine tone, bleeding amount, temperature, GI effects, pain, and contraindications per protocol',
      sideEffects: 'fever, chills, nausea, vomiting, diarrhea, abdominal cramping',
    },
    {
      pattern: /oxytocin|pitocin/,
      representative: 'Oxytocin',
      category: 'uterotonic hormone',
      names: ['oxytocin', 'pitocin'],
      dose: 'IV/IM per PPH protocol',
      why: 'uterine contraction and postpartum bleeding control',
      action: 'stimulates uterine smooth muscle contraction',
      implications: 'monitor uterine tone, lochia/bleeding, blood pressure, fluid balance, and response to therapy',
      sideEffects: 'uterine cramping, nausea, hypotension, tachycardia, water intoxication with high/prolonged dosing',
    },
  ];

  const profile = profiles.find((p) => p.pattern.test(text));
  if (!profile) return null;

  const hasSpecificName = profile.names.some((name) => text.includes(name));
  const label = hasSpecificName
    ? clean(String(nameClass || ''))
    : `${profile.representative} (${profile.category})`;

  return { ...profile, label };
}

function buildGeneralDoseRouteEstimate(nameClass = '') {
  const profile = getMedicationClassProfile(nameClass);
  if (profile) return profile.dose;
  const text = clean(String(nameClass || '')).toLowerCase();
  if (!text) return 'PO daily-BID';
  if (/opioid|analgesic|pain/i.test(text)) return 'PO/IV q4-12h';
  if (/antibiotic|antibacterial|cef|amoxicillin/i.test(text)) return 'PO/IV q8-24h';
  return 'PO daily-BID';
}

function getMedicationClassEstimate(nameClass = '') {
  const profile = getMedicationClassProfile(nameClass);
  if (profile) {
    return {
      why: profile.why,
      action: profile.action,
      implications: profile.implications,
      sideEffects: profile.sideEffects,
    };
  }

  return {
    why: 'symptom control and condition management per active treatment plan',
    action: 'produces class-specific therapeutic effects through receptor or enzyme pathway modulation',
    implications: 'monitor therapeutic response, adverse effects, vital sign impact, and medication safety before administration',
    sideEffects: 'dizziness, GI upset, sedation or activation, headache, and class-specific adverse reactions',
  };
}

function formatDoseForConceptMap(doseRoute = '') {
  return clean(String(doseRoute || ''))
    .replace(/\s*\(estimate\)/gi, '')
    .replace(/\s+/g, ' ');
}

function buildMedicationOrderSummary(medRows = [], maxLength = 165) {
  const rows = medRows
    .map((med) => ({
      name: compactText(clean(String(med?.nameClass || '')), 28),
      dose: compactText(formatDoseForConceptMap(med?.doseRoute), 32),
    }))
    .filter((med) => med.name);
  if (!rows.length) return '';

  const full = rows
    .map((med) => [med.name, med.dose].filter(Boolean).join(' '))
    .join('; ');
  if (full.length <= maxLength) return full;

  const namesOnly = rows.map((med) => med.name).join('; ');
  const withReviewNote = `${namesOnly}; review dose/freq`;
  if (withReviewNote.length <= maxLength) return withReviewNote;
  return compactText(withReviewNote, maxLength);
}

function ensurePopulatedMeds(inputMeds = []) {
  const source = Array.isArray(inputMeds) && inputMeds.length ? inputMeds : MED_DEFAULT;
  const meds = capMedicationRows(expandMedicationRows(source, MED_TEMPLATE_ROW_CAP), MED_TEMPLATE_ROW_CAP);
  return meds.map((med) => {
    const originalNameClass = clean(String(med?.nameClass || ''))
      .replace(/\s*;\s*Miscellaneous\b/gi, '')
      .replace(/\bMiscellaneous\b/gi, '')
      .trim() || FALLBACK_NA;
    const classProfile = getMedicationClassProfile(originalNameClass);
    const nameClass = classProfile?.label || originalNameClass;
    const doseRoute = clean(String(med?.doseRoute || ''));
    const why = clean(String(med?.why || ''));
    const action = clean(String(med?.action || ''));
    const implications = clean(String(med?.implications || ''));
    const sideEffects = clean(String(med?.sideEffects || ''));
    const generated = getMedicationClassEstimate(nameClass);
    const aiGenerated = !doseRoute || !why || !action || !implications || !sideEffects;

    return {
      nameClass,
      doseRoute: doseRoute || buildGeneralDoseRouteEstimate(nameClass),
      why: why || generated.why,
      action: action || generated.action,
      implications: implications || generated.implications,
      sideEffects: sideEffects || generated.sideEffects,
      aiGenerated,
    };
  });
}

function buildDocText(fields, meds) {
  const forDoc = (value) => {
    const normalized = stripOutputReviewMarkers(value);
    return normalized || FALLBACK_NA;
  };
  const forOptionalDoc = (value) => stripOutputReviewMarkers(value);

  return `Nursing Process Worksheet for Students\n\nStudent Name: ${forDoc(fields.studentName)}\nDate: ${forDoc(fields.date)}\nWeek #: ${forDoc(fields.week)}\nPatient Initial/Age: ${forDoc(fields.patientInitialAge)}\nPatient Diagnosis: ${forDoc(fields.diagnosis)}\nAllergy: ${forDoc(fields.allergy)}\n\nAssessment of Patient's Status and Needs\n\nNeuro: ${forDoc(fields.neuro)}\nCardiovascular & IV Infusions: ${forDoc(fields.cardio)}\nRespiratory: ${forDoc(fields.respiratory)}\nVital Signs: ${forDoc(fields.vitals)}\nGastrointestinal (GI): ${forDoc(fields.gi)}\nGenitourinary (GU): ${forDoc(fields.gu)}\nSkin Integrity: ${forDoc(fields.skin)}\nFluid & Electrolyte Status: ${forDoc(fields.fluidElectrolytes)}\nPain Management: ${forDoc(fields.pain)}\nPsychosocial/Developmental Status: ${forDoc(fields.psychosocial)}\nPrescribed Activity: ${forDoc(fields.activity)}\nCultural/Spiritual: ${forDoc(fields.cultural)}\nPt/Family Educational Needs: ${forDoc(fields.educationNeeds)}\nSpecific Safety Need/Precautions: ${forDoc(fields.safety)}\nDiagnostic Test or Procedure Scheduled: ${forDoc(fields.diagnosticTests)}\nRecent Lab Results: ${forDoc(fields.labs)}\nDischarge Plan: ${forDoc(fields.dischargePlan)}\n\nNursing Diagnosis: ${forDoc(fields.nursingDiagnosis)}\nNursing Goal: ${forDoc(fields.goal)}\nPlan of Care: ${forDoc(fields.plan)}\nNursing Care Given/Intervention: ${forDoc(fields.intervention)}\nRationale for Intervention: ${forDoc(fields.rationale)}\nEvaluation/Outcome: ${forDoc(fields.evaluation)}\nReassessment: ${forDoc(fields.reassessment)}\n\nExplain how you applied the work of a select nursing theorist in the care of this patient:\n${forOptionalDoc(fields.theorist)}\n\nKnowledge gained from this clinical experience:\n${forOptionalDoc(fields.knowledgeGained)}\n\nHow were the course/clinical objectives met through this clinical experience?\n${forOptionalDoc(fields.courseObjectives)}\n\nMedication Log\n\n${meds.map((m, i) => `${i + 1}. ${forDoc(m.nameClass)}\nDose/Frequency/Route: ${forDoc(m.doseRoute)}\nWhy client is receiving this drug: ${forDoc(m.why)}\nAction of Drug: ${forDoc(m.action)}\nNursing Implications: ${forDoc(m.implications)}\nSide Effects: ${forDoc(m.sideEffects)}`).join('\n\n')}`;
}

function downloadFile(filename, content, type = 'text/plain;charset=utf-8') {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

async function saveBlobAsFile(blob, filename, mimeType) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function buildDatedOutputFilename(inputDate = '') {
  const raw = String(inputDate || '').trim();
  const parsed = raw ? new Date(raw) : null;
  const dateObj = parsed && !Number.isNaN(parsed.getTime()) ? parsed : new Date();
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `nursing-process-worksheet-${y}-${m}-${d}.docx`;
}

function buildDatedArticleFilename(inputDate = '') {
  const raw = String(inputDate || '').trim();
  const parsed = raw ? new Date(raw) : null;
  const dateObj = parsed && !Number.isNaN(parsed.getTime()) ? parsed : new Date();
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `clinical-article-finder-${y}-${m}-${d}.docx`;
}

function buildDatedConceptMapFilename(inputDate = '') {
  const dateObj = parseLocalDate(inputDate) || new Date();
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `clinical-concept-map-${y}-${m}-${d}.docx`;
}

function buildDatedCombinedPackageFilename(inputDate = '') {
  const raw = String(inputDate || '').trim();
  const parsed = raw ? new Date(raw) : null;
  const dateObj = parsed && !Number.isNaN(parsed.getTime()) ? parsed : new Date();
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `map-and-process-${y}-${m}-${d}.docx`;
}

function buildDatedNursingArticlePdfFilename(inputDate = '') {
  const raw = String(inputDate || '').trim();
  const parsed = raw ? new Date(raw) : null;
  const dateObj = parsed && !Number.isNaN(parsed.getTime()) ? parsed : new Date();
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `nursing-process-with-article-${y}-${m}-${d}.pdf`;
}

function buildDatedConceptMapPdfFilename(inputDate = '') {
  const dateObj = parseLocalDate(inputDate) || new Date();
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `clinical-concept-map-${y}-${m}-${d}.pdf`;
}

function extractRawValue(regex, rawText = '') {
  return clean(String(rawText || '').match(regex)?.[1] || '');
}

function extractBodyParts(documentXml = '') {
  const fullBody = String(documentXml || '').match(/<w:body[^>]*>[\s\S]*<\/w:body>/)?.[0];
  if (!fullBody) return null;

  const openTag = fullBody.match(/^<w:body[^>]*>/)?.[0] || '<w:body>';
  const closeTag = '</w:body>';
  const inner = fullBody.slice(openTag.length, fullBody.length - closeTag.length);

  const sectPrMatches = [...inner.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)];
  const lastSectPrMatch = sectPrMatches.length ? sectPrMatches[sectPrMatches.length - 1] : null;
  const sectPr = lastSectPrMatch?.[0] || '';
  const content = lastSectPrMatch
    ? `${inner.slice(0, lastSectPrMatch.index)}${inner.slice((lastSectPrMatch.index || 0) + sectPr.length)}`
    : inner;

  return {
    fullBody,
    openTag,
    content,
    sectPr,
  };
}

function mergeDocumentXmlBodies(primaryXml, secondaryXml) {
  const primary = extractBodyParts(primaryXml);
  const secondary = extractBodyParts(secondaryXml);
  if (!primary || !secondary) return primaryXml;

  const pageBreakRun = '<w:r><w:br w:type="page"/></w:r>';
  const sectionBreakParagraph = primary.sectPr
    ? `<w:p><w:pPr>${primary.sectPr}</w:pPr>${pageBreakRun}</w:p>`
    : `<w:p>${pageBreakRun}</w:p>`;
  const finalSectPr = secondary.sectPr || primary.sectPr;

  const mergedBody = `${primary.openTag}${primary.content}${sectionBreakParagraph}${secondary.content}${finalSectPr || ''}</w:body>`;
  return String(primaryXml || '').replace(primary.fullBody, mergedBody);
}

function buildConceptMapData(fields, meds, rawText = '') {
  const sex = fields.sex || extractRawValue(/\b(male|female|man|woman|non-binary|nonbinary)\b/i, rawText);
  const ht = fields.ht || extractRawValue(/\b(?:height|ht)\s*[:=-]?\s*([^,;\n]+)/i, rawText);
  const wt = fields.wt || extractRawValue(/\b(?:weight|wt)\s*[:=-]?\s*([^,;\n]+)/i, rawText);
  const bp = fields.bp || extractRawValue(/\b(?:bp|blood pressure)\s*[:=-]?\s*([^,;\n]+)/i, rawText);
  const pulse = fields.pulse || extractRawValue(/\b(?:pulse|hr)\s*[:=-]?\s*([^,;\n]+)/i, rawText);
  const resp = fields.resp || extractRawValue(/\b(?:resp(?:iratory)?(?: rate)?)\s*[:=-]?\s*([^,;\n]+)/i, rawText);
  const temp = fields.temp || extractRawValue(/\b(?:temp(?:erature)?)\s*[:=-]?\s*([^,;\n]+)/i, rawText);
  const fieldVitals = extractVitalsFromSummary(fields?.vitals || '');

  const medsText = buildMedicationOrderSummary(ensurePopulatedMeds(meds).slice(0, MED_TEMPLATE_ROW_CAP), 165);

  const pick = (...values) => {
    for (const value of values) {
      const next = clean(String(value || ''));
      if (next && !/^not assessed in case log\.?$/i.test(next)) return next;
    }
    return FALLBACK_VERIFY;
  };

  const verifyLine = (context) => `${FALLBACK_VERIFY} ${context}`;
  const pickCompact = (maxLength, ...values) => {
    const value = pick(...values);
    return value.startsWith(FALLBACK_VERIFY) ? value : compactText(value, maxLength);
  };
  const buildUniquePriorityDiagnoses = () => {
    const focused = buildFocusedPriorityDiagnoses(fields, rawText, { includeManual: true });
    return {
      nd1: focused[0] || toShortNursingDiagnosis(pickCompact(130, fields.nd1Diagnosis, fields.nursingDiagnosis, fields.diagnosis, verifyLine('Finalize priority diagnosis #1.'))),
      nd2: focused[1] || 'Risk for falls',
      nd3: focused[2] || 'Activity intolerance',
    };
  };

  const uniqueDiagnoses = buildUniquePriorityDiagnoses();

  return {
    studentName: fields.studentName || FALLBACK_NA,
    instructor: fields.facultyMeta || FALLBACK_NA,
    date: fields.date || FALLBACK_NA,
    clientName: fields.clientName || (fields.patientInitialAge ? `Patient (${fields.patientInitialAge})` : ''),
    age: fields.age || fields.patientInitialAge || FALLBACK_NA,
    sex: sex || FALLBACK_NA,
    ht: ht || FALLBACK_NA,
    wt: wt || FALLBACK_NA,
    diagnosis: pickCompact(120, fields.diagnosis),
    temp: pick(temp, fieldVitals.temp, verifyLine('Confirm charted temperature.')),
    pulse: pick(pulse, fieldVitals.pulse, verifyLine('Confirm charted pulse/HR.')),
    resp: pick(resp, fieldVitals.resp, fields.respiratory, verifyLine('Confirm charted respiratory rate.')),
    bp: pick(bp, fieldVitals.bp, verifyLine('Confirm charted blood pressure.')),
    pain: pickCompact(90, fields.pain),
    medicalHistory: pickCompact(120, fields.medicalHistory),
    surgicalHistory: pickCompact(90, fields.surgicalHistory),
    supportSystem: pickCompact(150, fields.supportSystem, verifyLine('Confirm support system and key contacts.')),
    responseHospitalization: pickCompact(150, fields.responseHospitalization, verifyLine('Confirm response to hospitalization.')),
    allergies: pickCompact(70, fields.allergies, fields.allergy),
    immunizations: pickCompact(70, fields.immunizations),
    sleepPattern: pickCompact(70, fields.sleepPattern),
    nutritionalStatus: pickCompact(70, fields.nutritionalStatus),
    allergiesImmunization: pickCompact(100, [fields.allergies, fields.immunizations].filter(Boolean).join('; '), fields.allergy),
    sleepNutritional: pickCompact(100, [fields.sleepPattern, fields.nutritionalStatus].filter(Boolean).join('; ')),
    assistanceAdls: pickCompact(110, fields.assistanceAdls),
    hygiene: pickCompact(90, fields.hygiene),
    medsPrior: pickCompact(120, fields.medsPrior, medsText, verifyLine('Confirm home medications prior to admission.')),
    genAppearance: pickCompact(120, fields.genAppearance, verifyLine('Confirm general appearance.')),
    ivLocation: pickCompact(70, fields.ivLocation),
    surgicalIncision: pickCompact(70, fields.surgicalIncision),
    orientation: pickCompact(90, fields.orientation, fields.neuro, verifyLine('Confirm orientation status.')),
    speech: pickCompact(90, fields.speech, verifyLine('Confirm speech findings.')),
    weakness: pickCompact(80, fields.weakness, verifyLine('Confirm motor weakness findings.')),
    skinTurgor: pickCompact(90, fields.skinTurgor, verifyLine('Confirm hydration/skin turgor.')),
    breathSounds: pickCompact(90, fields.breathSounds, fields.respiratory, verifyLine('Assess breath sounds.')),
    peripheralPulses: pickCompact(90, fields.peripheralPulses, verifyLine('Confirm peripheral pulse quality.')),
    edema: pickCompact(90, fields.edema),
    bowelSounds: pickCompact(90, fields.bowelSounds, fields.gi, verifyLine('Confirm bowel sound pattern.')),
    physicalOther: pickCompact(120, fields.physicalOther, verifyLine('Confirm additional physical findings.')),
    spiritualAssessment: pickCompact(100, fields.spiritualAssessment, fields.cultural),
    currentMedsPlan: pickCompact(165, fields.currentMedOrder, medsText, verifyLine('Confirm active meds and therapeutic plan.')),
    nd1Assessment: pickCompact(150, fields.nd1Assessment, fields.psychosocial, fields.neuro, verifyLine('Confirm assessment cues for diagnosis #1.')),
    nd1Diagnosis: uniqueDiagnoses.nd1,
    nd1Rationale: pickCompact(150, fields.nd1Rationale, fields.rationale, verifyLine('Add evidence-based rationale for #1.')),
    nd1Intervention: pickCompact(150, fields.nd1Intervention, fields.intervention, verifyLine('Specify intervention/action for #1.')),
    nd1Evaluation: pickCompact(150, fields.nd1Evaluation, fields.evaluation, verifyLine('Document response/evaluation for #1.')),
    nd2Assessment: pickCompact(150, fields.nd2Assessment, fields.safety, fields.activity, fields.cardio, verifyLine('Confirm assessment cues for diagnosis #2.')),
    nd2Diagnosis: uniqueDiagnoses.nd2,
    nd2Rationale: pickCompact(150, fields.nd2Rationale, fields.rationale, verifyLine('Add evidence-based rationale for #2.')),
    nd2Intervention: pickCompact(150, fields.nd2Intervention, fields.plan, verifyLine('Specify intervention/action for #2.')),
    nd2Evaluation: pickCompact(150, fields.nd2Evaluation, fields.reassessment, verifyLine('Document response/evaluation for #2.')),
    nd3Assessment: pickCompact(150, fields.nd3Assessment, fields.pain, fields.vitals, fields.respiratory, verifyLine('Confirm assessment cues for diagnosis #3.')),
    nd3Diagnosis: uniqueDiagnoses.nd3,
    nd3Rationale: pickCompact(150, fields.nd3Rationale, fields.rationale, verifyLine('Add evidence-based rationale for #3.')),
    nd3Intervention: pickCompact(150, fields.nd3Intervention, fields.intervention, verifyLine('Specify intervention/action for #3.')),
    nd3Evaluation: pickCompact(150, fields.nd3Evaluation, fields.evaluation, verifyLine('Document response/evaluation for #3.')),
    labTestName: pickCompact(70, fields.labTestName, verifyLine('Add key lab/diagnostic test name.')),
    labClientResults: pickCompact(70, fields.labClientResults, verifyLine('Enter current client result.')),
    labNormalValue: pickCompact(70, fields.labNormalValue, verifyLine('Enter normal reference range.')),
    labInterpretation: pickCompact(90, fields.labInterpretation, verifyLine('Interpret significance to current condition.')),
    currentMedDate: pickCompact(40, fields.currentMedDate, fields.date, verifyLine('Enter current order date.')),
    currentMedOrder: pickCompact(165, fields.currentMedOrder, medsText, verifyLine('Enter active order/dose/frequency.')),
    currentMedIndication: pickCompact(80, fields.currentMedIndication, verifyLine('Enter medication indication for this client.')),
  };
}

function applyConceptMapFillToXml(xml, conceptData) {
  let output = String(xml || '');
  const once = (regex, replacement) => {
    let replaced = false;
    output = output.replace(regex, (match) => {
      if (replaced) return match;
      replaced = true;
      return replacement;
    });
  };
  const replaceNth = (regex, nth, replacement) => {
    let seen = 0;
    output = output.replace(regex, (match) => {
      seen += 1;
      return seen === nth ? replacement : match;
    });
  };
  const escapeRegex = (value = '') => String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const fillPrioritySection = (sectionHeading, nextSectionHeading, values) => {
    const startIdx = output.search(new RegExp(sectionHeading, 'i'));
    if (startIdx === -1) return;

    const tail = output.slice(startIdx);
    const endRel = nextSectionHeading ? tail.search(new RegExp(nextSectionHeading, 'i')) : -1;
    const endIdx = endRel === -1 ? output.length : startIdx + endRel;

    const before = output.slice(0, startIdx);
    let section = output.slice(startIdx, endIdx);
    const after = output.slice(endIdx);

    const add = (label, value) => {
      const safeValue = escapeXmlText(value || FALLBACK_NOT_ASSESSED);
      const rx = new RegExp(`(${escapeRegex(label)})(?![^<]{0,220}${escapeRegex(safeValue)})`, 'i');
      section = section.replace(rx, `$1: ${safeValue}`);
    };

    add('Client Assessment', values.assessment);
    add('Nursing Diagnosis', values.diagnosis);
    add('Scientific Rationale', values.rationale);
    add('Nursing Intervention/Action', values.intervention);
    add('Evaluation/Client Response', values.evaluation);

    output = `${before}${section}${after}`;
  };

  const map = [
    [/Student[’']s Name:\s*[^<]*/i, `Student's Name: ${escapeXmlText(conceptData.studentName)}`],
    [/Instructor:\s*[^<]*/i, `Instructor: ${escapeXmlText(conceptData.instructor)}`],
    [/Date:\s*[^<]*/i, `Date: ${escapeXmlText(conceptData.date)}`],
    [/Client[’']s Name:\s*[^<]*/i, `Client's Name: ${escapeXmlText(conceptData.clientName)}`],
    [/Age:\s*Sex:\s*[^<]*/i, `Age: ${escapeXmlText(conceptData.age)} Sex: ${escapeXmlText(conceptData.sex)}`],
    [/Diagnosis:\s*[^<]*/i, `Diagnosis: ${escapeXmlText(conceptData.diagnosis)}`],
    [/Medical:\s*[^<]*/i, `Medical: ${escapeXmlText(conceptData.medicalHistory)}`],
    [/Surgical:\s*[^<]*/i, `Surgical: ${escapeXmlText(conceptData.surgicalHistory)}`],
    [/Support system\/?significant others:\s*[^<]*/i, `Support system/significant others: ${escapeXmlText(conceptData.supportSystem)}`],
    [/Response to current hospitalization:\s*[^<]*/i, `Response to current hospitalization: ${escapeXmlText(conceptData.responseHospitalization)}`],
    [/Allergies:\s*Immunization:\s*[^<]*/i, `Allergies: Immunization: ${escapeXmlText(conceptData.allergiesImmunization)}`],
    [/Assistance with ADLs:\s*[^<]*/i, `Assistance with ADLs: ${escapeXmlText(conceptData.assistanceAdls)}`],
    [/Hygiene:\s*[^<]*/i, `Hygiene: ${escapeXmlText(conceptData.hygiene)}`],
    [/Medications prior to hospitalization:\s*[^<]*/i, `Medications prior to hospitalization: ${escapeXmlText(conceptData.medsPrior)}`],
    [/Gen\. Appearance:\s*[^<]*/i, `Gen. Appearance: ${escapeXmlText(conceptData.genAppearance)}`],
    [/Orientation:\s*[^<]*/i, `Orientation: ${escapeXmlText(conceptData.orientation)}`],
    [/Speech:\s*[^<]*/i, `Speech: ${escapeXmlText(conceptData.speech)}`],
    [/Skin Turgor:\s*[^<]*/i, `Skin Turgor: ${escapeXmlText(conceptData.skinTurgor)}`],
    [/Breath Sounds:\s*[^<]*/i, `Breath Sounds: ${escapeXmlText(conceptData.breathSounds)}`],
    [/Peripheral Pulses:\s*[^<]*/i, `Peripheral Pulses: ${escapeXmlText(conceptData.peripheralPulses)}`],
    [/Edema:\s*[^<]*/i, `Edema: ${escapeXmlText(conceptData.edema)}`],
    [/Bowel Sounds:\s*[^<]*/i, `Bowel Sounds: ${escapeXmlText(conceptData.bowelSounds)}`],
    [/Spiritual Assessment\s*[^<]*/i, `Spiritual Assessment: ${escapeXmlText(conceptData.spiritualAssessment)}`],
  ];

  for (const [regex, replacement] of map) {
    once(regex, replacement);
  }

  // Some templates split this label across runs: "Support system" + "/significant others:".
  // Fallback to target the second run directly so the value is reliably injected.
  once(/\/significant others:\s*[^<]*/i, `/significant others: ${escapeXmlText(conceptData.supportSystem)}`);

  fillPrioritySection('Priority Nursing Diagnosis #1', 'Priority Nursing Diagnosis #2', {
    assessment: conceptData.nd1Assessment,
    diagnosis: conceptData.nd1Diagnosis,
    rationale: conceptData.nd1Rationale,
    intervention: conceptData.nd1Intervention,
    evaluation: conceptData.nd1Evaluation,
  });

  fillPrioritySection('Priority Nursing Diagnosis #2', 'Priority Nursing Diagnosis #3', {
    assessment: conceptData.nd2Assessment,
    diagnosis: conceptData.nd2Diagnosis,
    rationale: conceptData.nd2Rationale,
    intervention: conceptData.nd2Intervention,
    evaluation: conceptData.nd2Evaluation,
  });

  fillPrioritySection('Priority Nursing Diagnosis #3', '', {
    assessment: conceptData.nd3Assessment,
    diagnosis: conceptData.nd3Diagnosis,
    rationale: conceptData.nd3Rationale,
    intervention: conceptData.nd3Intervention,
    evaluation: conceptData.nd3Evaluation,
  });

  const fillDiagOccurrence = (idx, values) => {
    replaceNth(/Client Assessment(?::\s*[^<]*)?/gi, idx, `Client Assessment: ${escapeXmlText(values.assessment || FALLBACK_NOT_ASSESSED)}`);
    replaceNth(/Nursing Diagnosis(?::\s*[^<]*)?/gi, idx, `Nursing Diagnosis: ${escapeXmlText(values.diagnosis || FALLBACK_NOT_ASSESSED)}`);
    replaceNth(/Scientific\s*Rationale(?::\s*[^<]*)?/gi, idx, `Scientific Rationale: ${escapeXmlText(values.rationale || FALLBACK_NOT_ASSESSED)}`);
    replaceNth(/Nursing Intervention\/Action(?::\s*[^<]*)?/gi, idx, `Nursing Intervention/Action: ${escapeXmlText(values.intervention || FALLBACK_NOT_ASSESSED)}`);
    replaceNth(/Evaluation\/Client Response(?::\s*[^<]*)?/gi, idx, `Evaluation/Client Response: ${escapeXmlText(values.evaluation || FALLBACK_NOT_ASSESSED)}`);
  };

  fillDiagOccurrence(1, {
    assessment: conceptData.nd1Assessment,
    diagnosis: conceptData.nd1Diagnosis,
    rationale: conceptData.nd1Rationale,
    intervention: conceptData.nd1Intervention,
    evaluation: conceptData.nd1Evaluation,
  });
  fillDiagOccurrence(2, {
    assessment: conceptData.nd2Assessment,
    diagnosis: conceptData.nd2Diagnosis,
    rationale: conceptData.nd2Rationale,
    intervention: conceptData.nd2Intervention,
    evaluation: conceptData.nd2Evaluation,
  });
  fillDiagOccurrence(3, {
    assessment: conceptData.nd3Assessment,
    diagnosis: conceptData.nd3Diagnosis,
    rationale: conceptData.nd3Rationale,
    intervention: conceptData.nd3Intervention,
    evaluation: conceptData.nd3Evaluation,
  });

  replaceNth(/Test Name(?::\s*[^<]*)?/gi, 1, `Test Name: ${escapeXmlText(conceptData.labTestName)}`);
  replaceNth(/Client[’']s Results(?::\s*[^<]*)?/gi, 1, `Client’s Results: ${escapeXmlText(conceptData.labClientResults)}`);
  replaceNth(/Normal Value(?::\s*[^<]*)?/gi, 1, `Normal Value: ${escapeXmlText(conceptData.labNormalValue)}`);
  replaceNth(/Interpretation(?::\s*[^<]*)?/gi, 1, `Interpretation: ${escapeXmlText(conceptData.labInterpretation)}`);

  replaceNth(/Date(?::\s*[^<]*)?/gi, 2, `Date: ${escapeXmlText(conceptData.currentMedDate)}`);
  replaceNth(/Order\/dosage\/Frequency(?::\s*[^<]*)?/gi, 1, `Order/dosage/Frequency: ${escapeXmlText(conceptData.currentMedOrder)}`);
  replaceNth(/Indication(?::\s*[^<]*)?/gi, 1, `Indication: ${escapeXmlText(conceptData.currentMedIndication)}`);

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(output, 'application/xml');
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(ns, 'p'));
  const paragraphText = (p) => Array.from(p.getElementsByTagNameNS(ns, 't')).map((n) => n.textContent || '').join('').replace(/\u00a0/g, ' ').trim();

  const applyVitalLine = (matcher, nextText) => {
    const target = paragraphs.find((p) => matcher.test(paragraphText(p)));
    if (!target) return;
    setParagraphText(target, nextText);
  };

  applyVitalLine(/^temp\s*-?\s*$/i, `Temp- ${conceptData.temp || FALLBACK_NOT_ASSESSED}`);
  applyVitalLine(/^pulse\s*-?\s*$/i, `Pulse- ${conceptData.pulse || FALLBACK_NOT_ASSESSED}`);
  applyVitalLine(/^resp\s*-?\s*$/i, `Resp- ${conceptData.resp || FALLBACK_NOT_ASSESSED}`);
  applyVitalLine(/^bp\s*-?\s*$/i, `BP- ${conceptData.bp || FALLBACK_NOT_ASSESSED}`);
  applyVitalLine(/^pain\s*-?\s*$/i, `Pain- ${conceptData.pain || FALLBACK_NOT_ASSESSED}`);

  return new XMLSerializer().serializeToString(xmlDoc);
}

function stripUnderlineRunsFromConceptMapXml(xml) {
  // Remove explicit underline run properties to improve readability of dense placeholder text.
  return String(xml || '')
    .replace(/<w:u\b[^>]*\/>/gi, '')
    .replace(/<w:u\b[^>]*><\/w:u>/gi, '');
}

async function buildFilledConceptMapZip(templateUrl, fields, meds, rawText = '') {
  const arrayBuffer = await loadTemplateArrayBuffer(templateUrl);
  let zip;
  try {
    zip = new PizZip(arrayBuffer);
  } catch {
    throw new Error('Could not parse concept map DOCX template.');
  }

  const documentXmlFile = zip.file('word/document.xml');
  if (!documentXmlFile) throw new Error('Concept map DOCX is missing word/document.xml.');

  const conceptData = buildConceptMapData(fields, meds, rawText);
  const updatedXml = applyConceptMapFillToXml(documentXmlFile.asText(), conceptData);
  const cleanedXml = stripUnderlineRunsFromConceptMapXml(updatedXml);
  zip.file('word/document.xml', cleanedXml);

  return zip;
}

async function exportConceptMapTemplate(templateUrl, fields, meds, rawText = '') {
  const zip = await buildFilledConceptMapZip(templateUrl, fields, meds, rawText);

  const out = zip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  const filename = buildDatedConceptMapFilename(fields?.date);
  await saveBlobAsFile(out, filename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}

async function renderPdfTemplatePageToImage(templateUrl) {
  const pdfjsLib = await import('pdfjs-dist');
  configurePdfJsWorker(pdfjsLib);

  const sourcePdf = await pdfjsLib.getDocument(templateUrl).promise;
  const page = await sourcePdf.getPage(1);
  const viewport = page.getViewport({ scale: 2.5 });
  const canvas = document.createElement('canvas');
  const context = canvas.getContext('2d');
  canvas.width = viewport.width;
  canvas.height = viewport.height;
  await page.render({ canvasContext: context, viewport }).promise;
  return {
    dataUrl: canvas.toDataURL('image/png'),
    width: page.view[2] - page.view[0],
    height: page.view[3] - page.view[1],
  };
}

async function exportConceptMapPdf(fields, meds, rawText = '', options = {}) {
  const builtData = buildConceptMapData(fields, meds, rawText);
  const data = options.blankPriorityNursingSections
    ? {
        ...builtData,
        nd1Assessment: '',
        nd1Diagnosis: '',
        nd1Rationale: '',
        nd1Intervention: '',
        nd1Evaluation: '',
        nd2Assessment: '',
        nd2Diagnosis: '',
        nd2Rationale: '',
        nd2Intervention: '',
        nd2Evaluation: '',
        nd3Assessment: '',
        nd3Diagnosis: '',
        nd3Rationale: '',
        nd3Intervention: '',
        nd3Evaluation: '',
      }
    : builtData;
  const template = await renderPdfTemplatePageToImage(DEFAULT_CONCEPT_MAP_PDF_TEMPLATE_URL);
  const pdf = new jsPDF({ orientation: 'landscape', unit: 'pt', format: [template.width, template.height] });
  pdf.addImage(template.dataUrl, 'PNG', 0, 0, template.width, template.height);

  const color = (r = 17, g = 24, b = 39) => pdf.setTextColor(r, g, b);
  const cover = (x, y, width, height) => {
    pdf.setFillColor(255, 255, 255);
    pdf.rect(x, y, width, height, 'F');
  };
  const font = (size = 7, style = 'normal') => {
    pdf.setFont('helvetica', style);
    pdf.setFontSize(size);
    pdf.setLineHeightFactor(1.05);
    color();
  };
  const text = (value, x, y, options = {}) => {
    const size = options.size || 7;
    const width = options.width || 80;
    const maxLines = options.maxLines || 2;
    font(size, options.style || 'normal');
    const printableValue = /^need to verify/i.test(clean(String(value || ''))) ? 'Verify' : stripOutputReviewMarkers(value);
    const wrapped = pdf.splitTextToSize(compactText(printableValue, options.maxChars || 120), width).slice(0, maxLines);
    pdf.text(wrapped, x, y);
  };
  const boxText = (value, box, options = {}) => {
    const printableValue = /^need to verify/i.test(clean(String(value || ''))) ? 'Verify' : stripOutputReviewMarkers(value);
    const clipped = compactText(printableValue, options.maxChars || box.maxChars || 120);
    if (!clipped) return;

    let size = options.size || box.size || 5.1;
    const minSize = options.minSize || box.minSize || 3.4;
    const lineHeightFactor = options.lineHeightFactor || 1.05;
    let wrapped = [];
    let maxLines = 1;

    while (size >= minSize) {
      font(size, options.style || box.style || 'normal');
      maxLines = Math.max(1, Math.floor(box.height / (size * lineHeightFactor)));
      wrapped = pdf.splitTextToSize(clipped, box.width);
      if (wrapped.length <= maxLines) break;
      size -= 0.25;
    }

    font(size, options.style || box.style || 'normal');
    const lines = wrapped.slice(0, maxLines);
    pdf.text(lines, box.x, box.y + size);
  };
  const tiny = (value, x, y, width, maxLines = 2, maxChars = 110) => text(value, x, y, { size: 5.1, width, maxLines, maxChars });
  const micro = (value, x, y, width, maxLines = 1, maxChars = 70) => text(value, x, y, { size: 4.3, width, maxLines, maxChars });

  text(data.studentName, 180, 64, { size: 8, width: 115, maxLines: 1, maxChars: 45 });
  text(data.date, 338, 62, { size: 8, width: 80, maxLines: 1, maxChars: 24 });
  text(data.instructor, 696, 60, { size: 6.2, width: 88, maxLines: 1, maxChars: 24 });

  micro(data.allergies, 52, 93, 82, 1, 26);
  micro(data.immunizations, 146, 93, 90, 1, 28);
  micro(data.sleepPattern, 64, 114, 96, 1, 30);
  micro(data.nutritionalStatus, 160, 114, 78, 1, 28);
  micro(data.assistanceAdls, 64, 137, 128, 1, 36);
  micro(data.hygiene, 64, 158, 118, 1, 36);
  micro(data.medsPrior, 54, 196, 172, 2, 75);

  tiny(data.medicalHistory, 330, 91, 130, 3, 115);
  tiny(data.surgicalHistory, 330, 121, 130, 2, 95);
  tiny(data.supportSystem, 330, 150, 130, 3, 90);
  tiny(data.responseHospitalization, 330, 174, 130, 3, 90);

  micro(data.genAppearance, 560, 101, 142, 1, 42);
  micro(data.ivLocation, 552, 122, 70, 1, 28);
  micro(data.surgicalIncision, 660, 122, 70, 1, 24);
  micro(data.orientation, 558, 137, 72, 1, 30);
  micro(data.speech, 638, 137, 52, 1, 24);
  micro(data.weakness, 688, 137, 64, 1, 22);
  micro(data.skinTurgor, 558, 158, 72, 1, 28);
  micro(data.breathSounds, 648, 158, 72, 1, 28);
  micro(data.peripheralPulses, 558, 180, 82, 1, 30);
  micro(data.edema, 676, 180, 66, 1, 26);
  micro(data.bowelSounds, 558, 201, 82, 1, 30);
  micro(data.physicalOther, 672, 196, 62, 1, 26);
  micro(`S. Assessment: ${data.spiritualAssessment}`, 610, 222, 128, 1, 52);

  tiny(data.labTestName, 34, 283, 50, 2, 50);
  tiny(data.labClientResults, 88, 283, 58, 2, 55);
  tiny(data.labNormalValue, 151, 283, 55, 2, 45);
  tiny(data.labInterpretation, 215, 283, 52, 3, 70);

  tiny(data.currentMedDate, 514, 278, 42, 1, 30);
  boxText(data.currentMedOrder, { x: 566, y: 274, width: 126, height: 82, size: 4.3, minSize: 3.35, maxChars: 165 });
  boxText(data.currentMedIndication, { x: 700, y: 274, width: 55, height: 90, size: 5.1, minSize: 3.5, maxChars: 80 });

  text(data.clientName, 386, 230, { size: 6.2, width: 80, maxLines: 1, maxChars: 24 });
  text(data.age, 342, 252, { size: 7, width: 22, maxLines: 1, maxChars: 10 });
  text(data.sex, 384, 252, { size: 7, width: 18, maxLines: 1, maxChars: 8 });
  text(data.ht, 420, 252, { size: 7, width: 32, maxLines: 1, maxChars: 12 });
  text(data.wt, 458, 252, { size: 7, width: 38, maxLines: 1, maxChars: 12 });
  text(data.diagnosis, 370, 282, { size: 6.2, width: 115, maxLines: 2, maxChars: 80 });
  text(data.temp, 372, 308, { size: 7, width: 80, maxLines: 1, maxChars: 35 });
  text(data.pulse, 372, 321, { size: 7, width: 80, maxLines: 1, maxChars: 35 });
  text(data.resp, 372, 334, { size: 7, width: 80, maxLines: 1, maxChars: 35 });
  text(data.bp, 372, 347, { size: 7, width: 80, maxLines: 1, maxChars: 35 });
  text(data.pain, 372, 360, { size: 7, width: 105, maxLines: 1, maxChars: 45 });

  const priority = (x, values) => {
    boxText(values.assessment, { x: x + 8, y: 436, width: 68, height: 45, size: 5.1, minSize: 3.4, maxChars: 115 });
    boxText(values.diagnosis, { x: x + 86, y: 436, width: 68, height: 45, size: 5.1, minSize: 3.4, maxChars: 42 });
    boxText(values.rationale, { x: x + 162, y: 436, width: 68, height: 45, size: 5.1, minSize: 3.4, maxChars: 115 });
    boxText(values.intervention, { x: x + 4, y: 500, width: 100, height: 86, size: 5.1, minSize: 3.4, maxChars: 130 });
    boxText(values.evaluation, { x: x + 146, y: 490, width: 88, height: 96, size: 5.1, minSize: 3.4, maxChars: 115 });
  };
  priority(15, {
    assessment: data.nd1Assessment,
    diagnosis: data.nd1Diagnosis,
    rationale: data.nd1Rationale,
    intervention: data.nd1Intervention,
    evaluation: data.nd1Evaluation,
  });
  priority(280, {
    assessment: data.nd2Assessment,
    diagnosis: data.nd2Diagnosis,
    rationale: data.nd2Rationale,
    intervention: data.nd2Intervention,
    evaluation: data.nd2Evaluation,
  });
  priority(530, {
    assessment: data.nd3Assessment,
    diagnosis: data.nd3Diagnosis,
    rationale: data.nd3Rationale,
    intervention: data.nd3Intervention,
    evaluation: data.nd3Evaluation,
  });

  pdf.save(buildDatedConceptMapPdfFilename(fields?.date));
}

function toTemplateData(fields, meds) {
  const medRows = Array.from({ length: MED_TEMPLATE_ROW_CAP }, (_, index) => meds[index] || EMPTY_MED);

  return {
    ...fields,
    // Alias keys for template friendliness.
    student_name: fields.studentName,
    week_number: fields.week,
    patient_initial_age: fields.patientInitialAge,
    patient_diagnosis: fields.diagnosis,
    vital_signs: fields.vitals,
    psychosocial_developmental_status: fields.psychosocial,
    cultural_spiritual: fields.cultural,
    educational_needs: fields.educationNeeds,
    safety_precautions: fields.safety,
    diagnostic_test_scheduled: fields.diagnosticTests,
    recent_lab_results: fields.labs,
    nursing_goal: fields.goal,
    plan_of_care: fields.plan,
    intervention_given: fields.intervention,
    outcome_evaluation: fields.evaluation,
    theorist_applied: fields.theorist,
    course_objectives_met: fields.courseObjectives,
    medications: meds.map((m, i) => ({ ...m, index: i + 1 })),
    medicationCount: meds.length,
    med1: medRows[0],
    med2: medRows[1],
    med3: medRows[2],
    med4: medRows[3],
    med5: medRows[4],
    med6: medRows[5],
    generatedAt: new Date().toISOString(),
  };
}

async function loadTemplateArrayBuffer(templateUrl) {
  const response = await fetch(templateUrl);
  if (!response.ok) {
    throw new Error(`Default template not found at ${templateUrl}. Add the converted .docx template to public/templates.`);
  }
  const arrayBuffer = await response.arrayBuffer();

  const bytes = new Uint8Array(arrayBuffer);
  const isZipHeader = bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b;
  if (!isZipHeader) {
    throw new Error('Template file is not a valid .docx zip file. Convert your .doc to .docx and place it at public/templates/nursing-process-template.docx.');
  }

  return arrayBuffer;
}

function setParagraphText(paragraphNode, text) {
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const safeText = sanitizeTextForDocument(text);
  const textNodes = Array.from(paragraphNode.getElementsByTagNameNS(ns, 't'));
  if (!textNodes.length) {
    let run = paragraphNode.getElementsByTagNameNS(ns, 'r')[0];
    if (!run) {
      run = paragraphNode.ownerDocument.createElementNS(ns, 'w:r');
      paragraphNode.appendChild(run);
    }
    const t = paragraphNode.ownerDocument.createElementNS(ns, 'w:t');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    t.textContent = text;
    run.appendChild(t);
    return;
  }
  textNodes[0].textContent = safeText;
  for (let i = 1; i < textNodes.length; i += 1) {
    textNodes[i].textContent = '';
  }
}

function setParagraphValueAfterLabel(paragraphNode, value, labelRegex) {
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const runs = Array.from(paragraphNode.getElementsByTagNameNS(ns, 'r'));
  const runTextNodes = runs.map((run) => Array.from(run.getElementsByTagNameNS(ns, 't')));
  const runTexts = runTextNodes.map((nodes) => nodes.map((n) => n.textContent || '').join(''));
  const full = runTexts.join('');

  if (!labelRegex.test(full)) {
    return false;
  }

  let labelEndRunIndex = runTexts.findIndex((text) => text.includes(':'));
  if (labelEndRunIndex === -1) {
    const colonIdx = full.indexOf(':');
    if (colonIdx === -1) return false;
    const before = full.slice(0, colonIdx + 1);
    setParagraphText(paragraphNode, `${before} ${value}`);
    return true;
  }

  let valueRunIndex = -1;
  for (let i = labelEndRunIndex + 1; i < runTexts.length; i += 1) {
    if ((runTexts[i] || '').trim()) {
      valueRunIndex = i;
      break;
    }
  }
  if (valueRunIndex === -1) {
    valueRunIndex = Math.min(labelEndRunIndex + 1, runTexts.length - 1);
  }

  for (let i = 0; i < runTexts.length; i += 1) {
    const nodes = runTextNodes[i];
    if (!nodes.length) continue;
    if (i < valueRunIndex) continue;
    if (i === valueRunIndex) {
      nodes[0].textContent = ` ${value}`;
      for (let j = 1; j < nodes.length; j += 1) nodes[j].textContent = '';
    } else {
      for (const node of nodes) node.textContent = '';
    }
  }

  return true;
}

function setTableCellText(cellNode, text) {
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const paragraphs = Array.from(cellNode.getElementsByTagNameNS(ns, 'p'));
  if (!paragraphs.length) {
    const p = cellNode.ownerDocument.createElementNS(ns, 'w:p');
    cellNode.appendChild(p);
    setParagraphText(p, text);
    return;
  }
  setParagraphText(paragraphs[0], text);
  for (let i = 1; i < paragraphs.length; i += 1) {
    setParagraphText(paragraphs[i], '');
  }
}

function stripResearchParagraphsFromXml(xml) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xml, 'application/xml');
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(ns, 'p'));
  const paragraphText = (p) => Array.from(p.getElementsByTagNameNS(ns, 't')).map((n) => n.textContent || '').join('');

  const researchPatterns = [
    /Research Article Title:/i,
    /The Journey of Adolescent Paranoia/i,
    /^Article Link:/i,
    /pmc\.ncbi\.nlm\.nih\.gov/i,
    /^Why I chose this article:/i,
    /^I chose this article because/i,
    /persecutory thinking/i,
    /patient.?s story of feeling persecuted/i,
    /^APA Citation:/i,
    /Bird,\s*J\.\s*C\./i,
    /Psychology and Psychotherapy/i,
    /doi\.org\//i,
  ];

  let changed = false;
  for (const p of paragraphs) {
    const text = paragraphText(p);
    if (researchPatterns.some((rx) => rx.test(text))) {
      setParagraphText(p, '');
      changed = true;
    }
  }

  return changed ? new XMLSerializer().serializeToString(xmlDoc) : xml;
}

function applyResearchCleanupToZip(zip) {
  const xmlFiles = zip.file(/^word\/(document|header\d+|footer\d+)\.xml$/) || [];
  for (const file of xmlFiles) {
    const original = file.asText();
    const cleaned = stripResearchParagraphsFromXml(original);
    if (cleaned !== original) {
      zip.file(file.name, cleaned);
    }
  }
}

function escapeXmlText(text = '') {
  const sanitized = String(text || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '');
  return sanitized
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildWordTextParagraph(text = '', { bold = false } = {}) {
  const safe = escapeXmlText(text);
  if (!safe) {
    return '<w:p/>';
  }
  return `<w:p><w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${safe}</w:t></w:r></w:p>`;
}

function buildArticleAppendixXml(articleResult, options = {}) {
  const { includePageBreak = true } = options;
  if (!articleResult || !clean(String(articleResult.title || articleResult.link || articleResult.why || articleResult.citation || ''))) {
    return '';
  }

  const parts = [
    buildWordTextParagraph('Nursing Process Worksheet for Students', { bold: true }),
    buildWordTextParagraph('Scholarly Article Attachment', { bold: true }),
    '<w:p/>',
    buildWordTextParagraph('Research Article Title:', { bold: true }),
    buildWordTextParagraph(articleResult.title || ''),
    '<w:p/>',
    buildWordTextParagraph('Article Link:', { bold: true }),
    buildWordTextParagraph(articleResult.link || ''),
    '<w:p/>',
    buildWordTextParagraph('Why I chose this article:', { bold: true }),
    ...String(articleResult.why || '').split(/\n+/).map((line) => buildWordTextParagraph(line || '')),
    '<w:p/>',
    buildWordTextParagraph('APA Citation:', { bold: true }),
    ...String(articleResult.citation || '').split(/\n+/).map((line) => buildWordTextParagraph(line || '')),
  ];

  if (includePageBreak) {
    parts.unshift('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
  }

  return parts.join('');
}

function appendArticleAppendixToDocumentXml(documentXml, articleResult) {
  const appendixXml = buildArticleAppendixXml(articleResult, { includePageBreak: true });
  if (!appendixXml) return documentXml;
  if (!documentXml.includes('</w:body>')) return documentXml;
  return documentXml.replace('</w:body>', `${appendixXml}</w:body>`);
}

function buildArticleOnlyDocumentXml(documentXml, articleResult) {
  const appendixXml = buildArticleAppendixXml(articleResult, { includePageBreak: false });
  if (!appendixXml) return documentXml;

  const bodyMatch = documentXml.match(/<w:body[^>]*>[\s\S]*?<\/w:body>/);
  if (!bodyMatch) return documentXml;

  const fullBody = bodyMatch[0];
  const openTagMatch = fullBody.match(/^<w:body[^>]*>/);
  if (!openTagMatch) return documentXml;
  const openTag = openTagMatch[0];
  const closeTag = '</w:body>';
  const oldInner = fullBody.slice(openTag.length, fullBody.length - closeTag.length);
  const sectPrMatch = oldInner.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
  const sectPr = sectPrMatch ? sectPrMatch[0] : '';
  const newBody = `${openTag}${appendixXml}${sectPr}${closeTag}`;

  return documentXml.replace(fullBody, newBody);
}

function applyArticleAppendixToZip(zip, articleResult) {
  if (!articleResult) return;
  const documentXmlFile = zip.file('word/document.xml');
  if (!documentXmlFile) return;
  const original = documentXmlFile.asText();
  const updated = appendArticleAppendixToDocumentXml(original, articleResult);
  if (updated !== original) {
    zip.file('word/document.xml', updated);
  }
}

async function exportArticleOnlyDocxTemplate(templateUrl, articleResult, fields = {}) {
  if (!articleResult) {
    throw new Error('No article output available. Run Step 5 article search first.');
  }

  const arrayBuffer = await loadTemplateArrayBuffer(templateUrl);
  let zip;
  try {
    zip = new PizZip(arrayBuffer);
  } catch {
    throw new Error('Could not parse template DOCX for article-only export.');
  }

  const documentXmlFile = zip.file('word/document.xml');
  if (!documentXmlFile) throw new Error('Template DOCX is missing word/document.xml.');

  const articleOnlyXml = buildArticleOnlyDocumentXml(documentXmlFile.asText(), articleResult);
  zip.file('word/document.xml', articleOnlyXml);

  const out = zip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  const filename = buildDatedArticleFilename(fields?.date);
  await saveBlobAsFile(out, filename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}

function applyLabelFillToXml(xml, fields, meds = []) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xml, 'application/xml');
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const paragraphs = Array.from(xmlDoc.getElementsByTagNameNS(ns, 'p'));
  const paragraphText = (p) => Array.from(p.getElementsByTagNameNS(ns, 't')).map((n) => n.textContent || '').join('');
  const findParagraph = (regex) => paragraphs.find((p) => regex.test(paragraphText(p)));

  const mappings = [
    [/Students Name:/i, () => `Students Name: ${fields.studentName || FALLBACK_NA}    Date: ${fields.date || FALLBACK_NA}    Week # ${fields.week || FALLBACK_NA}`],
    [/^Allergy:/i, () => fields.allergy || FALLBACK_NOT_ASSESSED],
    [/^Neuro:/i, () => fields.neuro || FALLBACK_NOT_ASSESSED],
    [/^Cardiovascular/i, () => fields.cardio || FALLBACK_NOT_ASSESSED],
    [/^Respiratory:/i, () => fields.respiratory || FALLBACK_NOT_ASSESSED],
    [/^Vital Signs:/i, () => fields.vitals || FALLBACK_NOT_ASSESSED],
    [/^Gastrointestinal \(GI\):/i, () => fields.gi || FALLBACK_NOT_ASSESSED],
    [/^Genitourinary \(GU\):/i, () => fields.gu || FALLBACK_NOT_ASSESSED],
    [/^Skin/i, () => fields.skin || FALLBACK_NOT_ASSESSED],
    [/^Fluid\s*&\s*Electrolyte Status:/i, () => fields.fluidElectrolytes || FALLBACK_NOT_ASSESSED],
    [/^Pain Management:/i, () => fields.pain || FALLBACK_NOT_ASSESSED],
    [/^Psychosocial\/Developmental Status:/i, () => fields.psychosocial || FALLBACK_NOT_ASSESSED],
    [/^Prescribed Activity/i, () => fields.activity || FALLBACK_NOT_ASSESSED],
    [/^Cultural\/Spiritual:/i, () => fields.cultural || FALLBACK_NOT_ASSESSED],
    [/^Pt\/Family Educational Needs:/i, () => fields.educationNeeds || FALLBACK_NOT_ASSESSED],
    [/^Specific Safety Need\/Precautions:/i, () => fields.safety || FALLBACK_NOT_ASSESSED],
    [/^Diagnostic Test or/i, () => fields.diagnosticTests || FALLBACK_NOT_ASSESSED],
    [/^Recent Lab Results/i, () => fields.labs || FALLBACK_NOT_ASSESSED],
    [/^Discharge Plan:/i, () => fields.dischargePlan || FALLBACK_NOT_ASSESSED],
    [/^Nursing Diagnosis:/i, () => fields.nursingDiagnosis || FALLBACK_NOT_ASSESSED],
    [/^Nursing Goal:/i, () => fields.goal || FALLBACK_NOT_ASSESSED],
    [/^Plan of Care:/i, () => fields.plan || FALLBACK_NOT_ASSESSED],
    [/^Intervention\s*:/i, () => fields.intervention || FALLBACK_NOT_ASSESSED],
    [/^Rational for Intervention:/i, () => fields.rationale || FALLBACK_NOT_ASSESSED],
    [/^Evaluation\/Outcome/i, () => fields.evaluation || FALLBACK_NOT_ASSESSED],
    [/^Reassessment:/i, () => fields.reassessment || FALLBACK_NOT_ASSESSED],
  ];

  for (const [regex, builder] of mappings) {
    const p = findParagraph(regex);
    if (!p) continue;
    if (/Students Name:/i.test(String(regex))) {
      setParagraphText(p, builder());
      continue;
    }
    const didSetStyledValue = setParagraphValueAfterLabel(p, builder(), regex);
    if (!didSetStyledValue) {
      setParagraphText(p, builder());
    }
  }

  const patientAgeDiagnosisParagraph = findParagraph(/Patient.?s Initial\s*and Age:/i);
  if (patientAgeDiagnosisParagraph) {
    const runs = Array.from(patientAgeDiagnosisParagraph.getElementsByTagNameNS(ns, 'r'));
    const runTextNodes = runs.map((run) => Array.from(run.getElementsByTagNameNS(ns, 't')));
    const runTexts = runTextNodes.map((nodes) => nodes.map((n) => n.textContent || '').join(''));
    const ageRunIndex = runTexts.findIndex((t) => /^\d+/.test((t || '').trim()) || /years?\s*old/i.test(t || ''));
    const diagnosisLabelRunIndex = runTexts.findIndex((t) => /Diagnosis:/i.test(t || ''));
    const diagnosisValueRunIndex = diagnosisLabelRunIndex !== -1
      ? runTexts.findIndex((t, idx) => idx > diagnosisLabelRunIndex && (t || '').trim())
      : -1;

    if (ageRunIndex !== -1 && runTextNodes[ageRunIndex]?.length) {
      runTextNodes[ageRunIndex][0].textContent = fields.patientInitialAge || FALLBACK_NA;
      for (let i = 1; i < runTextNodes[ageRunIndex].length; i += 1) runTextNodes[ageRunIndex][i].textContent = '';
    }
    if (diagnosisValueRunIndex !== -1 && runTextNodes[diagnosisValueRunIndex]?.length) {
      runTextNodes[diagnosisValueRunIndex][0].textContent = fields.diagnosis || FALLBACK_NOT_ASSESSED;
      for (let i = 1; i < runTextNodes[diagnosisValueRunIndex].length; i += 1) runTextNodes[diagnosisValueRunIndex][i].textContent = '';
    }
  }

  const theoristPromptIdx = paragraphs.findIndex((p) => /select nursing theorist/i.test(paragraphText(p)));
  if (theoristPromptIdx !== -1 && paragraphs[theoristPromptIdx + 1]) {
    setParagraphText(paragraphs[theoristPromptIdx + 1], fields.theorist || '');
  }

  const knowledgePromptIdx = paragraphs.findIndex((p) => /Knowledge gained from this/i.test(paragraphText(p)));
  if (knowledgePromptIdx !== -1 && paragraphs[knowledgePromptIdx + 1]) {
    setParagraphText(paragraphs[knowledgePromptIdx + 1], fields.knowledgeGained || '');
  }

  const objectivePromptIdx = paragraphs.findIndex((p) => /How were the course\/clinical objectives/i.test(paragraphText(p)));
  if (objectivePromptIdx !== -1 && paragraphs[objectivePromptIdx + 1]) {
    setParagraphText(paragraphs[objectivePromptIdx + 1], fields.courseObjectives || '');
  }

  // Remove this stock line from the psych template so medication rows focus on filled content only.
  const selectedOnCaseLogIdx = paragraphs.findIndex((p) => /Selected on case log/i.test(paragraphText(p)));
  if (selectedOnCaseLogIdx !== -1) {
    setParagraphText(paragraphs[selectedOnCaseLogIdx], '');
  }

  const researchPatterns = [
    /Research Article Title:/i,
    /The Journey of Adolescent Paranoia/i,
    /^Article Link:/i,
    /pmc\.ncbi\.nlm\.nih\.gov/i,
    /^Why I chose this article:/i,
    /^I chose this article because/i,
    /persecutory thinking/i,
    /patient.?s story of feeling persecuted/i,
    /^APA Citation:/i,
    /Bird,\s*J\.\s*C\./i,
    /Psychology and Psychotherapy/i,
    /doi\.org\//i,
  ];
  for (const p of paragraphs) {
    const text = paragraphText(p);
    if (researchPatterns.some((rx) => rx.test(text))) {
      setParagraphText(p, '');
    }
  }

  const medRows = ensurePopulatedMeds(meds).slice(0, MED_TEMPLATE_ROW_CAP);
  const tables = Array.from(xmlDoc.getElementsByTagNameNS(ns, 'tbl'));
  const medTable = tables.find((tbl) => {
    const text = Array.from(tbl.getElementsByTagNameNS(ns, 't')).map((n) => n.textContent || '').join(' ');
    return /Name of Drug, Prescribed Dose, Frequency\s*&\s*Route/i.test(text)
      && /Why is client receiving this drug/i.test(text)
      && /Nursing Implications/i.test(text);
  });

  if (medTable) {
    const rows = Array.from(medTable.getElementsByTagNameNS(ns, 'tr'));
    for (let r = 1; r < rows.length; r += 1) {
      const cells = Array.from(rows[r].getElementsByTagNameNS(ns, 'tc'));
      if (cells.length < 5) continue;
      const med = medRows[r - 1] || EMPTY_MED;
      const nameAndRoute = [clean(med.nameClass), clean(med.doseRoute)].filter(Boolean).join(' | ');
      setTableCellText(cells[0], nameAndRoute);
      setTableCellText(cells[1], clean(med.why));
      setTableCellText(cells[2], clean(med.action));
      setTableCellText(cells[3], clean(med.implications));
      setTableCellText(cells[4], clean(med.sideEffects));
    }
  }

  return new XMLSerializer().serializeToString(xmlDoc);
}

async function buildFilledNursingZip(templateUrl, fields, meds, articleResult = null) {
  const verification = await getTemplateVerification(templateUrl);
  const arrayBuffer = await loadTemplateArrayBuffer(templateUrl);

  if (verification.placeholderCount === 0 && verification.hasCoreWorksheetLabels) {
    let zip;
    try {
      zip = new PizZip(arrayBuffer);
    } catch {
      throw new Error('Could not parse template DOCX for fallback fill mode.');
    }
    const documentXmlFile = zip.file('word/document.xml');
    if (!documentXmlFile) throw new Error('Template DOCX is missing word/document.xml.');
    const updatedXml = applyLabelFillToXml(documentXmlFile.asText(), fields, meds);
    zip.file('word/document.xml', updatedXml);
    applyResearchCleanupToZip(zip);
    applyArticleAppendixToZip(zip, articleResult);
    return zip;
  }

  let zip;
  try {
    zip = new PizZip(arrayBuffer);
  } catch {
    throw new Error('Could not read template as a valid .docx file. Make sure nursing-process-template.docx is a real DOCX file (not legacy .doc).');
  }
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render(toTemplateData(fields, meds));

  const renderedZip = doc.getZip();
  applyResearchCleanupToZip(renderedZip);
  applyArticleAppendixToZip(renderedZip, articleResult);

  return renderedZip;
}

async function exportFilledDocxTemplate(templateUrl, fields, meds, articleResult = null) {
  const outputFilename = buildDatedOutputFilename(fields?.date);
  const renderedZip = await buildFilledNursingZip(templateUrl, fields, meds, articleResult);

  const out = renderedZip.generate({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });

  await saveBlobAsFile(out, outputFilename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}

async function getTemplateVerification(templateUrl) {
  const arrayBuffer = await loadTemplateArrayBuffer(templateUrl);
  let zip;
  try {
    zip = new PizZip(arrayBuffer);
  } catch {
    throw new Error('Template DOCX is invalid and cannot be parsed.');
  }

  const documentXmlFile = zip.file('word/document.xml');
  if (!documentXmlFile) {
    throw new Error('Template DOCX is missing word/document.xml.');
  }
  const documentXml = documentXmlFile.asText();
  const placeholderMatches = documentXml.match(/\{[#\/]?[a-zA-Z0-9_.]+\}/g) || [];
  const placeholders = Array.from(new Set(placeholderMatches));
  const missingRequired = REQUIRED_TEMPLATE_PLACEHOLDERS.filter((tag) => !placeholders.includes(tag));
  const hasCoreWorksheetLabels = [
    'Students Name:',
    'Date:',
    'Diagnosis:',
    'Nursing Diagnosis:',
    'Nursing Goal:',
    'Plan of Care:',
    'Intervention',
  ].every((label) => documentXml.includes(label));

  const digestBuffer = await crypto.subtle.digest('SHA-256', arrayBuffer);
  const digestHex = Array.from(new Uint8Array(digestBuffer))
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join('');

  return {
    fingerprint: digestHex.slice(0, 16),
    sizeBytes: arrayBuffer.byteLength,
    checkedAt: new Date().toLocaleString(),
    placeholderCount: placeholders.length,
    placeholderSample: placeholders.slice(0, 12),
    missingRequired,
    hasCoreWorksheetLabels,
  };
}

async function extractPdfText(file, onProgress) {
  const pdfjsLib = await import('pdfjs-dist');
  configurePdfJsWorker(pdfjsLib);

  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  let fullText = '';
  if (onProgress) onProgress(10, 'Reading PDF pages...');
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const strings = content.items.map((item) => item.str).join(' ');
    fullText += `\n${strings}\n`;
    if (onProgress) {
      const progress = Math.min(95, Math.round((i / pdf.numPages) * 100));
      onProgress(progress, `Extracting page ${i} of ${pdf.numPages}...`);
    }
  }
  if (onProgress) onProgress(100, 'PDF text extraction complete.');
  return normalizeSpacedPdfText(fullText) || fullText;
}

function normalizeAiFields(rawFields = {}) {
  const normalized = {};
  for (const [key] of FIELD_SECTIONS) {
    normalized[key] = clean(String(rawFields[key] || ''));
  }
  return normalized;
}

function sanitizeParsedFields(fields = {}, sourceText = '') {
  const localDate = extractEncounterDate(sourceText);
  const localAge = extractAgeFromText(sourceText);
  const localSex = extractSexFromText(sourceText);
  const localDiagnosis = extractSimpleDiagnosis(fields.diagnosis, sourceText);
  const localMedicalHistory = extractMedicalHistorySummary(sourceText);

  return {
    ...fields,
    date: localDate || clean(fields.date),
    age: localAge || clean(fields.age),
    patientInitialAge: localAge || clean(fields.patientInitialAge),
    sex: localSex || clean(fields.sex),
    diagnosis: localDiagnosis || cleanClinicalValue(fields.diagnosis, 40),
    medicalHistory: localMedicalHistory || cleanClinicalValue(fields.medicalHistory, 140),
    psychosocial: cleanClinicalValue(fields.psychosocial, 220),
    genAppearance: cleanClinicalValue(fields.genAppearance, 140),
  };
}

function sanitizeSimulationParsedFields(fields = {}) {
  return {
    ...fields,
    date: clean(fields.date),
    age: clean(fields.age),
    patientInitialAge: clean(fields.patientInitialAge),
    sex: clean(fields.sex),
    diagnosis: cleanClinicalValue(fields.diagnosis, 40),
    medicalHistory: cleanClinicalValue(fields.medicalHistory, 140),
    psychosocial: cleanClinicalValue(fields.psychosocial, 220),
    genAppearance: cleanClinicalValue(fields.genAppearance, 140),
  };
}

function mergeNewCaseFields(prev = {}, parsedFields = {}) {
  return {
    ...DEFAULT_STATE,
    ...parsedFields,
    studentName: prev.studentName || parsedFields.studentName || DEFAULT_STATE.studentName,
    week: parsedFields.week || '',
    semesterMeta: prev.semesterMeta || parsedFields.semesterMeta || '',
    courseMeta: prev.courseMeta || parsedFields.courseMeta || '',
    facultyMeta: prev.facultyMeta || parsedFields.facultyMeta || 'Karen Colombo',
    siteMeta: prev.siteMeta || parsedFields.siteMeta || '',
  };
}

function normalizeAiMeds(rawMeds = []) {
  if (!Array.isArray(rawMeds)) return [];
  const meds = rawMeds
    .map((med) => {
      const next = {};
      for (const key of MED_KEYS) {
        next[key] = clean(String(med?.[key] || ''));
      }
      return next;
    })
    .filter((med) => Object.values(med).some(Boolean));
  return capMedicationRows(meds);
}

async function parseCaseTextWithAi({ apiKey, caseText, metadata }) {
  const fieldSchemaProperties = Object.fromEntries(FIELD_SECTIONS.map(([key]) => [key, { type: 'string' }]));
  const medSchemaProperties = Object.fromEntries(MED_KEYS.map((key) => [key, { type: 'string' }]));

  const safeCaseText = clean(caseText).slice(0, AI_MAX_CASE_TEXT_CHARS);
  const systemPrompt = 'You extract nursing worksheet fields from a case log. Only use evidence found in the provided case text. If unknown, return an empty string. Do not invent details. Provide concise clinical phrasing for nursing diagnosis, goal, plan, intervention, rationale, evaluation, reassessment only when supported by case text. For medications, output one medication per array item (do not group classes). If side effects/actions are not in the case text, use concise medication-appropriate clinical guidance.';
  const userPrompt = `Metadata selected by user:\n- Student Name: ${metadata.studentName}\n- Semester: ${metadata.semesterMeta}\n- Course: ${metadata.courseMeta}\n- Clinical Faculty: ${metadata.facultyMeta}\n- Clinical Site: ${metadata.siteMeta}\n\nCase text to parse:\n${safeCaseText}`;

  const sendRequest = async (body) => {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`OpenAI request failed (${response.status}): ${errorText}`);
    }

    return response.json();
  };

  const extractMessageText = (data) => {
    const content = data?.choices?.[0]?.message?.content;
    const text = Array.isArray(content) ? content.map((part) => part?.text || '').join('') : content;
    return String(text || '').trim();
  };

  const parseJsonFromText = (text) => {
    if (!text) return {};
    try {
      return JSON.parse(text);
    } catch {
      const match = text.match(/\{[\s\S]*\}/);
      if (!match) throw new Error('AI response was not valid JSON.');
      return JSON.parse(match[0]);
    }
  };

  let parsed;
  try {
    const data = await sendRequest({
      model: 'gpt-4o-mini',
      temperature: 0.1,
      response_format: {
        type: 'json_schema',
        json_schema: {
          name: 'nursing_process_parse',
          schema: {
            type: 'object',
            additionalProperties: false,
            properties: {
              fields: {
                type: 'object',
                additionalProperties: false,
                properties: fieldSchemaProperties,
                required: FIELD_SECTIONS.map(([key]) => key),
              },
              medications: {
                type: 'array',
                items: {
                  type: 'object',
                  additionalProperties: false,
                  properties: medSchemaProperties,
                  required: MED_KEYS,
                },
              },
            },
            required: ['fields', 'medications'],
          },
        },
      },
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt },
      ],
    });
    parsed = parseJsonFromText(extractMessageText(data));
  } catch (primaryErr) {
    const fallbackData = await sendRequest({
      model: 'gpt-4o-mini',
      temperature: 0.1,
      response_format: { type: 'json_object' },
      messages: [
        { role: 'system', content: `${systemPrompt} Return valid JSON only with keys: fields, medications.` },
        { role: 'user', content: userPrompt },
      ],
    });
    parsed = parseJsonFromText(extractMessageText(fallbackData));
    if (!parsed?.fields) {
      throw primaryErr;
    }
  }

  return {
    fields: normalizeAiFields(parsed.fields),
    medications: normalizeAiMeds(parsed.medications),
  };
}

export default function App() {
  const [fields, setFields] = useState(DEFAULT_STATE);
  const [medications, setMedications] = useState(MED_DEFAULT);
  const [rawText, setRawText] = useState('');
  const [simulationNotesText, setSimulationNotesText] = useState('');
  const [simulationFileName, setSimulationFileName] = useState('');
  const [simulationParsed, setSimulationParsed] = useState(false);
  const [simulationStatus, setSimulationStatus] = useState('Load a simulation PDF, then parse it into the concept map.');
  const [loading, setLoading] = useState(false);
  const [aiLoading, setAiLoading] = useState(false);
  const [openAiApiKey, setOpenAiApiKey] = useState('');
  const [apiKeyVerified, setApiKeyVerified] = useState(false);
  const [hasSavedApiKey, setHasSavedApiKey] = useState(false);
  const [parseProgress, setParseProgress] = useState(0);
  const [parseProgressLabel, setParseProgressLabel] = useState('');
  const [status, setStatus] = useState('Step 1: Complete student/clinical details to unlock Case Intake automatically.');
  const [sequenceStep, setSequenceStep] = useState(1);
  const [isFinalized, setIsFinalized] = useState(false);
  const [outputProgress, setOutputProgress] = useState(0);
  const [outputStatus, setOutputStatus] = useState('Idle. Generate final output when ready.');
  const [outputLastRun, setOutputLastRun] = useState('');
  const [outputError, setOutputError] = useState('');
  const [templateReady, setTemplateReady] = useState(false);
  const [templateFingerprint, setTemplateFingerprint] = useState('');
  const [templateSizeBytes, setTemplateSizeBytes] = useState(0);
  const [templateCheckedAt, setTemplateCheckedAt] = useState('');
  const [templateCheckError, setTemplateCheckError] = useState('');
  const [templatePlaceholderCount, setTemplatePlaceholderCount] = useState(0);
  const [templatePlaceholderSample, setTemplatePlaceholderSample] = useState([]);
  const [templateMissingRequired, setTemplateMissingRequired] = useState([]);
  const [viewMode, setViewMode] = useState('split');
  const [showStep3Fields, setShowStep3Fields] = useState(false);
  const [showStep3Meds, setShowStep3Meds] = useState(false);
  const [articlePrompt, setArticlePrompt] = useState('');
  const [articleSecondaryPrompt, setArticleSecondaryPrompt] = useState('');
  const [articleLoading, setArticleLoading] = useState(false);
  const [articleWhyLoading, setArticleWhyLoading] = useState(false);
  const [articleProgress, setArticleProgress] = useState(0);
  const [articleStatus, setArticleStatus] = useState('Add a focus prompt and run article search after final output.');
  const [articleResult, setArticleResult] = useState(null);
  const [articleError, setArticleError] = useState('');
  const [articleHistory, setArticleHistory] = useState([]);
  const [sessionFacultyName, setSessionFacultyName] = useState('');
  const fileInputRef = useRef(null);
  const simulationFileInputRef = useRef(null);
  const hasAutoAdvancedStep1Ref = useRef(false);

  const metadataReady = useMemo(() => {
    const required = ['studentName', 'facultyMeta'];
    return required.every((k) => fields[k]?.trim());
  }, [fields]);

  const missing = useMemo(() => {
    const required = ['studentName', 'date', 'patientInitialAge', 'diagnosis', 'psychosocial', 'nursingDiagnosis', 'goal', 'plan', 'intervention'];
    return required.filter((k) => !fields[k]?.trim());
  }, [fields]);
  const fieldReviewItems = useMemo(() => getFieldReviewItems(fields), [fields]);
  const medicationReviewItems = useMemo(() => getMedicationReviewItems(medications), [medications]);
  const reviewItemCount = fieldReviewItems.length + medicationReviewItems.length;
  const conceptMapData = useMemo(() => buildConceptMapData(fields, medications, rawText), [fields, medications, rawText]);
  const conceptMapReviewItems = useMemo(() => Object.entries(conceptMapData)
    .map(([key, value]) => {
      if (isPriorityNursingKey(key)) return null;
      const reason = getReviewReasonForValue(value);
      return reason ? { key, label: CONCEPT_MAP_OUTPUT_LABELS[key] || key, reason } : null;
    })
    .filter(Boolean), [conceptMapData]);
  const workflowSignals = useMemo(() => [
    {
      label: 'Set Up',
      detail: metadataReady ? 'Student and faculty ready' : 'Add student name',
      state: metadataReady ? 'complete' : 'active',
    },
    {
      label: 'Case PDF',
      detail: rawText.trim() ? 'PDF text captured' : 'Upload a case PDF',
      state: rawText.trim() ? 'complete' : (metadataReady ? 'active' : 'pending'),
    },
    {
      label: 'Auto Fill',
      detail: sequenceStep >= 3 ? 'Concept map fields populated' : 'Fill the map from the case',
      state: sequenceStep >= 3 ? 'complete' : (rawText.trim() ? 'active' : 'pending'),
    },
    {
      label: 'Export',
      detail: sequenceStep >= 3 ? 'Download the formatted PDF' : 'Ready after auto fill',
      state: sequenceStep >= 3 ? 'active final' : 'pending',
    },
  ], [metadataReady, rawText, sequenceStep]);

  const docPreview = useMemo(() => buildDocText(fields, medications), [fields, medications]);
  const articlePreview = useMemo(() => {
    if (!articleResult) return '';
    return [
      'Research Article Title:',
      articleResult.title,
      '',
      'Article Link:',
      articleResult.link,
      '',
      'Why I chose this article:',
      articleResult.why,
      '',
      'APA Citation:',
      articleResult.citation,
    ].join('\n');
  }, [articleResult]);
  const canAddMedicationRow = medications.length < MED_TEMPLATE_ROW_CAP;
  const updateField = (key, value) => setFields((prev) => ({ ...prev, [key]: value }));
  const hasApiKey = Boolean(openAiApiKey.trim());
  const facultyOptions = useMemo(() => {
    const options = [...FACULTY_OPTIONS];
    [fields.facultyMeta, sessionFacultyName].forEach((name) => {
      const normalized = normalizeFacultyName(name);
      if (normalized && !options.includes(normalized)) options.push(normalized);
    });
    return options;
  }, [fields.facultyMeta, sessionFacultyName]);

  const applySessionFacultyName = () => {
    const nextName = normalizeFacultyName(sessionFacultyName);
    if (!nextName) {
      setStatus('Enter an instructor name to use for this session.');
      return;
    }
    setSessionFacultyName(nextName);
    updateField('facultyMeta', nextName);
    setStatus(`${nextName} is selected for this session only.`);
  };

  useEffect(() => {
    setHasSavedApiKey(Boolean(localStorage.getItem(API_KEY_STORAGE_KEY)));
  }, []);

  useEffect(() => {
    if (sequenceStep >= 3) {
      setShowStep3Fields(true);
      setShowStep3Meds(true);
    }
  }, [sequenceStep]);

  useEffect(() => {
    const normalizedFaculty = normalizeFacultyName(fields.facultyMeta);
    if (fields.facultyMeta && normalizedFaculty !== fields.facultyMeta) {
      updateField('facultyMeta', normalizedFaculty);
    }
  }, [fields.facultyMeta]);

  useEffect(() => {
    if (metadataReady && sequenceStep === 1 && !hasAutoAdvancedStep1Ref.current) {
      hasAutoAdvancedStep1Ref.current = true;
      setSequenceStep(2);
      setStatus('Step 2 unlocked: case intake is ready. Upload a PDF or paste case text to continue.');
    }
  }, [metadataReady, sequenceStep]);

  const checkTemplateReady = async () => {
    setTemplateCheckError('');
    setTemplatePlaceholderCount(0);
    setTemplatePlaceholderSample([]);
    setTemplateMissingRequired([]);
    try {
      const data = await getTemplateVerification(DEFAULT_CONCEPT_MAP_TEMPLATE_URL);
      setTemplateReady(true);
      setTemplateFingerprint(data.fingerprint);
      setTemplateSizeBytes(data.sizeBytes);
      setTemplateCheckedAt(data.checkedAt);
      setTemplatePlaceholderCount(data.placeholderCount);
      setTemplatePlaceholderSample(data.placeholderSample);
      setTemplateMissingRequired([]);
      setOutputStatus(`Concept map template verified (${data.fingerprint}). Ready for export.`);
    } catch (err) {
      setTemplateReady(false);
      setTemplateFingerprint('');
      setTemplateSizeBytes(0);
      setTemplateCheckedAt('');
      setTemplatePlaceholderCount(0);
      setTemplatePlaceholderSample([]);
      setTemplateMissingRequired([]);
      setTemplateCheckError(err.message || 'Template check failed.');
      setOutputStatus('Concept map template verification failed. Export is blocked.');
    }
  };

  useEffect(() => {
    checkTemplateReady();
  }, []);

  const setProgress = (value, label) => {
    setParseProgress(value);
    setParseProgressLabel(label || '');
  };

  const applyParsedText = (text) => {
    const parsed = parseCaseText(text);
    const sanitizedFields = sanitizeParsedFields(parsed.fields, text);
    const bestMatchedFields = ensureConceptMapFields(clearPriorityNursingFields(sanitizedFields), text, parsed.medications);
    setFields((prev) => mergeNewCaseFields(prev, bestMatchedFields));
    setMedications(ensurePopulatedMeds(parsed.medications));
    setRawText(parsed.rawText);
    setSequenceStep((prev) => Math.max(prev, 3));
    setIsFinalized(false);
    setStatus(
      parsed.medications.length > MED_TEMPLATE_ROW_CAP
        ? `Source parse complete. Medication rows are capped at ${MED_TEMPLATE_ROW_CAP}; extra entries were not added. Review flagged fields before export.`
        : 'Source parse complete. Concept-map values were added; app-generated items are flagged for review.'
    );
  };

  const applySimulationNotesText = (text, sourceHint = '') => {
    const parsed = parseSimulationNotesText(text, sourceHint);
    const sanitizedFields = sanitizeSimulationParsedFields(parsed.fields);
    const bestMatchedFields = ensureConceptMapFields(sanitizedFields, parsed.rawText, parsed.medications);
    setFields((prev) => mergeNewCaseFields(prev, bestMatchedFields));
    setMedications(ensurePopulatedMeds(parsed.medications));
    setRawText(parsed.rawText);
    setSimulationNotesText(parsed.rawText);
    setSimulationParsed(true);
    setSequenceStep((prev) => Math.max(prev, 3));
    setIsFinalized(false);
    setSimulationStatus('Simulation notes parsed into the concept map. Review fields, then download the beta PDF.');
    setStatus('Simulation notes beta filled the concept map. Review highlighted fields before exporting.');
    setOutputStatus('Simulation notes beta applied. Export when the fields look right.');
  };

  const handleSimulationFile = async (file) => {
    if (!file) return;
    setLoading(true);
    setProgress(5, 'Preparing simulation PDF extraction...');
    setSimulationFileName(file.name || '');
    setSimulationParsed(false);
    setSimulationStatus('Extracting simulation notes from PDF...');
    setStatus('Extracting simulation PDF text...');
    try {
      const text = await extractPdfText(file, (value, label) => setProgress(value, label));
      setSimulationNotesText(clean(text));
      setProgress(100, 'Simulation PDF text ready.');
      setSimulationStatus('Simulation notes loaded. Click Parse Simulation Notes to fill the concept map.');
      setStatus('Simulation PDF text extracted in the beta section.');
    } catch (err) {
      console.error(err);
      setProgress(0, '');
      setSimulationStatus('Simulation PDF extraction failed. Paste the notes into the beta source box.');
      setStatus('Simulation PDF extraction failed. Paste notes manually in the beta section.');
    } finally {
      setLoading(false);
    }
  };

  const clearSimulationBeta = () => {
    setSimulationNotesText('');
    setSimulationFileName('');
    setSimulationParsed(false);
    setSimulationStatus('Load a simulation PDF, then parse it into the concept map.');
    setFields((prev) => clearPriorityNursingFields({
      ...prev,
      nd1Assessment: '',
      nd1Diagnosis: '',
      nd1Rationale: '',
      nd1Intervention: '',
      nd1Evaluation: '',
      nd2Assessment: '',
      nd2Diagnosis: '',
      nd2Rationale: '',
      nd2Intervention: '',
      nd2Evaluation: '',
      nd3Assessment: '',
      nd3Diagnosis: '',
      nd3Rationale: '',
      nd3Intervention: '',
      nd3Evaluation: '',
    }));
    setStatus('Simulation beta cleared.');
  };

  const handleFile = async (file) => {
    if (!file) return;
    setLoading(true);
    setProgress(5, 'Preparing PDF extraction...');
    setStatus('Extracting text from PDF...');
    setFields((prev) => mergeNewCaseFields(prev, {}));
    setMedications(MED_DEFAULT.map((m) => ({ ...m })));
    setIsFinalized(false);
    setOutputProgress(0);
    setOutputStatus('Idle. Generate final output when ready.');
    setOutputLastRun('');
    setOutputError('');
    try {
      const text = await extractPdfText(file, (value, label) => setProgress(value, label));
      setRawText(clean(text));
      setIsFinalized(false);
      setProgress(100, 'PDF text ready.');
      setStatus('PDF text extracted. Click Auto Fill From Case PDF to fill the concept map.');
    } catch (err) {
      console.error(err);
      setProgress(0, '');
      setStatus('PDF extraction failed. Paste the case text manually below.');
    } finally {
      setLoading(false);
    }
  };

  const applyAiParsedText = async () => {
    if (!openAiApiKey.trim()) {
      setStatus('Enter your OpenAI API key to use AI parsing.');
      return;
    }
    if (!rawText.trim()) {
      setStatus('Upload a PDF or paste case text before AI parsing.');
      return;
    }

    setAiLoading(true);
    setProgress(10, 'Preparing AI parsing request...');
    setStatus('AI parsing in progress...');
    try {
      setProgress(30, 'Sending case text to AI parser...');
      const parsed = await parseCaseTextWithAi({
        apiKey: openAiApiKey.trim(),
        caseText: rawText,
        metadata: {
          studentName: fields.studentName,
          semesterMeta: fields.semesterMeta,
          courseMeta: fields.courseMeta,
          facultyMeta: fields.facultyMeta,
          siteMeta: fields.siteMeta,
        },
      });
      setProgress(75, 'Applying AI output to worksheet fields...');
      const localParsed = parseCaseText(rawText);
      const sanitizedFields = sanitizeParsedFields({
        ...localParsed.fields,
        ...parsed.fields,
      }, rawText);
      const bestMatchedFields = ensureConceptMapFields(clearPriorityNursingFields(sanitizedFields), rawText, parsed.medications);

      setFields((prev) => mergeNewCaseFields(prev, bestMatchedFields));
      setMedications(ensurePopulatedMeds(parsed.medications));
      setSequenceStep((prev) => Math.max(prev, 3));
      setIsFinalized(false);
      setApiKeyVerified(true);
      setProgress(100, 'AI parsing complete.');
      setStatus(
        parsed.medications.length > MED_TEMPLATE_ROW_CAP
          ? `AI source parse complete. Medication rows are capped at ${MED_TEMPLATE_ROW_CAP}; extra entries were not added. Review flagged fields before export.`
          : 'AI source parse complete. Concept-map values were added; AI-generated items are flagged for review.'
      );
    } catch (err) {
      console.error(err);
      setApiKeyVerified(false);
      setProgress(0, '');
      setStatus(`AI parsing failed: ${err.message}`);
    } finally {
      setAiLoading(false);
    }
  };

  const saveApiKeyLocal = () => {
    if (!openAiApiKey.trim()) {
      setStatus('Enter an OpenAI API key first, then save it.');
      return;
    }
    localStorage.setItem(API_KEY_STORAGE_KEY, openAiApiKey.trim());
    setHasSavedApiKey(true);
    setApiKeyVerified(false);
    setStatus('API key saved locally in this browser.');
  };

  const loadSavedApiKey = () => {
    const raw = localStorage.getItem(API_KEY_STORAGE_KEY);
    if (!raw) {
      setStatus('No saved API key found.');
      return;
    }

    setOpenAiApiKey(raw);
    setApiKeyVerified(false);
    setStatus('Saved API key loaded for this session.');
  };

  const clearSavedApiKey = () => {
    localStorage.removeItem(API_KEY_STORAGE_KEY);
    setHasSavedApiKey(false);
    setStatus('Saved API key removed from this browser.');
  };

  const downloadFormattedOutput = async () => {
    if (!isFinalized) {
      setStatus('Generate final output first, then download formatted output.');
      return;
    }
    if (!templateReady) {
      setStatus('Template is not ready. Ensure the required template file exists and is valid.');
      return;
    }

    setLoading(true);
    setOutputError('');
    setOutputStatus('Building formatted nursing DOCX from required template...');
    setOutputProgress(20);
    try {
      setOutputProgress(55);
      await exportFilledDocxTemplate(DEFAULT_TEMPLATE_URL, fields, medications, null);
      setOutputProgress(100);
      const stamp = new Date().toLocaleString();
      setOutputLastRun(stamp);
      setOutputStatus(`Nursing template downloaded successfully at ${stamp}.`);
      setStatus('Nursing template DOCX downloaded successfully.');
    } catch (err) {
      console.error(err);
      setOutputProgress(0);
      setOutputError(err.message || 'Unknown export error.');
      setOutputStatus('Formatted output download failed.');
      setStatus(`Formatted output download failed. ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const downloadFormattedOutputWithArticle = async () => {
    if (!isFinalized) {
      setStatus('Generate final output first, then download nursing template + article.');
      return;
    }
    if (!templateReady) {
      setStatus('Template is not ready. Ensure the required template file exists and is valid.');
      return;
    }
    if (!articleResult) {
      setStatus('Run Step 6 article search first, then download nursing template + article.');
      return;
    }

    setLoading(true);
    setOutputError('');
    setOutputStatus('Building formatted nursing DOCX with article appendix...');
    setOutputProgress(20);
    try {
      setOutputProgress(55);
      await exportFilledDocxTemplate(DEFAULT_TEMPLATE_URL, fields, medications, articleResult);
      setOutputProgress(100);
      const stamp = new Date().toLocaleString();
      setOutputLastRun(stamp);
      setOutputStatus(`Nursing + article DOCX downloaded successfully at ${stamp}.`);
      setStatus('Nursing + article DOCX downloaded successfully.');
    } catch (err) {
      console.error(err);
      setOutputProgress(0);
      setOutputError(err.message || 'Unknown export error.');
      setOutputStatus('Nursing + article download failed.');
      setStatus(`Nursing + article download failed. ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const downloadArticleOnlyOutput = async () => {
    if (!isFinalized) {
      setArticleStatus('Generate final output first, then download article output.');
      return;
    }
    if (!templateReady) {
      setArticleStatus('Template is not ready. Article DOCX export is blocked.');
      return;
    }
    if (!articleResult) {
      setArticleStatus('Run article search first, then download article output.');
      return;
    }

    setLoading(true);
    setArticleError('');
    setArticleProgress(20);
    setArticleStatus('Building article-only DOCX...');
    try {
      setArticleProgress(70);
      await exportArticleOnlyDocxTemplate(DEFAULT_TEMPLATE_URL, articleResult, fields);
      setArticleProgress(100);
      setArticleStatus('Article-only DOCX downloaded successfully.');
    } catch (err) {
      console.error(err);
      setArticleProgress(0);
      setArticleError(err.message || 'Article-only export failed.');
      setArticleStatus('Article-only DOCX download failed.');
    } finally {
      setLoading(false);
    }
  };

  const downloadNursingWithArticlePdf = async () => {
    if (!isFinalized) {
      setArticleStatus('Generate final output first, then download nursing + article PDF.');
      return;
    }
    if (!articleResult) {
      setArticleStatus('Run Step 6 article search first, then download nursing + article PDF.');
      return;
    }

    setLoading(true);
    setArticleError('');
    setArticleProgress(20);
    setArticleStatus('Preparing high-fidelity source DOCX for PDF conversion...');
    try {
      await exportFilledDocxTemplate(DEFAULT_TEMPLATE_URL, fields, medications, articleResult);
      setArticleProgress(100);
      const suggestedFilename = buildDatedNursingArticlePdfFilename(fields?.date);
      setArticleStatus(`Downloaded DOCX source. For closest 1:1 PDF, open in Word/Pages and Save as PDF as ${suggestedFilename}.`);
      setStatus('Downloaded Nursing + Article DOCX source for highest-fidelity PDF conversion.');
    } catch (err) {
      console.error(err);
      setArticleProgress(0);
      setArticleError(err.message || 'Nursing + article PDF export failed.');
      setArticleStatus('Nursing + article source DOCX export failed.');
    } finally {
      setLoading(false);
    }
  };

  const downloadConceptMapOutput = async () => {
    if (sequenceStep < 3) {
      setStatus('Parse a case first, then export the concept map.');
      return;
    }
    if (!templateReady) {
      setStatus('Concept map template is not ready. Check public/templates/concept-map-template.docx.');
      return;
    }
    setLoading(true);
    setOutputError('');
    setOutputProgress(15);
    setOutputStatus(`Building concept map DOCX from ${DEFAULT_CONCEPT_MAP_TEMPLATE_NAME}...`);
    try {
      setOutputProgress(60);
      await exportConceptMapTemplate(DEFAULT_CONCEPT_MAP_TEMPLATE_URL, fields, medications, rawText);
      setIsFinalized(true);
      setSequenceStep((prev) => Math.max(prev, 4));
      setOutputProgress(100);
      const stamp = new Date().toLocaleString();
      setOutputLastRun(stamp);
      setOutputStatus(`Concept map DOCX downloaded successfully at ${stamp}.`);
      setStatus('Concept map DOCX downloaded successfully.');
    } catch (err) {
      console.error(err);
      setOutputProgress(0);
      setOutputError(err.message || 'Concept map export failed.');
      setOutputStatus('Concept map export failed.');
      setStatus(`Concept map export failed. ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const downloadConceptMapPdfOutput = async ({ includePriorityNursingSections = false } = {}) => {
    if (sequenceStep < 3) {
      setStatus('Parse a case first, then export the formatted concept map source.');
      return;
    }
    if (!templateReady) {
      setStatus('Concept map template is not ready. Check public/templates/concept-map-template.docx.');
      return;
    }

    setLoading(true);
    setOutputError('');
    setOutputProgress(20);
    setOutputStatus(includePriorityNursingSections
      ? 'Building formatted concept map PDF with nursing priorities...'
      : 'Building formatted concept map PDF with blank nursing-priority sections...');
    try {
      await exportConceptMapPdf(fields, medications, rawText, {
        blankPriorityNursingSections: !includePriorityNursingSections,
      });
      setIsFinalized(true);
      setSequenceStep((prev) => Math.max(prev, 4));

      setOutputProgress(100);
      const stamp = new Date().toLocaleString();
      setOutputLastRun(stamp);
      setOutputStatus(`Concept map PDF downloaded successfully at ${stamp}.`);
      setStatus(includePriorityNursingSections
        ? 'Concept map PDF downloaded with nursing priorities included.'
        : 'Concept map PDF downloaded with nursing-priority sections blank.');
    } catch (err) {
      console.error(err);
      setOutputProgress(0);
      setOutputError(err.message || 'Concept map PDF export failed.');
      setOutputStatus('Concept map PDF export failed.');
      setStatus(`Concept map PDF export failed. ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const downloadFinalPdfSetSources = async () => {
    if (!isFinalized) {
      setStatus('Generate final output first, then download the final PDF set sources.');
      return;
    }
    if (!templateReady) {
      setStatus('Template is not ready. Final PDF set source export is blocked.');
      return;
    }
    if (!articleResult) {
      setStatus('Run Step 6 article search first, then download the final PDF set sources.');
      return;
    }

    setLoading(true);
    setOutputError('');
    setArticleError('');
    setOutputProgress(20);
    setArticleProgress(20);
    setOutputStatus('Preparing nursing + article source DOCX...');
    setArticleStatus('Preparing final PDF set source files...');

    try {
      await exportFilledDocxTemplate(DEFAULT_TEMPLATE_URL, fields, medications, articleResult);
      setOutputProgress(60);
      setOutputStatus('Preparing concept map source DOCX...');

      await exportConceptMapTemplate(DEFAULT_CONCEPT_MAP_TEMPLATE_URL, fields, medications, rawText);
      setOutputProgress(100);
      setArticleProgress(100);

      const nursingPdfName = buildDatedNursingArticlePdfFilename(fields?.date);
      const conceptPdfName = buildDatedConceptMapPdfFilename(fields?.date);
      const stamp = new Date().toLocaleString();

      setOutputLastRun(stamp);
      setOutputStatus(`Final PDF set source DOCX files downloaded at ${stamp}. Convert to PDF as ${nursingPdfName} and ${conceptPdfName}.`);
      setArticleStatus('Final PDF set source DOCX files downloaded successfully.');
      setStatus('Final PDF set sources downloaded. Open both DOCX files in Word/Pages and Save as PDF.');
    } catch (err) {
      console.error(err);
      setOutputProgress(0);
      setArticleProgress(0);
      setOutputError(err.message || 'Final PDF set source export failed.');
      setOutputStatus('Final PDF set source export failed.');
      setArticleStatus('Final PDF set source export failed.');
      setStatus(`Final PDF set source export failed. ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const downloadCombinedFullDocx = async () => {
    if (!isFinalized) {
      setStatus('Generate final output first, then download the combined full DOCX.');
      return;
    }
    if (!templateReady) {
      setStatus('Template is not ready. Combined full DOCX export is blocked.');
      return;
    }
    if (!articleResult) {
      setStatus('Run Step 6 article search first, then download the combined full DOCX.');
      return;
    }

    setLoading(true);
    setOutputError('');
    setOutputProgress(20);
    setOutputStatus('Building nursing + article section for combined DOCX...');

    try {
      const nursingZip = await buildFilledNursingZip(DEFAULT_TEMPLATE_URL, fields, medications, articleResult);
      setOutputProgress(55);
      setOutputStatus('Building concept map section for combined DOCX...');

      const conceptZip = await buildFilledConceptMapZip(DEFAULT_CONCEPT_MAP_TEMPLATE_URL, fields, medications, rawText);
      const nursingDoc = nursingZip.file('word/document.xml');
      const conceptDoc = conceptZip.file('word/document.xml');
      if (!nursingDoc || !conceptDoc) {
        throw new Error('Could not locate document.xml in one of the source DOCX files.');
      }

      const mergedDocumentXml = mergeDocumentXmlBodies(nursingDoc.asText(), conceptDoc.asText());
      nursingZip.file('word/document.xml', mergedDocumentXml);

      const out = nursingZip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      const filename = buildDatedCombinedPackageFilename(fields?.date);
      await saveBlobAsFile(out, filename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

      setOutputProgress(100);
      const stamp = new Date().toLocaleString();
      setOutputLastRun(stamp);
      setOutputStatus(`Combined full DOCX downloaded successfully at ${stamp}.`);
      setStatus('Combined full DOCX downloaded successfully (nursing process + article + concept map).');
    } catch (err) {
      console.error(err);
      setOutputProgress(0);
      setOutputError(err.message || 'Combined full DOCX export failed.');
      setOutputStatus('Combined full DOCX export failed.');
      setStatus(`Combined full DOCX export failed. ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const continueToCaseIntake = () => {
    if (!metadataReady) {
      setStatus('Complete student name, semester, course, faculty, and clinical site before continuing.');
      return;
    }
    setSequenceStep((prev) => Math.max(prev, 2));
    hasAutoAdvancedStep1Ref.current = true;
    setStatus('Step 2: Upload a case PDF or paste case text, then parse to fill worksheet fields.');
  };

  const fillSpacingTestData = () => {
    const testFields = {
      studentName: fields.studentName || 'Sample Student',
      facultyMeta: fields.facultyMeta || 'Karen Colombo',
      date: '5/28/2026',
      clientName: 'J.D.',
      age: '84',
      patientInitialAge: '84',
      sex: 'F',
      ht: "5'4\"",
      wt: '120 lb',
      diagnosis: 'Weakness',
      temp: '98.6 F',
      pulse: '88 bpm',
      resp: '18/min',
      bp: '145/85',
      pain: '3/10',
      allergies: 'None',
      immunizations: 'Up to date',
      sleepPattern: 'Normal sleep pattern',
      nutritionalStatus: 'Nourished',
      assistanceAdls: 'Needs assistance',
      hygiene: 'Assisted hygiene',
      medsPrior: 'Valsartan; amlodipine; Eliquis',
      medicalHistory: 'Atrial fibrillation; CHF; hypertension; AV block; possible UTI',
      surgicalHistory: 'None',
      supportSystem: 'Family involved in care',
      responseHospitalization: 'Cooperative; mild anxiety about weakness',
      genAppearance: 'Fatigued; cooperative',
      ivLocation: 'Left arm',
      surgicalIncision: 'None',
      orientation: 'A&O x3',
      speech: 'Clear',
      weakness: 'Generalized',
      skinTurgor: 'Decreased',
      breathSounds: 'Clear',
      peripheralPulses: 'Present',
      edema: 'None noted',
      bowelSounds: 'Active',
      physicalOther: 'Fall risk',
      spiritualAssessment: 'No needs stated',
      labTestName: 'WBC / urinalysis',
      labClientResults: 'WBC 14; urine WBCs/bacteria; hematuria',
      labNormalValue: 'WBC 4.5-11.0; UA negative bacteria',
      labInterpretation: 'Possible infection/UTI; monitor hematuria',
      currentMedDate: '5/28/2026',
      currentMedOrder: 'Valsartan; amlodipine; furosemide; Eliquis; cefdinir',
      currentMedIndication: 'Manage cardiac history; treat possible UTI',
      nd1Assessment: withPriorityPrompt('nd1Assessment', 'Elevated WBC and urine findings suggest infection.'),
      nd1Diagnosis: withPriorityPrompt('nd1Diagnosis', 'Impaired urinary elimination'),
      nd1Rationale: withPriorityPrompt('nd1Rationale', 'Abnormal urine findings and elevated WBC support urinary infection concerns.'),
      nd1Intervention: 'Monitor temp, urine, WBC trends, and antibiotics as ordered.',
      nd1Evaluation: 'Patient remains monitored; response pending.',
      nd2Assessment: withPriorityPrompt('nd2Assessment', 'Age, weakness, and Eliquis increase fall/bleeding risk.'),
      nd2Diagnosis: withPriorityPrompt('nd2Diagnosis', 'Risk for falls'),
      nd2Rationale: withPriorityPrompt('nd2Rationale', 'Weakness and anticoagulation increase injury risk.'),
      nd2Intervention: 'Maintain fall precautions and assist ambulation.',
      nd2Evaluation: 'No injury noted during shift.',
      nd3Assessment: withPriorityPrompt('nd3Assessment', 'Mild pain 3/10 with fatigue and SOB.'),
      nd3Diagnosis: withPriorityPrompt('nd3Diagnosis', 'Activity intolerance'),
      nd3Rationale: withPriorityPrompt('nd3Rationale', 'Fatigue/SOB can limit activity tolerance.'),
      nd3Intervention: 'Cluster care, pace activity, monitor respiratory status.',
      nd3Evaluation: 'Tolerance requires reassessment.',
    };

    const testMeds = capMedicationRows([
      { ...EMPTY_MED, nameClass: 'Valsartan', doseRoute: 'PO daily', why: 'Hypertension/CHF' },
      { ...EMPTY_MED, nameClass: 'Eliquis', doseRoute: 'PO BID', why: 'Atrial fibrillation anticoagulation' },
      { ...EMPTY_MED, nameClass: 'Cefdinir', doseRoute: 'PO as ordered', why: 'Possible UTI/infection' },
    ]);

    setFields((prev) => {
      const next = { ...prev };
      Object.entries(testFields).forEach(([key, value]) => {
        if (!clean(String(next[key] || ''))) next[key] = value;
      });
      return next;
    });
    setMedications((prev) => {
      const hasExisting = Array.isArray(prev) && prev.some((med) => Object.values(med || {}).some((value) => clean(String(value || ''))));
      return hasExisting ? prev : testMeds;
    });
    setRawText((prev) => prev || 'Spacing test data generated by the app for concept-map PDF layout review.');
    setSequenceStep(4);
    setIsFinalized(true);
    setOutputError('');
    setOutputStatus('Blank spacing-test fields filled. Existing case values were preserved.');
    setStatus('Blank spacing-test fields filled without overwriting parsed/manual case values.');
  };

  const finalizeOutput = async () => {
    if (sequenceStep < 3) {
      setStatus('Parse a case first so worksheet fields and medication log are populated.');
      return;
    }
    if (!templateReady) {
      setStatus('Template is not ready. Ensure the required template file exists and is valid.');
      return;
    }

    setLoading(true);
    setOutputError('');
    setOutputProgress(10);
    setOutputStatus('Checking required fields and preparing final text output...');
    setStatus('Generating final output...');
    try {
      setOutputProgress(45);
      setIsFinalized(true);
      setSequenceStep(4);
      setOutputProgress(100);
      setOutputStatus('Final output ready. Review fields or click Download PDF.');
      setStatus('Step 4 complete: final text output generated.');
    } catch (err) {
      console.error(err);
      setOutputProgress(0);
      setOutputError(err.message || 'Unknown generation error.');
      setOutputStatus('Final output generation failed.');
      setStatus(`Final output generation failed. ${err.message}`);
    } finally {
      setLoading(false);
    }
  };

  const updateMedication = (idx, key, value) => {
    setMedications((prev) => prev.map((m, i) => (i === idx ? { ...m, [key]: value, aiGenerated: false } : m)));
  };

  const developPriorityNursingBeta = () => {
    if (sequenceStep < 3) {
      setStatus('Auto fill a case first, then use the beta nursing priorities option.');
      return;
    }
    const priorityFields = buildPriorityNursingFields(fields, rawText);
    setFields((prev) => ({ ...prev, ...priorityFields }));
    setShowStep3Fields(true);
    setSequenceStep((prev) => Math.max(prev, 3));
    setIsFinalized(false);
    setStatus('Beta nursing priorities added. These are AI-generated testing support and must be reviewed before export.');
    setOutputStatus('Nursing priority boxes filled by beta helper. Verify/edit before downloading PDF.');
  };

  const addMedicationRow = () => {
    if (!canAddMedicationRow) {
      setStatus(`Medication log is capped at ${MED_TEMPLATE_ROW_CAP} rows for this template.`);
      return;
    }
    setMedications((prev) => [...prev, structuredClone(EMPTY_MED)]);
  };

  const updateArticleWhy = (value) => {
    setArticleResult((prev) => (prev ? { ...prev, why: value } : prev));
  };

  const generateArticleWhyWithAi = async () => {
    if (!articleResult) {
      setArticleStatus('Find an article first, then generate the AI relevance write-up.');
      return;
    }
    if (!openAiApiKey.trim()) {
      setArticleStatus('Enter an OpenAI API key to generate AI write-up text.');
      return;
    }

    setArticleWhyLoading(true);
    setArticleError('');
    setArticleStatus('Generating AI write-up for why this article is relevant...');
    try {
      const aiWhy = await draftArticleRelevanceWithAi({
        apiKey: openAiApiKey,
        fields,
        rawText,
        userPrompt: articlePrompt,
        articleResult,
      });
      if (aiWhy) {
        setArticleResult((prev) => (prev ? { ...prev, why: aiWhy } : prev));
      }
      setArticleStatus('AI write-up generated. You can edit the "Why it relates" text before export.');
    } catch (err) {
      console.error(err);
      setArticleError(err.message || 'AI write-up generation failed.');
      setArticleStatus('AI write-up generation failed.');
    } finally {
      setArticleWhyLoading(false);
    }
  };

  const findArticle = async (preferNew = false) => {
    if (sequenceStep < 4) {
      setArticleStatus('Generate final output first, then run article search.');
      return;
    }

    setArticleLoading(true);
    setArticleError('');
    setArticleProgress(8);
    setArticleStatus('Preparing scholarly search query...');

    try {
      setArticleProgress(22);
      setArticleStatus('Searching Europe PMC for scholarly/scientific sources...');
      const activePrompt = preferNew
        ? clean(articleSecondaryPrompt || articlePrompt || '')
        : '';
      const excludedLinks = preferNew
        ? Array.from(new Set([
          ...articleHistory.map((item) => item?.link),
          articleResult?.link,
        ].filter(Boolean)))
        : [];
      const excludedTitles = preferNew
        ? Array.from(new Set([
          ...articleHistory.map((item) => item?.title),
          articleResult?.title,
        ].filter(Boolean)))
        : [];

      const result = await findScholarlyArticle({
        fields,
        rawText,
        userPrompt: activePrompt,
        excludedLinks,
        excludedTitles,
      });
      let nextResult = result;
      if (openAiApiKey.trim()) {
        try {
          setArticleStatus('Article found. Generating AI write-up for relevance...');
          const aiWhy = await draftArticleRelevanceWithAi({
            apiKey: openAiApiKey,
            fields,
            rawText,
            userPrompt: activePrompt,
            articleResult: result,
          });
          if (aiWhy) nextResult = { ...result, why: aiWhy };
        } catch (err) {
          console.error(err);
        }
      }
      setArticleProgress(76);
      setArticleStatus('Building relevance summary and APA citation...');
      setArticleResult(nextResult);
      setArticleHistory((prev) => {
        const next = [...prev, { title: nextResult.title, link: nextResult.link }];
        const deduped = [];
        const seen = new Set();
        for (const item of next) {
          const key = `${clean(String(item?.title || '')).toLowerCase()}|${clean(String(item?.link || '')).toLowerCase()}`;
          if (!key || seen.has(key)) continue;
          seen.add(key);
          deduped.push(item);
        }
        return deduped;
      });
      setArticleProgress(100);
      if (preferNew && nextResult?.isRepeat) {
        setArticleStatus('No different result was available for this prompt right now; showing the best available article again. Try changing the focus prompt.');
      } else {
        setArticleStatus('Scholarly article found. Review/edit the write-up and copy the result below.');
      }
    } catch (err) {
      setArticleResult(null);
      setArticleProgress(0);
      setArticleError(err.message || 'Article search failed.');
      setArticleStatus('Article search failed. Try a different focus prompt.');
    } finally {
      setArticleLoading(false);
    }
  };

  const renderConceptMapField = ([key, label]) => {
    const reviewReason = getReviewReasonForValue(fields[key]);
    const exportLimit = CONCEPT_MAP_FIELD_LIMITS[key];
    const fieldLength = clean(String(fields[key] || '')).length;
    const priorityPrompt = PRIORITY_FIELD_PROMPTS[key] || '';
    const overExportLimit = Boolean(exportLimit && fieldLength > exportLimit);
    const highlightReview = Boolean(reviewReason && ['Verify', 'AI generated', 'Missing'].includes(reviewReason)) || overExportLimit;
    const priorityDiagnosisOptions = [
      'Deficient fluid volume',
      'Ineffective tissue perfusion',
      'Impaired urinary elimination',
      'Risk for maternal injury',
      'Risk for infection',
      'Risk for birth injury',
      'Acute pain',
      'Anxiety',
      'Deficient knowledge',
      'Impaired fetal gas exchange',
      'Impaired physical mobility',
      'Activity intolerance',
      'Risk for falls',
      'Risk for bleeding',
      'Fatigue',
      'Decreased cardiac output',
      'Impaired skin integrity',
    ];
    const selectOptionsByKey = {
      orientation: ['A&O x4', 'A&O x3', 'A&O x2', 'A&O x1', 'Alert, confused', 'Lethargic', 'Unable to assess'],
      skinTurgor: ['Normal', 'Decreased', 'Poor', 'Unable to assess'],
      breathSounds: ['Clear', 'Diminished', 'Crackles', 'Wheezes', 'Rhonchi', 'Unable to assess'],
      peripheralPulses: ['Present', '+2', 'Weak', 'Diminished', 'Absent', 'Unable to assess'],
      edema: ['None noted', 'Trace', '+1', '+2', '+3', '+4', 'Need to assess'],
      bowelSounds: ['Present', 'Active', 'Hypoactive', 'Hyperactive', 'Absent', 'Unable to assess'],
      speech: ['Clear', 'Slurred', 'Delayed', 'Nonverbal', 'Unable to assess'],
      weakness: ['None noted', 'Generalized weakness', 'Left-sided weakness', 'Right-sided weakness', 'Bilateral weakness'],
      surgicalIncision: ['None', 'Clean/dry/intact', 'Dressing present', 'Drainage noted', 'Need to assess'],
      ivLocation: ['None', 'Left arm', 'Right arm', 'Left hand', 'Right hand', 'Forearm', 'Need to assess'],
      nd1Diagnosis: priorityDiagnosisOptions,
      nd2Diagnosis: priorityDiagnosisOptions,
      nd3Diagnosis: priorityDiagnosisOptions,
    };
    const useInput = ['studentName', 'date', 'clientName', 'age', 'sex', 'ht', 'wt', 'diagnosis', 'temp', 'pulse', 'resp', 'bp', 'allergies', 'immunizations', 'currentMedDate'].includes(key);
    const quickOptionsByKey = {
      allergies: ['None'],
      immunizations: ['Up to date', 'None'],
      assistanceAdls: ['Independent', 'Needs assistance'],
      surgicalHistory: ['None'],
      sleepPattern: ['Normal sleep pattern', 'Irregular sleep pattern'],
      nutritionalStatus: ['Nourished', 'Malnourished'],
    };

    const compactField = ['date', 'age', 'sex', 'ht', 'wt', 'temp', 'pulse', 'resp', 'bp', 'pain', 'currentMedDate'].includes(key);

    return (
      <div className={`field field-${key} ${compactField ? 'compact-field' : ''} ${highlightReview ? 'needs-review' : ''}`} key={key}>
        <label className="field-label-row">
          <span>{label}</span>
          <span className="field-badges">
            {reviewReason && <span className={`badge ${highlightReview ? 'danger' : 'review'}`}>{reviewReason}</span>}
            {overExportLimit && <span className="badge danger">PDF limit {exportLimit}</span>}
          </span>
        </label>
        {key === 'pain' ? (
          <select value={fields[key] || ''} onChange={(e) => updateField(key, e.target.value)}>
            <option value="">Leave blank</option>
            <option value="None">None</option>
            {Array.from({ length: 11 }, (_, score) => (
              <option key={score} value={`${score}/10`}>{score}/10</option>
            ))}
            <option value="3-4/10">3-4/10</option>
          </select>
        ) : selectOptionsByKey[key] ? (
          <select value={fields[key] || ''} onChange={(e) => updateField(key, e.target.value)}>
            <option value="">Leave blank</option>
            {fields[key] && !selectOptionsByKey[key].includes(fields[key]) && (
              <option value={fields[key]}>{fields[key]}</option>
            )}
            {selectOptionsByKey[key].map((option) => (
              <option key={option} value={option}>{option}</option>
            ))}
          </select>
        ) : useInput ? (
          <input type="text" value={fields[key] || ''} placeholder={priorityPrompt} onChange={(e) => updateField(key, e.target.value)} />
        ) : (
          <textarea value={fields[key] || ''} placeholder={priorityPrompt} onChange={(e) => updateField(key, e.target.value)} />
        )}
        {exportLimit && (
          <div className={`field-hint ${overExportLimit ? 'danger' : ''}`}>
            PDF target: {fieldLength}/{exportLimit} chars
          </div>
        )}
        {quickOptionsByKey[key] && (
          <div className="btn-row" style={{ marginTop: 8 }}>
            {quickOptionsByKey[key].map((option) => (
              <button key={option} type="button" className="btn" onClick={() => updateField(key, option)}>{option}</button>
            ))}
          </div>
        )}
      </div>
    );
  };

  const typhonGuide = (
    <details className="guide-panel guide-hero" open>
      <summary>
        <span><FileText size={17} /> Start Here: Typhon Case Log Guide</span>
        <ChevronDown size={18} />
      </summary>
      <div className="guide-body">
        <div className="guide-copy">
          <strong>Use the exported Typhon case log as your source.</strong>
          <p>The builder fills the concept-map draft from that log, then you review the marked fields and export the finished PDF. Think of this as a guided first pass, not a locked form.</p>
        </div>
        <div className="guide-screens" aria-label="Onboarding walkthrough">
          <div className="guide-screen">
            <div className="mini-window">
              <div className="mini-bar"><span></span><span></span><span></span></div>
              <div className="mini-upload">
                <Upload size={18} />
                <strong>Upload Case PDF</strong>
                <small>Typhon case log</small>
              </div>
            </div>
            <strong>1. Add the log</strong>
            <p>Upload the case PDF, or paste the case text if the PDF is hard to read.</p>
          </div>
          <div className="guide-screen">
            <div className="mini-window">
              <div className="mini-bar"><span></span><span></span><span></span></div>
              <div className="mini-actions">
                <span className="mini-button dark">Upload</span>
                <span className="mini-button blue">Auto Fill</span>
              </div>
              <div className="mini-progress"><i></i><i></i><i></i></div>
            </div>
            <strong>2. Run Auto Fill</strong>
            <p>The app places the best available values into the concept-map fields.</p>
          </div>
          <div className="guide-screen">
            <div className="mini-window">
              <div className="mini-bar"><span></span><span></span><span></span></div>
              <div className="mini-field"><b>Allergies</b><em className="mini-flag gold">Verify</em></div>
              <div className="mini-field"><b>Edema</b><em className="mini-flag rose">Missing</em></div>
              <div className="mini-field"><b>Medication</b><em className="mini-flag plum">AI generated</em></div>
            </div>
            <strong>3. Review flags</strong>
            <p>Marked values need your eyes before the final export.</p>
          </div>
          <div className="guide-screen">
            <div className="mini-window">
              <div className="mini-bar"><span></span><span></span><span></span></div>
              <div className="mini-map">
                <span></span><span></span><span></span><span></span>
                <strong>PDF</strong>
              </div>
            </div>
            <strong>4. Export the map</strong>
            <p>Download the formatted concept map PDF after the fields look right.</p>
          </div>
        </div>
        <div className="guide-flags">
          <span className="guide-flag missing">Missing: no clear source value was found.</span>
          <span className="guide-flag verify">Verify: check this against the chart or case log.</span>
          <span className="guide-flag generated">AI generated: review before submitting.</span>
          <span className="guide-flag limit">PDF limit: shorten text so it fits the box.</span>
        </div>
      </div>
    </details>
  );

  return (
    <div className={`app-shell view-${viewMode}`}>
      <div className="container">
        <div className="topbar">
          <div className="title">
            <h1>Concept Map Builder</h1>
            <p>This concept builder helps create well-formatted, accurate concept maps with a smoother workflow. Output depends on the case log you provide: a clearer case log gives better results, and manual entry is available anywhere you need to review, correct, or complete details.</p>
          </div>
          <div className="view-toggle" aria-label="View mode">
            <button
              type="button"
              className={viewMode === 'split' ? 'active' : ''}
              onClick={() => setViewMode('split')}
            >
              <Columns2 size={16} />
              Split
            </button>
            <button
              type="button"
              className={viewMode === 'vertical' ? 'active' : ''}
              onClick={() => setViewMode('vertical')}
            >
              <Rows3 size={16} />
              Vertical
            </button>
          </div>
          <input ref={fileInputRef} type="file" accept="application/pdf" hidden onChange={(e) => handleFile(e.target.files?.[0])} />
          <input ref={simulationFileInputRef} type="file" accept="application/pdf" hidden onChange={(e) => handleSimulationFile(e.target.files?.[0])} />
        </div>

        {typhonGuide}

        <div className="workflow">
          <section className="card setup-card">
            <div className="card-header"><h2 className="card-title">Progress</h2></div>
            <div className="card-content">
              <div className="flow-panel">
                <div className="signal-list">
                  {workflowSignals.map((item) => (
                    <div className={`signal-item ${item.state}`} key={item.label}>
                      <span className="signal-dot">
                        {item.state === 'complete' ? <CheckCircle2 size={18} /> : <CircleDashed size={18} />}
                      </span>
                      <span>
                        <strong>{item.label}</strong>
                        <small>{item.detail}</small>
                      </span>
                    </div>
                  ))}
                </div>
                <div className="live-status">{status}</div>
              </div>
              <div>
                <div className="split-2">
                  <div className="field-inline">
                    <label>Student Name</label>
                    <input type="text" value={fields.studentName || ''} onChange={(e) => updateField('studentName', e.target.value)} />
                  </div>
                  <div className="field-inline">
                    <label>Clinical Faculty</label>
                    <select value={fields.facultyMeta || ''} onChange={(e) => updateField('facultyMeta', e.target.value)}>
                      <option value="">Select faculty</option>
                      {facultyOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                    </select>
                    <div className="inline-save-row" style={{ marginTop: 8 }}>
                      <input
                        type="text"
                        value={sessionFacultyName}
                        placeholder="Temporary instructor name"
                        onChange={(e) => setSessionFacultyName(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter') {
                            e.preventDefault();
                            applySessionFacultyName();
                          }
                        }}
                      />
                      <button type="button" className="btn" onClick={applySessionFacultyName}>Use</button>
                    </div>
                    <div className="field-hint">Custom instructor names are kept for this browser session only.</div>
                  </div>
                </div>
                {!metadataReady && (
                  <div className="warning">
                    <strong><AlertTriangle size={16} style={{verticalAlign:'text-bottom', marginRight:6}} />Setup missing</strong>
                    <div>Student name and instructor are used on the concept map header.</div>
                  </div>
                )}
              </div>
            </div>
          </section>

          <section className="card intake-card">
            <div className="card-header"><h2 className="card-title"><FileText size={18} style={{verticalAlign:'text-bottom', marginRight:8}} />Case Intake</h2></div>
            <div className="card-content">
              <div className="intake-actions">
                <button className="btn primary" onClick={() => fileInputRef.current?.click()} disabled={loading || aiLoading}><Upload size={16} />Upload Case PDF</button>
                <button className="btn primary-soft" onClick={() => applyParsedText(rawText)} disabled={!rawText.trim() || loading || aiLoading}><Wand2 size={16} />Auto Fill From Case PDF</button>
              </div>
              <details className="source-panel">
                <summary>
                  <span>{rawText.trim() ? 'Source text captured' : 'Paste case text manually'}</span>
                  <ChevronDown size={18} />
                </summary>
                <textarea className="large-text" value={rawText} onChange={(e) => setRawText(e.target.value)} placeholder="PDF text appears here. You can paste case-log text if upload is not available." />
              </details>
              {(loading || aiLoading || parseProgress > 0) && (
                <div className="progress-panel">
                  <strong>{parseProgressLabel || 'Progress'}</strong>
                  <progress value={parseProgress} max={100} style={{ width: '100%', height: 12 }} />
                </div>
              )}
              {!!conceptMapReviewItems.length && sequenceStep >= 3 && (
                <div className="warning">
                  <strong><AlertTriangle size={16} style={{verticalAlign:'text-bottom', marginRight:6}} />Missing or AI-generated concept-map fields</strong>
                  <div className="badges">
                    {conceptMapReviewItems.slice(0, 18).map((item) => (
                      <span key={item.key} className="badge review">{item.label}: {item.reason}</span>
                    ))}
                    {conceptMapReviewItems.length > 18 && <span className="badge review">+{conceptMapReviewItems.length - 18} more</span>}
                  </div>
                </div>
              )}
            </div>
          </section>

          <div className="review-export-stack">
            <section className="card review-card">
              <div className="card-header"><h2 className="card-title">Concept Map Fields</h2></div>
              <div className="card-content">
                {MAIN_CONCEPT_MAP_FIELD_GROUPS.map((group) => (
                  <details className="field-group" key={group.title}>
                    <summary>
                      <h3>{group.title}</h3>
                      <Plus size={18} />
                    </summary>
                    <div className="field-group-grid">
                      {group.fields.map(renderConceptMapField)}
                    </div>
                  </details>
                ))}
                <details className="field-group">
                  <summary>
                    <span>Medications</span>
                    <Plus size={18} />
                  </summary>
                  <div className="field">
                    <label className="field-label-row">
                      <span>Medication Review</span>
                    {!!medicationReviewItems.length && <span className="badge review">Review</span>}
                    </label>
                    <div className="med-grid">
                      {medications.length ? medications.map((med, idx) => (
                        <div className="med-card" key={idx}>
                          <h4>Medication {idx + 1}</h4>
                          <div className="field-inline"><label>Name / Class</label><input type="text" value={med.nameClass || ''} onChange={(e) => updateMedication(idx, 'nameClass', e.target.value)} /></div>
                          <div className="field-inline"><label>Dose / Route</label><input type="text" value={med.doseRoute || ''} onChange={(e) => updateMedication(idx, 'doseRoute', e.target.value)} /></div>
                          <div className="field-inline"><label>Indication</label><textarea value={med.why || ''} onChange={(e) => updateMedication(idx, 'why', e.target.value)} /></div>
                        </div>
                      )) : (
                        <div className="status">No medication entries parsed yet.</div>
                      )}
                    </div>
                    <button className="btn" onClick={addMedicationRow} disabled={!canAddMedicationRow}>Add Medication</button>
                  </div>
                </details>
              </div>
            </section>

            <section className="card priority-card">
              <div className="card-header"><h2 className="card-title">Nursing Diagnosis</h2></div>
              <div className="card-content">
                {PRIORITY_NURSING_FIELD_GROUPS.map((group) => (
                  <details className="field-group" key={group.title}>
                    <summary>
                      <h3>{group.title}</h3>
                      <Plus size={18} />
                    </summary>
                    <div className="field-group-grid">
                      {group.fields.map(renderConceptMapField)}
                    </div>
                  </details>
                ))}
                <details className="ai-support-panel">
                  <summary>
                    <span><Wand2 size={16} />Optional AI-generated support</span>
                    <Plus size={18} />
                  </summary>
                  <div className="ai-support-body">
                    <strong>This is a testing aid only.</strong>
                    <p>It can draft nursing-priority text from the case log, but it is not meant to replace your own clinical judgment, instructor requirements, or manual review.</p>
                    <button className="btn" onClick={developPriorityNursingBeta} disabled={sequenceStep < 3 || loading || aiLoading}>
                      <Wand2 size={16} />Fill nursing priorities for review
                    </button>
                  </div>
                </details>
              </div>
            </section>

            <section className="card export-card">
              <div className="card-header"><h2 className="card-title">Export</h2></div>
              <div className="card-content">
                <div className="status export-status">
                  <strong>Export Status</strong>
                  <div>{outputStatus}</div>
                  <progress value={outputProgress} max={100} style={{ width: '100%', height: 12 }} />
                  {outputLastRun && <div>Last export: {outputLastRun}</div>}
                  {outputError && <div style={{ color: '#991b1b' }}>Error: {outputError}</div>}
                </div>
                <div className="export-choice">
                  <div>
                    <strong>Normal class map</strong>
                    <span>Downloads the concept map with the nursing-priority boxes blank.</span>
                  </div>
                  <button className="btn primary" onClick={() => downloadConceptMapPdfOutput({ includePriorityNursingSections: false })} disabled={sequenceStep < 3 || loading || !templateReady}>
                    <Download size={16} />Download PDF
                  </button>
                </div>
                <div className="export-choice">
                  <div>
                    <strong>Include nursing priorities</strong>
                    <span>Uses the separate beta/manual nursing-priority fields above.</span>
                  </div>
                  <button className="btn" onClick={() => downloadConceptMapPdfOutput({ includePriorityNursingSections: true })} disabled={sequenceStep < 3 || loading || !templateReady}>
                    <Download size={16} />Download With Priorities
                  </button>
                </div>
              </div>
            </section>

            <section className="card simulation-beta-card">
              <div className="card-header"><h2 className="card-title"><Wand2 size={18} style={{verticalAlign:'text-bottom', marginRight:8}} />Nontraditional Notes Beta</h2></div>
              <div className="card-content">
                <div className="status export-status">
                  <strong>Simulation Beta Status</strong>
                  <div>{simulationStatus}</div>
                </div>
                <div className="beta-source-actions beta-action-grid">
                  <button className="btn primary" onClick={() => simulationFileInputRef.current?.click()} disabled={loading || aiLoading}>
                    <Upload size={16} />Load Simulation PDF
                  </button>
                  <button className="btn test-action" onClick={clearSimulationBeta} disabled={loading || aiLoading}>
                    Clear
                  </button>
                  <button className="btn primary-soft" onClick={() => applySimulationNotesText(simulationNotesText, simulationFileName)} disabled={!simulationNotesText.trim() || loading || aiLoading}>
                    <Wand2 size={16} />Parse Simulation PDF
                  </button>
                  <button className="btn" onClick={() => downloadConceptMapPdfOutput({ includePriorityNursingSections: true })} disabled={!simulationParsed || loading || !templateReady}>
                    <Download size={16} />Download Beta PDF
                  </button>
                </div>
                <div className="simulation-diagnosis-editor">
                  <h3>Nursing Diagnoses For This Simulation</h3>
                  {PRIORITY_NURSING_FIELD_GROUPS.map((group) => (
                    <details className="field-group" key={`simulation-${group.title}`} open>
                      <summary>
                        <h3>{group.title}</h3>
                        <Plus size={18} />
                      </summary>
                      <div className="field-group-grid">
                        {group.fields.map(renderConceptMapField)}
                      </div>
                    </details>
                  ))}
                </div>
                <details className="source-panel">
                  <summary>
                    <span>{simulationNotesText.trim() ? 'Simulation source loaded' : 'Paste simulation notes manually'}</span>
                    <ChevronDown size={18} />
                  </summary>
                  <textarea
                    className="large-text"
                    value={simulationNotesText}
                    onChange={(e) => {
                      setSimulationNotesText(e.target.value);
                      setSimulationFileName('');
                      setSimulationParsed(false);
                      setSimulationStatus('Simulation notes changed. Parse again before downloading the beta PDF.');
                    }}
                    placeholder="Simulation PDF text appears here. You can paste vSim notes or narrative simulation notes here."
                  />
                </details>
              </div>
            </section>
          </div>
        </div>
      </div>
    </div>
  );
}
