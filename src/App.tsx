/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  UserPlus,
  FileSpreadsheet,
  RotateCcw,
  Trash2,
  Printer,
  CheckCircle2,
  ExternalLink,
  AlertTriangle,
  Award,
  Save,
  Edit3,
  Check,
  Plus,
  X,
  Sliders,
  Search,
  Upload,
  Download,
  FileText,
  ClipboardList,
} from 'lucide-react';

export function is1to5Class(cls: string, level?: '1-5' | '6-9'): boolean {
  if (cls && ['1', '2', '3', '4', '5'].includes(cls)) return true;
  if (cls && ['6', '7', '8', '9'].includes(cls)) return false;
  return level === '1-5';
}

export const CLASSES_1_5 = ['1', '2', '3', '4', '5'];
export const CLASSES_6_9 = ['6', '7', '8', '9'];

export const SUBJECTS_6_9 = [
  'ಪ್ರಥಮ ಭಾಷೆ',
  'ದ್ವಿತೀಯ ಭಾಷೆ',
  'ತೃತೀಯ ಭಾಷೆ',
  'ಗಣಿತ',
  'ವಿಜ್ಞಾನ',
  'ಸಮಾಜ ವಿಜ್ಞಾನ',
  'ದೈಹಿಕ ಶಿಕ್ಷಣ / ಆರೋಗ್ಯ ಶಿಕ್ಷಣ',
];

export const SUBJECTS_1_5 = [
  'ಪ್ರಥಮ ಭಾಷೆ',
  'ದ್ವಿತೀಯ ಭಾಷೆ',
  'ತೃತೀಯ ಭಾಷೆ',
  'ಗಣಿತ',
  'ಪರಿಸರ ವಿಜ್ಞಾನ',
];

export function getSubjectsForClass(cls: string, level?: '1-5' | '6-9'): string[] {
  return is1to5Class(cls, level) ? SUBJECTS_1_5 : SUBJECTS_6_9;
}

export const CO_SKILLS_6_9 = [
  'ಭಾ. ಸಾ ಕೌಶಲಗಳು',
  'ಸಂಘಟನಾ ಕೌಶಲಗಳು',
  'ವೈಜ್ಞಾನಿಕ ಕೌಶಲಗಳು',
  'ಲಲಿತ ಕಲೆ',
  'ಸೃಜನಶೀಲತೆ',
  'ಮನೋಧಾರಣೆ',
  'ಮೌಲ್ಯಗಳು',
];

export const CO_SKILLS_1_5 = [
  'ದೈಹಿಕ ಮತ್ತು ಆರೋಗ್ಯ ಶಿಕ್ಷಣ',
  'ಕಲೆ ಮತ್ತು ಕರಕುಶಲತೆ',
  'ಸಂಗೀತ ಮತ್ತು ನೃತ್ಯ',
  'ಸಾ.ಉ.ಉ.ಕಾ (S.U.P.W)',
  'ಮೌಲ್ಯಗಳು',
];

export function getCoSkillsForClass(cls: string, level?: '1-5' | '6-9'): string[] {
  return is1to5Class(cls, level) ? CO_SKILLS_1_5 : CO_SKILLS_6_9;
}

export const ASSESSMENT_PARTS_6_9 = [
  { name: 'ರೂ-1', pct: '10%' },
  { name: 'ರೂ-2', pct: '10%' },
  { name: 'ಸಂ-1', pct: '30%' },
  { name: 'ಒಟ್ಟು', pct: '50%' },
  { name: 'ರೂ-3', pct: '10%' },
  { name: 'ರೂ-4', pct: '10%' },
  { name: 'ಸಂ-2', pct: '30%' },
  { name: 'ಒಟ್ಟು', pct: '50%' },
  { name: 'ಒಟ್ಟು', pct: '100%' },
];

export const ASSESSMENT_PARTS_1_5 = [
  { name: 'ರೂ-1', pct: '15%' },
  { name: 'ರೂ-2', pct: '15%' },
  { name: 'ಸಂ-1', pct: '20%' },
  { name: 'ಒಟ್ಟು', pct: '50%' },
  { name: 'ರೂ-3', pct: '15%' },
  { name: 'ರೂ-4', pct: '15%' },
  { name: 'ಸಂ-2', pct: '20%' },
  { name: 'ಒಟ್ಟು', pct: '50%' },
  { name: 'ಅ.ಒಟ್ಟು', pct: '100%' },
];

export function getAssessmentPartsForClass(cls: string, level?: '1-5' | '6-9') {
  return is1to5Class(cls, level) ? ASSESSMENT_PARTS_1_5 : ASSESSMENT_PARTS_6_9;
}

export function getMaxMarksConfig(cls: string, level?: '1-5' | '6-9') {
  if (is1to5Class(cls, level)) {
    return {
      maxPerSub: [15, 15, 20, 50, 15, 15, 20, 50, 100],
      maxTotal: [75, 75, 100, 250, 75, 75, 100, 250, 500],
      maxWeightPct: [15, 15, 20, 50, 15, 15, 20, 50, 100],
      maxGrand: 500,
    };
  }
  return {
    maxPerSub: [10, 10, 30, 50, 10, 10, 30, 50, 100],
    maxTotal: [70, 70, 210, 350, 70, 70, 210, 350, 700],
    maxWeightPct: [10, 10, 30, 50, 10, 10, 30, 50, 100],
    maxGrand: 700,
  };
}

const EDITABLE_ASSESSMENT_INDICES = [0, 1, 2, 4, 5, 6];
const CALCULATED_ASSESSMENT_INDICES = [3, 7, 8];

function calcWeightPercent(val: number, idx: number, maxTotalList: number[], maxWeightPctList: number[]): number {
  const maxT = maxTotalList[idx];
  if (!maxT || maxT === 0) return 0;
  return Math.round((val / maxT) * maxWeightPctList[idx]);
}

function renderSubjectHeader(
  sub: string,
  isPrint = false,
  languageNames?: { lang1: string; lang2: string; lang3: string },
  onLanguageChange?: (langKey: 'lang1' | 'lang2' | 'lang3', value: string) => void
) {
  if (sub === 'ಸಮಾಜ ವಿಜ್ಞಾನ') {
    return (
      <div
        className="flex flex-col items-center justify-center text-center w-full leading-[1.15]"
        style={{ fontSize: isPrint ? '10.5px' : '13px' }}
      >
        <span>ಸಮಾಜ</span>
        <span>ವಿಜ್ಞಾನ</span>
      </div>
    );
  }
  if (sub.includes('ದೈಹಿಕ') && (sub.includes('ಆರೋಗ್ಯ') || sub.includes('ಶಿಕ್ಷಣ'))) {
    return (
      <div
        className="flex flex-col items-center justify-center text-center w-full leading-[1.05]"
        style={{ fontSize: isPrint ? '9px' : '11px' }}
        title="ದೈಹಿಕ ಶಿಕ್ಷಣ / ಆರೋಗ್ಯ ಶಿಕ್ಷಣ"
      >
        <span>ದೈಹಿಕ ಶಿಕ್ಷಣ /</span>
        <span>ಆರೋಗ್ಯ ಶಿಕ್ಷಣ</span>
      </div>
    );
  }
  if (sub === 'ಪ್ರಥಮ ಭಾಷೆ' || sub === 'ವಿಷಯ - 1' || sub === 'ಪ್ರಥಮ ವಿಷಯ') {
    const customLang = languageNames?.lang1 || '';
    if (isPrint) {
      return (
        <div className="flex flex-col items-center justify-center text-center w-full leading-normal pt-0.5 px-0.5 overflow-visible">
          <span className="whitespace-nowrap font-black text-black inline-block" style={{ fontSize: '10px', fontWeight: 900 }}>
            ಪ್ರಥಮ ಭಾಷೆ
          </span>
          {customLang ? (
            <span
              className="font-black text-black truncate max-w-full leading-tight mt-0.5 inline-block px-0.5"
              style={{ fontSize: '9px', fontWeight: 900 }}
            >
              ({customLang})
            </span>
          ) : null}
        </div>
      );
    }
    return (
      <div className="flex flex-col items-center justify-center text-center w-full leading-normal pt-1 pb-0.5 px-0.5 overflow-visible">
        <span
          className="whitespace-nowrap font-black text-[12.5px] leading-normal text-gray-950 inline-block"
          style={{ fontWeight: 900 }}
        >
          ಪ್ರಥಮ ಭಾಷೆ
        </span>
        <input
          id="langInput-lang1"
          type="text"
          className={`w-[98px] text-center border-0 border-b-2 border-solid ${customLang.trim() ? 'border-blue-600 bg-white/90' : 'border-amber-500 bg-amber-50/90 ring-1 ring-amber-400'} hover:bg-white focus:bg-white text-[11px] font-black text-gray-950 rounded-none px-1 py-0 h-[19px] mt-0.5 outline-none transition-colors placeholder:text-gray-600 placeholder:font-black placeholder:text-[10px]`}
          placeholder="ವಿಷಯ *"
          value={customLang}
          onChange={(e) => onLanguageChange?.('lang1', e.target.value)}
          title="ಪ್ರಥಮ ಭಾಷೆಯ ವಿಷಯ ಹೆಸರು (ಅಗತ್ಯ)"
          aria-label="ಪ್ರಥಮ ಭಾಷೆ"
        />
      </div>
    );
  }
  if (sub === 'ದ್ವಿತೀಯ ಭಾಷೆ' || sub === 'ವಿಷಯ - 2' || sub === 'ದ್ವಿತೀಯ ವಿಷಯ') {
    const customLang = languageNames?.lang2 || '';
    if (isPrint) {
      return (
        <div className="flex flex-col items-center justify-center text-center w-full leading-normal pt-0.5 px-0.5 overflow-visible">
          <span className="whitespace-nowrap font-black text-black inline-block" style={{ fontSize: '10px', fontWeight: 900 }}>
            ದ್ವಿತೀಯ ಭಾಷೆ
          </span>
          {customLang ? (
            <span
              className="font-black text-black truncate max-w-full leading-tight mt-0.5 inline-block px-0.5"
              style={{ fontSize: '9px', fontWeight: 900 }}
            >
              ({customLang})
            </span>
          ) : null}
        </div>
      );
    }
    return (
      <div className="flex flex-col items-center justify-center text-center w-full leading-normal pt-1 pb-0.5 px-0.5 overflow-visible">
        <span
          className="whitespace-nowrap font-black text-[12.5px] leading-normal text-gray-950 inline-block"
          style={{ fontWeight: 900 }}
        >
          ದ್ವಿತೀಯ ಭಾಷೆ
        </span>
        <input
          id="langInput-lang2"
          type="text"
          className={`w-[98px] text-center border-0 border-b-2 border-solid ${customLang.trim() ? 'border-blue-600 bg-white/90' : 'border-amber-500 bg-amber-50/90 ring-1 ring-amber-400'} hover:bg-white focus:bg-white text-[11px] font-black text-gray-950 rounded-none px-1 py-0 h-[19px] mt-0.5 outline-none transition-colors placeholder:text-gray-600 placeholder:font-black placeholder:text-[10px]`}
          placeholder="ವಿಷಯ *"
          value={customLang}
          onChange={(e) => onLanguageChange?.('lang2', e.target.value)}
          title="ದ್ವಿತೀಯ ಭಾಷೆಯ ವಿಷಯ ಹೆಸರು (ಅಗತ್ಯ)"
          aria-label="ದ್ವಿತೀಯ ಭಾಷೆ"
        />
      </div>
    );
  }
  if (sub === 'ತೃತೀಯ ಭಾಷೆ' || sub === 'ವಿಷಯ - 3' || sub === 'ತೃತೀಯ ವಿಷಯ') {
    const customLang = languageNames?.lang3 || '';
    if (isPrint) {
      return (
        <div className="flex flex-col items-center justify-center text-center w-full leading-normal pt-0.5 px-0.5 overflow-visible">
          <span className="whitespace-nowrap font-black text-black inline-block" style={{ fontSize: '10px', fontWeight: 900 }}>
            ತೃತೀಯ ಭಾಷೆ
          </span>
          {customLang ? (
            <span
              className="font-black text-black truncate max-w-full leading-tight mt-0.5 inline-block px-0.5"
              style={{ fontSize: '9px', fontWeight: 900 }}
            >
              ({customLang})
            </span>
          ) : null}
        </div>
      );
    }
    return (
      <div className="flex flex-col items-center justify-center text-center w-full leading-normal pt-1 pb-0.5 px-0.5 overflow-visible">
        <span
          className="whitespace-nowrap font-black text-[12.5px] leading-normal text-gray-950 inline-block"
          style={{ fontWeight: 900 }}
        >
          ತೃತೀಯ ಭಾಷೆ
        </span>
        <input
          id="langInput-lang3"
          type="text"
          className={`w-[98px] text-center border-0 border-b-2 border-solid ${customLang.trim() ? 'border-blue-600 bg-white/90' : 'border-amber-500 bg-amber-50/90 ring-1 ring-amber-400'} hover:bg-white focus:bg-white text-[11px] font-black text-gray-950 rounded-none px-1 py-0 h-[19px] mt-0.5 outline-none transition-colors placeholder:text-gray-600 placeholder:font-black placeholder:text-[10px]`}
          placeholder="ವಿಷಯ *"
          value={customLang}
          onChange={(e) => onLanguageChange?.('lang3', e.target.value)}
          title="ತೃತೀಯ ಭಾಷೆಯ ವಿಷಯ ಹೆಸರು (ಅಗತ್ಯ)"
          aria-label="ತೃತೀಯ ಭಾಷೆ"
        />
      </div>
    );
  }
  return (
    <div
      className="flex items-center justify-center text-center w-full leading-[1.15]"
      style={{ fontSize: isPrint ? '11px' : '13px' }}
    >
      <span>{sub}</span>
    </div>
  );
}

export function renderCoSkillHeader(co: string, isPrint = false) {
  let line1 = co;
  let line2 = '';

  // Class 1-5 co-skills split into 2 lines
  if (co === 'ದೈಹಿಕ ಮತ್ತು ಆರೋಗ್ಯ ಶಿಕ್ಷಣ') {
    line1 = 'ದೈಹಿಕ ಮತ್ತು';
    line2 = 'ಆರೋಗ್ಯ ಶಿಕ್ಷಣ';
  } else if (co === 'ಕಲೆ ಮತ್ತು ಕರಕುಶಲತೆ' || co === 'ಕಲೆ ಮತ್ತು ಕರಕುಶಲ') {
    line1 = 'ಕಲೆ ಮತ್ತು';
    line2 = 'ಕರಕುಶಲ';
  } else if (co === 'ಸಂಗೀತ ಮತ್ತು ನೃತ್ಯ') {
    line1 = 'ಸಂಗೀತ ಮತ್ತು';
    line2 = 'ನೃತ್ಯ';
  } else if (co === 'ಸಾ.ಉ.ಉ.ಕಾ (S.U.P.W)') {
    line1 = 'ಸಾ.ಉ.ಉ.ಕಾ';
    line2 = '(S.U.P.W)';
  } else if (
    co === 'ಭಾ. ಸಾ ಕೌಶಲಗಳು' ||
    co === 'ಸಂಘಟನಾ ಕೌಶಲಗಳು' ||
    co === 'ವೈಜ್ಞಾನಿಕ ಕೌಶಲಗಳು' ||
    co === 'ಲಲಿತ ಕಲೆ' ||
    co === 'ಸೃಜನಶೀಲತೆ' ||
    co === 'ಮನೋಧಾರಣೆ' ||
    co === 'ಮೌಲ್ಯಗಳು'
  ) {
    // Class 6-9 co-skills stay in 1 line as previously
    line1 = co;
    line2 = '';
  }

  const vStyle: React.CSSProperties = {
    writingMode: 'vertical-rl',
    transform: 'rotate(180deg)',
    transformOrigin: 'center center',
    whiteSpace: 'nowrap',
    fontSize: isPrint ? '9.5px' : '11.5px',
    fontWeight: 700,
    lineHeight: 1.1,
    margin: '0',
    textAlign: 'center',
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 'auto',
    flex: '0 0 auto',
  };

  if (line2) {
    return (
      <div
        className="co-skill-flex flex flex-row items-center justify-center gap-1 w-full h-full mx-auto text-center"
        style={{
          display: 'flex',
          flexDirection: 'row',
          alignItems: 'center',
          justifyContent: 'center',
          textAlign: 'center',
          width: '100%',
          height: '100%',
          margin: '0 auto',
          gap: '2px',
        }}
      >
        <div style={vStyle} className="co-skill-line">{line1}</div>
        <div style={vStyle} className="co-skill-line">{line2}</div>
      </div>
    );
  }

  return (
    <div
      className="co-skill-flex flex flex-row items-center justify-center w-full h-full mx-auto text-center"
      style={{
        display: 'flex',
        flexDirection: 'row',
        alignItems: 'center',
        justifyContent: 'center',
        textAlign: 'center',
        width: '100%',
        height: '100%',
        margin: '0 auto',
      }}
    >
      <div style={vStyle} className="co-skill-line">{line1}</div>
    </div>
  );
}

const ASSESSMENT_LABELS = [
  'ರೂ-1 10%',
  'ರೂ-2 10%',
  'ಸಂ-1 30%',
  'ಒಟ್ಟು 50%',
  'ರೂ-3 10%',
  'ರೂ-4 10%',
  'ಸಂ-2 30%',
  'ಒಟ್ಟು 50%',
  'ಒಟ್ಟು 100%',
];

const ASSESSMENT_PARTS = [
  { name: 'ರೂ-1', pct: '10%' },
  { name: 'ರೂ-2', pct: '10%' },
  { name: 'ಸಂ-1', pct: '30%' },
  { name: 'ಒಟ್ಟು', pct: '50%' },
  { name: 'ರೂ-3', pct: '10%' },
  { name: 'ರೂ-4', pct: '10%' },
  { name: 'ಸಂ-2', pct: '30%' },
  { name: 'ಒಟ್ಟು', pct: '50%' },
  { name: 'ಒಟ್ಟು', pct: '100%' },
];

const CO_SKILLS = [
  'ಭಾ. ಸಾ ಕೌಶಲಗಳು',
  'ಸಂಘಟನಾ ಕೌಶಲಗಳು',
  'ವೈಜ್ಞಾನಿಕ ಕೌಶಲಗಳು',
  'ಲಲಿತ ಕಲೆ',
  'ಸೃಜನಶೀಲತೆ',
  'ಮನೋಧಾರಣೆ',
  'ಮೌಲ್ಯಗಳು',
];

const MAX_PER_SUB = [10, 10, 30, 50, 10, 10, 30, 50, 100];
const MAX_TOTAL = [70, 70, 210, 350, 70, 70, 210, 350, 700];
const MAX_WEIGHT_PCT = [10, 10, 30, 50, 10, 10, 30, 50, 100];

function calcPercentExact(val: number, max: number): number {
  return (val / max) * 100;
}

export interface GradeItemRule {
  grade: string;
  min10: number;
  max10: number;
  text10: string;
  min15: number;
  max15: number;
  text15: string;
  min20: number;
  max20: number;
  text20: string;
  min30: number;
  max30: number;
  text30: string;
  min50: number;
  max50: number;
  text50: string;
  min100: number;
  max100: number;
  text100: string;
}

export const DEFAULT_GRADE_RULES: GradeItemRule[] = [
  {
    grade: 'A+',
    min10: 9,
    max10: 10,
    text10: '9 - 10',
    min15: 13.5,
    max15: 15,
    text15: '13.5 - 15',
    min20: 18,
    max20: 20,
    text20: '18 - 20',
    min30: 27,
    max30: 30,
    text30: '27 - 30',
    min50: 45,
    max50: 50,
    text50: '45 - 50',
    min100: 90,
    max100: 100,
    text100: '90 - 100',
  },
  {
    grade: 'A',
    min10: 7,
    max10: 8.9,
    text10: '7 - 8.9',
    min15: 10.5,
    max15: 13.4,
    text15: '10.5 - 13.4',
    min20: 14,
    max20: 17.9,
    text20: '14 - 17.9',
    min30: 22,
    max30: 26.9,
    text30: '22 - 26.9',
    min50: 35,
    max50: 44.9,
    text50: '35 - 44.9',
    min100: 70,
    max100: 89.9,
    text100: '70 - 89.9',
  },
  {
    grade: 'B+',
    min10: 5,
    max10: 6.9,
    text10: '5 - 6.9',
    min15: 7.5,
    max15: 10.4,
    text15: '7.5 - 10.4',
    min20: 10,
    max20: 13.9,
    text20: '10 - 13.9',
    min30: 16,
    max30: 21.9,
    text30: '16 - 21.9',
    min50: 25,
    max50: 34.9,
    text50: '25 - 34.9',
    min100: 50,
    max100: 69.9,
    text100: '50 - 69.9',
  },
  {
    grade: 'B',
    min10: 3,
    max10: 4.9,
    text10: '3 - 4.9',
    min15: 4.5,
    max15: 7.4,
    text15: '4.5 - 7.4',
    min20: 6,
    max20: 9.9,
    text20: '6 - 9.9',
    min30: 10,
    max30: 15.9,
    text30: '10 - 15.9',
    min50: 15,
    max50: 24.9,
    text50: '15 - 24.9',
    min100: 30,
    max100: 49.9,
    text100: '30 - 49.9',
  },
  {
    grade: 'C',
    min10: 0,
    max10: 2.9,
    text10: '0 - 2.9',
    min15: 0,
    max15: 4.4,
    text15: '0 - 4.4',
    min20: 0,
    max20: 5.9,
    text20: '0 - 5.9',
    min30: 0,
    max30: 9.9,
    text30: '0 - 9.9',
    min50: 0,
    max50: 14.9,
    text50: '0 - 14.9',
    min100: 0,
    max100: 29.9,
    text100: '0 - 29.9',
  },
];

export function parseMinFromText(text: string): number {
  if (!text) return 0;
  const trimmed = text.trim();
  if (
    trimmed.includes('ಕಡಿಮೆ') ||
    trimmed.includes('<') ||
    trimmed.toLowerCase().includes('below') ||
    trimmed.toLowerCase().includes('less')
  ) {
    return 0;
  }
  const match = trimmed.match(/(\d+(\.\d+)?)/);
  if (match) {
    const val = parseFloat(match[1]);
    return isNaN(val) ? 0 : val;
  }
  return 0;
}

export function parseMaxFromText(text: string): number {
  if (!text) return 0;
  const match = text.match(/[-–—]\s*(\d+(\.\d+)?)/);
  if (match) {
    const val = parseFloat(match[1]);
    return isNaN(val) ? 0 : val;
  }
  const matchAll = text.match(/(\d+(\.\d+)?)/g);
  if (matchAll && matchAll.length > 1) {
    return parseFloat(matchAll[matchAll.length - 1]) || 0;
  }
  return 0;
}

export function normalizeGradeRules(rawList: unknown): GradeItemRule[] {
  if (!Array.isArray(rawList) || rawList.length === 0) return DEFAULT_GRADE_RULES;
  return rawList.map((r, i) => {
    const isBottom = i === rawList.length - 1;
    const min100 = Number(r.min100) || 0;
    const max100 = r.max100 !== undefined ? Number(r.max100) : (isBottom ? 29.9 : parseMaxFromText(r.text100) || 100);

    const min10 = Number(r.min10) || 0;
    const max10 = r.max10 !== undefined ? Number(r.max10) : (isBottom ? 2.9 : parseMaxFromText(r.text10) || 10);

    const min15 = r.min15 !== undefined ? Number(r.min15) : Math.round(min100 * 0.15 * 10) / 10;
    const max15 = r.max15 !== undefined ? Number(r.max15) : (isBottom ? 4.4 : parseMaxFromText(r.text15) || Math.round(max100 * 0.15 * 10) / 10);

    const min20 = r.min20 !== undefined ? Number(r.min20) : Math.round(min100 * 0.20 * 10) / 10;
    const max20 = r.max20 !== undefined ? Number(r.max20) : (isBottom ? 5.9 : parseMaxFromText(r.text20) || Math.round(max100 * 0.20 * 10) / 10);

    const min30 = Number(r.min30) || 0;
    const max30 = r.max30 !== undefined ? Number(r.max30) : (isBottom ? 9.9 : parseMaxFromText(r.text30) || 30);

    const min50 = Number(r.min50) || 0;
    const max50 = r.max50 !== undefined ? Number(r.max50) : (isBottom ? 14.9 : parseMaxFromText(r.text50) || 50);

    return {
      grade: r.grade || (isBottom ? 'C' : 'A'),
      min10,
      max10,
      text10: `${min10} - ${max10}`,
      min15,
      max15,
      text15: `${min15} - ${max15}`,
      min20,
      max20,
      text20: `${min20} - ${max20}`,
      min30,
      max30,
      text30: `${min30} - ${max30}`,
      min50,
      max50,
      text50: `${min50} - ${max50}`,
      min100,
      max100,
      text100: `${min100} - ${max100}`,
    };
  });
}

export function getGradeBadgeClass(grade: string): string {
  const g = (grade || '').trim().toUpperCase();
  if (g.startsWith('A+')) return 'bg-emerald-600 text-white';
  if (g.startsWith('A')) return 'bg-blue-600 text-white';
  if (g.startsWith('B+')) return 'bg-amber-500 text-white';
  if (g.startsWith('B')) return 'bg-orange-500 text-white';
  if (g.startsWith('C')) return 'bg-rose-600 text-white';
  if (g.startsWith('D')) return 'bg-red-700 text-white';
  return 'bg-indigo-600 text-white';
}

export function getGradeCardBadgeClass(grade: string): string {
  const g = (grade || '').trim().toUpperCase();
  if (g.startsWith('A+')) return 'bg-emerald-100 text-emerald-800 border-emerald-300';
  if (g.startsWith('A')) return 'bg-blue-100 text-blue-800 border-blue-300';
  if (g.startsWith('B+')) return 'bg-amber-100 text-amber-800 border-amber-300';
  if (g.startsWith('B')) return 'bg-orange-100 text-orange-800 border-orange-300';
  if (g.startsWith('C')) return 'bg-rose-100 text-rose-800 border-rose-300';
  if (g.startsWith('D')) return 'bg-red-100 text-red-800 border-red-300';
  return 'bg-indigo-100 text-indigo-800 border-indigo-300';
}

function getGrade(pct: number, rules: GradeItemRule[] = DEFAULT_GRADE_RULES): string {
  if (!rules || rules.length === 0) return 'C';
  for (let i = 0; i < rules.length - 1; i++) {
    const rule = rules[i];
    if (pct >= rule.min100) return rule.grade;
  }
  const lastRule = rules[rules.length - 1];
  return lastRule ? lastRule.grade : 'C';
}

function toInt(v: unknown): number {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  if (v === '' || v === null || v === undefined || v === '-') return 0;
  const n = parseInt(String(v), 10);
  return isNaN(n) ? 0 : n;
}

function getGradeFromMarks(
  marks: number,
  maxMarks: number,
  rules: GradeItemRule[] = DEFAULT_GRADE_RULES
): string {
  if (!rules || rules.length === 0) return 'C';
  if (marks === 0) {
    const lastRule = rules[rules.length - 1];
    return lastRule ? lastRule.grade : 'C';
  }

  for (let i = 0; i < rules.length - 1; i++) {
    const rule = rules[i];
    if (maxMarks === 10) {
      if (marks >= rule.min10) return rule.grade;
    } else if (maxMarks === 15) {
      const minVal = rule.min15 !== undefined ? rule.min15 : (rule.min100 * 0.15);
      if (marks >= minVal) return rule.grade;
    } else if (maxMarks === 20) {
      const minVal = rule.min20 !== undefined ? rule.min20 : (rule.min100 * 0.20);
      if (marks >= minVal) return rule.grade;
    } else if (maxMarks === 30) {
      if (marks >= rule.min30) return rule.grade;
    } else if (maxMarks === 50) {
      if (marks >= rule.min50) return rule.grade;
    } else if (maxMarks === 70) {
      if (marks >= (rule.min10 * 7)) return rule.grade;
    } else if (maxMarks === 75) {
      const minVal = rule.min15 !== undefined ? (rule.min15 * 5) : (rule.min100 * 0.75);
      if (marks >= minVal) return rule.grade;
    } else if (maxMarks === 100) {
      if (marks >= rule.min100) return rule.grade;
    } else if (maxMarks === 210) {
      if (marks >= Math.round((rule.min30 / 30) * 210)) return rule.grade;
    } else if (maxMarks === 250) {
      if (marks >= Math.round((rule.min50 / 50) * 250)) return rule.grade;
    } else if (maxMarks === 350) {
      if (marks >= Math.round((rule.min50 / 50) * 350)) return rule.grade;
    } else if (maxMarks === 500) {
      if (marks >= Math.round((rule.min100 / 100) * 500)) return rule.grade;
    } else if (maxMarks === 700) {
      if (marks >= Math.round((rule.min100 / 100) * 700)) return rule.grade;
    } else {
      const pct = (marks / maxMarks) * 100;
      if (pct >= rule.min100) return rule.grade;
    }
  }

  const lastRule = rules[rules.length - 1];
  return lastRule ? lastRule.grade : 'C';
}

function getSubGrade(
  marks: number,
  maxMarks: number,
  rules: GradeItemRule[] = DEFAULT_GRADE_RULES
): string {
  if (marks === 0) return '';
  return getGradeFromMarks(marks, maxMarks, rules);
}

export interface SubScore {
  m: number | string;
  g: string;
}

export interface AssessmentItem {
  label: string;
  subs: SubScore[];
  total: {
    m: number | string;
    p: number | string;
    g: string;
  };
  is50?: boolean;
  is100?: boolean;
}

export interface Student {
  id: number;
  name: string;
  admission: string;
  gender?: string;
  dob: string;
  sats: string;
  pen: string;
  attendance: {
    s1Total: number;
    s2Total: number;
    s1Present: number;
    s2Present: number;
  };
  co1: string[];
  co2: string[];
  result: string;
  assessments: AssessmentItem[];
}

function formatDateDisplay(d?: string): string {
  if (!d) return '-';
  const parts = d.trim().split('-');
  if (parts.length === 3 && parts[0].length === 4) {
    return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  return d;
}

function recalculateAssessments(
  assessments: AssessmentItem[],
  rules: GradeItemRule[] = DEFAULT_GRADE_RULES,
  cls = '',
  level?: '1-5' | '6-9'
): AssessmentItem[] {
  const subjects = getSubjectsForClass(cls, level);
  const { maxPerSub, maxTotal, maxWeightPct } = getMaxMarksConfig(cls, level);

  const next: AssessmentItem[] = assessments.map((a) => ({
    ...a,
    subs: subjects.map((_, i) => ({
      m: a.subs[i]?.m ?? '',
      g: a.subs[i]?.g ?? '',
    })),
    total: { ...a.total },
  }));

  const getSubTotal = (idx: number) =>
    next[idx].subs.reduce((acc, curr) => acc + toInt(curr.m), 0);

  // 1. Calculate grades for editable individual assessments
  EDITABLE_ASSESSMENT_INDICES.forEach((idx) => {
    const hasAnyInput = next[idx].subs.some((s) => s.m !== '' && s.m !== '-' && s.m !== undefined);
    const sum = getSubTotal(idx);
    next[idx].total.m = hasAnyInput ? sum : '';
    next[idx].total.p = hasAnyInput ? calcWeightPercent(sum, idx, maxTotal, maxWeightPct) : '';
    next[idx].total.g = hasAnyInput ? getGradeFromMarks(sum, maxTotal[idx], rules) : '';

    next[idx].subs.forEach((sub, sIdx) => {
      if (sub.m === '' || sub.m === '-' || sub.m === undefined) {
        next[idx].subs[sIdx].g = '';
      } else {
        const mark = toInt(sub.m);
        if (mark === 0) {
          const lastRule = rules[rules.length - 1];
          next[idx].subs[sIdx].g = lastRule ? lastRule.grade : 'C';
        } else {
          next[idx].subs[sIdx].g = getGradeFromMarks(mark, maxPerSub[idx], rules);
        }
      }
    });
  });

  // 2. Sem 1 Total (Index 3: FA1 + FA2 + SA1)
  for (let i = 0; i < subjects.length; i++) {
    const hasInput = [0, 1, 2].some((c) => {
      const val = next[c].subs[i]?.m;
      return val !== '' && val !== '-' && val !== undefined;
    });
    if (hasInput) {
      const m = toInt(next[0].subs[i]?.m) + toInt(next[1].subs[i]?.m) + toInt(next[2].subs[i]?.m);
      next[3].subs[i].m = m;
      next[3].subs[i].g = getGradeFromMarks(m, maxPerSub[3], rules);
    } else {
      next[3].subs[i].m = '';
      next[3].subs[i].g = '';
    }
  }
  const hasSem1 = [0, 1, 2].some(
    (c) => getSubTotal(c) > 0 || next[c].subs.some((s) => s.m !== '' && s.m !== '-' && s.m !== undefined)
  );
  if (hasSem1) {
    const sem1Total = next[3].subs.reduce((acc, curr) => acc + toInt(curr.m), 0);
    next[3].total.m = sem1Total;
    next[3].total.p = calcWeightPercent(sem1Total, 3, maxTotal, maxWeightPct);
    next[3].total.g = getGradeFromMarks(sem1Total, maxTotal[3], rules);
  } else {
    next[3].total.m = '';
    next[3].total.p = '';
    next[3].total.g = '';
  }

  // 3. Sem 2 Total (Index 7: FA3 + FA4 + SA2)
  for (let i = 0; i < subjects.length; i++) {
    const hasInput = [4, 5, 6].some((c) => {
      const val = next[c].subs[i]?.m;
      return val !== '' && val !== '-' && val !== undefined;
    });
    if (hasInput) {
      const m = toInt(next[4].subs[i]?.m) + toInt(next[5].subs[i]?.m) + toInt(next[6].subs[i]?.m);
      next[7].subs[i].m = m;
      next[7].subs[i].g = getGradeFromMarks(m, maxPerSub[7], rules);
    } else {
      next[7].subs[i].m = '';
      next[7].subs[i].g = '';
    }
  }
  const hasSem2 = [4, 5, 6].some(
    (c) => getSubTotal(c) > 0 || next[c].subs.some((s) => s.m !== '' && s.m !== '-' && s.m !== undefined)
  );
  if (hasSem2) {
    const sem2Total = next[7].subs.reduce((acc, curr) => acc + toInt(curr.m), 0);
    next[7].total.m = sem2Total;
    next[7].total.p = calcWeightPercent(sem2Total, 7, maxTotal, maxWeightPct);
    next[7].total.g = getGradeFromMarks(sem2Total, maxTotal[7], rules);
  } else {
    next[7].total.m = '';
    next[7].total.p = '';
    next[7].total.g = '';
  }

  // 4. Grand Total 100% (Index 8: Sem1 + Sem2)
  for (let i = 0; i < subjects.length; i++) {
    const hasInput = [0, 1, 2, 4, 5, 6].some((c) => {
      const val = next[c].subs[i]?.m;
      return val !== '' && val !== '-' && val !== undefined;
    });
    if (hasInput) {
      const m = toInt(next[3].subs[i]?.m) + toInt(next[7].subs[i]?.m);
      next[8].subs[i].m = m;
      next[8].subs[i].g = getGradeFromMarks(m, maxPerSub[8], rules);
    } else {
      next[8].subs[i].m = '';
      next[8].subs[i].g = '';
    }
  }
  const hasAnyMarks = hasSem1 || hasSem2;
  if (hasAnyMarks) {
    const grandTotal = next[8].subs.reduce((acc, curr) => acc + toInt(curr.m), 0);
    next[8].total.m = grandTotal;
    next[8].total.p = calcWeightPercent(grandTotal, 8, maxTotal, maxWeightPct);
    next[8].total.g = getGradeFromMarks(grandTotal, maxTotal[8], rules);
  } else {
    next[8].total.m = '';
    next[8].total.p = '';
    next[8].total.g = '';
  }

  return next;
}

const createEmptyAssessmentsHelper = (cls = '', level?: '1-5' | '6-9'): AssessmentItem[] => {
  const parts = getAssessmentPartsForClass(cls, level);
  const subjects = getSubjectsForClass(cls, level);
  return parts.map((part, idx) => {
    const is50 = idx === 3 || idx === 7;
    const is100 = idx === 8;
    return {
      label: `${part.name} ${part.pct}`,
      subs: subjects.map(() => ({ m: '', g: '' })),
      total: { m: '', p: '', g: '' },
      ...(is50 ? { is50: true } : {}),
      ...(is100 ? { is100: true } : {}),
    };
  });
};

const INITIAL_STUDENTS: Student[] = [
  {
    id: 1,
    name: '',
    admission: '',
    dob: '',
    sats: '',
    pen: '',
    attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
    co1: Array(7).fill(''),
    co2: Array(7).fill(''),
    result: '',
    assessments: createEmptyAssessmentsHelper(),
  },
  {
    id: 2,
    name: '',
    admission: '',
    dob: '',
    sats: '',
    pen: '',
    attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
    co1: Array(7).fill(''),
    co2: Array(7).fill(''),
    result: '',
    assessments: createEmptyAssessmentsHelper(),
  },
  {
    id: 3,
    name: '',
    admission: '',
    dob: '',
    sats: '',
    pen: '',
    attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
    co1: Array(7).fill(''),
    co2: Array(7).fill(''),
    result: '',
    assessments: createEmptyAssessmentsHelper(),
  },
];

function isStudentFilled(s: Student): boolean {
  if (!s) return false;
  if (s.name && s.name.trim() !== '') return true;
  if (s.admission && s.admission.trim() !== '') return true;
  if (s.gender && s.gender.trim() !== '') return true;
  if (s.dob && s.dob.trim() !== '') return true;
  if (s.sats && s.sats.trim() !== '') return true;
  if (s.pen && s.pen.trim() !== '') return true;
  if (s.result && s.result.trim() !== '') return true;

  if (s.attendance) {
    if (
      (s.attendance.s1Total || 0) > 0 ||
      (s.attendance.s2Total || 0) > 0 ||
      (s.attendance.s1Present || 0) > 0 ||
      (s.attendance.s2Present || 0) > 0
    ) {
      return true;
    }
  }

  if (s.co1 && s.co1.some((c) => c && String(c).trim() !== '')) return true;
  if (s.co2 && s.co2.some((c) => c && String(c).trim() !== '')) return true;

  if (s.assessments) {
    for (const a of s.assessments) {
      if (a.subs) {
        for (const sub of a.subs) {
          if (sub.m !== undefined && sub.m !== null && String(sub.m).trim() !== '') {
            return true;
          }
        }
      }
    }
  }

  return false;
}

export default function App() {
  const [gradeRules, setGradeRules] = useState<GradeItemRule[]>(() => {
    try {
      const saved = localStorage.getItem('cce_grade_rules_custom');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return normalizeGradeRules(parsed);
      }
    } catch (e) {
      console.warn('Could not load saved grade rules', e);
    }
    return DEFAULT_GRADE_RULES;
  });
  const [isEditingGradeRules, setIsEditingGradeRules] = useState(false);
  const [editableRules, setEditableRules] = useState<GradeItemRule[]>([]);

  const [classLevel, setClassLevel] = useState<'1-5' | '6-9'>(() => {
    try {
      const savedClass = localStorage.getItem('cce_selected_class') || '';
      if (savedClass && ['1', '2', '3', '4', '5'].includes(savedClass)) return '1-5';
      const savedLevel = localStorage.getItem('cce_class_level');
      if (savedLevel === '1-5' || savedLevel === '6-9') return savedLevel;
    } catch (e) {}
    return '6-9';
  });

  const [selectedClass, setSelectedClass] = useState(() => {
    try {
      return localStorage.getItem('cce_selected_class') || '';
    } catch (e) {
      return '';
    }
  });

  const ALL_CLASSES_LIST = React.useMemo(() => {
    return classLevel === '1-5' ? CLASSES_1_5 : CLASSES_6_9;
  }, [classLevel]);

  const [students, setStudents] = useState<Student[]>(() => {
    let initialRules = DEFAULT_GRADE_RULES;
    try {
      const savedRules = localStorage.getItem('cce_grade_rules_custom');
      if (savedRules) {
        const parsed = JSON.parse(savedRules);
        if (Array.isArray(parsed) && parsed.length > 0) initialRules = normalizeGradeRules(parsed);
      }
    } catch (e) {}

    const curClass = (() => {
      try {
        return localStorage.getItem('cce_selected_class') || '';
      } catch (e) {
        return '';
      }
    })();

    if (curClass) {
      try {
        const classSaved = localStorage.getItem(`cce_students_class_${curClass}`);
        if (classSaved) {
          const parsed = JSON.parse(classSaved);
          if (Array.isArray(parsed) && parsed.length > 0) {
            return parsed.map((s: Student) => ({
              ...s,
              assessments: recalculateAssessments(s.assessments, initialRules, curClass),
            }));
          }
        }
      } catch (e) {}
    }

    try {
      const savedStudents = localStorage.getItem('cce_students');
      if (savedStudents) {
        const parsed = JSON.parse(savedStudents);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed.map((s: Student) => ({
            ...s,
            assessments: recalculateAssessments(s.assessments, initialRules, curClass || '6'),
          }));
        }
      }
    } catch (e) {}

    return INITIAL_STUDENTS.map((s) => ({
      ...s,
      assessments: recalculateAssessments(s.assessments, initialRules, curClass || '6'),
    }));
  });

  const [year, setYear] = useState(() => {
    try {
      return localStorage.getItem('cce_year') || '';
    } catch (e) {
      return '';
    }
  });

  const [school, setSchool] = useState(() => {
    try {
      return localStorage.getItem('cce_school') || '';
    } catch (e) {
      return '';
    }
  });

  const [diseCode, setDiseCode] = useState(() => {
    try {
      return localStorage.getItem('cce_dise_code') || '';
    } catch (e) {
      return '';
    }
  });

  const [searchQuery, setSearchQuery] = useState('');
  const [lastSavedTime, setLastSavedTime] = useState<string>(() => {
    const now = new Date();
    return now.toLocaleTimeString('en-US', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true,
    });
  });
  const [isSaving, setIsSaving] = useState(false);
  const saveTimerRef = useRef<NodeJS.Timeout | null>(null);

  const [notification, setNotification] = useState('');
  const [alertModal, setAlertModal] = useState<{
    show: boolean;
    title: string;
    message: string;
    maxAllowed?: number;
    onClose?: () => void;
  } | null>(null);
  const [resetModal, setResetModal] = useState<{
    show: boolean;
    type: 'all' | 'marks';
    resetGradeRulesToDefault: boolean;
  } | null>(null);
  const [showGradeModal, setShowGradeModal] = useState(false);
  const [deleteStudentModal, setDeleteStudentModal] = useState<{ index: number; name: string } | null>(null);
  const [hoveredStudentIdx, setHoveredStudentIdx] = useState<number | null>(null);

  // Bulk Student Import States
  const [showBulkImportModal, setShowBulkImportModal] = useState(false);
  const [importTab, setImportTab] = useState<'file' | 'text'>('file');
  const [pastedText, setPastedText] = useState('');
  const [parsedImportStudents, setParsedImportStudents] = useState<Partial<Student>[]>([]);
  const [importMode, setImportMode] = useState<'append' | 'replace'>('append');
  const [importStatusMsg, setImportStatusMsg] = useState<{ type: 'success' | 'error' | 'info'; text: string } | null>(null);

  const [languageNames, setLanguageNames] = useState<{ lang1: string; lang2: string; lang3: string }>(() => {
    const curClass = (() => {
      try {
        return localStorage.getItem('cce_selected_class') || '';
      } catch (e) {
        return '';
      }
    })();

    if (curClass) {
      try {
        const classLangs = localStorage.getItem(`cce_language_names_class_${curClass}`);
        if (classLangs) return JSON.parse(classLangs);
      } catch (e) {}
    }

    try {
      const saved = localStorage.getItem('cce_language_names');
      if (saved) return JSON.parse(saved);
    } catch (e) {}
    return { lang1: '', lang2: '', lang3: '' };
  });

  const isSwitchingClassRef = useRef(false);

  // Auto-save all state variables to localStorage on every change (mapped per class)
  useEffect(() => {
    if (isSwitchingClassRef.current) {
      isSwitchingClassRef.current = false;
      return;
    }

    const doSave = () => {
      try {
        if (selectedClass) {
          localStorage.setItem(`cce_students_class_${selectedClass}`, JSON.stringify(students));
          localStorage.setItem(`cce_language_names_class_${selectedClass}`, JSON.stringify(languageNames));
        }
        localStorage.setItem('cce_students', JSON.stringify(students));
        localStorage.setItem('cce_selected_class', selectedClass);
        localStorage.setItem('cce_school', school);
        localStorage.setItem('cce_year', year);
        localStorage.setItem('cce_dise_code', diseCode);
        localStorage.setItem('cce_language_names', JSON.stringify(languageNames));
        const now = new Date();
        const formattedTime = now.toLocaleTimeString('en-US', {
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
          hour12: true,
        });
        setLastSavedTime(formattedTime);

        // Whole line blink while entering data
        setIsSaving(true);
        if (saveTimerRef.current) clearTimeout(saveTimerRef.current);
        saveTimerRef.current = setTimeout(() => {
          setIsSaving(false);
        }, 1200);
      } catch (e) {
        console.warn('Auto-save error', e);
      }
    };
    doSave();

    const handleBeforeUnload = () => {
      doSave();
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    window.addEventListener('pagehide', handleBeforeUnload);
    return () => {
      window.removeEventListener('beforeunload', handleBeforeUnload);
      window.removeEventListener('pagehide', handleBeforeUnload);
    };
  }, [students, selectedClass, school, year, diseCode, languageNames]);

  const handleLanguageChange = useCallback((langKey: 'lang1' | 'lang2' | 'lang3', value: string) => {
    setLanguageNames((prev) => {
      const updated = { ...prev, [langKey]: value };
      try {
        if (selectedClass) {
          localStorage.setItem(`cce_language_names_class_${selectedClass}`, JSON.stringify(updated));
        }
        localStorage.setItem('cce_language_names', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });
  }, [selectedClass]);

  const isCustomRulesActive =
    JSON.stringify(gradeRules) !== JSON.stringify(DEFAULT_GRADE_RULES);

  const gradeTooltipText = gradeRules
    .map((r) => `${r.grade}: ${r.text100 || `${r.min100}%+`}`)
    .join(', ');

  const handleStartEditRules = useCallback(() => {
    setEditableRules(normalizeGradeRules(JSON.parse(JSON.stringify(gradeRules))));
    setIsEditingGradeRules(true);
  }, [gradeRules]);

  const handleAddRuleRow = useCallback(() => {
    setEditableRules((prev) => {
      const copy = prev.map((r) => ({ ...r }));
      const round1 = (n: number) => Math.round(n * 10) / 10;

      if (copy.length === 0) {
        return [
          {
            grade: 'D',
            min10: 0,
            max10: 1.9,
            text10: '0 - 1.9',
            min15: 0,
            max15: 2.9,
            text15: '0 - 2.9',
            min20: 0,
            max20: 3.9,
            text20: '0 - 3.9',
            min30: 0,
            max30: 5.9,
            text30: '0 - 5.9',
            min50: 0,
            max50: 9.9,
            text50: '0 - 9.9',
            min100: 0,
            max100: 19.9,
            text100: '0 - 19.9',
          },
        ];
      }

      const prevLastIndex = copy.length - 1;
      const above = copy[prevLastIndex];

      const newMax10 = above.max10 > 2.0 ? 1.9 : Math.max(0.5, round1(above.max10 - 1.0));
      const newMax15 = above.max15 > 3.0 ? 2.9 : Math.max(0.8, round1(above.max15 - 1.5));
      const newMax20 = above.max20 > 4.0 ? 3.9 : Math.max(1.0, round1(above.max20 - 2.0));
      const newMax30 = above.max30 > 7.0 ? 5.9 : Math.max(1.0, round1(above.max30 - 3.0));
      const newMax50 = above.max50 > 12.0 ? 9.9 : Math.max(2.0, round1(above.max50 - 5.0));
      const newMax100 = above.max100 > 22.0 ? 19.9 : Math.max(5.0, round1(above.max100 - 10.0));

      const aboveMin10 = round1(newMax10 + 0.1);
      const aboveMin15 = round1(newMax15 + 0.1);
      const aboveMin20 = round1(newMax20 + 0.1);
      const aboveMin30 = round1(newMax30 + 0.1);
      const aboveMin50 = round1(newMax50 + 0.1);
      const aboveMin100 = round1(newMax100 + 0.1);

      copy[prevLastIndex] = {
        ...above,
        min10: aboveMin10,
        text10: `${aboveMin10} - ${above.max10}`,
        min15: aboveMin15,
        text15: `${aboveMin15} - ${above.max15}`,
        min20: aboveMin20,
        text20: `${aboveMin20} - ${above.max20}`,
        min30: aboveMin30,
        text30: `${aboveMin30} - ${above.max30}`,
        min50: aboveMin50,
        text50: `${aboveMin50} - ${above.max50}`,
        min100: aboveMin100,
        text100: `${aboveMin100} - ${above.max100}`,
      };

      const lastGrade = above.grade.trim();
      let nextGrade = 'D';
      if (lastGrade === 'A+' || lastGrade === 'A') nextGrade = 'B';
      else if (lastGrade === 'B+' || lastGrade === 'B') nextGrade = 'C';
      else if (lastGrade === 'C+' || lastGrade === 'C') nextGrade = 'D';
      else if (lastGrade.length === 1 && lastGrade >= 'A' && lastGrade < 'Z') {
        nextGrade = String.fromCharCode(lastGrade.charCodeAt(0) + 1);
      } else {
        nextGrade = `D${copy.length}`;
      }

      const newRow: GradeItemRule = {
        grade: nextGrade,
        min10: 0,
        max10: newMax10,
        text10: `0 - ${newMax10}`,
        min15: 0,
        max15: newMax15,
        text15: `0 - ${newMax15}`,
        min20: 0,
        max20: newMax20,
        text20: `0 - ${newMax20}`,
        min30: 0,
        max30: newMax30,
        text30: `0 - ${newMax30}`,
        min50: 0,
        max50: newMax50,
        text50: `0 - ${newMax50}`,
        min100: 0,
        max100: newMax100,
        text100: `0 - ${newMax100}`,
      };

      return [...copy, newRow];
    });
  }, []);

  const handleRemoveRuleRow = useCallback((index: number) => {
    setEditableRules((prev) => {
      if (prev.length <= 1) {
        return prev;
      }
      return prev.filter((_, i) => i !== index);
    });
  }, []);

  const handleUpdateGradeName = useCallback((index: number, val: string) => {
    setEditableRules((prev) => {
      const copy = [...prev];
      copy[index] = { ...copy[index], grade: val };
      return copy;
    });
  }, []);

  const handleUpdateMinMax = useCallback((
    index: number,
    field: '10' | '15' | '20' | '30' | '50' | '100',
    bound: 'min' | 'max',
    val: string
  ) => {
    setEditableRules((prev) => {
      const copy = [...prev];
      const item = { ...copy[index] };
      const num = val === '' ? 0 : parseFloat(val);
      const safeNum = isNaN(num) ? 0 : num;

      if (field === '10') {
        if (bound === 'min') item.min10 = safeNum;
        else item.max10 = safeNum;
        item.text10 = `${item.min10} - ${item.max10}`;
      } else if (field === '15') {
        if (bound === 'min') item.min15 = safeNum;
        else item.max15 = safeNum;
        item.text15 = `${item.min15} - ${item.max15}`;
      } else if (field === '20') {
        if (bound === 'min') item.min20 = safeNum;
        else item.max20 = safeNum;
        item.text20 = `${item.min20} - ${item.max20}`;
      } else if (field === '30') {
        if (bound === 'min') item.min30 = safeNum;
        else item.max30 = safeNum;
        item.text30 = `${item.min30} - ${item.max30}`;
      } else if (field === '50') {
        if (bound === 'min') item.min50 = safeNum;
        else item.max50 = safeNum;
        item.text50 = `${item.min50} - ${item.max50}`;
      } else if (field === '100') {
        if (bound === 'min') item.min100 = safeNum;
        else item.max100 = safeNum;
        item.text100 = `${item.min100} - ${item.max100}`;
      }

      copy[index] = item;
      return copy;
    });
  }, []);

  const showAlert = useCallback((title: string, message: string, maxAllowed?: number, onClose?: () => void) => {
    setAlertModal({
      show: true,
      title,
      message,
      maxAllowed,
      onClose,
    });
  }, []);

  const handleCloseAlertModal = useCallback(() => {
    const onCloseCb = alertModal?.onClose;
    setAlertModal(null);
    if (onCloseCb) {
      setTimeout(() => {
        onCloseCb();
      }, 50);
    }
  }, [alertModal]);

  const isClassSelected = Boolean(selectedClass && selectedClass.trim());
  const isLang1Filled = Boolean(languageNames.lang1 && languageNames.lang1.trim());
  const isLang2Filled = Boolean(languageNames.lang2 && languageNames.lang2.trim());
  const isLang3Filled = Boolean(languageNames.lang3 && languageNames.lang3.trim());
  const isEntryAllowed = isClassSelected && isLang1Filled && isLang2Filled && isLang3Filled;

  const verifyPrerequisites = useCallback((e?: React.SyntheticEvent): boolean => {
    if (!isEntryAllowed) {
      if (e) {
        e.preventDefault();
        e.stopPropagation();
      }
      const missing: string[] = [];
      if (!isClassSelected) missing.push('ತರಗತಿ ಆಯ್ಕೆ (Class Select)');
      if (!isLang1Filled) missing.push('ಪ್ರಥಮ ಭಾಷೆಯ ಹೆಸರು (1st Language Name)');
      if (!isLang2Filled) missing.push('ದ್ವಿತೀಯ ಭಾಷೆಯ ಹೆಸರು (2nd Language Name)');
      if (!isLang3Filled) missing.push('ತೃತೀಯ ಭಾಷೆಯ ಹೆಸರು (3rd Language Name)');

      const targetFocusId = !isClassSelected
        ? 'headerClassSelect'
        : !isLang1Filled
        ? 'langInput-lang1'
        : !isLang2Filled
        ? 'langInput-lang2'
        : 'langInput-lang3';

      showAlert(
        'ಅಗತ್ಯ ವಿವರಗಳು ಅಪೂರ್ಣವಾಗಿವೆ!',
        `ನಮೂದು (Data Entry) ಪ್ರಾರಂಭಿಸಲು ಮೊದಲು ಈ ಕೆಳಗಿನ ವಿವರಗಳನ್ನು ಪೂರ್ಣಗೊಳಿಸಿ:\n\n• ${missing.join('\n• ')}\n\n('ಸರಿ (OK)' ಕ್ಲಿಕ್ ಮಾಡಿ, ಅದರಲ್ಲಿ ಭಾಷೆಯ ಹೆಸರು ನಮೂದಿಸಿ).`,
        undefined,
        () => {
          const el = document.getElementById(targetFocusId) as HTMLElement | null;
          if (el) {
            el.focus();
            if ('select' in el && typeof (el as HTMLInputElement).select === 'function') {
              (el as HTMLInputElement).select();
            }
          }
        }
      );
      return false;
    }
    return true;
  }, [isEntryAllowed, isClassSelected, isLang1Filled, isLang2Filled, isLang3Filled, showAlert]);

  const handleSaveGradeRules = useCallback(() => {
    if (editableRules.length === 0) {
      showAlert('ದೋಷ!', 'ಕನಿಷ್ಠ ಒಂದು ಶ್ರೇಣಿಯ ನಿಯಮವಿರಬೇಕು.');
      return;
    }
    for (const r of editableRules) {
      if (!r.grade.trim()) {
        showAlert('ದೋಷ!', 'ದಯವಿಟ್ಟು ಎಲ್ಲಾ ಶ್ರೇಣಿಗಳಿಗೆ ಹೆಸರನ್ನು (Grade Name) ನಮೂದಿಸಿ.');
        return;
      }
    }

    const curCls = selectedClass || '6';
    const { maxGrand } = getMaxMarksConfig(curCls);

    const cleanedRules: GradeItemRule[] = editableRules.map((r) => ({
      grade: r.grade.trim(),
      min10: Number(r.min10) || 0,
      max10: Number(r.max10) || 0,
      text10: `${r.min10} - ${r.max10}`,
      min15: Number(r.min15) || 0,
      max15: Number(r.max15) || 0,
      text15: `${r.min15} - ${r.max15}`,
      min20: Number(r.min20) || 0,
      max20: Number(r.max20) || 0,
      text20: `${r.min20} - ${r.max20}`,
      min30: Number(r.min30) || 0,
      max30: Number(r.max30) || 0,
      text30: `${r.min30} - ${r.max30}`,
      min50: Number(r.min50) || 0,
      max50: Number(r.max50) || 0,
      text50: `${r.min50} - ${r.max50}`,
      min100: Number(r.min100) || 0,
      max100: Number(r.max100) || 0,
      text100: `${r.min100} - ${r.max100}`,
    }));

    setGradeRules(cleanedRules);
    try {
      localStorage.setItem('cce_grade_rules_custom', JSON.stringify(cleanedRules));
    } catch (e) {
      console.error(e);
    }

    // Immediately link to main table: Recalculate ALL existing students
    setStudents((prev) =>
      prev.map((s) => {
        const updatedAssessments = recalculateAssessments(s.assessments, cleanedRules, curCls);
        let updatedResult = s.result;
        const grandTotalAssessment = updatedAssessments[8];
        const grandSum = toInt(grandTotalAssessment?.total?.m);
        if (grandSum > 0) {
          const finalPct = calcPercentExact(grandSum, maxGrand);
          updatedResult = finalPct >= 35 ? 'ಉತ್ತೀರ್ಣ' : 'ಅನುತ್ತೀರ್ಣ';
        }
        return {
          ...s,
          assessments: updatedAssessments,
          result: updatedResult,
        };
      })
    );

    setIsEditingGradeRules(false);
    setNotification('✅ ಶ್ರೇಣಿ ಮಾನದಂಡ ಮತ್ತು ವ್ಯಾಪ್ತಿಗಳನ್ನು ಉಳಿಸಲಾಗಿದೆ ಮತ್ತು ಮುಖ್ಯ ಕೋಷ್ಟಕಕ್ಕೆ ಲಿಂಕ್ ಮಾಡಲಾಗಿದೆ!');
    setTimeout(() => setNotification(''), 4500);
  }, [editableRules, showAlert, selectedClass]);

  const handleResetGradeRules = useCallback(() => {
    const defaults = JSON.parse(JSON.stringify(DEFAULT_GRADE_RULES));
    const curCls = selectedClass || '6';
    const { maxGrand } = getMaxMarksConfig(curCls);

    setGradeRules(defaults);
    setEditableRules(defaults);
    try {
      localStorage.removeItem('cce_grade_rules_custom');
    } catch (e) {
      console.error(e);
    }

    setStudents((prev) =>
      prev.map((s) => {
        const updatedAssessments = recalculateAssessments(s.assessments, defaults, curCls);
        let updatedResult = s.result;
        const grandTotalAssessment = updatedAssessments[8];
        const grandSum = toInt(grandTotalAssessment?.total?.m);
        if (grandSum > 0) {
          const finalPct = calcPercentExact(grandSum, maxGrand);
          updatedResult = finalPct >= 35 ? 'ಉತ್ತೀರ್ಣ' : 'ಅನುತ್ತೀರ್ಣ';
        }
        return {
          ...s,
          assessments: updatedAssessments,
          result: updatedResult,
        };
      })
    );

    setIsEditingGradeRules(false);
    setNotification('🔄 ಮೂಲ CCE ಶ್ರೇಣಿ ಮಾನದಂಡಗಳಿಗೆ ಮರುಹೊಂದಿಸಲಾಗಿದೆ ಮತ್ತು ಕೋಷ್ಟಕಕ್ಕೆ ಲಿಂಕ್ ಮಾಡಲಾಗಿದೆ!');
    setTimeout(() => setNotification(''), 4500);
  }, [selectedClass]);

  useEffect(() => {
    if (!alertModal?.show) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Enter' || e.key === 'Escape' || e.key === ' ') {
        e.preventDefault();
        handleCloseAlertModal();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [alertModal, handleCloseAlertModal]);
  const activeClassKeys = React.useMemo(() => {
    const activeKeys = ALL_CLASSES_LIST.filter((clsKey) => {
      if (clsKey === selectedClass) {
        return students.some(
          (s) =>
            (s.name && s.name.trim() !== '') ||
            s.admission ||
            (s.assessments && s.assessments.some((a) => a.subs && a.subs.some((sub) => sub.m !== '')))
        );
      }
      try {
        const saved = localStorage.getItem(`cce_students_class_${clsKey}`);
        if (saved) {
          const parsed = JSON.parse(saved);
          if (Array.isArray(parsed) && parsed.length > 0) {
            return parsed.some(
              (s: Student) =>
                (s.name && s.name.trim() !== '') ||
                s.admission ||
                (s.assessments && s.assessments.some((a: any) => a.subs && a.subs.some((sub: any) => sub.m !== '')))
            );
          }
        }
      } catch (e) {}
      return false;
    });

    if (activeKeys.length > 0) {
      return activeKeys.sort();
    }
    return selectedClass ? [selectedClass] : ['6'];
  }, [selectedClass, students, ALL_CLASSES_LIST]);

  const getClassData = useCallback(
    (clsKey: string) => {
      if (clsKey === selectedClass) {
        return {
          cls: clsKey,
          students,
          languageNames,
        };
      }

      let classStudents: Student[] = [];
      try {
        const classSaved = localStorage.getItem(`cce_students_class_${clsKey}`);
        if (classSaved) {
          const parsed = JSON.parse(classSaved);
          if (Array.isArray(parsed) && parsed.length > 0) {
            classStudents = parsed.map((s: Student) => ({
              ...s,
              assessments: recalculateAssessments(s.assessments, gradeRules, clsKey),
            }));
          }
        }
      } catch (e) {}

      if (!classStudents || classStudents.length === 0) {
        const coCount = getCoSkillsForClass(clsKey).length;
        classStudents = Array.from({ length: 10 }, (_, i) => ({
          id: i + 1,
          name: '',
          admission: '',
          gender: '',
          dob: '',
          sats: '',
          pen: '',
          attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
          co1: Array(coCount).fill(''),
          co2: Array(coCount).fill(''),
          result: '',
          assessments: recalculateAssessments(createEmptyAssessmentsHelper(clsKey), gradeRules, clsKey),
        }));
      }

      let classLangs = { lang1: '', lang2: '', lang3: '' };
      try {
        const classLangsSaved = localStorage.getItem(`cce_language_names_class_${clsKey}`);
        if (classLangsSaved) {
          classLangs = JSON.parse(classLangsSaved);
        }
      } catch (e) {}

      return {
        cls: clsKey,
        students: classStudents,
        languageNames: classLangs,
      };
    },
    [selectedClass, students, languageNames, gradeRules]
  );

  const checkHasStudentDataForPrint = useCallback((): boolean => {
    const isStudentFilled = (s: Student) => {
      if (!s) return false;
      if (s.name && s.name.trim() !== '') return true;
      if (s.admission && s.admission.trim() !== '') return true;
      if (s.sats && s.sats.trim() !== '') return true;
      if (s.pen && s.pen.trim() !== '') return true;
      if (s.assessments && s.assessments.some((a) => a.subs && a.subs.some((sub) => sub.m !== ''))) return true;
      if (s.co1 && s.co1.some((g) => g !== '')) return true;
      if (s.co2 && s.co2.some((g) => g !== '')) return true;
      return false;
    };

    if (students.some(isStudentFilled)) {
      return true;
    }

    const targetClasses = classLevel === '1-5' ? CLASSES_1_5 : CLASSES_6_9;
    for (const clsKey of targetClasses) {
      if (clsKey === selectedClass) continue;
      try {
        const saved = localStorage.getItem(`cce_students_class_${clsKey}`);
        if (saved) {
          const parsed = JSON.parse(saved);
          if (Array.isArray(parsed) && parsed.some(isStudentFilled)) {
            return true;
          }
        }
      } catch (e) {}
    }

    return false;
  }, [students, classLevel, selectedClass]);

  const handlePrint = useCallback(() => {
    if (!verifyPrerequisites()) return;

    if (!checkHasStudentDataForPrint()) {
      showAlert(
        '⚠️ ವಿದ್ಯಾರ್ಥಿ ಮಾಹಿತಿ ಅಪೂರ್ಣ!',
        `ಮುದ್ರಣ (Print) ಮಾಡಲು ಕನಿಷ್ಠ ಒಬ್ಬ ವಿದ್ಯಾರ್ಥಿಯ ವಿವರಗಳನ್ನು ಭರ್ತಿ ಮಾಡಿರಬೇಕು!\n\n(${classLevel === '1-5' ? '1 ರಿಂದ 5 ನೇ' : '6 ರಿಂದ 9 ನೇ'} ತರಗತಿಯ ವಹಿಯಲ್ಲಿ ಕನಿಷ್ಠ 1 ವಿದ್ಯಾರ್ಥಿಯ ಹೆಸರು, ದಾಖಲಾತಿ ಸಂಖ್ಯೆ ಅಥವಾ ಅಂಕಗಳನ್ನು ನಮೂದಿಸಿ ನಂತರ ಪ್ರಿಂಟ್ ಮಾಡಿ.)`
      );
      return;
    }

    if (selectedClass) {
      try {
        localStorage.setItem(`cce_students_class_${selectedClass}`, JSON.stringify(students));
        localStorage.setItem(`cce_language_names_class_${selectedClass}`, JSON.stringify(languageNames));
      } catch (e) {}
    }

    const prevTitle = document.title;
    document.title = '';
    window.focus();

    const titleMsg = is1to5Class(selectedClass, classLevel)
      ? '1 ರಿಂದ 5 ನೇ ತರಗತಿಯ ನಿರಂತರ ಮತ್ತು ವ್ಯಾಪಕ ಮೌಲ್ಯಮಾಪನ ಕ್ರೋಡೀಕೃತ ವಹಿ'
      : '6 ರಿಂದ 9 ನೇ ತರಗತಿಯ ನಿರಂತರ ಮತ್ತು ವ್ಯಾಪಕ ಮೌಲ್ಯಮಾಪನ ಕ್ರೋಡೀಕೃತ ವಹಿ';

    setTimeout(() => {
      window.print();
      setTimeout(() => {
        document.title = prevTitle || titleMsg;
      }, 500);
    }, 100);
  }, [selectedClass, classLevel, students, languageNames, verifyPrerequisites, checkHasStudentDataForPrint, showAlert]);


  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p') {
        e.preventDefault();
        e.stopPropagation();
        handlePrint();
      }
    };
    window.addEventListener('keydown', handleKeyDown, true);
    return () => window.removeEventListener('keydown', handleKeyDown, true);
  }, [handlePrint]);

  useEffect(() => {
    const handleBeforePrint = () => {
      document.title = '';
    };
    const handleAfterPrint = () => {
      document.title = is1to5Class(selectedClass, classLevel)
        ? '1 ರಿಂದ 5 ನೇ ತರಗತಿಯ ನಿರಂತರ ಮತ್ತು ವ್ಯಾಪಕ ಮೌಲ್ಯಮಾಪನ ಕ್ರೋಡೀಕೃತ ವಹಿ'
        : '6 ರಿಂದ 9 ನೇ ತರಗತಿಯ ನಿರಂತರ ಮತ್ತು ವ್ಯಾಪಕ ಮೌಲ್ಯಮಾಪನ ಕ್ರೋಡೀಕೃತ ವಹಿ';
    };
    window.addEventListener('beforeprint', handleBeforePrint);
    window.addEventListener('afterprint', handleAfterPrint);
    return () => {
      window.removeEventListener('beforeprint', handleBeforePrint);
      window.removeEventListener('afterprint', handleAfterPrint);
    };
  }, [selectedClass, classLevel]);

  const createEmptyAssessments = (): AssessmentItem[] => {
    return createEmptyAssessmentsHelper(selectedClass || '6');
  };

  const createEmptyStudent = useCallback((id: number, rules: GradeItemRule[] = gradeRules, cls = selectedClass, level = classLevel): Student => {
    const effectiveCls = cls || (level === '1-5' ? '1' : '6');
    const coCount = getCoSkillsForClass(effectiveCls, level).length;
    return {
      id,
      name: '',
      admission: '',
      gender: '',
      dob: '',
      sats: '',
      pen: '',
      attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
      co1: Array(coCount).fill(''),
      co2: Array(coCount).fill(''),
      result: '',
      assessments: recalculateAssessments(createEmptyAssessmentsHelper(effectiveCls, level), rules, effectiveCls, level),
    };
  }, [gradeRules, selectedClass, classLevel]);

  const handleClassChange = useCallback((newClass: string) => {
    // 1. Save current class state
    if (selectedClass) {
      try {
        localStorage.setItem(`cce_students_class_${selectedClass}`, JSON.stringify(students));
        localStorage.setItem(`cce_language_names_class_${selectedClass}`, JSON.stringify(languageNames));
      } catch (e) {
        console.warn('Failed to save current class before switch', e);
      }
    }

    if (!newClass) {
      setSelectedClass('');
      try {
        localStorage.setItem('cce_selected_class', '');
      } catch (e) {}
      return;
    }

    // Auto sync level state
    if (is1to5Class(newClass)) {
      setClassLevel('1-5');
      try { localStorage.setItem('cce_class_level', '1-5'); } catch (e) {}
    } else {
      setClassLevel('6-9');
      try { localStorage.setItem('cce_class_level', '6-9'); } catch (e) {}
    }

    // 2. Load target class data
    let targetStudents: Student[] | null = null;
    let targetLanguages: { lang1: string; lang2: string; lang3: string } | null = null;

    try {
      const saved = localStorage.getItem(`cce_students_class_${newClass}`);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          targetStudents = parsed.map((s: Student) => ({
            ...s,
            assessments: recalculateAssessments(s.assessments, gradeRules, newClass),
          }));
        }
      }
    } catch (e) {}

    try {
      const savedLangs = localStorage.getItem(`cce_language_names_class_${newClass}`);
      if (savedLangs) {
        targetLanguages = JSON.parse(savedLangs);
      }
    } catch (e) {}

    if (!targetStudents || targetStudents.length === 0) {
      targetStudents = [1, 2, 3].map((id) => createEmptyStudent(id, gradeRules, newClass));
    }

    if (!targetLanguages) {
      targetLanguages = { lang1: '', lang2: '', lang3: '' };
    }

    try {
      localStorage.setItem(`cce_students_class_${newClass}`, JSON.stringify(targetStudents));
      localStorage.setItem(`cce_language_names_class_${newClass}`, JSON.stringify(targetLanguages));
      localStorage.setItem('cce_students', JSON.stringify(targetStudents));
      localStorage.setItem('cce_language_names', JSON.stringify(targetLanguages));
      localStorage.setItem('cce_selected_class', newClass);
    } catch (e) {}

    isSwitchingClassRef.current = true;

    setSelectedClass(newClass);
    setStudents(targetStudents);
    setLanguageNames(targetLanguages);

    const now = new Date();
    const formattedTime = now.toLocaleTimeString('en-US', {
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hour12: true,
    });
    setLastSavedTime(formattedTime);

    setNotification(`ತರಗತಿ ${newClass} ರ ದಾಖಲೆಗಳು ಲೋಡ್ ಆಗಿವೆ! (Class ${newClass} data loaded)`);
    setTimeout(() => setNotification(''), 3000);
  }, [selectedClass, students, languageNames, gradeRules, createEmptyStudent]);

  const handleLevelToggle = useCallback((newLevel: '1-5' | '6-9') => {
    setClassLevel(newLevel);
    try {
      localStorage.setItem('cce_class_level', newLevel);
    } catch (e) {}
    const targetClasses = newLevel === '1-5' ? CLASSES_1_5 : CLASSES_6_9;
    if (!targetClasses.includes(selectedClass)) {
      handleClassChange(targetClasses[0]);
    }
  }, [selectedClass, handleClassChange]);

  const handleAddStudent = () => {
    if (!verifyPrerequisites()) return;
    setStudents((prev) => {
      const newId = prev.length + 1;
      const curCls = selectedClass || (classLevel === '1-5' ? '1' : '6');
      const coCount = getCoSkillsForClass(curCls, classLevel).length;
      let newAssessments = createEmptyAssessmentsHelper(curCls, classLevel);
      newAssessments = recalculateAssessments(newAssessments, gradeRules, curCls, classLevel);
      const newStudent: Student = {
        id: newId,
        name: '',
        admission: '',
        gender: '',
        dob: '',
        sats: '',
        pen: '',
        attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
        co1: Array(coCount).fill(''),
        co2: Array(coCount).fill(''),
        result: '',
        assessments: newAssessments,
      };
      return [...prev, newStudent];
    });

    setTimeout(() => {
      const container = document.querySelector('.table-wrap');
      if (container) container.scrollTop = container.scrollHeight;
      window.scrollTo({ top: document.body.scrollHeight, behavior: 'smooth' });
    }, 80);
  };

  const handleRequestDeleteStudent = (sIdx: number) => {
    const stu = students[sIdx];
    const stuName = stu?.name?.trim();
    const displayName = stuName ? stuName : `ಕ್ರಮ ಸಂಖ್ಯೆ ${sIdx + 1} (ವಿದ್ಯಾರ್ಥಿ #${stu?.id || sIdx + 1})`;
    setDeleteStudentModal({ index: sIdx, name: displayName });
  };

  const confirmDeleteStudent = () => {
    if (!deleteStudentModal) return;
    const sIdx = deleteStudentModal.index;
    const curCls = selectedClass || (classLevel === '1-5' ? '1' : '6');
    setStudents((prev) => {
      if (prev.length <= 1) {
        return [{
          ...prev[0],
          name: '',
          admission: '',
          dob: '',
          sats: '',
          pen: '',
          attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
          assessments: recalculateAssessments(createEmptyAssessmentsHelper(curCls, classLevel), gradeRules, curCls, classLevel),
        }];
      }
      return prev.filter((_, idx) => idx !== sIdx).map((s, idx) => ({ ...s, id: idx + 1 }));
    });
    setNotification('🗑️ ವಿದ್ಯಾರ್ಥಿಯನ್ನು ಯಶಸ್ವಿಯಾಗಿ ತೆಗೆದುಹಾಕಲಾಗಿದೆ');
    setTimeout(() => setNotification(''), 3000);
    setDeleteStudentModal(null);
  };

  const handleResetAll = useCallback(() => {
    setResetModal({
      show: true,
      type: 'all',
      resetGradeRulesToDefault: true,
    });
  }, []);

  const handleResetMarks = useCallback(() => {
    setResetModal({
      show: true,
      type: 'marks',
      resetGradeRulesToDefault: true,
    });
  }, []);

  const handleConfirmReset = useCallback(() => {
    if (!resetModal) return;

    let targetGradeRules = gradeRules;

    if (resetModal.resetGradeRulesToDefault) {
      const defaults = JSON.parse(JSON.stringify(DEFAULT_GRADE_RULES));
      targetGradeRules = defaults;
      setGradeRules(defaults);
      setEditableRules(defaults);
      try {
        localStorage.removeItem('cce_grade_rules_custom');
      } catch (e) {
        console.error(e);
      }
    }

    const defaultCls = selectedClass || (classLevel === '1-5' ? '1' : '6');
    const coCount = getCoSkillsForClass(defaultCls, classLevel).length;

    if (resetModal.type === 'all') {
      setYear('');
      setSchool('');
      setDiseCode('');
      setSelectedClass('');
      setLanguageNames({ lang1: '', lang2: '', lang3: '' });
      try {
        localStorage.removeItem('cce_dise_code');
        localStorage.removeItem('cce_selected_class');
        localStorage.removeItem('cce_school');
        localStorage.removeItem('cce_year');
        localStorage.removeItem('cce_students');
        localStorage.removeItem('cce_language_names');
        ['1', '2', '3', '4', '5', '6', '7', '8', '9'].forEach((cls) => {
          localStorage.removeItem(`cce_students_class_${cls}`);
          localStorage.removeItem(`cce_language_names_class_${cls}`);
        });
      } catch (e) {}
      setStudents([1, 2, 3].map((id) => createEmptyStudent(id, targetGradeRules, defaultCls, classLevel)));
      setNotification(
        resetModal.resetGradeRulesToDefault
          ? '🔄 ಎಲ್ಲಾ ಡೇಟಾ ಮರುಹೊಂದಿಸಲಾಗಿದೆ ಮತ್ತು ಅಧಿಕೃತ ಶ್ರೇಣಿ ಕೋಷ್ಟಕವನ್ನು ಮೂಲ ಡಿಫಾಲ್ಟ್‌ಗೆ ಮರುಸ್ಥಾಪಿಸಲಾಗಿದೆ!'
          : '🔄 ಎಲ್ಲಾ ಡೇಟಾ ಮರುಹೊಂದಿಸಲಾಗಿದೆ! (ಕಸ್ಟಮ್ ಶ್ರೇಣಿ ನಿಯಮಗಳನ್ನು ಉಳಿಸಿಕೊಳ್ಳಲಾಗಿದೆ)'
      );
    } else {
      setStudents((prev) =>
        prev.map((s) => ({
          ...s,
          attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
          co1: Array(coCount).fill(''),
          co2: Array(coCount).fill(''),
          result: '',
          assessments: recalculateAssessments(createEmptyAssessmentsHelper(defaultCls, classLevel), targetGradeRules, defaultCls, classLevel),
        }))
      );
      setNotification(
        resetModal.resetGradeRulesToDefault
          ? '🔄 ಎಲ್ಲಾ ಅಂಕಗಳನ್ನು ಮರುಹೊಂದಿಸಲಾಗಿದೆ ಮತ್ತು ಅಧಿಕೃತ ಶ್ರೇಣಿ ಕೋಷ್ಟಕವನ್ನು ಮೂಲ ಡಿಫಾಲ್ಟ್‌ಗೆ ಮರುಸ್ಥಾಪಿಸಲಾಗಿದೆ!'
          : '🔄 ಎಲ್ಲಾ ಅಂಕಗಳನ್ನು ಮರುಹೊಂದಿಸಲಾಗಿದೆ! (ಕಸ್ಟಮ್ ಶ್ರೇಣಿ ನಿಯಮಗಳನ್ನು ಉಳಿಸಿಕೊಳ್ಳಲಾಗಿದೆ)'
      );
    }

    setResetModal(null);
    setTimeout(() => setNotification(''), 4500);
  }, [resetModal, gradeRules, classLevel, selectedClass, createEmptyStudent]);

  // Sync window global helpers for external invocations
  useEffect(() => {
    (window as unknown as Record<string, unknown>).resetAllData = handleResetAll;
    (window as unknown as Record<string, unknown>).resetMarksOnly = handleResetMarks;
  }, [handleResetAll, handleResetMarks]);

  const handleExportExcel = () => {
    const curCls = selectedClass || '6';
    const subjects = getSubjectsForClass(curCls);
    const coSkills = getCoSkillsForClass(curCls);
    const parts = getAssessmentPartsForClass(curCls);
    const { maxGrand } = getMaxMarksConfig(curCls);

    const filename = `CCE_Class${selectedClass || 'Register'}_${(school.trim() || 'School').replace(/[\/\\:*?"<>|]+/g, '_')}_${diseCode ? `DISE_${diseCode}_` : ''}${year || '2026-27'}`;
    const headerRow: string[] = [
      'ಕ್ರ.ಸಂ.',
      'ಹೆಸರು',
      'ದಾಖಲಾತಿ ಸಂಖ್ಯೆ',
      'ಲಿಂಗ',
      'ಜನ್ಮ ದಿನಾಂಕ',
      'SATS No',
      'PEN No',
      '1ನೇ ಸೆಮಿ ಶಾಲೆ ದಿನ',
      '1ನೇ ಸೆಮಿ ಹಾಜರಾತಿ',
      '2ನೇ ಸೆಮಿ ಶಾಲೆ ದಿನ',
      '2ನೇ ಸೆಮಿ ಹಾಜರಾತಿ',
      'ಒಟ್ಟು ಶಾಲೆ ದಿನ',
      'ಒಟ್ಟು ಹಾಜರಾತಿ',
    ];

    parts.forEach((assPart) => {
      const assLbl = `${assPart.name} ${assPart.pct}`;
      subjects.forEach((sub) => {
        headerRow.push(`${assLbl} - ${sub} ಅಂಕ`, `${assLbl} - ${sub} ಶ್ರೇಣಿ`);
      });
      headerRow.push(`${assLbl} - ಒಟ್ಟು ಅಂಕ`, `${assLbl} - %`, `${assLbl} - ಶ್ರೇಣಿ`);
    });

    coSkills.forEach((co) => headerRow.push(`ಸಹಪಠ್ಯ 1ನೇ - ${co}`));
    coSkills.forEach((co) => headerRow.push(`ಸಹಪಠ್ಯ 2ನೇ - ${co}`));
    coSkills.forEach((co) => headerRow.push(`ಸಹಪಠ್ಯ ಅಂತಿಮ - ${co}`));
    headerRow.push('ಫಲಿತಾಂಶ');

    const dataRows: (string | number)[][] = [headerRow];

    students.forEach((stu) => {
      const totDays = (stu.attendance.s1Total || 0) + (stu.attendance.s2Total || 0);
      const totPres = (stu.attendance.s1Present || 0) + (stu.attendance.s2Present || 0);
      const row: (string | number)[] = [
        stu.id,
        stu.name,
        stu.admission,
        stu.gender || '-',
        formatDateDisplay(stu.dob),
        stu.sats,
        stu.pen,
        stu.attendance.s1Total,
        stu.attendance.s1Present,
        stu.attendance.s2Total,
        stu.attendance.s2Present,
        totDays,
        totPres,
      ];

      stu.assessments.forEach((ass) => {
        subjects.forEach((_, subIdx) => {
          const sub = ass.subs[subIdx] || { m: '', g: '' };
          row.push(sub.m, sub.g);
        });
        row.push(ass.total.m, ass.total.p, ass.total.g);
      });

      coSkills.forEach((_, idx) => row.push(stu.co1[idx] || ''));
      coSkills.forEach((_, idx) => row.push(stu.co2[idx] || ''));

      const finalCo = coSkills.map((_, idx) =>
        stu.co1[idx] === 'A' && stu.co2[idx] === 'A' ? 'A' : 'B'
      );
      row.push(...finalCo);

      const grandTotalMarks = toInt(stu.assessments[8]?.total?.m);
      const pct = calcPercentExact(grandTotalMarks, maxGrand);
      const pctRound = Math.round(pct);
      const gr = getGrade(pct, gradeRules);
      let resCell = '-';
      const resLabel = stu.result || (grandTotalMarks > 0 ? 'ಉತ್ತೀರ್ಣ' : '');
      if (resLabel) {
        resCell = grandTotalMarks > 0
          ? `${resLabel} = ( ${pctRound}% - ${gr || 'C'} )`
          : resLabel;
      }
      row.push(resCell);

      dataRows.push(row);
    });

    try {
      const worksheet = XLSX.utils.aoa_to_sheet(dataRows);
      const colWidths = headerRow.map((_, idx) => ({
        wch: idx < 1 ? 6 : idx === 1 ? 24 : idx < 12 ? 14 : 10,
      }));
      worksheet['!cols'] = colWidths;
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, `CCE Class ${selectedClass}`);
      XLSX.writeFile(workbook, `${filename}.xlsx`);
      setNotification(`${filename}.xlsx ರಫ್ತು ಮಾಡಲಾಗಿದೆ ✓`);
      setTimeout(() => setNotification(''), 3500);
      return;
    } catch (err) {
      console.warn('XLSX writing error, falling back to UTF-8 CSV', err);
    }

    // CSV Fallback
    const escapeCsv = (val: unknown) => {
      const s = String(val ?? '');
      if (s.includes('"') || s.includes(',') || s.includes('\n')) {
        return `"${s.replace(/"/g, '""')}"`;
      }
      return s;
    };
    const csvContent = '\uFEFF' + dataRows.map((r) => r.map(escapeCsv).join(',')).join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${filename}.csv`;
    document.body.appendChild(link);
    link.click();
    setTimeout(() => {
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }, 150);
    setNotification(`${filename}.csv ರಫ್ತು ಮಾಡಲಾಗಿದೆ ✓`);
    setTimeout(() => setNotification(''), 3500);
  };

  const handleDownloadImportTemplate = () => {
    const headerRow = [
      'ಕ್ರಮ ಸಂಖ್ಯೆ (Sl. No.)',
      'ವಿದ್ಯಾರ್ಥಿಯ ಹೆಸರು (Student Name) *',
      'ದಾಖಲಾತಿ ಸಂಖ್ಯೆ (Admission No)',
      'ಲಿಂಗ (Gender - M/F/O)',
      'ಜನ್ಮ ದಿನಾಂಕ (DOB - YYYY-MM-DD)',
      'SATS ಸಂಖ್ಯೆ (SATS No)',
      'PEN ಸಂಖ್ಯೆ (PEN No)',
    ];
    const sampleRows = [
      [1, 'ರಮೇಶ್ ಬಿ (Ramesh B)', '101', 'M', '2014-05-15', '123456789', '987654321012'],
      [2, 'ಕಾವ್ಯ ಎಂ (Kavya M)', '102', 'F', '2014-08-20', '987654321', '123456789012'],
    ];

    try {
      const ws = XLSX.utils.aoa_to_sheet([headerRow, ...sampleRows]);
      ws['!cols'] = [{ wch: 18 }, { wch: 30 }, { wch: 22 }, { wch: 20 }, { wch: 24 }, { wch: 20 }, { wch: 20 }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Student Import Template');
      XLSX.writeFile(wb, 'CCE_Student_Import_Template.xlsx');
      setNotification('📥 ಮಾದರಿ ಆಮದು ಟೆಂಪ್ಲೇಟ್ ಡೌನ್‌ಲೋಡ್ ಆಗಿದೆ ✓');
      setTimeout(() => setNotification(''), 3500);
    } catch (e) {
      console.error('Template download error', e);
    }
  };

  const normalizeGender = (val: unknown): string => {
    if (val === undefined || val === null) return '';
    const s = String(val).trim().toLowerCase();
    if (!s) return '';
    if (
      s === 'm' ||
      s === 'male' ||
      s === 'boy' ||
      s === 'boys' ||
      s === 'ಗಂಡು' ||
      s === 'gandu' ||
      s === 'ಬಾಲಕ' ||
      s === '1' ||
      s.startsWith('m') ||
      s.startsWith('g') ||
      s.startsWith('b')
    ) {
      return 'ಗಂಡು';
    }
    if (
      s === 'f' ||
      s === 'female' ||
      s === 'girl' ||
      s === 'girls' ||
      s === 'ಹೆಣ್ಣು' ||
      s === 'hennu' ||
      s === 'ಬಾಲಕಿ' ||
      s === '2' ||
      s.startsWith('f') ||
      s.startsWith('h')
    ) {
      return 'ಹೆಣ್ಣು';
    }
    if (s === 'o' || s === 'other' || s === 'ಇತರೆ' || s === '3') {
      return 'ಇತರೆ';
    }
    return String(val).trim();
  };

  const parseRowFields = (row: any[]): Partial<Student> | null => {
    if (!row || !Array.isArray(row) || row.length === 0) return null;
    const cleanStr = (val: any) => (val !== undefined && val !== null ? String(val).trim() : '');

    // Ignore header row
    const rowStr = row.map(cleanStr).join(' ').toLowerCase();
    if (
      rowStr.includes('ಹೆಸರು') ||
      rowStr.includes('student name') ||
      rowStr.includes('ದಾಖಲಾತಿ') ||
      rowStr.includes('admission')
    ) {
      return null;
    }

    let name = '';
    let admission = '';
    let gender = '';
    let dob = '';
    let sats = '';
    let pen = '';

    const cells = row.map(cleanStr).filter((c) => c !== '');
    if (cells.length === 0) return null;

    let cellIdx = 0;
    // Skip Sl. No if present
    if (/^\d+$/.test(cells[0]) && cells.length > 1 && !/^\d{9,12}$/.test(cells[0])) {
      cellIdx = 1;
    }

    name = cells[cellIdx] || '';

    // Smart cell type detection for remaining cells
    for (let i = cellIdx + 1; i < cells.length; i++) {
      const val = cells[i];
      const g = normalizeGender(val);

      if (!gender && (g === 'ಗಂಡು' || g === 'ಹೆಣ್ಣು' || g === 'ಇತರೆ')) {
        gender = g;
      } else if (!dob && (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(val) || /^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(val))) {
        dob = val;
      } else if (!sats && /^\d{9}$/.test(val)) {
        sats = val;
      } else if (!pen && /^\d{11,12}$/.test(val)) {
        pen = val;
      } else if (!admission) {
        admission = val;
      }
    }

    if (!name && !admission && !sats) return null;
    return { name, admission, gender, dob, sats, pen };
  };

  const handleBulkFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });

        if (!rows || rows.length === 0) {
          setImportStatusMsg({ type: 'error', text: '❌ ಫೈಲ್ ಖಾಲಿಯಾಗಿದೆ.' });
          return;
        }

        // Smart header detection
        let headerRowIdx = -1;
        let colNameIdx = -1;
        let colAdmIdx = -1;
        let colGenIdx = -1;
        let colDobIdx = -1;
        let colSatsIdx = -1;
        let colPenIdx = -1;

        for (let r = 0; r < Math.min(5, rows.length); r++) {
          const rowCells = (rows[r] || []).map((c) => String(c || '').trim().toLowerCase());
          const nameIdx = rowCells.findIndex((c) => c.includes('ಹೆಸರು') || c.includes('name'));
          if (nameIdx >= 0) {
            headerRowIdx = r;
            colNameIdx = nameIdx;
            colAdmIdx = rowCells.findIndex((c) => c.includes('ದಾಖಲಾತಿ') || c.includes('admission') || c.includes('reg'));
            colGenIdx = rowCells.findIndex((c) => c.includes('ಲಿಂಗ') || c.includes('gender') || c.includes('sex') || c === 'm/f');
            colDobIdx = rowCells.findIndex((c) => c.includes('ದಿನಾಂಕ') || c.includes('dob') || c.includes('birth'));
            colSatsIdx = rowCells.findIndex((c) => c.includes('sats'));
            colPenIdx = rowCells.findIndex((c) => c.includes('pen'));
            break;
          }
        }

        const parsed: Partial<Student>[] = [];

        if (headerRowIdx >= 0 && colNameIdx >= 0) {
          const cleanStr = (val: any) => (val !== undefined && val !== null ? String(val).trim() : '');
          for (let r = headerRowIdx + 1; r < rows.length; r++) {
            const row = rows[r];
            if (!row || !Array.isArray(row) || row.length === 0) continue;
            const name = cleanStr(row[colNameIdx]);
            const admission = colAdmIdx >= 0 ? cleanStr(row[colAdmIdx]) : '';
            const gender = normalizeGender(colGenIdx >= 0 ? row[colGenIdx] : '');
            const dob = colDobIdx >= 0 ? cleanStr(row[colDobIdx]) : '';
            const sats = colSatsIdx >= 0 ? cleanStr(row[colSatsIdx]) : '';
            const pen = colPenIdx >= 0 ? cleanStr(row[colPenIdx]) : '';
            if (name || admission || sats) {
              parsed.push({ name, admission, gender, dob, sats, pen });
            }
          }
        } else {
          rows.forEach((row) => {
            const res = parseRowFields(row);
            if (res) parsed.push(res);
          });
        }

        if (parsed.length > 0) {
          setParsedImportStudents(parsed);
          setImportStatusMsg({
            type: 'success',
            text: `✓ ${parsed.length} ವಿದ್ಯಾರ್ಥಿಗಳನ್ನು ಫೈಲ್‌ನಿಂದ ಯಶಸ್ವಿಯಾಗಿ ಗುರುತಿಸಲಾಗಿದೆ!`,
          });
        } else {
          setImportStatusMsg({
            type: 'error',
            text: '❌ ಫೈಲ್‌ನಲ್ಲಿ ವಿದ್ಯಾರ್ಥಿ ವಿವರಗಳು ಕಂಡುಬಂದಿಲ್ಲ. ದಯವಿಟ್ಟು ಮಾದರಿ ಟೆಂಪ್ಲೇಟ್ ಬಳಸಿ.',
          });
        }
      } catch (err) {
        setImportStatusMsg({
          type: 'error',
          text: '❌ ಫೈಲ್ ಓದಲು ಸಾಧ್ಯವಾಗಲಿಲ್ಲ. ದಯವಿಟ್ಟು ಮಾನ್ಯ Excel (.xlsx/.xls) ಅಥವಾ CSV ಫೈಲ್ ಆಯ್ಕೆಮಾಡಿ.',
        });
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  const handleParsePastedText = () => {
    if (!pastedText.trim()) {
      setImportStatusMsg({ type: 'error', text: '❌ ದಯವಿಟ್ಟು ಪಠ್ಯ ಪ್ರದೇಶದಲ್ಲಿ ವಿದ್ಯಾರ್ಥಿಗಳ ಪಟ್ಟಿಯನ್ನು ಪೇಸ್ಟ್ ಮಾಡಿ.' });
      return;
    }

    const lines = pastedText.split('\n');
    const parsed: Partial<Student>[] = [];

    lines.forEach((line) => {
      if (!line.trim()) return;
      const parts = line.split(/[\t,|]/).map((s) => s.trim());
      const res = parseRowFields(parts);
      if (res) parsed.push(res);
    });

    if (parsed.length > 0) {
      setParsedImportStudents(parsed);
      setImportStatusMsg({
        type: 'success',
        text: `✓ ${parsed.length} ವಿದ್ಯಾರ್ಥಿಗಳನ್ನು ಪೇಸ್ಟ್ ಮಾಡಿದ ಪಠ್ಯದಿಂದ ಗುರುತಿಸಲಾಗಿದೆ!`,
      });
    } else {
      setImportStatusMsg({
        type: 'error',
        text: '❌ ಪೇಸ್ಟ್ ಮಾಡಿದ ಪಠ್ಯದಲ್ಲಿ ಮಾನ್ಯ ವಿದ್ಯಾರ್ಥಿ ಸಾಲುಗಳು ಕಂಡುಬಂದಿಲ್ಲ.',
      });
    }
  };

  const handleConfirmBulkImport = () => {
    if (!parsedImportStudents || parsedImportStudents.length === 0) return;
    if (!verifyPrerequisites()) return;

    const curCls = selectedClass || '6';
    const coCount = getCoSkillsForClass(curCls).length;

    const newImported: Student[] = parsedImportStudents.map((item, idx) => {
      const emptyAssessments = createEmptyAssessmentsHelper(curCls);
      const recalculated = recalculateAssessments(emptyAssessments, gradeRules, curCls);
      return {
        id: idx + 1,
        name: (item.name || '').trim(),
        admission: (item.admission || '').trim(),
        gender: normalizeGender(item.gender),
        dob: (item.dob || '').trim(),
        sats: (item.sats || '').trim(),
        pen: (item.pen || '').trim(),
        attendance: { s1Total: 0, s2Total: 0, s1Present: 0, s2Present: 0 },
        co1: Array(coCount).fill(''),
        co2: Array(coCount).fill(''),
        result: '',
        assessments: recalculated,
      };
    });

    setStudents((prev) => {
      const isTableEmpty = !prev.some(isStudentFilled);
      const effectiveMode = isTableEmpty ? 'replace' : importMode;
      let updated: Student[];
      if (effectiveMode === 'append') {
        updated = [...prev, ...newImported];
      } else {
        updated = newImported;
      }
      return updated.map((s, i) => ({
        ...s,
        id: i + 1,
        co1: s.co1 && s.co1.length === coCount ? s.co1 : Array(coCount).fill(''),
        co2: s.co2 && s.co2.length === coCount ? s.co2 : Array(coCount).fill(''),
        assessments: recalculateAssessments(s.assessments, gradeRules, curCls),
      }));
    });

    setNotification(
      `✓ ${newImported.length} ವಿದ್ಯಾರ್ಥಿಗಳನ್ನು ಯಶಸ್ವಿಯಾಗಿ ತರಗತಿ ${curCls} ಗೆ ಆಮದು ಮಾಡಲಾಗಿದೆ (ಕ್ರ.ಸಂ. 1 ರಿಂದ ಆರಂಭ)!`
    );
    setTimeout(() => setNotification(''), 4000);

    setShowBulkImportModal(false);
    setParsedImportStudents([]);
    setPastedText('');
    setImportStatusMsg(null);
  };

  const updateStudentField = (index: number, field: keyof Student, value: unknown) => {
    if (!verifyPrerequisites()) return;
    setStudents((prev) => {
      const copy = [...prev];
      copy[index] = { ...copy[index], [field]: value };
      return copy;
    });
  };

  const updateAttendance = (
    index: number,
    field: 's1Total' | 's2Total' | 's1Present' | 's2Present',
    val: string
  ) => {
    if (!verifyPrerequisites()) return;
    const num = parseInt(val, 10);
    const safeNum = isNaN(num) ? 0 : num;
    setStudents((prev) => {
      const copy = [...prev];
      copy[index] = {
        ...copy[index],
        attendance: { ...copy[index].attendance, [field]: safeNum },
      };
      return copy;
    });
  };

  const updateCoGrade = (studentIdx: number, type: 'co1' | 'co2', skillIdx: number, val: string) => {
    if (!verifyPrerequisites()) return;
    setStudents((prev) => {
      const copy = [...prev];
      const s = { ...copy[studentIdx] };
      const arr = [...(type === 'co1' ? s.co1 : s.co2)];
      arr[skillIdx] = val;
      if (type === 'co1') s.co1 = arr;
      else s.co2 = arr;
      copy[studentIdx] = s;
      return copy;
    });
  };

  const updateMark = (
    studentIdx: number,
    assIdx: number,
    subIdx: number,
    valStr: string
  ) => {
    if (!verifyPrerequisites()) return;
    if (CALCULATED_ASSESSMENT_INDICES.includes(assIdx)) return;

    const curCls = selectedClass || '6';
    const { maxPerSub, maxGrand } = getMaxMarksConfig(curCls);
    const maxAllowed = maxPerSub[assIdx] || 10;
    const subjects = getSubjectsForClass(curCls);
    const subName = subjects[subIdx] || 'ವಿಷಯ';
    const parts = getAssessmentPartsForClass(curCls);
    const assLabel = parts[assIdx]?.name || 'ಪರೀಕ್ಷೆ';

    if (valStr !== '' && valStr !== '-') {
      const trimmed = valStr.trim();
      if (!/^\d+$/.test(trimmed)) {
        showAlert(
          'ಅಮಾನ್ಯ ನಮೂದು!',
          `ದಯವಿಟ್ಟು ಕೇವಲ ಸಂಖ್ಯೆಗಳನ್ನು (0 ರಿಂದ ${maxAllowed}) ಮಾತ್ರ ನಮೂದಿಸಿ.`,
          maxAllowed
        );
        return;
      }

      const parsed = parseInt(trimmed, 10);
      if (parsed > maxAllowed) {
        showAlert(
          `ಗರಿಷ್ಠ ಅಂಕ ಮೀರಿದೆ! (ಗರಿಷ್ಠ: ${maxAllowed})`,
          `${assLabel} (${subName}) ಕ್ಕೆ ಗರಿಷ್ಠ ${maxAllowed} ಅಂಕಗಳನ್ನು ಮಾತ್ರ ನಮೂದಿಸಲು ಅವಕಾಶವಿದೆ. ನೀವು ನಮೂದಿಸಲು ಪ್ರಯತ್ನಿಸಿದ ಅಂಕ: ${parsed}`,
          maxAllowed
        );
        return;
      }
      if (parsed < 0) {
        showAlert(
          'ಅಮಾನ್ಯ ಅಂಕ!',
          `ಅಂಕಗಳು 0 ಕ್ಕಿಂತ ಕಡಿಮೆಯಾಗಿರಬಾರದು.`,
          maxAllowed
        );
        return;
      }
    }

    setStudents((prev) => {
      const copy = [...prev];
      const s = {
        ...copy[studentIdx],
        assessments: copy[studentIdx].assessments.map((a) => ({
          ...a,
          subs: a.subs.map((sub) => ({ ...sub })),
          total: { ...a.total },
        })),
      };

      const targetAss = s.assessments[assIdx];
      const targetSub = { ...(targetAss.subs[subIdx] || { m: '', g: '' }) };

      if (valStr === '' || valStr === '-') {
        targetSub.m = '';
        targetSub.g = '';
      } else {
        const parsed = parseInt(valStr.trim(), 10);
        targetSub.m = isNaN(parsed) ? '' : parsed;
        const validNum = isNaN(parsed) ? 0 : parsed;
        if (targetSub.m === 0) {
          const lastRule = gradeRules[gradeRules.length - 1];
          targetSub.g = lastRule ? lastRule.grade : 'C';
        } else {
          targetSub.g = getSubGrade(validNum, maxPerSub[assIdx], gradeRules);
        }
      }
      targetAss.subs[subIdx] = targetSub;

      s.assessments[assIdx] = targetAss;
      s.assessments = recalculateAssessments(s.assessments, gradeRules, curCls);

      // Automatic result determination
      const grandTotalAssessment = s.assessments[8];
      const grandSum = toInt(grandTotalAssessment?.total?.m);
      if (grandSum > 0) {
        const finalPct = calcPercentExact(grandSum, maxGrand);
        s.result = finalPct >= 35 ? 'ಉತ್ತೀರ್ಣ' : 'ಅನುತ್ತೀರ್ಣ';
      } else if (
        !s.assessments.some(
          (a) => a.subs && a.subs.some((sub) => sub.m !== '' && sub.m !== '-')
        )
      ) {
        s.result = '';
      }

      copy[studentIdx] = s;
      return copy;
    });
  };

  const filteredStudentList = React.useMemo(() => {
    if (!searchQuery.trim()) {
      return students.map((stu, originalIdx) => ({ stu, originalIdx }));
    }
    const q = searchQuery.trim().toLowerCase();
    const qClean = q.replace(/[-/.\s]/g, '');
    return students
      .map((stu, originalIdx) => ({ stu, originalIdx }))
      .filter(({ stu, originalIdx }) => {
        const idMatch = String(stu.id).includes(q) || String(originalIdx + 1).includes(q);
        const nameMatch = (stu.name || '').toLowerCase().includes(q);
        const admMatch = (stu.admission || '').toLowerCase().includes(q);
        const satsMatch = (stu.sats || '').toLowerCase().includes(q);
        const penMatch = (stu.pen || '').toLowerCase().includes(q);
        
        // Search by DOB (Raw, Formatted dd-mm-yyyy, or numbers without slashes/hyphens)
        const rawDob = (stu.dob || '').toLowerCase();
        const formattedDob = formatDateDisplay(stu.dob).toLowerCase();
        const dobMatch =
          rawDob.includes(q) ||
          formattedDob.includes(q) ||
          (qClean.length > 0 &&
            (rawDob.replace(/[-/.\s]/g, '').includes(qClean) ||
              formattedDob.replace(/[-/.\s]/g, '').includes(qClean)));

        const genderMatch = (stu.gender || '').toLowerCase().includes(q);

        return idMatch || nameMatch || admMatch || satsMatch || penMatch || dobMatch || genderMatch;
      });
  }, [students, searchQuery]);

  return (
    <div
      className="w-full min-h-screen bg-[#f3f1e8] py-2 md:py-3 print:bg-white print:p-0 relative flex flex-col items-center justify-start overflow-x-hidden"
      style={{ width: '100%', maxWidth: '100vw', paddingTop: 'var(--safe-area-inset-top, 0px)' }}
    >
      {/* Alert Modal / Warning Dialog */}
      {alertModal?.show && (
        <div
          role="alertdialog"
          aria-modal="true"
          className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-black/50 backdrop-blur-[2px] no-print animate-in fade-in duration-200"
          onClick={handleCloseAlertModal}
        >
          <div
            className="bg-white rounded-xl shadow-2xl border-2 border-red-600 max-w-md w-full p-6 text-center transform transition-all scale-100"
            onClick={(e) => e.stopPropagation()}
            style={{ fontFamily: "'Noto Sans Kannada', system-ui, sans-serif" }}
          >
            <div className="w-14 h-14 mx-auto mb-3.5 rounded-full bg-red-100 text-red-600 flex items-center justify-center shadow-inner">
              <AlertTriangle className="w-8 h-8 stroke-[2.5]" />
            </div>
            <h3 className="text-xl font-extrabold text-red-700 mb-2">
              {alertModal.title}
            </h3>
            <p className="text-[14.5px] font-bold text-gray-800 mb-4 leading-relaxed whitespace-pre-line">
              {alertModal.message}
            </p>
            {alertModal.maxAllowed !== undefined && (
              <div className="mb-5 py-2 px-3 bg-red-50 rounded-md text-[13px] font-bold text-red-800 border border-red-200 inline-block">
                {is1to5Class(selectedClass)
                  ? `ಗರಿಷ್ಠ ಮಿತಿ: ${alertModal.maxAllowed} ಅಂಕಗಳು (ರೂಪಣಾತ್ಮಕ: 15, ಸಂಕಲನಾತ್ಮಕ: 20)`
                  : `ಗರಿಷ್ಠ ಮಿತಿ: ${alertModal.maxAllowed} ಅಂಕಗಳು (ರೂಪಣಾತ್ಮಕ: 10, ಸಂಕಲನಾತ್ಮಕ: 30)`}
              </div>
            )}
            <div>
              <button
                type="button"
                autoFocus
                onClick={handleCloseAlertModal}
                className="w-full py-2.5 px-4 bg-red-600 hover:bg-red-700 active:bg-red-800 text-white font-bold rounded-lg shadow-md transition-colors cursor-pointer text-base"
              >
                ಸರಿ (OK)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Reset Confirmation Modal (ಎಲ್ಲವನ್ನೂ / ಅಂಕಗಳನ್ನು ಮರುಹೊಂದಿಸಿ) */}
      {resetModal?.show && (
        <div
          role="dialog"
          aria-modal="true"
          className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-black/60 backdrop-blur-[3px] no-print animate-in fade-in duration-200"
          onClick={() => setResetModal(null)}
        >
          <div
            className={`bg-white rounded-2xl shadow-2xl border-2 max-w-lg w-full p-6 transform transition-all scale-100 ${
              resetModal.type === 'all' ? 'border-red-600' : 'border-amber-500'
            }`}
            onClick={(e) => e.stopPropagation()}
            style={{ fontFamily: "'Noto Sans Kannada', system-ui, sans-serif" }}
          >
            {/* Header Icon & Title */}
            <div className="flex items-start gap-3.5">
              <div
                className={`w-12 h-12 rounded-xl flex items-center justify-center shrink-0 shadow-inner ${
                  resetModal.type === 'all'
                    ? 'bg-red-100 text-red-600'
                    : 'bg-amber-100 text-amber-700'
                }`}
              >
                {resetModal.type === 'all' ? (
                  <Trash2 className="w-7 h-7 stroke-[2.2]" />
                ) : (
                  <RotateCcw className="w-7 h-7 stroke-[2.2]" />
                )}
              </div>
              <div className="flex-1">
                <h3
                  className={`text-lg font-black leading-snug ${
                    resetModal.type === 'all' ? 'text-red-700' : 'text-amber-800'
                  }`}
                >
                  {resetModal.type === 'all'
                    ? '⚠️ ಎಲ್ಲವನ್ನೂ ಮರುಹೊಂದಿಸಿ (Reset All Data)'
                    : '⚠️ ಅಂಕಗಳನ್ನು ಮಾತ್ರ ಮರುಹೊಂದಿಸಿ (Reset Marks Only)'}
                </h3>
                <p className="text-[13px] font-bold text-gray-700 mt-1 leading-relaxed">
                  {resetModal.type === 'all'
                    ? 'ಶಾಲೆಯ ಹೆಸರು, ತರಗತಿ, ವರ್ಷ ಹಾಗೂ ಎಲ್ಲಾ ವಿದ್ಯಾರ್ಥಿಗಳ ವಿವರ, ಅಂಕಗಳು, ಫಲಿತಾಂಶಗಳು ಸಂಪೂರ್ಣವಾಗಿ ಅಳಿಸಲ್ಪಡುತ್ತವೆ.'
                    : 'ಎಲ್ಲಾ ವಿದ್ಯಾರ್ಥಿಗಳ ನಮೂದಿಸಲಾದ FA/SA ಅಂಕಗಳು, ಹಾಜರಾತಿ, ಸಹಪಠ್ಯ ಶ್ರೇಣಿಗಳು ಮತ್ತು ಫಲಿತಾಂಶಗಳು ಅಳಿಸಲ್ಪಡುತ್ತವೆ (ವಿದ್ಯಾರ್ಥಿ ವಿವರಗಳು ಉಳಿಯುತ್ತವೆ).'}
                </p>
              </div>
            </div>

            {/* Official Grade Rules Criteria Option (ಅಧಿಕೃತ ಶ್ರೇಣಿ ನಿಗದಿ ಮಾನದಂಡ ಕೋಷ್ಟಕ) */}
            <div className="mt-4 pt-3.5 border-t border-gray-200">
              <div className="bg-indigo-50/80 border border-indigo-200 rounded-xl p-3.5">
                <div className="flex items-center justify-between gap-2 mb-2">
                  <div className="flex items-center gap-1.5 font-black text-[13.5px] text-indigo-950">
                    <span>📊</span>
                    <span>ಅಧಿಕೃತ ಶ್ರೇಣಿ ನಿಗದಿ ಮಾನದಂಡ ಕೋಷ್ಟಕ:</span>
                  </div>
                  {isCustomRulesActive ? (
                    <span className="text-[10px] font-extrabold bg-amber-200 text-amber-900 px-2 py-0.5 rounded-full border border-amber-300">
                      ಕಸ್ಟಮೈಸ್ ಮಾಡಲಾಗಿದೆ
                    </span>
                  ) : (
                    <span className="text-[10px] font-bold bg-gray-200 text-gray-700 px-2 py-0.5 rounded-full">
                      ಮೂಲ ಡಿಫಾಲ್ಟ್
                    </span>
                  )}
                </div>

                <label className="flex items-start gap-2.5 cursor-pointer select-none group mt-1">
                  <input
                    type="checkbox"
                    checked={resetModal.resetGradeRulesToDefault}
                    onChange={(e) =>
                      setResetModal({
                        ...resetModal,
                        resetGradeRulesToDefault: e.target.checked,
                      })
                    }
                    className="mt-1 w-4 h-4 text-indigo-600 rounded border-gray-300 focus:ring-indigo-500 cursor-pointer"
                  />
                  <div className="text-[12.5px] leading-snug">
                    <span className="font-extrabold block text-gray-950 group-hover:text-indigo-900 transition-colors">
                      {resetModal.resetGradeRulesToDefault
                        ? '✔️ ಕಸ್ಟಮೈಸ್ ಮಾಡಿದ ಶ್ರೇಣಿಗಳನ್ನು ಅಳಿಸಿ ಮೂಲ ಡಿಫಾಲ್ಟ್ ಡೇಟಾ ಇರಿಸು (Erase custom & keep default)'
                        : '❌ ಈಗಿರುವ ಕಸ್ಟಮ್ ಶ್ರೇಣಿ ನಿಯಮಗಳನ್ನೇ ಹಾಗೆಯೇ ಇಟ್ಟುಕೋ (Keep customized rules)'}
                    </span>
                    <span className="text-[11.5px] text-gray-600 mt-1 block">
                      {resetModal.resetGradeRulesToDefault
                        ? 'ನೀವು ಬದಲಾಯಿಸಿದ ಕನಿಷ್ಠ-ಗರಿಷ್ಠ ಶ್ರೇಣಿ ಮೌಲ್ಯಗಳು ಅಳಿಸಿಹೋಗಿ, ಕರ್ನಾಟಕ CCE ಮೂಲ ಮಾನದಂಡಗಳು (A+ 90-100, A 70-89.9, ಇತ್ಯಾದಿ) ಮರುಸ್ಥಾಪನೆಯಾಗುತ್ತವೆ.'
                        : 'ಕೋಷ್ಟಕದಲ್ಲಿ ನೀವು ಸಿದ್ಧಪಡಿಸಿದ ಕಸ್ಟಮೈಸ್ ಶ್ರೇಣಿ ಮೌಲ್ಯಗಳು ಬದಲಾಗದೆ ಹಾಗೆಯೇ ಉಳಿಯುತ್ತವೆ.'}
                    </span>
                  </div>
                </label>
              </div>
            </div>

            {/* Warning Note */}
            <p className="mt-3 text-[11.5px] font-semibold text-rose-600 text-center">
              * ಈ ಕ್ರಿಯೆಯನ್ನು ಹಿಂಪಡೆಯಲು ಸಾಧ್ಯವಿಲ್ಲ. ಮುಂದುವರಿಯಲು ಖಚಿತಪಡಿಸಿ.
            </p>

            {/* Action Buttons */}
            <div className="mt-4 flex items-center justify-end gap-2.5">
              <button
                type="button"
                onClick={() => setResetModal(null)}
                className="px-4 py-2 text-[13px] font-bold rounded-lg border border-gray-300 text-gray-700 hover:bg-gray-100 active:bg-gray-200 transition-colors cursor-pointer"
              >
                ರದ್ದುಮಾಡಿ (Cancel)
              </button>
              <button
                type="button"
                autoFocus
                onClick={handleConfirmReset}
                className={`px-5 py-2 text-[13px] font-extrabold text-white rounded-lg shadow-md transition-all cursor-pointer flex items-center gap-1.5 ${
                  resetModal.type === 'all'
                    ? 'bg-red-600 hover:bg-red-700 active:bg-red-800'
                    : 'bg-amber-600 hover:bg-amber-700 active:bg-amber-800'
                }`}
              >
                {resetModal.type === 'all' ? (
                  <Trash2 className="w-4 h-4" />
                ) : (
                  <RotateCcw className="w-4 h-4" />
                )}
                <span>
                  {resetModal.type === 'all'
                    ? 'ಹೌದು, ಎಲ್ಲವನ್ನೂ ರಿಸೆಟ್ ಮಾಡಿ'
                    : 'ಹೌದು, ಅಂಕಗಳನ್ನು ರಿಸೆಟ್ ಮಾಡಿ'}
                </span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete Student Confirmation Modal */}
      {deleteStudentModal && (
        <div
          role="dialog"
          aria-modal="true"
          className="fixed inset-0 z-[9999] flex items-center justify-center p-4 bg-black/60 backdrop-blur-[3px] no-print animate-in fade-in duration-200"
          onClick={() => setDeleteStudentModal(null)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl border-2 border-red-500 max-w-md w-full p-6 text-center transform transition-all scale-100"
            onClick={(e) => e.stopPropagation()}
            style={{ fontFamily: "'Noto Sans Kannada', system-ui, sans-serif" }}
          >
            <div className="w-14 h-14 mx-auto mb-3.5 rounded-full bg-red-100 text-red-600 flex items-center justify-center shadow-inner">
              <Trash2 className="w-7 h-7 stroke-[2.5]" />
            </div>
            <h3 className="text-xl font-extrabold text-red-700 mb-2">
              ವಿದ್ಯಾರ್ಥಿ ತೆಗೆದುಹಾಕಿ (Delete Student)
            </h3>
            <p className="text-[14.5px] font-bold text-gray-800 mb-2 leading-relaxed">
              ಖಂಡಿತವಾಗಿ <span className="text-red-700 underline font-extrabold">{deleteStudentModal.name}</span> ಅವರ ವಿವರಗಳನ್ನು ಮತ್ತು ಎಲ್ಲಾ ಅಂಕಗಳನ್ನು ಕೋಷ್ಟಕದಿಂದ ತೆಗೆದುಹಾಕಬೇಕೆ?
            </p>
            <p className="text-[12px] text-gray-500 mb-5">
              (ತಪ್ಪಾಗಿ ಸೇರಿಸಲಾಗಿದ್ದರೆ ಈ ವಿದ್ಯಾರ್ಥಿಯನ್ನು ಕೋಷ್ಟಕದಿಂದ ಶಾಶ್ವತವಾಗಿ ತೆಗೆದುಹಾಕಲಾಗುತ್ತದೆ)
            </p>
            <div className="flex items-center justify-center gap-3">
              <button
                type="button"
                onClick={() => setDeleteStudentModal(null)}
                className="flex-1 py-2.5 px-4 bg-gray-100 hover:bg-gray-200 text-gray-800 font-bold rounded-lg transition-colors cursor-pointer text-sm"
              >
                ರದ್ದುಮಾಡಿ (Cancel)
              </button>
              <button
                type="button"
                autoFocus
                onClick={confirmDeleteStudent}
                className="flex-1 py-2.5 px-4 bg-red-600 hover:bg-red-700 active:bg-red-800 text-white font-extrabold rounded-lg shadow-md transition-colors cursor-pointer text-sm flex items-center justify-center gap-1.5"
              >
                <Trash2 className="w-4 h-4" />
                <span>ತೆಗೆದುಹಾಕಿ (Delete)</span>
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Bulk Student Import Modal (ವಿದ್ಯಾರ್ಥಿಗಳ ಬಲ್ಕ್ ಆಮದು) */}
      {showBulkImportModal && (
        <div
          role="dialog"
          aria-modal="true"
          className="fixed inset-0 z-[9999] flex items-center justify-center p-3 md:p-6 bg-black/60 backdrop-blur-[3px] no-print animate-in fade-in duration-200 overflow-y-auto"
          onClick={() => setShowBulkImportModal(false)}
        >
          <div
            className="bg-white rounded-2xl shadow-2xl border-2 border-sky-600 max-w-3xl w-full p-4 md:p-6 text-left transform transition-all my-auto max-h-[92vh] overflow-y-auto"
            onClick={(e) => e.stopPropagation()}
            style={{ fontFamily: "'Noto Sans Kannada', system-ui, sans-serif" }}
          >
            {/* Header */}
            <div className="flex items-center justify-between pb-3 border-b border-gray-200 mb-4">
              <div className="flex items-center gap-2 text-sky-800">
                <Upload className="w-6 h-6 stroke-[2.5]" />
                <h3 className="text-lg md:text-xl font-extrabold">
                  📥 ವಿದ್ಯಾರ್ಥಿಗಳ ಬಲ್ಕ್ ಆಮದು (Bulk Import Students - Class {selectedClass || '1-9'})
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setShowBulkImportModal(false)}
                className="text-gray-400 hover:text-gray-700 hover:bg-gray-100 p-1 rounded-lg transition-colors cursor-pointer"
                title="ಮುಚ್ಚಿ"
              >
                <X className="w-6 h-6" />
              </button>
            </div>

            {/* Template Download Option */}
            <div className="mb-4 p-3 bg-sky-50 border border-sky-200 rounded-xl flex flex-wrap items-center justify-between gap-2">
              <div>
                <p className="text-xs md:text-sm font-bold text-sky-950">
                  💡 Excel ಟೆಂಪ್ಲೇಟ್ ಬಳಸಿ ಸುಲಭವಾಗಿ ಆಮದು ಮಾಡಿ:
                </p>
                <p className="text-[11px] text-sky-800 font-medium">
                  ಮಾದರಿ ಫೈಲ್ ಡೌನ್‌ಲೋಡ್ ಮಾಡಿ, ವಿದ್ಯಾರ್ಥಿ ವಿವರಗಳನ್ನು ಭರ್ತಿ ಮಾಡಿ ನಂತರ ಇಲ್ಲಿ ಅಪ್‌ಲೋಡ್ ಮಾಡಿ.
                </p>
              </div>
              <button
                type="button"
                onClick={handleDownloadImportTemplate}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-sky-700 hover:bg-sky-800 text-white font-bold text-xs rounded-lg shadow transition-colors cursor-pointer"
              >
                <Download className="w-4 h-4" />
                <span>ಮಾದರಿ ಟೆಂಪ್ಲೇಟ್ ಡೌನ್‌ಲೋಡ್ (Sample XLSX)</span>
              </button>
            </div>

            {/* Import Method Tabs */}
            <div className="flex items-center gap-2 border-b border-gray-200 mb-4">
              <button
                type="button"
                onClick={() => {
                  setImportTab('file');
                  setImportStatusMsg(null);
                }}
                className={`py-2 px-4 text-xs md:text-sm font-extrabold border-b-2 transition-colors cursor-pointer flex items-center gap-1.5 ${
                  importTab === 'file'
                    ? 'border-sky-600 text-sky-800 bg-sky-50/50 rounded-t-lg'
                    : 'border-transparent text-gray-500 hover:text-gray-800'
                }`}
              >
                <FileSpreadsheet className="w-4 h-4" />
                <span>1. Excel / CSV ಫೈಲ್ ಆಯ್ಕೆ</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setImportTab('text');
                  setImportStatusMsg(null);
                }}
                className={`py-2 px-4 text-xs md:text-sm font-extrabold border-b-2 transition-colors cursor-pointer flex items-center gap-1.5 ${
                  importTab === 'text'
                    ? 'border-sky-600 text-sky-800 bg-sky-50/50 rounded-t-lg'
                    : 'border-transparent text-gray-500 hover:text-gray-800'
                }`}
              >
                <ClipboardList className="w-4 h-4" />
                <span>2. ಪಠ್ಯ ಪೇಸ್ಟ್ ಮಾಡಿ (Copy-Paste)</span>
              </button>
            </div>

            {/* Tab 1: File Upload */}
            {importTab === 'file' && (
              <div className="space-y-3 mb-4">
                <label className="block p-5 border-2 border-dashed border-sky-300 hover:border-sky-500 rounded-xl bg-sky-50/30 hover:bg-sky-50/70 text-center cursor-pointer transition-all">
                  <Upload className="w-8 h-8 mx-auto text-sky-600 mb-2" />
                  <span className="text-sm font-bold text-gray-800 block">
                    ಇಲ್ಲಿ ಕ್ಲಿಕ್ ಮಾಡಿ ನಿಮ್ಮ Excel (.xlsx, .xls) ಅಥವಾ CSV ಫೈಲ್ ಆಯ್ಕೆಮಾಡಿ
                  </span>
                  <span className="text-xs text-gray-500 block mt-1">
                    (ಕ್ರಮ ಸಂಖ್ಯೆ, ವಿದ್ಯಾರ್ಥಿ ಹೆಸರು, ದಾಖಲಾತಿ, ಲಿಂಗ, ಜನ್ಮ ದಿನಾಂಕ, SATS, PEN ಅಂಕಣಗಳು ಸ್ವಯಂಚಾಲಿತವಾಗಿ ಗುರುತಿಸಲ್ಪಡುತ್ತವೆ)
                  </span>
                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="hidden"
                    onChange={handleBulkFileUpload}
                  />
                </label>
              </div>
            )}

            {/* Tab 2: Copy Paste Text */}
            {importTab === 'text' && (
              <div className="space-y-3 mb-4">
                <label className="block text-xs font-bold text-gray-700">
                  ವಿದ್ಯಾರ್ಥಿಗಳ ಪಟ್ಟಿಯನ್ನು ಇಲ್ಲಿ ಪೇಸ್ಟ್ ಮಾಡಿ (Excel, Word, WhatsApp ಅಥವಾ SATS ನಿಂದ ಕಾಪಿ ಮಾಡಿದ ಪಟ್ಟಿ):
                </label>
                <textarea
                  rows={5}
                  value={pastedText}
                  onChange={(e) => setPastedText(e.target.value)}
                  placeholder={`ಉದಾಹರಣೆಗೆ (ಪ್ರತಿ ಸಾಲಿಗೆ ಒಬ್ಬ ವಿದ್ಯಾರ್ಥಿ):
ರಮೇಶ್ ಬಿ 	 101 	 M 	 2014-05-15 	 123456789
ಕಾವ್ಯ ಎಂ 	 102 	 F 	 2014-08-20 	 987654321`}
                  className="w-full p-2.5 text-xs font-mono border border-gray-300 rounded-lg focus:ring-2 focus:ring-sky-500 focus:outline-none"
                />
                <button
                  type="button"
                  onClick={handleParsePastedText}
                  className="px-4 py-2 bg-sky-700 hover:bg-sky-800 text-white font-bold text-xs rounded-lg shadow transition-colors cursor-pointer flex items-center gap-1.5"
                >
                  <FileText className="w-4 h-4" />
                  <span>ಪಠ್ಯ ವಿಶ್ಲೇಷಿಸಿ (Parse List)</span>
                </button>
              </div>
            )}

            {/* Status Message */}
            {importStatusMsg && (
              <div
                className={`p-3 rounded-lg text-xs font-bold mb-4 flex items-center gap-2 ${
                  importStatusMsg.type === 'success'
                    ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                    : importStatusMsg.type === 'error'
                    ? 'bg-red-50 text-red-800 border border-red-200'
                    : 'bg-sky-50 text-sky-800 border border-sky-200'
                }`}
              >
                <span>{importStatusMsg.text}</span>
              </div>
            )}

            {/* Parsed Students Preview Table */}
            {parsedImportStudents.length > 0 && (
              <div className="mb-5 space-y-3">
                <div className="flex items-center justify-between">
                  <h4 className="text-xs font-black text-gray-900">
                    📋 ಗುರುತಿಸಲಾದ ವಿದ್ಯಾರ್ಥಿಗಳ ಪೂರ್ವವೀಕ್ಷಣೆ ({parsedImportStudents.length} ವಿದ್ಯಾರ್ಥಿಗಳು):
                  </h4>
                  <div className="flex items-center gap-4 text-xs font-bold text-gray-700">
                    <label className="inline-flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="radio"
                        name="importMode"
                        value="append"
                        checked={importMode === 'append'}
                        onChange={() => setImportMode('append')}
                        className="text-sky-600 focus:ring-sky-500"
                      />
                      <span>ಹಾಲಿ ಪಟ್ಟಿಗೆ ಸೇರಿಸಿ (Append)</span>
                    </label>
                    <label className="inline-flex items-center gap-1.5 cursor-pointer">
                      <input
                        type="radio"
                        name="importMode"
                        value="replace"
                        checked={importMode === 'replace'}
                        onChange={() => setImportMode('replace')}
                        className="text-sky-600 focus:ring-sky-500"
                      />
                      <span className="text-red-700 font-extrabold">ಹಾಲಿ ಪಟ್ಟಿಯನ್ನು ಬದಲಾಯಿಸಿ (Replace)</span>
                    </label>
                  </div>
                </div>

                <div className="max-h-[220px] overflow-y-auto border border-gray-300 rounded-lg shadow-inner">
                  <table className="w-full text-xs text-left border-collapse">
                    <thead className="bg-sky-100 text-sky-950 sticky top-0 font-extrabold border-b border-sky-300">
                      <tr>
                        <th className="p-2 border-r border-sky-200 w-10 text-center">#</th>
                        <th className="p-2 border-r border-sky-200">ವಿದ್ಯಾರ್ಥಿ ಹೆಸರು</th>
                        <th className="p-2 border-r border-sky-200">ದಾಖಲಾತಿ</th>
                        <th className="p-2 border-r border-sky-200">ಲಿಂಗ</th>
                        <th className="p-2 border-r border-sky-200">DOB</th>
                        <th className="p-2 border-r border-sky-200">SATS No</th>
                        <th className="p-2">PEN No</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200 bg-white">
                      {parsedImportStudents.map((st, idx) => (
                        <tr key={idx} className="hover:bg-sky-50/50 transition-colors">
                          <td className="p-1.5 text-center font-bold text-gray-500 border-r border-gray-200">
                            {idx + 1}
                          </td>
                          <td className="p-1.5 font-bold text-gray-900 border-r border-gray-200">
                            {st.name || '-'}
                          </td>
                          <td className="p-1.5 font-medium text-gray-700 border-r border-gray-200">
                            {st.admission || '-'}
                          </td>
                          <td className="p-1.5 font-medium text-gray-700 border-r border-gray-200">
                            {st.gender || '-'}
                          </td>
                          <td className="p-1.5 font-medium text-gray-700 border-r border-gray-200">
                            {st.dob || '-'}
                          </td>
                          <td className="p-1.5 font-medium text-gray-700 border-r border-gray-200">
                            {st.sats || '-'}
                          </td>
                          <td className="p-1.5 font-medium text-gray-700">
                            {st.pen || '-'}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Footer Modal Actions */}
            <div className="flex items-center justify-end gap-3 pt-3 border-t border-gray-200">
              <button
                type="button"
                onClick={() => {
                  setShowBulkImportModal(false);
                  setParsedImportStudents([]);
                  setImportStatusMsg(null);
                }}
                className="py-2 px-4 bg-gray-100 hover:bg-gray-200 text-gray-800 font-bold rounded-lg text-xs transition-colors cursor-pointer"
              >
                ರದ್ದುಮಾಡಿ (Cancel)
              </button>
              <button
                type="button"
                disabled={parsedImportStudents.length === 0}
                onClick={handleConfirmBulkImport}
                className={`py-2 px-5 font-extrabold rounded-lg text-xs shadow-md transition-all cursor-pointer flex items-center gap-1.5 ${
                  parsedImportStudents.length > 0
                    ? 'bg-sky-700 hover:bg-sky-800 active:bg-sky-900 text-white'
                    : 'bg-gray-300 text-gray-500 cursor-not-allowed'
                }`}
              >
                <Check className="w-4 h-4" />
                <span>
                  ✓ {parsedImportStudents.length} ವಿದ್ಯಾರ್ಥಿಗಳನ್ನು ಆಮದು ಮಾಡಿ (Confirm Import)
                </span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Grade Rules Modal (ಕರ್ನಾಟಕ CCE ಶ್ರೇಣಿ ನಿಯಮಗಳ ಕೋಷ್ಟಕ) */}
      {showGradeModal && (
        <div
          role="dialog"
          aria-modal="true"
          className="fixed inset-0 z-[9999] flex items-center justify-center p-3 md:p-6 bg-black/60 backdrop-blur-[2px] no-print animate-in fade-in duration-200 overflow-y-auto"
          onClick={() => setShowGradeModal(false)}
        >
          <div
            className="bg-white rounded-xl shadow-2xl border-2 border-indigo-600 max-w-4xl w-full p-4 md:p-6 text-left transform transition-all my-auto max-h-[92vh] overflow-y-auto"
            onClick={(e) => e.stopPropagation()}
            style={{ fontFamily: "'Noto Sans Kannada', system-ui, sans-serif" }}
          >
            <div className="flex items-center justify-between pb-3 border-b border-gray-200 mb-4">
              <div className="flex items-center gap-2 text-indigo-700">
                <Award className="w-6 h-6 stroke-[2.5]" />
                <h3 className="text-lg md:text-xl font-extrabold">
                  ಕರ್ನಾಟಕ CCE ಶ್ರೇಣಿ ನಿಯಮಗಳ ಕೋಷ್ಟಕ (DSERT {is1to5Class(selectedClass) ? '1 ರಿಂದ 5 ನೇ ತರಗತಿ' : '6 ರಿಂದ 9 ನೇ ತರಗತಿ'})
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setShowGradeModal(false)}
                className="w-8 h-8 rounded-full bg-gray-100 hover:bg-gray-200 text-gray-700 font-bold flex items-center justify-center transition-colors cursor-pointer text-lg"
                title="ಮುಚ್ಚಿ"
              >
                ✕
              </button>
            </div>

            {/* Official Primary Grade Rules Table */}
            <div className="mb-4 overflow-hidden rounded-lg border-2 border-indigo-200 shadow-sm bg-white">
              <div className="bg-gradient-to-r from-indigo-700 via-indigo-600 to-purple-700 px-4 py-3 text-white flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="font-extrabold text-[14.5px] tracking-wide flex items-center gap-1.5">
                    📋 ಅಧಿಕೃತ ಶ್ರೇಣಿ ನಿಗದಿ ಮಾನದಂಡ ಕೋಷ್ಟಕ
                  </span>
                  <span className="text-[11px] bg-white/20 px-2.5 py-0.5 rounded-full font-extrabold text-amber-200 border border-amber-300/40">
                    ತರಗತಿ: {selectedClass ? `${selectedClass} ನೇ (${is1to5Class(selectedClass) ? '1-5' : '6-9'})` : '6-9'}
                  </span>
                  {isCustomRulesActive ? (
                    <span className="text-[11px] bg-emerald-500/30 text-emerald-100 border border-emerald-400/50 px-2.5 py-0.5 rounded-full font-bold flex items-center gap-1 shadow-xs">
                      ✨ ಕಸ್ಟಮ್ ಶ್ರೇಣಿಗಳು ಲಿಂಕ್ ಆಗಿವೆ (Linked)
                    </span>
                  ) : (
                    <span className="text-[11px] bg-white/20 px-2 py-0.5 rounded font-semibold">
                      CCE ಅಧಿಕೃತ ಮಾನದಂಡ
                    </span>
                  )}
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                  {!isEditingGradeRules ? (
                    <>
                      {isCustomRulesActive && (
                        <button
                          type="button"
                          onClick={handleResetGradeRules}
                          className="px-2.5 py-1 text-[12px] font-bold rounded bg-white/20 hover:bg-white/30 text-white flex items-center gap-1 transition-colors cursor-pointer"
                          title="ಮೂಲ ಡಿಫಾಲ್ಟ್ ನಿಯಮಗಳಿಗೆ ಮರುಹೊಂದಿಸಿ"
                        >
                          <RotateCcw className="w-3.5 h-3.5" />
                          <span>ಡಿಫಾಲ್ಟ್‌ಗೆ ಮರುಹೊಂದಿಸಿ</span>
                        </button>
                      )}
                      <button
                        type="button"
                        onClick={handleStartEditRules}
                        className="px-3 py-1.5 text-[12.5px] font-extrabold rounded-md bg-amber-400 hover:bg-amber-300 active:bg-amber-500 text-gray-950 flex items-center gap-1.5 shadow-sm transition-all cursor-pointer"
                        title="ಶ್ರೇಣಿ ಮತ್ತು ಅಂಕಗಳ ವ್ಯಾಪ್ತಿಯನ್ನು ಸಂಪಾದಿಸಿ"
                      >
                        <Edit3 className="w-4 h-4 stroke-[2.5]" />
                        <span>✏️ ಶ್ರೇಣಿ & ವ್ಯಾಪ್ತಿ ಸಂಪಾದಿಸಿ (Edit)</span>
                      </button>
                    </>
                  ) : (
                    <>
                      <button
                        type="button"
                        onClick={handleAddRuleRow}
                        className="px-2.5 py-1.5 text-[12px] font-bold rounded-md bg-white/20 hover:bg-white/30 text-white flex items-center gap-1 transition-colors cursor-pointer"
                        title="ಹೊಸ ಶ್ರೇಣಿ ಸಾಲು ಸೇರಿಸಿ"
                      >
                        <Plus className="w-3.5 h-3.5 stroke-[2.5]" />
                        <span>ಸಾಲು ಸೇರಿಸಿ</span>
                      </button>
                      <button
                        type="button"
                        onClick={handleResetGradeRules}
                        className="px-2.5 py-1.5 text-[12px] font-bold rounded-md bg-red-500/80 hover:bg-red-600 text-white flex items-center gap-1 transition-colors cursor-pointer"
                        title="ಮೂಲ ಡಿಫಾಲ್ಟ್ ನಿಯಮಗಳಿಗೆ ಮರುಹೊಂದಿಸಿ"
                      >
                        <RotateCcw className="w-3.5 h-3.5" />
                        <span>ಡಿಫಾಲ್ಟ್‌</span>
                      </button>
                      <button
                        type="button"
                        onClick={() => setIsEditingGradeRules(false)}
                        className="px-2.5 py-1.5 text-[12px] font-bold rounded-md bg-gray-600 hover:bg-gray-500 text-white flex items-center gap-1 transition-colors cursor-pointer"
                      >
                        <X className="w-3.5 h-3.5 stroke-[2.5]" />
                        <span>ರದ್ದು</span>
                      </button>
                      <button
                        type="button"
                        onClick={handleSaveGradeRules}
                        className="px-3.5 py-1.5 text-[12.5px] font-extrabold rounded-md bg-emerald-500 hover:bg-emerald-400 active:bg-emerald-600 text-white flex items-center gap-1.5 shadow-md transition-all cursor-pointer"
                        title="ಉಳಿಸಿ ಮತ್ತು ಮುಖ್ಯ ಕೋಷ್ಟಕಕ್ಕೆ ಲಿಂಕ್ ಮಾಡಿ"
                      >
                        <Save className="w-4 h-4 stroke-[2.5]" />
                        <span>💾 ಉಳಿಸಿ & ಲಿಂಕ್ ಮಾಡಿ</span>
                      </button>
                    </>
                  )}
                </div>
              </div>

              {/* Instructions banner when editing */}
              {isEditingGradeRules && (
                <div className="bg-amber-50 border-b border-amber-200 px-4 py-2.5 text-[12.5px] text-amber-900 flex items-start gap-2">
                  <span className="text-base leading-none">💡</span>
                  <div>
                    <span className="font-extrabold">ಸಂಪಾದನೆ ಸೂಚನೆ:</span> ಇಲ್ಲಿ ನೀವು ಶ್ರೇಣಿಯ ಹೆಸರು ಹಾಗೂ ಪ್ರತಿ ವಿಭಾಗದ ({is1to5Class(selectedClass) ? '15%, 20%' : '10%, 30%'}, 50%, 100%) ಕೇವಲ <strong className="text-indigo-900">ಕನಿಷ್ಠ (Min)</strong> ಮತ್ತು <strong className="text-indigo-900">ಗರಿಷ್ಠ (Max)</strong> ಬಾಕ್ಸ್‌ಗಳಲ್ಲಿ ಅಂಕಗಳನ್ನು ಟೈಪ್ ಮಾಡಿ. ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ ತಂತಾನೇ ತಯಾರಾಗುತ್ತದೆ (ಉದಾ: ಕನಿಷ್ಠ 0 ಹಾಗೂ ಗರಿಷ್ಠ 2.9 ನೀಡಿದಾಗ <strong>0 - 2.9</strong> ರೇಂಜ್ ಆಗುತ್ತದೆ).{' '}
                    <strong className="text-emerald-800">"ಉಳಿಸಿ & ಲಿಂಕ್ ಮಾಡಿ"</strong> ಕ್ಲಿಕ್ ಮಾಡಿದ ತಕ್ಷಣ ಮುಖ್ಯ ರಿಜಿಸ್ಟರ್‌ನ ಎಲ್ಲಾ ವಿದ್ಯಾರ್ಥಿಗಳ ಫಲಿತಾಂಶಗಳು ತಕ್ಷಣವೇ ಮರು-ಲೆಕ್ಕಾಚಾರವಾಗುತ್ತವೆ!
                  </div>
                </div>
              )}

              {/* Table rendering: Edit mode vs View mode */}
              <div className="overflow-x-auto">
                {isEditingGradeRules ? (
                  <table className="w-full text-center border-collapse text-[13px]">
                    <thead>
                      <tr className="bg-indigo-50/90 border-b-2 border-indigo-200 text-indigo-950 font-extrabold text-[13px]">
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[12%]">
                          ಶ್ರೇಣಿ (Grade)
                        </th>
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[21%]">
                          {is1to5Class(selectedClass) ? '15% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (FA)' : '10% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (FA)'}
                          <div className="text-[11px] font-normal text-indigo-700">ಕನಿಷ್ಠ — ಗರಿಷ್ಠ</div>
                        </th>
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[21%]">
                          {is1to5Class(selectedClass) ? '20% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (SA)' : '30% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (SA)'}
                          <div className="text-[11px] font-normal text-indigo-700">ಕನಿಷ್ಠ — ಗರಿಷ್ಠ</div>
                        </th>
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[21%]">
                          50% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (Sem)
                          <div className="text-[11px] font-normal text-indigo-700">ಕನಿಷ್ಠ — ಗರಿಷ್ಠ</div>
                        </th>
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[21%]">
                          100% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (Total)
                          <div className="text-[11px] font-normal text-indigo-700">ಕನಿಷ್ಠ % — ಗರಿಷ್ಠ %</div>
                        </th>
                        <th className="py-2.5 px-2 text-center w-[4%]">
                          ಕ್ರಿಯೆ
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200 bg-white">
                      {editableRules.map((rule, idx) => {
                        const is1to5 = is1to5Class(selectedClass);
                        const minFA = is1to5 ? rule.min15 : rule.min10;
                        const maxFA = is1to5 ? rule.max15 : rule.max10;
                        const fieldFA = is1to5 ? '15' : '10';

                        const minSA = is1to5 ? rule.min20 : rule.min30;
                        const maxSA = is1to5 ? rule.max20 : rule.max30;
                        const fieldSA = is1to5 ? '20' : '30';

                        return (
                          <tr key={idx} className="hover:bg-indigo-50/30 transition-colors">
                            {/* Grade Name */}
                            <td className="py-2.5 px-2 border-r border-gray-200">
                              <div className="flex items-center justify-center gap-1.5">
                                <span className={`inline-block px-2 py-0.5 rounded font-extrabold text-[12px] shadow-xs ${getGradeBadgeClass(rule.grade)}`}>
                                  {rule.grade || '—'}
                                </span>
                                <input
                                  type="text"
                                  value={rule.grade}
                                  onChange={(e) => handleUpdateGradeName(idx, e.target.value)}
                                  className="w-16 text-center font-extrabold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                  placeholder="ಶ್ರೇಣಿ"
                                />
                              </div>
                            </td>

                            {/* FA Range (Min & Max Boxes: 15% for 1-5, 10% for 6-9) */}
                            <td className="py-2.5 px-2 border-r border-gray-200">
                              <div className="mb-1 text-[11px] font-bold text-indigo-700 bg-indigo-50/80 rounded px-2 py-0.5 inline-block border border-indigo-100">
                                ವ್ಯಾಪ್ತಿ: {minFA} - {maxFA}
                              </div>
                              <div className="flex items-start justify-center gap-1">
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={minFA}
                                    onChange={(e) => handleUpdateMinMax(idx, fieldFA, 'min', e.target.value)}
                                    className="w-16 text-center font-bold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                    placeholder="0"
                                  />
                                  <span className="text-[10.5px] font-semibold text-gray-500 mt-0.5">ಕನಿಷ್ಠ</span>
                                </div>
                                <span className="font-bold text-gray-400 mt-1">—</span>
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={maxFA}
                                    onChange={(e) => handleUpdateMinMax(idx, fieldFA, 'max', e.target.value)}
                                    className="w-16 text-center font-bold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                    placeholder={is1to5 ? '15' : '10'}
                                  />
                                  <span className="text-[10.5px] font-semibold text-gray-500 mt-0.5">ಗರಿಷ್ಠ</span>
                                </div>
                              </div>
                            </td>

                            {/* SA Range (Min & Max Boxes: 20% for 1-5, 30% for 6-9) */}
                            <td className="py-2.5 px-2 border-r border-gray-200">
                              <div className="mb-1 text-[11px] font-bold text-indigo-700 bg-indigo-50/80 rounded px-2 py-0.5 inline-block border border-indigo-100">
                                ವ್ಯಾಪ್ತಿ: {minSA} - {maxSA}
                              </div>
                              <div className="flex items-start justify-center gap-1">
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={minSA}
                                    onChange={(e) => handleUpdateMinMax(idx, fieldSA, 'min', e.target.value)}
                                    className="w-16 text-center font-bold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                    placeholder="0"
                                  />
                                  <span className="text-[10.5px] font-semibold text-gray-500 mt-0.5">ಕನಿಷ್ಠ</span>
                                </div>
                                <span className="font-bold text-gray-400 mt-1">—</span>
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={maxSA}
                                    onChange={(e) => handleUpdateMinMax(idx, fieldSA, 'max', e.target.value)}
                                    className="w-16 text-center font-bold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                    placeholder={is1to5 ? '20' : '30'}
                                  />
                                  <span className="text-[10.5px] font-semibold text-gray-500 mt-0.5">ಗರಿಷ್ಠ</span>
                                </div>
                              </div>
                            </td>

                            {/* 50% Range (Min & Max Boxes) */}
                            <td className="py-2.5 px-2 border-r border-gray-200">
                              <div className="mb-1 text-[11px] font-bold text-indigo-700 bg-indigo-50/80 rounded px-2 py-0.5 inline-block border border-indigo-100">
                                ವ್ಯಾಪ್ತಿ: {rule.min50} - {rule.max50}
                              </div>
                              <div className="flex items-start justify-center gap-1">
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={rule.min50}
                                    onChange={(e) => handleUpdateMinMax(idx, '50', 'min', e.target.value)}
                                    className="w-16 text-center font-bold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                    placeholder="0"
                                  />
                                  <span className="text-[10.5px] font-semibold text-gray-500 mt-0.5">ಕನಿಷ್ಠ</span>
                                </div>
                                <span className="font-bold text-gray-400 mt-1">—</span>
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={rule.max50}
                                    onChange={(e) => handleUpdateMinMax(idx, '50', 'max', e.target.value)}
                                    className="w-16 text-center font-bold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                    placeholder="50"
                                  />
                                  <span className="text-[10.5px] font-semibold text-gray-500 mt-0.5">ಗರಿಷ್ಠ</span>
                                </div>
                              </div>
                            </td>

                            {/* 100% Range (Min & Max Boxes) */}
                            <td className="py-2.5 px-2 border-r border-gray-200">
                              <div className="mb-1 text-[11px] font-bold text-indigo-700 bg-indigo-50/80 rounded px-2 py-0.5 inline-block border border-indigo-100">
                                ವ್ಯಾಪ್ತಿ: {rule.min100} - {rule.max100}%
                              </div>
                              <div className="flex items-start justify-center gap-1">
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={rule.min100}
                                    onChange={(e) => handleUpdateMinMax(idx, '100', 'min', e.target.value)}
                                    className="w-16 text-center font-bold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                    placeholder="0"
                                  />
                                  <span className="text-[10.5px] font-semibold text-gray-500 mt-0.5">ಕನಿಷ್ಠ %</span>
                                </div>
                                <span className="font-bold text-gray-400 mt-1">—</span>
                                <div className="flex flex-col items-center">
                                  <input
                                    type="number"
                                    step="0.1"
                                    value={rule.max100}
                                    onChange={(e) => handleUpdateMinMax(idx, '100', 'max', e.target.value)}
                                    className="w-16 text-center font-bold border-2 border-indigo-200 rounded px-1.5 py-1 text-sm bg-white focus:outline-none focus:border-indigo-600 focus:ring-1 focus:ring-indigo-500"
                                    placeholder="100"
                                  />
                                  <span className="text-[10.5px] font-semibold text-gray-500 mt-0.5">ಗರಿಷ್ಠ %</span>
                                </div>
                              </div>
                            </td>

                            {/* Action (Delete) */}
                            <td className="py-2.5 px-2 text-center">
                              <button
                                type="button"
                                disabled={editableRules.length <= 1}
                                onClick={() => handleRemoveRuleRow(idx)}
                                className="p-1.5 rounded text-red-500 hover:text-red-700 hover:bg-red-50 disabled:opacity-30 disabled:cursor-not-allowed transition-colors"
                                title="ಸಾಲು ತೆಗೆದುಹಾಕಿ"
                              >
                                <Trash2 className="w-4 h-4 mx-auto" />
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                ) : (
                  <table className="w-full text-center border-collapse text-[13px]">
                    <thead>
                      <tr className="bg-indigo-50/80 border-b-2 border-indigo-200 text-indigo-950 font-extrabold text-[13px]">
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[12%]">
                          ಶ್ರೇಣಿ
                        </th>
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[22%]">
                          {is1to5Class(selectedClass) ? '15% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (FA)' : '10% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (FA)'}
                        </th>
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[22%]">
                          {is1to5Class(selectedClass) ? '20% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (SA)' : '30% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (SA)'}
                        </th>
                        <th className="py-2.5 px-3 border-r border-indigo-200 text-center w-[22%]">
                          50% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (Sem)
                        </th>
                        <th className="py-2.5 px-3 text-center w-[22%]">
                          100% ಕ್ಕೆ ಅಂಕಗಳ ವ್ಯಾಪ್ತಿ (Total)
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200">
                      {gradeRules.map((rule, idx) => {
                        const is1to5 = is1to5Class(selectedClass);
                        const textFA = is1to5 ? rule.text15 : rule.text10;
                        const textSA = is1to5 ? rule.text20 : rule.text30;

                        return (
                          <tr
                            key={idx}
                            className={`hover:bg-indigo-50/40 transition-colors ${
                              idx === gradeRules.length - 1 ? 'bg-red-50/20' : ''
                            }`}
                          >
                            <td className="py-2.5 px-3 border-r border-gray-200 font-extrabold">
                              <span
                                className={`inline-block px-3 py-1 rounded-md font-extrabold text-[13px] shadow-xs ${getGradeBadgeClass(
                                  rule.grade
                                )}`}
                              >
                                {rule.grade}
                              </span>
                            </td>
                            <td
                              className={`py-2.5 px-3 border-r border-gray-200 font-bold text-[14px] ${
                                idx === gradeRules.length - 1 ? 'text-rose-700' : 'text-gray-900'
                              }`}
                            >
                              {textFA}
                            </td>
                            <td
                              className={`py-2.5 px-3 border-r border-gray-200 font-bold text-[14px] ${
                                idx === gradeRules.length - 1 ? 'text-rose-700' : 'text-gray-900'
                              }`}
                            >
                              {textSA}
                            </td>
                            <td
                              className={`py-2.5 px-3 border-r border-gray-200 font-bold text-[14px] ${
                                idx === gradeRules.length - 1 ? 'text-rose-700' : 'text-gray-900'
                              }`}
                            >
                              {rule.text50}
                            </td>
                            <td
                              className={`py-2.5 px-3 font-bold text-[14px] ${
                                idx === gradeRules.length - 1 ? 'text-rose-700' : 'text-gray-900'
                              }`}
                            >
                              {rule.text100}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )}
              </div>

              {/* Edit mode bottom action toolbar */}
              {isEditingGradeRules && (
                <div className="bg-indigo-50/80 px-4 py-3 border-t border-indigo-200 flex flex-wrap items-center justify-between gap-2">
                  <button
                    type="button"
                    onClick={handleAddRuleRow}
                    className="px-3 py-1.5 rounded-lg border-2 border-dashed border-indigo-400 hover:border-indigo-600 bg-white hover:bg-indigo-50 text-indigo-700 font-extrabold text-[13px] flex items-center gap-1.5 transition-all cursor-pointer shadow-xs"
                  >
                    <Plus className="w-4 h-4 stroke-[2.5]" />
                    <span>+ ಮತ್ತೊಂದು ಶ್ರೇಣಿ ಸೇರಿಸಿ (Add Grade Row)</span>
                  </button>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setIsEditingGradeRules(false)}
                      className="px-4 py-2 rounded-lg bg-gray-200 hover:bg-gray-300 text-gray-800 font-bold text-[13px] transition-colors cursor-pointer"
                    >
                      ರದ್ದುಗೊಳಿಸಿ (Cancel)
                    </button>
                    <button
                      type="button"
                      onClick={handleSaveGradeRules}
                      className="px-5 py-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white font-extrabold text-[13px] flex items-center gap-2 shadow-md transition-all cursor-pointer"
                    >
                      <Save className="w-4 h-4 stroke-[2.5]" />
                      <span>💾 ಉಳಿಸಿ ಮತ್ತು ಮುಖ್ಯ ಕೋಷ್ಟಕಕ್ಕೆ ಲಿಂಕ್ ಮಾಡಿ (Save & Link)</span>
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Assessment Cards Overview */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
              {/* Card 1: FA Marks */}
              <div className="bg-blue-50/70 border border-blue-200 rounded-lg p-3 shadow-xs">
                <div className="font-extrabold text-blue-900 text-[13px] border-b border-blue-200 pb-1.5 mb-2 text-center">
                  1. ರೂಪಣಾತ್ಮಕ (FA1 - FA4)
                  <div className="text-[11.5px] font-semibold text-blue-700">
                    {is1to5Class(selectedClass) ? '15% ಕ್ಕೆ (ಗರಿಷ್ಠ 15 ಅಂಕ)' : '10% ಕ್ಕೆ (ಗರಿಷ್ಠ 10 ಅಂಕ)'}
                  </div>
                </div>
                <div className="space-y-1 text-[12.5px]">
                  {gradeRules.map((r, i) => (
                    <div
                      key={i}
                      className={`flex justify-between items-center py-0.5 ${
                        i < gradeRules.length - 1 ? 'border-b border-blue-100' : ''
                      }`}
                    >
                      <span
                        className={`font-bold ${
                          i === gradeRules.length - 1 ? 'text-rose-700' : 'text-gray-800'
                        }`}
                      >
                        {is1to5Class(selectedClass) ? r.text15 : r.text10}
                      </span>
                      <span
                        className={`font-bold px-2 py-0.5 rounded border text-[11px] ${getGradeCardBadgeClass(
                          r.grade
                        )}`}
                      >
                        {r.grade}
                      </span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Card 2: SA Marks */}
              <div className="bg-purple-50/70 border border-purple-200 rounded-lg p-3 shadow-xs">
                <div className="font-extrabold text-purple-900 text-[13px] border-b border-purple-200 pb-1.5 mb-2 text-center">
                  2. ಸಂಕಲನಾತ್ಮಕ (SA1, SA2)
                  <div className="text-[11.5px] font-semibold text-purple-700">
                    {is1to5Class(selectedClass) ? '20% ಕ್ಕೆ (ಗರಿಷ್ಠ 20 ಅಂಕ)' : '30% ಕ್ಕೆ (ಗರಿಷ್ಠ 30 ಅಂಕ)'}
                  </div>
                </div>
                <div className="space-y-1 text-[12.5px]">
                  {gradeRules.map((r, i) => (
                    <div
                      key={i}
                      className={`flex justify-between items-center py-0.5 ${
                        i < gradeRules.length - 1 ? 'border-b border-purple-100' : ''
                      }`}
                    >
                      <span
                        className={`font-bold ${
                          i === gradeRules.length - 1 ? 'text-rose-700' : 'text-gray-800'
                        }`}
                      >
                        {is1to5Class(selectedClass) ? r.text20 : r.text30}
                      </span>
                      <span
                        className={`font-bold px-2 py-0.5 rounded border text-[11px] ${getGradeCardBadgeClass(
                          r.grade
                        )}`}
                      >
                        {r.grade}
                      </span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Card 3: Sem Total 50 Marks */}
              <div className="bg-amber-50/70 border border-amber-200 rounded-lg p-3 shadow-xs">
                <div className="font-extrabold text-amber-900 text-[13px] border-b border-amber-200 pb-1.5 mb-2 text-center">
                  3. ಸೆಮಿಸ್ಟರ್ ಒಟ್ಟು (Sem 1, 2)
                  <div className="text-[11.5px] font-semibold text-amber-700">50% ಕ್ಕೆ (ಗರಿಷ್ಠ 50 ಅಂಕ)</div>
                </div>
                <div className="space-y-1 text-[12.5px]">
                  {gradeRules.map((r, i) => (
                    <div
                      key={i}
                      className={`flex justify-between items-center py-0.5 ${
                        i < gradeRules.length - 1 ? 'border-b border-amber-100' : ''
                      }`}
                    >
                      <span
                        className={`font-bold ${
                          i === gradeRules.length - 1 ? 'text-rose-700' : 'text-gray-800'
                        }`}
                      >
                        {r.text50}
                      </span>
                      <span
                        className={`font-bold px-2 py-0.5 rounded border text-[11px] ${getGradeCardBadgeClass(
                          r.grade
                        )}`}
                      >
                        {r.grade}
                      </span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Card 4: Grand Total 100 Marks */}
              <div className="bg-emerald-50/70 border border-emerald-200 rounded-lg p-3 shadow-xs">
                <div className="font-extrabold text-emerald-900 text-[13px] border-b border-emerald-200 pb-1.5 mb-2 text-center">
                  4. ವಾರ್ಷಿಕ ಮಹಾ ಒಟ್ಟು
                  <div className="text-[11.5px] font-semibold text-emerald-700">100% ಕ್ಕೆ (ಗರಿಷ್ಠ 100 ಅಂಕ)</div>
                </div>
                <div className="space-y-1 text-[12.5px]">
                  {gradeRules.map((r, i) => (
                    <div
                      key={i}
                      className={`flex justify-between items-center py-0.5 ${
                        i < gradeRules.length - 1 ? 'border-b border-emerald-100' : ''
                      }`}
                    >
                      <span
                        className={`font-bold ${
                          i === gradeRules.length - 1 ? 'text-rose-700' : 'text-gray-800'
                        }`}
                      >
                        {r.text100}
                      </span>
                      <span
                        className={`font-bold px-2 py-0.5 rounded border text-[11px] ${getGradeCardBadgeClass(
                          r.grade
                        )}`}
                      >
                        {r.grade}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            {/* Row Totals Reference */}
            <div className="bg-gray-50 border border-gray-200 rounded-lg p-3 text-[12.5px] text-gray-700 mb-4">
              <div className="font-bold text-gray-900 mb-1 flex items-center justify-between">
                <div className="flex items-center gap-1.5">
                  <span>📊</span>
                  <span>ಒಟ್ಟು ಅಂಕಗಳ ಶ್ರೇಣಿ ವ್ಯಾಪ್ತಿ (Row Totals Reference):</span>
                </div>
                <span className="text-[11px] text-gray-500 font-normal">
                  (7 ವಿಷಯಗಳ ಒಟ್ಟು ಅಂಕಗಳ ಆಧಾರದ ಮೇಲೆ)
                </span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-2 text-[12px]">
                <div className="p-2 bg-white rounded border border-gray-200">
                  <span className="font-bold text-blue-900">FA ಒಟ್ಟು (70 ಅಂಕ):</span>
                  <div className="text-[11px] text-gray-600 mt-0.5 leading-snug">
                    {gradeRules.map((r, i) => {
                      const min70 = Math.round((r.min10 / 10) * 70);
                      const nextMin = i > 0 ? Math.round((gradeRules[i - 1].min10 / 10) * 70) - 1 : 70;
                      const txt = i === gradeRules.length - 1 ? `<${Math.round((gradeRules[i - 1]?.min10 / 10) * 70) || min70}` : `${min70}-${nextMin}`;
                      return `${r.grade}: ${txt}`;
                    }).join(' | ')}
                  </div>
                </div>
                <div className="p-2 bg-white rounded border border-gray-200">
                  <span className="font-bold text-purple-900">SA ಒಟ್ಟು (210 ಅಂಕ):</span>
                  <div className="text-[11px] text-gray-600 mt-0.5 leading-snug">
                    {gradeRules.map((r, i) => {
                      const min210 = Math.round((r.min30 / 30) * 210);
                      const nextMin = i > 0 ? Math.round((gradeRules[i - 1].min30 / 30) * 210) - 1 : 210;
                      const txt = i === gradeRules.length - 1 ? `<${Math.round((gradeRules[i - 1]?.min30 / 30) * 210) || min210}` : `${min210}-${nextMin}`;
                      return `${r.grade}: ${txt}`;
                    }).join(' | ')}
                  </div>
                </div>
                <div className="p-2 bg-white rounded border border-gray-200">
                  <span className="font-bold text-amber-900">ಸೆಮಿಸ್ಟರ್ ಒಟ್ಟು (350 ಅಂಕ):</span>
                  <div className="text-[11px] text-gray-600 mt-0.5 leading-snug">
                    {gradeRules.map((r, i) => {
                      const min350 = Math.round((r.min50 / 50) * 350);
                      const nextMin = i > 0 ? Math.round((gradeRules[i - 1].min50 / 50) * 350) - 1 : 350;
                      const txt = i === gradeRules.length - 1 ? `<${Math.round((gradeRules[i - 1]?.min50 / 50) * 350) || min350}` : `${min350}-${nextMin}`;
                      return `${r.grade}: ${txt}`;
                    }).join(' | ')}
                  </div>
                </div>
                <div className="p-2 bg-white rounded border border-gray-200">
                  <span className="font-bold text-emerald-900">ವಾರ್ಷಿಕ ಒಟ್ಟು (700 ಅಂಕ):</span>
                  <div className="text-[11px] text-gray-600 mt-0.5 leading-snug">
                    {gradeRules.map((r, i) => {
                      const min700 = Math.round((r.min100 / 100) * 700);
                      const nextMin = i > 0 ? Math.round((gradeRules[i - 1].min100 / 100) * 700) - 1 : 700;
                      const txt = i === gradeRules.length - 1 ? `<${Math.round((gradeRules[i - 1]?.min100 / 100) * 700) || min700}` : `${min700}-${nextMin}`;
                      return `${r.grade}: ${txt}`;
                    }).join(' | ')}
                  </div>
                </div>
              </div>
            </div>

            <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-gray-200">
              <div className="text-[12px] text-gray-600 flex items-center gap-1.5">
                {isCustomRulesActive ? (
                  <span className="text-emerald-700 font-bold flex items-center gap-1">
                    <span>✨</span> ಕಸ್ಟಮ್ ಶ್ರೇಣಿ ನಿಯಮಗಳು ಸಕ್ರಿಯವಾಗಿವೆ (ಮುಖ್ಯ ಕೋಷ್ಟಕಕ್ಕೆ ಲಿಂಕ್ ಆಗಿದೆ)
                  </span>
                ) : (
                  <span>ಸರ್ಕಾರಿ CCE ಪ್ರಮಾಣಿತ ಶ್ರೇಣಿ ಮಾನದಂಡ</span>
                )}
              </div>
              <div className="flex items-center gap-2">
                {!isEditingGradeRules && (
                  <button
                    type="button"
                    onClick={handleStartEditRules}
                    className="py-2 px-4 bg-amber-500 hover:bg-amber-600 text-gray-950 font-extrabold rounded-lg shadow-sm transition-colors cursor-pointer text-sm flex items-center gap-1.5"
                  >
                    <Edit3 className="w-4 h-4 stroke-[2.5]" />
                    <span>ಶ್ರೇಣಿ & ವ್ಯಾಪ್ತಿ ಸಂಪಾದಿಸಿ (Edit)</span>
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => setShowGradeModal(false)}
                  className="py-2 px-6 bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white font-bold rounded-lg shadow-sm transition-colors cursor-pointer text-sm"
                >
                  ಸರಿ, ಅರ್ಥವಾಯಿತು (Close)
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
      <style>{`
        @page { size: A4 landscape; margin: 0; }
        * { box-sizing: border-box; scrollbar-width: thin !important; }
        @keyframes autoSaveBlink {
          0%, 100% {
            opacity: 1;
            background-color: #86efac !important;
            border-color: #15803d !important;
            box-shadow: 0 0 10px rgba(34, 197, 94, 0.75) !important;
          }
          50% {
            opacity: 0.45;
            background-color: #dcfce7 !important;
            border-color: #22c55e !important;
            box-shadow: 0 0 2px rgba(34, 197, 94, 0.15) !important;
          }
        }
        .auto-save-blink {
          animation: autoSaveBlink 0.6s infinite ease-in-out !important;
        }
        html, body {
          margin: 0;
          background: #f3f1e8;
          -webkit-print-color-adjust: exact;
          print-color-adjust: exact;
          font-family: 'Noto Sans Kannada', system-ui, sans-serif;
          overflow-x: hidden !important;
          scrollbar-width: thin !important;
          scrollbar-color: #94a3b8 #f1f5f9 !important;
        }
        .report-shell {
          width: calc(100% - 24px) !important;
          max-width: 1580px !important;
          margin-left: auto !important;
          margin-right: auto !important;
          background: #fff;
          box-shadow: 0 4px 24px rgba(0,0,0,0.10);
          border: 1px solid #000;
          box-sizing: border-box;
          overflow-x: hidden !important;
        }
        @media (min-width: 1620px) {
          .report-shell {
            width: calc(100% - 32px) !important;
            max-width: 1580px !important;
            margin: 0 auto !important;
          }
        }
        /* Sleek 6px scrollbar on right side */
        ::-webkit-scrollbar {
          width: 6px !important;
          height: 0px !important;
        }
        ::-webkit-scrollbar-track {
          background: #f1f5f9 !important;
        }
        ::-webkit-scrollbar-thumb {
          background: #94a3b8 !important;
          border-radius: 9999px !important;
        }
        ::-webkit-scrollbar-thumb:hover {
          background: #64748b !important;
        }
        /* Completely remove bottom horizontal slider in screen */
        ::-webkit-scrollbar:horizontal {
          display: none !important;
          height: 0px !important;
          width: 0px !important;
        }
        .table-wrap {
          overflow-x: hidden !important;
          scrollbar-width: none !important;
          -ms-overflow-style: none !important;
          background: #fff;
          width: 100% !important;
          max-width: 100% !important;
        }
        .table-wrap::-webkit-scrollbar {
          display: none !important;
          height: 0px !important;
          width: 0px !important;
        }
        .table-wrap table {
          margin: 0 auto;
        }

        /* Interactive Row Light Color Highlight */
        #studentRowsContainer tbody tr.interactive-row:hover,
        #studentRowsContainer tbody tr.interactive-row:focus-within {
          background-color: #dbeafe !important;
        }
        #studentRowsContainer tbody tr.interactive-row:hover > td.fa-cell,
        #studentRowsContainer tbody tr.interactive-row:hover > td.sa-cell,
        #studentRowsContainer tbody tr.interactive-row:hover > td.anka-total-cell,
        #studentRowsContainer tbody tr.interactive-row:focus-within > td.fa-cell,
        #studentRowsContainer tbody tr.interactive-row:focus-within > td.sa-cell,
        #studentRowsContainer tbody tr.interactive-row:focus-within > td.anka-total-cell {
          background-color: #dbeafe !important;
        }
        #studentRowsContainer tbody tr.interactive-row:hover input:not(:focus),
        #studentRowsContainer tbody tr.interactive-row:focus-within input:not(:focus) {
          background-color: transparent !important;
        }
        #studentRowsContainer tbody tr.interactive-row input:focus,
        .anka-mark-input:focus,
        .attendance-cell input:focus,
        .fa-cell input:focus,
        .sa-cell input:focus {
          outline: 1px solid #2563eb !important;
          border-color: #2563eb !important;
          outline-offset: -1px;
          z-index: 20;
          position: relative;
        }
        .vtxt { writing-mode: vertical-rl; transform: rotate(180deg); white-space: nowrap; letter-spacing:0.2px; display:inline-flex !important; align-items:center !important; justify-content:center !important; text-align:center !important; vertical-align:middle !important; margin:0 auto !important; color:#000; }
        thead, thead tr, thead th { text-align: center !important; vertical-align: middle !important; }
        input { font-family: inherit; }
        input:focus { outline: 1px solid #2563eb !important; outline-offset: -1px; z-index:20; position:relative; }
        table { border-collapse: collapse; table-layout: fixed; }
        th, td { border:1px solid #000 !important; color:#000; }
        .header-input { border:1px solid #000 !important; background:#fff; color:#000; font-weight:700; border-radius:2px; height:26px; padding:0 6px; font-family: inherit; }
        .header-input::placeholder { color:#777 !important; opacity:1 !important; font-weight:500; }
        .header-select { border:1px solid #000 !important; background:#fff; color:#000; font-weight:800; border-radius:2px; height:26px; padding:0 6px 0 8px; font-family: inherit; }
        .header-select:focus { background:#ffffff !important; outline:2px solid #2563eb !important; border-color:#2563eb !important; outline-offset:-2px; }
        .header-input:focus { background:#ffffff !important; outline:2px solid #2563eb !important; border-color:#2563eb !important; outline-offset:-2px; }
        .co-cell { background:#DCFCE7 !important; padding:0 !important; position:relative; vertical-align:middle !important; }
        .co-cell-2 { background:#DCFCE7 !important; padding:0 !important; position:relative; vertical-align:middle !important; }
        .co-select-direct {
          width:100%; height:100%; min-height:36px;
          background:#DCFCE7 !important;
          border:0 !important;
          outline:0 !important;
          box-shadow:none !important;
          color:#000; font-weight:800 !important; font-size:15px !important;
          text-align:center; text-align-last:center;
          cursor:pointer;
          appearance:none !important;
          -webkit-appearance:none !important;
          -moz-appearance:none !important;
          padding:0;
          display:flex;
          align-items:center;
          justify-content:center;
        }
        .co-select-direct:hover {
          background:#DCFCE7 !important;
          border:0 !important;
        }
        .co-select-direct:focus {
          outline: 1px solid #2563eb !important;
          border-color: #2563eb !important;
          outline-offset: -1px;
          z-index: 20;
          position: relative;
        }
        .co-select-direct option {
          background:#DCFCE7 !important;
          color:#000 !important;
          border:0 !important;
          outline:0 !important;
        }
        .co-wrap {
          width:100%; height:100%; min-height:36px;
          display:flex; align-items:center; justify-content:center;
          position:relative;
          background:transparent;
        }
        .co-wrap::after {
          content:'▼';
          position:absolute;
          right:3px; top:50%; transform:translateY(-50%);
          font-size:9px; color:#000; 
          opacity:0; 
          pointer-events:none;
          transition: opacity 0.12s ease;
        }
        .co-cell:hover .co-wrap::after,
        .co-cell:focus-within .co-wrap::after,
        .co-cell-2:hover .co-wrap::after,
        .co-cell-2:focus-within .co-wrap::after {
          opacity:0.75;
        }
        .co-final {
          font-weight:800 !important; font-size:15px !important; text-align:center;
          color:#000; line-height:1;
          display:flex; align-items:center; justify-content:center;
          width:100%; height:100%; min-height:36px;
          background:#c0facd !important;
        }
        #studentRowsContainer .moulankan-cell {
          background: #FCE7F3 !important;
        }
        #studentRowsContainer .moulankan-cell,
        #studentRowsContainer .moulankan-cell * {
          white-space: nowrap !important;
          font-size: 11.5px !important;
        }
        td.stu-cell,
        .stu-cell,
        .stu-detail-wrap,
        .stu-detail-wrap *,
        .stu-row {
          background-color: #f4eae0 !important;
        }
        .stu-detail-wrap input:not(:focus),
        .stu-detail-wrap select:not(:focus),
        .stu-row input:not(:focus),
        .stu-row select:not(:focus),
        .stu-input-name:not(:focus),
        input.stu-input-name:not(:focus),
        .stu-input-date:not(:focus) {
          background-color: transparent !important;
          background: transparent !important;
        }
        .stu-detail-wrap input:focus,
        .stu-detail-wrap select:focus,
        .stu-row input:focus,
        .stu-row select:focus,
        .stu-input-name:focus,
        input.stu-input-name:focus,
        .stu-input-date:focus,
        .anka-mark-input:focus,
        .attendance-cell input:focus,
        .fa-cell input:focus,
        .sa-cell input:focus,
        .co-select-direct:focus,
        #studentRowsContainer input:focus,
        #studentRowsContainer select:focus {
          background-color: #ffffff !important;
          background: #ffffff !important;
          outline: 1px solid #2563eb !important;
          border-color: #2563eb !important;
          outline-offset: -1px !important;
          color: #000000 !important;
          z-index: 30 !important;
          position: relative !important;
        }
        .stu-detail-wrap { width:100% !important; max-width:100% !important; min-height:324px; display:flex; flex-direction:column; justify-content:space-around; padding:6px 4px; box-sizing:border-box !important; }
        .stu-row { border-bottom:1px dotted #000; padding:4px 0; display:flex; align-items:center; min-height:44px; width:100% !important; box-sizing:border-box !important; }
        .stu-row:last-child { border-bottom:0; }
        .stu-label { font-size:13px; font-weight:700; color:#000; white-space:nowrap; padding:0 3px 0 4px; flex-shrink:0; }
        .stu-input { flex:1; border:0 !important; outline:0 !important; padding:4px 4px; font-size:14.5px; font-weight:700; color:#000; min-width:0; }
        .stu-input::placeholder,
        .stu-input-name::placeholder,
        input.stu-input-name::placeholder {
          color: #6b7280 !important;
          opacity: 0.85 !important;
          font-weight: 600 !important;
          font-size: 14px !important;
          font-style: normal !important;
        }
        .anka-mark-input::placeholder,
        .attendance-cell input::placeholder,
        .fa-cell input::placeholder,
        .sa-cell input::placeholder {
          color: #374151 !important;
          opacity: 0.85 !important;
          font-weight: 700 !important;
          text-align: center !important;
        }
        #studentRowsContainer tbody tr.interactive-row input:focus,
        #studentRowsContainer tbody tr.interactive-row select:focus,
        #studentRowsContainer input:focus,
        #studentRowsContainer select:focus,
        .anka-mark-input:focus,
        .attendance-cell input:focus,
        .fa-cell input:focus,
        .sa-cell input:focus,
        .co-select-direct:focus {
          outline: 1px solid #2563eb !important;
          border: 1px solid #2563eb !important;
          border-color: #2563eb !important;
          outline-offset: -1px !important;
          box-shadow: none !important;
          color: #000000 !important;
          z-index: 30 !important;
          position: relative !important;
        }
        .stu-input-name,
        input.stu-input-name {
          font-size: 16.5px !important;
          font-weight: 800 !important;
          padding: 4px 4px !important;
          color: #000 !important;
        }
        .stu-hdr { border:0.6px solid #000 !important; box-sizing:border-box !important; }
        .sem-cell { border:0.6px solid #000 !important; box-sizing:border-box !important; }
        .sem-hdr { border:0.6px solid #000 !important; box-sizing:border-box !important; }
        .stu-input-date {
          flex:1;
          border:0 !important;
          outline:0 !important;
          padding:2px 2px !important;
          font-size:13.5px;
          font-weight:700;
          color:#000;
          min-width:0;
          width:100%;
          height:30px;
          cursor:pointer;
          font-family: inherit;
          line-height:normal !important;
          color-scheme: light;
          box-sizing:border-box !important;
        }
        .print-stu-thick-border td {
          border-bottom: 1.5px solid #000 !important;
        }
        .stu-input-date::-webkit-calendar-picker-indicator {
          cursor:pointer;
          opacity:0.95;
          width:18px;
          height:18px;
          padding:0;
          margin-left:2px;
          margin-right:1px;
          background-size:16px 16px;
          background-position:center;
          filter: brightness(0) saturate(100%);
        }
        .stu-input-date::-webkit-datetime-edit { padding:2px 1px 0 1px !important; color:#000; line-height:normal !important; }
        .stu-input-date::-webkit-datetime-edit-fields-wrapper { padding:0 !important; line-height:normal !important; }
        .anka-th {
          font-size:13px !important;
          font-weight:700 !important;
          background:#eff6ff !important;
          text-align:center !important;
          vertical-align:middle !important;
        }
        .anka-mark-input {
          font-size:15px !important;
          font-weight:700 !important;
          background: transparent !important;
          border: 0 !important;
        }
        .cce-main-table {
          table-layout: fixed !important;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
        }
        #printableArea td.print-stu-thick-bottom,
        #printableArea tr.print-stu-thick-bottom td {
          border-bottom: 1.5px solid #000 !important;
        }
        @media print {
          @page {
            size: A4 landscape;
            margin: 0;
          }
          html, body {
            background: #fff !important;
            margin: 0 !important;
            padding: 0 !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
            font-family: 'Noto Sans Kannada', Arial, sans-serif !important;
          }
          .no-print,
          .report-shell,
          #printableArea .paper-header {
            display: none !important;
          }
          #printableArea {
            display: block !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
            box-sizing: border-box !important;
            background: #fff !important;
          }
          #printableArea div {
            overflow: visible !important;
          }
          #printableArea .print-page {
            page-break-after: always !important;
            break-after: page !important;
            page-break-inside: avoid !important;
            break-inside: avoid !important;
            width: 100% !important;
            box-sizing: border-box !important;
            margin: 0 !important;
          }
          #printableArea .print-page-1to5,
          #printableArea .print-page.print-page-1to5,
          #printableArea div.print-page-1to5 {
            padding: 8mm 10mm 5mm 10mm !important;
            box-sizing: border-box !important;
          }
          #printableArea .print-page-6to9,
          #printableArea .print-page.print-page-6to9,
          #printableArea div.print-page-6to9 {
            padding: 8mm 10mm 10mm 10mm !important;
          }
          #printableArea .print-page:last-child {
            page-break-after: auto !important;
            break-after: auto !important;
          }
          #printableArea .print-page-header {
            display: block !important;
            margin-bottom: 3px !important;
            width: 100% !important;
          }
          #printableArea .print-signatures-row {
            display: flex !important;
            font-size: 10.5px !important;
            font-weight: 700 !important;
            margin-top: 14px !important;
            padding-top: 4px !important;
            border-top: none !important;
          }
          #printableArea table {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            table-layout: fixed !important;
            border-collapse: collapse !important;
            box-sizing: border-box !important;
          }
          thead,
          thead tr,
          thead th,
          #printableArea thead,
          #printableArea thead tr,
          #printableArea thead th {
            text-align: center !important;
            vertical-align: middle !important;
          }
          #printableArea th {
            border: 0.5px solid #000 !important;
            background: #fff !important;
            color: #000 !important;
            min-width: 0 !important;
            max-width: none !important;
            padding: 1px 0 !important;
            overflow: hidden !important;
            vertical-align: middle !important;
            text-align: center !important;
            box-sizing: border-box !important;
          }
          #printableArea thead th {
            font-size: 10px !important;
            font-weight: 700 !important;
            text-align: center !important;
            vertical-align: middle !important;
            padding: 1px 0 !important;
          }
          #printableArea .print-page-header > div:first-child {
            font-size: 14.5px !important;
          }
          #printableArea thead th.print-sec-hdr,
          #printableArea thead th.print-sec-hdr * {
            font-size: 13px !important;
          }
          #printableArea thead th.print-co-hdr,
          #printableArea thead th.print-co-hdr * {
            font-size: 13px !important;
          }
          #printableArea thead th *,
          #printableArea thead th div,
          #printableArea thead th span {
            text-align: center !important;
            vertical-align: middle !important;
            margin-left: auto !important;
            margin-right: auto !important;
          }
          #printableArea thead th div:not(.co-skill-line) {
            text-align: center !important;
            vertical-align: middle !important;
            margin: 0 auto !important;
            width: 100% !important;
          }
          #printableArea tbody tr {
            height: 20.8px !important;
          }
          #printableArea .print-page-1to5 tbody tr {
            height: 21.2px !important;
          }
          #printableArea tbody td:not([rowspan]),
          #printableArea tbody td[rowspan="1"] {
            height: 20.8px !important;
          }
          #printableArea .print-page-1to5 tbody td:not([rowspan]),
          #printableArea .print-page-1to5 tbody td[rowspan="1"] {
            height: 21.2px !important;
          }
          #printableArea tbody td {
            border: 0.5px solid #000 !important;
            background: #fff;
            color: #000 !important;
            min-width: 0 !important;
            max-width: none !important;
            overflow: hidden !important;
            box-sizing: border-box !important;
            text-align: center !important;
            vertical-align: middle !important;
            padding: 0 !important;
            line-height: normal !important;
            font-size: 12.5px !important;
          }
          #printableArea tbody td[rowspan] {
            height: auto !important;
            vertical-align: middle !important;
            text-align: center !important;
          }
          #printableArea tbody td:not(.stu-print-td) {
            text-align: center !important;
            vertical-align: middle !important;
          }
          #printableArea tbody td:not(.stu-print-td) * {
            text-align: center !important;
            vertical-align: middle !important;
          }
          #printableArea tbody td:not(.stu-print-td) > div:not(.vtxt) {
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            text-align: center !important;
            margin: 0 auto !important;
            width: 100% !important;
            height: 100% !important;
          }
          #printableArea .moulankan-cell,
          #printableArea .moulankan-cell *,
          #printableArea .moulankan-pct-cell,
          #printableArea .moulankan-pct-cell *,
          #printableArea .moulankan-print-cell,
          #printableArea .moulankan-print-cell *,
          #printableArea .moulankan-pct-print-cell,
          #printableArea .moulankan-pct-print-cell * {
            background: transparent !important;
            white-space: nowrap !important;
            font-size: 9.5px !important;
            font-weight: 700 !important;
            letter-spacing: -0.3px !important;
            padding: 0 1px !important;
            text-align: center !important;
            vertical-align: middle !important;
          }
          #printableArea tr.row-highlight-print td.moulankan-cell,
          #printableArea tr.row-highlight-print td.moulankan-pct-cell,
          #printableArea tr.row-highlight-print td.moulankan-print-cell,
          #printableArea tr.row-highlight-print td.moulankan-pct-print-cell,
          #printableArea tr.row-highlight-print-50 td.moulankan-cell,
          #printableArea tr.row-highlight-print-50 td.moulankan-pct-cell,
          #printableArea tr.row-highlight-print-50 td.moulankan-print-cell,
          #printableArea tr.row-highlight-print-50 td.moulankan-pct-print-cell,
          #printableArea tr.row-highlight-print-100 td.moulankan-cell,
          #printableArea tr.row-highlight-print-100 td.moulankan-pct-cell,
          #printableArea tr.row-highlight-print-100 td.moulankan-print-cell,
          #printableArea tr.row-highlight-print-100 td.moulankan-pct-print-cell {
            background: #e5e7eb !important;
          }
          #printableArea td.attendance-cell,
          #printableArea td.attendance-cell * {
            font-size: 10px !important;
            text-align: center !important;
            vertical-align: middle !important;
          }
          #printableArea td.stu-print-td {
            text-align: left !important;
            vertical-align: middle !important;
            padding: 6px 8px !important;
          }
          #printableArea .stu-print-wrap {
            display: flex !important;
            flex-direction: column !important;
            justify-content: center !important;
            align-items: flex-start !important;
            text-align: left !important;
            width: 100% !important;
            height: 100% !important;
          }
          #printableArea .stu-print-name {
            font-size: 12.5px !important;
            font-weight: 800 !important;
            text-align: left !important;
            margin-bottom: 8px !important;
            line-height: 1.3 !important;
            white-space: normal !important;
            word-break: break-word !important;
            overflow-wrap: break-word !important;
            display: -webkit-box !important;
            -webkit-line-clamp: 2 !important;
            -webkit-box-orient: vertical !important;
            overflow: hidden !important;
            max-height: 34px !important;
          }
          #printableArea .stu-print-detail {
            font-size: 10px !important;
            font-weight: 600 !important;
            text-align: left !important;
            margin-bottom: 7px !important;
            line-height: 1.4 !important;
          }
          #printableArea .stu-print-detail:last-child {
            margin-bottom: 0 !important;
          }
          #printableArea .sub-mark-normal {
            font-weight: 400 !important;
          }
          #printableArea .sub-mark-bold {
            font-weight: 800 !important;
          }
          #printableArea .sub-grade-normal {
            font-weight: 400 !important;
          }
          #printableArea .co-grade-td {
            font-weight: 800 !important;
            font-size: 11.5px !important;
            text-align: center !important;
            vertical-align: middle !important;
          }
          #printableArea tbody tr.row-highlight-print td,
          #printableArea tbody tr.row-highlight-print-50 td,
          #printableArea tbody tr.row-highlight-print-100 td,
          #printableArea tbody tr td.highlight-print-cell,
          #printableArea tbody tr td.highlight-print-cell-50,
          #printableArea td.highlight-print-cell,
          #printableArea td.highlight-print-cell-50 {
            background: #e5e7eb !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          #printableArea tbody .vtxt {
            writing-mode: vertical-rl !important;
            transform: rotate(180deg) !important;
            white-space: nowrap !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            text-align: center !important;
            vertical-align: middle !important;
            margin: 0 auto !important;
            line-height: 1 !important;
          }
          #printableArea tbody td .sem-vtxt {
            font-size: 10px !important;
          }
          #printableArea tbody td .result-vtxt {
            font-size: 12.5px !important;
            max-height: 170px !important;
          }
          #printableArea thead .vtxt {
            writing-mode: vertical-rl !important;
            transform: rotate(180deg) !important;
            white-space: nowrap !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            text-align: center !important;
            vertical-align: middle !important;
            margin: 0 auto !important;
            font-size: 10px !important;
            line-height: 1 !important;
          }
          #printableArea .co-vtxt {
            font-size: 9px !important;
          }
          #printableArea .print-co-hdr,
          #printableArea .print-co-hdr * {
            white-space: nowrap !important;
            font-size: 13px !important;
            font-weight: 700 !important;
            line-height: 1.1 !important;
          }
          #printableArea th.print-co-skill-th {
            text-align: center !important;
            vertical-align: middle !important;
            padding: 1px 0 !important;
            overflow: visible !important;
          }
          #printableArea th.print-co-skill-th > div,
          #printableArea th.print-co-skill-th .co-skill-flex {
            display: flex !important;
            flex-direction: row !important;
            align-items: center !important;
            justify-content: center !important;
            text-align: center !important;
            width: 100% !important;
            height: 100% !important;
            margin: 0 auto !important;
            gap: 2px !important;
          }
          #printableArea th.print-co-skill-th .co-skill-line,
          #printableArea th.print-co-skill-th > div > div {
            display: inline-flex !important;
            writing-mode: vertical-rl !important;
            transform: rotate(180deg) !important;
            transform-origin: center center !important;
            white-space: nowrap !important;
            width: auto !important;
            max-width: none !important;
            min-width: 0 !important;
            margin: 0 !important;
            padding: 0 !important;
            flex: 0 0 auto !important;
            font-size: 9.5px !important;
            font-weight: 700 !important;
            line-height: 1.1 !important;
            align-items: center !important;
            justify-content: center !important;
            text-align: center !important;
          }
        }
      `}</style>

      <div className="report-shell">
        {/* Header Section */}
        <div style={{ borderBottom: '1px solid #000' }}>
          <div
            className="py-[8px] px-4 flex items-center justify-between flex-wrap gap-2"
            style={{ background: '#fef9c3' }}
          >
            <div className="w-6 hidden md:block"></div>
            <h1
              className="font-extrabold text-[19px] sm:text-[21px] md:text-[22px] leading-tight tracking-tight text-center flex-1"
              style={{ color: '#000', fontWeight: 800, fontSize: '21px', lineHeight: '25px' }}
            >
              {classLevel === '1-5'
                ? '1 ರಿಂದ 5 ನೇ ತರಗತಿಯ ನಿರಂತರ ಮತ್ತು ವ್ಯಾಪಕ ಮೌಲ್ಯಮಾಪನ ಕ್ರೋಡೀಕೃತ ವಹಿ'
                : '6 ರಿಂದ 9 ನೇ ತರಗತಿಯ ನಿರಂತರ ಮತ್ತು ವ್ಯಾಪಕ ಮೌಲ್ಯಮಾಪನ ಕ್ರೋಡೀಕೃತ ವಹಿ'}
            </h1>
            {/* English Auto-save status badge with whole line blink */}
            <div
              className={`inline-flex items-center gap-2 px-3.5 py-1 rounded-full text-[12.5px] font-extrabold shadow-sm no-print select-none transition-colors duration-200 ${
                isSaving ? 'auto-save-blink' : ''
              }`}
              style={{
                background: isSaving ? '#86efac' : '#ecfdf5',
                color: isSaving ? '#052e16' : '#047857',
                border: '1.5px solid #10b981',
                boxSizing: 'border-box',
                height: '30px',
              }}
              title="Auto Save active. All entries are saved continuously in local storage."
            >
              <span className={`w-2.5 h-2.5 rounded-full ${isSaving ? 'bg-emerald-800 animate-ping' : 'bg-emerald-500 animate-ping'}`}></span>
              <CheckCircle2 className="w-4 h-4 text-emerald-700 stroke-[2.5]" />
              <span className="whitespace-nowrap font-black tracking-wide">
                Auto Saved {selectedClass ? `(Class ${selectedClass})` : ''}
              </span>
              {lastSavedTime && (
                <span className="text-[12px] text-emerald-950 font-black tracking-wide">
                  • {lastSavedTime}
                </span>
              )}
            </div>
          </div>

          <div
            className="flex items-center justify-between gap-3 px-6 py-[6px] text-[15px] font-bold flex-wrap lg:flex-nowrap"
            style={{ background: '#d9e8d5', borderTop: '1px solid #000', color: '#000', fontSize: '15px', fontWeight: 700 }}
          >
            {/* 1. Class Level Toggle & Select */}
            <div className="flex-1 min-w-[240px] flex items-center justify-center gap-2">
              <div className="inline-flex items-center gap-0.5 bg-white/90 p-0.5 rounded-md border border-gray-400 shadow-xs no-print">
                <button
                  type="button"
                  onClick={() => handleLevelToggle('1-5')}
                  className={`px-2 py-0.5 text-[12px] font-black rounded transition-all cursor-pointer ${
                    classLevel === '1-5'
                      ? 'bg-blue-700 text-white shadow-xs'
                      : 'text-gray-700 hover:text-black hover:bg-gray-200'
                  }`}
                  title="1 ರಿಂದ 5 ನೇ ತರಗತಿ"
                >
                  1-5 ನೇ
                </button>
                <button
                  type="button"
                  onClick={() => handleLevelToggle('6-9')}
                  className={`px-2 py-0.5 text-[12px] font-black rounded transition-all cursor-pointer ${
                    classLevel === '6-9'
                      ? 'bg-blue-700 text-white shadow-xs'
                      : 'text-gray-700 hover:text-black hover:bg-gray-200'
                  }`}
                  title="6 ರಿಂದ 9 ನೇ ತರಗತಿ"
                >
                  6-9 ನೇ
                </button>
              </div>

              <span className="whitespace-nowrap">ತರಗತಿ :</span>
              <select
                id="headerClassSelect"
                className={`header-select text-[15px] min-w-[95px] text-center ${
                  !selectedClass ? 'ring-2 ring-amber-500 bg-amber-50/90' : ''
                }`}
                style={{
                  fontSize: '15px',
                  fontWeight: selectedClass ? 700 : 500,
                  color: selectedClass ? '#000' : '#777',
                }}
                value={selectedClass}
                onChange={(e) => handleClassChange(e.target.value)}
                aria-label="ತರಗತಿ ಆಯ್ಕೆ"
              >
                <option value="" style={{ color: '#777', fontWeight: 500 }}>
                  ತರಗತಿ ಆಯ್ಕೆ *
                </option>
                {(classLevel === '1-5' ? CLASSES_1_5 : CLASSES_6_9).map((cls) => (
                  <option key={cls} value={cls} style={{ color: '#000', fontWeight: 700 }}>
                    {cls}
                  </option>
                ))}
              </select>
            </div>

            {/* 2. School */}
            <div className="flex-1 min-w-[220px] flex items-center justify-center gap-1.5">
              <span className="text-[15px] font-bold" style={{ fontSize: '15px', fontWeight: 700 }}>
                ಶಾಲೆ:
              </span>
              <input
                id="schoolName"
                className="header-input text-[15px] w-full max-w-[220px] tracking-wide text-center"
                style={{ fontSize: '15px', fontWeight: 700 }}
                value={school}
                onChange={(e) => setSchool(e.target.value)}
                placeholder="ಶಾಲೆಯ ಹೆಸರು"
                title="ಶಾಲೆಯ ಹೆಸರು"
                aria-label="ಶಾಲೆಯ ಹೆಸರು"
              />
            </div>

            {/* 3. DISE Code */}
            <div className="flex-1 min-w-[220px] flex items-center justify-center gap-1.5">
              <span className="text-[15px] font-bold" style={{ fontSize: '15px', fontWeight: 700 }}>
                DISE Code:
              </span>
              <input
                id="diseCode"
                className="header-input text-[15px] w-full max-w-[190px] tracking-wide text-center"
                style={{ fontSize: '15px', fontWeight: 700 }}
                value={diseCode}
                maxLength={11}
                onChange={(e) => {
                  const val = e.target.value;
                  setDiseCode(val);
                  try {
                    localStorage.setItem('cce_dise_code', val);
                  } catch (err) {}
                }}
                placeholder="29010100101"
                title="11 ಅಂಕಿಗಳ ಡೈಸ್ ಕೋಡ್"
                aria-label="DISE Code"
              />
            </div>

            {/* 4. Year */}
            <div className="flex-1 min-w-[170px] flex items-center justify-center gap-1.5">
              <span className="text-[15px] font-bold" style={{ fontSize: '15px', fontWeight: 700 }}>
                ವರ್ಷ:
              </span>
              <input
                id="yearSelect"
                className="header-input text-[15px] w-[110px] text-center"
                style={{ fontSize: '15px', fontWeight: 700 }}
                value={year}
                onChange={(e) => setYear(e.target.value)}
                placeholder="2026-27"
                title="ಶೈಕ್ಷಣಿಕ ವರ್ಷ ನಮೂದಿಸಿ"
                aria-label="ಶೈಕ್ಷಣಿಕ ವರ್ಷ"
              />
            </div>
          </div>

          {notification && (
            <div
              className="px-3 py-1.5 text-[12px] font-bold text-[#065f46] bg-[#d1fae5] border-t border-[#000] flex items-center gap-2 no-print"
              style={{ borderTop: '1px solid #000' }}
            >
              <CheckCircle2 className="w-4 h-4 text-emerald-700" />
              <span>{notification}</span>
            </div>
          )}
        </div>

        {/* Toolbar Buttons */}
        <div
          className="flex justify-start gap-2 px-3 py-[8px] bg-white no-print flex-wrap items-center"
          style={{ borderBottom: '1px solid #000' }}
        >
          <button
            id="addStudentBtn"
            type="button"
            onClick={handleAddStudent}
            className="inline-flex items-center gap-1.5 px-3 py-[5px] rounded-[6px] text-[13px] font-bold text-white shadow-sm hover:brightness-110 active:brightness-95 transition-all cursor-pointer"
            style={{ background: '#16a34a' }}
            title="ಹೊಸ ವಿದ್ಯಾರ್ಥಿ ಸೇರಿಸಿ"
          >
            <UserPlus className="w-4 h-4" />
            <span>ವಿದ್ಯಾರ್ಥಿ ಸೇರಿಸಿ</span>
          </button>

          <button
            id="bulkImportBtn"
            type="button"
            onClick={() => {
              if (!verifyPrerequisites()) return;
              setShowBulkImportModal(true);
            }}
            className="inline-flex items-center gap-1.5 px-3 py-[5px] rounded-[6px] text-[13px] font-bold text-white shadow-sm hover:brightness-110 active:brightness-95 transition-all cursor-pointer"
            style={{ background: '#0284c7' }}
            title="Excel / CSV / ಪಠ್ಯದ ಮೂಲಕ ಅನೇಕ ವಿದ್ಯಾರ್ಥಿಗಳನ್ನು ಒಟ್ಟಿಗೆ ಆಮದು ಮಾಡಿ"
          >
            <Upload className="w-4 h-4" />
            <span>ಬಲ್ಕ್ ಆಮದು (Import)</span>
          </button>

          {/* Student Search */}
          <div className="relative flex items-center min-w-[130px] max-w-[165px]">
            <Search className="w-3.5 h-3.5 text-gray-500 absolute left-2 pointer-events-none" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="ಹುಡುಕಿ (ಹೆಸರು/ದಾ.ಸಂ/DOB)..."
              className="w-full pl-7 pr-6 py-[3px] text-[12px] bg-white border border-gray-400 rounded-[5px] font-medium text-black focus:outline-none focus:border-blue-600 focus:ring-1 focus:ring-blue-500"
              title="ವಿದ್ಯಾರ್ಥಿಯ ಹೆಸರು, ದಾಖಲಾತಿ ಸಂಖ್ಯೆ, DOB (ಜನ್ಮ ದಿನಾಂಕ), SATS, PEN ಅಥವಾ ಕ್ರಮ ಸಂಖ್ಯೆಯ ಮೂಲಕ ಹುಡುಕಿ"
              aria-label="ವಿದ್ಯಾರ್ಥಿ ಹುಡುಕಿ"
            />
            {searchQuery && (
              <button
                type="button"
                onClick={() => setSearchQuery('')}
                className="absolute right-1.5 text-gray-400 hover:text-gray-700 cursor-pointer"
                title="ಹುಡುಕಾಟ ತೆರವುಗೊಳಿಸಿ"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
          {searchQuery && (
            <span className="text-[12px] font-bold text-gray-600">
              ({filteredStudentList.length}/{students.length})
            </span>
          )}

          <button
            id="exportExcelBtn"
            type="button"
            onClick={handleExportExcel}
            className="inline-flex items-center gap-1.5 px-3 py-[5px] rounded-[6px] text-[13px] font-bold text-white shadow-sm hover:brightness-110 active:brightness-95 transition-all cursor-pointer"
            style={{ background: '#2563eb' }}
            title="Excel ಗೆ ರಫ್ತು ಮಾಡಿ"
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>Excel ರಫ್ತು</span>
          </button>

          <button
            id="printBtn"
            type="button"
            onClick={handlePrint}
            className="inline-flex items-center gap-1.5 px-3 py-[5px] rounded-[6px] text-[13px] font-bold text-white shadow-sm hover:brightness-110 active:brightness-95 transition-all cursor-pointer"
            style={{ background: '#059669' }}
            title="A4 ಮುದ್ರಣ ಪೂರ್ವವೀಕ್ಷಣೆ ಹಾಗೂ ಪ್ರಿಂಟ್"
          >
            <Printer className="w-4 h-4" />
            <span>ಮುದ್ರಣ (Print)</span>
          </button>


          <button
            id="gradeRulesBtn"
            type="button"
            onClick={() => setShowGradeModal(true)}
            className="inline-flex items-center gap-1.5 px-3 py-[5px] rounded-[6px] text-[13px] font-bold text-white shadow-sm hover:brightness-110 active:brightness-95 transition-all cursor-pointer"
            style={{ background: '#7c3aed' }}
            title="ಕರ್ನಾಟಕ CCE ಶ್ರೇಣಿ ನಿಯಮಗಳ ಕೋಷ್ಟಕ (Grade Rules Chart)"
          >
            <Award className="w-4 h-4" />
            <span>ಶ್ರೇಣಿ ನಿಯಮಗಳು</span>
          </button>

          <a
            id="cleanTabBtn"
            href={typeof window !== 'undefined' ? window.location.href : '#'}
            target="_blank"
            rel="noopener noreferrer"
            className="inline-flex items-center gap-1.5 px-3 py-[5px] rounded-[6px] text-[13px] font-bold text-white shadow-sm hover:brightness-110 active:brightness-95 transition-all cursor-pointer"
            style={{ background: '#0284c7' }}
            title="Google AI Studio ಟೂಲ್‌ಬಾರ್, Remix, Device ಇಲ್ಲದೆ ಸ್ವತಂತ್ರ ಹೊಸ ವಿಂಡೋದಲ್ಲಿ ತೆರೆಯಿರಿ"
          >
            <ExternalLink className="w-4 h-4" />
            <span>ಹೊಸ ಟ್ಯಾಬ್ (Clean Tab)</span>
          </a>

          <button
            id="resetMarksBtn"
            type="button"
            onClick={handleResetMarks}
            className="inline-flex items-center gap-1.5 px-3 py-[5px] rounded-[6px] text-[13px] font-bold text-white shadow-sm hover:brightness-110 active:brightness-95 transition-all cursor-pointer"
            style={{ background: '#ea580c' }}
            title="ಅಂಕಗಳನ್ನು ಮರುಹೊಂದಿಸಿ"
          >
            <RotateCcw className="w-4 h-4" />
            <span>ಅಂಕಗಳನ್ನು ರೀಸೆಟ್</span>
          </button>

          <button
            id="resetAllBtn"
            type="button"
            onClick={handleResetAll}
            className="inline-flex items-center gap-1.5 px-3 py-[5px] rounded-[6px] text-[13px] font-bold text-white shadow-sm hover:brightness-110 active:brightness-95 transition-all cursor-pointer"
            style={{ background: '#dc2626' }}
            title="ಎಲ್ಲವನ್ನೂ ಮರುಹೊಂದಿಸಿ"
          >
            <Trash2 className="w-4 h-4" />
            <span>ಎಲ್ಲವನ್ನೂ ರೀಸೆಟ್</span>
          </button>
        </div>

        {/* Required Setup Warning Banner */}
        {!isEntryAllowed && (
          <div
            className="px-4 py-2 bg-amber-50 border-b border-amber-300 text-amber-950 flex items-center justify-between text-xs sm:text-[13px] font-bold no-print animate-pulse"
            style={{ borderBottom: '1px solid #f59e0b' }}
          >
            <div className="flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-amber-600 flex-shrink-0" />
              <span>
                ⚠️ ನಮೂದು (Data Entry) ಪ್ರಾರಂಭಿಸಲು ಮೊದಲು ತರಗತಿ ಹಾಗೂ 3 ಭಾಷೆಗಳನ್ನು (ಪ್ರಥಮ, ದ್ವಿತೀಯ, ತೃತೀಯ ಭಾಷೆ) ಆಯ್ಕೆ/ನಮೂದಿಸಿ!
              </span>
            </div>
            <div className="flex items-center gap-2 flex-shrink-0">
              {!isClassSelected && (
                <button
                  type="button"
                  onClick={() => document.getElementById('headerClassSelect')?.focus()}
                  className="bg-amber-200 text-amber-900 px-2.5 py-0.5 rounded text-[11px] font-extrabold hover:bg-amber-300 cursor-pointer"
                >
                  ತರಗತಿ ಆಯ್ಕೆಮಾಡಿ ➜
                </button>
              )}
              {(!isLang1Filled || !isLang2Filled || !isLang3Filled) && (
                <button
                  type="button"
                  onClick={() => {
                    if (!isLang1Filled) document.getElementById('langInput-lang1')?.focus();
                    else if (!isLang2Filled) document.getElementById('langInput-lang2')?.focus();
                    else document.getElementById('langInput-lang3')?.focus();
                  }}
                  className="bg-amber-200 text-amber-900 px-2.5 py-0.5 rounded text-[11px] font-extrabold hover:bg-amber-300 cursor-pointer"
                >
                  ಭಾಷೆಗಳು ನಮೂದಿಸಿ ➜
                </button>
              )}
            </div>
          </div>
        )}

        {/* Master CCE Register Table */}
        {(() => {
          const curSubjects = getSubjectsForClass(selectedClass, classLevel);
          const curCoSkills = getCoSkillsForClass(selectedClass, classLevel);
          const curParts = getAssessmentPartsForClass(selectedClass, classLevel);
          const { maxPerSub, maxGrand } = getMaxMarksConfig(selectedClass, classLevel);

          return (
            <div
              className="table-wrap flex justify-center w-full"
              id="studentRowsContainer"
              onClickCapture={(e) => {
                const target = e.target as HTMLElement;

                if (target.id?.startsWith('langInput-')) {
                  if (!isClassSelected) {
                    e.preventDefault();
                    e.stopPropagation();
                    target.blur?.();
                    showAlert(
                      'ತರಗತಿ ಆಯ್ಕೆ ಮಾಡಿ!',
                      'ವಿಷಯ / ಭಾಷೆಯ ಹೆಸರನ್ನು ನಮೂದಿಸುವ ಮೊದಲು ದಯವಿಟ್ಟು ಮೇಲ್ಭಾಗದಲ್ಲಿ ತರಗತಿಯನ್ನು (Class) ಆಯ್ಕೆ ಮಾಡಿ.',
                      undefined,
                      () => {
                        document.getElementById('headerClassSelect')?.focus();
                      }
                    );
                  }
                  return;
                }

                if (!isEntryAllowed) {
                  if (target.closest('tbody') || target.closest('.stu-cell') || target.tagName === 'INPUT' || target.tagName === 'SELECT') {
                    e.preventDefault();
                    e.stopPropagation();
                    target.blur?.();
                    verifyPrerequisites(e);
                  }
                }
              }}
              onFocusCapture={(e) => {
                const target = e.target as HTMLElement;
                if (target.id?.startsWith('langInput-')) return;

                if (!isEntryAllowed) {
                  if (target.closest('tbody') || target.closest('.stu-cell')) {
                    e.preventDefault();
                    e.stopPropagation();
                    target.blur?.();
                    verifyPrerequisites(e);
                  }
                }
              }}
              onKeyDownCapture={(e) => {
                const target = e.target as HTMLElement;
                if (target.id?.startsWith('langInput-')) return;

                if (!isEntryAllowed) {
                  if (target.closest('tbody') || target.closest('.stu-cell')) {
                    e.preventDefault();
                    e.stopPropagation();
                    target.blur?.();
                    verifyPrerequisites(e);
                  }
                }
              }}
            >
              <table
                className="text-[13px] leading-[1.1] mx-auto w-full cce-main-table"
                style={{ tableLayout: 'fixed', width: '100%', maxWidth: '100%' }}
              >
                <colgroup>
                  <col style={{ width: '1.8%' }} />
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '13.7%' : '11.65%' }} />
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.05%' }} />
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.05%' }} />
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.05%' }} />
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.7%' }} />
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.4%' }} />
                  {Array.from({ length: curSubjects.length * 2 }).map((_, i) => (
                    <col
                      key={`sub-${i}`}
                      style={{
                        width: is1to5Class(selectedClass, classLevel)
                          ? '3.38%'
                          : (i < 6 || i === 12 || i === 13 ? '3.2%' : '2.55%'),
                      }}
                    />
                  ))}
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.5%' }} />
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.5%' }} />
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.5%' }} />
                  {Array.from({ length: curCoSkills.length }).map((_, i) => (
                    <col key={`co-${i}`} style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '2.05%' }} />
                  ))}
                  <col style={{ width: is1to5Class(selectedClass, classLevel) ? '3.38%' : '3.0%' }} />
                </colgroup>

                <thead>
                  {/* Header Row 1 */}
                  <tr style={{ height: '34px' }}>
                    <th
                      rowSpan={3}
                      style={{
                        background: '#fbe4b9',
                        border: '1px solid #000',
                        padding: '2px',
                        verticalAlign: 'middle',
                      }}
                    >
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '11px', fontWeight: 700, minHeight: '88px' }}>
                        ಕ್ರ.ಸಂ.
                      </div>
                    </th>
                    <th
                      rowSpan={3}
                      className="stu-hdr"
                      style={{
                        background: '#fbe4b9',
                        border: '1px solid #000',
                        padding: '4px',
                        verticalAlign: 'middle',
                        textAlign: 'center',
                        boxSizing: 'border-box',
                      }}
                    >
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center' }}>
                        <div style={{ fontSize: '15px', fontWeight: 700, color: '#000', lineHeight: '16px' }}>
                          ವಿದ್ಯಾರ್ಥಿ ವಿವರ
                        </div>
                        <div style={{ fontSize: '12.5px', fontWeight: 500, color: '#000', marginTop: '2px' }}>
                          ಹೆಸರು / ದಾಖಲಾತಿ / ಲಿಂಗ
                        </div>
                        <div style={{ fontSize: '12px', fontWeight: 500, color: '#000', marginTop: '1px' }}>
                          / DOB / SATS / PEN
                        </div>
                      </div>
                    </th>
                    <th
                      rowSpan={3}
                      className="sem-hdr"
                      style={{
                        background: '#fbcfe5',
                        border: '1px solid #000',
                        padding: '2px',
                        verticalAlign: 'middle',
                        textAlign: 'center',
                        boxSizing: 'border-box',
                      }}
                    >
                      <div
                        className="vtxt mx-auto font-bold"
                        style={{
                          fontSize: '13px',
                          fontWeight: 700,
                          minHeight: '88px',
                          textAlign: 'center',
                        }}
                      >
                        ಸೆಮಿಸ್ಟರ್
                      </div>
                    </th>
                    <th
                      rowSpan={3}
                      style={{
                        background: '#fbcfe5',
                        border: '1px solid #000',
                        padding: '2px',
                        verticalAlign: 'middle',
                        fontSize: '13px',
                        textAlign: 'center',
                      }}
                    >
                      <div
                        className="vtxt mx-auto font-bold"
                        style={{
                          fontSize: '13px',
                          fontWeight: 700,
                          minHeight: '88px',
                          lineHeight: '1.1',
                          textAlign: 'center',
                        }}
                      >
                        ಶಾಲೆ ನಡೆದ ದಿನ
                      </div>
                    </th>
                    <th
                      rowSpan={3}
                      style={{
                        background: '#fbcfe5',
                        border: '1px solid #000',
                        padding: '2px',
                        verticalAlign: 'middle',
                        textAlign: 'center',
                      }}
                    >
                      <div
                        className="vtxt mx-auto font-bold"
                        style={{
                          fontSize: '13px',
                          fontWeight: 700,
                          minHeight: '88px',
                          lineHeight: '1.1',
                          textAlign: 'center',
                        }}
                      >
                        ಹಾಜರಾದ ದಿನ
                      </div>
                    </th>
                    <th
                      rowSpan={3}
                      style={{
                        background: '#fbcfe5',
                        border: '1px solid #000',
                        padding: '2px',
                        verticalAlign: 'middle',
                        textAlign: 'center',
                      }}
                    >
                      <div
                        className="vtxt mx-auto font-bold"
                        style={{
                          fontSize: '11px',
                          fontWeight: 700,
                          minHeight: '88px',
                          textAlign: 'center',
                        }}
                      >
                        ಮೌಲ್ಯಾಂಕನ
                      </div>
                    </th>
                    <th
                      rowSpan={3}
                      style={{
                        background: '#fbcfe5',
                        border: '1px solid #000',
                        padding: '2px',
                        verticalAlign: 'middle',
                        textAlign: 'center',
                      }}
                    >
                      <div
                        className="vtxt mx-auto font-bold"
                        style={{
                          fontSize: '11px',
                          fontWeight: 700,
                          minHeight: '88px',
                          textAlign: 'center',
                        }}
                      >
                        ಶೇ. ಪ್ರಮಾಣ
                      </div>
                    </th>
                    <th
                      colSpan={curSubjects.length * 2 + 3}
                      title={gradeTooltipText}
                      style={{
                        background: '#dbeafe',
                        border: '1px solid #000',
                        fontSize: '15px',
                        fontWeight: 700,
                        height: '34px',
                        textAlign: 'center',
                        verticalAlign: 'middle',
                        cursor: 'help',
                      }}
                    >
                      ವಿದ್ಯಾರ್ಥಿಯ ಶೈಕ್ಷಣಿಕ ಸಾಧನೆ - ಪಠ್ಯ ಭಾಗಕ್ಕೆ ಸಂಬಂಧಿಸಿದಂತೆ
                    </th>
                    <th
                      colSpan={curCoSkills.length}
                      style={{
                        background: '#c0facd',
                        border: '1px solid #000',
                        fontSize: '15px',
                        fontWeight: 700,
                        height: '34px',
                        textAlign: 'center',
                        verticalAlign: 'middle',
                        lineHeight: '1.1',
                        whiteSpace: 'nowrap',
                      }}
                    >
                      ಸಹಪಠ್ಯ ವಿಷಯದ ಶ್ರೇಣಿಗಳು
                    </th>
                    <th
                      rowSpan={3}
                      style={{
                        background: '#8af6f8',
                        border: '1px solid #000',
                        padding: '2px',
                        verticalAlign: 'middle',
                        textAlign: 'center',
                      }}
                    >
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '11px', fontWeight: 700, minHeight: '88px' }}>
                        ಫಲಿತಾಂಶ
                      </div>
                    </th>
                  </tr>

                  {/* Header Row 2 */}
                  <tr style={{ height: '42px' }}>
                    {curSubjects.map((sub) => (
                      <th
                        key={sub}
                        colSpan={2}
                        style={{
                          background: '#dbeafe',
                          border: '1px solid #000',
                          fontSize: '13px',
                          fontWeight: 700,
                          textAlign: 'center',
                          verticalAlign: 'middle',
                          padding: '2px 1px',
                        }}
                      >
                        {renderSubjectHeader(sub, false, languageNames, handleLanguageChange)}
                      </th>
                    ))}
                    <th
                      colSpan={3}
                      title={gradeTooltipText}
                      style={{
                        background: '#dbeafe',
                        border: '1px solid #000',
                        fontSize: '14px',
                        fontWeight: 700,
                        textAlign: 'center',
                        verticalAlign: 'middle',
                        padding: '4px 2px',
                        cursor: 'help',
                      }}
                    >
                      ಒಟ್ಟು
                    </th>
                    {curCoSkills.map((co) => (
                      <th
                        key={co}
                        rowSpan={2}
                        style={{
                          background: '#c0facd',
                          border: '1px solid #000',
                          padding: '2px 1px',
                          verticalAlign: 'middle',
                          textAlign: 'center',
                        }}
                      >
                        {renderCoSkillHeader(co, false)}
                      </th>
                    ))}
                  </tr>

                  {/* Header Row 3 */}
                  <tr style={{ height: '30px' }}>
                    {curSubjects.map((sub, i) => (
                      <React.Fragment key={`${sub}-${i}`}>
                        <th
                          className="anka-th"
                          style={{
                            background: '#eff6ff',
                            border: '1px solid #000',
                            fontSize: '12px',
                            fontWeight: 700,
                            textAlign: 'center',
                            verticalAlign: 'middle',
                            padding: '3px 1px',
                          }}
                        >
                          ಅಂಕ
                        </th>
                        <th
                          title={gradeTooltipText}
                          style={{
                            background: '#eff6ff',
                            border: '1px solid #000',
                            fontSize: '12px',
                            fontWeight: 700,
                            textAlign: 'center',
                            verticalAlign: 'middle',
                            padding: '3px 1px',
                            cursor: 'help',
                          }}
                        >
                          ಶ್ರೇಣಿ
                        </th>
                      </React.Fragment>
                    ))}
                <th
                  className="anka-th"
                  style={{
                    background: '#eff6ff',
                    border: '1px solid #000',
                    fontSize: '12px',
                    fontWeight: 700,
                    textAlign: 'center',
                    verticalAlign: 'middle',
                    padding: '3px 1px',
                  }}
                >
                  ಅಂಕ
                </th>
                <th
                  className="anka-th"
                  style={{
                    background: '#eff6ff',
                    border: '1px solid #000',
                    fontSize: '12px',
                    fontWeight: 700,
                    textAlign: 'center',
                    verticalAlign: 'middle',
                    padding: '3px 1px',
                  }}
                >
                  ಶೇ.
                </th>
                <th
                  title={gradeTooltipText}
                  style={{
                    background: '#eff6ff',
                    border: '1px solid #000',
                    fontSize: '12px',
                    fontWeight: 700,
                    textAlign: 'center',
                    verticalAlign: 'middle',
                    padding: '3px 1px',
                    cursor: 'help',
                  }}
                >
                  ಶ್ರೇಣಿ
                </th>
              </tr>
            </thead>

            <tbody>
              {filteredStudentList.length === 0 ? (
                <tr>
                  <td colSpan={28} className="py-12 text-center text-gray-600 font-bold bg-amber-50/50">
                    <div className="flex flex-col items-center justify-center gap-2">
                      <Search className="w-8 h-8 text-gray-400 stroke-[2]" />
                      <span className="text-[16px] text-gray-800">
                        "{searchQuery}" ಗೆ ಯಾವುದೇ ವಿದ್ಯಾರ್ಥಿ ಕಂಡುಬಂದಿಲ್ಲ
                      </span>
                      <button
                        type="button"
                        onClick={() => setSearchQuery('')}
                        className="mt-1 text-xs px-3.5 py-1.5 bg-blue-600 text-white rounded-md hover:bg-blue-700 font-bold cursor-pointer shadow-xs transition-colors"
                      >
                        ಹುಡುಕಾಟ ತೆರವುಗೊಳಿಸಿ (Clear Search)
                      </button>
                    </div>
                  </td>
                </tr>
              ) : (
                filteredStudentList.map(({ stu, originalIdx }) => {
                  const sIdx = originalIdx;
                  return stu.assessments.map((ass, aIdx) => {
                    const is50 = !!ass.is50;
                    const is100 = !!ass.is100 || aIdx === 8;
                    const totDays = (stu.attendance.s1Total || 0) + (stu.attendance.s2Total || 0);
                    const totPres = (stu.attendance.s1Present || 0) + (stu.attendance.s2Present || 0);
                    const isCalc = CALCULATED_ASSESSMENT_INDICES.includes(aIdx);
                    const isInteractingWithStudent = hoveredStudentIdx === sIdx;
                    const bgRow = is100 ? '#e5e7eb' : is50 ? '#f3f4f6' : '#ffffff';
                    const bgLabel = '#FCE7F3';
                    const bgTotal = is100 ? '#e5e7eb' : is50 ? '#f3f4f6' : isCalc ? '#e5e7eb' : bgRow;
                    const bgCell = is100 ? '#e5e7eb' : is50 ? '#f3f4f6' : isCalc ? '#f3f4f6' : bgRow;
                    const bgGradeCell = is100 ? '#e5e7eb' : is50 ? '#f3f4f6' : '#f3f4f6';

                    return (
                      <tr
                        key={`${sIdx}-${aIdx}`}
                        className="student-data-row interactive-row transition-colors"
                        onMouseEnter={() => setHoveredStudentIdx(sIdx)}
                        onMouseLeave={() => setHoveredStudentIdx((prev) => (prev === sIdx ? null : prev))}
                        onFocus={() => setHoveredStudentIdx(sIdx)}
                        onBlur={(e) => {
                          const currentTarget = e.currentTarget;
                          requestAnimationFrame(() => {
                            if (!currentTarget.contains(document.activeElement)) {
                              setHoveredStudentIdx((prev) => (prev === sIdx ? null : prev));
                            }
                          });
                        }}
                        style={{
                          height: '36px',
                          background: is100 ? '#e5e7eb' : is50 ? '#f3f4f6' : isCalc ? '#eef2ff' : bgRow,
                        }}
                      >
                        {/* Student ID */}
                        {aIdx === 0 && (
                          <td
                            rowSpan={9}
                            className="text-center font-bold"
                            style={{
                              verticalAlign: 'middle',
                              background: '#f4eae0',
                              border: '1px solid #000',
                              fontSize: '14.5px',
                              padding: '4px 1px',
                            }}
                          >
                            {stu.id}
                          </td>
                        )}

                        {/* Student Info Box */}
                        {aIdx === 0 && (
                          <td
                            rowSpan={9}
                            className="p-0 stu-cell"
                            style={{
                              verticalAlign: 'middle',
                              background: '#f4eae0',
                              border: '1px solid #000',
                              padding: '0',
                              height: '324px',
                              boxSizing: 'border-box',
                              overflow: 'hidden',
                            }}
                          >
                          <div className="stu-detail-wrap" style={{ minHeight: '324px', width: '100%' }}>
                            <div className="stu-row relative group/name" style={{ minHeight: '44px' }}>
                              <input
                                className="stu-input stu-input-name font-extrabold pr-7"
                                style={{
                                  fontSize: '16.5px',
                                  fontWeight: 800,
                                  padding: '4px 28px 4px 4px',
                                  height: '38px',
                                  color: '#000',
                                }}
                                value={stu.name}
                                onChange={(e) => updateStudentField(sIdx, 'name', e.target.value)}
                                placeholder="ವಿದ್ಯಾರ್ಥಿಯ ಹೆಸರು"
                                aria-label="ವಿದ್ಯಾರ್ಥಿಯ ಹೆಸರು"
                                title="ವಿದ್ಯಾರ್ಥಿಯ ಹೆಸರು"
                              />
                              <button
                                type="button"
                                onClick={() => handleRequestDeleteStudent(sIdx)}
                                className={`absolute right-1 top-1/2 -translate-y-1/2 ${
                                  isInteractingWithStudent ? 'opacity-100 scale-100' : 'opacity-0 scale-90'
                                } group-hover/name:opacity-100 group-focus-within/name:opacity-100 hover:bg-red-600 hover:text-white text-red-600 bg-red-100 hover:scale-105 active:scale-95 rounded p-1 transition-all duration-150 cursor-pointer shadow-xs no-print flex items-center justify-center`}
                                title={`${stu.name ? stu.name : `ವಿದ್ಯಾರ್ಥಿ #${stu.id}`} ತೆಗೆದುಹಾಕಿ (Delete)`}
                                aria-label="ವಿದ್ಯಾರ್ಥಿ ತೆಗೆದುಹಾಕಿ"
                              >
                                <Trash2 className="w-3.5 h-3.5 stroke-[2.5]" />
                              </button>
                            </div>
                            <div className="stu-row">
                              <span className="stu-label">ದಾ . ಸಂ:</span>
                              <input
                                className="stu-input"
                                value={stu.admission}
                                onChange={(e) => updateStudentField(sIdx, 'admission', e.target.value)}
                                placeholder="ದಾಖಲಾತಿ ಸಂಖ್ಯೆ"
                              />
                            </div>
                            <div className="stu-row">
                              <span className="stu-label">ಲಿಂಗ:</span>
                              <select
                                className={`stu-input cursor-pointer ${!stu.gender ? 'text-gray-500 font-semibold' : 'text-black font-bold'}`}
                                value={normalizeGender(stu.gender)}
                                onChange={(e) => updateStudentField(sIdx, 'gender', e.target.value)}
                                style={{
                                  background: 'transparent',
                                  fontSize: '13.5px',
                                  padding: '2px 4px',
                                  color: !stu.gender ? '#6b7280' : '#000000',
                                  fontWeight: !stu.gender ? 600 : 700,
                                }}
                                title="ವಿದ್ಯಾರ್ಥಿಯ ಲಿಂಗ (Gender)"
                              >
                                <option value="" style={{ color: '#6b7280', fontWeight: 600 }}>
                                  ಆಯ್ಕೆ (ಲಿಂಗ)
                                </option>
                                <option value="ಗಂಡು" style={{ color: '#000000', fontWeight: 700 }}>
                                  ಗಂಡು
                                </option>
                                <option value="ಹೆಣ್ಣು" style={{ color: '#000000', fontWeight: 700 }}>
                                  ಹೆಣ್ಣು
                                </option>
                                <option value="ಇತರೆ" style={{ color: '#000000', fontWeight: 700 }}>
                                  ಇತರೆ
                                </option>
                              </select>
                            </div>
                            <div className="stu-row">
                              <span className="stu-label">ಜನ್ಮ ದಿನಾಂಕ:</span>
                              <input
                                type="date"
                                className="stu-input-date"
                                value={stu.dob}
                                onChange={(e) => updateStudentField(sIdx, 'dob', e.target.value)}
                                title="ಒಂದು ಕ್ಲಿಕ್ ನಲ್ಲಿ ಕ್ಯಾಲೆಂಡರ್ ತೆರೆಯುತ್ತದೆ"
                              />
                            </div>
                            <div className="stu-row">
                              <span className="stu-label">SATS No:</span>
                              <input
                                className="stu-input"
                                value={stu.sats}
                                onChange={(e) => updateStudentField(sIdx, 'sats', e.target.value)}
                                placeholder="123456789"
                              />
                            </div>
                            <div className="stu-row">
                              <span className="stu-label">PEN No:</span>
                              <input
                                className="stu-input"
                                value={stu.pen}
                                onChange={(e) => updateStudentField(sIdx, 'pen', e.target.value)}
                                placeholder="12345678901"
                              />
                            </div>
                          </div>
                        </td>
                      )}

                      {/* Semester 1 & 2 Badges */}
                      {aIdx === 0 && (
                        <td
                          rowSpan={4}
                          className="sem-cell"
                          style={{
                            verticalAlign: 'middle',
                            background: '#FCE7F3',
                            border: '1px solid #000',
                            textAlign: 'center',
                            padding: '3px 1px',
                            boxSizing: 'border-box',
                          }}
                        >
                          <div
                            className="vtxt mx-auto font-bold"
                            style={{ width: '32px', height: '100px', fontSize: '14.5px', fontWeight: 700 }}
                          >
                            1ನೇ
                          </div>
                        </td>
                      )}
                      {aIdx === 4 && (
                        <td
                          rowSpan={5}
                          className="sem-cell"
                          style={{
                            verticalAlign: 'middle',
                            background: '#FCE7F3',
                            border: '1px solid #000',
                            textAlign: 'center',
                            padding: '3px 1px',
                            boxSizing: 'border-box',
                          }}
                        >
                          <div
                            className="vtxt mx-auto font-bold"
                            style={{ width: '32px', height: '140px', fontSize: '14.5px', fontWeight: 700 }}
                          >
                            2ನೇ
                          </div>
                        </td>
                      )}

                      {/* Attendance Inputs: Sem 1 & Sem 2 */}
                      {aIdx === 0 && (
                        <td
                          rowSpan={4}
                          className="attendance-cell"
                          style={{
                            verticalAlign: 'middle',
                            background: '#FCE7F3',
                            border: '1px solid #000',
                            textAlign: 'center',
                            fontSize: '12.5px',
                            padding: '2px 1px',
                          }}
                        >
                          <input
                            className="w-full h-full min-h-[140px] text-center bg-transparent border-0 p-0 text-[12.5px] font-bold"
                            style={{ color: '#000', fontSize: '12.5px', fontWeight: 700 }}
                            value={stu.attendance.s1Total || ''}
                            placeholder="-"
                            onChange={(e) => updateAttendance(sIdx, 's1Total', e.target.value)}
                          />
                        </td>
                      )}
                      {aIdx === 4 && (
                        <td
                          rowSpan={4}
                          className="attendance-cell"
                          style={{
                            verticalAlign: 'middle',
                            background: '#FCE7F3',
                            border: '1px solid #000',
                            textAlign: 'center',
                            fontWeight: 700,
                            fontSize: '12.5px',
                            padding: '2px 1px',
                          }}
                        >
                          <input
                            className="w-full h-full min-h-[140px] text-center bg-transparent border-0 p-0 text-[12.5px] font-bold"
                            style={{ color: '#000', fontSize: '12.5px', fontWeight: 700 }}
                            value={stu.attendance.s2Total || ''}
                            placeholder="-"
                            onChange={(e) => updateAttendance(sIdx, 's2Total', e.target.value)}
                          />
                        </td>
                      )}

                      {aIdx === 0 && (
                        <td
                          rowSpan={4}
                          className="attendance-cell"
                          style={{
                            verticalAlign: 'middle',
                            background: '#FCE7F3',
                            border: '1px solid #000',
                            textAlign: 'center',
                            padding: '2px 1px',
                            fontSize: '12.5px',
                          }}
                        >
                          <input
                            className="w-full h-full min-h-[140px] text-center bg-transparent border-0 p-0 text-[12.5px] font-bold"
                            style={{ color: '#000', fontSize: '12.5px', fontWeight: 700 }}
                            value={stu.attendance.s1Present || ''}
                            placeholder="-"
                            onChange={(e) => updateAttendance(sIdx, 's1Present', e.target.value)}
                          />
                        </td>
                      )}
                      {aIdx === 4 && (
                        <td
                          rowSpan={4}
                          className="attendance-cell"
                          style={{
                            verticalAlign: 'middle',
                            background: '#FCE7F3',
                            border: '1px solid #000',
                            textAlign: 'center',
                            fontWeight: 700,
                            fontSize: '12.5px',
                            padding: '2px 1px',
                          }}
                        >
                          <input
                            className="w-full h-full min-h-[140px] text-center bg-transparent border-0 p-0 text-[12.5px] font-bold"
                            style={{ color: '#000', fontSize: '12.5px', fontWeight: 700 }}
                            value={stu.attendance.s2Present || ''}
                            placeholder="-"
                            onChange={(e) => updateAttendance(sIdx, 's2Present', e.target.value)}
                          />
                        </td>
                      )}

                      {/* Final Row: Total Attendance Sum */}
                      {aIdx === 8 && (
                        <React.Fragment>
                          <td
                            className="attendance-cell"
                            style={{
                              verticalAlign: 'middle',
                              background: '#FCE7F3',
                              border: '1px solid #000',
                              textAlign: 'center',
                              fontWeight: 700,
                              fontSize: '11.5px',
                              color: '#000',
                              padding: '2px 1px',
                              height: '36px',
                            }}
                          >
                            <div
                              style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', width: '100%', height: '100%', minHeight: '36px', background: '#FCE7F3' }}
                              title={`ಶಾಲೆ ನಡೆದ ಒಟ್ಟು: ${stu.attendance.s1Total} + ${stu.attendance.s2Total} = ${totDays}`}
                            >
                              {totDays || '-'}
                            </div>
                          </td>
                          <td
                            className="attendance-cell"
                            style={{
                              verticalAlign: 'middle',
                              background: '#FCE7F3',
                              border: '1px solid #000',
                              textAlign: 'center',
                              fontWeight: 700,
                              fontSize: '11.5px',
                              color: '#000',
                              padding: '2px 1px',
                              height: '36px',
                            }}
                          >
                            <div
                              style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', width: '100%', height: '100%', minHeight: '36px', background: '#FCE7F3' }}
                              title={`ಹಾಜರಾದ ಒಟ್ಟು: ${stu.attendance.s1Present} + ${stu.attendance.s2Present} = ${totPres}`}
                            >
                              {totPres || '-'}
                            </div>
                          </td>
                        </React.Fragment>
                      )}

                      {/* Assessment Label Cell (ಮೌಲ್ಯಾಂಕನ) */}
                      <td
                        className={`moulankan-cell ${is50 || is100 ? '' : 'fa-cell'}`}
                        style={{
                          background: bgLabel,
                          border: '1px solid #000',
                          textAlign: 'center',
                          verticalAlign: 'middle',
                          padding: '3px 1px',
                          fontWeight: 700,
                          fontSize: '12px',
                          whiteSpace: 'nowrap',
                          width: '42px',
                          minWidth: '42px',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '30px', lineHeight: '1.2', whiteSpace: 'nowrap', fontSize: '12px', fontWeight: 700 }}>
                          {getAssessmentPartsForClass(selectedClass)[aIdx]?.name || ass.label}
                        </div>
                      </td>

                      {/* Assessment Pct Cell (ಶೇ. ಪ್ರಮಾಣ) */}
                      <td
                        className={`moulankan-cell ${is50 || is100 ? '' : 'fa-cell'}`}
                        style={{
                          background: bgLabel,
                          border: '1px solid #000',
                          textAlign: 'center',
                          verticalAlign: 'middle',
                          padding: '3px 1px',
                          fontWeight: 700,
                          fontSize: '11px',
                          whiteSpace: 'nowrap',
                          width: '36px',
                          minWidth: '36px',
                        }}
                      >
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: '30px', lineHeight: '1.2', whiteSpace: 'nowrap', fontSize: '11px', fontWeight: 700 }}>
                          {getAssessmentPartsForClass(selectedClass)[aIdx]?.pct || ''}
                        </div>
                      </td>

                      {/* Subject Marks and Grades */}
                      {ass.subs.map((sub, subIdx) => (
                        <React.Fragment key={`${sIdx}-${aIdx}-${subIdx}`}>
                          <td
                            className="fa-cell sa-cell"
                            style={{
                              background: bgCell,
                              border: '1px solid #000',
                              padding: '1px',
                              verticalAlign: 'middle',
                              textAlign: 'center',
                              opacity: isCalc && !is100 ? 0.9 : 1,
                            }}
                          >
                            <input
                              readOnly={isCalc}
                              tabIndex={isCalc ? -1 : 0}
                              type="text"
                              inputMode="numeric"
                              maxLength={2}
                              className={`w-full h-[32px] text-center border-0 p-0 text-[15.5px] anka-mark-input ${isCalc ? 'font-bold' : 'font-normal'}`}
                                style={{
                                color: '#000',
                                background: is100 ? '#e5e7eb' : is50 ? '#f3f4f6' : isCalc ? '#f3f4f6' : 'transparent',
                                cursor: isCalc ? 'not-allowed' : 'text',
                                fontSize: '15.5px',
                                fontWeight: isCalc ? 700 : 400,
                              }}
                              value={sub.m}
                              placeholder="-"
                              onChange={(e) => updateMark(sIdx, aIdx, subIdx, e.target.value)}
                              title={isCalc ? 'ಸ್ವಯಂಚಾಲಿತ ಲೆಕ್ಕಾಚಾರ - FA/SA ಅಂಕಗಳಿಂದ' : `ಗರಿಷ್ಠ ${MAX_PER_SUB[aIdx]} ಅಂಕ`}
                            />
                          </td>
                          <td
                            className="fa-cell sa-cell"
                            style={{
                              background: bgGradeCell,
                              border: '1px solid #000',
                              padding: '1px',
                              verticalAlign: 'middle',
                              textAlign: 'center',
                              opacity: 0.95,
                            }}
                          >
                            <input
                              readOnly
                              tabIndex={-1}
                              className={`w-full h-[32px] text-center border-0 p-0 text-[15.5px] ${isCalc ? 'font-extrabold' : 'font-normal'}`}
                              style={{
                                color: '#000',
                                background: 'transparent',
                                cursor: 'not-allowed',
                                fontSize: '15.5px',
                                fontWeight: isCalc ? 800 : 400,
                              }}
                              value={sub.g || ''}
                              placeholder="-"
                              title="ಶ್ರೇಣಿ ಸ್ವಯಂಚಾಲಿತ"
                            />
                          </td>
                        </React.Fragment>
                      ))}

                      {/* Assessment Totals */}
                      <td
                        className="fa-cell sa-cell anka-total-cell"
                        style={{
                          background: bgCell,
                          border: '1px solid #000',
                          padding: '1px',
                          verticalAlign: 'middle',
                          textAlign: 'center',
                          opacity: isCalc && !is100 ? 0.9 : 1,
                        }}
                      >
                        <input
                          readOnly
                          tabIndex={-1}
                          className="w-full h-[32px] text-center border-0 p-0 text-[15.5px] font-bold anka-mark-input"
                          style={{
                            color: '#000',
                            background: 'transparent',
                            cursor: 'not-allowed',
                            fontSize: '15.5px',
                            fontWeight: 700,
                          }}
                          value={ass.total.m !== '' && ass.total.m !== undefined ? ass.total.m : ''}
                          placeholder="-"
                          title="ಒಟ್ಟು ಅಂಕ ಸ್ವಯಂಚಾಲಿತ"
                        />
                      </td>
                      <td
                        className="fa-cell sa-cell"
                        style={{
                          background: bgCell,
                          border: '1px solid #000',
                          padding: '1px',
                          verticalAlign: 'middle',
                          textAlign: 'center',
                          opacity: isCalc && !is100 ? 0.9 : 1,
                        }}
                      >
                        <input
                          readOnly
                          tabIndex={-1}
                          className="w-full h-[32px] text-center border-0 p-0 text-[15.5px] font-bold anka-mark-input"
                          style={{
                            color: '#000',
                            background: 'transparent',
                            cursor: 'not-allowed',
                            fontSize: '15.5px',
                            fontWeight: 700,
                          }}
                          value={ass.total.p !== '' && ass.total.p !== undefined ? ass.total.p : ''}
                          placeholder="-"
                          title="% ಸ್ವಯಂಚಾಲಿತ"
                        />
                      </td>
                      <td
                        className="fa-cell sa-cell"
                        style={{
                          background: bgGradeCell,
                          border: '1px solid #000',
                          padding: '1px',
                          verticalAlign: 'middle',
                          textAlign: 'center',
                          opacity: 0.95,
                        }}
                      >
                        <input
                          readOnly
                          tabIndex={-1}
                          className={`w-full h-[32px] text-center border-0 p-0 text-[15.5px] ${isCalc ? 'font-extrabold' : 'font-bold'}`}
                          style={{
                            color: '#000',
                            background: 'transparent',
                            cursor: 'not-allowed',
                            fontSize: '15.5px',
                            fontWeight: isCalc ? 800 : 700,
                          }}
                          value={ass.total.g || ''}
                          placeholder="-"
                          title="ಶ್ರೇಣಿ ಸ್ವಯಂಚಾಲಿತ"
                        />
                      </td>

                      {/* Co-scholastic Sem 1 */}
                      {aIdx === 0 &&
                        stu.co1.slice(0, curCoSkills.length).map((coVal, coIdx) => (
                          <td
                            key={`co1-${sIdx}-${coIdx}`}
                            rowSpan={4}
                            className="co-cell"
                            style={{ verticalAlign: 'middle', border: '1px solid #000', padding: '1px' }}
                          >
                            <div className="co-wrap">
                              <select
                                className="co-select-direct font-bold"
                                value={coVal}
                                onChange={(e) => updateCoGrade(sIdx, 'co1', coIdx, e.target.value)}
                                aria-label={`Co1 ${curCoSkills[coIdx] || ''}`}
                              >
                                <option value="">-</option>
                                <option value="A">A</option>
                                <option value="B">B</option>
                              </select>
                            </div>
                          </td>
                        ))}

                      {/* Co-scholastic Sem 2 */}
                      {aIdx === 4 &&
                        stu.co2.slice(0, curCoSkills.length).map((coVal, coIdx) => (
                          <td
                            key={`co2-${sIdx}-${coIdx}`}
                            rowSpan={4}
                            className="co-cell-2"
                            style={{ verticalAlign: 'middle', border: '1px solid #000', padding: '1px' }}
                          >
                            <div className="co-wrap">
                              <select
                                className="co-select-direct font-bold"
                                value={coVal}
                                onChange={(e) => updateCoGrade(sIdx, 'co2', coIdx, e.target.value)}
                                aria-label={`Co2 ${curCoSkills[coIdx] || ''}`}
                              >
                                <option value="">-</option>
                                <option value="A">A</option>
                                <option value="B">B</option>
                              </select>
                            </div>
                          </td>
                        ))}

                      {/* Co-scholastic Final Row */}
                      {aIdx === 8 &&
                        curCoSkills.map((_, coIdx) => {
                          const c1 = stu.co1[coIdx] || '';
                          const c2 = stu.co2[coIdx] || '';
                          const finalCo = !c1 && !c2 ? '-' : c1 === 'A' && c2 === 'A' ? 'A' : 'B';
                          return (
                            <td
                              key={`final-co-${sIdx}-${coIdx}`}
                              style={{
                                verticalAlign: 'middle',
                                border: '1px solid #000',
                                background: '#c0facd',
                                textAlign: 'center',
                                padding: '2px 1px',
                                height: '36px',
                              }}
                            >
                              <div
                                className="co-final font-bold"
                                title={finalCo !== '-' ? `Final: ${c1} + ${c2} = ${finalCo} (A+A=A else B)` : 'ಖಾಲಿ'}
                              >
                                {finalCo}
                              </div>
                            </td>
                          );
                        })}

                      {/* Overall Result Banner */}
                      {aIdx === 0 &&
                        (() => {
                          const grandMarks = toInt(stu.assessments[8]?.total?.m);
                          const { maxGrand } = getMaxMarksConfig(selectedClass || '6');
                          const pct = grandMarks ? calcPercentExact(grandMarks, maxGrand) : 0;
                          const pctRound = grandMarks ? Math.round(pct) : 0;
                          const gr = grandMarks ? getGrade(pct, gradeRules) : '';
                          const resLabel = stu.result || (grandMarks ? 'ಉತ್ತೀರ್ಣ' : '');
                          const resText = resLabel
                            ? grandMarks
                              ? `${resLabel} = ( ${pctRound}% - ${gr || 'C'} )`
                              : resLabel
                            : '-';
                          return (
                            <td
                              rowSpan={9}
                              className="result-cell"
                              style={{
                                verticalAlign: 'middle',
                                background: '#cdf9fb',
                                border: '1px solid #000',
                                textAlign: 'center',
                                padding: '2px 0',
                                width: '36px',
                                minWidth: '36px',
                                maxWidth: '36px',
                                height: '324px',
                              }}
                            >
                              <div
                                style={{
                                  writingMode: 'vertical-rl',
                                  transform: 'rotate(180deg)',
                                  textOrientation: 'mixed',
                                  whiteSpace: 'nowrap',
                                  display: 'flex',
                                  alignItems: 'center',
                                  justifyContent: 'center',
                                  width: '32px',
                                  height: '100%',
                                  minHeight: '324px',
                                  fontSize: '15.5px',
                                  fontWeight: 900,
                                  background: '#cdf9fb',
                                  color: '#000',
                                  lineHeight: 1.1,
                                  letterSpacing: '0.2px',
                                  margin: '0 auto',
                                }}
                                title={resText}
                              >
                                {resText}
                              </div>
                            </td>
                          );
                        })()}
                    </tr>
                  );
                });
              })
            )}
            </tbody>
          </table>
        </div>
      );
    })()}
  </div>

      {/* Calibrated 100% A4 landscape register (Visible only during print) */}
      <div id="printableArea" className="hidden print:block w-full bg-white text-black p-0 m-0">
        {activeClassKeys.map((clsKey, index) => {
          const classData = getClassData(clsKey);
          return (
            <div
              key={`printable-class-${clsKey}`}
              style={{
                pageBreakBefore: index > 0 ? 'always' : 'auto',
                breakBefore: index > 0 ? 'page' : 'auto',
              }}
            >
              <PrintableRegisterContent
                students={classData.students}
                year={year}
                selectedClass={classData.cls}
                school={school}
                diseCode={diseCode}
                gradeRules={gradeRules}
                languageNames={classData.languageNames}
              />
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PrintableRegisterContent({
  students,
  year,
  selectedClass,
  school,
  diseCode = '',
  gradeRules = DEFAULT_GRADE_RULES,
  languageNames,
}: {
  students: Student[];
  year: string;
  selectedClass: string;
  school: string;
  diseCode?: string;
  gradeRules?: GradeItemRule[];
  languageNames?: { lang1: string; lang2: string; lang3: string };
}) {
  const subjects = getSubjectsForClass(selectedClass);
  const coSkills = getCoSkillsForClass(selectedClass);
  const parts = getAssessmentPartsForClass(selectedClass);
  const { maxGrand } = getMaxMarksConfig(selectedClass);

  const gradeTooltipText = gradeRules
    .map((r) => `${r.grade}: ${r.text100 || `${r.min100}%+`}`)
    .join(', ');

  // Only show students who have data filled in print view (omit blank/unfilled placeholder rows)
  const studentsToPrint = React.useMemo(() => {
    const filled = students.filter(isStudentFilled);
    return filled.length > 0 ? filled : students;
  }, [students]);

  // Auto-calculate statistics for the summary table
  const summaryStats = React.useMemo(() => {
    const listToCount = studentsToPrint.filter(isStudentFilled).length > 0
      ? studentsToPrint.filter(isStudentFilled)
      : studentsToPrint;

    const isBoy = (s: Student) => {
      const g = (s.gender || '').trim().toLowerCase();
      return g === 'ಗಂಡು' || g === 'boy' || g === 'boys' || g === 'm' || g === 'male';
    };

    const isGirl = (s: Student) => {
      const g = (s.gender || '').trim().toLowerCase();
      return g === 'ಹೆಣ್ಣು' || g === 'girl' || g === 'girls' || g === 'f' || g === 'female';
    };

    const hasAppeared = (s: Student) => {
      const totPres = (s.attendance?.s1Present || 0) + (s.attendance?.s2Present || 0);
      const grandMarks = toInt(s.assessments?.[8]?.total?.m);
      const hasMarks = s.assessments?.some((a) => a.subs?.some((sub) => sub.m !== ''));
      return totPres > 0 || grandMarks > 0 || hasMarks;
    };

    const isPassed = (s: Student) => {
      if (s.result === 'ಉತ್ತೀರ್ಣ' || s.result === 'Pass' || s.result === 'PASSED') return true;
      if (s.result === 'ಅನುತ್ತೀರ್ಣ' || s.result === 'Fail' || s.result === 'FAILED') return false;
      const grandMarks = toInt(s.assessments?.[8]?.total?.m);
      if (!grandMarks) return false;
      const pct = calcPercentExact(grandMarks, maxGrand);
      return pct >= 35;
    };

    const isFailed = (s: Student) => {
      if (s.result === 'ಅನುತ್ತೀರ್ಣ' || s.result === 'Fail' || s.result === 'FAILED') return true;
      if (s.result === 'ಉತ್ತೀರ್ಣ' || s.result === 'Pass' || s.result === 'PASSED') return false;
      const grandMarks = toInt(s.assessments?.[8]?.total?.m);
      if (!grandMarks) return false;
      const pct = calcPercentExact(grandMarks, maxGrand);
      return pct < 35;
    };

    const totalBoys = listToCount.filter(isBoy).length;
    const totalGirls = listToCount.filter(isGirl).length;
    const totalStudents = listToCount.length;

    const appearedBoys = listToCount.filter((s) => isBoy(s) && hasAppeared(s)).length;
    const appearedGirls = listToCount.filter((s) => isGirl(s) && hasAppeared(s)).length;
    const appearedTotal = listToCount.filter(hasAppeared).length;

    const passedBoys = listToCount.filter((s) => isBoy(s) && isPassed(s)).length;
    const passedGirls = listToCount.filter((s) => isGirl(s) && isPassed(s)).length;
    const passedTotal = listToCount.filter(isPassed).length;

    const failedBoys = listToCount.filter((s) => isBoy(s) && isFailed(s)).length;
    const failedGirls = listToCount.filter((s) => isGirl(s) && isFailed(s)).length;
    const failedTotal = listToCount.filter(isFailed).length;

    const passPctBoys = appearedBoys > 0 ? `${((passedBoys / appearedBoys) * 100).toFixed(1)}%` : '0%';
    const passPctGirls = appearedGirls > 0 ? `${((passedGirls / appearedGirls) * 100).toFixed(1)}%` : '0%';
    const passPctTotal = appearedTotal > 0 ? `${((passedTotal / appearedTotal) * 100).toFixed(1)}%` : '0%';

    return {
      totalBoys,
      totalGirls,
      totalStudents,
      appearedBoys,
      appearedGirls,
      appearedTotal,
      passedBoys,
      passedGirls,
      passedTotal,
      failedBoys,
      failedGirls,
      failedTotal,
      passPctBoys,
      passPctGirls,
      passPctTotal,
    };
  }, [studentsToPrint]);

  // Chunk students into groups of maximum 3 students per page
  const chunkSize = 3;
  const studentPages: Student[][] = [];
  if (studentsToPrint.length === 0) {
    studentPages.push([]);
  } else {
    for (let i = 0; i < studentsToPrint.length; i += chunkSize) {
      studentPages.push(studentsToPrint.slice(i, i + chunkSize));
    }
  }

  return (
    <div className="w-full">
      {studentPages.map((pageStudents, pageIdx) => {
        return (
          <div
            key={`print-page-${pageIdx}`}
            className={`print-page ${is1to5Class(selectedClass) ? 'print-page-1to5' : 'print-page-6to9'} w-full text-black bg-white`}
            style={{
              pageBreakAfter: pageIdx < studentPages.length - 1 ? 'always' : 'auto',
              breakAfter: pageIdx < studentPages.length - 1 ? 'page' : 'auto',
              boxSizing: 'border-box' as const,
              padding: is1to5Class(selectedClass) ? '8mm 10mm 5mm 10mm' : '8mm 10mm 10mm 10mm',
            }}
          >
            {/* Main Title and School / Class / DISE / Year Header */}
            <div className="print-page-header mb-1 text-black">
              <div className="text-center font-bold text-[13px] leading-tight tracking-wide" style={{ fontSize: '13px' }}>
                {is1to5Class(selectedClass)
                  ? '1 ರಿಂದ 5 ನೇ ತರಗತಿಯ ನಿರಂತರ ಮತ್ತು ವ್ಯಾಪಕ ಮೌಲ್ಯಮಾಪನ ಕ್ರೋಡೀಕೃತ ವಹಿ'
                  : '6 ರಿಂದ 9 ನೇ ತರಗತಿಯ ನಿರಂತರ ಮತ್ತು ವ್ಯಾಪಕ ಮೌಲ್ಯಮಾಪನ ಕ್ರೋಡೀಕೃತ ವಹಿ'}
              </div>
              <div className="flex items-center justify-between text-[10.5px] font-bold mt-1 px-1 pb-1">
                <div style={{ flex: 1, textAlign: 'left' }}>
                  <span className="font-semibold text-neutral-800">ತರಗತಿ: </span>
                  <span className="font-bold">{selectedClass ? `${selectedClass} ನೇ` : '______'}</span>
                </div>
                <div style={{ flex: 1.2, textAlign: 'center' }}>
                  <span className="font-semibold text-neutral-800">ಶಾಲೆಯ ಹೆಸರು: </span>
                  <span className="font-bold">{school.trim() || '____________________'}</span>
                </div>
                <div style={{ flex: 1.1, textAlign: 'center' }}>
                  <span className="font-semibold text-neutral-800">DISE Code: </span>
                  <span className="font-bold">{diseCode.trim() || '___________'}</span>
                </div>
                <div style={{ flex: 1, textAlign: 'right' }}>
                  <span className="font-semibold text-neutral-800">ಶೈಕ್ಷಣಿಕ ವರ್ಷ: </span>
                  <span className="font-bold">{year || '2026-27'}</span>
                </div>
              </div>
            </div>

            {/* Printable Register Table */}
            <div className="w-full overflow-x-hidden print:overflow-visible">
              <table
                className="text-[10px] leading-[1.05] w-full"
                style={{
                  tableLayout: 'fixed',
                  width: '100%',
                  borderCollapse: 'collapse',
                  border: '0.6px solid #000',
                }}
              >
                <colgroup>
                  <col style={{ width: '1.8%' }} />
                  <col style={{ width: is1to5Class(selectedClass) ? '14.44%' : '11.65%' }} />
                  <col style={{ width: is1to5Class(selectedClass) ? '2.8%' : '2.05%' }} />
                  <col style={{ width: is1to5Class(selectedClass) ? '3.0%' : '2.05%' }} />
                  <col style={{ width: is1to5Class(selectedClass) ? '3.0%' : '2.05%' }} />
                  <col style={{ width: is1to5Class(selectedClass) ? '3.2%' : '2.7%' }} />
                  <col style={{ width: is1to5Class(selectedClass) ? '3.04%' : '2.4%' }} />
                  {Array.from({ length: subjects.length * 2 }).map((_, i) => (
                    <col
                      key={`prev-sub-${pageIdx}-${i}`}
                      style={{
                        width: is1to5Class(selectedClass)
                          ? '3.7%'
                          : (i < 6 || i === 12 || i === 13 ? '3.2%' : '2.55%'),
                      }}
                    />
                  ))}
                  <col style={{ width: is1to5Class(selectedClass) ? '3.0%' : '2.5%' }} />
                  <col style={{ width: is1to5Class(selectedClass) ? '3.0%' : '2.5%' }} />
                  <col style={{ width: is1to5Class(selectedClass) ? '3.0%' : '2.5%' }} />
                  {Array.from({ length: coSkills.length }).map((_, i) => (
                    <col key={`prev-co-${pageIdx}-${i}`} style={{ width: is1to5Class(selectedClass) ? '3.7%' : '2.05%' }} />
                  ))}
                  <col style={{ width: is1to5Class(selectedClass) ? '4.22%' : '3.0%' }} />
                </colgroup>

                <thead>
                  <tr style={{ height: '22px' }}>
                    <th rowSpan={3} style={{ border: '0.6px solid #000', background: '#fff', verticalAlign: 'middle', padding: '1px', textAlign: 'center' }}>
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '10px', textAlign: 'center' }}>ಕ್ರ.ಸಂ.</div>
                    </th>
                    <th rowSpan={3} className="stu-print-hdr" style={{ border: '0.6px solid #000', background: '#fff', verticalAlign: 'middle', textAlign: 'center', padding: '1px 2px', boxSizing: 'border-box' }}>
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%', height: '100%' }}>
                        <div style={{ fontSize: '11px', fontWeight: 700, textAlign: 'center', lineHeight: '1.2' }}>ವಿದ್ಯಾರ್ಥಿ ವಿವರ</div>
                        <div style={{ fontSize: '8.5px', fontWeight: 600, textAlign: 'center', lineHeight: '1.2', marginTop: '2px' }}>ಹೆಸರು / ದಾಖಲಾತಿ / ಲಿಂಗ</div>
                        <div style={{ fontSize: '8.5px', fontWeight: 600, textAlign: 'center', lineHeight: '1.2', marginTop: '1px' }}>/ DOB / SATS / PEN</div>
                      </div>
                    </th>
                    <th rowSpan={3} style={{ border: '0.6px solid #000', background: '#fff', verticalAlign: 'middle', padding: '1px', textAlign: 'center', boxSizing: 'border-box' }}>
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '10px', textAlign: 'center' }}>ಸೆಮಿಸ್ಟರ್</div>
                    </th>
                    <th rowSpan={3} style={{ border: '0.6px solid #000', background: '#fff', verticalAlign: 'middle', padding: '1px', textAlign: 'center' }}>
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '10px', textAlign: 'center' }}>ಶಾಲೆ ನಡೆದ ದಿನ</div>
                    </th>
                    <th rowSpan={3} style={{ border: '0.6px solid #000', background: '#fff', verticalAlign: 'middle', padding: '1px', textAlign: 'center' }}>
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '10px', textAlign: 'center' }}>ಹಾಜರಾದ ದಿನ</div>
                    </th>
                    <th rowSpan={3} style={{ border: '0.6px solid #000', background: '#fff', verticalAlign: 'middle', padding: '1px', textAlign: 'center' }}>
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '9.5px', textAlign: 'center' }}>ಮೌಲ್ಯಾಂಕನ</div>
                    </th>
                    <th rowSpan={3} style={{ border: '0.6px solid #000', background: '#fff', verticalAlign: 'middle', padding: '1px', textAlign: 'center' }}>
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '9px', textAlign: 'center' }}>ಶೇ. ಪ್ರಮಾಣ</div>
                    </th>
                    <th colSpan={subjects.length * 2 + 3} className="print-sec-hdr" style={{ border: '0.6px solid #000', background: '#fff', fontSize: '14.5px', fontWeight: 700, textAlign: 'center', verticalAlign: 'middle', height: '22px', padding: '1px' }} title={gradeTooltipText}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%', height: '100%', fontSize: '14.5px' }}>
                        ವಿದ್ಯಾರ್ಥಿಯ ಶೈಕ್ಷಣಿಕ ಸಾಧನೆ - ಪಠ್ಯ ಭಾಗಕ್ಕೆ ಸಂಬಂಧಿಸಿದಂತೆ
                      </div>
                    </th>
                    <th colSpan={coSkills.length} className="print-co-hdr" style={{ border: '0.6px solid #000', background: '#fff', fontSize: '13px', fontWeight: 700, textAlign: 'center', verticalAlign: 'middle', height: '22px', padding: '1px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%', height: '100%', fontSize: '13px', whiteSpace: 'nowrap' }}>
                        ಸಹಪಠ್ಯ ವಿಷಯದ ಶ್ರೇಣಿಗಳು
                      </div>
                    </th>
                    <th rowSpan={3} style={{ border: '0.6px solid #000', background: '#fff', verticalAlign: 'middle', padding: '1px', textAlign: 'center' }}>
                      <div className="vtxt mx-auto font-bold" style={{ fontSize: '10px', textAlign: 'center' }}>ಫಲಿತಾಂಶ</div>
                    </th>
                  </tr>

                  <tr style={{ height: is1to5Class(selectedClass) ? '54px' : '46px' }}>
                    {subjects.map((sub) => (
                      <th
                        key={`prev-hdr-${pageIdx}-${sub}`}
                        colSpan={2}
                        style={{
                          border: '0.6px solid #000',
                          background: '#fff',
                          fontSize: '9.5px',
                          fontWeight: 700,
                          textAlign: 'center',
                          verticalAlign: 'middle',
                          padding: '1px 0',
                        }}
                      >
                        {renderSubjectHeader(sub, true, languageNames)}
                      </th>
                    ))}
                    <th colSpan={3} style={{ border: '0.6px solid #000', background: '#fff', fontSize: '10.5px', fontWeight: 700, textAlign: 'center', verticalAlign: 'middle', padding: '1px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%' }}>
                        ಒಟ್ಟು
                      </div>
                    </th>
                    {coSkills.map((co) => (
                      <th
                        key={`prev-co-hdr-${pageIdx}-${co}`}
                        rowSpan={2}
                        className="print-co-skill-th"
                        style={{
                          border: '0.6px solid #000',
                          background: '#fff',
                          verticalAlign: 'middle',
                          textAlign: 'center',
                          padding: '1px',
                          height: is1to5Class(selectedClass) ? '78px' : '66px',
                          boxSizing: 'border-box',
                        }}
                      >
                        {renderCoSkillHeader(co, true)}
                      </th>
                    ))}
                  </tr>

                  <tr style={{ height: is1to5Class(selectedClass) ? '24px' : '20px' }}>
                    {subjects.map((sub, i) => (
                      <React.Fragment key={`prev-sub-sub-${pageIdx}-${sub}-${i}`}>
                        <th style={{ border: '0.6px solid #000', background: '#fff', fontSize: '11.5px', fontWeight: 700, textAlign: 'center', verticalAlign: 'middle', padding: '1px 0' }}>
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%' }}>ಅಂಕ</div>
                        </th>
                        <th style={{ border: '0.6px solid #000', background: '#fff', fontSize: '11.5px', fontWeight: 700, textAlign: 'center', verticalAlign: 'middle', padding: '1px 0' }} title={gradeTooltipText}>
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%' }}>ಶ್ರೇಣಿ</div>
                        </th>
                      </React.Fragment>
                    ))}
                    <th style={{ border: '0.6px solid #000', background: '#fff', fontSize: '11.5px', fontWeight: 700, textAlign: 'center', verticalAlign: 'middle', padding: '1px 0' }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%' }}>ಅಂಕ</div>
                    </th>
                    <th style={{ border: '0.6px solid #000', background: '#fff', fontSize: '11.5px', fontWeight: 700, textAlign: 'center', verticalAlign: 'middle', padding: '1px 0' }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%' }}>ಶೇ.</div>
                    </th>
                    <th style={{ border: '0.6px solid #000', background: '#fff', fontSize: '11.5px', fontWeight: 700, textAlign: 'center', verticalAlign: 'middle', padding: '1px 0' }} title={gradeTooltipText}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', textAlign: 'center', width: '100%' }}>ಶ್ರೇಣಿ</div>
                    </th>
                  </tr>
                </thead>

                <tbody>
                  {pageStudents.map((stu, sIdx) =>
                    stu.assessments.map((ass, aIdx) => {
                      const totDays = (stu.attendance.s1Total || 0) + (stu.attendance.s2Total || 0);
                      const totPres = (stu.attendance.s1Present || 0) + (stu.attendance.s2Present || 0);
                      const is50 = aIdx === 3 || aIdx === 7 || !!ass.is50;
                      const is100 = aIdx === 8 || !!ass.is100;
                      const is50Or100 = is50 || is100;
                      const rowBg = is50Or100 ? '#e5e7eb' : '#ffffff';
                      return (
                        <tr
                          key={`prev-row-${pageIdx}-${sIdx}-${aIdx}`}
                          className={is50Or100 ? 'row-highlight-print' : ''}
                          style={{ height: is1to5Class(selectedClass) ? '21.2px' : '20.8px', background: rowBg }}
                        >
                          {aIdx === 0 && (
                            <td rowSpan={9} className="print-stu-thick-bottom" style={{ border: '0.6px solid #000', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000', borderLeft: '1.5px solid #000', background: '#fff', textAlign: 'center', fontWeight: 700, verticalAlign: 'middle', fontSize: '11px', padding: '1px 1px' }}>
                              {stu.id}
                            </td>
                          )}
                          {aIdx === 0 && (
                            <td
                              rowSpan={9}
                              className="stu-print-td print-stu-thick-bottom"
                              style={{
                                border: '0.6px solid #000',
                                borderTop: '1.5px solid #000',
                                borderRight: '0.6px solid #000',
                                borderBottom: '1.5px solid #000',
                                background: '#fff',
                                verticalAlign: 'middle',
                                textAlign: 'left',
                                padding: '6px 8px',
                                boxSizing: 'border-box',
                              }}
                            >
                              <div className="stu-print-wrap flex flex-col items-start justify-center text-left w-full mx-auto" style={{ textAlign: 'left' }}>
                                <div
                                  className="stu-print-name font-bold text-[12.5px] leading-tight text-left w-full"
                                  style={{
                                    fontSize: '12.5px',
                                    marginBottom: '8px',
                                    textAlign: 'left',
                                    whiteSpace: 'normal',
                                    wordBreak: 'break-word',
                                    overflowWrap: 'break-word',
                                    display: '-webkit-box',
                                    WebkitLineClamp: 2,
                                    WebkitBoxOrient: 'vertical',
                                    overflow: 'hidden',
                                    lineHeight: 1.3,
                                    maxHeight: '34px',
                                  }}
                                  title={stu.name}
                                >
                                  {stu.name || '-'}
                                </div>
                                <div className="stu-print-detail text-[10px] text-left w-full" style={{ fontSize: '10px', marginBottom: '4px', textAlign: 'left', lineHeight: 1.3 }}>
                                  ದಾ . ಸಂ: {stu.admission || '-'}
                                </div>
                                <div className="stu-print-detail text-[10px] text-left w-full" style={{ fontSize: '10px', marginBottom: '4px', textAlign: 'left', lineHeight: 1.3 }}>
                                  ಲಿಂಗ: {stu.gender || '-'}
                                </div>
                                <div className="stu-print-detail text-[10px] text-left w-full" style={{ fontSize: '10px', marginBottom: '4px', textAlign: 'left', lineHeight: 1.3 }}>
                                  DOB: {formatDateDisplay(stu.dob)}
                                </div>
                                <div className="stu-print-detail text-[10px] text-left w-full" style={{ fontSize: '10px', marginBottom: '4px', textAlign: 'left', lineHeight: 1.3 }}>
                                  SATS: {stu.sats || '-'}
                                </div>
                                <div className="stu-print-detail text-[10px] text-left w-full" style={{ fontSize: '10px', textAlign: 'left', lineHeight: 1.3 }}>
                                  PEN: {stu.pen || '-'}
                                </div>
                              </div>
                            </td>
                          )}
                          {aIdx === 0 && (
                            <td rowSpan={4} style={{ border: '0.6px solid #000', borderTop: '1.5px solid #000', borderLeft: '0.6px solid #000', background: '#fff', textAlign: 'center', verticalAlign: 'middle', padding: '1px 1px', boxSizing: 'border-box' }}>
                              <div className="vtxt sem-vtxt mx-auto font-bold text-[10px]">1ನೇ</div>
                            </td>
                          )}
                          {aIdx === 4 && (
                            <td rowSpan={5} className="print-stu-thick-bottom" style={{ border: '0.6px solid #000', borderLeft: '0.6px solid #000', borderBottom: '1.5px solid #000', background: '#fff', textAlign: 'center', verticalAlign: 'middle', padding: '1px 1px', boxSizing: 'border-box' }}>
                              <div className="vtxt sem-vtxt mx-auto font-bold text-[10px]">2ನೇ</div>
                            </td>
                          )}
                          {aIdx === 0 && (
                            <td rowSpan={4} className="attendance-cell" style={{ border: '0.6px solid #000', borderTop: '1.5px solid #000', background: '#fff', textAlign: 'center', verticalAlign: 'middle', fontWeight: 700, fontSize: '10px', padding: '1px 1px' }}>
                              {stu.attendance.s1Total || '-'}
                            </td>
                          )}
                          {aIdx === 4 && (
                            <td rowSpan={4} className="attendance-cell" style={{ border: '0.6px solid #000', background: '#fff', textAlign: 'center', verticalAlign: 'middle', fontWeight: 700, fontSize: '10px', padding: '1px 1px' }}>
                              {stu.attendance.s2Total || '-'}
                            </td>
                          )}
                          {aIdx === 0 && (
                            <td rowSpan={4} className="attendance-cell" style={{ border: '0.6px solid #000', borderTop: '1.5px solid #000', background: '#fff', textAlign: 'center', verticalAlign: 'middle', fontWeight: 700, fontSize: '10px', padding: '1px 1px' }}>
                              {stu.attendance.s1Present || '-'}
                            </td>
                          )}
                          {aIdx === 4 && (
                            <td rowSpan={4} className="attendance-cell" style={{ border: '0.6px solid #000', background: '#fff', textAlign: 'center', verticalAlign: 'middle', fontWeight: 700, fontSize: '10px', padding: '1px 1px' }}>
                              {stu.attendance.s2Present || '-'}
                            </td>
                          )}
                          {aIdx === 8 && (
                            <>
                              <td className="attendance-cell highlight-print-cell print-stu-thick-bottom" style={{ border: '0.6px solid #000', borderBottom: '1.5px solid #000', background: '#e5e7eb', textAlign: 'center', verticalAlign: 'middle', fontWeight: 700, fontSize: '10px', padding: '1px 1px' }}>
                                {totDays || '-'}
                              </td>
                              <td className="attendance-cell highlight-print-cell print-stu-thick-bottom" style={{ border: '0.6px solid #000', borderBottom: '1.5px solid #000', background: '#e5e7eb', textAlign: 'center', verticalAlign: 'middle', fontWeight: 700, fontSize: '10px', padding: '1px 1px' }}>
                                {totPres || '-'}
                              </td>
                            </>
                          )}
                          {/* Assessment Name Cell (ಮೌಲ್ಯಾಂಕನ) */}
                          <td
                            className={`moulankan-print-cell ${is50Or100 ? 'highlight-print-cell font-bold' : ''} ${aIdx === 8 ? 'print-stu-thick-bottom' : ''}`}
                            style={{
                              border: '0.6px solid #000',
                              borderTop: aIdx === 0 ? '1.5px solid #000' : '0.6px solid #000',
                              borderBottom: aIdx === 8 ? '1.5px solid #000' : '0.6px solid #000',
                              background: rowBg,
                              textAlign: 'center',
                              verticalAlign: 'middle',
                              fontWeight: 700,
                              fontSize: '9.5px',
                              padding: '0 1px',
                              whiteSpace: 'nowrap',
                            }}
                          >
                            <div style={{ whiteSpace: 'nowrap', textAlign: 'center', width: '100%', fontSize: '9.5px', fontWeight: 700, lineHeight: 1.15 }}>
                              {parts[aIdx]?.name || ass.label}
                            </div>
                          </td>
                          {/* Assessment Pct Cell (ಶೇ. ಪ್ರಮಾಣ) */}
                          <td
                            className={`moulankan-pct-print-cell ${is50Or100 ? 'highlight-print-cell font-bold' : ''} ${aIdx === 8 ? 'print-stu-thick-bottom' : ''}`}
                            style={{
                              border: '0.6px solid #000',
                              borderTop: aIdx === 0 ? '1.5px solid #000' : '0.6px solid #000',
                              borderBottom: aIdx === 8 ? '1.5px solid #000' : '0.6px solid #000',
                              background: rowBg,
                              textAlign: 'center',
                              verticalAlign: 'middle',
                              fontWeight: 700,
                              fontSize: '9px',
                              padding: '0 1px',
                              whiteSpace: 'nowrap',
                            }}
                          >
                            <div style={{ whiteSpace: 'nowrap', textAlign: 'center', width: '100%', fontSize: '9px', fontWeight: 700, lineHeight: 1.15 }}>
                              {parts[aIdx]?.pct || ''}
                            </div>
                          </td>
                          {ass.subs.map((sub, subIdx) => (
                            <React.Fragment key={`prev-score-${pageIdx}-${sIdx}-${aIdx}-${subIdx}`}>
                              <td
                                className={`sub-score-td ${is50Or100 ? 'sub-mark-bold highlight-print-cell' : 'sub-mark-normal'} ${aIdx === 8 ? 'print-stu-thick-bottom' : ''}`}
                                style={{
                                  border: '0.6px solid #000',
                                  borderTop: aIdx === 0 ? '1.5px solid #000' : '0.6px solid #000',
                                  borderBottom: aIdx === 8 ? '1.5px solid #000' : '0.6px solid #000',
                                  background: rowBg,
                                  textAlign: 'center',
                                  verticalAlign: 'middle',
                                  fontWeight: is50Or100 ? 800 : 400,
                                  fontSize: '12.5px',
                                  padding: '1px 1px',
                                }}
                              >
                                {sub.m !== '' ? sub.m : '-'}
                              </td>
                              <td
                                className={`sub-score-td ${is50Or100 ? 'sub-mark-bold highlight-print-cell' : 'sub-grade-normal'} ${aIdx === 8 ? 'print-stu-thick-bottom' : ''}`}
                                style={{
                                  border: '0.6px solid #000',
                                  borderTop: aIdx === 0 ? '1.5px solid #000' : '0.6px solid #000',
                                  borderBottom: aIdx === 8 ? '1.5px solid #000' : '0.6px solid #000',
                                  background: rowBg,
                                  textAlign: 'center',
                                  verticalAlign: 'middle',
                                  fontWeight: is50Or100 ? 800 : 400,
                                  fontSize: '12.5px',
                                  padding: '1px 1px',
                                }}
                              >
                                {sub.g || '-'}
                              </td>
                            </React.Fragment>
                          ))}
                          <td
                            className={`${is50Or100 ? 'highlight-print-cell sub-mark-bold' : ''} ${aIdx === 8 ? 'print-stu-thick-bottom' : ''}`}
                            style={{ border: '0.6px solid #000', borderTop: aIdx === 0 ? '1.5px solid #000' : '0.6px solid #000', borderBottom: aIdx === 8 ? '1.5px solid #000' : '0.6px solid #000', background: rowBg, textAlign: 'center', verticalAlign: 'middle', fontWeight: is50Or100 ? 800 : 700, fontSize: '12.5px', padding: '1px 1px' }}
                          >
                            {ass.total.m !== '' ? ass.total.m : '-'}
                          </td>
                          <td
                            className={`${is50Or100 ? 'highlight-print-cell sub-mark-bold' : ''} ${aIdx === 8 ? 'print-stu-thick-bottom' : ''}`}
                            style={{ border: '0.6px solid #000', borderTop: aIdx === 0 ? '1.5px solid #000' : '0.6px solid #000', borderBottom: aIdx === 8 ? '1.5px solid #000' : '0.6px solid #000', background: rowBg, textAlign: 'center', verticalAlign: 'middle', fontWeight: is50Or100 ? 800 : 700, fontSize: '12.5px', padding: '1px 1px' }}
                          >
                            {ass.total.p !== '' ? ass.total.p : '-'}
                          </td>
                          <td
                            className={`${is50Or100 ? 'highlight-print-cell sub-mark-bold' : ''} ${aIdx === 8 ? 'print-stu-thick-bottom' : ''}`}
                            style={{ border: '0.6px solid #000', borderTop: aIdx === 0 ? '1.5px solid #000' : '0.6px solid #000', borderBottom: aIdx === 8 ? '1.5px solid #000' : '0.6px solid #000', background: rowBg, textAlign: 'center', verticalAlign: 'middle', fontWeight: is50Or100 ? 800 : 700, fontSize: '12.5px', padding: '1px 1px' }}
                          >
                            {ass.total.g || '-'}
                          </td>
                          {aIdx === 0 &&
                            stu.co1.slice(0, coSkills.length).map((coVal, coIdx) => (
                              <td
                                key={`prev-co1-${pageIdx}-${sIdx}-${coIdx}`}
                                rowSpan={4}
                                className="co-grade-td"
                                style={{
                                  border: '0.6px solid #000',
                                  borderTop: '1.5px solid #000',
                                  background: '#fff',
                                  textAlign: 'center',
                                  verticalAlign: 'middle',
                                  fontWeight: 700,
                                  fontSize: '11px',
                                  padding: '1px 1px',
                                }}
                              >
                                {coVal || '-'}
                              </td>
                            ))}
                          {aIdx === 4 &&
                            stu.co2.slice(0, coSkills.length).map((coVal, coIdx) => (
                              <td
                                key={`prev-co2-${pageIdx}-${sIdx}-${coIdx}`}
                                rowSpan={4}
                                className="co-grade-td"
                                style={{
                                  border: '0.6px solid #000',
                                  background: '#fff',
                                  textAlign: 'center',
                                  verticalAlign: 'middle',
                                  fontWeight: 700,
                                  fontSize: '11px',
                                  padding: '1px 1px',
                                }}
                              >
                                {coVal || '-'}
                              </td>
                            ))}
                          {aIdx === 8 &&
                            coSkills.map((_, coIdx) => {
                              const c1 = stu.co1[coIdx] || '';
                              const c2 = stu.co2[coIdx] || '';
                              const finalCo = !c1 && !c2 ? '-' : c1 === 'A' && c2 === 'A' ? 'A' : 'B';
                              return (
                                <td
                                  key={`prev-co-fin-${pageIdx}-${sIdx}-${coIdx}`}
                                  className="co-grade-td highlight-print-cell print-stu-thick-bottom"
                                  style={{
                                    border: '0.6px solid #000',
                                    borderBottom: '1.5px solid #000',
                                    background: '#f0f0f0',
                                    textAlign: 'center',
                                    verticalAlign: 'middle',
                                    fontWeight: 700,
                                    fontSize: '11px',
                                    padding: '1px 1px',
                                  }}
                                >
                                  {finalCo}
                                </td>
                              );
                            })}
                          {aIdx === 0 &&
                            (() => {
                              const grandMarks = toInt(stu.assessments[8]?.total?.m);
                              const pct = grandMarks ? calcPercentExact(grandMarks, maxGrand) : 0;
                              const pctRound = grandMarks ? Math.round(pct) : 0;
                              const gr = grandMarks ? getGrade(pct, gradeRules) : '';
                              const resLabel = stu.result || (grandMarks ? 'ಉತ್ತೀರ್ಣ' : '');
                              const resText = resLabel
                                ? grandMarks
                                  ? `${resLabel} = ( ${pctRound}% - ${gr || 'C'} )`
                                  : resLabel
                                : '-';
                              return (
                                <td rowSpan={9} className="print-stu-thick-bottom" style={{ border: '0.6px solid #000', borderTop: '1.5px solid #000', borderBottom: '1.5px solid #000', borderRight: '1.5px solid #000', background: '#fff', textAlign: 'center', verticalAlign: 'middle', padding: '1px 0' }}>
                                  <div className="vtxt result-vtxt mx-auto font-bold" style={{ fontSize: '12.5px', lineHeight: 1, maxHeight: '170px', overflow: 'hidden' }} title={resText}>
                                    {resText}
                                  </div>
                                </td>
                              );
                            })()}
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Class Summary Table & Signatures Row - shown only on the last page for each class */}
            {pageIdx === studentPages.length - 1 && (
              <div className="mt-3 pt-1 text-black" style={{ breakInside: 'avoid', pageBreakInside: 'avoid' }}>
                {/* Summary Table aligned under Pratham Bhashe to Ottu Shrene columns */}
                <div className="flex flex-col items-center mx-auto" style={{ marginLeft: '22%', width: '50%', minWidth: '450px' }}>
                  <div className="text-[12px] font-bold mb-1.5 text-black text-center w-full">
                    ತರಗತಿಯ ಫಲಿತಾಂಶ ಘೋಷವಾರ
                  </div>
                  <table
                    className="text-[11px] w-full text-black"
                    style={{
                      borderCollapse: 'collapse',
                      border: '1px solid #000',
                      textAlign: 'center',
                    }}
                  >
                    <thead>
                      <tr style={{ background: '#e5e7eb', height: '25px' }}>
                        <th style={{ border: '0.6px solid #000', padding: '3px 8px', textAlign: 'left', fontWeight: 700 }}>ವಿವರ</th>
                        <th style={{ border: '0.6px solid #000', padding: '3px 8px', width: '85px', fontWeight: 700 }}>ಬಾಲಕರು</th>
                        <th style={{ border: '0.6px solid #000', padding: '3px 8px', width: '85px', fontWeight: 700 }}>ಬಾಲಕಿಯರು</th>
                        <th style={{ border: '0.6px solid #000', padding: '3px 8px', width: '85px', fontWeight: 700 }}>ಒಟ್ಟು</th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr style={{ height: '22px' }}>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', textAlign: 'left', fontWeight: 600 }}>ದಾಖಲಾದ ಒಟ್ಟು ವಿದ್ಯಾರ್ಥಿಗಳು</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.totalBoys}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.totalGirls}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 800 }}>{summaryStats.totalStudents}</td>
                      </tr>
                      <tr style={{ height: '22px' }}>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', textAlign: 'left', fontWeight: 600 }}>ಪರೀಕ್ಷೆಗೆ ಹಾಜರಾದವರು</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.appearedBoys}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.appearedGirls}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 800 }}>{summaryStats.appearedTotal}</td>
                      </tr>
                      <tr style={{ height: '22px' }}>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', textAlign: 'left', fontWeight: 600 }}>ಉತ್ತೀರ್ಣರಾದವರು</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.passedBoys}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.passedGirls}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 800 }}>{summaryStats.passedTotal}</td>
                      </tr>
                      <tr style={{ height: '22px' }}>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', textAlign: 'left', fontWeight: 600 }}>ಅನುತ್ತೀರ್ಣರಾದವರು</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.failedBoys}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.failedGirls}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 800 }}>{summaryStats.failedTotal}</td>
                      </tr>
                      <tr style={{ background: '#f9fafb', height: '22px' }}>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', textAlign: 'left', fontWeight: 700 }}>ಉತ್ತೀರ್ಣತೆಯ ಶೇಕಡಾವಾರು (%)</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.passPctBoys}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 700 }}>{summaryStats.passPctGirls}</td>
                        <td style={{ border: '0.6px solid #000', padding: '3px 8px', fontWeight: 800 }}>{summaryStats.passPctTotal}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* Date & Signatures row below analysis table in one line with proper space */}
                <div className="w-full flex items-center justify-between text-[11px] font-bold mt-5 pt-1 px-3 text-black">
                  <span>ದಿನಾಂಕ: _________________</span>
                  <span>ತರಗತಿ ಶಿಕ್ಷಕರ ಸಹಿ: _________________</span>
                  <span>ಮುಖ್ಯೋಪಾಧ್ಯಾಯರ ಸಹಿ ಹಾಗೂ ಶಾಲಾ ಮೊಹರು: _________________</span>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
