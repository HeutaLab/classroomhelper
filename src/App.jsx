import React, { useState, useEffect, useRef, useMemo, useCallback, createContext, useContext } from 'react';
import { Download, Plus, Trash2, Users, ClipboardList, Settings2, ChevronRight, Loader2, Upload, X, Star, Calendar as CalendarIcon, Camera, MessageSquare, Sun, Moon, Contrast, Smartphone, BarChart3, TrendingUp } from 'lucide-react';
import * as XLSX from 'xlsx';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell, LineChart, Line } from 'recharts';

const LEVELS = [
  { name: 'Not started', fill: 'transparent', tint: 'transparent', text: 'var(--ink-soft)' },
  { name: 'Developing', fill: 'var(--coral)', tint: 'var(--coral-bg)', text: 'var(--coral-text)' },
  { name: 'Secure', fill: 'var(--green)', tint: 'var(--green-bg)', text: 'var(--green-text)' },
  { name: 'Mastered', fill: 'var(--indigo)', tint: 'var(--indigo-bg)', text: 'var(--indigo-text)' },
];
const LEVEL_LETTERS = ['', 'D', 'S', 'M'];
const HIGH_ACHIEVER_PCT = 85;
const NEEDS_PUSH_PCT = 40;

function computeStudentFlag(studentId, criteriaCols, records, termId, recKey) {
  if (!criteriaCols.length) return null;
  let total = 0;
  let hasAny = false;
  criteriaCols.forEach((c) => {
    const lvl = records[recKey(studentId, termId, c.id)] ?? 0;
    total += lvl;
    if (lvl > 0) hasAny = true;
  });
  const pct = (total / (criteriaCols.length * 3)) * 100;
  if (pct >= HIGH_ACHIEVER_PCT) return 'high';
  if (hasAny && pct < NEEDS_PUSH_PCT) return 'push';
  return null;
}

function StudentFlagBadge({ flag, size = 13 }) {
  if (!flag) return null;
  if (flag === 'high') {
    return <span title="High achiever — mostly Mastered here" style={{ display: 'inline-flex', flexShrink: 0 }}><Star size={size} color="var(--indigo)" fill="var(--indigo)" /></span>;
  }
  return <span title="May need a push — stuck below Secure across most of this view" style={{ display: 'inline-flex', flexShrink: 0 }}><TrendingUp size={size} color="var(--coral)" /></span>;
}


const INK = 'var(--ink)';
const PAPER = 'var(--cream)';
const CARD = 'var(--card)';
const LINE = 'var(--line)';
const MUTED = 'var(--ink-soft)';
const ACCENT = 'var(--accent)';
const AMBER = 'var(--amber)';

const ThemeContext = createContext('pale');

const THEME_CSS = `
:root, [data-theme="pale"] {
  --cream: #F7F5EF; --card: #FFFFFF; --ink: #20241F; --ink-soft: #8B8778; --ink-muted: #B7B2A2; --line: #E1DDD0;
  --coral: #C98A2C; --coral-bg: #FBEEDA; --coral-text: #7A4C14;
  --green: #2F7D6B; --green-bg: #E1F0EC; --green-text: #1B4B40;
  --indigo: #4B3F72; --indigo-bg: #EAE7F2; --indigo-text: #2E2648;
  --amber: #C98A2C; --accent: #4B3F72; --accent-hover: #3B3159; --border-width: 1px;
  --row-stripe: rgba(0,0,0,0.035); --hover-bg: rgba(0,0,0,0.05);
}
[data-theme="dark"] {
  --cream: #17181C; --card: #1F2024; --ink: #F2F0EA; --ink-soft: #B8B6AE; --ink-muted: #75746D; --line: #34353A;
  --coral: #FF8A6B; --coral-bg: #4A2A1F; --coral-text: #FFCBB8;
  --green: #4FC585; --green-bg: #1D3A2A; --green-text: #A8E8C1;
  --indigo: #8A7BF0; --indigo-bg: #2C2650; --indigo-text: #C9C0FA;
  --amber: #F4B85C; --accent: #8A7BF0; --accent-hover: #A192F5; --border-width: 1px;
  --row-stripe: rgba(255,255,255,0.045); --hover-bg: rgba(255,255,255,0.07);
}
[data-theme="pastel"] {
  --cream: #FDF9F5; --card: #FFFFFF; --ink: #3A3733; --ink-soft: #8A867D; --ink-muted: #C4C0B5; --line: #F0E9DC;
  --coral: #FFA98F; --coral-bg: #FFEDE5; --coral-text: #B5502E;
  --green: #7FCBA0; --green-bg: #E4F5EA; --green-text: #2E7A50;
  --indigo: #A79CF0; --indigo-bg: #EFECFC; --indigo-text: #5B4FB8;
  --amber: #F5C97E; --accent: #9B8FDE; --accent-hover: #8577CC; --border-width: 1px;
  --row-stripe: rgba(0,0,0,0.025); --hover-bg: rgba(0,0,0,0.04);
}
[data-theme="hc"] {
  --cream: #FFFFFF; --card: #FFFFFF; --ink: #000000; --ink-soft: #1A1A1A; --ink-muted: #4D4D4D; --line: #000000;
  --coral: #C43E00; --coral-bg: #FFFFFF; --coral-text: #C43E00;
  --green: #146C34; --green-bg: #FFFFFF; --green-text: #146C34;
  --indigo: #3423B0; --indigo-bg: #FFFFFF; --indigo-text: #3423B0;
  --amber: #8A5A00; --accent: #000000; --accent-hover: #262626; --border-width: 2px;
  --row-stripe: rgba(0,0,0,0.06); --hover-bg: rgba(0,0,0,0.09);
}
`;

const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-4);

const emptyStructure = { classes: [], subjects: [], terms: [], topics: [], criteria: [], fields: [], yearGroups: [], departments: [] };
const FIELD_TYPES = {
  note: { label: 'Note', icon: 'note' },
  rating: { label: 'Rating', icon: 'rating' },
  date: { label: 'Date', icon: 'date' },
  photo: { label: 'Photo', icon: 'photo' },
};

function compressImage(file, maxDim = 900, quality = 0.6) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new window.Image();
      img.onload = () => {
        let { width, height } = img;
        if (width > maxDim || height > maxDim) {
          if (width > height) { height = Math.round(height * (maxDim / width)); width = maxDim; }
          else { width = Math.round(width * (maxDim / height)); height = maxDim; }
        }
        const canvas = document.createElement('canvas');
        canvas.width = width; canvas.height = height;
        canvas.getContext('2d').drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL('image/jpeg', quality));
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function loadKey(key, fallback, shared = true) {
  try {
    const res = await window.storage.get(key, shared);
    if (res && res.value) return JSON.parse(res.value);
    return fallback;
  } catch (e) {
    return fallback;
  }
}

async function saveKey(key, value, shared = true) {
  try {
    await window.storage.set(key, JSON.stringify(value), shared);
  } catch (e) {
    console.error('storage save failed', key, e);
  }
}

function useDebouncedSave(key, value, ready, delay = 500, shared = true) {
  const timer = useRef(null);
  const first = useRef(true);
  useEffect(() => {
    if (!ready) return;
    if (first.current) { first.current = false; return; }
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => saveKey(key, value, shared), delay);
    return () => timer.current && clearTimeout(timer.current);
  }, [key, value, ready, delay, shared]);
}

function Pill({ active, onClick, children, tone }) {
  return (
    <button
      onClick={onClick}
      style={{
        padding: '6px 14px',
        borderRadius: 999,
        border: `var(--border-width) solid ${active ? ACCENT : LINE}`,
        background: active ? ACCENT : CARD,
        color: active ? '#fff' : INK,
        fontSize: 13,
        fontWeight: 500,
        cursor: 'pointer',
        whiteSpace: 'nowrap',
        flexShrink: 0,
      }}
    >
      {children}
    </button>
  );
}

const THEME_OPTIONS = [
  { id: 'pale', label: 'Pale', Icon: Sun },
  { id: 'dark', label: 'Dark', Icon: Moon },
  { id: 'pastel', label: 'Pastel', Icon: Sun },
  { id: 'hc', label: 'Hi-contrast', Icon: Contrast },
];

function ThemeToggle({ theme, setTheme }) {
  return (
    <div style={{ display: 'flex', background: PAPER, borderRadius: 999, padding: 3, border: `var(--border-width) solid ${LINE}`, gap: 2 }}>
      {THEME_OPTIONS.map(({ id, label, Icon }) => (
        <button
          key={id}
          onClick={() => setTheme(id)}
          title={label}
          style={{
            display: 'flex', alignItems: 'center', gap: 4, padding: '5px 8px', borderRadius: 999,
            border: 'none', background: theme === id ? ACCENT : 'transparent', color: theme === id ? '#fff' : MUTED,
            fontSize: 11.5, fontWeight: 500, cursor: 'pointer', whiteSpace: 'nowrap',
          }}
        >
          <Icon size={12} />
        </button>
      ))}
    </div>
  );
}

function IconBtn({ onClick, title, danger, children }) {
  return (
    <button
      onClick={onClick}
      title={title}
      style={{
        border: 'none',
        background: 'transparent',
        color: danger ? '#9A4A3A' : MUTED,
        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        padding: 4,
        borderRadius: 6,
      }}
      onMouseEnter={(e) => (e.currentTarget.style.background = 'var(--hover-bg)')}
      onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
    >
      {children}
    </button>
  );
}

function TextAdd({ placeholder, onAdd }) {
  const [val, setVal] = useState('');
  const submit = () => {
    const v = val.trim();
    if (!v) return;
    onAdd(v);
    setVal('');
  };
  return (
    <div style={{ display: 'flex', gap: 6 }}>
      <input
        value={val}
        onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => e.key === 'Enter' && submit()}
        placeholder={placeholder}
        style={{
          flex: 1,
          fontSize: 13,
          padding: '7px 10px',
          borderRadius: 7,
          border: `var(--border-width) solid ${LINE}`,
          outline: 'none',
          background: CARD,
        }}
      />
      <button
        onClick={submit}
        style={{
          border: `var(--border-width) solid ${LINE}`,
          background: CARD,
          borderRadius: 7,
          padding: '0 10px',
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          color: INK,
        }}
      >
        <Plus size={15} />
      </button>
    </div>
  );
}

function FieldAdd({ onAdd }) {
  const [name, setName] = useState('');
  const [type, setType] = useState('note');
  const submit = () => {
    const v = name.trim();
    if (!v) return;
    onAdd(v, type);
    setName('');
  };
  return (
    <div style={{ display: 'flex', gap: 6 }}>
      <input
        value={name}
        onChange={(e) => setName(e.target.value)}
        onKeyDown={(e) => e.key === 'Enter' && submit()}
        placeholder="Add extra column, e.g. Model Level"
        style={{ flex: 1, fontSize: 13, padding: '7px 10px', borderRadius: 7, border: `var(--border-width) solid ${LINE}`, outline: 'none', background: CARD }}
      />
      <select
        value={type}
        onChange={(e) => setType(e.target.value)}
        style={{ fontSize: 12, padding: '0 6px', borderRadius: 7, border: `var(--border-width) solid ${LINE}`, background: CARD, color: INK }}
      >
        {Object.entries(FIELD_TYPES).map(([key, v]) => <option key={key} value={key}>{v.label}</option>)}
      </select>
      <button
        onClick={submit}
        style={{ border: `var(--border-width) solid ${LINE}`, background: CARD, borderRadius: 7, padding: '0 10px', cursor: 'pointer', display: 'flex', alignItems: 'center', color: INK }}
      >
        <Plus size={15} />
      </button>
    </div>
  );
}

function LevelPill({ level, height = 20, width }) {
  const theme = useContext(ThemeContext);
  const isEmpty = level === 0;
  const isHc = theme === 'hc';
  const colorVar = level === 1 ? 'var(--coral)' : level === 2 ? 'var(--green)' : level === 3 ? 'var(--indigo)' : 'var(--ink-muted)';

  if (isHc) {
    return (
      <div
        style={{
          height, width: width || '100%', maxWidth: width || undefined, margin: width ? undefined : '0 auto',
          borderRadius: height / 2,
          background: CARD,
          border: `2px ${isEmpty ? 'dashed' : 'solid'} ${colorVar}`,
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          fontSize: Math.max(9, Math.round(height * 0.5)), fontWeight: 700, color: colorVar, lineHeight: 1,
        }}
      >
        {LEVEL_LETTERS[level]}
      </div>
    );
  }
  return (
    <div
      style={{
        height,
        width: width || '100%',
        maxWidth: width || undefined,
        margin: width ? undefined : '0 auto',
        borderRadius: height / 2,
        background: isEmpty ? 'transparent' : LEVELS[level].fill,
        border: isEmpty ? `2px dashed var(--ink-muted)` : 'none',
        boxShadow: isEmpty ? 'none' : 'inset 0 0 0 1px rgba(0,0,0,0.06)',
      }}
    />
  );
}

function LevelSegmentedPill({ level, onSet, height = 26 }) {
  const theme = useContext(ThemeContext);
  const isHc = theme === 'hc';
  const isEmpty = level === 0;
  const currentColor = level === 1 ? 'var(--coral)' : level === 2 ? 'var(--green)' : level === 3 ? 'var(--indigo)' : 'transparent';

  return (
    <div
      style={{
        display: 'flex',
        height,
        width: '100%',
        borderRadius: height / 2,
        overflow: 'hidden',
        border: isEmpty ? '2px dashed var(--ink-muted)' : isHc ? `2px solid ${currentColor}` : 'none',
        background: isHc ? CARD : 'transparent',
        boxShadow: isEmpty || isHc ? 'none' : 'inset 0 0 0 1px rgba(0,0,0,0.06)',
      }}
    >
      {[1, 2, 3].map((seg) => {
        const filled = level >= seg;
        const segColor = seg === 1 ? 'var(--coral)' : seg === 2 ? 'var(--green)' : 'var(--indigo)';
        return (
          <button
            key={seg}
            onClick={(e) => { e.stopPropagation(); onSet(level === seg ? 0 : seg); }}
            title={LEVELS[seg].name}
            style={{
              flex: 1,
              border: 'none',
              borderLeft: seg > 1 ? `1px solid ${filled ? 'rgba(255,255,255,0.4)' : 'var(--line)'}` : 'none',
              cursor: 'pointer',
              padding: 0,
              background: filled ? (isHc ? CARD : currentColor) : 'transparent',
              color: isHc ? segColor : 'transparent',
              fontSize: Math.max(9, Math.round(height * 0.42)),
              fontWeight: 700,
              lineHeight: 1,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}
          >
            {isHc && filled && seg === level ? LEVEL_LETTERS[seg] : ''}
          </button>
        );
      })}
    </div>
  );
}

export default function App() {
  const [ready, setReady] = useState(false);
  const [theme, setTheme] = useState('pale');
  const [structure, setStructure] = useState(emptyStructure);
  const [studentsByClass, setStudentsByClass] = useState({});
  const [records, setRecords] = useState({});
  const [fieldValues, setFieldValues] = useState({});
  const [fieldModal, setFieldModal] = useState(null); // { fieldId, studentId, type }
  const [view, setView] = useState('assess');
  const [classId, setClassId] = useState(null);
  const [termId, setTermId] = useState(null);
  const [subjectId, setSubjectId] = useState(null);
  const [topicFilterId, setTopicFilterId] = useState('');
  const [assessMode, setAssessMode] = useState('grid');
  const [focusIndex, setFocusIndex] = useState(0);
  const [tapStudentIndex, setTapStudentIndex] = useState(0);
  const [setupSubjectId, setSetupSubjectId] = useState(null);
  const [setupClassId, setSetupClassId] = useState(null);
  const [exporting, setExporting] = useState(false);
  const [importing, setImporting] = useState(false);
  const [importMsg, setImportMsg] = useState(null);

  useEffect(() => {
    (async () => {
      const [s, st, r, fv, savedTheme] = await Promise.all([
        loadKey('structure', emptyStructure),
        loadKey('students', {}),
        loadKey('records', {}),
        loadKey('fieldValues', {}),
        loadKey('theme', 'pale', false),
      ]);
      setStructure({ ...s, fields: s.fields || [], yearGroups: s.yearGroups || [], departments: s.departments || [] });
      setStudentsByClass(st);
      setRecords(r);
      setFieldValues(fv);
      setTheme(savedTheme);
      setClassId(s.classes[0]?.id ?? null);
      setTermId(s.terms[0]?.id ?? null);
      setSubjectId(s.subjects[0]?.id ?? null);
      setSetupClassId(s.classes[0]?.id ?? null);
      setSetupSubjectId(s.subjects[0]?.id ?? null);
      setReady(true);
    })();
  }, []);

  useDebouncedSave('structure', structure, ready);
  useDebouncedSave('students', studentsByClass, ready);
  useDebouncedSave('records', records, ready, 300);
  useDebouncedSave('fieldValues', fieldValues, ready, 300);
  useDebouncedSave('theme', theme, ready, 200, false);

  // keep selections valid
  useEffect(() => {
    if (!ready) return;
    if (!structure.classes.find((c) => c.id === classId)) setClassId(structure.classes[0]?.id ?? null);
    if (!structure.terms.find((t) => t.id === termId)) setTermId(structure.terms[0]?.id ?? null);
    if (!structure.subjects.find((s) => s.id === subjectId)) setSubjectId(structure.subjects[0]?.id ?? null);
    if (!structure.classes.find((c) => c.id === setupClassId)) setSetupClassId(structure.classes[0]?.id ?? null);
    if (!structure.subjects.find((s) => s.id === setupSubjectId)) setSetupSubjectId(structure.subjects[0]?.id ?? null);
  }, [structure, ready]);

  const students = studentsByClass[classId] || [];

  const subjectsForTerm = useMemo(() => {
    if (!termId) return structure.subjects;
    const withData = structure.subjects.filter((s) => {
      const hasTopic = structure.topics.some((t) => t.subjectId === s.id && (!t.termId || t.termId === termId));
      const hasGeneralCriteria = structure.criteria.some((c) => c.subjectId === s.id && !c.topicId && (!c.termId || c.termId === termId));
      return hasTopic || hasGeneralCriteria;
    });
    return withData.length ? withData : structure.subjects;
  }, [structure.subjects, structure.topics, structure.criteria, termId]);

  useEffect(() => {
    if (!ready) return;
    if (subjectsForTerm.length && !subjectsForTerm.find((s) => s.id === subjectId)) {
      setSubjectId(subjectsForTerm[0].id);
    }
  }, [termId, subjectsForTerm, ready]);

  const criteriaForSubject = useMemo(
    () => structure.criteria.filter((c) => c.subjectId === subjectId),
    [structure.criteria, subjectId]
  );
  const topicsForSubject = useMemo(
    () => structure.topics.filter((t) => t.subjectId === subjectId && (!t.termId || t.termId === termId)),
    [structure.topics, subjectId, termId]
  );
  const orderedColumns = useMemo(() => {
    const withTopic = topicsForSubject.flatMap((t) => {
      const crit = criteriaForSubject.filter((c) => c.topicId === t.id).map((c) => ({ ...c, topicName: t.name, kind: 'criteria' }));
      const fields = structure.fields.filter((f) => f.topicId === t.id).map((f) => ({ ...f, topicName: t.name, kind: 'field' }));
      return [...crit, ...fields];
    });
    const noTopic = criteriaForSubject
      .filter((c) => !c.topicId && (!c.termId || c.termId === termId))
      .map((c) => ({ ...c, topicName: 'General', kind: 'criteria' }));
    return [...withTopic, ...noTopic];
  }, [topicsForSubject, criteriaForSubject, structure.fields, termId]);

  const criteriaColumns = useMemo(() => orderedColumns.filter((c) => c.kind === 'criteria'), [orderedColumns]);

  const filteredColumns = useMemo(
    () => (topicFilterId ? orderedColumns.filter((c) => (c.topicId || 'general') === topicFilterId) : orderedColumns),
    [orderedColumns, topicFilterId]
  );
  const filteredCriteriaColumns = useMemo(() => filteredColumns.filter((c) => c.kind === 'criteria'), [filteredColumns]);

  useEffect(() => { setFocusIndex(0); setTapStudentIndex(0); }, [subjectId, termId, topicFilterId, classId]);
  useEffect(() => {
    if (focusIndex >= filteredCriteriaColumns.length) setFocusIndex(Math.max(0, filteredCriteriaColumns.length - 1));
  }, [filteredCriteriaColumns, focusIndex]);
  useEffect(() => { setTopicFilterId(''); }, [subjectId, termId]);
  useEffect(() => {
    if (tapStudentIndex >= students.length) setTapStudentIndex(Math.max(0, students.length - 1));
  }, [students, tapStudentIndex]);

  const recKey = (studentId, term, critId) => `${studentId}|${term}|${critId}`;
  const fieldKey = (fieldId, studentId) => `${fieldId}|${studentId}`;
  const photoStorageKey = (fieldId, studentId) => `photo:${fieldId}:${studentId}`;

  const cycleLevel = (studentId, critId) => {
    const key = recKey(studentId, termId, critId);
    setRecords((prev) => {
      const cur = prev[key] ?? 0;
      const next = (cur + 1) % LEVELS.length;
      return { ...prev, [key]: next };
    });
  };

  const setLevel = (studentId, critId, newLevel) => {
    const key = recKey(studentId, termId, critId);
    setRecords((prev) => ({ ...prev, [key]: prev[key] === newLevel ? 0 : newLevel }));
  };

  const bulkSetLevel = (critId, level) => {
    setRecords((prev) => {
      const next = { ...prev };
      students.forEach((stu) => { next[recKey(stu.id, termId, critId)] = level; });
      return next;
    });
  };

  const bulkSetLevelForStudent = (studentId, criteriaIds, level) => {
    setRecords((prev) => {
      const next = { ...prev };
      criteriaIds.forEach((critId) => { next[recKey(studentId, termId, critId)] = level; });
      return next;
    });
  };

  const setFieldValue = (fieldId, studentId, value) => {
    setFieldValues((prev) => ({ ...prev, [fieldKey(fieldId, studentId)]: value }));
  };

  const handlePhotoFile = async (fieldId, studentId, file) => {
    try {
      const dataUrl = await compressImage(file);
      await saveKey(photoStorageKey(fieldId, studentId), dataUrl);
      setFieldValue(fieldId, studentId, true);
    } catch (e) {
      console.error('photo save failed', e);
    }
  };
  const removePhoto = async (fieldId, studentId) => {
    try { await window.storage.delete(photoStorageKey(fieldId, studentId), true); } catch (e) {}
    setFieldValue(fieldId, studentId, false);
    setFieldModal(null);
  };

  const coverageFor = (studentId) => {
    if (criteriaColumns.length === 0) return null;
    const total = criteriaColumns.reduce((sum, c) => sum + (records[recKey(studentId, termId, c.id)] ?? 0), 0);
    return Math.round((total / (criteriaColumns.length * 3)) * 100);
  };

  // ---- setup mutators ----
  const addItem = (arrKey, extra = {}) =>
    setStructure((s) => ({ ...s, [arrKey]: [...s[arrKey], { id: uid(), ...extra }] }));
  const removeItem = (arrKey, id) =>
    setStructure((s) => ({ ...s, [arrKey]: s[arrKey].filter((x) => x.id !== id) }));

  const addClass = (name) => {
    const id = uid();
    setStructure((s) => ({ ...s, classes: [...s.classes, { id, name }] }));
    setStudentsByClass((sc) => ({ ...sc, [id]: [] }));
    setSetupClassId(id);
  };
  const removeClass = (id) => {
    setStructure((s) => ({ ...s, classes: s.classes.filter((c) => c.id !== id) }));
    setStudentsByClass((sc) => {
      const copy = { ...sc };
      delete copy[id];
      return copy;
    });
  };
  const togglePinnedSubject = (classId, subjectId) => {
    setStructure((s) => ({
      ...s,
      classes: s.classes.map((c) => {
        if (c.id !== classId) return c;
        const pinned = c.pinnedSubjectIds || [];
        const has = pinned.includes(subjectId);
        return { ...c, pinnedSubjectIds: has ? pinned.filter((id) => id !== subjectId) : [...pinned, subjectId] };
      }),
    }));
  };
  const addStudent = (cId, name) =>
    setStudentsByClass((sc) => ({ ...sc, [cId]: [...(sc[cId] || []), { id: uid(), name }] }));
  const removeStudent = (cId, sId) =>
    setStudentsByClass((sc) => ({ ...sc, [cId]: (sc[cId] || []).filter((s) => s.id !== sId) }));
  const setStudentBirthTerm = (cId, sId, birthTerm) =>
    setStudentsByClass((sc) => ({ ...sc, [cId]: (sc[cId] || []).map((s) => (s.id === sId ? { ...s, birthTerm: birthTerm || undefined } : s)) }));
  const setClassYearGroup = (classId, yearGroup) =>
    setStructure((s) => ({ ...s, classes: s.classes.map((c) => (c.id === classId ? { ...c, yearGroup: yearGroup || undefined } : c)) }));

  const addSubject = (name) => {
    const id = uid();
    setStructure((s) => ({ ...s, subjects: [...s.subjects, { id, name }] }));
    setSetupSubjectId(id);
  };
  const removeSubject = (id) =>
    setStructure((s) => {
      const removedTopicIds = s.topics.filter((t) => t.subjectId === id).map((t) => t.id);
      return {
        ...s,
        subjects: s.subjects.filter((x) => x.id !== id),
        topics: s.topics.filter((t) => t.subjectId !== id),
        criteria: s.criteria.filter((c) => c.subjectId !== id),
        fields: s.fields.filter((f) => !removedTopicIds.includes(f.topicId)),
      };
    });
  const setSubjectDepartment = (subjectId, department) =>
    setStructure((s) => ({ ...s, subjects: s.subjects.map((sub) => (sub.id === subjectId ? { ...sub, department: department || undefined } : sub)) }));
  const addTopic = (subId, termId2, name) => addItem('topics', { subjectId: subId, termId: termId2 || null, name });
  const removeTopic = (id) =>
    setStructure((s) => ({
      ...s,
      topics: s.topics.filter((t) => t.id !== id),
      criteria: s.criteria.filter((c) => c.topicId !== id),
      fields: s.fields.filter((f) => f.topicId !== id),
    }));
  const addCriteria = (subId, topicId, name, termId2) => addItem('criteria', { subjectId: subId, topicId: topicId || null, termId: (!topicId && termId2) ? termId2 : null, name });
  const addField = (topicId, name, type) => addItem('fields', { topicId, name, type });
  const removeField = (id) => removeItem('fields', id);
  const [showResetConfirm, setShowResetConfirm] = useState(false);
  const resetAll = () => setShowResetConfirm(true);
  const performReset = () => {
    setStructure(emptyStructure);
    setStudentsByClass({});
    setRecords({});
    setFieldValues({});
    setShowResetConfirm(false);
  };

  // ---- bulk import ----
  const downloadTemplate = () => {
    const wb = XLSX.utils.book_new();
    const readme = XLSX.utils.aoa_to_sheet([
      ['How this sheet works'],
      [''],
      ['One row = one thing. Don\u2019t put a student AND a success criteria on the same row.'],
      [''],
      ['To add a STUDENT: fill in Class + Student only. Leave every other column blank.'],
      ['To add a SUCCESS CRITERIA: fill in Subject + Success criteria (Term/Topic optional). Leave Class + Student blank.'],
      ['To add an EXTRA COLUMN (Photo, Rating, Date, Note): fill in Subject + Topic + Extra column name + Extra column type. Leave Success criteria blank.'],
      [''],
      ['Go to the Import tab, delete the EXAMPLE rows, and add your own below the headers.'],
    ]);
    readme['!cols'] = [{ wch: 100 }];
    XLSX.utils.book_append_sheet(wb, readme, 'Read me first');

    const ws = XLSX.utils.aoa_to_sheet([
      ['Class', 'Student', 'Term', 'Subject', 'Topic', 'Success criteria', 'Extra column name', 'Extra column type'],
      ['EXAMPLE — delete me', 'Jane Smith', '', '', '', '', '', ''],
      ['EXAMPLE — delete me', 'Sam Patel', '', '', '', '', '', ''],
      ['', '', 'Term 1', 'EXAMPLE — delete me', '', 'Understands the design brief', '', ''],
      ['', '', 'Term 2', 'EXAMPLE — delete me', 'TinkerCAD', 'Can create a basic 3D model', '', ''],
      ['', '', 'Term 2', 'EXAMPLE — delete me', 'TinkerCAD', 'Can apply materials and colour', '', ''],
      ['', '', 'Term 2', 'EXAMPLE — delete me', 'TinkerCAD', '', 'Model level', 'Rating'],
      ['', '', 'Term 2', 'EXAMPLE — delete me', 'TinkerCAD', '', 'Photo evidence', 'Photo'],
      ['', '', 'Term 2', 'EXAMPLE — delete me', 'TinkerCAD', '', 'Date seen', 'Date'],
      ['', '', 'Term 3', '', '', '', '', ''],
    ]);
    ws['!cols'] = [{ wch: 18 }, { wch: 18 }, { wch: 10 }, { wch: 20 }, { wch: 14 }, { wch: 34 }, { wch: 18 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws, 'Import');
    XLSX.writeFile(wb, 'coverage-import-template.xlsx');
  };

  const FIELD_TYPE_ALIASES = {
    note: 'note', text: 'note', comment: 'note', notes: 'note',
    rating: 'rating', star: 'rating', stars: 'rating', number: 'rating', score: 'rating',
    date: 'date',
    photo: 'photo', image: 'photo', picture: 'photo', camera: 'photo',
  };
  const mapFieldType = (raw) => FIELD_TYPE_ALIASES[raw.trim().toLowerCase()] || 'note';

  const cell = (row, ...names) => {
    for (const n of names) {
      const key = Object.keys(row).find((k) => k.trim().toLowerCase() === n);
      if (key && String(row[key]).trim()) return String(row[key]).trim();
    }
    return '';
  };
  const findOrCreate = (arr, matchFn, createFn) => {
    let item = arr.find(matchFn);
    if (!item) { item = createFn(); arr.push(item); }
    return item;
  };

  // ---- paste importers ----
  const importRoster = (text) => {
    const blocks = text
      .split(/\n\s*\n/)
      .map((b) => b.split('\n').map((l) => l.trim()).filter(Boolean))
      .filter((b) => b.length > 0);
    const newClasses = [...structure.classes];
    const newStudentsByClass = { ...studentsByClass };
    let addedClasses = 0, addedStudents = 0;
    blocks.forEach((lines) => {
      const className = lines[0];
      if (!className) return;
      const beforeLen = newClasses.length;
      const cls = findOrCreate(newClasses, (c) => c.name.toLowerCase() === className.toLowerCase(), () => ({ id: uid(), name: className }));
      if (newClasses.length > beforeLen) addedClasses++;
      if (!newStudentsByClass[cls.id]) newStudentsByClass[cls.id] = [];
      lines.slice(1).forEach((name) => {
        const exists = newStudentsByClass[cls.id].some((s) => s.name.toLowerCase() === name.toLowerCase());
        if (!exists) {
          newStudentsByClass[cls.id] = [...newStudentsByClass[cls.id], { id: uid(), name }];
          addedStudents++;
        }
      });
    });
    setStructure((s) => ({ ...s, classes: newClasses }));
    setStudentsByClass(newStudentsByClass);
    return { addedClasses, addedStudents };
  };

  const saveTopicContent = ({ termName, subjectName, topicName, criteriaText, columns }) => {
    const subj = subjectName.trim();
    if (!subj) return { error: 'Enter a subject name.' };
    const top = topicName.trim();
    const validColumns = columns.filter((c) => c.name.trim());
    if (validColumns.length && !top) return { error: 'Extra columns need a topic — enter one above.' };

    const newSubjects = [...structure.subjects];
    const newTerms = [...structure.terms];
    const newTopics = [...structure.topics];
    const newCriteria = [...structure.criteria];
    const newFields = [...structure.fields];

    let rowTermId = null;
    const tName = termName.trim();
    if (tName) {
      const term = findOrCreate(newTerms, (t) => t.name.toLowerCase() === tName.toLowerCase(), () => ({ id: uid(), name: tName }));
      rowTermId = term.id;
    }
    const subject = findOrCreate(newSubjects, (s) => s.name.toLowerCase() === subj.toLowerCase(), () => ({ id: uid(), name: subj }));
    let topicId = null;
    if (top) {
      const topic = findOrCreate(
        newTopics,
        (t) => t.subjectId === subject.id && (t.termId || null) === rowTermId && t.name.toLowerCase() === top.toLowerCase(),
        () => ({ id: uid(), subjectId: subject.id, termId: rowTermId, name: top })
      );
      topicId = topic.id;
    }

    let addedCriteria = 0;
    criteriaText.split('\n').map((l) => l.trim()).filter(Boolean).forEach((line) => {
      const before = newCriteria.length;
      findOrCreate(
        newCriteria,
        (c) => c.subjectId === subject.id && (c.topicId || null) === topicId && c.name.toLowerCase() === line.toLowerCase(),
        () => ({ id: uid(), subjectId: subject.id, topicId, termId: topicId ? null : rowTermId, name: line })
      );
      if (newCriteria.length > before) addedCriteria++;
    });

    let addedFields = 0;
    if (topicId) {
      validColumns.forEach((c) => {
        const before = newFields.length;
        findOrCreate(newFields, (f) => f.topicId === topicId && f.name.toLowerCase() === c.name.trim().toLowerCase(), () => ({ id: uid(), topicId, name: c.name.trim(), type: c.type }));
        if (newFields.length > before) addedFields++;
      });
    }

    setStructure({ classes: structure.classes, subjects: newSubjects, terms: newTerms, topics: newTopics, criteria: newCriteria, fields: newFields });
    return { addedCriteria, addedFields };
  };

  const importFile = async (file) => {
    setImporting(true);
    setImportMsg(null);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array' });
      const rows = wb.SheetNames.flatMap((name) => XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' }));

      const newClasses = [...structure.classes];
      const newSubjects = [...structure.subjects];
      const newTerms = [...structure.terms];
      const newTopics = [...structure.topics];
      const newCriteria = [...structure.criteria];
      const newFields = [...structure.fields];
      const newStudentsByClass = { ...studentsByClass };

      let addedStudents = 0, addedClasses = 0, addedTerms = 0, addedTopics = 0, addedCriteria = 0, addedFields = 0;

      rows.forEach((row) => {
        const className = cell(row, 'class');
        const studentName = cell(row, 'student', 'name');
        const subjectName = cell(row, 'subject');
        const topicName = cell(row, 'topic');
        const critName = cell(row, 'success criteria', 'criteria', 'success criterion', 'learning intention');
        const termName = cell(row, 'term');
        const extraColName = cell(row, 'extra column name', 'column name', 'field name');
        const extraColType = cell(row, 'extra column type', 'column type', 'field type');

        let rowTermId = null;
        if (termName) {
          const beforeTerms = newTerms.length;
          const term = findOrCreate(newTerms, (t) => t.name.toLowerCase() === termName.toLowerCase(), () => ({ id: uid(), name: termName }));
          if (newTerms.length > beforeTerms) addedTerms++;
          rowTermId = term.id;
        }

        if (className && studentName) {
          const beforeClasses = newClasses.length;
          const cls = findOrCreate(newClasses, (c) => c.name.toLowerCase() === className.toLowerCase(), () => ({ id: uid(), name: className }));
          if (newClasses.length > beforeClasses) addedClasses++;
          if (!newStudentsByClass[cls.id]) newStudentsByClass[cls.id] = [];
          const exists = newStudentsByClass[cls.id].some((s) => s.name.toLowerCase() === studentName.toLowerCase());
          if (!exists) {
            newStudentsByClass[cls.id] = [...newStudentsByClass[cls.id], { id: uid(), name: studentName }];
            addedStudents++;
          }
        }

        if (subjectName && extraColName && topicName) {
          const subject = findOrCreate(newSubjects, (s) => s.name.toLowerCase() === subjectName.toLowerCase(), () => ({ id: uid(), name: subjectName }));
          const beforeTopics = newTopics.length;
          const topic = findOrCreate(
            newTopics,
            (t) => t.subjectId === subject.id && (t.termId || null) === rowTermId && t.name.toLowerCase() === topicName.toLowerCase(),
            () => ({ id: uid(), subjectId: subject.id, termId: rowTermId, name: topicName })
          );
          if (newTopics.length > beforeTopics) addedTopics++;
          const beforeFields = newFields.length;
          findOrCreate(
            newFields,
            (f) => f.topicId === topic.id && f.name.toLowerCase() === extraColName.toLowerCase(),
            () => ({ id: uid(), topicId: topic.id, name: extraColName, type: mapFieldType(extraColType) })
          );
          if (newFields.length > beforeFields) addedFields++;
        }

        if (subjectName && critName) {
          const subject = findOrCreate(newSubjects, (s) => s.name.toLowerCase() === subjectName.toLowerCase(), () => ({ id: uid(), name: subjectName }));
          let topicId = null;
          if (topicName) {
            const beforeTopics = newTopics.length;
            const topic = findOrCreate(
              newTopics,
              (t) => t.subjectId === subject.id && (t.termId || null) === rowTermId && t.name.toLowerCase() === topicName.toLowerCase(),
              () => ({ id: uid(), subjectId: subject.id, termId: rowTermId, name: topicName })
            );
            if (newTopics.length > beforeTopics) addedTopics++;
            topicId = topic.id;
          }
          const beforeLen = newCriteria.length;
          findOrCreate(
            newCriteria,
            (c) => c.subjectId === subject.id && (c.topicId || null) === topicId && c.name.toLowerCase() === critName.toLowerCase(),
            () => ({ id: uid(), subjectId: subject.id, topicId, termId: topicId ? null : rowTermId, name: critName })
          );
          if (newCriteria.length > beforeLen) addedCriteria++;
        }
      });

      setStructure({ classes: newClasses, subjects: newSubjects, terms: newTerms, topics: newTopics, criteria: newCriteria, fields: newFields });
      setStudentsByClass(newStudentsByClass);
      const parts = [];
      if (addedStudents) parts.push(`${addedStudents} student${addedStudents === 1 ? '' : 's'}`);
      if (addedClasses) parts.push(`${addedClasses} new class${addedClasses === 1 ? '' : 'es'}`);
      if (addedTerms) parts.push(`${addedTerms} term${addedTerms === 1 ? '' : 's'}`);
      if (addedTopics) parts.push(`${addedTopics} topic${addedTopics === 1 ? '' : 's'}`);
      if (addedCriteria) parts.push(`${addedCriteria} success criteria`);
      if (addedFields) parts.push(`${addedFields} extra column${addedFields === 1 ? '' : 's'}`);
      setImportMsg(parts.length ? `Imported ${parts.join(', ')}.` : 'No new rows found — check your column headers match the template.');
    } catch (e) {
      setImportMsg('Could not read that file. Make sure it is the downloaded template, filled in.');
    } finally {
      setImporting(false);
    }
  };

  // ---- export ----
  const runExport = async () => {
    setExporting(true);
    try {
      const wb = XLSX.utils.book_new();
      const terms = structure.terms.length ? structure.terms : [{ id: '_none', name: 'All' }];

      const cols = [];
      terms.forEach((term) => {
        structure.subjects.forEach((subj) => {
          const subjCriteria = structure.criteria.filter((c) => c.subjectId === subj.id);
          const subjTopics = structure.topics.filter((t) => t.subjectId === subj.id && (!t.termId || t.termId === term.id));
          const grouped = [
            ...subjTopics.map((t) => ({
              topicName: t.name,
              criteria: subjCriteria.filter((c) => c.topicId === t.id),
              fields: structure.fields.filter((f) => f.topicId === t.id),
            })),
            { topicName: 'General', criteria: subjCriteria.filter((c) => !c.topicId && (!c.termId || c.termId === term.id)), fields: [] },
          ].filter((g) => g.criteria.length || g.fields.length);
          grouped.forEach((g) => {
            g.criteria.forEach((c) => cols.push({ kind: 'criteria', termName: term.name, subjectName: subj.name, topicName: g.topicName, colName: c.name, termId: term.id, critId: c.id }));
            g.fields.forEach((f) => cols.push({ kind: 'field', termName: term.name, subjectName: subj.name, topicName: g.topicName, colName: f.name, fieldId: f.id, type: f.type }));
          });
        });
      });
      const criteriaColsAll = cols.filter((c) => c.kind === 'criteria');

      // Cross-class comparison sheet — flat, single header row, no merges, so a couple of
      // columns can be selected and turned into a chart in Excel/Sheets in a few clicks.
      if (structure.classes.length && criteriaColsAll.length) {
        const compHeader = ['Class', 'Students', 'Avg coverage %', ...criteriaColsAll.map((c) => `${c.termName} · ${c.subjectName} > ${c.topicName} > ${c.colName}`)];
        const compRows = [compHeader];
        structure.classes.forEach((cls) => {
          const roster = studentsByClass[cls.id] || [];
          if (!roster.length) return;
          const avgCov = Math.round(roster.reduce((sum, stu) => {
            const t = criteriaColsAll.reduce((s2, c) => s2 + (records[recKey(stu.id, c.termId, c.critId)] ?? 0), 0);
            return sum + (t / (criteriaColsAll.length * 3)) * 100;
          }, 0) / roster.length);
          const pctCells = criteriaColsAll.map((c) => {
            const met = roster.filter((stu) => (records[recKey(stu.id, c.termId, c.critId)] ?? 0) >= 2).length;
            return Math.round((met / roster.length) * 100);
          });
          compRows.push([cls.name, roster.length, avgCov, ...pctCells]);
        });
        const compWs = XLSX.utils.aoa_to_sheet(compRows);
        compWs['!cols'] = [{ wch: 14 }, { wch: 10 }, { wch: 14 }, ...criteriaColsAll.map(() => ({ wch: 30 }))];
        XLSX.utils.book_append_sheet(wb, compWs, 'Class Comparison');
      }

      structure.classes.forEach((cls) => {
        const row0 = ['Student', 'Coverage %', 'Flag', ...cols.map((c) => c.termName)];
        const row1 = ['', '', '', ...cols.map((c) => c.subjectName)];
        const row2 = ['', '', '', ...cols.map((c) => c.topicName)];
        const row3 = ['', '', '', ...cols.map((c) => c.colName)];
        const rows = [row0, row1, row2, row3];

        const criteriaCols = cols.filter((c) => c.kind === 'criteria');
        const computeExportFlag = (studentId) => {
          if (!criteriaCols.length) return '';
          let total = 0, hasAny = false;
          criteriaCols.forEach((c) => {
            const lvl = records[recKey(studentId, c.termId, c.critId)] ?? 0;
            total += lvl;
            if (lvl > 0) hasAny = true;
          });
          const pct = (total / (criteriaCols.length * 3)) * 100;
          if (pct >= HIGH_ACHIEVER_PCT) return 'High achiever';
          if (hasAny && pct < NEEDS_PUSH_PCT) return 'Needs a push';
          return '';
        };
        const roster = studentsByClass[cls.id] || [];
        roster.forEach((stu) => {
          const cells = cols.map((c) => {
            if (c.kind === 'criteria') return LEVELS[records[recKey(stu.id, c.termId, c.critId)] ?? 0].name;
            const v = fieldValues[fieldKey(c.fieldId, stu.id)];
            if (c.type === 'photo') return v ? 'Yes' : '';
            if (c.type === 'rating') return v || '';
            return v || '';
          });
          const covTotal = criteriaCols.length ? criteriaCols.reduce((sum, c) => sum + (records[recKey(stu.id, c.termId, c.critId)] ?? 0), 0) : 0;
          const cov = criteriaCols.length ? Math.round((covTotal / (criteriaCols.length * 3)) * 100) : '';
          rows.push([stu.name, cov, computeExportFlag(stu.id), ...cells]);
        });

        if (roster.length) {
          const avgCov = Math.round(roster.reduce((sum, stu) => {
            const t = criteriaCols.reduce((s2, c) => s2 + (records[recKey(stu.id, c.termId, c.critId)] ?? 0), 0);
            return sum + (criteriaCols.length ? (t / (criteriaCols.length * 3)) * 100 : 0);
          }, 0) / roster.length);
          const summaryCells = cols.map((c) => {
            if (c.kind !== 'criteria') return '';
            const counts = [0, 0, 0, 0];
            roster.forEach((stu) => { counts[records[recKey(stu.id, c.termId, c.critId)] ?? 0]++; });
            const met = counts[2] + counts[3];
            const pct = Math.round((met / roster.length) * 100);
            return `N:${counts[0]} D:${counts[1]} S:${counts[2]} M:${counts[3]} · ${pct}% met`;
          });
          rows.push(['Class breakdown', avgCov, '', ...summaryCells]);
        }

        const ws = XLSX.utils.aoa_to_sheet(rows);
        const merges = [
          { s: { r: 0, c: 0 }, e: { r: 3, c: 0 } },
          { s: { r: 0, c: 1 }, e: { r: 3, c: 1 } },
          { s: { r: 0, c: 2 }, e: { r: 3, c: 2 } },
        ];
        [0, 1, 2].forEach((r) => {
          let start = 3;
          for (let c = 4; c <= cols.length + 2; c++) {
            const curVal = c <= cols.length + 2 ? rows[r][c] : Symbol('end');
            const prevVal = rows[r][c - 1];
            if (curVal !== prevVal) {
              if (c - 1 > start) merges.push({ s: { r, c: start }, e: { r, c: c - 1 } });
              start = c;
            }
          }
        });
        ws['!merges'] = merges;
        ws['!cols'] = [{ wch: 20 }, { wch: 10 }, { wch: 14 }, ...cols.map(() => ({ wch: 14 }))];
        const sheetName = cls.name.slice(0, 31) || 'Class';
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
      });
      if (structure.classes.length === 0) {
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['No classes yet']]), 'Empty');
      }
      XLSX.writeFile(wb, 'attainment-tracker.xlsx');
    } finally {
      setExporting(false);
    }
  };

  if (!ready) {
    return (
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 40, color: MUTED, fontSize: 14 }}>
        <Loader2 size={16} className="spin" />
        Loading tracker…
        <style>{`.spin{animation:spin 1s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
      </div>
    );
  }

  return (
    <ThemeContext.Provider value={theme}>
    <div data-theme={theme} style={{ fontFamily: 'ui-sans-serif, system-ui, sans-serif', background: PAPER, color: INK, minHeight: 400, borderRadius: 12, overflow: 'hidden', border: `var(--border-width) solid ${LINE}` }}>
      <style>{THEME_CSS}</style>
      {/* header */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 18px', borderBottom: `var(--border-width) solid ${LINE}`, background: CARD, flexWrap: 'wrap', gap: 10 }}>
        <div style={{ display: 'flex', alignItems: 'baseline', gap: 10 }}>
          <span style={{ fontFamily: 'Georgia, serif', fontSize: 19, fontWeight: 700, letterSpacing: -0.3, color: INK }}>Your Classroom Helper</span>
          <span style={{ fontSize: 12, color: MUTED }}>assess in a tap</span>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <ThemeToggle theme={theme} setTheme={setTheme} />
          <div style={{ display: 'flex', background: PAPER, borderRadius: 999, padding: 3, border: `var(--border-width) solid ${LINE}` }}>
            <Pill active={view === 'assess'} onClick={() => setView('assess')}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}><ClipboardList size={13} /> Assess</span>
            </Pill>
            <Pill active={view === 'analytics'} onClick={() => setView('analytics')}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}><BarChart3 size={13} /> Analytics</span>
            </Pill>
            <Pill active={view === 'setup'} onClick={() => setView('setup')}>
              <span style={{ display: 'flex', alignItems: 'center', gap: 5 }}><Settings2 size={13} /> Setup</span>
            </Pill>
          </div>
          <button
            onClick={runExport}
            disabled={exporting}
            style={{
              display: 'flex', alignItems: 'center', gap: 6, padding: '7px 13px', borderRadius: 8,
              border: 'none', background: ACCENT, color: '#fff', fontSize: 13, fontWeight: 500, cursor: 'pointer',
            }}
          >
            {exporting ? <Loader2 size={14} className="spin" /> : <Download size={14} />}
            Export workbook
          </button>
        </div>
      </div>

      {view === 'assess' ? (
        <AssessView
          structure={structure}
          students={students}
          classId={classId} setClassId={setClassId}
          termId={termId} setTermId={setTermId}
          subjectId={subjectId} setSubjectId={setSubjectId}
          subjectsForTerm={subjectsForTerm}
          topicsForSubject={topicsForSubject}
          topicFilterId={topicFilterId} setTopicFilterId={setTopicFilterId}
          orderedColumns={orderedColumns}
          filteredColumns={filteredColumns}
          filteredCriteriaColumns={filteredCriteriaColumns}
          records={records}
          recKey={recKey}
          cycleLevel={cycleLevel}
          setLevel={setLevel}
          bulkSetLevel={bulkSetLevel}
          bulkSetLevelForStudent={bulkSetLevelForStudent}
          coverageFor={coverageFor}
          setView={setView}
          assessMode={assessMode} setAssessMode={setAssessMode}
          focusIndex={focusIndex} setFocusIndex={setFocusIndex}
          tapStudentIndex={tapStudentIndex} setTapStudentIndex={setTapStudentIndex}
          fieldValues={fieldValues}
          fieldKey={fieldKey}
          setFieldValue={setFieldValue}
          handlePhotoFile={handlePhotoFile}
          removePhoto={removePhoto}
          fieldModal={fieldModal}
          setFieldModal={setFieldModal}
          addField={addField}
        />
      ) : view === 'analytics' ? (
        <AnalyticsView structure={structure} studentsByClass={studentsByClass} records={records} recKey={recKey} fieldValues={fieldValues} fieldKey={fieldKey} />
      ) : (
        <SetupView
          structure={structure}
          studentsByClass={studentsByClass}
          setupClassId={setupClassId} setSetupClassId={setSetupClassId}
          setupSubjectId={setupSubjectId} setSetupSubjectId={setSetupSubjectId}
          addClass={addClass} removeClass={removeClass}
          addStudent={addStudent} removeStudent={removeStudent} setStudentBirthTerm={setStudentBirthTerm}
          setClassYearGroup={setClassYearGroup}
          addSubject={addSubject} removeSubject={removeSubject} setSubjectDepartment={setSubjectDepartment}
          addTopic={addTopic} removeTopic={removeTopic}
          addCriteria={addCriteria} removeCriteria={(id) => removeItem('criteria', id)}
          addField={addField} removeField={removeField}
          togglePinnedSubject={togglePinnedSubject}
          addTerm={(name) => addItem('terms', { name })}
          removeTerm={(id) => removeItem('terms', id)}
          addYearGroup={(name) => addItem('yearGroups', { name })}
          removeYearGroup={(id) => removeItem('yearGroups', id)}
          addDepartment={(name) => addItem('departments', { name })}
          removeDepartment={(id) => removeItem('departments', id)}
          downloadTemplate={downloadTemplate}
          importFile={importFile}
          importing={importing}
          importMsg={importMsg}
          setImportMsg={setImportMsg}
          resetAll={resetAll}
          importRoster={importRoster}
          saveTopicContent={saveTopicContent}
        />
      )}
      {showResetConfirm && (
        <ConfirmModal
          title="Clear all data?"
          message="This removes every class, student, subject, criteria, extra column and recorded attainment. This cannot be undone."
          confirmLabel="Clear everything"
          danger
          onConfirm={performReset}
          onCancel={() => setShowResetConfirm(false)}
        />
      )}
      <style>{`.spin{animation:spin 1s linear infinite}@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
    </ThemeContext.Provider>
  );
}

function AssessView({ structure, students, classId, setClassId, termId, setTermId, subjectId, setSubjectId, subjectsForTerm, topicsForSubject, topicFilterId, setTopicFilterId, orderedColumns, filteredColumns, filteredCriteriaColumns, records, recKey, cycleLevel, setLevel, bulkSetLevel, bulkSetLevelForStudent, coverageFor, setView, assessMode, setAssessMode, focusIndex, setFocusIndex, tapStudentIndex, setTapStudentIndex, fieldValues, fieldKey, setFieldValue, handlePhotoFile, removePhoto, fieldModal, setFieldModal, addField }) {
  const [addColumnFor, setAddColumnFor] = useState(null);
  const [bulkConfirm, setBulkConfirm] = useState(null);
  const [showBreakdown, setShowBreakdown] = useState(false);
  const displayColumns = useMemo(() => {
    const out = [];
    let i = 0;
    while (i < filteredColumns.length) {
      const topicId = filteredColumns[i].topicId || null;
      const topicName = filteredColumns[i].topicName;
      let j = i;
      while (j < filteredColumns.length && (filteredColumns[j].topicId || null) === topicId) { out.push(filteredColumns[j]); j++; }
      if (topicId) out.push({ kind: 'addColumn', id: `add-${topicId}`, topicId, topicName });
      i = j;
    }
    return out;
  }, [filteredColumns]);

  const [photoThumbs, setPhotoThumbs] = useState({});
  const photoThumbsRef = useRef({});
  useEffect(() => {
    const photoFields = displayColumns.filter((c) => c.kind === 'field' && c.type === 'photo');
    if (!photoFields.length) return;
    let cancelled = false;
    photoFields.forEach((f) => {
      students.forEach((stu) => {
        const key = fieldKey(f.id, stu.id);
        if (!fieldValues[key]) return;
        if (key in photoThumbsRef.current) return;
        photoThumbsRef.current[key] = null;
        loadKey(`photo:${f.id}:${stu.id}`, null).then((url) => {
          if (cancelled) return;
          photoThumbsRef.current[key] = url;
          setPhotoThumbs((prev) => ({ ...prev, [key]: url }));
        });
      });
    });
    return () => { cancelled = true; };
  }, [displayColumns, students, fieldValues, fieldKey]);

  if (structure.classes.length === 0) {
    return <EmptyState text="Add a class in Setup to start assessing." onAction={() => setView('setup')} />;
  }
  const topicOptions = [{ id: '', name: 'All topics' }, ...topicsForSubject.map((t) => ({ id: t.id, name: t.name }))];
  const currentClass = structure.classes.find((c) => c.id === classId);
  const pinnedSubjects = (currentClass?.pinnedSubjectIds || [])
    .map((sid) => structure.subjects.find((s) => s.id === sid))
    .filter(Boolean);
  const subjectDropdownOptions = subjectsForTerm.some((s) => s.id === subjectId)
    ? subjectsForTerm
    : [...subjectsForTerm, ...structure.subjects.filter((s) => s.id === subjectId)];
  return (
    <div>
      <div style={{ display: 'flex', gap: 6, padding: '10px 18px', overflowX: 'auto', borderBottom: `var(--border-width) solid ${LINE}` }}>
        {structure.classes.map((c) => (
          <Pill key={c.id} active={c.id === classId} onClick={() => setClassId(c.id)}>{c.name}</Pill>
        ))}
      </div>
      {pinnedSubjects.length > 0 && (
        <div style={{ display: 'flex', gap: 6, padding: '8px 18px', overflowX: 'auto', borderBottom: `var(--border-width) solid ${LINE}`, background: 'var(--row-stripe)' }}>
          {pinnedSubjects.map((s) => (
            <Pill key={s.id} active={s.id === subjectId} onClick={() => setSubjectId(s.id)}>{s.name}</Pill>
          ))}
        </div>
      )}
      <div style={{ display: 'flex', gap: 16, padding: '10px 18px', alignItems: 'center', flexWrap: 'wrap', borderBottom: `var(--border-width) solid ${LINE}`, background: CARD }}>
        <FilterSelect label="Term" value={termId} onChange={setTermId} options={structure.terms} />
        <FilterSelect label="Subject" value={subjectId} onChange={setSubjectId} options={subjectDropdownOptions} />
        <FilterSelect label="Topic" value={topicFilterId} onChange={setTopicFilterId} options={topicOptions} allowEmpty />
        <div style={{ display: 'flex', background: PAPER, borderRadius: 999, padding: 3, border: `var(--border-width) solid ${LINE}` }}>
          <Pill active={assessMode === 'grid'} onClick={() => setAssessMode('grid')}>Grid</Pill>
          <Pill active={assessMode === 'focus'} onClick={() => setAssessMode('focus')}>Focus</Pill>
          <Pill active={assessMode === 'tap'} onClick={() => setAssessMode('tap')}><span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><Smartphone size={12} /> Tap</span></Pill>
        </div>
        <button
          onClick={() => setShowBreakdown((v) => !v)}
          style={{ display: 'flex', alignItems: 'center', gap: 5, border: `var(--border-width) solid ${showBreakdown ? ACCENT : LINE}`, background: showBreakdown ? ACCENT : CARD, color: showBreakdown ? '#fff' : MUTED, borderRadius: 999, padding: '6px 12px', fontSize: 12.5, cursor: 'pointer', fontWeight: 500 }}
        >
          <BarChart3 size={13} /> Class breakdown
        </button>
        <div style={{ display: 'flex', gap: 12, marginLeft: 'auto', fontSize: 11, color: MUTED, flexWrap: 'wrap' }}>
          {LEVELS.map((l, i) => (
            <span key={l.name} style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
              <LevelPill level={i} width={22} height={11} />
              {l.name}
            </span>
          ))}
          <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><Star size={12} color="var(--indigo)" fill="var(--indigo)" /> High achiever</span>
          <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}><TrendingUp size={12} color="var(--coral)" /> Needs a push</span>
        </div>
      </div>

      {showBreakdown && (
        <ClassBreakdownPanel columns={filteredCriteriaColumns} students={students} termId={termId} records={records} recKey={recKey} />
      )}

      {structure.subjects.length === 0 ? (
        <EmptyState text="Add a subject and success criteria in Setup." onAction={() => setView('setup')} />
      ) : orderedColumns.length === 0 ? (
        <EmptyState text="No success criteria for this subject yet. Add some in Setup." onAction={() => setView('setup')} />
      ) : students.length === 0 ? (
        <EmptyState text="No students in this class yet. Add them in Setup." onAction={() => setView('setup')} />
      ) : assessMode === 'focus' ? (
        filteredCriteriaColumns.length === 0 ? (
          <EmptyState text="No success criteria in this topic to focus on — try Grid, or add criteria in Setup." onAction={() => setView('setup')} />
        ) : (
          <FocusPanel
            columns={filteredCriteriaColumns}
            students={students}
            termId={termId}
            records={records}
            recKey={recKey}
            cycleLevel={cycleLevel}
            setLevel={setLevel}
            bulkSetLevel={bulkSetLevel}
            focusIndex={focusIndex}
            setFocusIndex={setFocusIndex}
          />
        )
      ) : assessMode === 'tap' ? (
        <TapAssessPanel
          columns={filteredColumns}
          students={students}
          termId={termId}
          records={records}
          recKey={recKey}
          cycleLevel={cycleLevel}
          setLevel={setLevel}
          bulkSetLevelForStudent={bulkSetLevelForStudent}
          coverageFor={coverageFor}
          studentIndex={tapStudentIndex}
          setStudentIndex={setTapStudentIndex}
          fieldValues={fieldValues}
          fieldKey={fieldKey}
          setFieldValue={setFieldValue}
          handlePhotoFile={handlePhotoFile}
          setFieldModal={setFieldModal}
        />
      ) : (
        <div style={{ overflowX: 'auto' }}>
          <table style={{ borderCollapse: 'collapse', width: '100%', fontSize: 12 }}>
            <thead>
              <tr>
                <th rowSpan={2} style={thStyle(true)}>Student</th>
                <th rowSpan={2} style={{ ...thStyle(true), width: 70 }}>Coverage</th>
                {groupConsecutive(displayColumns.map((c) => c.topicName)).map((g, i) => (
                  <th key={i} colSpan={g.span} style={{ ...thStyle(), borderBottom: `var(--border-width) solid ${LINE}`, background: CARD }}>{g.value}</th>
                ))}
              </tr>
              <tr>
                {displayColumns.map((c) => (
                  c.kind === 'addColumn' ? (
                    <th key={c.id} style={{ ...thStyle(), minWidth: 26, maxWidth: 26, padding: '6px 1px' }}>
                      <button
                        onClick={() => setAddColumnFor({ topicId: c.topicId, topicName: c.topicName })}
                        title={`Add a column to ${c.topicName}`}
                        style={{ border: `1px dashed ${LINE}`, background: CARD, borderRadius: 5, width: 20, height: 20, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer', color: MUTED, margin: '0 auto' }}
                      >
                        <Plus size={12} />
                      </button>
                    </th>
                  ) : (
                    <th key={c.id} style={{ ...thStyle(), minWidth: c.kind === 'field' ? fieldColWidth(c.type) : 96, maxWidth: c.kind === 'field' ? fieldColWidth(c.type) : 130, padding: c.kind === 'field' && c.type === 'note' ? '6px 8px' : '8px 6px', background: c.kind === 'field' ? 'var(--row-stripe)' : undefined, verticalAlign: 'bottom' }}>
                      {c.kind === 'field' ? (
                        c.type === 'note' ? (
                          <div style={{ fontSize: 11, color: MUTED, display: 'flex', alignItems: 'center', gap: 5, justifyContent: 'flex-start', textAlign: 'left' }}>
                            <FieldTypeIcon type={c.type} size={12} />
                            {c.name}
                          </div>
                        ) : (
                          <div style={{ fontSize: 10.5, color: MUTED, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 3 }}>
                            <FieldTypeIcon type={c.type} size={12} />
                            <span style={{ whiteSpace: 'normal', wordBreak: 'normal', overflowWrap: 'break-word', lineHeight: 1.25, textAlign: 'center' }}>{c.name}</span>
                          </div>
                        )
                      ) : (
                        <div>
                          <div title={c.name} style={{ whiteSpace: 'normal', wordBreak: 'normal', overflowWrap: 'break-word', overflow: 'hidden', display: '-webkit-box', WebkitLineClamp: 4, WebkitBoxOrient: 'vertical', fontSize: 11, lineHeight: 1.3, textAlign: 'center', fontWeight: 400, color: MUTED }}>
                            {c.name}
                          </div>
                          {students.length > 0 && (
                            <div style={{ display: 'flex', gap: 4, justifyContent: 'center', marginTop: 7 }}>
                              {LEVELS.map((lvl, i) => (
                                <button
                                  key={i}
                                  onClick={() => setBulkConfirm({ critId: c.id, level: i, critName: c.name, levelName: lvl.name })}
                                  title={`Set everyone to ${lvl.name}`}
                                  style={{
                                    width: 15, height: 15, borderRadius: 4, cursor: 'pointer', padding: 0,
                                    border: i === 0 ? `1.5px dashed var(--ink-muted)` : 'none',
                                    background: i === 0 ? 'transparent' : lvl.fill,
                                  }}
                                />
                              ))}
                            </div>
                          )}
                        </div>
                      )}
                    </th>
                  )
                ))}
              </tr>
            </thead>
            <tbody>
              {students.map((stu, si) => {
                const cov = coverageFor(stu.id);
                return (
                  <tr key={stu.id} style={{ background: si % 2 ? 'var(--row-stripe)' : 'transparent' }}>
                    <td style={{ ...tdStyle, fontWeight: 500, position: 'sticky', left: 0, background: si % 2 ? 'var(--row-stripe)' : CARD, textAlign: 'left', padding: '6px 12px' }}>
                      <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        {stu.name}
                        <StudentFlagBadge flag={computeStudentFlag(stu.id, filteredCriteriaColumns, records, termId, recKey)} />
                      </span>
                    </td>
                    <td style={{ ...tdStyle, color: MUTED, fontVariantNumeric: 'tabular-nums' }}>{cov}%</td>
                    {displayColumns.map((c) => {
                      if (c.kind === 'addColumn') {
                        return <td key={c.id} style={{ ...tdStyle, minWidth: 26, background: 'var(--row-stripe)' }} />;
                      }
                      if (c.kind === 'field') {
                        return (
                          <td key={c.id} style={{ ...tdStyle, minWidth: fieldColWidth(c.type), maxWidth: fieldColWidth(c.type), width: fieldColWidth(c.type), overflow: 'hidden', padding: c.type === 'note' ? '5px 6px' : tdStyle.padding, background: 'var(--row-stripe)' }}>
                            <FieldCell
                              field={c}
                              value={fieldValues[fieldKey(c.id, stu.id)]}
                              thumbUrl={photoThumbs[fieldKey(c.id, stu.id)]}
                              setValue={(v) => setFieldValue(c.id, stu.id, v)}
                              onPhotoFile={(file) => handlePhotoFile(c.id, stu.id, file)}
                              onOpenPhoto={() => setFieldModal({ fieldId: c.id, studentId: stu.id, type: 'photo', fieldName: c.name, studentName: stu.name })}
                              onOpenNote={() => setFieldModal({ fieldId: c.id, studentId: stu.id, type: 'note', fieldName: c.name, studentName: stu.name })}
                            />
                          </td>
                        );
                      }
                      const lvl = records[recKey(stu.id, termId, c.id)] ?? 0;
                      return (
                        <td
                          key={c.id}
                          title={LEVELS[lvl].name}
                          style={{ ...tdStyle, background: LEVELS[lvl].tint, minWidth: 96, maxWidth: 130, padding: '7px 10px' }}
                        >
                          <LevelSegmentedPill level={lvl} onSet={(v) => setLevel(stu.id, c.id, v)} height={24} />
                        </td>
                      );
                    })}
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr style={{ position: 'sticky', bottom: 0, background: CARD, zIndex: 2 }}>
                <td style={{ ...tdStyle, position: 'sticky', left: 0, background: CARD, textAlign: 'left', padding: '8px 12px', fontWeight: 600, fontSize: 11.5, color: MUTED, borderTop: `2px solid ${LINE}` }}>Class breakdown</td>
                <td style={{ ...tdStyle, borderTop: `2px solid ${LINE}`, color: MUTED, fontSize: 11 }}>
                  {students.length ? `${Math.round(students.reduce((sum, s) => sum + (coverageFor(s.id) ?? 0), 0) / students.length)}% avg` : '—'}
                </td>
                {displayColumns.map((c) => {
                  if (c.kind !== 'criteria') {
                    return <td key={c.id} style={{ ...tdStyle, borderTop: `2px solid ${LINE}`, minWidth: c.kind === 'addColumn' ? 26 : fieldColWidth(c.type) }} />;
                  }
                  const counts = [0, 0, 0, 0];
                  students.forEach((s) => { counts[records[recKey(s.id, termId, c.id)] ?? 0]++; });
                  const met = counts[2] + counts[3];
                  const pct = students.length ? Math.round((met / students.length) * 100) : 0;
                  return (
                    <td key={c.id} style={{ ...tdStyle, borderTop: `2px solid ${LINE}`, minWidth: 96, maxWidth: 130, padding: '7px 10px' }} title={`Not started ${counts[0]} · Developing ${counts[1]} · Secure ${counts[2]} · Mastered ${counts[3]}`}>
                      <div style={{ display: 'flex', height: 7, borderRadius: 4, overflow: 'hidden', border: `var(--border-width) solid ${LINE}`, marginBottom: 4 }}>
                        {[0, 1, 2, 3].map((lvl) => counts[lvl] > 0 && (
                          <div key={lvl} style={{ flex: counts[lvl], background: lvl === 0 ? 'var(--line)' : LEVELS[lvl].fill }} />
                        ))}
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'center', gap: 5, fontSize: 9.5, fontWeight: 600, marginBottom: 2 }}>
                        <span style={{ color: 'var(--ink-muted)' }}>{counts[0]}</span>
                        <span style={{ color: 'var(--coral)' }}>{counts[1]}</span>
                        <span style={{ color: 'var(--green)' }}>{counts[2]}</span>
                        <span style={{ color: 'var(--indigo)' }}>{counts[3]}</span>
                      </div>
                      <div style={{ fontSize: 9, color: MUTED }}>{pct}% met</div>
                    </td>
                  );
                })}
              </tr>
            </tfoot>
          </table>
        </div>
      )}
      {addColumnFor && (
        <AddColumnModal
          topicName={addColumnFor.topicName}
          onClose={() => setAddColumnFor(null)}
          onAdd={(name, type) => { addField(addColumnFor.topicId, name, type); setAddColumnFor(null); }}
        />
      )}
      {bulkConfirm && (
        <ConfirmModal
          title="Set everyone at once?"
          message={`Set all ${students.length} students to "${bulkConfirm.levelName}" for "${bulkConfirm.critName}"?\n\nThis overwrites any existing marks for this criteria — tap individual students afterwards for exceptions (absent, still developing, already mastered).`}
          confirmLabel={`Set all to ${bulkConfirm.levelName}`}
          onConfirm={() => { bulkSetLevel(bulkConfirm.critId, bulkConfirm.level); setBulkConfirm(null); }}
          onCancel={() => setBulkConfirm(null)}
        />
      )}
      {fieldModal && (
        <FieldModal
          modal={fieldModal}
          onClose={() => setFieldModal(null)}
          value={fieldValues[fieldKey(fieldModal.fieldId, fieldModal.studentId)]}
          setValue={(v) => setFieldValue(fieldModal.fieldId, fieldModal.studentId, v)}
          onRemovePhoto={() => removePhoto(fieldModal.fieldId, fieldModal.studentId)}
          onPhotoFile={(file) => handlePhotoFile(fieldModal.fieldId, fieldModal.studentId, file)}
        />
      )}
    </div>
  );
}

function fieldColWidth(type) {
  if (type === 'rating') return 96;
  if (type === 'date') return 100;
  if (type === 'note') return 150;
  return 78;
}

function FieldTypeIcon({ type, size = 13, color }) {
  const c = color || MUTED;
  if (type === 'rating') return <Star size={size} color={c} />;
  if (type === 'date') return <CalendarIcon size={size} color={c} />;
  if (type === 'photo') return <Camera size={size} color={c} />;
  return <MessageSquare size={size} color={c} />;
}

function FieldCell({ field, value, thumbUrl, setValue, onPhotoFile, onOpenPhoto, onOpenNote }) {
  if (field.type === 'rating') {
    const rating = value || 0;
    return (
      <div style={{ display: 'flex', gap: 1, justifyContent: 'center' }}>
        {[1, 2, 3, 4, 5].map((n) => (
          <span key={n} onClick={() => setValue(rating === n ? 0 : n)} style={{ cursor: 'pointer', color: n <= rating ? AMBER : 'var(--ink-muted)' }}>
            <Star size={13} fill={n <= rating ? AMBER : 'none'} />
          </span>
        ))}
      </div>
    );
  }
  if (field.type === 'date') {
    return (
      <input
        type="date"
        value={value || ''}
        onChange={(e) => setValue(e.target.value)}
        style={{ fontSize: 11, border: `var(--border-width) solid ${LINE}`, borderRadius: 5, padding: '3px 4px', width: '100%', background: CARD, color: INK }}
      />
    );
  }
  if (field.type === 'photo') {
    if (!value) {
      return <div style={{ display: 'flex', justifyContent: 'center' }}><PhotoPicker onFile={onPhotoFile} /></div>;
    }
    return (
      <button onClick={onOpenPhoto} style={{ border: 'none', background: 'transparent', cursor: 'pointer', display: 'flex', justifyContent: 'center', width: '100%', padding: 0 }} title="View photo">
        {thumbUrl ? (
          <img src={thumbUrl} alt="" style={{ width: 30, height: 30, borderRadius: 6, objectFit: 'cover', border: `var(--border-width) solid ${LINE}`, display: 'block' }} />
        ) : (
          <span style={{ width: 26, height: 26, borderRadius: 6, background: 'var(--green-bg)', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
            <Camera size={13} color="var(--green)" />
          </span>
        )}
      </button>
    );
  }
  const hasNote = value && String(value).trim();
  return (
    <button
      onClick={onOpenNote}
      title={hasNote ? value : 'Add note'}
      style={{
        border: `var(--border-width) solid ${hasNote ? 'var(--indigo)' : LINE}`,
        background: hasNote ? 'var(--indigo-bg)' : CARD,
        borderRadius: 7,
        cursor: 'pointer',
        width: '100%',
        minHeight: 30,
        maxHeight: 54,
        padding: '5px 7px',
        textAlign: 'left',
        fontSize: 11.5,
        lineHeight: 1.35,
        color: hasNote ? INK : MUTED,
        display: '-webkit-box',
        WebkitLineClamp: 3,
        WebkitBoxOrient: 'vertical',
        overflow: 'hidden',
        wordBreak: 'break-word',
        whiteSpace: 'normal',
      }}
    >
      {hasNote ? value : 'Add note'}
    </button>
  );
}

function PhotoPicker({ onFile, big }) {
  const ref = useRef(null);
  const size = big ? 44 : 22;
  return (
    <>
      <span
        onClick={(e) => { e.stopPropagation(); ref.current?.click(); }}
        style={{ width: size, height: size, borderRadius: big ? 10 : 5, background: 'var(--hover-bg)', display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' }}
      >
        <Camera size={big ? 18 : 13} color={MUTED} />
      </span>
      <input
        ref={ref}
        type="file"
        accept="image/*"
        capture="environment"
        style={{ display: 'none' }}
        onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ''; }}
      />
    </>
  );
}

function ConfirmModal({ title, message, confirmLabel = 'Confirm', danger, onConfirm, onCancel }) {
  return (
    <div onClick={onCancel} style={{ position: 'fixed', inset: 0, background: '#00000055', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 60, padding: 20 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: CARD, borderRadius: 12, padding: 18, width: '100%', maxWidth: 380, display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div style={{ fontSize: 14, fontWeight: 600, color: INK }}>{title}</div>
        <div style={{ fontSize: 13, color: MUTED, whiteSpace: 'pre-line', lineHeight: 1.45 }}>{message}</div>
        <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
          <button onClick={onCancel} style={{ border: `var(--border-width) solid ${LINE}`, background: CARD, color: INK, borderRadius: 7, padding: '8px 14px', fontSize: 13, cursor: 'pointer' }}>Cancel</button>
          <button onClick={onConfirm} style={{ border: 'none', background: danger ? '#9A4A3A' : ACCENT, color: '#fff', borderRadius: 7, padding: '8px 14px', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

function AddColumnModal({ topicName, onClose, onAdd }) {
  const [name, setName] = useState('');
  const [type, setType] = useState('note');
  const submit = () => {
    const v = name.trim();
    if (!v) return;
    onAdd(v, type);
  };
  return (
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: '#00000055', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 50, padding: 20 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: CARD, borderRadius: 12, padding: 18, width: '100%', maxWidth: 340, display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div style={{ fontSize: 14, fontWeight: 500 }}>Add column to {topicName}</div>
          <IconBtn onClick={onClose} title="Close"><X size={16} /></IconBtn>
        </div>
        <input
          autoFocus
          value={name}
          onChange={(e) => setName(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && submit()}
          placeholder="Column name, e.g. Model level"
          style={{ fontSize: 13, padding: '8px 10px', borderRadius: 7, border: `var(--border-width) solid ${LINE}`, outline: 'none' }}
        />
        <div style={{ display: 'flex', gap: 6 }}>
          {Object.entries(FIELD_TYPES).map(([key, v]) => (
            <button
              key={key}
              onClick={() => setType(key)}
              style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4, padding: '8px 4px', borderRadius: 8, border: `var(--border-width) solid ${type === key ? INK : LINE}`, background: type === key ? INK : CARD, color: type === key ? '#fff' : INK, cursor: 'pointer', fontSize: 11 }}
            >
              <FieldTypeIcon type={key} size={14} color={type === key ? '#fff' : MUTED} />
              {v.label}
            </button>
          ))}
        </div>
        <button onClick={submit} style={{ border: 'none', background: ACCENT, color: '#fff', borderRadius: 7, padding: '9px 14px', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}>Add column</button>
      </div>
    </div>
  );
}

function FieldModal({ modal, onClose, value, setValue, onRemovePhoto, onPhotoFile }) {
  const [photoUrl, setPhotoUrl] = useState(null);
  const [loadingPhoto, setLoadingPhoto] = useState(false);
  useEffect(() => {
    if (modal.type === 'photo' && value) {
      setLoadingPhoto(true);
      loadKey(`photo:${modal.fieldId}:${modal.studentId}`, null).then((d) => { setPhotoUrl(d); setLoadingPhoto(false); });
    }
  }, [modal.fieldId, modal.studentId, modal.type, value]);

  return (
    <div onClick={onClose} style={{ position: 'fixed', inset: 0, background: '#00000055', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 50, padding: 20 }}>
      <div onClick={(e) => e.stopPropagation()} style={{ background: CARD, borderRadius: 12, padding: 18, width: '100%', maxWidth: 380, display: 'flex', flexDirection: 'column', gap: 12 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <div>
            <div style={{ fontSize: 14, fontWeight: 500 }}>{modal.studentName}</div>
            <div style={{ fontSize: 12, color: MUTED }}>{modal.fieldName}</div>
          </div>
          <IconBtn onClick={onClose} title="Close"><X size={16} /></IconBtn>
        </div>

        {modal.type === 'note' && (
          <>
            <textarea
              autoFocus
              value={value || ''}
              onChange={(e) => setValue(e.target.value)}
              rows={4}
              style={{ width: '100%', fontSize: 13, padding: 10, borderRadius: 8, border: `var(--border-width) solid ${LINE}`, resize: 'vertical', fontFamily: 'inherit' }}
              placeholder="Type a note…"
            />
            <button onClick={onClose} style={{ alignSelf: 'flex-end', border: 'none', background: ACCENT, color: '#fff', borderRadius: 7, padding: '7px 14px', fontSize: 13, cursor: 'pointer' }}>Done</button>
          </>
        )}

        {modal.type === 'photo' && (
          <>
            {loadingPhoto ? (
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, color: MUTED, fontSize: 13, padding: 20, justifyContent: 'center' }}><Loader2 size={16} className="spin" /> Loading photo…</div>
            ) : photoUrl ? (
              <img src={photoUrl} alt="evidence" style={{ width: '100%', borderRadius: 8, display: 'block' }} />
            ) : (
              <div style={{ fontSize: 12, color: MUTED, padding: 20, textAlign: 'center' }}>No photo yet.</div>
            )}
            <div style={{ display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
              {value && <button onClick={onRemovePhoto} style={{ border: `var(--border-width) solid ${LINE}`, background: CARD, color: '#9A4A3A', borderRadius: 7, padding: '7px 14px', fontSize: 13, cursor: 'pointer' }}>Remove</button>}
              <label style={{ border: 'none', background: ACCENT, color: '#fff', borderRadius: 7, padding: '7px 14px', fontSize: 13, cursor: 'pointer' }}>
                {value ? 'Retake' : 'Take photo'}
                <input type="file" accept="image/*" capture="environment" style={{ display: 'none' }} onChange={(e) => { const f = e.target.files?.[0]; if (f) { onPhotoFile(f); setPhotoUrl(null); onClose(); } }} />
              </label>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

function FocusPanel({ columns, students, termId, records, recKey, cycleLevel, setLevel, bulkSetLevel, focusIndex, setFocusIndex }) {
  const [bulkConfirm, setBulkConfirm] = useState(null);
  if (columns.length === 0) return null;
  const col = columns[Math.min(focusIndex, columns.length - 1)];
  const go = (delta) => setFocusIndex((i) => Math.max(0, Math.min(columns.length - 1, i + delta)));
  return (
    <div style={{ padding: '14px 18px' }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 10 }}>
        <button
          onClick={() => go(-1)}
          disabled={focusIndex === 0}
          style={{ width: 40, height: 40, borderRadius: 8, border: `var(--border-width) solid ${LINE}`, background: CARD, fontSize: 16, cursor: focusIndex === 0 ? 'default' : 'pointer', color: focusIndex === 0 ? 'var(--ink-muted)' : INK }}
        >‹</button>
        <div style={{ flex: 1, textAlign: 'center' }}>
          <div style={{ fontSize: 11, color: MUTED }}>{col.topicName} · {focusIndex + 1} of {columns.length}</div>
          <div style={{ fontSize: 16, fontWeight: 500 }}>{col.name}</div>
        </div>
        <button
          onClick={() => go(1)}
          disabled={focusIndex === columns.length - 1}
          style={{ width: 40, height: 40, borderRadius: 8, border: `var(--border-width) solid ${LINE}`, background: CARD, fontSize: 16, cursor: focusIndex === columns.length - 1 ? 'default' : 'pointer', color: focusIndex === columns.length - 1 ? 'var(--ink-muted)' : INK }}
        >›</button>
      </div>
      {students.length > 0 && (
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, marginBottom: 14 }}>
          <span style={{ fontSize: 11, color: MUTED }}>Set everyone:</span>
          {LEVELS.map((lvl, i) => (
            <button
              key={i}
              onClick={() => setBulkConfirm({ level: i, levelName: lvl.name })}
              title={`Set all ${students.length} students to ${lvl.name}`}
              style={{
                width: 18, height: 18, borderRadius: 5, cursor: 'pointer', padding: 0,
                border: i === 0 ? `1.5px dashed var(--ink-muted)` : 'none',
                background: i === 0 ? 'transparent' : lvl.fill,
              }}
            />
          ))}
        </div>
      )}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {students.map((stu) => {
          const lvl = records[recKey(stu.id, termId, col.id)] ?? 0;
          return (
            <div
              key={stu.id}
              style={{
                display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 14,
                padding: '12px 16px', borderRadius: 10, border: `var(--border-width) solid ${LINE}`,
                background: LEVELS[lvl].tint, minHeight: 56, textAlign: 'left',
              }}
            >
              <span style={{ fontSize: 15, fontWeight: 500, color: INK, display: 'flex', alignItems: 'center', gap: 7 }}>
                {stu.name}
                <StudentFlagBadge flag={computeStudentFlag(stu.id, columns, records, termId, recKey)} />
              </span>
              <div style={{ width: 150, flexShrink: 0 }}>
                <LevelSegmentedPill level={lvl} onSet={(v) => setLevel(stu.id, col.id, v)} height={30} />
              </div>
            </div>
          );
        })}
      </div>
      {bulkConfirm && (
        <ConfirmModal
          title="Set everyone at once?"
          message={`Set all ${students.length} students to "${bulkConfirm.levelName}" for "${col.name}"?\n\nThis overwrites any existing marks for this criteria — tap individual students afterwards for exceptions (absent, still developing, already mastered).`}
          confirmLabel={`Set all to ${bulkConfirm.levelName}`}
          onConfirm={() => { bulkSetLevel(col.id, bulkConfirm.level); setBulkConfirm(null); }}
          onCancel={() => setBulkConfirm(null)}
        />
      )}
    </div>
  );
}

function TapAssessPanel({ columns, students, termId, records, recKey, cycleLevel, setLevel, bulkSetLevelForStudent, coverageFor, studentIndex, setStudentIndex, fieldValues, fieldKey, setFieldValue, handlePhotoFile, setFieldModal }) {
  const touchStartX = useRef(null);
  const pillRefs = useRef({});
  const idx = Math.min(studentIndex, Math.max(0, students.length - 1));
  const [bulkConfirm, setBulkConfirm] = useState(null);

  useEffect(() => {
    const el = pillRefs.current[students[idx]?.id];
    if (el) el.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
  }, [idx, students]);

  if (students.length === 0) return null;
  const stu = students[idx];
  const cov = coverageFor(stu.id);
  const go = (delta) => setStudentIndex((i) => Math.max(0, Math.min(students.length - 1, i + delta)));
  const criteriaOnly = columns.filter((c) => c.kind === 'criteria');
  const flag = computeStudentFlag(stu.id, criteriaOnly, records, termId, recKey);

  const groups = [];
  columns.forEach((c) => {
    const last = groups[groups.length - 1];
    if (last && last.topicName === c.topicName) last.items.push(c);
    else groups.push({ topicName: c.topicName, items: [c] });
  });

  const onTouchStart = (e) => { touchStartX.current = e.touches[0].clientX; };
  const onTouchEnd = (e) => {
    if (touchStartX.current == null) return;
    const dx = e.changedTouches[0].clientX - touchStartX.current;
    if (Math.abs(dx) > 60) go(dx > 0 ? -1 : 1);
    touchStartX.current = null;
  };

  return (
    <div style={{ padding: '12px 14px 24px' }} onTouchStart={onTouchStart} onTouchEnd={onTouchEnd}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 14, position: 'sticky', top: 0, background: PAPER, zIndex: 3, padding: '4px 0' }}>
        <button
          onClick={() => go(-1)}
          disabled={idx === 0}
          style={{ width: 48, height: 48, flexShrink: 0, borderRadius: 12, border: `var(--border-width) solid ${LINE}`, background: CARD, fontSize: 20, cursor: idx === 0 ? 'default' : 'pointer', color: idx === 0 ? 'var(--ink-muted)' : INK }}
        >‹</button>
        <div style={{ flex: 1, textAlign: 'center', minWidth: 0 }}>
          <div style={{ fontSize: 11.5, color: MUTED }}>{idx + 1} of {students.length} · {cov}% coverage</div>
          <div style={{ fontSize: 19, fontWeight: 600, color: INK, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 7 }}>
            {stu.name}
            <StudentFlagBadge flag={flag} size={16} />
          </div>
        </div>
        <button
          onClick={() => go(1)}
          disabled={idx === students.length - 1}
          style={{ width: 48, height: 48, flexShrink: 0, borderRadius: 12, border: `var(--border-width) solid ${LINE}`, background: CARD, fontSize: 20, cursor: idx === students.length - 1 ? 'default' : 'pointer', color: idx === students.length - 1 ? 'var(--ink-muted)' : INK }}
        >›</button>
      </div>

      <div style={{ display: 'flex', gap: 6, overflowX: 'auto', paddingBottom: 8, marginBottom: 12, position: 'sticky', top: 60, background: PAPER, zIndex: 2, WebkitOverflowScrolling: 'touch' }}>
        {students.map((s, i) => (
          <button
            key={s.id}
            ref={(el) => { pillRefs.current[s.id] = el; }}
            onClick={() => setStudentIndex(i)}
            style={{
              flexShrink: 0, padding: '6px 12px', borderRadius: 999, fontSize: 12.5, fontWeight: 500, cursor: 'pointer',
              border: `var(--border-width) solid ${i === idx ? ACCENT : LINE}`,
              background: i === idx ? ACCENT : CARD,
              color: i === idx ? '#fff' : INK,
              whiteSpace: 'nowrap',
            }}
          >
            {s.name.split(' ')[0]}
          </button>
        ))}
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 20 }}>
        {groups.map((g) => {
          const groupCriteriaIds = g.items.filter((c) => c.kind === 'criteria').map((c) => c.id);
          return (
          <div key={g.topicName}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8, flexWrap: 'wrap', gap: 8 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: MUTED, textTransform: 'uppercase', letterSpacing: 0.4 }}>{g.topicName}</div>
              {groupCriteriaIds.length > 0 && (
                <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                  <span style={{ fontSize: 10.5, color: MUTED }}>Set all here:</span>
                  {LEVELS.map((lvl, i) => (
                    <button
                      key={i}
                      onClick={() => setBulkConfirm({ level: i, levelName: lvl.name, topicName: g.topicName, criteriaIds: groupCriteriaIds })}
                      title={`Set all ${g.topicName} criteria to ${lvl.name} for ${stu.name}`}
                      style={{
                        width: 16, height: 16, borderRadius: 4, cursor: 'pointer', padding: 0,
                        border: i === 0 ? `1.5px dashed var(--ink-muted)` : 'none',
                        background: i === 0 ? 'transparent' : lvl.fill,
                      }}
                    />
                  ))}
                </div>
              )}
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {g.items.map((c) => {
                if (c.kind === 'field') {
                  return (
                    <div key={c.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10, padding: '12px 14px', borderRadius: 12, border: `var(--border-width) solid ${LINE}`, background: CARD, minHeight: 56 }}>
                      <span style={{ fontSize: 14.5, fontWeight: 500, color: INK, display: 'flex', alignItems: 'center', gap: 8 }}>
                        <FieldTypeIcon type={c.type} size={14} />
                        {c.name}
                      </span>
                      <div style={{ flexShrink: 0 }}>
                        {c.type === 'rating' ? (
                          <div style={{ display: 'flex', gap: 3 }}>
                            {[1, 2, 3, 4, 5].map((n) => {
                              const rating = fieldValues[fieldKey(c.id, stu.id)] || 0;
                              return (
                                <span key={n} onClick={() => setFieldValue(c.id, stu.id, rating === n ? 0 : n)} style={{ cursor: 'pointer', color: n <= rating ? AMBER : 'var(--ink-muted)' }}>
                                  <Star size={22} fill={n <= rating ? AMBER : 'none'} />
                                </span>
                              );
                            })}
                          </div>
                        ) : c.type === 'date' ? (
                          <input
                            type="date"
                            value={fieldValues[fieldKey(c.id, stu.id)] || ''}
                            onChange={(e) => setFieldValue(c.id, stu.id, e.target.value)}
                            style={{ fontSize: 14, border: `var(--border-width) solid ${LINE}`, borderRadius: 8, padding: '8px 10px', background: CARD, color: INK }}
                          />
                        ) : c.type === 'photo' ? (
                          fieldValues[fieldKey(c.id, stu.id)] ? (
                            <button onClick={() => setFieldModal({ fieldId: c.id, studentId: stu.id, type: 'photo', fieldName: c.name, studentName: stu.name })} style={{ border: 'none', background: 'var(--green-bg)', borderRadius: 10, width: 44, height: 44, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' }}>
                              <Camera size={18} color="var(--green)" />
                            </button>
                          ) : (
                            <PhotoPicker onFile={(file) => handlePhotoFile(c.id, stu.id, file)} big />
                          )
                        ) : (
                          <button
                            onClick={() => setFieldModal({ fieldId: c.id, studentId: stu.id, type: 'note', fieldName: c.name, studentName: stu.name })}
                            style={{ border: `var(--border-width) solid ${LINE}`, background: PAPER, borderRadius: 8, padding: '8px 12px', fontSize: 13, cursor: 'pointer', color: fieldValues[fieldKey(c.id, stu.id)] ? INK : MUTED, maxWidth: 160, textAlign: 'left', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}
                          >
                            {fieldValues[fieldKey(c.id, stu.id)] || 'Add note'}
                          </button>
                        )}
                      </div>
                    </div>
                  );
                }
                const lvl = records[recKey(stu.id, termId, c.id)] ?? 0;
                return (
                  <div
                    key={c.id}
                    style={{
                      display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 10,
                      padding: '14px 16px', borderRadius: 12, border: `var(--border-width) solid ${LINE}`,
                      background: LEVELS[lvl].tint, minHeight: 58, textAlign: 'left', width: '100%',
                    }}
                  >
                    <span style={{ fontSize: 14.5, fontWeight: 500, color: INK }}>{c.name}</span>
                    <div style={{ width: 160, flexShrink: 0 }}>
                      <LevelSegmentedPill level={lvl} onSet={(v) => setLevel(stu.id, c.id, v)} height={34} />
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
          );
        })}
      </div>
      {bulkConfirm && (
        <ConfirmModal
          title="Set this topic at once?"
          message={`Set all "${bulkConfirm.topicName}" criteria to "${bulkConfirm.levelName}" for ${stu.name}?\n\nThis overwrites any existing marks for these criteria — adjust individually afterwards if needed.`}
          confirmLabel={`Set all to ${bulkConfirm.levelName}`}
          onConfirm={() => { bulkSetLevelForStudent(stu.id, bulkConfirm.criteriaIds, bulkConfirm.level); setBulkConfirm(null); }}
          onCancel={() => setBulkConfirm(null)}
        />
      )}
    </div>
  );
}

function ClassBreakdownPanel({ columns, students, termId, records, recKey }) {
  const total = students.length;
  if (columns.length === 0 || total === 0) {
    return (
      <div style={{ padding: '14px 18px', background: CARD, borderBottom: `var(--border-width) solid ${LINE}`, fontSize: 12.5, color: MUTED }}>
        Nothing to break down yet for this view.
      </div>
    );
  }
  return (
    <div style={{ padding: '14px 18px', background: CARD, borderBottom: `var(--border-width) solid ${LINE}` }}>
      <div style={{ fontSize: 11.5, fontWeight: 600, color: MUTED, marginBottom: 12, textTransform: 'uppercase', letterSpacing: 0.4 }}>
        Class breakdown · {total} student{total === 1 ? '' : 's'}
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
        {columns.map((c) => {
          const counts = [0, 0, 0, 0];
          students.forEach((s) => { counts[records[recKey(s.id, termId, c.id)] ?? 0]++; });
          const met = counts[2] + counts[3];
          const pct = Math.round((met / total) * 100);
          return (
            <div key={c.id}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: 5, gap: 10 }}>
                <span style={{ fontSize: 12.5, color: INK, fontWeight: 500 }}>{c.name}</span>
                <span style={{ fontSize: 11.5, color: MUTED, flexShrink: 0 }}>{pct}% met ({met}/{total})</span>
              </div>
              <div style={{ display: 'flex', height: 12, borderRadius: 6, overflow: 'hidden', border: `var(--border-width) solid ${LINE}` }}>
                {[0, 1, 2, 3].map((lvl) => counts[lvl] > 0 && (
                  <div
                    key={lvl}
                    title={`${LEVELS[lvl].name}: ${counts[lvl]}`}
                    style={{ flex: counts[lvl], background: lvl === 0 ? 'var(--line)' : LEVELS[lvl].fill }}
                  />
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function groupConsecutive(arr) {
  const out = [];
  arr.forEach((v) => {
    if (out.length && out[out.length - 1].value === v) out[out.length - 1].span++;
    else out.push({ value: v, span: 1 });
  });
  return out;
}

const thStyle = (sticky) => ({
  padding: '6px 8px',
  fontWeight: 500,
  color: MUTED,
  textAlign: 'center',
  borderBottom: `var(--border-width) solid ${LINE}`,
  background: sticky ? CARD : undefined,
  position: sticky ? 'sticky' : undefined,
  left: sticky ? 0 : undefined,
  zIndex: sticky ? 2 : 1,
  fontSize: 11.5,
});
const tdStyle = { padding: '5px 6px', textAlign: 'center', borderBottom: `var(--border-width) solid ${LINE}` };

function FilterSelect({ label, value, onChange, options }) {
  return (
    <label style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: MUTED }}>
      {label}
      <select
        value={value || ''}
        onChange={(e) => onChange(e.target.value)}
        style={{ fontSize: 13, padding: '5px 8px', borderRadius: 7, border: `var(--border-width) solid ${LINE}`, background: CARD, color: INK }}
      >
        {options.length === 0 && <option value="">None</option>}
        {options.map((o) => <option key={o.id} value={o.id}>{o.name}</option>)}
      </select>
    </label>
  );
}

function EmptyState({ text, onAction }) {
  return (
    <div style={{ padding: '50px 20px', textAlign: 'center', color: MUTED, fontSize: 13 }}>
      <div style={{ marginBottom: 10 }}>{text}</div>
      <button onClick={onAction} style={{ border: `var(--border-width) solid ${LINE}`, background: CARD, borderRadius: 8, padding: '7px 14px', fontSize: 13, cursor: 'pointer', color: INK }}>
        Go to Setup
      </button>
    </div>
  );
}

function Section({ title, children }) {
  return (
    <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 10, padding: 14, display: 'flex', flexDirection: 'column', gap: 10 }}>
      <div style={{ fontSize: 13, fontWeight: 500 }}>{title}</div>
      {children}
    </div>
  );
}

function ListRow({ label, onRemove }) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '5px 8px', borderRadius: 6, fontSize: 13 }}
      onMouseEnter={(e) => (e.currentTarget.style.background = 'var(--hover-bg)')}
      onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
    >
      <span>{label}</span>
      <IconBtn onClick={onRemove} danger title="Remove"><Trash2 size={13} /></IconBtn>
    </div>
  );
}

function DataListInput({ label, value, onChange, options, placeholder, autoFocus }) {
  const id = useRef(`dl-${Math.random().toString(36).slice(2, 9)}`).current;
  return (
    <label style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 12, color: MUTED, flex: '1 1 160px', minWidth: 140 }}>
      {label}
      <input
        autoFocus={autoFocus}
        list={id}
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        style={{ fontSize: 13, padding: '8px 10px', borderRadius: 7, border: `var(--border-width) solid ${LINE}`, color: INK, background: CARD }}
      />
      <datalist id={id}>
        {options.map((o) => <option key={o} value={o} />)}
      </datalist>
    </label>
  );
}

function RosterImporter({ onImport }) {
  const [text, setText] = useState('');
  const [msg, setMsg] = useState(null);
  const submit = () => {
    if (!text.trim()) return;
    const result = onImport(text);
    const parts = [];
    if (result.addedClasses) parts.push(`${result.addedClasses} new class${result.addedClasses === 1 ? '' : 'es'}`);
    if (result.addedStudents) parts.push(`${result.addedStudents} student${result.addedStudents === 1 ? '' : 's'}`);
    setMsg(parts.length ? `Added ${parts.join(', ')}.` : 'Nothing new — those names already exist, or the format looks off.');
    setText('');
  };
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
      <textarea
        value={text}
        onChange={(e) => setText(e.target.value)}
        placeholder={'4D\nRichard Corbett\nLena Kim\n\n5K\nTom Baker\nJerry Lee'}
        rows={6}
        style={{ fontSize: 13, padding: 10, borderRadius: 8, border: `var(--border-width) solid ${LINE}`, fontFamily: 'inherit', resize: 'vertical' }}
      />
      <div style={{ fontSize: 11.5, color: MUTED }}>First line of each group is the class name. One student per line underneath. Leave a blank line between classes — paste as many classes as you like at once.</div>
      <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
        <button onClick={submit} style={{ border: 'none', background: ACCENT, color: '#fff', borderRadius: 7, padding: '8px 16px', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}>Add to roster</button>
        {msg && <span style={{ fontSize: 12, color: MUTED }}>{msg}</span>}
      </div>
    </div>
  );
}

function TopicImporter({ terms, subjects, topics, onSave }) {
  const [termName, setTermName] = useState('');
  const [subjectName, setSubjectName] = useState('');
  const [topicName, setTopicName] = useState('');
  const [criteriaText, setCriteriaText] = useState('');
  const [columns, setColumns] = useState([]);
  const [msg, setMsg] = useState(null);

  const addColumnRow = () => { if (columns.length < 4) setColumns([...columns, { name: '', type: 'note' }]); };
  const updateColumn = (i, patch) => setColumns(columns.map((c, idx) => (idx === i ? { ...c, ...patch } : c)));
  const removeColumn = (i) => setColumns(columns.filter((_, idx) => idx !== i));

  const submit = () => {
    const result = onSave({ termName, subjectName, topicName, criteriaText, columns });
    if (result.error) { setMsg(result.error); return; }
    const parts = [];
    if (result.addedCriteria) parts.push(`${result.addedCriteria} success criteria`);
    if (result.addedFields) parts.push(`${result.addedFields} extra column${result.addedFields === 1 ? '' : 's'}`);
    setMsg(parts.length ? `Saved ${parts.join(', ')}.` : 'Saved — nothing new to add.');
    setCriteriaText('');
    setColumns([]);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
      <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
        <DataListInput label="Term (optional)" value={termName} onChange={setTermName} options={terms.map((t) => t.name)} placeholder="e.g. Term 2" />
        <DataListInput label="Subject" value={subjectName} onChange={setSubjectName} options={subjects.map((s) => s.name)} placeholder="e.g. Design Technology" autoFocus />
        <DataListInput label="Topic (optional)" value={topicName} onChange={setTopicName} options={topics.map((t) => t.name)} placeholder="e.g. TinkerCAD" />
      </div>
      <div>
        <div style={{ fontSize: 12, color: MUTED, marginBottom: 4 }}>Success criteria — one per line</div>
        <textarea
          value={criteriaText}
          onChange={(e) => setCriteriaText(e.target.value)}
          placeholder={'Can create a basic 3D model\nCan apply materials and colour'}
          rows={4}
          style={{ width: '100%', fontSize: 13, padding: 10, borderRadius: 8, border: `var(--border-width) solid ${LINE}`, fontFamily: 'inherit', resize: 'vertical' }}
        />
      </div>
      <div>
        <div style={{ fontSize: 12, color: MUTED, marginBottom: 6 }}>Extra columns for this topic — optional, up to 4. Name them however you like (e.g. "PB" for Personal Best).</div>
        {columns.map((c, i) => (
          <div key={i} style={{ display: 'flex', gap: 6, marginBottom: 6, alignItems: 'center' }}>
            <input
              value={c.name}
              onChange={(e) => updateColumn(i, { name: e.target.value })}
              placeholder="Column name"
              style={{ flex: 1, fontSize: 13, padding: '7px 9px', borderRadius: 6, border: `var(--border-width) solid ${LINE}` }}
            />
            <select value={c.type} onChange={(e) => updateColumn(i, { type: e.target.value })} style={{ fontSize: 12, padding: '7px', borderRadius: 6, border: `var(--border-width) solid ${LINE}`, background: CARD, color: INK }}>
              {Object.entries(FIELD_TYPES).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
            <IconBtn onClick={() => removeColumn(i)} danger title="Remove"><Trash2 size={13} /></IconBtn>
          </div>
        ))}
        {columns.length < 4 && (
          <button onClick={addColumnRow} style={{ border: `1px dashed ${LINE}`, background: CARD, borderRadius: 7, padding: '6px 12px', fontSize: 12, cursor: 'pointer', color: MUTED }}>
            + Add extra column
          </button>
        )}
      </div>
      <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
        <button onClick={submit} style={{ border: 'none', background: ACCENT, color: '#fff', borderRadius: 7, padding: '8px 16px', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}>Save</button>
        {msg && <span style={{ fontSize: 12, color: MUTED }}>{msg}</span>}
      </div>
    </div>
  );
}

const CHART_CLASS_COLORS = ['var(--accent)', 'var(--coral)', 'var(--green)', 'var(--indigo)', 'var(--amber)', 'var(--ink-soft)'];
const FLAG_COLORS = { 'High achiever': 'var(--indigo)', 'On track': 'var(--green)', 'Needs a push': 'var(--coral)', 'Not started': 'var(--ink-muted)' };

function InfoTooltip({ text }) {
  const [open, setOpen] = useState(false);
  return (
    <span style={{ position: 'relative', display: 'inline-flex', marginLeft: 5 }}>
      <button
        onClick={(e) => { e.stopPropagation(); setOpen((o) => !o); }}
        onMouseEnter={() => setOpen(true)}
        onMouseLeave={() => setOpen(false)}
        aria-label="What does this mean?"
        style={{
          border: `1px solid ${MUTED}`, borderRadius: '50%', width: 14, height: 14, minWidth: 14,
          fontSize: 9.5, lineHeight: '12px', background: 'transparent', color: MUTED, cursor: 'pointer',
          padding: 0, display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700,
        }}
      >?</button>
      {open && (
        <div
          onMouseEnter={() => setOpen(true)}
          onMouseLeave={() => setOpen(false)}
          style={{
            position: 'absolute', bottom: '140%', left: 0, zIndex: 30, width: 220,
            background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 8,
            padding: '9px 11px', fontSize: 11.5, lineHeight: 1.4, color: INK,
            boxShadow: '0 6px 18px rgba(0,0,0,0.22)', fontWeight: 400,
          }}
        >
          {text}
        </div>
      )}
    </span>
  );
}

function MiniTrendCard({ topicName, data }) {
  const curr = data[data.length - 1];
  const prev = data[data.length - 2];
  return (
    <div style={{ border: `var(--border-width) solid ${LINE}`, borderRadius: 10, padding: 12 }}>
      <div style={{ fontSize: 12, fontWeight: 600, color: INK, marginBottom: 8 }}>{topicName}</div>
      {data.length === 1 ? (
        <div style={{ textAlign: 'center', padding: '6px 0' }}>
          <div style={{ fontSize: 20, fontWeight: 700, color: INK }}>{curr.avg}%</div>
          <div style={{ fontSize: 10.5, color: MUTED, marginTop: 2 }}>{curr.name} only so far</div>
        </div>
      ) : data.length === 2 ? (
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 12 }}>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: 10, color: MUTED }}>{prev.name}</div>
            <div style={{ fontSize: 16, fontWeight: 700, color: MUTED }}>{prev.avg}%</div>
          </div>
          <span style={{ fontSize: 14, color: curr.avg > prev.avg ? 'var(--green)' : curr.avg < prev.avg ? 'var(--coral)' : 'var(--ink-muted)' }}>
            {curr.avg > prev.avg ? '↑' : curr.avg < prev.avg ? '↓' : '→'}
          </span>
          <div style={{ textAlign: 'center' }}>
            <div style={{ fontSize: 10, color: MUTED }}>{curr.name}</div>
            <div style={{ fontSize: 20, fontWeight: 700, color: INK }}>{curr.avg}%</div>
          </div>
        </div>
      ) : (
        <div style={{ width: '100%', height: 110 }}>
          <ResponsiveContainer>
            <LineChart data={data}>
              <XAxis dataKey="name" tick={{ fontSize: 9.5, fill: 'var(--ink-soft)' }} />
              <YAxis domain={[0, 100]} tick={{ fontSize: 9.5, fill: 'var(--ink-soft)' }} width={28} />
              <Tooltip contentStyle={{ background: 'var(--card)', border: `1px solid var(--line)`, borderRadius: 8, fontSize: 11 }} />
              <Line type="monotone" dataKey="avg" stroke="var(--accent)" strokeWidth={2} dot={{ r: 3 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

function HeadlineCard({ label, value, tooltip }) {
  return (
    <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 10, padding: '12px 14px' }}>
      <div style={{ fontSize: 20, fontWeight: 700, color: INK, lineHeight: 1.15 }}>{value}</div>
      <div style={{ fontSize: 11, color: MUTED, marginTop: 3, display: 'flex', alignItems: 'center' }}>
        {label}
        <InfoTooltip text={tooltip} />
      </div>
    </div>
  );
}

function AnalyticsView({ structure, studentsByClass, records, recKey, fieldValues, fieldKey }) {
  const [lensMode, setLensMode] = useState('class');
  const [studentClassId, setStudentClassId] = useState('');
  const [studentId, setStudentId] = useState('');
  const [termId, setTermId] = useState(structure.terms[0]?.id ?? '');
  const [subjectId, setSubjectId] = useState(structure.subjects[0]?.id ?? '');
  const [topicFilterId, setTopicFilterId] = useState('');
  const [classId, setClassId] = useState('');

  useEffect(() => {
    if (!studentClassId && structure.classes[0]) setStudentClassId(structure.classes[0].id);
  }, [structure.classes]);
  useEffect(() => {
    const roster = studentsByClass[studentClassId] || [];
    if (!roster.find((s) => s.id === studentId)) setStudentId(roster[0]?.id ?? '');
  }, [studentClassId, studentsByClass]);

  useEffect(() => {
    if (!termId && structure.terms[0]) setTermId(structure.terms[0].id);
    if (!subjectId && structure.subjects[0]) setSubjectId(structure.subjects[0].id);
  }, [structure.terms, structure.subjects]);

  const topicsForSubject = useMemo(
    () => structure.topics.filter((t) => t.subjectId === subjectId && (!t.termId || t.termId === termId)),
    [structure.topics, subjectId, termId]
  );
  const criteriaForSubject = useMemo(() => {
    const withTopic = topicsForSubject.flatMap((t) => structure.criteria.filter((c) => c.topicId === t.id).map((c) => ({ ...c, topicName: t.name })));
    const noTopic = structure.criteria.filter((c) => c.subjectId === subjectId && !c.topicId && (!c.termId || c.termId === termId)).map((c) => ({ ...c, topicName: 'General' }));
    return [...withTopic, ...noTopic];
  }, [topicsForSubject, structure.criteria, subjectId, termId]);
  const criteriaInScope = useMemo(
    () => (topicFilterId ? criteriaForSubject.filter((c) => (c.topicId || 'general') === topicFilterId) : criteriaForSubject),
    [criteriaForSubject, topicFilterId]
  );

  const classesInScope = classId ? structure.classes.filter((c) => c.id === classId) : structure.classes;

  // Chart 1: % met per criteria, one series per class
  const barData = criteriaInScope.map((c) => {
    const row = { name: c.name.length > 28 ? c.name.slice(0, 26) + '…' : c.name, full: c.name };
    classesInScope.forEach((cls) => {
      const roster = studentsByClass[cls.id] || [];
      const met = roster.filter((s) => (records[recKey(s.id, termId, c.id)] ?? 0) >= 2).length;
      row[cls.name] = roster.length ? Math.round((met / roster.length) * 100) : 0;
    });
    return row;
  });

  // Chart 2: flag distribution across students in scope
  const flagCounts = { 'High achiever': 0, 'On track': 0, 'Needs a push': 0, 'Not started': 0 };
  classesInScope.forEach((cls) => {
    (studentsByClass[cls.id] || []).forEach((stu) => {
      if (!criteriaInScope.length) return;
      let total = 0, hasAny = false;
      criteriaInScope.forEach((c) => {
        const lvl = records[recKey(stu.id, termId, c.id)] ?? 0;
        total += lvl;
        if (lvl > 0) hasAny = true;
      });
      const pct = (total / (criteriaInScope.length * 3)) * 100;
      if (!hasAny) flagCounts['Not started']++;
      else if (pct >= HIGH_ACHIEVER_PCT) flagCounts['High achiever']++;
      else if (pct < NEEDS_PUSH_PCT) flagCounts['Needs a push']++;
      else flagCounts['On track']++;
    });
  });
  const pieData = Object.entries(flagCounts).filter(([, v]) => v > 0).map(([name, value]) => ({ name, value }));
  const totalStudentsInScope = Object.values(flagCounts).reduce((a, b) => a + b, 0);
  const notStartedCount = flagCounts['Not started'];
  const assessedCount = totalStudentsInScope - notStartedCount;
  const assessedPct = totalStudentsInScope ? Math.round((assessedCount / totalStudentsInScope) * 100) : 0;
  const assessedBreakdown = ['Needs a push', 'On track', 'High achiever']
    .filter((name) => flagCounts[name] > 0)
    .map((name) => ({ name, value: flagCounts[name] }));

  // Headline decision cards
  const allPairs = []; // { level }
  classesInScope.forEach((cls) => {
    (studentsByClass[cls.id] || []).forEach((stu) => {
      criteriaInScope.forEach((c) => {
        allPairs.push(records[recKey(stu.id, termId, c.id)] ?? 0);
      });
    });
  });
  const curriculumCoveragePct = allPairs.length ? Math.round((allPairs.reduce((a, b) => a + b, 0) / (allPairs.length * 3)) * 100) : 0;
  const completionPct = allPairs.length ? Math.round((allPairs.filter((l) => l > 0).length / allPairs.length) * 100) : 0;
  const secureMasteredPct = allPairs.length ? Math.round((allPairs.filter((l) => l >= 2).length / allPairs.length) * 100) : 0;

  const topicIdsInScope = [...new Set(criteriaInScope.map((c) => c.topicId).filter(Boolean))];
  const fieldsInScope = structure.fields.filter((f) => topicIdsInScope.includes(f.topicId));
  let evidenceCapturedPct = 0;
  if (fieldsInScope.length && totalStudentsInScope) {
    let withEvidence = 0;
    classesInScope.forEach((cls) => {
      (studentsByClass[cls.id] || []).forEach((stu) => {
        const hasAny = fieldsInScope.some((f) => fieldValues[fieldKey(f.id, stu.id)]);
        if (hasAny) withEvidence++;
      });
    });
    evidenceCapturedPct = Math.round((withEvidence / totalStudentsInScope) * 100);
  }

  const criteriaAttentionCount = criteriaInScope.filter((c) => {
    let met = 0, n = 0;
    classesInScope.forEach((cls) => {
      (studentsByClass[cls.id] || []).forEach((stu) => {
        n++;
        if ((records[recKey(stu.id, termId, c.id)] ?? 0) >= 2) met++;
      });
    });
    return n > 0 && (met / n) * 100 < 50;
  }).length;

  const attentionList = criteriaInScope
    .map((c) => {
      let met = 0, total = 0;
      const classesAffected = [];
      classesInScope.forEach((cls) => {
        const roster = studentsByClass[cls.id] || [];
        if (!roster.length) return;
        let classMet = 0;
        roster.forEach((stu) => {
          total++;
          if ((records[recKey(stu.id, termId, c.id)] ?? 0) >= 2) { met++; classMet++; }
        });
        if (classMet < roster.length) classesAffected.push(cls.name);
      });
      const pct = total ? Math.round((met / total) * 100) : 0;
      return { id: c.id, name: c.name, topicName: c.topicName, pct, studentsAffected: total - met, classesAffected, total };
    })
    .filter((r) => r.total > 0 && r.pct < 50)
    .sort((a, b) => a.pct - b.pct)
    .slice(0, 8);

  let consistencyLabel = null;
  let consistencyDetail = '';
  if (classesInScope.length > 1 && criteriaInScope.length) {
    const classPcts = classesInScope.map((cls) => {
      const roster = studentsByClass[cls.id] || [];
      if (!roster.length) return null;
      const t = roster.reduce((sum, stu) => sum + criteriaInScope.reduce((s2, c) => s2 + (records[recKey(stu.id, termId, c.id)] ?? 0), 0), 0);
      return Math.round((t / (roster.length * criteriaInScope.length * 3)) * 100);
    }).filter((v) => v !== null);
    if (classPcts.length > 1) {
      const spread = Math.max(...classPcts) - Math.min(...classPcts);
      consistencyLabel = spread < 15 ? 'Strong' : spread < 30 ? 'Mixed' : 'Uneven';
      consistencyDetail = `The furthest-along class is at ${Math.max(...classPcts)}%, the furthest-behind is at ${Math.min(...classPcts)}% — a gap of ${spread} percentage points. The bigger the gap, the more one class may need extra support to catch up with the others.`;
    }
  }

  // Optional: coverage by birth term (relative-age effect), only shown if any student has it set
  const birthTermBuckets = { Autumn: [], Spring: [], Summer: [] };
  classesInScope.forEach((cls) => {
    (studentsByClass[cls.id] || []).forEach((stu) => {
      if (stu.birthTerm && birthTermBuckets[stu.birthTerm]) birthTermBuckets[stu.birthTerm].push(stu);
    });
  });
  const anyBirthTermSet = Object.values(birthTermBuckets).some((arr) => arr.length > 0);
  const birthTermStats = anyBirthTermSet
    ? Object.entries(birthTermBuckets).map(([term, studs]) => {
        if (!studs.length || !criteriaInScope.length) return { term, count: studs.length, pct: 0 };
        const total = studs.reduce((sum, stu) => sum + criteriaInScope.reduce((s2, c) => s2 + (records[recKey(stu.id, termId, c.id)] ?? 0), 0), 0);
        return { term, count: studs.length, pct: Math.round((total / (studs.length * criteriaInScope.length * 3)) * 100) };
      })
    : [];

  // Chart 3: avg coverage % per term (trend), for this subject/topic + class scope
  const trendDataAll = structure.terms.map((term) => {
    const termTopics = structure.topics.filter((t) => t.subjectId === subjectId && (!t.termId || t.termId === term.id));
    let termCriteria = termTopics.flatMap((t) => structure.criteria.filter((c) => c.topicId === t.id));
    termCriteria = termCriteria.concat(structure.criteria.filter((c) => c.subjectId === subjectId && !c.topicId && (!c.termId || c.termId === term.id)));
    if (topicFilterId) termCriteria = termCriteria.filter((c) => (c.topicId || 'general') === topicFilterId);
    let sum = 0, n = 0, hasEngagement = false;
    classesInScope.forEach((cls) => {
      (studentsByClass[cls.id] || []).forEach((stu) => {
        if (!termCriteria.length) return;
        const t = termCriteria.reduce((s2, c) => s2 + (records[recKey(stu.id, term.id, c.id)] ?? 0), 0);
        sum += (t / (termCriteria.length * 3)) * 100;
        n++;
        if (termCriteria.some((c) => (records[recKey(stu.id, term.id, c.id)] ?? 0) > 0)) hasEngagement = true;
      });
    });
    return { name: term.name, avg: n ? Math.round(sum / n) : 0, hasEngagement };
  });
  // Only terms with real assessed engagement count as data points — an unassessed future
  // term (all zeros) doesn't get plotted at all, rather than looking like a drop-off.
  const trendData = trendDataAll.filter((t) => t.hasEngagement);
  const trendCurrent = trendData[trendData.length - 1];
  const trendPrevious = trendData[trendData.length - 2];

  // When "All topics" is selected, a single blended line would average unrelated topics
  // together (e.g. Football only ever taught in Term 2, T-Ball only in Term 3) as if they
  // were one continuous skill. Instead, compute one honest trend per topic.
  const topicTrends = !topicFilterId
    ? [...structure.topics.filter((t) => t.subjectId === subjectId), { id: null, name: 'General' }]
        .map((topic) => {
          const critList = topic.id
            ? structure.criteria.filter((c) => c.topicId === topic.id)
            : structure.criteria.filter((c) => c.subjectId === subjectId && !c.topicId);
          if (!critList.length) return null;
          const perTerm = structure.terms
            .map((term) => {
              let sum = 0, n = 0, hasEngagement = false;
              classesInScope.forEach((cls) => {
                (studentsByClass[cls.id] || []).forEach((stu) => {
                  const t = critList.reduce((s2, c) => s2 + (records[recKey(stu.id, term.id, c.id)] ?? 0), 0);
                  sum += (t / (critList.length * 3)) * 100;
                  n++;
                  if (critList.some((c) => (records[recKey(stu.id, term.id, c.id)] ?? 0) > 0)) hasEngagement = true;
                });
              });
              return { name: term.name, avg: n ? Math.round(sum / n) : 0, hasEngagement };
            })
            .filter((t) => t.hasEngagement);
          return perTerm.length ? { topicName: topic.name, data: perTerm } : null;
        })
        .filter(Boolean)
    : [];

  // Stage 9: plain-language summary sentences, deterministic from the numbers above —
  // no free text generation, just templates filled in from real data.
  const narrativeSentences = [];
  if (allPairs.length) {
    if (secureMasteredPct >= 70) narrativeSentences.push(`Most students are Secure or Mastered here (${secureMasteredPct}%).`);
    else if (secureMasteredPct >= 40) narrativeSentences.push(`Results are mixed so far — ${secureMasteredPct}% Secure or Mastered.`);
    else if (completionPct >= 20) narrativeSentences.push(`Most students are still Developing or not yet started (${secureMasteredPct}% Secure or Mastered so far).`);
  }
  if (attentionList.length > 0) {
    const worst = attentionList[0];
    const wherePart = worst.classesAffected.length > 1 ? ` across ${worst.classesAffected.length} classes` : worst.classesAffected.length === 1 ? ` in ${worst.classesAffected[0]}` : '';
    narrativeSentences.push(`"${worst.name}" is the weakest area right now, at ${worst.pct}%${wherePart}.`);
  }
  if (consistencyLabel === 'Uneven') narrativeSentences.push('Progress is uneven between classes — a moderation check might help before moving on.');
  else if (consistencyLabel === 'Strong' && classesInScope.length > 1) narrativeSentences.push('Classes are progressing at a similar pace.');
  if (fieldsInScope.length > 0) {
    if (evidenceCapturedPct >= 60) narrativeSentences.push(`Evidence capture is strong — ${evidenceCapturedPct}% of students have something attached (photo, note, rating or date).`);
    else if (evidenceCapturedPct > 0 && evidenceCapturedPct < 30) narrativeSentences.push(`Evidence capture is patchy so far — only ${evidenceCapturedPct}% of students have something attached.`);
  }
  if (trendData.length >= 2) {
    const delta = trendCurrent.avg - trendPrevious.avg;
    if (delta >= 4) narrativeSentences.push(`This is up ${delta} points on ${trendPrevious.name}.`);
    else if (delta <= -4) narrativeSentences.push(`This is down ${Math.abs(delta)} points on ${trendPrevious.name} — worth a look.`);
  }

  if (structure.classes.length === 0 || structure.subjects.length === 0) {
    return <EmptyState text="Add classes and success criteria in Setup to see analytics here." onAction={() => {}} />;
  }

  const studentRoster = studentsByClass[studentClassId] || [];
  const selectedStudent = studentRoster.find((s) => s.id === studentId);

  return (
    <div style={{ padding: '16px 18px', display: 'flex', flexDirection: 'column', gap: 22 }}>
      <div style={{ display: 'flex', background: PAPER, borderRadius: 999, padding: 3, border: `var(--border-width) solid ${LINE}`, width: 'fit-content' }}>
        <Pill active={lensMode === 'class'} onClick={() => setLensMode('class')}>Class view</Pill>
        <Pill
          active={lensMode === 'student'}
          onClick={() => { if (classId) setStudentClassId(classId); setLensMode('student'); }}
        >Student view</Pill>
      </div>

      {lensMode === 'student' ? (
        <>
          <div style={{ display: 'flex', gap: 6, overflowX: 'auto', paddingBottom: 2 }}>
            {structure.classes.map((c) => (
              <Pill key={c.id} active={c.id === studentClassId} onClick={() => setStudentClassId(c.id)}>{c.name}</Pill>
            ))}
          </div>
          {studentRoster.length > 0 && (
            <div style={{ display: 'flex', gap: 6, overflowX: 'auto', paddingBottom: 2 }}>
              {studentRoster.map((s) => (
                <Pill key={s.id} active={s.id === studentId} onClick={() => setStudentId(s.id)}>{s.name.split(' ')[0]}</Pill>
              ))}
            </div>
          )}
          {selectedStudent ? (
            <StudentProfileView
              key={selectedStudent.id}
              student={selectedStudent}
              className={structure.classes.find((c) => c.id === studentClassId)?.name}
              structure={structure}
              records={records}
              recKey={recKey}
              fieldValues={fieldValues}
              fieldKey={fieldKey}
            />
          ) : (
            <EmptyState text="Add a student in Setup to see their profile here." onAction={() => {}} />
          )}
        </>
      ) : (
      <>
      <div style={{ display: 'flex', gap: 16, alignItems: 'center', flexWrap: 'wrap' }}>
        <FilterSelect label="Term" value={termId} onChange={setTermId} options={structure.terms} />
        <FilterSelect label="Subject" value={subjectId} onChange={setSubjectId} options={structure.subjects} />
        <FilterSelect label="Topic" value={topicFilterId} onChange={setTopicFilterId} options={[{ id: '', name: 'All topics' }, ...topicsForSubject]} />
        <FilterSelect label="Class" value={classId} onChange={setClassId} options={[{ id: '', name: 'All classes' }, ...structure.classes]} />
      </div>

      {narrativeSentences.length > 0 && (
        <div style={{ background: 'var(--indigo-bg)', border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: '14px 16px' }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
            {narrativeSentences.map((s, i) => (
              <div key={i} style={{ fontSize: 13, color: INK, lineHeight: 1.5 }}>{s}</div>
            ))}
          </div>
        </div>
      )}

      {criteriaInScope.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 10 }}>
          <HeadlineCard
            label="Curriculum coverage"
            value={`${curriculumCoveragePct}%`}
            tooltip="How far the class has progressed through this topic overall, from not started through to fully mastered. A higher number means further along."
          />
          <HeadlineCard
            label="Assessment completion"
            value={`${completionPct}%`}
            tooltip="How much of the class has actually been checked and marked so far, whatever the result was. A higher number means fewer students still waiting to be assessed."
          />
          <HeadlineCard
            label="Secure or Mastered"
            value={`${secureMasteredPct}%`}
            tooltip="Of the marks given so far, how many show the student has properly grasped it (Secure) or fully mastered it. This is about how well they're doing, not how much has been marked."
          />
          <HeadlineCard
            label="Evidence captured"
            value={fieldsInScope.length ? `${evidenceCapturedPct}%` : '—'}
            tooltip={fieldsInScope.length ? "How many students have something extra attached as proof of their work — a photo, a written note, a rating or a date — not just a level." : 'This topic doesn\u2019t have any extra columns set up yet (like Photo, Note, Rating or Date), so there\u2019s nothing to show here.'}
          />
          <HeadlineCard
            label="Criteria needing attention"
            value={String(criteriaAttentionCount)}
            tooltip="How many goals in this topic have fewer than half the class Secure or Mastered yet. These are the ones most worth a closer look — see the list below for exactly which ones."
          />
          <HeadlineCard
            label="Class consistency"
            value={consistencyLabel || '—'}
            tooltip={consistencyLabel ? consistencyDetail : 'Choose "All classes" with more than one class selected to see whether they\u2019re all progressing at a similar pace, or very differently.'}
          />
        </div>
      )}

      {criteriaInScope.length > 0 && (
        <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4, color: INK, display: 'flex', alignItems: 'center' }}>
            Where should we look first?
            <InfoTooltip text="These are the criteria in this scope where fewer than half the class has reached Secure or Mastered yet. That can happen for a lot of reasons — timing, sequencing, absence, or assessment still in progress — it isn't pointing at any one cause. Worth a look, not a verdict." />
          </div>
          <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 14 }}>Lowest % Secure/Mastered first.</div>
          {attentionList.length === 0 ? (
            <div style={{ fontSize: 12.5, color: MUTED, padding: '20px 10px', textAlign: 'center' }}>Nothing below 50% right now in this scope — nice work.</div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              {attentionList.map((r) => (
                <div key={r.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12, padding: '10px 12px', borderRadius: 8, border: `var(--border-width) solid ${LINE}`, flexWrap: 'wrap' }}>
                  <div style={{ minWidth: 0, flex: 1 }}>
                    <div style={{ fontSize: 13, color: INK, fontWeight: 500 }}>{r.name}</div>
                    <div style={{ fontSize: 11, color: MUTED, marginTop: 2 }}>
                      {r.topicName} · {r.studentsAffected} student{r.studentsAffected === 1 ? '' : 's'} not yet Secure
                      {r.classesAffected.length > 1 ? ` · ${r.classesAffected.length} classes` : r.classesAffected.length === 1 ? ` · ${r.classesAffected[0]}` : ''}
                    </div>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexShrink: 0 }}>
                    <span style={{ fontSize: 14, fontWeight: 700, color: INK }}>{r.pct}%</span>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: 11.5, color: 'var(--coral)', background: 'var(--coral-bg)', borderRadius: 999, padding: '3px 9px', fontWeight: 500 }}>
                      Revisit
                      <InfoTooltip text="Shown because fewer students have reached Secure or Mastered here yet. This can reflect timing, sequencing, absence, or assessment still in progress — not any one cause." />
                    </span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {criteriaInScope.length === 0 ? (
        <EmptyState text="No success criteria in this scope yet." onAction={() => {}} />
      ) : (
        <>
          <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4, color: INK }}>Which success criteria need a closer look{classesInScope.length > 1 ? ', by class' : ''}?</div>
            <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 10 }}>Secure or Mastered counts as met.</div>
            <div style={{ width: '100%', height: Math.max(220, criteriaInScope.length * 46) }}>
              <ResponsiveContainer>
                <BarChart data={barData} layout="vertical" margin={{ left: 10, right: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" horizontal={false} />
                  <XAxis type="number" domain={[0, 100]} tick={{ fontSize: 11, fill: 'var(--ink-soft)' }} unit="%" />
                  <YAxis type="category" dataKey="name" width={170} tick={{ fontSize: 11, fill: 'var(--ink-soft)' }} />
                  <Tooltip
                    contentStyle={{ background: 'var(--card)', border: `1px solid var(--line)`, borderRadius: 8, fontSize: 12 }}
                    labelFormatter={(_, payload) => payload?.[0]?.payload?.full || ''}
                  />
                  {classesInScope.length > 1 && <Legend wrapperStyle={{ fontSize: 12 }} />}
                  {classesInScope.map((cls, i) => (
                    <Bar key={cls.id} dataKey={cls.name} fill={CHART_CLASS_COLORS[i % CHART_CLASS_COLORS.length]} radius={[0, 4, 4, 0]} barSize={classesInScope.length > 1 ? 12 : 18} />
                  ))}
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'minmax(260px, 1fr) minmax(300px, 1.3fr)', gap: 16 }}>
            <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4, color: INK }}>How are students doing overall?</div>
              <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 10 }}>{totalStudentsInScope} student{totalStudentsInScope === 1 ? '' : 's'} in scope.</div>
              {totalStudentsInScope === 0 ? (
                <div style={{ fontSize: 12.5, color: MUTED, padding: '40px 10px', textAlign: 'center' }}>No students in scope yet.</div>
              ) : (
                <>
                  <div style={{ marginBottom: 18 }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11.5, color: MUTED, marginBottom: 5 }}>
                      <span>Assessed so far</span>
                      <span>{assessedCount} of {totalStudentsInScope} · {assessedPct}%</span>
                    </div>
                    <div style={{ height: 10, borderRadius: 5, overflow: 'hidden', background: 'var(--line)' }}>
                      <div style={{ height: '100%', width: `${assessedPct}%`, background: ACCENT, borderRadius: 5 }} />
                    </div>
                  </div>

                  {assessedCount === 0 ? (
                    <div style={{ fontSize: 12.5, color: MUTED, padding: '10px 0', textAlign: 'center' }}>Nobody in this scope has been assessed yet.</div>
                  ) : (
                    <>
                      <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 5 }}>Of those assessed, how are they doing?</div>
                      <div style={{ display: 'flex', height: 34, borderRadius: 8, overflow: 'hidden', border: `var(--border-width) solid ${LINE}` }}>
                        {assessedBreakdown.map((entry) => (
                          <div
                            key={entry.name}
                            title={`${entry.name}: ${entry.value} (${Math.round((entry.value / assessedCount) * 100)}%)`}
                            style={{ flex: entry.value, background: FLAG_COLORS[entry.name] }}
                          />
                        ))}
                      </div>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 14, marginTop: 14 }}>
                        {assessedBreakdown.map((entry) => (
                          <div key={entry.name} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12 }}>
                            <span style={{ width: 10, height: 10, borderRadius: 3, background: FLAG_COLORS[entry.name], flexShrink: 0 }} />
                            <span style={{ color: INK }}>{entry.name}</span>
                            <span style={{ color: MUTED }}>{entry.value} ({Math.round((entry.value / assessedCount) * 100)}%)</span>
                          </div>
                        ))}
                      </div>
                    </>
                  )}
                </>
              )}
            </div>

            <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4, color: INK }}>What's changed since last term?</div>
              <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 10 }}>Progress over time for this subject{topicFilterId ? ' / topic' : ''}. Only terms with real assessed data are counted.</div>
              {!topicFilterId && topicTrends.length > 1 ? (
                <>
                  <div style={{ fontSize: 11, color: MUTED, marginBottom: 10 }}>
                    Shown per topic rather than blended together, since different topics (like a term-long unit taught only once) aren't the same thing to track as one line.
                  </div>
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 10 }}>
                    {topicTrends.map((tt) => (
                      <MiniTrendCard key={tt.topicName} topicName={tt.topicName} data={tt.data} />
                    ))}
                  </div>
                </>
              ) : trendData.length === 0 ? (
                <div style={{ fontSize: 12.5, color: MUTED, padding: '30px 10px', textAlign: 'center' }}>No terms have been assessed yet in this scope.</div>
              ) : trendData.length === 1 ? (
                <div style={{ padding: '20px 10px', textAlign: 'center' }}>
                  <div style={{ fontSize: 28, fontWeight: 700, color: INK }}>{trendCurrent.avg}%</div>
                  <div style={{ fontSize: 11.5, color: MUTED, marginTop: 4 }}>{trendCurrent.name} — the only assessed term so far. Add another term's worth of data to see a trend.</div>
                </div>
              ) : trendData.length === 2 ? (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 24, padding: '18px 10px' }}>
                  <div style={{ textAlign: 'center' }}>
                    <div style={{ fontSize: 11.5, color: MUTED }}>{trendPrevious.name}</div>
                    <div style={{ fontSize: 24, fontWeight: 700, color: MUTED }}>{trendPrevious.avg}%</div>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
                    <span style={{ fontSize: 20, color: trendCurrent.avg > trendPrevious.avg ? 'var(--green)' : trendCurrent.avg < trendPrevious.avg ? 'var(--coral)' : 'var(--ink-muted)' }}>
                      {trendCurrent.avg > trendPrevious.avg ? '↑' : trendCurrent.avg < trendPrevious.avg ? '↓' : '→'}
                    </span>
                    <span style={{ fontSize: 11, color: MUTED }}>{trendCurrent.avg - trendPrevious.avg > 0 ? '+' : ''}{trendCurrent.avg - trendPrevious.avg} pts</span>
                  </div>
                  <div style={{ textAlign: 'center' }}>
                    <div style={{ fontSize: 11.5, color: MUTED }}>{trendCurrent.name}</div>
                    <div style={{ fontSize: 28, fontWeight: 700, color: INK }}>{trendCurrent.avg}%</div>
                  </div>
                </div>
              ) : (
                <div style={{ width: '100%', height: 220 }}>
                  <ResponsiveContainer>
                    <LineChart data={trendData}>
                      <CartesianGrid strokeDasharray="3 3" stroke="var(--line)" />
                      <XAxis dataKey="name" tick={{ fontSize: 11, fill: 'var(--ink-soft)' }} />
                      <YAxis domain={[0, 100]} tick={{ fontSize: 11, fill: 'var(--ink-soft)' }} unit="%" />
                      <Tooltip contentStyle={{ background: 'var(--card)', border: `1px solid var(--line)`, borderRadius: 8, fontSize: 12 }} />
                      <Line type="monotone" dataKey="avg" stroke="var(--accent)" strokeWidth={2.5} dot={{ r: 4 }} name="Avg coverage %" />
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              )}
            </div>
          </div>

          {anyBirthTermSet && (
            <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
              <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 4, color: INK, display: 'flex', alignItems: 'center' }}>
                Coverage by birth term
                <InfoTooltip text="Groups students by which term of the school year they were born in, to check for the relative-age effect — younger-in-year children (Summer-born) sometimes appear behind simply due to age, not ability. Only students with a birth term set in Setup are included." />
              </div>
              <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 14 }}>Only students with a birth term set in Setup are included here.</div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                {birthTermStats.filter((b) => b.count > 0).map((b) => (
                  <div key={b.term}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: INK, marginBottom: 4 }}>
                      <span>{b.term}-born <span style={{ color: MUTED }}>({b.count} student{b.count === 1 ? '' : 's'})</span></span>
                      <span style={{ color: MUTED }}>{b.pct}%</span>
                    </div>
                    <div style={{ height: 9, borderRadius: 5, overflow: 'hidden', background: 'var(--line)' }}>
                      <div style={{ height: '100%', width: `${b.pct}%`, background: ACCENT, borderRadius: 5 }} />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )}
        </>
      )}
      </>
      )}
    </div>
  );
}

function computeStudentSubjectBuckets(studentId, structure, records, recKey) {
  return structure.subjects
    .map((subject) => {
      const topics = structure.topics.filter((t) => t.subjectId === subject.id);
      const allCrit = [
        ...topics.flatMap((t) => structure.criteria.filter((c) => c.topicId === t.id).map((c) => ({ ...c, topicName: t.name }))),
        ...structure.criteria.filter((c) => c.subjectId === subject.id && !c.topicId).map((c) => ({ ...c, topicName: 'General' })),
      ];
      const strengths = [];
      const developing = [];
      allCrit.forEach((c) => {
        let level = 0;
        structure.terms.forEach((term) => {
          const l = records[recKey(studentId, term.id, c.id)] ?? 0;
          if (l > level) level = l;
        });
        if (level >= 2) strengths.push({ ...c, level });
        else if (level === 1) developing.push({ ...c, level });
      });
      return { subject, strengths, developing };
    })
    .filter((g) => g.strengths.length || g.developing.length);
}

function computeStudentEvidence(studentId, structure, fieldValues, fieldKey) {
  return structure.fields
    .map((f) => {
      const val = fieldValues[fieldKey(f.id, studentId)];
      if (!val) return null;
      const topic = structure.topics.find((t) => t.id === f.topicId);
      return { field: f, value: val, topicName: topic?.name || '' };
    })
    .filter(Boolean);
}

function computeSubjectTrend(studentId, subject, structure, records, recKey) {
  const critList = [
    ...structure.topics.filter((t) => t.subjectId === subject.id).flatMap((t) => structure.criteria.filter((c) => c.topicId === t.id)),
    ...structure.criteria.filter((c) => c.subjectId === subject.id && !c.topicId),
  ];
  if (!critList.length) return null;
  const perTerm = structure.terms
    .map((term) => {
      let sum = 0, hasEngagement = false;
      critList.forEach((c) => {
        const lvl = records[recKey(studentId, term.id, c.id)] ?? 0;
        sum += lvl;
        if (lvl > 0) hasEngagement = true;
      });
      return { name: term.name, avg: Math.round((sum / (critList.length * 3)) * 100), hasEngagement };
    })
    .filter((t) => t.hasEngagement);
  return perTerm.length ? { subjectName: subject.name, data: perTerm } : null;
}

function StudentProfileView({ student, className, structure, records, recKey, fieldValues, fieldKey }) {
  const buckets = useMemo(() => computeStudentSubjectBuckets(student.id, structure, records, recKey), [student.id, structure, records, recKey]);
  const evidence = useMemo(() => computeStudentEvidence(student.id, structure, fieldValues, fieldKey), [student.id, structure, fieldValues, fieldKey]);
  const subjectTrends = useMemo(
    () => structure.subjects.map((s) => computeSubjectTrend(student.id, s, structure, records, recKey)).filter(Boolean),
    [student.id, structure, records, recKey]
  );

  const [photoThumbs, setPhotoThumbs] = useState({});
  const photoThumbsRef = useRef({});
  useEffect(() => {
    let cancelled = false;
    evidence.filter((e) => e.field.type === 'photo').forEach((e) => {
      const key = e.field.id;
      if (key in photoThumbsRef.current) return;
      photoThumbsRef.current[key] = null;
      loadKey(`photo:${e.field.id}:${student.id}`, null).then((url) => {
        if (cancelled) return;
        photoThumbsRef.current[key] = url;
        setPhotoThumbs((prev) => ({ ...prev, [key]: url }));
      });
    });
    return () => { cancelled = true; };
  }, [evidence, student.id]);

  const totalStrengths = buckets.reduce((sum, b) => sum + b.strengths.length, 0);
  const totalDeveloping = buckets.reduce((sum, b) => sum + b.developing.length, 0);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
      <div>
        <div style={{ fontSize: 20, fontWeight: 700, color: INK }}>{student.name}</div>
        <div style={{ fontSize: 12.5, color: MUTED }}>{className} · everything recorded so far, all subjects</div>
      </div>

      <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: INK, marginBottom: 10 }}>Strengths</div>
        {totalStrengths === 0 ? (
          <div style={{ fontSize: 12.5, color: MUTED }}>Nothing marked Secure or Mastered yet.</div>
        ) : (
          buckets.filter((b) => b.strengths.length).map((b) => (
            <div key={b.subject.id} style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: MUTED, marginBottom: 5 }}>{b.subject.name}</div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
                {b.strengths.map((c) => (
                  <div key={c.id} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: INK }}>
                    <LevelPill level={c.level} width={14} height={14} />
                    <span>{c.name}</span>
                  </div>
                ))}
              </div>
            </div>
          ))
        )}
      </div>

      <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: INK, marginBottom: 10 }}>Currently developing</div>
        {totalDeveloping === 0 ? (
          <div style={{ fontSize: 12.5, color: MUTED }}>Nothing currently in progress.</div>
        ) : (
          buckets.filter((b) => b.developing.length).map((b) => (
            <div key={b.subject.id} style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 12, fontWeight: 600, color: MUTED, marginBottom: 5 }}>{b.subject.name}</div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
                {b.developing.map((c) => (
                  <div key={c.id} style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, color: INK }}>
                    <LevelPill level={c.level} width={14} height={14} />
                    <span>{c.name}</span>
                  </div>
                ))}
              </div>
            </div>
          ))
        )}
      </div>

      <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: INK, marginBottom: 10 }}>Evidence collected</div>
        {evidence.length === 0 ? (
          <div style={{ fontSize: 12.5, color: MUTED }}>No notes, ratings, dates or photos attached yet.</div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {evidence.map((e) => (
              <div key={e.field.id} style={{ display: 'flex', alignItems: 'center', gap: 10, fontSize: 13 }}>
                <FieldTypeIcon type={e.field.type} size={14} color={MUTED} />
                <span style={{ color: MUTED, minWidth: 130 }}>{e.topicName} · {e.field.name}</span>
                {e.field.type === 'photo' ? (
                  photoThumbs[e.field.id] ? (
                    <img src={photoThumbs[e.field.id]} alt="" style={{ width: 32, height: 32, borderRadius: 6, objectFit: 'cover' }} />
                  ) : (
                    <span style={{ color: MUTED, fontSize: 11.5 }}>Loading…</span>
                  )
                ) : e.field.type === 'rating' ? (
                  <span style={{ color: AMBER }}>{'★'.repeat(e.value)}{'☆'.repeat(Math.max(0, 5 - e.value))}</span>
                ) : (
                  <span style={{ color: INK }}>{e.value}</span>
                )}
              </div>
            ))}
          </div>
        )}
      </div>

      {subjectTrends.length > 0 && (
        <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 16 }}>
          <div style={{ fontSize: 13, fontWeight: 600, color: INK, marginBottom: 4 }}>Progress over time</div>
          <div style={{ fontSize: 11, color: MUTED, marginBottom: 10 }}>This student's own trend, by subject — not compared to anyone else.</div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))', gap: 10 }}>
            {subjectTrends.map((st) => (
              <MiniTrendCard key={st.subjectName} topicName={st.subjectName} data={st.data} />
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function SetupView({ structure, studentsByClass, setupClassId, setSetupClassId, setupSubjectId, setSetupSubjectId, addClass, removeClass, addStudent, removeStudent, setStudentBirthTerm, setClassYearGroup, addSubject, removeSubject, setSubjectDepartment, addTopic, removeTopic, addCriteria, removeCriteria, addField, removeField, addTerm, removeTerm, addYearGroup, removeYearGroup, addDepartment, removeDepartment, downloadTemplate, importFile, importing, importMsg, setImportMsg, resetAll, importRoster, saveTopicContent, togglePinnedSubject }) {
  const roster = studentsByClass[setupClassId] || [];
  const topics = structure.topics.filter((t) => t.subjectId === setupSubjectId);
  const criteriaNoTopic = structure.criteria.filter((c) => c.subjectId === setupSubjectId && !c.topicId);
  const fileRef = useRef(null);
  const [setupTermId, setSetupTermId] = useState('');
  const [showRosterList, setShowRosterList] = useState(false);
  const [showCriteriaList, setShowCriteriaList] = useState(false);
  const [showSpreadsheet, setShowSpreadsheet] = useState(false);
  const hasData = structure.classes.length > 0 || structure.subjects.length > 0 || structure.terms.length > 0;
  const totalStudents = Object.values(studentsByClass).reduce((sum, arr) => sum + arr.length, 0);
  const onFilePicked = (e) => {
    const f = e.target.files?.[0];
    if (f) importFile(f);
    e.target.value = '';
  };

  return (
    <div style={{ padding: 16, display: 'grid', gridTemplateColumns: '1fr', gap: 16, maxWidth: 640, margin: '0 auto' }}>
      {hasData && (
        <button
          onClick={resetAll}
          style={{ alignSelf: 'flex-end', border: 'none', background: 'transparent', color: '#9A4A3A', fontSize: 12, cursor: 'pointer', textDecoration: 'underline', padding: 2 }}
        >
          Starting fresh? Clear all existing data
        </button>
      )}

      <Section title="1 · Classes & students">
        <RosterImporter onImport={importRoster} />
        <button
          onClick={() => setShowRosterList((v) => !v)}
          style={{ display: 'flex', alignItems: 'center', gap: 5, border: 'none', background: 'transparent', color: MUTED, fontSize: 12, cursor: 'pointer', padding: '4px 0', alignSelf: 'flex-start' }}
        >
          {showRosterList ? 'Hide' : 'Show'} classes & students ({structure.classes.length} classes, {totalStudents} students)
          <ChevronRight size={11} style={{ transform: showRosterList ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s' }} />
        </button>
        {showRosterList && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.4fr', gap: 14, alignItems: 'start', marginTop: 4 }}>
            <div style={{ border: `var(--border-width) solid ${LINE}`, borderRadius: 8, padding: 10 }}>
              <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 6 }}>Classes</div>
              {structure.classes.map((c) => (
                <div key={c.id} onClick={() => setSetupClassId(c.id)} style={{ cursor: 'pointer' }}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 6, padding: '6px 8px', borderRadius: 6, background: c.id === setupClassId ? 'var(--hover-bg)' : 'transparent', fontSize: 13 }}>
                    <span style={{ display: 'flex', alignItems: 'center', gap: 6, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}><Users size={13} color={MUTED} />{c.name}</span>
                    <select
                      value={c.yearGroup || ''}
                      onChange={(e) => setClassYearGroup(c.id, e.target.value)}
                      onClick={(e) => e.stopPropagation()}
                      title="Year group (optional) — for future Year Group analytics only. Add new ones in the Year groups list above."
                      style={{ fontSize: 11, padding: '3px 6px', borderRadius: 5, border: `var(--border-width) solid ${LINE}`, background: CARD, color: c.yearGroup ? INK : MUTED, width: 92, flexShrink: 0 }}
                    >
                      <option value="">Year group —</option>
                      {structure.yearGroups.map((yg) => <option key={yg.id} value={yg.name}>{yg.name}</option>)}
                    </select>
                    <IconBtn onClick={(e) => { e.stopPropagation(); removeClass(c.id); }} danger title="Remove class"><Trash2 size={13} /></IconBtn>
                  </div>
                </div>
              ))}
              <div style={{ marginTop: 6 }}><TextAdd placeholder="Add one class" onAdd={addClass} /></div>
            </div>
            <div style={{ border: `var(--border-width) solid ${LINE}`, borderRadius: 8, padding: 10 }}>
              <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 6 }}>
                {structure.classes.find((c) => c.id === setupClassId)?.name ? `Students — ${structure.classes.find((c) => c.id === setupClassId).name}` : 'Students'}
              </div>
              {setupClassId ? (
                <>
                  {roster.length === 0 && <div style={{ fontSize: 12, color: MUTED }}>No students yet.</div>}
                  {roster.map((s) => (
                    <div key={s.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8, padding: '5px 8px', borderRadius: 6, fontSize: 13 }}
                      onMouseEnter={(e) => (e.currentTarget.style.background = 'var(--hover-bg)')}
                      onMouseLeave={(e) => (e.currentTarget.style.background = 'transparent')}
                    >
                      <span style={{ flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{s.name}</span>
                      <select
                        value={s.birthTerm || ''}
                        onChange={(e) => setStudentBirthTerm(setupClassId, s.id, e.target.value)}
                        title="Birth term (optional) — for relative-age analytics only, never shown while assessing"
                        style={{ fontSize: 11, padding: '3px 5px', borderRadius: 5, border: `var(--border-width) solid ${LINE}`, background: CARD, color: s.birthTerm ? INK : MUTED, flexShrink: 0 }}
                      >
                        <option value="">Birth term —</option>
                        <option value="Autumn">Autumn born</option>
                        <option value="Spring">Spring born</option>
                        <option value="Summer">Summer born</option>
                      </select>
                      <IconBtn onClick={() => removeStudent(setupClassId, s.id)} danger title="Remove"><Trash2 size={13} /></IconBtn>
                    </div>
                  ))}
                  <div style={{ marginTop: 6 }}><TextAdd placeholder="Add one student" onAdd={(name) => addStudent(setupClassId, name)} /></div>
                </>
              ) : (
                <div style={{ fontSize: 12, color: MUTED }}>Select or add a class first.</div>
              )}
            </div>
          </div>
        )}
        {showRosterList && setupClassId && structure.subjects.length > 0 && (
          <div style={{ border: `var(--border-width) dashed ${LINE}`, borderRadius: 8, padding: 10, marginTop: 10 }}>
            <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 2 }}>
              Pinned subjects for {structure.classes.find((c) => c.id === setupClassId)?.name} — quick-switch pills in Assess
            </div>
            <div style={{ fontSize: 11, color: MUTED, marginBottom: 8 }}>
              Star the subjects this class actually rotates through, so switching between them mid-day is one tap instead of a dropdown search.
            </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
              {structure.subjects.map((s) => {
                const cls = structure.classes.find((c) => c.id === setupClassId);
                const pinned = (cls?.pinnedSubjectIds || []).includes(s.id);
                return (
                  <button
                    key={s.id}
                    onClick={() => togglePinnedSubject(setupClassId, s.id)}
                    style={{
                      display: 'flex', alignItems: 'center', gap: 5, border: `var(--border-width) solid ${pinned ? ACCENT : LINE}`,
                      background: pinned ? 'var(--indigo-bg)' : CARD, color: pinned ? INK : MUTED,
                      borderRadius: 999, padding: '5px 10px', fontSize: 12, cursor: 'pointer',
                    }}
                  >
                    <Star size={11} color={pinned ? AMBER : MUTED} fill={pinned ? AMBER : 'none'} />
                    {s.name}
                  </button>
                );
              })}
            </div>
          </div>
        )}
      </Section>

      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
        <span style={{ fontSize: 12, color: MUTED }}>Terms:</span>
        {structure.terms.map((t) => (
          <span key={t.id} style={{ display: 'flex', alignItems: 'center', gap: 4, border: `var(--border-width) solid ${LINE}`, borderRadius: 999, padding: '3px 9px', fontSize: 12, background: CARD }}>
            {t.name}
            <IconBtn onClick={() => removeTerm(t.id)} danger title="Remove"><Trash2 size={11} /></IconBtn>
          </span>
        ))}
        <span style={{ minWidth: 160 }}><TextAdd placeholder="Add a term" onAdd={addTerm} /></span>
      </div>

      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
        <span style={{ fontSize: 12, color: MUTED }} title="One shared list everyone picks from — set the exact spelling once here (e.g. 'Year 4') so it can't drift into 'Y4' / '4' elsewhere.">Year groups:</span>
        {structure.yearGroups.map((yg) => (
          <span key={yg.id} style={{ display: 'flex', alignItems: 'center', gap: 4, border: `var(--border-width) solid ${LINE}`, borderRadius: 999, padding: '3px 9px', fontSize: 12, background: CARD }}>
            {yg.name}
            <IconBtn onClick={() => removeYearGroup(yg.id)} danger title="Remove"><Trash2 size={11} /></IconBtn>
          </span>
        ))}
        <span style={{ minWidth: 160 }}><TextAdd placeholder="Add a year group" onAdd={addYearGroup} /></span>
      </div>

      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', alignItems: 'center' }}>
        <span style={{ fontSize: 12, color: MUTED }} title="One shared list — groups subjects together (e.g. Technology). Optional; only useful in secondary/international settings.">Departments:</span>
        {structure.departments.map((d) => (
          <span key={d.id} style={{ display: 'flex', alignItems: 'center', gap: 4, border: `var(--border-width) solid ${LINE}`, borderRadius: 999, padding: '3px 9px', fontSize: 12, background: CARD }}>
            {d.name}
            <IconBtn onClick={() => removeDepartment(d.id)} danger title="Remove"><Trash2 size={11} /></IconBtn>
          </span>
        ))}
        <span style={{ minWidth: 160 }}><TextAdd placeholder="Add a department" onAdd={addDepartment} /></span>
      </div>

      <Section title="2 · Subjects, topics, success criteria & extra columns">
        <TopicImporter terms={structure.terms} subjects={structure.subjects} topics={structure.topics} onSave={saveTopicContent} />
        <button
          onClick={() => setShowCriteriaList((v) => !v)}
          style={{ display: 'flex', alignItems: 'center', gap: 5, border: 'none', background: 'transparent', color: MUTED, fontSize: 12, cursor: 'pointer', padding: '4px 0', alignSelf: 'flex-start' }}
        >
          {showCriteriaList ? 'Hide' : 'Show'} subjects & criteria ({structure.subjects.length} subjects, {structure.criteria.length} criteria, {structure.fields.length} extra columns)
          <ChevronRight size={11} style={{ transform: showCriteriaList ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s' }} />
        </button>
        {showCriteriaList && (
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1.4fr', gap: 14, alignItems: 'start', marginTop: 4 }}>
            <div style={{ border: `var(--border-width) solid ${LINE}`, borderRadius: 8, padding: 10 }}>
              <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 6 }}>Subjects</div>
              {structure.subjects.map((s) => (
                <div key={s.id} onClick={() => setSetupSubjectId(s.id)} style={{ cursor: 'pointer' }}>
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 6, padding: '6px 8px', borderRadius: 6, background: s.id === setupSubjectId ? 'var(--hover-bg)' : 'transparent', fontSize: 13 }}>
                    <span style={{ minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{s.name}</span>
                    <select
                      value={s.department || ''}
                      onChange={(e) => setSubjectDepartment(s.id, e.target.value)}
                      onClick={(e) => e.stopPropagation()}
                      title="Department (optional) — groups subjects together for future Department analytics. Add new ones in the Departments list above."
                      style={{ fontSize: 11, padding: '3px 6px', borderRadius: 5, border: `var(--border-width) solid ${LINE}`, background: CARD, color: s.department ? INK : MUTED, width: 100, flexShrink: 0 }}
                    >
                      <option value="">Department —</option>
                      {structure.departments.map((d) => <option key={d.id} value={d.name}>{d.name}</option>)}
                    </select>
                    <IconBtn onClick={(e) => { e.stopPropagation(); removeSubject(s.id); }} danger title="Remove subject"><Trash2 size={13} /></IconBtn>
                  </div>
                </div>
              ))}
              <div style={{ marginTop: 6 }}><TextAdd placeholder="Add one subject" onAdd={addSubject} /></div>
            </div>

            <div style={{ border: `var(--border-width) solid ${LINE}`, borderRadius: 8, padding: 10 }}>
              <div style={{ fontSize: 11.5, color: MUTED, marginBottom: 6 }}>
                {structure.subjects.find((s) => s.id === setupSubjectId)?.name ? `Success criteria — ${structure.subjects.find((s) => s.id === setupSubjectId).name}` : 'Success criteria'}
              </div>
              {setupSubjectId ? (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                  {structure.terms.length > 0 && (
                    <FilterSelect
                      label="New topics go in"
                      value={setupTermId}
                      onChange={setSetupTermId}
                      options={[{ id: '', name: 'Every term (no specific term)' }, ...structure.terms]}
                    />
                  )}
                  {topics.map((t) => (
                    <div key={t.id} style={{ border: `var(--border-width) solid ${LINE}`, borderRadius: 8, padding: 8 }}>
                      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 4 }}>
                        <span style={{ fontSize: 12.5, fontWeight: 500, display: 'flex', alignItems: 'center', gap: 4 }}>
                          <ChevronRight size={12} />{t.name}
                          <span style={{ fontWeight: 400, color: MUTED, fontSize: 11 }}>
                            · {structure.terms.find((tm) => tm.id === t.termId)?.name || 'every term'}
                          </span>
                        </span>
                        <IconBtn onClick={() => removeTopic(t.id)} danger title="Remove topic"><Trash2 size={12} /></IconBtn>
                      </div>
                      {structure.criteria.filter((c) => c.topicId === t.id).map((c) => (
                        <ListRow key={c.id} label={c.name} onRemove={() => removeCriteria(c.id)} />
                      ))}
                      <div style={{ marginTop: 4 }}>
                        <TextAdd placeholder="Add success criteria" onAdd={(name) => addCriteria(setupSubjectId, t.id, name)} />
                      </div>
                      {structure.fields.filter((f) => f.topicId === t.id).length > 0 && (
                        <div style={{ marginTop: 8, paddingTop: 8, borderTop: `1px dashed ${LINE}` }}>
                          {structure.fields.filter((f) => f.topicId === t.id).map((f) => (
                            <ListRow
                              key={f.id}
                              label={<span style={{ display: 'flex', alignItems: 'center', gap: 5 }}><FieldTypeIcon type={f.type} size={12} />{f.name} <span style={{ color: MUTED, fontSize: 10.5 }}>· {FIELD_TYPES[f.type].label}</span></span>}
                              onRemove={() => removeField(f.id)}
                            />
                          ))}
                        </div>
                      )}
                      <div style={{ marginTop: 6 }}>
                        <FieldAdd onAdd={(name, type) => addField(t.id, name, type)} />
                      </div>
                    </div>
                  ))}
                  {criteriaNoTopic.length > 0 && (
                    <div style={{ border: `1px dashed ${LINE}`, borderRadius: 8, padding: 8 }}>
                      <div style={{ fontSize: 12.5, fontWeight: 500, marginBottom: 4, color: MUTED }}>General (no topic)</div>
                      {criteriaNoTopic.map((c) => (
                        <ListRow key={c.id} label={`${c.name} · ${structure.terms.find((tm) => tm.id === c.termId)?.name || 'every term'}`} onRemove={() => removeCriteria(c.id)} />
                      ))}
                    </div>
                  )}
                  <TextAdd placeholder="Add topic, e.g. Fractions" onAdd={(name) => addTopic(setupSubjectId, setupTermId, name)} />
                  <TextAdd placeholder="Add criteria without a topic" onAdd={(name) => addCriteria(setupSubjectId, null, name, setupTermId)} />
                </div>
              ) : (
                <div style={{ fontSize: 12, color: MUTED }}>Select or add a subject first.</div>
              )}
            </div>
          </div>
        )}
      </Section>

      <button
        onClick={() => setShowSpreadsheet((v) => !v)}
        style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, border: 'none', background: 'transparent', color: MUTED, fontSize: 12.5, cursor: 'pointer', padding: 6 }}
      >
        {showSpreadsheet ? 'Hide spreadsheet upload' : 'Rolling out a whole year group? Upload a spreadsheet instead'}
        <ChevronRight size={12} style={{ transform: showSpreadsheet ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s' }} />
      </button>
      {showSpreadsheet && (
        <div style={{ background: CARD, border: `var(--border-width) solid ${LINE}`, borderRadius: 12, padding: 20, display: 'flex', flexDirection: 'column', gap: 12, alignItems: 'center', textAlign: 'center' }}>
          <div style={{ fontSize: 13, color: MUTED, maxWidth: 420 }}>
            For large bulk uploads: download the sheet, paste in your students and success criteria, upload it back.
          </div>
          <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', justifyContent: 'center' }}>
            <button
              onClick={downloadTemplate}
              style={{ display: 'flex', alignItems: 'center', gap: 7, padding: '9px 16px', borderRadius: 9, border: `var(--border-width) solid ${LINE}`, background: CARD, fontSize: 13, fontWeight: 500, cursor: 'pointer', color: INK }}
            >
              <Download size={14} /> Download sheet
            </button>
            <button
              onClick={() => fileRef.current?.click()}
              disabled={importing}
              style={{ display: 'flex', alignItems: 'center', gap: 7, padding: '9px 16px', borderRadius: 9, border: 'none', background: ACCENT, color: '#fff', fontSize: 13, fontWeight: 500, cursor: 'pointer' }}
            >
              {importing ? <Loader2 size={14} className="spin" /> : <Upload size={14} />}
              Upload it back
            </button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={onFilePicked} style={{ display: 'none' }} />
          </div>
          {importMsg && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: 13, background: PAPER, borderRadius: 8, padding: '9px 12px', width: '100%', justifyContent: 'center' }}>
              <span>{importMsg}</span>
              <IconBtn onClick={() => setImportMsg(null)} title="Dismiss"><X size={13} /></IconBtn>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
