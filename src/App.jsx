import { useState, useRef, useEffect } from "react";

const uid = () => '_' + Math.random().toString(36).slice(2, 9);

const PALETTE = [
  { bg: '#1b3d6e', text: '#fff', light: '#d2e5f5' },
  { bg: '#1a5c38', text: '#fff', light: '#c4edd6' },
  { bg: '#6d1f78', text: '#fff', light: '#e8d3f4' },
  { bg: '#b55000', text: '#fff', light: '#ffddb2' },
  { bg: '#1a5c5c', text: '#fff', light: '#c4edee' },
  { bg: '#7a1a2e', text: '#fff', light: '#f5c8d2' },
  { bg: '#3a5a00', text: '#fff', light: '#d8f0a5' },
  { bg: '#33337a', text: '#fff', light: '#d0d0f5' },
];
const pc = idx => PALETTE[(idx ?? 0) % PALETTE.length];

// ── Results Code Schema (Mass Ingenuity) ──────────────────────────────────────
// Format: {mapNum:2}{type:2}{seq:2}{letter:1} e.g. 01OP01A
const pad2 = n => String(Math.max(1, n)).padStart(2, '0');
const nextLetter = n => String.fromCharCode(65 + Math.min(n, 25)); // 0→A, 1→B …
const mkMeasure = (text = '', code = '') => ({ id: uid(), code, text, measureOwner: '' });
const procCodeStr  = (mapNum, type, n) => `${pad2(mapNum || 1)}${type}${pad2(n)}`;
const measCodeStr  = (mapNum, type, procN, mIdx) => `${pad2(mapNum || 1)}${type}${pad2(procN)}${nextLetter(mIdx)}`;
const outcomeCodeStr = (mapNum, n) => `${pad2(mapNum || 1)}OM${pad2(n)}`;
const outMeasCodeStr = (mapNum, outN, mIdx) => `${pad2(mapNum || 1)}OM${pad2(outN)}${nextLetter(mIdx)}`;

// ── Shared column builder ─────────────────────────────────────────────────────
function buildAllProcs(data) {
  const { keyGoals, operatingProcesses, supportingProcesses } = data;
  const allProcs = [], seen = new Set();
  for (const goal of keyGoals) {
    for (const p of operatingProcesses) if (p.goalId === goal.id) { allProcs.push({ proc: p, side: 'op', goalId: goal.id }); seen.add(p.id); }
    for (const p of supportingProcesses) if (p.goalId === goal.id) { allProcs.push({ proc: p, side: 'sp', goalId: goal.id }); seen.add(p.id); }
  }
  for (const p of operatingProcesses)  if (!seen.has(p.id)) { allProcs.push({ proc: p, side: 'op', goalId: null }); seen.add(p.id); }
  for (const p of supportingProcesses) if (!seen.has(p.id)) { allProcs.push({ proc: p, side: 'sp', goalId: null }); seen.add(p.id); }
  return allProcs;
}

function buildGoalSpans(allProcs, keyGoals) {
  const spans = []; let i = 0;
  while (i < allProcs.length) {
    const gid = allProcs[i].goalId || null;
    const goal = gid ? keyGoals.find(g => g.id === gid) : null;
    let count = 0;
    while (i < allProcs.length && (allProcs[i].goalId || null) === gid) { count++; i++; }
    spans.push({ goal, gid, count });
  }
  return spans;
}

function sortedOutcomes(data) {
  return [...data.outcomes].sort((a, b) => {
    const ai = data.keyGoals.findIndex(g => g.id === a.goalId);
    const bi = data.keyGoals.findIndex(g => g.id === b.goalId);
    return (ai === -1 ? 999 : ai) - (bi === -1 ? 999 : bi);
  });
}

function buildOutSpans(sortedOut, keyGoals) {
  const spans = []; let i = 0;
  while (i < sortedOut.length) {
    const gid = sortedOut[i].goalId || null;
    const goal = gid ? keyGoals.find(g => g.id === gid) : null;
    let count = 0;
    while (i < sortedOut.length && sortedOut[i].goalId === gid) { count++; i++; }
    spans.push({ goal, gid, count });
  }
  return spans;
}

function applyUpdate(obj, path, value) {
  const [key, ...rest] = path;
  if (rest.length === 0) return { ...obj, [key]: value };
  const child = obj[key];
  if (Array.isArray(child)) {
    const [id, ...deepRest] = rest;
    return { ...obj, [key]: child.map(item => item.id === id ? applyUpdate(item, deepRest, value) : item) };
  }
  return { ...obj, [key]: applyUpdate(child || {}, rest, value) };
}

// ── INIT with Results code schema ─────────────────────────────────────────────
const INIT = {
  mapNum: '01', orgName: 'Your Organization Name', logo: null,
  layoutOutcomesFirst: false, layoutGroupByType: false,
  mission: 'Enter your mission statement - the fundamental purpose your organization serves.',
  vision: 'Enter your vision - what you aspire to become or achieve.',
  values: 'Integrity | Excellence | Service | Innovation | Collaboration',
  keyGoals: [
    { id: 'g1', name: 'Strategic Goal One', colorIdx: 0 },
    { id: 'g2', name: 'Strategic Goal Two', colorIdx: 1 },
    { id: 'g3', name: 'Strategic Goal Three', colorIdx: 2 },
  ],
  operatingProcesses: [
    { id: 'op1', code: '01OP01', name: 'First Operating Process',  goalId: 'g1', processOwner: 'Owner Name', subProcesses: ['Doing activity A', 'Doing activity B', 'Doing activity C'], processMeasures: [mkMeasure('Timeliness metric', '01OP01A')] },
    { id: 'op2', code: '01OP02', name: 'Second Operating Process', goalId: 'g1', processOwner: 'Owner Name', subProcesses: ['Doing activity A', 'Doing activity B'],                   processMeasures: [mkMeasure('Volume metric',    '01OP02A')] },
    { id: 'op3', code: '01OP03', name: 'Third Operating Process',  goalId: 'g2', processOwner: 'Owner Name', subProcesses: ['Doing activity A', 'Doing activity B'],                   processMeasures: [mkMeasure('Quality metric',   '01OP03A')] },
    { id: 'op4', code: '01OP04', name: 'Fourth Operating Process', goalId: 'g3', processOwner: 'Owner Name', subProcesses: ['Doing activity A', 'Doing activity B'],                   processMeasures: [mkMeasure('Accuracy metric',  '01OP04A')] },
  ],
  supportingProcesses: [
    { id: 'sp1', code: '01SP01', name: 'Leadership & Strategy', goalId: 'g1', processOwner: 'Owner Name', subProcesses: ['Setting direction', 'Managing performance', 'Driving improvement'], processMeasures: [mkMeasure('Engagement score', '01SP01A')] },
    { id: 'sp2', code: '01SP02', name: 'Finance & Operations',  goalId: 'g2', processOwner: 'Owner Name', subProcesses: ['Managing budget', 'Managing procurement', 'Financial reporting'],  processMeasures: [mkMeasure('Budget variance',  '01SP02A')] },
    { id: 'sp3', code: '01SP03', name: 'People & Culture',      goalId: 'g3', processOwner: 'Owner Name', subProcesses: ['Recruiting talent', 'Developing staff', 'Retaining people'],       processMeasures: [mkMeasure('Time to fill',     '01SP03A')] },
  ],
  outcomes: [
    { id: 'o1', code: '01OM01', name: 'Outcome One',   goalId: 'g1', owner: 'Owner Name', measures: [mkMeasure('Measure A', '01OM01A'), mkMeasure('Measure B', '01OM01B')] },
    { id: 'o2', code: '01OM02', name: 'Outcome Two',   goalId: 'g1', owner: 'Owner Name', measures: [mkMeasure('Measure C', '01OM02A')] },
    { id: 'o3', code: '01OM03', name: 'Outcome Three', goalId: 'g2', owner: 'Owner Name', measures: [mkMeasure('Measure D', '01OM03A'), mkMeasure('Measure E', '01OM03B')] },
    { id: 'o4', code: '01OM04', name: 'Outcome Four',  goalId: 'g3', owner: 'Owner Name', measures: [mkMeasure('Measure F', '01OM04A')] },
  ],
};

// Truly empty template — used by Reset
const BLANK = {
  mapNum: '01', orgName: '', logo: null,
  layoutOutcomesFirst: false, layoutGroupByType: false,
  mission: '', vision: '', values: '',
  keyGoals: [], operatingProcesses: [], supportingProcesses: [], outcomes: [],
};

// ── UI Primitives ─────────────────────────────────────────────────────────────
const iStyle = { width: '100%', padding: '7px 10px', border: '0.5px solid var(--color-border-secondary)', borderRadius: 6, fontSize: 13, fontFamily: 'inherit', background: 'var(--color-background-primary)', color: 'var(--color-text-primary)', boxSizing: 'border-box' };

function Field({ label, hint, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ fontSize: 11, fontWeight: 500, color: 'var(--color-text-secondary)', textTransform: 'uppercase', letterSpacing: 0.4, marginBottom: hint ? 3 : 5 }}>{label}</div>
      {hint && <div style={{ fontSize: 11, color: 'var(--color-text-tertiary)', marginBottom: 5, lineHeight: 1.4 }}>{hint}</div>}
      {children}
    </div>
  );
}

function Btn({ onClick, children, v = 'default', style: s = {}, title, disabled }) {
  const styles = {
    default: { bg: 'transparent', color: 'var(--color-text-secondary)', border: '0.5px solid var(--color-border-secondary)' },
    primary: { bg: '#1b3d6e', color: '#fff', border: 'none' },
    success: { bg: '#1a5c38', color: '#fff', border: 'none' },
    danger:  { bg: 'transparent', color: 'var(--color-text-danger)', border: '0.5px solid var(--color-border-danger)' },
    ghost:   { bg: 'rgba(255,255,255,0.18)', color: '#fff', border: '0.5px solid rgba(255,255,255,0.35)' },
    del:     { bg: 'rgba(180,40,40,0.55)', color: '#fff', border: 'none' },
  };
  const st = styles[v] || styles.default;
  return <button title={title} disabled={disabled} onClick={onClick} style={{ background: st.bg, color: st.color, border: st.border, padding: '5px 10px', borderRadius: 5, fontSize: 12, cursor: disabled ? 'default' : 'pointer', opacity: disabled ? 0.5 : 1, fontFamily: 'inherit', ...s }}>{children}</button>;
}

// ── Foundation Form ───────────────────────────────────────────────────────────
function FoundationForm({ data, setData }) {
  const f = k => v => setData(d => ({ ...d, [k]: v }));
  const logoRef = useRef(null);
  const handleLogo = e => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => setData(d => ({ ...d, logo: ev.target.result }));
    reader.readAsDataURL(file);
  };
  return (
    <div style={{ maxWidth: 560 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '120px 1fr', gap: '0 20px' }}>
        <Field label="Map number" hint="2-digit prefix used in all codes (01-99)">
          <input value={data.mapNum || '01'} onChange={e => f('mapNum')(e.target.value.replace(/\D/g, '').slice(0,2) || '01')} style={{ ...iStyle, maxWidth: 80 }} maxLength={2} />
        </Field>
        <Field label="Organization name">
          <input value={data.orgName} onChange={e => f('orgName')(e.target.value)} style={iStyle} />
        </Field>
      </div>
      <Field label="Logo" hint="PNG or JPG - shown in the map header">
        <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
          <Btn onClick={() => logoRef.current?.click()}>Upload logo</Btn>
          {data.logo && <Btn onClick={() => setData(d => ({ ...d, logo: null }))} v="danger">Remove</Btn>}
          <input ref={logoRef} type="file" accept="image/*" style={{ display: 'none' }} onChange={handleLogo} />
        </div>
        {data.logo && <img src={data.logo} alt="logo preview" style={{ maxHeight: 56, maxWidth: 220, marginTop: 8, objectFit: 'contain', display: 'block' }} />}
      </Field>
      <Field label="Mission" hint="Why your organization exists"><textarea value={data.mission} onChange={e => f('mission')(e.target.value)} rows={3} style={{ ...iStyle, resize: 'vertical' }} /></Field>
      <Field label="Vision"  hint="What you aspire to become"><textarea  value={data.vision}  onChange={e => f('vision')(e.target.value)}  rows={3} style={{ ...iStyle, resize: 'vertical' }} /></Field>
      <Field label="Values"  hint="Guiding principles - separate with | or commas"><textarea value={data.values} onChange={e => f('values')(e.target.value)} rows={2} style={{ ...iStyle, resize: 'vertical' }} /></Field>

      {/* ── Layout options ── */}
      <div style={{ marginTop: 8, padding: 14, border: '0.5px solid var(--color-border-secondary)', borderRadius: 6, background: 'var(--color-background-secondary)' }}>
        <div style={{ fontSize: 11, fontWeight: 600, textTransform: 'uppercase', letterSpacing: 0.4, color: 'var(--color-text-secondary)', marginBottom: 12 }}>Map Layout Options</div>
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0 28px' }}>
          <Field label="Outcomes position">
            {[['below', 'Below core processes (default)'], ['above', 'Above core processes']].map(([val, label]) => (
              <label key={val} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, fontSize: 12, cursor: 'pointer', color: 'var(--color-text-primary)' }}>
                <input type="radio" name="outcomesPos" value={val}
                  checked={(data.layoutOutcomesFirst ? 'above' : 'below') === val}
                  onChange={() => setData(d => ({ ...d, layoutOutcomesFirst: val === 'above' }))} />
                {label}
              </label>
            ))}
          </Field>
          <Field label="Process grouping">
            {[['goal', 'By key goal (goal-first)'], ['type', 'By type (Operating / Supporting)']].map(([val, label]) => (
              <label key={val} style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8, fontSize: 12, cursor: 'pointer', color: 'var(--color-text-primary)' }}>
                <input type="radio" name="procGroup" value={val}
                  checked={(data.layoutGroupByType ? 'type' : 'goal') === val}
                  onChange={() => setData(d => ({ ...d, layoutGroupByType: val === 'type' }))} />
                {label}
              </label>
            ))}
          </Field>
        </div>
      </div>
    </div>
  );
}

// ── Goals Form ────────────────────────────────────────────────────────────────
function GoalsForm({ data, setData }) {
  const setG = (id, k, v) => setData(d => ({ ...d, keyGoals: d.keyGoals.map(g => g.id === id ? { ...g, [k]: v } : g) }));
  const addG  = () => setData(d => ({ ...d, keyGoals: [...d.keyGoals, { id: uid(), name: 'New goal', colorIdx: d.keyGoals.length % PALETTE.length }] }));
  const remG  = id => setData(d => ({
    ...d, keyGoals: d.keyGoals.filter(g => g.id !== id),
    operatingProcesses: d.operatingProcesses.map(p => p.goalId === id ? { ...p, goalId: null } : p),
    supportingProcesses: d.supportingProcesses.map(p => p.goalId === id ? { ...p, goalId: null } : p),
    outcomes: d.outcomes.map(o => o.goalId === id ? { ...o, goalId: null } : o),
  }));
  return (
    <div style={{ maxWidth: 560 }}>
      <div style={{ fontSize: 12, color: 'var(--color-text-secondary)', marginBottom: 16, lineHeight: 1.5 }}>Strategic pillars that group processes and outcomes. Each becomes a colored column header.</div>
      {data.keyGoals.map(goal => {
        const color = pc(goal.colorIdx);
        return (
          <div key={goal.id} style={{ border: '0.5px solid var(--color-border-secondary)', borderLeft: `4px solid ${color.bg}`, borderRadius: 6, padding: 14, marginBottom: 10, background: 'var(--color-background-primary)' }}>
            <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
              <input value={goal.name} onChange={e => setG(goal.id, 'name', e.target.value)} style={{ ...iStyle, flex: 1 }} />
              <Btn onClick={() => remG(goal.id)} v="danger">✕</Btn>
            </div>
            <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
              <span style={{ fontSize: 11, color: 'var(--color-text-secondary)' }}>Color:</span>
              {PALETTE.map((p, i) => <button key={i} onClick={() => setG(goal.id, 'colorIdx', i)} style={{ width: 20, height: 20, borderRadius: 4, background: p.bg, border: goal.colorIdx === i ? '2px solid var(--color-text-primary)' : '2px solid transparent', cursor: 'pointer', padding: 0 }} />)}
            </div>
          </div>
        );
      })}
      <Btn onClick={addG} v="primary">+ Add goal</Btn>
    </div>
  );
}

// ── Process Form ──────────────────────────────────────────────────────────────
function ProcessForm({ data, setData, type, title, desc, prefix, typeCode }) {
  const procs = data[type];
  const setF = (id, k, v) => setData(d => ({ ...d, [type]: d[type].map(p => p.id === id ? { ...p, [k]: v } : p) }));
  const setSub = (id, i, v) => setData(d => ({ ...d, [type]: d[type].map(p => { if (p.id !== id) return p; const a = [...p.subProcesses]; a[i] = v; return { ...p, subProcesses: a }; }) }));
  const addSub = id => setData(d => ({ ...d, [type]: d[type].map(p => p.id === id ? { ...p, subProcesses: [...p.subProcesses, ''] } : p) }));
  const remSub = (id, i) => setData(d => ({ ...d, [type]: d[type].map(p => p.id === id ? { ...p, subProcesses: p.subProcesses.filter((_, j) => j !== i) } : p) }));

  const setMeasure = (pid, mid, field, v) => setData(d => ({ ...d, [type]: d[type].map(p => p.id !== pid ? p : { ...p, processMeasures: p.processMeasures.map(m => m.id === mid ? { ...m, [field]: v } : m) }) }));
  const addMeasure = pid => setData(d => {
    const prcs = d[type]; const procIdx = prcs.findIndex(p => p.id === pid); const proc = prcs[procIdx];
    const code = measCodeStr(d.mapNum, typeCode, procIdx + 1, proc.processMeasures.length);
    return { ...d, [type]: prcs.map(p => p.id !== pid ? p : { ...p, processMeasures: [...p.processMeasures, mkMeasure('', code)] }) };
  });
  const remMeasure = (pid, mid) => setData(d => ({ ...d, [type]: d[type].map(p => p.id !== pid ? p : { ...p, processMeasures: p.processMeasures.filter(m => m.id !== mid) }) }));

  const addP = () => setData(d => {
    const n = d[type].length + 1;
    const code = procCodeStr(d.mapNum, typeCode, n);
    const firstMeasCode = measCodeStr(d.mapNum, typeCode, n, 0);
    return { ...d, [type]: [...d[type], { id: uid(), code, name: 'New process', goalId: d.keyGoals[0]?.id ?? null, processOwner: '', subProcesses: ['Sub-process'], processMeasures: [mkMeasure('', firstMeasCode)] }] };
  });
  const remP  = id  => setData(d => ({ ...d, [type]: d[type].filter(p => p.id !== id) }));
  const moveP = (id, dir) => setData(d => { const a = [...d[type]]; const i = a.findIndex(p => p.id === id); const j = i + dir; if (j < 0 || j >= a.length) return d; [a[i], a[j]] = [a[j], a[i]]; return { ...d, [type]: a }; });

  const secBg = { background: 'var(--color-background-secondary)', borderRadius: 5, padding: '10px 12px', marginBottom: 10 };

  return (
    <div style={{ maxWidth: 720 }}>
      <div style={{ fontSize: 12, color: 'var(--color-text-secondary)', marginBottom: 16, lineHeight: 1.5 }}>{desc}</div>
      {procs.map(proc => {
        const g = data.keyGoals.find(g => g.id === proc.goalId);
        const color = g ? pc(g.colorIdx) : { bg: '#888', light: '#eee' };
        return (
          <div key={proc.id} style={{ border: '0.5px solid var(--color-border-secondary)', borderRadius: 6, marginBottom: 16, overflow: 'hidden', background: 'var(--color-background-primary)' }}>
            <div style={{ background: color.bg, padding: '7px 10px', display: 'flex', gap: 6, alignItems: 'center' }}>
              <input value={proc.code} onChange={e => setF(proc.id, 'code', e.target.value)} style={{ ...iStyle, width: 90, flexShrink: 0, background: 'rgba(255,255,255,0.15)', color: '#fff', border: '0.5px solid rgba(255,255,255,0.3)', fontSize: 12 }} placeholder="Code" />
              <input value={proc.name} onChange={e => setF(proc.id, 'name', e.target.value)} style={{ ...iStyle, flex: 1, background: 'rgba(255,255,255,0.15)', color: '#fff', border: '0.5px solid rgba(255,255,255,0.3)' }} />
              <Btn onClick={() => moveP(proc.id, -1)} v="ghost" title="Move up">↑</Btn>
              <Btn onClick={() => moveP(proc.id,  1)} v="ghost" title="Move down">↓</Btn>
              <Btn onClick={() => remP(proc.id)} v="del" title="Remove">✕</Btn>
            </div>
            <div style={{ padding: 14 }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0 20px' }}>
                <Field label="Assigned to goal">
                  <select value={proc.goalId || ''} onChange={e => setF(proc.id, 'goalId', e.target.value || null)} style={iStyle}>
                    <option value="">— Unassigned —</option>
                    {data.keyGoals.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}
                  </select>
                </Field>
                <Field label="Process owner">
                  <input value={proc.processOwner} onChange={e => setF(proc.id, 'processOwner', e.target.value)} style={iStyle} placeholder="Name..." />
                </Field>
              </div>
              <div style={secBg}>
                <div style={{ fontSize: 11, fontWeight: 600, color: 'var(--color-text-secondary)', textTransform: 'uppercase', letterSpacing: 0.4, marginBottom: 8 }}>Sub-processes / activities</div>
                {proc.subProcesses.map((s, i) => (
                  <div key={i} style={{ display: 'flex', gap: 5, marginBottom: 5 }}>
                    <input value={s} onChange={e => setSub(proc.id, i, e.target.value)} style={{ ...iStyle, flex: 1, fontSize: 12 }} placeholder={`Activity ${i + 1}`} />
                    <Btn onClick={() => remSub(proc.id, i)}>✕</Btn>
                  </div>
                ))}
                <Btn onClick={() => addSub(proc.id)}>+ Activity</Btn>
              </div>
              <div style={secBg}>
                <div style={{ fontSize: 11, fontWeight: 600, color: 'var(--color-text-secondary)', textTransform: 'uppercase', letterSpacing: 0.4, marginBottom: 8 }}>Process measures</div>
                {proc.processMeasures.map(m => (
                  <div key={m.id} style={{ border: '0.5px solid var(--color-border-secondary)', borderRadius: 5, padding: '8px 10px', marginBottom: 8, background: 'var(--color-background-primary)' }}>
                    <div style={{ display: 'flex', gap: 6, marginBottom: 6 }}>
                      <input value={m.code} onChange={e => setMeasure(proc.id, m.id, 'code', e.target.value)} style={{ ...iStyle, width: 100, flexShrink: 0, fontSize: 12 }} placeholder="Code" />
                      <input value={m.text} onChange={e => setMeasure(proc.id, m.id, 'text', e.target.value)} style={{ ...iStyle, flex: 1, fontSize: 12 }} placeholder="Measure description" />
                      <Btn onClick={() => remMeasure(proc.id, m.id)}>✕</Btn>
                    </div>
                    <input value={m.measureOwner} onChange={e => setMeasure(proc.id, m.id, 'measureOwner', e.target.value)} style={{ ...iStyle, fontSize: 12 }} placeholder="Measure owner..." />
                  </div>
                ))}
                <Btn onClick={() => addMeasure(proc.id)}>+ Measure</Btn>
              </div>
            </div>
          </div>
        );
      })}
      <Btn onClick={addP} v="primary">+ Add {title}</Btn>
    </div>
  );
}

// ── Outcomes Form ─────────────────────────────────────────────────────────────
function OutcomesForm({ data, setData }) {
  const setO  = (id, k, v) => setData(d => ({ ...d, outcomes: d.outcomes.map(o => o.id === id ? { ...o, [k]: v } : o) }));
  const setMeasure = (oid, mid, field, v) => setData(d => ({ ...d, outcomes: d.outcomes.map(o => o.id !== oid ? o : { ...o, measures: o.measures.map(m => m.id === mid ? { ...m, [field]: v } : m) }) }));
  const addMeasure = oid => setData(d => {
    const outs = d.outcomes; const oIdx = outs.findIndex(o => o.id === oid); const o = outs[oIdx];
    const code = outMeasCodeStr(d.mapNum, oIdx + 1, o.measures.length);
    return { ...d, outcomes: outs.map(x => x.id !== oid ? x : { ...x, measures: [...x.measures, mkMeasure('', code)] }) };
  });
  const remMeasure = (oid, mid) => setData(d => ({ ...d, outcomes: d.outcomes.map(o => o.id !== oid ? o : { ...o, measures: o.measures.filter(m => m.id !== mid) }) }));
  const addO = () => setData(d => {
    const n = d.outcomes.length + 1;
    const code = outcomeCodeStr(d.mapNum, n);
    const firstMeasCode = outMeasCodeStr(d.mapNum, n, 0);
    return { ...d, outcomes: [...d.outcomes, { id: uid(), code, name: `Outcome ${n}`, goalId: d.keyGoals[0]?.id ?? null, owner: '', measures: [mkMeasure('', firstMeasCode)] }] };
  });
  const remO = id => setData(d => ({ ...d, outcomes: d.outcomes.filter(o => o.id !== id) }));
  const secBg = { background: 'var(--color-background-secondary)', borderRadius: 5, padding: '10px 12px', marginBottom: 10 };

  return (
    <div style={{ maxWidth: 720 }}>
      <div style={{ fontSize: 12, color: 'var(--color-text-secondary)', marginBottom: 16, lineHeight: 1.5 }}>Strategic outcomes — what success looks like, grouped by goal.</div>
      {data.outcomes.map(o => {
        const g = data.keyGoals.find(g => g.id === o.goalId);
        const color = g ? pc(g.colorIdx) : { bg: '#888', light: '#eee' };
        return (
          <div key={o.id} style={{ border: '0.5px solid var(--color-border-secondary)', borderRadius: 6, marginBottom: 14, overflow: 'hidden', background: 'var(--color-background-primary)' }}>
            <div style={{ background: color.bg, padding: '7px 10px', display: 'flex', gap: 6, alignItems: 'center' }}>
              <input value={o.code} onChange={e => setO(o.id, 'code', e.target.value)} style={{ ...iStyle, width: 100, flexShrink: 0, background: 'rgba(255,255,255,0.15)', color: '#fff', border: '0.5px solid rgba(255,255,255,0.3)', fontSize: 12 }} placeholder="Code" />
              <input value={o.name} onChange={e => setO(o.id, 'name', e.target.value)} style={{ ...iStyle, flex: 1, background: 'rgba(255,255,255,0.15)', color: '#fff', border: '0.5px solid rgba(255,255,255,0.3)' }} />
              <Btn onClick={() => remO(o.id)} v="del">✕</Btn>
            </div>
            <div style={{ padding: 14 }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '0 20px' }}>
                <Field label="Assigned to goal">
                  <select value={o.goalId || ''} onChange={e => setO(o.id, 'goalId', e.target.value || null)} style={iStyle}>
                    <option value="">— Unassigned —</option>
                    {data.keyGoals.map(g => <option key={g.id} value={g.id}>{g.name}</option>)}
                  </select>
                </Field>
                <Field label="Outcome owner">
                  <input value={o.owner} onChange={e => setO(o.id, 'owner', e.target.value)} style={iStyle} placeholder="Name..." />
                </Field>
              </div>
              <div style={secBg}>
                <div style={{ fontSize: 11, fontWeight: 600, color: 'var(--color-text-secondary)', textTransform: 'uppercase', letterSpacing: 0.4, marginBottom: 8 }}>Outcome measures</div>
                {o.measures.map(m => (
                  <div key={m.id} style={{ border: '0.5px solid var(--color-border-secondary)', borderRadius: 5, padding: '8px 10px', marginBottom: 8, background: 'var(--color-background-primary)' }}>
                    <div style={{ display: 'flex', gap: 6, marginBottom: 6 }}>
                      <input value={m.code} onChange={e => setMeasure(o.id, m.id, 'code', e.target.value)} style={{ ...iStyle, width: 100, flexShrink: 0, fontSize: 12 }} placeholder="Code" />
                      <input value={m.text} onChange={e => setMeasure(o.id, m.id, 'text', e.target.value)} style={{ ...iStyle, flex: 1, fontSize: 12 }} placeholder="Measure description" />
                      <Btn onClick={() => remMeasure(o.id, m.id)}>✕</Btn>
                    </div>
                    <input value={m.measureOwner} onChange={e => setMeasure(o.id, m.id, 'measureOwner', e.target.value)} style={{ ...iStyle, fontSize: 12 }} placeholder="Measure owner..." />
                  </div>
                ))}
                <Btn onClick={() => addMeasure(o.id)}>+ Measure</Btn>
              </div>
            </div>
          </div>
        );
      })}
      <Btn onClick={addO} v="primary">+ Add outcome</Btn>
    </div>
  );
}

// ── Inline Editable Cell ──────────────────────────────────────────────────────
function EC({ path, value, editPath, setEditPath, setData, dark }) {
  const key = JSON.stringify(path);
  const editing = editPath === key;
  const isList = Array.isArray(value);
  const display = isList ? value.map(v => '- ' + v).join('\n') : (value || '');
  const raw     = isList ? value.join('\n')                    : (value || '');
  if (editing) {
    return (
      <textarea autoFocus defaultValue={raw} rows={isList ? Math.max(value.length, 2) : 2}
        style={{ width: '100%', fontSize: 'inherit', fontFamily: 'inherit', background: dark ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.04)', color: 'inherit', border: '1px solid rgba(128,128,128,0.4)', borderRadius: 3, padding: '3px 5px', resize: 'vertical', boxSizing: 'border-box' }}
        onBlur={e => { const raw = e.target.value; const val = isList ? raw.split('\n').map(s => s.trim()).filter(Boolean) : raw; setData(d => applyUpdate(d, path, val)); setEditPath(null); }}
        onKeyDown={e => { if (e.key === 'Escape') setEditPath(null); if (!isList && e.key === 'Enter') { e.preventDefault(); e.target.blur(); } }}
      />
    );
  }
  return (
    <div onClick={() => setEditPath(key)} title="Click to edit" style={{ cursor: 'pointer', minHeight: 18, whiteSpace: 'pre-line', lineHeight: 1.55, borderRadius: 2, padding: '1px 2px' }}>
      {display || <span style={{ opacity: 0.35, fontStyle: 'italic' }}>—</span>}
    </div>
  );
}

// ── Map Table ─────────────────────────────────────────────────────────────────
function MapTable({ data, editPath, setEditPath, setData, mapRef }) {
  const { keyGoals } = data;
  const groupByType   = data.layoutGroupByType   ?? false;
  const outcomesFirst = data.layoutOutcomesFirst ?? false;

  // ── Build column order ────────────────────────────────────────────────────
  const goalOrder = Object.fromEntries(keyGoals.map((g, i) => [g.id, i]));
  const byGoal    = (a, b) => (goalOrder[a.goalId] ?? 999) - (goalOrder[b.goalId] ?? 999);

  let allProcs = [], goalHeaderCells = [], opCount = 0, spCount = 0;
  if (groupByType) {
    // Type-first: all ops (sorted by goal) then all sps (sorted by goal)
    const sOp = [...data.operatingProcesses].sort(byGoal);
    const sSp = [...data.supportingProcesses].sort(byGoal);
    allProcs = [
      ...sOp.map(p => ({ proc: p, side: 'op', goalId: p.goalId })),
      ...sSp.map(p => ({ proc: p, side: 'sp', goalId: p.goalId })),
    ];
    opCount = Math.max(sOp.length, 1);
    spCount = Math.max(sSp.length, 1);
    const opSpans = buildGoalSpans(sOp.map(p => ({ goalId: p.goalId })), keyGoals);
    const spSpans = buildGoalSpans(sSp.map(p => ({ goalId: p.goalId })), keyGoals);
    goalHeaderCells = [...opSpans, ...spSpans]; // goals may appear in both sections
  } else {
    // Goal-first: for each goal, ops then sps
    const seen = new Set();
    for (const goal of keyGoals) {
      for (const p of data.operatingProcesses)  if (p.goalId === goal.id) { allProcs.push({ proc: p, side: 'op', goalId: goal.id }); seen.add(p.id); }
      for (const p of data.supportingProcesses) if (p.goalId === goal.id) { allProcs.push({ proc: p, side: 'sp', goalId: goal.id }); seen.add(p.id); }
    }
    for (const p of data.operatingProcesses)  if (!seen.has(p.id)) { allProcs.push({ proc: p, side: 'op', goalId: null }); seen.add(p.id); }
    for (const p of data.supportingProcesses) if (!seen.has(p.id)) { allProcs.push({ proc: p, side: 'sp', goalId: null }); seen.add(p.id); }
    goalHeaderCells = buildGoalSpans(allProcs.map(x => ({ goalId: x.goalId })), keyGoals);
  }

  const sortedOut  = sortedOutcomes(data);
  const outSpans   = buildOutSpans(sortedOut, keyGoals);
  const procCols   = allProcs.length;
  const totalCols  = Math.max(procCols, 3);
  const fillerCols = totalCols - procCols;
  const outCount   = Math.max(sortedOut.length, 1);

  const SB = 28, COL = 162;
  const totalW  = SB + totalCols * COL;
  const outColW = Math.max(Math.floor((totalW - SB) / outCount), 120);

  // When outcomes are first the header must align with the outcomes table directly below it.
  // When processes are first it aligns with the process table.
  const hdrAlignOutcomes = outcomesFirst && sortedOut.length > 0;
  const hdrCols   = hdrAlignOutcomes ? Math.max(outCount, 3) : totalCols;
  const hdrColW   = hdrAlignOutcomes ? outColW : COL;
  const hdrW      = SB + hdrCols * hdrColW;
  const hdrFiller = hdrCols - (hdrAlignOutcomes ? outCount : procCols);
  const hdrGoals  = hdrAlignOutcomes ? outSpans : goalHeaderCells;

  const c1 = Math.ceil(hdrCols / 3);
  const c2 = Math.ceil((hdrCols - c1) / 2);
  const c3 = hdrCols - c1 - c2;

  const B    = '0.5px solid #bbb';
  const cell = (ex = {}) => ({ border: B, padding: '5px 7px', verticalAlign: 'top', fontSize: 10, lineHeight: 1.55, ...ex });
  const hdr  = (bg, fg = '#fff', ex = {}) => ({ border: B, background: bg, color: fg, fontWeight: 500, textAlign: 'center', verticalAlign: 'middle', fontSize: 11, padding: '7px 6px', ...ex });
  const lbl  = text => <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 0.6, color: '#777', marginBottom: 2 }}>{text}</div>;

  const goalColor = gid => { const g = keyGoals.find(g => g.id === gid); return g ? pc(g.colorIdx) : { bg: '#666', text: '#fff', light: '#ddd' }; };
  const opById    = Object.fromEntries(data.operatingProcesses.map(p => [p.id, true]));
  const procType  = ({ proc }) => opById[proc.id] ? 'operatingProcesses' : 'supportingProcesses';
  const ecProps   = { editPath, setEditPath, setData };
  const OP_BG = '#2c3e50', SP_BG = '#445566';

  const SbCell = ({ bg, label, rowSpan }) => (
    <td rowSpan={rowSpan} style={{ border: B, background: bg, width: SB, minWidth: SB, maxWidth: SB, padding: 0, verticalAlign: 'middle', textAlign: 'center' }}>
      <span style={{ writingMode: 'vertical-rl', transform: 'rotate(180deg)', display: 'inline-block', color: '#fff', fontSize: 10, fontWeight: 700, textTransform: 'uppercase', letterSpacing: 1, whiteSpace: 'nowrap' }}>{label}</span>
    </td>
  );

  // colGroup for the process table
  const colGroup = (
    <colgroup>
      <col style={{ width: SB }} />
      {Array(totalCols).fill(0).map((_, i) => <col key={i} style={{ width: COL }} />)}
    </colgroup>
  );

  // ── Header table: always Foundations; Key Goals only when processes come first ──
  // When outcomes are first, Key Goals above them is redundant (goal context already
  // shown in the outcomes table's own goal-spans row). Key Goals then appears at the
  // top of the process table instead.
  const headerTable = (
    <table className="map-table" style={{ borderCollapse: 'collapse', width: hdrW, tableLayout: 'fixed' }}>
      <colgroup>
        <col style={{ width: SB }} />
        {Array(hdrCols).fill(0).map((_, i) => <col key={i} style={{ width: hdrColW }} />)}
      </colgroup>
      <tbody>
        {/* Rows 0-1: Foundations (always) */}
        <tr>
          <SbCell bg="#0d1b2a" label="Foundations" rowSpan={2} />
          <td colSpan={hdrCols} style={hdr('#0d1b2a', '#fff', { fontSize: 16, padding: '8px 12px', textAlign: 'left' })}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              {data.logo && <img src={data.logo} alt="logo" style={{ maxHeight: 36, maxWidth: 100, objectFit: 'contain', flexShrink: 0 }} />}
              <EC path={['orgName']} value={data.orgName} dark {...ecProps} />
            </div>
          </td>
        </tr>
        <tr>
          <td colSpan={c1} style={cell({ background: '#e8f0fb', fontSize: 16, lineHeight: 1.5, padding: '12px 14px' })}>{lbl('Mission')}<EC path={['mission']} value={data.mission} {...ecProps} /></td>
          <td colSpan={c2} style={cell({ background: '#1b3d6e', color: '#fff', textAlign: 'center', fontSize: 16, lineHeight: 1.5, padding: '12px 14px' })}>{lbl('Vision')}<EC path={['vision']} value={data.vision} dark {...ecProps} /></td>
          <td colSpan={c3} style={cell({ background: '#1e1e1e', color: '#fff', textAlign: 'center', fontSize: 16, lineHeight: 1.5, padding: '12px 14px' })}>{lbl('Values')}<EC path={['values']} value={data.values} dark {...ecProps} /></td>
        </tr>
        {/* Row 2: Key Goals — minHeight ensures rotated label fits at 10px font */}
        <tr style={{ minHeight: 90 }}>
          <SbCell bg="#1b3d6e" label="Key Goals" rowSpan={1} />
          {hdrGoals.length > 0
            ? hdrGoals.map(({ goal, gid, count }, i) => {
                const color = goal ? pc(goal.colorIdx) : { bg: '#666', text: '#fff' };
                return (
                  <td key={(gid ?? 'g') + i} colSpan={count} style={hdr(color.bg, color.text, { fontSize: 14, padding: '9px 8px', minHeight: 90, height: 90 })}>
                    {goal ? <EC path={['keyGoals', goal.id, 'name']} value={goal.name} dark {...ecProps} /> : <span style={{ opacity: 0.6, fontStyle: 'italic' }}>Unassigned</span>}
                  </td>
                );
              })
            : <td colSpan={hdrCols} style={hdr('#888', '#fff', { fontStyle: 'italic', fontSize: 11 })}>Add key goals in the Editor</td>
          }
          {hdrFiller > 0 && hdrGoals.length > 0 && <td colSpan={hdrFiller} style={hdr('#aaa', '#fff')} />}
        </tr>
      </tbody>
    </table>
  );

  // ── Outcomes table ─────────────────────────────────────────────────────────
  const outColGroup = (
    <colgroup>
      <col style={{ width: SB }} />
      {sortedOut.map((_, i) => <col key={i} style={{ width: outColW }} />)}
    </colgroup>
  );

  const outcomesTable = sortedOut.length > 0 ? (
    <table className="map-table" style={{ borderCollapse: 'collapse', width: SB + outCount * outColW, tableLayout: 'fixed', borderTop: '2px solid #888', marginTop: 3 }}>
      {outColGroup}
      <tbody>
        {/* Goal spans row — only when outcomes are at the bottom.
            SbCell rowSpan covers this row + the 3 content rows below = 4.
            When outcomes are at top there is no goal spans row so rowSpan = 3. */}
        {!outcomesFirst ? (
          <tr>
            <SbCell bg="#555" label="Outcomes" rowSpan={4} />
            {outSpans.map(({ goal, gid, count }, i) => {
              const color = goal ? pc(goal.colorIdx) : { bg: '#666', text: '#fff' };
              return <td key={(gid ?? 'u') + i} colSpan={count} style={hdr(color.bg, color.text, { fontSize: 13, padding: '9px 8px' })}>{goal?.name ?? 'Unassigned'}</td>;
            })}
          </tr>
        ) : (
          <tr>
            <SbCell bg="#555" label="Outcomes" rowSpan={3} />
          </tr>
        )}
        <tr>
          {sortedOut.map(o => {
            const color = goalColor(o.goalId);
            return (
              <td key={o.id} style={cell({ background: color.light, textAlign: 'center', fontSize: 11 })}>
                {o.code && <div style={{ fontSize: 10, color: '#777', marginBottom: 1 }}>{o.code}</div>}
                <div style={{ fontWeight: 500 }}><EC path={['outcomes', o.id, 'name']} value={o.name} {...ecProps} /></div>
                {o.owner && <div style={{ fontSize: 10, fontStyle: 'italic', color: '#555', marginTop: 3, borderTop: '0.5px solid rgba(0,0,0,0.1)', paddingTop: 3 }}><EC path={['outcomes', o.id, 'owner']} value={o.owner} {...ecProps} /></div>}
              </td>
            );
          })}
        </tr>
        <tr>
          {sortedOut.map(o => {
            const color = goalColor(o.goalId);
            return (
              <td key={o.id} style={cell({ background: color.light, fontSize: 10 })}>
                {lbl('Outcome Measures')}
                {o.measures.length === 0
                  ? <span style={{ opacity: 0.35, fontStyle: 'italic' }}>—</span>
                  : o.measures.map((m, mi) => (
                    <div key={m.id} style={{ marginBottom: mi < o.measures.length - 1 ? 5 : 0 }}>
                      <div>{m.code && <span style={{ fontWeight: 700, fontSize: 10, marginRight: 3 }}>{m.code}</span>}{m.text || <span style={{ opacity: 0.35, fontStyle: 'italic' }}>—</span>}</div>
                      {m.measureOwner && <div style={{ fontSize: 10, fontStyle: 'italic', color: '#555', marginTop: 1 }}>{m.measureOwner}</div>}
                    </div>
                  ))}
              </td>
            );
          })}
        </tr>
      </tbody>
    </table>
  ) : null;

  // ── Core Processes table ───────────────────────────────────────────────────
  const processTable = (
    <table className="map-table" style={{ borderCollapse: 'collapse', width: totalW, tableLayout: 'fixed', borderTop: '2px solid #888', marginTop: 3 }}>
      {colGroup}
      <tbody>
        {/* Goal spans row + sidebar — when processes are at bottom, SbCell spans this row too (rowSpan 5).
            When processes are at top, no goal spans row, SbCell rowSpan stays 4. */}
        {outcomesFirst ? (
          <tr>
            <SbCell bg="#3a3a3a" label="Core Processes" rowSpan={5} />
            {goalHeaderCells.length > 0
              ? goalHeaderCells.map(({ goal, gid, count }, i) => {
                  const color = goal ? pc(goal.colorIdx) : { bg: '#666', text: '#fff' };
                  return (
                    <td key={(gid ?? 'g') + i} colSpan={count} style={hdr(color.bg, color.text, { fontSize: 13, padding: '9px 8px' })}>
                      {goal?.name ?? 'Unassigned'}
                    </td>
                  );
                })
              : <td colSpan={totalCols} style={hdr('#888', '#fff', { fontStyle: 'italic', fontSize: 11 })}>Add key goals in the Editor</td>
            }
            {fillerCols > 0 && goalHeaderCells.length > 0 && <td colSpan={fillerCols} style={hdr('#aaa', '#fff')} />}
          </tr>
        ) : null}
        {/* Op/Sp section header — SbCell starts here when processes are at top (rowSpan 4) */}
        <tr>
          {!outcomesFirst && <SbCell bg="#3a3a3a" label="Core Processes" rowSpan={4} />}
          {groupByType ? (
            <>
              <td colSpan={opCount} style={hdr(OP_BG, '#fff', { fontSize: 10, letterSpacing: 0.8, textTransform: 'uppercase', padding: '5px 6px' })}>Operating Processes</td>
              <td colSpan={spCount} style={hdr(SP_BG, '#fff', { fontSize: 10, letterSpacing: 0.8, textTransform: 'uppercase', padding: '5px 6px' })}>Supporting Processes</td>
              {fillerCols > 0 && <td colSpan={fillerCols} style={hdr('#888', '#fff')} />}
            </>
          ) : (
            <>
              {allProcs.map(({ proc, side }) => (
                <td key={proc.id} style={hdr(side === 'op' ? OP_BG : SP_BG, '#fff', { fontSize: 10, letterSpacing: 0.8, textTransform: 'uppercase', padding: '4px' })}>
                  {side === 'op' ? 'Operating' : 'Supporting'}
                </td>
              ))}
              {fillerCols > 0 && <td colSpan={fillerCols} style={hdr('#888', '#fff', { fontSize: 10, padding: '4px' })} />}
            </>
          )}
        </tr>
        {/* Row 1: Process name + owner */}
        <tr>
          {allProcs.map(({ proc, goalId }) => {
            const color = goalColor(goalId);
            const type  = procType({ proc });
            return (
              <td key={proc.id} style={hdr(color.bg, color.text, { fontSize: 12, textAlign: 'left', verticalAlign: 'top', padding: '6px 7px' })}>
                {proc.code && <div style={{ fontSize: 10, opacity: 0.7, fontWeight: 400, letterSpacing: 0.3, marginBottom: 1 }}>{proc.code}</div>}
                <EC path={[type, proc.id, 'name']} value={proc.name} dark {...ecProps} />
                {proc.processOwner && (
                  <div style={{ fontSize: 10, fontStyle: 'italic', opacity: 0.8, marginTop: 3, borderTop: '0.5px solid rgba(255,255,255,0.25)', paddingTop: 3 }}>
                    <EC path={[type, proc.id, 'processOwner']} value={proc.processOwner} dark {...ecProps} />
                  </div>
                )}
              </td>
            );
          })}
          {fillerCols > 0 && <td colSpan={fillerCols} style={hdr('#888', '#fff', { padding: '6px 7px' })} />}
        </tr>
        {/* Row 2: Sub-processes */}
        <tr>
          {allProcs.map(({ proc }) => {
            const type = procType({ proc });
            return (
              <td key={proc.id} style={cell({ background: '#fff', fontSize: 10.5 })}>
                <EC path={[type, proc.id, 'subProcesses']} value={proc.subProcesses} {...ecProps} />
              </td>
            );
          })}
          {fillerCols > 0 && <td colSpan={fillerCols} style={cell({ background: '#fff', fontSize: 10.5 })} />}
        </tr>
        {/* Row 3: Process measures + owners */}
        <tr>
          {allProcs.map(({ proc, goalId }) => {
            const color = goalColor(goalId);
            return (
              <td key={proc.id} style={cell({ background: color.light, fontSize: 10 })}>
                {lbl('Process Measures')}
                {proc.processMeasures.length === 0
                  ? <span style={{ opacity: 0.35, fontStyle: 'italic' }}>—</span>
                  : proc.processMeasures.map((m, mi) => (
                    <div key={m.id} style={{ marginBottom: mi < proc.processMeasures.length - 1 ? 5 : 0 }}>
                      <div>{m.code && <span style={{ fontWeight: 700, fontSize: 10, marginRight: 3 }}>{m.code}</span>}{m.text || <span style={{ opacity: 0.35, fontStyle: 'italic' }}>—</span>}</div>
                      {m.measureOwner && <div style={{ fontSize: 10, fontStyle: 'italic', color: '#555', marginTop: 1 }}>{m.measureOwner}</div>}
                    </div>
                  ))}
              </td>
            );
          })}
          {fillerCols > 0 && <td colSpan={fillerCols} style={cell({ background: '#f5f5f5' })} />}
        </tr>
      </tbody>
    </table>
  );

  return (
    <div className="map-print-root" ref={mapRef}>
      {headerTable}
      {outcomesFirst  ? outcomesTable : null}
      {processTable}
      {!outcomesFirst ? outcomesTable : null}
    </div>
  );
}


// ── Storage helpers ───────────────────────────────────────────────────────────
const STORE_KEY = 'fundamentals_map_v1';

async function saveToStorage(data) {
  const payload = JSON.stringify(data);
  if (window.storage) {
    const result = await window.storage.set(STORE_KEY, payload);
    if (!result) throw new Error('storage.set returned null');
  } else {
    localStorage.setItem(STORE_KEY, payload);
  }
}

async function loadFromStorage() {
  if (window.storage) {
    const result = await window.storage.get(STORE_KEY);
    return result ? JSON.parse(result.value) : null;
  } else {
    const raw = localStorage.getItem(STORE_KEY);
    return raw ? JSON.parse(raw) : null;
  }
}

// ── Load PptxGenJS via script injection ───────────────────────────────────────

function loadScript(src) {
  return new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error('Failed to load: ' + src));
    document.head.appendChild(s);
  });
}

async function exportToPptx(data) {
  try {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pptxgenjs/3.12.0/pptxgen.bundle.js');
    const PG = window.PptxGenJS;
    if (!PG) throw new Error('PptxGenJS not loaded');
    const pres = new PG();
    pres.layout = 'LAYOUT_WIDE';
    const slide = pres.addSlide();
    slide.background = { color: 'F8F9FA' };
    slide.addText(data.orgName || 'Fundamentals Map', { x: 0.5, y: 0.5, fontSize: 24, bold: true, color: '0D1B2A' });
    slide.addText('Use Export JSON to get your data, then contact your admin to generate the full PPTX.', { x: 0.5, y: 1.5, fontSize: 12, color: '555555' });
    await pres.writeFile({ fileName: (data.orgName || 'Map').replace(/[^a-z0-9]/gi, '_') + '_Fundamentals_Map.pptx' });
  } catch(e) {
    console.error('PPTX export error:', e);
    alert('PowerPoint export failed: ' + e.message);
  }
}


// ── SVG Export ────────────────────────────────────────────────────────────────
function exportToSvg(data) {
  const PAL = [
    { bg: '#1b3d6e', text: '#fff', light: '#d2e5f5' },
    { bg: '#1a5c38', text: '#fff', light: '#c4edd6' },
    { bg: '#6d1f78', text: '#fff', light: '#e8d3f4' },
    { bg: '#b55000', text: '#fff', light: '#ffddb2' },
    { bg: '#1a5c5c', text: '#fff', light: '#c4edee' },
    { bg: '#7a1a2e', text: '#fff', light: '#f5c8d2' },
    { bg: '#3a5a00', text: '#fff', light: '#d8f0a5' },
    { bg: '#33337a', text: '#fff', light: '#d0d0f5' },
  ];
  const pc = idx => PAL[(idx || 0) % PAL.length];
  const gColor = gid => { const g = data.keyGoals.find(g => g.id === gid); return g ? pc(g.colorIdx) : { bg: '#666', text: '#fff', light: '#ddd' }; };
  const esc = s => (s || '').toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  // Column order — mirrors MapTable
  const goalOrder = Object.fromEntries(data.keyGoals.map((g, i) => [g.id, i]));
  const byGoal = (a, b) => (goalOrder[a.goalId] || 999) - (goalOrder[b.goalId] || 999);
  const groupByType = data.layoutGroupByType || false;
  const outFirst = data.layoutOutcomesFirst || false;

  const spanBuilder = arr => {
    const spans = []; let i = 0;
    while (i < arr.length) {
      const gid = arr[i].goalId || null;
      const goal = gid ? data.keyGoals.find(g => g.id === gid) : null;
      let c = 0;
      while (i < arr.length && (arr[i].goalId || null) === gid) { c++; i++; }
      spans.push({ goal, gid, count: c });
    }
    return spans;
  };

  let allProcs = [], goalHdrCells = [], opCount = 0, spCount = 0;
  if (groupByType) {
    const sOp = [...data.operatingProcesses].sort(byGoal);
    const sSp = [...data.supportingProcesses].sort(byGoal);
    allProcs = [...sOp.map(p => ({ proc: p, side: 'op', goalId: p.goalId })), ...sSp.map(p => ({ proc: p, side: 'sp', goalId: p.goalId }))];
    opCount = Math.max(sOp.length, 1); spCount = Math.max(sSp.length, 1);
    goalHdrCells = [...spanBuilder(sOp.map(p => ({ goalId: p.goalId }))), ...spanBuilder(sSp.map(p => ({ goalId: p.goalId })))];
  } else {
    const seen = new Set();
    for (const g of data.keyGoals) {
      for (const p of data.operatingProcesses)  if (p.goalId === g.id) { allProcs.push({ proc: p, side: 'op', goalId: g.id }); seen.add(p.id); }
      for (const p of data.supportingProcesses) if (p.goalId === g.id) { allProcs.push({ proc: p, side: 'sp', goalId: g.id }); seen.add(p.id); }
    }
    for (const p of data.operatingProcesses)  if (!seen.has(p.id)) { allProcs.push({ proc: p, side: 'op', goalId: null }); seen.add(p.id); }
    for (const p of data.supportingProcesses) if (!seen.has(p.id)) { allProcs.push({ proc: p, side: 'sp', goalId: null }); seen.add(p.id); }
    goalHdrCells = spanBuilder(allProcs.map(x => ({ goalId: x.goalId })));
  }

  const sortedOut = [...data.outcomes].sort((a, b) => (goalOrder[a.goalId] || 999) - (goalOrder[b.goalId] || 999));
  const outSpans  = spanBuilder(sortedOut.map(o => ({ goalId: o.goalId })));

  const procCols  = allProcs.length;
  const totCols   = Math.max(procCols, 3);
  const fillCols  = totCols - procCols;
  const outCount  = Math.max(sortedOut.length, 1);

  // Dimensions
  const SB = 44, COL = 162, PAD = 5, F = 11; // F = base font size
  const W      = SB + totCols * COL;
  const outColW = Math.max(Math.floor((W - SB) / outCount), 120);
  const outW   = SB + outCount * outColW;

  const hdrAlignOut = outFirst && sortedOut.length > 0;
  const hdrCols  = hdrAlignOut ? Math.max(outCount, 3) : totCols;
  const hdrColW  = hdrAlignOut ? outColW : COL;
  const hdrW     = SB + hdrCols * hdrColW;
  const hdrFill  = hdrCols - (hdrAlignOut ? outCount : procCols);
  const hdrGoals = hdrAlignOut ? outSpans : goalHdrCells;
  const c1 = Math.ceil(hdrCols / 3), c2 = Math.ceil((hdrCols - c1) / 2), c3 = hdrCols - c1 - c2;

  // Text wrapping — handles literal newlines in data by splitting on both \n and spaces
  const wrap = (text, maxW, fs) => {
    if (!text) return [];
    const charW = fs * 0.52;
    const maxC  = Math.max(1, Math.floor((maxW - PAD * 2) / charW));
    // First split on literal newlines to get paragraphs, then wrap each
    const paragraphs = (text + '').replace(/\r/g, '').split('\n');
    const lines = [];
    for (const para of paragraphs) {
      if (!para.trim()) { lines.push(''); continue; }
      const words = para.split(' ');
      let line = '';
      for (const w of words) {
        const t = line ? line + ' ' + w : w;
        if (t.length > maxC && line) { lines.push(line); line = w; }
        else line = t;
      }
      if (line) lines.push(line);
    }
    return lines;
  };

  // Compute height needed to fit wrapped text
  const textH = (text, colW, fs, lineH) => {
    const lines = wrap(text, colW, fs);
    return Math.max(lines.length, 1) * fs * (lineH || 1.45) + PAD * 2;
  };

  // Draw text lines into SVG
  const drawText = (x, ry, colW, text, fs, color, bold, align, lineH) => {
    if (!text) return;
    const lh   = (lineH || 1.45) * fs;
    const lines = wrap(text, colW, fs);
    const anch  = align === 'center' ? 'middle' : 'start';
    const tx    = align === 'center' ? x + colW / 2 : x + PAD;
    lines.forEach((ln, i) => {
      els.push('<text x="' + tx + '" y="' + (ry + PAD + fs + i * lh) + '" font-size="' + fs + '" font-family="Arial,sans-serif" fill="' + color + '" text-anchor="' + anch + '"' + (bold ? ' font-weight="bold"' : '') + '>' + esc(ln) + '</text>');
    });
  };

  const rect = (x, ry, w, h, fill) => {
    els.push('<rect x="' + x + '" y="' + ry + '" width="' + w + '" height="' + h + '" fill="' + fill + '" stroke="#bbb" stroke-width="0.5"/>');
  };

  const vLabel = (x, ry, w, h, text) => {
    const cx = x + w / 2, cy = ry + h / 2;
    els.push('<text x="' + cx + '" y="' + cy + '" font-size="7.5" font-family="Arial,sans-serif" fill="#fff" text-anchor="middle" dominant-baseline="middle" font-weight="bold" letter-spacing="0.8" transform="rotate(-90,' + cx + ',' + cy + ')">' + esc(text.toUpperCase()) + '</text>');
  };

  const els = [];
  let y = 0;

  // ── HEADER (Foundations + Key Goals) ────────────────────────────────────
  const H_ORG  = 38;
  // MVV — fixed large font (18px), generous fixed height
  const MVV_FS = 18;
  const MVV_LABEL_H = 18;
  const H_MVV = Math.max(
    MVV_LABEL_H + textH(data.mission, c1 * hdrColW, MVV_FS, 1.45),
    MVV_LABEL_H + textH(data.vision,  c2 * hdrColW, MVV_FS, 1.45),
    MVV_LABEL_H + textH(data.values,  c3 * hdrColW, MVV_FS, 1.45),
    100
  );
  const H_GOAL = 80;

  // Org name row
  rect(0, y, hdrW, H_ORG + H_MVV, '#0d1b2a');
  vLabel(0, y, SB, H_ORG + H_MVV, 'Foundations');
  rect(SB, y, hdrW - SB, H_ORG, '#0d1b2a');
  drawText(SB + PAD, y, hdrW - SB - PAD, data.orgName || '', 20, '#fff', true, 'left');
  y += H_ORG;

  // MVV row — 18px bold text, height sized to fit content
  const mvvCells = [
    { span: c1, bg: '#e8f0fb', text: data.mission, color: '#111', label: 'MISSION', lc: '#1b3d6e' },
    { span: c2, bg: '#1b3d6e', text: data.vision,  color: '#fff', label: 'VISION',  lc: '#ccc' },
    { span: c3, bg: '#1e1e1e', text: data.values,  color: '#fff', label: 'VALUES',  lc: '#ccc' },
  ];
  let mx = SB;
  mvvCells.forEach(({ span, bg, text, color, label, lc }) => {
    const cw = span * hdrColW;
    rect(mx, y, cw, H_MVV, bg);
    // Label
    els.push('<text x="' + (mx + PAD) + '" y="' + (y + MVV_LABEL_H - 4) + '" font-size="9" font-family="Arial,sans-serif" fill="' + lc + '" font-weight="bold" letter-spacing="0.8">' + label + '</text>');
    // Content text at MVV_FS starting below label
    const lh = MVV_FS * 1.45;
    const lines = wrap(text, cw, MVV_FS);
    lines.forEach((ln, i) => {
      els.push('<text x="' + (mx + PAD) + '" y="' + (y + MVV_LABEL_H + MVV_FS + i * lh) + '" font-size="' + MVV_FS + '" font-family="Arial,sans-serif" fill="' + color + '">' + esc(ln) + '</text>');
    });
    mx += cw;
  });
  y += H_MVV;

  // Key Goals row
  rect(0, y, SB, H_GOAL, '#1b3d6e');
  vLabel(0, y, SB, H_GOAL, 'Key Goals');
  let gx = SB;
  hdrGoals.forEach(({ goal, count }) => {
    const cw = count * hdrColW;
    const c  = goal ? pc(goal.colorIdx) : { bg: '#666', text: '#fff' };
    rect(gx, y, cw, H_GOAL, c.bg);
    drawText(gx, y, cw, goal ? goal.name : 'Unassigned', 13, c.text, true, 'center');
    gx += cw;
  });
  if (hdrFill > 0) rect(gx, y, hdrFill * hdrColW, H_GOAL, '#888');
  y += H_GOAL;

  // ── DRAW OUTCOMES ────────────────────────────────────────────────────────
  const drawOutcomes = startY => {
    let oy = startY;
    const H_OG  = 38;
    // Compute row heights dynamically
    const H_ON  = Math.max(...sortedOut.map(o => {
      const nameH  = textH(o.name, outColW, 10, 1.4);
      const ownerH = o.owner ? textH(o.owner, outColW, 8, 1.35) + 4 : 0;
      return nameH + ownerH + PAD * 2 + 8; // extra top padding for code label
    }), 50);
    const H_OM  = Math.max(...sortedOut.map(o => {
      if (!o.measures || !o.measures.length) return 30;
      return o.measures.reduce((acc, m) => acc + textH((m.code ? m.code + '  ' : '') + (m.text || ''), outColW, 8, 1.35) + (m.measureOwner ? 12 : 0), PAD * 2 + 14);
    }), 44);
    const outH = H_OG + H_ON + H_OM;

    // Sidebar
    rect(0, oy, SB, outH, '#555');
    vLabel(0, oy, SB, outH, 'Outcomes');

    // Goal spans
    let ox = SB;
    outSpans.forEach(({ goal, count }) => {
      const cw = count * outColW;
      const c  = goal ? pc(goal.colorIdx) : { bg: '#666', text: '#fff' };
      rect(ox, oy, cw, H_OG, c.bg);
      drawText(ox, oy, cw, goal ? goal.name : 'Unassigned', 12, c.text, true, 'center');
      ox += cw;
    });
    oy += H_OG;

    // Outcome names + owners — stacked: code -> name lines -> owner lines
    sortedOut.forEach((o, i) => {
      const c = gColor(o.goalId);
      const ox2 = SB + i * outColW;
      rect(ox2, oy, outColW, H_ON, c.light);
      let textY = oy + PAD;
      if (o.code) {
        textY += 8;
        els.push('<text x="' + (ox2 + outColW / 2) + '" y="' + textY + '" font-size="7" font-family="Arial,sans-serif" fill="#777" text-anchor="middle">' + esc(o.code) + '</text>');
        textY += 4;
      }
      // Name (bold, center)
      const nameLines = wrap(o.name, outColW, 10);
      nameLines.forEach(ln => {
        textY += 13;
        els.push('<text x="' + (ox2 + outColW / 2) + '" y="' + textY + '" font-size="11" font-family="Arial,sans-serif" fill="#222" text-anchor="middle" font-weight="bold">' + esc(ln) + '</text>');
      });
      // Owner (italic, center, below name)
      if (o.owner) {
        textY += 6;
        const ownerLines = wrap(o.owner, outColW, 8);
        ownerLines.forEach(ln => {
          textY += 11;
          els.push('<text x="' + (ox2 + outColW / 2) + '" y="' + textY + '" font-size="9" font-family="Arial,sans-serif" fill="#555" text-anchor="middle" font-style="italic">' + esc(ln) + '</text>');
        });
      }
    });
    oy += H_ON;

    // Outcome measures
    sortedOut.forEach((o, i) => {
      const c = gColor(o.goalId);
      const ox2 = SB + i * outColW;
      rect(ox2, oy, outColW, H_OM, c.light);
      els.push('<text x="' + (ox2 + PAD) + '" y="' + (oy + PAD + 9) + '" font-size="7.5" font-family="Arial,sans-serif" fill="#777" font-weight="bold" letter-spacing="0.5">OUTCOME MEASURES</text>');
      let my = oy + PAD + 8 + 11;
      (o.measures || []).forEach(m => {
        if (my > oy + H_OM - 8) return;
        const label = (m.code ? m.code + '  ' : '') + (m.text || '');
        const lns   = wrap(label, outColW, 8);
        lns.forEach(ln => {
          if (my > oy + H_OM - 8) return;
          els.push('<text x="' + (ox2 + PAD) + '" y="' + my + '" font-size="8" font-family="Arial,sans-serif" fill="#222">' + esc(ln) + '</text>');
          my += 11;
        });
        if (m.measureOwner && my <= oy + H_OM - 8) {
          els.push('<text x="' + (ox2 + PAD + 6) + '" y="' + my + '" font-size="7.5" font-family="Arial,sans-serif" fill="#555" font-style="italic">' + esc(m.measureOwner) + '</text>');
          my += 10;
        }
      });
    });
    return oy + H_OM;
  };

  // ── DRAW PROCESSES ───────────────────────────────────────────────────────
  const drawProcesses = startY => {
    let py = startY;

    // Dynamic row heights: measure tallest cell in each row
    const H_OPS  = 22; // Op/Sp label row
    const H_PN   = Math.max(...allProcs.map(({ proc }) => {
      const nameH  = textH(proc.name, COL, 11, 1.35);
      const ownerH = proc.processOwner ? textH(proc.processOwner, COL, 8, 1.35) + 8 : 0;
      return nameH + ownerH + 16; // 16 = code label + padding
    }), 44);
    const H_SUB  = Math.max(...allProcs.map(({ proc }) => {
      const subs = proc.subProcesses || [];
      if (!subs.length) return 30;
      return subs.reduce((acc, s) => acc + textH(s, COL, 9, 1.35), PAD * 2);
    }), 60);
    const H_PM   = Math.max(...allProcs.map(({ proc }) => {
      const ms = proc.processMeasures || [];
      if (!ms.length) return 30;
      return ms.reduce((acc, m) => acc + textH((m.code ? m.code + '  ' : '') + (m.text || ''), COL, 8, 1.35) + (m.measureOwner ? 12 : 0), PAD * 2 + 16);
    }), 44);

    const procH = H_OPS + H_PN + H_SUB + H_PM;

    // Sidebar
    rect(0, py, SB, procH, '#3a3a3a');
    vLabel(0, py, SB, procH, 'Core Processes');

    // Op/Sp label row
    allProcs.forEach(({ proc, side }, i) => {
      const bg = side === 'op' ? '#2c3e50' : '#445566';
      const px = SB + i * COL;
      rect(px, py, COL, H_OPS, bg);
      drawText(px, py - 2, COL, side === 'op' ? 'Operating' : 'Supporting', 9, '#fff', false, 'center');
    });
    if (fillCols > 0) rect(SB + procCols * COL, py, fillCols * COL, H_OPS, '#888');
    py += H_OPS;

    // Process name + owner
    allProcs.forEach(({ proc, goalId }, i) => {
      const c  = gColor(goalId);
      const px = SB + i * COL;
      rect(px, py, COL, H_PN, c.bg);
      if (proc.code) els.push('<text x="' + (px + PAD) + '" y="' + (py + PAD + 7) + '" font-size="7" font-family="Arial,sans-serif" fill="rgba(255,255,255,0.7)">' + esc(proc.code) + '</text>');
      drawText(px, py + 6, COL, proc.name, 11, c.text, true, 'left');
      if (proc.processOwner) {
        const owY = py + H_PN - (wrap(proc.processOwner, COL, 8).length * 11) - PAD;
        els.push('<line x1="' + px + '" y1="' + owY + '" x2="' + (px + COL) + '" y2="' + owY + '" stroke="rgba(255,255,255,0.25)" stroke-width="0.5"/>');
        wrap(proc.processOwner, COL, 8).forEach((ln, li) => {
          els.push('<text x="' + (px + PAD) + '" y="' + (owY + 10 + li * 11) + '" font-size="8" font-family="Arial,sans-serif" fill="rgba(255,255,255,0.85)" font-style="italic">' + esc(ln) + '</text>');
        });
      }
    });
    if (fillCols > 0) rect(SB + procCols * COL, py, fillCols * COL, H_PN, '#888');
    py += H_PN;

    // Sub-processes
    allProcs.forEach(({ proc }, i) => {
      const px = SB + i * COL;
      rect(px, py, COL, H_SUB, '#fff');
      let sy = py + PAD + 10;
      (proc.subProcesses || []).forEach(s => {
        if (sy > py + H_SUB - 8) return;
        const lns = wrap(s, COL - PAD * 2 - 8, 9);
        lns.forEach((ln, li) => {
          if (sy > py + H_SUB - 8) return;
          els.push('<text x="' + (px + PAD + (li === 0 ? 0 : 8)) + '" y="' + sy + '" font-size="10" font-family="Arial,sans-serif" fill="#222">' + (li === 0 ? '&#8226; ' : '') + esc(ln) + '</text>');
          sy += 14;
        });
      });
    });
    if (fillCols > 0) rect(SB + procCols * COL, py, fillCols * COL, H_SUB, '#f5f5f5');
    py += H_SUB;

    // Process measures
    allProcs.forEach(({ proc, goalId }, i) => {
      const c  = gColor(goalId);
      const px = SB + i * COL;
      rect(px, py, COL, H_PM, c.light);
      els.push('<text x="' + (px + PAD) + '" y="' + (py + PAD + 9) + '" font-size="7.5" font-family="Arial,sans-serif" fill="#777" font-weight="bold" letter-spacing="0.5">PROCESS MEASURES</text>');
      let my = py + PAD + 8 + 12;
      (proc.processMeasures || []).forEach(m => {
        if (my > py + H_PM - 8) return;
        const label = (m.code ? m.code + '  ' : '') + (m.text || '');
        const lns   = wrap(label, COL, 8);
        lns.forEach(ln => {
          if (my > py + H_PM - 8) return;
          els.push('<text x="' + (px + PAD) + '" y="' + my + '" font-size="8" font-family="Arial,sans-serif" fill="#222">' + esc(ln) + '</text>');
          my += 11;
        });
        if (m.measureOwner && my <= py + H_PM - 8) {
          els.push('<text x="' + (px + PAD + 6) + '" y="' + my + '" font-size="7.5" font-family="Arial,sans-serif" fill="#555" font-style="italic">' + esc(m.measureOwner) + '</text>');
          my += 11;
        }
      });
    });
    if (fillCols > 0) rect(SB + procCols * COL, py, fillCols * COL, H_PM, '#f5f5f5');
    return py + H_PM;
  };

  // ── Draw in layout order ─────────────────────────────────────────────────
  if (outFirst) {
    y = drawOutcomes(y);
    y = drawProcesses(y);
  } else {
    y = drawProcesses(y);
    y = drawOutcomes(y);
  }

  const svgW = Math.max(W, outW, hdrW);
  const nl = String.fromCharCode(10);
  // Use viewBox so SVG scales to fit — remove fixed width/height so viewer controls zoom
  const svgStr = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ' + svgW + ' ' + y + '" style="width:100%;height:auto">' + nl + els.join(nl) + nl + '</svg>';

  const blob = new Blob([svgStr], { type: 'image/svg+xml' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (data.orgName || 'Map').replace(/[^a-z0-9]/gi, '_') + '_Fundamentals_Map.svg';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// exportToPdf is handled via printMode state in the App component (see handlePdf)

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [data,          setData]          = useState(INIT);
  const [tab,           setTab]           = useState('editor');
  const [section,       setSection]       = useState('foundation');
  const [editPath,      setEditPath]      = useState(null);
  const [saveStatus,    setSaveStatus]    = useState(null);
  const [pptxStatus,    setPptxStatus]    = useState(null);
  const [svgStatus,     setSvgStatus]     = useState(null);
  const [showReset,     setShowReset]     = useState(false);
  const [printMode,     setPrintMode]     = useState(false);
  const [fullscreen,    setFullscreen]    = useState(false);
  const [mapScale,      setMapScale]      = useState(1);
  const fsContainerRef = useRef(null);
  const [exportModal,   setExportModal]   = useState(false);
  const [copied,        setCopied]        = useState(false);
  const fileRef  = useRef(null);
  const mapRef   = useRef(null);

  // Load from storage on mount
  useEffect(() => {
    loadFromStorage().then(saved => { if (saved) setData(saved); }).catch(() => {});
  }, []);

  // Escape key exits fullscreen
  useEffect(() => {
    const onKey = e => { if (e.key === 'Escape' && fullscreen) setFullscreen(false); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [fullscreen]);

  // Compute scale to fit map in fullscreen container (no scrollbars)
  useEffect(() => {
    if (!fullscreen || !fsContainerRef.current || !mapRef.current) { setMapScale(1); return; }
    const computeScale = () => {
      const container = fsContainerRef.current;
      const map       = mapRef.current;
      if (!container || !map) return;
      // Available space: container minus toolbar (~44px)
      const availW = container.clientWidth  - 32;
      const availH = container.clientHeight - 60;
      const mapW   = map.scrollWidth;
      const mapH   = map.scrollHeight;
      const scale  = Math.min(availW / mapW, availH / mapH, 1);
      setMapScale(scale);
    };
    // Wait for render then compute
    const t = setTimeout(computeScale, 50);
    window.addEventListener('resize', computeScale);
    return () => { clearTimeout(t); window.removeEventListener('resize', computeScale); };
  }, [fullscreen, data]);

  const handleSave = async () => {
    setSaveStatus('saving');
    try { await saveToStorage(data); setSaveStatus('saved'); setTimeout(() => setSaveStatus(null), 2500); }
    catch (e) { setSaveStatus('error'); setTimeout(() => setSaveStatus(null), 3000); }
  };

  const handlePptx = async () => {
    setPptxStatus('loading');
    try { await exportToPptx(data); setPptxStatus(null); }
    catch (e) { console.error(e); setPptxStatus('error'); setTimeout(() => setPptxStatus(null), 3000); }
  };

  const handleSvg = () => {
    setSvgStatus('loading');
    try { exportToSvg(data); setSvgStatus(null); }
    catch(e) { console.error(e); setSvgStatus('error'); setTimeout(() => setSvgStatus(null), 3000); }
  };

  const handlePdf = () => {
    setTab('map');
    setTimeout(() => setPrintMode(true), 100);
  };

  // Export JSON: show modal with copyable text (programmatic downloads blocked in sandbox)
  const openExportModal = () => { setCopied(false); setExportModal(true); };
  const copyJSON = () => {
    const json = JSON.stringify(data, null, 2);
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(json).then(() => { setCopied(true); setTimeout(() => setCopied(false), 2000); });
    } else {
      // fallback: select textarea text
      const ta = document.getElementById('_json_export_ta');
      if (ta) { ta.select(); document.execCommand('copy'); setCopied(true); setTimeout(() => setCopied(false), 2000); }
    }
  };

  const importJSON = e => {
    const file = e.target.files[0]; if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => { try { setData(JSON.parse(ev.target.result)); setTab('editor'); } catch { alert('Invalid JSON file.'); } };
    reader.readAsText(file); e.target.value = '';
  };

  const confirmReset = () => { setData(BLANK); setTab('editor'); setShowReset(false); };

  const SECTIONS = [['foundation','Foundation'],['goals','Key goals'],['operating','Operating processes'],['supporting','Supporting processes'],['outcomes','Outcomes']];
  const navBtn = (t, label) => (
    <button key={t} onClick={() => setTab(t)} style={{ background: tab === t ? 'var(--color-background-info)' : 'transparent', color: tab === t ? 'var(--color-text-info)' : 'var(--color-text-secondary)', border: '0.5px solid ' + (tab === t ? 'var(--color-border-info)' : 'var(--color-border-secondary)'), borderRadius: 5, padding: '4px 12px', fontSize: 12, cursor: 'pointer', fontFamily: 'inherit' }}>{label}</button>
  );

  const saveLabel = saveStatus === 'saving' ? 'Saving...' : saveStatus === 'saved' ? 'Saved' : saveStatus === 'error' ? 'Error!' : 'Save';
  const saveV     = saveStatus === 'saved' ? 'success' : saveStatus === 'error' ? 'danger' : 'default';
  const pptxLabel = pptxStatus === 'loading' ? 'Saving...' : pptxStatus === 'error' ? 'Failed' : 'Export to PowerPoint';
  const pdfLabel  = 'Export to PDF';

  const overlay = { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: 'rgba(0,0,0,0.55)', zIndex: 1000, display: 'flex', alignItems: 'center', justifyContent: 'center' };
  const modal   = { background: '#ffffff', borderRadius: 10, padding: 24, maxWidth: 520, width: '90%', boxShadow: '0 8px 40px rgba(0,0,0,0.35)', border: '1px solid #ddd' };

  return (
    <div style={{ fontFamily: 'var(--font-sans, system-ui)', background: 'var(--color-background-primary)', minHeight: 500 }}>
      <style>{`
        @media print {
          .no-print { display: none !important; }
          @page { size: 14in 8.5in landscape; margin: 0.3in; }
          body { margin: 0 !important; padding: 0 !important; }
          .map-print-root { width: 100% !important; overflow: visible !important; }
          .map-table { width: 100% !important; table-layout: fixed !important; font-size: 8pt !important; }
          .map-table td { padding: 3px 4px !important; }
        }
      `}</style>

      {/* ── Reset confirm dialog ── */}
      {showReset && (
        <div style={overlay} onClick={() => setShowReset(false)}>
          <div style={modal} onClick={e => e.stopPropagation()}>
            <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 8, color: '#111' }}>Reset to blank map?</div>
            <div style={{ fontSize: 13, color: '#555', marginBottom: 20, lineHeight: 1.5 }}>This will clear all goals, processes, outcomes, and settings. This cannot be undone.</div>
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end' }}>
              <Btn onClick={() => setShowReset(false)}>Cancel</Btn>
              <Btn onClick={confirmReset} v="danger">Reset</Btn>
            </div>
          </div>
        </div>
      )}

      {/* ── Export JSON modal ── */}
      {exportModal && (
        <div style={overlay} onClick={() => setExportModal(false)}>
          <div style={{ ...modal, maxWidth: 640 }} onClick={e => e.stopPropagation()}>
            <div style={{ fontWeight: 600, fontSize: 15, marginBottom: 6, color: '#111' }}>Export map data (JSON)</div>
            <div style={{ fontSize: 12, color: '#555', marginBottom: 10 }}>Copy the text below and save it as a <code>.json</code> file to back up or share your map.</div>
            <textarea
              id="_json_export_ta"
              readOnly
              value={JSON.stringify(data, null, 2)}
              style={{ width: '100%', height: 260, fontFamily: 'monospace', fontSize: 11, border: '1px solid #ddd', borderRadius: 6, padding: 10, boxSizing: 'border-box', resize: 'vertical', background: '#f8f8f8', color: '#222' }}
              onFocus={e => e.target.select()}
            />
            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 12 }}>
              <Btn onClick={() => setExportModal(false)}>Close</Btn>
              <Btn onClick={copyJSON} v="primary">{copied ? 'Copied!' : 'Copy to clipboard'}</Btn>
            </div>
          </div>
        </div>
      )}

      {/* ── Print preview overlay ── */}
      {printMode && (
        <div style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, background: '#fff', zIndex: 2000, overflowY: 'auto', overflowX: 'auto', padding: 20 }}>
          <div style={{ position: 'sticky', top: 0, background: '#1b3d6e', color: '#fff', padding: '10px 16px', borderRadius: 6, marginBottom: 16, display: 'flex', alignItems: 'center', gap: 16, zIndex: 10 }} className="no-print">
            <span style={{ fontSize: 14, fontWeight: 600 }}>Print Preview</span>
            <span style={{ fontSize: 13, opacity: 0.85 }}>Press <strong>Ctrl+P</strong> (Windows) or <strong>Cmd+P</strong> (Mac) — choose <em>Legal</em> paper, <em>Landscape</em>, then Save as PDF</span>
            <div style={{ flex: 1 }} />
            <button onClick={() => setPrintMode(false)} style={{ background: 'rgba(255,255,255,0.2)', color: '#fff', border: '1px solid rgba(255,255,255,0.4)', borderRadius: 5, padding: '4px 14px', fontSize: 12, cursor: 'pointer' }}>Close</button>
          </div>
          <MapTable data={data} editPath={null} setEditPath={() => {}} setData={setData} mapRef={mapRef} />
        </div>
      )}

      {/* ── Top bar ── */}
      <div className="no-print" style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '9px 14px', borderBottom: '0.5px solid var(--color-border-tertiary)', background: 'var(--color-background-secondary)', flexWrap: 'wrap' }}>
        <span style={{ fontWeight: 500, fontSize: 13, marginRight: 4, color: 'var(--color-text-primary)' }}>Fundamentals Map Builder</span>
        {navBtn('editor', 'Editor')}
        {navBtn('map', 'Preview map')}
        {tab === 'map' && (
          <button onClick={() => setFullscreen(f => !f)} title={fullscreen ? 'Exit full screen' : 'Full screen'} style={{ background: 'transparent', border: '0.5px solid var(--color-border-secondary)', borderRadius: 5, padding: '4px 10px', fontSize: 13, cursor: 'pointer', lineHeight: 1 }}>
            {fullscreen ? 'X  Exit full screen' : '[ ]  Full screen'}
          </button>
        )}
        <div style={{ flex: 1 }} />
        <Btn onClick={handleSave} v={saveV} disabled={saveStatus === 'saving'}>{saveLabel}</Btn>
        <Btn onClick={handlePptx} v="primary" disabled={pptxStatus === 'loading'}>{pptxLabel}</Btn>
        <Btn onClick={handleSvg} disabled={svgStatus === 'loading'}>{svgStatus === 'loading' ? 'Generating...' : 'Export SVG'}</Btn>
        <Btn onClick={openExportModal}>Export JSON</Btn>
        <Btn onClick={() => fileRef.current?.click()}>Import JSON</Btn>
        <Btn onClick={() => setShowReset(true)} v="danger">Reset</Btn>
        <input type="file" ref={fileRef} style={{ display: 'none' }} accept=".json" onChange={importJSON} />
      </div>

      {/* ── Editor ── */}
      {tab === 'editor' && (
        <div style={{ display: 'flex', minHeight: 600 }}>
          <div className="no-print" style={{ width: 178, flexShrink: 0, borderRight: '0.5px solid var(--color-border-tertiary)', background: 'var(--color-background-secondary)', paddingTop: 6 }}>
            {SECTIONS.map(([k, label]) => (
              <button key={k} onClick={() => setSection(k)} style={{ display: 'block', width: '100%', textAlign: 'left', padding: '9px 16px', border: 'none', background: section === k ? 'var(--color-background-primary)' : 'transparent', color: section === k ? 'var(--color-text-primary)' : 'var(--color-text-secondary)', borderLeft: section === k ? '3px solid var(--color-text-info)' : '3px solid transparent', fontSize: 12, cursor: 'pointer', fontFamily: 'inherit', lineHeight: 1.4 }}>{label}</button>
            ))}
          </div>
          <div style={{ flex: 1, padding: 24, overflowX: 'auto' }}>
            {section === 'foundation' && <FoundationForm data={data} setData={setData} />}
            {section === 'goals'      && <GoalsForm data={data} setData={setData} />}
            {section === 'operating'  && <ProcessForm data={data} setData={setData} type="operatingProcesses"  typeCode="OP" title="operating process"  desc="Core work your organization does to deliver its mission." prefix="OP" />}
            {section === 'supporting' && <ProcessForm data={data} setData={setData} type="supportingProcesses" typeCode="SP" title="supporting process" desc="Internal processes that enable and sustain operating work." prefix="SP" />}
            {section === 'outcomes'   && <OutcomesForm data={data} setData={setData} />}
          </div>
        </div>
      )}

      {/* ── Map ── */}
      {tab === 'map' && (
        <div ref={fsContainerRef} style={fullscreen ? { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, zIndex: 1500, background: '#fff', display: 'flex', flexDirection: 'column', overflow: 'hidden' } : {}}>
          <div className="no-print" style={{ padding: '7px 14px', borderBottom: '0.5px solid var(--color-border-tertiary)', display: 'flex', gap: 8, alignItems: 'center', background: 'var(--color-background-secondary)', flexShrink: 0 }}>
            {fullscreen && (
              <button onClick={() => setFullscreen(false)} style={{ background: 'transparent', border: '0.5px solid var(--color-border-secondary)', borderRadius: 5, padding: '4px 10px', fontSize: 12, cursor: 'pointer', marginRight: 4 }}>
                X Exit full screen
              </button>
            )}
            <span style={{ fontSize: 11, color: 'var(--color-text-tertiary)' }}>
              {fullscreen ? 'Esc to exit - Click any cell to edit' : 'Click any cell to edit inline - Enter to save - Esc to cancel'}
            </span>
            <div style={{ flex: 1 }} />
            <Btn onClick={handlePdf}>{pdfLabel}</Btn>
          </div>
          {fullscreen ? (
            /* Fullscreen: scale map to fit with no scrollbars */
            <div style={{ flex: 1, display: 'flex', alignItems: 'flex-start', justifyContent: 'flex-start', overflow: 'hidden', padding: 16 }}>
              <div style={{ transformOrigin: 'top left', transform: 'scale(' + mapScale + ')' }}>
                <MapTable data={data} editPath={editPath} setEditPath={setEditPath} setData={setData} mapRef={mapRef} />
              </div>
            </div>
          ) : (
            <div style={{ overflowX: 'auto', padding: 16 }}>
              <MapTable data={data} editPath={editPath} setEditPath={setEditPath} setData={setData} mapRef={mapRef} />
            </div>
          )}
        </div>
      )}
    </div>
  );
}
