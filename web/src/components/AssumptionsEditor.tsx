import { useState } from 'react'
import type { AssumptionKey, AssumptionMeta, Assumptions } from '../types'
import styles from './AssumptionsEditor.module.css'

interface Props {
  assumptions: Assumptions
  meta: AssumptionMeta[]
  onChange: <K extends AssumptionKey>(key: K, value: Assumptions[K]) => void
  onGenerate: () => void
  generating: boolean
}

const GROUPS: Array<{ id: AssumptionMeta['group']; title: string }> = [
  { id: 'deal', title: 'Deal Structure' },
  { id: 'operating', title: 'Operating' },
  { id: 'debt', title: 'Debt Terms' },
  { id: 'exit', title: 'Exit' },
]

function displayValue(value: number, format: AssumptionMeta['format']): string {
  if (format === 'percent') return (value * 100).toFixed(1)
  if (format === 'years') return String(Math.round(value))
  return value.toFixed(1)
}

function parseInput(
  raw: string,
  format: AssumptionMeta['format'],
): number | null {
  const n = Number(raw)
  if (Number.isNaN(n)) return null
  if (format === 'percent') return n / 100
  return n
}

export function AssumptionsEditor({
  assumptions,
  meta,
  onChange,
  onGenerate,
  generating,
}: Props) {
  const [collapsed, setCollapsed] = useState(false)
  const byKey = Object.fromEntries(meta.map((m) => [m.key, m])) as Record<
    AssumptionKey,
    AssumptionMeta
  >

  return (
    <section className={`tile tile-assumptions collapsible ${collapsed ? 'collapsed' : ''}`}>
      <button
        type="button"
        className="collapsible-header"
        onClick={() => setCollapsed((c) => !c)}
      >
        Assumptions Editor
        <span>{collapsed ? '+' : '−'}</span>
      </button>
      <h2 className="tile-title">Assumptions Editor</h2>
      <div className="collapsible-body">
        {GROUPS.map((g) => (
          <div key={g.id} className={styles.group}>
            <h3 className={styles.groupTitle}>{g.title}</h3>
            <div className={styles.fields}>
              {meta
                .filter((m) => m.group === g.id)
                .map((m) => {
                  const val = assumptions[m.key]
                  const fallback = m.flag === 'defaulted' || m.flag === 'substituted'
                  return (
                    <label key={m.key} className={styles.field}>
                      <span className="field-label">{m.label}</span>
                      <div className={styles.inputWrap}>
                        <input
                          className={`input ${fallback ? 'fallback' : ''}`}
                          type="number"
                          step={m.format === 'years' ? 1 : 0.1}
                          value={displayValue(val, m.format)}
                          onChange={(e) => {
                            const parsed = parseInput(e.target.value, m.format)
                            if (parsed !== null) onChange(m.key, parsed as Assumptions[typeof m.key])
                          }}
                        />
                        <span className={styles.suffix}>
                          {m.format === 'percent' ? '%' : m.format === 'multiple' ? 'x' : 'yr'}
                        </span>
                      </div>
                      <span className="field-hint">{byKey[m.key]?.source}</span>
                    </label>
                  )
                })}
            </div>
          </div>
        ))}

        <button
          type="button"
          className="btn btn-primary btn-block"
          onClick={onGenerate}
          disabled={generating}
        >
          {generating ? 'Generating…' : 'Generate Model'}
        </button>
      </div>
    </section>
  )
}
