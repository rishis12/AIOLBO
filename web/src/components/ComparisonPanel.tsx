import { useMemo, useState } from 'react'
import { compareSlots, mockCommentary } from '../mock/engine'
import type { ComparisonSlot } from '../types'
import styles from './ComparisonPanel.module.css'

interface Props {
  savedModels: ComparisonSlot[]
  hasLlmKey: boolean
  onNeedKey: () => void
}

export function ComparisonPanel({ savedModels, hasLlmKey, onNeedKey }: Props) {
  const [collapsed, setCollapsed] = useState(false)
  const [idA, setIdA] = useState('')
  const [idB, setIdB] = useState('')
  const [commentary, setCommentary] = useState<string | null>(null)

  const slotA = savedModels.find((s) => s.id === idA) ?? null
  const slotB = savedModels.find((s) => s.id === idB) ?? null

  const comparison = useMemo(() => {
    if (!slotA || !slotB) return null
    return compareSlots(slotA, slotB)
  }, [slotA, slotB])

  return (
    <section className={`tile tile-comparison collapsible ${collapsed ? 'collapsed' : ''}`}>
      <button
        type="button"
        className="collapsible-header"
        onClick={() => setCollapsed((c) => !c)}
      >
        Comparison Panel
        <span>{collapsed ? '+' : '−'}</span>
      </button>
      <h2 className="tile-title">Comparison Panel</h2>
      <div className="collapsible-body">
        <div className={styles.slots}>
          <label className={styles.slot}>
            <span className="field-label">File A</span>
            <select
              className="select"
              value={idA}
              onChange={(e) => {
                setIdA(e.target.value)
                setCommentary(null)
              }}
            >
              <option value="">Select generated model…</option>
              {savedModels.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.label}
                </option>
              ))}
            </select>
          </label>
          <label className={styles.slot}>
            <span className="field-label">File B</span>
            <select
              className="select"
              value={idB}
              onChange={(e) => {
                setIdB(e.target.value)
                setCommentary(null)
              }}
            >
              <option value="">Select generated model…</option>
              {savedModels.map((s) => (
                <option key={s.id} value={s.id}>
                  {s.label}
                </option>
              ))}
            </select>
          </label>
        </div>

        {savedModels.length === 0 && (
          <p className={styles.empty}>Generate at least one model to populate comparison slots.</p>
        )}

        {comparison && (
          <>
            <div className={styles.mode}>
              Mode:{' '}
              <strong>
                {comparison.mode === 'scenario' ? 'Scenario Comparison' : 'Company Comparison'}
              </strong>
            </div>
            <div className={styles.tableWrap}>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>Metric</th>
                    <th>Value A</th>
                    <th>Value B</th>
                    <th>Change</th>
                  </tr>
                </thead>
                <tbody>
                  {comparison.deltas.map((d) => (
                    <tr key={d.metric}>
                      <td>{d.metric}</td>
                      <td className="mono">{d.valueA}</td>
                      <td className="mono">{d.valueB}</td>
                      <td className="mono">{d.change}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <button
              type="button"
              className="btn btn-secondary"
              style={{ marginTop: '0.85rem' }}
              onClick={() => {
                if (!hasLlmKey) {
                  onNeedKey()
                  return
                }
                setCommentary(mockCommentary(comparison.mode, comparison.deltas))
              }}
            >
              Generate Commentary
            </button>

            {!hasLlmKey && (
              <p className={styles.keyHint}>Requires an LLM API key in Settings (session only).</p>
            )}

            {commentary && <p className={styles.commentary}>{commentary}</p>}
          </>
        )}
      </div>
    </section>
  )
}
