import styles from './TickerInput.module.css'
import { SOURCE_HEALTH } from '../mock/data'
import type { AnalyzeFailure } from '../types'

interface Props {
  value: string
  onChange: (v: string) => void
  onAnalyze: () => void
  loading: boolean
  error: AnalyzeFailure | null
}

export function TickerInput({ value, onChange, onAnalyze, loading, error }: Props) {
  return (
    <section className={`tile tile-ticker ${styles.hero}`}>
      <h2 className="tile-title">Ticker Input</h2>
      <form
        className={styles.row}
        onSubmit={(e) => {
          e.preventDefault()
          if (value.trim()) onAnalyze()
        }}
      >
        <input
          className={`input ${styles.input}`}
          type="text"
          value={value}
          onChange={(e) => onChange(e.target.value.toUpperCase())}
          placeholder="Enter ticker (e.g., AAPL)"
          aria-label="Ticker symbol"
          autoComplete="off"
          spellCheck={false}
          disabled={loading}
        />
        <button className="btn btn-primary" type="submit" disabled={loading || !value.trim()}>
          {loading ? 'Analyzing…' : 'Analyze'}
        </button>
      </form>

      <div className={styles.chips}>
        <span className={`chip ${SOURCE_HEALTH.secEdgar ? 'ok' : 'bad'}`}>
          SEC EDGAR {SOURCE_HEALTH.secEdgar ? '✓' : '✗'}
        </span>
        <span className={`chip ${SOURCE_HEALTH.twelveData ? 'ok' : 'bad'}`}>
          Twelve Data {SOURCE_HEALTH.twelveData ? '✓' : '✗'}
        </span>
      </div>

      {error && (
        <div className="error-banner" role="alert">
          <strong>Validation failed — {error.ticker}</strong>
          <ul>
            {error.validation.disqualifyingReasons.map((r) => (
              <li key={r}>{r}</li>
            ))}
            {error.validation.missingHard.map((r) => (
              <li key={r}>{r}</li>
            ))}
          </ul>
          <p style={{ margin: '0.65rem 0 0.25rem', fontWeight: 600 }}>Next steps</p>
          <ul>
            {error.nextSteps.map((s) => (
              <li key={s}>{s}</li>
            ))}
          </ul>
        </div>
      )}
    </section>
  )
}
