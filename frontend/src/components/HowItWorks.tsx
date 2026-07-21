import styles from './HowItWorks.module.css'

export function HowItWorks() {
  return (
    <section className={`tile tile-how ${styles.wrap}`}>
      <h2 className="tile-title">How it works</h2>
      <ol className={styles.steps}>
        <li>
          <strong>Analyze</strong> a public ticker — SEC financials + price, then validation.
        </li>
        <li>
          <strong>Edit assumptions</strong> — the 14 independent LBO inputs, pre-filled from history.
        </li>
        <li>
          <strong>Generate</strong> a formula-driven Excel model, feasibility score, and sensitivity grid.
        </li>
        <li>
          <strong>Compare</strong> scenarios or companies and optionally request LLM commentary (BYOK).
        </li>
      </ol>
      <p className={styles.note}>
        This is a UI shell with demo data. Try <span className="mono">AAPL</span>,{' '}
        <span className="mono">CCL</span>, or <span className="mono">FIZZ</span>. Use{' '}
        <span className="mono">JPM</span> / <span className="mono">ARM</span> to see validation failures.
      </p>
    </section>
  )
}
