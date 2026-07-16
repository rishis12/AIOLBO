import { useState } from 'react'
import type { CompanySnapshot } from '../types'
import styles from './CompanySnapshot.module.css'

interface Props {
  snapshot: CompanySnapshot
}

function formatCap(n: number): string {
  if (n >= 1e12) return `$${(n / 1e12).toFixed(2)}T`
  if (n >= 1e9) return `$${(n / 1e9).toFixed(1)}B`
  if (n >= 1e6) return `$${(n / 1e6).toFixed(0)}M`
  return `$${n.toLocaleString()}`
}

function formatShares(n: number): string {
  if (n >= 1e9) return `${(n / 1e9).toFixed(2)}B`
  if (n >= 1e6) return `${(n / 1e6).toFixed(1)}M`
  return n.toLocaleString()
}

export function CompanySnapshot({ snapshot }: Props) {
  const [collapsed, setCollapsed] = useState(false)
  const v = snapshot.validation
  const maxRev = Math.max(...snapshot.revenueHistory.map((p) => p.revenue), 1)
  const flags = [
    ...v.defaultsApplied.map((t) => ({ type: 'defaulted' as const, text: t })),
    ...v.substituteWarnings.map((t) => ({ type: 'substituted' as const, text: t })),
  ]

  return (
    <section className={`tile tile-snapshot collapsible ${collapsed ? 'collapsed' : ''}`}>
      <button
        type="button"
        className="collapsible-header"
        onClick={() => setCollapsed((c) => !c)}
      >
        Company Snapshot
        <span>{collapsed ? '+' : '−'}</span>
      </button>
      <h2 className="tile-title">Company Snapshot</h2>
      <div className="collapsible-body">
        <div className={styles.head}>
          <div>
            <div className={styles.name}>{snapshot.companyName}</div>
            <div className={styles.tickerRow}>
              <span className={`mono ${styles.ticker}`}>{snapshot.ticker}</span>
              <span className="badge badge-sector">
                SIC {snapshot.sicCode} · {snapshot.sicDescription}
              </span>
            </div>
          </div>
          <div className={styles.quality}>
            <span className={`status-dot ${v.status}`} />
            <span className={styles.qualityLabel}>
              {v.status === 'pass' ? 'Pass' : v.status === 'degraded' ? 'Degraded' : 'Fail'}
            </span>
          </div>
        </div>

        <div className={styles.metrics}>
          <div>
            <div className={styles.metricLabel}>Price</div>
            <div className={`mono ${styles.metricVal}`}>${snapshot.currentPrice.toFixed(2)}</div>
          </div>
          <div>
            <div className={styles.metricLabel}>Shares</div>
            <div className={`mono ${styles.metricVal}`}>{formatShares(snapshot.sharesOutstanding)}</div>
          </div>
          <div>
            <div className={styles.metricLabel}>Market Cap</div>
            <div className={`mono ${styles.metricVal}`}>{formatCap(snapshot.marketCap)}</div>
          </div>
        </div>

        <div className={styles.chartLabel}>5-year revenue</div>
        <div className={styles.bars} role="img" aria-label="Five year revenue trend">
          {snapshot.revenueHistory.map((p) => (
            <div key={p.year} className={styles.barCol}>
              <div
                className={styles.bar}
                style={{ height: `${Math.max(8, (p.revenue / maxRev) * 100)}%` }}
                title={`${p.year}: ${formatCap(p.revenue)}`}
              />
              <span className={styles.barYear}>{String(p.year).slice(2)}</span>
            </div>
          ))}
        </div>

        {flags.length > 0 && (
          <div className={styles.flags}>
            {flags.map((f) => (
              <span
                key={f.text}
                className={`badge ${f.type === 'defaulted' ? 'badge-defaulted' : 'badge-substituted'}`}
              >
                {f.type === 'defaulted' ? 'Defaulted' : 'Substituted'}: {f.text}
              </span>
            ))}
          </div>
        )}
      </div>
    </section>
  )
}
