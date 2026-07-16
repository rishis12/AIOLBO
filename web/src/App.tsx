import { useState } from 'react'
import { AssumptionsEditor } from './components/AssumptionsEditor'
import { CompanySnapshot } from './components/CompanySnapshot'
import { ComparisonPanel } from './components/ComparisonPanel'
import { HowItWorks } from './components/HowItWorks'
import { ResultsSummary } from './components/ResultsSummary'
import { SensitivityHeatmap } from './components/SensitivityHeatmap'
import { SettingsModal } from './components/SettingsModal'
import { SkeletonTile } from './components/SkeletonTile'
import { TickerInput } from './components/TickerInput'
import { useAppState } from './hooks/useAppState'
import { useSessionKeys } from './hooks/useSessionKeys'

export default function App() {
  const [settingsOpen, setSettingsOpen] = useState(false)
  const { keys, save, clear, hasLlmKey } = useSessionKeys()
  const state = useAppState()

  const showReady = state.phase === 'ready' || state.phase === 'generated'
  const showGenerated = state.phase === 'generated'

  return (
    <div className="app">
      <header className="app-header">
        <div className="brand">
          <span className="brand-name">AIO LBO</span>
          <span className="brand-tag">Leveraged buyout modeling</span>
        </div>
        <button
          type="button"
          className="icon-btn"
          aria-label="Settings"
          title="Settings"
          onClick={() => setSettingsOpen(true)}
        >
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <circle cx="12" cy="12" r="3" />
            <path d="M12 1v2M12 21v2M4.22 4.22l1.42 1.42M18.36 18.36l1.42 1.42M1 12h2M21 12h2M4.22 19.78l1.42-1.42M18.36 5.64l1.42-1.42" />
          </svg>
        </button>
      </header>

      <div className="dashboard-grid">
        <TickerInput
          value={state.tickerInput}
          onChange={state.setTickerInput}
          onAnalyze={() => state.analyze(state.tickerInput)}
          loading={state.phase === 'loading'}
          error={state.error}
        />

        {state.phase === 'empty' && <HowItWorks />}

        {state.phase === 'loading' && (
          <>
            <SkeletonTile className="tile-snapshot" />
            <SkeletonTile className="skeleton-mid" />
          </>
        )}

        {showReady && state.snapshot && (
          <CompanySnapshot snapshot={state.snapshot} />
        )}

        {showReady && state.assumptions && (
          <AssumptionsEditor
            assumptions={state.assumptions}
            meta={state.assumptionMeta}
            onChange={state.updateAssumption}
            onGenerate={state.generate}
            generating={state.generating}
          />
        )}

        {showGenerated && state.results && (
          <ResultsSummary
            results={state.results}
            onDownload={state.downloadExcel}
            onReport={() => state.setShowReport((v) => !v)}
            showReport={state.showReport}
          />
        )}

        {showGenerated && state.sensitivity.length > 0 && (
          <SensitivityHeatmap cells={state.sensitivity} />
        )}

        {showGenerated && (
          <ComparisonPanel
            savedModels={state.savedModels}
            hasLlmKey={hasLlmKey}
            onNeedKey={() => {
              state.showToast('Add an LLM API key in Settings')
              setSettingsOpen(true)
            }}
          />
        )}
      </div>

      <SettingsModal
        open={settingsOpen}
        keys={keys}
        onClose={() => setSettingsOpen(false)}
        onSave={save}
        onClear={clear}
      />

      {state.toast && <div className="toast">{state.toast}</div>}
    </div>
  )
}
