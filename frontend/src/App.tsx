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

type Tab = 'report' | 'comparison'

export default function App() {
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [activeTab, setActiveTab] = useState<Tab>('report')
  const { keys, save, clear, hasLlmKey } = useSessionKeys()
  const state = useAppState({ llmProvider: keys.llmProvider, llmApiKey: keys.llmApiKey })

  const showReady = state.phase === 'ready' || state.phase === 'generated'
  const showGenerated = state.phase === 'generated'

  // Only show tabs once we have at least one generated model
  const showTabs = state.savedModels.length > 0

  return (
    <div className="app">
      <header className="app-header">
        <div className="brand">
          <span className="brand-name">AIO LBO</span>
          <span className="brand-tag">Leveraged buyout modeling</span>
        </div>
        <button
          type="button"
          className="settings-btn"
          aria-label="API Key Settings"
          title="API Key Settings"
          onClick={() => setSettingsOpen(true)}
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M21 2l-2 2m-7.61 7.61a5.5 5.5 0 1 1-7.778 7.778 5.5 5.5 0 0 1 7.777-7.777zm0 0L15.5 7.5m0 0l3 3L22 7l-3-3m-3.5 3.5L19 4" />
          </svg>
          <span>API Key Settings</span>
        </button>
      </header>

      {showTabs && (
        <nav className="tab-bar">
          <button
            type="button"
            className={`tab-btn ${activeTab === 'report' ? 'active' : ''}`}
            onClick={() => setActiveTab('report')}
          >
            Report Generation
          </button>
          <button
            type="button"
            className={`tab-btn ${activeTab === 'comparison' ? 'active' : ''}`}
            onClick={() => setActiveTab('comparison')}
          >
            Comparison
            {state.savedModels.length > 0 && (
              <span className="tab-badge">{state.savedModels.length}</span>
            )}
          </button>
        </nav>
      )}

      {activeTab === 'report' && (
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
            <CompanySnapshot
              snapshot={state.snapshot}
              userOverrides={state.userOverrides}
              onUpdateOverride={state.updateOverride}
            />
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
              hasLlmKey={hasLlmKey}
              onNeedKey={() => {
                state.showToast('Add an LLM API key in Settings')
                setSettingsOpen(true)
              }}
            />
          )}

          {showGenerated && state.sensitivity.length > 0 && (
            <SensitivityHeatmap cells={state.sensitivity} />
          )}
        </div>
      )}

      {activeTab === 'comparison' && (
        <div className="comparison-page">
          <ComparisonPanel
            savedModels={state.savedModels}
            hasLlmKey={hasLlmKey}
            llmProvider={keys.llmProvider}
            llmApiKey={keys.llmApiKey}
            onNeedKey={() => {
              state.showToast('Add an LLM API key in Settings')
              setSettingsOpen(true)
            }}
          />
        </div>
      )}

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
