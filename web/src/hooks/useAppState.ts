import { useCallback, useState } from 'react'
import { buildSensitivityGrid, mockAnalyze, mockGenerate } from '../mock/engine'
import type {
  AnalyzeFailure,
  AppPhase,
  AssumptionMeta,
  Assumptions,
  CompanySnapshot,
  ComparisonSlot,
  ModelResults,
  SensitivityCell,
} from '../types'

export function useAppState() {
  const [phase, setPhase] = useState<AppPhase>('empty')
  const [tickerInput, setTickerInput] = useState('')
  const [error, setError] = useState<AnalyzeFailure | null>(null)
  const [snapshot, setSnapshot] = useState<CompanySnapshot | null>(null)
  const [assumptions, setAssumptions] = useState<Assumptions | null>(null)
  const [assumptionMeta, setAssumptionMeta] = useState<AssumptionMeta[]>([])
  const [results, setResults] = useState<ModelResults | null>(null)
  const [sensitivity, setSensitivity] = useState<SensitivityCell[]>([])
  const [generating, setGenerating] = useState(false)
  const [savedModels, setSavedModels] = useState<ComparisonSlot[]>([])
  const [showReport, setShowReport] = useState(false)
  const [toast, setToast] = useState<string | null>(null)

  const showToast = useCallback((msg: string) => {
    setToast(msg)
    window.setTimeout(() => setToast(null), 2800)
  }, [])

  const analyze = useCallback(async (raw: string) => {
    setPhase('loading')
    setError(null)
    setResults(null)
    setSensitivity([])
    setShowReport(false)
    setSnapshot(null)
    setAssumptions(null)

    const result = await mockAnalyze(raw)
    if (!result.ok) {
      setError(result)
      setPhase('error')
      return
    }

    setSnapshot(result.snapshot)
    setAssumptions(result.assumptions)
    setAssumptionMeta(result.assumptionMeta)
    setTickerInput(result.snapshot.ticker)
    setPhase('ready')
  }, [])

  const updateAssumption = useCallback(<K extends keyof Assumptions>(key: K, value: Assumptions[K]) => {
    setAssumptions((prev) => (prev ? { ...prev, [key]: value } : prev))
  }, [])

  const generate = useCallback(async () => {
    if (!snapshot || !assumptions) return
    setGenerating(true)
    const r = await mockGenerate(snapshot.ticker, assumptions, snapshot.companyName)
    const grid = buildSensitivityGrid(assumptions, r.moic)
    setResults(r)
    setSensitivity(grid)
    setPhase('generated')
    setGenerating(false)

    const slot: ComparisonSlot = {
      id: `${snapshot.ticker}-${Date.now()}`,
      label: `${snapshot.ticker} · ${assumptions.leverageMultiple.toFixed(1)}x lev`,
      ticker: snapshot.ticker,
      results: r,
      assumptions: { ...assumptions },
    }
    setSavedModels((prev) => [slot, ...prev].slice(0, 8))
  }, [snapshot, assumptions])

  const downloadExcel = useCallback(() => {
    if (!snapshot) return
    // Demo placeholder file — not a real workbook
    const blob = new Blob(
      [`AIO LBO demo placeholder for ${snapshot.ticker}. Connect the Python backend to generate a real .xlsx.`],
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' },
    )
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `LBO_${snapshot.ticker}_demo.txt`
    a.click()
    URL.revokeObjectURL(url)
    showToast('Demo download — not a real Excel model yet')
  }, [snapshot, showToast])

  return {
    phase,
    tickerInput,
    setTickerInput,
    error,
    snapshot,
    assumptions,
    assumptionMeta,
    results,
    sensitivity,
    generating,
    savedModels,
    showReport,
    setShowReport,
    toast,
    showToast,
    analyze,
    updateAssumption,
    generate,
    downloadExcel,
  }
}
