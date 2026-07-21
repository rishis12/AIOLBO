import { useCallback, useState } from 'react'
import type { LlmProvider, SessionKeys } from '../types'

const STORAGE_KEY = 'aio-lbo-session-keys'

const DEFAULTS: SessionKeys = {
  llmProvider: 'gemini',
  llmApiKey: '',
}

function read(): SessionKeys {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY)
    if (!raw) return { ...DEFAULTS }
    return { ...DEFAULTS, ...JSON.parse(raw) }
  } catch {
    return { ...DEFAULTS }
  }
}

export function useSessionKeys() {
  const [keys, setKeys] = useState<SessionKeys>(() => read())

  const save = useCallback((next: SessionKeys) => {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(next))
    setKeys(next)
  }, [])

  const clear = useCallback(() => {
    sessionStorage.removeItem(STORAGE_KEY)
    setKeys({ ...DEFAULTS })
  }, [])

  const hasLlmKey = Boolean(keys.llmApiKey.trim())

  return { keys, save, clear, hasLlmKey }
}

export type { LlmProvider }
