import { useEffect, useState } from 'react'
import type { LlmProvider, SessionKeys } from '../types'

interface Props {
  open: boolean
  keys: SessionKeys
  onClose: () => void
  onSave: (keys: SessionKeys) => void
  onClear: () => void
}

export function SettingsModal({ open, keys, onClose, onSave, onClear }: Props) {
  const [draft, setDraft] = useState(keys)

  useEffect(() => {
    if (open) setDraft(keys)
  }, [open, keys])

  useEffect(() => {
    if (!open) return
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [open, onClose])

  if (!open) return null

  return (
    <div className="modal-backdrop" role="presentation" onClick={onClose}>
      <div
        className="modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="settings-title"
        onClick={(e) => e.stopPropagation()}
      >
        <h2 id="settings-title">Settings / API Keys</h2>
        <p className="modal-disclosure">
          Keys are stored in this browser session only and are never persisted server-side. Closing the
          tab clears them with the session.
        </p>

        <div className="form-stack">
          <label>
            <span className="field-label">SEC Contact Email (required)</span>
            <input
              className="input"
              type="email"
              value={draft.secEmail}
              onChange={(e) => setDraft({ ...draft, secEmail: e.target.value })}
              placeholder="you@company.com"
              autoComplete="off"
            />
          </label>

          <label>
            <span className="field-label">Twelve Data API Key</span>
            <input
              className="input"
              type="password"
              value={draft.twelveDataKey}
              onChange={(e) => setDraft({ ...draft, twelveDataKey: e.target.value })}
              placeholder="Required for live price"
              autoComplete="off"
            />
          </label>

          <label>
            <span className="field-label">LLM Provider</span>
            <select
              className="select"
              value={draft.llmProvider}
              onChange={(e) =>
                setDraft({ ...draft, llmProvider: e.target.value as LlmProvider })
              }
            >
              <option value="anthropic">Anthropic</option>
              <option value="openai">OpenAI</option>
              <option value="gemini">Gemini</option>
            </select>
          </label>

          <label>
            <span className="field-label">LLM API Key</span>
            <input
              className="input"
              type="password"
              value={draft.llmApiKey}
              onChange={(e) => setDraft({ ...draft, llmApiKey: e.target.value })}
              placeholder="BYOK — never stored server-side"
              autoComplete="off"
            />
          </label>
        </div>

        <div className="modal-actions">
          <button
            type="button"
            className="btn btn-secondary"
            onClick={() => {
              onClear()
              onClose()
            }}
          >
            Clear
          </button>
          <button type="button" className="btn btn-secondary" onClick={onClose}>
            Cancel
          </button>
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => {
              onSave(draft)
              onClose()
            }}
          >
            Save
          </button>
        </div>
      </div>
    </div>
  )
}
