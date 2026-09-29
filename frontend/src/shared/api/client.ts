import type { BalanceEntry, GenerateResponse, TemplateSummary } from './types'

// В сборке (docker compose) фронт и API за одним nginx: путь относительный.
// В dev-режиме Vite бэкенд поднят отдельно на 8000.
const API_BASE_URL =
  import.meta.env.VITE_API_BASE_URL ?? (import.meta.env.DEV ? 'http://localhost:8000' : '')

export function resolveUrl(path: string): string {
  return path.startsWith('http') ? path : `${API_BASE_URL}${path}`
}

async function request<T>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(`${API_BASE_URL}${path}`, init)
  if (!res.ok) {
    const body = await res.json().catch(() => null)
    throw new Error(body?.detail ?? `Запрос ${path} завершился с ошибкой ${res.status}`)
  }
  return res.json() as Promise<T>
}

export function fetchTemplates(): Promise<TemplateSummary[]> {
  return request('/api/templates')
}

export function fetchModels(): Promise<{ models: string[] }> {
  return request('/api/models')
}

export function fetchBalance(): Promise<{ balance: BalanceEntry[] }> {
  return request('/api/balance')
}

export interface GenerationProgress {
  active: boolean
  stage: string
  done: number
  total: number
}

export function fetchProgress(): Promise<GenerationProgress> {
  return request('/api/progress')
}

export function uploadTemplate(file: File, model?: string | null): Promise<TemplateSummary> {
  const form = new FormData()
  form.append('file', file)
  if (model) form.append('model', model)
  return request('/api/templates', { method: 'POST', body: form })
}

export function deleteTemplate(templateId: string): Promise<{ deleted: string }> {
  return request(`/api/templates/${templateId}`, { method: 'DELETE' })
}

export function generatePresentation(
  templateId: string,
  brief: string,
  model?: string | null,
): Promise<GenerateResponse> {
  return request('/api/generate', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ template_id: templateId, brief, model: model || null }),
  })
}

/** Generation from a content package: a .zip of brief.md (+ package.json,
 * facts.md, data/*.csv, images/*). The deck's numbers are checked against the
 * package's; mismatches come back in `warnings`. */
export function generateFromPackage(
  templateId: string,
  file: File,
  model?: string | null,
): Promise<GenerateResponse> {
  const form = new FormData()
  form.append('template_id', templateId)
  form.append('file', file)
  if (model) form.append('model', model)
  return request('/api/generate/package', { method: 'POST', body: form })
}

/** Исправить выбранные находки аудита: текст этих слайдов пишется заново с
 * теми же ограничениями, колода пересобирается и проходит аудит заново. */
export function fixSelected(generationId: string, slides: number[]): Promise<GenerateResponse> {
  return request(`/api/generate/${generationId}/fix`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ slides }),
  })
}
