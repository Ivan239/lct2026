import type { GenerateResponse, VariantsResponse } from '@/shared/api/types'

/** История генераций — в localStorage этого браузера. Хранятся ответы сервера
 * (ссылки на превью и файлы, план, находки аудита); сами файлы лежат на сервере.
 * localStorage может быть недоступен (приватное окно, запрет сайта) — тогда
 * история просто пуста, страница работает как раньше. */

const KEY = 'slidegen:history'
const LIMIT = 30

export interface HistoryEntry {
  id: string
  createdAt: string
  templateName: string
  title: string
  single?: GenerateResponse
  variants?: VariantsResponse
}

export function loadHistory(): HistoryEntry[] {
  try {
    const raw = localStorage.getItem(KEY)
    const parsed = raw ? (JSON.parse(raw) as HistoryEntry[]) : []
    return Array.isArray(parsed) ? parsed : []
  } catch {
    return []
  }
}

function store(entries: HistoryEntry[]): HistoryEntry[] {
  const trimmed = entries.slice(0, LIMIT)
  try {
    localStorage.setItem(KEY, JSON.stringify(trimmed))
  } catch {
    /* переполнение или запрет — история не сохранится, генерация не страдает */
  }
  return trimmed
}

function titleOf(response: GenerateResponse | undefined): string {
  return response?.plan?.[0]?.title || 'Без названия'
}

export function entryFromSingle(response: GenerateResponse, templateName: string): HistoryEntry {
  return {
    id: response.generation_id,
    createdAt: new Date().toISOString(),
    templateName,
    title: titleOf(response),
    single: response,
  }
}

export function entryFromVariants(response: VariantsResponse, templateName: string): HistoryEntry {
  const first = response.variants.find((v) => !v.error)
  return {
    id: response.variants.map((v) => v.generation_id ?? v.variant).join('+'),
    createdAt: new Date().toISOString(),
    templateName,
    title: titleOf(first),
    variants: response,
  }
}

export function addToHistory(entry: HistoryEntry): HistoryEntry[] {
  return store([entry, ...loadHistory().filter((e) => e.id !== entry.id)])
}

/** После «Исправить выбранные» в истории хранится исправленная колода. */
export function updateHistory(id: string, patch: Partial<HistoryEntry>): HistoryEntry[] {
  return store(loadHistory().map((e) => (e.id === id ? { ...e, ...patch } : e)))
}

export function removeFromHistory(id: string): HistoryEntry[] {
  return store(loadHistory().filter((e) => e.id !== id))
}
