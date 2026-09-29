import type { HistoryEntry } from '@/shared/history/history'
import styles from './HistoryPanel.module.scss'

function when(iso: string): string {
  const d = new Date(iso)
  return d.toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' })
}

/** Прошлые генерации этого браузера: открыть результат снова или убрать из списка. */
export function HistoryPanel({
  entries,
  currentId,
  onOpen,
  onRemove,
}: {
  entries: HistoryEntry[]
  currentId: string | null
  onOpen: (entry: HistoryEntry) => void
  onRemove: (id: string) => void
}) {
  if (entries.length === 0) return null
  return (
    <details className={styles.wrap}>
      <summary className={styles.summary}>История генераций ({entries.length})</summary>
      <ul className={styles.list}>
        {entries.map((entry) => (
          <li key={entry.id} className={entry.id === currentId ? styles.current : undefined}>
            <button className={styles.open} onClick={() => onOpen(entry)}>
              <span className={styles.title}>{entry.title}</span>
              <span className={styles.meta}>
                {when(entry.createdAt)} · {entry.templateName} ·{' '}
                {entry.variants ? 'три варианта' : `${entry.single?.slides.length ?? 0} слайдов`}
              </span>
            </button>
            <button
              className={styles.remove}
              onClick={() => onRemove(entry.id)}
              aria-label="Убрать из истории"
              title="Убрать из истории"
            >
              ×
            </button>
          </li>
        ))}
      </ul>
    </details>
  )
}
