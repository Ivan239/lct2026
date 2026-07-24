import styles from './BriefEditor.module.scss'

interface BriefEditorProps {
  value: string
  onChange: (value: string) => void
}

export function BriefEditor({ value, onChange }: BriefEditorProps) {
  return (
    <textarea
      className={styles.textarea}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      spellCheck={false}
      placeholder="Опишите продукт: проблему, решение, преимущества, метрики, сравнение с конкурентами..."
    />
  )
}
