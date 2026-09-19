import styles from './BriefEditor.module.scss'

interface BriefEditorProps {
  value: string
  onChange: (value: string) => void
  /** A content package is chosen: the brief comes from it, this field is not used. */
  disabled?: boolean
}

export function BriefEditor({ value, onChange, disabled }: BriefEditorProps) {
  return (
    <textarea
      className={styles.textarea}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      spellCheck={false}
      disabled={disabled}
      title={disabled ? 'Бриф берётся из контент-пакета' : undefined}
      placeholder="Опишите продукт: проблему, решение, преимущества, метрики, сравнение с конкурентами..."
    />
  )
}
