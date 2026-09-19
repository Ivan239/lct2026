import { useRef } from 'react'
import styles from './PackageInput.module.scss'

interface PackageInputProps {
  file: File | null
  onChange: (file: File | null) => void
  disabled?: boolean
}

/** A content package instead of the free-text brief: brief.md plus facts and
 * tables. The service checks every number on the slides against it. */
export function PackageInput({ file, onChange, disabled }: PackageInputProps) {
  const inputRef = useRef<HTMLInputElement>(null)

  return (
    <div className={styles.row}>
      {file ? (
        <div className={styles.chosen}>
          <span className={styles.badge}>Контент-пакет</span>
          <span className={styles.fileName}>{file.name}</span>
          <button
            type="button"
            className={styles.clear}
            onClick={() => onChange(null)}
            disabled={disabled}
            title="Убрать пакет и вернуться к брифу"
          >
            ×
          </button>
        </div>
      ) : (
        <button
          type="button"
          className={styles.pick}
          onClick={() => inputRef.current?.click()}
          disabled={disabled}
        >
          или загрузите контент-пакет .zip
        </button>
      )}
      <span className={styles.hint}>
        brief.md, facts.md, data/*.csv — цифры на слайдах сверяются с пакетом
      </span>
      <input
        ref={inputRef}
        type="file"
        accept=".zip"
        hidden
        onChange={(e) => {
          const picked = e.target.files?.[0]
          if (picked) onChange(picked)
          e.target.value = ''
        }}
      />
    </div>
  )
}
