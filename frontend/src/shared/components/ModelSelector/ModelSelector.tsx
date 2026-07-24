import { useEffect, useState } from 'react'
import { fetchModels } from '@/shared/api/client'
import styles from './ModelSelector.module.scss'

interface ModelSelectorProps {
  value: string | null
  onChange: (model: string | null) => void
  label: string
}

export function ModelSelector({ value, onChange, label }: ModelSelectorProps) {
  const [models, setModels] = useState<string[]>([])

  useEffect(() => {
    fetchModels()
      .then((res) => setModels(res.models))
      .catch(() => setModels([]))
  }, [])

  return (
    <label className={styles.wrap}>
      <span className={styles.label}>{label}</span>
      <select
        className={styles.select}
        value={value ?? ''}
        onChange={(e) => onChange(e.target.value || null)}
      >
        <option value="">Авто (по умолчанию)</option>
        {models.map((m) => (
          <option key={m} value={m}>
            {m}
          </option>
        ))}
      </select>
    </label>
  )
}
