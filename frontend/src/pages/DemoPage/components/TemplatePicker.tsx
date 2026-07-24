import { useRef, useState } from 'react'
import type { TemplateSummary } from '@/shared/api/types'
import { resolveUrl } from '@/shared/api/client'
import { Spinner } from '@/shared/components/Spinner/Spinner'
import styles from './TemplatePicker.module.scss'

interface TemplatePickerProps {
  templates: TemplateSummary[]
  selectedId: string | null
  onSelect: (id: string) => void
  onUpload: (file: File) => void
  onDelete: (id: string) => void
  uploading: boolean
  uploadLabel?: string
}

export function TemplatePicker({
  templates,
  selectedId,
  onSelect,
  onUpload,
  onDelete,
  uploading,
  uploadLabel,
}: TemplatePickerProps) {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const [confirmId, setConfirmId] = useState<string | null>(null)

  return (
    <div className={styles.grid}>
      {templates.map((template) => (
        <div key={template.id} className={styles.cardWrap}>
          <button
            type="button"
            className={`${styles.card} ${selectedId === template.id ? styles.cardActive : ''}`}
            onClick={() => onSelect(template.id)}
          >
            {template.slides[0] && (
              <img className={styles.thumb} src={resolveUrl(template.slides[0])} alt={template.name} />
            )}
            <span className={styles.name}>{template.name}</span>
            <span className={styles.meta}>
              {new Set(Object.values(template.archetypes)).size} архетипов ·{' '}
              {Object.keys(template.archetypes).length} слайдов
            </span>
          </button>

          {!template.is_preset && (
            <button
              type="button"
              className={styles.deleteBtn}
              title="Удалить шаблон"
              onClick={(e) => {
                e.stopPropagation()
                setConfirmId(template.id)
              }}
            >
              ×
            </button>
          )}

          {confirmId === template.id && (
            <div className={styles.confirm} onClick={(e) => e.stopPropagation()}>
              <span>Удалить «{template.name}»?</span>
              <div className={styles.confirmActions}>
                <button
                  type="button"
                  className={styles.confirmDelete}
                  onClick={() => {
                    onDelete(template.id)
                    setConfirmId(null)
                  }}
                >
                  Удалить
                </button>
                <button type="button" className={styles.confirmCancel} onClick={() => setConfirmId(null)}>
                  Отмена
                </button>
              </div>
            </div>
          )}
        </div>
      ))}

      <button
        type="button"
        className={styles.uploadCard}
        onClick={() => fileInputRef.current?.click()}
        disabled={uploading}
      >
        {uploading ? (
          <Spinner label={uploadLabel ?? 'Разбираем шаблон...'} />
        ) : (
          <>
            <span className={styles.uploadIcon}>+</span>
            <span>Загрузить свой .pptx</span>
          </>
        )}
      </button>
      <input
        ref={fileInputRef}
        type="file"
        accept=".pptx"
        hidden
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) onUpload(file)
          e.target.value = ''
        }}
      />
    </div>
  )
}
