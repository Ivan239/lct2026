import { useEffect, useState } from 'react'
import { Container } from '@/shared/components/Container/Container'
import { Button } from '@/shared/components/Button/Button'
import { Spinner } from '@/shared/components/Spinner/Spinner'
import { ModelSelector } from '@/shared/components/ModelSelector/ModelSelector'
import { BalanceBadge } from '@/shared/components/BalanceBadge/BalanceBadge'
import {
  deleteTemplate,
  fetchBalance,
  fetchProgress,
  fetchTemplates,
  generatePresentation,
  uploadTemplate,
} from '@/shared/api/client'
import type { BalanceEntry, GenerateResponse, TemplateSummary } from '@/shared/api/types'
import { TemplatePicker } from './components/TemplatePicker'
import { BriefEditor } from './components/BriefEditor'
import { GenerationResult } from './components/GenerationResult'
import { EXAMPLE_BRIEF } from './exampleBrief'
import styles from './DemoPage.module.scss'

function startProgressPolling(setLabel: (label: string) => void, fallback: string): number {
  setLabel(fallback)
  return window.setInterval(async () => {
    try {
      const p = await fetchProgress()
      if (p.active && p.stage) {
        const counter = p.total > 0 ? ` (${p.done + 1}/${p.total})` : ''
        setLabel(`${p.stage}${counter}`)
      }
    } catch {
      /* прогресс — только украшение, ошибки поллинга молча пропускаем */
    }
  }, 1200)
}

function usePersistedModel(storageKey: string): [string | null, (value: string | null) => void] {
  const [value, setValue] = useState<string | null>(() => localStorage.getItem(storageKey))

  function update(next: string | null) {
    setValue(next)
    if (next) localStorage.setItem(storageKey, next)
    else localStorage.removeItem(storageKey)
  }

  return [value, update]
}

export function DemoPage() {
  const [templates, setTemplates] = useState<TemplateSummary[]>([])
  const [selectedId, setSelectedId] = useState<string | null>(null)
  const [brief, setBrief] = useState(EXAMPLE_BRIEF)
  const [uploading, setUploading] = useState(false)
  const [generating, setGenerating] = useState(false)
  const [result, setResult] = useState<GenerateResponse | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [balance, setBalance] = useState<BalanceEntry[] | null>(null)
  const [progressLabel, setProgressLabel] = useState('Собираем презентацию...')
  const [uploadLabel, setUploadLabel] = useState('Разбираем шаблон...')
  const [templateModel, setTemplateModel] = usePersistedModel('slidegen:templateModel')
  const [generateModel, setGenerateModel] = usePersistedModel('slidegen:generateModel')

  useEffect(() => {
    fetchTemplates()
      .then((list) => {
        setTemplates(list)
        if (list.length > 0) setSelectedId(list[0].id)
      })
      .catch((e: Error) => setError(e.message))
    fetchBalance()
      .then((res) => setBalance(res.balance))
      .catch(() => {})
  }, [])

  async function handleUpload(file: File) {
    setUploading(true)
    setError(null)
    const poll = startProgressPolling(setUploadLabel, 'Разбираем шаблон...')
    try {
      const template = await uploadTemplate(file, templateModel)
      setTemplates((prev) => [...prev, template])
      setSelectedId(template.id)
      if (template.balance) setBalance(template.balance)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      window.clearInterval(poll)
      setUploading(false)
    }
  }

  async function handleDelete(id: string) {
    setError(null)
    try {
      await deleteTemplate(id)
      setTemplates((prev) => prev.filter((t) => t.id !== id))
      setSelectedId((prev) => (prev === id ? null : prev))
    } catch (e) {
      setError((e as Error).message)
    }
  }

  async function handleGenerate() {
    if (!selectedId) return
    setGenerating(true)
    setError(null)
    setResult(null)
    const poll = startProgressPolling(setProgressLabel, 'Собираем презентацию...')
    try {
      const response = await generatePresentation(selectedId, brief, generateModel)
      setResult(response)
      if (response.balance) setBalance(response.balance)
    } catch (e) {
      setError((e as Error).message)
    } finally {
      window.clearInterval(poll)
      setGenerating(false)
    }
  }

  return (
    <Container>
      <div className={styles.page}>
        <div className={styles.headerRow}>
          <h1 className={styles.heading}>Демо: от брифа до презентации</h1>
          <BalanceBadge balance={balance} />
        </div>

        <section className={styles.section}>
          <div className={styles.sectionHead}>
            <h2 className={styles.sectionTitle}>1. Выберите шаблон</h2>
            <ModelSelector label="Модель для разбора:" value={templateModel} onChange={setTemplateModel} />
          </div>
          <TemplatePicker
            templates={templates}
            selectedId={selectedId}
            onSelect={setSelectedId}
            onUpload={handleUpload}
            onDelete={handleDelete}
            uploading={uploading}
            uploadLabel={uploadLabel}
          />
        </section>

        <section className={styles.section}>
          <h2 className={styles.sectionTitle}>2. Опишите продукт</h2>
          <BriefEditor value={brief} onChange={setBrief} />
        </section>

        <div className={styles.generateRow}>
          <ModelSelector label="Модель для генерации:" value={generateModel} onChange={setGenerateModel} />
          <Button onClick={handleGenerate} disabled={!selectedId || generating}>
            {generating ? <Spinner label={progressLabel} /> : 'Сгенерировать презентацию'}
          </Button>
        </div>

        {error && <div className={styles.error}>{error}</div>}

        {result && <GenerationResult result={result} />}
      </div>
    </Container>
  )
}
