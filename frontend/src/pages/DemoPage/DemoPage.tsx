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
  generateFromPackage,
  generatePresentation,
  generateVariants,
  generateVariantsFromPackage,
  uploadTemplate,
} from '@/shared/api/client'
import type { BalanceEntry, GenerateResponse, TemplateSummary, VariantsResponse } from '@/shared/api/types'
import { TemplatePicker } from './components/TemplatePicker'
import { BriefEditor } from './components/BriefEditor'
import { PackageInput } from './components/PackageInput'
import { GenerationResult } from './components/GenerationResult'
import { VariantsResult } from './components/VariantsResult'
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
  const [packageFile, setPackageFile] = useState<File | null>(null)
  const [uploading, setUploading] = useState(false)
  const [generating, setGenerating] = useState(false)
  const [result, setResult] = useState<GenerateResponse | null>(null)
  const [variants, setVariants] = useState<VariantsResponse | null>(null)
  const [threeVariants, setThreeVariants] = useState(true)
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
    setVariants(null)
    const poll = startProgressPolling(
      setProgressLabel,
      threeVariants ? 'Собираем три варианта...' : 'Собираем презентацию...',
    )
    try {
      if (threeVariants) {
        const response = packageFile
          ? await generateVariantsFromPackage(selectedId, packageFile, generateModel)
          : await generateVariants(selectedId, brief, generateModel)
        setVariants(response)
      } else {
        const response = packageFile
          ? await generateFromPackage(selectedId, packageFile, generateModel)
          : await generatePresentation(selectedId, brief, generateModel)
        setResult(response)
        if (response.balance) setBalance(response.balance)
      }
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
          <div className={packageFile ? styles.briefDisabled : undefined}>
            <BriefEditor value={brief} onChange={setBrief} disabled={packageFile !== null} />
          </div>
          <PackageInput file={packageFile} onChange={setPackageFile} disabled={generating} />
        </section>

        <div className={styles.generateRow}>
          <ModelSelector label="Модель для генерации:" value={generateModel} onChange={setGenerateModel} />
          <label className={styles.variantsToggle}>
            <input
              type="checkbox"
              checked={threeVariants}
              onChange={(e) => setThreeVariants(e.target.checked)}
              disabled={generating}
            />
            Три варианта вёрстки (компактный, визуальный, подробный)
          </label>
          <Button onClick={handleGenerate} disabled={!selectedId || generating}>
            {generating ? <Spinner label={progressLabel} /> : 'Сгенерировать презентацию'}
          </Button>
        </div>

        {error && <div className={styles.error}>{error}</div>}

        {variants && <VariantsResult response={variants} onBalance={setBalance} />}

        {result && (
          <GenerationResult
            key={result.generation_id}
            result={result}
            onFixed={(fixed) => {
              setResult(fixed)
              if (fixed.balance) setBalance(fixed.balance)
            }}
          />
        )}
      </div>
    </Container>
  )
}
