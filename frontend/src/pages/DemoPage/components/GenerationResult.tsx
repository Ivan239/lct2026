import { useState } from 'react'
import type { GenerateResponse } from '@/shared/api/types'
import { fixSelected, resolveUrl } from '@/shared/api/client'
import { Button } from '@/shared/components/Button/Button'
import { SlideGallery } from './SlideGallery'
import styles from './GenerationResult.module.scss'

const FIX_LABEL: Record<string, string> = {
  numbers_not_in_source: 'переписать без цифр, которых нет в источнике',
  sparse: 'переписать текст слайда',
}

export function GenerationResult({
  result,
  onFixed,
}: {
  result: GenerateResponse
  onFixed: (fixed: GenerateResponse) => void
}) {
  const warnings = result.warnings ?? []
  // По умолчанию отмечено всё найденное: пользователь снимает то, что считает
  // нормой (например, выведенную им самим цифру).
  const [chosen, setChosen] = useState<Set<number>>(() => new Set(warnings.map((w) => w.slide)))
  const [fixing, setFixing] = useState(false)
  const [fixError, setFixError] = useState<string | null>(null)

  function toggle(slide: number) {
    setChosen((prev) => {
      const next = new Set(prev)
      if (next.has(slide)) next.delete(slide)
      else next.add(slide)
      return next
    })
  }

  async function handleFix() {
    setFixing(true)
    setFixError(null)
    try {
      onFixed(await fixSelected(result.generation_id, [...chosen]))
    } catch (e) {
      setFixError((e as Error).message)
    } finally {
      setFixing(false)
    }
  }

  return (
    <div className={styles.wrap}>
      <div className={styles.head}>
        <h3>План сборки</h3>
        <div className={styles.downloads}>
          <a href={resolveUrl(result.download_url)} download target="_blank" rel="noopener noreferrer">
            <Button>Скачать .pptx</Button>
          </a>
          {result.pdf_url && (
            <a href={resolveUrl(result.pdf_url)} target="_blank" rel="noopener noreferrer">
              <Button>.pdf</Button>
            </a>
          )}
          {result.html_url && (
            <a href={resolveUrl(result.html_url)} target="_blank" rel="noopener noreferrer">
              <Button>.html</Button>
            </a>
          )}
        </div>
      </div>

      {(result.substituted_fonts?.length ?? 0) > 0 && (
        <p className={styles.packageLine}>
          В превью, .pdf и .html шрифты {result.substituted_fonts!.join(', ')} заменены на {result.fallback_font ?? 'Arial'} —
          их нет на сервере рендера. В .pptx шрифты оригинальные.
        </p>
      )}

      {result.package && (
        <p className={styles.packageLine}>
          По контент-пакету «{result.package.title ?? 'без названия'}»: фактов {result.package.facts},
          цифр для сверки {result.package.numbers}
          {result.package.tables.length > 0 && <>, таблицы: {result.package.tables.join(', ')}</>}
        </p>
      )}

      <ol className={styles.plan}>
        {result.plan.map((item, i) => (
          <li key={i}>
            <span className={styles.archetype}>{item.archetype}</span>
            <span className={styles.planTitle}>{item.title}</span>
          </li>
        ))}
      </ol>

      {result.skipped.length > 0 && (
        <div className={styles.skipped}>
          <strong>Не поместилось в шаблон:</strong>
          <ul>
            {result.skipped.map((item, i) => (
              <li key={i}>
                {item.type}: {item.title}
              </li>
            ))}
          </ul>
        </div>
      )}

      {(result.fixed_slides?.length ?? 0) > 0 && (
        <p className={styles.fixedLine}>
          Исправлено слайдов: {result.fixed_slides!.join(', ')}. Колода пересобрана и проверена заново.
        </p>
      )}
      {(result.not_fixed?.length ?? 0) > 0 && (
        <p className={styles.packageLine}>
          Не исправлено: {result.not_fixed!.map((x) => `слайд ${x.slide} — ${x.reason}`).join('; ')}
        </p>
      )}

      {warnings.length > 0 && (
        <div className={styles.warnings}>
          <strong>Аудит нашёл — отметьте, что исправить:</strong>
          <ul className={styles.findings}>
            {warnings.map((w, i) => (
              <li key={i}>
                <label>
                  <input
                    type="checkbox"
                    checked={chosen.has(w.slide)}
                    onChange={() => toggle(w.slide)}
                    disabled={fixing}
                  />
                  <span>
                    Слайд {w.slide}: {w.details}
                    {FIX_LABEL[w.kind] && <em> → {FIX_LABEL[w.kind]}</em>}
                  </span>
                </label>
              </li>
            ))}
          </ul>
          <div className={styles.fixRow}>
            <Button onClick={handleFix} disabled={fixing || chosen.size === 0}>
              {fixing ? 'Исправляем…' : `Исправить выбранные (${chosen.size})`}
            </Button>
            {fixError && <span className={styles.fixError}>{fixError}</span>}
          </div>
        </div>
      )}

      <h3 className={styles.slidesHeading}>Результат</h3>
      <SlideGallery slides={result.slides} />
    </div>
  )
}
