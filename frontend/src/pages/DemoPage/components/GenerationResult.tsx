import type { GenerateResponse } from '@/shared/api/types'
import { resolveUrl } from '@/shared/api/client'
import { Button } from '@/shared/components/Button/Button'
import { SlideGallery } from './SlideGallery'
import styles from './GenerationResult.module.scss'

export function GenerationResult({ result }: { result: GenerateResponse }) {
  return (
    <div className={styles.wrap}>
      <div className={styles.head}>
        <h3>План сборки</h3>
        <a href={resolveUrl(result.download_url)} download>
          <Button>Скачать .pptx</Button>
        </a>
      </div>

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

      {(result.warnings?.length ?? 0) > 0 && (
        <div className={styles.warnings}>
          <strong>Стоит взглянуть:</strong>
          <ul>
            {result.warnings!.map((w, i) => (
              <li key={i}>
                Слайд {w.slide}: {w.details}
              </li>
            ))}
          </ul>
        </div>
      )}

      <h3 className={styles.slidesHeading}>Результат</h3>
      <SlideGallery slides={result.slides} />
    </div>
  )
}
