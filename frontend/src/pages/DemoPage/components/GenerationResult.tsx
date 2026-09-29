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
        <div className={styles.downloads}>
          <a href={resolveUrl(result.download_url)} download>
            <Button>Скачать .pptx</Button>
          </a>
          {result.pdf_url && (
            <a href={resolveUrl(result.pdf_url)} download>
              <Button>.pdf</Button>
            </a>
          )}
          {result.html_url && (
            <a href={resolveUrl(result.html_url)} download>
              <Button>.html</Button>
            </a>
          )}
        </div>
      </div>

      {(result.substituted_fonts?.length ?? 0) > 0 && (
        <p className={styles.packageLine}>
          В превью, .pdf и .html шрифты {result.substituted_fonts!.join(', ')} заменены на Arial —
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
