import { resolveUrl } from '@/shared/api/client'
import styles from './SlideGallery.module.scss'

export function SlideGallery({ slides }: { slides: string[] }) {
  return (
    <div className={styles.gallery}>
      {slides.map((slide, i) => (
        <img key={slide} className={styles.slide} src={resolveUrl(slide)} alt={`Слайд ${i + 1}`} />
      ))}
    </div>
  )
}
