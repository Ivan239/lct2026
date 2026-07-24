import { Link } from 'react-router-dom'
import { Button } from '@/shared/components/Button/Button'
import styles from './Hero.module.scss'

export function Hero() {
  return (
    <section className={styles.hero}>
      <p className={styles.eyebrow}>Автоматизация корпоративных презентаций</p>
      <h1 className={styles.title}>
        Ваш фирменный шаблон.
        <br />
        Новый контент за минуты, а не часы.
      </h1>
      <p className={styles.subtitle}>
        СлайдоГен разбирает существующий PPTX-шаблон компании, извлекает дизайн-систему и типовые
        паттерны слайдов, а затем сам собирает новую презентацию из вашего контента — без участия
        дизайнера и без ручной вёрстки.
      </p>
      <div className={styles.actions}>
        <Link to="/demo">
          <Button>Попробовать демо</Button>
        </Link>
      </div>
    </section>
  )
}
