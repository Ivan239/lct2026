import { Link } from 'react-router-dom'
import { Button } from '@/shared/components/Button/Button'
import styles from './Cta.module.scss'

export function Cta() {
  return (
    <section className={styles.section}>
      <div className={styles.box}>
        <h2>Посмотрите на реальную генерацию</h2>
        <p>Выберите шаблон, вставьте бриф о продукте — и получите готовую презентацию за секунды.</p>
        <Link to="/demo">
          <Button>Перейти к демо</Button>
        </Link>
      </div>
    </section>
  )
}
