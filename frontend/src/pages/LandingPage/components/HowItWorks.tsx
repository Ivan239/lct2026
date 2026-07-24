import styles from './HowItWorks.module.scss'

const STEPS = [
  {
    n: '01',
    title: 'Разбор шаблона',
    text: 'Читаем OOXML-структуру PPTX: тема, шрифты, геометрия фигур. Рендерим слайды и через vision-модель классифицируем архетипы: title, bullet-list, stats, comparison…',
  },
  {
    n: '02',
    title: 'Разбор контента',
    text: 'Сырой бриф о продукте превращается LLM в структурированные content-блоки того же словаря архетипов, что и у шаблона.',
  },
  {
    n: '03',
    title: 'Сопоставление',
    text: 'Каждый content-блок подбирает себе подходящий слайд-архетип из доступных в конкретном шаблоне.',
  },
  {
    n: '04',
    title: 'Генерация',
    text: 'Контент подставляется в структуру шаблона, стиль (шрифты, цвета, layout) остаётся нетронутым. На выходе — готовый .pptx.',
  },
]

export function HowItWorks() {
  return (
    <section className={styles.section}>
      <h2 className={styles.heading}>Как это работает</h2>
      <div className={styles.steps}>
        {STEPS.map((step, i) => (
          <div key={step.n} className={styles.step}>
            <div className={styles.stepHeader}>
              <span className={styles.number}>{step.n}</span>
              {i < STEPS.length - 1 && <span className={styles.connector} />}
            </div>
            <h3>{step.title}</h3>
            <p>{step.text}</p>
          </div>
        ))}
      </div>
    </section>
  )
}
