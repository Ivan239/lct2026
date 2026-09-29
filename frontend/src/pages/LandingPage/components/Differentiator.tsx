import styles from './Differentiator.module.scss'

const ROWS = [
  {
    label: 'Дизайн',
    ours: 'Учится на вашем реальном PPTX-шаблоне',
    others: 'Готовые generic-темы, без учёта бренда',
  },
  {
    label: 'Стиль',
    ours: 'Точное воспроизведение фирменной айдентики',
    others: 'Собственный формат карточек, свой редактор',
  },
  {
    label: 'Результат',
    ours: 'Стандартный .pptx, открывается где угодно',
    others: 'Нужно экспортировать/адаптировать из своего формата',
  },
  {
    label: 'LLM-стек',
    ours: 'Открытые модели до 35B (Qwen3-32B, Apache 2.0) — свой сервер или российский провайдер',
    others: 'Закрытые зарубежные модели по подписке',
  },
]

export function Differentiator() {
  return (
    <section className={styles.section}>
      <h2 className={styles.heading}>Чем это отличается от Gamma и подобных</h2>
      <div className={styles.table}>
        <div className={styles.row}>
          <div className={styles.cellLabel} />
          <div className={styles.cellHeadOurs}>СлайдоГен</div>
          <div className={styles.cellHeadOthers}>Обычные AI-тулы</div>
        </div>
        {ROWS.map((row) => (
          <div key={row.label} className={styles.row}>
            <div className={styles.cellLabel}>{row.label}</div>
            <div className={styles.cellOurs}>{row.ours}</div>
            <div className={styles.cellOthers}>{row.others}</div>
          </div>
        ))}
      </div>
    </section>
  )
}
