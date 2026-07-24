import styles from './Problem.module.scss'

const POINTS = [
  {
    title: 'Часы ручной вёрстки',
    text: 'Каждая новая презентация требует, чтобы кто-то вручную подгонял контент под фирменный шаблон.',
  },
  {
    title: 'Дизайнеры — узкое место',
    text: 'Типовая, повторяющаяся задача отнимает время у дизайнеров вместо содержательной работы.',
  },
  {
    title: 'Разнобой в оформлении',
    text: 'У разных авторов результат получается визуально несогласованным — единого стандарта нет.',
  },
]

export function Problem() {
  return (
    <section className={styles.section}>
      <h2 className={styles.heading}>Знакомая проблема</h2>
      <div className={styles.grid}>
        {POINTS.map((p) => (
          <div key={p.title} className={styles.card}>
            <h3>{p.title}</h3>
            <p>{p.text}</p>
          </div>
        ))}
      </div>
    </section>
  )
}
