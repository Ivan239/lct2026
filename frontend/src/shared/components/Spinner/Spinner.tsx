import styles from './Spinner.module.scss'

export function Spinner({ label }: { label?: string }) {
  return (
    <div className={styles.wrap}>
      <span className={styles.spinner} />
      {label && <span className={styles.label}>{label}</span>}
    </div>
  )
}
