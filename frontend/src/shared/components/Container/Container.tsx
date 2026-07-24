import type { ReactNode } from 'react'
import styles from './Container.module.scss'

export function Container({ children }: { children: ReactNode }) {
  return <div className={styles.container}>{children}</div>
}
