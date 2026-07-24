import type { BalanceEntry } from '@/shared/api/types'
import styles from './BalanceBadge.module.scss'

function formatTokens(n: number): string {
  if (n >= 1000) return `${Math.round(n / 1000)}K`
  return String(n)
}

export function BalanceBadge({ balance }: { balance: BalanceEntry[] | null | undefined }) {
  if (!balance || balance.length === 0) return null

  return (
    <div className={styles.wrap} title="Остаток токенов GigaChat по тарифам">
      <span className={styles.dot} />
      {balance.map((entry) => (
        <span key={entry.model} className={styles.entry}>
          {entry.model}: <strong>{formatTokens(entry.tokens)}</strong>
        </span>
      ))}
    </div>
  )
}
