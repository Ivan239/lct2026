import { useState } from 'react'
import type { BalanceEntry, GenerateResponse, VariantsResponse } from '@/shared/api/types'
import { GenerationResult } from './GenerationResult'
import styles from './VariantsResult.module.scss'

/** Три варианта вёрстки одной колоды — вкладками; у каждого свои превью,
 * скачивания и находки аудита с исправлением. */
export function VariantsResult({
  response,
  onBalance,
  onChange,
}: {
  response: VariantsResponse
  onBalance: (balance: BalanceEntry[]) => void
  onChange?: (next: VariantsResponse) => void
}) {
  const [items, setItems] = useState(response.variants)
  const [active, setActive] = useState(0)
  const current = items[active]

  function replace(index: number, fixed: GenerateResponse) {
    const next = items.map((item, i) =>
      i === index ? { ...fixed, variant: item.variant, variant_title: item.variant_title } : item,
    )
    setItems(next)
    onChange?.({ ...response, variants: next })
    if (fixed.balance) onBalance(fixed.balance)
  }

  return (
    <div className={styles.wrap}>
      <p className={styles.summary}>
        Три варианта вёрстки собраны за {response.seconds} с на одном шаблоне и одном контенте.
        Различается план колоды: число слайдов, состав и порядок; правила шаблона у всех одни.
      </p>
      <div className={styles.tabs} role="tablist">
        {items.map((item, i) => (
          <button
            key={item.variant ?? i}
            role="tab"
            aria-selected={i === active}
            className={i === active ? styles.tabActive : styles.tab}
            onClick={() => setActive(i)}
          >
            {item.variant_title ?? item.variant}
            <span className={styles.count}>
              {item.error ? 'ошибка' : `${item.slides.length} слайдов`}
            </span>
          </button>
        ))}
      </div>
      {current.error ? (
        <div className={styles.error}>Вариант не собрался: {current.error}</div>
      ) : (
        <GenerationResult
          key={current.generation_id}
          result={current}
          onFixed={(fixed) => replace(active, fixed)}
        />
      )}
    </div>
  )
}
