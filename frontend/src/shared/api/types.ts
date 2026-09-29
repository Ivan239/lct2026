export type Archetype =
  | 'title'
  | 'section_divider'
  | 'bullet_list'
  | 'stats_kpi'
  | 'two_column_comparison'
  | 'image_caption'
  | 'quote'
  | 'agenda'
  | 'closing'
  | 'other'

export interface BalanceEntry {
  model: string
  tokens: number
}

export interface TemplateSummary {
  id: string
  name: string
  is_preset: boolean
  slides: string[]
  archetypes: Record<string, Archetype>
  balance?: BalanceEntry[] | null
}

export interface PlanItem {
  type: string
  title: string | null
  archetype: Archetype
}

export interface SkippedItem {
  type: string
  title: string | null
}

export interface GenerationWarning {
  slide: number
  kind: string
  details: string
}

/** What the service read from a content package (docs/CONTENT_PACKAGE.md). */
export interface PackageSummary {
  title: string | null
  purpose: string | null
  facts: number
  numbers: number
  tables: string[]
  images: number
}

export interface GenerateResponse {
  generation_id: string
  download_url: string
  pdf_url?: string
  html_url?: string
  substituted_fonts?: string[]
  slides: string[]
  plan: PlanItem[]
  skipped: SkippedItem[]
  warnings?: GenerationWarning[]
  balance?: BalanceEntry[] | null
  package?: PackageSummary
}
