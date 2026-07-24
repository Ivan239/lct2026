import { Container } from '@/shared/components/Container/Container'
import { Hero } from './components/Hero'
import { Problem } from './components/Problem'
import { HowItWorks } from './components/HowItWorks'
import { Differentiator } from './components/Differentiator'
import { Cta } from './components/Cta'

export function LandingPage() {
  return (
    <Container>
      <Hero />
      <Problem />
      <HowItWorks />
      <Differentiator />
      <Cta />
    </Container>
  )
}
