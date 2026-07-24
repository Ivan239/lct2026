import { NavLink, Route, Routes } from 'react-router-dom'
import { LandingPage } from '@/pages/LandingPage/LandingPage'
import { DemoPage } from '@/pages/DemoPage/DemoPage'
import { Container } from '@/shared/components/Container/Container'
import styles from './App.module.scss'

export function App() {
  return (
    <>
      <header className={styles.header}>
        <Container>
          <div className={styles.headerInner}>
            <NavLink to="/" className={styles.logo}>
              Слайдо<span>Ген</span>
            </NavLink>
            <nav className={styles.nav}>
              <NavLink
                to="/"
                end
                className={({ isActive }) => (isActive ? styles.navLinkActive : styles.navLink)}
              >
                О продукте
              </NavLink>
              <NavLink
                to="/demo"
                className={({ isActive }) => (isActive ? styles.navLinkActive : styles.navLink)}
              >
                Демо
              </NavLink>
            </nav>
          </div>
        </Container>
      </header>

      <main className={styles.main}>
        <Routes>
          <Route path="/" element={<LandingPage />} />
          <Route path="/demo" element={<DemoPage />} />
        </Routes>
      </main>

      <footer className={styles.footer}>
        <Container>
          <p>СлайдоГен — учебный прототип. Генерация на GigaChat, российский LLM-стек.</p>
        </Container>
      </footer>
    </>
  )
}
