import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import { UIProvider } from './components/ui'
import { ErrorBoundary } from './components/ErrorBoundary'
import './styles.css'

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <ErrorBoundary>
      <UIProvider>
        <App />
      </UIProvider>
    </ErrorBoundary>
  </React.StrictMode>,
)
