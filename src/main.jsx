import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import './index.css'

import { registerSW } from 'virtual:pwa-register'

const updateSW = registerSW({
  immediate: true,

  onRegisteredSW(swUrl, registration) {
    console.log('PROLI Marketing 360 registrada:', swUrl)

    if (registration) {
      registration.update()
    }
  },

  onNeedRefresh() {
    console.log('Nueva versión de PROLI disponible')
    updateSW(true)
  },

  onOfflineReady() {
    console.log('PROLI Marketing 360 está lista como aplicación')
  },

  onRegisterError(error) {
    console.error('Error registrando PROLI:', error)
  },
})

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)
