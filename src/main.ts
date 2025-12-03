import './assets/main.css'

import { VueQueryPlugin } from '@tanstack/vue-query'
import { createPinia } from 'pinia'
import piniaPluginPersistedstate from 'pinia-plugin-persistedstate'
import { createApp } from 'vue'

import App from './App.vue'
import router from './router'

/**
 * Start the application
 * Wait for Office.js to be ready before mounting Vue app
 */
window.Office.onReady((info) => {
  if (import.meta.env.DEV) {
    const isOfficeContext = info.host && info.platform
    if (isOfficeContext) {
      console.log(
        `%c[Office.js]%c Ready (Host: ${info.host}, Platform: ${info.platform})`,
        'color: #56b6c2',
        'color: #98c379',
      )
    } else {
      console.log(
        '%c[Office.js]%c Running in browser mode (Office context not available)',
        'color: #56b6c2',
        'color: #d19a66',
      )
    }
  }

  const app = createApp(App)
  const pinia = createPinia()

  pinia.use(piniaPluginPersistedstate)

  app.use(pinia)
  app.use(router)
  app.use(VueQueryPlugin, {
    queryClientConfig: {
      defaultOptions: {
        queries: {
          staleTime: 1000 * 60 * 5, // 5분
          gcTime: 1000 * 60 * 10, // 10분
          refetchOnWindowFocus: false,
          retry: 1,
        },
      },
    },
  })

  app.mount('#app')
})
