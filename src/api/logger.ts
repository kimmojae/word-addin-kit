import type { AxiosError, AxiosRequestConfig, AxiosResponse } from 'axios'

const STYLES = {
  white: 'color: #ffffff',
  yellow: 'color: #ffcb6b', // [API] 라벨
  green: 'color: #98c379', // GET, 200 OK (MSW 초록)
  blue: 'color: #61afef', // POST
  orange: 'color: #d19a66', // PUT
  purple: 'color: #c678dd', // PATCH
  red: 'color: #e06c75', // DELETE, 에러 (MSW 빨강)
} as const

function getMethodStyle(method: string): string {
  const styles: Record<string, string> = {
    GET: STYLES.green,
    POST: STYLES.blue,
    PUT: STYLES.orange,
    PATCH: STYLES.purple,
    DELETE: STYLES.red,
  }
  return styles[method.toUpperCase()] || STYLES.green
}

function getStatusStyle(status: number): string {
  if (status >= 400) return STYLES.red
  if (status >= 300) return STYLES.orange
  return STYLES.green
}

function getStatusText(status: number): string {
  const statusTexts: Record<number, string> = {
    200: 'OK',
    201: 'Created',
    204: 'No Content',
    400: 'Bad Request',
    401: 'Unauthorized',
    403: 'Forbidden',
    404: 'Not Found',
    500: 'Internal Server Error',
    502: 'Bad Gateway',
    503: 'Service Unavailable',
  }
  return statusTexts[status] || ''
}

interface PendingRequest {
  method: string
  startTime: number
  params?: unknown
  data?: unknown
}

const pendingRequests = new Map<string, PendingRequest>()

/**
 * API 요청 로거
 */
export function logRequest(config: AxiosRequestConfig): void {
  if (import.meta.env.DEV) {
    const { method = 'GET', url = '', params, data } = config
    pendingRequests.set(url, {
      method,
      startTime: Date.now(),
      params,
      data,
    })
  }
}

/**
 * API 응답 로거
 */
export function logResponse(response: AxiosResponse): void {
  if (import.meta.env.DEV) {
    const { config, status, data } = response
    const { url = '' } = config
    const pending = pendingRequests.get(url)

    if (pending) {
      const method = pending.method.toUpperCase()
      const statusText = getStatusText(status)

      console.groupCollapsed(
        `%c[API]%c %c${method}%c ${url} %c(${status} ${statusText})`,
        STYLES.yellow,
        STYLES.white,
        getMethodStyle(method),
        STYLES.white,
        getStatusStyle(status),
      )

      console.log('Request', {
        url,
        method,
        headers: config.headers || {},
        body: pending.data || '',
      })

      console.log('Response', data)
      console.groupEnd()

      pendingRequests.delete(url)
    }
  }
}

/**
 * API 에러 로거
 */
export function logError(error: AxiosError): void {
  if (import.meta.env.DEV) {
    const { config, response } = error
    const url = config?.url || 'UNKNOWN'
    const pending = pendingRequests.get(url)

    if (pending) {
      const method = pending.method.toUpperCase()
      const status = response?.status
      const statusText = status ? getStatusText(status) : 'Network Error'
      const statusDisplay = status ? `${status} ${statusText}` : statusText

      console.groupCollapsed(
        `%c[API]%c %c${method}%c ${url} %c(${statusDisplay})`,
        STYLES.yellow,
        STYLES.white,
        getMethodStyle(method),
        STYLES.white,
        STYLES.red,
      )

      console.log('Request', {
        url,
        method,
        headers: config?.headers || {},
        body: pending.data || '',
      })

      console.log('Response', response?.data)
      console.groupEnd()

      pendingRequests.delete(url)
    }
  }
}
