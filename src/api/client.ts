import type { ApiError } from '@/types/common'
import axios, { type AxiosError, type AxiosInstance, type InternalAxiosRequestConfig } from 'axios'
import { logError, logRequest, logResponse } from './logger'

/**
 * Axios 인스턴스 생성
 * baseURL = VITE_API_BASE_URL + VITE_API_PREFIX
 */
export const apiClient: AxiosInstance = axios.create({
  baseURL: (import.meta.env.VITE_API_BASE_URL || '') + (import.meta.env.VITE_API_PREFIX || '/api'),
  headers: {
    'Content-Type': 'application/json',
  },
  timeout: 30000, // 30초
})

/**
 * Request Interceptor
 * - 인증 토큰 자동 추가
 * - 요청 로깅
 */
apiClient.interceptors.request.use(
  (config: InternalAxiosRequestConfig) => {
    // 인증 토큰 추가 (Pinia Store에서 가져오기)
    // Note: interceptor에서는 useAuthStore()를 직접 호출할 수 없으므로
    // localStorage에서 가져옵니다 (Pinia persist가 저장한 값)
    const persistedAuth = localStorage.getItem('auth')
    if (persistedAuth) {
      try {
        const authData = JSON.parse(persistedAuth)
        const token = authData.token
        if (token && config.headers) {
          config.headers.Authorization = `Bearer ${token}`
        }
      } catch (e) {
        console.error('Failed to parse auth data:', e)
      }
    }

    // 요청 로깅
    logRequest(config)

    return config
  },
  (error: AxiosError) => {
    logError(error)
    return Promise.reject(error)
  },
)

/**
 * Response Interceptor
 * - 응답 로깅
 * - 에러 처리
 * - 401 에러 시 자동 로그아웃 처리
 */
apiClient.interceptors.response.use(
  (response) => {
    // 응답 로깅
    logResponse(response)
    return response
  },
  (error: AxiosError<ApiError>) => {
    // 에러 로깅
    logError(error)

    // 401 Unauthorized - 토큰 만료 또는 인증 실패
    if (error.response?.status === 401) {
      // Store 초기화 (persist가 자동으로 localStorage도 제거)
      // Note: interceptor에서는 useAuthStore()를 직접 호출할 수 없으므로
      // localStorage를 직접 제거합니다
      localStorage.removeItem('auth')

      // 로그인 페이지로 리다이렉트
      // router를 사용할 수 없으므로 window.location 사용
      // if (!window.location.pathname.includes('/login')) {
      //   window.location.href = '/login'
      // }
    }

    // 에러 응답 정규화
    const apiError: ApiError = {
      message: error.response?.data?.message || error.message || 'An error occurred',
      status: error.response?.status || 500,
      code: error.response?.data?.code,
      details: error.response?.data?.details,
    }

    return Promise.reject(apiError)
  },
)

/**
 * API 헬퍼 함수들
 */

/**
 * GET 요청
 */
export async function get<T>(url: string, params?: Record<string, unknown>) {
  const response = await apiClient.get<T>(url, { params })
  return response
}

/**
 * POST 요청
 */
export async function post<T>(url: string, data?: unknown) {
  const response = await apiClient.post<T>(url, data)
  return response
}

/**
 * PUT 요청
 */
export async function put<T>(url: string, data?: unknown) {
  const response = await apiClient.put<T>(url, data)
  return response
}

/**
 * PATCH 요청
 */
export async function patch<T>(url: string, data?: unknown) {
  const response = await apiClient.patch<T>(url, data)
  return response
}

/**
 * DELETE 요청
 */
export async function del<T>(url: string) {
  const response = await apiClient.delete<T>(url)
  return response
}
