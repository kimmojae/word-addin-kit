import * as authAPI from '@/api/auth'
import type { LoginRequest } from '@/types/api/auth'
import type { ActionResult } from '@/types/common'
import type { User } from '@/types/models/user'
import { defineStore } from 'pinia'

export const useAuthStore = defineStore(
  'auth',
  () => {
    // ===== State =====
    const token = ref<string | null>(null)
    const user = ref<User | null>(null)
    const loading = ref(false)
    const error = ref<string | null>(null)

    // ===== Computed =====
    const isAuthenticated = computed(() => !!token.value)

    // ===== Private Methods =====
    const resetState = () => {
      token.value = null
      user.value = null
      error.value = null
    }

    const handleError = (err: unknown): string => {
      const apiError = err as { message: string }
      error.value = apiError.message
      return apiError.message
    }

    // ===== Actions =====
    /**
     * 로그인
     */
    const login = async (credentials: LoginRequest): Promise<ActionResult> => {
      try {
        loading.value = true
        error.value = null

        const response = await authAPI.login(credentials)
        const { access_token, user: authUser } = response.data.data

        // Store에 저장 (persist로 자동 localStorage 저장)
        token.value = access_token
        user.value = authUser

        return { success: true }
      } catch (err: unknown) {
        return { success: false, error: handleError(err) }
      } finally {
        loading.value = false
      }
    }

    /**
     * 로그아웃
     */
    const logout = async (): Promise<void> => {
      try {
        loading.value = true
        await authAPI.logout()
        resetState()
      } finally {
        loading.value = false
      }
    }

    /**
     * Protected 리소스 가져오기
     */
    const getProtected = async (): Promise<ActionResult<{ message: string }>> => {
      try {
        loading.value = true
        error.value = null

        const response = await authAPI.getProtected()
        return { success: true, data: response.data.data }
      } catch (err: unknown) {
        return { success: false, error: handleError(err) }
      } finally {
        loading.value = false
      }
    }

    return {
      // State
      token,
      user,
      loading,
      error,
      isAuthenticated,

      // Actions
      login,
      logout,
      getProtected,
    }
  },
  {
    persist: true,
  },
)
