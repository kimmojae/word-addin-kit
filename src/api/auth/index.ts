import type { LoginRequest, LoginResponse, ProtectedResponse } from '@/types/api/auth'
import type { ApiResponse } from '@/types/common'
import { get, post } from '../client'

/**
 * 로그인 API
 *
 * Note: Store의 login() 메서드를 사용하는 것을 권장합니다.
 * 이 함수는 Store 내부에서 호출됩니다.
 */
export const login = async (credentials: LoginRequest) => {
  return post<ApiResponse<LoginResponse>>('/auth/login', credentials)
}

/**
 * 로그아웃 API
 *
 * Note: Store의 logout() 메서드를 사용하는 것을 권장합니다.
 * 이 함수는 Store 내부에서 호출됩니다.
 */
export const logout = async () => {
  // 실제 API에서는 서버에 로그아웃 요청을 보낼 수 있습니다
  // 현재는 클라이언트 측에서만 처리됩니다
}

/**
 * 인증이 필요한 리소스 가져오기 API
 *
 * Note: Store의 getProtected() 메서드를 사용하는 것을 권장합니다.
 * 이 함수는 Store 내부에서 호출됩니다.
 */
export const getProtected = async () => {
  return get<ApiResponse<ProtectedResponse>>('/protected')
}
