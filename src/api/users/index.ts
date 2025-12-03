import type { CreateUserRequest, UpdateUserRequest } from '@/types/api/users'
import type { ApiResponse } from '@/types/common'
import type { User } from '@/types/models/user'
import { del, get, post, put } from '../client'

/**
 * 사용자 목록 조회
 */
export async function getUsers() {
  return get<ApiResponse<User[]>>('/users')
}

/**
 * 사용자 상세 조회
 */
export async function getUser(id: number) {
  return get<ApiResponse<User>>(`/users/${id}`)
}

/**
 * 사용자 생성
 */
export async function createUser(data: CreateUserRequest) {
  return post<ApiResponse<User>>('/users', data)
}

/**
 * 사용자 수정
 */
export async function updateUser(id: number, data: UpdateUserRequest) {
  return put<ApiResponse<User>>(`/users/${id}`, data)
}

/**
 * 사용자 삭제
 */
export async function deleteUser(id: number) {
  return del<ApiResponse<null>>(`/users/${id}`)
}
