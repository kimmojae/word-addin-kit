import type { CreateUserRequest, UpdateUserRequest } from '@/types/api/users'
import { useMutation, useQueryClient } from '@tanstack/vue-query'
import { createUser, deleteUser, updateUser } from './index'

/**
 * 사용자 생성 Mutation
 */
export function useCreateUserMutation() {
  const queryClient = useQueryClient()

  return useMutation({
    mutationFn: (data: CreateUserRequest) => createUser(data),
    onSuccess: () => {
      // 사용자 목록 쿼리 무효화 → 자동 재조회
      queryClient.invalidateQueries({ queryKey: ['users'] })
    },
  })
}

/**
 * 사용자 수정 Mutation
 */
export function useUpdateUserMutation() {
  const queryClient = useQueryClient()

  return useMutation({
    mutationFn: ({ id, data }: { id: number; data: UpdateUserRequest }) => updateUser(id, data),
    onSuccess: (_, variables) => {
      // 사용자 목록 및 상세 쿼리 무효화
      queryClient.invalidateQueries({ queryKey: ['users'] })
      queryClient.invalidateQueries({ queryKey: ['users', { id: variables.id }] })
    },
  })
}

/**
 * 사용자 삭제 Mutation
 */
export function useDeleteUserMutation() {
  const queryClient = useQueryClient()

  return useMutation({
    mutationFn: (id: number) => deleteUser(id),
    onSuccess: (_, id) => {
      // 사용자 목록 쿼리 무효화
      queryClient.invalidateQueries({ queryKey: ['users'] })
      // 삭제된 사용자 상세 쿼리 제거
      queryClient.removeQueries({ queryKey: ['users', { id }] })
    },
  })
}
