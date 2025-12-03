import { useQuery } from '@tanstack/vue-query'
import type { MaybeRefOrGetter } from 'vue'
import { toValue } from 'vue'
import { getUser, getUsers } from './index'

/**
 * 사용자 목록 조회 Query
 */
export function useUsersQuery() {
  return useQuery({
    queryKey: ['users'],
    queryFn: async () => {
      const response = await getUsers()
      return response.data.data
    },
  })
}

/**
 * 사용자 상세 조회 Query
 */
export function useUserQuery(id: MaybeRefOrGetter<number>) {
  return useQuery({
    queryKey: ['users', { id: toValue(id) }],
    queryFn: async () => {
      const response = await getUser(toValue(id))
      return response.data.data
    },
    enabled: () => !!toValue(id), // id가 있을 때만 쿼리 실행
  })
}
