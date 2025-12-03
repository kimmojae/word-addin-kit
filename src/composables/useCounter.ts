/**
 * Counter Composable Example
 *
 * 이 composable은 카운터 로직을 재사용 가능한 형태로 캡슐화합니다.
 * 여러 컴포넌트에서 독립적인 카운터 인스턴스를 생성할 수 있습니다.
 *
 * @example
 * ```vue
 * <script setup lang="ts">
 * const { count, increment, decrement, reset } = useCounter(0)
 * </script>
 *
 * <template>
 *   <div>
 *     <p>Count: {{ count }}</p>
 *     <button @click="increment">+</button>
 *     <button @click="decrement">-</button>
 *     <button @click="reset">Reset</button>
 *   </div>
 * </template>
 * ```
 */
export function useCounter(initialValue = 0) {
  const count = ref(initialValue)

  const increment = () => {
    count.value++
  }

  const decrement = () => {
    count.value--
  }

  const reset = () => {
    count.value = initialValue
  }

  const isPositive = computed(() => count.value > 0)
  const isNegative = computed(() => count.value < 0)
  const isZero = computed(() => count.value === 0)

  return {
    // State
    count: readonly(count),

    // Actions
    increment,
    decrement,
    reset,

    // Computed
    isPositive,
    isNegative,
    isZero,
  }
}
