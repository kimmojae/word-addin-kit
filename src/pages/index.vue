<route lang="yaml">
meta:
  title: Word API Kit
</route>

<script setup lang="ts">
const categories = [
  {
    title: '기본 API',
    description: '텍스트, 단락, 서식 등 기본적인 Word API 기능',
    path: '/examples/basic',
    icon: '📝',
  },
  {
    title: '문서 구조',
    description: '섹션, 헤더/푸터, 테이블, 이미지 등',
    path: '/examples/structure',
    icon: '🏗️',
  },
  {
    title: 'Content Controls',
    description: '체크박스, 드롭다운, 날짜 선택기 등',
    path: '/examples/controls',
    icon: '🎛️',
  },
  {
    title: '고급 기능',
    description: '주석, 각주, 이벤트, 자동화 등',
    path: '/examples/advanced',
    icon: '⚡',
  },
]
</script>

<template>
  <div class="max-w-6xl mx-auto p-8">
    <!-- Categories Grid -->
    <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mb-12">
      <RouterLink
        v-for="category in categories"
        :key="category.path"
        :to="category.path"
        class="bg-white dark:bg-gray-800 border border-gray-200 dark:border-gray-700 rounded-xl p-6 hover:border-purple-500 dark:hover:border-purple-400 hover:shadow-lg transition-all duration-300 hover:-translate-y-1"
      >
        <div class="text-5xl mb-4">{{ category.icon }}</div>
        <h2 class="text-xl font-semibold mb-2 text-gray-900 dark:text-gray-100">
          {{ category.title }}
        </h2>
        <p class="text-sm text-gray-600 dark:text-gray-400 leading-relaxed">
          {{ category.description }}
        </p>
      </RouterLink>
    </div>

    <!-- Quick Links -->
    <div
      class="bg-gray-50 dark:bg-gray-800 rounded-xl p-6 border border-gray-200 dark:border-gray-700"
    >
      <h3 class="text-lg font-semibold mb-4 text-gray-900 dark:text-gray-100">빠른 시작</h3>
      <ul class="space-y-3">
        <li>
          <a
            href="https://learn.microsoft.com/javascript/api/word"
            target="_blank"
            class="text-purple-600 dark:text-purple-400 hover:text-purple-800 dark:hover:text-purple-300 transition-colors"
          >
            📚 공식 문서
          </a>
        </li>
        <li>
          <RouterLink
            to="/about"
            class="text-purple-600 dark:text-purple-400 hover:text-purple-800 dark:hover:text-purple-300 transition-colors"
          >
            ℹ️ 프로젝트 소개
          </RouterLink>
        </li>
      </ul>
    </div>
  </div>
</template>
