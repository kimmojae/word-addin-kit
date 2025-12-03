<script setup lang="ts">
import { BookOpen, Code, Package, Rocket, Wrench } from 'lucide-vue-next'
</script>

<template>
  <WelcomeItem>
    <template #icon>
      <BookOpen :size="20" />
    </template>
    <template #heading>Documentation</template>

    이 프로젝트의 가이드 문서를 확인하세요:
    <a href="/docs/PAGE_TEMPLATE.md" target="_blank" rel="noopener">페이지 템플릿 가이드</a>,
    <a href="/docs/COMPONENT_TEMPLATE.md" target="_blank" rel="noopener">컴포넌트 템플릿 가이드</a>,
    <a href="/docs/VSCODE_SNIPPETS.md" target="_blank" rel="noopener">VSCode 스니펫</a>을
    참고하세요.
  </WelcomeItem>

  <WelcomeItem>
    <template #icon>
      <Wrench :size="20" />
    </template>
    <template #heading>Tech Stack</template>

    이 프로젝트는
    <a href="https://vite.dev/" target="_blank" rel="noopener">Vite</a> +
    <a href="https://vuejs.org/" target="_blank" rel="noopener">Vue 3</a>로 구성되어 있습니다.
    스타일링은
    <a href="https://tailwindcss.com/" target="_blank" rel="noopener">Tailwind CSS v4</a>를
    사용합니다. IDE는
    <a href="https://code.visualstudio.com/" target="_blank" rel="noopener">VSCode</a> +
    <a
      href="https://marketplace.visualstudio.com/items?itemName=Vue.volar"
      target="_blank"
      rel="noopener"
    >
      Vue - Official </a
    >을 권장합니다.
  </WelcomeItem>

  <WelcomeItem>
    <template #icon>
      <Package :size="20" />
    </template>
    <template #heading>Features</template>

    <ul class="list-disc list-inside space-y-1">
      <li>
        <strong>Auto Import</strong>: Vue API, 컴포넌트, Composables 자동 임포트
        (unplugin-auto-import)
      </li>
      <li>
        <strong>File-based Routing</strong>: pages/ 폴더 기반 자동 라우팅 (unplugin-vue-router)
      </li>
      <li><strong>Layout System</strong>: 레이아웃 자동 적용 (vite-plugin-vue-layouts)</li>
      <li>
        <strong>State Management</strong>:
        <a href="https://pinia.vuejs.org/" target="_blank" rel="noopener">Pinia</a> + 상태 영속화
      </li>
      <li><strong>Dark Mode</strong>: VueUse useDark() 기반 다크모드 지원 (우측 상단 토글)</li>
      <li><strong>TypeScript</strong>: 전체 타입 안정성 지원</li>
    </ul>
  </WelcomeItem>

  <WelcomeItem>
    <template #icon>
      <Code :size="20" />
    </template>
    <template #heading>Developer Tools</template>

    개발 편의를 위한 도구들:
    <a href="https://github.com/antfu/eslint-config" target="_blank" rel="noopener">ESLint</a> +
    <a href="https://prettier.io/" target="_blank" rel="noopener">Prettier</a>로 코드 스타일 자동화,
    <a href="https://vueuse.org/" target="_blank" rel="noopener">VueUse</a>로 다양한 Composables
    제공, <a href="https://iconify.design/" target="_blank" rel="noopener">Iconify</a>와
    <a href="https://lucide.dev/" target="_blank" rel="noopener">Lucide</a>로 아이콘 사용 가능.
  </WelcomeItem>

  <WelcomeItem>
    <template #icon>
      <Rocket :size="20" />
    </template>
    <template #heading>Quick Start</template>

    <div class="space-y-2">
      <p><code class="bg-background-soft px-2 py-1 rounded">pnpm install</code> - 의존성 설치</p>
      <p><code class="bg-background-soft px-2 py-1 rounded">pnpm dev</code> - 개발 서버 시작</p>
      <p><code class="bg-background-soft px-2 py-1 rounded">pnpm build</code> - 프로덕션 빌드</p>
      <p>
        <code class="bg-background-soft px-2 py-1 rounded">pnpm lint</code> - 코드 스타일 검사/수정
      </p>
    </div>
  </WelcomeItem>
</template>
