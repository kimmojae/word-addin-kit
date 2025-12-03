<route lang="yaml">
meta:
  title: 기본 API
</route>

<script setup lang="ts">
import { ref } from 'vue'

const status = ref('')
const selectedText = ref('')

// 텍스트 삽입
async function insertText() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      body.insertText('Hello from Word API Kit!', Word.InsertLocation.end)
      await context.sync()
      status.value = '✅ 텍스트가 삽입되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 단락 삽입
async function insertParagraph() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const paragraph = body.insertParagraph(
        '이것은 새로운 단락입니다.',
        Word.InsertLocation.end,
      )
      paragraph.alignment = Word.Alignment.left
      await context.sync()
      status.value = '✅ 단락이 삽입되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 선택 영역 볼드 처리
async function makeBold() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      range.font.bold = true
      await context.sync()
      status.value = '✅ 선택한 텍스트를 볼드 처리했습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 선택 영역 이탤릭 처리
async function makeItalic() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      range.font.italic = true
      await context.sync()
      status.value = '✅ 선택한 텍스트를 이탤릭 처리했습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 선택 영역 하이라이트
async function highlightSelection() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      range.font.highlightColor = '#FFFF00'
      await context.sync()
      status.value = '✅ 선택한 텍스트를 하이라이트했습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 선택 영역 색상 변경
async function changeColor() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      range.font.color = '#FF0000'
      await context.sync()
      status.value = '✅ 선택한 텍스트 색상을 변경했습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 선택 영역 읽기
async function getSelection() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      context.load(range, 'text')
      await context.sync()
      selectedText.value = range.text
      status.value = '✅ 선택한 텍스트를 읽었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 제목 1 스타일 적용
async function applyHeading1() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      const paragraph = range.paragraphs.getFirst()
      paragraph.styleBuiltIn = Word.BuiltInStyleName.heading1
      await context.sync()
      status.value = '✅ 제목 1 스타일을 적용했습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// HTML 삽입
async function insertHTML() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const html = `
        <h2>HTML로 삽입된 제목</h2>
        <p>이것은 <strong>볼드</strong>와 <em>이탤릭</em>이 있는 텍스트입니다.</p>
        <ul>
          <li>항목 1</li>
          <li>항목 2</li>
          <li>항목 3</li>
        </ul>
      `
      body.insertHtml(html, Word.InsertLocation.end)
      await context.sync()
      status.value = '✅ HTML이 삽입되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 텍스트 검색
async function searchText() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const searchResults = body.search('API', {
        matchCase: false,
        matchWholeWord: false,
      })
      context.load(searchResults)
      await context.sync()

      // 검색 결과 하이라이트
      searchResults.items.forEach((item) => {
        item.font.highlightColor = '#FFFF00'
      })
      await context.sync()
      status.value = `✅ "${searchResults.items.length}"개의 결과를 찾아 하이라이트했습니다.`
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 문서 초기화
async function clearDocument() {
  try {
    await Word.run(async (context) => {
      // 본문 초기화
      const body = context.document.body
      body.clear()

      // 모든 섹션의 헤더/푸터 초기화
      const sections = context.document.sections
      context.load(sections, 'items')
      await context.sync()

      sections.items.forEach((section) => {
        const headerTypes = [
          Word.HeaderFooterType.primary,
          Word.HeaderFooterType.firstPage,
          Word.HeaderFooterType.evenPages,
        ]

        headerTypes.forEach((type) => {
          try {
            section.getHeader(type).clear()
            section.getFooter(type).clear()
          } catch (e) {
            // 일부 헤더/푸터가 없을 수 있음
          }
        })
      })

      await context.sync()
      status.value = '✅ 문서가 초기화되었습니다. (본문, 헤더, 푸터)'
      selectedText.value = ''
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}
</script>

<template>
  <div class="max-w-4xl mx-auto p-8">
    <header class="mb-8">
      <h1 class="text-3xl font-bold mb-2 text-gray-900 dark:text-gray-100">
        기본 API
      </h1>
      <p class="text-gray-600 dark:text-gray-400">
        텍스트, 단락, 서식 등 기본적인 Word API 기능을 테스트합니다.
      </p>
    </header>

    <!-- Status -->
    <div
      v-if="status"
      class="mb-6 p-4 rounded-lg bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800"
    >
      <p class="text-sm text-blue-900 dark:text-blue-200">{{ status }}</p>
    </div>

    <!-- 텍스트 & 단락 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        📝 텍스트 & 단락
      </h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="insertText"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          텍스트 삽입
        </button>
        <button
          @click="insertParagraph"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          단락 삽입
        </button>
        <button
          @click="insertHTML"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          HTML 삽입
        </button>
        <button
          @click="applyHeading1"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          제목 1 스타일 적용
        </button>
      </div>
    </section>

    <!-- 서식 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        🎨 서식 (선택 영역에 적용)
      </h2>
      <div class="grid grid-cols-2 md:grid-cols-4 gap-3">
        <button
          @click="makeBold"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors font-bold"
        >
          Bold
        </button>
        <button
          @click="makeItalic"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors italic"
        >
          Italic
        </button>
        <button
          @click="highlightSelection"
          class="px-4 py-2 bg-yellow-500 hover:bg-yellow-600 text-white rounded-lg transition-colors"
        >
          Highlight
        </button>
        <button
          @click="changeColor"
          class="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg transition-colors"
        >
          Red Color
        </button>
      </div>
    </section>

    <!-- 검색 & 읽기 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        🔍 검색 & 읽기
      </h2>
      <div class="space-y-3">
        <button
          @click="searchText"
          class="w-full px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors"
        >
          "API" 텍스트 검색 및 하이라이트
        </button>
        <button
          @click="getSelection"
          class="w-full px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors"
        >
          선택 영역 읽기
        </button>
        <div
          v-if="selectedText"
          class="p-4 bg-gray-50 dark:bg-gray-800 rounded-lg border border-gray-200 dark:border-gray-700"
        >
          <p class="text-sm font-semibold mb-1 text-gray-700 dark:text-gray-300">
            선택된 텍스트:
          </p>
          <p class="text-gray-900 dark:text-gray-100">{{ selectedText }}</p>
        </div>
      </div>
    </section>

    <!-- 문서 초기화 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">🗑️ 문서 관리</h2>
      <button
        @click="clearDocument"
        class="w-full px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg transition-colors"
      >
        문서 초기화 (모든 내용 삭제)
      </button>
    </section>

    <!-- 네비게이션 -->
    <div class="mt-12 pt-6 border-t border-gray-200 dark:border-gray-700">
      <RouterLink
        to="/"
        class="text-purple-600 dark:text-purple-400 hover:text-purple-800 dark:hover:text-purple-300 transition-colors"
      >
        ← 홈으로 돌아가기
      </RouterLink>
    </div>
  </div>
</template>
