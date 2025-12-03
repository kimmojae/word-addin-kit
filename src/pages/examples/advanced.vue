<route lang="yaml">
meta:
  title: 고급 기능
</route>

<script setup lang="ts">
import { ref } from 'vue'

const status = ref('')
const eventLog = ref<string[]>([])

// 각주 추가
async function addFootnote() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      const footnote = range.insertFootnote('이것은 각주 내용입니다.')

      context.load(footnote, 'body/text')
      await context.sync()

      status.value = `✅ 각주가 추가되었습니다: "${footnote.body.text}"`
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 미주 추가
async function addEndnote() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      const endnote = range.insertEndnote('이것은 미주 내용입니다.')

      context.load(endnote, 'body/text')
      await context.sync()

      status.value = `✅ 미주가 추가되었습니다: "${endnote.body.text}"`
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 문서 속성 읽기
async function getDocumentProperties() {
  try {
    await Word.run(async (context) => {
      const properties = context.document.properties
      context.load(properties, 'title,author,subject,keywords')
      await context.sync()

      status.value = `✅ 문서 속성:
        제목: ${properties.title || '(없음)'}
        작성자: ${properties.author || '(없음)'}
        주제: ${properties.subject || '(없음)'}
        키워드: ${properties.keywords || '(없음)'}`
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 문서 속성 설정
async function setDocumentProperties() {
  try {
    await Word.run(async (context) => {
      const properties = context.document.properties
      properties.title = 'Word API Kit 테스트 문서'
      properties.author = 'Word API Kit'
      properties.subject = 'API 테스트'
      properties.keywords = 'Word, API, JavaScript'

      await context.sync()
      status.value = '✅ 문서 속성이 설정되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 사용자 정의 속성 추가
async function addCustomProperty() {
  try {
    await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties

      customProps.add('ProjectName', 'Word API Kit')
      customProps.add('Version', '1.0.0')
      customProps.add('Status', 'In Progress')

      await context.sync()
      status.value = '✅ 사용자 정의 속성이 추가되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 사용자 정의 속성 읽기
async function readCustomProperty() {
  try {
    await Word.run(async (context) => {
      const customProps = context.document.properties.customProperties
      context.load(customProps, 'items')
      await context.sync()

      if (customProps.items.length > 0) {
        const propsList = customProps.items
          .map((prop) => {
            context.load(prop, 'key,value')
          })
          .join('')

        await context.sync()

        const propsText = customProps.items
          .map((prop) => `${prop.key}: ${prop.value}`)
          .join('\n        ')

        status.value = `✅ 사용자 정의 속성:\n        ${propsText}`
      } else {
        status.value = '⚠️ 사용자 정의 속성이 없습니다.'
      }
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 단락 추가 이벤트 등록
async function registerParagraphAddedEvent() {
  try {
    await Word.run(async (context) => {
      context.document.onParagraphAdded.add((eventArgs) => {
        eventLog.value.unshift(`[${new Date().toLocaleTimeString()}] 단락이 추가되었습니다.`)
        return context.sync()
      })

      await context.sync()
      status.value = '✅ 단락 추가 이벤트가 등록되었습니다. 새 단락을 추가해보세요!'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 단락 변경 이벤트 등록
async function registerParagraphChangedEvent() {
  try {
    await Word.run(async (context) => {
      context.document.onParagraphChanged.add((eventArgs) => {
        eventLog.value.unshift(`[${new Date().toLocaleTimeString()}] 단락이 변경되었습니다.`)
        return context.sync()
      })

      await context.sync()
      status.value = '✅ 단락 변경 이벤트가 등록되었습니다. 단락을 수정해보세요!'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 이벤트 로그 초기화
function clearEventLog() {
  eventLog.value = []
  status.value = '✅ 이벤트 로그가 초기화되었습니다.'
}

// 문서 저장
async function saveDocument() {
  try {
    await Word.run(async (context) => {
      context.document.save()
      await context.sync()
      status.value = '✅ 문서가 저장되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 모든 단락 정보 가져오기
async function getAllParagraphs() {
  try {
    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs
      context.load(paragraphs, 'items')
      await context.sync()

      status.value = `✅ 문서에 총 ${paragraphs.items.length}개의 단락이 있습니다.`
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
      eventLog.value = []
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}
</script>

<template>
  <div class="max-w-4xl mx-auto p-8">
    <header class="mb-8">
      <h1 class="text-3xl font-bold mb-2 text-gray-900 dark:text-gray-100">고급 기능</h1>
      <p class="text-gray-600 dark:text-gray-400">
        주석, 각주, 이벤트, 문서 속성 등 고급 Word API 기능을 테스트합니다.
      </p>
    </header>

    <!-- Status -->
    <div
      v-if="status"
      class="mb-6 p-4 rounded-lg bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800 whitespace-pre-wrap"
    >
      <p class="text-sm text-blue-900 dark:text-blue-200">{{ status }}</p>
    </div>

    <!-- 각주 & 미주 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">📌 각주 & 미주</h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="addFootnote"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          각주 추가
        </button>
        <button
          @click="addEndnote"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          미주 추가
        </button>
      </div>
      <p class="text-xs text-gray-500 dark:text-gray-400 mt-2">
        * 텍스트를 선택한 후 버튼을 클릭하세요
      </p>
    </section>

    <!-- 문서 속성 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">📋 문서 속성</h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
        <button
          @click="getDocumentProperties"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors"
        >
          문서 속성 읽기
        </button>
        <button
          @click="setDocumentProperties"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors"
        >
          문서 속성 설정
        </button>
      </div>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="addCustomProperty"
          class="px-4 py-2 bg-indigo-500 hover:bg-indigo-600 text-white rounded-lg transition-colors"
        >
          사용자 정의 속성 추가
        </button>
        <button
          @click="readCustomProperty"
          class="px-4 py-2 bg-indigo-500 hover:bg-indigo-600 text-white rounded-lg transition-colors"
        >
          사용자 정의 속성 읽기
        </button>
      </div>
    </section>

    <!-- 이벤트 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">⚡ 이벤트</h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
        <button
          @click="registerParagraphAddedEvent"
          class="px-4 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg transition-colors"
        >
          단락 추가 이벤트 등록
        </button>
        <button
          @click="registerParagraphChangedEvent"
          class="px-4 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg transition-colors"
        >
          단락 변경 이벤트 등록
        </button>
      </div>

      <!-- 이벤트 로그 -->
      <div
        class="bg-gray-50 dark:bg-gray-800 rounded-lg p-4 border border-gray-200 dark:border-gray-700"
      >
        <div class="flex justify-between items-center mb-2">
          <h3 class="text-sm font-semibold text-gray-700 dark:text-gray-300">이벤트 로그</h3>
          <button
            @click="clearEventLog"
            class="text-xs px-2 py-1 bg-gray-200 dark:bg-gray-700 hover:bg-gray-300 dark:hover:bg-gray-600 rounded transition-colors"
          >
            초기화
          </button>
        </div>
        <div
          v-if="eventLog.length === 0"
          class="text-sm text-gray-400 dark:text-gray-500 text-center py-4"
        >
          이벤트가 없습니다
        </div>
        <div v-else class="space-y-1 max-h-40 overflow-y-auto">
          <div
            v-for="(log, index) in eventLog"
            :key="index"
            class="text-xs text-gray-700 dark:text-gray-300"
          >
            {{ log }}
          </div>
        </div>
      </div>
    </section>

    <!-- 문서 조작 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">💾 문서 조작</h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="getAllParagraphs"
          class="px-4 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-lg transition-colors"
        >
          모든 단락 정보
        </button>
        <button
          @click="saveDocument"
          class="px-4 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-lg transition-colors"
        >
          문서 저장
        </button>
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
