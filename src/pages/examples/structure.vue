<route lang="yaml">
meta:
  title: 문서 구조
</route>

<script setup lang="ts">
import { ref } from 'vue'

const status = ref('')

// 테이블 생성
async function createTable() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body

      const data = [
        ['이름', '나이', '직업'],
        ['홍길동', '30', '개발자'],
        ['김철수', '25', '디자이너'],
        ['이영희', '28', '기획자'],
      ]

      const table = body.insertTable(data.length, data[0].length, Word.InsertLocation.end, data)

      // 테이블 스타일
      table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light

      // 헤더 행 스타일
      table.headerRowCount = 1

      // rows를 먼저 load
      const rows = table.rows
      context.load(rows)
      await context.sync()

      const headerRow = rows.items[0]
      headerRow.font.bold = true
      headerRow.font.color = '#FFFFFF'
      headerRow.shadingColor = '#4472C4'

      await context.sync()
      status.value = '✅ 테이블이 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 섹션 정보 가져오기
async function getSectionInfo() {
  try {
    await Word.run(async (context) => {
      const sections = context.document.sections
      context.load(sections, 'items')
      await context.sync()

      status.value = `✅ 문서에 총 ${sections.items.length}개의 섹션이 있습니다.`

      // 각 섹션의 body 정보 표시
      if (sections.items.length > 0) {
        const section = sections.items[0]
        const body = section.body
        context.load(body, 'text')
        await context.sync()

        const textLength = body.text.length
        status.value += `\n첫 번째 섹션의 텍스트 길이: ${textLength}자`
      }
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 헤더 추가
async function addHeader() {
  try {
    await Word.run(async (context) => {
      const section = context.document.sections.getFirst()
      const header = section.getHeader(Word.HeaderFooterType.primary)

      header.clear()
      const paragraph = header.insertParagraph('문서 제목 - 헤더', Word.InsertLocation.start)
      paragraph.alignment = Word.Alignment.centered
      paragraph.font.size = 10
      paragraph.font.color = '#666666'

      await context.sync()
      status.value = '✅ 헤더가 추가되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 푸터 추가
async function addFooter() {
  try {
    await Word.run(async (context) => {
      const section = context.document.sections.getFirst()
      const footer = section.getFooter(Word.HeaderFooterType.primary)

      footer.clear()
      const paragraph = footer.insertParagraph('페이지 ', Word.InsertLocation.start)
      paragraph.alignment = Word.Alignment.centered
      paragraph.font.size = 10
      paragraph.font.color = '#666666'

      await context.sync()
      status.value = '✅ 푸터가 추가되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 이미지 삽입 (샘플 Base64 이미지)
async function insertImage() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body

      // 작은 투명 PNG (1x1 픽셀)
      const base64Image =
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

      const picture = body.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end)

      picture.width = 100
      picture.height = 100
      picture.altTextTitle = '샘플 이미지'
      picture.altTextDescription = 'API로 삽입된 이미지'

      await context.sync()
      status.value = '✅ 이미지가 삽입되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 번호 매기기 목록
async function createNumberedList() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body

      const paragraph1 = body.insertParagraph('첫 번째 항목', Word.InsertLocation.end)
      const list = paragraph1.startNewList()

      // list.id를 로드
      context.load(list, 'id')
      await context.sync()

      const paragraph2 = body.insertParagraph('두 번째 항목', Word.InsertLocation.end)
      paragraph2.attachToList(list.id, 0)

      const paragraph3 = body.insertParagraph('세 번째 항목', Word.InsertLocation.end)
      paragraph3.attachToList(list.id, 0)

      await context.sync()
      status.value = '✅ 번호 매기기 목록이 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 페이지 나누기
async function insertPageBreak() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      body.insertBreak(Word.BreakType.page, Word.InsertLocation.end)
      await context.sync()
      status.value = '✅ 페이지 나누기가 삽입되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 책갈피 추가 (ContentControl로 구현)
async function addBookmark() {
  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection()
      context.load(range, 'text')
      await context.sync()

      if (!range.text || range.text.trim() === '') {
        status.value = '⚠️ 텍스트를 먼저 선택하세요.'
        return
      }

      // ContentControl을 책갈피처럼 사용
      const control = range.insertContentControl()
      control.tag = 'MyBookmark'
      control.title = '책갈피'
      control.appearance = Word.ContentControlAppearance.tags
      control.color = '#ADD8E6'

      await context.sync()
      status.value = `✅ 책갈피가 추가되었습니다. (텍스트: "${range.text}")`
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 책갈피로 이동
async function goToBookmark() {
  try {
    await Word.run(async (context) => {
      const controls = context.document.contentControls.getByTag('MyBookmark')
      context.load(controls, 'items')
      await context.sync()

      if (controls.items.length === 0) {
        status.value = '⚠️ 책갈피를 찾을 수 없습니다. 먼저 텍스트를 선택하고 "책갈피 추가"를 클릭하세요.'
        return
      }

      const control = controls.items[0]
      control.select()
      await context.sync()
      status.value = '✅ 책갈피로 이동했습니다.'
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
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}
</script>

<template>
  <div class="max-w-4xl mx-auto p-8">
    <header class="mb-8">
      <h1 class="text-3xl font-bold mb-2 text-gray-900 dark:text-gray-100">문서 구조</h1>
      <p class="text-gray-600 dark:text-gray-400">
        섹션, 헤더/푸터, 테이블, 이미지 등 문서 구조 관련 API를 테스트합니다.
      </p>
    </header>

    <!-- Status -->
    <div
      v-if="status"
      class="mb-6 p-4 rounded-lg bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800"
    >
      <p class="text-sm text-blue-900 dark:text-blue-200">{{ status }}</p>
    </div>

    <!-- 테이블 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">📊 테이블</h2>
      <div class="space-y-3">
        <button
          @click="createTable"
          class="w-full px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          테이블 생성 (3x4)
        </button>
      </div>
    </section>

    <!-- 섹션 & 페이지 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        📄 섹션 & 페이지
      </h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="getSectionInfo"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors"
        >
          섹션 정보 가져오기
        </button>
        <button
          @click="insertPageBreak"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors"
        >
          페이지 나누기
        </button>
      </div>
    </section>

    <!-- 헤더 & 푸터 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        📑 헤더 & 푸터
      </h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="addHeader"
          class="px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors"
        >
          헤더 추가
        </button>
        <button
          @click="addFooter"
          class="px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors"
        >
          푸터 추가
        </button>
      </div>
    </section>

    <!-- 이미지 & 미디어 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">🖼️ 이미지</h2>
      <div class="space-y-3">
        <button
          @click="insertImage"
          class="w-full px-4 py-2 bg-pink-600 hover:bg-pink-700 text-white rounded-lg transition-colors"
        >
          이미지 삽입 (Base64)
        </button>
      </div>
    </section>

    <!-- 목록 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">📝 목록</h2>
      <div class="space-y-3">
        <button
          @click="createNumberedList"
          class="w-full px-4 py-2 bg-orange-600 hover:bg-orange-700 text-white rounded-lg transition-colors"
        >
          번호 매기기 목록 생성
        </button>
      </div>
    </section>

    <!-- 책갈피 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">🔖 책갈피</h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="addBookmark"
          class="px-4 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-lg transition-colors"
        >
          책갈피 추가
        </button>
        <button
          @click="goToBookmark"
          class="px-4 py-2 bg-teal-600 hover:bg-teal-700 text-white rounded-lg transition-colors"
        >
          책갈피로 이동
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
