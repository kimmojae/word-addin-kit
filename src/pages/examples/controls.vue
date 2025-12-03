<route lang="yaml">
meta:
  title: Content Controls
</route>

<script setup lang="ts">
import { ref } from 'vue'

const status = ref('')

// 일반 텍스트 컨트롤
async function createTextControl() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const contentControl = body.insertContentControl(Word.ContentControlType.plainText)
      contentControl.tag = 'textControl'
      contentControl.title = '이름 입력'
      contentControl.placeholderText = '여기에 이름을 입력하세요'
      contentControl.appearance = Word.ContentControlAppearance.boundingBox
      contentControl.cannotEdit = false
      contentControl.cannotDelete = false

      await context.sync()
      status.value = '✅ 텍스트 컨트롤이 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 체크박스 컨트롤
async function createCheckbox() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const paragraph = body.insertParagraph('', Word.InsertLocation.end)
      const checkbox = paragraph.insertContentControl(Word.ContentControlType.checkBox)
      checkbox.title = '동의 여부'
      paragraph.insertText(' 개인정보 수집에 동의합니다.', Word.InsertLocation.end)

      await context.sync()
      status.value = '✅ 체크박스가 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 드롭다운 리스트
async function createDropdown() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const range = body.getRange(Word.RangeLocation.end)

      // 드롭다운 리스트 생성
      const dropdown = range.insertContentControl(Word.ContentControlType.dropDownList)
      dropdown.title = '부서 선택'
      dropdown.tag = 'dropdown'
      dropdown.appearance = Word.ContentControlAppearance.boundingBox

      await context.sync()

      // 옵션 추가 - dropDownListContentControl 객체 사용
      if (dropdown.dropDownListContentControl) {
        dropdown.dropDownListContentControl.addListItem('개발팀', 'dev')
        dropdown.dropDownListContentControl.addListItem('디자인팀', 'design')
        dropdown.dropDownListContentControl.addListItem('기획팀', 'planning')
        dropdown.dropDownListContentControl.addListItem('마케팅팀', 'marketing')
      }

      await context.sync()
      status.value = '✅ 드롭다운 리스트가 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 콤보박스
async function createComboBox() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const range = body.getRange(Word.RangeLocation.end)

      // 콤보박스 생성
      const comboBox = range.insertContentControl(Word.ContentControlType.comboBox)
      comboBox.title = '직급 선택'
      comboBox.tag = 'combobox'
      comboBox.appearance = Word.ContentControlAppearance.boundingBox

      await context.sync()

      // 옵션 추가
      if (comboBox.comboBoxContentControl) {
        comboBox.comboBoxContentControl.addListItem('사원', 'staff')
        comboBox.comboBoxContentControl.addListItem('대리', 'assistant')
        comboBox.comboBoxContentControl.addListItem('과장', 'manager')
        comboBox.comboBoxContentControl.addListItem('부장', 'director')
      }

      await context.sync()
      status.value = '✅ 콤보박스가 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 날짜 선택기
async function createDatePicker() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const datePicker = body.insertContentControl(Word.ContentControlType.datePicker)
      datePicker.title = '날짜 선택'
      datePicker.placeholderText = '날짜를 선택하세요'

      await context.sync()
      status.value = '✅ 날짜 선택기가 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 그림 컨트롤
async function createPictureControl() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const pictureControl = body.insertContentControl(Word.ContentControlType.picture)
      pictureControl.title = '이미지 영역'

      await context.sync()
      status.value = '✅ 그림 컨트롤이 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 리치 텍스트 컨트롤
async function createRichTextControl() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      const richText = body.insertContentControl(Word.ContentControlType.richText)
      richText.tag = 'richTextControl'
      richText.title = '서식 있는 텍스트'
      richText.placeholderText = '여기에 서식 있는 텍스트를 입력하세요'
      richText.appearance = Word.ContentControlAppearance.tags

      await context.sync()
      status.value = '✅ 리치 텍스트 컨트롤이 생성되었습니다.'
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 컨트롤 값 읽기
async function readControlValue() {
  try {
    await Word.run(async (context) => {
      const controls = context.document.body.contentControls.getByTag('textControl')
      context.load(controls, 'items/text')
      await context.sync()

      if (controls.items.length > 0) {
        const text = controls.items[0].text
        status.value = `✅ 컨트롤 값: "${text}"`
      } else {
        status.value = '⚠️ textControl 태그를 가진 컨트롤이 없습니다.'
      }
    })
  } catch (error) {
    status.value = `❌ 에러: ${error}`
  }
}

// 양식 생성 (종합)
async function createForm() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body

      // 제목
      const title = body.insertParagraph('신청서', Word.InsertLocation.end)
      title.styleBuiltIn = Word.BuiltInStyleName.heading1

      body.insertParagraph('', Word.InsertLocation.end)

      // 이름
      const namePara = body.insertParagraph('이름: ', Word.InsertLocation.end)
      const nameRange = namePara.getRange(Word.RangeLocation.end)
      const nameControl = nameRange.insertContentControl(Word.ContentControlType.plainText)
      nameControl.title = '이름'
      nameControl.placeholderText = '이름을 입력하세요'

      body.insertParagraph('', Word.InsertLocation.end)

      // 신청일
      const datePara = body.insertParagraph('신청일: ', Word.InsertLocation.end)
      const dateRange = datePara.getRange(Word.RangeLocation.end)
      const dateControl = dateRange.insertContentControl(Word.ContentControlType.datePicker)
      dateControl.title = '신청일'

      body.insertParagraph('', Word.InsertLocation.end)

      // 부서
      const deptPara = body.insertParagraph('부서: ', Word.InsertLocation.end)
      const deptRange = deptPara.getRange(Word.RangeLocation.end)
      const deptControl = deptRange.insertContentControl(Word.ContentControlType.plainText)
      deptControl.title = '부서'
      deptControl.placeholderText = '부서를 입력하세요'

      body.insertParagraph('', Word.InsertLocation.end)

      // 동의 - 텍스트 먼저 추가하고 체크박스 삽입
      const agreePara = body.insertParagraph(' 개인정보 수집에 동의합니다.', Word.InsertLocation.end)
      const agreeStartRange = agreePara.getRange(Word.RangeLocation.start)
      const agreeCheck = agreeStartRange.insertContentControl(Word.ContentControlType.checkBox)
      agreeCheck.title = '동의'

      await context.sync()
      status.value = '✅ 양식이 생성되었습니다.'
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
      <h1 class="text-3xl font-bold mb-2 text-gray-900 dark:text-gray-100">Content Controls</h1>
      <p class="text-gray-600 dark:text-gray-400">
        체크박스, 드롭다운, 날짜 선택기 등 다양한 Content Control을 테스트합니다.
      </p>
    </header>

    <!-- Status -->
    <div
      v-if="status"
      class="mb-6 p-4 rounded-lg bg-blue-50 dark:bg-blue-900/20 border border-blue-200 dark:border-blue-800"
    >
      <p class="text-sm text-blue-900 dark:text-blue-200">{{ status }}</p>
    </div>

    <!-- 기본 컨트롤 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        📝 기본 컨트롤
      </h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="createTextControl"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          텍스트 컨트롤
        </button>
        <button
          @click="createRichTextControl"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          리치 텍스트 컨트롤
        </button>
        <button
          @click="createCheckbox"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          체크박스
        </button>
        <button
          @click="createPictureControl"
          class="px-4 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-lg transition-colors"
        >
          그림 컨트롤
        </button>
      </div>
    </section>

    <!-- 선택 컨트롤 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        📋 선택 컨트롤
      </h2>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
        <button
          @click="createDropdown"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors"
        >
          드롭다운 리스트
        </button>
        <button
          @click="createComboBox"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors"
        >
          콤보박스
        </button>
        <button
          @click="createDatePicker"
          class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors"
        >
          날짜 선택기
        </button>
      </div>
    </section>

    <!-- 컨트롤 조작 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        🔧 컨트롤 조작
      </h2>
      <div class="space-y-3">
        <button
          @click="readControlValue"
          class="w-full px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg transition-colors"
        >
          텍스트 컨트롤 값 읽기
        </button>
      </div>
    </section>

    <!-- 종합 예제 -->
    <section class="mb-8">
      <h2 class="text-xl font-semibold mb-4 text-gray-900 dark:text-gray-100">
        📄 종합 예제
      </h2>
      <div class="space-y-3">
        <button
          @click="createForm"
          class="w-full px-4 py-2 bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-700 hover:to-indigo-700 text-white rounded-lg transition-colors"
        >
          양식 생성 (이름, 날짜, 부서, 동의)
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
