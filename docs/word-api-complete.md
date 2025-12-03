# Word JavaScript API 완전 가이드

## 목차
1. [API 구조](#api-구조)
2. [핵심 객체](#핵심-객체)
3. [문서 구조 객체](#문서-구조-객체)
4. [서식 및 스타일](#서식-및-스타일)
5. [미디어 및 객체](#미디어-및-객체)
6. [문서 속성](#문서-속성)
7. [이벤트](#이벤트)
8. [고급 기능](#고급-기능)
9. [패턴 및 모범 사례](#패턴-및-모범-사례)

---

## API 구조

Word JavaScript API는 계층적 구조로 되어 있습니다.

```
Document
  ├─ Section(s)
  │   ├─ Header
  │   ├─ Footer
  │   └─ Body
  │       ├─ Paragraph(s)
  │       ├─ ContentControl(s)
  │       ├─ Table(s)
  │       ├─ InlinePicture(s)
  │       └─ Range(s)
  ├─ CustomProperties
  └─ Settings
```

---

## 핵심 객체

### 1. Document

문서 전체를 나타냅니다.

```typescript
await Word.run(async (context) => {
  const doc = context.document

  // 문서 바디 접근
  const body = doc.body

  // 문서 속성 접근
  const properties = doc.properties
  context.load(properties)
  await context.sync()

  console.log('제목:', properties.title)
  console.log('작성자:', properties.author)
})
```

**주요 속성:**
- `body`: 문서 본문
- `sections`: 섹션 컬렉션
- `properties`: 문서 속성
- `contentControls`: 콘텐츠 컨트롤
- `customXmlParts`: 사용자 정의 XML

**주요 메서드:**
- `getSelection()`: 현재 선택 영역 가져오기
- `save()`: 문서 저장

### 2. Body

문서의 본문 영역입니다.

```typescript
const body = context.document.body

// 텍스트 삽입
body.insertText('Hello World', Word.InsertLocation.start)

// 단락 삽입
body.insertParagraph('New paragraph', Word.InsertLocation.end)

// HTML 삽입
body.insertHtml('<p><strong>Bold</strong> text</p>', Word.InsertLocation.end)

// 파일 삽입 (Base64)
body.insertFileFromBase64(base64String, Word.InsertLocation.end)
```

**InsertLocation 옵션:**
- `start`: 시작 부분
- `end`: 끝 부분
- `before`: 앞에
- `after`: 뒤에
- `replace`: 대체

### 3. Paragraph

단락을 나타냅니다.

```typescript
// 모든 단락 가져오기
const paragraphs = body.paragraphs
context.load(paragraphs)
await context.sync()

// 첫 번째 단락
const firstParagraph = paragraphs.items[0]

// 스타일 적용
firstParagraph.styleBuiltIn = Word.BuiltInStyleName.heading1

// 정렬
firstParagraph.alignment = Word.Alignment.centered

// 줄 간격
firstParagraph.lineSpacing = 1.5

// 들여쓰기
firstParagraph.leftIndent = 36 // 포인트 단위
firstParagraph.firstLineIndent = 18
```

**주요 속성:**
- `text`: 단락 텍스트
- `styleBuiltIn`: 기본 스타일
- `alignment`: 정렬 (left, right, centered, justified)
- `font`: 폰트 설정
- `lineSpacing`: 줄 간격
- `leftIndent`, `rightIndent`: 들여쓰기
- `firstLineIndent`: 첫 줄 들여쓰기

**주요 메서드:**
- `insertParagraph()`: 단락 삽입
- `insertBreak()`: 페이지/섹션 나누기
- `delete()`: 삭제
- `select()`: 선택

### 4. Range

연속된 콘텐츠 영역 (텍스트, 공백, 테이블, 이미지 등)

```typescript
// 현재 선택 영역
const range = context.document.getSelection()
context.load(range, 'text')
await context.sync()

console.log('선택된 텍스트:', range.text)

// 서식 적용
range.font.bold = true
range.font.color = '#FF0000'
range.font.size = 16
range.font.highlightColor = '#FFFF00'

// 텍스트 검색
const searchResults = body.search('keyword', {
  matchCase: false,
  matchWholeWord: true,
  matchWildcards: false
})
context.load(searchResults)
await context.sync()

// 검색 결과에 하이라이트
searchResults.items.forEach(item => {
  item.font.highlightColor = '#FFFF00'
})
```

**검색 옵션:**
- `matchCase`: 대소문자 구분
- `matchWholeWord`: 단어 단위 검색
- `matchWildcards`: 와일드카드 사용
- `matchPrefix`: 접두사 일치
- `matchSuffix`: 접미사 일치

**와일드카드 패턴:**
- `?`: 임의의 한 문자
- `*`: 임의의 문자열
- `[abc]`: a, b, c 중 하나
- `[a-z]`: a부터 z까지
- `[!a-z]`: a부터 z 제외
- `{n}`: 정확히 n번 반복
- `{n,}`: n번 이상 반복
- `{n,m}`: n번~m번 반복

### 5. ContentControl

편집 가능한 콘텐츠 컨트롤

```typescript
// Content Control 생성
const contentControl = body.insertContentControl()
contentControl.tag = 'myControl'
contentControl.title = 'My Control'
contentControl.appearance = Word.ContentControlAppearance.boundingBox
contentControl.cannotEdit = false
contentControl.cannotDelete = false
contentControl.insertText('Content here', Word.InsertLocation.start)

// 기존 Content Control 찾기
const controls = body.contentControls.getByTag('myControl')
context.load(controls)
await context.sync()

if (controls.items.length > 0) {
  const control = controls.items[0]
  control.font.color = '#0000FF'
}
```

**주요 속성:**
- `tag`: 식별용 태그
- `title`: 제목
- `appearance`: 외관 (boundingBox, tags, hidden)
- `cannotEdit`: 편집 금지
- `cannotDelete`: 삭제 금지
- `color`: 색상
- `placeholderText`: 플레이스홀더 텍스트

**이벤트:**
- `onDataChanged`: 데이터 변경 시
- `onDeleted`: 삭제 시
- `onEntered`: 포커스 진입 시
- `onExited`: 포커스 이탈 시
- `onSelectionChanged`: 선택 변경 시

### 6. Table

표를 나타냅니다.

```typescript
// 테이블 생성
const data = [
  ['이름', '나이', '직업'],
  ['홍길동', '30', '개발자'],
  ['김철수', '25', '디자이너']
]

const table = body.insertTable(
  data.length,
  data[0].length,
  Word.InsertLocation.end,
  data
)

// 테이블 스타일
table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light

// 헤더 행 스타일
table.headerRowCount = 1
const headerRow = table.rows.items[0]
headerRow.font.bold = true
headerRow.font.color = '#FFFFFF'
headerRow.shadingColor = '#4472C4'

// 셀 접근 및 수정
const cell = table.getCell(1, 0)
cell.body.insertText('Updated name', Word.InsertLocation.replace)

// 행/열 추가
table.addRows(Word.InsertLocation.end, 2, [
  ['이름3', '28', '기획자'],
  ['이름4', '32', '마케터']
])
table.addColumns(Word.InsertLocation.end, 1, ['비고', '', ''])

// 행/열 삭제
table.deleteRows(2, 1) // 2번 행부터 1개 삭제
table.deleteColumns(3, 1) // 3번 열부터 1개 삭제

// 테이블 속성
table.width = 500 // 너비
table.alignment = Word.Alignment.centered
```

**주요 메서드:**
- `addRows()`: 행 추가
- `addColumns()`: 열 추가
- `deleteRows()`: 행 삭제
- `deleteColumns()`: 열 삭제
- `getCell()`: 셀 가져오기
- `clear()`: 내용 지우기
- `delete()`: 테이블 삭제

---

## 문서 구조 객체

### 7. Section

문서 섹션 (페이지 설정, 머리글/바닥글 등)

```typescript
// 첫 번째 섹션 가져오기
const section = context.document.sections.getFirst()

// 페이지 설정
section.pageWidth = 595 // A4: 595pt
section.pageHeight = 842 // A4: 842pt

// 여백 설정 (포인트 단위)
section.topMargin = 72
section.bottomMargin = 72
section.leftMargin = 72
section.rightMargin = 72

await context.sync()
```

### 8. Header / Footer

머리글 및 바닥글

```typescript
// 머리글 가져오기
const header = section.getHeader(Word.HeaderFooterType.primary)

// 머리글에 텍스트 추가
header.insertParagraph('문서 제목', Word.InsertLocation.start)

// 머리글 서식
header.font.size = 10
header.font.color = '#666666'

// 바닥글
const footer = section.getFooter(Word.HeaderFooterType.primary)
footer.insertParagraph('페이지 ', Word.InsertLocation.start)

// 짝수/홀수 페이지 다른 머리글
const evenHeader = section.getHeader(Word.HeaderFooterType.evenPages)
evenHeader.insertParagraph('짝수 페이지 머리글', Word.InsertLocation.start)

// 첫 페이지 다른 머리글
const firstHeader = section.getHeader(Word.HeaderFooterType.firstPage)
firstHeader.insertParagraph('첫 페이지 머리글', Word.InsertLocation.start)
```

**HeaderFooterType:**
- `primary`: 기본 (홀수 페이지)
- `evenPages`: 짝수 페이지
- `firstPage`: 첫 페이지

---

## 서식 및 스타일

### 9. Font

폰트 설정

```typescript
const range = context.document.getSelection()

// 기본 속성
range.font.name = 'Arial'
range.font.size = 14
range.font.bold = true
range.font.italic = true
range.font.underline = Word.UnderlineType.single
range.font.strikeThrough = true
range.font.subscript = false
range.font.superscript = false

// 색상
range.font.color = '#FF0000' // 글자색
range.font.highlightColor = '#FFFF00' // 하이라이트

// 간격
range.font.characterSpacing = 2 // 문자 간격 (포인트)

await context.sync()
```

**UnderlineType:**
- `none`, `single`, `double`, `dotted`, `dashed`, `wave`

### 10. List

번호 매기기 및 글머리 기호

```typescript
// 글머리 기호 목록 생성
const paragraph = body.insertParagraph('항목 1', Word.InsertLocation.end)
const list = paragraph.startNewList()

// 추가 항목
body.insertParagraph('항목 2', Word.InsertLocation.end)
  .attachToList(list.id, 0) // 레벨 0

body.insertParagraph('하위 항목', Word.InsertLocation.end)
  .attachToList(list.id, 1) // 레벨 1 (들여쓰기)

// 번호 매기기
const numberedPara = body.insertParagraph('번호 1', Word.InsertLocation.end)
numberedPara.startNewList()
numberedPara.listItem.listString // "1."
```

### 11. Style

사용자 정의 스타일

```typescript
// 기본 스타일 적용
paragraph.styleBuiltIn = Word.BuiltInStyleName.heading1

// 사용 가능한 기본 스타일
// - normal
// - heading1 ~ heading9
// - title, subtitle
// - quote, emphasis, strong
// - listParagraph
```

---

## 미디어 및 객체

### 12. InlinePicture

인라인 이미지

```typescript
// Base64 이미지 삽입
const base64Image = 'iVBORw0KGgoAAAANSUhEUg...' // Base64 문자열

const picture = body.insertInlinePictureFromBase64(
  base64Image,
  Word.InsertLocation.end
)

// 이미지 크기 조정
picture.width = 200
picture.height = 150

// 이미지 설명
picture.altTextTitle = '이미지 제목'
picture.altTextDescription = '이미지 설명'

// 이미지 하이퍼링크
picture.hyperlink = 'https://example.com'

await context.sync()
```

**주요 속성:**
- `width`, `height`: 크기
- `lockAspectRatio`: 비율 고정
- `altTextTitle`, `altTextDescription`: 대체 텍스트
- `hyperlink`: 하이퍼링크

---

## 문서 속성

### 13. DocumentProperties

문서 메타데이터

```typescript
const properties = context.document.properties
context.load(properties)
await context.sync()

// 읽기
console.log('제목:', properties.title)
console.log('주제:', properties.subject)
console.log('작성자:', properties.author)
console.log('카테고리:', properties.category)
console.log('설명:', properties.comments)
console.log('키워드:', properties.keywords)

// 수정
properties.title = '새 문서 제목'
properties.author = '홍길동'
properties.subject = '문서 주제'
await context.sync()
```

### 14. CustomProperty

사용자 정의 속성

```typescript
// 사용자 정의 속성 추가
const customProps = context.document.properties.customProperties

customProps.add('ProjectName', 'My Project')
customProps.add('Version', '1.0.0')
customProps.add('Status', 'Draft')

await context.sync()

// 사용자 정의 속성 읽기
const projectName = customProps.getItem('ProjectName')
context.load(projectName, 'value')
await context.sync()

console.log('프로젝트명:', projectName.value)

// 삭제
customProps.getItem('Status').delete()
await context.sync()
```

### 15. CustomXmlPart

사용자 정의 XML 데이터

```typescript
// XML 추가
const xmlString = '<data><name>홍길동</name><age>30</age></data>'
const customXml = context.document.customXmlParts.add(xmlString)

await context.sync()

// XML 읽기
const xmlParts = context.document.customXmlParts
context.load(xmlParts)
await context.sync()

const firstXml = xmlParts.items[0]
context.load(firstXml, 'xml')
await context.sync()

console.log('XML:', firstXml.xml)

// XML 삭제
firstXml.delete()
await context.sync()
```

---

## 이벤트

### 16. Document 이벤트

```typescript
// 단락 추가 이벤트
context.document.onParagraphAdded.add((eventArgs) => {
  console.log('단락 추가됨')
  return context.sync()
})

// 단락 변경 이벤트
context.document.onParagraphChanged.add((eventArgs) => {
  console.log('단락 변경됨')
  return context.sync()
})

// 단락 삭제 이벤트
context.document.onParagraphDeleted.add((eventArgs) => {
  console.log('단락 삭제됨')
  return context.sync()
})

// Content Control 추가 이벤트
context.document.onContentControlAdded.add(async (eventArgs) => {
  await Word.run(async (ctx) => {
    console.log('Content Control 추가됨')
    await ctx.sync()
  })
})
```

**Document 이벤트 목록:**
- `onParagraphAdded`: 단락 추가
- `onParagraphChanged`: 단락 변경
- `onParagraphDeleted`: 단락 삭제
- `onContentControlAdded`: Content Control 추가
- `onAnnotationClicked`: 주석 클릭
- `onAnnotationHovered`: 주석 호버
- `onAnnotationInserted`: 주석 삽입
- `onAnnotationRemoved`: 주석 삭제

### 17. ContentControl 이벤트

```typescript
const contentControl = body.contentControls.getByTag('myTag').getFirst()

// 데이터 변경 이벤트
contentControl.onDataChanged.add((eventArgs) => {
  console.log('Content Control 데이터 변경됨')
  return context.sync()
})

// 포커스 진입/이탈
contentControl.onEntered.add(() => console.log('진입'))
contentControl.onExited.add(() => console.log('이탈'))

// 삭제 이벤트
contentControl.onDeleted.add(() => console.log('삭제됨'))

await context.sync()
```

---

## 고급 기능

### 18. Annotation (주석)

```typescript
const paragraph = context.document.getSelection().paragraphs.getFirst()

// 주석 옵션
const popupOptions: Word.CritiquePopupOptions = {
  brandingTextResourceId: 'Brand',
  subtitleResourceId: 'Subtitle',
  titleResourceId: 'Title',
  suggestions: ['제안 1', '제안 2', '제안 3']
}

// 주석 생성
const critique: Word.Critique = {
  colorScheme: Word.CritiqueColorScheme.red,
  start: 0,
  length: 10,
  popupOptions: popupOptions
}

const annotationSet: Word.AnnotationSet = {
  critiques: [critique]
}

const annotationIds = paragraph.insertAnnotations(annotationSet)

await context.sync()
console.log('주석 ID:', annotationIds.value)
```

**CritiqueColorScheme:**
- `red`, `green`, `blue`, `lavender`, `berry`

### 19. Field (필드)

```typescript
// 페이지 번호 필드
const field = body.insertField(
  Word.InsertLocation.end,
  Word.FieldType.page
)

// 날짜 필드
const dateField = body.insertField(
  Word.InsertLocation.start,
  Word.FieldType.date
)

await context.sync()
```

**FieldType:**
- `page`: 페이지 번호
- `date`: 날짜
- `time`: 시간
- `author`: 작성자
- `fileName`: 파일명
- `numPages`: 전체 페이지 수

### 20. Bookmark (책갈피)

```typescript
// 책갈피 추가
const range = context.document.getSelection()
const bookmark = range.insertBookmark('myBookmark')

await context.sync()

// 책갈피로 이동
const bookmarks = context.document.body.bookmarks
const targetBookmark = bookmarks.getByName('myBookmark')
const bookmarkRange = targetBookmark.getRange()
bookmarkRange.select()

await context.sync()
```

---

## 패턴 및 모범 사례

### 기본 패턴

모든 Word API 작업은 `Word.run()` 안에서 실행됩니다:

```typescript
await Word.run(async (context) => {
  // 1. 객체 가져오기
  const body = context.document.body

  // 2. 작업 수행
  body.insertParagraph('Hello World', Word.InsertLocation.end)

  // 3. 속성 로드가 필요한 경우
  context.load(body, 'text')

  // 4. 동기화 (서버와 통신)
  await context.sync()

  // 5. 로드된 속성 사용
  console.log(body.text)
})
```

### context.load()와 context.sync()

**왜 필요한가?**

Word API는 **배치(batch) 방식**으로 동작합니다:
- 명령을 모아서 한 번에 실행 (성능 최적화)
- `context.sync()`를 호출할 때 실제로 실행됨

**올바른 사용:**

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  // ❌ 이렇게 하면 에러!
  // console.log(body.text) // Error: 아직 로드 안 됨

  // ✅ 올바른 방법
  context.load(body, 'text') // 로드할 속성 지정
  await context.sync()       // 실제 로드 실행
  console.log(body.text)     // 이제 사용 가능!
})
```

**성능 최적화:**

```typescript
// ❌ 비효율적 - 루프 안에서 sync()
for (let i = 0; i < items.length; i++) {
  items[i].font.bold = true
  await context.sync() // 매번 서버 통신!
}

// ✅ 효율적 - 한 번만 sync()
for (let i = 0; i < items.length; i++) {
  items[i].font.bold = true
}
await context.sync() // 한 번만 통신!
```

### 에러 처리

```typescript
async function safeWordOperation() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body
      body.insertText('Hello', Word.InsertLocation.end)
      await context.sync()
    })
  } catch (error) {
    console.error('Word API 에러:', error)

    if (error instanceof OfficeExtension.Error) {
      console.log('코드:', error.code)
      console.log('메시지:', error.message)
      console.log('상세:', error.debugInfo)
    }
  }
}
```

### 실전 예제 모음

**1. 선택된 텍스트 볼드 처리**

```typescript
async function makeBold() {
  await Word.run(async (context) => {
    const range = context.document.getSelection()
    range.font.bold = true
    await context.sync()
  })
}
```

**2. 제목 추가**

```typescript
async function insertHeading(text: string) {
  await Word.run(async (context) => {
    const body = context.document.body
    const paragraph = body.insertParagraph(text, Word.InsertLocation.start)
    paragraph.styleBuiltIn = Word.BuiltInStyleName.heading1
    await context.sync()
  })
}
```

**3. 텍스트 검색 및 하이라이트**

```typescript
async function highlightText(keyword: string) {
  await Word.run(async (context) => {
    const body = context.document.body
    const searchResults = body.search(keyword, {
      matchCase: false,
      matchWholeWord: true
    })

    context.load(searchResults)
    await context.sync()

    searchResults.items.forEach(item => {
      item.font.highlightColor = '#FFFF00'
    })

    await context.sync()

    console.log(`${searchResults.items.length}개 찾음`)
  })
}
```

**4. 테이블 생성**

```typescript
async function createTable() {
  await Word.run(async (context) => {
    const body = context.document.body

    const data = [
      ['이름', '나이', '직업'],
      ['홍길동', '30', '개발자'],
      ['김철수', '25', '디자이너']
    ]

    const table = body.insertTable(
      data.length,
      data[0].length,
      Word.InsertLocation.end,
      data
    )

    // 헤더 스타일
    table.headerRowCount = 1
    const headerRow = table.rows.items[0]
    headerRow.font.bold = true
    headerRow.font.color = '#FFFFFF'
    headerRow.shadingColor = '#4472C4'

    await context.sync()
  })
}
```

**5. HTML 삽입**

```typescript
async function insertHTML() {
  await Word.run(async (context) => {
    const body = context.document.body

    const html = `
      <h1>제목</h1>
      <p>이것은 <strong>볼드</strong>와 <em>이탤릭</em>이 있는 텍스트입니다.</p>
      <ul>
        <li>항목 1</li>
        <li>항목 2</li>
      </ul>
    `

    body.insertHtml(html, Word.InsertLocation.end)
    await context.sync()
  })
}
```

**6. 이미지 삽입**

```typescript
async function insertImage(base64Image: string) {
  await Word.run(async (context) => {
    const body = context.document.body
    const picture = body.insertInlinePictureFromBase64(
      base64Image,
      Word.InsertLocation.end
    )

    picture.width = 200
    picture.height = 150
    picture.altTextTitle = '이미지 제목'

    await context.sync()
  })
}
```

**7. 머리글/바닥글 설정**

```typescript
async function setHeaderFooter() {
  await Word.run(async (context) => {
    const section = context.document.sections.getFirst()

    // 머리글
    const header = section.getHeader(Word.HeaderFooterType.primary)
    header.clear()
    header.insertParagraph('회사명 - 문서 제목', Word.InsertLocation.start)

    // 바닥글
    const footer = section.getFooter(Word.HeaderFooterType.primary)
    footer.clear()
    footer.insertParagraph('페이지 ', Word.InsertLocation.start)

    await context.sync()
  })
}
```

---

## 참고 자료

- [Word JavaScript API 공식 레퍼런스](https://learn.microsoft.com/javascript/api/word)
- [Office Add-in 샘플](https://github.com/OfficeDev/Office-Add-in-samples)
- [Context7 라이브러리](/officedev/office-js-docs-pr)
