# Word JavaScript API - 고급 기능

이 문서는 `word-api-complete.md`에서 다루지 않은 고급 기능들을 설명합니다.

## 목차
1. [특수 Content Control](#특수-content-control)
2. [Table 고급 기능](#table-고급-기능)
3. [Comment (댓글)](#comment-댓글)
4. [각주/미주 (NoteItem)](#각주미주-noteitem)
5. [문서 비교 및 검토](#문서-비교-및-검토)

---

## 특수 Content Control

기본 ContentControl 외에도 특수한 타입의 Content Control이 있습니다.

### CheckboxContentControl (체크박스)

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  // 체크박스 Content Control 삽입
  const checkbox = body.insertContentControl(Word.ContentControlType.checkBox)
  checkbox.title = '동의 여부'
  checkbox.appearance = Word.ContentControlAppearance.tags

  await context.sync()
})
```

### DatePickerContentControl (날짜 선택기)

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  const datePicker = body.insertContentControl(Word.ContentControlType.datePicker)
  datePicker.title = '날짜 선택'
  datePicker.placeholderText = '날짜를 선택하세요'

  await context.sync()
})
```

### DropDownListContentControl (드롭다운 목록)

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  const dropdown = body.insertContentControl(Word.ContentControlType.dropDownList)
  dropdown.title = '옵션 선택'
  dropdown.placeholderText = '선택하세요'

  // 옵션 추가
  dropdown.insertListItem('옵션 1', 'value1', Word.InsertLocation.end)
  dropdown.insertListItem('옵션 2', 'value2', Word.InsertLocation.end)
  dropdown.insertListItem('옵션 3', 'value3', Word.InsertLocation.end)

  await context.sync()
})
```

### ComboBoxContentControl (콤보박스)

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  const comboBox = body.insertContentControl(Word.ContentControlType.comboBox)
  comboBox.title = '콤보박스'
  comboBox.placeholderText = '선택 또는 입력'

  // 옵션 추가
  comboBox.insertListItem('항목 1', 'item1', Word.InsertLocation.end)
  comboBox.insertListItem('항목 2', 'item2', Word.InsertLocation.end)

  await context.sync()
})
```

### PictureContentControl (그림 컨트롤)

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  const pictureControl = body.insertContentControl(Word.ContentControlType.picture)
  pictureControl.title = '이미지 영역'

  await context.sync()
})
```

### BuildingBlockGalleryContentControl (빌딩 블록 갤러리)

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  const buildingBlock = body.insertContentControl(
    Word.ContentControlType.buildingBlockGallery
  )
  buildingBlock.title = '빌딩 블록'

  await context.sync()
})
```

**주요 속성:**
- `type`: ContentControl 타입
- `checked`: 체크박스의 체크 상태 (CheckboxContentControl만)
- `datePickerFormat`: 날짜 형식 (DatePickerContentControl만)
- `listItems`: 목록 항목 (DropDownList, ComboBox)

---

## Table 고급 기능

### TableRow / TableCell / TableColumn

```typescript
await Word.run(async (context) => {
  const body = context.document.body
  const table = body.tables.getFirst()

  // 행 가져오기
  const rows = table.rows
  context.load(rows)
  await context.sync()

  // 첫 번째 행
  const firstRow = rows.items[0]
  firstRow.font.bold = true
  firstRow.shadingColor = '#CCCCCC'

  // 특정 셀 접근
  const cell = firstRow.cells.items[0]
  cell.body.insertText('셀 내용', Word.InsertLocation.replace)
  cell.width = 100

  // 셀 병합
  const startCell = table.getCell(0, 0)
  const endCell = table.getCell(0, 2)
  const mergedCell = startCell.merge(endCell)

  await context.sync()
})
```

**TableRow 주요 메서드:**
- `insertRows()`: 행 삽입
- `delete()`: 행 삭제
- `merge()`: 셀 병합

**TableCell 주요 속성:**
- `body`: 셀 본문 (Body 객체)
- `width`: 셀 너비
- `columnWidth`: 열 너비
- `horizontalAlignment`: 가로 정렬
- `verticalAlignment`: 세로 정렬

### Table 스타일링

```typescript
await Word.run(async (context) => {
  const table = context.document.body.tables.getFirst()

  // 테두리 설정
  table.getBorder(Word.BorderLocation.top).type = Word.BorderType.single
  table.getBorder(Word.BorderLocation.top).width = 2
  table.getBorder(Word.BorderLocation.top).color = '#000000'

  // 모든 테두리
  table.getBorder(Word.BorderLocation.all).type = Word.BorderType.single

  // 셀 여백
  table.setCellPadding(
    Word.CellPaddingLocation.all,
    10 // 포인트 단위
  )

  await context.sync()
})
```

**BorderLocation:**
- `top`, `bottom`, `left`, `right`
- `insideHorizontal`, `insideVertical`
- `all`, `inside`, `outside`

---

## Comment (댓글)

Word API에서 Comment는 **미리보기(Preview) 상태**입니다.

### Comment 추가

```typescript
await Word.run(async (context) => {
  // 선택 영역에 댓글 추가
  const range = context.document.getSelection()
  const comment = range.insertComment('이 부분을 검토해주세요.')

  context.load(comment, 'id,content,authorName')
  await context.sync()

  console.log('댓글 ID:', comment.id)
  console.log('내용:', comment.content)
  console.log('작성자:', comment.authorName)
})
```

### Comment 읽기

```typescript
await Word.run(async (context) => {
  const comments = context.document.body.getComments()
  context.load(comments, 'items/id,items/content,items/authorName,items/creationDate')
  await context.sync()

  comments.items.forEach(comment => {
    console.log(`[${comment.authorName}] ${comment.content}`)
    console.log(`작성일: ${comment.creationDate}`)
  })
})
```

### CommentReply (댓글 답글)

```typescript
await Word.run(async (context) => {
  const comments = context.document.body.getComments()
  const firstComment = comments.getFirst()

  // 답글 추가
  const reply = firstComment.reply('동의합니다.')

  context.load(reply, 'content,authorName')
  await context.sync()

  console.log(`[${reply.authorName}] ${reply.content}`)
})
```

### Comment 삭제

```typescript
await Word.run(async (context) => {
  const comments = context.document.body.getComments()
  const firstComment = comments.getFirst()

  // 댓글 삭제 (모든 답글 포함)
  firstComment.delete()

  await context.sync()
})
```

### Comment 이벤트

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  // 댓글 추가 이벤트
  body.onCommentAdded.add((event) => {
    console.log('댓글 추가됨:', event.commentDetails)
    return context.sync()
  })

  // 댓글 변경 이벤트
  body.onCommentChanged.add((event) => {
    console.log('댓글 변경됨:', event.commentDetails)
    return context.sync()
  })

  // 댓글 삭제 이벤트
  body.onCommentDeleted.add((event) => {
    console.log('댓글 삭제됨:', event.commentDetails)
    return context.sync()
  })

  await context.sync()
})
```

**주요 이벤트:**
- `onCommentAdded`: 댓글 추가 시
- `onCommentChanged`: 댓글/답글 변경 시
- `onCommentDeleted`: 댓글 삭제 시
- `onCommentSelected`: 댓글 선택 시
- `onCommentDeselected`: 댓글 선택 해제 시

---

## 각주/미주 (NoteItem)

### 각주 추가

```typescript
await Word.run(async (context) => {
  const range = context.document.getSelection()

  // 각주 추가
  const footnote = range.insertFootnote('이것은 각주 내용입니다.')

  context.load(footnote, 'body/text,reference')
  await context.sync()

  console.log('각주 내용:', footnote.body.text)
  console.log('참조 번호:', footnote.reference)
})
```

### 미주 추가

```typescript
await Word.run(async (context) => {
  const range = context.document.getSelection()

  // 미주 추가
  const endnote = range.insertEndnote('이것은 미주 내용입니다.')

  context.load(endnote, 'body/text,reference')
  await context.sync()

  console.log('미주 내용:', endnote.body.text)
})
```

### 각주/미주 읽기

```typescript
await Word.run(async (context) => {
  const body = context.document.body

  // 모든 각주
  const footnotes = body.footnotes
  context.load(footnotes, 'items/body/text,items/reference')
  await context.sync()

  footnotes.items.forEach((note, index) => {
    console.log(`각주 ${index + 1}: ${note.body.text}`)
  })

  // 모든 미주
  const endnotes = body.endnotes
  context.load(endnotes, 'items/body/text')
  await context.sync()

  endnotes.items.forEach((note, index) => {
    console.log(`미주 ${index + 1}: ${note.body.text}`)
  })
})
```

### 각주/미주 삭제

```typescript
await Word.run(async (context) => {
  const footnotes = context.document.body.footnotes
  const firstFootnote = footnotes.getFirst()

  firstFootnote.delete()

  await context.sync()
})
```

**NoteItem 주요 속성:**
- `body`: 각주/미주 본문 (Body 객체)
- `reference`: 참조 번호
- `type`: 각주/미주 타입

---

## 문서 비교 및 검토

### TrackedChange (변경 내용 추적)

**참고:** TrackedChange API는 현재 **읽기 전용**입니다. 변경 내용 추적 켜기/끄기, 수락/거부는 아직 지원되지 않습니다.

```typescript
await Word.run(async (context) => {
  // 변경 내용 추적된 항목 읽기 (API가 있다면)
  // 현재는 이벤트를 통해 변경 감지만 가능

  context.document.onParagraphChanged.add((event) => {
    console.log('단락 변경됨')
    console.log('소스:', event.source) // 'local' 또는 'remote'
    return context.sync()
  })

  await context.sync()
})
```

### 문서 변경 이벤트

```typescript
await Word.run(async (context) => {
  // 단락 추가
  context.document.onParagraphAdded.add((event) => {
    console.log('단락 추가됨')
    return context.sync()
  })

  // 단락 변경
  context.document.onParagraphChanged.add((event) => {
    console.log('단락 변경됨')
    console.log('이벤트 소스:', event.source)
    return context.sync()
  })

  // 단락 삭제
  context.document.onParagraphDeleted.add((event) => {
    console.log('단락 삭제됨')
    return context.sync()
  })

  await context.sync()
})
```

**이벤트 소스 (event.source):**
- `local`: 로컬 사용자가 변경
- `remote`: 공동 작업자가 변경

### 공동 작업 (Coauthoring)

```typescript
await Word.run(async (context) => {
  // 공동 작업자 정보
  const authors = context.document.properties.author
  context.load(authors)
  await context.sync()

  console.log('작성자:', authors)
})
```

---

## 실전 예제

### 예제 1: 체크리스트 생성

```typescript
async function createChecklist(items: string[]) {
  await Word.run(async (context) => {
    const body = context.document.body

    for (const item of items) {
      const paragraph = body.insertParagraph('', Word.InsertLocation.end)
      const checkbox = paragraph.insertContentControl(Word.ContentControlType.checkBox)
      checkbox.title = item

      paragraph.insertText(` ${item}`, Word.InsertLocation.end)
    }

    await context.sync()
  })
}

// 사용
createChecklist([
  '요구사항 확인',
  '디자인 검토',
  '코드 구현',
  '테스트 완료'
])
```

### 예제 2: 양식 생성

```typescript
async function createForm() {
  await Word.run(async (context) => {
    const body = context.document.body

    // 제목
    const title = body.insertParagraph('신청서', Word.InsertLocation.end)
    title.styleBuiltIn = Word.BuiltInStyleName.heading1

    // 이름 입력
    body.insertParagraph('이름: ', Word.InsertLocation.end)
    const nameControl = body.insertContentControl(Word.ContentControlType.plainText)
    nameControl.title = '이름'
    nameControl.placeholderText = '이름을 입력하세요'

    // 날짜 선택
    body.insertParagraph('신청일: ', Word.InsertLocation.end)
    const dateControl = body.insertContentControl(Word.ContentControlType.datePicker)
    dateControl.title = '신청일'
    dateControl.placeholderText = '날짜 선택'

    // 옵션 선택
    body.insertParagraph('부서: ', Word.InsertLocation.end)
    const deptControl = body.insertContentControl(Word.ContentControlType.dropDownList)
    deptControl.title = '부서'
    deptControl.insertListItem('개발팀', 'dev', Word.InsertLocation.end)
    deptControl.insertListItem('디자인팀', 'design', Word.InsertLocation.end)
    deptControl.insertListItem('기획팀', 'planning', Word.InsertLocation.end)

    // 동의 체크박스
    const agreePara = body.insertParagraph('', Word.InsertLocation.end)
    const agreeCheck = agreePara.insertContentControl(Word.ContentControlType.checkBox)
    agreeCheck.title = '동의'
    agreePara.insertText(' 개인정보 수집에 동의합니다.', Word.InsertLocation.end)

    await context.sync()
  })
}
```

### 예제 3: 테이블 서식 자동화

```typescript
async function formatTable() {
  await Word.run(async (context) => {
    const table = context.document.body.tables.getFirst()

    // 헤더 행 스타일
    table.headerRowCount = 1
    const headerRow = table.rows.items[0]
    headerRow.font.bold = true
    headerRow.font.color = '#FFFFFF'
    headerRow.shadingColor = '#4472C4'
    headerRow.font.size = 12

    // 짝수 행 배경색
    const rows = table.rows
    context.load(rows)
    await context.sync()

    rows.items.forEach((row, index) => {
      if (index > 0 && index % 2 === 0) {
        row.shadingColor = '#F2F2F2'
      }
    })

    // 테두리
    table.getBorder(Word.BorderLocation.all).type = Word.BorderType.single
    table.getBorder(Word.BorderLocation.all).color = '#CCCCCC'

    // 너비 자동 조정
    table.width = 500
    table.alignment = Word.Alignment.centered

    await context.sync()
  })
}
```

---

## 제한사항 및 주의사항

### API 버전 확인

일부 기능은 특정 Word 버전에서만 지원됩니다:

```typescript
if (Office.context.requirements.isSetSupported('WordApi', '1.5')) {
  // 각주/미주 기능 사용 가능
  await insertFootnote()
} else {
  console.log('각주 기능은 WordApi 1.5 이상에서 지원됩니다.')
}
```

**주요 버전별 기능:**
- WordApi 1.1: 기본 기능
- WordApi 1.3: ContentControl, Table 확장
- WordApi 1.4: 문서 설정
- WordApi 1.5: 각주/미주, Style
- WordApi 1.7: Annotation

### Preview API 사용

```typescript
// Preview API는 프로덕션에서 주의해서 사용
if (Office.context.requirements.isSetSupported('WordApiPreview')) {
  // Comment API 등 사용
}
```

---

## 참고 자료

- [Word API 요구사항 세트](https://learn.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-api-requirement-sets)
- [Word API 미리보기](https://learn.microsoft.com/office/dev/add-ins/reference/requirement-sets/word-preview-apis)
- [Office Add-in 샘플](https://github.com/OfficeDev/Office-Add-in-samples)
