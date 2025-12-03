# Memory History in Office Add-ins

## 왜 Memory History를 사용하는가?

Office Add-in은 **iframe** 내부에서 실행되는 특수한 환경입니다. 이 환경에서는 일반적인 웹 애플리케이션과 다른 라우팅 전략이 필요합니다.

## Office Add-in의 실행 환경

```
Word Application
  └─ WebView/Browser Control
      └─ <iframe src="https://localhost:5174/index.html">
          └─ Vue App 실행
```

## Vue Router의 히스토리 모드 비교

### 1. createWebHistory (HTML5 History API)

**사용 예:**
```javascript
const router = createRouter({
  history: createWebHistory(),
  routes: [...]
})
```

**URL 변화:**
```
https://localhost:5174/
https://localhost:5174/settings
https://localhost:5174/about
```

**문제점:**
- ❌ iframe 내에서 URL이 변경되면 전체 페이지가 새로고침될 수 있음
- ❌ Office.js API 연결이 끊어짐
- ❌ Add-in이 재초기화되어 상태 손실

### 2. createWebHashHistory (Hash Mode)

**사용 예:**
```javascript
const router = createRouter({
  history: createWebHashHistory(),
  routes: [...]
})
```

**URL 변화:**
```
https://localhost:5174/#/
https://localhost:5174/#/settings
https://localhost:5174/#/about
```

**문제점:**
- ⚠️ 해시 변경이 Office 환경에서 예상치 못한 동작 유발 가능
- ⚠️ 일부 Office 버전에서 불안정
- ⚠️ URL이 변경되므로 여전히 리스크 존재

### 3. createMemoryHistory (Memory Mode) ✅ 권장

**사용 예:**
```javascript
const router = createRouter({
  history: createMemoryHistory(),
  routes: [...]
})
```

**URL 변화:**
```
https://localhost:5174/index.html (항상 동일)
```

**장점:**
- ✅ URL이 절대 변하지 않음
- ✅ 페이지 새로고침 없음
- ✅ Office.js 연결 안정적으로 유지
- ✅ 브라우저 뒤로가기/앞으로가기 버튼 영향 없음
- ✅ 모든 Office 버전에서 안정적

## Memory History 작동 원리

Memory History는 **브라우저 URL을 변경하지 않고** 메모리(JavaScript 객체)에만 히스토리를 저장합니다.

### 내부 구조 (개념)

```javascript
const memoryHistory = {
  stack: ['/'],           // 방문한 경로들을 배열로 저장
  currentIndex: 0         // 현재 위치
}
```

### 페이지 이동 시나리오

**1. 초기 상태**
```javascript
stack: ['/']
currentIndex: 0
URL: https://localhost:5174/index.html
렌더링: HomePage 컴포넌트
```

**2. router.push('/settings') 실행**
```javascript
stack: ['/', '/settings']
currentIndex: 1
URL: https://localhost:5174/index.html (변화 없음!)
렌더링: SettingsPage 컴포넌트
```

**3. router.push('/about') 실행**
```javascript
stack: ['/', '/settings', '/about']
currentIndex: 2
URL: https://localhost:5174/index.html (여전히 동일)
렌더링: AboutPage 컴포넌트
```

**4. router.back() 실행**
```javascript
stack: ['/', '/settings', '/about']
currentIndex: 1 (인덱스만 변경)
URL: https://localhost:5174/index.html (여전히 동일)
렌더링: SettingsPage 컴포넌트 (복원)
```

## 실제 코드 예시

### src/router/index.ts

```typescript
import { createRouter, createMemoryHistory } from 'vue-router'
import { setupLayouts } from 'virtual:generated-layouts'
import { routes } from 'vue-router/auto-routes'

const router = createRouter({
  // Office Add-in에서는 반드시 createMemoryHistory 사용
  history: createMemoryHistory(import.meta.env.BASE_URL),
  routes: setupLayouts(routes),
})

export default router
```

### 컴포넌트에서 사용

```vue
<template>
  <div>
    <button @click="goToSettings">설정으로 이동</button>
    <button @click="goBack">뒤로가기</button>
  </div>
</template>

<script setup lang="ts">
import { useRouter } from 'vue-router'

const router = useRouter()

// URL 변경 없이 컴포넌트만 교체됨
const goToSettings = () => {
  router.push('/settings')
}

const goBack = () => {
  router.back()
}
</script>
```

## 주의사항

### 1. 브라우저 네비게이션 버튼

Memory History를 사용하면 브라우저의 뒤로가기/앞으로가기 버튼이 **작동하지 않습니다**.
URL이 변경되지 않기 때문에 브라우저 히스토리에 기록되지 않습니다.

**해결책:**
- Add-in 내부에 자체 네비게이션 버튼 구현
- `router.back()`, `router.forward()` 사용

### 2. 북마크/URL 공유

Memory History는 URL이 변경되지 않으므로:
- ❌ 특정 페이지를 북마크할 수 없음
- ❌ URL을 복사해서 공유할 수 없음

하지만 Office Add-in 특성상 이런 기능이 필요하지 않으므로 문제되지 않습니다.

### 3. 새로고침 시 동작

페이지를 새로고침하면 항상 초기 페이지(`/`)로 돌아갑니다.
메모리에만 저장되므로 새로고침 시 히스토리가 초기화됩니다.

**필요 시 해결책:**
```typescript
// localStorage에 현재 경로 저장
router.afterEach((to) => {
  localStorage.setItem('lastRoute', to.path)
})

// 앱 시작 시 복원
const lastRoute = localStorage.getItem('lastRoute')
if (lastRoute) {
  router.replace(lastRoute)
}
```

## 요약

| 히스토리 모드 | URL 변경 | Office Add-in 호환성 | 권장 여부 |
|-------------|---------|-------------------|----------|
| createWebHistory | ✅ 변경 | ❌ 불안정 | ❌ |
| createWebHashHistory | ✅ 해시 변경 | ⚠️ 제한적 | ⚠️ |
| createMemoryHistory | ❌ 변경 없음 | ✅ 완벽 | ✅ |

**결론:** Office Add-in 개발 시 반드시 `createMemoryHistory`를 사용하세요.
