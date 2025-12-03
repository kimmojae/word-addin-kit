<script setup lang="ts">
import AppNavigation from '@/components/AppNavigation.vue'
import { useDark } from '@vueuse/core'

useDark()
</script>

<template>
  <AppNavigation />
  <RouterView />
</template>
