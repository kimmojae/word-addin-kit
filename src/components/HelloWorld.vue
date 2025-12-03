<script setup lang="ts">
interface Props {
  msg: string
}

defineProps<Props>()
</script>

<template>
  <div class="text-center lg:text-left">
    <h1 class="green font-medium text-[2.6rem] relative -top-2.5">{{ msg }}</h1>
    <h3 class="text-xl">
      You've successfully created a project with
      <a href="https://vite.dev/" target="_blank" rel="noopener">Vite</a> +
      <a href="https://vuejs.org/" target="_blank" rel="noopener">Vue 3</a>.
    </h3>
  </div>
</template>
