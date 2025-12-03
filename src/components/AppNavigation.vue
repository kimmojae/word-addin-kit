<script setup lang="ts"></script>

<template>
  <header
    class="max-h-screen leading-relaxed lg:flex lg:items-center lg:pr-[calc(var(--section-gap)/2)]"
  >
    <div class="flex justify-end">
      <ThemeToggle />
    </div>

    <div class="lg:flex lg:flex-wrap lg:items-start lg:gap-8 lg:w-full">
      <nav class="w-full text-center text-xs lg:text-left lg:-ml-4 lg:text-base lg:py-4 lg:mt-4">
        <RouterLink
          to="/"
          class="inline-block px-4 border-l border-border first:border-l-0 text-gray-400 dark:text-gray-400 transition-colors hover:text-foreground [&.router-link-exact-active]:text-accent [&.router-link-exact-active]:font-semibold"
        >
          Home
        </RouterLink>
        <RouterLink
          to="/examples/basic"
          class="inline-block px-4 border-l border-border first:border-l-0 text-gray-400 dark:text-gray-400 transition-colors hover:text-foreground [&.router-link-active]:text-accent [&.router-link-active]:font-semibold"
        >
          Basic
        </RouterLink>
        <RouterLink
          to="/examples/structure"
          class="inline-block px-4 border-l border-border first:border-l-0 text-gray-400 dark:text-gray-400 transition-colors hover:text-foreground [&.router-link-active]:text-accent [&.router-link-active]:font-semibold"
        >
          Structure
        </RouterLink>
        <RouterLink
          to="/examples/controls"
          class="inline-block px-4 border-l border-border first:border-l-0 text-gray-400 dark:text-gray-400 transition-colors hover:text-foreground [&.router-link-active]:text-accent [&.router-link-active]:font-semibold"
        >
          Controls
        </RouterLink>
        <RouterLink
          to="/examples/advanced"
          class="inline-block px-4 border-l border-border first:border-l-0 text-gray-400 dark:text-gray-400 transition-colors hover:text-foreground [&.router-link-active]:text-accent [&.router-link-active]:font-semibold"
        >
          Advanced
        </RouterLink>
      </nav>
    </div>
  </header>
</template>
