<template>
  <div
    class="item mt-8 flex relative lg:mt-0 lg:py-2 lg:pl-[calc(var(--spacing-section-gap)/2)] before:lg:content-[''] before:lg:border-l before:lg:border-border before:lg:absolute before:lg:left-0 before:lg:bottom-[calc(50%+25px)] before:lg:h-[calc(50%-25px)] first:before:lg:hidden after:lg:content-[''] after:lg:border-l after:lg:border-border after:lg:absolute after:lg:left-0 after:lg:top-[calc(50%+25px)] after:lg:h-[calc(50%-25px)] last:after:lg:hidden"
  >
    <i
      class="flex items-center justify-center w-8 h-8 text-foreground lg:absolute lg:top-[calc(50%-25px)] lg:-left-[26px] lg:border lg:border-border lg:bg-background lg:rounded-lg lg:w-[50px] lg:h-[50px]"
    >
      <slot name="icon"></slot>
    </i>
    <div class="flex-1 ml-4">
      <h3 class="text-xl font-medium mb-2 text-foreground">
        <slot name="heading"></slot>
      </h3>
      <slot></slot>
    </div>
  </div>
</template>
