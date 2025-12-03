import { computed } from 'vue'

/**
 * Vue composable for Office.js integration
 * Provides reactive access to Office context and common Word operations
 */
export function useOffice() {
  /**
   * Check if running in Office context
   */
  const isOffice = computed(() => {
    return typeof Office !== 'undefined' && Office.context !== undefined
  })

  /**
   * Get Office context
   */
  const context = computed(() => {
    return isOffice.value ? Office.context : null
  })

  /**
   * Insert text at current cursor position in Word
   */
  async function insertText(text: string): Promise<void> {
    if (!isOffice.value) {
      throw new Error('Office.js not available')
    }

    return Word.run(async (context) => {
      const selection = context.document.getSelection()
      selection.insertText(text, Word.InsertLocation.replace)
      await context.sync()
    })
  }

  /**
   * Get selected text from Word document
   */
  async function getSelectedText(): Promise<string> {
    if (!isOffice.value) {
      throw new Error('Office.js not available')
    }

    return Word.run(async (context) => {
      const selection = context.document.getSelection()
      selection.load('text')
      await context.sync()
      return selection.text
    })
  }

  /**
   * Run Word API batch operation
   */
  async function runWordBatch<T>(
    callback: (context: Word.RequestContext) => Promise<T>,
  ): Promise<T> {
    if (!isOffice.value) {
      throw new Error('Office.js not available')
    }

    return Word.run(callback)
  }

  return {
    // State
    isOffice,
    context,

    // Common operations
    insertText,
    getSelectedText,
    runWordBatch,
  }
}
