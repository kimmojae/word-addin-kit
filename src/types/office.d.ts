/// <reference types="@types/office-js" />

/**
 * Global Office.js type augmentations
 */

declare global {
  interface Window {
    Office: typeof Office
  }
}

/**
 * Custom Office-related types
 */

export interface OfficeEnvironment {
  isOffice: boolean
  host: Office.HostType | null
  platform: Office.PlatformType | null
}

export interface WordSelection {
  text: string
  html: string
  isEmpty: boolean
}

export {}
