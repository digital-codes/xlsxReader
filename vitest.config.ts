import { defineConfig } from 'vitest/config';

export default defineConfig({
  assetsInclude: ['**/*.xlsx'],
  test: {
    environment: 'jsdom', // needed for domparser to parse xml content
    globals: true,         // Enable global test methods like describe, it
    // environment: 'happy-dom', // or 'jsdom'
    //setupFiles: './test/setup.ts', // Include setup file
    coverage: {
      reporter: ['text', 'html'],
    },
    browser: {
      provider: "webdriverio", //'playwright', // or 'webdriverio'
      enabled: true,
      headless: true
      // name: 'firefox', // browser name is required
    },

  },
});

