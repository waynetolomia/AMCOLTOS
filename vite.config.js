import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  // Replace 'TOS-maker' with your actual GitHub repository name!
  base: '/TOS-maker/', 
})