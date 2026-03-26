import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import basicSsl from '@vitejs/plugin-basic-ssl'

export default defineConfig({
  plugins: [
    react(),
    basicSsl() // This creates the secure connection needed for the camera
  ],
  server: {
    host: true // This allows you to access the site from your phone on the same Wi-Fi
  }
})