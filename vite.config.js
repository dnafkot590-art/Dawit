import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    host: true,      // network ላይ ያሰራጫል — ስልክ ሊደርሰው ይችላል
    port: 5173,
    strictPort: false,
  },
});
