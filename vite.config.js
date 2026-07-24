import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { VitePWA } from "vite-plugin-pwa";

// Produkce běží na /vykazprce/, testovací verze na /vykazprce/test/.
// Cestu i prostředí nastavuje GitHub Actions přes proměnné VITE_BASE a VITE_APP_ENV.
const BASE = process.env.VITE_BASE || "/vykazprce/";
const IS_TEST = process.env.VITE_APP_ENV === "test";

export default defineConfig({
  base: BASE,
  plugins: [
    react(),
    VitePWA({
      // V testovací verzi service worker nezapínáme, aby se nepral s produkční instalací
      disable: IS_TEST,
      registerType: "autoUpdate",
      includeAssets: ["icon-192.png", "icon-512.png", "apple-touch-icon.png"],
      manifest: {
        name: "Výkaz práce – tým",
        short_name: "Výkazy",
        description: "Týmové vykazování pracovního času s exportem do firemní šablony",
        lang: "cs",
        display: "standalone",
        orientation: "portrait",
        theme_color: "#1C2530",
        background_color: "#F3F5F7",
        icons: [
          { src: "icon-192.png", sizes: "192x192", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png" },
          { src: "icon-512.png", sizes: "512x512", type: "image/png", purpose: "maskable" },
        ],
      },
      workbox: {
        globPatterns: ["**/*.{js,css,html,png,svg,ico,json}"],
        maximumFileSizeToCacheInBytes: 5 * 1024 * 1024,
        // Produkční service worker se nesmí míchat do testovací verze v podsložce
        navigateFallbackDenylist: [/\/test\//],
      },
    }),
  ],
});
