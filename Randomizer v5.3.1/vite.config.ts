import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import fs from 'fs';
import {defineConfig, loadEnv} from 'vite';
import { viteSingleFile } from 'vite-plugin-singlefile';

// 离线文件生成插件
function generateOfflineFile() {
  return {
    name: 'generate-offline-file',
    closeBundle() {
      const distPath = path.resolve(process.cwd(), 'dist');
      const indexPath = path.join(distPath, 'index.html');
      const targetFileName = 'offline.html';
      const version = 'V5.3.1';
      const rootTargetPath = path.join(process.cwd(), targetFileName);
      const distTargetPath = path.join(distPath, targetFileName);
      const publicPath = path.join(process.cwd(), 'public');
      const publicTargetPath = path.join(publicPath, targetFileName);

      if (!fs.existsSync(indexPath)) return;

      let content = fs.readFileSync(indexPath, 'utf-8');
      
      // 离线化处理
      let processed = content
        .replace(/\s+crossorigin(="[^"]*")?/gi, '')
        .replace(/href="\//g, 'href="./')
        .replace(/src="\//g, 'src="./');

      // 注入离线标记和版本号
      processed = processed.replace('</body>', `<script>window.__OFFLINE_VERSION__ = true; window.__APP_VERSION__ = "Randomizer V5.3.1";</script></body>`);
      processed = processed.replace('<title>无序打乱器</title>', `<title>无序打乱器</title>`);

      // 1. 写入根目录 (用于 ZIP 导出)
      fs.writeFileSync(rootTargetPath, processed);
      
      // 2. 写入 dist 目录 (用于生产环境下载)
      fs.writeFileSync(distTargetPath, processed);
      
      // 3. 写入 public 目录 (用于本地预览)
      if (!fs.existsSync(publicPath)) fs.mkdirSync(publicPath, { recursive: true });
      fs.writeFileSync(publicTargetPath, processed);

      console.log(`\n✅ 离线版已生成: ${targetFileName}\n`);
    }
  };
}

export default defineConfig(({mode}) => {
  const env = loadEnv(mode, '.', '');
  return {
    base: '/',
    build: {
      assetsInlineLimit: 100000000,
      chunkSizeWarningLimit: 100000000,
      cssCodeSplit: false,
      modulePreload: {
        polyfill: false,
      },
      rollupOptions: {
        output: {
          inlineDynamicImports: true,
        },
      },
    },
    plugins: [
      react(), 
      tailwindcss(), 
      mode === 'production' ? viteSingleFile() : null, 
      mode === 'production' ? generateOfflineFile() : null
    ].filter(Boolean) as any,
    define: {
      'process.env.GEMINI_API_KEY': JSON.stringify(env.GEMINI_API_KEY || ""),
    },
    resolve: {
      alias: {
        '@': path.resolve(__dirname, '.'),
      },
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modifyâfile watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
    },
  };
});
