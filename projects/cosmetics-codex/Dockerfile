# node:24-alpine（不是 20）因為 store.js 用內建的 node:sqlite，Node 22.13 起才不必加
# --experimental-sqlite。維持本 repo「零 npm 依賴」的家法，SQLite 就得由 runtime 自己提供。
#
# ⚠️ COPY 路徑相對 repo 根：railway up 的建置脈絡是 git root（不是本資料夾），
# 由 service 變數 RAILWAY_DOCKERFILE_PATH 指到本檔案。根目錄的 .dockerignore
# 必須留有 `!projects/cosmetics-codex` 例外，否則以下 COPY 會找不到檔案。
#
# data/ 是初始 166 筆種子資料，第一次啟動時匯入 /data volume 的 SQLite；
# 之後以資料庫為準，種子不會覆蓋線上編輯過的內容（見 store.js 的 seedIfEmpty）。
FROM node:24-alpine
WORKDIR /app
COPY projects/cosmetics-codex/index.html /app/index.html
COPY projects/cosmetics-codex/server.js  /app/server.js
COPY projects/cosmetics-codex/store.js   /app/store.js
COPY projects/cosmetics-codex/data       /app/data
EXPOSE 8080
CMD ["node", "/app/server.js"]
