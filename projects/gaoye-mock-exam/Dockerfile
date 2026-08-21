# node:24-alpine（不是 20）因為 store.js 用內建的 node:sqlite，Node 22.13 起才不必加
# --experimental-sqlite。維持本 repo「零 npm 依賴」的家法，SQLite 就得由 runtime 自己提供。
#
# ⚠️ COPY 路徑相對 repo 根：railway up 的建置脈絡是 git root（不是本資料夾），
# 由 service 變數 RAILWAY_DOCKERFILE_PATH 指到本檔案。根目錄的 .dockerignore
# 必須留有 `!projects/gaoye-mock-exam` 例外，否則這三個 COPY 會找不到檔案。
FROM node:24-alpine
WORKDIR /app
COPY projects/gaoye-mock-exam/index.html /app/index.html
COPY projects/gaoye-mock-exam/server.js  /app/server.js
COPY projects/gaoye-mock-exam/auth.js    /app/auth.js
COPY projects/gaoye-mock-exam/store.js   /app/store.js
EXPOSE 8080
CMD ["node", "/app/server.js"]
