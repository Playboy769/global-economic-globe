# 零 npm 依賴，node:24-alpine 內建 fetch 已足夠代理政府公開資料源。
# ⚠️ railway up 的建置脈絡是 git root（不是本資料夾），由 service 變數
# RAILWAY_DOCKERFILE_PATH 指到本檔案。根目錄的 .dockerignore 必須留有
# `!research/macro-tracker-2026` 例外，否則以下 COPY 會找不到檔案。
FROM node:24-alpine
WORKDIR /app
COPY research/macro-tracker-2026/index.html /app/index.html
COPY research/macro-tracker-2026/server.js  /app/server.js
EXPOSE 8080
CMD ["node", "/app/server.js"]
