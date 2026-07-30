# node:24 (not 20) because analytics.js uses the built-in node:sqlite module, which only
# stops requiring --experimental-sqlite from Node 22.13 onward. Keeping the zero-npm-
# dependency house style means the runtime has to supply SQLite itself.
FROM node:24-alpine
COPY app/OutsideFramework /app
COPY server.js /server.js
COPY auth.js /auth.js
COPY analytics.js /analytics.js
EXPOSE 8080
CMD ["node", "/server.js"]
