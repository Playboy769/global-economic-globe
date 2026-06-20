FROM nginx:1.25-alpine
COPY claude/OutsideFramework/index.html /usr/share/nginx/html/index.html
COPY claude/GlobalEco/index.html /usr/share/nginx/html/globe/index.html
COPY nginx.conf /etc/nginx/conf.d/default.conf
EXPOSE 8080
CMD ["nginx", "-g", "daemon off;"]
