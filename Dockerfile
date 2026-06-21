FROM nginx:1.25-alpine
ENV PORT=8080
COPY app/OutsideFramework/index.html /usr/share/nginx/html/index.html
COPY nginx.conf /etc/nginx/templates/default.conf.template
EXPOSE 8080
