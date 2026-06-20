FROM nginx:1.25-alpine
# Railway injects $PORT at runtime; default to 8080 for local `docker run`.
ENV PORT=8080
COPY claude/OutsideFramework/index.html /usr/share/nginx/html/index.html
COPY claude/GlobalEco/index.html /usr/share/nginx/html/globe/index.html
# Placed in templates/ so the nginx entrypoint runs envsubst (only ${PORT}) into conf.d/.
COPY nginx.conf /etc/nginx/templates/default.conf.template
EXPOSE 8080
