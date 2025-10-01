# Этап сборки
FROM oven/bun:latest AS builder

WORKDIR /app

COPY . .

RUN bun install --frozen-lockfile
RUN bun run build

# Этап деплоя с nginx
FROM nginx:alpine

# Копируем собранные файлы фронта
COPY --from=builder /app/dist /usr/share/nginx/html

# Копируем конфигурацию nginx
COPY nginx.conf /etc/nginx/conf.d/default.conf

EXPOSE 3400

CMD ["nginx", "-g", "daemon off;"]