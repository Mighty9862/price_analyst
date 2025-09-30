# Multi-stage Dockerfile

# Этап 1: Сборка приложения
FROM eclipse-temurin:22-jdk AS builder
WORKDIR /app

# Устанавливаем maven
RUN apt-get update && apt-get install -y maven

# Копируем pom и исходники
COPY pom.xml .
COPY src ./src

# Сборка с кешем для ~/.m2
RUN --mount=type=cache,target=/root/.m2 \
    mvn clean package -DskipTests

# Этап 2: Runtime-образ для продакшена
FROM eclipse-temurin:22-jre-alpine
WORKDIR /app

# Создаём пользователя для запуска приложения
RUN addgroup -g 1001 -S appgroup && \
    adduser -S appuser -u 1001 -G appgroup

# Копируем собранный jar из builder stage
COPY --from=builder --chown=appuser:appgroup /app/target/*.jar app.jar

# Переключаемся на пользователя appuser
USER appuser

# Порт приложения
EXPOSE 8400

# Переменные окружения для JVM - оптимизации для больших данных
ENV JAVA_OPTS="-Xmx2g -Xms1g -XX:+UseG1GC -XX:MaxGCPauseMillis=200 -XX:ParallelGCThreads=4 -XX:ConcGCThreads=2"

# Запуск приложения
ENTRYPOINT ["sh", "-c", "java $JAVA_OPTS -jar /app/app.jar"]