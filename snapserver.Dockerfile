FROM debian:bookworm-slim
RUN apt-get update && apt-get install -y curl libatomic1 && \
    curl -L -o /tmp/snapserver.deb \
      "https://github.com/badaix/snapcast/releases/download/v0.35.0/snapserver_0.35.0-1_amd64_bookworm.deb" && \
    apt-get install -y /tmp/snapserver.deb && \
    rm /tmp/snapserver.deb && \
    apt-get remove -y curl && \
    rm -rf /var/lib/apt/lists/*
EXPOSE 1704 1705 1780
CMD ["snapserver", "-c", "/etc/snapserver.conf"]
