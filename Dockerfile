# The remote server (tracker-mcp remote) for Cloud Run or any container host.
# Local use does not need this: run `uvx tracker-mcp` instead.
FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1 PIP_DISABLE_PIP_VERSION_CHECK=1
WORKDIR /app
COPY pyproject.toml README.md LICENSE ./
COPY src ./src
COPY examples ./examples
RUN pip install . && useradd --system --uid 10001 tracker
USER tracker
ENV TRACKER_MCP_DATA_DIR=/data PORT=8080
EXPOSE 8080
CMD ["tracker-mcp", "remote"]
