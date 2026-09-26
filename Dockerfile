# ExcelFinder -- containerized CLI build.
#
# The engine is pure standard library, so the image needs no pip installs and
# builds fully offline once the base image is cached.
#
#   docker build -t excelfinder:1.0.0 .
#   docker run --rm -v "D:\报表:/data:ro" excelfinder:1.0.0 /data --query 预算
#
# For machines without internet, ship the image as a file (see build_docker.ps1):
#   docker save excelfinder:1.0.0 -o excelfinder-1.0.0.tar

FROM python:3.12-slim

LABEL org.opencontainers.image.title="ExcelFinder" \
      org.opencontainers.image.description="Fast local search across Excel/CSV files by file name and cell content" \
      org.opencontainers.image.version="1.0.0"

# Nothing in the engine shells out or needs a compiler.
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONIOENCODING=utf-8 \
    LANG=C.UTF-8

WORKDIR /opt/excelfinder
COPY src/excelfinder_core.py /opt/excelfinder/
COPY src/excelfinder_cli.py  /opt/excelfinder/

# Mount the folder you want to search here (read-only is enough).
VOLUME ["/data"]

# Runs as a console tool: pass a directory, then any options.
ENTRYPOINT ["python", "/opt/excelfinder/excelfinder_cli.py"]
CMD ["/data", "--query", "预算", "--limit", "20"]
