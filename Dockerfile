FROM python:3.12-slim
WORKDIR /app
COPY . .
RUN pip install --no-cache-dir -e .
RUN mkdir -p data/incoming data/outgoing data/backups
EXPOSE 8000
CMD ["uvicorn", "ndr_api.main:app", "--host", "0.0.0.0", "--port", "8000"]
