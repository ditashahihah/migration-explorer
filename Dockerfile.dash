FROM python:3.11-slim

WORKDIR /app

COPY requirements-dash.txt .
RUN pip install --no-cache-dir -r requirements-dash.txt

COPY . .

EXPOSE 8501

CMD ["gunicorn", "-b", "0.0.0.0:8501", "--workers", "2", "--timeout", "120", "dash_app:server"]
