FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV PIP_NO_CACHE_DIR=1

WORKDIR /app

COPY requirements_dashboard.txt /app/requirements_dashboard.txt
RUN pip install -r /app/requirements_dashboard.txt
RUN python -c "import PIL, streamlit, plotly, pandas, xlrd, openpyxl, matplotlib; print('dashboard deps OK')"

COPY . /app

EXPOSE 10000

CMD ["sh", "-c", "streamlit run dashboard_achados_criticos.py --server.address 0.0.0.0 --server.port ${PORT:-10000} --server.headless true --server.enableCORS=false --server.enableXsrfProtection=false --server.maxUploadSize=400 --browser.gatherUsageStats false --theme.base dark --theme.primaryColor '#667eea' --theme.backgroundColor '#0e1117' --theme.secondaryBackgroundColor '#262730'"]
