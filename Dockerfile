FROM python:3.12-slim

COPY requirements.txt ./
COPY dush.py ./

RUN pip3 install --no-cache-dir -r requirements.txt

ENV PYTHONUNBUFFERED=1

CMD ["python", "-u", "./dush.py" ]