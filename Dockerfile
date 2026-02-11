FROM python:3.11-slim
1
WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY bitrix_bot.py .
COPY main.py .
COPY LabelsMarksGenerator/ ./LabelsMarksGenerator/

RUN mkdir -p LabelsMarksGenerator/img/logos \
    LabelsMarksGenerator/img/certificates \
    LabelsMarksGenerator/img/mark_images \
    LabelsMarksGenerator/input \
    LabelsMarksGenerator/output

EXPOSE 5000

CMD ["python", "bitrix_bot.py"]