# sample_status_separator

Appllication separates samples by status and saves to excel file

### create virtual environment

```bash
python -m venv venv
```

make shure your environment is activated

### requirements

```bash
pip install -r requirements.txt
```

### packaging using pyinstaller

write command below in terminal

```bash
pyinstaller --onefile --windowed --add-data 'blan.png' --icon=blan.ico main.py
```