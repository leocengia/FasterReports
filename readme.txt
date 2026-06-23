Faster Reports — Python Excel Automation
=========================================

Python 3.13 | pandas | openpyxl | xlsxwriter

Struttura
---------
raw_data/           Dati grezzi di input (CSV/XLSX)
reports_templates/  Template Excel di riferimento
src/                Codice Python per la generazione dei report
output/             Report Excel generati (escluso da git)

Setup
-----
Opzione A — doppio click (consigliata):
  setup.bat        crea il venv e installa tutte le dipendenze automaticamente

Opzione B — manuale:
1. python -m venv .venv
2. .venv\Scripts\activate
3. pip install -r requirements.txt

Uso
---
Ogni script in src/ legge da raw_data/ e scrive in output/.
Esempio: python src/agent_performance_score.py
