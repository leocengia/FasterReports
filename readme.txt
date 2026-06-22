Faster Reports — Python Excel Automation
=========================================

Python 3.13 | pandas | openpyxl | xlsxwriter

Struttura
---------
raw_data/           Dati grezzi di input (CSV/XLSX)
reports_templates/  Template Excel di riferimento
src/                Codice Python per la generazione dei report
output/             Report Excel generati (inclusi nei commit)

Setup
-----
1. python -m venv .venv
2. .venv\Scripts\activate
3. pip install -r requirements.txt

Uso
---
Ogni script in src/ legge da raw_data/ e scrive in output/.
Esempio: python src/agent_performance_score.py



tutti, ma in questo ordine: 
- DATASET: raw data grezzi;
- Case Type Analysis: la pivot serve solo per filtrare meglio il dataset ed effettuare eventuali check rapidi sul report. Le tabelle di case_type e Primary_category devono essere create dinamicamente. Servono per mostrare nei grafici a barre e linee (che trovi subito sotto nel foglio) i case types più rilevanti (ovvero quelli con la coppia più impattante di KPI utili. Di solito seleziono questi case types tramite un'altro dei report che voglio automatizzare "wow_CaseType_Deppdive", penso potrebbe essere utile includere anche quello nel più generico "case_KPI_report". Lo stile dei grafici combinati a barre e linee è salvato in un modello, che se vuoi posso fornirti senza problemi. 
- xLori & Costa (B. info): la parte del nome del foglio "(B. info)" deve cambiare in base al case type che giudichiamo più rilevante di settimana in settimana. Il foglio contiene tutti i casi di quel case type contenuti nei raw data e tutti i relativi parametri come da intestazione delle colonne;
- Dynamic Targets: questo è il foglio tutt'ora meno sviluppato, è fondamentale prima di tutto capire quello che voglio che faccia: voglio che su richiesta dell'utente (quando utilizza il file batch) possano essere estratti dei nuovi Dynamic Target per i case type più problematici (tipo i peggiori 3 case type . Il metodo di Dynamic Target Acquisition che ho pensato, per ora, utilizza i raw data per fare la media pesata dei 15 migliori in un determinato case type problematico
