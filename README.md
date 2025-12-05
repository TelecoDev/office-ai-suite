# OfficeAI Suite  
### Local AI Add-ins for Word, Excel and Outlook (Powered by Ollama)

[![Status](https://img.shields.io/badge/status-active-brightgreen.svg)]()
[![License](https://img.shields.io/badge/license-private-lightgrey.svg)]()
[![Ollama](https://img.shields.io/badge/Ollama-local%20AI-blue.svg)]()
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue.svg)]()
[![Office](https://img.shields.io/badge/Microsoft_Office-Word%20%7C%20Excel%20%7C%20Outlook-orange.svg)]()
[![Made by TelecoDev](https://img.shields.io/badge/made%20by-TelecoDev-black.svg)]()

---

OfficeAI Suite è una raccolta di componenti aggiuntivi **completamente locali** per Microsoft Office, progettati per integrare potenti funzionalità AI senza inviare alcun dato in cloud.

La suite utilizza **Ollama** come motore AI interno, garantendo:
- riservatezza completa dei dati,
- prestazioni elevate,
- nessuna dipendenza da servizi esterni.

---

# 🧠 Funzionalità principali

### 📝 Word Add-in
- Generazione contenuti professionali  
- Riscrittura e ottimizzazione testi  
- Riassunti contestuali  
- Inserimento diretto nel documento  
- Zero cloud, tutto locale  

### 📊 Excel Add-in
- Funzioni AI per analisi testuale  
- Supporto contestuale alle celle  
- Taskpane intelligente  
- Generazione automatizzata di testo  

### 📧 Outlook Add-in
- Generazione email professionali  
- Riscrittura tono e stile  
- Riassunto conversazioni email  
- Inserimento automatico nel corpo messaggio  

---

# 🧱 Architettura Tecnica

        ┌──────────────────────────────┐
        │          OfficeAI Suite      │
        │  Word | Excel | Outlook Add-ins  │
        └──────────────────────────────┘
                       │
                       ▼
            ┌───────────────────┐
            │   Office JS API   │
            └───────────────────┘
                       │
                       ▼
      ┌─────────────────────────────────────┐
      │      Backend Locale (Ollama)        │
      │  - LLaMA3 8B / 12B                  │
      │  - API HTTP su http://localhost     │
      └─────────────────────────────────────┘
                       │
                       ▼
          Nessun dato lascia il sistema

---

# 📂 Struttura repository

office-ai-suite/
│
├── word-addin/ # Add-in Word (React + Office JS)
├── excel-addin/ # Add-in Excel (React + Office JS)
└── outlook-addin/ # Add-in Outlook (React + Office JS)

Ogni add-in è indipendente e contiene:
- `manifest.xml`
- `package.json`
- Taskpane React
- Comandi personalizzati
- Webpack config

---

# 📦 Requisiti

- Node.js LTS  
- Yeoman Office Generator  
- Microsoft Office Desktop  
- Ollama (modello consigliato: LLaMA 3 – 8B)  
- Windows 10 / 11  

---

# ⚙️ Setup ambiente sviluppo

### 1️⃣ Clona la repository

git clone https://github.com/TelecoDev/office-ai-suite.git
cd office-ai-suite
2️⃣ Installa le dipendenze per ogni add-in
Word

cd word-addin
npm install
npm start

Excel

cd excel-addin
npm install
npm start

Outlook

cd outlook-addin
npm install
npm start

🔒 Privacy & Sicurezza
OfficeAI Suite è pensata per ambienti aziendali:

I dati non lasciano mai la macchina locale

Nessun traffico verso servizi cloud

Nessuna dipendenza da OpenAI o API esterne

Perfect-fit per contesti ISO 27001

🛣 Roadmap
 Miglioramento UI con Fluent Design

 Integrazione modello selezionabile dinamicamente

 Logging locale richieste AI

 Add-in PowerPoint

 Setup automatico tramite installer

🔐 Licenza
Repository privata.
