# BOM Extractor (Azure Document Intelligence + GPT-4o)

A production-ready web application that extracts Bill of Materials (BOM) data from PDF documents and images with industry-grade accuracy. It leverages **Azure Document Intelligence (Layout Model)** for native table understanding and **Azure OpenAI GPT-4o** for intelligent field mapping and data inference.

## Features
- **File Support:** Upload PDF, PNG, JPG, and TIFF files.
- **Azure Document Intelligence:** Native parsing of complex engineering tables, wire harness diagrams, and notes with high spatial accuracy.
- **GPT-4o Intelligence:** 
  - Dynamic mapping of structured text to BOM schema.
  - Intelligent inference of Commodity, Type, and Unit of Measure (UOM).
  - Handles wire color code mapping (e.g., OG → Orange).
  - Integrates tolerances and alternate parts from drawing notes.
- **Structured Output:** Guaranteed JSON schema via GPT-4o's `json_object` format.
- **Standardized Excel Export:** Generates styled Excel files with drawing metadata (Drawing No, Name) and sorted BOM items.
- **Token Metrics:** Real-time tracking of API usage and estimated costs.
- **Modern UI:** Premium dark theme with glassmorphism and real-time pipeline visualization.

---

## Tech Stack & Libraries

### Backend

| Library | Purpose |
|---------|---------|
| **[FastAPI](https://fastapi.tiangolo.com/)** | Async web framework for API endpoints |
| **[azure-ai-documentintelligence](https://pypi.org/project/azure-ai-documentintelligence/)** | Azure SDK for layout and table extraction |
| **[openai](https://github.com/openai/openai-python)** | SDK for Azure OpenAI GPT-4o extraction |
| **[openpyxl](https://openpyxl.readthedocs.io/)** | Professional Excel report generation |
| **[python-dotenv](https://pypi.org/project/python-dotenv/)** | Environment variable management |

### Frontend

| Technology | Purpose |
|------------|---------|
| **HTML5/CSS3** | Modern, responsive dark-mode UI |
| **Vanilla JS** | Async upload and real-time status updates |

---

## Prerequisites

### Azure Services
1. **Azure Document Intelligence**: A resource in the Azure Portal (F0 Free tier or S0).
2. **Azure OpenAI**: A resource with a **GPT-4o** deployment.

---

## Project Structure

```
BOM_Extraction_POC/
├── app.py                  # Main Application (Azure DI + GPT-4o pipeline)
├── static/                 # Frontend assets (HTML, CSS, JS)
├── .env                    # Credentials and Config
├── requirements.txt        # Python dependencies
├── README.md               # This file
├── uploads/                # Temporary storage for uploaded files
└── outputs/                # Generated Excel reports
```

---

## Installation & Setup

### Step 1: Clone and Navigate
```powershell
cd c:\Users\40045610\AI\Bom_Extraction_POC
```

### Step 2: Setup Environment
```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

### Step 3: Configure `.env`
Ensure your `.env` contains:
```env
AZURE_OPENAI_ENDPOINT=...
AZURE_OPENAI_API_KEY=...
AZURE_OPENAI_DEPLOYMENT=...
AZURE_OPENAI_API_VERSION=2024-12-01-preview

AZURE_DI_ENDPOINT=...
AZURE_DI_KEY=...
```

---

## Running the Application

1. **Start Server**: `.venv\Scripts\python.exe app.py`
2. **Access Web UI**: [http://localhost:8000](http://localhost:8000)

---

## Extraction Workflow

1. **Azure DI Layout**: Extracts raw text while preserving table structures and spatial headers.
2. **GPT-4o Mapping**: A unified prompt processes the structured text to identify parts, wires, and metadata.
3. **Validation**: Deduplication and sorting (1, 2, 3...) of extracted items.
4. **Export**: Styled Excel with Drawing Number and Name in the header.

### Extracted Fields

| Field | Description |
|-------|-------------|
| Drawing No/Name | Extracted from the title block |
| Item | Sorted line number |
| Part Number | Extracted part codes (validated) |
| Manufacturer | Standardized brand names |
| Description | Human-readable part names |
| Qty | Quantities with associated tolerances |
| Commodity/Type | Inferred based on industry standards |
| Notes | Connector mappings and alternate parts |
