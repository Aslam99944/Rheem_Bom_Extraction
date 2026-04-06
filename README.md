# BOM Extractor (Hybrid Vision Pipeline)

A production-ready web application that extracts Bill of Materials (BOM) data from complex engineering documents with industry-grade accuracy. It leverages a **Hybrid Vision Pipeline** combining **Azure Document Intelligence** for structured data and **GPT-4o Vision** for spatial diagram understanding.

## Features
- **Hybrid Extraction:** Processes both structured text (from Azure DI) and page images (from PyMuPDF) in a single unified AI model.
- **Diagram Understanding:** Correctly assigns dimensions and callouts to wires and components based on their spatial orientation in diagrams.
- **Automated Wire Summation:** Automatically sums lengths of identical wire types (same AWG and color) from both tables and diagram labels.
- **Strict Formatting:** Standardized uppercase descriptions and specialized wire naming (e.g., `18AWG,RED WIRE`).
- **Description Cleaning:** Automatically strips embedded part numbers and manufacturer names from descriptions into their respective columns.
- **Standardized Excel Export:** Professional reports with drawing metadata headers and sorted/validated BOM items.
- **Token Metrics:** Real-time tracking of API usage and estimated costs for both text and vision tokens.

---

## Tech Stack & Libraries

### Backend

| Library | Purpose |
|---------|---------|
| **[FastAPI](https://fastapi.tiangolo.com/)** | High-performance async web framework |
| **[PyMuPDF (fitz)](https://pypi.org/project/PyMuPDF/)** | High-quality PDF-to-image conversion for Vision |
| **[azure-ai-documentintelligence](https://pypi.org/project/azure-ai-documentintelligence/)** | Native parsing of layouts and tables |
| **[openai](https://github.com/openai/openai-python)** | Unified interface for GPT-4o Vision |
| **[openpyxl](https://openpyxl.readthedocs.io/)** | Professional Excel report generation |

### Frontend

| Technology | Purpose |
|------------|---------|
| **HTML5/CSS3** | Premium dark-mode UI with glassmorphism |
| **Vanilla JS** | Real-time pipeline visualization and async uploads |

---

## Technical Architecture

The application uses a **Dual-Input Hybrid Pipeline**:

1. **Azure DI Layout**: Extracts high-accuracy table cells, paragraphs, and annotations as structured text.
2. **Page Conversion**: PyMuPDF converts every PDF page into a high-resolution PNG image at 200 DPI.
3. **GPT-4o Vision**: Receives BOTH the text and the images. It uses text for character-perfect part numbers and images for spatial context (like wire routing).
4. **Validation Engine**: Cleans, upcases, sums, and sorts the data before final export.

---

## Formatting Standards

To ensure production consistency, the system enforces:
- **Descriptions**: All component descriptions are converted to **UPPERCASE**.
- **Wire Format**: `{AWG}AWG,{COLOR} WIRE` (e.g., `18AWG,RED WIRE`).
- **Cleaning**: "TE", "Molex", or part numbers are automatically stripped from descriptions once extracted to their own columns.
- **Labels**: General labels are standardized to simply `LABEL`.

---

## Project Structure

```
BOM_Extraction_POC/
├── app.py                  # Hybrid Pipeline (DI + Vision + Validation)
├── static/                 # Frontend UI (index.html, style.css, script.js)
├── .env                    # Azure & OpenAI Credentials
├── requirements.txt        # Backend dependencies
├── README.md               # Documentation
├── uploads/                # Temporary PDF storage
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
3. **Process**: Upload an engineering drawing and watch the hybrid pipeline in action.
