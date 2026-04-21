# SDG Analyzer

## Project Overview

`SDG Analyzer` is a multimodal document and image analysis project developed to identify how strongly a given input aligns with the **17 United Nations Sustainable Development Goals (SDGs)**.

The system accepts:

- Text pasted directly into the application
- Documents such as `PDF`, `DOCX`, `TXT`, `MD`, and `HTML`
- Images such as `PNG`, `JPG`, `JPEG`, `BMP`, `TIFF`, and `WEBP`

The application then analyzes the extracted content and produces:

- Ranked SDG relevance scores
- A top SDG match with confidence
- Detected keywords
- Detected visual concepts from images
- Charts and visual summaries
- Exportable reports in `TXT` and `DOCX`

This project was designed as an academic-style intelligent analysis system with a modern interactive user interface, multimodal processing pipeline, and project-report-ready documentation.

---

## Project Purpose

The main purpose of this website is to help users quickly understand which SDGs are most relevant to a document, report, article, image, poster, or scanned material.

This can be useful in:

- Academic projects related to sustainability
- NGO and social impact documentation
- Policy analysis
- CSR reporting
- Educational demonstrations
- Research classification

## What We Developed
We developed a **full-stack SDG analysis web application** with:
- A Flask backend
- A responsive frontend
- OCR support for image text
- Vision-based image concept detection
- A weighted multimodal scoring system
- A polished Word report export

### What we actually built
- A **curated SDG knowledge base** using manually organized keyword and concept mappings
- A **rule-based text scoring engine**
- A **multimodal analysis pipeline**
- A **zero-shot vision integration** using a pretrained model

> Yes, the project uses AI-assisted multimodal analysis, but it is not a custom-trained LLM. It combines a curated SDG knowledge base, OCR, and a pretrained zero-shot vision model to classify document and image content against the 17 SDGs.

### Can we describe it as a knowledge base?
Yes. A better and more accurate term than "small language model database" is:
- **SDG knowledge base**
- **SDG concept repository**
- **curated SDG rule base**
- **multimodal SDG inference system**

> The project does not train a language model. Instead, it uses curated SDG keywords for text inference and a pretrained zero-shot vision model for image understanding. The final decision comes from a scoring and weighting engine designed specifically for SDG mapping.


















## Core Features
### 1. Text-Based SDG Analysis

The system reads textual input from uploaded files or pasted text and compares the content against SDG-specific keyword sets.

It returns:

- Matched keywords
- SDG-wise scores
- Ranked SDG results
- Summary of findings

### 2. OCR for Images

The system can read text from image files using OCR.

This helps analyze:

- Posters
- Scanned documents
- Infographics
- Screenshots
- Image-based reports

OCR is implemented using `easyocr`, which avoids the need for a separate Tesseract installation.

### 3. Vision-Based Image Understanding

The system does not rely only on OCR. It also inspects the visual content of an image.

Examples of detected concepts:

- Solar panels
- Wind turbines
- Factory
- Classroom
- Hospital
- Flood disaster
- Forest
- Crop field
- Recycling bins
- Ocean pollution

These visual concepts are then mapped to the relevant SDGs.

### 4. Multimodal Score Fusion

The final SDG ranking is not based on only one source.

The system combines:

- Text signal
- OCR-extracted text signal
- Visual concept signal

The current weighted logic is:

- If both text and image signals exist, both are blended
- If only text exists, text drives the result
- If only visual evidence exists, vision drives the result

This makes the project more robust than a plain keyword-only analyzer.

### 5. Interactive Dashboard

The frontend displays:

- Top SDG card
- Confidence score bar
- SDG ranking cards
- Keyword tags
- Visual concept tags
- Signal blend information
- Doughnut chart
- Bar chart

### 6. Word Report Export

The system can generate a polished `DOCX` report containing:

- Title and project heading
- Executive summary
- Highlights
- Ranked SDG table
- Keywords section
- Visual concepts section
- Chart images embedded directly into the report
- Notes section for OCR or vision limitations

This is useful for project submission, presentation, and documentation.











## Technologies Used
## Backend

- `Python`
- `Flask`
- `Flask-CORS`

## Text Extraction

- `pdfplumber` for PDF extraction
- `python-docx` for DOCX reading and DOCX export generation
- `BeautifulSoup4` for HTML extraction

## Image Processing

- `Pillow` for image preprocessing
- `easyocr` for OCR from images

## AI / Vision
- `transformers`
- `torch`
- Zero-shot image classification using a pretrained CLIP-based model

## Frontend
- `HTML5`
- `CSS3`
- `Vanilla JavaScript`
- `Chart.js`

## Styling / UI
- `DM Sans`
- `Italiana`
- `Orbitron`
- Glassmorphism / liquid-glass design concepts


### File Roles
- `app.py`
  Main Flask backend, routing, text extraction, OCR, vision analysis, score fusion, and report export.

- `index.html`
  Main frontend page containing the cinematic intro, upload interface, results dashboard, charts, export buttons, and theme system.

- `requirements.txt`
  Python dependency list for the project.

- `README.md`
  Complete project documentation.















## How the System Works
## Step 1: Input Acquisition

The user either:

- Uploads a file
- Uploads an image
- Pastes text directly

## Step 2: Content Extraction

Depending on the file type:

- `PDF` -> text extracted using `pdfplumber`
- `DOCX` -> paragraphs and table text extracted using `python-docx`
- `HTML` -> visible text extracted using `BeautifulSoup`
- `TXT/MD` -> decoded directly
- `Images` -> OCR text extracted using `easyocr`

## Step 3: Visual Concept Detection

For image inputs, the system also runs a zero-shot classifier using a curated list of sustainability-related concept labels.

## Step 4: SDG Mapping

The extracted text and detected visual concepts are mapped to SDGs using:

- `SDG_KEYWORDS`
- `VISION_CONCEPTS`

## Step 5: Scoring

Scores are computed for each SDG and ranked from highest to lowest.

## Step 6: Final Presentation

The system shows:

- Top SDG
- Confidence score
- Charts
- Keyword list
- Visual concept list
- Summary

## Step 7: Export

The result can be exported as:

- `TXT`
- `DOCX`








## Input Formats Supported
Currently supported:

- `PDF`
- `DOCX`
- `TXT`
- `MD`
- `HTML`
- `HTM`
- `PNG`
- `JPG`
- `JPEG`
- `BMP`
- `TIFF`
- `WEBP`





## API Endpoints

| Method | Route | Purpose |
|-------|-------|---------|
| `GET` | `/` | Loads the main website |
| `POST` | `/analyze` | Analyzes uploaded file or pasted text |
| `POST` | `/export/txt` | Exports the analysis as a text report |
| `POST` | `/export/json` | Exports raw JSON data |
| `POST` | `/export/docx` | Exports a formatted Word report |

---











## How to Run the Project

Open terminal in the project folder and run:

```powershell
1.)
python -m pip install -r requirements.txt
2.)
python app.py
```

Then open:

```text
http://localhost:5000
```

## Dependencies
The current project uses the following main dependencies:
- `flask>=3.0.0`
- `flask-cors>=4.0.0`
- `pdfplumber>=0.10.0`
- `python-docx>=1.1.0`
- `beautifulsoup4>=4.12.0`
- `lxml>=5.0.0`
- `Pillow>=10.0.0`
- `torch>=2.2.0`
- `transformers>=4.40.0`
- `easyocr>=1.7.1`





## Current Limitations
- Old `.doc` files are not directly supported
- Accuracy depends on the quality of keywords and concept mappings
- OCR may be weaker on very noisy or low-quality images
- Vision results depend on the pretrained model and label design
- No persistent database is currently used
- No user login or history tracking is included







## Viva Questions and Suggested Answers

### 1. Is this project based on machine learning?

Yes, partially. It uses a pretrained zero-shot vision model for image understanding, but the text analysis is mainly based on a curated SDG keyword knowledge base and scoring engine.

### 2. Did you build a language model?

No. We did not train a language model. We built a rule-based SDG analysis engine and combined it with OCR and a pretrained vision model.

### 3. Did you use a database?

No traditional database is used in the current version. The project relies on in-code SDG definitions, keyword mappings, and concept mappings.

### 4. Why did you use Flask?

Flask is lightweight, simple to integrate with Python-based text processing, and suitable for rapid development of a full-stack academic project.

### 8. What is the most innovative part of the project?

The combination of text analysis, OCR, visual concept detection, weighted score fusion, and polished report export inside one interactive website.







---




## Conclusion

`SDG Analyzer` is more than a simple document classifier. It is a multimodal sustainability intelligence tool that combines:

- Document understanding
- OCR
- Image understanding
- SDG-specific knowledge mapping
- Interactive visual presentation
- Report generation

Although it does not train its own language model or use a conventional database, it successfully demonstrates how curated knowledge, computer vision, and full-stack development can be combined to solve a real sustainability analysis problem in a practical and presentation-ready way.