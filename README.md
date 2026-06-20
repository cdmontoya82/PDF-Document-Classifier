PDF Document Classifier
A local, privacy-preserving tool that classifies scanned PDF documents using

OCR and rule-based pattern matching. It reads each PDF, identifies its type

(memo, invoice, contract, promissory note, hotel reservation, and more),

extracts key metadata such as dates and document numbers, renames the file

according to a consistent convention, and produces an Excel summary report.
All processing runs locally — no documents leave your machine.
Features
·	Local OCR via PaddleOCR (no cloud calls)
·	Rule-based classification with a content-based fallback
·	Metadata extraction: dates, ID numbers, license plates, document numbers, vehicle units
·	Automatic, collision-safe file renaming
·	Excel report of every processed file, with a backup path if the report is locked
·	Parallel processing with a Windows/Anaconda-safe multiprocessing guard
Requirements
·	Python 3.9+
·	Poppler (required by pdf2image)
o	Windows (Anaconda): Poppler ships under …/Library/bin
o	macOS:brew install poppler
o	Linux:apt-get install poppler-utils
·	Python packages listed in requirements.txt
Installation
git clone https://github.com/<your-username>/pdf-classifier.git
cd pdf-classifier
pip install -r requirements.txt

Configuration
Configuration is driven by environment variables (with sensible defaults):
Variable	Description	Default
PDF_CLASSIFIER_BASE_DIR	Base working directory	./workspace next to the repo
POPPLER_PATH	Path to Poppler binaries (Windows/Anaconda)	System PATH

The base directory is expected to contain an unclassified/ folder with your

input PDFs. The tool creates classified/, not_classified/, reports/, and

logs/ automatically.
workspace/
├── unclassified/      # put your input PDFs here
├── classified/        # successfully classified & renamed files
├── not_classified/    # files that could not be classified
├── reports/           # Excel summary reports
└── logs/              # run logs

Usage
python src/classifier.py

Customizing classification rules
Classification is rule-driven. Edit CLASSIFICATION_RULES and SMART_RULES

in src/classifier.py to match your own document types and language. The

shipped patterns are examples; replace the keywords and naming formats with

ones that fit your dataset.
Notes on OCR language
The OCR model is configured for Spanish (lang="es"). Change the lang

argument in ocr_image_text() for other languages supported by PaddleOCR.
License
Released under the MIT License. See LICENSE.
