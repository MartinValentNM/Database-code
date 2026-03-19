Taxonomic DATABASE – Hyolitha 
GitHub Repository — README
A comprehensive, AI-assisted taxonomic database of the extinct marine invertebrate class Hyolitha, built from systematic extraction of 194 paleontological publications. The project combines a structured Microsoft Access database application with a reproducible data pipeline documented from raw literature to a publication-ready dataset of 1,455 taxon records.
This work was supported by the Ministry of Culture of the Czech Republic (DKRVO 2024–2028/2.I.cb, National Museum, 00023272).
What the Project Does
This project provides:
•	A structured taxonomic database covering all known Hyolitha taxa, with records spanning morphological descriptions, stratigraphic occurrences, geographic distributions, synonymies, and bibliographic references — extracted from primary literature and normalized into a relational schema.
•	A Microsoft Access GUI application (frmDark) for searching, editing, comparing, and exporting taxon records, built entirely from VBA form-builder packages that can be reproduced from source in any Access installation.
•	A reproducible data pipeline consisting of sixteen numbered steps — from Google Colab preprocessing through SQL view construction to the full VBA form stack — each documented with source code and installation instructions.
•	AI-assisted extraction methodology using large language models (DeepSeek-V3 and Claude) to extract structured data from unstructured paleontological literature.

The database schema consists of six normalized tables (Taxa, Taxa_Taxonomie, Taxa_Popis, Taxa_Material, Taxa_Vyskyt, Taxa_Poznamky) joined through a master query (vw_Complete_Taxa) that exposes all fields for search and export.

Why the Project Is Useful
Hyolitha are an extinct group of Cambrian–Permian marine invertebrates whose taxonomy has accumulated across more than 150 years of literature in multiple languages (English, French, German, Russian, Chinese). No comprehensive, machine-readable taxonomic database existed prior to this project.

This repository is useful for:
•	Paleontologists and systematists needing a searchable, cross-referenced record of Hyolitha taxonomy with morphological, stratigraphic, and geographic data in a single application.
•	Biodiversity informatics researchers interested in LLM-assisted structured data extraction from scientific literature as a reproducible methodology.
•	Database developers looking for a complete example of a VBA-based Access application built programmatically from source code, including dark-themed forms, tabbed detail views, batch export, and undo/history functionality.
•	Digital humanities and open science practitioners as a case study in converting legacy multilingual scientific literature into a FAIR-compliant structured dataset.
Getting Started
Requirements
•	Microsoft Access 2016 or later (.accdb format)
•	Microsoft Excel (for the preprocessing step)
•	Google account with access to Google Colab (for Step 01)
•	Windows (required for the WSH/VBScript installer in Step 02)

Build Steps
The project is assembled in eleven sequential steps. Each step has a corresponding source file in the repository.
Complete Build Sequence — All 16 Steps
01–05 · Data preparation and database creation (Python, VBScript, SQL)
06–11 · Core VBA packages — main form, compare, import/export, utilities, preview, detail editor
12 · BALIK_6_7 — filter presets, scope-aware export, bulk edit, frmMergePreview, frmValidator fix
13 · BALIK_8 — frmTextEditor for Compare, CROSS-SECTION / CARDINAL_PROCESSES fix in Compare
14 · BALIK_9 — BulkRename preview crash fix, hardcoded column index map for Compare fields
15 · BALIK_10 — IMPORT XLSX / BIBTEX / DIFF VIEWER buttons, Ctrl+B shortcut, report overflow fix
16 · BALIK_11 — NOTES field in TAXONOMY tab, left sidebar navigation, CARDINAL filter-clear fix


Step	File	Technology	Purpose
01	Colab_create_vw_Complete_Taxa.txt	Python / Google Colab	Prepare structured Excel with 6 thematic sheets
02	code_create_database.wsf	VBScript / WSH	Create Access DB and import all data from Excel
03	code_access_table.txt	SQL (Access)	Master query joining all 6 tables
04	create_column_compare.txt	SQL (Access)	Add Compare Yes/No flag column to Taxa
05	code_create_table_history.txt	SQL (Access)	Create audit/history table for undo support
06	code_main_form.txt	VBA (Access)	Package 1 — main search form frmDark + subform
07	code_compare_merge.txt	VBA (Access)	Package 2 — side-by-side taxon comparison form
08	code_export_import.txt	VBA (Access)	Package 3 — CSV/XLSX import, DOCX/BibTeX export
09	code_utilities.txt	VBA (Access)	Package 4 — Find/Replace, Validator, History, Stats
10	code_preview.txt	VBA (Access)	Package 5 — Quick Preview panel + printable report
11	code_detail.txt	VBA (Access)	Detail form frmDetailTaxa — full tabbed record editor
12	BALIK_6-7.bas	VBA (Access)	filter presets, scope-aware export, bulk edit, frmMergePreview, frmValidator fix
13	BALIK_8.bas	VBA (Access)	frmTextEditor for Compare, CROSS-SECTION / CARDINAL_PROCESSES fix in Compare
14	BALIK_9.bas	VBA (Access)	BulkRename preview crash fix, hardcoded column index map for Compare fields
15	BALIK_10.bas	VBA (Access)	IMPORT XLSX / BIBTEX / DIFF VIEWER buttons, Ctrl+B shortcut, report overflow fix
16	BALIK_11.bas	VBA (Access)	NOTES field in TAXONOMY tab, left sidebar navigation, CARDINAL filter-clear fix

Getting Help
•	Build Guide: A detailed Word document (Hyolitha_Database_Build_Guide.docx) is included in this repository. It covers every step with annotated code, installation instructions, known issues, troubleshooting, and a complete database schema reference.
•	Issues: If you encounter a bug or have a question, please open a GitHub Issue with a description of the step you are on, the error message, and your Access version.

Who Maintains This Project
This project is developed and maintained by Martin Valent at the National Museum (Národní muzeum), Department of Palaeontology, as part of a project supported by the Ministry of Culture of the Czech Republic (DKRVO 2024–2028/2.I.c, National Museum, 00023272).

The methodology and dataset are described in a manuscript submitted to Biodiversity Data Journal (BDJ). A preprint and the full dataset will be deposited on Zenodo upon acceptance.

Contributions, corrections, and additions from the paleontological community are welcome. Please open an Issue or submit a Pull Request with a description of the proposed change and its source reference.

License
The source code in this repository (VBA packages, Python scripts, SQL queries) is released under the MIT License. The taxonomic data are made available under CC BY 4.0. See LICENSE.txt for full terms.
