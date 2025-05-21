# ProofEase Web Application

## Overview

ProofEase Web Application is a Flask-based web interface designed to manage and orchestrate a document proofreading workflow. It allows users to track Word documents, initiate processing steps (simulated AI proofreading, simulated manual review, and actual content extraction for proofreading list generation), and manage the overall status of each document. The system is designed to interact with a conceptual Windows Worker Service that handles the heavy lifting of Microsoft Word document processing.

## Architecture

The system comprises the following key components:

1.  **Flask Web Application (`app.py`)**:
    *   Provides the user interface for managing documents and workflows.
    *   Handles user requests and orchestrates calls to the Windows Worker Service.
    *   Interacts with the MySQL database to store and retrieve metadata.

2.  **MySQL Database**:
    *   Stores metadata about documents, including filenames, upload times, processing status, error messages, and paths to generated files.
    *   Key tables: `file_records`, `tmp_document_contents`, `document_contents`, `document_content_chunks`.

3.  **Conceptual Windows Worker Service**:
    *   A separate (conceptual) service running on a Windows machine with Microsoft Word installed.
    *   Exposes an API that the Flask app calls to perform tasks like:
        *   Extracting content from Word documents (`word_extractor_worker.py`).
        *   Generating proofreading advice lists (`advice_generator_worker.py`).
    *   Interacts directly with the MySQL database via `database_operations_worker.py` to store extracted content and update status.
    *   Requires access to a file server to download original documents and a shared location to store generated files (images, proofreading lists).

4.  **File Storage / Server**:
    *   **Original Documents**: A server (e.g., `http://124.223.68.89:7777/`) where uploaded Word documents are stored and accessible via URL by the Windows Worker Service for download.
    *   **Extracted Images**: Stored on the Windows Worker machine in a configured base directory (e.g., `C:\proofease_worker_data\images`), organized by `file_record_id`.
    *   **Generated Proofreading Lists**: Stored on the Windows Worker machine in a configured base directory (e.g., `C:\proofease_worker_generated_lists`), organized by `file_record_id`. The Flask app receives the path to these files.

## Features

*   **Document Listing**: Displays a list of documents from the `file_records` table with their current status, size, and upload time.
*   **Status-Based Actions**: Provides actions based on the document's current status (e.g., Start AI Proofreading, Start Manual Review, Start Extraction, Start Generation, Download).
*   **Simulated AI Proofreading**: Placeholder action to simulate sending a document for AI review and updating its status.
*   **Simulated Manual Review**: Placeholder action to simulate sending a document for manual review and updating its status.
*   **Word Document Processing Orchestration**:
    *   **Content Extraction**: Initiates content extraction on the Windows Worker. The worker downloads the specified Word document, parses its content (text, tables, images), and stores the structured data in the `tmp_document_contents` table.
    *   **Proofreading List Generation**: Initiates the generation of a "Proofreading Advice List" Word document on the Windows Worker. The worker fetches extracted content and AI suggestions (conceptually from `document_content_chunks`) to produce this report.
*   **Download Simulation**: Simulates downloading the generated proofreading list by providing a link that (conceptually) would stream the file from the worker.

## Core Code Components

*   **`app.py`**: The main Flask application file. Defines routes, handles web requests, interacts with `database_operations.py`, and orchestrates calls (simulated) to the Windows Worker Service.
*   **`database_operations.py`**: Module used by the Flask app to interact with the MySQL database (e.g., fetching file records, updating statuses).
*   **`templates/index.html`**: The main HTML template for the web interface, using Jinja2 for dynamic content rendering.
*   **`word_extractor_worker.py`**: Script intended for the Windows Worker. Downloads a Word document, extracts its content using `win32com`, and saves the extracted data to the `tmp_document_contents` table via `database_operations_worker.py`.
*   **`advice_generator_worker.py`**: Script intended for the Windows Worker. Fetches data from `tmp_document_contents`, `document_contents`, and `document_content_chunks` (AI suggestions) to generate a "Proofreading Advice List" Word document using `win32com`. Saves the generated document to a shared location.
*   **`database_operations_worker.py`**: Module used by the worker scripts (`word_extractor_worker.py`, `advice_generator_worker.py`) to interact with the MySQL database.

## Setup and Running

1.  **Database Setup**:
    *   Ensure a MySQL server is running and accessible.
    *   Create the `ShenJiao` database (or as configured).
    *   Create tables: `file_records`, `tmp_document_contents`, `document_contents`, `document_content_chunks` with appropriate schemas.
    *   Update `DB_CONFIG` in `app.py` (and ensure worker scripts can be configured similarly) with your MySQL credentials.

2.  **Flask Application Setup**:
    *   Install Python dependencies: `pip install -r requirements.txt`
    *   Set environment variables if used (e.g., `FLASK_SECRET_KEY`, worker IP/paths).
    *   Run the Flask app: `python app.py`. The app will typically run on `http://0.0.0.0:5000/`.

3.  **Conceptual Windows Worker Service Setup**:
    *   **This part is conceptual and not implemented as a service in this codebase.**
    *   A Windows machine with Microsoft Word and Python is required.
    *   Install Python dependencies for the worker: `pip install -r requirements.txt` (or a worker-specific subset).
    *   The worker scripts (`word_extractor_worker.py`, `advice_generator_worker.py`) would need to be wrapped in an API (e.g., using Flask, FastAPI, or another web framework) that listens for HTTP requests from the main Flask app.
    *   The worker service needs its own database configuration to connect to MySQL.
    *   Ensure the worker has network access to the file server for downloading original documents and to the shared directories for storing extracted images and generated lists.

4.  **File Server**:
    *   A separate file server (e.g., a simple HTTP server, Nginx, Apache) is needed to host the original Word documents so they can be downloaded by the worker via URL. The example URL used in the app is `http://124.223.68.89:7777/`. The `filepath` column in the `file_records` table should store paths relative to this base URL.

## Workflow Overview

1.  **File Record Creation**: (Manual/External) A record for a Word document is created in the `file_records` table, including its `filepath` (URL accessible path) and initial status (e.g., 'uploaded').
2.  **User Interaction**: The user views the document list in the Flask web app.
3.  **AI Proofreading (Simulated)**: User clicks "AI审校". App updates status to 'AIReviewPending', then 'AISuccess'.
4.  **Manual Review (Simulated)**: User clicks "人工审核". App updates status to 'ManualReviewPending', then 'Pending'.
5.  **Content Extraction**:
    *   User clicks "导出审校清单 (Start Extraction)".
    *   Flask app updates status to `ExtractionPending`.
    *   Flask app (conceptually) calls the `/run_extraction` API endpoint on the Windows Worker Service, providing `file_record_id`, `document_url` (to the original Word file), and `image_output_dir_base`.
    *   The worker's `word_extractor_worker.py` downloads the document, parses it, and saves content to `tmp_document_contents`.
    *   The worker updates the `file_records` status to `ExtractionSuccess` or `ExtractionFailed`.
6.  **Proofreading List Generation**:
    *   User clicks "导出审校清单 (Start Generation)" (available if extraction was successful).
    *   Flask app updates status to `GenerationPending`.
    *   Flask app (conceptually) calls the `/run_generation` API endpoint on the Windows Worker Service, providing `file_record_id` and `output_base_dir` for the generated list.
    *   The worker's `advice_generator_worker.py` fetches data from DB tables, generates the Word document, and saves it.
    *   The worker updates the `file_records` status to `GenerationSuccess` (and stores `proof_list_filepath`) or `GenerationFailed`.
7.  **Download (Simulated)**:
    *   User clicks "Download Proofread List".
    *   Flask app (conceptually) calls a `/fetch_generated_file` API on the worker, providing the `proof_list_filepath`.
    *   The worker would stream the file back. (Currently, this is simulated with a flash message).

## Dependencies

All Python dependencies are listed in `requirements.txt`. Key dependencies include:

*   Flask
*   mysql-connector-python
*   requests
*   pywin32 (for Windows worker scripts)
*   Pillow (for image processing on worker)
*   openpyxl, pandas (potentially used by worker scripts, or for data handling if extended)

Install using: `pip install -r requirements.txt`

## Future Work / Notes

*   **Implement Windows Worker Service API**: The current interaction with the worker is simulated. A proper HTTP API (e.g., using Flask or FastAPI) needs to be built on the Windows worker to receive requests from the main Flask app and execute `word_extractor_worker.py` and `advice_generator_worker.py`.
*   **Actual AI and Manual Review**: The AI and Manual review steps are currently placeholders that only change status. Real implementation would involve integrating with AI services and providing interfaces for manual review.
*   **Asynchronous Operations**: For long-running tasks like document processing, implement asynchronous task queues (e.g., Celery, RQ) to prevent blocking web requests and provide better user experience.
*   **Robust File Handling**: Implement actual file upload functionality in the Flask app and secure file serving.
*   **Actual File Download**: The download of the generated proofreading list is simulated. Implement actual file streaming from the worker or a shared accessible location.
*   **Security**: Enhance security (CSRF protection, input validation, secure credential management, authentication/authorization).
*   **Configuration Management**: Move hardcoded configurations (paths, URLs) to environment variables or configuration files.
*   **Error Handling and Logging**: Further improve detailed error tracking and user feedback.
*   **Database Schema**: Review and refine database schema for optimal performance and data integrity.
*   **User Interface**: Enhance the UI for better usability and richer display of information.
*   **Testing**: Add comprehensive unit and integration tests.