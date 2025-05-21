from flask import Flask, render_template, url_for, flash, redirect, request
import database_operations as db_ops
import os
import json # For simulating API call payload
import requests # For simulating API call
import logging # For application logging

app = Flask(__name__)

# --- Logging Configuration ---
# In a production environment, you might want to configure logging more extensively,
# e.g., to a file, with rotation, different levels for different modules.
# For this exercise, basic console logging for INFO and above will be set.
if not app.debug: # Only apply this when not in debug mode, or adjust as needed
    app.logger.setLevel(logging.INFO)
    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    # You can add a formatter for more detailed logs
    # formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    # stream_handler.setFormatter(formatter)
    # app.logger.addHandler(stream_handler)
else: # In debug mode, Flask's default logger is usually sufficient and more verbose.
    app.logger.setLevel(logging.DEBUG)

app.logger.info("Flask application initialized with basic logging.")


# Configure database connection
# WARNING: Storing credentials directly in code is not recommended for production.
# Consider using environment variables or a configuration file.
DB_CONFIG = {
    'host': '124.223.68.89',
    'user': 'root',
    'password': 'Mjhu666777;',
    'database': 'ShenJiao'
}

# Pass the DB_CONFIG to the database_operations module
db_ops.DB_CONFIG = DB_CONFIG

# Ensure the 'uploads' and 'processed' directories exist
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')
PROCESSED_FOLDER = os.path.join(os.getcwd(), 'processed')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER


@app.route('/')
def index():
    """
    Main page, displays a list of uploaded files and their status.
    """
    try:
        file_records = db_ops.get_all_file_records_for_display()
        if file_records is None: # Check if db_ops call failed
            flash("Error fetching file records from the database. Please try again later.", "danger")
            file_records = [] # Ensure template gets an iterable

        # Convert filesize to a more readable format (e.g., KB, MB)
        for record in file_records: # This loop is safe even if file_records is empty
            if record.get('filesize') is not None: # Check if filesize is not None
                record['filesize_readable'] = convert_filesize(record['filesize'])
            else:
                record['filesize_readable'] = "N/A" # Or some other placeholder
        return render_template('index.html', file_records=file_records)
    except Exception as e:
        app.logger.error(f"Error in route {request.path}: {e}")
        flash("An unexpected error occurred while loading the main page. Please try again.", "danger")
        return render_template('index.html', file_records=[]) # Provide empty list to template

def convert_filesize(size_bytes):
    """Converts filesize in bytes to a more readable format (KB, MB, GB)."""
    if size_bytes == 0:
        return "0B"
    import math
    size_name = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return f"{s} {size_name[i]}"

# Placeholder routes for actions - to be implemented later

# Configuration for the conceptual Windows Worker API
WINDOWS_WORKER_IP = os.environ.get("WINDOWS_WORKER_IP", "192.168.1.100") # Example, replace with actual or config
WINDOWS_WORKER_PORT = os.environ.get("WINDOWS_WORKER_PORT", "8080") # Example port
WORKER_API_ENDPOINT_EXTRACT = f"http://{WINDOWS_WORKER_IP}:{WINDOWS_WORKER_PORT}/run_extraction"
WORKER_API_ENDPOINT_GENERATE = f"http://{WINDOWS_WORKER_IP}:{WINDOWS_WORKER_PORT}/run_generation"
WORKER_API_ENDPOINT_FETCH_FILE = f"http://{WINDOWS_WORKER_IP}:{WINDOWS_WORKER_PORT}/fetch_generated_file" # New endpoint for fetching

# This is the base directory *on the worker machine* where images will be stored.
# The worker will create subdirectories like file_record_id/ under this base.
WORKER_IMAGE_OUTPUT_DIR_BASE = os.environ.get("WORKER_IMAGE_OUTPUT_DIR_BASE", "C:\\proofease_worker_data\\images")
# This is the base directory *on the worker machine* where generated proofreading list documents will be stored.
WORKER_GENERATED_LISTS_DIR_BASE = os.environ.get("WORKER_GENERATED_LISTS_DIR_BASE", "C:\\proofease_worker_generated_lists")


@app.route('/export_proofread_list/start_extraction/<file_record_id>')
def start_extraction_for_export(file_record_id):
    """
    Initiates the content extraction process on the Windows worker
    for a given file_record_id.
    """
    try:
        record = db_ops.get_file_record(file_record_id)

        if record is None: # Check for DB error from get_file_record
             flash(f"Database error while fetching record {file_record_id}. Please try again.", "danger")
             return redirect(url_for('index'))
        if not record: # Check for record not found (empty dict/list or specific False from db_ops)
            flash(f"Error: File record {file_record_id} not found.", "danger")
            return redirect(url_for('index'))

    # Assuming 'filepath' stores the relative path like 'uploads/username/filename.docx'
    # And the Flask app serves these files from a specific base URL.
    # For this simulation, we'll use the provided server IP and a conceptual 'files' endpoint.
    # In a real setup, 'filepath' might be a direct URL or need transformation.
    
    # Constructing the document_url based on the assumption that 'filepath'
    # is the path relative to a known base URL where files are served.
    # Example: record['filepath'] = "uploads/user1/my_document.docx"
    # Server base URL for files: "http://124.223.68.89:7777/" (as per prompt example)
    # document_url will be "http://124.223.68.89:7777/uploads/user1/my_document.docx"
    
    # Using current request's host for base URL if files are served by this app.
    # If files are on a different server, this needs to be adjusted.
    # For now, let's assume files are served from the same IP as this app, under a '/files/' route (conceptual)
    # file_server_base_url = request.url_root.replace(request.script_root, '') # http://<host>:<port>/
    # For the given example "http://124.223.68.89:7777/", we'll use that directly.
    # This implies the 'filepath' column in 'file_records' stores the path *after* this base.
    
    file_server_base_url = "http://124.223.68.89:7777/" # As per example in prompt for worker
    
    if not record.get('filepath'):
        flash(f"Error: Filepath not available for record {file_record_id}. Cannot initiate extraction.", "danger")
        if not db_ops.update_file_record_status(file_record_id, "ExtractionFailed", error_message="Filepath missing in DB record"):
            flash("Additionally, failed to update status in database due to another error.", "danger")
        return redirect(url_for('index'))

    document_url = file_server_base_url + record['filepath'].lstrip('/') # Ensure no double slashes if filepath starts with /
    original_filename = record.get('original_filename', 'Unknown Filename')

    app.logger.info(f"Initiating extraction for {original_filename} (ID: {file_record_id}). Document URL: {document_url}")

    # Update status to 'ExtractionPending'
    update_success = db_ops.update_file_record_status(file_record_id, "ExtractionPending")
    if not update_success:
        flash(f"Database error: Failed to update status to ExtractionPending for '{original_filename}'. Extraction not initiated.", "danger")
        return redirect(url_for('index'))

    # Simulate making an HTTP POST request to the Windows worker API
    worker_payload = {
        "file_record_id": file_record_id,
        "document_url": document_url,
        "image_output_dir_base": WORKER_IMAGE_OUTPUT_DIR_BASE,
        # The DB_CONFIG for the worker should be securely managed and known by the worker,
        # not directly passed in this request for security reasons.
        # The worker script itself will have its DB_CONFIG.
    }

    print(f"[INFO] Simulating call to Windows Worker API: {WORKER_API_ENDPOINT_EXTRACT}")
    print(f"[INFO] Worker Payload: {json.dumps(worker_payload)}")

    try:
        # This is a simulation. In a real scenario, you'd make an actual HTTP request.
        # response = requests.post(WORKER_API_ENDPOINT_EXTRACT, json=worker_payload, timeout=10) # 10s timeout
        # response.raise_for_status() # Raise an exception for HTTP errors
        print(f"[SIMULATION] Successfully sent extraction request to worker for {original_filename}.")
        flash(f"Extraction process initiated for '{original_filename}'. The status will update automatically. Please refresh the page after some time.", "info")
    except requests.exceptions.RequestException as e:
        app.logger.error(f"Simulated call to Windows Worker API for extraction failed for {original_filename}: {e}")
        flash(f"Error initiating extraction for '{original_filename}': Could not connect to worker. Please contact admin.", "danger")
        if not db_ops.update_file_record_status(file_record_id, "ExtractionFailed", error_message=f"Worker API call failed: {str(e)}"):
            flash("Additionally, failed to update status to ExtractionFailed in database.", "danger")
    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred while starting extraction. Please try again.", "danger")
    
    return redirect(url_for('index'))


@app.route('/trigger_ai_proofread/<file_record_id>')
def trigger_ai_proofread(file_record_id):
    """
    Simulates triggering and completing the AI proofreading process.
    """
    try:
        record = db_ops.get_file_record(file_record_id)
        original_filename = record.get('original_filename', f"ID: {file_record_id}") if record else f"ID: {file_record_id}"

        if record is None: # Check for DB error from get_file_record
             flash(f"Database error while fetching record {file_record_id} for AI Proofreading. Please try again.", "danger")
             return redirect(url_for('index'))
        if not record:
            flash(f"Error: File record {file_record_id} not found. Cannot start AI Proofreading.", "danger")
            return redirect(url_for('index'))

    # Validate status - e.g., 'uploaded', 'AIReviewFailed', or 'Pending' (if re-processing from a general pending state)
    # For simplicity, let's allow from 'uploaded' or 'AIReviewFailed'
    allowed_statuses_for_ai_review = ['uploaded', 'AIReviewFailed', 'Pending', 'ExtractionSuccess']
    if record.get('status') not in allowed_statuses_for_ai_review:
        flash(f"AI Proofreading for '{original_filename}' can only be started if status is one of {allowed_statuses_for_ai_review}. Current status: '{record.get('status')}'.", "warning")
        return redirect(url_for('index'))

    app.logger.info(f"Simulating AI Proofreading for {original_filename} (ID: {file_record_id}).")

    # 1. Update status to 'AIReviewPending'
    update_pending_success = db_ops.update_file_record_status(file_record_id, "AIReviewPending")
    if not update_pending_success:
        flash(f"Database error: Failed to update status to AIReviewPending for '{original_filename}'. AI Proofreading not fully initiated.", "danger")
        return redirect(url_for('index'))
    
    # 2. Simulate AI processing delay (optional for UX, not strictly needed for simulation)
    # import time
    # time.sleep(1) # Simulate work

    # 3. Simulate completion - update to 'AISuccess'
    update_success_success = db_ops.update_file_record_status(file_record_id, "AISuccess", error_message=None)
    if update_success_success:
        flash(f"AI Proofreading for '{original_filename}' simulated and marked as successful.", "success")
    else:
        flash(f"Database error: Failed to update status to AISuccess for '{original_filename}' after simulation.", "danger")
        # Status might remain 'AIReviewPending'
    
    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred while triggering AI proofreading. Please try again.", "danger")

    return redirect(url_for('index'))


@app.route('/trigger_manual_review/<file_record_id>')
def trigger_manual_review(file_record_id):
    """
    Simulates triggering and completing the Manual Review process.
    After this, status becomes 'Pending' (indicating ready for export list generation by user).
    """
    try:
        record = db_ops.get_file_record(file_record_id)
        original_filename = record.get('original_filename', f"ID: {file_record_id}") if record else f"ID: {file_record_id}"

        if record is None: # Check for DB error
            flash(f"Database error while fetching record {file_record_id} for Manual Review. Please try again.", "danger")
            return redirect(url_for('index'))
        if not record:
            flash(f"Error: File record {file_record_id} not found. Cannot start Manual Review.", "danger")
            return redirect(url_for('index'))

    # Validate status - typically after AI success or if manual review failed previously
    allowed_statuses_for_manual_review = ['AISuccess', 'ManualReviewFailed']
    if record.get('status') not in allowed_statuses_for_manual_review:
        flash(f"Manual Review for '{original_filename}' can only be started if status is one of {allowed_statuses_for_manual_review}. Current status: '{record.get('status')}'.", "warning")
        return redirect(url_for('index'))

    app.logger.info(f"Simulating Manual Review for {original_filename} (ID: {file_record_id}).")

    # 1. Update status to 'ManualReviewPending'
    update_pending_success = db_ops.update_file_record_status(file_record_id, "ManualReviewPending")
    if not update_pending_success:
        flash(f"Database error: Failed to update status to ManualReviewPending for '{original_filename}'. Manual Review not fully initiated.", "danger")
        return redirect(url_for('index'))

    # 2. Simulate Manual Review completion - update to 'Pending'
    update_complete_success = db_ops.update_file_record_status(file_record_id, "Pending", error_message=None)
    if update_complete_success:
        flash(f"Manual Review for '{original_filename}' simulated. Document is now in 'Pending' state, ready for next actions (e.g., Export Proofreading List).", "success")
    else:
        flash(f"Database error: Failed to update status to Pending for '{original_filename}' after Manual Review simulation.", "danger")
        # Status might remain 'ManualReviewPending'

    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred while triggering manual review. Please try again.", "danger")
        
    return redirect(url_for('index'))


@app.route('/export_proofread_list/start_generation/<file_record_id>')
def start_generation_for_export(file_record_id):
    """
    Initiates the advice document generation process on the Windows worker
    for a given file_record_id.
    """
    try:
        record = db_ops.get_file_record(file_record_id)
        original_filename = record.get('original_filename', f"ID: {file_record_id}") if record else f"ID: {file_record_id}"

        if record is None: # DB error check
            flash(f"Database error while fetching record {file_record_id} for generation. Please try again.", "danger")
            return redirect(url_for('index'))
        if not record:
            flash(f"Error: File record {file_record_id} not found. Cannot start generation.", "danger")
            return redirect(url_for('index'))

    # Check current status - should be 'ExtractionSuccess' or allow retry from 'GenerationFailed'
    if record.get('status') not in ['ExtractionSuccess', 'GenerationFailed']:
        flash(f"Error: Advice document generation can only be started if content extraction was successful or previous generation failed. Current status for '{original_filename}' is '{record.get('status')}'.", "warning")
        return redirect(url_for('index'))

    app.logger.info(f"Initiating advice document generation for {original_filename} (ID: {file_record_id}).")

    # Update status to 'GenerationPending'
    update_pending_success = db_ops.update_file_record_status(file_record_id, "GenerationPending")
    if not update_pending_success:
        flash(f"Database error: Failed to update status to GenerationPending for '{original_filename}'. Generation not initiated.", "danger")
        return redirect(url_for('index'))

    # Simulate making an HTTP POST request to the Windows worker API for generation
    worker_payload = {
        "file_record_id": file_record_id,
        "output_base_dir": WORKER_GENERATED_LISTS_DIR_BASE,
        # DB_CONFIG is known by the worker, not passed in request for security.
    }

    print(f"[INFO] Simulating call to Windows Worker API for generation: {WORKER_API_ENDPOINT_GENERATE}")
    print(f"[INFO] Worker Payload for generation: {json.dumps(worker_payload)}")

    try:
        # This is a simulation. In a real scenario, you'd make an actual HTTP request.
        # response = requests.post(WORKER_API_ENDPOINT_GENERATE, json=worker_payload, timeout=10)
        # response.raise_for_status()
        print(f"[SIMULATION] Successfully sent generation request to worker for {original_filename}.")
        flash(f"Advice document generation initiated for '{original_filename}'. The status will update automatically. Please refresh the page after some time.", "info")
    except requests.exceptions.RequestException as e:
        app.logger.error(f"Simulated call to Windows Worker API for generation failed for {original_filename}: {e}")
        flash(f"Error initiating generation for '{original_filename}': Could not connect to worker. Please contact admin.", "danger")
        if not db_ops.update_file_record_status(file_record_id, "GenerationFailed", error_message=f"Worker API call for generation failed: {str(e)}"):
            flash("Additionally, failed to update status to GenerationFailed in database.", "danger")
    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred while starting generation. Please try again.", "danger")
    
    return redirect(url_for('index'))


@app.route('/download_proofreading_list/<file_record_id>')
def download_proofreading_list(file_record_id):
    """
    Simulates initiating a download of the generated proofreading list from the worker.
    """
    try:
        record = db_ops.get_file_record(file_record_id)
        original_filename = record.get('original_filename', f"ID: {file_record_id}") if record else f"ID: {file_record_id}"

        if record is None: # DB error check
            flash(f"Database error while fetching record {file_record_id} for download. Please try again.", "danger")
            return redirect(url_for('index'))
        if not record:
            flash(f"Error: File record {file_record_id} not found. Cannot download.", "danger")
            return redirect(url_for('index'))

    # Validate status and filepath
    # 'GenerationSuccess' is the new status indicating the file is ready.
    # 'exportListAlready' was from previous design, can be kept for compatibility or removed if not used.
    if record.get('status') not in ['GenerationSuccess', 'exportListAlready']:
        flash(f"Error: Proofreading list for '{original_filename}' is not ready for download. Current status: {record.get('status')}'.", "warning")
        return redirect(url_for('index'))

    proof_list_filepath_on_worker = record.get('proof_list_filepath')
    if not proof_list_filepath_on_worker:
        flash(f"Error: Filepath for the proofreading list of '{original_filename}' is missing. Cannot download.", "danger")
        if not db_ops.update_file_record_status(file_record_id, "GenerationFailed", error_message="Proof list filepath missing after GenerationSuccess status."):
            flash("Additionally, failed to update status to GenerationFailed in database.", "danger")
        return redirect(url_for('index'))

    # Construct a user-friendly download filename
    base_name, ext = os.path.splitext(original_filename)
    download_filename = f"{base_name}_prooflist.docx" # Assuming it's always docx

    # Simulate making an HTTP GET request to the Windows worker API to fetch the file
    # The worker would use proof_list_filepath_on_worker to locate and stream the file.
    print(f"[INFO] Simulating call to Worker API to fetch file: {WORKER_API_ENDPOINT_FETCH_FILE}")
    print(f"  Params: filepath={proof_list_filepath_on_worker}")
    print(f"  User would download as: {download_filename}")

    # In a real scenario, this would involve:
    # 1. Making the request: `response = requests.get(WORKER_API_ENDPOINT_FETCH_FILE, params={'filepath': proof_list_filepath_on_worker}, stream=True)`
    # 2. Checking `response.status_code`
    # 3. Returning a `Response` object with the file stream and appropriate headers:
    #    `return Response(response.iter_content(chunk_size=8192), headers={'Content-Disposition': f'attachment; filename="{download_filename}"', 'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'})`
    
    flash(f"Simulated Download: The file '{download_filename}' for '{original_filename}' would be streamed from the worker. (Actual download not implemented in this step).", "info")
    # For now, just redirect as actual download is not part of this step.
    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred while preparing the download. Please try again.", "danger")

    return redirect(url_for('index'))


@app.route('/upload', methods=['POST'])
def upload_file_action():
    # Logic for handling file uploads - Placeholder
    try:
        flash("File upload functionality is not yet implemented.", "warning")
    except Exception as e:
        app.logger.error(f"Error in route {request.path}: {e}")
        flash("An unexpected error occurred. Please try again.", "danger")
    return redirect(url_for('index')) # Or to an upload page

@app.route('/process/<file_record_id>') # Changed to string to match others, cast if needed
def process_file_action(file_record_id):
    # Logic for initiating file processing - Placeholder
    try:
        flash(f"Processing for file ID {file_record_id} is not yet implemented.", "warning")
    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred. Please try again.", "danger")
    return redirect(url_for('index'))


@app.route('/download_original/<file_record_id>')
def download_original_action(file_record_id):
    # Logic for downloading original file - Placeholder
    try:
        flash(f"Download of original file ID {file_record_id} is not yet implemented.", "warning")
        # Here you would:
        # 1. Get file record from DB.
        # 2. Construct path to original file (e.g., in UPLOAD_FOLDER).
        # 3. Use send_from_directory to send the file.
    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred. Please try again.", "danger")
    return redirect(url_for('index'))

@app.route('/download_proofread/<file_record_id>') # This was the old placeholder
def download_proofread_action(file_record_id): # Will now be an alias or removed
    try:
        # This route can now be an alias to the new download_proofreading_list
        # or deprecated if the new URL is used directly in templates.
        # For now, let's make it redirect to the new one to ensure old links (if any) still work.
        app.logger.info(f"Redirecting from old /download_proofread/{file_record_id} to new endpoint.")
        # flash("Redirecting to the new download endpoint...", "debug") # Optional: for debugging
    except Exception as e: # Should not happen in a simple redirect, but for consistency
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred during redirection. Please try again.", "danger")
        return redirect(url_for('index')) # Fallback redirect
    return redirect(url_for('download_proofreading_list', file_record_id=file_record_id))

@app.route('/view_report/<file_record_id>')
def view_report_action(file_record_id):
    # Logic for viewing processing report (if any) - Placeholder
    try:
        flash(f"Viewing report for file ID {file_record_id} is not yet implemented.", "warning")
    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred. Please try again.", "danger")
    return redirect(url_for('index'))

@app.route('/delete/<file_record_id>', methods=['POST']) # Should be POST for destructive actions
def delete_file_action(file_record_id):
    # Logic for deleting a file record and associated files - Placeholder
    try:
        # IMPORTANT: Add CSRF protection for POST requests like this in a real app.
        flash(f"Deletion of file ID {file_record_id} is not yet implemented.", "warning")
        # Here you would:
        # 1. Get file record.
        # 2. Delete associated files from disk (original, proofread list, images).
        # 3. Delete DB records (file_records, tmp_document_contents, etc.).
    except Exception as e:
        app.logger.error(f"Error in route {request.path} for ID {file_record_id}: {e}")
        flash("An unexpected error occurred. Please try again.", "danger")
    return redirect(url_for('index'))

if __name__ == '__main__':
    # Create dummy data for testing UI if DB is empty or for development
    # This part will not run when deployed with a WSGI server like Gunicorn
    if not db_ops.get_all_file_records_for_display():
        print("Database is empty or connection failed. Consider adding dummy data for UI testing if needed.")
        # Example: db_ops.create_file_record("test_document.docx", 102400, status="uploaded")
        # Example: db_ops.create_file_record("another_doc.docx", 204800, status="processed", proof_list_filepath="processed/another_doc_proofread.docx")

    
    # Secret key for flashing messages
    app.secret_key = os.environ.get("FLASK_SECRET_KEY", "super_secret_dev_key") # Change for production!

    app.run(debug=True, host='0.0.0.0', port=5000)
