import mysql.connector

# DB_CONFIG will be set by the worker script (e.g., word_extractor_worker.py) at runtime.
DB_CONFIG = None

def get_db_connection_worker():
    """Establishes a connection to the MySQL database using worker's DB_CONFIG."""
    if DB_CONFIG is None:
        print("[ERROR] DB_CONFIG is not set for the worker. Cannot connect to database.")
        raise ValueError("DB_CONFIG is not set. It should be configured by the calling worker script.")
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        # print("[INFO] Database connection established successfully by worker.")
        return conn
    except mysql.connector.Error as err:
        print(f"[ERROR] Worker failed to connect to database: {err}")
        # In a real app, you might want to log this error or raise it
        return None

def add_tmp_document_content_batch(file_record_id: str, elements: list[dict]):
    """
    Deletes existing entries for the file_record_id in tmp_document_contents
    and then inserts a batch of new elements.
    """
    if not elements:
        print(f"[INFO] No elements provided to add for file_record_id: {file_record_id}. Skipping DB operation.")
        return

    conn = get_db_connection_worker()
    if conn is None:
        print(f"[ERROR] Cannot add document content for {file_record_id}: No database connection.")
        raise ConnectionError("Failed to connect to the database for batch insert.")

    cursor = conn.cursor()
    
    try:
        # 1. Delete existing entries for this file_record_id
        delete_query = "DELETE FROM tmp_document_contents WHERE file_record_id = %s"
        cursor.execute(delete_query, (file_record_id,))
        print(f"[INFO] Deleted {cursor.rowcount} existing tmp_document_contents rows for file_record_id: {file_record_id}")

        # 2. Insert new elements in a batch
        insert_query = """
        INSERT INTO tmp_document_contents 
        (file_record_id, element_type, content_id, text_content, `level`, pageNo) 
        VALUES (%s, %s, %s, %s, %s, %s)
        """
        
        data_to_insert = []
        for elem in elements:
            data_to_insert.append((
                file_record_id, # This comes from the function argument
                elem.get("element_type"),
                elem.get("content_id"),
                elem.get("text_content"),
                elem.get("level"), # Can be None
                elem.get("pageNo") # Can be None
            ))

        cursor.executemany(insert_query, data_to_insert)
        conn.commit()
        print(f"[INFO] Successfully inserted {cursor.rowcount} new rows into tmp_document_contents for file_record_id: {file_record_id}")

    except mysql.connector.Error as err:
        print(f"[ERROR] Database error during batch operation for {file_record_id}: {err}")
        conn.rollback() # Rollback changes on error
        raise # Re-raise the exception to be caught by the worker
    except Exception as e_gen:
        print(f"[ERROR] A general error occurred during batch DB operation for {file_record_id}: {e_gen}")
        conn.rollback()
        raise
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
            # print("[INFO] Worker database connection closed.")

def _fetch_contents(table_name: str, file_record_id: str, element_type: str = None) -> list[dict]:
    """Helper function to fetch contents from a specified table."""
    conn = get_db_connection_worker()
    if conn is None:
        print(f"[ERROR] Cannot fetch from {table_name} for {file_record_id}: No database connection.")
        return []

    cursor = conn.cursor(dictionary=True)
    records = []
    try:
        if element_type:
            query = f"SELECT * FROM {table_name} WHERE file_record_id = %s AND element_type = %s ORDER BY id ASC;"
            cursor.execute(query, (file_record_id, element_type))
        else:
            query = f"SELECT * FROM {table_name} WHERE file_record_id = %s ORDER BY id ASC;"
            cursor.execute(query, (file_record_id,))
        records = cursor.fetchall()
        print(f"[INFO] Fetched {len(records)} rows from {table_name} for file_record_id: {file_record_id}" + (f" and element_type: {element_type}" if element_type else ""))
    except mysql.connector.Error as err:
        print(f"[ERROR] Database error fetching from {table_name} for {file_record_id}: {err}")
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
    return records

def get_tmp_document_contents_by_file_id(file_record_id: str, element_type: str = None) -> list[dict]:
    """Fetches records from 'tmp_document_contents' for a given file_record_id, optionally filtered by element_type."""
    return _fetch_contents("tmp_document_contents", file_record_id, element_type)

def get_document_contents_by_file_id(file_record_id: str, element_type: str = None) -> list[dict]:
    """Fetches records from 'document_contents' for a given file_record_id, optionally filtered by element_type."""
    return _fetch_contents("document_contents", file_record_id, element_type)

def get_document_content_chunks_by_file_id(file_record_id: str) -> list[dict]:
    """Fetches all records from 'document_content_chunks' for a given file_record_id."""
    # Assuming element_type is not typically used for chunks, so it's not a parameter here.
    # If filtering by type for chunks is needed, this function can be updated or use _fetch_contents.
    conn = get_db_connection_worker()
    if conn is None:
        print(f"[ERROR] Cannot fetch chunks for {file_record_id}: No database connection.")
        return []

    cursor = conn.cursor(dictionary=True)
    query = "SELECT * FROM document_content_chunks WHERE file_record_id = %s ORDER BY id ASC;"
    records = []
    try:
        cursor.execute(query, (file_record_id,))
        records = cursor.fetchall()
        print(f"[INFO] Fetched {len(records)} rows from document_content_chunks for file_record_id: {file_record_id}")
    except mysql.connector.Error as err:
        print(f"[ERROR] Database error fetching chunks for {file_record_id}: {err}")
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
    return records


if __name__ == '__main__':
    # This block is for testing this module directly.
    # Requires DB_CONFIG to be set manually for the test.
    print("Running database_operations_worker.py directly (for testing purposes)...")

    # --- Configuration for Direct Test ---
    # WARNING: Hardcoding credentials is not secure for production.
    TEST_DB_CONFIG_WORKER = {
        'host': '124.223.68.89',
        'user': 'root',
        'password': 'Mjhu666777;', # Replace with your actual password
        'database': 'ShenJiao'
    }
    DB_CONFIG = TEST_DB_CONFIG_WORKER # Set the global config for this test run

    test_file_id_for_add = "worker-db-test-add-002"
    sample_elements_for_add = [
        {"element_type": "paragraph", "content_id": "p1", "text_content": "This is a test paragraph from worker for add.", "level": None, "pageNo": 1},
        {"element_type": "heading", "content_id": "h1", "text_content": "Test Heading 1 for add", "level": 1, "pageNo": 1},
    ]

    print(f"Attempting to add/update elements for file_record_id: {test_file_id_for_add}")
    try:
        add_tmp_document_content_batch(test_file_id_for_add, sample_elements_for_add)
        print(f"\n[SUCCESS] Test data processed for {test_file_id_for_add}.")
        # ... (rest of the add_tmp_document_content_batch test messages) ...
    except Exception as e_add:
        print(f"\n[FAILURE] During add_tmp_document_content_batch test: {e_add}")

    # Test the new get functions
    # Assuming 'test-doc-003' is a file_record_id that has data in tmp_document_contents,
    # document_contents, and document_content_chunks from previous worker tests or manual insertion.
    test_file_id_for_get = "test-doc-003" 
    print(f"\n--- Testing Get Functions for file_record_id: {test_file_id_for_get} ---")
    
    print(f"\nFetching from tmp_document_contents (all types) for {test_file_id_for_get}:")
    tmp_contents_all = get_tmp_document_contents_by_file_id(test_file_id_for_get)
    if tmp_contents_all:
        print(f"  Found {len(tmp_contents_all)} records. First record: {tmp_contents_all[0] if tmp_contents_all else 'N/A'}")
    else:
        print(f"  No records found in tmp_document_contents for {test_file_id_for_get}.")

    print(f"\nFetching from tmp_document_contents (type 'main_paragraph') for {test_file_id_for_get}:")
    tmp_contents_para = get_tmp_document_contents_by_file_id(test_file_id_for_get, element_type="main_paragraph") # Example element type
    if tmp_contents_para:
        print(f"  Found {len(tmp_contents_para)} 'main_paragraph' records.")
    else:
        print(f"  No 'main_paragraph' records found in tmp_document_contents for {test_file_id_for_get}.")

    print(f"\nFetching from document_contents (all types) for {test_file_id_for_get}:")
    doc_contents_all = get_document_contents_by_file_id(test_file_id_for_get)
    if doc_contents_all:
        print(f"  Found {len(doc_contents_all)} records. First record: {doc_contents_all[0] if doc_contents_all else 'N/A'}")
    else:
        print(f"  No records found in document_contents for {test_file_id_for_get}.")
        print(f"  Note: This table might be populated by a different process or later step in a full workflow.")


    print(f"\nFetching from document_content_chunks for {test_file_id_for_get}:")
    chunks = get_document_content_chunks_by_file_id(test_file_id_for_get)
    if chunks:
        print(f"  Found {len(chunks)} chunk records. First chunk 'id': {chunks[0]['id'] if chunks else 'N/A'}, 'ai_content' preview: {chunks[0].get('ai_content', '')[:100] + '...' if chunks and chunks[0].get('ai_content') else 'N/A'}")
    else:
        print(f"  No records found in document_content_chunks for {test_file_id_for_get}.")
        print(f"  Note: This table is expected to contain AI suggestions in JSON format for the advice generator.")

    print("\n--- End of Get Functions Test ---")

    print("\ndatabase_operations_worker.py test run finished.")
    print("Ensure mysql-connector-python is installed (`pip install mysql-connector-python`).")
