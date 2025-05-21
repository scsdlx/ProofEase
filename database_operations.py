import mysql.connector

# DB_CONFIG will be imported from app.py or defined here if running independently
DB_CONFIG = None

def get_db_connection():
    """Establishes a connection to the MySQL database."""
    if DB_CONFIG is None:
        raise ValueError("DB_CONFIG is not set. Please configure it in app.py")
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        return conn
    except mysql.connector.Error as err:
        # In a real app, app.logger.error(f"DB Connection Error: {err}") might be used if logger is configured
        print(f"DB Error in get_db_connection: {err}")
        return None

def get_all_file_records_for_display():
    """
    Fetches all file records from the 'file_records' table for display.
    Returns a list of dictionaries, where each dictionary represents a file record.
    """
    conn = get_db_connection()
    if conn is None:
        return []  # Return empty list if connection failed

    cursor = conn.cursor(dictionary=True)
    query = """
    SELECT id, original_filename, filesize, upload_time, status, proof_list_filepath
    FROM file_records
    ORDER BY upload_time DESC;
    """
    records = []
    try:
        cursor.execute(query)
        records = cursor.fetchall()
    except mysql.connector.Error as err:
        print(f"DB Error in get_all_file_records_for_display: {err}")
        # records will remain []
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
    return records

def get_file_record(file_record_id: str) -> dict | None:
    """
    Fetches a single complete file record by its ID.
    Returns a dictionary representing the file record, or None if not found or error.
    """
    conn = get_db_connection()
    if conn is None:
        return None

    cursor = conn.cursor(dictionary=True)
    query = "SELECT * FROM file_records WHERE id = %s;"
    record = None
    try:
        cursor.execute(query, (file_record_id,))
        record = cursor.fetchone()
    except mysql.connector.Error as err:
        print(f"DB Error in get_file_record for ID {file_record_id}: {err}")
        # record will remain None
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()
    return record

def update_file_record_status(file_record_id: str, new_status: str, error_message: str = None, proof_list_filepath: str = None) -> bool:
    """
    Updates the status, error_message, and optionally proof_list_filepath of a file record.
    """
    conn = get_db_connection()
    if conn is None:
        return False

    cursor = conn.cursor()
    query = """
    UPDATE file_records 
    SET status = %s, error_message = %s, proof_list_filepath = %s
    WHERE id = %s;
    """
    # If proof_list_filepath is not provided, we might want to keep its existing value.
    # However, the current function signature implies setting it.
    # For this implementation, if None, it will set the DB field to NULL (or its default).
    # If only status and error_message should be updated, the query or logic needs adjustment.
    # Let's assume for now that all three are being explicitly set.
    # A more robust version might fetch the record first if only partial updates are common.
    # For now, this matches the function signature. Consider if proof_list_filepath should only be updated in specific statuses.

    # If proof_list_filepath is not being updated, fetch the current value first or adjust the query.
    # For this iteration, we'll update it. If it's None, it will set the DB field to NULL.
    # A common pattern is to only update fields that are explicitly passed.
    # Let's refine: if proof_list_filepath is None, we should not update it.
    
    current_record = get_file_record(file_record_id)
    if not current_record:
        return False # Record to update not found

    final_proof_list_filepath = proof_list_filepath if proof_list_filepath is not None else current_record.get('proof_list_filepath')

    try:
        cursor.execute(query, (new_status, error_message, final_proof_list_filepath, file_record_id))
        conn.commit()
        return cursor.rowcount > 0 # True if a row was updated
    except mysql.connector.Error as err:
        print(f"DB Error in update_file_record_status for ID {file_record_id}: {err}")
        conn.rollback()
        return False
    finally:
        if conn.is_connected():
            cursor.close()
            conn.close()

def create_file_record(original_filename, filesize, status="uploaded", file_hash=None, storage_path=None):
    """Creates a new file record in the database."""
    # To be implemented
    pass

if __name__ == '__main__':
    # Example Usage (requires DB_CONFIG to be set appropriately if run directly)
    # This is for testing purposes; DB_CONFIG would typically be set by app.py
    # For direct testing, you could temporarily define DB_CONFIG here:
    # DB_CONFIG = {'host': '124.223.68.89', 'user': 'root', 'password': 'Mjhu666777;', 'database': 'ShenJiao'}
    # print("Attempting to fetch records directly from database_operations.py (ensure DB_CONFIG is set for this test)")
    # records = get_all_file_records_for_display()
    # if records:
    #     for record in records:
    #         print(record)
    # else:
    #     print("No records found or connection failed.")
    print("database_operations.py loaded. Contains DB interaction functions.")
    print("To test functions directly, uncomment and configure DB_CONFIG within the __main__ block.")
