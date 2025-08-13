# word_parser_for_material.py

import os
import re
import win32com.client as win32
import pythoncom
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - [word_parser_for_material] - %(message)s')

def clean_text(text):
    """Removes invalid XML characters from text."""
    # This regex is more robust for database insertion
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)

def parse_word_to_db(doc_path, material_id, num_levels_to_extract, db_cursor):
    """
    Parses a Word document by OutlineLevel, and inserts the hierarchical content
    into the 'material_contents' table using the provided database cursor.

    Args:
        doc_path (str): The local path to the Word document.
        material_id (int): The ID of the material this content belongs to.
        num_levels_to_extract (int): The number of heading levels to parse.
        db_cursor: An active database cursor for executing SQL commands.

    Returns:
        dict: A summary of the operation.
    """
    word_app = None
    coinitialized = False
    result_summary = {"success": False, "message": "", "rows_inserted": 0}

    try:
        # --- Initialize COM and Word ---
        pythoncom.CoInitialize()
        coinitialized = True
        logging.info("Starting Word application for parsing...")
        word_app = win32.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = False

        # --- Open Document ---
        doc = word_app.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
        logging.info(f"Successfully opened document: {os.path.basename(doc_path)}")

        # --- Aggregate Content (similar to parseWord2Excel) ---
        doc_content_aggregator = {}
        current_headings = {f'标题{i}': '' for i in range(1, 10)}

        # Find the actual starting level of headings in the document
        all_levels_in_doc = {p.OutlineLevel for p in doc.Paragraphs if 1 <= p.OutlineLevel <= 9}
        if not all_levels_in_doc:
            doc.Close(SaveChanges=False)
            result_summary["message"] = "Warning: No outline-level headings (1-9) found in the document. No content was parsed."
            result_summary["success"] = True # Success in the sense that the process ran without error
            return result_summary
        
        min_level_in_doc = min(all_levels_in_doc)
        logging.info(f"Document's top heading level detected as: {min_level_in_doc}. Parsing up to {num_levels_to_extract} levels from there.")

        # Iterate through paragraphs to aggregate content under headings
        for para in doc.Paragraphs:
            # Extract text with automatic numbering
            raw_text = para.Range.Text
            list_string = para.Range.ListFormat.ListString
            full_text = f"{list_string} {raw_text}" if list_string else raw_text
            para_text = clean_text(full_text).strip()

            if not para_text:
                continue

            level = para.OutlineLevel
            if 1 <= level <= 9:
                # Update current heading state
                current_headings[f'标题{level}'] = para_text
                # Reset lower-level headings
                for L in range(level + 1, 10):
                    current_headings[f'标题{L}'] = ''

            # Create the key for the aggregator dictionary
            key_headings = []
            for j in range(num_levels_to_extract):
                absolute_level_num = min_level_in_doc + j
                if absolute_level_num <= 9:
                    heading_text = current_headings.get(f'标题{absolute_level_num}', '')
                    key_headings.append(heading_text)
            key_tuple = tuple(key_headings)

            # Ignore content that doesn't fall under any specified heading
            if not any(key_tuple):
                continue
            
            # Aggregate content
            if key_tuple not in doc_content_aggregator:
                doc_content_aggregator[key_tuple] = []
            doc_content_aggregator[key_tuple].append(para_text)

        doc.Close(SaveChanges=False)
        logging.info("Content aggregation complete. Preparing for database insertion.")

        # --- Insert into Database ---
        # Sort items to ensure parent nodes are processed before children
        sorted_items = sorted(doc_content_aggregator.items(), key=lambda item: item[0])

        parent_id_map = {}  # Maps a heading tuple to its database ID
        sequence_counters = {} # Maps a parent_id to its child sequence number
        rows_inserted = 0
        now_utc = datetime.utcnow()

        for headings_tuple, content_list in sorted_items:
            # Determine the level and title for the current node
            current_level = 0
            current_title = ""
            for i, h in enumerate(headings_tuple):
                if h:
                    current_level = i + 1
                    current_title = h
            
            if not current_title: # Skip if there's no title for this entry
                continue

            # Find parent_id
            parent_tuple = headings_tuple[:current_level - 1]
            parent_id = parent_id_map.get(parent_tuple, None)

            # Get sequence number
            seq_key = parent_id if parent_id is not None else 'root'
            sequence = sequence_counters.get(seq_key, 0) + 1
            sequence_counters[seq_key] = sequence

            # Prepare redundant title fields (title1, title2, etc.)
            title_fields = {f'title{i+1}': None for i in range(8)}
            for i in range(current_level):
                if i < 8:
                    title_fields[f'title{i+1}'] = headings_tuple[i][:255] # Truncate to fit VARCHAR(255)
            
            # Join content
            full_content = "\n".join(content_list).strip()

            # Prepare SQL insertion
            sql = """
                INSERT INTO material_contents
                (material_id, parent_id, level, sequence, title, 
                 title1, title2, title3, title4, title5, title6, title7, title8,
                 content, created_at, updated_at)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            """
            params = (
                material_id, parent_id, current_level, sequence, current_title[:255],
                title_fields['title1'], title_fields['title2'], title_fields['title3'], title_fields['title4'],
                title_fields['title5'], title_fields['title6'], title_fields['title7'], title_fields['title8'],
                full_content, now_utc, now_utc
            )

            db_cursor.execute(sql, params)
            new_id = db_cursor.lastrowid
            rows_inserted += 1
            
            # Store the new ID for potential children
            parent_id_map[headings_tuple[:current_level]] = new_id

        result_summary["success"] = True
        result_summary["message"] = f"Successfully parsed and inserted {rows_inserted} content sections."
        result_summary["rows_inserted"] = rows_inserted
        logging.info(result_summary["message"])

    except pythoncom.com_error as e:
        err_msg = f"A COM error occurred during Word processing: {e}"
        logging.error(err_msg, exc_info=True)
        result_summary["message"] = f"处理Word文档时发生内部错误 (COM): {e.args[2][2] if len(e.args) > 2 and e.args[2] else str(e)}"
        raise  # Re-raise to be caught by the API handler for DB rollback
    except Exception as e:
        err_msg = f"An unexpected error occurred: {e}"
        logging.error(err_msg, exc_info=True)
        result_summary["message"] = f"解析过程中发生未知错误: {e}"
        raise # Re-raise for rollback
    finally:
        # --- Cleanup ---
        if word_app:
            try:
                word_app.Quit(SaveChanges=False)
                logging.info("Word application has been closed.")
            except Exception as e_quit:
                logging.error(f"Error while quitting Word: {e_quit}")
        if coinitialized:
            pythoncom.CoUninitialize()
    
    return result_summary