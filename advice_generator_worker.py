import json
import win32com.client
import os
import pythoncom
import datetime # For timestamp in output filename

# Assuming database_operations_worker.py is in the same directory or PYTHONPATH
import database_operations_worker as db_worker

# --- LCS and alignment functions (from original genShenJiaoAdvice.py) ---
def _calculate_lcs_and_reconstruct(s1: str, s2: str) -> tuple[str, int]:
    n = len(s1)
    m = len(s2)
    if n == 0 or m == 0: return "", 0
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            if s1[i-1] == s2[j-1]: dp[i][j] = dp[i-1][j-1] + 1
            else: dp[i][j] = max(dp[i-1][j], dp[i][j-1])
    lcs_length = dp[n][m]
    if lcs_length == 0: return "", 0
    lcs_chars = []
    i, j = n, m
    while i > 0 and j > 0:
        if s1[i-1] == s2[j-1]: lcs_chars.append(s1[i-1]); i -= 1; j -= 1
        elif dp[i-1][j] > dp[i][j-1]: i -= 1
        else: j -= 1
    return "".join(reversed(lcs_chars)), lcs_length

def get_alignment_details(s1: str, s2: str) -> tuple[str, int, float]:
    s1 = str(s1) if s1 is not None else ""
    s2 = str(s2) if s2 is not None else ""
    common_sequence, common_length = _calculate_lcs_and_reconstruct(s1, s2)
    len_s1 = len(s1); len_s2 = len(s2)
    if len_s1 == 0 and len_s2 == 0: similarity_score = 1.0
    elif common_length == 0 or (len_s1 + len_s2 == 0): similarity_score = 0.0
    else: similarity_score = (2 * common_length) / (len_s1 + len_s2)
    return common_sequence, common_length, similarity_score

# --- Word color constants (RGB) and other Word constants ---
WD_COLOR_RED = 255
WD_COLOR_DARK_GREEN = 32768 # RGB(0, 128, 0)
WD_COLOR_LIGHT_GREEN_BG = 14476252 # RGB(220, 240, 220)

WD_COLLAPSE_END = 0
WD_COLOR_INDEX_AUTO = 0
WD_NO_HIGHLIGHT = 0

# --- Status Translation ---
STATUS_TRANSLATION = {
    "pending": "待修改",
    "accepted": "已接受修改建议",
    "denied": "已拒绝修改建议",
    "accepted-edited": "已手动修改"
}

def generate_advice_document_from_db(file_record_id: str, output_base_dir: str, db_config: dict) -> tuple[bool, str]:
    """
    Generates a Word document with proofreading advice by fetching data from the database.
    """
    pythoncom.CoInitialize()
    db_worker.DB_CONFIG = db_config # Configure DB for this worker instance

    word_app = None
    doc = None

    try:
        # 1. Data Fetching
        print(f"[INFO] Fetching data for file_record_id: {file_record_id}")
        
        # Fetch from tmp_document_contents (analogous to word_content_analysis.xlsx)
        # Assuming 'main_paragraph' is the equivalent type. Adjust if worker stores differently.
        wca_paragraphs_raw = db_worker.get_tmp_document_contents_by_file_id(file_record_id, element_type="main_paragraph")
        if not wca_paragraphs_raw:
            # Try fetching without element_type if specific type yields no results (broader search)
            wca_paragraphs_raw = db_worker.get_tmp_document_contents_by_file_id(file_record_id)
            if not wca_paragraphs_raw:
                 print(f"[WARNING] No 'main_paragraph' or any other elements found in tmp_document_contents for {file_record_id}.")
                 # Depending on strictness, could return False here or try to proceed if other data exists.
                 # For now, let's assume paragraphs are essential.
                 # return False, f"No paragraph content found in tmp_document_contents for {file_record_id}."


        # Fetch from document_contents (analogous to document_contents.xlsx)
        # Assuming 'paragraph' is the equivalent type.
        dc_paragraphs_raw = db_worker.get_document_contents_by_file_id(file_record_id, element_type="paragraph")
        if not dc_paragraphs_raw:
            dc_paragraphs_raw = db_worker.get_document_contents_by_file_id(file_record_id) # Broader search
            if not dc_paragraphs_raw:
                print(f"[WARNING] No 'paragraph' or any other elements found in document_contents for {file_record_id}.")
                # This might be acceptable if suggestions are linked directly to tmp_document_contents via content_id.
                # The original script logic implies matching wca_paragraphs to dc_paragraphs.
                # If dc_paragraphs are essential for matching, this could be an error.
                # For now, we'll allow it to proceed and see if matches are found.

        # Fetch from document_content_chunks (for AI suggestions)
        dcc_chunks_raw = db_worker.get_document_content_chunks_by_file_id(file_record_id)
        if not dcc_chunks_raw:
            return False, f"No AI suggestion chunks (document_content_chunks) found for {file_record_id}."

        # Basic data transformation (similar to original script's fillna and astype)
        wca_paragraphs = []
        for p_dict in wca_paragraphs_raw:
            p_dict['text_content'] = str(p_dict.get('text_content', ''))
            wca_paragraphs.append(p_dict)

        dc_paragraphs = []
        for p_dict in dc_paragraphs_raw:
            p_dict['text_content'] = str(p_dict.get('text_content', ''))
            dc_paragraphs.append(p_dict)
        
        print(f"[INFO] Fetched {len(wca_paragraphs)} 'wca_paragraphs', {len(dc_paragraphs)} 'dc_paragraphs', {len(dcc_chunks_raw)} 'dcc_chunks'.")

        # 2. AI Suggestion Parsing
        all_suggestions = []
        for chunk_row in dcc_chunks_raw:
            try:
                ai_content_json = chunk_row.get('ai_content')
                if ai_content_json and isinstance(ai_content_json, str):
                    suggestions = json.loads(ai_content_json)
                    if isinstance(suggestions, list):
                        for sugg in suggestions:
                            # Validate that suggestion has '材料id' (content_id it refers to)
                            if isinstance(sugg, dict) and '材料id' in sugg:
                                all_suggestions.append(sugg)
                            else:
                                print(f"[WARNING] Suggestion missing '材料id' or not a dict in chunk id {chunk_row.get('id')}: {sugg}")
                    else:
                         print(f"[WARNING] Parsed 'ai_content' is not a list in chunk id {chunk_row.get('id')}: {suggestions}")
                elif chunk_row.get('ai_content') is None:
                     print(f"[WARNING] Empty 'ai_content' in document_content_chunks (id: {chunk_row.get('id', 'N/A')})")
                else: # Not a string, could be already parsed by DB connector if JSON type in DB
                    if isinstance(chunk_row.get('ai_content'), list):
                        suggestions = chunk_row.get('ai_content')
                        for sugg in suggestions:
                            if isinstance(sugg, dict) and '材料id' in sugg:
                                 all_suggestions.append(sugg)
                    else:
                        print(f"[WARNING] 'ai_content' has unexpected type (id: {chunk_row.get('id', 'N/A')}) - type: {type(ai_content_json)}. Content: {ai_content_json}")

            except json.JSONDecodeError as e:
                print(f"[ERROR] JSON parsing failed for 'ai_content' in chunk id {chunk_row.get('id')}: {e}. Content: {chunk_row.get('ai_content')}")
            except TypeError as e:
                 print(f"[ERROR] TypeError during suggestion parsing for chunk id {chunk_row.get('id')}: {e}. Content: {chunk_row.get('ai_content')}")
        
        if not all_suggestions:
            print(f"[WARNING] No valid suggestions parsed from document_content_chunks for {file_record_id}.")
            # Decide if this is a hard error or if an empty report should be generated.
            # For now, let's generate an empty/minimal report if no suggestions.

        print(f"[INFO] Total valid suggestions parsed: {len(all_suggestions)}")

        # 3. Initialize Word Application and Document
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False # Run in background
            doc = word_app.Documents.Add()
        except Exception as e_word_init:
            return False, f"Failed to initialize Word application: {e_word_init}"

        first_suggestion_written_to_doc = False

        # 4. Core Logic: Match paragraphs and generate Word content
        # This loop iterates through 'tmp_document_contents' paragraphs (wca_paragraphs)
        for wca_para_data in wca_paragraphs:
            source_text = wca_para_data['text_content']
            page_no = wca_para_data.get('pageNo', 'N/A') # Get page number
            # wca_content_id = wca_para_data.get('content_id') # ID from tmp_document_contents

            # Find the best matching paragraph from 'document_contents' (dc_paragraphs)
            # This matching logic is from the original script.
            # The assumption is that AI suggestions in 'dcc_chunks_raw' use '材料id' that
            # corresponds to 'content_id' from 'document_contents' (dc_paragraphs).
            
            best_match_dc_content_id = None
            best_match_dc_text_content = ""
            max_similarity_score = 0.0

            # If dc_paragraphs is empty, this loop will be skipped.
            for dc_para_data in dc_paragraphs:
                target_text = dc_para_data['text_content']
                # The 'content_id' from dc_para_data is what AI suggestions should link to.
                # dc_content_id_for_match = dc_para_data.get('content_id')
                
                _, _, similarity = get_alignment_details(source_text, target_text)
                
                # Threshold from original script was 0.75
                if similarity >= 0.75 and similarity > max_similarity_score:
                    max_similarity_score = similarity
                    best_match_dc_content_id = dc_para_data.get('content_id') # This is the '材料id'
                    best_match_dc_text_content = target_text
            
            if not best_match_dc_content_id:
                # print(f"[DEBUG] No strong match found in 'document_contents' for tmp_content paragraph (content_id: {wca_content_id}, text: '{source_text[:50]}...'). Skipping suggestions for this one.")
                continue # Skip if no good match in document_contents
            
            # Filter suggestions that are relevant to this matched dc_paragraph
            relevant_suggestions_for_dc_para = [s for s in all_suggestions if s.get('材料id') == best_match_dc_content_id]

            if not relevant_suggestions_for_dc_para:
                continue # No suggestions for this specific matched paragraph

            # --- Start writing to Word doc for this matched paragraph and its suggestions ---
            for suggestion_item in relevant_suggestions_for_dc_para:
                json_original_text_from_suggestion = str(suggestion_item.get('原始内容', ''))
                json_modified_text_from_suggestion = str(suggestion_item.get('修改后内容', ''))
                json_status_raw = str(suggestion_item.get('status', 'N/A')) # e.g., "pending", "accepted"
                json_reason_text = str(suggestion_item.get('出错原因', '无原因说明')) # Note: field name is "出错原因"

                translated_status = STATUS_TRANSLATION.get(json_status_raw, json_status_raw)

                if first_suggestion_written_to_doc:
                    hr_para = doc.Paragraphs.Add()
                    try: hr_para.Range.InsertHorizontalLine()
                    except: hr_para.Range.Text = "------------------------------------------------------------\n"
                else:
                    first_suggestion_written_to_doc = True
                
                # Page number and original content header
                para_page_ref = doc.Paragraphs.Add().Range
                para_page_ref.Text = f"页码：{page_no}\n" # Page number from wca_para_data
                
                doc.Paragraphs.Add().Range.Text = "原始内容：" # Header for the original content block
                
                # Get the last paragraph's range to append content inline
                current_doc_range = doc.Paragraphs.Last.Range 
                current_doc_range.Collapse(WD_COLLAPSE_END) # Collapse to end to append

                # The 'best_match_dc_text_content' is the full original paragraph from document_contents
                # The 'json_original_text_from_suggestion' is the specific part AI focused on.
                original_para_text_full = str(best_match_dc_text_content)

                # Highlighting logic from original script
                if json_original_text_from_suggestion and json_original_text_from_suggestion in original_para_text_full:
                    start_idx = original_para_text_full.find(json_original_text_from_suggestion)
                    end_idx = start_idx + len(json_original_text_from_suggestion)

                    part_before_highlight = original_para_text_full[:start_idx]
                    part_to_delete_highlight = original_para_text_full[start_idx:end_idx] # Text to be "deleted"
                    part_after_highlight = original_para_text_full[end_idx:]

                    # Write part before
                    current_doc_range.InsertAfter(part_before_highlight)
                    current_doc_range.Collapse(WD_COLLAPSE_END)

                    # Write "deleted" part (strikethrough, red)
                    current_doc_range.InsertAfter(part_to_delete_highlight)
                    rng_deleted_text = doc.Range(current_doc_range.End - len(part_to_delete_highlight), current_doc_range.End)
                    rng_deleted_text.Font.Color = WD_COLOR_RED
                    rng_deleted_text.Font.StrikeThrough = True
                    current_doc_range.Collapse(WD_COLLAPSE_END)

                    # Write "added" part (green, background highlight)
                    current_doc_range.InsertAfter(json_modified_text_from_suggestion)
                    rng_added_text = doc.Range(current_doc_range.End - len(json_modified_text_from_suggestion), current_doc_range.End)
                    rng_added_text.Font.Color = WD_COLOR_DARK_GREEN
                    rng_added_text.Shading.BackgroundPatternColor = WD_COLOR_LIGHT_GREEN_BG
                    rng_added_text.Font.StrikeThrough = False # Ensure no strikethrough for added
                    current_doc_range.Collapse(WD_COLLAPSE_END)
                    
                    # Write status
                    status_display_text = f"【{translated_status}】"
                    current_doc_range.InsertAfter(status_display_text)
                    rng_status_text = doc.Range(current_doc_range.End - len(status_display_text), current_doc_range.End)
                    rng_status_text.Font.ColorIndex = WD_COLOR_INDEX_AUTO
                    rng_status_text.Shading.BackgroundPatternColorIndex = WD_NO_HIGHLIGHT
                    rng_status_text.Font.StrikeThrough = False
                    current_doc_range.Collapse(WD_COLLAPSE_END)

                    # Write part after
                    current_doc_range.InsertAfter(part_after_highlight)
                    if part_after_highlight: # Reset formatting for the part after
                        rng_after_text = doc.Range(current_doc_range.End - len(part_after_highlight), current_doc_range.End)
                        rng_after_text.Font.ColorIndex = WD_COLOR_INDEX_AUTO
                        rng_after_text.Shading.BackgroundPatternColorIndex = WD_NO_HIGHLIGHT
                        rng_after_text.Font.StrikeThrough = False
                    current_doc_range.Collapse(WD_COLLAPSE_END)
                else:
                    # Fallback if '原始内容' from JSON is not found in the matched dc_paragraph text
                    print(f"[WARNING] JSON '原始内容' ('{json_original_text_from_suggestion}') not found in matched document_contents paragraph ('{original_para_text_full[:50]}...'). Appending suggestion details.")
                    current_doc_range.InsertAfter(original_para_text_full) # Insert the full original paragraph
                    current_doc_range.Collapse(WD_COLLAPSE_END)

                    current_doc_range.InsertAfter(" (建议修改为：")
                    current_doc_range.Collapse(WD_COLLAPSE_END)
                    current_doc_range.InsertAfter(json_modified_text_from_suggestion)
                    rng_added_fallback = doc.Range(current_doc_range.End - len(json_modified_text_from_suggestion), current_doc_range.End)
                    rng_added_fallback.Font.Color = WD_COLOR_DARK_GREEN
                    rng_added_fallback.Shading.BackgroundPatternColor = WD_COLOR_LIGHT_GREEN_BG
                    rng_added_fallback.Font.StrikeThrough = False
                    current_doc_range.Collapse(WD_COLLAPSE_END)
                    current_doc_range.InsertAfter("）")
                    current_doc_range.Collapse(WD_COLLAPSE_END)
                    
                    status_display_text_fb = f"【{translated_status}】"
                    current_doc_range.InsertAfter(status_display_text_fb)
                    rng_status_fb = doc.Range(current_doc_range.End - len(status_display_text_fb), current_doc_range.End)
                    rng_status_fb.Font.ColorIndex = WD_COLOR_INDEX_AUTO
                    rng_status_fb.Shading.BackgroundPatternColorIndex = WD_NO_HIGHLIGHT
                    rng_status_fb.Font.StrikeThrough = False
                    current_doc_range.Collapse(WD_COLLAPSE_END)

                current_doc_range.InsertAfter("\n") # Newline after the content block

                # Add reason
                doc.Paragraphs.Add().Range.Text = f"原因：{json_reason_text}\n"
        
        # --- End of loops, prepare to save ---
        if not first_suggestion_written_to_doc:
            doc.Paragraphs.Add().Range.Text = "未找到与原始文档内容匹配的有效审校建议。"
            print(f"[INFO] No suggestions were written to the document for {file_record_id}.")

        # 5. Output Document
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"prooflist_{file_record_id}_{timestamp}.docx"
        if not os.path.exists(output_base_dir):
            os.makedirs(output_base_dir)
        full_output_path = os.path.join(output_base_dir, output_filename)

        try:
            doc.SaveAs(full_output_path)
            print(f"[SUCCESS] Proofreading advice list saved to: {full_output_path}")
            return True, full_output_path
        except Exception as e_save:
            return False, f"Failed to save Word document: {e_save}"

    except mysql.connector.Error as db_err:
        print(f"[FATAL] Database error during advice generation for {file_record_id}: {db_err}")
        return False, f"Database error: {db_err}"
    except ConnectionError as conn_err: # Raised by our db_ops if connection fails initially
        print(f"[FATAL] Database connection error for {file_record_id}: {conn_err}")
        return False, f"DB connection error: {conn_err}"
    except Exception as e_main:
        import traceback
        print(f"[FATAL] An unexpected error occurred in advice_generator_worker for {file_record_id}: {e_main}")
        traceback.print_exc()
        return False, f"Unexpected error: {e_main}"
    finally:
        if doc:
            try: doc.Close(False) # Close document without saving changes (already saved with SaveAs)
            except Exception as e_close: print(f"[WARNING] Error closing document: {e_close}")
        if word_app:
            try: word_app.Quit()
            except Exception as e_quit: print(f"[WARNING] Error quitting Word application: {e_quit}")
        pythoncom.CoUninitialize()


if __name__ == '__main__':
    print("Running advice_generator_worker.py directly (for testing purposes)...")

    # --- Configuration for Direct Test ---
    TEST_DB_CONFIG_ADVICE_GEN = {
        'host': '124.223.68.89',
        'user': 'root',
        'password': 'Mjhu666777;', # Replace with your actual password
        'database': 'ShenJiao'
    }
    
    # Replace with a file_record_id that has:
    # 1. Entries in `tmp_document_contents` (from word_extractor_worker)
    # 2. Entries in `document_contents` (if your workflow populates this separately, or use same as tmp for test)
    # 3. Entries in `document_content_chunks` with valid AI suggestion JSON in `ai_content`.
    # test_file_record_id_advice = "test-doc-003" # Example ID from previous tests
    test_file_record_id_advice = "test-doc-003" # Use an ID you know has data

    # Directory on this (worker) machine to save the generated .docx file
    test_output_dir = os.path.join(os.getcwd(), "generated_advice_docs")
    if not os.path.exists(test_output_dir):
        os.makedirs(test_output_dir)

    print(f"Test parameters for advice generation:")
    print(f"  File Record ID: {test_file_record_id_advice}")
    print(f"  Output Directory: {test_output_dir}")
    print(f"  DB Host: {TEST_DB_CONFIG_ADVICE_GEN['host']}, DB Name: {TEST_DB_CONFIG_ADVICE_GEN['database']}")
    
    # Ensure dependencies are installed:
    # pip install pywin32 mysql-connector-python
    
    input(f"Ensure that file_record_id '{test_file_record_id_advice}' has relevant data in tmp_document_contents, document_contents, and document_content_chunks. Press Enter to continue...")

    success_gen, message_gen = generate_advice_document_from_db(
        test_file_record_id_advice,
        test_output_dir,
        TEST_DB_CONFIG_ADVICE_GEN
    )

    if success_gen:
        print(f"\n[SUCCESS] Advice document generation test completed.")
        print(f"  Generated document path: {message_gen}")
    else:
        print(f"\n[FAILURE] Advice document generation test failed.")
        print(f"  Error/Message: {message_gen}")

    print("\nadvice_generator_worker.py test run finished.")
