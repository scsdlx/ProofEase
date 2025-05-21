import win32com.client
import os
from PIL import ImageGrab
import pythoncom
import pandas as pd
import re

# (Constants and other helper functions remain the same)
WD_OUTLINE_LEVEL_BODY_TEXT = 10
EXCEL_CELL_CHAR_LIMIT= 32767    # Excel cell character limit    

try:
    win32com.client.gencache.EnsureDispatch("Word.Application")
    WD_OUTLINE_LEVEL_BODY_TEXT = win32com.client.constants.wdOutlineLevelBodyText
except AttributeError:
    print("[WARNING] win32com.client.constants.wdOutlineLevelBodyText not found, using hardcoded value 10.")
except Exception as e_gencache:
    print(f"[WARNING] gencache.EnsureDispatch or constant loading failed: {e_gencache}")

# ... (WD_STORY_TYPES, WD_SHAPE_TYPES, EXCEL_CELL_CHAR_LIMIT, clean_text_for_excel, save_image_from_clipboard, get_page_number_from_range, format_table_for_excel) ...
def clean_text_for_excel(text): # ...
    if not isinstance(text, str): text = str(text)
    cleaned_text = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\U00010000-\U0010FFFF]', '', text)
    if len(cleaned_text) > EXCEL_CELL_CHAR_LIMIT:
        return cleaned_text[:EXCEL_CELL_CHAR_LIMIT] + "... (truncated)"
    return cleaned_text

def save_image_from_clipboard(filepath): # ...
    try:
        image = ImageGrab.grabclipboard()
        if image:
            if image.mode == 'RGBA' or image.mode == 'P': image = image.convert('RGB')
            image.save(filepath, 'PNG')
            return True
    except Exception: pass
    return False

def get_page_number_from_range(item_range): # ...
    if item_range:
        try: return item_range.Information(win32com.client.constants.wdActiveEndPageNumber)
        except Exception: pass
    return None 

def format_table_for_excel(table_content_list_of_lists): # ...
    if not table_content_list_of_lists: return ""
    formatted_rows = []
    for row in table_content_list_of_lists:
        cleaned_row = [clean_text_for_excel(str(cell_content)) for cell_content in row]
        formatted_rows.append(" | ".join(cleaned_row))
    return "\n".join(formatted_rows)


def _reconstruct_text_with_note_references(owner_range, note_collection_getter, 
                                           page_level_note_counts): # MODIFIED: Accept page_level_note_counts
    original_text_with_cr = owner_range.Text
    owner_range_start_offset = owner_range.Start

    notes_in_range_data = []
    # ... (Data collection for notes_in_range_data - same as before, ensuring "page_of_reference" is collected)
    try:
        for note_obj in note_collection_getter(owner_range):
            try:
                if not (hasattr(note_obj, "Reference") and note_obj.Reference and 
                        hasattr(note_obj.Reference, "Start") and hasattr(note_obj.Reference, "End")):
                    continue

                ref_start_doc = note_obj.Reference.Start
                ref_end_doc = note_obj.Reference.End

                if not (owner_range_start_offset <= ref_start_doc < ref_end_doc <= owner_range.End):
                    continue
                
                ref_start_in_owner = ref_start_doc - owner_range_start_offset
                ref_end_in_owner = ref_end_doc - owner_range_start_offset

                if not (0 <= ref_start_in_owner < ref_end_in_owner <= len(original_text_with_cr)):
                    continue
                
                page_of_reference = get_page_number_from_range(note_obj.Reference)

                notes_in_range_data.append({
                    "ref_start_in_owner": ref_start_in_owner,
                    "ref_end_in_owner": ref_end_in_owner,
                    "page_of_reference": page_of_reference,
                })
            except Exception: continue
    except Exception:
        return original_text_with_cr.strip() 

    if not notes_in_range_data:
        return original_text_with_cr.strip()
    # ... (Sort notes_in_range_data - same as before)
    notes_in_range_data.sort(key=lambda x: x["ref_start_in_owner"])


    new_text_parts = []
    last_pos_in_owner = 0
    
    # --- PER-PAGE NUMBERING LOGIC ---
    # page_level_note_counts is now passed in and MODIFIED by this function.
    # It will persist across calls within the same parse_range_content scope.
    
    for note_data in notes_in_range_data:
        if note_data["ref_start_in_owner"] < last_pos_in_owner: continue
        if note_data["ref_start_in_owner"] > len(original_text_with_cr): break

        new_text_parts.append(original_text_with_cr[last_pos_in_owner : note_data["ref_start_in_owner"]])
        
        current_page_for_this_note = note_data["page_of_reference"]
        mark_to_display = "?" 

        if current_page_for_this_note is not None:
            # Use and update the passed-in page_level_note_counts
            page_level_note_counts[current_page_for_this_note] = page_level_note_counts.get(current_page_for_this_note, 0) + 1
            mark_to_display = str(page_level_note_counts[current_page_for_this_note])
        
        new_text_parts.append(f"[{mark_to_display}]")

        last_pos_in_owner = note_data["ref_end_in_owner"]
        if last_pos_in_owner > len(original_text_with_cr):
            last_pos_in_owner = len(original_text_with_cr)

    if last_pos_in_owner < len(original_text_with_cr):
        new_text_parts.append(original_text_with_cr[last_pos_in_owner:])
    
    return "".join(new_text_parts)


def parse_range_content(doc_range, word_app, output_elements, image_dir, img_counter, 
                        element_prefix=""):
    processed_table_ids = set()
    if not doc_range: return

    # --- MODIFIED: Initialize page_local_footnote_counts here ---
    # This dictionary will track footnote counts per page *for the current doc_range being processed*
    # (e.g., for doc.Content, or for a specific header/footer range)
    page_local_footnote_counts = {}

    try:
        paragraphs_collection = doc_range.Paragraphs
    except Exception: return

    for para_idx, para in enumerate(paragraphs_collection):
        try:
            para_range_obj = para.Range
            
            final_para_text_for_output = _reconstruct_text_with_note_references(
                para_range_obj, 
                lambda r: r.Footnotes,
                page_local_footnote_counts # MODIFIED: Pass the dictionary
            ).strip()

            current_page = get_page_number_from_range(para_range_obj)
            is_in_table = False
            try: is_in_table = para_range_obj.Information(win32com.client.constants.wdWithInTable)
            except Exception: pass

            if is_in_table:
                try:
                    table = para_range_obj.Tables(1)
                    if table.ID not in processed_table_ids:
                        table_data = []
                        for r_idx, row in enumerate(table.Rows):
                            row_data_cells = []
                            for c_idx, cell in enumerate(row.Cells):
                                cell_range_obj = cell.Range
                                # Table cells will also use the same page_local_footnote_counts
                                # from the parent parse_range_content call. This means numbering
                                # will be continuous across paragraphs and then into table cells on the same page.
                                final_cell_text_intermediate = _reconstruct_text_with_note_references(
                                    cell_range_obj,
                                    lambda r: r.Footnotes,
                                    page_local_footnote_counts # MODIFIED: Pass the dictionary
                                )
                                final_cell_text = final_cell_text_intermediate.strip().replace('\r\x07', '').replace('\x07', '')
                                row_data_cells.append(final_cell_text)
                            table_data.append(row_data_cells)
                        
                        output_elements.append({
                            "type": f"{element_prefix}table", "id": table.ID,
                            "content_data": table_data, "rows": table.Rows.Count,
                            "columns": table.Columns.Count, "page_number": get_page_number_from_range(table.Range), "level": None
                        })
                        processed_table_ids.add(table.ID)
                except Exception: pass
                continue 

            if final_para_text_for_output:
                # ... (heading/paragraph classification and appending to output_elements - same as before) ...
                style_name = para.Style.NameLocal
                outline_level_val = WD_OUTLINE_LEVEL_BODY_TEXT
                try: outline_level_val = para.OutlineLevel
                except Exception: pass

                element_data = {
                    "text": final_para_text_for_output, "style": style_name,
                    "page_number": current_page, "level": None 
                }
                if 1 <= outline_level_val <= 9:
                    element_data["type"] = f"{element_prefix}heading"
                    element_data["level"] = outline_level_val
                else:
                    element_data["type"] = f"{element_prefix}paragraph"
                output_elements.append(element_data)

            if para_range_obj.InlineShapes.Count > 0:
                # ... (inline shape logic - same as before) ...
                for i_shape_idx, i_shape in enumerate(para_range_obj.InlineShapes):
                    img_counter[0] += 1
                    shape_page = get_page_number_from_range(i_shape.Range)
                    try:
                        if i_shape.Type == win32com.client.constants.wdInlineShapePicture:
                            img_filename = f"inline_image_{img_counter[0]}.png"; img_filepath = os.path.join(image_dir, img_filename)
                            path_to_save = img_filepath; saved_successfully = False
                            try:
                                i_shape.Select(); word_app.Selection.Copy()
                                if save_image_from_clipboard(img_filepath): saved_successfully = True
                            except: pass                             
                            output_elements.append({
                                "type": f"{element_prefix}inline_image" if saved_successfully else f"{element_prefix}inline_image_extraction_failed",
                                "path": path_to_save, "info": "" if saved_successfully else "Extraction error.",
                                "width": i_shape.Width, "height": i_shape.Height, "page_number": shape_page, "level": None
                            })
                    except: pass

            # Endnotes: If you want similar per-page numbering for endnotes, you'd need a
            # separate page_local_endnote_counts dictionary managed similarly.
            if para_range_obj.Endnotes.Count > 0:
                 for en_ref_obj in para_range_obj.Endnotes:
                    try:
                        output_elements.append({
                            "type": f"{element_prefix}endnote_marker_info",
                            "text": f"Endnote Ref: [{en_ref_obj.Reference.Text.strip()}] (Doc Index: {en_ref_obj.Index})",
                            "page_number": get_page_number_from_range(en_ref_obj.Reference), "level": None
                        })
                    except: pass
        except: pass


# --- parse_word_document --- (No changes needed here, it calls parse_range_content correctly)
def parse_word_document(doc_path, image_output_dir="word_images"): # Same as before
    pythoncom.CoInitialize()
    word_app = None; doc = None; output_elements = []; img_counter = [0]
    if not os.path.exists(image_output_dir): os.makedirs(image_output_dir)

    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)

        print("[INFO] Pre-processing Footnotes Content (for separate listing)...")
        if doc.Footnotes.Count > 0:
            for fn_in_doc in doc.Footnotes:
                output_elements.append({
                    "type": "footnote_content_definition", "doc_wide_index": fn_in_doc.Index,
                    "text": fn_in_doc.Range.Text.strip(),
                    "reference_mark_in_content_area": fn_in_doc.Reference.Text.strip(),
                    "page_number": get_page_number_from_range(fn_in_doc.Range), "level": None
                })
        
        print("[INFO] Pre-processing Endnotes Content (for separate listing)...")
        if doc.Endnotes.Count > 0:
            for en_in_doc in doc.Endnotes:
                output_elements.append({
                    "type": "endnote_content_definition", "doc_wide_index": en_in_doc.Index,
                    "text": en_in_doc.Range.Text.strip(),
                    "reference_mark_in_content_area": en_in_doc.Reference.Text.strip(),
                    "page_number": get_page_number_from_range(en_in_doc.Range), "level": None
                })

        print("[INFO] Parsing Main Document Content (doc.Content)...")
        parse_range_content(doc.Content, word_app, output_elements, image_folder, img_counter)

        print("[INFO] Parsing Floating Shapes (doc.Shapes)...")
        for shape_idx, shape in enumerate(doc.Shapes):
            # ... (floating shape logic as before, ensuring parse_range_content is called without stores) ...
            img_counter[0] += 1 # Ensure counter is incremented for each shape
            anchor_text_preview = "N/A"; shape_page = None
            if hasattr(shape, "Anchor") and shape.Anchor:
                shape_page = get_page_number_from_range(shape.Anchor)
                if hasattr(shape.Anchor, "Paragraphs") and shape.Anchor.Paragraphs.Count > 0:
                    try: anchor_text_preview = shape.Anchor.Paragraphs(1).Range.Text[:30].strip() + "..."
                    except: pass
            
            try:
                shape_type = shape.Type
                if shape_type == WD_SHAPE_TYPES["msoPicture"]:
                     img_filename = f"floating_image_{img_counter[0]}.png"; img_filepath = os.path.join(image_dir, img_filename)
                     path_to_save = img_filepath; saved_successfully = False
                     try: 
                        shape.Select(); 
                        word_app.Selection.Copy();
                        if save_image_from_clipboard(img_filepath): saved_successfully = True
                     except: pass
                     output_elements.append({
                         "type": "floating_image" if saved_successfully else "floating_image_extraction_failed",
                         "path": path_to_save, "info": "" if saved_successfully else "Extraction error.",
                         "anchor_info": anchor_text_preview, "page_number": shape_page, "level": None })
                elif shape_type == WD_SHAPE_TYPES["msoTextBox"] or \
                     (shape_type == WD_SHAPE_TYPES.get("msoAutoShape") and hasattr(shape, "TextFrame") and shape.TextFrame and hasattr(shape.TextFrame,"HasText") and shape.TextFrame.HasText):
                    text_content_of_shape_element = ""; text_range_to_parse_recursively = None
                    if hasattr(shape, "TextFrame") and shape.TextFrame:
                        tf = shape.TextFrame
                        if hasattr(tf, "HasText") and tf.HasText and hasattr(tf, "TextRange"):
                            try: text_range_to_parse_recursively = tf.TextRange; text_content_of_shape_element = text_range_to_parse_recursively.Text.strip()
                            except: text_content_of_shape_element = "[Error accessing TextRange]"
                    output_elements.append({
                        "type": "textbox" if shape_type == WD_SHAPE_TYPES["msoTextBox"] else "autoshape_with_text",
                        "text": text_content_of_shape_element, "anchor_info": anchor_text_preview, 
                        "page_number": shape_page, "level": None })
                    if text_range_to_parse_recursively:
                        parse_range_content(text_range_to_parse_recursively, word_app, output_elements, 
                                            image_dir, img_counter, element_prefix=f"shape{shape_idx+1}_text_") # Pass correct args
                else:
                    shape_type_name = "unknown_mso_shape"
                    for name, val in WD_SHAPE_TYPES.items():
                        if shape_type == val: shape_type_name = name; break
                    output_elements.append({
                        "type": "other_floating_shape", "mso_shape_type_id": shape_type,
                        "mso_shape_type_name": shape_type_name, "name": shape.Name if hasattr(shape, "Name") else "N/A",
                        "page_number": shape_page, "anchor_info": anchor_text_preview, "level": None,
                        "text": f"[Other floating shape: {shape_type_name}]" })
            except: pass


        print("[INFO] Parsing Headers and Footers...")
        for section_idx, section in enumerate(doc.Sections):
            header_footer_types_map = { "Primary": win32com.client.constants.wdHeaderFooterPrimary, "FirstPage": win32com.client.constants.wdHeaderFooterFirstPage, "EvenPages": win32com.client.constants.wdHeaderFooterEvenPages }
            for hf_kind_name, hf_collection_obj_getter in [("Header", section.Headers), ("Footer", section.Footers)]:
                for hf_type_name, hf_type_constant in header_footer_types_map.items():
                    try:
                        hf_object = hf_collection_obj_getter(hf_type_constant)
                        if hf_object.Exists:
                            hf_range = None
                            try: hf_range = hf_object.Range
                            except Exception as e_hf_range:
                                output_elements.append({
                                    "type": f"section{section_idx+1}_{hf_kind_name.lower()}_{hf_type_name.lower()}_inaccessible_range",
                                    "text": f"H/F exists but Range inaccessible. Error: {e_hf_range}", "page_number": None, "level": None })
                                continue 
                            if hf_range and hf_range.Text.strip():
                                element_name_prefix = f"section{section_idx+1}_{hf_kind_name.lower()}_{hf_type_name.lower()}_"
                                parse_range_content(hf_range, word_app, output_elements, # Pass correct args
                                                    image_dir, img_counter, element_prefix=element_name_prefix)
                    except: pass
    except Exception as e_main: print(f"An error: {e_main}"); import traceback; traceback.print_exc()
    finally:
        if doc: 
            try: doc.Close(False); 
            except: pass
        if word_app: 
            try: word_app.Quit(); 
            except: pass
        pythoncom.CoUninitialize()
    return output_elements

# --- save_elements_to_excel --- (No changes)
def save_elements_to_excel(elements_list, excel_filepath, document_id): # Same as previous correct version
    excel_data = []
    content_item_id = 0

    for elem_idx, elem in enumerate(elements_list):
        content_item_id += 1
        elem_type = elem.get("type", "unknown")
        text_content_raw = elem.get("text", "") 
        level_val = elem.get("level") 

        if ("image" in elem_type or "failed" in elem_type) and "path" in elem :
            text_content_raw = elem.get("path", "")
            if not text_content_raw and "info" in elem: text_content_raw = elem.get("info", "")
        elif "table" in elem_type and "content_data" in elem:
            text_content_raw = format_table_for_excel(elem.get("content_data", []))
        
        if not text_content_raw and "info" in elem: text_content_raw = elem.get("info", "")
        
        final_text_content = clean_text_for_excel(text_content_raw)
        page_no_val = elem.get("page_number", "")

        row = {
            "file_record_id": document_id, "element_type": elem_type,
            "content_id": content_item_id, "text_content": final_text_content,
            "level": level_val if level_val is not None else "",
            "pageNo": page_no_val if page_no_val is not None else ""
        }
        excel_data.append(row)

    if not excel_data: print("No data to save to Excel."); return
    df = pd.DataFrame(excel_data)
    try:
        df = df[["file_record_id", "element_type", "content_id", "text_content", "level", "pageNo"]]
    except KeyError as e_key: print(f"[WARNING] Excel column reorder failed: {e_key}.")
    try:
        df.to_excel(excel_filepath, index=False, engine='openpyxl')
        print(f"\n[SUCCESS] Extracted data saved to: {os.path.abspath(excel_filepath)}")
    except Exception as e: print(f"\n[ERROR] Could not save Excel: {e}"); import traceback; traceback.print_exc()

# --- __main__ --- (No changes)
if __name__ == "__main__": # Same as before
    # doc_file_path = r"节选：《教师数字素养提升与应用》.docx"   
    doc_file_path = r"节选：创建有意识的机器 250422.docx"
    excel_output_path = "word_content_analysis.xlsx"
    image_folder = "parsed_word_images_excel"
    if doc_file_path == r"YOUR_WORD_DOCUMENT_PATH.docx" or not os.path.exists(doc_file_path):
        print(f"Error: Document not found or path not updated: '{doc_file_path}'")
    else:
        print(f"Starting parsing for: {doc_file_path}")
        file_id_for_excel = os.path.basename(doc_file_path)
        extracted_elements = parse_word_document(doc_file_path, image_output_dir=image_folder)
        print(f"\n--- Parsed {len(extracted_elements)} elements ---")
        save_elements_to_excel(extracted_elements, excel_output_path, file_id_for_excel)