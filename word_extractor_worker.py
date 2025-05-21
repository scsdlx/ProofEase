import win32com.client
import os
import pythoncom
import pandas as pd # Though we are not creating Excel, pandas might be useful for other things, or can be removed if not.
import re
import requests # For downloading the file
import shutil # For deleting temporary directories

# Assuming database_operations_worker.py is in the same directory or PYTHONPATH
import database_operations_worker as db_worker

# Constants from original script (can be adjusted or moved)
WD_OUTLINE_LEVEL_BODY_TEXT = 10
EXCEL_CELL_CHAR_LIMIT = 32767  # Still relevant for text cleaning

# Attempt to load constants from win32com, with fallbacks
try:
    win32com.client.gencache.EnsureDispatch("Word.Application")
    WD_OUTLINE_LEVEL_BODY_TEXT = win32com.client.constants.wdOutlineLevelBodyText
    WD_STORY_TYPES = { # For iterating through different parts of the document
        "wdMainTextStory": win32com.client.constants.wdMainTextStory,
        "wdFootnotesStory": win32com.client.constants.wdFootnotesStory,
        "wdEndnotesStory": win32com.client.constants.wdEndnotesStory,
        "wdCommentsStory": win32com.client.constants.wdCommentsStory,
        "wdTextFrameStory": win32com.client.constants.wdTextFrameStory,
        "wdEvenPagesHeaderStory": win32com.client.constants.wdEvenPagesHeaderStory,
        "wdPrimaryHeaderStory": win32com.client.constants.wdPrimaryHeaderStory,
        "wdEvenPagesFooterStory": win32com.client.constants.wdEvenPagesFooterStory,
        "wdPrimaryFooterStory": win32com.client.constants.wdPrimaryFooterStory,
        "wdFirstPageHeaderStory": win32com.client.constants.wdFirstPageHeaderStory,
        "wdFirstPageFooterStory": win32com.client.constants.wdFirstPageFooterStory,
    }
    WD_SHAPE_TYPES = {
        "msoPicture": 13, # win32com.client.constants.msoPicture (if MSO constants are available)
        "msoTextBox": 17, # win32com.client.constants.msoTextBox
        "msoAutoShape": 1, # win32com.client.constants.msoAutoShape
        # Add other mso shape types if needed
    }
    WD_TABLE_BEHAVIOR = {
        "wdWord9TableBehavior": 1 # win32com.client.constants.wdWord9TableBehavior
    }
    WD_SAVE_FORMAT = {
        "wdFormatDocument": 0 # win32com.client.constants.wdFormatDocument
    }
    WD_INLINE_SHAPE_TYPE = {
        "wdInlineShapePicture": 3 # win32com.client.constants.wdInlineShapePicture
    }

except AttributeError as e:
    print(f"[WARNING] Could not load some win32com.client.constants: {e}. Using hardcoded values where available.")
    # Hardcoded fallbacks if constants aren't found (less ideal)
    WD_STORY_TYPES = {
        "wdMainTextStory": 1, "wdFootnotesStory": 2, "wdEndnotesStory": 3,
        "wdCommentsStory": 4, "wdTextFrameStory": 5, "wdEvenPagesHeaderStory": 6,
        "wdPrimaryHeaderStory": 7, "wdEvenPagesFooterStory": 8, "wdPrimaryFooterStory": 9,
        "wdFirstPageHeaderStory": 10, "wdFirstPageFooterStory": 11,
    }
    WD_SHAPE_TYPES = {"msoPicture": 13, "msoTextBox": 17, "msoAutoShape": 1}
    WD_TABLE_BEHAVIOR = {"wdWord9TableBehavior": 1}
    WD_SAVE_FORMAT = {"wdFormatDocument": 0}
    WD_INLINE_SHAPE_TYPE = {"wdInlineShapePicture": 3 }


except Exception as e_gencache:
    print(f"[WARNING] gencache.EnsureDispatch or constant loading failed: {e_gencache}. Some features might not work as expected.")


def clean_text_for_db(text):
    if not isinstance(text, str):
        text = str(text)
    # Remove characters that are generally problematic for XML/DBs, except common ones like tabs/newlines
    cleaned_text = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\U00010000-\U0010FFFF]', '', text)
    # Truncate if necessary, though modern DBs handle large text well, this is from original excel limit
    if len(cleaned_text) > EXCEL_CELL_CHAR_LIMIT: # Re-evaluate if this limit is necessary for DB
        return cleaned_text[:EXCEL_CELL_CHAR_LIMIT] + "... (truncated)"
    return cleaned_text

def save_image_from_clipboard_win32(filepath):
    # This function might need to be adapted if PIL/Pillow is not available or suitable in worker context
    # For now, assuming it's similar to original.
    # Ensure Pillow is a dependency: pip install Pillow
    from PIL import ImageGrab
    try:
        image = ImageGrab.grabclipboard()
        if image:
            if image.mode == 'RGBA' or image.mode == 'P':
                image = image.convert('RGB')
            image.save(filepath, 'PNG')
            print(f"[INFO] Image saved to {filepath} from clipboard.")
            return True
    except Exception as e:
        print(f"[ERROR] Failed to save image from clipboard: {e}")
    return False

def get_page_number_from_range_win32(item_range):
    if item_range:
        try:
            return item_range.Information(win32com.client.constants.wdActiveEndPageNumber)
        except Exception as e:
            # print(f"[DEBUG] Could not get page number: {e}") # Too verbose for normal operation
            pass
    return None # Return None if page number can't be determined

def format_table_for_db(table_content_list_of_lists):
    if not table_content_list_of_lists:
        return ""
    # Convert table data to a simple string representation for DB storage
    # For more structured storage, consider JSON string or separate table relations.
    formatted_rows = []
    for row in table_content_list_of_lists:
        cleaned_row = [clean_text_for_db(str(cell_content)) for cell_content in row]
        formatted_rows.append(" | ".join(cleaned_row))
    return "\n".join(formatted_rows)


def _reconstruct_text_with_note_references_db(owner_range, note_collection_getter, page_level_note_counts, note_type_prefix):
    original_text_with_cr = owner_range.Text
    owner_range_start_offset = owner_range.Start
    notes_in_range_data = []

    try:
        for note_obj_idx, note_obj in enumerate(note_collection_getter(owner_range)):
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
                page_of_reference = get_page_number_from_range_win32(note_obj.Reference)
                notes_in_range_data.append({
                    "ref_start_in_owner": ref_start_in_owner,
                    "ref_end_in_owner": ref_end_in_owner,
                    "page_of_reference": page_of_reference,
                    "original_mark": note_obj.Reference.Text.strip() # Store original mark
                })
            except Exception as e_inner_note:
                print(f"[DEBUG] Inner note processing error: {e_inner_note}")
                continue
    except Exception as e_outer_note:
        print(f"[DEBUG] Outer note collection error: {e_outer_note}")
        return original_text_with_cr.strip(), [] # Return empty list for related_notes

    if not notes_in_range_data:
        return original_text_with_cr.strip(), []

    notes_in_range_data.sort(key=lambda x: x["ref_start_in_owner"])
    
    new_text_parts = []
    related_notes_for_element = [] # To store note definitions related to this text block
    last_pos_in_owner = 0
    
    for note_data in notes_in_range_data:
        if note_data["ref_start_in_owner"] < last_pos_in_owner: continue
        new_text_parts.append(original_text_with_cr[last_pos_in_owner:note_data["ref_start_in_owner"]])
        
        current_page_for_this_note = note_data["page_of_reference"]
        note_id_suffix = "?"
        if current_page_for_this_note is not None:
            page_level_note_counts[current_page_for_this_note] = page_level_note_counts.get(current_page_for_this_note, 0) + 1
            note_id_suffix = str(page_level_note_counts[current_page_for_this_note])
        
        # Use original mark from document if available, otherwise generate one
        # The visual mark in text will be e.g. "[PageX-1]"
        # The actual note content will be stored separately.
        display_mark = f"[{note_type_prefix}{current_page_for_this_note if current_page_for_this_note else '?'}-{note_id_suffix}]"
        new_text_parts.append(display_mark)
        
        # This part is conceptual for linking note content.
        # Actual note content is extracted globally later.
        # related_notes_for_element.append({
        #     "type": f"{note_type_prefix}_reference_in_text",
        #     "reference_mark_text": display_mark, # The mark we inserted
        #     "original_doc_mark": note_data["original_mark"],
        #     "page_number": current_page_for_this_note,
        #     "level": None
        # })

        last_pos_in_owner = note_data["ref_end_in_owner"]

    if last_pos_in_owner < len(original_text_with_cr):
        new_text_parts.append(original_text_with_cr[last_pos_in_owner:])
    
    return "".join(new_text_parts).strip(), related_notes_for_element


def parse_range_content_for_db(doc_range, word_app, file_record_id, image_storage_dir_for_file,
                               img_counter, content_item_counter, page_level_footnote_counts, page_level_endnote_counts,
                               element_prefix=""):
    """
    Parses content from a Word document range (e.g., main content, header, footer, textbox)
    and prepares it for database insertion.
    """
    elements_for_db = []
    processed_table_ids = set() # To avoid processing tables multiple times if they span paragraphs

    if not doc_range:
        return elements_for_db

    try:
        paragraphs_collection = doc_range.Paragraphs
    except Exception as e:
        print(f"[ERROR] Could not access Paragraphs collection in range: {e}")
        return elements_for_db

    for para_idx, para in enumerate(paragraphs_collection):
        try:
            para_range_obj = para.Range
            current_page = get_page_number_from_range_win32(para_range_obj)
            
            # Process footnotes within this paragraph
            # The page_level_footnote_counts is passed in and modified by _reconstruct_text_with_note_references_db
            para_text_with_footnotes, _ = _reconstruct_text_with_note_references_db(
                para_range_obj,
                lambda r: r.Footnotes,
                page_level_footnote_counts,
                "FN" # FN for FootNote
            )
            
            # Process endnotes within this paragraph (if logic requires per-paragraph endnote handling)
            # Typically, endnotes are document-wide, but references can appear anywhere.
            # For now, let's assume we are just marking their references in text like footnotes.
            final_para_text_for_output, _ = _reconstruct_text_with_note_references_db(
                para_range_obj, # This range might need to be the one already processed for footnotes
                lambda r: r.Endnotes, # Or use the para_text_with_footnotes as input if nesting references
                page_level_endnote_counts, # Separate counter for endnotes
                "EN" # EN for EndNote
            )
            # If _reconstruct_text_with_note_references_db was chained, use the text from previous step.
            # This example assumes they are processed independently on the original para_range_obj text.
            # For true chaining, the function would need to accept text and operate on it.
            # For simplicity, let's use the text that has footnote references, then add endnote references.
            # This part needs careful thought on how to handle overlapping footnote/endnote refs in same spot.
            # Current _reconstruct function isn't designed for chaining text inputs.
            # So, we'll use the result from footnote processing for further checks.
            # A more robust way would be to get all note positions first, then reconstruct.
            
            # For now, using para_text_with_footnotes as the base for endnote processing is not what the current
            # _reconstruct function does. It always starts from para_range_obj.Text.
            # Let's just take the text that has footnote marks, and then separately check for endnote marks
            # (though _reconstruct will put its own marks). This is a simplification.
            # The ideal way is to get all note objects (footnotes & endnotes), sort them by position,
            # then iterate through text and insert references.

            final_para_text_for_output = para_text_with_footnotes # Defaulting to text with footnote marks
            # If you also want to mark endnotes:
            # final_para_text_for_output, _ = _reconstruct_text_with_note_references_db(para_range_obj, lambda r: r.Endnotes, page_level_endnote_counts, "EN")
            # This would overwrite footnote marks if not handled carefully.
            # The original script processed footnotes and endnotes separately at the document level for content,
            # and then _reconstruct_text_with_note_references handled inserting marks into paragraphs/cells.
            # We should follow that: parse_word_document_for_db will first extract all footnote/endnote content.
            # Then, when parsing paragraphs/cells here, we call _reconstruct for footnotes and then for endnotes
            # on the *same* range text, using separate counters. This means a piece of text might get two marks
            # if a footnote and endnote ref are in the same original text span.
            # This is what the original script did.

            # Let's re-evaluate: the original script's _reconstruct was called for Footnotes.
            # Endnotes were listed separately but not necessarily marked inline via _reconstruct in the same way.
            # The prompt asks for similar processing.
            # The provided extractWordElement.py *did* call _reconstruct_text_with_note_references for Footnotes.
            # Endnotes were listed as "endnote_marker_info" if found in para_range_obj.Endnotes.Count > 0
            # but not run through _reconstruct.
            # Let's stick to the original script's behavior first for footnotes.

            is_in_table = False
            try:
                is_in_table = para_range_obj.Information(win32com.client.constants.wdWithInTable)
            except Exception: pass

            if is_in_table:
                try:
                    table = para_range_obj.Tables(1)
                    table_id_str = f"{element_prefix}table-{content_item_counter[0]}" # Unique ID for the table element
                    if table.ID not in processed_table_ids: # table.ID might not be unique across ranges
                        table_content_data = []
                        for r_idx, row in enumerate(table.Rows):
                            row_data_cells = []
                            for c_idx, cell in enumerate(row.Cells):
                                cell_range_obj = cell.Range
                                # Process footnotes within this cell, using the main page_level_footnote_counts
                                cell_text_with_footnotes, _ = _reconstruct_text_with_note_references_db(
                                    cell_range_obj, lambda r: r.Footnotes, page_level_footnote_counts, "FN"
                                )
                                # Potentially process endnotes in cell too if needed
                                final_cell_text = cell_text_with_footnotes.strip().replace('\r\x07', '').replace('\x07', '') # Clean cell endings
                                row_data_cells.append(final_cell_text)
                            table_content_data.append(row_data_cells)
                        
                        content_item_counter[0] += 1
                        elements_for_db.append({
                            "file_record_id": file_record_id,
                            "element_type": f"{element_prefix}table",
                            "content_id": table_id_str,
                            "text_content": clean_text_for_db(format_table_for_db(table_content_data)),
                            "level": None,
                            "pageNo": get_page_number_from_range_win32(table.Range) # Page of table start
                        })
                        processed_table_ids.add(table.ID) # Mark this Word table object ID as processed
                except Exception as e_table:
                    print(f"[ERROR] Failed to process table: {e_table}")
                continue # Skip separate paragraph processing for text already in table

            if final_para_text_for_output: # Ensure there's text after footnote processing
                content_item_counter[0] += 1
                para_id_str = f"{element_prefix}paragraph-{content_item_counter[0]}"
                style_name = ""
                try: style_name = para.Style.NameLocal
                except: pass
                
                outline_level_val = WD_OUTLINE_LEVEL_BODY_TEXT # Default
                try: outline_level_val = para.OutlineLevel
                except: pass

                element_type = f"{element_prefix}paragraph"
                heading_level = None
                if 1 <= outline_level_val <= 9: # Is it a heading?
                    element_type = f"{element_prefix}heading"
                    heading_level = outline_level_val
                
                elements_for_db.append({
                    "file_record_id": file_record_id,
                    "element_type": element_type,
                    "content_id": para_id_str,
                    "text_content": clean_text_for_db(final_para_text_for_output),
                    "level": heading_level,
                    "pageNo": current_page
                })

            # Inline Shapes (Images) in paragraph
            if para_range_obj.InlineShapes.Count > 0:
                for i_shape_idx, i_shape in enumerate(para_range_obj.InlineShapes):
                    img_counter[0] += 1
                    content_item_counter[0] += 1
                    inline_image_id = f"{element_prefix}inline_image-{content_item_counter[0]}"
                    shape_page = get_page_number_from_range_win32(i_shape.Range)
                    img_filename = f"inline_img_{file_record_id}_{img_counter[0]}.png"
                    img_filepath = os.path.join(image_storage_dir_for_file, img_filename)
                    saved_successfully = False
                    try:
                        if i_shape.Type == WD_INLINE_SHAPE_TYPE["wdInlineShapePicture"]: # Check if it's a picture
                            # For InlineShape, you might need to select, copy, then paste/save from clipboard
                            i_shape.Select()
                            word_app.Selection.Copy()
                            if save_image_from_clipboard_win32(img_filepath):
                                saved_successfully = True
                            else: # Fallback for linked pictures or other issues
                                # Try to export if OLEFormat available (less common for typical inline images)
                                # Or if i_shape has a LinkFormat and it's a linked picture, get SourceFullName
                                print(f"[DEBUG] Clipboard save failed for inline image {img_counter[0]}. Type: {i_shape.Type}")

                        # Add placeholder even if save failed, to acknowledge its existence
                        elements_for_db.append({
                            "file_record_id": file_record_id,
                            "element_type": f"{element_prefix}inline_image" if saved_successfully else f"{element_prefix}inline_image_extraction_failed",
                            "content_id": inline_image_id,
                            "text_content": clean_text_for_db(img_filepath if saved_successfully else f"Failed to extract inline image {img_counter[0]}"),
                            "level": None,
                            "pageNo": shape_page
                        })
                    except Exception as e_ishape:
                        content_item_counter[0] += 1 # Ensure counter increment even on error
                        elements_for_db.append({
                            "file_record_id": file_record_id,
                            "element_type": f"{element_prefix}inline_image_error",
                            "content_id": f"{element_prefix}inline_image_error-{content_item_counter[0]}",
                            "text_content": clean_text_for_db(f"Error processing inline shape: {e_ishape}"),
                            "level": None, "pageNo": shape_page
                        })
            
            # Endnote markers (just noting their existence, content is global)
            # The original script had a section for this, let's replicate.
            if para_range_obj.Endnotes.Count > 0:
                 for en_ref_obj_idx, en_ref_obj in enumerate(para_range_obj.Endnotes):
                    try:
                        content_item_counter[0] += 1
                        elements_for_db.append({
                            "file_record_id": file_record_id,
                            "element_type": f"{element_prefix}endnote_reference_marker", # More specific type
                            "content_id": f"{element_prefix}endnote_ref-{content_item_counter[0]}",
                            # Text indicates where the reference is, not the endnote content itself.
                            "text_content": clean_text_for_db(f"Endnote Ref Mark: [{en_ref_obj.Reference.Text.strip()}] (Doc Index: {en_ref_obj.Index})"),
                            "level": None,
                            "pageNo": get_page_number_from_range_win32(en_ref_obj.Reference)
                        })
                    except Exception as e_en_marker:
                        print(f"[DEBUG] Error processing endnote marker in para: {e_en_marker}")


        except Exception as e_para:
            print(f"[ERROR] Failed to process paragraph {para_idx}: {e_para}")
            content_item_counter[0] += 1 # Increment to maintain unique IDs on error
            elements_for_db.append({
                "file_record_id": file_record_id,
                "element_type": f"{element_prefix}paragraph_error",
                "content_id": f"{element_prefix}error-{content_item_counter[0]}",
                "text_content": clean_text_for_db(f"Error processing paragraph: {e_para}"),
                "level": None, "pageNo": current_page if 'current_page' in locals() else None
            })
            
    return elements_for_db


def parse_word_document_for_db(file_record_id, temp_doc_path, image_output_dir_base):
    """
    Main parsing function. Extracts content from the Word document and structures it for DB insertion.
    """
    pythoncom.CoInitialize()
    word_app = None
    doc = None
    all_elements_for_db = []
    
    # Counters for generating unique IDs and image names
    img_counter = [0] # Mutable list to pass by reference
    content_item_counter = [0] # Global counter for all content items (para, table, image)

    # Per-page footnote/endnote numbering for references in text (e.g., [FN1-1], [EN1-1])
    # These are modified by _reconstruct_text_with_note_references_db
    page_level_footnote_counts = {} 
    page_level_endnote_counts = {} # If we decide to mark endnotes similarly

    # Create a specific directory for this file's images
    image_storage_dir_for_file = os.path.join(image_output_dir_base, str(file_record_id))
    if not os.path.exists(image_storage_dir_for_file):
        os.makedirs(image_storage_dir_for_file)
    else: # Clean out old images for this file_record_id if any
        shutil.rmtree(image_storage_dir_for_file)
        os.makedirs(image_storage_dir_for_file)

    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False # Run in background
        doc = word_app.Documents.Open(os.path.abspath(temp_doc_path), ReadOnly=True)

        # 1. Extract Global Footnote and Endnote Content Definitions First
        # These are stored separately as their own element_type.
        # Their references in the main text will be marked like [FN_Page1_Item1]
        print(f"[INFO] Extracting Footnote definitions for doc: {file_record_id}")
        if doc.Footnotes.Count > 0:
            for fn_idx, fn_in_doc in enumerate(doc.Footnotes):
                content_item_counter[0] += 1
                fn_id = f"footnote_def-{content_item_counter[0]}"
                try:
                    all_elements_for_db.append({
                        "file_record_id": file_record_id,
                        "element_type": "footnote_definition",
                        "content_id": fn_id,
                        "text_content": clean_text_for_db(f"[{fn_in_doc.Reference.Text.strip()}] {fn_in_doc.Range.Text.strip()}"),
                        "level": None, # Not applicable
                        "pageNo": get_page_number_from_range_win32(fn_in_doc.Range) # Page where footnote text appears
                    })
                except Exception as e_fn_def:
                     print(f"[ERROR] Extracting footnote definition {fn_idx}: {e_fn_def}")


        print(f"[INFO] Extracting Endnote definitions for doc: {file_record_id}")
        if doc.Endnotes.Count > 0:
            for en_idx, en_in_doc in enumerate(doc.Endnotes):
                content_item_counter[0] += 1
                en_id = f"endnote_def-{content_item_counter[0]}"
                try:
                    all_elements_for_db.append({
                        "file_record_id": file_record_id,
                        "element_type": "endnote_definition",
                        "content_id": en_id,
                        "text_content": clean_text_for_db(f"[{en_in_doc.Reference.Text.strip()}] {en_in_doc.Range.Text.strip()}"),
                        "level": None, # Not applicable
                        "pageNo": get_page_number_from_range_win32(en_in_doc.Range) # Page where endnote text appears
                    })
                except Exception as e_en_def:
                    print(f"[ERROR] Extracting endnote definition {en_idx}: {e_en_def}")
        
        # 2. Parse Main Document Content (doc.Content)
        # This will populate elements_for_db with paragraphs, tables, inline images from the main body
        # And it will use page_level_footnote_counts to mark references in text.
        print(f"[INFO] Parsing Main Document Content (doc.Content) for doc: {file_record_id}...")
        all_elements_for_db.extend(
            parse_range_content_for_db(doc.Content, word_app, file_record_id, image_storage_dir_for_file,
                                       img_counter, content_item_counter, page_level_footnote_counts, page_level_endnote_counts,
                                       element_prefix="main_")
        )

        # 3. Parse Headers and Footers
        print(f"[INFO] Parsing Headers and Footers for doc: {file_record_id}...")
        for section_idx, section in enumerate(doc.Sections):
            header_footer_types_map = {
                "Primary": win32com.client.constants.wdHeaderFooterPrimary,
                "FirstPage": win32com.client.constants.wdHeaderFooterFirstPage,
                "EvenPages": win32com.client.constants.wdHeaderFooterEvenPages
            }
            for hf_kind_name, hf_collection_obj_getter in [("Header", section.Headers), ("Footer", section.Footers)]:
                for hf_type_name, hf_type_constant in header_footer_types_map.items():
                    try:
                        hf_object = hf_collection_obj_getter(hf_type_constant)
                        if hf_object.Exists:
                            hf_range = hf_object.Range
                            if hf_range and hf_range.Text.strip(): # Check if there's actual content
                                prefix = f"section{section_idx+1}_{hf_kind_name.lower()}_{hf_type_name.lower()}_"
                                # Each header/footer is a separate "mini-document" in terms of note numbering context if needed
                                # However, standard Word footnotes are usually tied to the main document body pages.
                                # For now, pass the main doc's page_level_footnote_counts.
                                # If headers/footers can have their own independent footnotes (unlikely for standard Word usage),
                                # then new count dictionaries would be needed here.
                                all_elements_for_db.extend(
                                    parse_range_content_for_db(hf_range, word_app, file_record_id, image_storage_dir_for_file,
                                                               img_counter, content_item_counter, page_level_footnote_counts, page_level_endnote_counts,
                                                               element_prefix=prefix)
                                )
                    except Exception as e_hf:
                        print(f"[ERROR] Processing {hf_kind_name} {hf_type_name} for section {section_idx+1}: {e_hf}")
        
        # 4. Parse Floating Shapes (Textboxes, Floating Images) in doc.Shapes
        # These are anchored to the document but not inline with text flow.
        print(f"[INFO] Parsing Floating Shapes (doc.Shapes) for doc: {file_record_id}...")
        for shape_idx, shape in enumerate(doc.Shapes):
            img_counter[0] += 1 # Increment general image counter
            content_item_counter[0] += 1 # Increment general content counter
            
            shape_page = None
            anchor_text_preview = "N/A (Floating)"
            try: 
                if hasattr(shape, "Anchor") and shape.Anchor:
                    shape_page = get_page_number_from_range_win32(shape.Anchor)
            except: pass

            try:
                shape_type_val = shape.Type
                
                if shape_type_val == WD_SHAPE_TYPES.get("msoPicture") or \
                   shape_type_val == WD_SHAPE_TYPES.get("msoLinkedPicture"): # Handle pictures
                    floating_img_id = f"floating_image-{content_item_counter[0]}"
                    img_filename = f"floating_img_{file_record_id}_{img_counter[0]}.png"
                    img_filepath = os.path.join(image_storage_dir_for_file, img_filename)
                    saved_successfully = False
                    try:
                        shape.Select()
                        word_app.Selection.Copy() # Copy the selected shape
                        if save_image_from_clipboard_win32(img_filepath):
                            saved_successfully = True
                        else: # Fallback for linked pictures
                            if shape_type_val == WD_SHAPE_TYPES.get("msoLinkedPicture") and shape.LinkFormat:
                                # This is a linked picture, SourceFullName might be the path
                                # However, direct saving from LinkFormat isn't straightforward with win32com for all cases
                                print(f"[DEBUG] Clipboard save failed for floating linked image {img_counter[0]}. Link source: {shape.LinkFormat.SourceFullName if shape.LinkFormat else 'N/A'}")
                            else:
                                print(f"[DEBUG] Clipboard save failed for floating msoPicture {img_counter[0]}.")

                    except Exception as e_fimg_save:
                        print(f"[ERROR] Saving floating image {img_counter[0]} failed: {e_fimg_save}")
                    
                    all_elements_for_db.append({
                        "file_record_id": file_record_id,
                        "element_type": "floating_image" if saved_successfully else "floating_image_extraction_failed",
                        "content_id": floating_img_id,
                        "text_content": clean_text_for_db(img_filepath if saved_successfully else f"Failed to extract floating image {img_counter[0]}"),
                        "level": None, "pageNo": shape_page
                    })

                elif shape_type_val == WD_SHAPE_TYPES.get("msoTextBox") or \
                     (shape_type_val == WD_SHAPE_TYPES.get("msoAutoShape") and shape.TextFrame and shape.TextFrame.HasText):
                    # It's a textbox or an autoshape with text
                    textbox_id = f"textbox-{content_item_counter[0]}"
                    text_content_of_shape = "[No TextFrame or Text]"
                    text_range_to_parse = None
                    if hasattr(shape, "TextFrame") and shape.TextFrame and hasattr(shape.TextFrame, "TextRange"):
                        try:
                            text_range_to_parse = shape.TextFrame.TextRange
                            text_content_of_shape = text_range_to_parse.Text.strip() # Get raw text for now
                        except Exception as e_tf_text:
                            text_content_of_shape = f"[Error accessing TextRange: {e_tf_text}]"
                    
                    # Add a primary element for the textbox itself
                    all_elements_for_db.append({
                        "file_record_id": file_record_id,
                        "element_type": "textbox" if shape_type_val == WD_SHAPE_TYPES.get("msoTextBox") else "autoshape_with_text",
                        "content_id": textbox_id,
                        "text_content": clean_text_for_db(text_content_of_shape[:255]), # Store preview or full if simple
                        "level": None, "pageNo": shape_page
                    })
                    
                    # Recursively parse content within the textbox if it has complex content
                    if text_range_to_parse and text_content_of_shape: # Check if text_content_of_shape is not empty/error
                         all_elements_for_db.extend(
                            parse_range_content_for_db(text_range_to_parse, word_app, file_record_id, image_storage_dir_for_file,
                                                       img_counter, content_item_counter, page_level_footnote_counts, page_level_endnote_counts,
                                                       element_prefix=f"shape{shape_idx+1}_text_")
                        )
                else: # Other types of floating shapes
                    other_shape_id = f"other_floating_shape-{content_item_counter[0]}"
                    shape_type_name = "unknown_mso_shape"
                    for name, val_const in WD_SHAPE_TYPES.items():
                        if shape_type_val == val_const: shape_type_name = name; break
                    
                    all_elements_for_db.append({
                        "file_record_id": file_record_id,
                        "element_type": "other_floating_shape",
                        "content_id": other_shape_id,
                        "text_content": clean_text_for_db(f"Other floating shape: {shape_type_name}, Name: {shape.Name if hasattr(shape, 'Name') else 'N/A'}"),
                        "level": None, "pageNo": shape_page
                    })

            except Exception as e_shape:
                content_item_counter[0] +=1 # ensure counter increment
                all_elements_for_db.append({
                    "file_record_id": file_record_id,
                    "element_type": "floating_shape_error",
                    "content_id": f"shape_error-{content_item_counter[0]}",
                    "text_content": clean_text_for_db(f"Error processing floating shape {shape_idx}: {e_shape}"),
                    "level": None, "pageNo": shape_page
                })
        
        print(f"[INFO] Completed parsing for doc: {file_record_id}. Total elements: {len(all_elements_for_db)}")

    except Exception as e_main_parsing:
        print(f"[FATAL] Main parsing loop error for {file_record_id} on doc {temp_doc_path}: {e_main_parsing}")
        import traceback
        traceback.print_exc()
        return False, f"Main parsing error: {e_main_parsing}", [] # Return empty list on major failure
    finally:
        if doc:
            try: doc.Close(False) # Close document without saving changes
            except Exception as e_close: print(f"[WARNING] Error closing document: {e_close}")
        if word_app:
            try: word_app.Quit()
            except Exception as e_quit: print(f"[WARNING] Error quitting Word application: {e_quit}")
        pythoncom.CoUninitialize()
    
    return True, f"Successfully parsed {len(all_elements_for_db)} elements.", all_elements_for_db


def download_file(url, destination_folder, file_record_id):
    """Downloads a file from a URL to a local temporary path."""
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    
    # Derive filename from URL or use a generic one
    filename = url.split('/')[-1]
    if not filename: filename = f"{file_record_id}_temp_document.docx" # Fallback filename
    # Add file_record_id to ensure uniqueness if multiple workers download same named file to shared temp
    temp_filename = f"{file_record_id}_{filename}"
    local_filepath = os.path.join(destination_folder, temp_filename)

    try:
        print(f"[INFO] Downloading {url} to {local_filepath}...")
        with requests.get(url, stream=True) as r:
            r.raise_for_status() # Raises HTTPError for bad responses (4XX or 5XX)
            with open(local_filepath, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
        print(f"[INFO] File downloaded successfully: {local_filepath}")
        return local_filepath
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] Failed to download file from {url}. Error: {e}")
        return None
    except Exception as e_file_write:
        print(f"[ERROR] Failed to write downloaded file to {local_filepath}. Error: {e_file_write}")
        return None


def process_word_document_for_worker(file_record_id: str, document_url: str, image_output_dir_base: str, db_config: dict):
    """
    Main worker function: downloads, parses Word doc, and saves extracted content to DB.
    Returns: tuple (success: bool, message: str)
    """
    # 1. Configure DB operations for this worker context
    db_worker.DB_CONFIG = db_config # Pass runtime DB config to worker's DB module

    # 2. Download the file
    temp_download_dir = os.path.join(os.getcwd(), "temp_downloads_worker") # Worker's temp download location
    downloaded_doc_path = download_file(document_url, temp_download_dir, file_record_id)

    if not downloaded_doc_path:
        return False, f"Failed to download document from URL: {document_url}"

    # 3. Parse the document
    # The image_output_dir_base is where subfolders like 'image_output_dir_base/file_record_id/' will be created.
    success_parsing, message_parsing, extracted_elements = parse_word_document_for_db(
        file_record_id,
        downloaded_doc_path,
        image_output_dir_base
    )

    # 4. Clean up downloaded file
    try:
        os.remove(downloaded_doc_path)
        print(f"[INFO] Deleted temporary downloaded file: {downloaded_doc_path}")
    except OSError as e_remove:
        print(f"[WARNING] Could not delete temporary file {downloaded_doc_path}: {e_remove}")

    if not success_parsing:
        return False, f"Parsing failed for {file_record_id}: {message_parsing}"

    if not extracted_elements:
        return False, f"No elements extracted from {file_record_id}, or parsing error occurred early."
        
    # 5. Save to Database
    print(f"[INFO] Attempting to save {len(extracted_elements)} elements to DB for file_record_id: {file_record_id}")
    try:
        db_worker.add_tmp_document_content_batch(file_record_id, extracted_elements)
        message_db = f"Successfully saved {len(extracted_elements)} elements to database."
        print(f"[INFO] {message_db}")
        # The 'message_or_image_path_info' could be more structured if needed,
        # e.g., a JSON string with total elements, path to images if they are served by worker etc.
        # For now, just the count.
        return True, f"Processing complete for {file_record_id}. {message_parsing} {message_db}"
    except Exception as e_db:
        print(f"[FATAL] Database operation failed for {file_record_id}: {e_db}")
        import traceback
        traceback.print_exc()
        # Potentially try to clean up images if DB fails, though they might be useful for retry
        return False, f"Database operation failed: {e_db}"


if __name__ == '__main__':
    # This is an example of how to call the worker function.
    # In a real scenario, this script would be invoked by a separate process/service.
    print("Running word_extractor_worker.py directly (for testing purposes)...")

    # --- Configuration for Direct Test ---
    # WARNING: Hardcoding credentials is not secure for production.
    # These should come from a secure config or environment variables.
    TEST_DB_CONFIG = {
        'host': '124.223.68.89',
        'user': 'root',
        'password': 'Mjhu666777;', # Replace with your actual password
        'database': 'ShenJiao'
    }
    
    # Example: A file_record_id that exists in your `file_records` table
    # and has a valid URL to a .docx file.
    # Replace with a real file_record_id and a publicly accessible .docx URL for testing.
    test_file_record_id = "test-doc-003" # Example ID
    # IMPORTANT: This URL must be a direct link to a .docx file and accessible by the worker.
    # Using a placeholder, replace with a real one for testing.
    # test_document_url = "http://example.com/path/to/your/document.docx"
    # For local testing, you might need to serve a file locally e.g. using `python -m http.server`
    # in a directory with a test doc and then use `http://localhost:8000/your_document.docx`
    # Or, use a file from a known public source if available.
    
    # A local file can be "uploaded" to a local file server for testing the download
    # For example, put "test_document.docx" in a folder, cd to that folder, run:
    # python -m http.server 8000
    # Then the URL would be "http://localhost:8000/test_document.docx" (if worker runs on same machine)
    # If your worker is on a different machine, ensure the server is accessible (0.0.0.0) and firewall allows.
    
    # test_document_url = "https://calibre-ebook.com/downloads/demos/demo.docx" # Example public docx
    test_document_url = "http://127.0.0.1:8000/节选：创建有意识的机器 250422.docx" # Assuming local server running

    # Base directory on the worker where images for each document will be stored
    # e.g., C:\proofease_worker_data\images or /opt/proofease_worker_data/images
    test_image_output_base = os.path.join(os.getcwd(), "worker_image_output")
    if not os.path.exists(test_image_output_base):
        os.makedirs(test_image_output_base)

    print(f"Test parameters:")
    print(f"  File Record ID: {test_file_record_id}")
    print(f"  Document URL: {test_document_url}")
    print(f"  Image Output Base: {test_image_output_base}")
    print(f"  DB Host: {TEST_DB_CONFIG['host']}, DB Name: {TEST_DB_CONFIG['database']}")

    # Make sure mysql.connector is installed: pip install mysql-connector-python
    # Make sure pywin32 is installed: pip install pywin32
    # Make sure requests is installed: pip install requests
    # Make sure Pillow is installed: pip install Pillow

    # --- Run the Test ---
    # Note: Running this directly requires a running MySQL server accessible at TEST_DB_CONFIG
    # and a Word document accessible at test_document_url.
    # Also, this machine must be Windows with MS Word installed.
    
    # Check if a local server is needed/expected for the test URL
    if "127.0.0.1:8000" in test_document_url or "localhost:8000" in test_document_url:
        print("\n[TESTING NOTE] This test uses a local URL (e.g., http://127.0.0.1:8000).")
        print("Ensure you have a local HTTP server running in the directory containing your test .docx file.")
        print("Example: cd to your_docs_folder && python -m http.server 8000")
        input("Press Enter to continue if your local server is running, or Ctrl+C to abort...")


    success, message = process_word_document_for_worker(
        test_file_record_id,
        test_document_url,
        test_image_output_base,
        TEST_DB_CONFIG
    )

    if success:
        print(f"\n[SUCCESS] Test processing finished.")
        print(f"  Message: {message}")
        # Verify by checking the `tmp_document_contents` table for entries with `file_record_id = test_file_record_id`
        # Also check the `test_image_output_base/test_file_record_id/` directory for any extracted images.
    else:
        print(f"\n[FAILURE] Test processing failed.")
        print(f"  Message: {message}")

    print("\nword_extractor_worker.py test run finished.")
    print("Dependencies: requests, pywin32, mysql-connector-python, Pillow, (openpyxl if pandas is used for more than just stubs)")
