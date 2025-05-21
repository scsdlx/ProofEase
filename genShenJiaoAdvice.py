import pandas as pd
import json
import win32com.client
# import win32com.client.constants as wdConstants # REMOVED THIS LINE
import os

# --- LCS 和对齐函数 (保持不变) ---
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

# --- Word颜色常量 (RGB) ---
WD_COLOR_RED = 255
WD_COLOR_DARK_GREEN = 32768 # RGB(0, 128, 0)
WD_COLOR_LIGHT_GREEN_BG = 13421772 # A lighter pastel green: RGB(204, 255, 204). Original was 65280 (pure green)
                                  # Let's try a common light green: RGB(144, 238, 144) -> 9498256
                                  # Using a slightly more desaturated light green: RGB(220, 240, 220) -> 14476252

# --- Word Specific Constants (integer values) ---
WD_COLLAPSE_END = 0
WD_COLOR_INDEX_AUTO = 0     # For Font.ColorIndex (Automatic color)
WD_NO_HIGHLIGHT = 0         # For Shading.BackgroundPatternColorIndex (No background highlight)


# --- Status Translation ---
STATUS_TRANSLATION = {
    "pending": "待修改",
    "accepted": "已接受修改建议",
    "denied": "已拒绝修改建议",
    "accepted-edited": "已手动修改"
}

# --- 主逻辑 ---
def main():
    try:
        wca_df = pd.read_excel("word_content_analysis.xlsx")
        dc_df = pd.read_excel("document_contents.xlsx")
        dcc_df = pd.read_excel("document_content_chunks.xlsx")
    except FileNotFoundError as e:
        print(f"错误：找不到 Excel 文件 - {e}")
        return
    except Exception as e:
        print(f"读取 Excel 文件时出错：{e}")
        return

    wca_paragraphs = wca_df[wca_df['element_type'] == 'paragraph'].copy()
    dc_paragraphs = dc_df[dc_df['element_type'] == 'paragraph'].copy()

    wca_paragraphs['text_content'] = wca_paragraphs['text_content'].fillna('').astype(str)
    dc_paragraphs['text_content'] = dc_paragraphs['text_content'].fillna('').astype(str)
    
    all_suggestions = []
    for _, row in dcc_df.iterrows():
        try:
            if pd.notna(row['ai_content']):
                suggestions_json = row['ai_content']
                # Sometimes the JSON might be a string representation of a list of strings,
                # instead of a list of dicts. Add a check.
                if isinstance(suggestions_json, str):
                    try:
                        suggestions = json.loads(suggestions_json)
                    except json.JSONDecodeError as e_inner:
                        print(f"警告：内部 JSON 解析 'ai_content' 失败 (id: {row.get('id', 'N/A')}) - {e_inner}. 内容: {suggestions_json}")
                        continue # Skip this row if inner parsing fails
                elif isinstance(suggestions_json, (list, dict)): # Already parsed by pandas? Unlikely for complex JSON.
                    suggestions = suggestions_json
                else:
                    print(f"警告：'ai_content' 具有意外类型 (id: {row.get('id', 'N/A')}) - type: {type(suggestions_json)}. 内容: {suggestions_json}")
                    continue

                if isinstance(suggestions, list):
                    for sugg in suggestions:
                        if isinstance(sugg, dict) and '材料id' in sugg:
                             all_suggestions.append(sugg)
            else:
                print(f"警告：在 document_content_chunks.xlsx 中发现空的 'ai_content' (id: {row.get('id', 'N/A')})")
        except json.JSONDecodeError as e: # This handles if json.loads(row['ai_content']) fails directly
            print(f"警告：解析 'ai_content' JSON 失败 (id: {row.get('id', 'N/A')}) - {e}. 内容: {row['ai_content']}")
        except TypeError as e:
            print(f"警告：'ai_content' 类型错误 (id: {row.get('id', 'N/A')}) - {e}. 内容: {row['ai_content']}")

    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        doc = word_app.Documents.Add()
    except Exception as e:
        print(f"初始化 Word 失败: {e}")
        return

    first_suggestion_written = False

    for index, wca_row in wca_paragraphs.iterrows():
        source_text = wca_row['text_content']
        page_no = wca_row['pageNo']

        best_match_content_id = None
        best_match_dc_text = ""
        max_similarity = 0.0

        for _, dc_row in dc_paragraphs.iterrows():
            target_text = dc_row['text_content']
            _, _, similarity = get_alignment_details(source_text, target_text)
            if similarity >= 0.75 and similarity > max_similarity:
                max_similarity = similarity
                best_match_content_id = dc_row['content_id']
                best_match_dc_text = target_text

        if not best_match_content_id:
            continue
        
        relevant_suggestions = [sugg for sugg in all_suggestions if sugg.get('材料id') == best_match_content_id]

        if not relevant_suggestions:
            continue
        
        for sugg_data in relevant_suggestions:
            json_original_content = str(sugg_data.get('原始内容', ''))
            json_modified_content = str(sugg_data.get('修改后内容', ''))
            json_status_raw = str(sugg_data.get('status', 'N/A'))
            json_reason = str(sugg_data.get('出错原因', '无原因说明'))

            translated_status = STATUS_TRANSLATION.get(json_status_raw, json_status_raw)

            if first_suggestion_written:
                hr_para = doc.Paragraphs.Add()
                try:
                    hr_para.Range.InsertHorizontalLine()
                except:
                    hr_para.Range.Text = "------------------------------------------------------------\n"
            else:
                first_suggestion_written = True
            
            para_page = doc.Paragraphs.Add().Range
            para_page.Text = f"页码：{page_no}\n"
            
            para_content_header = doc.Paragraphs.Add().Range
            para_content_header.Text = "原始内容：" 
            
            current_inline_range = doc.Paragraphs.Last.Range
            current_inline_range.Collapse(WD_COLLAPSE_END) # Use defined constant

            original_doc_content_str = str(best_match_dc_text) if pd.notna(best_match_dc_text) else ""

            if json_original_content and json_original_content in original_doc_content_str:
                start_index = original_doc_content_str.find(json_original_content)
                end_index = start_index + len(json_original_content)

                part_before = original_doc_content_str[:start_index]
                part_to_mark = original_doc_content_str[start_index:end_index]
                part_after = original_doc_content_str[end_index:]

                current_inline_range.InsertAfter(part_before)
                current_inline_range.Collapse(WD_COLLAPSE_END)

                current_inline_range.InsertAfter(part_to_mark)
                rng_delete = doc.Range(current_inline_range.End - len(part_to_mark), current_inline_range.End)
                rng_delete.Font.Color = WD_COLOR_RED
                rng_delete.Font.StrikeThrough = True
                current_inline_range.Collapse(WD_COLLAPSE_END)

                current_inline_range.InsertAfter(json_modified_content)
                rng_add = doc.Range(current_inline_range.End - len(json_modified_content), current_inline_range.End)
                rng_add.Font.Color = WD_COLOR_DARK_GREEN
                rng_add.Shading.BackgroundPatternColor = WD_COLOR_LIGHT_GREEN_BG
                rng_add.Font.StrikeThrough = False
                current_inline_range.Collapse(WD_COLLAPSE_END)

                status_text_formatted = f"【{translated_status}】"
                current_inline_range.InsertAfter(status_text_formatted)
                rng_status = doc.Range(current_inline_range.End - len(status_text_formatted), current_inline_range.End)
                rng_status.Font.ColorIndex = WD_COLOR_INDEX_AUTO # Use defined constant
                rng_status.Shading.BackgroundPatternColorIndex = WD_NO_HIGHLIGHT # Use defined constant
                rng_status.Font.StrikeThrough = False
                current_inline_range.Collapse(WD_COLLAPSE_END)
                
                current_inline_range.InsertAfter(part_after)
                # Only apply default formatting if part_after is not empty
                if part_after:
                    rng_part_after = doc.Range(current_inline_range.End - len(part_after), current_inline_range.End)
                    rng_part_after.Font.ColorIndex = WD_COLOR_INDEX_AUTO
                    rng_part_after.Shading.BackgroundPatternColorIndex = WD_NO_HIGHLIGHT
                    rng_part_after.Font.StrikeThrough = False
                current_inline_range.Collapse(WD_COLLAPSE_END)

            else: 
                print(f"警告：JSON中的“原始内容” ('{json_original_content}') 未在文档原始内容 ('{original_doc_content_str[:50]}...') 中找到。将仅附加建议。")
                current_inline_range.InsertAfter(original_doc_content_str)
                current_inline_range.Collapse(WD_COLLAPSE_END)

                current_inline_range.InsertAfter(" （建议修改为：") 
                current_inline_range.Collapse(WD_COLLAPSE_END)
                current_inline_range.InsertAfter(json_modified_content)
                rng_add_fallback = doc.Range(current_inline_range.End - len(json_modified_content), current_inline_range.End)
                rng_add_fallback.Font.Color = WD_COLOR_DARK_GREEN
                rng_add_fallback.Shading.BackgroundPatternColor = WD_COLOR_LIGHT_GREEN_BG
                rng_add_fallback.Font.StrikeThrough = False
                current_inline_range.Collapse(WD_COLLAPSE_END)
                current_inline_range.InsertAfter("）")
                current_inline_range.Collapse(WD_COLLAPSE_END)

                status_text_formatted = f"【{translated_status}】"
                current_inline_range.InsertAfter(status_text_formatted)
                rng_status_fallback = doc.Range(current_inline_range.End - len(status_text_formatted), current_inline_range.End)
                rng_status_fallback.Font.ColorIndex = WD_COLOR_INDEX_AUTO
                rng_status_fallback.Shading.BackgroundPatternColorIndex = WD_NO_HIGHLIGHT
                rng_status_fallback.Font.StrikeThrough = False
                current_inline_range.Collapse(WD_COLLAPSE_END)
            
            current_inline_range.InsertAfter("\n")

            para_reason = doc.Paragraphs.Add().Range 
            para_reason.Text = f"原因：{json_reason}\n"


    output_filename = "审校建议清单_v3.docx" 
    full_output_path = os.path.abspath(output_filename)
    try:
        doc.SaveAs(full_output_path)
        print(f"审校建议清单已保存到: {full_output_path}")
    except Exception as e:
        print(f"保存 Word 文档失败: {e}")
    finally:
        if 'doc' in locals() and doc:
            doc.Close(False)
        if 'word_app' in locals() and word_app:
            word_app.Quit()
        # It's good practice to release COM objects, though Python's GC usually handles it.
        # Make sure they exist before trying to delete if an early error occurred.
        if 'doc' in locals(): del doc
        if 'word_app' in locals(): del word_app

if __name__ == "__main__":
    main()