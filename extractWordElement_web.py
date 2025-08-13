# extractWordElement_web.py
import win32com.client
import os
# from PIL import ImageGrab # ImageGrab 不再需要，因为我们不提取图片了
import pythoncom
# import pandas as pd # Pandas 不再在此脚本中使用
import re
import logging # 用于更好的日志记录
import mysql.connector # 导入 mysql.connector 以便在此文件内创建连接
from db_config import DB_CONFIG # 导入数据库配置

# 默认为 INFO 级别。开发时可以改为 logging.DEBUG 查看详细日志。
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - [extractWordElement_web] - %(message)s')

# --- 常量 ---
WD_OUTLINE_LEVEL_BODY_TEXT = 10 # 正文文本的大纲级别

try:
    win32com.client.gencache.EnsureDispatch("Word.Application") # 确保 Word 类型库已生成
    WD_OUTLINE_LEVEL_BODY_TEXT = win32com.client.constants.wdOutlineLevelBodyText
    WD_CONSTANTS = win32com.client.constants # Word 常量对象
except AttributeError:
    logging.warning("[extractWordElement_web] win32com.client.constants 不完全可用，必要时使用硬编码值。")
    # 定义一个虚拟常量类作为后备
    class DummyConstants:
        wdOutlineLevelBodyText = 10
        wdActiveEndPageNumber = 3 # 获取范围所在页码的常量
        wdWithInTable = 12        # 判断范围是否在表格内的常量
    WD_CONSTANTS = DummyConstants()
except Exception as e_gencache:
    logging.warning(f"[extractWordElement_web] gencache.EnsureDispatch 或常量加载失败: {e_gencache}")
    # 再次定义虚拟常量类作为最终后备
    class DummyConstants:
        wdOutlineLevelBodyText = 10; wdActiveEndPageNumber = 3; wdWithInTable = 12
    WD_CONSTANTS = DummyConstants()


# --- 辅助函数 ---
def clean_text_for_db(text):
    """清理文本，移除无效的XML字符以便存入数据库。"""
    if not isinstance(text, str): text = str(text) # 确保是字符串
    cleaned_text = re.sub(r'[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\U00010000-\U0010FFFF]', '', text)
    return cleaned_text

def get_page_number_from_range(item_range):
    """从 Word Range 对象获取页码。"""
    if item_range:
        try:
            return item_range.Information(WD_CONSTANTS.wdActiveEndPageNumber)
        except Exception as e:
            logging.warning(f"无法获取页码: {e}")
    return None

def format_table_for_db(table_content_list_of_lists):
    """将表格内容（列表的列表）格式化为适合数据库存储的字符串。"""
    if not table_content_list_of_lists: return ""
    formatted_rows = []
    for row in table_content_list_of_lists:
        cleaned_row = [clean_text_for_db(str(cell_content)) for cell_content in row]
        formatted_rows.append(" | ".join(cleaned_row))
    return "\n".join(formatted_rows)


def _reconstruct_text_with_note_references(owner_range, note_collection_getter, page_level_note_counts):
    """
    重构给定范围的文本，将脚注/尾注引用（如[1], [2]）插入文本中。
    """
    original_text_with_cr = owner_range.Text
    owner_range_start_offset = owner_range.Start
    notes_in_range_data = []

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
                    logging.warning(f"注释引用位置越界: "
                                    f"起始: {ref_start_in_owner}, 结束: {ref_end_in_owner}, 文本长度: {len(original_text_with_cr)}")
                    continue

                page_of_reference = get_page_number_from_range(note_obj.Reference)
                notes_in_range_data.append({
                    "ref_start_in_owner": ref_start_in_owner,
                    "ref_end_in_owner": ref_end_in_owner,
                    "page_of_reference": page_of_reference,
                })
            except Exception as e_note_inner:
                logging.warning(f"处理单个注释时出错: {e_note_inner}")
    except Exception as e_note_coll:
        logging.error(f"访问注释集合时出错: {e_note_coll}")
        return original_text_with_cr.strip().replace('\r', '\n').replace('\x07', '')

    if not notes_in_range_data:
        return original_text_with_cr.strip().replace('\r', '\n').replace('\x07', '')

    notes_in_range_data.sort(key=lambda x: x["ref_start_in_owner"])

    new_text_parts = []; last_pos_in_owner = 0
    for note_data in notes_in_range_data:
        if note_data["ref_start_in_owner"] < last_pos_in_owner:
            continue
        if note_data["ref_start_in_owner"] > len(original_text_with_cr):
            break

        new_text_parts.append(original_text_with_cr[last_pos_in_owner : note_data["ref_start_in_owner"]])

        current_page_for_this_note = note_data["page_of_reference"]; mark_to_display = "?"
        if current_page_for_this_note is not None:
            page_level_note_counts[current_page_for_this_note] = page_level_note_counts.get(current_page_for_this_note, 0) + 1
            mark_to_display = str(page_level_note_counts[current_page_for_this_note])
        new_text_parts.append(f"[{mark_to_display}]")

        last_pos_in_owner = note_data["ref_end_in_owner"]
        if last_pos_in_owner > len(original_text_with_cr):
            last_pos_in_owner = len(original_text_with_cr)

    if last_pos_in_owner < len(original_text_with_cr):
        new_text_parts.append(original_text_with_cr[last_pos_in_owner:])

    return "".join(new_text_parts).strip().replace('\r', '\n').replace('\x07', '')


def parse_range_content(doc_range, output_elements,
                        page_local_footnote_counts,
                        element_prefix=""):
    """
    解析给定 Word Range 中的内容，提取段落、标题和表格。
    """
    processed_table_ids = set()
    if not doc_range: return

    try:
        paragraphs_collection = doc_range.Paragraphs
    except Exception as e:
        logging.error(f"无法从此范围获取 Paragraphs 集合: {e}")
        return

    for para_idx, para in enumerate(paragraphs_collection):
        try:
            para_range_obj = para.Range

            final_para_text_for_output = _reconstruct_text_with_note_references(
                para_range_obj, lambda r: r.Footnotes, page_local_footnote_counts
            )

            current_page = get_page_number_from_range(para_range_obj); is_in_table = False
            try:
                is_in_table = para_range_obj.Information(WD_CONSTANTS.wdWithInTable)
            except Exception as e:
                logging.warning(f"检查 IsInTable 时出错: {e}")

            if is_in_table:
                try:
                    table = para_range_obj.Tables(1)
                    if table.ID not in processed_table_ids:
                        table_data = []
                        for r_idx in range(1, table.Rows.Count + 1):
                            row_data_cells = []; row = table.Rows(r_idx)
                            for c_idx in range(1, row.Cells.Count + 1):
                                cell = row.Cells(c_idx); cell_range_obj = cell.Range
                                final_cell_text_intermediate = _reconstruct_text_with_note_references(
                                    cell_range_obj, lambda r: r.Footnotes, page_local_footnote_counts
                                )
                                final_cell_text = final_cell_text_intermediate.strip().replace('\r\x07', '').replace('\x07', '').replace('\r', '\n')
                                row_data_cells.append(final_cell_text)
                            table_data.append(row_data_cells)
                        output_elements.append({
                            "type": f"{element_prefix}table", "id": table.ID,
                            "content_data": table_data, "rows": table.Rows.Count,
                            "columns": table.Columns.Count,
                            "page_number": get_page_number_from_range(table.Range), "level": None
                        })
                        processed_table_ids.add(table.ID)
                except Exception as e:
                    logging.error(f"处理表格时出错: {e}", exc_info=True)
                continue

            if final_para_text_for_output:
                style_name = "Normal";
                try: style_name = para.Style.NameLocal
                except Exception: pass # nosec B110
                outline_level_val = WD_CONSTANTS.wdOutlineLevelBodyText
                try: outline_level_val = para.OutlineLevel
                except Exception: pass # nosec B110
                element_data = {"text": final_para_text_for_output, "style": style_name,
                                "page_number": current_page, "level": None}
                if 1 <= outline_level_val <= 9:
                    element_data["type"] = f"{element_prefix}heading"; element_data["level"] = int(outline_level_val)
                else: element_data["type"] = f"{element_prefix}paragraph"
                output_elements.append(element_data)

        except pythoncom.com_error as e_com:
            logging.error(f"处理段落 {para_idx+1} 时发生COM错误: {e_com}", exc_info=True)
        except Exception as e_para:
            logging.error(f"处理段落 {para_idx+1} 时发生一般错误: {e_para}", exc_info=True)


def parse_word_document_to_elements(doc_path, word_app, image_output_dir_unused):
    """
    解析Word文档，提取段落、标题和表格元素。
    """
    doc = None
    output_elements = []
    main_content_footnotes_counts = {}

    try:
        abs_doc_path = os.path.abspath(doc_path)
        logging.debug(f"正在打开文档: {abs_doc_path}")
        doc = word_app.Documents.Open(abs_doc_path, ReadOnly=True, AddToRecentFiles=False)

        logging.debug("正在解析主文档内容 (doc.Content)...")
        parse_range_content(doc.Content, output_elements,
                            main_content_footnotes_counts)

    except pythoncom.com_error as e_com_main:
        logging.error(f"Word处理期间发生主COM错误: {e_com_main}", exc_info=True); raise
    except Exception as e_main:
        logging.error(f"Word处理期间发生主错误: {e_main}", exc_info=True); raise
    finally:
        if doc:
            try: doc.Close(False); logging.debug("已关闭Word文档。")
            except Exception as e_close: logging.error(f"关闭文档时出错: {e_close}")
    return output_elements


def save_elements_to_db(db_cursor, elements_list, file_record_id):
    """将提取的元素列表保存到数据库的 document_contents 表。"""
    if not elements_list:
        logging.info("没有元素需要保存到数据库。")
        return

    try:
        sql_delete = "DELETE FROM document_contents WHERE file_record_id = %s"
        db_cursor.execute(sql_delete, (file_record_id,))
        logging.debug(f"已清除 file_record_id: {file_record_id} 的现有 document_contents 记录。受影响行数: {db_cursor.rowcount}")
    except Exception as e_delete:
        logging.error(f"删除旧 document_contents 记录时出错: {e_delete}")
        # 根据错误处理策略，可能需要在这里回滚或抛出异常
        # 由于此函数现在由 run_extraction 调用，其中有事务处理，这里只记录并允许其传播
        raise

    sql = """
        INSERT INTO document_contents
        (file_record_id, element_type, content_id, text_content, sequence_order, level, page_no)
        VALUES (%s, %s, %s, %s, %s, %s, %s)
    """
    data_to_insert = []
    for elem_idx, elem in enumerate(elements_list):
        content_item_id = str(elem_idx + 1); elem_type = elem.get("type", "unknown")[:20]
        text_content_raw = ""
        if "table" in elem_type and "content_data" in elem:
            text_content_raw = format_table_for_db(elem.get("content_data", []))
        elif "text" in elem:
            text_content_raw = elem.get("text", "")
        final_text_content = clean_text_for_db(text_content_raw)
        level_val = elem.get("level"); page_no_val = elem.get("page_number")
        try: level_db = int(level_val) if level_val is not None else None
        except ValueError:
            logging.warning(f"元素 {content_item_id} 的级别值 '{level_val}' 无效。设置为 NULL。"); level_db = None
        try: page_no_db = int(page_no_val) if page_no_val is not None else None
        except ValueError:
            logging.warning(f"元素 {content_item_id} 的页码值 '{page_no_val}' 无效。设置为 NULL。"); page_no_db = None
        row_tuple = (file_record_id, elem_type, content_item_id, final_text_content, elem_idx + 1, level_db, page_no_db)
        data_to_insert.append(row_tuple)
    try:
        db_cursor.executemany(sql, data_to_insert) # 批量插入数据
        logging.info(f"成功向 document_contents 插入 {db_cursor.rowcount} 行数据。")
    except Exception as e:
        logging.error(f"向 document_contents 插入数据时出错: {e}", exc_info=True)
        raise


def run_extraction(doc_path, file_record_id, image_dir_unused):
    """
    运行Word文档内容提取的主函数。
    数据库连接在此函数内部建立和关闭。
    """
    word_app = None
    coinitialized = False
    db_conn_local = None  # 本地数据库连接
    db_cursor_local = None # 本地数据库游标
    try:
        pythoncom.CoInitialize()
        coinitialized = True
        logging.debug("COM 已初始化。")
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        word_app.DisplayAlerts = 0

        # 解析Word文档元素 (不涉及数据库连接)
        extracted_elements = parse_word_document_to_elements(doc_path, word_app, image_dir_unused)
        logging.info(f"从Word文档解析了 {len(extracted_elements)} 个元素 (文件ID: {file_record_id})。")

        # 在数据库操作前建立连接
        logging.debug(f"为文件ID {file_record_id} 的数据库操作建立本地连接...")
        db_conn_local = mysql.connector.connect(**DB_CONFIG)
        db_cursor_local = db_conn_local.cursor()
        logging.debug(f"文件ID {file_record_id} 的本地数据库连接已建立。")

        # 将提取的元素保存到数据库 (使用本地连接)
        save_elements_to_db(db_cursor_local, extracted_elements, file_record_id)
        db_conn_local.commit() # 提交数据库事务
        logging.info(f"{file_record_id} 的数据已提交到 document_contents。")

    except Exception as e:
        logging.error(f"{file_record_id} 的 run_extraction 过程中出错: {e}", exc_info=True)
        if db_conn_local and db_conn_local.is_connected(): # 检查连接是否已建立且仍连接
            try:
                db_conn_local.rollback() #发生错误时回滚事务
                logging.info(f"由于错误，文件ID {file_record_id} 的数据库事务已回滚。")
            except Exception as e_rollback:
                # 即使回滚失败（例如，如果连接在错误发生时已经断开），也要记录
                logging.error(f"文件ID {file_record_id} 回滚时出错: {e_rollback}")
        raise # 将异常向上层抛出
    finally:
        # 清理数据库资源 (本地连接)
        if db_cursor_local:
            try:
                db_cursor_local.close()
            except Exception as e_cur_close:
                logging.error(f"关闭本地数据库游标时出错 (文件ID: {file_record_id}): {e_cur_close}")
        if db_conn_local and db_conn_local.is_connected():
            try:
                db_conn_local.close()
                logging.debug(f"文件ID {file_record_id} 的本地数据库连接已关闭。")
            except Exception as e_conn_close:
                logging.error(f"关闭本地数据库连接时出错 (文件ID: {file_record_id}): {e_conn_close}")

        # 清理COM和Word资源
        if word_app:
            try:
                word_app.Quit(0)
                logging.debug(f"Word应用程序已退出 (文件ID: {file_record_id})。")
            except Exception as e_quit:
                logging.error(f"退出Word时出错 (文件ID: {file_record_id}): {e_quit}")
            del word_app
        if coinitialized:
            try:
                pythoncom.CoUninitialize()
                logging.debug(f"COM 已反初始化 (文件ID: {file_record_id})。")
            except Exception as e_com_uninit:
                 logging.error(f"COM反初始化时出错 (文件ID: {file_record_id}): {e_com_uninit}")


if __name__ == "__main__":
    # logging.getLogger().setLevel(logging.DEBUG)
    print("此脚本旨在通过Flask应用导入并运行。")
    # ... (独立测试代码示例保持不变，但现在 run_extraction 不再需要 db_conn, db_cursor 参数)