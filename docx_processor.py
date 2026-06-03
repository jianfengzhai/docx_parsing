# docx解析方案: 以XML的思路来进行docx文档的解析
import zipfile
import re
import os
import json
import logging
import argparse
from datetime import datetime
from bs4 import BeautifulSoup

logging.basicConfig(level=logging.WARNING, format='%(levelname)s: %(message)s')


class EnterpriseDocxProcessor:
    def __init__(self, file_path, output_dir, enable_ocr=False, review_mode='final'):
        self.document_path = file_path
        self.output_directory = output_dir
        self.ocr_enabled = enable_ocr
        self.review_processing_mode = review_mode  # 'final', 'original', 'both'
        self._numbering_counters = {}

    def log_warning(self, message):
        logging.warning(message)

    def extract_comprehensive_content(self):
        """综合内容提取的主入口"""
        with zipfile.ZipFile(self.document_path, 'r') as archive:
            # 核心处理流程
            extraction_result = {
                'basic_content': self.extract_basic_elements(archive),
                'heading_numbers': self.extract_heading_numbering(archive),
                'heading_styles': self.extract_heading_hierarchy(archive),
                'review_content': self.process_review_changes(archive)
            }

        return self.merge_extraction_results(extraction_result)

    # 文本内容的精确提取
    def extract_basic_elements(self, archive):
        try:
            document_xml = archive.read('word/document.xml')
        except KeyError:
            self.log_warning("缺少 word/document.xml")
            return {'paragraphs': [], 'tables': [], 'images': {}}

        paragraphs = self.extract_text_content(document_xml)
        tables = self.extract_table_structures(document_xml)

        images = {}
        try:
            rels_path = 'word/_rels/document.xml.rels'
            if rels_path in archive.namelist():
                relations_xml = archive.read(rels_path)
                images = self.process_image_extraction(archive, relations_xml)
        except Exception as e:
            self.log_warning(f"图片提取失败: {e}")

        return {
            'paragraphs': paragraphs,
            'tables': tables,
            'images': images
        }

    def extract_text_content(self, document_xml):
        """提取文本内容，保持段落结构"""
        content_elements = []
        xml_parser = BeautifulSoup(document_xml, 'xml')

        for paragraph in xml_parser.find_all('w:p'):
            paragraph_text = self.extract_paragraph_text(paragraph)
            if paragraph_text.strip():
                content_elements.append({
                    'type': 'paragraph',
                    'content': paragraph_text,
                    'style_info': self.extract_paragraph_style(paragraph)
                })

        return content_elements

    def extract_paragraph_text(self, paragraph_element):
        """从段落中提取纯文本，处理各种文本节点"""
        text_parts = []

        for run_element in paragraph_element.find_all('w:r'):
            # 处理普通文本
            text_nodes = run_element.find_all('w:t')
            for text_node in text_nodes:
                text_parts.append(text_node.get_text())

            # 处理特殊字符（制表符、换行符等）
            if run_element.find('w:tab'):
                text_parts.append('\t')
            if run_element.find('w:br'):
                text_parts.append('\n')

        return ''.join(text_parts)

    def extract_paragraph_style(self, paragraph):
        style_info = {}
        pPr = paragraph.find('w:pPr')
        if not pPr:
            return style_info

        pStyle = pPr.find('w:pStyle')
        if pStyle:
            style_info['style_id'] = pStyle.get('w:val', '')

        jc = pPr.find('w:jc')
        if jc:
            style_info['alignment'] = jc.get('w:val', '')

        ind = pPr.find('w:ind')
        if ind:
            style_info['indent'] = {
                'left': ind.get('w:left', ''),
                'right': ind.get('w:right', ''),
                'first_line': ind.get('w:firstLine', ''),
                'hanging': ind.get('w:hanging', '')
            }

        spacing = pPr.find('w:spacing')
        if spacing:
            style_info['spacing'] = {
                'before': spacing.get('w:before', ''),
                'after': spacing.get('w:after', ''),
                'line': spacing.get('w:line', ''),
                'line_rule': spacing.get('w:lineRule', '')
            }

        outlineLvl = pPr.find('w:outlineLvl')
        if outlineLvl:
            style_info['outline_level'] = int(outlineLvl.get('w:val', '0'))

        return style_info

    def extract_table_structures(self, document_xml):
        """提取表格，保持完整的行列结构"""
        tables = []
        xml_parser = BeautifulSoup(document_xml, 'xml')

        for table_element in xml_parser.find_all('w:tbl'):
            table_data = {
                'type': 'table',
                'rows': [],
                'properties': self.extract_table_properties(table_element)
            }

            for row_element in table_element.find_all('w:tr'):
                row_data = []
                for cell_element in row_element.find_all('w:tc'):
                    cell_content = self.extract_cell_content(cell_element)
                    cell_properties = self.extract_cell_properties(cell_element)
                    row_data.append({
                        'content': cell_content,
                        'properties': cell_properties
                    })
                table_data['rows'].append(row_data)

            tables.append(table_data)

        return tables

    def extract_table_properties(self, table_element):
        properties = {}
        tblPr = table_element.find('w:tblPr')
        if not tblPr:
            return properties

        tblStyle = tblPr.find('w:tblStyle')
        if tblStyle:
            properties['style'] = tblStyle.get('w:val', '')

        tblW = tblPr.find('w:tblW')
        if tblW:
            properties['width'] = {
                'value': tblW.get('w:w', ''),
                'type': tblW.get('w:type', '')
            }

        jc = tblPr.find('w:jc')
        if jc:
            properties['alignment'] = jc.get('w:val', '')

        tblBorders = tblPr.find('w:tblBorders')
        if tblBorders:
            borders = {}
            for border_tag in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = tblBorders.find(f'w:{border_tag}')
                if border:
                    borders[border_tag] = {
                        'val': border.get('w:val', ''),
                        'sz': border.get('w:sz', ''),
                        'color': border.get('w:color', '')
                    }
            properties['borders'] = borders

        return properties

    def extract_cell_content(self, cell_element):
        text_parts = []
        for paragraph in cell_element.find_all('w:p'):
            para_text = self.extract_paragraph_text(paragraph)
            if para_text.strip():
                text_parts.append(para_text)
        return '\n'.join(text_parts)

    def extract_cell_properties(self, cell_element):
        properties = {}
        tcPr = cell_element.find('w:tcPr')
        if not tcPr:
            return properties

        gridSpan = tcPr.find('w:gridSpan')
        if gridSpan:
            properties['col_span'] = int(gridSpan.get('w:val', '1'))

        vMerge = tcPr.find('w:vMerge')
        if vMerge:
            properties['vertical_merge'] = vMerge.get('w:val', 'continue')

        tcW = tcPr.find('w:tcW')
        if tcW:
            properties['width'] = {
                'value': tcW.get('w:w', ''),
                'type': tcW.get('w:type', '')
            }

        vAlign = tcPr.find('w:vAlign')
        if vAlign:
            properties['vertical_alignment'] = vAlign.get('w:val', '')

        return properties

    def process_image_extraction(self, archive, relations_xml):
        """图片提取与OCR处理"""
        image_registry = {}

        if not self.ocr_enabled:
            return {}  # 如果不需要OCR，直接返回空

        relations_parser = BeautifulSoup(relations_xml, 'xml')

        for relation in relations_parser.find_all('Relationship'):
            if 'image' in relation.get('Type', '').lower():
                image_id = relation.get('Id')
                image_target = relation.get('Target')

                # 提取并保存图片
                image_path = self.extract_image_file(archive, image_target, image_id)
                if image_path:
                    # 处理特殊格式转换
                    processed_path = self.handle_special_image_formats(image_path)
                    image_registry[image_id] = processed_path

        # 批量OCR处理
        return self.batch_ocr_processing(image_registry)

    def extract_image_file(self, archive, target, image_id):
        if target.startswith('/'):
            image_internal_path = target.lstrip('/')
        else:
            image_internal_path = f'word/{target}'

        try:
            image_data = archive.read(image_internal_path)
        except KeyError:
            try:
                alt_path = target.replace('word/', '')
                image_data = archive.read(f'word/{alt_path}')
            except KeyError:
                self.log_warning(f"无法找到图片文件: {image_internal_path}")
                return None

        image_dir = os.path.join(self.output_directory, 'images')
        os.makedirs(image_dir, exist_ok=True)

        filename = os.path.basename(image_internal_path)
        save_path = os.path.join(image_dir, f'{image_id}_{filename}')

        with open(save_path, 'wb') as f:
            f.write(image_data)

        return save_path

    def handle_special_image_formats(self, image_path):
        ext = os.path.splitext(image_path)[1].lower()
        if ext in ('.emf', '.wmf'):
            self.log_warning(f"图片格式 {ext} 可能需要额外转换工具，当前保留原路径: {image_path}")
        return image_path

    def batch_ocr_processing(self, image_registry):
        if not image_registry:
            return {}

        ocr_results = {}
        for image_id, image_path in image_registry.items():
            try:
                ocr_results[image_id] = self._perform_ocr(image_path)
            except Exception as e:
                self.log_warning(f"OCR处理失败 [{image_id}]: {e}")
                ocr_results[image_id] = ''

        return ocr_results

    def _perform_ocr(self, image_path):
        try:
            import pytesseract
            from PIL import Image
            img = Image.open(image_path)
            return pytesseract.image_to_string(img, lang='chi_sim+eng')
        except ImportError:
            self.log_warning("pytesseract 或 Pillow 未安装，跳过OCR处理")
            return ''
        except Exception as e:
            self.log_warning(f"OCR识别失败: {e}")
            return ''

    def extract_heading_numbering(self, archive):
        """提取标题编号系统

        编号来源有两种：
        1. 自动编号：通过 numbering.xml 中的编号定义 + styles.xml 中的 numId 关联计算
        2. 手动编号：标题文本中直接包含的编号（如 "6.8 替代供应商逻辑更新"），
           当样式没有关联自动编号时，通过正则从文本中提取

        样式与编号的关联也有两种路径：
        - 正向关联：styles.xml 中样式定义的 w:pPr/w:numId 直接引用编号ID
        - 反向关联：numbering.xml 中 w:lvl/w:pStyle 引用样式ID（部分文档只有这种关联方式）
        """
        try:
            numbering_xml = archive.read('word/numbering.xml')
            styles_xml = archive.read('word/styles.xml')
            document_xml = archive.read('word/document.xml')
        except KeyError as e:
            self.log_warning(f"缺少必要文件: {e}")
            return {}

        # 解析编号定义，同时提取 pStyle 反向映射（numbering.xml 中 w:lvl/w:pStyle -> numId）
        numbering_definitions, pstyle_to_num = self.parse_numbering_definitions(numbering_xml)

        # 解析样式与编号的关联，合并正向关联（styles.xml 中的 numId）和反向关联（pstyle_to_num）
        style_numbering_map = self.parse_style_numbering_relationships(
            styles_xml, numbering_definitions, pstyle_to_num
        )

        # 解析标题样式定义，用于回退到文本正则提取编号
        heading_styles = self.parse_heading_style_definitions(styles_xml)

        # 提取实际标题编号：优先使用自动编号，回退到文本正则提取
        heading_numbers = self.extract_actual_heading_numbers(
            document_xml, numbering_definitions, style_numbering_map, heading_styles
        )

        return heading_numbers

    def parse_numbering_definitions(self, numbering_xml):
        """解析 numbering.xml 中的编号定义

        返回:
            definitions: {num_id: {ilvl: {format, start_value, text_pattern, ...}}}
            pstyle_to_num: {style_id: {num_id, start_level, levels}}
                从 numbering.xml 中 w:lvl/w:pStyle 反向映射到编号定义。
                部分文档（如飞书导出的docx）不在 styles.xml 中设置 numId，
                而是在 numbering.xml 的 w:lvl 中通过 w:pStyle 引用样式ID，
                需要通过此反向映射建立样式与编号的关联。
        """
        definitions = {}
        pstyle_to_num = {}
        xml_parser = BeautifulSoup(numbering_xml, 'xml')

        abstract_nums = {}
        abstract_pstyles = {}
        for abstract_num in xml_parser.find_all('w:abstractNum'):
            abstract_id = abstract_num.get('w:abstractNumId')
            levels = {}
            for lvl in abstract_num.find_all('w:lvl'):
                ilvl = int(lvl.get('w:ilvl', '0'))
                level_info = {}

                numFmt = lvl.find('w:numFmt')
                if numFmt:
                    level_info['format'] = numFmt.get('w:val', 'decimal')

                start = lvl.find('w:start')
                if start:
                    level_info['start_value'] = int(start.get('w:val', '1'))

                lvlText = lvl.find('w:lvlText')
                if lvlText:
                    level_info['text_pattern'] = lvlText.get('w:val', '%1')

                lvlJc = lvl.find('w:lvlJc')
                if lvlJc:
                    level_info['alignment'] = lvlJc.get('w:val', 'left')

                isLgl = lvl.find('w:isLgl')
                if isLgl and isLgl.get('w:val', '0') == '1':
                    level_info['is_legal'] = True

                # 提取 w:lvl 中的 w:pStyle，建立反向映射
                # 例: <w:lvl w:ilvl="0"><w:pStyle w:val="3"/> 表示样式ID="3" 使用此编号的0级
                pStyle = lvl.find('w:pStyle')
                if pStyle:
                    pstyle_val = pStyle.get('w:val', '')
                    if pstyle_val:
                        if abstract_id not in abstract_pstyles:
                            abstract_pstyles[abstract_id] = []
                        abstract_pstyles[abstract_id].append({
                            'style_id': pstyle_val,
                            'ilvl': ilvl
                        })

                levels[ilvl] = level_info

            abstract_nums[abstract_id] = levels

        # 建立 abstractNumId -> numId 的映射，用于将 pStyle 反向映射到具体的 numId
        abstract_to_num = {}
        for num in xml_parser.find_all('w:num'):
            num_id = num.get('w:numId')
            abstract_ref = num.find('w:abstractNumId')
            if abstract_ref:
                abstract_id = abstract_ref.get('w:val')
                if abstract_id in abstract_nums:
                    definitions[num_id] = abstract_nums[abstract_id]
                    abstract_to_num[abstract_id] = num_id

        # 将 pStyle 引用转换为 style_id -> num_info 的直接映射
        for abstract_id, pstyles in abstract_pstyles.items():
            num_id = abstract_to_num.get(abstract_id)
            if num_id and num_id in definitions:
                for ps in pstyles:
                    pstyle_to_num[ps['style_id']] = {
                        'num_id': num_id,
                        'start_level': ps['ilvl'],
                        'levels': definitions[num_id]
                    }

        return definitions, pstyle_to_num

    def parse_style_numbering_relationships(self, styles_xml, numbering_definitions, pstyle_to_num=None):
        """解析样式与编号的关联关系

        合并两种关联来源：
        1. 正向关联：styles.xml 中样式定义的 w:pPr/w:numId 直接引用编号ID
        2. 反向关联：pstyle_to_num 中从 numbering.xml 的 w:lvl/w:pStyle 提取的映射

        正向关联优先，反向关联仅补充正向关联中不存在的样式。
        """
        style_numbering_map = {}
        xml_parser = BeautifulSoup(styles_xml, 'xml')

        for style in xml_parser.find_all('w:style'):
            style_id = style.get('w:styleId')
            pPr = style.find('w:pPr')
            if not pPr:
                continue

            numId_elem = pPr.find('w:numId')
            if numId_elem:
                num_id = numId_elem.get('w:val')
                if num_id and num_id in numbering_definitions:
                    ilvl_elem = pPr.find('w:ilvl')
                    start_level = int(ilvl_elem.get('w:val', '0')) if ilvl_elem else 0
                    style_numbering_map[style_id] = {
                        'num_id': num_id,
                        'start_level': start_level,
                        'levels': numbering_definitions[num_id]
                    }

        # 合并反向关联：仅补充正向关联中不存在的样式，正向关联优先
        if pstyle_to_num:
            for style_id, num_info in pstyle_to_num.items():
                if style_id not in style_numbering_map:
                    style_numbering_map[style_id] = num_info

        return style_numbering_map

    def extract_actual_heading_numbers(self, document_xml, numbering_definitions, style_numbering_map, heading_styles=None):
        """提取文档中实际的标题编号

        采用两级策略：
        1. 优先使用自动编号：当段落的样式在 style_numbering_map 中时，
           根据编号定义和计数器计算编号字符串
        2. 回退到文本正则提取：当标题样式没有关联自动编号时，
           从标题文本中用正则提取编号（如 "6.8.1 供应国资格" -> "6.8.1"）

        返回:
            {para_idx: {'number': 编号字符串, 'text': 段落文本}}
        """
        heading_numbers = {}
        xml_parser = BeautifulSoup(document_xml, 'xml')
        counters = {}

        heading_style_ids = set()
        if heading_styles:
            heading_style_ids = set(heading_styles.keys())

        for para_idx, paragraph in enumerate(xml_parser.find_all('w:p')):
            pPr = paragraph.find('w:pPr')
            if not pPr:
                continue

            pStyle = pPr.find('w:pStyle')
            if not pStyle:
                continue

            style_id = pStyle.get('w:val', '')

            # 策略1: 样式在编号映射中，使用自动编号计算
            if style_id in style_numbering_map:
                num_info = style_numbering_map[style_id]
                num_id = num_info['num_id']
                levels = num_info['levels']

                numPr = pPr.find('w:numPr')
                if numPr:
                    ilvl_elem = numPr.find('w:ilvl')
                    level = int(ilvl_elem.get('w:val', '0')) if ilvl_elem else 0
                else:
                    level = num_info['start_level']

                level_info = levels.get(level, {})
                current_val = self.maintain_numbering_counters(level, num_id, counters, level_info)

                text_pattern = level_info.get('text_pattern', '%1')
                number_str = self._format_number(current_val, level, text_pattern, levels, counters, num_id)

                heading_numbers[para_idx] = {
                    'number': number_str,
                    'text': self.extract_paragraph_text(paragraph).strip()
                }
            # 策略2: 样式不在编号映射中但是标题样式，回退到文本正则提取编号
            elif style_id in heading_style_ids:
                text = self.extract_paragraph_text(paragraph).strip()
                extracted = self._extract_number_from_text(text)
                if extracted:
                    heading_numbers[para_idx] = {
                        'number': extracted,
                        'text': text
                    }

        return heading_numbers

    def _extract_number_from_text(self, text):
        """从标题文本中用正则提取编号（回退方案）

        支持的编号格式：
        - "6.8 标题" / "6.8.1 标题"  -> 提取 "6.8" / "6.8.1"
        - "1. 标题"                   -> 提取 "1"
        - "（1）标题" / "(1) 标题"    -> 提取 "1"
        - "第一章标题" / "第二节标题"  -> 提取 "一" / "二"
        """
        patterns = [
            r'^(\d+(?:\.\d+)*)\s',
            r'^(\d+(?:\.\d+)*)\.',
            r'^(\d+(?:\.\d+)*)\s',
            r'^[（(](\d+(?:\.\d+)*)[）)]',
            r'^第([一二三四五六七八九十百千\d]+)[章节条]',
        ]
        for pattern in patterns:
            match = re.match(pattern, text)
            if match:
                return match.group(1)
        return None

    def _format_number(self, current_val, level, text_pattern, levels, counters, num_id):
        format_type = levels.get(level, {}).get('format', 'decimal')
        formatted = self._apply_number_format(current_val, format_type)

        if '%' not in text_pattern:
            return text_pattern

        result = text_pattern
        for lvl in range(level + 1):
            placeholder = f'%{lvl + 1}'
            if placeholder in result:
                lvl_val = counters.get(num_id, {}).get(lvl, 1)
                lvl_format = levels.get(lvl, {}).get('format', 'decimal')
                result = result.replace(placeholder, self._apply_number_format(lvl_val, lvl_format))

        return result

    def _apply_number_format(self, value, format_type):
        if format_type == 'decimal':
            return str(value)
        elif format_type == 'lowerLetter':
            return chr(ord('a') + value - 1)
        elif format_type == 'upperLetter':
            return chr(ord('A') + value - 1)
        elif format_type == 'lowerRoman':
            return self._to_roman(value).lower()
        elif format_type == 'upperRoman':
            return self._to_roman(value)
        elif format_type == 'chineseCounting':
            cn_nums = '零一二三四五六七八九十'
            if value <= 10:
                return cn_nums[value]
            elif value < 20:
                return f'十{cn_nums[value - 10]}' if value > 10 else '十'
            else:
                return str(value)
        elif format_type == 'ideographTraditional':
            cn_nums = '零壹贰叁肆伍陆柒捌玖拾'
            if value <= 10:
                return cn_nums[value]
            return str(value)
        else:
            return str(value)

    def _to_roman(self, num):
        val = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1]
        syms = ['M', 'CM', 'D', 'CD', 'C', 'XC', 'L', 'XL', 'X', 'IX', 'V', 'IV', 'I']
        result = ''
        for i in range(len(val)):
            while num >= val[i]:
                result += syms[i]
                num -= val[i]
        return result

    def maintain_numbering_counters(self, level, numbering_id, counters, level_info):
        """维护多级编号的计数器状态"""
        if numbering_id not in counters:
            counters[numbering_id] = {}

        # 初始化或递增当前级别
        if level not in counters[numbering_id]:
            start_val = level_info.get('start_value', 1)
            counters[numbering_id][level] = start_val
        else:
            counters[numbering_id][level] += 1

        # 重置更深层级的计数器
        levels_to_reset = [lv for lv in list(counters[numbering_id].keys()) if lv > level]
        for reset_level in levels_to_reset:
            del counters[numbering_id][reset_level]

        return counters[numbering_id][level]

    def extract_heading_hierarchy(self, archive):
        """提取标题层级结构 - 开源包不支持的功能"""
        try:
            styles_xml = archive.read('word/styles.xml')
            document_xml = archive.read('word/document.xml')
        except KeyError:
            return []

        # 解析样式定义
        heading_styles = self.parse_heading_style_definitions(styles_xml)

        # 提取文档中的标题内容
        document_headings = self.extract_document_headings(document_xml, heading_styles)

        return document_headings

    def parse_heading_style_definitions(self, styles_xml):
        """解析标题样式定义"""
        heading_styles = {}
        xml_parser = BeautifulSoup(styles_xml, 'xml')

        for style in xml_parser.find_all('w:style'):
            style_id = style.get('w:styleId')
            style_type = style.get('w:type')

            # 只处理段落样式
            if style_type == 'paragraph':
                style_name = self.get_style_name(style)

                # 识别标题样式
                if self.is_heading_style(style_name, style_id):
                    heading_level = self.extract_heading_level(style_name, style_id)
                    heading_styles[style_id] = {
                        'name': style_name,
                        'level': heading_level,
                        'formatting': self.extract_style_formatting(style)
                    }


        return heading_styles

    def extract_heading_level(self, style_name, style_id):
        """从样式信息中提取标题级别"""
        if style_name:
            # 匹配 "标题 1", "Heading 2" 等模式
            level_match = re.search(r'(\d+)', style_name)
            if level_match:
                return int(level_match.group(1))

        if style_id:
            # 匹配 "heading1", "toc1" 等模式
            level_match = re.search(r'(\d+)', style_id.lower())
            if level_match:
                return int(level_match.group(1))

        # 默认返回1级标题
        return 1

    def extract_document_headings(self, document_xml, heading_styles):
        headings = []
        xml_parser = BeautifulSoup(document_xml, 'xml')

        for paragraph in xml_parser.find_all('w:p'):
            pPr = paragraph.find('w:pPr')
            if not pPr:
                continue

            pStyle = pPr.find('w:pStyle')
            if not pStyle:
                continue

            style_id = pStyle.get('w:val', '')
            if style_id in heading_styles:
                text = self.extract_paragraph_text(paragraph)
                if text.strip():
                    headings.append({
                        'level': heading_styles[style_id]['level'],
                        'text': text.strip(),
                        'style_id': style_id,
                        'style_name': heading_styles[style_id]['name']
                    })

        return headings

    def get_style_name(self, style_element):
        name_elem = style_element.find('w:name')
        if name_elem:
            return name_elem.get('w:val', '')
        return ''

    def is_heading_style(self, style_name, style_id):
        heading_keywords = ['heading', '标题', 'heading ', '标题 ', 'toc']
        name_lower = (style_name or '').lower()
        id_lower = (style_id or '').lower()

        for keyword in heading_keywords:
            if keyword in name_lower or keyword in id_lower:
                return True

        if re.match(r'heading\d+', id_lower):
            return True

        return False

    def extract_style_formatting(self, style_element):
        formatting = {}

        rPr = style_element.find('w:rPr')
        if rPr:
            rFonts = rPr.find('w:rFonts')
            if rFonts:
                formatting['font'] = {
                    'ascii': rFonts.get('w:ascii', ''),
                    'east_asia': rFonts.get('w:eastAsia', ''),
                    'h_ansi': rFonts.get('w:hAnsi', '')
                }

            sz = rPr.find('w:sz')
            if sz:
                formatting['font_size_half_pt'] = int(sz.get('w:val', '0'))

            b = rPr.find('w:b')
            if b:
                formatting['bold'] = b.get('w:val', 'true') != 'false'

            i = rPr.find('w:i')
            if i:
                formatting['italic'] = i.get('w:val', 'true') != 'false'

            u = rPr.find('w:u')
            if u:
                formatting['underline'] = u.get('w:val', '')

            color = rPr.find('w:color')
            if color:
                formatting['color'] = color.get('w:val', '')

        pPr = style_element.find('w:pPr')
        if pPr:
            jc = pPr.find('w:jc')
            if jc:
                formatting['alignment'] = jc.get('w:val', '')

            spacing = pPr.find('w:spacing')
            if spacing:
                formatting['spacing'] = {
                    'before': spacing.get('w:before', ''),
                    'after': spacing.get('w:after', ''),
                    'line': spacing.get('w:line', '')
                }

        return formatting

    def process_review_changes(self, archive):
        """处理文档中的审阅修改内容

        审阅内容包括两类：
        1. 修订标记：w:ins（插入）和 w:del（删除），存储在 document.xml 中
        2. 评论：w:commentRangeStart/w:commentReference 引用，内容存储在 comments.xml 中
        需要同时读取 document.xml 和 comments.xml 才能完整提取审阅信息。
        """
        try:
            document_xml = archive.read('word/document.xml')
        except KeyError:
            return []

        # 读取 comments.xml，部分文档的审阅内容以评论形式存在而非修订标记
        comments_xml = None
        if 'word/comments.xml' in archive.namelist():
            try:
                comments_xml = archive.read('word/comments.xml')
            except KeyError:
                pass

        review_processor = ReviewContentProcessor(self.review_processing_mode)
        processed_content = review_processor.process_document_reviews(document_xml, comments_xml)

        return processed_content

    def merge_extraction_results(self, extraction_results):
        """整合所有提取结果为统一格式"""
        basic_content = extraction_results.get('basic_content', {})
        heading_numbers = extraction_results.get('heading_numbers', {})
        heading_styles = extraction_results.get('heading_styles', [])
        review_content = extraction_results.get('review_content', [])

        structured_elements = []
        for para in basic_content.get('paragraphs', []):
            structured_elements.append({
                'type': 'paragraph',
                'content': para.get('content', ''),
                'style_info': para.get('style_info', {})
            })
        for table in basic_content.get('tables', []):
            structured_elements.append({
                'type': 'table',
                'rows': table.get('rows', []),
                'properties': table.get('properties', {})
            })

        heading_outline = []
        for heading in heading_styles:
            outline_entry = {
                'level': heading.get('level', 1),
                'text': heading.get('text', ''),
                'style_id': heading.get('style_id', ''),
                'style_name': heading.get('style_name', ''),
                'number': ''
            }
            heading_outline.append(outline_entry)

        # 通过文本匹配将编号合并到标题大纲
        # heading_numbers 的 key 是文档段落绝对索引，heading_outline 是标题列表，
        # 两者索引不对应，因此通过段落文本进行匹配关联
        number_by_text = {}
        for para_idx, info in heading_numbers.items():
            text = info.get('text', '')
            number_by_text[text] = info.get('number', '')

        for entry in heading_outline:
            outline_text = entry['text']
            # 精确匹配
            if outline_text in number_by_text:
                entry['number'] = number_by_text[outline_text]
            else:
                # 模糊匹配：标题文本可能是编号文本的子串或反之
                for num_text, num_val in number_by_text.items():
                    if outline_text and num_text and (outline_text in num_text or num_text in outline_text):
                        entry['number'] = num_val
                        break

        # 统计审阅内容：遍历 changes 列表统计各类型数量
        # review_content 中每个 item 可能是 type='review'（含 changes 列表）或 type='comment'（孤立评论）
        review_summary = {
            'total_changes': len(review_content),
            'insertions': 0,
            'deletions': 0,
            'comments': 0,
            'details': review_content
        }
        for item in review_content:
            if item.get('type') == 'comment':
                review_summary['comments'] += 1
            for change in item.get('changes', []):
                if change.get('type') == 'insertion':
                    review_summary['insertions'] += 1
                elif change.get('type') == 'deletion':
                    review_summary['deletions'] += 1
                elif change.get('type') == 'comment':
                    review_summary['comments'] += 1

        final_result = {
            'metadata': {
                'processing_mode': self.review_processing_mode,
                'ocr_enabled': self.ocr_enabled,
                'extraction_timestamp': datetime.now().isoformat()
            },
            'content': {
                # 'structured_elements': [],# TODO
                # 'heading_outline': {},# TODO
                # 'review_summary': {}# TODO
                'structured_elements': structured_elements,
                'heading_outline': heading_outline,
                'review_summary': review_summary,
                'images': basic_content.get('images', {})
            }
        }

        return final_result


class ReviewContentProcessor:
    def __init__(self, processing_mode='final'):
        """
        初始化审阅内容处理器

        Args:
            processing_mode: 处理模式
                - 'final': 只保留最终文本
                - 'original': 只保留原始文本
                - 'both': 保留原始和修改后的文本
                - 'track': 保留完整的修改轨迹
        """
        self.mode = processing_mode

    def log_warning(self, message):
        logging.warning(message)

    def process_document_reviews(self, document_xml, comments_xml=None):
        """处理文档中的所有审阅内容（修订 + 评论）

        流程：
        1. 解析 comments.xml 为 {id: {author, date, content}} 映射
        2. 遍历文档段落，提取修订标记和评论引用
        3. 补充未被段落引用的孤立评论
        """
        xml_parser = BeautifulSoup(document_xml, 'xml')
        processed_items = []

        comments_map = {}
        if comments_xml:
            comments_map = self._parse_comments_map(comments_xml)

        for paragraph in xml_parser.find_all('w:p'):
            paragraph_content = self.process_paragraph_reviews(paragraph, comments_map)
            if paragraph_content:
                processed_items.append(paragraph_content)

        # 补充未被段落引用的孤立评论（评论内容在 comments.xml 中但段落中无引用标记）
        if comments_map:
            for comment_id, comment_info in comments_map.items():
                already_included = any(
                    item.get('comment_id') == comment_id or
                    any(c.get('comment_id') == comment_id for c in item.get('changes', []))
                    for item in processed_items
                )
                if not already_included:
                    processed_items.append({
                        'type': 'comment',
                        'comment_id': comment_id,
                        'author': comment_info['author'],
                        'date': comment_info['date'],
                        'content': comment_info['content']
                    })

        return processed_items

    def _parse_comments_map(self, comments_xml):
        """解析 comments.xml 为 {id: {author, date, content}} 映射表"""
        comments_map = {}
        xml_parser = BeautifulSoup(comments_xml, 'xml')
        for comment in xml_parser.find_all('w:comment'):
            cid = comment.get('w:id', '')
            text_parts = []
            for text_node in comment.find_all('w:t'):
                text_parts.append(text_node.get_text())
            comments_map[cid] = {
                'author': comment.get('w:author', 'Unknown'),
                'date': comment.get('w:date', ''),
                'content': ''.join(text_parts)
            }
        return comments_map

    def process_paragraph_reviews(self, paragraph, comments_map=None):
        """处理单个段落中的审阅内容

        提取三类审阅标记：
        1. w:ins - 插入内容
        2. w:del - 删除内容
        3. w:commentRangeStart/w:commentReference - 评论引用
           评论的具体内容需要从 comments_map 中查找（来源于 comments.xml）
        """
        insertions = paragraph.find_all('w:ins')
        deletions = paragraph.find_all('w:del')
        comment_starts = paragraph.find_all('w:commentRangeStart')
        comment_refs = paragraph.find_all('w:commentReference')

        if not insertions and not deletions and not comment_starts and not comment_refs:
            return None

        result = {'type': 'review', 'changes': []}

        for ins in insertions:
            change = {
                'type': 'insertion',
                'author': ins.get('w:author', 'Unknown'),
                'date': ins.get('w:date', ''),
                'content': self.extract_revision_text(ins)
            }
            result['changes'].append(change)

        for dele in deletions:
            change = {
                'type': 'deletion',
                'author': dele.get('w:author', 'Unknown'),
                'date': dele.get('w:date', ''),
                'content': self.extract_revision_text(dele)
            }
            result['changes'].append(change)

        # 从 comments_map 中查找段落引用的评论内容
        if comments_map:
            seen_ids = set()
            # 通过 w:commentRangeStart 定位评论作用范围的起始位置
            for cr in comment_starts:
                cid = cr.get('w:id', '')
                if cid and cid in comments_map and cid not in seen_ids:
                    seen_ids.add(cid)
                    cinfo = comments_map[cid]
                    result['changes'].append({
                        'type': 'comment',
                        'comment_id': cid,
                        'author': cinfo['author'],
                        'date': cinfo['date'],
                        'content': cinfo['content']
                    })
            # 通过 w:commentReference 获取评论引用（与 commentRangeStart 互补）
            for cr in comment_refs:
                cid = cr.get('w:id', '')
                if cid and cid in comments_map and cid not in seen_ids:
                    seen_ids.add(cid)
                    cinfo = comments_map[cid]
                    result['changes'].append({
                        'type': 'comment',
                        'comment_id': cid,
                        'author': cinfo['author'],
                        'date': cinfo['date'],
                        'content': cinfo['content']
                    })

        if not result['changes']:
            return None

        if self.mode == 'final':
            result['resolved_text'] = ' '.join(
                c['content'] for c in result['changes'] if c['type'] == 'insertion'
            )
        elif self.mode == 'original':
            result['resolved_text'] = ' '.join(
                c['content'] for c in result['changes'] if c['type'] == 'deletion'
            )
        elif self.mode == 'both':
            result['resolved_text'] = {
                'original': ' '.join(
                    c['content'] for c in result['changes'] if c['type'] == 'deletion'
                ),
                'modified': ' '.join(
                    c['content'] for c in result['changes'] if c['type'] == 'insertion'
                )
            }
        elif self.mode == 'track':
            pass

        return result

    def extract_revision_text(self, revision_element):
        text_parts = []
        for text_node in revision_element.find_all('w:t'):
            text_parts.append(text_node.get_text())
        return ''.join(text_parts)

    def extract_collaboration_history(self, archive):
        """提取协作历史信息"""
        collaboration_data = {
            'authors': set(),
            'change_timeline': [],
            'comment_threads': []
        }

        try:
            document_xml = archive.read('word/document.xml')
            comments_xml = archive.read('word/comments.xml') if 'word/comments.xml' in archive.namelist() else None

            # 提取修改历史
            self.extract_change_timeline(document_xml, collaboration_data)

            # 提取评论信息
            if comments_xml:
                self.extract_comments(comments_xml, collaboration_data)

        except Exception as e:
            self.log_warning(f"协作历史提取失败: {e}")

        collaboration_data['authors'] = list(collaboration_data['authors'])
        return collaboration_data

    def extract_change_timeline(self, document_xml, collaboration_data):
        """提取修改时间线"""
        xml_parser = BeautifulSoup(document_xml, 'xml')

        # 查找所有修订元素
        for revision in xml_parser.find_all(['w:ins', 'w:del']):
            change_info = {
                'type': 'insertion' if revision.name == 'w:ins' else 'deletion',
                'author': revision.get('w:author', 'Unknown'),
                'date': revision.get('w:date', ''),
                'content': self.extract_revision_text(revision)
            }

            collaboration_data['authors'].add(change_info['author'])
            collaboration_data['change_timeline'].append(change_info)

    def extract_comments(self, comments_xml, collaboration_data):
        xml_parser = BeautifulSoup(comments_xml, 'xml')

        for comment in xml_parser.find_all('w:comment'):
            comment_info = {
                'id': comment.get('w:id', ''),
                'author': comment.get('w:author', 'Unknown'),
                'date': comment.get('w:date', ''),
                'content': ''
            }

            text_parts = []
            for text_node in comment.find_all('w:t'):
                text_parts.append(text_node.get_text())
            comment_info['content'] = ''.join(text_parts)

            collaboration_data['comment_threads'].append(comment_info)
            collaboration_data['authors'].add(comment_info['author'])


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='DOCX文档解析工具')
    parser.add_argument('input', help='输入docx文件路径')
    parser.add_argument('-o', '--output', default='./output', help='输出目录 (默认: ./output)')
    parser.add_argument('--ocr', action='store_true', help='启用OCR图片识别')
    parser.add_argument('--review-mode', choices=['final', 'original', 'both', 'track'],
                        default='final', help='审阅处理模式 (默认: final)')
    parser.add_argument('--export', choices=['json', 'text'], default='json', help='导出格式 (默认: json)')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f'错误: 文件不存在 - {args.input}')
        exit(1)

    os.makedirs(args.output, exist_ok=True)

    processor = EnterpriseDocxProcessor(
        file_path=args.input,
        output_dir=args.output,
        enable_ocr=args.ocr,
        review_mode=args.review_mode
    )

    result = processor.extract_comprehensive_content()

    if args.export == 'json':
        output_path = os.path.join(args.output, 'result.json')
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f'结果已导出至: {output_path}')
    else:
        output_path = os.path.join(args.output, 'result.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f"文档解析结果\n{'=' * 50}\n\n")

            f.write(f"处理时间: {result['metadata']['extraction_timestamp']}\n")
            f.write(f"审阅模式: {result['metadata']['processing_mode']}\n")
            f.write(f"OCR启用: {result['metadata']['ocr_enabled']}\n\n")

            f.write(f"标题大纲\n{'-' * 30}\n")
            for heading in result['content']['heading_outline']:
                indent = '  ' * (heading['level'] - 1)
                number = heading.get('number', '')
                f.write(f"{indent}{number} {heading['text']}\n")

            f.write(f"\n内容元素\n{'-' * 30}\n")
            for element in result['content']['structured_elements']:
                if element['type'] == 'paragraph':
                    f.write(f"{element['content']}\n")
                elif element['type'] == 'table':
                    for row in element.get('rows', []):
                        cells = [cell.get('content', '') for cell in row]
                        f.write(' | '.join(cells) + '\n')
                    f.write('\n')

            review = result['content']['review_summary']
            f.write(f"\n审阅摘要\n{'-' * 30}\n")
            f.write(f"总修改数: {review['total_changes']}\n")
            f.write(f"插入: {review['insertions']}\n")
            f.write(f"删除: {review['deletions']}\n")

        print(f'结果已导出至: {output_path}')
