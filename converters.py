import os
import docx
from PyQt5.QtCore import QThread, pyqtSignal
from utils import (extract_text_simple, extract_text_with_sections, 
                  extract_text_from_pdf, convert_md_to_word, 
                  convert_md_to_pdf, merge_markdown_files)

class ToMarkdownThread(QThread):
    """将Word/PDF文档转换为Markdown的线程"""
    update_progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    file_progress = pyqtSignal(int, int)  # current_file, total_files
    
    def __init__(self, file_list, output_dir, mode='simple', merge_output=False, file_type='word'):
        super().__init__()
        self.file_list = file_list
        self.output_dir = output_dir
        self.mode = mode  # 'simple' 或 'sections'
        self.merge_output = merge_output  # 是否合并输出
        self.file_type = file_type  # 'word' 或 'pdf'
        
    def process_pdf_file(self, pdf_path, file_name):
        """处理PDF文件转换"""
        self.update_progress.emit(40, f"正在提取PDF内容: {file_name}.pdf")
        
        # 提取PDF文本
        markdown_text = extract_text_from_pdf(pdf_path)
        
        # 保存为Markdown文件
        md_path = os.path.join(self.output_dir, f"{file_name}.md")
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(markdown_text)
            
        self.update_progress.emit(100, f"已完成PDF转换: {file_name}.md")
        return 1, 0, markdown_text  # 返回处理的文件数和章节数
        
    def process_simple_mode(self, doc, file_name):
        """处理简单模式的文档转换"""
        markdown_text = extract_text_simple(doc)
        md_path = os.path.join(self.output_dir, f"{file_name}.md")
        
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(markdown_text)
            
        return 1, 0, markdown_text  # 返回处理的文件数、章节数和转换的文本
    
    def process_sections_mode(self, doc, file_name):
        """处理分割模式的文档转换"""
        sections = extract_text_with_sections(doc)
        section_count = len(sections)
        self.update_progress.emit(60, f"文档解析完成，发现 {section_count} 个章节")
        
        # 为每个文件创建子目录
        file_dir = os.path.join(self.output_dir, file_name)
        os.makedirs(file_dir, exist_ok=True)
        
        # 保存每个章节
        all_content = []  # 用于可能的合并输出
        
        for idx, (title, content) in enumerate(sections.items()):
            # 创建安全的文件名
            safe_title = title.replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
            md_path = os.path.join(file_dir, f"{safe_title}.md")
            
            with open(md_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            all_content.append(content)
                
            progress = 60 + ((idx + 1) / section_count) * 30
            self.update_progress.emit(int(progress), f"已创建文件: {safe_title}.md")
            
        return 1, section_count, "\n\n---\n\n".join(all_content)  # 返回处理的文件数、章节数和合并内容
        
    def run(self):
        try:
            # 确保输出目录存在
            os.makedirs(self.output_dir, exist_ok=True)
            
            total_files = len(self.file_list)
            processed_count = 0
            total_sections = 0
            merged_content = []  # 用于合并内容
            
            for idx, file_path in enumerate(self.file_list):
                try:
                    # 更新处理文件进度
                    self.file_progress.emit(idx + 1, total_files)
                    
                    # 获取文件名（不含扩展名）
                    file_name = os.path.splitext(os.path.basename(file_path))[0]
                    file_ext = os.path.splitext(file_path)[1].lower()
                    
                    if self.file_type == 'pdf':
                        # PDF处理
                        self.update_progress.emit(10, f"正在处理PDF文件: {file_name}.pdf")
                        _, _, pdf_content = self.process_pdf_file(file_path, file_name)
                        
                        if self.merge_output:
                            # 添加文件标题和分隔符
                            merged_content.append(f"# {file_name}\n\n{pdf_content}")
                        
                        processed_count += 1
                        
                    else:
                        # Word处理
                        self.update_progress.emit(10, f"正在处理文件: {file_name}.docx")
                        
                        # 加载文档
                        self.update_progress.emit(30, "正在加载文档...")
                        doc = docx.Document(file_path)
                        
                        # 根据模式提取文本
                        self.update_progress.emit(50, "正在提取文档内容...")
                        
                        # 处理文档
                        if self.mode == 'simple':
                            files, sections, content = self.process_simple_mode(doc, file_name)
                            if self.merge_output:
                                # 添加文件标题和内容到合并列表
                                merged_content.append(f"# {file_name}\n\n{content}")
                            self.update_progress.emit(100, f"已完成转换: {file_name}.md")
                        else:
                            files, sections, content = self.process_sections_mode(doc, file_name)
                            if self.merge_output:
                                # 添加文件标题和内容到合并列表
                                merged_content.append(f"# {file_name}\n\n{content}")
                            self.update_progress.emit(100, f"已完成转换: {file_name} ({sections}个章节)")
                        
                        processed_count += files
                        total_sections += sections
                    
                except Exception as e:
                    self.update_progress.emit(0, f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
            
            # 如果需要合并输出
            if self.merge_output and merged_content:
                merged_md_path = os.path.join(self.output_dir, "合并文档.md")
                with open(merged_md_path, 'w', encoding='utf-8') as f:
                    f.write("\n\n---\n\n".join(merged_content))
                self.update_progress.emit(100, f"已创建合并文档: 合并文档.md")
            
            # 完成消息
            if self.file_type == 'pdf':
                self.finished.emit(True, f"成功转换 {processed_count} 个PDF文件！")
            elif self.mode == 'simple':
                msg = f"成功转换 {processed_count} 个Word文件！"
                if self.merge_output:
                    msg += " 并已合并为一个文档"
                self.finished.emit(True, msg)
            else:
                msg = f"成功转换 {processed_count} 个Word文件，共 {total_sections} 个章节！"
                if self.merge_output:
                    msg += " 并已合并为一个文档"
                self.finished.emit(True, msg)
            
        except Exception as e:
            self.finished.emit(False, f"转换失败: {str(e)}")

class FromMarkdownThread(QThread):
    """将Markdown文档转换为Word/PDF的线程"""
    update_progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)
    file_progress = pyqtSignal(int, int)  # current_file, total_files
    
    def __init__(self, file_list, output_dir, target_format='word', merge_output=False):
        super().__init__()
        self.file_list = file_list
        self.output_dir = output_dir
        self.target_format = target_format  # 'word' 或 'pdf'
        self.merge_output = merge_output  # 是否合并输出
        
    def run(self):
        try:
            # 确保输出目录存在
            os.makedirs(self.output_dir, exist_ok=True)
            
            total_files = len(self.file_list)
            processed_count = 0
            
            # 如果需要合并，先合并Markdown文件
            if self.merge_output and len(self.file_list) > 1:
                self.update_progress.emit(10, "正在合并Markdown文件...")
                
                # 合并所有Markdown文件
                merged_md_path = os.path.join(self.output_dir, "合并文档.md")
                merge_markdown_files(self.file_list, merged_md_path)
                
                self.update_progress.emit(40, "已合并Markdown文件，开始转换...")
                
                # 转换合并后的文件
                if self.target_format == 'word':
                    output_path = os.path.join(self.output_dir, "合并文档.docx")
                    self.update_progress.emit(50, "正在转换为Word文档...")
                    success = convert_md_to_word(merged_md_path, output_path)
                else:
                    output_path = os.path.join(self.output_dir, "合并文档.pdf")
                    self.update_progress.emit(50, "正在转换为PDF文档...")
                    success = convert_md_to_pdf(merged_md_path, output_path)
                
                if success:
                    self.update_progress.emit(100, f"已完成合并转换: {os.path.basename(output_path)}")
                    processed_count = 1
                else:
                    self.update_progress.emit(0, "合并文档转换失败")
            
            # 单独处理每个文件
            else:
                for idx, md_path in enumerate(self.file_list):
                    try:
                        # 更新处理文件进度
                        self.file_progress.emit(idx + 1, total_files)
                        
                        # 获取文件名（不含扩展名）
                        file_name = os.path.splitext(os.path.basename(md_path))[0]
                        
                        if self.target_format == 'word':
                            # 转Word
                            output_path = os.path.join(self.output_dir, f"{file_name}.docx")
                            self.update_progress.emit(20, f"正在将 {file_name}.md 转换为Word...")
                            success = convert_md_to_word(md_path, output_path)
                        else:
                            # 转PDF
                            output_path = os.path.join(self.output_dir, f"{file_name}.pdf")
                            self.update_progress.emit(20, f"正在将 {file_name}.md 转换为PDF...")
                            success = convert_md_to_pdf(md_path, output_path)
                        
                        if success:
                            self.update_progress.emit(100, f"已完成转换: {os.path.basename(output_path)}")
                            processed_count += 1
                        else:
                            self.update_progress.emit(0, f"转换 {file_name} 失败")
                        
                    except Exception as e:
                        self.update_progress.emit(0, f"处理文件 {os.path.basename(md_path)} 时出错: {str(e)}")
            
            # 完成消息
            format_name = "Word" if self.target_format == 'word' else "PDF"
            if self.merge_output and total_files > 1:
                self.finished.emit(True, f"已将 {total_files} 个Markdown文件合并并转换为{format_name}文档！")
            else:
                self.finished.emit(True, f"成功转换 {processed_count} 个Markdown文件为{format_name}！")
            
        except Exception as e:
            self.finished.emit(False, f"转换失败: {str(e)}") 