import os
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from docx import Document
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

class WordToPdfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word转PDF转换器")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        
        # 设置中文字体
        self.font = ("SimHei", 10)
        
        # 创建主框架
        self.main_frame = tk.Frame(root)
        self.main_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
        
        # 创建标题
        self.title_label = tk.Label(self.main_frame, text="Word转PDF转换器", font=("SimHei", 16, "bold"))
        self.title_label.pack(pady=10)
        
        # 文件路径显示
        self.file_path_var = tk.StringVar()
        self.file_path_var.set("未选择文件")
        self.file_path_entry = tk.Entry(self.main_frame, textvariable=self.file_path_var, width=50, font=self.font)
        self.file_path_entry.pack(pady=10, fill=tk.X)
        
        # 选择文件按钮
        self.select_button = tk.Button(self.main_frame, text="选择Word文件", command=self.select_file, font=self.font)
        self.select_button.pack(pady=10)
        
        # 转换按钮
        self.convert_button = tk.Button(self.main_frame, text="转换为PDF", command=self.convert_to_pdf, font=self.font)
        self.convert_button.pack(pady=10)
        
        # 状态标签
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        self.status_label = tk.Label(self.main_frame, textvariable=self.status_var, font=self.font)
        self.status_label.pack(pady=20)
    
    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Word文件",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.status_var.set(f"已选择文件: {os.path.basename(file_path)}")
    
    def convert_to_pdf(self):
        # 在方法开始处就定义dpi变量，确保在所有执行路径中都有定义
        dpi = 300
        
        file_path = self.file_path_var.get()
        
        if file_path == "未选择文件":
            messagebox.showerror("错误", "请先选择一个Word文件")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("错误", "所选文件不存在")
            return
        
        try:
            self.status_var.set("正在转换...")
            self.root.update()
            
            # 转换文件
            pdf_path = os.path.splitext(file_path)[0] + ".pdf"
            
            # 注册Windows系统中的中文字体
            # 尝试使用几种常见的中文字体
            font_registered = False
            for font_name in ['SimHei', 'Microsoft YaHei', 'SimSun']:
                try:
                    # 检查系统字体是否存在
                    import matplotlib.font_manager as fm
                    fonts = [f.name for f in fm.fontManager.ttflist]
                    if font_name in fonts:
                        # 注册字体
                        pdfmetrics.registerFont(TTFont(font_name, f"C:\\Windows\\Fonts\\{font_name}.ttf"))
                        font_registered = True
                        break
                except:
                    # 如果matplotlib不可用，尝试直接注册
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, f"C:\\Windows\\Fonts\\{font_name}.ttf"))
                        font_registered = True
                        break
                    except:
                        continue
            
            # 使用python-docx读取Word文档
            doc = Document(file_path)
            
            # 创建PDF文档
            pdf = SimpleDocTemplate(pdf_path, pagesize=A4)
            
            # 获取样式并设置中文字体
            styles = getSampleStyleSheet()
            
            # 创建各种支持中文的样式
            custom_styles = {}
            if font_registered:
                base_font = font_name
                # 创建普通段落样式
                custom_styles['Normal'] = ParagraphStyle(
                    'CustomNormal',
                    parent=styles['Normal'],
                    fontName=base_font,
                    fontSize=12,
                    spaceAfter=12
                )
                
                # 创建各级标题样式
                heading_sizes = [20, 18, 16, 14, 13, 12]
                heading_fonts = [base_font, base_font, base_font, base_font, base_font, base_font]
                
                for i in range(1, 7):
                    custom_styles[f'Heading{i}'] = ParagraphStyle(
                        f'CustomHeading{i}',
                        parent=styles.get(f'Heading{i}', styles['Heading1']),
                        fontName=heading_fonts[i-1],
                        fontSize=heading_sizes[i-1],
                        bold=True,
                        spaceAfter=18,
                        spaceBefore=12,
                        leftIndent=0 if i == 1 else (i-1) * 20
                    )
                
                # 创建目录样式
                custom_styles['TOC'] = ParagraphStyle(
                    'CustomTOC',
                    parent=styles['Normal'],
                    fontName=base_font,
                    fontSize=12,
                    spaceAfter=6
                )
                
                # 为目录项创建不同级别的样式
                for i in range(1, 5):
                    custom_styles[f'TOC{i}'] = ParagraphStyle(
                        f'CustomTOC{i}',
                        parent=styles['Normal'],
                        fontName=base_font,
                        fontSize=12,
                        leftIndent=i * 30,
                        spaceAfter=6
                    )
            else:
                # 如果无法注册中文字体，使用默认样式
                for style_name in ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6']:
                    custom_styles[style_name] = styles.get(style_name, styles['Normal'])
            
            # 创建内容列表
            flowables = []
            
            # 收集文档中的标题，用于重建目录
            headings = []
            for para in doc.paragraphs:
                if para.style.name.startswith('Heading'):
                    level_match = re.search(r'\d+', para.style.name)
                    if level_match:
                        level = int(level_match.group())
                        headings.append((level, para.text))
            
            # 检查是否有目录内容
            has_toc = any('目录' in para.text or 'Contents' in para.text for para in doc.paragraphs)
            
            # 如果文档中提到了目录，添加重建的目录
            if has_toc:
                # 添加目录标题
                toc_title = Paragraph("目录", custom_styles.get('Heading1', custom_styles['Normal']))
                flowables.append(toc_title)
                flowables.append(Spacer(1, 0.3*inch))
                
                # 添加目录项
                for level, text in headings:
                    if level <= 4:  # 只包含前4级标题
                        toc_text = text
                        toc_item = Paragraph(toc_text, custom_styles.get(f'TOC{level}', custom_styles['TOC']))
                        flowables.append(toc_item)
                
                flowables.append(Spacer(1, 0.5*inch))
            
            # 处理每个段落并保留格式
            for para in doc.paragraphs:
                # 跳过目录占位文本（如果有的话）
                if has_toc and ('目录' in para.text or 'Contents' in para.text) and para.style.name.startswith('Heading'):
                    continue
                
                if para.text.strip():
                    # 根据段落样式选择适当的样式
                    if para.style.name in custom_styles:
                        # 使用对应的自定义样式
                        pdf_style = custom_styles[para.style.name]
                    elif para.style.name.startswith('Heading'):
                        # 处理可能的其他标题级别
                        level_match = re.search(r'\d+', para.style.name)
                        if level_match:
                            level = int(level_match.group())
                            if level <= 6:
                                pdf_style = custom_styles.get(f'Heading{level}', custom_styles['Normal'])
                            else:
                                pdf_style = custom_styles['Normal']
                        else:
                            pdf_style = custom_styles['Normal']
                    else:
                        # 创建基于原始段落格式的样式
                        base_style = custom_styles['Normal']
                        
                        # 尝试获取Word段落的格式信息
                        # 注意：python-docx对格式的支持有限，这里只能获取部分基本格式
                        try:
                            # 创建新的段落样式
                            formatted_style = ParagraphStyle(
                                'CustomFormatted',
                                parent=base_style
                            )
                            
                            # 设置段落对齐方式
                            alignment_map = {
                                0: 'LEFT',
                                1: 'CENTER',
                                2: 'RIGHT',
                                3: 'JUSTIFY'
                            }
                            if hasattr(para.paragraph_format, 'alignment') and para.paragraph_format.alignment is not None:
                                if para.paragraph_format.alignment in alignment_map:
                                    formatted_style.alignment = alignment_map[para.paragraph_format.alignment]
                            
                            # 设置左缩进
                            if hasattr(para.paragraph_format, 'left_indent') and para.paragraph_format.left_indent:
                                # 确保inch已正确导入和定义
                                if 'inch' in locals():
                                    formatted_style.leftIndent = para.paragraph_format.left_indent.inches * inch
                            
                            # 设置行距
                            if hasattr(para.paragraph_format, 'space_after') and para.paragraph_format.space_after:
                                formatted_style.spaceAfter = para.paragraph_format.space_after.pt
                            
                            pdf_style = formatted_style
                        except:
                            pdf_style = base_style
                    
                    # 创建段落并添加到内容列表
                    formatted_text = para.text
                    # 替换特殊字符，确保PDF生成正常
                    formatted_text = formatted_text.replace('\t', '    ')
                    
                    flowables.append(Paragraph(formatted_text, pdf_style))
                    
                    # 添加适当的间距
                    if para.style.name.startswith('Heading'):
                        flowables.append(Spacer(1, 0.2*inch))
            
            # 处理表格并保留格式
            for table in doc.tables:
                data = []
                # 获取表格数据
                for row in table.rows:
                    data_row = []
                    for cell in row.cells:
                        data_row.append(cell.text.strip())
                    data.append(data_row)
                
                if data:
                    # 创建表格
                    table_obj = Table(data)
                    # 添加样式
                    table_font = font_name if font_registered else 'Helvetica'
                    table_style = TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # 改为左对齐更符合中文阅读习惯
                        ('FONTNAME', (0, 0), (-1, -1), table_font),  # 设置表格中文字体
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                        ('TOPPADDING', (0, 0), (-1, -1), 8),
                        ('LEFTPADDING', (0, 0), (-1, -1), 10),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ])
                    table_obj.setStyle(table_style)
                    flowables.append(table_obj)
                    flowables.append(Spacer(1, 0.5*inch))
            
            # 构建PDF
            pdf.build(flowables)
            
            self.status_var.set(f"转换成功! PDF已保存至: {os.path.basename(pdf_path)}")
            messagebox.showinfo("成功", f"文件已成功转换为PDF\n保存路径: {pdf_path}")
            
        except Exception as e:
            self.status_var.set(f"转换失败: {str(e)}")
            messagebox.showerror("错误", f"转换过程中发生错误:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WordToPdfConverter(root)
    root.mainloop()