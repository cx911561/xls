import os
import xlrd
import chardet
import codecs
from collections import defaultdict

def detect_file_encoding(file_path):
    """检测文件的编码格式"""
    with open(file_path, 'rb') as f:
        raw_data = f.read(10000)
    
    result = chardet.detect(raw_data)
    encoding = result['encoding']
    confidence = result['confidence']
    
    print(f"检测到编码: {encoding} (置信度: {confidence:.2f})")
    
    return encoding

def fix_chinese_encoding(text, source_encoding='utf-8'):
    """修复中文乱码"""
    if text is None:
        return ""
    
    text = str(text)
    
    common_encodings = ['utf-8', 'gbk', 'gb2312', 'big5', 'cp1252']
    
    if source_encoding:
        try:
            return text.encode(source_encoding, errors='ignore').decode('utf-8')
        except:
            pass
    
    for encoding in common_encodings:
        try:
            decoded = text.encode('latin-1').decode(encoding)
            if is_valid_chinese(decoded):
                return decoded
        except:
            continue
    
    return text

def is_valid_chinese(text):
    """检查字符串是否包含有效中文"""
    chinese_count = 0
    for char in text:
        if '\u4e00' <= char <= '\u9fff':
            chinese_count += 1
    
    return chinese_count > 0

def read_xls_file(file_path, encoding=None):
    """读取XLS文件并处理编码问题"""
    print(f"\n正在处理文件: {file_path}")
    
    if not encoding:
        encoding = detect_file_encoding(file_path)
    
    try:
        workbook = xlrd.open_workbook(file_path, encoding_override=encoding)
    except Exception as e:
        print(f"使用编码 {encoding} 打开失败，尝试其他编码...")
        for alt_encoding in ['gbk', 'gb2312', 'utf-8', 'big5']:
            try:
                workbook = xlrd.open_workbook(file_path, encoding_override=alt_encoding)
                encoding = alt_encoding
                print(f"成功使用编码 {alt_encoding}")
                break
            except:
                continue
        else:
            print(f"所有编码尝试失败，使用默认编码")
            workbook = xlrd.open_workbook(file_path)
    
    return workbook

def parse_and_calculate(file_path):
    """解析XLS文件并计算每列数字之和"""
    try:
        workbook = read_xls_file(file_path)
        sheet = workbook.sheet_by_index(0)
        
        print(f"\n工作表名称: {sheet.name}")
        print(f"行数: {sheet.nrows}, 列数: {sheet.ncols}")
        
        headers = []
        for col in range(sheet.ncols):
            header = sheet.cell_value(0, col)
            header = fix_chinese_encoding(header)
            headers.append(header)
        
        column_sums = defaultdict(float)
        column_types = {}
        
        for col in range(sheet.ncols):
            col_sum = 0.0
            has_number = False
            
            for row in range(1, sheet.nrows):
                cell_value = sheet.cell_value(row, col)
                
                if isinstance(cell_value, (int, float)):
                    col_sum += cell_value
                    has_number = True
                elif isinstance(cell_value, str) and cell_value.strip():
                    try:
                        col_sum += float(cell_value.strip())
                        has_number = True
                    except ValueError:
                        pass
            
            column_sums[col] = col_sum
            column_types[col] = 'numeric' if has_number else 'text'
        
        print("\n" + "="*60)
        print(f"{'列名':<20} {'类型':<10} {'数字求和':<20}")
        print("="*60)
        
        for col in range(sheet.ncols):
            header = headers[col]
            col_type = column_types[col]
            col_sum = column_sums[col]
            
            if col_type == 'numeric':
                sum_str = f"{col_sum:,.2f}"
            else:
                sum_str = "(无数字)"
            
            print(f"{header:<20} {col_type:<10} {sum_str:<20}")
        
        print("="*60)
        
        return {
            'file_name': os.path.basename(file_path),
            'sheet_name': sheet.name,
            'headers': headers,
            'column_sums': dict(column_sums),
            'column_types': dict(column_types)
        }
    
    except Exception as e:
        print(f"解析文件失败: {str(e)}")
        return None

def batch_process_xls_files(folder_path):
    """批量处理文件夹中的所有XLS文件"""
    results = []
    
    if not os.path.isdir(folder_path):
        print(f"错误: {folder_path} 不是有效的文件夹")
        return results
    
    xls_files = [f for f in os.listdir(folder_path) 
                 if f.lower().endswith('.xls') or f.lower().endswith('.xlsx')]
    
    if not xls_files:
        print("未找到XLS文件")
        return results
    
    print(f"找到 {len(xls_files)} 个Excel文件")
    print("-" * 40)
    
    for filename in xls_files:
        file_path = os.path.join(folder_path, filename)
        result = parse_and_calculate(file_path)
        if result:
            results.append(result)
    
    return results

def export_results(results, output_file='results.txt'):
    """导出结果到文件"""
    with codecs.open(output_file, 'w', encoding='utf-8') as f:
        for result in results:
            f.write(f"文件名: {result['file_name']}\n")
            f.write(f"工作表: {result['sheet_name']}\n")
            f.write("-" * 40 + "\n")
            
            for i, header in enumerate(result['headers']):
                col_type = result['column_types'][i]
                col_sum = result['column_sums'][i]
                
                if col_type == 'numeric':
                    sum_str = f"{col_sum:,.2f}"
                else:
                    sum_str = "(无数字)"
                
                f.write(f"{header}\t{col_type}\t{sum_str}\n")
            
            f.write("\n")
    
    print(f"\n结果已导出到: {output_file}")

if __name__ == '__main__':
    print("=" * 60)
    print("          XLS文件中文乱码处理工具")
    print("=" * 60)
    
    import argparse
    parser = argparse.ArgumentParser(description='处理XLS文件中文乱码并计算列求和')
    parser.add_argument('path', help='XLS文件路径或包含XLS文件的文件夹路径')
    parser.add_argument('--output', default='results.txt', help='输出文件路径')
    args = parser.parse_args()
    
    input_path = args.path
    
    if os.path.isfile(input_path):
        if input_path.lower().endswith(('.xls', '.xlsx')):
            result = parse_and_calculate(input_path)
            if result:
                export_results([result], args.output)
        else:
            print("错误: 请提供XLS或XLSX文件")
    
    elif os.path.isdir(input_path):
        results = batch_process_xls_files(input_path)
        if results:
            export_results(results, args.output)
    
    else:
        print(f"错误: {input_path} 不存在")