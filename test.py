import openpyxl

def process_cell(cell_value):
    # 第一项处理: 删除,和)前的字符，但保留符号本身
    result = []
    skip = False
    for i in range(len(cell_value)):
        if skip:
            skip = False
            continue
        if cell_value[i] in {',', ')'} and i > 0:
            result.pop()
        result.append(cell_value[i])
        skip = cell_value[i] in {',', ')'}

    # 第二项处理: 插入数字1, 2, 3,... 在(前
    processed = []
    count = 0
    for char in result:
        if char == '(':
            processed.append(str(count + 1))
            count += 1
        processed.append(char)

    return ''.join(processed)

# 获取文件路径
input_file_path = input('请输入Excel文件的输入路径: ')
output_file_path = input('请输入Excel文件的输出路径: ')

# 打开Excel文件
workbook = openpyxl.load_workbook(input_file_path)

# 遍历每个工作簿
for sheet in workbook:
    # 遍历每个单元格
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.value = process_cell(cell.value)

# 保存修改后的Excel文件
workbook.save(output_file_path)

input_file_path = r'D:\BaiduSyncdisk\工作文件\卡本\projects\新疆和田\0阶段成果\PD\0103-和田地区CCER项目集体权属土地清单.xlsx'
output_file_path = r'D:\BaiduSyncdisk\工作文件\卡本\projects\新疆和田\0阶段成果\PD\0104-和田地区CCER项目集体权属土地清单.xlsx '
