import json

def import_list_from_file(file_path):
        # 初始化一个空列表
    with open(file_path, 'r') as file:
        # 读取文件的所有行
        content_raw = file.read()
    
    return json.loads(content_raw)
