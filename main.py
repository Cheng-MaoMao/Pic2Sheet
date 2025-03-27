import os
from openai import OpenAI
import openpyxl
from datetime import datetime
import json

class ImageAnalyzer:
    def __init__(self, provider="阿里", api_key=None):
        config_path = "config.json"
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        except Exception:
            # 使用默认配置
            self.config = {
                "阿里": {
                    "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
                    "model": "qwen-vl-max-latest"
                },
                "火山引擎": {
                    "base_url": "https://ark.cn-beijing.volces.com/api/v3",
                    "model": "doubao-1-5-vision-pro-32k-250115"
                }
            }
        
        if provider not in self.config:
            raise ValueError("不支持的提供商")
            
        provider_config = self.config[provider]
        self.provider = provider
        self.client = OpenAI(
            api_key=api_key,
            base_url=provider_config["base_url"]
        )
        self.model = provider_config["model"]
        
    def analyze_image(self, image_path):
        prompt = """请分析图片内容：
1. 分析图片，如果图片包含表格，请严格按照图片中的表格格式返回数据(保留空的内容)，以JSON数组形式，例如：
[["表头1", "表头2"], ["数据1", "数据2"]]

2. 分析图片，如果图片不包含表格，请分析内容并将相关信息整理成表格形式，以JSON数组返回，例如：
[["类别", "内容"], ["姓名", "张三"], ["年龄", "25"]]

请确保返回格式为标准JSON数组。"""

        if image_path.startswith(('http://', 'https://')):
            image_url = image_path
        else:
            import base64
            with open(image_path, "rb") as image_file:
                image_data = base64.b64encode(image_file.read()).decode('utf-8')
                image_url = f"data:image/jpeg;base64,{image_data}"
        
        try:
            completion = self.client.chat.completions.create(
                model=self.model,  # 使用配置文件中的模型名称
                temperature = 0.7,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": image_url}}
                    ]
                }]
            )
            
            response_data = json.loads(completion.model_dump_json())
            content = response_data['choices'][0]['message']['content']
            
            # 打印原始返回内容，用于调试
            print("原始返回内容:", content)
            print("提供商:", self.provider)
            print("使用的模型:", self.model)
            
            # 尝试清理和格式化内容
            content = content.strip()
            if content.startswith('```') and content.endswith('```'):
                content = content[3:-3].strip()
            if content.startswith('json'):
                content = content[4:].strip()
                
            return content
            
        except Exception as e:
            print(f"分析图片时发生错误: {str(e)}")
            return None

class ExcelHandler:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.workbook = None
        self.sheet = None
        self._init_workbook()
        
    def _init_workbook(self):
        if os.path.exists(self.excel_path):
            self.workbook = openpyxl.load_workbook(self.excel_path)
        else:
            self.workbook = openpyxl.Workbook()
        self.sheet = self.workbook.active
        
    def write_data(self, data):
        try:
            if isinstance(data, str):
                # 尝试解析JSON数组
                try:
                    parsed_data = json.loads(data)
                    if not isinstance(parsed_data, list):
                        raise ValueError("数据格式不是数组")
                    
                    # 清空当前工作表
                    self.sheet.delete_rows(1, self.sheet.max_row)
                    
                    # 直接按照JSON数组格式写入数据
                    for row_idx, row_data in enumerate(parsed_data, start=1):
                        for col_idx, value in enumerate(row_data, start=1):
                            cell_value = str(value) if value else ""  # 处理空值
                            self.sheet.cell(row=row_idx, column=col_idx, value=cell_value)
                    
                    # 设置表头样式
                    if parsed_data:
                        for col_idx in range(1, len(parsed_data[0]) + 1):
                            cell = self.sheet.cell(row=1, column=col_idx)
                            cell.font = openpyxl.styles.Font(bold=True)
                    
                    # 自动调整列宽
                    for column in self.sheet.columns:
                        max_length = 0
                        column = [cell for cell in column if cell.value]
                        for cell in column:
                            try:
                                max_length = max(max_length, len(str(cell.value)))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 40)  # 限制最大宽度
                        self.sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                    
                except json.JSONDecodeError as e:
                    print(f"JSON解析错误: {str(e)}")
                    return False
                    
                self.workbook.save(self.excel_path)
                return True
                
        except Exception as e:
            print(f"写入Excel时发生错误: {str(e)}")
            return False

def main():
    # 初始化图片分析器
    analyzer = ImageAnalyzer()
    
    # 测试本地图片
    local_image = r"C:\Users\Cheng-MaoMao\Desktop\PictureTransferForm\微信图片_20250327203511.jpg"
    
    # 根据图片路径生成Excel文件名
    image_name = os.path.splitext(os.path.basename(local_image))[0]
    excel_path = os.path.join(os.path.dirname(local_image), f"{image_name}.xlsx")
    
    # 创建Excel处理器
    excel_handler = ExcelHandler(excel_path)
    
    # 分析图片并处理返回结果
    result = analyzer.analyze_image(local_image)
    if result:
        try:
            # 打印处理前的结果
            print("处理前的结果:", result)
            
            # 移除可能的干扰字符
            result = result.strip().lstrip('\ufeff')
            
            # 确保result不为空且是有效的JSON
            if result:
                try:
                    # 验证JSON格式
                    json.loads(result)
                    success = excel_handler.write_data(result)
                    if success:
                        print(f"表格已保存至: {excel_path}")
                    else:
                        print("表格保存失败")
                except json.JSONDecodeError as e:
                    print(f"JSON格式无效: {str(e)}")
                    print("原始数据:", result)
            else:
                print("分析结果为空")
        except Exception as e:
            print(f"处理结果时出错: {str(e)}")
            print("错误的数据:", result)
    else:
        print("图片分析失败")

if __name__ == "__main__":
    main()