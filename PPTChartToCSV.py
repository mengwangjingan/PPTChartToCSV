import xml.etree.ElementTree as ET
import csv
from collections import defaultdict

def generate_ordered_output(xml_file, output_csv):
    """严格保持XML原始顺序的最终版本"""
    namespaces = {
        "pkg": "http://schemas.microsoft.com/office/2006/xmlPackage",
        "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
    }

    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        # 数据结构
        series_data = defaultdict(dict)  # {系列名称: {分类: 值}}
        ordered_categories = []          # 保持XML中的原始顺序
        seen_categories = set()          # 用于去重
        all_series_names = []            # 系列名称按出现顺序存储

        # 第一次遍历：收集所有分类的原始顺序
        for series in root.findall(".//c:ser", namespaces):
            # 提取分类数据
            for pt in series.findall(".//c:cat//c:pt", namespaces):
                v_elem = pt.find("c:v", namespaces)
                if v_elem is not None and v_elem.text:
                    category = v_elem.text.strip()
                    if category not in seen_categories:
                        ordered_categories.append(category)
                        seen_categories.add(category)

        # 第二次遍历：收集系列数据
        for series in root.findall(".//c:ser", namespaces):
            # 提取系列名称
            name_elem = series.find(".//c:tx/c:strRef/c:strCache/c:pt/c:v", namespaces)# or \series.find(".//c:tx/c:v", namespaces)
            series_name = name_elem.text.strip() if name_elem is not None else f"Series_{len(all_series_names)+1}"
            all_series_names.append(series_name)

            # 构建分类到值的映射
            category_value_map = {}
            categories = []
            values = []
            
            # 提取分类
            for pt in series.findall(".//c:cat//c:pt", namespaces):
                v_elem = pt.find("c:v", namespaces)
                if v_elem is not None and v_elem.text:
                    categories.append(v_elem.text.strip())
            
            # 提取数值
            for pt in series.findall(".//c:val//c:pt", namespaces):
                v_elem = pt.find("c:v", namespaces)
                if v_elem is not None and v_elem.text:
                    try:
                        values.append(float(v_elem.text.strip()))
                    except ValueError:
                        values.append(v_elem.text.strip())

            # 建立映射
            for cat, val in zip(categories, values):
                category_value_map[cat] = val
            
            # 保存到数据集
            series_data[series_name] = category_value_map

        # 写入CSV
        with open(output_csv, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            
            # 表头
            writer.writerow(["Category"] + all_series_names)
            
            # 数据行
            for category in ordered_categories:
                row = [category]
                for series_name in all_series_names:
                    row.append(series_data[series_name].get(category, "N/A"))
                writer.writerow(row)

        print(f"文件已生成：{output_csv}")
        return True

    except Exception as e:
        print(f"处理失败：{str(e)}")
        return False

# 使用示例
if __name__ == "__main__":
    xml_path = "test.xml"
    csv_path = "test.csv"
    generate_ordered_output(xml_path, csv_path)
