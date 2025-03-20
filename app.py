import os
import re
import shutil
import zipfile
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st
import pandas as pd
from io import BytesIO

# 加载规则文件
def load_rules(file_path):
    """加载规则段（强制小写验证）"""
    try:
        with open(file_path, 'r') as f:
            raw_rules = {line.strip().lower() for line in f if line.strip()}  # 规则强制小写
            
        mandatory_rules = {'m100-1.2w', 'f1w', 'f2w', 'f3w', 'f4w', 'f5w', 'f6w', 'm100-8w', 'm100-9w'}
        missing = mandatory_rules - raw_rules
        if missing:
            raise ValueError(f"规则文件缺少必须条目：{', '.join(missing)}")
        return raw_rules
    except Exception as e:
        st.error(f"规则加载失败：{str(e)}")
        raise

# 检查图片文件命名并生成报告
def check_and_fix_folders(root_folder, rules):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "严格小写验证报告"
    sheet.append(["文件夹", "问题文件", "错误类型", "正确示例"])

    yellow_fill = PatternFill(start_color="FFFF00", fill_type="solid")
    for cell in sheet[1]:
        cell.fill = yellow_fill

    errors_found = False

    file_pattern = re.compile(
        r'^[a-z\-]+-'    # 关键词段（必须小写）
        r'([a-z0-9.\-]+)'  # 规则段（必须小写）
        r'\.jpg$'        # 扩展名（必须小写）
    )

    report_data = []

    for folder in os.listdir(root_folder):
        folder_path = os.path.join(root_folder, folder)
        if not os.path.isdir(folder_path):
            continue

        errors = []
        all_files = os.listdir(folder_path)
        found_rules = set()

        # 判断是否为有效的 SKU 文件夹
        if len(folder) not in (20, 17):
            errors.append({"文件": "-", "错误": "文件夹命名错误", "正确示例": "长度应为20位或17位"})

        # 检查每个文件
        for file in all_files:
            if not file.endswith('.jpg'):
                if file.lower().endswith('.jpg'):
                    errors.append({"文件": file, "错误": "扩展名大小写错误", "正确示例": "必须使用小写 .jpg"})
                continue

            if file != file.lower():
                errors.append({"文件": file, "错误": "文件名大小写错误", "正确示例": "所有字母必须小写"})
                continue

            match = file_pattern.match(file)
            if not match:
                errors.append({"文件": file, "错误": "命名结构错误", "正确示例": "示例：pipe-shelves-m100-1.2w.jpg"})
                continue

            rule_segment = match.group(1)
            if rule_segment not in rules:
                errors.append({"文件": file, "错误": "未授权规则段", "正确示例": "允许的规则如：m100-1.2w"})
            else:
                found_rules.add(rule_segment)

        # 必须包含的规则段检查
        mandatory_check = {'m100-1.2w', 'f1w', 'f2w', 'f3w', 'f4w', 'f5w', 'f6w', 'm100-8w', 'm100-9w'}
        missing = mandatory_check - found_rules
        if missing:
            errors.append({"文件": ", ".join(missing), "错误": "缺少核心文件", "正确示例": "必须存在这些规则段"})

        if errors:
            errors_found = True
            for error in errors:
                sheet.append([folder, error["文件"], error["错误"], error["正确示例"]])
            os.rename(folder_path, os.path.join(root_folder, f"INVALID_{folder}"))

    if errors_found:
        output = BytesIO()
        wb.save(output)
        return output
    else:
        return None

# Streamlit 部分
st.title("图片命名规则检查工具")
st.write(
    "本工具支持上传一个包含图片文件夹的 `.zip` 文件，解压后检查图片命名规则并生成检查报告。"
)

# 上传规则文件
rules_file = st.file_uploader("上传规则文件 (TXT 格式)", type=["txt"])
if rules_file:
    rule_text = rules_file.read().decode("utf-8")
    rules = {line.strip().lower() for line in rule_text.splitlines() if line.strip()}

# 上传 .zip 文件夹
uploaded_file = st.file_uploader("请上传包含图片的 .zip 文件", type=["zip"])

if uploaded_file and rules_file:
    # 创建临时处理目录
    temp_dir = "temp_uploaded"
    os.makedirs(temp_dir, exist_ok=True)

    # 保存上传的 .zip 文件
    zip_path = os.path.join(temp_dir, "uploaded.zip")
    with open(zip_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # 解压 .zip 文件
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        zip_ref.extractall(temp_dir)
    
    st.success("文件解压完成！检查开始...")
    
    # 执行检查并生成报告
    report = check_and_fix_folders(temp_dir, rules)
    if report:
        st.success("检查完成！发现问题，生成了报告。")
        st.download_button(
            label="下载检查报告",
            data=report.getvalue(),
            file_name="检查报告.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("检查完成！未发现问题。")

    # 清理临时目录
    shutil.rmtree(temp_dir)
