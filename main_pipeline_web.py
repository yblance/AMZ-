import os
import shutil
import re
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter

import fitz  # PyMuPDF
import pdfplumber
from PIL import Image


# ==========================================
# 辅助函数：固化工作簿中的所有公式为数值
# ==========================================
def solidify_workbook_formulas(file_path):
    """
    将 Excel 文件中所有公式单元格替换为计算结果（需要先用 data_only=True 加载）。
    注意：此方法要求文件中已经存在计算后的缓存值（即之前用 Excel 保存过）。
    为了彻底避免依赖，我们在实际业务中尽量不创建公式，而是直接写值。
    这个函数保留备用，但在本修改版中不会调用，因为我们已经直接写值了。
    """
    pass  # 不再使用


# ==========================================
# 模块 1: FBA 箱单裁剪
# ==========================================
def crop_fba_pdfs(input_folder: Path, output_folder: Path):
    print("\n--- [1/6] 开始执行 FBA 箱单裁剪 ---")
    coordinates = [0, 0, 1250, 1250]
    processed_count = 0
    cropped_pdfs_dir = output_folder / "裁切好的PDF文件"
    cropped_pdfs_dir.mkdir(parents=True, exist_ok=True)

    for file_name in os.listdir(input_folder):
        if not file_name.lower().endswith('.pdf'):
            continue
        pdf_path = input_folder / file_name
        extracted_folder = output_folder / f"temp_images_{pdf_path.stem}"
        extracted_folder.mkdir(exist_ok=True)

        try:
            pdf_doc = fitz.open(pdf_path)
            for page_num in range(pdf_doc.page_count):
                page = pdf_doc.load_page(page_num)
                rect = fitz.Rect(coordinates)
                pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72), clip=rect)
                img_path = extracted_folder / f"{pdf_path.stem}_{page_num + 1}.png"
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                cropped = img.crop((coordinates[0], coordinates[1], coordinates[2], coordinates[3]))
                cropped.save(img_path)
            pdf_doc.close()

            new_pdf_path = cropped_pdfs_dir / f"new-{file_name}"
            pdf_doc_new = fitz.open()
            image_files = sorted([f for f in os.listdir(extracted_folder) if f.lower().endswith('.png')])
            for img_file in image_files:
                img_full = extracted_folder / img_file
                img_rgb = Image.open(img_full).convert('RGB')
                pixmap = fitz.Pixmap(str(img_full))  # 注意：fitz.Pixmap 支持直接读图片
                pdf_page = pdf_doc_new.new_page(width=img_rgb.width, height=img_rgb.height)
                pdf_page.insert_image(fitz.Rect(0, 0, img_rgb.width, img_rgb.height), pixmap=pixmap)
            pdf_doc_new.save(new_pdf_path, deflate=True)
            pdf_doc_new.close()
            processed_count += 1
            shutil.rmtree(extracted_folder)
        except Exception as e:
            print(f"⚠️ 裁切 {file_name} 失败: {e}")
    print(f"✅ FBA箱单裁剪完成，成功处理 {processed_count} 个文件")


# ==========================================
# 模块 2: FBA 箱单统计
# ==========================================
def generate_fba_stats(input_folder: Path, output_excel_path: Path):
    print("\n--- [2/6] 开始执行 FBA 箱单统计 ---")
    df_data = []

    for file_name in sorted(os.listdir(input_folder)):
        if not file_name.lower().endswith('.pdf'):
            continue
        pdf_file = input_folder / file_name
        fba_warehouse = None

        with pdfplumber.open(pdf_file) as pdf:
            if pdf.pages:
                text = pdf.pages[0].extract_text()
                if text:
                    for line in text.split('\n'):
                        if '-' in line and 'Created' in line:
                            fba_warehouse = line[line.find('-') + 1:line.find('Created')].strip()
                            break

        sku_data = {}
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                lines = text.split('\n')
                idx = 0
                while idx < len(lines):
                    line = lines[idx]
                    if 'Single SKU' in line:
                        if idx + 1 < len(lines):
                            sku_name = lines[idx + 1].strip()
                            additional_data = lines[idx + 3].strip() if idx + 3 < len(lines) else ""
                            if sku_name in sku_data:
                                sku_data[sku_name]['Count'] += 1
                                if additional_data:
                                    sku_data[sku_name]['Additional Data'] = additional_data
                            else:
                                sku_data[sku_name] = {
                                    'File Name': file_name,
                                    'SKU Name': sku_name,
                                    'Additional Data': additional_data,
                                    'Count': 1,
                                    'FBA Warehouse': fba_warehouse
                                }
                            idx += 1
                    elif 'Mixed SKUs' in line:
                        sku_name = 'Mixed SKUs'
                        additional_items = []
                        j = idx + 1
                        while j < len(lines) and 'SKU' not in lines[j]:
                            additional_items.append(lines[j].strip())
                            j += 1
                        additional_data = "、".join([x for x in additional_items if x])
                        if sku_name in sku_data:
                            sku_data[sku_name]['Count'] += 1
                            if sku_data[sku_name]['Additional Data']:
                                sku_data[sku_name]['Additional Data'] += "、" + additional_data
                            else:
                                sku_data[sku_name]['Additional Data'] = additional_data
                        else:
                            sku_data[sku_name] = {
                                'File Name': file_name,
                                'SKU Name': sku_name,
                                'Additional Data': additional_data,
                                'Count': 1,
                                'FBA Warehouse': fba_warehouse
                            }
                        idx = j
                    else:
                        idx += 1

        df_data.extend(list(sku_data.values()))

    if not df_data:
        print("⚠️ 未提取到统计数据")
        return

    df = pd.DataFrame(df_data)

    def combine_additional(series):
        items = set()
        for x in series:
            if pd.notna(x) and str(x).strip():
                for i in str(x).split('、'):
                    if i.strip():
                        items.add(i.strip())
        return '、'.join(sorted(items))

    sku_grouped = df.groupby('SKU Name').agg({
        'Count': 'sum',
        'Additional Data': combine_additional
    }).reset_index()

    warehouse_grouped = df.groupby('FBA Warehouse').agg({
        'Count': 'sum',
        'File Name': lambda x: ', '.join(x.unique())
    }).reset_index()
    warehouse_grouped.columns = ['FBA 仓库', '总箱数', '箱单文件']

    new_df = pd.DataFrame({
        'SKU': sku_grouped['SKU Name'],
        'Total Count': sku_grouped['Count'],
        'Additional Data': sku_grouped['Additional Data']
    })

    with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='详细数据')
        new_df.to_excel(writer, index=False, sheet_name='详细数据', startcol=len(df.columns) + 1)
        warehouse_grouped.to_excel(writer, index=False, sheet_name='仓库统计')
    print(f"✅ 统计完成，已生成: {output_excel_path.name}")

# ==========================================
# 模块 3 & 4 (究极融合版): 纯 Python 双引擎匹配 (防错加强版)
# ==========================================
def process_data_pure_python(stats_file, mixed_file, template_file, output_path):
    print("\n--- [3&4/6] 启动纯 Python 计算引擎 (展开混装 & 双重匹配) ---")
    
    # 1. 读取基础统计数据
    df_stats = pd.read_excel(stats_file, sheet_name='详细数据')
    
    # 2. 读取混装映射
    mix_mapping = {}
    if mixed_file.exists():
        wb_mixed = pd.read_excel(mixed_file, sheet_name=1, header=None) 
        skus = wb_mixed.iloc[5:14, 0].dropna().tolist()
        
        for col_idx in range(1, wb_mixed.shape[1]):
            code = wb_mixed.iloc[15, col_idx]
            if pd.notna(code) and str(code).strip():
                code_str = str(code).strip()
                items = []
                for i, sku in enumerate(skus):
                    qty = wb_mixed.iloc[5 + i, col_idx]
                    if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                        items.append({'SKU Name': sku, 'Per_Box_Qty': qty})
                if items:
                    mix_mapping[code_str] = items

    # 3. 展开 Mixed SKUs 行
    expanded_rows = []
    for _, row in df_stats.iterrows():
        sku_name = str(row.get('SKU Name', '')).strip()
        
        if sku_name == 'Mixed SKUs':
            additional_data = str(row.get('Additional Data', ''))
            codes = [c.strip() for c in additional_data.split('、') if c.strip()]
            box_count = row.get('Count', 1)
            
            matched_items = []
            for code in codes:
                if code in mix_mapping:
                    matched_items.extend(mix_mapping[code])
            
            if matched_items:
                fraction_val = round(len([c for c in codes if c in mix_mapping]) / len(matched_items), 1)
                for item in matched_items:
                    new_row = row.copy()
                    new_row['SKU Name'] = item['SKU Name']
                    new_row['总数'] = box_count * item['Per_Box_Qty'] 
                    new_row['箱数'] = fraction_val 
                    expanded_rows.append(new_row)
            else:
                expanded_rows.append(row)
        else:
            row['总数'] = row['Count'] 
            row['箱数'] = row['Count']
            expanded_rows.append(row)

    df_expanded = pd.DataFrame(expanded_rows)

    # ==========================================
    # 🚀 核心一：读取【商品管理】匹配 SKU 资料
    # ==========================================
    print("正在读取《商品管理》资料库并进行智能匹配...")
    df_master = pd.read_excel(template_file, sheet_name='商品管理')

    # 暴力清除表头前后的隐藏空格，防止由于 Excel 输入不规范导致报错
    df_master.columns = df_master.columns.astype(str).str.strip()
    df_expanded.columns = df_expanded.columns.astype(str).str.strip()

    # 智能定位 SKU 列名
    possible_sku_names = ['SKU Name', 'SKU', '商品货号', '货号', '商品SKU', 'ASIN', '商品编码']
    master_sku_col = next((col for col in possible_sku_names if col in df_master.columns), None)

    if master_sku_col is None:
        raise ValueError(f"❌ 匹配失败！在《商品管理》中找不到SKU列。\n🤔 系统读到的表头是: {list(df_master.columns)}")
    
    if master_sku_col != 'SKU Name':
        df_master = df_master.rename(columns={master_sku_col: 'SKU Name'})

    # 第一次匹配：商品信息
    df_final = pd.merge(df_expanded, df_master, on='SKU Name', how='left')

    # ==========================================
    # 🚀 核心二：读取【FBA仓库代码表】匹配地址资料
    # ==========================================
    print("正在读取《FBA仓库代码表》匹配仓库地址...")
    try:
        df_warehouse = pd.read_excel(template_file, sheet_name='FBA仓库代码表')
        df_warehouse.columns = df_warehouse.columns.astype(str).str.strip()

        # 智能定位 仓库代码 列名
        possible_wh_names = ['FBA Warehouse', 'FBA', '仓库代码', '仓库代码/FBA code', 'FBA Code', '目的仓']
        master_wh_col = next((col for col in possible_wh_names if col in df_warehouse.columns), None)

        if master_wh_col:
            if master_wh_col != 'FBA Warehouse':
                df_warehouse = df_warehouse.rename(columns={master_wh_col: 'FBA Warehouse'})
            
            # 去重：防止仓库表里有重复项导致数据行数翻倍
            df_warehouse = df_warehouse.drop_duplicates(subset=['FBA Warehouse'])
            # 第二次匹配：仓库信息
            df_final = pd.merge(df_final, df_warehouse, on='FBA Warehouse', how='left')
        else:
            print(f"⚠️ 未找到仓库代码列，仓库表头为: {list(df_warehouse.columns)}")
    except Exception as e:
        print(f"⚠️ 《FBA仓库代码表》读取失败，部分地址信息将为空: {e}")

    # ==========================================
    # 纯 Python 自动核算 (数值固化)
    # ==========================================
    df_final['单价'] = pd.to_numeric(df_final['单价'], errors='coerce').fillna(0)
    df_final['总价'] = df_final['总数'] * df_final['单价']
    
    if '单个毛重' in df_final.columns:
        df_final['总毛重'] = df_final['总数'] * pd.to_numeric(df_final['单个毛重'], errors='coerce').fillna(0)
    if '单个净重' in df_final.columns:
        df_final['总净重'] = df_final['总数'] * pd.to_numeric(df_final['单个净重'], errors='coerce').fillna(0)
    if '单个体积' in df_final.columns:
        df_final['体积'] = df_final['总数'] * pd.to_numeric(df_final['单个体积'], errors='coerce').fillna(0)

    # 兜底填充：如果商品或仓库表里没写这些字段，强制给定默认值
    default_values = {
        '国家编码': 'US', '货币编码': 'USD', 
        '离境口岸': '深圳', '起运港': 'SHENZHEN', '境内货源地': '深圳'
    }
    for col, val in default_values.items():
        if col not in df_final.columns:
            df_final[col] = val

    # 导出底层表
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='明细')
        
    print(f"✅ 纯 Python 计算完成，已生成底层明细表: {output_path.name}")
    return df_final

# ==========================================
# 模块 5: 创建报关单（无公式，直接填充数值）
# ==========================================
def generate_customs_docs(info_file_path: Path, template_file_path: Path, output_dir: Path):
    print("\n--- [5/6] 开始生成 AMZ 多份报关单 ---")
    df_all = pd.read_excel(info_file_path, sheet_name='明细')

    thin_side = Side(style='thin')
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    now = datetime.now()
    date_inv = now.strftime('%d-%b-%y')
    date_decl = f"申报日期\n{now.strftime('%Y.%m.%d')}"
    date_contract_ship = (now + timedelta(days=7)).strftime('%Y.%m.%d')

    grouped = df_all.groupby('FBA Warehouse')

    for wh_code, df_group in grouped:
        df_group = df_group.reset_index(drop=True)
        n = len(df_group)
        output_name = f"报关文件{wh_code}.xlsx"
        wb = openpyxl.load_workbook(template_file_path)

        # ---- 发票工作表 ----
        ws_inv = wb['发票']
        ws_inv['G5'] = date_inv
        ws_inv['B7'] = df_group.loc[0, '离境口岸'] if '离境口岸' in df_group else ''
        ws_inv['B8'] = df_group.loc[0, '起运港'] if '起运港' in df_group else ''
        ws_inv['B9'] = df_group.loc[0, '目的地'] if '目的地' in df_group else ''
        ws_inv['C9'] = wh_code
        ws_inv['B10'] = df_group.loc[0, '英文目的地'] if '英文目的地' in df_group else ''
        file_name_full = str(df_group.loc[0, 'File Name'])
        parts = file_name_full.split('-')
        ws_inv['G8'] = parts[2] if len(parts) > 2 else ""

        # 填充发票明细行（直接数值）
        for i in range(n):
            row = 13 + i
            ws_inv.cell(row=row, column=2, value=df_group.loc[i, '商品统称'] if '商品统称' in df_group else '')
            ws_inv.cell(row=row, column=4, value=df_group.loc[i, '总数'] if '总数' in df_group else 0)
            ws_inv.cell(row=row, column=5, value="个")
            unit_price = df_group.loc[i, '单价'] if '单价' in df_group else 0
            total_qty = df_group.loc[i, '总数'] if '总数' in df_group else 0
            ws_inv.cell(row=row, column=6, value=unit_price)
            ws_inv.cell(row=row, column=7, value=total_qty * unit_price)

        total_row_inv = 12 + n
        # 合计行
        ws_inv.cell(row=total_row_inv, column=1, value="合计（TOTAL）")
        ws_inv.cell(row=total_row_inv, column=4, value=f"=SUM(D13:D{12+n})")
        ws_inv.cell(row=total_row_inv, column=7, value=f"=SUM(G13:G{12+n})")
        # 合并明细行第一列（不包含合计行）
        if n > 0:
            ws_inv.merge_cells(start_row=13, start_column=1, end_row=11+n, end_column=1)
        # 应用边框
        for r in range(11, total_row_inv + 1):
            for c in range(1, 8):
                ws_inv.cell(row=r, column=c).border = thin_border

        # ---- 装箱单工作表 ----
        ws_pk = wb['装箱单']
        ws_pk['B10'] = df_group.loc[0, '仓库地址全'] if '仓库地址全' in df_group else ''
        for i in range(n):
            row = 14 + i
            ws_pk.cell(row=row, column=2, value=df_group.loc[i, '商品统称'] if '商品统称' in df_group else '')
            ws_pk.cell(row=row, column=3, value=df_group.loc[i, '总数'] if '总数' in df_group else 0)
            ws_pk.cell(row=row, column=4, value="PCS")
            ws_pk.cell(row=row, column=5, value=df_group.loc[i, '总毛重'] if '总毛重' in df_group else 0)
            ws_pk.cell(row=row, column=6, value=df_group.loc[i, '总净重'] if '总净重' in df_group else 0)
            ws_pk.cell(row=row, column=7, value=df_group.loc[i, '体积'] if '体积' in df_group else 0)

        total_row_pk = 13 + n
        ws_pk.cell(row=total_row_pk, column=1, value="合计（TOTAL）")
        ws_pk.cell(row=total_row_pk, column=3, value=f"=SUM(C14:C{13+n})")
        ws_pk.cell(row=total_row_pk, column=4, value="PCS")
        ws_pk.cell(row=total_row_pk, column=5, value=f"=SUM(E14:E{13+n})")
        ws_pk.cell(row=total_row_pk, column=6, value=f"=SUM(F14:F{13+n})")
        ws_pk.cell(row=total_row_pk, column=7, value=f"=SUM(G14:G{13+n})")
        if n > 0:
            ws_pk.merge_cells(start_row=14, start_column=1, end_row=12+n, end_column=1)
        for r in range(12, total_row_pk + 1):
            for c in range(1, 8):
                ws_pk.cell(row=r, column=c).border = thin_border

        # ---- 报关单工作表 ----
        ws_decl = wb['报关单']
        ws_decl['J3'] = date_decl
        ws_decl['J3'].alignment = Alignment(wrapText=True)
        ws_decl['A7'] = ws_inv['G8'].value
        ws_decl['J7'] = ws_inv['B9'].value
        total_boxes = int(round(df_group['箱数'].sum(), 0)) if '箱数' in df_group else 0
        ws_decl['D9'] = total_boxes
        ws_decl['F9'] = f"=装箱单!E{total_row_pk}"
        ws_decl['H9'] = f"=装箱单!F{total_row_pk}"
        ws_decl['A12'] = ws_inv['C9'].value

        for i in range(n):
            odd_r = 15 + 2*i
            ws_decl.cell(row=odd_r, column=1, value=i+1)
            ws_decl.cell(row=odd_r, column=2, value=df_group.loc[i, 'HS'] if 'HS' in df_group else '')
            ws_decl.cell(row=odd_r, column=3, value=df_group.loc[i, '商品统称'] if '商品统称' in df_group else '')
            ws_decl.cell(row=odd_r, column=4, value=df_group.loc[i, '总数'] if '总数' in df_group else 0)
            ws_decl.cell(row=odd_r, column=5, value="个")
            ws_decl.cell(row=odd_r, column=6, value=df_group.loc[i, '单价'] if '单价' in df_group else 0)
            ws_decl.cell(row=odd_r, column=7, value=df_group.loc[i, '总数'] * df_group.loc[i, '单价'] if '总数' in df_group and '单价' in df_group else 0)
            ws_decl.cell(row=odd_r, column=8, value="USD")
            ws_decl.cell(row=odd_r, column=9, value="中国")
            ws_decl.cell(row=odd_r, column=10, value="美国")
            ws_decl.cell(row=odd_r, column=11, value=df_group.loc[i, '境内货源地'] if '境内货源地' in df_group else '')

            even_r = 16 + 2*i
            ws_decl.cell(row=even_r, column=2, value=df_group.loc[i, '申报要素'] if '申报要素' in df_group else '')
            ws_decl.merge_cells(start_row=even_r, start_column=2, end_row=even_r, end_column=12)

        footer_r1 = 14 + 2*n + 1
        footer_r2 = 14 + 2*n + 2
        ws_decl.cell(row=footer_r1, column=1, value="报关人员                        报关人员证号                                电话            兹申明对以上内容承担如实申报、依法纳税之法律责任")
        ws_decl.merge_cells(start_row=footer_r1, start_column=1, end_row=footer_r1, end_column=9)
        ws_decl.cell(row=footer_r2, column=1, value="申报单位                                                                                        申报单位（签章）")
        ws_decl.merge_cells(start_row=footer_r2, start_column=1, end_row=footer_r2, end_column=9)
        ws_decl.cell(row=footer_r1, column=10, value="海关批注及签章")
        ws_decl.merge_cells(start_row=footer_r1, start_column=10, end_row=footer_r2, end_column=12)
        ws_decl.cell(row=footer_r1, column=10).alignment = Alignment(horizontal='center', vertical='center')

        for r in range(14, footer_r2 + 1):
            for c in range(1, 13):
                ws_decl.cell(row=r, column=c).border = thin_border

        # ---- 合同工作表 ----
        ws_ct = wb['合同']
        ws_ct['H4'] = date_inv
        # 合并明细行第一列（不包含合计行）
        if n > 0:
            ws_ct.merge_cells(start_row=12, start_column=1, end_row=10+n, end_column=1)
        for i in range(n):
            row = 12 + i
            ws_ct.cell(row=row, column=2, value=df_group.loc[i, '商品统称'] if '商品统称' in df_group else '')
            ws_ct.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
            ws_ct.cell(row=row, column=5, value=df_group.loc[i, '总数'] if '总数' in df_group else 0)
            ws_ct.cell(row=row, column=7, value=df_group.loc[i, '单价'] if '单价' in df_group else 0)
            ws_ct.cell(row=row, column=6, value="个")
            ws_ct.cell(row=row, column=8, value=df_group.loc[i, '总数'] * df_group.loc[i, '单价'] if '总数' in df_group and '单价' in df_group else 0)

        total_row_ct = 11 + n
        ws_ct.cell(row=total_row_ct, column=1, value="合计（TOTAL）")
        ws_ct.merge_cells(start_row=total_row_ct, start_column=1, end_row=total_row_ct, end_column=4)
        ws_ct.cell(row=total_row_ct, column=5, value=f"=SUM(E12:E{11+n})")
        ws_ct.cell(row=total_row_ct, column=6, value="个")
        ws_ct.cell(row=total_row_ct, column=8, value=f"=SUM(H12:H{11+n})")

        ws_ct.cell(row=total_row_ct+2, column=1, value="2. Payment Terms: T/T")
        ws_ct.cell(row=total_row_ct+3, column=1, value="3. Shipment: Prompt ")
        ws_ct.cell(row=total_row_ct+3, column=3, value=date_contract_ship)
        ws_ct.cell(row=total_row_ct+3, column=4, value="前")
        ws_ct.cell(row=total_row_ct+4, column=1, value="4. Package: Total")
        ws_ct.cell(row=total_row_ct+5, column=1, value="5. Shipping port: ")
        ws_ct.cell(row=total_row_ct+5, column=3, value=ws_inv['B8'].value)
        ws_ct.cell(row=total_row_ct+6, column=1, value="6. Destination: ")
        ws_ct.cell(row=total_row_ct+6, column=3, value=ws_inv['B10'].value)
        ws_ct.cell(row=total_row_ct+7, column=1, value="Please sign and return by fax.")

        for r in range(10, total_row_ct + 1):
            for c in range(1, 9):
                ws_ct.cell(row=r, column=c).border = thin_border

        wb.save(output_dir / output_name)
        wb.close()
    print("✅ 多份报关单生成完成！")


# ==========================================
# 模块 6: 创建推单信息表（无公式）
# ==========================================
def generate_booking_info(source_file: Path, template_file: Path, output_file: Path):
    print("\n--- [6/6] 开始生成推单舱单 ---")
    dtype_dic = {
        '申报单位': str,
        '法定单位': str,
        'HS': str,
        '国家编码': str,
        '货币编码': str
    }
    df_source = pd.read_excel(source_file, sheet_name='明细', dtype=dtype_dic)

    def extract_order_no(filename):
        if pd.isna(filename) or str(filename).lower() == 'nan':
            return ""
        parts = str(filename).split("-")
        return parts[2] if len(parts) >= 3 else ""

    df_body = pd.DataFrame()
    df_body['FBA Warehouse'] = df_source['FBA Warehouse']
    df_body['*海外仓订仓单编号'] = df_source['File Name'].apply(extract_order_no)
    df_body['*商品序号'] = df_body.groupby('FBA Warehouse').cumcount() + 1
    df_body['*商品名称'] = df_source['商品统称'] if '商品统称' in df_source else ''
    df_body['*商品货号'] = df_source['SKU Name'] if 'SKU Name' in df_source else ''
    df_body['*商品编码'] = df_source['HS'] if 'HS' in df_source else ''
    df_body['*规格型号'] = df_source['申报要素'] if '申报要素' in df_source else ''
    df_body['*申报数量'] = pd.to_numeric(df_source['总数'], errors='coerce').fillna(0)
    df_body['*申报计量单位'] = df_source['申报单位'] if '申报单位' in df_source else ''
    df_body['*法定计量单位'] = df_source['法定单位'] if '法定单位' in df_source else ''
    single_weight = pd.to_numeric(df_source['单个净重'], errors='coerce').fillna(0) if '单个净重' in df_source else 0
    df_body['*法定数量'] = np.where(
        df_body['*申报计量单位'] != df_body['*法定计量单位'],
        df_body['*申报数量'] * single_weight,
        df_body['*申报数量']
    )
    df_body['*单价'] = pd.to_numeric(df_source['单价'], errors='coerce').fillna(0) if '单价' in df_source else 0
    df_body['*总价'] = df_body['*单价'] * df_body['*申报数量']

    cols_body = ['*海外仓订仓单编号', '*商品序号', '*商品名称', '*商品货号', '*商品编码',
                 '*规格型号', '*申报数量', '*申报计量单位', '*法定计量单位', '*法定数量',
                 '*单价', '*总价']
    df_body_final = df_body[cols_body]

    df_unique_wh = df_source.drop_duplicates(subset=['FBA Warehouse']).copy()
    df_header = pd.DataFrame()
    df_header['*海外仓名称'] = df_unique_wh['FBA Warehouse']
    df_header['*海外仓订仓单编号'] = df_unique_wh['File Name'].apply(extract_order_no)
    df_header['*海外仓地址'] = df_unique_wh['仓库地址全'] if '仓库地址全' in df_unique_wh else ''
    df_header['*收货人国家'] = df_unique_wh['国家编码'] if '国家编码' in df_unique_wh else ''
    df_header['*币制'] = df_unique_wh['货币编码'] if '货币编码' in df_unique_wh else ''
    total_price_map = df_body.groupby('*海外仓订仓单编号')['*总价'].sum().to_dict()
    df_header['*商品总价'] = df_header['*海外仓订仓单编号'].map(total_price_map)
    df_header['*电商平台'] = "亚马逊"
    df_header['*电商平台代码'] = "无"
    cols_header = ['*海外仓名称', '*海外仓订仓单编号', '*海外仓地址', '*商品总价',
                   '*电商平台', '*电商平台代码', '*收货人国家', '*币制']
    df_header_final = df_header[cols_header]

    workbook = openpyxl.load_workbook(template_file)
    text_format_columns = {'*申报计量单位', '*法定计量单位', '*商品编码', '*收货人国家', '*币制', '*海外仓订仓单编号'}

    def write_df_to_sheet(ws, df):
        # 获取表头映射
        header_map = {}
        for col_idx, cell in enumerate(ws[1], start=1):
            if cell.value:
                header_map[str(cell.value).strip()] = col_idx
        start_row = ws.max_row + 1
        for row_idx, row_data in enumerate(df.to_dict('records'), start=start_row):
            for col_name, value in row_data.items():
                if col_name in header_map:
                    col_letter = get_column_letter(header_map[col_name])
                    cell = ws[f'{col_letter}{row_idx}']
                    if pd.isna(value) or str(value).lower() == 'nan':
                        cell.value = ""
                    else:
                        cell.value = value
                    if col_name in text_format_columns:
                        cell.number_format = '@'

    if '订仓单表体信息' in workbook.sheetnames:
        write_df_to_sheet(workbook['订仓单表体信息'], df_body_final)
    if '订仓单表头信息' in workbook.sheetnames:
        write_df_to_sheet(workbook['订仓单表头信息'], df_header_final)

    workbook.save(output_file)
    print(f"✅ 推单舱单生成完毕: {output_file.name}")