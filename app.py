import streamlit as st
import tempfile
import zipfile
import shutil
from pathlib import Path
from datetime import datetime
import os

# 导入修改后的业务模块
from main_pipeline_web import (
    crop_fba_pdfs,
    generate_fba_stats,
    fill_detail_base,
    process_mixed_skus,
    generate_customs_docs,
    generate_booking_info
)

st.set_page_config(page_title="AMZ 报关自动化工具", layout="wide")
st.title("📦 AMZ 报关文件一键处理")
st.markdown("上传 FBA 箱单 PDF 文件和可选混装表，系统将自动生成裁切后的 PDF、统计表、报关单、推单舱单等结果。")

# 创建一个临时工作目录
work_dir = Path(tempfile.mkdtemp())
input_dir = work_dir / "input"
temp_dir = work_dir / "temp"
template_dir = work_dir / "templates"
input_dir.mkdir()
temp_dir.mkdir()
template_dir.mkdir()

# 将模板文件复制到临时目录（注意：你需要将模板文件放在项目下的 "templates" 文件夹中）
# 这里假设模板文件与 app.py 同级目录下的 templates 文件夹内
local_template_dir = Path(__file__).parent / "templates"
if local_template_dir.exists():
    for tpl_file in local_template_dir.glob("*.xlsx"):
        shutil.copy(tpl_file, template_dir / tpl_file.name)
else:
    st.error("未找到模板文件夹 'templates'，请确保该文件夹存在并包含三个模板文件。")
    st.stop()

uploaded_pdfs = st.file_uploader(
    "请上传 FBA 箱单 PDF 文件（可多选）",
    type=["pdf"],
    accept_multiple_files=True
)

mixed_file = st.file_uploader(
    "如果存在混装表，请上传（可选）",
    type=["xlsx"]
)

if st.button("🚀 开始处理", type="primary"):
    if not uploaded_pdfs:
        st.error("至少上传一个 PDF 文件")
        st.stop()

    # 保存上传的 PDF 到 input_dir
    for pdf in uploaded_pdfs:
        with open(input_dir / pdf.name, "wb") as f:
            f.write(pdf.getbuffer())

    if mixed_file:
        with open(input_dir / "混装表.xlsx", "wb") as f:
            f.write(mixed_file.getbuffer())

    # 执行流水线
    try:
        with st.status("流水线运行中，请稍候...", expanded=True) as status:
            st.write("1/6 裁剪 PDF ...")
            crop_fba_pdfs(input_dir, temp_dir)

            stats_path = temp_dir / "统计结果.xlsx"
            st.write("2/6 统计箱单...")
            generate_fba_stats(input_dir, stats_path)

            detail_filled = temp_dir / "填充好的产品明细.xlsx"
            st.write("3/6 填充基础明细...")
            fill_detail_base(stats_path, template_dir / "产品明细模板.xlsx", detail_filled)

            final_detail = detail_filled
            if (input_dir / "混装表.xlsx").exists():
                st.write("4/6 处理混装信息...")
                final_detail = temp_dir / "处理后_产品明细.xlsx"
                process_mixed_skus(input_dir / "混装表.xlsx", detail_filled, final_detail)
            else:
                st.write("4/6 未检测到混装表，跳过...")

            st.write("5/6 生成报关单...")
            generate_customs_docs(final_detail, template_dir / "报关文件模版.xlsx", temp_dir)

            st.write("6/6 生成推单舱单...")
            booking_out = temp_dir / "推单舱单.xlsx"
            generate_booking_info(final_detail, template_dir / "推单订仓单导入模板.xlsx", booking_out)

            status.update(label="处理完成！", state="complete")

        # 打包所有生成的文件
        zip_path = work_dir / "结果.zip"
        with zipfile.ZipFile(zip_path, 'w') as zf:
            for f in temp_dir.rglob("*"):
                if f.is_file():
                    zf.write(f, arcname=f.relative_to(temp_dir))

        with open(zip_path, "rb") as f:
            st.download_button(
                label="📥 下载所有结果 (ZIP)",
                data=f,
                file_name=f"报关结果_{datetime.now().strftime('%Y%m%d%H%M%S')}.zip",
                mime="application/zip"
            )
        st.success("✅ 所有步骤执行成功！")
    except Exception as e:
        st.error(f"处理失败: {e}")
        st.exception(e)
    finally:
        # 可选：保留临时目录用于调试，正式环境可删除
        # shutil.rmtree(work_dir, ignore_errors=True)
        pass