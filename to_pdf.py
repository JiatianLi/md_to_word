import os
import sys
import shutil
import pypandoc
import tempfile

# ==== 可配置部分 ====
if len(sys.argv) > 1:
    DOCS_DIR = sys.argv[1]
else:
    DOCS_DIR = r"D:\docs\gcs_doc\Inbox\doc_edit\docs"  # 默认路径

OUTPUT_PDF = os.path.join(DOCS_DIR, "knowledgebase.pdf")
CHINESE_FONT = "Microsoft YaHei"

# ==== 核心修复 ====
def fix_image_paths(md_text, current_md_file_path):
    """修复图片路径"""
    lines = md_text.splitlines()
    fixed_lines = []
    current_md_dir = os.path.dirname(current_md_file_path)
    for line in lines:
        if "![" in line and "](" in line:
            start = line.find("](")
            end = line.find(")", start)
            if start != -1 and end != -1:
                img_rel_path = line[start + 2:end].strip()
                if img_rel_path.lower().startswith(("http://", "https://")):
                    fixed_lines.append(line)
                    continue

                img_abs_path = os.path.abspath(os.path.join(current_md_dir, img_rel_path))
                if not os.path.exists(img_abs_path):
                    print(f"⚠️ 警告：图片文件不存在 → {img_abs_path}")

                img_abs_path = img_abs_path.replace("\\", "/")
                fixed_line = line[:start + 2] + img_abs_path + line[end:]
                fixed_lines.append(fixed_line)
                continue
        fixed_lines.append(line)
    return "\n".join(fixed_lines)


def collect_markdown_files(directory):
    """收集 md 文件"""
    md_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".md"):
                md_files.append(os.path.join(root, file))
    md_files.sort()
    return md_files


def merge_markdown_files(md_files):
    """合并 Markdown"""
    merged_text = ""
    for i, md_file in enumerate(md_files, start=1):
        print(f"📄 合并文件 {i}/{len(md_files)}：{os.path.basename(md_file)}")
        with open(md_file, "r", encoding="utf-8") as f:
            content = f.read()
        fixed_content = fix_image_paths(content, md_file)
        rel_path = os.path.relpath(md_file, DOCS_DIR).replace("\\", "/")
        safe_rel_path = rel_path.replace("_", r"\_")
        merged_text += f"\n\n---\n\n# {safe_rel_path}\n\n{fixed_content}\n"
    return merged_text


def detect_pdf_engine():
    """检测 PDF 引擎"""
    if shutil.which("xelatex"):
        print("✅ 检测到 XeLaTeX，可正常生成中文 PDF。")
        return "xelatex"
    elif shutil.which("wkhtmltopdf"):
        print("⚙️ 使用 wkhtmltopdf。")
        return "wkhtmltopdf"
    else:
        print("❌ 未检测到 PDF 引擎，请安装 xelatex 或 wkhtmltopdf。")
        exit(1)


def convert_to_pdf(merged_text, pdf_engine):
    """转换为 PDF"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".md", mode="w", encoding="utf-8") as tmp_md:
        tmp_md.write(merged_text)
        tmp_md_path = tmp_md.name

    docs_dir_fixed = DOCS_DIR.replace("\\", "/")
    extra_args = [
        "--standalone",
        f"--resource-path={docs_dir_fixed}",
        "--toc",
        "--toc-depth=3",
        "--pdf-engine", pdf_engine,
        "--variable", "geometry:a4paper",
        "--variable", "margin=1in"
    ]
    if pdf_engine == "xelatex":
        extra_args += [
            "--variable", f"mainfont={CHINESE_FONT}",
            "--variable", "CJKmainfont=Microsoft YaHei",
            "--variable", "geometry=margin=1in"
        ]

    print(f"🧩 使用引擎：{pdf_engine}")
    print(f"🧾 输出文件：{OUTPUT_PDF}")

    pypandoc.convert_file(
        tmp_md_path,
        "pdf",
        outputfile=OUTPUT_PDF,
        extra_args=extra_args,
    )
    os.remove(tmp_md_path)
    print(f"✅ PDF 导出完成：{OUTPUT_PDF}")


def main():
    print(f"📂 正在扫描目录：{DOCS_DIR}")
    md_files = collect_markdown_files(DOCS_DIR)
    if not md_files:
        print("❌ 未找到任何.md文件。")
        return

    merged_text = merge_markdown_files(md_files)
    pdf_engine = detect_pdf_engine()
    convert_to_pdf(merged_text, pdf_engine)


if __name__ == "__main__":
    main()
