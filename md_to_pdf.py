import os
import shutil
import pypandoc
import tempfile

# ==== 可配置部分 ====
DOCS_DIR = r"D:\docs\gcs_doc\Inbox\doc_edit\docs"   # 所有 md 文件的根目录
OUTPUT_PDF = r"D:\docs\knowledgebase.pdf"            # 输出 PDF 文件路径
CHINESE_FONT = "Microsoft YaHei"                     # PDF 中文字体（系统中要存在）
MAX_IMAGE_WIDTH = "0.9\textwidth"  # 图片最大宽度（占页面宽度的90%）
MAX_IMAGE_HEIGHT = "0.8\textheight"  # 图片最大高度（占页面高度的80%）

# ==== 核心修复：增强路径处理和图片尺寸控制 ====
def fix_image_paths(md_text, current_md_file_path):
    """修复图片路径，同时添加尺寸控制避免超出页面"""
    lines = md_text.splitlines()
    fixed_lines = []
    current_md_dir = os.path.dirname(current_md_file_path)

    for line in lines:
        if "![" in line and "](" in line:
            start = line.find("](")
            end = line.find(")", start)
            if start != -1 and end != -1:
                img_rel_path = line[start + 2:end].strip()
                # 跳过网络图片
                if img_rel_path.lower().startswith(("http://", "https://")):
                    # 为网络图片添加尺寸控制（如果没有）
                    if "{" not in line:
                        line = line[:end+1] + f"{{width={MAX_IMAGE_WIDTH}, height={MAX_IMAGE_HEIGHT}, keepaspectratio}}" + line[end+1:]
                    fixed_lines.append(line)
                    continue

                # 转换相对路径为绝对路径
                img_abs_path = os.path.abspath(os.path.join(current_md_dir, img_rel_path))
                if not os.path.exists(img_abs_path):
                    print(f"⚠️ 警告：图片文件不存在 → {img_abs_path}")

                # 处理路径中的特殊字符
                img_abs_path = img_abs_path.replace("\\", "/")
                # 构建带尺寸控制的图片语法
                fixed_line = line[:start + 2] + img_abs_path + line[end:]
                # 仅在没有尺寸设置时添加（避免重复）
                if "{" not in fixed_line:
                    fixed_line = fixed_line[:end+1] + f"{{width={MAX_IMAGE_WIDTH}, height={MAX_IMAGE_HEIGHT}, keepaspectratio}}" + fixed_line[end+1:]
                fixed_lines.append(fixed_line)
                continue
        fixed_lines.append(line)
    return "\n".join(fixed_lines)


def collect_markdown_files(directory):
    """递归收集所有.md文件，并检查文件名合法性"""
    md_files = []
    invalid_chars = r'\/:*?"<>|'  # Windows系统不允许的文件名字符
    
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".md"):
                # 检查文件名是否包含特殊字符
                for c in invalid_chars:
                    if c in file:
                        print(f"⚠️ 警告：文件名包含特殊字符 '{c}'，可能导致错误 → {file}")
                md_files.append(os.path.join(root, file))
    
    # 按文件路径排序（确保合并顺序一致）
    md_files.sort()
    return md_files


def merge_markdown_files(md_files):
    """合并多个Markdown文件并修复图片路径，处理标题中的特殊字符"""
    merged_text = ""
    for i, md_file in enumerate(md_files, start=1):
        print(f"  📄 合并文件 {i}/{len(md_files)}：{os.path.basename(md_file)}")
        with open(md_file, "r", encoding="utf-8") as f:
            content = f.read()
        
        # 修复当前文件中的图片路径
        fixed_content = fix_image_paths(content, md_file)
        
        # 处理标题中的路径（替换反斜杠为正斜杠，避免LaTeX解析错误）
        rel_path = os.path.relpath(md_file, DOCS_DIR).replace("\\", "/")
        # 对标题中的特殊字符进行LaTeX转义（主要处理下划线和反斜杠）
        safe_rel_path = rel_path.replace("_", r"\_")  # 下划线在LaTeX中是特殊字符
        
        # 添加分隔线和文件标题（作为一级标题）
        merged_text += f"\n\n---\n\n# {safe_rel_path}\n\n{fixed_content}\n"
    
    return merged_text


def detect_pdf_engine():
    """检测可用的 PDF 引擎"""
    if shutil.which("xelatex"):
        print("✅ 检测到 XeLaTeX，可正常生成中文 PDF。")
        return "xelatex"
    elif shutil.which("wkhtmltopdf"):
        print("⚙️ 未检测到 XeLaTeX，自动切换为 wkhtmltopdf（轻量引擎）。")
        return "wkhtmltopdf"
    else:
        print("❌ 未检测到任何 PDF 引擎，请安装以下任意一个：\n"
              "  - MiKTeX（含 xelatex）→ https://miktex.org/download\n"
              "  - wkhtmltopdf → https://wkhtmltopdf.org/downloads.html")
        exit(1)


def convert_to_pdf(merged_text, pdf_engine):
    """执行 PDF 转换（带详细日志输出）"""
    print("🚀 正在执行 Pandoc 转换...")

    # 将内容写入临时 md 文件（使用UTF-8编码确保中文正常）
    with tempfile.NamedTemporaryFile(delete=False, suffix=".md", mode="w", encoding="utf-8") as tmp_md:
        tmp_md.write(merged_text)
        tmp_md_path = tmp_md.name

    # 提前处理路径中的反斜杠（避免f-string中出现反斜杠）
    docs_dir_fixed = DOCS_DIR.replace("\\", "/")

    # 构建Pandoc参数（增加全局图片尺寸控制）
    extra_args = [
        "--standalone",
        f"--resource-path={docs_dir_fixed}",  # 使用预处理后的路径
        "--toc",  # 自动生成目录
        "--toc-depth=3",
        "--pdf-engine", pdf_engine,
        "--variable", "geometry:a4paper",  # 使用A4纸
        "--variable", "margin=1in",       # 设置页边距
        "--variable", f"graphicxopts=width={MAX_IMAGE_WIDTH}, height={MAX_IMAGE_HEIGHT}, keepaspectratio",  # 全局图片设置
    ]

    # 如果是 xelatex，加入中文字体设置
    if pdf_engine == "xelatex":
        extra_args += [
            "--variable", f"mainfont={CHINESE_FONT}",
            "--variable", "sansfont=SimHei",
            "--variable", "monofont=Consolas",
            "--variable", "CJKmainfont=Microsoft YaHei",
            "--variable", "geometry=margin=1in"
        ]

    # 输出命令行信息（用于调试）
    print(f"🧩 使用引擎：{pdf_engine}")
    print(f"🧾 输出文件：{OUTPUT_PDF}")

    # 调用 Pandoc 转换
    try:
        pypandoc.convert_file(
            tmp_md_path,
            "pdf",
            outputfile=OUTPUT_PDF,
            extra_args=extra_args,
        )
    except RuntimeError as e:
        print("❌ PDF 生成失败：", e)
        exit(1)
    finally:
        # 清理临时文件
        if os.path.exists(tmp_md_path):
            os.remove(tmp_md_path)

    print(f"\n✅ 导出完成：{OUTPUT_PDF}")


def main():
    print(f"📂 正在扫描目录：{DOCS_DIR}")
    md_files = collect_markdown_files(DOCS_DIR)
    
    if not md_files:
        print("❌ 未找到任何.md文件，请检查目录路径。")
        return

    print(f"✅ 找到 {len(md_files)} 个Markdown文件，正在合并...")
    merged_text = merge_markdown_files(md_files)

    print(f"📝 正在导出 PDF 文件：{OUTPUT_PDF}")
    pdf_engine = detect_pdf_engine()
    convert_to_pdf(merged_text, pdf_engine)


if __name__ == "__main__":
    main()