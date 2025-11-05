import os
import sys
import pypandoc

# ==== 可配置部分 ====
if len(sys.argv) > 1:
    DOCS_DIR = sys.argv[1]
else:
    DOCS_DIR = r"D:\docs\gcs_doc\Inbox\doc_edit\docs"  # 默认路径

OUTPUT_DOCX = os.path.join(DOCS_DIR, "knowledgebase.docx")

# ==== 核心修复：正确解析相对路径和空格 ====
def fix_image_paths(md_text, current_md_file_path):
    """修复图片路径：根据当前md文件位置解析相对路径，处理空格"""
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

                img_abs_path = img_abs_path.replace("/", "\\")
                fixed_line = line[:start + 2] + img_abs_path + line[end:]
                fixed_lines.append(fixed_line)
                continue
        fixed_lines.append(line)
    return "\n".join(fixed_lines)


def collect_markdown_files(directory):
    """递归收集所有.md文件"""
    md_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".md"):
                md_files.append(os.path.join(root, file))
    md_files.sort()
    return md_files


def merge_markdown_files(md_files):
    """合并多个Markdown文件并修复图片路径"""
    merged_text = ""
    for md_file in md_files:
        with open(md_file, "r", encoding="utf-8") as f:
            content = f.read()
        fixed_content = fix_image_paths(content, md_file)
        rel_path = os.path.relpath(md_file, DOCS_DIR)
        merged_text += f"\n\n---\n\n# {rel_path}\n\n{fixed_content}\n"
    return merged_text


def main():
    print(f"📂 正在扫描目录：{DOCS_DIR}")
    md_files = collect_markdown_files(DOCS_DIR)
    if not md_files:
        print("❌ 未找到任何.md文件，请检查目录路径。")
        return

    print(f"✅ 找到 {len(md_files)} 个Markdown文件，正在合并...")
    merged_text = merge_markdown_files(md_files)

    print(f"📝 正在导出Word文件：{OUTPUT_DOCX}")
    extra_args = [
        "--standalone",
        f"--resource-path={DOCS_DIR}"
    ]
    pypandoc.convert_text(
        merged_text,
        "docx",
        format="md",
        outputfile=OUTPUT_DOCX,
        extra_args=extra_args
    )

    print(f"\n✅ 导出完成：{OUTPUT_DOCX}")


if __name__ == "__main__":
    main()
