import os
import urllib.parse
import pypandoc

# ==== 可配置部分 ====
DOCS_DIR = r"D:\docs\gcs_doc\Inbox\doc_edit\docs"  # 确保这是所有md文件的根目录
OUTPUT_DOCX = r"D:\docs\knowledgebase.docx"

# ==== 核心修复：正确解析相对路径和空格 ====
def fix_image_paths(md_text, current_md_file_path):
    """修复图片路径：根据当前md文件位置解析相对路径，处理空格"""
    lines = md_text.splitlines()
    fixed_lines = []
    # 当前md文件所在的文件夹路径（用于解析相对路径）
    current_md_dir = os.path.dirname(current_md_file_path)
    
    for line in lines:
        if "![" in line and "](" in line:
            # 提取图片路径部分（![描述](路径)）
            start = line.find("](")
            end = line.find(")", start)
            if start != -1 and end != -1:
                img_rel_path = line[start + 2:end].strip()
                # 跳过网络图片
                if img_rel_path.lower().startswith(("http://", "https://")):
                    fixed_lines.append(line)
                    continue
                
                # 关键1：根据当前md文件位置，计算图片的绝对路径
                # 例如：当前md在 docs/xxx/ 下，图片是 ../assets/xxx.png → 转换为 docs/assets/xxx.png
                img_abs_path = os.path.abspath(os.path.join(current_md_dir, img_rel_path))
                
                # 检查文件是否真的存在（调试用）
                if not os.path.exists(img_abs_path):
                    print(f"⚠️ 警告：图片文件不存在 → {img_abs_path}")
                
                # 关键2：处理路径中的空格（Windows中需要保留空格，而非编码为%20）
                # 转换为Windows可识别的路径格式（反斜杠）
                img_abs_path = img_abs_path.replace("/", "\\")
                
                # 关键3：生成pandoc能识别的本地路径（不使用file://协议，直接用绝对路径）
                # 格式：D:\docs\...\图片名.png（保留空格）
                fixed_line = line[:start + 2] + img_abs_path + line[end:]
                fixed_lines.append(fixed_line)
                continue
        # 非图片行直接保留
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
        # 传入当前md文件的完整路径，用于解析相对路径
        fixed_content = fix_image_paths(content, md_file)
        # 添加文件名作为章节标题
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
    # 关键参数：指定资源根目录，帮助pandoc查找文件
    extra_args = [
        "--standalone",
        f"--resource-path={DOCS_DIR}"  # 资源根目录（与你的DOCS_DIR一致）
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