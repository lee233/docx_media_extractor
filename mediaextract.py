#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import zipfile
import shutil
import tempfile
from xml.etree import ElementTree as ET

def extract_media(docx_path):
    if not os.path.isfile(docx_path):
        print("指定的文件不存在。请确保拖放一个有效的 .docx 文件。")
        return

    if not docx_path.lower().endswith('.docx'):
        print("请选择一个 .docx 文件。")
        return

    # 获取文件目录和文件名（不含扩展名）
    file_dir = os.path.dirname(docx_path)
    file_name = os.path.splitext(os.path.basename(docx_path))[0]

    # 创建输出文件夹路径
    output_folder = os.path.join(file_dir, file_name)
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 创建临时文件夹用于解压
    temp_dir = tempfile.mkdtemp(prefix='docx_extract_')

    try:
        # 解压 .docx 文件（实际上是一个 ZIP 文件）
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        document_xml_path = os.path.join(temp_dir, 'word', 'document.xml')
        rels_xml_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
        media_dir = os.path.join(temp_dir, 'word', 'media')

        if not os.path.exists(document_xml_path):
            print("无法找到 document.xml 文件。")
            return

        if not os.path.exists(rels_xml_path):
            print("无法找到 relationships 文件。")
            return

        if not os.path.exists(media_dir):
            print("无法找到媒体资源。请确保文档中包含媒体文件。")
            return

        # 解析 relationships文件，建立rId到媒体文件的映射
        rels_tree = ET.parse(rels_xml_path)
        rels_root = rels_tree.getroot()
        namespaces_rels = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}

        rId_to_media = {}
        for rel in rels_root.findall('rel:Relationship', namespaces=namespaces_rels):
            rId = rel.get('Id')
            target = rel.get('Target')
            if target.startswith('media/'):
                media_file = os.path.basename(target)
                rId_to_media[rId] = media_file

        # 解析document.xml，按顺序收集媒体文件的rId
        doc_tree = ET.parse(document_xml_path)
        doc_root = doc_tree.getroot()
        namespaces_doc = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        media_rIds_in_order = []
        for blip in doc_root.findall('.//w:drawing//a:blip', namespaces=namespaces_doc):
            embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if embed and embed in rId_to_media:
                media_rIds_in_order.append(rId_to_media[embed])

        if not media_rIds_in_order:
            print("文档中没有找到任何媒体资源。")
            return

        # 遍历媒体文件，按照顺序复制并重命名
        count = 1
        for media_file in media_rIds_in_order:
            src_path = os.path.join(media_dir, media_file)
            if not os.path.exists(src_path):
                print(f"媒体文件 {media_file} 未找到，跳过。")
                continue
            _, ext = os.path.splitext(media_file)
            new_name = f"{count}{ext}"
            dest_path = os.path.join(output_folder, new_name)
            shutil.copy2(src_path, dest_path)
            count += 1

        print(f"所有媒体资源已成功提取到 \"{output_folder}\" 文件夹中。")

    except zipfile.BadZipFile:
        print("文件不是一个有效的 .docx 文件或已损坏。")
    except ET.ParseError as e:
        print(f"XML 解析错误：{e}")
    except Exception as e:
        print(f"发生错误：{e}")
    finally:
        # 删除临时文件夹
        shutil.rmtree(temp_dir)

def main():
    if len(sys.argv) < 2:
        print("请将一个 .docx 文件拖放到此脚本上运行。")
        input("按任意键退出...")
        return

    docx_path = sys.argv[1]
    extract_media(docx_path)
    print("操作完成，窗口即将关闭...")

if __name__ == "__main__":
    main()
