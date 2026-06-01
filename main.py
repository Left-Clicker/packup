#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import sys
try:
    import certifi
    os.environ.setdefault("SSL_CERT_FILE", certifi.where())
    os.environ.setdefault("REQUESTS_CA_BUNDLE", certifi.where())
except Exception:
    pass

import pandas as pd
import argostranslate.package
import argostranslate.settings
import argostranslate.translate
import chardet
import tkinter as tk
from tkinter import simpledialog, filedialog
from langdetect import detect
from pathlib import Path


# In[2]:

APP_NAME = "文档翻译助手"

def get_runtime_base_dir():
    if hasattr(sys, '_MEIPASS'):
        return Path(getattr(sys, '_MEIPASS', Path.cwd()))

    return Path(__file__).resolve().parent

def get_user_lang_package_dir():
    user_home = Path.home()
    if os.name == 'nt':
        local_app_data = os.getenv("LOCALAPPDATA")
        if local_app_data:
            return Path(local_app_data) / "argos-translate" / "packages"

        return user_home / "AppData" / "Local" / "argos-translate" / "packages"

    return user_home / ".local" / "share" / "argos-translate" / "packages"

def get_lang_package_dirs():
    runtime_base_dir = get_runtime_base_dir()
    script_dir = Path(__file__).resolve().parent
    package_dirs = []

    if hasattr(sys, '_MEIPASS'):
        package_dirs.extend([
            runtime_base_dir / "argos-translate" / "packages",
            runtime_base_dir / "packages",
        ])

    package_dirs.extend([
        script_dir / "argos-translate" / "packages",
        script_dir / "packages",
        get_user_lang_package_dir(),
    ])

    unique_dirs = []
    for package_dir in package_dirs:
        if package_dir not in unique_dirs:
            unique_dirs.append(package_dir)

    return unique_dirs

def configure_argos_package_dirs():
    primary_package_dir = get_user_lang_package_dir()
    package_dirs = get_lang_package_dirs()
    for package_dir in package_dirs:
        package_dir.mkdir(parents=True, exist_ok=True)

    argostranslate.settings.package_data_dir = primary_package_dir
    argostranslate.settings.package_dirs = package_dirs
    os.environ["ARGOS_PACKAGES_DIR"] = str(primary_package_dir)
    return primary_package_dir, package_dirs

progress_window = None

def log(message):
    print(message)
    if progress_window:
        progress_window.log(message)

class ProgressWindow:
    def __init__(self, root):
        self.window = tk.Toplevel(root)
        self.window.title(APP_NAME)
        self.window.geometry("640x420")
        self.window.minsize(520, 320)
        self.window.attributes("-topmost", False)
        self.window.protocol("WM_DELETE_WINDOW", self.window.withdraw)

        self.status_var = tk.StringVar(value="准备开始处理...")
        status_label = tk.Label(self.window, textvariable=self.status_var, anchor="w")
        status_label.pack(fill="x", padx=14, pady=(12, 6))

        self.log_text = tk.Text(self.window, height=18, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=14, pady=6)

        self.close_button = tk.Button(self.window, text="处理中...", state="disabled", command=self.window.destroy)
        self.close_button.pack(padx=14, pady=(6, 12))

        self.window.update()

    def log(self, message):
        self.status_var.set(message)
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.window.update_idletasks()
        self.window.update()

    def finish(self, message):
        self.log(message)
        self.close_button.configure(text="关闭", state="normal")
        self.window.protocol("WM_DELETE_WINDOW", self.window.destroy)
        self.window.update()

lang_package_dir, lang_package_dirs = configure_argos_package_dirs()
log(f"语言包下载目录：{lang_package_dir}")
log(f"语言包搜索目录：{', '.join(str(package_dir) for package_dir in lang_package_dirs)}")
translation_cache = {}
installed_languages = []
target_lang = None
en_lang = None
available_packages = []
download_attempted_pairs = set()

def detect_encoding(file_path):
    with open(file_path, 'rb') as f:
        result = chardet.detect(f.read(65536))
    return result['encoding'] or 'utf-8'

def get_translator(source_lang, destination_lang):
    if source_lang is None or destination_lang is None:
        return None

    try:
        return source_lang.get_translation(destination_lang)
    except Exception:
        return None

def refresh_installed_languages():
    global installed_languages, target_lang, en_lang

    installed_languages = argostranslate.translate.get_installed_languages()
    target_lang = next((lang for lang in installed_languages if lang.code == "zh"), None)
    en_lang = next((lang for lang in installed_languages if lang.code == "en"), None)

def load_available_packages():
    global available_packages

    if available_packages:
        return available_packages

    try:
        available_packages = argostranslate.package.get_available_packages()
    except Exception as e:
        log(f"获取在线语言包列表失败：{e}")
        available_packages = []

    return available_packages

def install_translation_package(from_code, to_code):
    package_pair = (from_code, to_code)
    if package_pair in download_attempted_pairs:
        return False

    download_attempted_pairs.add(package_pair)
    package = next(
        (
            package
            for package in load_available_packages()
            if package.from_code == from_code and package.to_code == to_code
        ),
        None,
    )

    if package is None:
        log(f"未找到可下载语言包：{from_code} -> {to_code}")
        return False

    try:
        log(f"正在下载并安装语言包：{from_code} -> {to_code}")
        lang_package_dir.mkdir(parents=True, exist_ok=True)
        argostranslate.package.install_from_path(package.download())
        refresh_installed_languages()
        log(f"已安装语言包：{from_code} -> {to_code}")
        return True
    except Exception as e:
        log(f"下载或安装语言包 {from_code} -> {to_code} 时出错：{e}")
        return False

def translate_text(text):
    text = text.strip()
    if not text:
        return ''

    if text in translation_cache:
        return translation_cache[text]

    try:
        detected_lang = detect(text)
        log(f"检测到语言：{detected_lang}")

        if detected_lang.startswith('zh'):
            translation_cache[text] = text
            return text
        
        source_lang = next((lang for lang in installed_languages if lang.code == detected_lang), None)
        if not source_lang and install_translation_package(detected_lang, "zh"):
            source_lang = next((lang for lang in installed_languages if lang.code == detected_lang), None)

        if not source_lang:
            log(f"未找到源语言 {detected_lang} 的翻译器，跳过翻译。")
            translation_cache[text] = text
            return text

        direct_translator = get_translator(source_lang, target_lang)
        if not direct_translator and install_translation_package(detected_lang, "zh"):
            source_lang = next((lang for lang in installed_languages if lang.code == detected_lang), None)
            direct_translator = get_translator(source_lang, target_lang)

        if direct_translator:
            translated = direct_translator.translate(text)
            log(f"翻译成中文：{text} -> {translated}")
            translation_cache[text] = translated
            return translated

        if not en_lang:
            install_translation_package("en", "zh")

        if install_translation_package(detected_lang, "en"):
            source_lang = next((lang for lang in installed_languages if lang.code == detected_lang), None)

        to_english_translator = get_translator(source_lang, en_lang)
        en_translator = get_translator(en_lang, target_lang)
        if not to_english_translator or not en_translator:
            log(f"未找到 {detected_lang} -> zh 或 {detected_lang} -> en -> zh 的翻译器，跳过翻译。")
            translation_cache[text] = text
            return text

        text_in_english = to_english_translator.translate(text)
        log(f"翻译成英文：{text} -> {text_in_english}")
        translated = en_translator.translate(text_in_english)
        log(f"翻译成中文：{text_in_english} -> {translated}")

        translation_cache[text] = translated
        return translated
    except Exception as e:
        log(f"翻译文本 '{text}' 时出错：{e}")
        translation_cache[text] = text
        return text

def translate_and_create_new_excel(input_file, output_file, columns_to_translate):
    excel_file = pd.ExcelFile(input_file)
    log(f"读取文件：{input_file}")

    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(input_file, sheet_name=sheet_name)
        log(f"处理工作表：{sheet_name}")

        for col in columns_to_translate:
            if col in df.columns:
                translated_col_name = f"{col}_translated"
                df[translated_col_name] = df[col].apply(lambda x: translate_text(str(x)) if pd.notna(x) else '')
                log(f"列 {col} 已翻译")

        df.to_excel(writer, sheet_name=sheet_name, index=False)
        log(f"已处理并保存工作表：{sheet_name}")

    writer.close()
    log(f"已保存新的 Excel 文件：{output_file}")

def initialize_languages():
    refresh_installed_languages()
    load_available_packages()

    if installed_languages:
        log(f"已安装语言：{', '.join(lang.code for lang in installed_languages)}")
    else:
        log("当前没有本地语言包，将在翻译时按需下载。")

    return True

def process_file(filename, folder_path, output_dir, columns_to_translate):
    file_path = os.path.join(folder_path, filename)
    base_filename = os.path.splitext(filename)[0]
    output_filename = f"{base_filename}-翻译结果.xlsx"
    output_file_path = os.path.join(output_dir, output_filename)
    file_ext = os.path.splitext(filename)[1].lower()
    log(f"目标输出文件路径：{output_file_path}")

    if file_ext == '.xlsx':
        log(f"处理 Excel 文件：{file_path}")
        translate_and_create_new_excel(file_path, output_file_path, columns_to_translate)
        log(f"已保存为：{output_file_path}")
        return

    if file_ext == '.csv':
        encoding = detect_encoding(file_path)
        log(f"处理 CSV 文件：{file_path}，检测到的编码：{encoding}")
        file1 = pd.read_csv(file_path, encoding=encoding)
        temp_file_path = os.path.join(output_dir, base_filename + ".xlsx")
        file1.to_excel(temp_file_path, index=False)
        try:
            translate_and_create_new_excel(temp_file_path, output_file_path, columns_to_translate)
            log(f"已保存为：{output_file_path}")
        finally:
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)

def main():
    global progress_window

    root = tk.Tk()
    root.title(APP_NAME)
    root.withdraw()

    folder_path = filedialog.askdirectory(title="选择输入文件夹")
    if not folder_path:
        log("没有选择输入文件夹")
        return

    output_dir = filedialog.askdirectory(title="选择输出文件夹")
    if not output_dir:
        log("没有选择输出文件夹")
        return

    columns_input = simpledialog.askstring("输入列名", "请输入要翻译的列名，多个用英文逗号分隔：", parent=root)
    if not columns_input:
        log("没有输入列名")
        return

    columns_to_translate = [col.strip() for col in columns_input.split(',') if col.strip()]
    if not columns_to_translate:
        log("没有有效列名")
        return

    progress_window = ProgressWindow(root)
    os.makedirs(output_dir, exist_ok=True)

    if not initialize_languages():
        return

    processed_count = 0
    error_count = 0

    for filename in os.listdir(folder_path):
        if not filename.startswith('待翻译-'):
            continue

        try:
            process_file(filename, folder_path, output_dir, columns_to_translate)
            processed_count += 1
        except Exception as e:
            error_count += 1
            log(f"处理文件 {filename} 时出错：{e}")

    if processed_count == 0 and error_count == 0:
        progress_window.finish("没有找到以“待翻译-”开头的 Excel 或 CSV 文件。")
    elif error_count:
        progress_window.finish(
            f"处理完成：成功 {processed_count} 个，失败 {error_count} 个。请查看上方日志了解失败原因。"
        )
    else:
        progress_window.finish(f"处理完成：成功翻译 {processed_count} 个文件。")

    root.mainloop()

if __name__ == "__main__":
    main()

