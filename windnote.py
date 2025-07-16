# -*- coding: utf-8 -*-

"""
A note-taking and web article collection tool built with Python and PyQt6.

Features:
- Organize notes in a folder structure.
- Import web articles from URLs, with an advanced mode for sites requiring logins.
- Edit notes in Markdown with a live preview.
- Search and filter notes by title, summary, or favorite status.
- Sort notes by creation/modification date or name.
- Customize themes, fonts, and other settings.
- Export notes to PDF or DOCX formats (requires Pandoc).
- Automatically detects and uses system proxy on Windows for downloading web drivers.

To use this script, please install the required libraries:
pip install PyQt6 PyQt6-WebEngine requests beautifulsoup4 markdownify pypandoc python-docx markdown selenium webdriver-manager
"""
import sys
import os
import re
import json
import shutil
import requests
import pypandoc
import markdown
import time
from datetime import datetime
from urllib.parse import urlparse, parse_qs

if sys.platform == 'win32':
    import winreg

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextEdit, QPushButton, QFileDialog, QMessageBox, QInputDialog,
    QSplitter, QLabel, QFrame, QTreeWidget, QTreeWidgetItem, QMenu,
    QDialog, QFormLayout, QComboBox, QCheckBox, QLineEdit, QColorDialog
)
from PyQt6.QtGui import (
    QFont, QAction, QActionGroup, QDrag, QIcon, QShortcut,
    QKeySequence, QTextCharFormat, QColor
)
from PyQt6.QtCore import Qt, QUrl, QMimeData, QCoreApplication

from PyQt6.QtWebEngineWidgets import QWebEngineView

from bs4 import BeautifulSoup
from markdownify import markdownify as md

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.common.exceptions import WebDriverException, InvalidSessionIdException

from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager

CONFIG_FILE = "config.json"
DEFAULT_NOTES_DIR = "MyNotes"
DEFAULT_IMAGES_DIR = "MyNotes/images"
DRIVER_DIR = "drivers"
CHROME_DRIVER_PATH = os.path.join(DRIVER_DIR, "chromedriver.exe")
EDGE_DRIVER_PATH = os.path.join(DRIVER_DIR, "msedgedriver.exe")

TRANSLATIONS = {
    "window_title": {"中文": "Windnote", "English": "Python Note & Article Organizer"},
    "file_menu": {"中文": "文件", "English": "File"},
    "import_note_menu": {"中文": "导入笔记", "English": "Import Note"},
    "import_url_quick": {"中文": "从URL导入 (快速)", "English": "From URL (Quick)"},
    "import_browser_advanced": {"中文": "从浏览器导入 (高级)", "English": "From Browser (Advanced)"},
    "export_as_menu": {"中文": "导出为", "English": "Export As"},
    "settings_menu": {"中文": "设置", "English": "Settings"},
    "set_image_folder": {"中文": "设置图片文件夹", "English": "Set Image Folder"},
    "font_settings": {"中文": "字体设置", "English": "Font Settings"},
    "set_bold_color": {"中文": "设置加粗颜色", "English": "Set Bold Color"},
    "set_browser_path": {"中文": "设置浏览器路径", "English": "Set Browser Path"},
    "select_browser_menu": {"中文": "选择浏览器", "English": "Select Browser"},
    "theme_menu": {"中文": "主题", "English": "Theme"},
    "language_menu": {"中文": "语言", "English": "Language"},
    "search_placeholder": {"中文": "在此输入以进行搜索...", "English": "Search here..."},
    "filter_all_notes": {"中文": "所有笔记", "English": "All Notes"},
    "filter_favorites_only": {"中文": "只看收藏", "English": "Favorites Only"},
    "filter_by_title": {"中文": "按标题", "English": "By Title"},
    "filter_by_summary": {"中文": "按摘要", "English": "By Summary"},
    "sort_by_label": {"中文": "排序:", "English": "Sort by:"},
    "sort_mod_desc": {"中文": "修改日期 (降序)", "English": "Date Modified (Newest)"},
    "sort_mod_asc": {"中文": "修改日期 (升序)", "English": "Date Modified (Oldest)"},
    "sort_cre_desc": {"中文": "创建日期 (降序)", "English": "Date Created (Newest)"},
    "sort_cre_asc": {"中文": "创建日期 (升序)", "English": "Date Created (Oldest)"},
    "sort_name_asc": {"中文": "名称 (升序)", "English": "Name (A-Z)"},
    "sort_name_desc": {"中文": "名称 (降序)", "English": "Name (Z-A)"},
    "save_note_button": {"中文": "保存笔记", "English": "Save Note"},
    "note_saved_success": {"中文": "笔记 '{note_name}' 已保存。", "English": "Note '{note_name}' has been saved."},
    "import_success": {"中文": "文章 '{title}' 已成功导入！",
                       "English": "Article '{title}' has been imported successfully!"},
    "import_url_dialog_title": {"中文": "导入笔记 (快速)", "English": "Import Note (Quick)"},
    "import_url_dialog_label": {"中文": "请输入网页链接:", "English": "Enter a web page URL:"},
    "import_failed": {"中文": "导入失败", "English": "Import Failed"},
    "error": {"中文": "错误", "English": "Error"},
    "same_name_exists": {"中文": "同名笔记已存在于当前文件夹。",
                         "English": "A note with the same name already exists in this folder."},
    "export_select_note_prompt": {"中文": "请先选择一篇要导出的笔记。", "English": "Please select a note to export."},
    "export_to": {"中文": "导出为 {format}", "English": "Export to {format}"},
    "export_file_type": {"中文": "{format} 文件", "English": "{format} Files"},
    "all_files": {"中文": "所有文件", "English": "All Files"},
    "export_success": {"中文": "笔记已成功导出到:\n{path}", "English": "Note exported successfully to:\n{path}"},
    "export_failed": {"中文": "导出失败", "English": "Export Failed"},
    "export_pandoc_error": {"中文": "导出时发生错误: {e}\n\n请确保已正确安装 Pandoc。",
                            "English": "An error occurred during export: {e}\n\nPlease ensure Pandoc is installed correctly."},
    "select_image_folder_title": {"中文": "选择图片存储文件夹", "English": "Select Image Storage Folder"},
    "image_folder_updated": {"中文": "新图片的存储文件夹已更新为:\n{dir_name}",
                             "English": "Image storage folder updated to:\n{dir_name}"},
    "settings_saved": {"中文": "设置成功", "English": "Settings Saved"},
    "select_browser_exe": {"中文": "选择 {browser} 可执行文件", "English": "Select {browser} Executable"},
    "browser_path_set": {"中文": "{browser} 路径已设置为:\n{path}",
                         "English": "{browser} path has been set to:\n{path}"},
    "new_note": {"中文": "新建笔记", "English": "New Note"},
    "new_folder": {"中文": "新建文件夹", "English": "New Folder"},
    "enter_note_name": {"中文": "请输入笔记名称:", "English": "Enter note name:"},
    "enter_folder_name": {"中文": "请输入文件夹名称:", "English": "Enter folder name:"},
    "rename": {"中文": "重命名", "English": "Rename"},
    "enter_new_name": {"中文": "请输入新名称:", "English": "Enter new name:"},
    "delete": {"中文": "删除", "English": "Delete"},
    "confirm_delete": {"中文": "确认删除", "English": "Confirm Deletion"},
    "confirm_delete_message": {"中文": "您确定要删除 '{item_name}' 吗？",
                               "English": "Are you sure you want to delete '{item_name}'?"},
    "unpin": {"中文": "取消置顶", "English": "Unpin"},
    "pin_to_top": {"中文": "置顶", "English": "Pin to Top"},
    "unfavorite": {"中文": "取消收藏", "English": "Unfavorite"},
    "add_to_favorites": {"中文": "收藏", "English": "Add to Favorites"},
    "edit_summary": {"中文": "编辑摘要", "English": "Edit Summary"},
    "enter_summary": {"中文": "请输入摘要:", "English": "Enter summary:"},
    "created_date_label": {"中文": "创建", "English": "Created"},
    "modified_date_label": {"中文": "修改", "English": "Modified"},
    "date_unavailable": {"中文": "日期信息不可用", "English": "Date information unavailable"},
    "font_settings_title": {"中文": "字体设置", "English": "Font Settings"},
    "chinese_font_label": {"中文": "中文字体:", "English": "Chinese Font:"},
    "english_font_label": {"中文": "英文字体:", "English": "English Font:"},
    "apply_button": {"中文": "应用", "English": "Apply"},
    "cancel_button": {"中文": "取消", "English": "Cancel"},
    "bold_color_title": {"中文": "选择加粗文字的颜色", "English": "Select Color for Bold Text"},
    "bold_color_success": {"中文": "加粗颜色已设置为 {color_name}", "English": "Bold text color set to {color_name}"},
    "advanced_import_title": {"中文": "App专用浏览器模式", "English": "Advanced Import (Browser Mode)"},
    "status_idle": {"中文": "状态：未启动", "English": "Status: Idle"},
    "launch_browser_button": {"中文": "① 启动App专用浏览器", "English": "① Launch Dedicated Browser"},
    "scrape_page_button": {"中文": "② 抓取当前页面", "English": "② Scrape Current Page"},
    "status_running": {"中文": "状态：专用浏览器已启动。\n请在该浏览器中完成登录并打开目标页面。",
                       "English": "Status: Browser running.\nPlease log in and navigate to the target page."},
    "browser_running_button": {"中文": "✓ 专用浏览器已启动", "English": "✓ Browser is Running"},
    "launch_failed": {"中文": "启动失败", "English": "Launch Failed"},
    "scrape_failed": {"中文": "抓取失败", "English": "Scrape Failed"},
    "scrape_error_browser_closed": {"中文": "抓取失败：App专用浏览器窗口已被关闭。请重新启动。",
                                    "English": "Scrape failed: The browser window was closed. Please relaunch it."},
    "network_failure": {"中文": "网络连接失败", "English": "Network Failure"},
    "driver_fallback_message": {
        "中文": "无法在线自动下载驱动程序。\n\n已成功切换到备用模式，使用位于 'drivers' 文件夹下的本地驱动。\n请确保此驱动版本与您的浏览器匹配。",
        "English": "Could not download the web driver automatically.\n\nSuccessfully switched to fallback mode using the local driver in the 'drivers' folder.\nPlease ensure this driver matches your browser version."},
    "driver_failed_title": {"中文": "<b>启动浏览器失败</b>", "English": "<b>Failed to Launch Browser</b>"},
    "driver_failed_desc": {"中文": "<p>自动下载驱动失败，且在 'drivers' 文件夹中未找到备用驱动。</p>",
                           "English": "<p>Automatic driver download failed, and no fallback driver was found in the 'drivers' folder.</p>"},
    "driver_failed_manual_steps": {"中文": "<p><b>请按以下步骤手动解决:</b></p>",
                                   "English": "<p><b>Please follow these steps to resolve the issue:</b></p>"},
    "driver_failed_step1_title": {"中文": "<b>查看您的 {browser_choice} 浏览器版本:</b><br>",
                                  "English": "<b>Check your {browser_choice} browser version:</b><br>"},
    "driver_failed_step1_desc": {
        "中文": "- 复制并在浏览器地址栏打开: <b>{check_version_url}</b><br>- 记下版本号 (例如: 126.0.6478.127)。",
        "English": "- Copy and open this in your browser: <b>{check_version_url}</b><br>- Note the version number (e.g., 126.0.6478.127)."},
    "driver_failed_step2_title": {"中文": "<b>下载对应的驱动程序:</b><br>",
                                  "English": "<b>Download the corresponding driver:</b><br>"},
    "driver_failed_step2_desc": {
        "中文": "- 打开驱动下载页面: <a href='{driver_download_url}'>{driver_download_url}</a><br>- 根据您刚才记下的版本号，下载最匹配的驱动压缩包。",
        "English": "- Go to: <a href='{driver_download_url}'>{driver_download_url}</a><br>- Download the driver that matches your browser version."},
    "driver_failed_step3_title": {"中文": "<b>放置驱动文件:</b><br>", "English": "<b>Place the driver file:</b><br>"},
    "driver_failed_step3_desc": {
        "中文": "- 解压下载的文件，找到其中的 <b>{driver_exe}</b>。<br>- 将这个 .exe 文件放入程序目录下的 <b>drivers</b> 文件夹内。",
        "English": "- Unzip the downloaded file and find <b>{driver_exe}</b>.<br>- Move this .exe file into the <b>drivers</b> folder in the application directory."},
    "driver_failed_step4_title": {"中文": "<b>重启本程序。</b>", "English": "<b>Restart this program.</b>"},
    "language_changed_title": {"中文": "语言已更改", "English": "Language Changed"},
    "restart_to_apply": {"中文": "请重启程序以应用所有语言更改。",
                         "English": "Please restart the application to apply all language changes."},
    "success": {"中文": "成功", "English": "Success"},
    "move_failed": {"中文": "移动失败", "English": "Move Failed"},
    "destination_must_be_folder": {"中文": "目标必须是一个文件夹。", "English": "Destination must be a folder."},
    "destination_exists": {"中文": "目标文件夹已存在同名文件或文件夹。",
                           "English": "A file or folder with the same name already exists in the destination."},
    "name_cannot_be_empty": {"中文": "名称不能为空。", "English": "Name cannot be empty."},
    "rename_exists": {"中文": "同名文件或文件夹已存在。",
                      "English": "A file or folder with the same name already exists."},
    "rename_failed": {"中文": "重命名失败: {e}", "English": "Rename failed: {e}"},
    "tip": {"中文": "提示", "English": "Tip"}
}


def get_system_proxy():
    if sys.platform != 'win32':
        return None
    try:
        internet_settings = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                           r'Software\Microsoft\Windows\CurrentVersion\Internet Settings')
        proxy_enable, _ = winreg.QueryValueEx(internet_settings, 'ProxyEnable')
        if not proxy_enable:
            winreg.CloseKey(internet_settings)
            return None
        proxy_server, _ = winreg.QueryValueEx(internet_settings, 'ProxyServer')
        winreg.CloseKey(internet_settings)
        if proxy_server:
            proxy_url = f"http://{proxy_server}"
            return {'http': proxy_url, 'https': proxy_url}
    except FileNotFoundError:
        return None
    except Exception as e:
        print(f"Error reading system proxy: {e}")
        return None
    return None


THEMES = {
    "Default Light": {
        "style": "QTreeWidget::item:selected { background-color: #dbeafe; } QTreeWidget::item:selected QLabel, QTreeWidget::item:selected QLabel[summary_label=\"true\"], QTreeWidget::item:selected QLabel[dates_label=\"true\"] { background-color: transparent; color: #1f2937; }"},
    "Dark": {
        "style": "QWidget { background-color: #2d2d2d; color: #f0f0f0; } QTreeWidget, QTextEdit, QWebEngineView { background-color: #252525; border: 1px solid #444;} QMenuBar, QMenu { background-color: #2d2d2d; color: #f0f0f0; } QMenuBar::item:selected, QMenu::item:selected { background-color: #4a4a4a; } QPushButton { background-color: #4a4a4a; border: 1px solid #555; padding: 5px; } QPushButton:hover { background-color: #5a5a5a; } QLineEdit, QComboBox { background-color: #4a4a4a; padding: 3px; border: 1px solid #555; } QSplitter::handle { background-color: #444; } QLabel { color: #f0f0f0; } NoteItemWidget QLabel[summary_label=\"true\"] { color: #a0a0a0; } NoteItemWidget QLabel[dates_label=\"true\"] { color: #888; } QTreeWidget::item:selected { background-color: #4a4a4a; } QTreeWidget::item:selected QLabel, QTreeWidget::item:selected QLabel[summary_label=\"true\"], QTreeWidget::item:selected QLabel[dates_label=\"true\"] { background-color: transparent; color: #f0f0f0; }"},
    "Light Blue": {
        "style": "QWidget { background-color: #eaf2f8; color: #333; } QMenu::item:selected { background-color: #cce0ff; } QLineEdit, QComboBox { border: 1px solid #c0c0c0; padding: 2px; } QTreeWidget::item:selected { background-color: #cce0ff; } QTreeWidget::item:selected QLabel, QTreeWidget::item:selected QLabel[summary_label=\"true\"], QTreeWidget::item:selected QLabel[dates_label=\"true\"] { background-color: transparent; color: #1f2937; }"},
    "Green": {
        "style": "QWidget { background-color: #e8f8f5; color: #333; } QMenu::item:selected { background-color: #cceee8; } QLineEdit, QComboBox { border: 1px solid #c0c0c0; padding: 2px; } QTreeWidget::item:selected { background-color: #cceee8; } QTreeWidget::item:selected QLabel, QTreeWidget::item:selected QLabel[summary_label=\"true\"], QTreeWidget::item:selected QLabel[dates_label=\"true\"] { background-color: transparent; color: #1f2937; }"},
    "Yellow": {
        "style": "QWidget { background-color: #fef9e7; color: #333; } QMenu::item:selected { background-color: #fcf2d4; } QLineEdit, QComboBox { border: 1px solid #c0c0c0; padding: 2px; } QTreeWidget::item:selected { background-color: #fcf2d4; } QTreeWidget::item:selected QLabel, QTreeWidget::item:selected QLabel[summary_label=\"true\"], QTreeWidget::item:selected QLabel[dates_label=\"true\"] { background-color: transparent; color: #1f2937; }"},
    "Newspaper": {
        "style": "QWidget { background-color: #fdf5e6; color: #4a3c2a; } QTextEdit, QWebEngineView { background-color: #faf0e0; border: 1px solid #dcd2bf; } QMenu::item:selected { background-color: #f2e8d9; } QLineEdit, QComboBox { border: 1px solid #dcd2bf; padding: 2px; background-color: #faf0e0; } QTreeWidget::item:selected { background-color: #f2e8d9; } QTreeWidget::item:selected QLabel, QTreeWidget::item:selected QLabel[summary_label=\"true\"], QTreeWidget::item:selected QLabel[dates_label=\"true\"] { background-color: transparent; color: #4a3c2a; }"},
    "Cyberpunk": {
        "style": "QWidget { background-color: #0d0221; color: #00f0c0; } QTreeWidget, QTextEdit, QWebEngineView { background-color: #000; border: 1px solid #ff00ff;} QPushButton { background-color: #240046; color: #00f0c0; border: 1px solid #ff00ff; } QMenu::item:selected { background-color: #5a0094; } QLineEdit, QComboBox { background-color: #240046; color: #00f0c0; border: 1px solid #ff00ff; padding: 3px; } QLabel { color: #00f0c0; } NoteItemWidget QLabel[summary_label=\"true\"] { color: #9a9a9a; } NoteItemWidget QLabel[dates_label=\"true\"] { color: #666; } QTreeWidget::item:selected { background-color: #5a0094; } QTreeWidget::item:selected QLabel, QTreeWidget::item:selected QLabel[summary_label=\"true\"], QTreeWidget::item:selected QLabel[dates_label=\"true\"] { background-color: transparent; color: #e0e0e0; }"},
    "Letter": {
        "style": "QWidget { background-color: #f5f5dc; color: #5b4636; } QMenu::item:selected { background-color: #e9e9d0; } QLineEdit, QComboBox { border: 1px solid #c0c0c0; padding: 2px; background-color: #f5f5dc; } QTreeWidget::item:selected { background-color: #e9e9d0; } QTreeWidget::item:selected QLabel, QTreeWidget::item:selected QLabel[summary_label=\"true\"], QTreeWidget::item:selected QLabel[dates_label=\"true\"] { background-color: transparent; color: #5b4636; }"},
}


class SeleniumManager:
    def __init__(self, config):
        self.config = config
        self.driver = None
        self.base_converter = BaseConverter(config['images_dir'])
        self.main_window = None

    def tr(self, key, **kwargs):
        lang = self.config.get('language', '中文')
        template = TRANSLATIONS.get(key, {}).get(lang, TRANSLATIONS.get(key, {}).get('English', f'<{key}>'))
        return template.format(**kwargs)

    def launch_or_get_browser(self):
        if self.driver:
            try:
                _ = self.driver.window_handles
                print("Browser instance is already running.")
                return None
            except (WebDriverException, InvalidSessionIdException):
                print("Old browser instance detected as closed. Launching a new one.")
                self.driver = None

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        proxies = get_system_proxy()
        original_proxies = {}
        if proxies:
            print(f"Applying system proxy: {proxies['http']}")
            for key, value in proxies.items():
                original_proxies[key.upper()] = os.environ.get(key.upper())
                os.environ[key.upper()] = value
        try:
            browser_choice = self.config.get('browser', 'Chrome')
            service = None
            try:
                print(f"Attempting to get {browser_choice} driver online...")
                if browser_choice == 'Edge':
                    driver_path = EdgeChromiumDriverManager().install()
                    service = EdgeService(executable_path=driver_path)
                    options = webdriver.EdgeOptions()
                else:
                    driver_path = ChromeDriverManager().install()
                    service = ChromeService(executable_path=driver_path)
                    options = webdriver.ChromeOptions()
                print("Online driver acquired successfully.")
            except Exception as e:
                print(f"Online driver acquisition failed: {e}")
                print("Switching to offline fallback mode...")
                local_driver_path = EDGE_DRIVER_PATH if browser_choice == 'Edge' else CHROME_DRIVER_PATH
                if os.path.exists(local_driver_path):
                    if self.main_window:
                        QMessageBox.warning(self.main_window, self.tr('network_failure'),
                                            self.tr('driver_fallback_message'))
                    if browser_choice == 'Edge':
                        service = EdgeService(executable_path=local_driver_path)
                        options = webdriver.EdgeOptions()
                    else:
                        service = ChromeService(executable_path=local_driver_path)
                        options = webdriver.ChromeOptions()
                else:
                    if browser_choice == 'Edge':
                        check_version_url = "edge://settings/help"
                        driver_download_url = "https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/"
                        driver_exe = 'msedgedriver.exe'
                    else:
                        check_version_url = "chrome://settings/help"
                        driver_download_url = "https://googlechromelabs.github.io/chrome-for-testing/"
                        driver_exe = 'chromedriver.exe'

                    error_message = (f"{self.tr('driver_failed_title')}"
                                     f"{self.tr('driver_failed_desc')}"
                                     f"{self.tr('driver_failed_manual_steps')}"
                                     f"<ol>"
                                     f"<li>{self.tr('driver_failed_step1_title', browser_choice=browser_choice)}{self.tr('driver_failed_step1_desc', check_version_url=check_version_url)}</li>"
                                     f"<li>{self.tr('driver_failed_step2_title')}{self.tr('driver_failed_step2_desc', driver_download_url=driver_download_url)}</li>"
                                     f"<li>{self.tr('driver_failed_step3_title')}{self.tr('driver_failed_step3_desc', driver_exe=driver_exe)}</li>"
                                     f"<li>{self.tr('driver_failed_step4_title')}</li>"
                                     f"</ol>")
                    return error_message

            profile_key = "app_edge_profile" if browser_choice == 'Edge' else "app_chrome_profile"
            profile_dir = os.path.join(os.getcwd(), profile_key)
            options.add_argument(f"user-data-dir={profile_dir}")
            if browser_choice == 'Edge':
                self.driver = webdriver.Edge(service=service, options=options)
            else:
                self.driver = webdriver.Chrome(service=service, options=options)
            print(f"Successfully launched a dedicated {browser_choice} instance.")
            return None
        except Exception as e:
            return f"An unknown error occurred while launching the browser: {e}"
        finally:
            if proxies:
                print("Cleaning up temporary proxy settings...")
                for key in proxies:
                    env_key = key.upper()
                    if original_proxies[env_key] is None:
                        if env_key in os.environ: del os.environ[env_key]
                    else:
                        os.environ[env_key] = original_proxies[env_key]
            QApplication.restoreOverrideCursor()

    def scrape_current_page(self):
        if not self.driver:
            return None, self.tr('browser_not_connected'), None
        try:
            html = self.driver.page_source
            url = self.driver.current_url
            soup = BeautifulSoup(html, 'html.parser')
            title, content = self.base_converter._process_html(soup, base_url=url)
            return title, content, None
        except (WebDriverException, InvalidSessionIdException):
            self.driver = None
            return None, self.tr('scrape_error_browser_closed'), None
        except Exception as e:
            return None, f"{self.tr('scrape_failed')}: {e}", None

    def quit_browser(self):
        if self.driver:
            try:
                self.driver.quit()
                print("Dedicated browser has been closed.")
            except Exception as e:
                print(f"Error closing browser: {e}")
            finally:
                self.driver = None


class NoteManager:
    def __init__(self, notes_dir, images_dir, config):
        self.notes_dir = notes_dir
        self.images_dir = images_dir
        self.config = config
        self.metadata_file = os.path.join(self.notes_dir, "metadata.json")
        os.makedirs(self.notes_dir, exist_ok=True)
        os.makedirs(self.images_dir, exist_ok=True)
        os.makedirs(DRIVER_DIR, exist_ok=True)
        self.metadata = self._load_metadata()

    def tr(self, key, **kwargs):
        lang = self.config.get('language', '中文')
        template = TRANSLATIONS.get(key, {}).get(lang, TRANSLATIONS.get(key, {}).get('English', f'<{key}>'))
        return template.format(**kwargs)

    def _load_metadata(self):
        if os.path.exists(self.metadata_file):
            with open(self.metadata_file, 'r', encoding='utf-8') as f:
                try:
                    return json.load(f)
                except json.JSONDecodeError:
                    return {}
        return {}

    def _save_metadata(self):
        with open(self.metadata_file, 'w', encoding='utf-8') as f:
            json.dump(self.metadata, f, ensure_ascii=False, indent=4)

    def get_item_metadata(self, path):
        rel_path = os.path.relpath(path, self.notes_dir).replace('\\', '/')
        default_meta = {
            'created_at': datetime.fromtimestamp(os.path.getctime(path)).isoformat(),
            'modified_at': datetime.fromtimestamp(os.path.getmtime(path)).isoformat(),
            'summary': '', 'is_pinned': False, 'is_favorite': False
        }
        meta = self.metadata.get(rel_path, {});
        default_meta.update(meta)
        return default_meta

    def get_note_content(self, path):
        if os.path.exists(path) and os.path.isfile(path):
            with open(path, 'r', encoding='utf-8') as f: return f.read()
        return ""

    def save_note(self, path, content):
        rel_path = os.path.relpath(path, self.notes_dir).replace('\\', '/')
        if rel_path not in self.metadata:
            self.metadata[rel_path] = {'created_at': datetime.now().isoformat(), 'is_pinned': False,
                                       'is_favorite': False}
        self.metadata[rel_path]['modified_at'] = datetime.now().isoformat()
        if not self.metadata[rel_path].get('summary'):
            self.metadata[rel_path]['summary'] = content[:100].replace('\n', ' ') + '...'
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        self._save_metadata()

    def update_summary(self, path, summary):
        rel_path = os.path.relpath(path, self.notes_dir).replace('\\', '/')
        if rel_path in self.metadata:
            self.metadata[rel_path]['summary'] = summary;
            self._save_metadata()

    def toggle_pinned(self, path):
        rel_path = os.path.relpath(path, self.notes_dir).replace('\\', '/')
        if rel_path in self.metadata:
            self.metadata[rel_path]['is_pinned'] = not self.metadata[rel_path].get('is_pinned', False);
            self._save_metadata()

    def toggle_favorite(self, path):
        rel_path = os.path.relpath(path, self.notes_dir).replace('\\', '/')
        if rel_path in self.metadata:
            self.metadata[rel_path]['is_favorite'] = not self.metadata[rel_path].get('is_favorite', False);
            self._save_metadata()

    def create_item(self, parent_dir, name, is_folder=False, content=None):
        path = os.path.join(parent_dir, name)
        if os.path.exists(path): return None
        if is_folder:
            os.makedirs(path)
        else:
            initial_content = content if content is not None else f"# {os.path.splitext(name)[0]}\n"
            self.save_note(path, initial_content)
        return path

    def delete_item(self, path):
        rel_path_prefix = os.path.relpath(path, self.notes_dir).replace('\\', '/')
        if os.path.isdir(path):
            shutil.rmtree(path)
            for p in list(self.metadata.keys()):
                if p.startswith(rel_path_prefix): del self.metadata[p]
        else:
            os.remove(path)
            if rel_path_prefix in self.metadata: del self.metadata[rel_path_prefix]
        self._save_metadata()

    def move_item(self, source_path, dest_dir):
        if not os.path.isdir(dest_dir): return None, self.tr('destination_must_be_folder')
        dest_path = os.path.join(dest_dir, os.path.basename(source_path))
        if os.path.exists(dest_path): return None, self.tr('destination_exists')
        try:
            shutil.move(source_path, dest_path)
            old_rel_path_prefix = os.path.relpath(source_path, self.notes_dir).replace('\\', '/')
            new_rel_path_prefix = os.path.relpath(dest_path, self.notes_dir).replace('\\', '/')
            for p in list(self.metadata.keys()):
                if p == old_rel_path_prefix or p.startswith(old_rel_path_prefix + os.sep):
                    new_p = p.replace(old_rel_path_prefix, new_rel_path_prefix, 1)
                    self.metadata[new_p] = self.metadata.pop(p)
            self._save_metadata();
            return dest_path, None
        except Exception as e:
            return None, f"{self.tr('move_failed')}: {e}"

    def rename_item(self, old_path, new_name):
        if not new_name: return None, self.tr('name_cannot_be_empty')
        parent_dir = os.path.dirname(old_path);
        new_path = os.path.join(parent_dir, new_name)
        if os.path.exists(new_path): return None, self.tr('rename_exists')
        try:
            os.rename(old_path, new_path)
            old_rel_path_prefix = os.path.relpath(old_path, self.notes_dir).replace('\\', '/')
            new_rel_path_prefix = os.path.relpath(new_path, self.notes_dir).replace('\\', '/')
            for p in list(self.metadata.keys()):
                if p == old_rel_path_prefix or p.startswith(old_rel_path_prefix + os.sep):
                    new_p = p.replace(old_rel_path_prefix, new_rel_path_prefix, 1)
                    self.metadata[new_p] = self.metadata.pop(p)
            self._save_metadata();
            return new_path, None
        except Exception as e:
            return None, self.tr('rename_failed', e=e)


class NoteItemWidget(QWidget):
    def __init__(self, path, metadata, tr_func, parent=None):
        super().__init__(parent)
        self.tr = tr_func
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 6, 8, 6)
        layout.setSpacing(3)
        top_line_layout = QHBoxLayout()
        title_text = f"<b>{os.path.basename(os.path.splitext(path)[0])}</b>"
        self.title_label = QLabel(title_text)
        self.title_label.setWordWrap(True)
        icons_text = ""
        if metadata.get('is_pinned'): icons_text += "📌 "
        if metadata.get('is_favorite'): icons_text += "⭐"
        self.icons_label = QLabel(icons_text)
        self.icons_label.setAlignment(Qt.AlignmentFlag.AlignTop)
        top_line_layout.addWidget(self.title_label, 1)
        top_line_layout.addWidget(self.icons_label, 0)
        summary_text = metadata.get('summary', '...')
        self.summary_label = QLabel(summary_text)
        self.summary_label.setProperty("summary_label", True)
        self.summary_label.setWordWrap(True)
        try:
            created_date = datetime.fromisoformat(metadata.get('created_at')).strftime('%Y-%m-%d')
            modified_date = datetime.fromisoformat(metadata.get('modified_at')).strftime('%Y-%m-%d')
            dates_text = f"{self.tr('created_date_label')}: {created_date} | {self.tr('modified_date_label')}: {modified_date}"
        except (ValueError, TypeError):
            dates_text = self.tr('date_unavailable')
        self.dates_label = QLabel(dates_text)
        self.dates_label.setProperty("dates_label", True)
        layout.addLayout(top_line_layout)
        layout.addWidget(self.summary_label)
        layout.addWidget(self.dates_label)
        self.setLayout(layout)


class BaseConverter:
    def __init__(self, images_dir):
        self.images_dir = images_dir
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}

    def _download_image(self, url):
        try:
            parsed_url = urlparse(url);
            qs = parse_qs(parsed_url.query)
            img_format = qs.get('wx_fmt', ['jpeg'])[0] if 'wx_fmt' in qs else url.split('.')[-1].split('?')[0]
            if len(img_format) > 4: img_format = 'jpg'
            filename = f"{datetime.now().strftime('%Y%m%d%H%M%S%f')}.{img_format}"
            filepath = os.path.join(self.images_dir, filename)
            img_response = requests.get(url, headers=self.headers, stream=True, timeout=10)
            img_response.raise_for_status()
            with open(filepath, 'wb') as f:
                for chunk in img_response.iter_content(1024): f.write(chunk)
            return filename
        except Exception as e:
            print(f"Failed to download image: {url}, Error: {e}");
            return None

    def _process_html(self, soup, base_url=""):
        title_tag = soup.find('h1') or soup.find('h2', class_='rich_media_title')
        title = title_tag.get_text(strip=True) if title_tag else "Untitled Article"
        title = re.sub(r'[\\/*?:"<>|]', "", title)
        content_div = soup.find('div', id='js_content') or soup.find('article') or soup.find('main') or soup.body
        if not content_div: raise ValueError("Could not find the main content area of the article.")
        for img_tag in content_div.find_all('img'):
            img_url = img_tag.get('data-src') or img_tag.get('src')
            if not img_url: continue
            if not img_url.startswith(('http://', 'https://')):
                from urllib.parse import urljoin
                img_url = urljoin(base_url, img_url)
            img_name = self._download_image(img_url)
            if img_name:
                img_tag.attrs.clear()
                local_path = os.path.join(os.path.basename(self.images_dir), img_name).replace("\\", "/")
                img_tag['src'] = local_path;
                img_tag['alt'] = "image"
        return title, md(str(content_div), heading_style="ATX", escape_style=True)


class RequestsConverter(BaseConverter):
    def convert_from_url(self, url):
        try:
            response = requests.get(url, headers=self.headers, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')
            title, markdown_content = self._process_html(soup, base_url=url)
            return title, markdown_content, None
        except Exception as e:
            return None, f"Conversion failed: {e}", None


class AdvancedImportDialog(QDialog):
    def __init__(self, selenium_manager, main_window):
        super().__init__(main_window)
        self.selenium_manager = selenium_manager
        self.main_window = main_window
        self.tr = self.main_window.tr
        self.selenium_manager.main_window = main_window
        self.setWindowTitle(self.tr('advanced_import_title'))
        self.setMinimumWidth(350)
        self.layout = QVBoxLayout(self)
        self.status_label = QLabel()
        self.status_label.setStyleSheet("font-weight: bold;")
        self.launch_button = QPushButton()
        self.scrape_button = QPushButton(self.tr('scrape_page_button'))
        self.layout.addWidget(self.status_label)
        self.layout.addWidget(self.launch_button)
        self.layout.addWidget(self.scrape_button)
        self.launch_button.clicked.connect(self.launch_browser)
        self.scrape_button.clicked.connect(self.scrape_page)
        self.update_ui()

    def update_ui(self):
        if self.selenium_manager.driver:
            self.status_label.setText(self.tr('status_running'))
            self.scrape_button.setEnabled(True)
            self.launch_button.setText(self.tr('browser_running_button'))
        else:
            self.status_label.setText(self.tr('status_idle'))
            self.scrape_button.setEnabled(False)
            self.launch_button.setText(self.tr('launch_browser_button'))

    def launch_browser(self):
        error = self.selenium_manager.launch_or_get_browser()
        if error:
            QMessageBox.critical(self, self.tr('launch_failed'), error)
        self.update_ui()

    def scrape_page(self):
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        title, content, error = self.selenium_manager.scrape_current_page()
        QApplication.restoreOverrideCursor()
        if error:
            QMessageBox.critical(self, self.tr('scrape_failed'), error)
            if not self.selenium_manager.driver:
                self.update_ui()
            return
        note_name = f"{title}.md"
        parent_dir = self.main_window.get_selected_dir()
        path = self.main_window.note_manager.create_item(parent_dir, note_name, content=content)
        if not path:
            QMessageBox.warning(self, self.tr('error'), self.tr('same_name_exists'))
        else:
            self.main_window.load_notes_tree()
            QMessageBox.information(self, self.tr('success'), self.tr('import_success', title=title))
            self.accept()


class DraggableTreeWidget(QTreeWidget):
    def __init__(self, parent=None, note_manager=None, main_window=None):
        super().__init__(parent)
        self.note_manager = note_manager;
        self.main_window = main_window
        self.setDragDropMode(self.DragDropMode.InternalMove);
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)

    def startDrag(self, supportedActions):
        item = self.currentItem()
        if item and item.data(0, Qt.ItemDataRole.UserRole) and os.path.isfile(item.data(0, Qt.ItemDataRole.UserRole)):
            mime_data = QMimeData();
            mime_data.setText(item.data(0, Qt.ItemDataRole.UserRole))
            drag = QDrag(self);
            drag.setMimeData(mime_data);
            drag.exec(Qt.DropAction.MoveAction)

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        source_path = event.mimeData().text()
        if not source_path:
            event.ignore()
            return
        target_item = self.itemAt(event.position().toPoint())
        dest_dir = self.note_manager.notes_dir
        if target_item:
            target_path = target_item.data(0, Qt.ItemDataRole.UserRole)
            if target_path:
                dest_dir = target_path if os.path.isdir(target_path) else os.path.dirname(target_path)
        if dest_dir and not (os.path.isdir(source_path) and dest_dir.startswith(source_path)):
            if os.path.dirname(source_path) == dest_dir:
                event.ignore()
                return
            _, error = self.note_manager.move_item(source_path, dest_dir)
            if error:
                QMessageBox.warning(self, self.main_window.tr('move_failed'), error)
                event.ignore()
            else:
                self.main_window.load_notes_tree()
                event.acceptProposedAction()
        else:
            event.ignore()


class FontSettingsDialog(QDialog):
    def __init__(self, config, tr_func, parent=None):
        super().__init__(parent)
        self.config = config;
        self.tr = tr_func
        self.setWindowTitle(self.tr('font_settings_title'));
        layout = QFormLayout(self)
        self.chinese_font_combo = QComboBox();
        self.english_font_combo = QComboBox()
        self.chinese_font_combo.addItems(["宋体", "黑体", "楷体", "仿宋", "微软雅黑"])
        self.english_font_combo.addItems(
            ["Arial", "Times New Roman", "Verdana", "Courier New", "Georgia", "Comic Sans MS"])
        self.chinese_font_combo.setCurrentText(self.config.get('chinese_font', '宋体'))
        self.english_font_combo.setCurrentText(self.config.get('english_font', 'Arial'))
        layout.addRow(self.tr('chinese_font_label'), self.chinese_font_combo);
        layout.addRow(self.tr('english_font_label'), self.english_font_combo)
        buttons = QHBoxLayout();
        ok_button = QPushButton(self.tr('apply_button'));
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton(self.tr('cancel_button'));
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(ok_button);
        buttons.addWidget(cancel_button);
        layout.addRow(buttons)

    def get_selected_fonts(self):
        return {'chinese_font': self.chinese_font_combo.currentText(),
                'english_font': self.english_font_combo.currentText()}


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_note_path = None
        self.config = self._load_app_config()
        self.note_manager = NoteManager(self.config['notes_dir'], self.config['images_dir'], self.config)
        self.requests_converter = RequestsConverter(self.config['images_dir'])
        self.selenium_manager = SeleniumManager(self.config)
        self.selenium_manager.main_window = self
        self.init_ui()
        self.apply_styles()
        self.load_notes_tree()

    def tr(self, key, **kwargs):
        lang = self.config.get('language', '中文')
        template = TRANSLATIONS.get(key, {}).get(lang, TRANSLATIONS.get(key, {}).get('English', f"<{key}>"))
        return template.format(**kwargs)

    def toggle_bold(self):
        cursor = self.note_editor.textCursor()
        if not cursor.hasSelection(): return
        selected_text = cursor.selectedText()
        if selected_text.startswith('**') and selected_text.endswith('**'):
            cursor.insertText(selected_text[2:-2])
        else:
            cursor.insertText(f"**{selected_text}**")

    def toggle_italic(self):
        cursor = self.note_editor.textCursor()
        if not cursor.hasSelection(): return
        selected_text = cursor.selectedText()
        if (selected_text.startswith('*') and selected_text.endswith('*')) and not selected_text.startswith('**'):
            cursor.insertText(selected_text[1:-1])
        elif selected_text.startswith('_') and selected_text.endswith('_'):
            cursor.insertText(selected_text[1:-1])
        else:
            cursor.insertText(f"*{selected_text}*")

    def toggle_underline(self):
        cursor = self.note_editor.textCursor()
        if not cursor.hasSelection(): return
        selected_text = cursor.selectedText()
        if selected_text.startswith('<u>') and selected_text.endswith('</u>'):
            cursor.insertText(selected_text[3:-4])
        else:
            cursor.insertText(f"<u>{selected_text}</u>")

    def set_bold_color(self):
        current_color_hex = self.config.get('bold_color', '#000000')
        initial_color = QColor(current_color_hex)
        color = QColorDialog.getColor(initial=initial_color, parent=self, title=self.tr('bold_color_title'))
        if color.isValid():
            self.config['bold_color'] = color.name()
            self._save_app_config(self.config)
            self.update_preview()
            QMessageBox.information(self, self.tr('success'), self.tr('bold_color_success', color_name=color.name()))

    def import_from_url(self, use_selenium=False):
        if not use_selenium:
            url, ok = QInputDialog.getText(self, self.tr('import_url_dialog_title'), self.tr('import_url_dialog_label'))
            if not (ok and url): return
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            title, content, error = self.requests_converter.convert_from_url(url)
            QApplication.restoreOverrideCursor()
            if error:
                QMessageBox.critical(self, self.tr('import_failed'), error)
            else:
                note_name = f"{title}.md";
                parent_dir = self.get_selected_dir()
                path = self.note_manager.create_item(parent_dir, note_name, content=content)
                if not path:
                    QMessageBox.warning(self, self.tr('error'), self.tr('same_name_exists'))
                else:
                    self.load_notes_tree();
                    QMessageBox.information(self, self.tr('success'), self.tr('import_success', title=title))
        else:
            dialog = AdvancedImportDialog(self.selenium_manager, self)
            dialog.exec()

    def closeEvent(self, event):
        self.selenium_manager.quit_browser()
        self.save_current_note(show_message=False)
        event.accept()

    def _load_app_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
        else:
            config = {}

        defaults = {
            'notes_dir': DEFAULT_NOTES_DIR, 'images_dir': DEFAULT_IMAGES_DIR,
            'chinese_font': '宋体', 'english_font': 'Arial', 'theme': 'Default Light',
            'browser': 'Chrome', 'chrome_binary_path': '', 'edge_binary_path': '',
            'bold_color': '#000000', 'language': '中文'
        }

        is_new_config = not os.path.exists(CONFIG_FILE)
        for key, value in defaults.items():
            if key not in config:
                config[key] = value

        if is_new_config:
            self._save_app_config(config)

        return config

    def _save_app_config(self, config):
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(config, f, ensure_ascii=False, indent=4)

    def init_ui(self):
        self.setWindowTitle(self.tr('window_title'));
        self.setGeometry(100, 100, 1400, 900)
        menubar = self.menuBar();
        file_menu = menubar.addMenu(self.tr('file_menu'));
        import_menu = file_menu.addMenu(self.tr('import_note_menu'))
        import_action = QAction(self.tr('import_url_quick'), self);
        import_action.triggered.connect(lambda: self.import_from_url(False))
        import_selenium_action = QAction(self.tr('import_browser_advanced'), self);
        import_selenium_action.triggered.connect(lambda: self.import_from_url(True))
        import_menu.addAction(import_action);
        import_menu.addAction(import_selenium_action)
        export_menu = file_menu.addMenu(self.tr('export_as_menu'))
        export_pdf_action = QAction("PDF", self);
        export_pdf_action.triggered.connect(lambda: self.export_note('pdf'))
        export_docx_action = QAction("Word (.docx)", self);
        export_docx_action.triggered.connect(lambda: self.export_note('docx'))
        export_menu.addAction(export_pdf_action);
        export_menu.addAction(export_docx_action)
        settings_menu = menubar.addMenu(self.tr('settings_menu'))
        set_img_dir_action = QAction(self.tr('set_image_folder'), self);
        set_img_dir_action.triggered.connect(self.set_image_directory)
        set_font_action = QAction(self.tr('font_settings'), self);
        set_font_action.triggered.connect(self.open_font_settings)
        set_bold_color_action = QAction(self.tr('set_bold_color'), self)
        set_bold_color_action.triggered.connect(self.set_bold_color)
        set_browser_path_action = QAction(self.tr('set_browser_path'), self);
        set_browser_path_action.triggered.connect(self.set_browser_path)

        browser_menu = settings_menu.addMenu(self.tr('select_browser_menu'))
        browser_group = QActionGroup(self);
        browser_group.setExclusive(True)
        chrome_action = QAction("Chrome", self, checkable=True);
        chrome_action.triggered.connect(lambda: self.set_browser('Chrome'))
        edge_action = QAction("Edge", self, checkable=True);
        edge_action.triggered.connect(lambda: self.set_browser('Edge'))
        browser_group.addAction(chrome_action);
        browser_group.addAction(edge_action)
        browser_menu.addAction(chrome_action);
        browser_menu.addAction(edge_action)
        if self.config.get('browser') == 'Edge':
            edge_action.setChecked(True)
        else:
            chrome_action.setChecked(True)

        theme_menu = settings_menu.addMenu(self.tr('theme_menu'))
        self.theme_group = QActionGroup(self);
        self.theme_group.setExclusive(True)
        for theme_name in THEMES.keys():
            theme_action = QAction(theme_name, self, checkable=True);
            theme_action.triggered.connect(lambda checked, name=theme_name: self.change_theme(name))
            if self.config.get('theme') == theme_name: theme_action.setChecked(True)
            theme_menu.addAction(theme_action)
            self.theme_group.addAction(theme_action)

        language_menu = settings_menu.addMenu(self.tr('language_menu'))
        language_group = QActionGroup(self)
        language_group.setExclusive(True)
        zh_action = QAction("中文", self, checkable=True)
        zh_action.triggered.connect(lambda: self.change_language("中文"))
        language_group.addAction(zh_action)
        language_menu.addAction(zh_action)
        en_action = QAction("English", self, checkable=True)
        en_action.triggered.connect(lambda: self.change_language("English"))
        language_group.addAction(en_action)
        language_menu.addAction(en_action)
        if self.config.get('language') == 'English':
            en_action.setChecked(True)
        else:
            zh_action.setChecked(True)

        settings_menu.addActions([set_img_dir_action, set_font_action, set_bold_color_action, set_browser_path_action])

        central_widget = QWidget();
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget);
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        left_panel = QFrame();
        left_layout = QVBoxLayout(left_panel)
        filter_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(self.tr('search_placeholder'))
        self.search_input.textChanged.connect(self.load_notes_tree)
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(
            [self.tr('filter_all_notes'), self.tr('filter_favorites_only'), self.tr('filter_by_title'),
             self.tr('filter_by_summary')])
        self.filter_combo.currentIndexChanged.connect(self.load_notes_tree)
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(self.filter_combo)
        left_layout.addLayout(filter_layout)
        sort_layout = QHBoxLayout();
        sort_layout.addWidget(QLabel(self.tr('sort_by_label')))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(
            [self.tr('sort_mod_desc'), self.tr('sort_mod_asc'), self.tr('sort_cre_desc'), self.tr('sort_cre_asc'),
             self.tr('sort_name_asc'), self.tr('sort_name_desc')])
        self.sort_combo.currentIndexChanged.connect(self.load_notes_tree)
        sort_layout.addWidget(self.sort_combo);
        left_layout.addLayout(sort_layout)
        self.notes_tree_widget = DraggableTreeWidget(note_manager=self.note_manager, main_window=self)
        self.notes_tree_widget.setHeaderHidden(True);
        self.notes_tree_widget.itemClicked.connect(self.on_note_selected)
        self.notes_tree_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.notes_tree_widget.customContextMenuRequested.connect(self.show_tree_context_menu)
        left_layout.addWidget(self.notes_tree_widget)
        right_splitter = QSplitter(Qt.Orientation.Horizontal)
        editor_panel = QFrame();
        editor_layout = QVBoxLayout(editor_panel)
        self.note_editor = QTextEdit()
        self.note_editor.document().contentsChanged.connect(self.update_preview)
        editor_layout.addWidget(self.note_editor)
        preview_panel = QFrame();
        preview_layout = QVBoxLayout(preview_panel)
        self.preview_area = QWebEngineView();
        preview_layout.addWidget(self.preview_area)
        right_splitter.addWidget(editor_panel);
        right_splitter.addWidget(preview_panel);
        right_splitter.setSizes([700, 700])
        main_splitter.addWidget(left_panel);
        main_splitter.addWidget(right_splitter);
        main_splitter.setSizes([400, 1000])
        main_layout.addWidget(main_splitter)
        save_button = QPushButton(self.tr('save_note_button'));
        save_button.clicked.connect(self.save_current_note)
        main_layout.addWidget(save_button, 0, Qt.AlignmentFlag.AlignRight)
        QShortcut(QKeySequence("Ctrl+B"), self, self.toggle_bold)
        QShortcut(QKeySequence("Ctrl+I"), self, self.toggle_italic)
        QShortcut(QKeySequence("Ctrl+U"), self, self.toggle_underline)

    def apply_styles(self):
        theme_name = self.config.get('theme', 'Default Light')
        theme_data = THEMES.get(theme_name, THEMES['Default Light'])
        stylesheet = theme_data.get("style", "")
        eng_font = self.config.get('english_font', 'Arial')
        cn_font = self.config.get('chinese_font', '宋体')
        app_font = QFont(eng_font)
        QApplication.instance().setFont(app_font)
        self.setStyleSheet(stylesheet)
        self.note_editor.setFont(QFont(cn_font, 12))
        self.update_preview()

    def change_theme(self, theme_name):
        self.config['theme'] = theme_name;
        self._save_app_config(self.config)
        self.apply_styles()

    def change_language(self, lang):
        if self.config.get('language', '中文') == lang:
            return
        self.config['language'] = lang
        self._save_app_config(self.config)
        QMessageBox.information(self, self.tr('language_changed_title'), self.tr('restart_to_apply'))

    def set_browser(self, browser_name):
        self.config['browser'] = browser_name;
        self._save_app_config(self.config)
        self.selenium_manager.config['browser'] = browser_name

    def load_notes_tree(self):
        self.notes_tree_widget.clear()
        self._populate_tree(self.notes_tree_widget, self.note_manager.notes_dir)
        self.notes_tree_widget.expandAll()

    def _populate_tree(self, parent_item, path):
        items = os.listdir(path)
        folders = sorted([i for i in items if os.path.isdir(os.path.join(path, i)) and not i.startswith('.')])
        files = [i for i in items if
                 os.path.isfile(os.path.join(path, i)) and not i.startswith('.') and i != 'metadata.json']

        search_text = self.search_input.text().lower()
        filter_index = self.filter_combo.currentIndex()
        filter_options = ['filter_all_notes', 'filter_favorites_only', 'filter_by_title', 'filter_by_summary']
        current_filter_key = filter_options[filter_index]

        file_infos = []
        for f in files:
            full_path = os.path.join(path, f)
            meta = self.note_manager.get_item_metadata(full_path)
            meta['name'] = f
            meta['path'] = full_path
            file_infos.append(meta)

        filtered_infos = []
        is_all_notes_search = current_filter_key in ['filter_all_notes', 'filter_favorites_only']
        if not search_text and current_filter_key == "filter_all_notes":
            filtered_infos = file_infos
        else:
            for info in file_infos:
                if current_filter_key == "filter_favorites_only" and not info.get('is_favorite', False):
                    continue
                if search_text:
                    match = False
                    if current_filter_key in ["filter_all_notes", "filter_favorites_only", "filter_by_title"]:
                        if search_text in info['name'].lower(): match = True
                    if not match and current_filter_key in ["filter_all_notes", "filter_favorites_only",
                                                            "filter_by_summary"]:
                        if search_text in info.get('summary', '').lower(): match = True
                    if not match: continue
                filtered_infos.append(info)

        sort_index = self.sort_combo.currentIndex()
        sort_keys = ['sort_mod_desc', 'sort_mod_asc', 'sort_cre_desc', 'sort_cre_asc', 'sort_name_asc',
                     'sort_name_desc']
        sort_map = {"sort_mod_desc": ('modified_at', True), "sort_mod_asc": ('modified_at', False),
                    "sort_cre_desc": ('created_at', True), "sort_cre_asc": ('created_at', False),
                    "sort_name_asc": ('name', False), "sort_name_desc": ('name', True)}
        sort_key, reverse = sort_map.get(sort_keys[sort_index], ('modified_at', True))
        if sort_key != 'name':
            filtered_infos.sort(key=lambda x: x.get(sort_key, ''), reverse=reverse)
        else:
            filtered_infos.sort(key=lambda x: x['name'], reverse=reverse)
        filtered_infos.sort(key=lambda x: x.get('is_pinned', False), reverse=True)

        for folder_name in folders:
            if folder_name in ["images", "app_edge_profile", "app_chrome_profile", "drivers"]: continue
            folder_path = os.path.join(path, folder_name);
            folder_item = QTreeWidgetItem(parent_item, [folder_name])
            folder_item.setData(0, Qt.ItemDataRole.UserRole, folder_path)
            folder_item.setIcon(0, QIcon(self.style().standardIcon(self.style().StandardPixmap.SP_DirIcon)))
            folder_item.setFlags(folder_item.flags() & ~Qt.ItemFlag.ItemIsDragEnabled)
            self._populate_tree(folder_item, folder_path)
        for info in filtered_infos:
            item = QTreeWidgetItem(parent_item);
            item.setData(0, Qt.ItemDataRole.UserRole, info['path'])
            item_widget = NoteItemWidget(info['path'], info, self.tr)
            self.notes_tree_widget.setItemWidget(item, 0, item_widget);
            item.setSizeHint(0, item_widget.sizeHint())

    def on_note_selected(self, item, column):
        path = item.data(0, Qt.ItemDataRole.UserRole)
        if path and os.path.isfile(path):
            if self.current_note_path and self.note_editor.document().isModified():
                self.save_current_note(show_message=False)
            self.current_note_path = path
            content = self.note_manager.get_note_content(path)
            self.note_editor.setText(content)
            self.note_editor.document().setModified(False)
            self.update_preview()

    def update_preview(self):
        if not hasattr(self, 'note_manager'): return
        markdown_text = self.note_editor.toPlainText()
        eng_font = self.config.get('english_font', 'Arial')
        cn_font = self.config.get('chinese_font', '宋体')
        theme_name = self.config.get('theme', 'Default Light')
        bold_color = self.config.get('bold_color', '#000000')
        bold_style = f"<style>strong, b {{ color: {bold_color} !important; }}</style>"
        theme_css = ""
        if theme_name == "Dark":
            theme_css = "<style>body { background-color: #252525; color: #f0f0f0; }</style>"
        elif theme_name == "Cyberpunk":
            theme_css = "<style>body { background-color: #000; color: #00f0c0; }</style>"
        elif theme_name == "Newspaper":
            theme_css = "<style>body { background-color: #faf0e0; color: #333; }</style>"
        font_style = f"<style>body {{ font-family: '{eng_font}', '{cn_font}'; font-size: 16px; }}</style>"
        html = theme_css + font_style + bold_style + markdown.markdown(markdown_text,
                                                                       extensions=['fenced_code', 'tables'])
        base_url = QUrl.fromLocalFile(os.path.abspath(self.note_manager.notes_dir) + os.path.sep)
        self.preview_area.setHtml(html, baseUrl=base_url)

    def save_current_note(self, show_message=True):
        if self.current_note_path and self.note_editor.document().isModified():
            content = self.note_editor.toPlainText()
            self.note_manager.save_note(self.current_note_path, content)
            self.note_editor.document().setModified(False)
            self.load_notes_tree()
            if show_message:
                QMessageBox.information(self, self.tr('success'), self.tr('note_saved_success',
                                                                          note_name=os.path.basename(
                                                                              self.current_note_path)))

    def _create_reference_docx(self, filepath):
        try:
            from docx import Document
            from docx.shared import Pt
            from docx.oxml.ns import qn
        except ImportError:
            QMessageBox.critical(self, self.tr('error'),
                                 "python-docx library not found. Please install it via 'pip install python-docx'.")
            return

        document = Document()

        cn_font = self.config.get('chinese_font', '宋体')
        en_font = self.config.get('english_font', 'Arial')

        # --- Style for Normal Text (正文) ---
        style = document.styles['Normal']
        font = style.font
        font.name = en_font
        font.size = Pt(11)
        r = style._element.rPr.rFonts
        r.set(qn('w:eastAsia'), cn_font)

        # --- Style for Headings 1-4 (标题1-4) ---
        for i in range(1, 5):
            try:
                h_style = document.styles[f'Heading {i}']
                h_font = h_style.font
                h_font.name = en_font
                h_font.bold = True
                h_r = h_style._element.rPr.rFonts
                h_r.set(qn('w:eastAsia'), cn_font)
                if i == 1:
                    h_font.size = Pt(16)
                elif i == 2:
                    h_font.size = Pt(14)
                else:
                    h_font.size = Pt(12)
            except KeyError:
                print(f"Heading {i} style not found in base template, skipping.")

        document.save(filepath)

    def export_note(self, format_type):
        if not self.current_note_path:
            QMessageBox.warning(self, self.tr('tip'), self.tr('export_select_note_prompt'));
            return

        self.save_current_note(show_message=False)
        default_filename = os.path.splitext(os.path.basename(self.current_note_path))[0] + f".{format_type}"
        save_path, _ = QFileDialog.getSaveFileName(self, self.tr('export_to', format=format_type.upper()),
                                                   default_filename,
                                                   f"{self.tr('export_file_type', format=format_type.upper())} (*.{format_type});;{self.tr('all_files')} (*)")

        if save_path:
            ref_docx_path = None
            try:
                extra_args = [f'--resource-path={self.note_manager.notes_dir}']

                if format_type == 'docx':
                    ref_docx_path = os.path.join(os.getcwd(), "temp_reference_for_export.docx")
                    self._create_reference_docx(ref_docx_path)
                    extra_args.append(f'--reference-doc={ref_docx_path}')

                elif format_type == 'pdf':
                    extra_args.extend(['--pdf-engine=xelatex', '-V', f'mainfont={self.config["chinese_font"]}'])

                pypandoc.convert_file(self.current_note_path, format_type, outputfile=save_path, extra_args=extra_args)
                QMessageBox.information(self, self.tr('success'), self.tr('export_success', path=save_path))
            except Exception as e:
                QMessageBox.critical(self, self.tr('export_failed'), self.tr('export_pandoc_error', e=e))
            finally:
                if ref_docx_path and os.path.exists(ref_docx_path):
                    try:
                        os.remove(ref_docx_path)
                    except OSError as e:
                        print(f"Error removing temporary reference docx: {e}")

    def set_image_directory(self):
        dir_name = QFileDialog.getExistingDirectory(self, self.tr('select_image_folder_title'),
                                                    self.config['images_dir'])
        if dir_name:
            self.config['images_dir'] = dir_name;
            self._save_app_config(self.config)
            self.note_manager.images_dir = dir_name
            self.requests_converter.images_dir = dir_name
            self.selenium_manager.base_converter.images_dir = dir_name
            QMessageBox.information(self, self.tr('settings_saved'), self.tr('image_folder_updated', dir_name=dir_name))

    def set_browser_path(self):
        browser = self.config.get('browser', 'Chrome')
        path, _ = QFileDialog.getOpenFileName(self, self.tr('select_browser_exe', browser=browser), "",
                                              "Executable Files (*.exe)")
        if path:
            config_key = 'edge_binary_path' if browser == 'Edge' else 'chrome_binary_path'
            self.config[config_key] = path
            self._save_app_config(self.config)
            self.selenium_manager.config[config_key] = path
            QMessageBox.information(self, self.tr('success'), self.tr('browser_path_set', browser=browser, path=path))

    def open_font_settings(self):
        dialog = FontSettingsDialog(self.config, self.tr, self)
        if dialog.exec():
            self.config.update(dialog.get_selected_fonts());
            self._save_app_config(self.config)
            self.apply_styles()

    def get_selected_dir(self):
        selected_item = self.notes_tree_widget.currentItem()
        if selected_item:
            path = selected_item.data(0, Qt.ItemDataRole.UserRole)
            if path: return path if os.path.isdir(path) else os.path.dirname(path)
        return self.note_manager.notes_dir

    def show_tree_context_menu(self, pos):
        item = self.notes_tree_widget.itemAt(pos);
        menu = QMenu()
        new_note_action = menu.addAction(self.tr('new_note'));
        new_folder_action = menu.addAction(self.tr('new_folder'))
        if item:
            menu.addSeparator()
            path = item.data(0, Qt.ItemDataRole.UserRole)
            if path and os.path.isfile(path):
                meta = self.note_manager.get_item_metadata(path)
                pin_text = self.tr('unpin') if meta.get('is_pinned') else self.tr('pin_to_top')
                fav_text = self.tr('unfavorite') if meta.get('is_favorite') else self.tr('add_to_favorites')
                pin_action = menu.addAction(pin_text);
                fav_action = menu.addAction(fav_text)
                edit_summary_action = menu.addAction(self.tr('edit_summary'))
            if path:
                rename_action = menu.addAction(self.tr('rename'));
                delete_action = menu.addAction(self.tr('delete'))
        action = menu.exec(self.notes_tree_widget.mapToGlobal(pos));
        parent_dir = self.get_selected_dir()
        if action == new_note_action:
            name, ok = QInputDialog.getText(self, self.tr('new_note'), self.tr('enter_note_name'))
            if ok and name: self.note_manager.create_item(parent_dir, f"{name}.md"); self.load_notes_tree()
        elif action == new_folder_action:
            name, ok = QInputDialog.getText(self, self.tr('new_folder'), self.tr('enter_folder_name'))
            if ok and name: self.note_manager.create_item(parent_dir, name, is_folder=True); self.load_notes_tree()
        elif item and 'rename_action' in locals() and action == rename_action:
            old_name = os.path.basename(path);
            new_name, ok = QInputDialog.getText(self, self.tr('rename'), self.tr('enter_new_name'), text=old_name)
            if ok and new_name != old_name:
                _, error = self.note_manager.rename_item(path, new_name)
                if error:
                    QMessageBox.warning(self, self.tr('error'), error)
                else:
                    self.load_notes_tree()
        elif item and 'delete_action' in locals() and action == delete_action:
            reply = QMessageBox.question(self, self.tr('confirm_delete'),
                                         self.tr('confirm_delete_message', item_name=os.path.basename(path)),
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                self.note_manager.delete_item(path)
                if path == self.current_note_path: self.current_note_path = None; self.note_editor.clear(); self.preview_area.setHtml(
                    "")
                self.load_notes_tree()
        elif item and 'pin_action' in locals() and action == pin_action:
            self.note_manager.toggle_pinned(path);
            self.load_notes_tree()
        elif item and 'fav_action' in locals() and action == fav_action:
            self.note_manager.toggle_favorite(path);
            self.load_notes_tree()
        elif item and 'edit_summary_action' in locals() and action == edit_summary_action:
            self.edit_summary(path)

    def edit_summary(self, path):
        meta = self.note_manager.get_item_metadata(path)
        new_summary, ok = QInputDialog.getMultiLineText(self, self.tr('edit_summary'), self.tr('enter_summary'),
                                                        text=meta.get('summary', ''))
        if ok: self.note_manager.update_summary(path, new_summary); self.load_notes_tree()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MainWindow()
    main_win.show()
    sys.exit(app.exec())