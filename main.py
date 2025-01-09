#!/usr/bin/env python3
# -*- coding: utf-8 -*-


from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QLabel, QPushButton, QFileDialog, QMessageBox,
                            QLineEdit, QHBoxLayout)
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QIcon
import json
import pandas as pd
from pathlib import Path
import sys
import os

class TaskConfigApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings("YourCompany", "YourApp")  # 设置应用程序名称
        self.last_file_path = self.settings.value("lastFilePath", "")  # 获取上次文件路径
        self.task_name = ""
        self.task_types = []
        self.phones_per_type = 0
        
        # 设置应用图标
        icon_path = Path(__file__).parent / "assets" / "app_icon.png"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("任务配置工具")
        self.setGeometry(100, 100, 500, 400)
        self.setStyleSheet("""
            QMainWindow, QWidget {
                background-color: white;
            }
            QLabel {
                font-size: 14px;
                color: black;
                background: transparent;
            }
            QLabel#task_info, QLabel#file_info {
                font-size: 18px;
                font-weight: bold;
            }
            QLabel#task_info span {
                color: #FF4500;
                font-weight: bold;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #ccc;
                border-radius: 3px;
                min-height: 25px;
                background-color: white;
                color: black;
            }
            QPushButton {
                padding: 8px 15px;
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 3px;
                min-height: 30px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        
        # 创建主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(15)
        
        # 添加顶部弹性空间
        self.main_layout.addStretch(2)  # 增加顶部权重
        
        # 创建一个容器来包含标题和按钮
        container = QVBoxLayout()
        container.setSpacing(20)  # 设置标题和按钮之间的间距
        
        # 第一阶段组件
        self.file_label = QLabel("请选择任务模板文件(Excel)")
        self.file_label.setObjectName("file_info")
        self.file_label.setAlignment(Qt.AlignCenter)
        container.addWidget(self.file_label)
        
        self.select_button = QPushButton("选择文件")
        self.select_button.setCursor(Qt.PointingHandCursor)
        self.select_button.clicked.connect(self.select_excel_file)
        container.addWidget(self.select_button)
        
        # 将容器添加到主布局
        self.main_layout.addLayout(container)
        
        # 添加底部弹性空间
        self.main_layout.addStretch(3)  # 增加底部权重
        
        # 第二阶段组件
        self.task_info_label = QLabel()
        self.task_info_label.setObjectName("task_info")  # 设置对象名以应用特定样式
        self.task_info_label.setAlignment(Qt.AlignCenter)
        self.task_info_label.hide()
        self.main_layout.addWidget(self.task_info_label)
        
        self.phones_label = QLabel("请输入每种类型分配的手机数量(1或2):")
        self.phones_label.hide()
        self.main_layout.addWidget(self.phones_label)
        
        self.phones_entry = QLineEdit()
        self.phones_entry.setPlaceholderText("输入1或2")
        self.phones_entry.hide()
        self.main_layout.addWidget(self.phones_entry)
        
        self.range_label = QLabel("请输入需要使用的手机编号范围(例如: 1-8):")
        self.range_label.hide()
        self.main_layout.addWidget(self.range_label)
        
        self.range_entry = QLineEdit()
        self.range_entry.setPlaceholderText("例如: 1-8")
        self.range_entry.hide()
        self.main_layout.addWidget(self.range_entry)
        
        # 在确认按钮前添加一个水平布局来放置两个按钮
        button_layout = QHBoxLayout()
        
        # 返回按钮
        self.back_button = QPushButton("返回上一步")
        self.back_button.setCursor(Qt.PointingHandCursor)
        self.back_button.clicked.connect(self.back_to_first_stage)
        self.back_button.hide()
        button_layout.addWidget(self.back_button)
        
        # 确认按钮
        self.confirm_button = QPushButton("确认配置")
        self.confirm_button.setCursor(Qt.PointingHandCursor)
        self.confirm_button.clicked.connect(self.validate_and_process)
        self.confirm_button.hide()
        button_layout.addWidget(self.confirm_button)
        
        # 将按钮布局添加到主布局
        self.main_layout.addLayout(button_layout)
        
        # 添加底部弹性空间
        self.main_layout.addStretch()
        
    def select_excel_file(self):
        options = QFileDialog.Options()
        if self.last_file_path:  # 如果有上次的路径，使用它
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "选择Excel文件",
                self.last_file_path,  # 使用上次的路径
                "Excel files (*.xlsx *.xls)",
                options=options
            )
        else:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "选择Excel文件",
                "",
                "Excel files (*.xlsx *.xls)",
                options=options
            )
        
        if file_path:
            self.last_file_path = file_path  # 记住当前选择的路径
            self.settings.setValue("lastFilePath", self.last_file_path)  # 保存路径
            self.process_excel_file(file_path)
    
    def get_folders_json_path(self):
        """根据操作系统返回 folders.json 的路径"""
        if sys.platform == 'win32':  # Windows
            return Path.home() / "AppData" / "Roaming" / "flip" / "folders.json"
        elif sys.platform == 'darwin':  # macOS
            return Path.home() / "Library" / "Application Support" / "flip" / "folders.json"
        else:  # 其他系统
            raise OSError("Unsupported operating system")

    def create_backup(self, json_path):
        """创建配置文件的备份，并进行错误处理"""
        try:
            backup_path = json_path.parent / "folders_bak.json"
            
            # 检查是否有足够的磁盘空间
            import shutil
            if shutil.disk_usage(str(json_path.parent)).free < os.path.getsize(str(json_path)):
                raise OSError("磁盘空间不足，无法创建备份")
            
            # 如果备份文件已存在，先检查是否可写
            if backup_path.exists() and not os.access(str(backup_path), os.W_OK):
                raise PermissionError(f"无法写入备份文件: {backup_path}")
            
            # 创建备份
            shutil.copy2(json_path, backup_path)
            return backup_path
            
        except Exception as e:
            raise RuntimeError(f"创建备份文件时出错: {str(e)}")

    def process_excel_file(self, file_path):
        try:
            # 获取并验证 folders.json 路径
            json_path = self.get_folders_json_path()
            
            # 创建备份
            backup_path = self.create_backup(json_path)
            print(f"已创建配置文件备份: {backup_path}")
            
            # 读取 Excel 文件
            try:
                df = pd.read_excel(file_path)
            except Exception as e:
                raise ValueError(f"Excel文件格式错误: {str(e)}")
            
            # 清理列名（移除空格并转换为小写）
            df.columns = df.columns.str.strip().str.lower()
            
            # 验证必需的列（使用小写进行比较）
            required_columns = ['平台', '账号', '标题']
            required_columns_lower = [col.lower() for col in required_columns]
            missing_columns = [col for col in required_columns if col.lower() not in df.columns]
            if missing_columns:
                raise ValueError(f"Excel文件缺少必需的列: {', '.join(missing_columns)}")
            
            if df.empty:
                raise ValueError("Excel文件中没有数据")
            
            self.task_name = Path(file_path).stem
            self.task_types = df.apply(
                lambda row: (
                    str(row['平台']).strip(),
                    str(row['账号']).strip(),
                    str(row['标题']).strip()
                ),
                axis=1
            ).unique()
            
            if not self.task_types.size:
                raise ValueError("未能从Excel文件中提取到任何任务类型")
            
            # 更新界面，使用 HTML 格式确保数字显示为红色
            count = len(self.task_types)
            self.task_info_label.setText(
                f"检测到 <span style='color: #FF4500; font-weight: bold;'>{count}</span> 种不同的任务类型"
            )
            self.show_second_stage()
            
        except (FileNotFoundError, PermissionError, ValueError, RuntimeError) as e:
            QMessageBox.critical(self, "错误", str(e))
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理Excel文件时出现未知错误：\n{str(e)}")

    def validate_and_process(self):
        try:
            phones_per_type = int(self.phones_entry.text().strip())
            if phones_per_type not in [1, 2]:
                raise ValueError("每种类型的手机数量必须为1或2")
                
            phone_range = self.range_entry.text().strip()
            if not phone_range:
                raise ValueError("请输入手机编号范围")
                
            try:
                phone_numbers = self.parse_phone_range(phone_range)
            except:
                raise ValueError("手机编号范围格式错误，请使用类似 '1-8' 或 '01-08' 的格式")
                
            expected_count = len(self.task_types) * phones_per_type
            if len(phone_numbers) != expected_count:
                raise ValueError(
                    f"手机数量不匹配！需要 {expected_count} 部手机，"
                    f"但输入范围包含 {len(phone_numbers)} 部"
                )
                
            self.update_json_file(phone_numbers, phones_per_type)
            
        except ValueError as e:
            QMessageBox.warning(self, "输入错误", str(e))
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理配置时出错：\n{str(e)}")
            
    def parse_phone_range(self, range_str):
        range_str = range_str.strip()
        if '-' in range_str:
            start, end = range_str.split('-')
            start = int(start.lstrip('0') or '0')
            end = int(end.lstrip('0') or '0')
            if start > end:
                raise ValueError("起始编号不能大于结束编号")
            return list(range(start, end + 1))
        else:
            return [int(range_str.lstrip('0') or '0')]
            
    def show_success_message(self):
        # 创建自定义消息框
        msg = QMessageBox(self)
        msg.setWindowTitle("成功")
        msg.setText("配置已成功保存！原文件已备份！")
        
        # 设置图标
        icon_path = Path(__file__).parent / "assets" / "app_icon.icns."
        if icon_path.exists():
            icon = QIcon(str(icon_path))
            msg.setIconPixmap(icon.pixmap(32, 32))  # 设置较小的图标尺寸
        
        # 设置窗口图标（标题栏图标）
        msg.setWindowIcon(self.windowIcon())
        
        msg.exec_()

    def update_json_file(self, phone_numbers, phones_per_type):
        try:
            # 获取 folders.json 路径
            json_path = self.get_folders_json_path()
            
            with open(json_path, 'r', encoding='utf-8') as f:
                json_data = json.load(f)
            
            phone_idx = 0
            for task_type in self.task_types:
                for _ in range(phones_per_type):
                    phone_num = phone_numbers[phone_idx]
                    phone_id = f"Phone{phone_num:02d}"
                    
                    for item in json_data:
                        if item['name'].startswith(phone_id):
                            item['name'] = f"{phone_id}_{self.task_name}"
                            item['targetPath'] = self.task_name
                            if 'extra' not in item:
                                item['extra'] = {}
                            item['extra']['platform'] = task_type[0]
                            item['extra']['account'] = task_type[1]
                            item['extra']['title'] = task_type[2]
                            break
                            
                    phone_idx += 1
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
                
            # 使用自定义的成功消息框
            self.show_success_message()
            self.close()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存配置时出错：{str(e)}")

    def back_to_first_stage(self):
        # 清空输入
        self.phones_entry.clear()
        self.range_entry.clear()
        
        # 隐藏所有组件
        while self.main_layout.count():
            item = self.main_layout.takeAt(0)
            if item.widget():
                item.widget().hide()
            elif item.layout():
                # 清理子布局
                while item.layout().count():
                    child = item.layout().takeAt(0)
                    if child.widget():
                        child.widget().hide()
        
        # 重新添加第一阶段组件
        self.main_layout.addStretch()  # 顶部弹性空间
        self.file_label.show()
        self.main_layout.addWidget(self.file_label)
        self.select_button.show()
        self.main_layout.addWidget(self.select_button)
        self.main_layout.addStretch()  # 底部弹性空间

    def show_second_stage(self):
        # 隐藏第一阶段的组件
        self.file_label.hide()
        self.select_button.hide()
        
        # 清除所有布局
        while self.main_layout.count():
            item = self.main_layout.takeAt(0)
            if item.widget():
                item.widget().hide()
            elif item.layout():
                while item.layout().count():
                    child = item.layout().takeAt(0)
                    if child.widget():
                        child.widget().hide()
        
        # 添加顶部弹性空间
        self.main_layout.addStretch()
        
        # 显示第二阶段组件并添加到布局
        self.task_info_label.show()
        self.main_layout.addWidget(self.task_info_label)
        
        self.phones_label.show()
        self.main_layout.addWidget(self.phones_label)
        
        self.phones_entry.show()
        self.main_layout.addWidget(self.phones_entry)
        
        self.range_label.show()
        self.main_layout.addWidget(self.range_label)
        
        self.range_entry.show()
        self.main_layout.addWidget(self.range_entry)
        
        # 创建按钮布局
        button_layout = QHBoxLayout()
        self.back_button.show()
        self.confirm_button.show()
        button_layout.addWidget(self.back_button)
        button_layout.addWidget(self.confirm_button)
        self.main_layout.addLayout(button_layout)
        
        # 添加底部弹性空间
        self.main_layout.addStretch()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 设置应用程序图标
    icon_path = Path(__file__).parent / "assets" / "app_icon.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    
    window = TaskConfigApp()
    window.show()
    sys.exit(app.exec_())