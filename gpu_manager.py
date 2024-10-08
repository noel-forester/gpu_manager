import sys
import os
import winreg
from PyQt5.QtWidgets import (QApplication, QMainWindow, QSystemTrayIcon, 
                             QTableWidget, QTableWidgetItem, QMenu, QAction, 
                             QPushButton, QVBoxLayout, QHBoxLayout, QWidget, QMessageBox, QComboBox, QHeaderView, QFileDialog)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import (Qt, QEvent)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GPU Manager")
        self.setGeometry(100, 100, 600, 400)
        
        # Table for applications and GPU assignments
        self.table = QTableWidget(0, 3)  # Initially empty, rows will be added dynamically
        self.table.setHorizontalHeaderLabels(["Name", "Current GPU", "Select GPU"])
        
        # カラム幅の設定（リサイズモードを指定）
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)   # Nameカラムは拡張
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Current GPUは内容に合わせる
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Select GPUも内容に合わせる

        # Buttons
        self.add_btn = QPushButton("Add")
        self.delete_btn = QPushButton("Delete")
        self.refresh_btn = QPushButton("Refresh")
        self.apply_btn = QPushButton("Apply")
        self.exit_btn = QPushButton("Exit")
        
        
        
        # Layout for buttons
        button_layout_top = QHBoxLayout()  # 1行目のボタン配置
        button_layout_top.addWidget(self.add_btn)
        button_layout_top.addWidget(self.delete_btn)

        button_layout_bottom = QHBoxLayout()  # 2行目のボタン配置
        button_layout_bottom.addWidget(self.refresh_btn)
        button_layout_bottom.addWidget(self.apply_btn)
        button_layout_bottom.addWidget(self.exit_btn)

        # Main layout
        layout = QVBoxLayout()
        layout.addWidget(self.table)
        layout.addLayout(button_layout_top)
        layout.addLayout(button_layout_bottom)
        
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        
        # Button actions
        self.apply_btn.clicked.connect(self.apply_settings)
        self.refresh_btn.clicked.connect(self.refresh_table)
        self.add_btn.clicked.connect(self.add_entry)
        self.delete_btn.clicked.connect(self.delete_entry)
        self.exit_btn.clicked.connect(self.exit_app)
        
        # Load current settings from registry
        self.load_registry_settings()

    def exit_app(self):
        self.tray_icon.hide()  # トレイアイコンを非表示にする
        self.main_window.close()  # メインウィンドウを閉じる
        QApplication.quit()  # アプリケーションを終了



    def load_registry_settings(self):
        # テーブルの初期化
        self.table.setRowCount(0)
        
        # レジストリからGPU設定を読み込み
        registry_path = r"Software\Microsoft\DirectX\UserGpuPreferences"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path) as key:
                i = 0
                while True:
                    try:
                        # レジストリの値を読み取る
                        full_app_path, gpu_setting, _ = winreg.EnumValue(key, i)
                        app_name = os.path.basename(full_app_path)  # 実行ファイル名だけにする
                        
                        # GPU設定が "GpuPreference=0;" の形式である場合
                        if gpu_setting.startswith("GpuPreference="):
                            gpu_value = gpu_setting.split('=')[1].replace(";", "")
                            self.add_app_to_table(full_app_path, app_name, gpu_value)
                        
                        i += 1
                    except OSError:
                        break
        except FileNotFoundError:
            print("レジストリキーが見つかりませんでした")

    def format_gpu_value(self, gpu_value):
        # GPU値を説明付きテキストに変換
        if gpu_value == "0":
            return "auto(0)"
        elif gpu_value == "1":
            return "iGPU(1)"
        elif gpu_value == "2":
            return "dGPU(2)"
        return gpu_value

    def add_app_to_table(self, full_app_path, app_name, gpu_value):
        # テーブルに新しい行を追加して情報を表示
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        
        # アプリ名（左寄せ）
        app_item = QTableWidgetItem(app_name)
        app_item.setData(Qt.UserRole, full_app_path)  # フルパスを保持
        app_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.table.setItem(row_position, 0, app_item)
        
        # 現在のGPU設定を変換して表示（左寄せ）
        gpu_item = QTableWidgetItem(self.format_gpu_value(gpu_value))
        gpu_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.table.setItem(row_position, 1, gpu_item)
        
        # GPU選択用のコンボボックス（左寄せ）
        combo = QComboBox()
        combo.addItems(["auto(0)", "iGPU(1)", "dGPU(2)"])  # 表示値
        combo.setCurrentText(self.format_gpu_value(gpu_value))  # 現在の設定を選択状態に
        self.table.setCellWidget(row_position, 2, combo)

    def apply_settings(self):
        # テーブルから情報を取得してレジストリに保存
        registry_path = r"Software\Microsoft\DirectX\UserGpuPreferences"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path, 0, winreg.KEY_SET_VALUE) as key:
            for row in range(self.table.rowCount()):
                # フルパスと選択されたGPU設定を取得
                app_item = self.table.item(row, 0)
                full_app_path = app_item.data(Qt.UserRole)  # フルパスを保持していたデータを取得
                selected_gpu_text = self.table.cellWidget(row, 2).currentText()
                
                # 選択された表示値を数字に戻す
                selected_gpu = selected_gpu_text[-2]  # "auto(0)" -> "0" のように末尾の数値を取り出す
                
                # レジストリに保存（"GpuPreference=数字;"の形式で）
                winreg.SetValueEx(key, full_app_path, 0, winreg.REG_SZ, f"GpuPreference={selected_gpu};")
        
        QMessageBox.information(self, "Apply", "Settings applied!")
        
        # Apply後にテーブルを自動更新
        self.refresh_table()

    def refresh_table(self):
        # テーブルを更新（内容を再読み込み）
        self.load_registry_settings()

    def add_entry(self):
        # ファイル選択ダイアログを開く
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter("Executable Files (*.exe)")
        if file_dialog.exec_():
            exe_path = file_dialog.selectedFiles()[0]
            app_name = os.path.basename(exe_path)
            
            # レジストリに新規エントリを追加
            registry_path = r"Software\Microsoft\DirectX\UserGpuPreferences"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path, 0, winreg.KEY_SET_VALUE) as key:
                    winreg.SetValueEx(key, exe_path, 0, winreg.REG_SZ, "GpuPreference=0;")
                
                QMessageBox.information(self, "Add", f"{app_name} added successfully!")
                # テーブルを更新
                self.refresh_table()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to add {app_name}: {e}")

    def delete_entry(self):
        # 選択された行を取得して削除
        selected_row = self.table.currentRow()
        if selected_row < 0:
            QMessageBox.warning(self, "Delete", "No entry selected.")
            return
        
        # レジストリから選択されたエントリを削除
        app_item = self.table.item(selected_row, 0)
        full_app_path = app_item.data(Qt.UserRole)
        registry_path = r"Software\Microsoft\DirectX\UserGpuPreferences"
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, registry_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, full_app_path)
            
            QMessageBox.information(self, "Delete", f"{app_item.text()} deleted successfully!")
            # テーブルから行を削除
            self.table.removeRow(selected_row)
        except FileNotFoundError:
            QMessageBox.critical(self, "Error", "The selected entry was not found.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to delete {app_item.text()}: {e}")


    def closeEvent(self, event):
        # ウィンドウの×ボタンが押されたときに最小化して非表示にする
        event.ignore()  # イベントを無視して閉じない
        self.hide()  # ウィンドウを非表示にする
    
    def changeEvent(self, event):
        # ウィンドウが最小化されたときに非表示にする
        if event.type() == QEvent.WindowStateChange and self.isMinimized():
            self.hide()  # ウィンドウを非表示にする
            event.accept()

class TrayApp(QApplication):
    def __init__(self, sys_argv):
        super().__init__(sys_argv)
        
        # PyInstallerでの実行ファイルかどうかを確認し、一時ディレクトリを使用
        if getattr(sys, 'frozen', False):  # PyInstallerで実行ファイル化されている場合
            app_icon_path = os.path.join(sys._MEIPASS, "icon.png")
        else:  # 通常のPythonスクリプトの場合
            app_icon_path = "icon.png"
        
        # アプリケーションのアイコン設定
        app_icon = QIcon(app_icon_path)
        self.setWindowIcon(app_icon)
        
        # MainWindowインスタンス
        self.main_window = MainWindow()
        
        # 起動時にメインウィンドウを表示
        self.main_window.show()
        
        # System tray icon
        self.tray_icon = QSystemTrayIcon(app_icon, parent=self)
        self.tray_icon.setToolTip("GPU Manager")
        
        # Tray menu
        self.tray_menu = QMenu()
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.exit_app)
        self.tray_menu.addAction(exit_action)
        
        # コンテキストメニューをトレイアイコンに設定
        self.tray_icon.setContextMenu(self.tray_menu)
        
        # 左クリックと右クリックでメニューが表示されるように設定
        self.tray_icon.activated.connect(self.on_tray_icon_click)
        self.tray_icon.show()
    
    def on_tray_icon_click(self, reason):
        if reason == QSystemTrayIcon.Trigger:  # 左クリック
            # ウィンドウを表示してアクティブにする
            self.main_window.showNormal()
            self.main_window.raise_()
            self.main_window.activateWindow()
        elif reason == QSystemTrayIcon.Context:  # 右クリック
            # 右クリックメニューを表示
            self.tray_menu.exec_(self.tray_icon.geometry().center())
    
    def exit_app(self):
        self.tray_icon.hide()  # トレイアイコンを非表示にする
        self.main_window.close()  # メインウィンドウを閉じる
        QApplication.quit()  # アプリケーションを終了

if __name__ == '__main__':
    app = TrayApp(sys.argv)
    sys.exit(app.exec_())
