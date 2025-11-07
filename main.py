import customtkinter as ctk
import winreg
import ctypes
import sys
import os
import json
import subprocess
from datetime import datetime
from tkinter import filedialog, messagebox
from typing import Dict, Any, Optional

# 管理者権限チェック
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# 管理者権限で再起動
def run_as_admin():
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        sys.exit()

class RegistryManager:
    """レジストリ操作を管理するクラス"""
    
    @staticmethod
    def read_value(key_path: str, value_name: str, root=winreg.HKEY_CURRENT_USER) -> Optional[Any]:
        """レジストリ値を読み取る"""
        try:
            with winreg.OpenKey(root, key_path, 0, winreg.KEY_READ) as key:
                value, _ = winreg.QueryValueEx(key, value_name)
                return value
        except FileNotFoundError:
            return None
        except Exception as e:
            print(f"読み取りエラー: {e}")
            return None
    
    @staticmethod
    def write_value(key_path: str, value_name: str, value: Any, value_type: int, root=winreg.HKEY_CURRENT_USER):
        """レジストリ値を書き込む"""
        try:
            with winreg.CreateKeyEx(root, key_path, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, value_name, 0, value_type, value)
            return True
        except Exception as e:
            print(f"書き込みエラー: {e}")
            return False
    
    @staticmethod
    def key_exists(key_path: str, root=winreg.HKEY_CURRENT_USER) -> bool:
        """レジストリキーが存在するかチェック"""
        try:
            with winreg.OpenKey(root, key_path, 0, winreg.KEY_READ):
                return True
        except FileNotFoundError:
            return False

class BackupManager:
    """設定のバックアップを管理するクラス"""
    
    def __init__(self, backup_file="settings_backup.json"):
        self.backup_file = backup_file
    
    def save_backup(self, settings: Dict[str, Any]):
        """設定をバックアップ"""
        backup_data = {
            "timestamp": datetime.now().isoformat(),
            "settings": settings
        }
        try:
            with open(self.backup_file, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"バックアップエラー: {e}")
            return False
    
    def load_backup(self) -> Optional[Dict[str, Any]]:
        """バックアップを読み込む"""
        try:
            if os.path.exists(self.backup_file):
                with open(self.backup_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
        except Exception as e:
            print(f"バックアップ読み込みエラー: {e}")
            return None

class SettingItem:
    """設定項目を表す基底クラス"""
    
    def __init__(self, name: str, description: str):
        self.name = name
        self.description = description
        self.current_value = None
        self.new_value = None
        self.modified = False
    
    def scan_current_value(self):
        """現在の設定値をスキャン"""
        raise NotImplementedError
    
    def apply_setting(self):
        """設定を適用"""
        raise NotImplementedError
    
    def get_backup_data(self) -> Dict[str, Any]:
        """バックアップデータを取得"""
        return {
            "name": self.name,
            "current_value": self.current_value
        }

class RegistrySettingItem(SettingItem):
    """レジストリ設定項目"""
    
    def __init__(self, name: str, description: str, key_path: str, value_name: str, 
                 value_type: int, enabled_value: Any, disabled_value: Any, 
                 root=winreg.HKEY_CURRENT_USER, labels=("有効", "無効")):
        super().__init__(name, description)
        self.key_path = key_path
        self.value_name = value_name
        self.value_type = value_type
        self.enabled_value = enabled_value
        self.disabled_value = disabled_value
        self.root = root
        self.labels = labels
    
    def scan_current_value(self):
        """現在の設定値をスキャン"""
        value = RegistryManager.read_value(self.key_path, self.value_name, self.root)
        if value == self.enabled_value:
            self.current_value = "enabled"
        elif value == self.disabled_value:
            self.current_value = "disabled"
        else:
            self.current_value = "unknown"
        return self.current_value
    
    def apply_setting(self):
        """設定を適用"""
        if self.new_value == "enabled":
            return RegistryManager.write_value(
                self.key_path, self.value_name, self.enabled_value, self.value_type, self.root
            )
        elif self.new_value == "disabled":
            return RegistryManager.write_value(
                self.key_path, self.value_name, self.disabled_value, self.value_type, self.root
            )
        return False

class PoleToWinApp(ctk.CTk):
    """メインアプリケーション"""
    
    def __init__(self):
        super().__init__()
        
        # ウィンドウ設定
        self.title("Pole To Win No11 - Windows 11 最適化ツール")
        self.geometry("900x700")
        
        # テーマ設定
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # マネージャー初期化
        self.backup_manager = BackupManager()
        
        # 設定項目リスト
        self.settings = self._initialize_settings()
        
        # UI要素の辞書
        self.setting_widgets = {}
        self.modified_items = set()
        
        # UIの構築
        self._build_ui()
        
        # 初期スキャン
        self.scan_all_settings()
    
    def _initialize_settings(self) -> Dict[str, SettingItem]:
        """設定項目を初期化"""
        settings = {
            # Bing検索連携
            "bing_search": RegistrySettingItem(
                name="Bing検索連携",
                description="検索ボックスとBingの連携",
                key_path=r"Software\Policies\Microsoft\Windows\Explorer",
                value_name="DisableSearchBoxSuggestions",
                value_type=winreg.REG_DWORD,
                enabled_value=0,
                disabled_value=1
            ),
            
            # フォルダ自動検出
            "folder_type": RegistrySettingItem(
                name="フォルダ自動検出",
                description="フォルダの種類の自動検出機能",
                key_path=r"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags\AllFolders\Shell",
                value_name="FolderType",
                value_type=winreg.REG_SZ,
                enabled_value="Generic",
                disabled_value="NotSpecified"
            ),
            
            # 広告ID
            "ad_id": RegistrySettingItem(
                name="広告ID",
                description="個人用広告の表示",
                key_path=r"Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo",
                value_name="Enabled",
                value_type=winreg.REG_DWORD,
                enabled_value=1,
                disabled_value=0
            ),
            
            # 透明効果
            "transparency": RegistrySettingItem(
                name="透明効果",
                description="ウィンドウの透明効果",
                key_path=r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                value_name="EnableTransparency",
                value_type=winreg.REG_DWORD,
                enabled_value=1,
                disabled_value=0,
                labels=("オン", "オフ")
            ),
            
            # タスクバー配置
            "taskbar_align": RegistrySettingItem(
                name="タスクバー配置",
                description="タスクバーアイコンの配置",
                key_path=r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
                value_name="TaskbarAl",
                value_type=winreg.REG_DWORD,
                enabled_value=1,  # 中央
                disabled_value=0,  # 左
                labels=("中央揃え", "左揃え")
            ),
            
            # タスクビュー
            "task_view": RegistrySettingItem(
                name="タスクビュー",
                description="タスクバーのタスクビューボタン",
                key_path=r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
                value_name="ShowTaskViewButton",
                value_type=winreg.REG_DWORD,
                enabled_value=1,
                disabled_value=0,
                labels=("表示", "非表示")
            ),
            
            # 右クリックメニュー
            "context_menu": RegistrySettingItem(
                name="右クリックメニュー",
                description="エクスプローラーの右クリックメニュー",
                key_path=r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32",
                value_name="",
                value_type=winreg.REG_SZ,
                enabled_value="",  # 従来仕様
                disabled_value="default",  # Windows 11仕様
                labels=("従来仕様", "Windows11仕様")
            ),
            
            # オプション診断データ
            "optional_diagnostic": RegistrySettingItem(
                name="オプション診断データ",
                description="Microsoftに送信する診断データ",
                key_path=r"Software\Microsoft\Windows\CurrentVersion\Diagnostics\DiagTrack",
                value_name="ShowedToastAtLevel",
                value_type=winreg.REG_DWORD,
                enabled_value=3,
                disabled_value=1,
                labels=("送信する", "最小限")
            ),
        }
        
        return settings
    
    def _build_ui(self):
        """UIを構築"""
        # ヘッダー
        header = ctk.CTkLabel(
            self, 
            text="Pole To Win No11", 
            font=ctk.CTkFont(size=24, weight="bold")
        )
        header.pack(pady=20)
        
        # 警告ラベル（設定変更時に表示）
        self.warning_label = ctk.CTkLabel(
            self,
            text="",
            font=ctk.CTkFont(size=12),
            text_color="orange"
        )
        self.warning_label.pack(pady=5)
        
        # メインコンテナ（スクロール可能）
        main_container = ctk.CTkScrollableFrame(self, width=850, height=400)
        main_container.pack(pady=10, padx=20, fill="both", expand=True)
        
        # 設定項目を表示
        for setting_id, setting in self.settings.items():
            self._create_setting_widget(main_container, setting_id, setting)
        
        # ボタンフレーム
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=20, padx=20, fill="x")
        
        # 設定適用ボタン
        apply_btn = ctk.CTkButton(
            button_frame,
            text="選択した項目の設定を適用",
            command=self.apply_selected_settings,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40
        )
        apply_btn.pack(side="left", padx=5, expand=True, fill="x")
        
        # すべて適用ボタン
        apply_all_btn = ctk.CTkButton(
            button_frame,
            text="すべての設定を適用",
            command=self.apply_all_settings,
            font=ctk.CTkFont(size=14, weight="bold"),
            height=40,
            fg_color="green"
        )
        apply_all_btn.pack(side="left", padx=5, expand=True, fill="x")
        
        # 初期設定に戻すボタン
        reset_btn = ctk.CTkButton(
            button_frame,
            text="初期設定に戻す",
            command=self.reset_settings,
            font=ctk.CTkFont(size=14),
            height=40,
            fg_color="gray"
        )
        reset_btn.pack(side="left", padx=5, expand=True, fill="x")
        
        # システムボタンフレーム
        system_button_frame = ctk.CTkFrame(self)
        system_button_frame.pack(pady=10, padx=20, fill="x")
        
        # Explorer再起動ボタン
        explorer_btn = ctk.CTkButton(
            system_button_frame,
            text="Explorer.exeを再起動",
            command=self.restart_explorer,
            height=35
        )
        explorer_btn.pack(side="left", padx=5, expand=True, fill="x")
        
        # Windows再起動ボタン
        windows_btn = ctk.CTkButton(
            system_button_frame,
            text="Windowsを再起動",
            command=self.restart_windows,
            height=35,
            fg_color="red"
        )
        windows_btn.pack(side="left", padx=5, expand=True, fill="x")
    
    def _create_setting_widget(self, parent, setting_id: str, setting: SettingItem):
        """設定項目のウィジェットを作成"""
        # フレーム
        frame = ctk.CTkFrame(parent, corner_radius=10)
        frame.pack(pady=8, padx=10, fill="x")
        
        # 設定名ラベル
        name_label = ctk.CTkLabel(
            frame,
            text=setting.name,
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        name_label.grid(row=0, column=0, sticky="w", padx=15, pady=(10, 0))
        
        # 説明ラベル
        desc_label = ctk.CTkLabel(
            frame,
            text=setting.description,
            font=ctk.CTkFont(size=11),
            text_color="gray",
            anchor="w"
        )
        desc_label.grid(row=1, column=0, sticky="w", padx=15, pady=(0, 5))
        
        # ラジオボタン変数
        radio_var = ctk.StringVar(value="enabled")
        
        # ラジオボタンフレーム
        radio_frame = ctk.CTkFrame(frame, fg_color="transparent")
        radio_frame.grid(row=0, column=1, rowspan=2, padx=15, pady=10)
        
        # ラジオボタンラベルを取得
        if isinstance(setting, RegistrySettingItem):
            enabled_label, disabled_label = setting.labels
        else:
            enabled_label, disabled_label = "有効", "無効"
        
        # 有効ラジオボタン
        enabled_radio = ctk.CTkRadioButton(
            radio_frame,
            text=enabled_label,
            variable=radio_var,
            value="enabled",
            command=lambda: self._on_setting_changed(setting_id, radio_var)
        )
        enabled_radio.pack(side="left", padx=10)
        
        # 無効ラジオボタン
        disabled_radio = ctk.CTkRadioButton(
            radio_frame,
            text=disabled_label,
            variable=radio_var,
            value="disabled",
            command=lambda: self._on_setting_changed(setting_id, radio_var)
        )
        disabled_radio.pack(side="left", padx=10)
        
        # ウィジェット情報を保存
        self.setting_widgets[setting_id] = {
            "frame": frame,
            "name_label": name_label,
            "radio_var": radio_var,
            "enabled_radio": enabled_radio,
            "disabled_radio": disabled_radio
        }
    
    def _on_setting_changed(self, setting_id: str, radio_var: ctk.StringVar):
        """設定が変更されたときの処理"""
        setting = self.settings[setting_id]
        new_value = radio_var.get()
        
        # 現在値と異なる場合は変更フラグを立てる
        if new_value != setting.current_value:
            setting.modified = True
            setting.new_value = new_value
            self.modified_items.add(setting_id)
            
            # ラベルを太字に
            widgets = self.setting_widgets[setting_id]
            widgets["name_label"].configure(font=ctk.CTkFont(size=14, weight="bold", underline=True))
        else:
            setting.modified = False
            if setting_id in self.modified_items:
                self.modified_items.remove(setting_id)
            
            # ラベルを通常に
            widgets = self.setting_widgets[setting_id]
            widgets["name_label"].configure(font=ctk.CTkFont(size=14, weight="bold"))
        
        # 警告メッセージを更新
        self._update_warning_message()
    
    def _update_warning_message(self):
        """警告メッセージを更新"""
        if self.modified_items:
            self.warning_label.configure(
                text=f"設定を変更しようとしている項目があります。「設定を適用」ボタンで設定が反映されます。（{len(self.modified_items)}項目）"
            )
        else:
            self.warning_label.configure(text="")
    
    def scan_all_settings(self):
        """すべての設定をスキャン"""
        for setting_id, setting in self.settings.items():
            current_value = setting.scan_current_value()
            
            # ラジオボタンの値を更新
            if setting_id in self.setting_widgets:
                radio_var = self.setting_widgets[setting_id]["radio_var"]
                radio_var.set(current_value if current_value != "unknown" else "disabled")
    
    def apply_selected_settings(self):
        """選択した設定を適用"""
        if not self.modified_items:
            messagebox.showinfo("情報", "変更された設定項目がありません。")
            return
        
        if not is_admin():
            messagebox.showerror("エラー", "設定を適用するには管理者権限が必要です。")
            return
        
        # バックアップを作成
        backup_data = {}
        for setting_id in self.modified_items:
            setting = self.settings[setting_id]
            backup_data[setting_id] = setting.get_backup_data()
        
        self.backup_manager.save_backup(backup_data)
        
        # 設定を適用
        success_count = 0
        for setting_id in list(self.modified_items):
            setting = self.settings[setting_id]
            if setting.apply_setting():
                success_count += 1
                setting.modified = False
                self.modified_items.remove(setting_id)
                
                # ラベルを通常に戻す
                widgets = self.setting_widgets[setting_id]
                widgets["name_label"].configure(font=ctk.CTkFont(size=14, weight="bold"))
        
        self._update_warning_message()
        
        messagebox.showinfo("完了", f"{success_count}個の設定を適用しました。\n一部の設定は再起動後に反映されます。")
        
        # 設定を再スキャン
        self.scan_all_settings()
    
    def apply_all_settings(self):
        """すべての設定を適用"""
        if not is_admin():
            messagebox.showerror("エラー", "設定を適用するには管理者権限が必要です。")
            return
        
        # すべての設定項目の新しい値を設定
        for setting_id, setting in self.settings.items():
            radio_var = self.setting_widgets[setting_id]["radio_var"]
            setting.new_value = radio_var.get()
        
        # バックアップを作成
        backup_data = {}
        for setting_id, setting in self.settings.items():
            backup_data[setting_id] = setting.get_backup_data()
        
        self.backup_manager.save_backup(backup_data)
        
        # すべて適用
        success_count = 0
        for setting_id, setting in self.settings.items():
            if setting.apply_setting():
                success_count += 1
        
        self.modified_items.clear()
        self._update_warning_message()
        
        messagebox.showinfo("完了", f"{success_count}個の設定を適用しました。\n一部の設定は再起動後に反映されます。")
        
        # 設定を再スキャン
        self.scan_all_settings()
    
    def reset_settings(self):
        """初期設定に戻す"""
        backup_data = self.backup_manager.load_backup()
        
        if not backup_data:
            messagebox.showwarning("警告", "バックアップが見つかりません。")
            return
        
        if not is_admin():
            messagebox.showerror("エラー", "設定を復元するには管理者権限が必要です。")
            return
        
        response = messagebox.askyesno(
            "確認",
            f"バックアップ（{backup_data.get('timestamp', '不明')}）から設定を復元しますか？"
        )
        
        if response:
            settings_data = backup_data.get("settings", {})
            success_count = 0
            
            for setting_id, data in settings_data.items():
                if setting_id in self.settings:
                    setting = self.settings[setting_id]
                    # バックアップされた値に戻す
                    setting.new_value = data.get("current_value")
                    if setting.apply_setting():
                        success_count += 1
            
            messagebox.showinfo("完了", f"{success_count}個の設定を復元しました。")
            self.scan_all_settings()
    
    def restart_explorer(self):
        """Explorerを再起動"""
        response = messagebox.askyesno("確認", "Explorer.exeを再起動しますか？")
        if response:
            try:
                subprocess.run("taskkill /f /im explorer.exe", shell=True, check=True)
                subprocess.run("start explorer.exe", shell=True, check=True)
                messagebox.showinfo("完了", "Explorer.exeを再起動しました。")
            except Exception as e:
                messagebox.showerror("エラー", f"再起動に失敗しました: {e}")
    
    def restart_windows(self):
        """Windowsを再起動"""
        response = messagebox.askyesno("確認", "Windowsを再起動しますか？\n保存されていないデータは失われます。")
        if response:
            try:
                subprocess.run("shutdown /r /t 0", shell=True, check=True)
            except Exception as e:
                messagebox.showerror("エラー", f"再起動に失敗しました: {e}")

def main():
    """メイン関数"""
    # 管理者権限チェック（情報表示のみ）
    if not is_admin():
        print("警告: 管理者権限なしで起動しています。設定の適用には管理者権限が必要です。")
    
    # アプリケーション起動
    app = PoleToWinApp()
    app.mainloop()

if __name__ == "__main__":
    main()