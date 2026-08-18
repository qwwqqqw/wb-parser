import os
import sys
import webview
from core.config_manager import ConfigManager
from core.wb_parser import WBParser

class Api:
    def __init__(self):
        self._config = ConfigManager()
        self._parser = WBParser(self)
        self._window = None

    def open_excel(self, filepath):
        print(f"Открытие файла Excel: {filepath}")
        import os, platform
        try:
            if platform.system() == 'Windows':
                os.startfile(filepath)
        except Exception as e:
            print(f"Ошибка открытия Excel: {e}")
            
    def set_window(self, window):
        self._window = window

    def get_config(self):
        return self._config.get_all()

    def save_config(self, new_config):
        self._config.update(new_config)
        return {"status": "ok"}

    def get_categories(self):
        return self._parser.fetch_categories()

    def start_parsing(self, params):
        self._parser.start(params)
        return {"status": "started"}

    def stop_parsing(self):
        self._parser.stop()
        return {"status": "stopped"}

    def log(self, message, filepath=None):
        """Отправляет лог в интерфейс. Если передан filepath, сигнализирует о завершении парсинга."""
        import json
        if self._window:
            safe_msg = json.dumps(message)
            self._window.evaluate_js(f"addLog({safe_msg})")
            if filepath:
                self.notify_finished(filepath)

    def notify_finished(self, filepath=None):
        """Уведомляет интерфейс о завершении или остановке парсинга."""
        import json
        if self._window:
            safe_path = json.dumps(filepath) if filepath else "null"
            self._window.evaluate_js(f"parsingFinished({safe_path})")

def main():
    api = Api()
    
    if hasattr(sys, '_MEIPASS'):
        ui_path = os.path.join(sys._MEIPASS, 'ui', 'index.html')
    else:
        ui_path = os.path.join(os.path.dirname(__file__), 'ui', 'index.html')
    
    window = webview.create_window(
        'WB Parser', 
        url=f'file:///{ui_path.replace(os.sep, "/")}',
        js_api=api,
        maximized=True,
        min_size=(1024, 768)
    )
    api.set_window(window)
    
    webview.start(debug=True)

if __name__ == '__main__':
    main()
