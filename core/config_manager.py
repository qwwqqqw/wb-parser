import json
import os

CONFIG_FILE = 'config.json'

class ConfigManager:
    def __init__(self):
        self.config = self._load()

    def _load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Ошибка загрузки конфигурации: {e}")
                return self._default_config()
        return self._default_config()

    def _default_config(self):
        return {
            "theme": "dark",
            "seller_url": "",
            "mode": "all",  
            "price_min": 0,
            "price_max": 100000,
            "items_count": 100,
            "auth_mode": "no_auth",
            "merge_files": True,
            "captcha_mode": "manual",
            "selected_categories": []
        }

    def get_all(self):
        return self.config

    def update(self, new_config):
        self.config.update(new_config)
        self._save()

    def _save(self):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка сохранения конфигурации: {e}")
