import json
import os

def save_config(api_key, api_password, directory):
    config = {
        'api_key': api_key,
        'api_password': api_password,
        'directory': directory
    }

    try:
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
    except Exception as e:
        raise Exception(f"Error saving config: {e}")

def load_config():
    if not os.path.exists('config.json'):
        raise FileNotFoundError("Error: config.json file not found")

    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)

            required_keys = {'api_key', 'api_password', 'directory'}
            if not required_keys.issubset(config.keys()):
                raise KeyError("Missing one or more required keys in the config file")

            return (
                config.get('api_key'),
                config.get('api_password'),
                config.get('directory')
            )
    except Exception as e:
        raise Exception(f"Error loading config: {e}")
