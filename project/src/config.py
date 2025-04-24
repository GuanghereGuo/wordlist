import yaml

def load_config(config_path='../config/config.yaml'):
    """读取配置文件并返回配置字典"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        print(f"读取配置文件失败: {e}")
        return {}