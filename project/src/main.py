from config import load_config
from utils import setup_logging, create_backup, set_format
from processor import WordProcessor


def main():
    config = load_config()
    setup_logging(config.get('log_level', 'INFO'))
    file_path = config['file_path']
    create_backup(file_path)
    processor = WordProcessor(file_path, config)
    processor.process_all_sheets()
    processor.save()
    set_format(file_path, config)


if __name__ == "__main__":
    main()