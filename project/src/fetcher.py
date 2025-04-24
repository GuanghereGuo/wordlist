import requests
from bs4 import BeautifulSoup
import time
import logging

def get_word_info(word, config):
    """从有道词典获取单词信息"""
    url = f"http://dict.youdao.com/w/{word}/"
    headers = {'User-Agent': config['request']['user_agent']}
    retries = config['request']['retries']
    for attempt in range(retries):
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                phonetic = soup.find('span', class_='phonetic')
                phonetic_text = phonetic.text if phonetic else ''
                trans_container = soup.find('div', class_='trans-container')
                definition = ''
                if trans_container and trans_container.find('ul'):
                    definition = trans_container.find('ul').find('li').text or ''
                examples = soup.find('div', class_='examples')
                example_sentence = examples.find('p').text if examples and examples.find('p') else ''
                return phonetic_text, definition, example_sentence
            else:
                logging.warning(f"请求失败，状态码: {response.status_code}")
        except Exception as e:
            logging.error(f"获取单词 {word} 信息出错: {e}")
        if attempt < retries - 1:
            time.sleep(5)
    logging.error(f"单词 {word} 获取失败，已记录")
    with open('failed_words.txt', 'a', encoding='utf-8') as f:
        f.write(f"{word}\n")
    return None, None, None