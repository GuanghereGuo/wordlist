import pandas as pd
import re

# ================= 配置区域 =================
INPUT_FILE = 'words.xlsx'  # 你的文件名
OUTPUT_TEX = 'wordlist_content.tex' # 输出的中间内容文件
# ===========================================

def escape_latex(text):
    """转义 LaTeX 特殊字符，防止编译报错"""
    if pd.isna(text):
        return ""
    text = str(text)
    chars = {
        '&': r'\&',
        '%': r'\%',
        '$': r'\$',
        '#': r'\#',
        '_': r'\_',
        '{': r'\{',
        '}': r'\}',
        '~': r'\textasciitilde{}',
        '^': r'\textasciicircum{}',
        '\\': r'\textbackslash{}',
    }
    pattern = re.compile('|'.join(re.escape(key) for key in chars.keys()))
    return pattern.sub(lambda x: chars[x.group()], text)

def generate_latex_table():
    # 读取Excel的所有Sheet
    xls = pd.ExcelFile(INPUT_FILE)
    
    latex_content = ""

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # 确保列名对应（根据你的截图调整）
        # 假设列顺序是固定的：A=单词, B=音标, C=中文, D=例句, G=错误历史
        # 如果列名固定，也可以用列名索引：df['单词']
        
        # 提取数据，iloc[:, 0] 代表第一列
        data = df.iloc[:, [0, 1, 2, 3, 6]].copy() 
        data.columns = ['word', 'phonetic', 'meaning', 'example', 'history']
        
        # 开始构建当前 Sheet 的 LaTeX 代码
        # 添加一个小标题
        latex_content += f"\\section*{{List: {escape_latex(sheet_name)}}}\n"
        
        # 表格头
        latex_content += r"""
\begin{longtable}{p{3cm} p{4cm} p{8cm} p{2cm}}
\toprule
\textbf{Word} & \textbf{Meaning} & \textbf{Example} & \textbf{Note} \\
\midrule
\endfirsthead
\toprule
\textbf{Word} & \textbf{Meaning} & \textbf{Example} & \textbf{Note} \\
\midrule
\endhead
\bottomrule
\endfoot
"""
        
        # 遍历每一行生成内容
        for index, row in data.iterrows():
            word = escape_latex(row['word'])
            phonetic = escape_latex(row['phonetic'])
            meaning = escape_latex(row['meaning'])
            example = escape_latex(row['example'])
            history = escape_latex(row['history'])
            
            # 处理 N/A 或 空值
            if history == "nan": history = ""
            
            # 格式化：单词加粗，音标换行并变小
            # col1 = f"\\textbf{{{word}}} \\newline \\small{{[{phonetic}]}}"
            col1 = f"\\textbf{{{word}}} \\newline \\small{{\\ipafont [{phonetic}]}}"
            
            # 写入一行
            latex_content += f"{col1} & {meaning} & {example} & {history} \\\\ \n"
            latex_content += r"\cmidrule(l){1-4}" + "\n" # 添加行间细线，如果不喜欢可以去掉

        latex_content += r"\end{longtable}" + "\n\n"
        print(f"处理完成: {sheet_name}")

    # 保存到 tex 文件
    with open(OUTPUT_TEX, 'w', encoding='utf-8') as f:
        f.write(latex_content)
    print(f"成功生成: {OUTPUT_TEX}")

if __name__ == '__main__':
    generate_latex_table()