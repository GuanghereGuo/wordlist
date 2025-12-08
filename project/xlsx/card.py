import pandas as pd
import os
import random
from colorama import init, Fore, Style

# 初始化颜色
init(autoreset=True)

FILE_PATH = 'words.xlsx'

class FlashCard:
    def __init__(self, file_path):
        self.file_path = file_path
        try:
            self.excel_file = pd.ExcelFile(file_path)
        except FileNotFoundError:
            print(f"{Fore.RED}错误: 找不到文件 {file_path}")
            exit()
        self.sheets = self.excel_file.sheet_names
        self.df = None
        self.sheet_name = ""
        self.modified = False

    def clear_screen(self):
        """清屏，支持 Windows 和 Mac/Linux"""
        os.system('cls' if os.name == 'nt' else 'clear')

    def save_progress(self):
        """只保存对'生词度'列的修改"""
        if not self.modified:
            return

        print(f"\n{Fore.YELLOW}正在保存生词度状态到 Excel...{Style.RESET_ALL}")
        try:
            # 使用 openpyxl 引擎，以追加模式打开，替换当前 Sheet 数据
            with pd.ExcelWriter(self.file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                self.df.to_excel(writer, sheet_name=self.sheet_name, index=False)
            print(f"{Fore.GREEN}保存成功！{Style.RESET_ALL}")
        except Exception as e:
            print(f"{Fore.RED}保存失败 (请关闭 Excel 文件后重试): {e}{Style.RESET_ALL}")

    def run(self):
        # 1. 选择 Sheet
        print(f"{Fore.CYAN}=== 单词闪示卡 ==={Style.RESET_ALL}")
        for i, s in enumerate(self.sheets):
            print(f"{i+1}. {s}")
        
        try:
            idx = int(input("\n请选择单元 (数字): ")) - 1
            self.sheet_name = self.sheets[idx]
        except:
            print("无效输入")
            return

        # 读取数据
        self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_name)
        
        # 2. 选择模式
        print(f"\n{Fore.CYAN}=== 选择背诵模式 ==={Style.RESET_ALL}")
        print("1. 看英文 -> 想中文 (默认)")
        print("2. 看中文 -> 想英文")
        mode = input("请输入 (1/2): ").strip()
        is_en_to_cn = mode != '2'

        # 准备数据索引
        # 假设列顺序：Word(0), Phonetic(1), Meaning(2), Example(3), Status(4-生词度)
        # 你的截图中，生词度是 E 列，也就是索引 4
        col_word = self.df.columns[0]
        col_phonetic = self.df.columns[1]
        col_mean = self.df.columns[2]
        col_ex = self.df.columns[3]
        col_status = self.df.columns[4] # 生词度

        # 生成待背列表 (过滤掉空行)
        indices = list(self.df.dropna(subset=[col_word]).index)
        
        # 是否乱序？
        random.shuffle(indices)

        print(f"\n{Fore.GREEN}>>> 开始背诵 {self.sheet_name} (共 {len(indices)} 词) <<<{Style.RESET_ALL}")
        print("操作指南: 按 Enter 翻牌，输入 'q' 退出\n")
        input("按 Enter 开始...")

        count = 0
        for i in indices:
            count += 1
            row = self.df.iloc[i]
            word = str(row[col_word])
            mean = str(row[col_mean])
            phonetic = str(row[col_phonetic]) if pd.notna(row[col_phonetic]) else ""
            example = str(row[col_ex]) if pd.notna(row[col_ex]) else ""
            status = str(row[col_status]) if pd.notna(row[col_status]) else "未测试"

            self.clear_screen()
            print(f"{Fore.WHITE}进度: {count}/{len(indices)} | 当前状态: {status}{Style.RESET_ALL}\n")

            # === 正面 (问题) ===
            print(f"{Fore.CYAN + Style.BRIGHT}----------------------------------------")
            if is_en_to_cn:
                print(f"   {word}")  # 只显示英文
            else:
                print(f"   {mean}")  # 只显示中文
            print(f"----------------------------------------{Style.RESET_ALL}")
            
            cmd = input("\n>> 按 Enter 查看答案 (q 退出)...")
            if cmd.lower() == 'q':
                break

            # === 背面 (答案) ===
            print(f"\n{Fore.YELLOW}答案:{Style.RESET_ALL}")
            print(f"单词: {Fore.GREEN}{word}{Style.RESET_ALL}  [{phonetic}]")
            print(f"释义: {mean}")
            print(f"例句: {Fore.LIGHTBLACK_EX}{example}{Style.RESET_ALL}")
            
            # === 自我评估 ===
            print(f"\n{Fore.MAGENTA}你记住了吗？{Style.RESET_ALL}")
            print("y (Yes) = 认识 (标记为'已掌握')")
            print("n (No)  = 忘了 (标记为'需复习')")
            print("Enter   = 跳过 (不修改状态)")
            
            eval_input = input(">> ").lower().strip()

            if eval_input == 'y':
                self.df.at[i, col_status] = '已掌握'
                self.modified = True
            elif eval_input == 'n':
                self.df.at[i, col_status] = '需复习'
                self.modified = True
            
            # 这里可以加个小逻辑：如果选了 n，可以把这个 index 再 append 到 indices 列表末尾
            # 从而实现“本轮如果不认识，一会再问一遍”
            if eval_input == 'n':
                indices.append(i)
                print(f"{Fore.RED}>>> 已加入队尾，稍后重测{Style.RESET_ALL}")
                import time
                time.sleep(0.5)

        # self.save_progress()

if __name__ == "__main__":
    app = FlashCard(FILE_PATH)
    app.run()