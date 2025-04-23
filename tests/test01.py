import codecs
import xlwt

intlist = ['1','2','3','4','5','6','7','8','9','0']
worddict = {}
#{'word':[word,soundmark,[mean]]}

meanlist_ = []
meanlist = []
with open('./zzq.txt','r',encoding='utf-16') as wordfile:
    while True :
        wordline = wordfile.readline()
        if(wordline):
            if(wordline[0] in intlist):
                word = wordline.split('[')[0].split(',')[1].strip()
                if len(wordline.split('['))>3:
                    soundmark = ''
                    for soundmark_ in wordline.split('[')[1:]:
                        soundmark = soundmark +'['+soundmark_.strip()
                elif len(wordline.split('['))>1:
                    soundmark = '['+wordline.split('[')[1].strip()
                else:
                    soundmark = ''
                worddict[word]=[word,soundmark,[]]
            else:
                if '人名' not in wordline:
                    worddict[word][2].append(wordline)
                else:
                    pass
        else:
            break


#设置表格样式
def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style

# 写Excel
def write_excel(worddict):
    f = xlwt.Workbook()
    sheet1 = f.add_sheet('words',cell_overwrite_ok=True)
    for item,word in enumerate(  worddict):
        sheet1.write(item,0,worddict[word][0],set_style('Times New Roman',220,True))
        sheet1.write(item,1,worddict[word][1],set_style('Times New Roman',220,True))
        for item2,meaning in enumerate(  worddict[word][2]):
            sheet1.write(item,2+item2,meaning,set_style('Times New Roman',220,True))
    f.save('word.xls')


write_excel(worddict)
