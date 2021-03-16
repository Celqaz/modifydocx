# -*- coding: utf-8 -*
'''
@Author: 达达里昂
@Contact: libin9400@gmail.com
@Date: 2021-03-10 16:29:41
LastEditTime: 2021-03-13 10:46:45
@Description: Modify the docs. Without the tears.
'''
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import time
import os


class GetFileNames():
    def __init__(self, templatePath='template'):
        self.files = os.listdir(templatePath)
        self.filenames = []

    def getFileNames(self):
        for file_name in self.files:
            pureFileName = file_name.split(".", 1)
            # 排除子目录、非word类型文件及隐藏文件（"~"for MacOS，"." for Windows）。
            if (len(pureFileName) == 1) or (pureFileName[1] != "docx"
                                            or pureFileName[0][0] == "~" or pureFileName[0][0] == "."):
                continue
            self.filenames.append(pureFileName[0])
        return self.filenames


class ModifyFile():
    def __init__(self, preOutputName, customTextFileName, templateFileName):
        """初始化定义。同main()函数。

        Args:
            customTextFileName (str): 自定义文本文件名称。位于text文件夹下的.txt文件名称，不需输入.txt的拓展名。
            templateFileName (str): 模版文件名称。位于template文件夹下的.docx文件名称，不需输入.docx的拓展名
            preOutputName (str, 可选参数): 输出文件名称前缀，默认为空。输出的文件名称样式为“输出文件名称前缀+模版文件名称+时间标记.docx”，保存在output文件夹下。
        """
        self.customTextFileName = "text"+"/"+customTextFileName+".txt"
        self.templateFilePath = "template/"+templateFileName+".docx"
        file_time = time.strftime("%m%d-%H%M%S", time.localtime())
        #   如果用户定义的文件前缀为'',则返回空，否则返回“前缀_”的格式
        self.modPreOutputName = preOutputName + "_" if preOutputName != '' else ''
        self.outputFilePath = "output/" + \
            self.modPreOutputName + templateFileName + \
            "[" + file_time + "].docx"
        # run
        self.modify = self.handleDocx(self.getText())

    def getText(self):
        """获取text.txt中的文本

        Returns:
            dict: 自定义的变量名称及值
        """
        # 获取自定义文本的路径，并打开
        path = self.customTextFileName
        file = open(path, encoding="utf8")

        # 遍历读取自定义文本，并存储为dict
        DICT = {}
        for line in file.readlines():
            # 对字符串进行切片，按”：“切1次，分割成2段
            try:
                text = line.split("：", 1)
                # 在字段前后添加@，以符合模版文件中占位符的样式，从而方便查找替换
                DICT['@'+text[0]+'@'] = text[1].strip()
            except:
                # 忽略注释及空行
                continue
        return DICT

    def handleDocx(self, customText):
        """文件处理，打开模版文件，调用处理函数，另存为新文件

        Args:
            customText (dict): text.txt中用户自定义文本
            outputpath (str): 输出文件的存储相对路径
        """
        if self.templateFilePath.split(".")[1] == 'docx':
            document = Document(self.templateFilePath)
            document = self.replace_string(customText, document)
            document.save(self.outputFilePath)
            print('🟢导出成功ヾ(✿ﾟ▽ﾟ)ノ\n👉文件路径为：'+self.outputFilePath+'\n')

    def replace_string(self, customText, document):
        """遍历文档，将"@xxxx@"样式的占位文本，替换为用户在text.txt定义的文本

        Args:
            customText (dict): text.txt中用户自定义文本
            document (docx.document.Document): 模版docx文件

        Returns:
            [docx.document.Document]: 返回替换完成的docx文件
        """
        for key, value in customText.items():

            for p in document.paragraphs:

                if key in p.text:
                    newText = p.text.replace(key, value)
                    inline = p.runs
                    indexMark = []

                    # 将所有分词清空
                    for i in range(len(inline)):
                        inline[i].text = ''
                    # 将第一段文本替换为更改后的文本
                    inline[0].add_text(newText)
                    # 高亮修改后的段落
                    inline[0].font.highlight_color = WD_COLOR_INDEX.YELLOW

        return document


def main(preOutputName='', customTextFileName='text', templatePath='template'):
    """传入初始数据

    Args:
        customTextFileName (str): 自定义文本文件名称。位于text文件夹下的.txt文件名称，不需输入.txt的拓展名。
        templateFileName (str): 模版文件名称。位于template文件夹下的.docx文件名称，不需输入.docx的拓展名
        preOutputName (str, 可选参数): 输出文件名称前缀，默认为空。输出的文件名称样式为“输出文件名称前缀+模版文件名称+时间标记.docx”，保存在output文件夹下。
        textFont (str, 可选参数): 被替换进模版文件的文本的字体样式，默认为"楷体"。建议更改为和模版文件自身相同的字体。
    """
    start = time.time()
    # 调用GetFileNames类下的getFileNames()函数，获取模版文件夹下所有docx文件名称（不含子文件夹）。
    fileNamesList = GetFileNames(templatePath).getFileNames()
    print("📃目录载入完毕，共"+str(len(fileNamesList))+'个.docx文件。\n')
    count = 0
    fail_list = []

    for fileName in fileNamesList:
        print("🕓正在处理第"+str(count+1)+"个文件："+fileName)
        # ModifyFile(preOutputName, customTextFileName, fileName)

        try:
            ModifyFile(preOutputName, customTextFileName, fileName)
            count = count + 1
        except:
            print("🔴替换失败，文件名为："+fileName)
            fail_list.append(fileName)
    if count == len(fileNamesList):
        print('\n 🎉全部文件替换完成，共'+str(count)+'份文件。')
    else:
        print('\n🔴部分替换成功，其中：\n\t\t成功替换'+str(count)+'份文件。')
        print('\t\t'+str(len(fileNamesList)-count)+'份文件替换失败，文件名为：')
        for i in fail_list:
            print('\t\t'+i)
    end = time.time()
    print('用时: ', str(round(end - start, 3))+'s')


if __name__ == "__main__":
    inputOutputName = input("请输入文件保存的名称前缀(可留空)，完成后按Enter：")
    main(inputOutputName)
    # main()
