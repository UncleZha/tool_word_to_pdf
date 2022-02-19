from win32com.client import gencache
from win32com.client import constants, gencache
import os


# 创建PDF
def createPdf(wordPath, pdfPath):
    """
    word转pdf
    :param wordPath: word文件路径
    :param pdfPath:  生成pdf文件路径
    """
    word = gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(wordPath, ReadOnly=1)
    doc.ExportAsFixedFormat(pdfPath,
                            constants.wdExportFormatPDF,
                            Item=constants.wdExportDocumentWithMarkup,
                            CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    print(pdfPath)
    word.Quit(constants.wdDoNotSaveChanges)


# 遍历当前目录，并把Word文件转换为PDF
def wordToPdf():
    print("给小哈童做的:word文件批量转换成pdf小工具")
    print(" ")
    print("==========")
    print(" ")
    print("每个pdf成功保存时候会显示一行名称，如果没有出现一堆代码的话，就可以挂着玩别的去了，保留着这个窗口")
    print(" ")
    print("==========")
    print(" ")
    print("开始转换...")
    word_files_list = show_files("d:\\word_files", [])
    for word_file in word_files_list:
        pdf_name = os.path.splitext(word_file['name'])[0] + '.pdf'
        word_path = word_file['path']
        file_at_path = word_file['path'].split(word_file['name'])[0]
        pdf_path = 'd:\\out_pdf_files' + file_at_path.split("d:\\word_files")[1] + pdf_name
        createPdf(word_path, pdf_path)


def show_files(path, all_data):
    # 首先遍历当前目录所有文件及文件夹
    file_list = os.listdir(path)

    # 准备循环判断每个元素是否是文件夹还是文件，是文件的话，把名称传入list，是文件夹的话，递归
    for file in file_list:
        _data = {
            'path': '',
            'name': ''
        }
        # 利用os.path.join()方法取得路径全名，并存入cur_path变量，否则每次只能遍历一层目录
        cur_path = os.path.join(path, file)
        # 判断是否是文件夹
        if os.path.isdir(cur_path):
            pdf_path_now = 'd:\\out_pdf_files' + path.split("d:\\word_files")[1] + '\\' + file
            if not os.path.exists(pdf_path_now):
                os.makedirs(pdf_path_now)
            show_files(cur_path, all_data)
        else:
            if file.endswith((".doc", ".docx")):
                _data['path'] = cur_path
                _data['name'] = file
                all_data.append(_data)
    return all_data


# word转pdf
if __name__ == '__main__':
    # contents = show_files("d:\\word_files", [])
    # for content in contents:
    #     print(content)
    wordToPdf()
    os.system('pause')
