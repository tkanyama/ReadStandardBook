# pip install pdfminer
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.converter import TextConverter
from pdfminer.pdfpage import PDFPage
# from pdfminer.layout import LAParams, LTTextContainer
from pdfminer.layout import LAParams, LTTextContainer, LTContainer, LTTextBox, LTTextLine, LTChar,LTTextBoxHorizontal

# # pip install pdfrw
# from pdfrw import PdfReader
# from pdfrw.buildxobj import pagexobj
# from pdfrw.toreportlab import makerl

# pip install reportlab
from reportlab.pdfgen import canvas
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# pip install PyPDF2
from PyPDF2 import PdfReader, PdfWriter, PdfFileReader, PdfFileWriter # 名前が上とかぶるので別名を使用

# その他のimport
import os,time
import sys
import numpy as np
import logging
import MeCab
from tkinter import filedialog
from tkinter import messagebox
import fmrest
import json

#============================================================================
#  整数を表しているかどうかを判定する関数
#============================================================================
def isint(s):  
    try:
        int(s)  # 文字列を実際にint関数で変換してみる
    except ValueError:
        return False
    else:
        return True
    #end if
#end def


class FMRestAPI():
    def __init__(self):
        path = os.path.expanduser('~/init2.json')
        f = open(path, 'r')
        datajson = json.loads(f.read())
        f.close()
        USER = datajson['USER']
        PASSWORD = datajson['PASSWORD']
        self.serverAddress = "192.168.0.171"
        self.databaseName = "基準書検索"
        self.layoutName = "ページ情報"
    
        self.fms = fmrest.Server("https://{}".format(self.serverAddress),
                user=USER,
                password=PASSWORD,
                database=self.databaseName,
                layout=self.layoutName,
                api_version='vLatest',
                verify_ssl=False
                )
        self.fms.login()

    def insertrRecord(self, jsondata):
        res = self.fms.create_record(jsondata)
        print(res)
        return res
    #end def

    def insertPdf(self, Id, fieldName, pdfFile):
        with open(pdfFile, 'rb') as pdf_pf:
            result = self.fms.upload_container(Id, fieldName, pdf_pf)
        return result
    
    def insertPdf2(self, Id, fieldName, pdfFileIO):
        result = self.fms.upload_container(Id, fieldName, pdfFileIO)
        return result
    
#end class


class PdfPage2Text():
    def __init__(self):
        # 源真ゴシック等幅フォント
        # GEN_SHIN_GOTHIC_MEDIUM_TTF = "/Library/Fonts/GenShinGothic-Monospace-Medium.ttf"
        GEN_SHIN_GOTHIC_MEDIUM_TTF = "./Fonts/GenShinGothic-Monospace-Medium.ttf"
        self.fontname1 = 'GenShinGothic'
        # IPAexゴシックフォント
        # IPAEXG_TTF = "/Library/Fonts/ipaexg.ttf"
        IPAEXG_TTF = "./Fonts/ipaexg.ttf"
        self.fontname2 = 'ipaexg'
        
        # フォント登録
        pdfmetrics.registerFont(TTFont(self.fontname1, GEN_SHIN_GOTHIC_MEDIUM_TTF))
        pdfmetrics.registerFont(TTFont(self.fontname2, IPAEXG_TTF))
    #end def
    #*********************************************************************************



    def Page2Text(self, page, interpreter, device):
        
        #============================================================
        # 構造計算書がSS7の場合の処理
        #============================================================
        pageFlag = False
        ResultData = []
        interpreter.process_page(page)
        layout = device.get_result()
        
        texts = ""
        textArry = []
        for lt in layout:
            # LTTextContainerの場合だけ標準出力　断面算定表(杭基礎)
            # if isinstance(lt, LTTextContainer):
            if isinstance(lt, LTTextContainer):
                text = lt.get_text()
                texts += text
                textArry.append(text)
            #end if
        #next
        
        
        #==========================================================================
        #  検出結果を出力する
        return pageFlag, ResultData
    #end def           



    def LoadFile(self, filename):

        if filename =="" :
            return False
        #end if

        pdf_file = filename

        # PyPDF2のツールを使用してPDFのページ情報を読み取る。
        # PDFのページ数と各ページの用紙サイズを取得
        # try:
        #     with open(pdf_file, "rb") as input:
        #         reader = PdfReader(input)
        #         PageMax = len(reader.pages)     # PDFのページ数
        #         PaperSize = []
        #         for page in reader.pages:       # 各ページの用紙サイズの読取り
        #             p_size = page.mediabox
        #             page_xmin = float(page.mediabox.lower_left[0])
        #             page_ymin = float(page.mediabox.lower_left[1])
        #             page_xmax = float(page.mediabox.upper_right[0])
        #             page_ymax = float(page.mediabox.upper_right[1])
        #             PaperSize.append([page_xmax - page_xmin , page_ymax - page_ymin])
        #         #next
        #         input.close()
        #     #end with
        # except OSError as e:
        #     print(e)
        #     logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
        #     return False
        # except:
        #     logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
        #     return False
        # #end try


        # PDFMinerのツールの準備
        resourceManager = PDFResourceManager()
        # PDFから単語を取得するためのデバイス
        laparams = LAParams()               # パラメータインスタンス
        laparams.boxes_flow = None          # -1.0（水平位置のみが重要）から+1.0（垂直位置のみが重要）default 0.5
        laparams.word_margin = 0.1          # default 0.1
        laparams.char_margin = 2.0          # default 2.0
        laparams.line_margin = 0.5            # default 0.5
        device = PDFPageAggregator(resourceManager, laparams=laparams)
        # device = TextConverter(rmgr, outfp, laparams=lprms)    # TextConverterオブジェクトの取得
        # PDFから１文字ずつを取得するためのデバイス
        # device2 = PDFPageAggregator(resourceManager)

        pageText = []
        pageResultData = []
        pageNo = []
        tagger =MeCab.Tagger()
        tagger.parse('')
        # node = tagger.parseToNode(text)

        try:
            with open(pdf_file, 'rb') as fp:
                interpreter = PDFPageInterpreter(resourceManager, device)
                # interpreter2 = PDFPageInterpreter(resourceManager, device2)

                pageI = 0
                        
                for page in PDFPage.get_pages(fp):
                    pageI += 1
                    print("page={}:".format(pageI), end="")
                    
                    interpreter.process_page(page)
                    layout = device.get_result()
                    texts = ""
                    ResultData = []
                    for lt in layout:
                        # LTTextContainerの場合だけ標準出力　断面算定表(杭基礎)
                        if isinstance(lt, LTTextBoxHorizontal):
                            text = lt.get_text()
                            print(text)
                            texts += text
                        #end if
                    #next
                    lines = texts.splitlines()
                    lastLine = lines[len(lines)-1]
                    pn = lastLine.replace(" ","").replace("\n","")
                    if isint(pn):
                        pageNo.append(pn)

                    node = tagger.parseToNode(texts)
                    word_list=[]
                    while node:
                        word_type = node.feature.split(',')[0]
                        if word_type in ["名詞",'代名詞']:
                            word_list.append(node.surface)
                        node=node.next
                    #end with
                    word_chain=' '.join(word_list)
                    pageText.append(texts)
                    pageResultData.append(word_chain)


        
                #next

                fp.close()
            # end with

        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False, pageNo, pageText, pageResultData
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False, pageNo, pageText, pageResultData
        #end try


        # 使用したデバイスをクローズ
        device.close()

        return True, pageNo, pageText, pageResultData






#end class

def main():
    dir = os.getcwd()
    # inputRCPath = filedialog.askdirectory(initialdir=dir)
    inputRCPath = "/Users/kanyama/VS Code/MeCabPDF/2020年版黄色本（査読版）"
    folderfile = os.listdir(inputRCPath)
    print(folderfile)
    # BookName = os.path.basename(os.path.dirname(inputRCPath))
    BookName = os.path.basename(inputRCPath)
    
    if len(folderfile)>0:
        PT = PdfPage2Text()
        FMA = FMRestAPI()


        ChapterNames = []
        for file in folderfile:
            ChapterName = file.replace(".pdf","").replace(".PDF","")
            ChapterNames.append(ChapterName)
            flag,pageNo, pageText, pageResultData = PT.LoadFile(inputRCPath + "/" +file)
            if flag:
                n = len(pageText)
                if n>0 :
                    pageId = []
                    for i in range(n):
                        print(pageText[i])
                        pn = pageNo[i]
                        texts = pageText[i]
                        word_chain = pageResultData[i]
                        data = {
                            "ページ番号":pn,
                            "テキスト全文":texts,
                            "テキスト固有名詞":word_chain,
                            "基準文書の名称":BookName,
                            "章の名称":ChapterName
                        }
                        Id = FMA.insertrRecord(data)
                        if Id>0 :
                            pageId.append(Id)
                        #end if

                        with open(inputRCPath + "/" +file,'rb') as f:
                            pdfReader = PdfReader(f)
                            file_object = pdfReader.pages[i]
                            
                            pdf_file = "./pdf/pdftmp.pdf"
                            with open(pdf_file,'wb') as f2: #(10)
                                pdfWriter1 = PdfWriter(f2)
                                pdfWriter1.add_page(file_object) #(11)
                                pdfWriter1.write(f2)
                                f2.close()

                            f.close()
                        #end with

                        fieldName = "pdf"
                        res = FMA.insertPdf(Id, fieldName, pdf_file)

                    #next
                    if os.path.isfile(pdf_file):
                        os.remove(pdf_file)
                #end if
            #end if
        #next
    #end if

if __name__ == '__main__':
    main()
                
