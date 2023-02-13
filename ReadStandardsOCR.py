

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
# from reportlab.pdfgen import canvas
# from reportlab.pdfbase.cidfonts import UnicodeCIDFont
# from reportlab.pdfbase import pdfmetrics
# from reportlab.pdfbase.ttfonts import TTFont
# from reportlab.lib.units import mm

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

from pdf2image import convert_from_path
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import pyocr
import re
from pathlib import Path

import cv2

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
        # pdfmetrics.registerFont(TTFont(self.fontname1, GEN_SHIN_GOTHIC_MEDIUM_TTF))
        # pdfmetrics.registerFont(TTFont(self.fontname2, IPAEXG_TTF))
    #end def
    #*********************************************************************************

        
        

    def OCRFile(self, filename, bitflag=False):
        if filename =="" :
            return False
        #end if

        pdf_file = filename

        #OCRエンジンを取得する
        tools = pyocr.get_available_tools()
        if len(tools) == 0:
            print("OCRエンジンが指定されていません")
            sys.exit(1)
        else:
            tool = tools[0]
        # print(text)
        pdfFileName = filename
        dpi0 = 300      # 数値を読み取る場合のDPI

        fname = Path(pdfFileName).stem
        cname = fname[:fname.find("_")]
        pn = fname[fname.find("_")+1:]
        print(fname,cname,pn)

        if isint(pn):
            stpage = int(pn)
        #end if

        pageText = []
        pageResultData = []
        pageNo = []
        pdfKind = []
        kind = "スキャンデータ"
        tagger =MeCab.Tagger()
        tagger.parse('')
        
        try:

            with open(pdfFileName, "rb") as input:
                reader = PdfReader(input)
                PageMax = len(reader.pages)
                input.close()
            #end with
            images = convert_from_path(pdfFileName,dpi=dpi0,first_page=1,last_page=PageMax)
                
            i = -1
            for image in images:
                i += 1
                
                if bitflag:
                    # 2値画像に変換に変換する場合
                    #データ形式をnumpy配列に変換
                    img = np.array(image)
                    # モノクロ・グレースケール画像へ変換（2値化前の画像処理）
                    img_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                    # 2値化（Binarization）：白（1）黒（0）のシンプルな2値画像に変換
                    retval, img_binary = cv2.threshold(img_gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
                    # データ形式をCV2形式変換
                    image2 = Image.fromarray(img_binary)
                    texts = tool.image_to_string(
                        image2,
                        lang='jpn+eng',
                        builder=pyocr.builders.TextBuilder(tesseract_layout=6)
                    )
                else:
                    texts = tool.image_to_string(
                        image,
                        lang='jpn+eng',
                        builder=pyocr.builders.TextBuilder(tesseract_layout=6)
                    )
                #end if

                print(i+1)
                
                # print (texts)
                pagen = stpage + i
                pageNo.append(pagen)

                lines = texts.splitlines()
                pmax = len(lines)
                # lastLine = lines[len(lines)-1]
                # pn = lastLine.replace(" ","").replace("\n","")
                # if isint(pn):
                #     pageNo.append(pn)
                texts2 = ""
                j = 0
                for line in lines:
                    j += 1
                    if j> 1 and i<pmax:
                        line2 = ""
                        sflag = False
                        for k in range(len(line)):
                            c = line[k]
                            if sflag:
                                if c != " ":
                                    line2 += c
                                    sflag = False
                                #end if
                            else:
                                if c == " ":
                                    line2 += c
                                    sflag = True
                                else:
                                    line2 += c
                                #end if
                            #end if
                        #next
                        texts2 += line2.replace("|","")
                    # end if
                #next

                node = tagger.parseToNode(texts2)
                word_list=[]
                while node:
                    word_type = node.feature.split(',')[0]
                    if word_type in ["助詞"]:
                        if node.surface == "の":
                            word_list.append(node.surface)
                        #end if
                    #end if
                    if word_type in ["名詞",'代名詞']:
                        word_list.append(node.surface)
                    #end if
                    node=node.next
                #end while
                word_chain=''.join(word_list)
                pageText.append(texts2)
                pageResultData.append(word_chain)
                pdfKind.append(kind)
            #next
        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False, pageNo, pageText, pageResultData, pdfKind
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False, pageNo, pageText, pageResultData, pdfKind
        #end try

        return True, pageNo, pageText, pageResultData, pdfKind
    
    # end def






    def LoadFile(self, filename):

        if filename =="" :
            return False
        #end if

        pdf_file = filename

        # PDFMinerのツールの準備
        resourceManager = PDFResourceManager()
        # PDFから単語を取得するためのデバイス
        laparams = LAParams()               # パラメータインスタンス
        laparams.boxes_flow = None          # -1.0（水平位置のみが重要）から+1.0（垂直位置のみが重要）default 0.5
        laparams.word_margin = 0.1          # default 0.1
        laparams.char_margin = 2.0          # default 2.0
        laparams.line_margin = 0.5            # default 0.5
        device = PDFPageAggregator(resourceManager, laparams=laparams)

        pageText = []
        pageResultData = []
        pageNo = []
        pdfKind = []
        kind = "テキストを含むPDF"
        tagger =MeCab.Tagger()
        tagger.parse('')
        
        try:
            with open(pdf_file, 'rb') as fp:
                interpreter = PDFPageInterpreter(resourceManager, device)

                pageI = 0
                pn2 = 0
                for page in PDFPage.get_pages(fp):
                    pageI += 1
                    print("page={}:".format(pageI), end="")
                    
                    interpreter.process_page(page)
                    layout = device.get_result()
                    texts = ""
                    ResultData = []
                    for lt in layout:
                        if isinstance(lt, LTTextBoxHorizontal):
                            text = lt.get_text()
                            # print(text)
                            texts += text
                        #end if
                    #next
                    lines = texts.splitlines()
                    pmax = len(lines)
                    if pageI == 1:
                        lastLine = lines[len(lines)-1]
                        pn = lastLine.replace(" ","").replace("\n","")
                        if isint(pn):
                            pn2 = int(pn)
                        else:
                            pn2 = 1
                        #end if
                    else:
                        pn2 += 1
                    #end id

                    pageNo.append(pn2)
                    
                    texts2 = ""
                    i = 0
                    for line in lines:
                        i += 1
                        if i> 1 and i<pmax:
                            texts2 += line
                        # end if
                    #next

                    node = tagger.parseToNode(texts2)
                    word_list=[]
                    while node:
                        word_type = node.feature.split(',')[0]
                        if word_type in ["助詞"]:
                            if node.surface == "の":
                                word_list.append(node.surface)
                            #end if
                        #end if
                        if word_type in ["名詞",'代名詞']:
                            word_list.append(node.surface)
                        node=node.next
                    #end while
                    word_chain=''.join(word_list)
                    pageText.append(texts2)
                    pageResultData.append(word_chain)
                    pdfKind.append(kind)
                #next
                fp.close()
            # end with

        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False, pageNo, pageText, pageResultData, pdfKind
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False, pageNo, pageText, pageResultData, pdfKind
        #end try

        # 使用したデバイスをクローズ
        device.close()

        return True, pageNo, pageText, pageResultData, pdfKind

#end class

def main():
    
    dir = os.getcwd()
    inputRCPath = filedialog.askdirectory(initialdir=dir)
    # inputRCPath = "/Users/kanyama/VS Code/MeCabPDF/2020年版黄色本（スキャン）"
    folderfile = os.listdir(inputRCPath)
    print(folderfile)
    # BookName = os.path.basename(os.path.dirname(inputRCPath))
    BookName = os.path.basename(inputRCPath)
    
    if len(folderfile)>0:
        time_sta = time.time()  # 開始時刻の記録
        PT = PdfPage2Text()
        FMA = FMRestAPI()


        ChapterNames = []
        for file in folderfile:
            if file.find("_")>0:
                fname = Path(file).stem
                ChapterName = fname[:fname.find("_")]        
                ChapterNames.append(ChapterName)
                flag,pageNo, pageText, pageResultData, pdfKind  = PT.OCRFile(inputRCPath + "/" +file,bitflag=False)
            else:
                ChapterName = file.replace(".pdf","").replace(".PDF","")
                ChapterNames.append(ChapterName)
                flag,pageNo, pageText, pageResultData, pdfKind = PT.LoadFile(inputRCPath + "/" +file)

            if flag:
                n = len(pageText)
                if n>0 :
                    pageId = []
                    for i in range(n):
                        print(pageText[i])
                        pn = pageNo[i]
                        texts = pageText[i]
                        word_chain = pageResultData[i]
                        kind = pdfKind[i]
                        data = {
                            "ページ番号":pn,
                            "テキスト全文":texts,
                            "テキスト固有名詞":word_chain,
                            "基準文書の名称":BookName,
                            "章の名称":ChapterName,
                            "pdfの種類":kind,
                            "最初のページ":pageNo[0]
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
    t1 = time.time() - time_sta
    print("time = {} sec".format(t1))
    #end if

if __name__ == '__main__':
    main()
                
"""
自宅からのVPN接続環境において、
テキストを含むPDFの場合、60ページの読み込みに約90秒,610ページの読み込みに約585秒
スキャンデータのPDFの場合、60ページの読み込みに約280秒

所内の有線LAN接続環境において、
テキストを含むPDFの場合、60ページの読み込みに約21秒,610ページの読み込みに約207秒
スキャンデータのPDFの場合、60ページの読み込みに約243秒
イメージを二値化すると321秒
"""