'''
機能
PDFファイルを読込み、OpenAIにより要約させたテキストファイルを保存する
UIは作成していないためコマンドで実行する必要あり
クリエイティブ・コモンズ・ライセンス BY-NC-SA
参考文献
https://www.kkaneko.jp/ai/chatgpt/summary.html

ChatGPT 3.5 turbo の利用では APIキーが必要である
[利用法]
コマンドプロンプトから本人のAPIキーやファイルの情報を使用し、要約したいテキストファイルを同じフォルダ内に設置して以下のプロンプトを実行する
************************************
python openai_arrange.py --input 対象のpdf名（.pdfなしで） --api_key 自分のopenAIのAPIキー
************************************
契約しているChatGPTのモデルが違う場合は--model で指定できますが動作未確認. 詳細はOpenAIの公式ドキュメントをご参照ください
要約速度は契約のプランによって変化する
'''

# 必要なPdfminer.sixモジュールのクラスをインポート
from calendar import c
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.layout import LAParams
from io import StringIO
import argparse
# Pdfminer.sixモジュールのクラスをインポート
# PDF本体情報を扱うための機能や属性を提供するクラス
from pdfminer.pdfdocument import PDFDocument, PDFNoOutlines
# 構文解析を実行するクラス
from pdfminer.pdfparser import PDFParser

import argparse
import openai
import sys
import textwrap
import time
import math
MAX_CHUNK_LENGTH = 6000
DEBUG_PRINT = True
CHARACTERS = 1000 #要約の長さを調整している


# 標準組込み関数open()でモード指定をbinaryでFileオブジェクトを取得

def get_arguments():
    parser = argparse.ArgumentParser(description='ChatGPT Text Refinement from pdf to summary txt.')
    parser.add_argument('--input', type=str, required=True,
                        help='Input file path')
    #parser.add_argument('--output', type=str, required=True, default='output.txt',
    #                    help='Output file path')
    parser.add_argument('--api_key', type=str, required=True,
                        help='OpenAI API Key')
    parser.add_argument('--model', type=str, default="gpt-3.5-turbo", 
                        help='GPT model')

    args = parser.parse_args()
    return args

def send_messages(model, content, chunk):
    messages = [
        {"role": "system", "content": content},
        {"role": "user", "content": chunk}
    ]

    if DEBUG_PRINT:
        print("----------------------------------------------------")
        print(messages)

    try:
        # APIにリクエスト
        response = openai.ChatCompletion.create(
            model=model,
            messages=messages,
        )
        print('please wait......')

    except openai.api_resources.abstract.Error as e:
        print(f"Failed to send request to OpenAI API: {str(e)}")
        sys.exit(1)

    return response['choices'][0]['message']['content']


def handle_chunk(model, content, current_chunk):
    results = []
    # テキストの長さが長すぎる場合、その位置で分割
    chunks = textwrap.wrap(
        current_chunk, width=MAX_CHUNK_LENGTH, break_long_words=True)
    for chunk in chunks:
        # APIにリクエスト
        response = send_messages(model, content, chunk)
        if DEBUG_PRINT:
            print("----------------------------------------------------")
            print("response,", response)
        # レスポンスをリストに追加
        results.append(response)
        # APIのレート制限. 1分間に60000トークン以下
        # time.sleep(20)
    return results

def request(model, content, text):
    sentences = text.split("\n")

    results = []
    current_chunk = ""

    for sentence in sentences:
        # 一定の長さに達するまで文章を追加
        if len(current_chunk) + len(sentence) < (MAX_CHUNK_LENGTH - 100):
            current_chunk += sentence + "\n"
        else:
            if len(current_chunk) > 0:
                # ここでチャンクをさらに分割して要約
                chunk_results = handle_chunk(model, content, current_chunk)
                results.extend(chunk_results)
            # 現在のチャンクをリセット．次の文を設定
            current_chunk = sentence + "\n"

    # 最終チャンクの処理
    if current_chunk:
        chunk_results = handle_chunk(model, content, current_chunk)
        results.extend(chunk_results)

    if DEBUG_PRINT:
        print("----------------------------------------------------")
        print("results,", results)
    # 結合して最終的な結果を作成
    final_result = "\n".join(results)
    return final_result

# テキストを半分に分割する関数
def split_text(text):
    length = len(text)
    midpoint = length // 2
    # 半分に分割して返す
    return text[:midpoint], text[midpoint:]



def extract_txt(filename, output_filename):
    # 標準組込み関数open()でモード指定をbinaryでFileオブジェクトを取得
    fp = open(filename, 'rb')

    # 出力先をPythonコンソールするためにIOストリームを取得
    #outfp = StringIO() 
    #fileに出力する
    outfp = open(output_filename, 'w', encoding='utf-8')


    # 各種テキスト抽出に必要なPdfminer.sixのオブジェクトを取得する処理

    rmgr = PDFResourceManager() # PDFResourceManagerオブジェクトの取得、インスタンス作成
    lprms = LAParams()          # LAParamsオブジェクトの取得
    device = TextConverter(rmgr, outfp, laparams=lprms)    # TextConverterオブジェクトの取得
    iprtr = PDFPageInterpreter(rmgr, device) # PDFPageInterpreterオブジェクトの取得

    # PDFファイルから1ページずつ解析(テキスト抽出)処理する
    for page in PDFPage.get_pages(fp):
        iprtr.process_page(page)

    #text = outfp.getvalue()  # Pythonコンソールへの出力内容を取得

    outfp.close()  # I/Oストリームを閉じる
    device.close() # TextConverterオブジェクトの解放
    fp.close()     #  Fileストリームを閉じる
    #print(text)  # Jupyterの出力ボックスに表示する

def extract_info(filename):
    # PDFファイルの構文解析を行うプログラム
    # 標準組込み関数open()で解析対象PDFのFileオブジェクトを"Binary"モードで取得
    fp = open(filename, 'rb')

    # PDFParserオブジェクトの取得
    parser = PDFParser(fp)

    # PDFDocumentオブジェクトの取得
    doc = PDFDocument(parser)

    # ----------------------------------------------------------------------
    # 【PDFDocumentオブジェクトの属性確認】

    # ➀PDFの構成情報の取得
    print(doc.catalog)
    # >> {'Type': /'Catalog', 'Pages': <PDFObjRef:1>, 'Outlines': <PDFObjRef:206>, 'PageMode': /'UseOutlines'}

    # ➁PDFの属性情報の取得
    print(doc.info)
    # >> { 'Author': b'atsushi', 'CreationDate': b"D:20210321143519+09'00'",
    #      'ModDate': b"D:20210321143519+09'00'", 'Producer': b'Microsoft: Print To PDF',
    #  'Title': b'Microsoft Word - pdfminer_sample1.docx'}

    # ➂コンテンツ抽出の可否
    print(doc.is_extractable)
    # >> True

    # ----------------------------------------------------------------------
    # 【PDFDocumentオブジェクトのメソッド】

    # ＜目次＞のテキストを抽出する
    try:
        outlines = doc.get_outlines() # get_outlines()メソッドはGeneraterを戻す
        for outline in outlines:
            level = outline[0]    # 目次の階層を取得 <インデックス0>
            title = outline[1]    # 目次のコンテンツを取得 <インデックス1>
            print(level,title)

    except PDFNoOutlines: # 目次がないPDFの場合のエラー処理対策
        print("このコンテンツには目次はありません")


def main():
    args = get_arguments()

    # OpenAIのAPIキーを設定
    openai.api_key = args.api_key
    # ファイル名は，コマンドライン引数
    filename = args.input + ('.pdf')
    output_filename = args.input + ('.txt')
    
    # コマンドライン引数から使用する GPT モデル名を取得
    model = args.model

    #PDFからtxtにする
    extract_txt(filename, output_filename)
    extract_info(filename)

    #以下でtxtから要約txtに
    try:
        with open(output_filename, 'r', encoding='utf-8') as file:
            text = file.read()
    except FileNotFoundError:
        print(f"The file {filename} was not found. Put it in a same folder.")
        sys.exit(1)
    
    prompt = f'"Please provide a summary of the supplied Japanese or English text. The summary should be {CHARACTERS} characters max. Do not generate any future sentences. The output should be a single paragraph and should be written in Japanese.'

    final_result = request(
        model,
        prompt, text)

    if DEBUG_PRINT:
        print("----------------------------------------------------")
        print("final_result,", final_result)

    # final_result の長さが指定された文字数を超える場合に処理を行うループ
    while(len(final_result) > CHARACTERS+500):
        # final_result を半分に分割する
        text1, text2 = split_text(final_result)
        
        # ChatGPT に要求して半分ごとに処理する
        final_result1 = request("gpt-3.5-turbo", prompt, text1)
        final_result2 = request("gpt-3.5-turbo", prompt, text2)
        
        # 結果を結合する
        final_result = final_result1 + final_result2
        
        if DEBUG_PRINT:
            print("----------------------------------------------------")  
            print("final_result:", final_result)

    # 全体を整えるように ChatGPT に頼む
    final_result = request(
        "gpt-3.5-turbo",
        "As a professional proofreader, your task is to refine and correct the provided text, preserving its original meaning. The refined output should not contain any predictions in it, and it should be consolidated into a single paragraph. Importantly, the output must remain in Japanese and not exceed " + str(CHARACTERS) + " characters.",
        final_result)
    
    if DEBUG_PRINT:
        print("----------------------------------------------------")  
        print("final_result,", final_result)
        
    


    try:
        # 結果をファイルに書き込む
        with open(output_filename, 'w', encoding='utf-8') as output_file:
            output_file.write(final_result.replace('\r', '').replace('\n', ''))
        print("\033[32m結果が保存されました．入力ファイル名は", filename, "出力ファイル名は", output_filename, "\033[0m")
    except IOError as e:
        print("\033[31mファイルへの書き込みに失敗しました．エラー内容:", str(e), "\033[0m")
        sys.exit(1)




if __name__ == "__main__":
    main()