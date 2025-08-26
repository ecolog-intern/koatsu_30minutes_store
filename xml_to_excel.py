import os
from time import sleep
import win32com.client
from pathlib import Path

#filesにxmlのパスを指定するとすべてExcelになる
#実行はコマンドプロンプトでこのファイル自体を実行する

files = [r"\\f-hikari02\法人事業本部\事業開発部\02.アライアンス事業部\31.ガスでん\27.インターン用\20.高圧30分値\高圧実行\フォルダ保存用\関東\202508\20250818\W401202025081800000100.xml", r"\\f-hikari02\法人事業本部\事業開発部\02.アライアンス事業部\31.ガスでん\27.インターン用\20.高圧30分値\高圧実行\フォルダ保存用\九州\202508\20250818\W401202025081800000000.xml", r"\\f-hikari02\法人事業本部\事業開発部\02.アライアンス事業部\31.ガスでん\27.インターン用\20.高圧30分値\高圧実行\フォルダ保存用\中国\202508\20250818\W401202025081800000100.xml", r"\\f-hikari02\法人事業本部\事業開発部\02.アライアンス事業部\31.ガスでん\27.インターン用\20.高圧30分値\高圧実行\フォルダ保存用\中部\202508\20250818\W401202025081800000001.xml", r"\\f-hikari02\法人事業本部\事業開発部\02.アライアンス事業部\31.ガスでん\27.インターン用\20.高圧30分値\高圧実行\フォルダ保存用\東北\202508\20250818\W401202025081800000000.xml", r"\\f-hikari02\法人事業本部\事業開発部\02.アライアンス事業部\31.ガスでん\27.インターン用\20.高圧30分値\高圧実行\フォルダ保存用\北陸\202508\20250818\W401202025081800000000.xml"]
for extracted_file in files:
    # Path オブジェクト化
    p = Path(extracted_file)

    # フォルダパス（ファイルの親ディレクトリ）
    folder_path = str(p.parent)

    # yesterday_str （親フォルダ名から取得）
    yesterday_str = p.parent.name

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # バックグラウンドで実行

    workbook = excel.Workbooks.OpenXML(
        Filename=extracted_file,
        LoadOption=1  
    )

    # 昨日の日付を用いて保存先Excelファイルの名前を決定
    filename = f"{yesterday_str}.xlsx"
    excel_file = os.path.join(folder_path, filename)

    # Excelファイルとして保存（FileFormat=51は.xlsx形式）
    workbook.SaveAs(excel_file, FileFormat=51)
    workbook.Close(SaveChanges=False)
    excel.Quit()
    print(extracted_file)
