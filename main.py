from flask import Flask, render_template, request , send_file
from zipfile import ZipFile, ZIP_DEFLATED
import re
from scraping import KeywordScraper
import excel_or_csv

# flaskアプリの明示
templates_path = "templates/"
static_path = "static/"
app = Flask(__name__ , template_folder=templates_path, static_folder=static_path)

# パスの定義
output_csv_folder = static_path + "csv/"
output_excel_folder = static_path + "excel/"
output_csv_path = output_csv_folder + "output.csv"
output_excel_path = output_excel_folder + "output.xlsx"
output_zip_csv_path = output_csv_folder + "output_csv.zip"
csv_path_list = []


@app.route('/', methods=['GET', 'POST'])
def suggest_vol():
    return render_template("suggest_vol.html")

@app.route('/vol_result', methods=['GET', 'POST'])
def vol_result():
    if request.method == 'POST':
        main_keyword_str = request.form['main_keyword']
        main_keyword_list = re.split(r'\s+', main_keyword_str.strip())
        url = 'https://ruri-co.biz-samurai.com'
        keyword_scraper = KeywordScraper()
        many_kw_output_list = keyword_scraper.scraping(url , main_keyword_list)        

        main_kw_count = len(many_kw_output_list)
        
        # csvファイルをローカルディレクトリに保存
        if main_kw_count > 1:
            # csvファイルをzip化
            csv_path_list = excel_or_csv.many_list_to_csv(many_kw_output_list)
            with ZipFile(output_zip_csv_path, 'w', ZIP_DEFLATED) as zip_file:
                for csv_path in csv_path_list:
                    zip_file.write(csv_path , csv_path)
        else:
            excel_or_csv.list_to_csv(many_kw_output_list[0] , output_csv_path)

        # Excelファイルをローカルディレクトリに保存
        if main_kw_count > 1:
            # 複数のExcelシート化
            excel_or_csv.many_list_to_excel(many_kw_output_list , output_excel_path)
        else:
            excel_or_csv.list_to_excel(many_kw_output_list[0] , output_excel_path)

            
        form_flag = True
        return render_template(
            "suggest_vol.html" ,
            form_flag = form_flag ,
            main_kw_count = main_kw_count ,
        )



@app.route('/csv_download')
def csv_download():
    return send_file(output_csv_path , as_attachment=True)

@app.route('/zip_csv_download')
def zip_csv_download():
    return send_file(output_zip_csv_path , as_attachment=True)

@app.route('/excel_download')
def excel_download():
    return send_file(output_excel_path , as_attachment=True)



if __name__ == "__main__":
    port_number = 6601
    app.run(port = port_number , debug=True)

