from glob import glob
from os import path

import pdfkit


def html_to_pdf(htmls, to_file):
    # # 将wkhtmltopdf.exe程序绝对路径传入config对象
    # path_wkthmltopdf = r'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'
    # config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)
    # # 设置方式
    # options = {
    #     'encoding': "utf-8"
    # }
    # # 生成pdf文件，to_file为文件路径
    # pdfkit.from_file(htmls, to_file, configuration=config, options=options)
    pdfkit.from_file(htmls, to_file)
    print(f'html_to_pdf {to_file} done.')


# html_to_pdf(['0.html','1.html','2.html'],'out_2.pdf')

# print(glob(r"data/rongchuang/*/*/*_doc"))


html_dirs = glob(r"data/rongchuang/*/*/*_doc") + glob(r"data/rongchuang/*/*/*_docx") + glob(
    r"data/rongchuang/*/*/*_xls") + glob(r"data/rongchuang/*/*/*_xlsx")
# print(html_dirs)
for html_dir in html_dirs:
    html_files = sorted(glob(html_dir + r"/*.html"), key=path.getmtime)
    basename_dir = path.basename(html_dir)
    # print(path.basename(html_dir))
    # print(html_files, len(html_files))
    output_file = path.join(html_dir, f"{basename_dir}.pdf")
    # print(output_file)
    # input()
    html_to_pdf(html_files, output_file)

