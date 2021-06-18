# 批量转换pdf为xlsx
pdf2xlsx_bat

## 安装第三方依赖
python -m pip install --upgrade pip

pip install pdfplumber

(替代操作：python -m pip install --upgrade pdfplumber)

pip install xlwt

## 修改配置文件
替换pdf_folder的‘d:/pdf’为你的转换PDF表文件路径

替换xlsx_folder的‘d:/xlsx’为你的xlsx文件输出目标路径

## 执行
python pdf2xlsx_bat.py
