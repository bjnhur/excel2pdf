import pandas as pd
import pdfkit


def excel_to_pdf(excel_path, pdf_path, wkhtmltopdf_path):
    # Excel 파일을 HTML로 변환
    df = pd.read_excel(excel_path)
    html_str = df.to_html()

    # pdfkit 설정: wkhtmltopdf 실행 파일 경로 지정
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

    # HTML을 PDF로 변환
    pdfkit.from_string(html_str, pdf_path, configuration=config)
    print(f"Excel file saved as PDF at {pdf_path}")


# wkhtmltopdf 경로 (절대 경로)
wkhtmltopdf_path = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"

# 예시 사용
excel_path = "results.xlsx"
pdf_path = "results.pdf"
excel_to_pdf(excel_path, pdf_path, wkhtmltopdf_path)
