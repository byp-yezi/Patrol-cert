import ssl
import openpyxl
from openpyxl import workbook, load_workbook
from cryptography import x509
from cryptography.hazmat.backends import default_backend
import socket
from datetime import datetime


def get_certificate_expiration(domain):
    try:
        with socket.create_connection((domain, 443), timeout=3) as sock:
            cert = ssl.get_server_certificate((domain, 443))
            cert_bytes = cert.encode('utf-8')
            cert_obj = x509.load_pem_x509_certificate(cert_bytes, default_backend())
            expiration_date = cert_obj.not_valid_after
            return expiration_date
    except (ssl.CertificateError, socket.gaierror, socket.timeout, TimeoutError, ConnectionRefusedError,
            ConnectionResetError, OSError) as e:
        return e


def main():
    # 打开Excel文件
    workbook = openpyxl.load_workbook('domains.xlsx')

    # 选择要操作的工作表
    sheet = 'abc.com'
    worksheet = workbook[sheet]

    # 获取域名列表，假设域名列表从第1行开始（行索引从1开始）
    domain_column = 1  # 列A
    start_row = 1

    # 计算当前时间
    current_time = datetime.now()

    # 循环遍历域名列表
    for row in range(start_row, worksheet.max_row + 1):
        domain_cell = worksheet.cell(row=row, column=domain_column)
        domain = domain_cell.value
        print(domain)

        expiration_date = get_certificate_expiration(domain)

        if isinstance(expiration_date, datetime):
            # 计算过期日期与当前日期之间的差距
            days_until_expiration = (expiration_date - current_time).days

            # 将过期日期写入Excel文件
            expiration_date_cell = worksheet.cell(row=row, column=domain_column + 1)  # 在B列写入过期日期
            expiration_date_cell.value = expiration_date

            # 如果过期时间小于30天，将单元格标记为红色
            if days_until_expiration < 30:
                red_fill = openpyxl.styles.PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                expiration_date_cell.fill = red_fill
        else:
            # 标注错误信息并标记为红色
            error_message = str(expiration_date)
            error_cell = worksheet.cell(row=row, column=domain_column + 1)
            error_cell.value = error_message
            error_fill = openpyxl.styles.PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
            error_cell.fill = error_fill

    # 保存Excel文件
    results_file = f'{current_time.strftime("%Y-%m-%d_%H-%M-%S")}_{sheet}.xlsx'
    workbook.save(results_file)

    # 打开已保存的工作簿并设置活动工作表
    saved_workbook = load_workbook(results_file)
    sheet_to_open = saved_workbook[sheet]
    saved_workbook.active = sheet_to_open
    saved_workbook.save(results_file)


if __name__ == '__main__':
    main()