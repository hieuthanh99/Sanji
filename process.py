from selenium.webdriver.remote.webdriver import By
import selenium.webdriver.support.expected_conditions as EC  # noqa
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import Select
import datetime
import xlwings as xw


def process(sheet, data):
    default_data = 'DataMacDinh.xlsx'
    stay30days = ['1Thang1Lan', '1T1L', '一月单次', '1ThangNhieuLan', '1TNL', '一月多次']
    stay30 = False
    for s in stay30days:
        if s in sheet.strip():
            stay30 = True
            break
    if stay30:
        stay = 30
    else:
        stay = 90

    multi_entry = ['1ThangNhieuLan', '1TNL', '一月多次', '3ThangNhieuLan', '3TNL', '三月多次']
    multi = False
    for m in multi_entry:
        if m in sheet.strip():
            multi = True
            break

    ws = xw.Book(default_data).sheets['Data']
    obj = uc.Chrome()
    obj.get(
        "https://evisa.immigration.gov.vn/vi_VN/web/guest/khai-thi-thuc-dien-tu/cap-thi-thuc-dien-tu?p_p_id=khaithithucdientu_WAR_eVisaportlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-2&p_p_col_count=1&_khaithithucdientu_WAR_eVisaportlet_view=insert")
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ten").send_keys(data[1])
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ho_tt22").send_keys(data[2])
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ngaySinh").send_keys(datetime.datetime.strptime(
        data[3], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y"))
    if data[4] == "M":
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_gioiTinh-nam").click()
    else:
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_gioiTinh-nu").click()

    Select(obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_quocTichHienTai")).select_by_value(data[5])
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_tonGiao").send_keys("No")
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_email").send_keys(ws.range("A3").value)
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_reEmail").send_keys(ws.range("A3").value)

    if multi:
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_passport_mutil_tt22").click()
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_fromDate").clear()
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_fromDate").send_keys(datetime.datetime.strptime(
        data[6], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y"))
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_toDate").clear()

    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_toDate").send_keys((datetime.datetime.strptime(
        data[6], "%Y-%m-%d %H:%M:%S") + datetime.timedelta(days=stay)).strftime("%d/%m/%Y"))
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_soHoChieu").send_keys(data[7])
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_issueDateTt22").send_keys(datetime.datetime.strptime(
        data[8], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y"))
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ngayCoGiaTri").send_keys(datetime.datetime.strptime(
        data[9], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y"))

    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_diaChiLienHeTt22").send_keys(data[10])

    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_soNgayTamTru").clear()

    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_soNgayTamTru").send_keys(str(stay))
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ngayNhapCanh").clear()
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ngayNhapCanh").send_keys(datetime.datetime.strptime(
        data[6], "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%Y"))
    Select(obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_cuaKhauNhapCanh")).select_by_value(ws.range(
        "B3").value)
    Select(obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_cuaKhauXuatCanh")).select_by_value(ws.range(
        "C3").value)
    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_diaChi").send_keys(ws.range("D3").value)
    Select(obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_maTtTT")).select_by_value(str(int(ws.range(
        "E3").value)))

    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_chiphidukien_tt22").send_keys("1000")

    obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_agree").click()

