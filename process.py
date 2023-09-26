from selenium import webdriver
from selenium.webdriver.remote.webdriver import By
import selenium.webdriver.support.expected_conditions as EC  # noqa
from time import sleep
import undetected_chromedriver as uc
import xlwings as xw
from selenium.webdriver.support.ui import Select
import datetime


def process(date, type_visa):
    day, month, year = [int(item) for item in date]
    if day < 10:
        day = '0' + str(day)
    else:
        day = str(day)
    if month < 10:
        month = '0' + str(month)
    else:
        month = str(month)

    filename = str(year) + "/" + month + "/" + day + "/" + str(year) + str(month) + str(day) + ".xlsx"
    if type_visa == '1':
        sheet = "1Thang1Lan (1T1L)"
        stay = 30
    elif type_visa == '2':
        sheet = "1ThangNhieuLan (1TNL)"
        stay = 30
    elif type_visa == '3':
        sheet = "3Thang1Lan (3T1L)"
        stay = 90
    elif type_visa == '4':
        sheet = "3ThangNhieuLan (3TNL)"
        stay = 90

    if type_visa == 'q' or type_visa == 'Q':
        exit()
    else:
        obj = uc.Chrome()
        ws = xw.Book(filename).sheets[sheet]
        obj.get(
            "https://evisa.immigration.gov.vn/vi_VN/web/guest/khai-thi-thuc-dien-tu/cap-thi-thuc-dien-tu?p_p_id=khaithithucdientu_WAR_eVisaportlet&p_p_lifecycle=0&p_p_state=normal&p_p_mode=view&p_p_col_id=column-2&p_p_col_count=1&_khaithithucdientu_WAR_eVisaportlet_view=insert")
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ten").send_keys(ws.range("B2").value)
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ho_tt22").send_keys(ws.range("C2").value)
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ngaySinh").send_keys(ws.range(
            "D2").value.strftime("%d/%m/%Y"))
        if ws.range("E2").value == "M":
            obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_gioiTinh-nam").click()
        else:
            obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_gioiTinh-nu").click()

        Select(obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_quocTichHienTai")).select_by_value(ws.range(
            "F2").value)
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_tonGiao").send_keys("No")
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_email").send_keys("tkface02@gmail.com")
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_reEmail").send_keys("tkface02@gmail.com")

        if type_visa == '2' or type_visa == '4':
            obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_passport_mutil_tt22").click()
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_fromDate").clear()
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_fromDate").send_keys(ws.range(
            "G2").value.strftime("%d/%m/%Y"))
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_toDate").clear()

        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_toDate").send_keys((ws.range("G2").value +
            datetime.timedelta(days=stay)).strftime("%d/%m/%Y"))
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_soHoChieu").send_keys(ws.range("H2").value)
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_issueDateTt22").send_keys(
            ws.range("I2").value.strftime("%d/%m/%Y"))
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ngayCoGiaTri").send_keys(
            ws.range("J2").value.strftime("%d/%m/%Y"))

        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_diaChiLienHeTt22").send_keys(ws.range("K2").value)

        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_soNgayTamTru").clear()

        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_soNgayTamTru").send_keys(str(stay))
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ngayNhapCanh").clear()
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_ngayNhapCanh").send_keys(ws.range(
            "G2").value.strftime("%d/%m/%Y"))
        Select(obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_cuaKhauNhapCanh")).select_by_value(ws.range(
            "L2").value)
        Select(obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_cuaKhauXuatCanh")).select_by_value(ws.range(
            "M2").value)
        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_diaChi").send_keys(ws.range("N2").value)
        Select(obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_maTtTT")).select_by_value(str(int(ws.range(
            "O2").value)))

        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_chiphidukien_tt22").send_keys(ws.range("P2").value)

        obj.find_element(By.ID, "_khaithithucdientu_WAR_eVisaportlet_agree").click()

        sleep(100000)
        obj.quit()