import selenium.webdriver.support.expected_conditions as EC  # noqa
import signal
import readchar

from process import process


def handler(signum, frame):
    msg = "\nBạn có muốn thoát không/Do you want to exit? Y/N\n"
    print(msg, end="", flush=True)
    res = readchar.readchar()
    if res == 'y' or res == 'Y':
        print("Đang thoát/Exiting...")
        exit(0)


def main():
    signal.signal(signal.SIGINT, handler)
    type_visa = ''
    command = ['1', '2', '3', '4', 'q', 'Q']

    date = input('Nhập ngày tháng định dạng dd/mm/yyyy: ').split('/')
    while type_visa not in command:
        type_visa = input("Chọn loại visa cần xử lý/Please select a visa type to process:\n"
                          "1. 1 tháng 1 lần/1 month (Single) \n"
                          "2. 1 tháng nhiều lần/1 month (Multi)\n"
                          "3. 3 tháng 1 lần/3 months (Single)\n"
                          "4. 3 tháng nhiều lần/3 months (Multi)\n"
                          "Lựa chọn/Select: ")
        if type_visa in command:
            if type_visa == 'q' or type_visa == 'Q':
                print('Đang thoát/Exiting...')
            else:
                print('Đang mở trình duyệt Google Chrome/Opening Google Chrome... (Nhấn Ctrl + C để thoát)')
        else:
            print('Lựa chọn không hợp lệ/Command not found\n')

    process(date, type_visa)



if __name__ == "__main__":
    main()
