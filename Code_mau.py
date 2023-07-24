import time
import pandas as pd
import json
import re
import telnetlib
from ansi2html import Ansi2HTMLConverter
import html2text
import sys

# Thông tin kết nối Telnet
# hostname = '10.153.15.37'
# port = 23  # Thay thế bằng cổng Telnet
# username = 'rftestsdc'
# password = 'rftestsdc'

global config

with open("config.json") as json_config_file:
    config = json.load(json_config_file)

def read_excel_to_json(file_path):
    # Đọc file Excel
    df = pd.read_excel(file_path)
    # Chuyển đổi dữ liệu thành mảng JSON
    json_data = df.to_json(orient='records')

    # Trả về mảng JSON
    return json.loads(json_data)

def save_json_to_excel(json_data, file_path):
    # Tạo DataFrame từ mảng JSON
    df = pd.DataFrame(json_data)
    
    # Lưu DataFrame thành file Excel
    df.to_excel(file_path, index=False)

def sendValue(tn, value, timeout=0):
    time.sleep(timeout)
    tn.write(value.encode('ascii'))
    if value != b'\x1b[B':
        time.sleep(0.5)
        output = tn.read_very_eager().decode('utf-8')
        converter = Ansi2HTMLConverter()
        cleaned_output = converter.convert(output)
        plain_text = html2text.html2text(cleaned_output)

        # Lưu chuỗi plain_text sang file log
        # ...
        if config['log'] == 'on':
            with open('log.txt', 'a') as f:
                f.write(plain_text+'\n<------------new screen---------->\n')

        # In kết quả đã được xử lý
        # print(plain_text)
        # print('<------------newline---------->')
        return plain_text

# Tạo kết nối Telnet
tn = telnetlib.Telnet(config['hostname'], config['port'])
tn.set_debuglevel(1)  # Chỉ để debug, có thể loại bỏ dòng này khi hoạt động bình thường

if config['need_login'] == 'yes':
    # login
    sendValue(tn, '\r', timeout=4)  # Vì đây là login đầu tiên nên phải đợi 3s để load hoàn toàn
    sendValue(tn, config['operator_code'])
    sendValue(tn, '\r')
    sendValue(tn, config['app_username'])
    sendValue(tn, '\r')
    sendValue(tn, config['app_password'])
    sendValue(tn, '\r')
    sendValue(tn, '\r')
    sendValue(tn, '\r')
    sendValue(tn, '\r')

excels = pd.read_excel(config['excel_file_name'])

door = excels.iloc[3, 1][6:].strip()

# Bỏ hàng từ 1 đến 13
column_names = excels.iloc[12]
excels = excels.iloc[13:-1]
excels.columns = column_names

# Bỏ cột bằng chỉ số cột (ở đây ta bỏ cột đầu tiên, tức cột 0)
excels = excels.drop(excels.columns[0], axis=1)

# tạo mảng cho cột kết quả
result_column = []

for index, row in excels.iterrows():
    sendValue(tn, row[config['location_column_name']])
    output = sendValue(tn, '\r')

    # Lay noi dung man hinh de tach chuoi
    isset_pick = False
    for row in output.split("\n"):
        if "Pick" in row:
            isset_pick = True

            row_array = row.strip().split(" ")
            if int(row_array[1]) <= row[config['quantity_column_name']]:
                sendValue(tn, b'\x1b[B')
                sendValue(tn, '\r')
                sendValue(tn, row_array[1])
                sendValue(tn, b'\x1b[B')
                sendValue(tn, row[config['container_column_name']])  # vị trí nhập Container là cột Pack ID trong file excel
                sendValue(tn, b'\x1b[B')
                check_locn_lane = sendValue(tn, '\r')
                
                # Lưu kết quả row_array[1] vào cột config['result_column_name']
                # ...
                # row[config['result_column_name']] = row_array[1]
                result_column.append(row_array[1])
                # Xuất kết quả ra màn hình: "Picked row_array[1] CS from location config['location_column_name']"
                # ...
                print("Picked "+row_array[1]+" CS from location "+config['location_column_name'])
            else:
                sendValue(tn, b'\x1b[B')
                sendValue(tn, b'\x1b[B')
                sendValue(tn, '\r')
                sendValue(tn, row[config['quantity_column_name']])
                sendValue(tn, '\r')
                sendValue(tn, row[config['quantity_column_name']])
                sendValue(tn, b'\x1b[B')
                sendValue(tn, row[config['container_column_name']])
                sendValue(tn, b'\x1b[B')
                check_locn_lane = sendValue(tn, '\r')
                
                # Lưu kết quả row_array[1] vào cột config['result_column_name']
                # ...
                # row[config['result_column_name']] = row_array[1]
                result_column.append(row_array[1])
                # Xuất kết quả ra màn hình: "Picked row_array[1] CS from location config['location_column_name']"
                # ...
                print("Picked "+row_array[1]+" CS from location "+config['location_column_name'])
                
            # Kiểm tra màn hình hiện tại có dòng "Scan Locn/Lane:" hay không?
            for row_locn_lane in check_locn_lane.split("\n"):
                if "Locn/Lane" in row_locn_lane:
                    # Nhập cửa
                    sendValue(tn, door)
                    sendValue(tn, '\r')
                    
                    # Xuất kết quả ra màn hình: "Taked item/s to config['export_door_column_name']"
                    # ...
                    print("Taked item/s to "+door)
                
            # Check Pack ID ở ô kế tiếp, nếu trùng thì nhập From Location, nếu không trùng thì End
            # Có thể sẽ không cần đoạn code này, vì đang test thì có thể nhập đồng thời nhiều Pack ID, khi nào khác cửa mới nhập
            # if excels[i][config['container_column_name']] != excels[i+1][config['container_column_name']]:
            #    sendValue(tn, b'\x1b[B')
            #    sendValue(tn, b'\x1b[B')
            #    sendValue(tn, '\r')
            #    sendValue(tn, excels[i][config['export_door_column_name']])
            #    sendValue(tn, '\r')
                
            break

    # Nếu isset_pick = False thì có nghĩa là màn hình này là màn hình báo lỗi
    if isset_pick == False:
        # row[config['result_column_name']] = "Error!!!"
        result_column.append("Error!!!")
        excels[config['result_column_name']] = result_column
        # save_json_to_excel(excels, 'Result_'+config['excel_file_name'])
        excels.to_excel('Result_'+config['excel_file_name'], index=False)

        print(config['location_column_name']+" error!!!")
        input('Press any key to close...')
        sys.exit(1)

# Ghi xuống file excel
excels[config['result_column_name']] = result_column
excels.to_excel('Result_'+config['excel_file_name'], index=False)
# save_json_to_excel(excels, 'Result_'+config['excel_file_name'])
# ansi_escape = re.compile(r'\x1B\[[0-?]*[ -/]*[@-~]')
# output_clean = ansi_escape.sub('', output)
# Nhận kết quả

# Đóng kết nối
tn.close()

print('Done!')
input('Press any key to close...')
sys.exit(1)
