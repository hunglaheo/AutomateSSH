import pandas as pd
import json

def read_excel_to_json(file_path):
	# Đọc file Excel
    df = pd.read_excel(file_path)

    print(df.iloc[3,1][6:].strip())

    #string[6:].strip()

    # Bỏ hàng từ 1 đến 13
    column_names = df.iloc[12]
    df = df.iloc[13:-1]
    df.columns = column_names

    # Bỏ cột bằng chỉ số cột (ở đây ta bỏ cột đầu tiên, tức cột 0)
    df = df.drop(df.columns[0], axis=1)

    #print(df)

    for index, row in df.iterrows():
        print(row['Qty'])
        #for column in df.columns:
            #print(f"{column}: {row[column]}")
        #print()  # In một dòng trống sau mỗi hàng

	# Chuyển đổi dữ liệu thành mảng JSON
    json_data = df.to_json(orient='records')

	# Trả về mảng JSON
    return json.loads(json_data)

excels = read_excel_to_json('test1.xlsx')

#print(excels)