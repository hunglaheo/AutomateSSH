import pandas as pd

# Đường dẫn tới file Excel của bạn
file_path = 'test1.xlsx'

# Đọc file Excel vào DataFrame
df = pd.read_excel(file_path)

#Lấy thông tin cột Door
print(df.iloc[3,1])
# Bỏ hàng từ 1 đến 13

column_names = df.iloc[12]
df = df.iloc[13:-1]
df.columns = column_names


# Bỏ cột bằng chỉ số cột (ở đây ta bỏ cột đầu tiên, tức cột 0)
df = df.drop(df.columns[0], axis=1)
# tạo mảng cho cột kết quả
result_column = []
# vòng lặp for để lướt từng phần tử
for index, row in df.iterrows():
    for column in df.columns:
        print(f"{column}: {row[column]}")
    # thêm kết quả vào
    result_column.append("oke")
    print()  # In một dòng trống sau mỗi hàng
# add kết quả vào df
df["kết quả"] = result_column
# Lưu DataFrame thành file Excel mới
output_file_path = 'file_moi.xlsx'
df.to_excel(output_file_path, index=False)