import pandas as pd
import xlwings as xw
from openpyxl import load_workbook
import shutil


def change_sheets_name(file_path):

    ss = load_workbook(file_path)
    # writer = pd.ExcelWriter(file_path, engine = 'openpyxl')

    excel_data_df = pd.read_excel(file_path, sheet_name='Liste Références', index_col=None)

    sheet_mapping = excel_data_df.set_index('Unnamed: 0')['Numéro du client'].to_dict()

    i = 0

    hyper_value = []

    # Đổi tên sheet
    for k, v in sheet_mapping.items():
        i += 1

        sheet_mapping[k] = v.split(": ")[1]
        new_sheet_name = v.split(": ")[1]

        hyper_value.append(f"=LIEN_HYPERTEXTE(\"#\"&\"'{new_sheet_name}'!A1\", \"click ici\")") 
        # hyper_value.append(f"=HYPERLINK(\"#\"&\"'{new_sheet_name}'!A1\", \"click ici\")") 

        try:
            ss_sheet = ss[k]
            ss_sheet.title = new_sheet_name
        except Exception as e:
            continue

    ss.save(file_path)

    # Thêm link
    # excel_data_df["Hyper"] = hyper_value

    # # with pd.ExcelWriter(file_path, engine='xlsxwriter', mode='a', if_sheet_exists='replace') as writer:  
    # with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:  
    #     excel_data_df.to_excel(writer, sheet_name='Liste Références')

    df = pd.DataFrame(hyper_value, columns=["Hyper"])
    app = xw.App(visible=False)
    wb = xw.Book(file_path)  
    ws = wb.sheets['Liste Références']

    ws.range('G1').options(index=False).value = df

    ws.range("B1").copy()
    ws.range("G1").paste(paste="formats")
    ws.api.Application.CutCopyMode = False

    wb.save()
    wb.close()
    app.quit()




def print_sheets_name(file_path):
    
    xl = pd.ExcelFile(file_path)

    sheet_list = xl.sheet_names

    print(sheet_list)


def print_whole_sheet(file_path, sheet_name):

    excel_data_df = pd.read_excel(file_path, sheet_name=sheet_name)

    print(excel_data_df)


def copy_backup (file_path):
    shutil.copyfile(file_path, f"{file_path.split('.')[0]}-bkup.{file_path.split('.')[1]}")


if __name__ == "__main__":

    file_name = input("Nhập tên file (file phải cùng thư mục): ")
    # file_name = "./data/FormRename.xlsx"

    print("Backup file...")
    copy_backup(file_path=file_name)

    print_whole_sheet(file_path=file_name, sheet_name="Liste Références")

    print("Tên các sheet trước thay đổi")
    print_sheets_name(file_path=file_name)

    print("Thực hiện đổi tên sheet theo thông tin trong sheet 'Liste Références'. Xin hãy chờ ...")

    change_sheets_name(file_path=file_name)
    print("Tên các sheet sau thay đổi")
    print_sheets_name(file_path=file_name)

    print_whole_sheet(file_name, "Liste Références")

    input("Hoàn thành, hãy kiểm tra file.")
