import os
import requests
from datetime import datetime
from openpyxl import load_workbook
import pandas as pd
import shutil


def reformat_sheet(path, filename):
    file_path = os.path.join(path, filename)
    wb = load_workbook(filename=file_path)
    ws = wb.active
    # iterate thru all cells and if hyperlink found attempt modification of cell
    for row in ws.rows:
        for cell in row:
            try:
                if len(cell.hyperlink.target) > 0:
                    cell.value = "".join([cell.value, "|", cell.hyperlink.target])
    # Join cell.value and hyperlink target into string (optionally just assign the hyperlink.target to the cell.value
            except:
                pass
    temp_filename = os.path.join(path, "temp" + filename)
    wb.save(temp_filename)

    # read with pandas
    data = pd.read_excel(temp_filename)
    # take DataSeries and rsplit by "|" and expand to 2 columns
    hyper1 = (data.sfz_photo_1.str.rsplit("|", expand=True))
    hyper1.columns = ["sfz_photo_1_name", "sfz_photo_1_link"]
    hyper2 = (data.sfz_photo_2.str.rsplit("|", expand=True))
    hyper2.columns = ["sfz_photo_2_name", "sfz_photo_2_link"]
    # join them back to dataframe on index
    data_new = data[["uploader_id", "name", "sfz_number"]]
    data_new = data_new.join(hyper1, how="left").join(hyper2, how="left")
    new_filename = os.path.join(path, "new_" + filename)
    data_new.to_excel(new_filename)
    print("Successfully saved " + new_filename)
    return new_filename


def save_imgs(path, filename):
    file_path = os.path.join(path, filename)
    df = pd.read_excel(file_path)
    root_path = os.path.join(path, "sfz_package")
    if not os.path.exists(root_path):
        os.mkdir(root_path)
    for index, row in df.iterrows():
        new_path = os.path.join(root_path, str(row["name"])+"_"+str(row["sfz_number"]))
        if not os.path.exists(new_path):
            os.mkdir(new_path)
        img_1_filepath = os.path.join(new_path, "正面.jpeg")
        with open(img_1_filepath, 'wb') as handle:
            response = requests.get(row["sfz_photo_1_link"], stream=True)
            if not response.ok:
                print(response)
            for block in response.iter_content(1024):
                if not block:
                    break
                handle.write(block)
        img_2_filepath = os.path.join(new_path, "反面.jpeg")
        with open(img_2_filepath, 'wb') as handle:
            response = requests.get(row["sfz_photo_2_link"], stream=True)
            if not response.ok:
                print(response)
            for block in response.iter_content(1024):
                if not block:
                    break
                handle.write(block)
    now = datetime.now()
    dt_string = now.strftime("%Y%m%d_%H%M%S")
    new_filename = os.path.join(path, 'sfz_package_'+dt_string)
    shutil.make_archive(new_filename, 'zip', root_path)
    print("Successfully saved " + new_filename)
    return new_filename


def extract_and_save(path, filename):
    new_sheet = reformat_sheet(path, filename)
    zipped_file = save_imgs(path, new_sheet)
    print("Successfully saved " + zipped_file)
    return 0


if __name__ == "__main__":
    dir_path = input("Path: ")
    file = input("Filename: ")
    extract_and_save(dir_path, file)
