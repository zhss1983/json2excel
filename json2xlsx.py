import os

from support import load_json_data
from converter import WorksheetWriter

input_json_file_name = "./output.json"
output_xlsx_file_name = "./output.xlsx"

if not os.path.exists(input_json_file_name):
    raise FileExistsError(f"Не найден файл: {input_json_file_name}")
data = load_json_data(input_json_file_name)
if not data:
    raise ValueError("Данные в файле отсутствуют.")

with WorksheetWriter(output_xlsx_file_name) as writer:
    writer.write(data)

# input_output_file_names = {
#     "./test.json": "./test.xlsx",
#     "./output.json": "./output.xlsx",
# }
#
# for input_json_file_name, output_xlsx_file_name in input_output_file_names.items():
#     if not os.path.exists(input_json_file_name):
#         raise FileExistsError(f"Не найден файл: {input_json_file_name}")
#     data = load_json_data(input_json_file_name)
#     if not data:
#         raise ValueError("Данные в файле отсутствуют.")
#
#     with WorksheetWriter(output_xlsx_file_name) as writer:
#         writer.write(data)
