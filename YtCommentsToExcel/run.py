import YtCommentsToExcel
import sys
import os.path

args = sys.argv

json_path = args[1]
xlsx_path = args[2]

if not os.path.exists(json_path):
    print(f"File {json_path} not found!")
    quit(-1)

YtCommentsToExcel.ytdlp_json_to_excel_group_by_author(json_path, xlsx_path)

print("Done!")
