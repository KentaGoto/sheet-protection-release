import flet as ft
import os
import zipfile
import re
import pathlib
from pathlib import Path


def remove_sheet_protection(xlsx_file):
    # Create temporary directory.
    temp_dir = Path("temp_unzip")
    temp_dir.mkdir(exist_ok=True)

    # Unzip the xlsx file.
    with zipfile.ZipFile(xlsx_file, 'r') as zfile:
        zfile.extractall(temp_dir)

    # Edit worksheet XML file.
    ws_dir = temp_dir / "xl" / "worksheets"
    sheet_protection_pattern = re.compile(r'<sheetProtection.*?>')

    for ws_file in ws_dir.glob("*.xml"):
        with open(ws_file, "r", encoding="utf-8") as f:
            content = f.read()

        # Remove sheet protection tags.
        content = sheet_protection_pattern.sub('', content)

        # Save changes made.
        with open(ws_file, "w", encoding="utf-8") as f:
            f.write(content)

    # Recompress modified files into new xlsx files.
    new_xlsx_file = xlsx_file.with_stem(xlsx_file.stem + "_unprotected")
    with zipfile.ZipFile(new_xlsx_file, 'w', zipfile.ZIP_DEFLATED) as zfile:
        for root, _, files in os.walk(temp_dir):
            for file in files:
                file_path = Path(root) / file
                zfile.write(file_path, file_path.relative_to(temp_dir))

    # Delete temporary directory.
    for root, dirs, files in os.walk(temp_dir, topdown=False):
        for file in files:
            os.remove(Path(root) / file)
        for dir in dirs:
            os.rmdir(Path(root) / dir)
    os.rmdir(temp_dir)
    

# Recursively process directories.
def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)

def main(page):

    dir_path = ft.TextField(label="Directory", autofocus=True)
    dir = ft.Column()
    done = ft.Column()
    progress = ft.ProgressBar(width=400, color="amber", bgcolor="#eeeeee")

    def btn_click(e):
        page.add(
            progress
        )
        s = dir_path.value
        s = s.strip('\"')
        
        dir.controls.append(ft.Text(f"{s}"))
        dir_path.value = ""
        
        for i in all_files(s):
            xlsx = pathlib.Path(i)
            if xlsx.suffix == ".xlsx":
                xlsx_file = Path(xlsx)
                remove_sheet_protection(xlsx_file)
        
        page.remove(progress) # Hide prgress bar.
        done.controls.append(ft.Text(value="Done!"))
        
        page.update()
        dir_path.focus()
        
    page.add(
        dir_path,
        ft.ElevatedButton("Run", on_click=btn_click),
        dir,
        done,
    )

ft.app(target=main)