#경로 설정 : 순서 변경 금지
import sys, os
app_path = os.path.join(os.path.dirname(sys.argv[0])) #현재 위치 확인
os.chdir(app_path) #경로 변경

# 옵션 ------------------------------------------------------
path_search = app_path
path_out = app_path
as_is = "import img_rc"
to_be = ""
# -----------------------------------------------------------

#라이브러리
from pathlib import Path
from time import sleep
import shutil

work_list =[]
for i in os.listdir(path_search):
    if i.endswith(".ui") or i.endswith(".qrc"):
        work_list.append(i)

def ui_convert(i):
  os.system("pyside2-uic " + i + " -o " + path_out + "/" + Path(i).stem + ".py")

def qrc_convert(i):
  os.system("pyside2-rcc " + i + " -o " + path_out + "/" + Path(i).stem + ".py")

def replace_text(i):
  with open(path_out + "/" + Path(i).stem + ".py", 'r', encoding="utf-8") as file :
    filedata = file.read()
  filedata = filedata.replace(as_is, to_be)
  with open(path_out + "/" + Path(i).stem + ".py", 'w', encoding="utf-8") as file:
    file.write(filedata)

for i in work_list:
  print(i + " 완료")
  if i.endswith(".ui"):ui_convert(i)
  elif i.endswith(".qrc"):qrc_convert(i)
  replace_text(i)

shutil.rmtree("__pycache__") #폴더삭제

sleep(0.2)