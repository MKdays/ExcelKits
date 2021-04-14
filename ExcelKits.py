#-*- coding: utf-8 -*-

#Title
app_ver = "ExcelKits 0.4.1" #메인윈도우 제목
app_sub_box = "Screenshot : Box" #서브윈도우 제목
app_sub_drag = "Screenshot : Drag" #서브윈도우 제목
qt_class_name = "Qt5152QWindowIcon" #QT 클래스네임

#경로 설정 : 순서 변경 금지
import sys, os
sys.path.append(os.path.join(os.path.dirname(sys.argv[0]), "Lib"))  #환경변수 추가, 경로 앞부분 슬래시 금지
app_path = os.path.join(os.path.dirname(sys.argv[0])) #현재 실행 위치 확인
os.chdir(app_path) #작업 경로를 실행 위치로 변경
os.add_dll_directory(app_path+"/Lib/") #라이브러리 경로 추가

#Python 라이브러리
from pathlib import Path
import re #캡쳐 파일명 특수문자 필터
from time import * #스크린샷 시간 표시
import webbrowser #브라우저
from itertools import chain #리스트 데이터
from datetime import * #위크넘버 확인
import configparser #ini 설정 저장

#QT
from PySide2.QtWidgets import *
from PySide2.QtGui import *
from PySide2 import QtCore, QtWidgets, QtGui, QtXml, QtUiTools

#QT리소스
from qt_resource import resource

#Pywin32 라이브러리
from win32api import GetCursorPos #마우스 좌표 추적
import win32com.client #엑셀 ppt 제어
import win32gui #윈도우창 제어
import win32con #윈도우창 제어
import pythoncom #ROT

#기타 라이브러리
from system_hotkey import SystemHotkey #단축키

#중복실행 방지
def run_check():
    single_run = win32gui.FindWindow(qt_class_name, app_ver) #실행 중인 Qt윈도우
    if single_run == 0 : pass #실행중이 없으면 패스
    else : #실행중이라면
        win32gui.ShowWindow(single_run, win32con.SW_SHOWNORMAL)
        win32gui.SetForegroundWindow(single_run)
        sys.exit() #종료
    del single_run
run_check()

class CLASS_UI_LOADER(QtUiTools.QUiLoader):
    def __init__(self, base_instance):
        QtUiTools.QUiLoader.__init__(self, base_instance)
        self.base_instance = base_instance
    def createWidget(self, class_name, parent=None, name=""):
        if parent is None and self.base_instance:
            return self.base_instance
        else:
            widget = QtUiTools.QUiLoader.createWidget(self, class_name, parent, name)
            if self.base_instance:
                setattr(self.base_instance, name, widget)
            return widget    

#CLASS : 메인윈도우
class CLASS_MAINWINDOW(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui_setup()
        self.ui_basic()
        self.ui_shadow()
        self.settings()
        self.definitions()
        self.menu_signal()
        self.signals()

    def ui_setup(self):
        loader = CLASS_UI_LOADER(self)
        loader.load(app_path+"/Lib/main.ui")

    def ui_basic(self):
        self.setWindowTitle(app_ver) #윈도우 제목표시줄
        self.setWindowIcon(QIcon(":/__resource__/image/icon_main.png")) #아이콘 경로

    def ui_shadow(self):
        self.setFixedSize(self.frame.width()+20,self.frame.height()+20) #윈도우 사이즈 : Qframe한변 +10px
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground) #윈도우 불필요 영역 투명처리 및 클릭가능하게
        shadow = QtWidgets.QGraphicsDropShadowEffect(self, blurRadius=10, offset=(QtCore.QPointF(0,0)))
        self.container = QtWidgets.QWidget(self)
        self.container.setStyleSheet("background-color: white;")
        self.container.setGeometry(self.rect().adjusted(10, 10, -10, -10)) #Qframe한변 +10px 보정
        self.container.setGraphicsEffect(shadow)
        self.container.lower()

    def settings(self):
        self.setWindowFlags(QtCore.Qt.Window|QtCore.Qt.FramelessWindowHint|QtCore.Qt.WindowMinMaxButtonsHint)#윈도우 타이틀/프레임 숨기기
        self.tab_main.setStyleSheet("QTabWidget::pane {background: white;border: 0px solid;margin-right: 1px;margin-bottom: 1px;}\
            QTabBar::tab{color: blue;font: 30px; height: 30px;background: transparent;border: 0px solid;width: 0;}") #ui파일에서 tab 위젯을 숨김
        self.range_input.setVisible(False) #토글 텍스트박스 초기 숨김처리
        self.ws_str_in.setVisible(False) #토글 텍스트박스 초기 숨김처리
        self.sc_load_run() #설정불러오기 : ini
        self.context_menu() #컨텍스트 메뉴 정의
        #테이블 헤더 사이즈 자동 맞춤
        header = self.file_control.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)

    def definitions(self): #인스턴스/변수 정의
        self.excel_file_list = [] #파일리스트
        self.file_place = "c:/" #파일다이얼로그 초기 위치
        self.instance_title = CLASS_TITLE(self) #제목표시줄 인스턴스, 드래그 기능 구현
        self.instance_title.setFixedSize(390,50) #제목표시줄 사이즈 : 디자인 맞게 조정필요
        self.instance_message = CLASS_MESSAGE()
        self.instance_message.setParent(self)
        self.instance_message.setWindowFlags(QtCore.Qt.FramelessWindowHint | QtCore.Qt.Dialog)

    def menu_signal(self):
        self.radio_0.pressed.connect(lambda : self.radio_run(0,"Chart Style"))
        self.radio_1.pressed.connect(lambda : self.radio_run(1,"Merge"))
        self.radio_2.pressed.connect(lambda : self.radio_run(2,"Stack"))
        self.radio_3.pressed.connect(lambda : self.radio_run(3,"Function"))
        self.radio_4.pressed.connect(lambda : self.radio_run(4,"PowerPoint"))
        self.radio_5.pressed.connect(lambda : self.radio_run(5,"ScreenShot"))
        self.radio_6.pressed.connect(lambda : self.radio_run(6,"Shortcut"))
        self.radio_7.pressed.connect(lambda : self.radio_run(7,"About"))

    def signals(self):
        self.color_opt1.clicked.connect(self.color_opt1_run) #차트 색상표1
        self.color_opt7.clicked.connect(self.color_opt7_run) #차트 색상표7
        self.color_opt8.clicked.connect(self.color_opt8_run) #차트 색상표8
        self.click_opt1.clicked.connect(self.opt1_run) #차트 실행1
        self.click_opt2.clicked.connect(self.opt2_run) #차트 실행2
        self.click_opt3.clicked.connect(self.opt3_run) #차트 실행3
        self.click_opt4.clicked.connect(self.opt4_run) #차트 실행4
        self.click_opt5.clicked.connect(self.opt5_run) #차트 실행5
        self.click_opt6.clicked.connect(self.opt6_run) #차트 실행6
        self.click_opt7.clicked.connect(self.opt7_run) #차트함실행7
        self.click_opt8.clicked.connect(self.opt8_run) #차트 실행8
        self.file_open_btn.clicked.connect(self.file_open) #파일오픈
        self.file_remove_btn.clicked.connect(self.file_remove) #파일리스트삭제
        self.file_run_btn.clicked.connect(self.file_run) #파일병합실행
        self.btn_only_list.clicked.connect(lambda : self.only_list_run(True)) #한줄정렬
        self.btn_sel_only_list.clicked.connect(self.only_list_run) #선택영역 한줄정렬
        self.stack_btn.clicked.connect(self.stack_run) #데이터쌓기
        self.pit_btn.clicked.connect(self.pit_run) #피치셀렉트
        self.sep_btn.clicked.connect(self.sep_run) #행분할
        self.chart_btn.clicked.connect(self.l_cnt_run) #차트갯수 자동입력
        self.gr_col_btn.clicked.connect(self.gr_col_run) #그룹 컬러
        self.cln_btn.clicked.connect(self.cln_run) #공백제거
        self.t_box_fix_btn.clicked.connect(self.t_box_fix_run) #ppt font_box_fixed
        self.t_box_free_btn.clicked.connect(self.t_box_free_run) #ppt font_box_free
        self.x_p_btn.clicked.connect(self.x_p_run) #ppt 도형 위치
        self.h_w_btn.clicked.connect(self.h_w_run) #ppt 도형 크기
        self.ctq_btn.clicked.connect(self.ctq_run) #CTQ
        self.array_btn.clicked.connect(self.array_run) #데이터 재배치
        self.sc_box_btn.clicked.connect(self.sc_box_run) #스크린샷 버튼 : 박스
        self.sc_open_btn.clicked.connect(self.sc_open_run) #스크린샷 폴더 열기
        self.cbox_top.stateChanged.connect(self.ontop_run) #항상위로 체크박스
        self.hot1_in.activated.connect(self.hot_run) #단축키 설정
        self.hot2_in.activated.connect(self.hot_run) #단축키 설정
        self.min_btn.clicked.connect(self.min_run) #윈도우 최소화
        self.app_close_btn.clicked.connect(self.app_close_run) #윈도우 닫기
        self.icons8_btn.clicked.connect(lambda : self.url_run("https://icons8.com"))
        self.blog_btn.clicked.connect(lambda : self.url_run("https://blog.naver.com/eliase"))
        self.src_btn.clicked.connect(lambda : self.url_run("https://github.com/MKdays/ExcelKits"))
        self.test_btn.clicked.connect(self.test_run) #테스트버튼
        self.sc_btn_1.clicked.connect(lambda : self.url_run(self.sc_url_in_1.text()))
        self.sc_btn_2.clicked.connect(lambda : self.url_run(self.sc_url_in_2.text()))
        self.sc_btn_3.clicked.connect(lambda : self.url_run(self.sc_url_in_3.text()))
        self.sc_btn_4.clicked.connect(lambda : self.url_run(self.sc_url_in_4.text()))
        self.sc_btn_5.clicked.connect(lambda : self.url_run(self.sc_url_in_5.text()))
        self.sc_btn_6.clicked.connect(lambda : self.url_run(self.sc_url_in_6.text()))
        self.sc_save_btn.clicked.connect(self.sc_save_run)

    def radio_run(self,num,text):
        self.tab_main.setCurrentIndex(num)
        self.title.setText(text)
        getattr(self, "radio_%s"%(num)).setChecked(True)

    #이스터에그
    def test_run(self):
        self.instance_message.popup("Hello!","<B>오늘은 %s주차입니다.</B><br><br>Jan(1) Feb(2) Mar(3) Apr(4) May(5) Jun(6)<br>Jul(7) Aug(8) Sep(9) Oct(10) Nov(11) Dec(12)"%(datetime.today().strftime("%Y-%m-%d (%a) %V")), 1)

    #설정저장 : ini
    def sc_save_run(self):
        try:Path(app_path+"/Lib/excelkits_settings").mkdir() #폴더 없으면 만들기
        except:pass #폴더 있으면 패스
        try :
            config = configparser.ConfigParser()
            config_file_path = app_path + "/Lib/excelkits_settings/settings.ini"
            config.read(config_file_path, encoding="utf-8")

            #Section_N 작성
            for i in range (0,10): #숫자만 수정할 것
                try:config.add_section("Section_"+str(i))
                except:pass

            #입력 or 수정
            for i in range (1,7):
                config.set("Section_"+str(i), "Title", getattr(self, "sc_title_in_"+str(i)).text().replace("%","%%")) #%를 %%로 변환하여 저장
                config.set("Section_"+str(i), "URL", getattr(self, "sc_url_in_"+str(i)).text().replace("%","%%"))

            #입력 or 수정
            for i in range (1,19):
                config.set("Section_0", "Sym_"+str(i), getattr(self, "sym_"+str(i)).text().replace("%","%%")) #%를 %%로 변환하여 저장

            #저장
            configFile = open(config_file_path, "w", encoding="utf-8")
            config.write(configFile)
            configFile.close()

            self.instance_message.popup("알림","설정이 저장되었습니다.", 1)

        except Exception as e:
            self.instance_message.popup("알림",str(e), 1)

    #설정불러오기 : ini
    def sc_load_run(self):
        try:
            config = configparser.ConfigParser()
            config_file_path = app_path + "/Lib/excelkits_settings/settings.ini"
            config.read(config_file_path, encoding="utf-8")
            for i in range (1,7):
                a = config["Section_"+str(i)]["title"]
                b = config["Section_"+str(i)]["url"]
                getattr(self, "sc_title_in_%s"%i).setText(a)
                getattr(self, "sc_url_in_%s"%i).setText(b)
            for i in range (1,19):
                getattr(self, "sym_%s"%i).setText(config["Section_0"]["sym_"+str(i)])
        except:pass

    #컨텍스트 메뉴
    def context_menu(self):
        self.file_control.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        table_copy_run = QAction("Copy List", self.file_control)
        table_clear_run = QAction("Clear All", self.file_control)
        self.file_control.addAction(table_copy_run)
        self.file_control.addAction(table_clear_run)
        table_copy_run.triggered.connect(self.table_copy)
        table_clear_run.triggered.connect(self.table_clear)
    def table_copy(self):
        if self.excel_file_list == []:
            self.instance_message.popup("알림","리스트가 비어있습니다.", 1)
            return
        clip_b=""
        for i in self.excel_file_list:
            clip_b = clip_b + "%s\n"%(i)
        QApplication.clipboard().setText(clip_b)
        self.instance_message.popup("알림","%s개의 리스트가 클립보드에 복사되었습니다."%(len(self.excel_file_list)), 1)

    def table_clear(self):
        self.excel_file_list = [] #리스트 초기화
        self.file_control.setRowCount(0) #Row 초기화
        self.file_control.clearContents() #테이블 내용 클리어

    #함수 : QT 이벤트
    def keyPressEvent(self, event): #키보드 이벤트
        if event.key() == QtCore.Qt.Key_F20: #F20키가 눌리면
            self.sc_drag_run() #스크린샷 함수 실행

    #함수 : 웹브라우저 열기
    def url_run(self,url):
        try : webbrowser.open(url)
        except Exception as e:
            self.instance_message.popup("알림",str(e), 1)

    #함수 : 단축키설정
    def hot_run(self):
        instance_hkey.hkey_update(self.hot1_in.currentText(), self.hot2_in.currentText())

    #함수 : 윈도우 최소화/닫기
    def min_run(self): #윈도우 최소화
        self.showMinimized()
    def app_close_run(self): #윈도우 닫기
        self.close()

    #함수 : ALWAYS ON TOP 기능
    def ontop_run(self, state):
        hwnd = win32gui.FindWindow(None, app_ver) #실행중 프로그램 찾기
        if state == QtCore.Qt.Checked:
            win32gui.SetWindowPos(hwnd,win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)
        else:
            win32gui.SetWindowPos(hwnd,win32con.HWND_NOTOPMOST,0,0,0,0,win32con.SWP_NOMOVE + win32con.SWP_NOSIZE)

    #함수 : 스크린샷 폴더 열기
    def sc_open_run(self):
        try:Path(app_path+"/Lib/Screenshot").mkdir() #스크린샷 폴더 없으면 만들기
        except:pass #스크린샷 폴더 있으면 패스
        webbrowser.open(app_path+"/Lib/Screenshot")

    #함수 : 스크린샷 GUI 전체화면에 띄우기
    def sc_drag_run(self): #드래그모드
        self.sc_file = re.sub('[*|:"></?\n]', '', self.sc_file_in.text()).replace('\\','') #특수문자 필터링, 줄바꿈필터링, 역슬러쉬 필터링
        self.ftype = self.ftype_in.currentText()
        self.showMinimized() #스크린 확보 목적 메인윈도우 최소화, 주의 : 포커스를 가져갈 수 있음
        self.instance_drag = CLASS_DRAG_MAIN() #스크린샷 클래스 박스
        self.instance_drag.showFullScreen() #스크린샷 윈도우 띄움, 포커스 가져옴
        self.instance_drag.setGeometry(QApplication.desktop().geometry()) #멀티모니터의 좌표/사이즈를 반영.. 메인모니터의 좌상단좌표 (0,0)
        hwnd = win32gui.FindWindow(None, app_sub_drag) #실행중 프로그램 찾기
        win32gui.SetWindowPos(hwnd,win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE + win32con.SWP_NOSIZE) #윈도우 항상위로

    def sc_box_run(self): #박스모드
        self.sc_file = re.sub('[*|:"></?\n]', '', self.sc_file_in.text()).replace('\\','') #특수문자 필터링, 줄바꿈필터링, 역슬러쉬 필터링
        self.ftype = self.ftype_in.currentText()
        self.showMinimized() #스크린 확보 목적 메인윈도우 최소화, 주의 : 포커스를 가져갈 수 있음
        self.instance_box = CLASS_BOX_MAIN() #스크린샷 클래스 드래그
        self.instance_box.showFullScreen() #스크린샷 윈도우 띄움, 포커스 가져옴
        self.instance_box.setGeometry(QApplication.desktop().geometry()) #멀티모니터의 좌표/사이즈를 반영.. 메인모니터의 좌상단좌표 (0,0)
        hwnd = win32gui.FindWindow(None, app_sub_box) #실행중 프로그램 찾기
        win32gui.SetWindowPos(hwnd,win32con.HWND_TOPMOST,0,0,0,0,win32con.SWP_NOMOVE + win32con.SWP_NOSIZE) #윈도우 항상위로

    #함수 : 엑셀 바인딩
    def excel_dispatch(self):
        global excel, wb, ws, excel_connect
        excel_connect = False #맨 위에서 선언해줘야 오류나지 않음

        if win32gui.FindWindow("XLMAIN", None) == 0 : #검증1. 엑셀이 실행중인가?
            self.pre_return()
            self.instance_message.popup("알림","실행 중인 엑셀 워크시트가 없습니다.", 1)
            return()

        target_excel = win32gui.FindWindow("XLMAIN", None)
        target_excel_title = win32gui.GetWindowText(target_excel)

        for i in pythoncom.GetRunningObjectTable():
            try_instance = i.GetDisplayName(pythoncom.CreateBindCtx(0), None)
            try:
                wb_try = win32com.client.GetObject(try_instance)
                wb_try.ActiveSheet.Select() #검증2. 편집 여부 체크?
                if target_excel == wb_try.Application.Hwnd: #핸들값이 일치하면 바인딩
                    wb = win32com.client.GetObject(try_instance)
                    excel=wb.Application
                    ws = wb.ActiveSheet
                    excel_connect = True
                    return #정상 바인딩, 함수 중단

            except Exception as e:
                if str(type(e)) == "<class 'pywintypes.com_error'>": #오피스 아님 or 읽기 전용 or 시트 인식 실패"
                    pass
                elif str(type(e)) == "<class 'AttributeError'>": #"오피스이지만 엑셀이 아님 or 시트 편집 중이거나"
                    pass
                else:
                    pass

        self.pre_return() #검증3. 반복문 완료시에도 바인딩 실패 케이스
        self.instance_message.popup("알림", f"{target_excel_title}<br>위 파일의 시트가 인식되지 않았습니다.<br>팝업되는 항목을 점검해주세요.", 1)
        self.instance_message.popup("알림","1) 엑셀에 대화창이 열려있는 경우<br>2) 셀값을 현재 편집 중인 경우<br>3) 읽기전용으로 중복으로 파일이 열린 경우", 1)

    #함수 : 엑셀 완료/중단시 종료 전처리
    def pre_return(self):
        try: #메인앱 활성화
            win32gui.ShowWindow(win32gui.FindWindow(qt_class_name, app_ver), win32con.SW_SHOWNORMAL)
            win32gui.SetForegroundWindow(win32gui.FindWindow(qt_class_name, app_ver))
        except:pass
        global excel, wb, ws, excel_connect
        try:excel.ScreenUpdating = True #스크린 복구
        except:pass
        try:del excel, wb, ws, excel_connect #Com 인스턴스 삭제
        except:pass

    #함수 : PPT Font Box Fixed
    def t_box_fix_run(self):
        ppt_run_check = win32gui.FindWindow("PPTFrameClass", None) #실행 엑셀 체크
        if ppt_run_check == 0 : #실행중인 엑셀이 없으면 실행하지 않음
            self.instance_message.popup("알림","실행 전 PPT 프로그램을 먼저 실행해주세요.", 1)
            return
        try:
            ppt = win32com.client.dynamic.Dispatch("PowerPoint.Application")
            shape = ppt.ActiveWindow.Selection.ShapeRange
            shape.TextFrame.AutoSize = 0
        except:
            self.instance_message.popup("알림","PPT에서 도형이나 텍스트 상자를 선택 후 실행해주세요.", 1)
            return

    #함수 : PPT Font Box Free
    def t_box_free_run(self):
        ppt_run_check = win32gui.FindWindow("PPTFrameClass", None) #실행 엑셀 체크
        if ppt_run_check == 0 : #실행중인 엑셀이 없으면 실행하지 않음
            self.instance_message.popup("알림","실행 전 PPT 프로그램을 먼저 실행해주세요.", 1)
            return
        try:
            ppt = win32com.client.dynamic.Dispatch("PowerPoint.Application")
            shape = ppt.ActiveWindow.Selection.ShapeRange
            shape.TextFrame.AutoSize = 1
        except:
            self.instance_message.popup("알림","PPT에서 도형이나 텍스트 상자를 선택 후 실행해주세요.", 1)
            return

    #함수 : PPT 도형 위치
    def x_p_run(self):
        ppt_run_check = win32gui.FindWindow("PPTFrameClass", None) #실행 엑셀 체크
        if ppt_run_check == 0 : #실행중인 엑셀이 없으면 실행하지 않음
            self.instance_message.popup("알림","실행 전 PPT 프로그램을 먼저 실행해주세요.", 1)
            return
        try:
            ppt = win32com.client.dynamic.Dispatch("PowerPoint.Application")
            shape = ppt.ActiveWindow.Selection.ShapeRange
            shape.Left = float(self.x_p_in.text())*28.3464566929
            shape.Top = float(self.y_p_in.text())*28.3464566929
        except:
            self.instance_message.popup("알림","PPT에서 도형이나 텍스트 상자를 선택 후 실행해주세요.", 1)
            return

    #함수 : PPT 도형 크기
    def h_w_run(self):
        ppt_run_check = win32gui.FindWindow("PPTFrameClass", None) #실행 엑셀 체크
        if ppt_run_check == 0 : #실행중인 엑셀이 없으면 실행하지 않음
            self.instance_message.popup("알림","실행 전 PPT 프로그램을 먼저 실행해주세요.", 1)
            return
        try:
            ppt = win32com.client.dynamic.Dispatch("PowerPoint.Application")
            shape = ppt.ActiveWindow.Selection.ShapeRange
            shape.LockAspectRatio = False #Ratio unlock
            shape.Height = float(self.h_in.text())*28.3464566929
            shape.Width = float(self.w_in.text())*28.3464566929
        except:
            self.instance_message.popup("알림","PPT에서 도형이나 텍스트 상자를 선택 후 실행해주세요.", 1)
            return

    #함수 : 엑셀 공백 제거 / 프로그레스바 금지

    def cln_run_sub(self, arg): #공백제거 함수
        try: return arg.strip()
        except: return arg

    def cln_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        self.instance_message.popup("확인","작업을 실행합니다.", 2)
        if self.QnA == 2:
            self.pre_return()
            return
        try:
            usdrng = ws.Range(ws.UsedRange, ws.Range("A1:A2")) #빈시트 오류방지 A1A2
            height = usdrng.Rows.Count
            space = usdrng.Columns.Count
            var_raw = usdrng.Formula #엑셀값을 읽어옴
            var = [i for x in var_raw for i in x] #튜플을 리스트로 변환

            data_semi = list(map(self.cln_run_sub, var)) #모든 리스트의 공백 제거

            n = 0
            m = space
            data_final = []

            for i in range(height):
                data_final.append(data_semi[n:m])
                n = n+space
                m = m+space

            usdrng.Formula = data_final
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 그룹 컬러
    def gr_col_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        self.instance_message.popup("확인","작업을 실행합니다.", 2)
        if self.QnA == 2:
            self.pre_return()
            return
        try:

            var = excel.Selection.Value2
            var_cnt = len(var)
            switch = 1
            memo_st = excel.Selection.Address.split(":")[0]
            start_cell = ws.Range(memo_st)
            wid_cnt = len(var[0])

            n = var_cnt+wid_cnt-2 #테이터 갯수 파악
            pbar_value_step = round(10000/n+0.5) #프로그레스바
            pbar_value = 0 #프로그레스바

            for i in range (1,var_cnt):
                if var[i][0] == var[i-1][0]:pass #동일
                else:switch = switch*-1 #전환
                if switch == 1:
                    start_cell = excel.Union(start_cell,ws.Range(memo_st).Resize(i+1,1))
                pbar_value += pbar_value_step #프로그레스바
                QApplication.processEvents() #프로그레스바
                if pbar_value > 10000:self.pbar3.setValue(10000)
                else : self.pbar3.setValue(pbar_value) #프로그레스바

            switch = 1
            for i in range (1,wid_cnt):
                if var[0][i] == var[0][i-1]:pass #동일
                else:switch = switch*-1 #전환
                if switch == 1:
                    start_cell = excel.Union(start_cell,ws.Range(memo_st).Resize(1,i+1))
                pbar_value += pbar_value_step #프로그레스바
                QApplication.processEvents() #프로그레스바
                if pbar_value > 10000:self.pbar3.setValue(10000)
                else : self.pbar3.setValue(pbar_value) #프로그레스바

            start_cell.Select()

            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 행분할
    def sep_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        self.instance_message.popup("확인","선택한 영역의 첫번째 열 기준으로 작업을 실행합니다.", 2)
        if self.QnA == 2:
            self.pre_return()
            return
        try:
            excel.ScreenUpdating = False #스크린업데이트 끄기
            cell_r = ws.Range(excel.Selection.Address.split(":")[0]).Row
            cell_c = ws.Range(excel.Selection.Address.split(":")[0]).Column
            sep_t = self.sep_t_in.text() #분할 기준 문자 지정
            n = excel.Selection.Rows.Count #테이터 갯수 파악
            pbar_value_step = round(10000/n+0.5) #프로그레스바
            pbar_value = 0 #프로그레스바
            for i in range (n, 0, -1): #밑에서 위로 올라감
                w_cnt = 0
                T_CNT_SUB = 0
                T = ws.Cells(cell_r, cell_c).Offset(i,1).Value2 #실행될 타겟 열
                try:
                    T_CNT = T.count(sep_t)
                except:
                    T_CNT = 0
                try:
                    T_list = T.split(sep_t)
                except:
                    pass
                if T_CNT > 0 :
                    T_CNT_SUB = T_CNT+1
                while T_CNT > 0 :
                    excel.Rows(cell_r+i-1).EntireRow.Insert()
                    excel.Rows(cell_r+i-1).Value2 = excel.Rows(cell_r+i).Value2
                    T_CNT = T_CNT-1
                while T_CNT_SUB > 0 :
                    excel.Cells(cell_r+w_cnt, cell_c).Offset(i,1).Value2 = T_list[w_cnt]
                    w_cnt = w_cnt+1
                    T_CNT_SUB = T_CNT_SUB-1
                pbar_value += pbar_value_step #프로그레스바
                QApplication.processEvents() #프로그레스바
                if pbar_value > 10000:self.pbar3.setValue(10000)
                else : self.pbar3.setValue(pbar_value) #프로그레스바
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 피치 셀렉트
    def pit_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            r1 = ws.Range(excel.Selection.Address.split(":")[0]).Row
            c1 = ws.Range(excel.Selection.Address.split(":")[0]).Column
            r2 = excel.Selection.Rows.Count
            c2 = excel.Selection.Columns.Count
            pit_n = int(self.pit_n_in.text())
            x_pit = int(self.x_pit_in.text())
            y_pit = int(self.y_pit_in.text())
            select_r = ws.Range(ws.Cells(r1, c1), ws.Cells(r1,c1).Offset(r2,c2))
            st_cell = ws.Cells(r1+1, c1+1).Offset

            n = pit_n-1 #테이터 갯수 파악
            pbar_value_step = round(10000/n+0.5) #프로그레스바
            pbar_value = 0 #프로그레스바

            for x in range (1, pit_n):
                select_r_add = ws.Range(st_cell(x*y_pit,x*x_pit), st_cell(r2,c2).Offset(x*y_pit,x*x_pit))
                select_r = excel.Union(select_r,select_r_add)
                pbar_value += pbar_value_step #프로그레스바
                QApplication.processEvents() #프로그레스바
                if pbar_value > 10000:self.pbar3.setValue(10000)
                else : self.pbar3.setValue(pbar_value) #프로그레스바
            select_r.Select()

            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 데이터 쌓기
    def stack_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        self.instance_message.popup("확인","작업을 실행합니다.", 2)
        if self.QnA == 2:
            self.pre_return()
            return
        try:
            #시트 복제 Start
            ws_new = wb.Worksheets.Add() #워크시트 생성 및 이름지정
            df_height = ws.Range(ws.Cells(1,1), ws.UsedRange).Rows.Count #old워크시트의 row 갯수
            df_width = ws.Range(ws.Cells(1,1), ws.UsedRange).Columns.Count #old워크시트의 col 갯수
            ws_new.Range(ws_new.Cells(1,1),ws_new.Cells(df_height, df_width)).Value2 = ws.Range(ws.Cells(1,1), ws.UsedRange).Value2 #데이터 복사
            #시트 복제 End
            start_cell = ws_new.Range(self.start_cell_input.text())
            data_head = int(self.data_head_input.text())
            n = int(self.table_head_input.text())
            length = int(self.length_input.text())
            data_pitch = int(self.data_pitch_input.text())
            data_n = int(self.data_n_input.text())
            yyy=start_cell.Row
            xxx=start_cell.Column
            for i in range (1,data_n+1):
                ini = excel.Cells(yyy,xxx).Offset(1,1+(-1+i)*data_pitch) #기준셀 XY간소화
                for k in range(1,n):
                    out_data = ws_new.Range(ini.Offset(1,1+k*data_head), ini.Offset(length,k*data_head+data_head))
                    in_data = ws_new.Range(ws_new.Cells(yyy, xxx).Offset(1+k*length,1+(-1+i)*data_pitch),ws_new.Cells(yyy, xxx).Offset(1+k*length+length-1, data_head +(-1+i)*data_pitch))
                    in_data.Value2 = out_data.Value2
            a = ws_new.Range(excel.Cells(yyy, xxx + data_head+(-1+data_n)*data_pitch), excel.Cells(yyy, xxx + (-1+data_n)*data_pitch-1+data_head*n))
            for i in range(1, data_n+1):
                b = ws_new.Range(excel.Cells(yyy, xxx + data_head+(-1+i)*data_pitch), excel.Cells(yyy, xxx + (-1+i)*data_pitch-1+data_head*n))
                a = excel.Union(a, b)
            a.Select()
            excel.Selection.EntireColumn.Delete()
            try:
                wb.ActiveSheet.Name = "Result"
            except:
                try:
                    wb.ActiveSheet.Name = ("Result_"+str(wb.Sheets.Count))
                except:
                    wb.ActiveSheet.Name = ("Result__"+str(wb.Sheets.Count))
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 한줄 정열
    def only_list_run(self, option = False):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        self.instance_message.popup("확인","작업을 실행합니다.", 2)
        if self.QnA == 2:
            self.pre_return()
            return
        try:
            if option == True : ws.UsedRange.Select() #옵션 선택시 시트 전체선택
            if excel.Selection.SpecialCells(2).Count < 2 : raise NotImplementedError #데이터 2개 미만시 에러처리
            var = excel.Selection.Value2
            col_cnt = excel.Selection.Columns.Count
            datas = [(data[i],) for i in range(col_cnt) for data in var if data[i] != None]
            ws_new = wb.Worksheets.Add()
            try:
                wb.ActiveSheet.Name = "Result"
            except:
                try:
                    wb.ActiveSheet.Name = ("Result_"+str(wb.Sheets.Count))
                except:
                    wb.ActiveSheet.Name = ("Result__"+str(wb.Sheets.Count))
            ws_new.Range("A1:A%s"%(len(datas))).Value2 = datas

            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 Merge
    def file_open(self):
        excel_file_add_list = QFileDialog.getOpenFileNames(self, '파일을 선택하세요.', self.file_place, "Sheet(*.csv;*.xlsx;*.xls;*.xlsm);;All(*.*)")[0] #튜플(파일명,필터링정보)
        self.excel_file_list = self.excel_file_list + excel_file_add_list #파일리스트추가
        self.excel_file_list = [self.excel_file_list[i] for i in range(len(self.excel_file_list)) if not self.excel_file_list[i] in self.excel_file_list[:i]] #리스트 중복값 제거
        try: self.file_place = str(Path(excel_file_add_list[0]).parent) #최근 폴더 기억하기
        except:pass #파일 선택하지 않고 종료한 경우
        self.table_update()

    def file_remove(self):
        selected_list = self.file_control.selectedIndexes() #선택된 항목명 확인
        selected_list_row = set( i.row() for i in selected_list )

        for i in sorted(selected_list_row, reverse=True):
            del self.excel_file_list[i]

        self.table_update()

    def table_update(self):
        self.file_control.clearContents() #테이블 초기화
        self.file_control.setRowCount(len(self.excel_file_list)) #갯수 맞춰서 테이블갯수 생성
        for e, i in enumerate(self.excel_file_list): #테이블 업데이트 실행
            self.file_control.setItem(e,0,QTableWidgetItem(str(e+1)))
            self.file_control.setItem(e,1,QTableWidgetItem(Path(i).name))


    def file_run(self):
        if len(self.excel_file_list) == 0: #리스트가 비어있으면 실행하지 않음
            return
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        self.instance_message.popup("확인","작업을 실행합니다.", 2)
        if self.QnA == 2:
            self.pre_return()
            return
        try:
            wb = excel.Workbooks.Add()
            try:
                wb.Worksheets("Sheet1").Name = "Result"
            except:
                wb.Worksheets("Sheet1").Name = ("Result_"+str(wb.Sheets.Count))
            ws = wb.ActiveSheet
            df_space = 0
            excel.ScreenUpdating = False #스크린업데이트 끄기
            pbar_value_step = round(10000/len(self.excel_file_list)+0.5) #프로그레스바
            pbar_value = 0 #프로그레스바

            #체크박스 로직
            opt_direction = "next" #방향
            if self.d_in.isChecked(): opt_direction = "down"

            opt_range = False #범위 타겟
            if self.cbox_range.isChecked():opt_range = self.range_input.text()

            if self.ws_target_cbox.isChecked():ws_tar = self.ws_str_in.text() #워크시트 타겟
            else: ws_tar = int(self.ws_in.text())

            for i in (self.excel_file_list):
                wb2 = excel.Workbooks.Open(i, ReadOnly=True) #새워크북 열기
                try :
                    ws2 = wb2.Worksheets(ws_tar) #새워크시트 설정
                except Exception as e:
                    self.pre_return()
                    if str(e).count("NoneType")>0:self.instance_message.popup("알림","%s<br>동일한 파일명이 현재 실행중이어서 작업이 중단되었습니다."%i, 1)
                    else:self.instance_message.popup("알림","%s<br>현재 파일에 Sheet %s가 없어서 작업이 중단되었습니다."%(i,ws_tar), 1)
                    job_success = False
                    return
                if opt_range == False:
                    df_height = ws2.Range(ws2.Cells(1,1), ws2.UsedRange).Rows.Count #새워크시트의 row 갯수
                    df_width = ws2.Range(ws2.Cells(1,1), ws2.UsedRange).Columns.Count #새워크시트의 col 갯수
                else:
                    df_height = ws2.Range(opt_range).Rows.Count #새워크시트의 row 갯수
                    df_width = ws2.Range(opt_range).Columns.Count #새워크시트의 col 갯수
                if opt_direction == "next": #next 방향으로 파일 합치기
                    if opt_range == False:
                        ws.Range(ws.Cells(2,1+df_space),ws.Cells(1+df_height,df_space+df_width)).Value2 = [tuple(i if i == None else str(i).replace("=","_is_") for i in z) for z in ws2.Range(ws2.UsedRange, ws2.Range("A1:A2")).Value2] #빈시트 오류방지 A1A2
                    else:
                        ws.Range(ws.Cells(2,1+df_space),ws.Cells(1+df_height,df_space+df_width)).Value2 = [tuple(i if i == None else str(i).replace("=","_is_") for i in z) for z in ws2.Range(ws2.Range(opt_range),ws2.Range(opt_range).Resize(1,2)).Value2] #1개셀 오류장비 resize
                    ws.Range(ws.Cells(1, 1+df_space), ws.Cells(1, df_space + df_width)).Value2 = Path(i).name #제목줄에 파일명 입력
                    ws.Range(ws.Cells(1, 1+df_space), ws.Cells(1, df_space + df_width)).Interior.Color = 6250335 #제목줄 색상
                    ws.Range(ws.Cells(1, 1+df_space), ws.Cells(1, df_space + df_width)).Font.ColorIndex = 2 #제목줄 색상
                    df_space = df_space + df_width + 1 #다음 리스트 작업을 위한 df_space 업데이트

                else: #down 방향으로 파일 합치기
                    if opt_range == False:
                        ws.Range(ws.Cells(1+df_space,2),ws.Cells(df_space+df_height,1+df_width)).Value2 = [tuple(i if i == None else str(i).replace("=","_is_") for i in z) for z in ws2.Range(ws2.UsedRange, ws2.Range("A1:A2")).Value2] #빈시트 오류방지 A1A2
                    else:
                        ws.Range(ws.Cells(1+df_space,2),ws.Cells(df_space+df_height,1+df_width)).Value2 = [tuple(i if i == None else str(i).replace("=","_is_") for i in z) for z in ws2.Range(ws2.Range(opt_range),ws2.Range(opt_range).Resize(1,2)).Value2] #1개셀 오류장비 resize
                    ws.Range(ws.Cells(1+df_space,1), ws.Cells(df_space + df_height,1)).Value2 = Path(i).name #제목줄에 파일명 입력
                    ws.Range(ws.Cells(1+df_space,1), ws.Cells(df_space + df_height,1)).Interior.Color = 6250335 #제목줄 색상
                    ws.Range(ws.Cells(1+df_space,1), ws.Cells(df_space + df_height,1)).Font.ColorIndex = 2 #제목줄 색상
                    df_space = df_space + df_height + 1 #다음 리스트 작업을 위한 df_space 업데이트
                wb2.Close(False) #저장없이 종료
                pbar_value += pbar_value_step #프로그레스바
                QApplication.processEvents() #프로그레스바
                if pbar_value > 10000:self.pbar.setValue(10000)
                else : self.pbar.setValue(pbar_value) #프로그레스바
                job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 차트 컬러
    def l_cnt_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                l_cnt = excel.ActiveChart.SeriesCollection().Count
                self.input_start.setValue(1)
                self.input_end.setValue(l_cnt)
                self.instance_message.popup("알림","차트 범위 "+str(l_cnt)+"개 적용했습니다.", 1)
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    def color_opt1_run(self):
        color = QColorDialog.getColor(parent = self)
        if color.isValid():
            self.selectedColor = color
            opt1_r = color.red()
            opt1_g = color.green()
            opt1_b = color.blue()
            self.opt1_rgb = opt1_r + opt1_g*256 + opt1_b*256*256
            self.frame_2.setStyleSheet("background-color: rgb("+str(opt1_r)+","+str(opt1_g)+","+str(opt1_b)+");border-style:solid;border-width:1px;border-color: rgb(200, 200, 200);")

    def color_opt7_run(self):
        color = QColorDialog.getColor(parent = self)
        if color.isValid():
            self.selectedColor = color
            opt7_r = color.red()
            opt7_g = color.green()
            opt7_b = color.blue()
            self.opt7_rgb = opt7_r + opt7_g*256 + opt7_b*256*256
            self.frame_4.setStyleSheet("background-color: rgb("+str(opt7_r)+","+str(opt7_g)+","+str(opt7_b)+");border-style:solid;border-width:1px;border-color: rgb(200, 200, 200);")

    def color_opt8_run(self):
        color = QColorDialog.getColor(parent = self)
        if color.isValid():
            self.selectedColor = color
            opt8_r = color.red()
            opt8_g = color.green()
            opt8_b = color.blue()
            self.opt8_rgb = opt8_r + opt8_g*256 + opt8_b*256*256
            self.frame_3.setStyleSheet("background-color: rgb("+str(opt8_r)+","+str(opt8_g)+","+str(opt8_b)+");border-style:solid;border-width:1px;border-color: rgb(200, 200, 200);")

    def opt1_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            start = self.input_start.value() #반복시작변수
            end = self.input_end.value() #반복끝변수
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                excel.ScreenUpdating = False
                for i in range(int(start),int(end)+1,1): #반복함수
                    excel.ActiveChart.SeriesCollection(i).Border.Color = self.opt1_rgb
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    def opt2_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            start = self.input_start.value() #반복시작변수
            end = self.input_end.value() #반복끝변수
            opt2 = self.input_opt2.text() #변수
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                excel.ScreenUpdating =False
                for i in range(int(start),int(end)+1,1): #반복함수
                    excel.ActiveChart.SeriesCollection(i).Format.Line.Weight = opt2 #함수실행
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    def opt3_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            start = self.input_start.value() #반복시작변수
            end = self.input_end.value() #반복끝변수
            opt3 = self.input_opt3.currentIndex() #변수
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                excel.ScreenUpdating =False
                if opt3 == 0 :
                    for i in range(int(start),int(end)+1,1): #반복함수
                        excel.ActiveChart.SeriesCollection(i).Format.Line.Visible = 0 #함수실행
                else :
                    for i in range(int(start),int(end)+1,1): #반복함수
                        excel.ActiveChart.SeriesCollection(i).Format.Line.Visible = -1 #함수실행
                        excel.ActiveChart.SeriesCollection(i).Format.Line.DashStyle = opt3 #함수실행
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    def opt4_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            start = self.input_start.value() #반복시작변수
            end = self.input_end.value() #반복끝변수
            opt4 = self.input_opt4.currentIndex() #변수
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                excel.ScreenUpdating =False
                for i in range(int(start),int(end)+1,1): #반복함수
                    excel.ActiveChart.SeriesCollection(i).Format.Line.Transparency = opt4/10 #함수실행
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    def opt5_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            start = self.input_start.value() #반복시작변수
            end = self.input_end.value() #반복끝변수
            opt5 = self.input_opt5.currentIndex() #변수
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                excel.ScreenUpdating =False
                for i in range(int(start),int(end)+1,1): #반복함수
                    excel.ActiveChart.SeriesCollection(i).MarkerStyle = opt5 #함수실행
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    def opt6_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            start = self.input_start.value() #반복시작변수
            end = self.input_end.value() #반복끝변수
            opt6 = self.input_opt6.text() #변수
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                excel.ScreenUpdating =False
                for i in range(int(start),int(end)+1,1): #반복함수
                    excel.ActiveChart.SeriesCollection(i).MarkerSize = opt6 #함수실행
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    def opt7_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            start = self.input_start.value() #반복시작변수
            end = self.input_end.value() #반복끝변수
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                excel.ScreenUpdating =False
                for i in range(int(start),int(end)+1,1): #반복함수
                    excel.ActiveChart.SeriesCollection(i).MarkerForegroundColor = self.opt7_rgb
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    def opt8_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        try:
            start = self.input_start.value() #반복시작변수
            end = self.input_end.value() #반복끝변수
            if excel.ActiveChart is None:
                self.instance_message.popup("알림","차트 선택 후 실행해주세요.", 1)
            else:
                excel.ScreenUpdating =False
                for i in range(int(start),int(end)+1,1): #반복함수
                    excel.ActiveChart.SeriesCollection(i).MarkerBackgroundColor = self.opt8_rgb
            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True:pass
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 PPK
    def ctq_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        self.instance_message.popup("확인","작업을 실행합니다.", 2)
        if self.QnA == 2:
            self.pre_return()
            return
        try:
            ws = wb.Worksheets.Add() #워크시트 생성 및 이름지정
            excel.ScreenUpdating = False #스크린업데이트 끄기
            try:
                wb.ActiveSheet.Name = "PPK"
            except:
                try:
                    wb.ActiveSheet.Name = ("PPK_"+str(wb.Sheets.Count))
                except:
                    wb.ActiveSheet.Name = ("PPK__"+str(wb.Sheets.Count))
            ws.Range("A3").Value2 = "Target"
            ws.Range("A4").Value2 = "USL"
            ws.Range("A5").Value2 = "LSL"
            ws.Range("A6").Value2 = "3 Sigma"
            ws.Range("A7").Value2 = "Sample N"
            ws.Range("A8").Value2 = "Average"
            ws.Range("A9").Value2 = "Min"
            ws.Range("A10").Value2 = "Max"
            ws.Range("A11").Value2 = "Ppk"
            ws.Range("A13").Value2 = 1
            ws.Range("A14").Value2 = 2
            ws.Range("A15").Value2 = 3
            ws.Range("A16").Value2 = 4
            ws.Range("A17").Value2 = "..."
            ws.Range("B1").Value2 = "CD_1"
            ws.Range("B2").Value2 = "Lot_A"
            ws.Range("B3").Value2 = 10
            ws.Range("B4").Value2 = 15
            ws.Range("B5").Value2 = 5
            ws.Range("B6").Value2 = "=STDEV(B13:B65536)*3"
            ws.Range("B7").Value2 = "=COUNT(B13:B65536)"
            ws.Range("B8").Value2 = "=AVERAGE(B13:B65536)"
            ws.Range("B9").Value2 = "=MIN(B13:B65536)"
            ws.Range("B10").Value2 = "=MAX(B13:B65536)"
            ws.Range("B11").Value2 = "=MIN((B4-B8)/B6,(B8-B5)/B6)"
            ws.Range("B13").Value2 = "=RANDBETWEEN(5,15)"
            ws.Range("B14").Value2 = "=RANDBETWEEN(5,15)"
            ws.Range("B15").Value2 = "=RANDBETWEEN(5,15)"
            ws.Range("B16").Value2 = "=RANDBETWEEN(5,15)"
            ws.Range("B17").Value2 = "=RANDBETWEEN(5,15)"
            ws.Cells.HorizontalAlignment = 3 #셀 중간 맞춤
            ws.Cells.VerticalAlignment = 2 #셀 중간 맞춤
            ws.Range("A1:A11").Interior.Color = 8421504
            ws.Range("A1:A11").Font.ColorIndex = 2
            ws.Range("A13:A17").Interior.Color = 8421504
            ws.Range("A13:A17").Font.ColorIndex = 2
            ws.Range("B3:B5").Interior.Color = 16247773
            ws.Range("B6:B11").Interior.Color = 15652797
            ws.Range("B3:B11").NumberFormatLocal = "0.00" #소수점 표현

            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

    #함수 : 엑셀 데이터 재배치
    def array_run(self):
        global excel, wb, ws, excel_connect
        self.excel_dispatch()
        if excel_connect == False :
            self.pre_return()
            return
        self.instance_message.popup("확인","선택한 영역의 교차값을 X, Y, Value 형태로 재배치합니다.", 2)
        if self.QnA == 2:
            self.pre_return()
            return

        try:
            if excel.Selection.Columns.Count < 2 or excel.Selection.Rows.Count < 2 : raise NotImplementedError #2x2 미선택시 에러처리
            if excel.Selection.SpecialCells(2).Count < 2 : raise NotImplementedError #데이터 2개 미만시 에러처리
            x_max = excel.Selection.Columns.Count-1 #선택영역의 열갯수
            y_max = excel.Selection.Rows.Count-1 #선택영역의 행갯수
            total_l = y_max*x_max
            var = excel.Selection.Value2

            datas_x = [[data] for data in var[0]] #keep 세로축
            datas_y = [[data[0]] for data in var] #keep 세로축

            del datas_x[0], datas_y[0]

            datas_x_total = [i for i in datas_x for s in range(y_max)]
            datas_y_total = list(chain.from_iterable([datas_y for i in range(x_max)]))

            var = var[1:len(var)]
            data_total = [[data[i]] for i in range(1,x_max+1) for data in var]

            ws_new = wb.Worksheets.Add()
            try:
                wb.ActiveSheet.Name = "Result"
            except:
                try:
                    wb.ActiveSheet.Name = ("Result_"+str(wb.Sheets.Count))
                except:
                    wb.ActiveSheet.Name = ("Result__"+str(wb.Sheets.Count))

            ws_new.Range(excel.Cells(2,1).Resize(1,1), excel.Cells(2,1).Resize(total_l,1)).Value2 = datas_x_total
            ws_new.Range(excel.Cells(2,2).Resize(1,1), excel.Cells(2,2).Resize(total_l,1)).Value2 = datas_y_total
            ws_new.Range(excel.Cells(2,3).Resize(1,1), excel.Cells(2,3).Resize(total_l,1)).Value2 = data_total
            ws_new.Range("A1").Value2 = "X"
            ws_new.Range("B1").Value2 = "Y"
            ws_new.Range("C1").Value2 = "Value"

            job_success = True
        except:
            job_success = False
        finally:
            self.pre_return()
            if job_success == True: self.instance_message.popup("알림","작업을 완료했습니다.", 1)
            else: self.instance_message.popup("알림","설정을 확인해주세요", 1)

#CLASS : 박스
box_x_wid, box_y_wid = 200, 200
class CLASS_BOX_MAIN(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(app_sub_box) #제목표시줄
        self.setWindowIcon(QIcon(":/__resource__/image/icon_sub.png")) #아이콘 경로
        self.setWindowOpacity(0.5) #윈도우창 투명도 (0~1)
        self.setMouseTracking(True) #마우스이동 추적활성화
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint) #제목표시줄 감추기
        self.setCursor(QtGui.QCursor(QtCore.Qt.CrossCursor)) #마우스 커서 변경
        self.captext_box = QLabel("",self) #라벨 생성
        self.captext_box.setFont(QFont("Arial",20))
        self.captext_box.setMouseTracking(True) #라벨에 마우스좌표 트래킹 허용

    def mouseMoveEvent(self, event): #마우스 이벤트
        #self.activateWindow() #마우스 움직일때 현재창 활성화
        self.update() #마우스 움직일때 현재 윈도우화면 갱신

    def paintEvent(self, event): #페인트 이벤트
        desk_d_x = QApplication.desktop().geometry().left() #메인모니터 x축 보정값 설청
        desk_d_y = QApplication.desktop().geometry().top() #메인모니터 y축 보정값 설정
        box_mouse_x,box_mouse_y = GetCursorPos()[0]-desk_d_x,GetCursorPos()[1]-desk_d_y #캡쳐박스좌표+보정
        qp = QtGui.QPainter(self)
        br = QtGui.QBrush(QtGui.QColor(52, 152, 219, 255)) #캡쳐박스영역 색깔 0~255(R,G,B,알파)
        qp.setBrush(br)
        qp.drawRect(QtCore.QRect(box_mouse_x, box_mouse_y, box_x_wid, box_y_wid)) #사각형 그리기
        self.captext_box.setText(str(box_x_wid)+" x "+str(box_y_wid)) #박스사이즈정보
        self.captext_box.setGeometry(QtCore.QRect(box_mouse_x, box_mouse_y-70, 1000, 100)) #캡쳐좌표문자의 텍스트박스 위치/크기

    def mousePressEvent(self, event): #마우스 클릭 이벤트
        if event.button() == QtCore.Qt.RightButton: #우클릭하면 실행취소
            self.close()
            del instance_mainwindow.instance_box

    def mouseReleaseEvent(self, event): #마우스 이벤트
        self.close() #현재 윈도우를 닫음
        self.sc_save() #스크린샷저장 실행

    def sc_save(self): #스크린샷저장
        sc_x0 = GetCursorPos()[0]
        sc_y0 = GetCursorPos()[1]
        sc_xw = box_x_wid
        sc_yw = box_y_wid
        if sc_xw < 0:
            sc_x0 = sc_x0 + sc_xw + 1
            sc_xw = sc_xw*-1
        else:pass
        if sc_yw < 0:
            sc_y0 = sc_y0 + sc_yw + 1
            sc_yw = sc_yw*-1
        else:pass
        screenshot_memory = QApplication.primaryScreen().grabWindow(0,sc_x0,sc_y0,sc_xw,sc_yw)
        img_name = instance_mainwindow.sc_file+"_"+strftime('%m%d_%I%M%S', localtime()) #%Y_%m%d_%I%M%S 기본날짜형식
        try:Path(app_path+"/Lib/Screenshot").mkdir() #스크린샷 폴더 없으면 만들기
        except:pass #스크린샷 폴더 있으면 패스
        try:screenshot_memory.save(app_path+"/Lib/Screenshot/"+img_name+"."+instance_mainwindow.ftype, instance_mainwindow.ftype) #png(무손실), bmp(고화질), jpg(저화질)
        except:pass
        QApplication.clipboard().setPixmap(screenshot_memory) #클립보드로 복사
        instance_mainwindow.showNormal() #메인윈도우 복구
        del instance_mainwindow.instance_box

    def keyPressEvent(self, event): #키보드 이벤트
        global box_x_wid, box_y_wid #사각형 그리기
        if event.key() == QtCore.Qt.Key_Escape: #ESC누르면 종료
            self.close()
            del instance_mainwindow.instance_box
        elif (event.key() == QtCore.Qt.Key_Left) and (event.modifiers() & QtCore.Qt.ControlModifier): box_x_wid = box_x_wid-1
        elif event.key() == QtCore.Qt.Key_Left: box_x_wid = box_x_wid-40
        elif (event.key() == QtCore.Qt.Key_Right) and (event.modifiers() & QtCore.Qt.ControlModifier): box_x_wid = box_x_wid+1
        elif event.key() == QtCore.Qt.Key_Right: box_x_wid = box_x_wid+40
        elif (event.key() == QtCore.Qt.Key_Up) and (event.modifiers() & QtCore.Qt.ControlModifier): box_y_wid = box_y_wid-1
        elif event.key() == QtCore.Qt.Key_Up: box_y_wid = box_y_wid-40
        elif (event.key() == QtCore.Qt.Key_Down) and (event.modifiers() & QtCore.Qt.ControlModifier): box_y_wid = box_y_wid+1
        elif event.key() == QtCore.Qt.Key_Down: box_y_wid = box_y_wid+40
        elif event.modifiers() & QtCore.Qt.ControlModifier : pass #컨트롤키는 박스사이즈 조정해야 하므로 꺼지지 않도록 설정
        else:
            self.close() #나머지 버튼은 모두 종료
            del instance_mainwindow.instance_box
        self.update() #키보드 누들때 현재 윈도우화면 갱신

#CLASS : 드래그
class CLASS_DRAG_MAIN(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(app_sub_drag) #제목표시줄
        self.setWindowIcon(QIcon(":/__resource__/image/icon_sub.png")) #아이콘 경로
        self.setWindowOpacity(0.5) #윈도우창 투명도 (0~1)
        self.setMouseTracking(True) #마우스이동 추적활성화
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint) #제목표시줄 감추기
        self.setCursor(QtGui.QCursor(QtCore.Qt.CrossCursor)) #마우스 커서 변경
        self.click_true = 0

    def mouseMoveEvent(self, event): #마우스 이벤트
        #self.activateWindow() #마우스 움직일때 현재창 활성화 [중요]
        self.update() #마우스 움직일때 현재 윈도우화면 갱신

    def paintEvent(self, event): #페인트 이벤트
        desk_d_x = QApplication.desktop().geometry().left() #메인모니터 x축 보정값 설청
        desk_d_y = QApplication.desktop().geometry().top() #메인모니터 y축 보정값 설정
        if self.click_true == 1:
            drag_paint_x_wid,drag_paint_y_wid = GetCursorPos()[0]-self.drag_click_x,GetCursorPos()[1]-self.drag_click_y
            qp = QtGui.QPainter(self)
            br = QtGui.QBrush(QtGui.QColor(52, 152, 219, 255)) #캡쳐박스영역 색깔 0~255(R,G,B,알파)
            qp.setBrush(br)
            qp.drawRect(QtCore.QRect(self.drag_click_x-desk_d_x, self.drag_click_y-desk_d_y, drag_paint_x_wid, drag_paint_y_wid)) #사각형 그리기
        else:pass

    def mousePressEvent(self, event): #마우스 클릭 이벤트
        if event.button() == QtCore.Qt.RightButton: #우클릭하면 실행취소
            self.close()
            del instance_mainwindow.instance_drag
        self.click_true = 1
        self.drag_click_x,self.drag_click_y = GetCursorPos()[0],GetCursorPos()[1] #캡쳐박스좌표+보정

    def mouseReleaseEvent(self, event): #마우스 릴리즈 이벤트
        self.click_true = 0
        self.drag_rel_x,self.drag_rel_y = GetCursorPos()[0],GetCursorPos()[1]
        self.close() #현재 윈도우를 닫음
        self.sc_save() #스크린샷저장 실행

    def sc_save(self): #스크린샷저장
        sc_x0 = self.drag_click_x
        sc_y0 = self.drag_click_y
        sc_xw = self.drag_rel_x-self.drag_click_x
        sc_yw = self.drag_rel_y-self.drag_click_y
        if sc_xw < 0:
            sc_x0 = sc_x0 + sc_xw
            sc_xw = sc_xw*-1 + 1
        else:sc_xw = sc_xw + 1
        if sc_yw < 0:
            sc_y0 = sc_y0 + sc_yw
            sc_yw = sc_yw*-1 + 1
        else:sc_yw = sc_yw + 1
        screenshot_memory = QApplication.primaryScreen().grabWindow(0,sc_x0,sc_y0,sc_xw,sc_yw)
        img_name = instance_mainwindow.sc_file+"_"+strftime('%m%d_%I%M%S', localtime()) #%Y_%m%d_%I%M%S 기본날짜형식
        try:Path(app_path+"/Lib/Screenshot").mkdir() #스크린샷 폴더 없으면 만들기
        except:pass #스크린샷 폴더 있으면 패스
        try:screenshot_memory.save(app_path+"/Lib/Screenshot/"+img_name+"."+instance_mainwindow.ftype, instance_mainwindow.ftype) #png(무손실), bmp(고화질), jpg(저화질)
        except:pass
        QApplication.clipboard().setPixmap(screenshot_memory) #클립보드로 복사
        instance_mainwindow.showNormal() #메인윈도우 복구
        del instance_mainwindow.instance_drag

    def keyPressEvent(self, event): #키보드 이벤트
        self.click_true = 0
        if event.key() == QtCore.Qt.Key_Escape: #ESC누르면 종료
            self.close()
            del instance_mainwindow.instance_drag
        else:
            self.close() #나머지 버튼은 모두 종료
            del instance_mainwindow.instance_drag
        self.update() #키보드 누들때 현재 윈도우화면 갱신

#CLASS : 타이틀
class CLASS_TITLE(QWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.parent.is_moving = True
            self.parent.offset = event.pos()
    def mouseMoveEvent(self, event):
        try:
            if self.parent.is_moving:self.parent.move(event.globalPos()-self.parent.offset)
        except:pass

#CLASS : 메시지박스
class CLASS_MESSAGE(QDialog):
    def __init__(self):
        super().__init__()
        self.ui_setup()
        self.definitions()
        self.ui_shadow()
        self.signals()

    def ui_setup(self):
        loader = CLASS_UI_LOADER(self)
        loader.load(app_path+"/Lib/msgbox.ui")
        
    def definitions(self): #인스턴스/변수 정의
        self.instance_title = CLASS_TITLE(self) #제목표시줄 인스턴스
        self.instance_title.setFixedSize(370,50) #제목표시줄 사이즈 : 디자인 맞게 조정필요

    def ui_shadow(self):
        self.setFixedSize(self.frame.width()+20,self.frame.height()+20) #윈도우 사이즈 : Qframe한변 +10px
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground) #윈도우 불필요 영역 투명처리 및 클릭가능하게
        shadow = QtWidgets.QGraphicsDropShadowEffect(self, blurRadius=10, offset=(QtCore.QPointF(0,0)))
        self.container = QtWidgets.QWidget(self)
        self.container.setStyleSheet("background-color: white;")
        self.container.setGeometry(self.rect().adjusted(10, 10, -10, -10)) #Qframe한변 +10px 보정
        self.container.setGraphicsEffect(shadow)
        self.container.lower()

    def signals(self):
        self.yes_btn.clicked.connect(self.yes_btn_run)
        self.no_btn.clicked.connect(self.no_btn_run)
        self.msg_close_btn.clicked.connect(self.msg_close_btn_run)

    def popup(self,title,text,select):
        self.move(instance_mainwindow.pos().x()+50,instance_mainwindow.pos().y()+100) #메시지박스 위치
        self.title.setText(title)
        self.text.setHtml("<p align=center vertical-align=middle >%s</p>"  %text) #HTML
        if select == 1:
            self.yes_btn.move(155,120)
            self.yes_btn.setText("OK")
            self.msg_close_btn.setFocus()
            self.no_btn.close()
        else:
            self.yes_btn.move(110,120)
            self.no_btn.move(200,120)
            self.yes_btn.setText("Yes")
            self.yes_btn.setFocus()
            self.no_btn.show()
        super().exec_()

    def yes_btn_run(self):
        self.close()
        instance_mainwindow.QnA = 1 #1YES

    def msg_close_btn_run(self):
        self.close()
        instance_mainwindow.QnA = 2 #2NO

    def keyPressEvent(self, event):
        if event.key() == QtCore.Qt.Key_Escape: self.msg_close_btn_run()

    def no_btn_run(self):
        self.msg_close_btn_run()

#CLASS : 단축키
class CLASS_HKEY:
    def __init__(self):
        self.hk = SystemHotkey() #인스턴스 생성
        self.hot1 = instance_mainwindow.hot1_in.currentText()
        self.hot2 = instance_mainwindow.hot2_in.currentText()
        self.hk.register((self.hot1, self.hot2),callback=lambda x: self.hkey_run())

    def hkey_update(self, arg_hot1, arg_hot2):
        self.hk.unregister((self.hot1, self.hot2))
        self.hot1 = arg_hot1
        self.hot2 = arg_hot2
        self.hk.register((self.hot1, self.hot2),callback=lambda x: self.hkey_run())

    def hkey_run(self):
        hwnd = win32gui.FindWindow(None, app_ver)
        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_F20)

#인스턴스 실행
app = QApplication(sys.argv)
instance_mainwindow = CLASS_MAINWINDOW()
instance_mainwindow.show()
instance_hkey = CLASS_HKEY()
sys.exit(app.exec_())
