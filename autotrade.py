import sys
import copy
from collections import deque
from queue import Queue
import datetime

from loguru import logger
import pandas as pd
import numpy as np

from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import Qt, QSettings, QTimer, QCoreApplication, QAbstractTableModel
from PyQt5.QAxContainer import QAxWidget
from PyQt5 import QtGui, uic

form_class = uic.loadUiType("main.ui")[0]

class PandasModel(QAbstractTableModel): # PandasModel은 테이블 뷰를 만들어주는 클래스
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[section]
        if orientation == Qt.Vertical and role == Qt.DisplayRole:
            return self._data.index[section]
        return None

    def setData(self, index, value, role):
        # 항상 False를 반환하여 편집을 비활성화
        return False

    def flags(self, index):
        return Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable


class KiwoomAPI(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show()

        self.conditionInPushButton.clicked.connect(self.condition_in)
        self.conditionOutPushButton.clicked.connect(self.condition_out)
        self.settings = QSettings('My company', 'myApp')
        # My company, myApp에 setting 저장(buyAmountLineEdit, goalReturnLineEdit, stopLossLineEdit) windows 레지스르리에 등록
        self.load_settings()
        # self.setWindowIcon(QtGui.QIcon('icon.ico'))

        self.max_send_per_sec: int = 4 # 초당 TR 호출 최대 4번
        self.max_send_per_minute: int = 55 # 분당 TR 호출 최대 55번
        self.max_send_per_hour: int = 950 # 시간당 TR 호출 최대 950번
        self.last_tr_send_times = deque(maxlen=self.max_send_per_hour)
        self.tr_req_queue = Queue()
        self.orders_queue = Queue()

        self.account_num = None # 계좌번호 초기화
        self.unfinished_order_num_to_info_dict = dict() # 미체결 수량 리스트
        self.stock_code_to_info_dict = dict()

        self.scrnum = 5000
        self.using_condition_name = ""
        #self.realtime_reqisted_codes = []
        self.realtime_reqisted_codes = set()
        self.condition_name_to_condition_idx_dict = dict() # 조건 검색식을 저장해두는 부분
        self.registed_condition_df = pd.DataFrame(columns=["화면번호", "조건식이름"])
        self.registed_conditions_list = []

        # 계좌정보를 담을 dataframe
        self.account_info_df  = pd.DataFrame(
            columns=[
                "종목명",
                "매매가능수량",
                "보유수량",
                "매입가",
                "현재가",
                "수익률",
            ]
        )
        self.is_updated_realtime_watchlist = False
        self.stock_code_to_sell_price_dict = dict()
        try:
            self.realtime_watchlist_df = pd.read_pickle("./realtime_watchlist_df.pkl")
        except FileNotFoundError:
            self.realtime_watchlist_df = pd.DataFrame(
                columns=["종목명", "현재가", "평균단가", "목표가", "손절가", "수익률", "매수기반조건식", "보유수량", "매수주문완료여부"]
            )

        self.registeredTableView

        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1") #  kiwoom api activ x 를 연동시키는 방법
        self._set_signal_slots() # 키움증권 API와 내부 매소드를 연동
        self._login()

        self.timer1 = QTimer()
        self.timer2 = QTimer()
        self.timer3 = QTimer()
        self.timer4 = QTimer()
        self.timer5 = QTimer()
        self.timer6 = QTimer()
        self.timer7 = QTimer()
        self.timer8 = QTimer()

        self.timer1.timeout.connect(self.update_pandas_models)
        self.timer2.timeout.connect(self.send_tr_request)
        self.timer3.timeout.connect(self.send_orders)
        self.timer4.timeout.connect(self.request_get_account_balance)
        self.timer5.timeout.connect(self.request_current_order_info)
        self.timer6.timeout.connect(self.save_settings)
        self.timer7.timeout.connect(self.check_unfinished_orders)
        self.timer8.timeout.connect(self.check_outliers)


    def check_outliers(self):
        pop_list = []
        for row in self.realtime_watchlist_df.itertuples():
            stock_code = getattr(row, "Index")
            목표가 = getattr(row, "목표가")
            손절가 = getattr(row, "손절가")
            보유수량 = getattr(row, "보유수량")

            # if np.isnan(목표가) or np.isnan(손절가):
            #     pop_list.append(stock_code)
            if (isinstance(목표가, (int, float)) and np.isnan(목표가)) or (isinstance(손절가, (int, float)) and np.isnan(손절가)):
                pop_list.append(stock_code)

        for stock_code in pop_list:
            logger.info(f"종목코드: {stock_code}, Outlier!! Pop!!")
            self.realtime_watchlist_df.drop(stock_code, inplace=True)

    def check_unfinished_orders(self):
        pop_list = []
        for order_num, stock_info_dict in self.unfinished_order_num_to_info_dict.items():
            주문번호 = order_num["주문번호"]
            종목코드 = stock_info_dict["종목코드"]
            주문체결시간 = stock_info_dict["주문체결시간"]
            미체결수량 = stock_info_dict["미체결수량"]
            주문가격 = stock_info_dict["주문가격"]

            order_time = datetime.datetime.now().replace(
                # '시간' 문자열이 "HHMMSS" 형식(예: "153010"은 15시 30분 10초를 의미)으로 제공된다고 가정합니다.
                hour=int(주문체결시간[:-4]),
                minute=int(주문체결시간[-4:-2]),
                second=int(주문체결시간[-2:]),
            )

            # 실 투자시 지정가 매도 주석 처리 TODO: 실 투자시 지정가 매도 주석처리
            정정주문가격 = self.stock_code_to_sell_price_dict.get(종목코드, None)
            if not 정정주문가격:
                logger.info(f"종목코드: {종목코드}, 최우선 매수 호가X 주문 실폐!!")
                return
            # basic.info.dict = self.stock_code_to_info_dict.get(종목코드, None)
            # if not basic.info.dict:
            #     logger.info(f"종목코드: {종목코드}, 기본정보X 정정주문 실폐!!")
            #     return
            # 정정주문가격 = basic_info_dict['하한가']
            if self.now_time - order_time >= datetime.timedelta(seconds=10):
            # if 주문구분 == "매도" and self.now_time - order_time >= datetime.timedelta(seconds=10):
                # 지정가 매도 주문이후 10초안에 미체결시 시장가 매도 정정 주문
                logger.info(f"종목코드: {종목코드}, 주문번호: {주문번호}, 지정가 매도 정정 주문!!")
                self.orders_queue.put(
                    [
                        "매도정정주문",
                        self._get_screen_num(),
                        self.account_num,
                        6,
                        종목코드,
                        미체결수량,
                        정정주문가격,
                        "00",
                        주문번호,
                    ]
                )

            # 실 투자시 시장가 매도 주석해제 TODO: 실 투자시 시장가 매도 주석해제
            # if self.now_time - order_time >= datetime.timedelta(seconds=10):
            #     # 시장가 매도 주문이후 10초안에 미체결시 시장가 매도 정정 주문
            #     logger.info(f"종목코드: {종목코드}, 주문번호: {주문번호}, 시장가 매도 정정 주문!!")
            #     self.orders_queue.put(
            #         [
            #             "매도정정주문",
            #             self._get_screen_num(),
            #             self.account_num,
            #             6,
            #             종목코드,
            #             미체결수량,
            #             "",
            #             "03",
            #             주문번호,
            #         ]
            #     )
                pop_list.append(주문번호)
        for order_num in pop_list:
            self.unfinished_order_num_to_info_dict.pop(order_num)
            self.save_settings()

    def load_settings(self):
        self.resize(self.settings.value("size", self.size()))
        self.move(self.settings.value("pos", self.pos()))
        self.buyAmountLineEdit.setText(self.settings.value("buyAmountLineEdit", defaultValue="100000", type=str))
        self.goalReturnLineEdit.setText(self.settings.value("goalReturnLineEdit", defaultValue="2.5", type=str))
        self.stopLossLineEdit.setText(self.settings.value("stopLossLineEdit", defaultValue="-2.5", type=str))

    def save_pickle(self):
        self.realtime_watchlist_df.to_pickle("./realtime_watchlist_df.pkl")
        self.realtime_watchlist_df.to_csv("./realtime_watchlist_df.csv")

    def request_current_order_info(self): # 미체결 처리
        self.tr_req_queue.put([self.get_current_order_info])

    def update_pandas_models(self): # registed_condition_df의 DataTable을 보여주는 함수
        pd_model = PandasModel(self.registed_condition_df) # 조건 검색식 목록 뷰
        self.registeredTableView.setModel(pd_model)
        pd_model2 = PandasModel(self.realtime_watchlist_df) # 실시간 조건 검색 편입 목록 뷰
        self.watchListTableView.setModel(pd_model2)
        self.account_info_df
        pd_model3 = PandasModel(self.account_info_df) # 실시간 계좌정보 목록 뷰
        self.accountTableView.setModel(pd_model3)

    def condition_in(self): # 조건 검색식 편입
        condition_name = self.conditionComboBox.currentText()
        condition_idx = self.condition_name_to_condition_idx_dict.get(condition_name, None)
        if not condition_idx:
            logger.info(f"잘못된 조건 검색식 이름! 다시 선택하세요!!")
            return
        else:
            logger.info(f"{condition_name}  실시간 조건 검색 등록 요청!!")
            self.send_condition(self._get_screen_num(), condition_name, condition_idx, 1)

    def condition_out(self): # 조건 검색식 편출
        condition_name = self.conditionComboBox.currentText()
        condition_idx = self.condition_name_to_condition_idx_dict.get(condition_name, None)
        if not condition_idx:
            logger.info(f"잘못된 조건 검색식 이름! 다시 선택하세요!!")
            return
        elif condition_idx in self.registed_condition_df.index:
            logger.info(f"{condition_name}  실시간 조건 검색 편출!!")
            self.registed_condition_df.drop(condition_idx, inplace=True)
        else:
            logger.info(f"조건식 편출 실패")
            return

    def _set_signal_slots(self): # 키움 API와 연동을 위한 Slot
        self.kiwoom.OnEventConnect.connect(self._event_connect)
        self.kiwoom.OnReceiveRealData.connect(self._receive_realdata)
        self.kiwoom.OnReceiveConditionVer.connect(self._receive_condition)
        self.kiwoom.OnReceiveRealCondition.connect(self._receive_real_condition)
        self.kiwoom.OnReceiveTrData.connect(self.receive_tr_data)
        self.kiwoom.OnReceiveChejanData.connect(self.receive_chejandata)
        self.kiwoom.OnReceiveMsg.connect(self.receive_msg)

    def receive_msg(self, sScrNo, sRQName, sTrCode, sMsg):
        logger.info(f"Received MSG: 화면번호: {sScrNo}, 사용자 구분명: {sRQName}, TR이름: {sTrCode}, 메세지: {sMsg}")

    def get_current_order_info(self):
        self.set_input_value("계좌번호", self.account_num)
        self.set_input_value("체결구분", "1")
        self.set_input_value("매매구분", "0")
        self.comm_rq_data("opw00075_req", "opw00075", 0, self._get_screen_num())

    def request_get_account_balance(self): #
        self.tr_req_queue.put([self.get_account_balance]) # 계좌정보를 10초에 한번 요청

    def send_tr_request(self): # TR요청 진행
        self.now_time = datetime.datetime.now()
        if self.is_check_tr_req_condition() and not self.tr_req_queue.empty():
            request_func, *func_args = self.tr_req_queue.get()
            logger.info(f"Excuting TR request function: {request_func}")
            request_func(*func_args) if func_args else request_func()
            self.last_tr_send_times.append(self.now_time)

    def get_account_info(self): # 계좌번호를 받아오는 함수
        account_nums = str(self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"]).rstrip(';'))
        logger.info(f"계좌번호 리스트: {account_nums}")
        self.account_num = account_nums.split(';')[0]
        logger.info(f"사용 계좌 번호: {self.account_num}")
        self.accountNumComboBox.addItems([x for x in account_nums.split(';') if x != '']) # 콤보 박스에 split해서 넣어준다

    def get_account_balance(self): # 계좌 정보 조회
        if self.is_check_tr_req_condition():
            self.set_input_value("계좌번호", self.accountNumComboBox.currentText())
            self.set_input_value("계좌번호", self.account_num)
            self.set_input_value("비밀번호", "")
            self.set_input_value("비밀번호입력매체구분", "00")
            # self.comm_rq_data("opw00018_req", "opw00018", 0, self._get_screen_num())
            self.comm_rq_data("opw00018_req", "opw00018", 0, self._get_screen_num())


    def receive_tr_data(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext, nDataLength, sErrorCode, sMessage,
                        sSplmMsg): # 체결 데이터
        # sScrNo: 화면번호, sRQName: 사용자 구분명, sTrCode: TR이름, sRecordName: 레코드 이름, sPrevNext: 연속조회 유무를 판단하는 값 0: 연속(추가조회)데이터 없음, 2:연속(추가조회) 데이터 있음
        # 조회요청 응답을 받거나 조회 데이터를 수신했을때 호출합니다.
        # 조회 데이터는 이 이벤트에서 GetCommData()함수를 이용해서 얻어올 수 있습니다.
        logger.info(
            f"Receive TR data sScrNo: {sScrNo}, sRQName: {sRQName}, "
            f"sTrCode: {sTrCode}, sRecordName: {sRecordName}, sPrevNext: {sPrevNext}, "
            f"nDataLength: {nDataLength}, sErrorCode: {sErrorCode}, sMessage: {sMessage}, sSplmMsg: {sSplmMsg}"
        )
        try:
            if sRQName == "opw00018_req":
                self.on_opw00018_req(sTrCode, sRQName)
            elif sRQName == "opt10075_req":
                self.on_opt10075_req(sTrCode, sRQName)
            elif sRQName == "opt10001_req":
                self.on_opt10001_req(sTrCode, sRQName)
        except Exception as e:
            logger.exception(e)


    def get_basic_stock_info(self, stock_code):
        if self.is_check_tr_req_condition():
            self.set_input_value("종목코드", stock_code)
            self.comm_rq_data(f"opt10001_req", "opt10001", 0, self._get_screen_num())

    def get_chejandata(self, nFid):
        ret = self.kiwoom.dynamicCall("GetChejanData(int)", nFid)
        return ret

    def receive_chejandata(self, sGubun, nItemCnt, sFIdList): #  실시간 체결 결과 요청 함수(체결 접수와 체결 결과)
        # sGubun: 체결구분 접수와 체결시 '0'값, 국내주식 잔고변경은 '1'값, 파생잔고변경은 '4'
        if sGubun == "0":
            종목코드 = self.get_chejandata(9001).replace("A", "").strip()
            종목명 = self.get_chejandata(302).strip()
            주문체결시간 = self.get_chejandata(908).strip()
            주문수량 = 0 if len(self.get_chejandata(900)) == 0 else int (self.get_chejandata(900))
            주문가격 = 0 if len(self.get_chejandata(901)) == 0 else int (self.get_chejandata(901))
            체결수량 = 0 if len(self.get_chejandata(911)) == 0 else int (self.get_chejandata(911))
            체결가격 = 0 if len(self.get_chejandata(910)) == 0 else int (self.get_chejandata(910))
            미체결수량 = 0 if len(self.get_chejandata(902)) == 0 else int (self.get_chejandata(902))
            주문구분 = self.get_chejandata(905).replace("+", "").replace("-", "").strip()
            매매구분 = self.get_chejandata(906).strip()
            단위체결가 = 0 if len(self.get_chejandata(914)) == 0 else int (self.get_chejandata(914))
            단위체결량 = 0 if len(self.get_chejandata(915)) == 0 else int (self.get_chejandata(915))
            원주문번호 = self.get_chejandata(904).strip()
            주문번호 = self.get_chejandata(9203).strip()
            logger.info(f" Receive chejandata! 주문체결시간: {주문체결시간}, 종목코드: {종목코드}, "
                        f"종목명: {종목명}, 주문수량: {주문수량}, 주문가격: {주문가격}, 체결수량: {체결수량}, 체결가격: {체결가격}, "
                        f"주문구분: {주문구분}, 미체결수량: {미체결수량}, 매매구분: {매매구분}, 단위체결가: {단위체결가}, "
                        f"단위체결량: {단위체결량}, 주문번호: {주문번호}, 원주문번호: {원주문번호}")
            if 주문구분 == "매수" and 체결수량 > 0:
                self.realtime_watchlist_df.loc[종목코드, "보유수량"] = 체결수량


            if 주문구분 in ("매도", "매도정정"): # 미체결 주문 처리
                self.unfinished_order_num_to_info_dict[주문번호] = dict(
                    종목코드=종목코드,
                    미체결수량=미체결수량,
                    주문가격=주문가격,
                    주문체결시간=주문체결시간,
                )
                if 미체결수량 == 0: # 미체결수량이 '0'이면 모두체결된 것이므로 unfinished_order_num_to_info_dict에서 제거
                    self.unfinished_order_num_to_info_dict.pop(주문번호)

        if sGubun == 1:
            logger.info("잔고통보")

    def _login(self):
        ret = self.kiwoom.dynamicCall("CommConnect()")
        if ret == 0:
            logger.info("로그인 창 열기 성공!!")

    def _event_connect(self, err_code):
        if err_code == 0:
            logger.info("로그인 성공!!")
            self._after_login()
        else:
            raise Exception("로그인 실폐!!")

    def _after_login(self): # 로그인이 끝나면 바로 실행되는 함수
        self.get_account_info()
        logger.info("조건 검색 정보 요청")
        self.kiwoom.dynamicCall("GetConditionLoad()") # 조건 검색 정보 요청

        self.timer1.start(300) # 0.3초마다 한번 실행
        self.timer2.start(10) # 0.01초마다 한번 실행
        self.timer3.start(10) # 0.01초마다 한번 실행
        self.timer4.start(5000) # 5초마다 한번 실행
        self.timer5.start(60000) # 60초마다 한번 실행
        self.timer6.start(30000) # 30초마다 한번 실행
        self.timer7.start(100) # 0.1초마다 한번 실행
        self.timer8.start(1000)  # 1초마다 한번 실행

    def _receive_condition(self): # 조건 검색식 받는 함수
        condition_info = self.kiwoom.dynamicCall("GetConditionNameList()").split(';')
        for condition_name_idx_str in condition_info:
            if len(condition_name_idx_str) == 0:
                continue
            condition_idx, condition_name = condition_name_idx_str.split('^')
            self.condition_name_to_condition_idx_dict[condition_name] = condition_idx
            # print(condition_idx, condition_name)
            # if condition_name == self.using_condition_name:
            #     self.send_condition(self._get_screen_num(), condition_name, condition_idx, 1)
        self.conditionComboBox.addItems(self.condition_name_to_condition_idx_dict.keys())

    def _get_screen_num(self): # 실시간 화면 번호를 만드는 함수
        self.scrnum += 1 # 화면 번호 요청시 1씩 더해지며 화면 번호를 만든다
        if self.scrnum > 5150: # 5150을 초과하면 5000으로 초기화
            self.scrnum = 5000
        return str(self.scrnum)

    def send_condition(self, scr_Num, condition_name, condition_idx, n_search): # 조건 검색식 등록
        # n_search : 조회구분 0:조건검색만, 1:조건검색+실시간 조건검색
        result = self.kiwoom.dynamicCall(
            "SendCondition(QString, QString, int, int)",
            scr_Num, condition_name, condition_idx, n_search
        )
        if result == 1:
            logger.info(f"{condition_name} 조건 검색 등록!!")
            self.registed_condition_df.loc[condition_idx] = {"화면번호": scr_Num, "조건식이름": condition_name}
            self.registed_conditions_list.append(condition_name)
        elif result != 1 and condition_name in self.registed_conditions_list:
            logger.info(f"{condition_name} 조건검색 이미 등록 완료!!")
            self.registed_condition_df.loc[condition_idx] = {"화면번호": scr_Num, "조건식이름": condition_name}
        else:
            logger.info(f"{condition_name} 조건 검색 등록 실패!!")

    def send_condition_stop(self, scr_Num, condition_name, condition_idx): # 조건 검색식 실시간 해제
        logger.info(f"{condition_name} 조건 검색 실시간 해제!!")
        # self.kiwoom.dynamicCall(
        #     "SendConditionStop(QString, QString, int)",
        #     scr_Num, condition_name, condition_idx
        # )

    def _receive_real_condition(self, strCode, strType, strConditionName, strConditionIndex): # 실시간 검색된 종묵을 편입 또는 이탈
        # strType: 이벤트 종류, "I":종목편입, "D":종목 이탈
        # strConditionName: 조건식 이름
        # strConditionIndex: 조건식 인덱스

        logger.info(f"Received real condition, {strCode}, {strType}, {strConditionName}, {strConditionIndex}")
        if strConditionIndex.zfill(3) not in self.registed_condition_df.index.to_list():
            logger.info(f"조건명: {strConditionName}, 편입 조건식에 해당 안됨 Pass")
            return

        if strType == "I" and strCode not in self.realtime_watchlist_df.index.to_list():
            if strCode not in self.realtime_reqisted_codes:
                self.register_code_to_realtime_list(strCode) # 실시간 체결 등록
            name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [strCode])  # 종목코드에 해당하는 종목명을 전달

            self.realtime_watchlist_df.loc[strCode] = {
                '종목명': name,
                '현재가': None,
                '평균단가': None,
                '목표가': None,
                '손절가': None,
                '수익률': None,
                '매수기반조건식': strConditionName,
                '보유수량': 0,
                '매수주문완료여부': False,
            }
            self.tr_req_queue.put([self.get_basic_stock_info, strCode])
            # TODO:매수 주문 진행


        # logger.info(f"Received real condition, {strCode}, {strType}, {strConditionName}, {strConditionIndex}")
        # if strConditionIndex.zfill(3) not in self.registered_condition_df.index.to_list():
        #     logger.info(f"조건명: {strConditionName}, 편입 조건식에 해당 안됨 Pass")
        #     return
        # if strType == "I" and strCode not in self.realtime_watchlist_df.Index.to_list():
        #     self.register_code_to_realtime_list(strCode)  # 실시간 체결 등록
        # name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [strCode])  # 종목코드에 해당하는 종목명을 전달
        #
        # self.realtime_watchlist_df.loc[strCode] = {
        #     '종목명': name,
        #     '현재가': None,
        #     '평균단가': None,
        #     '목표가': None,
        #     '손절가': None,
        #     '수익률': None,
        #     '매수기반조건식': strConditionName,
        #     '보유수량': 0,
        #     '매수주문완료여부': False,
        # }
        # self.tr_req_queue.put([self.get_basic_stock_info, strCode])

    def get_comm_realdata(self, strCode, nFid):
        # 실시간시세 데이터 수신 이벤트인 OnReceiveRealData() 가 발생될때 실시간데이터를 얻어오는 함수입니다.
        return self.kiwoom.dynamicCall("GetCommRealData(QString, int)", strCode, nFid)

    def _receive_realdata(self, sJongmokCode, sRealType, sRealData): # 실시간으로 주식 체결을 체크하는 함수
        if sRealType == "주식체결":
            self.now_time = datetime.datetime.now()
            now_price = int(self.get_comm_realdata(sRealType, 10).replace('-', '')) # 현재가
            최우선매수호가 = int(self.get_comm_realdata(sRealType, 28).replace('-', '')) # 최우선 매수 호가
            self.stock_code_to_sell_price_dict[sJongmokCode] = 최우선매수호가

            if sJongmokCode in self.realtime_watchlist_df.index.to_list():
                if not self.realtime_watchlist_df.loc[sJongmokCode, "매수주문완료여부"]:
                    goal_price = now_price * (1 + float(self.goalReturnLineEdit.text()) / 100)
                    stoploss_price = now_price * (1 + float(self.stopLossLineEdit.text()) / 100)
                    self.realtime_watchlist_df.loc[sJongmokCode, "목표가"] = goal_price
                    self.realtime_watchlist_df.loc[sJongmokCode, "손절가"] = stoploss_price
                    order_amount = int(self.buyAmountLineEdit.text()) // now_price

                    # if self.current_available_buy_amount_krw < int(self.buyAmountLineEdit.text()):
                    #     logger.info(f"주문 가능 금액: {self.current_available_buy_amount_krw: ,}원: 금액 부족으로 매수 X")
                    #     return
                    # order_amount = int(self, buyAmountLineEdit.text()) // now_price

                    if order_amount < 1:
                        logger.info(f"종목코드: {sJongmokCode}, 주문 수량 부족으로 매수 진행 안됨!!")
                        return
                    self.orders_queue.put(
                        [
                            "시장가매수주문",
                            self._get_screen_num(),
                            self.accountNumComboBox.currentText(),
                            1,
                            sJongmokCode,
                            order_amount,
                            "",
                            "03",
                            "",
                        ],
                    )
                    self.realtime_watchlist_df.loc[sJongmokCode, "매수주문완료여부"] = True
                self.realtime_watchlist_df.loc[sJongmokCode, "현재가"] = now_price
                mean_buy_price = self.realtime_watchlist_df.loc[sJongmokCode, "평균단가"]
                if mean_buy_price is not None:
                    self.realtime_watchlist_df.loc[sJongmokCode, "수익률"] = round(
                        (now_price - mean_buy_price) / mean_buy_price * 100 - 0.21,
                        2,
                    )

                보유수량 = int(copy.deepcopy(self.realtime_watchlist_df.loc[sJongmokCode, "보유수량"]))
                if 보유수량 > 0 and now_price < self.realtime_watchlist_df.loc[sJongmokCode, '손절가']:
                    logger.info(f"종목코드: {sJongmokCode} 매도 진행!! (손절)")
                    # basic_info_dict = self.stock_code_to_info_dict.get(sJongmokCode, None)
                    # if not basic_info_dict:
                    #     logger.info(f"종목코드: {sJongmokCode}, 기본정보X 정정주문 실폐!!")
                    #     return
                    # 주문가격 = basic_info_dict['하한가']

                    주문가격 = self.stock_code_to_sell_price_dict.get(sJongmokCode, None)
                    if not 주문가격:
                        logger.info(f"종목코드: {sJongmokCode}, 최우선 매수 호가X 주문 실폐!!")
                        return

                    self.orders_queue.put(
                        [
                            "매도주문",
                            self._get_screen_num(),
                            self.account_num,
                            2,
                            sJongmokCode,
                            보유수량,
                            주문가격,
                            "00",
                            "",
                        ]
                    )

                    # 실투자시 시장가 매도 주석해제
                    # logger.info(f"종목코드: {sJongmokCode} 시장가 매도 진행!!")
                    # self.orders_queue.put(
                    #     [
                    #         "시장가매도주문",
                    #         self._get_screen_num(),
                    #         self.account_num,
                    #         2,
                    #         sJongmokCode,
                    #         self.realtime_watchlist_df.loc[sJongmokCode, "보유수량"],
                    #         "",
                    #         "03",
                    #         "",
                    #     ],
                    # )
                    # self.registed_condition_df.drop(sJongmokCode, inplace=True) #체결 완료시 drop으로 registed_condition_df에서 삭제
                    # registed_condition_df에서 sJongmokCode가 존재하는지 확인 후 삭제
                    if sJongmokCode in self.registed_condition_df.index:
                        self.registed_condition_df.drop(sJongmokCode, inplace=True)

                elif 보유수량 > 0 and now_price > self.realtime_watchlist_df.loc[sJongmokCode, "목표가"]:
                    logger.info(f"종목코드: {sJongmokCode} 매도 진행(익절 )!!")

                    self.orders_queue.put(
                        [
                            "지정가매도주문",
                            self._get_screen_num(),
                            self.account_num,
                            2,
                            sJongmokCode,
                            보유수량,
                            now_price,
                            "00",
                            "",
                        ],
                    )
                    # self.registed_condition_df.drop(sJongmokCode, inplace=True) #체결 완료시 drop으로 registed_condition_df에서 삭제
                    # registed_condition_df에서 sJongmokCode가 존재하는지 확인 후 삭제
                    if sJongmokCode in self.registed_condition_df.index:
                        self.registed_condition_df.drop(sJongmokCode, inplace=True)
                    else:
                        logger.info(f"종목코드: {sJongmokCode}는 registed_condition_df에 존재하지 않음. 삭제 스킵.")



            # if sJongmokCode in self.realtime_watchlist_df.index.to_list():
            #     if not self.realtime_watchlist_df.loc[sJongmokCode, "매수주문완료여부"]:
            #         goal_price = now_price * (1 + float(self.goalReturnLineEdit.text()) / 100)
            #         stoploss_price = now_price * (1 + float(self.stopLossLineEdit.text()) / 100)
            #         self.realtime_watchlist_df.loc[sJongmokCode, "목표가"] = goal_price
            #         self.realtime_watchlist_df.loc[sJongmokCode, "손절가"] = stoploss_price
            #         if self.current_available_buy_amount_krw < int(self.buyAmountLineEdit.text()):
            #             logger.info(f"주문 가능 금액: {self.current_available_buy_amount_krw: ,}원: 금액 부족으로 매수 X")
            #             return
            #         order_amount = int(self, buyAmountLineEdit.text()) // now_price
            #         if order_amount < 1:
            #             logger.info(f"종목코드: {sJongmokCode}, 주문수량 부족으로 매수 진행 X")
            #             return
            #         self.orders_queue.put(
            #             [
            #                 "시장가매수주문",
            #                 self._get_screen_num(),
            #                 self.accountNumComboBox.currentText(),
            #                 1,
            #                 sJongmokCode,
            #                 order_amount,
            #                 "",
            #                 "03",
            #                 "",
            #                 ],
            #         )
            #         self.realtime_watchlist_df.loc[sJongmokCode, "매수주문완료여부"] = True
            #
            #     self.realtime_watchlist_df.loc[sJongmokCode, '현재가'] = now_price
            #     mean_buy_price = self.realtime_watchlist_df.loc[sJongmokCode, '평균단가']
            #     if mean_buy_price is not None:
            #         self.realtime_watchlist_df.loc[sJongmokCode, '수익률'] = round(
            #             (now_price - mean_buy_price) / mean_buy_price * 100 - 0.21,
            #             2,
            #         )
            #     보유수량 = int(copy.deepcopy(self.realtime_watchlist_df.loc[sJongmokCode, '보유수량']))
            #     if 보유수량 > 0 and now_price < self.realtime_watchlist_df.loc[sJongmokCode, '손절가']:
            #         logger.info(f"종목코드: {sJongmokCode} 매도 진행!! (손절)")
            #         basic_info_dict = self.stock_code_to_info_dict.get(sJongmokCode, None)
            #         if not basic_info_dict:
            #             logger.info(f"종목코드: {sJongmokCode}, 기본정보X 정정주문 실폐!!")
            #             return
            #         주문가격 = basic_info_dict['하한가']
            #         self.orders_queue.put(
            #
            #         )

    def send_orders(self): # 주문을 보내는 함수
        self.now_time = datetime.datetime.now()

        if self.is_check_tr_req_condition() and not self.orders_queue.empty():
            sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo = self.orders_queue.get()
            ret = self.send_order(sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo)
            if ret == 0:
                logger.info(f"{sRQName} 주문 접수 성공!!")
            self.last_tr_send_times.append(self.now_time)

    def send_order(self, sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo):
        return self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                       [sRQName, sScreenNo, sAccNo, nOrderType, sCode, nQty, nPrice, sHogaGb, sOrgOrderNo])
        # [SendOrder() 함수]
        #
        # sRQName: 사용자 구분명 (OnReceiveTrData에서 받을 이름으로!
        # sScreenNo: 화면번호,  sAccNo: 계좌번호 10자리
        # nOrderType, 주문유형
        # 1: 신규매수, 2: 신규매도, 3: 매수취소, 4: 매도취소, 5: 매수정정, 6: 매도정정, 7: 프로그램매매 매수, 8: 프로그매매 매도
        # sCode: 종목코드(6자리), nQty: 주문수량, nPrice: 주문가격, sHogaGb: 거래구분(혹은 호가구분)은 아래 참고
        # sOrgOrderNo: 원주문번호.신규주문에는 공백 입력, 정정 / 취소시 입력합니다.
        # [거래구분]
        # 00 : 지정가, 03 : 시장가, 05 : 조건부지정가, 06 : 최유리지정가, 07 : 최우선지정가,, 10 : 지정가IOC, 13 : 시장가IOC
        # 16 : 최유리IOC, 20 : 지정가FOK, 23 : 시장가FOK, 26 : 최유리FOK, 61 : 장전시간외종가, 62 : 시간외단일가매매, 81 : 장후시간외종가
        # ※ 모의투자에서는 지정가 주문과 시장가 주문만 가능합니다.
        # 예시 -> SendOrder("주식주문", _get_screen_num(), cbo계좌.Text.Trim(), 1, 종목코드, 수량, 현재가, "00", ""):

    def set_real(self, scrNum, strCodeList, strFidList, strRealType):
        self.kiwoom.dynamicCall("SetRealReg(QString, QString, QString, QString)", scrNum, strCodeList, strFidList, strRealType)

    def register_code_to_realtime_list(self, code):
        fid_list = "10;12;20;28"
        # "10": "현재가", "12": "등락율", "20": "체결시간", "28": "(최우선)매수호가"
        if len(code) != 0:
            self.realtime_reqisted_codes.add(code)
            self.set_real(self._get_screen_num(), code, fid_list, "I")
            logger.info(f"{code}, 실시간 등록 완료!!")

    def is_check_tr_req_condition(self): # TR요청시 제한되는 부분을 감시하는 함수
        now_time = datetime.datetime.now()
        if len(self.last_tr_send_times) >= self.max_send_per_sec and \
            now_time - self.last_tr_send_times[-self.max_send_per_sec] < datetime.timedelta(milliseconds=1000):
            logger.info(f"초 단위 TR 요청 제한! Wait for time to send!!")
            return  False
        elif len(self.last_tr_send_times) >= self.max_send_per_minute and \
            self.now_time - self.last_tr_send_times[-self.max_send_per_minute] < datetime.timedelta(minutes=1):
            logger.info(f"분 단위 TR 요청 제한! Wait for time to send!!")
            return False
        elif len(self.last_tr_send_times) >= self.max_send_per_hour and \
            now_time - self.last_tr_send_times[-self.max_send_per_hour] < datetime.timedelta(minutes=60):
            logger.info(f"시간 단위 TR 요청 제한! Wait for time to send!!")
            return False
        else:
            return True

    def get_comm_data(self, strTrCode, strRecordName, nIdex, strItemName):
        ret = self.kiwoom.dynamicCall("GetCommData(QString, QString, int, QString)", strTrCode, strRecordName, nIdex,
                                      strItemName)
        return ret.strip()

    def set_input_value(self, id, value):
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", id, value)

    def comm_rq_data(self, rqname, trcode, next, screen_no):
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", rqname, trcode, next, screen_no)

    def save_settings(self):
        # Write window size and position to config file
        self.settings.setValue("size", self.size())
        self.settings.setValue("pos", self.pos())
        self.settings.setValue('buyAmountLineEdit', self.buyAmountLineEdit.text())
        self.settings.setValue('goalReturnLineEdit', self.goalReturnLineEdit.text())
        self.settings.setValue('stopLossLineEdit', self.stopLossLineEdit.text())
        self.realtime_watchlist_df.to_pickle("./realtime_watchlist_df.pkl")

    def _get_repeat_cnt(self, trcode, rqname):
        ret = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", trcode, rqname)
        return ret

    def on_opt10001_req(self, sTrCode, sRQName):
        종목코드 = self.get_comm_data(sTrCode, sRQName, 0, "종목코드").replace("A", "").strip()
        상한가 = abs(int(self.get_comm_data(sTrCode, sRQName, 0, "상한가")))
        하한가 = abs(int(self.get_comm_data(sTrCode, sRQName, 0, "하한가")))
        self.stock_code_to_info_dict[종목코드] = dict(상한가=상한가, 하한가=하한가)

    def on_opt10075_req(self, sTrCode, sRQName):
        cnt = self._get_repeat_cnt(sTrCode, sRQName)
        for i in range(cnt):
            주문번호 = self.get_comm_data(sTrCode, sRQName, i, "주문번호").strip()
            미체결수량 = int(self.get_comm_data(sTrCode, sRQName, i, "미체결수량"))
            주문가격 = int(self.get_comm_data(sTrCode, sRQName, i, "주문가격"))
            종목코드 = self.get_comm_data(sTrCode, sRQName, i, "종목코드").strip()
            주문구분 = self.get_comm_data(sTrCode, sRQName, i, "주문구분").replace("+", "").replace("-", "").strip()
            시간 = self.get_comm_data(sTrCode, sRQName, i, "시간").strip()
            order_time = datetime.datetime.now().replace(
                # '시간' 문자열이 "HHMMSS" 형식(예: "153010"은 15시 30분 10초를 의미)으로 제공된다고 가정합니다.
                hour=int(시간[:-4]),
                minute=int(시간[-4:-2]),
                second=int(시간[-2:]),
            )

            정정주문가격 = self.stock_code_to_sell_price_dict.get(종목코드, None)
            if not 정정주문가격:
                logger.info(f"종목코드: {종목코드}, 최우선 매수 호가X 주문 실폐!!")
                return
            # basic.info.dict = self.stock_code_to_info_dict.get(종목코드, None)
            # if not basic.info.dict:
            #     logger.info(f"종목코드: {종목코드}, 기본정보X 정정주문 실폐!!")
            #     return
            # 정정주문가격 = basic_info_dict['하한가']
            if 주문구분 in ("매도", "매도정정") and self.now_time - order_time >= datetime.timedelta(seconds=10):
            # if 주문구분 == "매도" and self.now_time - order_time >= datetime.timedelta(seconds=10):
                # 지정가 매도 주문이후 10초안에 미체결시 시장가 매도 정정 주문
                logger.info(f"종목코드: {종목코드}, 주문번호: {주문번호}, 시장가 매도 정정 주문!!")
                self.orders_queue.put(
                    [
                        "매도정정주문",
                        self._get_screen_num(),
                        self.account_num,
                        6,
                        종목코드,
                        미체결수량,
                        정정주문가격,
                        "00",
                        주문번호,
                    ]
                )

            # basic.info.dict = self.stock_code_to_info_dict.get(종목코드, None)
            # if not basic.info.dict:
            #     logger.info(f"종목코드: {종목코드}, 기본정보X 정정주문 실폐!!")
            #     return
            # 정정주문가격 = basic_info_dict['하한가']
            # if 주문구분 in ("매도", "매도정정") and self.now_time - order_time >= datetime.timedelta(seconds=10):
            #     logger.info(f"종목코드: {종목코드}, 주문번호: {주문번호}, 미체결수량: {미체결수량}, 시장가 매도 정정 주문!!")

    def on_opw00018_req(self, sTrCode, sRQName):
        현재평가잔고 = int(self.get_comm_data(sTrCode, sRQName, 0, "추정예탁자산"))
        logger.info(f"현재평가잔고 : {현재평가잔고: ,}원")
        self.currentBalanceLabel.setText(f"현재 평가 잔고: {현재평가잔고: ,}원")
        cnt = self._get_repeat_cnt(sTrCode, sRQName)
        self.account_info_df = pd.DataFrame(
            columns=[
                "종목명",
                "매매가능수량",
                "보유수량",
                "매입가",
                "현재가",
                "수익률",
            ]
        )

        current_account_code_list = []
        current_filled_amount_krw = 0
        for i in range(cnt):
            종목코드 = self.get_comm_data(sTrCode, sRQName, i, "종목번호").replace("A", "").strip()
            current_account_code_list.append(종목코드)
            종목명 = self.get_comm_data(sTrCode, sRQName, i, "종목명")
            매매가능수량 = int(self.get_comm_data(sTrCode, sRQName, i, "매매가능수량"))
            보유수량 = int(self.get_comm_data(sTrCode, sRQName, i, "보유수량"))
            현재가 = int(self.get_comm_data(sTrCode, sRQName, i, "현재가"))
            매입가 = int(self.get_comm_data(sTrCode, sRQName, i, "매입가"))
            # 수익률 = int(self.get_comm_data(sTrCode, sRQName, i, "수익률(%)"))
            수익률 = int(float(self.get_comm_data(sTrCode, sRQName, i, "수익률(%)")))
            logger.info(
                f"종목코드: {종목코드}, 종목명: {종목명}, 매매가능수량: {매매가능수량}, 보유수량: {보유수량}, 매입가: {매입가}, 수익률: {수익률}"
            )
            current_filled_amount_krw += 보유수량 * 현재가
            if 종목코드 in self.realtime_watchlist_df.index.to_list():
                self.realtime_watchlist_df.loc[종목코드, "종목명"] = 종목명
                self.realtime_watchlist_df.loc[종목코드, "평균단가"] = 매입가
                self.realtime_watchlist_df.loc[종목코드, "보유수량"] = 보유수량
            self.account_info_df.loc[종목코드] = {
                "종목명": 종목명,
                "매매가능수량": 매매가능수량,
                "현재가": 현재가,
                "보유수량": 보유수량,
                "매입가": 매입가,
                "수익률": 수익률,
            }
        self.current_available_buy_amount_krw = 현재평가잔고 - current_filled_amount_krw
        if not self.is_updated_realtime_watchlist:
            for 종목코드 in current_account_code_list:
                self.register_code_to_realtime_list(종목코드)
            self.is_updated_realtime_watchlist = True
            realtime_tracking_code_list = self.realtime_watchlist_df.index.to_list()
            for stock_code in realtime_tracking_code_list:
                if stock_code not in current_account_code_list:
                    self.realtime_watchlist_df.drop(stock_code, inplace=True)
                    logger.info(f"종목코드: {stock_code} self.realtime_watchlist_df 에서 drop!!")

    @ staticmethod
    def get_sell_price(now_price):
        now_price = int(now_price)
        if now_price < 2000:
            return now_price
        elif 5000 > now_price >= 2000:
            return now_price - now_price % 5
        elif now_price >= 5000 and now_price < 20000:
            return now_price - now_price % 10
        elif now_price >= 20000 and now_price < 50000:
            return now_price - now_price % 50
        elif now_price >= 50000 and now_price < 200000:
            return now_price - now_price % 100
        elif now_price >= 200000 and now_price < 500000:
            return now_price - now_price % 500
        else:
            return now_price - now_price % 1000


#PyQt 디버깅용 코드
sys._excepthook = sys.excepthook

def my_exception_hook(exctype, value, traceback):
    # print the error and traceback
    print(exctype, value, traceback)
    # Call the normal Exception hook after
    sys._excepthook(exctype, value, traceback)
    sys.exit(1)

# Set the exception hook to our wrapping function
sys.excepthook = my_exception_hook


if __name__ == '__main__':
    app = QApplication(sys.argv)
    kiwoom_api = KiwoomAPI()
    sys.exit(app.exec_())