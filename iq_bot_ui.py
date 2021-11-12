from PyQt5 import QtCore, QtGui, QtWidgets

from PyQt5.QtGui import QCursor
from datetime import datetime, timedelta
import dateutil.parser as dparser

import os
import sys

from iqoptionapi.stable_api import IQ_Option
import pathlib
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from threading import Thread
import time

class Ui_IqOptionBot(object):
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    stop_thread = False
    login = False
    SAMPLE_SPREADSHEET_ID_input = ''
    bot_timeinterval = 6
    sheet_header_offset = 2
    trade_Sheet_data = {}
    current_active_options = {}
    balance_information = {}
    account_type = "PRACTICE"

    def setupUi(self, IqOptionBot):
        IqOptionBot.setObjectName("IqOptionBot")
        IqOptionBot.resize(458, 489)
        IqOptionBot.setAutoFillBackground(False)
        IqOptionBot.setStyleSheet("background-color:rgb(42, 42, 42)")
        self.centralwidget = QtWidgets.QWidget(IqOptionBot)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(40, 40, 381, 121))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.groupBox.setFont(font)
        self.groupBox.setAutoFillBackground(False)
        self.groupBox.setStyleSheet("color: rgb(255, 198, 109);\n"
"background-color:rgb(42, 42, 42);\n"
"")
        self.groupBox.setObjectName("groupBox")
        self.lineEdit = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit.setGeometry(QtCore.QRect(70, 30, 141, 31))
        self.lineEdit.setStyleSheet("color:rgb(255, 255, 255)")
        self.lineEdit.setObjectName("lineEdit")
        self.label = QtWidgets.QLabel(self.groupBox)
        self.label.setGeometry(QtCore.QRect(10, 40, 47, 16))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label.setFont(font)
        self.label.setAutoFillBackground(False)
        self.label.setStyleSheet("color:rgb(85, 255, 0)")
        self.label.setObjectName("label")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit_2.setGeometry(QtCore.QRect(70, 70, 141, 31))
        self.lineEdit_2.setStyleSheet("color:rgb(255, 255, 255)")
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.label_2 = QtWidgets.QLabel(self.groupBox)
        self.label_2.setGeometry(QtCore.QRect(10, 80, 47, 16))
        font = QtGui.QFont()
        font.setPointSize(8)
        self.label_2.setFont(font)
        self.label_2.setAutoFillBackground(False)
        self.label_2.setStyleSheet("color:rgb(0, 255, 0)")
        self.label_2.setObjectName("label_2")
        self.pushButton = QtWidgets.QPushButton(self.groupBox)
        self.pushButton.setGeometry(QtCore.QRect(260, 30, 75, 23))
        self.pushButton.setStyleSheet("QPushButton#pushButton {\n     background-color:rgb(85, 170, 0);\n     border-radius:5;\n     color:rgb(255, 255, 255);\n}\n\n\nQPushButton#pushButton:hover {\n     background-color:rgb(85, 170, 200);\n     cusor:pointer\n}\n")
        self.pushButton.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
        self.pushButton.setObjectName("pushButton")
        self.label_3 = QtWidgets.QLabel(self.groupBox)
        self.label_3.setGeometry(QtCore.QRect(260, 80, 47, 13))
        self.label_3.setObjectName("label_3")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit_3.setGeometry(QtCore.QRect(300, 70, 51, 31))
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_2.setGeometry(QtCore.QRect(40, 280, 381, 101))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.groupBox_2.setFont(font)
        self.groupBox_2.setStyleSheet("color:rgb(255, 198, 109)")
        self.groupBox_2.setObjectName("groupBox_2")
        self.label_4 = QtWidgets.QLabel(self.groupBox_2)
        self.label_4.setGeometry(QtCore.QRect(20, 30, 131, 16))
        self.label_4.setStyleSheet("color:rgb(85, 255, 0)")
        self.label_4.setObjectName("label_4")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.groupBox_2)
        self.lineEdit_4.setGeometry(QtCore.QRect(20, 50, 71, 31))
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.label_5 = QtWidgets.QLabel(self.groupBox_2)
        self.label_5.setGeometry(QtCore.QRect(240, 30, 47, 13))
        self.label_5.setObjectName("label_5")
        self.lineEdit_5 = QtWidgets.QLineEdit(self.groupBox_2)
        self.lineEdit_5.setGeometry(QtCore.QRect(180, 50, 181, 31))
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(120, 400, 75, 23))
        self.pushButton_2.setStyleSheet("background-color:rgb(85, 170, 0, 0.5);\n"
"border-radius:5;\n"
"color:rgb(255, 255, 255);")
        self.pushButton_2.setCursor(QCursor(QtCore.Qt.ForbiddenCursor))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.setEnabled(False)
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(250, 400, 75, 23))
        self.pushButton_3.setStyleSheet("background-color:rgb(217, 0, 3, 0.5);\n"
"border-radius:5;\n"
"color:rgb(255, 255, 255);")
        self.pushButton_3.setCursor(QCursor(QtCore.Qt.ForbiddenCursor))
        self.pushButton_3.setObjectName("pushButton_3")
        self.groupBox_3 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_3.setGeometry(QtCore.QRect(40, 180, 381, 81))
        self.pushButton_3.setEnabled(False)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.groupBox_3.setFont(font)
        self.groupBox_3.setStyleSheet("color:rgb(255, 198, 109)")
        self.groupBox_3.setObjectName("groupBox_3")
        self.lineEdit_6 = QtWidgets.QLineEdit(self.groupBox_3)
        self.lineEdit_6.setGeometry(QtCore.QRect(20, 30, 341, 31))
        self.lineEdit_6.setObjectName("lineEdit_6")
        IqOptionBot.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(IqOptionBot)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 458, 21))
        self.menubar.setObjectName("menubar")
        IqOptionBot.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(IqOptionBot)
        self.statusbar.setObjectName("statusbar")
        IqOptionBot.setStatusBar(self.statusbar)
        self.pushButton.clicked.connect(self.login)
        self.pushButton_2.clicked.connect(self.run)
        self.pushButton_3.clicked.connect(self.stop_allthreading)

        self.retranslateUi(IqOptionBot)
        QtCore.QMetaObject.connectSlotsByName(IqOptionBot)

    def retranslateUi(self, IqOptionBot):
        _translate = QtCore.QCoreApplication.translate
        IqOptionBot.setWindowTitle(_translate("IqOptionBot", "MainWindow"))
        self.groupBox.setTitle(_translate("IqOptionBot", "Login Widget"))
        self.label.setText(_translate("IqOptionBot", "User Emalil"))
        self.label_2.setText(_translate("IqOptionBot", "Password"))
        self.pushButton.setText(_translate("IqOptionBot", "LOGIN"))
        self.label_3.setText(_translate("IqOptionBot", "Status:"))
        self.groupBox_2.setTitle(_translate("IqOptionBot", "IQ Bot Setting"))
        self.label_4.setText(_translate("IqOptionBot", "Bot Check Time Interval(S)"))
        self.label_5.setText(_translate("IqOptionBot", "Status:"))
        self.pushButton_2.setText(_translate("IqOptionBot", "Run"))
        self.pushButton_3.setText(_translate("IqOptionBot", "Stop"))
        self.groupBox_3.setTitle(_translate("IqOptionBot", "Google Spread Sheet Key"))

    def login(self):
        self.lineEdit_3.setText('Logging');
        self.lineEdit_3.repaint()
        self.pushButton.setEnabled(False)
        username = self.lineEdit.text().strip()
        password = self.lineEdit_2.text().strip()
        self.API_connection = IQ_Option('apurbo.hello88@gmail.com', 'jjhpph_2021922st@art')
        iqch1, iqch2 = self.API_connection.connect()  # connect to iqoption
        balance_type = "PRACTICE"
        self.API_connection.change_balance(balance_type)
        if iqch1 == True:
            self.lineEdit_3.setText('Success')
            self.lineEdit_3.setStyleSheet("color:rgb(85, 170, 0);")
            self.pushButton_2.setEnabled(True)
            self.pushButton_2.setStyleSheet("QPushButton#pushButton_2 {\n     background-color:rgb(85, 170, 0);\n     border-radius:5;\n     color:rgb(255, 255, 255);\n}\n\n\nQPushButton#pushButton_2:hover {\n     background-color:rgb(85, 170, 200);\n     cusor:pointer\n}\n")
            self.pushButton_2.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
            self.pushButton_3.setEnabled(True)
            self.pushButton_3.setStyleSheet("QPushButton#pushButton_3 {\n     background-color:rgb(217, 0, 3);\n     border-radius:5;\n     color:rgb(255, 255, 255);\n}\n\n\nQPushButton#pushButton_3:hover {\n     background-color:rgb(85, 170, 200);\n     cusor:pointer\n}\n")
            self.pushButton_3.setCursor(QCursor(QtCore.Qt.PointingHandCursor))
            self.pushButton.setEnabled(True)
            self.login = True

        else:
            self.lineEdit_3.setText('Fail')
            self.lineEdit_3.setStyleSheet("color:rgb(255, 0, 0);")
            self.pushButton_2.setEnabled(False)
            self.pushButton_2.setStyleSheet("background-color:rgb(85, 170, 0, 0.5);\n"
                                            "border-radius:5;\n"
                                            "color:rgb(255, 255, 255);")
            self.pushButton_3.setStyleSheet("background-color:rgb(217, 0, 3, 0.5);\n"
                                            "border-radius:5;\n"
                                            "color:rgb(255, 255, 255);")
            self.pushButton_3.setEnabled(False)
            self.pushButton.setEnabled(True)

        return


    def run(self):
        if not self.login:
            return
        # self.SAMPLE_SPREADSHEET_ID_input = self.lineEdit_6.text().strip()
        self.SAMPLE_SPREADSHEET_ID_input = '1YTQPSIGq0Wt5_5_Dan6bghDmTYl2u_zGwzA7IJH79EI'
        bot_timeinterval = self.lineEdit_4.text().strip()
        if not self.SAMPLE_SPREADSHEET_ID_input or not bot_timeinterval or not bot_timeinterval.isnumeric():
            self.lineEdit_5.setText('Error')
            self.lineEdit_5.setStyleSheet("color:rgb(255, 0, 0);")
        else:
            if self.bot_timeinterval > 6:
                self.bot_timeinterval = int(bot_timeinterval)
            self.lineEdit_5.setText('Working')
            self.lineEdit_5.setStyleSheet("color:rgb(85, 170, 0);")

            self.stop_thread = False
            trading_thread = Thread(target=self.trading_run)
            trading_thread.start()
            # balance_market_thread = Thread(target=self.balance_market_run)
            # balance_market_thread.start()

    def trading_run(self):
        if self.API_connection.check_connect() == False:
            check, reason = self.API_connection.connect()
            if not check:
                self.throw_error_exception('Can\'t log in IQOption!')

        while not self.stop_thread:
            values = []
            UpdateValues = []
            try:
                self.authenticate_google()
                result = self.service.spreadsheets().values().get(spreadsheetId=self.SAMPLE_SPREADSHEET_ID_input,range='TRADE2!A1:T').execute()
                values = result.get('values', [])
            except HttpError as err:
                self.throw_error_exception('Error: Connecting to Google Sheet!')
            if values and len(values) > self.sheet_header_offset:
                for row in values[self.sheet_header_offset:]:
                    try:

                        if row[1]:
                            if row[1] == 'Open':
                                if row[2] and dparser.parse(row[2]) > datetime.now():
                                    continue
                                # Get the instrument_type
                                instrument_id = "BTCUSD"
                                instrument_type = "crypto"
                                # Get the side/direction
                                side = "buy"
                                if row[5] == 'BUY' or row[5] == 'buy' or row[5] == 'Buy':
                                    side = "buy"
                                elif row[5] == 'SELL' or row[5] == 'sell' or row[5] == 'Sell':
                                    side = 'sell'
                                # Get the Invest amount
                                amount = 0.01  # input how many Amount you want to play

                                # "leverage"="Multiplier"
                                leverage = 3  # you can get more information in get_available_leverages()

                                type = "market"  # input:"market"/"limit"/"stop"

                                # for type="limit"/"stop"

                                # only working by set type="limit"
                                limit_price = None  # input:None/value(float/int)

                                # only working by set type="stop"
                                stop_price = None  # input:None/value(float/int)

                                # "percent"=Profit Percentage
                                # "price"=Asset Price
                                # "diff"=Profit in Money
                                if row[9]:
                                    take_profit_kind = "percent"  # input:None/"price"/"diff"/"percent"
                                    take_profit_value = float(row[9].replace('%', ''))  # input:None/value(float/int)
                                else:
                                    take_profit_kind = None
                                    take_profit_value = None
                                if row[10]:
                                    stop_lose_kind = "percent"  # input:None/"price"/"diff"/"percent"
                                    stop_lose_value = float(row[10].replace('%', ''))  # input:None/value(float/int)
                                else:
                                    stop_lose_kind = None
                                    stop_lose_value = None
                                # "use_trail_stop"="Trailing Stop"
                                use_trail_stop = True  # True/False

                                # "auto_margin_call"="Use Balance to Keep Position Open"
                                auto_margin_call = False  # True/False
                                use_token_for_commission = False  # True/False

                                check, order_id = self.API_connection.buy_order(instrument_type=instrument_type,
                                                                                instrument_id=instrument_id,
                                                                                side=side, amount=amount,
                                                                                leverage=leverage,
                                                                                type=type, limit_price=limit_price,
                                                                                stop_price=stop_price,
                                                                                stop_lose_value=stop_lose_value,
                                                                                stop_lose_kind=stop_lose_kind,
                                                                                take_profit_value=take_profit_value,
                                                                                take_profit_kind=take_profit_kind,
                                                                                use_trail_stop=use_trail_stop,
                                                                                auto_margin_call=auto_margin_call,
                                                                                use_token_for_commission=use_token_for_commission)
                                while self.API_connection.get_async_order(order_id) == None:
                                    pass
                                order_data = self.API_connection.get_async_order(order_id)
                                print(order_data)
                                # if check:
                                #     print(self.API_connection.get_order(order_id))
                                #     print(self.API_connection.get_positions("crypto"))
                                #     print(self.API_connection.get_position_history("crypto"))
                                #     print(self.API_connection.get_available_leverages("crypto", "BTCUSD"))
                                #     print(self.API_connection.close_position(order_id))
                                #     print(self.API_connection.get_overnight_fee("crypto", "BTCUSD"))

                            elif row[1] == 'Pending' or row[1] == 'Active':
                                self.current_active_options[row[0]] = row[13]
                                # print(self.current_active_options)
                            elif row[1] == 'cancel':
                                print(row)
                    except err:
                        print(err.error_details)
                        continue;

            else:
                self.throw_error_exception('Error: Invalid Google Sheet Data!')
            time.sleep(self.bot_timeinterval)

    def balance_market_run(self):

        if self.API_connection.check_connect() == False:
            check, reason = self.API_connection.connect()
            if not check:
                self.stop_thread = True
        ACTIVES = "EURUSD"
        duration = 1

        while not self.stop_thread:

            balances = self.API_connection.get_profile_ansyc()["balances"]

            for balance in balances:
                if balance["type"] == 1:
                    self.balance_information['REAL'] = balance
                if balance["type"] == 4:
                    self.balance_information['PRACTICE'] = balance

            time.sleep(10)

    def throw_error_exception(self, error_message):
        self.stop_thread = True
        self.SAMPLE_SPREADSHEET_ID_input = ''
        self.lineEdit_5.setText(error_message)
        self.lineEdit_5.setStyleSheet("color: rgb(255, 0, 0);")

    def stop_allthreading(self):
        self.stop_thread = True
        self.bot_timeinterval = 0
        self.SAMPLE_SPREADSHEET_ID_input = ''
        self.lineEdit_5.setText('Stopped')
        self.lineEdit_5.setStyleSheet("color: rgb(255, 198, 109);")

    def authenticate_google(self):
        """Shows basic usage of the Sheets API.
                    Prints values from a sample spreadsheet.
                    """
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(os.path.join(sys.path[0], 'token.json')):
            creds = Credentials.from_authorized_user_file(os.path.join(sys.path[0], 'token.json'), self.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    os.path.dirname(os.path.realpath(__file__)) + '/credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(os.path.join(sys.path[0], 'token.json'), 'w') as token:
                token.write(creds.to_json())

        self.service = build('sheets', 'v4', credentials=creds)
