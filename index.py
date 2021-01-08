from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from PyQt5.uic import loadUiType
import threading
from time import sleep

from utils import *

ui,_ = loadUiType('window.ui');


class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self);
        self.setupUi(self);
        self.setWindowTitle("PVT SYSTEM");
        self.serial_connection = SerialConnection(
            self.data_table_widget, 
            self.statusBar,
            self.connection_control_btn
            )
        
        # call functions on init
        self.init_view();
        self.thread_control();
        self.handle_button_event();

    
    def init_view(self):
        self.baud_selected_box.setCurrentIndex(2);
        pass
    
    def thread_control(self):
        port_check_tread = threading.Thread(target=update_port_view,args=(self.serial_connection,self.port_selected_box,),daemon=True);

        port_check_tread.start();

    def handle_button_event(self):
        self.connection_control_btn.clicked.connect(
            lambda:connection_btn_clicked(
                self.serial_connection,
                self.connection_control_btn,
                self.port_selected_box,
                self.baud_selected_box
                ));
        self.clear_data_btn.clicked.connect(self.serial_connection.clear_btn_handle);
        self.export_exel_btn.clicked.connect(self.serial_connection.export_exel_btn_handle);


        


def main():
    app = QApplication(sys.argv);
    window = MainApp()
    window.show();
    app.exec_();

if __name__ == '__main__':
    main()
