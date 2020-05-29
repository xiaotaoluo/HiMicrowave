import sys
import os
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']

from PyQt5 import QtCore, QtWidgets, QtGui
from Ui_Main import Ui_MainWindow
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

import numpy as np
import xlsxwriter
import pyqtgraph as pg
import scipy.interpolate as spi

class HiMicrowave(QMainWindow,Ui_MainWindow):
    def __init__(self, parent=None):
        super(HiMicrowave,self).__init__(parent)
        self.setupUi(self)
        self.register_event()

        pg.setConfigOptions(antialias=True)
        pg.setConfigOption('background','w')
        self.plotgraph = pg.PlotWidget(title="S-Parameter")
        self.legendtext = self.plotgraph.addLegend()
        self.Waveform.addWidget(self.plotgraph)
        self.WaveformBox.setLayout(self.Waveform)


    def register_event(self):
        self.ImportButton.clicked.connect(self.import_btn_click)

    def import_btn_click(self):
        filename, filetype = QFileDialog.getOpenFileNames(self, 'Import Files', os.getcwd(), 
                                                            's2p files(*.s2p)')
        if len(filename) == 0:
            return
        else:
            s2p_data = []
            if len(filename) == 1:
                self.StatusLabel.setText(filename[0])
                self.NumLabel.setText('Num of File: 1')
            else:
                self.StatusLabel.setText(filename[0] + '\n ... \n' + filename[-1])
                self.NumLabel.setText('Num of Files: ' + str(len(filename)))
            for i in range(len(filename)):
                s2p_data.append(self.import_data(filename[i]))

            new_freq = np.arange(s2p_data[0][...,0][0]/1000000.0 , s2p_data[0][...,0][-1]/1000000.0,0.1)
            ipo3 = spi.splrep(s2p_data[0][...,0]/1000000.0,s2p_data[0][...,3],k=1)
            iy3 = spi.splev(new_freq,ipo3)
            self.plotgraph.plot(new_freq,iy3, pen=(0, 255, 0), name="new_S21")


            self.plotgraph.plot(s2p_data[0][...,0]/1000000.0,s2p_data[0][...,1], pen=(242, 242, 0), name="S11")
            self.plotgraph.plot(s2p_data[0][...,0]/1000000.0,s2p_data[0][...,7], pen=(255, 64, 64), name="S22")
            self.plotgraph.plot(s2p_data[0][...,0]/1000000.0,s2p_data[0][...,3], pen=(0, 127, 255), name="S21")


    def import_data(self,file_name):
        _s2p = np.loadtxt(file_name, dtype=float, comments=['!', '#'])
        return _s2p
        # _freq = _s2p[..., 0]
        # _mag_s11 = _s2p[..., 1]
        # _ang_s11 = _s2p[..., 2]
        # _mag_s21 = _s2p[..., 3]
        # _ang_s21 = _s2p[..., 4]
        # _mag_s12 = _s2p[..., 5]
        # _ang_s12 = _s2p[..., 6]
        # _mag_s22 = _s2p[..., 7]
        # _ang_s22 = _s2p[..., 8]






if __name__ == '__main__':
    app = QApplication(sys.argv)
    mainwindow = HiMicrowave()
    mainwindow.show()
    sys.exit(app.exec_())