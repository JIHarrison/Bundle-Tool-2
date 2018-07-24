from PyQt5 import QtCore, QtGui, QtWidgets
import sys
from mailmerge import MailMerge
import itertools
import datetime
import openpyxl
import pandas as pd

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(444, 755)
        MainWindow.setMinimumSize(QtCore.QSize(426, 487))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.parts_number_label = QtWidgets.QLabel(self.centralwidget)
        self.parts_number_label.setGeometry(QtCore.QRect(130, 10, 91, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.parts_number_label.setFont(font)
        self.parts_number_label.setObjectName("parts_number_label")
        self.shop_hours_cost_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.shop_hours_cost_lineEdit.setGeometry(QtCore.QRect(300, 440, 113, 20))
        self.shop_hours_cost_lineEdit.setObjectName("shop_hours_cost_lineEdit")
        self.spinBox = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox.setGeometry(QtCore.QRect(240, 40, 42, 22))
        self.spinBox.setObjectName("spinBox")
        self.tubesheet_unit_cost_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.tubesheet_unit_cost_lineEdit.setGeometry(QtCore.QRect(300, 40, 113, 20))
        self.tubesheet_unit_cost_lineEdit.setObjectName("tubesheet_unit_cost_lineEdit")
        self.materials_cost_label = QtWidgets.QLabel(self.centralwidget)
        self.materials_cost_label.setGeometry(QtCore.QRect(310, 340, 91, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.materials_cost_label.setFont(font)
        self.materials_cost_label.setObjectName("materials_cost_label")
        self.spinBox_2 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_2.setGeometry(QtCore.QRect(240, 80, 42, 22))
        self.spinBox_2.setObjectName("spinBox_2")
        self.spinBox_4 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_4.setGeometry(QtCore.QRect(240, 160, 42, 22))
        self.spinBox_4.setObjectName("spinBox_4")
        self.part_number_lineEdit_6 = QtWidgets.QLineEdit(self.centralwidget)
        self.part_number_lineEdit_6.setGeometry(QtCore.QRect(120, 240, 113, 20))
        self.part_number_lineEdit_6.setObjectName("part_number_lineEdit_6")
        self.hex_nuts_unit_cost_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.hex_nuts_unit_cost_lineEdit.setGeometry(QtCore.QRect(300, 280, 113, 20))
        self.hex_nuts_unit_cost_lineEdit.setObjectName("hex_nuts_unit_cost_lineEdit")
        self.calculated_materials_label = QtWidgets.QLabel(self.centralwidget)
        self.calculated_materials_label.setGeometry(QtCore.QRect(310, 360, 91, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.calculated_materials_label.setFont(font)
        self.calculated_materials_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.calculated_materials_label.setObjectName("calculated_materials_label")
        self.studs_label = QtWidgets.QLabel(self.centralwidget)
        self.studs_label.setGeometry(QtCore.QRect(20, 240, 47, 13))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.studs_label.setFont(font)
        self.studs_label.setObjectName("studs_label")
        self.redraw_cost_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.redraw_cost_lineEdit.setGeometry(QtCore.QRect(300, 400, 113, 20))
        self.redraw_cost_lineEdit.setObjectName("redraw_cost_lineEdit")
        self.spinBox_5 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_5.setGeometry(QtCore.QRect(240, 200, 42, 22))
        self.spinBox_5.setObjectName("spinBox_5")
        self.gaskets2_label = QtWidgets.QLabel(self.centralwidget)
        self.gaskets2_label.setGeometry(QtCore.QRect(20, 200, 47, 13))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.gaskets2_label.setFont(font)
        self.gaskets2_label.setObjectName("gaskets2_label")
        self.markup_SpinBox = QtWidgets.QDoubleSpinBox(self.centralwidget)
        self.markup_SpinBox.setGeometry(QtCore.QRect(240, 550, 61, 22))
        self.markup_SpinBox.setSingleStep(0.25)
        self.markup_SpinBox.setObjectName("markup_SpinBox")
        self.tubes_label = QtWidgets.QLabel(self.centralwidget)
        self.tubes_label.setGeometry(QtCore.QRect(20, 80, 47, 13))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.tubes_label.setFont(font)
        self.tubes_label.setObjectName("tubes_label")
        self.shop_bundle_label = QtWidgets.QLabel(self.centralwidget)
        self.shop_bundle_label.setGeometry(QtCore.QRect(20, 440, 91, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.shop_bundle_label.setFont(font)
        self.shop_bundle_label.setObjectName("shop_bundle_label")
        self.redraw_label = QtWidgets.QLabel(self.centralwidget)
        self.redraw_label.setGeometry(QtCore.QRect(20, 400, 47, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.redraw_label.setFont(font)
        self.redraw_label.setObjectName("redraw_label")
        self.part_number_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.part_number_lineEdit.setGeometry(QtCore.QRect(120, 40, 113, 20))
        self.part_number_lineEdit.setObjectName("part_number_lineEdit")
        self.tubes_unit_cost_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.tubes_unit_cost_lineEdit.setGeometry(QtCore.QRect(300, 80, 113, 20))
        self.tubes_unit_cost_lineEdit.setObjectName("tubes_unit_cost_lineEdit")
        self.part_number_lineEdit_4 = QtWidgets.QLineEdit(self.centralwidget)
        self.part_number_lineEdit_4.setGeometry(QtCore.QRect(120, 160, 113, 20))
        self.part_number_lineEdit_4.setObjectName("part_number_lineEdit_4")
        self.gaskets_unit_cost_lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.gaskets_unit_cost_lineEdit_2.setGeometry(QtCore.QRect(300, 200, 113, 20))
        self.gaskets_unit_cost_lineEdit_2.setObjectName("gaskets_unit_cost_lineEdit_2")
        self.spinBox_3 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_3.setGeometry(QtCore.QRect(240, 120, 42, 22))
        self.spinBox_3.setObjectName("spinBox_3")
        self.spinBox_6 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_6.setGeometry(QtCore.QRect(240, 240, 42, 22))
        self.spinBox_6.setObjectName("spinBox_6")
        self.markup_label = QtWidgets.QLabel(self.centralwidget)
        self.markup_label.setGeometry(QtCore.QRect(240, 530, 47, 13))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.markup_label.setFont(font)
        self.markup_label.setObjectName("markup_label")
        self.hex_label = QtWidgets.QLabel(self.centralwidget)
        self.hex_label.setGeometry(QtCore.QRect(20, 280, 51, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.hex_label.setFont(font)
        self.hex_label.setObjectName("hex_label")
        self.tubesheet_label = QtWidgets.QLabel(self.centralwidget)
        self.tubesheet_label.setGeometry(QtCore.QRect(6, 40, 61, 20))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.tubesheet_label.setFont(font)
        self.tubesheet_label.setObjectName("tubesheet_label")
        self.studs_unit_cost_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.studs_unit_cost_lineEdit.setGeometry(QtCore.QRect(300, 240, 113, 20))
        self.studs_unit_cost_lineEdit.setObjectName("studs_unit_cost_lineEdit")
        self.part_number_lineEdit_7 = QtWidgets.QLineEdit(self.centralwidget)
        self.part_number_lineEdit_7.setGeometry(QtCore.QRect(120, 280, 113, 20))
        self.part_number_lineEdit_7.setObjectName("part_number_lineEdit_7")
        self.baffles_label = QtWidgets.QLabel(self.centralwidget)
        self.baffles_label.setGeometry(QtCore.QRect(20, 120, 47, 13))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.baffles_label.setFont(font)
        self.baffles_label.setObjectName("baffles_label")
        self.spinBox_9 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_9.setGeometry(QtCore.QRect(240, 400, 42, 22))
        self.spinBox_9.setObjectName("spinBox_9")
        self.spinBox_7 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_7.setGeometry(QtCore.QRect(240, 280, 42, 22))
        self.spinBox_7.setObjectName("spinBox_7")
        self.part_number_lineEdit_2 = QtWidgets.QLineEdit(self.centralwidget)
        self.part_number_lineEdit_2.setGeometry(QtCore.QRect(120, 80, 113, 20))
        self.part_number_lineEdit_2.setObjectName("part_number_lineEdit_2")
        self.baffles_unit_cost_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.baffles_unit_cost_lineEdit.setGeometry(QtCore.QRect(300, 120, 113, 20))
        self.baffles_unit_cost_lineEdit.setObjectName("baffles_unit_cost_lineEdit")
        self.per_item_label = QtWidgets.QLabel(self.centralwidget)
        self.per_item_label.setGeometry(QtCore.QRect(310, 10, 81, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.per_item_label.setFont(font)
        self.per_item_label.setObjectName("per_item_label")
        self.spinBox_10 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_10.setGeometry(QtCore.QRect(240, 440, 42, 22))
        self.spinBox_10.setObjectName("spinBox_10")
        self.gaskets_label = QtWidgets.QLabel(self.centralwidget)
        self.gaskets_label.setGeometry(QtCore.QRect(20, 160, 47, 13))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.gaskets_label.setFont(font)
        self.gaskets_label.setObjectName("gaskets_label")
        self.part_number_lineEdit_3 = QtWidgets.QLineEdit(self.centralwidget)
        self.part_number_lineEdit_3.setGeometry(QtCore.QRect(120, 120, 113, 20))
        self.part_number_lineEdit_3.setObjectName("part_number_lineEdit_3")
        self.part_number_lineEdit_5 = QtWidgets.QLineEdit(self.centralwidget)
        self.part_number_lineEdit_5.setGeometry(QtCore.QRect(120, 200, 113, 20))
        self.part_number_lineEdit_5.setObjectName("part_number_lineEdit_5")
        self.gaskets_unit_cost_lineEdit_1 = QtWidgets.QLineEdit(self.centralwidget)
        self.gaskets_unit_cost_lineEdit_1.setGeometry(QtCore.QRect(300, 160, 113, 20))
        self.gaskets_unit_cost_lineEdit_1.setObjectName("gaskets_unit_cost_lineEdit_1")
        self.calculate_pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.calculate_pushButton.setGeometry(QtCore.QRect(340, 670, 91, 23))
        self.calculate_pushButton.setObjectName("calculate_pushButton")
        self.total_cost_label = QtWidgets.QLabel(self.centralwidget)
        self.total_cost_label.setGeometry(QtCore.QRect(310, 480, 71, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.total_cost_label.setFont(font)
        self.total_cost_label.setObjectName("total_cost_label")
        self.calculated_total_label = QtWidgets.QLabel(self.centralwidget)
        self.calculated_total_label.setGeometry(QtCore.QRect(310, 500, 91, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.calculated_total_label.setFont(font)
        self.calculated_total_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.calculated_total_label.setObjectName("calculated_total_label")
        self.rep_cost_label = QtWidgets.QLabel(self.centralwidget)
        self.rep_cost_label.setGeometry(QtCore.QRect(310, 580, 71, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.rep_cost_label.setFont(font)
        self.rep_cost_label.setObjectName("rep_cost_label")
        self.calculated_rep_cost_label = QtWidgets.QLabel(self.centralwidget)
        self.calculated_rep_cost_label.setGeometry(QtCore.QRect(310, 600, 91, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.calculated_rep_cost_label.setFont(font)
        self.calculated_rep_cost_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.calculated_rep_cost_label.setObjectName("calculated_rep_cost_label")
        self.list_price_label = QtWidgets.QLabel(self.centralwidget)
        self.list_price_label.setGeometry(QtCore.QRect(310, 620, 71, 16))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.list_price_label.setFont(font)
        self.list_price_label.setObjectName("list_price_label")
        self.calculated_list_price_label = QtWidgets.QLabel(self.centralwidget)
        self.calculated_list_price_label.setGeometry(QtCore.QRect(310, 640, 91, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.calculated_list_price_label.setFont(font)
        self.calculated_list_price_label.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.calculated_list_price_label.setObjectName("calculated_list_price_label")
        self.markup_SpinBox_2 = QtWidgets.QDoubleSpinBox(self.centralwidget)
        self.markup_SpinBox_2.setGeometry(QtCore.QRect(240, 630, 61, 22))
        self.markup_SpinBox_2.setSingleStep(0.05)
        self.markup_SpinBox_2.setObjectName("markup_SpinBox_2")
        self.markup_label_2 = QtWidgets.QLabel(self.centralwidget)
        self.markup_label_2.setGeometry(QtCore.QRect(240, 610, 47, 13))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.markup_label_2.setFont(font)
        self.markup_label_2.setObjectName("markup_label_2")
        self.misc_label = QtWidgets.QLabel(self.centralwidget)
        self.misc_label.setGeometry(QtCore.QRect(150, 310, 81, 21))
        font = QtGui.QFont()
        font.setBold(True)
        font.setWeight(75)
        self.misc_label.setFont(font)
        self.misc_label.setObjectName("misc_label")
        self.spinBox_8 = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_8.setGeometry(QtCore.QRect(240, 310, 42, 22))
        self.spinBox_8.setObjectName("spinBox_8")
        self.misc_unit_cost_lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.misc_unit_cost_lineEdit.setGeometry(QtCore.QRect(300, 310, 113, 20))
        self.misc_unit_cost_lineEdit.setObjectName("misc_unit_cost_lineEdit")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 444, 21))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionSave = QtWidgets.QAction(MainWindow)
        self.actionSave.setObjectName("actionSave")
        self.menuFile.addAction(self.actionSave)
        self.menubar.addAction(self.menuFile.menuAction())

        self.actionExport_to_Excel = QtWidgets.QAction(MainWindow)
        self.actionExport_to_Excel.setObjectName("actionExport_to_Excel")
        self.menuFile.addAction(self.actionExport_to_Excel)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        MainWindow.setTabOrder(self.part_number_lineEdit, self.spinBox)
        MainWindow.setTabOrder(self.spinBox, self.tubesheet_unit_cost_lineEdit)
        MainWindow.setTabOrder(self.tubesheet_unit_cost_lineEdit, self.part_number_lineEdit_2)
        MainWindow.setTabOrder(self.part_number_lineEdit_2, self.spinBox_2)
        MainWindow.setTabOrder(self.spinBox_2, self.tubes_unit_cost_lineEdit)
        MainWindow.setTabOrder(self.tubes_unit_cost_lineEdit, self.part_number_lineEdit_3)
        MainWindow.setTabOrder(self.part_number_lineEdit_3, self.spinBox_3)
        MainWindow.setTabOrder(self.spinBox_3, self.baffles_unit_cost_lineEdit)
        MainWindow.setTabOrder(self.baffles_unit_cost_lineEdit, self.part_number_lineEdit_4)
        MainWindow.setTabOrder(self.part_number_lineEdit_4, self.spinBox_4)
        MainWindow.setTabOrder(self.spinBox_4, self.gaskets_unit_cost_lineEdit_1)
        MainWindow.setTabOrder(self.gaskets_unit_cost_lineEdit_1, self.part_number_lineEdit_5)
        MainWindow.setTabOrder(self.part_number_lineEdit_5, self.spinBox_5)
        MainWindow.setTabOrder(self.spinBox_5, self.gaskets_unit_cost_lineEdit_2)
        MainWindow.setTabOrder(self.gaskets_unit_cost_lineEdit_2, self.part_number_lineEdit_6)
        MainWindow.setTabOrder(self.part_number_lineEdit_6, self.spinBox_6)
        MainWindow.setTabOrder(self.spinBox_6, self.studs_unit_cost_lineEdit)
        MainWindow.setTabOrder(self.studs_unit_cost_lineEdit, self.part_number_lineEdit_7)
        MainWindow.setTabOrder(self.part_number_lineEdit_7, self.spinBox_7)
        MainWindow.setTabOrder(self.spinBox_7, self.hex_nuts_unit_cost_lineEdit)
        MainWindow.setTabOrder(self.hex_nuts_unit_cost_lineEdit, self.spinBox_8)
        MainWindow.setTabOrder(self.spinBox_8, self.misc_unit_cost_lineEdit)
        MainWindow.setTabOrder(self.misc_unit_cost_lineEdit, self.spinBox_9)
        MainWindow.setTabOrder(self.spinBox_9, self.redraw_cost_lineEdit)
        MainWindow.setTabOrder(self.redraw_cost_lineEdit, self.spinBox_10)
        MainWindow.setTabOrder(self.spinBox_10, self.shop_hours_cost_lineEdit)
        MainWindow.setTabOrder(self.shop_hours_cost_lineEdit, self.markup_SpinBox)
        MainWindow.setTabOrder(self.markup_SpinBox, self.markup_SpinBox_2)
        MainWindow.setTabOrder(self.markup_SpinBox_2, self.calculate_pushButton)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.parts_number_label.setText(_translate("MainWindow", "Parts Numbers"))
        self.shop_hours_cost_lineEdit.setText(_translate("MainWindow", "0.0"))
        self.tubesheet_unit_cost_lineEdit.setText(_translate("MainWindow", "0.0"))
        self.materials_cost_label.setText(_translate("MainWindow", "Materials Cost:"))
        self.hex_nuts_unit_cost_lineEdit.setText(_translate("MainWindow", "0.0"))
        self.calculated_materials_label.setText(_translate("MainWindow", "0.0"))
        self.studs_label.setText(_translate("MainWindow", "Studs"))
        self.redraw_cost_lineEdit.setText(_translate("MainWindow", "0.0"))
        self.gaskets2_label.setText(_translate("MainWindow", "Gaskets"))
        self.tubes_label.setText(_translate("MainWindow", "Tubes"))
        self.shop_bundle_label.setText(_translate("MainWindow", "Shop Hours"))
        self.redraw_label.setText(_translate("MainWindow", "Redraw"))
        self.tubes_unit_cost_lineEdit.setText(_translate("MainWindow", "0.0"))
        self.gaskets_unit_cost_lineEdit_2.setText(_translate("MainWindow", "0.0"))
        self.markup_label.setText(_translate("MainWindow", "Mark Up"))
        self.hex_label.setText(_translate("MainWindow", "Hex Nuts"))
        self.tubesheet_label.setText(_translate("MainWindow", "Tubesheet"))
        self.studs_unit_cost_lineEdit.setText(_translate("MainWindow", "0.0"))
        self.baffles_label.setText(_translate("MainWindow", "Baffles"))
        self.baffles_unit_cost_lineEdit.setText(_translate("MainWindow", "0.0"))
        self.per_item_label.setText(_translate("MainWindow", "Cost Per Item"))
        self.gaskets_label.setText(_translate("MainWindow", "Gaskets"))
        self.gaskets_unit_cost_lineEdit_1.setText(_translate("MainWindow", "0.0"))
        self.calculate_pushButton.setText(_translate("MainWindow", "Calculate"))
        self.total_cost_label.setText(_translate("MainWindow", "Total Cost:"))
        self.calculated_total_label.setText(_translate("MainWindow", "0.0"))
        self.rep_cost_label.setText(_translate("MainWindow", "Rep Cost:"))
        self.calculated_rep_cost_label.setText(_translate("MainWindow", "0.0"))
        self.list_price_label.setText(_translate("MainWindow", "List Price:"))
        self.calculated_list_price_label.setText(_translate("MainWindow", "0.0"))
        self.markup_label_2.setText(_translate("MainWindow", "Mark Up"))
        self.misc_label.setText(_translate("MainWindow", "Miscellaneous"))
        self.misc_unit_cost_lineEdit.setText(_translate("MainWindow", "0.0"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.actionSave.setText(_translate("MainWindow", "Save"))
        self.actionSave.setShortcut(_translate("MainWindow", "Ctrl+S"))
        self.actionExport_to_Excel.setText(_translate("MainWindow", "Export to Excel"))




        ########################################################################
        # copy below before replacing designer generated code for entire class: Ui_misc_Dialog
        ########################################################################


        # connections to functions
        self.tubesheet_unit_cost_lineEdit.editingFinished.connect(self.set_tubesheet_unit_cost)
        self.tubes_unit_cost_lineEdit.editingFinished.connect(self.set_tubes_unit_cost)
        self.baffles_unit_cost_lineEdit.editingFinished.connect(self.set_baffles_unit_cost)
        self.gaskets_unit_cost_lineEdit_1.editingFinished.connect(self.set_gaskets_unit_cost)
        self.gaskets_unit_cost_lineEdit_2.editingFinished.connect(self.set_gaskets2_unit_cost)
        self.studs_unit_cost_lineEdit.editingFinished.connect(self.set_studs_unit_cost)
        self.hex_nuts_unit_cost_lineEdit.editingFinished.connect(self.set_hex_unit_cost)
        self.misc_unit_cost_lineEdit.editingFinished.connect(self.set_misc_unit_cost)
        self.redraw_cost_lineEdit.editingFinished.connect(self.set_redraw_unit_cost)
        self.shop_hours_cost_lineEdit.editingFinished.connect(self.set_shop_hrs_unit_cost)
        self.part_number_lineEdit.editingFinished.connect(self.set_tubesheet_parts)
        self.part_number_lineEdit_2.editingFinished.connect(self.set_tubes_parts)
        self.part_number_lineEdit_3.editingFinished.connect(self.set_baffles_parts)
        self.part_number_lineEdit_4.editingFinished.connect(self.set_gaskets1_parts)
        self.part_number_lineEdit_5.editingFinished.connect(self.set_gaskets2_parts)
        self.part_number_lineEdit_6.editingFinished.connect(self.set_studs_parts)
        self.part_number_lineEdit_7.editingFinished.connect(self.set_hex_parts)

        self.spinBox.editingFinished.connect(self.set_tubesheet_qty)
        self.spinBox_2.editingFinished.connect(self.set_tubes_qty)
        self.spinBox_3.editingFinished.connect(self.set_baffles_qty)
        self.spinBox_4.editingFinished.connect(self.set_gaskets_qty)
        self.spinBox_5.editingFinished.connect(self.set_gaskets2_qty)
        self.spinBox_6.editingFinished.connect(self.set_studs_qty)
        self.spinBox_7.editingFinished.connect(self.set_hex_qty)
        self.spinBox_8.editingFinished.connect(self.set_misc_qty)
        self.spinBox_9.editingFinished.connect(self.set_redraw_qty)
        self.spinBox_10.editingFinished.connect(self.set_shop_hrs_qty)
        self.markup_SpinBox.editingFinished.connect(self.set_markup1)
        self.markup_SpinBox_2.editingFinished.connect(self.set_markup2)

        self.calculate_pushButton.clicked.connect(self.calculate_total_costs)


        self.actionSave.triggered.connect(self.file_save)
        self.actionExport_to_Excel.triggered.connect(self.excel_export)

        # lineEdit validators restricting input to integers and 2 decimal places
        self.baffles_unit_cost_lineEdit.setInputMask('')
        self.misc_unit_cost_lineEdit.setInputMask('')
        self.tubesheet_unit_cost_lineEdit.setInputMask('')
        self.tubes_unit_cost_lineEdit.setInputMask('')
        self.gaskets_unit_cost_lineEdit_1.setInputMask('')
        self.gaskets_unit_cost_lineEdit_2.setInputMask('')
        self.studs_unit_cost_lineEdit.setInputMask('')
        self.hex_nuts_unit_cost_lineEdit.setInputMask('')
        self.redraw_cost_lineEdit.setInputMask('')
        self.shop_hours_cost_lineEdit.setInputMask('')

        regexp = QtCore.QRegExp('^|[0-9]*(\.[0-9][0-9]?)?')
        validator = QtGui.QRegExpValidator(regexp)

        self.baffles_unit_cost_lineEdit.setValidator(validator)
        self.misc_unit_cost_lineEdit.setValidator(validator)
        self.tubesheet_unit_cost_lineEdit.setValidator(validator)
        self.tubes_unit_cost_lineEdit.setValidator(validator)
        self.gaskets_unit_cost_lineEdit_1.setValidator(validator)
        self.gaskets_unit_cost_lineEdit_2.setValidator(validator)
        self.studs_unit_cost_lineEdit.setValidator(validator)
        self.hex_nuts_unit_cost_lineEdit.setValidator(validator)
        self.redraw_cost_lineEdit.setValidator(validator)
        self.shop_hours_cost_lineEdit.setValidator(validator)
        self.baffles_unit_cost_lineEdit.setValidator(validator)

        self.baffles_unit_cost_lineEdit.setCursorPosition(0)

    def create_dicts(self, qty, cost, totals):
        docx_dict.clear()
        docx_dict.extend([{'item': i, 'parts_number': p, 'qty': q, 'unit_cost': c, 'total_cost': str(t)} for i, p, q, c, t in
         itertools.zip_longest(item, parts, qty, cost, totals, fillvalue='N/A')])

    def create_merge2(self, rep, list_price, markup, material_cost, redraw_eng_cost, markup_multiplier2):
        merge2.clear()
        date_time = datetime.date.today()
        merge2.update({'rep_cost': str(rep), 'list_price': str(list_price), 'markup': str(markup), 'materials_cost': str(material_cost), 'complete_cost':str(redraw_eng_cost), 'markup2':str(markup_multiplier2), 'date_time':str(date_time)})

    def write_final_options(self, name):
        template = "test-print.docx"
        document = MailMerge(template)
        document.merge_rows('item', docx_dict)
        document.merge(**merge2)
        if name.endswith('.docx'):
            document.write(name)
        else:
            document.write(name + '.docx')
        document.close()

    def file_save(self):
        name = QtWidgets.QFileDialog.getSaveFileName()
        self.write_final_options(str(name[0]))

    def excel_export(self):
        pass

    def set_tubesheet_qty(self):
        item_qtys[0] = self.spinBox.text()

    def set_tubes_qty(self):
        item_qtys[1] = self.spinBox_2.text()

    def set_baffles_qty(self):
        item_qtys[2] = self.spinBox_3.text()

    def set_gaskets_qty(self):
        item_qtys[3] = self.spinBox_4.text()

    def set_gaskets2_qty(self):
        item_qtys[4] = self.spinBox_5.text()

    def set_studs_qty(self):
        item_qtys[5] = self.spinBox_6.text()

    def set_hex_qty(self):
        item_qtys[6] = self.spinBox_7.text()

    def set_misc_qty(self):
        item_qtys[7] = self.spinBox_8.text()

    def set_redraw_qty(self):
        item_qtys[8] = self.spinBox_9.text()

    def set_shop_hrs_qty(self):
        item_qtys[9] = self.spinBox_10.text()

    def set_tubesheet_unit_cost(self):
        unit_costs[0] = self.tubesheet_unit_cost_lineEdit.text()

    def set_tubes_unit_cost(self):
        unit_costs[1] = self.tubes_unit_cost_lineEdit.text()

    def set_baffles_unit_cost(self):
        unit_costs[2] = self.baffles_unit_cost_lineEdit.text()

    def set_gaskets_unit_cost(self):
        unit_costs[3] = self.gaskets_unit_cost_lineEdit_1.text()

    def set_gaskets2_unit_cost(self):
        unit_costs[4] = self.gaskets_unit_cost_lineEdit_2.text()

    def set_studs_unit_cost(self):
        unit_costs[5] = self.studs_unit_cost_lineEdit.text()

    def set_hex_unit_cost(self):
        unit_costs[6] = self.hex_nuts_unit_cost_lineEdit.text()

    def set_misc_unit_cost(self):
        unit_costs[7] = self.misc_unit_cost_lineEdit.text()

    def set_redraw_unit_cost(self):
        unit_costs[8] = self.redraw_cost_lineEdit.text()

    def set_shop_hrs_unit_cost(self):
        unit_costs[9] = self.shop_hours_cost_lineEdit.text()

    def set_markup1(self):
        markup_multiplier = self.markup_SpinBox.text()

    def set_markup2(self):
        markup_multiplier2 = self.markup_SpinBox_2.text()

    def set_tubesheet_parts(self):
        parts_numbers[0] = self.part_number_lineEdit.text()

    def set_tubes_parts(self):
        parts_numbers[1] = self.part_number_lineEdit_2.text()

    def set_baffles_parts(self):
        parts_numbers[2] = self.part_number_lineEdit_3.text()

    def set_gaskets1_parts(self):
        parts_numbers[3] = self.part_number_lineEdit_4.text()

    def set_gaskets2_parts(self):
        parts_numbers[4] = self.part_number_lineEdit_5.text()

    def set_studs_parts(self):
        parts_numbers[5] = self.part_number_lineEdit_6.text()

    def set_hex_parts(self):
        parts_numbers[6] = self.part_number_lineEdit_7.text()

    def costs_qtys(self):
        for x in unit_costs:
            if x == "":
                x = 0
                unit_costs_float.append(float(x))
            else:
                unit_costs_float.append(float(x))
        for y in item_qtys:
            item_qtys_float.append(float(y))
        total_costs = [x * y for x, y in zip(unit_costs_float, item_qtys_float)]
        item_total = total_costs
        materials_cost = sum(total_costs[:-2])
        material_cost = materials_cost
        redraw_eng_cost = sum(total_costs)
        markup_multiplier = self.markup_SpinBox.text()
        markup_multiplier2 = self.markup_SpinBox_2.text()
        rep_cost = (redraw_eng_cost * float(markup_multiplier))
        list_price = rep_cost + (rep_cost * float(markup_multiplier2))
        self.calculated_total_label.setText(str(redraw_eng_cost))
        self.calculated_materials_label.setText(str(materials_cost))
        self.calculated_rep_cost_label.setText(str(rep_cost))
        self.calculated_list_price_label.setText(str(list_price))
        self.create_dicts(item_qtys, unit_costs, total_costs)
        self.create_merge2(rep_cost, list_price, markup_multiplier, material_cost, redraw_eng_cost, markup_multiplier2)
        del total_costs[:]
        del unit_costs_float[:]
        del item_qtys_float[:]

    def calculate_total_costs(self):
        self.costs_qtys()

    def create_dataframe(self):
        pass

parts_numbers = ["", "", "", "", "", "", ""]
materials_cost = 0.0
unit_costs = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
unit_costs_float = []
item_qtys = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
item_qtys_float = []
markup_multiplier2 = 0.0
markup_multiplier = 0.0

item = ['Tubesheet', 'Tubes', 'Baffles', 'Gaskets', 'Gaskets', 'Studs', 'Hex Nuts', 'Miscellaneous', 'Redraw', 'Shop Hours']
qty = item_qtys
unit_cost = unit_costs
item_total = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
keys = ['item', 'part_number', 'qty', 'unit_cost', 'item_total']
parts = parts_numbers
docx_dict = []
merge2 = {}

class ApplicationWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(ApplicationWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    application = ApplicationWindow()
    application.show()
    sys.exit(app.exec_())
