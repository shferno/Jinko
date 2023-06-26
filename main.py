import sys
from PyQt5.QtWidgets import QApplication, QWidget
from templates.weight import Ui_WeightRatioCalculator
import datetime
import pandas as pd
import os
from time import sleep
# from functools import partial1
path = "" # excel we want to store
class MyMainForm(QWidget, Ui_WeightRatioCalculator):
    def __init__(self):
        super(MyMainForm, self).__init__()
        self.setupUi(self)
        self.retranslateUi(self)
        self.Submit.clicked.connect(self.cal_rat)

    def cal_rat(self):
        ca = self.CAV.text()
        cb = self.CBV.text()
        ta = self.TAV.text()
        tb = self.TBV.text()
        # if not (ca.isnumeric() and cb.isnumeric and ta.isnumeric() and tb.isnumeric()):
        #     self.clear_form()
        mf = self.MFV.text()
        ins = self.INSPV.text()
        raw_time = datetime.datetime.now().strftime("%Y/%m/%d/%H/%M/%S")
        time = raw_time.split("/")
        try:
            ratio = round((float(ta) - float(ca)) / (float(tb) - float(cb)), 2)
            self.RV.append(f"{raw_time}\nthe mixture ratio of glue is  *** {ratio} *** \n")
            self.clear_form()
        except Exception:
            self.RV.append("XX Type in the wrong data, please do it again")
            # sleep(2)
            # self.RV.clear()
        df = pd.DataFrame(
            data = [[f"{time[1]}/{time[2]}/{time[0]}", f"{time[3]}:{time[4]}:{time[5]}", mf, ca, ta, float(ta) - float(ca), cb, tb, float(tb) - float(cb), ratio, "Y" if ratio >= 5.5 and ratio <= 6.0 else "N", ins]],
            columns = ["Date", "Time", "Silicon Gel Filling Manufacturer", "Cup Weight A1", "Total Weight A2", "Silicone Gel Weight A", "Cup Weight B1", "Total Weight B2", "Silicone Gel Weight", "Ration", "Whether qualified", "Inspection Personnel"]
        )
        if os.path.exists(os.getcwd() + r"\potting_glue_record.xlsx"):
            data = pd.read_excel(os.getcwd() + r"\potting_glue_record.xlsx")
            data = pd.concat([data, df], ignore_index = True)
            data.to_excel(os.getcwd() + r"\potting_glue_record.xlsx", index = None)
        else:
            df.to_excel(os.getcwd() + r"\potting_glue_record.xlsx", index = None)
        print("Success")
    def clear_form(self):
        self.CAV.clear()
        self.CBV.clear()
        self.TAV.clear()
        self.TBV.clear()
        self.MFV.clear()
        self.INSPV.clear()
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWin = MyMainForm()
    myWin.show()
    sys.exit(app.exec())