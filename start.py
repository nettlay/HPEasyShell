from Test_Scripts.easyshell import *
import os
from openpyxl import load_workbook
import pysnooper


class Test:
    def __init__(self):
        easyshelltest = EasyShellTest()
        self.path = easyshelltest.path
        self.result = True  # sigle test case result
        self.casepath = easyshelltest.casepath
        self.logpath = easyshelltest.log_path
        self.misc = easyshelltest.misc
        self.testset = easyshelltest.testset

    def test(self):
        wb = load_workbook(self.testset)
        sheets = wb.sheetnames  # 获得表单名字
        ws = wb[sheets[0]]
        rows = ws.max_row
        for i in range(2, rows + 1):
            self.result = True
            testName = ws.cell(row=i, column=1).value
            result = ws.cell(row=i, column=2).value
            if testName is None:
                break
            if '#' in testName:
                continue
            if result == 'PASS' or result == 'FAIL' or result == 'N/A':
                continue
            else:
                self.case = os.path.join(self.casepath, '{}.xlsx'.format(testName))
                self.runTestcase(self.case)
            if self.result:
                ws.cell(row=i, column=2).value = "PASS"
                wb.save(self.testset)
            else:
                ws.cell(row=i, column=2).value = "FAIL"
                wb.save(self.testset)

    def runTestcase(self, name):
        wb = load_workbook(name)
        sheets = wb.sheetnames  # 获得表单名字
        ws = wb[sheets[0]]
        rows = ws.max_row
        for i in range(2, rows + 1):
            checkPoint = ws.cell(row=i, column=1).value
            result = ws.cell(row=i, column=4).value
            command = ws.cell(row=i, column=2).value
            value = ws.cell(row=i, column=3).value
            if checkPoint is None:
                break
            if '#' in checkPoint:
                continue
            if result == 'PASS':
                continue
            if result == 'FAIL':
                self.result = False
                continue
            if checkPoint.upper() == "Y":
                rs = eval(command)
                print(rs)
                if value is None:
                    if rs:
                        ws.cell(row=i, column=4).value = "PASS"
                        wb.save(name)
                        EasyShellTest().Logfile("--->[Pass]:{} check".format(command))
                    else:
                        self.result = False
                        ws.cell(row=i, column=4).value = "FAIL"
                        wb.save(name)
                        EasyShellTest().Logfile("--->[Fail]:{} check, Expect:{},Actual:{}".format(command, "True", rs))
                else:
                    if '$NOT' in str(value).upper():
                        value = str(value).upper().replace('$NOT', '').strip()
                        if str(rs).upper() != value:
                            ws.cell(row=i, column=4).value = "PASS"
                            wb.save(name)
                            EasyShellTest().Logfile("--->[Pass]:{} check".format(command))
                        else:
                            self.result = False
                            ws.cell(row=i, column=4).value = "FAIL"
                            wb.save(name)
                            EasyShellTest().Logfile("--->[Fail]:{} check, Expect:{},Actual:{}".format(command, value, rs))
                    elif '$BETWEEN' in str(value).upper():
                        value = str(value).upper().replace('$NOT', '').strip()
                        rs = int(rs)
                        v1, v2 = value.split(',')
                        v1 = int(v1.strip())
                        v2 = int(v2.strip())
                        if v1 < rs < v2:
                            ws.cell(row=i, column=4).value = "PASS"
                            wb.save(name)
                            EasyShellTest().Logfile("--->[Pass]:{} check".format(command))
                        else:
                            self.result = False
                            ws.cell(row=i, column=4).value = "FAIL"
                            wb.save(name)
                            EasyShellTest().Logfile("--->[Fail]:{} check, Expect:{},Actual:{}".format(command, value, rs))
                    else:
                        if str(rs).upper() == str(value).upper():
                            ws.cell(row=i, column=4).value = "PASS"
                            wb.save(name)
                            EasyShellTest().Logfile("--->[Pass]:{} check".format(command))
                        else:
                            self.result = False
                            ws.cell(row=i, column=4).value = "FAIL"
                            wb.save(name)
                            EasyShellTest().Logfile("--->[Fail]:{} check, Expect:{},Actual:{}".format(command, value, rs))
            elif checkPoint.upper() == 'N':
                if 'Reboot' in str(command):
                    ws.cell(row=i, column=4).value = "PASS"
                    wb.save(name)
                    exec(command)
                else:
                    print(command)
                    exec(command)
                    ws.cell(row=i, column=4).value = "PASS"
                    wb.save(name)
                    print('save no check')
            else:
                continue
        wb.save(name)


if __name__ == '__main__':
    Test().test()
    import uiautomation