import uiautomation
import ruamel.yaml as yaml
import time

time.sleep(5)
elemnt = uiautomation.ControlFromCursor()

print(elemnt)
#
# f = open("testcase.yaml")
# rs = yaml.safe_load(f)
# f.close()
# print(rs['TestSet'])
# testitems=rs['TestSet']
#
# for i in dict(testitems).keys():
#     if  testitems[i] == "PASS" or testitems[i] == 'FAIL':
#         continue
#     if i == "test4":
#         testitems[i] = "FAIL"
#     else:
#         testitems[i] = "PASS"
# with open('testcase.yaml','w+') as f1:
#     yaml.dump(rs,f1, default_flow_style=False,allow_unicode=True)

