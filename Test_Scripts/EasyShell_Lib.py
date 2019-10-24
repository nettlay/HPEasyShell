import platform
from Library.CommonLib import QAUtils, TxtUtils, getElementByType
import os

if not os.path.exists('c:\\svc'):
    os.mkdir("C:\\svc")
if not os.path.exists('c:\\svc\\svcconfig.ini'):
    os.system('echo {}\\test.txt>c:\\svc\\svcconfig.ini'.format(os.getcwd()))
file_path = TxtUtils('c:\\svc\\svcconfig.ini').get_source().strip().split(':', 1)[1]
print(file_path)
file_path = os.path.dirname(file_path)
ElementlibPath = os.path.join(file_path, 'Configuration\\elementLib.ini')


def getElementMapping(filepath=ElementlibPath):
    """
    # element format:
    # define name:"name":automationid:controltype
    # eg.: OKButton:"OK":Button----->By name
    #      CancelButton:btnCancel:Button----->By automationId
    :param filepath: ElementlibPath
    :return: element
    """
    mappingDict = {}
    lines = TxtUtils(filepath).get_lines()
    for line in lines:
        if line[0] == '#':
            continue
        items = line.strip().split(":", 1)
        mappingDict[items[0]] = items[1]
    return mappingDict


def getElement(name, **kwargs):
    # name is defined name, format: defined name:"Name"/AutomationId:ControlType
    elementId = getElementMapping()[name].split(':')[0]
    controltype = getElementMapping()[name].split(':')[1].upper()
    if elementId.startswith('"') and elementId.endswith('"'):
        return getElementByType(controltype, RegexName=elementId.replace('"', ''), **kwargs)
    else:
        return getElementByType(controltype, AutomationId=elementId, **kwargs)


class CommonUtils(QAUtils):
    @staticmethod
    def SwitchToUser():
        QAUtils.SwitchUser("User", "User", "")

    @staticmethod
    def SwitchToAdmin():
        logon_user = platform.version()
        if logon_user.split(".")[0] == "10":
            QAUtils.SwitchUser("Admin", "Admin", "")
        else:
            QAUtils.SwitchUser("Administrator", "Administrator", "")

    @staticmethod
    def launchFromControl():
        QAUtils.LaunchAppFromControl("HP Easy Shell Configuration")

    @staticmethod
    def launchFromPath():
        QAUtils.LaunchAppFromFile("C:\\Program Files\\HP\\HP Easy Shell\\HPEasyShellConfig.exe")

    @staticmethod
    def install(path):
        os.system('msiexec.exe /q /i {}'.format(path))
