import servicemanager
import socket
import sys, os
import win32event
import win32security
import win32service
import win32serviceutil
import win32api, win32con
import win32ts
import win32process
import shutil


class TestService(win32serviceutil.ServiceFramework):
    _svc_name_ = "AutoTestSvc"
    _svc_display_name_ = "AutoTestSvc"
    _svc_description_ = "Start Application as system account."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.filePath = ''
        self.loadConfig()

    def loadConfig(self):
        if os.path.exists("c:\\svc\\svcconfig.ini"):
            pass
        else:
            if os.path.exists("C:\\svc"):
                pass
            else:
                os.mkdir("C:\\svc")
            curPth=os.getcwd()
            configPath = os.path.join(curPth,'svcconfig.ini')
            shutil.copy(configPath, "c:\\svc\\svcconfig.ini")
        with open("c:\\svc\\svcconfig.ini") as f:
            rs = f.readlines()
            for line in rs:
                if line.split(':',1)[0].strip().upper() == 'FILEPATH':
                    self.filePath = line.split(':',1)[1].strip()
        print(self.filePath)

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def runProcessInteractWithUser(self):
        hToken = win32security.OpenProcessToken(win32api.GetCurrentProcess(),
                                                win32con.TOKEN_DUPLICATE | win32con.TOKEN_ADJUST_DEFAULT |
                                                win32con.TOKEN_QUERY | win32con.TOKEN_ASSIGN_PRIMARY)
        hNewToken = win32security.DuplicateTokenEx(hToken, win32security.SecurityImpersonation,
                                                   win32security.TOKEN_ALL_ACCESS, win32security.TokenPrimary)
        # Create sid level
        sessionId = win32ts.WTSGetActiveConsoleSessionId()
        win32security.SetTokenInformation(hNewToken, win32security.TokenSessionId, sessionId)
        priority = win32con.NORMAL_PRIORITY_CLASS | win32con.CREATE_NEW_CONSOLE
        startup = win32process.STARTUPINFO()
        handle, thread_id, pid, tid = win32process.CreateProcessAsUser(hNewToken, self.filePath,
                                                                       None, None, None, False, priority, None, None,
                                                                       startup)
        # with open('C:\\svc\\TestService.log', 'a') as f:
        #     f.write("-----------------\n%s\n%s\n%s\n%s\n%\n---------------" % (handle, thread_id, pid, tid,self.filePath))

    def SvcDoRun(self):
        rc = None
        i = 0
        while rc != win32event.WAIT_OBJECT_0:
            i += 1
            if i == 5:
                self.runProcessInteractWithUser()
            with open('C:\\svc\\TestService.log', 'a') as f:
                f.write('test service stopped...{index}{path}\n'.format(index=i, path=self.filePath))
            rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)
        with open('C:\\svc\\TestService.log', 'a') as f:
            f.write('test service stopped...{index}{path}\n'.format(index=i, path=self.filePath))
    # def SvcDoRun(self):
    #     rc = None
    #     while rc != win32event.WAIT_OBJECT_0:
    #         with open('d:\\TestService.log', 'a') as f:
    #             f.write('test service running...\n')
    #         rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TestService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TestService)
