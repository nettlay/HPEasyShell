import uiautomation

wnd = uiautomation.WindowControl(RegexName='Administrator: Command Prompt.*')
print(wnd)
wnd.GetWindowPattern().Close()
