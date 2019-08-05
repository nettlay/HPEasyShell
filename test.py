import os
import uiautomation

print(uiautomation.ButtonControl(AutomationId="UserSettingsButton").IsOffScreen)