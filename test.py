import uiautomation


t = uiautomation.ButtonControl(AutomationId='DisplayTilesToggle').GetTogglePattern()
print(t.ToggleState)

# uiautomation.ButtonControl(AutomationId='buttonExport').GetTogglePattern().Toggle()
