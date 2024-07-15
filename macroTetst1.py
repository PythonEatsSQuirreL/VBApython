import win32com.client as win32

import comtypes, comtypes.client

xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.Visible = True
ss = xl.Workbooks.Add()
sh = ss.ActiveSheet

xlmodule = ss.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule

sCode = '''sub VBAMacro()
       msgbox "VBA Macro called"
      end sub'''

xlmodule.CodeModule.AddFromString(sCode)