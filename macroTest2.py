import win32com.client as win32

xl = win32.gencache.EnsureDispatch('Excel.Application')
xl.Visible = True
ss = xl.Workbooks.Add()
xlmodule = ss.VBProject.VBComponents.Add(1)
xlmodule.Name = 'testing123'
code = '''sub TestMacro()
    msgbox "Testing 1 2 3"
    end sub'''
xlmodule.CodeModule.AddFromString(code)
ss.Application.Run('testing123.TestMacro')

