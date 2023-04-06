If Not IsObject(application) Then

   Set SapGuiAuto  = GetObject("SAPGUI")

   Set application = SapGuiAuto.GetScriptingEngine

End If

If Not IsObject(connection) Then

   Set connection = application.Children(0)

End If

If Not IsObject(session) Then

   Set session    = connection.Children(0)

End If

If IsObject(WScript) Then

   WScript.ConnectObject session,     "on"

   WScript.ConnectObject application, "on"

End If

Dim intRow

Set objExcel = CreateObject("Excel.Application")

Set objWorkbook = objExcel.Workbooks.Open("C:\Users\username\Desktop\filename.xlsx") ' file direction

intRow = 2

Do Until objExcel.Cells(intRow,1).Value = ""

session.findById("wnd[0]").maximize

session.findById("wnd[0]/tbar[0]/okcd").text = "/npfcg"

session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtAGR_NAME_NEU").text = objExcel.Cells(intRow,1).Value    ' Role name in column 1 (from which the copy is made)

session.findById("wnd[0]/usr/ctxtAGR_NAME_NEU").caretPosition = 30

session.findById("wnd[0]/tbar[1]/btn[23]").press

session.findById("wnd[1]/usr/ctxtP_DEST").text = objExcel.Cells(intRow,2).Value    ' Column 2 role name (new role)

session.findById("wnd[1]/usr/ctxtP_DEST").setFocus

session.findById("wnd[1]/usr/ctxtP_DEST").caretPosition = 2

session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[0]/btn[3]").press

intRow = intRow + 1

Loop

objExcel.Quit
