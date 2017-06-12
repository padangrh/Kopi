Attribute VB_Name = "Module1"
Public fMainForm As FrmMain
Public con As New ADODB.Connection
Public rsSupplier As ADODB.Recordset
Public status
Public username As String
Public Setting_Object As Object

' **********************************************
' Posiflex usbpd.dll DLL
' **********************************************
Public Declare Function WritePD _
    Lib "usbpd.dll" _
    (ByVal data As String, ByVal Length As Long) _
As Long

Public Declare Function WritePD80 _
    Lib "usbpd.dll" Alias "WritePD" _
    (ByRef data As Any, ByVal Length As Long) _
As Long

Public Declare Function PdState _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function OpenUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function CloseUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Declare Sub Sleep Lib "kernel32" _
   (ByVal dwMilliseconds As Long)
   
Public Sub connect()
    
    
'    con.ConnectionString = "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=foodcourt1"
    con.ConnectionString = "Provider=MSDASQL.1;Password=" & Setting_Object.Item("DB_Pw") & ";Persist Security Info=True;User ID=" & Setting_Object.Item("DB_Id") & ";Data Source=" & Setting_Object.Item("DB_Name")
    con.Open
End Sub

Public Function getSupplier(kode As String) As Boolean
    If rsSupplier Is Nothing Then
        Set rsSupplier = con.Execute("select * from tbsuplier")
    'Else
        'rsSupplier.MoveFirst
    End If
    
    Dim found As Boolean
    found = False
    If Not rsSupplier.EOF Then
        rsSupplier.MoveFirst
        Do While Not rsSupplier.EOF
            If kode = rsSupplier!kdsuplier Then
                found = True
                Exit Do
            End If
            rsSupplier.MoveNext
        Loop
    End If
    getSupplier = found
End Function

Public Function priceToNum(price As String) As Long
    price = Replace(price, ".", "")
    price = Replace(price, ",", "")
    priceToNum = Val(price)
End Function

Public Function isMaster() As Boolean
    isMaster = (status = "Master")
End Function

Public Function isSPV() As Boolean
    isSPV = (status = "Supervisor")
End Function

Sub Main()
    Set fMainForm = New FrmMain
    fMainForm.Show
End Sub


'Public Function NewWindowProc(ByVal hwnd As Long, ByVal msg _
'    As Long, ByVal wParam As Long, ByVal lParam As Long) As _
'    Long
'Const WM_NCDESTROY = &H82
'Const WM_NOTIFY = &H4E
'Const LVN_FIRST = -100&
'Const LVN_BEGINDRAG = (LVN_FIRST - 9)
'
'Dim nm_hdr As NMHDR
'
'    ' If we're being destroyed,
'    ' restore the original WindowProc.
'    If msg = WM_NCDESTROY Then
'        SetWindowLong _
'            hwnd, GWL_WNDPROC, _
'            OldWindowProc
'    ElseIf msg = WM_NOTIFY Then
'        ' Copy info into the NMHDR structure.
'        CopyMemory nm_hdr, ByVal lParam, Len(nm_hdr)
'
'        ' See if this is the start of a drag.
'        If nm_hdr.Code = LVN_BEGINDRAG Then
'            ' A drag is beginning. Ignore this event.
'            ' Indicate we have handled this.
'            NewWindowProc = 1
'            ' Do nothing else.
'            Exit Function
'        End If
'    End If
'
'    NewWindowProc = CallWindowProc( _
'        OldWindowProc, hwnd, msg, wParam, _
'        lParam)
'End Function
