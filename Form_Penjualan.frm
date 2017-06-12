VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form_Penjualan 
   BackColor       =   &H0000C000&
   Caption         =   "Transaksi Penjualan"
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20250
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form11"
   ScaleHeight     =   13046.83
   ScaleMode       =   0  'User
   ScaleWidth      =   24147.85
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11760
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   100
      ImageHeight     =   100
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Penjualan.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV_Barang 
      Height          =   2055
      Left            =   240
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama"
         Object.Width           =   2130
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kode"
         Object.Width           =   2130
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Kategori"
         Object.Width           =   2130
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Harga"
         Object.Width           =   2130
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "KodeSupplier"
         Object.Width           =   2130
      EndProperty
   End
   Begin MSComctlLib.ListView LV_Kategori 
      Height          =   2055
      Left            =   240
      TabIndex        =   20
      Top             =   1320
      Width           =   19455
      _ExtentX        =   34316
      _ExtentY        =   3625
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kategori"
         Object.Width           =   2130
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ID"
         Object.Width           =   2130
      EndProperty
   End
   Begin MSComctlLib.ListView list_nama 
      Height          =   2295
      Left            =   4440
      TabIndex        =   18
      Top             =   5040
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode"
         Object.Width           =   2976
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama"
         Object.Width           =   7440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Harga"
         Object.Width           =   2976
      EndProperty
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   615
      Left            =   17160
      TabIndex        =   15
      Top             =   4440
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   1085
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt_jumlah"
      BuddyDispid     =   196609
      OrigLeft        =   19062
      OrigTop         =   2863
      OrigRight       =   19364
      OrigBottom      =   3596
      Max             =   9999
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   12360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.ListView lv_jual 
      Height          =   4575
      Left            =   360
      TabIndex        =   13
      Top             =   5520
      Width           =   19275
      _ExtentX        =   33999
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Barang"
         Object.Width           =   4464
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   14879
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Harga"
         Object.Width           =   4464
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Jumlah"
         Object.Width           =   4464
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   5208
      EndProperty
   End
   Begin VB.TextBox txt_jumlah 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   614
      Left            =   15960
      TabIndex        =   12
      Text            =   "12345678901234"
      Top             =   4440
      Width           =   1239
   End
   Begin VB.TextBox txt_harga 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   13200
      TabIndex        =   11
      Text            =   "12345678901234"
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txt_nama 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4440
      TabIndex        =   8
      Text            =   "12345678901234"
      Top             =   4440
      Width           =   8415
   End
   Begin VB.TextBox txt_kode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   600
      TabIndex        =   5
      Text            =   "12345678901234"
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox txt_total 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15360
      TabIndex        =   2
      Text            =   "00.000.000"
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Faktur:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   19
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackColor       =   &H0080FF80&
      Caption         =   "[F1 -> Print] [F2 / ESC -> Hide]  [Delete -> Hapus 1 baris]  [Shift + Delete -> Hapus Semua]  [F4 -> Tutup]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   10320
      Width           =   15015
   End
   Begin VB.Label Label8 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15960
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Harga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13200
      TabIndex        =   9
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Barang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label lbl_user 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Richard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Kasir:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Rp."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   14280
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label lbl_faktur 
      BackColor       =   &H0080FF80&
      BackStyle       =   0  'Transparent
      Caption         =   "A123456"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackColor       =   &H0080FF80&
      Height          =   1695
      Left            =   360
      TabIndex        =   16
      Top             =   3600
      Width           =   19275
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FF80&
      Height          =   1095
      Left            =   360
      TabIndex        =   17
      Top             =   120
      Width           =   19275
   End
End
Attribute VB_Name = "Form_Penjualan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsbarang As ADODB.Recordset
Dim txt_nama_toggle As Boolean
'Dim flag_LV_Barang As Boolean
Dim icon_Index As Integer

Private Sub Form_Load()
    icon_Index = 0
    lbl_user = username
    txt_total = 0
    'flag_LV_Barang = False
    kosongkan
    
    'load image
    Dim rsKat As ADODB.Recordset
    Dim n As Integer
    n = 1
    Set rsKat = con.Execute("Select * from tbkategori order by urutan asc")
    
    rsKat.MoveFirst
    
    Do While Not rsKat.EOF
        n = n + 1
        Set fso = New FileSystemObject
        If fso.FileExists(App.Path & "\" & rsKat!gambar) Then
            Call ImageList1.ListImages.Add(n, rsKat!kode, LoadPicture(App.Path & "\" & rsKat!gambar))
        Else
            Call ImageList1.ListImages.Add(n, rsKat!kode, ImageList1.ListImages(1).Picture)
        End If
        rsKat.MoveNext
    Loop
    
    generate_Kategori
    txt_nama_toggle = False
    Set rsbarang = con.Execute("select * from tbbarang")
        
    Dim namafile, file_data, huruf As String
    Dim angka As Long
    namafile = App.Path & "\faktur.txt"
    Open namafile For Input As #1
    While Not EOF(1)
        Input #1, data
        file_data = data
        huruf = Left(file_data, 1)
        angka = Val(Mid(file_data, 2, 20))
        lbl_faktur = huruf + CStr(angka + 1)
    Wend
Close #1

    On Error GoTo ErrorFound
    MSComm1.CommPort = 3
    MSComm1.Settings = "9600,N,8,1"
    MSComm1.PortOpen = True
    
    If MSComm1.PortOpen Then
        MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
        MSComm1.Output = "Selamat Datang      Kasir: " + username
    End If
ErrorFound:
        'nothing happens
    On Error GoTo 0

    reload_List
    
      'test
'    OldWindowProc = SetWindowLong( _
'        hwnd, GWL_WNDPROC, _
'        AddressOf NewWindowProc)

End Sub

Private Sub Form_unload(cancel As Integer)
     If MSComm1.PortOpen = True Then
      Do While MSComm1.OutBufferCount > 0
          DoEvents
       Loop
       MSComm1.PortOpen = False
    End If
End Sub

Private Sub Form_KeyDown(key As Integer, Shift As Integer)
    If key = 112 Then
        If lv_jual.ListItems.count > 0 Then
            Form_Print.Show
            Form_Print.Init lbl_faktur, txt_total, True
            Me.Enabled = False
            If MSComm1.PortOpen Then
                MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
                MSComm1.Output = "Total Belanja:      " + txt_total.Text
            End If
            LV_Barang.Visible = False
        Else
            MsgBox "Faktur masih kosong"
        End If
    End If
    If key = 113 Or key = 27 Then
        
        'flag_LV_Barang = True
        LV_Barang.Visible = False
        txt_kode.SetFocus
    End If
    
    
    If lv_jual.ListItems.count <> 0 Then
        If key = 46 Then
            If Shift = 1 Then
                txt_total = "0"
                lv_jual.ListItems.Clear
            Else
                txt_total = Format(priceToNum(txt_total) - priceToNum(lv_jual.SelectedItem.SubItems(4)), "###,###,##0")
                lv_jual.ListItems.Remove (lv_jual.SelectedItem.index)
                'set barang terakhir untuk dihapus
                If lv_jual.ListItems.count <> 0 Then lv_jual.ListItems(lv_jual.ListItems.count).Selected = True
            End If
        End If
    End If
    
    If key = 115 Then
        If MsgBox("Tutup form transaksi?", vbYesNo) = vbYes Then
            Unload Me
        End If
    End If
End Sub

Private Sub kosongkan()
    txt_kode.Text = ""
    txt_nama.Text = ""
    txt_harga.Text = ""
    txt_jumlah.Text = 1
    list_nama.Visible = False
End Sub

Private Sub list_nama_lostfocus()
    list_nama.Visible = False
End Sub

Private Sub list_nama_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        list_nama_DblClick
    End If
End Sub

Private Sub LV_Barang_ItemClick(ByVal item As MSComctlLib.ListItem)
        
        Dim found As Boolean
        Dim i As Integer
        found = False
        i = 1
        
        Do While i <= lv_jual.ListItems.count
            
            If lv_jual.ListItems(i).Text = item.SubItems(1) Then
                found = True
                lv_jual.ListItems(i).SubItems(3) = Val(lv_jual.ListItems(i).SubItems(3)) + 1
                lv_jual.ListItems(i).SubItems(4) = priceToNum(lv_jual.ListItems(i).SubItems(4)) + priceToNum(item.SubItems(3))
                lv_jual.ListItems(i).SubItems(4) = Format(lv_jual.ListItems(i).SubItems(4), "###,###,##0")
                Exit Do
            End If
            i = i + 1
        Loop
        
        Dim subtotal As String
        subtotal = Format(item.SubItems(3), "###,###,##0")
        
        If found = False Then
            Dim Item1 As ListItem
            Set Item1 = lv_jual.ListItems.Add(, , item.SubItems(1))
            Item1.SubItems(1) = item.Text
            Item1.SubItems(2) = Format(item.SubItems(3), "###,###,##0")
            Item1.SubItems(3) = "1"
            Item1.SubItems(4) = subtotal
        End If
        
        txt_total.Text = Format(priceToNum(txt_total) + priceToNum(subtotal), "###,###,##0")
        
        If MSComm1.PortOpen Then
            MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
            Dim baris1, baris2 As String
            baris1 = txt_jumlah.Text + " " + txt_nama.Text
            If Len(baris1) < 20 Then
               Do While (Len(baris1) < 20)
                baris1 = baris1 + " "
               Loop
            Else
                baris1 = Left$(baris1, 20)
            End If
            
            MSComm1.Output = baris1
              
            Dim spaces As Integer
            spaces = 20 - (Len(subtotal) + Len(txt_total.Text) + 2)
            Do While (Len(baris2) < spaces)
                baris2 = baris2 + " "
            Loop
            baris2 = subtotal + baris2 + "(" + txt_total.Text + ")"
            MSComm1.Output = baris2
        End If
        
        kosongkan
        If lv_jual.ListItems.count <> 0 Then lv_jual.ListItems(lv_jual.ListItems.count).Selected = True
        txt_kode.SetFocus
End Sub

Private Sub LV_Kategori_ItemClick(ByVal item As MSComctlLib.ListItem)
    'If flag_LV_Barang = False Then
    If CInt(item.index) <> icon_Index Then
        LV_Barang.Visible = True
        'MsgBox icon_Index & "X" & Item.index
'        Select Case item.index
'        Case 1
'            icon_Index = 1
'            Call generate_ListIcons(1, "Hot Drinks")
'        Case 2
'            icon_Index = 2
'            Call generate_ListIcons(2, "Flavored Latte")
'        Case 3
'            icon_Index = 3
'            Call generate_ListIcons(3, "Ice Blended")
'        Case 4
'            icon_Index = 4
'            Call generate_ListIcons(4, "Tea")
'        End Select
        icon_Index = item.index
        Call generate_ListIcons(item.index, item.Text)
    Else
        icon_Index = 0
        LV_Barang.Visible = False
    End If
'    Else
'        flag_LV_Barang = False
'    End If
    'generate_Kategori
End Sub

Private Sub LV_Kategori_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And LV_Barang.Visible = True Then
        LV_Barang.SetFocus
    End If
End Sub

Private Sub txt_nama_Change()
    
    If txt_nama.Text <> "" And txt_nama_toggle = False Then
        list_nama.Visible = True
        reload_List
    Else
        list_nama.Visible = False
        txt_nama_toggle = False
    End If
End Sub

Private Sub txt_nama_LostFocus()
    If Not Me.ActiveControl Is Nothing Then
        If Not Me.ActiveControl.Name = "list_nama" Then
            list_nama.Visible = False
        End If
    End If
End Sub

Private Sub list_nama_DblClick()
    If getItemByID(list_nama.SelectedItem.Text) Then
        txt_kode.Text = rsbarang!kode
        txt_nama.Text = rsbarang!nama
        txt_harga.Text = Format(rsbarang!harga_jual, "###,###,##0")
        list_nama.Visible = False
        txt_jumlah.SetFocus
        txt_jumlah.SelLength = Len(txt_jumlah.Text)
    End If
End Sub

Private Sub txt_jumlah_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        If Len(txt_jumlah) > 4 Then
            txt_jumlah = ""
            Exit Sub
        End If
    
        If txt_harga = "" Then
            MsgBox "Barang tidak valid"
            Exit Sub
        End If
        
        If Val(txt_jumlah.Text) < 1 Then
            MsgBox "Jumlah tidak valid"
            Exit Sub
        End If
        
        Dim found As Boolean
        Dim i As Integer
        found = False
        i = 1
        
        Do While i <= lv_jual.ListItems.count
            If lv_jual.ListItems(i).Text = rsbarang!kode Then
                found = True
                lv_jual.ListItems(i).SubItems(3) = Val(lv_jual.ListItems(i).SubItems(3)) + Val(txt_jumlah.Text)
                lv_jual.ListItems(i).SubItems(4) = priceToNum(lv_jual.ListItems(i).SubItems(4)) + Val(txt_jumlah.Text) * priceToNum(txt_harga)
                lv_jual.ListItems(i).SubItems(4) = Format(lv_jual.ListItems(i).SubItems(4), "###,###,##0")
                Exit Do
            End If
            i = i + 1
        Loop
        
        Dim subtotal As String
        subtotal = Format(rsbarang!harga_jual * Val(txt_jumlah), "###,###,##0")
        
        If found = False Then
            Dim item As ListItem
            Set item = lv_jual.ListItems.Add(, , rsbarang!kode)
            item.SubItems(1) = rsbarang!nama
            item.SubItems(2) = Format(rsbarang!harga_jual, "###,###,##0")
            item.SubItems(3) = txt_jumlah.Text
            item.SubItems(4) = subtotal
        End If
        
        txt_total.Text = Format(priceToNum(txt_total) + priceToNum(subtotal), "###,###,##0")
        
        If MSComm1.PortOpen Then
            MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
            Dim baris1, baris2 As String
            baris1 = txt_jumlah.Text + " " + txt_nama.Text
            If Len(baris1) < 20 Then
               Do While (Len(baris1) < 20)
                baris1 = baris1 + " "
               Loop
            Else
                baris1 = Left$(baris1, 20)
            End If
            
            MSComm1.Output = baris1
              
            Dim spaces As Integer
            spaces = 20 - (Len(subtotal) + Len(txt_total.Text) + 2)
            Do While (Len(baris2) < spaces)
                baris2 = baris2 + " "
            Loop
            baris2 = subtotal + baris2 + "(" + txt_total.Text + ")"
            MSComm1.Output = baris2
        End If
        
        kosongkan
        reload_List
        If lv_jual.ListItems.count <> 0 Then lv_jual.ListItems(lv_jual.ListItems.count).Selected = True
        txt_kode.SetFocus
    End If
End Sub

Private Sub txt_kode_KeyDown(key As Integer, Shift As Integer)
    If key = 13 Then
        txt_nama_toggle = True
        Dim kode As String
        kode = Trim(txt_kode.Text)
        If getItemByID(kode) Then
            txt_nama.Text = rsbarang!nama
            txt_harga.Text = Format(rsbarang!harga_jual, "###,###,##0")
            txt_jumlah.SetFocus
            txt_jumlah.SelLength = Len(txt_jumlah.Text)
        Else
            MsgBox ("Kode ini tidak terdaftar")
        End If
    ElseIf Len(txt_nama) > 0 Then
        txt_nama = ""
        txt_harga = ""
    End If
End Sub

Private Function getItemByID(kode As String) As Boolean
    rsbarang.MoveFirst
    Do While Not rsbarang.EOF
        If rsbarang!kode = kode Then
            getItemByID = True
            Exit Function
        End If
        rsbarang.MoveNext
    Loop
    getItemByID = False
End Function

Private Sub txt_nama_KeyDown(key As Integer, Shift As Integer)
    If key = 40 Then
        list_nama.Visible = True
        list_nama.SetFocus
    ElseIf key = 13 And list_nama.Visible = True Then
        list_nama.SetFocus
    End If
    
    
End Sub

Public Sub nextFaktur()
    Dim namafile, huruf As String
    Dim angka As Long
    Me.Enabled = True
    huruf = Left(lbl_faktur, 1)
    angka = Val(Mid(lbl_faktur, 2, 20))
    
    namafile = App.Path & "\faktur.txt"
    Open namafile For Output As #1
    Print #1, lbl_faktur
    Close #1
    
    lbl_faktur = huruf + CStr(angka + 1)
    lv_jual.ListItems.Clear
    txt_total = "0"
    kosongkan
    txt_kode.SetFocus
    Form_List_Jual.refreshlist
    
    If MSComm1.PortOpen Then
        MSComm1.Output = Chr$(&H1B) + Chr$(&H49) + Chr$(&HC)
        MSComm1.Output = "Selamat Datang      Kasir: " + username
    End If
End Sub

Public Sub reload_List()
'pindahan generate list barang
    list_nama.ListItems.Clear
    'list_nama.Visible = True
    Dim rsFilter As ADODB.Recordset
    Set rsFilter = con.Execute("select * from tbbarang where nama like '%" & txt_nama.Text & "%'")
    
    If rsFilter.EOF Then
        list_nama.Visible = False
        Exit Sub
    End If
    
    rsFilter.MoveFirst
    Do While Not rsFilter.EOF
        Dim mitem As ListItem
        Set mitem = list_nama.ListItems.Add(, , rsFilter!kode)
        mitem.SubItems(1) = rsFilter!nama
        mitem.SubItems(2) = "Rp. " + Format(rsFilter!harga_jual, "###,###,##0")
        rsFilter.MoveNext
    Loop
    
    Set rsFilter = Nothing
'end pindahan list barang
End Sub

Public Sub generate_Kategori()
    
    LV_Kategori.ListItems.Clear
    'list_nama.Visible = True
    Dim rsKategori As ADODB.Recordset
    Set rsKategori = con.Execute("select * from tbkategori order by urutan asc")
    
    If rsKategori.EOF Then
        LV_Kategori.Visible = False
        Exit Sub
    End If
    Dim i As Integer
    i = 0
    rsKategori.MoveFirst
    Do While Not rsKategori.EOF
        If i < 5 Then i = i + 1
        Dim mitem As ListItem
        Set mitem = LV_Kategori.ListItems.Add(, , rsKategori!kode, i + 1)
        rsKategori.MoveNext
    Loop
    
    Set rsKategori = Nothing
End Sub

Public Sub generate_ListIcons(i As Integer, temp_kategori As String)
    LV_Barang.ListItems.Clear
    
    Dim rstemp As ADODB.Recordset
    Dim counter_subitem As Integer
    Set rstemp = con.Execute("select * from tbbarang where kategori = '" & temp_kategori & "' order by urutan")
    counter_subitem = 0
    
    
    If Not rstemp.EOF Then
        Dim mitem As ListItem
        Do While Not rstemp.EOF
            Set mitem = LV_Barang.ListItems.Add(, , rstemp!nama, i + 1)
            mitem.SubItems(1) = rstemp!kode
            mitem.SubItems(2) = rstemp!kategori
            mitem.SubItems(3) = rstemp!harga_jual
            mitem.SubItems(4) = rstemp!kdsuplier
            rstemp.MoveNext
            counter_subitem = counter_subitem + 1
        Loop
    End If
    
    If counter_subitem > 18 Then
        LV_Barang.Height = 7356
    ElseIf counter_subitem > 9 Then
        LV_Barang.Height = 4903
    Else
        LV_Barang.Height = 2451.873
    End If
    
    Set rstemp = Nothing
End Sub
