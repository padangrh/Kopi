VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Kategori 
   Caption         =   "Kategori"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_Browse 
      Caption         =   "Browse"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   2520
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btn_Reset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txt_Urutan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   3240
      Width           =   2895
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   915
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   1614
      ButtonWidth     =   3043
      ButtonHeight    =   1455
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.ListBox list_kategori 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox txt_Kategori 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txt_Gambar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   2520
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   49
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Kategori.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Kategori.frx":0D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Kategori.frx":1DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Kategori.frx":2CF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Default = 99 (smallest displayed first)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   11
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label lbl_Urutan 
      BackStyle       =   0  'Transparent
      Caption         =   "Urutan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori Manager"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lbl_Gambar 
      BackStyle       =   0  'Transparent
      Caption         =   "Gambar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lbl_Kategori 
      BackStyle       =   0  'Transparent
      Caption         =   "Kategori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "Form_Kategori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rsKategori As ADODB.Recordset

Private Sub Command1_Click()
    cd.InitDir = App.Path
    cd.ShowOpen
End Sub

Private Sub btn_Browse_Click()
    OpenShowFile
End Sub

Private Sub btn_Reset_Click()
    txt_Kategori.Text = ""
    txt_Gambar.Text = ""
    txt_Urutan.Text = ""
End Sub

Private Sub Form_Load()
    reload
End Sub


Private Sub reload()
    list_kategori.Clear
    Set rsKategori = con.Execute("select * from tbkategori order by urutan asc")
    If rsKategori.EOF Then Exit Sub
    
    rsKategori.MoveFirst
    Do While Not rsKategori.EOF
        list_kategori.AddItem (rsKategori!kode)
        rsKategori.MoveNext
    Loop
End Sub

Private Function getKategori(kategori_id As String) As Boolean
    Dim found As Boolean
    found = False
    rsKategori.MoveFirst
    Do While Not rsKategori.EOF
        If rsKategori!kode = kategori_id Then
            found = True
            Exit Do
        End If
        rsKategori.MoveNext
    Loop
    
    getKategori = found
End Function

Private Sub list_kategori_Click()
    If getKategori(list_kategori.Text) Then
        txt_Kategori = rsKategori!kode
        txt_Gambar = rsKategori!gambar
        txt_Urutan = rsKategori!urutan
    Else
        reset
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: tambah
        Case 2: ubah
        Case 3: hapus
        Case 4: keluar
    End Select
End Sub

Private Sub tambah()
    If txt_Kategori = "" Then
        MsgBox "Kategori harus diisi"
        Exit Sub
    End If
    
    If getKategori(txt_Kategori) Then
        MsgBox "Kategori sudah ada"
        Exit Sub
    End If
    
    If txt_Urutan.Text = "" Or IsNumeric(txt_Urutan.Text) = False Or Len(txt_Urutan.Text) > 6 Then
        txt_Urutan.Text = 99
    End If
    
    con.Execute ("insert into tbkategori values('" & txt_Kategori & "', '" & txt_Gambar & "', '" & txt_Urutan & "')")
    reload
    reset
End Sub

Private Sub ubah()
    If txt_Kategori = "" Then
        MsgBox "Kategori harus diisi"
        Exit Sub
    End If
    
    If Not getKategori(txt_Kategori) Then
        MsgBox "Kategori tidak ditemukan"
        Exit Sub
    End If
    
    If txt_Urutan.Text = "" Or IsNumeric(txt_Urutan.Text) = False Then
        txt_Urutan.Text = 99
    End If
    
    con.Execute ("update tbkategori set gambar = '" & txt_Gambar & "', urutan = '" & txt_Urutan & "' where kode = '" & txt_Kategori & "'")
    reload
    reset
End Sub

Private Sub hapus()
    con.Execute ("delete from tbkategori where kode = '" & txt_Kategori & "'")
    reload
    reset
End Sub

Private Sub keluar()
    Unload Me
End Sub

Private Sub OpenShowFile()
    On Error GoTo OpenShowFileError
    Dim FName As String, FNumb As Integer
    Dim FileContents As String
    cd.CancelError = True
    'cd.Filter = "Text Files *.txt|*.txt"
    cd.InitDir = App.Path
    cd.FileName = vbNullString
    cd.ShowOpen
    If cd.FileName = vbNullString Then Exit Sub
    
    txt_Gambar.Text = cd.FileTitle
    
'    FName = cd.FileName
'    FNumb = FreeFile
'    Open FName For Input As #FNumb
'    FileContents = Input(FileLen(FName), #FNumb)
'    Close #FNumb
'    Text1.Text = FileContents
    Exit Sub
OpenShowFileError:
    If Err.Number = 32755 Then Exit Sub 'user pressed cancel
End Sub

