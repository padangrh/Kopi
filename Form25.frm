VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form form_Rekap 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Rekap"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form25"
   ScaleHeight     =   5670
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btn_close 
      BackColor       =   &H0080C0FF&
      Caption         =   "Close"
      Height          =   495
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton btn_PrintBon 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cetak Struk"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   1455
   End
   Begin MSComctlLib.ListView lv_jual 
      Height          =   3855
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   6800
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
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Barang"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Harga"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Jumlah"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.Label lbl_Kode 
      BackStyle       =   0  'Transparent
      Caption         =   "Nomor Faktur : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "form_Rekap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kode_Faktur As String
Public temp_Total As String

Public Sub fill_List()
    Dim rstbjual As ADODB.Recordset
    
    Set rstbjual = con.Execute("Select * from tbjual where nobukti = '" & kode_Faktur & "'")
    
    If Not rstbjual.EOF Then
        rstbjual.MoveFirst
        
        Do While Not rstbjual.EOF
            Dim mitem As ListItem
            Set mitem = lv_jual.ListItems.Add(, , rstbjual!kode)
            mitem.SubItems(1) = rstbjual!nama_barang
            mitem.SubItems(2) = rstbjual!harga_jual
            mitem.SubItems(3) = rstbjual!jumlah_jual
            mitem.SubItems(4) = rstbjual!harga_jual * rstbjual!jumlah_jual
            rstbjual.MoveNext
        Loop
    End If
    
    lbl_Kode.Caption = lbl_Kode.Caption & kode_Faktur
    
End Sub

Private Sub btn_close_Click()
    Unload Me
End Sub

Private Sub btn_PrintBon_Click()
    Form_Print.Show
    Form_Print.Init kode_Faktur, temp_Total, False
    
End Sub

Private Sub Form_unload(cancel As Integer)
    Unload Form_Print
End Sub
