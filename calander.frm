VERSION 5.00
Begin VB.Form frmCalander 
   Appearance      =   0  'Flat
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Nehaj Calander"
   ClientHeight    =   4500
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "calander.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "calander.frx":0442
   ScaleHeight     =   4500
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbMonth 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "calander.frx":4FFA
      Left            =   210
      List            =   "calander.frx":502E
      TabIndex        =   37
      Text            =   "Select Month"
      Top             =   1225
      Width           =   1775
   End
   Begin VB.VScrollBar scrlYear 
      Height          =   250
      Left            =   2740
      TabIndex        =   1
      Top             =   1260
      Width           =   255
   End
   Begin VB.TextBox txtyear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   0
      Text            =   "200"
      Top             =   1225
      Width           =   975
   End
   Begin VB.Label lblToday 
      BackStyle       =   0  'Transparent
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   39
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblexit 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   4250
      TabIndex        =   38
      ToolTipText     =   "Exit"
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   1
      Left            =   810
      TabIndex        =   36
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   34
      Left            =   3140
      TabIndex        =   35
      Top             =   3800
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   33
      Left            =   2670
      TabIndex        =   34
      Top             =   3800
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   32
      Left            =   2200
      TabIndex        =   33
      Top             =   3800
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   31
      Left            =   1735
      TabIndex        =   32
      Top             =   3800
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   30
      Left            =   1275
      TabIndex        =   31
      Top             =   3800
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   29
      Left            =   840
      TabIndex        =   30
      Top             =   3795
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   28
      Left            =   360
      TabIndex        =   29
      Top             =   3800
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   27
      Left            =   3140
      TabIndex        =   28
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   26
      Left            =   2670
      TabIndex        =   27
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   25
      Left            =   2200
      TabIndex        =   26
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   24
      Left            =   1735
      TabIndex        =   25
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   23
      Left            =   1275
      TabIndex        =   24
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   22
      Left            =   810
      TabIndex        =   23
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   21
      Left            =   360
      TabIndex        =   22
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   20
      Left            =   3140
      TabIndex        =   21
      Top             =   2910
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   19
      Left            =   2670
      TabIndex        =   20
      Top             =   2910
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   18
      Left            =   2200
      TabIndex        =   19
      Top             =   2910
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   17
      Left            =   1735
      TabIndex        =   18
      Top             =   2910
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   16
      Left            =   1275
      TabIndex        =   17
      Top             =   2910
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   15
      Left            =   810
      TabIndex        =   16
      Top             =   2910
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   14
      Left            =   360
      TabIndex        =   15
      Top             =   2910
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   13
      Left            =   3140
      TabIndex        =   14
      Top             =   2490
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   12
      Left            =   2670
      TabIndex        =   13
      Top             =   2490
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   11
      Left            =   2200
      TabIndex        =   12
      Top             =   2490
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   10
      Left            =   1735
      TabIndex        =   11
      Top             =   2490
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   9
      Left            =   1275
      TabIndex        =   10
      Top             =   2490
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   8
      Left            =   810
      TabIndex        =   9
      Top             =   2490
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   7
      Left            =   360
      TabIndex        =   8
      Top             =   2490
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   6
      Left            =   3140
      TabIndex        =   7
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   5
      Left            =   2670
      TabIndex        =   6
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   4
      Left            =   2200
      TabIndex        =   5
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   3
      Left            =   1735
      TabIndex        =   4
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   2
      Left            =   1275
      TabIndex        =   3
      Top             =   2040
      Width           =   345
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   350
   End
End
Attribute VB_Name = "frmCalander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim varday As Integer
Private Sub cmbMonth_Change()

Dim ln As Integer
    
    For a = 0 To cmbMonth.ListCount - 1
        ln = Len(cmbMonth.Text)
        If UCase(cmbMonth.Text) = Left(UCase(cmbMonth.List(a)), ln) Then
            cmbMonth.ListIndex = a
            cmbMonth.SelStart = ln
            cmbMonth.SelLength = Len(cmbMonth.Text) - ln
            Exit Sub
        End If
    Next a
    
End Sub
 
Private Sub cmbMonth_Click()
    Calander
End Sub

Private Sub Form_Load()

    varday = Day(Date)
    cmbMonth.ListIndex = Month(Date) - 1
    scrlYear.Value = Year(Date)
    
End Sub



Private Sub lblDay_Click(Index As Integer)
If IsNumeric(lblDay(Index).Caption) Then
    varday = val(lblDay(Index).Caption)
    Calander
End If
End Sub

Private Sub lblexit_Click()
End
End Sub

Private Sub lblToday_Click()
    varday = Day(Date)
    cmbMonth.ListIndex = Month(Date) - 1
    scrlYear.Value = Year(Date)
    Calander
End Sub

Private Sub scrlYear_Change()
    txtyear.Text = scrlYear.Value
    
    Calander
End Sub
Private Sub Calander()
Dim n As Integer, t As Integer
n = 1

        t = Dnum(1, cmbMonth.ListIndex + 1, CLng(txtyear.Text))
    
        For a = 0 To lblDay.Count - 1
            
            lblDay(a).Caption = ""
            lblDay(a).ForeColor = &HFFF&
        
            If a >= t And a < t + cmbMonth.ItemData(cmbMonth.ListIndex) Then
                If n = varday Then lblDay(a).ForeColor = vbYellow
                 lblDay(a).Caption = n
                n = n + 1
            End If
        Next a
        
        
        If n <= cmbMonth.ItemData(cmbMonth.ListIndex) Then
            For a = 0 To cmbMonth.ItemData(cmbMonth.ListIndex) - n
                If n = varday Then lblDay(a).ForeColor = vbYellow
                 lblDay(a).Caption = n
                n = n + 1
            Next a
        End If
End Sub

Private Sub txtyear_Change()
    If IsNumeric(txtyear.Text) And val(txtyear.Text) >= 0 Then
        scrlYear.Value = val(txtyear.Text)
        Calander
    Else
        txtyear.Text = scrlYear.Value
    End If
End Sub


' Main Algorthm


Private Function Dnum(dae As Integer, mon As Integer, Yea As Long) As Integer

    Dim days(1 To 12) As Integer
    Dim val As Integer

    Yea = Abs(Yea)
    val = 0

    cmbMonth.ItemData(1) = 28


    If (Yea) Mod 4 = 0 And Yea <> 0 Then cmbMonth.ItemData(1) = 29    'if leap then feb is of 29



    val = Fix((Yea - 1) / 4)                            'add days for all leap years


    val = val + (((Yea Mod 7) * (365 Mod 7)) Mod 7)     'get all days till given year

        For a = 0 To (mon - 2)
            val = val + cmbMonth.ItemData(a)                      'add previous month's days
        Next a

    val = val + dae                                     'add current days


    val = val Mod 7                                     'now calculate the date day

    Dnum = val

cmbMonth.ItemData(1) = 28
End Function

