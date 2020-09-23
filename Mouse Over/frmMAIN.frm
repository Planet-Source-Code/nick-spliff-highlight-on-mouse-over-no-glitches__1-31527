VERSION 5.00
Begin VB.Form frmMAIN 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mouse Over Things"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMOUSEOVER 
      Interval        =   1
      Left            =   4200
      Top             =   4440
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Step 4: The Result"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Step 3 : How Can We Correct This Problem?"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3165
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Step 2 : Why Does This Happen?"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2400
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Step 1: Understand The Glitch."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label Label4 
      Caption         =   $"frmMAIN.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   4695
   End
   Begin VB.Label lblEMAIL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Email: spliff@wideboys.co.uk"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   4560
      Width           =   2085
   End
   Begin VB.Label lblWEBSITE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Website: http://www.spliff.wideboys.co.uk"
      Height          =   195
      Left            =   945
      TabIndex        =   3
      Top             =   4920
      Width           =   3045
   End
   Begin VB.Label lblEXAMPLE 
      Caption         =   $"frmMAIN.frx":009F
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   $"frmMAIN.frx":014D
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMAIN.frx":01DB
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   4815
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEXAMPLE.ForeColor = vbBlack
    lblEXAMPLE.FontUnderline = False
End Sub

Private Sub Label3_Click()

End Sub

Private Sub lblEXAMPLE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEXAMPLE.ForeColor = vbBlue
    lblEXAMPLE.FontUnderline = True
End Sub

Private Sub tmrMOUSEOVER_Timer()
    '### FOR THE EMAIL LABEL
    If MouseHasLeftObject(frmMAIN, lblEMAIL) = False Then
        'By adding this IF statement, we stop updating if it is already like it.
        If lblEMAIL.FontUnderline = False Then
            lblEMAIL.ForeColor = vbBlue
            lblEMAIL.FontUnderline = True
        End If
    End If
    
    If MouseHasLeftObject(frmMAIN, lblEMAIL) = True Then
        'By adding this IF statement, we stop updating if it is already like it.
        If lblEMAIL.FontUnderline = True Then
            lblEMAIL.ForeColor = vbBlack
            lblEMAIL.FontUnderline = False
         End If
    End If
    
    '### FOR THE WEBSITE LABEL
    If MouseHasLeftObject(frmMAIN, lblWEBSITE) = False Then
        'By adding this IF statement, we stop updating if it is already like it.
        If lblWEBSITE.FontUnderline = False Then
            lblWEBSITE.ForeColor = vbBlue
            lblWEBSITE.FontUnderline = True
        End If
    End If
    
    If MouseHasLeftObject(frmMAIN, lblWEBSITE) = True Then
        'By adding this IF statement, we stop updating if it is already like it.
        If lblWEBSITE.FontUnderline = True Then
            lblWEBSITE.ForeColor = vbBlack
            lblWEBSITE.FontUnderline = False
         End If
    End If
End Sub
