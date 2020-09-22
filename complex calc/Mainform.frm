VERSION 5.00
Begin VB.Form Mainform 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complex Calculator"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ANS"
      Height          =   375
      Index           =   5
      Left            =   1560
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Mainform.frx":0000
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SQR"
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "/"
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "*"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "-"
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "0"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Text            =   "0"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "i"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   14
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "i"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   12
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   11
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ans As Complex

Private Sub Command1_Click(Index As Integer)
Dim t1 As Complex, t2 As Complex
On Error GoTo err:
t1 = MakeComplex(Text1, Text2)
t2 = MakeComplex(Text3, Text4)

Select Case Index
    Case 0 'add
        ans = C_ADD(t1, t2)
    Case 1 'subtract
        ans = C_sub(t1, t2)
    Case 2 'multiply
        ans = Cmult(t1, t2)
    Case 3 'devide
        ans = CDev(t1, t2)
    Case 4 ' square root
        ans = C_SQR(t1)
    Case 5 ' move the answer to term 1's position
        Text1 = ans.Real
        Text2 = ans.Imag
End Select
Text5 = GenerateString(ans)
Exit Sub
err:
MsgBox "Error" & vbNewLine & Error$(err)

End Sub

Private Sub Label2_Click(Index As Integer)

End Sub
