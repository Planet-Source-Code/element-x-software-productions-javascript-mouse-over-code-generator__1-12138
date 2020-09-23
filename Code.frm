VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "JavaScript Mouse Over Maker"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4605
   Icon            =   "Code.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Clear Fields"
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtLoadingText 
      Height          =   375
      Left            =   1680
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtMouseDown 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtNormalImage 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtMouseOver 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtCode 
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2760
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Code"
      Default         =   -1  'True
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Element-X Software"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      MouseIcon       =   "Code.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "Loading Text:"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Link URL: "
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "MouseDown Image: "
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "MouseOver Image: "
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Normal Image:"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
txtCode.Text = "<script language='JavaScript'>" & vbNewLine
txtCode.Text = txtCode.Text + "<!--" & vbNewLine
txtCode.Text = txtCode.Text + "function click396()" & vbNewLine
txtCode.Text = txtCode.Text + "{" & vbNewLine
txtCode.Text = txtCode.Text + "var URL = '" & txtURL.Text & "';" & vbNewLine
txtCode.Text = txtCode.Text + "var link = '' + URL;" & vbNewLine
txtCode.Text = txtCode.Text + "if (link != '') window.location = URL;" & vbNewLine
txtCode.Text = txtCode.Text + "}" & vbNewLine
txtCode.Text = txtCode.Text + "function mouseover396()" & vbNewLine
txtCode.Text = txtCode.Text + "{" & vbNewLine
txtCode.Text = txtCode.Text + "if (document.images) {document.img396.src = MouseOverImage396.src;}" & vbNewLine
txtCode.Text = txtCode.Text + "}" & vbNewLine
txtCode.Text = txtCode.Text + "function mouseout396()" & vbNewLine
txtCode.Text = txtCode.Text + "{" & vbNewLine
txtCode.Text = txtCode.Text + "if (document.images) {document.img396.src = MouseOutImage396.src;}" & vbNewLine
txtCode.Text = txtCode.Text + "}" & vbNewLine
txtCode.Text = txtCode.Text + "function mousedown396()" & vbNewLine
txtCode.Text = txtCode.Text + "{" & vbNewLine
txtCode.Text = txtCode.Text + "if (document.images) {document.img396.src = MouseDownImage396.src;}" & vbNewLine
txtCode.Text = txtCode.Text + "}" & vbNewLine
txtCode.Text = txtCode.Text + "if (document.images)" & vbNewLine
txtCode.Text = txtCode.Text + "{" & vbNewLine
txtCode.Text = txtCode.Text + "MouseOverImage396 = new Image();" & vbNewLine
txtCode.Text = txtCode.Text + "MouseOutImage396 = new Image();" & vbNewLine
txtCode.Text = txtCode.Text + "MouseDownImage396 = new Image();" & vbNewLine
txtCode.Text = txtCode.Text + "MouseOverImage396.src = '" & txtMouseOver.Text & "';" & vbNewLine
txtCode.Text = txtCode.Text + "MouseOutImage396.src = '" & txtNormalImage.Text & "';" & vbNewLine
txtCode.Text = txtCode.Text + "MouseDownImage396.src = '" & txtMouseDown.Text & "';" & vbNewLine
txtCode.Text = txtCode.Text + "}" & vbNewLine
txtCode.Text = txtCode.Text + "// -->" & vbNewLine
txtCode.Text = txtCode.Text + "</script>" & vbNewLine
txtCode.Text = txtCode.Text + "<a href='javascript:click396()' onMouseOver='mouseover396();' onMouseOut='mouseout396();' onMouseDown='mousedown396();'>" & vbNewLine
txtCode.Text = txtCode.Text + "<img name='img396' src='" & txtNormalImage.Text & "' alt='" & txtLoadingText.Text & "' border=0 vspace=0 hspace=0></a>" & vbNewLine

End Sub

Private Sub Command2_Click()
txtURL.Text = ""
txtMouseOver.Text = ""
txtNormalImage.Text = ""
txtLoadingText.Text = ""
txtCode.Text = ""
txtMouseDown.Text = ""
End Sub

Private Sub Label6_Click()
MsgBox "Visit: http://www.angelfire.com/oh4/elementx", vbOKOnly + vbInformation, "Visit Our Web Site"
End Sub
