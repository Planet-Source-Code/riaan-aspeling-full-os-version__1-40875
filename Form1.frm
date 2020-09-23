VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "OS Version Info"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim cOS As New clsOS
    txt.Text = "Name         : " & cOS.OS_Name & " " & cOS.OS_ProductType & vbCrLf & _
               "Version      : " & cOS.OS_Version & vbCrLf & _
               "Build        : " & cOS.OS_Build & vbCrLf & _
               "Additional   : " & cOS.OS_Additional & vbCrLf & _
               "Service Pack : " & cOS.OS_ServicePack & vbCrLf & _
               "Suites       : " & cOS.OS_Suite
    Set cOS = Nothing
End Sub

