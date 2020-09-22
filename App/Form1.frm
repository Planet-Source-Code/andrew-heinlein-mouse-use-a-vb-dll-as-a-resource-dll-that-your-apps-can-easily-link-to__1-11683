VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   720
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      Height          =   1095
      Left            =   1440
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1035
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim ResourceDLL As String
    
    'The resouce DLL should be in the same directory as this app
    'so let's set the path:
    If Right(App.Path, 1) <> "\" Then
        ResourceDLL = App.Path & "\" & "VbRes.dll"
    Else
        ResourceDLL = App.Path & "VbRes.dll"
    End If
    
    'see if the file is where its supposed to be:
    If Dir(ResourceDLL) = "" Then
        MsgBox "Missing the VbRes.DLL needed to demonstrait!"
        Exit Sub
    End If
    
    'yadda yadda:
    Form1.Caption = ResourceDLL
    List1.Clear
    
    'Use our great functions to get the BMP from the DLL
    LoadResourceToPicBox ResourceDLL, 101, Picture1
    LoadResourceToPicBox ResourceDLL, 102, Picture2
    LoadResourceToPicBox ResourceDLL, 103, Picture3
    
    'use our great functions to get the string resources.
    List1.AddItem LoadResourceString(ResourceDLL, 101)
    List1.AddItem LoadResourceString(ResourceDLL, 102)
    List1.AddItem LoadResourceString(ResourceDLL, 103)
    List1.AddItem LoadResourceString(ResourceDLL, 104)

End Sub

