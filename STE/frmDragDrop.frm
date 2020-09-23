VERSION 5.00
Begin VB.Form frmDragDrop 
   Caption         =   "Shortcut Target Extractor"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblDragFilesHere 
      Alignment       =   2  'Center
      Caption         =   "Drag Shortcut (LNK) Files Here"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   39
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   360
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   0
      Width           =   3870
   End
End
Attribute VB_Name = "frmDragDrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Drag drop code
    'not commented because it really
    'is only here to facilitate what
    'goes on in the module
    Dim i As Long
    On Error GoTo errHandler
        For i = 1 To Data.Files.Count
            If UCase(Right(Data.Files(i), 4)) = ".LNK" Then
                MsgBox GetLNKTarget(Data.Files(i))
            End If
        Next i
errHandler:
    Exit Sub
End Sub

Private Sub lblDragFilesHere_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Drag drop code
    'not commented because it really
    'is only here to facilitate what
    'goes on in the module
    Dim i As Long
    On Error GoTo errHandler
        For i = 1 To Data.Files.Count
            If UCase(Right(Data.Files(i), 4)) = ".LNK" Then
                MsgBox GetLNKTarget(Data.Files(i))
            End If
        Next i
errHandler:
    Exit Sub
End Sub
