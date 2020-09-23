VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lotus Mailer"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Test"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    StartNewMail
    SetMailSubject "TEST"
    SetMailSendTo "Spam@Spam.com"
    SetMailFrom "MadSpammer@Spam.com", "Super Spammer!"
    InsertRichTextLine "Line 1", COLOR_BLUE, 0, 0, True, True, True
    InsertRichTextLine "Line 2", COLOR_DARK_RED, 0, 0, True, True, True
    CreateTable 3, 5
    InsertTableValue "Monday", 1, COLOR_DARK_RED, 0, 12, True, False
    InsertTableValue "Tuesday", 2, COLOR_DARK_RED, 0, 12, True, False
    InsertTableValue "Wednesday", 3, COLOR_DARK_RED, 0, 12, True, False
    InsertTableValue "Thursday", 4, COLOR_DARK_RED, 0, 12, True, False
    InsertTableValue "Friday", 5, COLOR_DARK_RED, 0, 12, True, False
    AttachFile App.Path & "\attachment.txt"
    SendMail
End Sub

