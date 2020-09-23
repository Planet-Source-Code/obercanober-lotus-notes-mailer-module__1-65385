Attribute VB_Name = "ModLotusMailer"
    
    
    '********************************************************************************************
    'Lotus Notes Mailer Module Created May 17th, 2006 for Status Update E-mail Automation
    'Lotus Mailer tested with Lotus Notes Release 6.5.1
    'Module created by Jeff Olson (Copyright © Ober-Soft.com)
    'Reference
    'http://www-128.ibm.com/developerworks/lotus/library/ls-COM_Access/#N101CC
    'http://www-128.ibm.com/developerworks/lotus/library/ls-ND6_LSrichtext/
    '********************************************************************************************
    
    
    
Dim ln_Line As New NotesRichTextItem
Dim ln_Doc As New NotesDocument
Dim ln_session As New NotesSession
Dim ln_rtnav As New NotesRichTextNavigator
Dim ln_dir As NotesDbDirectory
Dim ln_db As NotesDatabase
Dim ln_rtstyle As NotesRichTextStyle
Private MailStarted As Boolean
Public Sub StartNewMail()
    MailStarted = True
    Set ln_session = CreateObject("Lotus.NotesSession")
    Call ln_session.Initialize
    Set ln_rtstyle = ln_session.CreateRichTextStyle
    Set ln_dir = ln_session.GetDbDirectory("")
    Set ln_db = ln_dir.OpenMailDatabase
    Set ln_Doc = ln_db.CreateDocument
    Set ln_Line = ln_Doc.CreateRichTextItem("Body")
    Set ln_rtnav = ln_Line.CreateNavigator
End Sub
Public Sub AttachFile(TheFilePath)
    Dim lnATTACHMENT As Object
    Set lnATTACHMENT = ln_Line.EmbedObject(1454, "", TheFilePath, "Sample")
End Sub
Public Sub SetMailFrom(TheFromAddress As String, TheDisplaySent As String)
    If MailStarted = False Then Exit Sub
    ln_Doc.ReplaceItemValue "Principal", TheFromAddress
    ln_Doc.ReplaceItemValue "DisplaySent", TheDisplaySent
End Sub
Public Sub SetMailSubject(TheSubject As String)
    If MailStarted = False Then Exit Sub
    ln_Doc.ReplaceItemValue "Subject", TheSubject
End Sub
Public Sub SetMailSendTo(SendTo As String)
    If MailStarted = False Then Exit Sub
    ln_Doc.ReplaceItemValue "SendTo", SendTo
End Sub
Public Sub InsertRichTextLine(TheLine As String, Optional TheColor As COLORS, Optional TheFont As RT_FONTS, Optional TheSize As Long, Optional DoBold As Boolean, Optional DoItalic As Boolean, Optional DoNewLine As Boolean)
    If MailStarted = False Then Exit Sub
    If TheColor <> 0 Then ln_rtstyle.NotesColor = TheColor
    If TheFont <> 0 Then ln_rtstyle.NotesFont = TheFont
    If TheSize <> 0 Then ln_rtstyle.FontSize = TheSize
    If DoBold = True Then ln_rtstyle.Bold = 1
    If DoItalic = True Then ln_rtstyle.Italic = 1
    Call ln_Line.AppendStyle(ln_rtstyle)
    Call ln_Line.AppendText(TheLine)
    If DoNewLine = True Then Call ln_Line.AppendText(vbNewLine)
End Sub
Public Sub CreateTable(RowCount As Long, ColumnCount As Long)
    If MailStarted = False Then Exit Sub
    Call ln_Line.AppendTable(RowCount, ColumnCount)
End Sub
Public Sub InsertTableValue(TheValue As String, TheFieldNumber As Long, Optional TheColor As COLORS, Optional TheFont As RT_FONTS, Optional TheSize As Long, Optional DoBold As Boolean, Optional DoItalic As Boolean, Optional BackGroundColor As COLORS)
    If MailStarted = False Then Exit Sub
    If TheColor <> 0 Then ln_rtstyle.NotesColor = TheColor
    If TheFont <> 0 Then ln_rtstyle.NotesFont = TheFont
    If TheSize <> 0 Then ln_rtstyle.FontSize = TheSize
    If DoBold = True Then ln_rtstyle.Bold = 1
    If DoItalic = True Then ln_rtstyle.Italic = 1
    Call ln_rtnav.FindNthElement(RTELEM_TYPE_TABLECELL, TheFieldNumber)
    Call ln_Line.BeginInsert(ln_rtnav)
    Call ln_Line.AppendText(TheValue)
    Call ln_Line.EndInsert
End Sub
Public Sub SendMail()
    If MailStarted = False Then Exit Sub
    ln_Doc.Send False
    MailStarted = False
    Set ln_session = Nothing
    Set ln_rtstyle = Nothing
    Set ln_dir = Nothing
    Set ln_db = Nothing
    Set ln_Doc = Nothing
    Set ln_Line = Nothing
    Set ln_rtnav = Nothing
End Sub
