Attribute VB_Name = "SetLanguage"
' Name: SetLanguage
' Type: PowerPoint AddIn VBS code
' Language: Visual Basic
' Developed by Rafal Urniaz
' Purpose: Set language to all elements in PowerPoint presentation even if the default set language command doesn't work properly
' Based on code here : http://www.pptfaq.com/FAQ00031 Create_an_ADD-IN_with_TOOLBARS_that_run_macros.htm

Sub Auto_Open()

    Dim oToolbar As CommandBar
    Dim oButton As CommandBarButton
    Dim InfoButton As CommandBarButton
    
    ' Give the toolbar a name
    Dim MyToolbar As String
    MyToolbar = "Set language"

    On Error Resume Next
    
    ' so that it doesn't stop on the next line if the toolbar's already there

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=MyToolbar, _
        Position:=msoBarFloating, Temporary:=True)
        
    If Err.Number <> 0 Then
          ' The toolbar's already there, so we have nothing to do
          Exit Sub
    End If

    On Error GoTo ErrorHandler

    ' Now add a button to the new toolbar
    Set oButton = oToolbar.Controls.Add(Type:=msoControlButton)
    Set InfoButton = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties

    With oButton

         .DescriptionText = "Set language to all slides and objects"
          'Tooltip text when mouse if placed over button

         .Caption = "Set language"
         'Text if Text in Icon is chosen

         .OnAction = "SetLanguage"
          'Runs the Sub SetLanguage() code when clicked

         .Style = msoButtonIconAndCaptionBelow
          ' Button displays as icon, not text or both

         .FaceId = 790  '610 '4166
          ' chooses icon #52 from the available Office icons
             
    End With
    
    With InfoButton
         .DescriptionText = "Explains how to use the solution"
         .Caption = "How to use it"
         .OnAction = "ShowInfo"
         .Style = msoButtonIcon
         .FaceId = 487
         .Width = 150
    End With
    
    ' Repeat the above for as many more buttons as you need to add
    ' Be sure to change the .OnAction property at least for each new button

    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.Width = 500
    oToolbar.Visible = True
    
NormalExit:
    Exit Sub   ' so it doesn't go on to run the errorhandler code

ErrorHandler:
     'Just in case there is an error
     MsgBox Err.Number & vbCrLf & Err.Description
     Resume NormalExit:
End Sub

Sub SetLanguage()

'Determine Which Shape is Active
  If Application.Windows.Count >= 1 Then
  
  Dim ID As String
  
 ' ID = ActiveShape.TextFrame.TextRange.LanguageID
 ' This general solution works better
 
 ID = ActiveWindow.Selection.TextRange.LanguageID
 
 'Call set language
 
  SetLang (ID)

 Else
     MsgBox "There is no opened document yet!", vbExclamation, "No active document found ..."
    ' MsgBox "There is no object currently selected!", vbExclamation, "No Selection Found"
 End If
 
End Sub

Sub SetLang(ID As Integer)

' set language for all slides and notes:

Dim scount, j, k, fcount

scount = ActivePresentation.Slides.Count

For j = 1 To scount
    fcount = ActivePresentation.Slides(j).Shapes.Count
        For k = 1 To fcount
        'change all shapes:
            If ActivePresentation.Slides(j).Shapes(k).HasTextFrame Then
               ActivePresentation.Slides(j).Shapes(k).TextFrame _
                    .TextRange.LanguageID = ID
            End If
        Next k
            
        'change notes:
        fcount = ActivePresentation.Slides(j).NotesPage.Shapes.Count
            
        For k = 1 To fcount
        'change all shapes:
            If ActivePresentation.Slides(j).NotesPage.Shapes(k).HasTextFrame Then
            ActivePresentation.Slides(j).NotesPage.Shapes(k).TextFrame _
                .TextRange.LanguageID = ID
            End If
        Next k
        
Next j
End Sub

Sub ShowInfo()
    'Dim Message As String
    
        InfoForm.Show
    
    'Message = "This Add-In sets appropriate language to all elements (shapes, text, boxes etc.) in PowerPoint presentation even the default set language command doesn't do it properly." + vbNewLine + vbNewLine + "Set any language by:" + vbNewLine + vbNewLine + "   1) Open appropriate presentation" + vbNewLine + "   2) Select text, shape or any different item" + vbNewLine + "   3) Set language of interest" + vbNewLine + "   4) Clock 'Set language' button on the ribbon panel" + vbNewLine + "   5) Done!" + vbNewLine + vbNewLine + "Set custom language by:" + vbNewLine + vbNewLine + "  1)  Simply click flag and language name on the ribbon panel" + vbNewLine + vbNewLine + "---" + vbNewLine + "Add-In solution developed by" + vbNewLine + "BioTesseract Cambridge Bioinformatics Solutions" + vbNewLine + "Contact at biotesseract@biotesseract.com"
    'MsgBox Message, vbInformation, "About ..."
End Sub

