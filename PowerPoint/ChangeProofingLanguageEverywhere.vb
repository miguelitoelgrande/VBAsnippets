
Sub Auto_Open()
' save the PowerPoint containing this macro/module as .pptm AND .ppam
' go to File->Options->AddIns and add the .ppam file as Add-In.
' the custom Toolbar should appear.
    Dim oToolbar As CommandBar
    Dim oButton1 As CommandBarButton
    Dim oButton2 As CommandBarButton
    Dim oButton3 As CommandBarButton
    Dim MyToolbar As String

    ' Give the toolbar a name
    MyToolbar = "Set Proofing Language"

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
    Set oButton1 = oToolbar.Controls.Add(Type:=msoControlButton)
    Set oButton2 = oToolbar.Controls.Add(Type:=msoControlButton)
    Set oButton3 = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
    With oButton1
         .DescriptionText = "Lang button" 'Tooltip text when mouse if placed over button
         .Caption = "UK"   'Text if Text in Icon is chosen
         .OnAction = "ClickButtonUK"  'Runs the Sub Button1() code when clicked
         .Style = msoButtonIconAndCaptionBelow ' msoButtonIcon
         .TooltipText = "Set Proofing Language UK (all slide objects)"
         .FaceId = 2566 '1610  ' chooses icon #52 from the available Office icons
          ' http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
          ' http://officeone.mvps.org/faceid/FaceId.ppa
    End With

    ' And set some of the button's properties
    With oButton2
         .DescriptionText = "Lang button" 'Tooltip text when mouse if placed over button
         .Caption = "US"   'Text if Text in Icon is chosen
         .OnAction = "ClickButtonUS"  'Runs the Sub Button1() code when clicked
         .Style = msoButtonIconAndCaptionBelow ' msoButtonIcon
         .TooltipText = "Set Proofing Language US (all slide objects)"
         .FaceId = 2566 '1610  ' chooses icon #52 from the available Office icons
          ' http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
          ' http://officeone.mvps.org/faceid/FaceId.ppa
    End With
    
        ' And set some of the button's properties
    With oButton3
         .DescriptionText = "Lang button" 'Tooltip text when mouse if placed over button
         .Caption = "DE"   'Text if Text in Icon is chosen
         .OnAction = "ClickButtonDE"  'Runs the Sub Button1() code when clicked
         .Style = msoButtonIconAndCaptionBelow ' msoButtonIcon
         .TooltipText = "Set Proofing Language DE (all slide objects)"
         .FaceId = 2566 '1610  ' chooses icon #52 from the available Office icons
          ' http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
          ' http://officeone.mvps.org/faceid/FaceId.ppa
    End With

    ' Repeat the above for as many more buttons as you need to add
    ' Be sure to change the .OnAction property at least for each new button

    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.Visible = True

NormalExit:
    Exit Sub   ' so it doesn't go on to run the errorhandler code

ErrorHandler:
     'Just in case there is an error
     MsgBox Err.Number & vbCrLf & Err.Description
     Resume NormalExit:
End Sub

Sub ClickButtonUK()
    Call ChangeSpellCheckingLanguage(msoLanguageIDEnglishUK)
End Sub

Sub ClickButtonUS()
    Call ChangeSpellCheckingLanguage(msoLanguageIDEnglishUS)
End Sub

Sub ClickButtonDE()
    Call ChangeSpellCheckingLanguage(msoLanguageIDGerman)
End Sub


Private Sub ChangeSpellCheckingLanguageSIMPLE(langID As Integer)
' idea: http://answers.microsoft.com/en-us/office/forum/office_2007-powerpoint/setting-language-for-entire-presentation-all-at/4f8ab731-9fcf-4b64-9433-2bac1b51eff3?auth=1
    Dim j As Integer, k As Integer, scount As Integer, fcount As Integer

    scount = ActivePresentation.Slides.Count
    For j = 1 To scount
        fcount = ActivePresentation.Slides(j).Shapes.Count
        For k = 1 To fcount
            If ActivePresentation.Slides(j).Shapes(k).HasTextFrame Then
                ActivePresentation.Slides(j).Shapes(k) _
                .TextFrame.TextRange.LanguageID = langID ' msoLanguageIDEnglishUK
                ' more languages: https://msdn.microsoft.com/en-us/library/aa432635.aspx
                ' msoLanguageIDEnglishUK, msoLanguageIDEnglishUS, msoLanguageIDGerman
            End If
        Next k
    Next j

End Sub


Private Sub ChangeSpellCheckingLanguage(langID As Integer)
' http://stackoverflow.com/questions/4735765/powerpoint-2007-set-language-on-tables-charts-etc-that-contains-text
    On Error Resume Next
    Dim gi As GroupShapes '<-this was added. used below
    'lang = "English"
    'lang = "German"
    ' more languages: https://msdn.microsoft.com/en-us/library/aa432635.aspx
    'If lang = "English" Then
    '    lang = msoLanguageIDEnglishUK ' msoLanguageIDEnglishUK
    'ElseIf lang = "Norwegian" Then
    '    lang = msoLanguageIDGerman
    'End If
    lang = langID
    'Set default language in application
    'ActivePresentation.DefaultLanguageID = lang

    'Set language in each textbox in each slide
    For Each oSlide In ActivePresentation.Slides
        Dim oShape As Shape
        For Each oShape In oSlide.Shapes
            'Check first if it is a table
            If oShape.HasTable Then
                For r = 1 To oShape.Table.Rows.Count
                    For c = 1 To oShape.Table.Columns.Count
                    oShape.Table.Cell(r, c).Shape.TextFrame.TextRange.LanguageID = lang
                    Next
                Next
            Else
                Set gi = oShape.GroupItems
                'Check if it is a group of shapes
                If Not gi Is Nothing Then
                    oShape.TextFrame.TextRange.LanguageID = lang ' apply to group as well...
                    If oShape.GroupItems.Count > 0 Then
                        For i = 0 To oShape.GroupItems.Count - 1
                            oShape.GroupItems(i).TextFrame.TextRange.LanguageID = lang
                        Next
                    End If
                'it's none of the above, it's just a simple shape, change the language ID
                Else
                    oShape.TextFrame.TextRange.LanguageID = lang
                End If
            End If
        Next
    Next
End Sub



