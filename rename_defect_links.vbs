Sub renameDefectLinks()
	' Define RTC domain.
	Dim rtcDomain As String = "https://rtc.domain.com:9443/ccm/web/projects/"
	Dim projectName As String = "PROJECT"
	
    ' Define Outlook dependencies.
    Dim inspector As Outlook.inspector = Application.ActiveInspector
    Dim editor As Word.Document = inspector.WordEditor
    Dim selection As Word.selection = editor.Application.selection
    
	' For every defect link in the email replace the link text up to the defect ID with "defect #".
    With selection.Range
        With selection.Find
            .Text = rtcDomain & projectName & "#action=com.ibm.team.workitem.viewWorkItem&id="
            .Replacement.Text = "defect #"
            .MatchWildcards = True
            .Execute Replace:=wdReplaceAll
        End With
    End With
    
    ' Clear selection.
    editor.Application.selection.Move
End Sub