'https://github.com/shishouyuan/HandyRefVBA

'A handy way to insert Cross Reference in MS Word
'Author: Shouyuan Shi @ South China University of Technology
'E-mail: shishouyuan@outlook.com
'Creating Date: 2021/5/11


'用于在Word里方便地添加交叉引用
'作者: 史守圆 @ 华南理工大学
'E-mail: shishouyuan@outlook.com
'创建时期: 2021/5/11


Const HandyRefVersion = "20230616.1914.VBA"

Const TEXT_HandyRefGithubUrl = "https://github.com/shishouyuan/HandyRefVBA"
Const TEXT_HandyRefZhihuUrl = "https://zhuanlan.zhihu.com/p/373677845"

Const BookmarkPrefix = "_HandyRef"
Const RefBrokenCommentTitle = "$HANDYREF_REFERENCE_BROKEN_COMMENT$"

#Const HandyRef_Lang = "zh-cn"

Const BrokenRefNumPosHolder = "#"

#If HandyRef_Lang = "zh-cn" Then

    Const TEXT_HandyRefAppName = "HandyRef-快引"
    Const TEXT_HandyRefAuthor = "史守圆 @ 华南理工大学"
    Const TEXT_HandyRefDescription = "为 Word 提供一个快速添加交叉引用的方式。"
    Const TEXT_CreateReferencePoint_NothingSelected = "请先选中要引用的内容!"
    Const TEXT_InsertCrossReferenceField_NoRefPoint = "当前没有设置引用点!"
    Const TEXT_InsertCrossReferenceField_CannotCrossFile = "不支持跨文件引用!"
    Const TEXT_VersionPrompt = "版本："
    Const TEXT_NonCommecialPrompt = "仅限非商业用途"
    Const TEXT_RefBrokenComment = "引用源丢失！"
    Const TEXT_BrokenRefFoundPrompt = "发现了 " & BrokenRefNumPosHolder & " 个损坏的引用，已为其添加批注。"
    Const TEXT_NoBrokenRefFoundPrompt = "没有发现损坏的引用。"
    Const TEXT_RefBrokenCommentClearedPrompt = "引用损坏批注已清除。"
    Const TEXT_RefCheckingForWholeDocPrompt = "当前没有选中的内容，检查整个文档吗？" & vbCrLf & "这可能需要一些时间。"
    Const TEXT_ClearRefBrokenCommentForWholeDocPrompt = "当前没有选中的内容，清除整个文档中的引用损坏批注吗？"
    Const TEXT_UnknownErrOccurredPrompt = "遇到错误："
    Const TEXT_ActionName_CreateSource = "创建引用源"
    Const TEXT_ActionName_InsertReference = "交叉引用"
    Const TEXT_ActionName_CheckReference = "检查引用"
    Const TEXT_ActionName_ClearRefBrokenComment = "清除批注"
    

#Else

    Const TEXT_HandyRefAppName = "HandyRef"
    Const TEXT_HandyRefAuthor = "Shouyuan Shi @ South China University of Technology"
    Const TEXT_HandyRefDescription = "Provide a handy way to insert Cross Reference in MS Word."
    Const TEXT_CreateReferencePoint_NothingSelected = "Nothing selected!"
    Const TEXT_InsertCrossReferenceField_NoRefPoint = "No Reference Point Selected!"
    Const TEXT_InsertCrossReferenceField_CannotCrossFile = "Cross file reference is not supported!"
    Const TEXT_VersionPrompt = "Version: "
    Const TEXT_NonCommecialPrompt = "Only for NON-COMMERCIAL use."
    Const TEXT_RefBrokenComment = "Reference Broken!"
    Const TEXT_BrokenRefFoundPrompt = BrokenRefNumPosHolder & " broken reference found. Comments are attached."
    Const TEXT_NoBrokenRefFoundPrompt = "No broken reference found."
    Const TEXT_RefBrokenCommentClearedPrompt = "Reference broken comments cleared."
    Const TEXT_RefCheckingForWholeDocPrompt = "Nothing is selected. Check the whole document?" & vbCrLf & "This may take a while."
    Const TEXT_ClearRefBrokenCommentForWholeDocPrompt = "Nothing is selected. Clear reference broken comments for the whole document?"
    Const TEXT_UnknownErrOccurredPrompt = "Error occurred:"
    Const TEXT_ActionName_CreateSource = "Create Source"
    Const TEXT_ActionName_InsertReference = "Insert Reference"
    Const TEXT_ActionName_CheckReference = "Check Reference"
    Const TEXT_ActionName_ClearRefBrokenComment = "Clear Comments"
    
#End If

Public Enum RefTypes
    Normal = 0
    ParaNumber = 1
    PageNumber = 2
    RelativePosition = 3
End Enum

Private selectedRange As Range
Private selectedBM As Bookmark
Private selectedIsNote As Boolean
Private selectedHeading As Boolean
Private editEnabled As Boolean


Private ribbonUI As IRibbonUI
Private helper As helper
Public Sub HandyRef_OnLoad(ByVal rb As IRibbonUI)
    Set ribbonUI = rb
    Set helper = New helper
    Set helper.App = Application
End Sub

Public Sub HandyRef_UpdateRibbonState()
    ribbonUI.Invalidate
End Sub

Public Sub HandyRef_GetEnabled(ByVal control As IRibbonControl, ByRef enabled)
    On Error GoTo noDoc
    editEnabled = Not Application.ActiveWindow.Document Is Nothing
    enabled = editEnabled
    Exit Sub
    
noDoc:
    editEnabled = False
    enabled = editEnabled
End Sub


Private Function FormatUndoRecordText(s As String) As String
    FormatUndoRecordText = s & "-" & TEXT_HandyRefAppName
End Function

Private Sub ShowUnknowErrorPrompt(e As ErrObject)
    MsgBox TEXT_UnknownErrOccurredPrompt & vbCrLf & e.Description, vbOKOnly + vbExclamation, TEXT_HandyRefAppName
End Sub

Public Sub HandyRef_CreateReferencePoint_RibbonFun(ByVal control As IRibbonControl) ' wrap the function to match the signature called by ribbion
    HandyRef_CreateReferencePoint
End Sub

Public Sub HandyRef_CreateReferencePoint()

    Dim rg As Range
    Set rg = Application.Selection.Range
   
    selectedIsNote = False
    Set selectedRange = rg
    Set selectedBM = Nothing
    
    On Error Resume Next    'will cause error when accessing endnotes property when the range is in footnote section
    If rg.Endnotes.Count = 0 Then
    On Error GoTo exitSub
        If rg.Footnotes.Count = 1 Then
            Dim fn As Footnote
            Set fn = rg.Footnotes.Item(1)
            If rg.InRange(fn.Range) Or rg.InRange(fn.Reference) Or Not fn.Reference.InRange(rg) Then
                selectedIsNote = True
                Set selectedRange = fn.Reference
            End If
        End If
    End If
    
    On Error Resume Next
    If rg.Footnotes.Count = 0 Then
    On Error GoTo exitSub
        If rg.Endnotes.Count = 1 Then
        Dim en As Endnote
            Set en = rg.Endnotes.Item(1)
            If rg.InRange(en.Range) Or rg.InRange(en.Reference) Or Not en.Reference.InRange(rg) Then
                selectedIsNote = True
                Set selectedRange = en.Reference
            End If
        End If
    End If

exitSub:
    If rg.End = rg.Start And Not selectedIsNote Then
        Set selectedRange = Nothing
        Set selectedBM = Nothing
        MsgBox TEXT_CreateReferencePoint_NothingSelected, vbOKOnly + vbInformation, TEXT_HandyRefAppName
    End If
    
    HandyRef_UpdateRibbonState
    
End Sub

Public Sub HandyRef_InsertCrossReferenceField_SplitButton_GetEnabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = editEnabled And Not selectedRange Is Nothing
End Sub

Public Sub HandyRef_InsertCrossReferenceField_Menu_GetVisible(ByVal control As IRibbonControl, ByRef enabled)
    enabled = editEnabled And Not selectedRange Is Nothing And Not selectedIsNote
End Sub

Public Sub HandyRef_InsertCrossReferenceField_Normal_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_InsertCrossReferenceField_With_Type RefTypes.Normal
End Sub

Public Sub HandyRef_InsertCrossReferenceField_ParaNumber_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_InsertCrossReferenceField_With_Type RefTypes.ParaNumber
End Sub

Public Sub HandyRef_InsertCrossReferenceField_PageNumber_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_InsertCrossReferenceField_With_Type RefTypes.PageNumber
End Sub

Public Sub HandyRef_InsertCrossReferenceField_RelativePosition_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_InsertCrossReferenceField_With_Type RefTypes.RelativePosition
End Sub



Private Function GetTimeStamp() As String
    'Date variables are stored as IEEE 64-bit (8-byte) floating-point numbers
    'When other numeric types are converted to Date, values to the left of the decimal represent date information,
    'while values to the right of the decimal represent time. Midnight is 0 and midday is 0.5.
    'Double (double-precision floating-point) variables are stored as IEEE 64-bit (8-byte) floating-point numbers
    GetTimeStamp = Replace(CStr(CDbl(Now)), ".", "")
End Function

Public Sub HandyRef_InsertCrossReferenceField()
    HandyRef_InsertCrossReferenceField_With_Type RefTypes.Normal
End Sub


Private Sub HandyRef_InsertCrossReferenceField_With_Type(refType As RefTypes)

    On Error GoTo errHandle
    Application.UndoRecord.StartCustomRecord FormatUndoRecordText(TEXT_ActionName_InsertReference)
    
    Dim setToFirstPara As Boolean
    If refType = RefTypes.ParaNumber Then
        setToFirstPara = True
    Else
        setToFirstPara = False
    End If
    
    Dim bmValid As Boolean
    bmValid = False
    Dim targetRange As Range
    
    If setToFirstPara Then
        Set targetRange = selectedRange.Paragraphs.First.Range
    Else
        Set targetRange = selectedRange
    End If

    If Not selectedBM Is Nothing Then
        If Application.IsObjectValid(selectedBM) Then
            If selectedBM.Parent Is ActiveDocument Then
                If selectedBM.Range.IsEqual(targetRange) Then
                    bmValid = True
                End If
            Else
crossFile:
                MsgBox TEXT_InsertCrossReferenceField_CannotCrossFile, vbOKOnly + vbInformation, TEXT_HandyRefAppName
                GoTo exitSub
            End If
        Else ' it's possible the bookmark is deleted by the user, but the range remaind.
            Set selectedBM = Nothing
        End If
    End If
    If Not bmValid Then
        If targetRange Is Nothing Then
            GoTo emptyRange
        ElseIf Not Application.IsObjectValid(targetRange) Or targetRange.Start = targetRange.End Then
emptyRange:
            Set selectedRange = Nothing
            MsgBox TEXT_InsertCrossReferenceField_NoRefPoint, vbOKOnly + vbInformation, TEXT_HandyRefAppName
            GoTo exitSub
        ElseIf Not targetRange.Document Is ActiveDocument Then
            GoTo crossFile
        Else
            Dim oldbm As Bookmark
            Dim bmi As Bookmark
            Dim bmShowHiddenOld As Boolean
            bmShowHiddenOld = targetRange.Bookmarks.ShowHidden
            
            'search for existing bookmark reference the same range
            targetRange.Bookmarks.ShowHidden = True
            For Each bmi In targetRange.Bookmarks
                If bmi.Range.IsEqual(targetRange) And bmi.Name Like BookmarkPrefix & "#*" Then
                    Set oldbm = bmi
                    Exit For
                End If
            Next bmi
            targetRange.Bookmarks.ShowHidden = bmShowHiddenOld
            
            If Not oldbm Is Nothing Then
                Set selectedBM = oldbm
            Else
                'create new bookmark using timestamp as its name
                Set selectedBM = targetRange.Bookmarks.Add(BookmarkPrefix & GetTimeStamp(), targetRange)
            End If
            
            bmValid = True
        End If

    End If
    
    If bmValid Then
        If selectedIsNote Then
            If refType <> RefTypes.Normal Then
                MsgBox "Action not supported for footnote or endnote."
                GoTo exitSub
            End If
            ActiveDocument.Fields.Add Selection.Range, WdFieldType.wdFieldNoteRef, selectedBM.Name & " \h "
        Else
            Select Case refType
            Case RefTypes.Normal
                ActiveDocument.Fields.Add Selection.Range, WdFieldType.wdFieldRef, selectedBM.Name & " \h"
            Case RefTypes.ParaNumber
                ActiveDocument.Fields.Add Selection.Range, WdFieldType.wdFieldRef, selectedBM.Name & " \h \w"
            Case RefTypes.PageNumber
                ActiveDocument.Fields.Add Selection.Range, WdFieldType.wdFieldPageRef, selectedBM.Name & " \h"
            Case RefTypes.RelativePosition
                ActiveDocument.Fields.Add Selection.Range, WdFieldType.wdFieldRef, selectedBM.Name & " \h \p"
            End Select
        End If
    End If
    
exitSub:
    Application.UndoRecord.EndCustomRecord
    HandyRef_UpdateRibbonState
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt Err
    GoTo exitSub
    
End Sub

Public Sub HandyRef_ClearRefBrokenComment_RibbonFun(ByVal control As IRibbonControl)
    
    If Application.Selection.End = Application.Selection.Start Then
        If MsgBox(TEXT_ClearRefBrokenCommentForWholeDocPrompt, vbOKCancel + vbQuestion, TEXT_HandyRefAppName) = vbOK Then
            HandyRef_ClearRefBrokenComment ActiveDocument.Range
        Else
            Exit Sub
        End If
    Else
        HandyRef_ClearRefBrokenComment Application.Selection.Range
    End If
    MsgBox TEXT_RefBrokenCommentClearedPrompt, vbOKOnly + vbInformation, TEXT_HandyRefAppName
    
End Sub

Public Sub HandyRef_ClearRefBrokenComment(targetRange As Range)
    Application.UndoRecord.StartCustomRecord FormatUndoRecordText(TEXT_ActionName_ClearRefBrokenComment)
    On Error GoTo errHandle
    
    Dim cmt As Comment
    Dim s As String
    For Each cmt In targetRange.Comments
        If cmt.Reference.InRange(targetRange) Then 'targetRange.Comments will also return comments before targentRange. may be a bug or misunderstanding
            s = cmt.Range.Paragraphs.Last.Range.Text
            s = Replace(s, vbCr, "")
            s = Replace(s, vbLf, "")
            s = Trim(s)
            If StrComp(s, RefBrokenCommentTitle) = 0 Then
                cmt.DeleteRecursively
                
            End If
        End If
    Next cmt
    
exitSub:
    Application.UndoRecord.EndCustomRecord
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt Err
    GoTo exitSub
    
End Sub
 
Public Sub HandyRef_CheckForBrokenRef_RibbonFun(ByVal control As IRibbonControl)
    If Application.Selection.End - Application.Selection.Start = 0 Then
        If MsgBox(TEXT_RefCheckingForWholeDocPrompt, vbOKCancel + vbQuestion, TEXT_HandyRefAppName) = vbOK Then
            HandyRef_CheckForBrokenRef ActiveDocument.Range
        End If
    Else
        HandyRef_CheckForBrokenRef Application.Selection.Range
    End If
    
End Sub

Public Sub HandyRef_CheckForBrokenRef(checkingRange As Range)
    
    Dim oldScreenUpdating As Boolean
    oldScreenUpdating = Application.ScreenUpdating
    
    On Error GoTo errHandle
    
    Application.ScreenUpdating = False
    
    Application.UndoRecord.StartCustomRecord FormatUndoRecordText(TEXT_ActionName_CheckReference)
    
    HandyRef_ClearRefBrokenComment checkingRange
    
    Static refRegExp As Object
    Static refRegExp0 As Object
    If refRegExp Is Nothing Then
        Set refRegExp = CreateObject("VBScript.RegExp")
        With refRegExp
            .Global = False
            .IgnoreCase = True
            .Pattern = "^\s*(?:NOTE|PAGE)?REF.*\s([^\s\\]+).*"
            '.Pattern = "^\s*(?:NOTE)?REF.*?(?<!\\\*)\s+([^\s\\]+).*"
        End With
        
        Set refRegExp0 = CreateObject("VBScript.RegExp")
        With refRegExp0
            .Global = True
            .IgnoreCase = True
            .Pattern = "\\[*@#]\s*[^\s\\]*"
        End With
        
    End If
    
    Dim brokenCount As Integer
    brokenCount = 0
    
    Dim fd As Field
    Dim bmName As String
    
    For Each fd In checkingRange.Fields
        If fd.Type = wdFieldRef Or fd.Type = wdFieldNoteRef Or fd.Type = wdFieldPageRef Then
            Set r = refRegExp.Execute(refRegExp0.Replace(fd.code.Text, ""))
            Dim isBroken As Boolean
            isBroken = True
            If r.Count > 0 Then
                bmName = r.Item(0).SubMatches(0)
                If ActiveDocument.Bookmarks.Exists(bmName) Then
                    isBroken = False
                End If
            End If
            If isBroken Then
                brokenCount = brokenCount + 1
                Dim cmt As Comment
                Set cmt = fd.code.Comments.Add(fd.code)
                With cmt.Range
                    .InsertAfter TEXT_RefBrokenComment
                    .InsertParagraphAfter
                    .InsertAfter RefBrokenCommentTitle
                End With
                
                With cmt.Range.Paragraphs.First.Range
                    .Bold = True
                    .HighlightColorIndex = wdYellow
                End With
            End If
        End If
    Next fd
    
    If brokenCount = 0 Then
        MsgBox TEXT_NoBrokenRefFoundPrompt, vbOKOnly + vbInformation, TEXT_HandyRefAppName
    Else
        
        MsgBox Replace(TEXT_BrokenRefFoundPrompt, BrokenRefNumPosHolder, CStr(brokenCount)), vbOKOnly + vbInformation, TEXT_HandyRefAppName
        ActiveWindow.View.SplitSpecial = wdPaneNone
        ActiveWindow.View.SplitSpecial = wdPaneRevisions
    End If
    
exitSub:
    Application.ScreenUpdating = oldScreenUpdating
    Application.UndoRecord.EndCustomRecord
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt Err
    GoTo exitSub
    
End Sub



Public Sub HandyRef_About_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_About
End Sub

Public Sub HandyRef_About()
    MsgBox TEXT_HandyRefAppName + vbCrLf _
    + TEXT_HandyRefDescription + vbCrLf _
    + TEXT_NonCommecialPrompt + vbCrLf + vbCrLf _
    + TEXT_VersionPrompt + HandyRefVersion + vbCrLf _
    + TEXT_HandyRefAuthor + vbCrLf _
    + TEXT_HandyRefGithubUrl, _
    vbOKOnly + vbInformation, TEXT_HandyRefAppName
End Sub

Public Sub HandyRef_GetLatestVersion_Github_RibbonFun(ByVal control As IRibbonControl)
    On Error GoTo errHandle
       
    Shell "explorer.exe " & TEXT_HandyRefGithubUrl
    
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt Err
    
End Sub


Public Sub HandyRef_GetLatestVersion_Zhihu_RibbonFun(ByVal control As IRibbonControl)
    On Error GoTo errHandle
       
    Shell "explorer.exe " & TEXT_HandyRefZhihuUrl
    
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt Err
    
End Sub
