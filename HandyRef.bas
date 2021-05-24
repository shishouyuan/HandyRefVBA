'https://github.com/shishouyuan/HandyRefVBA

'A handy way to insert Cross Reference in MS Word
'Author: Shouyuan Shi @ South China University of Technology
'E-mail: shishouyuan@outlook.com
'Creating Date: 2021/5/11


'用于在Word里方便地添加交叉引用
'作者: 史守圆 @ 华南理工大学
'E-mail: shishouyuan@outlook.com
'创建时期: 2021/5/11


Const HandyRefVersion = "20210524.1259"

Const TEXT_HandyRefGithubUrl = "https://github.com/shishouyuan/HandyRefVBA"

Const BookmarkPrefix = "_HandyRef"
Const RefBrokenCommentTitle = "$HANDYREF_REFERENCE_BROKEN_COMMENT$"

#Const HandyRef_Lang = "en-us"

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


Public selectedBM As Bookmark
Public lastBMRefered As Boolean

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
    Application.UndoRecord.StartCustomRecord FormatUndoRecordText(TEXT_ActionName_CreateSource)
    On Error GoTo errHandle
    
    Dim rg As Range
    Set rg = Application.Selection.Range

     
    If rg.End - rg.Start = 0 Then
        MsgBox TEXT_CreateReferencePoint_NothingSelected, vbOKOnly + vbInformation, TEXT_HandyRefAppName
        GoTo exitSub
    End If
   
    If Not selectedBM Is Nothing Then
        If Not Application.IsObjectValid(selectedBM) Then
            Set selectedBM = Nothing    'set to Nothing when the bookmark is deleted by user
        ElseIf rg.IsEqual(selectedBM.Range) Then
            GoTo exitSub  'same range, thus the same bookmark remained
        Else
            If Not lastBMRefered Then
                selectedBM.Delete   'delete unreferenced bookmark
                Set selectedBM = Nothing
            End If
        End If
    End If

    Dim oldbm As Bookmark
    Dim bmi As Bookmark
    Dim bmShowHiddenOld As Boolean
    bmShowHiddenOld = rg.Bookmarks.ShowHidden
    
    'search for existing bookmark reference the same range
    rg.Bookmarks.ShowHidden = True
    For Each bmi In rg.Bookmarks
        If bmi.Range.IsEqual(rg) And bmi.Name Like BookmarkPrefix & "#*" Then
            Set oldbm = bmi
            Exit For
        End If
    Next bmi
    rg.Bookmarks.ShowHidden = bmShowHiddenOld
    
    If Not oldbm Is Nothing Then
        Set selectedBM = oldbm
        lastBMRefered = True
    Else
        'create new bookmark using timestamp as its name
        Set selectedBM = rg.Bookmarks.Add(BookmarkPrefix & CLngLng(Now * 1000000), rg)
        lastBMRefered = False
    End If
    
exitSub:
    Application.UndoRecord.EndCustomRecord
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt err
    GoTo exitSub
    
End Sub


Public Sub HandyRef_InsertCrossReferenceField_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_InsertCrossReferenceField
End Sub

Public Sub HandyRef_InsertCrossReferenceField()
    Application.UndoRecord.StartCustomRecord FormatUndoRecordText(TEXT_ActionName_InsertReference)
    On Error GoTo errHandle
    
    If Not selectedBM Is Nothing Then
        If Application.IsObjectValid(selectedBM) Then
            If selectedBM.Parent Is ActiveDocument Then
                ActiveDocument.Fields.Add Selection.Range, WdFieldType.wdFieldRef, selectedBM.Name & " \h"
                lastBMRefered = True
            Else
                MsgBox TEXT_InsertCrossReferenceField_CannotCrossFile, vbOKOnly + vbInformation, TEXT_HandyRefAppName
            End If
        Else
            Set selectedBM = Nothing
            GoTo noRefPointPrompt
        End If
    Else
noRefPointPrompt:
        MsgBox TEXT_InsertCrossReferenceField_NoRefPoint, vbOKOnly + vbInformation, TEXT_HandyRefAppName
    End If
    
exitSub:
    Application.UndoRecord.EndCustomRecord
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt err
    GoTo exitSub
    
End Sub

Public Sub HandyRef_ClearRefBrokenComment_RibbonFun(ByVal control As IRibbonControl)
    
    If Application.Selection.End - Application.Selection.Start = 0 Then
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
    ShowUnknowErrorPrompt err
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
    Application.UndoRecord.StartCustomRecord FormatUndoRecordText(TEXT_ActionName_CheckReference)
    On Error GoTo errHandle
    
    HandyRef_ClearRefBrokenComment checkingRange
    
    Static refRegExp As Object
    If refRegExp Is Nothing Then
        Set refRegExp = CreateObject("VBScript.RegExp")
        With refRegExp
            .Global = False
            .IgnoreCase = True
            .Pattern = "\s*REF\s+([^\s]+)\s*.*"
        End With
    End If
    
    Dim brokenCount As Integer
    brokenCount = 0
    
    Dim fd As Field
    Dim bmName As String
    Dim cmt As Comment
    For Each fd In checkingRange.Fields
        If fd.Type = wdFieldRef Then
            Set r = refRegExp.Execute(fd.Code.Text)
            If r.Count > 0 Then
                bmName = r.Item(0).SubMatches(0)
                If Not ActiveDocument.Bookmarks.Exists(bmName) Then
                
                    brokenCount = brokenCount + 1
                    
                    Set cmt = fd.Code.Comments.Add(fd.Code)
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
        End If
    Next fd
    
    If brokenCount = 0 Then
        MsgBox TEXT_NoBrokenRefFoundPrompt, vbOKOnly + vbInformation, TEXT_HandyRefAppName
    Else
        MsgBox Replace(TEXT_BrokenRefFoundPrompt, BrokenRefNumPosHolder, CStr(brokenCount)), vbOKOnly + vbInformation, TEXT_HandyRefAppName
    End If
    
exitSub:
    Application.UndoRecord.EndCustomRecord
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt err
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

Public Sub HandyRef_GetLatestVersion_RibbonFun(ByVal control As IRibbonControl)
    On Error GoTo errHandle
       
    Shell "explorer.exe " & TEXT_HandyRefGithubUrl
    
    Exit Sub
    
errHandle:
    ShowUnknowErrorPrompt err
    
End Sub

