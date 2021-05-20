'https://github.com/shishouyuan/HandyRefVBA
'A handy way to insert Cross Reference in MS Word
'Author: Shouyuan Shi @ South China University of Technology
'E-mail: shishouyuan@outlook.com
'Creating Date: 2021/5/11

'Usage:
'Step 1: Select the contents that needed to be referenced and run macro CreateReferencePoint.
'Step 2: Select the point you want to insert cross reference and run macro InsertCrossReferenceField.

'用于在Word里方便地添加交叉引用
'作者: 史守圆 @ 华南理工大学
'E-mail: shishouyuan@outlook.com
'创建时期: 2021/5/11

'用法:
'步骤1：选中要被引用的内容然后执行宏 CreateReferencePoint。
'步骤2：选中想要插入交插引用的地方然后执行宏 InsertCrossReferenceField。


Const HandyRefVersion = "20210520.1434"

Const TEXT_HandyRefGithubUrl = "https://github.com/shishouyuan/HandyRefVBA"



Const BookmarkPrefix = "_HandyRef"
Const RefBrokenCommentTitle = "$HANDYREF_REFERENCE_BROKEN_COMMENT$"

#Const HandyRef_Lang = "en-us"

Const BrokenRefNumPosHolder = "#"  '数量占位符
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
    Const TEXT_NoBrokenRefFoundPrompt = "没有发现损坏的索引。"
    Const TEXT_RefBrokenCommentClearedPrompt = "引用损坏批注已清除。"

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
    Const TEXT_BrokenRefFoundPrompt = BrokenRefNumPosHolder & " broken reference found, and comments are attached."
    Const TEXT_NoBrokenRefFoundPrompt = "No broken reference found."
    Const TEXT_RefBrokenCommentClearedPrompt = "Reference broken comments cleared."
    
#End If


Public selectedBM As Bookmark
Public lastBMRefered As Boolean


Public Sub HandyRef_CreateReferencePoint_RibbonFun(ByVal control As IRibbonControl) ' wrap the function to match the signature called by ribbion
    HandyRef_CreateReferencePoint
End Sub

Public Sub HandyRef_CreateReferencePoint()
    
    Dim rg As Range
    Set rg = Application.Selection.Range

     
    If rg.End - rg.Start = 0 Then
        MsgBox TEXT_CreateReferencePoint_NothingSelected, vbOKOnly, TEXT_HandyRefAppName
        Exit Sub
    End If
   
    If Not selectedBM Is Nothing Then
        If Not Application.IsObjectValid(selectedBM) Then
            Set selectedBM = Nothing    'set to Nothing when the bookmark is deleted by user
        ElseIf rg.IsEqual(selectedBM.Range) Then
            Exit Sub  'same range, thus the same bookmark remained
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
    
End Sub


Public Sub HandyRef_InsertCrossReferenceField_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_InsertCrossReferenceField
End Sub

Public Sub HandyRef_InsertCrossReferenceField()
    If Not selectedBM Is Nothing Then
        If Application.IsObjectValid(selectedBM) Then
            If selectedBM.Parent Is ActiveDocument Then
                ActiveDocument.Fields.Add Selection.Range, WdFieldType.wdFieldRef, selectedBM.Name
                lastBMRefered = True
            Else
                MsgBox TEXT_InsertCrossReferenceField_CannotCrossFile, vbOKOnly, TEXT_HandyRefAppName
            End If
        Else
            Set selectedBM = Nothing
            GoTo noRefPointPrompt
        End If
    Else
noRefPointPrompt:
        MsgBox TEXT_InsertCrossReferenceField_NoRefPoint, vbOKOnly, TEXT_HandyRefAppName
    End If
End Sub

Public Sub HandyRef_ClearRefBrokenComment_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_ClearRefBrokenComment
    MsgBox TEXT_RefBrokenCommentClearedPrompt, vbOKOnly, TEXT_HandyRefAppName
End Sub

Public Sub HandyRef_ClearRefBrokenComment()
    Dim cmt As Comment
    Dim s As String
    For Each cmt In ActiveDocument.Comments
        s = cmt.Range.Paragraphs.Last.Range.Text
        s = Replace(s, vbCr, "")
        s = Replace(s, vbLf, "")
        If StrComp(s, RefBrokenCommentTitle) = 0 Then
            cmt.DeleteRecursively
        End If
    Next cmt
End Sub
 
Public Sub HandyRef_CheckForBrokenRef_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_CheckForBrokenRef
End Sub

Public Sub HandyRef_CheckForBrokenRef()

    HandyRef_ClearRefBrokenComment
    
    Dim refRegExp As Object
    Set refRegExp = CreateObject("VBScript.RegExp")
    With refRegExp
        .Global = False
        .IgnoreCase = True
        .Pattern = "\s*REF\s+([^\s]+)\s*.*"
    End With
    
    Dim brokenCount As Integer
    brokenCount = 0
    
    Dim fd As Field
    Dim bmName As String
    Dim cmt As Comment
    For Each fd In ActiveDocument.Fields
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
        MsgBox TEXT_NoBrokenRefFoundPrompt, vbOKOnly, TEXT_HandyRefAppName
    Else
        MsgBox Replace(TEXT_BrokenRefFoundPrompt, BrokenRefNumPosHolder, CStr(brokenCount)), vbOKOnly, TEXT_HandyRefAppName
    End If
    
End Sub



Public Sub HandyRef_About_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_About
End Sub

Public Sub HandyRef_About()

    MsgBox TEXT_HandyRefAppName + vbCrLf + TEXT_HandyRefDescription + vbCrLf + TEXT_NonCommecialPrompt + vbCrLf + vbCrLf + TEXT_HandyRefAuthor + vbCrLf + TEXT_VersionPrompt + HandyRefVersion + vbCrLf + TEXT_HandyRefGithubUrl, vbOKOnly, TEXT_HandyRefAppName

End Sub


