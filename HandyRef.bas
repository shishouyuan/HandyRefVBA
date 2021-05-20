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


Const HandyRefVersion = "20210520.0940"

Const TEXT_HandyRefGithubUrl = "https://github.com/shishouyuan/HandyRefVBA"



#Const HandyRef_Lang = "zh-cn"

#If HandyRef_Lang = "zh-cn" Then

    Const TEXT_HandyRefAppName = "HandyRef-快引"
    Const TEXT_HandyRefAuthor = "史守圆 @ 华南理工大学"
    Const TEXT_HandyRefDescription = "为 Word 提供一个快速添加交叉引用的方式。"
    Const TEXT_CreateReferencePoint_NothingSelected = "请先选中要引用的内容!"
    Const TEXT_InsertCrossReferenceField_NoRefPoint = "当前没有设置引用点!"
    Const TEXT_InsertCrossReferenceField_CannotCrossFile = "不支持跨文件引用!"
    Const TEXT_VersionPrompt = "版本："
    Const TEXT_NonCommecialPrompt = "仅限非商业用途"

#Else

    Const TEXT_HandyRefAppName = "HandyRef"
    Const TEXT_HandyRefAuthor = "Shouyuan Shi @ South China University of Technology"
    Const TEXT_HandyRefDescription = "Provide a handy way to insert Cross Reference in MS Word."
    Const TEXT_CreateReferencePoint_NothingSelected = "Nothing Selected!"
    Const TEXT_InsertCrossReferenceField_NoRefPoint = "No Reference Point selected!"
    Const TEXT_InsertCrossReferenceField_CannotCrossFile = "Cross file reference not supported!"
    Const TEXT_VersionPrompt = "Version: "
    Const TEXT_NonCommecialPrompt = "Only for NON-COMMERCIAL use."

#End If

Const BookmarkPrefix = "_HandyRef"
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


Public Sub HandyRef_About_RibbonFun(ByVal control As IRibbonControl)
    HandyRef_About
End Sub

Public Sub HandyRef_About()

    MsgBox TEXT_HandyRefAppName + vbCrLf + TEXT_HandyRefDescription + vbCrLf + TEXT_NonCommecialPrompt + vbCrLf + vbCrLf + TEXT_HandyRefAuthor + vbCrLf + TEXT_VersionPrompt + HandyRefVersion + vbCrLf + TEXT_HandyRefGithubUrl, vbOKOnly, TEXT_HandyRefAppName

End Sub

