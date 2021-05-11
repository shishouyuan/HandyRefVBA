'https://github.com/shishouyuan/HandyRefVBA
'A handy way to insert Cross Reference in MS Word
'Author: Shouyuan Shi @ South China University of Technology
'E-mail: shishouyuan@outlook.com
'Creating Date: 2021/5/11

'Usage:
'Step 1: Select the contents that needed to be referenced and run macro CreateReferencePoint.
'Step 2: Select the point you want to insert cross reference and run macro InsertCrossReferenceField.
'Tips: You can add keyboard shortcut to speed up the process. Search for "add keyboard shortcut for macro in word" to get help.

'用于在Word里方便地添加交叉引用
'作者: 史守圆 @ 华南理工大学
'E-mail: shishouyuan@outlook.com
'创建时期: 2021/5/11

'用法:
'步骤1：选中要被引用的内容然后执行宏 CreateReferencePoint。
'步骤2：选中想要插入交插引用的地方然后执行宏 InsertCrossReferenceField。
'提示: 可以通过给键盘快捷方式给这两个宏来提高操作速度. 搜索 "Word 给宏添加快捷键" 获取帮助。

Dim selectedBM As Bookmark
Dim lastBMRefered As Boolean

Sub CreateReferencePoint()
        
    Dim rg As Range
    Set rg = Selection.Range
    
    If Not selectedBM Is Nothing Then
        If Not IsObjectValid(selectedBM) Then
            Set selectedBM = Nothing    'set to Nothing when the bookmark is deleted by user
        ElseIf rg.IsEqual(selectedBM.Range) Then
            Exit Sub  'same range, thus the same bookmark returned
        Else
            If Not lastBMRefered Then
                selectedBM.Delete   'delete unreferenced bookmark
                Set selectedBM = Nothing
            End If
        End If
    End If
    
    Dim oldbm As Bookmark
    Dim bmi As Bookmark
    
    'search for existing bookmark reference the same range
    Bookmarks.ShowHidden = True
    For Each bmi In Bookmarks
        If bmi.Range.IsEqual(rg) Then
            Set oldbm = bmi
            Exit For
        End If
    Next bmi
    Bookmarks.ShowHidden = False
    
    If Not oldbm Is Nothing Then
        Set selectedBM = oldbm
        lastBMRefered = True
    Else
        'create new bookmark using timestamp as its name
        Set selectedBM = Bookmarks.Add("_SSYRef" & CLngLng(Now * 1000000), rg)
        lastBMRefered = False
    End If
    
End Sub

Sub InsertCrossReferenceField()
    If Not selectedBM Is Nothing Then
        Fields.Add Selection.Range, WdFieldType.wdFieldRef, selectedBM.Name
        lastBMRefered = True
    Else
        MsgBox "No Reference Point selected!", vbOKOnly, "Add Cross Reference"
    End If
End Sub