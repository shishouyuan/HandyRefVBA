//https://github.com/shishouyuan/HandyRefVBA

//A handy way to insert Cross Reference in MS Word and WPS
//Author: Shouyuan Shi @ South China University of Technology
//E-mail: shishouyuan@outlook.com
//Creating Date: 2021/5/11


//用于在Word里方便地添加交叉引用
//作者: 史守圆 @ 华南理工大学
//E-mail: shishouyuan@outlook.com
//创建时期: 2021/5/11


var HandyRefVersion

var TEXT_HandyRefGithubUrl
var BookmarkPrefix
var RefBrokenCommentTitle
var HandyRef_Lang
var BrokenRefNumPosHolder

var TEXT_HandyRefAppName
var TEXT_HandyRefAuthor
var TEXT_HandyRefDescription
var TEXT_CreateReferencePoint_nullSelected
var TEXT_InsertCrossReferenceField_NoRefPoint
var TEXT_InsertCrossReferenceField_CannotCrossFile
var TEXT_VersionPrompt
var TEXT_NonCommecialPrompt
var TEXT_RefBrokenComment
var TEXT_BrokenRefFoundPrompt
var TEXT_NoBrokenRefFoundPrompt
var TEXT_RefBrokenCommentClearedPrompt
var TEXT_RefCheckingForWholeDocPrompt
var TEXT_ClearRefBrokenCommentForWholeDocPrompt
var TEXT_UnknownErrOccurredPrompt
var TEXT_ActionName_CreateSource
var TEXT_ActionName_InsertReference
var TEXT_ActionName_CheckReference
var TEXT_ActionName_ClearRefBrokenComment

function HandyRef_OnLoad(ribbonUI) {

    HandyRefVersion = "20210620.1556.JS"

    TEXT_HandyRefGithubUrl = "https://github.com/shishouyuan/HandyRefVBA"

    BookmarkPrefix = "_HandyRef"
    RefBrokenCommentTitle = "$HANDYREF_REFERENCE_BROKEN_COMMENT$"

    HandyRef_Lang = "zh-cn"

    BrokenRefNumPosHolder = "#"

    if (HandyRef_Lang == "zh-cn") {

        TEXT_HandyRefAppName = "HandyRef-快引"
        TEXT_HandyRefAuthor = "史守圆 @ 华南理工大学"
        TEXT_HandyRefDescription = "为 Word 提供一个快速添加交叉引用的方式。"
        TEXT_CreateReferencePoint_nullSelected = "请先选中要引用的内容!"
        TEXT_InsertCrossReferenceField_NoRefPoint = "当前没有设置引用点!"
        TEXT_InsertCrossReferenceField_CannotCrossFile = "不支持跨文件引用!"
        TEXT_VersionPrompt = "版本："
        TEXT_NonCommecialPrompt = "仅限非商业用途"
        TEXT_RefBrokenComment = "引用源丢失！"
        TEXT_BrokenRefFoundPrompt = "发现了 " + BrokenRefNumPosHolder + " 个损坏的引用，已为其添加批注。"
        TEXT_NoBrokenRefFoundPrompt = "没有发现损坏的引用。"
        TEXT_RefBrokenCommentClearedPrompt = "引用损坏批注已清除。"
        TEXT_RefCheckingForWholeDocPrompt = "当前没有选中的内容，将检查整个文档。" + '\r\n' + "这可能需要一些时间。"
        TEXT_ClearRefBrokenCommentForWholeDocPrompt = "当前没有选中的内容，将清除整个文档中的引用损坏批注。"
        TEXT_UnknownErrOccurredPrompt = "遇到错误："
        TEXT_ActionName_CreateSource = "创建引用源"
        TEXT_ActionName_InsertReference = "交叉引用"
        TEXT_ActionName_CheckReference = "检查引用"
        TEXT_ActionName_ClearRefBrokenComment = "清除批注"


    } else {

        TEXT_HandyRefAppName = "HandyRef"
        TEXT_HandyRefAuthor = "Shouyuan Shi @ South China University of Technology"
        TEXT_HandyRefDescription = "Provide a handy way to insert Cross Reference in MS Word."
        TEXT_CreateReferencePoint_nullSelected = "Nothing selected!"
        TEXT_InsertCrossReferenceField_NoRefPoint = "No Reference Point Selected!"
        TEXT_InsertCrossReferenceField_CannotCrossFile = "Cross file reference is ! supported!"
        TEXT_VersionPrompt = "Version: "
        TEXT_NonCommecialPrompt = "Only for NON-COMMERCIAL use."
        TEXT_RefBrokenComment = "Reference Broken!"
        TEXT_BrokenRefFoundPrompt = BrokenRefNumPosHolder + " broken reference found. Comments are attached."
        TEXT_NoBrokenRefFoundPrompt = "No broken reference found."
        TEXT_RefBrokenCommentClearedPrompt = "Reference broken comments cleared."
        TEXT_RefCheckingForWholeDocPrompt = "Nothing is selected. The whole document will be checked." + '\r\n' + "This may take a while."
        TEXT_ClearRefBrokenCommentForWholeDocPrompt = "Nothing is selected. Reference broken comments in the whole document will be cleared."
        TEXT_UnknownErrOccurredPrompt = "Error occurred:"
        TEXT_ActionName_CreateSource = "Create Source"
        TEXT_ActionName_InsertReference = "Insert Reference"
        TEXT_ActionName_CheckReference = "Check Reference"
        TEXT_ActionName_ClearRefBrokenComment = "Clear Comments"

    }
}


function HandyRef_GetEnabled(control) {
    return true
    //return ActiveDocument!=null
}


var selectedBM
var selectedRange
var selectedIsNote


function HandyRef_FormatUndoRecordText(s) {
    return s + "-" + TEXT_HandyRefAppName
}

function HandyRef_ShowUnknowErrorPrompt(e) {
    alert(TEXT_UnknownErrOccurredPrompt + '\r\n' + e.Description)
}

function HandyRef_CreateReferencePoint_RibbonFun(control) { // wrap the function to match the signature called by ribbion
    HandyRef_CreateReferencePoint()
}

function HandyRef_CreateReferencePoint() {
    var rg = Application.Selection.Range

    selectedIsNote = false
    selectedRange = rg
    selectedBM = null
    if (rg.Endnotes.Count == 0 && rg.Footnotes.Count == 1) {
        var fn = rg.Footnotes.Item(1)
        if (rg.InRange(fn.Range) || rg.InRange(fn.Reference) || !fn.Reference.InRange(rg)) {
            selectedIsNote = true
            selectedRange = fn.Reference
        }
    }
    else if (rg.Footnotes.Count == 0 && rg.Endnotes.Count == 1) {
        var en = rg.Endnotes.Item(1)
        if (rg.InRange(en.Range) || rg.InRange(en.Reference) || !en.Reference.InRange(rg)) {
            selectedIsNote = true
            selectedRange = en.Reference
        }
    }

    if (rg.End == rg.Start && !selectedIsNote) {
        selectedRange = null
        alert(TEXT_CreateReferencePoint_nullSelected)
    }

}


function HandyRef_InsertCrossReferenceField_RibbonFun(control) {
    HandyRef_InsertCrossReferenceField()
}



function HandyRef_InsertCrossReferenceField() {
    try {
        Application.UndoRecord.StartCustomRecord(HandyRef_FormatUndoRecordText(TEXT_ActionName_InsertReference))

        var bmValid = false
        if (selectedBM) {
            if (Application.IsObjectValid(selectedBM)) {
                if (selectedBM.Parent == ActiveDocument) {
                    bmValid = true
                }
                else {
                    alert(TEXT_InsertCrossReferenceField_CannotCrossFile)
                    return
                }
            }
            else {// it's possible the bookmark is deleted by the user, but the range remaind.
                selectedBM = null
            }
        }
        if (!bmValid) {
            if (!selectedRange || !Application.IsObjectValid(selectedRange) || selectedRange.Start == selectedRange.End) {
                selectedRange = null
                alert(TEXT_InsertCrossReferenceField_NoRefPoint)
                return
            }
            else if (selectedRange.Document != ActiveDocument) {
                alert(TEXT_InsertCrossReferenceField_CannotCrossFile)
                return
            }
            else {
                var oldbm// As Bookmark
                var bmShowHiddenOld = selectedRange.Bookmarks.ShowHidden

                //search for existing bookmark reference the same range
                var bmRegExp = new RegExp(BookmarkPrefix + "\\d+")
                selectedRange.Bookmarks.ShowHidden = true
                for (var i = 1; i <= selectedRange.Bookmarks.Count; i++) {
                    var bmi = selectedRange.Bookmarks.Item(i)
                    if (bmi.Range.IsEqual(selectedRange) && bmRegExp.test(bmi.Name)) {
                        oldbm = bmi
                        break
                    }
                }
                selectedRange.Bookmarks.ShowHidden = bmShowHiddenOld

                if (oldbm) {
                    selectedBM = oldbm
                }
                else {
                    //create new bookmark using timestamp as its name
                    selectedBM = selectedRange.Bookmarks.Add(BookmarkPrefix + new Date().getTime(), selectedRange)
                }
                bmValid = true

            }
        }
        if (bmValid) {
            if (selectedIsNote) {
                ActiveDocument.Fields.Add(Selection.Range, wdFieldNoteRef, selectedBM.Name + " \\h")
            }
            else {
                ActiveDocument.Fields.Add(Selection.Range, wdFieldRef, selectedBM.Name + " \\h")
            }
        }
    }
    catch (err) {
        HandyRef_ShowUnknowErrorPrompt(err.message)
    }
    finally {
        Application.UndoRecord.EndCustomRecord()
    }

}

function HandyRef_ClearRefBrokenComment_RibbonFun(control) {
    if (Application.Selection.End - Application.Selection.Start == 0) {
        alert(TEXT_ClearRefBrokenCommentForWholeDocPrompt)
        HandyRef_ClearRefBrokenComment(ActiveDocument.Range())
    }
    else {
        HandyRef_ClearRefBrokenComment(Application.Selection.Range)
    }
    alert(TEXT_RefBrokenCommentClearedPrompt)

}

function HandyRef_ClearRefBrokenComment(targetRange) {
    try {
        Application.UndoRecord.StartCustomRecord(HandyRef_FormatUndoRecordText(TEXT_ActionName_ClearRefBrokenComment))

        var toDelete = []
        for (var i = 1; i <= targetRange.Comments.Count; i++) {
            var cmt = targetRange.Comments.Item(i)
            var s = cmt.Range.Paragraphs.Last.Range.Text.trim()
            if (s.indexOf(RefBrokenCommentTitle) == s.length - RefBrokenCommentTitle.length) {
                toDelete.push(cmt)
            }

        }
        for (var i in toDelete) {
            toDelete[i].DeleteRecursively()
        }
    }
    catch (err) {
        HandyRef_ShowUnknowErrorPrompt(err.message)
    }
    finally {
        Application.UndoRecord.EndCustomRecord()
    }

}

function HandyRef_CheckForBrokenRef_RibbonFun(control) {

    if (Application.Selection.End - Application.Selection.Start == 0) {
        alert(TEXT_RefCheckingForWholeDocPrompt)
        HandyRef_CheckForBrokenRef(ActiveDocument.Range())
    }
    else {
        HandyRef_CheckForBrokenRef(Application.Selection.Range)
    }

}

function HandyRef_CheckForBrokenRef(checkingRange) {
    var oldScreenUpdating = Application.ScreenUpdating
    try {
        Application.ScreenUpdating = false
        Application.UndoRecord.StartCustomRecord(HandyRef_FormatUndoRecordText(TEXT_ActionName_CheckReference))
        HandyRef_ClearRefBrokenComment(checkingRange)

        //var refRegExp = /^\s*(?:NOTE)?REF.*?(?<!\\\*)\s+([^\s\\]+).*/i
        var refRegExp = /^\s*(NOTE){0,1}REF.*\s([^\s\\]+).*/i
        var refRegExp0 = /\\[*@#]\s*[^\s\\]*/g

        var brokenCount = 0

        for (var i = 1; i <= checkingRange.Fields.Count; i++) {
            var fd = checkingRange.Fields.Item(i)

            if (fd.Type == wdFieldRef || fd.Type == wdFieldNoteRef) {
                r = refRegExp.exec(fd.Code.Text.replace(refRegExp0, ""))
                var isBroken = true
                if (r.length > 0) {
                    var bmName = r[2]
                    if (ActiveDocument.Bookmarks.Exists(bmName)) {
                        isBroken = false
                    }
                }
                if (isBroken) {
                    brokenCount = brokenCount + 1

                    var cmt = fd.Code.Comments.Add(fd.Code)
                    var t = cmt.Range
                    t.InsertAfter(TEXT_RefBrokenComment)
                    t.InsertParagraphAfter()
                    t.InsertAfter(RefBrokenCommentTitle)

                    t = cmt.Range.Paragraphs.First.Range
                    t.Bold = true
                    t.HighlightColorIndex = wdYellow
                }
            }
        }

        if (brokenCount == 0) {
            alert(TEXT_NoBrokenRefFoundPrompt)
        }
        else {
            alert(TEXT_BrokenRefFoundPrompt.replace(BrokenRefNumPosHolder, brokenCount))
            try {
                ActiveWindow.View.SplitSpecial = wdPaneNone
            }
            catch (err) { }
            ActiveWindow.View.SplitSpecial = wdPaneRevisions
        }
    }
    catch (err) {
        HandyRef_ShowUnknowErrorPrompt(err.message)
    }
    finally {
        Application.ScreenUpdating = oldScreenUpdating
        Application.UndoRecord.EndCustomRecord()
    }

}

function HandyRef_About_RibbonFun(control) {
    HandyRef_About()
}

function HandyRef_About() {
    alert(TEXT_HandyRefAppName + '\r\n' + TEXT_HandyRefDescription + '\r\n' + TEXT_NonCommecialPrompt + '\r\n\r\n' + TEXT_VersionPrompt + HandyRefVersion + '\r\n' + TEXT_HandyRefAuthor + '\r\n' + TEXT_HandyRefGithubUrl)//)
}

function HandyRef_GetLatestVersion_RibbonFun(control) {
    try {
        Shell("explorer.exe " + TEXT_HandyRefGithubUrl, jsNormalFocus)
    }
    catch (err) {
        HandyRef_ShowUnknowErrorPrompt(err)
    }
}

