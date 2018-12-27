# Use WPF controls in Office solutions
https://docs.microsoft.com/de-de/visualstudio/vsto/using-wpf-controls-in-office-solutions?view=vs-2017#host-wpf-controls-by-using-the-elementhost-class

# Excel specifications and limits
https://support.office.com/en-us/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3

````
Characters in a header or footer => 255
````

# excel interop

## vba


````
Option Explicit

Const PICT_VAAT = "\\dr-trm00\home\lip1dr3\07_logo\bosch_wortmarke_ohne_slogan.jpg"
Const cDepartment = "PADD/Facility Management"
Const cManager = "Heinz Böse"
Const cCompany = "ACME Company"
Const cAuthor = "Ludwig -s-chnitzel"

Sub Format_Header_Footer()
'
' Format_Header_Footer Macro
' Macro recorded 7/22/99 by Peter List
'
' revision: 2
' edit: 08/14/2000
' changes:
' changes hard-coded department value "PLPC" to document property "Department"
' the value is now changable through the file property access
'
    Dim sh As Object, _
        Title, Subject, Manager, Author, FontName, Company, Owner, Department, _
        LeftFooterString, LeftFooterStringCurrent, CenterFooterString, RightFooterString, _
        Sensitivity, SelectionType As String, _
        LenStrg As Integer

    If TypeName(Application.Selection) = "Nothing" Then
        SelectionType = "Worksheet"
    Else
        SelectionType = TypeName(Application.Selection.Parent)
    End If

    Title = ActiveWorkbook.BuiltinDocumentProperties("Title")
    Subject = ActiveWorkbook.BuiltinDocumentProperties("Subject")
    Author = ActiveWorkbook.BuiltinDocumentProperties("Author")
    Manager = ActiveWorkbook.BuiltinDocumentProperties("Manager")
    Company = ActiveWorkbook.BuiltinDocumentProperties("Company")
    Department = DocPropDepartment()
    Owner = DocPropOwner()
    Sensitivity = DocPropSensitivity()
    
    If Len(Author) = 0 Then
        Author = Application.UserName
        ActiveWorkbook.BuiltinDocumentProperties("Author") = Author
    End If
    
    Company = cCompany
    ActiveWorkbook.BuiltinDocumentProperties("Company") = Company
    
    
    'Define Fontname and Size
    FontName = "&""Tahoma""&6"

    'Building RightFooterString
    RightFooterString = FontName
    RightFooterString = RightFooterString & "File: &F" & vbCr & "Print Date: &D &T"

    'Building LeftFooterString
    LeftFooterString = ""

    LeftFooterString = LeftFooterString & FontName

    If Len(Company) > 0 Then
        LeftFooterString = LeftFooterString & _
                           "Org.Unit:" & Space(10 - Len("Company:")) & Company
    End If

    If Len(Company) > 0 And Len(Department) > 0 Then
        LeftFooterString = LeftFooterString & "; "
    End If

    If Len(Department) > 0 Then
        LeftFooterString = LeftFooterString & Department
    End If

    LeftFooterString = LeftFooterString & vbCr    'break line and new line

    If Len(Owner) > 0 Then
        LeftFooterString = LeftFooterString & _
                           "Owner:" & Space(10 - Len("Owner:")) & Owner & vbCr
    Else
        If Len(Author) > 0 Then
            LeftFooterString = LeftFooterString & _
                               "Author:" & Space(10 - Len("Author:")) & Author & vbCr
        End If
    End If

    If Len(Manager) > 0 Then
        LeftFooterString = LeftFooterString & _
                           "Manager:" & Space(10 - Len("Manager:")) & Manager & vbCr
    End If

    If Len(Sensitivity) > 0 Then
        LeftFooterString = LeftFooterString & _
                           "" & Sensitivity & ""    'document classification
    Else
        LeftFooterString = LeftFooterString & _
                           "This document has no specific confidentiality."
    End If


    'format each marked sheet
    For Each sh In ActiveWorkbook.Windows(1).SelectedSheets
        If SelectionType = "Chart" Then Set sh = Selection.Parent
        With sh.PageSetup
            CenterFooterString = .CenterFooter
            If Len(LeftFooterString) > (300 - Len(RightFooterString) - Len(CenterFooterString)) Then
                LeftFooterStringCurrent = Left(LeftFooterString, (300 - Len(RightFooterString) - Len(CenterFooterString))) & "..."
            Else
                LeftFooterStringCurrent = LeftFooterString
            End If

            Call InsertLeftHeaderPicture
            ' .CenterHeader = ""
            .LeftHeader = FontName & "&10&B" & Title & vbCr & "&B" & Subject & "&10"
            .LeftFooter = ""
            .LeftFooter = LeftFooterString
            '.CenterFooter = ""
            .RightFooter = RightFooterString
            .CenterFooter = FontName & "Page &P of &N"
            .ScaleWithDocHeaderFooter = False
            .AlignMarginsHeaderFooter = True
        End With
    Next
End Sub
Function DocPropOwner() As String
    Dim inputstring, Author, strSuggest As String
    On Error GoTo propertyError    'go to user dialog if no Owner Property exist




    DocPropOwner = ActiveWorkbook.CustomDocumentProperties("Owner")


    Exit Function

propertyError:

    'get Author from build in doc props
    Author = ActiveWorkbook.BuiltinDocumentProperties("Author")

    ' determine suggestion you will make
    If Len(Author) > 0 Then
        strSuggest = Author
    Else
        strSuggest = Application.UserName
    End If


    ' get user input or cancelation
    inputstring = InputBox("There is no owner of document currently." & vbCr _
                         & "You can define a owner by typing it in the next row." & vbCr _
                         & "If you push Cancel or do not plug in anything no owner will be stated." _
                , "Determine Document Owner", strSuggest)

    If Len(inputstring) > 0 Then
        ActiveWorkbook.CustomDocumentProperties.Add "Owner", False, msoPropertyTypeString, inputstring, False
        DocPropOwner = ActiveWorkbook.CustomDocumentProperties("Owner")
    Else
        DocPropOwner = ""
    End If
End Function
Function DocPropSensitivity() As String
    Dim inputstring As String
    On Error GoTo propertyError    'go to user dialog if no Owner Property exist
    DocPropSensitivity = ActiveWorkbook.CustomDocumentProperties("Sensitivity")

    Exit Function

propertyError:
    ' get user input or cancelation
    inputstring = InputBox("There is no Sensitivity of this Document stated." & vbCr _
                         & "You can define a owner by completing following sentence:" & vbCr _
                         & vbCr & "Confidentiality:", "Determine Document Sensitivity", "C-SC1")

    If Len(inputstring) > 0 Then
        ActiveWorkbook.CustomDocumentProperties.Add "Sensitivity", False, msoPropertyTypeString, inputstring, False
        DocPropSensitivity = ActiveWorkbook.CustomDocumentProperties("Sensitivity")
    Else
        DocPropSensitivity = ""
    End If

End Function
Function DocPropDepartment() As String
    Dim inputstring As String
    Dim propKey As String
    Dim propKeyDefVal As String    'default value for property key

    'determine property key & default value
    propKey = "Department"
    propKeyDefVal = cDepartment

    On Error GoTo propertyError    'go to user dialog if no Owner Property exist
    DocPropDepartment = ActiveWorkbook.CustomDocumentProperties(propKey)

    If Len(Trim(DocPropDepartment)) > 0 Then Exit Function

    inputstring = InputBox("There is no department of document filed currently." & vbCr _
                         & "You can define a department by typing it in the next row." & vbCr _
                         & "If you push Cancel or do not plug in anything then department display is empty." _
                , "Determine Document Department", propKeyDefVal)

    If Len(inputstring) > 0 Then
        ActiveWorkbook.CustomDocumentProperties.Add propKey, False, msoPropertyTypeString, inputstring, False
        DocPropDepartment = ActiveWorkbook.CustomDocumentProperties(propKey)
    Else
        DocPropDepartment = ""
    End If

    Exit Function

propertyError:
    ' get user input or cancelation
    inputstring = InputBox("There is no department of document filed currently." & vbCr _
                         & "You can define a department by typing it in the next row." & vbCr _
                         & "If you push Cancel or don't enter anything no department will be displayed." _
                , "Determine Document Department", propKeyDefVal)

    If Len(inputstring) > 0 Then
        ActiveWorkbook.CustomDocumentProperties.Add propKey, False, msoPropertyTypeString, inputstring, False
        DocPropDepartment = ActiveWorkbook.CustomDocumentProperties(propKey)
    Else
        DocPropDepartment = ""
    End If

End Function
Private Sub InsertLeftHeaderPicture()
    Dim myLH_Pic As Graphic

    Set myLH_Pic = ActiveSheet.PageSetup.RightHeaderPicture


    With myLH_Pic
        .FileName = PICT_VAAT
        .LockAspectRatio = msoTrue
        .Height = 25
        '.Width = 463.5
        '.Brightness = 0.36
        '.ColorType = msoPictureGrayscale
        '.Contrast = 0.39
        '.CropBottom = -14.4
        '.CropLeft = -28.8
        '.CropRight = -14.4
        '.CropTop = 21.6
    End With

    ' Enable the image to show up in the left header.
    ActiveSheet.PageSetup.RightHeader = "&G"

    Set myLH_Pic = Nothing

End Sub

Sub OpenDialogProperties()
    Dim dlganswer As Boolean
    dlganswer = Application.Dialogs(xlDialogProperties).Show
End Sub


Sub testuser()
    MsgBox "Current user is " & Application.UserName
End Sub
````


## XlBuiltInDialog Enum 
https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.xlbuiltindialog?view=excel-pia

Fields
_xlDialogChartSourceData 	541 	

Displays the dialog box described in the constant name.
_xlDialogPhonetic 	538 	

Displays the dialog box described in the constant name.
xlDialogActivate 	103 	

Displays the dialog box described in the constant name.
xlDialogActiveCellFont 	476 	

Displays the dialog box described in the constant name.
xlDialogAddChartAutoformat 	390 	

Displays the dialog box described in the constant name.
xlDialogAddinManager 	321 	

Displays the dialog box described in the constant name.
xlDialogAlignment 	43 	

Displays the dialog box described in the constant name.
xlDialogApplyNames 	133 	

Displays the dialog box described in the constant name.
xlDialogApplyStyle 	212 	

Displays the dialog box described in the constant name.
xlDialogAppMove 	170 	

Displays the dialog box described in the constant name.
xlDialogAppSize 	171 	

Displays the dialog box described in the constant name.
xlDialogArrangeAll 	12 	

Displays the dialog box described in the constant name.
xlDialogAssignToObject 	213 	

Displays the dialog box described in the constant name.
xlDialogAssignToTool 	293 	

Displays the dialog box described in the constant name.
xlDialogAttachText 	80 	

Displays the dialog box described in the constant name.
xlDialogAttachToolbars 	323 	

Displays the dialog box described in the constant name.
xlDialogAutoCorrect 	485 	

Displays the dialog box described in the constant name.
xlDialogAxes 	78 	

Displays the dialog box described in the constant name.
xlDialogBorder 	45 	

Displays the dialog box described in the constant name.
xlDialogCalculation 	32 	

Displays the dialog box described in the constant name.
xlDialogCellProtection 	46 	

Displays the dialog box described in the constant name.
xlDialogChangeLink 	166 	

Displays the dialog box described in the constant name.
xlDialogChartAddData 	392 	

Displays the dialog box described in the constant name.
xlDialogChartLocation 	527 	

Displays the dialog box described in the constant name.
xlDialogChartOptionsDataLabelMultiple 	724 	

Displays the dialog box described in the constant name.
xlDialogChartOptionsDataLabels 	505 	

Displays the dialog box described in the constant name.
xlDialogChartOptionsDataTable 	506 	

Displays the dialog box described in the constant name.
xlDialogChartSourceData 	540 	

Displays the dialog box described in the constant name.
xlDialogChartTrend 	350 	

Displays the dialog box described in the constant name.
xlDialogChartType 	526 	

Displays the dialog box described in the constant name.
xlDialogChartWizard 	288 	

Displays the dialog box described in the constant name.
xlDialogCheckboxProperties 	435 	

Displays the dialog box described in the constant name.
xlDialogClear 	52 	

Displays the dialog box described in the constant name.
xlDialogColorPalette 	161 	

Displays the dialog box described in the constant name.
xlDialogColumnWidth 	47 	

Displays the dialog box described in the constant name.
xlDialogCombination 	73 	

Displays the dialog box described in the constant name.
xlDialogConditionalFormatting 	583 	

Displays the dialog box described in the constant name.
xlDialogConsolidate 	191 	

Displays the dialog box described in the constant name.
xlDialogCopyChart 	147 	

Displays the dialog box described in the constant name.
xlDialogCopyPicture 	108 	

Displays the dialog box described in the constant name.
xlDialogCreateList 	796 	

Displays the dialog box described in the constant name.
xlDialogCreateNames 	62 	

Displays the dialog box described in the constant name.
xlDialogCreatePublisher 	217 	

Displays the dialog box described in the constant name.
xlDialogCreateRelationship 	1272 	

Displays the dialog box described in the constant name.
xlDialogCustomizeToolbar 	276 	

Displays the dialog box described in the constant name.
xlDialogCustomViews 	493 	

Displays the dialog box described in the constant name.
xlDialogDataDelete 	36 	

Displays the dialog box described in the constant name.
xlDialogDataLabel 	379 	

Displays the dialog box described in the constant name.
xlDialogDataLabelMultiple 	723 	

Displays the dialog box described in the constant name.
xlDialogDataSeries 	40 	

Displays the dialog box described in the constant name.
xlDialogDataValidation 	525 	

Displays the dialog box described in the constant name.
xlDialogDefineName 	61 	

Displays the dialog box described in the constant name.
xlDialogDefineStyle 	229 	

Displays the dialog box described in the constant name.
xlDialogDeleteFormat 	111 	

Displays the dialog box described in the constant name.
xlDialogDeleteName 	110 	

Displays the dialog box described in the constant name.
xlDialogDemote 	203 	

Displays the dialog box described in the constant name.
xlDialogDisplay 	27 	

Displays the dialog box described in the constant name.
xlDialogDocumentInspector 	862 	

Document Inspector dialog box
xlDialogEditboxProperties 	438 	

Displays the dialog box described in the constant name.
xlDialogEditColor 	223 	

Displays the dialog box described in the constant name.
xlDialogEditDelete 	54 	

Displays the dialog box described in the constant name.
xlDialogEditionOptions 	251 	

Displays the dialog box described in the constant name.
xlDialogEditSeries 	228 	

Displays the dialog box described in the constant name.
xlDialogErrorbarX 	463 	

Displays the dialog box described in the constant name.
xlDialogErrorbarY 	464 	

Displays the dialog box described in the constant name.
xlDialogErrorChecking 	732 	

Displays the dialog box described in the constant name.
xlDialogEvaluateFormula 	709 	

Displays the dialog box described in the constant name.
xlDialogExternalDataProperties 	530 	

Displays the dialog box described in the constant name.
xlDialogExtract 	35 	

Displays the dialog box described in the constant name.
xlDialogFileDelete 	6 	

Displays the dialog box described in the constant name.
xlDialogFileSharing 	481 	

Displays the dialog box described in the constant name.
xlDialogFillGroup 	200 	

Displays the dialog box described in the constant name.
xlDialogFillWorkgroup 	301 	

Displays the dialog box described in the constant name.
xlDialogFilter 	447 	

Displays the dialog box described in the constant name.
xlDialogFilterAdvanced 	370 	

Displays the dialog box described in the constant name.
xlDialogFindFile 	475 	

Displays the dialog box described in the constant name.
xlDialogFont 	26 	

Displays the dialog box described in the constant name.
xlDialogFontProperties 	381 	

Displays the dialog box described in the constant name.
xlDialogFormatAuto 	269 	

Displays the dialog box described in the constant name.
xlDialogFormatChart 	465 	

Displays the dialog box described in the constant name.
xlDialogFormatCharttype 	423 	

Displays the dialog box described in the constant name.
xlDialogFormatFont 	150 	

Displays the dialog box described in the constant name.
xlDialogFormatLegend 	88 	

Displays the dialog box described in the constant name.
xlDialogFormatMain 	225 	

Displays the dialog box described in the constant name.
xlDialogFormatMove 	128 	

Displays the dialog box described in the constant name.
xlDialogFormatNumber 	42 	

Displays the dialog box described in the constant name.
xlDialogFormatOverlay 	226 	

Displays the dialog box described in the constant name.
xlDialogFormatSize 	129 	

Displays the dialog box described in the constant name.
xlDialogFormatText 	89 	

Displays the dialog box described in the constant name.
xlDialogFormulaFind 	64 	

Displays the dialog box described in the constant name.
xlDialogFormulaGoto 	63 	

Displays the dialog box described in the constant name.
xlDialogFormulaReplace 	130 	

Displays the dialog box described in the constant name.
xlDialogFunctionWizard 	450 	

Displays the dialog box described in the constant name.
xlDialogGallery3dArea 	193 	

Displays the dialog box described in the constant name.
xlDialogGallery3dBar 	272 	

Displays the dialog box described in the constant name.
xlDialogGallery3dColumn 	194 	

Displays the dialog box described in the constant name.
xlDialogGallery3dLine 	195 	

Displays the dialog box described in the constant name.
xlDialogGallery3dPie 	196 	

Displays the dialog box described in the constant name.
xlDialogGallery3dSurface 	273 	

Displays the dialog box described in the constant name.
xlDialogGalleryArea 	67 	

Displays the dialog box described in the constant name.
xlDialogGalleryBar 	68 	

Displays the dialog box described in the constant name.
xlDialogGalleryColumn 	69 	

Displays the dialog box described in the constant name.
xlDialogGalleryCustom 	388 	

Displays the dialog box described in the constant name.
xlDialogGalleryDoughnut 	344 	

Displays the dialog box described in the constant name.
xlDialogGalleryLine 	70 	

Displays the dialog box described in the constant name.
xlDialogGalleryPie 	71 	

Displays the dialog box described in the constant name.
xlDialogGalleryRadar 	249 	

Displays the dialog box described in the constant name.
xlDialogGalleryScatter 	72 	

Displays the dialog box described in the constant name.
xlDialogGoalSeek 	198 	

Displays the dialog box described in the constant name.
xlDialogGridlines 	76 	

Displays the dialog box described in the constant name.
xlDialogImportTextFile 	666 	

Displays the dialog box described in the constant name.
xlDialogInsert 	55 	

Displays the dialog box described in the constant name.
xlDialogInsertHyperlink 	596 	

Displays the dialog box described in the constant name.
xlDialogInsertNameLabel 	496 	

Displays the dialog box described in the constant name.
xlDialogInsertObject 	259 	

Displays the dialog box described in the constant name.
xlDialogInsertPicture 	342 	

Displays the dialog box described in the constant name.
xlDialogInsertTitle 	380 	

Displays the dialog box described in the constant name.
xlDialogLabelProperties 	436 	

Displays the dialog box described in the constant name.
xlDialogListboxProperties 	437 	

Displays the dialog box described in the constant name.
xlDialogMacroOptions 	382 	

Displays the dialog box described in the constant name.
xlDialogMailEditMailer 	470 	

Displays the dialog box described in the constant name.
xlDialogMailLogon 	339 	

Displays the dialog box described in the constant name.
xlDialogMailNextLetter 	378 	

Displays the dialog box described in the constant name.
xlDialogMainChart 	85 	

Displays the dialog box described in the constant name.
xlDialogMainChartType 	185 	

Displays the dialog box described in the constant name.
xlDialogManageRelationships 	1271 	

Displays the dialog box described in the constant name.
xlDialogMenuEditor 	322 	

Displays the dialog box described in the constant name.
xlDialogMove 	262 	

Displays the dialog box described in the constant name.
xlDialogMyPermission 	834 	

Displays the dialog box described in the constant name.
xlDialogNameManager 	977 	

NameManager dialog box
xlDialogNew 	119 	

Displays the dialog box described in the constant name.
xlDialogNewName 	978 	

NewName dialog box
xlDialogNewWebQuery 	667 	

Displays the dialog box described in the constant name.
xlDialogNote 	154 	

Displays the dialog box described in the constant name.
xlDialogObjectProperties 	207 	

Displays the dialog box described in the constant name.
xlDialogObjectProtection 	214 	

Displays the dialog box described in the constant name.
xlDialogOpen 	1 	

Displays the dialog box described in the constant name.
xlDialogOpenLinks 	2 	

Displays the dialog box described in the constant name.
xlDialogOpenMail 	188 	

Displays the dialog box described in the constant name.
xlDialogOpenText 	441 	

Displays the dialog box described in the constant name.
xlDialogOptionsCalculation 	318 	

Displays the dialog box described in the constant name.
xlDialogOptionsChart 	325 	

Displays the dialog box described in the constant name.
xlDialogOptionsEdit 	319 	

Displays the dialog box described in the constant name.
xlDialogOptionsGeneral 	356 	

Displays the dialog box described in the constant name.
xlDialogOptionsListsAdd 	458 	

Displays the dialog box described in the constant name.
xlDialogOptionsME 	647 	

Displays the dialog box described in the constant name.
xlDialogOptionsTransition 	355 	

Displays the dialog box described in the constant name.
xlDialogOptionsView 	320 	

Displays the dialog box described in the constant name.
xlDialogOutline 	142 	

Displays the dialog box described in the constant name.
xlDialogOverlay 	86 	

Displays the dialog box described in the constant name.
xlDialogOverlayChartType 	186 	

Displays the dialog box described in the constant name.
xlDialogPageSetup 	7 	

Displays the dialog box described in the constant name.
xlDialogParse 	91 	

Displays the dialog box described in the constant name.
xlDialogPasteNames 	58 	

Displays the dialog box described in the constant name.
xlDialogPasteSpecial 	53 	

Displays the dialog box described in the constant name.
xlDialogPatterns 	84 	

Displays the dialog box described in the constant name.
xlDialogPermission 	832 	

Displays the dialog box described in the constant name.
xlDialogPhonetic 	656 	

Displays the dialog box described in the constant name.
xlDialogPivotCalculatedField 	570 	

Displays the dialog box described in the constant name.
xlDialogPivotCalculatedItem 	572 	

Displays the dialog box described in the constant name.
xlDialogPivotClientServerSet 	689 	

Displays the dialog box described in the constant name.
xlDialogPivotFieldGroup 	433 	

Displays the dialog box described in the constant name.
xlDialogPivotFieldProperties 	313 	

Displays the dialog box described in the constant name.
xlDialogPivotFieldUngroup 	434 	

Displays the dialog box described in the constant name.
xlDialogPivotShowPages 	421 	

Displays the dialog box described in the constant name.
xlDialogPivotSolveOrder 	568 	

Displays the dialog box described in the constant name.
xlDialogPivotTableOptions 	567 	

Displays the dialog box described in the constant name.
xlDialogPivotTableSlicerConnections 	1183 	

Reserved for internal use.
xlDialogPivotTableWhatIfAnalysisSettings 	1153 	

Reserved for internal use.
xlDialogPivotTableWizard 	312 	

Displays the dialog box described in the constant name.
xlDialogPlacement 	300 	

Displays the dialog box described in the constant name.
xlDialogPrint 	8 	

Displays the dialog box described in the constant name.
xlDialogPrinterSetup 	9 	

Displays the dialog box described in the constant name.
xlDialogPrintPreview 	222 	

Displays the dialog box described in the constant name.
xlDialogPromote 	202 	

Displays the dialog box described in the constant name.
xlDialogProperties 	474 	

Displays the dialog box described in the constant name.
xlDialogPropertyFields 	754 	

Displays the dialog box described in the constant name.
xlDialogProtectDocument 	28 	

Displays the dialog box described in the constant name.
xlDialogProtectSharing 	620 	

Displays the dialog box described in the constant name.
xlDialogPublishAsWebPage 	653 	

Displays the dialog box described in the constant name.
xlDialogPushbuttonProperties 	445 	

Displays the dialog box described in the constant name.
xlDialogRecommendedPivotTables 	1258 	

Displays the dialog box described in the constant name.
xlDialogReplaceFont 	134 	

Displays the dialog box described in the constant name.
xlDialogRoutingSlip 	336 	

Displays the dialog box described in the constant name.
xlDialogRowHeight 	127 	

Displays the dialog box described in the constant name.
xlDialogRun 	17 	

Displays the dialog box described in the constant name.
xlDialogSaveAs 	5 	

Displays the dialog box described in the constant name.
xlDialogSaveCopyAs 	456 	

Displays the dialog box described in the constant name.
xlDialogSaveNewObject 	208 	

Displays the dialog box described in the constant name.
xlDialogSaveWorkbook 	145 	

Displays the dialog box described in the constant name.
xlDialogSaveWorkspace 	285 	

Displays the dialog box described in the constant name.
xlDialogScale 	87 	

Displays the dialog box described in the constant name.
xlDialogScenarioAdd 	307 	

Displays the dialog box described in the constant name.
xlDialogScenarioCells 	305 	

Displays the dialog box described in the constant name.
xlDialogScenarioEdit 	308 	

Displays the dialog box described in the constant name.
xlDialogScenarioMerge 	473 	

Displays the dialog box described in the constant name.
xlDialogScenarioSummary 	311 	

Displays the dialog box described in the constant name.
xlDialogScrollbarProperties 	420 	

Displays the dialog box described in the constant name.
xlDialogSearch 	731 	

Displays the dialog box described in the constant name.
xlDialogSelectSpecial 	132 	

Displays the dialog box described in the constant name.
xlDialogSendMail 	189 	

Displays the dialog box described in the constant name.
xlDialogSeriesAxes 	460 	

Displays the dialog box described in the constant name.
xlDialogSeriesOptions 	557 	

Displays the dialog box described in the constant name.
xlDialogSeriesOrder 	466 	

Displays the dialog box described in the constant name.
xlDialogSeriesShape 	504 	

Displays the dialog box described in the constant name.
xlDialogSeriesX 	461 	

Displays the dialog box described in the constant name.
xlDialogSeriesY 	462 	

Displays the dialog box described in the constant name.
xlDialogSetBackgroundPicture 	509 	

Displays the dialog box described in the constant name.
xlDialogSetManager 	1109 	

Reserved for internal use.
xlDialogSetMDXEditor 	1208 	

Reserved for internal use.
xlDialogSetPrintTitles 	23 	

Displays the dialog box described in the constant name.
xlDialogSetTupleEditorOnColumns 	1108 	

Reserved for internal use.
xlDialogSetTupleEditorOnRows 	1107 	

Reserved for internal use.
xlDialogSetUpdateStatus 	159 	

Displays the dialog box described in the constant name.
xlDialogShowDetail 	204 	

Displays the dialog box described in the constant name.
xlDialogShowToolbar 	220 	

Displays the dialog box described in the constant name.
xlDialogSize 	261 	

Displays the dialog box described in the constant name.
xlDialogSlicerCreation 	1182 	

Reserved for internal use.
xlDialogSlicerPivotTableConnections 	1184 	

Reserved for internal use.
xlDialogSlicerSettings 	1179 	

Reserved for internal use.
xlDialogSort 	39 	

Displays the dialog box described in the constant name.
xlDialogSortSpecial 	192 	

Displays the dialog box described in the constant name.
xlDialogSparklineInsertColumn 	1134 	

Reserved for internal use.
xlDialogSparklineInsertLine 	1133 	

Reserved for internal use.
xlDialogSparklineInsertWinLoss 	1135 	

Reserved for internal use.
xlDialogSplit 	137 	

Displays the dialog box described in the constant name.
xlDialogStandardFont 	190 	

Displays the dialog box described in the constant name.
xlDialogStandardWidth 	472 	

Displays the dialog box described in the constant name.
xlDialogStyle 	44 	

Displays the dialog box described in the constant name.
xlDialogSubscribeTo 	218 	

Displays the dialog box described in the constant name.
xlDialogSubtotalCreate 	398 	

Displays the dialog box described in the constant name.
xlDialogSummaryInfo 	474 	

Displays the dialog box described in the constant name.
xlDialogTable 	41 	

Displays the dialog box described in the constant name.
xlDialogTabOrder 	394 	

Displays the dialog box described in the constant name.
xlDialogTextToColumns 	422 	

Displays the dialog box described in the constant name.
xlDialogUnhide 	94 	

Displays the dialog box described in the constant name.
xlDialogUpdateLink 	201 	

Displays the dialog box described in the constant name.
xlDialogVbaInsertFile 	328 	

Displays the dialog box described in the constant name.
xlDialogVbaMakeAddin 	478 	

Displays the dialog box described in the constant name.
xlDialogVbaProcedureDefinition 	330 	

Displays the dialog box described in the constant name.
xlDialogView3d 	197 	

Displays the dialog box described in the constant name.
xlDialogWebOptionsBrowsers 	773 	

Displays the dialog box described in the constant name.
xlDialogWebOptionsEncoding 	686 	

Displays the dialog box described in the constant name.
xlDialogWebOptionsFiles 	684 	

Displays the dialog box described in the constant name.
xlDialogWebOptionsFonts 	687 	

Displays the dialog box described in the constant name.
xlDialogWebOptionsGeneral 	683 	

Displays the dialog box described in the constant name.
xlDialogWebOptionsPictures 	685 	

Displays the dialog box described in the constant name.
xlDialogWindowMove 	14 	

Displays the dialog box described in the constant name.
xlDialogWindowSize 	13 	

Displays the dialog box described in the constant name.
xlDialogWorkbookAdd 	281 	

Displays the dialog box described in the constant name.
xlDialogWorkbookCopy 	283 	

Displays the dialog box described in the constant name.
xlDialogWorkbookInsert 	354 	

Displays the dialog box described in the constant name.
xlDialogWorkbookMove 	282 	

Displays the dialog box described in the constant name.
xlDialogWorkbookName 	386 	

Displays the dialog box described in the constant name.
xlDialogWorkbookNew 	302 	

Displays the dialog box described in the constant name.
xlDialogWorkbookOptions 	284 	

Displays the dialog box described in the constant name.
xlDialogWorkbookProtect 	417 	

Displays the dialog box described in the constant name.
xlDialogWorkbookTabSplit 	415 	

Displays the dialog box described in the constant name.
xlDialogWorkbookUnhide 	384 	

Displays the dialog box described in the constant name.
xlDialogWorkgroup 	199 	

Displays the dialog box described in the constant name.
xlDialogWorkspace 	95 	

Displays the dialog box described in the constant name.
xlDialogZoom 	256 	

Displays the dialog box described in the constant name.