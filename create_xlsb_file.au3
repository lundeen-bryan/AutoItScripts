#include <Excel.au3>
#include <MsgBoxConstants.au3>

; Create application object
Local $oExcel = _Excel_Open()
  If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookOpen Example", "Error creating the Excel application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

; Create a new workbook with only 1 worksheet
Local $sTabs = InputBox("Please enter a number", "How many tabs do you need?", "3")
If StringIsInt($sTabs) = 0 Then
  MsgBox($MB_SYSTEMMODAL, "Number of tabs", "The number of tabs must not be text. Please try again")
  $sTabs = InputBox("Please enter a number", "How many tabs do you need?", "3")
Else
  $sTabs = Int($sTabs)
EndIf
Local $oWorkbook = _Excel_BookNew($oExcel, Int($sTabs))
  If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookNew Example 1", "Error creating new workbook." & @CRLF & "@error = " & @error & ", @extended = " & @extended)

; Choose where to save the file
Local Const $sMessage = "Select a folder"

; Display an open dialog to select a file.
Local $sFileSelectFolder = FileSelectFolder($sMessage, "")
If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "No folder was selected.")
Else
        ; Display the selected folder.
        MsgBox($MB_SYSTEMMODAL, "", "You chose the following folder:" & @CRLF & $sFileSelectFolder)
EndIf


; Save the workbook as xlsb format
Local $sWorkbook = $sFileSelectFolder & "\index.xlsb"
_Excel_BookSaveAs($oWorkbook, $sWorkbook, $xlExcel12, True)
If @error Then Exit MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSaveAs Example 1", "Error saving workbook to '" & $sWorkbook & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
MsgBox($MB_SYSTEMMODAL, "Excel UDF: _Excel_BookSaveAs Example 1", "Workbook successfully saved as '" & $sWorkbook & "'.")

; #ABOUT# ==================================================================================
; Date: .............. 2022-07-02
; Title: ............. create_excel.au3
; Language: .......... English
; Website: ........... weburl
; Description: ....... Creates an Excel xlsb file in current directory
; Installs to: ....... 1486452/XL database
; Contact Author: .... lundeen-bryan
; Copyright © ........ n/a 2022. All rights reserved.
; Includes: .......... Excel.au3, MsgBoxConstants.au3
; Syntax: ............ n/a
; Parameters: ........ n/a
; Return values: ..... n/a
;
;
; ==========================================================================================
