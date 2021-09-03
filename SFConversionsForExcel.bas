Attribute VB_Name = "SFConversionsForExcel"
' ======================================================================================
' @file     SFConversionsForExcel
' @ingroup  Excel Addins
' @author   Jack D. Pond (jack.pond@psitex.com)
'
' Date:     2021-09-01
'
' @copyright  2009-2021 Jack D. Pond
' @description Some commonly useful Salesforce Conversions
'   SFConvertId15to18 - Convert Salesforce 15 character ID to 18 character Case Safe Id (equivalent of SF CASESAFEID())
'   SFConvertId18to15 - Convert Salesforce 18 character to 18 Case Safe Id to 15 Character Id
'
' ======================================================================================
Public Const SF_ALL_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ012345"
Public Const SF_UPPER_CHARS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


Public Function SFConvertId15to18(ShortId As String) As String
Attribute SFConvertId15to18.VB_Description = "Convert a Salesforce 18 character Case Safe Id to 15 Character Id"
    If Len(ShortId) <> 15 Then
        SFConvertId15to18 = "Short ID should be 15 characters"
        Err.Raise 500, , SFConvertId15to18
        Exit Function
    End If
    
    Dim idExtension As String
    Dim i As Integer, thisCharVal As Integer
    thisCharVal = 0
    For i = 15 To 1 Step -1
        thisCharVal = (2 * thisCharVal) + Sgn(InStr(1, SF_UPPER_CHARS, Mid(ShortId, i, 1), vbBinaryCompare))
        If i Mod 5 = 1 Then
            idExtension = Mid(SF_ALL_CHARS, thisCharVal + 1, 1) + idExtension
            thisCharVal = 0
        End If
    Next i
    SFConvertId15to18 = ShortId + idExtension
End Function
Public Function SFConvertId18to15(LongId As String) As String
Attribute SFConvertId18to15.VB_Description = "Convert Salesforce 18 character Case Safe Id to 15 character ID"
    If Len(ShortId) <> 18 Then
        SFConvertId18to15 = "Long ID should be 18 characters"
        Err.Raise 500, , SFConvertId18to15
        Exit Function
    End If
    SFConvertId18to15 = Left(LongId, 15)
End Function
