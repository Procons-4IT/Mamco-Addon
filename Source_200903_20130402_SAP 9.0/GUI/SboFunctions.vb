Option Explicit On 
'*********************************************************************
'* TECHDEMO ADDON CLASS DESCRIPTION                                  *
'* assembly name:       TechDemoAddonBase                            *
'* classname:           NumbersInStr                                 *
'* classtype:           non abstract                                 *
'*********************************************************************
'* created by:          Lutz Morrien                                 *
'* company:             ocb GmbH, Ahaus                              *
'*                                                                   *
'* date of last change: 01-21-2004                                   *
'* last change by:      Lutz Morrien                                 *
'*                                                                   *
'*********************************************************************
'* Class description:                                                *
'* Class provides useful functions to read out fields including      *
'* descriptions (Might be obsolete after version 6.5).               *
'*                                                                   *
'*********************************************************************
'* list of changes and additions:                                    *
'*                                                                   *
'*********************************************************************
Public Class SBOFunctions
    Public Shared sDecimalSeparator As String = "."
    Public Shared sThousandsSeparator As String = ","
    Public Shared oNumberProvider As System.Globalization.NumberFormatInfo = New System.Globalization.NumberFormatInfo

    Public Shared Function Convert2SAPBool(ByVal Value As String) As String
        Dim blnFlag As Boolean
        Try
            blnFlag = CBool(Value)
            If blnFlag Then
                Return "Y"
            Else
                Return "N"
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Function
    Public Shared Function CleanNumberAsString(ByVal StringToFilter As String) As String
        ' This function takes a string as input and filters out all chars
        ' which do not appear in the string GoodChars
        ' Example: Input  "220,00 EUR"
        '          Output "220,00"

        Dim intCharCount As Integer
        Dim GoodChars As String
        Dim strResult As String

        GoodChars = "12345677890" + sDecimalSeparator

        'go through string
        For intCharCount = 0 To StringToFilter.Length - 1
            'if char at current location is a "good char",
            ' then move it to result string
            Dim sChar = StringToFilter.Chars(intCharCount)
            'If sChar = "." Then
            'strResult = strResult & ","
            'Else
            If GoodChars.IndexOf(sChar) >= 0 Then
                strResult = strResult & sChar
            End If
        Next
        If strResult <> "" Then
            Return strResult
        Else
            Return "0"
        End If
    End Function
    Public Shared Function CleanNumberAsDouble(ByVal MixedString As String) As Double
        Return CDbl(CleanNumberAsString(MixedString))
    End Function
    Public Shared Function CleanNumberAsSingle(ByVal MixedString As String) As Single
        Return CSng(CleanNumberAsString(MixedString))
    End Function
End Class
