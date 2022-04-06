Imports Microsoft.Office.Interop

Public Class ExportExcel

    Public Sub exl2pdfOnly(ByRef excelBook1 As Microsoft.Office.Interop.Excel.Workbook, ByVal path As String, ByVal File As String)
        Try
            excelBook1.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, path & "\" & File,
                                       Excel.XlFixedFormatQuality.xlQualityStandard, True, False, 1, 1, False)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Export Excel Error!")
        End Try

    End Sub

    Public Sub exl2pdf(ByRef excelBook1 As Microsoft.Office.Interop.Excel.Workbook, ByVal path As String, ByVal File As String, ByRef closeApp As Microsoft.Office.Interop.Excel.ApplicationClass)
        Try
            excelBook1.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, path & "\" & File,
                                       Excel.XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, From:=1, [To]:=1, OpenAfterPublish:=False)
            excelBook1.Close(SaveChanges:=False)
            closeApp.Quit()
            closeApp = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Export Excel Error!")
        End Try

    End Sub

    ''' <summary>
    ''' Save the First Sheet of the Excel file as PDF document
    ''' </summary>
    ''' <param name="CompleteFilePath">Complete file path of Excel including file name and Extension</param>
    ''' <param name="SaveFilePath">Destination File with file path and File name. No need of extension.</param>
    ''' <param name="errMsg">Returns any error message as string</param>
    Public Sub ExcelSheet1_SaveAsPdf(ByVal CompleteFilePath As String, ByVal SaveFilePath As String, ByRef errMsg As String)
        Try
            Dim excelApp1 As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Dim excelBook1 As Microsoft.Office.Interop.Excel.Workbook
            excelBook1 = excelApp1.Workbooks.Open(CompleteFilePath, 0, [ReadOnly]:=False, 5,
            System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            False, System.Reflection.Missing.Value, System.Reflection.Missing.Value,
            True, False, System.Reflection.Missing.Value, False)
            Dim excelSheets1 As Microsoft.Office.Interop.Excel.Sheets = excelBook1.Sheets
            Dim wSheet1 As Microsoft.Office.Interop.Excel.Worksheet = CType(excelSheets1(1), Excel.Worksheet)
            excelBook1.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, SaveFilePath,
                                       Excel.XlFixedFormatQuality.xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, From:=1, [To]:=1, OpenAfterPublish:=False)
            excelBook1.Close(SaveChanges:=False)
            excelApp1.Quit()
            excelApp1 = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Export Excel Error!")
        End Try

    End Sub

End Class
