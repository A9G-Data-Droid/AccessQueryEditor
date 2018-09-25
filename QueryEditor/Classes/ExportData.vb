Imports System.IO
Imports System.Text
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Microsoft.Office.Interop.Excel

''' <summary>Contains different manners to export a DataTable to an Excel Sheet</summary>
Public Class ExportData

#Region "Export By Execl"

    ''' <summary>
    '''     Exports a DataTable to an Excel Sheet using the Excel Application itself
    ''' </summary>
    ''' <param name="fileName">the destination file</param>
    ''' <param name="exportDataTable">the data to export</param>
    Public Shared Function ExportByExcel(fileName As String, exportDataTable As Data.DataTable) As Boolean
        Dim dataColumns As Integer = exportDataTable.Columns.Count

        If exportDataTable Is Nothing Or dataColumns = 0 Then
            Debug.WriteLine("Attempted to export empty table?")
            Return False
        End If

        ' Get headers
        Dim header(dataColumns - 1) as Object
        For i = 0 To dataColumns - 1
            header(i) = exportDataTable.Columns(i).ColumnName
        Next

        ' Get Data
        Dim dataRows = exportDataTable.Rows.Count
        Dim excelCells(dataRows - 1, dataColumns - 1) As Object
        For row = 0 To dataRows - 1
            For column = 0 To dataColumns - 1
                excelCells(row, column) = exportDataTable.Rows(row)(column).ToString
            Next
        Next

        Dim excelExport As New Application With {
                .Visible = False,
                .DisplayAlerts = False
                } ' This opens excel

        Try 'to make Excel sheet
            excelExport.Workbooks.Add
            Dim excelSheet As Worksheet = excelExport.ActiveSheet
            excelSheet.Name = exportDataTable.TableName

            ' Write headers
            Dim headerRange As Range = excelSheet.Range(excelSheet.Cells(1, 1), excelSheet.Cells(1, dataColumns))
            HeaderRange.Value = Header
            ' Write the data!
            excelSheet.Range(excelSheet.Cells(2, 1), excelSheet.Cells(dataRows + 1, dataColumns)).Value = excelCells
            ' Create Table
            dim excelTable As ListObject = excelSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
                                                                      excelSheet.UsedRange, Type.Missing, XlYesNoGuess.xlYes)
            excelTable.TableStyle = "TableStyleDark1"

            If fileName IsNot Nothing And fileName IsNot String.Empty Then excelSheet.SaveAs(fileName)

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            excelExport.Visible = True
            
            'If excelBook IsNot Nothing Then excelBook.Close()
            'excelExport.Quit()
        End Try

        Return True
    End Function

    'Private Shared Sub PopulateSheet(dt As Data.DataTable, oCells As Range)
    '    Dim dRow As DataRow
    '    Dim dataArray() As Object
    '    Dim count As Integer
    '    Dim columnCount As Integer

    '    'Output Columns Headers

    '    For columnCount = 0 To dt.Columns.Count - 1
    '        oCells(1, columnCount + 1) = dt.Columns(columnCount).ToString
    '        With CType(oCells(1, columnCount + 1), Range)
    '            .Interior.ColorIndex = 6
    '            .Interior.Pattern = 1

    '            .Font.Name = "Arial"
    '            .Font.Size = 12
    '            .Font.Bold = True
    '            .Font.ColorIndex = 5

    '        End With
    '    Next

    '    'Output Data
    '    For count = 0 To dt.Rows.Count - 1
    '        dRow = dt.Rows.Item(count)
    '        dataArray = dRow.ItemArray
    '        For columnCount = 0 To dt.Columns.Count - 1
    '            oCells(count + 2, columnCount + 1) = dataArray(columnCount).ToString
    '        Next

    '        If count Mod 20 = 0 Then
    '            'just to make a little bit rest for the system
    '            ' and this condition is to not take long time for this loop
    '            System.Windows.Forms.Application.DoEvents()
    '        End If

    '    Next
    'End Sub

#End Region

#Region "Export By Html"

    ''' <summary>
    '''     Exports a DataTable to HTML
    ''' </summary>
    ''' <param name="fileName">the destination file</param>
    ''' <param name="dt">the data to export</param>
    Public Shared Function ExportByHtml(fileName As String, dt As Data.DataTable) As Boolean
        Dim grid As New DataGrid
        With grid
            .HeaderStyle.Font.Bold = True
            .DataSource = dt
            .GridLines = WebControls.GridLines.Both
            .DataMember = dt.TableName
            .DataBind()
        End With

        Try 'to make a file
            Using sw As New StreamWriter(fileName)
                Using hw As New HtmlTextWriter(sw)
                    grid.RenderControl(hw)
                End Using
            End Using

            Return True
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Return False
        End Try
    End Function

#End Region

#Region "Export By XML"

    ''' <summary>
    '''     exports a DataTable to an Excel Sheet using "XML SpreadSheet" formatting
    '''     to create an Excel-Distinguished formatted file
    ''' </summary>
    ''' <param name="fileName">the destination file</param>
    ''' <param name="dt">the data to export</param>
    Public Shared Function ExportByXml(fileName As String, dt As Data.DataTable) As Boolean
        Dim rowCount = 0
        Dim sheetCount = 1
        Dim startExcelXml As New StringBuilder(1024)
        With startExcelXml
            .Append("<?xml version=""1.0""?>")
            .Append("<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""")
            .Append(" xmlns:o=""urn:schemas-microsoft-com:office:office""")
            .Append(" xmlns:x=""urn:schemas-    microsoft-com:office:excel""")
            .Append(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">")
            .Append(" <Styles>")
            .Append(
                "	<Style ss:ID=""Default"" ss:Name=""Normal"">  <Alignment ss:Vertical=""Bottom""/> <Font/> <Interior/> <NumberFormat/> <Protection/> </Style>")
            .Append(
                " 	<Style ss:ID=""BoldColumn"">  <Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/>   <Font x:Family=""Swiss"" ss:Bold=""1""/>  </Style>")
            .Append("	<Style ss:ID=""StringLiteral""> <NumberFormat ss:Format=""@""/> </Style>")
            .Append("	<Style ss:ID=""Decimal""> <NumberFormat ss:Format=""0.0000""/> </Style>")
            .Append("	<Style ss:ID=""Integer""> <NumberFormat ss:Format=""0""/> </Style>")
            .Append("	<Style ss:ID=""DateLiteral""> <NumberFormat ss:Format=""mm/dd/yyyy;@""/> </Style>")
            .Append(" </Styles>" & Environment.NewLine)
            .Append("<Worksheet ss:Name=""Output-Sheet" & sheetCount.ToString & """>")
        End With

        Const endExcelXml = "</Workbook>"
        Try 'to make a file
            Using excelDoc = New StreamWriter(fileName)
                With excelDoc
                    .WriteLine(startExcelXML.ToString)
                    .WriteLine("<Table>")
                    .WriteLine("<Row>")
                    For Each col As DataColumn In dt.Columns
                        .Write(ControlChars.Tab)
                        .Write("<Cell ss:StyleID=""BoldColumn""><Data ss:Type=""String"">")
                        .Write(Col.ColumnName)
                        .WriteLine("</Data></Cell>")
                    Next

                    .WriteLine("</Row>")
                    For Each row As DataRow In dt.Rows
                        rowCount += 1
                        'if the number of rows is > 64000 create a new page to continue output
                        If rowCount = 64000 Then
                            rowCount = 0
                            sheetCount += 1
                            .Write("</Table>")
                            .Write(" </Worksheet>")
                            .Write("<Worksheet ss:Name=""Sheet" + sheetCount + """>")
                            .Write("<Table>")
                        End If

                        .WriteLine("<Row>")
                        For ColI = 0 To dt.Columns.Count - 1
                            .Write(ControlChars.Tab)
                            Select Case dt.Columns(ColI).DataType.Name.ToUpper
                                Case GetType(String).Name.ToUpper
                                    Dim xmlRow = Row(ColI).ToString()
                                    xmlRow = xmlRow.Trim()
                                    xmlRow = xmlRow.Replace("&", "&")
                                    xmlRow = xmlRow.Replace(">", ">")
                                    xmlRow = xmlRow.Replace("<", "<")
                                    .Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                    .Write(xmlRow)
                                    .WriteLine("</Data></Cell>")
                                Case GetType(Date).Name.ToUpper
                                    'Excel has a specific Date Format of YYYY-MM-DD followed by  
                                    'the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                                    'The Following Code puts the date stored in XMLDate 
                                    'to the format above
                                    Dim xmlDate = CType(Row(ColI), DateTime)
                                    Dim xmlDateToString as String 'Excel Converted Date
                                    xmlDateToString = XMLDate.Year.ToString() & "-" &
                                                      If(XMLDate.Month < 10, "0" & XMLDate.Month.ToString(),
                                                         XMLDate.Month.ToString()) & "-" &
                                                      If(XMLDate.Day < 10, "0" & XMLDate.Day.ToString(),
                                                         XMLDate.Day.ToString()) & "T" &
                                                      If(XMLDate.Hour < 10, "0" & XMLDate.Hour.ToString(),
                                                         XMLDate.Hour.ToString()) & ":" &
                                                      If(XMLDate.Minute < 10, "0" & XMLDate.Minute.ToString(),
                                                         XMLDate.Minute.ToString()) & ":" &
                                                      If(XMLDate.Second < 10, "0" & XMLDate.Second.ToString(),
                                                         XMLDate.Second.ToString()) & ".000"

                                    .Write("<Cell ss:StyleID=""DateLiteral""><Data ss:Type=""DateTime"">")
                                    .Write(xmlDateToString)
                                    .WriteLine("</Data></Cell>")
                                Case "System.Boolean".ToUpper
                                    .Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                    .Write(Row(ColI).ToString())
                                    .WriteLine("</Data></Cell>")

                                Case GetType(Byte).Name.ToUpper, GetType(Int16).Name.ToUpper,
                                    GetType(Int32).Name.ToUpper, GetType(Int64).Name.ToUpper,
                                    GetType(Integer).Name.ToUpper, GetType(Long).Name.ToUpper

                                    .Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                                    .Write(Row(ColI).ToString())
                                    .WriteLine("</Data></Cell>")
                                Case GetType(Decimal).Name.ToUpper, GetType(Single).Name.ToUpper,
                                    GetType(Double).Name.ToUpper

                                    .Write("<Cell ss:StyleID=""Decimal""><Data ss:Type=""Number"">")
                                    .Write(Row(ColI).ToString())
                                    .WriteLine("</Data></Cell>")
                                Case GetType(DBNull).Name.ToUpper
                                    .Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                    .Write("")
                                    .WriteLine("</Data></Cell>")
                                Case "Byte[]".ToUpper
                                    .Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                    .Write("System.Byte[]")
                                    .WriteLine("</Data></Cell>")
                                Case Else
                                    .Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                                    .Write("ERROR HAPPENED")
                                    .WriteLine("</Data></Cell>")
                            End Select
                        Next
                        .WriteLine("</Row>")
                    Next
                    .WriteLine("</Table>")
                    .Write(" </Worksheet>")
                    .Write(endExcelXML)
                    .Close()
                End With
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region
End Class
