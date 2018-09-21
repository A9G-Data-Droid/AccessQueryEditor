Imports Microsoft.Office.Interop.Excel
Imports System.Web.UI
Imports System.IO

''' <summary>Contains different manners to export a DataTable to an Excel Sheet</summary>
Public Class Excel

#Region "Export By Execl"
    ''' <summary>
    ''' exports a DataTable to an Excel Sheet using the Excel Application itself
    ''' </summary>
    ''' <param name="fileName">the destination file</param>
    ''' <param name="dt">the data to export</param>
    Public Shared Function ExportByExecl(ByVal fileName As String, ByVal dt As Data.DataTable) As Boolean
        Dim excelExport As New Microsoft.Office.Interop.Excel.Application
        Dim excelBook As Workbook
        Dim excelSheets As Sheets
        Dim excelSheet As Worksheet
        Dim excelCells As Range


        Try
            'next are to avoid the "Old format or invalid type library" Bug
            'see this bug on local MSDN on this link
            'ms-help://MS.VSCC.v90/MS.MSDNQTR.v90.en/enu_kboffdev/offdev/320369.htm
            excelExport.UserControl = True
            Dim oldCI As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
            System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en-US")
            '----------end bug avoiding---

            excelExport.Visible = False
            excelExport.DisplayAlerts = False
            excelBook = excelExport.Workbooks.Add
            excelSheets = excelBook.Worksheets
            excelSheet = CType(excelSheets.Item(1), Worksheet)
            excelSheet.Name = "OutPut"
            excelCells = excelSheet.Cells

            PopulateSheet(dt, excelCells)

            excelSheet.SaveAs(fileName)

            Return True

        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        Finally
            excelBook.Close()
            excelExport.Quit()


            excelExport = Nothing : excelBook = Nothing : excelSheets = Nothing
            excelSheet = Nothing : excelCells = Nothing

            System.GC.Collect()

        End Try
    End Function

    Private Shared Sub PopulateSheet(ByVal dt As System.Data.DataTable, ByVal oCells As Range)
        Dim dRow As DataRow
        Dim dataArray() As Object
        Dim count As Integer
        Dim column_count As Integer

        'Output Columns Headers
        For column_count = 0 To dt.Columns.Count - 1
            oCells(1, column_count + 1) = dt.Columns(column_count).ToString
            With CType(oCells(1, column_count + 1), Range)
                .Interior.ColorIndex = 6
                .Interior.Pattern = 1

                .Font.Name = "Arial"
                .Font.Size = 12
                .Font.Bold = True
                .Font.ColorIndex = 5

            End With
        Next

        'Output Data
        For count = 0 To dt.Rows.Count - 1
            dRow = dt.Rows.Item(count)
            dataArray = dRow.ItemArray
            For column_count = 0 To dt.Columns.Count - 1
                oCells(count + 2, column_count + 1) = dataArray(column_count).ToString
            Next

            If count Mod 20 = 0 Then
                'just to make a little bit rest for the system
                ' and this condition is to not take long time for this loop
                System.Windows.Forms.Application.DoEvents()
            End If

        Next
    End Sub
#End Region

#Region "Export By Html"

    ''' <summary>
    ''' exports a DataTable to an Excel Sheet using "HTML Table" 
    ''' that Excel 2003 thankfully recognizes 
    ''' </summary>
    ''' <param name="fileName">the destination file</param>
    ''' <param name="dt">the data to export</param>
    Public Shared Function ExportByHtml(ByVal fileName As String, ByVal dt As Data.DataTable) As Boolean
        Try

            Dim grid As New WebControls.DataGrid
            grid.HeaderStyle.Font.Bold = True
            grid.DataSource = dt
            grid.GridLines = WebControls.GridLines.Both
            grid.DataMember = dt.TableName

            grid.DataBind()

            Using sw As New StreamWriter(fileName)
                Using hw As New HtmlTextWriter(sw)
                    grid.RenderControl(hw)
                End Using
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

#Region "Export By XML"
    ''' <summary>
    ''' exports a DataTable to an Excel Sheet using "XML SpreadSheet" formating 
    ''' to create an Excel-Distinguished formatted file
    ''' </summary>
    ''' <param name="fileName">the destination file</param>
    ''' <param name="dt">the data to export</param>
    Public Shared Function ExportByXML(ByVal fileName As String, ByVal dt As Data.DataTable) As Boolean
        Try

            Dim excelDoc As New System.IO.StreamWriter(fileName)

            Dim startExcelXML As New System.Text.StringBuilder(1024)

            startExcelXML.Append("<?xml version=""1.0""?>")
            startExcelXML.Append("<Workbook xmlns=""urn:schemas-microsoft-com:office:spreadsheet""")
            startExcelXML.Append(" xmlns:o=""urn:schemas-microsoft-com:office:office""")
            startExcelXML.Append(" xmlns:x=""urn:schemas-    microsoft-com:office:excel""")
            startExcelXML.Append(" xmlns:ss=""urn:schemas-microsoft-com:office:spreadsheet"">")
            startExcelXML.Append(" <Styles>")
            startExcelXML.Append("	<Style ss:ID=""Default"" ss:Name=""Normal"">  <Alignment ss:Vertical=""Bottom""/> <Font/> <Interior/> <NumberFormat/> <Protection/> </Style>")
            startExcelXML.Append(" 	<Style ss:ID=""BoldColumn"">  <Interior ss:Color=""#FFFF00"" ss:Pattern=""Solid""/>   <Font x:Family=""Swiss"" ss:Bold=""1""/>  </Style>")
            startExcelXML.Append("	<Style ss:ID=""StringLiteral""> <NumberFormat ss:Format=""@""/> </Style>")
            startExcelXML.Append("	<Style ss:ID=""Decimal""> <NumberFormat ss:Format=""0.0000""/> </Style>")
            startExcelXML.Append("	<Style ss:ID=""Integer""> <NumberFormat ss:Format=""0""/> </Style>")
            startExcelXML.Append("	<Style ss:ID=""DateLiteral""> <NumberFormat ss:Format=""mm/dd/yyyy;@""/> </Style>")
            startExcelXML.Append(" </Styles>" & Environment.NewLine)


            Const endExcelXML = "</Workbook>"

            Dim rowCount As Integer = 0
            Dim sheetCount As Integer = 1
            startExcelXML.Append("<Worksheet ss:Name=""Output-Sheet" & sheetCount.ToString & """>")


            excelDoc.WriteLine(startExcelXML.ToString)

            excelDoc.WriteLine("<Table>")
            excelDoc.WriteLine("<Row>")
            For Each Col As DataColumn In dt.Columns
                excelDoc.Write(ControlChars.Tab)
                excelDoc.Write("<Cell ss:StyleID=""BoldColumn""><Data ss:Type=""String"">")
                excelDoc.Write(Col.ColumnName)
                excelDoc.WriteLine("</Data></Cell>")
            Next

            excelDoc.WriteLine("</Row>")
            For Each Row As DataRow In dt.Rows

                rowCount += 1
                'if the number of rows is > 64000 create a new page to continue output
                If rowCount = 64000 Then

                    rowCount = 0
                    sheetCount += 1
                    excelDoc.Write("</Table>")
                    excelDoc.Write(" </Worksheet>")
                    excelDoc.Write("<Worksheet ss:Name=""Sheet" + sheetCount + """>")
                    excelDoc.Write("<Table>")
                End If

                excelDoc.WriteLine("<Row>")

                For ColI As Integer = 0 To dt.Columns.Count - 1
                    excelDoc.Write(ControlChars.Tab)

                    Select Case dt.Columns(ColI).DataType.Name.ToUpper

                        Case GetType(String).Name.ToUpper

                            Dim XMLstring = Row(ColI).ToString()
                            XMLstring = XMLstring.Trim()
                            XMLstring = XMLstring.Replace("&", "&")
                            XMLstring = XMLstring.Replace(">", ">")
                            XMLstring = XMLstring.Replace("<", "<")
                            excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")

                            excelDoc.Write(XMLstring)
                            excelDoc.WriteLine("</Data></Cell>")

                        Case GetType(Date).Name.ToUpper
                            'Excel has a specific Date Format of YYYY-MM-DD followed by  
                            'the letter 'T' then hh:mm:sss.lll Example 2005-01-31T24:01:21.000
                            'The Following Code puts the date stored in XMLDate 
                            'to the format above
                            Dim XMLDate As DateTime = CType(Row(ColI), DateTime)
                            Dim XMLDatetoString = "" 'Excel Converted Date
                            XMLDatetoString = XMLDate.Year.ToString() & "-" & _
                            If(XMLDate.Month < 10, "0" & XMLDate.Month.ToString(), _
                               XMLDate.Month.ToString()) & "-" & _
                            If(XMLDate.Day < 10, "0" & XMLDate.Day.ToString(), _
                                XMLDate.Day.ToString()) & "T" & _
                            If(XMLDate.Hour < 10, "0" & XMLDate.Hour.ToString(), _
                                XMLDate.Hour.ToString()) & ":" & _
                            If(XMLDate.Minute < 10, "0" & XMLDate.Minute.ToString(), _
                                XMLDate.Minute.ToString()) & ":" & _
                            If(XMLDate.Second < 10, "0" & XMLDate.Second.ToString(), _
                                XMLDate.Second.ToString()) & ".000"

                            excelDoc.Write("<Cell ss:StyleID=""DateLiteral""><Data ss:Type=""DateTime"">")
                            excelDoc.Write(XMLDatetoString)
                            excelDoc.WriteLine("</Data></Cell>")

                        Case "System.Boolean".ToUpper
                            excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                            excelDoc.Write(Row(ColI).ToString())
                            excelDoc.WriteLine("</Data></Cell>")

                        Case GetType(Byte).Name.ToUpper, GetType(Int16).Name.ToUpper, _
                             GetType(Int32).Name.ToUpper, GetType(Int64).Name.ToUpper, _
                             GetType(Integer).Name.ToUpper, GetType(Long).Name.ToUpper


                            excelDoc.Write("<Cell ss:StyleID=""Integer""><Data ss:Type=""Number"">")
                            excelDoc.Write(Row(ColI).ToString())
                            excelDoc.WriteLine("</Data></Cell>")

                        Case GetType(Decimal).Name.ToUpper, GetType(Single).Name.ToUpper, _
                             GetType(Double).Name.ToUpper

                            excelDoc.Write("<Cell ss:StyleID=""Decimal""><Data ss:Type=""Number"">")
                            excelDoc.Write(Row(ColI).ToString())
                            excelDoc.WriteLine("</Data></Cell>")

                        Case GetType(DBNull).Name.ToUpper
                            excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                            excelDoc.Write("")
                            excelDoc.WriteLine("</Data></Cell>")


                        Case "Byte[]".ToUpper
                            excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                            excelDoc.Write("System.Byte[]")
                            excelDoc.WriteLine("</Data></Cell>")

                        Case Else

                            excelDoc.Write("<Cell ss:StyleID=""StringLiteral""><Data ss:Type=""String"">")
                            excelDoc.Write("ERROR HAPPENED")
                            excelDoc.WriteLine("</Data></Cell>")

                    End Select
                Next
                excelDoc.WriteLine("</Row>")
            Next
            excelDoc.WriteLine("</Table>")
            excelDoc.Write(" </Worksheet>")
            excelDoc.Write(endExcelXML)
            excelDoc.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region
End Class
