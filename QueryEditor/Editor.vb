#Region "Imports"

Imports System.ComponentModel
Imports System.IO
Imports System.Threading.Tasks
Imports QueryEditor.BindingDataToNode
Imports QueryEditor.ColoringWords
Imports QueryEditor.Queries

#End Region

Public Class Editor
    Private _findReplaceForm As FindReplace
    Private _progressMessage As Working

    ' TODO: Make this private. Requires proper passing between classes.
    Public Shared ReadOnly Property CurrentDb As DataBase = New DataBase("", "") _
    ' this declaration won't open the database

    Private readonly Property TrimmedSql As String
        Get
            trimmedSql = CType(Invoke(Function() As String
                Return If(Rtext_Doc.SelectionLength > 0,
                          Rtext_Doc.SelectedText,
                          Rtext_Doc.Text).Trim()
            End Function), String)
        End Get
    End Property

    ''' <summary>
    '''     I MADE THESE PRIVATE.
    '''     OLD COMMENT:
    '''     one of the variables that holds an Index in the "ImageList" of the "TrView_dbObjects"
    '''     in the main form ....  Input them as Public variables
    '''     so I can change their values without affecting codes that use them
    ''' </summary>
    Private _textIcon As Integer

    Private _numberIcon As Integer

    Private _dateIcon As Integer

    Private _booleanIcon As Integer

    Private _binaryArrayIcon As Integer

    Private _dbIcon As Integer

    Private _tableIcon As Integer

    Private _queryIcon As Integer

    Private _groupIcon As Integer

    Private _functionIcon As Integer

' ReSharper disable once MemberCanBePrivate.Global
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Threading.Thread.CurrentThread.TrySetApartmentState(Threading.ApartmentState.STA)
    End Sub

#Region "Actions"

    Private Sub NewQuery()
        With Rtext_Doc
            If .TextModified Then
                Dim saveAnswer = .Ask4Save()
                Select Case SaveAnswer
                    Case MsgBoxResult.Ok
                        .Save()
                    Case MsgBoxResult.Cancel
                        Exit Sub
                    Case MsgBoxResult.No
                        'nothing to do
                End Select
            End If
        End With

        Using newQueryForm As New NewQuery
            With NewQueryForm
                If .ShowDialog() = DialogResult.OK Then
                    If .ClearOutput Then
                        ClearOutput()
                    End If

                    If .ClearParams Then
                        CurrentDb.Query.QueryParams.Clear()
                    End If
                    Rtext_Doc.Clear()
                End If
            End With
        End Using
    End Sub

    Private Async Function Connect() As Task
        Using udlForm As New ConnectionDialog
            With udlForm
                .PassWord = CurrentDb.PassWord
                .DataBasePath = CurrentDb.Path
            End With

            If udlForm.ShowDialog() = DialogResult.OK Then
                If _progressMessage is Nothing Then _progressMessage = New Working

                _progressMessage.Message = "Fetching DB Schema..."
                _progressMessage.Show
                Try
                    CurrentDb.PassWord = udlForm.PassWord
                    CurrentDb.Path = udlForm.DataBasePath
                    status_Connection.Text = CurrentDb.Path

                    ' We must get data before filling the objects
                    Await Task.Run(Function () CurrentDb.RefreshSchema())

                    FillDbObjects
                Catch ex As Exception
                    MsgBox(ex.Message)
                Finally
                    _progressMessage.Hide()
                End Try
            End If
        End Using
    End Function

    Private Async Function Execute(theQuery As String) As Task
        If CurrentDb.Path = "" Then
            ShowResults2MsgTab("Not connected to the DataBase", True)
            Return
        End If

        CurrentDb.Query = New SqlQuery(theQuery) ' Me.Rtext_Doc.Query.Sql()

        If CurrentDb.Query.Sql = "" Then
            ShowResults2MsgTab("No Sql statement is supplied to execute ........")
            Return
        End If

        Invoke(Sub() status_CurTask.Text = "Executing.....")

        Try
            If CBool(CurrentDb.Query.QueryType And SqlQueryType.ParameterizedSqlQuery) Then
                'in case we want to do some operation when the Current query is a Parameterized one
                If CurrentDb.Query.QueryParams Is Nothing OrElse
                   CurrentDb.Query.QueryParams.Count = 0 Then

                    Throw New Exception("Parameterized query , but Parameters are not set yet")
                End If
            End If

            If CBool(CurrentDb.Query.QueryType And SqlQueryType.SelectSqlQuery) Then
                Dim dt As DataTable
                dt = Await CurrentDb.GetDataTable(CurrentDb.Query)

                ShowResults2GridTab(dt)
                ShowResults2MsgTab("The command completed successfully.", showIt := False)
            ElseIf CBool(CurrentDb.Query.QueryType And SqlQueryType.NoneQuerySqlQuery) Then
                Dim affected = Await CurrentDb.ExecuteSQLCommand(CurrentDb.Query)
                ShowResults2MsgTab(Affected & " Row(s) affected")
            Else
                ShowResults2MsgTab("Error Encountered")
            End If

        Catch ex As Exception
            ShowResults2MsgTab(ex.Message)
        Finally
            Invoke(Sub() status_CurTask.Text = "Ready")
        End Try
    End function

#Region "Out Put Handling"

    Private Sub ShowHideOutput(Optional ByVal outputVisible As Boolean = True)

        Split_Doc_And_Output.Panel2Collapsed = Not OutputVisible
        mnu_ShowOutPutPanel.Checked = OutputVisible
    End Sub

    Public Sub ShowResults2MsgTab(msg As String,
                                  Optional ByVal showIt As Boolean = True,
                                  Optional ByVal append As Boolean = False
                                  )
        Invoke(Sub()
            txt_OutPut.Text = If(Append,
                                 txt_OutPut.Text & Environment.NewLine & MSG,
                                 MSG)
            txt_OutPut.Text += vbTab & Now.ToLongTimeString

            If ShowIt Then
                Tabs_Output.SelectedTab = TabPage_OutPutTxt
            End If

            txt_OutPut.SelectionLength = 0 'remove selection

            If Split_Doc_And_Output.Panel2Collapsed Then
                ShowHideOutput(True)
            End If
        End Sub)
    End Sub

    Private Sub ShowResults2GridTab(results As DataTable)
        Invoke(Sub()
            dgv_OutPut.DataSource = Results

            Tabs_Output.SelectedTab = TabPage_OutPutDGV

            If Split_Doc_And_Output.Panel2Collapsed Then
                ShowHideOutput(True)
            End If
        End Sub)
    End Sub

    Private Sub ClearOutput()
        dgv_OutPut.Columns.Clear()
        dgv_OutPut.DataSource = Nothing
        txt_OutPut.Clear()
    End Sub

#End Region

    Private Sub FillDbObjects() 'As Task
        status_CurTask.Text = "Fetching DB Schema..."

        Dim dbName = Path.GetFileNameWithoutExtension(CurrentDb.Path)
        Dim trvNodeTables As TreeNode
        Dim trvNodeQueries As TreeNode
        Dim folderBinding As New NodeBindingData(NodeTypes.FolderNode)
        Dim tableBinding As New NodeBindingData(NodeTypes.TableNode)
        Dim queryBinding As New NodeBindingData(NodeTypes.QueryNode)
        Dim fieldBinding As New NodeBindingData(NodeTypes.FieldNode)
        ' This wasn't in use?
        'Dim FunctionBinding As New NodeBindingData(NodeTypes.FunctionNode)

        'to Update the Text of the DataBase Node with the DataBaseName
        With TrView_dbObjects.Nodes(0)
            .Nodes.Clear()
            .Text = DbName

            trvNodeTables = .Nodes.Add(DbName & "Tables", "Tables (" & CurrentDb.TableCount & ")", _groupIcon,
                                       _groupIcon)
            trvNodeTables.Tag = FolderBinding
            trvNodeTables.ToolTipText = trvNodeTables.Text

            trvNodeQueries = .Nodes.Add(DbName & "Queries", "Queries (" & CurrentDb.QueryCount & ")", _groupIcon,
                                        _groupIcon)
            trvNodeQueries.Tag = FolderBinding
            trvNodeQueries.ToolTipText = trvNodeQueries.Text
        End With

        Try 'to add Tables           
            For Each dataTable In CurrentDb.Tables.OrderBy(Function(T) T.TableName)
                Dim tblNode As TreeNode = trvNodeTables.Nodes.Add(dataTable.TableName, dataTable.TableName, _tableIcon,
                                                                  _tableIcon)
                tblNode.ToolTipText = dataTable.TableName
                tblNode.ContextMenuStrip = CntxtMnu_TrView_dbObjects_Tblqry
                tblNode.Tag = tableBinding

                Dim asterisk = tblNode.Nodes.Add("All_" & dataTable.TableName & "_Cols", "*")
                asterisk.ToolTipText = "*"
                asterisk.Tag = FieldBinding

                dim dataColumns = dataTable.Columns
                For Each col As DataColumn In (From c As DataColumn In dataColumns Select c Order By c.ColumnName)
                    tblNode.Nodes.Add(GetNodeByDataType(Col, FieldBinding))
                Next
            Next

            'to add Queries
            For Each qry In CurrentDb.Queries
                Dim qryNode As TreeNode = trvNodeQueries.Nodes.Add(qry.TableName, qry.TableName, _queryIcon, _queryIcon)
                qryNode.ToolTipText = qry.TableName
                qryNode.ContextMenuStrip = CntxtMnu_TrView_dbObjects_Tblqry
                qryNode.Tag = QueryBinding

                Dim asterisk = qryNode.Nodes.Add("All_" & qry.TableName & "_Cols", "*")
                asterisk.ToolTipText = "*"
                asterisk.Tag = FieldBinding

                Dim queryColumns = qry.Columns
                For Each col As DataColumn In (From c As DataColumn In queryColumns Select c Order By c.ColumnName)
                    qryNode.Nodes.Add(GetNodeByDataType(Col, FieldBinding))
                Next
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Invoke(Sub() status_CurTask.Text = "Ready")
        End Try
    End Sub

    Private Function GetNodeByDataType(col As DataColumn,
                                       bindingObj As NodeBindingData) As TreeNode

        Dim node As New TreeNode(Col.ColumnName)
        Dim imgIndex as Integer

        Select Case Col.DataType.Name.ToUpper
            Case GetType(String).Name.ToUpper, GetType(Guid).Name.ToUpper
                ImgIndex = _textIcon
            Case GetType(Boolean).Name.ToUpper
                ImgIndex = _booleanIcon
            Case GetType(Byte).Name.ToUpper, GetType(Int16).Name.ToUpper,
                GetType(Int32).Name.ToUpper, GetType(Int64).Name.ToUpper,
                GetType(Integer).Name.ToUpper, GetType(Long).Name.ToUpper,
                GetType(Decimal).Name.ToUpper, GetType(Single).Name.ToUpper,
                GetType(Double).Name.ToUpper

                ImgIndex = _numberIcon
            Case GetType(Date).Name.ToUpper
                ImgIndex = _dateIcon
            Case "Byte[]".ToUpper
                ImgIndex = _binaryArrayIcon
            Case Else
                ImgIndex = 0
        End Select

        node.ToolTipText = Col.ColumnName
        node.ImageIndex = ImgIndex
        node.SelectedImageIndex = ImgIndex
        node.Tag = BindingObj

        Return node
    End Function

    Private Shared Sub BringColorsFromSettings()
        ColoredWords.SqlLiteralOperatorsColor = My.Settings.LiteralOperatorsColor
        ColoredWords.FunctionsColor = My.Settings.FunctionsColor
        ColoredWords.SqlOperatorsColor = My.Settings.OperatorsColor
        ColoredWords.KeyWordsColor = My.Settings.KeyWordsColor
        ColoredWords.StringsColor = My.Settings.StringsColor
        ColoredWords.CommentsColor = My.Settings.CommentsColor
        ColoredWords.SqlPunctuationColor = My.Settings.PunctuationColor
    End Sub

#End Region

#Region "Rtext RichText"
    ' TODO: These were not implemented. 
    '
    'Private Shared Sub Rtext_Doc_DragDrop(sender As Object, e As DragEventArgs) Handles Rtext_Doc.DragDrop
    '    Dim dragedTxt As String = e.Data.GetData(GetType(String))
    '    If DragedTxt IsNot Nothing Then
    '        'Me.Rtext_Doc.SelectedText = DragedTxt
    '        'Me.Rtext_Doc.RecolorWord("Rtext_Doc_DragDrop", True)
    '    End If
    'End Sub
    '
    'Private Shared Sub Rtext_Doc_DragEnter(sender As Object, e As DragEventArgs) Handles Rtext_Doc.DragEnter
    '    'If e.Data.GetDataPresent(DataFormats.Text) Then
    '    '    e.Effect = DragDropEffects.Copy
    '    'Else
    '    '    e.Effect = DragDropEffects.None
    '    'End If
    'End Sub

    Private Sub Rtext_Doc_KeyDown(sender As Object, e As KeyEventArgs) _
        Handles Rtext_Doc.KeyDown

        If e.Shift = True AndAlso e.KeyCode = Keys.Insert Then
            e.SuppressKeyPress = True
            Rtext_Doc.Paste()
        End If
    End Sub


    Private Sub Rtext_Doc_OverType_Changed(sender As Object, overTypeState As Boolean) _
        Handles Rtext_Doc.OverTypeChanged
        Status_Ins_Ovr.Text = If(OverTypeState, "OVR", "INS")
    End Sub

    Private Sub Rtext_Doc_SelectionChanged(sender As Object, e As EventArgs) Handles Rtext_Doc.SelectionChanged
        status_CurCol.Text = "Col " & Rtext_Doc.CursorColumnNumber
        status_CurLine.Text = "Ln " & Rtext_Doc.CurrLineNumber + 1
    End Sub

#End Region

#Region "TrView_dbObjects"

    Private Shared Sub TrView_dbObjects_DragEnter(sender As Object, e As DragEventArgs) _
        Handles TrView_dbObjects.DragEnter
        e.Effect = e.AllowedEffect 'to change the Cursor
    End Sub

    Private Sub TrView_dbObjects_ItemDrag(sender As Object, e As ItemDragEventArgs) Handles TrView_dbObjects.ItemDrag
        Dim node = CType(e.Item, TreeNode)

        If Node.Tag Is Nothing Then Exit Sub

        Dim bindingData = CType(Node.Tag, NodeBindingData)
        Select Case bindingData.NodeType
            Case NodeTypes.TableNode, NodeTypes.QueryNode
                TrView_dbObjects.DoDragDrop(" [" & Node.Text.Trim & "] ", DragDropEffects.All)

            Case NodeTypes.FieldNode
                If Node.Text.Trim = "*" Then ' we can't put [*] , it has to be just "*"
                    TrView_dbObjects.DoDragDrop(" * ", DragDropEffects.All)
                Else
                    TrView_dbObjects.DoDragDrop(" [" & Node.Parent.Text & "].[" & Node.Text.Trim & "] ",
                                                DragDropEffects.All)
                End If
            Case NodeTypes.FunctionNode
                TrView_dbObjects.DoDragDrop(" " & Node.Text.Trim & " ( ) ", DragDropEffects.All)
            Case Else
                Exit Sub
        End Select
    End Sub

#End Region

#Region "dgv_OutPut"

    Private Sub Dgv_OutPut_RowsCountChanged(sender As Object, e As EventArgs) Handles dgv_OutPut.RowsCountChanged
        'this way, we don't need to update the status bar
        'when a gird is filled or cleared, it will automatically
        'update it with the new Rows count
        status_OutPutCount.Text = dgv_OutPut.Rows.Count & " Row(s)"
    End Sub

    Private Async Sub Btn_Exec_Click(sender As Object, e As EventArgs) Handles btn_Exec.Click
        Await Execute(trimmedSql)
    End Sub

#End Region

#Region "Form Events Handlers"

    Private Sub Frm_Editor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        My.Settings.Reload()
        BringColorsFromSettings()
        Rtext_Doc.Font = My.Settings.EditorFont

        Rtext_Doc.AllowDrop = True ' sorry , but I can't make it at design time

        '-------------
        ToolBar_Connect.Image = New Bitmap(GetType(BindingSource), "BindingSource.bmp")
        mnu_Connect.Image = ToolBar_Connect.Image
        Cmnu_DB_Connect.Image = ToolBar_Connect.Image
        '------------

        'populate the TreeView ImageList with proper Images and set the Indices
        'to use them later
        Dim imgLst As New ImageList
        With ImgLst.Images
            .Add(New Bitmap(1, 1))
            .Add(My.Resources.Folder)
            _groupIcon = .Count - 1
            .Add(My.Resources.FunctionImg)
            _functionIcon = .Count - 1
            .Add(My.Resources.DB.ToBitmap)
            _dbIcon = .Count - 1
            .Add(New Bitmap(GetType(DataGrid), "DataGrid.bmp"))
            _tableIcon = .Count - 1
            .Add(New Bitmap(GetType(DataGrid), "DataGrid.bmp"))
            _queryIcon = .Count - 1
            .Add(New Bitmap(GetType(TextBox), "TextBox.bmp"))
            _textIcon = .Count - 1
            .Add(New Bitmap(GetType(NumericUpDown), "NumericUpDown.bmp"))
            _numberIcon = .Count - 1
            .Add(New Bitmap(GetType(DateTimePicker), "DateTimePicker.bmp"))
            _dateIcon = .Count - 1
            .Add(New Bitmap(GetType(CheckBox), "CheckBox.bmp"))
            _booleanIcon = .Count - 1
            .Add(My.Resources.BinaryObject)
            _binaryArrayIcon = .Count - 1
        End With
        ImgLst.TransparentColor = Color.Fuchsia
        TrView_dbObjects.ImageList = ImgLst

        'DB PlaceHolder
        TrView_dbObjects.Nodes(0).ImageIndex = _dbIcon
        TrView_dbObjects.Nodes(0).SelectedImageIndex = _dbIcon
        TrView_dbObjects.Nodes(0).Tag = New NodeBindingData(NodeTypes.DataBaseNode)
        '-------------

        With TrView_dbObjects.Nodes(1)
            .ImageIndex = _functionIcon
            .SelectedImageIndex = _functionIcon
            .Tag = New NodeBindingData(NodeTypes.FolderNode)

            For Each groupNode As TreeNode In .Nodes
                groupNode.ImageIndex = _groupIcon
                groupNode.SelectedImageIndex = _groupIcon
                groupNode.Tag = New NodeBindingData(NodeTypes.FolderNode)

                For Each itemNode As TreeNode In groupNode.Nodes
                    ItemNode.ImageIndex = _functionIcon
                    ItemNode.SelectedImageIndex = _functionIcon
                    ItemNode.Tag = New NodeBindingData(NodeTypes.FunctionNode)
                Next
            Next
        End With

        Show()
        Rtext_Doc.Focus()
    End Sub

    Private Sub Frm_Editor_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If ((e.CloseReason And CloseReason.ApplicationExitCall) = CloseReason.ApplicationExitCall) Or
           ((e.CloseReason And CloseReason.UserClosing) = CloseReason.UserClosing) Then

            If Rtext_Doc.TextModified Then
                Dim questAnswer As MsgBoxResult
                QuestAnswer = Rtext_Doc.Ask4Save()

                Select Case QuestAnswer
                    Case MsgBoxResult.Yes
                        Rtext_Doc.Save()
                    Case MsgBoxResult.Cancel
                        e.Cancel = True
                    Case MsgBoxResult.No
                        ' nothing to do ....... just close nice and easy
                End Select
            End If

            If Not e.Cancel Then
                'next is to save My.Settings before the app closes
                My.Settings.LiteralOperatorsColor = ColoredWords.SqlLiteralOperatorsColor
                My.Settings.FunctionsColor = ColoredWords.FunctionsColor
                My.Settings.OperatorsColor = ColoredWords.SqlOperatorsColor
                My.Settings.KeyWordsColor = ColoredWords.KeyWordsColor
                My.Settings.StringsColor = ColoredWords.StringsColor
                My.Settings.CommentsColor = ColoredWords.CommentsColor
                My.Settings.PunctuationColor = ColoredWords.SqlPunctuationColor

                My.Settings.EditorFont = Rtext_Doc.Font
                My.Settings.Save()
            End If
        End If
    End Sub

#End Region

#Region "File"

    Private Sub Mnu_File_NewQuery_Click(sender As Object, e As EventArgs) Handles mnu_File_NewQuery.Click
        NewQuery()
    End Sub

    Private Async Sub Mnu_Connect_Click(sender As Object, e As EventArgs) _
        Handles mnu_Connect.Click, ToolBar_Connect.Click, Cmnu_DB_Connect.Click

        'Dim backTask As New Thread(New ThreadStart(AddressOf Connect))
        'backTask.SetApartmentState(ApartmentState.STA)
        'backTask.Start()
        Await Connect()
    End Sub

    Private Sub Mnu_Open_Click(sender As Object, e As EventArgs) _
        Handles mnu_Open.Click, ToolBar_Open.Click
        status_CurTask.Text = "Fetching the new file.........."
        Rtext_Doc.OpenNewFile()
        status_CurTask.Text = "Ready"
    End Sub

    Private Sub Mnu_SaveAs_Click(sender As Object, e As EventArgs) Handles mnu_SaveAs.Click
        Rtext_Doc.SaveAs()
    End Sub

    Private Sub Mnu_Save_Click(sender As Object, e As EventArgs) _
        Handles mnu_Save.Click, ToolBar_Save.Click

        Rtext_Doc.Save()
    End Sub

    Private Sub Mnu_Exit_Click(sender As Object, e As EventArgs) Handles mnu_Exit.Click
        Close
        'Application.Exit()  ' Ew, gross
    End Sub

#End Region

#Region "View"

    Private Sub Mnu_ObjectsBrowser_Click(sender As Object, e As EventArgs) Handles mnu_ObjectsBrowser.Click
        Split_ObjBrowser_And_Others.Panel1Collapsed = Not Split_ObjBrowser_And_Others.Panel1Collapsed
        mnu_ObjectsBrowser.Checked = Not mnu_ObjectsBrowser.Checked
    End Sub

    Private Sub Mnu_ShowOutPutPanel_Click(sender As Object, e As EventArgs) Handles mnu_ShowOutPutPanel.Click
        mnu_ShowOutPutPanel.Checked = Not mnu_ShowOutPutPanel.Checked
        ShowHideOutput(mnu_ShowOutPutPanel.Checked)
    End Sub

#End Region

#Region "Format"

    Private Sub Mnu_WordWrap_CheckedChanged(sender As Object, e As EventArgs) Handles mnu_WordWrap.CheckedChanged
        Rtext_Doc.WordWrap = mnu_WordWrap.Checked
    End Sub

    Private Sub Mnu_Font_Click(sender As Object, e As EventArgs) Handles mnu_Font.Click
        Using fontDlg As New FontDialog
            With FontDlg
                .ShowColor = False
                .ShowEffects = False
                .ShowApply = False 'don't wanna add EventHandler 4 it , so he has to click "OK"
                .Font = Rtext_Doc.Font
                If .ShowDialog() = DialogResult.OK Then
                    Rtext_Doc.Font = .Font
                End If
            End With
        End Using
    End Sub

    Private Sub EditorOptions_Click(sender As Object, e As EventArgs) _
        Handles Cmnu_EditorOptions.Click, mnu_EditorOptions.Click

        Using editorOptions As New EditorOptions
            With EditorOptions
                If .ShowDialog() = DialogResult.OK Then
                    ColoredWords.Initiate()
                    Rtext_Doc.RecolorWords()
                End If
            End With
        End Using
    End Sub

#End Region

#Region "Edit"

    Private Sub Mnu_Undo_Click(sender As Object, e As EventArgs) _
        Handles mnu_Undo.Click, Cmnu_Undo.Click

        If Rtext_Doc.CanUndo Then
            Rtext_Doc.Undo()
        End If
    End Sub

    Private Sub Mnu_Redo_Click(sender As Object, e As EventArgs) _
        Handles mnu_Redo.Click, Cmnu_Redo.Click
        If Rtext_Doc.CanRedo Then
            Rtext_Doc.Redo()
        End If
    End Sub

    Private Sub Edit_DropDownOpening() _
        Handles mnu_Edit.DropDownOpening, ContextMnu_RText.Opening

        mnu_Undo.Enabled = Rtext_Doc.CanUndo
        mnu_Redo.Enabled = Rtext_Doc.CanRedo

        Cmnu_Undo.Enabled = Rtext_Doc.CanUndo
        Cmnu_Redo.Enabled = Rtext_Doc.CanRedo
        '-------------------


        'I could use it every time like it is ... 
        'but this is for more optimization
        'it will calculate it only once
        Dim hasSelection = CBool(Rtext_Doc.SelectedText.Length > 0)
        mnu_Cut.Enabled = HasSelection
        mnu_Copy.Enabled = HasSelection

        Cmnu_Cut.Enabled = HasSelection
        Cmnu_Copy.Enabled = HasSelection

        '---------------------------------
        'next is to Enable/Disable the [Select All] menu item
        mnu_SelectAll.Enabled = CBool(Rtext_Doc.Text <> "")
        Cmnu_SelectAll.Enabled = mnu_SelectAll.Enabled
    End Sub

    Private Sub Mnu_Cut_Click(sender As Object, e As EventArgs) _
        Handles mnu_Cut.Click, Cmnu_Cut.Click, ToolBar_Cut.Click
        Try
            Rtext_Doc.Cut()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Mnu_Copy_Click(sender As Object, e As EventArgs) _
        Handles mnu_Copy.Click, Cmnu_Copy.Click, ToolBar_Copy.Click
        Try
            Rtext_Doc.Copy()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Paste_Clicks(sender As Object, e As EventArgs) _
        Handles mnu_Paste.Click, Cmnu_Paste.Click, ToolBar_Paste.Click

        Rtext_Doc.Paste()
        'Try

        '    If Clipboard.ContainsText Then
        '        'we're not gonna use Me.Rtext_Doc.Paste()
        '        'cuz it's gonna past it as it is 
        '        '(it could has some Formatting , user is pasting from 'MS Word' or something)
        '        'so we'll deal with it our way
        '        Me.Rtext_Doc.SelectedText = Clipboard.GetText(TextDataFormat.Text)

        '    End If
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub SelectAll_Click() _
        Handles Cmnu_SelectAll.Click, mnu_SelectAll.Click

        Rtext_Doc.SelectAll()
    End Sub

    Private Sub Mnu_FindReplace_Click(sender As Object, e As EventArgs) _
        Handles mnu_Replace.Click, mnu_Find.Click

        If Not FindReplace.Opened Then
            _findReplaceForm = New FindReplace(Rtext_Doc) With {
                .Owner = Me
                }
            _findReplaceForm.SelectTab(sender Is mnu_Replace)

            _findReplaceForm.Show()
        Else
            _findReplaceForm.Activate()
            _findReplaceForm.SelectTab(sender Is mnu_Replace)
        End If
    End Sub

#End Region

#Region "Output"

    Private Sub Mnu_ClearOutPut_Click(sender As Object, e As EventArgs) Handles mnu_ClearOutPut.Click
        If MsgBox("Are you sure you want to clear all outputs?",
                  MsgBoxStyle.Question Or MsgBoxStyle.YesNo,
                  "Caution") = MsgBoxResult.Yes Then

            ClearOutput()
        End If
    End Sub

    Private Async Sub Mnu_Export2Excel(sender As Object, e As EventArgs) _
        Handles mnu_Export2ExcelByHTML.Click, mnu_Export2ExcelByXML.Click,
                mnu_Export2ExcelByExcelApp.Click

        Using saveFile As New SaveFileDialog
            With saveFile
                .Filter = "Excel 2003 Books (*.xls)|*.xls"
                .Title = "Select the location and the name of the File to export to"

                If .ShowDialog = DialogResult.OK Then
                    ' The following declaration must match the Handles type above
                    Dim sourceMenuItem = CType(sender, ToolStripMenuItem)

                    Cursor = Cursors.WaitCursor
                    dgv_OutPut.Cursor = Cursors.WaitCursor
                    status_CurTask.Text = "Exporting To Excel. This operation could take a while."
                    If _progressMessage is Nothing Then _progressMessage = New Working

                    _progressMessage.Message = status_CurTask.Text
                    _progressMessage.Show
                    Try 'Export method
                        Dim dt = CType(dgv_OutPut.DataSource, DataTable)
                        Select Case sourceMenuItem.Name
                            Case mnu_Export2ExcelByExcelApp.Name
                                Await Task.Run(Function() ExportData.ExportByExcel(.FileName, dt))

                                status_CurTask.Text = "Done"

                                ' Forced cleanup of Excl COM object
                                GC.Collect()
                                GC.WaitForPendingFinalizers()
                                GC.Collect()
                                GC.WaitForPendingFinalizers()
                            Case mnu_Export2ExcelByHTML.Name
                                If ExportData.ExportByHtml(.FileName, dt) Then
                                    MsgBox("Creating HTML File Done", MsgBoxStyle.Information, "Success")
                                End If
                            Case mnu_Export2ExcelByXML.Name
                                If ExportData.ExportByXML(.FileName, dt) Then
                                    MsgBox("Creating XML File Done", MsgBoxStyle.Information, "Success")
                                End If

                            Case Else
                                MsgBox("Error: Can't determine export method?", MsgBoxStyle.Exclamation, "Failure")
                        End Select
                    Catch ex As Exception
                        MsgBox("Error writing to Excel file." &
                               Environment.NewLine & Environment.NewLine &
                               "Technical Info:" & Environment.NewLine & ex.Message,
                               MsgBoxStyle.Exclamation, "Caution")
                        Exit Sub
                    Finally
                        status_CurTask.Text = "Ready"
                        Cursor = Cursors.Default
                        dgv_OutPut.Cursor = Cursors.Default
                        _progressMessage.Hide()
                    End Try
                End If
            End With
        End Using
    End Sub

    Private Sub Mnu_Output_DropDownOpening(sender As Object, e As EventArgs) Handles mnu_Output.DropDownOpening
        mnu_ClearOutPut.Enabled = CBool(dgv_OutPut.Rows.Count > 0)
        mnu_Export2ExcelByHTML.Enabled = CBool(dgv_OutPut.Rows.Count > 0)
        mnu_Export2ExcelByXML.Enabled = CBool(dgv_OutPut.Rows.Count > 0)
        mnu_Export2ExcelByExcelApp.Enabled = CBool(dgv_OutPut.Rows.Count > 0)
    End Sub

#End Region

#Region "Query"

    Private Async Sub Mnu_Exec_Click(sender As Object, e As EventArgs) Handles mnu_Exec.Click
        Await Execute(TrimmedSql)
    End Sub

    Private Shared Sub Mnu_Query_Parameters_Click(sender As Object, e As EventArgs) Handles mnu_Query_Parameters.Click

        Dim paramsForm As New FrmAddParams
        ParamsForm.ShowDialog()
    End Sub

    Private Sub Mnu_Convert2Code_Click(sender As Object, e As EventArgs) Handles mnu_Convert2Code.Click
        Using converterForm As New Text2Code
            With ConverterForm
                .Text2Convert = Rtext_Doc.Query.Sql
                .ShowDialog()
            End With
        End Using
    End Sub

#End Region

#Region "Help"

    Private Shared Sub Mnu_TutorialFeatures_Click(sender As Object, e As EventArgs) Handles mnu_TutorialFeatures.Click
        Using sveDlg As New SaveFileDialog
            With SveDlg
                .Filter = "Archive files (*.RAR)|*.RAR"
                .Title = "Select a folder to save the tutorial file to"
                If SveDlg.ShowDialog = DialogResult.OK Then
                    My.Computer.FileSystem.WriteAllBytes(SveDlg.FileName, My.Resources.Tutorial, False)
                End If
            End With
        End Using
    End Sub

    Private Shared Sub Mnu_About_Click(sender As Object, e As EventArgs) Handles mnu_About.Click
        dim thisBox as New AboutBox
        thisBox.ShowDialog()
    End Sub

#End Region

#Region "Context Menu Events Handlers"

    Private Async Sub Cmnu_DB_Refresh_Click(sender As Object, e As EventArgs) Handles Cmnu_DB_Refresh.Click
        If _progressMessage is Nothing Then _progressMessage = New Working
        _progressMessage.Message = "Refreshing DB Schema..."
        _progressMessage.Show()
        Try
            Await CurrentDb.RefreshSchema()

            FillDbObjects
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
        Finally
            _progressMessage.Hide()
        End Try
    End Sub

    Private Sub CntxtMnu_TrView_dbObjects_DB_Opening(sender As Object, e As CancelEventArgs) _
        Handles CntxtMnu_TrView_dbObjects_DB.Opening
        Cmnu_DB_Refresh.Enabled = CBool(CurrentDb.Path <> "")
    End Sub

    Private Sub CntxtMnu_TrView_dbObjects_TblQry_GenSelectALLScript_Click(sender As Object, e As EventArgs) _
        Handles CntxtMnu_TrView_dbObjects_TblQry_GenSelectALLScript.Click
        Try
            If TrView_dbObjects.LastUsedNode IsNot Nothing Then

                Dim bind = CType(TrView_dbObjects.LastUsedNode.Tag, NodeBindingData)
                If Bind.NodeType = NodeTypes.TableNode Or
                   Bind.NodeType = NodeTypes.QueryNode Then

                    Rtext_Doc.SelectedText = "Select * From [" & TrView_dbObjects.LastUsedNode.Text & "] ;"
                End If

            End If
        Catch ex As Exception
            MsgBox("Error encountered", MsgBoxStyle.Exclamation, "Caution")
        End Try
    End Sub

    Private Sub CntxtMnu_TrView_dbObjects_Tblqry_GenDrop_Click(sender As Object, e As EventArgs) _
        Handles CntxtMnu_TrView_dbObjects_Tblqry_GenDrop.Click
        Try
            If TrView_dbObjects.LastUsedNode IsNot Nothing Then

                Dim bind = CType(TrView_dbObjects.LastUsedNode.Tag, NodeBindingData)
                If Bind.NodeType = NodeTypes.TableNode Or
                   Bind.NodeType = NodeTypes.QueryNode Then

                    Rtext_Doc.SelectedText = "Drop Table [" & TrView_dbObjects.LastUsedNode.Text & "] ;"
                End If

            End If
        Catch ex As Exception
            MsgBox("Error encountered", MsgBoxStyle.Exclamation, "Caution")
        End Try
    End Sub

#End Region
End Class