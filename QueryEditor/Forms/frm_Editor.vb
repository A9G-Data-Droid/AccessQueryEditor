
#Region "Imports"
Imports System.Text.RegularExpressions
Imports QueryEditor.ColoringWords
Imports QueryEditor.BindingDataToNode
Imports QueryEditor.Queries
Imports System.Data.OleDb
Imports System.IO
#End Region

Public Class frm_Editor

    Private FindReplaceForm As frm_FindReplace

#Region "Added Functions/Methods"

    Sub NewQuery(ByVal sender As Object)
        With Rtext_Doc

            If .TextModified Then
                Dim SaveAnswer = .Ask4Save()
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


        Using NewQueryForm As New frm_NewQuery
            With NewQueryForm
                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                    If .ClearOutput Then
                        ClearOutput("mnu_File_NewQuery_Click")
                    End If

                    If .ClearParams Then
                        Query.QueryParams.Clear()
                    End If
                    Me.Rtext_Doc.Clear()
                End If
            End With

        End Using

    End Sub

    Sub Connect(ByVal sender As Object)

        Using Udl_Form As New frm_UDL
            With Udl_Form
                .PassWord = DB.PassWord
                .DataBasePath = DB.Path

                If .ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim Task As String = "Fetching DB Schema.........."
                    frm_Working.Task = Task
                    frm_Working.Show(Task)
                    Try
                        DB.PassWord = .PassWord
                        DB.Path = .DataBasePath

                        status_Connection.Text = DB.Path
                        FillDbObjects()
                    Catch ex As Exception
                        MsgBox(ex.Message)

                    Finally
                        frm_Working.Close()
                    End Try
                End If
            End With
        End Using

    End Sub

    Sub Exec(ByVal sender As Object)

        If DB.Path = "" Then
            ShowResults2MsgTab("Not connected to the DataBase", True)
            Exit Sub
        End If
        Query.Sql = Me.Rtext_Doc.Query.Sql()

        If Query.Sql = "" Then
            ShowResults2MsgTab("No Sql statement is supplied to execute ........")
            Exit Sub
        End If

        Me.status_CurTask.Text = "Executing....."
        Application.DoEvents()

        Try
            Dim Errors As String = ""

            If CBool(Query.QueryType And SqlQueryType.ParameterizedSqlQuery) Then
                'in case we wana do some operation when the Current query is a Parameterized one

                If Query.QueryParams Is Nothing OrElse _
                   Query.QueryParams.Count = 0 Then

                    Throw New Exception("Parameterized query , but Parameters are not set yet")
                End If
            End If

            If CBool(Query.QueryType And SqlQueryType.SelectSqlQuery) Then

                Dim dt As New DataTable

                dt = DB.GetDataTable(Query, Errors)

                If Errors = "" Then
                    ShowResults2GridTab(dt)
                    ShowResults2MsgTab("The command completed successfully.", ShowIt:=False)
                Else
                    ShowResults2MsgTab(Errors)
                End If

            ElseIf CBool(Query.QueryType And SqlQueryType.NoneQuerySqlQuery) Then

                Dim Affected = DB.ExecuteSQLCommand(Query, Errors)

                ShowResults2MsgTab(Errors & Environment.NewLine & _
                                   Affected & " Row(s) affected")

            Else
                ShowResults2MsgTab("Error Encountered")
            End If



        Catch ex As Exception
            ShowResults2MsgTab(ex.Message)
        Finally
            Me.status_CurTask.Text = "Ready"
        End Try

    End Sub

#Region "Out Put Handling"
    Sub ShowHideOutput(Optional ByVal OutputVisible As Boolean = True)

        Split_Doc_And_Output.Panel2Collapsed = Not OutputVisible
        mnu_ShowOutPutPanel.Checked = OutputVisible
    End Sub

    Sub ShowResults2MsgTab(ByVal MSG As String, _
                           Optional ByVal ShowIt As Boolean = True, _
                           Optional ByVal Append As Boolean = False _
                           )

        Me.txt_OutPut.Text = If(Append, _
                                Me.txt_OutPut.Text & Environment.NewLine & MSG, _
                                MSG)
        Me.txt_OutPut.Text += vbTab & Now.ToLongTimeString

        If ShowIt Then
            Me.Tabs_Output.SelectedTab = Me.TabPage_OutPutTxt
        End If

        Me.txt_OutPut.SelectionLength = 0 'remove selection

        If Split_Doc_And_Output.Panel2Collapsed Then
            ShowHideOutput(True)
        End If
    End Sub

    Sub ShowResults2GridTab(ByVal Results As DataTable)
        Me.dgv_OutPut.DataSource = Results

        Me.Tabs_Output.SelectedTab = TabPage_OutPutDGV

        If Split_Doc_And_Output.Panel2Collapsed Then
            ShowHideOutput(True)
        End If
    End Sub

    Sub ClearOutput(ByVal sender As Object)
        Me.dgv_OutPut.Columns.Clear()
        Me.dgv_OutPut.DataSource = Nothing
        Me.txt_OutPut.Clear()
    End Sub
#End Region

    Sub FillDbObjects()
        Me.status_CurTask.Text = "Fetching DB Schema....."
        Application.DoEvents()

        Dim DbName = Path.GetFileNameWithoutExtension(DB.Path)

        Dim trvNode_Tables, trvNode_Queries As TreeNode
        Try
            Dim FolderBinding As New NodeBindingData(NodeTypes.FolderNode)
            Dim TalbeBinding As New NodeBindingData(NodeTypes.TableNode)
            Dim QueryBinding As New NodeBindingData(NodeTypes.QueryNode)
            Dim FieldBinding As New NodeBindingData(NodeTypes.FieldNode)
            Dim FunctionBinding As New NodeBindingData(NodeTypes.FunctionNode)

            'next is to Update the Text of the DataBase Node with the DataNaseName
            'and add the "Tables" and "Queries" Nodes ... as headers
            With TrView_dbObjects.Nodes(0)
                .Nodes.Clear()
                .Text = DbName
                trvNode_Tables = .Nodes.Add(DbName & "Tables", "Tables (" & DB.Tables.Count() & ")", GroupIcon, GroupIcon)
                trvNode_Tables.Tag = FolderBinding
                trvNode_Tables.ToolTipText = trvNode_Tables.Text

                trvNode_Queries = .Nodes.Add(DbName & "Queries", "Queries (" & DB.Queries.Count() & ")", GroupIcon, GroupIcon)
                trvNode_Queries.Tag = FolderBinding
                trvNode_Queries.ToolTipText = trvNode_Queries.Text

            End With


            'next is to add Tables
            For Each Tbl In DB.Tables.OrderBy(Function(T) T.TableName)
                Application.DoEvents()
                Dim tbl_Node As TreeNode = trvNode_Tables.Nodes.Add(Tbl.TableName, Tbl.TableName, TableIcon, TableIcon)
                tbl_Node.ToolTipText = Tbl.TableName
                tbl_Node.ContextMenuStrip = CntxtMnu_TrView_dbObjects_Tblqry
                tbl_Node.Tag = TalbeBinding

                Dim asterisk = tbl_Node.Nodes.Add("All_" & Tbl.TableName & "_Cols", "*")
                asterisk.ToolTipText = "*"
                asterisk.Tag = FieldBinding

                For Each Col As DataColumn In (From c As DataColumn In Tbl.Columns Select c Order By c.ColumnName)
                    Application.DoEvents()
                    tbl_Node.Nodes.Add(GetNodeByDataType(Col, FieldBinding))
                Next
            Next

            'next is to add Queries
            For Each qry In DB.Queries
                Application.DoEvents()
                Dim qry_Node As TreeNode = trvNode_Queries.Nodes.Add(qry.TableName, qry.TableName, QueryIcon, QueryIcon)
                qry_Node.ToolTipText = qry.TableName
                qry_Node.ContextMenuStrip = CntxtMnu_TrView_dbObjects_Tblqry
                qry_Node.Tag = QueryBinding

                Dim asterisk = qry_Node.Nodes.Add("All_" & qry.TableName & "_Cols", "*")
                asterisk.ToolTipText = "*"
                asterisk.Tag = FieldBinding

                For Each Col As DataColumn In (From c As DataColumn In qry.Columns Select c Order By c.ColumnName)
                    Application.DoEvents()
                    qry_Node.Nodes.Add(GetNodeByDataType(Col, FieldBinding))
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            Me.status_CurTask.Text = "Ready"
        End Try

    End Sub

    Function GetNodeByDataType(ByVal Col As DataColumn, _
                               ByVal BindingObj As NodeBindingData) As TreeNode

        Dim _Node As New TreeNode(Col.ColumnName)
        Dim ImgIndex As Integer = 0

        Select Case Col.DataType.Name.ToUpper

            Case GetType(String).Name.ToUpper, GetType(System.Guid).Name.ToUpper
                ImgIndex = TextIcon

            Case GetType(Boolean).Name.ToUpper
                ImgIndex = BooleanIcon

            Case GetType(Byte).Name.ToUpper, GetType(Int16).Name.ToUpper, _
                 GetType(Int32).Name.ToUpper, GetType(Int64).Name.ToUpper, _
                 GetType(Integer).Name.ToUpper, GetType(Long).Name.ToUpper, _
                 GetType(Decimal).Name.ToUpper, GetType(Single).Name.ToUpper, _
                 GetType(Double).Name.ToUpper

                ImgIndex = NumberIcon

            Case GetType(Date).Name.ToUpper
                ImgIndex = DateIcon

            Case "Byte[]".ToUpper
                ImgIndex = BinarryArrayIcon
            Case Else
                ImgIndex = 0
        End Select


        _Node.ToolTipText = Col.ColumnName
        _Node.ImageIndex = ImgIndex
        _Node.SelectedImageIndex = ImgIndex
        _Node.Tag = BindingObj

        Return _Node
    End Function
 

    Sub BringColorsFromSettings()
        ColoringWords.ColoredWords.SqlLiteralOperatorsColor = My.Settings.LiteralOperatorsColor
        ColoringWords.ColoredWords.FunctionsColor = My.Settings.FunctionsColor
        ColoringWords.ColoredWords.SqlOperatorsColor = My.Settings.OperatorsColor
        ColoringWords.ColoredWords.KeyWordsColor = My.Settings.KeyWordsColor
        ColoringWords.ColoredWords.StringsColor = My.Settings.StringsColor
        ColoringWords.ColoredWords.CommentsColor = My.Settings.CommentsColor
        ColoringWords.ColoredWords.SqlPunctuationColor = My.Settings.PunctuationColor

    End Sub

#End Region

#Region "Controls Events Handlers"

#Region "Rtext RichText"

#Region "Drag Drop Stuff"

    Private Sub Rtext_Doc_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Rtext_Doc.DragDrop
        'Application.DoEvents()
        Dim DragedTxt As String = e.Data.GetData(GetType(String))
        If DragedTxt IsNot Nothing Then
            'Me.Rtext_Doc.SelectedText = DragedTxt
            'Me.Rtext_Doc.RecolorWord("Rtext_Doc_DragDrop", True)

        End If
    End Sub

    Private Sub Rtext_Doc_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Rtext_Doc.DragEnter

        'If e.Data.GetDataPresent(DataFormats.Text) Then
        '    e.Effect = DragDropEffects.Copy
        'Else
        '    e.Effect = DragDropEffects.None
        'End If
    End Sub
#End Region

    Private Sub Rtext_Doc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) _
        Handles Rtext_Doc.KeyDown

        If e.Shift = True AndAlso e.KeyCode = Keys.Insert Then
            e.SuppressKeyPress = True
            Me.Rtext_Doc.Paste()
        End If

    End Sub


    Private Sub Rtext_Doc_OverType_Changed(ByVal sender As Object, ByVal OverTypeState As Boolean) Handles Rtext_Doc.OverType_Changed
        Me.Status_Ins_Ovr.Text = If(OverTypeState, "OVR", "INS") 
    End Sub

    Private Sub Rtext_Doc_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Rtext_Doc.SelectionChanged
        Me.status_CurCol.Text = "Col " & Me.Rtext_Doc.CurrColNumber
        Me.status_CurLine.Text = "Ln " & Me.Rtext_Doc.CurrLineNumber + 1
    End Sub

#End Region

#Region "TrView_dbObjects"

#Region "Drag Drop Stuff"

    Private Sub TrView_dbObjects_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TrView_dbObjects.DragEnter
        e.Effect = e.AllowedEffect 'to change the Cursor
    End Sub

    Private Sub TrView_dbObjects_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles TrView_dbObjects.ItemDrag

        Dim Node = CType(e.Item, TreeNode)

        If Node.Tag Is Nothing Then Exit Sub

        Dim Binde = CType(Node.Tag, NodeBindingData)

        Select Case Binde.NodeType
            Case NodeTypes.TableNode, NodeTypes.QueryNode
                TrView_dbObjects.DoDragDrop(" [" & Node.Text.Trim & "] ", DragDropEffects.All)

            Case NodeTypes.FieldNode
                If Node.Text.Trim = "*" Then ' we can't put [*] , it has to be just "*"
                    TrView_dbObjects.DoDragDrop(" * ", DragDropEffects.All)
                Else
                    TrView_dbObjects.DoDragDrop(" [" & Node.Parent.Text & "].[" & Node.Text.Trim & "] ", DragDropEffects.All)
                End If
            Case NodeTypes.FunctionNode
                TrView_dbObjects.DoDragDrop(" " & Node.Text.Trim & " ( ) ", DragDropEffects.All)
            Case Else
                Exit Sub
        End Select

    End Sub
#End Region

#End Region

#Region "dgv_OutPut"

    Private Sub dgv_OutPut_RowsCountChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dgv_OutPut.RowsCountChanged
        'this way, we don't need to update the status bar
        'when a gird is filled or cleared, it will automatically
        'update it with the new Rows count
        status_OutPutCount.Text = dgv_OutPut.Rows.Count & " Row(s)"

    End Sub
#End Region

    Private Sub btn_Exec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Exec.Click
        Exec(sender)
    End Sub

#End Region

#Region "Form Events Handlers"

    Private Sub frm_Editor_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
        My.Settings.Reload()

        BringColorsFromSettings()

        Me.Rtext_Doc.Font = My.Settings.EditorFont

        Me.Rtext_Doc.AllowDrop = True ' sorry , but I can't make it at design time

        '-------------
        Me.ToolBar_Connect.Image = New Bitmap(GetType(BindingSource), "BindingSource.bmp")
        Me.mnu_Connect.Image = Me.ToolBar_Connect.Image
        Me.Cmnu_DB_Connect.Image = Me.ToolBar_Connect.Image
        '------------


        'populate the TreeView ImageList with proper Images and set the Indices
        'to use them later

        Dim ImgLst As New ImageList
        With ImgLst.Images
            .Add(New Bitmap(1, 1))


            .Add(My.Resources.Folder) : GroupIcon = .Count - 1

            .Add(My.Resources.FunctionImg) : FunctionIcon = .Count - 1
            .Add(My.Resources.DB.ToBitmap) : DBIcon = .Count - 1
            .Add(New Bitmap(GetType(DataGrid), "DataGrid.bmp")) : TableIcon = .Count - 1
            .Add(New Bitmap(GetType(DataGrid), "DataGrid.bmp")) : QueryIcon = .Count - 1
            .Add(New Bitmap(GetType(TextBox), "TextBox.bmp")) : TextIcon = .Count - 1
            .Add(New Bitmap(GetType(NumericUpDown), "NumericUpDown.bmp")) : NumberIcon = .Count - 1
            .Add(New Bitmap(GetType(DateTimePicker), "DateTimePicker.bmp")) : DateIcon = .Count - 1
            .Add(New Bitmap(GetType(CheckBox), "CheckBox.bmp")) : BooleanIcon = .Count - 1
            .Add(My.Resources.BinaryObject) : BinarryArrayIcon = .Count - 1



        End With
        ImgLst.TransparentColor = Color.Fuchsia
        Me.TrView_dbObjects.ImageList = ImgLst

        'DB PlaceHolder
        Me.TrView_dbObjects.Nodes(0).ImageIndex = DBIcon
        Me.TrView_dbObjects.Nodes(0).SelectedImageIndex = DBIcon
        Me.TrView_dbObjects.Nodes(0).Tag = New NodeBindingData(NodeTypes.DataBaseNode)

        '-------------

        With Me.TrView_dbObjects.Nodes(1)

            .ImageIndex = FunctionIcon
            .SelectedImageIndex = FunctionIcon
            .Tag = New NodeBindingData(NodeTypes.FolderNode)


            For Each GropNode As TreeNode In .Nodes
                GropNode.ImageIndex = GroupIcon
                GropNode.SelectedImageIndex = GroupIcon
                GropNode.Tag = New NodeBindingData(NodeTypes.FolderNode)

                For Each ItemNode As TreeNode In GropNode.Nodes
                    ItemNode.ImageIndex = FunctionIcon
                    ItemNode.SelectedImageIndex = FunctionIcon
                    ItemNode.Tag = New NodeBindingData(NodeTypes.FunctionNode)

                Next
            Next

        End With

        Show()
        Me.Rtext_Doc.Focus()
    End Sub

    Private Sub frm_Editor_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If (e.CloseReason And CloseReason.ApplicationExitCall = CloseReason.ApplicationExitCall) Or _
           (e.CloseReason And CloseReason.UserClosing = CloseReason.UserClosing) Then

            If Me.Rtext_Doc.TextModified Then

                Dim QuestAnswer As MsgBoxResult
                QuestAnswer = Me.Rtext_Doc.Ask4Save()

                Select Case QuestAnswer
                    Case MsgBoxResult.Yes
                        Me.Rtext_Doc.Save()
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

                My.Settings.EditorFont = Me.Rtext_Doc.Font
                My.Settings.Save()
            End If


        End If



    End Sub
#End Region

#Region "Menues Commands Events Handlers"
#Region "Main Menue"

#Region "File"

    Private Sub mnu_File_NewQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_File_NewQuery.Click
        NewQuery("mnu_File_NewQuery_Click")
    End Sub

    Private Sub mnu_Connect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles mnu_Connect.Click, ToolBar_Connect.Click, Cmnu_DB_Connect.Click
        Connect(sender)
    End Sub

    Private Sub mnu_Open_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
            Handles mnu_Open.Click, ToolBar_Open.Click
        Me.status_CurTask.Text = "Fetching the new file.........."
        Rtext_Doc.OpenNewFile()
        Me.status_CurTask.Text = "Ready"
    End Sub

    Private Sub mnu_SaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_SaveAs.Click
        Me.Rtext_Doc.SaveAs()
    End Sub

    Private Sub mnu_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
            Handles mnu_Save.Click, ToolBar_Save.Click

        Me.Rtext_Doc.Save()
    End Sub

    Private Sub mnu_Exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_Exit.Click
        Application.Exit()
    End Sub
#End Region

#Region "View"

    Private Sub mnu_ObjectsBrowser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_ObjectsBrowser.Click
        Split_ObjBrowser_And_Others.Panel1Collapsed = Not Split_ObjBrowser_And_Others.Panel1Collapsed
        mnu_ObjectsBrowser.Checked = Not mnu_ObjectsBrowser.Checked
    End Sub

    Private Sub mnu_ShowOutPutPanel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_ShowOutPutPanel.Click
        mnu_ShowOutPutPanel.Checked = Not mnu_ShowOutPutPanel.Checked
        ShowHideOutput(mnu_ShowOutPutPanel.Checked)

    End Sub

#End Region

#Region "Format"
    Private Sub mnu_WordWrap_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnu_WordWrap.CheckedChanged
        Me.Rtext_Doc.WordWrap = mnu_WordWrap.Checked
    End Sub

    Private Sub mnu_Font_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_Font.Click
        Using FontDlg As New FontDialog

            With FontDlg
                .ShowColor = False
                .ShowEffects = False
                .ShowApply = False 'don't wanna add EventHandler 4 it , so he has to click "OK"
                .Font = Me.Rtext_Doc.Font
                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                    Me.Rtext_Doc.Font = .Font
                End If
            End With

        End Using

    End Sub

    Private Sub EditorOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
          Handles Cmnu_EditorOptions.Click, mnu_EditorOptions.Click

        Using EditorOptions As New frm_EditorOptions
            With EditorOptions
                If .ShowDialog() = Windows.Forms.DialogResult.OK Then
                    ColoredWords.Initiate(Me.Name & "_ChangesEditorSettings")
                    Me.Rtext_Doc.RecolorWords("ContainerForm_ChangesEditorSettings")
                End If

            End With
        End Using

    End Sub

#End Region

#Region "Edit"
    Private Sub mnu_Undo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
            Handles mnu_Undo.Click, Cmnu_Undo.Click

        If Me.Rtext_Doc.CanUndo Then
            Me.Rtext_Doc.Undo()
        End If
    End Sub

    Private Sub mnu_Redo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                    Handles mnu_Redo.Click, Cmnu_Redo.Click
        If Me.Rtext_Doc.CanRedo Then
            Me.Rtext_Doc.Redo()
        End If
    End Sub

    Private Sub Edit_DropDownOpening() _
            Handles mnu_Edit.DropDownOpening, ContextMnu_RText.Opening

        Me.mnu_Undo.Enabled = Me.Rtext_Doc.CanUndo
        Me.mnu_Redo.Enabled = Me.Rtext_Doc.CanRedo

        Me.Cmnu_Undo.Enabled = Me.Rtext_Doc.CanUndo
        Me.Cmnu_Redo.Enabled = Me.Rtext_Doc.CanRedo
        '-------------------


        'I could use it every time like it is ... 
        'but this is for more optimization
        'it will calculate it only once
        Dim HasSelection As Boolean = CBool(Me.Rtext_Doc.SelectedText.Length > 0)
        Me.mnu_Cut.Enabled = HasSelection
        Me.mnu_Copy.Enabled = HasSelection

        Me.Cmnu_Cut.Enabled = HasSelection
        Me.Cmnu_Copy.Enabled = HasSelection

        '---------------------------------
        'next is to Enable/Disable the [Select All] menu item
        Me.mnu_SelectAll.Enabled = CBool(Me.Rtext_Doc.Text <> "")
        Me.Cmnu_SelectAll.Enabled = Me.mnu_SelectAll.Enabled

    End Sub

    Private Sub mnu_Cut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
            Handles mnu_Cut.Click, Cmnu_Cut.Click, ToolBar_Cut.Click
        Try
            Me.Rtext_Doc.Cut()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub mnu_Copy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
                Handles mnu_Copy.Click, Cmnu_Copy.Click, ToolBar_Copy.Click
        Try
            Me.Rtext_Doc.Copy()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Paste_Clicks(ByVal sender As System.Object, ByVal e As System.EventArgs) _
            Handles mnu_Paste.Click, Cmnu_Paste.Click, ToolBar_Paste.Click

        Me.Rtext_Doc.Paste()
        'Try

        '    If Clipboard.ContainsText Then
        '        'we're not gonna use Me.Rtext_Doc.Paste()
        '        'cuz it's gonna past it as it is 
        '        '(it could has some Formating , user is pasting from 'MS Word' or somethig)
        '        'so we'll deal with it our way
        '        Me.Rtext_Doc.SelectedText = Clipboard.GetText(TextDataFormat.Text)

        '    End If
        'Catch ex As Exception

        'End Try
    End Sub

    Private Sub SelectAll_Click() _
            Handles Cmnu_SelectAll.Click, mnu_SelectAll.Click

        Me.Rtext_Doc.SelectAll()

    End Sub

    Private Sub mnu_FindReplace_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles mnu_Replace.Click, mnu_Find.Click

        If Not frm_FindReplace.Opened Then
            FindReplaceForm = New frm_FindReplace(Rtext_Doc)

            FindReplaceForm.Owner = Me
            FindReplaceForm.SelectTab(If(sender Is mnu_Replace, True, False))

            FindReplaceForm.Show()
        Else
            FindReplaceForm.Activate()
            FindReplaceForm.SelectTab(If(sender Is mnu_Replace, True, False))
        End If
    End Sub

#End Region

#Region "Output"

    Private Sub mnu_ClearOutPut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_ClearOutPut.Click
        If MsgBox("Are you sure you want to clear all outputs?", _
                   MsgBoxStyle.Question Or MsgBoxStyle.YesNo, _
                   "Caution") = MsgBoxResult.Yes Then

            ClearOutput("mnu_ClearOutPut_Click")
        End If

    End Sub

    Private Sub mnu_Export2Excel(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles mnu_Export2ExcelByHTML.Click, mnu_Export2ExcelByXML.Click, _
                mnu_Export2ExcelByExcelApp.Click

        Using SaveFile As New SaveFileDialog
            With SaveFile

                .Filter = "Excel 2003 Books (*.xls)|*.xls"
                .Title = "Select the location and the name of the File to export to"

                If .ShowDialog = System.Windows.Forms.DialogResult.OK Then
                    MsgBox("This operation could take a while ...." & _
                           Environment.NewLine & Environment.NewLine & _
                           "Please wait.........", , "Caution")

                    Me.Cursor = Cursors.WaitCursor
                    Me.dgv_OutPut.Cursor = Cursors.WaitCursor
                    Me.status_CurTask.Text = "Exporting To Excel......."
                    Application.DoEvents()
                    frm_Working.Show("Exporting To Excel.......")
                    Try
                        Dim dt = CType(Me.dgv_OutPut.DataSource, DataTable)
                        Select Case sender.Name
                            Case mnu_Export2ExcelByExcelApp.Name
                                If Excel.ExportByExecl(.FileName, dt) Then
                                    MsgBox("Creating Excel File Done", MsgBoxStyle.Information, "Success")
                                End If

                            Case mnu_Export2ExcelByHTML.Name
                                If Excel.ExportByHtml(.FileName, dt) Then
                                    MsgBox("Creating Excel File Done", MsgBoxStyle.Information, "Success")
                                End If

                            Case mnu_Export2ExcelByXML.Name
                                If Excel.ExportByXML(.FileName, dt) Then
                                    MsgBox("Creating Excel File Done", MsgBoxStyle.Information, "Success")
                                End If
                        End Select


                    Catch ex As Exception

                        MsgBox("Seems like Excel 2003 is not installed or damaged" & Environment.NewLine & _
                               "You're not gonna be able to export it unless you have it installed." & Environment.NewLine & Environment.NewLine & _
                               "Technical Info:" & Environment.NewLine & ex.Message, _
                              MsgBoxStyle.Exclamation, "Caution")
                        Exit Sub
                        Application.DoEvents()
                    Finally

                        Me.status_CurTask.Text = "Ready"
                        Me.Cursor = Cursors.Default
                        Me.dgv_OutPut.Cursor = Cursors.Default
                        frm_Working.Close()

                    End Try

                End If
            End With
        End Using
    End Sub

    Private Sub mnu_Output_DropDownOpening(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnu_Output.DropDownOpening

        Me.mnu_ClearOutPut.Enabled = CBool(Me.dgv_OutPut.Rows.Count > 0)
        Me.mnu_Export2ExcelByHTML.Enabled = CBool(Me.dgv_OutPut.Rows.Count > 0)
        Me.mnu_Export2ExcelByXML.Enabled = CBool(Me.dgv_OutPut.Rows.Count > 0)
        Me.mnu_Export2ExcelByExcelApp.Enabled = CBool(Me.dgv_OutPut.Rows.Count > 0)

    End Sub

#End Region

#Region "Query"

    Private Sub mnu_Exec_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_Exec.Click
        Exec(sender)
    End Sub

    Private Sub mnu_Query_Parameters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_Query_Parameters.Click

        Dim ParamsForm As New frm_AddParams
        ParamsForm.ShowDialog()

    End Sub

    Private Sub mnu_Convert2Code_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_Convert2Code.Click
        Using ConverterForm As New frm_Text2Code
            With ConverterForm
                .Text2Convert = Me.Rtext_Doc.Query.Sql
                .ShowDialog()
            End With
        End Using
    End Sub

#End Region

#Region "Help"

    Private Sub mnu_TutorialFeatures_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_TutorialFeatures.Click

        Using SveDlg As New SaveFileDialog
            With SveDlg
                .Filter = "Archive files (*.RAR)|*.RAR"
                .Title = "Select a folder to save the tutorial file to"
                If SveDlg.ShowDialog = Windows.Forms.DialogResult.OK Then

                    My.Computer.FileSystem.WriteAllBytes(SveDlg.FileName, _
                                                         My.Resources.Tutorial, False)
                End If
            End With
        End Using

    End Sub

    Private Sub mnu_About_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnu_About.Click
        AboutBox.ShowDialog()
    End Sub

#End Region
#End Region

#Region "Context Menues Events Handlers"

#Region "CntxtMnu_TrView_dbObjects_DB Events Handlers"
    Private Sub Cmnu_DB_Refresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmnu_DB_Refresh.Click

        frm_Working.Show("Refreshing DB Schema.......")
        Try
            DB.RefreshSchema()
            FillDbObjects()

        Catch ex As Exception
            Application.DoEvents()
        Finally
            frm_Working.Close()
        End Try
    End Sub
    Private Sub CntxtMnu_TrView_dbObjects_DB_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles CntxtMnu_TrView_dbObjects_DB.Opening
        Cmnu_DB_Refresh.Enabled = CBool(DB.Path <> "")

    End Sub
#End Region

#Region "CntxtMnu_TrView_dbObjects_TblQry Events Handlers"

    Private Sub CntxtMnu_TrView_dbObjects_TblQry_GenSelectALLScript_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CntxtMnu_TrView_dbObjects_TblQry_GenSelectALLScript.Click
        Try
            If TrView_dbObjects.LastUsedNode IsNot Nothing Then

                Dim Bind = CType(TrView_dbObjects.LastUsedNode.Tag, NodeBindingData)
                If Bind.NodeType = NodeTypes.TableNode Or _
                   Bind.NodeType = NodeTypes.QueryNode Then

                    Me.Rtext_Doc.SelectedText = "Select * From [" & TrView_dbObjects.LastUsedNode.Text & "] ;"
                End If

            End If
        Catch ex As Exception
            MsgBox("Error encountered", MsgBoxStyle.Exclamation, "Caution")
        End Try
    End Sub

    Private Sub CntxtMnu_TrView_dbObjects_Tblqry_GenDrop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CntxtMnu_TrView_dbObjects_Tblqry_GenDrop.Click
        Try
            If TrView_dbObjects.LastUsedNode IsNot Nothing Then

                Dim Bind = CType(TrView_dbObjects.LastUsedNode.Tag, NodeBindingData)
                If Bind.NodeType = NodeTypes.TableNode Or _
                   Bind.NodeType = NodeTypes.QueryNode Then

                    Rtext_Doc.SelectedText = "Drop Table [" & TrView_dbObjects.LastUsedNode.Text & "] ;"
                End If

            End If
        Catch ex As Exception
            MsgBox("Error encountered", MsgBoxStyle.Exclamation, "Caution")
        End Try
    End Sub

#End Region

#End Region

#End Region

 

End Class