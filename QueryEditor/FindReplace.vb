Imports System.Text
Imports QueryEditor.ColoringWords

Public Class FindReplace
    Private WithEvents _richText As New DevRichTextBox

#Region "Fields"

    ''' <summary>used to determine if there is more than one instance of this Form-class opened</summary>
    Public Shared Opened As Boolean = False

    ''' <summary> holds the found matches of the search </summary>
    Dim _founds As List(Of Integer)

    ''' <summary>holds the navigation pos in found matches</summary>
    Dim _findPos As Integer = 0

    ''' <summary>turns true if user changes one/more of find elements</summary>
    Dim _reFind As Boolean = True

    ''' <summary>
    '''     user can't change the Height , But I can ,
    '''     so this is used to set the height that user has to stick with
    ''' </summary>
    Dim _forcedHeight As Single = 0

    ''' <summary>
    '''     user can't change the Height , But I can ,
    '''     this is var toggles to determine who's changing the height now
    ''' </summary>
    Dim _heightChangeable As Boolean

    ''' <summary>
    '''     this var is used when we perform some movements ,
    '''     it's used in more than place,
    '''     so we made it as a public to be able to change it
    '''     without affecting codes that use it
    ''' </summary>
    ''' <remarks></remarks>
    Const SpaceBetweenTools = 40

#End Region

#Region "Form Events Handlers"

    ''' <summary>
    '''     creates a new instance of the Form-class
    ''' </summary>
    ''' <param name="targetRichTextBox">just a pointer to the RichTextBox that will be searched </param>
    Public Sub New(targetRichTextBox As DevRichTextBox)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _richText = TargetRichTextBox
        _forcedHeight = btn_ExpandOptions.Bottom + Status.Height + SpaceBetweenTools
        Opened = True
    End Sub

    Private Shared Sub Frm_Find_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Opened = False
    End Sub

    Private Sub Frm_Find_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown, MyBase.KeyUp
        If e.KeyCode = Keys.Escape Then
            Close()
        End If

        If e.Control Then
            If e.KeyCode = Keys.H Then
                TabFindReplace.SelectedTab = TabPage_Replace
            ElseIf e.KeyCode = Keys.F Then
                TabFindReplace.SelectedTab = TabPage_Find
            End If
        End If
    End Sub

    Private Sub Frm_Find_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        If _heightChangeable Then
            ' now I'm changing the size then let me do what I want
        Else
            ' now the user is trying to change the size
            ' you can change the Width, but not the Height

            SetMaxMinSize()
        End If
    End Sub

    Private Sub Frm_Find_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        lbl_Cover.BorderStyle = BorderStyle.None ' it was FixedSingle just to see it in the design time
        Txt_FindStr_TextChanged(txt_FindStr, e) 't set the "Enabled" property of "btn_FindNext"

        '-----------------------

        chk_MatchCase.Checked = False
        chk_SearchUp.Checked = False
    End Sub

#End Region

#Region "Controls Events Handlers"

    Private Sub Btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Close()
    End Sub

    Private Sub Btn_FindNext_Click(sender As Object, e As EventArgs) Handles btn_FindNext.Click
        If _reFind Then
            _founds = GetFounds()
            If _founds Is Nothing Then Exit Sub ' no row is found , a msg is already shown

            _findPos = 0
            _reFind = False
        End If

        ' none of Find  conditions is changed so show the next found row
        If Not chk_SearchUp.Checked Then 'Down
            If _findPos + 1 <= _founds.Count - 1 Then
                _findPos += 1
            Else
                _findPos = 0
            End If
        Else 'chk_SearchUp.Checked = True ie : Up 
            If _findPos - 1 >= 0 Then
                _findPos -= 1
            Else
                _findPos = _founds.Count - 1
            End If
        End If

        FindStatus.Text = String.Format("{0} Of {1}", _findPos + 1, _founds.Count)
        CType(Owner, Editor).Rtext_Doc.Select(_founds(_findPos),
                                                 txt_FindStr.Text.Length)
    End Sub

    Private Sub Btn_Replace_Click(sender As Object, e As EventArgs) Handles btn_Replace.Click
        Replace(txt_FindStr.Text, txt_ReplaceWith.Text, False)
    End Sub

    Private Sub Btn_ReplaceALL_Click(sender As Object, e As EventArgs) Handles btn_ReplaceALL.Click
        Replace(txt_FindStr.Text, txt_ReplaceWith.Text, True)
    End Sub

#End Region

#Region "TabFindReplace"

    Private Sub TabFindReplace_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles TabFindReplace.SelectedIndexChanged

        Pnl_Btns.Parent = TabFindReplace.SelectedTab
        Select Case TabFindReplace.SelectedIndex
            Case 0 'Find
                btn_Replace.Hide()
                btn_ReplaceALL.Hide()
                lbl_ReplaceWith_lbl.Hide()
                txt_ReplaceWith.Hide()

            Case 1 'Replace
                btn_Replace.Show()
                btn_ReplaceALL.Show()
                lbl_ReplaceWith_lbl.Show()
                txt_ReplaceWith.Show()

        End Select

        _reFind = True
    End Sub

    Public Sub SelectTab(replaceOperation As Boolean)

        If ReplaceOperation Then
            TabFindReplace.SelectedTab = TabPage_Replace
            TabFindReplace_SelectedIndexChanged(TabFindReplace, New EventArgs)

        Else 'it's a find operation,
            'but there is no need to change the selected tab 
            'cuz it's already on the find one

        End If
    End Sub

#End Region

#Region "btn_ExpandOptions"

    Private Sub btn_ExpandOptions_Click(sender As Object, e As EventArgs) Handles btn_ExpandOptions.Click
        btn_ExpandOptions.Enabled = False ' prevent user from clicking again till I finish
        _heightChangeable = True
        'Me.MinimumSize = New Size(0, 0) ' to be able to minimize it the way I want
        'Me.MaximumSize = New Size(0, 0) ' to be able to minimize it the way I want

        Select Case btn_ExpandOptions.Text
            Case "+"
                _forcedHeight = Grop_FindOptions.Bottom + Status.Height + SpaceBetweenTools
                Height = _forcedHeight
                btn_ExpandOptions.Text = "-"
                lbl_Cover.Visible = False 'show the unwanted sides of Group_FindOptions                
            Case "-"
                _forcedHeight = btn_ExpandOptions.Bottom + Status.Height + SpaceBetweenTools
                Height = _forcedHeight
                btn_ExpandOptions.Text = "+"
                lbl_Cover.Visible = True ' to cover the unwanted sides of the Group_FindOptions
        End Select

        SetMaxMinSize()
        _heightChangeable = False
        btn_ExpandOptions.Enabled = True
    End Sub

    ' Not sure what this was for:
    'Private Sub Btn_ExpandOptions_Paint(sender As Object, e As PaintEventArgs)
    '    Dim w, h As Single
    '    W = btn_ExpandOptions.ClientRectangle.Width
    '    H = btn_ExpandOptions.ClientRectangle.Height

    '    If sender.Text = "-" Then
    '        ControlPaint.DrawCaptionButton(
    '            e.Graphics, New Rectangle(0, 0, W, H),
    '            CaptionButton.Restore, ButtonState.Normal
    '            )
    '    Else
    '        ControlPaint.DrawCaptionButton(
    '            e.Graphics, New Rectangle(0, 0, W, H),
    '            CaptionButton.Maximize, ButtonState.Normal
    '            )
    '    End If
    'End Sub

#End Region

#Region "txt_FindStr"

    Private Sub Txt_FindStr_TextChanged(sender As Object, e As EventArgs) Handles txt_FindStr.TextChanged
        btn_FindNext.Enabled = (txt_FindStr.Text <> String.Empty)
    End Sub

    Private Sub Txt_FindStr_KeyUp(sender As Object, e As KeyEventArgs) Handles txt_FindStr.KeyUp
        If e.KeyCode = Keys.Return Then
            btn_FindNext.PerformClick()
        End If
    End Sub

#End Region

#Region "Added Subs"

    Private Sub SetMaxMinSize()
        MaximumSize = New Size(700, _forcedHeight)
        MinimumSize = New Size(347, MaximumSize.Height)
    End Sub

    Private Shared Sub MsgNotFound(wantedItem As String)
        MsgBox("No Results detected " &
               "for '" & WantedItem & "'.",
               MsgBoxStyle.OkOnly Or MsgBoxStyle.Information,
               "Find")
    End Sub

    Private Sub Replace(findWhat As String, replaceWith As String, all As Boolean)
        _founds = GetFounds()
        If _founds IsNot Nothing Then
            If MsgBox("Please confirm the Replacement." & Environment.NewLine &
                      If(All, _founds.Count, 1) & " Match(s) will be affected" & Environment.NewLine & ControlChars.Lf &
                      "Replace ?", MsgBoxStyle.YesNo, "Replacement Confirmation") = MsgBoxResult.Yes Then

                If All Then
                    Dim replacer As New StringBuilder(_richText.Text)
                    Dim addedLen = 0
                    For Each Index In _founds
                        replacer.Remove(AddedLen + Index, FindWhat.Length)
                        replacer.Insert(AddedLen + Index, ReplaceWith)

                        AddedLen += ReplaceWith.Length - FindWhat.Length
                    Next

                    _richText.Text = replacer.ToString
                Else 'All = False 
                    _richText.Select(_founds(_findPos), FindWhat.Length)
                    _richText.SelectedText = ReplaceWith
                End If
            End If
        End If
    End Sub

    Private Function GetFounds() As List(Of Integer)
        Dim rtn As New List(Of Integer)

        Dim preSelectionStart = _richText.SelectionStart
        Dim preSelectionLen = _richText.SelectionLength

        '-------------------------
        Dim searchOptions = RichTextBoxFinds.None

        SearchOptions = SearchOptions Or If(chk_MatchCase.Checked,
                                            RichTextBoxFinds.MatchCase,
                                            SearchOptions)

        SearchOptions = SearchOptions Or If(chk_MatchWholeWord.Checked,
                                            RichTextBoxFinds.WholeWord,
                                            SearchOptions)

        With _richText
            .DisableDrawing() 'using the RtextBox.Find will Select found elements

            Dim tmpIndex = .Find(txt_FindStr.Text, .Text.Length, SearchOptions)
            While tmpIndex > - 1
                If Rtn.Contains(tmpIndex) Then Exit While
                Rtn.Add(tmpIndex)
                tmpIndex = .Find(txt_FindStr.Text, tmpIndex + 1, .Text.Length, SearchOptions)
            End While

            .EnableDrawing() 'using the RtextBox.Find will Select found elements , so Enable Drawing after disabling it
            .Select(PreSelectionStart, PreSelectionLen)
        End With

        If IsNothing(Rtn) OrElse Rtn.Count = 0 Then

            FindStatus.Text = "No Results"
            MsgNotFound(txt_FindStr.Text)
            Return Nothing
        Else
            FindStatus.Text = Rtn.Count & " Match(s)"
        End If
        ' it gets here when Rtn has a value
        Return Rtn
    End Function

    Private Sub ReFind_Changed() _
        Handles txt_FindStr.TextChanged, _richText.TextChanged,
                chk_MatchCase.CheckedChanged,
                chk_MatchWholeWord.CheckedChanged,
                chk_SearchUp.CheckedChanged

        _reFind = True
    End Sub

#End Region
End Class