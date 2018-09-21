Imports System.Text.RegularExpressions
Imports System.Text

Public Class frm_FindReplace

    Private WithEvents RtextBox As New ColoringWords.DevRichTextBox

#Region "Fields"

    ''' <summary>used to determine if there is more than one instance of this Form-class opened</summary>
    Public Shared Opened As Boolean = False

    ''' <summary> holds the found matches of the search </summary>
    Dim Founds As List(Of Integer)

    ''' <summary>holds the navigation pos in found matches</summary>
    Dim FindPos As Integer = 0

    ''' <summary>turns true if user changes one/more of find elements</summary>
    Dim ReFind As Boolean = True

    ''' <summary>
    ''' user can't change the Height , But I can , 
    ''' so this is used to set the height that user has to stick with
    ''' </summary>
    Dim ForcedHeight As Single = 0

    ''' <summary>
    ''' user can't change the Height , But I can , 
    ''' this is var toggles to determine who's changing the height now
    ''' </summary>
    Dim HeightChangeable As Boolean

    ''' <summary>
    ''' this var is used when we perform some movements ,
    ''' it's used in more than place,
    ''' so we made it as a public to be able to change it 
    ''' without affecting codes that use it
    ''' </summary>
    ''' <remarks></remarks>
    Const SpaceBetweenTools = 40

#End Region

#Region "Form Events Handlers"

    ''' <summary>
    '''creates a new instance of the Form-class
    ''' </summary>
    ''' <param name="TargetRichTextBox">just a pointer to the RichTextBox that will be searched </param>
    Public Sub New(ByVal TargetRichTextBox As ColoringWords.DevRichTextBox)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.RtextBox = TargetRichTextBox
        ForcedHeight = btn_ExpandOptions.Bottom + Status.Height + SpaceBetweenTools
        Opened = True

    End Sub

    Private Sub frm_Find_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Opened = False
    End Sub

    Private Sub frm_Find_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown, MyBase.KeyUp
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If

        If e.Control Then
            If e.KeyCode = Keys.H Then
                TabFindReplace.SelectedTab = TabPage_Replace
            ElseIf e.KeyCode = Keys.F Then
                TabFindReplace.SelectedTab = TabPage_Find
            End If
        End If
    End Sub

    Private Sub frm_Find_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.SizeChanged
        If Me.HeightChangeable Then
            ' now I'm changing the size then let me do what I want
        Else
            ' now the user is trying to change the size
            ' honey ..... there are limits
            ' you can change the Width , but the Height is mine

            SetMaxMinSize()

        End If
    End Sub

    Private Sub frm_Find_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        lbl_Cover.BorderStyle = BorderStyle.None ' it was FixedSingle just to see it in the design time
        txt_FindStr_TextChanged(txt_FindStr, e) 't set the "Enabled" property of "btn_FindNext"

        '-----------------------

        Me.chk_MatchCase.Checked = False
        Me.chk_SearchUp.Checked = False

    End Sub
#End Region

#Region "Controls Events Handlers"

    Private Sub btn_Close_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Close.Click
        Me.Close()
    End Sub

    Private Sub btn_FindNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_FindNext.Click
        If ReFind Then

            Founds = GetFounds()
            If Founds Is Nothing Then Exit Sub ' no row is found , a msg is already shown

            FindPos = 0
            ReFind = False
        End If

        ' none of Find  conditions is changed so show the next found row

        If Not chk_SearchUp.Checked Then 'Down
            If FindPos + 1 <= Founds.Count - 1 Then
                FindPos += 1
            Else
                FindPos = 0
            End If
        Else 'chk_SearchUp.Checked = True ie : Up 

            If FindPos - 1 >= 0 Then
                FindPos -= 1
            Else
                FindPos = Founds.Count - 1
            End If
        End If


        FindStatus.Text = String.Format("{0} Of {1}", FindPos + 1, Founds.Count)
        CType(Me.Owner, frm_Editor).Rtext_Doc.Select(Founds(FindPos), _
                                                     txt_FindStr.Text.Length)


    End Sub

    Private Sub btn_Replace_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_Replace.Click
        Replace(Me.txt_FindStr.Text, Me.txt_ReplaceWith.Text, False)
    End Sub

    Private Sub btn_ReplaceALL_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btn_ReplaceALL.Click
        Replace(Me.txt_FindStr.Text, Me.txt_ReplaceWith.Text, True)
    End Sub

#End Region

#Region "TabFindReplace"

    Private Sub TabFindReplace_SelectedIndexChanged(ByVal Sender As Object, ByVal e As System.EventArgs) Handles TabFindReplace.SelectedIndexChanged

        Me.Pnl_Btns.Parent = Me.TabFindReplace.SelectedTab
        Select Case Me.TabFindReplace.SelectedIndex
            Case 0 'Find

                Me.btn_Replace.Hide()
                Me.btn_ReplaceALL.Hide()
                Me.lbl_ReplaceWith_lbl.Hide()
                Me.txt_ReplaceWith.Hide()

            Case 1 'Replace
                Me.btn_Replace.Show()
                Me.btn_ReplaceALL.Show()
                Me.lbl_ReplaceWith_lbl.Show()
                Me.txt_ReplaceWith.Show()

        End Select

        ReFind = True
    End Sub

    Public Sub SelectTab(ByVal ReplaceOperation As Boolean)

        If ReplaceOperation Then
            TabFindReplace.SelectedTab = TabPage_Replace
            TabFindReplace_SelectedIndexChanged(TabFindReplace, New EventArgs)

        Else 'it's a finde operation,
            'but there is no need to change the selcted tab 
            'cuz it's already on the find one

        End If

    End Sub

#End Region

#Region "btn_ExpandOptions"

    Private Sub btn_ExpandOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_ExpandOptions.Click

        'this sub contains some loops , but they are not too long , 
        'so won't use "Application.DoEvents()"
        '"Application.DoEvents()" relieves the processor , 
        ' but it causes the form to redraw itself, 
        ' and that's bad because it has a lot of 
        'controls whoes "Anchor" Property is set to make more rules

        Me.HeightChangeable = True
        Me.MinimumSize = New Size(0, 0) ' to be able to miminize it the way I want
        Me.MaximumSize = New Size(0, 0) ' to be able to miminize it the way I want

        Me.btn_ExpandOptions.Enabled = False ' prevent user from clicking again till I finsh
        Select Case btn_ExpandOptions.Text
            Case "+"
                Me.ForcedHeight = Me.Grop_FindOptions.Bottom + Status.Height + SpaceBetweenTools

                btn_ExpandOptions.Text = "-"
                lbl_Cover.Visible = False 'show the unshown sides of Grop_FindOptions 

                Do While Me.Height + 20 <= Me.ForcedHeight
                    'Application.DoEvents()
                    Me.Height += 20
                Loop
                Me.Height = Me.ForcedHeight

            Case "-"

                Me.ForcedHeight = Me.btn_ExpandOptions.Bottom + Status.Height + SpaceBetweenTools
                btn_ExpandOptions.Text = "+"

                Do While Me.Height - 20 >= Me.ForcedHeight
                    'Application.DoEvents()
                    Me.Height -= 20
                Loop

                Me.Height = Me.ForcedHeight
                lbl_Cover.Visible = True ' to covere the unwanted sides of the Grop_FindOptions

        End Select

        Me.HeightChangeable = False
        Me.SetMaxMinSize()
        Me.btn_ExpandOptions.Enabled = True
    End Sub

    Private Sub btn_ExpandOptions_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim W, H As Single
        W = btn_ExpandOptions.ClientRectangle.Width
        H = btn_ExpandOptions.ClientRectangle.Height

        If sender.Text = "-" Then

            ControlPaint.DrawCaptionButton( _
                     e.Graphics, New Rectangle(0, 0, W, H), _
                     CaptionButton.Restore, ButtonState.Normal _
                                           )
        Else

            ControlPaint.DrawCaptionButton( _
                      e.Graphics, New Rectangle(0, 0, W, H), _
                      CaptionButton.Maximize, ButtonState.Normal _
                    )
        End If
    End Sub
#End Region

#Region "txt_FindStr"

    Private Sub txt_FindStr_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txt_FindStr.TextChanged
        Me.btn_FindNext.Enabled = (Me.txt_FindStr.Text <> String.Empty)

    End Sub

    Private Sub txt_FindStr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txt_FindStr.KeyUp
        If e.KeyCode = Keys.Return Then
            Me.btn_FindNext.PerformClick()
        End If
    End Sub

#End Region

#Region "Added Subs"

    Sub SetMaxMinSize()
        Me.MaximumSize = New Size(700, Me.ForcedHeight)
        Me.MinimumSize = New Size(347, Me.MaximumSize.Height)
    End Sub

    Sub MsgNotFound(ByVal WantedItem As String)
        MsgBox("No Results detected " & _
               "for '" & WantedItem & "'.", _
               MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, _
               "Find")
    End Sub

    Sub Replace(ByVal FindWhat As String, ByVal ReplaceWith As String, ByVal All As Boolean)

        Founds = GetFounds()

        If Founds Is Nothing Then
            'here we have to show a msg , but "GetFoundRows" deals with it
            Exit Sub
        End If

        '-------------

        If MsgBox("Please confirm the Replacement." & Environment.NewLine & _
             If(All, Founds.Count, 1) & " Match(s) will be affected" & _
             Environment.NewLine & ControlChars.Lf & _
             "Replace ?", MsgBoxStyle.YesNo, _
             "Replacement Confirmation") = MsgBoxResult.No Then

            Exit Sub
        End If


        If All Then


            Dim StrBldr As New StringBuilder(RtextBox.Text)

            Dim AddedLen = 0
            For Each Index In Founds
                StrBldr.Remove(AddedLen + Index, FindWhat.Length)
                StrBldr.Insert(AddedLen + Index, ReplaceWith)

                AddedLen += ReplaceWith.Length - FindWhat.Length
            Next
            RtextBox.Text = StrBldr.ToString
            StrBldr = Nothing

        Else 'All = False 
            RtextBox.Select(Founds(FindPos), FindWhat.Length)
            RtextBox.SelectedText = ReplaceWith
        End If


    End Sub

    Private Function GetFounds() As List(Of Integer)


        Dim Rtn As New List(Of Integer)

        Dim PreSelectionStart = RtextBox.SelectionStart
        Dim PreSelectionLen = RtextBox.SelectionLength

        '-------------------------
        Dim SearchOptions = RichTextBoxFinds.None

        SearchOptions = SearchOptions Or If(chk_MatchCase.Checked, _
                                            RichTextBoxFinds.MatchCase, _
                                            SearchOptions)

        SearchOptions = SearchOptions Or If(chk_MatchWholeWord.Checked, _
                                            RichTextBoxFinds.WholeWord, _
                                            SearchOptions)

        With RtextBox
            .DisableDrawing() 'using the RtextBox.Find will Select found elements

            Dim tmpIndex = .Find(txt_FindStr.Text, .Text.Length, SearchOptions)
            While tmpIndex > -1
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
             Handles txt_FindStr.TextChanged, RtextBox.TextChanged, _
                     chk_MatchCase.CheckedChanged, _
                     chk_MatchWholeWord.CheckedChanged, _
                     chk_SearchUp.CheckedChanged

        Me.ReFind = True

    End Sub

#End Region

End Class