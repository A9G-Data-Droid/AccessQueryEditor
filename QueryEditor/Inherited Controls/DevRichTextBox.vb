#Region "Imports"

Imports System.Text.RegularExpressions
Imports System.Text
Imports QueryEditor.Queries
Imports System.ComponentModel

#End Region

Namespace ColoringWords

    <ToolboxBitmap(GetType(System.Windows.Forms.RichTextBox))> _
    Public Class DevRichTextBox
        Inherits RichTextBox

#Region "Fields"
        ''' <summary>
        ''' in some certain stances we need to change (Internally) the Text property of the tool,
        ''' but changing the text will affect other actions
        ''' (like Undo and Redo ability) that are not Internal
        ''' so this variable will determine whether the current "Text-Changing Operation" is Internal or not
        ''' </summary>
        Private ToolInternalChanging As Boolean = False

        ''' <summary>
        ''' in some certain stances we need to Recolor all the text in the tool,
        ''' so this is the variable that determine whether to color the CurrentWord or all text
        ''' </summary>
        Private NeedColorAll As Boolean = False

        ''' <summary>
        ''' RecolorWord Sub will change the Selection many times ...., 
        ''' but Changing Selection will Raise OnTextChanged Event which in turn Calls RecolorWord Sub ,
        ''' even Changing the Font will Raise OnTextChanged Event ,......
        ''' thus we'll have a Recursive Calls ,....
        ''' so we use this var to tell the RecolorWord Sub not to Start Coloring again cuz it's already coloring
        ''' </summary>
        Private StartedColoring As Boolean = False
#End Region

#Region "Undo-Redo Handling and Operations"

#Region "UndoObject"
        ''' <summary>
        ''' The original Undo , Redo .....of the tool affect all actions 
        ''' even Font Properties and selecton Properties, 
        ''' so we have to make something instead
        ''' so this is a class to perform the Redo ,Undo
        '''</summary> 
        Private Class UndoObject

            ''' <summary>
            ''' Redo,Unndo is stored a Pair of info 
            ''' (the text it was, and the Caret position it was)
            ''' and this is a Structure to hold this info for each State
            ''' </summary>
            Structure UndoState
                Dim StateText As String
                Dim StateCaret As Integer
            End Structure

            ''' <summary> the count of Redo operations available </summary>
            Private Capcity As Byte = 10

            ''' <summary> the variable that holds the States informations</summary>
            Private UndoList As New Generic.LinkedList(Of UndoState)

            ''' <summary>
            ''' holds the current position of the Redo-Undo operations
            ''' when a new state is added to the end of Redo-Undo list it turns
            ''' to pre-last state (the one that is before the last one) 
            ''' </summary>
            Private _currentIndex As Byte


            ''' <summary>
            ''' adds a new Redo-Undo state at the end of the Redo-Undo list 
            ''' it removes the first one if the list becoms longer than the capcity
            ''' </summary>
            ''' <param name="StateText">the text of the added state</param>
            ''' <param name="StateCaret">the caret positions of the added state</param>
            Public Sub Add(ByVal StateText As String, ByVal StateCaret As Integer)

                UndoList.AddLast(New UndoState With {.StateText = StateText, _
                                                     .StateCaret = StateCaret})


                _currentIndex = UndoList.Count - 1

                If UndoList.Count > Capcity Then
                    UndoList.RemoveFirst()
                    _currentIndex -= 1
                End If

            End Sub

            ''' <summary>gets the previous State of the text to perform an Undo operation </summary>
            Public Function GetPrev() As UndoState

                Dim rtn As New UndoState
                If Me.CanUndo() Then
                    _currentIndex -= 1

                    Dim tmp = UndoList.Skip(_currentIndex).Take(1)
                    rtn = tmp(0)

                End If

                Return rtn
            End Function

            ''' <summary>gets a value indicating whether the tool can perform an Undo operation</summary>
            Public Function CanUndo() As Boolean
                Return CBool(_currentIndex > 0)
            End Function

            ''' <summary>gets the Next State of the text to perform a Redo operation </summary>
            Public Function GetNext() As UndoState

                Dim rtn As New UndoState
                If Me.CanRedo() Then
                    _currentIndex += 1

                    Dim tmp = UndoList.Skip(_currentIndex).Take(1)
                    rtn = tmp(0)

                End If

                Return rtn
            End Function

            ''' <summary>gets a value indicating whether the tool can perform a Redo operation</summary>
            Public Function CanRedo() As Boolean
                Return CBool(_currentIndex < Capcity And _
                             _currentIndex < UndoList.Count - 1 _
                             )
            End Function

            ''' <summary>clears the history of the Undo-Redo operations</summary>
            Public Sub ClearList()
                UndoList.Clear()
            End Sub

        End Class
#End Region

        ''' <summary>the list of the Redo-Undo available States to perform the Redo-Undo Operations</summary>
        Private UndoList As UndoObject


        ''' <summary>clears the history of the Undo-Redo operations</summary>
        Public Sub ClearUndoRedo()
            UndoList.ClearList()
        End Sub

        ''' <summary>
        ''' clears the text on the tool and clears the history of the Undo-Redo operations
        ''' </summary>
        Shadows Sub Clear()
            ClearUndoRedo()
            Me.Text = "" ' we cant use Me.Clear() .... 
            '               cuz it won't clear the text 
            '               it will not use the original Clear method....
            '               it will call itself recursively
        End Sub

        '''<summary>gets a value indicating whether the tool can perform an Undo operation</summary>
        Public Shadows Function CanUndo() As Boolean
            Return UndoList.CanUndo
        End Function

        ''' <summary>performs an Undo operation</summary>
        Shadows Sub Undo()

            If UndoList.CanUndo Then
                Dim State = UndoList.GetPrev
                DisableDrawing()
                NeedColorAll = True
                ToolInternalChanging = True
                Me.Text = State.StateText
                Me.SelectionStart = State.StateCaret
                ToolInternalChanging = False

                EnableDrawing()
            End If
        End Sub

        ''' <summary>gets a value indicating whether the tool can perform a Redo operation</summary>
        Public Shadows Function CanRedo() As Boolean
            Return UndoList.CanRedo
        End Function

        '''<summary>performs a Redo operation</summary>
        Shadows Sub Redo()
            If UndoList.CanRedo Then
                Dim State = UndoList.GetNext
                DisableDrawing()
                NeedColorAll = True
                ToolInternalChanging = True

                Me.Text = State.StateText
                Me.SelectionStart = State.StateCaret

                ToolInternalChanging = False
                EnableDrawing()

            End If
        End Sub

#End Region

#Region "Added Controls and their Events"
        ''' <summary>the ListBox that shows up when user asks for AutoComplete</summary>
        Private WithEvents AutoCompleteListBox As ListBox

        Public Sub New()
            MyBase.New()
            ColoredWords.Initiate("DevRichTextBox- Sub New")

            AutoCompleteListBox = New ListBox With {.Visible = False, _
                                                    .Sorted = True, _
                                                    .Cursor = Cursors.Arrow, _
                                                    .BorderStyle = Windows.Forms.BorderStyle.Fixed3D, _
                                                    .BackColor = Color.Aqua}

            Me.Controls.Add(AutoCompleteListBox)
            'next we're gonna re-set the Font Property of the contained ListBox
            'cuz when it get a child of the RichTextBox it will inherit its Font...
            'and that is not what we want ...
            AutoCompleteListBox.Font = New Font("Tahoma", 8, FontStyle.Regular)

            Me.UndoList = New UndoObject
        End Sub

        Private Sub AutoCompleteListBox_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles AutoCompleteListBox.KeyDown
            Select Case e.KeyCode
                Case Keys.Return, Keys.Space
                    InsertAutoCompleteWord()
                    'RecolorWord("AutoCompleteListBox_KeyDown - Case Keys.Return, Keys.Space")

                Case Keys.Escape
                    AutoCompleteListBox.Hide()
                    Focus()
            End Select
        End Sub

        Private Sub AutoCompleteListBox_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles AutoCompleteListBox.MouseDoubleClick
            InsertAutoCompleteWord()

            'RecolorWord("AutoCompleteListBox_MouseDoubleClick")
        End Sub

        Private Sub AutoCompleteListBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles AutoCompleteListBox.LostFocus
            AutoCompleteListBox.Hide()
        End Sub

        Sub InsertAutoCompleteWord()

            If AutoCompleteListBox.SelectedIndex >= 0 Then
                DisableDrawing()
                If SelectionLength = 0 Then 'oly if there is nothing selected
                    '                    cuz if there is somethng selected ...
                    '                    it's the thing will be replaced by the new inserted word

                    Dim _PrevSpace = PrevSpace
                    _PrevSpace = If(_PrevSpace > 0, _PrevSpace + 1, _PrevSpace) ' in case it's the first word in the Text
                    Me.Select(_PrevSpace, PrevWord.Length)

                End If

                Me.SelectedText = AutoCompleteListBox.SelectedItem

                AutoCompleteListBox.Hide()
                Focus()
                EnableDrawing()

            End If
        End Sub

        Sub SetAutoCompletePoint(Optional ByVal SohwIt As Boolean = False)

            Dim NewPos As New Point()
            With AutoCompleteListBox
                .Height = .ItemHeight * If(.Items.Count > 10, _
                                           10, _
                                           .Items.Count) + 5


                'agin with Height....
                'if the text area is too small ...
                If .Height + 10 > Me.Height Then
                    .Height = Me.Height - 10
                End If



                Dim sizeTmp As SizeF = CreateGraphics.MeasureString("A", Font) 'Qaht's the size of a Letter with this font
                Dim CaretPoint As Point = GetPositionFromCharIndex(SelectionStart)

                'now what if the 'NewPos' is at the EDGE ...!!!!???
                'we don't wanted to get out of sight

                NewPos.X = If(CaretPoint.X + .Width + 10 > Width, _
                              Width - .Width, CaretPoint.X)

                NewPos.Y = If(CaretPoint.Y + .Height + 10 + sizeTmp.Height > Height, _
                              Height - .Height - 10, CaretPoint.Y + sizeTmp.Height)

                '--------------------

                .Location = NewPos
                .BringToFront()

                If SohwIt Then
                    .Show()
                    .Focus()
                    .SelectedIndex = 0 ' it has items so no problem
                End If
            End With
        End Sub

#End Region

#Region "API Declarations"

        ''' <summary>
        ''' this will disables drawing in some control (whoes Handle is passed as a param)
        ''' so working with it will be faster (while no need to see the drawing things)
        ''' </summary>
        ''' <param name="hWnd">the handle of the control to disables drawing in</param>
        Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Integer) As Integer

#End Region

#Region "Added Events"
        ''' <summary>
        ''' occurs when the Insert Keyboard button is pressed telling to change the INS\OVR state is changed
        ''' </summary>
        ''' <param name="sender">the caller object</param>
        ''' <param name="OverTypeState">the new state of INS\OVR state</param>
        Public Event OverType_Changed(ByVal sender As Object, ByVal OverTypeState As Boolean)

#End Region

#Region "Added Properties"

        Private _FileName As String
        ''' <summary>full name of the file that the current text in the control belongs to </summary>
        <Category("Query"), _
        Description("Full name of the file that the current text in the control belongs to") _
        > _
        Public Property FileName() As String
            Get
                Return _FileName
            End Get
            Set(ByVal value As String)
                If _FileName <> value Then
                    _FileName = value
                End If
            End Set
        End Property

        '''<summary>Retrieves the Index of the last character in current Line</summary>
        <Browsable(False)> _
        Public ReadOnly Property GetLastCharIndexOfCurrentLine() As Integer
            Get
                Dim pos = Regex.Match(Text.Substring(SelectionStart), _
                                      "\f|\r|$", _
                                      RegexOptions.None).Index

                Return If(pos = 0, Text.Length, pos + SelectionStart)

            End Get

        End Property


        Private _CurrColNumber As Integer
        '''<summary>the current horizontal location of the Caret in the current Line</summary>
        <Browsable(False)> _
        Public ReadOnly Property CurrColNumber() As Integer
            Get
                Return _CurrColNumber
            End Get
        End Property

        Private _CurrLineNumber As Integer
        ''' <summary>the current Vertical location of the Caret</summary>
        <Browsable(False)> _
         Public ReadOnly Property CurrLineNumber() As Integer
            Get
                Return _CurrLineNumber
            End Get
        End Property

        ''' <summary>
        '''returns the last presence of any of non-word characters (; , ' \ ] [......etc) starting from the current position of the caret
        ''' </summary>
        ''' <param name="Starting">
        ''' the place to start searching , 
        ''' if it's dropped or has a value of -1 then searching will start from the Caret position
        ''' </param>
        <Browsable(False)> _
        Public ReadOnly Property PrevSpace(Optional ByVal Starting As Integer = -1) As Integer

            Get
                Starting = If(Starting = -1, SelectionStart, Starting)
                
                'Return Regex.Match(Text.Substring(0, Starting), _
                '                   "\W", _
                '                    RegexOptions.RightToLeft).Index


                Dim pos As Integer = Starting

                Dim _Matches = Regex.Matches(Text.Substring(0, Starting), _
                                          "\W", _
                                          RegexOptions.RightToLeft)

                If _Matches.Count > 0 Then
                    pos = _Matches.Item(0).Index
                Else
                    pos = 0 ' the start
                End If

                Return pos
            End Get

        End Property

        ''' <summary>
        '''returns the Next presence of any of non-word characters (; , ' \ ] [......etc) starting from the current position of the caret
        ''' </summary>
        <Browsable(False)> _
        Public ReadOnly Property NextSpace(Optional ByVal Starting As Integer = -1) As Integer
            Get
                Starting = If(Starting = -1, SelectionStart, Starting)
                Dim pos As Integer = Starting

                Dim _Matches = Regex.Matches(Text.Substring(Starting), _
                                          "\W", _
                                          RegexOptions.None)
                If _Matches.Count > 0 Then
                    pos = Starting + _Matches.Item(0).Index
                Else
                    pos = Text.Length   ' the end
                End If

                Return pos
            End Get
        End Property

        ''' <summary>
        '''returns the previous word in the text starting from the current position of the caret,
        '''the word is returned without the separator (which separates it from the other words)
        ''' </summary>
        <Browsable(False)> _
        Public ReadOnly Property PrevWord() As String
            Get
                Dim _PrevSpace = PrevSpace ' a local var to hold the value so we don't ask the 
                '                           Property 'PrevSpace' to calc it more thatn one time
                'next is to delete the "\W" before the word
                _PrevSpace = If(_PrevSpace > 0, _PrevSpace, _PrevSpace) ' in case it's the first word in the Text
                Return Text.Substring(_PrevSpace, SelectionStart - _PrevSpace).Trim
            End Get
        End Property

        ''' <summary>
        ''' The text that the property TextModified is set to True 
        ''' if the text in the tool is defferent from
        ''' </summary>
        Private _BaseText As String = ""
        'Private _TextModified As Boolean = False

        ''' <summary>
        ''' the already defined Property "Modified" changes to True when the text format is changed 
        ''' (size , font , color ...) ,... 
        ''' and being we want to save it as a plain text at the end , 
        ''' we can't count on the already defined Property "Modified" , 
        ''' so we had to define our own Property that turns to True only when 
        ''' the plain text is changed
        ''' </summary>
        <Browsable(False)> _
        Public ReadOnly Property TextModified() As Boolean
            Get
                Return _BaseText <> Text
            End Get
        End Property

        Private _OverType As Boolean = False
        ''' <summary>a value indicating the current state of the INS/OVR state of the this tool</summary>
        <Browsable(False)> _
        Public ReadOnly Property OverType() As Boolean
            Get
                Return _OverType
            End Get
        End Property

        <Browsable(False)> _
        Public ReadOnly Property Query() As SqlQuery
            Get
                Dim TrimedSql = If(SelectionLength > 0, _
                                   SelectedText, _
                                   Text).Trim()

                Return New SqlQuery(TrimedSql)
            End Get

        End Property
#End Region

#Region "Added Subs & Functions"


        ''' <summary>
        ''' this is the sub that is responsible for coloring words
        ''' </summary>
        ''' <param name="sender">a value to indicate the caller of this function </param>
        Public Sub ColorCurrentWord(ByVal sender As Object, _
                               Optional ByVal ColorAll As Boolean = False)


            If Me.DesignMode OrElse StartedColoring Then
                Exit Sub
            End If

            StartedColoring = True
            Dim SearchOption As RegexOptions = RegexOptions.IgnorePatternWhitespace Or _
                                               RegexOptions.IgnoreCase

            DisableDrawing()
            If Me.Text.Trim.Length = 0 Then GoTo Exiting


            Dim PrevSelStart = SelectionStart 'to put the caret the palce it was before start coloring
            Dim _PrevSpace, _NextSpace As Integer

            Dim Tet2Color As String = ""
            If ColorAll Then
                _PrevSpace = 0
                _NextSpace = Me.Text.Length '- 1
                Tet2Color = Me.Text
            Else
                'here we're gonna color the Previous Two word and the next Two words

                _PrevSpace = Me.PrevSpace 'get the First prevSpace
                _PrevSpace = Me.PrevSpace(_PrevSpace) ' get the Second prevSpace

                _NextSpace = Me.NextSpace 'get the First NextSpace
                _NextSpace = Me.NextSpace(_NextSpace) 'get the Secode NextSpace
                Tet2Color = Me.Text.Substring(_PrevSpace, _
                                              _NextSpace - _PrevSpace)

            End If



            'Dim Tet2Color As String = Me.Text.Substring(_PrevSpace, _
            '                                            _NextSpace - _PrevSpace)

            If Tet2Color.Length = 0 Then
                GoTo Exiting
            End If


            Me.ColorSection(_PrevSpace, Tet2Color.Length, Color.Black)
            Dim t = Now
            For Each clrdWrd In ColoredWords.GetColoredWords

                Try
                    For Each mtch As Match In Regex.Matches(Tet2Color, _
                                                           clrdWrd.SearchPattern, _
                                                           SearchOption)

                        Me.ColorSection(mtch.Index + _PrevSpace, _
                                        mtch.Length, clrdWrd.color _
                                        )
                    Next
                Catch ex As Exception
                    'Application.DoEvents()
                End Try

            Next

            Dim reg As Regex
            For Each clrdWrd In ColoredWords.OverAllColoredWords()
                reg = New Regex(clrdWrd.SearchPattern, SearchOption)

                For Each m As Match In reg.Matches(Me.Text)
                    Me.ColorSection(m.Index, m.Length, clrdWrd.color)
                Next
            Next



            SelectionLength = 0
            SelectionStart = PrevSelStart
            SelectionColor = Color.Black


Exiting:
            EnableDrawing()
            StartedColoring = False



        End Sub

        ''' <summary> colors all words in the tool </summary>
        ''' <param name="Sender">the caller that calls this sub</param>
        Public Sub RecolorWords(ByVal Sender As String)
            ColorCurrentWord(Sender, True)
            NeedColorAll = False
        End Sub

        ''' <summary>
        ''' Disables Drawing processes , so dealing with the tool internally becomes faster 
        ''' when drwaings are not really a big deal
        ''' </summary>
        Public Sub DisableDrawing()
            LockWindowUpdate(Me.Handle.ToInt32)
        End Sub

        ''' <summary>re-Enables Drawing processes after calling DisableDrawing</summary>
        Public Sub EnableDrawing()
            LockWindowUpdate(0)
        End Sub

        ''' <summary>
        ''' this function will color some section of the text in the control
        ''' with a specified color
        ''' </summary>
        ''' <param name="ColoringStart">location to start coloring from</param>
        ''' <param name="ColoringLen">location to end coloring with</param>
        ''' <param name="UseDisableDrawing">coloring use internally selection ...., 
        ''' so if you wana hide selection while coloring set this param to True othrwise to False,
        ''' But,if you're gonna call this sub several times in one sub rutine 
        ''' it's better for performance to call the DisableDrawing() before and EnableDrawing() yourself 
        ''' and set this param to False each time you call it.
        ''' </param>
        ''' <remarks></remarks>
        Public Sub ColorSection(ByVal ColoringStart As Integer, _
                                ByVal ColoringLen As Integer, _
                                ByVal Color As Drawing.Color, _
                                Optional ByVal UseDisableDrawing As Boolean = False)

            If UseDisableDrawing Then DisableDrawing()

            Me.Select(ColoringStart, ColoringLen)
            Me.SelectionColor = Color

            If UseDisableDrawing Then EnableDrawing()

        End Sub

#Region "Save & SaveAs & Open"

        ''' <summary>this will ask for a new FileName and then Call the Save sub rutine</summary>
        Public Sub SaveAs()
            Using SaveDlg As New SaveFileDialog
                With SaveDlg
                    .Filter = "Text Files (*.txt)|*.txt|Sql Files (*.sql)|*.sql"
                    .Title = "Select a location to save the file to"

                    If .ShowDialog = Windows.Forms.DialogResult.OK Then
                        _FileName = .FileName

                        Save() 'this actually does the save
                    End If

                End With
            End Using
        End Sub

        ''' <summary>
        ''' this will save the text to the pre specified FileName 
        ''' or call SaveAs sub rutine if it's not specified yet
        ''' </summary>
        Public Sub Save()
            If _FileName = "" Then
                SaveAs() ' this will call me agian when it gets the FileName
            Else
                Try
                    My.Computer.FileSystem.WriteAllText(FileName, Text, False)
                    Modified = False
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Caution")
                End Try

                _BaseText = Me.Text 'to set the Property TextModified to Fale

            End If
        End Sub

        ''' <summary>
        ''' asks user to supply a file (text file) path by showing an OpenFileDialog ,
        ''' and shows it's contents
        ''' </summary>
        Public Sub OpenNewFile()

            If TextModified Then
                Dim QuestAnswer As MsgBoxResult
                QuestAnswer = Ask4Save()

                Select Case QuestAnswer
                    Case MsgBoxResult.Yes
                        Save()

                    Case MsgBoxResult.No
                        ' nothing to do ....... 

                    Case MsgBoxResult.Cancel
                        Exit Sub

                End Select
            End If

            Using opnFile As New OpenFileDialog
                With opnFile
                    .Multiselect = False
                    .Filter = "Text Files (*.txt)|*.txt|Sql Files (*.sql)|*.sql"
                    If opnFile.ShowDialog = Windows.Forms.DialogResult.OK Then
                        Clear()
                        ClearUndo()
                        Text = My.Computer.FileSystem.ReadAllText(.FileName)
                        _BaseText = Text
                        FileName = .FileName
                    End If
                End With
            End Using
        End Sub

        ''' <summary> shows a message dialog to ask user if he/she wants to save the current text </summary>
        ''' <returns>returns what user selected a Yes or a No</returns> 
        Public Function Ask4Save(Optional ByVal MsgStyle As MsgBoxStyle = MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNoCancel) _
                                 As MsgBoxResult

            Dim QuestAnswer As MsgBoxResult
            QuestAnswer = MsgBox("Text is changed............." & Environment.NewLine & _
                                 "Do you want to save it ?", _
                                 MsgStyle, _
                                 "Cation")
            Return QuestAnswer
        End Function

#End Region

#End Region

#Region "Overrided & Shadowed Subs & Functions"

        Protected Overrides Sub OnFontChanged(ByVal e As System.EventArgs)
            MyBase.OnFontChanged(e)

            ColorCurrentWord("OnFontchanged")
        End Sub

        Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)


            'If _TextModified AndAlso Not Me.ToolInternalChanging Then
            If Not Me.ToolInternalChanging And Not StartedColoring Then
                UndoList.Add(Text, SelectionStart)
            End If


            MyBase.OnTextChanged(e)

            If NeedColorAll Then
                RecolorWords("OnTextChanged- NeedColorAll = True")
            Else
                If Not StartedColoring Then
                    ColorCurrentWord("OnTextChanged")
                End If
            End If


        End Sub

        Protected Overrides Sub OnSelectionChanged(ByVal e As System.EventArgs)

            _CurrColNumber = SelectionStart - GetFirstCharIndexOfCurrentLine()
            _CurrLineNumber = GetLineFromCharIndex(SelectionStart)
            MyBase.OnSelectionChanged(e)

        End Sub

        Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
            If e.KeyCode = Keys.Insert Then
                _OverType = Not _OverType
                RaiseEvent OverType_Changed(Me, _OverType)
            End If

            MyBase.OnKeyUp(e)

        End Sub

        Protected Overrides Sub OnKeyDown(ByVal e As KeyEventArgs)

            If e.KeyCode = Keys.Space And e.Control Then
                'next is the Scenario ..... :
                'it tries to bring the AutoComplete words (every things : Sql word or DB words)
                'that start with the word whose AutoComplete we're looking for 
                '
                'if it finds no match it brings all things could be written (no filter)
                'if it finds one match it inserts in the RichText (no need to show the Auto Complete List)
                'if it finds more than one matches it shows the Auto Complete List so user chooses what he/she wants


                e.SuppressKeyPress = True
                e.Handled = True ' don't let him to write the space

                Dim Word = PrevWord
                Dim AutoCompleteList = ColoredWords.GetAutoCompleteList(Word). _
                            Union(DB.GetAutoComplete(Word)).Union(Query.GetParamsAutoCompleteList(Word)) ' bring only things that start with Previous Word

                AutoCompleteListBox.Items.Clear()

                If AutoCompleteList.Count > 0 Then 'if it has AutoComplete words that start this word fill it
                    AutoCompleteListBox.Items.AddRange(AutoCompleteList.ToArray)
                Else 'it no match is found for AutoComplete , so show all Ojbects (Sql,Fun , Opr ,.....etc)

                    AutoCompleteList = ColoredWords.GetAutoCompleteList(""). _
                                       Union(DB.GetAutoComplete("")). _
                                       Union(Query.GetParamsAutoCompleteList("")) ' bring all things you have (they all start with "")

                    AutoCompleteListBox.Items.AddRange(AutoCompleteList.ToArray)

                End If


                'Now what to do after you filled the AutoCompleteList ?

                Select Case AutoCompleteList.Count
                    Case Is = 1 'if it's only one Word you don't need to show the list , 
                        '       just put it in the context in the RichText

                        AutoCompleteListBox.SelectedIndex = 0 'cuz the 'InsertAutoCompleteWord()' sub uses the 'SelectedItem'
                        InsertAutoCompleteWord()

                        'RecolorWord("ToolKeyDown - KeyCode = Space + Ctrl")
                    Case Is > 1 'but if it's more than one Word you show the list , 
                        '       so the user will choose what he/she wants

                        SetAutoCompletePoint(True)
                End Select


            End If


            MyBase.OnKeyDown(e)
        End Sub

        ''' <summary>
        ''' Pastes the contents of the Clipboard into the tool
        ''' </summary>
        Shadows Sub Paste()
            Try

                If Clipboard.ContainsText Then
                    'we're not gonna use Me.Paste()
                    'cuz it's gonna past it as it is 
                    '(it could has some Formating , 
                    '  user is pasting from 'MS Word' or somethig)
                    'so we'll deal with it our way
                    SelectedText = Clipboard.GetText(TextDataFormat.UnicodeText) 'this will ReColor all the text
                    'RecolorWords("Shadows Sub Paste")
                End If
            Catch ex As Exception

            End Try
        End Sub

        Protected Overrides Sub OnDragDrop(ByVal drgevent As System.Windows.Forms.DragEventArgs)

            MyBase.OnDragDrop(drgevent)
            NeedColorAll = True

        End Sub

        ''' <summary>
        ''' Moves the current selection in the Tool to the Clipboard 
        ''' </summary>
        Shadows Sub Cut()
            Try
                Clipboard.SetText(Me.SelectedText, TextDataFormat.Rtf)
                Me.SelectedText = ""
                RecolorWords("Shadows Sub Paste")

            Catch ex As Exception

            End Try
        End Sub

        Public Overrides Property SelectedText() As String
            Get
                Return MyBase.SelectedText
            End Get
            Set(ByVal value As String)
                NeedColorAll = True
                MyBase.SelectedText = value

            End Set
        End Property
#End Region

    End Class

    Public Class ColoredWord

        Public color As System.Drawing.Color
        Public SearchPattern As String
        Public Sub New()
            MyBase.New()

            color = Drawing.Color.Black
            SearchPattern = ""
        End Sub
    End Class

    Public Class ColoredWords

#Region "Fields"
        Private Shared _Initiated As Boolean
        Private Shared _Pattern As String = "(?<=^|\W)( KEYWORD_PLACEHOLDER ) (?=\W|$)"

        Private Shared _ColoredWords As IEnumerable(Of ColoredWord)
        Private Shared _AutoCopleteLst As IEnumerable(Of String)

        Private Shared _SqlKeyWords() As String = New String() _
                                    {"Parameters", "Transform", "Select", "All", "Distinct", "DistinctRow", "Top", "Percent", "As", _
                                     "Insert", "Delete", "Update", "Into", "Set", _
                                     "From", _
                                     "Values", "On", _
                                     "Where", "Having", _
                                     "Order", "Group", "By", "Pivot", _
                                     "Create", "Drop", "Add", "Alter", _
                                     "Table", "Column", "Proc", "Exec", _
                                     "AutoIncreament", _
                                     "Text", "Memo", "Number", "YesNo", "OLEObject", _
                                     "Short", "Long", "Single", "Double", "Currency", "DateTime", "Bit", "Byte", "GUID", "BigBinary", "LongBinary", "VarBinary", "LongText", "VarChar", "Decimal"}


        Private Shared _FunctionNames() As String = New String() {"LBound", "UBound", _
                                       "Asc", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "Chr", "Chr$", "CInt", "Clng", "CSng", "CStr", "CVar", "CVDate", "DateSerial", "DateValue", "Day", "FormatCurrency", "FormatDateTime", "FormatNumber", "FormatPercent", "GUIDFromString", "Hex", " Hex$", "Hour", "Minute", "Month", "Nz", "Oct", "Oct$", "Second", "Str", "Str$", "StrConv", "StringFromGUID", "TimeSerial", "TimeValue", "Val", "Weekday", "Year", _
                                       "CurrentUser", "Eval", "HyperlinkPart", "IMEStatus", "Partition", _
                                       "CDate", "CVDate", "Date", "Date$", "DateAdd", "DateDiff", "DatePart", "DateSerial", "DateValue", "Day", "Hour", "IsDate", "Minute", "Month", "MonthName", "Now", "Second", "Time", "Time$", "Timer", "TimeSerial", "TimeValue", "Weekday", "WeekdayName", "Year", _
                                       "DAvg", "DCount", "DFirst", "DLast", "DLookup", "DMax", "DMin", "DStDev", "DSum", "DVar", "DVarP", _
                                       "CVErr", "Err", "Error", "Error$", "IsError", _
                                       "DDB", "FV", "IPmt", "IRR", "MIRR", "NPer", "NPV", "Pmt", "PPmt", "PV", "Rate", "SLN", "SYD", _
                                       "IsArray", "IsDate", "IsEmpty", "IsError", "IsMissing", "IsNull", "IsNumeric", "IsObject", "TypeName", "VarType", _
                                       "Abs", "Atn", "Cos", "Exp", "Fix", "Int", "Log", "Rnd", "Round", "Sgn", "Sin", "Sqr", "Tan", _
                                       "Choose", "IIf", "Switch", _
                                       "Avg", "Count", "Max", "Min", "StDev", "StDevP", "Sum", "Var", "VarP", _
                                       "Asc", "Chr", "Chr$", "Format", "Format$", "GUIDFromString", "InStr", "InStrRev", "LCase", "LCase$", "Left", "Left$", "Len", "LTrim", "LTrim$", "Mid", "Mid$", "Replace", "Right", "Node73", "Right$", "RTrim", "RTrim$", "Space", "Space$", "StrComp", "StrConv", "String", "String$", "StringFromGUID", "StrReverse", "Trim", "Trim$", "UCase", "UCase$" _
                                       }

        Private Shared _SqlLiteralOperators() As String = New String() _
                {"Inner", "Outer", "Join", _
                 "And", "Or", "Xor", "Not", "Between", "Like", "In", "Eqv", "Imp", _
                 "Null", "Exists", "Mod"}

        Private Shared _SqlOperators() As String = New String() _
                               {"*", "^", "\", "/", "+", "-", ">", "<", "=", "&"}

        Private Shared _SqlPunctuation() As String = New String() {";", ",", ".", "(", ")"}

        Private Shared _Comments() As String = New String() {"(--.*(\n|\f|$)) | (/ \*  [\W+\w]{0,}?  \* /)"} 'notice the Lazy match for /* Comments */
        Private Shared _Strings() As String = New String() {"'[\W+\w]{0,}?'|""[\W+\w]{0,}?"""} ' notice the "Lazy Matching" for ' ' and for " "


#End Region

#Region "Properties"

        Public Shared ReadOnly Property GetColoredWords() As IEnumerable(Of ColoredWord)
            Get
                Return _ColoredWords
            End Get
        End Property

        Public Shared ReadOnly Property GetAutoCompleteList(ByVal word As String) As IEnumerable(Of String)
            Get
                Return From Item In _AutoCopleteLst _
                       Where Item.StartsWith(word, StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        Public Shared ReadOnly Property Initiated() As Boolean
            Get
                Return _Initiated
            End Get
        End Property

#Region "Colors"

        Private Shared _KeyWordsColor As Drawing.Color = Drawing.Color.Blue
        Public Shared Property KeyWordsColor() As Drawing.Color
            Get
                Return _KeyWordsColor
            End Get
            Set(ByVal value As Drawing.Color)
                If _KeyWordsColor <> value Then
                    _KeyWordsColor = value
                End If
            End Set
        End Property

        Private Shared _FunctionsColor As Drawing.Color = Drawing.Color.DeepPink
        Public Shared Property FunctionsColor() As Drawing.Color
            Get
                Return _FunctionsColor
            End Get
            Set(ByVal value As Drawing.Color)
                If _FunctionsColor <> value Then
                    _FunctionsColor = value
                End If
            End Set
        End Property

        Private Shared _SqlLiteralOperatorsColor As Drawing.Color = Drawing.Color.DarkGray
        Public Shared Property SqlLiteralOperatorsColor() As Drawing.Color
            Get
                Return _SqlLiteralOperatorsColor
            End Get
            Set(ByVal value As Drawing.Color)
                If _SqlLiteralOperatorsColor <> value Then
                    _SqlLiteralOperatorsColor = value
                End If
            End Set
        End Property

        Private Shared _SqlOperatorsColor As Drawing.Color = Drawing.Color.DarkGray
        Public Shared Property SqlOperatorsColor() As Drawing.Color
            Get
                Return _SqlOperatorsColor
            End Get
            Set(ByVal value As Drawing.Color)
                If _SqlOperatorsColor <> value Then
                    _SqlOperatorsColor = value
                End If
            End Set
        End Property

        Private Shared _SqlPunctuationColor As Drawing.Color = Drawing.Color.DarkGray
        Public Shared Property SqlPunctuationColor() As Drawing.Color
            Get
                Return _SqlPunctuationColor
            End Get
            Set(ByVal value As Drawing.Color)
                If _SqlPunctuationColor <> value Then
                    _SqlPunctuationColor = value
                End If
            End Set
        End Property

        Private Shared _CommentsColor As Drawing.Color = Drawing.Color.Green
        Public Shared Property CommentsColor() As Drawing.Color
            Get
                Return _CommentsColor
            End Get
            Set(ByVal value As Drawing.Color)
                If _CommentsColor <> value Then
                    _CommentsColor = value
                End If
            End Set
        End Property

        Private Shared _StringsColor As Drawing.Color = Drawing.Color.Red
        Public Shared Property StringsColor() As Drawing.Color
            Get
                Return _StringsColor
            End Get
            Set(ByVal value As Drawing.Color)
                If _StringsColor <> value Then
                    _StringsColor = value
                End If
            End Set
        End Property

#End Region

#End Region

#Region "Functions"
        Private Shared Function EscapeTxt4Pattern(ByVal plainTxt As String) As String

            Dim Fun_Replace = Function(m As Match) _
                                  "\" & m.Value
            Dim P As String = "\^| \\|	\[|	\]|	\+| \*|	\.|	\(|	\)|	\$"

            Return Regex.Replace(plainTxt, P, _
                                 Fun_Replace, _
                                 RegexOptions.IgnorePatternWhitespace)

        End Function

        Private Shared Function GetPattern4Words(ByVal Replacement As String) As String
            Return _Pattern.Replace("KEYWORD_PLACEHOLDER", Replacement)
        End Function

        Public Shared Function SqlKeyWords() As IEnumerable(Of ColoredWord)

            Return From KW In _SqlKeyWords _
                   Select New ColoredWord() With _
                            {.color = KeyWordsColor, _
                             .SearchPattern = GetPattern4Words(KW)}

        End Function

        Public Shared Function FunctionNames() As IEnumerable(Of ColoredWord)
            Return From Fun In _FunctionNames _
                   Select New ColoredWord() With _
                            {.color = FunctionsColor, _
                             .SearchPattern = GetPattern4Words(Fun.Replace("$", "$$"))}
        End Function

        Public Shared Function SqlOperators() As IEnumerable(Of ColoredWord)

            Return From Opr In _SqlOperators _
                   Select New ColoredWord() With _
                        {.color = SqlOperatorsColor, _
                         .SearchPattern = EscapeTxt4Pattern(Opr)}

        End Function

        Public Shared Function SqlLiteralOperators() As IEnumerable(Of ColoredWord)

            Return From Opr In _SqlLiteralOperators _
                   Select New ColoredWord() With _
                        {.color = SqlLiteralOperatorsColor, _
                         .SearchPattern = GetPattern4Words(EscapeTxt4Pattern(Opr))}

        End Function

        Public Shared Function SqlPunctuation() As IEnumerable(Of ColoredWord)

            Return From Pun In _SqlPunctuation _
                   Select New ColoredWord() With _
                        {.color = SqlPunctuationColor, _
                         .SearchPattern = EscapeTxt4Pattern(Pun)}

        End Function

        Public Shared Function Comments() As IEnumerable(Of ColoredWord)
            Return From com In _Comments _
                   Select New ColoredWord() With _
                            {.color = CommentsColor, _
                             .SearchPattern = com}


        End Function

        Public Shared Function Strings() As IEnumerable(Of ColoredWord)
            Return From s In _Strings _
                       Select New ColoredWord() With _
                       {.color = StringsColor, _
                        .SearchPattern = s}

        End Function

        Public Shared Function NonSqlChars() As IEnumerable(Of ColoredWord)

            Dim _NonSqlChars() As String = New String() {"\[|\]|'|:", "\[ [\W+\w]{0,}? \]"} 'notice the "Lazy Matching" for [ ]

            Return From NSC In _NonSqlChars _
                       Select New ColoredWord() With _
                       {.color = Drawing.Color.Black, _
                       .SearchPattern = NSC}


        End Function

        Public Shared Function GetCommentsPattern() As String
            Return String.Join(" ", _Comments)
        End Function

        Public Shared Function OverAllColoredWords() As IEnumerable(Of ColoredWord)
            Return NonSqlChars.Union(Strings).Union(Comments)

            ' this needs moooooooore attention .......
            'look at this line , 
            'and gues what colors would be ther:    
            '                   [ select ' from xT ] where ' what ever

        End Function

#End Region

#Region "Sub Routines"
        Public Shared Sub Initiate(ByVal sender As Object)

            _ColoredWords = SqlKeyWords(). _
                    Union(FunctionNames()). _
                    Union(SqlLiteralOperators()). _
                    Union(SqlPunctuation()). _
                    Union(SqlOperators())  '. _
            'Union(NonSqlChars())



            _AutoCopleteLst = _SqlKeyWords.Union(_FunctionNames). _
                                Union(_SqlLiteralOperators)

            _Initiated = True
        End Sub

        Public Shared Sub ResetColors()
            SqlLiteralOperatorsColor = Drawing.Color.DarkGray
            FunctionsColor = Drawing.Color.DeepPink
            SqlOperatorsColor = Drawing.Color.DarkGray
            KeyWordsColor = Drawing.Color.Blue
            StringsColor = Drawing.Color.Red
            CommentsColor = Drawing.Color.Green
        End Sub
#End Region

    End Class

End Namespace