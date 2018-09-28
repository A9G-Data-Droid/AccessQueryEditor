#Region "Imports"
Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports QueryEditor.Queries
#End Region

Namespace ColoringWords
    <ToolboxBitmap(GetType(RichTextBox))>
    Public Class DevRichTextBox
        Inherits RichTextBox

#Region "Fields"

        ''' <summary>
        '''     in some certain stances we need to change (Internally) the Text property of the tool,
        '''     but changing the text will affect other actions
        '''     (like Undo and Redo ability) that are not Internal
        '''     so this variable will determine whether the current "Text-Changing Operation" is Internal or not
        ''' </summary>
        Private _toolInternalChanging As Boolean = False

        ''' <summary>
        '''     in some certain stances we need to Recolor all the text in the tool,
        '''     so this is the variable that determine whether to color the CurrentWord or all text
        ''' </summary>
        Private _needColorAll As Boolean = False

        ''' <summary>
        '''     RecolorWord Sub will change the Selection many times ....,
        '''     but Changing Selection will Raise OnTextChanged Event which in turn Calls RecolorWord Sub ,
        '''     even Changing the Font will Raise OnTextChanged Event ,......
        '''     thus we'll have a Recursive Calls ,....
        '''     so we use this var to tell the RecolorWord Sub not to Start Coloring again cuz it's already coloring
        ''' </summary>
        Private _startedColoring As Boolean = False

#End Region

#Region "Undo-Redo Handling and Operations"

        ''' <summary>
        '''     The original Undo , Redo .....of the tool affect all actions
        '''     even Font Properties and selection Properties,
        '''     so we have to make something instead
        '''     so this is a class to perform the Redo ,Undo
        ''' </summary>
        Private Class UndoObject
            ''' <summary>
            '''     Redo,Undo is stored a Pair of info
            '''     (the text it was, and the Caret position it was)
            '''     and this is a Structure to hold this info for each State
            ''' </summary>
            Structure UndoState
                Dim StateText As String
                Dim StateCaret As Integer
            End Structure

            ''' <summary> the count of Redo operations available </summary>
            Private Const RedoCapacity As Byte = 10

            ''' <summary> the variable that holds the States information</summary>
            Private ReadOnly _undoList As New LinkedList(Of UndoState)

            ''' <summary>
            '''     holds the current position of the Redo-Undo operations
            '''     when a new state is added to the end of Redo-Undo list it turns
            '''     to pre-last state (the one that is before the last one)
            ''' </summary>
            Private _currentIndex As Integer


            ''' <summary>
            '''     adds a new Redo-Undo state at the end of the Redo-Undo list
            '''     it removes the first one if the list becomes longer than the capacity
            ''' </summary>
            ''' <param name="stateText">the text of the added state</param>
            ''' <param name="stateCaret">the caret positions of the added state</param>
            Public Sub Add(stateText As String, stateCaret As Integer)

                _undoList.AddLast(New UndoState With {.StateText = StateText, _
                                    .StateCaret = StateCaret})


                _currentIndex = _undoList.Count - 1

                If _undoList.Count > RedoCapacity Then
                    _undoList.RemoveFirst()
                    _currentIndex -= 1
                End If
            End Sub

            ''' <summary>gets the previous State of the text to perform an Undo operation </summary>
            Public Function GetPrev() As UndoState

                Dim rtn As New UndoState
                If CanUndo() Then
                    _currentIndex -= 1

                    Dim tmp = _undoList.Skip(_currentIndex).Take(1)
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
                If CanRedo() Then
                    _currentIndex += 1

                    Dim tmp = _undoList.Skip(_currentIndex).Take(1)
                    rtn = tmp(0)

                End If

                Return rtn
            End Function

            ''' <summary>gets a value indicating whether the tool can perform a Redo operation</summary>
            Public Function CanRedo() As Boolean
                Return CBool(_currentIndex < RedoCapacity And
                             _currentIndex < _undoList.Count - 1
                    )
            End Function

            ''' <summary>clears the history of the Undo-Redo operations</summary>
            Public Sub ClearList()
                _undoList.Clear()
            End Sub
        End Class

        ''' <summary>the list of the Redo-Undo available States to perform the Redo-Undo Operations</summary>
        Private ReadOnly _undoList As UndoObject

        ''' <summary>clears the history of the Undo-Redo operations</summary>
        Private Sub ClearUndoRedo()
            _undoList.ClearList()
        End Sub

        ''' <summary>
        '''     clears the text on the tool and clears the history of the Undo-Redo operations
        ''' </summary>
        Shadows Sub Clear()
            ClearUndoRedo()
            Text = "" ' we cant use Me.Clear() .... 
            '               cuz it won't clear the text 
            '               it will not use the original Clear method....
            '               it will call itself recursively
        End Sub

        '''<summary>gets a value indicating whether the tool can perform an Undo operation</summary>
        Public Shadows Function CanUndo() As Boolean
            Return _undoList.CanUndo
        End Function

        ''' <summary>performs an Undo operation</summary>
        Shadows Sub Undo()

            If _undoList.CanUndo Then
                Dim state = _undoList.GetPrev
                _needColorAll = True
                _toolInternalChanging = True
                Text = state.StateText
                SelectionStart = state.StateCaret
                _toolInternalChanging = False
            End If
        End Sub

        ''' <summary>gets a value indicating whether the tool can perform a Redo operation</summary>
        Public Shadows Function CanRedo() As Boolean
            Return _undoList.CanRedo
        End Function

        '''<summary>performs a Redo operation</summary>
        Shadows Sub Redo()
            If _undoList.CanRedo Then
                Dim state = _undoList.GetNext
                _needColorAll = True
                _toolInternalChanging = True

                Text = state.StateText
                SelectionStart = state.StateCaret

                _toolInternalChanging = False
            End If
        End Sub

#End Region

#Region "Added Controls and their Events"

        ''' <summary>the ListBox that shows up when user asks for AutoComplete</summary>
        Private WithEvents _autoComplete As ListBox

        Public Sub New()
            MyBase.New()
            ColoredWords.Initiate()

            _autoComplete = New ListBox With {.Visible = False, _
                .Sorted = True, _
                .Cursor = Cursors.Arrow, _
                .BorderStyle = BorderStyle.Fixed3D, _
                .BackColor = Color.Aqua}

            Controls.Add(_autoComplete)
            'next we're gonna re-set the Font Property of the contained ListBox
            'cuz when it get a child of the RichTextBox it will inherit its Font...
            'and that is not what we want ...
            _autoComplete.Font = New Font("Tahoma", 8, FontStyle.Regular)

            _undoList = New UndoObject
        End Sub

        Private Sub AutoCompleteListBox_KeyDown(sender As Object, e As KeyEventArgs) Handles _autoComplete.KeyDown
            Select Case e.KeyCode
                Case Keys.Return, Keys.Space
                    InsertAutoCompleteWord()
                    'RecolorWord("AutoCompleteListBox_KeyDown - Case Keys.Return, Keys.Space")

                Case Keys.Escape
                    _autoComplete.Hide()
                    Focus()
            End Select
        End Sub

        Private Sub AutoCompleteListBox_MouseDoubleClick(sender As Object, e As MouseEventArgs) _
            Handles _autoComplete.MouseDoubleClick
            InsertAutoCompleteWord()

            'RecolorWord("AutoCompleteListBox_MouseDoubleClick")
        End Sub

        Private Sub AutoCompleteListBox_LostFocus(sender As Object, e As EventArgs) _
            Handles _autoComplete.LostFocus
            _autoComplete.Hide()
        End Sub

        Private Sub InsertAutoCompleteWord()
            If _autoComplete.SelectedIndex >= 0 Then
                If SelectionLength = 0 Then 'oly if there is nothing selected
                    '                    cuz if there is something selected ...
                    '                    it's the thing will be replaced by the new inserted word

                    Dim prevSpace = PreviousSpace
                    prevSpace = If(prevSpace > 0, prevSpace + 1, prevSpace) _
                    ' in case it's the first word in the Text
                    Me.Select(prevSpace, PrevWord.Length)

                End If

                SelectedText = _autoComplete.SelectedItem.ToString()

                _autoComplete.Hide()
                Focus()
            End If
        End Sub

        Private Sub SetAutoCompletePoint(Optional ByVal showIt As Boolean = False)

            Dim newPos As New Point()
            With _autoComplete
                .Height = .ItemHeight*If(.Items.Count > 10,
                                         10,
                                         .Items.Count) + 5


                'again with Height....
                'if the text area is too small ...
                If .Height + 10 > Height Then
                    .Height = Height - 10
                End If


                Dim sizeTmp As SizeF = CreateGraphics.MeasureString("A", Font) _
                'What's the size of a Letter with this font
                Dim caretPoint As Point = GetPositionFromCharIndex(SelectionStart)

                'now what if the 'NewPos' is at the EDGE ...!!!!???
                'we don't wanted to get out of sight

                newPos.X = If(caretPoint.X + .Width + 10 > Width,
                              Width - .Width, caretPoint.X)

                newPos.Y = CInt(If(caretPoint.Y + .Height + 10 + sizeTmp.Height > Height,
                              Height - .Height - 10, caretPoint.Y + sizeTmp.Height))

                '--------------------

                .Location = newPos
                .BringToFront()

                If showIt Then
                    .Show()
                    .Focus()
                    .SelectedIndex = 0 ' it has items so no problem
                End If
            End With
        End Sub

#End Region

#Region "API Declarations"

        ''' <summary>
        '''     this will disables drawing in some control (whose Handle is passed as a param)
        '''     so working with it will be faster (while no need to see the drawing things)
        ''' </summary>
        ''' <param name="hWnd">the handle of the control to disables drawing in</param>
        Private Declare Function LockWindowUpdate Lib "user32"(hWnd As Integer) As Integer

#End Region

#Region "Added Events"

        ''' <summary>
        '''     occurs when the Insert Keyboard button is pressed telling to change the INS\OVR state is changed
        ''' </summary>
        ''' <param name="sender">the caller object</param>
        ''' <param name="overTypeState">the new state of INS\OVR state</param>
        Public Event OverTypeChanged(sender As Object, overTypeState As Boolean)

#End Region

#Region "Added Properties"

        Private _fileName As String

        ''' <summary>full name of the file that the current text in the control belongs to </summary>
        <Category("Query"),
            Description("Full name of the file that the current text in the control belongs to")
            >
        Public Property FileName As String
            Get
                Return _FileName
            End Get
            Set
                If _FileName <> value Then
                    _FileName = value
                End If
            End Set
        End Property

        '''<summary>Retrieves the Index of the last character in current Line</summary>
        <Browsable(False)>
        Public ReadOnly Property GetLastCharIndexOfCurrentLine As Integer
            Get
                Dim pos = Regex.Match(Text.Substring(SelectionStart),
                                      "\f|\r|$",
                                      RegexOptions.None).Index

                Return If(pos = 0, Text.Length, pos + SelectionStart)
            End Get
        End Property


        Private _cursorColumnNumber As Integer

        '''<summary>the current horizontal location of the Caret in the current Line</summary>
        <Browsable(False)>
        Public ReadOnly Property CursorColumnNumber As Integer
            Get
                Return _cursorColumnNumber
            End Get
        End Property

        Private _currLineNumber As Integer

        ''' <summary>the current Vertical location of the Caret</summary>
        <Browsable(False)>
        Public ReadOnly Property CurrLineNumber As Integer
            Get
                Return _CurrLineNumber
            End Get
        End Property

        ''' <summary>
        '''     returns the last presence of any of non-word characters (; , ' \ ] [......etc) starting from the current position
        '''     of the caret
        ''' </summary>
        ''' <param name="starting">
        '''     the place to start searching ,
        '''     if it's dropped or has a value of -1 then searching will start from the Caret position
        ''' </param>
        <Browsable(False)>
        Private ReadOnly Property PreviousSpace(Optional ByVal starting As Integer = - 1) As Integer
            Get
                Starting = If(Starting = - 1, SelectionStart, Starting)

                'Return Regex.Match(Text.Substring(0, Starting), _
                '                   "\W", _
                '                    RegexOptions.RightToLeft).Index

                Dim pos As Integer '= Starting

                Dim matches = Regex.Matches(Text.Substring(0, Starting), "\W",
                                             RegexOptions.RightToLeft)

                If matches.Count > 0 Then
                    pos = matches.Item(0).Index
                Else
                    pos = 0 ' the start
                End If

                Return pos
            End Get
        End Property

        ''' <summary>
        '''     returns the Next presence of any of non-word characters (; , ' \ ] [......etc) starting from the current position
        '''     of the caret
        ''' </summary>
        <Browsable(False)>
        Private ReadOnly Property NextSpace(Optional ByVal starting As Integer = - 1) As Integer
            Get
                Starting = If(Starting = - 1, SelectionStart, Starting)
                Dim pos As Integer '= Starting

                Dim matches = Regex.Matches(Text.Substring(Starting),
                                             "\W",
                                             RegexOptions.None)
                If matches.Count > 0 Then
                    pos = Starting + matches.Item(0).Index
                Else
                    pos = Text.Length   ' the end
                End If

                Return pos
            End Get
        End Property

        ''' <summary>
        '''     returns the previous word in the text starting from the current position of the caret,
        '''     the word is returned without the separator (which separates it from the other words)
        ''' </summary>
        <Browsable(False)>
        Private ReadOnly Property PrevWord As String
            Get
                Dim prevSpace = PreviousSpace ' a local var to hold the value so we don't ask the 
                '                           Property 'PrevSpace' to calc it more than one time
                'next is to delete the "\W" before the word

                prevSpace = If(prevSpace > 0, prevSpace, prevSpace + 1) ' in case it's the first word in the Text

                Return Text.Substring(prevSpace, SelectionStart - prevSpace).Trim
            End Get
        End Property

        ''' <summary>
        '''     The text that the property TextModified is set to True
        '''     if the text in the tool is different from
        ''' </summary>
        Private _baseText As String = ""
        'Private _TextModified As Boolean = False

        ''' <summary>
        '''     the already defined Property "Modified" changes to True when the text format is changed
        '''     (size , font , color ...) ,...
        '''     and being we want to save it as a plain text at the end ,
        '''     we can't count on the already defined Property "Modified" ,
        '''     so we had to define our own Property that turns to True only when
        '''     the plain text is changed
        ''' </summary>
        <Browsable(False)>
        Public ReadOnly Property TextModified As Boolean
            Get
                Return _BaseText <> Text
            End Get
        End Property

        Private _overType As Boolean = False

        ''' <summary>a value indicating the current state of the INS/OVR state of the this tool</summary>
        <Browsable(False)>
        Public ReadOnly Property OverType As Boolean
            Get
                Return _OverType
            End Get
        End Property

        <Browsable(False)>
        Public ReadOnly Property Query As SqlQuery
            Get
                Dim trimmedSql = If(SelectionLength > 0,
                                   SelectedText,
                                   Text).Trim()

                Return New SqlQuery(trimmedSql)
            End Get
        End Property

#End Region

#Region "Added Subs & Functions"

        ''' <summary>
        '''     this is the sub that is responsible for coloring words
        ''' </summary>
        Private Sub ColorCurrentWord(Optional ByVal colorAll As Boolean = False)

            If DesignMode OrElse _startedColoring OrElse Text.Trim.Length = 0 Then
                Exit Sub
            End If

            _startedColoring = True
            Const searchOption As RegexOptions = RegexOptions.IgnorePatternWhitespace Or
                                                 RegexOptions.IgnoreCase

            DisableDrawing()

            Dim prevSelStart = SelectionStart 'to put the caret the place it was before start coloring
            Dim prevSpace As Integer
            Dim afterSpace As Integer

            Dim text2Color as String
            If ColorAll Then
                prevSpace = 0
                'afterSpace = Me.Text.Length '- 1
                text2Color = Text
            Else
                'here we're gonna color the Previous Two word and the next Two words

                prevSpace = PreviousSpace 'get the First prevSpace
                prevSpace = PreviousSpace(prevSpace) ' get the Second prevSpace

                afterSpace = NextSpace 'get the First NextSpace
                afterSpace = NextSpace(afterSpace) 'get the Second NextSpace
                text2Color = Text.Substring(prevSpace, afterSpace - prevSpace)
            End If

            'Dim Tet2Color As String = Me.Text.Substring(_PrevSpace, _
            '                                            _NextSpace - _PrevSpace)
            If text2Color.Length <> 0 Then
                ColorSection(prevSpace, text2Color.Length, Color.Black)
                'Dim t = Now
                For Each coloredWord In ColoredWords.GetColoredWords
                    Try
                        For Each wordMatch As Match In Regex.Matches(text2Color,
                                                                coloredWord.SearchPattern,
                                                                searchOption)

                            ColorSection(wordMatch.Index + prevSpace,
                                            wordMatch.Length, coloredWord.color
                                            )
                        Next
                    Catch ex As Exception
                        Debug.WriteLine(ex.Message)
                    End Try
                Next

                Dim reg As Regex
                For Each clrdWrd In ColoredWords.OverAllColoredWords()
                    reg = New Regex(clrdWrd.SearchPattern, searchOption)

                    For Each m As Match In reg.Matches(Text)
                        ColorSection(m.Index, m.Length, clrdWrd.color)
                    Next
                Next

                SelectionLength = 0
                SelectionStart = prevSelStart
                SelectionColor = Color.Black
            End If

            EnableDrawing()
            _startedColoring = False
        End Sub

        ''' <summary> colors all words in the tool </summary>
        Public Sub RecolorWords()
            ColorCurrentWord(True)
            _needColorAll = False
        End Sub

        ''' <summary>
        '''     Disables Drawing processes , so dealing with the tool internally becomes faster
        '''     when drawings are not really a big deal
        ''' </summary>
        Public Sub DisableDrawing()
            LockWindowUpdate(Handle.ToInt32)
        End Sub

        ''' <summary>re-Enables Drawing processes after calling DisableDrawing</summary>
        Public Sub EnableDrawing()
            LockWindowUpdate(0)
        End Sub

        ''' <summary>
        '''     this function will color some section of the text in the control
        '''     with a specified color
        ''' </summary>
        ''' <param name="coloringStart">location to start coloring from</param>
        ''' <param name="coloringLen">location to end coloring with</param>
        ''' <param name="useDisableDrawing">
        '''     coloring use internally selection ....,
        '''     so if you wanna hide selection while coloring set this param to True otherwise to False,
        '''     But,if you're gonna call this sub several times in one sub routine
        '''     it's better for performance to call the DisableDrawing() before and EnableDrawing() yourself
        '''     and set this param to False each time you call it.
        ''' </param>
        ''' <remarks></remarks>
        Private Sub ColorSection(coloringStart As Integer,
                                coloringLen As Integer,
                                textColor As Color,
                                Optional ByVal useDisableDrawing As Boolean = False)

            If UseDisableDrawing Then DisableDrawing()

            [Select](ColoringStart, ColoringLen)
            SelectionColor = textColor

            If UseDisableDrawing Then EnableDrawing()
        End Sub

#End Region

#Region "Save & SaveAs & Open"

        ''' <summary>this will ask for a new FileName and then Call the Save sub routine</summary>
        Public Sub SaveAs()
            Using saveDlg As New SaveFileDialog
                With saveDlg
                    .Filter = "Text Files (*.txt)|*.txt|Sql Files (*.sql)|*.sql"
                    .Title = "Select a location to save the file to"

                    If .ShowDialog = DialogResult.OK Then
                        _FileName = .FileName

                        Save() 'this actually does the save
                    End If
                End With
            End Using
        End Sub

        ''' <summary>
        '''     this will save the text to the pre specified FileName
        '''     or call SaveAs sub routine if it's not specified yet
        ''' </summary>
        Public Sub Save()
            If _FileName = "" Then
                SaveAs() ' this will call me again when it gets the FileName
            Else
                Try
                    My.Computer.FileSystem.WriteAllText(FileName, Text, False)
                    Modified = False
                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Caution")
                End Try

                _BaseText = Text 'to set the Property TextModified to False
            End If
        End Sub

        ''' <summary>
        '''     asks user to supply a file (text file) path by showing an OpenFileDialog ,
        '''     and shows it's contents
        ''' </summary>
        Public Sub OpenNewFile()
            If TextModified Then
                Dim questAnswer As MsgBoxResult
                questAnswer = Ask4Save()

                Select Case questAnswer
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
                    If opnFile.ShowDialog = DialogResult.OK Then
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
        Public Function Ask4Save(
                                 Optional ByVal msgStyle As MsgBoxStyle =
                                    MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNoCancel) _
            As MsgBoxResult

            Dim questAnswer As MsgBoxResult
            questAnswer = MsgBox("Text is changed............." & Environment.NewLine &
                                 "Do you want to save it ?",
                                 msgStyle,
                                 "Cation")
            Return questAnswer
        End Function

#End Region

#Region "Overrided & Shadowed Subs & Functions"

        Protected Overrides Sub OnFontChanged(e As EventArgs)
            MyBase.OnFontChanged(e)

            ColorCurrentWord()
        End Sub

        Protected Overrides Sub OnTextChanged(e As EventArgs)
            'If _TextModified AndAlso Not Me.ToolInternalChanging Then
            If Not _toolInternalChanging And Not _startedColoring Then
                _undoList.Add(Text, SelectionStart)
            End If


            MyBase.OnTextChanged(e)

            If _needColorAll Then
                RecolorWords()
            Else
                If Not _startedColoring Then
                    ColorCurrentWord()
                End If
            End If
        End Sub

        Protected Overrides Sub OnSelectionChanged(e As EventArgs)

            _cursorColumnNumber = SelectionStart - GetFirstCharIndexOfCurrentLine()
            _CurrLineNumber = GetLineFromCharIndex(SelectionStart)
            MyBase.OnSelectionChanged(e)
        End Sub

        Protected Overrides Sub OnKeyUp(e As KeyEventArgs)
            If e.KeyCode = Keys.Insert Then
                _OverType = Not _OverType
                RaiseEvent OverTypeChanged(Me, _OverType)
            End If

            MyBase.OnKeyUp(e)
        End Sub

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)

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

                Dim word = PrevWord
                Dim autoCompleteList = ColoredWords.GetAutoCompleteList(word).
                        Union(Editor.CurrentDb.GetAutoComplete(word)).Union(Query.GetParamsAutoCompleteList(word)) _
                ' bring only things that start with Previous Word
                _autoComplete.Items.Clear()

                If autoCompleteList.Count > 0 Then 'if it has AutoComplete words that start this word fill it
                    _autoComplete.Items.AddRange(autoCompleteList.ToArray)
                Else 'it no match is found for AutoComplete , so show all Objects (Sql,Fun , Opr ,.....etc)

                    autoCompleteList = ColoredWords.GetAutoCompleteList("").
                        Union(Editor.CurrentDb.GetAutoComplete("")).
                        Union(Query.GetParamsAutoCompleteList("")) ' bring all things you have (they all start with "")

                    _autoComplete.Items.AddRange(autoCompleteList.ToArray)

                End If

                'Now what to do after you filled the AutoCompleteList ?
                Select Case autoCompleteList.Count
                    Case Is = 1 'if it's only one Word you don't need to show the list , 
                        '       just put it in the context in the RichText

                        _autoComplete.SelectedIndex = 0 _
                        'cuz the 'InsertAutoCompleteWord()' sub uses the 'SelectedItem'
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
        '''     Pastes the contents of the Clipboard into the tool
        ''' </summary>
        Shadows Sub Paste()
            Try
                If Clipboard.ContainsText Then
                    'we're not gonna use Me.Paste()
                    'cuz it's gonna past it as it is 
                    '(it could has some Formatting , 
                    '  user is pasting from 'MS Word' or something)
                    'so we'll deal with it our way
                    SelectedText = Clipboard.GetText(TextDataFormat.UnicodeText) 'this will ReColor all the text
                    Text = Text & SelectedText
                    'RecolorWords("Shadows Sub Paste")
                End If
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try
        End Sub

        Protected Overrides Sub OnDragDrop(dragEvent As DragEventArgs)
            MyBase.OnDragDrop(dragEvent)
            _needColorAll = True
        End Sub

        ''' <summary>
        '''     Moves the current selection in the Tool to the Clipboard
        ''' </summary>
        Shadows Sub Cut()
            Try
                Clipboard.SetText(SelectedText, TextDataFormat.Rtf)
                SelectedText = ""
                RecolorWords()
            Catch ex As Exception
                Debug.WriteLine(ex.Message)
            End Try
        End Sub

        'Dim _selectedText As String
        'Public Overrides Property SelectedText As String
        '    Get
        '        Return _selectedText
        '    End Get
        '    Set
        '        NeedColorAll = True
        '        _selectedText = value
        '    End Set
        'End Property

#End Region
    End Class

    Public Class ColoredWord
        Public Color As Color
        Public SearchPattern As String

        Public Sub New()
            MyBase.New()

            Color = Color.Black
            SearchPattern = ""
        End Sub
    End Class

    Public Class ColoredWords

#Region "Fields"

        Private Shared _initiated As Boolean
        Private Const Pattern As String = "(?<=^|\W)( KEYWORD_PLACEHOLDER ) (?=\W|$)"

        Private Shared _coloredWords As IEnumerable(Of ColoredWord)
        Private Shared _autoCompleteLst As IEnumerable(Of String)

        Private Shared ReadOnly AllSqlKeyWords() As String = New String() _
            {"Parameters", "Transform", "Select", "All", "Distinct", "DistinctRow", "Top", "Percent", "As",
             "Insert", "Delete", "Update", "Into", "Set",
             "From",
             "Values", "On",
             "Where", "Having",
             "Order", "Group", "By", "Pivot",
             "Create", "Drop", "Add", "Alter",
             "Table", "Column", "Proc", "Exec",
             "AutoIncrement",
             "Text", "Memo", "Number", "YesNo", "OLEObject",
             "Short", "Long", "Single", "Double", "Currency", "DateTime", "Bit", "Byte", "GUID", "BigBinary",
             "LongBinary", "VarBinary", "LongText", "VarChar", "Decimal"}

        Private Shared ReadOnly AllFunctionNames() As String = New String() {"LBound", "UBound",
                                                                           "Asc", "CBool", "CByte", "CCur", "CDate",
                                                                           "CDbl", "CDec", "Chr", "Chr$", "CInt", "CLng",
                                                                           "CSng", "CStr", "CVar", "CVDate",
                                                                           "DateSerial", "DateValue", "Day",
                                                                           "FormatCurrency", "FormatDateTime",
                                                                           "FormatNumber", "FormatPercent",
                                                                           "GUIDFromString", "Hex", " Hex$", "Hour",
                                                                           "Minute", "Month", "Nz", "Oct", "Oct$",
                                                                           "Second", "Str", "Str$", "StrConv",
                                                                           "StringFromGUID", "TimeSerial", "TimeValue",
                                                                           "Val", "Weekday", "Year",
                                                                           "CurrentUser", "Eval", "HyperlinkPart",
                                                                           "IMEStatus", "Partition",
                                                                           "CDate", "CVDate", "Date", "Date$", "DateAdd",
                                                                           "DateDiff", "DatePart", "DateSerial",
                                                                           "DateValue", "Day", "Hour", "IsDate",
                                                                           "Minute", "Month", "MonthName", "Now",
                                                                           "Second", "Time", "Time$", "Timer",
                                                                           "TimeSerial", "TimeValue", "Weekday",
                                                                           "WeekdayName", "Year",
                                                                           "DAvg", "DCount", "DFirst", "DLast",
                                                                           "DLookup", "DMax", "DMin", "DStDev", "DSum",
                                                                           "DVar", "DVarP",
                                                                           "CVErr", "Err", "Error", "Error$", "IsError",
                                                                           "DDB", "FV", "IPmt", "IRR", "MIRR", "NPer",
                                                                           "NPV", "Pmt", "PPmt", "PV", "Rate", "SLN",
                                                                           "SYD",
                                                                           "IsArray", "IsDate", "IsEmpty", "IsError",
                                                                           "IsMissing", "IsNull", "IsNumeric",
                                                                           "IsObject", "TypeName", "VarType",
                                                                           "Abs", "Atn", "Cos", "Exp", "Fix", "Int",
                                                                           "Log", "Rnd", "Round", "Sgn", "Sin", "Sqr",
                                                                           "Tan",
                                                                           "Choose", "IIf", "Switch",
                                                                           "Avg", "Count", "Max", "Min", "StDev",
                                                                           "StDevP", "Sum", "Var", "VarP",
                                                                           "Asc", "Chr", "Chr$", "Format", "Format$",
                                                                           "GUIDFromString", "InStr", "InStrRev",
                                                                           "LCase", "LCase$", "Left", "Left$", "Len",
                                                                           "LTrim", "LTrim$", "Mid", "Mid$", "Replace",
                                                                           "Right", "Node73", "Right$", "RTrim",
                                                                           "RTrim$", "Space", "Space$", "StrComp",
                                                                           "StrConv", "String", "String$",
                                                                           "StringFromGUID", "StrReverse", "Trim",
                                                                           "Trim$", "UCase", "UCase$"
                                                                          }

        Private Shared ReadOnly AllSqlLiteralOperators() As String = New String() _
            {"Inner", "Outer", "Join",
             "And", "Or", "Xor", "Not", "Between", "Like", "In", "Eqv", "Imp",
             "Null", "Exists", "Mod"}

        Private Shared ReadOnly AllSqlOperators() As String = New String() _
            {"*", "^", "\", "/", "+", "-", ">", "<", "=", "&"}

        Private Shared ReadOnly AllSqlPunctuation() As String = New String() {";", ",", ".", "(", ")"}

        Private Shared ReadOnly RegexComments() As String = New String() {"(--.*(\n|\f|$)) | (/ \*  [\W+\w]{0,}?  \* /)"} _
        'notice the Lazy match for /* Comments */
        Private Shared ReadOnly RegexStrings() As String = New String() {"'[\W+\w]{0,}?'|""[\W+\w]{0,}?"""} _
        ' notice the "Lazy Matching" for ' ' and for " "

#End Region

#Region "Properties"

        Public Shared ReadOnly Property GetColoredWords As IEnumerable(Of ColoredWord)
            Get
                Return _ColoredWords
            End Get
        End Property

        Public Shared ReadOnly Property GetAutoCompleteList(word As String) As IEnumerable(Of String)
            Get
                Return From item In _autoCompleteLst
                    Where item.StartsWith(word, StringComparison.OrdinalIgnoreCase)
            End Get
        End Property

        Public Shared ReadOnly Property Initiated As Boolean
            Get
                Return _Initiated
            End Get
        End Property

#End Region

#Region "Colors"

        Private Shared _keyWordsColor As Color = Color.Blue

        Public Shared Property KeyWordsColor As Color
            Get
                Return _KeyWordsColor
            End Get
            Set
                If _KeyWordsColor <> value Then
                    _KeyWordsColor = value
                End If
            End Set
        End Property

        Private Shared _functionsColor As Color = Color.DeepPink

        Public Shared Property FunctionsColor As Color
            Get
                Return _FunctionsColor
            End Get
            Set
                If _FunctionsColor <> value Then
                    _FunctionsColor = value
                End If
            End Set
        End Property

        Private Shared _sqlLiteralOperatorsColor As Color = Color.DarkGray

        Public Shared Property SqlLiteralOperatorsColor As Color
            Get
                Return _SqlLiteralOperatorsColor
            End Get
            Set
                If _SqlLiteralOperatorsColor <> value Then
                    _SqlLiteralOperatorsColor = value
                End If
            End Set
        End Property

        Private Shared _sqlOperatorsColor As Color = Color.DarkGray

        Public Shared Property SqlOperatorsColor As Color
            Get
                Return _SqlOperatorsColor
            End Get
            Set
                If _SqlOperatorsColor <> value Then
                    _SqlOperatorsColor = value
                End If
            End Set
        End Property

        Private Shared _sqlPunctuationColor As Color = Color.DarkGray

        Public Shared Property SqlPunctuationColor As Color
            Get
                Return _SqlPunctuationColor
            End Get
            Set
                If _SqlPunctuationColor <> value Then
                    _SqlPunctuationColor = value
                End If
            End Set
        End Property

        Private Shared _commentsColor As Color = Color.Green

        Public Shared Property CommentsColor As Color
            Get
                Return _CommentsColor
            End Get
            Set
                If _CommentsColor <> value Then
                    _CommentsColor = value
                End If
            End Set
        End Property

        Private Shared _stringsColor As Color = Color.Red

        Public Shared Property StringsColor As Color
            Get
                Return _StringsColor
            End Get
            Set
                If _StringsColor <> value Then
                    _StringsColor = value
                End If
            End Set
        End Property

#End Region

#Region "Functions"

        Private Shared Function EscapeTxt4Pattern(plainTxt As String) As String
            Dim funReplace = Function(m As Match) "\" & m.Value

            Const escapePattern = "\^| \\|	\[|	\]|	\+| \*|	\.|	\(|	\)|	\$"

            Return Regex.Replace(plainTxt, escapePattern,
                                 funReplace,
                                 RegexOptions.IgnorePatternWhitespace)
        End Function

        Private Shared Function GetPattern4Words(replacement As String) As String
            Return Pattern.Replace("KEYWORD_PLACEHOLDER", Replacement)
        End Function

        Public Shared Function SqlKeyWords() As IEnumerable(Of ColoredWord)
            Return From kw In AllSqlKeyWords
                Select New ColoredWord() With
                    {.color = KeyWordsColor,
                    .SearchPattern = GetPattern4Words(kw)}
        End Function

        Public Shared Function FunctionNames() As IEnumerable(Of ColoredWord)
            Return From fun In AllFunctionNames
                Select New ColoredWord() With
                    {.color = FunctionsColor,
                    .SearchPattern = GetPattern4Words(fun.Replace("$", "$$"))}
        End Function

        Public Shared Function SqlOperators() As IEnumerable(Of ColoredWord)

            Return From opr In AllSqlOperators
                Select New ColoredWord() With
                    {.color = SqlOperatorsColor,
                    .SearchPattern = EscapeTxt4Pattern(opr)}
        End Function

        Public Shared Function SqlLiteralOperators() As IEnumerable(Of ColoredWord)

            Return From opr In AllSqlLiteralOperators
                Select New ColoredWord() With
                    {.color = SqlLiteralOperatorsColor,
                    .SearchPattern = GetPattern4Words(EscapeTxt4Pattern(opr))}
        End Function

        Public Shared Function SqlPunctuation() As IEnumerable(Of ColoredWord)

            Return From pun In AllSqlPunctuation
                Select New ColoredWord() With
                    {.color = SqlPunctuationColor,
                    .SearchPattern = EscapeTxt4Pattern(pun)}
        End Function

        Public Shared Function Comments() As IEnumerable(Of ColoredWord)
            Return From com In regexComments
                Select New ColoredWord() With
                    {.color = CommentsColor,
                    .SearchPattern = com}
        End Function

        Public Shared Function Strings() As IEnumerable(Of ColoredWord)
            Return From s In regexStrings
                Select New ColoredWord() With
                    {.color = StringsColor,
                    .SearchPattern = s}
        End Function

        Public Shared Function NonSqlChars() As IEnumerable(Of ColoredWord)

            Dim regexNonSqlChars = New String() {"\[|\]|'|:", "\[ [\W+\w]{0,}? \]"} 'notice the "Lazy Matching" for [ ]

            Return From nsc In regexNonSqlChars _
                Select New ColoredWord() With _
                    {.color = Color.Black, _
                    .SearchPattern = nsc}
        End Function

        Public Shared Function GetCommentsPattern() As String
            Return String.Join(" ", regexComments)
        End Function

        Public Shared Function OverAllColoredWords() As IEnumerable(Of ColoredWord)
            Return NonSqlChars.Union(Strings).Union(Comments)

            'TODO: Fix this return
            'look at this line , 
            'and guess what colors would be there:    
            '                   [ select ' from xT ] where ' what ever
        End Function

#End Region

#Region "Sub Routines"

        Public Shared Sub Initiate()
            _ColoredWords = SqlKeyWords().
                Union(FunctionNames()).
                Union(SqlLiteralOperators()).
                Union(SqlPunctuation()).
                Union(SqlOperators())  '. _
            'Union(NonSqlChars())

            _autoCompleteLst = AllSqlKeyWords.Union(AllFunctionNames).
                Union(AllSqlLiteralOperators)

            _Initiated = True
        End Sub

        Public Shared Sub ResetColors()
            SqlLiteralOperatorsColor = Color.DarkGray
            FunctionsColor = Color.DeepPink
            SqlOperatorsColor = Color.DarkGray
            KeyWordsColor = Color.Blue
            StringsColor = Color.Red
            CommentsColor = Color.Green
        End Sub

#End Region
    End Class
End Namespace