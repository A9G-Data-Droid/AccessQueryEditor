<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FindReplace
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txt_FindStr = New System.Windows.Forms.TextBox
        Me.lbl_FindWhat_lbl = New System.Windows.Forms.Label
        Me.btn_FindNext = New System.Windows.Forms.Button
        Me.btn_Close = New System.Windows.Forms.Button
        Me.Status = New System.Windows.Forms.StatusStrip
        Me.FindStatus = New System.Windows.Forms.ToolStripStatusLabel
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel
        Me.txt_ReplaceWith = New System.Windows.Forms.TextBox
        Me.btn_Replace = New System.Windows.Forms.Button
        Me.btn_ReplaceALL = New System.Windows.Forms.Button
        Me.lbl_ReplaceWith_lbl = New System.Windows.Forms.Label
        Me.TabFindReplace = New System.Windows.Forms.TabControl
        Me.TabPage_Find = New System.Windows.Forms.TabPage
        Me.Pnl_Btns = New System.Windows.Forms.Panel
        Me.TabPage_Replace = New System.Windows.Forms.TabPage
        Me.btn_ExpandOptions = New System.Windows.Forms.Button
        Me.chk_MatchCase = New System.Windows.Forms.CheckBox
        Me.Grop_FindOptions = New System.Windows.Forms.GroupBox
        Me.chk_SearchUp = New System.Windows.Forms.CheckBox
        Me.chk_MatchWholeWord = New System.Windows.Forms.CheckBox
        Me.lbl_Cover = New System.Windows.Forms.Label
        Me.Status.SuspendLayout()
        Me.TabFindReplace.SuspendLayout()
        Me.TabPage_Find.SuspendLayout()
        Me.Pnl_Btns.SuspendLayout()
        Me.Grop_FindOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'txt_FindStr
        '
        Me.txt_FindStr.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_FindStr.Location = New System.Drawing.Point(2, 25)
        Me.txt_FindStr.Name = "txt_FindStr"
        Me.txt_FindStr.Size = New System.Drawing.Size(385, 20)
        Me.txt_FindStr.TabIndex = 0
        '
        'lbl_FindWhat_lbl
        '
        Me.lbl_FindWhat_lbl.AutoSize = True
        Me.lbl_FindWhat_lbl.Location = New System.Drawing.Point(2, 7)
        Me.lbl_FindWhat_lbl.Name = "lbl_FindWhat_lbl"
        Me.lbl_FindWhat_lbl.Size = New System.Drawing.Size(56, 13)
        Me.lbl_FindWhat_lbl.TabIndex = 3
        Me.lbl_FindWhat_lbl.Text = "Find What"
        '
        'btn_FindNext
        '
        Me.btn_FindNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_FindNext.Location = New System.Drawing.Point(397, 37)
        Me.btn_FindNext.Name = "btn_FindNext"
        Me.btn_FindNext.Size = New System.Drawing.Size(162, 25)
        Me.btn_FindNext.TabIndex = 2
        Me.btn_FindNext.Text = "&Find Next"
        Me.btn_FindNext.UseVisualStyleBackColor = True
        '
        'btn_Close
        '
        Me.btn_Close.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Close.Location = New System.Drawing.Point(397, 12)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(162, 25)
        Me.btn_Close.TabIndex = 6
        Me.btn_Close.Text = "&Close"
        Me.btn_Close.UseVisualStyleBackColor = True
        '
        'Status
        '
        Me.Status.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FindStatus, Me.ToolStripStatusLabel1})
        Me.Status.Location = New System.Drawing.Point(0, 167)
        Me.Status.Name = "Status"
        Me.Status.Size = New System.Drawing.Size(583, 22)
        Me.Status.TabIndex = 7
        '
        'FindStatus
        '
        Me.FindStatus.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
                    Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
                    Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.FindStatus.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.FindStatus.Margin = New System.Windows.Forms.Padding(0, 3, 0, 0)
        Me.FindStatus.Name = "FindStatus"
        Me.FindStatus.Size = New System.Drawing.Size(67, 19)
        Me.FindStatus.Text = "No Results"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.BorderSides = CType((((System.Windows.Forms.ToolStripStatusLabelBorderSides.Left Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Top) _
                    Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Right) _
                    Or System.Windows.Forms.ToolStripStatusLabelBorderSides.Bottom), System.Windows.Forms.ToolStripStatusLabelBorderSides)
        Me.ToolStripStatusLabel1.BorderStyle = System.Windows.Forms.Border3DStyle.Sunken
        Me.ToolStripStatusLabel1.Margin = New System.Windows.Forms.Padding(0, 3, 0, 0)
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(501, 19)
        Me.ToolStripStatusLabel1.Spring = True
        '
        'txt_ReplaceWith
        '
        Me.txt_ReplaceWith.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_ReplaceWith.Location = New System.Drawing.Point(2, 67)
        Me.txt_ReplaceWith.Name = "txt_ReplaceWith"
        Me.txt_ReplaceWith.Size = New System.Drawing.Size(385, 20)
        Me.txt_ReplaceWith.TabIndex = 1
        Me.txt_ReplaceWith.Visible = False
        '
        'btn_Replace
        '
        Me.btn_Replace.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Replace.Location = New System.Drawing.Point(397, 63)
        Me.btn_Replace.Name = "btn_Replace"
        Me.btn_Replace.Size = New System.Drawing.Size(81, 25)
        Me.btn_Replace.TabIndex = 4
        Me.btn_Replace.Text = "Replace"
        Me.btn_Replace.UseVisualStyleBackColor = True
        Me.btn_Replace.Visible = False
        '
        'btn_ReplaceALL
        '
        Me.btn_ReplaceALL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_ReplaceALL.Location = New System.Drawing.Point(478, 63)
        Me.btn_ReplaceALL.Name = "btn_ReplaceALL"
        Me.btn_ReplaceALL.Size = New System.Drawing.Size(81, 25)
        Me.btn_ReplaceALL.TabIndex = 5
        Me.btn_ReplaceALL.Text = "Replace All"
        Me.btn_ReplaceALL.UseVisualStyleBackColor = True
        Me.btn_ReplaceALL.Visible = False
        '
        'lbl_ReplaceWith_lbl
        '
        Me.lbl_ReplaceWith_lbl.AutoSize = True
        Me.lbl_ReplaceWith_lbl.Location = New System.Drawing.Point(2, 49)
        Me.lbl_ReplaceWith_lbl.Name = "lbl_ReplaceWith_lbl"
        Me.lbl_ReplaceWith_lbl.Size = New System.Drawing.Size(70, 13)
        Me.lbl_ReplaceWith_lbl.TabIndex = 13
        Me.lbl_ReplaceWith_lbl.Text = "Replace With"
        Me.lbl_ReplaceWith_lbl.Visible = False
        '
        'TabFindReplace
        '
        Me.TabFindReplace.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabFindReplace.Controls.Add(Me.TabPage_Find)
        Me.TabFindReplace.Controls.Add(Me.TabPage_Replace)
        Me.TabFindReplace.Location = New System.Drawing.Point(0, 0)
        Me.TabFindReplace.Name = "TabFindReplace"
        Me.TabFindReplace.SelectedIndex = 0
        Me.TabFindReplace.Size = New System.Drawing.Size(583, 131)
        Me.TabFindReplace.TabIndex = 15
        '
        'TabPage_Find
        '
        Me.TabPage_Find.Controls.Add(Me.Pnl_Btns)
        Me.TabPage_Find.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_Find.Name = "TabPage_Find"
        Me.TabPage_Find.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Find.Size = New System.Drawing.Size(575, 105)
        Me.TabPage_Find.TabIndex = 0
        Me.TabPage_Find.Text = "Find"
        Me.TabPage_Find.UseVisualStyleBackColor = True
        '
        'Pnl_Btns
        '
        Me.Pnl_Btns.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Pnl_Btns.Controls.Add(Me.btn_Close)
        Me.Pnl_Btns.Controls.Add(Me.btn_FindNext)
        Me.Pnl_Btns.Controls.Add(Me.lbl_FindWhat_lbl)
        Me.Pnl_Btns.Controls.Add(Me.btn_Replace)
        Me.Pnl_Btns.Controls.Add(Me.lbl_ReplaceWith_lbl)
        Me.Pnl_Btns.Controls.Add(Me.txt_FindStr)
        Me.Pnl_Btns.Controls.Add(Me.txt_ReplaceWith)
        Me.Pnl_Btns.Controls.Add(Me.btn_ReplaceALL)
        Me.Pnl_Btns.Location = New System.Drawing.Point(4, 4)
        Me.Pnl_Btns.Name = "Pnl_Btns"
        Me.Pnl_Btns.Size = New System.Drawing.Size(565, 94)
        Me.Pnl_Btns.TabIndex = 0
        '
        'TabPage_Replace
        '
        Me.TabPage_Replace.Location = New System.Drawing.Point(4, 22)
        Me.TabPage_Replace.Name = "TabPage_Replace"
        Me.TabPage_Replace.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage_Replace.Size = New System.Drawing.Size(575, 105)
        Me.TabPage_Replace.TabIndex = 1
        Me.TabPage_Replace.Text = "Replace"
        Me.TabPage_Replace.UseVisualStyleBackColor = True
        '
        'btn_ExpandOptions
        '
        Me.btn_ExpandOptions.Location = New System.Drawing.Point(5, 137)
        Me.btn_ExpandOptions.Name = "btn_ExpandOptions"
        Me.btn_ExpandOptions.Size = New System.Drawing.Size(20, 20)
        Me.btn_ExpandOptions.TabIndex = 6
        Me.btn_ExpandOptions.TabStop = False
        Me.btn_ExpandOptions.Text = "+"
        Me.btn_ExpandOptions.UseVisualStyleBackColor = True
        '
        'chk_MatchCase
        '
        Me.chk_MatchCase.AutoSize = True
        Me.chk_MatchCase.Location = New System.Drawing.Point(20, 32)
        Me.chk_MatchCase.Name = "chk_MatchCase"
        Me.chk_MatchCase.Size = New System.Drawing.Size(82, 17)
        Me.chk_MatchCase.TabIndex = 2
        Me.chk_MatchCase.Text = "Match Case"
        Me.chk_MatchCase.UseVisualStyleBackColor = True
        '
        'Grop_FindOptions
        '
        Me.Grop_FindOptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Grop_FindOptions.Controls.Add(Me.chk_SearchUp)
        Me.Grop_FindOptions.Controls.Add(Me.chk_MatchWholeWord)
        Me.Grop_FindOptions.Controls.Add(Me.chk_MatchCase)
        Me.Grop_FindOptions.Location = New System.Drawing.Point(8, 142)
        Me.Grop_FindOptions.Name = "Grop_FindOptions"
        Me.Grop_FindOptions.Size = New System.Drawing.Size(553, 128)
        Me.Grop_FindOptions.TabIndex = 4
        Me.Grop_FindOptions.TabStop = False
        Me.Grop_FindOptions.Text = "    Find Options "
        '
        'chk_SearchUp
        '
        Me.chk_SearchUp.AutoSize = True
        Me.chk_SearchUp.Location = New System.Drawing.Point(21, 81)
        Me.chk_SearchUp.Name = "chk_SearchUp"
        Me.chk_SearchUp.Size = New System.Drawing.Size(75, 17)
        Me.chk_SearchUp.TabIndex = 4
        Me.chk_SearchUp.Text = "Search Up"
        Me.chk_SearchUp.UseVisualStyleBackColor = True
        '
        'chk_MatchWholeWord
        '
        Me.chk_MatchWholeWord.AutoSize = True
        Me.chk_MatchWholeWord.Location = New System.Drawing.Point(20, 55)
        Me.chk_MatchWholeWord.Name = "chk_MatchWholeWord"
        Me.chk_MatchWholeWord.Size = New System.Drawing.Size(113, 17)
        Me.chk_MatchWholeWord.TabIndex = 3
        Me.chk_MatchWholeWord.Text = "Match whole word"
        Me.chk_MatchWholeWord.UseVisualStyleBackColor = True
        '
        'lbl_Cover
        '
        Me.lbl_Cover.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbl_Cover.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lbl_Cover.Location = New System.Drawing.Point(547, 137)
        Me.lbl_Cover.Name = "lbl_Cover"
        Me.lbl_Cover.Size = New System.Drawing.Size(25, 45)
        Me.lbl_Cover.TabIndex = 16
        '
        'frm_FindReplace
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(583, 189)
        Me.Controls.Add(Me.Status)
        Me.Controls.Add(Me.lbl_Cover)
        Me.Controls.Add(Me.btn_ExpandOptions)
        Me.Controls.Add(Me.Grop_FindOptions)
        Me.Controls.Add(Me.TabFindReplace)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.Name = "frm_FindReplace"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Find & Replace"
        Me.TopMost = True
        Me.Status.ResumeLayout(False)
        Me.Status.PerformLayout()
        Me.TabFindReplace.ResumeLayout(False)
        Me.TabPage_Find.ResumeLayout(False)
        Me.Pnl_Btns.ResumeLayout(False)
        Me.Pnl_Btns.PerformLayout()
        Me.Grop_FindOptions.ResumeLayout(False)
        Me.Grop_FindOptions.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_FindStr As System.Windows.Forms.TextBox
    Friend WithEvents lbl_FindWhat_lbl As System.Windows.Forms.Label
    Friend WithEvents btn_FindNext As System.Windows.Forms.Button
    Friend WithEvents btn_Close As System.Windows.Forms.Button
    Friend WithEvents FindStatus As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Status As System.Windows.Forms.StatusStrip
    Friend WithEvents txt_ReplaceWith As System.Windows.Forms.TextBox
    Friend WithEvents btn_Replace As System.Windows.Forms.Button
    Friend WithEvents btn_ReplaceALL As System.Windows.Forms.Button
    Friend WithEvents lbl_ReplaceWith_lbl As System.Windows.Forms.Label
    Friend WithEvents TabFindReplace As System.Windows.Forms.TabControl
    Friend WithEvents TabPage_Find As System.Windows.Forms.TabPage
    Friend WithEvents TabPage_Replace As System.Windows.Forms.TabPage
    Friend WithEvents Pnl_Btns As System.Windows.Forms.Panel
    Friend WithEvents btn_ExpandOptions As System.Windows.Forms.Button
    Friend WithEvents chk_MatchCase As System.Windows.Forms.CheckBox
    Friend WithEvents Grop_FindOptions As System.Windows.Forms.GroupBox
    Friend WithEvents lbl_Cover As System.Windows.Forms.Label
    Friend WithEvents chk_MatchWholeWord As System.Windows.Forms.CheckBox
    Friend WithEvents chk_SearchUp As System.Windows.Forms.CheckBox
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
End Class
