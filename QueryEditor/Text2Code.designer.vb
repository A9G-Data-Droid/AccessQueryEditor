<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Text2Code
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
        Me.txt_PlainText = New System.Windows.Forms.TextBox
        Me.txt_CodeText = New System.Windows.Forms.TextBox
        Me.btn_Convert = New System.Windows.Forms.Button
        Me.chk_StrBldr = New System.Windows.Forms.CheckBox
        Me.chk_AddNewLine = New System.Windows.Forms.CheckBox
        Me.lbl_Lang = New System.Windows.Forms.Label
        Me.comb_Lang = New System.Windows.Forms.ComboBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.grop_PlainText = New System.Windows.Forms.GroupBox
        Me.grop_OutPut = New System.Windows.Forms.GroupBox
        Me.btn_CpyOutPut = New System.Windows.Forms.Button
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.grop_PlainText.SuspendLayout()
        Me.grop_OutPut.SuspendLayout()
        Me.SuspendLayout()
        '
        'txt_PlainText
        '
        Me.txt_PlainText.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txt_PlainText.Location = New System.Drawing.Point(3, 17)
        Me.txt_PlainText.Multiline = True
        Me.txt_PlainText.Name = "txt_PlainText"
        Me.txt_PlainText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txt_PlainText.Size = New System.Drawing.Size(528, 121)
        Me.txt_PlainText.TabIndex = 0
        '
        'txt_CodeText
        '
        Me.txt_CodeText.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txt_CodeText.Location = New System.Drawing.Point(3, 17)
        Me.txt_CodeText.Multiline = True
        Me.txt_CodeText.Name = "txt_CodeText"
        Me.txt_CodeText.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txt_CodeText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txt_CodeText.Size = New System.Drawing.Size(528, 101)
        Me.txt_CodeText.TabIndex = 1
        '
        'btn_Convert
        '
        Me.btn_Convert.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Convert.Location = New System.Drawing.Point(447, 7)
        Me.btn_Convert.Name = "btn_Convert"
        Me.btn_Convert.Size = New System.Drawing.Size(84, 31)
        Me.btn_Convert.TabIndex = 2
        Me.btn_Convert.Text = "Convert"
        Me.btn_Convert.UseVisualStyleBackColor = True
        '
        'chk_StrBldr
        '
        Me.chk_StrBldr.AutoSize = True
        Me.chk_StrBldr.Checked = True
        Me.chk_StrBldr.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk_StrBldr.Location = New System.Drawing.Point(7, 12)
        Me.chk_StrBldr.Name = "chk_StrBldr"
        Me.chk_StrBldr.Size = New System.Drawing.Size(140, 17)
        Me.chk_StrBldr.TabIndex = 3
        Me.chk_StrBldr.Text = "Use StringBuilder object"
        Me.chk_StrBldr.UseVisualStyleBackColor = True
        '
        'chk_AddNewLine
        '
        Me.chk_AddNewLine.AutoSize = True
        Me.chk_AddNewLine.Location = New System.Drawing.Point(7, 31)
        Me.chk_AddNewLine.Name = "chk_AddNewLine"
        Me.chk_AddNewLine.Size = New System.Drawing.Size(87, 17)
        Me.chk_AddNewLine.TabIndex = 4
        Me.chk_AddNewLine.Text = "Add new line"
        Me.chk_AddNewLine.UseVisualStyleBackColor = True
        '
        'lbl_Lang
        '
        Me.lbl_Lang.AutoSize = True
        Me.lbl_Lang.Location = New System.Drawing.Point(174, 12)
        Me.lbl_Lang.Name = "lbl_Lang"
        Me.lbl_Lang.Size = New System.Drawing.Size(54, 13)
        Me.lbl_Lang.TabIndex = 5
        Me.lbl_Lang.Text = "Language"
        '
        'comb_Lang
        '
        Me.comb_Lang.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.comb_Lang.FormattingEnabled = True
        Me.comb_Lang.Items.AddRange(New Object() {"VB", "C#‎"})
        Me.comb_Lang.Location = New System.Drawing.Point(234, 12)
        Me.comb_Lang.Name = "comb_Lang"
        Me.comb_Lang.Size = New System.Drawing.Size(79, 21)
        Me.comb_Lang.TabIndex = 6
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.BackColor = System.Drawing.SystemColors.Control
        Me.SplitContainer1.Location = New System.Drawing.Point(4, 54)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.grop_PlainText)
        Me.SplitContainer1.Panel1MinSize = 50
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.grop_OutPut)
        Me.SplitContainer1.Panel2MinSize = 50
        Me.SplitContainer1.Size = New System.Drawing.Size(534, 266)
        Me.SplitContainer1.SplitterDistance = 141
        Me.SplitContainer1.TabIndex = 8
        '
        'grop_PlainText
        '
        Me.grop_PlainText.Controls.Add(Me.txt_PlainText)
        Me.grop_PlainText.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grop_PlainText.Location = New System.Drawing.Point(0, 0)
        Me.grop_PlainText.Name = "grop_PlainText"
        Me.grop_PlainText.Size = New System.Drawing.Size(534, 141)
        Me.grop_PlainText.TabIndex = 1
        Me.grop_PlainText.TabStop = False
        Me.grop_PlainText.Text = "Plain Text"
        '
        'grop_OutPut
        '
        Me.grop_OutPut.Controls.Add(Me.txt_CodeText)
        Me.grop_OutPut.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grop_OutPut.Location = New System.Drawing.Point(0, 0)
        Me.grop_OutPut.Name = "grop_OutPut"
        Me.grop_OutPut.Size = New System.Drawing.Size(534, 121)
        Me.grop_OutPut.TabIndex = 2
        Me.grop_OutPut.TabStop = False
        Me.grop_OutPut.Text = "Out Put"
        '
        'btn_CpyOutPut
        '
        Me.btn_CpyOutPut.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_CpyOutPut.Location = New System.Drawing.Point(357, 7)
        Me.btn_CpyOutPut.Name = "btn_CpyOutPut"
        Me.btn_CpyOutPut.Size = New System.Drawing.Size(84, 31)
        Me.btn_CpyOutPut.TabIndex = 9
        Me.btn_CpyOutPut.Text = "Copy Out Put"
        Me.btn_CpyOutPut.UseVisualStyleBackColor = True
        '
        'frm_Text2Code
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(542, 323)
        Me.Controls.Add(Me.btn_CpyOutPut)
        Me.Controls.Add(Me.lbl_Lang)
        Me.Controls.Add(Me.chk_StrBldr)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.chk_AddNewLine)
        Me.Controls.Add(Me.btn_Convert)
        Me.Controls.Add(Me.comb_Lang)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MinimumSize = New System.Drawing.Size(550, 350)
        Me.Name = "frm_Text2Code"
        Me.Text = "Text To Code"
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.grop_PlainText.ResumeLayout(False)
        Me.grop_PlainText.PerformLayout()
        Me.grop_OutPut.ResumeLayout(False)
        Me.grop_OutPut.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_PlainText As System.Windows.Forms.TextBox
    Friend WithEvents txt_CodeText As System.Windows.Forms.TextBox
    Friend WithEvents btn_Convert As System.Windows.Forms.Button
    Friend WithEvents chk_StrBldr As System.Windows.Forms.CheckBox
    Friend WithEvents chk_AddNewLine As System.Windows.Forms.CheckBox
    Friend WithEvents lbl_Lang As System.Windows.Forms.Label
    Friend WithEvents comb_Lang As System.Windows.Forms.ComboBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents grop_OutPut As System.Windows.Forms.GroupBox
    Friend WithEvents grop_PlainText As System.Windows.Forms.GroupBox
    Friend WithEvents btn_CpyOutPut As System.Windows.Forms.Button

End Class
