<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_EditorOptions
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.btn_PunctuationColor = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.btn_CommentsColor = New System.Windows.Forms.Button
        Me.btn_StringsColor = New System.Windows.Forms.Button
        Me.btn_OperatorsColor = New System.Windows.Forms.Button
        Me.btn_Literal_OperatorsColor = New System.Windows.Forms.Button
        Me.btn_FunctionsColor = New System.Windows.Forms.Button
        Me.btn_KeyWords_Color = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btn_OK = New System.Windows.Forms.Button
        Me.btn_Cancel = New System.Windows.Forms.Button
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.btn_PunctuationColor)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.btn_CommentsColor)
        Me.GroupBox1.Controls.Add(Me.btn_StringsColor)
        Me.GroupBox1.Controls.Add(Me.btn_OperatorsColor)
        Me.GroupBox1.Controls.Add(Me.btn_Literal_OperatorsColor)
        Me.GroupBox1.Controls.Add(Me.btn_FunctionsColor)
        Me.GroupBox1.Controls.Add(Me.btn_KeyWords_Color)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(243, 161)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Colors"
        '
        'btn_PunctuationColor
        '
        Me.btn_PunctuationColor.AutoSize = True
        Me.btn_PunctuationColor.Location = New System.Drawing.Point(15, 107)
        Me.btn_PunctuationColor.Name = "btn_PunctuationColor"
        Me.btn_PunctuationColor.Size = New System.Drawing.Size(100, 23)
        Me.btn_PunctuationColor.TabIndex = 7
        Me.btn_PunctuationColor.Text = "Punctuation"
        Me.ToolTip1.SetToolTip(Me.btn_PunctuationColor, "Example: 'some text in String Format'")
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(133, 126)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(100, 26)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Reset All"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btn_CommentsColor
        '
        Me.btn_CommentsColor.Location = New System.Drawing.Point(133, 78)
        Me.btn_CommentsColor.Name = "btn_CommentsColor"
        Me.btn_CommentsColor.Size = New System.Drawing.Size(100, 22)
        Me.btn_CommentsColor.TabIndex = 5
        Me.btn_CommentsColor.Text = "Comments"
        Me.btn_CommentsColor.UseVisualStyleBackColor = True
        '
        'btn_StringsColor
        '
        Me.btn_StringsColor.AutoSize = True
        Me.btn_StringsColor.Location = New System.Drawing.Point(15, 77)
        Me.btn_StringsColor.Name = "btn_StringsColor"
        Me.btn_StringsColor.Size = New System.Drawing.Size(100, 23)
        Me.btn_StringsColor.TabIndex = 4
        Me.btn_StringsColor.Text = "Strings"
        Me.ToolTip1.SetToolTip(Me.btn_StringsColor, "Example: 'some text in String Format'")
        '
        'btn_OperatorsColor
        '
        Me.btn_OperatorsColor.AutoSize = True
        Me.btn_OperatorsColor.Location = New System.Drawing.Point(133, 48)
        Me.btn_OperatorsColor.Name = "btn_OperatorsColor"
        Me.btn_OperatorsColor.Size = New System.Drawing.Size(100, 23)
        Me.btn_OperatorsColor.TabIndex = 3
        Me.btn_OperatorsColor.Text = "Operators"
        Me.ToolTip1.SetToolTip(Me.btn_OperatorsColor, "Example: ; > < = + .................")
        '
        'btn_Literal_OperatorsColor
        '
        Me.btn_Literal_OperatorsColor.AutoSize = True
        Me.btn_Literal_OperatorsColor.Location = New System.Drawing.Point(15, 48)
        Me.btn_Literal_OperatorsColor.Name = "btn_Literal_OperatorsColor"
        Me.btn_Literal_OperatorsColor.Size = New System.Drawing.Size(100, 23)
        Me.btn_Literal_OperatorsColor.TabIndex = 2
        Me.btn_Literal_OperatorsColor.Text = "Literal Operators"
        Me.ToolTip1.SetToolTip(Me.btn_Literal_OperatorsColor, "Example: Join And Inion.........")
        '
        'btn_FunctionsColor
        '
        Me.btn_FunctionsColor.AutoSize = True
        Me.btn_FunctionsColor.Location = New System.Drawing.Point(133, 19)
        Me.btn_FunctionsColor.Name = "btn_FunctionsColor"
        Me.btn_FunctionsColor.Size = New System.Drawing.Size(100, 23)
        Me.btn_FunctionsColor.TabIndex = 1
        Me.btn_FunctionsColor.Text = "Functions"
        Me.ToolTip1.SetToolTip(Me.btn_FunctionsColor, "Example: Date$ Format$ Left .......")
        '
        'btn_KeyWords_Color
        '
        Me.btn_KeyWords_Color.AutoSize = True
        Me.btn_KeyWords_Color.Location = New System.Drawing.Point(15, 19)
        Me.btn_KeyWords_Color.Name = "btn_KeyWords_Color"
        Me.btn_KeyWords_Color.Size = New System.Drawing.Size(100, 23)
        Me.btn_KeyWords_Color.TabIndex = 0
        Me.btn_KeyWords_Color.Text = "Key Words"
        Me.ToolTip1.SetToolTip(Me.btn_KeyWords_Color, "Example: Select  From Where .....")
        '
        'btn_OK
        '
        Me.btn_OK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btn_OK.Location = New System.Drawing.Point(115, 208)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(67, 27)
        Me.btn_OK.TabIndex = 1
        Me.btn_OK.Text = "&OK"
        Me.btn_OK.UseVisualStyleBackColor = True
        '
        'btn_Cancel
        '
        Me.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn_Cancel.Location = New System.Drawing.Point(188, 208)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(67, 27)
        Me.btn_Cancel.TabIndex = 2
        Me.btn_Cancel.Text = "&Cancel"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'frm_EditorOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(267, 244)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_EditorOptions"
        Me.Text = "Editor Options"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btn_OperatorsColor As System.Windows.Forms.Button
    Friend WithEvents btn_Literal_OperatorsColor As System.Windows.Forms.Button
    Friend WithEvents btn_FunctionsColor As System.Windows.Forms.Button
    Friend WithEvents btn_KeyWords_Color As System.Windows.Forms.Button
    Friend WithEvents btn_StringsColor As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btn_CommentsColor As System.Windows.Forms.Button
    Friend WithEvents btn_OK As System.Windows.Forms.Button
    Friend WithEvents btn_Cancel As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btn_PunctuationColor As System.Windows.Forms.Button
End Class
