<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AboutBox
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    Friend WithEvents btn_OK As System.Windows.Forms.Button

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btn_OK = New System.Windows.Forms.Button
        Me.txt_AuthorInf = New System.Windows.Forms.TextBox
        Me.grop_AuthorInf = New System.Windows.Forms.GroupBox
        Me.grop_AppInf = New System.Windows.Forms.GroupBox
        Me.txt_Description = New System.Windows.Forms.TextBox
        Me.lbl_Description_lbl = New System.Windows.Forms.Label
        Me.lbl_Release = New System.Windows.Forms.TextBox
        Me.lbl_Release_lbl = New System.Windows.Forms.Label
        Me.txt_Version = New System.Windows.Forms.TextBox
        Me.lbl_Version_lbl = New System.Windows.Forms.Label
        Me.grop_AuthorInf.SuspendLayout()
        Me.grop_AppInf.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_OK
        '
        Me.btn_OK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_OK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btn_OK.Location = New System.Drawing.Point(15, 446)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(83, 35)
        Me.btn_OK.TabIndex = 0
        Me.btn_OK.Text = "&OK"
        '
        'txt_AuthorInf
        '
        Me.txt_AuthorInf.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_AuthorInf.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_AuthorInf.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_AuthorInf.Location = New System.Drawing.Point(9, 22)
        Me.txt_AuthorInf.Margin = New System.Windows.Forms.Padding(6, 3, 3, 3)
        Me.txt_AuthorInf.Multiline = True
        Me.txt_AuthorInf.Name = "txt_AuthorInf"
        Me.txt_AuthorInf.ReadOnly = True
        Me.txt_AuthorInf.Size = New System.Drawing.Size(541, 115)
        Me.txt_AuthorInf.TabIndex = 2
        Me.txt_AuthorInf.TabStop = False
        Me.txt_AuthorInf.Text = "Yasser Daheek AKA vbnetskywalker" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Country" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "Syria - Hama" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Cell-Phone" & Global.Microsoft.VisualBasic.ChrW(9) & "963 - 0967398" & _
            "199" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Phone" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "033 - 435665" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Email" & Global.Microsoft.VisualBasic.ChrW(9) & Global.Microsoft.VisualBasic.ChrW(9) & "yasserdaheek@windowslive.com"
        '
        'grop_AuthorInf
        '
        Me.grop_AuthorInf.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grop_AuthorInf.Controls.Add(Me.txt_AuthorInf)
        Me.grop_AuthorInf.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grop_AuthorInf.Location = New System.Drawing.Point(15, 297)
        Me.grop_AuthorInf.Name = "grop_AuthorInf"
        Me.grop_AuthorInf.Size = New System.Drawing.Size(566, 143)
        Me.grop_AuthorInf.TabIndex = 3
        Me.grop_AuthorInf.TabStop = False
        Me.grop_AuthorInf.Text = "Author Info"
        '
        'grop_AppInf
        '
        Me.grop_AppInf.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grop_AppInf.Controls.Add(Me.txt_Description)
        Me.grop_AppInf.Controls.Add(Me.lbl_Description_lbl)
        Me.grop_AppInf.Controls.Add(Me.lbl_Release)
        Me.grop_AppInf.Controls.Add(Me.lbl_Release_lbl)
        Me.grop_AppInf.Controls.Add(Me.txt_Version)
        Me.grop_AppInf.Controls.Add(Me.lbl_Version_lbl)
        Me.grop_AppInf.Font = New System.Drawing.Font("Arial", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grop_AppInf.Location = New System.Drawing.Point(15, 12)
        Me.grop_AppInf.Name = "grop_AppInf"
        Me.grop_AppInf.Size = New System.Drawing.Size(566, 279)
        Me.grop_AppInf.TabIndex = 4
        Me.grop_AppInf.TabStop = False
        Me.grop_AppInf.Text = "Application Info"
        '
        'txt_Description
        '
        Me.txt_Description.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_Description.BackColor = System.Drawing.SystemColors.Control
        Me.txt_Description.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_Description.Location = New System.Drawing.Point(123, 109)
        Me.txt_Description.Multiline = True
        Me.txt_Description.Name = "txt_Description"
        Me.txt_Description.ReadOnly = True
        Me.txt_Description.Size = New System.Drawing.Size(428, 142)
        Me.txt_Description.TabIndex = 5
        '
        'lbl_Description_lbl
        '
        Me.lbl_Description_lbl.AutoSize = True
        Me.lbl_Description_lbl.Location = New System.Drawing.Point(5, 109)
        Me.lbl_Description_lbl.Name = "lbl_Description_lbl"
        Me.lbl_Description_lbl.Size = New System.Drawing.Size(111, 22)
        Me.lbl_Description_lbl.TabIndex = 4
        Me.lbl_Description_lbl.Text = "Description:"
        '
        'lbl_Release
        '
        Me.lbl_Release.BackColor = System.Drawing.SystemColors.Control
        Me.lbl_Release.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lbl_Release.Location = New System.Drawing.Point(121, 72)
        Me.lbl_Release.Name = "lbl_Release"
        Me.lbl_Release.ReadOnly = True
        Me.lbl_Release.Size = New System.Drawing.Size(429, 22)
        Me.lbl_Release.TabIndex = 3
        Me.lbl_Release.Text = "29 - June - 2009"
        '
        'lbl_Release_lbl
        '
        Me.lbl_Release_lbl.AutoSize = True
        Me.lbl_Release_lbl.Location = New System.Drawing.Point(30, 72)
        Me.lbl_Release_lbl.Name = "lbl_Release_lbl"
        Me.lbl_Release_lbl.Size = New System.Drawing.Size(86, 22)
        Me.lbl_Release_lbl.TabIndex = 2
        Me.lbl_Release_lbl.Text = "Release:"
        '
        'txt_Version
        '
        Me.txt_Version.BackColor = System.Drawing.SystemColors.Control
        Me.txt_Version.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txt_Version.Location = New System.Drawing.Point(123, 37)
        Me.txt_Version.Name = "txt_Version"
        Me.txt_Version.ReadOnly = True
        Me.txt_Version.Size = New System.Drawing.Size(427, 22)
        Me.txt_Version.TabIndex = 1
        Me.txt_Version.Text = "                     "
        '
        'lbl_Version_lbl
        '
        Me.lbl_Version_lbl.AutoSize = True
        Me.lbl_Version_lbl.Location = New System.Drawing.Point(36, 37)
        Me.lbl_Version_lbl.Name = "lbl_Version_lbl"
        Me.lbl_Version_lbl.Size = New System.Drawing.Size(80, 22)
        Me.lbl_Version_lbl.TabIndex = 0
        Me.lbl_Version_lbl.Text = "Version:"
        '
        'AboutBox
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btn_OK
        Me.ClientSize = New System.Drawing.Size(594, 493)
        Me.Controls.Add(Me.grop_AppInf)
        Me.Controls.Add(Me.grop_AuthorInf)
        Me.Controls.Add(Me.btn_OK)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AboutBox"
        Me.Padding = New System.Windows.Forms.Padding(9)
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.grop_AuthorInf.ResumeLayout(False)
        Me.grop_AuthorInf.PerformLayout()
        Me.grop_AppInf.ResumeLayout(False)
        Me.grop_AppInf.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txt_AuthorInf As System.Windows.Forms.TextBox
    Friend WithEvents grop_AuthorInf As System.Windows.Forms.GroupBox
    Friend WithEvents grop_AppInf As System.Windows.Forms.GroupBox
    Friend WithEvents lbl_Release As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Release_lbl As System.Windows.Forms.Label
    Friend WithEvents txt_Version As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Version_lbl As System.Windows.Forms.Label
    Friend WithEvents txt_Description As System.Windows.Forms.TextBox
    Friend WithEvents lbl_Description_lbl As System.Windows.Forms.Label

End Class
