<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_UDL
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.OK_Button = New System.Windows.Forms.Button
        Me.Cancel_Button = New System.Windows.Forms.Button
        Me.btn_Browse = New System.Windows.Forms.Button
        Me.txt_Path = New System.Windows.Forms.TextBox
        Me.txt_PW = New System.Windows.Forms.TextBox
        Me.chk_ShowPW_Chars = New System.Windows.Forms.CheckBox
        Me.lbl_DB_Path = New System.Windows.Forms.Label
        Me.lbl_Password = New System.Windows.Forms.Label
        Me.grp_Path_PW = New System.Windows.Forms.GroupBox
        Me.chk_BlankPW = New System.Windows.Forms.CheckBox
        Me.btn_TestCon = New System.Windows.Forms.Button
        Me.TableLayoutPanel1.SuspendLayout()
        Me.grp_Path_PW.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.68493!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 49.31507!))
        Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(336, 137)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OK_Button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.OK_Button.Location = New System.Drawing.Point(3, 3)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 0
        Me.OK_Button.Text = "OK"
        '
        'Cancel_Button
        '
        Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
        Me.Cancel_Button.Name = "Cancel_Button"
        Me.Cancel_Button.Size = New System.Drawing.Size(66, 23)
        Me.Cancel_Button.TabIndex = 1
        Me.Cancel_Button.Text = "Cancel"
        '
        'btn_Browse
        '
        Me.btn_Browse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Browse.Location = New System.Drawing.Point(433, 18)
        Me.btn_Browse.Name = "btn_Browse"
        Me.btn_Browse.Size = New System.Drawing.Size(28, 24)
        Me.btn_Browse.TabIndex = 1
        Me.btn_Browse.Text = "..."
        Me.btn_Browse.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_Browse.UseVisualStyleBackColor = True
        '
        'txt_Path
        '
        Me.txt_Path.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_Path.Location = New System.Drawing.Point(72, 19)
        Me.txt_Path.Name = "txt_Path"
        Me.txt_Path.Size = New System.Drawing.Size(356, 20)
        Me.txt_Path.TabIndex = 2
        '
        'txt_PW
        '
        Me.txt_PW.Location = New System.Drawing.Point(72, 62)
        Me.txt_PW.Name = "txt_PW"
        Me.txt_PW.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txt_PW.Size = New System.Drawing.Size(196, 20)
        Me.txt_PW.TabIndex = 3
        '
        'chk_ShowPW_Chars
        '
        Me.chk_ShowPW_Chars.AutoSize = True
        Me.chk_ShowPW_Chars.Location = New System.Drawing.Point(274, 55)
        Me.chk_ShowPW_Chars.Name = "chk_ShowPW_Chars"
        Me.chk_ShowPW_Chars.Size = New System.Drawing.Size(101, 17)
        Me.chk_ShowPW_Chars.TabIndex = 4
        Me.chk_ShowPW_Chars.Text = "Show Password"
        Me.chk_ShowPW_Chars.UseVisualStyleBackColor = True
        '
        'lbl_DB_Path
        '
        Me.lbl_DB_Path.AutoSize = True
        Me.lbl_DB_Path.Location = New System.Drawing.Point(4, 22)
        Me.lbl_DB_Path.Name = "lbl_DB_Path"
        Me.lbl_DB_Path.Size = New System.Drawing.Size(56, 13)
        Me.lbl_DB_Path.TabIndex = 5
        Me.lbl_DB_Path.Text = "Data Base"
        '
        'lbl_Password
        '
        Me.lbl_Password.AutoSize = True
        Me.lbl_Password.Location = New System.Drawing.Point(7, 63)
        Me.lbl_Password.Name = "lbl_Password"
        Me.lbl_Password.Size = New System.Drawing.Size(53, 13)
        Me.lbl_Password.TabIndex = 6
        Me.lbl_Password.Text = "Password"
        '
        'grp_Path_PW
        '
        Me.grp_Path_PW.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grp_Path_PW.Controls.Add(Me.chk_BlankPW)
        Me.grp_Path_PW.Controls.Add(Me.txt_Path)
        Me.grp_Path_PW.Controls.Add(Me.lbl_Password)
        Me.grp_Path_PW.Controls.Add(Me.btn_Browse)
        Me.grp_Path_PW.Controls.Add(Me.lbl_DB_Path)
        Me.grp_Path_PW.Controls.Add(Me.txt_PW)
        Me.grp_Path_PW.Controls.Add(Me.chk_ShowPW_Chars)
        Me.grp_Path_PW.Location = New System.Drawing.Point(15, 12)
        Me.grp_Path_PW.Name = "grp_Path_PW"
        Me.grp_Path_PW.Size = New System.Drawing.Size(467, 111)
        Me.grp_Path_PW.TabIndex = 7
        Me.grp_Path_PW.TabStop = False
        '
        'chk_BlankPW
        '
        Me.chk_BlankPW.AutoSize = True
        Me.chk_BlankPW.Location = New System.Drawing.Point(274, 76)
        Me.chk_BlankPW.Name = "chk_BlankPW"
        Me.chk_BlankPW.Size = New System.Drawing.Size(100, 17)
        Me.chk_BlankPW.TabIndex = 7
        Me.chk_BlankPW.Text = "Blank Password"
        Me.chk_BlankPW.UseVisualStyleBackColor = True
        '
        'btn_TestCon
        '
        Me.btn_TestCon.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btn_TestCon.Location = New System.Drawing.Point(15, 135)
        Me.btn_TestCon.Name = "btn_TestCon"
        Me.btn_TestCon.Size = New System.Drawing.Size(96, 28)
        Me.btn_TestCon.TabIndex = 11
        Me.btn_TestCon.Text = "Test Connection"
        Me.btn_TestCon.UseVisualStyleBackColor = True
        '
        'frm_UDL
        '
        Me.AcceptButton = Me.OK_Button
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Cancel_Button
        Me.ClientSize = New System.Drawing.Size(494, 174)
        Me.Controls.Add(Me.btn_TestCon)
        Me.Controls.Add(Me.grp_Path_PW)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(700, 200)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(6, 200)
        Me.Name = "frm_UDL"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Connect"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.grp_Path_PW.ResumeLayout(False)
        Me.grp_Path_PW.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents btn_Browse As System.Windows.Forms.Button
    Friend WithEvents txt_Path As System.Windows.Forms.TextBox
    Friend WithEvents txt_PW As System.Windows.Forms.TextBox
    Friend WithEvents chk_ShowPW_Chars As System.Windows.Forms.CheckBox
    Friend WithEvents lbl_DB_Path As System.Windows.Forms.Label
    Friend WithEvents lbl_Password As System.Windows.Forms.Label
    Friend WithEvents grp_Path_PW As System.Windows.Forms.GroupBox
    Friend WithEvents chk_BlankPW As System.Windows.Forms.CheckBox
    Friend WithEvents btn_TestCon As System.Windows.Forms.Button

End Class
