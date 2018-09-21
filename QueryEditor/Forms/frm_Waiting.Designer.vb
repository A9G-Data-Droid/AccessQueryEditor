<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Working
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Working))
        Me.lbl_Task = New System.Windows.Forms.Label
        Me.lbl_Wait_lbl = New System.Windows.Forms.Label
        Me.pic_Ico = New System.Windows.Forms.PictureBox
        Me.lbl_Task_lbl = New System.Windows.Forms.Label
        CType(Me.pic_Ico, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lbl_Task
        '
        Me.lbl_Task.AutoSize = True
        Me.lbl_Task.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Task.Location = New System.Drawing.Point(54, 69)
        Me.lbl_Task.Name = "lbl_Task"
        Me.lbl_Task.Size = New System.Drawing.Size(29, 13)
        Me.lbl_Task.TabIndex = 0
        Me.lbl_Task.Text = "Task"
        '
        'lbl_Wait_lbl
        '
        Me.lbl_Wait_lbl.AutoSize = True
        Me.lbl_Wait_lbl.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Wait_lbl.Location = New System.Drawing.Point(12, 102)
        Me.lbl_Wait_lbl.Name = "lbl_Wait_lbl"
        Me.lbl_Wait_lbl.Size = New System.Drawing.Size(119, 13)
        Me.lbl_Wait_lbl.TabIndex = 1
        Me.lbl_Wait_lbl.Text = "Please Wait.............."
        '
        'pic_Ico
        '
        Me.pic_Ico.Image = CType(resources.GetObject("pic_Ico.Image"), System.Drawing.Image)
        Me.pic_Ico.Location = New System.Drawing.Point(12, 11)
        Me.pic_Ico.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.pic_Ico.Name = "pic_Ico"
        Me.pic_Ico.Size = New System.Drawing.Size(297, 11)
        Me.pic_Ico.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pic_Ico.TabIndex = 2
        Me.pic_Ico.TabStop = False
        Me.pic_Ico.WaitOnLoad = True
        '
        'lbl_Task_lbl
        '
        Me.lbl_Task_lbl.AutoSize = True
        Me.lbl_Task_lbl.Location = New System.Drawing.Point(12, 69)
        Me.lbl_Task_lbl.Name = "lbl_Task_lbl"
        Me.lbl_Task_lbl.Size = New System.Drawing.Size(36, 13)
        Me.lbl_Task_lbl.TabIndex = 3
        Me.lbl_Task_lbl.Text = "Task :"
        '
        'frm_Working
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(321, 124)
        Me.ControlBox = False
        Me.Controls.Add(Me.lbl_Task_lbl)
        Me.Controls.Add(Me.pic_Ico)
        Me.Controls.Add(Me.lbl_Wait_lbl)
        Me.Controls.Add(Me.lbl_Task)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "frm_Working"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Working"
        Me.TopMost = True
        CType(Me.pic_Ico, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbl_Task As System.Windows.Forms.Label
    Friend WithEvents lbl_Wait_lbl As System.Windows.Forms.Label
    Friend WithEvents pic_Ico As System.Windows.Forms.PictureBox
    Friend WithEvents lbl_Task_lbl As System.Windows.Forms.Label
End Class
