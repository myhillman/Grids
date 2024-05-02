<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class KMLbyCQ
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
        ListBox1 = New ListBox()
        btnOK = New Button()
        btnCancel = New Button()
        SuspendLayout()
        ' 
        ' ListBox1
        ' 
        ListBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom
        ListBox1.FormattingEnabled = True
        ListBox1.ItemHeight = 15
        ListBox1.Location = New Point(63, 21)
        ListBox1.MultiColumn = True
        ListBox1.Name = "ListBox1"
        ListBox1.Size = New Size(167, 334)
        ListBox1.TabIndex = 0
        ' 
        ' btnOK
        ' 
        btnOK.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        btnOK.DialogResult = DialogResult.OK
        btnOK.Location = New Point(12, 402)
        btnOK.Name = "btnOK"
        btnOK.Size = New Size(131, 32)
        btnOK.TabIndex = 1
        btnOK.Text = "OK"
        btnOK.UseVisualStyleBackColor = True
        ' 
        ' btnCancel
        ' 
        btnCancel.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Location = New Point(149, 402)
        btnCancel.Name = "btnCancel"
        btnCancel.Size = New Size(125, 32)
        btnCancel.TabIndex = 2
        btnCancel.Text = "Cancel"
        btnCancel.UseVisualStyleBackColor = True
        ' 
        ' KMLbyCQ
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(294, 446)
        Controls.Add(btnCancel)
        Controls.Add(btnOK)
        Controls.Add(ListBox1)
        Name = "KMLbyCQ"
        Text = "KMLbyCQ"
        ResumeLayout(False)
    End Sub

    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents btnOK As Button
    Friend WithEvents btnCancel As Button
End Class
