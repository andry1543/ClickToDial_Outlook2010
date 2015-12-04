<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Click_to_Dial = New System.Windows.Forms.Button()
        Me.Click_to_Exit = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(142, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Ввведите номер телефона"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(31, 38)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(139, 20)
        Me.TextBox1.TabIndex = 0
        '
        'Click_to_Dial
        '
        Me.Click_to_Dial.Location = New System.Drawing.Point(12, 97)
        Me.Click_to_Dial.Name = "Click_to_Dial"
        Me.Click_to_Dial.Size = New System.Drawing.Size(75, 23)
        Me.Click_to_Dial.TabIndex = 4
        Me.Click_to_Dial.Text = "Позвонить"
        Me.Click_to_Dial.UseVisualStyleBackColor = True
        '
        'Click_to_Exit
        '
        Me.Click_to_Exit.Location = New System.Drawing.Point(114, 97)
        Me.Click_to_Exit.Name = "Click_to_Exit"
        Me.Click_to_Exit.Size = New System.Drawing.Size(75, 23)
        Me.Click_to_Exit.TabIndex = 5
        Me.Click_to_Exit.Text = "Отменить"
        Me.Click_to_Exit.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(201, 132)
        Me.Controls.Add(Me.Click_to_Exit)
        Me.Controls.Add(Me.Click_to_Dial)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Click to Call"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Click_to_Dial As System.Windows.Forms.Button
    Friend WithEvents Click_to_Exit As System.Windows.Forms.Button
End Class
