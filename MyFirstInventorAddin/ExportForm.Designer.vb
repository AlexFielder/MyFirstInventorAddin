<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExportForm
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
        Me.btExpPipes = New System.Windows.Forms.Button()
        Me.btExpSat = New System.Windows.Forms.Button()
        Me.btExpPdf = New System.Windows.Forms.Button()
        Me.btExpStl = New System.Windows.Forms.Button()
        Me.btExpStp = New System.Windows.Forms.Button()
        Me.btExpDWF = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btExpPipes
        '
        Me.btExpPipes.Location = New System.Drawing.Point(5, 142)
        Me.btExpPipes.Name = "btExpPipes"
        Me.btExpPipes.Size = New System.Drawing.Size(80, 22)
        Me.btExpPipes.TabIndex = 336
        Me.btExpPipes.Text = "Export Pipes"
        Me.btExpPipes.UseVisualStyleBackColor = True
        '
        'btExpSat
        '
        Me.btExpSat.Location = New System.Drawing.Point(5, 38)
        Me.btExpSat.Name = "btExpSat"
        Me.btExpSat.Size = New System.Drawing.Size(80, 22)
        Me.btExpSat.TabIndex = 335
        Me.btExpSat.Text = "Export SAT"
        Me.btExpSat.UseVisualStyleBackColor = True
        '
        'btExpPdf
        '
        Me.btExpPdf.Location = New System.Drawing.Point(5, 90)
        Me.btExpPdf.Name = "btExpPdf"
        Me.btExpPdf.Size = New System.Drawing.Size(80, 22)
        Me.btExpPdf.TabIndex = 334
        Me.btExpPdf.Text = "Export PDF"
        Me.btExpPdf.UseVisualStyleBackColor = True
        '
        'btExpStl
        '
        Me.btExpStl.Location = New System.Drawing.Point(5, 64)
        Me.btExpStl.Name = "btExpStl"
        Me.btExpStl.Size = New System.Drawing.Size(80, 22)
        Me.btExpStl.TabIndex = 333
        Me.btExpStl.Text = "Export STL"
        Me.btExpStl.UseVisualStyleBackColor = True
        '
        'btExpStp
        '
        Me.btExpStp.Location = New System.Drawing.Point(5, 12)
        Me.btExpStp.Name = "btExpStp"
        Me.btExpStp.Size = New System.Drawing.Size(80, 22)
        Me.btExpStp.TabIndex = 332
        Me.btExpStp.Text = "Export STP"
        Me.btExpStp.UseVisualStyleBackColor = True
        '
        'btExpDWF
        '
        Me.btExpDWF.Location = New System.Drawing.Point(5, 116)
        Me.btExpDWF.Name = "btExpDWF"
        Me.btExpDWF.Size = New System.Drawing.Size(80, 22)
        Me.btExpDWF.TabIndex = 337
        Me.btExpDWF.Text = "Export DWF"
        Me.btExpDWF.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(91, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 13)
        Me.Label1.TabIndex = 338
        Me.Label1.Text = "Export to stp file"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(91, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 13)
        Me.Label2.TabIndex = 339
        Me.Label2.Text = "Export to sat file"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(91, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(106, 13)
        Me.Label3.TabIndex = 340
        Me.Label3.Text = "Export for 3D printing"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(91, 88)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 29)
        Me.Label4.TabIndex = 341
        Me.Label4.Text = "Export to pdf for drawing or 3D pdf for model"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(91, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(85, 13)
        Me.Label5.TabIndex = 342
        Me.Label5.Text = "Export to dwf file"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(91, 140)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(138, 29)
        Me.Label6.TabIndex = 343
        Me.Label6.Text = "Export assembly parts/sub assys to individual stp files"
        '
        'ExportForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(225, 178)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btExpDWF)
        Me.Controls.Add(Me.btExpPipes)
        Me.Controls.Add(Me.btExpSat)
        Me.Controls.Add(Me.btExpPdf)
        Me.Controls.Add(Me.btExpStl)
        Me.Controls.Add(Me.btExpStp)
        Me.Name = "ExportForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "ExportOptions"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btExpPipes As Windows.Forms.Button
    Friend WithEvents btExpSat As Windows.Forms.Button
    Friend WithEvents btExpPdf As Windows.Forms.Button
    Friend WithEvents btExpStl As Windows.Forms.Button
    Friend WithEvents btExpStp As Windows.Forms.Button
    Friend WithEvents btExpDWF As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label6 As Windows.Forms.Label
End Class
