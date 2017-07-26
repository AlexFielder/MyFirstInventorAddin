<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class iPropertiesForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(iPropertiesForm))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.btUpdateAll = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btDefer = New System.Windows.Forms.Button()
        Me.tbMass = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.tbDensity = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.btITEM = New System.Windows.Forms.Button()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btShtMaterial = New System.Windows.Forms.Button()
        Me.btShtScale = New System.Windows.Forms.Button()
        Me.btDiaEng = New System.Windows.Forms.Button()
        Me.btDegEng = New System.Windows.Forms.Button()
        Me.btExpStp = New System.Windows.Forms.Button()
        Me.btExpStl = New System.Windows.Forms.Button()
        Me.btExpPdf = New System.Windows.Forms.Button()
        Me.FileLocation = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.tbDrawnBy = New System.Windows.Forms.TextBox()
        Me.tbEngineer = New System.Windows.Forms.TextBox()
        Me.tbStockNumber = New System.Windows.Forms.TextBox()
        Me.tbPartNumber = New System.Windows.Forms.TextBox()
        Me.tbDescription = New System.Windows.Forms.TextBox()
        Me.btExpSat = New System.Windows.Forms.Button()
        Me.ModelFileLocation = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.tbRevNo = New System.Windows.Forms.TextBox()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.btReNum = New System.Windows.Forms.Button()
        Me.btCopyPN = New System.Windows.Forms.Button()
        Me.btDegDes = New System.Windows.Forms.Button()
        Me.btDiaDes = New System.Windows.Forms.Button()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(46, 13)
        Me.Label1.TabIndex = 100
        Me.Label1.Text = "Part No."
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 37)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 101
        Me.Label2.Text = "Desc."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 13)
        Me.Label3.TabIndex = 102
        Me.Label3.Text = "Stock No."
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(3, 83)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(29, 13)
        Me.Label4.TabIndex = 104
        Me.Label4.Text = "Eng."
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(3, 106)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(32, 13)
        Me.Label5.TabIndex = 105
        Me.Label5.Text = "Mass"
        '
        'btUpdateAll
        '
        Me.btUpdateAll.AccessibleDescription = " up"
        Me.btUpdateAll.Location = New System.Drawing.Point(201, 124)
        Me.btUpdateAll.Name = "btUpdateAll"
        Me.btUpdateAll.Size = New System.Drawing.Size(78, 23)
        Me.btUpdateAll.TabIndex = 200
        Me.btUpdateAll.Text = "Update iProp"
        Me.btUpdateAll.UseVisualStyleBackColor = True
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(3, 174)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(67, 13)
        Me.Label6.TabIndex = 106
        Me.Label6.Text = "Drawn Date:"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(3, 60)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(118, 13)
        Me.Label8.TabIndex = 107
        Me.Label8.Text = "Defer Drawing Updates"
        '
        'btDefer
        '
        Me.btDefer.Location = New System.Drawing.Point(191, 55)
        Me.btDefer.Name = "btDefer"
        Me.btDefer.Size = New System.Drawing.Size(87, 23)
        Me.btDefer.TabIndex = 201
        Me.btDefer.Text = "Defer Switch"
        Me.btDefer.UseVisualStyleBackColor = True
        '
        'tbMass
        '
        Me.tbMass.Location = New System.Drawing.Point(73, 102)
        Me.tbMass.Name = "tbMass"
        Me.tbMass.ReadOnly = True
        Me.tbMass.Size = New System.Drawing.Size(116, 20)
        Me.tbMass.TabIndex = 5
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(3, 129)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(42, 13)
        Me.Label9.TabIndex = 108
        Me.Label9.Text = "Density"
        '
        'tbDensity
        '
        Me.tbDensity.Location = New System.Drawing.Point(73, 125)
        Me.tbDensity.Name = "tbDensity"
        Me.tbDensity.ReadOnly = True
        Me.tbDensity.Size = New System.Drawing.Size(115, 20)
        Me.tbDensity.TabIndex = 8
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(3, 151)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(30, 13)
        Me.Label11.TabIndex = 109
        Me.Label11.Text = "Matl:"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Location = New System.Drawing.Point(31, 151)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(44, 13)
        Me.Label12.TabIndex = 9
        Me.Label12.Tag = ""
        Me.Label12.Text = "Material"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker1.Location = New System.Drawing.Point(67, 170)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(116, 20)
        Me.DateTimePicker1.TabIndex = 301
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(3, 106)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(53, 13)
        Me.Label7.TabIndex = 110
        Me.Label7.Text = "Drawn By"
        '
        'btITEM
        '
        Me.btITEM.Location = New System.Drawing.Point(228, 147)
        Me.btITEM.Name = "btITEM"
        Me.btITEM.Size = New System.Drawing.Size(50, 23)
        Me.btITEM.TabIndex = 202
        Me.btITEM.Text = "#ITEM"
        Me.btITEM.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(212, 173)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(67, 13)
        Me.Label10.TabIndex = 302
        Me.Label10.Text = "Checked in?"
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.ImageLocation = ""
        Me.PictureBox2.Location = New System.Drawing.Point(199, 173)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(12, 12)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 304
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.Control
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(199, 173)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(12, 12)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 303
        Me.PictureBox1.TabStop = False
        '
        'btShtMaterial
        '
        Me.btShtMaterial.Location = New System.Drawing.Point(3, 124)
        Me.btShtMaterial.Name = "btShtMaterial"
        Me.btShtMaterial.Size = New System.Drawing.Size(75, 23)
        Me.btShtMaterial.TabIndex = 305
        Me.btShtMaterial.Text = "Sht Material"
        Me.btShtMaterial.UseVisualStyleBackColor = True
        '
        'btShtScale
        '
        Me.btShtScale.Location = New System.Drawing.Point(103, 124)
        Me.btShtScale.Name = "btShtScale"
        Me.btShtScale.Size = New System.Drawing.Size(75, 23)
        Me.btShtScale.TabIndex = 306
        Me.btShtScale.Text = "Sht Scale"
        Me.btShtScale.UseVisualStyleBackColor = True
        '
        'btDiaEng
        '
        Me.btDiaEng.Location = New System.Drawing.Point(35, 78)
        Me.btDiaEng.Name = "btDiaEng"
        Me.btDiaEng.Size = New System.Drawing.Size(16, 23)
        Me.btDiaEng.TabIndex = 307
        Me.btDiaEng.Text = "Ø"
        Me.btDiaEng.UseVisualStyleBackColor = True
        '
        'btDegEng
        '
        Me.btDegEng.Location = New System.Drawing.Point(52, 78)
        Me.btDegEng.Name = "btDegEng"
        Me.btDegEng.Size = New System.Drawing.Size(16, 23)
        Me.btDegEng.TabIndex = 308
        Me.btDegEng.Text = "°"
        Me.btDegEng.UseVisualStyleBackColor = True
        '
        'btExpStp
        '
        Me.btExpStp.Location = New System.Drawing.Point(3, 190)
        Me.btExpStp.Name = "btExpStp"
        Me.btExpStp.Size = New System.Drawing.Size(60, 23)
        Me.btExpStp.TabIndex = 311
        Me.btExpStp.Text = "Exp STP"
        Me.btExpStp.UseVisualStyleBackColor = True
        '
        'btExpStl
        '
        Me.btExpStl.Location = New System.Drawing.Point(147, 190)
        Me.btExpStl.Name = "btExpStl"
        Me.btExpStl.Size = New System.Drawing.Size(60, 23)
        Me.btExpStl.TabIndex = 312
        Me.btExpStl.Text = "Exp STL"
        Me.btExpStl.UseVisualStyleBackColor = True
        '
        'btExpPdf
        '
        Me.btExpPdf.Location = New System.Drawing.Point(218, 190)
        Me.btExpPdf.Name = "btExpPdf"
        Me.btExpPdf.Size = New System.Drawing.Size(60, 23)
        Me.btExpPdf.TabIndex = 313
        Me.btExpPdf.Text = "Exp PDF"
        Me.btExpPdf.UseVisualStyleBackColor = True
        '
        'FileLocation
        '
        Me.FileLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!)
        Me.FileLocation.Location = New System.Drawing.Point(3, 216)
        Me.FileLocation.Name = "FileLocation"
        Me.FileLocation.Size = New System.Drawing.Size(273, 35)
        Me.FileLocation.TabIndex = 314
        Me.FileLocation.Text = "File location"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Arial", 6.0!)
        Me.Label13.Location = New System.Drawing.Point(82, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(101, 10)
        Me.Label13.TabIndex = 315
        Me.Label13.Text = "iProperties Controller v10.01"
        '
        'tbDrawnBy
        '
        Me.tbDrawnBy.Location = New System.Drawing.Point(73, 102)
        Me.tbDrawnBy.Name = "tbDrawnBy"
        Me.tbDrawnBy.Size = New System.Drawing.Size(116, 20)
        Me.tbDrawnBy.TabIndex = 7
        '
        'tbEngineer
        '
        Me.tbEngineer.Location = New System.Drawing.Point(73, 79)
        Me.tbEngineer.Name = "tbEngineer"
        Me.tbEngineer.Size = New System.Drawing.Size(206, 20)
        Me.tbEngineer.TabIndex = 319
        '
        'tbStockNumber
        '
        Me.tbStockNumber.Location = New System.Drawing.Point(73, 56)
        Me.tbStockNumber.Name = "tbStockNumber"
        Me.tbStockNumber.Size = New System.Drawing.Size(116, 20)
        Me.tbStockNumber.TabIndex = 320
        '
        'tbPartNumber
        '
        Me.tbPartNumber.Location = New System.Drawing.Point(73, 11)
        Me.tbPartNumber.Name = "tbPartNumber"
        Me.tbPartNumber.Size = New System.Drawing.Size(116, 20)
        Me.tbPartNumber.TabIndex = 321
        '
        'tbDescription
        '
        Me.tbDescription.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!)
        Me.tbDescription.Location = New System.Drawing.Point(73, 34)
        Me.tbDescription.Name = "tbDescription"
        Me.tbDescription.Size = New System.Drawing.Size(205, 19)
        Me.tbDescription.TabIndex = 322
        '
        'btExpSat
        '
        Me.btExpSat.Location = New System.Drawing.Point(74, 190)
        Me.btExpSat.Name = "btExpSat"
        Me.btExpSat.Size = New System.Drawing.Size(60, 23)
        Me.btExpSat.TabIndex = 323
        Me.btExpSat.Text = "Exp SAT"
        Me.btExpSat.UseVisualStyleBackColor = True
        '
        'ModelFileLocation
        '
        Me.ModelFileLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.5!)
        Me.ModelFileLocation.Location = New System.Drawing.Point(3, 251)
        Me.ModelFileLocation.Name = "ModelFileLocation"
        Me.ModelFileLocation.Size = New System.Drawing.Size(273, 35)
        Me.ModelFileLocation.TabIndex = 324
        Me.ModelFileLocation.Text = "Model File location"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Location = New System.Drawing.Point(195, 106)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(50, 13)
        Me.Label14.TabIndex = 325
        Me.Label14.Text = "Rev. No."
        '
        'tbRevNo
        '
        Me.tbRevNo.Location = New System.Drawing.Point(246, 102)
        Me.tbRevNo.Name = "tbRevNo"
        Me.tbRevNo.Size = New System.Drawing.Size(32, 20)
        Me.tbRevNo.TabIndex = 326
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'btReNum
        '
        Me.btReNum.Location = New System.Drawing.Point(171, 147)
        Me.btReNum.Name = "btReNum"
        Me.btReNum.Size = New System.Drawing.Size(50, 23)
        Me.btReNum.TabIndex = 327
        Me.btReNum.Text = "BoM #"
        Me.btReNum.UseVisualStyleBackColor = True
        '
        'btCopyPN
        '
        Me.btCopyPN.Location = New System.Drawing.Point(192, 55)
        Me.btCopyPN.Name = "btCopyPN"
        Me.btCopyPN.Size = New System.Drawing.Size(87, 23)
        Me.btCopyPN.TabIndex = 328
        Me.btCopyPN.Text = "Copy Part No."
        Me.btCopyPN.UseVisualStyleBackColor = True
        '
        'btDegDes
        '
        Me.btDegDes.Location = New System.Drawing.Point(52, 32)
        Me.btDegDes.Name = "btDegDes"
        Me.btDegDes.Size = New System.Drawing.Size(16, 23)
        Me.btDegDes.TabIndex = 330
        Me.btDegDes.Text = "°"
        Me.btDegDes.UseVisualStyleBackColor = True
        '
        'btDiaDes
        '
        Me.btDiaDes.Location = New System.Drawing.Point(35, 32)
        Me.btDiaDes.Name = "btDiaDes"
        Me.btDiaDes.Size = New System.Drawing.Size(16, 23)
        Me.btDiaDes.TabIndex = 329
        Me.btDiaDes.Text = "Ø"
        Me.btDiaDes.UseVisualStyleBackColor = True
        '
        'iPropertiesForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(282, 296)
        Me.Controls.Add(Me.btDegDes)
        Me.Controls.Add(Me.btDiaDes)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.btCopyPN)
        Me.Controls.Add(Me.btShtMaterial)
        Me.Controls.Add(Me.btReNum)
        Me.Controls.Add(Me.tbRevNo)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.ModelFileLocation)
        Me.Controls.Add(Me.btExpSat)
        Me.Controls.Add(Me.tbDescription)
        Me.Controls.Add(Me.tbPartNumber)
        Me.Controls.Add(Me.btITEM)
        Me.Controls.Add(Me.FileLocation)
        Me.Controls.Add(Me.btExpPdf)
        Me.Controls.Add(Me.btExpStl)
        Me.Controls.Add(Me.btExpStp)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.btUpdateAll)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btShtScale)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.tbEngineer)
        Me.Controls.Add(Me.btDefer)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tbStockNumber)
        Me.Controls.Add(Me.tbDensity)
        Me.Controls.Add(Me.btDegEng)
        Me.Controls.Add(Me.btDiaEng)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.tbDrawnBy)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.tbMass)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "iPropertiesForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "iPropertiesForm"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents btUpdateAll As Windows.Forms.Button
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents btDefer As Windows.Forms.Button
    Friend WithEvents tbMass As Windows.Forms.TextBox
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents tbDensity As Windows.Forms.TextBox
    Friend WithEvents Label11 As Windows.Forms.Label
    Friend WithEvents Label12 As Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As Windows.Forms.DateTimePicker
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents btITEM As Windows.Forms.Button
    Friend WithEvents Label10 As Windows.Forms.Label
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents btShtMaterial As Windows.Forms.Button
    Friend WithEvents btShtScale As Windows.Forms.Button
    Friend WithEvents btDiaEng As Windows.Forms.Button
    Friend WithEvents btDegEng As Windows.Forms.Button
    Friend WithEvents btExpStp As Windows.Forms.Button
    Friend WithEvents btExpStl As Windows.Forms.Button
    Friend WithEvents btExpPdf As Windows.Forms.Button
    Friend WithEvents FileLocation As Windows.Forms.Label
    Friend WithEvents Label13 As Windows.Forms.Label
    Friend WithEvents tbDrawnBy As Windows.Forms.TextBox
    Friend WithEvents tbEngineer As Windows.Forms.TextBox
    Friend WithEvents tbStockNumber As Windows.Forms.TextBox
    Friend WithEvents tbPartNumber As Windows.Forms.TextBox
    Friend WithEvents tbDescription As Windows.Forms.TextBox
    Friend WithEvents btExpSat As Windows.Forms.Button
    Friend WithEvents ModelFileLocation As Windows.Forms.Label
    Friend WithEvents Label14 As Windows.Forms.Label
    Friend WithEvents tbRevNo As Windows.Forms.TextBox
    Friend WithEvents ErrorProvider1 As Windows.Forms.ErrorProvider
    Friend WithEvents btReNum As Windows.Forms.Button
    Friend WithEvents btCopyPN As Windows.Forms.Button
    Friend WithEvents btDegDes As Windows.Forms.Button
    Friend WithEvents btDiaDes As Windows.Forms.Button
End Class
