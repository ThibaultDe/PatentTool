<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MyUserControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.FindRefs = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.Number = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Text = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ReplaceAllRevs = New System.Windows.Forms.Button()
        Me.EnglishVersion = New System.Windows.Forms.CheckBox()
        Me.ContinuousScann = New System.Windows.Forms.CheckBox()
        Me.ReplaceRevs = New System.Windows.Forms.Button()
        Me.ReplaceDesc = New System.Windows.Forms.Button()
        Me.ReplaceAllDesc = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'FindRefs
        '
        Me.FindRefs.Location = New System.Drawing.Point(3, 3)
        Me.FindRefs.Name = "FindRefs"
        Me.FindRefs.Size = New System.Drawing.Size(113, 23)
        Me.FindRefs.TabIndex = 3
        Me.FindRefs.Text = "Refresh"
        Me.FindRefs.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Number, Me.Text})
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(3, 51)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(327, 668)
        Me.ListView1.TabIndex = 4
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'ReplaceAllRevs
        '
        Me.ReplaceAllRevs.Location = New System.Drawing.Point(174, 725)
        Me.ReplaceAllRevs.Name = "ReplaceAllRevs"
        Me.ReplaceAllRevs.Size = New System.Drawing.Size(156, 23)
        Me.ReplaceAllRevs.TabIndex = 5
        Me.ReplaceAllRevs.Text = "Replace All in Revs"
        Me.ReplaceAllRevs.UseVisualStyleBackColor = True
        '
        'EnglishVersion
        '
        Me.EnglishVersion.AutoSize = True
        Me.EnglishVersion.Location = New System.Drawing.Point(232, 7)
        Me.EnglishVersion.Name = "EnglishVersion"
        Me.EnglishVersion.Size = New System.Drawing.Size(98, 17)
        Me.EnglishVersion.TabIndex = 6
        Me.EnglishVersion.Text = "English Version"
        Me.EnglishVersion.UseVisualStyleBackColor = True
        '
        'ContinuousScann
        '
        Me.ContinuousScann.AutoSize = True
        Me.ContinuousScann.Location = New System.Drawing.Point(3, 28)
        Me.ContinuousScann.Name = "ContinuousScann"
        Me.ContinuousScann.Size = New System.Drawing.Size(113, 17)
        Me.ContinuousScann.TabIndex = 7
        Me.ContinuousScann.Text = "Continuous Scann"
        Me.ContinuousScann.UseVisualStyleBackColor = True
        '
        'ReplaceRevs
        '
        Me.ReplaceRevs.Location = New System.Drawing.Point(3, 725)
        Me.ReplaceRevs.Name = "ReplaceRevs"
        Me.ReplaceRevs.Size = New System.Drawing.Size(165, 23)
        Me.ReplaceRevs.TabIndex = 2
        Me.ReplaceRevs.Text = "Replace in Revs"
        Me.ReplaceRevs.UseVisualStyleBackColor = True
        '
        'ReplaceDesc
        '
        Me.ReplaceDesc.Location = New System.Drawing.Point(3, 754)
        Me.ReplaceDesc.Name = "ReplaceDesc"
        Me.ReplaceDesc.Size = New System.Drawing.Size(165, 23)
        Me.ReplaceDesc.TabIndex = 8
        Me.ReplaceDesc.Text = "Replace in Desc"
        Me.ReplaceDesc.UseVisualStyleBackColor = True
        '
        'ReplaceAllDesc
        '
        Me.ReplaceAllDesc.Location = New System.Drawing.Point(174, 754)
        Me.ReplaceAllDesc.Name = "ReplaceAllDesc"
        Me.ReplaceAllDesc.Size = New System.Drawing.Size(156, 23)
        Me.ReplaceAllDesc.TabIndex = 9
        Me.ReplaceAllDesc.Text = "Replace All in Desc"
        Me.ReplaceAllDesc.UseVisualStyleBackColor = True
        '
        'MyUserControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.ReplaceAllDesc)
        Me.Controls.Add(Me.ReplaceDesc)
        Me.Controls.Add(Me.ContinuousScann)
        Me.Controls.Add(Me.EnglishVersion)
        Me.Controls.Add(Me.ReplaceAllRevs)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.FindRefs)
        Me.Controls.Add(Me.ReplaceRevs)
        Me.Name = "MyUserControl"
        Me.Size = New System.Drawing.Size(336, 787)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FindRefs As Windows.Forms.Button
    Friend WithEvents ListView1 As Windows.Forms.ListView
    Friend WithEvents Number As Windows.Forms.ColumnHeader
    Friend WithEvents Text As Windows.Forms.ColumnHeader
    Friend WithEvents ReplaceAllRevs As Windows.Forms.Button
    Friend WithEvents EnglishVersion As Windows.Forms.CheckBox
    Friend WithEvents ContinuousScann As Windows.Forms.CheckBox
    Friend WithEvents ReplaceRevs As Windows.Forms.Button
    Friend WithEvents ReplaceDesc As Windows.Forms.Button
    Friend WithEvents ReplaceAllDesc As Windows.Forms.Button
End Class
