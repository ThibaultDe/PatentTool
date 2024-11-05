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
        Me.Replace = New System.Windows.Forms.Button()
        Me.FindRefs = New System.Windows.Forms.Button()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.Number = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Text = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Replace_All = New System.Windows.Forms.Button()
        Me.EnglishVersion = New System.Windows.Forms.CheckBox()
        Me.ContinuousScann = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Replace
        '
        Me.Replace.Location = New System.Drawing.Point(3, 725)
        Me.Replace.Name = "Replace"
        Me.Replace.Size = New System.Drawing.Size(165, 23)
        Me.Replace.TabIndex = 2
        Me.Replace.Text = "Replace"
        Me.Replace.UseVisualStyleBackColor = True
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
        'Replace_All
        '
        Me.Replace_All.Location = New System.Drawing.Point(174, 725)
        Me.Replace_All.Name = "Replace_All"
        Me.Replace_All.Size = New System.Drawing.Size(156, 23)
        Me.Replace_All.TabIndex = 5
        Me.Replace_All.Text = "Replace All"
        Me.Replace_All.UseVisualStyleBackColor = True
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
        'MyUserControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.ContinuousScann)
        Me.Controls.Add(Me.EnglishVersion)
        Me.Controls.Add(Me.Replace_All)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.FindRefs)
        Me.Controls.Add(Me.Replace)
        Me.Name = "MyUserControl"
        Me.Size = New System.Drawing.Size(336, 767)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Replace As Windows.Forms.Button
    Friend WithEvents FindRefs As Windows.Forms.Button
    Friend WithEvents ListView1 As Windows.Forms.ListView
    Friend WithEvents Number As Windows.Forms.ColumnHeader
    Friend WithEvents Text As Windows.Forms.ColumnHeader
    Friend WithEvents Replace_All As Windows.Forms.Button
    Friend WithEvents EnglishVersion As Windows.Forms.CheckBox
    Friend WithEvents ContinuousScann As Windows.Forms.CheckBox
End Class
