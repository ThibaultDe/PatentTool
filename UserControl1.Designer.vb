﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
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
        Me.SuspendLayout()
        '
        'Replace
        '
        Me.Replace.Location = New System.Drawing.Point(86, 725)
        Me.Replace.Name = "Replace"
        Me.Replace.Size = New System.Drawing.Size(163, 23)
        Me.Replace.TabIndex = 2
        Me.Replace.Text = "Replace"
        Me.Replace.UseVisualStyleBackColor = True
        '
        'FindRefs
        '
        Me.FindRefs.Location = New System.Drawing.Point(3, 3)
        Me.FindRefs.Name = "FindRefs"
        Me.FindRefs.Size = New System.Drawing.Size(123, 23)
        Me.FindRefs.TabIndex = 3
        Me.FindRefs.Text = "Find References"
        Me.FindRefs.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Number, Me.Text})
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(3, 32)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(327, 687)
        Me.ListView1.TabIndex = 4
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'MyUserControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.FindRefs)
        Me.Controls.Add(Me.Replace)
        Me.Name = "MyUserControl"
        Me.Size = New System.Drawing.Size(336, 767)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Replace As Windows.Forms.Button
    Friend WithEvents FindRefs As Windows.Forms.Button
    Friend WithEvents ListView1 As Windows.Forms.ListView
    Friend WithEvents Number As Windows.Forms.ColumnHeader
    Friend WithEvents Text As Windows.Forms.ColumnHeader
End Class