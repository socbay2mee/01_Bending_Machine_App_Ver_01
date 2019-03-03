<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Setup
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
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Button_OK = New System.Windows.Forms.Button()
        Me.Button_Cancel = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Txt_Cut_Angle = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Clock_In = New System.Windows.Forms.RadioButton()
        Me.Anti_In = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Clock_Out = New System.Windows.Forms.RadioButton()
        Me.Anti_Out = New System.Windows.Forms.RadioButton()
        Me.Txt_Thickness = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox6
        '
        Me.GroupBox6.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox6.Controls.Add(Me.Label1)
        Me.GroupBox6.Controls.Add(Me.ComboBox1)
        Me.GroupBox6.Location = New System.Drawing.Point(32, 27)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(399, 93)
        Me.GroupBox6.TabIndex = 51
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Setting"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(47, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 13)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "Type"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(120, 19)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(153, 21)
        Me.ComboBox1.TabIndex = 54
        '
        'Button_OK
        '
        Me.Button_OK.Location = New System.Drawing.Point(535, 540)
        Me.Button_OK.Name = "Button_OK"
        Me.Button_OK.Size = New System.Drawing.Size(75, 23)
        Me.Button_OK.TabIndex = 52
        Me.Button_OK.Text = "OK"
        Me.Button_OK.UseVisualStyleBackColor = True
        '
        'Button_Cancel
        '
        Me.Button_Cancel.Location = New System.Drawing.Point(632, 528)
        Me.Button_Cancel.Name = "Button_Cancel"
        Me.Button_Cancel.Size = New System.Drawing.Size(75, 23)
        Me.Button_Cancel.TabIndex = 53
        Me.Button_Cancel.Text = "Cancel"
        Me.Button_Cancel.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.Txt_Cut_Angle)
        Me.GroupBox1.Location = New System.Drawing.Point(455, 27)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(181, 93)
        Me.GroupBox1.TabIndex = 56
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "V-CUT Angle"
        '
        'Txt_Cut_Angle
        '
        Me.Txt_Cut_Angle.Location = New System.Drawing.Point(15, 38)
        Me.Txt_Cut_Angle.Name = "Txt_Cut_Angle"
        Me.Txt_Cut_Angle.Size = New System.Drawing.Size(100, 20)
        Me.Txt_Cut_Angle.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox2.Controls.Add(Me.GroupBox4)
        Me.GroupBox2.Controls.Add(Me.GroupBox3)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 328)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(604, 185)
        Me.GroupBox2.TabIndex = 57
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "V-CUT Margin Setup"
        '
        'GroupBox4
        '
        Me.GroupBox4.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.Clock_In)
        Me.GroupBox4.Controls.Add(Me.Anti_In)
        Me.GroupBox4.Location = New System.Drawing.Point(15, 102)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(563, 77)
        Me.GroupBox4.TabIndex = 58
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Inside"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(405, 23)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(49, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Direction"
        '
        'Clock_In
        '
        Me.Clock_In.AutoSize = True
        Me.Clock_In.Location = New System.Drawing.Point(470, 19)
        Me.Clock_In.Name = "Clock_In"
        Me.Clock_In.Size = New System.Drawing.Size(73, 17)
        Me.Clock_In.TabIndex = 3
        Me.Clock_In.TabStop = True
        Me.Clock_In.Text = "Clockwise"
        Me.Clock_In.UseVisualStyleBackColor = True
        '
        'Anti_In
        '
        Me.Anti_In.AutoSize = True
        Me.Anti_In.Location = New System.Drawing.Point(470, 54)
        Me.Anti_In.Name = "Anti_In"
        Me.Anti_In.Size = New System.Drawing.Size(93, 17)
        Me.Anti_In.TabIndex = 4
        Me.Anti_In.TabStop = True
        Me.Anti_In.Text = "Anti-clockwise"
        Me.Anti_In.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.Clock_Out)
        Me.GroupBox3.Controls.Add(Me.Anti_Out)
        Me.GroupBox3.Location = New System.Drawing.Point(15, 19)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(563, 77)
        Me.GroupBox3.TabIndex = 57
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Outside"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(405, 23)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Direction"
        '
        'Clock_Out
        '
        Me.Clock_Out.AutoSize = True
        Me.Clock_Out.Location = New System.Drawing.Point(470, 19)
        Me.Clock_Out.Name = "Clock_Out"
        Me.Clock_Out.Size = New System.Drawing.Size(73, 17)
        Me.Clock_Out.TabIndex = 3
        Me.Clock_Out.TabStop = True
        Me.Clock_Out.Text = "Clockwise"
        Me.Clock_Out.UseVisualStyleBackColor = True
        '
        'Anti_Out
        '
        Me.Anti_Out.AutoSize = True
        Me.Anti_Out.Location = New System.Drawing.Point(470, 54)
        Me.Anti_Out.Name = "Anti_Out"
        Me.Anti_Out.Size = New System.Drawing.Size(93, 17)
        Me.Anti_Out.TabIndex = 4
        Me.Anti_Out.TabStop = True
        Me.Anti_Out.Text = "Anti-clockwise"
        Me.Anti_Out.UseVisualStyleBackColor = True
        '
        'Txt_Thickness
        '
        Me.Txt_Thickness.Location = New System.Drawing.Point(86, 26)
        Me.Txt_Thickness.Name = "Txt_Thickness"
        Me.Txt_Thickness.Size = New System.Drawing.Size(100, 20)
        Me.Txt_Thickness.TabIndex = 0
        '
        'GroupBox5
        '
        Me.GroupBox5.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Controls.Add(Me.Txt_Thickness)
        Me.GroupBox5.Location = New System.Drawing.Point(32, 172)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(604, 123)
        Me.GroupBox5.TabIndex = 59
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Material Setup"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 29)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Thickness"
        '
        'Form_Setup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(719, 575)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button_Cancel)
        Me.Controls.Add(Me.Button_OK)
        Me.Controls.Add(Me.GroupBox6)
        Me.Name = "Form_Setup"
        Me.Text = "Form_Setup"
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents Button_OK As System.Windows.Forms.Button
    Friend WithEvents Button_Cancel As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Txt_Cut_Angle As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Anti_Out As System.Windows.Forms.RadioButton
    Friend WithEvents Clock_Out As System.Windows.Forms.RadioButton
    Friend WithEvents Txt_Thickness As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Clock_In As System.Windows.Forms.RadioButton
    Friend WithEvents Anti_In As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
End Class
