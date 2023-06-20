<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Dim resources As ComponentModel.ComponentResourceManager = New ComponentModel.ComponentResourceManager(GetType(Form1))
        ButtonExcelToKML = New Button()
        ButtonKMZToKML = New Button()
        ButtonCombineKML = New Button()
        GroupBoxExcelToKml = New GroupBox()
        Label1 = New Label()
        GroupBoxKMZtoKML = New GroupBox()
        Label = New Label()
        GroupBoxGenerateMasterKMLFile = New GroupBox()
        Label3 = New Label()
        PictureBox1 = New PictureBox()
        PictureBox2 = New PictureBox()
        GroupBoxExcelToKml.SuspendLayout()
        GroupBoxKMZtoKML.SuspendLayout()
        GroupBoxGenerateMasterKMLFile.SuspendLayout()
        CType(PictureBox1, ComponentModel.ISupportInitialize).BeginInit()
        CType(PictureBox2, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' ButtonExcelToKML
        ' 
        ButtonExcelToKML.Location = New Point(124, 80)
        ButtonExcelToKML.Name = "ButtonExcelToKML"
        ButtonExcelToKML.Size = New Size(84, 43)
        ButtonExcelToKML.TabIndex = 0
        ButtonExcelToKML.Text = "Excel to KML"
        ButtonExcelToKML.UseVisualStyleBackColor = True
        ' 
        ' ButtonKMZToKML
        ' 
        ButtonKMZToKML.Location = New Point(121, 80)
        ButtonKMZToKML.Name = "ButtonKMZToKML"
        ButtonKMZToKML.Size = New Size(84, 43)
        ButtonKMZToKML.TabIndex = 1
        ButtonKMZToKML.Text = "KMZ to KML"
        ButtonKMZToKML.UseVisualStyleBackColor = True
        ' 
        ' ButtonCombineKML
        ' 
        ButtonCombineKML.Location = New Point(121, 80)
        ButtonCombineKML.Name = "ButtonCombineKML"
        ButtonCombineKML.Size = New Size(84, 43)
        ButtonCombineKML.TabIndex = 2
        ButtonCombineKML.Text = "Mongoose KML"
        ButtonCombineKML.UseVisualStyleBackColor = True
        ' 
        ' GroupBoxExcelToKml
        ' 
        GroupBoxExcelToKml.Controls.Add(Label1)
        GroupBoxExcelToKml.Controls.Add(ButtonExcelToKML)
        GroupBoxExcelToKml.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        GroupBoxExcelToKml.ForeColor = Color.Navy
        GroupBoxExcelToKml.Location = New Point(15, 12)
        GroupBoxExcelToKml.Name = "GroupBoxExcelToKml"
        GroupBoxExcelToKml.Size = New Size(368, 129)
        GroupBoxExcelToKml.TabIndex = 3
        GroupBoxExcelToKml.TabStop = False
        GroupBoxExcelToKml.Text = "Generate KML from Excel"
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(6, 34)
        Label1.Name = "Label1"
        Label1.Size = New Size(321, 15)
        Label1.TabIndex = 1
        Label1.Text = "Select the Mongoose Excel File, A KML will be generated"
        ' 
        ' GroupBoxKMZtoKML
        ' 
        GroupBoxKMZtoKML.Controls.Add(Label)
        GroupBoxKMZtoKML.Controls.Add(ButtonKMZToKML)
        GroupBoxKMZtoKML.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        GroupBoxKMZtoKML.ForeColor = Color.Navy
        GroupBoxKMZtoKML.Location = New Point(18, 147)
        GroupBoxKMZtoKML.Name = "GroupBoxKMZtoKML"
        GroupBoxKMZtoKML.Size = New Size(365, 129)
        GroupBoxKMZtoKML.TabIndex = 4
        GroupBoxKMZtoKML.TabStop = False
        GroupBoxKMZtoKML.Text = "Convert KMZ file to KML files"
        ' 
        ' Label
        ' 
        Label.AutoSize = True
        Label.Location = New Point(16, 34)
        Label.Name = "Label"
        Label.Size = New Size(302, 30)
        Label.TabIndex = 2
        Label.Text = "Select the Parent Folder containing the KMZ files," & vbCrLf & "the KMZ files will be unzipped and extract their KML." & vbCrLf
        ' 
        ' GroupBoxGenerateMasterKMLFile
        ' 
        GroupBoxGenerateMasterKMLFile.Controls.Add(Label3)
        GroupBoxGenerateMasterKMLFile.Controls.Add(ButtonCombineKML)
        GroupBoxGenerateMasterKMLFile.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point)
        GroupBoxGenerateMasterKMLFile.ForeColor = Color.Navy
        GroupBoxGenerateMasterKMLFile.Location = New Point(18, 282)
        GroupBoxGenerateMasterKMLFile.Name = "GroupBoxGenerateMasterKMLFile"
        GroupBoxGenerateMasterKMLFile.Size = New Size(365, 129)
        GroupBoxGenerateMasterKMLFile.TabIndex = 5
        GroupBoxGenerateMasterKMLFile.TabStop = False
        GroupBoxGenerateMasterKMLFile.Text = "Generate Master KML file"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(10, 19)
        Label3.Name = "Label3"
        Label3.Size = New Size(312, 45)
        Label3.TabIndex = 3
        Label3.Text = "Select the Parent Folder Containing all of the KML files" & vbCrLf & "you want to combine. A new file named Mongoose.kml" & vbCrLf & "will be saved in the parent folder." & vbCrLf
        ' 
        ' PictureBox1
        ' 
        PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), Image)
        PictureBox1.Location = New Point(421, 12)
        PictureBox1.Name = "PictureBox1"
        PictureBox1.Size = New Size(153, 105)
        PictureBox1.TabIndex = 6
        PictureBox1.TabStop = False
        ' 
        ' PictureBox2
        ' 
        PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), Image)
        PictureBox2.Location = New Point(447, 315)
        PictureBox2.Name = "PictureBox2"
        PictureBox2.Size = New Size(100, 96)
        PictureBox2.TabIndex = 7
        PictureBox2.TabStop = False
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(607, 433)
        Controls.Add(PictureBox2)
        Controls.Add(PictureBox1)
        Controls.Add(GroupBoxGenerateMasterKMLFile)
        Controls.Add(GroupBoxKMZtoKML)
        Controls.Add(GroupBoxExcelToKml)
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        Name = "Form1"
        Text = "Mongoose"
        GroupBoxExcelToKml.ResumeLayout(False)
        GroupBoxExcelToKml.PerformLayout()
        GroupBoxKMZtoKML.ResumeLayout(False)
        GroupBoxKMZtoKML.PerformLayout()
        GroupBoxGenerateMasterKMLFile.ResumeLayout(False)
        GroupBoxGenerateMasterKMLFile.PerformLayout()
        CType(PictureBox1, ComponentModel.ISupportInitialize).EndInit()
        CType(PictureBox2, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub

    Friend WithEvents ButtonExcelToKML As Button
    Friend WithEvents ButtonKMZToKML As Button
    Friend WithEvents ButtonCombineKML As Button
    Friend WithEvents GroupBoxExcelToKml As GroupBox
    Friend WithEvents GroupBoxKMZtoKML As GroupBox
    Friend WithEvents GroupBoxGenerateMasterKMLFile As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents PictureBox2 As PictureBox
End Class
