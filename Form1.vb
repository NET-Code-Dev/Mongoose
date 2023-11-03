Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports OfficeOpenXml

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial

    End Sub

    Private Const FolderConfigFileName As String = "mongooseFolderConfig.txt"
    Private Const ExcelConfigFileName As String = "mongooseExcelConfig.txt"

    Private Function GetSavedPath(configFileName As String) As String
        Dim savedPath As String = String.Empty

        Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Dim fullPath As String = Path.Combine(appDataPath, configFileName)

        If File.Exists(fullPath) Then
            Try
                savedPath = File.ReadAllText(fullPath)
            Catch ex As Exception
                MessageBox.Show($"Could not read the config file: {ex.Message}")
            End Try
        Else
            MessageBox.Show($"Show me where it's at, and I'll remember for the next time!")
        End If

        Return savedPath
    End Function

    Private Sub SavePath(savedPath As String, configFileName As String)
        Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Dim mongoosePath As String = Path.Combine(appDataPath, "mongoose")
        Directory.CreateDirectory(mongoosePath) ' Create the directory if it does not exist

        Dim fullPath As String = Path.Combine(mongoosePath, configFileName)

        Try
            File.WriteAllText(fullPath, savedPath)
        Catch ex As Exception
            MessageBox.Show($"Could not save the directory path to the config file: {ex.Message}")
        End Try
    End Sub

    Private Function GetSavedFolderPath() As String
        Dim savedPath As String = String.Empty

        Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Dim configFilePath As String = Path.Combine(appDataPath, FolderConfigFileName)

        If File.Exists(configFilePath) Then
            Try
                savedPath = File.ReadAllText(configFilePath)
            Catch ex As Exception
                MessageBox.Show($"Could not read the config file: {ex.Message}")
            End Try
        Else
            MessageBox.Show($"Show me where it's at, and I'll remember for the next time!")
        End If

        Return savedPath
    End Function

    Private Sub SaveFolderPath(savedPath As String)
        Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Dim mongoosePath As String = Path.Combine(appDataPath, "mongoose")
        Directory.CreateDirectory(mongoosePath) ' Create the directory if it does not exist

        Dim configFilePath As String = Path.Combine(mongoosePath, FolderConfigFileName)

        Try
            File.WriteAllText(configFilePath, savedPath)
        Catch ex As Exception
            MessageBox.Show($"Could not save the folder path to the config file: {ex.Message}")
        End Try
    End Sub

    Private Function ShowFolderBrowserDialog(Optional defaultPath As String = "") As String
        Dim folderBrowserDialog As New FolderBrowserDialog()

        If Not String.IsNullOrEmpty(defaultPath) Then
            folderBrowserDialog.SelectedPath = defaultPath
        End If

        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            Return folderBrowserDialog.SelectedPath
        Else
            Return String.Empty
        End If
    End Function

    Private Function IsExcelFile(filePath As String) As Boolean
        Dim validExtensions As String() = {".xls", ".xlsx", ".xlsm"}
        Dim extension As String = Path.GetExtension(filePath)
        Return validExtensions.Contains(extension)
    End Function

    Private Function FileExistsWithWildcard(filePath As String) As Boolean
        Dim directoryPath As String = Path.GetDirectoryName(filePath)
        Dim fileName As String = Path.GetFileName(filePath)

        ' Replace the date portion with a wildcard
        Dim fileNameWildcard As String = fileName.Replace(Date.Today.ToString("yyyy-MM-dd"), "*")

        Dim files As String() = Directory.GetFiles(directoryPath, fileNameWildcard)

        Return files.Any()
    End Function

    Private Sub ButtonExcelToKML_Click(sender As Object, e As EventArgs) Handles ButtonExcelToKML.Click

        Dim excelFilePath As String = GetSavedPath(ExcelConfigFileName)

        ' If the saved path exists and is a valid Excel file, use it.
        If Not String.IsNullOrEmpty(excelFilePath) AndAlso FileExistsWithWildcard(excelFilePath) AndAlso IsExcelFile(excelFilePath) Then
            ' Use the saved path
        Else
            ' Prompt the user to select an Excel file.
            Dim openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                excelFilePath = openFileDialog.FileName
                SavePath(excelFilePath, ExcelConfigFileName)
            Else
                ' If no file is selected, exit the method.
                Return
            End If
        End If

        ' Declare these after we're sure we have a valid excelFilePath
        Dim kmlDirectoryPath As String = Path.GetDirectoryName(excelFilePath)

        ' Get the backup directory path for Personnel and EquipmentKML.kml
        Dim backupDirectoryPath As String = Path.Combine(kmlDirectoryPath, "_backups", "_Personnel_and_EquipmentKML")

        Using package As New ExcelPackage(New FileInfo(excelFilePath))
            Dim personnelWorksheet = package.Workbook.Worksheets("Personnel")
            Dim assetsWorksheet = package.Workbook.Worksheets("Pangea_Assets_Raw")

            Dim settings As New XmlWriterSettings()
            settings.Indent = True

            Using kmlMemoryStream As New MemoryStream()
                Using writer As XmlWriter = XmlWriter.Create(kmlMemoryStream, settings)
                    writer.WriteStartDocument()
                    writer.WriteStartElement("kml", "http://www.opengis.net/kml/2.2")
                    writer.WriteStartElement("Document")

                    ' Start of Personnel Folder
                    writer.WriteStartElement("Folder")
                    writer.WriteElementString("name", "Personnel")

                    ' Past City Folder
                    writer.WriteStartElement("Folder")
                    writer.WriteElementString("name", "Past City")

                    Dim i As Integer = 2
                    While personnelWorksheet.Cells(i, 1).Value IsNot Nothing
                        Dim employeeName As String = personnelWorksheet.Cells(i, 1).Value.ToString()
                        Dim cityLatitude As String = If(personnelWorksheet.Cells(i, 11).Value IsNot Nothing, personnelWorksheet.Cells(i, 11).Value.ToString(), "")
                        Dim cityLongitude As String = If(personnelWorksheet.Cells(i, 12).Value IsNot Nothing, personnelWorksheet.Cells(i, 12).Value.ToString(), "")

                        ' Create Pushpin for Past City
                        CreatePushpin(writer, employeeName, "", cityLatitude, cityLongitude)

                        i += 1
                    End While
                    ' End of Past City Folder
                    writer.WriteEndElement()

                    ' Current City Folder
                    writer.WriteStartElement("Folder")
                    writer.WriteElementString("name", "Current City")

                    i = 2 ' Re-initialize the variable i without re-declaring
                    While personnelWorksheet.Cells(i, 1).Value IsNot Nothing
                        Dim employeeName As String = personnelWorksheet.Cells(i, 1).Value.ToString()
                        Dim cityLatitude As String = If(personnelWorksheet.Cells(i, 13).Value IsNot Nothing, personnelWorksheet.Cells(i, 13).Value.ToString(), "")
                        Dim cityLongitude As String = If(personnelWorksheet.Cells(i, 14).Value IsNot Nothing, personnelWorksheet.Cells(i, 14).Value.ToString(), "")

                        ' Search for employee name in assetsWorksheet
                        Dim description As String = ""
                        Dim j As Integer = 2
                        While assetsWorksheet.Cells(j, 13).Value IsNot Nothing
                            Dim assetName As String = assetsWorksheet.Cells(j, 13).Value.ToString()
                            If assetName = employeeName Then
                                Dim headerC As String = assetsWorksheet.Cells(1, 3).Value.ToString()
                                Dim headerF As String = assetsWorksheet.Cells(1, 6).Value.ToString()
                                Dim valueC As String = assetsWorksheet.Cells(j, 3).Value.ToString()
                                Dim valueF As String = assetsWorksheet.Cells(j, 6).Value.ToString()
                                description &= $"{headerC}: {valueC}, {headerF}: {valueF}" & Environment.NewLine
                            End If
                            j += 1
                        End While

                        ' Create Pushpin for Current City
                        CreatePushpin(writer, employeeName, description, cityLatitude, cityLongitude)

                        i += 1
                    End While

                    ' End of Current City Folder
                    writer.WriteEndElement()

                    ' Future City Folder
                    writer.WriteStartElement("Folder")
                    writer.WriteElementString("name", "Future City")

                    i = 2 'Get coordinates for Future City
                    While personnelWorksheet.Cells(i, 1).Value IsNot Nothing
                        Dim employeeName As String = personnelWorksheet.Cells(i, 1).Value.ToString()
                        Dim cityLatitude As String = If(personnelWorksheet.Cells(i, 15).Value IsNot Nothing, personnelWorksheet.Cells(i, 15).Value.ToString(), "")
                        Dim cityLongitude As String = If(personnelWorksheet.Cells(i, 16).Value IsNot Nothing, personnelWorksheet.Cells(i, 16).Value.ToString(), "")

                        ' Create Pushpin for Future City
                        CreatePushpin(writer, employeeName, "", cityLatitude, cityLongitude)

                        i += 1
                    End While
                    ' End of Future City Folder
                    writer.WriteEndElement()

                    ' End of Personnel Folder
                    writer.WriteEndElement()

                    ' Start of Assets Folder
                    writer.WriteStartElement("Folder")
                    writer.WriteElementString("name", "Assets")

                    ' Subfolders for Assets: Equipment, Instruments, Safety Equipment, IT Device, Vehicle
                    Dim subFoldersAssets As New List(Of String) From {"Equipment", "Instruments", "Safety Equipment", "IT Device", "Vehicle"}

                    For Each folder In subFoldersAssets
                        ' Start of Sub Folder
                        writer.WriteStartElement("Folder")
                        writer.WriteElementString("name", folder)

                        Dim k As Integer = 2
                        While assetsWorksheet.Cells(k, 3).Value IsNot Nothing
                            Dim folderName As String = assetsWorksheet.Cells(k, 1).Value.ToString()
                            Dim itemName As String = assetsWorksheet.Cells(k, 3).Value.ToString()

                            ' Get header titles
                            Dim header2 As String = assetsWorksheet.Cells(1, 2).Value.ToString()
                            Dim header5 As String = assetsWorksheet.Cells(1, 5).Value.ToString()
                            Dim header6 As String = assetsWorksheet.Cells(1, 6).Value.ToString()
                            Dim header10 As String = assetsWorksheet.Cells(1, 10).Value.ToString()
                            Dim header13 As String = assetsWorksheet.Cells(1, 13).Value.ToString()

                            ' Get values
                            Dim value2 As String = assetsWorksheet.Cells(k, 2).Value.ToString()
                            Dim value5 As String = assetsWorksheet.Cells(k, 5).Value.ToString()
                            Dim value6 As String = assetsWorksheet.Cells(k, 6).Value.ToString()
                            Dim value10 As String = assetsWorksheet.Cells(k, 10).Value.ToString()
                            Dim value13 As String = assetsWorksheet.Cells(k, 13).Value.ToString()

                            ' Create description string
                            Dim description As String = $"{header13}: {value13}{vbCrLf}{header2}: {value2}{vbCrLf}{header5}: {value5}{vbCrLf}{header6}: {value6}{vbCrLf}{header10}: {value10}"

                            Dim personnelName As String = assetsWorksheet.Cells(k, 13).Value.ToString()

                            ' Find matching personnel for coordinates
                            Dim j As Integer = 2
                            Dim latitude As String = ""
                            Dim longitude As String = ""
                            While personnelWorksheet.Cells(j, 1).Value IsNot Nothing
                                If personnelWorksheet.Cells(j, 1).Value.ToString() = personnelName Then
                                    latitude = personnelWorksheet.Cells(j, 13).Value.ToString()
                                    longitude = personnelWorksheet.Cells(j, 14).Value.ToString()
                                    Exit While
                                End If
                                j += 1
                            End While

                            ' Create Pushpin only if the folder name matches
                            If folderName = folder Then
                                CreatePushpin(writer, itemName, description, latitude, longitude)
                            End If

                            k += 1
                        End While

                        ' End of Sub Folder
                        writer.WriteEndElement()
                    Next

                    ' End of Assets Folder
                    writer.WriteEndElement()

                    writer.WriteEndElement() ' End Document
                    writer.WriteEndElement() ' End kml
                    writer.WriteEndDocument()
                End Using

                Dim currentDate As String = DateTime.Now.ToString("yyyy-MM-dd")
                Dim newKMLFile As String = currentDate & "_Personnel_and_Equipment.kml"
                Dim kmlFilePath As String = Path.Combine(kmlDirectoryPath, newKMLFile)

                ' Get all files matching the pattern "????-??-??_Personnel_and_Equipment.kml"
                Dim directoryInfo As New DirectoryInfo(kmlDirectoryPath)
                Dim allFiles As FileInfo() = directoryInfo.GetFiles("????-??-??_Personnel_and_Equipment.kml")

                If allFiles.Length > 0 Then
                    ' Ensure backup directory exists
                    If Not Directory.Exists(backupDirectoryPath) Then
                        Directory.CreateDirectory(backupDirectoryPath)
                    End If

                    For Each file As FileInfo In allFiles
                        Dim originalFileName As String = file.Name
                        Dim backupFilePath As String = Path.Combine(backupDirectoryPath, originalFileName)

                        ' Check if a file with the same name already exists in the backup directory
                        Dim counter As Integer = 1
                        While System.IO.File.Exists(backupFilePath)
                            Dim fileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(originalFileName)
                            Dim extension As String = Path.GetExtension(originalFileName)
                            backupFilePath = Path.Combine(backupDirectoryPath, $"{fileNameWithoutExtension}_{counter}{extension}")
                            counter += 1
                        End While

                        ' Move the file to the backup directory
                        file.MoveTo(backupFilePath)
                    Next
                End If

                ' File path for the parent directory
                Dim newKMLFilePath As String = Path.Combine(kmlDirectoryPath, newKMLFile)

                ' Save the KML data to the new KML file in the parent directory
                CreateKMLFile(newKMLFilePath, kmlMemoryStream)

                ' File path for the _mongoose subdirectory
                Dim mongooseSubdirectoryPath As String = Path.Combine(kmlDirectoryPath, "_mongoose")
                Dim mongooseSubdirectoryFilePath As String = Path.Combine(mongooseSubdirectoryPath, newKMLFile)

                ' Ensure _mongoose subdirectory exists
                If Not Directory.Exists(mongooseSubdirectoryPath) Then
                    Directory.CreateDirectory(mongooseSubdirectoryPath)
                End If

                ' Save the KML data to the new KML file in the _mongoose subdirectory
                CreateKMLFile(mongooseSubdirectoryFilePath, kmlMemoryStream)

            End Using
        End Using

        MessageBox.Show("Mongoose was Successful!")

    End Sub

    'This code is part of a button click event handler. It first checks if a folder path has been saved, and if so, uses it. If not, it prompts the user to select a folder
    '. If a folder is selected, it is saved. Finally, the code calls a method to process the directory and displays a success message.
    Private Sub ButtonKMZToKML_Click(sender As Object, e As EventArgs) Handles ButtonKMZToKML.Click
        Dim folderPath As String = GetSavedFolderPath()

        ' If the saved folder path exists, use it. Otherwise, prompt the user to select a folder.
        If Not String.IsNullOrEmpty(folderPath) AndAlso Directory.Exists(folderPath) Then
            ' Use the saved folder path
        Else
            folderPath = ShowFolderBrowserDialog()
            If String.IsNullOrEmpty(folderPath) OrElse Not Directory.Exists(folderPath) Then
                ' If no folder is selected, exit the method.
                Return
            End If
            SaveFolderPath(folderPath)
        End If

        ' Continue with the logic for processing the directory
        ProcessDirectory(folderPath)

        MessageBox.Show("Mongoose Successfully converted KMZ's to KML's!")
    End Sub

    'This code is a subroutine that is triggered when a button is clicked. It is responsible for combining KML files into a single KML file. It first checks if a saved
    ' folder path exists, and if not, it prompts the user to select a folder. It then creates a new KML file with the current date in the folder path. It then gets all
    ' files matching the pattern "????-??-?? Mongoose.kml" and moves them to a backup directory. Finally, it generates a subfolder hierarchy KML file and displays a message
    ' box to the user.
    Private Sub ButtonCombineKML_Click(sender As Object, e As EventArgs) Handles ButtonCombineKML.Click
        Dim folderPath As String = GetSavedFolderPath()

        ' If the saved folder path exists, use it. Otherwise, prompt the user to select a folder.
        If Not String.IsNullOrEmpty(folderPath) AndAlso Directory.Exists(folderPath) Then
            ' Use the saved folder path
        Else
            folderPath = ShowFolderBrowserDialog()
            If String.IsNullOrEmpty(folderPath) OrElse Not Directory.Exists(folderPath) Then
                ' If no folder is selected, exit the method.
                Return
            End If
            SaveFolderPath(folderPath)
        End If

        Dim currentDate As String = DateTime.Now.ToString("yyyy-MM-dd")
        Dim newKMLFile As String = currentDate & " Mongoose.kml"
        Dim kmlFilePath As String = Path.Combine(folderPath, newKMLFile)

        ' Get the DirectoryInfo for the selected folder
        Dim directoryInfo As New DirectoryInfo(folderPath)

        ' Get all files matching the pattern "????-??-?? Mongoose.kml"
        Dim allFiles As FileInfo() = directoryInfo.GetFiles("????-??-?? Mongoose.kml")

        If allFiles.Length > 0 Then
            ' The _backups/_MongooseKML subdirectory path
            Dim backupDirectoryPath As String = Path.Combine(folderPath, "_backups", "_MongooseKML")

            For Each file As FileInfo In allFiles
                ' Get the original file name
                Dim originalFileName As String = file.Name

                ' Determine the backup file path
                Dim backupFilePath As String = Path.Combine(backupDirectoryPath, originalFileName)

                ' Check if a file with the same name already exists in the backup directory
                Dim counter As Integer = 1
                While System.IO.File.Exists(backupFilePath)
                    Dim fileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(originalFileName)
                    Dim extension As String = Path.GetExtension(originalFileName)
                    backupFilePath = Path.Combine(backupDirectoryPath, $"{fileNameWithoutExtension}_{counter}{extension}")
                    counter += 1
                End While

                ' Move the file to the backup directory
                file.MoveTo(backupFilePath)
            Next
        End If

        Try
            GenerateSubfolderHierarchyKML(folderPath, kmlFilePath)
            MessageBox.Show("Mongoose successfully created a combined KML file!")
        Catch ex As Exception
            MessageBox.Show("An error occurred: " + ex.Message)
        End Try
    End Sub

    'This subroutine is used to generate a KML file from a folder hierarchy. It takes in a XmlDocument, XmlElement, and a folder path as parameters. It then gets the folder
    ' name from the folder path and creates a new folder element if it is not the root directory. It then processes KML files within the folder and imports Document and
    ' Folder elements from the KML files into the XmlDocument. Finally, it recursively processes subdirectories.
    Private Sub CreateCombinedKML(ByRef doc As XmlDocument, ByRef parentElement As XmlElement, folderPath As String, Optional isFirstLevel As Boolean = False)
        Dim folderInfo As New DirectoryInfo(folderPath)
        Dim logFilePath As String = Path.Combine(folderPath, "error_log.txt") 'Define the log file path

        ' Process folders
        For Each subFolder As DirectoryInfo In folderInfo.GetDirectories()
            ' Skip folders that start with an underscore
            If Not subFolder.Name.StartsWith("_") Then
                Dim folderElement As XmlElement = doc.CreateElement("Folder")
                parentElement.AppendChild(folderElement)

                Dim folderNameElement As XmlElement = doc.CreateElement("name")
                folderNameElement.InnerText = subFolder.Name
                folderElement.AppendChild(folderNameElement)

                CreateCombinedKML(doc, folderElement, subFolder.FullName)

                ' Remove the empty folder if it has no child elements
                If folderElement.ChildNodes.Count = 1 Then
                    parentElement.RemoveChild(folderElement)
                End If
            End If
        Next

        ' Process KML and KMZ files
        Dim existingElements As New List(Of String)() ' Track existing names

        For Each file As FileInfo In folderInfo.GetFiles().Where(Function(f) f.Extension.ToLower() = ".kml" OrElse f.Extension.ToLower() = ".kmz")
            Try
                Dim fileElement As XmlElement = doc.CreateElement("Document")
                parentElement.AppendChild(fileElement)

                Dim fileNameElement As XmlElement = doc.CreateElement("name")
                fileNameElement.InnerText = Path.GetFileNameWithoutExtension(file.Name)

                If Not existingElements.Contains(fileNameElement.InnerText) Then
                    fileElement.AppendChild(fileNameElement)

                    If file.Extension.ToLower() = ".kml" Then
                        Try
                            Dim kmlDoc As New XmlDocument()
                            kmlDoc.Load(file.FullName)
                            Dim kmlRoot As XmlElement = kmlDoc.DocumentElement
                            For Each node As XmlNode In kmlRoot.ChildNodes
                                If node.Name <> "Document" Then
                                    fileElement.AppendChild(doc.ImportNode(node, True))
                                End If
                            Next

                            If fileElement.ChildNodes.Count = 1 Then
                                parentElement.RemoveChild(fileElement)
                            Else
                                existingElements.Add(fileNameElement.InnerText) ' Add name to existing elements list
                            End If
                        Catch xmlEx As XmlException
                            MessageBox.Show($"XML error in file {file.FullName}: {xmlEx.Message}")
                            LogError($"XML error in file {file.FullName}: {xmlEx.Message}", folderPath)
                        Catch ex As Exception
                            MessageBox.Show($"General error in file {file.FullName}: {ex.Message}")
                            LogError($"General error in file {file.FullName}: {ex.Message}", folderPath)
                        End Try
                    ElseIf file.Extension.ToLower() = ".kmz" Then
                        ' Handle KMZ files if needed
                    End If
                Else
                    parentElement.RemoveChild(fileElement) ' Remove duplicate file element
                End If
            Catch ex As Exception
                MessageBox.Show($"Error processing file {file.FullName}: {ex.Message}")
                LogError($"Error processing file {file.FullName}: {ex.Message}", folderPath)
            End Try
        Next

    End Sub

    'This subroutine is used to import folders and placemarks from a KML file into an XML document. It takes three parameters: a reference to the XML document, a reference
    ' to the parent element, and a reference to the KML element. The subroutine iterates through the child nodes Of the KML element. If the node Is an element, it
    ' checks if it is a folder or a placemark. If it is a folder, it checks if an existing folder with the same name exists in the parent element. If it does, it recursively
    ' calls the subroutine with the existing folder as the parent element. If it does not, it creates a new folder element in the parent element and imports the KML element into it.
    ' If the node is a placemark, it checks if an existing placemark with the same name exists in the parent element. If it does not, it creates a new placemark element in the parent
    ' element and imports the KML element into it.
    Private Sub ImportFoldersRecursively(doc As XmlDocument, parentElement As XmlElement, kmlElement As XmlElement)
        For Each childNode As XmlNode In kmlElement.ChildNodes
            If childNode.NodeType = XmlNodeType.Element Then
                Dim childElement As XmlElement = CType(childNode, XmlElement)

                If childElement.Name = "Folder" Then
                    Dim existingFolder = FindExistingElementByName(parentElement, childElement.Name, childElement.SelectSingleNode("name").InnerText)
                    If existingFolder IsNot Nothing Then
                        ImportFoldersRecursively(doc, existingFolder, childElement)
                    Else
                        Dim folderElement As XmlElement = doc.CreateElement("Folder")
                        parentElement.AppendChild(folderElement)
                        ImportElement(doc, folderElement, childElement)
                        ImportFoldersRecursively(doc, folderElement, childElement)
                    End If
                ElseIf childElement.Name = "Placemark" Then
                    Dim existingPlacemark = FindExistingElementByName(parentElement, childElement.Name, childElement.SelectSingleNode("name").InnerText)
                    If existingPlacemark Is Nothing Then
                        Dim placemarkElement As XmlElement = doc.CreateElement("Placemark")
                        parentElement.AppendChild(placemarkElement)
                        ImportElement(doc, placemarkElement, childElement)
                    End If
                End If
            End If
        Next
    End Sub

    'This function is used to search for an existing XML element by its name. It takes three parameters: parentElement (the parent element to search within), elementName
    ' (the name of the element to search for), and nameValue (the value of the element's name attribute). It then searches through the parent element's child elements
    ' for one with a name attribute that matches the nameValue parameter. If it finds a match, it returns the element, otherwise it returns nothing.
    Private Function FindExistingElementByName(parentElement As XmlElement, elementName As String, nameValue As String) As XmlElement
        Dim elements As XmlNodeList = parentElement.GetElementsByTagName(elementName)
        For Each element As XmlElement In elements
            If element.SelectSingleNode("name").InnerText = nameValue Then
                Return element
            End If
        Next
        Return Nothing
    End Function

    Private Sub ImportElement(doc As XmlDocument, destinationElement As XmlElement, sourceElement As XmlElement)

        'This subroutine Is used To copy an XML element from one document To another. It takes three parameters: an XmlDocument Object (doc) which Is the destination
        'document, an XmlElement object (destinationElement) which is the element in the destination document to which the source element will be appended, and an XmlElement
        'object (sourceElement) which is the element to be copied.

        'Loop through each attribute in the source element
        For Each attribute As XmlAttribute In sourceElement.Attributes

            'Create a new attribute with the same name
            Dim attributeCopy As XmlAttribute = doc.CreateAttribute(attribute.Name)

            'Set the value of the new attribute to the value of the source attribute
            attributeCopy.Value = attribute.Value

            'Append the new attribute to the destination element
            destinationElement.Attributes.Append(attributeCopy)

            'End the loop
        Next

        'Loop through each child node of the source element
        For Each childNode As XmlNode In sourceElement.ChildNodes

            'Check if the node is an element
            If childNode.NodeType = XmlNodeType.Element Then

                'Import the element into the destination document
                Dim importedElement As XmlElement = CType(doc.ImportNode(CType(childNode, XmlElement), True), XmlElement)

                'Append the imported element to the destination element
                destinationElement.AppendChild(importedElement)
            End If
        Next
    End Sub

    'This code creates a KML file from a MemoryStream. It first gets the current date in the format yyyy-mm-dd and creates a new file name based on that. It then combines
    ' the directory path and the new file name to create the full path. It then checks if the directory exists, and if not, creates it. Finally, it writes the KML data
    ' to the file.
    Private Sub CreateKMLFile(fullPath As String, kmlMemoryStream As MemoryStream) 'Method used in ButtonExcelToKML_Click
        ' Write the KML data to the file
        Using kmlFileStream As FileStream = File.OpenWrite(fullPath)
            kmlMemoryStream.WriteTo(kmlFileStream)
        End Using
    End Sub

    'This subroutine is used to create a pushpin in a KML file. It takes in five parameters: an XmlWriter object, a pushpin name, a description, a latitude, and a long
    'itude. It then uses the XmlWriter object to write the start of a Placemark element, followed by the name and description of the pushpin. It then writes the start
    ' of a Point element, followed by the coordinates of the pushpin (longitude, latitude, and 0). Finally, it writes the end of the Point and Placemark elements.
    Private Sub CreatePushpin(writer As XmlWriter, pushpinName As String, description As String, latitude As String, longitude As String) 'Method used in ButtonExcelToKML_Click
        writer.WriteStartElement("Placemark")
        writer.WriteElementString("name", pushpinName)
        writer.WriteElementString("description", description)
        writer.WriteStartElement("Point")
        writer.WriteElementString("coordinates", $"{longitude},{latitude},0")
        writer.WriteEndElement() ' End Point
        writer.WriteEndElement() ' End Placemark
    End Sub

    'This code is used to generate a subfolder hierarchy in a KML (Keyhole Markup Language) document. It creates a KML document, adds the necessary namespaces, and creates
    ' a main Folder element. It then calls the CreateCombinedKML method to create the subfolder hierarchy. Finally, it saves the KML document to the specified file.
    Private Sub GenerateSubfolderHierarchyKML(parentFolderPath As String, kmlFilePath As String) 'Method used in ButtonCombineKML_Click
        ' Create the KML document
        Dim doc As New XmlDocument()

        ' Create the root KML element with namespaces
        Dim kmlElement As XmlElement = doc.CreateElement("kml")
        doc.AppendChild(kmlElement)

        ' Add namespaces to kml element
        kmlElement.SetAttribute("xmlns", "http://www.opengis.net/kml/2.2")
        kmlElement.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        kmlElement.SetAttribute("xsi:schemaLocation", "http://www.opengis.net/kml/2.2 http://schemas.opengis.net/kml/2.2.0/ogckml22.xsd")

        ' Create the main Folder element
        Dim mainFolderElement As XmlElement = doc.CreateElement("Folder")
        kmlElement.AppendChild(mainFolderElement)

        ' Handling the KML files in the root directory
        Dim rootDirectoryKmlFiles As String() = Directory.GetFiles(parentFolderPath, "*.kml")
        For Each kmlFile As String In rootDirectoryKmlFiles
            Dim kmlDocument As New XmlDocument()
            kmlDocument.Load(kmlFile)

            Dim documentElement As XmlElement = TryCast(kmlDocument.GetElementsByTagName("Document")(0), XmlElement)
            If documentElement IsNot Nothing Then
                mainFolderElement.InnerXml &= documentElement.InnerXml
            End If
        Next

        ' Create the subfolder hierarchy
        CreateCombinedKML(doc, mainFolderElement, parentFolderPath, True)

        ' Save the KML document to the specified file
        doc.Save(kmlFilePath)
    End Sub

    'This code is a recursive method used to process a directory and all of its subdirectories. It begins by getting a list of all the files in the directory and processing
    ' each one. Then it gets a list of all the subdirectories in the directory and calls the ProcessDirectory method on each one. This allows the method to process all
    ' the files and subdirectories in the directory and its subdirectories.
    Private Sub ProcessDirectory(directory As String) 'Method used in ButtonKMZToKML_Click
        ' Process the list of files found in the directory
        Dim fileEntries As String() = System.IO.Directory.GetFiles(directory)
        For Each fileName As String In fileEntries
            ProcessFile(fileName)
        Next fileName

        ' Recurse into subdirectories of this directory
        Dim subdirectoryEntries As String() = System.IO.Directory.GetDirectories(directory)
        For Each subdirectory As String In subdirectoryEntries
            ProcessDirectory(subdirectory)
        Next subdirectory
    End Sub

    'This code is part of a larger program that processes a directory of files. This particular code is used to process files with a .kmz extension. It first creates a
    ' temporary folder and extracts the contents of the KMZ file into it. It then reads the contents of the doc.kml file and uses a regular expression to match the <Document
    ' ...> tag. It then replaces the last occurrence of the </Document> tag with </Folder>. Finally, it saves the modified contents as a .kml file and deletes the temporary
    ' folder.
    Private Sub ProcessFile(filePath As String) 'Method used in ProcessDirectory
        ' Check if the file has a .kmz extension
        If Path.GetExtension(filePath).ToLower() = ".kmz" Then
            Dim tempFolderName = Path.GetFileNameWithoutExtension(filePath) & "_" & Guid.NewGuid().ToString()
            Dim tempFolderPath As String = Path.Combine(Path.GetTempPath(), tempFolderName)

            Directory.CreateDirectory(tempFolderPath)

            ' Extract the contents of the KMZ file
            ZipFile.ExtractToDirectory(filePath, tempFolderPath)

            ' Process the doc.kml file
            Dim kmlFilePath As String = Path.Combine(tempFolderPath, "doc.kml")
            If File.Exists(kmlFilePath) Then
                Dim contents As String = File.ReadAllText(kmlFilePath)

                ' Regular expression to match <Document ...> tag
                Dim openTagPattern As String = "<Document\b[^>]*>"
                contents = Regex.Replace(contents, openTagPattern, "<Folder>", RegexOptions.IgnoreCase)

                ' Replace last occurrence of </Document> with </Folder>
                Dim closeTagPattern As String = "</Document>"
                contents = Regex.Replace(contents, closeTagPattern, "</Folder>", RegexOptions.IgnoreCase)

                ' Save modified contents as .kml
                Dim newKmlFilePath As String = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath) & ".kml")
                File.WriteAllText(newKmlFilePath, contents, Encoding.UTF8)
            End If

            ' Delete the temp folder
            Directory.Delete(tempFolderPath, True)
        End If
    End Sub

    Private Sub LogError(errorMessage As String, userFolderPath As String)
        ' Assuming 'userFolderPath' is the path you get which contains the 'Mongoose' directory
        ' Find the 'Mongoose' directory in the path
        Dim mongooseDirectory As String = userFolderPath
        While Not Path.GetFileName(mongooseDirectory).Equals("Mongoose", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrEmpty(Path.GetDirectoryName(mongooseDirectory))
            mongooseDirectory = Path.GetDirectoryName(mongooseDirectory)
        End While

        ' If 'Mongoose' directory is not found in the path, you can decide to default to a known location or throw an error
        If String.IsNullOrEmpty(mongooseDirectory) Then
            ' Handle the error accordingly, e.g., throw an exception or default to a known path
            Throw New DirectoryNotFoundException("The 'Mongoose' directory was not found in the provided path.")
        End If

        ' Construct the log file path within the 'Mongoose' directory
        Dim logFilePath As String = Path.Combine(mongooseDirectory, "error_log.txt")

        ' Check if the directory exists; if not, create it
        Dim logFileDirectory As String = Path.GetDirectoryName(logFilePath)
        If Not Directory.Exists(logFileDirectory) Then
            Directory.CreateDirectory(logFileDirectory)
        End If

        ' Append the error message to the log file with a timestamp
        Using writer As New StreamWriter(logFilePath, True)
            writer.WriteLine($"{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}: {errorMessage}")
        End Using
    End Sub


End Class