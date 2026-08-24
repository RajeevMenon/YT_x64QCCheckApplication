Imports System.Configuration
Imports System.IO
Imports System.Net.Http
Imports System.Text.Json
Imports System.Threading.Tasks
Imports Newtonsoft.Json

Public Class DocumentApiClient
    Private ReadOnly _baseUrl As String = Link.DocClient '"http://localhost:56101/" 'ConfigurationManager.AppSettings("ApiBaseUrl")
    Private ReadOnly _client As New HttpClient()

    Public Sub New(baseUrl As String)
        If String.IsNullOrWhiteSpace(baseUrl) Then
            Throw New Exception("ApiBaseUrl is missing from the application configuration file.")
        End If
        _client.BaseAddress = New Uri(baseUrl)
    End Sub

    ''' <summary>
    ''' Searches the MVC backend for files matching a wildcard within a specific store.
    ''' </summary>
    Public Async Function SearchFilesAsync(storeKey As String, subFolder As String, searchPattern As String) As Task(Of List(Of String))
        ' Append storeKey to the query parameters
        Dim url As String = $"Documents/SearchFiles?storeKey={Uri.EscapeDataString(storeKey)}&subFolder={Uri.EscapeDataString(subFolder)}&searchPattern={Uri.EscapeDataString(searchPattern)}"

        ' ADD THIS LINE TO DEBUG:
        'MessageBox.Show("Actual URL being called: " & New Uri(_client.BaseAddress, url).ToString())

        Dim response As HttpResponseMessage = Await _client.GetAsync(url)

        If response.IsSuccessStatusCode Then
            Dim jsonString As String = Await response.Content.ReadAsStringAsync()
            Return System.Text.Json.JsonSerializer.Deserialize(Of List(Of String))(jsonString)
        Else
            Throw New Exception($"Error searching files: {response.StatusCode} - {Await response.Content.ReadAsStringAsync()}")
        End If
    End Function

    ''' <summary>
    ''' Uploads multiple files simultaneously to a specific store on the MVC backend.
    ''' </summary>
    Public Async Function SaveMultipleFilesAsync(storeKey As String, subFolder As String, overwrite As Boolean, localFilePaths As List(Of String)) As Task
        ' Append storeKey to the query parameters
        Dim url As String = $"Documents/SaveMultipleFiles?storeKey={Uri.EscapeDataString(storeKey)}&subFolder={Uri.EscapeDataString(subFolder)}&overwrite={overwrite.ToString().ToLower()}"

        Using form As New MultipartFormDataContent()
            For Each filePath In localFilePaths
                If File.Exists(filePath) Then
                    Dim fileName As String = Path.GetFileName(filePath)
                    Dim mimeType As String = GetMimeType(fileName)

                    Dim fileBytes As Byte() = File.ReadAllBytes(filePath)
                    Dim fileContent As New ByteArrayContent(fileBytes)

                    ' Dynamically assign the correct content type instead of hardcoding PDF
                    fileContent.Headers.ContentType = Net.Http.Headers.MediaTypeHeaderValue.Parse(mimeType)

                    form.Add(fileContent, "files", fileName)
                End If
            Next

            Dim response As HttpResponseMessage = Await _client.PostAsync(url, form)

            If Not response.IsSuccessStatusCode Then
                Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
                Throw New Exception($"Failed to upload files: {errorMsg}")
            End If
        End Using
    End Function

    ''' <summary>
    ''' Determines the correct MIME type based on the file extension.
    ''' </summary>
    Private Function GetMimeType(fileName As String) As String
        Dim extension As String = Path.GetExtension(fileName).ToLowerInvariant()

        Select Case extension
            Case ".pdf"
                Return "application/pdf"
            Case ".png"
                Return "image/png"
            Case ".jpg", ".jpeg"
                Return "image/jpeg"
            Case ".txt"
                Return "text/plain"
            Case ".doc", ".docx"
                Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            Case ".xls", ".xlsx"
                Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Case ".csv"
                Return "text/csv"
            Case Else
                ' The standard default for unknown binary files
                Return "application/octet-stream"
        End Select
    End Function

    ''' <summary>
    ''' 1. Downloads a file from the API and saves it to a specified local directory.
    ''' </summary>
    Public Async Function DownloadToLocalFolderAsync(storeKey As String, subFolder As String, fileName As String, localSaveDirectory As String) As Task(Of String)
        Dim url As String = $"Documents/GetFile?storeKey={Uri.EscapeDataString(storeKey)}&subFolder={Uri.EscapeDataString(subFolder)}&fileName={Uri.EscapeDataString(fileName)}"
        Dim response As HttpResponseMessage = Await _client.GetAsync(url)

        If response.IsSuccessStatusCode Then
            If Not Directory.Exists(localSaveDirectory) Then
                Directory.CreateDirectory(localSaveDirectory)
            End If

            Dim localFilePath As String = Path.Combine(localSaveDirectory, fileName)
            Dim fileBytes As Byte() = Await response.Content.ReadAsByteArrayAsync()
            File.WriteAllBytes(localFilePath, fileBytes)

            Return localFilePath
        Else
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Error downloading file: {response.StatusCode} - {errorMsg}")
        End If
    End Function


    ''' <summary>
    ''' 2. Downloads a file from the API and saves it directly to the Windows Temporary folder.
    ''' Windows will eventually clean this up automatically. Useful for Process.Start().
    ''' </summary>
    Public Async Function DownloadToTempFolderAsync(storeKey As String, subFolder As String, fileName As String) As Task(Of String)
        Dim url As String = $"Documents/GetFile?storeKey={Uri.EscapeDataString(storeKey)}&subFolder={Uri.EscapeDataString(subFolder)}&fileName={Uri.EscapeDataString(fileName)}"
        Dim response As HttpResponseMessage = Await _client.GetAsync(url)

        If response.IsSuccessStatusCode Then
            ' Resolve the specific Windows Temp path for the current user
            Dim tempFilePath As String = Path.Combine(Path.GetTempPath(), fileName)

            Dim fileBytes As Byte() = Await response.Content.ReadAsByteArrayAsync()
            File.WriteAllBytes(tempFilePath, fileBytes)

            Return tempFilePath
        Else
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Error downloading file to Temp: {response.StatusCode} - {errorMsg}")
        End If
    End Function


    ''' <summary>
    ''' 3. Downloads a file from the API directly into a MemoryStream. 
    ''' The file never touches the hard drive. Perfect for 3rd-party PDF viewers.
    ''' </summary>
    Public Async Function DownloadToMemoryStreamAsync(storeKey As String, subFolder As String, fileName As String) As Task(Of MemoryStream)
        Dim url As String = $"Documents/GetFile?storeKey={Uri.EscapeDataString(storeKey)}&subFolder={Uri.EscapeDataString(subFolder)}&fileName={Uri.EscapeDataString(fileName)}"
        Dim response As HttpResponseMessage = Await _client.GetAsync(url)

        If response.IsSuccessStatusCode Then
            ' Read the bytes and load them into a new MemoryStream object
            Dim fileBytes As Byte() = Await response.Content.ReadAsByteArrayAsync()
            Return New MemoryStream(fileBytes)
        Else
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Error downloading file to MemoryStream: {response.StatusCode} - {errorMsg}")
        End If
    End Function

    ''' <summary>
    ''' Gets a list of all subdirectory names inside a specific folder path.
    ''' </summary>
    Public Async Function GetSubDirectoriesAsync(storeKey As String, subFolder As String) As Task(Of List(Of String))
        Dim url As String = $"Documents/GetSubDirectories?storeKey={Uri.EscapeDataString(storeKey)}&subFolder={Uri.EscapeDataString(subFolder)}"
        Dim response As HttpResponseMessage = Await _client.GetAsync(url)

        If response.IsSuccessStatusCode Then
            Dim jsonString As String = Await response.Content.ReadAsStringAsync()
            Return System.Text.Json.JsonSerializer.Deserialize(Of List(Of String))(jsonString)
        Else
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Error fetching subdirectories: {response.StatusCode} - {errorMsg}")
        End If
    End Function

    ''' <summary>
    ''' Creates a new subdirectory remotely on the server.
    ''' </summary>
    Public Async Function CreateDirectoryAsync(storeKey As String, subFolder As String, newFolderName As String) As Task
        Dim url As String = $"Documents/CreateDirectory?storeKey={Uri.EscapeDataString(storeKey)}&subFolder={Uri.EscapeDataString(subFolder)}&newFolderName={Uri.EscapeDataString(newFolderName)}"
        Dim response As HttpResponseMessage = Await _client.PostAsync(url, Nothing)

        If Not response.IsSuccessStatusCode Then
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Failed to create directory: {errorMsg}")
        End If
    End Function

    ''' <summary>
    ''' Deletes a subdirectory remotely on the server.
    ''' </summary>
    Public Async Function DeleteDirectoryAsync(storeKey As String, subFolder As String, folderToDelete As String, Optional recursive As Boolean = True) As Task
        Dim url As String = $"Documents/DeleteDirectory?storeKey={Uri.EscapeDataString(storeKey)}&subFolder={Uri.EscapeDataString(subFolder)}&folderToDelete={Uri.EscapeDataString(folderToDelete)}&recursive={recursive.ToString().ToLower()}"
        Dim response As HttpResponseMessage = Await _client.PostAsync(url, Nothing)

        If Not response.IsSuccessStatusCode Then
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Failed to delete directory: {errorMsg}")
        End If
    End Function

    ''' <summary>
    ''' Copies an existing file remotely to a new destination folder. Can also rename the file.
    ''' </summary>
    ''' <param name="storeKey">The root storage directory key.</param>
    ''' <param name="sourceSubFolder">The subfolder path of the original file.</param>
    ''' <param name="sourceFileName">The name of the file to copy.</param>
    ''' <param name="destSubFolder">The subfolder path where the file should be copied.</param>
    ''' <param name="destFileName">The new file name. Pass Nothing or "" to keep the original name.</param>
    ''' <param name="overwrite">If True, overwrites an existing file at the destination.</param>
    Public Async Function CopyFileAsync(storeKey As String, sourceSubFolder As String, sourceFileName As String, destSubFolder As String, destFileName As String, Optional overwrite As Boolean = False) As Task
        Dim url As String = $"Documents/CopyFile?storeKey={Uri.EscapeDataString(storeKey)}" &
                            $"&sourceSubFolder={Uri.EscapeDataString(If(sourceSubFolder, ""))}" &
                            $"&sourceFileName={Uri.EscapeDataString(sourceFileName)}" &
                            $"&destSubFolder={Uri.EscapeDataString(If(destSubFolder, ""))}" &
                            $"&destFileName={Uri.EscapeDataString(If(destFileName, ""))}" &
                            $"&overwrite={overwrite.ToString().ToLower()}"

        Dim response As HttpResponseMessage = Await _client.PostAsync(url, Nothing)

        If Not response.IsSuccessStatusCode Then
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Failed to copy file: {errorMsg}")
        End If
    End Function

    ''' <summary>
    ''' Moves an existing file remotely to a new destination folder. Can also rename the file.
    ''' </summary>
    ''' <param name="storeKey">The root storage directory key.</param>
    ''' <param name="sourceSubFolder">The subfolder path of the original file.</param>
    ''' <param name="sourceFileName">The name of the file to move.</param>
    ''' <param name="destSubFolder">The subfolder path where the file should be moved.</param>
    ''' <param name="destFileName">The new file name. Pass Nothing or "" to keep the original name.</param>
    ''' <param name="overwrite">If True, overwrites an existing file at the destination.</param>
    Public Async Function MoveFileAsync(storeKey As String, sourceSubFolder As String, sourceFileName As String, destSubFolder As String, destFileName As String, Optional overwrite As Boolean = False) As Task
        Dim url As String = $"Documents/MoveFile?storeKey={Uri.EscapeDataString(storeKey)}" &
                            $"&sourceSubFolder={Uri.EscapeDataString(If(sourceSubFolder, ""))}" &
                            $"&sourceFileName={Uri.EscapeDataString(sourceFileName)}" &
                            $"&destSubFolder={Uri.EscapeDataString(If(destSubFolder, ""))}" &
                            $"&destFileName={Uri.EscapeDataString(If(destFileName, ""))}" &
                            $"&overwrite={overwrite.ToString().ToLower()}"

        Dim response As HttpResponseMessage = Await _client.PostAsync(url, Nothing)

        If Not response.IsSuccessStatusCode Then
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Failed to move file: {errorMsg}")
        End If
    End Function

    ''' <summary>
    ''' Deletes an existing file remotely on the server.
    ''' </summary>
    ''' <param name="storeKey">The root storage directory key.</param>
    ''' <param name="subFolder">The subfolder path where the target file resides.</param>
    ''' <param name="fileName">The name of the file to delete.</param>
    Public Async Function DeleteFileAsync(storeKey As String, subFolder As String, fileName As String) As Task
        Dim url As String = $"Documents/DeleteFile?storeKey={Uri.EscapeDataString(storeKey)}" &
                            $"&subFolder={Uri.EscapeDataString(If(subFolder, ""))}" &
                            $"&fileName={Uri.EscapeDataString(fileName)}"

        Dim response As HttpResponseMessage = Await _client.PostAsync(url, Nothing)

        If Not response.IsSuccessStatusCode Then
            Dim errorMsg As String = Await response.Content.ReadAsStringAsync()
            Throw New Exception($"Failed to delete file: {errorMsg}")
        End If
    End Function

End Class
