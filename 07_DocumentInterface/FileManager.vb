Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms.LinkLabel

Public Class FileManager

    Private Const ApiBaseUrl As String = Link.DocClient
    Private Const StoreKey As String = Link.StoreKey
    Private ReadOnly _apiClient As New DocumentApiClient(ApiBaseUrl)

    ''' <summary>
    ''' Returns True if the file exists. Pass path relative to StoreKey (e.g., "OrderTag\report.pdf").
    ''' </summary>
    Public Async Function ExistsAsync(relativePath As String) As Task(Of Boolean)
        Dim subFolder As String = String.Empty
        Dim fileName As String = String.Empty

        ParsePath(relativePath, subFolder, fileName)

        Try
            Dim files As List(Of String) = Await _apiClient.SearchFilesAsync(StoreKey, subFolder, fileName)
            Return files IsNot Nothing AndAlso files.Contains(fileName, StringComparer.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Synchronous wrapper for ExistsAsync.
    ''' </summary>
    Public Function Exists(relativePath As String) As Boolean
        Return Task.Run(Function() ExistsAsync(relativePath)).GetAwaiter().GetResult()
    End Function

    ''' <summary>
    ''' Deletes the file. Pass path relative to StoreKey (e.g., "OrderTag\report.pdf").
    ''' </summary>
    Public Async Function DeleteAsync(relativePath As String) As Task
        Dim subFolder As String = String.Empty
        Dim fileName As String = String.Empty

        ParsePath(relativePath, subFolder, fileName)
        Await _apiClient.DeleteFileAsync(StoreKey, subFolder, fileName)
    End Function

    ''' <summary>
    ''' Synchronous wrapper for DeleteAsync.
    ''' </summary>
    Public Sub Delete(relativePath As String)
        Task.Run(Function() DeleteAsync(relativePath)).GetAwaiter().GetResult()
    End Sub

    ''' <summary>
    ''' Searches a subfolder for files matching a search pattern and returns 
    ''' an array of file paths including their subfolder path relative to the DocStore folder.
    ''' </summary>
    Public Async Function GetFilesAsync(subFolder As String, searchPattern As String) As Task(Of String())
        Try
            Dim cleanedSubFolder = If(subFolder, "").TrimStart("\"c, "/"c)
            Dim fileNames As List(Of String) = Await _apiClient.SearchFilesAsync(StoreKey, cleanedSubFolder, searchPattern)

            If fileNames Is Nothing Then
                Return New String() {}
            End If

            Dim results As New List(Of String)
            For Each fileName In fileNames
                If String.IsNullOrEmpty(cleanedSubFolder) Then
                    results.Add(fileName)
                Else
                    results.Add(Path.Combine(cleanedSubFolder, fileName))
                End If
            Next

            Return results.ToArray()
        Catch
            Return New String() {}
        End Try
    End Function

    ''' <summary>
    ''' Synchronous wrapper for GetFilesAsync.
    ''' </summary>
    Public Function GetFiles(subFolder As String, searchPattern As String) As String()
        Return Task.Run(Function() GetFilesAsync(subFolder, searchPattern)).GetAwaiter().GetResult()
    End Function

    Private Sub ParsePath(relativePath As String, ByRef subFolder As String, ByRef fileName As String)
        If String.IsNullOrWhiteSpace(relativePath) Then
            Throw New ArgumentException("Path cannot be empty.", NameOf(relativePath))
        End If

        Dim cleanedPath = relativePath.TrimStart("\"c, "/"c)
        Dim parts As String() = cleanedPath.Split(New Char() {"\"c, "/"c}, StringSplitOptions.RemoveEmptyEntries)

        If parts.Length = 0 Then
            Throw New ArgumentException("Invalid path format.", NameOf(relativePath))
        End If

        fileName = parts(parts.Length - 1)

        If parts.Length > 1 Then
            subFolder = String.Join("/", parts, 0, parts.Length - 1)
        Else
            subFolder = String.Empty
        End If
    End Sub

    ''' <summary>
    ''' Downloads a file from the server via the API into the Windows Temp folder and returns the local path. 
    ''' Pass path relative to StoreKey (e.g., "OrderTag\report.pdf").
    ''' </summary>
    Public Async Function GetLocalViewPathAsync(relativePath As String) As Task(Of String)
        Dim subFolder As String = String.Empty
        Dim fileName As String = String.Empty

        Try
            ParsePath(relativePath, subFolder, fileName)

            ' Calls your DocumentApiClient method under the hood
            Return Await _apiClient.DownloadToTempFolderAsync(StoreKey, subFolder, fileName)

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error in GetLocalViewPathAsync for '{relativePath}': {ex.Message}")
            Throw New Exception($"Failed to download file '{relativePath}' from server: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' Synchronous wrapper for GetLocalViewPathAsync.
    ''' </summary>
    Public Function GetLocalViewPath(relativePath As String) As String
        Return Task.Run(Function() GetLocalViewPathAsync(relativePath)).GetAwaiter().GetResult()
    End Function

    ''' <summary>
    ''' Saves/uploads a local file to the server via the API using SaveMultipleFilesAsync.
    ''' Pass relative destination directory path (e.g., "OrderTag\SubFolder") and local file path.
    ''' </summary>
    Public Async Function SaveAsync(relativeDestinationFolder As String, localFilePath As String, Optional overwrite As Boolean = True) As Task
        If String.IsNullOrWhiteSpace(localFilePath) Then
            Throw New ArgumentException("Local file path cannot be empty.", NameOf(localFilePath))
        End If

        If Not File.Exists(localFilePath) Then
            Throw New FileNotFoundException($"Local file not found: {localFilePath}")
        End If

        Dim filesList As New List(Of String) From {localFilePath}

        ' Clean up slashes to match API expectation
        Dim cleanedSubFolder = relativeDestinationFolder?.TrimStart("\"c, "/"c).Replace("\"c, "/"c)

        ' Calls the underlying API client
        Await _apiClient.SaveMultipleFilesAsync(StoreKey, cleanedSubFolder, overwrite, filesList)
    End Function

    ''' <summary>
    ''' Synchronous wrapper for saving a file using the async manager call.
    ''' </summary>
    Public Sub Save(relativeDestinationFolder As String, localFilePath As String, Optional overwrite As Boolean = True)
        Task.Run(Function() SaveAsync(relativeDestinationFolder, localFilePath, overwrite)).GetAwaiter().GetResult()
    End Sub
End Class