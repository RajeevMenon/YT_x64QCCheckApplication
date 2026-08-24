Imports System.IO
Imports System.Threading.Tasks

Public Class DirectoryManager

    Private Const ApiBaseUrl As String = Link.DocClient
    Private Const StoreKey As String = Link.StoreKey
    Private ReadOnly _apiClient As New DocumentApiClient(ApiBaseUrl)

    ''' <summary>
    ''' Returns True if the directory exists. Pass path relative to StoreKey (e.g., "OrderTag" or "OrderTag\SubFolder").
    ''' </summary>
    Public Async Function ExistsAsync(relativePath As String) As Task(Of Boolean)
        Dim parentSubFolder As String = String.Empty
        Dim targetFolderName As String = String.Empty

        ParseDirectoryPath(relativePath, parentSubFolder, targetFolderName)

        If String.IsNullOrEmpty(targetFolderName) Then
            Return True ' Root StoreKey exists conceptually
        End If

        Try
            Dim subDirs As List(Of String) = Await _apiClient.GetSubDirectoriesAsync(StoreKey, parentSubFolder)
            Return subDirs IsNot Nothing AndAlso subDirs.Contains(targetFolderName, StringComparer.OrdinalIgnoreCase)
        Catch ex As Exception
            ' Log or debug print the actual error so it's not silent
            System.Diagnostics.Debug.WriteLine($"Error checking directory: {ex.Message}")
            Throw ' Re-throw the exception if you want your client code to handle it
        End Try
    End Function

    ''' <summary>
    ''' Synchronous wrapper for ExistsAsync.
    ''' </summary>
    Public Function Exists(relativePath As String) As Boolean
        Return Task.Run(Function() ExistsAsync(relativePath)).GetAwaiter().GetResult()
    End Function

    ''' <summary>
    ''' Creates a new directory. Pass path relative to StoreKey.
    ''' </summary>
    Public Async Function CreateDirectoryAsync(relativePath As String) As Task
        Dim parentSubFolder As String = String.Empty
        Dim targetFolderName As String = String.Empty

        ParseDirectoryPath(relativePath, parentSubFolder, targetFolderName)

        If Not String.IsNullOrEmpty(targetFolderName) Then
            Await _apiClient.CreateDirectoryAsync(StoreKey, parentSubFolder, targetFolderName)
        End If
    End Function

    ''' <summary>
    ''' Synchronous wrapper for CreateDirectoryAsync.
    ''' </summary>
    Public Sub CreateDirectory(relativePath As String)
        Task.Run(Function() CreateDirectoryAsync(relativePath)).GetAwaiter().GetResult()
    End Sub

    ''' <summary>
    ''' Deletes the directory and its contents recursively. Pass path relative to StoreKey.
    ''' </summary>
    Public Async Function DeleteAsync(relativePath As String) As Task
        Dim parentSubFolder As String = String.Empty
        Dim targetFolderName As String = String.Empty

        ParseDirectoryPath(relativePath, parentSubFolder, targetFolderName)

        If Not String.IsNullOrEmpty(targetFolderName) Then
            Await _apiClient.DeleteDirectoryAsync(StoreKey, parentSubFolder, targetFolderName, recursive:=True)
        End If
    End Function

    ''' <summary>
    ''' Synchronous wrapper for DeleteAsync.
    ''' </summary>
    Public Sub Delete(relativePath As String)
        Task.Run(Function() DeleteAsync(relativePath)).GetAwaiter().GetResult()
    End Sub

    Private Sub ParseDirectoryPath(relativePath As String, ByRef parentSubFolder As String, ByRef targetFolderName As String)
        If String.IsNullOrWhiteSpace(relativePath) Then
            Throw New ArgumentException("Path cannot be empty.", NameOf(relativePath))
        End If

        Dim cleanedPath = relativePath.TrimStart("/"c, "\"c)
        Dim parts As String() = cleanedPath.Split(New Char() {"/"c, "\"c}, StringSplitOptions.RemoveEmptyEntries)

        If parts.Length = 0 Then
            targetFolderName = String.Empty
            parentSubFolder = String.Empty
            Return
        End If

        targetFolderName = parts(parts.Length - 1)

        If parts.Length > 1 Then
            parentSubFolder = String.Join("\", parts, 0, parts.Length - 1)
        Else
            parentSubFolder = String.Empty
        End If
    End Sub
End Class