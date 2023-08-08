Imports System.ComponentModel

Public Class AppControl


    ''' <summary>
    ''' Read and Return version number from Version.txt file saved in the path variable
    ''' </summary>
    ''' <param name="VersionFilePath">Path to the Version.txt file</param>
    ''' <returns></returns>
    Public Shared Function GetVersion(ByVal VersionFilePath As String) As String
        Try
            If System.IO.File.Exists(System.IO.Path.Combine(VersionFilePath, "Version.txt")) Then
                Dim Lines = System.IO.File.ReadAllLines(System.IO.Path.Combine(VersionFilePath, "Version.txt"))
                Dim ver = Lines(Lines.Length - 1).Split(vbTab)(0)
                Return ver
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Read and Return Settings as per data in the path variable
    ''' </summary>
    ''' <param name="SettingsFilePath">Path to the Settings.txt file</param>
    ''' <returns></returns>
    Public Shared Function GetSettings(ByVal SettingsFilePath As String) As Settings
        Try
            If System.IO.File.Exists(System.IO.Path.Combine(SettingsFilePath, "Settings.txt")) Then
                Dim Lines = System.IO.File.ReadAllLines(System.IO.Path.Combine(SettingsFilePath, "Settings.txt"))
                Dim SettingVar As New Settings
                For Each Line In Lines
                    For Each prop As PropertyDescriptor In TypeDescriptor.GetProperties(SettingVar)
                        If prop.Name = Line.Split(vbTab)(0) Then
                            Dim ReadVal As Boolean = False
                            For Each Item In Line.Split(vbTab)
                                If Item = "=" Then
                                    ReadVal = True
                                End If
                                If ReadVal = True And Item <> "=" Then
                                    prop.SetValue(SettingVar, Item)
                                End If
                            Next

                        End If
                    Next
                Next
                Return SettingVar
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Read and Return Settings as per data in the path variable
    ''' </summary>
    ''' <param name="SettingsFilePath">Path to the Settings.txt file</param>
    ''' <returns></returns>
    Public Shared Function GetSettingsNew(ByVal SettingsFilePath As String) As Settings
        Try
            If System.IO.File.Exists(System.IO.Path.Combine(SettingsFilePath, "Settings_new.txt")) Then
                Dim Lines = System.IO.File.ReadAllLines(System.IO.Path.Combine(SettingsFilePath, "Settings_new.txt"))
                Dim SettingVar As New Settings
                For Each Line In Lines
                    For Each prop As PropertyDescriptor In TypeDescriptor.GetProperties(SettingVar)
                        If prop.Name = Line.Split(vbTab)(0) Then
                            Dim ReadVal As Boolean = False
                            For Each Item In Line.Split(vbTab)
                                If Item = "=" Then
                                    ReadVal = True
                                End If
                                If ReadVal = True And Item <> "=" Then
                                    prop.SetValue(SettingVar, Item)
                                End If
                            Next

                        End If
                    Next
                Next
                Return SettingVar
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Read and Return Settings as per data in the path variable
    ''' </summary>
    ''' <param name="SettingsFilePath">Path to the Settings.txt file</param>
    ''' <returns></returns>
    Public Shared Function GetSettingsLive(ByVal SettingsFilePath As String) As Settings
        Try
            If System.IO.File.Exists(System.IO.Path.Combine(SettingsFilePath, "YTA_LineSettings(LIVE).txt")) Then
                Dim Lines = System.IO.File.ReadAllLines(System.IO.Path.Combine(SettingsFilePath, "YTA_LineSettings(LIVE).txt"))
                Dim SettingVar As New Settings
                For Each Line In Lines
                    For Each prop As PropertyDescriptor In TypeDescriptor.GetProperties(SettingVar)
                        If prop.Name = Line.Split(vbTab)(0) Then
                            Dim ReadVal As Boolean = False
                            For Each Item In Line.Split(vbTab)
                                If Item = "=" Then
                                    ReadVal = True
                                End If
                                If ReadVal = True And Item <> "=" Then
                                    prop.SetValue(SettingVar, Item)
                                End If
                            Next

                        End If
                    Next
                Next
                Return SettingVar
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Read and Return Settings as per data in the path variable
    ''' </summary>
    ''' <param name="SettingsFilePath">Path to the Settings.txt file</param>
    ''' <returns></returns>
    Public Shared Function GetSettingsDomain(ByVal SettingsFilePath As String) As Settings
        Try
            If System.IO.File.Exists(System.IO.Path.Combine(SettingsFilePath, "YTA_LineSettings(DOMAIN).txt")) Then
                Dim Lines = System.IO.File.ReadAllLines(System.IO.Path.Combine(SettingsFilePath, "YTA_LineSettings(DOMAIN).txt"))
                Dim SettingVar As New Settings
                For Each Line In Lines
                    For Each prop As PropertyDescriptor In TypeDescriptor.GetProperties(SettingVar)
                        If prop.Name = Line.Split(vbTab)(0) Then
                            Dim ReadVal As Boolean = False
                            For Each Item In Line.Split(vbTab)
                                If Item = "=" Then
                                    ReadVal = True
                                End If
                                If ReadVal = True And Item <> "=" Then
                                    prop.SetValue(SettingVar, Item)
                                End If
                            Next

                        End If
                    Next
                Next
                Return SettingVar
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Read and Return Settings as per data in the path variable
    ''' </summary>
    ''' <param name="SettingsFilePath">Path to the Settings.txt file</param>
    ''' <returns></returns>
    Public Shared Function GetSettingsLocal(ByVal SettingsFilePath As String) As Settings
        Try
            If System.IO.File.Exists(System.IO.Path.Combine(SettingsFilePath, "YTA_LineSettings(TEST).txt")) Then
                Dim Lines = System.IO.File.ReadAllLines(System.IO.Path.Combine(SettingsFilePath, "YTA_LineSettings(TEST).txt"))
                Dim SettingVar As New Settings
                For Each Line In Lines
                    For Each prop As PropertyDescriptor In TypeDescriptor.GetProperties(SettingVar)
                        If prop.Name = Line.Split(vbTab)(0) Then
                            Dim ReadVal As Boolean = False
                            For Each Item In Line.Split(vbTab)
                                If Item = "=" Then
                                    ReadVal = True
                                End If
                                If ReadVal = True And Item <> "=" Then
                                    prop.SetValue(SettingVar, Item)
                                End If
                            Next

                        End If
                    Next
                Next
                Return SettingVar
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class
