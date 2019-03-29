Public Class ClsReg
    '    Public Class clsAppSettings
    Private _configFileName As String 'local var to hold the config file name
    Private _configFileType As Config 'local var to hold the config file type (private or shared)
    Private _AlternatePath As String = String.Empty

    'config file options
    Public Enum Config
        SharedFile  'all users use the same config file
        PrivateFile 'each user has their own config file
        SystemFile 'all users use the same config file, config file put at System folder
        ApplicationFile '//Application folder
    End Enum

    'constructor
    Public Sub New(ByVal ConfigFileType As Config)
        _configFileType = ConfigFileType 'remember this setting

        InitializeConfigFile() 'setup the filename and location
    End Sub


    Public Sub New()

    End Sub

    'initialize the apps config file, create it if it doesn't exist
    Private Sub InitializeConfigFile()
        Dim sb As New System.Text.StringBuilder()
        'build the path\filename depending on the location of the config file
        Select Case _configFileType
            Case Config.PrivateFile 'each user has their own personal settings
                'use "\documents and settings\username\application data" for the config file directory
                sb.Append(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
            Case Config.SharedFile 'all users share the same settings
                'use "\documents and settings\All Users\application data" for the config file directory
                sb.Append(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
            Case Config.SystemFile
                'use "\Winnt\System32" for the config file directory
                sb.Append(Environment.GetFolderPath(Environment.SpecialFolder.System))
            Case Config.ApplicationFile
                '// application path
                sb.Append(getAppPath)
        End Select

        'add the product name
        sb.Append("\")
        sb.Append(Application.ProductName)

        'create the directory if it isn't there
        If Not IO.Directory.Exists(sb.ToString) Then
            IO.Directory.CreateDirectory(sb.ToString)
        End If

        'finish building the file name
        sb.Append("\")
        sb.Append(Application.ProductName)
        sb.Append(".config")

        _configFileName = sb.ToString 'completed config filename

        'if the file doesn't exist, create a blank xml
        If Not IO.File.Exists(_configFileName) Then
            Dim fn As New IO.StreamWriter(IO.File.Open(_configFileName, IO.FileMode.Create))
            fn.WriteLine("<?xml version=""1.0"" encoding=""utf-8""?>")
            fn.WriteLine("<configuration>")
            fn.WriteLine("  <appSettings>")
            fn.WriteLine("    <!--   User application and configured property settings go here.-->")
            fn.WriteLine("    <!--   Example: <add key=""settingName"" value=""settingValue""/> -->")
            fn.WriteLine("  </appSettings>")
            fn.WriteLine("</configuration>")
            fn.Close() 'all done
        End If
    End Sub

    Public Function GetConfigFileLocation() As String
        Return _configFileName
    End Function

    Private Sub DeleteRegistry(ByVal Key As String)
        Try
            Dim regKey As Microsoft.Win32.RegistryKey
            Dim Ret As String

            regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
            regKey.CreateSubKey(Application.ProductName)
            regKey.Close()

            regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\" & Application.ProductName, True)
            regKey.DeleteValue(Key)
            regKey.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Function getRegistry(ByVal key As String, Optional ByVal defValue As String = "") As String
        Try
            Dim regKey As Microsoft.Win32.RegistryKey
            Dim Ret As String

            regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
            regKey.CreateSubKey(Application.ProductName)
            regKey.Close()

            'regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\" & Application.ProductName, True)
            regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Office2011", True)


            'regKey.SetValue(key, value)
            Ret = regKey.GetValue(key, defValue)

            regKey.Close()

            Return Ret

        Catch ex As Exception

        End Try

    End Function

    'get an application setting by key value
    Public Function GetSetting(ByVal key As String, Optional ByVal defValue As String = "") As String

        Dim strResult As String

        strResult = getRegistry(key, defValue)
        If strResult <> "" Then
            Return strResult
        Else
            If defValue <> "" Then
                Return defValue
            Else
                Return ""
            End If
        End If

    End Function

    Private Sub SaveRegistry(ByVal Key As String, ByVal value As String)
        Try
            Dim regKey As Microsoft.Win32.RegistryKey
            Dim Ret As String

            regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE", True)
            'regKey.CreateSubKey(Application.ProductName)
            regKey.CreateSubKey("Office2011")
            regKey.Close()

            'regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\" & Application.ProductName, True)
            regKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("Software\Office2011", True)
            regKey.SetValue(Key, value)
            Ret = regKey.GetValue(Key, "")
            regKey.Close()
        Catch ex As Exception

        End Try
    End Sub

    'save an application setting, takes a key and a value
    Public Sub SaveSetting(ByVal key As String, ByVal value As String)

        SaveRegistry(key, value)

    End Sub

    'delete an application setting, takes a key and a value
    Public Sub DeleteSetting(ByVal key As String)

        DeleteRegistry(key)

    End Sub

    Private Function getAppPath() As String
        Dim s As String
        If _AlternatePath.Length = 0 Then
            s = IO.Path.GetDirectoryName( _
               Reflection.Assembly.GetExecutingAssembly.Location())
        Else
            s = IO.Path.GetDirectoryName(_AlternatePath)
        End If
        Return s
    End Function

End Class
