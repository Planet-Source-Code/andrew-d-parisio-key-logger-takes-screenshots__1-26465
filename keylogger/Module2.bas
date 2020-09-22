Attribute VB_Name = "Module2"
Public Exist As Boolean
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1


Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long


Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long


Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long


Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
    
'After this has been added also add the following code to your module file


Public Sub SetStringValue(Hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim i As Long
    'Create the key
    i = RegCreateKey(Hkey, strPath, keyhand)
    'Set the value
    i = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    'Close the key
    i = RegCloseKey(keyhand)
End Sub

Public Sub DeleteStringValue(Hkey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim i As Long
    'Open the key
    i = RegOpenKey(Hkey, strPath, keyhand)
    'Delete the value
    i = RegDeleteValue(keyhand, strValue)
    'Close the key
    i = RegCloseKey(keyhand)
End Sub


