Attribute VB_Name = "modRegistry"
Option Explicit

Global Const REG_SZ As Long = 1
Global Const REG_DWORD As Long = 4

Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

Global Const KEY_ALL_ACCESS = &H3F

Global Const REG_OPTION_NON_VOLATILE = 0

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
   
Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                 ByVal lpSubKey As String, _
                                 ByVal Reserved As Long, _
                                 ByVal lpClass As String, _
                                 ByVal dwOptions As Long, _
                                 ByVal samDesired As Long, _
                                 ByVal lpSecurityAttributes As Long, _
                                 phkResult As Long, _
                                 lpdwDisposition As Long) As Long
   
Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                               ByVal lpSubKey As String, _
                               ByVal ulOptions As Long, _
                               ByVal samDesired As Long, _
                               phkResult As Long) As Long

Declare Function RegQueryValueExString Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                  ByVal lpValueName As String, _
                                  ByVal lpReserved As Long, _
                                  lpType As Long, _
                                  ByVal lpData As String, _
                                  lpcbData As Long) As Long

Declare Function RegQueryValueExLong Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                  ByVal lpValueName As String, _
                                  ByVal lpReserved As Long, _
                                  lpType As Long, _
                                  lpData As Long, _
                                  lpcbData As Long) As Long
   
Declare Function RegQueryValueExNULL Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                  ByVal lpValueName As String, _
                                  ByVal lpReserved As Long, _
                                  lpType As Long, _
                                  ByVal lpData As Long, _
                                  lpcbData As Long) As Long

Declare Function RegSetValueExString Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, _
                                ByVal lpValueName As String, _
                                ByVal Reserved As Long, _
                                ByVal dwType As Long, _
                                ByVal lpValue As String, _
                                ByVal cbData As Long) As Long

Declare Function RegSetValueExLong Lib "advapi32.dll" _
        Alias "RegSetValueExA" (ByVal hKey As Long, _
                                ByVal lpValueName As String, _
                                ByVal Reserved As Long, _
                                ByVal dwType As Long, _
                                lpValue As Long, _
                                ByVal cbData As Long) As Long
                                      
Public Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
  '
  ' Creating A New Registry Key
  '
  ' CreateNewKey takes the name of the key to create, and the constant
  ' representing the predefined key to create the key under. The call
  ' to RegCreateKeyEx doesn't take advantage of the security mechanisms.
  '
  ' Examples:
  '
  '   CreateNewKey "TestKey", HKEY_CURRENT_USER
  '
  '   - will create a key called TestKey immediately under HKEY_CURRENT_USER.
  '
  '   CreateNewKey "TestKey\SubKey1\SubKey2", HKEY_LOCAL_MACHINE
  '
  '   - will create three-nested keys beginning with TestKey immediately under
  '     HKEY_LOCAL_MACHINE, SubKey1 subordinate to TestKey, and SubKey3 under SubKey2.
  '
  Dim hNewKey As Long         'handle to the new key
  Dim lRetVal As Long         'result of the RegCreateKeyEx function
  lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
  RegCloseKey (hNewKey)
End Sub

Public Sub SetKeyValue(lPredefinedKey As Long, _
                        sKeyName As String, _
                        sValueName As String, _
                        vValueSetting As Variant, _
                        lValueType As Long)
  '
  ' Setting / Modifying a Registry Value
  '
  ' SetKeyValue takes the key that the value will be associated with, the
  ' name of the value, the setting of the value, and the type of the value
  ' (the SetValueEx function only supports REG_SZ and REG_DWORD,  but this
  ' can be modified if necessary). Specifying a new value for an existing
  ' sValueName modifies the current setting of that value.
  '
  ' Example:
  '
  '   SetKeyValue HKEY_CURRENT_USER, "TestKey\SubKey1", "StringValue", "Hello", REG_SZ
  '
  '   - creates a value of type REG_SZ called "SubKey1" with the setting of "Hello".
  '     This value is associated with the key SubKey1 of "TestKey."  In this case,
  '     "TestKey" is a subkey of HKEY_CURRENT_USER. This call fails if "TestKey\SubKey1"
  '     does not exist. To avoid this problem, use a call to RegCreateKeyEx instead
  '     of a call to RegOpenKeyEx. RegCreateKeyEx opens a specified key if it already
  '     exists.
  '
  Dim lRetVal As Long         'result of the SetValueEx function
  Dim hKey As Long            'handle of open key
  ' open the specified key
  lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
  lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
  RegCloseKey (hKey)
End Sub
   
Public Function QueryValue(lPredefinedKey As Long, _
                           sKeyName As String, _
                           sValueName As String) As Variant
  '
  ' Querying a Registry Value
  '
  ' The next procedure can be used to ascertain the setting of an existing value.
  ' QueryValue takes the name of the key and the name of a value associated with
  ' that key and displays a message box with the corresponding value. It uses a
  ' call to the QueryValueEx wrapper function defined below, which supports only '
  ' REG_SZ and REG_DWORD types:
  '
  ' Example:
  '
  '   QueryValue HKEY_CURRENT_USER, "TestKey\SubKey1", "StringValue"
  '
  '   - displays a message box with the current setting of the "StringValue"
  '     value, and assumes that "StringValue" exists in the "TestKey\SubKey1" key.
  '     If the Value that you query does not exist, then QueryValue returns an
  '     error code of 2 - 'ERROR_BADKEY'.
  '
  Dim lRetVal As Long         'result of the API functions
  Dim hKey As Long            'handle of opened key
  Dim vValue As Variant       'setting of queried value
  lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
  lRetVal = QueryValueEx(hKey, sValueName, vValue)
  QueryValue = UCase(vValue)
  RegCloseKey (hKey)
End Function
   
Public Function SetValueEx(ByVal hKey As Long, _
                           sValueName As String, _
                           lType As Long, _
                           vValue As Variant) As Long
  Dim lValue As Long
  Dim sValue As String
  Select Case lType
         Case REG_SZ
              sValue = vValue & Chr$(0)
              SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
         Case REG_DWORD
              lValue = vValue
              SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
  End Select
End Function
   
Function QueryValueEx(ByVal lhKey As Long, _
                      ByVal szValueName As String, _
                      vValue As Variant) As Long
  Dim cch As Long
  Dim lrc As Long
  Dim lType As Long
  Dim lValue As Long
  Dim sValue As String

  On Error GoTo QueryValueExError

  ' Determine the size and type of data to be read
  lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
  
  If lrc <> ERROR_NONE Then Error 5
  
  Select Case lType
         ' For strings
         Case REG_SZ:
              sValue = String(cch, 0)
              lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
              If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
              Else
                 vValue = Empty
              End If
          ' For DWORDS
          Case REG_DWORD:
               lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
          Case Else
               'all other data types not supported
               lrc = -1
  End Select

QueryValueExExit:
  QueryValueEx = lrc
  Exit Function

QueryValueExError:
  Resume QueryValueExExit
   
End Function

