<div align="center">

## DSN Module and UserName


</div>

### Description

DSN Connection Creation or modification for Access or SQL Server

Included is a few functions et get UserName, ComputerName, DomainName, FullUserName (Requires NT or more)

By the way, GetServer, found at Microsoft
 
### More Info
 
DomainName and FullUserName does not work on win98, but who's still on 98???


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Patrick Garceau](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/patrick-garceau.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/patrick-garceau-dsn-module-and-username__1-58952/archive/master.zip)

### API Declarations

```
Option Explicit
Option Compare Text
Private Declare Function NetServerEnum Lib "netapi32" (ByVal ServerName As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal ServerType As Long, ByVal domain As Long, resume_handle As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nsize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nsize As Long) As Long
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv As Long) As Integer
Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal hEnv As Long, ByVal fDirection As Integer, ByVal szDSN$, ByVal cbDSNMax%, pcbDSN As Integer, ByVal szDescription As String, ByVal cbDescriptionMax As Integer, pcbDescription As Integer) As Integer
Private Declare Function SQLDrivers Lib "odbc32.dll" (ByVal hEnv As Long, ByVal fDirection As Integer, ByVal szDesc$, ByVal cbDescMax%, pcbDesc As Integer, ByVal szAttr$, ByVal cbAttrMax%, pcbAttr As Integer) As Integer
Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer
Private Const MAX_PREFERRED_LENGTH As Long = -1
Private Const NERR_SUCCESS As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234&
Public Enum SV_TYPE
  SV_TYPE_WORKSTATION = 1
  SV_TYPE_SERVER = 2
  SV_TYPE_SQLSERVER = 4
  SV_TYPE_DOMAIN_CTRL = 8
  SV_TYPE_DOMAIN_BAKCTRL = 16
  SV_TYPE_TIME_SOURCE = 32
  SV_TYPE_AFP = &H40
  SV_TYPE_NOVELL = &H80
  SV_TYPE_DOMAIN_MEMBER = &H100
  SV_TYPE_PRINTQ_SERVER = &H200
  SV_TYPE_DIALIN_SERVER = &H400
  SV_TYPE_XENIX_SERVER = &H800
  SV_TYPE_SERVER_UNIX = SV_TYPE_XENIX_SERVER
  SV_TYPE_NT = &H1000
  SV_TYPE_WFW = &H2000
  SV_TYPE_SERVER_MFPN = &H4000
  SV_TYPE_SERVER_NT = &H8000
  SV_TYPE_POTENTIAL_BROWSER = &H10000
  SV_TYPE_BACKUP_BROWSER = &H20000
  SV_TYPE_MASTER_BROWSER = &H40000
  SV_TYPE_DOMAIN_MASTER = &H80000
  SV_TYPE_SERVER_OSF = &H100000
  SV_TYPE_SERVER_VMS = &H200000
  SV_TYPE_WINDOWS = &H400000       'Windows95 and above
  SV_TYPE_DFS = &H800000         'Root of a DFS tree
  SV_TYPE_CLUSTER_NT = &H1000000     'NT Cluster
  SV_TYPE_TERMINALSERVER = &H2000000   'Terminal Server
  SV_TYPE_DCE = &H10000000         'IBM DSS
  SV_TYPE_ALTERNATE_XPORT = &H20000000   'rtn alternate transport
  SV_TYPE_LOCAL_LIST_ONLY = &H40000000   'rtn local only
  SV_TYPE_DOMAIN_ENUM = &H80000000
  SV_TYPE_ALL = &HFFFFFFFF
End Enum
Private Const SV_PLATFORM_ID_OS2    As Long = 400
Private Const SV_PLATFORM_ID_NT    As Long = 500
Private Type SERVER_INFO_100
 sv100_platform_id As Long
 sv100_name As Long
End Type
Public Enum ODBC_TYPE
  ODBC_USER_DNS = 1
  ODBC_SYSTEM_DNS = 2
End Enum
Public Enum DSN_DATABASE_TYPE
  MICROSOFT_ACCESS = 1
  MICROSOFT_SQL_SERVER = 2
End Enum
'Constants for SQLConfigDataSource.
Private Const ODBC_ADD_DSN = 1    ' Add user data source
Private Const ODBC_CONFIG_DSN = 2   ' Configure (edit) user data Source
Private Const ODBC_REMOVE_DSN = 3   ' Remove user data source
Private Const ODBC_ADD_SYS_DSN = 4  ' Add system data source
Private Const ODBC_CONFIG_SYS_DSN = 5 ' Configure (edit) system data Source
Private Const ODBC_REMOVE_SYS_DSN = 6 ' Remove system data source
Private Const vbAPINull As Long = 0& ' NULL Pointer
'Constants for SQLDataSources.
Private Const SQL_SUCCESS As Long = 0
Private Const SQL_FETCH_NEXT = 1
Private Const SQL_FETCH_FIRST_USER = 2
Private Const SQL_FETCH_FIRST_SYSTEM = 32
```


### Source Code

```
Public Function CreateDSN(ODBCType As ODBC_TYPE, DBType As DSN_DATABASE_TYPE, pstrDSN As String, pstrDesc As String, pstrPath As String, Optional pstrSQLServer As String) As Boolean
  Dim lngRet As Long
  Dim strDriver As String
  Dim strAttributes As String
  Select Case DBType
    Case MICROSOFT_ACCESS
      strDriver = "Microsoft Access Driver (*.mdb)" & Chr(0)
      strAttributes = "DSN=" & pstrDSN & Chr(0)
      strAttributes = strAttributes & "Description=" & pstrDesc & Chr(0)
      strAttributes = strAttributes & "Uid=Admin" & Chr(0) & "pwd=" & Chr(0)
      strAttributes = strAttributes & "DBQ=" & pstrPath & Chr(0)
    Case MICROSOFT_SQL_SERVER
      strDriver = "SQL Server" & Chr(0)
      strAttributes = "DSN=" & pstrDSN & Chr(0)
      strAttributes = strAttributes & "Description=" & pstrDesc & Chr(0)
      strAttributes = strAttributes & "SERVER=" & pstrSQLServer & Chr(0)
      strAttributes = strAttributes & "DATABASE=" & pstrPath & Chr(0)
      strAttributes = strAttributes & "Trusted_Connection=Yes" & Chr(0)
      '"SERVER=MySQL\0ADDRESS=MyServer\0NETWORK=dbmssocn\0"
  End Select
  If ODBCType = ODBC_USER_DNS Then
    lngRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)
  Else
    lngRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, strDriver, strAttributes)
  End If
  CreateDSN = (lngRet = 1)
End Function
Public Function ModifyDSN(ODBCType As ODBC_TYPE, DBType As DSN_DATABASE_TYPE, pstrDSN As String, pstrDesc As String, pstrPath As String, Optional pstrSQLServer As String) As Boolean
  Dim lngRet As Long
  Dim strDriver As String
  Dim strAttributes As String
  Select Case DBType
    Case MICROSOFT_ACCESS
      strDriver = "Microsoft Access Driver (*.mdb)" & Chr(0)
      strAttributes = "DSN=" & pstrDSN & Chr(0)
      strAttributes = strAttributes & "Description=" & pstrDesc & Chr(0)
      strAttributes = strAttributes & "Uid=Admin" & Chr(0) & "pwd=" & Chr(0)
      strAttributes = strAttributes & "DBQ=" & pstrPath & Chr(0)
    Case MICROSOFT_SQL_SERVER
      strDriver = "SQL Server" & Chr(0)
      strAttributes = "DSN=" & pstrDSN & Chr(0)
      strAttributes = strAttributes & "Description=" & pstrDesc & Chr(0)
      strAttributes = strAttributes & "SERVER=" & pstrSQLServer & Chr(0)
      strAttributes = strAttributes & "DATABASE=" & pstrPath & Chr(0)
      strAttributes = strAttributes & "Trusted_Connection=Yes" & Chr(0)
  End Select
  If ODBCType = ODBC_USER_DNS Then
    lngRet = SQLConfigDataSource(vbAPINull, ODBC_CONFIG_DSN, strDriver, strAttributes)
  Else
    lngRet = SQLConfigDataSource(vbAPINull, ODBC_CONFIG_SYS_DSN, strDriver, strAttributes)
  End If
  ModifyDSN = (lngRet = 1)
End Function
Public Function DeleteDSN(ODBCType As ODBC_TYPE, DBType As DSN_DATABASE_TYPE, pstrDSN As String) As Boolean
  Dim lngRet As Long
  Dim strDriver As String
  Dim strAttributes As String
  Select Case DBType
    Case MICROSOFT_ACCESS
      strDriver = "Microsoft Access Driver (*.mdb)" & Chr(0)
    Case MICROSOFT_SQL_SERVER
      strDriver = "SQL Server" & Chr(0)
  End Select
  strAttributes = "DSN=" & pstrDSN & Chr(0)
  If ODBCType = ODBC_USER_DNS Then
    lngRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)
  Else
    lngRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_SYS_DSN, strDriver, strAttributes)
  End If
  DeleteDSN = (lngRet = 1)
End Function
Public Function DetectDSN(ODBCType As ODBC_TYPE, pstrDSNName As String) As Boolean
  Dim intRet As Integer
  Dim strDSN As String
  Dim strDriver As String
  Dim intDSNLen As Integer
  Dim intDriverLen As Integer
  Dim lngEnvHandle As Long
  Dim blnFound As Boolean
  blnFound = False
  pstrDSNName = Trim$(pstrDSNName)
  intRet = SQLAllocEnv(lngEnvHandle)
  strDSN = Space(1024)
  strDriver = Space(1024)
  If ODBCType = ODBC_USER_DNS Then
    intRet = SQLDataSources(lngEnvHandle, SQL_FETCH_FIRST_USER, strDSN, 1024, intDSNLen, strDriver, 1024, intDriverLen)
  Else
    intRet = SQLDataSources(lngEnvHandle, SQL_FETCH_FIRST_SYSTEM, strDSN, 1024, intDSNLen, strDriver, 1024, intDriverLen)
  End If
  If intRet = SQL_SUCCESS Then
    If Trim$(strDSN) <> "" Then
      strDSN = Mid$(strDSN, 1, intDSNLen)
      If Trim$(strDSN) = pstrDSNName Then
        blnFound = True
      End If
    End If
    Do Until (intRet <> SQL_SUCCESS) Or blnFound
      strDSN = Space(1024)
      strDriver = Space(1024)
      intRet = SQLDataSources(lngEnvHandle, SQL_FETCH_NEXT, strDSN, 1024, intDSNLen, strDriver, 1024, intDriverLen)
      If Trim$(strDSN) <> "" Then
        strDSN = Mid$(strDSN, 1, intDSNLen)
        If Trim$(strDSN) = pstrDSNName Then
          blnFound = True
        End If
      End If
    Loop
  End If
  intRet = SQLFreeEnv(lngEnvHandle)
  DetectDSN = blnFound
End Function
Private Function GetServers(Optional ServerType As SV_TYPE = SV_TYPE_ALL) As String
 'lists all servers of the specified type
 'that are visible in a domain.
  Dim sDomain As String
  Dim bufptr     As Long
  Dim dwEntriesread  As Long
  Dim dwTotalentries As Long
  Dim dwResumehandle As Long
  Dim se100      As SERVER_INFO_100
  Dim success     As Long
  Dim nStructSize   As Long
  Dim cnt       As Long
  Dim St       As String
  nStructSize = LenB(se100)
 'Call passing MAX_PREFERRED_LENGTH to have the
 'API allocate required memory for the return values.
 '
 'The call is enumerating all machines on the
 'network (SV_TYPE_ALL); however, by Or'ing
 'specific bit masks for defined types you can
 'customize the returned data. For example, a
 'value of 0x00000003 combines the bit masks for
 'SV_TYPE_WORKSTATION (0x00000001) and
 'SV_TYPE_SERVER (0x00000002).
 '
 'dwServerName must be Null. The level parameter
 '(100 here) specifies the data structure being
 'used (in this case a SERVER_INFO_100 structure).
 '
 'The domain member is passed as Null, indicating
 'machines on the primary domain are to be retrieved.
 'If you decide to use this member, pass
 'StrPtr("domain name"), not the string itself.
  success = NetServerEnum(0&, _
              100, _
              bufptr, _
              MAX_PREFERRED_LENGTH, _
              dwEntriesread, _
              dwTotalentries, _
              ServerType, _
              0&, _
              dwResumehandle)
 'if all goes well
  If success = NERR_SUCCESS And _
   success <> ERROR_MORE_DATA Then
  'loop through the returned data, adding each
  'machine to the list
   For cnt = 0 To dwEntriesread - 1
    'get one chunk of data and cast
    'into an SERVER_INFO_100 struct
    'in order to add the name to a list
     CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize
     St = St & IIf(St = "", "", vbCrLf) & GetPointerToByteStringW(se100.sv100_name)
   Next
  End If
 'clean up regardless of success
  Call NetApiBufferFree(bufptr)
 'return entries as sign of success
  GetServers = St
End Function
Private Function GetPointerToByteStringW(ByVal dwData As Long) As String
  Dim tmp() As Byte
  Dim tmplen As Long
  If dwData <> 0 Then
   tmplen = lstrlenW(dwData) * 2
   If tmplen <> 0 Then
     ReDim tmp(0 To (tmplen - 1)) As Byte
     CopyMemory tmp(0), ByVal dwData, tmplen
     GetPointerToByteStringW = tmp
   End If
  End If
End Function
Function CurrentPrimaryDomainController() As String
  CurrentPrimaryDomainController = GetServers(SV_TYPE_DOMAIN_CTRL)
End Function
Function CurrentLogonUserName(Optional ByVal sUser As String = "") As String
  Dim sNom As String
  Dim sUserName As String
  Dim sPrenom As String
  On Error GoTo CurrentLogonUserName_Err
  If sUser = "" Then sUser = CurrentLogonUser()
  Dim MyObj As Object
  Set MyObj = GetObject("WinNT://" & CurrentPrimaryDomainController() & "/" & sUser & ",user")
  sUserName = MyObj.Fullname
  If InStr(sUserName, ",") > 0 Then
    sNom = Mid$(sUserName, 1, InStr(sUserName, ",") - 1)
    sNom = Trim$(sNom)
    sPrenom = Mid$(sUserName, InStr(sUserName, ",") + 1)
    sPrenom = Trim$(sPrenom)
    If sPrenom <> "" Then
      sUserName = sPrenom & " " & sNom
    Else
      sUserName = sNom
    End If
    sUserName = Trim$(sUserName)
  End If
CurrentLogonUserName_Err:
  If Err.Number <> 0 Then Err.Clear
  Set MyObj = Nothing
  If sUserName = "" Then sUserName = sUser
  CurrentLogonUserName = sUserName
End Function
Function CurrentLogonUser() As String
  Dim UserLoginName As String
  UserLoginName = Space(200)
  Call GetUserName(UserLoginName, 200)
  UserLoginName = Trim$(UserLoginName)
  UserLoginName = Mid$(UserLoginName, 1, Len(UserLoginName) - 1)
  CurrentLogonUser = UCase$(UserLoginName)
End Function
Function CurrentComputerName() As String
  Dim St As String
  St = Space(1024)
  Call GetComputerName(St, 1024)
  CurrentComputerName = Mid$(St, 1, InStr(St, Chr(0)) - 1)
End Function
```

