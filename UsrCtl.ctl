VERSION 5.00
Begin VB.UserControl UsrCtl 
   ClientHeight    =   4920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   ScaleHeight     =   4920
   ScaleWidth      =   6090
   ToolboxBitmap   =   "UsrCtl.ctx":0000
   Begin VB.Frame Frame1 
      Caption         =   "NT Domain Users/Groups"
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdGroup 
         Height          =   375
         Left            =   4440
         Picture         =   "UsrCtl.ctx":0312
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdGo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Height          =   375
         Left            =   4080
         Picture         =   "UsrCtl.ctx":0824
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.ListBox lstUsers 
         Height          =   2400
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtDomain 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Users :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Domain :"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
   End
End
Attribute VB_Name = "UsrCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' The following code is used to return all the NT user account and group names from your current
' domain,(Primary Domain Controller)

' Created this Control with the help of some people from PSC. (Can't remember their names)

' Good luck in using this code. Henk le Roux

Option Base 0

Private Declare Function NetShareDel Lib "netapi32.dll" (ByRef servername As Byte, ByRef netname As Byte, ByVal reserved As Long) As Long

Private Declare Function NetUserEnum Lib "netapi32.dll" (ByRef servername As Byte, ByVal level As Long, ByVal lFilter As Long, ByRef buffer As Long, ByVal prefmaxlen As Long, ByRef entriesread As Long, ByRef totalentries As Long, ByRef ResumeHandle As Long) As Long

Private Declare Function NetGroupEnumUsers Lib "netapi32.dll" Alias "NetGroupGetUsers" (ByRef servername As Byte, ByRef GroupName As Byte, ByVal level As Long, ByRef buffer As Long, ByVal prefmaxlen As Long, ByRef entriesread As Long, ByRef totalentries As Long, ByRef ResumeHandle As Long) As Long

Private Declare Function NetUserGetGroups Lib "netapi32.dll" (ByRef servername As Byte, ByRef username As Byte, ByVal level As Long, ByRef buffer As Long, ByVal prefmaxlen As Long, ByRef entriesread As Long, ByRef totalentries As Long) As Long

Private Declare Function NetQueryDisplayInformation Lib "netapi32.dll" (ByRef servername As Byte, ByVal level As Long, ByVal Index As Long, ByVal EntriesRequested As Long, ByVal PreferredMaximumLength As Long, ByRef ReturnedEntryCount As Long, ByRef SortedBuffer As Long) As Long

Private Declare Function NetUserGetInfo Lib "NETAPI32" (ByRef servername As Byte, ByRef username As Byte, ByVal level As Long, ByRef buffer As Long) As Long

Private Declare Function NetUserSetInfo Lib "NETAPI32" (ByRef servername As Byte, ByRef username As Byte, ByVal level As Long, ByRef buffer As TUser1006, ByRef parm_err As Long) As Long

Private Declare Function NetShareGetInfo Lib "NETAPI32" (ByRef servername As Byte, ByRef netname As Byte, ByVal level As Long, ByRef buffer As Long) As Long

Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long

Private Declare Function NetAPIBufferAllocate Lib "netapi32.dll" Alias "NetApiBufferAllocate" (ByVal ByteCount As Long, Ptr As Long) As Long

Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long

Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long

Private Declare Function StrToPtr Lib "kernel32" Alias "lstrcpyW" (ByVal Ptr As Long, Source As Byte) As Long

Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Private Declare Function NetGetDCName Lib "netapi32.dll" (ByRef servername As Byte, ByRef DomainName As Byte, ByRef buffer As Long) As Long

Private Declare Function NetGroupAddUser Lib "netapi32.dll" (ByRef servername As Byte, ByRef GroupName As Byte, ByRef username As Byte) As Long

Private Declare Function NetGroupDelUser Lib "netapi32.dll" (ByRef servername As Byte, ByRef GroupName As Byte, ByRef username As Byte) As Long


Const UF_SCRIPT = &H1
Const UF_ACCOUNTDISABLE = &H2
Const UF_HOMEDIR_REQUIRED = &H8
Const UF_LOCKOUT = &H10
Const UF_PASSWD_NOTREQD = &H20
Const UF_PASSWD_CANT_CHANGE = &H40

Const UF_TEMP_DUPLICATE_ACCOUNT = &H100
Const UF_NORMAL_ACCOUNT = &H200
Const UF_INTERDOMAIN_TRUST_ACCOUNT = &H800
Const UF_WORKSTATION_TRUST_ACCOUNT = &H1000
Const UF_SERVER_TRUST_ACCOUNT = &H2000

Const UF_DONT_EXPIRE_PASSWD = &H10000
Const UF_MNS_LOGON_ACCOUNT = &H20000

Const AF_OP_PRINT = 1
Const AF_OP_COMM = 2
Const AF_OP_SERVER = 4
Const AF_OP_ACCOUNTS = 8

Const DateFormat As String = "dd/mm/yyyy hh:nn:ss"

Private Type MungeLong
   x As Long
   Dummy As Integer
End Type

Private Type MungeInt
   XLo As Integer
   XHi As Integer
   Dummy As Integer
End Type

Private Type TUser1006
   ptrHomeDir As Long
End Type
      
Dim m_NTUserAccount As String

Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
          

Public Function GetPDC(Server As String, domain As String, PDC As String) As Long
   Dim Result As Long
   Dim SNArray() As Byte
   Dim DArray() As Byte
   Dim DCNPtr As Long
   Dim STRArray(100) As Byte

   SNArray = Server & vbNullChar
   DArray = domain & vbNullChar

   Result = NetGetDCName(SNArray(0), _
         DArray(0), DCNPtr)

   GetPDC = Result

   If Result = 0 Then
      Result = PtrToStr(STRArray(0), DCNPtr)
      PDC = Left(STRArray(), StrLen(DCNPtr))
   Else
      PDC = ""
   End If
   NetAPIBufferFree (DCNPtr)
End Function

      
Function QuickEnumerate(domain As String, level As Long, Size As Long, Data() As String) As Boolean
   Dim APIResult As Long
   Dim Result As Long
   Dim PDC As String
   Dim SNArray() As Byte
   Dim EntriesRequested As Long
   Dim PreferredMaximumLength As Long
   Dim ReturnedEntryCount As Long
   Dim SortedBuffer As Long
   Dim TempPtr As MungeLong
   Dim tempstr As MungeInt
   Dim STRArray(500) As Byte
   Dim i As Integer
   Dim Index As Long
   Dim NextIndex As Long
   Dim MoreData As Boolean
   Dim ArrayRoom As String
   Dim name As String
   Dim comment As String

   On Error GoTo RuntimeError
   
   ReDim Data(1 To 8000)
   ArrayRoom = 8000
   
   Result = GetPDC("", domain, PDC)
   If Result <> 0 Then GoTo HandleError
   Size = 0
   SNArray = PDC & vbNullChar
   Index = 0
   EntriesRequested = 500
   PreferredMaximumLength = 6000
   Do
      APIResult = NetQueryDisplayInformation(SNArray(0), level, Index, EntriesRequested, PreferredMaximumLength, ReturnedEntryCount, SortedBuffer)

      If APIResult <> 0 And APIResult <> 234 Then
         GoTo HandleError
      End If

      For i = 1 To ReturnedEntryCount
         Size = Size + 1
         If Size > ArrayRoom Then
            ArrayRoom = ArrayRoom + 2000
            ReDim Preserve Data(1 To ArrayRoom)
         End If
         Select Case level
            Case Is = 1

               Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 24, 2)
               Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 24 + 2, 2)
               LSet TempPtr = tempstr
               Result = PtrToStr(STRArray(0), TempPtr.x)
               Data(Size) = Left(STRArray, StrLen(TempPtr.x))

               Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 24 + 20, 2)
               Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 24 + 22, 2)
               LSet TempPtr = tempstr
               NextIndex = TempPtr.x

            Case Is = 2
               Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20, 2)
               Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 2, 2)
               LSet TempPtr = tempstr
               Result = PtrToStr(STRArray(0), TempPtr.x)
               name = Left(STRArray, StrLen(TempPtr.x))

               Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20 + 4, 2)
               Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 6, 2)
               LSet TempPtr = tempstr
               Result = PtrToStr(STRArray(0), TempPtr.x)
               comment = Left(STRArray, StrLen(TempPtr.x))

               Data(Size) = "1234567890123456789012"
               LSet Data(Size) = name
               Data(Size) = Data(Size) & comment

               Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20 + 16, 2)
               Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 18, 2)
               LSet TempPtr = tempstr
               NextIndex = TempPtr.x

            Case Is = 3
               Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20, 2)
               Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 2, 2)
               LSet TempPtr = tempstr
               Result = PtrToStr(STRArray(0), TempPtr.x)
               name = Left(STRArray, StrLen(TempPtr.x))

               Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20 + 4, 2)
               Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 6, 2)
               LSet TempPtr = tempstr
               Result = PtrToStr(STRArray(0), TempPtr.x)
               comment = Left(STRArray, StrLen(TempPtr.x))

               Data(Size) = "1234567890123456789012"
               LSet Data(Size) = name
               Data(Size) = Data(Size) & comment

               Result = PtrToInt(tempstr.XLo, SortedBuffer + (i - 1) * 20 + 16, 2)
               Result = PtrToInt(tempstr.XHi, SortedBuffer + (i - 1) * 20 + 18, 2)
               LSet TempPtr = tempstr
               NextIndex = TempPtr.x

         End Select
      Next i
      Result = NetAPIBufferFree(SortedBuffer)
      Index = NextIndex
   Loop Until APIResult = 0

   If Size > 0 Then
      ReDim Preserve Data(1 To Size)
   Else
      ReDim Preserve Data(1 To 1)
   End If

   QuickEnumerate = True
ExitHere:
   Exit Function
HandleError:
   On Error Resume Next
   QuickEnumerate = False
   GoTo ExitHere
RuntimeError:
   Resume HandleError
End Function


Private Sub cmdGo_Click()
Dim mbSuccess As Boolean
Dim Size As Long
Dim Data() As String
Dim i As Long

   lstUsers.Clear
   mbSuccess = QuickEnumerate(txtDomain, 1, Size, Data)
   If mbSuccess Then
      For i = 1 To Size
         lstUsers.AddItem (Data(i))
      Next i
      UsersShown = True
   Else
      MsgBox ("Error when trying to enumerate Users")
      UsersShown = False
   End If
End Sub


Private Sub cmdGroup_Click()
Dim mbSuccess As Boolean
Dim Size As Long
Dim Data() As String
Dim i As Long
   UsersShown = False
   lstUsers.Clear
   mbSuccess = QuickEnumerate(txtDomain, 3, Size, Data)
   If mbSuccess Then
      For i = 1 To Size
         lstUsers.AddItem (Data(i))
      Next i
   Else
      MsgBox ("Error when trying to enumerate groups")
   End If
End Sub

Private Sub lstUsers_Click()
    m_NTUserAccount = lstUsers.Text
End Sub


Public Property Get NTUserAccount() As String
    NTUserAccount = m_NTUserAccount
End Property


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Frame1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Frame1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Frame1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BackColor = Frame1.BackColor
End Sub


Private Sub UserControl_Resize()
    Frame1.Move 0, 0, ScaleWidth, ScaleHeight
    lstUsers.Move 900, 800, Abs(ScaleWidth - 1200), Abs(ScaleHeight - 1000)
    txtDomain.Move 900, 300, Abs(ScaleWidth - 1200 - cmdGo.Width - cmdGroup.Width), cmdGo.Height
    cmdGo.Left = txtDomain.Left + txtDomain.Width
    cmdGroup.Left = cmdGo.Left + cmdGo.Width
    cmdGo.Top = txtDomain.Top
    cmdGroup.Top = txtDomain.Top
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", Frame1.BackColor, &H8000000F)
End Sub

