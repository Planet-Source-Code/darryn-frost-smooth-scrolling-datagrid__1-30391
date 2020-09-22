<div align="center">

## Smooth Scrolling DataGrid


</div>

### Description

This code allows you to have the smooth-scrolling effect seen in better applications on your datagrids(could also apply to other scroll bars in VB)

When you grab the trackbar and move it, VB doesn't do anything until you let go. Or if you click on the trackbar itself, the grid just jumps.

This code shows you how to change these effects so that the grid(or text) will scroll smoothly as you drag or click.
 
### More Info
 
I used the "MsgBlaster" .Bas module and type library for the subclassing, they are available free on the net, and I will upload it in a seperate Listing here right after this one


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Darryn Frost](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/darryn-frost.md)
**Level**          |Intermediate
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/darryn-frost-smooth-scrolling-datagrid__1-30391/archive/master.zip)

### API Declarations

```
Public Type SCROLLINFO
 cbSize As Long
 fMask As Long
 nMin As Long
 nMax As Long
 nPage As Long
 nPos As Long
 nTrackPos As Long
End Type
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
' Scroll Bar Commands
Public Const SB_PAGEUP = 2
Public Const SB_PAGEDOWN = 3
Public Const SB_THUMBTRACK = 5
' Scroll Bar Constants
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)
```


### Source Code

```

'This is for a form with a datagrid
Option Explicit
Private m_Grid_Subclassed As Boolean
Private Const msCustomMessageName As String = "MsgBlasterCustomMessage"
Private mlCustomMessageID As Long
Private rglMsgIDs() As Long
Implements IMsgTarget
Private Sub Form_Load()
'Open a recordset and bind the grid to it here
 Call SubClassGrid
End Sub
Private Sub SubClassGrid()
On Error GoTo SubClass_Error
 If Not m_Grid_Subclassed = True Then
  'To prevent it from trying again, since that can cause problems
  m_Grid_Subclassed = True
  ' Register our custom message to get the message id.
  mlCustomMessageID = RegisterWindowMessage(msCustomMessageName)
  'The windows messages we are interested in are WM_VSCROLL and WM_HSCROLL
  ReDim rglMsgIDs(1 To 3) As Long
  rglMsgIDs(1) = WM_VSCROLL
  rglMsgIDs(2) = WM_HSCROLL
  rglMsgIDs(3) = mlCustomMessageID
  MsgBlaster.SubclassWindow DataGrid1.hWnd, Me, rglMsgIDs
 End If
Exit Sub
SubClass_Error:
 'Since this is not a critical error, just ignore it for the user
 Exit Sub
End Sub
Private Function IMsgTarget_OnMsg( _
 ByVal hWnd As Long, _
 ByVal msg As Long, _
 ByVal wParam As Long, _
 ByVal lParam As Long) As Long
 Dim LOBYTE As Integer
 Dim HIBYTE As Integer
 Dim nRes As Long
 Dim fEat As Boolean
 Dim intAction As Integer
 Dim pVert As Boolean
On Error GoTo SubClass_Error
  'If this is False, the message will be passed along the chain
  'If it is True, it will not be passed on
  fEat = False
  intAction = 0
  Select Case msg
    Case WM_VSCROLL
      nRes = MsgBlaster.GetHiLoByte(wParam, LOBYTE, HIBYTE)
      If LOBYTE = SB_THUMBTRACK Or LOBYTE = SB_PAGEDOWN Or LOBYTE = SB_PAGEUP Then
       fEat = True
       intAction = 1
       pVert = True
      End If
    Case WM_HSCROLL
      nRes = MsgBlaster.GetHiLoByte(wParam, LOBYTE, HIBYTE)
      If LOBYTE = SB_THUMBTRACK Then
       fEat = True
       intAction = 1
       pVert = False
      End If
    Case mlCustomMessageID
     'lstLog.AddItem msCustomMessageName & vbTab & "wParam=0x" & Hex$(wParam) & vbTab & "lParam=0x" & Hex$(lParam)
  End Select
  If fEat = False Then
    IMsgTarget_OnMsg = _
      MsgBlaster.CallOrigWndProc(hWnd, msg, wParam, lParam)
    Exit Function
  Else
    IMsgTarget_OnMsg = 1& 'Non-zero means we ate it
  End If
  If intAction = 1 Then SetScrollType pVert, LOBYTE
Exit Function
SubClass_Error:
 Exit Function
End Function
Private Sub SetScrollType(ByVal pVert As Boolean, ByVal pLoByte As Integer)
 Dim hWndVert As Long
 Dim hWndHorz As Long
 Dim typScroll As SCROLLINFO
 Dim i As Integer
 'Looking for Vertical scroll bar
 hWndVert = FindWindowEx(DataGrid1.hWnd, 0&, "ScrollBar", vbNullString)
 'Looking for Horizontal scroll bar
 hWndHorz = FindWindowEx(DataGrid1.hWnd, hWndVert, "ScrollBar", vbNullString)
 If pVert = True Then
  If Not hWndVert = 0 Then
    typScroll.cbSize = LenB(typScroll)
    typScroll.fMask = 31
   If GetScrollInfo(hWndVert, SB_CTL, typScroll) <> 0 Then
    Select Case pLoByte
     Case SB_THUMBTRACK
      DataGrid1.Scroll 0, typScroll.nTrackPos - typScroll.nPos
     Case SB_PAGEDOWN
      For i = 1 To DataGrid1.VisibleRows - 1
       DataGrid1.Scroll 0, 1
       Sleep 25
      Next i
     Case SB_PAGEUP
      For i = 1 To DataGrid1.VisibleRows - 1
       DataGrid1.Scroll 0, -1
       Sleep 25
      Next i
    End Select
   End If
  End If
 Else
  If Not hWndHorz = 0 Then
    typScroll.cbSize = LenB(typScroll)
    typScroll.fMask = 31
   If GetScrollInfo(hWndHorz, SB_CTL, typScroll) <> 0 Then
     DataGrid1.Scroll typScroll.nTrackPos - typScroll.nPos, 0
   End If
  End If
 End If
End Sub
```

