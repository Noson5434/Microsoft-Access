Attribute VB_Name = "modClipboard"
Option Compare Database
Option Explicit

' Module:       modClipboard
' Description:  This module provides functions to work with the Windows Clipboard in VBA, supporting
'               both 32-bit and 64-bit systems using appropriate Windows API calls. The functions allow you
'               to retrieve the contents of the clipboard, copy text to the clipboard, and clear the clipboard.

#If VBA7 Then
    ' Declaration of Windows API functions with PtrSafe keyword for 64-bit Office
    ' Clipboard Functions
    
    ' Opens the clipboard for examination and prevents other applications from modifying the clipboard content
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    ' Retrieves a handle to the specified clipboard format data
    Private Declare PtrSafe Function WinGetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
    ' Allocates a block of memory and returns a handle to the memory block
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    ' Locks a global memory object and returns a pointer to the first byte of the object's memory block
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    ' Copies a string to a buffer
    Private Declare PtrSafe Function lstrCopy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    ' Unlocks a global memory object, making it available for other processes to access
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    ' Closes the clipboard
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    ' Places data on the clipboard in a specified clipboard format
    Private Declare PtrSafe Function WinSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As LongPtr) As Long
    ' Empties the clipboard and frees handles to data in the clipboard
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
#Else
    ' Declaration of Windows API functions without PtrSafe keyword for 32-bit Office
    ' Clipboard Functions
    
    ' Opens the clipboard for examination and prevents other applications from modifying the clipboard content
    Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
    ' Retrieves a handle to the specified clipboard format data
    Private Declare Function WinGetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
    ' Allocates a block of memory and returns a handle to the memory block
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    ' Locks a global memory object and returns a pointer to the first byte of the object's memory block
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    ' Copies a string to a buffer
    Private Declare Function lstrCopy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    ' Unlocks a global memory object, making it available for other processes to access
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    ' Closes the clipboard
    Private Declare Function CloseClipboard Lib "user32" () As Long
    ' Places data on the clipboard in a specified clipboard format
    Private Declare Function WinSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    ' Empties the clipboard and frees handles to data in the clipboard
    Private Declare Function EmptyClipboard Lib "user32" () As Long
#End If

' GlobalAlloc flags - GHND: Combines GMEM_MOVEABLE and GMEM_ZEROINIT
Private Const GHND As Long = &H42
' Clipboard format - CF_TEXT: Plain text format
Private Const CF_TEXT As Long = 1
' Maximum size for the clipboard data (4,096 bytes)
Private Const mcintMaxSize As Integer = 4096

' Purpose:  Clears the clipboard by emptying its contents.
Public Sub ClearClipboardData()
    On Error GoTo Err_Handler

    ' Attempt to open the clipboard
    If OpenClipboard(0&) <> 0 Then
        ' Clear the Clipboard
        EmptyClipboard
        
        ' Close the clipboard
        CloseClipboard
    End If

Exit_Err_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Procedure: ClearClipboardData", vbCritical + vbOKOnly, "ClearClipboardData - Error"
    Resume Exit_Err_Handler
End Sub

' Purpose:  Retrieves the text contents from the clipboard.
' Returns:  The text contents of the clipboard as a String.
Public Function GetClipboardData() As String
    On Error GoTo Err_Handler

    #If VBA7 Then
        Dim lngClipMemory As LongPtr
    #Else
        Dim lngClipMemory As Long
    #End If
  
    Dim lngHandle As Long   ' Handle to the global memory holding clipboard text
    Dim strTemp As String   ' Temporary string to hold the retrieved text

    ' Attempt to open the clipboard
    If OpenClipboard(0&) <> 0 Then
    
        ' Get the handle to the global memory holding clipboard text
        lngHandle = WinGetClipboardData(CF_TEXT)
    
        ' Check if the memory allocation was successful
        If lngHandle <> 0 Then
      
            ' Lock the memory to obtain the string
            lngClipMemory = GlobalLock(lngHandle)
      
            ' If the memory was successfully locked
            strTemp = Space$(mcintMaxSize)
            lstrCopy strTemp, lngClipMemory
            GlobalUnlock lngHandle
      
            ' Strip off any null characters and trim the result
            strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
      
        End If
    
        ' Close the clipboard
        CloseClipboard
    End If
  
    ' Return the retrieved text
    GetClipboardData = strTemp

Exit_Err_Handler:
    Exit Function

Err_Handler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Procedure: GetClipboardData", vbCritical + vbOKOnly, "GetClipboardData - Error"
    Resume Exit_Err_Handler
End Function

' Purpose:  Writes the supplied string to the clipboard.
' Params:   strText - The text to write to the clipboard.
Public Sub SetClipboardData(ByVal strText As String)
    On Error GoTo Err_Handler

    #If VBA7 Then
        Dim lngHoldMem As LongPtr
        Dim lngGlobalMem As LongPtr
    #Else
        Dim lngHoldMem As Long
        Dim lngGlobalMem As Long
    #End If
  
    ' Allocate movable global memory
    lngHoldMem = GlobalAlloc(GHND, LenB(strText) + 1)
  
    ' Lock the memory to obtain a pointer
    lngGlobalMem = GlobalLock(lngHoldMem)
  
    ' Copy the string to the global memory
    lngGlobalMem = lstrCopy(lngGlobalMem, strText)
  
    ' Unlock the memory
    If GlobalUnlock(lngHoldMem) = 0 Then
    
        ' Open the clipboard to copy data to
        If OpenClipboard(0&) <> 0 Then
      
            ' Clear the clipboard
            EmptyClipboard
      
            ' Copy the data to the clipboard
            WinSetClipboardData CF_TEXT, lngHoldMem
            
            ' Close the clipboard
            CloseClipboard
        End If
    End If

Exit_Err_Handler:
    Exit Sub

Err_Handler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & "Error Description: " & Err.Description & vbCrLf & "Procedure: SetClipboardData", vbCritical + vbOKOnly, "SetClipboardData - Error"
    Resume Exit_Err_Handler
End Sub
