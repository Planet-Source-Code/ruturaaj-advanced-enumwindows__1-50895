Attribute VB_Name = "modEnumWindows"
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Enum EnumFilter
    [No_Filter]
    [Only_Enabled]
    [Only_Visible]
    [Only_Enabled_Visible]
    [Only_Enabled_NonVisible]
    [Only_Disabled_Visible]
    [Only_Disabled_NonVisible]
    [Only_Visible_WinTextNotEmpty]
End Enum

Public EnumCondition As EnumFilter
    

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    Dim sWinText As String
    Dim lngWinTextLen As Long
    Dim bIsWin As Boolean
    Dim bIsVisible As Boolean
    Dim bIsEnabled As Boolean
    Dim lstItem As ListItem
    Dim lstSubItem As ListSubItem
    Dim lstImage As Integer
    
       
    'See if Handle returns a Window ...
    If IsWindow(hwnd) = 0 Then bIsWin = False Else bIsWin = True
    
    'See if Window is Visible ...
    If bIsWin = True Then
        If IsWindowVisible(hwnd) = 0 Then bIsVisible = False Else bIsVisible = True
    
    'See if Window is Enabled
        If IsWindowEnabled(hwnd) = 0 Then bIsEnabled = False Else bIsEnabled = True
    
    'Get Window Text Length ...
        lngWinTextLen = GetWindowTextLength(hwnd)
    
    'Get Window Text ...
        sWinText = Space(lngWinTextLen)
        GetWindowText hwnd, sWinText, lngWinTextLen + 1
    
    End If
    
    
    'ImageList Images ...
    '1 : Enabled
    '2 : Visible
    '3 : Window
    With frmMain
        If bIsWin = True Then
            If bIsEnabled = True Then
                If bIsVisible = True Then
                    lstImage = 2
                Else
                    lstImage = 1
                End If
            Else
                lstImage = 3
            End If
        End If
    
        'List Sequence ...
        '1 : Window Handle
        '2 : Window Text
        '3 : bIsVisible
        '4 : bIsEnabled
        
        'Fill the ListView Control with Data ...
        Select Case EnumCondition
            Case No_Filter:
                Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                With lstItem.ListSubItems
    
                    If Trim(sWinText) <> "" Then
                        .Add , , sWinText
                    Else
                        .Add , , "- NA -"
                    End If
                    
                    If bIsVisible = True Then
                        .Add , , "Visible"
                        lstItem.ForeColor = vbRed
                        lstItem.Bold = True
                    Else
                        .Add , , "Not Visible"
                    End If
                    
                    If bIsEnabled = True Then
                        .Add , , "Enabled"
                    Else
                        .Add , , "Disabled"
                    End If
                
                End With
            
            Case Only_Visible
                If bIsVisible = True Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
    
                        If Trim(sWinText) <> "" Then
                        .Add , , sWinText
                        Else
                          .Add , , "- NA -"
                        End If
                        
                        .Add , , "Visible"
                        lstItem.ForeColor = vbRed
                        lstItem.Bold = True
                        
                        If bIsEnabled = True Then
                            .Add , , "Enabled"
                        Else
                            .Add , , "Disabled"
                        End If
                    End With
                End If
            
            Case Only_Enabled
                If bIsEnabled = True Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                        .Add , , sWinText
                        Else
                          .Add , , "- NA -"
                        End If
                        
                        If bIsVisible = True Then
                            .Add , , "Visible"
                            lstItem.ForeColor = vbRed
                            lstItem.Bold = True
                        Else
                            .Add , , "Not Visible"
                        End If
                        .Add , , "Enabled"
                    End With
                End If
                
            Case Only_Visible_WinTextNotEmpty
                If bIsVisible = True And Trim(sWinText) <> "" Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        .Add , , sWinText
                        .Add , , "Visible"
                        lstItem.ForeColor = vbRed
                        lstItem.Bold = True
                        
                        If bIsEnabled = True Then
                            .Add , , "Enabled"
                        Else
                            .Add , , "Disabled"
                        End If
                    End With
                End If
                
            Case Only_Enabled_Visible
                If bIsEnabled = True And bIsVisible = True Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                        .Add , , sWinText
                        Else
                          .Add , , "- NA -"
                        End If
                        .Add , , "Visible"
                        lstItem.ForeColor = vbRed
                        lstItem.Bold = True
                        .Add , , "Enabled"
                    End With
                End If
            
            Case Only_Enabled_NonVisible
                If bIsEnabled = True And bIsVisible = False Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                            .Add , , sWinText
                        Else
                            .Add , , "- NA -"
                        End If
                        .Add , , "Not Visible"
                        .Add , , "Enabled"
                    End With
                End If
                
            Case Only_Disabled_NonVisible
                If bIsEnabled = False And bIsVisible = False Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                            .Add , , sWinText
                        Else
                             .Add , , "- NA -"
                        End If
                        .Add , , "Not Visible"
                        .Add , , "Disabled"
                    End With
                End If
            
            Case Only_Disabled_Visible
                If bIsEnabled = False And bIsVisible = True Then
                    Set lstItem = .lstWinList.ListItems.Add(, , hwnd, lstImage, lstImage)
                    With lstItem.ListSubItems
                        If Trim(sWinText) <> "" Then
                            .Add , , sWinText
                        Else
                            .Add , , "- NA -"
                        End If
                        .Add , , "Visible"
                        lstItem.ForeColor = vbRed
                        lstItem.Bold = True
                        .Add , , "Disabled"
                    End With
                End If
        End Select
    End With
        
    'Continue same process ...
    EnumWindowsProc = True
End Function

Public Function GetWinInfo()
    'Clear existing entries in ListView ...
    frmMain.lstWinList.ListItems.Clear
    
    'Call EnumWindowsProc ...
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
End Function




