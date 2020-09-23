Attribute VB_Name = "modOwnerDrawnMenu"
Option Explicit

#Const UNICODE = False

Public RunningOnXP As Boolean

'*******************************************************************************************
'                                             ENUMS
'*******************************************************************************************

Private Enum MENU_DRAWING       ' Menu drawing constants
    ODS_SELECTED = &H1
    ODS_GRAYED = &H2
    ODS_HOTLIGHT = &H40
End Enum

Private Enum MENU_TYPE          ' Menu type constants
    MT_ITEM
    MT_SUB
End Enum
    
'*******************************************************************************************
'                                             TYPES
'*******************************************************************************************

Private Type BITMAP             ' BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type MEASUREITEMSTRUCT   ' MEASUREITEMSTRUCT
    CtlType    As Long
    CtlID      As Long
    itemID     As Long
    itemWidth  As Long
    itemHeight As Long
    itemData   As Long
End Type

Public Type RECT                ' RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type POINTAPI           ' POINT
    x As Long
    y As Long
End Type

Public Type DRAWITEMSTRUCT      ' DRAWITEMSTRUCT
    CtlType    As Long
    CtlID      As Long
    itemID     As Long
    itemAction As Long
    itemState  As Long
    hwndItem   As Long
    hDC        As Long
    rcItem     As RECT
    itemData   As Long
End Type

Public Type MENULIST            ' MENULIST
    hMenu      As Long          ' Menu handle that this item belongs to
    ItemName   As String        ' Menu item's name
    NameSize   As Size          ' Menu item's name size (cx, cy)
    hBitmap    As Long          ' Menu item's bitmap
    hMaskBmp   As Long          ' Handle to mask bitmap (for transparency effect)
    BitmapSize As Size          ' Menu item's bitmap size
    AccName    As String        ' Menu item's accelerator string
    AccSize    As Size          ' Menu item's accelerator string size
    Height     As Long          ' Total menu item height
    Width      As Long          ' Total menu item width
    PicPath    As String        ' The path of the menu picture
    Position   As Long          ' Position of the menu item (0..n-1)
    MenuType   As MENU_TYPE     ' Type of menu item
End Type

'*******************************************************************************************
'                                           CONSTANTS
'*******************************************************************************************

Private Const MF_BYPOSITION = &H400&
Private Const MF_POPUP = &H10&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_GRAYED = &H1&
Private Const MF_BITMAP = &H4&
Private Const MF_SYSMENU = &H2000&

Private Const COLOR_HIGHLIGHTTEXT = 14
Private Const COLOR_MENUTEXT = 7
Private Const COLOR_HIGHLIGHT = 13
Private Const COLOR_MENU = 4
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_BTNHIGHLIGHT = 20
Private Const COLOR_MENUBAR = 30
Private Const COLOR_MENUHILIGHT = 29

Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const WM_MENUCHAR = &H120
Private Const MNC_EXECUTE = &H2
Private Const MNC_IGNORE = &H0

Private Const DT_RIGHT = &H2
Private Const DT_LEFT = &H0

Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Private Const TRANSPARENT = &H1
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0
Private Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE
Private Const GWL_WNDPROC = -4
Private Const PS_SOLID = 0

Private Const SPI_GETFLATMENU = &H1022
Private Const MSAA_MENU_SIG = &HAA0DF00D

'*******************************************************************************************
'                                   API FUNCTION DECLARATIONS
'*******************************************************************************************

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GrayString Lib "user32" Alias "GrayStringA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpOutputFunc As Long, ByVal lpData As String, ByVal nCount As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
#If UNICODE Then
    Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
#Else
    Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
#End If
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'*******************************************************************************************
'                                          GLOBALS
'*******************************************************************************************

Dim oldWndProc       As Long         ' Previous WndProc address
Dim m_hWnd           As Long         ' Main window handle
Dim m_FontName       As String       ' Selected font name
Dim m_FontSize       As Long         ' Selected font size
Dim m_FontWeight     As Long         ' 700 if font is bold, 0 otherwise
Dim m_FontItalic     As Boolean      ' True if bold is italic, false otherwise
Dim m_hMenuBarBitmap As Long
Dim nItems           As Long         ' The number of menu items (0..nItems - 1)
Public Menus(100)    As MENULIST

'*******************************************************************************************
'                                         SUBCLASSING
'*******************************************************************************************

Private Function MyWndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case wMsg
    Case WM_MEASUREITEM
         Dim MIS As MEASUREITEMSTRUCT
        
         'get MEASUREITEMSTRUCT data into MIS
         CopyMemory MIS, ByVal lParam, Len(MIS)
                 
         'itemData contains the number of the menu item
         With Menus(MIS.itemData)
            ' Add 4 extra pixels (2 for each direction)
            MIS.itemHeight = Max(.NameSize.cy, Max(.AccSize.cy, .BitmapSize.cy)) + 4
            
            ' If this is a top level menu
            If .hMenu = Menus(0).hMenu Then
                ' Do not add extra space if there is no picture
                If .BitmapSize.cx = 0 Then
                    MIS.itemWidth = .NameSize.cx
                Else
                    MIS.itemWidth = .BitmapSize.cx + .NameSize.cx + 14
                End If
            Else
                ' Add 14 extra pixels (2 for each direction and 10 for the gap between the picture and the Text)
                MIS.itemWidth = .BitmapSize.cx + .NameSize.cx + .AccSize.cx + 14
            End If
         End With
        
         'copy data back
         CopyMemory ByVal lParam, MIS, Len(MIS)
        
         'return true (we have processed the message)
         MyWndProc = 1

    Case WM_DRAWITEM
         Dim DIS As DRAWITEMSTRUCT
         
         'get DRAWITEMSTRUCT data into DIS
         CopyMemory DIS, ByVal lParam, Len(DIS)
                  
'********************************************************************************
' Speak the menu item if it is a submenu
'If Menus(DIS.itemData).MenuType = MT_SUB Then
'    If (DIS.itemState And ODS_SELECTED) Or (DIS.itemState And ODS_HOTLIGHT) Then
'        Dim ToSpeak() As String
'        Dim FinalString As String
'        Dim i As Integer
'
'        FinalString = ""
'        ToSpeak = Split(Menus(DIS.itemData).ItemName, "&")
'
'        For i = 0 To UBound(ToSpeak)
'            FinalString = FinalString & ToSpeak(i)
'        Next i
'
'        frmVoice.Speak FinalString
'    End If
'End If
'********************************************************************************

         DrawMenuItem DIS.hDC, DIS.rcItem, DIS.itemData, DIS.itemState

         'return true (we have processed the message)
         MyWndProc = 1
    Case WM_MENUCHAR
         Dim Hi As Long
         Dim Lo As Long
         Dim ID As Long
         
        ' We need to process Top level menu shortcut keys
         GetHiLoWord wParam, Lo, Hi
         lParam = 0
         Select Case LCase(Chr(Lo))
         Case "f"
              Hi = -(frmMain.mnuFile.Visible) - 1
              SetHiLoWord lParam, Hi, MNC_EXECUTE
         Case "e"
              Hi = -(frmMain.mnuFile.Visible + frmMain.mnuEdit.Visible) - 1
              SetHiLoWord lParam, Hi, MNC_EXECUTE
         End Select

         MyWndProc = lParam
    Case Else
         'any other message sent to the menu Form will be processed by the default window procedure
         MyWndProc = CallWindowProc(oldWndProc, hWnd, wMsg, wParam, lParam)
    End Select
End Function

' Begin Subclassing
Public Sub SubClass(WndHwnd)
    oldWndProc = SetWindowLong(WndHwnd, GWL_WNDPROC, AddressOf MyWndProc)
End Sub

' End subclassing
Public Sub UnSubClass()
    If oldWndProc Then
        SetWindowLong m_hWnd, GWL_WNDPROC, oldWndProc
        oldWndProc = 0
    End If
End Sub

'*******************************************************************************************
'                                       MENU CREATION
'*******************************************************************************************

' Adds owner drawn menus to the menu bar
Public Sub AddMenus(hWnd As Long, Optional FontName As String, Optional FontSize As Long, Optional FontBold As Boolean, Optional FontItalic As Boolean)
    Dim hMainMenu  As Long
    
    UnSubClass
    
    ' Initialise the menu items count
    nItems = 0
    
    ' Save the information passed to globals
    m_hWnd = hWnd
    m_FontName = IIf(FontName = "", "Arial", FontName)
    m_FontSize = IIf(FontSize <= 0, 18, FontSize)
    m_FontWeight = IIf(FontBold, 700, 0)
    m_FontItalic = FontItalic
    
    ' Get the handle to the menu bar of the window
    hMainMenu = GetMenu(hWnd)
    
    ' Determine the look of the menu (XP flat style or normal)
    RunningOnXP = HasFlatMenu
    
    ' Subclass the main window
    SubClass hWnd
    
    ' Get the menu item information
    GetMenuItems hMainMenu, nItems
        
    ' Redraw the menu
    DrawMenuBar hWnd
End Sub

' Stops subclassing  and removes the menu resources
Public Sub RemoveMenus()
    Dim i As Long
    
    UnSubClass
    
    For i = 0 To nItems
        DeleteObject Menus(i).hBitmap
        DeleteObject Menus(i).hMaskBmp
    Next i
    
    If m_hMenuBarBitmap Then
        DeleteObject m_hMenuBarBitmap
        m_hMenuBarBitmap = 0
    End If
    
    Erase Menus
    
    nItems = 0
End Sub

' Recursive function that gets the menu items of the entire menu
Private Sub GetMenuItems(hMenu As Long, nItems As Long)
    Dim nMenuItems  As Long     ' number of menu items under the current menu
    Dim hMenuTemp   As Long     ' Temporary menu handle
    Dim hMenuBar    As Long     ' Handle of menu bar (top level menu)
    Dim Buf         As String   ' String buffer to store menu strings
    Dim ID          As Long     ' Menu item id
    Dim i           As Long     ' used with loops
    
    'Get the number of menuitems under the current menu
    nMenuItems = GetMenuItemCount(hMenu) - 1
    
    ' Get the menu bar handle
    hMenuBar = GetMenu(m_hWnd)
    
    For i = 0 To nMenuItems
        ' Get the menu id
        ID = GetMenuItemID(hMenu, i)
        
        ' Prepare a string buffer
        Buf = String(255, vbNullChar)
        
        ' If this is a popup menu
        If ID = -1 Then
            ' Get the sub menu's text
            GetMenuString hMenu, i, Buf, Len(Buf), MF_BYPOSITION
            Buf = Left(Buf, InStr(1, Buf, vbNullChar) - 1)
            
            ' Save menu info
            Menus(nItems).ItemName = Buf
            Menus(nItems).hMenu = hMenu
            Menus(nItems).Position = i
            Menus(nItems).MenuType = MT_SUB
                                    
            ' Update maximums
            UpdateMaximums (hMenu)
            
            ' Change the submenu to owner drawn menu
            hMenuTemp = GetSubMenu(hMenu, i)
            RemoveMenu hMenu, i, MF_BYPOSITION
            InsertMenu hMenu, i, MF_BYPOSITION Or MF_OWNERDRAW Or MF_POPUP, hMenuTemp, nItems
            
            ' Get the items of this sub menu
            nItems = nItems + 1
            GetMenuItems GetSubMenu(hMenu, i), nItems
        ' If this is a menu item
        Else
            ' Get the menu item's text
            GetMenuString hMenu, i, Buf, Len(Buf), MF_BYPOSITION

            ' if the item is not a separator
            If Mid(Buf, 1, 1) <> vbNullChar Then
                Buf = Left(Buf, InStr(1, Buf, vbNullChar) - 1)
                
                ' If it has a keyboard accelerator
                If InStr(1, Buf, vbTab) Then
                    Menus(nItems).ItemName = Left(Buf, InStr(1, Buf, vbTab) - 1)
                    Menus(nItems).AccName = Mid(Buf, InStr(1, Buf, vbTab) + 1)
                ' If it has no accelerator
                Else
                    Menus(nItems).ItemName = Buf
                End If
                
                ' Save the menu handle
                Menus(nItems).hMenu = hMenu
                Menus(nItems).Position = i
                Menus(nItems).MenuType = MT_ITEM
                                
                ' Update the maximums
                UpdateMaximums hMenu
                
                ' Change the menu item to owner drawn
                ModifyMenu hMenu, i, MF_BYPOSITION Or MF_OWNERDRAW Or GetMenuState(hMenu, i, MF_BYPOSITION), ID, nItems
                
                nItems = nItems + 1
            End If
        End If
    Next i
End Sub

' Gets the current menu item's sizes
Private Sub GetMenuItemSizes(NameSize As Size, AccSize As Size, PicSize As Size)
    Dim hDC          As Long    ' hdc to do calculations
    Dim hFont        As Long    ' handle to logical font
    Dim FontHeight   As Long    ' The height of the font
    Dim hBitmap      As Long    ' Bitmap handle
    Dim bmp          As BITMAP  ' Bitmap structure to save the bitmap object
    
    ' Create a window compatible dc
    hDC = CreateCompatibleDC(ByVal 0&)
    
    ' Create a font
    FontHeight = -MulDiv(m_FontSize, GetDeviceCaps(hDC, LOGPIXELSY), 72)
    hFont = CreateFont(FontHeight, 0, 0, 0, m_FontWeight, m_FontItalic, 0, 0, 0, 0, 0, 0, 0, m_FontName)
    
    ' Select the font in the device context
    hFont = SelectObject(hDC, hFont)
    
    ' Get the text size of the menu name and the accelerator
    ' for the given font name and size
    With Menus(nItems)
        GetTextExtentPoint32 hDC, .ItemName, Len(.ItemName), .NameSize
        GetTextExtentPoint32 hDC, .AccName, Len(.AccName), .AccSize
        
        ' Get the bitmap size - if there is a bitmap
        .hBitmap = LoadImage(ByVal 0&, .PicPath, IMAGE_BITMAP, ByVal 0&, ByVal 0&, LR_LOADFROMFILE)
        If .hBitmap <> 0 Then
            GetObject .hBitmap, Len(bmp), bmp
            .BitmapSize.cx = bmp.bmWidth
            .BitmapSize.cy = bmp.bmHeight
            
            ' Create the mask bitmaps
            CreateMaskBitmap .hBitmap, .hMaskBmp, .BitmapSize
        End If
        
        ' Save to return values
        NameSize = .NameSize
        AccSize = .AccSize
        PicSize = .BitmapSize
    End With
    
    ' Memory clean up
    DeleteObject (SelectObject(hDC, hFont))
    DeleteDC hDC
End Sub

' Finds and stores the maximum sizes of all the menu items
Private Sub UpdateMaximums(hCurrentMenu As Long)
    Dim MaxName      As Size    ' Maximum name string so far
    Dim MaxAccel     As Size    ' Maximum accelerator string so far
    Dim MaxPic       As Size    ' Maximum picture size so far
    Dim i            As Long    ' used with loops
    Dim ChangeAllY   As Boolean ' True when all the items heights need to be updated
    Dim ChangeAllX   As Boolean ' True when all the items widths need to be updated
    Dim hMainMenu    As Long    ' Handle to main menu
    
    ' Get the handle to the menu bar
    hMainMenu = GetMenu(m_hWnd)
    
    ' Get the menu item's sizes
    GetMenuItemSizes MaxName, MaxAccel, MaxPic
    
    ' Look in the menu info array for bigger menu items
    ChangeAllX = False
    ChangeAllY = False
    
    For i = 0 To nItems
        ' If this menu shares the same handle with the added one
        If Menus(i).hMenu = hCurrentMenu Then
            ' If the menu is in the menu bar
            If Menus(i).hMenu = hMainMenu Then
                ' If the added menu item has a smaller height
                If Menus(i).NameSize.cy > MaxName.cy Then
                    Menus(nItems).NameSize.cy = Menus(i).NameSize.cy
                    MaxName.cy = Menus(i).NameSize.cy
                Else
                    ChangeAllY = True
                End If
                
                ' Check picture - height matters
                If Menus(i).BitmapSize.cy > MaxPic.cy Then
                    Menus(nItems).BitmapSize.cy = Menus(i).BitmapSize.cy
                    MaxPic.cy = Menus(i).BitmapSize.cy
                Else
                    ChangeAllY = True
                End If
                
            ' The menu is a normal menuitem or submenu
            Else
                ' If stored menu width > added then change the added
                If Menus(i).NameSize.cx > MaxName.cx Then
                    Menus(nItems).NameSize.cx = Menus(i).NameSize.cx
                    MaxName.cx = Menus(i).NameSize.cx
                Else
                    ChangeAllX = True
                End If
                
                ' If stored accel width > added then change the added
                If Menus(i).AccSize.cx > MaxAccel.cx Then
                    Menus(nItems).AccSize.cx = Menus(i).AccSize.cx
                    MaxAccel.cx = Menus(i).AccSize.cx
                Else
                    ChangeAllX = True
                End If
                
                ' Check the picture - width matters
                If Menus(i).BitmapSize.cx > MaxPic.cx Then
                    Menus(nItems).BitmapSize.cx = Menus(i).BitmapSize.cx
                    MaxPic.cx = Menus(i).BitmapSize.cx
                Else
                    ChangeAllX = True
                End If
                
                ' Exit the loop
                Exit For
            End If
        End If
    Next i
    
    ' If the previous menus need updating
    If ChangeAllY Then
        For i = 0 To nItems
            With Menus(i)
                If .hMenu = hCurrentMenu Then
                   .NameSize.cy = MaxName.cy
                   .AccSize.cy = MaxAccel.cy
                   .BitmapSize.cy = MaxPic.cy
                End If
            End With
        Next
    End If
    
    If ChangeAllX Then
        For i = 0 To nItems
            With Menus(i)
                If .hMenu = hCurrentMenu Then
                   .NameSize.cx = MaxName.cx
                   .AccSize.cx = MaxAccel.cx
                   .BitmapSize.cx = MaxPic.cx
                End If
            End With
        Next
    End If
End Sub

' Adds a bitmap to the menu bar in order to resize it properly
Public Sub AddEmptyBitmapToMenuBar(hWnd As Long, nHeight As Long)
    Dim hMenu As Long
    Dim hBitmap As Long
    Dim hDC As Long
    
    ' Create a black and white bitmap
    hBitmap = CreateBitmap(1, nHeight, 1, 1, 0)
    
    ' Paint the bitmap white
    hDC = CreateCompatibleDC(0)
    hBitmap = SelectObject(hDC, hBitmap)
    PatBlt hDC, 0, 0, 1, nHeight, WHITENESS
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
    
    ' Add the white bitmap to the menu
    m_hMenuBarBitmap = hBitmap
    hMenu = GetMenu(hWnd)
    InsertMenu hMenu, -1, MF_BYPOSITION Or MF_BITMAP Or MF_GRAYED, 0, m_hMenuBarBitmap
End Sub

'*******************************************************************************************
'                                             DRAWING
'*******************************************************************************************

' Draws a menu item
Private Sub DrawMenuItem(hDC As Long, rc As RECT, Index As Long, Selected As MENU_DRAWING)
    Dim TextColor  As Long   ' Color of the menu text
    Dim BackColor  As Long   ' Background color of the menu
    Dim BackMode   As Long   ' Previous background mode of hDC
    Dim hFont      As Long   ' Handle to logical font
    Dim FontHeight As Long   ' Font height in pixels
    Dim bmp        As BITMAP ' Bitmap object
    Dim rcBitmap   As RECT   ' Rectangle with the bitmap size
    Dim TextRC     As RECT   ' Rectangle of the text
    Dim hPen1      As Long   ' Pen of the Dark side
    Dim hPen2      As Long   ' Pen of the Bright side
    Dim hPenPrev   As Long   ' May the previous Pen be with you
    Dim pt         As POINTAPI ' dummy point value passed to MoveToEx
    
    ' Set the background mode to transparent
    BackMode = SetBkMode(hDC, TRANSPARENT)
    
    ' Determine the menu text and background color
    ' If this is a top menu
    If Menus(Index).hMenu = Menus(0).hMenu Then
        If RunningOnXP Then
            If Selected And ODS_SELECTED Or Selected And ODS_HOTLIGHT Then
                BackColor = COLOR_MENUHILIGHT + 1
                TextColor = GetSysColor(COLOR_HIGHLIGHTTEXT)
            Else
                BackColor = COLOR_MENUBAR + 1
                TextColor = GetSysColor(COLOR_MENUTEXT)
            End If
        Else
            BackColor = COLOR_MENU + 1
            TextColor = COLOR_MENUTEXT
        End If
    Else
        BackColor = IIf(Selected And ODS_SELECTED, COLOR_HIGHLIGHT, COLOR_MENU) + 1
        TextColor = GetSysColor(IIf(Selected And ODS_SELECTED, COLOR_HIGHLIGHTTEXT, COLOR_MENUTEXT))
    End If
        
    ' Fill the background rectangle
    FillRect hDC, rc, BackColor
    
    ' If this is a menu bar item
    If Menus(Index).hMenu = Menus(0).hMenu Then
        ' And is not just painted normally
        If (Selected And ODS_SELECTED) Or (Selected And ODS_HOTLIGHT) Then
            If RunningOnXP Then
                hPen1 = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_HIGHLIGHT))
                hPen2 = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_HIGHLIGHT))
            Else
                ' Create two pens
                If Selected And ODS_HOTLIGHT Then
                    hPen1 = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT))
                    hPen2 = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW))
                ElseIf Selected And ODS_SELECTED Then
                    hPen1 = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW))
                    hPen2 = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNHIGHLIGHT))
                End If
            End If
            
            ' Draw the shadows
            MoveToEx hDC, rc.Left, rc.Bottom - 1, pt
            hPenPrev = SelectObject(hDC, hPen1)
            LineTo hDC, rc.Left, rc.Top
            LineTo hDC, rc.Right - 1, rc.Top
            SelectObject hDC, hPen2
            LineTo hDC, rc.Right - 1, rc.Bottom - 1
            LineTo hDC, rc.Left, rc.Bottom - 1
            
            ' Clean up
            SelectObject hDC, hPenPrev
            DeleteObject hPen1
            DeleteObject hPen2
        End If
    End If
    
    ' Calculate the position to place the bitmap
    GetObject Menus(Index).hBitmap, Len(bmp), bmp
    If Menus(Index).hMenu = Menus(0).hMenu And (Selected And ODS_SELECTED) Then
        rcBitmap.Left = ((Menus(Index).BitmapSize.cx - bmp.bmWidth) / 2) + rc.Left + 3
        rcBitmap.Top = (((rc.Bottom - rc.Top - 1) - bmp.bmHeight) / 2) + rc.Top + 1
    Else
        rcBitmap.Left = ((Menus(Index).BitmapSize.cx - bmp.bmWidth) / 2) + rc.Left + 2
        rcBitmap.Top = (((rc.Bottom - rc.Top) - bmp.bmHeight) / 2) + rc.Top
    End If
    rcBitmap.Bottom = bmp.bmHeight
    rcBitmap.Right = bmp.bmWidth
    
    ' Add the menu bitmap
    AddBitmap hDC, rcBitmap, BackColor, Index, (Selected And ODS_GRAYED)
        
    ' Calculate the font height and create a font, select it in the DC
    FontHeight = -MulDiv(m_FontSize, GetDeviceCaps(hDC, LOGPIXELSY), 72)
    hFont = CreateFont(FontHeight, 0, 0, 0, m_FontWeight, m_FontItalic, 0, 0, 0, 0, 0, 0, 0, m_FontName)
    hFont = SelectObject(hDC, hFont)
    
    ' Calculate the text coordinates
    With Menus(Index)
        If .hMenu = Menus(0).hMenu And (Selected And ODS_SELECTED) And Not RunningOnXP Then
            TextRC.Left = .BitmapSize.cx + 10 + rc.Left + 3
            TextRC.Top = (((rc.Bottom - rc.Top) - .NameSize.cy) / 2) + rc.Top + 3
        Else
            TextRC.Left = .BitmapSize.cx + 10 + rc.Left + 2
            TextRC.Top = (((rc.Bottom - rc.Top) - .NameSize.cy) / 2) + rc.Top + 2
        End If
        TextRC.Right = TextRC.Left + .NameSize.cx
        TextRC.Bottom = TextRC.Top + .NameSize.cy
    End With
    
    ' Add the text
    TextColor = SetTextColor(hDC, TextColor)
    AddName hDC, TextRC, Index, Selected
    
    ' Add the Accelerator
    AddAccel hDC, rc, Index, 20, Selected
    
    ' Memory cleanup and DC restoration
    SetBkMode hDC, BackMode
    TextColor = SetTextColor(hDC, TextColor)
    DeleteObject SelectObject(hDC, hFont)
End Sub

' Adds the bitmap to the menu item
Private Sub AddBitmap(hDC As Long, rc As RECT, ByVal BackColor As Long, Index As Long, Disabled As Boolean)
    Dim hdcPic    As Long   ' Off screen DC with the picture
    Dim hdcMask   As Long   ' Off screen DC with the Mask of the picture
    Dim hBitmap   As Long   ' Handle to bitmap
    Dim hMask     As Long   ' Handle to bitmap mask

    ' Create two off screen DCs
    hdcPic = CreateCompatibleDC(ByVal 0&)
    hdcMask = CreateCompatibleDC(ByVal 0&)
    
    ' Select the bitmap and the mask in them
    hBitmap = SelectObject(hdcPic, Menus(Index).hBitmap)
    hMask = SelectObject(hdcMask, Menus(Index).hMaskBmp)
    
    ' Change the backcolor for the BitBlt effect to work correctly
    BackColor = SetBkColor(hDC, GetSysColor(BackColor - 1))
    
    ' Print the image on the menu hdc
    If Not Disabled Then
        BitBlt hDC, rc.Left, rc.Top, rc.Right, rc.Bottom, hdcMask, 0, 0, vbSrcAnd
    End If
    BitBlt hDC, rc.Left, rc.Top, rc.Right, rc.Bottom, hdcPic, 0, 0, vbSrcPaint
    
    ' Restore the backcolor
    SetBkColor hDC, BackColor
    
    ' Memory clean up
    SelectObject hdcPic, hBitmap
    SelectObject hdcMask, hMask
    DeleteDC hdcPic
    DeleteDC hdcMask
End Sub

' Adds text to the menu item
Private Sub AddName(hDC As Long, TextRC As RECT, Index As Long, Disabled As MENU_DRAWING)
    Dim txColor As Long
    
    ' Print the text on the DC
    With Menus(Index)
        If Disabled And ODS_GRAYED Then
            ' Change the position and the color of the text
            TextRC.Top = TextRC.Top + 1
            TextRC.Left = TextRC.Left + 1
            txColor = SetTextColor(hDC, GetSysColor(COLOR_BTNHIGHLIGHT))
            
            If Not CBool(Disabled And ODS_SELECTED) Then
#If UNICODE Then
                DrawText hDC, StrPtr(.ItemName), Len(.ItemName), TextRC, DT_LEFT
#Else
                DrawText hDC, .ItemName, Len(.ItemName), TextRC, DT_LEFT
#End If
            End If
            
            ' Change the position and the color of the text
            TextRC.Top = TextRC.Top - 1
            TextRC.Left = TextRC.Left - 1
            SetTextColor hDC, GetSysColor(COLOR_BTNSHADOW)
#If UNICODE Then
            DrawText hDC, StrPtr(.ItemName), Len(.ItemName), TextRC, DT_LEFT
#Else
            DrawText hDC, .ItemName, Len(.ItemName), TextRC, DT_LEFT
#End If
            
            ' Restore the color
            SetTextColor hDC, txColor
        Else
#If UNICODE Then
            DrawText hDC, StrPtr(.ItemName), Len(.ItemName), TextRC, DT_LEFT
#Else
            DrawText hDC, .ItemName, Len(.ItemName), TextRC, DT_LEFT
#End If
        End If
    End With
End Sub

' Adds the accelerator to the menu dc
Private Sub AddAccel(hDC As Long, rc As RECT, Index As Long, ExtraSpace As Long, Disabled As MENU_DRAWING)
    Dim TextRC  As RECT
    Dim txColor As Long
    
    With Menus(Index)
        TextRC.Left = .BitmapSize.cx + .NameSize.cx + rc.Left + ExtraSpace + 1
        TextRC.Top = (((rc.Bottom - rc.Top) - .AccSize.cy) / 2) + rc.Top + 1
        TextRC.Right = TextRC.Left + .AccSize.cx
        TextRC.Bottom = TextRC.Top + .AccSize.cy
        
        If Disabled And ODS_GRAYED Then
            ' Change the position and the color of the text
            TextRC.Top = TextRC.Top + 1
            TextRC.Left = TextRC.Left + 1
            txColor = SetTextColor(hDC, GetSysColor(COLOR_BTNHIGHLIGHT))
            
            If Not CBool(Disabled And ODS_SELECTED) Then
#If UNICODE Then
                DrawText hDC, StrPtr(.AccName), Len(.AccName), TextRC, DT_RIGHT
#Else
                DrawText hDC, .AccName, Len(.AccName), TextRC, DT_RIGHT
#End If
            End If
            
            ' Change the position and the color of the text
            TextRC.Top = TextRC.Top - 1
            TextRC.Left = TextRC.Left - 1
            SetTextColor hDC, GetSysColor(COLOR_BTNSHADOW)

#If UNICODE Then
            DrawText hDC, StrPtr(.AccName), Len(.AccName), TextRC, DT_RIGHT
#Else
            DrawText hDC, .AccName, Len(.AccName), TextRC, DT_RIGHT
#End If
            
            ' Restore the color
            SetTextColor hDC, txColor
        Else
#If UNICODE Then
            DrawText hDC, StrPtr(.AccName), Len(.AccName), TextRC, DT_RIGHT
#Else
            DrawText hDC, .AccName, Len(.AccName), TextRC, DT_RIGHT
#End If
        End If
    End With
End Sub

' Create the mask bitmaps
Private Sub CreateMaskBitmap(hBitmap As Long, hBmpMask As Long, sz As Size)
    Dim hdcBitmap   As Long      ' DC with the image
    Dim hdcMask     As Long      ' DC with the mask bitmap of the image
    Dim hdcInvMask  As Long      ' DC with the inverse of the mask bitmap
    Dim hBmpInvMask As Long      ' Bitmap of the inverse of the mask bitmap
    Dim PrevColor   As Long      ' Previous color of the image DC
    
    ' Create three off screen DCs
    hdcBitmap = CreateCompatibleDC(ByVal 0&)
    hdcMask = CreateCompatibleDC(ByVal 0&)
    hdcInvMask = CreateCompatibleDC(ByVal 0&)
    
    ' Create two monochrome bitmaps
    hBmpMask = CreateBitmap(sz.cx, sz.cy, 1, 1, 0)
    hBmpInvMask = CreateBitmap(sz.cx, sz.cy, 1, 1, 0)
    
    ' Select the bitmaps in the DCs
    hBitmap = SelectObject(hdcBitmap, hBitmap)
    hBmpMask = SelectObject(hdcMask, hBmpMask)
    hBmpInvMask = SelectObject(hdcInvMask, hBmpInvMask)
    
    ' Change the background color of the image DC to magenta
    PrevColor = SetBkColor(hdcBitmap, vbMagenta)
    
    ' Create the mask image
    BitBlt hdcMask, 0, 0, sz.cx, sz.cy, hdcBitmap, 0, 0, vbSrcCopy
    
    ' Restore the back color
    SetBkColor hdcBitmap, PrevColor
    
    ' Create the inverse of the object mask
    BitBlt hdcInvMask, 0, 0, sz.cx, sz.cy, hdcMask, 0, 0, vbNotSrcCopy
    
    ' Mask out the transparent colored pixels on the bitmap.
    BitBlt hdcBitmap, 0, 0, sz.cx, sz.cy, hdcInvMask, 0, 0, vbSrcAnd
    
    ' Return the bitmaps
    hBmpMask = SelectObject(hdcMask, hBmpMask)
    hBitmap = SelectObject(hdcBitmap, hBitmap)
    
    ' Perform memory cleanup
    DeleteObject (SelectObject(hBmpInvMask, hdcInvMask))
    DeleteDC (hdcInvMask)
    DeleteDC (hdcMask)
    DeleteDC (hdcBitmap)
End Sub

'*********************************************************************************
'                                      VARIOUS
'*********************************************************************************

' Returns the maximum of two longs
Private Function Max(Value1 As Long, Value2 As Long) As Long
    Max = IIf(Value1 > Value2, Value1, Value2)
End Function

Private Sub GetHiLoWord(lParam As Long, LOWORD As Long, HIWORD As Long)
    LOWORD = lParam And &HFFFF&
    HIWORD = lParam \ &H10000 And &HFFFF&
End Sub

Private Sub SetHiLoWord(lParam As Long, LOWORD As Long, HIWORD As Long)
    lParam = LOWORD
    lParam = lParam Or (HIWORD * &H10000)
End Sub

' Checks the style of windows (if XP)
Private Function HasFlatMenu() As Boolean
    Dim x As Boolean

    SystemParametersInfo SPI_GETFLATMENU, ByVal 0&, x, ByVal 0&
    HasFlatMenu = x
End Function
