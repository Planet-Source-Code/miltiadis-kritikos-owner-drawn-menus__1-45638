VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Owner Drawn Menu Demo"
   ClientHeight    =   3675
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Optional constants (used for assigning images to the menu items)
Private Const MI_FILE = 0
Private Const MI_NEW = 1
Private Const MI_OPEN = 2
Private Const MI_QUIT = 3
Private Const MI_EDIT = 4
Private Const MI_CUT = 5
Private Const MI_COPY = 6
Private Const MI_PASTE = 7

Private Const FONT_NAME = "Comic Sans MS"
Private Const FONT_SIZE = 22
Private Const FONT_BOLD = False
Private Const FONT_ITALIC = False

' When the form loads
Private Sub Form_Load()
    CreateMenu FONT_NAME, FONT_SIZE, FONT_BOLD, FONT_ITALIC
End Sub

' Creates a menu
Private Sub CreateMenu(FontName As String, FontSize As Long, _
                       Optional FontBold As Boolean = False, _
                       Optional FontItalic As Boolean = False)
    
    Dim BasePath As String  ' Path of the images
    
    ' Construct the images path
    ' Check for "\" in case BasePath is a root folder (like C:\)
    ' If we don't check this we might end up with something like
    ' -->   C:\\Images\
    ' Which will point to an invalid folder and the images will
    ' not be loaded (no errors are generated in that case)
    BasePath = App.Path
    BasePath = BasePath & IIf(Right(BasePath, 1) = "\", "Images\", "\Images\")
    
    ' Add the menu bitmap paths
    Menus(MI_FILE).PicPath = BasePath & "File.bmp"
    Menus(MI_NEW).PicPath = BasePath & "New.bmp"
    Menus(MI_OPEN).PicPath = BasePath & "Open.bmp"
    Menus(MI_QUIT).PicPath = BasePath & "Quit.bmp"
    Menus(MI_EDIT).PicPath = BasePath & "Edit.bmp"
    Menus(MI_CUT).PicPath = BasePath & "Cut.bmp"
    Menus(MI_COPY).PicPath = BasePath & "Copy.bmp"
    Menus(MI_PASTE).PicPath = BasePath & "Paste.bmp"
    
    ' Add the menu to the window
    ' The second parameter (50) depends on the font height
    ' and is used for increasing the size of the Menubar
    ' when using OwnerDrawn menus. You should adjust it to
    ' the font size you use or make a function to generate it
    ' for you. I couldn't find any other way to increase the
    ' height of OwnerDrawn Menubars, I posted my question on
    ' the Microsoft Newsgroups but I got no reply.
    
    ' *BUG* Although this method of adding a bitmap does the
    ' job, when you make the width of the window too small so
    ' that the second top level menu goes under the first
    ' (in order to be visible), the height is not retained for
    ' the second menu item
    AddEmptyBitmapToMenuBar hWnd, 50
    AddMenus hWnd, FontName, FontSize, FontBold, FontItalic
End Sub

' When the user clicks on Quit
Private Sub mnuFileExit_Click()
    Unload Me
End Sub

' When the user clicks on New
Private Sub mnuFileNew_Click()
    MsgBox "New", vbInformation, "mnuFileNew"
End Sub

' When the user clicks on Open
Private Sub mnuFileOpen_Click()
    MsgBox "Open", vbInformation, "mnuFileOpen"
End Sub
