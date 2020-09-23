Quick Start
-----------
1)Add modOwnerDrawnMenu.bas to your project
2)Create menus using the VB menu editor
3)Edit WM_MENUCHAR message in MyWndProc, to process menubar keyboard shortcuts
4)Optional add picture paths to the menu array


Accessibility Issues
--------------------
For those who deal with MSAA, this menu might be useful.
It provides MSAA information and can be used with screen readers
However, there is a small problem. Submenus, including the top level menus,
do not provide MSAA information. 

I use this menu in a self voicing application and I found a way
to work around this problem by speaking when WM_DRAWITEM is sent
for a Submenu. For more information see commented code in MyWndProc

Although there are ways to solve this problem I will not be considering
them for now