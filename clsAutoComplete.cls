VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************************
'* Copyright © 2002 Corné van Dyk *
'* All Rights Reserved            *
'**********************************

'*****************************************************************
'***************************DISCLAIMER****************************
'*****************************************************************
'* No waranty of any kind, either explicit or implied is made as *
'* to the usability or the functionality of this code.           *
'* By using this code you indemnify the author of any and all    *
'* liability related to the use, or inability to use this code.  *
'*****************************************************************

'*****************************************************************************
'* THANK YOU VERY MUCH TO EVERYONE WHO HAS VOTED FOR THIS CONTROL AT PSC     *
'* http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=34618 *
'*****************************************************************************

'FEEDBACK
'========
'c@dtdn.com

'UPDATES
'=======
'http://www.dtdn.com/dev

'CREDITS
'=======
'I wish to extend credit to the following people for their feedback,
'suggestions and support in making clsAutoComplete better:
'John Couture   - Case Statement for optimal functionality
'John Couture   - Enterkey functionality
'Dave Buckner   - LimitToList functionality
'Gurmeet        - Backspace functionality
'Dave Buckner   - ForceListUsage functionality
'Baart Goossens - Multi Addition Bug
'Kaze           - Mouse Pointer Disapperance Fix for Win XP

'Thank you very much for your support!

'DESCRIPTION
'===========
'EASY TO IMPLEMENT FAST API Based Class featuring AutoComplete, AutoDropdown,
'Full Featured RowSource, ForceListUsage, LimitToList and AutoUpdateList.

'SYNOPSIS
'========
'This class is used to EASILY add AutoCompletion such as in Internet
'Explorer or MS Access to VB Combo Boxes.  It makes use of API calls
'which makes it EXTREMELY efficient.  It also gives you the ability
'to control various properties:
'1.  The AutoDropdown property which controls weather or not the dropdown
'    list is automatically dropped down when the user starts typing.
'2.  The RowSource and RowSourceField properties which allows the application
'    to use any Data control as a source for populating the combo control.
'3.  The ForceListUsage property which allows your application to force the
'    user to only use values in the dropdown list.  This property is overriding
'    to both LimitToList and AutoUpdateList.
'4.  The LimitToList property which controls weather or not the only values
'    accepted by the control are those in the dropdown list.  This property is
'    overriding to AutoUpdateList.
'5.  The AutoUpdateList property which allows your application to add new
'    values that are not in the dropdown list to the list.

'USAGE
'=====
'The demo project should be self explanatory, but basically all you
'have to do is add the class to your project and upon the Load event
'of your form, simply initialize a class instance for each of your
'combo box controls.
'All properties can easily be set in code.  In the example project I tied
'them to check boxes.  Play around with it to see how it works.

'REVISION HISTORY
'================
'1.00   - Original class featuring AutoComplete and the AutoDropdown option
'1.10   - Changed the range of If statements in the KeyUp sub to a more
'         optimal Case statement.
'       - Added functionality to make the Enterkey behave the same as the
'         Tab key in that it accepts the selected value and moves to the
'         next field.
'       - Added LimitToList functionality.  When the limit to list property
'         is set to True, the code will blank out any invalid values in the
'         control on the loss of focus.  If however the LimitToList property
'         is set to False, the value will remain.
'       - Added AutoUpdateList functionality.  When an invalid value is
'         entered and the LimitToList property is set to False this
'         property becomes usefull.  If set to True in such as case, it will
'         automatically add the newly added value to the list of the control.
'         It is important to notice that setting the "Sorted" property of your
'         ComboBox at design time has an effect on how this functionality
'         work.  If Sorted is set to True, the value is added to the list and
'         the list will automatically be sorted.  If however Sorted is set to
'         False, the value is simply added to the end of the original list.
'         It is recommended that you use Sorted = True.
'       - Swapped Integer variables for Long variables to enhance speed as
'         the use of Integers cause VB to do an internal conversion of the
'         value prior to using it.  A Long value being 32 bit does not need
'         to be converted.  Though the effect is not generally noticable, it
'         is all part of the Optimal Code Principle.  The tradeoff in this
'         case is 3 bytes per variable in extra size, i.e. in this Class's
'         case the class would occupy an extra 9 bytes but execution speed
'         would be quicker.
'1.20   - Added ForceListUsage property to allow the parent application to
'         force the user to use values from the dropdown list.  This property
'         is overriding to the AutoUpdateList property and also replaces the
'         LimitToList property functionality when it is used.
'       - Added RowSource for populating the dropdown list.  Made use of a
'         boolean flag to eliminate the need for the RowSourceField to be set
'         before the RowSource.  This allows the user to take any data control
'         and use it to populate the class control.
'       - Fixed a bug that caused valid values selected, while the
'         AutoUpdateList = True to be added to the dropdown multiple times.
'1.21   - Added code to prevent the mouse cursor from being hidden after the
'         combo was dropped down under Windows XP.

Public Event Message(ByVal strMessage As String)    'event to send status msg to parent

Private WithEvents memCBO As ComboBox       'member var to hold combo
Attribute memCBO.VB_VarHelpID = -1

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any _
        ) As Long                           'internal windows messaging API
Private Declare Function ShowCursor Lib "user32" ( _
        ByVal bShow As Long) As Long        'internal windows cursor API

Private Const CB_ERR = (-1)                 'combo error code
Private Const CB_FINDSTRING = &H14C         'hex value to invoke find
Private Const CB_FINDSTRINGEXACT = &H158    'hex value to invoke find with exact matching
Private Const CB_SHOWDROPDOWN = &H14F       'hex value to drop down combo

Private membAutoDropdown As Boolean         'hold AutoDropdown property
Private membLimitToList As Boolean          'hold LimitToList property
Private membAutoUpdateList As Boolean       'hold AutoUpdateList property
Private membForceListUsage As Boolean       'hold ForceListUsage property
Private memstrRowSourceField As String      'hold RowSourceField property
Private memdbRowSource As Data              'hold RowSource data control property
Private bRowSourceIncomplete As Boolean     'used to make RowSource and
                                            'RowSourceField property assignment
                                            'sequence irrelevent
'variables for v1.30 to add ADO functionality
'Private memconADOConnection As ADODB.Connection
'Private memrsADORecordSet As ADODB.Recordset
'Private memstrADODB As String
'Private memstrADOTable As String
'Private memstrADOWhereClause As String
'Private memstrADOOrder As String
'Private memstrADOSQL As String
'I decided to use Long values as apposed to Integer because a Long value is
'a 32 bit value.  All Integers that are used have to be converted internally
'by VB prior to using the values.  As a result, my code will be a couple of
'bytes largers, but the execution time will be quicker.
Private lngSelStart As Long                 'var to hold combo.SelStart val
Private lngSelLength As Long                'var to hold combo.SelLength val
Private lngMatchLoc As Long                 'var to hold mathing index val
Private strText As String                   'var to hold search string
Private lngKeyDownCount As Long             'var to hold number of keys

Public Property Let AutoDropdown(bAutoDropdown As Boolean)
    membAutoDropdown = bAutoDropdown        'set autodropdown property
End Property

Public Property Get AutoDropdown() As Boolean
    AutoDropdown = membAutoDropdown         'return autodropdown property
End Property

Public Property Let LimitToList(bLimitToList As Boolean)
    membLimitToList = bLimitToList          'set limittolist property
End Property

Public Property Get LimitToList() As Boolean
    LimitToList = membLimitToList           'return limittolist property
End Property

Public Property Let AutoUpdateList(bAutoUpdateList As Boolean)
    membAutoUpdateList = bAutoUpdateList    'set autoupdatelist property
End Property

Public Property Get AutoUpdateList() As Boolean
    AutoUpdateList = membAutoUpdateList     'return autoupdatelist property
End Property

Public Property Let ForceListUsage(bForceListUsage As Boolean)
    If bForceListUsage Then                 'is forcelistusage being set to true
        If membAutoUpdateList Then          'if autoupdate list is also true
            'send a warning that autoupdatelist no longer has any effect
            RaiseEvent Message("WARNING!  AutoUpdateList property is set but will not longer have any effect...")
        End If
    End If
    membForceListUsage = bForceListUsage    'set forcelistusage property
End Property

Public Property Get ForceListUsage() As Boolean
    ForceListUsage = membForceListUsage     'return forcelistusage property
End Property

Public Property Let RowSourceField(strRowSourceField As String)
    memstrRowSourceField = Trim$(strRowSourceField) 'set rowsourcefield property
    If bRowSourceIncomplete Then            'if the RowSource is already set
        bRowSourceIncomplete = False        'reset flag
        BuildDropdownFromRowSource          'build the dropdown
    End If
End Property

Public Property Get RowSourceField() As String
    RowSourceField = Trim$(memstrRowSourceField)    'return rowrourcefield property
End Property

Public Property Set RowSource(dbRowSource As Data)
    memCBO.Clear                            'clear dropdown
    Set memdbRowSource = dbRowSource        'set datasource property
    If memstrRowSourceField = "" Then       'if RowSourceField not set
        bRowSourceIncomplete = True         'set flag
    ElseIf Not dbRowSource Is Nothing Then  'otherwise RowSourceField is set
                                            'check if data control is valid
        BuildDropdownFromRowSource          'build the dropdown
    End If
End Property

Public Property Get RowSource() As Data
    Set RowSource = memdbRowSource          'return datasource property
End Property

Public Property Set LinkedComboBox(cboTargetComboBox As ComboBox)
    Set memCBO = cboTargetComboBox          'set the linked combo
End Property

Public Property Get LinkedComboBox() As ComboBox
    Set LinkedComboBox = memCBO             'return the linked combo
End Property

Private Sub Class_Initialize()
    membAutoDropdown = True                 'dropdown by default
    membLimitToList = False                 'do not limit to list by default
    membAutoUpdateList = False              'do not update list by default
    membForceListUsage = False              'do not force list usage by default
    memstrRowSourceField = ""               'initialize var
    Set memdbRowSource = Nothing            'initialize var
    bRowSourceIncomplete = False            'initialize var
    lngSelStart = 0                         'initialize var
    lngSelLength = 0                        'initialize var
    lngMatchLoc = 0                         'initialize var
    strText = ""                            'initialize var
    lngKeyDownCount = 0                     'initialize var
End Sub

Private Sub memCBO_KeyDown(KeyCode As Integer, Shift As Integer)
    lngKeyDownCount = lngKeyDownCount + 1   'incremented when key held down
End Sub

Private Sub memCBO_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    'do nothing if certain special keys is pressed
    '8 = BackSpace      16 = Shift          17 = Ctrl           18 = Alt
    '27 = Escape        33 = Page Up        34 = Page Down      35 = End
    '36 = Home          37 = Left Arrow     39 = Right Arrow    45 = Insert
    '91 = Windows Key   112 = F1            113 = F2            114 = F3
    '115 = F4           116 = F5            117 = F6            118 = F7
    '119 = F8           120 = F9            121 = F10           122 = F11
    '123 = F12
    Case 8, 16, 17, 18, 27, 33 To 37, 39, 45, 91, 112 To 123
        lngKeyDownCount = 0                 'key raised so reset
        Exit Sub
    Case vbKeyReturn
        lngKeyDownCount = 0                 'key raised so reset
        SendKeys "{TAB}"                    'duplicate Tab functionality
    Case Else
        With memCBO
            lngSelStart = .SelStart         'save combo.SelStart
            lngSelLength = .SelLength       'save combo.SelLength
            strText = .Text                 'save combo.Text
            lngMatchLoc = FindInList(strText, , False)  'search for match
            If lngMatchLoc <> CB_ERR Then   'if match found
                ShowCursor False
                ShowDropDown True           'drop down combo
                ShowCursor True
                .ListIndex = lngMatchLoc    'position cursor in dropped list
                .Text = .List(lngMatchLoc)  'set combo text
                .SelStart = lngSelStart     'set selected text
                .SelLength = Len(.Text) - lngSelStart
                lngKeyDownCount = 0         'key raised so reset
            Else                            'no match found
                If membForceListUsage Then  'force list usage
                    If lngKeyDownCount > 1 Then 'key was held down
                        .Text = Left(.Text, Len(.Text) - lngKeyDownCount)   'delete held down text
                        lngSelStart = .SelStart         'save combo.SelStart
                        lngSelLength = .SelLength       'save combo.SelLength
                        strText = .Text                 'save combo.Text
                        lngMatchLoc = FindInList(strText, , False)  'search for match
                        If lngMatchLoc <> CB_ERR Then   'if match found
                            ShowCursor False
                            ShowDropDown True           'drop down combo
                            ShowCursor True
                            .ListIndex = lngMatchLoc    'position cursor in dropped list
                            .Text = .List(lngMatchLoc)  'set combo text
                            .SelStart = lngSelStart     'set selected text
                            .SelLength = Len(.Text) - lngSelStart
                        End If
                    Else                    'key not held down
                        SendKeys "{BACKSPACE}"  'delete last character typed
                    End If
                    'send an error message to the parent application
                    RaiseEvent Message("Invalid value.  Valid value must be in the list...")
                Else
                    ShowDropDown False      'hide combo drop down
                    .Text = strText         'set combo text to original text
                    .SelStart = lngSelStart 'set selected text
                    .SelLength = lngSelLength
                End If
                lngKeyDownCount = 0         'key raised so reset
            End If
        End With
    End Select
End Sub

Private Sub memCBO_LostFocus()
    With memCBO
        If membLimitToList Then             'if you need to limit values
            .Text = .List(lngMatchLoc)      'set combo to index value if TAB pressed
        ElseIf membAutoUpdateList Then      'if val need to be added to list
            If Trim(.Text) <> "" Then       'ensure it is not blank
                If lngMatchLoc = CB_ERR Then 'check value from last match search
                    .AddItem strText        'add to list
                End If
            End If
        End If
    End With
End Sub

Private Function FindInList(strFindThis As String, _
        Optional lngIndexStart As Long = -1, _
        Optional bExactMatchRequired As Boolean = False) _
        As Long                             'function to search for match
                                            'if found, it returns the index
                                            'else it returns CB_ERR
    Dim wMsg As Long                        'temp var to hold type of search
    'check if exact match is required & set search type accordingly
    wMsg = IIf(bExactMatchRequired, CB_FINDSTRINGEXACT, CB_FINDSTRING)
    strFindThis = strFindThis & Chr(0)
    FindInList = SendMessage(memCBO.hwnd, wMsg, lngIndexStart, _
            ByVal strFindThis)              'invoke API search of combo
End Function

Private Sub ShowDropDown(bShowDropdown As Boolean)
    If membAutoDropdown Then                'if autodropdown property is set
        'make API call to drop down combo
        SendMessage memCBO.hwnd, CB_SHOWDROPDOWN, bShowDropdown, 0
        memCBO.Parent.MousePointer = 0      'show mouse again under XP
    End If
End Sub

Private Sub BuildDropdownFromRowSource()    'build the actual dropdown
    If Not memdbRowSource Is Nothing Then   'if rowsource is valid
        memdbRowSource.Refresh              'get recordset
        If Not memdbRowSource.Recordset Is Nothing Then     'if recordset is valid
            memdbRowSource.Recordset.MoveLast               'get record count
            If Not memdbRowSource.Recordset.RecordCount = 0 Then    'if records exist
                memdbRowSource.Recordset.MoveFirst          'ready to load data
                While Not memdbRowSource.Recordset.EOF      'get all records
                    memCBO.AddItem memdbRowSource.Recordset.Fields(memstrRowSourceField).Value  'add value
                    memdbRowSource.Recordset.MoveNext       'next record
                Wend
            End If
        End If
    End If
End Sub
