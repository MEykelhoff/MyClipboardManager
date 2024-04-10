; *****
;
; Author: Maarten Eykelhoff
;
; *****

#include <AutoItConstants.au3>
#include <ButtonConstants.au3>
#include <Clipboard.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <Misc.au3>
#include <ScreenCapture.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>

#include <Winapi.au3>
#include <WinAPIGdi.au3>
#include <WindowsConstants.au3>

; Constants
Global $version = "2024-04-10"
Global $AppTitle = "My Clipboard Manager v. " & $version
Global $clipboardDataIdColumn = 0
Global $clipboardDataColumn = 3
Global $antiScreenlockPeriod = 290000 ; Press key every n milliseconds.
; File locations
Global $directory = @AppDataDir & "\MyClipboardManager\"
Global $preloadEntriesFileLocation = $directory & "preloadedEntries.txt"
Global $settingsFileLocation = $directory & "settings.txt"
; Global Variables
Global  $lastAddedToCM = ""         ;Is used to decrease the amount of times retrieving the top entry. UBound is bad for perfomance.
Global  $clipboardData[0][4]        ;Listview cannot hold large strings, this will hold the actual data. Holds arrays with id, timestamp, starred and data.
Global  $listViewLastSelectedIndices = -1
Global  $clipboardDataID = 0
Global  $listViewFontSize = 11
Global  $previewFontSize = 14
Global  $antiScreenlockTimer = TimerInit()
Global  $antiScreenlockIsEnabled = True
Global  $lastSearchString = ""
_loadSettingsFromFile()

; Create ClipboardGUI
Opt("GUIOnEventMode", 1)
Global $clipboardGUI = GUICreate($AppTitle, 650, 420, 250, 146, $WS_MAXIMIZEBOX + $WS_MINIMIZEBOX + $WS_SIZEBOX + $WS_CAPTION)

Global $menuMenu = GUICtrlCreateMenu("Menu")
Global $menuSetListViewFontSize = GUICtrlCreateMenuItem("Set font size of the clipboard manager listview", $menuMenu)
Global $menuSetPreviewFontSize = GUICtrlCreateMenuItem("Set preview font size", $menuMenu)
Global $menuToggleAntiScreenlock = GUICtrlCreateMenuItem("Enable/Disable antiscreenlock", $menuMenu)
Global $menuClearClipboardData = GUICtrlCreateMenuItem("Clear Clipboard data (0)", $menuMenu)
Global $menuEditPreloadedDataFile = GUICtrlCreateMenuItem("Edit preloaded clipboard manager data entries", $menuMenu)
Global $menuExit = GUICtrlCreateMenuItem("Exit (Alt + F12)", $menuMenu)

Global $clipboardListView = GUICtrlCreateListView("", 5, 5, 640, 200)             ;, $GUI_BKCOLOR_LV_ALTERNATE)
Global $pasteButton = GUICtrlCreateButton("Paste", 5, 215, 50, 30, $BS_DEFPUSHBUTTON)                        ;( "text", left, top [, width [, height [, style = -1 [, exStyle = -1]]]] )
Global $storeInClipboardButton = GUICtrlCreateButton("To Clipboard", 55, 215, 80, 30)
Global $searchButton = GUICtrlCreateButton("Search", 135, 215, 70, 30)
Global $filterStarredCheckbox = GUICtrlCreateCheckbox("Starred only", 210, 215, 100, 30)
Global $deleteEntryButton = GUICtrlCreateButton("Delete entry", 310, 215, 80, 30)
Global $selectTopRowButton = GUICtrlCreateButton("^", 500, 215, 20, 20)
Global $previewLabel = GUICtrlCreateLabel("Preview        (Ctrl + Enter for linebreak)", 5, 250)
Global $previewEdit = GUICtrlCreateEdit("", 5, 265, 640, 125, $WS_VSCROLL + $WS_BORDER)
_GUICtrlListView_AddColumn($clipboardListView, "ID", 0)
_GUICtrlListView_AddColumn($clipboardListView, "Time", 80)
_GUICtrlListView_AddColumn($clipboardListView, "*", 20)
_GUICtrlListView_AddColumn($clipboardListView, "Data", 2400)

; Set gui resizing options
GUICtrlSetResizing($clipboardListView, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKHEIGHT)
GUICtrlSetResizing($pasteButton, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($storeInClipboardButton, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($searchButton, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($filterStarredCheckbox, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($deleteEntryButton, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($selectTopRowButton, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($previewLabel, $GUI_DOCKTOP + $GUI_DOCKLEFT + $GUI_DOCKSIZE)
GUICtrlSetResizing($previewEdit, $GUI_DOCKBORDERS)

; Set gui events
GUICtrlSetOnEvent($menuSetListViewFontSize, "_setListViewFontSize")
GUICtrlSetOnEvent($menuSetPreviewFontSize, "_setPreviewFontSize")
GUICtrlSetOnEvent($menuClearClipboardData, "_clearClipboardManager")
GUICtrlSetOnEvent($menuToggleAntiScreenlock, "_toggleAntiScreenlockEnabled")
GUICtrlSetOnEvent($menuEditPreloadedDataFile, "_EditPreloadedDataFile")
GUICtrlSetOnEvent($menuExit, "_Exit")

GUISetOnEvent($GUI_EVENT_CLOSE, "_hideClipboardManager")
GUICtrlSetOnEvent($storeInClipboardButton, "_storeSelectedRowToClipboard")
GUICtrlSetOnEvent($searchButton, "_clipboardManagerStartSearch")
GUICtrlSetOnEvent($filterStarredCheckbox, "_updateClipboardManagerGUI")
GUICtrlSetOnEvent($deleteEntryButton, "_deleteClipboardEntry")
GUICtrlSetOnEvent($selectTopRowButton, "_selectTopRow")
GUICtrlSetOnEvent($pasteButton, "_pasteFromClipboardManager")
GUIRegisterMsg($WM_NOTIFY, "WM_Notify_Events") ; To handle mouse inputs
WinSetOnTop($clipboardGUI, "", 1) ; GUI is on top of all screens
GUISetState(@SW_HIDE)

; Set gui styling
_setListViewFontSize($listViewFontSize)
_setPreviewFontSize($previewFontSize)
GUICtrlSetTip($pasteButton, "Enter")
GUICtrlSetTip($storeInClipboardButton, "Ctrl + Alt + c")
GUICtrlSetTip($searchButton, "Ctrl + f")
GUICtrlSetTip($selectTopRowButton, "Select top row")

; Create Tray
Opt("TrayMenuMode", 3) ;[Function: Hides the default AutoIt system tray menu.]
TrayCreateItem("***** " & $AppTitle & " *****")
Global $tbClipboardManager = TrayCreateItem("Show Clipboard Manager (Ctrl + Alt + v)")
$tbShowFiles = TrayCreateItem("Show My Clipboard Manager files")
TrayCreateItem("")
Global $tbExit = TrayCreateItem("Exit (Alt + F12)")

; Set hotkeys
HotKeySet("^!v", "_showClipboardManager")
HotKeySet("!{F12}", "_Exit")

; Preload My Clipboard Manager data
If FileExists($preloadEntriesFileLocation) Then
	$preloadEntriesFileHandle = FileOpen($preloadEntriesFileLocation)
	If @error = 0 Then
		While True
			$line = FileReadLine($preloadEntriesFileHandle)
			If @error <> 0 Then ExitLoop
			_toClipboardManager($line)
		WEnd
		FileClose($preloadEntriesFileHandle)
	EndIf
EndIf

TrayTip("My Clipboard manager", "   Hello, I'm started :)" & @CRLF & "Version: " & $version, 2, 1)
mainLoop()

Func mainLoop()
	Local $trayInputLoopTimer = TimerInit()
	Local $mainLoopTimer = TimerInit()
	Local $hotkeysSet = False
	While 1
		If TimerDiff($trayInputLoopTimer) > 200 Then ; Every n milliseconds. Larger numbers will reduce the tray menu responsiveness
			; Handle input through the tray menu.
			$trayInputLoopTimer = TimerInit()
			Switch TrayGetMsg()
				Case $tbExit
					_Exit()
				Case $tbClipboardManager
					_showClipboardManager()
				Case $tbShowFiles
					If FileExists($directory) Then
						Run(@WindowsDir & "\EXPLORER.EXE /n,/e," & $directory)
					Else
						TrayTip("My Clipboard manager", "No My Clipboard manager directory found @ "& $directory, 20, 1)
					EndIf
			EndSwitch
		EndIf

		If TimerDiff($mainLoopTimer) > 2000 Then ; Every n milliseconds.
			$mainLoopTimer = TimerInit()
			; Monitor clipboard changes
			If _clipboardDataChanged() Then _toClipboardManager(ClipGet())

			If WinActive($AppTitle) <> 0 And $hotkeysSet = False Then	;My Clipboard manager is active
				HotKeySet("^f","_clipboardManagerStartSearch")
				HotKeySet("^!c", "_storeSelectedRowToClipboard")
				$hotkeysSet = True
			ElseIf WinActive($AppTitle) = 0 And $hotkeysSet = True Then ;My Clipboard manager is not active
				HotKeySet("^f")
				HotKeySet("^!c")
				$hotkeysSet = False
			Endif
			; Handle AntiScreenlock
			_antiScreenlock()
		EndIf
	WEnd
EndFunc   ;==>mainLoop

; ***** Functions Menu
Func _setListViewFontSize($size = "")
	_hideClipboardManager()
	if Not IsDeclared("size") Then $listViewFontSize = InputBox("Set clipboard manager listview font size", "New font size", $listViewFontSize)
	GUICtrlSetFont($clipboardListView, $listViewFontSize, 0, 0, "Consolas")
EndFunc	;==>_setListViewFontSize

Func _setPreviewFontSize($size = "")
	_hideClipboardManager()
	if Not IsDeclared("size") Then $previewFontSize = InputBox("Set preview font size", "New font size", $previewFontSize)
	GUICtrlSetFont($previewEdit, $previewFontSize, 0, 0, "Courier New")
EndFunc   ;==>_setPreviewFontSize

Func _clearClipboardManager()
	$buttonPressed = MsgBox(4, "Delete all entries", "Are you sure you want to delete all entries?")
	_showClipboardManager()
	If $buttonPressed = 7 Then Return
	Local $newArray[0][3]
	$clipboardData = $newArray
	ClipPut("")
	$lastAddedToCM = ""
	_updateClipboardManagerGUI()
EndFunc   ;==>_clearClipboardManager

Func _EditPreloadedDataFile()
   If Not FileExists($preloadEntriesFileLocation) Then
		_FileCreate($preloadEntriesFileLocation)
		If @error <> 0 Then
			TrayTip("Preload data file","Unable to create file." & $preloadEntriesFileLocation, 2, 3)
			Return
		EndIf
		TrayTip("Preload data file","File not found. New file created at: " & $preloadEntriesFileLocation, 2, 1)
   EndIf
   Run(@comspec & ' /c start ' & $preloadEntriesFileLocation)
   If @error <> 0 Then
	   TrayTip("Preload data file","Unable to open file." & $preloadEntriesFileLocation, 2, 3)
	   Return
   EndIf
   TrayTip("Preload data file","Add entries." &@CRLF& "Every line in the file will be an entry in the clipboard manager.", 20, 1)
   _hideClipboardManager()
EndFunc	;==>_EditPreloadedDataFile

Func _Exit()
	_hideClipboardManager()
	If MsgBox(4, "Exit My Clipboard Manager", "Are you sure you want to exit My Clipboard Manager?") = $IDNO Then Return
	_saveSettingsToFile()
	TrayTip($AppTitle, "Bye bye", 10, 1)
	Sleep(2000)
	Exit
EndFunc   ;==>_Exit

Func _saveSettingsToFile()
	Local $settings = ["previewFontSize", "listViewFontSize"]
	Local $settingsFileHandle = FileOpen($settingsFileLocation, 2)
	If @error = 0 Then
		For $setting In $settings
			$line = $setting & ":" & Eval($setting)
			FileWriteLine($settingsFileHandle, $line)
		Next
		FileClose($settingsFileHandle)
	EndIf
EndFunc

Func _loadSettingsFromFile()
	If FileExists($settingsFileLocation) Then
		$settingsFileHandle = FileOpen($settingsFileLocation)
		If @error = 0 Then
			While True
				$line = FileReadLine($settingsFileHandle)
				If @error <> 0 Then ExitLoop
				$lineSplitted = StringSplit($line, ":")
				if($lineSplitted[0] = 2) Then
					Assign($lineSplitted[1],$lineSplitted[2])
				EndIf
			WEnd
			FileClose($settingsFileHandle)
		EndIf
	EndIf
EndFunc

; ***** Functions ClipboardManager

; Handle mouse inputs in the clipboardManager GUI
Func WM_Notify_Events($hWndGUI, $MsgID, $wParam, $lParam)
	#forceref $hWndGUI, $MsgID, $wParam
	Local $tagNMHDR, $event
	If $wParam = $clipboardListView Then
		$tagNMHDR = DllStructCreate("hwnd hWndFrom; uint_ptr IDFrom; int Code", $lParam)
		If @error Then Return
		$event = DllStructGetData($tagNMHDR, 3)
		;Handle mouse inputs
		Switch $event
			Case $NM_CLICK
				_updatePreview()
			Case $NM_DBLCLK ;Handle doubleclick
				_pasteFromClipboardManager()
			Case $NM_RCLICK
				_clipboardManagerStarItem()
		EndSwitch
	EndIf
EndFunc   ;==>WM_Notify_Events

Func _showClipboardManager()
	ControlFocus($clipboardGUI, "", $clipboardListView)
	GUISetState(@SW_HIDE)   ;added to set focus on the clipboardmanager window.
	_updateClipboardManagerGUI()
	GUISetState(@SW_SHOW)
EndFunc   ;==>_showClipboardManager

Func _updateClipboardManagerGUI()
	; Save the original selected row
	$selectedIndices = _GUICtrlListView_GetSelectedIndices($clipboardListView, True)
	_updateListView($clipboardData)
	If $selectedIndices[0] <> 0 Then
		;Select the original selected row
		_GUICtrlListView_SetItemSelected($clipboardListView, $selectedIndices[1])
		_GUICtrlListView_SetItemFocused($clipboardListView, $selectedIndices[1])
		_updatePreview()
	EndIf
	; update number of entries label
	GUICtrlSetData($menuClearClipboardData, "Clear Clipboard data (" & UBound($clipboardData) & ")")
EndFunc   ;==>_updateClipboardManagerGUI

Func _hideClipboardManager()
	GUISetState(@SW_HIDE)
EndFunc   ;==>_hideClipboardManager

; Check for selection changes in the listview
Func _listviewSelectionChanged()
	$selectedIndices = _GUICtrlListView_GetSelectedIndices($clipboardListView)
	If $listViewLastSelectedIndices = $selectedIndices Then Return False
	$listViewLastSelectedIndices = $selectedIndices
	Return True
EndFunc   ;==>_listviewSelectionChanged

; Repopulate the listview with the given data array. Star filter is applied.
Func _updateListView($listviewDataArray)
	_filterStarredData($listviewDataArray)
	_ArraySort($listviewDataArray, 1, 0, 0, 1)	; By time
	_GUICtrlListView_DeleteAllItems($clipboardListView)
	_GUICtrlListView_AddArray($clipboardListView, $listviewDataArray) ;Requires array to accept | characters
EndFunc ;==>_updateListView

Func _updatePreview($text = "")
	If $text = "" Then $text = _getSelectedTextFromClipboardManager()
	GUICtrlSetData($previewEdit, "" & $text)
EndFunc   ;==>_updatePreview

; Check for change in the clibboard
Func _clipboardDataChanged()
	$cbText = ClipGet()
	Return $cbText <> "" And $cbText <> $lastAddedToCM
EndFunc   ;==>_clipboardDataChanged

; Put $data to clipboard manager
Func _toClipboardManager($text)
	; Add the data if it is not empty.
	If $text <> "" Then
		$lastAddedToCM = $text
		; Loading the preview for long Strings will take a lot of time. Therefore, long Strings are cropped.
		$maxNumberOfCharacters = 120000
		If (StringLen($text) > $maxNumberOfCharacters) Then
			TrayTip("To clipboard manager", "WARNING"& @CRLF &"Too many characters will impact performance." & @CRLF & "Text is capped to " & $maxNumberOfCharacters & " characters.", 2, 2)
			$text = StringLeft($text, $maxNumberOfCharacters) & @CRLF & @CRLF & "[NOTE: cropped to " & $maxNumberOfCharacters & " characters]"
		EndIf
		; Search if there is a row with the same data. Update existing data row or create new.
		$iSubItem = 3 ;indicates the column to search
		$i = _ArraySearch($clipboardData, $text, 0, 0, 0, 0, 0, $iSubItem)
		If $i <> -1 Then
			$clipboardData[$i][1] = _NowTime(5)
		Else
			Local $newDataRow[1][4] = [[$clipboardDataID, _NowTime(5), "", $text]]
			_ArrayAdd($clipboardData, $newDataRow)
			$clipboardDataID += 1
		EndIf
		_updateClipboardManagerGUI()
	EndIf
EndFunc   ;==>_toClipboardManager

; Get text from the clipboardManager, put to clipboard and paste it. The pasted item is put at the top.
Func _pasteFromClipboardManager()
	_hideClipboardManager()
	$selectedRowDataId = _getSelectedDataIdFromListView()
	If $selectedRowDataId <> -1 Then
		$i = _ArraySearch($clipboardData, $selectedRowDataId, 0, 0, 0, 0, 0, $clipboardDataIdColumn)
		If $i <> -1 Then
			_ClipBoard_Open(0) ;
			_ClipBoard_SetData($clipboardData[$i][$clipboardDataColumn], $CF_TEXT)
			_ClipBoard_Close()
			Send("^v")  ;perform a Ctrl+v
			Return
		EndIf
	EndIf
	TrayTip($AppTitle, "Nothing pasted", 5)
EndFunc   ;==>_pasteFromClipboardManager

; Because the listview text is capped, the text must be fetched from $clipboardData Array.
; Retrieve the dataID from the listview to determine which row is selected.
; Based on the dataID, the actual text is retrieved from the $clipboardData Array.
Func _getSelectedTextFromClipboardManager()
	$selectedRowDataID = _getSelectedDataIdFromListView()
	; Within the $clipboardData array, search if there is a row with the same dataID.
	$arrayIndex = _ArraySearch($clipboardData, $selectedRowDataID, 0, 0, 0, 0, 0, $clipboardDataIdColumn)
	Return $arrayIndex <> -1 ? $clipboardData[$arrayIndex][3] : ""
EndFunc   ;==>_getSelectedTextFromClipboardManager

; Return the dataId of the selected row in the listview. Return -1 if no row is selected.
Func _getSelectedDataIdFromListView()
	$selectedListViewArray = _GUICtrlListView_GetItemTextArray($clipboardListView, -1) ;-1 will attempt to get the Currently Selected item.
	Return ($selectedListViewArray[0] > 0 ? $selectedListViewArray[1] : -1)
EndFunc

; stores the selected row value to the clipboard.
Func _storeSelectedRowToClipboard()
	$selectedRowText = _getSelectedTextFromClipboardManager()
	if $selectedRowText = "" Then Return
	ClipPut($selectedRowText)
	_hideClipboardManager()
	TrayTip("My Clipboard manager", "Copied text to clipboard.", 2, 1)
EndFunc   ;==>_storeSelectedRowToClipboard

; Retrieve the dataID from the listview to determine which row is selected.
Func _deleteClipboardEntry()
	$selectedRowDataID = _getSelectedDataIdFromListView()
	If $selectedRowDataID <> -1 Then
		$arrayIndex = _ArraySearch($clipboardData, $selectedRowDataID, 0, 0, 0, 0, 0, $clipboardDataIdColumn)
		If $arrayIndex <> -1 Then
			_ArrayDelete($clipboardData, $arrayIndex)
			_updatePreview(" ")
			_updateClipboardManagerGUI()
			Return
		EndIf
	EndIf
	TrayTip($AppTitle, "WARNING" & @CRLF & "Unable to delete.", 2, 2)
EndFunc   ;==>_deleteClipboardEntry

Func _selectTopRow()
	If _GUICtrlListView_GetItemCount($clipboardListView) > 0 Then
		_GUICtrlListView_ClickItem($clipboardListView, 0)
	EndIf
EndFunc   ;==>_selectTopRow

Func _clipboardManagerStarItem()
	$selectedRowDataID = _getSelectedDataIdFromListView()
	If $selectedRowDataID = -1 Then Return
	; Within the $clipboardData array, search if there is a row with the same dataID.
	$arrayIndex = _ArraySearch($clipboardData, $selectedRowDataID, 0, 0, 0, 0, 0, $clipboardDataIdColumn)
	; Star/unstar item
	If $arrayIndex <> -1 Then $clipboardData[$arrayIndex][2] = ($clipboardData[$arrayIndex][2] = "*" ? "" : "*")
	_updateClipboardManagerGUI()
EndFunc ;==>_clipboardManagerStarItem

Func _filterStarredData(ByRef $dataArray)
	If UBound($dataArray) = 0 Then Return
	If GUICtrlRead($filterStarredCheckbox) = $GUI_CHECKED Then
		$index = 0
		While $index < UBound($dataArray)
			if $dataArray[$index][2] = "" Then
				_ArrayDelete($dataArray, $index)
			Else
				$index += 1
			EndIf
		Wend
	EndIf
EndFunc ;==>_filterStarredData

; Search the clipboardmanager
Func _clipboardManagerStartSearch()
	_hideClipboardManager()	; Needs to hide or it will cover the inputbox.
	_clipboardManagerSearchForString(InputBox("Search My Clipboard Manager","What do you want to search?", $lastSearchString))
EndFunc ;==>_clipboardManagerStartSearch

Func _clipboardManagerSearchForString($searchString)
	_showClipboardManager()
	If $searchString = "" Then Return
	$lastSearchString = $searchString
	Local $listviewDataArray[0][4]
	Local $searchResult = _ArrayFindAll($clipboardData, $searchString, Default, Default, Default, 1, $clipboardDataColumn)
	If @error = 0 Then
		For $index In $searchResult
			_ArrayAdd($listviewDataArray, _ArrayExtract($clipboardData, $index, $index))
		Next
	EndIf
	_updateListView($listviewDataArray)
EndFunc ;==>_clipboardManagerSearchForString

; ***** Functions Anti Screen Lock
Func _antiScreenlock()
	If TimerDiff($antiScreenlockTimer) > $antiScreenlockPeriod And $antiScreenlockIsEnabled Then
		Send("{SCROLLLOCK}")
		Sleep(50)
		Send("{SCROLLLOCK}")
		$antiScreenlockTimer = TimerInit()
	EndIf
EndFunc

Func _toggleAntiScreenlockEnabled()
	$antiScreenlockIsEnabled = Not $antiScreenlockIsEnabled
	TrayTip($AppTitle, "Anti Screenlock"&@CRLF&"Active = " & $antiScreenlockIsEnabled, 2, 1)
EndFunc
