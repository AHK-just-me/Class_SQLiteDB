;=======================================================================================================================
; Script Function:  Sample script for SQLite.ahk
; AHK Version:      L 1.0.97.02 (U 32)
; Language:         English
; Tested on:        Win XPSP3, Win VistaSP2 (32 Bit)
; Author:           ich_L
; Version:          0.0.00.01/2011-05-01/ich_L
; Remarks:          As suggested by andricOn on
;                   http://www.autohotkey.com/forum/post-329350.html#329350
;=======================================================================================================================
; AHK Settings
;=======================================================================================================================
#NoEnv
#SingleInstance force
SetWorkingDir, %A_ScriptDir%
SetBatchLines, -1
;=======================================================================================================================
; Includes
;=======================================================================================================================
#Include *i SQLite_L.ahk
;=======================================================================================================================
; Start
;=======================================================================================================================
CBBSQL := "SELECT * FROM Test"
DBFileName := A_ScriptDir . "\TEST.DB"
CSVFileName := A_ScriptDir . "\TEST.CSV"
Title := "SQL Query/Command ListView Function GUI"
If FileExist(DBFileName) {
   SB_SetText("Deleting " . DBFileName)
   FileDelete, %DBFileName%
}
If FileExist(CSVFileName) {
   SB_SetText("Deleting " . CSVFileName)
   FileDelete, %CSVFileName%
}
Loop, 500 {
   FileAppend,
   (LTrim Join
      Näme%A_Index%%A_Tab%Fname%A_Index%%A_Tab%
      Phöne%A_Index%%A_Tab%Room%A_Index%`n
   ),%CSVFileName%, UTF-8-RAW
}
Gui, +LastFound +OwnDialogs +Disabled
Gui, Margin, 10, 10
Gui, Add, Text, w100 h20 0x200 vTX, SQL statement:
Gui, Add, ComboBox, x+0 ym w590 vSQL Sort, %CBBSQL%
GuiControlGet, P, Pos, SQL
GuiControl, Move, TX, h%PH%
Gui, Add, Button, ym w80 hp vRun gRunSQL Default, Run
Gui, Add, Text, xm h20 w100 0x200, Table name:
Gui, Add, Edit, x+0 yp w150 hp vTable, Test
Gui, Add, Button, Section x+10 yp wp hp gGetTable, SQLite_GetTable
Gui, Add, Button, x+10 yp wp hp gFetchData, SQLite_FetchData
Gui, Add, GroupBox, xm w780 h330 , Results
Gui, Add, ListView, xp+10 yp+18 w760 h300 vResultsLV,
Gui, Add, StatusBar,
Gui, Show, , %Title%
;=======================================================================================================================
; Use SQLITE3.DLL - Initialize and get lib version
;=======================================================================================================================
SB_SetText("SQLite_StartUp")
If !SQLite_Startup() {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0,  ERROR from STARTUP, %Msg%
   ExitApp
}
Sleep, 1000
SB_SetText("SQLite_LibVersion")
Version := SQLite_LibVersion()
WinSetTitle, %Title% - SQLite3.dll v %Version%
Sleep, 1000
;=======================================================================================================================
; Use SQLITE3.EXE - Create database and table
;=======================================================================================================================
; Commands =
; (Ltrim
;    CREATE TABLE Test (Name, Fname, Phone, Room, PRIMARY KEY(Name ASC, FName ASC));
;    .separator \t
;    .import '%CSVFileName%' Test
; )
; Output := ""
; SB_SetText("SQLite3exe: Creating database and table test importing 500 records")
; Sleep, 1000
; Start := A_TickCount
; If !SQLite_SQLiteExe(DBFileName, Commands, Output) {
;    Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
;    MsgBox, 0, ERROR from SQLITEEXE, %Msg%`n`n%Output%
;    ExitApp
; }
; SB_SetText("SQLite3exe: Done in " . (A_TickCount - Start) . " ms")
; Sleep, 1000
;=======================================================================================================================
; Use SQLITE3.DLL - Open/Create database and table
;=======================================================================================================================
If !(hDB := SQLite_OpenDB(DBFileName)) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from OPENDB, %Msg%
}
SQL := "CREATE TABLE Test (Name, Fname, Phone, Room, PRIMARY KEY(Name ASC, FName ASC));"
SB_SetText("SQLite_Exec: CREATE TABLE")
If !SQLite_Exec(hDB, SQL) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from EXEC, %Msg%
}
SB_SetText("SQLite_Exec: INSERT 1000 rows")
Start := A_TickCount
SQL := "BEGIN TRANSACTION;"
SQLite_Exec(hDB, SQL)
_SQL := "INSERT INTO Test VALUES('Name#', 'Fname#', 'Phone#', 'Room#');"
I := 501
Loop, 1000 {
   StringReplace, SQL, _SQL, #, %I%, All
   If !SQLite_Exec(hDB, SQL) {
      Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
      MsgBox, 0, ERROR from EXEC, %Msg%
   }
   I++
}
SQL := "COMMIT TRANSACTION;"
SQLite_Exec(hDB, SQL)
SB_SetText("INSERT 1000 rows: Done in " . (A_TickCount - Start) . " ms")
Sleep, 1000
;=======================================================================================================================
; Use SQLITE3.DLL - Query the Database
;=======================================================================================================================
SB_SetText("SQLite_LastInsertRowID")
Names := Result := ""
If !SQLite_LastInsertRowID(hDB, RowID) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from LASTINSERTROWID, %Msg%
}
Names := ["Last inserted RowID"]
Result := [RowID]
GoSub, ShowResult
Sleep, 1000
SQL := "SELECT COUNT(*) FROM Test;"
Names := Result := ""
Rows := Cols := 0
SB_SetText("SQLite_GetTable : " . SQL)
If !SQLite_GetTable(hDB, SQL, Rows, Cols, Names, Result) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from GETTABLE, %Msg%
}
GoSub, ShowResult
Sleep, 1000
;=======================================================================================================================
; Start of query using SQLite_Query()
;=======================================================================================================================
SQL := "SELECT * FROM Test;"
SB_SetText("SQLite_Query : " . SQL)
If !(hQuery := SQLite_Query(hDB, SQL)) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from QUERY, %Msg%
}
Names := Result := ""
SB_SetText("SQLite_FetchNames")
If !SQLite_FetchNames(hQuery, Names) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from FETCHNAMES, %Msg%
}
Result := Names.Clone()
Names := ""
Names := ["Column Names"]
GoSub, ShowResult
Sleep, 1000
Names := Result
Result := ""
Result := Array()
Rows := 0
Loop {
   SB_SetText("SQLite_FetchData " . A_Index)
   If !(RC := SQLite_FetchData(hQuery, Data)) {
      Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
      MsgBox, 0, ERROR from FETCHDATA, %Msg%
      Break
   }
   If (RC = -1)
      Break
   Result[A_Index] := Data
   Rows++
}
Gosub, ShowResult
SB_SetText("SQLite_QueryFinalize")
If !SQLite_QueryFinalize(hQuery) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from FINALIZE, %Msg%
}
;=======================================================================================================================
; End of query using SQLite_Query()
;=======================================================================================================================
Gui, -Disabled
Return
;=======================================================================================================================
; Gui Subs
;=======================================================================================================================
GuiClose:
GuiEscape:
SB_SetText("SQLite_CloseDB")
If !SQLite_CloseDB(hDB) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from CLOSEDB, %$Msg%
}
SB_SetText("SQLite_ShutDown")
If !SQLite_ShutDown() {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from SHUTDOWN, %Msg%
}
Gui, Destroy
ExitApp
;=======================================================================================================================
; Other Subs
;=======================================================================================================================
;=======================================================================================================================
; "One step" query using SQLite_GetTableClass()
;=======================================================================================================================
GetTable:
Gui, Submit, NoHide
Result := ""
SQL := "SELECT * FROM " . Table . ";"
SB_SetText("SQLite_GetTable : SELECT * FROM " . Table)
Start := A_TickCount
If !SQLite_GetTableClass(hDB, SQL, Result) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from GETTABLE, %Msg%
}
SB_SetText("SQLite_GetTable done in " . (A_TickCount - Start) . " ms")
GuiControl, -ReDraw, ResultsLV
LV_Delete()
ColCount := LV_GetCount("Column")
Loop, %ColCount%
   LV_DeleteCol(1)
If (Result.HasNames) {
   Loop, % Result.ColumnCount
      LV_InsertCol(A_Index,"", Result.ColumnNames[A_Index])
   If (Result.HasRows) {
      Loop, % Result.RowCount {
         RowCount := LV_Add("", "")
         Row := Result.NextRow()
         Loop, % Result.ColumnCount
            LV_Modify(RowCount, "Col" . A_Index, Row[A_Index])
      }
   }
   Loop, % Result.ColumnCount
      LV_ModifyCol(A_Index, "AutoHdr")
}
GuiControl, +ReDraw, ResultsLV
Return
;=======================================================================================================================
; Show results for prepared query using SQLite_FetchData
;=======================================================================================================================
FetchData:
Gui, Submit, NoHide
SQL := "SELECT * FROM " . Table . ";"
SB_SetText("SQLite_Query : " . SQL)
Start := A_TickCount
If !(hQuery := SQLite_Query(hDB, SQL)) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from QUERY, %Msg%
}
Names := Result := Data := ""
SB_SetText("SQLite_FetchNames")
If !SQLite_FetchNames(hQuery, Names) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from FETCHNAMES, %Msg%
}
SB_SetText("SQLite_FetchData")
Result := Array()
Loop, {
   If !(RC := SQLite_FetchData(hQuery, Data)) {
      Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
      MsgBox, 0, ERROR from FETCHDATA, %Msg%
      Break
   }
   If (RC = -1)
      Break
   Result[A_Index] := Data
}
SB_SetText("SQLite_QueryFinalize")
If !SQLite_QueryFinalize(hQuery) {
   Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
   MsgBox, 0, ERROR from FINALIZE, %Msg%
}
SB_SetText("SQLite_FetchData done in " . (A_TickCount - Start) . " ms")
GoSub, ShowResult
Return
;=======================================================================================================================
; Execute SQL-Statement
;=======================================================================================================================
RunSQL:
Gui, +OwnDialogs
GuiControlGet, SQL
If SQL Is Space
{
   SB_SetText("No text entered")
   Return
}
If !InStr("`n" . CBBSQL . "`n", "`n" . SQL . "`n") {
   GuiControl, , SQL, %SQL%
   CBBSQL .= "`n" . SQL
}
SQL := SQL . ";"
Names := Result := ""
Rows := Cols := 0
If RegExMatch(SQL, "i)^\s*SELECT\s") {
   SB_SetText("SQLite_GetTable : " . SQL)
   If !SQLite_GetTable(hDB, SQL, Rows, Cols, Names, Result) {
      Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
      MsgBox, 0, ERROR from GETTABLE, %Msg%
   }
} Else {
   SB_SetText("SQLite_Exec: " . SQL)
   If !SQLite_Exec(hDB, SQL) {
      Msg := "ErrorLevel: " . ErrorLevel . "`n" . SQLite_LastError()
      MsgBox, 0, ERROR from EXEC, %Msg%
   }
}
GoSub, ShowResult
Return
;=======================================================================================================================
; Show result
;=======================================================================================================================
ShowResult:
GuiControl, -ReDraw, ResultsLV
LV_Delete()
ColCount := LV_GetCount("Column")
Loop, %ColCount%
   LV_DeleteCol(1)
For I In Names
   LV_InsertCol(A_Index,"", Names[I])
ColCount := I
For I In Result {
   RowCount := LV_Add("", "")
   If IsObject(Result[I]) {
      For J In Result[I]
         LV_Modify(RowCount, "Col" . J, Result[I][J])
   } Else {
      LV_Modify(RowCount, "Col1", Result[I])
   }
}
Loop, %ColCount%
   LV_ModifyCol(A_Index, "AutoHdr")
GuiControl, +ReDraw, ResultsLV
Return