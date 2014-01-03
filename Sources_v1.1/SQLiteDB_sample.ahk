; ======================================================================================================================
; Script Function:  Sample script for Class_SQLiteDB.ahk
; AHK Version:      L 1.1.00.00 (U 32)
; Language:         English
; Tested on:        Win XPSP3, Win VistaSP2 (32 Bit)
; Author:           just me
; Version:          0.0.00.03/2011-05-24/just me
; ======================================================================================================================
; AHK Settings
; ======================================================================================================================
#NoEnv
; #Warn
#SingleInstance force
SetWorkingDir, %A_ScriptDir%
SetBatchLines, -1
OnExit, GuiClose
; ======================================================================================================================
; Includes
; ======================================================================================================================
#Include Class_SQLiteDB.ahk
; ======================================================================================================================
; Start & GUI
; ======================================================================================================================
CBBSQL := "SELECT * FROM Test"
DBFileName := A_ScriptDir . "\TEST.DB"
Title := "SQL Query/Command ListView Function GUI"
If FileExist(DBFileName) {
   SB_SetText("Deleting " . DBFileName)
   FileDelete, %DBFileName%
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
Gui, Add, Button, Section x+10 yp wp hp gGetTable, Get _Table
Gui, Add, Button, x+10 yp wp hp gGetRecordSet, Get _RecordSet
Gui, Add, GroupBox, xm w780 h330, Results
Gui, Add, ListView, xp+10 yp+18 w760 h300 vResultsLV +LV0x00010000
Gui, Add, StatusBar,
Gui, Show, , %Title%
; ======================================================================================================================
; Use Class SQLiteDB : Initialize and get lib version
; ======================================================================================================================
SB_SetText("SQLiteDB new")
DB := new SQLiteDB
Sleep, 1000
SB_SetText("Version")
Version := DB.Version
WinSetTitle, %Title% - SQLite3.dll v %Version%
Sleep, 1000
; ======================================================================================================================
; Use Class SQLiteDB : Open/Create database and table
; ======================================================================================================================
SB_SetText("OpenDB")
If !DB.OpenDB(DBFileName) {
   MsgBox, 16, SQLite Error, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
   ExitApp
}
Sleep, 1000
SB_SetText("Exec: CREATE TABLE")
SQL := "CREATE TABLE Test (Name, Fname, Phone, Room, PRIMARY KEY(Name ASC, FName ASC));"
If !DB.Exec(SQL)
   MsgBox, 16, SQLite Error, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
Sleep, 1000
SB_SetText("Exec: INSERT 1000 rows")
Start := A_TickCount
DB.Exec("BEGIN TRANSACTION;")
SQLStr := ""
_SQL := "INSERT INTO Test VALUES('Näme#', 'Fname#', 'Phone#', 'Room#');"
Loop, 1000 {
   StringReplace, SQL, _SQL, #, %A_Index%, All
   SQLStr .= SQL
}
If !DB.Exec(SQLStr)
   MsgBox, 16, SQLite Error, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
DB.Exec("COMMIT TRANSACTION;")
SQLStr := ""
SB_SetText("Exec: INSERT 1000 rows done in " . (A_TickCount - Start) . " ms")
Sleep, 1000
; ======================================================================================================================
; Use Class SQLiteDB : Using Exec() with callback function
; ======================================================================================================================
SB_SetText("Exec: Using a callback function")
SQL := "SELECT COUNT(*) FROM Test;"
DB.Exec(SQL, "SQLiteExecCallBack")
; ======================================================================================================================
; Use Class SQLiteDB : Get some informations
; ======================================================================================================================
SB_SetText("LastInsertRowID")
If !DB.LastInsertRowID(RowID)
   MsgBox, 16, SQLite Error, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
GuiControl, -ReDraw, ResultsLV
LV_Delete()
ColCount := LV_GetCount("Column")
Loop, %ColCount%
   LV_DeleteCol(1)
LV_InsertCol(1,"", "LastInsertedRowID")
LV_Add("", RowID)
GuiControl, +ReDraw, ResultsLV
Sleep, 1000
SQL := "SELECT COUNT(*) FROM Test;"
SB_SetText("SQLite_GetTable : " . SQL)
Result := ""
If !DB.GetTable(SQL, Result)
   MsgBox, 16, SQLite Error: GetTable, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
ShowTable(Result)
Sleep, 1000
; ======================================================================================================================
; Start of query using Query() : Get the column names for table Test
; ======================================================================================================================
SQL := "SELECT * FROM Test;"
SB_SetText("Query : " . SQL)
If !DB.Query(SQL, RecordSet)
   MsgBox, 16, SQLite Error: Query, % "Msg:`t" . RecordSet.ErrorMsg . "`nCode:`t" . RecordSet.ErrorCode
GuiControl, -ReDraw, ResultsLV
LV_Delete()
ColCount := LV_GetCount("Column")
Loop, %ColCount%
   LV_DeleteCol(1)
LV_InsertCol(1,"", "Column names")
Loop, % RecordSet.ColumnCount
   LV_Add("", RecordSet.ColumnNames[A_Index])
LV_ModifyCol(1, "AutoHdr")
RecordSet.Free()
GuiControl, +ReDraw, ResultsLV
; ======================================================================================================================
; End of query using Query()
; ======================================================================================================================
Gui, -Disabled
Return
; ======================================================================================================================
; Gui Subs
; ======================================================================================================================
GuiClose:
GuiEscape:
If !DB.CloseDB()
   MsgBox, 16, SQLite Error, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
Gui, Destroy
ExitApp
; ======================================================================================================================
; Other Subs
; ======================================================================================================================
; "One step" query using GetTable()
; ======================================================================================================================
GetTable:
Gui, Submit, NoHide
Result := ""
SQL := "SELECT * FROM " . Table . ";"
SB_SetText("GetTable: " . SQL)
Start := A_TickCount
If !DB.GetTable(SQL, Result)
   MsgBox, 16, SQLite Error, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
SB_SetText("GetTable: " . SQL . " done in " . (A_TickCount - Start) . " ms")
ShowTable(Result)
Return
; ======================================================================================================================
; Show results for prepared query using Query()
; ======================================================================================================================
GetRecordSet:
Gui, Submit, NoHide
SQL := "SELECT * FROM " . Table . ";"
SB_SetText("Query: " . SQL)
RecordSet := ""
Start := A_TickCount
If !DB.Query(SQL, RecordSet)
   MsgBox, 16, SQLite Error: Query, % "Msg:`t" . RecordSet.ErrorMsg . "`nCode:`t" . RecordSet.ErrorCode
ShowRecordSet(RecordSet)
RecordSet.Free()
SB_SetText("Query: " . SQL . " done in " . (A_TickCount - Start) . " ms")
Return
; ======================================================================================================================
; Execute SQL statement using Exec() / GetTable()
; ======================================================================================================================
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
If (SubStr(SQL, 0) <> ";")
   SQL .= ";"
Result := ""
If RegExMatch(SQL, "i)^\s*SELECT\s") {
   SB_SetText("GetTable: " . SQL)
   If !DB.GetTable(SQL, Result)
      MsgBox, 16, SQLite Error: GetTable, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
   Else
      ShowTable(Result)
   SB_SetText("GetTable: " . SQL . " done!")
} Else {
   SB_SetText("Exec: " . SQL)
   If !DB.Exec(SQL)
      MsgBox, 16, SQLite Error: Exec, % "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode
   Else
      SB_SetText("Exec: " . SQL . " done!")
}
Return
; ======================================================================================================================
; Exec() callback function sample
; ======================================================================================================================
SQLiteExecCallBack(DB, ColumnCount, ColumnValues, ColumnNames) {
   This := Object(DB)
   MsgBox, 0, %A_ThisFunc%
      , % "SQLite version: " . This.Version . "`n"
      . "SQL statement: " . StrGet(A_EventInfo) . "`n"
      . "Number of columns: " . ColumnCount . "`n" 
      . "Name of first column: " . StrGet(NumGet(ColumnNames + 0, "UInt"), "UTF-8") . "`n" 
      . "Value of first column: " . StrGet(NumGet(ColumnValues + 0, "UInt"), "UTF-8")
   Return 0
}
; ======================================================================================================================
; Show results
; ======================================================================================================================
ShowTable(Table) {
   Global
   Local ColCount, RowCount, Row
   GuiControl, -ReDraw, ResultsLV
   LV_Delete()
   ColCount := LV_GetCount("Column")
   Loop, %ColCount%
      LV_DeleteCol(1)
   If (Table.HasNames) {
      Loop, % Table.ColumnCount
         LV_InsertCol(A_Index,"", Table.ColumnNames[A_Index])
      If (Table.HasRows) {
         Loop, % Table.RowCount {
            RowCount := LV_Add("", "")
            Table.Next(Row)
            Loop, % Table.ColumnCount
               LV_Modify(RowCount, "Col" . A_Index, Row[A_Index])
         }
      }
      Loop, % Table.ColumnCount
         LV_ModifyCol(A_Index, "AutoHdr")
   }
   GuiControl, +ReDraw, ResultsLV
}
; ----------------------------------------------------------------------------------------------------------------------
ShowRecordSet(RecordSet) {
   Global
   Local ColCount, RowCount, Row, RC
   GuiControl, -ReDraw, ResultsLV
   LV_Delete()
   ColCount := LV_GetCount("Column")
   Loop, %ColCount%
      LV_DeleteCol(1)
   If (RecordSet.HasNames) {
      Loop, % RecordSet.ColumnCount
         LV_InsertCol(A_Index,"", RecordSet.ColumnNames[A_Index])
   }
   If (RecordSet.HasRows) {
      If (RecordSet.Next(Row) < 1) {
         MsgBox, 16, %A_ThisFunc%, % "Msg:`t" . RecordSet.ErrorMsg . "`nCode:`t" . RecordSet.ErrorCode
         Return
      }
      Loop {
         RowCount := LV_Add("", "")
         Loop, % RecordSet.ColumnCount
            LV_Modify(RowCount, "Col" . A_Index, Row[A_Index])
            RC := RecordSet.Next(Row)
      } Until (RC < 1)
   }
   If (RC = 0)
      MsgBox, 16, %A_ThisFunc%, % "Msg:`t" . RecordSet.ErrorMsg . "`nCode:`t" . RecordSet.ErrorCode
   Loop, % RecordSet.ColumnCount
      LV_ModifyCol(A_Index, "AutoHdr")
   GuiControl, +ReDraw, ResultsLV
}