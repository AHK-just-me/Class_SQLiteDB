; ======================================================================================================================
; Script Function:  BLOB sample script for Class_SQLiteDB.ahk
; AHK Version:      v2 alpha
; Tested on:        Win 7 Pro (x64)
; Author:           just me
; ======================================================================================================================
; AHK Settings
; ======================================================================================================================
#SingleInstance Force
#Warn
#Warn LocalSameAsGlobal, Off
SetWorkingDir(A_ScriptDir)
; ======================================================================================================================
; Includes
#Include Class_SQLiteDB.ahk
; ======================================================================================================================
; Get the Google logo or store a picture named Original.gif in the script's folder and comment this out
Try FileDelete("Original.gif")
Download("http://www.google.de/intl/de_ALL/images/logos/images_logo_lg.gif", "Original.gif")
; ======================================================================================================================
; Start
Try FileDelete("Blob.gif")
DBFileName := A_ScriptDir . "\TestBLOB.DB"
If FileExist(DBFileName) {
   SB_SetText("Deleting " . DBFileName)
   FileDelete(DBFileName)
}
; ======================================================================================================================
; Use Class SQLiteDB : Create new instance
DB := New SQLiteDB
; ======================================================================================================================
; Use Class SQLiteDB : Open/create database and table, insert a BLOB from a GIF file
If !DB.OpenDB(DBFileName) {
   MsgBox(16, "SQLite Error", "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode)
   ExitApp
}
HFILE := FileOpen("Original.gif", "r")
Size := HFILE.RawRead(BLOB, HFILE.Length)
HFILE.Close()
If !DB.Exec("CREATE TABLE Test (TextType, BlobType);")
   MsgBox(16, "SQLite Error", "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode)
DB.Exec("BEGIN TRANSACTION;")
; ? stands for an automatically numbered parameter (here: 1) to use in BlobArray
SQL := "INSERT INTO Test VALUES('Text', ?);"
; Create the BLOB array
BlobArray := [{Addr: &BLOB, Size: Size}]
If !DB.StoreBLOB(SQL, BlobArray)
   MsgBox(16, "SQLite Error", "Msg:`t" . DB.ErrorMsg . "`nCode:`t" . DB.ErrorCode)
DB.Exec("COMMIT TRANSACTION;")
; ======================================================================================================================
; Start of query using Query() : Get the BLOB from table Test
HFILE := FileOpen("Blob.gif", "w")
If !DB.Query("SELECT * FROM Test;", RecordSet)
   MsgBox(16, "SQLite Error: Query", "Msg:`t" . RecordSet.ErrorMsg . "`nCode:`t" . RecordSet.ErrorCode)
If (RecordSet.HasRows) {
   If (RecordSet.Next(Row) < 1) {
      MsgBox(16, A_ThisFunc, "Msg:`t" . RecordSet.ErrorMsg . "`nCode:`t" . RecordSet.ErrorCode)
   } Else {
      Loop, % RecordSet.ColumnCount {
         If IsObject(Row[A_Index]) {
            Size := Row[A_Index].Size
            Addr := Row[A_Index].GetAddress("Blob")
            If !(Addr) || !(Size) {
               MsgBox(0, "Error", "BlobAddr = %Addr% - BlobSize = %Size%")
            } Else {
               VarSetCapacity(MyBLOBVar, Size)
               DllCall("Kernel32.dll\RtlMoveMemory", "Ptr", &MyBLOBVar, "Ptr", Addr, "Ptr", Size)
               HFILE.RawWrite(&MyBLOBVar, Size)
            }
         }
      }
   }
}
RecordSet.Free()
HFILE.Close()
; ======================================================================================================================
; Show the result
Gui("Margin", 20, 20)
Gui("Add", "Text", "Section", "Original from Google:")
Gui("Add", "Pic", "y+5", "Original.gif")
Gui("Add", "Text", "ys", "Copy out of database:")
Gui("Add", "Pic", "y+5", "Blob.gif")
Gui("Show", , "Pictures")
Try FileDelete("Original.gif")
Try FileDelete("Blob.gif")
Return
GuiClose:
GuiEscape:
ExitApp