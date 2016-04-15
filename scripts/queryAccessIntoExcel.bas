' PATH=c:\lib\runVBAFilesInOffice;%PATH%
' set VBAMODDIR= y:\2016-03-Entwicklung-René\prototypes\01\VBAModules\
'
' runVBAFilesInOffice -excel %CD%\queryAccessIntoExcel %VBAMODDIR%common\File %VBAMODDIR%Database\SQL %VBAMODDIR%Database\ADOHelpers -c main c:\temp\sql.sql
'
' Runs the select statement in the given file and writes
' the returned values into an excel sheet.
'

sub main(accDB as string, fileWithSelectStatement as string) ' {

  dim sqlText as string
  sqlText = slurpFile(fileWithSelectStatement)
  sqlText = removeSQLComments(sqlText)

  dim con as ADODB.connection
  set con = openADOConnectionToAccess(accDB)

  dim rs as ADODB.recordSet
  set rs = con.execute(sqlText)

  range("A1").copyFromRecordset rs

  activeWorkbook.saved = true

end sub ' }
