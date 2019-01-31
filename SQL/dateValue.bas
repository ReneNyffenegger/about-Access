option explicit

sub dateValueTest() ' {

   dim db as dao.database
   set db =  application.currentDB

 '
 ' Drop table if it already exists:
 '
   if not isNull(dLookup("Name", "MSysObjects", "Name='dateValueTest'")) then db.execute("drop table dateValueTest")


 '
 ' Create test table …
 '
   db.execute("create table dateValueTest(d date, t varchar(19))")

 '
 ' … and fill some values in it:
 '
   db.execute("insert into dateValueTest values (#2012-12-12 12:12:12#, '2012-12-12 12:12:12')")
   db.execute("insert into dateValueTest values (#2001-01-01 00:00:00#, '2001-01-01 00:00:00')")
   db.execute("insert into dateValueTest values (#2023-12-31 23:59:59#, '2023-12-31 23-59:59')")


 '
 ' Use dateValue(…) to get only the date without time.
 '
   dim rs as dao.recordSet
   set rs = db.openRecordset("select dateValue(d) as dt, t from dateValueTest")
   do while not rs.eof ' {
      debug.print (rs!dt & " | " & rs!t)
      rs.moveNext
   loop

end sub ' }
