option explicit

sub selectCalculatedValues() ' {

   dim db as dao.database
   set db =  application.currentDB

 '
 ' Drop table if it already exists:
 '
   if not isNull(dLookup("Name", "MSysObjects", "Name='tq84_tab'")) then db.execute("drop table tq84_tab")


 '
 ' Create test table …
 '
   db.execute("create table tq84_tab(a number, b number, c number, d number)")

 '
 ' … and fill some values in it:
 '
   db.execute("insert into tq84_tab values (1, 1, 1, 1)")
   db.execute("insert into tq84_tab values (1, 1, 2, 4)")
   db.execute("insert into tq84_tab values (3, 1, 0, 1)")

   debug.print(" a b c d | a+b  a+b+c  a+b+c+d")
   debug.print("---------+--------------------")

   dim rs as dao.recordSet
   set rs = db.openRecordset("select                  " & _
                             "  a, b, c, d ,          " & _
                             "  a+b        as A_B,    " & _
                             "  A_B   + c  as A_B_C,  " & _
                             "  A_B_C + d  as A_B_C_D " & _
                             "from                    " & _
                             "  tq84_tab")
   do while not rs.eof ' {
      debug.print (" " & rs!a & " " & rs!b & " " & rs!c & " " & rs!d & " |   " & rs!A_B & "      " & rs!A_B_C & "        " & rs!A_B_C_D)
      rs.moveNext
   loop

end sub ' }
