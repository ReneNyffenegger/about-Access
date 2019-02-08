option explicit

sub main() ' {

    if not isNull(dLookup("Name", "MSysObjects", "Name='tq84_vw'")) then ' {
       doCmd.close acQuery, "tq84_vw", acSaveNo
       execSQL "drop view tq84_vw"
    end if ' }

    if not isNull(dLookup("Name", "MSysObjects", "Name='tq84_tab'")) then ' {
       doCmd.close acTable, "tq84_tab", acSaveNo
       execSQL "drop table tq84_tab"
    end if ' }

    execSQL "create table tq84_tab(num number)"

    execSQL "insert into tq84_tab values ( 1 )"
    execSQL "insert into tq84_tab values ( 2 )"
    execSQL "insert into tq84_tab values ( 3 )"

    execSQL "create view tq84_vw as select num, giveMeADate(num) as dat from tq84_tab"

    dim rs as dao.recordSet
    set rs = currentDB.openRecordset(   _
      "select                         " & _
      "  num,                         " & _
      "  dat                          " & _
      "from                           " & _
      "  tq84_vw                      " & _
      "order by                       " & _
      "  cDate(nz(dat, #9999-12-31#)) " )

    do while not rs.eof ' {
       debug.print(rs!num & "  " & rs!dat)
       rs.moveNext
    loop ' }

    set rs = nothing

end sub ' }


function giveMeADate(i as long) as variant ' {

    select case i
           case 1 : giveMeADate = cDate(#2011-01-30 01:02:03#)
           case 2 : giveMeADate = null
           case 3 : giveMeADate = cDate(#2012-12-01 18:17:16#)
    end select

end function ' }


sub execSQL(stmt as string) ' {
    currentProject.connection.execute(stmt)
end sub ' }
