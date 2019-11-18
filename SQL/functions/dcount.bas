option explicit

sub dcount_test() ' {

    dim db as dao.database
    set db = application.currentDB

    dropTableIfExists db, "tq84_dcount_test"

    db.execute "create table tq84_dcount_test(col_1 number, col_2 number, col_3 number)"

    db.execute "insert into tq84_dcount_test (col_1, col_2, col_3) values (1, 10, null)"
    db.execute "insert into tq84_dcount_test (col_1, col_2, col_3) values (2, 20,  101)"
    db.execute "insert into tq84_dcount_test (col_1, col_2, col_3) values (3, 10,  202)"
    db.execute "insert into tq84_dcount_test (col_1, col_2, col_3) values (4, 10,  303)"

    dim rs as dao.recordSet
    set rs = db.createQueryDef("",                                        _ 
       "select "                                                        & _
       "    count(  *                            ) as cnt           , " & _ 
       "    count(  col_1                        ) as cnt_col_1     , " & _
       "    count(  col_2                        ) as cnt_col_2     , " & _
       "    count(  col_3                        ) as cnt_col_3     , " & _
       "   dcount(""col_1"", ""tq84_dcount_test"") as cnt_dist_col_1, " & _
       "   dcount(""col_2"", ""tq84_dcount_test"") as cnt_dist_col_2, " & _
       "   dcount(""col_3"", ""tq84_dcount_test"") as cnt_dist_col_3  " & _
       "from "                                                          & _
       "   tq84_dcount_test").openRecordSet


    dim sep as string
    sep = chr(9) & "| "

    do while not rs.eof ' {
       debug.print(" count( *                          ) : " & rs!cnt      )
       debug.print(" count( col_1                      ) : " & rs!cnt_col_1)
       debug.print(" count( col_2                      ) : " & rs!cnt_col_2)
       debug.print(" count( col_3                      ) : " & rs!cnt_col_3)
       debug.print("dcount(""col_1"", ""tq84_dcount_test"")) : " & rs!cnt_dist_col_1)
       debug.print("dcount(""col_2"", ""tq84_dcount_test"")) : " & rs!cnt_dist_col_2)
       debug.print("dcount(""col_3"", ""tq84_dcount_test"")) : " & rs!cnt_dist_col_3)
'      debug.print(rs!cnt_col_1      & sep & _
'                  rs!cnt_col_2      & sep & _
'                  rs!cnt_col_3      & sep & _
'                  rs!cnt_dist_col_1 & sep & _
'                  rs!cnt_dist_col_2 & sep & _
'                  rs!cnt_dist_col_3 & sep)

       rs.moveNext
    loop ' }

end sub ' }

sub dropTableIfExists(db as dao.database, tableName as string) ' {
  on error goto err_
    db.execute("drop table " & tableName)
    exit sub
  err_:
    if err.number = 3376 then
     '
     ' Ignore »Table … does not exist«.
     '
       exit sub
    end if

    err.raise err.number, err.source, err.description

end sub ' }
'
' Output:
'
'   count( *                          ) : 4
'   count( col_1                      ) : 4
'   count( col_2                      ) : 4
'   count( col_3                      ) : 3
'  dcount("col_1", "tq84_dcount_test")) : 4
'  dcount("col_2", "tq84_dcount_test")) : 4
'  dcount("col_3", "tq84_dcount_test")) : 3
