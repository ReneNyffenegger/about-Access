option explicit

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

    cleanUpLastRun   db
    createTables     db
    insertValues     db

 '  ------------------------------------------------------
    db.execute                                           _
      "   update                                     " & _
      "     dest inner join                          " & _
      "     src             on src.id_dest = dest.id " & _
      "   set                                        " & _
      "     dest.txt = src.txt, "                      & _
      "     dest.num = src.num  "  ,                     _
          dbFailOnError
 '  ------------------------------------------------------

    selectValues     db

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

sub cleanUpLastRun(db as dao.database) ' {

    call dropTableIfExists(db, "dest")
    call dropTableIfExists(db, "src" )

end sub ' }

sub createTables(db as dao.database) ' {

    db.execute "create table dest(id      number, txt varchar(10), num long)", dbFailOnError
    db.execute "create table src (id_dest number, txt varchar(10), num long)", dbFailOnError

end sub ' }

sub insertValues(db as dao.database) ' {

    db.execute "insert into dest(id     , txt, num) values (1, 'foo'   , 111)"   , dbFailOnError
    db.execute "insert into dest(id     , txt, num) values (2, 'bar'   , 222)"   , dbFailOnError
    db.execute "insert into dest(id     , txt, num) values (3, 'baz'   , 333)"   , dbFailOnError

    db.execute "insert into src (id_dest, txt, num) values (1, 'aaa',   1)", dbFailOnError
    db.execute "insert into src (id_dest, txt, num) values (3, 'ccc',   3)", dbFailOnError  ' Will it be updated with
    db.execute "insert into src (id_dest, txt, num) values (3, 'CCC',  -3)", dbFailOnError  ' ccc or CCC, 3 or -3?
    db.execute "insert into src (id_dest, txt, num) values (4, 'ddd',   4)", dbFailOnError


end sub ' }

sub selectValues(db as dao.database) ' {

    dim stmt as queryDef
    set stmt = db.createQueryDef("", "select * from dest")

    dim rs as dao.recordSet
    set rs = stmt.openRecordSet

    do while not rs.eof ' {
       debug.print(rs!id & ": " & rs!txt & ", " & rs!num)
       rs.moveNext
    loop ' }

end sub ' }
