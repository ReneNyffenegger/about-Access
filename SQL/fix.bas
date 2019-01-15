option explicit

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

    dropTableIfExists db, "tq84_fix_test"

    db.execute "create table tq84_fix_test(num number, dt date)"
    db.execute "insert into tq84_fix_test (num, dt) values (14.05, #2018-10-03 09:55:18#)"
    db.execute "insert into tq84_fix_test (num, dt) values (24.69, #2018-05-18 17:12:33#)"

    dim rs as dao.recordSet
    set rs = db.createQueryDef("", "select num, dt, fix(num) as num_, fix(dt) as dt_ from tq84_fix_test").openRecordSet

    dim sep as string
    sep = chr(9) & "| "

    do while not rs.eof ' {
       debug.print(rs!num & sep & rs!dt & sep & rs!num_ & sep & rs!dt_)
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
