option explicit


type GUID ' {
  '
  '  Declared in rpcdce.h / included by rpc.h
  '
     Data1          as long
     Data2          as integer
     Data3          as integer
     Data4 (0 to 7) as byte
end  type ' }

declare function CoCreateGuid    lib "ole32" (pguid as GUID) as long
declare function StringFromGUID2 lib "ole32" (rguid as GUID, byVal lpOleChar as any, byVal cbmax as long) as long

function CoCreateGuid_ as GUID ' {

    if CoCreateGuid(CoCreateGuid_) <> 0 then
       MsgBox "Something went wrong with CoCreateGuid"
    end if

end function ' }

function StringFromGUID2_(rguid as GUID) as string ' {

    StringFromGUID2_ = space$(38)

    call StringFromGUID2 (rguid, strPtr(StringFromGUID2_), 38*2)

end function ' }

sub main() ' {

    dim db as dao.database
    set db = application.currentDB

    cleanUpLastRun db
    createTables   db
    insertValues   db
    selectValues   db

end sub ' }

sub cleanUpLastRun(db as dao.database) ' {

    dropTableIfExists db, "tq84_child" 
    dropTableIfExists db, "tq84_parent"

end sub ' }

sub createTables(db as dao.database) ' {

    db.execute(   _
      "create table tq84_parent ("                            & _
      "  id     guid primary key,"                            & _
      "  txt    char(10)        "                             & _
      ")")

    db.execute(   _
      "create table tq84_child ("                             & _
      "  id_parent guid       null references tq84_parent,"   & _
      "  txt    char(10)        "                             & _
      ")")

end sub ' }

sub insertValues(db as dao.database) ' {

'    dim stmtParent, stmtChild   as dao.queryDef
     dim stmtParent              as dao.queryDef
     set stmtParent = db.createQueryDef("",   _
       "parameters "         & _
       "  id  char(38), "    & _
       "  txt char(10); "    & _
       "insert into tq84_parent(id, txt) values ([id], [txt]) ")

     dim stmtChild   as dao.queryDef
     set stmtChild = db.createQueryDef("",   _
       "parameters "              & _
       "  id_parent  char(38),"   & _
       "  txt char(10)       ;"   & _
       "insert into tq84_child(id_parent, txt) values ([id_parent], [txt]) ")

'    dim guid_1, guid_2, guid_3, guid_4 as guid
     dim guid_1 as guid
     dim guid_2 as guid
     dim guid_3 as guid
     dim guid_4 as guid

     guid_1 = CoCreateGuid_
     guid_2 = CoCreateGuid_
     guid_3 = CoCreateGuid_
     guid_4 = CoCreateGuid_

     call insertValuesParent(stmtParent, guid_1, "one"    )
     call insertValuesParent(stmtParent, guid_2, "two"    )
     call insertValuesParent(stmtParent, guid_3, "three"  )

     call insertValuesChild (stmtChild , guid_1, "uno"    )
     call insertValuesChild (stmtChild , guid_1, "eins"   )
     call insertValuesChild (stmtChild , guid_3, "tre"    )
     call insertValuesChild (stmtChild , guid_4, "quattro") ' Note missing parent!

end sub ' }

sub insertValuesParent(stmt as dao.queryDef, id as guid, txt as string) ' {

     stmt.parameters!id  = StringFromGUID2_(id)
     stmt.parameters!txt = txt
     stmt.execute

end sub ' }

sub insertValuesChild(stmt as dao.queryDef, id_parent as guid, txt as string) ' {

  on error goto err_

     stmt.parameters!id_parent = StringFromGUID2_(id_parent)
     stmt.parameters!txt       = txt

   '
   ' Without dbFailOnError, the following stmt executes without
   ' throwing an error if id_parent does not refer to a record
   ' in tq84_parent - but the record is (obviously) not inserted!
   '
   ' Therefore, execute should, imho, always be used with
   ' dbFailOnError
   '
     stmt.execute dbFailOnError
     exit sub

  err_:
    debug.print "error in insertValuesChild(): " & err.description
    debug.print "  txt = " & txt

end sub ' }

sub selectValues(db as dao.database) ' {

    dim stmt as queryDef
    dim rs as dao.recordSet

    set stmt = db.createQueryDef("", _
      "select "                     & _
      "  p.txt as parent_txt, "     & _
      "  c.txt as child_txt   "     & _
      "from "                       & _
      "  tq84_parent p left join " & _
      "  tq84_child  c on p.id = c.id_parent")
    set rs = stmt.openRecordSet

    debug.print "Join parent - child"
    do while not rs.eof
       debug.print("  " & rs!parent_txt & ":  " & rs!child_txt)
       rs.moveNext
    loop

  ' ------------------------------------------

    debug.print "child"
    set stmt = db.createQueryDef("", _
      "select "                     & _
      "  c.txt as child_txt   "     & _
      "from "                       & _
      "  tq84_child  c")
    set rs = stmt.openRecordSet

    do while not rs.eof
       debug.print("  " & rs!child_txt)
       rs.moveNext
    loop

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
