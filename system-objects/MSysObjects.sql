select
   obj.Name,
   obj.Id,
-- obj.ParentId,
   par.Name        as parentName,
   par.Type        as parentType,
   par.Flags       as parentFlags,
   --
   switch( -- http://www.mendipdatasystems.co.uk/purpose-of-system-tables-2/4594446647
      obj.type = 1, 'Local table'       ,
      obj.type = 4, 'Linked ODBC table' ,
      obj.type = 6, 'Other linked table',
          1    = 1, '?'
   )                                     as typeDecoded,
   ---------------------------------------------------------------------------------------------------------------- TableDefAttributeEnum 
   iif( int(obj.Flags/(2^ 0)) mod 2 <> 0, 'x', '') as isDeeplyHidden, -- Bit  0              : Deeply hidden table (dbHiddenObject   )
   iif( int(obj.Flags/(2^ 1)) mod 2 <> 0, 'x', '') as isSystemTable , -- Bit  1              : System table
   iif( int(obj.Flags/(2^ 3)) mod 2 <> 0, 'x', '') as isHidden      , -- Bit  3              : Hidden
   iif( int(obj.Flags/(2^16)) mod 2 <> 0, 'x', '') as isExclusive   , -- Bit 16 (      65536):                     (dbAttachExclusive)
   iif( int(obj.Flags/(2^17)) mod 2 <> 0, 'x', '') as isPwd         , -- Bit 17 (     131072):                     (dbAttachSavePWD  )
   iif( int(obj.Flags/(2^29)) mod 2 <> 0, 'x', '') as isAttachedODBC, -- Bit 29 (  536870912): linked ODBC         (dbAttachedODBC   )
   iif( int(obj.Flags/(2^30)) mod 2 <> 0, 'x', '') as isLinkedTable , -- Bit 30 ( 1073741824): linked non-ODBC     (dbAttachedTable  )
   iif( int(obj.Flags/(2^31)) mod 2 <> 0, 'x', '') as isSystemObject, -- Bit 31 (-2147483646):                     (dbSystemObject   )
-- iif( int(obj.type/(2^0)) mod 2 <> 0, 'x', '') as isLocalTable, -- Bit 0: Local table
-- obj.Type,
   obj.Flags,
   obj.LvProp,
   obj.DateCreate,
   obj.DateUpdate
from
   MSysObjects     obj                           left join
   MSysObjects     par on obj.ParentId = par.Id
order by
   obj.Name
