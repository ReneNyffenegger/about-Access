select
   qry.Attribute,
   qry.Expression,
   qry.Flag,
   qry.LvExtra,
   qob.Name         as ObjectName,
   qry.Name1        as columnName, -- ?
   qry.Name2        as alias
from
   MSysQueries qry  left join
   MSysObjects qob on qry.ObjectId = qob.Id
order by
   qob.Name
