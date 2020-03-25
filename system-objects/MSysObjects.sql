select
   obj.Name,
   obj.Id,
-- obj.ParentId,
   par.Name        as parentName,
   par.Type        as parentType,
   par.Flags       as parentFlags,
   obj.Type,
   obj.Flags,
   obj.LvProp,
   obj.DateCreate,
   obj.DateUpdate
from
   MSysObjects     obj                           left join
   MSysObjects     par on obj.ParentId = par.Id
order by
   obj.DateCreate
