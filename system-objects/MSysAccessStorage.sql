select
   str.name,
   par.name     as parentName,
   ppr.name     as grandParentName,
   str.Id,
   str.Type,
   str.ParentId,
   str.DateCreate,
   str.DateUpdate,
   str.Lv
from (
   MSysAccessStorage str                            left join
   MSysAccessStorage par on str.ParentId = par.Id ) left join
   MSysAccessStorage ppr on par.ParentId = ppr.Id
