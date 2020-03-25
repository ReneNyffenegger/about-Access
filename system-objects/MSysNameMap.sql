select
   map.id,
   map.name,
   map.nameMap,
   map.type     as mapType,
   par.Name     as parentName,
   map.guid
-- obj.*
from (
   MSysNameMap map                              left join
   MsysObjects obj on map.name     = obj.name ) left join
   MsysObjects par on obj.parentId = par.id
