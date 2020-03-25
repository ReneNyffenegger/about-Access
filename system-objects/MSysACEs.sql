select
   obj.name, -- Same object can appear multiple times in result.
   ACM,
   FInheritable,
   SID
from
   MSysACEs     ace                            left join
   MSysObjects  obj on ace.ObjectId = obj.Id
