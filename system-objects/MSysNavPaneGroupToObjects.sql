select
   pgo.Name                   as pgr_ObjectName,
   pgr.Flags                  as pgr_Flags,
   pgr.Icon,
   pgr.Name                   as pgr_Name,
   pgr.Position               as pgr_Position,
   cat.Filter,
   cat.Flags                  as cat_Flags,
-- cat.Id                     as cat_Id,
   cat.Name                   as cat_Name,
   cat.Position               as cat_Position,
-- cat.SelectedObjectId,
   obs.Name                   as SelectedObjectName,
   cat.Type                   as cat_Type,
   grp.Flags,
   grp.GroupCategoryId,
   grp.Id,
   grp.Name,
   grp.[Object Type Group],
   grp.ObjectId,
   grp.Position
from (((
   MSysNavPaneGroupToObjects  pgr                                    left join
   MSysNavPaneGroups          grp on pgr.GroupId          = grp.Id ) left join
   MSysNavPaneGroupCategories cat on grp.GroupCategoryId  = cat.Id ) left join
   MSysObjects                pgo on pgr.ObjectId         = pgo.Id ) left join
   MSysObjects                obs on cat.SelectedObjectId = obs.Id
order by
   grp.Id,
   pgr.Id
