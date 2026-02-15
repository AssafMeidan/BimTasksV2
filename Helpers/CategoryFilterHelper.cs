using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace BimTasksV2.Helpers
{
    /// <summary>
    /// Provides predefined sets of BuiltInCategory and BuiltInParameter values
    /// used for element filtering in the FilterTree and ElementCalculation views.
    /// </summary>
    public static class CategoryFilterHelper
    {
        /// <summary>
        /// Returns a HashSet of 37 BuiltInCategory values commonly used in
        /// BIM quantity takeoff and element analysis operations.
        /// </summary>
        public static HashSet<BuiltInCategory> GetCategoriesToInclude()
        {
            return new HashSet<BuiltInCategory>
            {
                BuiltInCategory.OST_Areas,
                BuiltInCategory.OST_ArcWallRectOpening,
                BuiltInCategory.OST_CeilingOpening,
                BuiltInCategory.OST_Ceilings,
                BuiltInCategory.OST_Doors,
                BuiltInCategory.OST_ElectricalEquipment,
                BuiltInCategory.OST_Entourage,
                BuiltInCategory.OST_FloorOpening,
                BuiltInCategory.OST_Floors,
                BuiltInCategory.OST_Furniture,
                BuiltInCategory.OST_FurnitureSystems,
                BuiltInCategory.OST_GenericModel,
                BuiltInCategory.OST_LightingFixtures,
                BuiltInCategory.OST_Massing,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_Planting,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_Railings,
                BuiltInCategory.OST_RailingSystem,
                BuiltInCategory.OST_Ramps,
                BuiltInCategory.OST_Roofs,
                BuiltInCategory.OST_Rooms,
                BuiltInCategory.OST_Site,
                BuiltInCategory.OST_SpecialityEquipment,
                BuiltInCategory.OST_Sprinklers,
                BuiltInCategory.OST_Stairs,
                BuiltInCategory.OST_StairsLandings,
                BuiltInCategory.OST_StructuralColumns,
                BuiltInCategory.OST_StructuralFoundation,
                BuiltInCategory.OST_StructuralFraming,
                BuiltInCategory.OST_Topography,
                BuiltInCategory.OST_Toposolid,
                BuiltInCategory.OST_Truss,
                BuiltInCategory.OST_Walls,
                BuiltInCategory.OST_Windows
            };
        }

        /// <summary>
        /// Returns a HashSet of ~25 BuiltInParameter values representing
        /// dimensional and offset parameters relevant to quantity analysis.
        /// </summary>
        public static HashSet<BuiltInParameter> GetBuiltParamsToInclude()
        {
            return new HashSet<BuiltInParameter>
            {
                BuiltInParameter.OFFSET_FROM_REFERENCE_BASE,
                BuiltInParameter.ASSOCIATED_LEVEL_OFFSET,
                BuiltInParameter.WALL_BASE_OFFSET,
                BuiltInParameter.WALL_BASE_HEIGHT_PARAM,
                BuiltInParameter.WALL_TOP_OFFSET,
                BuiltInParameter.WINDOW_HEIGHT,
                BuiltInParameter.WINDOW_WIDTH,
                BuiltInParameter.COLUMN_TOP_ATTACHMENT_OFFSET_PARAM,
                BuiltInParameter.COLUMN_BASE_ATTACHMENT_OFFSET_PARAM,
                BuiltInParameter.DOOR_HEIGHT,
                BuiltInParameter.DOOR_WIDTH,
                BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM,
                BuiltInParameter.FAMILY_THICKNESS_PARAM,
                BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM,
                BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM,
                BuiltInParameter.GENERIC_HEIGHT,
                BuiltInParameter.GENERIC_WIDTH,
                BuiltInParameter.HANDRAIL_HEIGHT_PARAM,
                BuiltInParameter.ROOF_CONSTRAINT_OFFSET_PARAM,
                BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM,
                BuiltInParameter.STRUCTURAL_ELEVATION_AT_TOP,
                BuiltInParameter.INSTANCE_HEAD_HEIGHT_PARAM,
                BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM,
                BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM,
                BuiltInParameter.STAIRS_BASE_OFFSET,
            };
        }
    }
}
