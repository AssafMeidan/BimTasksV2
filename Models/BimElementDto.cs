using System.Text.Json.Serialization;

namespace BimTasksV2.Models
{
    /// <summary>
    /// Data Transfer Object for BIM Israel web application.
    /// Field names match the JSON specification at:
    /// https://docs.bim-israel.com/revit-json-export-specification
    /// Uses System.Text.Json attributes (not Newtonsoft.Json).
    /// </summary>
    public class BimElementDto
    {
        #region Identification (Required)

        /// <summary>Revit UniqueId - must be stable across exports</summary>
        [JsonPropertyName("ExternalId")]
        public string ExternalId { get; set; }

        /// <summary>Element instance name (from Mark or Name parameter)</summary>
        [JsonPropertyName("Name")]
        public string Name { get; set; }

        /// <summary>UniqueId of the host/parent element (e.g., curtain wall for panels/mullions, wall for doors/windows)</summary>
        [JsonPropertyName("BI_HostElementId")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public string HostElementId { get; set; }

        #endregion Identification (Required)

        #region Classification (Required)

        /// <summary>Revit category name (e.g., "Structural Foundations")</summary>
        [JsonPropertyName("Category")]
        public string Category { get; set; }

        /// <summary>Family name</summary>
        [JsonPropertyName("FamilyName")]
        public string FamilyName { get; set; }

        /// <summary>Type name</summary>
        [JsonPropertyName("TypeName")]
        public string TypeName { get; set; }

        /// <summary>Structural/Non-Structural function</summary>
        [JsonPropertyName("Function")]
        public string Function { get; set; }

        #endregion Classification (Required)

        #region Location

        /// <summary>Associated level name</summary>
        [JsonPropertyName("Level")]
        public string Level { get; set; }

        /// <summary>Level elevation in meters</summary>
        [JsonPropertyName("LevelElevation")]
        public double? LevelElevation { get; set; }

        /// <summary>Revit phase</summary>
        [JsonPropertyName("Phase")]
        public string Phase { get; set; }

        #endregion Location

        #region Geometry (Always from Revit)

        /// <summary>Volume in cubic meters (m3)</summary>
        [JsonPropertyName("Volume")]
        public double Volume { get; set; }

        /// <summary>Area in square meters (m2)</summary>
        [JsonPropertyName("Area")]
        public double Area { get; set; }

        /// <summary>Length in meters (m)</summary>
        [JsonPropertyName("Length")]
        public double Length { get; set; }

        /// <summary>Primary measurement unit (e.g., "m3", "m2", "m")</summary>
        [JsonPropertyName("Unit")]
        public string Unit { get; set; }

        #endregion Geometry (Always from Revit)

        #region Type Information

        /// <summary>Uniformat/OmniClass code</summary>
        [JsonPropertyName("AssemblyCode")]
        public string AssemblyCode { get; set; }

        /// <summary>Assembly description</summary>
        [JsonPropertyName("AssemblyDescription")]
        public string AssemblyDescription { get; set; }

        /// <summary>Type comments parameter</summary>
        [JsonPropertyName("TypeComments")]
        public string TypeComments { get; set; }

        /// <summary>Type description parameter</summary>
        [JsonPropertyName("TypeDescription")]
        public string TypeDescription { get; set; }

        /// <summary>Keynote value</summary>
        [JsonPropertyName("Keynote")]
        public string Keynote { get; set; }

        /// <summary>Type model information</summary>
        [JsonPropertyName("typeModel")]
        public string TypeModel { get; set; }

        #endregion Type Information

        #region BI_ Commercial Parameters - BOQ Assignment

        /// <summary>Contract line item code (e.g., "03-02-001")</summary>
        [JsonPropertyName("BI_BOQ_Code")]
        public string BoqCode { get; set; }

        /// <summary>Location/WBS zone (e.g., "Zone A", "Building 1")</summary>
        [JsonPropertyName("BI_Zone")]
        public string Zone { get; set; }

        /// <summary>Construction phase: FOUNDATIONS, FRAME, BUILDING, FINISHES</summary>
        [JsonPropertyName("BI_WorkStage")]
        public string WorkStage { get; set; }

        #endregion BI_ Commercial Parameters - BOQ Assignment

        #region BI_ Commercial Parameters - Quantity Handling

        /// <summary>Include in BOQ schedules</summary>
        [JsonPropertyName("BI_IsPayItem")]
        public bool? IsPayItem { get; set; }

        /// <summary>Which geometry to use: AREA, VOLUME, LENGTH, COUNT, COMP</summary>
        [JsonPropertyName("BI_QtyBasis")]
        public string QtyBasis { get; set; }

        /// <summary>Manual quantity override</summary>
        [JsonPropertyName("BI_QtyOverride")]
        public double? QtyOverride { get; set; }

        /// <summary>Quantity multiplier (default 1.0)</summary>
        [JsonPropertyName("BI_QtyMultiplier")]
        public double? QtyMultiplier { get; set; }

        #endregion BI_ Commercial Parameters - Quantity Handling

        #region BI_ Commercial Parameters - Pricing

        /// <summary>Price per unit</summary>
        [JsonPropertyName("BI_UnitPrice")]
        public double? UnitPrice { get; set; }

        #endregion BI_ Commercial Parameters - Pricing

        #region BI_ Commercial Parameters - Progress Tracking

        /// <summary>Execution progress percentage (0-100)</summary>
        [JsonPropertyName("BI_ExecPct_ToDate")]
        public double? ExecutionPercentage { get; set; }

        /// <summary>Payment progress percentage (0-100)</summary>
        [JsonPropertyName("BI_PaidPct_ToDate")]
        public double? PaidPercentage { get; set; }

        #endregion BI_ Commercial Parameters - Progress Tracking

        #region BI_ Commercial Parameters - Metadata

        /// <summary>Short note/comment</summary>
        [JsonPropertyName("BI_Note")]
        public string Note { get; set; }

        /// <summary>Links proxy element to source element</summary>
        [JsonPropertyName("BI_SourceElementId")]
        public string SourceElementId { get; set; }

        /// <summary>Role within assembly: DRILLING, CAGE, CONCRETE</summary>
        [JsonPropertyName("BI_ComponentRole")]
        public string ComponentRole { get; set; }

        #endregion BI_ Commercial Parameters - Metadata

        #region Legacy Fields (Hebrew Projects)

        /// <summary>Budget section code</summary>
        [JsonPropertyName("SeifChoze")]
        public string SeifChoze { get; set; }

        /// <summary>Budget section description</summary>
        [JsonPropertyName("SeifChozeDescription")]
        public string SeifChozeDescription { get; set; }

        /// <summary>Include in measurements</summary>
        [JsonPropertyName("IsMedida")]
        public bool? IsMedida { get; set; }

        /// <summary>Total contract price</summary>
        [JsonPropertyName("TotalContractPrice")]
        public double? TotalContractPrice { get; set; }

        /// <summary>Assigned subcontractor</summary>
        [JsonPropertyName("SubcontractorName")]
        public string SubcontractorName { get; set; }

        #endregion Legacy Fields (Hebrew Projects)
    }
}
