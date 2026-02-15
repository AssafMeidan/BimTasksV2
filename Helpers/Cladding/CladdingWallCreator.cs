using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BimTasksV2.Models;
using System;
using System.Collections.Generic;

namespace BimTasksV2.Helpers.Cladding
{
    /// <summary>
    /// Creates cladding (chipuy) walls parallel to existing walls.
    /// Supports creating walls on both sides, exterior only, or interior only.
    /// </summary>
    public static class CladdingWallCreator
    {
        /// <summary>
        /// Internal cladding wall type name as it appears in Revit.
        /// </summary>
        private const string InternalCladdingWallName = "w2";

        /// <summary>
        /// External cladding wall type name as it appears in Revit.
        /// </summary>
        private const string ExternalCladdingWallName = "w1";

        /// <summary>
        /// Specifies which cladding walls to create.
        /// </summary>
        public enum WallsToCreate
        {
            ExternalAndInternal = 0,
            ExternalOnly = 1,
            InternalOnly = 2
        }

        /// <summary>
        /// Adds cladding walls to both sides of the selected walls.
        /// Creates external and internal cladding, then joins them.
        /// </summary>
        public static void AddCladdingToBothSides(Document doc, List<Wall> selectedWalls)
        {
            if (selectedWalls.Count == 0)
            {
                TaskDialog.Show("Selection Error", "No valid walls were selected for cladding.");
                return;
            }

            var newExternalCladdingWalls = new List<MyWallMember>();
            var newInternalCladdingWalls = new List<MyWallMember>();

            using (Transaction tx = new Transaction(doc, "Create Cladding Walls"))
            {
                tx.Start();
                newExternalCladdingWalls = CreateCladdingWalls(doc, selectedWalls, WallsToCreate.ExternalOnly);
                newInternalCladdingWalls = CreateCladdingWalls(doc, selectedWalls, WallsToCreate.InternalOnly);
                tx.Commit();
            }

            using (Transaction tx2 = new Transaction(doc, "Join Cladding Walls"))
            {
                tx2.Start();
                JoinCladdingWalls(doc, newExternalCladdingWalls);
                JoinCladdingWalls(doc, newInternalCladdingWalls);
                tx2.Commit();
            }

            CladdingWallJoinHelper.ConnectCladdingWalls(doc, newInternalCladdingWalls);
            CladdingWallJoinHelper.ConnectCladdingWalls(doc, newExternalCladdingWalls);

            TaskDialog.Show("Success", "Cladding walls have been successfully added and joined.");
        }

        /// <summary>
        /// Adds cladding walls to the external (positive orientation) side only.
        /// </summary>
        public static void AddCladdingToExternalSide(Document doc, List<Wall> selectedWalls)
        {
            if (selectedWalls.Count == 0)
            {
                TaskDialog.Show("Selection Error", "No valid walls were selected for cladding.");
                return;
            }

            var newExternalCladdingWalls = new List<MyWallMember>();

            using (Transaction tx = new Transaction(doc, "Create Cladding Walls"))
            {
                tx.Start();
                newExternalCladdingWalls = CreateCladdingWalls(doc, selectedWalls, WallsToCreate.ExternalOnly);
                tx.Commit();
            }

            using (Transaction tx2 = new Transaction(doc, "Join Cladding Walls"))
            {
                tx2.Start();
                JoinCladdingWalls(doc, newExternalCladdingWalls);
                tx2.Commit();
            }

            CladdingWallJoinHelper.ConnectCladdingWalls(doc, newExternalCladdingWalls);

            TaskDialog.Show("Success", "Cladding walls have been successfully added and joined.");
        }

        /// <summary>
        /// Adds cladding walls to the internal (negative orientation) side only.
        /// </summary>
        public static void AddCladdingToInternalSide(Document doc, List<Wall> selectedWalls)
        {
            if (selectedWalls.Count == 0)
            {
                TaskDialog.Show("Selection Error", "No valid walls were selected for cladding.");
                return;
            }

            var newInternalCladdingWalls = new List<MyWallMember>();

            using (Transaction tx = new Transaction(doc, "Create Cladding Walls"))
            {
                tx.Start();
                newInternalCladdingWalls = CreateCladdingWalls(doc, selectedWalls, WallsToCreate.InternalOnly);
                tx.Commit();
            }

            using (Transaction tx2 = new Transaction(doc, "Join Cladding Walls"))
            {
                tx2.Start();
                JoinCladdingWalls(doc, newInternalCladdingWalls);
                tx2.Commit();
            }

            CladdingWallJoinHelper.ConnectCladdingWalls(doc, newInternalCladdingWalls);

            TaskDialog.Show("Success", "Cladding walls have been successfully added and joined.");
        }

        /// <summary>
        /// Creates cladding walls for each selected original wall on the specified side.
        /// </summary>
        /// <param name="doc">The Revit document.</param>
        /// <param name="selectedWalls">List of walls selected by the user.</param>
        /// <param name="wallsToCreate">Which side(s) to create cladding on.</param>
        /// <returns>List of MyWallMember wrapping the newly created cladding walls.</returns>
        public static List<MyWallMember> CreateCladdingWalls(Document doc, List<Wall> selectedWalls, WallsToCreate wallsToCreate)
        {
            var myCladdingWalls = new List<MyWallMember>();

            WallType externalCladdingWallType = WallHelper.GetWallTypeByName(doc, ExternalCladdingWallName);
            WallType internalCladdingWallType = WallHelper.GetWallTypeByName(doc, InternalCladdingWallName);

            if (externalCladdingWallType == null && wallsToCreate != WallsToCreate.InternalOnly)
                throw new InvalidOperationException($"External cladding wall type '{ExternalCladdingWallName}' not found.");
            if (internalCladdingWallType == null && wallsToCreate != WallsToCreate.ExternalOnly)
                throw new InvalidOperationException($"Internal cladding wall type '{InternalCladdingWallName}' not found.");

            foreach (Wall selectedWall in selectedWalls)
            {
                var wallParams = WallHelper.GetWallParameters(doc, selectedWall);
                if (wallParams.WallHeight <= 0.0)
                {
                    TaskDialog.Show("Invalid Wall Height", $"Wall ID {selectedWall.Id} has an invalid height. Skipping.");
                    continue;
                }

                // Calculate offsets
                double extWidth = externalCladdingWallType != null ? WallHelper.GetWallTypeWidth(externalCladdingWallType) : 0;
                double intWidth = internalCladdingWallType != null ? WallHelper.GetWallTypeWidth(internalCladdingWallType) : 0;
                var offsets = WallHelper.CalculateOffsets(wallParams.OriginalWallWidth, extWidth, intWidth);

                // Calculate offset vectors using wall orientation
                XYZ orientation = selectedWall.Orientation;
                XYZ offsetVectorPositive = orientation.Multiply(offsets.PositiveOffset);
                XYZ offsetVectorNegative = orientation.Multiply(-offsets.NegativeOffset);

                // Create translation transforms
                Transform translationPositive = Transform.CreateTranslation(offsetVectorPositive);
                Transform translationNegative = Transform.CreateTranslation(offsetVectorNegative);

                // Create new offset curves
                Curve originalCurve = (selectedWall.Location as LocationCurve).Curve;
                Curve newCurvePositive = originalCurve.CreateTransformed(translationPositive);
                Curve newCurveNegative = originalCurve.CreateTransformed(translationNegative);

                if (wallsToCreate == WallsToCreate.ExternalOnly)
                {
                    Wall newWall = WallHelper.CreateParallelWall(
                        doc, newCurvePositive, externalCladdingWallType,
                        wallParams.BaseLevelId, wallParams.WallHeight, wallParams.BaseOffset, selectedWall.Flipped);

                    if (newWall == null)
                    {
                        TaskDialog.Show("Wall Creation Error", $"Failed to create external cladding wall for wall ID {selectedWall.Id}.");
                        continue;
                    }

                    WallHelper.SetTopConstraint(doc, newWall, wallParams.TopConstraintId, wallParams.TopOffset, wallParams.IsUnconnected);
                    WallHelper.CopyAdditionalParameters(selectedWall, newWall);
                    WallHelper.SetWallLocationLine(newWall, WallLocationLineValue.FinishFaceInterior);
                    myCladdingWalls.Add(new MyWallMember(doc, newWall.Id, selectedWall, disallowEndJoin: true));
                }
                else if (wallsToCreate == WallsToCreate.InternalOnly)
                {
                    Wall newWall = WallHelper.CreateParallelWall(
                        doc, newCurveNegative, internalCladdingWallType,
                        wallParams.BaseLevelId, wallParams.WallHeight, wallParams.BaseOffset, selectedWall.Flipped);

                    if (newWall == null)
                    {
                        TaskDialog.Show("Wall Creation Error", $"Failed to create internal cladding wall for wall ID {selectedWall.Id}.");
                        continue;
                    }

                    WallHelper.SetTopConstraint(doc, newWall, wallParams.TopConstraintId, wallParams.TopOffset, wallParams.IsUnconnected);
                    WallHelper.CopyAdditionalParameters(selectedWall, newWall);
                    WallHelper.SetWallLocationLine(newWall, WallLocationLineValue.FinishFaceExterior);
                    myCladdingWalls.Add(new MyWallMember(doc, newWall.Id, selectedWall, disallowEndJoin: true));
                }
            }

            return myCladdingWalls;
        }

        /// <summary>
        /// Joins newly created cladding walls with their respective original walls.
        /// </summary>
        public static void JoinCladdingWalls(Document doc, List<MyWallMember> myCladdingWalls)
        {
            foreach (var claddingWall in myCladdingWalls)
            {
                if (claddingWall.OriginalWallToJoin == null)
                    continue;

                try
                {
                    if (!JoinGeometryUtils.AreElementsJoined(doc, claddingWall.OriginalWallToJoin, claddingWall.MyWall))
                    {
                        JoinGeometryUtils.JoinGeometry(doc, claddingWall.OriginalWallToJoin, claddingWall.MyWall);
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Join Error",
                        $"Failed to join cladding wall ID {claddingWall.MyWall.Id} with original wall: {ex.Message}");
                }
            }
        }
    }
}
