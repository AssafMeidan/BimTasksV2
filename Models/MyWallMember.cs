using Autodesk.Revit.DB;
using System;

namespace BimTasksV2.Models
{
    /// <summary>
    /// Wrapper around a Revit Wall element that provides lazy-loaded access
    /// to commonly used wall parameters. Also tracks the original wall
    /// that this wall was derived from (e.g., for cladding operations).
    /// </summary>
    public class MyWallMember
    {
        private readonly Document _document;
        private readonly ElementId _wallId;
        private Wall _wall;
        private Curve _wallCurve;

        // Lazy-loaded parameter backing fields
        private Parameter _unconnectedHeight;
        private Parameter _bottomIsAttached;
        private Parameter _topIsAttached;
        private Parameter _topOffset;
        private Parameter _baseOffset;
        private Parameter _baseConstraint;
        private Parameter _topConstraint;
        private Parameter _locationLine;
        private Parameter _length;
        private Parameter _area;
        private Parameter _volume;
        private Parameter _comments;

        /// <summary>
        /// The wall width (from Wall.Width).
        /// </summary>
        public double? Width => _wall?.Width;

        /// <summary>
        /// The underlying Revit Wall element.
        /// </summary>
        public Wall MyWall => _wall;

        /// <summary>
        /// The wall's location curve, lazily resolved.
        /// </summary>
        public Curve MyWallCurve => _wallCurve ??= (_wall.Location as LocationCurve)?.Curve
            ?? throw new InvalidOperationException("Wall location is not a LocationCurve.");

        /// <summary>
        /// The original wall this cladding wall was created from (for join operations).
        /// </summary>
        public Wall OriginalWallToJoin { get; set; }

        /// <summary>
        /// The ElementId of the original wall to join.
        /// </summary>
        public ElementId OriginalWallToJoinId { get; set; }

        #region Constructors

        /// <summary>
        /// Creates a MyWallMember from a Wall instance.
        /// </summary>
        public MyWallMember(Document doc, Wall wall)
            : this(doc, wall.Id, null as Wall)
        {
        }

        /// <summary>
        /// Creates a MyWallMember from a wall ElementId.
        /// </summary>
        public MyWallMember(Document doc, ElementId wallId)
            : this(doc, wallId, null as Wall)
        {
        }

        /// <summary>
        /// Creates a MyWallMember with an original wall reference specified by ElementId.
        /// </summary>
        public MyWallMember(Document doc, ElementId wallId, ElementId wallToJoinId)
            : this(doc, wallId, doc.GetElement(wallToJoinId) as Wall)
        {
        }

        /// <summary>
        /// Creates a MyWallMember with an original wall reference.
        /// </summary>
        public MyWallMember(Document doc, ElementId wallId, Wall wallToJoin)
        {
            _document = doc ?? throw new ArgumentNullException(nameof(doc));
            _wallId = wallId;
            _wall = doc.GetElement(wallId) as Wall
                ?? throw new ArgumentException("Invalid Wall Id", nameof(wallId));

            OriginalWallToJoin = wallToJoin;
            OriginalWallToJoinId = wallToJoin?.Id;

            if (_wall.Location is LocationCurve lc)
            {
                _wallCurve = lc.Curve;
            }
            else
            {
                throw new InvalidOperationException("Wall location is not a LocationCurve.");
            }
        }

        /// <summary>
        /// Creates a MyWallMember with optional disallow-end-join behavior.
        /// Used for cladding walls that should not auto-join at creation time.
        /// </summary>
        public MyWallMember(Document doc, ElementId wallId, Wall wallToJoin, bool disallowEndJoin)
        {
            _document = doc ?? throw new ArgumentNullException(nameof(doc));
            _wallId = wallId;
            _wall = doc.GetElement(wallId) as Wall
                ?? throw new ArgumentException("Invalid Wall Id", nameof(wallId));

            OriginalWallToJoin = wallToJoin;
            OriginalWallToJoinId = wallToJoin?.Id;

            if (_wall.Location is LocationCurve lc)
            {
                _wallCurve = lc.Curve;
            }
            else
            {
                throw new InvalidOperationException("Wall location is not a LocationCurve.");
            }

            if (disallowEndJoin)
            {
                WallUtils.DisallowWallJoinAtEnd(_wall, 0);
                WallUtils.DisallowWallJoinAtEnd(_wall, 1);
            }
        }

        #endregion Constructors

        #region Lazy-Loaded Parameter Properties

        public Parameter UnconnectedHeight => _unconnectedHeight ??= GetParameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
        public Parameter BottomIsAttached => _bottomIsAttached ??= GetParameter(BuiltInParameter.WALL_BOTTOM_IS_ATTACHED);
        public Parameter TopIsAttached => _topIsAttached ??= GetParameter(BuiltInParameter.WALL_TOP_IS_ATTACHED);
        public Parameter TopOffset => _topOffset ??= GetParameter(BuiltInParameter.WALL_TOP_OFFSET);
        public Parameter BaseOffset => _baseOffset ??= GetParameter(BuiltInParameter.WALL_BASE_OFFSET);
        public Parameter BaseConstraint => _baseConstraint ??= GetParameter(BuiltInParameter.WALL_BASE_CONSTRAINT);
        public Parameter TopConstraint => _topConstraint ??= GetParameter(BuiltInParameter.WALL_HEIGHT_TYPE);
        public Parameter LocationLine => _locationLine ??= GetParameter(BuiltInParameter.WALL_KEY_REF_PARAM);
        public Parameter Length => _length ??= GetParameter(BuiltInParameter.FAMILY_LINE_LENGTH_PARAM);
        public Parameter Area => _area ??= GetParameter(BuiltInParameter.HOST_AREA_COMPUTED);
        public Parameter Volume => _volume ??= GetParameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
        public Parameter Comments => _comments ??= GetParameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);

        #endregion Lazy-Loaded Parameter Properties

        #region Parameter Operations

        /// <summary>
        /// Retrieves a parameter from the wall.
        /// </summary>
        private Parameter GetParameter(BuiltInParameter param)
        {
            var parameter = _wall.get_Parameter(param);
            if (parameter == null)
            {
                throw new InvalidOperationException($"Parameter '{param}' not found on wall ID {_wall.Id}");
            }
            return parameter;
        }

        /// <summary>
        /// Sets a double parameter value for the wall.
        /// </summary>
        public void SetParameter(BuiltInParameter param, double value)
        {
            var parameter = GetParameter(param);
            if (!parameter.IsReadOnly)
            {
                parameter.Set(value);
            }
            else
            {
                throw new InvalidOperationException($"Unable to set parameter '{param}' for wall ID {_wall.Id}");
            }
        }

        /// <summary>
        /// Sets the top constraint for the wall.
        /// </summary>
        public void SetTopConstraint(ElementId constraintId, double offset)
        {
            TopConstraint.Set(constraintId);
            TopOffset.Set(offset);
        }

        /// <summary>
        /// Sets the base constraint for the wall.
        /// </summary>
        public void SetBaseConstraint(ElementId constraintId, double offset)
        {
            BaseConstraint.Set(constraintId);
            BaseOffset.Set(offset);
        }

        #endregion Parameter Operations
    }
}
