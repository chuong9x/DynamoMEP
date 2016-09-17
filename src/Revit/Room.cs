﻿using System;
using DB = Autodesk.Revit.DB;
using DynamoServices;
using Autodesk.DesignScript.Runtime;
using Autodesk.DesignScript.Geometry;
using Autodesk.DesignScript.Interfaces;
using Revit.GeometryConversion;
using RevitServices.Persistence;
using RevitServices.Transactions;
using System.Collections.Generic;
using System.Linq;

namespace Revit.Elements
{
    /// <summary>
    /// MEP Rooms
    /// </summary>
    [DynamoServices.RegisterForTrace]
    public class Pièce : Element, IGraphicItem
    {
        #region Internal Properties

        /// <summary>
        /// Internal reference to the Revit Element
        /// </summary>
        internal DB.Architecture.Room InternalRoom
        {
            get;
            private set;
        }

        /// <summary>
        /// Reference to the Element
        /// </summary>
        public override DB.Element InternalElement
        {
            get { return InternalRoom; }
        }

        /// <summary>
        /// Transform of the Element
        /// </summary>
        internal DB.Transform InternalTransform
        {
            get;
            private set;
        }

        internal List<DB.BoundarySegment> InternalBoundarySegments = new List<DB.BoundarySegment>();

        #endregion

        #region Private constructors

        /// <summary>
        /// Create from an existing Revit Element
        /// </summary>
        /// <param name="room">An existing Revit room</param>
        private Pièce(DB.Architecture.Room room)
        {
            SafeInit(() => InitRoom(room));
        }


        private Pièce(
            DB.Level level,
            DB.UV point)
        {
            SafeInit(() => InitRoom(level, point));
        }

        #endregion

        #region Helpers for private constructors

        /// <summary>
        /// Initialize a Room element
        /// </summary>
        /// <param name="room"></param>
        private void InitRoom(DB.Architecture.Room room)
        {
            InternalSetRoom(room);
        }


        private void InitRoom(DB.Level level, DB.UV point)
        {
            DB.Document document = DocumentManager.Instance.CurrentDBDocument;

            // This creates a new wall and deletes the old one
            TransactionManager.Instance.EnsureInTransaction(document);

            //Phase 1 - Check to see if the object exists and should be rebound
            var roomElem = ElementBinder.GetElementFromTrace<DB.Architecture.Room>(document);

            if (roomElem == null)
                roomElem = document.Create.NewRoom(level, point);

            InternalSetRoom(roomElem);

            TransactionManager.Instance.TransactionTaskDone();

            if (roomElem != null)
            {
                ElementBinder.CleanupAndSetElementForTrace(document, this.InternalElement);
            }
            else
            {
                ElementBinder.SetElementForTrace(this.InternalElement);
            }
        }

        #endregion

        #region Private mutators

        /// <summary>
        /// Set the internal Element, ElementId, and UniqueId
        /// </summary>
        /// <param name="room"></param>
        private void InternalSetRoom(DB.Architecture.Room room)
        {
            InternalRoom = room;
            InternalElementId = room.Id;
            InternalUniqueId = room.UniqueId;
            InternalBoundarySegments = GetBoundarySegment();
            InternalTransform = GetTransform();
        }

        private DB.Transform GetTransform()
        {
            if (InternalElement.Document.GetHashCode() == DocumentManager.Instance.CurrentDBDocument.GetHashCode())
            {
                return DB.Transform.Identity;
            }
            else
            {
                //Find the revit instance where we find the room
                DB.FilteredElementCollector collector = new DB.FilteredElementCollector(DocumentManager.Instance.CurrentDBDocument);
                List<DB.RevitLinkInstance> linkInstances = collector.OfCategory(DB.BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements().Cast<DB.RevitLinkInstance>().ToList();
                DB.RevitLinkInstance roomLinkInstance = linkInstances.FirstOrDefault();
                
                foreach (DB.RevitLinkInstance linkInstance in linkInstances)
                {
                    if (linkInstance.GetLinkDocument().GetHashCode() == InternalElement.Document.GetHashCode())
                    {
                        roomLinkInstance = linkInstance;
                        break;
                    }
                }

                return roomLinkInstance.GetTotalTransform();
            }
        }

        private List<DB.BoundarySegment> GetBoundarySegment()
        {
            List<DB.BoundarySegment> output = new List<DB.BoundarySegment>();
            DB.SpatialElementBoundaryOptions opt = new DB.SpatialElementBoundaryOptions();

            foreach (List<DB.BoundarySegment> segments in InternalRoom.GetBoundarySegments(opt))
            {
                foreach (DB.BoundarySegment segment in segments)
                {
                    output.Add(segment);
                }
            }

             return output.Distinct().ToList();
        }

        #endregion

        #region Public static constructors

        /// <summary>
        /// Créer une pièce
        /// à partir d'un niveau et d'un emplacement
        /// </summary>
        /// <param name="point">Emplacement de la pièce</param>
        /// <param name="niveau">Niveau de la pièce</param>
        /// <returns></returns>
        public static Pièce ParPointEtNiveau(Point point, Level niveau)
        {
            DB.Level revitLevel = niveau.InternalElement as DB.Level;
            DB.XYZ revitPoint = GeometryPrimitiveConverter.ToXyz(point);

            DB.UV uv = new DB.UV(revitPoint.X, revitPoint.Y);

            return new Pièce(revitLevel, uv);
        }

        /// <summary>
        /// Créer une pièce
        /// à partir d'un emplacement
        /// </summary>
        /// <param name="point">Emplacement de la pièce</param>
        /// <returns></returns>
        public static Pièce ParPoint(Point point)
        {
            DB.XYZ revitPoint = GeometryPrimitiveConverter.ToXyz(point);
            DB.Level revitLevel = GetNearestLevel(revitPoint);

            DB.UV uv = new DB.UV(revitPoint.X, revitPoint.Y);

            return new Pièce(revitLevel, uv);
        }

        /// <summary>
        /// Find the nearest level in the active document
        /// </summary>
        /// <param name="point">The reference point</param>
        /// <returns></returns>
        private static DB.Level GetNearestLevel(DB.XYZ point)
        {
            //find all level in the active document
            DB.Document doc = DocumentManager.Instance.CurrentDBDocument;

            DB.FilteredElementCollector collector = new DB.FilteredElementCollector(doc);
            List<DB.Level> activeLevels = collector.OfCategory(DB.BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements().Cast<DB.Level>().ToList();

            DB.Level nearestLevel = activeLevels.FirstOrDefault();
            double delta = Math.Abs(nearestLevel.ProjectElevation - point.Z);

            foreach (DB.Level currentLevel in activeLevels)
            {
                if (Math.Abs(currentLevel.ProjectElevation - point.Z) < delta)
                {
                    nearestLevel = currentLevel;
                    delta = Math.Abs(currentLevel.ProjectElevation - point.Z);
                }
            }

            return nearestLevel;
        }

        /// <summary>
        /// Créer un pièce
        /// à partir d'une pièce existante
        /// </summary>
        /// <param name="element">La pièce d'origine</param>
        /// <returns></returns>
        public static Pièce DepuisExistant(Element element)
        {
            if (element != null)
            {
                if (element.InternalElement.GetType() == typeof(DB.Architecture.Room))
                {
                    return new Pièce(element.InternalElement as DB.Architecture.Room);
                }
                else
                {
                    throw new ArgumentException("The Element is not a Room");
                }
            }
            else
            {
                throw new ArgumentException("An error occured");
            }
        }

        #endregion

        #region public properties

        /// <summary>
        /// Extraire les propriétés
        /// de la pièce
        /// </summary>
        /// <returns name="Nom">Le Nom de la pièce</returns>
        /// <returns name="Numéro">Le Numéro de la pièce</returns>
        [MultiReturn(new[] { "Name", "Number" })]
        public Dictionary<string, string> DonnéesIdentification()
        {
            return new Dictionary<string, string>()
                {
                    {"Nom",InternalRoom.Name},
                    {"Numéro",InternalRoom.Number}
                };
        }

        /// <summary>
        /// Determine si un element
        /// est dans la pièce
        /// </summary>
        public bool EstDansLaPièce(Element element)
        {
            DB.FamilyInstance familyInstance = element.InternalElement as DB.FamilyInstance;
            if (familyInstance != null)
            {
                if (familyInstance.HasSpatialElementCalculationPoint)
                {
                    DB.XYZ insertionPoint = familyInstance.GetSpatialElementCalculationPoint();

                    if (InternalRoom.IsPointInRoom(insertionPoint))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }

            DB.LocationPoint insertionLocationPoint = element.InternalElement.Location as DB.LocationPoint;
            if (insertionLocationPoint != null)
            {
                DB.XYZ insertionPoint = insertionLocationPoint.Point;

                if (InternalRoom.IsPointInRoom(insertionPoint))
                {
                    return true;
                }
            }

            return false;

        }

        /// <summary>
        /// Retourne une grille de points dans la pièce
        /// </summary>
        /// <param name="step">La distance entre deux points de la grille</param>
        public List<Point> Grille(double step)
        {
            step = UnitConverter.DynamoToHostFactor(DB.UnitType.UT_Length) * step;
            List<Point> grid = new List<Point>();

            DB.BoundingBoxXYZ bb = InternalElement.get_BoundingBox(null);

            for (double x = bb.Min.X;x<bb.Max.X;)
            {
                for (double y = bb.Min.Y; y < bb.Max.Y;)
                {
                    DB.XYZ point = new DB.XYZ(x, y, bb.Min.Z);
                    if (InternalRoom.IsPointInRoom(point))
                    {
                        grid.Add(GeometryPrimitiveConverter.ToPoint(InternalTransform.OfPoint(point)));
                    }
                    y = y + step;
                }

                x = x + step;
            }

            return grid;
        }

        /// <summary>
        /// Retourne les limites de la pièce
        /// </summary>
        public List<Element> ElementsLimites
        {
            get
            {
                List<Element> output = new List<Element>();
                DB.Document doc = InternalElement.Document;

                foreach (DB.BoundarySegment segment in InternalBoundarySegments)
                {
                    DB.Element boundaryElement = doc.GetElement(segment.ElementId);
                    if (boundaryElement.GetType() == typeof(DB.RevitLinkInstance))
                    {
                        DB.RevitLinkInstance linkInstance = boundaryElement as DB.RevitLinkInstance;
                        DB.Element linkBoundaryElement = linkInstance.GetLinkDocument().GetElement(segment.LinkElementId);
                        output.Add(ElementWrapper.ToDSType(linkBoundaryElement, true));
                    }
                    else
                    {
                        output.Add(ElementWrapper.ToDSType(boundaryElement, true));
                    }
                }

                output = output.Distinct().ToList();
                return output;
            }
        }

        /// <summary>
        /// Retourne le niveau associé à la pièce
        /// </summary>
        public Level Niveau
        {
            get
            {
                DB.Document doc = InternalElement.Document;
                DB.Element roomLevel = doc.GetElement(InternalElement.LevelId);

                return ElementWrapper.ToDSType(roomLevel, true) as Level;
            }
        }

        /// <summary>
        /// Retourne l'emplacement de la pièce
        /// </summary>
        public Point PointEmplacement
        {
            get
            {
                DB.LocationPoint locPoint = InternalElement.Location as DB.LocationPoint;
                return GeometryPrimitiveConverter.ToPoint(InternalTransform.OfPoint(locPoint.Point));
            }
        }

        /// <summary>
        /// Retrive family instance hosted in boundary elements
        /// This is the base function for Windows and Doors
        /// </summary>
        /// <param name="cat">The category of hosted elements</param>
        /// <returns></returns>
        private List<FamilyInstance> BoundaryFamilyInstance(DB.BuiltInCategory cat)
        {
                List<FamilyInstance> output = new List<FamilyInstance>();

                //the document of the room
                DB.Document doc = InternalElement.Document; // DocumentManager.Instance.CurrentDBDocument;

                //Find boundary elements and their associated document
                List<DB.ElementId> boundaryElements = new List<DB.ElementId>();
                List<DB.Document> boundaryDocuments = new List<DB.Document>();

                foreach (DB.BoundarySegment segment in InternalBoundarySegments)
                {
                    DB.Element boundaryElement = doc.GetElement(segment.ElementId);
                    if (boundaryElement.GetType() == typeof(DB.RevitLinkInstance))
                    {
                        DB.RevitLinkInstance linkInstance = boundaryElement as DB.RevitLinkInstance;
                        boundaryDocuments.Add(linkInstance.GetLinkDocument());
                        boundaryElements.Add(segment.LinkElementId);
                    }
                    else
                    {
                        boundaryDocuments.Add(doc);
                        boundaryElements.Add(segment.ElementId);
                    }
                }

                // Create a category filter
                DB.ElementCategoryFilter filter = new DB.ElementCategoryFilter(cat);
                // Apply the filter to the elements in these documents,
                // Use shortcut WhereElementIsNotElementType() to find family instances in all boundary documents
                boundaryDocuments = boundaryDocuments.Distinct().ToList();
                List<DB.FamilyInstance> familyInstances = new List<DB.FamilyInstance>();
                foreach (DB.Document boundaryDocument in boundaryDocuments)
                {
                    DB.FilteredElementCollector collector = new DB.FilteredElementCollector(boundaryDocument);
                    familyInstances.AddRange(collector.WherePasses(filter).WhereElementIsNotElementType().ToElements().Cast<DB.FamilyInstance>().ToList());
                }

                //Find all family instance hosted on a boundary element
                IEnumerable<DB.FamilyInstance> boundaryFamilyInstances = familyInstances.Where(s => boundaryElements.Contains(s.Host.Id));

                //loop on these boundary family instance to find to and from room
                foreach (DB.FamilyInstance boundaryFamilyInstance in boundaryFamilyInstances)
                {
                    DB.Phase familyInstancePhase = boundaryFamilyInstance.Document.GetElement(boundaryFamilyInstance.CreatedPhaseId) as DB.Phase;
                    if (boundaryFamilyInstance.get_FromRoom(familyInstancePhase) != null)
                    {
                        if (boundaryFamilyInstance.get_FromRoom(familyInstancePhase).Id == InternalRoom.Id)
                        {
                            output.Add(ElementWrapper.ToDSType(boundaryFamilyInstance, true) as FamilyInstance);
                            continue;
                        }
                    }

                    if (boundaryFamilyInstance.get_ToRoom(familyInstancePhase) != null)
                    {
                        if (boundaryFamilyInstance.get_ToRoom(familyInstancePhase).Id == InternalRoom.Id)
                        {
                            output.Add(ElementWrapper.ToDSType(boundaryFamilyInstance, true) as FamilyInstance);
                        }
                    }
                }

                output = output.Distinct().ToList();
                return output;
        }

        /// <summary>
        /// Retourne les fenêtres de la pièce
        /// </summary>
        public List<FamilyInstance> Fenêtres
        {
            get
            {
                return BoundaryFamilyInstance(DB.BuiltInCategory.OST_Windows);
            }
        }

        /// <summary>
        /// Retourne les portes de la pièce
        /// </summary>
        public List<FamilyInstance> Portes
        {
            get
            {
                return BoundaryFamilyInstance(DB.BuiltInCategory.OST_Doors);
            }
        }

        #endregion

        #region Internal static constructors

        /// <summary>
        /// Create a space from an existing reference
        /// </summary>
        /// <param name="room"></param>
        /// <param name="isRevitOwned"></param>
        /// <returns></returns>
        internal static Pièce FromExisting(DB.Architecture.Room room, bool isRevitOwned)
        {
            return new Pièce(room)
            {
                //IsRevitOwned = isRevitOwned
            };
        }

        #endregion

        #region Display Functions

        /// <summary>
        /// Display Spaces in the Dynamo interface
        /// </summary>
        /// <param name="package"></param>
        /// <param name="parameters"></param>
        [IsVisibleInDynamoLibrary(false)]
        public new void Tessellate(IRenderPackage package, TessellationParameters parameters)
        {
            //Ensure that the object is still alive
            if (!IsAlive) return;

            //Location Point
            DB.LocationPoint locPoint = InternalElement.Location as DB.LocationPoint;
            GeometryPrimitiveConverter.ToPoint(InternalTransform.OfPoint(locPoint.Point)).Tessellate(package, parameters);
            package.ApplyPointVertexColors(CreateColorByteArrayOfSize(package.PointVertexCount, 255, 0, 0, 0));

            //Boundaries
            foreach (DB.BoundarySegment segment in InternalBoundarySegments)
            {
                Curve crv = RevitToProtoCurve.ToProtoType(segment.GetCurve().CreateTransformed(InternalTransform));

                crv.Tessellate(package, parameters);

                if (package.LineVertexCount > 0)
                {
                    package.ApplyLineVertexColors(CreateColorByteArrayOfSize(package.LineVertexCount, 255, 0, 0, 0));
                }
            }
        }

        private static byte[] CreateColorByteArrayOfSize(int size, byte red, byte green, byte blue, byte alpha)
        {
            var arr = new byte[size * 4];
            for (var i = 0; i < arr.Length; i += 4)
            {
                arr[i] = red;
                arr[i + 1] = green;
                arr[i + 2] = blue;
                arr[i + 3] = alpha;
            }
            return arr;
        }

        /// <summary>
        /// OPTIONAL:
        /// Overriding ToString allows you to control what is
        /// displayed whenever the object's string representation
        /// is used. For example, ToString is called when the 
        /// object is displayed in a Watch node.
        /// </summary>
        /// <returns>The string representation of our object.</returns>
        public override string ToString()
        {
            return string.Format("Room N°{1} - {0}", InternalRoom.Name, InternalRoom.Number);
        }

        #endregion

    }
}