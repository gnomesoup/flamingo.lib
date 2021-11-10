from pyrevit import clr, DB
from math import ceil
import uuid

clr.AddReference("System")
from System.Collections.Generic import List

def MakeSolid(minPoint, maxPoint):
    """
    Create a solid cube used to check intersection points for placing
    spot coordinates.
    """
    pt0 = DB.XYZ(minPoint.X, minPoint.Y, minPoint.Z)
    pt1 = DB.XYZ(maxPoint.X, minPoint.Y, minPoint.Z)
    pt2 = DB.XYZ(maxPoint.X, maxPoint.Y, minPoint.Z)
    pt3 = DB.XYZ(minPoint.X, maxPoint.Y, minPoint.Z)
    curves = List[DB.Curve]()
    curves.Add(DB.Line.CreateBound(pt0, pt1))
    curves.Add(DB.Line.CreateBound(pt1, pt2))
    curves.Add(DB.Line.CreateBound(pt2, pt3))
    curves.Add(DB.Line.CreateBound(pt3, pt0))
    curveLoop = DB.CurveLoop.Create(curves)
    loopList = List[DB.CurveLoop]()
    loopList.Add(curveLoop)
    height = maxPoint.Z - minPoint.Z
    centerPoint = DB.XYZ(maxPoint.X - (maxPoint.X - minPoint.X)/2,
                         maxPoint.Y - (maxPoint.Y - minPoint.Y)/2,
                         maxPoint.Z - (maxPoint.Z - minPoint.Z)/2)
    solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        loopList, DB.XYZ.BasisZ, height
    )
    return solid

def GetFaceWithNormal(solid, normal):
    faces = (solid.Faces)
    for face in faces:
        faceNormal = face.ComputeNormal(DB.UV())
        if faceNormal.IsAlmostEqualTo(normal):
            return face

def GetMinMaxPoints(element, view):
    boundingBox = element.get_BoundingBox(view)
    if boundingBox:
        maximumPoint = boundingBox.Max
        minimumPoint = boundingBox.Min
    return[minimumPoint, maximumPoint]

def GetMidPointIntersections(
    doc, origin, vector, outline, gridSpacing, referenceList
):
    """
    Take a list of curves or walls and generate points where those curves
    intersect a grid produced off the bounding area of the elements
    and the spacing specified
    """
    # Get number of divisions in grid
    xLength = outline.MaximumPoint.X - outline.MinimumPoint.X
    xCount = ceil(xLength / gridSpacing)
    xGridsToMin = ceil((origin.X - outline.MinimumPoint.X) / gridSpacing)
    xMin = origin.X - (gridSpacing * xGridsToMin)
    yLength = outline.MaximumPoint.Y - outline.MinimumPoint.Y
    yCount = ceil(yLength / gridSpacing)
    yGridsToMin = ceil((origin.Y - outline.MinimumPoint.Y) / gridSpacing)
    yMin = origin.Y - (gridSpacing * yGridsToMin)

    # Generate a list of points representing interval markers
    points = {}
    if DB.XYZ.BasisX.IsAlmostEqualTo(vector):
        solidMin = DB.XYZ(
            xMin - gridSpacing,
            outline.MinimumPoint.Y,
            outline.MinimumPoint.Z - 1
        )
        solidMax = DB.XYZ(
            xMin,
            outline.MaximumPoint.Y,
            outline.MaximumPoint.Z + 1
        )
        pointList = [
            DB.XYZ((x * gridSpacing) + xMin, origin.Y, 0) 
            for x in range(int(xCount))
        ]
    else:
        solidMin = DB.XYZ(
            outline.MinimumPoint.X,
            yMin - gridSpacing,
            outline.MinimumPoint.Z - 1
        )
        solidMax = DB.XYZ(
            outline.MaximumPoint.X,
            yMin,
            outline.MaximumPoint.Z + 1
        )
        pointList = [
            DB.XYZ(origin.X, (x * gridSpacing) + yMin, 0) 
            for x in range(int(yCount))
        ]
    # Generate a solid to move around and find interections
    solid = MakeSolid(solidMin, solidMax)
    for point in pointList:
        translation = DB.Transform.CreateTranslation(
            point.Subtract(pointList[0])
        )
        iSolid = DB.SolidUtils.CreateTransformed(solid, translation)
        face = GetFaceWithNormal(iSolid, vector)
        for reference in referenceList:
            try:
                if type(reference) is DB.Grid:
                    curve = reference.Curve
                elif type(reference) is DB.Reference:
                    curve = doc.GetElement(reference)\
                        .GetGeometryObjectFromReference(reference)\
                        .AsCurve()
                else:
                    curve = reference.Location.Curve
            except Exception as e:
                curve = reference
            intersectInfo = clr.StrongBox[DB.IntersectionResultArray]()
            if str(face.Intersect(curve, intersectInfo)) == "Overlap":
                intersectionPoint = None
                for j in range(intersectInfo.Size):
                    intersectionPoint = intersectInfo.get_Item(j).XYZPoint
                    points[str(uuid.uuid1())] = {
                        "reference": reference,
                        "point": intersectionPoint,
                        "primary": False
                    }
        iSolid.Dispose()
    solid.Dispose()
    return points

def GetWallLocationCurve(wall, locationLine):
    wallCurve = wall.Location.Curve
    detailLineType = None
    wallType = wall.WallType
    if locationLine == "Core Face: Exterior":
        offsetLength = (wallType.Width)/2
        compoundStructure = wallType.GetCompoundStructure()
        outsideCoreIndex = compoundStructure.GetFirstCoreLayerIndex()
        if outsideCoreIndex > 0:
            for l in range(outsideCoreIndex):
                offsetLength = offsetLength - compoundStructure.GetLayerWidth(l)
        if wallCurve.GetType().Name == "Line":
            if wall.Flipped:
                referenceVector = DB.XYZ(0, 0, 1)
            else:
                referenceVector = DB.XYZ(0, 0, -1)
        else:
            curveVector = (wallCurve.GetEndPoint(0) - wallCurve.Center)\
                .Normalize()
            referenceVector = wallCurve.Normal
            if not curveVector.IsAlmostEqualTo(wall.Orientation):
                offsetLength = offsetLength * -1
        curveOut = wallCurve.CreateOffset(offsetLength, referenceVector)
        detailLineType = "solid"
    elif locationLine == "Core Face: Interior":
        offsetLength = (wallType.Width)/2
        compoundStructure = wallType.GetCompoundStructure()
        insideCoreIndex = compoundStructure.GetLastCoreLayerIndex()
        if insideCoreIndex > 0:
            for l in range(insideCoreIndex + 1, compoundStructure.LayerCount):
                offsetLength = offsetLength - compoundStructure.GetLayerWidth(l)
        if wallCurve.GetType().Name == "Line":
            if wall.Flipped:
                referenceVector = DB.XYZ(0, 0, -1)
            else:
                referenceVector = DB.XYZ(0, 0, 1)
        else:
            curveVector = (wallCurve.GetEndPoint(0) - wallCurve.Center)\
                .Normalize()
            referenceVector = wallCurve.Normal
            if curveVector.IsAlmostEqualTo(wall.Orientation):
                offsetLength = offsetLength * -1
        curveOut = wallCurve.CreateOffset(offsetLength, referenceVector)
        detailLineType = "solid"
    elif locationLine in ["Finished Face: Exterior",
                          "Finished Face: Interior"]:
        if locationLine == "Finished Face: Exterior":
            shellLayer = DB.ShellLayerType.Exterior
        else:
            shellLayer = DB.ShellLayerType.Interior
        faceReference = DB.HostObjectUtils.GetSideFaces(wall, shellLayer)
        face = wall.GetGeometryObjectFromReference(faceReference[0])
        # for faceCurve in face.GetEdgesAsCurveLoops():
        curveOut = face 
    else:
        curveOut = wallCurve
    return curveOut, detailLineType