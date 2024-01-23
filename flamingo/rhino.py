from pyrevit import script
import clr

# Get logger and output from script
LOGGER = script.get_logger()
OUTPUT = script.get_output()

# Add references to Rhino libraries
clr.AddReference("RhinoCommon")
clr.AddReference("RhinoInside.Revit")
from Rhino import RhinoDoc, DocObjects

# Function to find or add a layer in Rhino
def FindOrAddRhinoLayer(layerFullPath, rhinoDoc=None):
    # Log the function call
    LOGGER.debug(
        "FindOrAddRhinoLayer(layerFullPath={}, rhinoDoc={})".format(
            layerFullPath, rhinoDoc
        )
    )
    # If no document is provided, use the active document
    rhinoDoc = rhinoDoc or RhinoDoc.ActiveDoc
    # Try to find the layer by its full path
    layerIndex = rhinoDoc.Layers.FindByFullPath(layerFullPath, -1)
    # If the layer is found, return its index
    if not layerIndex < 0:
        LOGGER.debug("Layer Found: layerIndex = {}".format(layerIndex))
        return layerIndex
    # If the layer is not found, split the full path into individual layer names
    layerNameSplit = layerFullPath.split("::")
    parentLayerId = None
    # Iterate over the layer names
    for i, layerName in enumerate(layerNameSplit):
        # Construct the current path
        currentPath = "::".join(layerNameSplit[0 : i + 1])
        LOGGER.debug("currentPath = {}".format(currentPath))
        # Try to find the layer by the current path
        layerIndex = rhinoDoc.Layers.FindByFullPath(currentPath, -1)
        LOGGER.debug("parentLayerId = {}".format(parentLayerId))
        # If the layer is found, get its parent layer
        if not layerIndex < 0:
            LOGGER.debug("Matched: layerIndex = {}".format(layerIndex))
            parentLayer = rhinoDoc.Layers.FindIndex(layerIndex)
            LOGGER.debug("type(parentLayer) = {}".format(type(parentLayer)))
            LOGGER.debug("parentLayer.Id = {}".format(parentLayer.Id))
            # Set the parent layer id for the next iteration
            parentLayerId = parentLayer.Id
            continue
        # If the layer is not found, create a new layer
        childLayer = DocObjects.Layer()
        childLayer.Name = layerName
        # If there is a parent layer, set its id as the parent layer id of the new layer
        if parentLayerId:
            childLayer.ParentLayerId = parentLayerId
        # Add the new layer to the document
        layerIndex = rhinoDoc.Layers.Add(childLayer)
        # Get the new layer from the document
        childLayer = rhinoDoc.Layers.FindIndex(layerIndex)
        # Assert that the layer was added successfully
        assert layerIndex >= 0
        # Set the id of the new layer as the parent layer id for the next iteration
        parentLayerId = childLayer.Id
    # Return the index of the last layer added
    return layerIndex