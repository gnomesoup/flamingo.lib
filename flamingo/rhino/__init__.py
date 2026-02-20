from distutils.version import StrictVersion
from pyrevit import clr, script
import re
from System import Drawing
from System.IO import IOException

clr.AddReference("RhinoCommon")
clr.AddReference("RhinoInside.Revit")

from Rhino import RhinoApp, RhinoDoc

LOGGER = script.get_logger()
OUTPUT = script.get_output()

def check_rhino_version(minimum_version):
    """
    Check if current Rhino version meets minimum requirement
    Args:
        minimum_version (str): Minimum version required (e.g., "7.0")
    Returns:
        bool: True if current version meets or exceeds minimum, False otherwise
    """
    current_version = str(RhinoApp.Version)
    try:
        m = re.match(r"(\d+\.\d+)", current_version)
        if m:
            current_version = m.group(1)
        else:
            raise ValueError("Invalid version string")
        LOGGER.debug(StrictVersion(str(current_version)))
        LOGGER.debug(StrictVersion(minimum_version))
        return StrictVersion(str(current_version)) >= StrictVersion(minimum_version)
    except ValueError as e:
        LOGGER.debug("Version comparison error: {}".format(e))
        return False

def FindOrAddRhinoLayer(layer_name, rhinoDoc=None):
    """
    Find or add a layer in the Rhino document
    Args:
        layer_name (str): Name of the layer to find or add
        rhinoDoc (optional): Rhino document
    Returns:
        The found or newly added layer
    """
    rhinoDoc = rhinoDoc or RhinoDoc.ActiveDoc
    try:
        layer_index = rhinoDoc.Layers.FindByFullPath(layer_name, True)
        if layer_index >= 0:
            return rhinoDoc.Layers[layer_index]
        else:
            new_layer = rhinoDoc.Layers.Add(layer_name, Drawing.Color.White)
            return rhinoDoc.Layers[new_layer]
    except IOException as e:
        LOGGER.debug("Error accessing Rhino document: {}".format(e))
        return None