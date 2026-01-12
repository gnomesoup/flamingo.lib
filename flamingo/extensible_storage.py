from Autodesk.Revit import DB
from flamingo import FLAMINGO_VERSION
import json
from pyrevit import HOST_APP, script
from clr import StrongBox
import System
from System import String
from System.Collections.Generic import Dictionary, IDictionary

LOGGER = script.get_logger()

FLAMINGO_SCHEMA_GUID = "476690c0-e882-4b36-90a1-1c1261f33ba4"


def GetSchemaByName(schemaName, doc=None):
    doc = doc or HOST_APP.doc
    schemas = []
    for docSchema in DB.ExtensibleStorage.Schema.ListSchemas():
        if docSchema.SchemaName == schemaName:
            schema = docSchema
            schemas.append(schema)
    return schemas


def CreateSchemaBuilder(schemaGuid, schemaName, vendorId, doc=None):
    doc = doc or HOST_APP.doc
    if not type(schemaGuid) is System.Guid:
        schemaGuid = System.Guid(schemaGuid)
    public = DB.ExtensibleStorage.AccessLevel.Public
    schemaBuilder = DB.ExtensibleStorage.SchemaBuilder(schemaGuid)
    schemaBuilder.SetReadAccessLevel(public)
    schemaBuilder.SetWriteAccessLevel(public)
    schemaBuilder.SetSchemaName(schemaName)
    schemaBuilder.SetVendorId(vendorId)

    return schemaBuilder


def SetSchemaData(schema, fieldName, element, data):
    entity = DB.ExtensibleStorage.Entity(schema)
    field = schema.GetField(fieldName)
    entity.Set[System.String](field, data)
    element.SetEntity(entity)
    retrievedEntity = element.GetEntity(schema)
    data = retrievedEntity.Get[str](field)
    return data


def SetSchemaMapData(schema, fieldName, element, dictionary):
    entity = DB.ExtensibleStorage.Entity(schema)
    field = schema.GetField(fieldName)
    entity.Set[IDictionary[String, String]](field, dictionary)
    element.SetEntity(entity)
    data = entity.Get[IDictionary[String, String]](field)
    return data


def GetSchemaMapData(schema, fieldName, element):
    entity = element.GetEntity(schema)
    data = entity.Get[IDictionary[String, String]](schema.GetField(fieldName))
    return data


def GetFlamingoSchema(doc=None):
    doc = doc or HOST_APP.doc
    flamingoSchema = DB.ExtensibleStorage.Schema.Lookup(
        System.Guid(FLAMINGO_SCHEMA_GUID)
    )
    if flamingoSchema:
        LOGGER.debug("Returning existing Flamingo schema.")
        return flamingoSchema
    LOGGER.debug("Creating new Flamingo schema.")
    schemaBuilder = CreateSchemaBuilder(
        FLAMINGO_SCHEMA_GUID, "Flamingo", "Flamingo", doc
    )
    fieldBuilder = schemaBuilder.AddMapField("Settings", System.String, System.String)
    fieldBuilder.SetDocumentation("JSon string that includes a dictionary settings.")
    fieldBuilder = schemaBuilder.AddSimpleField("Version", System.String)
    fieldBuilder.SetDocumentation("Flamingo version information")
    flamingoSchema = schemaBuilder.Finish()
    entity = DB.ExtensibleStorage.Entity(flamingoSchema)
    entity.Set(flamingoSchema.GetField("Version"), FLAMINGO_VERSION)
    doc.ProjectInformation.SetEntity(entity)
    return flamingoSchema


def GetFlamingoSetting(settingKey, default=None, doc=None, schema=None):
    doc = doc or HOST_APP.doc
    schema = schema or GetFlamingoSchema(doc)
    projectInfo = doc.ProjectInformation
    settings = GetSchemaMapData(schema, "Settings", projectInfo)
    LOGGER.debug("Current settings: {}".format(settings))
    value = StrongBox[String]()
    if settings.TryGetValue(String(settingKey), value):
        return str(value.Value)
    LOGGER.debug("Setting '{}' not found, returning default.".format(settingKey))
    return default


def SetFlamingoSetting(settingKey, settingValue, doc=None, schema=None):
    doc = doc or HOST_APP.doc
    schema = schema or GetFlamingoSchema(doc)
    projectInfo = doc.ProjectInformation
    entity = projectInfo.GetEntity(schema)
    LOGGER.debug("Checking for existing settings...")
    if entity.RecognizedField(schema.GetField("Settings")):
        settings = entity.Get[IDictionary[String, String]](schema.GetField("Settings"))
        LOGGER.debug("Existing settings found. settings: {}".format(settings))
    else:
        LOGGER.debug("No existing settings found. Creating new settings dictionary...")
        settings = Dictionary[String, String]()
    settings[String(settingKey)] = String(settingValue)
    LOGGER.debug("settings: {}".format(settings))
    entity.Set[IDictionary[String, String]](schema.GetField("Settings"), settings)
    LOGGER.debug("Setting entity to project info")
    projectInfo.SetEntity(entity)
