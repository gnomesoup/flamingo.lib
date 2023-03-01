from Autodesk.Revit import DB
import json
from pyrevit import HOST_APP
import System

def GetSchemaByName(schemaName, doc=None):
    doc = doc or HOST_APP.doc
    schemas = []
    for docSchema in DB.ExtensibleStorage.Schema.ListSchemas():
        if docSchema.SchemaName == schemaName:
            schema = docSchema
            schemas.append(schema)
    return schemas

def CreateSchemaBuilder(schemaGuid, schemaName, doc=None):
    doc = doc or HOST_APP.doc
    if not type(schemaGuid) is System.Guid:
        schemaGuid = System.Guid(schemaGuid)
    public = DB.ExtensibleStorage.AccessLevel.Public
    schemaBuilder = DB.ExtensibleStorage.SchemaBuilder(schemaGuid)
    schemaBuilder.SetReadAccessLevel(public)
    schemaBuilder.SetWriteAccessLevel(public)
    schemaBuilder.SetSchemaName(schemaName)

    return schemaBuilder

def SetSchemaData(schema, fieldName, element, data):
    entity = DB.ExtensibleStorage.Entity(schema)
    field = schema.GetField(fieldName)
    entity.Set[System.String](field, data)
    element.SetEntity(entity)
    retrievedEntity = element.GetEntity(schema)
    data = retrievedEntity.Get[str](field)
    return data
    