# pyright: reportMissingImports = false
from pyrevit import DB, revit

from .field import schema_field
from .entity import Entity, transfer_field_data
from .schema import SchemaMeta, list_similar_schemas

ES = DB.ExtensibleStorage


simple_field = lambda value_type, spec_type_id=None, guid=None: schema_field(
    ES.ContainerType.Simple, value_type, spec_type_id=spec_type_id, guid=guid
)
array_field = lambda value_type, spec_type_id=None, guid=None: schema_field(
    ES.ContainerType.Array, value_type, spec_type_id=spec_type_id, guid=guid
)
map_field = lambda key_type, value_type, spec_type_id=None, guid=None: schema_field(
    ES.ContainerType.Map,
    value_type,
    key_type=key_type,
    spec_type_id=spec_type_id,
    guid=guid,
)


class BaseSchema(Entity):
    """Base class for user-defined schemas with Extensible Storage

    Simplifies schema creation and management in Revit by providing standard
    initialization, entity management, and context handling for updates

    Schemas are immutable and exist for the lifetime of the Revit project

    Attributes:
        guid (Guid or str):
            The unique identifier for the schema. Must be set in subclasses
        schema_name (str):
            The name of the schema. Default is the name of the subclass
        vendor_id (str, optional):
            Identifier for the schema vendor. Default is None
        application_guid (Guid or str, optional):
            Identifier for the application that created the schema. Default is None
        read_access_level (AccessLevel, optional):
            Access level for reading schema data. Default is public
        write_access_level (AccessLevel, optional):
            Access level for writing schema data. Default is public

    Args:
        element (Element): The Revit element from which to get an existing entity
        update (bool): Flag indicating whether to update the entity
            when entering the context manager

    Field Types:
        You can specify the field types using string representations in snake_case\n
        The following types are allowed:

        Double, Single, XYZ and UV types requires units to be specified.

        - `int32` (Int32)
        - `int16` (Int16)
        - `byte` (Byte)
        - `double` (Double)
        - `single` (Single)
        - `boolean` (Boolean)
        - `string` (String) - 16MB limit
        - `guid` (Guid)
        - `element_id` (ElementId)
        - `xyz` (XYZ)
        - `uv` (UV)
        - `entity` (Entity)

        Python types can also be used:
        - `int` -> `Int32`
        - `float` -> `Double`
        - `bool` -> `Boolean`
        - `str` -> `String`
    """

    __metaclass__ = SchemaMeta

    guid = None
    schema_name = None

    vendor_id = None
    application_guid = None

    read_access_level = ES.AccessLevel.Public
    write_access_level = ES.AccessLevel.Public

    def __init__(self, element, update=True, **kwargs):
        self.element = element
        self.update = update

        entity = self.element.GetEntity(type(self).schema)
        if not entity.IsValid():
            entity = ES.Entity(type(self).schema)

        super(BaseSchema, self).__init__(entity)

    def __enter__(self):
        if self.update:
            update_schema_entities(self.element, self._wrapped)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        tx_name = "Update Entity: {} % {}".format(self.element.Name, self.schema_name)
        with revit.Transaction(tx_name, self.element.Document):
            self.element.SetEntity(self._wrapped)


def update_schema_entities(element, entity, remove_old=True):
    """Update schema entities associated with a Revit element

    This function replaces outdated schema entities with the new entity data,
    optionally removing old entities after data transfer

    Args:
        element (Element): The Revit element containing the outdated entities
        entity (Entity): The new entity to update with
        remove_old (bool, optional): Flag to indicate if old entities should
            be removed after updating. Defaults to True
    """
    for schema in list_similar_schemas(entity.Schema):
        old_entity = element.GetEntity(schema)
        if not old_entity.IsValid():
            continue

        transfer_field_data(old_entity, entity)

        if remove_old:
            tx_name = "Delete Entity: {} % {}".format(element.Name, schema.SchemaName)
            with revit.Transaction(tx_name, element.Document):
                element.DeleteEntity(old_entity.Schema)
