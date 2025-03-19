# pyright: reportMissingImports = false
import System
from pyrevit import DB, script

from .field import FieldDescriptor
from .field_builder import add_field
from .schema_builder import initialize_schema

ES = DB.ExtensibleStorage
logger = script.get_logger()


class SchemaMeta(type):
    """
    Metaclass for defining schemas with Extensible Storage

    This metaclass is used to automatically configure schema-related attributes
    for classes that inherit from `BaseSchema`
    """

    def __new__(cls, name, bases, attrs):
        if name == "BaseSchema":
            return super(SchemaMeta, cls).__new__(cls, name, bases, attrs)

        attrs["guid"] = cls._validate_guid(attrs.get("guid", None))
        attrs["_schema"] = ES.Schema.Lookup(attrs["guid"])
        attrs["schema_name"] = name

        return super(SchemaMeta, cls).__new__(cls, name, bases, attrs)

    @staticmethod
    def _validate_guid(guid):
        if not guid:
            raise ValueError("Schema must have a GUID value")
        return System.Guid(guid) if isinstance(guid, str) else guid

    @property
    def schema(cls):
        """Get the revit extensible storage schema for the class"""
        if cls._schema:
            return cls._schema

        # Build the schema if it doesn't exist
        cls._schema = build_schema(cls)
        return cls._schema

    @property
    def entity(cls):
        """Creates a new entity corresponding to the schema"""
        return ES.Entity(cls.schema)

    @property
    def documentation(cls):
        """The overall description of the schema"""
        return cls.schema.Documentation

    @property
    def read_access_granted(cls):
        """Checks whether entities of this Schema may be retrieved by the current add-in"""
        return cls.schema.ReadAccessGranted()

    @property
    def write_access_granted(cls):
        """Checks whether entities of this Schema may be stored by the current add-in"""
        return cls.schema.WriteAccessGranted()

    @property
    def list_fields(cls):
        """The complete list of fields in the Schema, sorted by name"""
        return list(cls.schema.ListFields())

    def get_field(cls, name):
        """Gets a field of a given name from the schema"""
        return cls.schema.GetField(name)


def build_schema(cls):
    """Initialize an Extensible Storage schema based on a provided class

    Args:
        cls (Type[BaseSchema]): The schema class for which to build the schema\n
        The class should inherit from `BaseSchema`

    Returns:
        Schema: The finished Extensible Storage schema object
    """
    schema_builder = initialize_schema(
        cls.guid,
        cls.schema_name,
        cls.__doc__,
        read_access=cls.read_access_level,
        write_access=cls.write_access_level,
        vendor_id=cls.vendor_id,
        app_guid=cls.application_guid,
    )

    # Add fields to the schema builder
    for base_cls in cls.mro():
        for name, attr in base_cls.__dict__.items():
            if isinstance(attr, FieldDescriptor):
                add_field(schema_builder, attr)

    return schema_builder.Finish()


def list_similar_schemas(schema):
    """Get schemas with the same name as the provided schema, but a different GUID"""
    for s in ES.Schema.ListSchemas():
        if s.SchemaName == schema.SchemaName:
            if s.GUID != schema.GUID:
                yield s
