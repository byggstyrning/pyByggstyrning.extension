# pyright: reportMissingImports = false
import System
from System.Collections.Generic import Dictionary, IDictionary, IList, List

from pyrevit import DB, framework, revit, script


ES = DB.ExtensibleStorage
logger = script.get_logger()

ALLOWED_VALUE_TYPES = {
    "int32": System.Int32,
    "int16": System.Int16,
    "byte": System.Byte,
    "double": System.Double,
    "single": System.Single,
    "boolean": System.Boolean,
    "string": System.String,
    "guid": System.Guid,
    "element_id": DB.ElementId,
    "xyz": DB.XYZ,
    "uv": DB.UV,
    "entity": ES.Entity,
}


ALLOWED_KEY_TYPES = {
    "int32": System.Int32,
    "int16": System.Int16,
    "byte": System.Byte,
    "boolean": System.Boolean,
    "string": System.String,
    "guid": System.Guid,
    "element_id": DB.ElementId,
}


class Field(revit.BaseWrapper):
    """A wrapper for the Extensible Storage Field class"""

    def __init__(self, field):
        super(Field, self).__init__(field)

    @property
    def sub_schema_guid(self):
        """The GUID of the schema describing the subentities stored in this Field"""
        return self._wrapped.SubSchemaGUID

    @property
    def sub_entity_read_access_granted(self):
        """Checks whether there is read access to subentities storable in this field"""
        return self._wrapped.SubEntityReadAccessGranted()

    @property
    def sub_entity_write_access_granted(self):
        """Checks whether there is write access to subentities storable in this field"""
        return self._wrapped.SubEntityWriteAccessGranted()

    @property
    def spec_type_id(self):
        """Gets the spec describing the values stored in this field"""
        return self._wrapped.GetSpecTypeId()

    @property
    def field_type(self):
        """Gets the type of the field"""
        return determine_field_type(self._wrapped)

    def compatible_unit(self, unit_type_id):
        """Checks if the specified unit is compatible with the field description

        Args:
            unit_type_id (ForgeTypeId): The ForgeTypeId of the unit

        Returns:
            bool: True if the unit is compatible
        """
        return self._wrapped.CompatibleUnit(unit_type_id)


class FieldDescriptor(object):

    def __init__(self, name, container_type, value_type, key_type=None, **kwargs):
        self.name = name
        self.__doc__ = kwargs.get("documentation", "")

        # Validate the container type
        if not isinstance(container_type, ES.ContainerType):
            raise ValueError("Invalid field container type: {}".format(container_type))
        self.container_type = container_type

        # Validate the value type
        self.value_type = resolve_type(value_type, ALLOWED_VALUE_TYPES)

        # Validate the key type
        self.key_type = key_type
        if container_type == ES.ContainerType.Map:
            if key_type is None:
                raise ValueError("Missing key type for Map field")
            else:
                self.key_type = resolve_type(key_type, ALLOWED_KEY_TYPES)

        # Validate the spec type
        self.spec_type_id = kwargs.get("spec_type_id", None)
        if self.spec_type_id is not None:
            if not isinstance(self.spec_type_id, DB.ForgeTypeId):
                raise ValueError("Invalid spec type: {}".format(self.spec_type_id))

        # Validate the sub schema guid
        self.sub_schema_guid = kwargs.get("sub_schema_guid", None)
        if value_type == ES.Entity:
            if not self.sub_schema_guid:
                raise ValueError("A valid sub schema GUID must be provided")

            if isinstance(self.sub_schema_guid, str):
                self.sub_schema_guid = System.Guid(self.sub_schema_guid)

    def __get__(self, instance, owner):
        """Retrieve the value of the field from an instance of the schema.

        Args:
            instance (Entity): The entity that holds the value of the field.
                This is the `Entity` wrapper
            owner (Schema): The schema class that owns the field.
                This is the user-defined schema class that inherits from `BaseSchema`

        Returns:
            Field or Any: If `instance` is `None`, returns a `Field` object.
            Otherwise, returns the value of the field from the `instance`

        """
        field = owner.get_field(self.name)
        if instance is None:
            return Field(field)

        return instance.get(field)

    def __set__(self, instance, value):
        """
        Set the value of the field on an instance of the schema.

        Args:
            instance (Entity): The entity that holds the value of the field.
                This is the `Entity` wrapper
            value (Any): The value to set for the field. See `ALLOWED_VALUE_TYPES`
        """
        instance.set(self.name, value)


def schema_field(container_type, value_type, key_type=None, **kwargs):
    """A decorator to mark a method in a schema class as representing a field in the
    Extensible Storage schema

    Args:
        container_type (ContainerType): The type of container for the field\n
        value_type (Type): The value type to stored in the field\n
            See `ALLOWED_VALUE_TYPES`
        key_type (Type, optional): The key type for map fields. Defaults to None\n
            See `ALLOWED_KEY_TYPES`
        spec_type_id (ForgeTypeId, optional): An optional `ForgeTypeId` representing
            the field's specification type (for associating units with a field).
            Defaults to None, or `DB.SpecTypeId.Number` if required
        guid (str or System.Guid, optional): The schema GUID of the Entities
            that are intended to be stored in this field. Defaults to None

    Returns:
        FieldDescriptor: A descriptor for a schema field
    """
    spec_type_id = kwargs.get("spec_type_id", None)
    sub_schema_guid = kwargs.get("guid", None)

    def decorator(func):
        return FieldDescriptor(
            name=func.__name__,
            container_type=container_type,
            value_type=value_type,
            key_type=key_type,
            documentation=func.__doc__,
            spec_type_id=spec_type_id,
            sub_schema_guid=sub_schema_guid,
        )

    return decorator


def determine_field_type(field):
    """Determine the type of a field based on its container type.

    Args:
        field (Field): The field object for which to determine the type

    Raises:
        ValueError: Raised if the field's container type is unrecognized

    Returns:
        Type: The determined type for the field, can be one of the following:
            - `Type`: For simple fields
            - `IList[value_type]`: For array fields
            - `IDictionary[key_type, value_type]`: For map fields
    """
    if field.ContainerType == ES.ContainerType.Simple:
        return framework.get_type(field.ValueType)

    if field.ContainerType == ES.ContainerType.Array:
        value_type = framework.get_type(field.ValueType)
        return IList[value_type]

    if field.ContainerType == ES.ContainerType.Map:
        key_type = framework.get_type(field.KeyType)
        value_type = framework.get_type(field.ValueType)
        return IDictionary[key_type, value_type]

    raise ValueError("Invalid field container type: {}".format(field.ContainerType))


def convert_to_generic(value):
    """Convert a value to a generic type compatible with Extensible Storage"""
    if isinstance(value, (list, tuple, set)):
        # Determine the type of the elements from the first item
        value_type = framework.get_type(type(next(iter(value))))
        try:
            return List[value_type](value)
        except TypeError:
            raise

    if isinstance(value, dict):
        # Determine the types of keys and values from the by first item
        key_type = framework.get_type(type(next(iter(value.keys()))))
        value_type = framework.get_type(type(next(iter(value.values()))))
        try:
            return Dictionary[key_type, value_type](value)
        except TypeError:
            raise

    return value


def resolve_type(data_type, allowed_types):
    """Resolve and return the appropriate type based on a given input

    Args:
        data_type (Type or str): The data type to resolve\n
            Can be a `Type` or string provide in snake_case.
        allowed_types (dict): A dictionary mapping of the allowed data types

    Raises:
        ValueError: If the data type cannot be resolved or is not an allowed type

    Returns:
        Type: The resolved type corresponding to the input data type
    """

    if data_type in allowed_types.values():
        return data_type

    if data_type is int:
        return System.Int32
    if data_type is float:
        return System.Double
    if data_type is bool:
        return System.Boolean
    if data_type is str:
        return System.String

    if isinstance(data_type, str):
        if data_type in allowed_types:
            return allowed_types[data_type]

    raise ValueError("Invalid data type '{}'".format(data_type))
