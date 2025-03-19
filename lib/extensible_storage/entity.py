# pyright: reportMissingImports = false
import sys

from pyrevit import DB, revit, script

from .field import Field, determine_field_type, convert_to_generic

ES = DB.ExtensibleStorage
logger = script.get_logger()


class Entity(revit.BaseWrapper):
    """A wrapper for the Extensible Storage Entity class"""

    def __init__(self, entity=None):
        if entity is None:
            entity = ES.Entity()
        if not isinstance(entity, ES.Entity):
            entity = ES.Entity(entity)
        super(Entity, self).__init__(entity)

    @property
    def schema(self):
        """The schema describing this entity"""
        return self._wrapped.Schema

    @property
    def schema_guid(self):
        """The GUID of the schema describing this entity"""
        return self._wrapped.SchemaGUID

    @property
    def is_valid(self):
        """Checks whether this entity has a live schema corresponding to it"""
        return self._wrapped.IsValid()

    @property
    def read_access_granted(self):
        """Checks whether this entity may be retrieved by the current add-in"""
        return self._wrapped.ReadAccessGranted()

    @property
    def write_access_granted(self):
        """Checks whether this entity may be stored by the current add-in"""
        return self._wrapped.WriteAccessGranted()

    def recognized_field(self, field):
        """Checks whether a field belongs to the same schema as this entity"""

        if type(field) is str:
            field = self.schema.GetField(field)
        if isinstance(field, Field):
            field = field.unwrap()

        return self._wrapped.RecognizedField(field)

    def get(self, field, unit_type_id=None):
        """Retrieves the value of the field in the entity.
        This method will look up the field by name"""
        logger.debug("Getting field '{}'".format(field))

        if type(field) is str:
            field = self.schema.GetField(field)
        if isinstance(field, Field):
            field = field.unwrap()

        if unit_type_id is None:
            unit_type_id = get_default_unit_type_id(field)

        data_type = determine_field_type(field)
        value = self._wrapped.Get[data_type](field, unit_type_id)

        if field.ContainerType == ES.ContainerType.Array:
            return list(value)
        if field.ContainerType == ES.ContainerType.Map:
            return dict(value)
        if isinstance(value, DB.ExtensibleStorage.Entity):
            return Entity(value)

        return value

    def set(self, field, value, unit_type_id=None):
        """Stores the value of the field in the entity.
        This method will look up the field by name

        This method only modifies your copy of the Entity.
        Store the Entity in an element or another Entity to save the new value.
        Write access is checked when you try to save the Entity in an Element
        or another Entity
        """
        logger.debug("Setting field '{}' to {}".format(field, value))

        if type(field) is str:
            field = self.schema.GetField(field)
        if isinstance(field, Field):
            field = field.unwrap()

        if unit_type_id is None:
            unit_type_id = get_default_unit_type_id(field)

        data_type = determine_field_type(field)
        value = convert_to_generic(value)
        if isinstance(value, Entity):
            value = value.unwrap()

        try:
            self._wrapped.Set[data_type](field, value, unit_type_id)

        except TypeError:
            exc_type, exc_value, exc_traceback = sys.exc_info()
            msg = "Invalid type for field '{}': expected {} but got {}"
            e = TypeError(msg.format(field.FieldName, data_type, type(value)))
            raise (e, None, exc_traceback)

    def clear(self, field):
        """Resets the field to its default value\n

        Numeric fields: Default value is zero.\n
        Identifiers and Entities: Default value is invalid.\n
        Strings and Containers: Default value is empty.
        """
        if isinstance(field, Field):
            field = field.unwrap()
        self._wrapped.Clear(field)


def transfer_field_data(from_entity, to_entity):
    """Transfer field data from one schema entity to another.

    Copies the field values from the old entity to the new entity, provided that
    the field names and field types match.

    Args:
        from_entity (Entity): The source entity
        to_entity (Entity): The target entity
    """
    for field in to_entity.Schema.ListFields():
        old_field = from_entity.Schema.GetField(field.FieldName)
        if old_field is None:
            continue  # Field does not exist in the old entity

        field_data_type = determine_field_type(field)
        if field_data_type != old_field.ValueType:
            continue  # Field data types do not match

        # Set a default unit type id
        unit_type_id = get_default_unit_type_id(field)

        # Transfer the data
        old_value = from_entity.Get[field_data_type](field.FieldName, unit_type_id)
        to_entity.Set(field.FieldName, old_value, unit_type_id)


def get_default_unit_type_id(field):
    """Get the default unit type id for a field"""
    spec_type_id = field.GetSpecTypeId()

    # TODO: Add defaults for spec types
    if spec_type_id == DB.SpecTypeId.Number:
        return DB.UnitTypeId.General

    if DB.UnitUtils.IsMeasurableSpec(spec_type_id):
        return DB.UnitUtils.GetValidUnits(spec_type_id)[0]

    # Default to an empty unit type id
    return DB.ForgeTypeId()
