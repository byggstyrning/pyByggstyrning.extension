# pyright: reportMissingImports = false
from pyrevit import DB, script

ES = DB.ExtensibleStorage
logger = script.get_logger()


def add_field(schema_builder, field_descriptor):
    """Adds a single field to the schema based on the provided attributes

    Args:
        schema_builder (SchemaBuilder): An instance of the `SchemaBuilder` class
        field_descriptor (FieldDescriptor): An instance of the `FieldDescriptor` class,
            that contains the necessary field attributes
    """
    logger.debug("Adding field '{}'".format(field_descriptor.name))

    if not schema_builder.AcceptableName(field_descriptor.name):
        raise ValueError("Invalid field name '{}'".format(field_descriptor.name))

    field_builder = _add_field(schema_builder, field_descriptor)
    set_field_documentation(field_builder, field_descriptor.__doc__)
    set_field_spec_type_id(field_builder, field_descriptor.spec_type_id)
    set_sub_schema_guid(field_builder, field_descriptor.sub_schema_guid)
    return schema_builder


def _add_field(schema_builder, field_descriptor):
    """Adds a field to the schema and returns the field builder object"""
    if field_descriptor.container_type == ES.ContainerType.Simple:
        logger.debug(
            "Creating SimpleField with value type '{}'".format(
                field_descriptor.value_type
            )
        )
        return schema_builder.AddSimpleField(
            field_descriptor.name, field_descriptor.value_type
        )

    if field_descriptor.container_type == ES.ContainerType.Array:
        logger.debug(
            "Creating ArrayField with value type '{}'".format(
                field_descriptor.value_type
            )
        )
        return schema_builder.AddArrayField(
            field_descriptor.name, field_descriptor.value_type
        )

    if field_descriptor.container_type == ES.ContainerType.Map:
        logger.debug(
            "Creating MapField with key type '{}' and value type '{}'".format(
                field_descriptor.key_type, field_descriptor.value_type
            )
        )
        return schema_builder.AddMapField(
            field_descriptor.name,
            field_descriptor.key_type,
            field_descriptor.value_type,
        )

    raise ValueError(
        "Invalid container type '{}'".format(field_descriptor.container_type)
    )


def set_field_documentation(field_builder, documentation):
    """Sets the documentation for the field"""
    field_builder.SetDocumentation(documentation)
    logger.debug("Field documentation set to: {}".format(documentation))


def set_field_spec_type_id(field_builder, spec_type_id):
    """Sets the spec type for the field"""
    if field_builder.NeedsUnits():
        field_builder.SetSpec(spec_type_id or DB.SpecTypeId.Number)
        logger.debug(
            "Field spec type set to: {}".format(spec_type_id, DB.SpecTypeId.Number)
        )


def set_sub_schema_guid(field_builder, sub_schema_guid):
    """Sets the sub schema guid for the field""" ""
    if field_builder.NeedsSubSchemaGUID():
        field_builder.SetSubSchemaGUID(sub_schema_guid)
        logger.debug("Field sub schema guid set to: {}".format(sub_schema_guid))
