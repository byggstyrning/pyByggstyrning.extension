# pyright: reportMissingImports = false
import System
from pyrevit import DB, script

ES = DB.ExtensibleStorage
logger = script.get_logger()


def initialize_schema(guid, schema_name, documentation, **kwargs):
    """Initializes the schema with the provided information"""

    schema_builder = _create_schema_builder(guid)
    _set_schema_name(schema_builder, schema_name)
    _set_schema_documentation(schema_builder, documentation)
    _set_access_levels(schema_builder, kwargs)
    _set_optional_properties(schema_builder, kwargs)
    return schema_builder


def _create_schema_builder(guid):
    """Creates and returns a SchemaBuilder object with a validated GUID"""
    if not ES.SchemaBuilder.GUIDIsValid(guid):
        raise ValueError("Invalid schema GUID '{}'".format(guid))

    logger.debug("SchemaBuilder created with GUID: {}".format(guid))
    return ES.SchemaBuilder(guid)


def _set_schema_name(schema_builder, schema_name):
    """Sets the schema name after validating it"""
    if not schema_builder.AcceptableName(schema_name):
        raise ValueError("Invalid schema name '{}'".format(schema_name))

    schema_builder.SetSchemaName(schema_name)
    logger.debug("Schema name set to: {}".format(schema_name))


def _set_schema_documentation(schema_builder, documentation):
    """Sets the schema documentation"""
    schema_builder.SetDocumentation(documentation)
    logger.debug("Schema documentation set to: '{}'".format(documentation))


def _set_access_levels(schema_builder, kwargs):
    """Sets the read and write access levels for the schema"""
    read_access = kwargs.get("read_access", ES.AccessLevel.Public)
    if not isinstance(read_access, ES.AccessLevel):
        raise ValueError("Invalid read access level '{}'".format(read_access))

    schema_builder.SetReadAccessLevel(read_access)
    logger.debug("Read access level set to: {}".format(read_access))

    write_access = kwargs.get("write_access", ES.AccessLevel.Public)
    if not isinstance(write_access, ES.AccessLevel):
        raise ValueError("Invalid write access level '{}'".format(write_access))

    schema_builder.SetWriteAccessLevel(write_access)
    logger.debug("Write access level set to: {}".format(write_access))


def _set_optional_properties(schema_builder, kwargs):
    """Sets optional properties like Vendor ID and Application GUID"""
    vendor_id = kwargs.get("vendor_id", None)
    if vendor_id is not None:
        if not ES.SchemaBuilder.VendorIdIsValid(vendor_id):
            raise ValueError("Invalid Vendor ID '{}'".format(vendor_id))

        schema_builder.SetVendorId(vendor_id)
        logger.debug("Vendor ID set to: {}".format(vendor_id))

    app_guid = kwargs.get("app_guid", None)
    if app_guid is not None:
        if isinstance(app_guid, str):
            app_guid = System.Guid(app_guid)

        if not ES.SchemaBuilder.GUIDIsValid(app_guid):
            raise ValueError("Invalid Application GUID '{}' ".format(app_guid))

        schema_builder.SetApplicationGuid(app_guid)
        logger.debug("Application GUID set to: {}".format(app_guid))
