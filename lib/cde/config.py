# -*- coding: utf-8 -*-
"""Configuration for the CDE integration.

Holds the CDE backend base URL, the REST/GraphQL endpoint builders, the
per-user token cache location and the loopback redirect settings used by the
interactive OIDC login. The base URL can be overridden per user through the
pyRevit user config (section ``CDE``) so QA / staging backends can be targeted
without code changes.
"""
import os

from pyrevit import script

logger = script.get_logger()

# --- Defaults -------------------------------------------------------------

DEFAULT_BASE_URL = "https://nobelhub-api.byggstyrning.se"

# pyRevit user-config section / keys
CONFIG_SECTION = "CDE"
CONFIG_KEY_BASE_URL = "baseUrl"

# Loopback redirect used to capture the OIDC ``code``. Must match the URI
# whitelisted on the CDE backend.
LOOPBACK_HOST = "localhost"
LOOPBACK_PORT = 48800
LOOPBACK_CALLBACK_PATH = "/callback"
LOOPBACK_REDIRECT_URI = "http://{}:{}{}".format(
    LOOPBACK_HOST, LOOPBACK_PORT, LOOPBACK_CALLBACK_PATH)

# How many seconds to wait for the user to complete the browser login.
LOGIN_TIMEOUT_SECONDS = 300

# Treat a token as expired this many seconds before its real expiry, to avoid
# racing the backend clock.
TOKEN_EXPIRY_SKEW_SECONDS = 60

# Cloudflare on nobelhub-api blocks default Python-urllib signatures (Error 1010).
# Use a normal browser UA plus a product token the API owner can allowlist.
HTTP_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 "
    "pyByggstyrning-pyRevit/1.0"
)


def apply_http_headers(req, extra=None):
    """Apply default headers expected by the CDE API / Cloudflare edge."""
    req.add_header("User-Agent", HTTP_USER_AGENT)
    req.add_header("Accept", "application/json")
    for key, value in (extra or {}).items():
        req.add_header(key, value)


# --- Per-user base URL ----------------------------------------------------

def _get_user_config():
    try:
        from pyrevit.userconfig import user_config
        return user_config
    except Exception as ex:
        logger.debug("CDE: user_config unavailable: {}".format(ex))
        return None


def get_base_url():
    """Return the configured CDE base URL (user override or default)."""
    cfg = _get_user_config()
    if cfg is not None and hasattr(cfg, CONFIG_SECTION):
        try:
            section = getattr(cfg, CONFIG_SECTION)
            value = section.get_option(CONFIG_KEY_BASE_URL, default_value=None)
            if value:
                return value.rstrip("/")
        except Exception as ex:
            logger.debug("CDE: failed reading base url override: {}".format(ex))
    return DEFAULT_BASE_URL.rstrip("/")


def set_base_url(base_url):
    """Persist a per-user base URL override."""
    cfg = _get_user_config()
    if cfg is None:
        return False
    if not hasattr(cfg, CONFIG_SECTION):
        cfg.add_section(CONFIG_SECTION)
    section = getattr(cfg, CONFIG_SECTION)
    section.set_option(CONFIG_KEY_BASE_URL, (base_url or "").rstrip("/"))
    cfg.save_changes()
    return True


# --- Token cache ----------------------------------------------------------

def get_token_dir():
    """Directory where the CDE token cache lives (``%APPDATA%/pyBS``)."""
    appdata = os.getenv("APPDATA") or os.path.expanduser("~")
    return os.path.join(appdata, "pyBS")


def get_token_file():
    """Full path to the cached CDE token json."""
    return os.path.join(get_token_dir(), "cde_token.json")


# --- Endpoint builders ----------------------------------------------------

def auth_start_url(base_url=None):
    return "{}/api/v1/auth/azure/start".format(base_url or get_base_url())


def auth_exchange_url(base_url=None):
    return "{}/api/v1/auth/azure/exchange".format(base_url or get_base_url())


def projects_url(base_url=None):
    return "{}/api/v1/projects".format(base_url or get_base_url())


def live_drops_url(project_id, base_url=None):
    return "{}/api/v1/projects/{}/live-drops".format(
        base_url or get_base_url(), project_id)


def project_api_keys_url(project_id, base_url=None):
    return "{}/api/v1/projects/{}/api-keys".format(
        base_url or get_base_url(), project_id)


def graphql_url(base_url=None):
    return "{}/api/v2/graphql".format(base_url or get_base_url())


def graph_analysis_url(base_url=None):
    return "{}/api/v2/graph/analysis".format(base_url or get_base_url())


def graph_status_url(base_url=None):
    return "{}/api/v2/graph/status".format(base_url or get_base_url())


def graph_overview_url(base_url=None):
    return "{}/api/v2/graph/overview".format(base_url or get_base_url())
