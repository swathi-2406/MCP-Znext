"""
tc_plm.py - Teamcenter / zNext PLM MCP server (34 tools from "zNext PLM MCP QA02 - Available tools v74").
One file. Settings go in .env next to this file.

  Claude Desktop      python tc_plm.py
  Gemini (via Render) python tc_plm.py --setup     once: get your keys and your Gemini address
                      python tc_plm.py --agent     keep running (or: --autostart on)
  Render              start command:  python tc_plm.py --relay

.env (next to this file):
  TC_URL=http://your-aw-server:3000      TC_USER=...      TC_PASSWORD=...
  TCPLM_PROFILE=znext                    (or: standard  for plain Teamcenter)
  RELAY_URL / AGENT_KEY / AGENT_ID       written by --setup

Write tools never write on the first call: they return a preview + confirm_token. The same call with the
token (same values, after >= 5 s, within 10 min, used once) does the write.
"""

from __future__ import annotations

import base64
import csv
import functools
import hashlib
import io
import json
import os
import re
import requests
import secrets
import sys
import threading
import time
import tomllib
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from mcp.server.mcpserver import MCPServer
from typing import Any
from typing import Any, Callable
from urllib.parse import unquote, urlsplit
from urllib.parse import urlsplit

TOOLS: dict = {}


def tool(fn):
    """Marks a function as an MCP tool (registered in main(), locally or as a relay proxy)."""
    TOOLS[fn.__name__] = fn
    return fn


# =============================================================================
# 1. SETTINGS AND PROFILES
# =============================================================================


# Built-in profiles. Lines marked CONFIRM are best guesses: fix them here (or point TCPLM_PROFILE to your own .toml).
PROFILE_TEXT = {
    '_services': r'''
# SOA service names shared by both profiles. A profile can override any of them in its own [services].
[services]
login = "Core-2011-06-Session/login"
getTCSessionInfo = "Core-2007-01-Session/getTCSessionInfo"
getProperties = "Core-2006-03-DataManagement/getProperties"
setProperties = "Core-2010-09-DataManagement/setProperties"
findSavedQueries = "Query-2010-04-SavedQuery/findSavedQueries"
executeSavedQuery = "Query-2006-03-SavedQuery/executeSavedQuery"
performSearchViewModel = "Internal-AWS2-2023-06-Finder/performSearchViewModel4"
performSearch = "Query-2012-10-Finder/performSearch"
expandGRMRelationsForPrimary = "Core-2007-09-DataManagement/expandGRMRelationsForPrimary"
expandGRMRelationsForSecondary = "Core-2007-09-DataManagement/expandGRMRelationsForSecondary"
createRelations = "Core-2006-03-DataManagement/createRelations"
deleteRelations = "Core-2006-03-DataManagement/deleteRelations"
whereUsed = "Core-2007-01-DataManagement/whereUsed"
createBOMWindows = "Cad-2007-01-StructureManagement/createBOMWindows"
expandPSOneLevel = "Cad-2007-01-StructureManagement/expandPSOneLevel"
closeBOMWindows = "Cad-2007-01-StructureManagement/closeBOMWindows"
saveBOMWindows = "Cad-2008-06-StructureManagement/saveBOMWindows"
addOrUpdateChildrenToParentLine = "Bom-2008-06-StructureManagement/addOrUpdateChildrenToParentLine"
removeChildrenFromParentLine = "Bom-2008-06-StructureManagement/removeChildrenFromParentLine"
getDeclarativeStyleSheets = "Internal-AWS2-2016-04-DataManagement/getDeclarativeStyleSheets"
getAttachedPropDescs2 = "Core-2011-06-PropDescriptor/getAttachedPropDescs2"
getInitialLOVValues = "Core-2013-05-LOV/getInitialLOVValues"
generateNextValuesForProperties = "Core-2013-05-DataManagement/generateNextValuesForProperties"
createObjects = "Core-2008-06-DataManagement/createObjects"
createDatasets = "Core-2010-04-DataManagement/createDatasets"
commitDatasetFiles = "Core-2006-03-FileManagement/commitDatasetFiles"
getWorkflowTemplates = "Workflow-2013-05-Workflow/getWorkflowTemplates"
createInstance = "Workflow-2008-06-Workflow/createInstance"
performAction3 = "Workflow-2014-06-Workflow/performAction3"
addParticipants = "Participant-2018-11-Participant/addParticipants"

# File extension -> [dataset type, named reference] used by create_attachment.
[datasets]
pdf = ["PDF", "PDF_Reference"]
doc = ["MSWord", "word"]
docx = ["MSWordX", "word"]
xls = ["MSExcel", "excel"]
xlsx = ["MSExcelX", "excel"]
ppt = ["MSPowerPoint", "powerpoint"]
pptx = ["MSPowerPointX", "powerpoint"]
png = ["Image", "Image"]
jpg = ["Image", "Image"]
jpeg = ["Image", "Image"]
gif = ["Image", "Image"]
bmp = ["Image", "Image"]
tif = ["Image", "Image"]
txt = ["Text", "Text"]
csv = ["Text", "Text"]
zip = ["Zip", "ZIPFILE"]
"*" = ["MISC", "MISC_File"]

[limits]
inline_max_mb = 5
upload_max_mb = 25
''',
    'znext': r'''
# Profile for zNext (customised Teamcenter, QA02). Names follow the zNext tool sheet (v74).
# Anything marked CONFIRM is my best reading of the sheet; check it on QA02 once.

[profile]
name = "znext"
label = "zNext"
id_label = "GPN"
object_link = "{origin}/#/com.siemens.splm.clientfx.tcui.xrt.showObject?uid={uid}"

[access]
exec_users = []

[queries]
item = "Item..."
item_id_entry = "Item ID"
item_name_entry = "Name"
item_type_entry = "Type"
general = "General..."
general_name_entry = "Name"

[types]
part = "G4B_Part"
part_types = ["G4B_Part", "G4B_Capacitor", "Item"]
document = "G4B_Document"                  # CONFIRM
mpn = "G4B_MfgPart"
mpn_types = ["G4B_MfgPart"]
manufacturer = "G4B_Manufacturer"
plant = "G4B_MfgPlantCode"
product_line = "G4B_ProductLine"
mfg_test_types = ["G4B_MfgTestRevision"]   # CONFIRM

[change_types]
eco = "G4B_ECO"
mco = "G4B_MfgChangeOrder"                 # CONFIRM
dev = "G4B_DevReq"
prq = "G4B_PostRampQual"
pnr = "G4B_PNR"
cr = "G4B_CR"
mcr = "G4B_MCR"                            # CONFIRM

[change_prefixes]
"ECO-" = "eco"
"MCO-" = "mco"
"DEV-" = "dev"
"PRQ-" = "prq"
"PNR-" = "pnr"
"CR-" = "cr"
"MCR-" = "mcr"

[props]
id = "item_id"
rev = "item_revision_id"
name = "object_name"
desc = "object_desc"
type = "object_type"
status = "release_status_list"
released = "date_released"
owner = "owning_user"
group = "owning_group"
modified = "last_mod_date"
created = "creation_date"
uom = "uom_tag"
checked_out = "checked_out"
attributes = ["object_name", "object_desc", "object_type", "release_status_list", "date_released", "owning_user",
              "last_mod_date", "g4b_category", "g4b_sourcing", "g4b_make_buy", "g4b_rohs_status", "g4b_compliance",
              "g4b_lifecycle", "g4b_product_lines", "g4b_projects", "g4b_custom_off_shelf"]
overview = ["object_name", "object_desc", "object_type", "release_status_list", "g4b_category", "g4b_make_buy",
            "g4b_rohs_status", "g4b_lifecycle", "g4b_custom_off_shelf"]
custom_off_shelf = "g4b_custom_off_shelf"
custom_value = "Custom"
off_shelf_value = "Off-the-Shelf"

[relations]
aml = "G4B_GPN_MPNRelation"
vendor_docs = "G4B_MPN_VendorDocRelation"
attachments = ["IMAN_specification", "IMAN_reference", "IMAN_Rendering", "TC_Attaches", "IMAN_manifestation"]
attach_default = "IMAN_specification"
plants = "LocationInCompany"
impacted = "CMHasImpactedItem"
solution = "CMHasSolutionItem"
problem = "CMHasProblemItem"
reference = "CMReferences"
participants = "HasParticipant"
aml_secondary = "item"                   # MPN side of the AML relation: "item" or "revision"

[aml_fields]                               # CONFIRM property names on G4B_GPN_MPNRelation
preferred_status = "g4b_preferred_status"
qual_status = "g4b_qual_status"
allocation = "g4b_allocation"
reference_notes = "g4b_reference_notes"
target_milestone = "g4b_target_milestone"
build_phase = "g4b_build_phase"
vendor_documents = "g4b_vendor_documents"
plant_codes = "g4b_plant_codes"

[mpn_fields]
manufacturer = "g4b_manufacturer"          # CONFIRM
lifecycle = "g4b_lifecycle"
rohs = "g4b_rohs_status"

[bom_fields]                               # CONFIRM the g4b_* occurrence properties
find_no = "bl_sequence_no"
gpn = "bl_item_item_id"
revision = "bl_rev_item_revision_id"
description = "bl_rev_object_desc"
quantity = "bl_quantity"
uom = "bl_uom"
ref_designators = "bl_ref_designator"
notes = "bl_occ_notes"
function = "bl_occ_g4b_function"
primary_alternate = "bl_occ_g4b_primary_alternate"
reference = "bl_occ_g4b_reference"
fatp_input = "bl_occ_g4b_fatp_input"
status = "bl_rev_release_status_list"

[change_fields]
Status = "release_status_list"
Maturity = "CMMaturity"
Disposition = "CMDisposition"
"Change Type" = "g4b_change_type"
"Reason Code" = "g4b_reason_code"
Category = "g4b_category"
"Builds Impacted" = "g4b_build_impacted"
"Functional Teams" = "g4b_functional_teams"
"Product Lines" = "g4b_product_lines"
"Description of Change" = "g4b_desc_of_change"
Requestor = "fnd0Requestor"
Analyst = "fnd0ChangeAnalyst"

[create.eco]                               # zNext: ECO id is the name; title also saved as description
type = "G4B_ECO"
values_on = "revision"                     # CONFIRM: fields that the type does not declare go on the first revision
name_from = "_generated_id"

[create.eco.map]
title = "object_desc"
change_type = "g4b_change_type"
description_of_change = "g4b_desc_of_change"
category = "g4b_category"
reason_code = "g4b_reason_code"
product_lines = "g4b_product_lines"
build_impacted = "g4b_build_impacted"
functional_teams = "g4b_functional_teams"

[create.cr]
type = "G4B_CR"
values_on = "revision"                     # CONFIRM: fields that the type does not declare go on the first revision
name_from = "_generated_id"

[create.cr.map]
synopsis = "object_desc"
description = "g4b_description"
change_details = "g4b_change_details"
change_category = "g4b_change_category"
reason_code = "g4b_reason_code"
product_lines = "g4b_product_lines"
functional_team = "g4b_functional_team"

[create.pnr]
type = "G4B_PNR"
values_on = "revision"                     # CONFIRM: fields that the type does not declare go on the first revision
name_from = "title"

[create.pnr.map]
priority = "g4b_priority"
schematic_symbol_required = "g4b_schematic_symbol_required"
footprint_required = "g4b_footprint_required"
copy_from_existing_part = "g4b_copy_from_part"
comments = "g4b_comments"

[create.prq]
type = "G4B_PostRampQual"
values_on = "revision"                     # CONFIRM: fields that the type does not declare go on the first revision
name_from = "_generated_id"

[create.prq.map]
title = "g4b_title"
description = "object_desc"
product_area = "g4b_product_area"
prq_type = "g4b_prq_type"
description_of_change = "g4b_desc_of_change"

[create.manufacturer]
type = "G4B_Manufacturer"
name_from = "name"

[create.manufacturer.map]
manufacturer_type = "g4b_manufacturer_type"
description = "object_desc"
parent_company = "g4b_parent_company"

[create.product_line]
type = "G4B_ProductLine"
name_from = "name"
product_code_max = 6

[create.product_line.map]
product_code = "g4b_product_code"
pdp_build = "g4b_pdp_build"
build_phase = "g4b_build_phase"
product_area = "g4b_product_area"
product_categories = "g4b_product_categories"

[create.product_line.cascade]
NPI = ["Proto", "EVT", "DVT", "PVT", "MP"]
Sustaining = ["N/A"]

[create.mpn]
type = "G4B_MfgPart"
name_from = "mpn_id"

[create.mpn.map]
description = "object_desc"
manufacturer = "g4b_manufacturer"

[workflow]
perform_signoff_action = "SOA_EPM_perform_action"
approve = "SOA_EPM_approve"
reject = "SOA_EPM_reject"
complete_action = "SOA_EPM_complete_action"
complete_value = "SOA_EPM_completed"
yes_value = "SOA_EPM_true"
no_value = "SOA_EPM_false"
comment_tag = "[via zNext MCP]"

[participants]                             # CONFIRM against znext_model.yaml
eco = { Analyst = "ChangeSpecialist1", Approver = "G4B_ECOApprover" }
cr = { Analyst = "Analyst", Approver = "G4B_ECOApprover" }
dev = { Approver = "G4B_ECOApprover" }
prq = { Owner = "G4B_PRQOwner", DRI = "G4B_PRQDRI", "OK2 Approver" = "G4B_OK2Approver" }
pnr = { Requestor = "Requestor" }

[rules]
nonrev_change_types = ["Non-Rev"]
change_type_prop = "g4b_change_type"
working_status_values = ["Working", "Temporary"]
obsolete_status_values = ["Obsolete"]
bom_allowed_parent_child = []              # fill from znext_model.yaml rules.bom
open_pnr_blocks_submit = true              # 919216
[rules.mandatory]                          # CONFIRM: fields zNext requires before Submit
eco = ["g4b_change_type", "g4b_reason_code", "g4b_desc_of_change", "fnd0ChangeAnalyst"]
cr = ["g4b_change_details", "fnd0ChangeAnalyst"]
''',
    'standard': r'''
# Profile for a standard (out-of-the-box) Teamcenter, e.g. our dev server inhyd0015dc.
# Objects that only exist in zNext (PNR, PRQ, product line, MPN relation fields) are left
# empty here: the tools that need them answer "not available on this Teamcenter".

[profile]
name = "standard"
label = "Teamcenter"
id_label = "Item ID"                      # what users call the part number
# Link to an object in Active Workspace; {uid} is replaced.
object_link = "{origin}/#/com.siemens.splm.clientfx.tcui.xrt.showObject?uid={uid}"

[access]
exec_users = []                           # users with the "exec" tier (read across all data)

[queries]
item = "Item..."
item_id_entry = "Item ID"
item_name_entry = "Name"
item_type_entry = "Type"
general = "General..."
general_name_entry = "Name"

[types]
part = "Item"                             # default type for create_item
part_types = ["Item", "Part", "Design"]   # types treated as parts / components
document = "Document"
mpn = "ManufacturerPart"
mpn_types = ["ManufacturerPart", "CommercialPart"]
manufacturer = "Vendor"
plant = ""
product_line = ""
mfg_test_types = []

[change_types]                             # change kind -> Teamcenter item type ("" = not available)
eco = "ChangeNotice"
cr = "ChangeRequest"
mco = ""
dev = ""
prq = ""
pnr = ""
mcr = ""
pr = "ProblemReport"

[change_prefixes]                          # id prefix -> change kind (used to recognise ids)
"ECN-" = "eco"
"CN-" = "eco"
"CR-" = "cr"
"PR-" = "pr"

[props]
id = "item_id"
rev = "item_revision_id"
name = "object_name"
desc = "object_desc"
type = "object_type"
status = "release_status_list"
released = "date_released"
owner = "owning_user"
group = "owning_group"
modified = "last_mod_date"
created = "creation_date"
uom = "uom_tag"
checked_out = "checked_out"
attributes = ["object_name", "object_desc", "object_type", "release_status_list", "date_released",
              "owning_user", "owning_group", "creation_date", "last_mod_date"]
overview = ["object_name", "object_desc", "object_type", "release_status_list", "owning_user", "last_mod_date"]
custom_off_shelf = ""                      # zNext: g4b_custom_off_shelf
custom_value = "Custom"
off_shelf_value = "Off-the-Shelf"

[relations]
aml = "VMRepresents"
vendor_docs = ""
attachments = ["IMAN_specification", "IMAN_reference", "IMAN_Rendering", "TC_Attaches", "IMAN_manifestation"]
attach_default = "IMAN_specification"
plants = ""
impacted = "CMHasImpactedItem"
solution = "CMHasSolutionItem"
problem = "CMHasProblemItem"
reference = "CMReferences"
participants = "HasParticipant"
aml_secondary = "item"                   # MPN side of the AML relation: "item" or "revision"

[aml_fields]                               # field -> property on the AML relation ("" = not available)
preferred_status = ""
qual_status = ""
allocation = ""
reference_notes = ""
target_milestone = ""
build_phase = ""
vendor_documents = ""
plant_codes = ""

[mpn_fields]
manufacturer = "vm0manufacturer"
lifecycle = ""
rohs = ""

[bom_fields]                               # field -> BOM line property ("" = not available)
find_no = "bl_sequence_no"
gpn = "bl_item_item_id"
revision = "bl_rev_item_revision_id"
description = "bl_rev_object_name"
quantity = "bl_quantity"
uom = "bl_uom"
ref_designators = "bl_ref_designator"
notes = "bl_occ_notes"
function = ""
primary_alternate = ""
reference = ""
fatp_input = ""
status = "bl_rev_release_status_list"

[change_fields]                            # label -> property on the change revision
Status = "release_status_list"
Maturity = "CMMaturity"
Disposition = "CMDisposition"
Closure = "CMClosure"
Reason = "CMReason"
Synopsis = "object_desc"
Requestor = "fnd0Requestor"
Analyst = "fnd0ChangeAnalyst"

[create.eco]
type = "ChangeNotice"
values_on = "revision"                     # CONFIRM: fields that the type does not declare go on the first revision
name_from = "title"                        # object_name = title

[create.eco.map]
description = "object_desc"
description_of_change = "object_desc"

[create.cr]
type = "ChangeRequest"
values_on = "revision"                     # CONFIRM: fields that the type does not declare go on the first revision
name_from = "synopsis"

[create.cr.map]
description = "object_desc"
change_details = "object_desc"

[create.manufacturer]
type = "Vendor"
name_from = "name"

[create.manufacturer.map]
description = "object_desc"

[create.mpn]
type = "ManufacturerPart"
name_from = "mpn_id"

[create.mpn.map]
description = "object_desc"

[workflow]
perform_signoff_action = "SOA_EPM_perform_action"
approve = "SOA_EPM_approve"
reject = "SOA_EPM_reject"
complete_action = "SOA_EPM_complete_action"
complete_value = "SOA_EPM_completed"
yes_value = "SOA_EPM_true"
no_value = "SOA_EPM_false"
comment_tag = "[via MCP]"

[participants]                              # change kind -> {user label: participant type}
eco = { Analyst = "ChangeSpecialist1", Approver = "ProposedReviewer" }
cr = { Analyst = "Analyst", Approver = "ProposedReviewer" }

[rules]
nonrev_change_types = []
change_type_prop = ""
mandatory = {}                              # change kind -> [properties required before submit]
working_status_values = []
obsolete_status_values = ["Obsolete"]
bom_allowed_parent_child = []               # [["ParentType","ChildType"], ...]; empty = no check
''',
}
HERE = os.path.dirname(os.path.abspath(__file__))
APP_DIR = HERE


def load_settings() -> dict:
    """Env vars, then .env next to this file (env vars win)."""
    vals: dict[str, str] = {}
    path = os.environ.get("TCPLM_SETTINGS") or os.path.join(APP_DIR, ".env")
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, v = line.split("=", 1)
                    vals[k.strip()] = v.strip().strip('"').strip("'")
    for k, v in os.environ.items():
        if k.startswith("TC_") or k.startswith("TCPLM_"):
            vals[k] = v
    return vals


def _merge(a: dict, b: dict) -> dict:
    out = dict(a)
    for k, v in b.items():
        out[k] = _merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def load_profile(name_or_path: str) -> dict:
    """'znext', 'standard' (built in), or a path to your own .toml (merged over the built-in services)."""
    base = tomllib.loads(PROFILE_TEXT["_services"])
    if name_or_path in PROFILE_TEXT and not name_or_path.startswith("_"):
        return _merge(base, tomllib.loads(PROFILE_TEXT[name_or_path]))
    path = name_or_path if os.path.isabs(name_or_path) else os.path.join(APP_DIR, name_or_path)
    if not os.path.exists(path):
        raise SystemExit(f"TCPLM_PROFILE={name_or_path!r}: use znext, standard, or a path to a .toml file.")
    with open(path, "rb") as f:
        return _merge(base, tomllib.load(f))


class Config:
    def __init__(self, settings: dict | None = None):
        self.settings = settings if settings is not None else load_settings()
        self.p = load_profile(self.settings.get("TCPLM_PROFILE", "znext"))
        self.label = self.p["profile"]["label"]
        self.id_label = self.p["profile"]["id_label"]
        u = urlsplit(self.settings.get("TC_URL", ""))
        self.origin = f"{u.scheme}://{u.netloc}" if u.netloc else ""

    def __getitem__(self, k):
        return self.p[k]

    def get(self, section: str, key: str, default=None):
        return self.p.get(section, {}).get(key, default)

    def link(self, uid: str) -> str:
        if not uid or not self.origin:
            return ""
        return self.p["profile"]["object_link"].format(origin=self.origin, uid=uid)

    def change_kind(self, change_id: str) -> str:
        cid = change_id.upper()
        for prefix, kind in self.p.get("change_prefixes", {}).items():
            if cid.startswith(prefix.upper()):
                return kind
        return ""

    def change_type(self, kind: str) -> str:
        return self.p.get("change_types", {}).get(kind, "")

    def all_change_types(self) -> list[str]:
        return [t for t in self.p.get("change_types", {}).values() if t]


# =============================================================================
# 2. PREVIEW -> CONFIRM TOKENS
# =============================================================================
TTL = 600
MIN_DELAY = 5


def _canon(args: dict) -> str:
    return json.dumps({k: v for k, v in args.items() if k != "confirm_token"}, sort_keys=True, default=str)


class ConfirmError(Exception):
    pass


class ConfirmStore:
    def __init__(self, ttl: int = TTL, min_delay: int = MIN_DELAY):
        self.ttl, self.min_delay = ttl, min_delay
        self._t: dict[str, dict] = {}
        self._lock = threading.Lock()

    def issue(self, user: str, tool: str, args: dict, state: str = "") -> str:
        tok = secrets.token_urlsafe(18)
        with self._lock:
            now = time.time()
            self._t = {k: v for k, v in self._t.items() if v["exp"] > now}
            self._t[tok] = {"user": user, "tool": tool, "args": hashlib.sha256(_canon(args).encode()).hexdigest(),
                            "state": state, "issued": now, "exp": now + self.ttl}
        return tok

    def redeem(self, token: str, user: str, tool: str, args: dict) -> dict:
        """Validates and consumes the token. Returns the stored record (with 'state')."""
        with self._lock:
            rec = self._t.get(token)
            now = time.time()
            if not rec:
                raise ConfirmError("Unknown or already used confirm_token. Run the preview again.")
            if rec["exp"] < now:
                self._t.pop(token, None)
                raise ConfirmError("The confirm_token expired (10 minutes). Run the preview again.")
            if rec["user"] != user or rec["tool"] != tool:
                raise ConfirmError("This confirm_token belongs to another user or tool.")
            if rec["args"] != hashlib.sha256(_canon(args).encode()).hexdigest():
                raise ConfirmError("The values differ from the preview. Run the preview again with the new values.")
            if now - rec["issued"] < self.min_delay:
                raise ConfirmError("Confirmed too quickly. Show the preview to the user and wait for their explicit yes.")
            self._t.pop(token, None)
            return rec


def fingerprint(obj) -> str:
    return hashlib.sha256(json.dumps(obj, sort_keys=True, default=str).encode()).hexdigest()[:16]


# =============================================================================
# 3. TEAMCENTER CONNECTION (JSON SOA)
# =============================================================================
NULL = {"uid": "AAAAAAAAAAAAAA", "type": "unknownType"}


class TCError(RuntimeError):
    """Error from Teamcenter or from talking to it."""

    def __init__(self, message: str, code: str = "TC_ERROR"):
        super().__init__(message)
        self.code = code


def _permission(msg: str) -> bool:
    m = msg.lower()
    return any(w in m for w in ("access", "permission", "not authorized", "privilege", "denied"))


class TCClient:
    def __init__(self, settings: dict, services: dict):
        self.base_url = settings.get("TC_URL", "").rstrip("/")
        self.user = settings.get("TC_USER", "")
        self.password = settings.get("TC_PASSWORD", "")
        self.group = settings.get("TC_GROUP", "")
        self.role = settings.get("TC_ROLE", "")
        self.verify_ssl = str(settings.get("TC_VERIFY_SSL", "true")).lower() != "false"
        self.csrf_header = settings.get("TC_CSRF_HEADER", "")
        self.svc = services
        self.session = requests.Session()
        self.logged_in = False
        self._lock = threading.RLock()

    # ------------------------------------------------------------------ transport

    def _origin(self) -> str:
        u = urlsplit(self.base_url)
        return f"{u.scheme}://{u.netloc}"

    def _apply_csrf(self) -> bool:
        cookies = {c.name: c.value for c in self.session.cookies if c.name != "_csrf"}
        token = next((cookies[n] for n in ("XSRF-TOKEN", "CSRF-TOKEN", "csrftoken") if n in cookies), None)
        if token is None:
            token = next((v for k, v in cookies.items() if "xsrf" in k.lower() or "csrf" in k.lower()), None)
        if not token:
            return False
        for n in ([self.csrf_header] if self.csrf_header else ["X-XSRF-TOKEN", "X-CSRF-TOKEN"]):
            self.session.headers[n] = unquote(token)
        return True

    def _prime_csrf(self) -> None:
        for url in (self._origin() + "/", self._origin() + "/tc/", self.base_url):
            try:
                self.session.get(url, verify=self.verify_ssl, timeout=30)
            except requests.RequestException:
                continue
            if self._apply_csrf():
                return

    def _post(self, service: str, body: dict, retry_csrf: bool = True, timeout: float = 90) -> dict:
        if not self.base_url:
            raise TCError("TC_URL is not set.", "NOT_CONFIGURED")
        self.session.headers.setdefault("Origin", self._origin())
        self.session.headers.setdefault("Referer", self._origin() + "/")
        header = {"state": {"clientVersion": "10000.1.2", "clientID": "tcplm-mcp", "stateless": True,
                            "unloadObjects": True, "enableServerStateHeaders": True, "formatProperties": True},
                  "policy": {}}
        r = self.session.post(f"{self.base_url}/JsonRestServices/{service}", json={"header": header, "body": body},
                              verify=self.verify_ssl, timeout=timeout)
        self._apply_csrf()
        if r.status_code == 403 and "csrf" in r.text.lower() and retry_csrf:
            self._prime_csrf()
            return self._post(service, body, False, timeout)
        if r.status_code != 200:
            raise TCError(f"{service}: HTTP {r.status_code} {r.text[:300]}")
        data = r.json() if r.text else {}
        if isinstance(data, dict) and str(data.get(".QName", "")).endswith("Exception"):
            msg = "; ".join(m.get("message", "") for m in data.get("messages", [])) or str(data)[:300]
            raise TCError(f"{service}: {msg}", "PERMISSION_DENIED" if _permission(msg) else "TC_ERROR")
        return data

    def login(self) -> None:
        self._prime_csrf()
        self._post(self.svc["login"], {"credentials": {
            "user": self.user, "password": self.password, "group": self.group, "role": self.role,
            "locale": "", "descrimator": "tcplm-mcp"}})
        self.logged_in = True

    def call(self, key_or_service: str, body: dict, timeout: float = 90) -> dict:
        """Call a service by its [services] key (or a literal service name)."""
        service = self.svc.get(key_or_service, key_or_service)
        with self._lock:
            if not self.logged_in:
                self.login()
            try:
                return self._post(service, body, timeout=timeout)
            except TCError as e:
                t = str(e).lower()
                if any(w in t for w in ("session", "credential", "login", "not logged")):
                    self.logged_in = False
                    self.login()
                    return self._post(service, body, timeout=timeout)
                raise

    # ------------------------------------------------------------------ reply helpers

    @staticmethod
    def sd(resp: dict) -> dict:
        return resp.get("ServiceData") or {}

    @classmethod
    def model_objects(cls, resp: dict) -> dict:
        return cls.sd(resp).get("modelObjects") or resp.get("modelObjects") or {}

    @classmethod
    def errors(cls, resp: dict) -> list[str]:
        out = []
        for pe in cls.sd(resp).get("partialErrors") or []:
            for ev in pe.get("errorValues") or []:
                code = ev.get("code")
                out.append(f"{ev.get('message', '')}" + (f" ({code})" if code else ""))
        return out

    @classmethod
    def raise_errors(cls, resp: dict, what: str = "") -> None:
        errs = cls.errors(resp)
        if errs:
            msg = "; ".join(errs)
            raise TCError((what + ": " if what else "") + msg, "PERMISSION_DENIED" if _permission(msg) else "TC_ERROR")

    @staticmethod
    def ref(o: dict) -> dict:
        return {"uid": o["uid"], "type": o.get("type") or o.get("className") or "unknownType"}

    # ------------------------------------------------------------------ properties

    def get_props(self, objects: list[dict], names: list[str]) -> dict[str, dict]:
        """{uid: {"type", "className", "ui": {name: [..]}, "db": {name: [..]}}} for the objects asked, in order."""
        objects = [o for o in objects if o and o.get("uid") and o["uid"] != NULL["uid"]]
        if not objects:
            return {}
        out: dict[str, dict] = {}
        for i in range(0, len(objects), 200):
            chunk = objects[i:i + 200]
            resp = self.call("getProperties", {"objects": [self.ref(o) for o in chunk], "attributes": names})
            mos = self.model_objects(resp)
            for o in chunk:
                mo = mos.get(o["uid"])
                if not mo:
                    continue
                ui, db = {}, {}
                for n, p in (mo.get("props") or {}).items():
                    ui[n] = p.get("uiValues") if p.get("uiValues") is not None else p.get("dbValues") or []
                    db[n] = p.get("dbValues") if p.get("dbValues") is not None else []
                out[o["uid"]] = {"uid": o["uid"], "type": mo.get("type"), "className": mo.get("className"), "ui": ui, "db": db}
        return out

    @staticmethod
    def one(values) -> Any:
        if isinstance(values, list):
            if not values:
                return ""
            return values[0] if len(values) == 1 else values
        return values if values is not None else ""

    def set_props(self, obj: dict, values: dict[str, Any]) -> None:
        def as_list(v):
            return [str(x) for x in v] if isinstance(v, list) else [str(v)]
        resp = self.call("setProperties", {"info": [{"object": self.ref(obj), "timestamp": "",
                                                     "vecNameVal": [{"name": k, "values": as_list(v)} for k, v in values.items()]}],
                                           "options": []})
        self.raise_errors(resp, "setProperties")

    # ------------------------------------------------------------------ queries & search

    def find_saved_query(self, name: str) -> dict:
        resp = self.call("findSavedQueries", {"inputCriteria": [{"queryNames": [name], "queryDescs": [], "queryType": 0}]})
        qs = resp.get("savedQueries") or []
        if not qs:
            raise TCError(f'Saved query "{name}" not found.', "NOT_CONFIGURED")
        return qs[0]

    def run_saved_query(self, name: str, entries: dict[str, str], limit: int = 200) -> list[dict]:
        q = self.find_saved_query(name)
        resp = self.call("executeSavedQuery", {"query": self.ref(q), "entries": list(entries.keys()),
                                               "values": list(entries.values()), "limit": limit})
        return (resp.get("objects") or [])[:limit]

    def fulltext(self, text: str, types: list[str] | None, start: int, count: int) -> tuple[list[dict], int, dict]:
        """AWC full-text search. Returns (objects, total, type_facets)."""
        body = {
            "columnConfigInput": {"clientName": "AWClient", "clientScopeURI": "Awp0SearchResults"},
            "searchInput": {
                "providerName": "Awp0FullTextSearchProvider",
                "searchCriteria": {"searchString": text},
                "searchFilterFieldSortType": "Priority",
                "searchFilterMap6": ({"WorkspaceObject.object_type": [{"searchFilterType": "StringFilter", "stringValue": t} for t in types]}
                                     if types else {}),
                "startIndex": start, "maxToLoad": count, "maxToReturn": count,
                "attributesToInflate": [], "internalPropertyName": "", "cursor": {"startIndex": start, "endIndex": 0,
                                                                                  "startReached": False, "endReached": False},
            },
            "inflateProperties": False, "saveColumnConfigData": {}, "noServiceData": False,
        }
        resp = self.call("performSearchViewModel", body)
        objs = []
        raw = resp.get("searchResultsJSON")
        if isinstance(raw, str) and raw:
            import json
            try:
                objs = [{"uid": o.get("uid"), "type": o.get("type", "unknownType")} for o in json.loads(raw).get("objects", [])]
            except ValueError:
                objs = []
        if not objs:
            objs = [{"uid": o["uid"], "type": o.get("type", "unknownType")} for o in (resp.get("searchResults") or []) if o.get("uid")]
        facets = {}
        for cat in ((resp.get("searchFilterMap6") or {}).get("WorkspaceObject.object_type") or []):
            facets[cat.get("stringDisplayValue") or cat.get("stringValue")] = cat.get("count", 0)
        return objs, int(resp.get("totalFound") or len(objs)), facets

    # ------------------------------------------------------------------ relations

    def expand_relations(self, primaries: list[dict], relations: list[str], secondary: bool = False) -> dict[str, list[dict]]:
        """{primary uid: [{"relation": name, "other": ref, "rel": ref}]}"""
        if not primaries:
            return {}
        key = "expandGRMRelationsForSecondary" if secondary else "expandGRMRelationsForPrimary"
        body = {("secondaryObjects" if secondary else "primaryObjects"): [self.ref(p) for p in primaries],
                "pref": {"expItemRev": False, "returnRelations": True,
                         "info": [{"relationTypeName": r, "otherSideObjectTypes": []} for r in relations]}}
        resp = self.call(key, body)
        out: dict[str, list[dict]] = {p["uid"]: [] for p in primaries}
        for o in resp.get("output") or []:
            uid = (o.get("inputObject") or {}).get("uid")
            for rd in o.get("relationshipData") or []:
                for ro in rd.get("relationshipObjects") or []:
                    other = ro.get("otherSideObject")
                    if other and uid in out:
                        out[uid].append({"relation": rd.get("relationName"), "other": other, "rel": ro.get("relation")})
        return out

    def create_relation(self, rel_type: str, primary: dict, secondary: dict) -> dict:
        resp = self.call("createRelations", {"input": [{"clientId": "r", "relationType": rel_type, "primaryObject": self.ref(primary),
                                                        "secondaryObject": self.ref(secondary), "userData": NULL}]})
        self.raise_errors(resp, f"createRelations {rel_type}")
        out = resp.get("output") or [{}]
        return out[0].get("relation") or {}

    def delete_relation(self, rel_type: str, primary: dict, secondary: dict) -> None:
        resp = self.call("deleteRelations", {"input": [{"clientId": "r", "relationType": rel_type, "primaryObject": self.ref(primary),
                                                        "secondaryObject": self.ref(secondary), "userData": NULL}]})
        self.raise_errors(resp, f"deleteRelations {rel_type}")

    # ------------------------------------------------------------------ BOM

    def open_bom(self, rev: dict) -> tuple[dict, dict]:
        resp = self.call("createBOMWindows", {"info": [{
            "clientId": "w", "item": NULL, "itemRev": self.ref(rev), "bomView": NULL,
            "revRuleConfigInfo": {"clientId": "w", "revRule": NULL, "props": {
                "unitNo": -1, "date": "0001-01-01T00:00:00", "today": True, "endItem": NULL,
                "endItemRevision": NULL, "overrideFolders": []}},
            "objectForConfigure": NULL, "activeAssemblyArrangement": NULL}]})
        out = resp.get("output") or []
        if not out:
            raise TCError("Could not open a BOM window. " + "; ".join(self.errors(resp)))
        return out[0]["bomWindow"], out[0]["bomLine"]

    def close_bom(self, window: dict, save: bool = False) -> None:
        if save:
            resp = self.call("saveBOMWindows", {"bomWindows": [self.ref(window)]})
            self.raise_errors(resp, "saveBOMWindows")
        try:
            self.call("closeBOMWindows", {"bomWindows": [self.ref(window)]})
        except TCError:
            pass

    def expand_line(self, line: dict) -> list[dict]:
        """Children of one BOM line: [{"line": ref, "rev": ref|None}]"""
        resp = self.call("expandPSOneLevel", {"input": {"parentBomLines": [self.ref(line)], "excludeFilter": "None"},
                                              "pref": {"expItemRev": False, "info": []}})
        kids = []
        for o in resp.get("output") or []:
            for ch in o.get("children") or []:
                kids.append({"line": ch["bomLine"], "rev": ch.get("itemRevOfBOMLine")})
        return kids

    # ------------------------------------------------------------------ session

    def session_info(self) -> dict:
        resp = self.call("getTCSessionInfo", {})
        return resp


def call_timed(fn, *a, timeout_s: float = 45, **kw):
    """Run fn with a soft deadline (raises TCError if it took longer than timeout_s)."""
    t0 = time.time()
    out = fn(*a, **kw)
    if time.time() - t0 > timeout_s:
        raise TCError(f"Timed out after {int(timeout_s)} s.", "TIMEOUT")
    return out


# =============================================================================
# 4. SHARED HELPERS
# =============================================================================
class ToolError(Exception):
    def __init__(self, message: str, code: str = "ERROR", **extra):
        super().__init__(message)
        self.code, self.extra = code, extra


class Ctx:
    def __init__(self, cfg: Config, tc: TCClient):
        self.cfg, self.tc = cfg, tc
        self.confirm = ConfirmStore()
        self._cache: dict[str, tuple[float, Any]] = {}
        self._me: dict | None = None

    # ------------------------------------------------------------ cache
    def cached(self, key: str, ttl: float, fn: Callable[[], Any]) -> Any:
        hit = self._cache.get(key)
        if hit and hit[0] > time.time():
            return hit[1]
        val = fn()
        self._cache[key] = (time.time() + ttl, val)
        return val

    # ------------------------------------------------------------ identity
    def me(self) -> dict:
        if self._me is None:
            info = self.tc.session_info()
            user = info.get("user") or {}
            props = self.tc.get_props([user], ["user_id", "user_name", "taskinbox"]) if user.get("uid") else {}
            p = props.get(user.get("uid"), {})
            uid_str = str(self.tc.one(p.get("db", {}).get("user_id", [])) or self.tc.user)
            self._me = {"user_obj": user, "user_id": uid_str,
                        "user_name": str(self.tc.one(p.get("ui", {}).get("user_name", [])) or uid_str),
                        "group": info.get("group"), "role": info.get("role"),
                        "taskinbox": (p.get("db", {}).get("taskinbox") or [None])[0],
                        "server": info.get("serverVersion", ""), "raw": info}
        return self._me

    def user_key(self) -> str:
        return self.tc.user or "unknown"

    def tier(self) -> str:
        return "exec" if self.tc.user in (self.cfg.get("access", "exec_users") or []) else "user"

    # ------------------------------------------------------------ props
    def P(self, key: str) -> str:
        return self.cfg["props"][key]

    def ui(self, rec: dict, name: str) -> Any:
        return self.tc.one(rec.get("ui", {}).get(name, [])) if name else ""

    def db(self, rec: dict, name: str) -> list:
        return rec.get("db", {}).get(name, []) if name else []

    # ------------------------------------------------------------ lookup
    def find_items(self, item_id: str, limit: int = 25, types: list[str] | None = None) -> list[dict]:
        q = self.cfg["queries"]
        entries = {q["item_id_entry"]: item_id}
        objs = self.tc.run_saved_query(q["item"], entries, limit=limit)
        if types:
            props = self.tc.get_props(objs, [self.P("type")])
            objs = [o for o in objs if (props.get(o["uid"], {}).get("type") or "") in types
                    or self.ui(props.get(o["uid"], {}), self.P("type")) in types] or objs
        return objs

    def item(self, item_id: str, types: list[str] | None = None) -> dict:
        objs = self.find_items(item_id, limit=5, types=types)
        exact = []
        if objs:
            props = self.tc.get_props(objs, [self.P("id")])
            exact = [o for o in objs if str(self.ui(props.get(o["uid"], {}), self.P("id"))).upper() == item_id.upper()]
        if not exact:
            raise ToolError(f"{self.cfg.id_label} {item_id} not found in {self.cfg.label}.", "NOT_FOUND")
        return exact[0]

    def revisions(self, item: dict) -> list[dict]:
        p = self.tc.get_props([item], ["revision_list"]).get(item["uid"], {})
        refs = [{"uid": u, "type": "unknownType"} for u in self.db(p, "revision_list") if u]
        props = self.tc.get_props(refs, [self.P("rev"), self.P("status"), self.P("released"), self.P("name"),
                                         self.P("desc"), self.P("type"), self.P("modified")])
        out = []
        for r in refs:
            rec = props.get(r["uid"])
            if rec:
                status = rec["ui"].get(self.P("status")) or []
                out.append({"ref": {"uid": r["uid"], "type": rec.get("type") or "ItemRevision"},
                            "rev": str(self.ui(rec, self.P("rev"))), "status": [s for s in status if s],
                            "released": self.ui(rec, self.P("released")), "rec": rec})
        return out

    def pick_revision(self, revs: list[dict], revision: str, item_id: str) -> dict:
        if not revs:
            raise ToolError(f"{item_id} has no revisions you can see.", "NOT_FOUND")
        r = (revision or "latest_released").strip()
        if r.lower() == "latest":
            return revs[-1]
        if r.lower() in ("latest_released", "released"):
            working = set(self.cfg.get("rules", "working_status_values") or [])
            rel = [x for x in revs if x["status"] and not set(x["status"]) <= working]
            if not rel:
                raise ToolError(f"{item_id} has no released revision. Use revision='latest' for the newest one "
                                f"({revs[-1]['rev']}).", "NO_RELEASED_REVISION", available=[x["rev"] for x in revs])
            return rel[-1]
        for x in revs:
            if x["rev"].upper() == r.upper():
                return x
        raise ToolError(f"Revision {r} of {item_id} not found.", "NOT_FOUND", available=[x["rev"] for x in revs])

    def resolve(self, item_id: str, revision: str = "latest_released", types: list[str] | None = None) -> tuple[dict, dict]:
        """(item ref, revision dict) for an id + revision rule."""
        it = self.item(item_id, types)
        return it, self.pick_revision(self.revisions(it), revision, item_id)

    def change(self, change_id: str, revision: str = "latest") -> tuple[dict, dict, str]:
        kind = self.cfg.change_kind(change_id)
        types = [self.cfg.change_type(kind)] if kind and self.cfg.change_type(kind) else None
        it, rev = self.resolve(change_id, revision or "latest", types)
        return it, rev, kind

    def link(self, uid: str) -> str:
        return self.cfg.link(uid)

    def not_available(self, what: str) -> dict:
        return {"error": "NOT_AVAILABLE", "message": f"{what} is not available on this {self.cfg.label} "
                                                     f"(profile '{self.cfg['profile']['name']}'). Set it up in the profile file."}


def tool_errors(fn):
    """Turn exceptions into the error shapes the tool sheet describes."""
    @functools.wraps(fn)
    def wrap(*a, **kw):
        try:
            return fn(*a, **kw)
        except ConfirmError as e:
            return {"error": "CONFIRM_REFUSED", "message": str(e)}
        except ToolError as e:
            return {"error": e.code, "message": str(e), **e.extra}
        except TCError as e:
            if e.code == "PERMISSION_DENIED":
                return {"error": "PERMISSION_DENIED", "message": f"You don't have access to do this in Teamcenter: {e}"}
            return {"error": e.code, "message": str(e)}
    return wrap


def page_of(items: list, page: int, size: int, cap: int = 500) -> dict:
    size = max(1, min(size, cap))
    start = max(0, page) * size
    return {"total": len(items), "page": page, "page_size": size, "has_more": start + size < len(items),
            "items": items[start:start + size]}


# =============================================================================
# 5. MCP SERVER OBJECT
# =============================================================================
mcp = MCPServer(
    "teamcenter-plm",
    instructions=(
        "PLM tools for Teamcenter / zNext: search, parts (GPNs), BOMs, sourcing (AML/MPNs), changes (ECO, CR, ...), "
        "workflows and tasks. Write tools never change anything on the first call: they return a PREVIEW and a "
        "confirm_token. Show the preview to the user and call again with the same values and the token ONLY after "
        "the user explicitly says yes."),
)

_ctx: Ctx | None = None


def ctx() -> Ctx:
    global _ctx
    if _ctx is None:
        cfg = Config()
        _ctx = Ctx(cfg, TCClient(cfg.settings, cfg["services"]))
    return _ctx


def set_ctx(c: Ctx) -> None:
    global _ctx
    _ctx = c

TOOL_NAMES: list[str] = []


# =============================================================================
# 6. BOM / AML / CHANGE / WORKFLOW DATA
# =============================================================================
# ---------------------------------------------------------------- BOM

def bom_fields(ctx: Ctx) -> dict[str, str]:
    return {k: v for k, v in ctx.cfg["bom_fields"].items() if v}


def bom_lines(ctx: Ctx, rev: dict, depth: int = 1, max_lines: int = 5000) -> list[dict]:
    """Walks the BOM under a revision. Each line: level, parent, find_no, gpn, revision, quantity, ... + line/rev refs."""
    depth = max(1, min(depth, 5))
    fields = bom_fields(ctx)
    window, top = ctx.tc.open_bom(rev)
    out: list[dict] = []
    try:
        top_p = ctx.tc.get_props([top], [fields.get("gpn", "bl_item_item_id")]).get(top["uid"], {})
        frontier = [(top, str(ctx.ui(top_p, fields.get("gpn", ""))), 0)]
        while frontier and len(out) < max_lines:
            nxt = []
            for line, parent_id, lvl in frontier:
                kids = ctx.tc.expand_line(line)
                if not kids:
                    continue
                props = ctx.tc.get_props([k["line"] for k in kids], list(fields.values()))
                for k in kids:
                    p = props.get(k["line"]["uid"], {})
                    row = {"level": lvl + 1, "parent": parent_id}
                    for f, prop in fields.items():
                        v = ctx.ui(p, prop)
                        row[f] = v if not isinstance(v, list) else ", ".join(map(str, v))
                    row["_line"], row["_rev"] = k["line"], k.get("rev")
                    out.append(row)
                    if lvl + 1 < depth:
                        nxt.append((k["line"], str(row.get("gpn", "")), lvl + 1))
                    if len(out) >= max_lines:
                        break
            frontier = nxt
    finally:
        ctx.tc.close_bom(window)
    return out


def public(rows: list[dict]) -> list[dict]:
    return [{k: v for k, v in r.items() if not k.startswith("_")} for r in rows]


def bom_diff(a: list[dict], b: list[dict], cap: int = 500) -> dict:
    """Merges repeated parts per parent, then diffs A -> B."""
    def merge(rows):
        m: dict[tuple, dict] = {}
        for r in rows:
            key = (r.get("parent", ""), r.get("gpn", ""))
            if key in m:
                x = m[key]
                try:
                    x["quantity"] = str(float(x.get("quantity") or 0) + float(r.get("quantity") or 0)).rstrip("0").rstrip(".")
                except ValueError:
                    x["quantity"] = f"{x.get('quantity')}+{r.get('quantity')}"
                refs = sorted({s.strip() for s in f"{x.get('ref_designators','')},{r.get('ref_designators','')}".split(",") if s.strip()})
                x["ref_designators"] = ",".join(refs)
            else:
                m[key] = {k: v for k, v in r.items() if not k.startswith("_")}
        return m
    ma, mb = merge(a), merge(b)
    compare = ["quantity", "revision", "uom", "ref_designators", "find_no", "function", "primary_alternate"]
    added = [mb[k] for k in mb if k not in ma]
    removed = [ma[k] for k in ma if k not in mb]
    changed = []
    for k in ma.keys() & mb.keys():
        diffs = {f: {"from": ma[k].get(f, ""), "to": mb[k].get(f, "")} for f in compare
                 if str(ma[k].get(f, "")) != str(mb[k].get(f, ""))}
        if diffs:
            changed.append({"parent": k[0], "gpn": k[1], "changes": diffs})
    return {"added": added[:cap], "removed": removed[:cap], "changed": changed[:cap],
            "counts": {"added": len(added), "removed": len(removed), "changed": len(changed)},
            "truncated": max(len(added), len(removed), len(changed)) > cap}

# ---------------------------------------------------------------- sourcing / AML

def aml(ctx: Ctx, revs: list[dict]) -> dict[str, list[dict]]:
    """{rev uid: [{mpn, manufacturer, preferred_status, qual_status, ..., _mpn, _rel}]}"""
    rel = ctx.cfg.get("relations", "aml")
    if not rel or not revs:
        return {r["uid"]: [] for r in revs}
    exp = ctx.tc.expand_relations(revs, [rel])
    mpns = [x["other"] for v in exp.values() for x in v]
    rels = [x["rel"] for v in exp.values() for x in v if x.get("rel")]
    mf = ctx.cfg["mpn_fields"]
    mp = ctx.tc.get_props(mpns, [ctx.P("id"), ctx.P("name"), ctx.P("desc")] + [v for v in mf.values() if v])
    af = {k: v for k, v in ctx.cfg["aml_fields"].items() if v}
    rp = ctx.tc.get_props(rels, list(af.values())) if af else {}
    out = {}
    for uid, lst in exp.items():
        rows = []
        for x in lst:
            m = mp.get(x["other"]["uid"], {})
            row = {"mpn": ctx.ui(m, ctx.P("id")) or ctx.ui(m, ctx.P("name")), "description": ctx.ui(m, ctx.P("desc")),
                   "manufacturer": ctx.ui(m, mf.get("manufacturer", "")), "lifecycle": ctx.ui(m, mf.get("lifecycle", "")),
                   "rohs": ctx.ui(m, mf.get("rohs", ""))}
            r = rp.get((x.get("rel") or {}).get("uid", ""), {})
            for f, prop in af.items():
                row[f] = ctx.ui(r, prop)
            row["_mpn"], row["_rel"], row["link"] = x["other"], x.get("rel"), ctx.link(x["other"]["uid"])
            rows.append(row)
        out[uid] = rows
    return out

# ---------------------------------------------------------------- attachments

def attachments(ctx: Ctx, rev: dict, dataset_types: list[str] | None = None) -> list[dict]:
    rels = ctx.cfg.get("relations", "attachments") or []
    exp = ctx.tc.expand_relations([rev], rels).get(rev["uid"], [])
    ds = [x["other"] for x in exp]
    props = ctx.tc.get_props(ds, [ctx.P("name"), ctx.P("type"), ctx.P("modified"), ctx.P("owner"), "ref_names"])
    out = []
    for x in exp:
        p = props.get(x["other"]["uid"], {})
        dtype = str(ctx.ui(p, ctx.P("type")) or p.get("type") or "")
        if dataset_types and dtype not in dataset_types and (p.get("type") or "") not in dataset_types:
            continue
        out.append({"name": ctx.ui(p, ctx.P("name")), "type": dtype, "relation": x["relation"],
                    "modified": ctx.ui(p, ctx.P("modified")), "owner": ctx.ui(p, ctx.P("owner")),
                    "link": ctx.link(x["other"]["uid"]), "_ds": x["other"]})
    return out

# ---------------------------------------------------------------- change items

CM_RELATIONS = ("impacted", "solution", "problem", "reference")


def change_items(ctx: Ctx, change_rev: dict, cap: int = 200) -> dict[str, list[dict]]:
    names = {k: ctx.cfg.get("relations", k) for k in CM_RELATIONS if ctx.cfg.get("relations", k)}
    exp = ctx.tc.expand_relations([change_rev], list(names.values())).get(change_rev["uid"], [])
    objs = [x["other"] for x in exp]
    props = ctx.tc.get_props(objs, [ctx.P("id"), ctx.P("rev"), ctx.P("name"), ctx.P("desc"), ctx.P("status"), ctx.P("type")])
    out: dict[str, list[dict]] = {k: [] for k in names}
    back = {v: k for k, v in names.items()}
    for x in exp:
        kind = back.get(x["relation"])
        if not kind or len(out[kind]) >= cap:
            continue
        p = props.get(x["other"]["uid"], {})
        out[kind].append({"id": ctx.ui(p, ctx.P("id")), "revision": ctx.ui(p, ctx.P("rev")),
                          "name": ctx.ui(p, ctx.P("name")), "description": ctx.ui(p, ctx.P("desc")),
                          "status": ctx.ui(p, ctx.P("status")), "type": ctx.ui(p, ctx.P("type")),
                          "link": ctx.link(x["other"]["uid"]), "_obj": x["other"]})
    return out


def participants(ctx: Ctx, rev: dict) -> list[dict]:
    try:
        p = ctx.tc.get_props([rev], ["HasParticipant"]).get(rev["uid"], {})
    except TCError:
        return []
    refs = [{"uid": u, "type": "Participant"} for u in ctx.db(p, "HasParticipant") if u]
    pp = ctx.tc.get_props(refs, ["object_type", "assignee", "fnd0AssigneeUser"])
    out = []
    for r in refs:
        x = pp.get(r["uid"], {})
        out.append({"role": ctx.ui(x, "object_type") or x.get("type"),
                    "person": ctx.ui(x, "fnd0AssigneeUser") or ctx.ui(x, "assignee")})
    return out

# ---------------------------------------------------------------- workflow

TASK_PROPS = ["object_name", "object_type", "task_state", "state_value", "fnd0Assignee", "responsible_party",
              "creation_date", "fnd0StartDate", "child_tasks", "signoff_attachments", "valid_signoffs",
              "parent_process", "job_name", "root_target_attachments", "fnd0WorkflowInitiator", "fnd0Status"]
SIGNOFF_PROPS = ["group_member", "fnd0Assignee", "decision", "decision_date", "fnd0DecisionRequired", "object_name"]


def _age_days(s: str) -> float | None:
    s = (s or "").strip()
    for fmt in ("%d-%b-%Y %H:%M:%S", "%d-%b-%Y %H:%M", "%d-%b-%Y", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S"):
        try:
            d = datetime.strptime(s, fmt)
        except ValueError:
            continue
        now = datetime.now(d.tzinfo) if d.tzinfo else datetime.now()
        return round((now - d).total_seconds() / 86400, 1)
    return None


def workflows(ctx: Ctx, obj: dict) -> list[dict]:
    """Workflows on an object with their started tasks and sign-offs."""
    p = ctx.tc.get_props([obj], ["fnd0AllWorkflows", "process_stage_list", "fnd0MyWorkflowTasks"]).get(obj["uid"], {})
    job_uids = [u for u in ctx.db(p, "fnd0AllWorkflows") if u]
    stage = [{"uid": u, "type": "EPMTask"} for u in ctx.db(p, "process_stage_list") if u]
    jobs = []
    if job_uids:
        jp = ctx.tc.get_props([{"uid": u, "type": "EPMJob"} for u in job_uids], ["object_name", "root_task", "fnd0Status"])
        for u in job_uids:
            j = jp.get(u, {})
            root = (ctx.db(j, "root_task") or [None])[0]
            jobs.append({"name": ctx.ui(j, "object_name"), "status": ctx.ui(j, "fnd0Status"),
                         "root": {"uid": root, "type": "EPMTask"} if root else None})
    elif stage:
        jobs.append({"name": "", "status": "", "root": None, "stage": stage})
    out = []
    for j in jobs:
        tasks = []
        frontier = [j["root"]] if j.get("root") else list(j.get("stage", []))
        seen = set()
        while frontier:
            props = ctx.tc.get_props(frontier, TASK_PROPS)
            nxt = []
            for t in frontier:
                if t["uid"] in seen:
                    continue
                seen.add(t["uid"])
                tp = props.get(t["uid"], {})
                state = str(ctx.ui(tp, "task_state") or ctx.ui(tp, "fnd0Status"))
                if not j["name"]:
                    j["name"] = ctx.ui(tp, "job_name") or ctx.ui(tp, "parent_process")
                kids = [{"uid": u, "type": "EPMTask"} for u in ctx.db(tp, "child_tasks") if u]
                nxt.extend(kids)
                if state.lower() in ("started", "in progress", "4"):
                    sos = [{"uid": u, "type": "Signoff"} for u in (ctx.db(tp, "valid_signoffs") or ctx.db(tp, "signoff_attachments")) if u]
                    sp = ctx.tc.get_props(sos, SIGNOFF_PROPS)
                    signoffs = []
                    for s in sos:
                        x = sp.get(s["uid"], {})
                        dec = str(ctx.ui(x, "decision") or "No Decision")
                        signoffs.append({"person": ctx.ui(x, "group_member") or ctx.ui(x, "fnd0Assignee"),
                                         "decision": dec, "decided": dec.lower() not in ("no decision", "", "0"),
                                         "required": str(ctx.ui(x, "fnd0DecisionRequired") or "Required").lower() != "optional",
                                         "decision_date": ctx.ui(x, "decision_date"), "_signoff": s})
                    started = str(ctx.ui(tp, "fnd0StartDate") or ctx.ui(tp, "creation_date"))
                    tasks.append({"task": ctx.ui(tp, "object_name"), "type": ctx.ui(tp, "object_type") or tp.get("type"),
                                  "assignee": ctx.ui(tp, "fnd0Assignee") or ctx.ui(tp, "responsible_party"),
                                  "started": started, "age_days": _age_days(started), "signoffs": signoffs,
                                  "_task": t})
            frontier = [k for k in nxt if k["uid"] not in seen]
        out.append({"workflow": j["name"], "status": j["status"], "active_tasks": tasks})
    return out


def diagnose(wfs: list[dict]) -> str:
    if not wfs:
        return "No workflow has been started on this object. It needs to be submitted (submit_change)."
    lines = []
    for w in wfs:
        if not w["active_tasks"]:
            lines.append(f"Workflow '{w['workflow']}' has no active task (finished or not started).")
        for t in w["active_tasks"]:
            age = f" for {t['age_days']} days" if t.get("age_days") is not None else ""
            pending = [s for s in t["signoffs"] if not s["decided"] and s["required"]]
            if pending:
                who = ", ".join(str(s["person"]) for s in pending)
                lines.append(f"'{t['task']}' is waiting{age} on required sign-off from: {who}.")
            elif t["signoffs"]:
                lines.append(f"'{t['task']}' has all required sign-offs; it should move on (check optional ones or a condition).")
            elif not t.get("assignee"):
                lines.append(f"'{t['task']}' is active{age} but nobody is assigned. A workflow admin needs to assign it.")
            else:
                lines.append(f"'{t['task']}' is with {t['assignee']}{age}.")
    return " ".join(lines)


def inbox(ctx: Ctx, include_tracked: bool = False, cap: int = 500) -> list[dict]:
    me = ctx.me()
    if not me.get("taskinbox"):
        return []
    ib = {"uid": me["taskinbox"], "type": "TaskInbox"}
    names = ["tasks_to_perform"] + (["tasks_to_track"] if include_tracked else [])
    p = ctx.tc.get_props([ib], names).get(ib["uid"], {})
    refs = []
    for n in names:
        refs += [({"uid": u, "type": "EPMTask"}, "perform" if n == "tasks_to_perform" else "track") for u in ctx.db(p, n) if u]
    refs = refs[:cap]
    tp = ctx.tc.get_props([r for r, _ in refs], TASK_PROPS)
    targets = {u for r, _ in refs for u in ctx.db(tp.get(r["uid"], {}), "root_target_attachments")[:1] if u}
    tgt = ctx.tc.get_props([{"uid": u, "type": "unknownType"} for u in targets], [ctx.P("id"), ctx.P("rev"), ctx.P("name"), ctx.P("type")])
    out = []
    for r, mode in refs:
        x = tp.get(r["uid"], {})
        t_uid = (ctx.db(x, "root_target_attachments") or [None])[0]
        t = tgt.get(t_uid, {}) if t_uid else {}
        started = str(ctx.ui(x, "fnd0StartDate") or ctx.ui(x, "creation_date"))
        out.append({"task": ctx.ui(x, "object_name"), "task_type": ctx.ui(x, "object_type") or x.get("type"),
                    "workflow": ctx.ui(x, "job_name") or ctx.ui(x, "parent_process"), "mode": mode,
                    "target": f"{ctx.ui(t, ctx.P('id'))}" + (f"/{ctx.ui(t, ctx.P('rev'))}" if ctx.ui(t, ctx.P("rev")) else ""),
                    "target_name": ctx.ui(t, ctx.P("name")), "target_link": ctx.link(t_uid) if t_uid else "",
                    "started": started, "age_days": _age_days(started), "_task": r, "_target_uid": t_uid})
    out.sort(key=lambda x: (x["age_days"] if x["age_days"] is not None else 1e9))
    return out

# ---------------------------------------------------------------- page layout (Overview etc.)

def _find_xrt(obj: Any) -> list[str]:
    found = []
    if isinstance(obj, str) and "<rendering" in obj:
        found.append(obj)
    elif isinstance(obj, dict):
        for v in obj.values():
            found += _find_xrt(v)
    elif isinstance(obj, list):
        for v in obj:
            found += _find_xrt(v)
    return found


def layout(ctx: Ctx, obj: dict, type_name: str) -> dict:
    """{"pages": [titles], "sections": {page: [(section, [prop names])]}} from the Active Workspace style sheet.
    Cached 1 h per type. Empty dict if the service is not available."""
    def load():
        try:
            resp = ctx.tc.call("getDeclarativeStyleSheets", {
                "processEntries": [{"stylesheetType": "SUMMARY", "targetPage": "", "clientContext": {
                    "ActiveWorkspace:Location": "com.siemens.splm.clientfx.tcui.xrt.showObjectLocation",
                    "ActiveWorkspace:SubLocation": "com.siemens.splm.clientfx.tcui.xrt.showObjectSubLocation"}}],
                "input": [{"businessObject": ctx.tc.ref(obj), "businessObjectType": type_name, "styleSheetType": "SUMMARY",
                           "styleSheetLastModDate": "", "targetPage": "", "clientContext": {}}]})
        except TCError:
            return {}
        xrts = _find_xrt(resp)
        if not xrts:
            return {}
        try:
            root = ET.fromstring(xrts[0])
        except ET.ParseError:
            return {}
        pages, sections = [], {}
        for page in root.iter("page"):
            title = page.get("title") or page.get("titleKey") or "Page"
            pages.append(title)
            secs = []
            for sec in page.iter("section"):
                names = [p.get("name") for p in sec.iter("property") if p.get("name")]
                if names:
                    secs.append((sec.get("title") or sec.get("titleKey") or "", names))
            loose = [p.get("name") for p in page.findall("property") if p.get("name")]
            if loose:
                secs.insert(0, ("", loose))
            sections[title] = secs
        return {"pages": pages, "sections": sections}
    return ctx.cached(f"layout:{type_name}", 3600, load)


def read_page(ctx: Ctx, obj: dict, type_name: str, page: str, include_empty: bool, fallback_props: list[str]) -> dict:
    lay = layout(ctx, obj, type_name)
    sections = []
    if lay:
        want = next((t for t in lay["pages"] if t.lower() == (page or "Overview").lower()), lay["pages"][0] if lay["pages"] else "")
        secs = lay["sections"].get(want, [])
        names = sorted({n for _, ns in secs for n in ns})
        props = ctx.tc.get_props([obj], names).get(obj["uid"], {})
        for title, ns in secs:
            fields = []
            for n in ns:
                v = ctx.ui(props, n)
                if v in ("", [], None) and not include_empty:
                    continue
                fields.append({"label": _label(n), "name": n, "value": v})
            if fields:
                sections.append({"section": title, "fields": fields})
        return {"page": want, "tabs": lay["pages"], "sections": sections, "layout_source": "Active Workspace style sheet"}
    props = ctx.tc.get_props([obj], fallback_props).get(obj["uid"], {})
    fields = [{"label": _label(n), "name": n, "value": ctx.ui(props, n)} for n in fallback_props
              if include_empty or ctx.ui(props, n) not in ("", [], None)]
    return {"page": page or "Overview", "tabs": [], "sections": [{"section": "Properties", "fields": fields}],
            "layout_source": "profile field list (style sheet not available)"}


def _label(prop: str) -> str:
    p = re.sub(r"^(g4b_|fnd0|bl_|object_)", "", prop)
    return re.sub(r"[_]+", " ", re.sub(r"(?<=[a-z])(?=[A-Z])", " ", p)).strip().title()


# =============================================================================
# 7. READ TOOLS (18)
# =============================================================================
ID_LIKE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.\-]*\*?$")


def _looks_like_id(q: str) -> bool:
    q = q.strip()
    return bool(ID_LIKE.match(q)) and (any(c.isdigit() for c in q) or q.endswith("*") or "-" in q) and " " not in q


def _rev_summary(c, item_id: str, rev: dict) -> dict:
    rec = rev["rec"]
    return {"id": item_id, "revision": rev["rev"], "name": c.ui(rec, c.P("name")), "description": c.ui(rec, c.P("desc")),
            "type": c.ui(rec, c.P("type")), "status": ", ".join(rev["status"]) or "Working",
            "released": rev["released"], "link": c.link(rev["ref"]["uid"])}

# ======================================================================= identity

@tool
@tool_errors
def whoami() -> dict:
    """Shows who you are to zNext / Teamcenter: your user, access tier (user / exec), client app,
    allowed tools and session state. First call to check access; never returns part data."""
    c = ctx()
    me = c.me()
    names = c.tc.get_props([x for x in (me.get("group"), me.get("role")) if x], ["name", "role_name", "object_string"])
    def nm(o):
        p = names.get((o or {}).get("uid", ""), {})
        return c.ui(p, "name") or c.ui(p, "role_name") or c.ui(p, "object_string")
    pass
    tools = sorted(TOOL_NAMES)
    return {"user": me["user_id"], "name": me["user_name"], "group": nm(me.get("group")), "role": nm(me.get("role")),
            "access_tier": c.tier(), "client_app": "tcplm-mcp", "system": c.cfg.label,
            "profile": c.cfg["profile"]["name"], "server_version": me.get("server", ""),
            "session": "connected", "allowed_tools": tools or "all tools of this server"}

# ======================================================================= search

@tool
@tool_errors
def search_znext(query: str, object_types: list[str] | None = None, page: int = 0, page_size: int = 25) -> dict:
    """One search for anything: parts / GPNs, MPNs, change orders (ECO, MCO, DEV, PRQ, MCR), documents - by GPN or
    prefix, id, name, description or keywords, with wildcards (e.g. 'G650-0143*', 'ECO-A66*', 'DAUGHTER BOARD').
    GPNs / prefixes use a live query (finds objects created seconds ago); other text uses the full-text index
    (with type_facets), with a live fallback when the index has no hit. Each result says source: live | index.
    object_types: optional type filter, e.g. ['G4B_ChangeOrderRevision']. Max 25 per page."""
    c = ctx()
    page_size = max(1, min(page_size, 25))
    q = query.strip()
    results, facets, total, path = [], {}, 0, ""

    def live(text: str, by_name: bool) -> list[dict]:
        qc = c.cfg["queries"]
        if by_name:
            objs = c.tc.run_saved_query(qc["general"], {qc["general_name_entry"]: text if "*" in text else f"*{text}*"}, limit=200)
        else:
            objs = c.tc.run_saved_query(qc["item"], {qc["item_id_entry"]: text}, limit=200)
        return [(o, "live") for o in objs]

    if _looks_like_id(q):
        found, path = live(q, False), "live (saved query by id)"
    else:
        try:
            objs, total, facets = c.tc.fulltext(q, object_types, page * page_size, page_size)
            found, path = [(o, "index") for o in objs], "index (full-text)"
        except TCError:
            found = []
        if not found:
            found, path = live(q, True), "live fallback (saved query by name)"
    objs = [o for o, _ in found]
    props = c.tc.get_props(objs, [c.P("id"), c.P("rev"), c.P("name"), c.P("desc"), c.P("type"), c.P("status"), c.P("modified")])
    for o, src in found:
        p = props.get(o["uid"])
        if not p:
            continue
        t = p.get("type") or ""
        if object_types and t not in object_types and str(c.ui(p, c.P("type"))) not in object_types:
            continue
        results.append({"id": c.ui(p, c.P("id")), "revision": c.ui(p, c.P("rev")), "name": c.ui(p, c.P("name")),
                        "description": c.ui(p, c.P("desc")), "type": c.ui(p, c.P("type")) or t,
                        "status": c.ui(p, c.P("status")), "modified": c.ui(p, c.P("modified")),
                        "source": src, "link": c.link(o["uid"])})
    if path.startswith("index"):
        return {"query": q, "path": path, "total": total or len(results), "page": page, "page_size": page_size,
                "results": results, "type_facets": facets}
    pg = page_of(results, page, page_size, 25)
    return {"query": q, "path": path, "total": pg["total"], "page": page, "page_size": page_size,
            "has_more": pg["has_more"], "results": pg["items"]}

# ======================================================================= part details

@tool
@tool_errors
def get_component_overview(gpn: str, revision: str = "latest_released") -> dict:
    """One-call summary of a GPN: key attributes, AML summary (manufacturers, preferred MPNs) and level-1 BOM
    summary. Best first call for a GPN. revision: 'latest_released' (default), 'latest' or e.g. 'B'."""
    c = ctx()
    item, rev = c.resolve(gpn, revision)
    r = rev["ref"]
    attrs = c.tc.get_props([r], c.cfg["props"]["overview"]).get(r["uid"], {})
    srcs = aml(c, [r]).get(r["uid"], [])
    try:
        bom = bom_lines(c, r, 1, 500)
    except TCError as e:
        bom, bom_err = [], str(e)
    else:
        bom_err = ""
    pref = [s for s in srcs if str(s.get("preferred_status", "")).lower() == "preferred"]
    revs = c.revisions(item)
    return {**_rev_summary(c, gpn, rev),
            "attributes": {n: c.ui(attrs, n) for n in c.cfg["props"]["overview"] if c.ui(attrs, n) not in ("", [])},
            "revisions": [f"{x['rev']} ({', '.join(x['status']) or 'Working'})" for x in revs],
            "aml": {"mpn_count": len(srcs), "manufacturers": sorted({str(s["manufacturer"]) for s in srcs if s["manufacturer"]}),
                    "preferred_mpns": [f"{s['mpn']} ({s['manufacturer']})" for s in pref]},
            "bom_level1": {"lines": len(bom), "first_lines": public(bom[:10]), **({"error": bom_err} if bom_err else {})}}


@tool
@tool_errors
def get_components_attributes(gpns: list[str], revision: str = "latest_released", fields: list[str] | None = None) -> dict:
    """Attributes of 1 to 25 GPNs in one call (core + zNext g4b_* fields: category, sourcing, make/buy, RoHS,
    compliance, lifecycle); errors are reported per GPN instead of failing the call. Max 25 GPNs."""
    c = ctx()
    if not gpns or len(gpns) > 25:
        raise ToolError("Give 1 to 25 GPNs.", "BAD_INPUT")
    names = fields or c.cfg["props"]["attributes"]
    results, refs = [], []
    for g in gpns:
        try:
            _, rev = c.resolve(g, revision)
            refs.append((g, rev))
        except ToolError as e:
            results.append({"gpn": g, "error": e.code, "message": str(e)})
    props = c.tc.get_props([rv["ref"] for _, rv in refs], names)
    for g, rv in refs:
        p = props.get(rv["ref"]["uid"], {})
        results.append({"gpn": g, "revision": rv["rev"], "link": c.link(rv["ref"]["uid"]),
                        "attributes": {n: c.ui(p, n) for n in names}})
    order = {g: i for i, g in enumerate(gpns)}
    return {"count": len(results), "results": sorted(results, key=lambda x: order.get(x["gpn"], 0))}

# ======================================================================= structure

@tool
@tool_errors
def get_bom_structure(gpn: str, revision: str = "latest_released", depth: int = 1, page: int = 0, page_size: int = 200) -> dict:
    """BOM of an assembly: find no, GPN, revision, description, quantity, UOM, reference designators, function,
    primary/alternate. Depth 1-5, paginated; the BOM window is always closed afterwards."""
    c = ctx()
    _, rev = c.resolve(gpn, revision)
    lines = public(bom_lines(c, rev["ref"], depth))
    pg = page_of(lines, page, page_size, 1000)
    return {"gpn": gpn, "revision": rev["rev"], "depth": max(1, min(depth, 5)), "link": c.link(rev["ref"]["uid"]),
            "total_lines": pg["total"], "page": pg["page"], "has_more": pg["has_more"], "lines": pg["items"]}


@tool
@tool_errors
def compare_bom_revisions(gpn: str, revision_a: str, revision_b: str = "latest", depth: int = 1, gpn_b: str = "") -> dict:
    """Diff of two BOMs (two revisions of one assembly, or two assemblies with gpn_b): added, removed and changed
    lines (qty, revision, UOM, ref des, find no, function, primary/alt). Repeated parts merged per parent; max 500
    rows per list."""
    c = ctx()
    _, ra = c.resolve(gpn, revision_a)
    _, rb = c.resolve(gpn_b or gpn, revision_b)
    a, b = bom_lines(c, ra["ref"], depth), bom_lines(c, rb["ref"], depth)
    return {"a": f"{gpn}/{ra['rev']}", "b": f"{gpn_b or gpn}/{rb['rev']}", **bom_diff(a, b)}


@tool
@tool_errors
def get_where_used(gpn: str, revision: str = "latest_released", depth: int = 1, page: int = 0, page_size: int = 25,
                   released_only: bool = False) -> dict:
    """Parent assemblies that use a GPN (which boards / products contain this part), with links.
    Depth 1-2; paginated; released_only keeps only released parents."""
    c = ctx()
    _, rev = c.resolve(gpn, revision)
    resp = c.tc.call("whereUsed", {"objects": [c.tc.ref(rev["ref"])], "numLevels": max(1, min(depth, 2)),
                                   "whereUsedPrecise": False, "rule": {"uid": "AAAAAAAAAAAAAA", "type": "unknownType"}},
                     timeout=45)
    found = []
    for o in resp.get("output") or []:
        for inf in o.get("info") or []:
            pr = inf.get("parentItemRev")
            if pr:
                found.append((pr, inf.get("level", 1)))
    props = c.tc.get_props([p for p, _ in found], [c.P("id"), c.P("rev"), c.P("name"), c.P("status"), c.P("type")])
    rows, seen = [], set()
    for p, lvl in found:
        x = props.get(p["uid"], {})
        status = c.ui(x, c.P("status"))
        if p["uid"] in seen or (released_only and not status):
            continue
        seen.add(p["uid"])
        rows.append({"level": lvl, "id": c.ui(x, c.P("id")), "revision": c.ui(x, c.P("rev")), "name": c.ui(x, c.P("name")),
                     "type": c.ui(x, c.P("type")), "status": status or "Working", "link": c.link(p["uid"])})
    pg = page_of(rows, page, page_size, 200)
    return {"gpn": gpn, "revision": rev["rev"], "total": pg["total"], "page": page, "has_more": pg["has_more"], "parents": pg["items"]}

# ======================================================================= sourcing

@tool
@tool_errors
def get_item_sourcing_batch(gpns: list[str], revision: str = "latest_released") -> dict:
    """Item sourcing / AML (Manufacturers tab) of 1 to 25 GPNs: MPN, manufacturer, preferred / qualification,
    lifecycle and RoHS status; errors are reported per GPN. Max 25 GPNs."""
    c = ctx()
    if not gpns or len(gpns) > 25:
        raise ToolError("Give 1 to 25 GPNs.", "BAD_INPUT")
    if not c.cfg.get("relations", "aml"):
        return c.not_available("Sourcing (AML relation)")
    out, revs = [], []
    for g in gpns:
        try:
            _, rv = c.resolve(g, revision)
            revs.append((g, rv))
        except ToolError as e:
            out.append({"gpn": g, "error": e.code, "message": str(e)})
    amls = aml(c, [rv["ref"] for _, rv in revs])
    for g, rv in revs:
        rows = public(amls.get(rv["ref"]["uid"], []))
        out.append({"gpn": g, "revision": rv["rev"], "mpn_count": len(rows), "aml": rows})
    return {"count": len(out), "results": out}

# ======================================================================= changes & workflow

@tool
@tool_errors
def get_change_order(change_id: str, revision: str = "latest") -> dict:
    """Change object details (ECO, MCO, DEV, PRQ, MCR, CR): status, maturity, disposition, reason, builds and teams
    impacted, requestor / analyst, plus solution, impacted and problem items (up to 200 per relation)."""
    c = ctx()
    _, rev, kind = c.change(change_id, revision)
    r = rev["ref"]
    fields = c.cfg["change_fields"]
    p = c.tc.get_props([r], list(fields.values()) + [c.P("name"), c.P("desc"), c.P("type")]).get(r["uid"], {})
    items = change_items(c, r)
    return {"id": change_id, "kind": kind.upper() or c.ui(p, c.P("type")), "revision": rev["rev"],
            "name": c.ui(p, c.P("name")), "description": c.ui(p, c.P("desc")), "link": c.link(r["uid"]),
            "fields": {label: c.ui(p, prop) for label, prop in fields.items()},
            "participants": participants(c, r),
            **{k: public(v) for k, v in items.items()}}


@tool
@tool_errors
def get_workflow_status(object_id: str, revision: str = "latest") -> dict:
    """Why is a change or part workflow stuck: workflows, active tasks, assignees, sign-offs (decided / pending,
    required / optional), task age and a plain-language diagnosis. Your own access only."""
    c = ctx()
    _, rev = c.resolve(object_id, revision or "latest")
    wfs = workflows(c, rev["ref"])
    clean = [{"workflow": w["workflow"], "status": w["status"],
              "active_tasks": [{k: v for k, v in t.items() if not k.startswith("_")} | {
                  "signoffs": [{k: v for k, v in s.items() if not k.startswith("_")} for s in t["signoffs"]]}
                  for t in w["active_tasks"]]} for w in wfs]
    return {"object": f"{object_id}/{rev['rev']}", "link": c.link(rev["ref"]["uid"]), "workflows": clean,
            "diagnosis": diagnose(wfs)}


@tool
@tool_errors
def get_my_tasks(include_tracked: bool = False, page: int = 0, page_size: int = 25) -> dict:
    """Your workflow inbox: tasks to perform (and optionally tracked), newest first, with workflow, target object +
    link and age. Max 500 tasks. Approve / reject with act_on_task (or act_on_change for ECO sign-offs)."""
    c = ctx()
    tasks = public(inbox(c, include_tracked))
    pg = page_of(tasks, page, page_size, 100)
    return {"user": c.me()["user_id"], "total": pg["total"], "page": page, "has_more": pg["has_more"], "tasks": pg["items"]}


def _redline(c, change_id: str, revision: str, full: bool) -> dict:
    _, rev, kind = c.change(change_id, revision)
    r = rev["ref"]
    hp = c.tc.get_props([r], [c.P("name"), c.P("desc")] + list(c.cfg["change_fields"].values())).get(r["uid"], {})
    items = change_items(c, r, cap=80)
    sol_by_id = {}
    for s in items.get("solution", []):
        sol_by_id.setdefault(s["id"], s)
    attrs = c.cfg["props"]["attributes"]
    per_item, bom_done = [], 0
    for imp in items.get("impacted", []):
        sol = sol_by_id.get(imp["id"])
        entry: dict[str, Any] = {"id": imp["id"], "from": imp["revision"], "to": sol["revision"] if sol else "",
                                 "status_from": imp["status"], "status_to": sol["status"] if sol else "",
                                 "link": imp["link"]}
        if sol:
            pa = c.tc.get_props([imp["_obj"]], attrs).get(imp["_obj"]["uid"], {})
            pb = c.tc.get_props([sol["_obj"]], attrs).get(sol["_obj"]["uid"], {})
            entry["property_changes"] = {n: {"from": c.ui(pa, n), "to": c.ui(pb, n)} for n in attrs
                                         if str(c.ui(pa, n)) != str(c.ui(pb, n)) and n != c.P("status")}
            if bom_done < 20:
                try:
                    diff = bom_diff(bom_lines(c, imp["_obj"], 1, 2000), bom_lines(c, sol["_obj"], 1, 2000), 200)
                    entry["bom_changes"] = {k: diff[k] for k in ("added", "removed", "changed")}
                    bom_done += 1
                except TCError:
                    entry["bom_changes"] = "not available"
            if c.cfg.get("relations", "aml"):
                aa = {s["mpn"]: s for s in public(aml(c, [imp["_obj"]]).get(imp["_obj"]["uid"], []))}
                ab = {s["mpn"]: s for s in public(aml(c, [sol["_obj"]]).get(sol["_obj"]["uid"], []))}
                entry["aml_changes"] = {"added": [ab[k] for k in ab if k not in aa], "removed": [aa[k] for k in aa if k not in ab],
                                        "changed": [{"mpn": k, "from": aa[k], "to": ab[k]} for k in aa.keys() & ab.keys() if aa[k] != ab[k]]}
            fa = {(f["name"], f["type"]) for f in attachments(c, imp["_obj"])}
            fb = {(f["name"], f["type"]) for f in attachments(c, sol["_obj"])}
            entry["file_changes"] = {"added": sorted(f"{n} ({t})" for n, t in fb - fa), "removed": sorted(f"{n} ({t})" for n, t in fa - fb)}
        per_item.append(entry)
    wfs = workflows(c, r)
    team = [{"task": t["task"], "person": s["person"], "decision": s["decision"], "required": s["required"]}
            for w in wfs for t in w["active_tasks"] for s in t["signoffs"]]
    mine = [t for t in inbox(c) if t.get("_target_uid") == r["uid"]]
    header = {"id": change_id, "kind": kind.upper(), "revision": rev["rev"], "title": c.ui(hp, c.P("desc")) or c.ui(hp, c.P("name")),
              "fields": {k: c.ui(hp, v) for k, v in c.cfg["change_fields"].items() if c.ui(hp, v) not in ("", [])},
              "link": c.link(r["uid"])}
    fp = fingerprint({"items": per_item, "team": team})
    md = [f"### {change_id} ({header['kind']}) - {header['title']}", ""]
    md += [f"- **{k}:** {v}" for k, v in header["fields"].items()]
    md += ["", "| Item | Revision | Status | Changes |", "|---|---|---|---|"]
    for e in per_item:
        bits = []
        if e.get("property_changes"):
            bits.append("props: " + ", ".join(e["property_changes"]))
        bc = e.get("bom_changes")
        if isinstance(bc, dict) and any(bc.values()):
            bits.append(f"BOM +{len(bc['added'])}/-{len(bc['removed'])}/~{len(bc['changed'])}")
        ac = e.get("aml_changes")
        if ac and any(ac.values()):
            bits.append(f"AML +{len(ac['added'])}/-{len(ac['removed'])}/~{len(ac['changed'])}")
        fc = e.get("file_changes")
        if fc and any(fc.values()):
            bits.append(f"files +{len(fc['added'])}/-{len(fc['removed'])}")
        md.append(f"| {e['id']} | {e['from']} -> {e['to'] or '?'} | {e['status_from'] or 'Working'} -> {e['status_to'] or '?'} | "
                  f"{'; '.join(bits) or 'no changes found'} |")
    if team:
        md += ["", "**Sign-off team:** " + "; ".join(f"{t['person']} ({t['decision']})" for t in team)]
    md += ["", f"**Your pending sign-off:** {mine[0]['task'] if mine else 'none'}", f"[Open in {c.cfg.label}]({header['link']})"]
    return {"header": header, "items": per_item, "signoff_team": team,
            "my_pending_signoff": [{k: v for k, v in t.items() if not k.startswith("_")} for t in mine],
            "card_markdown": "\n".join(md), "redline_fingerprint": fp,
            "_rev": r, "_mine": mine, "_full": full}


@tool
@tool_errors
def get_change_redline(change_id: str, revision: str = "latest", format: str = "card") -> dict:
    """The ECO / MCO approval card: header, and per affected item the redline between impacted and solution
    revisions - revision/status change, property changes, BOM adds/removes/qty and ref-des changes, AML/MPN
    changes, drawing/file changes - plus the sign-off team and your own pending sign-off. Returns ready-to-show
    card_markdown and a redline fingerprint. format: 'card' (default) or 'full' (adds per-item data and approvers).
    Max 80 items per side; BOM redline for up to 20 assemblies; never cached."""
    c = ctx()
    rl = _redline(c, change_id, revision, format == "full")
    out = {k: v for k, v in rl.items() if not k.startswith("_")}
    if format != "full":
        out.pop("items", None)
    return out

# ======================================================================= files & export

@tool
@tool_errors
def get_attachments(gpn: str, revision: str = "latest_released", dataset_types: list[str] | None = None) -> dict:
    """Files attached to a GPN revision (drawings, datasheets, specs, CAD) with a link for each.
    Files are never downloaded: links only."""
    c = ctx()
    _, rev = c.resolve(gpn, revision)
    files = public(attachments(c, rev["ref"], dataset_types))
    return {"gpn": gpn, "revision": rev["rev"], "count": len(files), "files": files}


def _csv(rows: list[dict], max_rows: int) -> tuple[str, bool, int]:
    if not rows:
        return "", False, 0
    max_rows = max(1, min(max_rows, 20000))
    cols = list(dict.fromkeys(k for r in rows for k in r.keys()))
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=cols, extrasaction="ignore")
    w.writeheader()
    n, trunc = 0, len(rows) > max_rows
    for r in rows[:max_rows]:
        safe = {}
        for k, v in r.items():
            v = ", ".join(map(str, v)) if isinstance(v, list) else ("" if v is None else str(v))
            safe[k] = "'" + v if v[:1] in ("=", "+", "-", "@", "\t", "\r") else v   # formula-injection safe
        w.writerow(safe)
        n += 1
        if buf.tell() > 900_000:
            trunc = True
            break
    return buf.getvalue(), trunc, n


@tool
@tool_errors
def export_csv(kind: str, gpn: str = "", gpns: list[str] | None = None, revision: str = "latest_released", depth: int = 1,
               revision_b: str = "latest", released_only: bool = False, max_rows: int = 5000,
               user_requested: bool = False) -> dict:
    """Bulk CSV export (text + filename) for: bom, aml, where_used, attributes, bom_compare, my_tasks.
    Runs ONLY when the user explicitly asked for a file / CSV export (user_requested=true); otherwise refused.
    Never export or offer CSVs on your own: summarize, page and link instead. max_rows default 5000 (max 20000)."""
    c = ctx()
    if not user_requested:
        raise ToolError("CSV export runs only when the user explicitly asks for a file. Summarize the data instead.", "NOT_REQUESTED")
    k = kind.lower()
    if k == "bom":
        _, rev = c.resolve(gpn, revision)
        rows, name = public(bom_lines(c, rev["ref"], depth, 20000)), f"{gpn}_{rev['rev']}_bom.csv"
    elif k == "aml":
        res = get_item_sourcing_batch(gpns or [gpn], revision)
        rows = [{"gpn": r["gpn"], **a} for r in res.get("results", []) for a in r.get("aml", [])]
        name = f"{(gpns or [gpn])[0]}_aml.csv"
    elif k == "where_used":
        res = get_where_used(gpn, revision, depth, 0, 20000, released_only)
        rows, name = res.get("parents", []), f"{gpn}_where_used.csv"
    elif k == "attributes":
        res = get_components_attributes(gpns or [gpn], revision)
        rows = [{"gpn": r["gpn"], "revision": r.get("revision", ""), **r.get("attributes", {})} for r in res.get("results", [])]
        name = f"{(gpns or [gpn])[0]}_attributes.csv"
    elif k == "bom_compare":
        res = compare_bom_revisions(gpn, revision, revision_b, depth)
        rows = ([{"change": "added", **x} for x in res.get("added", [])] + [{"change": "removed", **x} for x in res.get("removed", [])]
                + [{"change": "changed", "parent": x["parent"], "gpn": x["gpn"],
                    **{f"{f}_from": v["from"] for f, v in x["changes"].items()}, **{f"{f}_to": v["to"] for f, v in x["changes"].items()}}
                   for x in res.get("changed", [])])
        name = f"{gpn}_bom_compare.csv"
    elif k == "my_tasks":
        rows, name = public(inbox(c, True)), "my_tasks.csv"
    else:
        raise ToolError("kind must be one of: bom, aml, where_used, attributes, bom_compare, my_tasks.", "BAD_INPUT")
    if isinstance(rows, dict) and rows.get("error"):
        return rows
    text, trunc, n = _csv(rows, max_rows)
    return {"filename": name, "rows": n, "truncated": trunc, "csv": text}

# ======================================================================= objects

def _find_named(c, name: str, types: list[str]) -> dict | None:
    qc = c.cfg["queries"]
    objs = c.tc.run_saved_query(qc["general"], {qc["general_name_entry"]: name}, limit=50)
    props = c.tc.get_props(objs, [c.P("name"), c.P("type")])
    exact = [o for o in objs if str(c.ui(props.get(o["uid"], {}), c.P("name"))).lower() == name.lower()
             and (not types or props.get(o["uid"], {}).get("type") in types or c.ui(props.get(o["uid"], {}), c.P("type")) in types)]
    return exact[0] if exact else None


@tool
@tool_errors
def get_object(object_id: str, object_type: str = "", page: str = "Overview", include_empty: bool = False,
               revision: str = "", include_parts: bool = False) -> dict:
    """Any object's page (default Overview), labeled like Active Workspace, plus its other tabs and a link.
    By id: GPN, document, MPN (+ GPNs using it, vendor documents), PNR / PRQ / CR / ECO (+ impacted / solution /
    problem items, participants, workflow summary). By name: manufacturer (+ plants; include_parts=true lists its
    parts, slow), plant (+ manufacturer), product line. object_type narrows a lookup (e.g. 'G4B_Manufacturer',
    'G4B_MfgPlantCode', 'G4B_ProductLine', or 'G4B_MfgPart' for an MPN whose id equals a GPN).
    Empty fields are hidden unless include_empty."""
    c = ctx()
    t = c.cfg["types"]
    named_types = [x for x in (t.get("manufacturer"), t.get("plant"), t.get("product_line")) if x]
    obj, kind, rel = None, "", {}
    if object_type and object_type in named_types:
        obj = _find_named(c, object_id, [object_type])
    if obj is None:
        try:
            types = [object_type] if object_type else None
            it, rv = c.resolve(object_id, revision or "latest", types)
            obj, kind = rv["ref"], c.cfg.change_kind(object_id) or "item"
        except ToolError:
            obj = _find_named(c, object_id, [object_type] if object_type else named_types)
            kind = "named"
    if obj is None:
        raise ToolError(f"Nothing called {object_id} found.", "NOT_FOUND")
    tp = c.tc.get_props([obj], [c.P("type")]).get(obj["uid"], {})
    type_name = tp.get("type") or object_type or ""
    fallback = list(dict.fromkeys(c.cfg["props"]["attributes"] + list(c.cfg["change_fields"].values())))
    pg = read_page(c, obj, type_name, page, include_empty, fallback)
    if kind in c.cfg.p.get("change_types", {}) or type_name.replace("Revision", "") in c.cfg.all_change_types():
        rel = {k: public(v) for k, v in change_items(c, obj).items()}
        rel["participants"] = participants(c, obj)
        rel["workflow"] = diagnose(workflows(c, obj))
    elif type_name.replace("Revision", "") in (t.get("mpn_types") or []) and c.cfg.get("relations", "aml"):
        side = obj
        if c.cfg.get("relations", "aml_secondary") != "revision":
            it = c.tc.get_props([obj], ["items_tag"]).get(obj["uid"], {})
            iu = (it.get("db", {}).get("items_tag") or [""])[0] if isinstance(it.get("db", {}).get("items_tag"), list) else it.get("db", {}).get("items_tag")
            if iu:
                side = {"uid": iu, "type": "Item"}
        using = c.tc.expand_relations([side], [c.cfg.get("relations", "aml")], secondary=True).get(side["uid"], [])
        ip = c.tc.get_props([x["other"] for x in using], [c.P("id"), c.P("rev"), c.P("name")])
        rel["used_by_gpns"] = [{"gpn": c.ui(ip.get(x["other"]["uid"], {}), c.P("id")), "revision": c.ui(ip.get(x["other"]["uid"], {}), c.P("rev")),
                                "link": c.link(x["other"]["uid"])} for x in using]
        if c.cfg.get("relations", "vendor_docs"):
            docs = c.tc.expand_relations([obj], [c.cfg.get("relations", "vendor_docs")]).get(obj["uid"], [])
            dp = c.tc.get_props([x["other"] for x in docs], [c.P("name")])
            rel["vendor_documents"] = [{"name": c.ui(dp.get(x["other"]["uid"], {}), c.P("name")), "link": c.link(x["other"]["uid"])} for x in docs]
    elif type_name == t.get("manufacturer") and c.cfg.get("relations", "plants"):
        pl = c.tc.expand_relations([obj], [c.cfg.get("relations", "plants")]).get(obj["uid"], [])
        pp = c.tc.get_props([x["other"] for x in pl], [c.P("name")])
        rel["plants"] = [{"name": c.ui(pp.get(x["other"]["uid"], {}), c.P("name")), "link": c.link(x["other"]["uid"])} for x in pl]
        if include_parts:
            try:
                objs, total, _ = c.tc.fulltext(object_id, t.get("mpn_types"), 0, 25)
                mp = c.tc.get_props(objs, [c.P("id"), c.P("name")])
                rel["parts"] = {"first_25": [c.ui(mp.get(o["uid"], {}), c.P("id")) for o in objs], "approx_total": total,
                                "note": "from the search index (approximate)"}
            except TCError as e:
                rel["parts"] = {"error": str(e)}
    elif type_name == t.get("plant") and c.cfg.get("relations", "plants"):
        mf = c.tc.expand_relations([obj], [c.cfg.get("relations", "plants")], secondary=True).get(obj["uid"], [])
        mp = c.tc.get_props([x["other"] for x in mf], [c.P("name")])
        rel["manufacturer"] = [c.ui(mp.get(x["other"]["uid"], {}), c.P("name")) for x in mf]
    return {"object": object_id, "type": type_name, "link": c.link(obj["uid"]), **pg, **({"related": rel} if rel else {})}

# ======================================================================= create form (read-only)

def _prop_descs(c, type_name: str, op: str) -> list[dict]:
    """Property descriptors for a type and operation (CREATE / EDIT): name, label, required, has_lov, default."""
    resp = c.tc.call("getAttachedPropDescs2", {"inputs": [{"typeName": type_name, "propNames": [],
                                                           "operationType": {"CREATE": 0, "REVISE": 1, "EDIT": 2}.get(op, 0)}]})
    out = []
    def walk(o):
        if isinstance(o, dict):
            name = o.get("propName") or o.get("propertyName")
            if name and ("displayName" in o or "propertyDisplayName" in o):
                cm = o.get("constantsMap") or {}
                out.append({"name": name, "label": o.get("displayName") or o.get("propertyDisplayName"),
                            "required": bool(o.get("isRequired") or o.get("required") or str(cm.get("required", "")).lower() == "true"),
                            "has_lov": bool(o.get("lovAttached") or o.get("lov") or o.get("hasLOV")),
                            "default": o.get("initialValue") or o.get("defaultValue") or cm.get("InitialValue", ""),
                            "max_length": o.get("maxLength") or o.get("maxStringLength")})
            else:
                for v in o.values():
                    walk(v)
        elif isinstance(o, list):
            for v in o:
                walk(v)
    walk(resp)
    seen, uniq = set(), []
    for d in out:
        if d["name"] not in seen:
            seen.add(d["name"])
            uniq.append(d)
    return uniq


def _lov(c, type_name: str, prop: str, op: str, filter_text: str = "", owner: dict | None = None, values: dict | None = None) -> list[str]:
    try:
        resp = c.tc.call("getInitialLOVValues", {"initialData": {
            "propertyName": prop, "lov": {"uid": "AAAAAAAAAAAAAA", "type": "unknownType"},
            "lovInput": {"owningObject": owner or {"uid": "AAAAAAAAAAAAAA", "type": "unknownType"}, "boName": type_name,
                         "operationName": op.title(), "propertyValues": {k: [str(v)] for k, v in (values or {}).items()}},
            "filterData": {"filterString": filter_text, "maxResults": 2000, "numberToReturn": 200, "order": 1, "sortPropertyName": ""}}})
    except TCError:
        return []
    vals = []
    for v in resp.get("lovValues") or []:
        disp = (v.get("propDisplayValues") or {}).get("lov_values") or (v.get("propInternalValues") or {}).get("lov_values") or []
        vals += [str(x) for x in disp]
    return vals


def _next_id(c, type_name: str, prop: str = "item_id") -> str:
    try:
        resp = c.tc.call("generateNextValuesForProperties", {"propertyNamingRuleInfo": [{
            "clientId": "n", "operationType": 1, "businessObjectTypeName": type_name,
            "propertyNameAttachedPattern": {prop: ""}, "additionalInputParams": {}}]})
    except TCError:
        return ""
    for g in resp.get("generatedValues") or []:
        gv = (g.get("generatedValues") or {}).get(prop) or {}
        if gv.get("nextValue"):
            return gv["nextValue"]
    return ""


@tool
@tool_errors
def prepare_create_item(description: str = "", similar_to: str = "", object_type: str = "", lov_filters: dict | None = None) -> dict:
    """Step 1 of creating a part / item: the create form for the right type (picked from a similar GPN, a type or a
    description) - fields, which are required, allowed values, defaults, the auto-assigned GPN base number - plus
    possible duplicates. Nothing is created. lov_filters searches a long value list, e.g. {"g4b_product_lines": "A70"}.
    Then ask the user for missing values and call create_item."""
    c = ctx()
    t, how = object_type, "object_type"
    if not t and similar_to:
        it = c.item(similar_to)
        t, how = c.tc.get_props([it], [c.P("type")]).get(it["uid"], {}).get("type") or c.cfg["types"]["part"], f"same type as {similar_to}"
    dups = []
    if description:
        try:
            objs, _, _ = c.tc.fulltext(description, None, 0, 25)
        except TCError:
            objs = []
        props = c.tc.get_props(objs, [c.P("id"), c.P("rev"), c.P("name"), c.P("desc"), c.P("type")])
        types = Counter()
        for o in objs:
            p = props.get(o["uid"], {})
            tn = (p.get("type") or "").replace("Revision", "")
            if tn in (c.cfg["types"]["part_types"] or []) or not c.cfg["types"]["part_types"]:
                types[tn] += 1
                dups.append({"id": c.ui(p, c.P("id")), "revision": c.ui(p, c.P("rev")), "name": c.ui(p, c.P("name")),
                             "description": c.ui(p, c.P("desc")), "type": tn, "link": c.link(o["uid"])})
        if not t and types:
            t, how = types.most_common(1)[0][0], "most common type among similar parts"
    if not t:
        t, how = c.cfg["types"]["part"], "profile default part type"
    item_fields = _prop_descs(c, t, "CREATE")
    rev_fields = _prop_descs(c, t + "Revision", "CREATE")
    filters = lov_filters or {}
    for f in item_fields + rev_fields:
        if f["has_lov"]:
            owner_t = t if f in item_fields else t + "Revision"
            vals = _lov(c, owner_t, f["name"], "CREATE", filters.get(f["name"], ""))
            f["allowed_values"] = vals[:200]
            f["allowed_values_truncated"] = len(vals) > 200
    cos = c.cfg["props"]["custom_off_shelf"]
    notes = []
    if cos:
        notes.append(f"'{cos}' (Custom vs Off-the-Shelf) is required. Custom: the MPN id will be the GPN. "
                     "Off-the-Shelf: ask the user for the manufacturer's MPN id and manufacturer (create_mpn after create).")
    return {"object_type": t, "type_chosen_by": how, "auto_assigned_id": _next_id(c, t, c.P("id")) or "(assigned on create)",
            "item_fields": item_fields, "revision_fields": rev_fields,
            "required": [f["name"] for f in item_fields + rev_fields if f["required"]],
            "possible_duplicates": dups[:10], "notes": notes,
            "next_step": "Ask the user for the missing required values, then call create_item (it returns a preview first)."}


# =============================================================================
# 8. WRITE HELPERS
# =============================================================================
NEXT = ("Show this preview to the user. Only after they explicitly say yes, call this tool again with exactly the "
        "same values plus confirm_token. The token is single use and expires in 10 minutes.")


def two_step(tool: str, args: dict, preview: Callable[[], tuple[dict, str]], write: Callable[[dict], dict]) -> dict:
    """preview() -> (preview dict, state fingerprint); write(preview) -> result. Rebuilds the preview on confirm and
    refuses if the state changed since the user saw it."""
    c = ctx()
    token = args.get("confirm_token") or ""
    if not token:
        pv, state = preview()
        if pv.get("blocking"):
            return {"preview": pv, "ready": False,
                    "message": "Fix the problems listed under 'blocking' first; nothing was created or changed."}
        return {"preview": pv, "ready": True, "confirm_token": c.confirm.issue(c.user_key(), tool, args, state),
                "expires_in_seconds": 600, "next": NEXT}
    rec = c.confirm.redeem(token, c.user_key(), tool, args)
    pv, state = preview()
    if pv.get("blocking"):
        raise ToolError("Not done: " + "; ".join(map(str, pv["blocking"])), "REFUSED")
    if rec.get("state") and state != rec["state"]:
        raise ToolError("The data changed since the preview (someone else edited it). Run the preview again.", "STALE")
    return {"done": True, **write(pv)}


def descriptors(c: Ctx, type_name: str, op: str) -> dict[str, dict]:
    try:
        return {d["name"]: d for d in _prop_descs(c, type_name, op)}
    except TCError:
        return {}


def validate(c: Ctx, type_name: str, values: dict, op: str, owner: dict | None = None,
             required_check: bool = True) -> tuple[dict, list[str], list[str]]:
    """Maps labels -> property names, checks required fields and value lists.
    Returns (values by property name, blocking problems, warnings)."""
    descs = descriptors(c, type_name, op)
    by_label = {str(d.get("label", "")).lower(): n for n, d in descs.items()}
    out, blocking, warn = {}, [], []
    for k, v in (values or {}).items():
        name = k if k in descs or not descs else by_label.get(k.lower(), k)
        if descs and name not in descs:
            blocking.append(f"'{k}' is not a field of {type_name}.")
            continue
        out[name] = v
        d = descs.get(name, {})
        if d.get("has_lov") and v not in ("", None, []):
            allowed = _lov(c, type_name, name, op, "", owner, out)
            vals = v if isinstance(v, list) else [v]
            bad = [x for x in vals if allowed and str(x) not in allowed]
            if bad:
                blocking.append(f"{d.get('label', name)}: {bad} not allowed. Allowed: {allowed[:30]}"
                                + (" ..." if len(allowed) > 30 else ""))
        ml = d.get("max_length")
        if ml and isinstance(v, str) and len(v) > int(ml):
            blocking.append(f"{d.get('label', name)} is longer than {ml} characters.")
    if required_check and descs:
        missing = [d.get("label") or n for n, d in descs.items() if d.get("required") and out.get(n) in (None, "", [])
                   and n not in ("item_id", "item_revision_id") and not d.get("default")]
        if missing:
            blocking.append("Required fields missing: " + ", ".join(missing))
    return out, blocking, warn


def _props_block(values: dict) -> dict:
    sp, sap, tp = {}, {}, {}
    for k, v in values.items():
        if v in (None, ""):
            continue
        if isinstance(v, dict) and v.get("uid"):
            tp[k] = v
        elif isinstance(v, list):
            sap[k] = [str(x) for x in v]
        elif isinstance(v, bool):
            sp[k] = "true" if v else "false"
        else:
            sp[k] = str(v)
    return {"stringProps": sp, "stringArrayProps": sap, "tagProps": tp}


def create_object(c: Ctx, type_name: str, values: dict, rev_values: dict | None = None, rev_type: str = "") -> list[dict]:
    data = {"boName": type_name, **_props_block(values)}
    if rev_values is not None:
        data["compoundCreateInput"] = {"revision": [{"boName": rev_type or type_name + "Revision", **_props_block(rev_values)}]}
    resp = c.tc.call("createObjects", {"input": [{"clientId": "c", "data": data}]})
    c.tc.raise_errors(resp, f"create {type_name}")
    objs = []
    for o in resp.get("output") or []:
        objs += o.get("objects") or []
    if not objs:
        objs = [{"uid": u, "type": "unknownType"} for u in (c.tc.sd(resp).get("created") or [])]
    if not objs:
        raise ToolError(f"Teamcenter did not create the {type_name}.", "TC_ERROR")
    return objs


def read_back(c: Ctx, objs: list[dict]) -> dict:
    props = c.tc.get_props(objs, [c.P("id"), c.P("rev"), c.P("name"), c.P("type")])
    item, rev = {}, {}
    for o in objs:
        p = props.get(o["uid"], {})
        if c.ui(p, c.P("rev")):
            rev = rev or {"uid": o["uid"], "revision": c.ui(p, c.P("rev")), "link": c.link(o["uid"])}
        if c.ui(p, c.P("id")) or c.ui(p, c.P("name")):
            item = item or {"uid": o["uid"], "id": c.ui(p, c.P("id")), "name": c.ui(p, c.P("name")),
                            "type": p.get("type") or c.ui(p, c.P("type")), "link": c.link(o["uid"])}
    return {"created": item, "revision": rev}


def find_named(c: Ctx, name: str, types: list[str]) -> dict | None:
    qc = c.cfg["queries"]
    try:
        objs = c.tc.run_saved_query(qc["general"], {qc["general_name_entry"]: name}, limit=50)
    except TCError:
        objs = []
    props = c.tc.get_props(objs, [c.P("name"), c.P("id"), c.P("type")])
    for o in objs:
        p = props.get(o["uid"], {})
        names = {str(c.ui(p, c.P("name"))).lower(), str(c.ui(p, c.P("id"))).lower()}
        if name.lower() in names and (not types or p.get("type") in types or c.ui(p, c.P("type")) in types):
            return o
    return None


# =============================================================================
# 9. CREATE TOOLS
# =============================================================================
# ======================================================================= item

@tool
@tool_errors
def create_item(object_type: str, values: dict, revision_values: dict | None = None, confirm_token: str = "") -> dict:
    """Creates a part / item. Without confirm_token: validated PREVIEW + single-use confirm_token (nothing created).
    With the token (after the user's explicit yes): reserves the GPN base number and creates the item + first
    revision. values = item fields (e.g. object_desc, g4b_prefix, uom_tag); revision_values = revision fields
    (e.g. g4b_product_lines: ["A70"], g4b_make_buy: "Buy"). Leave the id empty to auto-assign it.
    Custom vs Off-the-Shelf (g4b_custom_off_shelf) is required on zNext; after creating, continue with create_mpn.
    Call prepare_create_item first to get the form."""
    c = ctx()
    args = dict(object_type=object_type, values=values, revision_values=revision_values, confirm_token=confirm_token)
    cos = c.cfg["props"]["custom_off_shelf"]

    def preview():
        v, b1, w1 = validate(c, object_type, values, "CREATE")
        rv, b2, w2 = validate(c, object_type + "Revision", revision_values or {}, "CREATE")
        blocking = b1 + b2
        choice = (v.get(cos) or rv.get(cos)) if cos else None
        if cos and not choice:
            blocking.append(f"'{cos}' is required: Custom or Off-the-Shelf.")
        pv = {"action": f"Create a {object_type}", "id": v.get(c.P("id")) or "auto-assigned on create",
              "values": v, "revision_values": rv, "warnings": w1 + w2}
        if choice:
            pv["mpn_next_step"] = ("Custom: create_mpn will use the new GPN as the MPN id." if str(choice) == c.cfg["props"]["custom_value"]
                                   else "Off-the-Shelf: ask the user for the manufacturer's MPN id and manufacturer, then create_mpn.")
        if blocking:
            pv["blocking"] = blocking
        return pv, ""

    def write(pv):
        v = dict(pv["values"])
        if not v.get(c.P("id")):
            nid = _next_id(c, object_type, c.P("id"))
            if nid:
                v[c.P("id")] = nid
        objs = create_object(c, object_type, v, pv["revision_values"])
        out = read_back(c, objs)
        out["next"] = "Attach the MPN with create_mpn (preview first)." if cos else ""
        return out
    return two_step("create_item", args, preview, write)

# ======================================================================= profile-driven creates

def _split_item_rev(c, item_type: str, values: dict, default: str = "item") -> tuple[dict, dict]:
    """Puts each value on the item or on its first revision. A field the revision type declares goes on the
    revision, else one the item type declares goes on the item; unknown fields follow the profile's values_on.
    id, name and description always stay on the item."""
    item_d = descriptors(c, item_type, "CREATE")
    rev_d = descriptors(c, item_type + "Revision", "CREATE")
    keep = {c.P("id"), c.P("name"), "object_desc"}
    item_v, rev_v = {}, {}
    for k, v in values.items():
        if k in keep:
            item_v[k] = v
        elif k in rev_d:
            rev_v[k] = v
        elif k in item_d:
            item_v[k] = v
        else:
            (rev_v if default == "revision" else item_v)[k] = v
    return item_v, rev_v


def _profile_create(tool: str, kind: str, params: dict, extra_values: dict | None, confirm_token: str,
                    required: list[str], checks=None) -> dict:
    c = ctx()
    spec = c.cfg.p.get("create", {}).get(kind)
    if not spec or not spec.get("type"):
        return c.not_available(f"Creating a {kind.upper()}")
    t = spec["type"]
    args = dict(params, extra_values=extra_values, confirm_token=confirm_token)

    def preview():
        blocking = [f"'{r}' is required." for r in required if params.get(r) in (None, "", [])]
        mapping = spec.get("map", {})
        item_vals, ignored = {}, []
        for k, v in params.items():
            if v in (None, "", []) or k == spec.get("name_from"):
                continue
            if k in mapping:
                item_vals[mapping[k]] = v
            elif k not in ("manufacturer_id",):
                ignored.append(k)
        item_vals.update(extra_values or {})
        if spec.get("name_from") and spec["name_from"] != "_generated_id":
            item_vals[c.P("name")] = params.get(spec["name_from"], "")
        if kind == "manufacturer":
            item_vals[c.P("id")] = params.get("manufacturer_id") or params.get("name")
        if checks:
            blocking += checks(item_vals)
        v, b, w = validate(c, t, item_vals, "CREATE", required_check=False)
        blocking += b
        pv = {"action": f"Create a {kind.upper() if len(kind) <= 3 else kind.replace('_', ' ')} ({t})",
              "id": "assigned on create" if spec.get("name_from") == "_generated_id" or not v.get(c.P("id")) else v.get(c.P("id")),
              "values": v, "warnings": w + ([f"Not stored on this Teamcenter (no field configured): {ignored}"] if ignored else [])}
        if blocking:
            pv["blocking"] = blocking
        return pv, ""

    def write(pv):
        v = dict(pv["values"])
        if spec.get("name_from") == "_generated_id" or not v.get(c.P("id")):
            nid = _next_id(c, t, c.P("id"))
            if nid:
                v[c.P("id")] = nid
                if spec.get("name_from") == "_generated_id":
                    v[c.P("name")] = nid
        item_v, rev_v = _split_item_rev(c, t, v, spec.get("values_on", "item"))
        objs = create_object(c, t, item_v, rev_v)
        return read_back(c, objs)
    return two_step(tool, args, preview, write)


@tool
@tool_errors
def create_change_order(title: str, change_type: str, description_of_change: str = "", category: str = "",
                        reason_code: str = "", product_lines: list[str] | None = None, build_impacted: str = "",
                        functional_teams: str = "", description: str = "", confirm_token: str = "") -> dict:
    """Creates an ECO (change order): title + change type required; optional reason, category, product lines,
    build, team, description of change. The ECO id (e.g. ECO-A#####) is assigned on create; on zNext the id is the
    name, so the title is also saved as the description. Preview first, then confirm with the token."""
    p = dict(title=title, change_type=change_type, description_of_change=description_of_change, category=category,
             reason_code=reason_code, product_lines=product_lines, build_impacted=build_impacted,
             functional_teams=functional_teams, description=description)
    return _profile_create("create_change_order", "eco", p, None, confirm_token, ["title", "change_type"])


@tool
@tool_errors
def create_cr(synopsis: str, description: str = "", change_details: str = "", change_category: str = "",
              reason_code: str = "", product_lines: list[str] | None = None, functional_team: str = "",
              extra_values: dict | None = None, confirm_token: str = "") -> dict:
    """Creates a Change Request (CR): synopsis required; optional description, change details, change category,
    reason code, product lines, functional team impacted. The CR id is assigned on create. For an ECO use
    create_change_order. Preview first, then confirm."""
    p = dict(synopsis=synopsis, description=description, change_details=change_details, change_category=change_category,
             reason_code=reason_code, product_lines=product_lines, functional_team=functional_team)
    return _profile_create("create_cr", "cr", p, extra_values, confirm_token, ["synopsis"])


@tool
@tool_errors
def create_pnr(title: str, priority: str = "", schematic_symbol_required: str = "", footprint_required: str = "",
               copy_from_existing_part: str = "", comments: str = "", extra_values: dict | None = None,
               confirm_token: str = "") -> dict:
    """Creates a PNR (part number request): title, priority, schematic symbol / footprint required, copy from
    existing part, comments. The PNR id (PNR-NNNNNN) is assigned on create. Preview first, then confirm."""
    p = dict(title=title, priority=priority, schematic_symbol_required=schematic_symbol_required,
             footprint_required=footprint_required, copy_from_existing_part=copy_from_existing_part, comments=comments)
    return _profile_create("create_pnr", "pnr", p, extra_values, confirm_token, ["title"])


@tool
@tool_errors
def create_prq(title: str, description: str, product_area: str, prq_type: str, description_of_change: str = "",
               extra_values: dict | None = None, confirm_token: str = "") -> dict:
    """Creates a PRQ (Post-Ramp Qualification): title, description, product area, PRQ type, optional description of
    change. The PRQ id (PRQ-ANNNNN) is assigned on create and used as the name. Preview first, then confirm."""
    p = dict(title=title, description=description, product_area=product_area, prq_type=prq_type,
             description_of_change=description_of_change)
    return _profile_create("create_prq", "prq", p, extra_values, confirm_token, ["title", "description", "product_area", "prq_type"])


@tool
@tool_errors
def create_manufacturer(name: str, manufacturer_type: str = "", description: str = "", parent_company: str = "",
                        manufacturer_id: str = "", extra_values: dict | None = None, confirm_token: str = "") -> dict:
    """Creates a manufacturer (ID = name unless manufacturer_id is given; manufacturer type from the system's list,
    e.g. Component Supplier, Contract Manufacturer; description; optional parent company). zNext also creates its
    plant automatically. Preview first, then confirm."""
    c = ctx()

    def checks(vals):
        t = c.cfg["types"].get("manufacturer")
        if t and find_named(c, manufacturer_id or name, [t]):
            return [f"A manufacturer called '{manufacturer_id or name}' already exists."]
        return []
    p = dict(name=name, manufacturer_type=manufacturer_type, description=description, parent_company=parent_company,
             manufacturer_id=manufacturer_id)
    return _profile_create("create_manufacturer", "manufacturer", p, extra_values, confirm_token, ["name"], checks)


@tool
@tool_errors
def create_product_line(name: str, product_code: str = "", pdp_build: str = "", build_phase: str = "",
                        product_area: str = "", product_categories: str = "", extra_values: dict | None = None,
                        confirm_token: str = "") -> dict:
    """Creates a product line: name, PDP Build, Build Phase (depends on PDP Build: NPI -> Proto/EVT/DVT/PVT...,
    Sustaining -> N/A), Product Area/BU, Product Category, optional product code (max 6). New product lines show in
    search after indexing (a few minutes). Preview first, then confirm."""
    c = ctx()
    spec = c.cfg.p.get("create", {}).get("product_line", {})

    def checks(vals):
        b = []
        mx = spec.get("product_code_max")
        if mx and product_code and len(product_code) > int(mx):
            b.append(f"Product code can be at most {mx} characters.")
        casc = spec.get("cascade") or {}
        if pdp_build and build_phase and pdp_build in casc and build_phase not in casc[pdp_build]:
            b.append(f"Build Phase '{build_phase}' is not valid for PDP Build '{pdp_build}'. Allowed: {casc[pdp_build]}")
        return b
    p = dict(name=name, product_code=product_code, pdp_build=pdp_build, build_phase=build_phase,
             product_area=product_area, product_categories=product_categories)
    return _profile_create("create_product_line", "product_line", p, extra_values, confirm_token, ["name"], checks)

# ======================================================================= MPN

@tool
@tool_errors
def create_mpn(gpn: str, manufacturer: str = "", mpn_id: str = "", description: str = "", gpn_revision: str = "latest",
               confirm_token: str = "") -> dict:
    """Creates a manufacturer part (MPN) and attaches it to a GPN's AML. Custom part: the MPN id = the GPN id.
    Off-the-Shelf part: the user gives the manufacturer's MPN id and manufacturer. An existing MPN is only attached.
    The manufacturer must already exist (else create_manufacturer first). Preview first, then confirm."""
    c = ctx()
    spec = c.cfg.p.get("create", {}).get("mpn")
    rel = c.cfg.get("relations", "aml")
    if not spec or not rel:
        return c.not_available("Creating MPNs / AML")
    args = dict(gpn=gpn, manufacturer=manufacturer, mpn_id=mpn_id, description=description, gpn_revision=gpn_revision,
                confirm_token=confirm_token)
    cos = c.cfg["props"]["custom_off_shelf"]
    mpn_types = c.cfg["types"].get("mpn_types") or [spec["type"]]

    def preview():
        blocking, warn = [], []
        item, rev = c.resolve(gpn, gpn_revision)
        choice = ""
        if cos:
            p = c.tc.get_props([item, rev["ref"]], [cos])
            choice = str(c.ui(p.get(rev["ref"]["uid"], {}), cos) or c.ui(p.get(item["uid"], {}), cos))
        the_id = mpn_id
        if choice == c.cfg["props"]["custom_value"]:
            if mpn_id and mpn_id.upper() != gpn.upper():
                blocking.append(f"{gpn} is Custom, so its MPN id must be the GPN ({gpn}), not {mpn_id}.")
            the_id = gpn
        elif choice == c.cfg["props"]["off_shelf_value"] or not cos:
            if not mpn_id:
                blocking.append("Off-the-Shelf part: ask the user for the manufacturer's MPN id.")
            if not manufacturer:
                blocking.append("Off-the-Shelf part: ask the user for the manufacturer.")
        elif cos:
            blocking.append(f"{gpn} has no Custom / Off-the-Shelf value; set '{cos}' first (update_attributes).")
        mfr = None
        if manufacturer:
            mfr = find_named(c, manufacturer, [c.cfg["types"].get("manufacturer")] if c.cfg["types"].get("manufacturer") else [])
            if not mfr:
                blocking.append(f"Manufacturer '{manufacturer}' does not exist. Create it first with create_manufacturer.")
        existing = None
        if the_id:
            try:
                existing = c.item(the_id, mpn_types)
            except ToolError:
                existing = None
        attached = [r["mpn"] for r in aml(c, [rev["ref"]]).get(rev["ref"]["uid"], [])]
        if the_id and the_id in attached:
            blocking.append(f"MPN {the_id} is already on the AML of {gpn}.")
        pv = {"action": ("Attach existing MPN" if existing else "Create MPN and attach") + f" to {gpn}/{rev['rev']}",
              "gpn": gpn, "gpn_revision": rev["rev"], "custom_or_off_shelf": choice or "n/a", "mpn_id": the_id,
              "manufacturer": manufacturer, "description": description, "existing_mpn": bool(existing), "warnings": warn,
              "_rev": rev["ref"], "_mfr": mfr, "_existing": existing}
        if blocking:
            pv["blocking"] = blocking
        return pv, ""

    def write(pv):
        mpn = pv["_existing"]
        if not mpn:
            vals = {c.P("id"): pv["mpn_id"], c.P("name"): pv["mpn_id"]}
            if description:
                vals[spec.get("map", {}).get("description", c.P("desc"))] = description
            mprop = spec.get("map", {}).get("manufacturer")
            if mprop and pv["_mfr"]:
                vals[mprop] = pv["_mfr"]
            objs = create_object(c, spec["type"], vals, {})
            created = read_back(c, objs)
            mpn = {"uid": created["created"]["uid"], "type": spec["type"]}
            if c.cfg.get("relations", "aml_secondary") == "revision" and created["revision"]:
                mpn = {"uid": created["revision"]["uid"], "type": spec["type"] + "Revision"}
        c.tc.create_relation(rel, pv["_rev"], mpn)
        now = [r["mpn"] for r in aml(c, [pv["_rev"]]).get(pv["_rev"]["uid"], [])]
        return {"attached": pv["mpn_id"], "to": f"{pv['gpn']}/{pv['gpn_revision']}", "aml_now": now, "link": c.link(mpn["uid"])}

    res = two_step("create_mpn", args, preview, write)
    if "preview" in res:
        res["preview"] = {k: v for k, v in res["preview"].items() if not k.startswith("_")}
    return res

# ======================================================================= attachments (upload link + FMS)

_uploads: dict[str, dict] = {}
_upload_server: dict = {}


class _UploadHandler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _store(self):
        uid = self.path.strip("/").split("/")[-1]
        rec = _uploads.get(uid)
        if not rec or rec.get("data") is not None or rec["exp"] < time.time():
            self.send_response(410); self.end_headers(); self.wfile.write(b"Upload link unknown, used or expired.\n"); return
        n = int(self.headers.get("Content-Length") or 0)
        if n > rec["max"]:
            self.send_response(413); self.end_headers(); self.wfile.write(b"File too large.\n"); return
        rec["data"] = self.rfile.read(n)
        self.send_response(200); self.end_headers(); self.wfile.write(b"Uploaded. Tell the assistant to continue.\n")

    do_PUT = _store
    do_POST = _store


def _upload_link(user: str, max_bytes: int) -> tuple[str, str]:
    if not _upload_server:
        srv = ThreadingHTTPServer(("127.0.0.1", 0), _UploadHandler)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        _upload_server["port"] = srv.server_address[1]
    uid = secrets.token_urlsafe(12)
    _uploads[uid] = {"user": user, "exp": time.time() + 600, "max": max_bytes, "data": None}
    return uid, f"http://127.0.0.1:{_upload_server['port']}/upload/{uid}"


@tool
@tool_errors
def create_attachment(target: str, file_name: str, revision: str = "latest", description: str = "",
                      file_size_bytes: int = 0, upload_id: str = "", file_base64: str = "", confirm_token: str = "",
                      file_path: str = "") -> dict:
    """Attaches a file (PDF, Word, Excel, PowerPoint, images, text, zip...) to a part / document / object revision
    (Attachments tab), uploading it like Active Workspace. Step 1: get a one-time upload link (upload the file
    with the given curl command), or send small files (<= 5 MB) inline as file_base64, or give file_path on this PC.
    Step 2: preview. Step 3: confirm -> file attached. Max 25 MB; bigger files: use Active Workspace (link given)."""
    c = ctx()
    lim = c.cfg.p.get("limits", {})
    max_b, inline_b = int(lim.get("upload_max_mb", 25)) * 2**20, int(lim.get("inline_max_mb", 5)) * 2**20
    _, rev = c.resolve(target, revision)
    if file_size_bytes and file_size_bytes > max_b:
        return {"error": "TOO_LARGE", "message": f"Files over {max_b // 2**20} MB must be attached in Active Workspace.",
                "link": c.link(rev["ref"]["uid"])}
    data = None
    if upload_id:
        rec = _uploads.get(upload_id)
        if not rec or rec["user"] != c.user_key() or rec["exp"] < time.time():
            raise ToolError("Upload link unknown or expired. Ask for a new one.", "BAD_INPUT")
        if rec["data"] is None:
            raise ToolError("Nothing has been uploaded to that link yet.", "NOT_READY")
        data = rec["data"]
    elif file_base64:
        data = base64.b64decode(file_base64)
        if len(data) > inline_b:
            raise ToolError(f"Inline files are limited to {inline_b // 2**20} MB. Use an upload link.", "TOO_LARGE")
    elif file_path:
        if not os.path.isfile(file_path):
            raise ToolError(f"No file at {file_path}.", "NOT_FOUND")
        if os.path.getsize(file_path) > max_b:
            return {"error": "TOO_LARGE", "message": f"Over {max_b // 2**20} MB: attach it in Active Workspace.",
                    "link": c.link(rev["ref"]["uid"])}
        with open(file_path, "rb") as f:
            data = f.read()
    if data is None:
        uid, url = _upload_link(c.user_key(), max_b)
        return {"step": "upload", "upload_id": uid, "upload_url": url, "expires_in_seconds": 600,
                "curl": f'curl -X PUT --data-binary "@{file_name}" {url}',
                "next": "Upload the file to upload_url, then call create_attachment again with upload_id."}
    ext = file_name.rsplit(".", 1)[-1].lower() if "." in file_name else ""
    ds_map = c.cfg.p.get("datasets", {})
    ds_type, ref_name = ds_map.get(ext) or ds_map.get("*") or ["MISC", "MISC_File"]
    digest = hashlib.sha256(data).hexdigest()[:16]
    args = dict(target=target, file_name=file_name, revision=revision, description=description, upload_id=upload_id,
                file_sha=digest, confirm_token=confirm_token, file_path=file_path)
    rel = c.cfg.get("relations", "attach_default") or "IMAN_specification"

    def preview():
        existing = [f["name"] for f in attachments(c, rev["ref"])]
        warn = [f"A file named {file_name} is already attached; a second one will be added."] if file_name in existing else []
        return {"action": f"Attach {file_name} ({len(data):,} bytes) to {target}/{rev['rev']}", "dataset_type": ds_type,
                "relation": rel, "description": description, "warnings": warn}, ""

    def write(pv):
        resp = c.tc.call("createDatasets", {"input": [{
            "clientId": "d", "container": c.tc.ref(rev["ref"]), "relationType": rel, "name": file_name,
            "description": description, "type": ds_type, "toolUsed": "",
            "datasetFileInfos": [{"clientId": "f", "fileName": file_name, "namedReferencedName": ref_name,
                                  "isText": ds_type == "Text", "allowReplace": True}]}]})
        c.tc.raise_errors(resp, "createDatasets")
        out = (resp.get("datasetOutput") or [{}])[0]
        ds = out.get("dataset")
        tickets = [t for ci in (out.get("commitInfo") or []) for t in (ci.get("datasetFileTicketInfos") or [])]
        if not ds or not tickets:
            raise ToolError("Teamcenter did not return an upload ticket.", "TC_ERROR")
        t = tickets[0]
        fms = c.cfg.settings.get("TC_FMS_URL") or f"{c.tc.base_url}/fms/fmsupload/"
        up = c.tc.session.post(fms, files={"fmsFile": (file_name, data)}, data={"fmsTicket": t["ticket"]},
                               verify=c.tc.verify_ssl, timeout=300)
        if up.status_code >= 300:
            raise ToolError(f"File upload to FMS failed: HTTP {up.status_code}", "TC_ERROR")
        resp = c.tc.call("commitDatasetFiles", {"commitInput": [{"dataset": ds, "createNewVersion": True,
                                                                 "datasetFileTicketInfos": [t]}]})
        c.tc.raise_errors(resp, "commitDatasetFiles")
        if upload_id:
            _uploads.pop(upload_id, None)
        now = [f["name"] for f in attachments(c, rev["ref"])]
        return {"attached": file_name, "to": f"{target}/{rev['rev']}", "attachments_now": now, "link": c.link(ds["uid"])}
    return two_step("create_attachment", args, preview, write)


# =============================================================================
# 10. CHANGE / BOM / AML / WORKFLOW TOOLS
# =============================================================================
def _strip(res: dict) -> dict:
    if isinstance(res.get("preview"), dict):
        res["preview"] = {k: v for k, v in res["preview"].items() if not k.startswith("_")}
    return res

# ======================================================================= workflow

def _templates(c, rev: dict) -> list[str]:
    resp = c.tc.call("getWorkflowTemplates", {"input": [{"clientId": "t", "getFiltered": True,
                                                         "objects": [c.tc.ref(rev)], "group": ""}]})
    objs = [mo for mo in c.tc.model_objects(resp).values() if "Template" in str(mo.get("type", "")) + str(mo.get("className", ""))]
    if not objs:
        objs = [o for out in (resp.get("templatesOutput") or []) for o in (out.get("workflowTemplates") or [])]
    p = c.tc.get_props(objs, ["template_name", "object_name"])
    return sorted({str(c.ui(x, "template_name") or c.ui(x, "object_name")) for x in p.values()} - {""})


def _open_pnrs_on(c, part_rev: dict) -> list[str]:
    """Open (unreleased) PNRs that have this part as a solution item."""
    t = c.cfg.change_type("pnr")
    sol = c.cfg.get("relations", "solution")
    if not t or not sol:
        return []
    found = c.tc.expand_relations([part_rev], [sol], secondary=True).get(part_rev["uid"], [])
    p = c.tc.get_props([x["other"] for x in found], [c.P("id"), c.P("status")])
    out = []
    for x in found:
        rec = p.get(x["other"]["uid"], {})
        if str(rec.get("type", "")).startswith(t) and not c.ui(rec, c.P("status")):
            out.append(str(c.ui(rec, c.P("id"))))
    return out


@tool
@tool_errors
def submit_change(object_id: str, revision: str = "latest", workflow: str = "", comment: str = "", confirm_token: str = "") -> dict:
    """Submits ANY object (ECO, CR, PRQ, PNR, MCO, part, document ...) to a workflow. The workflows come from the
    system's own filtered list for that object and user (same as Active Workspace 'Submit to Workflow'). Several
    apply -> they are returned and you ask the user which one; one applies -> used. Before submitting, required
    fields and change rules are checked (missing fields are listed by label). Preview first, then confirm;
    afterwards get_workflow_status shows the first tasks."""
    c = ctx()
    _, rev = c.resolve(object_id, revision or "latest")
    r = rev["ref"]
    names = _templates(c, r)
    if not names:
        raise ToolError(f"No workflow can be started on {object_id}/{rev['rev']} by you.", "NO_WORKFLOW")
    if not workflow:
        if len(names) > 1:
            return {"choose_workflow": names, "message": "Several workflows apply. Ask the user which one, then call again with workflow=..."}
        workflow = names[0]
    elif workflow not in names:
        raise ToolError(f"'{workflow}' is not available for this object. Choose one of: {names}", "BAD_INPUT")
    args = dict(object_id=object_id, revision=revision, workflow=workflow, comment=comment, confirm_token=confirm_token)
    kind = c.cfg.change_kind(object_id)

    def preview():
        blocking, warn = [], []
        rules = c.cfg.p.get("rules", {})
        mand = (rules.get("mandatory") or {}).get(kind, [])
        if mand:
            descs = descriptors(c, r.get("type") or rev["rec"].get("type") or "", "EDIT")
            p = c.tc.get_props([r], mand).get(r["uid"], {})
            missing = [descs.get(n, {}).get("label") or _label(n) for n in mand if c.ui(p, n) in ("", [], None)]
            if missing:
                blocking.append("Fill these fields first (update_attributes / set_participant): " + ", ".join(missing))
        if kind:
            items = change_items(c, r)
            if kind == "eco" and not items.get("impacted"):
                blocking.append("An ECO needs at least one impacted item (add_change_items).")
            ctp = rules.get("change_type_prop")
            if ctp and kind == "eco":
                ctype = str(c.ui(c.tc.get_props([r], [ctp]).get(r["uid"], {}), ctp))
                if ctype in (rules.get("nonrev_change_types") or []):
                    bad = [s["id"] for s in items.get("solution", []) if s["status"]]
                    if bad:
                        blocking.append(f"Non-Rev ECO: solution items must have a blank release status (919102): {bad}")
            if rules.get("open_pnr_blocks_submit") and kind == "eco":
                for it in items.get("impacted", []):
                    pnrs = _open_pnrs_on(c, it["_obj"])
                    if pnrs:
                        blocking.append(f"{it['id']} has an open PNR ({', '.join(pnrs)}) (919216).")
        pv = {"action": f"Start '{workflow}' on {object_id}/{rev['rev']}", "comment": comment, "warnings": warn}
        if blocking:
            pv["blocking"] = blocking
        return pv, ""

    def write(pv):
        resp = c.tc.call("createInstance", {"startImmediately": True, "observerKey": "", "name": f"{object_id}/{rev['rev']} - {workflow}",
                                            "subject": "", "description": comment, "contextData": {
                                                "processTemplate": workflow, "attachmentCount": 1, "attachments": [r["uid"]],
                                                "attachmentTypes": [1], "dependencyList": [], "subscribeToEvents": False,
                                                "subscribeEventCount": 0, "eventTypeNames": [], "deadlineDate": "0001-01-01T00:00:00",
                                                "container": "", "processOwner": "", "processAssignmentList": ""}})
        errs = c.tc.errors(resp)
        if errs:
            return {"started": False, "status": "NOT STARTED", "reason": "; ".join(errs)}
        wfs = workflows(c, r)
        return {"started": True, "workflow": workflow, "first_tasks": [t["task"] for w in wfs for t in w["active_tasks"]],
                "diagnosis": diagnose(wfs), "link": c.link(r["uid"])}
    return two_step("submit_change", args, preview, write)

# ======================================================================= attributes

@tool
@tool_errors
def update_attributes(object_id: str, values: dict, revision: str = "latest", confirm_token: str = "") -> dict:
    """Edits fields of ANY object revision (part, ECO, CR, PRQ, PNR, document ...). Keys are the labels users see
    (e.g. 'Change Details', 'Reason Code') or property names; values are checked against the system's field
    definitions and allowed-value lists. Preview shows current -> new, then confirm. Use it to fill fields a
    workflow requires before submit_change. Participants (Analyst ...) go through set_participant."""
    c = ctx()
    _, rev = c.resolve(object_id, revision or "latest")
    r = rev["ref"]
    t = rev["rec"].get("type") or r.get("type")
    args = dict(object_id=object_id, values=values, revision=revision, confirm_token=confirm_token)

    def preview():
        v, blocking, warn = validate(c, t, values, "EDIT", owner=r, required_check=False)
        cur = c.tc.get_props([r], list(v.keys())).get(r["uid"], {})
        changes = {n: {"from": c.ui(cur, n), "to": nv} for n, nv in v.items()}
        pv = {"action": f"Update {object_id}/{rev['rev']}", "changes": changes, "warnings": warn}
        if not changes:
            blocking.append("No fields to change.")
        if blocking:
            pv["blocking"] = blocking
        return pv, fingerprint({n: c.ui(cur, n) for n in v})

    def write(pv):
        c.tc.set_props(r, {n: ch["to"] for n, ch in pv["changes"].items()})
        now = c.tc.get_props([r], list(pv["changes"].keys())).get(r["uid"], {})
        return {"updated": f"{object_id}/{rev['rev']}", "now": {n: c.ui(now, n) for n in pv["changes"]}, "link": c.link(r["uid"])}
    return two_step("update_attributes", args, preview, write)

# ======================================================================= participants

def _people(c, user: str) -> list[dict]:
    """Group memberships of a user (people picker)."""
    try:
        resp = c.tc.call("performSearch", {"searchInput": {
            "providerName": "Awp0ResourceProvider", "maxToLoad": 50, "maxToReturn": 50, "startIndex": 0,
            "searchCriteria": {"searchString": user, "resourceProviderContentType": "Users", "group": "", "role": ""},
            "searchFilterFieldSortType": "Alphabetical", "searchFilterMap": {}, "searchSortCriteria": [], "attributesToInflate": []}})
        objs = [o for o in (resp.get("searchResults") or []) if o.get("uid")]
    except TCError:
        objs = []
    props = c.tc.get_props(objs, ["user", "group", "role", "user_id", "object_string"])
    out = []
    for o in objs:
        p = props.get(o["uid"], {})
        if (p.get("type") or o.get("type")) != "GroupMember":
            continue
        out.append({"ref": {"uid": o["uid"], "type": "GroupMember"}, "user": c.ui(p, "user"),
                    "member": f"{c.ui(p, 'group')}/{c.ui(p, 'role')}"})
    return out


@tool
@tool_errors
def set_participant(object_id: str, role: str, user: str, revision: str = "latest", member: str = "", confirm_token: str = "") -> dict:
    """Adds a participant to a change object (CR, ECO, PRQ, PNR ...), e.g. role 'Analyst' or 'Approver'.
    The role is resolved from the exact participant type, the role configured for the object type, or a unique
    label; otherwise the candidates are listed (never guessed). The person (LDAP / user id) is found with the people
    picker; if they have several group/role memberships, ask the user which one (member='Group/Role').
    Preview, then confirm."""
    c = ctx()
    _, rev = c.resolve(object_id, revision or "latest")
    kind = c.cfg.change_kind(object_id)
    roles = (c.cfg.p.get("participants") or {}).get(kind, {})
    types = set(roles.values())
    if role in types:
        ptype = role
    else:
        hits = [t for lbl, t in roles.items() if lbl.lower() == role.lower()] or \
               [t for lbl, t in roles.items() if role.lower() in lbl.lower()]
        if len(set(hits)) != 1:
            return {"choose_role": sorted(roles.keys()) or "no roles configured for this object type",
                    "message": f"Role '{role}' is ambiguous or unknown for {kind.upper() or 'this object'}. Ask the user which one."}
        ptype = hits[0]
    people = _people(c, user)
    if not people:
        raise ToolError(f"No user '{user}' found in the people picker.", "NOT_FOUND")
    if member:
        people = [p for p in people if p["member"].lower() == member.lower()] or people
    if len(people) > 1 and not member:
        return {"choose_member": [p["member"] for p in people], "message": f"{user} has several memberships; ask which one (member=...)."}
    gm = people[0]
    args = dict(object_id=object_id, role=role, user=user, revision=revision, member=member, confirm_token=confirm_token)

    def preview():
        return {"action": f"Add {gm['user'] or user} ({gm['member']}) as {role} ({ptype}) on {object_id}/{rev['rev']}"}, ""

    def write(pv):
        resp = c.tc.call("addParticipants", {"addParticipantInfo": [{"itemRev": c.tc.ref(rev["ref"]), "participantInfo": [
            {"clientId": "p", "assignee": gm["ref"], "participantType": ptype, "additionalInfo": {}}]}]})
        c.tc.raise_errors(resp, "addParticipants")
        return {"participants_now": participants(c, rev["ref"]), "link": c.link(rev["ref"]["uid"])}
    return two_step("set_participant", args, preview, write)

# ======================================================================= change items

def _rel(c, relation: str) -> str:
    name = c.cfg.get("relations", relation.lower())
    if relation.lower() not in CM_RELATIONS or not name:
        raise ToolError("relation must be impacted, solution, problem or reference.", "BAD_INPUT")
    return name


@tool
@tool_errors
def add_change_items(change_id: str, gpns: list[str], relation: str = "impacted", revision: str = "latest_released",
                     confirm_token: str = "") -> dict:
    """Adds items (up to 25 GPNs) to any change (ECO, MCO, CR, Deviation, PRQ, PNR): relation impacted (default;
    on ECOs the system creates the working solution revision), solution, problem or reference. Pre-checks: no
    temporary/working revisions, MPNs or Mfg-test revisions as impacted; no Obsolete items on Rev ECOs; one revision
    per GPN; the change must be editable; warns when an Approved ECO won't create working revisions.
    Default revision = latest released (use 'latest' for unreleased parts). Preview, then confirm."""
    c = ctx()
    if not gpns or len(gpns) > 25:
        raise ToolError("Give 1 to 25 GPNs.", "BAD_INPUT")
    rel = _rel(c, relation)
    _, crev, kind = c.change(change_id, "latest")
    args = dict(change_id=change_id, gpns=gpns, relation=relation, revision=revision, confirm_token=confirm_token)
    rules = c.cfg.p.get("rules", {})

    def preview():
        blocking, warn, add, already = [], [], [], []
        cp = c.tc.get_props([crev["ref"]], [c.P("status"), "CMMaturity", rules.get("change_type_prop") or c.P("type")]).get(crev["ref"]["uid"], {})
        if crev["status"]:
            blocking.append(f"{change_id} is released ({', '.join(crev['status'])}) and can't be edited.")
        maturity = str(c.ui(cp, "CMMaturity"))
        if kind == "eco" and relation == "impacted" and maturity.lower() == "approved":
            warn.append("This ECO is Approved: adding impacted items will not create working revisions.")
        is_rev_eco = kind == "eco" and str(c.ui(cp, rules.get("change_type_prop", ""))) not in (rules.get("nonrev_change_types") or [])
        current = change_items(c, crev["ref"]).get(relation.lower(), [])
        cur_ids = {x["id"]: x["revision"] for x in current}
        seen = set()
        for g in gpns:
            if g.upper() in seen:
                blocking.append(f"{g} is listed twice (one revision per GPN).")
                continue
            seen.add(g.upper())
            try:
                it, rv = c.resolve(g, revision)
            except ToolError as e:
                blocking.append(f"{g}: {e}")
                continue
            t = (rv["rec"].get("type") or "").replace("Revision", "")
            if g in cur_ids:
                (already if cur_ids[g] == rv["rev"] else blocking).append(
                    f"{g}/{cur_ids[g]}" if cur_ids[g] == rv["rev"] else f"{g} is already on the change as revision {cur_ids[g]} (one revision per GPN).")
                continue
            if relation == "impacted":
                if t in (c.cfg["types"].get("mpn_types") or []):
                    blocking.append(f"{g} is an MPN; MPNs can't be impacted items.")
                    continue
                if t in (c.cfg["types"].get("mfg_test_types") or []):
                    blocking.append(f"{g}/{rv['rev']} is a Mfg-test revision; it can't be an impacted item.")
                    continue
                if not rv["status"] or set(rv["status"]) <= set(rules.get("working_status_values") or []):
                    blocking.append(f"{g}/{rv['rev']} is a working/temporary revision; impacted items must be released "
                                    f"(or pick another revision).")
                    continue
                if is_rev_eco and set(rv["status"]) & set(rules.get("obsolete_status_values") or []):
                    blocking.append(f"{g}/{rv['rev']} is Obsolete; not allowed on a Rev ECO.")
                    continue
            add.append({"gpn": g, "revision": rv["rev"], "status": ", ".join(rv["status"]) or "Working", "_ref": rv["ref"]})
        pv = {"action": f"Add to {change_id} as {relation}", "add": add, "already_on_change": already, "warnings": warn}
        if not add and not blocking:
            blocking.append("Nothing to add.")
        if blocking:
            pv["blocking"] = blocking
        return pv, fingerprint(sorted(cur_ids.items()))

    def write(pv):
        for a in pv["add"]:
            c.tc.create_relation(rel, crev["ref"], a["_ref"])
        items = change_items(c, crev["ref"])
        return {"added": [f"{a['gpn']}/{a['revision']}" for a in pv["add"]],
                f"{relation}_now": [f"{x['id']}/{x['revision']}" for x in items.get(relation.lower(), [])],
                "solution_now": [f"{x['id']}/{x['revision']}" for x in items.get("solution", [])], "link": c.link(crev["ref"]["uid"])}
    res = two_step("add_change_items", args, preview, write)
    if isinstance(res.get("preview"), dict):
        res["preview"]["add"] = [{k: v for k, v in a.items() if not k.startswith("_")} for a in res["preview"].get("add", [])]
    return res


@tool
@tool_errors
def remove_change_items(change_id: str, gpns: list[str], relation: str = "impacted", confirm_token: str = "") -> dict:
    """Removes items from a change's impacted (default), solution, problem or reference list by GPN. On ECOs,
    removing an impacted item makes the system DELETE its working solution revision (with edits, BOM changes,
    attachments): the preview warns; refused if that revision is checked out (919178). Preview, then confirm."""
    c = ctx()
    rel = _rel(c, relation)
    _, crev, kind = c.change(change_id, "latest")
    args = dict(change_id=change_id, gpns=gpns, relation=relation, confirm_token=confirm_token)

    def preview():
        blocking, warn = [], []
        items = change_items(c, crev["ref"])
        cur = {x["id"].upper(): x for x in items.get(relation.lower(), [])}
        sol = {x["id"].upper(): x for x in items.get("solution", [])}
        rm = []
        for g in gpns:
            x = cur.get(g.upper())
            if not x:
                blocking.append(f"{g} is not in the {relation} list of {change_id}.")
                continue
            rm.append({"gpn": x["id"], "revision": x["revision"], "_obj": x["_obj"]})
            if kind == "eco" and relation == "impacted" and g.upper() in sol:
                s = sol[g.upper()]
                co = str(c.ui(c.tc.get_props([s["_obj"]], [c.P("checked_out")]).get(s["_obj"]["uid"], {}), c.P("checked_out")))
                if co.upper() in ("Y", "TRUE", "YES"):
                    blocking.append(f"{g}/{s['revision']} (working solution revision) is checked out (919178).")
                else:
                    warn.append(f"Removing {g} will DELETE its working solution revision {g}/{s['revision']} "
                                "with all its edits, BOM changes and attachments.")
        pv = {"action": f"Remove from {change_id} ({relation})", "remove": rm, "warnings": warn}
        if blocking:
            pv["blocking"] = blocking
        return pv, fingerprint(sorted(cur.keys()))

    def write(pv):
        for x in pv["remove"]:
            c.tc.delete_relation(rel, crev["ref"], x["_obj"])
        items = change_items(c, crev["ref"])
        return {"removed": [f"{x['gpn']}/{x['revision']}" for x in pv["remove"]],
                f"{relation}_now": [f"{x['id']}/{x['revision']}" for x in items.get(relation.lower(), [])]}
    res = two_step("remove_change_items", args, preview, write)
    if isinstance(res.get("preview"), dict):
        res["preview"]["remove"] = [{k: v for k, v in a.items() if not k.startswith("_")} for a in res["preview"].get("remove", [])]
    return res

# ======================================================================= tasks & sign-offs

def _my_signoff(c, task: dict):
    tp = c.tc.get_props([task], ["valid_signoffs", "signoff_attachments"]).get(task["uid"], {})
    sos = [{"uid": u, "type": "Signoff"} for u in (c.db(tp, "valid_signoffs") or c.db(tp, "signoff_attachments")) if u]
    sp = c.tc.get_props(sos, ["group_member", "fnd0Assignee", "decision"])
    me = c.me()
    for s in sos:
        x = sp.get(s["uid"], {})
        who = str(c.ui(x, "group_member") or c.ui(x, "fnd0Assignee"))
        if me["user_id"].lower() in who.lower() or me["user_name"].lower() in who.lower():
            return s, str(c.ui(x, "decision"))
    return None, ""


def _perform(c, task: dict, decision: str, comment: str, signoff: dict | None) -> None:
    wf = c.cfg["workflow"]
    tag = wf.get("comment_tag", "")
    comment = f"{comment} {tag}".strip() if comment else tag
    if signoff:
        action, value, support = wf["perform_signoff_action"], wf["reject" if decision == "reject" else "approve"], signoff
    elif decision in ("yes", "no"):
        action, value, support = wf["complete_action"], wf["yes_value" if decision == "yes" else "no_value"], None
    else:
        action, value, support = wf["complete_action"], wf["complete_value"], None
    resp = c.tc.call("performAction3", {"input": [{"clientId": "a", "actionableObject": c.tc.ref(task), "action": action,
                                                   "password": "", "supportingValue": value,
                                                   "supportingObject": support or NULL,
                                                   "propertyNameValues": {"comments": [comment]}}]})
    c.tc.raise_errors(resp, "performAction3")


def _task_kind(t: dict) -> str:
    tt = str(t.get("task_type", "")).lower()
    if "condition" in tt:
        return "condition"
    if "signoff" in tt or "review" in tt or "acknowledge" in tt or "perform" in tt:
        return "signoff"
    return "do"


@tool
@tool_errors
def act_on_task(object_id: str, decision: str, task: str = "", comment: str = "", confirm_token: str = "") -> dict:
    """Acts on ANY task in your own inbox, matched by its target as listed by get_my_tasks: approve / reject a
    sign-off on any object (manufacturer approvals, PNR / PRQ / CR / ECO reviews; reject needs a reason) or complete
    a Do task. decision: approve / reject / complete / yes / no. Condition tasks have workflow-specific choices: the
    link is returned instead. If several tasks match, ask which (task=...). Preview, then confirm."""
    c = ctx()
    decision = decision.lower().strip()
    if decision not in ("approve", "reject", "complete", "yes", "no"):
        raise ToolError("decision must be approve, reject, complete, yes or no.", "BAD_INPUT")
    if decision == "reject" and not comment.strip():
        raise ToolError("A reject needs a comment with the reason.", "BAD_INPUT")
    tasks = [t for t in inbox(c) if str(t["target"]).upper().startswith(object_id.upper())]
    if task:
        tasks = [t for t in tasks if task.lower() in str(t["task"]).lower()]
    if not tasks:
        raise ToolError(f"No task of yours on {object_id}. Check get_my_tasks.", "NOT_FOUND")
    if len(tasks) > 1:
        return {"choose_task": [t["task"] for t in tasks], "message": "Several tasks match; ask the user which one (task=...)."}
    t = tasks[0]
    kind = _task_kind(t)
    if kind == "condition":
        return {"task": t["task"], "kind": "condition", "link": t["target_link"],
                "message": "This is a condition task with workflow-specific choices; complete it in Active Workspace."}
    args = dict(object_id=object_id, decision=decision, task=task, comment=comment, confirm_token=confirm_token)

    def preview():
        so, dec = _my_signoff(c, t["_task"]) if kind == "signoff" else (None, "")
        blocking = []
        if kind == "signoff" and not so:
            blocking.append("You have no sign-off on this task.")
        if kind == "signoff" and decision not in ("approve", "reject"):
            blocking.append("A sign-off takes approve or reject.")
        if kind == "do" and decision not in ("complete", "yes", "no"):
            blocking.append("A Do task takes complete.")
        pv = {"action": f"{decision.title()} '{t['task']}' on {t['target']}", "workflow": t["workflow"], "comment": comment,
              "target_link": t["target_link"], "_so": so}
        if blocking:
            pv["blocking"] = blocking
        return pv, fingerprint([t["task"], dec])

    def write(pv):
        _perform(c, t["_task"], decision, comment, pv["_so"])
        still = [x for x in inbox(c) if x["_task"]["uid"] == t["_task"]["uid"]]
        return {"recorded": decision, "task": t["task"], "still_in_inbox": bool(still)}
    return _strip(two_step("act_on_task", args, preview, write))


@tool
@tool_errors
def act_on_change(change_id: str, decision: str, comment: str = "", task: str = "", confirm_token: str = "") -> dict:
    """Approve / acknowledge / reject YOUR OWN pending sign-off on an ECO / MCO (CCB Review, Acknowledge, ...).
    Without confirm_token: a PREVIEW (the change's approval card + which task will be signed) and a confirm_token;
    show card_markdown and ask the user. Then call again with the SAME values and the token. Reject needs a comment.
    Refused if the redline or sign-off changed since the preview. task: only if you have several pending sign-offs."""
    c = ctx()
    decision = decision.lower().strip()
    if decision == "acknowledge":
        decision = "approve"
    if decision not in ("approve", "reject"):
        raise ToolError("decision must be approve, acknowledge or reject.", "BAD_INPUT")
    if decision == "reject" and not comment.strip():
        raise ToolError("A reject needs a comment with the reason.", "BAD_INPUT")
    args = dict(change_id=change_id, decision=decision, comment=comment, task=task, confirm_token=confirm_token)

    def preview():
        rl = _redline(c, change_id, "latest", False)
        mine = rl["_mine"]
        if task:
            mine = [m for m in mine if task.lower() in str(m["task"]).lower()]
        blocking = []
        if not mine:
            blocking.append("You have no pending sign-off on this change.")
        elif len(mine) > 1:
            blocking.append("Several pending sign-offs: pass task=... (" + ", ".join(m["task"] for m in mine) + ")")
        so = None
        if len(mine) == 1:
            so, _ = _my_signoff(c, mine[0]["_task"])
            if not so:
                blocking.append("Your sign-off on that task was not found.")
        pv = {"action": f"{decision.title()} {change_id}" + (f" on '{mine[0]['task']}'" if len(mine) == 1 else ""),
              "comment": comment, "card_markdown": rl["card_markdown"], "_task": mine[0]["_task"] if len(mine) == 1 else None, "_so": so}
        if blocking:
            pv["blocking"] = blocking
        return pv, rl["redline_fingerprint"] + ":" + ",".join(sorted(m["task"] for m in mine))

    def write(pv):
        _perform(c, pv["_task"], decision, comment, pv["_so"])
        after = _redline(c, change_id, "latest", False)
        return {"recorded": decision, "change": change_id, "signoff_team_now": after["signoff_team"]}
    return _strip(two_step("act_on_change", args, preview, write))

# ======================================================================= BOM

@tool
@tool_errors
def update_bom(gpn: str, add: list[dict] | None = None, edit: list[dict] | None = None, remove: list[dict] | None = None,
               revision: str = "latest", confirm_token: str = "") -> dict:
    """Adds, edits and removes BOM lines (child components) under a part, normally its ECO working revision
    (default 'latest'). add: [{gpn, revision?, quantity, find_no, ref_designators, notes, function,
    primary_alternate, reference, fatp_input}]; edit: [{gpn and/or find_no, fields to change, new_find_no}];
    remove: [{gpn or find_no}]. All changes of one call are applied in one BOM window and saved together; if any
    step is refused nothing is saved. Pre-checks: Non-Rev ECO working revisions are frozen (919199), allowed
    parent/child types (919110), no duplicate child (919126), no blank quantity (919078); child release status is
    a warning. Stale-check before saving and read-back after. Preview, then confirm."""
    c = ctx()
    add, edit, remove = add or [], edit or [], remove or []
    if not (add or edit or remove):
        raise ToolError("Nothing to do: give add, edit or remove.", "BAD_INPUT")
    _, prev = c.resolve(gpn, revision or "latest")
    pr = prev["ref"]
    bf = bom_fields(c)
    rules = c.cfg.p.get("rules", {})
    args = dict(gpn=gpn, add=add, edit=edit, remove=remove, revision=revision, confirm_token=confirm_token)

    def find_line(lines, spec):
        hits = [l for l in lines if (not spec.get("gpn") or str(l.get("gpn")).upper() == str(spec["gpn"]).upper())
                and (not spec.get("find_no") or str(l.get("find_no")) == str(spec["find_no"]))]
        return hits

    def preview():
        blocking, warn, plan = [], [], []
        sol = c.cfg.get("relations", "solution")
        if sol and rules.get("change_type_prop"):
            for x in c.tc.expand_relations([pr], [sol], secondary=True).get(pr["uid"], []):
                ct = str(c.ui(c.tc.get_props([x["other"]], [rules["change_type_prop"]]).get(x["other"]["uid"], {}), rules["change_type_prop"]))
                if ct in (rules.get("nonrev_change_types") or []):
                    blocking.append(f"{gpn}/{prev['rev']} is the working revision of a Non-Rev ECO; its BOM is frozen (919199).")
        if prev["status"] and not set(prev["status"]) <= set(rules.get("working_status_values") or []):
            blocking.append(f"{gpn}/{prev['rev']} is released ({', '.join(prev['status'])}); edit its ECO working revision instead.")
        lines = bom_lines(c, pr, 1, 5000)
        existing = {str(l.get("gpn")).upper() for l in lines}
        allowed = [tuple(x) for x in (rules.get("bom_allowed_parent_child") or [])]
        ptype = (prev["rec"].get("type") or "").replace("Revision", "")
        seen = set()
        for a in add:
            g = str(a.get("gpn", ""))
            if not g:
                blocking.append("Every add needs a gpn."); continue
            if g.upper() in existing or g.upper() in seen:
                blocking.append(f"{g} is already in the BOM (919126). Use edit to change it."); continue
            seen.add(g.upper())
            if str(a.get("quantity", "")).strip() == "":
                blocking.append(f"{g}: quantity is required (919078)."); continue
            try:
                it, cr = c.resolve(g, a.get("revision") or "latest_released")
            except ToolError as e:
                blocking.append(f"{g}: {e}"); continue
            ctype = (cr["rec"].get("type") or "").replace("Revision", "")
            if allowed and (ptype, ctype) not in allowed:
                blocking.append(f"{ctype} can't be a child of {ptype} (919110)."); continue
            if not cr["status"]:
                warn.append(f"{g}/{cr['rev']} is not released (919010 at Rev ECO submit).")
            props = {bf[k]: str(v) for k, v in a.items() if k in bf and k not in ("gpn", "revision", "description", "status") and v not in (None, "")}
            unknown = [k for k in a if k not in bf and k not in ("gpn", "revision")]
            if unknown:
                warn.append(f"{g}: fields not available on this Teamcenter, ignored: {unknown}")
            plan.append({"op": "add", "gpn": g, "revision": cr["rev"], "set": props, "_item": it, "_rev": cr["ref"]})
        for e in edit:
            hits = find_line(lines, e)
            if len(hits) != 1:
                blocking.append(f"edit {e}: {'no' if not hits else 'several'} matching line(s); give gpn and find_no."); continue
            props = {bf[k]: str(v) for k, v in e.items() if k in bf and k not in ("gpn", "find_no", "revision", "description", "status")}
            if e.get("new_find_no"):
                props[bf["find_no"]] = str(e["new_find_no"])
            if "quantity" in e and str(e["quantity"]).strip() == "":
                blocking.append(f"{hits[0]['gpn']}: quantity can't be blank (919078)."); continue
            if not props:
                blocking.append(f"edit {e}: no fields to change."); continue
            plan.append({"op": "edit", "gpn": hits[0]["gpn"], "find_no": hits[0].get("find_no"),
                         "from": {k: hits[0].get(k) for k, p in bf.items() if p in props}, "set": props})
        for r in remove:
            hits = find_line(lines, r)
            if not hits:
                blocking.append(f"remove {r}: no matching line."); continue
            for h in hits:
                plan.append({"op": "remove", "gpn": h["gpn"], "find_no": h.get("find_no")})
        pv = {"action": f"Change the BOM of {gpn}/{prev['rev']}", "changes": plan, "warnings": warn}
        if blocking:
            pv["blocking"] = blocking
        return pv, fingerprint(public(lines))

    def write(pv):
        window, top = c.tc.open_bom(pr)
        saved = False
        try:
            kids = c.tc.expand_line(top)
            kp = c.tc.get_props([k["line"] for k in kids], [bf["gpn"], bf["find_no"]])
            def line_of(g, fn):
                for k in kids:
                    p = kp.get(k["line"]["uid"], {})
                    if str(c.ui(p, bf["gpn"])).upper() == str(g).upper() and (fn in (None, "") or str(c.ui(p, bf["find_no"])) == str(fn)):
                        return k["line"]
                raise ToolError(f"Line {g} {fn or ''} disappeared; nothing saved.", "STALE")
            adds = [p for p in pv["changes"] if p["op"] == "add"]
            if adds:
                resp = c.tc.call("addOrUpdateChildrenToParentLine", {"inputs": [{"parentLine": c.tc.ref(top), "viewType": "", "items": [
                    {"item": c.tc.ref(a["_item"]), "itemRev": c.tc.ref(a["_rev"]), "bomline": NULL, "itemLineProperties": a["set"]}
                    for a in adds]}]})
                c.tc.raise_errors(resp, "addOrUpdateChildrenToParentLine")
            for p in pv["changes"]:
                if p["op"] == "edit":
                    c.tc.set_props(line_of(p["gpn"], p.get("find_no")), p["set"])
            rms = [line_of(p["gpn"], p.get("find_no")) for p in pv["changes"] if p["op"] == "remove"]
            if rms:
                resp = c.tc.call("removeChildrenFromParentLine", {"bomlines": [c.tc.ref(x) for x in rms]})
                c.tc.raise_errors(resp, "removeChildrenFromParentLine")
            c.tc.close_bom(window, save=True)
            saved = True
        finally:
            if not saved:
                c.tc.close_bom(window, save=False)
        after = public(bom_lines(c, pr, 1, 5000))
        have = {str(l.get("gpn")).upper() for l in after}
        problems = [f"{p['gpn']} not found after save" for p in pv["changes"] if p["op"] == "add" and p["gpn"].upper() not in have] + \
                   [f"{p['gpn']} still present" for p in pv["changes"] if p["op"] == "remove" and p["gpn"].upper() in have]
        return {"saved": True, "verified": not problems, "problems": problems, "bom_now": after, "link": c.link(pr["uid"])}
    res = two_step("update_bom", args, preview, write)
    if isinstance(res.get("preview"), dict):
        res["preview"]["changes"] = [{k: v for k, v in p.items() if not k.startswith("_")} for p in res["preview"].get("changes", [])]
    return res

# ======================================================================= AML

@tool
@tool_errors
def update_aml(gpn: str, mpn: str, values: dict | None = None, detach: bool = False, revision: str = "latest",
               confirm_token: str = "") -> dict:
    """Changes an MPN's entry on a GPN's AML (Manufacturers tab) or detaches the MPN (the MPN itself stays).
    Fields: preferred_status (Preferred / Alternate / NRND / Non Preferred), qual_status (Qualified / Qual in
    Process / Tested / Untested / Disqualified / N/A), allocation, reference_notes, target_milestone, build_phase,
    vendor_documents, plant_codes; values are checked against the live value lists. Warns if no Preferred MPN would
    remain. Preview, then confirm."""
    c = ctx()
    rel = c.cfg.get("relations", "aml")
    if not rel:
        return c.not_available("The AML")
    af = {k: v for k, v in c.cfg["aml_fields"].items() if v}
    _, rev = c.resolve(gpn, revision or "latest")
    args = dict(gpn=gpn, mpn=mpn, values=values, detach=detach, revision=revision, confirm_token=confirm_token)

    def preview():
        blocking, warn = [], []
        rows = aml(c, [rev["ref"]]).get(rev["ref"]["uid"], [])
        row = next((r for r in rows if str(r["mpn"]).upper() == mpn.upper()), None)
        if not row:
            return {"action": "update AML", "blocking": [f"{mpn} is not on the AML of {gpn}/{rev['rev']}. On it: {[r['mpn'] for r in rows]}"]}, ""
        changes = {}
        if not detach:
            unknown = [k for k in (values or {}) if k not in af]
            if unknown:
                blocking.append(f"Unknown AML fields {unknown}. Available here: {sorted(af)}")
            for k, v in (values or {}).items():
                if k in af:
                    allowed = []
                    try:
                        pass
                        allowed = _lov(c, rel, af[k], "EDIT", "", row["_rel"])
                    except TCError:
                        pass
                    if allowed and str(v) not in allowed:
                        blocking.append(f"{k}: '{v}' not allowed. Allowed: {allowed}")
                    changes[af[k]] = {"field": k, "from": row.get(k, ""), "to": v}
            if not changes and not blocking:
                blocking.append("No values to change.")
        remaining_pref = [r for r in rows if r is not row and str(r.get("preferred_status", "")).lower() == "preferred"]
        becomes_pref = str((values or {}).get("preferred_status", row.get("preferred_status", ""))).lower() == "preferred"
        if c.cfg["aml_fields"].get("preferred_status") and not remaining_pref and (detach or not becomes_pref):
            warn.append(f"After this, {gpn} would have no Preferred MPN.")
        pv = {"action": (f"Detach {mpn} from" if detach else f"Update {mpn} on") + f" the AML of {gpn}/{rev['rev']}",
              "changes": list(changes.values()), "warnings": warn, "_row": row, "_props": {p: ch["to"] for p, ch in changes.items()}}
        if blocking:
            pv["blocking"] = blocking
        return pv, fingerprint({k: v for k, v in row.items() if not k.startswith("_")})

    def write(pv):
        row = pv["_row"]
        if detach:
            c.tc.delete_relation(rel, rev["ref"], row["_mpn"])
        else:
            c.tc.set_props(row["_rel"], pv["_props"])
        now = public(aml(c, [rev["ref"]]).get(rev["ref"]["uid"], []))
        return {"done": "detached" if detach else "updated", "aml_now": now}
    return _strip(two_step("update_aml", args, preview, write))


# =============================================================================
# 12. RUN MODES
#     python tc_plm.py                 Claude Desktop / Gemini CLI (stdio)
#     python tc_plm.py --list          print the tool names
#     python tc_plm.py --http          local HTTP  http://127.0.0.1:8765/mcp
#     python tc_plm.py --setup         one-time: get your keys from the Render relay
#     python tc_plm.py --agent         connect this PC to the Render relay (for Gemini)
#     python tc_plm.py --autostart on  start the agent automatically at Windows login (off to stop)
#     python tc_plm.py --relay         the Render side (Render runs this; it never talks to Teamcenter)
# =============================================================================

import inspect
import itertools
import json as _json
import traceback


def _local_server() -> MCPServer:
    """The real tools, talking to Teamcenter from this PC."""
    for name, fn in TOOLS.items():
        mcp.tool()(fn)
    TOOL_NAMES[:] = sorted(TOOLS)
    return mcp


def _env_path() -> str:
    return os.environ.get("TCPLM_SETTINGS") or os.path.join(APP_DIR, ".env")


def _save_env(updates: dict) -> None:
    """Writes keys into .env, keeping every other line as it is."""
    path = _env_path()
    lines = open(path, encoding="utf-8").read().splitlines() if os.path.exists(path) else []
    left = dict(updates)
    for i, line in enumerate(lines):
        k = line.split("=", 1)[0].strip()
        if k in left and not line.lstrip().startswith("#"):
            lines[i] = f"{k}={left.pop(k)}"
    lines += [f"{k}={v}" for k, v in left.items()]
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ----------------------------------------------------------------- agent (your PC)

def _log(msg: str) -> None:
    print(time.strftime("%H:%M:%S"), msg, flush=True)


def _run_job(job: dict) -> dict:
    name, args = job.get("tool"), job.get("args") or {}
    fn = TOOLS.get(name)
    if fn is None:
        return {"id": job["id"], "ok": False, "error": f"Unknown tool {name}. Update tc_plm.py on Render and on this PC."}
    try:
        result = fn(**args)
        _json.dumps(result, default=str)
        return {"id": job["id"], "ok": True, "result": result}
    except Exception as e:
        return {"id": job["id"], "ok": False, "error": f"{type(e).__name__}: {e}"}


def run_agent() -> None:
    s = load_settings()
    url, key = s.get("RELAY_URL", "").rstrip("/"), s.get("AGENT_KEY", "")
    if not url or not key:
        sys.exit("No RELAY_URL / AGENT_KEY in .env. Run:  python tc_plm.py --setup")
    if not s.get("TC_URL"):
        sys.exit("No TC_URL in .env. Add your Teamcenter address, user and password first.")
    TOOL_NAMES[:] = sorted(TOOLS)
    http = requests.Session()        # honours HTTPS_PROXY if your company needs one
    http.headers["Authorization"] = f"Bearer {key}"
    _log(f"Agent started ({len(TOOLS)} tools, profile {s.get('TCPLM_PROFILE', 'znext')}). Connecting out to {url} ...")
    backoff, connected = 2, False
    while True:
        try:
            r = http.get(f"{url}/agent/next", params={"wait": 25}, timeout=40)
            if r.status_code == 401:
                sys.exit("The relay refused AGENT_KEY (wrong, revoked, or MASTER_SECRET changed). Run --setup again.")
            r.raise_for_status()
            if not connected:
                _log("Connected. Waiting for requests from Gemini / Claude ...")
                connected = True
            backoff = 2
            if r.status_code == 204:
                continue
            job = r.json()
            shown = ", ".join(f"{k}={v!r}"[:60] for k, v in (job.get("args") or {}).items() if k not in ("file_base64",))
            _log(f"-> {job['tool']}({shown})")
            reply = _run_job(job)
            _log(f"<- {'ok' if reply['ok'] else 'ERROR: ' + reply['error']}")
            http.post(f"{url}/agent/result", data=_json.dumps(reply, default=str),
                      headers={"Content-Type": "application/json"}, timeout=60).raise_for_status()
        except KeyboardInterrupt:
            _log("Stopped.")
            return
        except requests.RequestException as e:
            connected = False
            _log(f"Connection problem ({type(e).__name__}). Retrying in {backoff}s ...")
            time.sleep(backoff)
            backoff = min(backoff * 2, 60)
        except Exception:
            traceback.print_exc()
            time.sleep(5)


def run_setup() -> None:
    s = load_settings()
    print("One-time setup for Gemini (via your Render relay).\n")
    url = (input(f"Relay address [{s.get('RELAY_URL', 'https://....onrender.com')}]: ").strip()
           or s.get("RELAY_URL", "")).rstrip("/")
    name = input("Your name (letters only, e.g. ravi): ").strip() or "user"
    invite = input("Invite code: ").strip()
    try:
        r = requests.post(f"{url}/register", json={"name": name, "invite": invite}, timeout=60)
        d = r.json()
    except Exception as e:
        sys.exit(f"Could not reach {url} ({e}). Is the Render service awake? Open it in a browser once and retry.")
    if r.status_code != 200:
        sys.exit(f"Relay said: {d.get('error', r.text)}")
    _save_env({"RELAY_URL": url, "AGENT_KEY": d["agent_key"], "AGENT_ID": d["id"]})
    gem = f"{url}/mcp?key={d['mcp_key']}"
    print(f"\nDone. Your id: {d['id']}  (saved to {_env_path()})")
    print("\nYour Gemini MCP address (keep it private, it works like a password):\n")
    print("   " + gem + "\n")
    print("Next:  python tc_plm.py --agent      (leave it running)")
    print("       python tc_plm.py --autostart on   (optional: start it automatically at login)")


def run_autostart(on: bool) -> None:
    if os.name != "nt":
        sys.exit("--autostart is for Windows. On other systems add  python tc_plm.py --agent  to your login items.")
    import winreg
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE)
    if on:
        pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        exe = pyw if os.path.exists(pyw) else sys.executable
        log = os.path.join(APP_DIR, "tc-plm-agent.log")
        cmd = f'"{exe}" "{os.path.abspath(__file__)}" --agent --log "{log}"'
        winreg.SetValueEx(key, "TC-PLM-Agent", 0, winreg.REG_SZ, cmd)
        print(f"The agent will start by itself at every login (no window). Log: {log}\nStarting it now ...")
        import subprocess
        subprocess.Popen([exe, os.path.abspath(__file__), "--agent", "--log", log],
                         creationflags=0x00000008 | 0x00000200)   # DETACHED_PROCESS | NEW_PROCESS_GROUP
    else:
        try:
            winreg.DeleteValue(key, "TC-PLM-Agent")
        except FileNotFoundError:
            pass
        print("Auto-start removed. (A running agent keeps running until you sign out or end pythonw.exe.)")


def _single_instance() -> None:
    """Stops a second agent on the same PC from stealing jobs from the first."""
    import socket
    global _LOCK
    _LOCK = socket.socket()
    try:
        _LOCK.bind(("127.0.0.1", int(os.environ.get("TCPLM_LOCK_PORT", "47614"))))
    except OSError:
        sys.exit("An agent is already running on this PC.")


# ----------------------------------------------------------------- relay (Render)

def run_relay() -> None:
    import asyncio
    import base64
    import contextvars
    import hmac
    import re as _re
    from urllib.parse import parse_qs

    import uvicorn
    from starlette.requests import Request
    from starlette.responses import JSONResponse, PlainTextResponse, Response

    E = os.environ
    TOOL_TIMEOUT = float(E.get("TOOL_TIMEOUT", "120"))
    STALE, PICKUP = 40, 35
    SECRET, INVITE = E.get("MASTER_SECRET", ""), E.get("INVITE_CODE", "")
    REVOKED = {x.strip() for x in E.get("REVOKED", "").split(",") if x.strip()}
    ID_RE = _re.compile(r"^[a-z0-9_-]{1,40}$")
    if not SECRET:
        print("WARNING: MASTER_SECRET is not set; nobody can connect.", flush=True)

    def sign(kind, uid):
        mac = hmac.new(SECRET.encode(), f"{kind}:{uid}".encode(), hashlib.sha256).digest()
        return base64.urlsafe_b64encode(mac).decode().rstrip("=")[:32]

    def who(token, kind):
        if not SECRET or token.count(".") != 2:
            return None
        k, uid, sig = token.split(".")
        if k != kind or not ID_RE.match(uid) or uid in REVOKED:
            return None
        return uid if hmac.compare_digest(sig, sign(kind, uid)) else None

    current = contextvars.ContextVar("current_user", default="")
    queues, waiting, picked, last_seen = {}, {}, {}, {}
    ids = itertools.count(1)
    online = lambda u: time.time() - last_seen.get(u, 0) < STALE
    q = lambda u: queues.setdefault(u, asyncio.Queue())

    async def relay(tool, args):
        user = current.get()
        if not online(user):
            return {"error": "AGENT_OFFLINE", "message": "Your Teamcenter agent is offline. On your PC run:  python tc_plm.py --agent"}
        jid = str(next(ids))
        fut = asyncio.get_running_loop().create_future()
        waiting[jid], picked[jid] = (user, fut), asyncio.Event()
        await q(user).put({"id": jid, "tool": tool, "args": args})
        try:
            try:
                await asyncio.wait_for(picked[jid].wait(), PICKUP)
            except asyncio.TimeoutError:
                last_seen.pop(user, None)
                return {"error": "AGENT_OFFLINE", "message": "Your agent did not pick up the request. Start  python tc_plm.py --agent  and retry."}
            reply = await asyncio.wait_for(fut, TOOL_TIMEOUT)
        except asyncio.TimeoutError:
            return {"error": "TIMEOUT", "message": f"No answer from your Teamcenter within {int(TOOL_TIMEOUT)} s."}
        finally:
            waiting.pop(jid, None)
            picked.pop(jid, None)
        if not reply.get("ok"):
            return {"error": "AGENT_ERROR", "message": reply.get("error", "Agent error")}
        res = reply.get("result")
        return {"count": len(res), "results": res} if isinstance(res, list) else res

    # Same 34 tools, same names, parameters and descriptions as the local server, built from this file.
    rmcp = MCPServer("teamcenter-plm", instructions=mcp.instructions)

    def proxy_for(name, fn):
        async def proxy(**kwargs):
            return await relay(name, kwargs)
        proxy.__name__, proxy.__doc__ = name, fn.__doc__
        proxy.__signature__ = inspect.signature(fn)
        proxy.__annotations__ = dict(getattr(fn, "__annotations__", {}))
        proxy.__module__ = fn.__module__
        return proxy

    for n, f in TOOLS.items():
        rmcp.tool()(proxy_for(n, f))

    @rmcp.custom_route("/agent/next", methods=["GET"])
    async def agent_next(request: Request) -> Response:
        user = current.get()
        last_seen[user] = time.time()
        try:
            job = await asyncio.wait_for(q(user).get(), min(float(request.query_params.get("wait", "25")), 50))
        except asyncio.TimeoutError:
            last_seen[user] = time.time()
            return Response(status_code=204)
        if job["id"] not in waiting:
            return Response(status_code=204)
        if await request.is_disconnected():
            q(user).put_nowait(job)
            last_seen.pop(user, None)
            return Response(status_code=204)
        picked[job["id"]].set()
        return JSONResponse(job)

    @rmcp.custom_route("/agent/result", methods=["POST"])
    async def agent_result(request: Request) -> Response:
        user = current.get()
        last_seen[user] = time.time()
        data = await request.json()
        entry = waiting.get(str(data.get("id")))
        if entry and entry[0] == user and not entry[1].done():
            entry[1].set_result(data)
        return JSONResponse({"ok": True})

    @rmcp.custom_route("/register", methods=["POST"])
    async def register(request: Request) -> Response:
        try:
            data = await request.json()
        except Exception:
            return JSONResponse({"error": "Send JSON: {name, invite}"}, status_code=400)
        if not SECRET or not INVITE:
            return JSONResponse({"error": "Sign-up is not enabled (set MASTER_SECRET and INVITE_CODE on Render)."}, status_code=403)
        if not hmac.compare_digest(str(data.get("invite", "")), INVITE):
            await asyncio.sleep(1.5)
            return JSONResponse({"error": "Wrong invite code."}, status_code=403)
        base = _re.sub(r"[^a-z0-9_-]", "-", str(data.get("name", "")).strip().lower())[:24].strip("-") or "user"
        uid = f"{base}-{secrets.token_hex(3)}"
        return JSONResponse({"id": uid, "mcp_key": f"mcp.{uid}.{sign('mcp', uid)}",
                             "agent_key": f"agent.{uid}.{sign('agent', uid)}"})

    @rmcp.custom_route("/", methods=["GET"])
    async def home(request: Request) -> Response:
        n = sum(1 for u in list(last_seen) if online(u))
        return PlainTextResponse(f"Teamcenter PLM relay ({len(TOOLS)} tools) is running. Agents online: {n}. MCP endpoint: /mcp")

    def token_of(scope):
        h = {k.decode().lower(): v.decode() for k, v in scope.get("headers", [])}
        a = h.get("authorization", "")
        if a.lower().startswith("bearer "):
            return a[7:].strip()
        return (parse_qs(scope.get("query_string", b"").decode()).get("key") or [""])[0]

    inner = rmcp.streamable_http_app(host="0.0.0.0", stateless_http=True)

    async def app(scope, receive, send):
        if scope["type"] != "http" or scope.get("path", "") in ("", "/", "/register"):
            return await inner(scope, receive, send)
        user = who(token_of(scope), "agent" if scope["path"].startswith("/agent/") else "mcp")
        if not user:
            await send({"type": "http.response.start", "status": 401, "headers": [(b"content-type", b"text/plain")]})
            return await send({"type": "http.response.body", "body": b"Missing, wrong or revoked key."})
        tok = current.set(user)
        try:
            return await inner(scope, receive, send)
        finally:
            current.reset(tok)

    uvicorn.run(app, host="0.0.0.0", port=int(E.get("PORT", "8000")), proxy_headers=True, forwarded_allow_ips="*")


# ----------------------------------------------------------------- main

def main() -> None:
    a = sys.argv[1:]
    if "--log" in a:
        i = a.index("--log")
        f = open(a[i + 1], "a", buffering=1, encoding="utf-8")
        sys.stdout = sys.stderr = f
    if "--relay" in a:
        return run_relay()
    if "--setup" in a:
        return run_setup()
    if "--autostart" in a:
        i = a.index("--autostart")
        return run_autostart((a[i + 1:] or ["on"])[0].lower() != "off")
    if "--agent" in a:
        _single_instance()
        return run_agent()
    server = _local_server()
    if "--list" in a:
        print("\n".join(TOOL_NAMES), f"\n{len(TOOL_NAMES)} tools")
        return
    if "--http" in a:
        import uvicorn
        uvicorn.run(server.streamable_http_app(host="127.0.0.1", stateless_http=True),
                    host="127.0.0.1", port=int(os.environ.get("TCPLM_PORT", "8765")))
        return
    server.run()


if __name__ == "__main__":
    main()
