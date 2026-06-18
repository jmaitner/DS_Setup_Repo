#!/usr/bin/env python3
"""
ds_schema.py  —  Single source of truth for converting between the
DS "DS Only" setup-sheet rows, the on-disk product JSON files, and the
flat product dict that ds_automation.py's generator functions expect.

Three public functions:

  sheet_row_to_product(ws, row)      DS Only row  -> nested product JSON dict
  product_to_automation_dict(prod)   product JSON -> flat dict (ds_automation)
  brand_slug(brand)                  "First4 Figures" -> "first4-figures"

The column map (the DS class) is imported from the vendored
ds_automation.py so the catalog and the transformer never drift apart.
"""

import re
from ds_automation import DS


# ──────────────────────────────────────────────────────────────
#  Value cleaning
# ──────────────────────────────────────────────────────────────

def _clean(value):
    """Blank cells -> None.  Strip strings.  Leave numbers as-is."""
    if value is None:
        return None
    if isinstance(value, str):
        s = value.strip()
        return s if s else None
    return value


def brand_slug(brand):
    """'First4 Figures' -> 'first4-figures'.  Used for folder names."""
    if not brand:
        return "unknown-brand"
    s = str(brand).strip().lower()
    s = re.sub(r"[^a-z0-9\s-]", "", s)
    s = re.sub(r"[\s_]+", "-", s)
    s = re.sub(r"-+", "-", s)
    return s.strip("-") or "unknown-brand"


# ──────────────────────────────────────────────────────────────
#  DS Only row  ->  nested product JSON
# ──────────────────────────────────────────────────────────────

def sheet_row_to_product(ws, row):
    """
    Read one data row of the 'DS Only' sheet (1-based row index) and
    return a nested product dict ready to serialise as JSON.
    Captures ALL 81 template columns for full fidelity.
    """
    def g(col):
        return _clean(ws.cell(row=row, column=col).value)

    image_urls = [g(c) for c in (
        DS.IMAGE_URL_1, DS.IMAGE_URL_2, DS.IMAGE_URL_3,
        DS.IMAGE_URL_4, DS.IMAGE_URL_5,
    )]
    image_urls = [u for u in image_urls if u]

    return {
        "ds_number": str(g(DS.DS_NUMBER)) if g(DS.DS_NUMBER) is not None else None,
        "brand": g(DS.BRAND),
        "vendor_item_number": g(DS.VENDOR_ITEM_NUMBER),
        "product_name": g(DS.PRODUCT_NAME),

        "identity": {
            "id_type": g(DS.ID_TYPE),
            "upc": str(g(DS.UPC)) if g(DS.UPC) is not None else None,
        },

        "pricing": {
            "drop_ship_cost": g(DS.DROP_SHIP_COST),
            "wholesale_cost": g(DS.WHOLESALE_COST),
            "ds_cost_domestic": g(DS.DS_COST_DOMESTIC),
            "import_cost": g(DS.IMPORT_COST),
            "msrp": g(DS.MSRP),
        },

        "sourcing": {
            "fob_point": g(DS.FOB_POINT),
            "country_of_origin": g(DS.COUNTRY_OF_ORIGIN),
            "harmonized_code": g(DS.HARMONIZED_CODE),
        },

        "content": {
            "bullets": [
                g(DS.BULLET_1), g(DS.BULLET_2), g(DS.BULLET_3),
                g(DS.BULLET_4), g(DS.BULLET_5),
            ],
            "keywords": g(DS.KEYWORDS),
            "description": g(DS.DESCRIPTION),
        },

        "attributes": {
            "material": g(DS.MATERIAL),
            "num_pieces": g(DS.NUM_PIECES),
            "whats_in_box": g(DS.WHATS_IN_BOX),
            "primary_color": g(DS.PRIMARY_COLOR),
            "secondary_color": g(DS.SECONDARY_COLOR),
            "min_age": g(DS.MIN_AGE),
            "max_age": g(DS.MAX_AGE),
            "gender": g(DS.GENDER),
            "assembly_required": g(DS.ASSEMBLY_REQUIRED),
            "assembly_instructions": g(DS.ASSEMBLY_INSTRUCTIONS),
        },

        "images": {
            "availability": g(DS.IMAGE_AVAILABILITY),
            "urls": image_urls,
        },

        "dimensions": {
            "item": {
                "weight": g(DS.ITEM_WEIGHT), "weight_unit": g(DS.ITEM_WEIGHT_UNIT),
                "length": g(DS.ITEM_LENGTH), "length_unit": g(DS.ITEM_LENGTH_UNIT),
                "width": g(DS.ITEM_WIDTH), "width_unit": g(DS.ITEM_WIDTH_UNIT),
                "height": g(DS.ITEM_HEIGHT), "height_unit": g(DS.ITEM_HEIGHT_UNIT),
            },
            "package": {
                "weight": g(DS.PKG_WEIGHT), "weight_unit": g(DS.PKG_WEIGHT_UNIT),
                "length": g(DS.PKG_LENGTH), "length_unit": g(DS.PKG_LENGTH_UNIT),
                "width": g(DS.PKG_WIDTH), "width_unit": g(DS.PKG_WIDTH_UNIT),
                "height": g(DS.PKG_HEIGHT), "height_unit": g(DS.PKG_HEIGHT_UNIT),
            },
            "master_case": {
                "case_qty": g(DS.CASE_QTY),
                "length": g(DS.MC_LENGTH), "length_unit": g(DS.MC_LENGTH_UNIT),
                "width": g(DS.MC_WIDTH), "width_unit": g(DS.MC_WIDTH_UNIT),
                "height": g(DS.MC_HEIGHT), "height_unit": g(DS.MC_HEIGHT_UNIT),
                "weight": g(DS.MC_WEIGHT), "weight_unit": g(DS.MC_WEIGHT_UNIT),
            },
        },

        "compliance": {
            "choking_hazard": g(DS.CHOKING_HAZARD),
            "lead_phthalates": g(DS.LEAD_PHTHALATES),
            "warranty_included": g(DS.WARRANTY_INCLUDED),
            "warranty_desc": g(DS.WARRANTY_DESC),
            "batteries_required": g(DS.BATTERIES_REQUIRED),
            "batteries_included": g(DS.BATTERIES_INCLUDED),
            "battery_cell_comp": g(DS.BATTERY_CELL_COMP),
            "battery_type_qty": g(DS.BATTERY_TYPE_QTY),
            "packaging_type": g(DS.PACKAGING_TYPE),
            "compliance_cert": g(DS.COMPLIANCE_CERT),
            "doc": g(DS.DOC),
            "sds": g(DS.SDS),
            "sds_url": g(DS.SDS_URL),
            "cpsia": g(DS.CPSIA),
            "test_reports": g(DS.TEST_REPORTS),
            "cpc": g(DS.CPC),
            "product_pics": g(DS.PRODUCT_PICS),
            "instructions": g(DS.INSTRUCTIONS),
            "letter_of_compliance": g(DS.LETTER_OF_COMPLIANCE),
        },

        "_meta": {
            "source_file": None,
            "imported_at": None,
            "supplier_id": None,
            "last_verified": None,
        },
    }


# ──────────────────────────────────────────────────────────────
#  product JSON  ->  flat dict for ds_automation generators
# ──────────────────────────────────────────────────────────────

def product_to_automation_dict(prod):
    """
    Flatten a nested product JSON dict into the exact key shape that
    ds_automation.read_template() produces, so the generator functions
    (generate_new_product, generate_shopify_xlsx, ...) accept it directly.
    """
    ident = prod.get("identity", {}) or {}
    pricing = prod.get("pricing", {}) or {}
    sourcing = prod.get("sourcing", {}) or {}
    content = prod.get("content", {}) or {}
    bullets = (content.get("bullets") or []) + [None] * 5
    attrs = prod.get("attributes", {}) or {}
    images = (prod.get("images", {}) or {}).get("urls", []) or []
    images = images + [None] * 5
    dims = prod.get("dimensions", {}) or {}
    item = dims.get("item", {}) or {}
    pkg = dims.get("package", {}) or {}
    mc = dims.get("master_case", {}) or {}
    comp = prod.get("compliance", {}) or {}

    return {
        "ds_number": prod.get("ds_number"),
        "vendor_item_number": prod.get("vendor_item_number"),
        "id_type": ident.get("id_type"),
        "upc": ident.get("upc"),
        "brand": prod.get("brand"),
        "product_name": prod.get("product_name"),
        "cost": pricing.get("wholesale_cost"),
        "msrp": pricing.get("msrp"),
        "drop_ship_cost": pricing.get("drop_ship_cost"),
        "country_of_origin": sourcing.get("country_of_origin"),

        "description": content.get("description"),
        "bullet1": bullets[0], "bullet2": bullets[1], "bullet3": bullets[2],
        "bullet4": bullets[3], "bullet5": bullets[4],
        "keywords": content.get("keywords"),

        "material": attrs.get("material"),
        "num_pieces": attrs.get("num_pieces"),
        "primary_color": attrs.get("primary_color"),
        "min_age": attrs.get("min_age"),
        "max_age": attrs.get("max_age"),
        "assembly_required": attrs.get("assembly_required"),

        "image_url_1": images[0], "image_url_2": images[1],
        "image_url_3": images[2], "image_url_4": images[3],
        "image_url_5": images[4],

        "item_weight": item.get("weight"), "item_weight_unit": item.get("weight_unit"),
        "item_length": item.get("length"), "item_width": item.get("width"),
        "item_height": item.get("height"),

        "pkg_weight": pkg.get("weight"), "pkg_weight_unit": pkg.get("weight_unit"),
        "pkg_length": pkg.get("length"), "pkg_length_unit": pkg.get("length_unit"),
        "pkg_width": pkg.get("width"), "pkg_width_unit": pkg.get("width_unit"),
        "pkg_height": pkg.get("height"), "pkg_height_unit": pkg.get("height_unit"),

        "case_qty": mc.get("case_qty"),
        "mc_length": mc.get("length"), "mc_length_unit": mc.get("length_unit"),
        "mc_width": mc.get("width"), "mc_width_unit": mc.get("width_unit"),
        "mc_height": mc.get("height"), "mc_height_unit": mc.get("height_unit"),
        "mc_weight": mc.get("weight"), "mc_weight_unit": mc.get("weight_unit"),

        "choking_hazard": comp.get("choking_hazard"),
        "batteries_required": comp.get("batteries_required"),
        "batteries_included": comp.get("batteries_included"),
        "battery_cell_comp": comp.get("battery_cell_comp"),
        "battery_type_qty": comp.get("battery_type_qty"),
    }
