"""Clean Road Intersection — the ArcGIS-built file vs the site's export (the
Reports vs ArcGIS matrix's Clean Road Intersection rows). The comparison lives
in `compare_clean_road_arcgis`; this module is the registry's handle on it."""
import arcgis_report_clean_intersection as _build
from compare_clean_road_arcgis import SIDE_A, SIDE_B, CleanRoadArcgisCompare

_COMPARE = CleanRoadArcgisCompare(_build)

REPORT_NAME = _COMPARE.REPORT_NAME
file_a_label = SIDE_A
file_b_label = SIDE_B
SHARED_HEADER = list(_build.SPEC.header)
CONTEXT_FIELDS = _COMPARE.context
EXPORT_CONTEXT_FIELDS = tuple(c for c in CONTEXT_FIELDS
                              if c not in _build.CONTEXT_COLUMNS)
compare = _COMPARE.compare
suggest_name = _COMPARE.suggest_name
_load_pair = _COMPARE._load_pair
