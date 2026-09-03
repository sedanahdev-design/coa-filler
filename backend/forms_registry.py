"""Registry of the customer document-forms (Feature 3). Each entry maps a
form_id to a human name and a `generate(customer, shipment, output_dir)`
function that returns {doc_key: Path}. Forms not yet implemented are listed
with generate=None so the frontend can show them as "coming soon" instead of
erroring.
"""
from __future__ import annotations

from typing import Callable, Dict, Optional

import forms_biesterfeld_dubai
import forms_biesterfeld_germany
import forms_sedanah_jordan
import forms_thosco
import forms_vigorous

FORMS: Dict[str, dict] = {
    "vigorous": {
        "name": "Vigorous Form",
        "docs": ["Invoice", "Packing List", "Certificate of Origin"],
        "generate": forms_vigorous.generate,
    },
    "thosco": {
        "name": "Thosco",
        "docs": ["Invoice", "Packing List", "Certificate of Origin"],
        "generate": forms_thosco.generate,
    },
    "biesterfeld_germany": {
        "name": "Biesterfeld Germany",
        "docs": ["Invoice", "Packing List", "Certificate of Origin"],
        "generate": forms_biesterfeld_germany.generate,
    },
    "biesterfeld_dubai": {
        "name": "Biesterfeld Dubai",
        "docs": ["Invoice Draft", "Packing Draft"],
        "generate": forms_biesterfeld_dubai.generate,
    },
    "sedanah_jordan": {
        "name": "Sedanah Jordan",
        "docs": ["Invoice", "Packing List"],
        "generate": forms_sedanah_jordan.generate,
    },
}


def list_forms() -> list:
    return [
        {"id": fid, "name": f["name"], "docs": f["docs"], "available": f["generate"] is not None}
        for fid, f in FORMS.items()
    ]


def get_generator(form_id: str) -> Optional[Callable]:
    entry = FORMS.get(form_id)
    return entry["generate"] if entry else None
