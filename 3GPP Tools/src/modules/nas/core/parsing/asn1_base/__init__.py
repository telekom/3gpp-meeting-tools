"""Modular ASN.1 base parser package."""

from .base_parser import BaseAsn1DocxParser
from .scanner import RE_ASN1_BOX_CAPTION

__all__ = ["BaseAsn1DocxParser", "RE_ASN1_BOX_CAPTION"]
