# WordprocessingML XML Namespaces
import re

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TAG_BODY = f"{W_NS}body"
TAG_P = f"{W_NS}p"
TAG_TBL = f"{W_NS}tbl"
TAG_TR = f"{W_NS}tr"
TAG_TC = f"{W_NS}tc"
TAG_T = f"{W_NS}t"
TAG_TAB = f"{W_NS}tab"
TAG_BR = f"{W_NS}br"
TAG_CR = f"{W_NS}cr"
TAG_HYPHEN = f"{W_NS}noBreakHyphen"
TAG_TCPR = f"{W_NS}tcPr"
TAG_GRIDSPAN = f"{W_NS}gridSpan"
TAG_VMERGE = f"{W_NS}vMerge"
TAG_PPR = f"{W_NS}pPr"
TAG_JC = f"{W_NS}jc"
RE_PART_INDEX = re.compile(r"_(\d+)_")
RE_SPEC_NUMBER = re.compile(r"(24|25|36|37|38)[._]?(301|501|331|413|423|412|473|463)")
RE_VERSION_STEM = re.compile(r"-([a-zA-Z0-9]{3})(?:_\d+.*)?$")
