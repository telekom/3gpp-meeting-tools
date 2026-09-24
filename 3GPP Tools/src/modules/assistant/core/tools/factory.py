from pathlib import Path
from modules.assistant.core.knowledge import ProtocolKnowledgeService, SpecSearchKnowledgeService, SpecificationKnowledgeService
from modules.assistant.core.tool_registry import ToolRegistry
from modules.assistant.core.tools.protocol_tools import build_protocol_tools
from modules.assistant.core.tools.spec_search_tools import build_spec_search_tools
from modules.assistant.core.tools.specification_tools import build_specification_tools

def build_read_only_tool_registry(*, specs_db_path: Path, spec_search_db_path: Path, protocol_db_path: Path) -> ToolRegistry:
    registry=ToolRegistry()
    services=(SpecificationKnowledgeService(specs_db_path),SpecSearchKnowledgeService(spec_search_db_path),ProtocolKnowledgeService(protocol_db_path))
    for d in (*build_specification_tools(services[0]),*build_spec_search_tools(services[1]),*build_protocol_tools(services[2])):
        registry.register(d)
    return registry
