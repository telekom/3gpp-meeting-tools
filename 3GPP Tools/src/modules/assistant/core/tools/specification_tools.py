from modules.assistant.core.tool_registry import ArgumentType, ToolArgument, ToolDefinition
def build_specification_tools(service):
    return (
        ToolDefinition("find_specifications","Search the local 3GPP specification catalogue by number, title, or topic.",(
            ToolArgument("query",ArgumentType.STRING,"Specification number, title fragment, or topic.",True,max_length=300),
            ToolArgument("limit",ArgumentType.INTEGER,"Maximum candidates.",default=8,minimum=1,maximum=20)),service.find_specifications),
        ToolDefinition("get_specification","Get catalogue metadata and known versions for one 3GPP specification.",(
            ToolArgument("specification",ArgumentType.STRING,"Specification number such as 29.244.",True,max_length=40),),service.get_specification),
    )
