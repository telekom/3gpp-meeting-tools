from modules.assistant.core.tool_registry import ArgumentType, ToolArgument, ToolDefinition
def build_spec_search_tools(service):
    return (
        ToolDefinition("search_spec_text","Search locally indexed 3GPP text; use get_spec_clause on a returned result_id.",(
            ToolArgument("query",ArgumentType.STRING,"Concise 3GPP phrase/terms.",True,max_length=500),
            ToolArgument("specification",ArgumentType.STRING,"Optional specification number.",max_length=40),
            ToolArgument("version",ArgumentType.STRING,"Optional exact indexed version.",max_length=40),
            ToolArgument("release",ArgumentType.INTEGER,"Optional Release number.",minimum=1,maximum=99),
            ToolArgument("clause",ArgumentType.STRING,"Optional clause-prefix filter.",max_length=80),
            ToolArgument("limit",ArgumentType.INTEGER,"Maximum compact matches.",default=6,minimum=1,maximum=20)),service.search_text),
        ToolDefinition("get_spec_clause","Read the exact clause represented by a search_spec_text result_id.",(
            ToolArgument("result_id",ArgumentType.STRING,"Opaque S-... result ID.",True,max_length=80),),service.get_clause),
    )
