from modules.assistant.core.tool_registry import ArgumentType, ToolArgument, ToolDefinition
def build_protocol_tools(service):
    return (
        ToolDefinition("list_protocols","List parser capabilities and locally imported structured protocol coverage.",(),service.list_protocols),
        ToolDefinition("list_protocol_messages","List messages/PDUs defined for a protocol. Use this when no specific message name is known.",(
            ToolArgument("protocol",ArgumentType.STRING,"Protocol name such as PFCP, NGAP, GTP-U, NR RRC, or 5GS NAS.",True,max_length=100),
            ToolArgument("version",ArgumentType.STRING,"Optional exact imported version.",max_length=40),
            ToolArgument("release",ArgumentType.INTEGER,"Optional Release number.",minimum=1,maximum=99),
            ToolArgument("limit",ArgumentType.INTEGER,"Maximum messages to return.",default=100,minimum=1,maximum=250)),service.list_messages),
        ToolDefinition("find_protocol_message","Find one specific structured protocol message/PDU and its fields/IEs. The message argument is required.",(
            ToolArgument("message",ArgumentType.STRING,"Specific message/PDU name or distinctive fragment.",True,max_length=300),
            ToolArgument("protocol",ArgumentType.STRING,"Optional protocol name.",max_length=100),
            ToolArgument("version",ArgumentType.STRING,"Optional exact imported version.",max_length=40),
            ToolArgument("release",ArgumentType.INTEGER,"Optional Release number.",minimum=1,maximum=99)),service.find_message),
        ToolDefinition("find_protocol_ie","Find a structured IE/field and message usage/definitions.",(
            ToolArgument("ie",ArgumentType.STRING,"IE or field name.",True,max_length=300),
            ToolArgument("protocol",ArgumentType.STRING,"Optional protocol name.",max_length=100),
            ToolArgument("version",ArgumentType.STRING,"Optional exact imported version.",max_length=40),
            ToolArgument("release",ArgumentType.INTEGER,"Optional Release number.",minimum=1,maximum=99),
            ToolArgument("search_descriptions",ArgumentType.BOOLEAN,"Also search descriptions.",default=True)),service.find_ie),
    )
