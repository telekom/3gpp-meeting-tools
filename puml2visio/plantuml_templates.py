# plantuml_templates.py

PLANTUML_TYPES = {
    # --- STANDARD UML DIAGRAMS ---
    "Sequence": {
        "url": "https://plantuml.com/en/sequence-diagram",
        "template": "@startuml\nautonumber\n\nAlice -> Bob: Authentication Request\nBob --> Alice: Authentication Response\n\nAlice -> Bob: Another authentication Request\nAlice <-- Bob: Another authentication Response\n@enduml"
    },
    "Use Case": {
        "url": "https://plantuml.com/en/use-case-diagram",
        "template": "@startuml\nleft to right direction\nactor User\nrectangle System {\n  User -- (Login)\n  User -- (Logout)\n}\n@enduml"
    },
    "Class": {
        "url": "https://plantuml.com/en/class-diagram",
        "template": "@startuml\nclass Car {\n  - String make\n  - String model\n  + startEngine()\n}\n\nclass Engine\n\nCar *-- Engine : contains\n@enduml"
    },
    "Object": {
        "url": "https://plantuml.com/en/object-diagram",
        "template": "@startuml\nobject \"user1: User\" as user1 {\n  name = \"Alice\"\n  id = 123\n}\n\nobject \"user2: User\" as user2 {\n  name = \"Bob\"\n  id = 456\n}\n\nuser1 -> user2 : follows\n@enduml"
    },
    "Activity (Beta)": {
        "url": "https://plantuml.com/en/activity-diagram-beta",
        "template": "@startuml\nstart\nif (Condition?) then (yes)\n  :Action 1;\nelse (no)\n  :Action 2;\nendif\nstop\n@enduml"
    },
    "Component": {
        "url": "https://plantuml.com/en/component-diagram",
        "template": "@startuml\npackage \"Some Group\" {\n  HTTP - [First Component]\n  [Another Component]\n}\n\nnode \"Other Groups\" {\n  FTP - [Second Component]\n  [First Component] --> FTP\n}\n@enduml"
    },
    "Deployment": {
        "url": "https://plantuml.com/en/deployment-diagram",
        "template": "@startuml\nnode \"Server\" {\n  artifact \"Application.war\"\n}\nnode \"Database\" {\n  database \"MySQL\"\n}\nServer --> Database : TCP/IP\n@enduml"
    },
    "State": {
        "url": "https://plantuml.com/en/state-diagram",
        "template": "@startuml\n[*] --> State1\nState1 --> [*]\nState1 : this is a string\nState1 : this is another string\n\nState1 -> State2\nState2 --> [*]\n@enduml"
    },
    "Timing": {
        "url": "https://plantuml.com/en/timing-diagram",
        "template": "@startuml\nrobust \"Web Browser\" as WB\nconcise \"Web User\" as WU\n\n@0\nWU is Idle\nWB is Idle\n\n@100\nWU is Waiting\nWB is Processing\n\n@300\nWB is Waiting\n@enduml"
    },

    # --- ARCHITECTURE & DATA STRUCTURES ---
    "Information Engineering (IE)": {
        "url": "https://plantuml.com/en/ie-diagram",
        "template": "@startuml\nentity \"User\" as user {\n  *user_id : number <<generated>>\n  --\n  *name : text\n  email : text\n}\n\nentity \"Order\" as order {\n  *order_id : number <<generated>>\n  --\n  *user_id : number <<FK>>\n  total_amount : number\n}\n\nuser ||..o{ order\n@enduml"
    },
    "ER Diagram (Chen's Notation)": {
        "url": "https://plantuml.com/en/er-diagram",
        "template": "@startuml\nentity \"Employee\" as emp\ndiamond \"Works For\" as works\nentity \"Department\" as dept\n\nemp - works : 1..N\nworks - dept : 1..1\n@enduml"
    },
    "JSON Data": {
        "url": "https://plantuml.com/en/json",
        "template": "@startjson\n{\n  \"firstName\": \"John\",\n  \"lastName\": \"Smith\",\n  \"isAlive\": true,\n  \"age\": 28,\n  \"address\": {\n    \"streetAddress\": \"21 2nd Street\",\n    \"city\": \"New York\"\n  }\n}\n@endjson"
    },
    "YAML Data": {
        "url": "https://plantuml.com/en/yaml",
        "template": "@startyaml\nserver:\n  port: 8080\n  host: localhost\ndatabase:\n  user: admin\n  password: secret\n  timeout: 5000\n@endyaml"
    },

    # --- ADVANCED / NON-UML DIAGRAMS ---
    "Network (nwdiag)": {
        "url": "https://plantuml.com/en/nwdiag",
        "template": "@startnwdiag\nnwdiag {\n  network dmz {\n    address = \"210.x.x.x/24\"\n    web01 [address = \"210.x.x.1\"];\n    web02 [address = \"210.x.x.2\"];\n  }\n  network internal {\n    address = \"172.x.x.x/24\";\n    web01 [address = \"172.x.x.1\"];\n    db01;\n  }\n}\n@endnwdiag"
    },
    "Rack (rackdiag)": {
        "url": "https://plantuml.com/en/nwdiag",
        "template": "@startrackdiag\nrackdiag {\n  16U;\n  1: UPS [2U];\n  3: DB Server;\n  4: Web Server;\n  5: Web Server;\n  7: Load Balancer;\n  8: L3 Switch;\n}\n@endrackdiag"
    },
    "Packet (packetdiag)": {
        "url": "https://plantuml.com/en/nwdiag",
        "template": "@startpacketdiag\npacketdiag {\n  colwidth = 32\n  node_height = 72\n\n  0-15: Source Port\n  16-31: Destination Port\n  32-63: Sequence Number\n  64-95: Acknowledgment Number\n  96-99: Data Offset\n  100-105: Reserved\n  106: URG [rotate = 270]\n  107: ACK [rotate = 270]\n  108: PSH [rotate = 270]\n  109: RST [rotate = 270]\n  110: SYN [rotate = 270]\n  111: FIN [rotate = 270]\n  112-127: Window\n  128-143: Checksum\n  144-159: Urgent Pointer\n  160-191: (Options and Padding)\n  192-223: data [colheight = 3]\n}\n@endpacketdiag"
    },
    "Wireframe / UI Mockup (Salt)": {
        "url": "https://plantuml.com/en/salt",
        "template": "@startsalt\n{\n  Login    | \"MyName   \"\n  Password | \"**** \"\n  [Cancel] | [  OK   ]\n}\n@endsalt"
    },
    "Files / Tree Diagram": {
        "url": "https://plantuml.com/en/salt",
        "template": "@startsalt\n{\n  T\n  + Workspace\n  ++ Project\n  +++ src\n  ++++ main.py\n  +++ data\n  ++++ config.json\n  ++ README.md\n}\n@endsalt"
    },
    "EBNF (Syntax Grammar)": {
        "url": "https://plantuml.com/en/ebnf",
        "template": "@startebnf\ntitle EBNF Diagram\nletter = \"A\" | \"B\" | \"C\" ;\ndigit = \"0\" | \"1\" | \"2\" ;\nalphanumeric = letter | digit ;\n@endebnf"
    },
    "Regex (Regular Expression)": {
        "url": "https://plantuml.com/en/regex",
        "template": "@startregex\ntitle Email Validator\n[a-zA-Z0-9_]+@[a-zA-Z0-9_]+\\.[a-zA-Z]{2,4}\n@endregex"
    },
    "Archimate": {
        "url": "https://plantuml.com/en/archimate-diagram",
        "template": "@startuml\narchimate #Technology \"VPN Server\" as vpn <<technology-device>>\narchimate #Technology \"Mobile App\" as mob <<technology-device>>\nvpn - mob\n@enduml"
    },
    "SDL (Telecom/Logic)": {
        "url": "https://plantuml.com/en/activity-diagram-beta#sdl",
        "template": "@startuml\n:Ready;\n:next(o)|\n:Receiving;\nsplit\n :nak(i)<\n :ack(o)>\nsplit again\n :ack(i)<\n :next(o)\nend split\n:wait;\n@enduml"
    },
    "Ditaa (ASCII Art)": {
        "url": "https://plantuml.com/en/ditaa",
        "template": "@startditaa\n+--------+   +-------+\n| cAAA   +---+Version|\n|  Data  |   |   V3  |\n|  Base  |   |cRED{d}|\n+---+----+   +-------+\n@endditaa"
    },
    "Mathematics (AsciiMath)": {
        "url": "https://plantuml.com/en/ascii-math",
        "template": "@startmath\nf(t)=(a_0)/2 + sum_(n=1)^ooa_ncos((npit)/L)+sum_(n=1)^oo b_n\\ sin((npit)/L)\n@endmath"
    },
    "Chart (Pie)": {
        "url": "https://plantuml.com/en/pie-chart",
        "template": "@startuml\npie title Programming Languages\n\"Java\" : 30\n\"Python\" : 40\n\"C++\" : 20\n\"JavaScript\" : 10\n@enduml"
    },

    # --- PROJECT MANAGEMENT ---
    "MindMap": {
        "url": "https://plantuml.com/en/mindmap-diagram",
        "template": "@startmindmap\n* Root concept\n** Sub-concept 1\n*** Idea A\n*** Idea B\n** Sub-concept 2\n@endmindmap"
    },
    "Work Breakdown Structure (WBS)": {
        "url": "https://plantuml.com/en/wbs-diagram",
        "template": "@startwbs\n* Project\n** Phase 1\n*** Task 1.1\n*** Task 1.2\n** Phase 2\n@endwbs"
    },
    "Gantt Chart": {
        "url": "https://plantuml.com/en/gantt-diagram",
        "template": "@startgantt\n[Prototype design] lasts 15 days\n[Test prototype] lasts 10 days\n[Test prototype] starts at [Prototype design]'s end\n@endgantt"
    },
    "Chronology / Timeline": {
        "url": "https://plantuml.com/en/gantt-diagram",
        "template": "@startgantt\nProject starts the 2026-06-01\n[Phase 1] lasts 10 days\n[Phase 2] lasts 5 days\n[Phase 2] starts at [Phase 1]'s end\n@endgantt"
    }
}