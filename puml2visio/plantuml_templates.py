# plantuml_templates.py

# A unified styling block injected into all UML templates to guarantee
# a flat, enterprise-ready look (Arial, white backgrounds, black lines).
COMMON_STYLE = '''skinparam defaultFontName Arial
skinparam shadowing false
skinparam monochrome true
skinparam BackgroundColor white
skinparam DefaultBackgroundColor white
skinparam NoteBackgroundColor white
skinparam ParticipantBackgroundColor white
skinparam ActivityBackgroundColor white
'''

PLANTUML_TYPES = {
    # --- STANDARD UML DIAGRAMS ---
    "Sequence": {
        "url": "https://plantuml.com/en/sequence-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
Alice -> Bob: Authentication Request
note right: This is a white note box
Bob --> Alice: Authentication Response

Alice -> Bob: Another authentication Request
Alice <-- Bob: Another authentication Response
@enduml'''
    },
    "Use Case": {
        "url": "https://plantuml.com/en/use-case-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
left to right direction
actor User

rectangle System {
  User -- (Login)
  User -- (Logout)
}
@enduml'''
    },
    "Class": {
        "url": "https://plantuml.com/en/class-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
class Car {
  - String make
  - String model
  + startEngine()
}

class Engine

Car *-- Engine : contains
@enduml'''
    },
    "Object": {
        "url": "https://plantuml.com/en/object-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
object "user1: User" as user1 {
  name = "Alice"
  id = 123
}

object "user2: User" as user2 {
  name = "Bob"
  id = 456
}

user1 -> user2 : follows
@enduml'''
    },
    "Activity (Beta)": {
        "url": "https://plantuml.com/en/activity-diagram-beta",
        "template": "@startuml\n" + COMMON_STYLE + '''
start
if (Condition?) then (yes)
  :Action 1;
else (no)
  :Action 2;
endif
stop
@enduml'''
    },
    "Component": {
        "url": "https://plantuml.com/en/component-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
package "Some Group" {
  HTTP - [First Component]
  [Another Component]
}

node "Other Groups" {
  FTP - [Second Component]
  [First Component] --> FTP
}
@enduml'''
    },
    "Deployment": {
        "url": "https://plantuml.com/en/deployment-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
node "Server" {
  artifact "Application.war"
}
node "Database" {
  database "MySQL"
}
Server --> Database : TCP/IP
@enduml'''
    },
    "State": {
        "url": "https://plantuml.com/en/state-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
[*] --> State1
State1 --> [*]
State1 : this is a string
State1 : this is another string

State1 -> State2
State2 --> [*]
@enduml'''
    },
    "Timing": {
        "url": "https://plantuml.com/en/timing-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
robust "Web Browser" as WB
concise "Web User" as WU

@0
WU is Idle
WB is Idle

@100
WU is Waiting
WB is Processing

@300
WB is Waiting
@enduml'''
    },

    # --- ARCHITECTURE & DATA STRUCTURES ---
    "Information Engineering (IE)": {
        "url": "https://plantuml.com/en/ie-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
entity "User" as user {
  *user_id : number <<generated>>
  --
  *name : text
  email : text
}

entity "Order" as order {
  *order_id : number <<generated>>
  --
  *user_id : number <<FK>>
  total_amount : number
}

user ||..o{ order
@enduml'''
    },
    "ER Diagram (Chen's Notation)": {
        "url": "https://plantuml.com/en/er-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
entity "Employee" as emp
diamond "Works For" as works
entity "Department" as dept

emp - works : 1..N
works - dept : 1..1
@enduml'''
    },
    "JSON Data": {
        "url": "https://plantuml.com/en/json",
        "template": '''@startjson
{
  "firstName": "John",
  "lastName": "Smith",
  "isAlive": true,
  "age": 28,
  "address": {
    "streetAddress": "21 2nd Street",
    "city": "New York"
  }
}
@endjson'''
    },
    "YAML Data": {
        "url": "https://plantuml.com/en/yaml",
        "template": '''@startyaml
server:
  port: 8080
  host: localhost
database:
  user: admin
  password: secret
  timeout: 5000
@endyaml'''
    },

    # --- ADVANCED / NON-UML DIAGRAMS ---
    "Network (nwdiag)": {
        "url": "https://plantuml.com/en/nwdiag",
        "template": '''@startnwdiag
nwdiag {
  network dmz {
    address = "210.x.x.x/24"
    web01 [address = "210.x.x.1"];
    web02 [address = "210.x.x.2"];
  }
  network internal {
    address = "172.x.x.x/24";
    web01 [address = "172.x.x.1"];
    db01;
  }
}
@endnwdiag'''
    },
    "Rack (rackdiag)": {
        "url": "https://plantuml.com/en/nwdiag",
        "template": '''@startrackdiag
rackdiag {
  16U;
  1: UPS [2U];
  3: DB Server;
  4: Web Server;
  5: Web Server;
  7: Load Balancer;
  8: L3 Switch;
}
@endrackdiag'''
    },
    "Packet (packetdiag)": {
        "url": "https://plantuml.com/en/nwdiag",
        "template": '''@startpacketdiag
packetdiag {
  colwidth = 32
  node_height = 72

  0-15: Source Port
  16-31: Destination Port
  32-63: Sequence Number
  64-95: Acknowledgment Number
  96-99: Data Offset
  100-105: Reserved
  106: URG [rotate = 270]
  107: ACK [rotate = 270]
  108: PSH [rotate = 270]
  109: RST [rotate = 270]
  110: SYN [rotate = 270]
  111: FIN [rotate = 270]
  112-127: Window
  128-143: Checksum
  144-159: Urgent Pointer
  160-191: (Options and Padding)
  192-223: data [colheight = 3]
}
@endpacketdiag'''
    },
    "Wireframe / UI Mockup (Salt)": {
        "url": "https://plantuml.com/en/salt",
        "template": '''@startsalt
{
  Login    | "MyName   "
  Password | "**** "
  [Cancel] | [  OK   ]
}
@endsalt'''
    },
    "Files / Tree Diagram": {
        "url": "https://plantuml.com/en/salt",
        "template": '''@startsalt
{
  T
  + Workspace
  ++ Project
  +++ src
  ++++ main.py
  +++ data
  ++++ config.json
  ++ README.md
}
@endsalt'''
    },
    "EBNF (Syntax Grammar)": {
        "url": "https://plantuml.com/en/ebnf",
        "template": '''@startebnf
title EBNF Diagram
letter = "A" | "B" | "C" ;
digit = "0" | "1" | "2" ;
alphanumeric = letter | digit ;
@endebnf'''
    },
    "Regex (Regular Expression)": {
        "url": "https://plantuml.com/en/regex",
        "template": '''@startregex
title Email Validator
[a-zA-Z0-9_]+@[a-zA-Z0-9_]+\.[a-zA-Z]{2,4}
@endregex'''
    },
    "Archimate": {
        "url": "https://plantuml.com/en/archimate-diagram",
        "template": "@startuml\n" + COMMON_STYLE + '''
archimate #Technology "VPN Server" as vpn <<technology-device>>
archimate #Technology "Mobile App" as mob <<technology-device>>
vpn - mob
@enduml'''
    },
    "SDL (Telecom/Logic)": {
        "url": "https://plantuml.com/en/activity-diagram-beta#sdl",
        "template": "@startuml\n" + COMMON_STYLE + '''
:Ready;
:next(o)|
:Receiving;
split
 :nak(i)<
 :ack(o)>
split again
 :ack(i)<
 :next(o)
end split
:wait;
@enduml'''
    },
    "Ditaa (ASCII Art)": {
        "url": "https://plantuml.com/en/ditaa",
        "template": '''@startditaa
+--------+   +-------+
| cAAA   +---+Version|
|  Data  |   |   V3  |
|  Base  |   |cRED{d}|
+---+----+   +-------+
@endditaa'''
    },
    "Mathematics (AsciiMath)": {
        "url": "https://plantuml.com/en/ascii-math",
        "template": '''@startmath
f(t)=(a_0)/2 + sum_(n=1)^ooa_ncos((npit)/L)+sum_(n=1)^oo b_n\ sin((npit)/L)
@endmath'''
    },
    "Chart (Pie)": {
        "url": "https://plantuml.com/en/pie-chart",
        "template": "@startuml\n" + COMMON_STYLE + '''
pie title Programming Languages
"Java" : 30
"Python" : 40
"C++" : 20
"JavaScript" : 10
@enduml'''
    },

    # --- PROJECT MANAGEMENT ---
    "MindMap": {
        "url": "https://plantuml.com/en/mindmap-diagram",
        "template": "@startmindmap\n" + COMMON_STYLE + '''
* Root concept
** Sub-concept 1
*** Idea A
*** Idea B
** Sub-concept 2
@endmindmap'''
    },
    "Work Breakdown Structure (WBS)": {
        "url": "https://plantuml.com/en/wbs-diagram",
        "template": "@startwbs\n" + COMMON_STYLE + '''
* Project
** Phase 1
*** Task 1.1
*** Task 1.2
** Phase 2
@endwbs'''
    },
    "Gantt Chart": {
        "url": "https://plantuml.com/en/gantt-diagram",
        "template": "@startgantt\n" + COMMON_STYLE + '''
[Prototype design] lasts 15 days
[Test prototype] lasts 10 days
[Test prototype] starts at [Prototype design]'s end
@endgantt'''
    },
    "Chronology / Timeline": {
        "url": "https://plantuml.com/en/gantt-diagram",
        "template": "@startgantt\n" + COMMON_STYLE + '''
Project starts the 2026-06-01
[Phase 1] lasts 10 days
[Phase 2] lasts 5 days
[Phase 2] starts at [Phase 1]'s end
@endgantt'''
    }
}