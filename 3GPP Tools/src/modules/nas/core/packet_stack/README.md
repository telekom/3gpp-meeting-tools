# Cross-Specification Packet Stack Package (`packet_stack/`)

Pure-Python engine assembling end-to-end multi-spec wire packet diagrams (starting with 5G User Plane `TS 29.281` + `TS 38.415`), computing deterministic 4-octet extension header padding, and exporting to HTML, Word, Markdown, Plain Text, and PlantUML `packetdiag`.

## Module Map

| File | Responsibility |
|---|---|
| [`models.py`](models.py) | `PacketField`, `PacketLayer`, `FlagOption`, `PacketScenarioPreset`, `PacketConfig`, `ComposedPacket` |
| [`up_catalog.py`](up_catalog.py) | 5G UP scenario presets (`N3/N9`, `Xn-U`, `F1-U`, `Dummy G-PDU`, `End Marker`), DB capability inspection, TS 38.415 frame builders |
| [`composer.py`](composer.py) | Transport & GTPv1-U builders, 4-octet `Padding` / `n` solver (`TS 29.281 Cl. 5.2.1`), extension header chaining, `compose_up_packet()` |
| [`ngap_catalog.py`](ngap_catalog.py) | N2 NGAP scenario presets, `38.413` message/IE & `24.501` NAS-PDU DB loaders, Stream 0 / choice inference |
| [`ngap_composer.py`](ngap_composer.py) | SCTP (`RFC 9260`, `TS 38.412` Port 38412/PPID 60, bundled `SACK`) + NGAP APER envelope + inline `NAS-PDU` (`24.501`) composer |
| [`registry.py`](registry.py) | Multi-stack `resolve_packet_target()`, `compose_packet()`, and preset family dispatchers |
| [`exporters.py`](exporters.py) | Theme-tokenized HTML, Word clipboard, Markdown, Plain Text, and PlantUML `@startpacketdiag` generator |
| [`__init__.py`](__init__.py) | Public package exports |

For full architectural methodology and extension roadmap, see [`docs/PROTOCOL_PACKET_STACK.md`](../../../../docs/PROTOCOL_PACKET_STACK.md).
