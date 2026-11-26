# Cross-Specification Protocol Packet Stack Methodology (`packet_stack/`)

Architectural blueprint and extension guide for assembling end-to-end on-the-wire 3GPP packet diagrams across multiple specifications (implemented in [`src/modules/nas/core/packet_stack/`](../src/modules/nas/core/packet_stack/) and surfaced in [`PacketStackWidget`](../src/modules/nas/ui/components/packet_stack_widget.py)).

---

## 1. Core Design Philosophy

3GPP specifications define each encapsulation layer in isolation (e.g., transport in TS 38.414 / RFCs, GTPv1-U headers & extension header framing in **TS 29.281**, and User Plane frames in **TS 38.415** or **TS 38.425**). Rather than hardcoding static bitmaps, `packet_stack/` models a packet as a **composed, ordered list of `PacketLayer` objects** populated from **pure-Python catalog builders + optional SQLite DB release inspection**, solved deterministically at render time:

1. **Decoupled Data Model ([`models.py`](../src/modules/nas/core/packet_stack/models.py)):**
   - `PacketField` stores bit/octet widths (`bits`, `is_variable`), presence (`M`/`C`/`O`), spec/clause origin, `flag_key` binding, and conditional trigger text.
   - `PacketLayer` groups fields by encapsulation role (`"transport"`, `"gtpu_header"`, `"gtpu_ext_wrapper"`, `"up_frame"`, `"ran_container"`, `"payload"`).
   - `FlagOption` + `PacketScenarioPreset` + `PacketConfig` drive interactive flag toggles, scenario switching, outer transport expansion, and release-specific availability (`is_release_present`).
   - `ComposedPacket` is the single immutable-snapshot output consumed by all renderers and exporters.

2. **Target Resolution & Release-Aware DB Enrichment ([`up_catalog.py`](../src/modules/nas/core/packet_stack/up_catalog.py)):**
   - `resolve_up_packet_target(spec_number, clause, ie_name, message_name)` maps any selected message or IE in `TS 29.281` or `TS 38.415` to a `{target_frame, preset_id, auto_enable_flags, highlight_field}` tuple (returning `None` for non-UP specs so the Inspector stays clean).
   - Selecting a conditional IE (e.g., `DL Sending Time Stamp` Cl. `5.5.3.9` or `BSSize` Cl. `5.5.3.30`) automatically enables its gating flag (`qmp=True`, `bssi=True`) and highlights the field in the rendered stack.
   - `load_up_release_capabilities(db, version_38415, version_29281)` queries `3gpp_protocol_data.db` (`nas_messages` + `message_ies` + `ie_definitions`) to check which optional fields exist in the selected TS 38.415 release (e.g. Rel-15 vs Rel-19) and pulls concise field descriptions. If a spec is not yet imported, built-in Rel-19 defaults provide seamless fallback.

3. **Deterministic GTPv1-U Framing & 4-Octet Padding Solver ([`composer.py`](../src/modules/nas/core/packet_stack/composer.py)):**
   - `compose_gtpu_extension_header()` wraps any inner frame in **TS 29.281 Clause 5.2.1** framing:
     - Prepends **`Extension Header Length`** (1 octet, in 4-octet units `n`).
     - Appends **`Next Extension Header Type`** (1 octet: chained to next extension header codepoint, or `0x00` to terminate).
     - Computes exact padding for fixed-length inner frames:
       - `raw_octets = 1 + inner_fixed_octets + 1`
       - `padding_octets = (4 - (raw_octets % 4)) % 4`
       - `total_octets = raw_octets + padding_octets`, `n = total_octets // 4`
     - For variable-length inner frames (e.g., `NR RAN Container`), inserts a symbolic `0–3 octets (4-octet align)` padding row.
   - Extension headers chain automatically (`0x03 Long PDCP SN` $\to$ `0x84 NR RAN Container` $\to$ `0x85 PDU Session Container` $\to$ `0x04 PDU Set Information Container` $\to$ `0x00`), updating Octet 12 (`Next Extension Header Type`) in the GTP-U header and `E`/`S` flags automatically.

4. **Multi-Format Exporters ([`exporters.py`](../src/modules/nas/core/packet_stack/exporters.py)):**
   - `_pack_layer_rows()` packs contiguous sub-octet bitfields into 8-bit visual rows and computes continuous 1-based wire octet offsets across all layers.
   - **Interactive HTML (`render_packet_html(..., interactive_links=True)`):** Uses `core.ui.theme.C` tokens and `action:packet_field:` / `action:open_spec:` links for click-to-inspect and Word opening.
   - **Word / Outlook HTML (`interactive_links=False`):** Strips `action:` links for clean pasting into TDocs, slides, and emails.
   - **PlantUML `packetdiag` (`render_packet_plantuml_packetdiag()`):** Emits verified `@startpacketdiag ... @endpacketdiag` source with contiguous bit offsets (`0-2:`, `3:`, ...), `rotate = 270` for narrow multi-char flags, and `colheight = 2` for wide/payload blocks — both clipboard-copyable and 1-click loadable into the **PlantUML IDE** tab.
   - **Markdown & Plain Text:** GFM tables and aligned monospace wire maps.

---

## 2. Currently Supported Protocols (v1 — 5G User Plane)

- **TS 29.281 (GTPv1-U):** Common Header (Cl. 5.1, Octets 1–12), Extension Header framing (Cl. 5.2.1), `Long PDCP PDU Number` (`0x03`, Cl. 5.2.2.2A), `NR RAN Container` (`0x84`, Cl. 5.2.2.6 placeholder), `PDU Session Container` (`0x85`, Cl. 5.2.2.7), `PDU Set Information Container` (`0x04`, Cl. 5.2.2.8), `G-PDU` (`255`), `End Marker` (`254`), and `Dummy G-PDU` (no T-PDU).
- **TS 38.415 (PDU Session & PDU Set UP):**
  - `5.5.2.1` **DL PDU SESSION INFORMATION (PDU Type 0)** — incl. `QMP`, `SNP`, `MSNP`, `PPP`, `RQI`, `QFI`, `PPI`, Rel-19 `BSSI`/`BSSize`, `TTNBI`/`TTNB`.
  - `5.5.2.2` **UL PDU SESSION INFORMATION (PDU Type 1)** — incl. `QMP`, `DL/UL/N3-N9 Delay`, `SNP`, `New IE Flag` + Rel-16/17/18/19 `New IE Flags Octet` (`D1 UL PDCP Delay`, `UL/DL Congestion`, `UL/DL Available Bitrate`).
  - `6.5.2.1` **DL PDU SET INFORMATION (PDU Type 0)** — incl. `EDB`, `EPDU`, `PSSI`, `QFI`, 10-bit split `PSSN`, `PSI`, `PSN`, `PSSize`.
- **Interface Presets:** `N3 / N9`, `N3 / N9 + PDU Set Chained`, `Xn-U`, `F1-U`, `QoS Monitoring Dummy G-PDU`, and `End Marker`.

---

## 3. To-Do & Extension Roadmap (How to Add Other Protocol Stacks)

Because `models.py` and `exporters.py` are completely protocol-agnostic, adding new stacks requires **zero changes** to the packing solver, HTML/Word/Markdown/PlantUML exporters, or `PacketStackWidget` shell:

### 3.1. Remaining User Plane Protocols (To-Do)
1. **TS 38.425 (NR User Plane Protocol on Xn-U / F1-U / E1):**
   - Add a dedicated parser for `TS 38.425` (`DL USER DATA` PDU Type 0, `DL DATA DELIVERY STATUS` PDU Type 1, `ASSISTANCE INFORMATION DATA` PDU Type 2) and plug its field builders into `up_catalog.py` to replace the current `build_nr_ran_container_fields()` placeholder inside `0x84 NR RAN Container`.
2. **LTE RAN User Plane Containers (`TS 36.425` X2-U, `TS 36.465` Xw-U):**
   - Map `0x81 RAN Container` and `0x83 Xw RAN Container` similarly to `0x84`.
3. **Inner T-PDU Header Expansion & Tunnel Variants:**
   - Optional expandable inner `IPv4`/`IPv6`/`Ethernet (802.3)`/`UDP`/`TCP` T-PDU header option in `build_transport_layers()` / payload builder, plus **MASQUE / HTTP/3 / QUIC** or **IPsec ESP** outer transport modes.
4. **Legacy / Additional GTP-U Extension Headers:**
   - `0xC0 PDCP PDU Number` (15-bit/12-bit), `0x20 Service Class Indicator`, and `UDP Port (0x40)` on `Error Indication`.

### 3.2. Control Plane & UE Policy Encapsulation Stacks (To-Do)
To extend this pattern beyond User Plane (e.g., into a generalized `stack_registry.py`):
1. **5GS NAS + UE Policy Stack (`TS 24.526 / 24.588 / 24.555` $\to$ `TS 24.501` $\to$ `TS 38.413` $\to$ `TS 38.412`):**
   - Outer `SCTP` (PPID 60 for NGAP) $\to$ `NGAP` (`DownlinkNASTransport`) $\to$ `5GS NAS` (`DL NAS TRANSPORT`, security-protected `5GMM` header + `Payload container` `9.11.3.39` / `MANAGE UE POLICY COMMAND` `D.5.1`) $\to$ `UE Policy Section Management List` (`D.6.2`) $\to$ `URSP Rule` (`TS 24.526 Cl. 5.2`).
2. **PFCP over UDP (`TS 29.244`):**
   - `IP` $\to$ `UDP (Dst Port 8805)` $\to$ `PFCP Header` (Cl. 7.2.2, `S` SEID flag toggling Octets 5–12) $\to$ selected PFCP Message grouped IEs.
3. **NR RRC Air Interface Stack (`TS 38.331` $\to$ `TS 38.323` PDCP $\to$ `TS 38.322` RLC):**
   - `RLC` header + `PDCP` header (`SN`, `MAC-I`) + `RRC` ASN.1 PDU summary.

**Recipe to add any new stack:**
1. Create `<domain>_catalog.py` returning `PacketLayer` / `FlagOption` / `PacketScenarioPreset` + `resolve_<domain>_packet_target(...)`.
2. Add `<domain>_composer.py` producing a `ComposedPacket` (reusing `PacketLayer` and `PacketField`).
3. Register the resolver in `PacketStackWidget.set_context()` (or a thin dispatcher in `packet_stack/registry.py`) — all UI toggles, octet packing, Word/Markdown/Plain Text/`@startpacketdiag` exports, and PlantUML IDE handoff work automatically.
