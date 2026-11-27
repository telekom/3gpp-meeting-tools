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
   - **Interactive HTML (`render_packet_html(..., interactive_links=True)`):** Uses `core.ui.theme.C` tokens, explicit HTML `<table width="100%">` + percentage `<th width="...">` attributes (required because Qt `QTextDocument` ignores CSS `width: 100%`), plain-text `Spec / Clause` right column, and `action:packet_field:` links on field/IE cells that inspect fields cleanly in the embedded Side Panel without blocking the GUI.
   - **Word / Outlook HTML (`interactive_links=False`):** Strips `action:` links and passes output through [`compact_html_for_word()`](../src/modules/nas/core/word_html.py) (`MsoNormal` zero-margin paragraphs + tight `mso-padding-alt`) so pasted Word tables keep compact row heights.
   - **PlantUML `packetdiag` (`render_packet_plantuml_packetdiag()`):** Emits verified `@startpacketdiag ... @endpacketdiag` source with contiguous bit offsets (`0-2:`, `3:`, ...), `rotate = 270` for narrow multi-char flags, and `colheight = 2` for wide/payload blocks — both clipboard-copyable and 1-click loadable into the **PlantUML IDE** tab.
   - **Markdown & Plain Text:** GFM tables and aligned monospace wire maps.

---

## 2. Currently Supported Protocols

### 2.1. 5G User Plane (`TS 29.281` + `TS 38.415`)
- **TS 29.281 (GTPv1-U):** Common Header (Cl. 5.1, Octets 1–12), Extension Header framing (Cl. 5.2.1), `Long PDCP PDU Number` (`0x03`, Cl. 5.2.2.2A), `NR RAN Container` (`0x84`, Cl. 5.2.2.6 placeholder), `PDU Session Container` (`0x85`, Cl. 5.2.2.7), `PDU Set Information Container` (`0x04`, Cl. 5.2.2.8), `G-PDU` (`255`), `End Marker` (`254`), and `Dummy G-PDU` (no T-PDU).
- **TS 38.415 (PDU Session & PDU Set UP):**
  - `5.5.2.1` **DL PDU SESSION INFORMATION (PDU Type 0)** — incl. `QMP`, `SNP`, `MSNP`, `PPP`, `RQI`, `QFI`, `PPI`, Rel-19 `BSSI`/`BSSize`, `TTNBI`/`TTNB`.
  - `5.5.2.2` **UL PDU SESSION INFORMATION (PDU Type 1)** — incl. `QMP`, `DL/UL/N3-N9 Delay`, `SNP`, `New IE Flag` + Rel-16/17/18/19 `New IE Flags Octet` (`D1 UL PDCP Delay`, `UL/DL Congestion`, `UL/DL Available Bitrate`).
  - `6.5.2.1` **DL PDU SET INFORMATION (PDU Type 0)** — incl. `EDB`, `EPDU`, `PSSI`, `QFI`, 10-bit split `PSSN`, `PSI`, `PSN`, `PSSize`.
- **Interface Presets:** `N3 / N9`, `N3 / N9 + PDU Set Chained`, `Xn-U`, `F1-U`, `QoS Monitoring Dummy G-PDU`, and `End Marker`.

### 2.2. N2 Control Plane (`TS 38.412` SCTP + `TS 38.413` NGAP + `TS 24.501` `NAS-PDU`)
Implemented via [`ngap_catalog.py`](../src/modules/nas/core/packet_stack/ngap_catalog.py), [`ngap_composer.py`](../src/modules/nas/core/packet_stack/ngap_composer.py), and [`registry.py`](../src/modules/nas/core/packet_stack/registry.py):
- **Outer IP + SCTP (`RFC 9260` + `TS 38.412` Cl. 7):** Outer IPv4/IPv6 (`Proto/Next = 132 [SCTP]`), 12-octet SCTP Common Header (`Dst Port = 38412`), optional bundled 16-octet `SACK` chunk, and 16-octet `DATA` chunk (`Stream 0` for non-UE vs `Stream ≠ 0` for UE-associated signalling, `PPID = 60`).
- **Bit-Exact APER Envelope + Variable IE Values (`TS 38.413` Cl. 9.4 + `ITU-T X.691`):** Models the deterministic APER PDU header (`CHOICE`, `procedureCode`, `criticality`, `Open-Type Length`, `protocolIEs` count) and each `ProtocolIE-Field` header (`IE ID [16b]`, `Criticality`, `Length`) bit-accurately while marking ASN.1 APER IE payloads cleanly as variable APER-encoded blocks.
- **Interactive Optional IE Pills & Inline `NAS-PDU` Expansion:** Mandatory (`M`) IEs are always present; Optional/Conditional (`O`/`C`) top-level IEs can be toggled via pills. Messages carrying `NAS-PDU` (`IE ID 38`) can expand any sample `TS 24.501` 5GS NAS message (with optional 7-octet `5GMM` Security Protected Header) inline.

### 2.3. Inline IE Definition Side Panel (`packet_ie_side_panel.py`, `packet_ie_resolver.py`)
- **Responsive Splitter Layout:** `PacketStackWidget` pairs the packet table `QTextBrowser` and [`PacketIESidePanel`](../src/modules/nas/ui/components/packet_ie_side_panel.py) inside `body_splitter` (Right-docked by default, with a `⬓` / `◫` Right $\leftrightarrow$ Bottom toggle persisted in `QSettings`).
- **Qt Width Rule:** `render_packet_html()` sets both HTML `width="100%"` on `<table>` and percentage `width` attributes on `<th>` columns (`Octet` `7%`, 8 bit columns `8.5%`, `Size` `9%`, `Spec / Clause` `16%`) because Qt 5 `QTextDocument` ignores CSS `style="width: 100%"`, allowing the packet table to reflow dynamically as the Side Panel opens or resizes.
- **Non-Blocking Click Routing:** Every link inside the packet table (field cells and `Spec / Clause` cells) routes to the Side Panel rather than leaving Packet mode or invoking Word COM. Explicit Word launch (`📝 Word`) uses [`open_document_async`](../src/modules/word_tools/core/document_opener_async.py) off the GUI thread.
- **Resolution & Candidate Rules (`resolve_packet_field()`):** Resolves `PacketField.def_name` / `ie_clause` (mapped from `type_reference` / `ie_definitions` in `ngap_catalog.py`), then probes ordered candidates (`def_name`, `"Criticality"` for criticality flags, stripped `" Value"`/`" Value Length"`/`" Len"`, embedded `IE ID = N (Name)`, and hyphen $\leftrightarrow$ space variants). Pure wire/external fields (e.g. `RFC 9260`, `X.691`) render a compact wire-field summary card. Reusable across future CP stacks (RRC / SDAP / PDCP still to-do).

### 2.4. 5GS NAS & Encapsulated Payloads (`TS 24.501` + UE Policies)
Implemented via [`nas_catalog.py`](../src/modules/nas/core/packet_stack/nas_catalog.py) and [`nas_composer.py`](../src/modules/nas/core/packet_stack/nas_composer.py):
- **5-Layer UE Policy Delivery Chain (`nas_ue_policy_ursp`):** Visually chains optional outer carrier $\to$ `5GMM Security Protected Header` (`7 B`, Cl. `9.1`) $\to$ `DL/UL NAS TRANSPORT` (`Payload container type = 0101`) $\to$ `MANAGE UE POLICY COMMAND` (`TS 24.501` Annex `D.5.1` / `D.6.2` `PLMN` + `UPSC` + `UE policy part type`) $\to$ inner `URSP Rule` (`TS 24.526` Cl. `5.2` Traffic Descriptor & Route Selection Descriptor toggles) or sibling policy specs (`ANDSP`, `V2XP` `24.588`, `ProSeP` `24.555`, `A2XP` `24.578`, `RSLPP` `24.514`).
- **5GSM, SOR & SMS Containers:** Expands `5GSM` (`TS 24.501` Cl. `8.3` e.g. `PDU SESSION ESTABLISHMENT ACCEPT`), `SOR Transparent Container` (Cl. `9.11.3.51`), and `SMS over NAS` (`TS 24.011` `CP-DATA`) inside `DL/UL NAS TRANSPORT`.
- **Selectable Carrier Envelope (`Carrier:`):** Switch between `N1 Standalone (NAS Only)`, `Over N2 (NGAP + SCTP)`, and `Over Uu (MAC → RLC AM → PDCP SRB2 → RRC DL/ULInformationTransfer + MAC-I)`.

---

## 3. To-Do & Extension Roadmap (Next Protocol Stacks)

Because [`models.py`](../src/modules/nas/core/packet_stack/models.py), [`exporters.py`](../src/modules/nas/core/packet_stack/exporters.py), and [`registry.py`](../src/modules/nas/core/packet_stack/registry.py) are protocol-agnostic:

### 3.1. Peer RAN3 SCTP/APER Interfaces (Direct Reuse of `ngap_composer.py`)
- **`XnAP` (`TS 38.423`, Port `38422`, PPID `61`)**, **`F1AP` (`TS 38.473`, Port `38472`, PPID `62`)**, **`E1AP` (`TS 38.463`, Port `38462`, PPID `64`)**, and **`S1AP` (`TS 36.413`, Port `36412`, PPID `18`)** share the exact `RAN3Asn1Parser` schema and SCTP/APER `ProtocolIE-Field` structure — they only need SCTP Port/PPID catalog entries and non-UE message prefixes.

### 3.2. Remaining User Plane Protocols
1. **`TS 38.425` (NR User Plane Protocol on Xn-U / F1-U / E1):** Add a parser for `DL USER DATA` / `DL DATA DELIVERY STATUS` and replace `build_nr_ran_container_fields()` inside `0x84 NR RAN Container`.
2. **LTE RAN Containers (`TS 36.425` X2-U, `TS 36.465` Xw-U)**, **inner T-PDU IP/Ethernet headers**, and **MASQUE / IPsec** outer transports.

### 3.3. Uu Air Interface Stacks (Requires L2 Header Catalogs First)
1. **Control Plane Uu (`RRC → PDCP → RLC → MAC`):** Requires adding `TS 38.323` (PDCP `SN` + `MAC-I`), `TS 38.322` (RLC), and `TS 38.321` (MAC subheaders) before wrapping `TS 38.331` RRC + `TS 24.501` `dedicatedNAS-Message`.
2. **User Plane Uu (`SDAP → PDCP → RLC → MAC`):** Requires `TS 37.324` (`SDAP` 1-octet `D/C` + `RDI` + `RQI` + `QFI` header) + `TS 38.323` (PDCP).
3. **PFCP over UDP (`TS 29.244`):** `IP → UDP (Dst Port 8805) → PFCP Header (Cl. 7.2.2, S SEID flag) → grouped PFCP IEs`.
