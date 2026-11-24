# `modules.word_tools.core.cr_pcr`

Modular 3GPP Change Request (CR) & Pseudo-CR (pCR) generator package decomposed from the former `cr_pcr_generator.py` monolith:

- [`models.py`](models.py): OpenXML namespaces, `ContributionMetadata`, `ChangeBlock`, contiguous heading block grouping, and affected-clause summaries.
- [`cr_builder.py`](cr_builder.py): Formal CR `docProps/custom.xml` + `word/document.xml` DOCPROPERTY/table/anchor builder (`build_cr_base_document`).
- [`pcr_builder.py`](pcr_builder.py): Official MCC-styled pseudo-CR builder (`build_pcr_base_document`) with boxed blue 14pt separators and 3.5cm hanging-indent metadata rows.
- [`engine.py`](engine.py): `CrPcrGeneratorEngine` (isolated fragment extraction, Word COM `InsertFile` splicing, 3GPP style/MIP lock-in) and `CrPcrGeneratorThread`.
