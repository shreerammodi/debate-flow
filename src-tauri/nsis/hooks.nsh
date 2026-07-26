; Windows has no MIME type hierarchy, so a .ebb file cannot declare itself a
; kind of JSON the way it does through UTI conformance on macOS and
; sub-class-of on Linux. These two registry values are the closest equivalent:
; PerceivedType makes Explorer and the search indexer treat a flow as text, and
; Content Type puts JSON editors in its "Open with" list. ebb stays the default
; handler either way - that association is written by the installer itself from
; bundle.fileAssociations.

!macro NSIS_HOOK_POSTINSTALL
  WriteRegStr SHCTX "Software\Classes\.ebb" "PerceivedType" "text"
  WriteRegStr SHCTX "Software\Classes\.ebb" "Content Type" "application/json"
!macroend

!macro NSIS_HOOK_PREUNINSTALL
  DeleteRegValue SHCTX "Software\Classes\.ebb" "PerceivedType"
  DeleteRegValue SHCTX "Software\Classes\.ebb" "Content Type"
!macroend
