Attribute VB_Name = "Config"
'==============================================================================
' Config.bas
'------------------------------------------------------------------------------
' Central configuration for the CAD Batch Converter.
'
' Everything a user may want to change lives here: folder paths, the target
' layer name, the extrude depth, tolerances, ProgIDs and behaviour switches.
' No magic numbers are used elsewhere in the project - every tunable value is
' declared once, here, as a strongly typed Public Const.
'
' Target hosts : AutoCAD 2026 (COM) + SolidWorks 2025 (API)
' Language     : VBA only (VBA7 / 64-bit)
' Required refs: AutoCAD 2026 Type Library, SolidWorks 2025 Type Library
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' 1. FOLDER PATHS
'    Trailing back-slashes are REQUIRED - other modules concatenate file names
'    directly onto these strings.
'------------------------------------------------------------------------------
Public Const SOURCE_FOLDER As String = _
    "U:\Engineering\CAD Services\Current Jobs\018 - HII - Base A's - Qty 5 - RFQ #6001358953\Drawings\UNIT 6033\250-HSLA-65\"

Public Const FILTERED_FOLDER As String = _
    "U:\Engineering\CAD Services\Current Jobs\018 - HII - Base A's - Qty 5 - RFQ #6001358953\Drawings\UNIT 6033\Working Model\FilteredDWGs\"

Public Const OUTPUT_FOLDER As String = _
    "U:\Engineering\CAD Services\Current Jobs\018 - HII - Base A's - Qty 5 - RFQ #6001358953\Drawings\UNIT 6033\Working Model\staging\"

'------------------------------------------------------------------------------
' 2. CORE PROCESS PARAMETERS
'------------------------------------------------------------------------------
' Only geometry on this layer survives the AutoCAD filter step.
Public Const TARGET_LAYER As String = "CUT-OUTSIDE STRAIGHT"

' Layer(s) that hold the words / part marking text (TEXT and MTEXT). The
' strings are harvested from these layers BEFORE the filter deletes them, then
' recreated on the part as a native SolidWorks sketch (see TextMarking.bas).
'   >>> SET THIS to the exact layer name(s) that hold the words. <<<
' One layer or several - separate multiple names with commas, e.g.:
'   "PIN STAMP TEXT, PART MARKING, ETCH"
' (spaces around the commas are ignored; layer names themselves may contain
' spaces). Leave it empty ("") to disable text marking entirely.
Public Const TEXT_LAYER As String = "PIN STAMP TEXT"

' DWG drawing units expressed in meters, used to place the harvested text at the
' same location as the imported outline. 0.0254 = inches, 0.001 = millimeters.
' If the words land in the wrong spot, this is the one value to change.
Public Const DWG_UNITS_TO_METERS As Double = 0.0254

' Extrusion depth, expressed in METERS because the SolidWorks API is always SI.
' 0.00635 m = 6.35 mm = 0.25 in.
Public Const EXTRUDE_DEPTH_METERS As Double = 0.00635

' Optionally render the words as a shallow ENGRAVED groove (a thin-slot cut
' along each stroke). DEFAULT False: the words and marking lines are REFERENCE
' ONLY - they stay as an un-consumed sketch and nothing is cut into the steel.
' Set True only if you ever want the text physically engraved on the model.
Public Const TEXT_ENGRAVE As Boolean = False
' Engrave depth into the top face, meters. 0.0002 m = 0.2 mm (~0.008 in).
Public Const TEXT_ENGRAVE_DEPTH_M As Double = 0.0002
' Engraved stroke width as a fraction of the text cap height. 0.12 reads like
' a pin-stamped line.
Public Const TEXT_STROKE_WIDTH_FRAC As Double = 0.12

' File mask used to enumerate the source folder.
Public Const DWG_FILESPEC As String = "*.dwg"

' File mask used by the text-stamp pass to enumerate finished parts.
Public Const SLDPRT_FILESPEC As String = "*.SLDPRT"

' Name of the batch log file (written into OUTPUT_FOLDER).
Public Const LOG_FILE_NAME As String = "BatchLog.txt"

'------------------------------------------------------------------------------
' 3. TOLERANCES
'------------------------------------------------------------------------------
' Gap that the DWG import will attempt to merge closed when snapping coincident
' end points together (helps close tiny gaps that would otherwise open a
' contour). 0.000254 m = 0.254 mm = 0.010 in.
Public Const IMPORT_MERGE_METERS As Double = 0.000254

' Grid used when comparing sketch end points during open-contour detection.
' Points closer than this are treated as the same vertex. 0.00001 m = 0.01 mm.
Public Const POINT_TOLERANCE_METERS As Double = 0.00001

' Lengths below this (in raw DWG units) are treated as zero by the
' native-sketch workaround: zero-length segments are skipped and a polyline
' bulge smaller than this is drawn as a straight line.
Public Const GEOM_EPS_DWG As Double = 0.0000001

'------------------------------------------------------------------------------
' 4. COM PROGIDS
'    Versioned ProgIDs are tried first; the generic ProgID is the fallback so
'    the project keeps working after a version upgrade.
'    AutoCAD 2025/2026 share major version 25  -> "AutoCAD.Application.25"
'    SolidWorks 2025 is major version 33        -> "SldWorks.Application.33"
'------------------------------------------------------------------------------
Public Const ACAD_PROGID_VERSIONED As String = "AutoCAD.Application.25"
Public Const ACAD_PROGID_GENERIC As String = "AutoCAD.Application"
Public Const SW_PROGID_VERSIONED As String = "SldWorks.Application.33"
Public Const SW_PROGID_GENERIC As String = "SldWorks.Application"

'------------------------------------------------------------------------------
' 5. BEHAVIOUR SWITCHES
'------------------------------------------------------------------------------
' Keep both applications hidden for speed (recommended for production runs).
' Flip to True only when debugging a specific file interactively.
Public Const APP_VISIBLE As Boolean = False

' Unlock any locked (non-target) layer so its entities can be deleted.
' The locked layer names are always reported in the log regardless.
Public Const UNLOCK_LOCKED_LAYERS As Boolean = True

' Leave the applications running when the batch finishes (False) or shut them
' down (True). Default is to leave them open so a shared/manual session is not
' killed unexpectedly.
Public Const QUIT_APPS_ON_FINISH As Boolean = False

' Show a single summary message box when the batch finishes. Off by default so
' the run is fully unattended.
Public Const SHOW_SUMMARY_DIALOG As Boolean = False

'------------------------------------------------------------------------------
' 6. SOLIDWORKS API ENUMERATION VALUES
'    Declared locally (with an SW_ prefix so they never clash with the type
'    library's own enum names) so the numeric contract passed to the API is
'    explicit and self-documenting. Values match the SolidWorks 2025 API.
'------------------------------------------------------------------------------
' swEndConditions_e.swEndCondBlind
Public Const SW_END_COND_BLIND As Long = 0
' swStartConditions_e.swStartSketchPlane
Public Const SW_START_SKETCH_PLANE As Long = 0
' swSaveAsVersion_e.swSaveAsCurrentVersion
Public Const SW_SAVEAS_CURRENT As Long = 0
' swSaveAsOptions_e.swSaveAsOptions_Silent
Public Const SW_SAVE_SILENT As Long = 1
' NOTE: the DWG import method is NOT a numeric constant here on purpose.
' SolidWorks_Import uses the real swconst enum swImportDxfDwg_ImportToPartSketch
' (SOLIDWORKS Constant Type Library, swconst.tlb - see README "Setup"), because
' a wrong numeric guess silently falls back to the DWG default, which is
' "create new DRAWING" - exactly the failure this project must never have.
' Argument string passed to LoadFile4 for a DWG/DXF import.
Public Const SW_IMPORT_ARGS As String = "r"
' swDocumentTypes_e.swDocPART - used when re-opening an SLDPRT in the text pass
' and to verify the DWG import actually produced a part.
Public Const SW_DOC_PART As Long = 1
' swDocumentTypes_e.swDocDRAWING - what the import must NOT produce.
Public Const SW_DOC_DRAWING As Long = 3
' swOpenDocOptions_e.swOpenDocOptions_Silent
Public Const SW_OPEN_SILENT As Long = 1
' swRebuildOnActivation_e.swDontRebuildActiveDoc - used when (re)activating the
' imported part before selection, since selection acts on the active document.
Public Const SW_ACTIVATE_NO_REBUILD As Long = 1
' swUserPreferenceStringValue_e.swDefaultTemplatePart - the user's default part
' template, used by the native-sketch workaround to create a fresh empty part.
Public Const SW_PREF_DEFAULT_PART_TEMPLATE As Long = 8
' Feature type name reported by an imported/normal 2D sketch.
Public Const SW_SKETCH_TYPENAME As String = "ProfileFeature"

' Feature type name reported by a reference plane (Front/Top/Right). The first
' one in the tree is the Front plane - locale independent, unlike its name.
Public Const SW_REFPLANE_TYPENAME As String = "RefPlane"

' swBodyType_e.swSolidBody - used when fetching the extruded solid's faces.
Public Const SW_SOLID_BODY As Long = 0

' A face counts as "the top" when its normal is within this of +Z (0,0,1).
Public Const FACE_NORMAL_TOL As Double = 0.001

' Sketch-text formatting passed to InsertSketchText.
' swTextAlign_e.swTextAlignLeft
Public Const SW_TEXT_ALIGN_LEFT As Long = 0
' Width / spacing / height as a percentage of the document font (100 = default).
Public Const SW_TEXT_WIDTH_PCT As Long = 100
Public Const SW_TEXT_SPACING_PCT As Long = 100
Public Const SW_TEXT_HEIGHT_PCT As Long = 100

'------------------------------------------------------------------------------
' 7. STRUCTURED PER-FILE RESULT
'    Passed by reference through the pipeline and then serialised to the log,
'    so every stage records its own success/failure independently.
'------------------------------------------------------------------------------
Public Type TFileResult
    FileName As String          ' source DWG file name (no path)
    AutoCadOK As Boolean        ' filter + audit + purge + save succeeded
    ImportOK As Boolean         ' SolidWorks import produced a part + sketch
    ExtrudeOK As Boolean        ' blind extrusion feature was created
    TextCount As Long           ' number of words harvested from TEXT_LAYER
    TextOK As Boolean           ' text sketch created on the part
    SaveOK As Boolean           ' SLDPRT written to the staging folder
    OpenContour As Boolean      ' open-contour detection flagged this sketch
    LockedLayers As String      ' comma-separated list of locked layers found
    Message As String           ' first error / diagnostic message, if any
    ElapsedSeconds As Double     ' wall-clock processing time for this file
End Type

'------------------------------------------------------------------------------
' 8. OUTLINE SEGMENT (native-sketch workaround)
'    One line / arc / circle harvested from the filtered DWG via AutoCAD COM,
'    carried to SolidWorks and redrawn as native sketch geometry. Coordinates
'    are in raw DWG units; the SolidWorks side scales by DWG_UNITS_TO_METERS.
'------------------------------------------------------------------------------
' TSegment.Kind values:
Public Const SEG_LINE As Long = 0
Public Const SEG_ARC As Long = 1
Public Const SEG_CIRCLE As Long = 2

Public Type TSegment
    Kind As Long                ' SEG_LINE / SEG_ARC / SEG_CIRCLE
    X1 As Double                ' start point (line/arc)
    Y1 As Double
    X2 As Double                ' end point (line/arc)
    Y2 As Double
    CX As Double                ' centre (arc/circle)
    CY As Double
    Radius As Double            ' circle only
    Direction As Long           ' arc only: +1 = counter-clockwise, -1 = CW
End Type

'------------------------------------------------------------------------------
' 9. HARVESTED TEXT MARK
'    One word/label captured from the DWG text layer, carried from the AutoCAD
'    stage to the SolidWorks stage. Coordinates are in raw DWG units.
'------------------------------------------------------------------------------
Public Type TTextMark
    Text As String              ' the string (MTEXT formatting codes stripped)
    X As Double                 ' insertion point X, DWG units
    Y As Double                 ' insertion point Y, DWG units
    Height As Double            ' text height, DWG units
    Rotation As Double          ' rotation, radians
End Type
