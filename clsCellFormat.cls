VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCellFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--- Basic attributes
Public InteriorColor As Long
Public FontColor As Long
Public Bold As Boolean
Public Italic As Boolean
Public Underline As Boolean
Public FontName As String
Public FontSize As Double
Public NumberFormat As String

'--- Alignment
Public HorizontalAlignment As Long
Public VerticalAlignment As Long
Public WrapText As Boolean

'--- Borders
Public BorderTopColor As Long
Public BorderBottomColor As Long
Public BorderLeftColor As Long
Public BorderRightColor As Long
Public BorderTopLineStyle As Long
Public BorderBottomLineStyle As Long
Public BorderLeftLineStyle As Long
Public BorderRightLineStyle As Long

'--- Priority (used for conflict resolution or ranking)
Public Priority As Long
