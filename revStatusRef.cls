VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "revStatusRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- Accessible Attributes ---

' Current Review Status
Public RevStatusCol As Range

' AutoReview Status Column
Public AutoReviewDropCol As Range

' Review Status set by a human
Public HumanSetStatusCol As Range

