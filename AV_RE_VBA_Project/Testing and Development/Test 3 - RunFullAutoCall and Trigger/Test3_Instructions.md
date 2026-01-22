# TEST 3: RunFullValidationMaster - Instructions

**Version:** Phase 2 - Test 3  
**Date:** 2026-01-19  
**Status:** Ready to Execute

---

## Setup Confirmed

✅ **Test data:** Ready  
✅ **Validation rules:** Populated (Phase 2B rules)  
✅ **Test scenarios:** Set up  
✅ **GlobalDebugOn:** Set to "ON"

---

## Module Requirements

### Required Modules (Test 2 Versions):
1. ✅ **AV_Core_v2.1_COMPLETE.bas** - From Test 2
2. ✅ **AV_UI_v2.1_Test2.bas** - From Test 2
3. **AV_Constants.bas** - Phase 2 version (if not already imported)
4. **AV_DataAccess.bas** - Phase 2 version (if not already imported)

### Phase 2 Modules (from project folder):
5. **av_engine_fixed.bas** - Use as AV_Engine
6. **av_format_fixed.bas** - Use as AV_Format
7. **av_validators_RoutingLayer.bas** - Use as AV_Validators
8. **av_validation_BusinessRules.bas** - Use as AV_ValidationRules

### Test Module:
9. **Test3_RunFullValidation.bas** - New (just created)

---

## Test Procedure

### Step 1: Pre-Flight Check

Run this first to verify everything is ready:

```vba
Test3_PreFlightCheck
```

**Expected Output:**
```
==========================================
TEST 3 - PRE-FLIGHT CHECK
==========================================

Check 1: GlobalDebugOn setting...
  ✅ PASS: GlobalDebugOn = ON

Check 2: AutoValidationCommentPrefixMappingTable...
  ✅ PASS: Table found with 9 validation functions

Check 3: AutoFormatOnFullValidation table...
  ✅ PASS: Table found with 3 format types

Check 4: Validation rule tables...
  ✅ GIWValidationTable (X rules)
  ✅ ElectricityPairValidation (X rules)
  ✅ PlumbingPairValidation (X rules)
  ✅ HeatSourcePairValidation (X rules)
  ✅ HeatSourceANYRefTable (X rules)

Check 5: Target data sheet...
  ✅ PASS: Target sheet 'ReviewTable' exists
     Data range: Row X to Y
     Total rows: Z

==========================================
PRE-FLIGHT CHECK: ✅ ALL SYSTEMS GO

Ready to run Test3_RunValidation
==========================================
```

**If any checks FAIL:** Fix the issue before proceeding.

---

### Step 2: Quick Validation (Recommended First)

Test with small sample (10 rows only):

```vba
Test3_QuickValidation
```

**What this does:**
- Temporarily sets row count to 10
- Runs full validation on 10 rows
- Restores original row count
- Good for quick verification

**Expected:**
- ValidationTrackerForm appears
- Progress messages appear
- Debug messages in Immediate Window (GlobalDebugOn = ON)
- Validation completes
- Check target sheet for results

---

### Step 3: Full Validation

Once quick test passes, run full validation:

```vba
Test3_RunValidation
```

**What this does:**
- Runs validation on all configured rows
- Shows elapsed time
- Reports completion status

**Monitor:**
1. **ValidationTrackerForm:** Progress updates
2. **Immediate Window:** Debug messages (if GlobalDebugOn = ON)
3. **Target Sheet:** Validation results applied

---

## What to Watch For

### During Validation

**ValidationTrackerForm should show:**
- "Initializing Full Validation Master"
- "Target sheet: [SheetName]"
- "Row range: X to Y"
- Progress updates every 10 rows
- "ADVANCED AUTO VALIDATIONS COMPLETE"
- "Standard menu accessible field validation Completed"

**Immediate Window (if GlobalDebugOn = ON):**
```
[DEBUG] AV_Core :: Table found: AutoValidationCommentPrefixMappingTable (9 rows)
[DEBUG] AV_Core :: Row 1 Processing: Validate_Column_GIWQuantity
[DEBUG] AV_Core :: Row 2 Processing: Validate_Column_GIWIncluded
...
[DEBUG] AV_Core :: Success: 9 | Skipped: 0
```

### After Validation

**Check target sheet:**
1. **Error cells:** Red background (or configured error format)
2. **Autocorrect cells:** Yellow background (or configured autocorrect format)
3. **Valid cells:** Default format
4. **Drop columns:** Validation messages with [[SYS_TAG ...]]
5. **Key column:** Formatted with highest priority error

---

## Expected Results

### ✅ Successful Test

**Immediate Window:**
```
==========================================
TEST 3 - RUNNING FULL VALIDATION
==========================================

Starting validation at hh:mm:ss
Watch ValidationTrackerForm for progress...

---[VALIDATION OUTPUT BELOW]---

[DEBUG messages from validation process]

---[VALIDATION OUTPUT ABOVE]---

Validation completed successfully
Elapsed time: X.XX seconds

==========================================
TEST 3 COMPLETE

NEXT STEPS:
1. Check ValidationTrackerForm for completion status
2. Review target sheet for validation results
3. Verify error/autocorrect formatting applied
4. Check drop columns for validation messages
==========================================
```

**ValidationTrackerForm:**
- All 3 checkboxes checked ✓
- Log shows all steps completed
- No error messages

**Target Sheet:**
- Errors flagged and formatted
- Auto-corrections applied
- Messages in drop columns
- Key column formatted

---

### ❌ Failed Test

**If error occurs:**
```
==========================================
❌ ERROR DURING VALIDATION
==========================================
Error #XXX: [Description]
Source: [Module]

Check ValidationTrackerForm for additional details
==========================================
```

**Troubleshooting:**
1. Note the error number and description
2. Check which module failed (Source)
3. Review ValidationTrackerForm log for context
4. Share error details for diagnosis

---

## Common Issues

### Issue: "Compile Error"
**Cause:** Module not imported or version mismatch  
**Solution:** Run `Test3_CheckModuleVersions` to verify

### Issue: "Table not found"
**Cause:** Missing configuration table  
**Solution:** Run `Test3_PreFlightCheck` to identify missing tables

### Issue: "Object required" error
**Cause:** AV_Core not properly initialized  
**Solution:** Verify AV_Core v2.1 COMPLETE is imported

### Issue: No debug messages appear
**Cause:** GlobalDebugOn not set correctly  
**Solution:** 
1. Check GlobalDebugOptions table shows "ON"
2. Run `AV_Core.InitDebugFlags True`
3. Check `AV_Core.GlobalDebugOn` = True

### Issue: Validation runs but no results
**Cause:** AutoValidate flags may be False  
**Solution:** Check AutoValidationCommentPrefixMappingTable - AutoValidate column should be "TRUE"

---

## Performance Expectations

**Quick Validation (10 rows):**
- Expected time: < 5 seconds
- Debug logging adds ~1-2 seconds

**Full Validation (depends on row count):**
- 100 rows: ~15-20 seconds
- 500 rows: ~60-90 seconds
- 1000 rows: ~2-3 minutes

**With GlobalDebugOn = OFF:**
- ~20-30% faster (no debug message overhead)

---

## After Test 3

### If Test Passes:
✅ Core validation workflow confirmed working  
✅ Error detection functional  
✅ Auto-correction functional  
✅ Format application working  
✅ Progress tracking working  

**Next:** Test 4 - Trigger-based validation (single cell changes)

### If Test Fails:
1. Share full error output
2. Share ValidationTrackerForm log
3. Share Immediate Window output
4. Share any specific error examples from data

---

## Quick Reference

```vba
' Check setup
Test3_PreFlightCheck

' Quick test (10 rows)
Test3_QuickValidation

' Full test (all rows)
Test3_RunValidation

' Check module versions
Test3_CheckModuleVersions

' Check debug status
Test2_ShowStatus
```

---

**Ready when you are!**
