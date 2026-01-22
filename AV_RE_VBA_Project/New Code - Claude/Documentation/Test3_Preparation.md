# Test 3: RunFullValidationMaster - Preparation

**Version:** Phase 2 - Test 3  
**Date:** 2026-01-19  
**Status:** Ready to Begin

---

## Test 1 & 2 Summary

### ✅ Test 1: GetAutoValidationMap
- **Status:** PASSED
- **Result:** Successfully loads 9 validation mappings
- **Bug Fixed:** Error #450 (missing Set keyword)

### ✅ Test 2: Debug Logger System  
- **Status:** PASSED
- **Result:** GlobalDebugOn controls debug output correctly
- **Bugs Fixed:** 
  - InitDebugFlags table structure mismatch
  - Form method name (setLMenuValCompletedCB)

---

## Test 3 Scope

**Purpose:** End-to-end validation with actual data

**What will be tested:**
1. Full validation workflow from start to finish
2. ValidationTrackerForm progress display
3. Error detection and reporting
4. Auto-correction functionality
5. Format application
6. Cancel button
7. Performance measurement

**Current modules ready:**
- ✅ AV_Core v2.1 COMPLETE (tested)
- ✅ AV_UI v2.1 Test2 (tested)
- ⏳ AV_Engine v2.1 (ready, not tested)
- ⏳ AV_Format v2.1 (ready, not tested)
- ⏳ AV_Validators v2.1 (ready, not tested)
- ⏳ AV_ValidationRules v2.1 (ready, not tested)

---

## Prerequisites for Test 3

### Data Setup
1. ✅ GlobalDebugOptions table configured
2. ✅ AutoValidationCommentPrefixMappingTable populated (9 validations)
3. ⏳ Target data table with test rows
4. ⏳ Validation rule tables populated:
   - GIWValidationTable
   - ElectricityPairValidation
   - PlumbingPairValidation
   - HeatSourcePairValidation
   - HeatSourceANYRefTable

### Configuration
1. ⏳ Set GlobalDebugOn = "ON" (for detailed logging)
2. ⏳ Prepare test data with known errors
3. ⏳ Prepare test data with valid values
4. ⏳ Prepare test data requiring auto-correction

---

## Test 3 Plan

### Phase 3.1: Basic Execution
**Test:** Run validation on 10-20 rows  
**Expected:**
- Form appears
- Progress updates every 10 rows
- Validation completes without errors
- Results logged in form

### Phase 3.2: Error Detection
**Test:** Rows with validation errors  
**Expected:**
- Errors flagged in drop columns
- Error formatting applied
- Key column formatted with highest priority
- Error messages clear and specific

### Phase 3.3: Auto-Correction
**Test:** Rows requiring auto-correction  
**Expected:**
- Corrections applied automatically
- Autocorrect formatting applied
- Drop column shows correction message
- Original values changed to corrected values

### Phase 3.4: Cancel Function
**Test:** Cancel validation mid-process  
**Expected:**
- Validation stops gracefully
- Partial results preserved
- Form reports cancellation
- No errors or crashes

### Phase 3.5: Performance
**Test:** Large dataset (100+ rows)  
**Expected:**
- Progress updates regular
- No timeout errors
- Reasonable completion time
- Memory stable

---

## Questions for User Before Test 3

1. **Test data ready?** Do you have a test dataset with:
   - Valid rows
   - Invalid rows (for error detection)
   - Rows needing auto-correction

2. **Validation rules populated?** Are the rule tables filled in?

3. **Expected behavior?** Any specific validation scenarios to test?

4. **Debug setting?** Set GlobalDebugOn = "ON" for detailed Test 3 logging?

---

## Next Steps

**When ready:**
1. Confirm prerequisites above
2. Set GlobalDebugOn = "ON"
3. Create Test3 module (similar to Test2)
4. Run RunFullValidationMaster
5. Review results
6. Fix any issues found

**Estimated time:** 15-30 minutes depending on data size

---

**Documentation Updated:** Master Documentation v2.3 created with all Test 1 & 2 results
