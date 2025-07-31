# Pipe Validation Test - All Buttons Show N/A Without Pipes

## Issue Fixed
Previously, single-value text (no pipes) would still show content in the first button. Now all buttons correctly show "N/A" when there are no pipe characters.

## Test Scenarios

### Test 1: Single Value (No Pipes)
**Data**: `Tourism WA`
**Expected Result**: All buttons show "N/A" and are disabled
- `"1: N/A"` (disabled, greyed)
- `"2: N/A"` (disabled, greyed)  
- `"3: N/A"` (disabled, greyed)
- ... (all buttons through 9)
- `"ID: N/A"` (disabled, greyed)

### Test 2: Single Value with Colon (No Pipes)
**Data**: `Tourism WA:TEST123`
**Expected Result**: 
- All segment buttons show "N/A" and are disabled
- Activation ID button shows "ID: TEST123" and is enabled

### Test 3: Valid Pipe-Delimited Data
**Data**: `FY24_26|Q1-4|Tourism WA`
**Expected Result**: 
- `"1: FY24_26"` (enabled)
- `"2: Q1-4"` (enabled)
- `"3: Tourism WA"` (enabled)
- `"4: N/A"` through `"9: N/A"` (disabled, greyed)
- `"ID: N/A"` (disabled, greyed)

### Test 4: After Extraction Workflow
**Initial Data**: `FY24_26|Q1-4|Tourism WA|WA|Marketing:TEST123`
**Step 1**: Form shows proper pipe-delimited previews
**Step 2**: Click "3: Tourism WA" button
**Step 3**: Cell becomes `Tourism WA` (no pipes)
**Expected**: Form refreshes to show all buttons as "N/A" except none are enabled

### Test 5: Edge Cases
**Empty String**: `""` → All buttons "N/A" 
**Just Pipes**: `|||` → Buttons show empty segments but are enabled if segments exist
**Just Colon**: `:TEST123` → All segment buttons "N/A", ID shows "TEST123"

## Technical Implementation

### Logic Change
**Before**:
```vba
' Split always returned at least one element, so Segment1 always got content
If UBound(segments) >= 0 Then result.Segment1 = Trim(segments(0))
```

**After**:
```vba
' Only parse segments if pipes are actually present
If InStr(mainContent, "|") > 0 Then
    ' Parse segments normally
Else
    ' Leave all segments empty (default to empty strings)
End If
```

### Expected Behavior Changes
- **No Pipes**: All segment buttons show "N/A" and are disabled
- **With Pipes**: Normal segment parsing and button enabling
- **Activation ID**: Works independently of pipe validation (colon-based)

## Benefits
✅ **Accurate State Representation**: Buttons reflect actual extractable segments
✅ **Clear Visual Feedback**: Users immediately see when no extractions are possible  
✅ **Prevents Confusion**: No misleading enabled buttons for non-segmented data
✅ **Consistent Logic**: Pipe requirement aligns with extraction functionality