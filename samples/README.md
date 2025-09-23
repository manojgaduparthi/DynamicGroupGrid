# PCF DynamicGroupGrid Control - Consumer Samples

This folder contains sample code showing how to consume outputs from the PCF dataset control and integrate with D365 native functionality.

## Quick Start

1. **Upload Web Resource**: Import `pcfConsumerSample.js` as a Script web resource
2. **Add to Form**: Include it in form libraries where the PCF control is used
3. **Call Functions**: Use from ribbon commands or form buttons

## Available Functions

```javascript
// Open first selected record
pcfConsumerSample.openSelectedRecord(executionContext, "yourControlName")

// Open all selected records  
pcfConsumerSample.openSelectedRecords(executionContext, "yourControlName")

// Delete selected records (with confirmation)
pcfConsumerSample.deleteSelectedRecords(executionContext, "yourControlName")

// Check if selection exists (for ribbon enable rules)
pcfConsumerSample.isPCFSelectionPresent(executionContext, "yourControlName")
```

## Ribbon Integration

For command bar integration:
1. Use Ribbon Workbench or Solution Explorer
2. Create button with enable rule pointing to `isPCFSelectionPresent`  
3. Add action calling desired function (e.g., `openSelectedRecord`)
4. Pass `PrimaryControl` and PCF control name as parameters

## Notes

- PCF control exposes `selectedRecordIds` (CSV) and `selectedRecordId` outputs
- Functions handle multiple record selection automatically
- Error handling included for production use