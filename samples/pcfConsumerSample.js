// Sample form script to consume outputs from the PCF dataset control
// Attach these functions to your form or ribbon commands. Replace controlName with the PCF control name on the form.

/**
 * Reads selected IDs from the PCF control and opens the first record using Xrm.Navigation.openForm
 * @param {FormContext} formContext
 * @param {string} controlName
 */
/**
 * Read selected IDs from the PCF control outputs. Returns array of GUIDs (strings).
 */
function _readSelectedIds(control) {
    if (!control) return [];
    const outputs = control.getOutputs && control.getOutputs();
    if (!outputs) return [];
    const idsCsv = outputs.selectedRecordIds || outputs.selectedRecordId || '';
    if (!idsCsv) return [];
    return (typeof idsCsv === 'string') ? idsCsv.split(',').filter(Boolean) : [];
}

function openSelectedRecord(formContext, controlName) {
    try {
        const control = formContext.getControl(controlName);
        const ids = _readSelectedIds(control);
        if (!ids || ids.length === 0) { console.warn('No selected ids'); return; }
        const id = ids[0];
        const entityName = formContext.data.entity.getEntityName();
        Xrm.Navigation.openForm({ entityName: entityName, entityId: id }).catch(e=>console.error(e));
    } catch (e) { console.error(e); }
}

function openSelectedRecords(formContext, controlName) {
    try {
        const control = formContext.getControl(controlName);
        const ids = _readSelectedIds(control);
        if (!ids || ids.length === 0) { alert('No records selected'); return; }
        const entityName = formContext.data.entity.getEntityName();
        ids.forEach(id=>{
            Xrm.Navigation.openForm({ entityName: entityName, entityId: id }).catch(e=>console.error(e));
        });
    } catch (e) { console.error(e); }
}

function deleteSelectedRecords(formContext, controlName) {
    try {
        const control = formContext.getControl(controlName);
        const ids = _readSelectedIds(control);
        if (!ids || ids.length === 0) { alert('No records selected'); return; }
        if (!confirm('Delete ' + ids.length + ' selected record(s)? This cannot be undone.')) return;
        const entityName = formContext.data.entity.getEntityName();
        const promises = ids.map(id => Xrm.WebApi.deleteRecord(entityName, id));
        Promise.all(promises).then(()=>{ alert('Deleted selected records.'); location.reload(); }).catch(err=>{ console.error(err); alert('Error deleting records'); });
    } catch (e) { console.error(e); }
}

function editSelectedRecord(formContext, controlName) {
    try {
        const control = formContext.getControl(controlName);
        const ids = _readSelectedIds(control);
        if (!ids || ids.length === 0) { console.warn('No selected ids'); return; }
        const id = ids[0];
        const entityName = formContext.data.entity.getEntityName();
        Xrm.Navigation.openForm({ entityName: entityName, entityId: id }).catch(e=>console.error(e));
    } catch (e) { console.error(e); }
}

function assignSelectedToUser(formContext, controlName, userId) {
    try {
        const control = formContext.getControl(controlName);
        const ids = _readSelectedIds(control);
        if (!ids || ids.length === 0) { alert('No records selected'); return; }
        const entityName = formContext.data.entity.getEntityName();
        const promises = ids.map(id => Xrm.WebApi.updateRecord(entityName, id, { 'ownerid@odata.bind': '/systemusers(' + userId + ')' }));
        Promise.all(promises).then(()=>{ alert('Assigned selected records.'); }).catch(err=>{ console.error(err); alert('Error assigning records'); });
    } catch (e) { console.error(e); }
}

/**
 * Show a simple CSV export by fetching visible fields client-side (requires Web API calls for fields).
 */
function exportSelectedFromPCF(formContext, controlName) {
    try {
        const control = formContext.getControl(controlName);
        const ids = _readSelectedIds(control);
        if (!ids || ids.length === 0) { alert('No records selected'); return; }
        alert('Selected IDs: ' + ids.join(','));
    } catch (e) { console.error(e); }
}

/**
 * Enable rule helper for Ribbon: returns true if PCF control has selection.
 * Usage in Ribbon Workbench: add a JavaScript enable rule pointing to this function.
 */
function isPCFSelectionPresent(formContext, controlName) {
    try {
        const control = formContext.getControl(controlName);
        const ids = _readSelectedIds(control);
        return ids && ids.length > 0;
    } catch (e) { return false; }
}

module.exports = { openSelectedRecord, openSelectedRecords, editSelectedRecord, deleteSelectedRecords, assignSelectedToUser, exportSelectedFromPCF, isPCFSelectionPresent };