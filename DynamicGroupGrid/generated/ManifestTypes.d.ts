/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    pageSize: ComponentFramework.PropertyTypes.WholeNumberProperty;
    enablePagination: ComponentFramework.PropertyTypes.TwoOptionsProperty;
    sampleDataSet: ComponentFramework.PropertyTypes.DataSet;
}
export interface IOutputs {
    selectedRecordId?: string;
    selectedRecordIds?: string;
    rowEvent?: string;
}
