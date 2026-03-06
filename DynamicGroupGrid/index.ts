import { IInputs, IOutputs } from "./generated/ManifestTypes";

interface ErrorContext {
    operation: string;
    severity: 'low' | 'medium' | 'high' | 'critical';
    recoverable: boolean;
    details?: Record<string, unknown>;
}

/**
 * DynamicGroupGrid PCF Control
 * 
 * Advanced grouped data grid component with responsive column layout,
 * intelligent field display, and optimized performance for large datasets.
 * 
 * Features:
 * - Dynamic column width distribution
 * - Grouped data visualization
 * - Intelligent pagination
 * - Responsive design
 * - Accessibility compliant
 * 
 * @version 3.9.12
 * @author Community PCF Control
 */
export class DynamicGroupGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    
    /**
     * CSS class name constants for consistent styling
     */
    private static readonly CSS = {
        ROOT: 'dynamic-group-grid-root',
        TOOLBAR: 'dynamic-group-grid-toolbar',
        GRID_HEADER: 'pcf-grid-header',
        GRID_COL: 'pcf-grid-col',
        GROUP_SECTION: 'pcf-group-section',
        GROUP_HEADER: 'pcf-group-header',
        GROUP_LIST: 'dynamic-group-grid-list',
        GROUP_ROW: 'pcf-group-row',
        CELL: 'pcf-cell',
        PAGINATION_CONTAINER: 'pcf-pagination-container',
        PAGINATION_INFO: 'pcf-pagination-info',
        PAGINATION_CONTROLS: 'pcf-pagination-controls',
        PAGINATION_BUTTON: 'pcf-pagination-button'
    } as const;

    /**
     * Configuration constants for layout and behavior
     */
    private static readonly CONFIG = {
        /** Default column width in pixels */
        DEFAULT_UNIFORM_WIDTH: 150,
        /** Minimum column width to prevent content cutoff */
        DEFAULT_MIN_WIDTH: 40,
        /** CSS flex fallback for responsive columns */
        DEFAULT_FALLBACK_FLEX: '1 1 120px',
        /** Default number of records per page */
        DEFAULT_PAGE_SIZE: 25,
        /** Maximum records to load per page */
        MAX_PAGE_SIZE: 5000,
        /** Reserved space for container padding and scrollbars */
        CONTAINER_PADDING: 50,
        /** Reserved layout offset for action columns and gutters */
        RESERVED_LAYOUT_OFFSET: 250,
        /** Proportion of width for last column */
        LAST_COLUMN_WIDTH_RATIO: 0.35,
        /** Proportion of width for other columns */
        OTHER_COLUMNS_WIDTH_RATIO: 0.65
    } as const;

    // runtime state for optimized updates
    private readonly _listenerDisposables: (() => void)[] = [];
    private _selectedGroupColumn: string | null = null;
    private _expandedGroups: Map<string, boolean> = new Map();
    private _selectedRecordId: string | null = null;
    private _notifyOutputChanged: (() => void) | null = null;
    private _selectedRecordIds: string[] = [];
    private _lastRowEvent: string | null = null;
    private _sortColumn: string | null = null;
    private _sortAsc = true;
    private readonly _colWidths = new Map<string, number>();
    private _colUniformWidth = 150;
    private _currentPage = 1;
    private _pageSize: number = DynamicGroupGrid.CONFIG.DEFAULT_PAGE_SIZE;
    private _totalRecords = 0;
    private _enablePagination = true;
    private _lastContext: ComponentFramework.Context<IInputs> | null = null;
    private _container!: HTMLDivElement;
    private _dataset!: ComponentFramework.PropertyTypes.DataSet;
    private readonly _lastAllocatedWidth = 0;

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this._container = document.createElement("div");
        this._container.className = DynamicGroupGrid.CSS.ROOT;
        container.appendChild(this._container);

        try {
            if (typeof context.mode?.trackContainerResize === 'function') {
                try {
                    context.mode.trackContainerResize(true);
                } catch { /* ignore */ }
            }
        } catch { /* ignore */ }

        this._selectedGroupColumn = null;
        this._expandedGroups = new Map();
        this._selectedRecordId = null;
        this._selectedRecordIds = [];
        this._notifyOutputChanged = notifyOutputChanged;
        this._currentPage = 1;
        this._pageSize = DynamicGroupGrid.CONFIG.DEFAULT_PAGE_SIZE;
        this._enablePagination = true;
        this._lastContext = null;
        this._dataset = context.parameters.sampleDataSet;

        // Configure pagination with large page size for optimal performance
        if (context.parameters.sampleDataSet?.paging) {
            context.parameters.sampleDataSet.paging.setPageSize(5000);
        }

        // Register window resize listener for responsive layout adjustments
        try {
            this.addListener(globalThis, 'resize', () => {
                try { this.adjustCollapsedHeaderWidths(); } catch { /* ignore */ }
            });
        } catch { /* ignore */ }
    }

    /**
     * Helper method to register and track DOM event listeners for proper cleanup.
     * @param el Element to attach listener to
     * @param evt Event type
     * @param fn Event handler function
     * @param options Optional event listener options
     */
    private addListener(el: EventTarget, evt: string, fn: EventListener, options?: boolean | AddEventListenerOptions): void {
        el.addEventListener(evt, fn, options);
        this._listenerDisposables.push(() => {
            try {
                el.removeEventListener(evt, fn, options);
            } catch { /* ignore */ }
        });
    }

    /**
     * Applies column width using proportional distribution algorithm.
     * @param el HTML element to apply width to
     * @param columnName Name of the column for width calculation
     */
    private applyColumnWidth(el: HTMLElement, columnName?: string): void {
        try {
            if (!columnName || !this._dataset || !this._lastContext) {
                return;
            }

            const columns = Array.from(this._dataset.columns);
            const columnIndex = columns.findIndex((col: ComponentFramework.PropertyHelper.DataSetApi.Column) => col.name === columnName || col.alias === columnName);
            
            if (columnIndex === -1) {
                return;
            }

            // Universal proportional approach - adapts to any container width and column count
            const containerWidth = this._lastContext?.mode?.allocatedWidth || 611;
            const availableWidth = containerWidth - DynamicGroupGrid.CONFIG.CONTAINER_PADDING;
            const columnCount = columns.length;
            
            // Calculate proportional width distribution
            let baseWidth: number;
            if (columnIndex === columns.length - 1) {
                // Last column gets larger allocation for longer content
                baseWidth = Math.floor(availableWidth * DynamicGroupGrid.CONFIG.LAST_COLUMN_WIDTH_RATIO);
            } else {
                // Other columns share remaining space equally
                const otherColumnsShare = availableWidth * DynamicGroupGrid.CONFIG.OTHER_COLUMNS_WIDTH_RATIO;
                baseWidth = Math.floor(otherColumnsShare / (columnCount - 1));
            }
            
            // Apply width constraints
            const minWidth = 80;
            const maxWidth = 250;
            const columnWidth = Math.max(minWidth, Math.min(maxWidth, baseWidth)) + 'px';
            
            el.style.width = columnWidth;
            el.style.maxWidth = columnWidth;
            el.style.minWidth = '80px';
            el.style.flex = 'none';               // No flex, fixed sizes
            el.style.overflow = 'hidden';         // Contain content
            el.style.textOverflow = 'ellipsis';   // Show ... for long content
            el.style.whiteSpace = 'nowrap';       // Keep text on one line
            el.style.textOverflow = 'clip';
        } catch (e) {
            this.handleError('applyColumnWidth', e);
        }
    }

    /**
     * Calculates proportional width distribution with the last column getting remaining space.
     * Implements robust width calculation with proper validation and error handling.
     */
    private getColumnWidthDistribution(): string[] | null {
        try {
            if (!this._dataset || !this._lastContext?.mode) {
                return null;
            }

            const columns = Array.from(this._dataset.columns);
            if (!columns.length) {
                return null;
            }

            const widthDistribution: string[] = [];
            
            // Validate allocated width
            const allocatedWidth = this._lastContext.mode.allocatedWidth;
            if (!allocatedWidth || allocatedWidth <= DynamicGroupGrid.CONFIG.CONTAINER_PADDING) {
                return null;
            }

            // Calculate available width minus padding
            const totalWidth = allocatedWidth - DynamicGroupGrid.CONFIG.RESERVED_LAYOUT_OFFSET;
            let widthSum = 0;

            // Calculate total visual size with validation
            for (const col of columns) {
                const visualSize = col.visualSizeFactor || 1; // Default to 1 if undefined
                widthSum += visualSize;
            }

            if (widthSum <= 0) {
                return null; // Prevent division by zero
            }

            let remainWidth = totalWidth;

            // Distribute widths proportionally
            for (let index = 0; index < columns.length; index++) {
                const col = columns[index];
                const visualSize = col.visualSizeFactor || 1;
                let widthPerCell = "";
                
                if (index === columns.length - 1) {
                    // Last column gets remaining space (minimum width protected)
                    const finalWidth = Math.max(remainWidth, DynamicGroupGrid.CONFIG.DEFAULT_MIN_WIDTH);
                    widthPerCell = finalWidth + "px";
                } else {
                    // Not last column - proportional width
                    const cellWidth = Math.max(
                        Math.round((visualSize / widthSum) * totalWidth),
                        DynamicGroupGrid.CONFIG.DEFAULT_MIN_WIDTH
                    );
                    remainWidth = Math.max(remainWidth - cellWidth, 0);
                    widthPerCell = cellWidth + "px";
                }
                
                widthDistribution.push(widthPerCell);
            }

            return widthDistribution;
        } catch (error) {
            this.handleError('calculating column width distribution', error);
            return null;
        }
    }

    /**
     * Handles column width adjustments for collapsed headers to maintain layout consistency.
     */
    private updateAllColumnWidths(): void {
        // No custom width calculations needed - let browser handle everything
        // This eliminates last column truncation issues
    }







    /**
     * Centralized error handling for the component.
     * Provides structured error logging and classification.
     * 
     * @param operation - Description of the operation that failed
     * @param error - The error that occurred
     * @param severity - Error severity level
     * @param recoverable - Whether the error allows continued operation
     */
    private handleError(
        operation: string, 
        error: unknown, 
        severity: 'low' | 'medium' | 'high' | 'critical' = 'medium',
        recoverable = true
    ): void {
        if (console?.error) {
            const errorMessage = error instanceof Error ? error.message : String(error);
            const context: ErrorContext = {
                operation,
                severity,
                recoverable,
                details: { 
                    timestamp: new Date().toISOString(),
                    component: 'DynamicGroupGrid',
                    version: '1.0.0'
                }
            };
            console.error(`[DynamicGroupGrid] ${operation} failed:`, errorMessage, context);
        }
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets,
     * global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._lastContext = context;
        this._dataset = context.parameters.sampleDataSet;

        // Browser handles width changes naturally

        const dataset = context.parameters.sampleDataSet;

        // Use dataset if available; early return if no data
        if (!dataset?.records || Object.keys(dataset.records).length === 0) {
            return;
        }

        // Handle pagination configuration
        if (dataset.paging?.hasNextPage) {
            try {
                dataset.paging.loadNextPage();
            } catch (error) {
                this.handleError("loading next page", error);
            }
            return;
        }

        const ds = dataset;

        // Read pagination configuration from context
        try {
            if (context.parameters.pageSize?.raw != null) {
                const configuredPageSize = Number(context.parameters.pageSize.raw);
                if (configuredPageSize > 0 && configuredPageSize <= DynamicGroupGrid.CONFIG.MAX_PAGE_SIZE) {
                    this._pageSize = configuredPageSize;
                }
            }

            if (context.parameters.enablePagination?.raw != null) {
                this._enablePagination = Boolean(context.parameters.enablePagination.raw);
            }
        } catch { /* ignore */ }

        // Read columns - only show columns that are configured in the view (like OOB subgrid)  
        const allColumns = ds.columns ? Object.values(ds.columns) : [];
        

        
        // Sort columns by order first (Microsoft best practice)
        // Create a copy before sorting to avoid mutating the original array
        const sortedColumns = [...allColumns].sort((a, b) => (a.order || 0) - (b.order || 0));
        
        // Microsoft PCF Best Practice: Respect view column configuration
        // Columns should only be displayed if they are explicitly included in the view
        const columns = sortedColumns.filter(col => {
            // Must have basic metadata
            const hasValidMetadata = col.name && col.displayName;
            
            // Microsoft Standard: Only show columns that have a valid order (indicating view inclusion)
            // Columns not in view typically have order = -1 or undefined
            const isInViewConfiguration = typeof col.order === 'number' && col.order >= 0;
            
            // Respect visibility settings
            const isVisible = col.isHidden !== true;
            
            // Exclude system fields
            const isNotSystemField = !col.name.startsWith('_') && 
                                    !col.name.startsWith('msft_') && 
                                    !col.name.includes('_base') &&
                                    col.name !== 'versionnumber' &&
                                    col.name !== 'timezoneruleversionnumber';
            
            return hasValidMetadata && isInViewConfiguration && isVisible && isNotSystemField;
        });



        // choose default group column: first column
        if (!this._selectedGroupColumn && columns.length > 0) {
            this._selectedGroupColumn = columns[0].name;
        }

        // show all columns; provide a slider to control uniform column width for horizontal exploration
        const colsToShowCount = columns.length;

        // restore persisted column widths (localStorage key based on control name)
        try {
            const key = 'pcf_col_widths_DynamicGroupGrid';
            const raw = localStorage.getItem(key);
            if (raw) {
                const parsed = JSON.parse(raw);
                for (const k of Object.keys(parsed)) {
                    this._colWidths.set(k, parsed[k]);
                }
            }

            // Load uniform width preference from state
            try {
                const uw = localStorage.getItem('dynamic_group_grid_uniform_width');
                if (uw) this._colUniformWidth = Number(uw) || this._colUniformWidth;
            } catch { /* ignore */ }
        } catch {
            // Ignore localStorage errors
        }

        // clear
        this._container.innerHTML = "";

        // build group-by dropdown
        const toolbar = document.createElement("div");
        toolbar.className = DynamicGroupGrid.CSS.TOOLBAR;

        const label = document.createElement("label");
        label.textContent = "Group by: ";
        toolbar.appendChild(label);

        const select = document.createElement("select");
        select.className = "dynamic-group-grid-select";

        for (const col of columns) {
            const opt = document.createElement("option");
            opt.value = col.name;
            // Use same Microsoft-recommended display name logic
            if (col.displayName?.trim()) {
                opt.text = col.displayName.trim();
            } else if (col.alias?.trim()) {
                opt.text = col.alias.trim();
            } else {
                // Use logical name as-is - Power Platform handles proper display names
                opt.text = col.name || '';
            }
            if (col.name === this._selectedGroupColumn) opt.selected = true;
            select.appendChild(opt);
        }

        this.addListener(select, "change", (evt: Event) => {
            const val = (evt.target as HTMLSelectElement).value;
            this._selectedGroupColumn = val;
            this._currentPage = 1; // Reset to first page when changing grouping
            this.updateView(context);
        });

        toolbar.appendChild(select);

        // Horizontal scrolling is handled by native browser scrollbar
        this._container.appendChild(toolbar);

        // Sort dataset records alphabetically by group column
        const entries = ds.sortedRecordIds ? ds.sortedRecordIds.map(id => ({
            id: String(id),
            rec: ds.records[id]
        })) : Object.keys(ds.records || {}).map(k => ({
            id: String(k),
            rec: ds.records[k]
        }));

        // Sort entries alphabetically by the group column value
        entries.sort((a, b) => {
            const valA = this.getRecordFieldValue(a.rec, this._selectedGroupColumn || '') || "(blank)";
            const valB = this.getRecordFieldValue(b.rec, this._selectedGroupColumn || '') || "(blank)";
            return valA.toString().localeCompare(valB.toString());
        });

        // Group the sorted entries by group column value
        const groups: Record<string, Array<{id: string; rec: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord}>> = {};
        for (const entry of entries) {
            const val = this.getRecordFieldValue(entry.rec, this._selectedGroupColumn || '') || "(blank)";
            if (!groups[val]) groups[val] = [];
            groups[val].push(entry);
        }

        this._totalRecords = entries.length;
        const allRecordIds = entries.map(e => e.id);

        // render a single header that shows columns
        const headerRow = document.createElement("div");
        headerRow.className = "pcf-grid-header";
        headerRow.style.width = '100%';          // Ensure header respects container width
        headerRow.style.maxWidth = '100%';
        
        // Add accessibility attributes for header
        headerRow.setAttribute('role', 'row');

        // add selection header checkbox
        const selectAllHeader = document.createElement('div');
        selectAllHeader.className = 'pcf-grid-col';
        const selectAllCb = document.createElement('input');
        selectAllCb.type = 'checkbox';
        selectAllCb.checked = !!(this._selectedRecordIds && this._selectedRecordIds.length > 0 && allRecordIds.length > 0 && this._selectedRecordIds.length === allRecordIds.length);
        
        this.addListener(selectAllCb, 'change', (evt: Event) => {
            evt.stopPropagation();
            if (!this._selectedRecordIds) this._selectedRecordIds = [];
            const checked = (evt.target as HTMLInputElement).checked;
            if (checked) {
                this._selectedRecordIds = allRecordIds.slice();
            } else {
                this._selectedRecordIds = [];
            }
            this._selectedRecordId = this._selectedRecordIds.length ? 
                this._selectedRecordIds[this._selectedRecordIds.length - 1] : null;
            
            // sync selection with host dataset so native ribbon sees it (pass dataset row ids)
            this._syncSelectionToHost(context, this._selectedRecordIds || []);
            if (this._notifyOutputChanged) this._notifyOutputChanged();
            
            // avoid full re-render; update checkboxes and classes
            this.updateSelectionVisuals();
        });
        
        selectAllHeader.appendChild(selectAllCb);
        headerRow.appendChild(selectAllHeader);

        // column headers
        const colsToShow = columns.slice(0, colsToShowCount);
        
        for (const c of colsToShow) {
            const hc = document.createElement('div');
            hc.className = 'pcf-grid-col';
            
            // apply stored/uniform width to header cell to match row cells
            try {
                this.applyColumnWidth(hc, (c.name || '').toLowerCase());
            } catch { /* ignore */ }
            
            // Microsoft recommended approach: Use displayName, fallback to alias, then name
            // Power Platform should provide proper displayName from entity metadata
            let headerText = '';
            if (c.displayName?.trim()) {
                headerText = c.displayName.trim();
            } else if (c.alias?.trim()) {
                headerText = c.alias.trim();
            } else {
                // Use name as-is (Microsoft handles localization at platform level)
                headerText = c.name || '';
            }
            

            
            hc.textContent = headerText;
            hc.title = `Sort by ${headerText}`;
            hc.dataset.colName = (c.name || '').toLowerCase();
            
            // Add accessibility attributes for column headers
            hc.setAttribute('role', 'columnheader');
            hc.setAttribute('tabindex', '0');
            let sortState = 'none';
            if (this._sortColumn === c.name) {
                sortState = this._sortAsc ? 'ascending' : 'descending';
            }
            hc.setAttribute('aria-sort', sortState);
            
            this.addListener(hc, 'click', (evt: Event) => {
                evt.stopPropagation();
                if (this._sortColumn === c.name) {
                    this._sortAsc = !this._sortAsc;
                } else {
                    this._sortColumn = c.name;
                    this._sortAsc = true;
                }
                this.updateView(context);
            });

            // sort indicator
            if (this._sortColumn === c.name) {
                const ind = document.createElement('span');
                ind.className = 'pcf-sort-ind';
                ind.textContent = this._sortAsc ? ' ▲' : ' ▼';
                hc.appendChild(ind);
            }

            // add resize handle
            const handle = document.createElement('div');
            handle.className = 'pcf-col-handle';
            let isDragging = false;

            this.addListener(handle, 'mousedown', (startEv: Event) => {
                const mouseEv = startEv as MouseEvent;
                mouseEv.preventDefault();
                mouseEv.stopPropagation();
                isDragging = true;

                const startX = mouseEv.clientX;
                const headerCell = hc;
                const startRect = headerCell.getBoundingClientRect();
                const startWidth = startRect.width;
                const colName = c.name;

                const onMove = (moveEv: Event) => {
                    if (!isDragging) return;
                    const mEv = moveEv as MouseEvent;
                    const delta = mEv.clientX - startX;
                    const newW = Math.max(40, Math.round(startWidth + delta));
                    headerCell.style.flex = `0 0 ${newW}px`;
                    this._colWidths.set(colName, newW);

                    // apply to existing row cells in the DOM for visual feedback
                    try {
                        const headerIndex = Array.prototype.indexOf.call(headerRow.children, headerCell);
                        const rows = this._container.querySelectorAll('.pcf-group-row');
                        for (const rEl of Array.from(rows)) {
                            const child = rEl.children[headerIndex] as HTMLElement;
                            if (child) this.applyColumnWidth(child, c.name.toLowerCase());
                        }
                    } catch {
                        // Ignore errors during resize
                    }
                };

                const onUp = () => {
                    isDragging = false;
                    try {
                        document.removeEventListener('mousemove', onMove);
                    } catch { /* ignore */ }
                    try {
                        document.removeEventListener('mouseup', onUp);
                    } catch { /* ignore */ }

                    // persist widths to localStorage
                    try {
                        const _key = 'pcf_col_widths_DynamicGroupGrid';
                        const obj: Record<string, number> = {};
                        for (const [k, v] of this._colWidths.entries()) {
                            obj[k] = v;
                        }
                        localStorage.setItem(_key, JSON.stringify(obj));
                    } catch {
                        // Ignore localStorage errors
                    }

                    // final updateSelectionVisuals to ensure layout applied
                    this.updateSelectionVisuals();
                };

                this.addListener(document, 'mousemove', onMove);
                this.addListener(document, 'mouseup', onUp);
            });

            hc.appendChild(handle);
            headerRow.appendChild(hc);
        }

        // create a horizontally scrollable wrapper for header + groups so all fields can be seen
        const gridWrapper = document.createElement('div');
        gridWrapper.className = 'pcf-grid-scroll-wrap';
        gridWrapper.style.overflowX = 'auto';
        gridWrapper.style.display = 'block';
        gridWrapper.style.width = '100%';           // Constrain to container width
        gridWrapper.style.maxWidth = '100%';        // Prevent overflow beyond container
        
        // Add accessibility attributes
        gridWrapper.setAttribute('role', 'grid');
        gridWrapper.setAttribute('aria-label', 'Dynamic data grid with grouping');
        gridWrapper.setAttribute('tabindex', '0');
        gridWrapper.appendChild(headerRow);
        this._container.appendChild(gridWrapper);

        // determine ordered list of group keys so group sections render in a predictable sort order
        const groupKeys = Object.keys(groups || {});
        
        try {
            // sort case-insensitive and numeric-aware
            groupKeys.sort((a = '', b = '') => {
                const cmp = a.localeCompare(b, undefined, {
                    sensitivity: 'base',
                    numeric: true
                });
                // if the current sort column is the same as the group-by column, respect the sort direction
                if (this._sortColumn === this._selectedGroupColumn) return this._sortAsc ? cmp : -cmp;
                // otherwise default to ascending by group key
                return cmp;
            });
        } catch { /* ignore sort errors and fall back to insertion order */ }

        for (const groupKey of groupKeys) {
            const groupDiv = document.createElement("div");
            groupDiv.className = "pcf-group-section";

            const header = document.createElement("div");
            header.className = "pcf-group-header";

            const expandButton = document.createElement("span");
            expandButton.className = "expand-toggle-btn";

            // all groups start collapsed; user can click to expand/collapse
            if (!this._expandedGroups.has(groupKey)) {
                this._expandedGroups.set(groupKey, false);
            }

            const isExpanded = this._expandedGroups.get(groupKey) === true;
            expandButton.textContent = isExpanded ? "−" : "+";
            expandButton.dataset.expanded = isExpanded.toString();
            header.appendChild(expandButton);

            const title = document.createElement("div");
            title.textContent = `${groupKey} (${groups[groupKey].length})`;
            header.appendChild(title);

            // Use addListener to ensure consistent cleanup tracking
            this.addListener(header, 'click', (evt: Event) => {
                evt.preventDefault();
                evt.stopPropagation();
                const cur = this._expandedGroups.get(groupKey) || false;
                this._expandedGroups.set(groupKey, !cur);
                this.updateView(context);
            });

            groupDiv.appendChild(header);

            const list = document.createElement("div");
            list.className = DynamicGroupGrid.CSS.GROUP_LIST;

            const expanded = this._expandedGroups.get(groupKey) === true;

            // hide the list entirely when collapsed to avoid rendering an empty placeholder row
            if (!expanded) {
                // CSS handles display: none with !important rule - no inline style needed
                // mark parent as collapsed for CSS targeting - use unique class name
                groupDiv.classList.add('pcf-group-collapsed');
            }

            if (expanded) {
                // optionally sort group rows
                const grpRows = groups[groupKey].slice();
                if (this._sortColumn) {
                    const sc = this._sortColumn;
                    grpRows.sort((a, b) => {
                        const va = this.getRecordFieldValue(a.rec, sc) || '';
                        const vb = this.getRecordFieldValue(b.rec, sc) || '';
                        if (va < vb) return this._sortAsc ? -1 : 1;
                        if (va > vb) return this._sortAsc ? 1 : -1;
                        return 0;
                    });
                }

                for (const entry of grpRows) {
                    const row = document.createElement("div");
                    row.className = "pcf-group-row";
                    const recordId = entry.id;
                    const r = entry.rec;

                    if (entry.id) row.dataset.recordId = entry.id;
                    
                    // Add accessibility attributes for data rows
                    row.setAttribute('role', 'row');
                    row.setAttribute('tabindex', '0');
                    const isSelected = !!(this._selectedRecordIds && recordId && this._selectedRecordIds.includes(recordId));
                    if (isSelected) {
                        row.setAttribute('aria-selected', 'true');
                    }

                    if (recordId && this._selectedRecordIds?.includes(recordId)) {
                        row.classList.add("selected");
                    }

                    // add selection checkbox cell
                    const selectCell = document.createElement('div');
                    selectCell.className = 'pcf-grid-col';
                    selectCell.dataset.colName = 'select';
                    const cb = document.createElement('input');
                    cb.type = 'checkbox';
                    cb.checked = !!(this._selectedRecordIds && recordId && this._selectedRecordIds.includes(recordId));

                    // prevent clicks on the checkbox from bubbling to the row click handler
                    this.addListener(cb, 'click', (evt: Event) => {
                        evt.stopPropagation();
                    });

                    this.addListener(cb, 'change', (evt: Event) => {
                        evt.stopPropagation();
                        if (!this._selectedRecordIds) this._selectedRecordIds = [];
                        const checked = (evt.target as HTMLInputElement).checked;
                        const idx = recordId ? this._selectedRecordIds.indexOf(recordId) : -1;
                        if (checked && idx === -1 && recordId) this._selectedRecordIds.push(recordId);
                        if (!checked && idx !== -1) this._selectedRecordIds.splice(idx, 1);
                        this._selectedRecordId = this._selectedRecordIds.length ? 
                            this._selectedRecordIds[this._selectedRecordIds.length - 1] : null;
                        
                        // sync selection with host dataset
                        this._syncSelectionToHost(context, this._selectedRecordIds || []);
                        if (this._notifyOutputChanged) this._notifyOutputChanged();
                        
                        // update visuals only
                        this.updateSelectionVisuals();
                    });

                    selectCell.appendChild(cb);
                    row.appendChild(selectCell);

                    // render a few columns values - use visible columns from dataset
                    const colsToShow = columns.slice(0, colsToShowCount);
                    for (const c of colsToShow) {
                        const cellDiv = document.createElement('div');
                        cellDiv.className = 'pcf-grid-col';
                        
                        // expose column name as attribute for targeted CSS
                        cellDiv.dataset.colName = (c.name || '').toLowerCase();
                        
                        // apply uniform/stored width via helper
                        try {
                            this.applyColumnWidth(cellDiv, (c.name || '').toLowerCase());
                        } catch { /* ignore */ }

                        const inner = document.createElement('span');
                        inner.className = 'pcf-cell';
                        const text = this.getRecordFieldValue(r, c.name) || '';

                        // Check if this is a lookup column and create appropriate display
                        const isLookup = this.isLookupColumn(c);
                        
                        if (isLookup) {
                            // For lookup fields, create a clickable link regardless of column position
                            const lookupRef = this.getLookupEntityReference(r, c.name);

                            
                            if (lookupRef?.id && text) {
                                const a = document.createElement('a');
                                
                                // Extract GUID from EntityReference
                                const entityId = typeof lookupRef.id === 'object' && lookupRef.id.guid 
                                    ? lookupRef.id.guid 
                                    : lookupRef.id.toString();
                                
                                // Set proper href and title for hover behavior
                                a.href = '#';
                                a.title = `Open ${lookupRef.name || text} (${lookupRef.logicalName}: ${entityId})`;
                                a.textContent = text;
                                a.className = 'pcf-lookup-link';
                                
                                this.addListener(a, 'click', (evt: any) => {
                                    evt.preventDefault();
                                    evt.stopPropagation();
                                    

                                    
                                    try {
                                        // Navigate to the lookup record using D365 navigation - CORRECTED FORMAT
                                        if (context.navigation) {
                                            context.navigation.openForm({
                                                entityName: lookupRef.logicalName,
                                                entityId: entityId  // Use extracted GUID string, not object
                                            });
                                        } else {
                                            // Emit event for manual handling if navigation fails
                                            this._lastRowEvent = JSON.stringify({
                                                type: 'lookupOpen',
                                                id: entityId,
                                                entityName: lookupRef.logicalName,
                                                fieldName: c.name
                                            });
                                            if (this._notifyOutputChanged) this._notifyOutputChanged();
                                        }
                                    } catch (error) {
                                        // Log detailed error for developers only - do not expose to output
                                        this.handleError('lookupNavigation', error, 'high', false);

                                        // Emit a safe, generic error event - no raw error details exposed
                                        this._lastRowEvent = JSON.stringify({
                                            type: 'lookupOpenFailed',
                                            id: entityId,
                                            entityName: lookupRef.logicalName,
                                            fieldName: c.name,
                                            errorCode: 'NAVIGATION_FAILED'
                                        });
                                        if (this._notifyOutputChanged) this._notifyOutputChanged();
                                    }
                                });
                                inner.appendChild(a);
                            } else {
                                // Display as plain text if no valid lookup reference
                                inner.textContent = text;
                            }
                        } else {
                            // Display non-lookup fields as plain text
                            inner.textContent = text;
                        }

                        cellDiv.appendChild(inner);
                        
                        // also set data attribute on inner for CSS selectors
                        inner.dataset.colName = (c.name || '').toLowerCase();
                        row.appendChild(cellDiv);
                    }

                    this.addListener(row, 'click', (evt: any) => {
                        // stop group toggle
                        evt.stopPropagation();
                        
                        // set multi-selection (Ctrl/Cmd toggle, Shift range)
                        if (!this._selectedRecordIds) this._selectedRecordIds = [];
                        
                        if (recordId) {
                            const isCtrl = evt.ctrlKey || evt.metaKey;
                            const isShift = evt.shiftKey;
                            
                            if (isShift && this._selectedRecordIds.length > 0) {
                                // select range from last selected to this
                                const last = this._selectedRecordIds[this._selectedRecordIds.length - 1];
                                const start = allRecordIds.indexOf(last);
                                const end = allRecordIds.indexOf(recordId);
                                if (start >= 0 && end >= 0) {
                                    const [s, e] = start <= end ? [start, end] : [end, start];
                                    const range = allRecordIds.slice(s, e + 1);
                                    this._selectedRecordIds = Array.from(new Set(this._selectedRecordIds.concat(range)));
                                }
                            } else if (isCtrl) {
                                // toggle
                                const idx = this._selectedRecordIds.indexOf(recordId);
                                if (idx === -1) this._selectedRecordIds.push(recordId);
                                else this._selectedRecordIds.splice(idx, 1);
                            } else {
                                // single select
                                this._selectedRecordIds = [recordId];
                            }
                            
                            // update primary single output for backward compatibility
                            this._selectedRecordId = this._selectedRecordIds.length ? 
                                this._selectedRecordIds[this._selectedRecordIds.length - 1] : null;
                            if (this._notifyOutputChanged) this._notifyOutputChanged();
                            
                            // sync selection with host dataset
                            this._syncSelectionToHost(context, this._selectedRecordIds || []);
                            
                            // avoid full re-render; update selection visuals
                            this.updateSelectionVisuals();
                        }
                    });

                    this.addListener(row, 'dblclick', (evt: any) => {
                        evt.stopPropagation();
                        if (recordId) {
                            try {
                                this.navigateToRecord(context, recordId);
                            } catch {
                                this._lastRowEvent = JSON.stringify({
                                    type: 'doubleClick',
                                    id: recordId
                                });
                                if (this._notifyOutputChanged) this._notifyOutputChanged();
                            }
                        }
                    });

                    list.appendChild(row);
                }
            }

            groupDiv.appendChild(list);
            gridWrapper.appendChild(groupDiv);
        }

        // OOB pagination: Let the dataset framework handle pagination automatically
        // No custom pagination controls needed - framework provides this

        // final sizing adjustment: ensure collapsed headers visually span the scrollable content
        try {
            this.adjustCollapsedHeaderWidths();
        } catch { /* ignore */ }

        // No forced width recalculations needed
    }

    /**
     * Gets paginated subset of entries based on current page settings.
     * @param allEntries Complete array of entries to paginate
     * @returns Paginated subset of entries
     */
    private getPaginatedEntries<T>(allEntries: T[]): T[] {
        if (!this._enablePagination || this._pageSize <= 0) {
            return allEntries;
        }

        const startIndex = (this._currentPage - 1) * this._pageSize;
        const endIndex = startIndex + this._pageSize;
        const result = allEntries.slice(startIndex, endIndex);
        return result;
    }

    // Render pagination controls
    private renderPaginationControls(): void {
        if (!this._enablePagination || this._totalRecords <= this._pageSize) {
            return;
        }

        const totalPages = Math.ceil(this._totalRecords / this._pageSize);
        if (totalPages <= 1) {
            return;
        }

        const paginationContainer = document.createElement('div');
        paginationContainer.className = DynamicGroupGrid.CSS.PAGINATION_CONTAINER;

        // Pagination info
        const paginationInfo = document.createElement('div');
        paginationInfo.className = DynamicGroupGrid.CSS.PAGINATION_INFO;
        const startRecord = (this._currentPage - 1) * this._pageSize + 1;
        const endRecord = Math.min(this._currentPage * this._pageSize, this._totalRecords);
        paginationInfo.textContent = `Page ${this._currentPage} of ${totalPages} - Showing records ${startRecord}-${endRecord} of ${this._totalRecords}`;
        paginationContainer.appendChild(paginationInfo);

        // Pagination controls
        const paginationControls = document.createElement('div');
        paginationControls.className = DynamicGroupGrid.CSS.PAGINATION_CONTROLS;

        // First page button
        const firstButton = this.createPaginationButton('<<', 1, this._currentPage === 1);
        paginationControls.appendChild(firstButton);

        // Previous page button
        const prevButton = this.createPaginationButton('<', this._currentPage - 1, this._currentPage === 1);
        paginationControls.appendChild(prevButton);

        // Page number buttons
        const startPage = Math.max(1, this._currentPage - 2);
        const endPage = Math.min(totalPages, this._currentPage + 2);
        for (let i = startPage; i <= endPage; i++) {
            const pageButton = this.createPaginationButton(String(i), i, false, i === this._currentPage);
            paginationControls.appendChild(pageButton);
        }

        // Next page button
        const nextButton = this.createPaginationButton('>', this._currentPage + 1, this._currentPage === totalPages);
        paginationControls.appendChild(nextButton);

        // Last page button
        const lastButton = this.createPaginationButton('>>', totalPages, this._currentPage === totalPages);
        paginationControls.appendChild(lastButton);

        paginationContainer.appendChild(paginationControls);
        this._container.appendChild(paginationContainer);
        
        // Add click handlers after buttons are in DOM to prevent them being lost
        setTimeout(() => {
            this.attachPaginationHandlers();
        }, 0);
    }

    // Create a pagination button
    private createPaginationButton(text: string, targetPage: number, disabled: boolean, isActive = false): HTMLButtonElement {
        const button = document.createElement('button');
        button.className = DynamicGroupGrid.CSS.PAGINATION_BUTTON;
        button.textContent = text;
        button.disabled = disabled;

        if (isActive) {
            button.classList.add('active');
        }

        // Onclick handler will be added later via attachPaginationHandlers()

        return button;
    }

    /**
     * Attaches click handlers to pagination buttons after they're rendered in the DOM.
     */
    private attachPaginationHandlers(): void {
        const paginationButtons = this._container.querySelectorAll('button.pcf-pagination-button:not([disabled])');
        
        for (const button of Array.from(paginationButtons)) {
            const btn = button as HTMLButtonElement;
            const text = btn.textContent || '';
            
            // Skip if already has handler
            if (btn.dataset.hasOnclick === 'true') {
                continue;
            }
            
            let targetPage = this._currentPage;
            
            // Determine target page based on button text
            if (text === '<<') {
                targetPage = 1;
            } else if (text === '<') {
                targetPage = this._currentPage - 1;
            } else if (text === '>') {
                targetPage = this._currentPage + 1;
            } else if (text === '>>') {
                const totalPages = Math.ceil(this._totalRecords / this._pageSize);
                targetPage = totalPages;
            } else if (!Number.isNaN(Number.parseInt(text))) {
                targetPage = Number.parseInt(text);
            }
            
            // Use addListener for consistent cleanup tracking
            this.addListener(btn, 'click', (evt: Event) => {
                evt.preventDefault();
                evt.stopPropagation();
                this.goToPage(targetPage);
            });

            // Mark as having handler
            btn.dataset.hasOnclick = 'true';
        }
    }



    /**
     * Navigate to a specific page in the dataset
     */
    private goToPage(page: number): void {
        if (!this._enablePagination) {
            return;
        }

        const totalPages = Math.ceil(this._totalRecords / this._pageSize);
        if (page < 1 || page > totalPages || page === this._currentPage) {
            return;
        }

        this._currentPage = page;

        if (this._lastContext) {
            this.updateView(this._lastContext);
        }
    }

    // Ensure collapsed header background spans the scroll area width to avoid a narrow white gap
    private adjustCollapsedHeaderWidths(): void {
        try {
            const headers = Array.from(this._container.querySelectorAll('.pcf-group-section.pcf-group-collapsed > .pcf-group-header'));
            for (const h of headers) {
                try {
                    // CLEAN FIX: Only set minimum required properties to prevent expansion
                    (h as HTMLElement).style.setProperty('min-width', '0', 'important');
                    (h as HTMLElement).style.setProperty('width', '100%', 'important');
                    // Apply flexible width constraints

                    // Ensure the parent group section uses natural width
                    const parentSection = h.parentElement;
                    if (parentSection?.classList.contains('pcf-group-section')) {
                        parentSection.style.setProperty('width', 'auto', 'important');
                        parentSection.style.setProperty('min-width', '0', 'important');
                        // Apply responsive width styling
                    }
                } catch {
                    // Silent failure - don't spam console
                }
            }
        } catch {
            // Silent failure - this is called frequently
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {
            selectedRecordId: this._selectedRecordId || undefined,
            selectedRecordIds: this._selectedRecordIds ? this._selectedRecordIds.join(',') : '',
            rowEvent: this._lastRowEvent || ''
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // remove DOM and cleanup tracked listeners
        try {
            for (const d of this._listenerDisposables) {
                try {
                    d();
                } catch { /* ignore */ }
            }
        } catch { /* ignore */ }

        if (this._container?.parentNode) {
            this._container.remove();
        }
    }

    // Update only selection visuals (checkboxes and selected class) to avoid full re-render
    private updateSelectionVisuals(): void {
        try {
            const rows = this._container.querySelectorAll('.pcf-group-row');
            for (const rEl of Array.from(rows)) {
                const row = rEl as HTMLElement;
                const rid = row.dataset.recordId;
                if (!rid) continue;
                
                const selected = this._selectedRecordIds?.includes(rid);
                if (selected) {
                    row.classList.add('selected');
                } else {
                    row.classList.remove('selected');
                }
                
                const cb = row.querySelector('input[type="checkbox"]') as HTMLInputElement;
                if (cb) cb.checked = !!selected;
            }
        } catch { /* ignore */ }
    }

    // helper to check if a column is a lookup field based on dataType
    private isLookupColumn(column: any): boolean {
        if (!column?.dataType) return false;
        const dataType = column.dataType.toLowerCase();
        return dataType.includes('lookup') || dataType === 'customer';
    }

    // helper to get lookup EntityReference from a record
    private getLookupEntityReference(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, fieldName: string): any {
        if (!fieldName || !record) return null;

        try {
            const recAsAny = record as any;
            
            // Try getValue first - should return EntityReference for lookups
            if (typeof recAsAny.getValue === 'function') {
                const rawValue = recAsAny.getValue(fieldName);
                
                if (rawValue && typeof rawValue === 'object' && rawValue.id) {
                    const entityName = rawValue.etn || rawValue.logicalName || rawValue.entityType;
                    const entityId = rawValue.id?.guid ?? rawValue.id;
                    const displayName = rawValue.name || this.getRecordFieldValue(record, fieldName);
                    
                    if (entityName && entityId) {
                        return {
                            id: entityId,
                            logicalName: entityName,
                            name: displayName
                        };
                    }
                }
            }

            // Try getNamedReference as fallback per Microsoft PCF documentation
            if (typeof recAsAny.getNamedReference === 'function') {
                const namedRef = recAsAny.getNamedReference(fieldName);
                
                if (namedRef?.id) {
                    const entityName = namedRef.etn;
                    const entityId = namedRef.id?.guid ?? namedRef.id;
                    const displayName = namedRef.name || this.getRecordFieldValue(record, fieldName);
                    
                    if (entityName && entityId) {
                        return {
                            id: entityId,
                            logicalName: entityName,
                            name: displayName
                        };
                    }
                }
            }
        } catch (err) {
            // Silently handle lookup reference errors - this is expected when field types don't match
            // or when the API methods are not available in different PCF runtime versions
            this.handleError('getLookupEntityReference', err, 'low');
        }
        
        return null;
    }

    // helper to safely read field value
    private getRecordFieldValue(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord, fieldName: string): string | null {
        if (!fieldName) return null;

        let value: any = null;

        if (record) {
            const recAsAny = record as any;
            if (typeof recAsAny.getFormattedValue === 'function') {
                try {
                    value = recAsAny.getFormattedValue(fieldName);
                } catch {
                    void 0;
                }
            }

            if ((value === null || value === undefined) && typeof recAsAny.getValue === "function") {
                const gv = recAsAny.getValue;
                try {
                    value = gv(fieldName);
                } catch {
                    value = null;
                }
            }
        }

        // if getValue not available or returned undefined, try direct property access
        if ((value === null || value === undefined) && record) {
            const recObj = record as any;
            if (recObj[fieldName] !== undefined) {
                value = recObj[fieldName];
            }
        }

        if (value === null || value === undefined) return null;

        if (typeof value === "object") {
            // Check for common formatted value patterns in PCF
            if ("formatted" in value && typeof value.formatted === "string") return value.formatted;
            if ("name" in value && typeof value.name === "string") return value.name;
            if ("value" in value && typeof value.value === "string") return value.value;
            try {
                return JSON.stringify(value);
            } catch {
                return String(value);
            }
        }

        return String(value);
    }

    // helper to resolve record id from various dataset shapes
    private getRecordId(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord): string | null {
        if (!record) return null;

        const rObj = record as any;
        const maybeGet = rObj.getRecordId;
        if (typeof maybeGet === 'function') {
            try {
                return maybeGet();
            } catch { /* fallthrough */ }
        }

        if (typeof rObj.recordId === 'string') return rObj.recordId;
        if (typeof rObj.id === 'string') return rObj.id;
        if (typeof rObj['$id'] === 'string') return rObj['$id'];

        return null;
    }

    // helper to resolve entity logical name/type from record object
    private getRecordEntityName(record: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord): string | null {
        if (!record) return null;

        const rObj = record as any;

        // try raw identifier shape first (some dataset records expose an internal _record.identifier)
        try {
            const raw = rObj['_record'];
            if (raw && typeof raw === 'object') {
                const ident = raw['identifier'];
                if (ident && typeof ident === 'object') {
                    const entityName = ident['etn'] || ident['entityType'] || ident['entityLogicalName'];
                    if (entityName && typeof entityName === 'string') return entityName;
                }
            }
        } catch { /* ignore */ }

        // common shapes
        if (typeof rObj['entityType'] === 'string') return rObj['entityType'];
        if (typeof rObj['entityLogicalName'] === 'string') return rObj['entityLogicalName'];

        // try named reference shape
        const maybeGetNamedRef = rObj.getNamedReference;
        if (typeof maybeGetNamedRef === 'function') {
            try {
                const ref = maybeGetNamedRef();
                if (ref && typeof ref === 'object') {
                    if (typeof ref['entityType'] === 'string') return ref['entityType'];
                    if (typeof ref['entityLogicalName'] === 'string') return ref['entityLogicalName'];
                }
            } catch { /* ignore */ }
        }

        return null;
    }

    // helper to sync selection with host dataset safely
    private _syncSelectionToHost(context: ComponentFramework.Context<IInputs>, ids: string[]): void {
        try {
            if (!ids) ids = [];
            const ds = context.parameters?.['sampleDataSet'];
            if (ds) {
                const maybeDs = ds as any;
                if (typeof maybeDs.setSelectedRecordIds === 'function') {
                    try {
                        const fn = maybeDs.setSelectedRecordIds;
                        const datasetKeys = (ids || []).filter(Boolean);
                        try {
                            fn(datasetKeys);
                        } catch { /* ignore */ }
                    } catch { /* ignore */ }
                }
            }
        } catch { /* ignore */ }
    }

    // helper to navigate/open a record using dataset named reference or record helpers
    private navigateToRecord(context: ComponentFramework.Context<IInputs>, datasetRowId: string): void {
        if (!datasetRowId) return;

        try {
            // navigate request
            const ds = context.parameters?.['sampleDataSet'];
            if (ds?.records?.[datasetRowId]) {
                const rec = ds.records[datasetRowId];

                try {
                    // Prefer raw identifier shape if present (fast path for dataset records)
                    try {
                        const raw = (rec as any)['_record'];
                        if (raw && typeof raw === 'object') {
                            const ident = raw['identifier'];
                            if (ident && typeof ident === 'object') {
                                const _entityName = ident['etn'] || ident['entityType'] || ident['entityLogicalName'];
                                const idObj = ident['id'];
                                const id = idObj && (idObj['guid'] || idObj['Id'] || idObj['value']) ? 
                                    String(idObj['guid'] || idObj['Id'] || idObj['value']) : undefined;

                                if (_entityName && id && 
                                    typeof context.navigation?.openForm === 'function') {
                                    try {
                                        // If this is the generic activitypointer, try to map to the actual activity entity using activitytypecode
                                        let finalEntity = _entityName;
                                        try {
                                            if (_entityName === 'activitypointer' && 
                                                typeof raw?.['fields'] === 'object') {
                                                const fields = raw['fields'];
                                                const activityType = fields['activitytypecode'];
                                                const atVal = activityType ? 
                                                    (activityType['value'] ?? activityType['formatted'] ?? activityType['label']) : undefined;
                                                const at = String(atVal ?? '').toLowerCase();

                                                const map: { [key: string]: string } = {
                                                    'phonecall': 'phonecall',
                                                    'task': 'task',
                                                    'email': 'email',
                                                    'appointment': 'appointment',
                                                    'letter': 'letter',
                                                    'fax': 'fax',
                                                    'serviceappointment': 'serviceappointment'
                                                };

                                                // activityType sometimes contains 'phonecall' or formatted 'Phone Call'
                                                for (const k of Object.keys(map)) {
                                                    if (at.includes(k)) {
                                                        finalEntity = map[k];
                                                    }
                                                }
                                            }
                                        } catch { /* activity mapping failed */ }

                                        context.navigation.openForm({
                                            entityName: finalEntity,
                                            entityId: id
                                        });
                                        return;
                                    } catch { /* ignore */ }
                                }
                            }
                        }
                    } catch { /* raw identifier check failed */ }

                    // try named reference if raw not available
                    const maybeNamed = (rec as any).getNamedReference;
                    const namedRef = typeof maybeNamed === 'function' ? maybeNamed() : null;
                    if (namedRef && typeof namedRef === 'object') {
                        // namedRef available
                        const _entityName2 = namedRef['entityName'] || namedRef['entityLogicalName'] || namedRef['entityType'];
                        const _id = namedRef['id'] || namedRef['entityId'] || namedRef['key'];

                        if (_entityName2 && _id && 
                            typeof context.navigation?.openForm === 'function') {
                            try {
                                context.navigation.openForm({
                                    entityName: _entityName2,
                                    entityId: _id
                                }).catch(() => { /* ignore */ });
                                return;
                            } catch { /* ignore */ }
                        }
                    }
                } catch { /* ignore */ }

                // Fallback: use record helper methods
                const rid = this.getRecordId(rec) || datasetRowId;
                const entityName = this.getRecordEntityName(rec) || null;

                if (rid && entityName && 
                    typeof context.navigation?.openForm === 'function') {
                    try {
                        context.navigation.openForm({
                            entityName,
                            entityId: rid
                        }).catch(() => { /* ignore */ });
                        return;
                    } catch { /* ignore */ }
                }

                // emit event fallback
                this._lastRowEvent = JSON.stringify({
                    type: 'openRequest',
                    id: rid || datasetRowId,
                    entityName: entityName
                });
                if (this._notifyOutputChanged) this._notifyOutputChanged();
            }
        } catch { /* ignore */ }
    }
}
