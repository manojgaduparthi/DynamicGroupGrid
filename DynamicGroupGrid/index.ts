import { IInputs, IOutputs } from "./generated/ManifestTypes";
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;

export class DynamicGroupGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    // Configurable constants and CSS class names
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
    };

    // Configurable values (magic numbers centralized)
    private static readonly DEFAULT_UNIFORM_WIDTH = 150;
    private static readonly DEFAULT_MIN_WIDTH = 40;
    private static readonly DEFAULT_FALLBACK_FLEX = '1 1 120px';
    private static readonly DEFAULT_PAGE_SIZE = 25;
    private static readonly MAX_PAGE_SIZE = 100;

    // runtime state for optimized updates
    private _listenerDisposables: (() => void)[] = [];
    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

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
            if (context && context.mode && typeof (context.mode as unknown as { trackContainerResize?: unknown }).trackContainerResize === 'function') {
                try { (context.mode as unknown as { trackContainerResize: (v: boolean) => void }).trackContainerResize(true); } catch { /* ignore */ }
            }
        } catch { /* ignore */ }
        this._selectedGroupColumn = null;
        this._expandedGroups = new Map<string, boolean>();
        this._selectedRecordId = null;
        this._selectedRecordIds = [];
        this._notifyOutputChanged = notifyOutputChanged;
        this._currentPage = 1;
        this._pageSize = DynamicGroupGrid.DEFAULT_PAGE_SIZE;
        this._enablePagination = true;
        this._lastContext = null;
        // track a single window resize listener to keep collapsed header widths in sync
        try {
            this.addListener(window, 'resize', () => {
                try { this.adjustCollapsedHeaderWidths(); } catch { /* ignore */ }
            });
        } catch { /* ignore */ }
    }

    // Helper to register and track DOM event listeners for cleanup
    private addListener(el: EventTarget, evt: string, fn: EventListenerOrEventListenerObject, options?: boolean | AddEventListenerOptions) {
        // Keep a typed reference for DOM add/remove; caller can pass specific mouse/keyboard handlers
        const listener = fn as EventListener;
        el.addEventListener(evt, listener, options);
        this._listenerDisposables.push(() => { try { el.removeEventListener(evt, listener, options); } catch { /* ignore */ } });
    }

    // Apply a consistent column width policy to both header and row cells
    private applyColumnWidth(el: HTMLElement, columnName: string | null) {
        try {
            const col = columnName || '';
            const storedW = this._colWidths.get(col);
            const uniform = this._colUniformWidth || DynamicGroupGrid.DEFAULT_UNIFORM_WIDTH;
            if (uniform && uniform > 0) {
                el.style.flex = `0 0 ${uniform}px`;
            } else if (storedW && storedW > 0) {
                el.style.flex = `0 0 ${storedW}px`;
            } else {
                el.style.flex = DynamicGroupGrid.DEFAULT_FALLBACK_FLEX;
                el.style.maxWidth = '400px';
                el.style.minWidth = `${DynamicGroupGrid.DEFAULT_MIN_WIDTH}px`;
            }
        } catch (e) { this.handleError('applyColumnWidth', e); }
    }

    // Centralized error handling
    private handleError(operation: string, error: unknown): void {
        // Log for debugging but don't expose to user
        if (console?.error) {
            console.error(`DynamicGroupGrid ${operation}:`, error);
        }
        // Could implement user-friendly error reporting here in the future
    }

    // Debug helpers removed for production build.


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this._lastContext = context;
        const dataset = context.parameters.sampleDataSet;
        // Use dataset if available; early return if no data
        if (!dataset || !dataset.records || Object.keys(dataset.records).length === 0) {
            return;
        }
        const ds = dataset;

        // Read pagination configuration from context
        try {
            if (context.parameters.pageSize && context.parameters.pageSize.raw !== null && context.parameters.pageSize.raw !== undefined) {
                const configuredPageSize = Number(context.parameters.pageSize.raw);
                if (configuredPageSize > 0 && configuredPageSize <= DynamicGroupGrid.MAX_PAGE_SIZE) {
                    this._pageSize = configuredPageSize;
                }
            }
            if (context.parameters.enablePagination && context.parameters.enablePagination.raw !== null && context.parameters.enablePagination.raw !== undefined) {
                this._enablePagination = Boolean(context.parameters.enablePagination.raw);
            }
        } catch { /* ignore */ }

        // read columns
        const columns = ds.columns ? (Object.values(ds.columns) as DataSetInterfaces.Column[]) : [];

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
                const parsed = JSON.parse(raw) as Record<string, number>;
                Object.keys(parsed).forEach(k => this._colWidths.set(k, parsed[k]));
            }
            // load uniform width preference
            try { const uw = localStorage.getItem('dynamic_group_grid_uniform_width'); if (uw) this._colUniformWidth = Number(uw) || this._colUniformWidth; } catch { /* ignore */ }
        } catch (e) { void e; }

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
        columns.forEach(col => {
            const opt = document.createElement("option");
            opt.value = col.name;
            opt.text = (col.displayName as unknown as string) || col.name;
            if (col.name === this._selectedGroupColumn) opt.selected = true;
            select.appendChild(opt);
        });
        this.addListener(select, "change", (evt: Event) => {
            const val = (evt.target as HTMLSelectElement).value;
            this._selectedGroupColumn = val;
            this._currentPage = 1; // Reset to first page when changing grouping
            this.updateView(context);
        });
        toolbar.appendChild(select);
    // Note: horizontal panning is available via native horizontal scrollbar; removed toolbar range control.
        this._container.appendChild(toolbar);

        // Debug toolbar removed for production build.

        // Removed bulk action toolbar per request — native subgrid command bar will be used instead.

        // gather records with their ids so we always use dataset ids (avoids unreliable getRecordId shapes)
        const allEntries = (ds.sortedRecordIds ? ds.sortedRecordIds.map(id => ({ id: String(id), rec: (ds.records as Record<string, unknown>)[id] as DataSetInterfaces.EntityRecord })) : Object.keys(ds.records || {}).map(k => ({ id: String(k), rec: (ds.records as Record<string, unknown>)[k] as DataSetInterfaces.EntityRecord })));
        
        this._totalRecords = allEntries.length;
        
        // Apply pagination if enabled
        const entries = this._enablePagination ? this.getPaginatedEntries(allEntries) : allEntries;
        
        const allRecordIds = entries.map((e: { id: string; rec: DataSetInterfaces.EntityRecord }) => e.id);
        const groups: Record<string, { id: string; rec: DataSetInterfaces.EntityRecord }[]> = {};
        entries.forEach((entry: { id: string; rec: DataSetInterfaces.EntityRecord }) => {
            const val = this.getRecordFieldValue(entry.rec, this._selectedGroupColumn) || "(blank)";
            if (!groups[val]) groups[val] = [];
            groups[val].push(entry);
        });

        // render a single header that shows columns
        const headerRow = document.createElement("div");
        headerRow.className = "pcf-grid-header";
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
            this._selectedRecordId = this._selectedRecordIds.length ? this._selectedRecordIds[this._selectedRecordIds.length - 1] : null;
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
        colsToShow.forEach(c => {
            const hc = document.createElement('div');
            hc.className = 'pcf-grid-col';
            // apply stored/uniform width to header cell to match row cells
            try { this.applyColumnWidth(hc, (c.name || '').toLowerCase()); } catch { /* ignore */ }
            hc.textContent = (c.displayName as unknown as string) || c.name;
            hc.title = `Sort by ${(c.displayName as unknown as string) || c.name}`;
            hc.setAttribute('data-col-name', (c.name || '').toLowerCase());
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
                const sEv = startEv as MouseEvent;
                sEv.preventDefault();
                sEv.stopPropagation();
                isDragging = true;
                const startX = sEv.clientX;
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
                        rows.forEach(rEl => {
                            const child = (rEl as HTMLElement).children[headerIndex] as HTMLElement | undefined;
                            if (child) this.applyColumnWidth(child as HTMLElement, c.name.toLowerCase());
                        });
                    } catch (e) { void e; }
                };
                const onUp = () => {
                    isDragging = false;
                    try { document.removeEventListener('mousemove', onMove); } catch { /* ignore */ }
                    try { document.removeEventListener('mouseup', onUp); } catch { /* ignore */ }
                    // persist widths to localStorage
                    try {
                        const key = 'pcf_col_widths_DynamicGroupGrid';
                        const obj: Record<string, number> = {};
                        this._colWidths.forEach((v,k)=> obj[k]=v);
                        localStorage.setItem(key, JSON.stringify(obj));
                    } catch (e) { void e; }
                    // final updateSelectionVisuals to ensure layout applied
                    this.updateSelectionVisuals();
                };
                this.addListener(document, 'mousemove', onMove);
                this.addListener(document, 'mouseup', onUp);
            });
            hc.appendChild(handle);
            headerRow.appendChild(hc);
        });
    // create a horizontally scrollable wrapper for header + groups so all fields can be seen
    const gridWrapper = document.createElement('div');
    gridWrapper.className = 'pcf-grid-scroll-wrap';
    gridWrapper.style.overflowX = 'auto';
    gridWrapper.style.display = 'block';
    gridWrapper.appendChild(headerRow);
    this._container.appendChild(gridWrapper);

        // determine ordered list of group keys so group sections render in a predictable sort order
        const groupKeys = Object.keys(groups || {});
        try {
            // sort case-insensitive and numeric-aware
            groupKeys.sort((a, b) => {
                const aa = a || '';
                const bb = b || '';
                const cmp = aa.localeCompare(bb, undefined, { sensitivity: 'base', numeric: true });
                // if the current sort column is the same as the group-by column, respect the sort direction
                if (this._sortColumn === this._selectedGroupColumn) return this._sortAsc ? cmp : -cmp;
                // otherwise default to ascending by group key
                return cmp;
            });
        } catch { /* ignore sort errors and fall back to insertion order */ }

        groupKeys.forEach(groupKey => {
            const groupDiv = document.createElement("div");
            groupDiv.className = "pcf-group-section";

            const header = document.createElement("div");
            header.className = "pcf-group-header";
            const chev = document.createElement("span");
            chev.className = "chev";
            // Triangle arrow is now fully handled by CSS for consistent cross-environment rendering
            // default groups to collapsed; user can click to expand
            if (!this._expandedGroups.has(groupKey)) this._expandedGroups.set(groupKey, false);
            if (!this._expandedGroups.get(groupKey)) chev.classList.add("collapsed");
            header.appendChild(chev);

            const title = document.createElement("div");
            title.textContent = `${groupKey} (${groups[groupKey].length})`;
            header.appendChild(title);
            this.addListener(header, 'click', () => {
                const cur = this._expandedGroups.get(groupKey) || false;
                this._expandedGroups.set(groupKey, !cur);
                this.updateView(context);
            });

            groupDiv.appendChild(header);

            const list = document.createElement("div");
            list.className = DynamicGroupGrid.CSS.GROUP_LIST;
            const expanded = this._expandedGroups.get(groupKey) !== false;
            // hide the list entirely when collapsed to avoid rendering an empty placeholder row
            if (!expanded) {
                (list as HTMLElement).style.display = 'none';
                // mark parent as collapsed for CSS targeting
                groupDiv.classList.add('collapsed');
            }
            if (expanded) {
                // optionally sort group rows
                const grpRows = groups[groupKey].slice();
                if (this._sortColumn) {
                    const sc = this._sortColumn;
                    grpRows.sort((a,b)=>{
                        const va = this.getRecordFieldValue(a.rec, sc) || '';
                        const vb = this.getRecordFieldValue(b.rec, sc) || '';
                        if (va < vb) return this._sortAsc ? -1 : 1;
                        if (va > vb) return this._sortAsc ? 1 : -1;
                        return 0;
                    });
                }
                grpRows.forEach(entry => {
                    const row = document.createElement("div");
                    row.className = "pcf-group-row";
                    const recordId = entry.id;
                    const r = entry.rec;
                    if (entry.id) row.setAttribute('data-record-id', entry.id);
                    if (recordId && this._selectedRecordIds && this._selectedRecordIds.indexOf(recordId) !== -1) row.classList.add("selected");
                    // add selection checkbox cell
                    const selectCell = document.createElement('div');
                    selectCell.className = 'pcf-grid-col';
                    selectCell.setAttribute('data-col-name', 'select');
                    const cb = document.createElement('input');
                    cb.type = 'checkbox';
                    cb.checked = this._selectedRecordIds && this._selectedRecordIds.indexOf(recordId!) !== -1;
                    // prevent clicks on the checkbox from bubbling to the row click handler
                    this.addListener(cb, 'click', (evt: Event) => { evt.stopPropagation(); });
                    this.addListener(cb, 'change', (evt: Event) => {
                        evt.stopPropagation();
                        if (!this._selectedRecordIds) this._selectedRecordIds = [];
                        const checked = (evt.target as HTMLInputElement).checked;
                        const idx = this._selectedRecordIds.indexOf(recordId!);
                        if (checked && idx === -1) this._selectedRecordIds.push(recordId!);
                        if (!checked && idx !== -1) this._selectedRecordIds.splice(idx,1);
                        this._selectedRecordId = this._selectedRecordIds.length ? this._selectedRecordIds[this._selectedRecordIds.length-1] : null;
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
                    colsToShow.forEach((c, idx) => {
                        const cellDiv = document.createElement('div');
                        cellDiv.className = 'pcf-grid-col';
                        // expose column name as attribute for targeted CSS
                        cellDiv.setAttribute('data-col-name', (c.name || '').toLowerCase());
                        // apply uniform/stored width via helper
                        try { this.applyColumnWidth(cellDiv, (c.name || '').toLowerCase()); } catch { /* ignore */ }

                        const inner = document.createElement('span');
                        inner.className = 'pcf-cell';
                        const text = this.getRecordFieldValue(r, c.name) || '';
                        if (idx === 0) {
                            const a = document.createElement('a');
                            a.href = '#';
                            a.textContent = text;
                            this.addListener(a, 'click', (evt: Event) => { evt.preventDefault(); evt.stopPropagation(); try { this.navigateToRecord(context, recordId); } catch { this._lastRowEvent = JSON.stringify({ type: 'openRequest', id: recordId }); if (this._notifyOutputChanged) this._notifyOutputChanged(); } });
                            inner.appendChild(a);
                        } else {
                            inner.textContent = text;
                        }
                        cellDiv.appendChild(inner);
                        // also set data attribute on inner for CSS selectors
                        inner.setAttribute('data-col-name', (c.name || '').toLowerCase());
                        row.appendChild(cellDiv);
                    });
                    this.addListener(row, 'click', (evt: Event) => {
                        // stop group toggle
                        evt.stopPropagation();
                        // set multi-selection (Ctrl/Cmd toggle, Shift range)
                        if (!this._selectedRecordIds) this._selectedRecordIds = [];
                        if (recordId) {
                            const isCtrl = (evt as MouseEvent).ctrlKey || (evt as MouseEvent).metaKey;
                            const isShift = (evt as MouseEvent).shiftKey;
                            if (isShift && this._selectedRecordIds.length > 0) {
                                // select range from last selected to this
                                const last = this._selectedRecordIds[this._selectedRecordIds.length - 1];
                                const start = allRecordIds.indexOf(last);
                                const end = allRecordIds.indexOf(recordId);
                                if (start >= 0 && end >= 0) {
                                    const [s,e] = start <= end ? [start,end] : [end,start];
                                    const range = allRecordIds.slice(s, e+1);
                                    this._selectedRecordIds = Array.from(new Set(this._selectedRecordIds.concat(range)));
                                }
                            }
                            else if (isCtrl) {
                                // toggle
                                const idx = this._selectedRecordIds.indexOf(recordId);
                                if (idx === -1) this._selectedRecordIds.push(recordId);
                                else this._selectedRecordIds.splice(idx,1);
                            }
                            else {
                                // single select
                                this._selectedRecordIds = [recordId];
                            }
                            // update primary single output for backward compatibility
                            this._selectedRecordId = this._selectedRecordIds.length ? this._selectedRecordIds[this._selectedRecordIds.length-1] : null;
                            if (this._notifyOutputChanged) this._notifyOutputChanged();
                            // sync selection with host dataset
                            this._syncSelectionToHost(context, this._selectedRecordIds || []);
                            // avoid full re-render; update selection visuals
                            this.updateSelectionVisuals();
                        }
                    });
                    this.addListener(row, 'dblclick', (evt: Event) => {
                        evt.stopPropagation();
                        if (recordId) {
                            try { this.navigateToRecord(context, recordId); } catch {
                                this._lastRowEvent = JSON.stringify({ type: 'doubleClick', id: recordId });
                                if (this._notifyOutputChanged) this._notifyOutputChanged();
                            }
                        }
                    });
                    list.appendChild(row);
                });
            }
            groupDiv.appendChild(list);
            gridWrapper.appendChild(groupDiv);
        });
        // horizontal panning is available via native scrollbar; no toolbar slider wiring necessary

        // Add pagination controls if enabled
        if (this._enablePagination && this._totalRecords > this._pageSize) {
            this.renderPaginationControls();
        }

        // final sizing adjustment: ensure collapsed headers visually span the scrollable content
        try { this.adjustCollapsedHeaderWidths(); } catch { /* ignore */ }
    }

    // Get paginated subset of entries
    private getPaginatedEntries(allEntries: { id: string; rec: DataSetInterfaces.EntityRecord }[]): { id: string; rec: DataSetInterfaces.EntityRecord }[] {
        if (!this._enablePagination || this._pageSize <= 0) {
            return allEntries;
        }
        
        const startIndex = (this._currentPage - 1) * this._pageSize;
        const endIndex = startIndex + this._pageSize;
        return allEntries.slice(startIndex, endIndex);
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
        const startRecord = ((this._currentPage - 1) * this._pageSize) + 1;
        const endRecord = Math.min(this._currentPage * this._pageSize, this._totalRecords);
        paginationInfo.textContent = `Showing ${startRecord}-${endRecord} of ${this._totalRecords} records`;
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
        
        if (!disabled) {
            this.addListener(button, 'click', (evt: Event) => {
                evt.preventDefault();
                evt.stopPropagation();
                this.goToPage(targetPage);
            });
        }
        
        return button;
    }

    // Navigate to a specific page
    private goToPage(page: number): void {
        if (!this._enablePagination) {
            return;
        }
        
        const totalPages = Math.ceil(this._totalRecords / this._pageSize);
        if (page < 1 || page > totalPages || page === this._currentPage) {
            return;
        }
        
        this._currentPage = page;
        // Find the context from the most recent updateView call
        if (this._lastContext) {
            this.updateView(this._lastContext);
        }
    }

    // Ensure collapsed header background spans the scroll area width to avoid a narrow white gap
    private adjustCollapsedHeaderWidths(): void {
        try {
            const gridWrap = this._container.querySelector('.pcf-grid-scroll-wrap') as HTMLElement | null;
            if (!gridWrap) return;
            // measure the scrollable content width (including any horizontal overflow)
            const contentWidth = gridWrap.scrollWidth || gridWrap.clientWidth || 0;
            const headers = Array.from(this._container.querySelectorAll('.pcf-group-section.collapsed > .pcf-group-header')) as HTMLElement[];
            headers.forEach(h => {
                // ensure header background covers the full scroll content area by setting a pseudo-full width filler via minWidth
                // use inline style so we don't require additional CSS selectors
                try {
                    h.style.minWidth = `${contentWidth}px`;
                } catch { /* ignore */ }
            });
        } catch { /* ignore */ }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return { selectedRecordId: this._selectedRecordId, selectedRecordIds: this._selectedRecordIds ? this._selectedRecordIds.join(',') : '', rowEvent: this._lastRowEvent || '' } as unknown as IOutputs;
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // remove DOM and cleanup tracked listeners
        try {
            this._listenerDisposables.forEach(d => { try { d(); } catch { /* ignore */ } });
        } catch { /* ignore */ }
        if (this._container && this._container.parentNode) {
            this._container.parentNode.removeChild(this._container);
        }
    }

    // Update only selection visuals (checkboxes and selected class) to avoid full re-render
    private updateSelectionVisuals(): void {
        try {
            const rows = this._container.querySelectorAll('.pcf-group-row');
            rows.forEach(rEl => {
                const row = rEl as HTMLElement;
                const rid = row.getAttribute('data-record-id');
                if (!rid) return;
                const selected = this._selectedRecordIds && this._selectedRecordIds.indexOf(rid) !== -1;
                if (selected) row.classList.add('selected'); else row.classList.remove('selected');
                const cb = row.querySelector('input[type="checkbox"]') as HTMLInputElement | null;
                if (cb) cb.checked = !!selected;
            });
        } catch { /* ignore */ }
    }

    // helper to safely read field value
    private getRecordFieldValue(record: DataSetInterfaces.EntityRecord, fieldName: string | null): string | null {
        if (!fieldName) return null;
        let value: unknown = null;
        if (record) {
            const recAsAny = record as unknown as { getFormattedValue?: unknown; getValue?: unknown };
            if (typeof recAsAny.getFormattedValue === 'function') {
                try { value = (recAsAny.getFormattedValue as (f: string) => unknown)(fieldName as string); } catch { void 0; }
            }
            if ((value === null || value === undefined) && typeof recAsAny.getValue === "function") {
                const gv = recAsAny.getValue as (f: string) => unknown;
                try {
                    value = gv(fieldName as string);
                }
                catch {
                    value = null;
                }
            }
        }
        // if getValue not available or returned undefined, try direct property access
        if ((value === null || value === undefined) && record) {
            const recObj = record as unknown as Record<string, unknown>;
            if (recObj[fieldName as string] !== undefined) {
                value = recObj[fieldName as string];
            }
        }

        if (value === null || value === undefined) return null;

        if (typeof value === "object") {
            const vObj = value as Record<string, unknown>;
            if ("formatted" in vObj && typeof vObj["formatted"] === "string") return vObj["formatted"] as string;
            if ("name" in vObj && typeof vObj["name"] === "string") return vObj["name"] as string;
            if ("value" in vObj && typeof vObj["value"] === "string") return vObj["value"] as string;
            try {
                return JSON.stringify(vObj);
            }
            catch {
                return String(value);
            }
        }

        return String(value);
    }

    private _container: HTMLDivElement;
    private _selectedGroupColumn: string | null;
    private _expandedGroups: Map<string, boolean>;
    private _selectedRecordId: string | null;
    private _notifyOutputChanged: (() => void) | null;
    private _selectedRecordIds: string[];
    private _lastRowEvent: string | null;
    private _sortColumn: string | null = null;
    private _sortAsc = true;
    private _colWidths: Map<string, number> = new Map<string, number>();
    private _colUniformWidth = 150;
    private _currentPage = 1;
    private _pageSize = DynamicGroupGrid.DEFAULT_PAGE_SIZE;
    private _totalRecords = 0;
    private _enablePagination = true;
    private _lastContext: ComponentFramework.Context<IInputs> | null = null;

    // helper to resolve record id from various dataset shapes
    private getRecordId(record: DataSetInterfaces.EntityRecord): string | null {
        if (!record) return null;
        const rObj = record as unknown as Record<string, unknown>;
        const maybeGet = (rObj as unknown as { getRecordId?: unknown }).getRecordId;
        if (typeof maybeGet === 'function') {
            try { return (maybeGet as () => unknown)() as string; } catch { /* fallthrough */ }
        }
        if (typeof rObj.recordId === 'string') return rObj.recordId as string;
        if (typeof rObj.id === 'string') return rObj.id as string;
        if (typeof rObj['$id'] === 'string') return rObj['$id'] as string;
        return null;
    }

    // helper to resolve entity logical name/type from record object
    private getRecordEntityName(record: DataSetInterfaces.EntityRecord): string | null {
        if (!record) return null;
        const rObj = record as unknown as Record<string, unknown>;
        // try raw identifier shape first (some dataset records expose an internal _record.identifier)
        try {
            const raw = rObj['_record'] as Record<string, unknown> | undefined;
            if (raw && typeof raw === 'object') {
                const ident = raw['identifier'] as Record<string, unknown> | undefined;
                if (ident && typeof ident === 'object') {
                    const entityName = (ident['etn'] || ident['entityType'] || ident['entityLogicalName']) as string | undefined;
                    if (entityName && typeof entityName === 'string') return entityName;
                }
            }
        } catch { /* ignore */ }
        // common shapes
            if (typeof rObj['entityType'] === 'string') return rObj['entityType'] as string | null;
            if (typeof rObj['entityLogicalName'] === 'string') return rObj['entityLogicalName'] as string | null;
        // try named reference shape
        const maybeGetNamedRef = (rObj as unknown as { getNamedReference?: unknown }).getNamedReference;
        if (typeof maybeGetNamedRef === 'function') {
            try {
                const ref = (maybeGetNamedRef as () => unknown)();
                if (ref && typeof ref === 'object') {
                    const rr = ref as Record<string, unknown>;
                    if (typeof rr['entityType'] === 'string') return rr['entityType'] as string;
                    if (typeof rr['entityLogicalName'] === 'string') return rr['entityLogicalName'] as string;
                }
            } catch { /* ignore */ }
        }
        return null;
    }

    // helper to sync selection with host dataset safely
    private _syncSelectionToHost(context: ComponentFramework.Context<IInputs>, ids: string[] | null | undefined): void {
        try {
            if (!ids) ids = [];
            const ds = (context.parameters && context.parameters['sampleDataSet']) as unknown as ComponentFramework.PropertyTypes.DataSet | undefined;
            if (ds) {
                const maybeDs = ds as unknown as { setSelectedRecordIds?: unknown };
                if (typeof maybeDs.setSelectedRecordIds === 'function') {
                    try {
                        const fn = (maybeDs.setSelectedRecordIds as unknown) as ((ids: unknown) => void);
                        const datasetKeys = (ids || []).filter(Boolean);
                        try { fn(datasetKeys); } catch { /* ignore */ }
                    } catch { /* ignore */ }
                }
            }
        } catch { /* ignore */ }
    }

    // ...existing code...

    // helper to navigate/open a record using dataset named reference or record helpers
    private navigateToRecord(context: ComponentFramework.Context<IInputs>, datasetRowId: string | null | undefined): void {
        if (!datasetRowId) return;
        try {
            // navigate request
            const ds = (context.parameters && context.parameters['sampleDataSet']) as unknown as ComponentFramework.PropertyTypes.DataSet | undefined;
            if (ds && ds.records && ds.records[datasetRowId]) {
                const rec = ds.records[datasetRowId] as unknown as DataSetInterfaces.EntityRecord;
                try {
                    // Prefer raw identifier shape if present (fast path for dataset records)
                    try {
                        const raw = (rec as unknown as Record<string, unknown>)['_record'] as Record<string, unknown> | undefined;
                        if (raw && typeof raw === 'object') {
                            const ident = raw['identifier'] as Record<string, unknown> | undefined;
                            if (ident && typeof ident === 'object') {
                                const entityName = (ident['etn'] || ident['entityType'] || ident['entityLogicalName']) as string | undefined;
                                const idObj = ident['id'] as Record<string, unknown> | undefined;
                                const id = idObj && (idObj['guid'] || idObj['Id'] || idObj['value']) ? String(idObj['guid'] || idObj['Id'] || idObj['value']) : undefined;
                                if (entityName && id && context && context.navigation && typeof context.navigation.openForm === 'function') {
                                    try {
                                        // If this is the generic activitypointer, try to map to the actual activity entity using activitytypecode
                                        let finalEntity = entityName;
                                        try {
                                            if (entityName === 'activitypointer' && raw && raw['fields'] && typeof raw['fields'] === 'object') {
                                                const fields = raw['fields'] as Record<string, unknown>;
                                                const activityType = fields['activitytypecode'] as Record<string, unknown> | undefined;
                                                const atVal = activityType ? (activityType['value'] ?? activityType['formatted'] ?? activityType['label']) : undefined;
                                                const at = String(atVal ?? '').toLowerCase();
                                                const map: Record<string,string> = { 'phonecall':'phonecall', 'task':'task', 'email':'email', 'appointment':'appointment', 'letter':'letter', 'fax':'fax', 'serviceappointment':'serviceappointment' };
                                                // activityType sometimes contains 'phonecall' or formatted 'Phone Call'
                                                Object.keys(map).forEach(k=>{ if (at.indexOf(k) !== -1) finalEntity = map[k]; });
                                            }
                                        } catch { /* activity mapping failed */ }
                                        (context.navigation.openForm as unknown as (opts:{entityName?:string,entityId?:string})=>unknown)({ entityName: finalEntity, entityId: id });
                                        return;
                                    } catch { /* ignore */ }
                                }
                            }
                        }
                    } catch { /* raw identifier check failed */ }
                    // try named reference if raw not available
                    const maybeNamed = (rec as unknown as { getNamedReference?: unknown }).getNamedReference;
                    const namedRef = typeof maybeNamed === 'function' ? (maybeNamed as () => unknown)() : null;
                    if (namedRef && typeof namedRef === 'object') {
                        // namedRef available
                        const nr = namedRef as Record<string, unknown>;
                        const entityName = (nr['entityName'] || nr['entityLogicalName'] || nr['entityType']) as string | undefined;
                        const id = (nr['id'] || nr['entityId'] || nr['key']) as string | undefined;
                        if (entityName && id && context && context.navigation && typeof context.navigation.openForm === 'function') {
                            try { (context.navigation.openForm as unknown as (opts:{entityName?:string,entityId?:string})=>unknown)({ entityName, entityId: id }); return; } catch { /* ignore */ }
                        }
                    }
                } catch { /* ignore */ }
                // fallback: try record helpers
                const rid = this.getRecordId(rec) || datasetRowId;
                const entityName = this.getRecordEntityName(rec) || null;
                if (rid && entityName && context && context.navigation && typeof context.navigation.openForm === 'function') {
                    try { (context.navigation.openForm as unknown as (opts:{entityName?:string,entityId?:string})=>unknown)({ entityName, entityId: rid }); return; } catch { /* ignore */ }
                }
                // emit event fallback
                this._lastRowEvent = JSON.stringify({ type: 'openRequest', id: rid || datasetRowId, entityName: entityName });
                if (this._notifyOutputChanged) this._notifyOutputChanged();
            }
        } catch { /* ignore */ }
    }

    // resolve dataset entry ids to platform canonical ids (GUIDs) when possible
    private _toCanonicalIds(context: ComponentFramework.Context<IInputs>, ids: string[] | null | undefined): string[] {
        const out: string[] = [];
        try {
            if (!ids) return out;
            const params = context.parameters as unknown as Record<string, unknown>;
            const dsRaw = params['sampleDataSet'] as unknown;
            if (!dsRaw || typeof dsRaw !== 'object') return ids.filter(Boolean);
            const ds = dsRaw as unknown as ComponentFramework.PropertyTypes.DataSet;
            ids.forEach(id => {
                try {
                    const rec = (ds.records as Record<string, unknown>)[id] as unknown as DataSetInterfaces.EntityRecord | undefined;
                    if (rec) {
                        // 1) try getRecordId()
                        const rid = this.getRecordId(rec);
                        if (rid) { out.push(rid); return; }
                        // 2) try named reference shape
                        try {
                            const maybeNamed = (rec as unknown as { getNamedReference?: unknown }).getNamedReference;
                            const namedRef = typeof maybeNamed === 'function' ? (maybeNamed as () => unknown)() : null;
                            if (namedRef && typeof namedRef === 'object') {
                                const nr = namedRef as Record<string, unknown>;
                                const idFld = (nr['id'] || nr['entityId'] || nr['key']);
                                if (idFld && typeof idFld === 'object') {
                                    const idObj = idFld as Record<string, unknown>;
                                    const gid = (idObj['guid'] || idObj['Id'] || idObj['value']);
                                    if (gid) { out.push(String(gid)); return; }
                                }
                                if (typeof idFld === 'string' && idFld) { out.push(idFld); return; }
                            }
                        } catch { /* ignore namedRef */ }
                        // 3) try internal raw identifier (common in some host shapes)
                        try {
                            const raw = (rec as unknown as Record<string, unknown>)['_record'] as Record<string, unknown> | undefined;
                            if (raw && typeof raw === 'object') {
                                const ident = raw['identifier'] as Record<string, unknown> | undefined;
                                if (ident && typeof ident === 'object') {
                                    const idObj = ident['id'] as unknown as Record<string, unknown> | undefined;
                                    const gid = idObj && ((idObj['guid'] as unknown) || (idObj['Id'] as unknown) || (idObj['value'] as unknown));
                                    if (gid) { out.push(String(gid)); return; }
                                }
                                // fallback: sometimes raw has an 'id' or 'Id' directly
                                const rawAsRecord = raw as unknown as Record<string, unknown>;
                                if (rawAsRecord['id'] && typeof rawAsRecord['id'] === 'string') { out.push(String(rawAsRecord['id'])); return; }
                            }
                        } catch { /* ignore raw */ }
                    }
                } catch { /* ignore */ }
                if (id) out.push(id);
            });
        } catch { /* ignore */ }
        return out;
        }
}
