/* eslint-disable */
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import EditorShell, { IEditorShellProps as IHelloWorldProps } from "./components/EditorShell";
import * as React from "react";

// Narrow response shapes for metadata queries
// minimal typed shapes used for safe extraction from unknown responses
type RawEntity = { logicalname?: string; LogicalName?: string; displayname?: { userLocalizedLabel?: { label?: string } } | undefined; DisplayName?: string };
type RawAttribute = { logicalname?: string; LogicalName?: string; displayname?: { userLocalizedLabel?: { label?: string } } | undefined; DisplayName?: string; attributetype?: string; AttributeType?: string };

export class FetchXMLEditor implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private latestFetchXml: string | null = null;
    private entityList: { logicalName: string; displayName: string }[] = [];
    private fieldCache: Record<string, { logicalName: string; displayName: string; attributeType: string }[]> = {};
    // no caching flag here; always allow refresh on demand

    // (response interfaces are declared at top-level)

    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    private extractLabel(obj: unknown): string | undefined {
        if (!obj || typeof obj !== 'object') return undefined;
        const o = obj as { userLocalizedLabel?: { label?: string } };
        if (o.userLocalizedLabel && typeof o.userLocalizedLabel.label === 'string') return o.userLocalizedLabel.label;
        return undefined;
    }

    private extractEntity(raw: unknown): { logicalName: string; displayName: string } {
        if (!raw || typeof raw !== 'object') return { logicalName: '', displayName: '' };
        const r = raw as Record<string, unknown>;
        const logical = (r['logicalname'] ?? r['LogicalName'] ?? '') as string;
        let display = logical;
        if (typeof r['DisplayName'] === 'string') display = r['DisplayName'] as string;
        else if (r['displayname']) {
            const dl = this.extractLabel(r['displayname']);
            if (dl) display = dl;
        }
        return { logicalName: String(logical), displayName: String(display) };
    }

    private extractAttribute(raw: unknown): { logicalName: string; displayName: string; attributeType: string } {
        if (!raw || typeof raw !== 'object') return { logicalName: '', displayName: '', attributeType: '' };
        const r = raw as Record<string, unknown>;
        const logical = (r['logicalname'] ?? r['LogicalName'] ?? '') as string;
        let display = logical;
        if (typeof r['DisplayName'] === 'string') display = r['DisplayName'] as string;
        else if (r['displayname']) {
            const dl = this.extractLabel(r['displayname']);
            if (dl) display = dl;
        }
        // Normalize Dataverse attribute types into simple types consumed by the QueryBuilder
        const attributeTypeRaw = String((r['attributetype'] ?? r['AttributeType'] ?? '') || '').toLowerCase();
        // Map common Dataverse metadata types to QueryBuilder-friendly types
        const map: Record<string, string> = {
            'string': 'string',
            'memo': 'string',
            'boolean': 'boolean',
            'datetime': 'date',
            'datetime2': 'date',
            'datetimeoffset': 'date',
            'dateandtime': 'date',
            'integer': 'number',
            'int': 'number',
            'decimal': 'number',
            'double': 'number',
            'money': 'number',
            'lookup': 'string',
            'owner': 'string',
            'customer': 'string',
            'partylist': 'string',
            'picklist': 'string',
            'optionset': 'string',
            'uniqueidentifier': 'string'
        };
        const normalized = map[attributeTypeRaw] ?? 'string';
        return { logicalName: String(logical), displayName: String(display), attributeType: normalized };
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        const current = context.parameters.sampleProperty.raw;
        // keep the most recent value for getOutputs
        this.latestFetchXml = current ?? null;

        // ensure we have an entity list cached
        if (!this.entityList.length) {
            this.loadEntities(context).catch(() => {
                // ignore errors; HelloWorld can still function with manual entity entry
            });
        }

        const props: IHelloWorldProps = {
            value: current ?? null,
            onChange: (fetchXml: string) => {
                // update internal cache and notify framework to fetch outputs
                this.latestFetchXml = fetchXml;
                this.notifyOutputChanged();
            },
            disabled: !!context.mode.isControlDisabled,
            entityList: this.entityList,
            getEntities: async () => {
                // always attempt to refresh entities from the API on request
                await this.loadEntities(context).catch(() => { /* ignore */ });
                return this.entityList;
            },
            getFieldsForEntity: async (logicalName: string) => {
                // return cached fields when available
                if (this.fieldCache[logicalName]) return this.fieldCache[logicalName];
                try {
                    const fields = await this.loadFields(context, logicalName);
                    this.fieldCache[logicalName] = fields;
                    return fields;
                } catch {
                    return [];
                }
            }
            ,
            // Debug helper for raw attribute metadata — returns the raw object or null
            getRawFieldsForEntity: async (logicalName: string) => {
                try {
                    const path = "/api/data/v9.2/EntityDefinitions(LogicalName='" + encodeURIComponent(logicalName) + "')/Attributes?$select=LogicalName,DisplayName,AttributeType";
                    // prefer performRequest when available
                    if ((context.webAPI as any)?.performRequest) {
                        const perf = await (context.webAPI as any).performRequest({ path, method: 'GET' } as any);
                        try { return typeof perf === 'string' ? JSON.parse(perf) : perf; } catch { return perf; }
                    }
                    // otherwise try fetch
                    const res = await fetch(path, { credentials: 'same-origin', headers: { Accept: 'application/json' } });
                    const json = await res.json().catch(() => null);
                    return json;
                } catch (e) {
                    return null;
                }
            }
        };

    return React.createElement(EditorShell, props as unknown as React.ComponentProps<typeof EditorShell>);
    }

    private async loadEntities(context: ComponentFramework.Context<IInputs>): Promise<void> {
        // Query global metadata for entity display names
        try {
            // Try a few common OData entity set names for metadata
            const candidates = [
                { set: 'EntityDefinitions', qs: "?$select=LogicalName,DisplayName&$filter=IsPrivate ne true" },
                { set: 'EntityDefinition', qs: "?$select=LogicalName,DisplayName&$filter=IsPrivate ne true" },
                { set: 'EntityMetadata', qs: "?$select=LogicalName,DisplayName&$filter=IsPrivate ne true" }
            ];
            let success = false;
            for (const c of candidates) {
                try {
                    console.log('[FetchXMLEditor] trying metadata set', c.set);
                    const resp = await context.webAPI.retrieveMultipleRecords(c.set, c.qs);
                    if (resp && Array.isArray((resp as any).entities)) {
                        this.entityList = (resp as any).entities.map((raw: unknown) => this.extractEntity(raw));
                        console.log('[FetchXMLEditor] loaded entities count=', this.entityList.length, 'via', c.set);
                        success = true;
                        break;
                    }
                } catch (innerErr) {
                    console.warn('[FetchXMLEditor] metadata set failed', c.set, innerErr);
                    // try next
                }
            }
            if (!success) {
                // Try a lower-level performRequest to get raw OData JSON in environments where retrieveMultipleRecords is restricted
                try {
                    const path = "/api/data/v9.2/EntityDefinitions?$select=LogicalName,DisplayName&$filter=IsPrivate%20ne%20true";
                    console.log('[FetchXMLEditor] attempting performRequest to', path);
                    // performRequest may not exist in some environments; prefer it when available
                    try {
                        if ((context.webAPI as any)?.performRequest && typeof (context.webAPI as any).performRequest === 'function') {
                            const perf = await (context.webAPI as any).performRequest({ path, method: 'GET' } as any);
                            console.log('[FetchXMLEditor] performRequest response (raw)', perf);
                            let parsed: any = perf;
                            try { if (typeof perf === 'string') parsed = JSON.parse(perf); } catch (e) { /* ignore */ }
                            const arr = parsed?.value ?? parsed?.entities ?? parsed;
                            if (Array.isArray(arr)) {
                                this.entityList = arr.map((raw: unknown) => this.extractEntity(raw));
                                console.log('[FetchXMLEditor] loaded entities count=', this.entityList.length, 'via performRequest');
                                success = true;
                            }
                        } else {
                            throw new Error('performRequest-not-available');
                        }
                    } catch (perfErr) {
                        console.warn('[FetchXMLEditor] performRequest failed or not available', perfErr);
                        // fallback: try a raw fetch against the relative OData endpoint
                        try {
                            const url = path; // relative path on same host
                            console.log('[FetchXMLEditor] attempting fetch to', url);
                            const res = await fetch(url, { credentials: 'same-origin', headers: { Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } });
                            const json = await res.json().catch((e) => { console.warn('[FetchXMLEditor] failed parsing fetch json', e); return null; });
                            console.log('[FetchXMLEditor] fetch response status=', res.status, 'ok=', res.ok, 'json=', json);
                            const arr = json?.value ?? json?.entities ?? json;
                            if (Array.isArray(arr)) {
                                this.entityList = arr.map((raw: unknown) => this.extractEntity(raw));
                                console.log('[FetchXMLEditor] loaded entities count=', this.entityList.length, 'via fetch fallback');
                                success = true;
                            } else {
                                console.warn('[FetchXMLEditor] fetch fallback returned unexpected shape', arr);
                            }
                        } catch (fetchErr) {
                            console.warn('[FetchXMLEditor] fetch fallback failed', fetchErr);
                        }
                        // final fallback: try the older 'Entity' collection via retrieveMultipleRecords
                        if (!success) {
                            try {
                                const resp2 = await context.webAPI.retrieveMultipleRecords('Entity', "?$select=LogicalName,DisplayName&$filter=IsPrivate ne true");
                                this.entityList = (resp2.entities ?? []).map((raw) => this.extractEntity(raw));
                                console.log('[FetchXMLEditor] loaded entities count=', this.entityList.length, 'via Entity fallback');
                                success = true;
                            } catch (finalErr) {
                                console.warn('[FetchXMLEditor] Entity fallback failed', finalErr);
                                throw finalErr;
                            }
                        }
                    }
                } catch (finalErr) {
                    throw finalErr;
                }
            }
        } catch (err) {
            // Ignore — this API may be restricted based on environment; leave entityList empty
            this.entityList = [];
            console.error('[FetchXMLEditor] failed to load entities', err);
        }
    }

    private async loadFields(context: ComponentFramework.Context<IInputs>, logicalName: string): Promise<{ logicalName: string; displayName: string; attributeType: string }[]> {
        try {
            // Use RetrieveMultiple on AttributeDefinitions for the entity
            // endpoint format: RetrieveMultiple on EntityDefinitions(LogicalName='account')/Attributes?$select=LogicalName,DisplayName,AttributeType
            // Query attribute metadata for the entity
            const attrCandidates = [
                { set: 'Attributes', qs: "?$select=LogicalName,DisplayName,AttributeType&$filter=EntityLogicalName eq '" + logicalName + "'" },
                { set: 'Attribute', qs: "?$select=LogicalName,DisplayName,AttributeType&$filter=EntityLogicalName eq '" + logicalName + "'" },
                { set: 'AttributeMetadata', qs: "?$select=LogicalName,DisplayName,AttributeType&$filter=EntityLogicalName eq '" + logicalName + "'" }
            ];
            let attrs: { logicalName: string; displayName: string; attributeType: string }[] = [];
            for (const c of attrCandidates) {
                try {
                    const resp = await context.webAPI.retrieveMultipleRecords(c.set, c.qs);
                    if (resp && Array.isArray((resp as any).entities)) {
                        attrs = (resp as any).entities.map((raw: unknown) => this.extractAttribute(raw));
                        console.log('[FetchXMLEditor] loaded fields for', logicalName, 'count=', attrs.length, 'via', c.set);
                        break;
                    }
                } catch (inner) {
                    console.warn('[FetchXMLEditor] attribute set failed', c.set, inner);
                }
            }
            // If no attrs found via retrieveMultipleRecords, try performRequest or fetch against EntityDefinitions(LogicalName='name')/Attributes
            if ((!attrs || attrs.length === 0) && logicalName) {
                try {
                    const path = "/api/data/v9.2/EntityDefinitions(LogicalName='" + encodeURIComponent(logicalName) + "')/Attributes?$select=LogicalName,DisplayName,AttributeType";
                    console.log('[FetchXMLEditor] attempting attribute metadata via performRequest/fetch to', path);
                    try {
                        if ((context.webAPI as any)?.performRequest && typeof (context.webAPI as any).performRequest === 'function') {
                            const perf = await (context.webAPI as any).performRequest({ path, method: 'GET' } as any);
                            console.log('[FetchXMLEditor] performRequest(attributes) raw=', perf);
                            let parsed: any = perf;
                            try { if (typeof perf === 'string') parsed = JSON.parse(perf); } catch (e) { /* ignore */ }
                            const arr = parsed?.value ?? parsed?.entities ?? parsed;
                            if (Array.isArray(arr)) {
                                attrs = arr.map((raw: unknown) => this.extractAttribute(raw));
                                console.log('[FetchXMLEditor] loaded fields for', logicalName, 'count=', attrs.length, 'via performRequest attributes');
                            }
                        } else {
                            throw new Error('performRequest-not-available');
                        }
                    } catch (perfErr) {
                        console.warn('[FetchXMLEditor] performRequest(attributes) failed or missing', perfErr);
                        try {
                            const res = await fetch(path, { credentials: 'same-origin', headers: { Accept: 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } });
                            const json = await res.json().catch((e) => { console.warn('[FetchXMLEditor] failed parsing attributes fetch json', e); return null; });
                            console.log('[FetchXMLEditor] fetch(attributes) status=', res.status, 'ok=', res.ok, 'json=', json);
                            const arr = json?.value ?? json?.entities ?? json;
                            if (Array.isArray(arr)) {
                                attrs = arr.map((raw: unknown) => this.extractAttribute(raw));
                                console.log('[FetchXMLEditor] loaded fields for', logicalName, 'count=', attrs.length, 'via fetch attributes');
                            }
                        } catch (fetchErr) {
                            console.warn('[FetchXMLEditor] fetch(attributes) failed', fetchErr);
                        }
                    }
                } catch (outer) {
                    console.warn('[FetchXMLEditor] attribute metadata final fallback failed', outer);
                }
            }
            return attrs;
        } catch (err) {
            console.error('[FetchXMLEditor] failed to load fields for', logicalName, err);
            return [];
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return { sampleProperty: this.latestFetchXml ?? undefined };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
