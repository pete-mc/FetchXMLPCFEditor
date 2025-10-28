import * as React from 'react';
import { QueryBuilderComponent, RuleModel } from '@syncfusion/ej2-react-querybuilder';
import { ColumnsModel } from '@syncfusion/ej2-querybuilder';

// Syncfusion styles (Fluent theme) and icons — required so the QueryBuilder buttons/icons render
import '@syncfusion/ej2-base/styles/fluent.css';
import '@syncfusion/ej2-react-querybuilder/styles/fluent.css';
import '@syncfusion/ej2-icons/styles/material.css';

// Note: CSS is packaged as local files and referenced by the manifest. Keep imports for bundling.

/**
 * Props for the FetchXmlQueryBuilder component.
 * - value: the FetchXML string stored in the bound text field (may be empty)
 * - onChange: called with the new FetchXML when the user updates the query
 * - disabled: optional UI disabled state
 */
export interface FetchXmlQueryBuilderProps {
  value?: string | null;
  onChange: (fetchXml: string) => void;
  disabled?: boolean;
  /** Optional externally-provided columns (from Dataverse metadata) */
  columns?: { field: string; label?: string; type?: string }[];
}

/**
 * Production-ready React component that embeds Syncfusion's QueryBuilder and
 * maps its rules to/from Microsoft Dataverse FetchXML stored in a single-line text field.
 *
 * Responsibilities:
 * - Parse props.value (FetchXML) on load and when it changes, initialize the QueryBuilder model.
 * - Convert QueryBuilder rules back into FetchXML and call props.onChange(fetchXml) when rules change.
 * - Provide a clean Fluent-like container so the control looks native in Power Apps.
 */
const FetchXmlQueryBuilder: React.FC<FetchXmlQueryBuilderProps> = ({ value, onChange, disabled, columns: propsColumns }) => {
  const [rule, setRule] = React.useState<RuleModel>({ condition: 'and', rules: [] });
  // Local column definition used to configure the QueryBuilder columns
  interface ColumnDef {
    field: string;
    label?: string;
    type?: string;
  }
  // Lightweight rule-like shape used for traversing/building rules
  interface RuleLike {
    field?: string;
    operator?: string;
    value?: string | number | boolean | (string | number)[] | undefined;
    condition?: string;
    rules?: RuleLike[];
  }
  const [fields, setFields] = React.useState<ColumnDef[]>([]);
  const qbRef = React.useRef<QueryBuilderComponent | null>(null);
  const entityNameRef = React.useRef<string>('entity');

  // Helper: escape XML values
  const escapeXml = (s: unknown) => {
    if (s === undefined || s === null) return '';
    let str: string;
    if (typeof s === 'string' || typeof s === 'number' || typeof s === 'boolean') {
      str = String(s);
    } else {
      // avoid object default stringification
      return '';
    }
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  };

  // Map common QueryBuilder operator names to FetchXML operator names.
  // This covers common cases. Extend if you use custom operators.
  const operatorToFetch = (op?: string) => {
    if (!op) return 'eq';
    const map: Record<string, string> = {
      equal: 'eq',
      '=': 'eq',
      notequal: 'ne',
      '!=': 'ne',
      lessthan: 'lt',
      greaterthan: 'gt',
      lessthanorequal: 'le',
      greaterthanorequal: 'ge',
      contains: 'like',
      notcontains: 'not-like',
      like: 'like',
      startswith: 'like',
      endswith: 'like',
      'in': 'in',
      'not-in': 'not-in',
      null: 'null',
      'not-null': 'not-null',
    };
    return map[op.toLowerCase()] || op;
  };

  // Reverse map if needed in the future (QueryBuilder will supply its own operator names).

  const parseFetchXml = (fetchXml?: string | null) => {
    const defaultResult: { rule: RuleModel; columns: ColumnDef[]; entityName: string } = { rule: { condition: 'and', rules: [] }, columns: [], entityName: '' };
    if (!fetchXml) return defaultResult;

    try {
      const dom = new DOMParser().parseFromString(fetchXml, 'text/xml');
      const fetchEl = dom.getElementsByTagName('fetch')[0];
      if (!fetchEl) return defaultResult;
      const entityEl = fetchEl.getElementsByTagName('entity')[0];
      if (!entityEl) return defaultResult;

      const entityName = entityEl.getAttribute('name') ?? '';

      // attributes
  const attrNodes = Array.from(entityEl.getElementsByTagName('attribute'));
  const columns: ColumnDef[] = attrNodes.map((n) => ({ field: n.getAttribute('name') ?? '', label: n.getAttribute('name') ?? '', type: 'string' }));

      // recursive filter parser
      const conditionToRule = (condEl: Element) => {
        const attribute = condEl.getAttribute('attribute') ?? '';
        const operator = condEl.getAttribute('operator') ?? 'eq';
        // value attribute or nested <value> elements
        let value: string | string[] | undefined;
        if (condEl.hasAttribute('value')) {
          value = condEl.getAttribute('value') ?? '';
        } else {
          const values = Array.from(condEl.getElementsByTagName('value'));
          if (values.length > 0) value = values.map((v) => v.textContent ?? '');
        }
        // strip % for like
        if (operator.toLowerCase() === 'like' && typeof value === 'string') {
          value = value.replace(/^%+/, '').replace(/%+$/, '');
        }
        return { field: attribute, operator, value } as RuleModel;
      };

      const filterToRule = (filterEl: Element): RuleModel => {
        const type = (filterEl.getAttribute('type') ?? 'and').toLowerCase() as 'and' | 'or';
        const rules: RuleModel[] = [];
        // direct condition children
        const conds = Array.from(filterEl.children).filter((c) => c.tagName === 'condition' || c.tagName === 'filter');
        conds.forEach((c) => {
          if (c.tagName === 'condition') rules.push(conditionToRule(c));
          else if (c.tagName === 'filter') rules.push(filterToRule(c));
        });
        return { condition: type, rules };
      };

      // top-level filters under entity
  const filterEls = Array.from(entityEl.getElementsByTagName('filter'));
      let rootRule: RuleModel = { condition: 'and', rules: [] };
      if (filterEls.length === 1) rootRule = filterToRule(filterEls[0]);
      else if (filterEls.length > 1) rootRule.rules = filterEls.map((f) => filterToRule(f));

      // add referenced fields found in conditions
      const referenced = new Set<string>();
      const collect = (r: RuleLike | undefined) => {
        if (!r) return;
        if (r.field) referenced.add(r.field);
        if (r.rules) r.rules.forEach((rr) => collect(rr));
      };
      collect(rootRule as unknown as RuleLike);
      referenced.forEach((f) => {
        if (!columns.find((c) => c.field === f)) columns.push({ field: f, label: f, type: 'string' });
      });

      entityNameRef.current = entityName || 'entity';
      return { rule: rootRule, columns, entityName };
    } catch (_e) {
      return defaultResult;
    }
  };

  // Convert a QueryBuilder RuleModel back into FetchXML string
  const buildFetchXml = (ruleModel: RuleModel, entityName = 'entity', attributes: ColumnDef[] = []) => {
    // build attributes xml
  const uniqueAttrs = new Set<string>(attributes.map((c) => c.field).filter(Boolean));

    // also include any referenced fields from rules
    const collect = (r: RuleLike | undefined) => {
      if (!r) return;
      if (r.field) uniqueAttrs.add(r.field);
      if (r.rules) r.rules.forEach((rr) => collect(rr));
    };
    // RuleModel is structurally compatible with RuleLike so this is safe
    collect(ruleModel as unknown as RuleLike);

    const attrsXml = Array.from(uniqueAttrs)
      .map((a) => `    <attribute name="${escapeXml(a)}" />`)
      .join('\n');

    // convert rule to xml recursively
    const ruleToXml = (r?: RuleLike): string => {
      if (!r) return '';
      if (r.rules && Array.isArray(r.rules)) {
        const type = (r.condition ?? 'and').toLowerCase();
        const children = r.rules.map((child) => ruleToXml(child)).filter(Boolean).join('\n');
        return `    <filter type="${escapeXml(type)}">\n${children}\n    </filter>`;
      }

      // Leaf condition
      const field = r.field ?? '';
      const operator = operatorToFetch(r.operator ?? '');
      const value = r.value;

      // No-value operators (null/not-null)
      if (!operator) return '';
      const opLower = operator.toLowerCase();
      if (opLower === 'null' || opLower === 'not-null') {
        return `    <condition attribute="${escapeXml(field)}" operator="${escapeXml(opLower)}" />`;
      }

      if (opLower === 'in' || opLower === 'not-in') {
        // value expected as array or comma-separated
        const vals = Array.isArray(value) ? value : String(value ?? '').split(',').map((s) => s.trim());
        const vs = vals.map((v) => `      <value>${escapeXml(v)}</value>`).join('\n');
        return `    <condition attribute="${escapeXml(field)}" operator="${escapeXml(opLower)}">\n${vs}\n    </condition>`;
      }

      // like/contains/startswith/endswith: map to 'like' and add % where appropriate
      if (opLower === 'like' || opLower === 'contains' || opLower === 'startswith' || opLower === 'endswith') {
        let s = String(value ?? '');
        // If operator was contains -> wrap both sides
        const opRaw = (r.operator ?? '').toString().toLowerCase();
        if (opRaw === 'contains') s = `%${s}%`;
        if (opRaw === 'startswith') s = `${s}%`;
        if (opRaw === 'endswith') s = `%${s}`;
        return `    <condition attribute="${escapeXml(field)}" operator="like" value="${escapeXml(s)}" />`;
      }

      // default simple condition with value attribute
      return `    <condition attribute="${escapeXml(field)}" operator="${escapeXml(opLower)}" value="${escapeXml(value ?? '')}" />`;
    };

  const filtersXml = ruleToXml(ruleModel as unknown as RuleLike) || '';

    const fetchXml = `<fetch mapping="logical">\n  <entity name="${escapeXml(entityName)}">\n${attrsXml}\n${filtersXml ? '\n' + filtersXml : ''}\n  </entity>\n</fetch>`;

    return fetchXml;
  };

  // Initialize from props.value
  React.useEffect(() => {
    const { rule: parsedRule, columns } = parseFetchXml(value);
    setRule(parsedRule);
    // If external columns provided, prefer those; otherwise use parsed columns
    if (propsColumns?.length) {
      setFields(propsColumns as ColumnDef[]);
    } else if (columns?.length) setFields(columns);
  }, [value]);

  // When external columns change (for example when the field checkboxes are toggled),
  // update the local fields and emit a new FetchXML that reflects the selected attributes.
  React.useEffect(() => {
  const cols = (propsColumns?.length) ? (propsColumns as ColumnDef[]) : fields;
    setFields(cols);
    try {
      const fetchXml = buildFetchXml(rule, entityNameRef.current || 'entity', cols);
      onChange(fetchXml);
    } catch (_e) {
      // ignore
    }
    // we purposely omit `rule` from deps to avoid an immediate loop; rule changes are handled by onRuleChange
  }, [propsColumns]);

  // Inject small CSS overrides to ensure Syncfusion popup/dropdown elements render correctly
  // inside the Power Apps host (avoid being clipped or positioned at top-left).
  React.useEffect(() => {
    const id = 'fetchxml-qb-popup-fix';
    if (document.getElementById(id)) return;
    const style = document.createElement('style');
    style.id = id;
    style.innerHTML = `
      /* Let Syncfusion compute popup position (absolute) — only increase z-index so popups appear above host chrome. */
      .e-popup, .e-dropdown-popup, .e-ddl.e-popup {
        position: absolute !important;
        z-index: 2147483000 !important; /* very high but below browser max */
        background: white !important; /* ensure visible against host chrome */
        border: 1px solid rgba(0,0,0,0.08) !important;
      }
      /* keep popup internals in normal flow */
      .e-popup .e-list-parent, .e-popup .e-content {
        position: relative !important;
      }
    `;
    document.head.appendChild(style);
    return () => { try { style.remove(); } catch { /* ignore */ } };
  }, []);

  // Memoize fields for QueryBuilder configuration (Syncfusion ColumnsModel ~ ColumnModel here)
  const qbColumns = React.useMemo<ColumnsModel[]>(() => {
    // Syncfusion QueryBuilder expects columns: { field: string, label?: string, type?: string }
    const opsForType = (t?: string) => {
  const tt = (t ?? 'string').toLowerCase();
      const mk = (v: string, text?: string) => ({ key: v, text: text ?? v });
      if (tt === 'date' || tt === 'datetime' || tt === 'datetimeoffset') {
        return [mk('eq', 'Equals'), mk('ne', 'Not equals'), mk('lt', 'Less than'), mk('gt', 'Greater than'), mk('le', 'Less or equal'), mk('ge', 'Greater or equal'), mk('between', 'Between'), mk('null', 'Null'), mk('not-null', 'Not null')];
      }
      if (tt === 'number' || tt === 'int' || tt === 'integer' || tt === 'decimal' || tt === 'double' || tt === 'money') {
        return [mk('eq', 'Equals'), mk('ne', 'Not equals'), mk('lt', 'Less than'), mk('gt', 'Greater than'), mk('le', 'Less or equal'), mk('ge', 'Greater or equal'), mk('in', 'In'), mk('not-in', 'Not in'), mk('null', 'Null'), mk('not-null', 'Not null')];
      }
      if (tt === 'boolean' || tt === 'bool') {
        return [mk('eq', 'Equals'), mk('ne', 'Not equals'), mk('null', 'Null'), mk('not-null', 'Not null')];
      }
      // default string
      return [mk('equal', 'Equals'), mk('notequal', 'Not equals'), mk('contains', 'Contains'), mk('notcontains', 'Not contains'), mk('startswith', 'Starts with'), mk('endswith', 'Ends with'), mk('null', 'Null'), mk('not-null', 'Not null')];
    };

    return fields.map((f) => {
      const type = f.type ?? 'string';
  const ops = opsForType(type) as ColumnsModel['operators'];
      // Provide primitive values for boolean fields so the QueryBuilder renders a proper dropdown
      const values = (type.toLowerCase() === 'boolean' || type.toLowerCase() === 'bool') ? [true, false] : undefined;
  const col: ColumnsModel = { field: f.field, label: f.label ?? f.field, type, operators: ops, values: values } as ColumnsModel;
      // ColumnsModel typing in this build may be loose; return the constructed ColumnsModel
      return col;
    });
  }, [fields]);

  // Called by QueryBuilder when rules change
  const onRuleChange = (args: unknown) => {
    // args may be a ChangeEvent with a 'value' property or a RuleModel directly
    let newRule: RuleModel;
    const isChangeEvent = (a: unknown): a is { value?: RuleModel } => {
      return typeof a === 'object' && a !== null && Object.prototype.hasOwnProperty.call(a, 'value');
    };

    if (isChangeEvent(args)) {
      newRule = args.value ?? ({} as RuleModel);
    } else {
      newRule = args as RuleModel;
    }
    setRule(newRule);
    try {
      // try to preserve entity name from incoming value if possible
      const parsed = parseFetchXml(value);
      const fetchXml = buildFetchXml(newRule, parsed.entityName ?? 'entity', fields);
      onChange(fetchXml);
    } catch (err) {
      // if conversion fails, send an empty fetch wrapper to avoid breaking PCF
      onChange('<fetch />');
    }
  };

  // Small native-looking container styles
  const containerStyle: React.CSSProperties = {
    borderRadius: 6,
    border: '1px solid rgba(0,0,0,0.08)',
    padding: 8,
    background: 'transparent',
    fontFamily: 'Segoe UI, Roboto, system-ui, -apple-system, "Helvetica Neue", Arial',
    minHeight: 120,
  };

  return (
    <div style={containerStyle} aria-disabled={disabled}>
      {qbColumns.length === 0 ? (
        <div style={{ color: '#666', padding: '12px 8px' }}>
          No attributes found in FetchXML. Add attributes to the FetchXML or use a default entity to start building a query.
        </div>
      ) : (
        <ErrorBoundary
          onCatch={(err: Error, info: React.ErrorInfo) => {
            // Attach a debug dump to window so it can be inspected in the host console
            try {
              const dump = {
                time: new Date().toISOString(),
                error: { message: err.message, stack: err.stack },
                info,
                rule,
                qbColumns,
                fields,
                qbRef: qbRef.current ? { hasInstance: true } : { hasInstance: false },
              };
              // eslint-disable-next-line @typescript-eslint/ban-ts-comment
              // @ts-ignore
              window.__fetchXmlQbLastError = dump;
              console.error('FetchXmlQueryBuilder caught error:', dump);
            } catch (e) {
              // best-effort
              console.error('Failed to write fetchxml-qb dump', e);
            }
          }}
        >
          <QueryBuilderComponent
            ref={(r: QueryBuilderComponent | null) => {
              qbRef.current = r;
            }}
            width="100%"
            rule={rule}
            columns={qbColumns}
            change={onRuleChange}
            readOnly={!!disabled}
          />
        </ErrorBoundary>
      )}
    </div>
  );
};

// Simple ErrorBoundary to catch rendering/runtime exceptions originating from
// Syncfusion internals (they throw during some DOM interactions in host). We
// capture a small debug dump and prevent the entire PCF control from unmounting.
class ErrorBoundary extends React.Component<{
  children: React.ReactNode;
  onCatch?: (err: Error, info: React.ErrorInfo) => void;
}> {
  state = { hasError: false };
  static getDerivedStateFromError() {
    return { hasError: true };
  }
  componentDidCatch(error: Error, info: React.ErrorInfo) {
    this.setState({ hasError: true });
    if (this.props.onCatch) this.props.onCatch(error, info);
  }
  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: 12, color: '#900', background: '#fff6f6', borderRadius: 6 }}>
          An internal error occurred while rendering the query editor. Open the browser console for details.
        </div>
      );
    }
    return this.props.children as React.ReactElement;
  }
}

export default FetchXmlQueryBuilder;
