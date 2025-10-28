/* eslint-disable */
import * as React from 'react';
import { observer } from 'mobx-react-lite';
import QueryBuilderStore, { IEntityListItem, IFieldMetadata } from '../viewmodels/QueryBuilderStore';
import EntitySelector from './EntitySelector';
import FieldPicker from './FieldPicker';
import QueryBuilderView from './QueryBuilderView';

export interface IEditorShellProps {
  value?: string | null;
  onChange: (fetchXml: string) => void;
  disabled?: boolean;
  entityList?: IEntityListItem[];
  getFieldsForEntity?: (logicalName: string) => Promise<IFieldMetadata[]>;
  getRawFieldsForEntity?: (logicalName: string) => Promise<any>;
  getEntities?: () => Promise<IEntityListItem[]>;
}

const EditorShellInner: React.FC<IEditorShellProps> = ({ value, onChange, disabled, entityList = [], getFieldsForEntity, getRawFieldsForEntity, getEntities }) => {
  const storeRef = React.useRef<QueryBuilderStore | null>(null);
  if (!storeRef.current) storeRef.current = new QueryBuilderStore();
  const store = storeRef.current;
  const [entitiesLoading, setEntitiesLoading] = React.useState(false);
  const [entitiesError, setEntitiesError] = React.useState<string | null>(null);

  const mountedRef = React.useRef(true);

  const load = React.useCallback(async () => {
    if (!mountedRef.current) return;
    setEntitiesLoading(true);
    setEntitiesError(null);
    if (typeof getEntities === 'function') {
      try {
        const list = await getEntities();
        if (!mountedRef.current) return;
        if (list?.length) store.setEntityList(list);
        else store.setEntityList(entityList);
      } catch (err) {
        if (!mountedRef.current) return;
        store.setEntityList(entityList);
        const msg = (err && typeof err === 'object' && 'message' in err) ? String((err as any).message) : 'Failed to load entities';
        setEntitiesError(msg);
  // also log for convenience
  console.error('[FetchXMLEditor] getEntities error', err);
      } finally {
        if (!mountedRef.current) return;
        setEntitiesLoading(false);
      }
    } else {
      store.setEntityList(entityList);
      if (!mountedRef.current) return;
      setEntitiesLoading(false);
    }
  }, [getEntities, entityList, store]);

  React.useEffect(() => {
    mountedRef.current = true;
    void load();
    return () => {
      mountedRef.current = false;
    };
  }, [load]);

  React.useEffect(() => {
    store.getFieldsForEntity = getFieldsForEntity;
  }, [getFieldsForEntity, store]);

  React.useEffect(() => {
    if (!value) return;
    const m = /<entity[^>]*name=["']([^"']+)["']/i.exec(value);
    if (m?.[1]) store.setSelectedEntity(m[1]);
  }, [value, store]);

  const columns = React.useMemo(() => {
    return store.selectedFieldNames.map((name) => {
      const meta = store.availableFields.find((f) => f.logicalName === name);
      return meta ? { field: meta.logicalName, label: meta.displayName || meta.logicalName, type: (meta.attributeType ?? 'string') } : { field: name, label: name, type: 'string' };
    });
  }, [store.selectedFieldNames, store.availableFields]);

  const paneStyle: React.CSSProperties = { width: 280, borderRight: '1px solid rgba(0,0,0,0.06)', paddingRight: 12 };

  const [debugJson, setDebugJson] = React.useState<any>(null);
  const [debugLoading, setDebugLoading] = React.useState(false);

  const runDebug = async () => {
    if (!store.selectedEntity || !getRawFieldsForEntity) return;
    setDebugLoading(true);
    try {
      const raw = await getRawFieldsForEntity(store.selectedEntity);
      setDebugJson(raw);
    } catch (e) {
      setDebugJson({ error: String(e) });
    } finally {
      setDebugLoading(false);
    }
  };

  return (
    <div style={{ display: 'flex' }}>
      <div style={paneStyle}>
        {entitiesError ? (
          <div style={{ marginBottom: 8, color: '#a00' }}>
            <div style={{ marginBottom: 6, fontSize: 12 }}>Failed to load metadata: {entitiesError}</div>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={() => { void load(); }} style={{ padding: '6px 8px' }}>Retry</button>
              <div style={{ alignSelf: 'center', color: '#666', fontSize: 12 }}>Or enter a logical name manually below.</div>
            </div>
          </div>
        ) : null}
  <EntitySelector entityList={store.entityList} value={store.selectedEntity} onChange={(v) => store.setSelectedEntity && store.setSelectedEntity(v)} loading={entitiesLoading} showManual={!entitiesLoading && store.entityList.length === 0} errorMessage={entitiesError ?? undefined} />
        <FieldPicker
          fields={store.availableFields}
          selected={store.selectedFieldNames}
          loading={store.loadingFields}
          onToggleField={(n) => store.toggleField(n)}
        />
      </div>
      <div style={{ flex: 1, paddingLeft: 12 }}>
        <QueryBuilderView value={value} onChange={onChange} disabled={disabled} columns={columns} />
      </div>
    </div>
  );
};

const EditorShell = observer(EditorShellInner);

export default EditorShell;
export { EditorShell };
