import * as React from 'react';
import { IEntityListItem } from '../viewmodels/QueryBuilderStore';

interface Props {
  entityList?: IEntityListItem[];
  value?: string;
  onChange: (value: string) => void;
  loading?: boolean;
  showManual?: boolean;
  errorMessage?: string;
}

const EntitySelector: React.FC<Props> = ({ entityList = [], value, onChange, loading, showManual, errorMessage }) => {
  return (
    <div style={{ marginBottom: 8 }}>
      <label style={{ display: 'block', marginBottom: 4, fontSize: 12 }}>Table</label>
      {showManual ? (
        <>
          {errorMessage ? <div style={{ color: '#a00', marginBottom: 6, fontSize: 12 }}>{errorMessage}</div> : null}
          <input value={value} onChange={(e) => onChange(e.target.value)} placeholder="Enter logical name (e.g. account)" style={{ width: '100%' }} />
        </>
      ) : (
        <select value={value} onChange={(e) => onChange(e.target.value)} style={{ width: '100%' }}>
          <option value="">-- select table --</option>
          {entityList && entityList.length > 0 ? (
            entityList.map((e) => (
              <option key={e.logicalName} value={e.logicalName}>{e.displayName || e.logicalName}</option>
            ))
          ) : (
            <option value="" disabled>{loading ? "Loading entities..." : "No entities available"}</option>
          )}
        </select>
      )}
    </div>
  );
};

export default EntitySelector;
