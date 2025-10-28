import * as React from 'react';
import { IFieldMetadata } from '../viewmodels/QueryBuilderStore';

interface Props {
  fields: IFieldMetadata[];
  selected: string[];
  loading?: boolean;
  onToggleField: (logicalName: string) => void;
}

const attributeTypeToColumnType = (t: string | undefined) => {
  if (!t) return 'string';
  const tt = t.toLowerCase();
  if (['string', 'memo', 'sstring', 'memo'].includes(tt)) return 'string';
  if (['integer', 'int', 'bigint', 'decimal', 'double', 'money'].includes(tt)) return 'number';
  if (tt === 'datetime' || tt === 'dateandtime' || tt === 'datetimeoffset') return 'date';
  if (tt === 'boolean' || tt === 'bool') return 'boolean';
  return 'string';
};

const FieldPicker: React.FC<Props> = ({ fields, selected, loading, onToggleField }) => {
  const visible = fields;

  return (
    <div>
      <div>
        <label style={{ display: 'block', marginBottom: 4, fontSize: 12 }}>Fields</label>
        <div style={{ maxHeight: 280, overflow: 'auto', border: '1px solid rgba(0,0,0,0.04)', padding: 8 }}>
          {loading ? (
            <div style={{ color: '#666' }}>Loading fields…</div>
          ) : visible.length === 0 ? (
            <div style={{ color: '#666' }}>No fields available for the selected table.</div>
          ) : (
            visible.map((f) => (
              <div key={f.logicalName} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <input type="checkbox" checked={selected.includes(f.logicalName)} onChange={() => onToggleField(f.logicalName)} />
                <div style={{ fontSize: 13 }}>{f.displayName || f.logicalName} <span style={{ color: '#888', fontSize: 11 }}>({attributeTypeToColumnType(f.attributeType)})</span></div>
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
};

export default FieldPicker;
