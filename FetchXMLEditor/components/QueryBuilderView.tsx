import * as React from 'react';
import FetchXmlQueryBuilder from '../FetchXmlQueryBuilder';

interface Props {
  value?: string | null;
  onChange: (fetchXml: string) => void;
  disabled?: boolean;
  columns?: { field: string; label?: string; type?: string }[];
}

const QueryBuilderView: React.FC<Props> = (props) => {
  return <FetchXmlQueryBuilder {...props} />;
};

export default QueryBuilderView;
