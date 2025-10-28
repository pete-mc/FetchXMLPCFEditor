import { makeAutoObservable } from 'mobx';

export interface IEntityListItem { logicalName: string; displayName: string }
export interface IFieldMetadata { logicalName: string; displayName: string; attributeType: string }

export class QueryBuilderStore {
  entityList: IEntityListItem[] = [];
  selectedEntity = '';
  availableFields: IFieldMetadata[] = [];
  selectedFieldNames: string[] = [];
  loadingFields = false;

  getFieldsForEntity?: (logicalName: string) => Promise<IFieldMetadata[]>;

  constructor() {
    makeAutoObservable(this, {}, { autoBind: true });
  }

  setEntityList(list: IEntityListItem[]) {
    this.entityList = list;
  }

  setSelectedEntity(name: string) {
    this.selectedEntity = name;
    void this.loadFieldsForSelectedEntity();
  }

  setSelectedFieldNames(names: string[]) {
    this.selectedFieldNames = Array.isArray(names) ? names.slice() : [];
  }

  async loadFieldsForSelectedEntity() {
    const name = this.selectedEntity;
    if (!name || !this.getFieldsForEntity) {
      this.availableFields = [];
      this.selectedFieldNames = [];
      return;
    }
    this.loadingFields = true;
    try {
      const f = await this.getFieldsForEntity(name);
      this.availableFields = f;
      if (!this.selectedFieldNames || this.selectedFieldNames.length === 0) {
        this.selectedFieldNames = f.slice(0, 8).map((ff) => ff.logicalName);
      }
    } catch {
      this.availableFields = [];
      this.selectedFieldNames = [];
    } finally {
      this.loadingFields = false;
    }
  }

  toggleField(name: string) {
    if (this.selectedFieldNames.includes(name)) this.selectedFieldNames = this.selectedFieldNames.filter((n) => n !== name);
    else this.selectedFieldNames = [...this.selectedFieldNames, name];
  }
}

export default QueryBuilderStore;
