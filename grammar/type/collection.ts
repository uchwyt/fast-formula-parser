/**
 * Represents unions of references (e.g. A1, A1:C5, ...)
 */
export class Collection {
  private _data: any[];
  private _refs: any[];

  /**
   * Create a new collection of references and their values
   * @param data Optional initial data array
   * @param refs Optional initial references array
   */
  constructor(data?: any[], refs?: any[]) {
    if (data == null && refs == null) {
      this._data = [];
      this._refs = [];
    } else {
      if (!data || !refs || data.length !== refs.length) {
        throw Error('Collection: data length should match references length.');
      }
      this._data = data;
      this._refs = refs;
    }
  }

  /**
   * Get the data array
   */
  get data(): any[] {
    return this._data;
  }

  /**
   * Get the references array
   */
  get refs(): any[] {
    return this._refs;
  }

  /**
   * Get the number of items in the collection
   */
  get length(): number {
    return this._data.length;
  }

  /**
   * Add data and references to this collection
   * @param obj Data value
   * @param ref Reference object
   */
  add(obj: any, ref: any): void {
    this._data.push(obj);
    this._refs.push(ref);
  }
}

export default Collection;
