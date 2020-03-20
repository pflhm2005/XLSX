export default class JimmyMap {
  constructor(defaultValue = null, needStringify = true) {
    this.map = new Map();
    this.needStringify = needStringify;
    this.uid = -1;
    this.defaultValue = defaultValue;
    if(this.defaultValue !== null) {
      if(this.needStringify) this.defaultValue = JSON.stringify(defaultValue);
      this.uid++;
      this.map.set(this.defaultValue, this.uid);
    }
  }
  LookOrInsert(key) {
    if(this.needStringify) key = JSON.stringify(key);
    let tar = this.map.get(key);
    if(tar === undefined) {
      this.uid++;
      this.map.set(key, this.uid);
      return this.uid;
    } else {
      return tar;
    }
  }
  clearUp() {
    this.map.clear();
    this.uid = -1;
    if(this.defaultValue !== null) {
      this.uid++;
      this.map.set(this.defaultValue, this.uid);
    }
  }
  size() {
    return this.map.size;
  }
}