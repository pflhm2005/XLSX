export default class JimmyMap {
  constructor(init = null, needStringify = true) {
    this.map = new Map();
    this.needStringify = needStringify;
    this.uid = -1;
    if(this.init !== null) {
      if(this.needStringify) this.init = JSON.stringify(init);
      this.uid++;
      this.map.set(this.init, this.uid);
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
    if(this.init !== null) {
      this.uid++;
      this.map.set(this.init, this.uid);
    }
  }
  size() {
    return this.map.size;
  }
}