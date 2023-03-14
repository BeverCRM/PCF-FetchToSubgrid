declare global {
  interface String {
    hashCode() : number;
  }
}

export interface String {
  hashCode(): number;
}

export default String.prototype.hashCode = function() {
  if (this.length === 0) return 0;

  let hash = 0;
  for (let i = 0; i < this.length; i++) {
    const charCode = this.charCodeAt(i);
    hash = ((hash << 5) - hash) + charCode;
    hash |= 0; // Convert to 32bit integer
  }
  return hash;
};
