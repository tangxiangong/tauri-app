export class Memory {
  totalMemory: number;
  usedMemory: number;
  totalSwap: number;
  usedSwap: number;
  constructor(
    totalMemory: number,
    usedMemory: number,
    totalSwap: number,
    usedSwap: number,
  ) {
    this.totalMemory = totalMemory;
    this.usedMemory = usedMemory;
    this.totalSwap = totalSwap;
    this.usedSwap = usedSwap;
  }
}
