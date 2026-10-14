import { ILookupOption } from '../models/repeatingRows';

export class LookupCache {
  private readonly store: Map<string, { at: number; options: ILookupOption[] }> = new Map();

  public constructor(private readonly ttlMs: number) {}

  public get(key: string, now: number): ILookupOption[] | undefined {
    const hit = this.store.get(key);
    if (!hit) {
      return undefined;
    }
    if (now - hit.at > this.ttlMs) {
      this.store.delete(key);
      return undefined;
    }
    return hit.options;
  }

  public set(key: string, options: ILookupOption[], now: number): void {
    this.store.set(key, { at: now, options });
  }
}
