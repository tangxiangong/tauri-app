import { invoke } from "@tauri-apps/api/core";
import { Memory } from "./types.ts";

export async function isSupported(): Promise<boolean> {
  return await invoke("is_supported");
}

export async function getMemory(): Promise<Memory> {
  return await invoke("get_memory");
}
