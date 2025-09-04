import { open } from "@tauri-apps/plugin-dialog";

export async function selectFile(): Promise<string | null> {
  return await open({
    multiple: false,
    directory: false,
  });
}
