<script lang="ts">
    import { getMemory } from "$lib/command";
    import { Memory } from "$lib/types";
    import { formatBytes } from "$lib/utils";

    let memory = $state<Memory>();

    async function refreshMemory() {
        memory = await getMemory();
    }
</script>

<div>
    {#if memory}
        <p>
            Usage: {formatBytes(memory.usedMemory)} / {formatBytes(
                memory.totalMemory,
            )}
        </p>
        <p>
            Swap: {formatBytes(memory.usedSwap)} / {formatBytes(
                memory.totalSwap,
            )}
        </p>
    {:else}
        {#await refreshMemory()}
            <p>Loading</p>
        {/await}
    {/if}
</div>
